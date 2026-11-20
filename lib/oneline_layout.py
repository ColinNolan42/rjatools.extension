# -*- coding: ascii -*-
"""
oneline_layout.py

One-Line rebuild step 4 (2026-11-18): places the schematic tree on a grid
and returns a flat draw list. Pure Python, no Revit import, so every layout
decision is tested in CPython. ASCII source.

    out = layout(sch, notes=None)
    out["prims"]   list of draw primitives, drafting view feet (1:100)
    out["bounds"]  (x0, y0, x1, y1)
    out["log"]     placement decisions (stretches, header rows)

Primitives (every one carries "box" = (x0, y0, x1, y1)):
    {"t": "line", "x0", "y0", "x1", "y1", "new": bool, "seg": sid}
    {"t": "text", "x", "y", "lines": [..], "align": "center"|"left",
     "underline_first": bool, "role": "seg"|"fixture"|"notes"|"utility",
     "seg" or "node"}
    {"t": "symbol", "role": "meter"|"equipment"|"cap"|"prv"|"iso"|"ansul"|
     "squiggle", "x", "y", "rot": 0|90|180|270, "node"}
    x, y of a symbol is its center (its family insertion point, except the
    cap, which carries "ix", "iy": the pipe end); of a text, the middle of
    its top edge (align center) or its top left corner (align left).

Rules (plan section 4, Colin 2026-11-18):
- Left to right, meter at the far left. No scale horizontally: one column
  per Segment.
- Elevation vertical by level band (Colin, 2026-11-18): the levels the
  meter and the equipment sit on each get a band. A band draws its
  equipment on its base row and the headers that run at that level's
  ceiling one row above, with drops down to the equipment. The meter sits
  on the base row of its level, so first floor equipment (kitchen, WH-1)
  shares the meter row and rooftop units sit higher.
- A roof band hangs UP: its header runs on the band's base row and the
  units rise from it, with the header labels below the line (Colin,
  2026-11-19). A band is the roof when its level name contains ROOF. Only
  when the model has NO level named ROOF, units more than
  ROOF_ABOVE_TOP_FT above the top level get a roof band of their own, and
  each such unit is listed in the One-Line warnings as a note (Colin,
  2026-11-20). Every other band hangs down. At a tie in the existing meter
  sits at the lowest level.
- The meter outlet rises to its header; nothing else runs along the meter
  row at the meter. A sub-header moved off its parent's row may sit on the
  meter row away from the meter (Colin, 2026-11-20: the approved G3C
  cookline manifold and G2 first floor boilers do).
- A tap to equipment is a vertical drop to its band's equipment row; a tap
  holding more junctions rises or drops to its own level's header row. A
  run that changes level becomes a riser too. A second leg the same way
  from one tee steps right first.
- A Segment label sits on the Segment's LAST drawn line: centered above a
  horizontal line, right of and centered on a vertical one (item 4).
- The last unit on a header turns up or down like the other units on that
  line, one column past the last tap; nothing runs straight into a unit
  (Colin, 2026-11-19, supersedes item 7).
- Every equipment end has a PRV symbol on the pipe just before it.
- Any clash of a line, label or symbol with what is already placed moves
  the junction one column right, or steps the leg further right.
"""

import math
import time

import gas_labels

# ---------------------------------------------------------------------------
# sizes, drafting view feet at 1:100 (3/32 in. text = 0.78 ft)
# ---------------------------------------------------------------------------
COL_FT = 10.0
ROW_FT = 9.0
TEXT_H_FT = 1.39            # first line, measured in Revit 2026-11-19 (3/32" Arial, 1:100)
LINE_PITCH_FT = 1.26        # each further line, measured (was 0.78 and 1.3: labels hit symbols)
CHAR_W_FT = 0.52
LABEL_GAP_FT = 0.4
LABEL_SIDE_FT = 0.6
EQUIP_HALF_W = 1.0          # RJA - P Symbols - Equipment, 0.242 in. wide
EQUIP_HALF_H = 0.56
CAP_HALF_W = 0.85           # RJA- Equipment Cap: one 0.2 in. line at its insertion
CAP_HALF_H = 0.15
PRV_HALF_L = 1.0            # along the pipe
PRV_HALF_T = 1.0            # across the pipe (dome side)
ISO_HALF_L = 1.0
ISO_HALF_T = 0.6
PRV_BEFORE_END_FT = 2.6     # PRV center back from the equipment symbol
VALVE_FROM_START_FT = 2.5
VALVE_PITCH_FT = 2.6
METER_R = 1.04              # RJA - P Symbols - Meter Annotation, 1/4 in.
UTILITY_RUN_FT = 5.0
UTILITY_DROP_FT = 3.0
SQUIGGLE_HALF_L = 0.5       # RJA - Squiggle along the line (0.97 ft at 1:100)
SQUIGGLE_HALF_W = 0.2       # across the line (0.32 ft)
NOTES_GAP_FT = 3.0
MAX_STRETCH = 12
MAX_ROWS_OUT = 12
BAND_FT = 20.0             # one level band: equipment row, drop, header row
LEVEL_TOL_FT = 0.25        # an element this close below a level is on it
# units this far above the top level are on a roof of their own, ONLY when
# no level is named ROOF (Colin, 2026-11-20): a ROOF level wins, and a model
# without one still draws its RTUs hanging up
ROOF_ABOVE_TOP_FT = 8.0
INLINE_Z_FT = 1.0          # a run end this close to the tee height is inline
ROOT_RISER_COL = 0.5       # meter outlet riser, in columns from the meter
EXISTING_COLS = 2          # tie in: existing meter to the new connection
EPS = 0.05
# placement checks before the search gives up: a diagram that cannot be laid
# out fails in seconds instead of hanging Revit. G2 and G3C need well under
# a tenth of this.
MAX_CHECKS = 20000
# and a wall clock limit: the button runs on Revit's main thread and cannot
# be interrupted (Grantham 4 froze Revit 2024, 2026-11-19)
MAX_SECONDS = 20.0
# two primitives farther apart than this cannot clash (largest margin used by
# clash() is SYMBOL_MARGIN_FT)
NEAR_FT = 1.0
LOG_LIMIT = 40              # layout messages kept for the console


class _OutOfBudget(Exception):
    pass

UTILITY_TEXT = ["GAS FROM UTILITY"]


def X(c):
    return c * COL_FT


def node_line_id(node_id):
    """Pseudo Segment id of a junction's node line (review n14)."""
    return "J{0}".format(node_id)


def Y(r):
    return r * ROW_FT


# ---------------------------------------------------------------------------
# geometry checks
# ---------------------------------------------------------------------------

def _text_box(x, y_top, lines, align):
    w = max([len(s) for s in lines] + [1]) * CHAR_W_FT
    h = (len(lines) - 1) * LINE_PITCH_FT + TEXT_H_FT
    if align == "center":
        return (x - w / 2.0, y_top - h, x + w / 2.0, y_top)
    return (x, y_top - h, x + w, y_top)


def _text_size(lines):
    w = max([len(s) for s in lines] + [1]) * CHAR_W_FT
    h = (len(lines) - 1) * LINE_PITCH_FT + TEXT_H_FT
    return w, h


def _boxes_overlap(a, b, eps=EPS):
    return (a[0] < b[2] - eps and b[0] < a[2] - eps and
            a[1] < b[3] - eps and b[1] < a[3] - eps)


def _seg_in_box(p, box, eps=EPS):
    """Does segment p (x0, y0, x1, y1) enter box shrunk by eps."""
    bx0, by0, bx1, by1 = box[0] + eps, box[1] + eps, box[2] - eps, box[3] - eps
    if bx0 >= bx1 or by0 >= by1:
        return False
    x0, y0, x1, y1 = p
    # Liang-Barsky clip
    dx, dy = x1 - x0, y1 - y0
    t0, t1 = 0.0, 1.0
    for pp, qq in ((-dx, x0 - bx0), (dx, bx1 - x0), (-dy, y0 - by0), (dy, by1 - y0)):
        if abs(pp) < 1e-12:
            if qq < 0:
                return False
        else:
            t = qq / pp
            if pp < 0:
                if t > t1:
                    return False
                if t > t0:
                    t0 = t
            else:
                if t < t0:
                    return False
                if t < t1:
                    t1 = t
    return t1 - t0 > 1e-9


def _cross(ax, ay, bx, by):
    return ax * by - ay * bx


def _near(p, q, tol=EPS):
    return abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol


def _lines_clash(a, b, eps=EPS):
    """Lines may meet only where BOTH end (every tee and corner is drawn
    that way). Any other touch, crossing or collinear overlap clashes."""
    ends_a = ((a[0], a[1]), (a[2], a[3]))
    ends_b = ((b[0], b[1]), (b[2], b[3]))
    shared = [p for p in ends_a if _near(p, ends_b[0]) or _near(p, ends_b[1])]
    if _segments_meet(a, b, eps):
        if not shared:
            return True
        # meeting only at the shared end point: fine unless they also run
        # along each other
        return _collinear_overlap(a, b) > eps
    return False


def _collinear_overlap(a, b):
    r = (a[2] - a[0], a[3] - a[1])
    s = (b[2] - b[0], b[3] - b[1])
    qp = (b[0] - a[0], b[1] - a[1])
    rr = r[0] * r[0] + r[1] * r[1]
    if rr < 1e-12 or abs(_cross(r[0], r[1], s[0], s[1])) > 1e-9 or \
            abs(_cross(qp[0], qp[1], r[0], r[1])) > 1e-9:
        return 0.0
    t0 = (qp[0] * r[0] + qp[1] * r[1]) / rr
    t1 = t0 + (s[0] * r[0] + s[1] * r[1]) / rr
    lo, hi = min(t0, t1), max(t0, t1)
    return max(0.0, min(hi, 1.0) - max(lo, 0.0)) * (rr ** 0.5)


def _segments_meet(a, b, eps=EPS):
    """True when the two segments come within eps of each other."""
    return _seg_dist(a, b) <= eps


def _pt_seg_dist(px, py, s):
    x0, y0, x1, y1 = s
    dx, dy = x1 - x0, y1 - y0
    L2 = dx * dx + dy * dy
    if L2 < 1e-12:
        return ((px - x0) ** 2 + (py - y0) ** 2) ** 0.5
    t = max(0.0, min(1.0, ((px - x0) * dx + (py - y0) * dy) / L2))
    cx, cy = x0 + t * dx, y0 + t * dy
    return ((px - cx) ** 2 + (py - cy) ** 2) ** 0.5


def _seg_dist(a, b):
    if _proper_cross(a, b):
        return 0.0
    return min(_pt_seg_dist(a[0], a[1], b), _pt_seg_dist(a[2], a[3], b),
               _pt_seg_dist(b[0], b[1], a), _pt_seg_dist(b[2], b[3], a))


def _proper_cross(a, b):
    def orient(p, q, r):
        return _cross(q[0] - p[0], q[1] - p[1], r[0] - p[0], r[1] - p[1])
    p1, p2 = (a[0], a[1]), (a[2], a[3])
    q1, q2 = (b[0], b[1]), (b[2], b[3])
    d1, d2 = orient(q1, q2, p1), orient(q1, q2, p2)
    d3, d4 = orient(p1, p2, q1), orient(p1, p2, q2)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0)) and \
        d1 != 0 and d2 != 0 and d3 != 0 and d4 != 0


SYMBOL_MARGIN_FT = 0.3


def _line_vs_symbol(l, box):
    """A line may come near a symbol only when it ends there (its own pipe);
    then it shall not run through the symbol."""
    m = SYMBOL_MARGIN_FT
    grown = (box[0] - m, box[1] - m, box[2] + m, box[3] + m)
    if not _seg_in_box(l, grown, 0.0):
        return False
    ends_at = (_inside(l[0], l[1], grown) or _inside(l[2], l[3], grown))
    if not ends_at:
        return True
    return _seg_in_box(l, box)


def _inside(x, y, b):
    return b[0] <= x <= b[2] and b[1] <= y <= b[3]


def clash(a, b):
    """True when two primitives overlap in a way a reader would notice."""
    if a["t"] == "line" and b["t"] == "line":
        if "joins" in a and "joins" in b and \
                b.get("seg") not in a["joins"] and _ends_touch(a, b):
            # G2 2026-11-19: AHU-2 turning up and the boiler riser met end
            # to end and read as one header; only connected Segments meet
            return True
        return _lines_clash(_pts(a), _pts(b))
    if a["t"] == "line" or b["t"] == "line":
        l, o = (a, b) if a["t"] == "line" else (b, a)
        if o["t"] == "symbol":
            if o.get("line_to_center"):
                p = _pts(l)
                c = (o["x"], o["y"])
                if _near((p[0], p[1]), c) or _near((p[2], p[3]), c):
                    return False
            return _line_vs_symbol(_pts(l), o["box"])
        return _seg_in_box(_pts(l), o["box"])
    return _boxes_overlap(a["box"], b["box"])


def _pts(l):
    return (l["x0"], l["y0"], l["x1"], l["y1"])


def _ends_touch(a, b):
    pa, pb = _pts(a), _pts(b)
    for e in ((pa[0], pa[1]), (pa[2], pa[3])):
        if _near(e, (pb[0], pb[1])) or _near(e, (pb[2], pb[3])):
            return True
    return False


def all_clashes(prims):
    out = []
    for i in range(len(prims)):
        for j in range(i + 1, len(prims)):
            if clash(prims[i], prims[j]):
                out.append((i, j))
    return out


# ---------------------------------------------------------------------------
# primitive builders
# ---------------------------------------------------------------------------

def _line(x0, y0, x1, y1, seg=None, new=True):
    return {"t": "line", "x0": x0, "y0": y0, "x1": x1, "y1": y1, "new": new,
            "seg": seg, "box": (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))}


def _text(x, y_top, lines, align="center", role="seg", underline_first=False,
          **kw):
    p = {"t": "text", "x": x, "y": y_top, "lines": list(lines),
         "align": align, "role": role, "underline_first": underline_first,
         "box": _text_box(x, y_top, lines, align)}
    p.update(kw)
    return p


def _symbol(role, x, y, rot, half_along, half_across, node=None):
    if rot in (90, 270):
        box = (x - half_across, y - half_along, x + half_across, y + half_along)
    else:
        box = (x - half_along, y - half_across, x + half_along, y + half_across)
    return {"t": "symbol", "role": role, "x": x, "y": y, "rot": rot,
            "node": node, "box": box}


def _squiggle(x, y_end):
    """Break symbol across the end of a vertical line at (x, y_end), the
    line running into its center (Colin, 2026-11-19). The family drawn at
    rot 90 lies across a vertical line; the draw layer centers its view box
    on (x, y_end). "line_to_center" lets the line end inside its box."""
    s = _symbol("squiggle", x, y_end, 90, SQUIGGLE_HALF_L, SQUIGGLE_HALF_W)
    s["box"] = (x - SQUIGGLE_HALF_L, y_end - SQUIGGLE_HALF_W,
                x + SQUIGGLE_HALF_L, y_end + SQUIGGLE_HALF_W)
    s["line_to_center"] = True
    return s


def _unit(ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    n = (dx * dx + dy * dy) ** 0.5
    return (dx / n, dy / n, n) if n > 1e-9 else (0.0, 0.0, 0.0)


def _rot_for(ux, uy):
    return 90 if abs(uy) > abs(ux) else 0


# ---------------------------------------------------------------------------
# layout engine
# ---------------------------------------------------------------------------

class _Engine(object):

    def __init__(self, sch):
        self.sch = sch
        self.segs = sch["segments"]
        self.log = []
        self.checks = 0
        self.t0 = time.time()
        self._bands(sch)
        self._adjacency(sch.get("root"))

    def _adjacency(self, root):
        """Segments whose lines may meet end to end: consecutive Segments
        of a branch, and the Segments at one tee. None is the meter service
        (utility, existing run), which meets the first Segment."""
        adj = {None: set([None])}

        def link(a, b):
            adj.setdefault(a, set([a])).add(b)
            adj.setdefault(b, set([b])).add(a)

        def rec(br, up):
            segs = br.get("segments") or []
            if not segs:
                return
            link(up, segs[0])
            for a, b in zip(segs, segs[1:]):
                link(a, b)
            end = br.get("end")
            if end and end["type"] == "junction":
                legs = list(end["taps"]) + [end["run"]]
                firsts = [l["segments"][0] for l in legs if l.get("segments")]
                # the junction's own node line, drawn along the header when
                # a leg steps right (review n14), meets all of them
                jid = node_line_id(end["node_id"])
                link(segs[-1], jid)
                for f in firsts:
                    link(segs[-1], f)
                    link(f, jid)
                    for g in firsts:
                        link(f, g)
                for l in legs:
                    rec(l, segs[-1])

        if root is not None:
            rec(root, None)
        self.adj = adj

    def tag(self, prims):
        """Give each line the Segments it may meet."""
        for p in prims:
            if p["t"] == "line" and "joins" not in p:
                p["joins"] = self.adj.get(p.get("seg"), set([p.get("seg")]))
        return prims

    # -- elevation bands (Colin, 2026-11-18) ------------------------------
    def _bands(self, sch):
        """Bands come from the levels the meter and the equipment sit on:
        band i draws its equipment at y = i * BAND_FT and the headers that
        run at that level's ceiling one row above. The meter shares the
        band of the first level equipment.

        Roof (Colin, 2026-11-20, review n31 and n32): a band is the roof
        when its level name contains ROOF (any case). A ROOF level wins:
        only when no level is named ROOF do units more than
        ROOF_ABOVE_TOP_FT above the top level get a roof band of their own
        (the height rule), and each of them is kept in roof_by_height for
        the console note. With a ROOF level, a unit high above the top level
        that is not on a ROOF band hangs down as normal. Roof units hang
        up, the rest hang down; the meter's elevation plays no part. At a
        tie in the existing meter is drawn at the lowest level, not at the
        picked pipe's elevation."""
        lv = sorted([(float(l["elev"]), str(l.get("name") or ""))
                     for l in sch.get("levels") or []])
        self.levels = [e for e, _n in lv]
        self.roof_levels = set(e for e, n in lv if "ROOF" in n.upper())
        self.top_elev = self.levels[-1] if self.levels else None
        self.top_name = lv[-1][1] if lv else ""
        # the height rule applies only with no level named ROOF
        self.height_rule = bool(self.levels) and not self.roof_levels
        oz = (sch.get("origin_xyz") or [0, 0, 0])[2]
        if sch.get("origin_kind") == "tie_in" and self.levels:
            oz = self.levels[0]
        self.origin_z = oz
        zs = [oz]
        self.roof_by_height = []      # (name, node id, ft above top level)
        for t in self._terms(sch.get("root")):
            if t.get("xyz"):
                z = t["xyz"][2]
                zs.append(z)
                if self._above_top(z):
                    self.roof_by_height.append(
                        (t.get("name") or "UNNAMED", t.get("node_id"),
                         float(z) - self.top_elev))
        used = set()
        for z in zs:
            used.add(self._level_key(z))
        self.band_elevs = sorted(used)
        self._band_index = dict((e, i) for i, e in enumerate(self.band_elevs))
        self.up_bands = set(i for i, e in enumerate(self.band_elevs)
                            if self._is_roof_key(e))
        # header rows of hang up bands: their Segment labels go below the
        # line, the units' own labels take the space above
        self.up_header_ys = [b * BAND_FT for b in self.up_bands]

    def hangs_up(self, z):
        return self.band(z) in self.up_bands

    @staticmethod
    def _level_of(z, levels):
        if not levels:
            return 0.0
        best = levels[0]
        for e in levels:
            if e <= z + LEVEL_TOL_FT:
                best = e
        return best

    def _above_top(self, z):
        """True when the height rule makes z a roof unit: no level is named
        ROOF and z is more than ROOF_ABOVE_TOP_FT above the top level."""
        return (self.height_rule and
                z > self.top_elev + ROOF_ABOVE_TOP_FT)

    def _level_key(self, z):
        """Elevation of the level z sits on, or (height rule, no ROOF
        level) the roof band above the top level for units more than
        ROOF_ABOVE_TOP_FT above it."""
        if self._above_top(z):
            return self.top_elev + ROOF_ABOVE_TOP_FT
        return self._level_of(z, self.levels)

    def _is_roof_key(self, e):
        if e in self.roof_levels:
            return True
        return (self.height_rule and
                e > self.top_elev + ROOF_ABOVE_TOP_FT - 1e-6)

    def roof_notes(self):
        """One note per unit the height rule made a roof unit, so the
        engineer can check it or rename a level (Colin, 2026-11-20)."""
        return [roof_height_note(name, nid, ft, self.top_name)
                for name, nid, ft in self.roof_by_height]

    def band(self, z):
        k = self._level_key(z)
        if k in self._band_index:
            return self._band_index[k]
        b = 0
        for i, e in enumerate(self.band_elevs):
            if e <= z + LEVEL_TOL_FT:
                b = i
        return b

    def y_equip(self, z):
        return self.band(z) * BAND_FT + (ROW_FT if self.hangs_up(z) else 0.0)

    def y_header(self, z):
        return self.band(z) * BAND_FT + (0.0 if self.hangs_up(z) else ROW_FT)

    def turn_y(self, end, y, rows=1):
        """Equipment row for terminal end reached from a header at y: its
        band's row, or rows rows off the header the way its band hangs. A
        unit never turns against the way its band hangs (review n13), and
        the turn is at least rows rows long (review n1)."""
        z = self._z(end)
        s = 1 if self.hangs_up(z) else -1
        ty = self.y_equip(z)
        need = rows * ROW_FT
        if abs(ty - y) < ROW_FT - 1e-6 or (ty - y) * s < 0:
            return y + s * need
        if abs(ty - y) < need - 1e-6:
            return y + (need if ty > y else -need)
        return ty

    # -- how long a drop must be (review n1) -------------------------------
    def _piece_need(self, seg, end=None):
        """Length one vertical piece needs: its valves from the start, its
        label, and the PRV before equipment at the end."""
        _w, h = _text_size(seg.get("label") or [""])
        k = len(seg.get("valves") or [])
        v = 0.0
        if k:
            v = VALVE_FROM_START_FT + (k - 1) * VALVE_PITCH_FT + \
                max(PRV_HALF_L, ISO_HALF_L)
        e = 0.0
        if end is not None and end.get("kind") == "equipment":
            e = PRV_BEFORE_END_FT + PRV_HALF_L
        return v + h + 2.0 * LABEL_GAP_FT + e

    def _piece_needs(self, segs, end):
        last = len(segs) - 1
        return [self._piece_need(s, end if i == last else None)
                for i, s in enumerate(segs)]

    @staticmethod
    def _rows_for(length):
        return max(1, int(math.ceil(length / ROW_FT - 1e-9)))

    # -- helpers -----------------------------------------------------------
    def fits(self, new, occ):
        self.checks += 1
        if self.checks > MAX_CHECKS or \
                time.time() - self.t0 > MAX_SECONDS:
            raise _OutOfBudget()
        self.tag(new)
        # only primitives near each other can clash: a box test first keeps
        # a large system (Grantham 4, 2026-11-19) from comparing every new
        # piece with everything already drawn
        m = NEAR_FT
        for a in new:
            ab = a["box"]
            x0, y0, x1, y1 = ab[0] - m, ab[1] - m, ab[2] + m, ab[3] + m
            for b in occ:
                bb = b["box"]
                if bb[0] > x1 or bb[2] < x0 or bb[1] > y1 or bb[3] < y0:
                    continue
                if clash(a, b):
                    return False
        for i in range(len(new)):
            for j in range(i + 1, len(new)):
                if clash(new[i], new[j]):
                    return False
        return True

    def _terms(self, br, out=None):
        out = [] if out is None else out
        if br is None or br.get("end") is None:
            return out
        end = br["end"]
        if end["type"] == "terminal":
            out.append(end)
        else:
            self._terms(end["run"], out)
            for t in end["taps"]:
                self._terms(t, out)
        return out

    @staticmethod
    def _z(item):
        return (item.get("xyz") or [0, 0, 0])[2]

    # -- one Segment drawn along a polyline -------------------------------
    def _seg_prims(self, seg, pts, end_inserts=None, start_trim=0.0):
        """Lines for a Segment along polyline pts (list of (x, y)), with its
        valves on the last piece, end_inserts (symbols near the end of the
        last piece, as (distance from end, half length, prim)), and its label
        centered on the last piece."""
        sid = seg["id"]
        new = not seg.get("is_existing")
        prims = []
        pieces = []
        for i in range(len(pts) - 1):
            pieces.append([pts[i], pts[i + 1]])
        if start_trim:
            (ax, ay), (bx, by) = pieces[0]
            ux, uy, _n = _unit(ax, ay, bx, by)
            pieces[0][0] = (ax + ux * start_trim, ay + uy * start_trim)
        last = pieces[-1]
        (ax, ay), (bx, by) = last
        ux, uy, n = _unit(ax, ay, bx, by)
        rot = _rot_for(ux, uy)
        cuts = []         # (t from start of last piece, half length)
        t = VALVE_FROM_START_FT
        for v in seg.get("valves") or []:
            if v["kind"] == "prv":
                sym = _symbol("prv", ax + ux * t, ay + uy * t, rot,
                              PRV_HALF_L, PRV_HALF_T, v["node_id"])
                cuts.append((t, PRV_HALF_L))
            else:
                sym = _symbol("ansul" if v["kind"] == "ansul" else "iso",
                              ax + ux * t, ay + uy * t, rot,
                              ISO_HALF_L, ISO_HALF_T, v["node_id"])
                cuts.append((t, ISO_HALF_L))
                if v["kind"] == "ansul":
                    # the hood shutoff is named, on the side away from the
                    # Segment label (Colin, 2026-11-19)
                    prims.append(self._valve_label(sym, rot, ay,
                                                   gas_labels.ANSUL_LABEL))
            prims.append(sym)
            t += VALVE_PITCH_FT
        for dist, half, prim in end_inserts or []:
            cuts.append((n - dist, half))
            prims.append(prim)
        # line pieces, the last one cut around its symbols
        for pc in pieces[:-1]:
            prims.append(_line(pc[0][0], pc[0][1], pc[1][0], pc[1][1], sid, new))
        cuts.sort()
        cur = 0.0
        for tc, half in cuts:
            a, b = cur, tc - half
            if b - a > 1e-6:
                prims.append(_line(ax + ux * a, ay + uy * a, ax + ux * b,
                                   ay + uy * b, sid, new))
            cur = tc + half
        if n - cur > 1e-6:
            prims.append(_line(ax + ux * cur, ay + uy * cur, bx, by, sid, new))
        # label on the last piece (item 4), centered on its longest stretch
        # clear of valve and PRV symbols
        lines = seg["label"]
        w, h = _text_size(lines)
        best = (0.0, n)
        if cuts:
            spans = []
            cur = 0.0
            for tc, half in cuts:
                spans.append((cur, tc - half))
                cur = tc + half
            spans.append((cur, n))
            best = max(spans, key=lambda s: s[1] - s[0])
        mid = (best[0] + best[1]) / 2.0
        mx, my = ax + ux * mid, ay + uy * mid
        if rot == 0 and any(abs(ay - uy_) < 1e-6
                            for uy_ in self.up_header_ys):
            prims.append(_text(mx, ay - LABEL_GAP_FT, lines, "center",
                               "seg", seg=sid))
        elif rot == 0:
            prims.append(_text(mx, ay + LABEL_GAP_FT + h, lines, "center",
                               "seg", seg=sid))
        else:
            prims.append(_text(ax + LABEL_SIDE_FT, my + h / 2.0, lines, "left",
                               "seg", seg=sid))
        return prims

    def _valve_label(self, sym, rot, ay, lines):
        """A named valve's label, opposite the Segment label: left of a
        vertical pipe; below a horizontal one, above on a header whose units
        hang up."""
        w, h = _text_size(lines)
        b = sym["box"]
        if rot != 0:
            return _text(b[0] - LABEL_SIDE_FT - w, sym["y"] + h / 2.0, lines,
                         "left", "valve", node=sym["node"])
        if any(abs(ay - uy_) < 1e-6 for uy_ in self.up_header_ys):
            return _text(sym["x"], b[3] + LABEL_GAP_FT + h, lines, "center",
                         "valve", node=sym["node"])
        return _text(sym["x"], b[1] - LABEL_GAP_FT, lines, "center", "valve",
                     node=sym["node"])

    # -- terminal at the end of a pipe ------------------------------------
    def _terminal_inserts(self, end, ex, ey, ux, uy):
        """Symbols at a pipe end arriving in direction (ux, uy): returns
        (end_inserts for _seg_prims, extra prims: equipment symbol, label)."""
        inserts = []
        extra = []
        rot = _rot_for(ux, uy)
        if end["kind"] == "equipment":
            half_t = EQUIP_HALF_H
            # symbol center sits past the pipe end by its half thickness
            sx, sy = ex + ux * half_t, ey + uy * half_t
            sym_rot = 90 if rot == 0 else 0
            sym = _symbol("equipment", sx, sy, sym_rot, EQUIP_HALF_W,
                          EQUIP_HALF_H, end["node_id"])
            if sym_rot == 90:
                sym["box"] = (sx - EQUIP_HALF_H, sy - EQUIP_HALF_W,
                              sx + EQUIP_HALF_H, sy + EQUIP_HALF_W)
            else:
                sym["box"] = (sx - EQUIP_HALF_W, sy - EQUIP_HALF_H,
                              sx + EQUIP_HALF_W, sy + EQUIP_HALF_H)
            extra.append(sym)
            prv = _symbol("prv", ex - ux * PRV_BEFORE_END_FT,
                          ey - uy * PRV_BEFORE_END_FT, rot, PRV_HALF_L,
                          PRV_HALF_T, end.get("modeled_prv_id"))
            prv["equipment_prv"] = True
            inserts.append((PRV_BEFORE_END_FT, PRV_HALF_L, prv))
            far = sym["box"]
        elif end["kind"] == "cap":
            sx, sy = ex + ux * CAP_HALF_H, ey + uy * CAP_HALF_H
            sym_rot = {(0, -1): 0, (0, 1): 180, (1, 0): 90, (-1, 0): 270}.get(
                (int(round(ux)), int(round(uy))), 0)
            sym = _symbol("cap", sx, sy, sym_rot, CAP_HALF_W, CAP_HALF_H,
                          end["node_id"])
            if sym_rot in (90, 270):
                sym["box"] = (sx - CAP_HALF_H, sy - CAP_HALF_W,
                              sx + CAP_HALF_H, sy + CAP_HALF_W)
            sym["ix"], sym["iy"] = ex, ey      # the cap line sits on the end
            extra.append(sym)
            far = sym["box"]
        else:
            far = (ex, ey, ex, ey)
        lines = end.get("label") or []
        if lines:
            w, h = _text_size(lines)
            if rot == 0:      # inline: label right of the symbol (item 7)
                if ux >= 0:
                    extra.append(_text(far[2] + LABEL_SIDE_FT, ey + h / 2.0,
                                       lines, "left", "fixture",
                                       underline_first=end["kind"] == "equipment",
                                       node=end["node_id"]))
                else:
                    extra.append(_text(far[0] - LABEL_SIDE_FT - w, ey + h / 2.0,
                                       lines, "left", "fixture",
                                       underline_first=end["kind"] == "equipment",
                                       node=end["node_id"]))
            elif uy < 0:      # drop down: label below
                extra.append(_text(ex, far[1] - LABEL_GAP_FT, lines, "center",
                                   "fixture",
                                   underline_first=end["kind"] == "equipment",
                                   node=end["node_id"]))
            else:             # drop up: label above
                extra.append(_text(ex, far[3] + LABEL_GAP_FT + h, lines,
                                   "center", "fixture",
                                   underline_first=end["kind"] == "equipment",
                                   node=end["node_id"]))
        return inserts, extra

    def _branch_prims(self, br, pts_list, end_xy, end_dir, start_trim=0.0):
        """Prims for every Segment of br along its polylines; the last one
        ends at the terminal when br ends at one."""
        prims = []
        segs = [self.segs[s] for s in br["segments"]]
        end = br["end"]
        for i, seg in enumerate(segs):
            inserts = None
            if i == len(segs) - 1 and end["type"] == "terminal":
                inserts, extra = self._terminal_inserts(
                    end, end_xy[0], end_xy[1], end_dir[0], end_dir[1])
                prims += extra
            prims += self._seg_prims(seg, pts_list[i], inserts,
                                     start_trim if i == 0 else 0.0)
        return prims

    # -- headers, legs, junctions -----------------------------------------
    def header(self, br, cx, y, occ, from_y=None, riser_dx=0.0,
               start_xy=None):
        """Lay br out from column cx to the right at height y. With from_y,
        the first Segment first runs vertically from from_y at column
        cx + riser_dx (reached from start_xy horizontally, for the meter).
        Returns the new occupancy list, or None."""
        n = len(br["segments"])
        end = br["end"]
        for stretch in range(0, MAX_STRETCH + 1):
            pts_list = []
            for i in range(n):
                c0 = cx + i
                c1 = cx + i + 1 + (stretch if i == n - 1 else 0)
                pts = [(X(c0), y), (X(c1), y)]
                if i == 0 and from_y is not None and abs(from_y - y) > 1e-6:
                    xr = X(cx + riser_dx)
                    pts = [(xr, from_y), (xr, y), (X(c1), y)]
                    if start_xy is not None:
                        pts = [start_xy] + pts
                elif i == 0 and start_xy is not None:
                    pts = [start_xy, (X(c1), y)]
                pts_list.append(pts)
            ce = cx + n + stretch
            end_xy, end_dir = (X(ce), y), (1.0, 0.0)
            if end["type"] == "terminal":
                # the last unit on a header turns like the others on that
                # line, never straight on (Colin, 2026-11-19); the turn is
                # long enough for its valves, label and PRV (review n1)
                rows = self._rows_for(self._piece_need(
                    self.segs[br["segments"][-1]], end))
                ty = self.turn_y(end, y, rows)
                pts_list[-1] = pts_list[-1] + [(X(ce), ty)]
                end_xy = (X(ce), ty)
                end_dir = (0.0, 1.0 if ty > y else -1.0)
            new = self._branch_prims(br, pts_list, end_xy, end_dir)
            if not self.fits(new, occ):
                continue
            trial = occ + new
            if end["type"] == "terminal":
                if stretch:
                    self.log.append("stretch {0} col(s) before {1}".format(
                        stretch, end.get("name")))
                return trial
            res = self.junction(end, ce, y, trial)
            if res is not None:
                if stretch:
                    self.log.append("stretch {0} col(s) before junction "
                                    "{1}".format(stretch, end["node_id"]))
                return res
        return None

    def _is_inline(self, leg, j):
        end = leg["end"]
        return (end["type"] == "terminal" and
                abs(self._z(end) - self._z(j)) <= INLINE_Z_FT)

    def junction(self, j, c, y, occ):
        legs = list(j["taps"]) + [j["run"]]
        run = j["run"]
        cont = None
        # the run carries on along this row when its next tee is on this
        # row, or on this tee's level (a header moved off its row keeps
        # going on the moved row: G2 first floor boilers, 2026-11-19)
        if run["end"]["type"] == "junction" and (
                abs(self.y_header(self._z(run["end"])) - y) < 1e-6 or
                self.band(self._z(run["end"])) == self.band(self._z(j))):
            cont = run
        trial = occ
        used = {}          # direction -> vertical legs already at this tee
        reach = 0          # columns of node line drawn right of c
        for leg in legs:
            if leg is cont:
                continue
            res = None
            d_key = None
            # a junction leg, or the header's own end unit, may step as far
            # as a branch can: G3C 2026-11-19, FUTURE after the cookline
            # manifold had to clear its whole span
            jogs = MAX_STRETCH if (leg["end"]["type"] == "junction" or
                                   leg is run) else 4
            # the run's own unit turns one column on, so the header reads
            # as continuing to it (Colin, 2026-11-19)
            first = 1 if (leg is run and leg["end"]["type"] == "terminal") \
                else 0
            for jog in range(first, jogs + first):
                res, d_key = self.leg(leg, c, y, trial, jog, used, reach,
                                      j["node_id"])
                if res is not None:
                    break
            if res is None:
                return None
            used[d_key] = used.get(d_key, 0) + 1
            if jog > first:
                self.log.append("leg stepped {0} col(s) right at junction "
                                "{1}".format(jog, j["node_id"]))
            reach = max(reach, jog)
            trial = res
        if cont is None:
            return trial
        # the run carries on from the end of the node line (review n14)
        return self.header(cont, c + reach, y, trial)

    def _node_lines(self, jid, c, y, reach, jog):
        """Unit pieces of a junction's node line along the header, from
        column c + reach to c + jog: a leg that steps right leaves from
        here, so no leg runs along the header itself (review n14). The
        line is new or existing like the Segment arriving at the junction
        (an existing main's tee keeps the existing line style)."""
        new = self._junction_is_new(jid)
        out = []
        for k in range(reach, jog):
            p = _line(X(c + k), y, X(c + k + 1), y, node_line_id(jid), new)
            out.append(p)
        return out

    def _junction_is_new(self, jid):
        """False when the Segment ending at junction jid is existing pipe,
        else True (new, also when no incoming Segment is found)."""
        if jid is None:
            return True
        for seg in self.segs.values():
            if seg.get("end_node") == jid:
                return not seg.get("is_existing")
        return True

    def leg(self, leg, c, y, occ, jog, used, reach=0, jid=None):
        """A leg leaving the header vertically: a drop or riser to its
        equipment, or a riser to its own header at its level. Returns
        (occupancy or None, direction)."""
        end = leg["end"]
        n = len(leg["segments"])
        ext = self._node_lines(jid, c, y, reach, jog)
        if end["type"] == "terminal":
            segs = [self.segs[s] for s in leg["segments"]]
            needs = self._piece_needs(segs, end)
            tot = sum(needs)
            ty = self.turn_y(end, y, self._rows_for(tot))
            d = 1 if ty > y else -1
            if jog < used.get(d, 0):
                return None, d
            x = X(c + jog)
            pts_list = []
            acc = 0.0
            for i in range(n):
                y0 = y + (ty - y) * acc / tot
                acc += needs[i]
                y1 = ty if i == n - 1 else y + (ty - y) * acc / tot
                pts_list.append([(x, y0), (x, y1)])
            new = ext + self._branch_prims(leg, pts_list, (x, ty),
                                           (0.0, float(d)))
            if self.fits(new, occ):
                return occ + new, d
            return None, d
        return self._header_leg(leg, c, y, occ, jog, used, ext)

    def _header_rows(self, y, s):
        """Rows a header on its parent's row may move to: the way its
        units hang first, then the other way, alternating. The meter row
        is allowed: a sub-header may sit on it away from the meter (Colin,
        2026-11-20, review n13; the approved G3C cookline manifold and G2
        first floor boilers do). The clash check keeps it off the meter
        and its outlet."""
        out = []
        for k in range(1, MAX_ROWS_OUT + 1):
            out += [y + s * k * ROW_FT, y - s * k * ROW_FT]
        return out

    def _header_leg(self, leg, c, y, occ, jog, used, ext):
        end = leg["end"]
        ty = self.y_header(self._z(end))
        cands = [ty]
        if abs(ty - y) < 1e-6:
            # a header on its parent's row moves off it, first the way its
            # units hang (G2 2026-11-19: first floor boilers off the first
            # floor main went up into the roof labels and failed)
            s = 1 if self.hangs_up(self._z(end)) else -1
            cands = self._header_rows(y, s)
        d = 1 if cands[0] > y else -1
        if jog < used.get(d, 0):
            return None, d
        if ext:
            if not self.fits(ext, occ):
                return None, d
            occ = occ + ext
        for cy in cands:
            res = self.header(leg, c + jog, cy, occ, from_y=y)
            if res is not None:
                if cy != ty:
                    self.log.append("header {0} moved off its level "
                                    "row".format(end["node_id"]))
                return res, (1 if cy > y else -1)
        return None, d


def roof_height_note(name, node_id, above_ft, top_name=""):
    """The console note for a unit drawn as a roof unit by the height rule
    (no level named ROOF, more than ROOF_ABOVE_TOP_FT above the top
    level)."""
    return ("NOTE: no level is named ROOF, so {0} (element {1}), {2:.1f} ft "
            "above the top level{3}, is drawn as a roof unit (more than "
            "{4:g} ft above it). Check it, or rename the roof level to "
            "include ROOF.".format(
                name, node_id, float(above_ft),
                " ({0})".format(top_name) if top_name else "",
                float(ROOF_ABOVE_TOP_FT)))


def layout(sch, notes=None, origin_kind="meter"):
    """Draw list for a schematic. notes: the notes block lines
    (gas_labels.notes_lines).

    Units made roof units by the height rule are listed as notes in
    sch["warnings"] (the list the One-Line button already prints) and in
    the returned "roof_notes"."""
    eng = _Engine(sch)
    roof_notes = eng.roof_notes()
    warnings = sch.setdefault("warnings", [])
    for note in roof_notes:
        if note not in warnings:
            warnings.append(note)
    occ = []
    # the meter's elevation: at a tie in, the lowest level (review n32)
    oz = eng.origin_z
    my = eng.y_equip(oz)
    # meter and utility stub
    existing = sch.get("existing")
    occ.append(_symbol("meter", 0.0, my, 0, METER_R, METER_R, sch["origin"]))
    ux1 = -(METER_R + UTILITY_RUN_FT)
    # at a tie in the meter and its service are existing: thin, like the
    # existing piping (Colin, 2026-11-18)
    occ.append(_line(-METER_R, my, ux1, my, None, new=not existing))
    occ.append(_line(ux1, my, ux1, my - UTILITY_DROP_FT, None,
                     new=not existing))
    occ.append(_squiggle(ux1, my - UTILITY_DROP_FT))
    occ.append(_text(ux1 - 1.0, my - UTILITY_DROP_FT - SQUIGGLE_HALF_W
                     - LABEL_GAP_FT, UTILITY_TEXT, "left", "utility"))
    eng.tag(occ)
    root = sch.get("root")
    start_col = 0
    start_xy = (METER_R, my)
    if existing:
        # Colin 2026-10-13: start at the existing meter, the existing pipe
        # (thin, labelled) to the new connection, then the existing piping
        # on to a squiggle, to the existing appliances.
        start_col = EXISTING_COLS
        cx = X(EXISTING_COLS)
        occ.append(_line(METER_R, my, cx, my, None, new=False))
        lines = existing["label"]
        w, h = _text_size(lines)
        occ.append(_text((METER_R + cx) / 2.0, my + LABEL_GAP_FT + h, lines,
                         "center", "existing"))
        occ.append(_line(cx, my, cx, my - UTILITY_DROP_FT, None, new=False))
        occ.append(_squiggle(cx, my - UTILITY_DROP_FT))
        occ.append(_text(cx - 1.0, my - UTILITY_DROP_FT - SQUIGGLE_HALF_W
                         - LABEL_GAP_FT, gas_labels.EXISTING_APPLIANCES,
                         "left", "existing"))
        start_xy = (cx, my)
    if root is not None:
        end = root["end"]
        if end["type"] == "junction":
            ty = eng.y_header(eng._z(end))
        elif eng.y_equip(eng._z(end)) == my:
            ty = my
        else:
            ty = eng.y_header(eng._z(end))
        try:
            res = eng.header(root, start_col, ty, occ, from_y=my,
                             riser_dx=ROOT_RISER_COL, start_xy=start_xy)
        except _OutOfBudget:
            # the search stops rather than hang Revit (G3C 2026-11-19 ran
            # past five minutes in IronPython)
            eng.log.append("search stopped after {0} placement checks, "
                           "{1:.0f} s".format(eng.checks,
                                              time.time() - eng.t0))
            res = None
        if res is None:
            eng.log.append("LAYOUT FAILED: no clash free placement found")
            res = occ
        occ = res
    x0 = min(p["box"][0] for p in occ)
    y0 = min(p["box"][1] for p in occ)
    x1 = max(p["box"][2] for p in occ)
    y1 = max(p["box"][3] for p in occ)
    if notes:
        top = y1 + NOTES_GAP_FT + (len(notes) - 1) * LINE_PITCH_FT + TEXT_H_FT
        occ.append(_text(x0, top, notes, "left", "notes"))
        y1 = top
    return {"prims": occ, "bounds": (x0, y0, x1, y1), "log": _short_log(eng.log),
            "roof_notes": roof_notes}


def failed(log):
    """True when the layout log says no placement was found."""
    return any(l.startswith("LAYOUT FAILED") for l in log)


def failure_message(log):
    """Alert text for a failed layout. A search stopped on the time or
    check budget did not prove an overlap and says so (review n43)."""
    stop = [l for l in log if l.startswith("search stopped")]
    if stop:
        return ("The layout search ran out of time ({0}) before it found a "
                "placement without overlaps. Nothing was drawn. No overlap "
                "was proven; the layout diagnostic is in the "
                "console.".format(stop[0][len("search stopped "):]))
    return ("The diagram could not be laid out without overlaps. Nothing "
            "was drawn. The layout diagnostic is in the console.")


def _short_log(log, limit=LOG_LIMIT):
    """Each message once, in order, at most limit lines: a long search logged
    the same move thousands of times and the pyRevit console tied up Revit
    printing it (Grantham 4, 2026-11-19)."""
    out, seen = [], set()
    for ln in log:
        if ln not in seen:
            seen.add(ln)
            out.append(ln)
    if len(out) > limit:
        # failures are always kept: the button checks for LAYOUT FAILED
        keep = [ln for ln in out if ln.startswith(("LAYOUT FAILED",
                                                    "search stopped"))]
        rest = [ln for ln in out if ln not in keep]
        extra = len(rest) - (limit - len(keep))
        out = keep + rest[:limit - len(keep)] + [
            "... {0} more layout messages".format(extra)]
    return out
