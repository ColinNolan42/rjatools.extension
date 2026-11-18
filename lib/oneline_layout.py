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
    {"t": "symbol", "role": "meter"|"equipment"|"cap"|"prv"|"iso"|
     "squiggle", "x", "y", "rot": 0|90|180|270, "node"}
    x, y of a symbol is its center; of a text, the middle of its top edge
    (align center) or its top left corner (align left).

Rules (plan section 4, Colin 2026-11-18):
- Left to right, meter at the far left on row 0. No scale: one column per
  Segment, one row per drop.
- The trunk (the longest path) runs along row 0. No other header ever sits
  on row 0 (item 9).
- A tap that ends at equipment is a vertical drop of one row; a tap that
  holds more junctions rises or drops to its own header row, at least two
  rows away so the parent's drops keep the row next to it. Up when its
  equipment sits above the tee, else down (elevation vertical).
- A Segment label sits on the Segment's LAST drawn line: centered above a
  horizontal line, right of and centered on a vertical one (item 4).
- An inline terminal (a header running straight into equipment) has its
  symbol turned 90 deg and its label to the right (item 7).
- Every equipment end has a PRV symbol on the pipe just before it.
- Any clash of a line, label or symbol with what is already placed moves
  the junction one column right, or the header one row further out.
"""

import gas_labels

# ---------------------------------------------------------------------------
# sizes, drafting view feet at 1:100 (3/32 in. text = 0.78 ft)
# ---------------------------------------------------------------------------
COL_FT = 10.0
ROW_FT = 9.0
TEXT_H_FT = 0.78
LINE_PITCH_FT = 1.3
CHAR_W_FT = 0.52
LABEL_GAP_FT = 0.4
LABEL_SIDE_FT = 0.6
EQUIP_HALF_W = 1.0          # RJA - P Symbols - Equipment, 0.242 in. wide
EQUIP_HALF_H = 0.56
CAP_HALF_W = 1.0
CAP_HALF_H = 0.5
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
NOTES_GAP_FT = 3.0
MAX_STRETCH = 12
MAX_ROWS_OUT = 12
EPS = 0.05

UTILITY_TEXT = ["GAS FROM UTILITY"]


def X(c):
    return c * COL_FT


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
        return _lines_clash(_pts(a), _pts(b))
    if a["t"] == "line" or b["t"] == "line":
        l, o = (a, b) if a["t"] == "line" else (b, a)
        if o["t"] == "symbol":
            return _line_vs_symbol(_pts(l), o["box"])
        return _seg_in_box(_pts(l), o["box"])
    return _boxes_overlap(a["box"], b["box"])


def _pts(l):
    return (l["x0"], l["y0"], l["x1"], l["y1"])


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

    # -- helpers -----------------------------------------------------------
    def fits(self, new, occ):
        for a in new:
            for b in occ:
                if clash(a, b):
                    return False
        for i in range(len(new)):
            for j in range(i + 1, len(new)):
                if clash(new[i], new[j]):
                    return False
        return True

    def _terms(self, br, out=None):
        out = [] if out is None else out
        end = br["end"]
        if end is None:
            return out
        if end["type"] == "terminal":
            out.append(end)
        else:
            self._terms(end["run"], out)
            for t in end["taps"]:
                self._terms(t, out)
        return out

    def _dir(self, tap, j):
        """+1 up when the tap's equipment sits above the tee, else -1."""
        jz = (j.get("xyz") or [0, 0, 0])[2]
        zs = [(t.get("xyz") or [0, 0, jz])[2] for t in self._terms(tap)]
        if not zs:
            return -1
        mz = sum(zs) / float(len(zs))
        return 1 if mz > jz + 0.5 else -1

    def _split(self, j, trunk):
        """Continuation along this header, and the taps in drawing order."""
        legs = list(j["taps"]) + [j["run"]]
        cont = j["run"]
        if trunk:
            for b in legs:
                if self.segs[b["segments"][0]]["on_longest"]:
                    cont = b
                    break
        taps = [b for b in legs if b is not cont]
        return cont, taps

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
                sym = _symbol("iso", ax + ux * t, ay + uy * t, rot,
                              ISO_HALF_L, ISO_HALF_T, v["node_id"])
                cuts.append((t, ISO_HALF_L))
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
        if rot == 0:
            prims.append(_text(mx, ay + LABEL_GAP_FT + h, lines, "center",
                               "seg", seg=sid))
        else:
            prims.append(_text(ax + LABEL_SIDE_FT, my + h / 2.0, lines, "left",
                               "seg", seg=sid))
        return prims

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

    # -- headers, taps, junctions -----------------------------------------
    def header(self, br, cx, r, occ, trunk, from_row=None, start_trim=0.0):
        """Lay br out from grid point (cx, r) to the right along row r. With
        from_row, the first Segment first runs vertically from (cx, from_row).
        Returns the new occupancy list, or None."""
        n = len(br["segments"])
        end = br["end"]
        for stretch in range(0, MAX_STRETCH + 1):
            pts_list = []
            for i in range(n):
                c0 = cx + i
                c1 = cx + i + 1 + (stretch if i == n - 1 else 0)
                pts = [(X(c0), Y(r)), (X(c1), Y(r))]
                if i == 0 and from_row is not None:
                    pts = [(X(cx), Y(from_row))] + pts
                pts_list.append(pts)
            ce = cx + n + stretch
            new = self._branch_prims(br, pts_list, (X(ce), Y(r)), (1.0, 0.0),
                                     start_trim)
            if not self.fits(new, occ):
                continue
            trial = occ + new
            if end["type"] == "terminal":
                if stretch:
                    self.log.append("stretch {0} col(s) before {1}".format(
                        stretch, end.get("name")))
                return trial
            res = self.junction(end, ce, r, trial, trunk)
            if res is not None:
                if stretch:
                    self.log.append("stretch {0} col(s) before junction "
                                    "{1}".format(stretch, end["node_id"]))
                return res
        return None

    def junction(self, j, c, r, occ, trunk):
        cont, taps = self._split(j, trunk)
        trial = occ
        for tap in taps:
            res = self.tap(tap, j, c, r, trial)
            if res is None:
                return None
            trial = res
        cont_trunk = trunk and self.segs[cont["segments"][0]]["on_longest"]
        return self.header(cont, c, r, trial, cont_trunk)

    def tap(self, tap, j, c, r, occ):
        d = self._dir(tap, j)
        end = tap["end"]
        n = len(tap["segments"])
        if end["type"] == "terminal":
            pts_list = []
            for i in range(n):
                pts_list.append([(X(c), Y(r + d * i)), (X(c), Y(r + d * (i + 1)))])
            new = self._branch_prims(tap, pts_list, (X(c), Y(r + d * n)),
                                     (0.0, float(d)))
            if self.fits(new, occ):
                return occ + new
            return None
        for k in range(2, MAX_ROWS_OUT + 1):
            rr = r + d * k
            if rr == 0:
                continue
            res = self.header(tap, c, rr, occ, False, from_row=r)
            if res is not None:
                if k > 2:
                    self.log.append("header {0} placed {1} rows out".format(
                        end["node_id"], k))
                return res
        return None


def layout(sch, notes=None, origin_kind="meter"):
    """Draw list for a schematic. notes: the notes block lines
    (gas_labels.notes_lines)."""
    eng = _Engine(sch)
    occ = []
    # meter and utility stub on the reserved meter row
    occ.append(_symbol("meter", 0.0, 0.0, 0, METER_R, METER_R, sch["origin"]))
    ux0 = -METER_R
    ux1 = -(METER_R + UTILITY_RUN_FT)
    occ.append(_line(ux0, 0.0, ux1, 0.0, None))
    occ.append(_line(ux1, 0.0, ux1, -UTILITY_DROP_FT, None))
    occ.append(_symbol("squiggle", ux1, -UTILITY_DROP_FT - 0.5, 90, 0.5, 0.4))
    occ.append(_text(ux1 - 1.0, -UTILITY_DROP_FT - 1.4, UTILITY_TEXT, "left",
                     "utility"))
    root = sch.get("root")
    if root is not None:
        res = eng.header(root, 0, 0, occ, True, start_trim=METER_R)
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
    return {"prims": occ, "bounds": (x0, y0, x1, y1), "log": eng.log}
