# -*- coding: ascii -*-
"""
oneline_schematic.py

One-Line rebuild step 3 (2026-11-18): the schematic tree, built from an
oneline_snapshot dict. Pure Python, no Revit import. ASCII source.

    sch = build(snap)

    sch["segments"]  {sid: Segment}   every drawn pipe run, own data
    sch["root"]      Branch           from the meter
    sch["warnings"]  [str]

    Segment  {"id", "eids" [pipe ids], "node_ids" [fittings passed],
              "start_node", "end_node", "size", "sizes_seen",
              "pipe_ft", "elbows", "tees", "length_ft", "mbh",
              "valves" [{"kind": "iso"|"prv", "node_id"}],
              "on_longest" bool, "branch_kind" "root"|"run"|"tap",
              "is_existing" bool, "label" [line1, line2]}
    Branch   {"kind" "root"|"run"|"tap", "segments" [sid], "end": End,
              "dir" [dx, dy, dz], "developed_ft" (to its farthest end)}
    End      {"type": "junction", "node_id", "xyz", "run": Branch,
              "taps": [Branch]}
           | {"type": "terminal", "node_id", "xyz", "kind" "equipment"|
              "cap"|"open", "name", "mbh", "label" [lines], "inline" bool,
              "equipment_prv" bool, "modeled_prv_id" id|None}

Rules (plan section 3.2, Colin 2026-11-18):
1. A Segment runs from a Junction (or the meter) to the next Junction or
   Terminal and owns exactly the pipes and fittings between them. No id is
   ever copied to another Segment (fixes items 1 and 2).
2. A Segment also splits where the pipe size changes, so it has one size.
3. length_ft = pipe + 5 ft per elbow passed + 3 ft per tee passed (Colin,
   2026-11-19; the length inside a fitting is not added). A tee's 3 ft goes
   on the Segment that ARRIVES at the tee: every path through the tee, straight
   on or into a branch, passes that Segment once, and each outgoing leg
   starts a new Segment, so it counts once. Same hop rule as
   pipe_graph._longest_run_from, so the segments on the longest path sum to
   the longest run.
4. mbh = cum_mbh of the Segment's first pipe (pipe_graph's number).
5. At a Junction the run is the leg collinear (within 10 deg) with the
   incoming leg; the rest are taps. No collinear leg, or more than one:
   the leg with the longest developed length is the run (fixes item 3).
6. Taps are drawn before the run, so the run's end is drawn last.
7. Every equipment Terminal shows a PRV before it (Colin, 2026-11-18); a
   modeled equipment PRV (not mid stream) maps to that same symbol and is
   not drawn again on the Segment.
"""

import math

import gas_labels
import shared_params


COLLINEAR_COS = math.cos(math.radians(10.0))
MIN_DIR_FT = 0.75


# ---------------------------------------------------------------------------
# small vector helpers
# ---------------------------------------------------------------------------

def _sub(a, b):
    return [a[0] - b[0], a[1] - b[1], a[2] - b[2]]


def _len(v):
    return math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])


def _unit(v):
    n = _len(v)
    if n <= 1e-9:
        return None
    return [v[0] / n, v[1] / n, v[2] / n]


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _dist(a, b):
    return _len(_sub(a, b))


def _far_end(edge, ref):
    """The end of a pipe farther from ref."""
    s, e = edge.get("start"), edge.get("end")
    if not s:
        return e
    if not e:
        return s
    return s if _dist(s, ref) >= _dist(e, ref) else e


# ---------------------------------------------------------------------------
# graph access
# ---------------------------------------------------------------------------

class _G(object):
    def __init__(self, snap):
        self.snap = snap
        self.nodes = snap["nodes"]
        self.edges = snap["edges"]
        self.out = {}
        for e in sorted(self.edges.values(), key=lambda x: x["id"]):
            if e.get("to") is not None:
                self.out.setdefault(str(e["from"]), []).append(("edge", e["id"]))
        for k, kids in (snap.get("children") or {}).items():
            for c in kids:
                self.out.setdefault(str(k), []).append(("child", c))
        self.path = set(int(i) for i in (snap.get("longest_path") or []))
        self.elbow_ft = float(snap.get("elbow_equiv_ft") or
                              shared_params.ELBOW_EQUIV_LENGTH_FT)
        self.tee_ft = float(snap.get("tee_equiv_ft") or
                            shared_params.TEE_EQUIV_LENGTH_FT)
        self.origin_phase = snap.get("origin_phase")

    def node(self, nid):
        return self.nodes[str(nid)]

    def edge(self, eid):
        return self.edges[str(eid)]

    def legs(self, nid):
        return list(self.out.get(str(nid), []))

    def is_tee(self, nid):
        """A fitting with 3 or more connectors that adds the tee equivalent
        length to a path through it (never an appliance or a PRV)."""
        n = self.node(nid)
        if "is_tee" in n:
            return bool(n["is_tee"])
        return bool(n.get("kind") == "tee" and not n.get("is_gas_fixture") and
                    not n.get("is_prv"))

    def next_node(self, hop):
        kind, hid = hop
        if kind == "edge":
            return self.edge(hid)["to"]
        return hid


def _valve_kind(node):
    if node.get("is_prv"):
        return "prv"
    if node.get("kind") != "fitting" or node.get("is_elbow") or \
            node.get("is_transition"):
        return None
    fam = (node.get("family") or "").lower()
    for kw in shared_params.ANSUL_VALVE_KEYWORDS:
        if kw in fam:
            return "ansul"
    for kw in shared_params.ISOLATION_VALVE_KEYWORDS:
        if kw in fam:
            return "iso"
    return None


# ---------------------------------------------------------------------------
# build
# ---------------------------------------------------------------------------

class _Builder(object):

    def __init__(self, snap):
        self.g = _G(snap)
        self.segments = {}
        self.warnings = []
        self._n = 0

    def _new_seg(self, start_node, kind):
        self._n += 1
        sid = "S{0}".format(self._n)
        seg = {"id": sid, "eids": [], "node_ids": [], "start_node": start_node,
               "end_node": None, "size": None, "sizes_seen": [],
               "pipe_ft": 0.0, "elbows": 0, "tees": 0,
               "length_ft": 0.0, "mbh": None, "valves": [],
               "on_longest": False, "branch_kind": kind,
               "is_existing": False, "label": None}
        self.segments[sid] = seg
        return seg

    def _close(self, seg, end_node):
        g = self.g
        seg["end_node"] = end_node
        seg["length_ft"] = seg["pipe_ft"] + seg["elbows"] * g.elbow_ft + \
            seg["tees"] * g.tee_ft
        if seg["mbh"] is None:
            seg["mbh"] = float(g.node(end_node).get("cum_mbh") or 0.0)
        if len(seg["sizes_seen"]) > 1:
            self.warnings.append("Segment {0} has more than one size: {1}".format(
                seg["id"], ", ".join(seg["sizes_seen"])))
        ids = set(seg["eids"])
        if ids:
            seg["on_longest"] = ids.issubset(g.path)
        else:
            seg["on_longest"] = (seg["start_node"] in g.path and
                                 end_node in g.path)
        seg["label"] = gas_labels.segment_label(seg["size"], seg["length_ft"],
                                                seg["mbh"])

    def _leg_dir(self, jid, hop):
        """Direction of a leg leaving junction jid, from a point at least
        MIN_DIR_FT away (transitions sit right at the tee)."""
        g = self.g
        j = g.node(jid)["xyz"]
        cur = hop
        for _i in range(6):
            kind, hid = cur
            if kind == "edge":
                p = _far_end(g.edge(hid), j)
                if p:
                    return _unit(_sub(p, j))
                nxt = g.edge(hid)["to"]
            else:
                nxt = hid
            if nxt is None:
                return None
            p = g.node(nxt)["xyz"]
            if p and _dist(p, j) >= MIN_DIR_FT:
                return _unit(_sub(p, j))
            legs = g.legs(nxt)
            if len(legs) != 1:
                return _unit(_sub(p, j)) if p else None
            cur = legs[0]
        return None

    def _in_dir(self, seg, jid):
        """Direction of travel arriving at junction jid along seg."""
        g = self.g
        j = g.node(jid)["xyz"]
        if seg["eids"]:
            p = _far_end(g.edge(seg["eids"][-1]), j)
        else:
            p = g.node(seg["start_node"])["xyz"]
        if not p:
            return None
        return _unit(_sub(j, p))

    def branch(self, start, hop, kind):
        """Walk one leg from node start along hop to its Junction or
        Terminal. Returns a Branch."""
        g = self.g
        br = {"kind": kind, "segments": [], "end": None, "dir": None,
              "developed_ft": 0.0}
        seg = self._new_seg(start, kind)
        br["segments"].append(seg["id"])
        prev = start
        visited = set([start])
        while True:
            k, hid = hop
            if k == "edge":
                e = g.edge(hid)
                seg["eids"].append(hid)
                seg["pipe_ft"] += float(e.get("length_ft") or 0.0)
                size = e.get("size")
                if size is not None and size not in seg["sizes_seen"]:
                    seg["sizes_seen"].append(size)
                if seg["size"] is None:
                    seg["size"] = size
                if seg["mbh"] is None:
                    seg["mbh"] = float(e.get("cum_mbh") or 0.0)
                if (g.origin_phase and e.get("phase") and
                        e.get("phase") != g.origin_phase):
                    seg["is_existing"] = True
                y = e["to"]
            else:
                y = hid
            if y in visited:
                self.warnings.append("Loop at element {0}".format(y))
                self._close(seg, y)
                br["end"] = self._terminal(y, "open")
                return br
            visited.add(y)
            node = g.node(y)
            if node.get("is_elbow"):
                seg["elbows"] += 1
            # The tee's 3 ft lands on the Segment arriving at it (see rule 3).
            if g.is_tee(y):
                seg["tees"] += 1

            if node.get("kind") == "fixture" or node.get("is_gas_fixture"):
                self._close(seg, y)
                br["end"] = self._terminal(y, None)
                return br

            legs = g.legs(y)
            if not legs:
                self._close(seg, y)
                br["end"] = self._terminal(y, "open")
                self.warnings.append("Open end at element {0}".format(y))
                return br
            if len(legs) > 1:
                self._close(seg, y)
                br["end"] = self._junction(y, seg)
                return br

            # pass through y
            vk = _valve_kind(node)
            if vk:
                seg["valves"].append({"kind": vk, "node_id": y,
                                      "midstream": bool(node.get("is_midstream_prv"))})
            nxt = legs[0]
            if nxt[0] == "edge":
                nsize = g.edge(nxt[1]).get("size")
                if seg["size"] is not None and nsize is not None and \
                        nsize != seg["size"]:
                    self._close(seg, y)
                    seg = self._new_seg(y, kind)
                    br["segments"].append(seg["id"])
                    prev = y
                    hop = nxt
                    continue
            seg["node_ids"].append(y)
            prev = y
            hop = nxt

    def _terminal(self, nid, forced):
        g = self.g
        n = g.node(nid)
        name = (n.get("name") or "").strip()
        mbh = float(n.get("load_mbh") or 0.0)
        if forced == "open":
            kind = "open"
            label = []
        elif gas_labels.is_future_name(name):
            kind = "cap"
            label = gas_labels.cap_label(name, mbh)
        else:
            kind = "equipment"
            label = gas_labels.fixture_label(name, mbh)
        return {"type": "terminal", "node_id": nid, "xyz": n.get("xyz"),
                "kind": kind, "name": name, "mbh": mbh, "label": label,
                "inline": False, "equipment_prv": kind == "equipment",
                "modeled_prv_id": None}

    def _junction(self, jid, in_seg):
        g = self.g
        vin = self._in_dir(in_seg, jid)
        legs = []
        for hop in g.legs(jid):
            b = self.branch(jid, hop, "tap")
            b["dir"] = self._leg_dir(jid, hop)
            _developed(b, self.segments)
            legs.append(b)
        run = None
        if vin is not None:
            col = [b for b in legs if b["dir"] is not None and
                   _dot(b["dir"], vin) >= COLLINEAR_COS]
            if len(col) == 1:
                run = col[0]
        if run is None:
            run = max(legs, key=lambda b: b["developed_ft"])
        taps = [b for b in legs if b is not run]
        # taps in a stable order: by angle about the incoming direction
        taps.sort(key=lambda b: (_tap_key(b, vin), b["segments"][0]))
        run["kind"] = "run"
        for sid in _branch_sids(run, self.segments, top_only=True):
            self.segments[sid]["branch_kind"] = "run"
        return {"type": "junction", "node_id": jid, "xyz": g.node(jid)["xyz"],
                "run": run, "taps": taps}

    def finish(self, root):
        _developed(root, self.segments)
        _mark_inline(root)
        _fold_equipment_prvs(root, self.segments)


def _tap_key(b, vin):
    d = b.get("dir")
    if d is None or vin is None:
        return 0.0
    # signed angle in plan from the incoming direction, then vertical last
    cross_z = vin[0] * d[1] - vin[1] * d[0]
    return -cross_z


def _branch_sids(br, segments, top_only=False):
    out = list(br["segments"])
    if top_only:
        return out
    end = br["end"]
    if end and end["type"] == "junction":
        out += _branch_sids(end["run"], segments)
        for t in end["taps"]:
            out += _branch_sids(t, segments)
    return out


def _developed(br, segments):
    own = sum(segments[s]["length_ft"] for s in br["segments"])
    end = br["end"]
    down = 0.0
    if end and end["type"] == "junction":
        kids = [end["run"]] + end["taps"]
        down = max(_developed(k, segments) for k in kids)
    br["developed_ft"] = own + down
    return br["developed_ft"]


def _mark_inline(br):
    end = br["end"]
    if end is None:
        return
    if end["type"] == "terminal":
        end["inline"] = br["kind"] in ("run", "root")
        return
    _mark_inline(end["run"])
    for t in end["taps"]:
        _mark_inline(t)


def _fold_equipment_prvs(br, segments):
    """A modeled PRV that is not mid stream, on the last pipe run into
    equipment, becomes that equipment's PRV symbol (drawn once)."""
    end = br["end"]
    if end is None:
        return
    if end["type"] == "terminal":
        if end["kind"] == "equipment":
            for sid in reversed(br["segments"]):
                seg = segments[sid]
                keep = []
                for v in seg["valves"]:
                    if v["kind"] == "prv" and not v.get("midstream") and \
                            end["modeled_prv_id"] is None:
                        end["modeled_prv_id"] = v["node_id"]
                    else:
                        keep.append(v)
                seg["valves"] = keep
        return
    _fold_equipment_prvs(end["run"], segments)
    for t in end["taps"]:
        _fold_equipment_prvs(t, segments)


def build(snap):
    b = _Builder(snap)
    origin = snap["origin_id"]
    legs = b.g.legs(origin)
    if len(legs) != 1:
        b.warnings.append("The meter has {0} outlets; expected 1.".format(
            len(legs)))
    root = None
    if legs:
        root = b.branch(origin, legs[0], "root")
        b.finish(root)
    existing = snap.get("existing")
    if existing:
        root_mbh = 0.0
        if root and root["segments"]:
            root_mbh = b.segments[root["segments"][0]]["mbh"] or 0.0
        existing = dict(existing)
        existing["label"] = gas_labels.segment_label(
            existing.get("size_nominal"), existing.get("length_ft"),
            float(existing.get("load_mbh") or 0.0) + root_mbh)
    sch = {"origin": origin, "root": root, "segments": b.segments,
           "existing": existing,
           "warnings": b.warnings,
           "origin_xyz": b.g.node(origin).get("xyz"),
           "levels": list(snap.get("levels") or []),
           "longest_run_ft": float(snap.get("longest_run_ft") or 0.0)}
    _balance_trunk_labels(sch)
    return sch


def balanced_feet(lengths, target):
    """Whole feet for each length, summing exactly to target (largest
    remainder: every length rounds down, the feet left over go to the
    largest fractions, ties to the earlier one). A real length never shows
    as 0."""
    floors = [int(math.floor(float(x))) for x in lengths]
    extra = int(target) - sum(floors)
    order = sorted(range(len(lengths)),
                   key=lambda i: (-(float(lengths[i]) - floors[i]), i))
    out = list(floors)
    for i in order[:max(0, extra)]:
        out[i] += 1
    for i, x in enumerate(lengths):
        if out[i] < 1 and float(x) > 0.0:
            out[i] = 1
    return out


def _balance_trunk_labels(sch):
    """The longest run's labels add up to the notes block TOTAL DEVELOPED
    LENGTH. Rounding each segment on its own showed the run 1 ft long
    (G3C 2026-11-20: 26 + 13 + 36 + 23 + 27 = 125 against 124). Branch
    labels off the longest run keep plain rounding."""
    trunk = trunk_segments(sch)
    items = [(s["length_ft"], s) for s in trunk]
    ex = sch.get("existing")
    if ex and ex.get("length_ft") is not None:
        items.insert(0, (float(ex["length_ft"]), None))
    if not items:
        return
    total = sum(float(x) for x, _s in items)
    target = gas_labels.feet(total) if total > 0 else 0
    feet = balanced_feet([x for x, _s in items], target)
    for n, (_x, seg) in zip(feet, items):
        if seg is None:
            # the existing run keeps its MBH line, only the length changes
            ex["label_ft"] = n
            ex["label"] = [gas_labels.pipe_label(ex.get("size_nominal"), n),
                           ex["label"][1]]
        else:
            seg["label_ft"] = n
            seg["label"] = gas_labels.segment_label(seg["size"], n,
                                                    seg["mbh"])


# ---------------------------------------------------------------------------
# reading the tree (used by layout, console and tests)
# ---------------------------------------------------------------------------

def walk(br, visit, depth=0):
    """Depth first in drawing order: each Junction's taps, then its run.
    visit(branch, depth) is called before the branch's children."""
    if br is None:
        return
    visit(br, depth)
    end = br["end"]
    if end and end["type"] == "junction":
        for t in end["taps"]:
            walk(t, visit, depth + 1)
        walk(end["run"], visit, depth)


def terminals_in_order(sch):
    out = []

    def _v(br, _d):
        e = br["end"]
        if e and e["type"] == "terminal":
            out.append(e)
    walk(sch["root"], _v)
    return out


def terminal_branch(sch, name):
    """The Branch ending at the terminal with this name."""
    found = []

    def _v(br, _d):
        e = br["end"]
        if e and e["type"] == "terminal" and e["name"] == name:
            found.append(br)
    walk(sch["root"], _v)
    return found[0] if found else None


def trunk_segments(sch):
    """Trunk Segments in order from the meter out (G2 2026-11-19: the
    console listed them in dict order)."""
    order = []
    if sch.get("root") is not None:
        walk(sch["root"], lambda br, d: order.extend(br["segments"]))
    seen = set(order)
    order += [sid for sid in sch["segments"] if sid not in seen]
    return [sch["segments"][sid] for sid in order
            if sch["segments"][sid]["on_longest"]]


def trunk_sum_ft(sch):
    """Trunk Segments plus, on a tie in, the entered existing run (Size Gas
    adds that length to the longest run)."""
    total = sum(s["length_ft"] for s in trunk_segments(sch))
    if sch.get("existing"):
        total += float(sch["existing"].get("length_ft") or 0.0)
    return total


def eid_counts(sch):
    counts = {}
    for s in sch["segments"].values():
        for e in s["eids"]:
            counts[e] = counts.get(e, 0) + 1
    return counts


def mbh_violations(sch, tol=0.01):
    """Segments whose MBH is larger than the segment before them on the
    path from the meter. Empty when MBH never rises away from the meter."""
    bad = []
    segs = sch["segments"]

    def _rec(br, upstream):
        prev = upstream
        for sid in br["segments"]:
            m = segs[sid]["mbh"]
            if prev is not None and m > prev + tol:
                bad.append(sid)
            prev = m
        end = br["end"]
        if end and end["type"] == "junction":
            for k in [end["run"]] + end["taps"]:
                _rec(k, prev)
    if sch["root"]:
        _rec(sch["root"], None)
    return bad
