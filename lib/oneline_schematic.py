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
              "trail" [pipe ends and fittings passed, xyz, in order],
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
import oneline_gate
import shared_params


COLLINEAR_COS = math.cos(math.radians(10.0))
MIN_DIR_FT = 0.75
# the incoming direction at a junction is taken from the last point passed
# at least this far from it (an elbow can sit 0.3 ft from the tee, review n47)
MIN_IN_FT = 0.25


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


def label_size(edge):
    """Size printed on a pipe: Size Gas's required size name when the drawn
    diameter matches it (CSST EHD-25 is not shown as EHD-23, which shares
    its diameter, Colin 2026-11-20), else the modeled size. Existing pipe
    always shows its modeled size (review n18)."""
    if not oneline_gate.is_existing_edge(edge) and             oneline_gate.diameter_matches(edge):
        return edge.get("required_size")
    return edge.get("size")


def _valve_kind(node):
    if node.get("is_prv"):
        return "prv"
    if node.get("kind") != "fitting" or node.get("is_elbow") or \
            node.get("is_transition"):
        return None
    # the one classifier (shared_params.accessory_kind). The walk sets
    # is_prv by the same rule, so a PRV name here was not taken as a PRV by
    # the walk: no symbol is invented for it.
    kind = shared_params.accessory_kind(node.get("family"))
    return None if kind == "prv" else kind


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
        seg = {"id": sid, "eids": [], "node_ids": [], "trail": [],
               "start_node": start_node,
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
        if not j:
            return None        # a joint with no location (pipe_graph)
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
        """Direction of travel arriving at junction jid along seg: from the
        last pipe end or fitting passed (an elbow after the last pipe
        counts, review n47), at least MIN_IN_FT from the junction."""
        g = self.g
        j = g.node(jid)["xyz"]
        if not j:
            return None
        start = g.node(seg["start_node"]).get("xyz")
        for p in reversed(list(seg.get("trail") or [])):
            if p and _dist(p, j) >= MIN_IN_FT:
                return _unit(_sub(j, p))
        if start and _dist(start, j) > 1e-9:
            return _unit(_sub(j, start))
        return None

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
                _add_trail(seg, e, g.node(prev).get("xyz"))
                size = label_size(e)
                if size is not None and size not in seg["sizes_seen"]:
                    seg["sizes_seen"].append(size)
                if seg["size"] is None:
                    seg["size"] = size
                if seg["mbh"] is None:
                    seg["mbh"] = float(e.get("cum_mbh") or 0.0)
                # existing = not created in the project's last phase, the
                # rule Size Gas uses (review n16), not the meter's phase
                if oneline_gate.is_existing_edge(e):
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
            if node.get("xyz"):
                seg["trail"].append(node["xyz"])
            vk = _valve_kind(node)
            if vk:
                seg["valves"].append({"kind": vk, "node_id": y,
                                      "midstream": bool(node.get("is_midstream_prv"))})
            nxt = legs[0]
            if nxt[0] == "edge":
                nsize = label_size(g.edge(nxt[1]))
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


def _add_trail(seg, edge, from_xyz):
    """Pipe ends in travel order: the end nearer the node it leaves first."""
    s, e = edge.get("start"), edge.get("end")
    ends = [p for p in (s, e) if p]
    if from_xyz and len(ends) == 2 and _dist(e, from_xyz) < _dist(s, from_xyz):
        ends = [e, s]
    seg["trail"].extend(ends)


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
           "origin_kind": snap.get("origin_kind") or "meter",
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
    # a 1 ft bump is taken back from the largest label(s), so the labels
    # still add up to target (review n48); ties give from the later label
    over = sum(out) - int(target)
    while over > 0:
        big = max(range(len(out)), key=lambda i: (out[i], i))
        if out[big] <= 1:
            break           # every label is already 1 ft: cannot balance
        out[big] -= 1
        over -= 1
    return out


def _balance_trunk_labels(sch):
    """The longest run's labels add up to the notes block TOTAL DEVELOPED
    LENGTH. Rounding each segment on its own showed the run 1 ft long
    (G3C 2026-11-20: 26 + 13 + 36 + 23 + 27 = 125 against 124). Branch
    labels off the longest run keep plain rounding.

    The target is gas_labels.feet(sch["longest_run_ft"]), the same number
    the notes print (snap["longest_run_ft"], which on a tie in already holds
    the entered existing length, pipe_graph.apply_existing_conditions; the
    existing run is one of the balanced labels). Only with no longest run
    is the target the rounded trunk sum. A target the labels cannot reach
    adds a warning to sch["warnings"]."""
    trunk = trunk_segments(sch)
    items = [(s["length_ft"], s) for s in trunk]
    ex = sch.get("existing")
    if ex and ex.get("length_ft") is not None:
        items.insert(0, (float(ex["length_ft"]), None))
    if not items:
        return
    total = sum(float(x) for x, _s in items)
    run_ft = float(sch.get("longest_run_ft") or 0.0)
    if run_ft > 0:
        target = gas_labels.feet(run_ft)
    else:
        target = gas_labels.feet(total) if total > 0 else 0
    feet = balanced_feet([x for x, _s in items], target)
    if sum(feet) != target:
        sch.setdefault("warnings", []).append(
            "Longest run labels add up to {0} ft, not the {1} ft TOTAL "
            "DEVELOPED LENGTH in the notes (trunk segments sum to {2:.1f} "
            "ft).".format(sum(feet), target, total))
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


def _tenths_split(values, target):
    """Whole tenths for each value (ft), summing exactly to target tenths
    (largest remainder, ties to the earlier one). Unlike balanced_feet no
    minimum is forced: a 0.0 ft run stays 0.0."""
    tenths = [float(v) * 10.0 for v in values]
    floors = [int(math.floor(t + 1e-9)) for t in tenths]
    extra = int(target) - sum(floors)
    order = sorted(range(len(values)),
                   key=lambda i: (-(tenths[i] - floors[i]), i))
    out = list(floors)
    if not out:
        return out
    if extra > 0:
        for i in order[:extra]:
            out[i] += 1
    elif extra < 0:
        # only when the target was held below the floors (kept for safety)
        for i in list(reversed(order))[:(-extra)]:
            out[i] = max(0, out[i] - 1)
    return out


def _node_names(sch):
    """node_id -> 'TEE' | fixture name, for naming a segment's two ends."""
    names = {}

    def _v(br, _d):
        e = br["end"]
        if e is None:
            return
        if e["type"] == "junction":
            names[e["node_id"]] = "TEE"
        else:
            names[e["node_id"]] = (e.get("name") or "").strip() or "OUTLET"
    if sch.get("root") is not None:
        walk(sch["root"], _v)
    return names


def developed_breakdown(sch, longest_run_ft=None, elbow_ft=None,
                        tee_ft=None):
    """The longest run as a table: pipe lengths per segment, then total
    pipe, elbows, tees and the developed length (Colin, 2026-11-20).

    Returns {"segments": [{"item", "tenths", "pipe_text", "from", "to",
    "size"}], "summary": [{"key", "item", "qty", "each", "length"}],
    "total_pipe_ft", "elbows", "tees", "elbow_ft", "tee_ft",
    "developed_ft" (as shown, tenths), "developed_whole", "farthest",
    "run_ft" (what the notes block uses), "mismatch" bool}

    Every figure shown adds exactly: the segment tenths sum to the total
    pipe length, and total pipe + elbows x 5 + tees x 3 is the developed
    length shown. The whole feet always equal gas_labels.feet of the run
    the notes block prints (longest_run_ft, else the rows' own sum): when
    plain rounding of the tenths would read 124.5 against a notes block of
    124, the total pipe is held one tenth lower.

    Tie in: the entered existing run is its own first row,
    METER TO NEW CONNECTION, EXISTING. An existing run has no fittings of
    its own to count (its length is the Size Gas entry).
    """
    elbow_ft = float(shared_params.ELBOW_EQUIV_LENGTH_FT
                     if elbow_ft is None else elbow_ft)
    tee_ft = float(shared_params.TEE_EQUIV_LENGTH_FT
                   if tee_ft is None else tee_ft)
    trunk = trunk_segments(sch)
    existing = sch.get("existing")
    names = _node_names(sch)

    # a name per trunk end node; repeats are numbered (TEE 1, TEE 2)
    kinds = []
    for s in trunk:
        nm = names.get(s["end_node"])
        kinds.append(nm if nm else "REDUCER")
    totals = {}
    for k in kinds:
        totals[k] = totals.get(k, 0) + 1
    seen = {}
    labels = []
    for k in kinds:
        if k in ("TEE", "REDUCER") and totals[k] > 1:
            seen[k] = seen.get(k, 0) + 1
            labels.append("{0} {1}".format(k, seen[k]))
        else:
            labels.append(k)

    start_name = "NEW CONNECTION" if existing else "METER"
    items = []
    if existing and existing.get("length_ft") is not None:
        items.append({"value": float(existing["length_ft"]),
                      "from": "METER", "to": "NEW CONNECTION, EXISTING",
                      "size": existing.get("size_nominal"),
                      "existing": True})
    prev = start_name
    for s, lab in zip(trunk, labels):
        items.append({"value": float(s["pipe_ft"]), "from": prev, "to": lab,
                      "size": s["size"], "existing": False})
        prev = lab
    for it in items:
        sz = gas_labels.pipe_label(it["size"], None) if it["size"] else ""
        if it["existing"]:
            head = "METER TO NEW CONNECTION, EXISTING"
        else:
            head = "{0} TO {1}".format(it["from"], it["to"])
        it["item"] = "{0}, {1}".format(head, sz) if sz else head

    elbows = sum(int(s["elbows"]) for s in trunk)
    tees = sum(int(s["tees"]) for s in trunk)
    elbow_total = elbows * elbow_ft
    tee_total = tees * tee_ft
    pipe_true = sum(it["value"] for it in items)
    run_true = pipe_true + elbow_total + tee_total
    run_ft = float(longest_run_ft) if longest_run_ft is not None else run_true
    want = gas_labels.feet(run_ft) if run_ft > 0 else 0
    elbow10 = int(round(elbow_total * 10.0))
    tee10 = int(round(tee_total * 10.0))
    equiv10 = elbow10 + tee10

    pipe10 = gas_labels._half_up(pipe_true * 10.0)
    for _i in range(12):
        have = gas_labels._half_up((pipe10 + equiv10) / 10.0)
        if run_ft <= 0 or have == want:
            break
        pipe10 += -1 if have > want else 1
    pipe10 = max(0, pipe10)
    dev10 = pipe10 + equiv10

    split = _tenths_split([it["value"] for it in items], pipe10)
    segs = []
    for it, t in zip(items, split):
        segs.append({"item": it["item"], "tenths": t,
                     "pipe_text": "{0:.1f}'".format(t / 10.0),
                     "from": it["from"], "to": it["to"], "size": it["size"]})

    def _f(t10):
        return "{0:.1f}'".format(t10 / 10.0)
    dev_whole = gas_labels._half_up(dev10 / 10.0)
    summary = [
        {"key": "pipe", "item": "TOTAL PIPE LENGTH", "qty": "", "each": "",
         "length": _f(pipe10)},
        {"key": "elbows", "item": "ELBOWS", "qty": str(elbows),
         "each": "{0:.0f}'".format(elbow_ft), "length": _f(elbow10)},
        {"key": "tees", "item": "TEES", "qty": str(tees),
         "each": "{0:.0f}'".format(tee_ft), "length": _f(tee10)},
        {"key": "developed", "item": "DEVELOPED LENGTH", "qty": "",
         "each": "",
         "length": "{0} ({1}')".format(_f(dev10), dev_whole)},
    ]
    farthest = None
    if trunk:
        farthest = names.get(trunk[-1]["end_node"])
        if farthest == "TEE":
            farthest = None
    return {"segments": segs, "summary": summary,
            "total_pipe_ft": pipe10 / 10.0, "elbows": elbows, "tees": tees,
            "elbow_ft": elbow10 / 10.0, "tee_ft": tee10 / 10.0,
            "developed_ft": dev10 / 10.0, "developed_whole": dev_whole,
            "farthest": farthest, "run_ft": run_ft,
            "mismatch": abs(run_true - run_ft) > 0.05}


def breakdown_table(bd):
    """(headers, rows) of a developed_breakdown for the console and the
    drafting view: segment rows, then the summary rows, the last (the
    developed length) shown as the totals row."""
    headers = ["ITEM", "QTY", "EQUIV EACH", "LENGTH"]
    rows = []
    for s in bd["segments"]:
        rows.append([s["item"], "", "", s["pipe_text"]])
    for r in bd["summary"]:
        rows.append([r["item"], r["qty"], r["each"], r["length"]])
    return headers, rows


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
