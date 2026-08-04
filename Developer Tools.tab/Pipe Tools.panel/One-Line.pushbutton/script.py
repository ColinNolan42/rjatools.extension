# -*- coding: ascii -*-
# One-Line.pushbutton/script.py
# Phase 2 - Gas Piping One-Line Diagram Generator
#
# Picks the gas meter, lets user pick the IFGC table (for notes),
# traverses the network, computes a schematic layout, and draws
# everything into a new Revit DraftingView.
#
# Layout rules:
#   - Trunk runs left-to-right (horizontal pipe length on x-axis).
#   - Branches go UP or DOWN based on fixture z-elevation vs. the meter.
#   - Vertical sections (risers/drops) advance y by LEVEL_HEIGHT (10 ft).
#   - Multiple branches at the same tee: same x, staggered y (Option A).
#
# IronPython 2.7 / PyRevit

import os
import sys
import math
import datetime

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    ElementTransformUtils,
    GraphicsStyle,
    Line,
    Arc,
    XYZ,
    ViewDrafting,
    ViewFamilyType,
    ViewFamily,
    TextNote,
    TextNoteType,
    TextRange,
    HorizontalTextAlignment,
    FilteredElementCollector,
    FamilySymbol,
    Transaction,
)
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List as _CSList

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

_lib_dir = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

import shared_params
import revit_helpers
import pipe_graph
import gas_tables
import sizing_engine
import ui_helpers


# ---------------------------------------------------------------------------
# Layout constants  (all in Revit feet = view coordinates at 1:100)
# ---------------------------------------------------------------------------
VIEW_SCALE       = 100      # DraftingView.Scale
LEVEL_HEIGHT     = 18.0    # ft vertical clearance per branch level (compressed)
MIN_SEGMENT_FT   = 10.0    # ft minimum horizontal segment so text doesn't overlap
SYMBOL_RADIUS    = 0.5     # ft  meter circle radius
FIXTURE_HW       = 1.5     # ft  half-width of 3-line fixture symbol (3 ft total)
FIXTURE_SPACING  = 0.7     # ft  gap between 3 fixture lines
FIXTURE_LABEL_GAP = 1.0   # ft  gap between outermost symbol line and nearest text edge
VALVE_HW          = 1.0     # ft  half-width of bowtie (2 ft total)
VALVE_HH          = 0.6     # ft  half-height of bowtie triangle
VALVE_GAP         = 0.3     # ft  gap between valve edge and fixture connection line (fallback)
VALVE_FIXTURE_GAP = 2.5     # ft  fixture baseline → nearest valve center
VALVE_PITCH       = 2.5     # ft  center-to-center between consecutive valve symbols
LABEL_ABOVE      = 1.2     # ft  above a horizontal pipe (must clear text height)
LABEL_RIGHT      = 0.6     # ft  right of a vertical pipe
UPSTREAM_H       = 6.0     # ft  horizontal stub left of meter
UPSTREAM_V       = 4.0     # ft  vertical drop of upstream stub
TEXT_HEIGHT_FT   = 0.78    # ft  3/32" x (100/12) at 1:100
TEXT_GAP         = TEXT_HEIGHT_FT * 2.0  # ft  between note lines
TEXT_CHAR_WIDTH_FT = TEXT_HEIGHT_FT * 0.6  # ft  approx glyph width for centering
NOTES_X_BASE     = -(UPSTREAM_H + SYMBOL_RADIUS + 2.0)
NOTES_Y_BASE     = LEVEL_HEIGHT + 4.0

ELBOW_EQUIV_FT   = 5.0  # ft per elbow, per IFGC A103.1 -- must match
                        # pipe_graph._find_longest_run's ELBOW_EQUIV_FT

_PRV_KW       = ("prv", "regulator", "regulating")
_ISOLATION_KW = ("valve", "ball", "gate", "check", "shutoff")
_VALVE_KW     = _PRV_KW + _ISOLATION_KW  # combined for legacy checks


# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------

def _node_z(graph, node_id, default=0.0):
    node = graph.nodes.get(node_id)
    if node and node.location_xyz:
        return float(node.location_xyz[2])
    return default


def _edges_from(graph, node_id):
    return [e for e in graph.edges.values() if e.from_node_id == node_id]


def _edge_developed_length(graph, edge):
    """Pipe length + 5ft if this edge's downstream node is an elbow.

    Matches pipe_graph._find_longest_run's elbow attribution: each is_elbow
    node is counted exactly once, via the single edge that terminates at it
    (the network is a tree, so every node has at most one parent edge).
    """
    length = edge.length_feet
    to_node = graph.nodes.get(edge.to_node_id)
    if to_node is not None and to_node.is_elbow:
        length += ELBOW_EQUIV_FT
    return length


def _trunk_edge_developed_lengths(graph, trunk_all_ids):
    """Map trunk edge_id -> developed length (pipe + elbow-equivalent).

    Walks path_element_ids in order. Each edge starts with
    _edge_developed_length (pipe length + 5ft if its own to_node is an
    elbow). Any node reached via node_children (zero-length connection,
    appears in the path with no preceding edge) that is itself an elbow
    has its 5ft equivalent added onto the most recently seen trunk edge --
    mirroring _find_longest_run's DFS, which counts every is_elbow node
    along the path exactly once regardless of how it was reached.
    """
    dev = {}
    last_edge_id = None
    pending_elbow_ft = 0.0
    for item in trunk_all_ids:
        if item in graph.edges:
            edge = graph.edges[item]
            dev[item] = _edge_developed_length(graph, edge) + pending_elbow_ft
            pending_elbow_ft = 0.0
            last_edge_id = item
        else:
            node = graph.nodes.get(item)
            if node is None:
                continue
            if (last_edge_id is not None
                    and graph.edges[last_edge_id].to_node_id == item):
                continue  # already counted via _edge_developed_length
            if node.is_elbow:
                if last_edge_id is not None:
                    dev[last_edge_id] += ELBOW_EQUIV_FT
                else:
                    pending_elbow_ft += ELBOW_EQUIV_FT
    return dev


def _find_fixture_z(graph, start_nid, trunk_set, default_z):
    """BFS downstream from start_nid (ignoring trunk edges) to find the first
    gas fixture node, then return its Revit z-elevation.

    This is used so branch direction (UP vs DOWN in the diagram) is driven by
    where the EQUIPMENT actually is, not where the first branch fitting is.
    A top-takeoff CSST stub that goes UP briefly before routing DOWN to a
    water heater at floor level will correctly read DOWN.

    Args:
        graph:      NetworkGraph
        start_nid:  node_id to start searching from
        trunk_set:  set of trunk edge element IDs (skip these)
        default_z:  z to return if no fixture is found

    Returns:
        float z-elevation in Revit feet
    """
    queue   = [start_nid]
    seen    = {start_nid}
    while queue:
        nid  = queue.pop(0)
        node = graph.nodes.get(nid)
        if node and node.is_gas_fixture:
            return _node_z(graph, nid, default_z)
        for edge in graph.edges.values():
            if edge.from_node_id == nid and edge.element_id not in trunk_set:
                if edge.to_node_id and edge.to_node_id not in seen:
                    seen.add(edge.to_node_id)
                    queue.append(edge.to_node_id)
        for child_nid in graph.node_children.get(nid, []):
            if child_nid not in seen:
                seen.add(child_nid)
                queue.append(child_nid)
    return default_z


def _branch_children(graph, nid, trunk_set, seen):
    """One-hop children of nid: real pipe edges (excluding trunk) plus
    zero-length node_children fitting-to-fitting connections, excluding
    anything already visited. Does not mutate seen -- the caller marks a
    child seen only once it actually commits to recursing into it.

    Returns a list of (child_nid, edge_len_ft, edge_ids).
    """
    out = []
    for edge in graph.edges.values():
        if (edge.from_node_id == nid
                and edge.element_id not in trunk_set
                and edge.to_node_id
                and edge.to_node_id not in seen):
            out.append((edge.to_node_id, _edge_developed_length(graph, edge),
                        [edge.element_id]))
    for child in graph.node_children.get(nid, []):
        if child not in seen:
            child_node = graph.nodes.get(child)
            extra = ELBOW_EQUIV_FT if (child_node and child_node.is_elbow) else 0.0
            out.append((child, extra, []))
    return out


def _build_tree(graph, nid, trunk_set, seen, seg_len, seg_eids, prv, iso):
    """Recursively build a tree rooted at nid.

    seg_len/seg_eids is the developed length/edge-ids of the single hop
    that led here from whichever node called this (the tee_candidate's
    branch edge for the very first call, or a real branch point for
    deeper calls) -- NOT a cumulative total from the very start.

    Passes straight through simple (single-child) fittings, accumulating
    length as it goes, and only creates a 'branch' tree node at a REAL
    branch point (2+ surviving non-trunk children). There is no depth
    limit: a cascade of any length is preserved intact as nested 'branch'
    nodes, and no fixture is ever dropped, unlike the old approach of
    flattening everything into one list with _trace_to_fixtures and
    trying to regroup it afterward.

    Returns one of:
      {"kind": "fixture", "fixture_nid":.., "seg_len":.., "seg_eids":..,
       "has_isolation":.., "has_prv":.., "cum_mbh":..}
      {"kind": "branch", "seg_len":.., "seg_eids":.., "has_isolation":..,
       "has_prv":.., "children": [tree,...]}
      None (dead end -- no fixture reachable this way)
    """
    node = graph.nodes.get(nid)
    if node is None:
        return None

    fname  = (node.family_name or "").lower()
    is_prv = any(kw in fname for kw in _PRV_KW)
    is_iso = (not is_prv) and any(kw in fname for kw in _ISOLATION_KW)
    prv = prv or is_prv
    iso = iso or is_iso

    if node.is_gas_fixture:
        return {"kind": "fixture", "fixture_nid": nid,
                "seg_len": seg_len, "seg_eids": seg_eids,
                "has_isolation": iso, "has_prv": prv,
                "cum_mbh": node.cumulative_load_mbh}

    children = _branch_children(graph, nid, trunk_set, seen)
    if not children:
        return None

    if len(children) == 1:
        child_nid, extra_len, extra_eids = children[0]
        seen.add(child_nid)
        return _build_tree(graph, child_nid, trunk_set, seen,
                            seg_len + extra_len, seg_eids + extra_eids,
                            prv, iso)

    subtrees = []
    for child_nid, extra_len, extra_eids in children:
        seen.add(child_nid)
        sub = _build_tree(graph, child_nid, trunk_set, seen,
                           extra_len, extra_eids, prv, iso)
        if sub is not None:
            subtrees.append(sub)

    if not subtrees:
        return None
    if len(subtrees) == 1:
        # Only one child actually led anywhere (the other(s) were dead
        # ends) -- not a real split, fold this hop's length into it.
        only = subtrees[0]
        only["seg_len"]  = seg_len + only["seg_len"]
        only["seg_eids"] = seg_eids + only["seg_eids"]
        return only

    return {"kind": "branch", "seg_len": seg_len, "seg_eids": seg_eids,
            "has_isolation": iso, "has_prv": prv, "children": subtrees}


def _tree_total_len(node):
    """Longest cumulative seg_len from this node down to any leaf fixture,
    inclusive of this node's own seg_len. Used only to rank slots by which
    one is longest (becomes the primary, continues straight) vs. the
    others (flat siblings, side-by-side)."""
    if node["kind"] == "fixture":
        return node["seg_len"]
    return node["seg_len"] + max(_tree_total_len(c) for c in node["children"])


def _tree_sum_mbh(node):
    """Sum of cum_mbh across every fixture reachable under this tree node.

    Used to label a new-main spine segment with the TOTAL load still
    flowing through it, before any of ITS OWN downstream taps are
    subtracted -- the same "cumulative MBH decreasing away from meter"
    convention the real trunk already uses for its own segments (see
    graph.edges[...].cumulative_load_mbh / pipe_graph._sum_load), just
    computed straight from the tree so it works even for a hop with no
    real pipe of its own to read a cumulative load off of.
    """
    if node["kind"] == "fixture":
        return node["cum_mbh"]
    return sum(_tree_sum_mbh(c) for c in node["children"])


def _tree_first_fixture(node):
    """Any one fixture_nid reachable under this node -- used only for the
    'already positioned by a sibling tee_candidate' dedupe check."""
    if node["kind"] == "fixture":
        return node["fixture_nid"]
    for c in node["children"]:
        f = _tree_first_fixture(c)
        if f is not None:
            return f
    return None


def _flatten_main(children, acc_len, acc_eids):
    """Flatten a list of sibling tree nodes (all hanging off the SAME
    real fork) into an ordered list of leaf (fixture-kind) tree dicts,
    each a shallow copy with "seg_len"/"seg_eids" REPLACED to mean "pipe
    since the PREVIOUS fixture in this flattened sequence" (not "since
    its own immediate parent fork").

    acc_len/acc_eids is whatever real pipe led INTO this fork from
    wherever the flattening started (0/[] for the very top-level call,
    since that hop is the riser, drawn separately by the caller).

    Order, confirmed against the real Fraser reference diagram for
    RTU-1..6 (drawn as ONE continuous flat main -- RTU-6, RTU-4, RTU-3,
    RTU-1, RTU-2, RTU-5, no nested "side main" row at all): any plain
    single-fixture sibling first, ascending by remaining length (reads
    as an ordinary tap -- shorter first, farther/"last tee" last);
    THEN any sibling that is itself a further branch (>= 2 fixtures,
    which by construction of _build_tree is any "branch"-kind child),
    in DESCENDING order of remaining length, each fully flattened in
    turn onto this SAME continuous sequence -- confirmed live: the
    longer {RTU-1,RTU-3} branch unrolls completely (itself ascending,
    RTU-3 then RTU-1) before the shorter {RTU-2,RTU-5} branch does the
    same, all six ending up as one flat sequence.

    Two or more children hanging off the exact same real fork all
    receive acc_len/acc_eids on the FIRST one only -- subsequent
    siblings reset to 0/[] since their own hop is measured from the
    shared fork itself, not from whichever sibling happened to be
    placed immediately before them in the flattened sequence. A
    "branch" sibling's OWN seg_len/seg_eids (the hop from this fork to
    IT) is folded into the accumulator handed to its children before
    recursing -- dropping this was a real bug: it silently zeroed out
    the developed length of everything past the first nested fork.
    """
    leaves = [c for c in children if c["kind"] == "fixture"]
    mains  = [c for c in children if c["kind"] == "branch"]
    out = []

    if not mains:
        ranked = sorted(children, key=_tree_total_len)
        for i, c in enumerate(ranked):
            a_len, a_eids = (acc_len, acc_eids) if i == 0 else (0.0, [])
            leaf = dict(c)
            leaf["seg_len"]  = a_len + c["seg_len"]
            leaf["seg_eids"] = a_eids + c["seg_eids"]
            out.append(leaf)
        return out

    ranked_leaves = sorted(leaves, key=_tree_total_len)
    for i, c in enumerate(ranked_leaves):
        a_len, a_eids = (acc_len, acc_eids) if i == 0 else (0.0, [])
        leaf = dict(c)
        leaf["seg_len"]  = a_len + c["seg_len"]
        leaf["seg_eids"] = a_eids + c["seg_eids"]
        out.append(leaf)

    ranked_mains = sorted(mains, key=_tree_total_len, reverse=True)
    for i, m in enumerate(ranked_mains):
        a_len, a_eids = (acc_len, acc_eids) if (i == 0 and not ranked_leaves) else (0.0, [])
        out.extend(_flatten_main(m["children"], a_len + m["seg_len"], a_eids + m["seg_eids"]))
    return out


def _emit_new_main(graph, tx, ty, tee_z, direc, depth, tree,
                    positions, branch_info, layout_log, spine_segments,
                    meter_z=0.0):
    """Draw an entire branch tree as ONE continuous sub-main ("new
    main"): a single riser up/down from the parent row to this main's
    own row, then a flat horizontal sequence of every fixture in it (see
    _flatten_main for the exact order), each with its own short stub to
    the equipment symbol -- real segment length + correctly decreasing
    cumulative MBH the whole way, exactly matching how the real Fraser
    reference diagram draws RTU-1..6.

    tree is the WHOLE tree node for this cascade (must be "branch" kind
    -- callers already handle the plain "fixture" case separately).
    """
    main_row_y = ty + direc * LEVEL_HEIGHT * depth

    if tree["seg_len"] > 0.01:
        spine_segments.append({
            "from_pos":  (tx, ty),
            "to_pos":    (tx, main_row_y),
            "eids":      tree["seg_eids"],
            "cum_mbh":   _tree_sum_mbh(tree),
            "length_ft": tree["seg_len"],
            "size":      "",
        })

    flat = _flatten_main(tree["children"], 0.0, [])
    tap_y = main_row_y + direc * LEVEL_HEIGHT
    cur_x = tx
    prev_x = tx

    for i, leaf in enumerate(flat):
        # Always advance, including the very first tap -- otherwise it
        # lands on the exact same x as the riser above, and the riser's
        # own label collides with this hop's label (both centered within
        # a few feet of tx). Still draw a plain connector for every hop so
        # the row reads as one continuous line even when there's nothing
        # new to label.
        cur_x += MIN_SEGMENT_FT
        # i==0's leaf["seg_len"] is that fixture's OWN perpendicular tap
        # stub (already reported on its own branch_info entry) -- the
        # riser already covers the real distance up to this first tee, so
        # there is never a real horizontal hop to label here, regardless
        # of leaf["seg_len"]. Only i>0 hops represent real pipe travel
        # between consecutive tees along this row.
        if i > 0 and leaf["seg_len"] > 0.01:
            remaining_mbh = sum(f["cum_mbh"] for f in flat[i:])
            spine_segments.append({
                "from_pos":  (prev_x, main_row_y),
                "to_pos":    (cur_x, main_row_y),
                "eids":      leaf["seg_eids"],
                "cum_mbh":   remaining_mbh,
                "length_ft": leaf["seg_len"],
                "size":      "",
            })
        else:
            spine_segments.append({
                "from_pos":  (prev_x, main_row_y),
                "to_pos":    (cur_x, main_row_y),
                "eids":      [],
                "cum_mbh":   None,   # sentinel: draw the line, skip the label
                "length_ft": 0.0,
                "size":      "",
            })

        positions[leaf["fixture_nid"]] = (cur_x, tap_y)
        branch_info.append({
            "tee_nid":         None,
            "tee_pos":         (cur_x, main_row_y),
            "fixture_nid":     leaf["fixture_nid"],
            "fixture_pos":     (cur_x, tap_y),
            "total_ft":        leaf["seg_len"],
            "branch_edge_ids": leaf["seg_eids"],
            "has_isolation":   leaf["has_isolation"],
            "has_prv":         leaf["has_prv"],
            "direc":           direc,
            "size":            "",
            "cum_mbh":         leaf["cum_mbh"],
            "sub_fixtures":    [],
        })
        layout_log.append({
            "tee_nid":    None,
            "tee_pos":    (cur_x, main_row_y),
            "tee_z":      tee_z,
            "edge_id":    leaf["seg_eids"][-1] if leaf["seg_eids"] else None,
            "to_nid":     leaf["fixture_nid"],
            "to_z":       _node_z(graph, leaf["fixture_nid"], meter_z),
            "fixture_z":  _node_z(graph, leaf["fixture_nid"], meter_z),
            "direc":      "UP" if direc > 0 else "DOWN",
            "result_pos": (cur_x, tap_y),
            "branch_y":   None,
        })
        prev_x = cur_x

    return cur_x


def _resolve_collisions(positions, branch_info, trunk_nodes, tee_candidates):
    """Expand the trunk diagram when an L-branch stub would overlap an adjacent branch.

    After layout, a sub-fixture stub from an L-branch may land within BRANCH_GAP
    of the next trunk tee's X.  The fix is to shift that tee (and every subsequent
    trunk node and branch) rightward, keeping the L-branch stubs in place.

    Iterates until no conflicts remain (handles cascading shifts).
    """
    BRANCH_GAP = FIXTURE_HW * 2.0 + 2.0  # ft min clearance between branch verticals

    trunk_candidate_nids = trunk_nodes | tee_candidates

    changed = True
    while changed:
        changed = False
        for bi in branch_info:
            if not bi.get("sub_fixtures"):
                continue
            src_tee_x = bi["tee_pos"][0]

            for sf in bi["sub_fixtures"]:
                stub_x = sf["stub_x"]

                for nid in trunk_candidate_nids:
                    if nid not in positions:
                        continue
                    nx = positions[nid][0]
                    if nx <= src_tee_x + 0.01:
                        continue
                    if abs(nx - stub_x) >= BRANCH_GAP:
                        continue

                    # Collision: shift this tee and everything to its right.
                    # The current L-branch stubs stay fixed; downstream tees move.
                    shift       = stub_x + BRANCH_GAP - nx
                    threshold_x = nx - 0.01

                    for node_id in trunk_candidate_nids:
                        if node_id in positions:
                            ox, oy = positions[node_id]
                            if ox >= threshold_x:
                                positions[node_id] = (ox + shift, oy)

                    for other_bi in branch_info:
                        if other_bi is bi:
                            continue
                        otx, oty = other_bi["tee_pos"]
                        if otx < threshold_x:
                            continue
                        other_bi["tee_pos"] = (otx + shift, oty)
                        ofx, ofy = other_bi["fixture_pos"]
                        other_bi["fixture_pos"] = (ofx + shift, ofy)
                        for osf in other_bi.get("sub_fixtures", []):
                            osf["stub_x"] += shift
                            sfn = osf["fixture_nid"]
                            if sfn in positions:
                                osx, osy = positions[sfn]
                                positions[sfn] = (osx + shift, osy)

                    changed = True
                    break
                if changed:
                    break
            if changed:
                break


def _compute_layout(graph):
    """Assign (x, y) view positions to every graph node via two-phase BFS.

    Phase 1: Walk path_element_ids (which interleaves node and edge IDs) to
             position every node on the main trunk, including direct fitting
             connections that have no pipe between them.
    Phase 2: BFS from ALL trunk nodes (not just from_node_ids of trunk edges)
             through both pipe edges and node_children to reach every branch
             node in the system.

    Returns:
        positions dict  {node_id: (x, y)}  in view feet
        trunk_set       set of pipe edge element IDs on the main trunk
        meter_nid       node_id of the gas meter
        meter_z         z-elevation of the meter in Revit feet
        layout_log      list of dicts recording every BFS branch decision
        spine_segments  list of dicts, one per "new main" pipe segment
                         drawn by _emit_new_main (the sub-main line
                         itself, distinct from any individual fixture tap)
    """
    # path_element_ids interleaves node and edge IDs:
    # [meter_nid, pipe1_id, node1_id, pipe2_id, node2_id, ..., fixture_nid]
    trunk_all_ids  = list(graph.longest_run["path_element_ids"])
    meter_nid      = trunk_all_ids[0]
    meter_z        = _node_z(graph, meter_nid)
    trunk_set      = set(eid for eid in trunk_all_ids if eid in graph.edges)

    positions = {meter_nid: (0.0, 0.0)}
    trunk_x   = 0.0
    prev_pos  = (0.0, 0.0)

    # ------------------------------------------------------------------
    # Phase 1: Walk the full trunk path (nodes AND pipe edges)
    # Apply MIN_SEGMENT_FT so labels never overlap adjacent trunk segments.
    # ------------------------------------------------------------------
    for item in trunk_all_ids[1:]:
        if item in graph.edges:
            edge     = graph.edges[item]
            from_pos = positions.get(edge.from_node_id, prev_pos)
            if edge.from_node_id not in positions:
                positions[edge.from_node_id] = from_pos
            from_z  = _node_z(graph, edge.from_node_id, meter_z)
            to_z    = _node_z(graph, edge.to_node_id,   meter_z)
            z_delta = to_z - from_z
            L       = max(edge.length_feet, 0.001)
            fx, fy  = from_pos
            if abs(z_delta) >= 0.5 * L:
                d = 1.0 if z_delta > 0 else -1.0
                new_pos = (fx, fy + d * LEVEL_HEIGHT)
            else:
                seg_len = max(edge.length_feet, MIN_SEGMENT_FT)
                trunk_x = fx + seg_len
                new_pos = (trunk_x, fy)
            positions[edge.to_node_id] = new_pos
            prev_pos = new_pos
        else:
            # Direct node-to-node connection (node_children) on the trunk
            positions[item] = prev_pos

    # Collect ALL trunk node IDs (from the path AND from edge endpoints)
    trunk_nodes = set()
    for item in trunk_all_ids:
        if item in graph.edges:
            e = graph.edges[item]
            trunk_nodes.add(e.from_node_id)
            trunk_nodes.add(e.to_node_id)
        else:
            trunk_nodes.add(item)

    # ------------------------------------------------------------------
    # Phase 2 (Simplified Schematic): For each trunk node with outgoing
    # branch edges, trace the entire branch to its fixture(s) and place
    # each fixture DIRECTLY above or below the trunk tee (same x).
    #
    # KEY: Branches in this Revit model start from node_children of trunk
    # tees (e.g. Transition fittings directly connected to a Tee), NOT from
    # the trunk tee node itself.  We must expand the tee candidate set to
    # include ALL node_children reachable from trunk nodes.
    # ------------------------------------------------------------------
    layout_log         = []
    branch_info        = []
    spine_segments     = []
    branch_counters    = {}
    trunk_fixture_nids = set()

    trunk_edge_ids_ordered = [i for i in trunk_all_ids if i in graph.edges]

    # Fixtures directly at the end of a trunk edge (no branch needed)
    for eid in trunk_edge_ids_ordered:
        edge    = graph.edges.get(eid)
        if edge is None:
            continue
        to_node = graph.nodes.get(edge.to_node_id)
        if to_node and to_node.is_gas_fixture:
            trunk_fixture_nids.add(edge.to_node_id)

    # Build tee_candidates: trunk nodes + their node_children (transitively).
    # Branches often start from Transition fittings that are node_children of
    # the Tee fitting, so we must include these child nodes.
    tee_candidates = set(trunk_nodes)
    worklist = list(trunk_nodes)
    while worklist:
        nid = worklist.pop()
        for child in graph.node_children.get(nid, []):
            if child not in tee_candidates:
                tee_candidates.add(child)
                # Give the child the same diagram position as its parent
                if nid in positions and child not in positions:
                    positions[child] = positions[nid]
                worklist.append(child)

    # For each candidate tee node, find outgoing non-trunk branches
    for tee_nid in tee_candidates:
        if tee_nid not in positions:
            continue
        tx, ty  = positions[tee_nid]
        tee_z   = _node_z(graph, tee_nid, meter_z)

        for branch_edge in _edges_from(graph, tee_nid):
            if branch_edge.element_id in trunk_set:
                continue
            if branch_edge.to_node_id is None:
                continue

            # Build the real tee tree downstream of this branch edge (no
            # flattening, no depth limit -- see _build_tree). Seed with the
            # branch edge's own developed length so an elbow right at the
            # first branch node is counted.
            branch_seed_len = _edge_developed_length(graph, branch_edge)
            seen = {branch_edge.to_node_id}
            tree = _build_tree(graph, branch_edge.to_node_id, trunk_set, seen,
                                branch_seed_len, [branch_edge.element_id],
                                False, False)
            if tree is None:
                continue
            # Skip if the primary fixture is already positioned by a sibling
            # tee_candidate (avoids duplicate branches from the same physical tee)
            first_fixture_nid = _tree_first_fixture(tree)
            if first_fixture_nid is None or positions.get(first_fixture_nid) is not None:
                continue

            # Direction from downstream fixture z vs this tee z
            first_fix_z = _node_z(graph, first_fixture_nid, meter_z)
            direc = 1.0 if first_fix_z > tee_z else -1.0

            depth = branch_counters.get(tee_nid, 0) + 1
            branch_counters[tee_nid] = depth

            if tree["kind"] == "fixture":
                # Simple single-fixture branch: no siblings anywhere down
                # this path.
                fix_y = ty + direc * LEVEL_HEIGHT * depth
                positions[tree["fixture_nid"]] = (tx, fix_y)
                branch_info.append({
                    "tee_nid":         tee_nid,
                    "tee_pos":         (tx, ty),
                    "fixture_nid":     tree["fixture_nid"],
                    "fixture_pos":     (tx, fix_y),
                    "total_ft":        tree["seg_len"],
                    "branch_edge_ids": tree["seg_eids"],
                    "has_isolation":   tree["has_isolation"],
                    "has_prv":         tree["has_prv"],
                    "direc":           direc,
                    "size":            "",
                    "cum_mbh":         tree["cum_mbh"],
                    "sub_fixtures":    [],
                })
                layout_log.append({
                    "tee_nid":    tee_nid,
                    "tee_pos":    (tx, ty),
                    "tee_z":      tee_z,
                    "edge_id":    branch_edge.element_id,
                    "to_nid":     tree["fixture_nid"],
                    "to_z":       first_fix_z,
                    "fixture_z":  first_fix_z,
                    "direc":      "UP" if direc > 0 else "DOWN",
                    "result_pos": (tx, fix_y),
                    "branch_y":   None,
                })
            else:
                # tree["children"] is always 2+ (that's what makes it a
                # 'branch' node) -- draw the whole cascade as ONE
                # continuous sub-main via _emit_new_main (a single riser
                # then a flat sequence of every fixture in it, no matter
                # how many real tees the cascade actually has).
                _emit_new_main(graph, tx, ty, tee_z, direc, depth, tree,
                                positions, branch_info, layout_log,
                                spine_segments, meter_z=meter_z)

    _resolve_collisions(positions, branch_info, trunk_nodes, tee_candidates)

    return (positions, trunk_set, meter_nid, meter_z,
            layout_log, branch_info, trunk_fixture_nids, spine_segments)


# ---------------------------------------------------------------------------
# Layout diagnostic formatter
# ---------------------------------------------------------------------------

def _format_layout_diagnostic(graph, positions, trunk_set, trunk_all_ids,
                               meter_nid, meter_z, layout_log, pipe_sizes):
    """Build a full multi-section diagnostic string for the terminal output.

    Mirrors the style of Diagnose and Size Gas debug output.
    """
    lines = []

    def row(s):
        lines.append(s)

    trunk_edge_ids = [i for i in trunk_all_ids if i in graph.edges]
    all_xy = list(positions.values())
    min_x = min(p[0] for p in all_xy) if all_xy else 0.0
    max_x = max(p[0] for p in all_xy) if all_xy else 0.0
    min_y = min(p[1] for p in all_xy) if all_xy else 0.0
    max_y = max(p[1] for p in all_xy) if all_xy else 0.0

    row("=== ONE-LINE LAYOUT DIAGNOSTIC ===")
    row("Meter node ID : {}  z = {:.2f} ft  diagram origin = (0.00, 0.00)".format(
        meter_nid, meter_z))
    row("Nodes positioned: {}/{}".format(len(positions), len(graph.nodes)))
    row("Trunk pipe edges: {}".format(len(trunk_edge_ids)))
    row("Diagram bounds  : x=[{:.1f}, {:.1f}]  y=[{:.1f}, {:.1f}]".format(
        min_x, max_x, min_y, max_y))
    row("")

    # ------ TRUNK PATH ------
    row("=== TRUNK PATH ({} edges) ===".format(len(trunk_edge_ids)))
    row(" {:>4}  {:>10}  {:>9}  {:>10}  {:>16}  {:>10}  {:>16}  {}".format(
        "idx", "edge_id", "length_ft", "from_nid", "from_pos",
        "to_nid", "to_pos", "type"))
    for i, eid in enumerate(trunk_edge_ids):
        edge     = graph.edges.get(eid)
        if edge is None:
            continue
        fp = positions.get(edge.from_node_id, (None, None))
        tp = positions.get(edge.to_node_id,   (None, None))
        fp_str = "({:6.1f},{:6.1f})".format(*fp) if fp[0] is not None else "UNPLACED"
        tp_str = "({:6.1f},{:6.1f})".format(*tp) if tp[0] is not None else "UNPLACED"
        from_z = _node_z(graph, edge.from_node_id, meter_z)
        to_z   = _node_z(graph, edge.to_node_id,   meter_z)
        z_delta = to_z - from_z
        L = max(edge.length_feet, 0.001)
        seg_type = "VERT" if abs(z_delta) >= 0.5 * L else "HORIZ"
        row(" {:>4}  {:>10}  {:>9.1f}  {:>10}  {:>16}  {:>10}  {:>16}  {}".format(
            i, eid, edge.length_feet,
            edge.from_node_id, fp_str,
            edge.to_node_id,   tp_str,
            seg_type))
    row("")

    # ------ BRANCH DECISIONS ------
    row("=== BRANCH DECISIONS (BFS log, {} entries) ===".format(len(layout_log)))
    row(" {:>10}  {:>14}  {:>8}  {:>10}  {:>8}  {:>8}  {:>10}  {:>12}  {:>14}".format(
        "tee_nid", "tee_pos", "tee_z", "edge_id", "to_nid",
        "to_z", "fixture_z", "direction", "result_pos"))
    for entry in layout_log:
        rp  = entry["result_pos"]
        tp  = entry["tee_pos"]
        fz  = entry.get("fixture_z", entry.get("to_z", 0.0))
        row(" {:>10}  {:>14}  {:>8.2f}  {:>10}  {:>8}  {:>8.2f}  {:>10.2f}  {:>12}  {:>14}".format(
            entry["tee_nid"],
            "({:.1f},{:.1f})".format(tp[0], tp[1]),
            entry["tee_z"],
            entry["edge_id"],
            entry["to_nid"],
            entry["to_z"],
            fz,
            entry["direc"],
            "({:.1f},{:.1f})".format(rp[0], rp[1])))
    row("")

    # ------ ALL NODES ------
    row("=== ALL NODES ({} total, {} positioned) ===".format(
        len(graph.nodes), len(positions)))
    row(" {:>10}  {:>8}  {:>24}  {:>8}  {:>8}  {:>28}  {:>14}".format(
        "node_id", "type", "family", "fixture", "MBH",
        "revit_xyz", "diagram_pos"))
    for nid, node in sorted(graph.nodes.items()):
        pos  = positions.get(nid)
        pos_str  = "({:.1f},{:.1f})".format(*pos) if pos else "UNPOSITIONED"
        xyz_str  = "({:.1f},{:.1f},{:.1f})".format(*node.location_xyz) \
                   if node.location_xyz else "-"
        fam  = (node.family_name or "-")[:22]
        row(" {:>10}  {:>8}  {:>24}  {:>8}  {:>8}  {:>28}  {:>14}".format(
            nid,
            (node.node_type or "-")[:8],
            fam,
            "YES" if node.is_gas_fixture else "no",
            "{:.1f}".format(node.gas_load_mbh),
            xyz_str[:28],
            pos_str))
    row("")

    # ------ ALL EDGES (drawing plan) ------
    row("=== ALL EDGES - DRAWING PLAN ({} edges) ===".format(len(graph.edges)))
    row(" {:>10}  {:>10}  {:>10}  {:>9}  {:>8}  {:>14}  {:>14}  {:>10}  {}".format(
        "edge_id", "from_nid", "to_nid", "length_ft", "MBH",
        "from_pos", "to_pos", "label?", "size"))
    for eid, edge in sorted(graph.edges.items()):
        if edge.to_node_id is None:
            continue
        fp  = positions.get(edge.from_node_id)
        tp  = positions.get(edge.to_node_id)
        fp_str = "({:.1f},{:.1f})".format(*fp) if fp else "UNPLACED"
        tp_str = "({:.1f},{:.1f})".format(*tp) if tp else "UNPLACED"
        nom  = pipe_sizes.get(eid, "")
        if fp is None or tp is None:
            label_flag = "SKIP(unpositioned)"
        elif edge.length_feet < 5.0:
            label_flag = "NO(<5ft)"
        else:
            label_flag = "YES"
        row(" {:>10}  {:>10}  {:>10}  {:>9.1f}  {:>8.1f}  {:>14}  {:>14}  {:>10}  {}".format(
            eid,
            edge.from_node_id or "-",
            edge.to_node_id or "-",
            edge.length_feet,
            edge.cumulative_load_mbh,
            fp_str, tp_str,
            label_flag,
            nom or "unsized"))

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Pipe size reading
# ---------------------------------------------------------------------------

_STEEL_NOMINALS = frozenset([
    "1/2","3/4","1","1-1/4","1-1/2","2","2-1/2","3","4","5","6","8","10","12"])


def _read_pipe_sizes(graph):
    """Read nominal pipe sizes from Revit model.

    Builds the inverse map in two passes so standard steel sizes always win
    over EHD/copper designations for the same diameter value.

    Returns {edge_element_id: nominal_size_str} for all readable edges.
    """
    inv = {}
    # Pass 1: non-steel (EHD, K&L, ACR, PE) -- populate but can be overwritten
    for nom, inches in sizing_engine.NOMINAL_TO_INCHES.items():
        if nom not in _STEEL_NOMINALS:
            inv[inches] = nom
    # Pass 2: standard steel -- overwrite any EHD/copper that shares the same inch value
    for nom, inches in sizing_engine.NOMINAL_TO_INCHES.items():
        if nom in _STEEL_NOMINALS:
            inv[inches] = nom
    sizes = {}
    for eid, edge in graph.edges.items():
        if edge.pipe is None:
            continue
        try:
            p = edge.pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            if p:
                dia_in  = p.AsDouble() * 12.0
                closest = min(inv.keys(), key=lambda k: abs(k - dia_in))
                if abs(closest - dia_in) < 0.1:
                    sizes[eid] = inv[closest]
        except Exception:
            pass
    return sizes


# ---------------------------------------------------------------------------
# Annotation symbol helpers
# ---------------------------------------------------------------------------

def _get_annotation_symbol(doc, family_name):
    """Return the first FamilySymbol whose Family.Name matches family_name."""
    for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if s.Family and s.Family.Name == family_name:
                return s
        except Exception:
            pass
    return None


def _activate_sym(doc, sym):
    """Activate a FamilySymbol if not already active. Returns sym or None."""
    if sym is None:
        return None
    if sym.IsActive:
        return sym
    t_act = Transaction(doc, "Activate " + (sym.Family.Name or "Symbol"))
    t_act.Start()
    try:
        sym.Activate()
        t_act.Commit()
        return sym
    except Exception:
        t_act.RollBack()
        return None


def _place_sym(doc, view, sym, x, y, rotate_90=False, rotate_180=False):
    """Place a generic annotation instance at (x, y) in the view.

    If rotate_90 is True, rotate the instance 90 degrees around the Z-axis
    through the placement point (used for symbols hanging off a horizontal
    trunk run, so the 3 lines run vertically).

    If rotate_180 is True, rotate the instance 180 degrees around the
    placement point instead. The "RJA - P Symbols - Equipment" family's
    own geometry always hangs BELOW its insertion point (confirmed via
    BoundingBoxXYZ against the live model: insertion point sits exactly on
    the box's Max.Y, i.e. the TOP line of the 3-line detail is the one
    nearest the connection). For a branch that takes off vertically UP
    from the main, the connection must land on the BOTTOM line instead, so
    the symbol needs a 180 degree flip to hang ABOVE its insertion point.
    Confirmed live: after RotateElement by pi about the insertion point,
    the box's Min.Y (not Max.Y) lands on the insertion point.

    rotate_90 and rotate_180 are mutually exclusive; rotate_90 wins if both
    are passed. Returns the placed element or None.
    """
    if sym is None:
        return None
    try:
        inst = doc.Create.NewFamilyInstance(XYZ(x, y, 0), sym, view)
        if inst is not None:
            axis = Line.CreateBound(XYZ(x, y, 0), XYZ(x, y, 1))
            if rotate_90:
                ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.pi / 2.0)
            elif rotate_180:
                ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.pi)
        return inst
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Schematic branch drawing
# ---------------------------------------------------------------------------

def _trunk_fixture_valves(graph, fix_nid, trunk_all_ids):
    """Return (has_prv, has_isolation) for a trunk-endpoint fixture.

    Scans the full trunk path (trunk_all_ids, which includes both edge IDs
    and node IDs) for valve-type nodes within 10 items upstream of fix_nid.
    Valve fittings are NODES in the path, not edges, so this function must
    search the ordered list rather than trunk_set (pipe edge IDs only).
    """
    has = [False, False]  # [has_prv, has_iso]  (IronPython 2.7: no nonlocal)

    def _check(nid):
        n = graph.nodes.get(nid)
        if not n:
            return
        fname = (n.family_name or "").lower()
        if any(kw in fname for kw in _PRV_KW):
            has[0] = True
        elif any(kw in fname for kw in _ISOLATION_KW):
            has[1] = True

    # Find where fix_nid first appears in the path (as a node, or as the
    # to_node / node_child-of-to_node of an edge).
    idx = None
    for i, item in enumerate(trunk_all_ids):
        if item == fix_nid:
            idx = i
            break
        if item in graph.edges:
            e = graph.edges[item]
            if e.to_node_id == fix_nid:
                idx = i
                break
            if fix_nid in graph.node_children.get(e.to_node_id, []):
                idx = i
                break

    if idx is None:
        return (False, False)

    # Inspect all items in the 10-element window immediately before idx.
    # This covers the valve fitting nodes that precede the fixture in the path.
    for item in trunk_all_ids[max(0, idx - 10): idx]:
        _check(item)
        if item in graph.edges:
            e = graph.edges[item]
            for child in graph.node_children.get(e.to_node_id, []):
                _check(child)
            for child in graph.node_children.get(e.from_node_id, []):
                _check(child)

    return (has[0], has[1])


def _draw_schematic_branch(doc, view, tee_x, tee_y, fix_x, fix_y,
                            direc, total_ft, size, has_isolation, has_prv,
                            fixture_node, tt_id,
                            valve_sym=None, prv_sym=None, equip_sym=None,
                            cap_sym=None, line_style=None):
    """Draw one simplified schematic branch from trunk tee to fixture.

    Diagram order from branch tee toward fixture:
      branch pipe -> [PRV] -> [isolation valve] -> 3-line fixture
    If only PRV: PRV takes the isolation slot (adjacent to fixture).
    If only isolation: isolation adjacent to fixture.
    If both: isolation adjacent to fixture, PRV one valve-width beyond.

    Uses project annotation families when available:
      valve_sym  -- RJA - P Symbols - Gate Valve  (rotated 90 degrees)
      prv_sym    -- RJA - P Symbols - Pressure Regulating Valve  (rotated 90 degrees)
      equip_sym  -- RJA - P Symbols - Equipment
      cap_sym    -- RJA - Equipment Cap, used instead of equip_sym when the
                    fixture is named "FUTURE" (see _is_future_fixture)
    """
    # Vertical segment from tee to fixture level
    _line(doc, view, tee_x, tee_y, tee_x, fix_y, line_style=line_style)

    # Horizontal segment if fixture is offset from tee
    if abs(fix_x - tee_x) > 0.01:
        _line(doc, view, tee_x, fix_y, fix_x, fix_y, line_style=line_style)

    going_up = fix_y > tee_y
    sign     = 1.0 if going_up else -1.0

    # Isolation valve: VALVE_FIXTURE_GAP ft from fixture baseline
    iso_y = fix_y - sign * VALVE_FIXTURE_GAP
    if has_isolation:
        v_inst = _place_sym(doc, view, valve_sym, tee_x, iso_y, rotate_90=True)
        if v_inst is None:
            _make_group(doc, _draw_valve_bowtie(doc, view, tee_x, iso_y))

    # PRV: one VALVE_PITCH slot further from fixture than isolation (or takes iso slot if alone)
    if has_prv:
        prv_y = (iso_y - sign * VALVE_PITCH) if has_isolation else iso_y
        p_inst = _place_sym(doc, view, prv_sym, tee_x, prv_y, rotate_90=True)
        if p_inst is None:
            _make_group(doc, _draw_valve_bowtie(doc, view, tee_x, prv_y))

    # Label on right side of the vertical segment
    mid_y = (tee_y + fix_y) / 2.0
    mbh_val  = int(round(fixture_node.gas_load_mbh)) if fixture_node else 0
    lft_rounded = int(round(total_ft))
    if size and lft_rounded > 0:
        lbl_line1 = '{}"G, {} FT'.format(size, lft_rounded)
    elif size:
        lbl_line1 = '{}\"G'.format(size)
    elif lft_rounded > 0:
        lbl_line1 = "{} FT".format(lft_rounded)
    else:
        lbl_line1 = None
    lbl = (lbl_line1 + "\n" + "{} MBH".format(mbh_val)) if lbl_line1 else "{} MBH".format(mbh_val)
    lbl_x = tee_x + (VALVE_HW + LABEL_RIGHT if (has_isolation or has_prv) else LABEL_RIGHT)
    _note(doc, view, lbl_x, mid_y, lbl, tt_id)

    # Equipment symbol at the fixture endpoint (cap symbol if FUTURE fixture)
    _place_equipment_symbol(doc, view, equip_sym, cap_sym, fix_x, fix_y, sign,
                             fixture_node.fixture_name if fixture_node else None,
                             rotate_180=going_up)

    # Fixture name + MBH as a SEPARATE TextNote (not grouped with symbol).
    # TextNote origin = top of text box; text flows DOWN in model space.
    # For going-up branches the origin must clear the outer line by one text
    # block height so the text bottom (not origin) sits above the symbol.
    if fixture_node:
        name     = fixture_node.fixture_name or "UNNAMED"
        label    = name + "\n" + "{} MBH".format(int(round(fixture_node.gas_load_mbh)))
        far_y    = fix_y + sign * 2 * FIXTURE_SPACING
        if sign > 0:  # going up: shift origin up by gap + 2 text lines
            lbl_y = far_y + FIXTURE_LABEL_GAP + 2 * TEXT_HEIGHT_FT
        else:         # going down: origin just below outer line, text flows further down
            lbl_y = far_y - FIXTURE_LABEL_GAP
        _note(doc, view, fix_x, lbl_y, label, tt_id, center_align=True,
              underline_len=len(name))


def _draw_schematic_branch_with_stubs(doc, view, bi, graph, tt_id,
                                       valve_sym=None, prv_sym=None, equip_sym=None,
                                       cap_sym=None, line_style=None):
    """Draw a branch where multiple fixtures share one branch off the trunk.

    The primary fixture (longest pipe path) hangs at the end of the main
    vertical.  Each shorter path branches off as a horizontal stub at the
    junction midpoint, keeping all fixture labels spatially separated.

    Valve order from tee toward fixture: PRV -> isolation -> 3-line fixture.
    """
    tee_x, tee_y  = bi["tee_pos"]
    fix_x, fix_y  = bi["fixture_pos"]
    has_isolation  = bi.get("has_isolation", False)
    has_prv        = bi.get("has_prv", False)
    sub_fixtures   = bi.get("sub_fixtures", [])
    size           = bi.get("size", "")
    total_ft       = bi.get("total_ft", 0)
    cum_mbh        = bi.get("cum_mbh", 0)
    going_up       = fix_y > tee_y
    sign           = 1.0 if going_up else -1.0

    # Main vertical: trunk tee to primary fixture level
    _line(doc, view, tee_x, tee_y, tee_x, fix_y, line_style=line_style)

    # Isolation valve: VALVE_FIXTURE_GAP ft from primary fixture baseline
    iso_y = fix_y - sign * VALVE_FIXTURE_GAP
    if has_isolation:
        v_inst = _place_sym(doc, view, valve_sym, tee_x, iso_y, rotate_90=True)
        if v_inst is None:
            _make_group(doc, _draw_valve_bowtie(doc, view, tee_x, iso_y))

    # PRV: one VALVE_PITCH slot further from fixture than isolation
    if has_prv:
        prv_y = (iso_y - sign * VALVE_PITCH) if has_isolation else iso_y
        p_inst = _place_sym(doc, view, prv_sym, tee_x, prv_y, rotate_90=True)
        if p_inst is None:
            _make_group(doc, _draw_valve_bowtie(doc, view, tee_x, prv_y))

    # L-shaped stubs for sub-fixtures: horizontal at junction_y then
    # vertical down to fixture_y (same level as primary fixture).
    for sf in sub_fixtures:
        jy    = sf["junction_y"]
        fy    = sf.get("fixture_y", fix_y)
        sx    = sf["stub_x"]
        sfnd  = graph.nodes.get(sf["fixture_nid"])
        sf_iso = sf.get("has_isolation", False)
        sf_prv = sf.get("has_prv", False)

        # Horizontal leg at junction_y
        _line(doc, view, tee_x, jy, sx, jy, line_style=line_style)
        # Vertical leg from junction to fixture level
        _line(doc, view, sx, jy, sx, fy, line_style=line_style)

        # Valves on stub vertical, VALVE_FIXTURE_GAP from stub fixture baseline
        stub_height = abs(fy - jy)
        sf_gap = min(VALVE_FIXTURE_GAP, stub_height * 0.5)
        sf_iso_y = fy - sign * sf_gap
        if sf_iso:
            sv_ins = _place_sym(doc, view, valve_sym, sx, sf_iso_y, rotate_90=True)
            if sv_ins is None:
                _make_group(doc, _draw_valve_bowtie(doc, view, sx, sf_iso_y))
        if sf_prv:
            sf_prv_y = (sf_iso_y - sign * VALVE_PITCH) if sf_iso else sf_iso_y
            sp_ins = _place_sym(doc, view, prv_sym, sx, sf_prv_y, rotate_90=True)
            if sp_ins is None:
                _make_group(doc, _draw_valve_bowtie(doc, view, sx, sf_prv_y))

        # Equipment symbol at fixture level (same Y as primary; cap if FUTURE)
        _place_equipment_symbol(doc, view, equip_sym, cap_sym, sx, fy, sign,
                                 sfnd.fixture_name if sfnd else None,
                                 rotate_180=going_up)

        if sfnd:
            name  = sfnd.fixture_name or "UNNAMED"
            label = name + "\n{} MBH".format(int(round(sfnd.gas_load_mbh)))
            far_y = fy + sign * 2 * FIXTURE_SPACING
            if sign > 0:
                lbl_y = far_y + FIXTURE_LABEL_GAP + 2 * TEXT_HEIGHT_FT
            else:
                lbl_y = far_y - FIXTURE_LABEL_GAP
            _note(doc, view, sx, lbl_y, label, tt_id, center_align=True,
                  underline_len=len(name))

        # Pipe label above the horizontal leg — uses remaining_ft (sub-tee to
        # stub fixture only, not the full path from trunk tee).
        sf_size = sf.get("remaining_size", "") or sf.get("size", "")
        sf_lft  = int(round(sf.get("remaining_ft", 0)))
        sf_mbh  = int(round(sfnd.gas_load_mbh)) if sfnd else 0
        if sf_size and sf_lft > 0:
            sub_l1 = '{}"G, {} FT'.format(sf_size, sf_lft)
        elif sf_size:
            sub_l1 = '{}"G'.format(sf_size)
        elif sf_lft > 0:
            sub_l1 = '{} FT'.format(sf_lft)
        else:
            sub_l1 = None
        if sub_l1:
            mid_x      = (tee_x + sx) / 2.0
            stub_lbl_y = jy + sign * LABEL_ABOVE
            _note(doc, view, mid_x, stub_lbl_y,
                  sub_l1 + "\n{} MBH".format(sf_mbh), tt_id, width=True)

    # Primary fixture symbol at end of main vertical (cap if FUTURE)
    primary_node = graph.nodes.get(bi["fixture_nid"])
    _place_equipment_symbol(doc, view, equip_sym, cap_sym, fix_x, fix_y, sign,
                             primary_node.fixture_name if primary_node else None,
                             rotate_180=going_up)

    if primary_node:
        name  = primary_node.fixture_name or "UNNAMED"
        label = name + "\n{} MBH".format(int(round(primary_node.gas_load_mbh)))
        far_y = fix_y + sign * 2 * FIXTURE_SPACING
        if sign > 0:
            lbl_y = far_y + FIXTURE_LABEL_GAP + 2 * TEXT_HEIGHT_FT
        else:
            lbl_y = far_y - FIXTURE_LABEL_GAP
        _note(doc, view, fix_x, lbl_y, label, tt_id, center_align=True,
              underline_len=len(name))

    # Two pipe labels on the main vertical when sub-fixtures are present:
    #   Label 1 (trunk tee → sub-tee junction): shared_ft + combined cum_mbh
    #   Label 2 (sub-tee junction → primary fixture): remaining_ft + primary MBH
    junction_y = sub_fixtures[0]["junction_y"] if sub_fixtures else None
    if junction_y is not None:
        shared_size   = bi.get("shared_size", "") or size
        shared_ft     = int(round(bi.get("shared_ft", 0)))
        combined_mbh  = int(round(cum_mbh))
        rem_size      = bi.get("remaining_size", "") or size
        rem_ft        = int(round(bi.get("remaining_ft", 0)))
        prim_mbh      = int(round(primary_node.gas_load_mbh)) if primary_node else 0

        # Label 1: on the shared section (tee_y → junction_y)
        mid1 = (tee_y + junction_y) / 2.0
        if shared_size and shared_ft > 0:
            l1 = '{}"G, {} FT'.format(shared_size, shared_ft)
        elif shared_size:
            l1 = '{}"G'.format(shared_size)
        elif shared_ft > 0:
            l1 = '{} FT'.format(shared_ft)
        else:
            l1 = None
        if l1:
            _note(doc, view, tee_x + LABEL_RIGHT, mid1,
                  l1 + "\n{} MBH".format(combined_mbh), tt_id)

        # Label 2: on the individual section (junction_y → fix_y)
        mid2 = (junction_y + fix_y) / 2.0
        if rem_size and rem_ft > 0:
            l2 = '{}"G, {} FT'.format(rem_size, rem_ft)
        elif rem_size:
            l2 = '{}"G'.format(rem_size)
        elif rem_ft > 0:
            l2 = '{} FT'.format(rem_ft)
        else:
            l2 = None
        if l2:
            _note(doc, view, tee_x + LABEL_RIGHT, mid2,
                  l2 + "\n{} MBH".format(prim_mbh), tt_id)
    else:
        # Fallback: single label if no sub-fixtures (shouldn't happen here)
        mid_y = (tee_y + fix_y) / 2.0
        lft   = int(round(total_ft))
        mbh   = int(round(cum_mbh))
        if size and lft > 0:
            l1 = '{}"G, {} FT'.format(size, lft)
        elif size:
            l1 = '{}"G'.format(size)
        elif lft > 0:
            l1 = '{} FT'.format(lft)
        else:
            l1 = None
        if l1:
            _note(doc, view, tee_x + LABEL_RIGHT, mid_y,
                  l1 + "\n{} MBH".format(mbh), tt_id)


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------

def _get_line_style(doc, style_name):
    """Return the GraphicsStyle with the given name, or None if not found."""
    for gs in FilteredElementCollector(doc).OfClass(GraphicsStyle):
        try:
            if gs.Name == style_name:
                return gs
        except Exception:
            pass
    return None


def _build_new_construction_eids(graph, doc):
    """Return set of edge element IDs whose pipe Phase Created name contains 'new'."""
    new_eids = set()
    for eid, edge in graph.edges.items():
        if edge.pipe is None:
            continue
        try:
            p = edge.pipe.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if p is None:
                continue
            phase_el = doc.GetElement(p.AsElementId())
            if phase_el is None:
                continue
            if "new" in (phase_el.Name or "").lower():
                new_eids.add(eid)
        except Exception:
            pass
    return new_eids


def _line(doc, view, x0, y0, x1, y1, line_style=None):
    """Draw a detail line and return the element (or None on failure).

    If line_style is a GraphicsStyle, it is applied to the detail curve.
    Pass None to use the view's default line style (thin).
    """
    try:
        if abs(x1 - x0) < 0.001 and abs(y1 - y0) < 0.001:
            return None
        dc = doc.Create.NewDetailCurve(
            view,
            Line.CreateBound(XYZ(x0, y0, 0), XYZ(x1, y1, 0)))
        if dc is not None and line_style is not None:
            try:
                dc.LineStyle = line_style
            except Exception:
                pass
        return dc
    except Exception:
        return None


def _note(doc, view, x, y, text, tt_id, width=None, center_align=False,
          underline_len=None):
    """Create a TextNote and return the element (or None on failure).

    If width is truthy, the insertion point is shifted left by roughly half
    the text's rendered width so the note appears horizontally centered on x
    (char-count approximation).

    If center_align is True, sets HorizontalTextAlignment.Center on the note
    so Revit centers the text on x directly -- more accurate than the width
    shift. Don't combine both; use one or the other.

    If underline_len is a positive int, the first underline_len characters
    of text are underlined via FormattedText -- the same result as
    selecting the text in Revit's text editor and clicking Underline.
    Used for equipment/fixture name labels (name is line 1, MBH is line 2;
    only the name gets underlined).
    """
    try:
        if width:
            max_chars = max(len(line) for line in text.split("\n"))
            x = x - (max_chars * TEXT_CHAR_WIDTH_FT) / 2.0
        tn = TextNote.Create(doc, view.Id, XYZ(x, y, 0), text, tt_id)
        if center_align and tn is not None:
            try:
                tn.HorizontalAlignment = HorizontalTextAlignment.Center
            except Exception:
                pass
        if underline_len and tn is not None:
            try:
                ft = tn.GetFormattedText()
                ft.SetUnderlineStatus(TextRange(0, underline_len), True)
                tn.SetFormattedText(ft)
            except Exception:
                pass
        return tn
    except Exception:
        return None


def _make_group(doc, elements):
    """Group a list of drawn elements into a Revit Detail Group."""
    ids = _CSList[ElementId]()
    for e in elements:
        if e is not None:
            try:
                ids.Add(e.Id)
            except Exception:
                pass
    if ids.Count > 1:
        try:
            doc.Create.NewGroup(ids)
        except Exception:
            pass


def _draw_meter_symbol(doc, view, cx, cy, tt_id, meter_sym=None):
    """Place meter symbol. Uses RJA - P Symbols - Meter when available."""
    inst = _place_sym(doc, view, meter_sym, cx, cy)
    if inst is None:
        # Fallback: draw circle + "M" label
        try:
            arc = Arc.Create(
                XYZ(cx, cy, 0),
                SYMBOL_RADIUS,
                0.0,
                2.0 * math.pi - 0.001,
                XYZ(1, 0, 0),
                XYZ(0, 1, 0))
            doc.Create.NewDetailCurve(view, arc)
        except Exception:
            pass
        _note(doc, view, cx - 0.15, cy - 0.25, "M", tt_id)


def _draw_upstream_stub(doc, view, cx, cy, squiggle_sym, tt_id, line_style=None):
    """Draw horizontal stub + vertical drop + squiggle (rotated 90 deg)."""
    sx = cx - SYMBOL_RADIUS
    # Horizontal stub going left from meter
    _line(doc, view, sx, cy, sx - UPSTREAM_H, cy, line_style=line_style)
    # Vertical drop to utility
    tip_x = sx - UPSTREAM_H
    _line(doc, view, tip_x, cy, tip_x, cy - UPSTREAM_V, line_style=line_style)
    # Squiggle at tip -- rotated 90 degrees so it reads vertically
    _place_sym(doc, view, squiggle_sym, tip_x, cy - UPSTREAM_V, rotate_90=True)
    # "GAS FROM UTILITY" label below the squiggle
    _note(doc, view, tip_x - 0.5, cy - UPSTREAM_V - 1.6, "GAS FROM\nUTILITY", tt_id)


def _draw_pipe_segment(doc, view, x0, y0, x1, y1, edge, pipe_sizes, tt_id,
                        force_label=False):
    _line(doc, view, x0, y0, x1, y1)

    # Skip very short pipes unless this is a trunk segment (force_label=True).
    if edge.length_feet < 5.0 and not force_label:
        return

    nom  = pipe_sizes.get(edge.element_id, "")
    lft  = int(round(edge.length_feet))
    mbh  = int(round(edge.cumulative_load_mbh))
    is_h = abs(y1 - y0) <= abs(x1 - x0)

    # Label: size + length on line 1, MBH on line 2 (no decimals)
    if nom:
        line1 = '{}"G, {} FT'.format(nom, lft)
    else:
        line1 = "{} FT".format(lft)
    label = line1 + "\n" + "{} MBH".format(mbh)

    if is_h:
        lx = (x0 + x1) / 2.0
        ly = max(y0, y1) + LABEL_ABOVE
        _note(doc, view, lx, ly, label, tt_id, width=True)
    else:
        lx = max(x0, x1) + LABEL_RIGHT
        ly = (y0 + y1) / 2.0
        _note(doc, view, lx, ly, label, tt_id)


def _is_future_fixture(fixture_name):
    """True if a fixture is a placeholder stub for future equipment (capped,
    no equipment connected yet) rather than a real, currently-served fixture."""
    return (fixture_name or "").strip().upper() == "FUTURE"


def _place_equipment_symbol(doc, view, equip_sym, cap_sym, cx, cy, sign,
                             fixture_name, rotate_90=False, rotate_180=False):
    """Place the equipment symbol at a fixture endpoint -- or, for a fixture
    named "FUTURE", the cap symbol instead (a capped stub reserved for
    reconnection in a future work phase, no equipment installed yet).

    Falls back to a drawn primitive when the corresponding project family
    isn't loaded: the normal 3-line equipment detail, or a squared-off "U"
    bracket for a FUTURE stub (open end centered on the pipe termination
    point, closed end pointing away from the trunk) -- rotate_90 orients
    the fallback lines vertically (trunk approaches horizontally);
    otherwise they're horizontal.
    """
    is_future = _is_future_fixture(fixture_name)
    sym  = cap_sym if is_future else equip_sym
    inst = _place_sym(doc, view, sym, cx, cy, rotate_90=rotate_90, rotate_180=rotate_180)
    if inst is not None:
        return

    if is_future:
        depth = 2 * FIXTURE_SPACING
        if rotate_90:
            far_x = cx + sign * depth
            elems = [
                _line(doc, view, cx, cy - FIXTURE_HW, far_x, cy - FIXTURE_HW),
                _line(doc, view, cx, cy + FIXTURE_HW, far_x, cy + FIXTURE_HW),
                _line(doc, view, far_x, cy - FIXTURE_HW, far_x, cy + FIXTURE_HW),
            ]
        else:
            far_y = cy + sign * depth
            elems = [
                _line(doc, view, cx - FIXTURE_HW, cy, cx - FIXTURE_HW, far_y),
                _line(doc, view, cx + FIXTURE_HW, cy, cx + FIXTURE_HW, far_y),
                _line(doc, view, cx - FIXTURE_HW, far_y, cx + FIXTURE_HW, far_y),
            ]
    else:
        elems = []
        for i in range(3):
            if rotate_90:
                xx = cx + sign * i * FIXTURE_SPACING
                elems.append(_line(doc, view, xx, cy - FIXTURE_HW, xx, cy + FIXTURE_HW))
            else:
                yy = cy + sign * i * FIXTURE_SPACING
                elems.append(_line(doc, view, cx - FIXTURE_HW, yy, cx + FIXTURE_HW, yy))
    _make_group(doc, elems)


def _draw_fixture_symbol(doc, view, cx, cy, going_up, node, tt_id):
    """Draw 3-line equipment symbol and label. Returns elements for grouping.

    Line 0 (outer, at cy) connects to the branch. Lines 1 and 2 extend
    AWAY from the trunk so the symbol hangs in the correct direction.
    """
    sign = 1.0 if going_up else -1.0
    elems = []
    for i in range(3):
        yy = cy + sign * i * FIXTURE_SPACING
        elems.append(_line(doc, view, cx - FIXTURE_HW, yy, cx + FIXTURE_HW, yy))

    name   = node.fixture_name or "UNNAMED"
    label  = name + "\n" + "{} MBH".format(int(round(node.gas_load_mbh)))
    far_y  = cy + sign * 2 * FIXTURE_SPACING
    ly     = far_y + sign * LABEL_ABOVE
    elems.append(_note(doc, view, cx, ly, label, tt_id, center_align=True))
    return elems


def _draw_valve_bowtie(doc, view, cx, cy):
    """Draw bowtie valve symbol. Returns elements for grouping."""
    elems = []
    for pts in [
        [(cx - VALVE_HW, cy + VALVE_HH),
         (cx - VALVE_HW, cy - VALVE_HH),
         (cx, cy)],
        [(cx + VALVE_HW, cy + VALVE_HH),
         (cx + VALVE_HW, cy - VALVE_HH),
         (cx, cy)],
    ]:
        for i in range(3):
            p1 = pts[i]
            p2 = pts[(i + 1) % 3]
            elems.append(_line(doc, view, p1[0], p1[1], p2[0], p2[1]))
    return elems


def _parse_pressure_drop(table_label):
    """Extract pressure drop string from a TABLE_OPTIONS label.

    Returns e.g. '0.3" W.C.' or '1 PSI', or '' if not found.
    """
    import re
    lbl = table_label or ""
    m = re.search(r'([\d.]+)"?\s*w\.c\.', lbl, re.IGNORECASE)
    if m:
        return '{}" W.C.'.format(m.group(1))
    m = re.search(r'([\d.]+)\s*psi\s*drop', lbl, re.IGNORECASE)
    if m:
        return '{} PSI'.format(m.group(1))
    return ""


def _draw_notes_block(doc, view, table_id, inlet_psi,
                      total_mbh, total_developed_ft, tt_id, notes_x, notes_y,
                      table_label=""):
    """Draw the 5-line notes block as a SINGLE multi-line TextNote."""
    pressure_drop = _parse_pressure_drop(table_label)
    if pressure_drop:
        loss_line = "MAX PRESSURE LOSS OF {} PER IFGC TABLE {}".format(
            pressure_drop, table_id)
    else:
        loss_line = "MAX PRESSURE LOSS PER IFGC TABLE {}".format(table_id)
    text = "\n".join([
        "CONTRACTOR SHALL SUBMIT APPLICATIONS TO UTILITY"
        " AND COORDINATE NEW METER SERVICE",
        "GAS PIPING SIZED FOR {} PSI".format(inlet_psi),
        loss_line,
        "TOTAL CONNECTED LOAD: {} MBH".format(int(round(total_mbh))),
        "TOTAL DEVELOPED LENGTH: {}'".format(int(round(total_developed_ft))),
    ])
    _note(doc, view, notes_x, notes_y, text, tt_id)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    output.print_md("# Gas Piping One-Line Diagram")
    output.print_md("---")

    revit_helpers.clear_log()

    # ------------------------------------------------------------------
    # STEP 1 - Pick gas meter
    # ------------------------------------------------------------------
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Select the gas meter element"
        )
        selected_element = doc.GetElement(ref.ElementId)
    except Exception:
        output.print_md("Selection cancelled.")
        return

    if selected_element is None:
        forms.alert("Could not retrieve selected element.",
                    title="One-Line - Selection Error")
        return

    output.print_md("**Selected:** Element ID {}".format(
        revit_helpers.eid_int(selected_element.Id)))

    # ------------------------------------------------------------------
    # STEP 2 - Validate meter
    # ------------------------------------------------------------------
    validation = revit_helpers.validate_selected_element(selected_element)
    if not validation["is_valid"]:
        forms.alert(
            "Please select the gas meter element.\n\n{}".format(
                validation["reason"]),
            title="One-Line - Invalid Selection"
        )
        return

    output.print_md(":white_check_mark: Meter validation passed.")

    # ------------------------------------------------------------------
    # STEP 3 - Select pipe material and IFGC table (populates notes block)
    # ------------------------------------------------------------------
    pipe_material, selected_table_label, elevation_ft = ui_helpers.show_table_picker(
        "One-Line - Select IFGC Table")
    if not pipe_material or not selected_table_label or elevation_ft is None:
        output.print_md("Cancelled at table selection. No changes made.")
        return

    selected_opt       = gas_tables.get_table_option_by_material_and_short_label(
        pipe_material, selected_table_label)
    table_id           = selected_opt["table_id"]
    inlet_pressure_psi = selected_opt["inlet_pressure_psi"]
    altitude_factor    = sizing_engine.altitude_derate_factor(elevation_ft)
    output.print_md("**Material:** {}  |  **Table:** {}  |  **Elevation:** {:.0f} ft".format(
        pipe_material, table_id, elevation_ft))

    # ------------------------------------------------------------------
    # STEP 4 - Traverse network
    # ------------------------------------------------------------------
    output.print_md("**Traversing network...**")
    try:
        graph = pipe_graph.build_network(selected_element, doc)
    except Exception as ex:
        forms.alert("Traversal failed:\n\n{}".format(str(ex)),
                    title="One-Line - Traversal Error")
        output.print_md(":cross_mark: {}".format(str(ex)))
        return

    output.print_md(":white_check_mark: {} nodes, {} pipe segments.".format(
        len(graph.nodes), len(graph.edges)))

    if graph.longest_run is None:
        forms.alert("Could not determine longest run. Run Diagnose first.",
                    title="One-Line - Error")
        return

    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
    if not fixture_nodes:
        forms.alert("No gas fixtures found. Cannot generate one-line diagram.",
                    title="One-Line - Error")
        return

    total_mbh          = sum(n.gas_load_mbh for n in fixture_nodes)
    longest_ft         = graph.longest_run["total_length_feet"]  # used for IFGC sizing
    total_developed_ft = longest_ft  # longest run (pipe + elbow equiv) per IFGC A103.1

    # ------------------------------------------------------------------
    # STEP 5 - Compute layout
    # ------------------------------------------------------------------
    output.print_md("**Computing layout...**")
    (positions, trunk_set, meter_nid, meter_z,
     layout_log, branch_info, trunk_fixture_nids,
     spine_segments) = _compute_layout(graph)
    n_positioned = len(positions)
    n_total      = len(graph.nodes)
    output.print_md(":white_check_mark: {}/{} nodes positioned.".format(
        n_positioned, n_total))

    # Compute diagram bounding box for notes placement and diagnostics
    all_xy   = list(positions.values())
    min_x_d  = min(p[0] for p in all_xy) if all_xy else 0.0
    max_x_d  = max(p[0] for p in all_xy) if all_xy else 0.0
    min_y_d  = min(p[1] for p in all_xy) if all_xy else 0.0
    max_y_d  = max(p[1] for p in all_xy) if all_xy else 0.0

    # Notes: left of upstream stub, above all diagram content
    notes_x = -(UPSTREAM_H + SYMBOL_RADIUS + 4.0)
    notes_y = max_y_d + 6.0

    # ------------------------------------------------------------------
    # STEP 6 - Read pipe sizes from model (needed for full diagnostic)
    # ------------------------------------------------------------------
    pipe_sizes  = _read_pipe_sizes(graph)
    sized_count = len(pipe_sizes)
    output.print_md("**Pipe sizes read:** {} of {} segments.".format(
        sized_count, len(graph.edges)))

    # Fill dominant size into each branch_info entry now that pipe_sizes is known
    for bi in branch_info:
        for eid in bi.get("branch_edge_ids", []):
            s = pipe_sizes.get(eid, "")
            if s:
                bi["size"] = s
                break
        # For multi-fixture branches: fill shared_size and remaining_size
        if bi.get("sub_fixtures"):
            for eid in bi.get("shared_eids", []):
                s = pipe_sizes.get(eid, "")
                if s:
                    bi["shared_size"] = s
                    break
            for eid in bi.get("remaining_eids", []):
                s = pipe_sizes.get(eid, "")
                if s:
                    bi["remaining_size"] = s
                    break
        for sf in bi.get("sub_fixtures", []):
            for eid in sf.get("branch_edge_ids", []):
                s = pipe_sizes.get(eid, "")
                if s:
                    sf["size"] = s
                    break
            for eid in sf.get("remaining_eids", []):
                s = pipe_sizes.get(eid, "")
                if s:
                    sf["remaining_size"] = s
                    break

    # Same dominant-size fill for the "new main" spine segments themselves
    # (the sub-main pipe between taps -- see _emit_new_main).
    for seg in spine_segments:
        for eid in seg.get("eids", []):
            s = pipe_sizes.get(eid, "")
            if s:
                seg["size"] = s
                break

    # Fallback for stub / side-takeoff branches whose pipe element wasn't written
    # by the sizing engine (e.g. short tee stubs, bottom take-offs).  Look up the
    # minimum IFGC size that handles the branch MBH demand at the system length.
    def _pick_fallback_size(demand_mbh, pairs):
        """Smallest nominal size in pairs whose capacity covers demand_mbh,
        altitude-derated per sizing_engine.altitude_derate_factor() so this
        fallback matches the same comparison Size Gas uses. Returns "" if
        no size in pairs is sufficient."""
        if altitude_factor > 0:
            demand_cfh_effective = demand_mbh / altitude_factor
        else:
            demand_cfh_effective = float("inf")
        for nom, cap in pairs:
            if cap is not None and cap >= demand_cfh_effective:
                return nom
        return ""

    try:
        _fb_sizes = gas_tables.list_pipe_sizes(table_id)
        _, _fb_caps = gas_tables.get_length_row(table_id, total_developed_ft)
        _fb_pairs = list(zip(_fb_sizes, _fb_caps))
    except Exception:
        _fb_pairs = []

    for bi in branch_info:
        if not bi["size"] and _fb_pairs:
            nom = _pick_fallback_size(bi["cum_mbh"], _fb_pairs)
            if nom:
                bi["size"] = nom
        if bi.get("sub_fixtures"):
            if not bi.get("shared_size") and _fb_pairs:
                nom = _pick_fallback_size(bi["cum_mbh"], _fb_pairs)
                if nom:
                    bi["shared_size"] = nom
            if not bi.get("remaining_size") and _fb_pairs:
                demand = bi.get("cum_mbh", 0) - sum(
                    s["cum_mbh"] for s in bi["sub_fixtures"])
                nom = _pick_fallback_size(demand, _fb_pairs)
                if nom:
                    bi["remaining_size"] = nom
        for sf in bi.get("sub_fixtures", []):
            if not sf.get("size") and _fb_pairs:
                nom = _pick_fallback_size(sf["cum_mbh"], _fb_pairs)
                if nom:
                    sf["size"] = nom
            if not sf.get("remaining_size") and _fb_pairs:
                nom = _pick_fallback_size(sf["cum_mbh"], _fb_pairs)
                if nom:
                    sf["remaining_size"] = nom

    for seg in spine_segments:
        if not seg["size"] and _fb_pairs and seg["cum_mbh"] is not None:
            nom = _pick_fallback_size(seg["cum_mbh"], _fb_pairs)
            if nom:
                seg["size"] = nom

    # ------------------------------------------------------------------
    # STEP 5c - Phase-based line style
    # ------------------------------------------------------------------
    new_construction_eids = _build_new_construction_eids(graph, doc)
    wide_line_style = _get_line_style(doc, "Line 4")
    if wide_line_style is None:
        output.print_md(":warning: Line style 'Line 4' not found - pipe segments will use default line weight.")
    else:
        output.print_md(":white_check_mark: Line style 'Line 4' loaded ({} New Construction segments).".format(
            len(new_construction_eids)))

    # ------------------------------------------------------------------
    # STEP 5b - Full layout diagnostic (copy/paste into conversation)
    # ------------------------------------------------------------------
    trunk_all_ids_diag = list(graph.longest_run["path_element_ids"])
    diag_text = _format_layout_diagnostic(
        graph, positions, trunk_set, trunk_all_ids_diag,
        meter_nid, meter_z, layout_log, pipe_sizes)
    output.print_md("---")
    output.print_md("## Layout Diagnostic - Copy and paste below this line")
    output.print_html(
        "<pre style='font-family:monospace;font-size:11px;'>{}</pre>".format(
            diag_text.replace("&", "&amp;")
                     .replace("<", "&lt;")
                     .replace(">", "&gt;")))
    output.print_md("---")
    output.print_md("**Pipe sizes read:** {} of {} segments.".format(
        sized_count, len(graph.edges)))

    # ------------------------------------------------------------------
    # STEP 7 - Locate Revit annotation resources
    # ------------------------------------------------------------------
    vft = next(
        (v for v in FilteredElementCollector(doc).OfClass(ViewFamilyType)
         if v.ViewFamily == ViewFamily.Drafting),
        None
    )
    if vft is None:
        forms.alert("No Drafting view type found in this project.",
                    title="One-Line - Error")
        return

    tt_id = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElementId()
    if tt_id is None or tt_id == ElementId.InvalidElementId:
        forms.alert("No TextNoteType found in this project.",
                    title="One-Line - Error")
        return

    # Locate project annotation families (RJA - P Symbols - *)
    squiggle_sym = _activate_sym(doc, _get_annotation_symbol(doc, "RJA- Squiggle"))
    if squiggle_sym is None:
        squiggle_sym = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - Squiggle"))
    meter_sym    = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Meter"))
    valve_sym    = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Gate Valve"))
    prv_sym      = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Pressure Regulating Valve"))
    if prv_sym is None:
        prv_sym  = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Pressure Reducing Valve"))
    equip_sym    = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Equipment"))
    cap_sym      = _activate_sym(doc, _get_annotation_symbol(doc, "RJA- Equipment Cap"))
    if cap_sym is None:
        cap_sym  = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - Equipment Cap"))

    found_syms = [("Squiggle",      squiggle_sym),
                  ("Meter",         meter_sym),
                  ("Gate Valve",    valve_sym),
                  ("PRV",           prv_sym),
                  ("Equipment",     equip_sym),
                  ("Equipment Cap", cap_sym)]
    for sym_label, sym in found_syms:
        if sym is None:
            output.print_md(":warning: {} symbol not found - using fallback.".format(
                sym_label))
        else:
            output.print_md(":white_check_mark: {} symbol loaded.".format(sym_label))

    # ------------------------------------------------------------------
    # STEP 8 - Create DraftingView
    # ------------------------------------------------------------------
    view_name = "Gas Piping One-Line - {}".format(
        datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S"))

    t_view = Transaction(doc, "RJA Tools - Create One-Line View")
    t_view.Start()
    try:
        view = ViewDrafting.Create(doc, vft.Id)
        view.Name = view_name
        view.Scale = VIEW_SCALE
        t_view.Commit()
    except Exception as ex:
        t_view.RollBack()
        forms.alert("Could not create DraftingView:\n\n{}".format(str(ex)),
                    title="One-Line - Error")
        output.print_md(":cross_mark: DraftingView failed: {}".format(str(ex)))
        return

    output.print_md(":white_check_mark: DraftingView created: **{}**".format(
        view_name))

    # ------------------------------------------------------------------
    # STEP 9 - Draw everything
    # ------------------------------------------------------------------
    output.print_md("**Drawing diagram...**")

    trunk_all_ids_draw = list(graph.longest_run["path_element_ids"])
    trunk_edge_ids_ord = [i for i in trunk_all_ids_draw if i in graph.edges]

    # Meter symbol and upstream stub drawn at the meter's diagram origin.
    # All trunk edges (including any initial riser) are drawn and labeled so
    # segment labels sum to the TOTAL DEVELOPED LENGTH in the notes block.
    mx, my = positions.get(meter_nid, (0.0, 0.0))

    t = Transaction(doc, "RJA Tools - Gas One-Line Diagram")
    t.Start()
    try:
        # a. Upstream stub + squiggle + "GAS FROM UTILITY" label
        #    Drawn at the distribution main level; stub goes left then DOWN
        #    to represent the underground utility service entry.
        #    Always New Construction (new meter connection).
        _draw_upstream_stub(doc, view, mx, my, squiggle_sym, tt_id,
                            line_style=wide_line_style)

        # b. Meter symbol (uses RJA - P Symbols - Meter when available)
        _draw_meter_symbol(doc, view, mx, my, tt_id, meter_sym)

        # c. Trunk pipe edges only.
        #    Branch edges are NOT drawn here -- they are replaced by the
        #    simplified schematic branches drawn in step (d).
        #    Upstream edges (riser from meter to distribution main) are drawn
        #    as plain lines; the upstream stub represents them schematically.
        #    Consecutive trunk edges with no branch tap or fixture between
        #    them, and equal cumulative load (i.e. no equipment takeoff in
        #    between -- just elbows/fittings), are merged into a single
        #    label, even when the run changes direction (horizontal-vertical-
        #    horizontal, etc).
        trunk_dev_lengths      = _trunk_edge_developed_lengths(graph, trunk_all_ids_draw)
        branch_point_positions = set(bi["tee_pos"] for bi in branch_info)

        drawn_edges = 0
        runs        = []
        current_run = []
        prev_geom   = None  # (x1, y1, to_node_id, cum_mbh)
        for eid in trunk_edge_ids_ord:
            edge = graph.edges[eid]
            if edge.to_node_id is None:
                continue
            pos_from = positions.get(edge.from_node_id)
            pos_to   = positions.get(edge.to_node_id)
            if pos_from is None or pos_to is None:
                continue
            x1, y1 = pos_to

            continues_run = False
            if prev_geom is not None:
                p_x1, p_y1, p_to_nid, p_mbh = prev_geom
                connects      = (p_to_nid == edge.from_node_id)
                same_mbh      = abs(edge.cumulative_load_mbh - p_mbh) < 0.01
                not_at_fix    = p_to_nid not in trunk_fixture_nids
                continues_run = connects and same_mbh and not_at_fix

            if continues_run:
                current_run.append(eid)
            else:
                if current_run:
                    runs.append(current_run)
                current_run = [eid]
            prev_geom = (x1, y1, edge.to_node_id, edge.cumulative_load_mbh)

        if current_run:
            runs.append(current_run)

        for run in runs:
            for eid in run:
                edge     = graph.edges[eid]
                pos_from = positions[edge.from_node_id]
                pos_to   = positions[edge.to_node_id]
                ls = wide_line_style if eid in new_construction_eids else None
                _line(doc, view, pos_from[0], pos_from[1], pos_to[0], pos_to[1],
                      line_style=ls)
                drawn_edges += 1

            first_edge = graph.edges[run[0]]
            last_edge  = graph.edges[run[-1]]
            # Label position uses the LAST segment so an L-shaped run
            # (riser + horizontal) labels the horizontal, not the corner.
            lbl_from_x, lbl_from_y = positions[last_edge.from_node_id]
            lbl_to_x,   lbl_to_y   = positions[last_edge.to_node_id]
            is_h = abs(lbl_to_y - lbl_from_y) <= abs(lbl_to_x - lbl_from_x)

            dev_len = sum(trunk_dev_lengths.get(e, graph.edges[e].length_feet)
                           for e in run)
            lft = int(round(dev_len))
            mbh = int(round(first_edge.cumulative_load_mbh))
            nom = ""
            for e in run:
                nom = pipe_sizes.get(e, "")
                if nom:
                    break
            if not nom and _fb_pairs:
                nom = _pick_fallback_size(mbh, _fb_pairs)

            if nom:
                line1 = '{}"G, {} FT'.format(nom, lft)
            else:
                line1 = "{} FT".format(lft)
            label = line1 + "\n" + "{} MBH".format(mbh)

            if is_h:
                lx = (lbl_from_x + lbl_to_x) / 2.0
                ly = max(lbl_from_y, lbl_to_y) + LABEL_ABOVE
                _note(doc, view, lx, ly, label, tt_id, width=True)
            else:
                lx = max(lbl_from_x, lbl_to_x) + LABEL_RIGHT
                ly = (lbl_from_y + lbl_to_y) / 2.0
                _note(doc, view, lx, ly, label, tt_id)

        # c2. "New main" spine segments (see _emit_new_main) -- the riser
        #     and the flat horizontal run of a branch that itself carries
        #     2+ fixtures (e.g. RTU-1..6). Drawn and labeled exactly like
        #     a trunk run above: real segment length, size, and
        #     cumulative MBH still flowing through that stretch.
        for seg in spine_segments:
            fx, fy = seg["from_pos"]
            tx2, ty2 = seg["to_pos"]
            seg_ls = wide_line_style if any(
                e in new_construction_eids for e in seg.get("eids", [])) else None
            _line(doc, view, fx, fy, tx2, ty2, line_style=seg_ls)

            if seg["cum_mbh"] is None:
                continue  # plain connector -- nothing new to report here

            lft = int(round(seg["length_ft"]))
            mbh = int(round(seg["cum_mbh"]))
            nom = seg.get("size", "")
            if nom:
                line1 = '{}"G, {} FT'.format(nom, lft)
            else:
                line1 = "{} FT".format(lft)
            label = line1 + "\n" + "{} MBH".format(mbh)
            lx = (fx + tx2) / 2.0
            ly = max(fy, ty2) + LABEL_ABOVE
            _note(doc, view, lx, ly, label, tt_id, width=True)

        # d. Schematic branches: one clean line per fixture from its trunk tee
        drawn_fixtures = 0
        drawn_valves   = 0
        for bi in branch_info:
            tee_x, tee_y = bi["tee_pos"]
            fix_x, fix_y = bi["fixture_pos"]
            node = graph.nodes.get(bi["fixture_nid"])
            if node is None:
                continue

            # Line style for this branch: wide if any of its pipe edges are
            # New Construction, default (None) otherwise.
            bi_ls = None
            if wide_line_style is not None:
                all_bi_eids = list(bi.get("branch_edge_ids", []))
                for sf in bi.get("sub_fixtures", []):
                    all_bi_eids.extend(sf.get("branch_edge_ids", []))
                if any(eid in new_construction_eids for eid in all_bi_eids):
                    bi_ls = wide_line_style

            if bi.get("sub_fixtures"):
                _draw_schematic_branch_with_stubs(
                    doc, view, bi, graph, tt_id,
                    valve_sym=valve_sym, prv_sym=prv_sym, equip_sym=equip_sym,
                    cap_sym=cap_sym, line_style=bi_ls)
                drawn_fixtures += 1 + len(bi["sub_fixtures"])
            else:
                _draw_schematic_branch(
                    doc, view,
                    tee_x, tee_y, fix_x, fix_y,
                    bi["direc"],
                    bi["total_ft"],
                    bi["size"],
                    bi.get("has_isolation", False),
                    bi.get("has_prv", False),
                    node,
                    tt_id,
                    valve_sym=valve_sym,
                    prv_sym=prv_sym,
                    equip_sym=equip_sym,
                    cap_sym=cap_sym,
                    line_style=bi_ls)
                drawn_fixtures += 1
            if bi.get("has_isolation") or bi.get("has_prv"):
                drawn_valves += 1

        # d2. Trunk-endpoint fixtures (e.g. MAU-1 at end of trunk).
        #     If the trunk runs INTO the fixture horizontally (e.g. RTU-2 at
        #     the end of the line), the 3-line equipment symbol is rotated
        #     90 degrees so the lines run vertically, perpendicular to the
        #     trunk -- per firm standard.
        for fix_nid in trunk_fixture_nids:
            pos  = positions.get(fix_nid)
            node = graph.nodes.get(fix_nid)
            if pos is None or node is None:
                continue
            cx, cy = pos

            incoming_edge = None
            for eid in trunk_set:
                e = graph.edges.get(eid)
                if e and e.to_node_id == fix_nid:
                    incoming_edge = e
                    break
            from_pos = positions.get(incoming_edge.from_node_id) if incoming_edge else None

            horiz_approach = True
            if from_pos is not None:
                horiz_approach = abs(cx - from_pos[0]) >= abs(cy - from_pos[1])

            name  = node.fixture_name or "UNNAMED"
            label = name + "\n" + "{} MBH".format(int(round(node.gas_load_mbh)))

            tf_has_prv, tf_has_isolation = _trunk_fixture_valves(graph, fix_nid, trunk_all_ids_draw)

            if horiz_approach:
                # Trunk runs horizontally into this fixture: rotate the
                # 3-line symbol 90 degrees (vertical lines, stacked along x
                # in the direction the trunk approached from).
                sign = 1.0 if (from_pos is None or cx >= from_pos[0]) else -1.0
                # Isolation VALVE_FIXTURE_GAP from fixture; PRV one VALVE_PITCH further
                iso_x = cx - sign * (FIXTURE_HW + VALVE_FIXTURE_GAP)
                if tf_has_isolation:
                    v_inst = _place_sym(doc, view, valve_sym, iso_x, cy)
                    if v_inst is None:
                        _make_group(doc, _draw_valve_bowtie(doc, view, iso_x, cy))
                if tf_has_prv:
                    prv_x = (iso_x - sign * VALVE_PITCH) if tf_has_isolation else iso_x
                    p_inst = _place_sym(doc, view, prv_sym, prv_x, cy)
                    if p_inst is None:
                        _make_group(doc, _draw_valve_bowtie(doc, view, prv_x, cy))
                _place_equipment_symbol(doc, view, equip_sym, cap_sym, cx, cy, sign,
                                         name, rotate_90=True)
                x_center = cx + sign * 1.0 * FIXTURE_SPACING
                lbl_y    = cy - FIXTURE_HW - FIXTURE_LABEL_GAP
                _note(doc, view, x_center, lbl_y, label, tt_id, center_align=True,
                      underline_len=len(name))
            else:
                # going_up must be relative to where the trunk pipe approaches
                # FROM, not an absolute cy>0 test -- otherwise a trunk that
                # itself sits above y=0 would flip every branch's connect
                # point to the wrong end of the 3-line symbol.
                going_up = cy > from_pos[1]
                sign     = 1.0 if going_up else -1.0
                # Isolation VALVE_FIXTURE_GAP from fixture; PRV one VALVE_PITCH further
                iso_y = cy - sign * VALVE_FIXTURE_GAP
                if tf_has_isolation:
                    v_inst = _place_sym(doc, view, valve_sym, cx, iso_y, rotate_90=True)
                    if v_inst is None:
                        _make_group(doc, _draw_valve_bowtie(doc, view, cx, iso_y))
                if tf_has_prv:
                    prv_y = (iso_y - sign * VALVE_PITCH) if tf_has_isolation else iso_y
                    p_inst = _place_sym(doc, view, prv_sym, cx, prv_y, rotate_90=True)
                    if p_inst is None:
                        _make_group(doc, _draw_valve_bowtie(doc, view, cx, prv_y))
                _place_equipment_symbol(doc, view, equip_sym, cap_sym, cx, cy, sign,
                                         name, rotate_180=going_up)
                far_y = cy + sign * 2 * FIXTURE_SPACING
                if sign > 0:
                    lbl_y = far_y + FIXTURE_LABEL_GAP + 2 * TEXT_HEIGHT_FT
                else:
                    lbl_y = far_y - FIXTURE_LABEL_GAP
                _note(doc, view, cx, lbl_y, label, tt_id, center_align=True,
                      underline_len=len(name))

            drawn_fixtures += 1

        # e. Notes block (single text box, positioned above diagram)
        _draw_notes_block(doc, view,
                          table_id           = table_id,
                          inlet_psi          = inlet_pressure_psi,
                          total_mbh          = total_mbh,
                          total_developed_ft = total_developed_ft,
                          tt_id              = tt_id,
                          notes_x            = notes_x,
                          notes_y            = notes_y,
                          table_label        = selected_opt.get("label", ""))

        t.Commit()

    except Exception as ex:
        t.RollBack()
        forms.alert(
            "Drawing transaction failed:\n\n{}".format(str(ex)),
            title="One-Line - Transaction Error"
        )
        output.print_md(":cross_mark: Transaction ERROR: {}".format(str(ex)))
        return

    # ------------------------------------------------------------------
    # STEP 10 - Open the view
    # ------------------------------------------------------------------
    try:
        uidoc.ActiveView = view
    except Exception:
        pass

    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| View | {} |".format(view_name))
    output.print_md("| Pipe segments drawn | {} |".format(drawn_edges))
    output.print_md("| Fixtures labeled | {} |".format(drawn_fixtures))
    output.print_md("| Valves drawn | {} |".format(drawn_valves))
    output.print_md("| Pipe sizes shown | {} |".format(sized_count))
    output.print_md("| IFGC table | {} |".format(table_id))
    output.print_md("| Total load | {:.1f} MBH |".format(total_mbh))
    output.print_md("| Longest run | {:.1f} ft |".format(longest_ft))
    output.print_md("")
    output.print_md(":white_check_mark: **One-line diagram generated.**")
    if sized_count == 0:
        output.print_md(
            ":warning: No pipe sizes found. Run Size Gas first to show "
            "nominal sizes on the diagram.")


main()
