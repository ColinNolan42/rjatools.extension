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
    Line,
    Arc,
    XYZ,
    ViewDrafting,
    ViewFamilyType,
    ViewFamily,
    TextNote,
    TextNoteType,
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


# ---------------------------------------------------------------------------
# Layout constants  (all in Revit feet = view coordinates at 1:100)
# ---------------------------------------------------------------------------
VIEW_SCALE       = 100      # DraftingView.Scale
LEVEL_HEIGHT     = 18.0    # ft vertical clearance per branch level (compressed)
MIN_SEGMENT_FT   = 10.0    # ft minimum horizontal segment so text doesn't overlap
SYMBOL_RADIUS    = 0.5     # ft  meter circle radius
FIXTURE_HW       = 1.5     # ft  half-width of 3-line fixture symbol (3 ft total)
FIXTURE_SPACING  = 0.5     # ft  gap between 3 fixture lines
VALVE_HW         = 1.0     # ft  half-width of bowtie (2 ft total)
VALVE_HH         = 0.6     # ft  half-height of bowtie triangle
LABEL_ABOVE      = 1.2     # ft  above a horizontal pipe (must clear text height)
LABEL_RIGHT      = 0.6     # ft  right of a vertical pipe
UPSTREAM_H       = 6.0     # ft  horizontal stub left of meter
UPSTREAM_V       = 4.0     # ft  vertical drop of upstream stub
TEXT_HEIGHT_FT   = 0.78    # ft  3/32" x (100/12) at 1:100
TEXT_GAP         = TEXT_HEIGHT_FT * 2.0  # ft  between note lines
NOTES_X_BASE     = -(UPSTREAM_H + SYMBOL_RADIUS + 2.0)
NOTES_Y_BASE     = LEVEL_HEIGHT + 4.0

_VALVE_KW = ("valve", "prv", "regulator", "ball", "gate", "check", "shutoff")


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


def _trace_to_fixtures(graph, start_nid, trunk_set):
    """Trace all pipe edges from start_nid (non-trunk) to find fixtures.

    Collapses every intermediate fitting, elbow, transition, and CSST run
    into one result entry per fixture.  This implements the firm standard
    where top-takeoff routing (rise from trunk, horizontal, drop to equip)
    is shown as a single schematic vertical line on the one-line diagram.

    Returns list of dicts:
      fixture_nid:     node_id of the fixture
      total_length_ft: sum of all pipe segment lengths from start to fixture
      branch_edge_ids: list of pipe edge element IDs in the branch path
      has_valve:       True if any valve-type fitting was traversed
      cum_mbh:         cumulative_load_mbh from the starting branch edge
    """
    results = []
    # stack: (nid, acc_len_ft, acc_edge_ids, acc_valve)
    stack = [(start_nid, 0.0, [], False)]
    seen  = {start_nid}

    while stack:
        nid, acc_len, acc_eids, acc_valve = stack.pop()
        node = graph.nodes.get(nid)
        if not node:
            continue

        fname    = (node.family_name or "").lower()
        is_valve = any(kw in fname for kw in _VALVE_KW)
        nv       = acc_valve or is_valve

        if node.is_gas_fixture:
            results.append({
                "fixture_nid":     nid,
                "total_length_ft": acc_len,
                "branch_edge_ids": acc_eids,
                "has_valve":       nv,
                "cum_mbh":         node.cumulative_load_mbh,
            })
            continue

        for edge in graph.edges.values():
            if (edge.from_node_id == nid
                    and edge.element_id not in trunk_set
                    and edge.to_node_id
                    and edge.to_node_id not in seen):
                seen.add(edge.to_node_id)
                stack.append((
                    edge.to_node_id,
                    acc_len + edge.length_feet,
                    acc_eids + [edge.element_id],
                    nv,
                ))

        for child in graph.node_children.get(nid, []):
            if child not in seen:
                seen.add(child)
                stack.append((child, acc_len, acc_eids, nv))

    return results


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

            # Trace entire branch to find fixture(s) and cumulative length
            fixtures = _trace_to_fixtures(graph, branch_edge.to_node_id, trunk_set)
            if not fixtures:
                continue
            # Skip if the primary fixture is already positioned by a sibling
            # tee_candidate (avoids duplicate branches from the same physical tee)
            if positions.get(fixtures[0]["fixture_nid"]) is not None:
                continue

            # Direction from downstream fixture z vs this tee z
            first_fix_z = _node_z(graph, fixtures[0]["fixture_nid"], meter_z)
            direc = 1.0 if first_fix_z > tee_z else -1.0

            depth = branch_counters.get(tee_nid, 0) + 1
            branch_counters[tee_nid] = depth
            fix_y = ty + direc * LEVEL_HEIGHT * depth

            for i, fix_info in enumerate(fixtures):
                # If multiple fixtures from one tee, spread them horizontally
                fix_x   = tx + i * MIN_SEGMENT_FT
                fix_nid = fix_info["fixture_nid"]
                positions[fix_nid] = (fix_x, fix_y)

                branch_info.append({
                    "tee_nid":         tee_nid,
                    "tee_pos":         (tx, ty),
                    "fixture_nid":     fix_nid,
                    "fixture_pos":     (fix_x, fix_y),
                    "total_ft":        fix_info["total_length_ft"],
                    "branch_edge_ids": fix_info["branch_edge_ids"],
                    "has_valve":       fix_info["has_valve"],
                    "direc":           direc,
                    "size":            "",  # filled after pipe_sizes known
                    "cum_mbh":         fix_info["cum_mbh"],
                })
                layout_log.append({
                    "tee_nid":    tee_nid,
                    "tee_pos":    (tx, ty),
                    "tee_z":      tee_z,
                    "edge_id":    branch_edge.element_id,
                    "to_nid":     fix_nid,
                    "to_z":       first_fix_z,
                    "fixture_z":  first_fix_z,
                    "direc":      "UP" if direc > 0 else "DOWN",
                    "result_pos": (fix_x, fix_y),
                    "branch_y":   None,
                })

    return (positions, trunk_set, meter_nid, meter_z,
            layout_log, branch_info, trunk_fixture_nids)


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


def _place_sym(doc, view, sym, x, y, rotate_90=False):
    """Place a generic annotation instance at (x, y) in the view.

    If rotate_90 is True, rotate the instance 90 degrees around the Z-axis
    through the placement point.  Returns the placed element or None.
    """
    if sym is None:
        return None
    try:
        inst = doc.Create.NewFamilyInstance(XYZ(x, y, 0), sym, view)
        if rotate_90 and inst is not None:
            axis = Line.CreateBound(XYZ(x, y, 0), XYZ(x, y, 1))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.pi / 2.0)
        return inst
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Schematic branch drawing
# ---------------------------------------------------------------------------

def _draw_schematic_branch(doc, view, tee_x, tee_y, fix_x, fix_y,
                            direc, total_ft, size, has_valve,
                            fixture_node, tt_id,
                            valve_sym=None, equip_sym=None):
    """Draw one simplified schematic branch from trunk tee to fixture.

    Uses project annotation families when available:
      valve_sym  -- RJA - P Symbols - Gate Valve  (rotated 90 degrees)
      equip_sym  -- RJA - P Symbols - Equipment

    Symbol label is placed as a separate TextNote (not grouped with symbol).
    """
    # Vertical segment from tee to fixture level
    _line(doc, view, tee_x, tee_y, tee_x, fix_y)

    # Horizontal segment if fixture is offset from tee
    if abs(fix_x - tee_x) > 0.01:
        _line(doc, view, tee_x, fix_y, fix_x, fix_y)

    # Valve symbol 40% down the vertical (rotated 90 to sit on branch line)
    if has_valve:
        val_y = tee_y + (fix_y - tee_y) * 0.4
        v_inst = _place_sym(doc, view, valve_sym, tee_x, val_y, rotate_90=True)
        if v_inst is None:
            # Fallback: drawn bowtie lines
            val_elems = _draw_valve_bowtie(doc, view, tee_x, val_y)
            _make_group(doc, val_elems)

    # Label on right side of the vertical segment
    mid_y = (tee_y + fix_y) / 2.0
    if size:
        lbl_line1 = '{}"G, {}\''.format(size, int(round(total_ft)))
    else:
        lbl_line1 = "{}\'".format(int(round(total_ft)))
    mbh_val = int(round(fixture_node.gas_load_mbh)) if fixture_node else 0
    lbl = lbl_line1 + "\n" + "{} MBH".format(mbh_val)
    lbl_x = tee_x + (VALVE_HW + LABEL_RIGHT if has_valve else LABEL_RIGHT)
    _note(doc, view, lbl_x, mid_y, lbl, tt_id)

    # Equipment symbol at the fixture endpoint
    going_up = fix_y > tee_y
    # Direction sign: +1 for upward branch (lines extend up), -1 for downward
    # (lines extend down away from trunk so the outer line connects to the branch)
    sign = 1.0 if going_up else -1.0

    e_inst = _place_sym(doc, view, equip_sym, fix_x, fix_y)
    if e_inst is None:
        # Fallback: drawn 3-line symbol.
        # Line 0 (outer, connects to branch) is at fix_y.
        # Lines 1 and 2 extend AWAY from trunk.
        sym_elems = []
        for i in range(3):
            yy = fix_y + sign * i * FIXTURE_SPACING
            sym_elems.append(_line(doc, view, fix_x - FIXTURE_HW, yy,
                                   fix_x + FIXTURE_HW, yy))
        _make_group(doc, sym_elems)

    # Fixture name + MBH as a SEPARATE TextNote (not grouped with symbol).
    # Label is placed beyond the outermost (far) line of the symbol.
    if fixture_node:
        name     = fixture_node.fixture_name or "UNNAMED"
        label    = name + "\n" + "{} MBH".format(int(round(fixture_node.gas_load_mbh)))
        far_y    = fix_y + sign * 2 * FIXTURE_SPACING  # outermost line position
        lbl_y    = far_y + sign * LABEL_ABOVE           # beyond that
        _note(doc, view, fix_x, lbl_y, label, tt_id)


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------

def _line(doc, view, x0, y0, x1, y1):
    """Draw a detail line and return the element (or None on failure)."""
    try:
        if abs(x1 - x0) < 0.001 and abs(y1 - y0) < 0.001:
            return None
        return doc.Create.NewDetailCurve(
            view,
            Line.CreateBound(XYZ(x0, y0, 0), XYZ(x1, y1, 0)))
    except Exception:
        return None


def _note(doc, view, x, y, text, tt_id):
    """Create a TextNote and return the element (or None on failure)."""
    try:
        return TextNote.Create(doc, view.Id, XYZ(x, y, 0), text, tt_id)
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


def _draw_upstream_stub(doc, view, cx, cy, squiggle_sym, tt_id):
    """Draw horizontal stub + vertical drop + squiggle (rotated 90 deg)."""
    sx = cx - SYMBOL_RADIUS
    # Horizontal stub going left from meter
    _line(doc, view, sx, cy, sx - UPSTREAM_H, cy)
    # Vertical drop to utility
    tip_x = sx - UPSTREAM_H
    _line(doc, view, tip_x, cy, tip_x, cy - UPSTREAM_V)
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
        line1 = '{}"G, {}\''.format(nom, lft)
    else:
        line1 = "{}\'".format(lft)
    label = line1 + "\n" + "{} MBH".format(mbh)

    if is_h:
        lx = (x0 + x1) / 2.0
        ly = max(y0, y1) + LABEL_ABOVE
    else:
        lx = max(x0, x1) + LABEL_RIGHT
        ly = (y0 + y1) / 2.0
    _note(doc, view, lx, ly, label, tt_id)


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
    elems.append(_note(doc, view, cx, ly, label, tt_id))
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


def _draw_notes_block(doc, view, table_id, inlet_psi,
                      total_mbh, longest_ft, tt_id, notes_x, notes_y):
    """Draw the 5-line notes block as a SINGLE multi-line TextNote."""
    text = "\n".join([
        "CONTRACTOR SHALL SUBMIT APPLICATIONS TO UTILITY"
        " AND COORDINATE NEW METER SERVICE",
        "GAS PIPING SIZED FOR {} PSI".format(inlet_psi),
        "MAX PRESSURE LOSS PER IFGC TABLE {}".format(table_id),
        "TOTAL CONNECTED LOAD: {} MBH".format(int(round(total_mbh))),
        "TOTAL PIPE LENGTH: {}'".format(int(round(longest_ft))),
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
        selected_element.Id.IntegerValue))

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
    # STEP 3 - Select IFGC table (populates notes block)
    # ------------------------------------------------------------------
    option_labels = gas_tables.get_table_option_labels()
    selected_label = forms.SelectFromList.show(
        option_labels,
        title="One-Line - Select IFGC Sizing Table",
        multiselect=False
    )
    if not selected_label:
        output.print_md("Cancelled at table selection. No changes made.")
        return

    selected_opt       = gas_tables.get_table_option_by_label(selected_label)
    table_id           = selected_opt["table_id"]
    inlet_pressure_psi = selected_opt["inlet_pressure_psi"]
    output.print_md("**Table:** {}".format(table_id))

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

    total_mbh    = sum(n.gas_load_mbh for n in fixture_nodes)
    longest_ft   = graph.longest_run["total_length_feet"]  # used for IFGC sizing
    # For the notes display use pipe_length_feet (actual pipe only, no elbow
    # equivalents) so the labeled segment lengths on the diagram add up to match.
    display_ft   = graph.longest_run.get("pipe_length_feet", longest_ft)

    # ------------------------------------------------------------------
    # STEP 5 - Compute layout
    # ------------------------------------------------------------------
    output.print_md("**Computing layout...**")
    (positions, trunk_set, meter_nid, meter_z,
     layout_log, branch_info, trunk_fixture_nids) = _compute_layout(graph)
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
    if tt_id is None or tt_id.IntegerValue < 0:
        forms.alert("No TextNoteType found in this project.",
                    title="One-Line - Error")
        return

    # Locate project annotation families (RJA - P Symbols - *)
    squiggle_sym = _activate_sym(doc, _get_annotation_symbol(doc, "RJA- Squiggle"))
    if squiggle_sym is None:
        squiggle_sym = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - Squiggle"))
    meter_sym    = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Meter"))
    valve_sym    = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Gate Valve"))
    prv_sym      = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Pressure Reducing Valve"))
    equip_sym    = _activate_sym(doc, _get_annotation_symbol(doc, "RJA - P Symbols - Equipment"))

    found_syms = [("Squiggle",      squiggle_sym),
                  ("Meter",         meter_sym),
                  ("Gate Valve",    valve_sym),
                  ("PRV",           prv_sym),
                  ("Equipment",     equip_sym)]
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
        datetime.datetime.now().strftime("%Y%m%d-%H%M"))

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

    # Find the distribution main start position.
    # If the trunk begins with a riser (VERT section between meter and the
    # horizontal main), place the meter symbol at the TOP of that riser so the
    # diagram shows a clean horizontal trunk.  The upstream stub then extends
    # DOWN from the meter circle to represent the utility service entry.
    trunk_all_ids_draw = list(graph.longest_run["path_element_ids"])
    trunk_edge_ids_ord = [i for i in trunk_all_ids_draw if i in graph.edges]

    # Identify edges that are "upstream" (initial riser + any leading stub).
    # These are trunk edges before the FIRST horizontal segment that follows a
    # VERT section.  They will still be drawn as lines but not labeled.
    upstream_draw_edges = set()
    main_origin         = positions.get(meter_nid, (0.0, 0.0))
    found_vert          = False
    for eid in trunk_edge_ids_ord:
        edge = graph.edges[eid]
        fp = positions.get(edge.from_node_id, (0.0, 0.0))
        tp = positions.get(edge.to_node_id,   (0.0, 0.0))
        seg_is_vert = abs(tp[1] - fp[1]) > abs(tp[0] - fp[0])
        if seg_is_vert:
            found_vert = True
            upstream_draw_edges.add(eid)
        elif found_vert:
            # First horizontal after the riser = distribution main start.
            main_origin = fp
            break
        else:
            upstream_draw_edges.add(eid)  # short stubs before the riser

    mx, my = main_origin

    t = Transaction(doc, "RJA Tools - Gas One-Line Diagram")
    t.Start()
    try:
        # a. Upstream stub + squiggle + "GAS FROM UTILITY" label
        #    Drawn at the distribution main level; stub goes left then DOWN
        #    to represent the underground utility service entry.
        _draw_upstream_stub(doc, view, mx, my, squiggle_sym, tt_id)

        # b. Meter symbol (uses RJA - P Symbols - Meter when available)
        _draw_meter_symbol(doc, view, mx, my, tt_id, meter_sym)

        # c. Trunk pipe edges only.
        #    Branch edges are NOT drawn here -- they are replaced by the
        #    simplified schematic branches drawn in step (d).
        #    Upstream edges (riser from meter to distribution main) are drawn
        #    as plain lines; the upstream stub represents them schematically.
        drawn_edges = 0
        for eid, edge in graph.edges.items():
            if eid not in trunk_set:
                continue  # branch edges: handled by _draw_schematic_branch
            if edge.to_node_id is None:
                continue
            pos_from = positions.get(edge.from_node_id)
            pos_to   = positions.get(edge.to_node_id)
            if pos_from is None or pos_to is None:
                continue
            x0, y0 = pos_from
            x1, y1 = pos_to
            if eid in upstream_draw_edges:
                continue  # skip: service connection shown by upstream stub
            _draw_pipe_segment(doc, view, x0, y0, x1, y1, edge, pipe_sizes, tt_id,
                                force_label=True)
            drawn_edges += 1

        # d. Schematic branches: one clean line per fixture from its trunk tee
        drawn_fixtures = 0
        drawn_valves   = 0
        for bi in branch_info:
            tee_x, tee_y = bi["tee_pos"]
            fix_x, fix_y = bi["fixture_pos"]
            node = graph.nodes.get(bi["fixture_nid"])
            if node is None:
                continue
            _draw_schematic_branch(
                doc, view,
                tee_x, tee_y, fix_x, fix_y,
                bi["direc"],
                bi["total_ft"],
                bi["size"],
                bi["has_valve"],
                node,
                tt_id,
                valve_sym=valve_sym,
                equip_sym=equip_sym)
            drawn_fixtures += 1
            if bi["has_valve"]:
                drawn_valves += 1

        # d2. Trunk-endpoint fixtures (e.g. MAU-1 at end of trunk)
        for fix_nid in trunk_fixture_nids:
            pos  = positions.get(fix_nid)
            node = graph.nodes.get(fix_nid)
            if pos is None or node is None:
                continue
            cx, cy = pos
            going_up = cy > 0.0
            sign     = 1.0 if going_up else -1.0
            e_inst = _place_sym(doc, view, equip_sym, cx, cy)
            if e_inst is None:
                sym_elems = []
                for i in range(3):
                    yy = cy + sign * i * FIXTURE_SPACING
                    sym_elems.append(_line(doc, view, cx - FIXTURE_HW, yy,
                                           cx + FIXTURE_HW, yy))
                _make_group(doc, sym_elems)
            name   = node.fixture_name or "UNNAMED"
            label  = name + "\n" + "{} MBH".format(int(round(node.gas_load_mbh)))
            far_y  = cy + sign * 2 * FIXTURE_SPACING
            lbl_y  = far_y + sign * LABEL_ABOVE
            _note(doc, view, cx, lbl_y, label, tt_id)
            drawn_fixtures += 1

        # e. Notes block (single text box, positioned above diagram)
        _draw_notes_block(doc, view,
                          table_id   = table_id,
                          inlet_psi  = inlet_pressure_psi,
                          total_mbh  = total_mbh,
                          longest_ft = display_ft,   # pipe length only (no elbows)
                          tt_id      = tt_id,
                          notes_x    = notes_x,
                          notes_y    = notes_y)

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
