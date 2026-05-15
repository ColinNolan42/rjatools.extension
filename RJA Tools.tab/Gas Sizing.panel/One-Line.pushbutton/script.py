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
LEVEL_HEIGHT     = 15.0    # ft vertical clearance per branch level (compressed)
MIN_SEGMENT_FT   = 6.0     # ft minimum horizontal segment so text doesn't overlap
SYMBOL_RADIUS    = 0.5     # ft  meter circle radius
FIXTURE_HW       = 0.75    # ft  half-width of 3-line fixture symbol
FIXTURE_SPACING  = 0.4     # ft  gap between 3 fixture lines
VALVE_HW         = 0.6     # ft  half-width of bowtie
VALVE_HH         = 0.35    # ft  half-height of bowtie triangle
LABEL_ABOVE      = 1.0     # ft  above a horizontal pipe (must clear text height)
LABEL_RIGHT      = 0.5     # ft  right of a vertical pipe
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
    # Phase 2: BFS from ALL trunk nodes to reach every branch node.
    #
    # Queue entries: (node_id, branch_y, branch_direc, branch_x)
    #   branch_y     = y of this branch level (None = on trunk)
    #   branch_direc = +1 up / -1 down for this branch
    #   branch_x     = current x within the branch
    #
    # Branch direction rules (top-takeoff aware):
    #   - First pipe off a trunk tee: direction = sign of (branch_end_z - tee_z).
    #     Most pipes take off upward from the trunk (top takeoff) and come back
    #     down to the equipment; the endpoint above the trunk -> UP.
    #   - ALL subsequent pipes within a branch always continue HORIZONTALLY,
    #     regardless of Revit z-delta.  This avoids showing the final drop-to-
    #     equipment as a second vertical segment in the schematic.
    #   - If a branch node has MORE THAN ONE outgoing non-trunk edge, the first
    #     is the horizontal continuation; each additional edge sub-branches by
    #     dropping another LEVEL_HEIGHT in branch_direc from the current y.
    # ------------------------------------------------------------------
    branch_counters = {}  # node_id -> branches dispatched from it
    visited = set(positions.keys())

    queue = [(nid, None, 1.0, positions[nid][0])
             for nid in sorted(trunk_nodes) if nid in positions]
    qi = 0
    while qi < len(queue):
        nid, branch_y, branch_direc, branch_x = queue[qi]
        qi += 1
        tx, ty = positions[nid]

        # Walk pipe edges from this node
        for edge in _edges_from(graph, nid):
            to_nid = edge.to_node_id
            if to_nid is None or to_nid in visited:
                continue
            if edge.element_id in trunk_set:
                continue
            visited.add(to_nid)

            if branch_y is None:
                # ---- First drop from a trunk tee ----
                # Direction based on branch endpoint z vs THIS tee's z.
                tee_z = _node_z(graph, nid, meter_z)
                to_z  = _node_z(graph, to_nid, tee_z)
                direc = 1.0 if to_z > tee_z else -1.0
                depth = branch_counters.get(nid, 0) + 1
                branch_counters[nid] = depth
                new_y   = ty + direc * LEVEL_HEIGHT * depth
                new_pos = (tx, new_y)
                queue.append((to_nid, new_y, direc, tx))
            else:
                # ---- Within a branch ----
                # First outgoing edge from this node continues horizontally.
                # Additional edges (branch splits) sub-drop in branch_direc.
                dispatched = branch_counters.get(nid, 0)
                if dispatched == 0:
                    new_x   = branch_x + max(edge.length_feet, MIN_SEGMENT_FT)
                    new_pos = (new_x, ty)
                    branch_counters[nid] = 1
                    queue.append((to_nid, branch_y, branch_direc, new_x))
                else:
                    depth   = dispatched + 1
                    branch_counters[nid] = depth
                    new_y   = ty + branch_direc * LEVEL_HEIGHT
                    new_pos = (tx, new_y)
                    queue.append((to_nid, new_y, branch_direc, tx))

            positions[to_nid] = new_pos

        # Walk node_children connections (no pipe, same position as parent)
        for child_nid in graph.node_children.get(nid, []):
            if child_nid in visited:
                continue
            visited.add(child_nid)
            positions[child_nid] = positions[nid]
            queue.append((child_nid, branch_y, branch_direc, branch_x))

    return positions, trunk_set, meter_nid, meter_z


# ---------------------------------------------------------------------------
# Pipe size reading
# ---------------------------------------------------------------------------

def _read_pipe_sizes(graph):
    """Read nominal pipe sizes from Revit model.

    Returns {edge_element_id: nominal_size_str} for all readable edges.
    """
    inv = {}
    for nom, inches in sizing_engine.NOMINAL_TO_INCHES.items():
        if inches not in inv:
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


def _draw_meter_symbol(doc, view, cx, cy, tt_id):
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


def _draw_upstream_stub(doc, view, cx, cy, squiggle_sym):
    sx = cx - SYMBOL_RADIUS
    # Horizontal stub going left
    _line(doc, view, sx, cy, sx - UPSTREAM_H, cy)
    # Vertical drop
    tip_x = sx - UPSTREAM_H
    _line(doc, view, tip_x, cy, tip_x, cy - UPSTREAM_V)
    # Squiggle at tip
    if squiggle_sym is not None:
        try:
            doc.Create.NewFamilyInstance(
                XYZ(tip_x, cy - UPSTREAM_V, 0), squiggle_sym, view)
        except Exception:
            pass


def _draw_pipe_segment(doc, view, x0, y0, x1, y1, edge, pipe_sizes, tt_id):
    _line(doc, view, x0, y0, x1, y1)

    nom  = pipe_sizes.get(edge.element_id, "")
    lft  = int(round(edge.length_feet))
    mbh  = round(edge.cumulative_load_mbh, 1)
    is_h = abs(y1 - y0) <= abs(x1 - x0)

    # Label format: "2-1/2\", 25'" on line 1, "350.0 MBH" on line 2
    if nom:
        line1 = '{}\", {}\''.format(nom, lft)
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
    """Draw 3-line equipment symbol and label. Returns elements for grouping."""
    elems = []
    for i in range(3):
        yy = cy + i * FIXTURE_SPACING
        elems.append(_line(doc, view, cx - FIXTURE_HW, yy, cx + FIXTURE_HW, yy))

    name  = node.fixture_name or "UNNAMED"
    label = name + "\n" + "{} MBH".format(round(node.gas_load_mbh, 1))
    top_y = cy + 2 * FIXTURE_SPACING
    if going_up:
        ly = top_y + LABEL_ABOVE
    else:
        ly = cy - LABEL_ABOVE * 1.5
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
                      total_mbh, longest_ft, tt_id):
    lines = [
        "CONTRACTOR SHALL SUBMIT APPLICATIONS TO UTILITY"
        " AND COORDINATE NEW METER SERVICE",
        "GAS PIPING SIZED FOR {} PSI".format(inlet_psi),
        "MAX PRESSURE LOSS PER IFGC TABLE {}".format(table_id),
        "TOTAL CONNECTED LOAD: {:.1f} MBH".format(total_mbh),
        "TOTAL DEVELOPED LENGTH: {}'".format(int(round(longest_ft))),
    ]
    for i, line in enumerate(lines):
        _note(doc, view,
              NOTES_X_BASE,
              NOTES_Y_BASE - i * TEXT_GAP,
              line, tt_id)


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

    total_mbh  = sum(n.gas_load_mbh for n in fixture_nodes)
    longest_ft = graph.longest_run["total_length_feet"]

    # ------------------------------------------------------------------
    # STEP 5 - Compute layout
    # ------------------------------------------------------------------
    output.print_md("**Computing layout...**")
    positions, trunk_set, meter_nid, meter_z = _compute_layout(graph)
    output.print_md(":white_check_mark: {} nodes positioned.".format(
        len(positions)))

    # ------------------------------------------------------------------
    # STEP 6 - Read pipe sizes from model
    # ------------------------------------------------------------------
    pipe_sizes = _read_pipe_sizes(graph)
    sized_count = len(pipe_sizes)
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

    squiggle_sym = next(
        (s for s in FilteredElementCollector(doc).OfClass(FamilySymbol)
         if "Squiggle" in (s.Family.Name if s.Family else "")),
        None
    )
    if squiggle_sym is None:
        output.print_md(
            ":warning: RJA-Squiggle annotation family not found in project. "
            "Upstream stub will be drawn without squiggle symbol.")
    elif not squiggle_sym.IsActive:
        t_act = Transaction(doc, "Activate Squiggle Symbol")
        t_act.Start()
        try:
            squiggle_sym.Activate()
            t_act.Commit()
        except Exception:
            t_act.RollBack()
            squiggle_sym = None

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

    mx, my = positions.get(meter_nid, (0.0, 0.0))

    t = Transaction(doc, "RJA Tools - Gas One-Line Diagram")
    t.Start()
    try:
        # a. Upstream stub + squiggle
        _draw_upstream_stub(doc, view, mx, my, squiggle_sym)

        # b. Meter circle + "M"
        _draw_meter_symbol(doc, view, mx, my, tt_id)

        # c. All pipe edges
        drawn_edges = 0
        for eid, edge in graph.edges.items():
            if edge.to_node_id is None:
                continue
            pos_from = positions.get(edge.from_node_id)
            pos_to   = positions.get(edge.to_node_id)
            if pos_from is None or pos_to is None:
                continue
            x0, y0 = pos_from
            x1, y1 = pos_to
            _draw_pipe_segment(doc, view, x0, y0, x1, y1, edge, pipe_sizes, tt_id)
            drawn_edges += 1

        # d. Fixture symbols and labels; valve/PRV symbols
        drawn_fixtures = 0
        drawn_valves   = 0
        for nid, node in graph.nodes.items():
            pos = positions.get(nid)
            if pos is None:
                continue
            cx, cy = pos

            if node.is_gas_fixture:
                # going_up = fixture z above the trunk (y=0 level)
                going_up = cy > 0.0
                fix_elems = _draw_fixture_symbol(
                    doc, view, cx, cy, going_up, node, tt_id)
                _make_group(doc, fix_elems)
                drawn_fixtures += 1
                continue

            fname = (node.family_name or "").lower()
            if any(kw in fname for kw in _VALVE_KW):
                val_elems = _draw_valve_bowtie(doc, view, cx, cy)
                _make_group(doc, val_elems)
                drawn_valves += 1

        # e. Notes block
        _draw_notes_block(doc, view,
                          table_id      = table_id,
                          inlet_psi     = inlet_pressure_psi,
                          total_mbh     = total_mbh,
                          longest_ft    = longest_ft,
                          tt_id         = tt_id)

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
