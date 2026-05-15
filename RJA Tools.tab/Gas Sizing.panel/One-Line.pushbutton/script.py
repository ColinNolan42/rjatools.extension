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
VIEW_SCALE      = 100      # DraftingView.Scale
LEVEL_HEIGHT    = 10.0     # ft per branch depth (not actual elevation)
SYMBOL_RADIUS   = 0.5      # ft  meter circle radius
FIXTURE_HW      = 0.75     # ft  half-width of 3-line fixture symbol
FIXTURE_SPACING = 0.3      # ft  gap between 3 fixture lines
VALVE_HW        = 0.5      # ft  half-width of bowtie
VALVE_HH        = 0.3      # ft  half-height of bowtie triangle
LABEL_ABOVE     = 0.8      # ft  above a horizontal pipe
LABEL_RIGHT     = 0.4      # ft  right of a vertical pipe
UPSTREAM_H      = 6.0      # ft  horizontal stub left of meter
UPSTREAM_V      = 4.0      # ft  vertical drop of upstream stub
TEXT_GAP        = 0.78 * 1.8   # ft  spacing between note lines (3/32" x 1.8 at 1:100)
NOTES_X_BASE    = -(UPSTREAM_H + SYMBOL_RADIUS + 2.0)
NOTES_Y_BASE    = LEVEL_HEIGHT + 2.0

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
    """Assign (x, y) view positions to every graph node.

    Returns:
        positions dict  {node_id: (x, y)}  in view feet
        trunk_set       set of edge element IDs on the main trunk
    """
    trunk_ordered = list(graph.longest_run["path_element_ids"])
    trunk_set     = set(trunk_ordered)

    # Meter node = from_node_id of the first trunk edge
    meter_edge = graph.edges[trunk_ordered[0]]
    meter_nid  = meter_edge.from_node_id
    meter_z    = _node_z(graph, meter_nid)

    positions = {meter_nid: (0.0, 0.0)}
    x = 0.0

    # Walk the trunk, classifying each edge as horizontal or vertical
    for eid in trunk_ordered:
        edge    = graph.edges.get(eid)
        if edge is None:
            continue
        from_z  = _node_z(graph, edge.from_node_id, meter_z)
        to_z    = _node_z(graph, edge.to_node_id,   meter_z)
        z_delta = to_z - from_z
        L       = max(edge.length_feet, 0.001)
        is_vert = abs(z_delta) >= 0.5 * L

        fx, fy = positions.get(edge.from_node_id, (x, 0.0))

        if is_vert:
            direc = 1.0 if z_delta > 0 else -1.0
            positions[edge.to_node_id] = (fx, fy + direc * LEVEL_HEIGHT)
        else:
            x += edge.length_feet
            positions[edge.to_node_id] = (x, fy)

    # Place branches off each trunk node
    branch_counters = {}
    for eid in trunk_ordered:
        edge = graph.edges.get(eid)
        if edge is None:
            continue
        tee_nid = edge.from_node_id
        tx, ty  = positions.get(tee_nid, (0.0, 0.0))
        _place_branches_from(graph, trunk_set, positions, meter_z,
                             tee_nid, tx, ty, branch_counters)

    # Propagate positions to nodes linked via node_children (no pipe between them)
    for parent_nid, child_nids in graph.node_children.items():
        if parent_nid not in positions:
            continue
        for child_nid in child_nids:
            if child_nid not in positions:
                positions[child_nid] = positions[parent_nid]

    return positions, trunk_set, meter_nid, meter_z


def _place_branches_from(graph, trunk_set, positions, meter_z,
                          tee_nid, tx, ty, branch_counters):
    for branch_edge in _edges_from(graph, tee_nid):
        if branch_edge.element_id in trunk_set:
            continue
        if branch_edge.to_node_id in positions:
            continue
        to_z  = _node_z(graph, branch_edge.to_node_id, meter_z)
        direc = 1.0 if to_z > meter_z else -1.0
        depth = branch_counters.get(tee_nid, 0) + 1
        branch_counters[tee_nid] = depth
        by = ty + direc * LEVEL_HEIGHT * depth
        positions[branch_edge.to_node_id] = (tx, by)
        _layout_subtree(graph, trunk_set, positions, meter_z,
                        branch_edge.to_node_id, tx, by, direc)


def _layout_subtree(graph, trunk_set, positions, meter_z,
                    start_nid, x_start, y_start, direc):
    """After a vertical drop to start_nid, walk its children.
    Horizontal children advance x; vertical children drop further."""
    x = x_start
    sub_counters = {}
    for edge in _edges_from(graph, start_nid):
        if edge.element_id in trunk_set:
            continue
        if edge.to_node_id in positions:
            continue
        from_z  = _node_z(graph, start_nid, 0.0)
        to_z    = _node_z(graph, edge.to_node_id, 0.0)
        z_delta = to_z - from_z
        L       = max(edge.length_feet, 0.001)
        is_vert = abs(z_delta) >= 0.5 * L

        if is_vert:
            depth = sub_counters.get(start_nid, 0) + 1
            sub_counters[start_nid] = depth
            sy = y_start + direc * LEVEL_HEIGHT * depth
            positions[edge.to_node_id] = (x, sy)
            _layout_subtree(graph, trunk_set, positions, meter_z,
                            edge.to_node_id, x, sy, direc)
        else:
            x += edge.length_feet
            positions[edge.to_node_id] = (x, y_start)
            _layout_subtree(graph, trunk_set, positions, meter_z,
                            edge.to_node_id, x, y_start, direc)


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
    try:
        if (x0 == x1 and y0 == y1):
            return
        doc.Create.NewDetailCurve(
            view,
            Line.CreateBound(XYZ(x0, y0, 0), XYZ(x1, y1, 0)))
    except Exception:
        pass


def _note(doc, view, x, y, text, tt_id):
    try:
        TextNote.Create(doc, view.Id, XYZ(x, y, 0), text, tt_id)
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

    if nom:
        line1 = '{}\"G, {}\''.format(nom, lft)
    else:
        line1 = '{}\''.format(lft)
    label = line1 + "\n" + '{} MBH'.format(mbh)

    if is_h:
        lx = (x0 + x1) / 2.0
        ly = max(y0, y1) + LABEL_ABOVE
    else:
        lx = max(x0, x1) + LABEL_RIGHT
        ly = (y0 + y1) / 2.0
    _note(doc, view, lx, ly, label, tt_id)


def _draw_fixture_symbol(doc, view, cx, cy, going_up, node, tt_id):
    for i in range(3):
        yy = cy + i * FIXTURE_SPACING
        _line(doc, view, cx - FIXTURE_HW, yy, cx + FIXTURE_HW, yy)

    name  = node.fixture_name or "UNNAMED"
    label = name + "\n" + '{} MBH'.format(round(node.gas_load_mbh, 1))
    top_y = cy + 2 * FIXTURE_SPACING
    if going_up:
        ly = top_y + LABEL_ABOVE
    else:
        ly = cy - LABEL_ABOVE * 1.5
    _note(doc, view, cx, ly, label, tt_id)


def _draw_valve_bowtie(doc, view, cx, cy):
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
            _line(doc, view, p1[0], p1[1], p2[0], p2[1])


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
                going_up = _node_z(graph, nid, meter_z) > meter_z
                _draw_fixture_symbol(doc, view, cx, cy, going_up, node, tt_id)
                drawn_fixtures += 1
                continue

            fname = (node.family_name or "").lower()
            if any(kw in fname for kw in _VALVE_KW):
                _draw_valve_bowtie(doc, view, cx, cy)
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
