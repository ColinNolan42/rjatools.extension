# -*- coding: ascii -*-
# sizing_engine.py
# IFGC Longest Run Method gas pipe sizing engine.
# Pure Python - no Revit API calls. All table lookups via gas_tables.py.
#
# IronPython 2.7

import gas_tables
import shared_params


# ---------------------------------------------------------------------------
# Nominal size -> decimal inches mapping
# Used when writing sizes back to Revit (Size Gas script converts to feet).
# Keys must match pipe_sizes_nominal strings in ifgc_gas_sizing_tables.json.
# ---------------------------------------------------------------------------

NOMINAL_TO_INCHES = {
    "1/2":   0.5,
    "3/4":   0.75,
    "1":     1.0,
    "1-1/4": 1.25,
    "1-1/2": 1.5,
    "2":     2.0,
    "2-1/2": 2.5,
    "3":     3.0,
    "4":     4.0,
    "5":     5.0,
    "6":     6.0,
    "8":     8.0,
    "10":    10.0,
    "12":    12.0,
}


# ---------------------------------------------------------------------------
# Main sizing function
# ---------------------------------------------------------------------------

def size_system(graph, pipe_material, inlet_pressure_psi):
    """Size every pipe segment using the IFGC Longest Run Method.

    Per IFGC A103.1:
      - One longest run length is used for ALL segments.
      - Each segment is sized for its cumulative downstream demand (CFH).
      - Smallest nominal size whose table capacity >= demand is selected.
      - 1 MBH = 1 CFH for natural gas at 1000 BTU/cf.

    Args:
        graph:               NetworkGraph from pipe_graph.build_network()
        pipe_material:       str  e.g. "Schedule 40 Steel"
        inlet_pressure_psi:  float  supply pressure at the meter

    Returns:
        dict with keys:
            sizes            {pipe_element_id (int): nominal_size (str)}
            table_id         str  e.g. "402.4(2)"
            longest_run_ft   float
            pipe_material    str
            inlet_pressure_psi  float
            segment_detail   list of dicts - one per sized segment

    Raises:
        ValueError: If longest run is missing, table not available, or any
                    pipe demand exceeds the maximum table capacity.
    """
    if graph.longest_run is None:
        raise ValueError(
            "Longest run not found. Run Diagnose and verify the system "
            "has at least one reachable fixture before sizing.")

    longest_run_ft = graph.longest_run["total_length_feet"]
    if longest_run_ft <= 0:
        raise ValueError(
            "Longest run is 0 ft. Verify meter connection and traversal.")

    table_id = gas_tables.select_table(inlet_pressure_psi, pipe_material)
    pipe_sizes = gas_tables.list_pipe_sizes(table_id)
    table_length_used, _ = gas_tables.get_length_row(table_id, longest_run_ft)

    sizes = {}
    segment_detail = []
    sizing_errors = []

    for edge in graph.edges.values():
        # Skip open-ended pipes - no downstream node to size for
        if edge.to_node_id is None:
            continue

        demand_mbh = edge.cumulative_load_mbh  # 1 MBH = 1 CFH

        # Zero demand: assign minimum available pipe size
        if demand_mbh <= 0:
            selected = pipe_sizes[0]
            capacity_at_size = gas_tables.get_capacity(
                table_id, longest_run_ft, selected)
            segment_detail.append({
                "pipe_id":       edge.element_id,
                "demand_mbh":    0.0,
                "selected_size": selected,
                "capacity_mbh":  capacity_at_size,
                "note":          "zero demand - minimum size assigned"
            })
            sizes[edge.element_id] = selected
            continue

        # Find smallest size whose capacity >= demand
        selected = None
        selected_capacity = None
        for size in pipe_sizes:
            try:
                capacity = gas_tables.get_capacity(
                    table_id, longest_run_ft, size)
                if capacity >= demand_mbh:
                    selected = size
                    selected_capacity = capacity
                    break
            except ValueError:
                # Size not available at this length - try the next size up
                continue

        if selected is None:
            sizing_errors.append(
                "Pipe {}: demand {:.1f} MBH exceeds max table capacity "
                "at {:.0f} ft in Table {}.".format(
                    edge.element_id, demand_mbh, table_length_used, table_id))
            continue

        sizes[edge.element_id] = selected
        segment_detail.append({
            "pipe_id":       edge.element_id,
            "demand_mbh":    round(demand_mbh, 1),
            "selected_size": selected,
            "capacity_mbh":  selected_capacity,
            "note":          ""
        })

    if sizing_errors:
        raise ValueError(
            "Sizing failed for {} pipe(s):\n{}".format(
                len(sizing_errors), "\n".join(sizing_errors)))

    return {
        "sizes":               sizes,
        "table_id":            table_id,
        "table_length_used_ft": table_length_used,
        "longest_run_ft":      longest_run_ft,
        "pipe_material":       pipe_material,
        "inlet_pressure_psi":  inlet_pressure_psi,
        "segment_detail":      segment_detail,
    }


# ---------------------------------------------------------------------------
# Downstream fixture helper
# ---------------------------------------------------------------------------

def _downstream_fixtures(start_node_id, graph):
    """Return list of fixture names reachable downstream from start_node_id."""
    names = []
    visited = set()
    stack = [start_node_id]

    while stack:
        nid = stack.pop()
        if nid is None or nid in visited:
            continue
        visited.add(nid)

        node = graph.nodes.get(nid)
        if node is None:
            continue

        if node.is_gas_fixture:
            names.append(node.fixture_name or "UNNAMED")
            continue

        for edge in graph.edges.values():
            if edge.from_node_id == nid and edge.to_node_id is not None:
                stack.append(edge.to_node_id)

        for child_id in graph.node_children.get(nid, []):
            stack.append(child_id)

    return names


# ---------------------------------------------------------------------------
# Diagnostic formatter
# ---------------------------------------------------------------------------

def format_sizing_output(sizing_result, graph):
    """Return a formatted string summarizing sizing results for the output window.

    Args:
        sizing_result: dict returned by size_system()
        graph:         NetworkGraph used for sizing

    Returns:
        str
    """
    lines = []

    # Total load and fixture count
    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
    total_mbh = sum(n.gas_load_mbh for n in fixture_nodes)

    lines.append("=== SIZING RESULTS ===")
    lines.append("Table:         {}".format(sizing_result["table_id"]))
    lines.append("Table row:     {} ft  (longest run {:.1f} ft rounded up)".format(
        sizing_result["table_length_used_ft"],
        sizing_result["longest_run_ft"]))
    lines.append("Pipe material: {}".format(sizing_result["pipe_material"]))
    lines.append("Inlet PSI:     {}".format(sizing_result["inlet_pressure_psi"]))
    lines.append("Total load:    {:.1f} MBH  |  {} fixtures".format(
        total_mbh, len(fixture_nodes)))
    lines.append("")

    sizes = sizing_result["sizes"]
    lines.append("=== PIPE SIZES ({} segments) ===".format(len(sizes)))

    for detail in sizing_result["segment_detail"]:
        edge = graph.edges.get(detail["pipe_id"])
        length_str = "{:.1f}'".format(edge.length_feet) if edge else "?"

        # Downstream fixture label
        fixture_label = ""
        if edge and edge.to_node_id is not None:
            fixtures = _downstream_fixtures(edge.to_node_id, graph)
            if len(fixtures) == 1:
                fixture_label = "  -> {}".format(fixtures[0])
            elif len(fixtures) > 1:
                fixture_label = "  -> trunk ({}: {})".format(
                    len(fixtures), ", ".join(fixtures[:3]) +
                    (" ..." if len(fixtures) > 3 else ""))

        note = "  ({})".format(detail["note"]) if detail["note"] else ""
        lines.append(
            "  [{}]  {:.1f} MBH  ->  {}\"  "
            "(cap {} MBH)  {}{}{}".format(
                detail["pipe_id"],
                detail["demand_mbh"],
                detail["selected_size"],
                int(detail["capacity_mbh"]),
                length_str,
                fixture_label,
                note))

    return "\n".join(lines)


def format_stub_report(skipped_stubs):
    """Return a formatted section listing fixture stub pipes that were skipped.

    Args:
        skipped_stubs: list of dicts with keys:
            pipe_id, fixture_name, demand_mbh, recommended_size

    Returns:
        str
    """
    if not skipped_stubs:
        return ""

    lines = ["=== FIXTURE STUB PIPES (not written - manually resize these) ==="]
    for s in skipped_stubs:
        lines.append(
            "  [{}]  {}  {:.1f} MBH  ->  recommend {}\"".format(
                s["pipe_id"],
                s["fixture_name"],
                s["demand_mbh"],
                s["recommended_size"]))
    return "\n".join(lines)
