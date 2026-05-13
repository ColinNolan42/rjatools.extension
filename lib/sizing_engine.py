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

        demand_cfh = edge.cumulative_load_mbh  # 1 MBH = 1 CFH

        # Zero demand: assign minimum available pipe size
        if demand_cfh <= 0:
            selected = pipe_sizes[0]
            capacity_at_size = gas_tables.get_capacity(
                table_id, longest_run_ft, selected)
            segment_detail.append({
                "pipe_id":       edge.element_id,
                "demand_cfh":    0.0,
                "selected_size": selected,
                "capacity_cfh":  capacity_at_size,
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
                if capacity >= demand_cfh:
                    selected = size
                    selected_capacity = capacity
                    break
            except ValueError:
                # Size not available at this length - try the next size up
                continue

        if selected is None:
            sizing_errors.append(
                "Pipe {}: demand {:.1f} CFH exceeds max table capacity "
                "at {:.0f} ft in Table {}.".format(
                    edge.element_id, demand_cfh, table_length_used, table_id))
            continue

        sizes[edge.element_id] = selected
        segment_detail.append({
            "pipe_id":       edge.element_id,
            "demand_cfh":    round(demand_cfh, 1),
            "selected_size": selected,
            "capacity_cfh":  selected_capacity,
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

    lines.append("=== SIZING RESULTS ===")
    lines.append("Table:         {}".format(sizing_result["table_id"]))
    lines.append("Table row:     {} ft  (longest run {:.1f} ft rounded up)".format(
        sizing_result["table_length_used_ft"],
        sizing_result["longest_run_ft"]))
    lines.append("Pipe material: {}".format(sizing_result["pipe_material"]))
    lines.append("Inlet PSI:     {}".format(sizing_result["inlet_pressure_psi"]))
    lines.append("")

    sizes = sizing_result["sizes"]
    lines.append("=== PIPE SIZES ({} segments) ===".format(len(sizes)))

    for detail in sizing_result["segment_detail"]:
        edge = graph.edges.get(detail["pipe_id"])
        length_str = "{:.1f}'".format(edge.length_feet) if edge else "?"
        note = "  ({})".format(detail["note"]) if detail["note"] else ""
        lines.append(
            "  [{}]  {:.1f} CFH  ->  {}\"  "
            "(cap {:.0f} CFH)  {}{}".format(
                detail["pipe_id"],
                detail["demand_cfh"],
                detail["selected_size"],
                detail["capacity_cfh"],
                length_str,
                note))

    return "\n".join(lines)
