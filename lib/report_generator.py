# report_generator.py
# Takes the NetworkGraph from pipe_graph.py and produces:
#   1. format_diagnostic_output()  -  formatted string printed to PyRevit output window
#   2. generate_one_line_data()    -  structured segment and fixture label data
#
# IronPython 2.7

import datetime
import shared_params
import revit_helpers


# =============================================================================
# DIAGNOSTIC OUTPUT  -  printed to PyRevit output window, copy/paste to debug
# =============================================================================

def format_diagnostic_output(graph, origin_element):
    """Return a formatted diagnostic string for printing to the PyRevit output window.

    Args:
        graph: NetworkGraph from pipe_graph.build_network()
        origin_element: The user-selected meter Revit element

    Returns:
        str
    """
    lines = []

    def section(title):
        lines.append("")
        lines.append("=== {} ===".format(title))

    # -------------------------------------------------------------------------
    # HEADER
    # -------------------------------------------------------------------------
    lines.append("=== DIAGNOSTIC REPORT ===")
    lines.append("Timestamp: {}".format(
        datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")))

    # -------------------------------------------------------------------------
    # SYSTEM ORIGIN
    # -------------------------------------------------------------------------
    section("SYSTEM ORIGIN")
    origin_id = origin_element.Id.IntegerValue
    origin_connectors = revit_helpers.get_connectors(origin_element)
    location = revit_helpers.get_element_location(origin_element)

    has_out = any(c["direction"] == "Out" for c in origin_connectors)
    out_connected = any(
        c["direction"] == "Out" and c["is_connected"] for c in origin_connectors)
    meter_valid = "PASS" if (has_out and out_connected) else "FAIL"

    lines.append("Element ID:  {}".format(origin_id))
    lines.append("Family:      {}".format(_safe_family_name(origin_element)))
    lines.append("Location:    {}".format(location))
    lines.append("Connectors:  {}".format(len(origin_connectors)))
    for c in origin_connectors:
        lines.append("  connector  direction={}  connected={}".format(
            c["direction"], c["is_connected"]))
    lines.append("Meter validation: {}".format(meter_valid))

    # -------------------------------------------------------------------------
    # FIXTURES
    # -------------------------------------------------------------------------
    fixtures = [n for n in graph.nodes.values() if n.is_gas_fixture]
    section("FIXTURES ({} found)".format(len(fixtures)))
    if not fixtures:
        lines.append("  NONE")
    for n in fixtures:
        result = "PASS" if n.gas_load_mbh > 0 else "WARN-load=0"
        lines.append("  [{}]  {}  {:.1f} MBH  loc={}  {}".format(
            n.element_id, n.fixture_name, n.gas_load_mbh, n.location_xyz, result))

    # -------------------------------------------------------------------------
    # PIPES
    # -------------------------------------------------------------------------
    section("PIPES ({} found)".format(len(graph.edges)))
    if not graph.edges:
        lines.append("  NONE")
    for e in graph.edges.values():
        to_display = "OPEN-END" if e.to_node_id is None else str(e.to_node_id)
        lines.append("  [{}]  {}\"  {:.1f}'  from={} to={}  cumulative={:.1f} MBH".format(
            e.element_id,
            round(e.diameter_inches, 4),
            e.length_feet,
            e.from_node_id,
            to_display,
            e.cumulative_load_mbh))

    # -------------------------------------------------------------------------
    # FITTINGS
    # -------------------------------------------------------------------------
    fittings = [n for n in graph.nodes.values()
                if n.node_type in ("tee", "fitting", "elbow")]
    section("FITTINGS ({} found)".format(len(fittings)))
    if not fittings:
        lines.append("  NONE")
    for n in fittings:
        lines.append("  [{}]  {}  connectors={}  branch_point={}".format(
            n.element_id, n.node_type, n.connector_count,
            n.node_type == "tee"))

    # -------------------------------------------------------------------------
    # NETWORK GRAPH
    # -------------------------------------------------------------------------
    section("NETWORK GRAPH  nodes={}  edges={}".format(
        len(graph.nodes), len(graph.edges)))
    for node in graph.nodes.values():
        lines.append("  node [{}]  {}  {:.1f} MBH".format(
            node.element_id, node.node_type, node.cumulative_load_mbh))
    for edge in graph.edges.values():
        to_display = "OPEN-END" if edge.to_node_id is None else str(edge.to_node_id)
        lines.append("  edge [{}]  {} -> {}  {:.1f}'  {:.1f} MBH".format(
            edge.element_id, edge.from_node_id, to_display,
            edge.length_feet, edge.cumulative_load_mbh))

    # -------------------------------------------------------------------------
    # LONGEST RUN
    # -------------------------------------------------------------------------
    section("LONGEST RUN")
    if graph.longest_run:
        lr = graph.longest_run
        lines.append("  Total:          {:.1f}'".format(lr["total_length_feet"]))
        lines.append("  Pipe length:    {:.1f}'".format(lr["pipe_length_feet"]))
        lines.append("  Elbows:         {}  ({:.1f}' equiv)".format(
            lr["elbow_count"], lr["elbow_equiv_length_feet"]))
        lines.append("  Farthest:       {} (ID {})".format(
            lr["farthest_fixture_name"], lr["farthest_fixture_id"]))
        lines.append("  Path IDs:       {}".format(lr["path_element_ids"]))
    else:
        lines.append("  NOT FOUND")

    # -------------------------------------------------------------------------
    # SYSTEM SUMMARY
    # -------------------------------------------------------------------------
    total_load = sum(n.gas_load_mbh for n in graph.nodes.values() if n.is_gas_fixture)
    longest_ft = graph.longest_run["total_length_feet"] if graph.longest_run else 0.0

    section("SYSTEM SUMMARY")
    lines.append("  Fixtures:       {}".format(len(fixtures)))
    lines.append("  Total load:     {:.1f} MBH".format(total_load))
    lines.append("  Pipe segments:  {}".format(len(graph.edges)))
    lines.append("  Longest run:    {:.1f}'".format(longest_ft))
    lines.append("  Spec gravity:   {}".format(shared_params.SPECIFIC_GRAVITY))

    # -------------------------------------------------------------------------
    # DISCONNECTED ELEMENTS
    # -------------------------------------------------------------------------
    section("DISCONNECTED ELEMENTS ({})".format(len(graph.disconnected)))
    if not graph.disconnected:
        lines.append("  None")
    for d in graph.disconnected:
        lines.append("  {}".format(d))

    # -------------------------------------------------------------------------
    # VALIDATION SUMMARY
    # -------------------------------------------------------------------------
    validation = _build_validation_summary(
        graph, fixtures, origin_connectors, total_load)

    section("VALIDATION SUMMARY")
    for check in validation["checks"]:
        lines.append("  {}  {}".format(check["result"], check["check"]))
    lines.append("")
    lines.append("  ready_for_sizing: {}".format(validation["ready_for_sizing"]))
    if validation["errors_list"]:
        lines.append("  ERRORS:   {}".format(", ".join(validation["errors_list"])))
    if validation["warnings_list"]:
        lines.append("  WARNINGS: {}".format(", ".join(validation["warnings_list"])))

    return "\n".join(lines)


# =============================================================================
# ONE-LINE DIAGRAM DATA
# =============================================================================

def generate_one_line_data(graph):
    """Build structured one-line diagram data from the graph.

    Returns:
        dict with keys: segments, fixtures, notes_block
    """
    total_load = sum(n.gas_load_mbh for n in graph.nodes.values() if n.is_gas_fixture)
    longest_ft = graph.longest_run["total_length_feet"] if graph.longest_run else 0.0

    segments = []
    for edge in graph.edges.values():
        size_str = _format_pipe_size(edge.diameter_inches)
        label_line1 = "{}G, {}'".format(size_str, int(round(edge.length_feet)))
        label_line2 = "{} MBH".format(round(edge.cumulative_load_mbh, 1))
        segments.append({
            "edge_id":              edge.element_id,
            "from_node":            edge.from_node_id,
            "to_node":              edge.to_node_id,
            "pipe_label_line1":     label_line1,
            "pipe_label_line2":     label_line2,
            "diameter_inches":      round(edge.diameter_inches, 4),
            "length_feet":          round(edge.length_feet, 2),
            "cumulative_load_mbh":  round(edge.cumulative_load_mbh, 2)
        })

    fixture_labels = []
    for node in graph.nodes.values():
        if node.is_gas_fixture:
            fixture_labels.append({
                "element_id":           node.element_id,
                "fixture_label_line1":  node.fixture_name,
                "fixture_label_line2":  "{} MBH".format(round(node.gas_load_mbh, 1))
            })

    notes_block = {
        "line1": "CONTRACTOR SHALL SUBMIT APPLICATIONS TO UTILITY AND COORDINATE NEW METER SERVICE",
        "line2": "GAS PIPING SIZED FOR [X] PSI",
        "line3": "MAX PRESSURE LOSS OF [X] PSI PER IFGC TABLE 402.4([X])",
        "line4": "TOTAL CONNECTED LOAD: {} MBH".format(round(total_load, 1)),
        "line5": "TOTAL DEVELOPED LENGTH: {}'".format(int(round(longest_ft)))
    }

    return {
        "segments":     segments,
        "fixtures":     fixture_labels,
        "notes_block":  notes_block
    }


# =============================================================================
# VALIDATION SUMMARY
# =============================================================================

def _build_validation_summary(graph, fixtures, origin_connectors, total_load):
    checks = []
    warnings = []
    errors = []

    def _check(name, passed, warn_not_error=False):
        result = "PASS" if passed else ("WARN" if warn_not_error else "FAIL")
        checks.append({"check": name, "result": result})
        if not passed:
            if warn_not_error:
                warnings.append(name)
            else:
                errors.append(name)

    has_out = any(c["direction"] == "Out" for c in origin_connectors)
    out_connected = any(
        c["direction"] == "Out" and c["is_connected"] for c in origin_connectors)
    _check("Meter has Out connector", has_out)
    _check("Meter Out connector is connected to piping", out_connected)
    _check("At least one fixture found", len(fixtures) > 0)
    _check("Total load > 0 MBH", total_load > 0)

    fixtures_missing_load = [n for n in fixtures if n.gas_load_mbh <= 0]
    _check("All fixtures have gas load > 0",
           len(fixtures_missing_load) == 0, warn_not_error=True)

    fixtures_missing_name = [
        n for n in fixtures if n.fixture_name in ("", "UNNAMED", None)]
    _check("All fixtures have Fixture_Name",
           len(fixtures_missing_name) == 0, warn_not_error=True)

    _check("No disconnected elements",
           len(graph.disconnected) == 0, warn_not_error=True)
    _check("Longest run identified", graph.longest_run is not None)

    return {
        "total_checks":     len(checks),
        "passed":           sum(1 for c in checks if c["result"] == "PASS"),
        "warnings":         len(warnings),
        "errors":           len(errors),
        "ready_for_sizing": len(errors) == 0,
        "checks":           checks,
        "warnings_list":    warnings,
        "errors_list":      errors
    }


# =============================================================================
# FORMATTING HELPERS
# =============================================================================

def _format_pipe_size(diameter_inches):
    """Convert decimal diameter in inches to fractional string (e.g. 2.5 -> '2-1/2"')."""
    size_map = {
        0.5:   '1/2"',
        0.75:  '3/4"',
        1.0:   '1"',
        1.25:  '1-1/4"',
        1.5:   '1-1/2"',
        2.0:   '2"',
        2.5:   '2-1/2"',
        3.0:   '3"',
        4.0:   '4"',
        5.0:   '5"',
        6.0:   '6"',
        8.0:   '8"',
        10.0:  '10"',
        12.0:  '12"',
    }
    closest = min(size_map.keys(), key=lambda k: abs(k - diameter_inches))
    return size_map[closest]


def _safe_family_name(element):
    try:
        return element.Symbol.Family.Name
    except Exception:
        try:
            return element.GetType().Name
        except Exception:
            return "Unknown"
