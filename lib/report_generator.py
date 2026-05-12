# report_generator.py
# Takes the NetworkGraph from pipe_graph.py and produces:
#   1. Diagnostic report JSON — debugging snapshot of what the traversal found
#   2. One-line diagram data — formatted segments and fixture labels
#
# IronPython 2.7

import datetime
import shared_params
import revit_helpers


def generate_diagnostic_report(graph, origin_element, revit_version, pyrevit_version):
    """Build the full diagnostic report dict from the network graph.

    Args:
        graph: NetworkGraph from pipe_graph.build_network()
        origin_element: The user-selected meter Revit element
        revit_version: str — e.g. "Revit 2024"
        pyrevit_version: str — e.g. "4.8.x"

    Returns:
        dict — the complete diagnostic report, ready for json.dumps()
    """

    report = {}

    # -------------------------------------------------------------------------
    # METADATA
    # -------------------------------------------------------------------------
    report["report_metadata"] = {
        "schema_version":       shared_params.REPORT_SCHEMA_VERSION,
        "tool_name":            shared_params.TOOL_NAME,
        "phase":                shared_params.PHASE,
        "timestamp":            datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "revit_version":        revit_version,
        "pyrevit_version":      pyrevit_version,
    }

    # -------------------------------------------------------------------------
    # SYSTEM ORIGIN — meter element
    # -------------------------------------------------------------------------
    origin_id = origin_element.Id.IntegerValue
    origin_node = graph.nodes.get(origin_id)
    origin_connectors = revit_helpers.get_connectors(origin_element)

    # Validation checks on the meter
    has_out = any(c["direction"] == "Out" for c in origin_connectors)
    out_connected = any(
        c["direction"] == "Out" and c["is_connected"] for c in origin_connectors)
    has_in = any(c["direction"] == "In" for c in origin_connectors)

    report["system_origin"] = {
        "element_id":       origin_id,
        "family_name":      _safe_family_name(origin_element),
        "location_xyz":     revit_helpers.get_element_location(origin_element),
        "connector_count":  len(origin_connectors),
        "connectors":       origin_connectors,
        "validation": {
            "has_out_connector":        has_out,
            "out_connector_connected":  out_connected,
            "has_in_connector":         has_in,
            "result":   "PASS" if (has_out and out_connected) else "FAIL"
        }
    }

    # -------------------------------------------------------------------------
    # FIXTURES
    # -------------------------------------------------------------------------
    fixtures = []
    for node in graph.nodes.values():
        if node.is_gas_fixture:
            fixtures.append({
                "element_id":       node.element_id,
                "fixture_name":     node.fixture_name,
                "family_name":      node.family_name,
                "location_xyz":     node.location_xyz,
                "gas_load_mbh":     node.gas_load_mbh,
                "gas_load_cfh":     node.gas_load_mbh,  # 1 MBH = 1 CFH at 1000 BTU/cf
                "connector_count":  node.connector_count,
                "validation": {
                    "has_gas_load":     node.gas_load_mbh > 0,
                    "has_fixture_name": node.fixture_name not in ("", "UNNAMED"),
                    "result": "PASS" if (node.gas_load_mbh > 0) else "WARN — load is 0"
                }
            })

    report["fixtures_found"] = fixtures

    # -------------------------------------------------------------------------
    # PIPES
    # -------------------------------------------------------------------------
    pipes = []
    for edge in graph.edges.values():
        pipes.append({
            "element_id":       edge.element_id,
            "diameter_inches":  round(edge.diameter_inches, 4),
            "length_feet":      round(edge.length_feet, 2),
            "from_node_id":     edge.from_node_id,
            "to_node_id":       edge.to_node_id,
            "cumulative_load_mbh": round(edge.cumulative_load_mbh, 2)
        })

    report["pipes_found"] = pipes

    # -------------------------------------------------------------------------
    # FITTINGS / TEES
    # -------------------------------------------------------------------------
    fittings = []
    for node in graph.nodes.values():
        if node.node_type in ("tee", "fitting"):
            fittings.append({
                "element_id":       node.element_id,
                "family_name":      node.family_name,
                "node_type":        node.node_type,
                "connector_count":  node.connector_count,
                "location_xyz":     node.location_xyz,
                "is_branch_point":  node.node_type == "tee"
            })

    report["fittings_found"] = fittings

    # -------------------------------------------------------------------------
    # NETWORK GRAPH — adjacency summary
    # -------------------------------------------------------------------------
    nodes_out = []
    for node in graph.nodes.values():
        nodes_out.append({
            "node_id":              node.element_id,
            "node_type":            node.node_type,
            "family_name":          node.family_name,
            "cumulative_load_mbh":  round(node.cumulative_load_mbh, 2)
        })

    edges_out = []
    for edge in graph.edges.values():
        edges_out.append({
            "edge_id":              edge.element_id,
            "from_node":            edge.from_node_id,
            "to_node":              edge.to_node_id,
            "length_feet":          round(edge.length_feet, 2),
            "diameter_inches":      round(edge.diameter_inches, 4),
            "cumulative_load_mbh":  round(edge.cumulative_load_mbh, 2)
        })

    report["network_graph"] = {
        "node_count":   len(graph.nodes),
        "edge_count":   len(graph.edges),
        "nodes":        nodes_out,
        "edges":        edges_out
    }

    # -------------------------------------------------------------------------
    # LONGEST RUN
    # -------------------------------------------------------------------------
    if graph.longest_run:
        lr = graph.longest_run
        report["longest_run"] = {
            "total_length_feet":    round(lr["total_length_feet"], 2),
            "path_element_ids":     lr["path_element_ids"],
            "farthest_fixture_id":  lr["farthest_fixture_id"],
            "farthest_fixture_name": lr["farthest_fixture_name"]
        }
    else:
        report["longest_run"] = None

    # -------------------------------------------------------------------------
    # SYSTEM SUMMARY
    # -------------------------------------------------------------------------
    total_load = sum(
        n.gas_load_mbh for n in graph.nodes.values() if n.is_gas_fixture)
    total_pipes = len(graph.edges)
    longest_ft = graph.longest_run["total_length_feet"] if graph.longest_run else 0.0

    report["system_summary"] = {
        "total_fixtures":       len(fixtures),
        "total_load_mbh":       round(total_load, 2),
        "total_load_cfh":       round(total_load, 2),
        "total_pipe_segments":  total_pipes,
        "longest_run_feet":     round(longest_ft, 2),
        "specific_gravity":     shared_params.SPECIFIC_GRAVITY
    }

    # -------------------------------------------------------------------------
    # ONE-LINE DIAGRAM DATA
    # -------------------------------------------------------------------------
    report["one_line_data"] = _build_one_line_data(graph, total_load, longest_ft)

    # -------------------------------------------------------------------------
    # DISCONNECTED ELEMENTS
    # -------------------------------------------------------------------------
    report["disconnected_elements"] = {
        "count":    len(graph.disconnected),
        "elements": graph.disconnected
    }

    # -------------------------------------------------------------------------
    # TRAVERSAL LOG
    # -------------------------------------------------------------------------
    report["traversal_log"] = graph.traversal_log

    # -------------------------------------------------------------------------
    # VALIDATION SUMMARY
    # -------------------------------------------------------------------------
    report["validation_summary"] = _build_validation_summary(
        graph, fixtures, origin_connectors, total_load)

    return report


# =============================================================================
# ONE-LINE DIAGRAM DATA
# =============================================================================

def _build_one_line_data(graph, total_load_mbh, longest_run_ft):
    """Build formatted one-line diagram data from the graph."""

    segments = []
    for edge in graph.edges.values():
        d = edge.diameter_inches
        size_str = _format_pipe_size(d)
        length_str = "{}\'".format(int(round(edge.length_feet)))
        label_line1 = "{}G, {}".format(size_str, length_str)
        label_line2 = "{} MBH".format(round(edge.cumulative_load_mbh, 1))

        segments.append({
            "edge_id":          edge.element_id,
            "from_node":        edge.from_node_id,
            "to_node":          edge.to_node_id,
            "pipe_label_line1": label_line1,
            "pipe_label_line2": label_line2,
            "diameter_inches":  round(edge.diameter_inches, 4),
            "length_feet":      round(edge.length_feet, 2),
            "cumulative_load_mbh": round(edge.cumulative_load_mbh, 2)
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
        "line4": "TOTAL CONNECTED LOAD: {} MBH".format(round(total_load_mbh, 1)),
        "line5": "TOTAL DEVELOPED LENGTH: {}'".format(int(round(longest_run_ft)))
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
    """Run all validation checks and return a summary dict."""

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

    # Meter checks
    has_out = any(c["direction"] == "Out" for c in origin_connectors)
    out_connected = any(
        c["direction"] == "Out" and c["is_connected"] for c in origin_connectors)
    _check("Meter has Out connector", has_out)
    _check("Meter Out connector is connected to piping", out_connected)

    # Fixture checks
    _check("At least one fixture found", len(fixtures) > 0)
    _check("Total load > 0 MBH", total_load > 0)

    fixtures_missing_load = [
        f for f in fixtures if f["gas_load_mbh"] <= 0]
    _check(
        "All fixtures have GAS_LOAD_MBH > 0",
        len(fixtures_missing_load) == 0,
        warn_not_error=True
    )

    fixtures_missing_name = [
        f for f in fixtures
        if f["fixture_name"] in ("", "UNNAMED", None)]
    _check(
        "All fixtures have FIXTURE_NAME",
        len(fixtures_missing_name) == 0,
        warn_not_error=True
    )

    # Graph checks
    _check("No disconnected elements", len(graph.disconnected) == 0,
           warn_not_error=True)
    _check("Longest run identified", graph.longest_run is not None)

    ready = len(errors) == 0

    return {
        "total_checks":     len(checks),
        "passed":           sum(1 for c in checks if c["result"] == "PASS"),
        "warnings":         len(warnings),
        "errors":           len(errors),
        "ready_for_sizing": ready,
        "checks":           checks,
        "warnings_list":    warnings,
        "errors_list":      errors
    }


# =============================================================================
# FORMATTING HELPERS
# =============================================================================

def _format_pipe_size(diameter_inches):
    """Convert a decimal diameter in inches to a fractional string.

    Examples:
        0.5  → "1/2\""
        0.75 → "3/4\""
        1.0  → "1\""
        1.25 → "1-1/4\""
        2.5  → "2-1/2\""
    """
    size_map = {
        0.5:    '1/2"',
        0.75:   '3/4"',
        1.0:    '1"',
        1.25:   '1-1/4"',
        1.5:    '1-1/2"',
        2.0:    '2"',
        2.5:    '2-1/2"',
        3.0:    '3"',
        4.0:    '4"',
        5.0:    '5"',
        6.0:    '6"',
        8.0:    '8"',
        10.0:   '10"',
        12.0:   '12"',
    }

    # Find the closest nominal size
    closest = min(size_map.keys(), key=lambda k: abs(k - diameter_inches))
    return size_map[closest]


def _safe_family_name(element):
    """Return family name or fallback string."""
    try:
        return element.Symbol.Family.Name
    except Exception:
        try:
            return element.GetType().Name
        except Exception:
            return "Unknown"
