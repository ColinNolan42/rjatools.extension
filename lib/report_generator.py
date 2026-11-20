# report_generator.py
# Takes the NetworkGraph from pipe_graph.py and produces:
#   1. format_diagnostic_output()  -  formatted string printed to PyRevit output window
#
# IronPython 2.7

import datetime
import shared_params
import revit_helpers
import pipe_graph


# =============================================================================
# DIAGNOSTIC OUTPUT  -  printed to PyRevit output window, copy/paste to debug
# =============================================================================

def format_diagnostic_output(graph, origin_element, mid_stream_prv=None,
                             tie_in_check=None):
    """Return a formatted diagnostic string for printing to the PyRevit output window.

    Args:
        graph: NetworkGraph from pipe_graph.build_network()
        origin_element: The user-selected meter, or on a tie in the first
            new pipe.
        mid_stream_prv: see gas_validation().
        tie_in_check: on a pipe pick, what tie_in.check_gas_start()
            returned.

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
    tie_in_pick = _is_tie_in(graph)
    # n=33: a pipe pick is the tie in pipe, not a meter
    section("SYSTEM ORIGIN - TIE-IN PIPE (first new pipe, not a meter)"
            if tie_in_pick else "SYSTEM ORIGIN")
    origin_id = revit_helpers.eid_int(origin_element.Id)
    # piping connectors only: an electrical connector on the meter family
    # is not listed or counted (n=30)
    origin_connectors = revit_helpers.get_piping_connectors(origin_element)
    location = revit_helpers.get_element_location(origin_element)

    lines.append("Element ID:  {}".format(origin_id))
    lines.append("Family:      {}".format(_safe_family_name(origin_element)))
    lines.append("Location:    {}".format(location))
    lines.append("Connectors:  {}".format(len(origin_connectors)))
    for c in origin_connectors:
        lines.append("  connector  direction={}  connected={}".format(
            c["direction"], c["is_connected"]))
    if tie_in_pick:
        if tie_in_check is not None:
            lines.append("Tie-in pipe validation: {}  {}".format(
                "PASS" if tie_in_check.get("ok") else "FAIL",
                tie_in_check.get("message", "")))
        else:
            lines.append("Tie-in pipe validation: not run (meter checks "
                         "skipped on a pipe pick)")
    else:
        has_out = any(c["direction"] == "Out" for c in origin_connectors)
        out_connected = any(c["direction"] == "Out" and c["is_connected"]
                            for c in origin_connectors)
        lines.append("Meter validation: {}".format(
            "PASS" if (has_out and out_connected) else "FAIL"))

    # -------------------------------------------------------------------------
    # FIXTURES
    # -------------------------------------------------------------------------
    fixtures = pipe_graph.gas_fixtures(graph)
    error_ids = pipe_graph.data_error_ids(graph)
    section("FIXTURES ({} found)".format(len(fixtures)))
    if not fixtures:
        lines.append("  NONE")
    for n in fixtures:
        # the walk's data errors are the one record (graph.data_errors)
        result = ("DATA-ERROR-load=0" if n.element_id in error_ids
                  else "PASS")
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
        lines.append("  [{}]  {}\"  {:.2f} FT  from={} to={}  cumulative={:.1f} MBH".format(
            e.element_id,
            round(e.diameter_inches, 4),
            e.length_feet,
            e.from_node_id,
            to_display,
            e.cumulative_load_mbh))
        start = e.start_xyz
        end   = e.end_xyz
        if start:
            lines.append("    start=({:.3f}, {:.3f}, {:.3f})".format(
                float(start[0]), float(start[1]), float(start[2])))
        if end:
            lines.append("    end=({:.3f}, {:.3f}, {:.3f})".format(
                float(end[0]), float(end[1]), float(end[2])))

    # -------------------------------------------------------------------------
    # FITTINGS
    # -------------------------------------------------------------------------
    fittings = [n for n in graph.nodes.values()
                if n.node_type in ("tee", "fitting", "elbow")]
    section("FITTINGS ({} found)".format(len(fittings)))
    if not fittings:
        lines.append("  NONE")
    for n in fittings:
        loc = n.location_xyz
        loc_str = ""
        if loc:
            loc_str = "  loc=({:.3f}, {:.3f}, {:.3f})".format(
                float(loc[0]), float(loc[1]), float(loc[2]))
        lines.append("  [{}]  {}  connectors={}  branch_point={}{}".format(
            n.element_id, n.node_type, n.connector_count,
            n.node_type == "tee", loc_str))

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
        lines.append("  Elbows:        {}  ({:.1f}' equiv, 5 ft each)".format(
            lr["elbow_count"], lr["elbow_equiv_length_feet"]))
        lines.append("  Tees:          {}  ({:.1f}' equiv, 3 ft each)".format(
            lr.get("tee_count", 0), lr.get("tee_equiv_length_feet", 0.0)))
        lines.append("  Farthest:       {} (ID {})".format(
            lr["farthest_fixture_name"], lr["farthest_fixture_id"]))
        lines.append("  Path IDs:       {}".format(lr["path_element_ids"]))
    else:
        lines.append("  NOT FOUND")
    for label, run in zone_run_rows(graph):
        lines.append("  {}: {}".format(label, zone_run_text(run)))

    # -------------------------------------------------------------------------
    # SYSTEM SUMMARY
    # -------------------------------------------------------------------------
    total_load = pipe_graph.total_connected_mbh(graph)
    longest_ft = graph.longest_run["total_length_feet"] if graph.longest_run else 0.0

    section("SYSTEM SUMMARY")
    lines.append("  Fixtures:       {}".format(len(fixtures)))
    lines.append("  Total load:     {:.1f} MBH".format(float(total_load)))
    lines.append("  Pipe segments:  {}".format(len(graph.edges)))
    lines.append("  Longest run:    {:.1f}'".format(float(longest_ft)))
    lines.append("  Spec gravity:   {}".format(shared_params.SPECIFIC_GRAVITY))

    # -------------------------------------------------------------------------
    # DISCONNECTED ELEMENTS
    # -------------------------------------------------------------------------
    section("DISCONNECTED ELEMENTS ({})".format(len(graph.disconnected)))
    if not graph.disconnected:
        lines.append("  None")
    for d in graph.disconnected:
        lines.append("  {}".format(d))

    appliances = existing_appliance_rows(graph)
    if appliances:
        section("EXISTING APPLIANCES ON NEW PIPING ({}, load counted)".format(
            len(appliances)))
        for eid, name in appliances:
            lines.append("  [{}]  {}".format(eid, name))
    tapped = list(getattr(graph, "tapped_pipes", None) or [])
    if tapped:
        section("TAPPED PIPES ({}, each branch charged the whole pipe)".format(
            len(tapped)))
        lines.append("  {}".format(", ".join(str(i) for i in tapped)))

    # -------------------------------------------------------------------------
    # VALIDATION SUMMARY
    # -------------------------------------------------------------------------
    validation = gas_validation(graph, origin_connectors, mid_stream_prv)

    section("VALIDATION SUMMARY")
    for check in validation["checks"]:
        lines.append("  {}  {}".format(check["result"], check["check"]))
    for d in validation["data_errors"]:
        lines.append("  DATA ERROR  {}".format(d))
    lines.append("")
    lines.append("  ready_for_sizing: {}".format(validation["ready_for_sizing"]))
    if validation["errors_list"]:
        lines.append("  ERRORS:   {}".format(", ".join(validation["errors_list"])))
    if validation["warnings_list"]:
        lines.append("  WARNINGS: {}".format(", ".join(validation["warnings_list"])))

    return "\n".join(lines)


# =============================================================================
# VALIDATION SUMMARY
# =============================================================================

def _is_tie_in(graph):
    return bool(getattr(graph, "origin_is_tie_in", False))


def data_error_lines(graph):
    """One line per gas fixture whose load is a data error (n=10/11), from
    the walk's graph.data_errors ('zero_load', 'bad_load', each with a
    message), the one record. pipe_graph records every gas fixture whose
    GAS_LOAD_MBH is 0, unreadable or negative, so the load is not tested
    again here."""
    lines = []
    for d in getattr(graph, "data_errors", None) or []:
        lines.append("{} (element {}): {}.".format(
            d.get("name") or "UNNAMED", d.get("element_id"),
            str(d.get("message") or d.get("kind") or "load not usable")
            .rstrip(".")))
    return lines


def existing_appliance_rows(graph):
    """[(element id, name)] of existing phase gas appliances a tie in run
    reached on new piping (their load is counted, never resized)."""
    rows = []
    for eid in getattr(graph, "existing_appliances", None) or []:
        node = graph.nodes.get(eid)
        name = None
        if node is not None:
            name = node.fixture_name or node.family_name
        rows.append((eid, name or "-"))
    return rows


def gas_validation(graph, origin_connectors=None, mid_stream_prv=None):
    """The one set of gas checks behind both the Diagnose summary and the
    VALIDATION SUMMARY block, so ready_for_sizing and the summary agree
    (n=42). Each check carries a message for the summary.

    Args:
        graph: NetworkGraph from pipe_graph.build_network().
        origin_connectors: revit_helpers.get_piping_connectors() of the
            picked meter. Not used on a tie in (pipe pick): the meter checks
            are skipped there (n=33), the pick was checked by
            tie_in.check_gas_start() as Size Gas does.
        mid_stream_prv: whether Size Gas would run with Mid Stream PRV
            ticked. None (Diagnose, zero input, does not know) or False
            makes nested mid-stream PRVs a warning; True makes them an
            error, as Size Gas stops on them only then.

    Returns:
        dict: checks [{check, result PASS/WARN/FAIL, message}], errors and
        warnings (messages), data_errors (one line per 0 MBH fixture),
        errors_list / warnings_list (check names), ready_for_sizing, and
        the counts.
    """
    checks = []
    warnings = []
    errors = []
    error_names = []
    warning_names = []

    def _check(name, passed, message, warn_not_error=False):
        result = "PASS" if passed else ("WARN" if warn_not_error else "FAIL")
        checks.append({"check": name, "result": result, "message": message})
        if not passed:
            if warn_not_error:
                warnings.append(message)
                warning_names.append(name)
            else:
                errors.append(message)
                error_names.append(name)

    fixtures = pipe_graph.gas_fixtures(graph)
    total_load = pipe_graph.total_connected_mbh(graph)

    if not _is_tie_in(graph):
        conns = origin_connectors or []
        has_out = any(c["direction"] == "Out" for c in conns)
        out_connected = any(
            c["direction"] == "Out" and c["is_connected"] for c in conns)
        _check("Meter has Out connector", has_out,
               "The meter has no Out connector.")
        _check("Meter Out connector is connected to piping", out_connected,
               "The meter's Out connector is not connected to piping.")
    _check("At least one fixture found", len(fixtures) > 0,
           "No gas fixtures found (IS_GAS_FIXTURE = Yes). Common causes: "
           "pipe caps replaced fixture families at terminations, fixture "
           "families not physically connected to the piping, or "
           "IS_GAS_FIXTURE not set to Yes.")
    _check("Total load > 0 MBH", total_load > 0, "Total load is 0 MBH.")

    data_errors = data_error_lines(graph)
    _check("No fixture has 0 MBH or an unreadable load (data error)",
           len(data_errors) == 0,
           "{} gas fixture(s) have a load of 0 MBH or one that cannot be "
           "read, a data error, listed under Data errors.".format(
               len(data_errors)), warn_not_error=True)

    tapped = list(getattr(graph, "tapped_pipes", None) or [])
    _check("No pipe carries a tap", not tapped,
           "{} pipe(s) carry a tap (more than one connection downstream): "
           "{}. Each branch is charged the whole pipe length and the pipe "
           "carries every branch's load (conservative). Break the pipe at "
           "the tap with a tee.".format(
               len(tapped), ", ".join(str(i) for i in tapped)),
           warn_not_error=True)

    missing_name = [n for n in fixtures
                    if n.fixture_name in ("", "UNNAMED", None)]
    _check("All fixtures have Fixture_Name", len(missing_name) == 0,
           "{} fixture(s) have no FIXTURE_NAME.".format(len(missing_name)),
           warn_not_error=True)

    _check("No disconnected elements", len(graph.disconnected) == 0,
           "{} disconnected element(s) found.".format(len(graph.disconnected)),
           warn_not_error=True)
    _check("Longest run identified", graph.longest_run is not None,
           "Longest run could not be determined.")

    nested = list(getattr(graph, "nested_prv_ids", None) or [])
    ids = ", ".join(str(i) for i in nested)
    if mid_stream_prv:
        _check("No nested mid-stream PRVs", not nested,
               "Mid-stream PRV(s) {} sit downstream of another mid-stream "
               "PRV. Size Gas supports a single step down only.".format(ids))
    else:
        _check("No nested mid-stream PRVs", not nested,
               "Mid-stream PRV(s) {} sit downstream of another mid-stream "
               "PRV. Size Gas stops on this when Mid Stream PRV is ticked "
               "(a single step down only); unticked, the whole system is "
               "sized on the meter's table.".format(ids),
               warn_not_error=True)

    return {
        "total_checks":     len(checks),
        "passed":           sum(1 for c in checks if c["result"] == "PASS"),
        "warnings":         len(warnings),
        "errors":           len(errors),
        "ready_for_sizing": len(errors) == 0,
        "checks":           checks,
        "warnings_list":    warning_names,
        "errors_list":      error_names,
        "error_messages":   errors,
        "warning_messages": warnings,
        "data_errors":      data_errors,
    }


# =============================================================================
# FORMATTING HELPERS
# =============================================================================

def gas_summary_counts(graph, kind_of):
    """Tee, other fitting and accessory counts for the Diagnose summary
    (n=36). kind_of(node) returns "PRV", "isolation valve" or None. An
    accessory is counted once, as an accessory, whether its family has two
    connectors (typed fitting) or three or more (typed tee, a regulator
    with a vent), so no count is understated or negative.

    Returns:
        dict: tees, fittings_other (counts), accessories [(node, kind)],
        prv_count, iso_count.
    """
    accessories = []
    tees = 0
    others = 0
    for n in graph.nodes.values():
        if n.node_type not in ("fitting", "tee"):
            continue
        kind = kind_of(n)
        if kind:
            accessories.append((n, kind))
        elif n.node_type == "tee":
            tees += 1
        else:
            others += 1
    return {"tees": tees, "fittings_other": others,
            "accessories": accessories,
            "prv_count": sum(1 for _n, k in accessories if k == "PRV"),
            "iso_count": sum(1 for _n, k in accessories
                             if k == "isolation valve")}


def zone_run_rows(graph):
    """[(label, run dict or None)] for each zone Size Gas sizes on its own
    run when a mid-stream PRV exists (graph.zone_runs), meter or tie in
    first, then each PRV. [] with no mid-stream PRV (n=62)."""
    zone_runs = getattr(graph, "zone_runs", None) or {}
    if not zone_runs:
        return []
    rows = []
    origin = getattr(graph, "origin_id", None)
    # Diagnose does not know whether the box will be ticked: the rows say
    # when Size Gas uses these runs.
    when = "(used when Mid Stream PRV is ticked)"
    first = ("Tie-in zone run " if _is_tie_in(graph)
             else "Meter zone run ") + when
    if origin in zone_runs:
        rows.append((first, zone_runs[origin]))
    for pid in getattr(graph, "midstream_prv_ids", None) or []:
        if pid in zone_runs:
            rows.append(("After PRV {} run {}".format(pid, when),
                         zone_runs[pid]))
    return rows


def zone_run_text(run):
    """One line for a zone run: '180.0 ft to RTU-2', or 'not found'."""
    if not run:
        return "not found"
    return "{:.1f} ft to {}".format(
        float(run.get("total_length_feet") or 0.0),
        run.get("farthest_fixture_name") or "-")


# Characters that break or hide a markdown table cell, as HTML entities so
# no markdown escaping rule is relied on (n=63).
_MD_CELL_ENTITIES = (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"),
                     ("|", "&#124;"), ("*", "&#42;"), ("_", "&#95;"),
                     ("`", "&#96;"), ("[", "&#91;"), ("]", "&#93;"),
                     ("\\", "&#92;"))


def md_cell(value):
    """Model text made safe for one print_md table cell: HTML special
    characters and markdown table and emphasis characters become entities,
    line breaks become spaces. None becomes ''."""
    if value is None:
        return ""
    try:
        text = value if isinstance(value, basestring) else str(value)  # noqa: F821
    except NameError:   # CPython 3
        text = value if isinstance(value, str) else str(value)
    for ch, ent in _MD_CELL_ENTITIES:
        text = text.replace(ch, ent)
    return text.replace("\r", " ").replace("\n", " ")


def _safe_family_name(element):
    """Family name through revit_helpers.get_element_family_name, else the
    element class name, else 'Unknown'."""
    name = revit_helpers.get_element_family_name(element)
    if name:
        return name
    try:
        return element.GetType().Name
    except Exception:
        return "Unknown"
