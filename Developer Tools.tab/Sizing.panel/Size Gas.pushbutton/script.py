# -*- coding: ascii -*-
# Size Gas.pushbutton/script.py
# Phase 2 - IFGC Longest Run Gas Pipe Sizing
# Startup dialog -> user picks the gas meter, or the first new pipe on an
# existing building -> sizing -> writes sizes to Revit
#
# IronPython 2.7 / PyRevit

import os
import sys

from pyrevit import script
from Autodesk.Revit.DB import BuiltInParameter, ElementId

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()
logger = script.get_logger()

# Add lib/ to path
_lib_dir = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

import console_report as cr
import gas_console
import shared_params
import revit_helpers
import pipe_graph
import report_generator
import gas_tables
import sizing_engine
import tie_in
import ui_helpers

TITLE = "Size Gas"


def _emit(html, md):
    """One themed section; the markdown twin prints if print_html fails."""
    cr.emit(output, html, lambda: output.print_md(md), logger, TITLE)


def _say(text, kind="info"):
    """One status or error line as a themed callout."""
    _emit(gas_console.message(text, kind), gas_console.message_md(text, kind))


def _alert(text, kind="error", yes_no=False):
    """The RJA themed popup (ui_helpers.rja_alert), in place of forms.alert."""
    return ui_helpers.rja_alert(TITLE, text, kind, yes_no)


# ---------------------------------------------------------------------------
# Pipe diameter write-back
# Three API approaches tried in order. First success wins for all pipes.
# ---------------------------------------------------------------------------

# Populated on first successful write - avoids re-trying failed approaches
_confirmed_approach = [None]


def _set_pipe_diameter(pipe, nominal_inches):
    """Set a pipe's nominal diameter and report the approach to the window.

    The three-approach write logic now lives in revit_helpers.set_pipe_diameter
    so Size Gas and Size Water share one implementation. This wrapper only records
    the approach that worked.

    The approach that worked is shown in the Summary, a failure in the write
    results table, so nothing prints per pipe.

    Args:
        pipe:           Revit Pipe element
        nominal_inches: float  nominal diameter in inches

    Returns:
        (success: bool, approach_name: str)
    """
    ok, name = revit_helpers.set_pipe_diameter(pipe, nominal_inches)

    if ok and _confirmed_approach[0] is None:
        _confirmed_approach[0] = name

    return ok, name


# ---------------------------------------------------------------------------
# Fitting resize - one size per fitting through revit_helpers.set_fitting_size
# ---------------------------------------------------------------------------

def _new_fit_counts():
    """The counts _resize_fittings fills. failures holds the text lines,
    failure_rows the same failures as data for the console table."""
    return {"resized": 0, "no_sized_pipe": 0, "existing": 0,
            "no_element": 0, "appliance": 0, "replaced": 0,
            "transitions": 0, "transitions_removed": 0, "prv": 0,
            "failures": [], "failure_rows": []}


def _fit_fail(counts, node, family, reason):
    """Record one fitting failure, as a text line and as a table row."""
    counts["failures"].append("Fitting {} ({}): {}".format(
        node.element_id, family, reason))
    counts["failure_rows"].append({"id": node.element_id,
                                   "family": family or "",
                                   "reason": reason})


def _transition_ids(graph):
    """Ids of the fittings that are transitions, read BEFORE the pipes are
    resized: Revit may delete a transition during the resize, after which
    its part type can no longer be read."""
    ids = set()
    for node in graph.nodes.values():
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        if node.is_gas_fixture:
            continue
        if node.is_transition or revit_helpers.is_transition(node.element):
            ids.add(node.element_id)
    return ids


def _resize_fittings(graph, new_sizes, doc, transition_ids=None,
                     existing_sizes=None):
    """Resize the new construction fittings to the pipes sized this run.

    For each fitting the target size is the highest-load connected pipe in
    new_sizes, which corresponds to the main run size at that fitting
    location. A new fitting that also joins existing phase pipes takes the
    largest of all its pipes, the existing ones at their modeled size from
    existing_sizes, so a new tee on an existing main never shrinks the main
    (n=3, Colin 2026-11-20). An existing phase fitting is never resized:
    existing stays the size it is, and the existing pipe audit reports a
    shortfall (Colin, 2026-11-02). The meter and every PRV are never
    resized (n=29).

    The write goes through revit_helpers.set_fitting_size with no
    per-connector map, so only the 'Nominal Radius' parameter is used.
    Connector.Radius writes are not attempted: they disconnect fittings
    (verified 2026-10-12).

    Args:
        graph:     NetworkGraph that was sized.
        new_sizes: {pipe element id: nominal size} for new construction pipes.
        doc:       active Document, for the phase check.

    Returns:
        dict with resized (int), no_sized_pipe (int, no new pipe sized this
        run touches it), existing (int, existing phase, left alone),
        no_element (int), appliance (int, equipment with no size
        parameter), transitions (int, left to Revit), transitions_removed
        (int, Revit removed it and its pipes now join), failures (list of
        str, real write failures) and failure_rows (the same as dicts).

    Transitions are never resized here (Colin, 2026-11-19): a transition
    takes its two sizes from its pipes, and Revit updates or removes it when
    the pipes change. transition_ids is from _transition_ids(), taken before
    the pipe resize.
    """
    counts = _new_fit_counts()
    transition_ids = transition_ids or set()
    done_ids = set()   # replacement fittings already sized this run

    # One guard per fitting: an exception on one fitting is one failure
    # line, never the rollback of every fitting (G7C, 2026-11-04).
    for node in graph.nodes.values():
        # A "joint" (two pipes end to end, no Revit element, keyed by a pipe
        # id) is never in this list, so it is never resized or looked up.
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        try:
            _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                                transition_ids, existing_sizes)
        except Exception as e:
            _fit_fail(counts, node, node.family_name, str(e))
            revit_helpers._log_entry(
                "ERROR", "_resize_fittings", node.element_id,
                "Resize raised: {}".format(str(e)))

    return counts


def _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                        transition_ids=(), existing_sizes=None):
    """Resize one fitting node. See _resize_fittings for the rules.

    The node's element is re-read first. Revit may have deleted it when the
    pipes were resized in the pipe transaction and placed another fitting;
    that replacement is found on the connected pipes and sized instead.
    """
    if node.element is None:
        counts["no_element"] += 1
        return

    # The meter and every PRV keep the size drawn (n=29).
    exempt = sizing_engine.fitting_resize_exempt(node, graph)
    if exempt:
        counts["prv"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Not resized: {}.".format(exempt))
        return

    element  = node.element
    replaced = False
    if not revit_helpers.is_valid_element(element):
        element = revit_helpers.current_element(doc, node.element_id)
    if element is None:
        pipes = [e.pipe for e in graph.edges.values()
                 if node.element_id in (e.from_node_id, e.to_node_id)
                 and e.pipe is not None]
        # A transition Revit removed because its pipes became the same size
        # is the expected result, not a failure: its pipes now join.
        if (node.element_id in transition_ids
                and (revit_helpers.pipes_joined(pipes)
                     or revit_helpers.ends_connected(pipes,
                                                     node.location_xyz))):
            counts["transitions_removed"] += 1
            revit_helpers._log_entry(
                "INFO", "_resize_fittings", node.element_id,
                "Transition removed or replaced by Revit, its pipe ends "
                "are connected.")
            return
        element, _d = revit_helpers.find_replacement_fitting(
            pipes, node.location_xyz, set(graph.nodes.keys()) | done_ids)
        if element is None:
            _fit_fail(counts, node, node.family_name,
                      "deleted by Revit when its pipes were resized, and no "
                      "replacement fitting was found on its pipes.")
            return
        replaced = True
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Deleted by Revit; replacement {} sized instead.".format(
                revit_helpers.eid_int(element.Id)))

    # Transitions are left to Revit (a replacement counts the same way).
    if revit_helpers.is_transition(element):
        counts["transitions"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Transition: size comes from its pipes, left to Revit.")
        return

    if not revit_helpers.is_new_construction(element, doc):
        counts["existing"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Existing phase (or phase unreadable): not resized.")
        return

    # The highest-load new pipe that touches it sets the size, never below
    # an existing pipe it joins; a valve with a transition on each side
    # reads the pipes past them (Horse Creek).
    target_size, via_transition = sizing_engine.fitting_target_size(
        node.element_id, graph, new_sizes, transition_ids, existing_sizes)
    if via_transition:
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "No sized pipe touches it; size {} read through a transition."
            .format(target_size))

    if target_size is None:
        counts["no_sized_pipe"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "No new construction pipe sized this run touches it: not "
            "resized.")
        return

    # An appliance is resized only when it is a cap style family with a
    # 'Nominal Radius' (the cap first, then the stub pipe, STEP 9). Real
    # equipment has no such parameter and is left alone, not a failure.
    if node.is_gas_fixture and _lookup_param(
            element, "Nominal Radius") is None:
        counts["appliance"] += 1
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Appliance with no 'Nominal Radius': not resized.")
        return

    nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(target_size)
    if nominal_inches is None:
        _fit_fail(counts, node, node.family_name,
                  "unrecognised nominal size '{}'.".format(target_size))
        return

    # piping connectors only on a Type swap (n=30): no pipe end is moved
    # onto an electrical connector of a gas valve
    ok, approach = revit_helpers.set_fitting_size(element, nominal_inches,
                                                  piping_only=True)
    if ok:
        counts["resized"] += 1
        if replaced:
            counts["replaced"] += 1
            done_ids.add(revit_helpers.eid_int(element.Id))
    else:
        _fit_fail(counts, node, node.family_name,
                  "could not set size to {}\" ({}).".format(
                      target_size, approach))


def _size_one_stub(s, graph):
    """Size one fixture stub pipe (STEP 9). True when it was written."""
    edge = graph.edges.get(s["pipe_id"])
    if (edge is None or edge.pipe is None
            or not revit_helpers.is_valid_element(edge.pipe)):
        revit_helpers._log_entry(
            "ERROR", "size_gas_stubs", s["pipe_id"],
            "Stub pipe NOT FOUND in graph, or deleted.")
        return False
    nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(
        s["recommended_size"])
    if nominal_inches is None:
        revit_helpers._log_entry(
            "ERROR", "size_gas_stubs", s["pipe_id"],
            "Unrecognised nominal size '{}'.".format(s["recommended_size"]))
        return False
    ok, _approach = _set_pipe_diameter(edge.pipe, nominal_inches)
    return ok


def _lookup_param(element, name):
    """element.LookupParameter(name), or None when it raises (logged)."""
    try:
        return element.LookupParameter(name)
    except Exception as e:
        revit_helpers._log_entry(
            "WARNING", "_lookup_param", revit_helpers.eid_int(element.Id),
            "LookupParameter('{}') raised: {}".format(name, str(e)))
        return None


# ---------------------------------------------------------------------------
# Diameter helpers
# ---------------------------------------------------------------------------

def _get_pipe_nominal_size(pipe, table_id):
    """The nominal size name, from table_id's own size list, for a pipe's
    drawn diameter, or None. A None is logged with the pipe id.

    The table's own list is used so a copper or CSST pipe reads as its own
    size name, never as the steel name sharing its diameter
    (sizing_engine.nominal_for_diameter).
    """
    pipe_id = revit_helpers.eid_int(pipe.Id)
    dia_in = revit_helpers.get_pipe_diameter_inches(pipe)
    if dia_in is None:
        revit_helpers._log_entry(
            "WARNING", "_get_pipe_nominal_size", pipe_id,
            "Diameter NOT READABLE, existing pipe not checked.")
        return None
    nominal = sizing_engine.nominal_for_diameter(table_id, dia_in)
    if nominal is None:
        revit_helpers._log_entry(
            "WARNING", "_get_pipe_nominal_size", pipe_id,
            "Diameter {:.3f} in. matches no size in Table {}.".format(
                float(dia_in), table_id))
    else:
        revit_helpers._log_entry(
            "INFO", "_get_pipe_nominal_size", pipe_id,
            "Diameter {:.3f} in. reads as {} in Table {}.".format(
                float(dia_in), nominal, table_id))
    return nominal


def _audit_existing_pipe(pipe_id, edge, recommended_nom, result,
                         overloaded_list, unchecked_list):
    """Check whether an existing pipe's current size handles the new load.

    If current capacity < demand, or the table publishes no capacity for
    its size at this run, appends a dict to overloaded_list. A pipe whose
    size cannot be matched to the table goes to unchecked_list. Both are
    shown after the sizing transaction.

    Args:
        pipe_id:         int  pipe element ID
        edge:            graph edge (has .pipe, .cumulative_load_mbh)
        recommended_nom: str  the IFGC-computed size (already in result["sizes"])
        result:          dict returned by sizing_engine.size_system()
        overloaded_list: list  mutated in-place
        unchecked_list:  list  mutated in-place
    """
    demand = edge.cumulative_load_mbh
    if demand <= 0:
        return  # no load -> no concern

    # The table and run this pipe was sized on (differs after a PRV).
    ctx = result.get("edge_context", {}).get(pipe_id, {})
    table_id = ctx.get("table_id", result["table_id"])
    run_ft = ctx.get("longest_run_ft", result["longest_run_ft"])

    current_nom = _get_pipe_nominal_size(edge.pipe, table_id)
    if current_nom is None:
        unchecked_list.append({"pipe_id": pipe_id, "table_id": table_id})
        return

    # Demand in the table's units: CFH via heat content (the propane value
    # on 402.4(37)), or the MBH load itself on a propane table 402.4(25) to
    # (36) (Colin, 2026-11-04, 2026-11-20).
    demand_cfh, units = sizing_engine.demand_in_table_units(
        demand, table_id,
        result.get("heat_content_btu_per_cf",
                   shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF),
        result.get("propane_heat_content_btu_per_cf"))
    try:
        current_cap = float(gas_tables.get_capacity(
            table_id, run_ft, current_nom))
    except ValueError as e:
        # The table publishes no capacity for this size at this run.
        revit_helpers._log_entry(
            "WARNING", "_audit_existing_pipe", pipe_id,
            "No published capacity: {}".format(str(e)))
        current_cap = None

    if current_cap is None or current_cap < demand_cfh:
        overloaded_list.append({
            "pipe_id":        pipe_id,
            "current_nom":    current_nom,
            "demand_mbh":     float(demand),
            "current_cap":    current_cap,
            "units":          units,
            "recommended_nom": recommended_nom,
        })


# (no startup dialog helpers - table options are loaded from gas_tables.py)


# ---------------------------------------------------------------------------
# Run checks (small functions so they port to lib/gas_run.py on headless)
# ---------------------------------------------------------------------------

def _propane_heat_content(choice):
    """The propane heat content from the dialog choice, BTU/CF. The key is
    'propane_heat_content_btu_per_cf'; a dialog without the box gives the
    firm default (Colin, 2026-11-20)."""
    value = choice.get("propane_heat_content_btu_per_cf")
    if value is None:
        return shared_params.DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF
    return float(value)


def _table_limit_check(table_id, graph, heat_content, propane_heat_content):
    """Table Intended Use load limit warnings (2024 IFGC Table 402.4(13):
    150 CFH through the house line regulator). The meter carries the whole
    system, so the total is the new fixtures plus the existing load entered
    for Connect into Existing (n=35/52, Colin 2026-11-20)."""
    total_mbh = pipe_graph.total_connected_mbh(graph)
    existing = getattr(graph, "existing", None) or {}
    total_mbh += float(existing.get("load_mbh") or 0.0)
    demand, _units = sizing_engine.demand_in_table_units(
        total_mbh, table_id, heat_content, propane_heat_content)
    return sizing_engine.table_limit_warnings(table_id, demand, total_mbh)


def _existing_pipe_sizes(graph, new_sizes, result, doc):
    """{pipe id: modeled nominal size or None} for the existing phase pipes
    in the graph, read from the drawn diameter on the table of each pipe's
    own system (sizing_engine.edge_table_id, by its run_key). A new fitting
    is never sized below these (n=3). An open ended pipe has no system table
    and is skipped (logged), never read on the meter table."""
    sizes = {}
    for pipe_id, edge in graph.edges.items():
        if pipe_id in new_sizes or edge.pipe is None:
            continue
        if not revit_helpers.is_valid_element(edge.pipe):
            continue
        if revit_helpers.is_new_construction(edge.pipe, doc):
            continue    # a new pipe with no size this run (open end)
        table_id = sizing_engine.edge_table_id(result, edge)
        if table_id is None:
            revit_helpers._log_entry(
                "INFO", "_existing_pipe_sizes", pipe_id,
                "Existing pipe with no downstream node (or no system "
                "table): its size is not read.")
            continue
        sizes[pipe_id] = _get_pipe_nominal_size(edge.pipe, table_id)
    return sizes


def _graph_notices(graph, result):
    """(text, kind) lines for the pipes left unsized because they lead only
    to 0 MBH (data error) fixtures. The data errors and tapped pipes
    themselves print in the pre-sizing warnings (gas_console)."""
    lines = []
    for z in result.get("unsized_zero_load") or []:
        lines.append((
            "Pipe {} leads only to 0 MBH fixture(s) {} and was NOT sized or "
            "written. Fix the fixture load and run Size Gas again.".format(
                z["pipe_id"], ", ".join(str(i) for i in z["fixture_ids"])),
            "error"))
    return lines


# ---------------------------------------------------------------------------
# Writes. Each step is one transaction through revit_helpers.
# run_in_transaction: warnings deleted, an error rolls the step back with no
# Revit dialog, and the Commit status is checked (n=4). A step that did not
# commit reports zero successes, and the steps that depend on it do not run.
# ---------------------------------------------------------------------------

PIPE_STEP = "pipe sizes"
FITTING_STEP = "fittings"
STUB_STEP = "fixture stub pipes"


def _new_pipe_writes():
    return {"success_count": 0, "fail_count": 0, "fail_rows": [],
            "skipped_stubs": [], "existing_skip_count": 0,
            "existing_overloaded": [], "existing_unchecked": [],
            "new_sizes": {}, "committed": False, "error": None}


def _write_one_pipe(w, pipe_id, nominal_size, graph, result, doc):
    """Write or queue one pipe of result["sizes"] (STEP 7)."""
    edge = graph.edges.get(pipe_id)
    if edge is None or edge.pipe is None:
        w["fail_rows"].append({
            "id": pipe_id, "family": "Pipe",
            "reason": "edge or pipe element not found in graph."})
        w["fail_count"] += 1
        return

    # Determine phase before anything else. One test for new construction,
    # failing closed: an unreadable phase is existing and is never resized
    # (Colin, 2026-11-02).
    is_new = revit_helpers.is_new_construction(edge.pipe, doc)

    to_node = graph.nodes.get(edge.to_node_id)
    if to_node is not None and to_node.is_gas_fixture:
        # An existing phase appliance fed by new piping keeps its stub pipe
        # as drawn: it is audited, never resized (review n=26), so it is
        # not in new_sizes either (no fitting is sized from it).
        existing_appliance = to_node.element_id in (
            getattr(graph, "existing_appliances", None) or [])
        if is_new and not existing_appliance:
            w["new_sizes"][pipe_id] = nominal_size
            # New-construction stub: queued for STEP 9 (cap-first resize)
            w["skipped_stubs"].append({
                "pipe_id":          pipe_id,
                "fixture_name":     to_node.fixture_name or "UNNAMED",
                "demand_mbh":       edge.cumulative_load_mbh,
                "recommended_size": nominal_size,
            })
        else:
            # Existing stub, or the stub of an existing appliance: skip
            # entirely, run capacity audit only
            w["existing_skip_count"] += 1
            _audit_existing_pipe(pipe_id, edge, nominal_size, result,
                                 w["existing_overloaded"],
                                 w["existing_unchecked"])
        return

    if not is_new:
        # Existing distribution pipe: skip resize, audit capacity
        w["existing_skip_count"] += 1
        _audit_existing_pipe(pipe_id, edge, nominal_size, result,
                             w["existing_overloaded"], w["existing_unchecked"])
        return
    w["new_sizes"][pipe_id] = nominal_size

    nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(nominal_size)
    if nominal_inches is None:
        w["fail_rows"].append({
            "id": pipe_id, "family": "Pipe",
            "reason": "unrecognised nominal size '{}'.".format(nominal_size)})
        w["fail_count"] += 1
        return

    ok, _approach = _set_pipe_diameter(edge.pipe, nominal_inches)
    if ok:
        w["success_count"] += 1
    else:
        w["fail_count"] += 1
        w["fail_rows"].append({
            "id": pipe_id, "family": "Pipe",
            "reason": "all API approaches failed ({}\").".format(
                nominal_size)})


def _write_pipe_sizes(doc, graph, result):
    """STEP 7: write the new construction pipe sizes in one transaction.
    Returns the _new_pipe_writes() dict, with committed False and zero
    successes when Revit rolled the transaction back."""
    w = _new_pipe_writes()

    def action():
        for pipe_id, nominal_size in result["sizes"].items():
            _write_one_pipe(w, pipe_id, nominal_size, graph, result, doc)

    err = revit_helpers.run_in_transaction(
        doc, "RJA Tools - Size Gas Pipes", action)
    if err is None:
        w["committed"] = True
        return w
    revit_helpers._log_entry("ERROR", "_write_pipe_sizes", None,
                             "Pipe transaction rolled back: {}".format(err))
    attempted = w["success_count"] + w["fail_count"]
    w["success_count"] = 0
    w["fail_count"] = attempted
    w["error"] = err
    w["fail_rows"].append({
        "id": "-", "family": "All pipes",
        "reason": "transaction rolled back, no pipe size was written: "
                  "{}".format(err)})
    return w


def _write_fitting_sizes(doc, graph, new_sizes, transition_ids,
                         existing_sizes, pipes_committed):
    """STEP 8: resize the fittings in one transaction, only when the pipe
    sizes committed. Returns (fit counts, committed)."""
    if not pipes_committed:
        fit = _new_fit_counts()
        fit["failures"].append(
            "Fittings not resized: the pipe sizes were rolled back.")
        fit["failure_rows"].append({
            "id": "-", "family": "All fittings",
            "reason": "not run, the pipe sizes were rolled back."})
        return fit, False

    holder = {"fit": _new_fit_counts()}

    def action():
        holder["fit"] = _resize_fittings(graph, new_sizes, doc,
                                         transition_ids, existing_sizes)

    err = revit_helpers.run_in_transaction(
        doc, "RJA Tools - Resize Pipe Fittings", action)
    if err is None:
        return holder["fit"], True
    revit_helpers._log_entry("ERROR", "_write_fitting_sizes", None,
                             "Fitting transaction rolled back: {}".format(err))
    fit = _new_fit_counts()
    fit["failures"].append("Fitting transaction rolled back: {}".format(err))
    fit["failure_rows"].append({
        "id": "-", "family": "All fittings",
        "reason": "transaction rolled back: {}".format(err)})
    return fit, False


def _write_stub_sizes(doc, graph, skipped_stubs, fittings_committed):
    """STEP 9: size the fixture stub pipes in one transaction, after their
    caps were sized in STEP 8, so only when STEP 8 committed.

    Returns {"success", "fail", "rows", "rolled_back"}. A rolled back or
    skipped step counts every stub as failed (n=49)."""
    st = {"success": 0, "fail": 0, "rows": [], "rolled_back": False}
    if not skipped_stubs:
        return st
    if not fittings_committed:
        st["fail"] = len(skipped_stubs)
        st["rows"].append({
            "id": "-", "family": "Fixture stub pipes",
            "reason": "not run: the fittings and caps were not sized first "
                      "(their transaction did not commit)."})
        return st

    def action():
        for s in skipped_stubs:
            # One guard per stub: one bad stub never rolls back the rest.
            try:
                if _size_one_stub(s, graph):
                    st["success"] += 1
                else:
                    st["fail"] += 1
                    st["rows"].append({
                        "id": s["pipe_id"], "family": "Fixture stub pipe",
                        "reason": "size {}\" not written (see log)".format(
                            s["recommended_size"])})
            except Exception as e:
                st["fail"] += 1
                revit_helpers._log_entry(
                    "ERROR", "size_gas_stubs", s["pipe_id"],
                    "Stub resize raised: {}".format(str(e)))
                st["rows"].append({
                    "id": s["pipe_id"], "family": "Fixture stub pipe",
                    "reason": str(e)})

    err = revit_helpers.run_in_transaction(
        doc, "RJA Tools - Size Fixture Stub Pipes", action)
    if err is None:
        return st
    revit_helpers._log_entry("ERROR", "_write_stub_sizes", None,
                             "Stub transaction rolled back: {}".format(err))
    st["success"] = 0
    st["fail"] = len(skipped_stubs)
    st["rolled_back"] = True
    # the per-stub rows describe writes Revit undid: only the rollback row
    # stands
    st["rows"] = []
    st["rows"].append({
        "id": "-", "family": "Fixture stub pipes",
        "reason": "transaction rolled back: {}".format(err)})
    return st


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    revit_helpers.clear_log()
    _confirmed_approach[0] = None
    # revit_helpers keeps its own cache and stays loaded between runs, so it
    # has to be reset alongside the local one.
    revit_helpers.reset_pipe_diameter_approach()

    # ------------------------------------------------------------------
    # STEP 1 - Startup dialog: select pipe material and IFGC table
    # ------------------------------------------------------------------
    choice = ui_helpers.show_size_gas_dialog("Size Gas - Select IFGC Table")
    if choice is None:
        _say("Cancelled at table selection. No changes were made to the "
             "model.")
        return
    pipe_material           = choice["pipe_material"]
    selected_table_label    = choice["table_label"]
    heat_content_btu_per_cf = choice["heat_content"]
    propane_heat_content    = _propane_heat_content(choice)
    mid_stream_prv          = choice["mid_stream_prv"]

    selected_opt       = gas_tables.get_table_option_by_material_and_short_label(
        pipe_material, selected_table_label)
    table_id           = selected_opt["table_id"]
    inlet_pressure_psi = selected_opt["inlet_pressure_psi"]

    downstream_opt      = None
    downstream_table_id = None
    if mid_stream_prv:
        downstream_opt = gas_tables.get_table_option_by_material_and_short_label(
            pipe_material, choice["downstream_table_label"])
        downstream_table_id = downstream_opt["table_id"]

    # The inputs section prints once the pick is known (STEP 3), so the
    # table, material, gas, heat content and start show together.
    inputs = {
        "lib": revit_helpers.lib_folder(),
        "material": pipe_material,
        "table": selected_opt["label"].split("  [")[0],
        "table_id": table_id,
        "gas": selected_opt["gas"],
        # gas_console.heat_content_text prints what the table used: the
        # propane value on 402.4(37), none on an MBH table (n=24)
        "heat_content": heat_content_btu_per_cf,
        "propane_heat_content": propane_heat_content,
        "downstream_table": (downstream_opt["label"].split("  [")[0]
                             if downstream_opt is not None else None),
        "downstream_table_id": downstream_table_id,
    }

    # Scope, chosen in the dialog before the pick (Colin, 2026-11-20): None
    # is New Construction, else the existing gas piping entered there.
    connect_existing = choice.get("existing")

    # ------------------------------------------------------------------
    # STEP 2 - Pick the gas meter, or the first new pipe on an existing
    # building (Colin, 2026-10-13). A meter runs exactly as before. A pipe
    # starts the run at the connection to existing. Connect into Existing
    # takes the first new pipe only.
    # ------------------------------------------------------------------
    # the shared pick step, same words and window as One-Line (2026-11-19)
    selected_element = ui_helpers.pick_element(
        uidoc, doc, "Size Gas",
        ui_helpers.PICK_GAS_EXISTING if connect_existing
        else ui_helpers.PICK_GAS)
    if selected_element is None:
        pick_error = ui_helpers.last_pick_error()
        if pick_error:
            _say("The pick failed: {}. No changes were made to the "
                 "model.".format(pick_error), "error")
        else:
            _say("Selection cancelled. No changes were made to the "
                 "model.")
        return

    inputs["selected"] = "Element ID {}".format(
        revit_helpers.eid_int(selected_element.Id))

    # ------------------------------------------------------------------
    # STEP 3 - Validate the pick. For a pipe, find its existing side. The
    # existing conditions, assumed not modeled, came from the dialog.
    # ------------------------------------------------------------------
    # One check for Size Gas, One-Line and Diagnose (tie_in.check_gas_start).
    existing = None
    start = tie_in.check_gas_start(selected_element, doc,
                                   pipe_only=bool(connect_existing))
    if not start["ok"]:
        _alert(start["message"] + "\n\nNothing was changed.", "error")
        return
    upstream = start["upstream"]
    inputs["check"] = start["message"]
    if start["kind"] == "tie_in":
        # New Construction from a pipe: full new construction, new meter.
        existing = tie_in.existing_from_scope(connect_existing)
        inputs["existing"] = tie_in.existing_inputs_text(existing)

    _emit(gas_console.sizegas_inputs(inputs),
          gas_console.sizegas_inputs_md(inputs))

    # ------------------------------------------------------------------
    # STEP 4 - Traverse piping network
    # ------------------------------------------------------------------
    try:
        graph = pipe_graph.build_network(selected_element, doc, upstream)
    except Exception as e:
        _alert("Traversal failed:\n\n{}".format(str(e)), "error")
        _say("Traversal ERROR: {}".format(str(e)), "error")
        return

    trav = {"nodes": len(graph.nodes), "segments": len(graph.edges),
            "existing_connections": list(graph.existing_connections or [])}
    _emit(gas_console.traversal(trav), gas_console.traversal_md(trav))
    if getattr(graph, "existing_appliances", None):
        appliance_rows = report_generator.existing_appliance_rows(graph)
        _emit(gas_console.existing_appliances(appliance_rows),
              gas_console.existing_appliances_md(appliance_rows))

    fixture_nodes = pipe_graph.gas_fixtures(graph)
    if len(fixture_nodes) == 0:
        _alert(
            "No gas fixtures found (IS_GAS_FIXTURE = Yes).\n\n"
            "Common causes:\n"
            "  1. Pipe caps were inadvertently changed when resizing all "
            "piping in Revit -- verify cap families have not replaced fixture "
            "families at pipe terminations.\n"
            "  2. Fixture families are not physically connected to the piping.\n"
            "  3. IS_GAS_FIXTURE parameter is not set to Yes on the appliances.\n\n"
            "Run Diagnose to see which terminal elements were found.",
            "error")
        return

    zero_load = pipe_graph.zero_load_stop_message(graph)
    if zero_load is not None:
        _alert(zero_load, "error")
        _say("Stopped: total connected load is 0 MBH. No changes were made "
             "to the model.", "error")
        return

    if graph.longest_run is None:
        _alert(
            "Could not determine longest run. "
            "Run Diagnose first and resolve all errors.",
            "error")
        return

    # Existing building: the existing length joins the developed length that
    # picks the table row, before anything is sized or zoned.
    new_construction = tie_in.is_new_construction(existing)
    if tie_in.existing_for_sizing(existing) is not None:
        pipe_graph.apply_existing_conditions(
            graph, existing["size_nominal"], existing["length_ft"],
            existing["load_mbh"])

    # ------------------------------------------------------------------
    # STEP 4b - Pressure regulating valves
    # PRVs are auto-detected by family name. Only a MID-STREAM one (more than
    # PRV_MIDSTREAM_MIN_DOWNSTREAM_FT of pipe downstream of it) is a step
    # down; a regulator at its equipment is that equipment's own and is
    # ignored. Nothing changes unless "Mid Stream PRV" is checked.
    # ------------------------------------------------------------------
    min_ft         = shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT
    prv_nodes      = [graph.nodes[i] for i in graph.prv_ids]
    midstream_prvs = [graph.nodes[i] for i in graph.midstream_prv_ids]
    prv_warnings   = []

    # the same stop and words as One-Line (pipe_graph; the PRV keywords come
    # from shared_params.PRV_FAMILY_KEYWORDS)
    prv_stop = pipe_graph.midstream_prv_stop_message(graph, mid_stream_prv)
    if prv_stop is not None:
        _alert(prv_stop, "warn")
        return

    if mid_stream_prv and graph.nested_prv_ids:
        _alert(
            "Mid-stream PRV(s) {} sit downstream of another mid-stream "
            "PRV.\n\nOnly a single step down is supported.".format(
                ", ".join(str(i) for i in graph.nested_prv_ids)),
            "error")
        return

    if prv_nodes:
        prv_rows = [{
            "id": n.element_id, "family": n.family_name,
            "downstream_ft": n.downstream_pipe_ft or 0.0,
            "treated": ("mid-stream step down" if n.is_midstream_prv else
                        "equipment PRV, ignored (within {:g} ft)".format(
                            min_ft))} for n in prv_nodes]
        _emit(gas_console.prv_table(prv_rows),
              gas_console.prv_table_md(prv_rows))

        if midstream_prvs and not mid_stream_prv:
            prv_warnings.append(
                "{} mid-stream PRV(s) found but Mid Stream PRV is unchecked - "
                "the whole system is sized as one system on the meter's "
                "table, including the piping after the regulator(s).".format(
                    len(midstream_prvs)))

    # ------------------------------------------------------------------
    # STEP 5 - IFGC sizing calculation
    # ------------------------------------------------------------------
    try:
        result = sizing_engine.size_system(
            graph, pipe_material, inlet_pressure_psi, table_id,
            heat_content_btu_per_cf=heat_content_btu_per_cf,
            downstream_table_id=downstream_table_id,
            propane_heat_content_btu_per_cf=propane_heat_content)
    except ValueError as e:
        _alert("Sizing failed:\n\n{}".format(str(e)), "error")
        _say("Sizing ERROR: {}".format(str(e)), "error")
        return

    # Sizing results: the basis, then one table per system.
    total_load = pipe_graph.total_connected_mbh(graph)
    units = gas_tables.table_units(result["table_id"])
    _emit(gas_console.sizing_basis(result, total_load, len(fixture_nodes),
                                   units),
          gas_console.sizing_basis_md(result, total_load, len(fixture_nodes),
                                      units))
    rows = sizing_engine.sizing_rows(result, graph)
    _emit(gas_console.sizing_tables(rows, result),
          gas_console.sizing_tables_md(rows, result))
    for text, kind in _graph_notices(graph, result):
        _say(text, kind)

    # Existing pipe from the meter to the new connection: checked, never
    # sized. A shortfall is a warning the designer has to accept below.
    existing_check    = sizing_engine.check_existing_pipe(result, graph)
    existing_warnings = []
    if existing_check is not None:
        _emit(gas_console.existing_check(existing_check),
              gas_console.existing_check_md(existing_check))
        if not existing_check["adequate"]:
            if existing_check["capacity_cfh"] is None:
                existing_warnings.append(
                    "The entered existing pipe size {}\" is not in Table {} at "
                    "{:.0f} ft, so it could not be checked.".format(
                        existing_check["size_nominal"],
                        existing_check["table_id"],
                        float(existing_check["run_ft"])))
            else:
                # Capacity and demand are in the table's units: CFH, or MBH
                # for a propane table 402.4(25) to (36) (Colin, 2026-11-04).
                e_units = existing_check.get("units", "CFH")
                if e_units == "MBH":
                    load_text = "{:.0f} MBH".format(
                        float(existing_check["total_load_mbh"]))
                else:
                    load_text = "{:.0f} {} ({:.0f} MBH)".format(
                        float(existing_check["demand_cfh"]), e_units,
                        float(existing_check["total_load_mbh"]))
                existing_warnings.append(
                    "The existing {}\" pipe carries {:.0f} {} at {:.0f} ft "
                    "developed length, but existing plus new load is {}. "
                    "The existing pipe is NOT adequate and was not "
                    "resized.".format(
                        existing_check["size_nominal"],
                        float(existing_check["capacity_cfh"]), e_units,
                        float(existing_check["run_ft"]),
                        load_text))

    # ------------------------------------------------------------------
    # STEP 6 - Pre-sizing validation warning
    # ------------------------------------------------------------------
    # 0 MBH fixtures are listed from graph.data_errors only (the walk's
    # one record), not tested again here
    fixture_nodes  = pipe_graph.gas_fixtures(graph)
    unnamed        = [n for n in fixture_nodes
                      if n.fixture_name in ("", "UNNAMED", None)]

    # Table Intended Use load limits (2024 IFGC Table 402.4(13): 150 CFH),
    # on the new load plus any existing load entered (n=35/52).
    table_warnings = _table_limit_check(
        table_id, graph, heat_content_btu_per_cf, propane_heat_content)

    non_gas  = list(graph.non_gas_skipped)              # F1
    bare_end = pipe_graph.non_fixture_ends(graph)       # F2

    groups = gas_console.build_warning_groups(
        existing=existing_warnings, prv=prv_warnings,
        table_limit=table_warnings,
        unnamed=[(n.element_id, n.gas_load_mbh) for n in unnamed],
        non_gas=list(non_gas),
        bare_end=[(n.element_id, n.family_name) for n in bare_end],
        disconnected=list(graph.disconnected),
        n_fixtures=len(fixture_nodes),
        data_errors=list(getattr(graph, "data_errors", None) or []),
        tapped=list(getattr(graph, "tapped_pipes", None) or []))
    if groups:
        _emit(gas_console.presize_warnings(groups),
              gas_console.presize_warnings_md(groups))
        proceed = _alert(
            "{} pre-sizing warning(s) found (see the output window). "
            "Proceed with sizing anyway?".format(len(groups)),
            "warn", yes_no=True)
        if not proceed:
            _say("Sizing cancelled at the pre-sizing warnings. No changes "
                 "were made to the model.")
            return

    # ------------------------------------------------------------------
    # STEP 7 - Write sizes to Revit via Transaction
    # ------------------------------------------------------------------
    _emit(gas_console.write_start(
        "RBS_PIPE_NOMINAL_DIAMETER, RBS_PIPE_DIAMETER_PARAM, "
        "LookupParameter(Diameter)"),
        gas_console.write_start_md(
            "RBS_PIPE_NOMINAL_DIAMETER, RBS_PIPE_DIAMETER_PARAM, "
            "LookupParameter(Diameter)"))
    # Transitions are identified now, before the resize: Revit may delete one
    # while the pipes change size (Colin, 2026-11-19).
    transition_ids = _transition_ids(graph)

    w = _write_pipe_sizes(doc, graph, result)
    rolled_back = []
    if not w["committed"]:
        rolled_back.append(PIPE_STEP)
        _alert(
            "The pipe sizes could not be committed and Revit rolled them "
            "back. No size was written to the model.\n\n{}".format(
                w["error"]),
            "error")
        _say("Pipe transaction rolled back: {}".format(w["error"]), "error")

    # ------------------------------------------------------------------
    # Existing pipe audit results. Printing is guarded so a formatting
    # error can never stop Transactions 2 and 3 (Colin, 2026-11-02).
    # ------------------------------------------------------------------
    try:
        # new construction: print only if the walk still met existing phase
        # pipes, so a skip is never silent
        if not new_construction or w["existing_skip_count"]:
            _emit(gas_console.existing_audit(w["existing_skip_count"],
                                             w["existing_overloaded"],
                                             w["existing_unchecked"]),
                  gas_console.existing_audit_md(w["existing_skip_count"],
                                                w["existing_overloaded"],
                                                w["existing_unchecked"]))
    except Exception as e:
        _say("Existing pipe audit could not be printed: {}".format(str(e)),
             "warn")

    # ------------------------------------------------------------------
    # STEP 8 - Resize fittings (separate transaction), only when the pipe
    # sizes committed. New construction fittings only; one that joins an
    # existing pipe is never sized below it. Transitions are counted and
    # left to Revit; the meter and PRVs are never resized.
    # ------------------------------------------------------------------
    existing_sizes = {}
    if w["committed"]:
        existing_sizes = _existing_pipe_sizes(graph, w["new_sizes"], result,
                                              doc)
    fit, fit_committed = _write_fitting_sizes(
        doc, graph, w["new_sizes"], transition_ids, existing_sizes,
        w["committed"])
    if w["committed"] and not fit_committed:
        rolled_back.append(FITTING_STEP)

    # ------------------------------------------------------------------
    # STEP 9 - Resize fixture stub pipes (after cap Nominal Radius is set
    # in STEP 8, so Revit sees no size conflict and skips the family
    # replacement). Not run when STEP 8 did not commit.
    # ------------------------------------------------------------------
    st = _write_stub_sizes(doc, graph, w["skipped_stubs"], fit_committed)
    if st["rolled_back"]:
        rolled_back.append(STUB_STEP)

    try:
        failures = gas_console.failure_rows(
            w["fail_rows"], fit.get("failure_rows") or [], st["rows"])
        wr = {"pipes_sized": w["success_count"], "stubs_sized": st["success"],
              "fit": fit, "failures": failures}
        _emit(gas_console.write_results(wr), gas_console.write_results_md(wr))
    except Exception as e:
        _say("Write results could not be printed: {}. See the summary "
             "below.".format(str(e)), "warn")

    try:
        _print_summary(result, existing_check, w, st, fit, rolled_back)
    except Exception as e:
        _say("Summary could not be printed: {}. Check the write results "
             "above.".format(str(e)), "warn")


def _summary_info(result, existing_check, w, st, fit, rolled_back):
    """The info dict gas_console.summary reads. Counts come from the
    transactions as committed: a rolled back step has zero successes and is
    named in rolled_back (n=4/49)."""
    return {
        "success": w["success_count"], "stub_success": st["success"],
        "stub_fail": st["fail"], "existing_skip": w["existing_skip_count"],
        "existing_overloaded": len(w["existing_overloaded"]),
        "existing_unchecked": len(w["existing_unchecked"]),
        "fit": fit, "fail_count": w["fail_count"],
        # no approach stands when Revit rolled the pipe sizes back
        "approach": (_confirmed_approach[0] if w.get("committed")
                     else None),
        "table_id": result["table_id"],
        # gas_console.heat_content_text prints what the table used (n=24)
        "heat_content": result["heat_content_btu_per_cf"],
        "propane_heat_content": result.get(
            "propane_heat_content_btu_per_cf"),
        "existing_check": existing_check,
        "zones": result.get("zones") or [],
        "downstream": bool(result.get("downstream_table_id")),
        "longest_run": result["longest_run_ft"],
        "table_row": result["table_length_used_ft"],
        "rolled_back": list(rolled_back),
        "stub_rolled_back": bool(st.get("rolled_back")),
    }


def _print_summary(result, existing_check, w, st, fit, rolled_back):
    """The run summary table for the output window."""
    info = _summary_info(result, existing_check, w, st, fit, rolled_back)
    _emit(gas_console.summary(info), gas_console.summary_md(info))

main()
