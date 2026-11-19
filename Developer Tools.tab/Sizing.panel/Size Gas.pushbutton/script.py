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
from Autodesk.Revit.DB import BuiltInParameter, ElementId, Transaction

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
            "transitions": 0, "transitions_removed": 0,
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


def _resize_fittings(graph, new_sizes, doc, transition_ids=None):
    """Resize the new construction fittings to the pipes sized this run.

    For each fitting the target size is the highest-load connected pipe in
    new_sizes, which corresponds to the main run size at that fitting
    location. new_sizes holds new construction pipes only, so an existing
    pipe never sets a fitting's size. An existing phase fitting is never
    resized: existing stays the size it is, and the existing pipe audit
    reports a shortfall (Colin, 2026-11-02).

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
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        try:
            _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                                transition_ids)
        except Exception as e:
            _fit_fail(counts, node, node.family_name, str(e))
            revit_helpers._log_entry(
                "ERROR", "_resize_fittings", node.element_id,
                "Resize raised: {}".format(str(e)))

    return counts


def _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids,
                        transition_ids=()):
    """Resize one fitting node. See _resize_fittings for the rules.

    The node's element is re-read first. Revit may have deleted it when the
    pipes were resized in the pipe transaction and placed another fitting;
    that replacement is found on the connected pipes and sized instead.
    """
    if node.element is None:
        counts["no_element"] += 1
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
                and revit_helpers.pipes_joined(pipes)):
            counts["transitions_removed"] += 1
            revit_helpers._log_entry(
                "INFO", "_resize_fittings", node.element_id,
                "Transition removed by Revit (pipes now the same size).")
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

    # Find the highest-load connected new pipe to determine target size
    target_size = None
    max_load    = -1
    for edge in graph.edges.values():
        if (edge.from_node_id == node.element_id or
                edge.to_node_id == node.element_id):
            if (edge.element_id in new_sizes and
                    edge.cumulative_load_mbh > max_load):
                max_load    = edge.cumulative_load_mbh
                target_size = new_sizes[edge.element_id]

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

    ok, approach = revit_helpers.set_fitting_size(element, nominal_inches)
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

    # Demand in the table's units: CFH via heat content, or the MBH load
    # itself on a propane table 402.4(25) to (36) (Colin, 2026-11-04).
    demand_cfh, units = sizing_engine.demand_in_table_units(
        demand, table_id, result.get("heat_content_btu_per_cf",
                                     shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))
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
        "heat_content": heat_content_btu_per_cf,
        "downstream_table": (downstream_opt["label"].split("  [")[0]
                             if downstream_opt is not None else None),
        "downstream_table_id": downstream_table_id,
    }

    # ------------------------------------------------------------------
    # STEP 2 - Pick the gas meter, or the first new pipe on an existing
    # building (Colin, 2026-10-13). A meter runs exactly as before. A pipe
    # starts the run at the connection to existing.
    # ------------------------------------------------------------------
    # the shared pick step, same words and window as One-Line (2026-11-19)
    selected_element = ui_helpers.pick_element(
        uidoc, doc, "Size Gas", ui_helpers.PICK_GAS)
    if selected_element is None:
        _say("Selection cancelled. No changes were made to the model.")
        return

    inputs["selected"] = "Element ID {}".format(
        revit_helpers.eid_int(selected_element.Id))

    # ------------------------------------------------------------------
    # STEP 3 - Validate the pick. For a pipe, find its existing side and ask
    # for the existing conditions, which are assumed not modeled.
    # ------------------------------------------------------------------
    upstream = None
    existing = None
    if revit_helpers.is_pipe(selected_element):
        gas_check = pipe_graph.check_gas_pipe(selected_element)
        if not gas_check["ok"]:
            _alert(
                "This is not a gas pipe: {}.\n\nPick the gas meter, or the "
                "first new GAS pipe at the connection to existing.".format(
                    gas_check["reason"]),
                "error")
            return
        upstream = tie_in.resolve_new_pipe_upstream(selected_element, doc)
        if upstream["status"] != "ok":
            _alert(
                "Could not tell which end of the picked pipe connects to the "
                "existing gas piping: {}.\n\nPick the first new gas pipe, the "
                "one with one end at the existing connection (open, capped, "
                "or joined to existing phase piping) and the other end on "
                "new construction.".format(upstream["reason"]),
                "error")
            return
        inputs["check"] = ("New gas piping connection. Upstream end found "
                           "because {}.".format(upstream["reason"]))

        picked_id = revit_helpers.eid_int(selected_element.Id)
        existing = ui_helpers.show_existing_gas_dialog(
            "Size Gas - Existing Conditions",
            gas_tables.list_pipe_sizes(table_id),
            tie_in.recall_existing(picked_id))
        if existing is None:
            _say("Cancelled at existing conditions. No changes were made to "
                 "the model.")
            return
        if not tie_in.is_new_construction(existing):
            tie_in.remember_existing(picked_id, existing)
        inputs["existing"] = tie_in.existing_inputs_text(existing)
    else:
        validation = revit_helpers.validate_selected_element(selected_element)
        if not validation["is_valid"]:
            _alert(
                "Please select the gas meter element, or the first new gas "
                "pipe at the connection to existing.\n\n{}".format(
                    validation["reason"]),
                "error")
            return

        inputs["check"] = "Meter validation passed."

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

    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
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

    if mid_stream_prv and not midstream_prvs:
        if prv_nodes:
            why = ("{} regulator(s) were found, but each has {:g} ft of pipe "
                   "or less downstream, so they are treated as equipment "
                   "PRVs and ignored.".format(len(prv_nodes), min_ft))
        else:
            why = ("No pressure regulating valve was found. A PRV is "
                   "recognised by 'PRV', 'regulator' or 'regulating' in its "
                   "family name; check it is connected to the piping.")
        _alert(
            "Mid Stream PRV is checked, but there is no mid-stream "
            "regulator in this system.\n\n{}\n\nUncheck Mid Stream PRV "
            "to size the system as one.".format(why),
            "warn")
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
            downstream_table_id=downstream_table_id)
    except ValueError as e:
        _alert("Sizing failed:\n\n{}".format(str(e)), "error")
        _say("Sizing ERROR: {}".format(str(e)), "error")
        return

    # Sizing results: the basis, then one table per system.
    total_load = sum(n.gas_load_mbh for n in fixture_nodes)
    units = gas_tables.table_units(result["table_id"])
    _emit(gas_console.sizing_basis(result, total_load, len(fixture_nodes),
                                   units),
          gas_console.sizing_basis_md(result, total_load, len(fixture_nodes),
                                      units))
    rows = sizing_engine.sizing_rows(result, graph)
    _emit(gas_console.sizing_tables(rows, result),
          gas_console.sizing_tables_md(rows, result))

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
    fixture_nodes  = [n for n in graph.nodes.values() if n.is_gas_fixture]
    no_load        = [n for n in fixture_nodes if n.gas_load_mbh <= 0]
    unnamed        = [n for n in fixture_nodes
                      if n.fixture_name in ("", "UNNAMED", None)]

    # Table Intended Use load limits (2024 IFGC Table 402.4(13): 150 CFH).
    # The meter carries the whole system, so the total load is checked.
    total_mbh = sum(n.gas_load_mbh for n in fixture_nodes)
    table_warnings = sizing_engine.table_limit_warnings(
        table_id,
        sizing_engine.mbh_to_cfh(total_mbh, heat_content_btu_per_cf),
        total_mbh)

    non_gas  = list(graph.non_gas_skipped)              # F1
    bare_end = pipe_graph.non_fixture_ends(graph)       # F2

    groups = gas_console.build_warning_groups(
        existing=existing_warnings, prv=prv_warnings,
        table_limit=table_warnings,
        no_load=[(n.fixture_name, n.element_id) for n in no_load],
        unnamed=[(n.element_id, n.gas_load_mbh) for n in unnamed],
        non_gas=list(non_gas),
        bare_end=[(n.element_id, n.family_name) for n in bare_end],
        disconnected=list(graph.disconnected),
        n_fixtures=len(fixture_nodes))
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
    success_count       = 0
    fail_count          = 0
    fail_rows           = []   # pipe failures, {"id", "family", "reason"}
    skip_count          = 0
    skipped_stubs       = []
    existing_skip_count = 0
    existing_overloaded = []  # pipes that are Existing but undersized for new load
    existing_unchecked  = []  # existing pipes whose size is not in the table
    new_sizes           = {}  # new construction pipes only, for the fittings

    # Transitions are identified now, before the resize: Revit may delete one
    # while the pipes change size (Colin, 2026-11-19).
    transition_ids = _transition_ids(graph)

    t = Transaction(doc, "RJA Tools - Size Gas Pipes")
    t.Start()

    try:
        for pipe_id, nominal_size in result["sizes"].items():
            edge = graph.edges.get(pipe_id)
            if edge is None or edge.pipe is None:
                fail_rows.append({
                    "id": pipe_id, "family": "Pipe",
                    "reason": "edge or pipe element not found in graph."})
                fail_count += 1
                continue

            # Determine phase before anything else. One test for new
            # construction, failing closed: an unreadable phase is existing
            # and is never resized (Colin, 2026-11-02).
            is_new = revit_helpers.is_new_construction(edge.pipe, doc)
            if is_new:
                new_sizes[pipe_id] = nominal_size

            to_node = graph.nodes.get(edge.to_node_id)
            is_stub = (to_node is not None and to_node.is_gas_fixture)

            if is_stub:
                if is_new:
                    # New-construction stub: queue for STEP 9 (cap-first resize)
                    skip_count += 1
                    skipped_stubs.append({
                        "pipe_id":          pipe_id,
                        "fixture_name":     to_node.fixture_name or "UNNAMED",
                        "demand_mbh":       edge.cumulative_load_mbh,
                        "recommended_size": nominal_size,
                    })
                else:
                    # Existing stub: skip entirely, run capacity audit only
                    existing_skip_count += 1
                    _audit_existing_pipe(
                        pipe_id, edge, nominal_size, result,
                        existing_overloaded, existing_unchecked)
                continue

            if not is_new:
                # Existing distribution pipe: skip resize, audit capacity
                existing_skip_count += 1
                _audit_existing_pipe(
                    pipe_id, edge, nominal_size, result,
                    existing_overloaded, existing_unchecked)
                continue

            nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(nominal_size)
            if nominal_inches is None:
                fail_rows.append({
                    "id": pipe_id, "family": "Pipe",
                    "reason": "unrecognised nominal size '{}'.".format(
                        nominal_size)})
                fail_count += 1
                continue

            ok, approach = _set_pipe_diameter(edge.pipe, nominal_inches)

            if ok:
                success_count += 1
            else:
                fail_count += 1
                fail_rows.append({
                    "id": pipe_id, "family": "Pipe",
                    "reason": "all API approaches failed ({}\").".format(
                        nominal_size)})

        # Regenerate inside the transaction so Revit propagates pipe size
        # changes to connected fittings (elbows, tees) before committing.
        doc.Regenerate()
        t.Commit()

    except Exception as e:
        t.RollBack()
        _alert(
            "Transaction failed - no sizes were written to the model.\n\n"
            "{}".format(str(e)),
            "error")
        _say("Transaction ERROR: {}".format(str(e)), "error")
        return

    # ------------------------------------------------------------------
    # Existing pipe audit results. Printing is guarded so a formatting
    # error can never stop Transactions 2 and 3 (Colin, 2026-11-02).
    # ------------------------------------------------------------------
    try:
        # new construction: print only if the walk still met existing phase
        # pipes, so a skip is never silent
        if not new_construction or existing_skip_count:
            _emit(gas_console.existing_audit(existing_skip_count,
                                             existing_overloaded,
                                             existing_unchecked),
                  gas_console.existing_audit_md(existing_skip_count,
                                                existing_overloaded,
                                                existing_unchecked))
    except Exception as e:
        _say("Existing pipe audit could not be printed: {}".format(str(e)),
             "warn")

    # ------------------------------------------------------------------
    # STEP 8 - Resize fittings (separate transaction). New construction
    # fittings only, sized from new construction pipes only. Transitions
    # are counted and left to Revit.
    # ------------------------------------------------------------------
    fit = _new_fit_counts()

    t2 = Transaction(doc, "RJA Tools - Resize Pipe Fittings")
    t2.Start()
    try:
        fit = _resize_fittings(graph, new_sizes, doc, transition_ids)
        doc.Regenerate()
        t2.Commit()
    except Exception as e:
        t2.RollBack()
        fit = _new_fit_counts()
        fit["failures"].append(
            "Fitting transaction rolled back: {}".format(str(e)))
        fit["failure_rows"].append({
            "id": "-", "family": "All fittings",
            "reason": "transaction rolled back: {}".format(str(e))})

    # ------------------------------------------------------------------
    # STEP 9 - Resize fixture stub pipes (after cap Nominal Radius is set)
    # Cap Nominal Radius was updated in the fitting transaction above.
    # Revit may now see no size conflict and skip the family replacement.
    # ------------------------------------------------------------------
    stub_success = 0
    stub_fail    = 0
    stub_rows    = []

    if skipped_stubs:
        t3 = Transaction(doc, "RJA Tools - Size Fixture Stub Pipes")
        t3.Start()
        try:
            for s in skipped_stubs:
                # One guard per stub: one bad stub never rolls back the rest.
                try:
                    if _size_one_stub(s, graph):
                        stub_success += 1
                    else:
                        stub_fail += 1
                        stub_rows.append({
                            "id": s["pipe_id"], "family": "Fixture stub pipe",
                            "reason": "size {}\" not written (see log)".format(
                                s["recommended_size"])})
                except Exception as e:
                    stub_fail += 1
                    revit_helpers._log_entry(
                        "ERROR", "size_gas_stubs", s["pipe_id"],
                        "Stub resize raised: {}".format(str(e)))
                    stub_rows.append({
                        "id": s["pipe_id"], "family": "Fixture stub pipe",
                        "reason": str(e)})
            doc.Regenerate()
            t3.Commit()
        except Exception as e:
            t3.RollBack()
            stub_success = 0
            stub_rows.append({
                "id": "-", "family": "Fixture stub pipes",
                "reason": "transaction rolled back: {}".format(str(e))})

    try:
        failures = gas_console.failure_rows(
            fail_rows, fit.get("failure_rows") or [], stub_rows)
        wr = {"pipes_sized": success_count, "stubs_sized": stub_success,
              "fit": fit, "failures": failures}
        _emit(gas_console.write_results(wr), gas_console.write_results_md(wr))
    except Exception as e:
        _say("Write results could not be printed: {}. The sizes were "
             "written.".format(str(e)), "warn")

    try:
        _print_summary(result, existing_check, success_count, stub_success,
                       stub_fail, existing_skip_count, existing_overloaded,
                       existing_unchecked, fit, fail_count)
    except Exception as e:
        _say("Summary could not be printed: {}. The sizes above were "
             "written.".format(str(e)), "warn")


def _print_summary(result, existing_check, success_count, stub_success,
                   stub_fail, existing_skip_count, existing_overloaded,
                   existing_unchecked, fit, fail_count):
    """The run summary table for the output window."""
    info = {
        "success": success_count, "stub_success": stub_success,
        "stub_fail": stub_fail, "existing_skip": existing_skip_count,
        "existing_overloaded": len(existing_overloaded),
        "existing_unchecked": len(existing_unchecked),
        "fit": fit, "fail_count": fail_count,
        "approach": _confirmed_approach[0],
        "table_id": result["table_id"],
        "heat_content": result["heat_content_btu_per_cf"],
        "existing_check": existing_check,
        "zones": result.get("zones") or [],
        "downstream": bool(result.get("downstream_table_id")),
        "longest_run": result["longest_run_ft"],
        "table_row": result["table_length_used_ft"],
    }
    _emit(gas_console.summary(info), gas_console.summary_md(info))


main()
