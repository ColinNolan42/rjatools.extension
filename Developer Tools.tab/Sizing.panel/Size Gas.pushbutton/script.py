# -*- coding: ascii -*-
# Size Gas.pushbutton/script.py
# Phase 2 - IFGC Longest Run Gas Pipe Sizing
# Startup dialog -> user picks the gas meter, or the first new pipe on an
# existing building -> sizing -> writes sizes to Revit
#
# IronPython 2.7 / PyRevit

import os
import sys

from pyrevit import script, forms
from Autodesk.Revit.DB import BuiltInParameter, ElementId, Transaction
from Autodesk.Revit.UI.Selection import ObjectType

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib_dir = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

import shared_params
import revit_helpers
import pipe_graph
import gas_tables
import sizing_engine
import tie_in
import ui_helpers


# ---------------------------------------------------------------------------
# Pipe diameter write-back
# Three API approaches tried in order. First success wins for all pipes.
# ---------------------------------------------------------------------------

# Populated on first successful write - avoids re-trying failed approaches
_confirmed_approach = [None]


def _set_pipe_diameter(pipe, nominal_inches):
    """Set a pipe's nominal diameter and report the approach to the window.

    The three-approach write logic now lives in revit_helpers.set_pipe_diameter
    so Size Gas and Size Water share one implementation. This wrapper only adds
    the pyRevit output messages, which behave exactly as before.

    Args:
        pipe:           Revit Pipe element
        nominal_inches: float  nominal diameter in inches

    Returns:
        (success: bool, approach_name: str)
    """
    already_confirmed = _confirmed_approach[0] is not None

    ok, name = revit_helpers.set_pipe_diameter(pipe, nominal_inches)

    if ok and not already_confirmed:
        _confirmed_approach[0] = name
        output.print_md(
            ":white_check_mark: API approach confirmed: **{}**".format(name))
    elif not ok:
        output.print_md(
            ":cross_mark: Pipe {}: all three API approaches failed.".format(
                revit_helpers.eid_int(pipe.Id)))

    return ok, name


# ---------------------------------------------------------------------------
# Fitting resize - one size per fitting through revit_helpers.set_fitting_size
# ---------------------------------------------------------------------------

def _resize_fittings(graph, new_sizes, doc):
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
        parameter) and failures (list of str, real write failures).
    """
    counts = {"resized": 0, "no_sized_pipe": 0, "existing": 0,
              "no_element": 0, "appliance": 0, "replaced": 0,
              "failures": []}
    done_ids = set()   # replacement fittings already sized this run

    # One guard per fitting: an exception on one fitting is one failure
    # line, never the rollback of every fitting (G7C, 2026-11-04).
    for node in graph.nodes.values():
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        try:
            _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids)
        except Exception as e:
            counts["failures"].append(
                "Fitting {} ({}): {}".format(
                    node.element_id, node.family_name, str(e)))
            revit_helpers._log_entry(
                "ERROR", "_resize_fittings", node.element_id,
                "Resize raised: {}".format(str(e)))

    return counts


def _resize_one_fitting(node, graph, new_sizes, doc, counts, done_ids):
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
        element, _d = revit_helpers.find_replacement_fitting(
            pipes, node.location_xyz, set(graph.nodes.keys()) | done_ids)
        if element is None:
            counts["failures"].append(
                "Fitting {} ({}): deleted by Revit when its pipes were "
                "resized, and no replacement fitting was found on its "
                "pipes.".format(node.element_id, node.family_name))
            return
        replaced = True
        revit_helpers._log_entry(
            "INFO", "_resize_fittings", node.element_id,
            "Deleted by Revit; replacement {} sized instead.".format(
                revit_helpers.eid_int(element.Id)))

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
        counts["failures"].append(
            "Fitting {} ({}): unrecognised nominal size '{}'.".format(
                node.element_id, node.family_name, target_size))
        return

    ok, approach = revit_helpers.set_fitting_size(element, nominal_inches)
    if ok:
        counts["resized"] += 1
        if replaced:
            counts["replaced"] += 1
            done_ids.add(revit_helpers.eid_int(element.Id))
    else:
        counts["failures"].append(
            "Fitting {} ({}): could not set size to {}\" ({}).".format(
                node.element_id, node.family_name, target_size, approach))


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
    output.print_md("# Size Gas - IFGC Longest Run Method")
    output.print_md("Tools loaded from `{}`".format(revit_helpers.lib_folder()))
    output.print_md("---")

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
        output.print_md(
            "Cancelled at table selection. No changes were made to the model.")
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

    output.print_md("**Table:**     {} ({})".format(
        table_id, selected_opt["label"].split("  [")[0]))
    if downstream_opt is not None:
        output.print_md("**Downstream of PRV:** {} ({})".format(
            downstream_table_id, downstream_opt["label"].split("  [")[0]))
    output.print_md("**Material:**  {}".format(pipe_material))
    output.print_md("**Gas:**       {}".format(selected_opt["gas"]))
    output.print_md("**Heat Content of Gas:** {:.0f} BTU/CF".format(heat_content_btu_per_cf))

    # ------------------------------------------------------------------
    # STEP 2 - Pick the gas meter, or the first new pipe on an existing
    # building (Colin, 2026-10-13). A meter runs exactly as before. A pipe
    # starts the run at the connection to existing.
    # ------------------------------------------------------------------
    forms.alert(
        "Select meter or new gas piping connection.",
        title="Size Gas")
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Select meter or new gas piping connection"
        )
        selected_element = doc.GetElement(ref.ElementId)
    except Exception:
        output.print_md("Selection cancelled. No element was picked.")
        return

    if selected_element is None:
        forms.alert(
            "Could not retrieve selected element.",
            title="Size Gas - Selection Error"
        )
        return

    output.print_md("**Selected:** Element ID {}".format(
        revit_helpers.eid_int(selected_element.Id)))

    # ------------------------------------------------------------------
    # STEP 3 - Validate the pick. For a pipe, find its existing side and ask
    # for the existing conditions, which are assumed not modeled.
    # ------------------------------------------------------------------
    upstream = None
    existing = None
    if revit_helpers.is_pipe(selected_element):
        gas_check = pipe_graph.check_gas_pipe(selected_element)
        if not gas_check["ok"]:
            forms.alert(
                "This is not a gas pipe: {}.\n\nPick the gas meter, or the "
                "first new GAS pipe at the connection to existing.".format(
                    gas_check["reason"]),
                title="Size Gas - Invalid Selection"
            )
            return
        upstream = tie_in.resolve_new_pipe_upstream(selected_element, doc)
        if upstream["status"] != "ok":
            forms.alert(
                "Could not tell which end of the picked pipe connects to the "
                "existing gas piping: {}.\n\nPick the first new gas pipe, the "
                "one with one end at the existing connection (open, capped, "
                "or joined to existing phase piping) and the other end on "
                "new construction.".format(upstream["reason"]),
                title="Size Gas - Invalid Selection"
            )
            return
        output.print_md(
            ":white_check_mark: New gas piping connection. Upstream end "
            "found because {}.".format(upstream["reason"]))

        picked_id = revit_helpers.eid_int(selected_element.Id)
        existing = ui_helpers.show_existing_gas_dialog(
            "Size Gas - Existing Conditions",
            gas_tables.list_pipe_sizes(table_id),
            tie_in.recall_existing(picked_id))
        if existing is None:
            output.print_md(
                "Cancelled at existing conditions. No changes were made to "
                "the model.")
            return
        tie_in.remember_existing(picked_id, existing)
        output.print_md(
            "**Existing (entered):** {}\" pipe, {:.1f} ft from the meter to "
            "the new connection, {:.1f} MBH existing load.".format(
                existing["size_nominal"], float(existing["length_ft"]),
                float(existing["load_mbh"])))
    else:
        validation = revit_helpers.validate_selected_element(selected_element)
        if not validation["is_valid"]:
            forms.alert(
                "Please select the gas meter element, or the first new gas "
                "pipe at the connection to existing.\n\n{}".format(
                    validation["reason"]),
                title="Size Gas - Invalid Selection"
            )
            return

        output.print_md(":white_check_mark: Meter validation passed.")

    # ------------------------------------------------------------------
    # STEP 4 - Traverse piping network
    # ------------------------------------------------------------------
    output.print_md("**Traversing network...**")
    try:
        graph = pipe_graph.build_network(selected_element, doc, upstream)
    except Exception as e:
        forms.alert(
            "Traversal failed:\n\n{}".format(str(e)),
            title="Size Gas - Traversal Error"
        )
        output.print_md(":cross_mark: Traversal ERROR: {}".format(str(e)))
        return

    output.print_md(":white_check_mark: {} nodes, {} pipe segments.".format(
        len(graph.nodes), len(graph.edges)))

    if graph.existing_connections:
        output.print_md(
            ":information_source: {} connection(s) to Existing phase piping "
            "were not followed ({}). The existing piping is assumed adequate "
            "and is not sized. A new pipe that ends at one carries no load in "
            "this run.".format(
                len(graph.existing_connections),
                ", ".join(str(i) for i in graph.existing_connections[:10])))

    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
    if len(fixture_nodes) == 0:
        forms.alert(
            "No gas fixtures found (IS_GAS_FIXTURE = Yes).\n\n"
            "Common causes:\n"
            "  1. Pipe caps were inadvertently changed when resizing all "
            "piping in Revit -- verify cap families have not replaced fixture "
            "families at pipe terminations.\n"
            "  2. Fixture families are not physically connected to the piping.\n"
            "  3. IS_GAS_FIXTURE parameter is not set to Yes on the appliances.\n\n"
            "Run Diagnose to see which terminal elements were found.",
            title="Size Gas - No Fixtures"
        )
        return

    if graph.longest_run is None:
        forms.alert(
            "Could not determine longest run. "
            "Run Diagnose first and resolve all errors.",
            title="Size Gas - Sizing Error"
        )
        return

    # Existing building: the existing length joins the developed length that
    # picks the table row, before anything is sized or zoned.
    if existing is not None:
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
        forms.alert(
            "Mid Stream PRV is checked, but there is no mid-stream "
            "regulator in this system.\n\n{}\n\nUncheck Mid Stream PRV "
            "to size the system as one.".format(why),
            title="Size Gas - No Mid-Stream PRV"
        )
        return

    if mid_stream_prv and graph.nested_prv_ids:
        forms.alert(
            "Mid-stream PRV(s) {} sit downstream of another mid-stream "
            "PRV.\n\nOnly a single step down is supported.".format(
                ", ".join(str(i) for i in graph.nested_prv_ids)),
            title="Size Gas - Nested PRVs"
        )
        return

    if prv_nodes:
        output.print_md("---")
        output.print_md("## Pressure Regulating Valves")
        output.print_md("| Element ID | Family | Pipe downstream | Treated as |")
        output.print_md("| --- | --- | --- | --- |")
        for n in prv_nodes:
            output.print_md("| {} | {} | {:.1f} ft | {} |".format(
                n.element_id, n.family_name, n.downstream_pipe_ft or 0.0,
                "mid-stream step down" if n.is_midstream_prv
                else "equipment PRV, ignored (within {:g} ft)".format(min_ft)))

        if midstream_prvs and not mid_stream_prv:
            prv_warnings.append(
                "{} mid-stream PRV(s) found but Mid Stream PRV is unchecked - "
                "the whole system is sized as one system on the meter's "
                "table, including the piping after the regulator(s).".format(
                    len(midstream_prvs)))

    # ------------------------------------------------------------------
    # STEP 5 - IFGC sizing calculation
    # ------------------------------------------------------------------
    output.print_md("**Running IFGC sizing...**")
    try:
        result = sizing_engine.size_system(
            graph, pipe_material, inlet_pressure_psi, table_id,
            heat_content_btu_per_cf=heat_content_btu_per_cf,
            downstream_table_id=downstream_table_id)
    except ValueError as e:
        forms.alert(
            "Sizing failed:\n\n{}".format(str(e)),
            title="Size Gas - Sizing Error"
        )
        output.print_md(":cross_mark: Sizing ERROR: {}".format(str(e)))
        return

    # Print sizing summary to output window
    output.print_md("---")
    diag = sizing_engine.format_sizing_output(result, graph)

    # Existing pipe from the meter to the new connection: checked, never
    # sized. A shortfall is a warning the designer has to accept below.
    existing_check    = sizing_engine.check_existing_pipe(result, graph)
    existing_warnings = []
    if existing_check is not None:
        diag += "\n\n" + "\n".join(
            sizing_engine.format_existing_check(existing_check))
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
                units = existing_check.get("units", "CFH")
                if units == "MBH":
                    load_text = "{:.0f} MBH".format(
                        float(existing_check["total_load_mbh"]))
                else:
                    load_text = "{:.0f} {} ({:.0f} MBH)".format(
                        float(existing_check["demand_cfh"]), units,
                        float(existing_check["total_load_mbh"]))
                existing_warnings.append(
                    "The existing {}\" pipe carries {:.0f} {} at {:.0f} ft "
                    "developed length, but existing plus new load is {}. "
                    "The existing pipe is NOT adequate and was not "
                    "resized.".format(
                        existing_check["size_nominal"],
                        float(existing_check["capacity_cfh"]), units,
                        float(existing_check["run_ft"]),
                        load_text))
    output.print_html(
        "<pre style='font-family:monospace;font-size:12px;'>{}</pre>".format(
            diag.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")))


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

    if (no_load or unnamed or graph.disconnected or prv_warnings
            or existing_warnings or table_warnings):
        output.print_md("---")
        output.print_md("## :warning: Pre-Sizing Warnings")

        for w in existing_warnings:
            output.print_md("**Existing gas pipe:** {}".format(w))

        for w in prv_warnings:
            output.print_md("**PRV:** {}".format(w))

        for w in table_warnings:
            output.print_md("**IFGC table limit:** {}".format(w))

        if no_load:
            output.print_md(
                "**{}/{} fixture(s) have 0 MBH load** - "
                "connected pipes will be sized at minimum (3/4\"):".format(
                    len(no_load), len(fixture_nodes)))
            for n in no_load:
                output.print_md("  - {} (ID {})".format(
                    n.fixture_name or "UNNAMED", n.element_id))

        if unnamed:
            output.print_md(
                "**{}/{} fixture(s) have no name** "
                "(will appear as UNNAMED on one-line diagram):".format(
                    len(unnamed), len(fixture_nodes)))
            for n in unnamed:
                output.print_md("  - ID {}  {:.1f} MBH".format(
                    n.element_id, n.gas_load_mbh))

        if graph.disconnected:
            output.print_md(
                "**{} disconnected element(s)** - "
                "not included in sizing:".format(len(graph.disconnected)))
            for d in graph.disconnected:
                output.print_md("  - Element ID {}".format(d))

        proceed = forms.alert(
            "Warnings found (see output window). Proceed with sizing anyway?",
            title="Size Gas - Warnings",
            yes=True,
            no=True
        )
        if not proceed:
            output.print_md(
                "Sizing cancelled at pre-sizing warnings dialog. "
                "No changes were made to the model.")
            return

    # ------------------------------------------------------------------
    # STEP 7 - Write sizes to Revit via Transaction
    # ------------------------------------------------------------------
    output.print_md("---")
    output.print_md("**Writing sizes to Revit model...**")
    output.print_md(
        "Attempting three API approaches in order: "
        "RBS_PIPE_NOMINAL_DIAMETER -> "
        "RBS_PIPE_DIAMETER_PARAM -> "
        "LookupParameter(Diameter)")
    success_count       = 0
    fail_count          = 0
    fail_list           = []
    skip_count          = 0
    skipped_stubs       = []
    existing_skip_count = 0
    existing_overloaded = []  # pipes that are Existing but undersized for new load
    existing_unchecked  = []  # existing pipes whose size is not in the table
    new_sizes           = {}  # new construction pipes only, for the fittings

    t = Transaction(doc, "RJA Tools - Size Gas Pipes")
    t.Start()

    try:
        for pipe_id, nominal_size in result["sizes"].items():
            edge = graph.edges.get(pipe_id)
            if edge is None or edge.pipe is None:
                fail_list.append(
                    "Pipe {}: edge or pipe element not found in graph.".format(
                        pipe_id))
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
                fail_list.append(
                    "Pipe {}: unrecognised nominal size '{}'.".format(
                        pipe_id, nominal_size))
                fail_count += 1
                continue

            ok, approach = _set_pipe_diameter(edge.pipe, nominal_inches)

            if ok:
                success_count += 1
            else:
                fail_count += 1
                fail_list.append(
                    "Pipe {}: all API approaches failed "
                    "({}\").".format(pipe_id, nominal_size))

        # Regenerate inside the transaction so Revit propagates pipe size
        # changes to connected fittings (elbows, tees) before committing.
        doc.Regenerate()
        t.Commit()

    except Exception as e:
        t.RollBack()
        forms.alert(
            "Transaction failed - no sizes were written to the model.\n\n"
            "{}".format(str(e)),
            title="Size Gas - Transaction Error"
        )
        output.print_md(":cross_mark: Transaction ERROR: {}".format(str(e)))
        return

    # ------------------------------------------------------------------
    # Existing pipe audit results. Printing is guarded so a formatting
    # error can never stop Transactions 2 and 3 (Colin, 2026-11-02).
    # ------------------------------------------------------------------
    try:
        _print_existing_audit(existing_skip_count, existing_overloaded,
                              existing_unchecked)
    except Exception as e:
        output.print_md(
            ":warning: Existing pipe audit could not be printed: {}".format(
                str(e)))

    # ------------------------------------------------------------------
    # STEP 8 - Resize fittings (separate transaction). New construction
    # fittings only, sized from new construction pipes only.
    # ------------------------------------------------------------------
    output.print_md("**Resizing fittings (Nominal Radius parameter)...**")
    fit = {"resized": 0, "no_sized_pipe": 0, "existing": 0,
           "no_element": 0, "appliance": 0, "replaced": 0, "failures": []}
    fit_committed = False

    t2 = Transaction(doc, "RJA Tools - Resize Pipe Fittings")
    t2.Start()
    try:
        fit = _resize_fittings(graph, new_sizes, doc)
        doc.Regenerate()
        t2.Commit()
        fit_committed = True
    except Exception as e:
        t2.RollBack()
        output.print_md(
            ":warning: Fitting resize transaction failed: {}".format(str(e)))
        fit = {"resized": 0, "no_sized_pipe": 0, "existing": 0,
               "no_element": 0, "appliance": 0, "replaced": 0,
               "failures": ["Fitting transaction rolled back: {}".format(
                   str(e))]}

    if fit_committed:
        try:
            output.print_md(
                ":white_check_mark: Fittings: {} resized, {} existing left "
                "alone, {} with no new pipe sized this run, {} appliance(s) "
                "with no size parameter, {} failed.".format(
                    fit["resized"], fit["existing"], fit["no_sized_pipe"],
                    fit["appliance"], len(fit["failures"])))
            if fit.get("replaced"):
                output.print_md(
                    ":information_source: {} fitting(s) were replaced by "
                    "Revit when their pipes were resized; the replacement "
                    "was sized.".format(fit["replaced"]))
            for f in fit["failures"]:
                output.print_md(":warning: {}".format(f))
        except Exception as e:
            output.print_md(
                ":warning: Fitting results could not be printed: {}".format(
                    str(e)))

    # ------------------------------------------------------------------
    # STEP 9 - Resize fixture stub pipes (after cap Nominal Radius is set)
    # Cap Nominal Radius was updated in the fitting transaction above.
    # Revit may now see no size conflict and skip the family replacement.
    # ------------------------------------------------------------------
    stub_success = 0
    stub_fail    = 0

    if skipped_stubs:
        output.print_md("**Sizing fixture stub pipes (cap already resized)...**")
        t3 = Transaction(doc, "RJA Tools - Size Fixture Stub Pipes")
        t3.Start()
        stub_committed = False
        try:
            for s in skipped_stubs:
                # One guard per stub: one bad stub never rolls back the rest.
                try:
                    if _size_one_stub(s, graph):
                        stub_success += 1
                    else:
                        stub_fail += 1
                except Exception as e:
                    stub_fail += 1
                    revit_helpers._log_entry(
                        "ERROR", "size_gas_stubs", s["pipe_id"],
                        "Stub resize raised: {}".format(str(e)))
                    output.print_md(":warning: Stub pipe {}: {}".format(
                        s["pipe_id"], str(e)))
            doc.Regenerate()
            t3.Commit()
            stub_committed = True
        except Exception as e:
            t3.RollBack()
            stub_success = 0
            output.print_md(
                ":warning: Stub pipe transaction failed: {}".format(str(e)))
        if stub_committed:
            output.print_md(
                ":white_check_mark: Stub pipes: {} sized.".format(stub_success))
            if stub_fail:
                output.print_md(
                    ":warning: {} stub pipe(s) failed.".format(stub_fail))

    try:
        _print_summary(result, existing_check, success_count, stub_success,
                       stub_fail, existing_skip_count, existing_overloaded,
                       existing_unchecked, fit, fail_count, fail_list)
    except Exception as e:
        output.print_md(
            ":warning: Summary could not be printed: {}. The sizes above "
            "were written.".format(str(e)))


def _print_existing_audit(existing_skip_count, existing_overloaded,
                          existing_unchecked):
    """The existing pipe audit table for the output window."""
    if existing_skip_count <= 0:
        return
    output.print_md("---")
    if existing_overloaded:
        output.print_md(
            "## :warning: Existing Pipes: Load Exceeds Current Size")
        output.print_md(
            "{} existing pipe(s) were skipped. "
            "{} of them carry more load than their current size can handle "
            "at the system's longest run length, or the table publishes no "
            "capacity for their size at that length. "
            "These pipes were **not** resized. Coordinate with the project "
            "team before upsizing existing piping.".format(
                existing_skip_count, len(existing_overloaded)))
        output.print_md("")
        output.print_md(
            "| Pipe ID | Current Size | New Load | Current Capacity"
            " | Recommended |")
        output.print_md("| --- | --- | --- | --- | --- |")
        for e in existing_overloaded:
            cap = e["current_cap"]
            cap_text = ("none published" if cap is None
                        else "{:.1f} {}".format(float(cap),
                                                e.get("units", "CFH")))
            output.print_md(
                "| {} | {}\" | {:.1f} MBH | {} | {}\" |".format(
                    e["pipe_id"],
                    e["current_nom"],
                    float(e["demand_mbh"]),
                    cap_text,
                    e["recommended_nom"]))
    else:
        output.print_md(
            ":white_check_mark: {} existing pipe(s) skipped, "
            "all adequately sized for the new load.".format(
                existing_skip_count))
    if existing_unchecked:
        output.print_md(
            ":warning: {} existing pipe(s) could not be checked: the drawn "
            "diameter matches no size in the table ({}).".format(
                len(existing_unchecked),
                ", ".join("{} [Table {}]".format(u["pipe_id"], u["table_id"])
                          for u in existing_unchecked[:10])))


def _print_summary(result, existing_check, success_count, stub_success,
                   stub_fail, existing_skip_count, existing_overloaded,
                   existing_unchecked, fit, fail_count, fail_list):
    """The run summary table for the output window."""
    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| Distribution pipes sized | {} |".format(success_count))
    output.print_md("| Fixture stub pipes sized | {} |".format(stub_success))
    if stub_fail:
        output.print_md("| Fixture stub pipe failures | {} |".format(stub_fail))
    output.print_md("| Existing pipes skipped | {} |".format(existing_skip_count))
    output.print_md("| Existing pipes undersized | {} |".format(
        len(existing_overloaded)))
    if existing_unchecked:
        output.print_md("| Existing pipes not checked | {} |".format(
            len(existing_unchecked)))
    output.print_md("| Fittings + caps resized | {} |".format(fit["resized"]))
    if fit.get("replaced"):
        output.print_md("| Fittings replaced by Revit, replacement sized | "
                        "{} |".format(fit["replaced"]))
    output.print_md("| Existing fittings left alone | {} |".format(
        fit["existing"]))
    output.print_md("| Fittings with no new pipe sized | {} |".format(
        fit["no_sized_pipe"]))
    output.print_md("| Fitting failures | {} |".format(len(fit["failures"])))
    output.print_md("| Pipe failures | {} |".format(fail_count))
    output.print_md("| API approach used | {} |".format(
        _confirmed_approach[0] or "None - all failed"))
    output.print_md("| IFGC table | {} |".format(result["table_id"]))
    if existing_check is not None:
        output.print_md(
            "| Existing pipe (entered) | {}\", {:.1f} ft, {:.1f} MBH existing "
            "load, {} |".format(
                existing_check["size_nominal"],
                float(existing_check["length_ft"]),
                float(existing_check["existing_load_mbh"]),
                "adequate" if existing_check["adequate"] else "NOT ADEQUATE"))
    output.print_md("| Heat Content of Gas | {:.0f} BTU/CF |".format(
        float(result["heat_content_btu_per_cf"])))
    if result.get("downstream_table_id"):
        # One row per system: the meter side, then each PRV.
        for z in result["zones"]:
            output.print_md("| {} | Table {}, run {:.1f} ft (row {} ft) |".format(
                z["label"], z["table_id"], float(z["run_ft"]), z["row_ft"]))
    else:
        output.print_md("| Longest run | {:.1f} ft |".format(
            float(result["longest_run_ft"])))
        output.print_md("| Table row used | {} ft |".format(
            result["table_length_used_ft"]))

    if fail_list:
        output.print_md("---")
        output.print_md("## :cross_mark: Pipe Failures")
        for f in fail_list:
            output.print_md("- {}".format(f))

    if fail_count == 0:
        output.print_md("---")
        output.print_md(
            ":white_check_mark: **{} distribution pipes + {} stub pipes + "
            "{} fittings/caps sized.**".format(
                success_count, stub_success, fit["resized"]))
    else:
        output.print_md("---")
        output.print_md(
            ":warning: {}/{} pipes failed. "
            "Check failures above.".format(fail_count, success_count + fail_count))


main()
