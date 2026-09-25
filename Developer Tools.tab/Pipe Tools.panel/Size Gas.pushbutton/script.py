# -*- coding: ascii -*-
# Size Gas.pushbutton/script.py
# Phase 2 - IFGC Longest Run Gas Pipe Sizing
# User picks gas meter -> startup dialog -> sizing -> writes sizes to Revit
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
# Fitting resize - set each connector radius to match its connected pipe
# ---------------------------------------------------------------------------

def _resize_fittings(graph, result_sizes, doc):
    """Resize parametric fittings by setting the 'Nominal Radius' instance
    parameter. Nominal Radius = nominal pipe diameter / 2 (in feet).

    For each fitting the target size is the highest-load connected pipe,
    which corresponds to the main run size at that fitting location.

    Returns:
        (resized_count, skipped_count, fail_list)
    """
    resized  = 0
    skipped  = 0
    failures = []

    for node in graph.nodes.values():
        if node.node_type not in ("tee", "fitting", "elbow", "fixture"):
            continue
        if node.element is None:
            skipped += 1
            continue

        # Find the highest-load connected pipe to determine target size
        target_size = None
        max_load    = -1
        for edge in graph.edges.values():
            if (edge.from_node_id == node.element_id or
                    edge.to_node_id == node.element_id):
                if (edge.element_id in result_sizes and
                        edge.cumulative_load_mbh > max_load):
                    max_load    = edge.cumulative_load_mbh
                    target_size = result_sizes[edge.element_id]

        if target_size is None:
            skipped += 1
            continue

        nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(target_size)
        if nominal_inches is None:
            skipped += 1
            continue

        # Nominal Radius = nominal diameter / 2, converted to feet
        target_radius_feet = (nominal_inches / 2.0) / 12.0

        ok = False

        # Approach 1: LookupParameter("Nominal Radius") - instance param on
        # Generic Standard elbows and tees
        try:
            param = node.element.LookupParameter("Nominal Radius")
            if param is not None and not param.IsReadOnly:
                param.Set(target_radius_feet)
                ok = True
        except Exception:
            pass

        # Approach 2: connector.Radius fallback
        if not ok:
            try:
                cm = node.element.ConnectorManager
                if cm is not None:
                    for connector in cm.Connectors:
                        try:
                            connector.Radius = target_radius_feet
                            ok = True
                        except Exception:
                            pass
            except Exception:
                pass

        if ok:
            resized += 1
        else:
            failures.append(
                "Fitting {} ({}): could not set size to {}".format(
                    node.element_id, node.family_name, target_size))
            skipped += 1

    return resized, skipped, failures


# ---------------------------------------------------------------------------
# Phase and diameter helpers
# ---------------------------------------------------------------------------

def _get_pipe_phase_name(pipe, doc):
    """Return the Phase Created name for a pipe, or None if unreadable."""
    try:
        p = pipe.get_Parameter(BuiltInParameter.PHASE_CREATED)
        if p is None:
            return None
        phase_elem = doc.GetElement(p.AsElementId())
        if phase_elem is None:
            return None
        return phase_elem.Name
    except Exception:
        return None


def _get_pipe_nominal_size(pipe):
    """Return the nominal size string for a pipe's current Revit diameter, or None.

    Steel nominal designations (1/2, 3/4 ... 12) take priority over EHD/K&L
    when multiple nominal sizes map to the same decimal inch value.
    """
    try:
        param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if param is None:
            param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_NOMINAL_DIAMETER)
        if param is None:
            return None
        dia_in = param.AsDouble() * 12.0  # Revit feet -> inches
        _STEEL = frozenset([
            "1/2","3/4","1","1-1/4","1-1/2","2","2-1/2",
            "3","4","5","6","8","10","12"])
        inv = {}
        for nom, inches in sizing_engine.NOMINAL_TO_INCHES.items():
            if nom not in _STEEL:
                inv[inches] = nom
        for nom, inches in sizing_engine.NOMINAL_TO_INCHES.items():
            if nom in _STEEL:
                inv[inches] = nom
        if not inv:
            return None
        closest = min(inv.keys(), key=lambda k: abs(k - dia_in))
        return inv[closest] if abs(closest - dia_in) < 0.1 else None
    except Exception:
        return None


def _audit_existing_pipe(pipe_id, edge, recommended_nom, result, overloaded_list):
    """Check whether an existing pipe's current size handles the new load.

    If current capacity < demand, appends a dict to overloaded_list so a
    warning is shown after the sizing transaction.

    Args:
        pipe_id:         int  pipe element ID
        edge:            graph edge (has .pipe, .cumulative_load_mbh)
        recommended_nom: str  the IFGC-computed size (already in result["sizes"])
        result:          dict returned by sizing_engine.size_system()
        overloaded_list: list  mutated in-place
    """
    demand = edge.cumulative_load_mbh
    if demand <= 0:
        return  # no load -> no concern

    current_nom = _get_pipe_nominal_size(edge.pipe)
    if current_nom is None:
        return  # can't read current size -> skip check

    demand_cfh = sizing_engine.mbh_to_cfh(
        demand, result.get("heat_content_btu_per_cf",
                            shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF))
    # The table and run this pipe was sized on (differs after a PRV).
    ctx = result.get("edge_context", {}).get(pipe_id, {})
    try:
        current_cap = gas_tables.get_capacity(
            ctx.get("table_id", result["table_id"]),
            ctx.get("longest_run_ft", result["longest_run_ft"]),
            current_nom)
    except ValueError:
        # Nominal not in this table (e.g. copper table selected for steel pipe)
        current_cap = 0.0

    if current_cap < demand_cfh:
        overloaded_list.append({
            "pipe_id":        pipe_id,
            "current_nom":    current_nom,
            "demand_mbh":     demand,
            "current_cap":    current_cap,
            "recommended_nom": recommended_nom,
        })


# (no startup dialog helpers - table options are loaded from gas_tables.py)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    output.print_md("# Size Gas - IFGC Longest Run Method")
    output.print_md("---")

    revit_helpers.clear_log()
    _confirmed_approach[0] = None
    # revit_helpers keeps its own cache and stays loaded between runs, so it
    # has to be reset alongside the local one.
    revit_helpers.reset_pipe_diameter_approach()

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
    # STEP 2 - Validate meter
    # ------------------------------------------------------------------
    validation = revit_helpers.validate_selected_element(selected_element)
    if not validation["is_valid"]:
        forms.alert(
            "Please select the gas meter element.\n\n{}".format(
                validation["reason"]),
            title="Size Gas - Invalid Selection"
        )
        return

    output.print_md(":white_check_mark: Meter validation passed.")

    # ------------------------------------------------------------------
    # STEP 3 - Startup dialog: select pipe material and IFGC table
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
    # STEP 4 - Traverse piping network
    # ------------------------------------------------------------------
    output.print_md("**Traversing network...**")
    try:
        graph = pipe_graph.build_network(selected_element, doc)
    except Exception as e:
        forms.alert(
            "Traversal failed:\n\n{}".format(str(e)),
            title="Size Gas - Traversal Error"
        )
        output.print_md(":cross_mark: Traversal ERROR: {}".format(str(e)))
        return

    output.print_md(":white_check_mark: {} nodes, {} pipe segments.".format(
        len(graph.nodes), len(graph.edges)))

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

    if no_load or unnamed or graph.disconnected or prv_warnings:
        output.print_md("---")
        output.print_md("## :warning: Pre-Sizing Warnings")

        for w in prv_warnings:
            output.print_md("**PRV:** {}".format(w))

        if no_load:
            output.print_md(
                "**{}/{} fixture(s) have 0 MBH load** - "
                "connected pipes will be sized at minimum (1/2\"):".format(
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

            # Determine phase before anything else
            phase_name = _get_pipe_phase_name(edge.pipe, doc)
            is_new     = (phase_name is None or phase_name == "New Construction")

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
                        pipe_id, edge, nominal_size, result, existing_overloaded)
                continue

            if not is_new:
                # Existing distribution pipe: skip resize, audit capacity
                existing_skip_count += 1
                _audit_existing_pipe(
                    pipe_id, edge, nominal_size, result, existing_overloaded)
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
    # Existing pipe audit results
    # ------------------------------------------------------------------
    if existing_skip_count > 0:
        output.print_md("---")
        if existing_overloaded:
            output.print_md(
                "## :warning: Existing Pipes: Load Exceeds Current Size")
            output.print_md(
                "{} existing pipe(s) were skipped. "
                "{} of them carry more load than their current size can handle "
                "at the system's longest run length. "
                "These pipes were **not** resized. Coordinate with the project "
                "team before upsizing existing piping.".format(
                    existing_skip_count, len(existing_overloaded)))
            output.print_md("")
            output.print_md(
                "| Pipe ID | Current Size | New Load | Current Capacity"
                " (CFH) | Recommended |")
            output.print_md("| --- | --- | --- | --- | --- |")
            for e in existing_overloaded:
                output.print_md(
                    "| {} | {}\" | {:.1f} MBH | {:.1f} | {}\" |".format(
                        e["pipe_id"],
                        e["current_nom"],
                        e["demand_mbh"],
                        e["current_cap"],
                        e["recommended_nom"]))
        else:
            output.print_md(
                ":white_check_mark: {} existing pipe(s) skipped — "
                "all adequately sized for the new load.".format(
                    existing_skip_count))

    # ------------------------------------------------------------------
    # STEP 8 - Resize fittings (separate transaction)
    # ------------------------------------------------------------------
    output.print_md("**Resizing fittings (Nominal Radius parameter)...**")
    fit_resized  = 0
    fit_skipped  = 0
    fit_failures = []

    t2 = Transaction(doc, "RJA Tools - Resize Pipe Fittings")
    t2.Start()
    try:
        fit_resized, fit_skipped, fit_failures = _resize_fittings(
            graph, result["sizes"], doc)
        doc.Regenerate()
        t2.Commit()
        output.print_md(
            ":white_check_mark: Fittings: {} resized, {} skipped.".format(
                fit_resized, fit_skipped))
        if fit_failures:
            for f in fit_failures:
                output.print_md(":warning: {}".format(f))
    except Exception as e:
        t2.RollBack()
        output.print_md(
            ":warning: Fitting resize transaction failed: {}".format(str(e)))

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
        try:
            for s in skipped_stubs:
                edge = graph.edges.get(s["pipe_id"])
                if edge is None or edge.pipe is None:
                    stub_fail += 1
                    continue
                nominal_inches = sizing_engine.NOMINAL_TO_INCHES.get(
                    s["recommended_size"])
                if nominal_inches is None:
                    stub_fail += 1
                    continue
                ok, _approach = _set_pipe_diameter(edge.pipe, nominal_inches)
                if ok:
                    stub_success += 1
                else:
                    stub_fail += 1
            doc.Regenerate()
            t3.Commit()
            output.print_md(
                ":white_check_mark: Stub pipes: {} sized.".format(stub_success))
            if stub_fail:
                output.print_md(
                    ":warning: {} stub pipe(s) failed.".format(stub_fail))
        except Exception as e:
            t3.RollBack()
            output.print_md(
                ":warning: Stub pipe transaction failed: {}".format(str(e)))

    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| Distribution pipes sized | {} |".format(success_count))
    output.print_md("| Fixture stub pipes sized | {} |".format(stub_success))
    output.print_md("| Existing pipes skipped | {} |".format(existing_skip_count))
    output.print_md("| Existing pipes undersized | {} |".format(
        len(existing_overloaded)))
    output.print_md("| Fittings + caps resized | {} |".format(fit_resized))
    output.print_md("| Fitting failures | {} |".format(fit_skipped))
    output.print_md("| Pipe failures | {} |".format(fail_count))
    output.print_md("| API approach used | {} |".format(
        _confirmed_approach[0] or "None - all failed"))
    output.print_md("| IFGC table | {} |".format(result["table_id"]))
    output.print_md("| Heat Content of Gas | {:.0f} BTU/CF |".format(
        result["heat_content_btu_per_cf"]))
    if result.get("downstream_table_id"):
        # One row per system: the meter side, then each PRV.
        for z in result["zones"]:
            output.print_md("| {} | Table {}, run {:.1f} ft (row {} ft) |".format(
                z["label"], z["table_id"], z["run_ft"], z["row_ft"]))
    else:
        output.print_md("| Longest run | {:.1f} ft |".format(
            result["longest_run_ft"]))
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
                success_count, stub_success, fit_resized))
    else:
        output.print_md("---")
        output.print_md(
            ":warning: {}/{} pipes failed. "
            "Check failures above.".format(fail_count, success_count + fail_count))


main()
