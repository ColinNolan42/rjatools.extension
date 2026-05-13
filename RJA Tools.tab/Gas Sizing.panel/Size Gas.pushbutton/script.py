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


# ---------------------------------------------------------------------------
# Pipe diameter write-back
# Three API approaches tried in order. First success wins for all pipes.
# ---------------------------------------------------------------------------

# Populated on first successful write - avoids re-trying failed approaches
_confirmed_approach = [None]


def _set_pipe_diameter(pipe, nominal_inches):
    """Try multiple Revit API approaches to set a pipe's nominal diameter.

    Args:
        pipe:           Revit Pipe element
        nominal_inches: float  nominal diameter in inches

    Returns:
        (success: bool, approach_name: str)
    """
    nominal_feet = nominal_inches / 12.0
    pipe_id = pipe.Id.IntegerValue

    # If a working approach was already confirmed this run, use it directly
    if _confirmed_approach[0] is not None:
        return _apply_approach(_confirmed_approach[0], pipe, nominal_feet)

    # --- Approach 1: RBS_PIPE_NOMINAL_DIAMETER ---
    ok, name = _apply_approach("RBS_PIPE_NOMINAL_DIAMETER", pipe, nominal_feet)
    if ok:
        _confirmed_approach[0] = name
        output.print_md(
            ":white_check_mark: API approach confirmed: **{}**".format(name))
        return True, name

    # --- Approach 2: RBS_PIPE_DIAMETER_PARAM ---
    ok, name = _apply_approach("RBS_PIPE_DIAMETER_PARAM", pipe, nominal_feet)
    if ok:
        _confirmed_approach[0] = name
        output.print_md(
            ":white_check_mark: API approach confirmed: **{}**".format(name))
        return True, name

    # --- Approach 3: LookupParameter Diameter ---
    ok, name = _apply_approach("LookupParameter", pipe, nominal_feet)
    if ok:
        _confirmed_approach[0] = name
        output.print_md(
            ":white_check_mark: API approach confirmed: **{}**".format(name))
        return True, name

    output.print_md(
        ":cross_mark: Pipe {}: all three API approaches failed.".format(pipe_id))
    return False, "FAILED"


def _apply_approach(approach_name, pipe, nominal_feet):
    """Apply one specific approach. Returns (success, approach_name)."""
    try:
        if approach_name == "RBS_PIPE_NOMINAL_DIAMETER":
            param = pipe.get_Parameter(
                BuiltInParameter.RBS_PIPE_NOMINAL_DIAMETER)

        elif approach_name == "RBS_PIPE_DIAMETER_PARAM":
            param = pipe.get_Parameter(
                BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)

        elif approach_name == "LookupParameter":
            param = pipe.LookupParameter("Diameter")

        else:
            return False, approach_name

        if param is None:
            return False, approach_name

        if param.IsReadOnly:
            return False, approach_name

        param.Set(nominal_feet)
        return True, approach_name

    except Exception:
        return False, approach_name


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
        if node.node_type not in ("tee", "fitting", "elbow"):
            continue
        if node.is_gas_fixture:
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
# Startup dialog helpers
# ---------------------------------------------------------------------------

_PRESSURE_OPTIONS = [
    "< 2 PSI  (Table 402.4(2), 0.5 in. w.c. drop)",
    "2 PSI    (Table 402.4(5), 1.0 PSI drop)",
    "3 PSI    (Table 402.4(6), 2.0 PSI drop)",
    "5 PSI    (Table 402.4(7), 3.5 PSI drop)",
]

_PRESSURE_PSI = {
    _PRESSURE_OPTIONS[0]: 1.5,
    _PRESSURE_OPTIONS[1]: 2.0,
    _PRESSURE_OPTIONS[2]: 3.0,
    _PRESSURE_OPTIONS[3]: 5.0,
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    output.print_md("# Size Gas - IFGC Longest Run Method")
    output.print_md("---")

    revit_helpers.clear_log()
    _confirmed_approach[0] = None

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
        selected_element.Id.IntegerValue))

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
    # STEP 3 - Startup dialog: pipe material then inlet pressure
    # ------------------------------------------------------------------
    # If only one material is supported, skip the dialog and use it directly.
    # Show a dialog only when multiple materials are available.
    materials = sorted(gas_tables.SUPPORTED_MATERIALS)
    if len(materials) == 1:
        pipe_material = materials[0]
        output.print_md("**Material:** {} (only supported material)".format(
            pipe_material))
    else:
        pipe_material = forms.CommandSwitchWindow.show(
            materials,
            message="Select pipe material:"
        )
        if not pipe_material:
            output.print_md(
                "Cancelled at pipe material dialog. "
                "No changes were made to the model.")
            return

    pressure_choice = forms.SelectFromList.show(
        _PRESSURE_OPTIONS,
        title="Size Gas - Select Inlet Pressure",
        multiselect=False
    )
    if not pressure_choice:
        output.print_md(
            "Cancelled at inlet pressure dialog. "
            "No changes were made to the model.")
        return

    inlet_pressure_psi = _PRESSURE_PSI[pressure_choice]
    output.print_md("**Inlet pressure:**  {} PSI".format(inlet_pressure_psi))

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

    if graph.longest_run is None:
        forms.alert(
            "Could not determine longest run. "
            "Run Diagnose first and resolve all errors.",
            title="Size Gas - Sizing Error"
        )
        return

    # ------------------------------------------------------------------
    # STEP 5 - IFGC sizing calculation
    # ------------------------------------------------------------------
    output.print_md("**Running IFGC sizing...**")
    try:
        result = sizing_engine.size_system(graph, pipe_material, inlet_pressure_psi)
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

    if no_load or unnamed or graph.disconnected:
        output.print_md("---")
        output.print_md("## :warning: Pre-Sizing Warnings")

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
    output.print_md(
        "Note: pipes directly connected to fixture families are skipped "
        "to preserve custom fixture parameters.")

    success_count = 0
    skip_count    = 0
    fail_count    = 0
    fail_list     = []

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

            # Skip pipes directly connected to gas fixtures - resizing these
            # causes Revit to replace the fixture cap family, losing custom
            # parameters (GAS_LOAD_MBH, FIXTURE_NAME, IS_GAS_FIXTURE).
            to_node = graph.nodes.get(edge.to_node_id)
            if to_node is not None and to_node.is_gas_fixture:
                skip_count += 1
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
    # STEP 7 - Resize fittings (separate transaction)
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
    # STEP 8 - Summary
    # ------------------------------------------------------------------
    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| Pipes sized and written | {} |".format(success_count))
    output.print_md("| Fixture stub pipes skipped | {} |".format(skip_count))
    output.print_md("| Fittings resized | {} |".format(fit_resized))
    output.print_md("| Pipe failures | {} |".format(fail_count))
    output.print_md("| API approach used | {} |".format(
        _confirmed_approach[0] or "None - all failed"))
    output.print_md("| IFGC table | {} |".format(result["table_id"]))
    output.print_md("| Longest run | {:.1f} ft |".format(
        result["longest_run_ft"]))
    output.print_md("| Table row used | {} ft |".format(
        result["table_length_used_ft"]))

    if skip_count:
        output.print_md(
            "_Note: {} fixture stub pipe(s) were not resized to preserve "
            "custom fixture family parameters._".format(skip_count))

    if fail_list:
        output.print_md("---")
        output.print_md("## :cross_mark: Failures")
        for f in fail_list:
            output.print_md("- {}".format(f))

    if fail_count == 0:
        output.print_md("---")
        output.print_md(
            ":white_check_mark: **{} pipes sized and written. "
            "{} fixture stubs preserved.**".format(success_count, skip_count))
    else:
        output.print_md("---")
        output.print_md(
            ":warning: {}/{} pipes failed. "
            "Check failures above.".format(fail_count, success_count + fail_count))


main()
