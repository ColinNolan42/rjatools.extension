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
        output.print_md("Selection cancelled.")
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
    pipe_material = forms.CommandSwitchWindow.show(
        sorted(gas_tables.SUPPORTED_MATERIALS),
        message="Select pipe material:"
    )
    if not pipe_material:
        output.print_md("Cancelled.")
        return

    pressure_choice = forms.CommandSwitchWindow.show(
        _PRESSURE_OPTIONS,
        message="Select inlet pressure at meter:"
    )
    if not pressure_choice:
        output.print_md("Cancelled.")
        return

    inlet_pressure_psi = _PRESSURE_PSI[pressure_choice]

    output.print_md("**Material:**        {}".format(pipe_material))
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
    # STEP 6 - Write sizes to Revit via Transaction
    # ------------------------------------------------------------------
    output.print_md("---")
    output.print_md("**Writing sizes to Revit model...**")
    output.print_md(
        "Attempting three API approaches in order: "
        "RBS_PIPE_NOMINAL_DIAMETER -> "
        "RBS_PIPE_DIAMETER_PARAM -> "
        "LookupParameter(Diameter)")

    success_count = 0
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
    # STEP 7 - Summary
    # ------------------------------------------------------------------
    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| Pipes sized successfully | {} |".format(success_count))
    output.print_md("| Failures | {} |".format(fail_count))
    output.print_md("| API approach used | {} |".format(
        _confirmed_approach[0] or "None - all failed"))
    output.print_md("| IFGC table | {} |".format(result["table_id"]))
    output.print_md("| Longest run | {:.1f} ft |".format(
        result["longest_run_ft"]))
    output.print_md("| Table row used | {} ft |".format(
        result["table_length_used_ft"]))

    if fail_list:
        output.print_md("---")
        output.print_md("## :cross_mark: Failures")
        for f in fail_list:
            output.print_md("- {}".format(f))

    if fail_count == 0:
        output.print_md("---")
        output.print_md(
            ":white_check_mark: **All {} pipes sized and written to model.**".format(
                success_count))
    else:
        output.print_md("---")
        output.print_md(
            ":warning: {}/{} pipes failed. "
            "Check failures above.".format(fail_count, success_count + fail_count))


main()
