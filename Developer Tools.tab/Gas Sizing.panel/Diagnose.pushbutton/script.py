# -*- coding: ascii -*-
# Diagnose.pushbutton/script.py
# Entry point for the Diagnose button.
# User pre-selects the gas meter in Revit, then clicks this button.
# Zero dialogs. Zero prompts. Pre-selection is the only input.
#
# IronPython 2.7 / PyRevit

import os
import sys

from pyrevit import script, forms
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
import report_generator


def main():
    output.print_md("# Gas Piping Diagnostic")
    output.print_md("---")

    revit_helpers.clear_log()

    # ------------------------------------------------------------------
    # STEP 1 - Prompt user to pick the gas meter
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
            "Could not retrieve the selected element. Please try again.",
            title="Diagnose - Selection Error"
        )
        return

    output.print_md("**Selected:** Element ID {}".format(
        selected_element.Id.IntegerValue))

    # ------------------------------------------------------------------
    # STEP 2 -Validate selection
    # ------------------------------------------------------------------
    validation = revit_helpers.validate_selected_element(selected_element)

    if not validation["is_valid"]:
        forms.alert(
            "Please select the gas meter element.\n\n{}".format(
                validation["reason"]),
            title="Diagnose -Invalid Selection"
        )
        return

    output.print_md(":white_check_mark: Meter validation passed.")

    # ------------------------------------------------------------------
    # STEP 3 -Traverse piping network
    # ------------------------------------------------------------------
    try:
        graph = pipe_graph.build_network(selected_element, doc)
    except Exception as e:
        forms.alert(
            "Traversal failed:\n\n{}".format(str(e)),
            title="Diagnose -Traversal Error"
        )
        output.print_md(":cross_mark: Traversal ERROR: {}".format(str(e)))
        return

    output.print_md(":white_check_mark: Traversal complete -{} nodes, {} pipe segments.".format(
        len(graph.nodes), len(graph.edges)))

    # ------------------------------------------------------------------
    # STEP 4 -Print diagnostic output (copy/paste into conversation)
    # ------------------------------------------------------------------
    output.print_md("---")
    output.print_md("## Diagnostic Output -Copy and paste below this line")
    diagnostic_text = report_generator.format_diagnostic_output(graph, selected_element)
    output.print_html("<pre style='font-family:monospace;font-size:12px;'>{}</pre>".format(
        diagnostic_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")))

    # ------------------------------------------------------------------
    # STEP 5 -Generate one-line diagram data (used by One-Line button)
    # ------------------------------------------------------------------
    one_line = report_generator.generate_one_line_data(graph)

    # ------------------------------------------------------------------
    # STEP 6 -Print summary
    # ------------------------------------------------------------------
    fixtures = [n for n in graph.nodes.values() if n.is_gas_fixture]
    total_load = sum(n.gas_load_mbh for n in fixtures)
    longest_ft = graph.longest_run["total_length_feet"] if graph.longest_run else 0.0
    farthest   = graph.longest_run["farthest_fixture_name"] if graph.longest_run else "-"
    lr = graph.longest_run or {}

    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| Fixtures found | {} |".format(len(fixtures)))
    output.print_md("| Total load | {:.1f} MBH |".format(total_load))
    output.print_md("| Pipe segments | {} |".format(len(graph.edges)))
    output.print_md("| Longest run | {:.1f} ft |".format(longest_ft))
    output.print_md("| Elbows on longest run | {} x 5 ft = {:.0f} ft |".format(
        lr.get("elbow_count", 0), lr.get("elbow_equiv_length_feet", 0.0)))
    output.print_md("| Farthest fixture | {} |".format(farthest))

    if fixtures:
        output.print_md("---")
        output.print_md("## Fixtures")
        output.print_md("| Name | Load (MBH) | Element ID |")
        output.print_md("| --- | --- | --- |")
        for n in fixtures:
            output.print_md("| {} | {:.1f} | {} |".format(
                n.fixture_name, n.gas_load_mbh, n.element_id))

    # ------------------------------------------------------------------
    # STEP 7 -Errors and warnings
    # ------------------------------------------------------------------
    errors   = []
    warnings = []

    if len(fixtures) == 0:
        errors.append("No gas fixtures found.")
    if total_load <= 0:
        errors.append("Total load is 0 MBH.")
    if graph.longest_run is None:
        errors.append("Longest run could not be determined.")
    if graph.disconnected:
        warnings.append("{} disconnected element(s) found.".format(
            len(graph.disconnected)))

    if errors:
        output.print_md("---")
        output.print_md("## :cross_mark: Errors -Fix Before Sizing")
        for e in errors:
            output.print_md("- {}".format(e))

    if warnings:
        output.print_md("---")
        output.print_md("## :warning: Warnings")
        for w in warnings:
            output.print_md("- {}".format(w))

    if not errors:
        output.print_md("---")
        output.print_md(":white_check_mark: **System is ready for Phase 2 sizing.**")


main()
