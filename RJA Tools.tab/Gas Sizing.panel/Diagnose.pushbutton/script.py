# Diagnose.pushbutton/script.py
# Entry point for the Diagnose button in Revit.
# User selects the gas meter -> clicks Diagnose -> report is saved and summary shown.
# No dialogs. No prompts. The selected element is the only input.
#
# IronPython 2.7 / PyRevit

import os
import sys
import json
import datetime

# --- PyRevit imports ---
from pyrevit import script
from pyrevit import forms

# --- Revit API ---
from Autodesk.Revit.DB import Document

# --- PyRevit document/UI handles ---
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# --- PyRevit output window ---
output = script.get_output()

# --- Add lib/ folder to path so shared modules are importable ---
# PyRevit should handle this automatically if lib/ is in the extension root.
# This block is a fallback in case the path is not added automatically.
_script_dir = os.path.dirname(__file__)
_lib_dir = os.path.normpath(os.path.join(_script_dir, '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

# --- Shared modules ---
import shared_params
import revit_helpers
import pipe_graph
import report_generator


def main():
    """Main entry point. Called when the user clicks Diagnose."""

    output.print_md("# Gas Piping Diagnostic")
    output.print_md("---")

    # -------------------------------------------------------------------------
    # STEP 1  -  Get selected element
    # -------------------------------------------------------------------------
    revit_helpers.clear_log()

    selection = uidoc.Selection.GetElementIds()

    if not selection or len(selection) == 0:
        forms.alert(
            "No element selected.\n\nPlease select the gas meter element, then click Diagnose.",
            title="Diagnose  -  No Selection"
        )
        return

    if len(selection) > 1:
        forms.alert(
            "Multiple elements selected.\n\nPlease select only the gas meter element.",
            title="Diagnose  -  Multiple Elements Selected"
        )
        return

    element_id = list(selection)[0]
    selected_element = doc.GetElement(element_id)

    output.print_md("**Selected element:** ID {}".format(element_id.IntegerValue))

    # -------------------------------------------------------------------------
    # STEP 2  -  Validate the selected element
    # -------------------------------------------------------------------------
    output.print_md("**Validating selection...**")

    validation = revit_helpers.validate_selected_element(selected_element)

    if not validation["is_valid"]:
        forms.alert(
            "Invalid selection: {}\n\nPlease select the gas meter element.".format(
                validation["reason"]),
            title="Diagnose  -  Invalid Selection"
        )
        output.print_md(
            ":cross_mark: Validation FAILED: {}".format(validation["reason"]))
        return

    output.print_md(":white_check_mark: Meter validation passed.")
    output.print_md(
        "Connector directions found: {}".format(validation["connector_summary"]))

    # -------------------------------------------------------------------------
    # STEP 3  -  Traverse the piping network
    # -------------------------------------------------------------------------
    output.print_md("**Traversing gas piping network...**")

    try:
        graph = pipe_graph.build_network(selected_element, doc)
    except Exception as e:
        forms.alert(
            "Traversal failed with error:\n\n{}".format(str(e)),
            title="Diagnose  -  Traversal Error"
        )
        output.print_md(":cross_mark: Traversal ERROR: {}".format(str(e)))
        return

    output.print_md(":white_check_mark: Traversal complete.")
    output.print_md(
        "Found: **{}** nodes, **{}** pipe segments.".format(
            len(graph.nodes), len(graph.edges)))

    # -------------------------------------------------------------------------
    # STEP 4  -  Get environment info
    # -------------------------------------------------------------------------
    try:
        revit_version = "Revit {}".format(
            doc.Application.VersionNumber)
    except Exception:
        revit_version = "Unknown"

    try:
        from pyrevit import versionmgr
        pyrevit_version = str(versionmgr.get_pyrevit_version())
    except Exception:
        pyrevit_version = "Unknown"

    # -------------------------------------------------------------------------
    # STEP 5  -  Generate diagnostic report
    # -------------------------------------------------------------------------
    output.print_md("**Generating diagnostic report...**")

    try:
        report = report_generator.generate_diagnostic_report(
            graph=graph,
            origin_element=selected_element,
            revit_version=revit_version,
            pyrevit_version=pyrevit_version
        )
    except Exception as e:
        forms.alert(
            "Report generation failed:\n\n{}".format(str(e)),
            title="Diagnose  -  Report Error"
        )
        output.print_md(":cross_mark: Report generation ERROR: {}".format(str(e)))
        return

    # -------------------------------------------------------------------------
    # STEP 6  -  Save report to desktop
    # -------------------------------------------------------------------------
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "gas_diagnostic_{}.json".format(timestamp)
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        report_path = os.path.join(desktop, filename)

        with open(report_path, "w") as f:
            json.dump(report, f, indent=2)

        output.print_md(
            ":white_check_mark: Report saved to: **{}**".format(report_path))

    except Exception as e:
        output.print_md(
            ":warning: Could not save report to desktop: {}".format(str(e)))
        report_path = "Not saved"

    # -------------------------------------------------------------------------
    # STEP 7  -  Print summary to output window
    # -------------------------------------------------------------------------
    summary = report.get("system_summary", {})
    validation_summary = report.get("validation_summary", {})
    longest_run = report.get("longest_run") or {}

    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("| Item | Value |")
    output.print_md("| --- | --- |")
    output.print_md("| Total fixtures found | {} |".format(
        summary.get("total_fixtures", 0)))
    output.print_md("| Total connected load | {} MBH |".format(
        summary.get("total_load_mbh", 0)))
    output.print_md("| Total pipe segments | {} |".format(
        summary.get("total_pipe_segments", 0)))
    output.print_md("| Longest run | {} ft |".format(
        summary.get("longest_run_feet", 0)))
    output.print_md("| Farthest fixture | {} |".format(
        longest_run.get("farthest_fixture_name", " - ")))
    output.print_md("| Revit version | {} |".format(revit_version))

    # Fixtures list
    fixtures = report.get("fixtures_found", [])
    if fixtures:
        output.print_md("---")
        output.print_md("## Fixtures Found")
        output.print_md("| Name | Load (MBH) | Element ID |")
        output.print_md("| --- | --- | --- |")
        for f in fixtures:
            output.print_md("| {} | {} | {} |".format(
                f.get("fixture_name", " - "),
                f.get("gas_load_mbh", 0),
                f.get("element_id", " - ")
            ))

    # Validation results
    output.print_md("---")
    output.print_md("## Validation")

    checks = validation_summary.get("checks", [])
    for c in checks:
        icon = ":white_check_mark:" if c["result"] == "PASS" else (
            ":warning:" if c["result"] == "WARN" else ":cross_mark:")
        output.print_md("{} {}".format(icon, c["check"]))

    errors = validation_summary.get("errors_list", [])
    warnings = validation_summary.get("warnings_list", [])

    if errors:
        output.print_md("---")
        output.print_md("## :cross_mark: Errors  -  Fix Before Sizing")
        for e in errors:
            output.print_md("- {}".format(e))

    if warnings:
        output.print_md("---")
        output.print_md("## :warning: Warnings")
        for w in warnings:
            output.print_md("- {}".format(w))

    if not errors:
        output.print_md("---")
        output.print_md(
            ":white_check_mark: **System is ready for Phase 2 sizing.**")

    output.print_md("---")
    output.print_md("*Report saved to: {}*".format(report_path))


# --- Run ---
main()
