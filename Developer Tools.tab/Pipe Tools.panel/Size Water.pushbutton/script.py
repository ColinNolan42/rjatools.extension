# -*- coding: ascii -*-
# Size Water.pushbutton/script.py
# Domestic water pipe sizing per the 2024 IPC.
#
# User picks the RPZ (backflow preventer) -> one startup dialog -> traversal,
# sizing, write-back and report. No prompts after the dialog.
#
# Sizes BOTH cold water (on TOTAL fixture units) and hot water (on HOT fixture
# units). The only system not sized is the hot water RETURN line, which is
# sized on circulation flow rather than on fixture units.
#
# IronPython 2.7 / pyRevit

import os
import sys
import datetime

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    BuiltInParameter, Transaction, FilteredElementCollector)
from Autodesk.Revit.DB.Plumbing import PipingSystemType
from Autodesk.Revit.UI.Selection import ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

_lib_dir = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

import shared_params
import revit_helpers
import water_graph
import water_sizing_engine
import water_report
import water_drafting
import ui_helpers


def _project_info():
    """Prefill the report header from Revit Project Information."""
    info = {"job": "", "job_number": "", "by": ""}
    try:
        project = doc.ProjectInformation
        name = project.get_Parameter(BuiltInParameter.PROJECT_NAME)
        number = project.get_Parameter(BuiltInParameter.PROJECT_NUMBER)
        if name is not None:
            info["job"] = name.AsString() or ""
        if number is not None:
            info["job_number"] = number.AsString() or ""
    except Exception as exc:
        output.print_md(
            "Could not read Project Information ({}). "
            "Fill the header in by hand.".format(str(exc)))
    try:
        info["by"] = doc.Application.Username or ""
    except Exception:
        pass
    return info


def _hot_water_system_types():
    """Every PipingSystemType classified as Domestic Hot Water.

    Both the hot supply and the recirculation system classify the same way, so
    the user has to say which is which. Names are shown to the user, never
    matched against in code.
    """
    found = []
    try:
        for element in FilteredElementCollector(doc).OfClass(PipingSystemType):
            try:
                classification = str(element.SystemClassification)
            except Exception:
                continue
            if classification == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
                found.append((revit_helpers.eid_int(element.Id), element.Name))
    except Exception as exc:
        output.print_md(
            "Could not list piping system types ({}).".format(str(exc)))
    found.sort(key=lambda pair: pair[1])
    return found


def _create_drafting_view(graph, sizing, header, settings):
    """Build the WSFU take-off drafting view, and optionally a sheet for it.

    Its own transaction, so a failure here can never roll back or block the
    pipe sizing write that follows.
    """
    transaction = Transaction(doc, "RJA Tools - Water WSFU Take-Off")
    transaction.Start()
    try:
        view, content_h, content_w = water_drafting.build_wsfu_view(
            doc, graph, header, sizing)

        if view is None:
            transaction.RollBack()
            output.print_md(
                "Could not create the drafting view. The project needs a "
                "Drafting view family type and at least one Text Note type.")
            return

        output.print_md("Created drafting view **{}**.".format(view.Name))

        if settings["place_on_sheet"]:
            stamp = datetime.datetime.now().strftime("%m%d-%H%M")
            sheet, viewport = water_drafting.place_on_sheet(
                doc, view, content_h, content_w,
                "WS-{}".format(stamp),
                "Domestic Water WSFU Take-Off {}".format(stamp))
            if sheet is None:
                output.print_md(
                    "The drafting view was created, but no sheet was made: "
                    "this project has no title block type loaded.")
            else:
                output.print_md(
                    "Placed on sheet **{} - {}**.".format(
                        sheet.SheetNumber, sheet.Name))
                if viewport is None:
                    output.print_md(
                        "The sheet was made but the viewport could not be "
                        "placed. Drag the view on by hand.")

        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        output.print_md(
            "Drafting view failed, nothing was added to the model: "
            "{}".format(str(exc)))


def main():
    output.print_md("# Size Water")

    # --- 1. Pick the RPZ ---
    try:
        reference = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Select the RPZ / backflow preventer to size from")
    except Exception:
        output.print_md("Cancelled. Nothing was changed.")
        return

    origin = doc.GetElement(reference.ElementId)
    if origin is None:
        forms.alert("Could not read the selected element.",
                    title="Size Water")
        return

    origin_id = revit_helpers.eid_int(origin.Id)
    output.print_md("Origin: element **{}**".format(origin_id))

    # Warn, but do not block: the origin only has to be a real starting point
    # on the cold water network.
    origin_connectors = revit_helpers.get_connectors(origin)
    has_cold = any(c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER
                   for c in origin_connectors)
    if not has_cold:
        proceed = forms.alert(
            "The element you picked has no Domestic Cold Water connector, so "
            "the traversal will probably find nothing.\n\n"
            "Pick the RPZ / backflow preventer on the cold water main.\n\n"
            "Continue anyway?",
            title="Size Water", yes=True, no=True)
        if not proceed:
            return

    # --- 2. One dialog ---
    settings = ui_helpers.show_water_dialog(
        "Size Water - 2024 IPC", _project_info(), _hot_water_system_types())
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return

    # --- 3. Traverse ---
    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc)

    if not graph.fixture_ids:
        forms.alert(
            "No water fixtures were found on the network from that element.\n\n"
            "Check that the fixtures carry IS_WATER_FIXTURE and that the "
            "piping is actually connected.",
            title="Size Water")

    # --- 4. Size ---
    sizing = water_sizing_engine.size_network(
        graph,
        apply_minimums=settings["apply_minimums"],
        return_system_type_ids=settings["return_system_type_ids"])

    # --- 5. Report ---
    header = {
        "date": datetime.date.today().strftime("%Y-%m-%d"),
        "job": settings["job"],
        "job_number": settings["job_number"],
        "by": settings["by"],
    }
    report = water_report.build_report(graph, sizing, header)
    print(report)

    # --- 6. Drafting view ---
    # Runs even in report-only mode: report-only means "do not change pipe
    # sizes", not "produce nothing". The take-off table is a deliverable.
    if settings["create_drafting_view"]:
        _create_drafting_view(graph, sizing, header, settings)

    # --- 7. Write back ---
    if settings["report_only"]:
        output.print_md(
            "**Report only.** No sizes were written to the model.")
        return

    sized = sizing["totals"]["sized_count"]
    if sized == 0:
        output.print_md("Nothing to write: no pipe was sized.")
        return

    try:
        write = water_sizing_engine.write_sizes(
            doc, graph, sizing, lambda d, name: Transaction(d, name))
    except Exception as exc:
        forms.alert(
            "Transaction failed. No sizes were written to the model.\n\n"
            "{}".format(str(exc)),
            title="Size Water - Transaction Error")
        output.print_md("Transaction ERROR: {}".format(str(exc)))
        return

    output.print_md(
        "Wrote **{}** pipe size(s). {} failed.".format(
            write["written"], write["failed"]))
    for error in write["errors"][:20]:
        output.print_md("  " + error)

    output.print_md(
        "Fittings are NOT resized by this build. Check transitions at tees "
        "and elbows before issuing.")


main()
