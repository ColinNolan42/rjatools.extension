# -*- coding: ascii -*-
# water_run.py
# The shared Size Water / Water Report pipeline.
#
# Both buttons do the same work up to the last step, so the whole run lives
# here and each pushbutton is a few lines. Keeping it in one place is the
# point: a duplicated pipeline is how the two would end up reporting
# different numbers for the same system.
#
#   Water Report -> everything, then the take-off on a drafting view and a
#                   sheet. The model is not touched.
#   Size Water   -> everything, then the new sizes written to the pipes. The
#                   report goes to the pyRevit window only.
#
# Firm standards are assumed rather than asked about on every run: minimum
# pipe sizes always apply, and the report always goes on a drafting view and
# a sheet. The only runtime question is which piping system types are the hot
# water RETURN, which no standard can answer because Revit classifies a
# recirculation system exactly like the hot supply.
#
# IronPython 2.7 / pyRevit

import datetime

from Autodesk.Revit.DB import (
    BuiltInParameter, Transaction, FilteredElementCollector)
from Autodesk.Revit.DB.Plumbing import PipingSystemType
from Autodesk.Revit.UI.Selection import ObjectType

import shared_params
import revit_helpers
import water_graph
import water_checks
import water_sizing_engine
import water_report
import water_drafting
import ui_helpers


MODE_REPORT = ui_helpers.MODE_REPORT
MODE_SIZE = ui_helpers.MODE_SIZE


def run(doc, uidoc, output, forms):
    """Run the tool. The dialog's two action buttons choose what happens.

    Returns True if it finished.
    """
    output.print_md("# Size Water")

    origin = _pick_origin(uidoc, doc, output, forms)
    if origin is None:
        return False

    settings = ui_helpers.show_water_dialog(
        "Size Water - 2024 IPC",
        _project_info(doc, output),
        _hot_water_system_types(doc, output))
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False

    mode = settings["mode"]
    title = "Water Report" if mode == MODE_REPORT else "Size Water"

    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc)

    # Completeness check, before anything is sized or written. A half-modelled
    # system should fail loudly rather than be quietly sized.
    checks = water_checks.check_system(
        graph, return_system_type_ids=settings["return_system_type_ids"])
    print("\n".join(water_checks.format_checks(checks)))

    if not checks["ready_for_sizing"]:
        detail = "\n".join("  - " + f.message for f in checks["errors"][:6])
        proceed = forms.alert(
            "The system is not complete. {} error(s) found:\n\n{}\n\n"
            "Continuing would produce numbers that do not describe the whole "
            "system. Continue anyway?".format(len(checks["errors"]), detail),
            title="{} - Incomplete System".format(title), yes=True, no=True)
        if not proceed:
            output.print_md("Stopped. Nothing was changed.")
            return False

    # Sizes are always computed, in both modes. The report shows them; only
    # Size Water writes them.
    sizing = water_sizing_engine.size_network(
        graph,
        apply_minimums=True,
        return_system_type_ids=settings["return_system_type_ids"])

    header = {
        "date": datetime.date.today().strftime("%Y-%m-%d"),
        "job": settings["job"],
        "job_number": settings["job_number"],
        "by": settings["by"],
    }
    print(water_report.build_report(graph, sizing, header))

    if mode == MODE_REPORT:
        _create_drafting_view(doc, output, graph, sizing, header)
        output.print_md(
            "Report only. No pipe size was written to the model. Use "
            "**Size Water** to apply them.")
        return True

    _write_sizes(doc, output, forms, graph, sizing)
    return True


# ---------------------------------------------------------------------------
# Steps
# ---------------------------------------------------------------------------

def _pick_origin(uidoc, doc, output, forms):
    try:
        reference = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Select the RPZ / backflow preventer to start from")
    except Exception:
        output.print_md("Cancelled. Nothing was changed.")
        return None

    origin = doc.GetElement(reference.ElementId)
    if origin is None:
        forms.alert("Could not read the selected element.", title="Water")
        return None

    output.print_md("Origin: element **{}**".format(
        revit_helpers.eid_int(origin.Id)))

    # Warn but do not block. water_checks reports this properly once the
    # traversal has run; this is just an early exit for an obvious mis-pick.
    connectors = revit_helpers.get_connectors(origin)
    has_cold = any(
        c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER
        for c in connectors)
    if not has_cold:
        proceed = forms.alert(
            "The element you picked has no Domestic Cold Water connector, so "
            "the traversal will probably find nothing.\n\n"
            "Pick the RPZ / backflow preventer on the cold water main.\n\n"
            "Continue anyway?",
            title="Water", yes=True, no=True)
        if not proceed:
            return None
    return origin


def _project_info(doc, output):
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
            "Could not read Project Information ({}). Fill the header in by "
            "hand.".format(str(exc)))
    try:
        info["by"] = doc.Application.Username or ""
    except Exception:
        pass
    return info


def _hot_water_system_types(doc, output):
    """Every PipingSystemType classified as Domestic Hot Water, with a count
    of how many pipes in the model actually use it.

    Both the hot supply and the recirculation system classify identically, so
    the user has to say which is which. The pipe count is shown next to each
    name because it is what makes the answer obvious at a glance without the
    tool hardcoding any project's naming: on a real job the return carries far
    fewer pipes than the supply (11 against 66 on Grantham 4).

    Returns a list of (element_id, name, pipe_count).
    """
    found = []
    try:
        for element in FilteredElementCollector(doc).OfClass(PipingSystemType):
            try:
                classification = str(element.SystemClassification)
            except Exception:
                continue
            if classification == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
                name = (revit_helpers.get_element_type_name(element) or
                        _safe_name(element))
                found.append([revit_helpers.eid_int(element.Id), name, 0])
    except Exception as exc:
        output.print_md(
            "Could not list piping system types ({}).".format(str(exc)))
        return []

    by_name = {}
    for entry in found:
        by_name[entry[1]] = entry

    try:
        from Autodesk.Revit.DB import BuiltInCategory
        pipes = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_PipeCurves)
                 .WhereElementIsNotElementType())
        for pipe in pipes:
            try:
                param = pipe.get_Parameter(
                    BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
                if param is None:
                    continue
                entry = by_name.get(param.AsValueString())
                if entry is not None:
                    entry[2] += 1
            except Exception:
                continue
    except Exception as exc:
        output.print_md(
            "Could not count pipes per system type ({}).".format(str(exc)))

    found.sort(key=lambda row: row[1])
    return [tuple(row) for row in found]


def _safe_name(element):
    """Last-resort name read. .Name throws under IronPython often enough that
    it is never the first choice, but a PipingSystemType has no type
    parameter to read instead."""
    try:
        return element.Name
    except Exception:
        return "(unnamed system type)"


def _create_drafting_view(doc, output, graph, sizing, header):
    """The take-off on a drafting view and a new sheet, per firm standard.

    Its own transaction so a failure here cannot roll anything else back.
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

        stamp = datetime.datetime.now().strftime("%m%d-%H%M")
        sheet, viewport = water_drafting.place_on_sheet(
            doc, view, content_h, content_w,
            "WS-{}".format(stamp),
            "Domestic Water WSFU Take-Off {}".format(stamp))
        if sheet is None:
            output.print_md(
                "The drafting view was created, but no sheet was made: this "
                "project has no title block type loaded.")
        else:
            output.print_md("Placed on sheet **{} - {}**.".format(
                sheet.SheetNumber, sheet.Name))
            if viewport is None:
                output.print_md(
                    "The sheet was made but the viewport could not be placed. "
                    "Drag the view on by hand.")
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        output.print_md(
            "Drafting view failed, nothing was added to the model: "
            "{}".format(str(exc)))


def _write_sizes(doc, output, forms, graph, sizing):
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

    output.print_md("Wrote **{}** pipe size(s). {} failed.".format(
        write["written"], write["failed"]))
    for error in write["errors"][:20]:
        output.print_md("  " + error)

    # Fittings do not follow their pipes in Revit, so elbows and tees keep the
    # size they were drawn at until they are set explicitly. Separate
    # transaction: a fitting family that refuses to resize must not roll back
    # the pipe sizes that already committed.
    try:
        fittings = water_sizing_engine.write_fitting_sizes(
            doc, graph, sizing, lambda d, name: Transaction(d, name))
        output.print_md("Resized **{}** fitting(s). {} skipped.".format(
            fittings["resized"], fittings["skipped"]))
        for error in fittings["errors"][:10]:
            output.print_md("  " + error)
        if fittings["skipped"]:
            output.print_md(
                "A skipped fitting either touches no sized pipe, or its "
                "family exposes no size parameter the tool can set. Check "
                "those by hand before issuing.")
    except Exception as exc:
        output.print_md(
            "Fitting resize failed, pipe sizes are unaffected: {}".format(
                str(exc)))
