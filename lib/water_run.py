# -*- coding: ascii -*-
# water_run.py
# The Size Water pipeline.
#
# ONE action. Pick the RPZ, answer the one dialog, and the run does the whole
# job in order:
#
#   traverse -> detect the return system -> dialog -> check the system ->
#   size -> print the report -> write the sizes -> WSFU take-off on a
#   drafting view and a sheet
#
# Two check boxes say what a run does. SIZING is always on. WSFU CALCULATIONS
# is optional, and produces the take-off: the check ON the sizing, how the
# fixture unit count that drove every pipe size gets verified against the
# fixtures actually in the model.
#
# The traversal runs BEFORE the dialog, because the dialog reports the hot
# water RETURN system rather than asking for it, and that is read off the
# System Types of the pipes the traversal found.
#
# Firm standards are assumed rather than asked about on every run: minimum
# pipe sizes always apply.
#
# IronPython 2.7 / pyRevit

import datetime

from Autodesk.Revit.DB import BuiltInParameter, Transaction
from Autodesk.Revit.UI.Selection import ObjectType

import shared_params
import revit_helpers
import water_graph
import water_checks
import water_sizing_engine
import water_report
import water_drafting
import ui_helpers


def run(doc, uidoc, output, forms):
    """Run the tool: size the model and produce the take-off that checks it.

    Returns True if it finished.
    """
    output.print_md("# Size Water")

    origin = _pick_origin(uidoc, doc, output, forms)
    if origin is None:
        return False

    # Traverse BEFORE the dialog. The dialog reports which hot water System
    # Type is the return, and that is read off the pipes the traversal found,
    # so the network has to exist before the dialog can say anything true
    # about it.
    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc)

    detection = water_graph.detect_return_system_types(graph)
    _print_return_detection(output, detection)

    settings = ui_helpers.show_water_dialog(
        "Size Water - 2024 IPC", _project_info(doc, output), detection)
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False

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
            title="Size Water - Incomplete System", yes=True, no=True)
        if not proceed:
            output.print_md("Stopped. Nothing was changed.")
            return False

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

    # Sizes first, because writing them is the job. The take-off follows in
    # its own transaction, so a project missing a drafting view family type or
    # a title block costs the user a sheet, never the sizes already committed.
    _write_sizes(doc, output, forms, graph, sizing)

    if settings["wsfu_calcs"]:
        _create_drafting_view(doc, output, graph, sizing, header)
        output.print_md(
            "Check the take-off against the fixtures in the model. The "
            "fixture unit totals on it are what drove every pipe size.")
    else:
        output.print_md(
            "WSFU Calculations were not requested, so no drafting view or "
            "sheet was made. The take-off above is print only.")
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


def _print_return_detection(output, detection):
    """Say in the window which hot water System Type was taken as the return.

    System CLASSIFICATION is DomesticHotWater for the hot supply and for the
    recirculation system alike, so it can never separate them. The System TYPE
    does: they are two different PipingSystemType elements. The traversal
    records each pipe's System Type, so the return is worked out from the
    model rather than asked about.
    """
    candidates = (detection or {}).get("candidates") or []
    if not candidates:
        output.print_md(
            "No hot water piping found on this network, so there is no "
            "return system to identify.")
        return

    output.print_md("## Hot water System Types found")
    for entry in candidates:
        label = "RETURN" if entry.get("detected") else "SUPPLY"
        reasons = entry.get("reasons") or []
        output.print_md("- **{}** - `{}`, {} pipe(s){}".format(
            label, entry["name"], entry["pipe_count"],
            ", because " + "; ".join(reasons) if reasons else ""))

    if not detection.get("certain") and any(
            e.get("detected") for e in candidates):
        output.print_md(
            "> No recirculation pump and no return-to-heater connection was "
            "found, so the RETURN above is a best guess from the System Type "
            "name and pipe count. A system marked RETURN is left unsized. "
            "Check it against the model before using these sizes.")


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
