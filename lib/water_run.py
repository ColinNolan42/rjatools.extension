# -*- coding: ascii -*-
# water_run.py
# The Size Water pipeline.
#
# ONE action, in the order the user sees it (Colin, 2026-10-13): the Size
# Water dialog first, then a choice of what to pick (RPZ or a new pipe at a
# tee to existing), then the pick, then the whole job:
#
#   dialog -> choose RPZ / tee -> pick -> traverse -> detect the return
#   system -> split supply from return -> check the system -> circuit setter
#   flows -> size -> print the report -> write the sizes
#
# Sizing is the only job. The WSFU take-off drafting view and sheet belong to
# Water Performance (Colin, 2026-10-21); the take-off still prints here.
#
# The dialog opens before the network is known, so it cannot grey out the
# recirculation inputs. They stay live and are used only if the traversal
# finds hot water return piping. The return system itself is still detected
# from the pipes' System Types after the pick, never asked about.
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
import tie_in
import water_graph
import water_checks
import water_sizing_engine
import water_recirc
import water_tables
import water_report
import ui_helpers


def run(doc, uidoc, output, forms):
    """Run the tool: size the model and produce the take-off that checks it.

    Returns True if it finished.
    """
    output.print_md("# Size Water")

    settings = ui_helpers.show_water_dialog(
        "Size Water - 2024 IPC", project_info(doc, output), None)
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False

    origin, upstream = pick_origin(uidoc, doc, output, forms)
    if origin is None:
        return False

    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc, upstream)

    detection = water_graph.detect_return_system_types(graph)
    _print_return_detection(output, detection)

    # Split supply from return before anything is counted. The hot walk goes
    # round a recirculation loop from both ends, so until the return pipes are
    # taken out of the supply tree the hot fixture units can land on the
    # wrong side of it.
    # Only a return that something confirmed. One picked on pipe count alone
    # is reported and left unsized, never acted on.
    water_graph.apply_return_system_types(
        graph, detection["detected"] - detection["unconfirmed"])
    # Detected, not chosen: the same set the dialog used to carry through.
    settings["return_system_type_ids"] = set(detection["detected"])

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

    # Circuit setter flows and the flow in every return pipe. None when the
    # system has no return, so the report and the engine say nothing about it.
    recirc = None
    if settings["return_system_type_ids"]:
        recirc = water_recirc.analyze(
            graph,
            heat_loss_btu_per_100ft=settings["recirc_heat_loss_btu_per_100ft"],
            delta_t_f=settings["recirc_delta_t_f"])

    sizing = water_sizing_engine.size_network(
        graph,
        apply_minimums=True,
        return_system_type_ids=settings["return_system_type_ids"],
        recirc=recirc)

    header = {
        "date": datetime.date.today().strftime("%Y-%m-%d"),
        "job": settings["job"],
        "job_number": settings["job_number"],
        "by": settings["by"],
    }
    print(water_report.build_report(graph, sizing, header, recirc))

    _warn_over_table(output, forms, sizing)
    _write_sizes(doc, output, forms, graph, sizing)
    output.print_md(
        "Run Water Performance with WSFU Calculations for the take-off on a "
        "drafting view and a sheet.")
    return True


# ---------------------------------------------------------------------------
# Steps
# ---------------------------------------------------------------------------

ORIGIN_RPZ = "RPZ"
ORIGIN_TEE = "New pipe connection (TEE)"


def pick_origin(uidoc, doc, output, forms):
    """Pick where the run starts.

    Returns (origin element, upstream). upstream is None for an RPZ, and what
    tie_in.resolve_tee_branch() found for a tee on an existing main. Returns
    (None, None) when the pick is cancelled or cannot be used.
    """
    choice = ui_helpers.show_start_point(
        "Domestic Water - Select Start Point", ORIGIN_RPZ, ORIGIN_TEE)
    if not choice:
        output.print_md("Cancelled. Nothing was changed.")
        return None, None

    prompt = ("Select the RPZ / backflow preventer" if choice == ORIGIN_RPZ
              else "Select the tee on the existing cold water main")
    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, prompt)
    except Exception:
        output.print_md("Cancelled. Nothing was changed.")
        return None, None

    origin = doc.GetElement(reference.ElementId)
    if origin is None:
        forms.alert("Could not read the selected element.", title="Water")
        return None, None

    output.print_md("Origin: element **{}**".format(
        revit_helpers.eid_int(origin.Id)))

    if choice == ORIGIN_TEE:
        return _tee_origin(origin, doc, output, forms)

    if revit_helpers.is_pipe(origin):
        forms.alert("RPZ was chosen, but a pipe was picked. Pick the RPZ / "
                    "backflow preventer, or run again and choose New pipe "
                    "connection (TEE).", title="Size Water")
        output.print_md("Stopped. Nothing was changed.")
        return None, None

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
            return None, None
    return origin, None


def _tee_origin(tee, doc, output, forms):
    """A tee on the existing cold water main, picked on an existing building.

    The tee shows both sides at once: the existing main runs through it and
    the new branch leaves it (Colin, 2026-10-13). Only the branch is walked.
    The existing main and RPZ are assumed adequate for the added load.
    """
    def _stop(message):
        forms.alert(message, title="Size Water")
        output.print_md("Stopped. Nothing was changed.")
        return None, None

    if revit_helpers.is_pipe(tee):
        return _stop("New pipe connection (TEE) was chosen, but a pipe was "
                     "picked. Pick the tee fitting on the existing cold water "
                     "main where the new piping branches off.")

    cold = [c for c in revit_helpers.get_connectors(tee)
            if c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER]
    if len(cold) < 3:
        return _stop("The element picked is not a cold water tee (it has {} "
                     "cold water connector(s), a tee has 3). Pick the tee on "
                     "the existing cold water main.".format(len(cold)))

    upstream = tie_in.resolve_tee_branch(
        tee, doc, shared_params.SYSTEM_DOMESTIC_COLD_WATER)
    upstream["tee"] = True
    if upstream["status"] != "ok":
        return _stop("Could not find the new branch of the picked tee: {}.\n\n"
                     "The new branch is told from the existing main by phase: "
                     "the pipe or fitting on the branch shall be in the "
                     "project's last phase (New Construction), and the main on "
                     "either side in an earlier phase, open or capped.".format(
                         upstream["reason"]))

    output.print_md(
        "Starting at the **tee** on the existing main. Only its new "
        "construction branch is sized, because {}. The existing cold water "
        "main is assumed adequate for the added load and is not "
        "checked.".format(upstream["reason"]))
    return tee, upstream


def project_info(doc, output):
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
            "name and pipe count. A return identified by name is sized on "
            "circulation flow instead of fixture units. One picked on pipe "
            "count alone is left unsized. Check it against the model before "
            "using these sizes.")


def _warn_over_table(output, forms, sizing):
    """Warn when a load is above the largest size in the firm table.

    Those pipes are NOT sized and keep their drawn size, so the run carries
    on, but the user shall know before trusting the result.
    """
    over = [s for s in sizing["segments"] if s.exceeds_table]
    if not over:
        return
    largest = water_tables.firm_sizing_rows()[-1]
    limits = "{}\" carries up to {} WSFU cold, {} WSFU hot, {:.1f} gpm return".format(
        largest["nominal_size"],
        water_tables.fmt_number(largest["max_wsfu_8fps_cold"]),
        water_tables.fmt_number(largest["max_wsfu_5fps_hot"]),
        float(water_tables.return_gpm_rows()[-1]["max_gpm"]))
    lines = ["  - Pipe {}: {}".format(s.element_id, s.flag) for s in over[:8]]
    if len(over) > 8:
        lines.append("  ... and {} more".format(len(over) - 8))
    message = ("{} pipe(s) carry more load than the firm sizing table "
               "covers ({}).\n\n{}\n\nThese pipes were NOT sized and keep "
               "their drawn size. Size them by hand.".format(
                   len(over), limits, "\n".join(lines)))
    output.print_md("> **WARNING:** " + message.replace("\n", "  \n"))
    forms.alert(message, title="Size Water - Over the Table", warn_icon=True)


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

    # Fixture stubs follow the pipe sized onto them (RJA fixture family Size
    # Override). Own transaction, same reason as the fittings.
    try:
        stubs = water_sizing_engine.write_fixture_connection_sizes(
            doc, graph, sizing, lambda d, name: Transaction(d, name))
        output.print_md(
            "Fixture stubs: **{}** set to a larger pipe size, {} back to the "
            "Type size, {} skipped.".format(
                stubs["set"], stubs["cleared"], stubs["skipped"]))
        for error in stubs["errors"][:10]:
            output.print_md("  " + error)
    except Exception as exc:
        output.print_md(
            "Fixture stub sizing failed, pipe sizes are unaffected: {}".format(
                str(exc)))
