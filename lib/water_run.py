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

    The button path: dialog, start point choice, pick, then the same core the
    headless entry uses. Returns True if it finished.
    """
    import headless
    prompter = headless.UiPrompter(forms)
    output.print_md("# Size Water")

    settings = ui_helpers.show_water_dialog(
        "Size Water - 2024 IPC", project_info(doc, output), None)
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False

    origin, upstream = pick_origin(uidoc, doc, output, forms)
    if origin is None:
        return False

    outcome = size_core(doc, settings, origin, upstream, output, prompter,
                        write=True,
                        transaction_factory=lambda d, name: Transaction(d, name))
    return outcome["status"] == "ok"


def size_core(doc, settings, origin, upstream, output, prompter, write=True,
              transaction_factory=None):
    """Traverse, check, size, report and (write=True) write the sizes.

    Shared by the button and run_headless. The yes/no question goes through
    prompter.confirm (qid "incomplete_system"); HeadlessPrompter raises
    NeedsDecision there, before anything is written.

    Returns dict: status ("ok" or "cancelled"), graph, detection, checks,
    recirc, sizing, writes (None when nothing was written), errors.
    """
    outcome = {"status": "ok", "graph": None, "detection": None,
               "checks": None, "recirc": None, "sizing": None,
               "writes": None, "errors": []}

    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc, upstream)
    outcome["graph"] = graph

    detection = water_graph.detect_return_system_types(graph)
    outcome["detection"] = detection
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
    outcome["checks"] = checks
    _plain(output, "\n".join(water_checks.format_checks(checks)))

    if not checks["ready_for_sizing"]:
        detail = "\n".join("  - " + f.message for f in checks["errors"][:6])
        proceed = prompter.confirm(
            "incomplete_system",
            "The system is not complete. {} error(s) found:\n\n{}\n\n"
            "Continuing would produce numbers that do not describe the whole "
            "system. Continue anyway?".format(len(checks["errors"]), detail),
            title="Size Water - Incomplete System", suggested="stop")
        if not proceed:
            output.print_md("Stopped. Nothing was changed.")
            outcome["status"] = "cancelled"
            return outcome

    # Circuit setter flows and the flow in every return pipe. None when the
    # system has no return, so the report and the engine say nothing about it.
    recirc = None
    if settings["return_system_type_ids"]:
        recirc = water_recirc.analyze(
            graph,
            heat_loss_btu_per_100ft=settings["recirc_heat_loss_btu_per_100ft"],
            delta_t_f=settings["recirc_delta_t_f"])
    outcome["recirc"] = recirc

    sizing = water_sizing_engine.size_network(
        graph,
        apply_minimums=True,
        return_system_type_ids=settings["return_system_type_ids"],
        recirc=recirc)
    outcome["sizing"] = sizing

    header = {
        "date": datetime.date.today().strftime("%Y-%m-%d"),
        "job": settings["job"],
        "job_number": settings["job_number"],
        "by": settings["by"],
    }
    _plain(output,
           water_report.build_compact_report(graph, sizing, header, recirc))

    _warn_over_table(output, prompter, sizing)
    if write:
        outcome["writes"] = _write_sizes(doc, output, prompter, graph, sizing,
                                         transaction_factory, outcome["errors"])
    else:
        _print_existing_upsizes(output, sizing)
        output.print_md("Plan only (write=False): nothing was written.")
    output.print_md(
        "Run Water Performance with WSFU Calculations for the take-off on a "
        "drafting view and a sheet.")
    return outcome


def _plain(output, text):
    """Plain text block. The pyRevit window takes print() as before; a
    headless.RecordingOutput keeps it as a code block."""
    import headless
    if isinstance(output, headless.RecordingOutput):
        output.print_code(text)
    else:
        print(text)


# ---------------------------------------------------------------------------
# Headless entry (pyRevit agent runtime: no dialog, no pick, no alert)
# ---------------------------------------------------------------------------

TOOL = "size_water"
START_RPZ = "rpz"
START_TEE = "tee"
DETAILS = ("flagged", "all")


def default_settings(doc=None):
    """What the dialog prefills: the Project Information header and the firm
    recirculation defaults."""
    info = {"job": "", "job_number": "", "by": ""}
    if doc is not None:
        import headless
        info = project_info(doc, headless.RecordingOutput())
    info.update({
        "recirc_heat_loss_btu_per_100ft":
            float(shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT),
        "recirc_delta_t_f": float(shared_params.DEFAULT_RECIRC_DELTA_T_F),
    })
    return info


def merge_settings(overrides, doc=None):
    """default_settings with the caller's values on top. An unknown key is an
    error, so a typo never runs silently on the default."""
    settings = default_settings(doc)
    unknown = sorted(set(overrides or {}) - set(settings))
    if unknown:
        raise ValueError("unknown Size Water setting(s): {}".format(
            ", ".join(unknown)))
    settings.update(overrides or {})
    return settings


def validate_settings(settings):
    """Problems as text, empty when the settings can run. Same limits as the
    dialog."""
    problems = []
    try:
        if float(settings["recirc_heat_loss_btu_per_100ft"]) <= 0:
            problems.append("recirc_heat_loss_btu_per_100ft shall be above 0")
    except (TypeError, ValueError):
        problems.append("recirc_heat_loss_btu_per_100ft is not a number")
    max_dt = float(shared_params.MAX_RECIRC_DELTA_T_F)
    try:
        dt = float(settings["recirc_delta_t_f"])
        if dt <= 0 or dt > max_dt:
            problems.append("recirc_delta_t_f shall be above 0 and at most "
                            "{:.0f} F".format(max_dt))
    except (TypeError, ValueError):
        problems.append("recirc_delta_t_f is not a number")
    return problems


def run_headless(doc, start_id, start=START_RPZ, settings=None,
                 decisions=None, write=True, detail="flagged",
                 include_log=False, output=None):
    """Size Water with no UI. Returns the headless.result envelope.

    start_id: the RPZ (start "rpz") or the tee on the existing cold main
    (start "tee"). decisions: {"incomplete_system" | "no_cold_connector":
    "continue" | "stop"}. write=False sizes and reports only (run_query);
    write=True writes pipes, fittings and fixture stubs, each transaction
    with the warning deleting failure handler (run_modify). detail "flagged"
    lists pipes whose size changes, with a flag, unsized, over the table or
    existing and too small; "all" lists every pipe. output: optional window to forward the report to.
    """
    import headless
    recorder = headless.RecordingOutput(forward=output)
    state = {"prompter": None}

    def done(status, data=None, errors=None, decision_exc=None):
        log = recorder.log if include_log else None
        prompter = state["prompter"]
        warnings = list(prompter.warnings) if prompter else []
        if decision_exc is not None:
            return headless.needs_decision(TOOL, decision_exc, warnings, log)
        return headless.result(TOOL, status, data=data, warnings=warnings,
                               errors=errors, log=log)

    try:
        state["prompter"] = headless.HeadlessPrompter(decisions)
        settings = merge_settings(settings, doc)
    except ValueError as exc:
        return done("refused", errors=[str(exc)])
    problems = validate_settings(settings)
    if start not in (START_RPZ, START_TEE):
        problems.append("start shall be 'rpz' or 'tee'")
    if detail not in DETAILS:
        problems.append("detail shall be 'flagged' or 'all'")
    if problems:
        return done("refused", errors=problems)

    recorder.print_md("# Size Water (headless)")
    try:
        origin, upstream, stop = resolve_origin(
            doc, _element(doc, start_id), start, state["prompter"], recorder)
        if stop == _DECLINED:
            return done("cancelled")
        if stop:
            return done("refused", errors=[stop])
        if write:
            factory, check_tx = revit_helpers.quiet_transaction_factory()
        else:
            factory, check_tx = None, (lambda: [])
        outcome = size_core(doc, settings, origin, upstream, recorder,
                            state["prompter"], write=write,
                            transaction_factory=factory)
    except headless.NeedsDecision as exc:
        return done("needs_decision", decision_exc=exc)
    except Exception as exc:
        return done("error", errors=[revit_helpers.exc_text(exc)])

    errors = list(outcome["errors"]) + check_tx()
    data = _envelope_data(outcome, origin, upstream, start, write, detail)
    return done(outcome["status"], data=data, errors=errors)


def _element(doc, element_id):
    """Element from an integer id (Int64 for Revit 2024+), None if absent."""
    from Autodesk.Revit.DB import ElementId
    try:
        from System import Int64
        eid = ElementId(Int64(int(element_id)))
    except Exception:
        eid = ElementId(int(element_id))
    try:
        return doc.GetElement(eid)
    except Exception:
        return None


def _envelope_data(outcome, origin, upstream, start, write, detail):
    """The JSON data block of the headless result."""
    sizing = outcome["sizing"] or {"segments": [], "totals": {}, "flags": []}
    checks = outcome["checks"] or {"errors": [], "warnings": [],
                                   "ready_for_sizing": None}
    detection = outcome["detection"] or {}
    recirc = outcome["recirc"]
    pipes = []
    for s in sizing["segments"]:
        # An existing pipe left at its size is listed only when it needs
        # upsizing; a fine one is not a finding.
        changed = (not s.keep_existing and s.size_inches is not None and
                   s.existing_size_in is not None and
                   abs(float(s.size_inches) - float(s.existing_size_in)) > 1e-6)
        listed = (detail == "all" or changed or s.flag or s.nominal_size is None
                  or s.exceeds_table or s.upsize_needed)
        if not listed:
            continue
        pipes.append({
            "id": s.element_id, "system": s.system,
            "wsfu": s.demand_wsfu, "gpm": s.demand_gpm,
            "circ_gpm": s.circ_gpm, "is_return": s.is_return,
            "drawn_in": s.existing_size_in, "size": s.nominal_size,
            "size_in": s.size_inches, "changed": changed,
            "rule": s.rule, "flag": s.flag,
            "length_ft": s.length_feet, "fixtures": s.fixture_count,
            "exceeds_table": s.exceeds_table,
            "keep_existing": s.keep_existing,
            "required": s.required_nominal, "upsize_needed": s.upsize_needed,
        })
    data = {
        "origin": {"id": revit_helpers.eid_int(origin.Id), "start": start,
                   "tee_reason": (upstream or {}).get("reason")},
        "write": bool(write),
        "ready_for_sizing": checks["ready_for_sizing"],
        "check_errors": [f.message for f in checks["errors"]],
        "check_warnings": [f.message for f in checks["warnings"]],
        "return_system_types": [
            {"name": c.get("name"), "detected": bool(c.get("detected")),
             "pipes": c.get("pipe_count"), "reasons": c.get("reasons") or []}
            for c in (detection.get("candidates") or [])],
        "return_certain": detection.get("certain"),
        "totals": sizing.get("totals"),
        "flags": sizing.get("flags"),
        "pipes_listed": detail,
        "pipes": pipes,
        "writes": outcome["writes"],
    }
    if recirc:
        data["recirc"] = {
            "total_gpm": recirc.get("total_gpm"),
            "root_gpm": recirc.get("root_gpm"),
            "circuit_setters": len(recirc.get("legs") or []),
            "unrooted_pipe_ids": recirc.get("unrooted_pipe_ids"),
            "heat_loss_btu_per_100ft": recirc.get("heat_loss_btu_per_100ft"),
            "delta_t_f": recirc.get("delta_t_f"),
        }
    return data


# ---------------------------------------------------------------------------
# Steps
# ---------------------------------------------------------------------------

ORIGIN_RPZ = "RPZ"
ORIGIN_TEE = "New pipe connection (TEE)"
_DECLINED = "declined"


def pick_origin(uidoc, doc, output, forms):
    """Pick where the run starts (button path).

    Returns (origin element, upstream). upstream is None for an RPZ, and what
    tie_in.resolve_tee_branch() found for a tee on an existing main. Returns
    (None, None) when the pick is cancelled or cannot be used.
    """
    import headless
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

    start = START_TEE if choice == ORIGIN_TEE else START_RPZ
    origin, upstream, stop = resolve_origin(
        doc, origin, start, headless.UiPrompter(forms), output)
    if stop == _DECLINED:
        return None, None
    if stop:
        forms.alert(stop, title="Size Water")
        output.print_md("Stopped. Nothing was changed.")
        return None, None
    return origin, upstream


def resolve_origin(doc, element, start, prompter, output=None):
    """Check a start element and resolve where the walk begins.

    Shared by the button pick and run_headless (and later Water Performance
    and Diagnose). Returns (origin, upstream, stop). stop is None when the
    run can go on; _DECLINED when the missing cold connector question was
    answered no; otherwise the message that stopped the run.
    """
    if element is None:
        return None, None, "Could not read the selected element."
    if start == START_TEE:
        return _tee_origin(element, doc, output)

    if revit_helpers.is_pipe(element):
        return None, None, (
            "RPZ was chosen, but a pipe was picked. Pick the RPZ / backflow "
            "preventer, or run again and choose New pipe connection (TEE).")

    # Warn but do not block. water_checks reports this properly once the
    # traversal has run; this is just an early exit for an obvious mis-pick.
    connectors = revit_helpers.get_connectors(element)
    has_cold = any(
        c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER
        for c in connectors)
    if not has_cold:
        proceed = prompter.confirm(
            "no_cold_connector",
            "The element you picked has no Domestic Cold Water connector, so "
            "the traversal will probably find nothing.\n\n"
            "Pick the RPZ / backflow preventer on the cold water main.\n\n"
            "Continue anyway?",
            title="Water", suggested="stop")
        if not proceed:
            return None, None, _DECLINED
    return element, None, None


def _tee_origin(tee, doc, output=None):
    """A tee on the existing cold water main, picked on an existing building.

    The tee shows both sides at once: the existing main runs through it and
    the new branch leaves it (Colin, 2026-10-13). Only the branch is walked.
    The existing main and RPZ are assumed adequate for the added load.
    Returns (origin, upstream, stop message or None).
    """
    if revit_helpers.is_pipe(tee):
        return None, None, (
            "New pipe connection (TEE) was chosen, but a pipe was picked. Pick "
            "the tee fitting on the existing cold water main where the new "
            "piping branches off.")

    cold = [c for c in revit_helpers.get_connectors(tee)
            if c.get("system_type") == shared_params.SYSTEM_DOMESTIC_COLD_WATER]
    if len(cold) < 3:
        return None, None, (
            "The element picked is not a cold water tee (it has {} cold water "
            "connector(s), a tee has 3). Pick the tee on the existing cold "
            "water main.".format(len(cold)))

    upstream = tie_in.resolve_tee_branch(
        tee, doc, shared_params.SYSTEM_DOMESTIC_COLD_WATER)
    upstream["tee"] = True
    if upstream["status"] != "ok":
        return None, None, (
            "Could not find the new branch of the picked tee: {}.\n\n"
            "The new branch is told from the existing main by phase: the pipe "
            "or fitting on the branch shall be in the project's last phase "
            "(New Construction), and the main on either side in an earlier "
            "phase, open or capped.".format(upstream["reason"]))

    if output is not None:
        output.print_md(
            "Starting at the **tee** on the existing main. Only its new "
            "construction branch is sized, because {}. The existing cold "
            "water main is assumed adequate for the added load and is not "
            "checked.".format(upstream["reason"]))
    return tee, upstream, None


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


def _warn_over_table(output, prompter, sizing):
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
    prompter.notify(message, title="Size Water - Over the Table",
                    warn_icon=True)


def _print_existing_upsizes(output, sizing):
    """Existing phase pipes are left alone; say in red which need upsizing."""
    held = [s for s in sizing["segments"] if s.keep_existing]
    if not held:
        return
    upsize = [s for s in held if s.upsize_needed]
    output.print_md("{} EXISTING phase pipe(s) were left at their drawn "
                    "size.".format(len(held)))
    if not upsize:
        output.print_md("None of them needs upsizing for the new load.")
        return
    rows = ["<li>Pipe {}: {} existing, new load calls for {}</li>".format(
        s.element_id, s.nominal_size, s.required_nominal) for s in upsize]
    output.print_html(
        "<div style='color:#C00000;font-weight:bold'>EXISTING PIPING TO BE "
        "UPSIZED FOR THE NEW LOAD ({} pipe(s)):<ul>{}</ul></div>".format(
            len(upsize), "".join(rows)))


def _write_sizes(doc, output, prompter, graph, sizing, transaction_factory,
                 errors):
    """Write pipes, fittings and fixture stubs, each in its own transaction.
    Returns a summary dict; every failure is also appended to errors."""
    _print_existing_upsizes(output, sizing)
    summary = {"written": 0, "failed": 0, "fittings": None, "stubs": None}
    sized = sizing["totals"]["sized_count"]
    if sized == 0:
        output.print_md("Nothing to write: no pipe was sized.")
        return summary

    try:
        write = water_sizing_engine.write_sizes(
            doc, graph, sizing, transaction_factory)
    except Exception as exc:
        prompter.notify(
            "Transaction failed. No sizes were written to the model.\n\n"
            "{}".format(str(exc)),
            title="Size Water - Transaction Error")
        output.print_md("Transaction ERROR: {}".format(str(exc)))
        errors.append("Pipe write failed, nothing written: {}".format(
            revit_helpers.exc_text(exc)))
        return summary

    summary["written"] = write["written"]
    summary["failed"] = write["failed"]
    errors.extend(write["errors"])
    output.print_md("Wrote **{}** pipe size(s). {} failed.".format(
        write["written"], write["failed"]))
    for error in write["errors"][:5]:
        output.print_md("  " + error)

    # Fittings do not follow their pipes in Revit, so elbows and tees keep the
    # size they were drawn at until they are set explicitly. Separate
    # transaction: a fitting family that refuses to resize must not roll back
    # the pipe sizes that already committed.
    try:
        fittings = water_sizing_engine.write_fitting_sizes(
            doc, graph, sizing, transaction_factory)
        summary["fittings"] = {
            "resized": fittings["resized"],
            "transitions": fittings.get("transitions", 0),
            "skipped": fittings["skipped"]}
        errors.extend(fittings["errors"])
        output.print_md(
            "Resized **{}** fitting(s). {} Revit transition(s) follow their "
            "pipes. {} skipped.".format(
                fittings["resized"], fittings.get("transitions", 0),
                fittings["skipped"]))
        for error in fittings["errors"]:
            output.print_md("**{}**".format(error))
        if fittings["skipped"]:
            output.print_md(
                "A skipped fitting touches no sized pipe, is existing phase, "
                "or its family has no size parameter. Check by hand.")
    except Exception as exc:
        errors.append("Fitting resize failed: {}".format(
            revit_helpers.exc_text(exc)))
        output.print_md(
            "Fitting resize failed, pipe sizes are unaffected: {}".format(
                str(exc)))

    # Fixture stubs follow the pipe sized onto them (RJA fixture family Size
    # Override). Own transaction, same reason as the fittings.
    try:
        stubs = water_sizing_engine.write_fixture_connection_sizes(
            doc, graph, sizing, transaction_factory)
        summary["stubs"] = {"set": stubs["set"], "cleared": stubs["cleared"],
                            "skipped": stubs["skipped"]}
        errors.extend(stubs["errors"])
        output.print_md(
            "Fixture stubs: **{}** set to a larger pipe size, {} back to the "
            "Type size, {} skipped.".format(
                stubs["set"], stubs["cleared"], stubs["skipped"]))
        for error in stubs["errors"][:3]:
            output.print_md("  " + error)
        if len(stubs["errors"]) > 3:
            output.print_md("  ... and {} more.".format(
                len(stubs["errors"]) - 3))
    except Exception as exc:
        errors.append("Fixture stub sizing failed: {}".format(
            revit_helpers.exc_text(exc)))
        output.print_md(
            "Fixture stub sizing failed, pipe sizes are unaffected: {}".format(
                str(exc)))
    return summary
