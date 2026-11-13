# -*- coding: ascii -*-
# water_perf_run.py
# The Water Performance pipeline. READ ONLY apart from one duplicated view.
#
#   dialog -> choose RPZ / tee -> pick -> traverse -> split return -> size
#   (in memory, nothing written) -> judge each pipe against the firm table ->
#   copy the active plan view and color it -> print the flagged pipes
#
# Pass/fail is the drawn size against the firm table size. Computed speed is
# shown, and raises a pipe at table size to yellow. See water_performance.py.
#
# IronPython 2.7 / pyRevit

import datetime
import re

from Autodesk.Revit.DB import (
    Color, ElementId, FamilySymbol, FilteredElementCollector,
    OverrideGraphicSettings, StorageType, Transaction, ViewDuplicateOption,
    ViewPlan)

import shared_params
import revit_helpers
import water_graph
import water_recirc
import water_sizing_engine
import water_tables
import water_performance
import water_run
import water_drafting
import water_pressure
import water_pump
import water_checks
import water_flagging
import water_system_errors
import hvac_graph
import ui_helpers

_COLORS = {
    water_performance.GREEN:  Color(0, 200, 0),
    water_performance.YELLOW: Color(255, 215, 0),
    water_performance.RED:    Color(210, 40, 40),
    water_performance.PURPLE: Color(140, 40, 200),
    water_performance.GRAY:   Color(160, 160, 160),
}


def run(doc, uidoc, output, forms):
    """Run the tool (button path): dialog, plan view check, start point pick,
    then the same core the headless entry uses. Returns True if it finished."""
    output.print_md("# Water Performance")

    view = doc.ActiveView
    settings = ui_helpers.show_water_performance_dialog(
        "Water Performance - 2024 IPC", water_run.project_info(doc, output))
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False
    stop = check_plan_view(view, settings)
    if stop:
        forms.alert(stop, title="Water Performance")
        return False

    origin, upstream = water_run.pick_origin(uidoc, doc, output, forms)
    if origin is None:
        return False

    outcome = perf_core(doc, settings, origin, upstream, view, output)
    return outcome["status"] == "ok"


def check_plan_view(view, settings, make_sheets=True):
    """None when the view can take the colored copy, else why not. Only the
    size and velocity check with sheets colors a copy of the plan."""
    if not (make_sheets and settings.get("velocity")):
        return None
    if not isinstance(view, ViewPlan) or view.IsTemplate:
        return ("Open the floor plan to check, then run again. Pipe size "
                "and velocity colors a copy of the active plan view.")
    return None


def perf_core(doc, settings, origin, upstream, view, output, make_sheets=True):
    """Traverse, size in memory, judge, and (make_sheets=True) make the
    colored view and sheets. Shared by the button and run_headless.

    Returns dict: status ("ok" or "refused"), message, graph, sizing, recirc,
    results, items, checks, system_errors, pressure, pumps, warnings and
    created (view and sheet ids made this run).
    """
    outcome = {"status": "ok", "message": "", "graph": None, "sizing": None,
               "recirc": None, "results": {}, "items": [], "checks": None,
               "system_errors": [], "pressure": None, "pumps": None,
               "warnings": [], "created": {}}

    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc, upstream)
    outcome["graph"] = graph

    # Same order as Size Water, so the return is out of the supply tree
    # before demand is counted.
    detection = water_graph.detect_return_system_types(graph)
    water_graph.apply_return_system_types(
        graph, detection["detected"] - detection["unconfirmed"])
    return_ids = set(detection["detected"])

    recirc = None
    if return_ids:
        recirc = water_recirc.analyze(
            graph,
            heat_loss_btu_per_100ft=settings["recirc_heat_loss_btu_per_100ft"],
            delta_t_f=settings["recirc_delta_t_f"])
    outcome["recirc"] = recirc

    # In memory only. write_sizes is never called.
    sizing = water_sizing_engine.size_network(
        graph, apply_minimums=True, return_system_type_ids=return_ids,
        recirc=recirc)
    outcome["sizing"] = sizing

    if not sizing["segments"]:
        output.print_md("No pipes were found on this system. Nothing was "
                        "changed.")
        outcome["status"] = "refused"
        outcome["message"] = "No pipes were found on this system."
        return outcome
    if sizing.get("flush_valve_system"):
        output.print_md("The system serves a flush valve fixture: cold water "
                        "demand uses the flush valve column of Table "
                        "E103.3(3).")

    results = {}
    if settings["velocity"]:
        for segment in sizing["segments"]:
            node = graph.nodes.get(segment.element_id)
            if node is None:
                continue
            hot = (segment.is_return or
                   node.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER)
            key = water_tables.SYSTEM_HOT if hot else water_tables.SYSTEM_COLD
            results[segment.element_id] = water_performance.classify(
                segment, key, settings["speed_tolerance"])
    outcome["results"] = results

    if settings.get("pressure_loss") or settings.get("pump_duty"):
        _read_device_drops(graph)
        _read_levels(doc, graph)
    # Each extra check fails on its own: an error here is printed and the
    # colored view and report the user asked for still get made.
    pressure = None
    if settings.get("pressure_loss"):
        try:
            pressure = water_pressure.analyze(graph, sizing,
                                              settings["start_psi"])
        except Exception as exc:
            text = "Pressure loss failed, not reported: {}".format(str(exc))
            outcome["warnings"].append(text)
            output.print_md(":warning: " + text)
    outcome["pressure"] = pressure
    pumps = None
    if settings.get("pump_duty") and recirc is not None:
        try:
            pumps = water_pump.duty_points(
                graph, recirc, settings["circuit_setter_psi"],
                settings["pump_safety_factor"])
        except Exception as exc:
            text = "Pump duty point failed, not reported: {}".format(str(exc))
            outcome["warnings"].append(text)
            output.print_md(":warning: " + text)
    outcome["pumps"] = pumps

    # The model checks and the numbered list of problem runs (one row per run,
    # the first pipe of it), then the system errors that can make the
    # results wrong (Colin, 2026-10-26).
    checks = water_checks.check_system(graph, return_ids)
    items = water_flagging.flag_items(graph, results) if results else []
    system_errors = water_system_errors.collect(
        graph, sizing, recirc, checks, items, pressure, pumps)
    outcome["checks"] = checks
    outcome["items"] = items
    outcome["system_errors"] = system_errors

    stamp = datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")
    created = outcome["created"]
    new_view = None
    if results:
        if make_sheets:
            fitting_labels = _fitting_labels(graph, results)
            new_view = _make_colored_view(doc, output, view, graph, results,
                                          fitting_labels, stamp, items)
            if new_view is not None:
                created["colored_view_id"] = revit_helpers.eid_int(new_view.Id)
        _print_report(output, results, sizing, settings["speed_tolerance"],
                      items, system_errors)
    if pressure is not None:
        _print_pressure(output, pressure, settings["start_psi"])
    if pumps:
        _print_pumps(output, pumps, settings)
    elif settings.get("pump_duty"):
        output.print_md("No hot water return was found, so there is no "
                        "recirculation pump duty point.")
    if not make_sheets:
        output.print_md("Results only (make_sheets=False): no view or sheet "
                        "was made.")
        return outcome

    extra = [t for t in [_pressure_table(pressure, settings["start_psi"])] +
             _pump_tables(pumps, settings) if t]
    # The WSFU take-off goes on the colored plan sheet, under the flagged
    # table (Colin, 2026-10-21). Only without that sheet does it get its own.
    wsfu = None
    if settings["wsfu_calcs"]:
        wsfu = (graph, sizing, {
            "date": datetime.date.today().strftime("%Y-%m-%d"),
            "job": settings["job"],
            "job_number": settings["job_number"],
            "by": settings["by"],
        }, recirc)
    if new_view is not None:
        sheet = _make_sheet(doc, output, view, new_view, results, stamp,
                            settings["speed_tolerance"], extra, wsfu, items,
                            system_errors)
        _record(created, "sheet", sheet)
    else:
        if extra:
            _record(created, "tables",
                    _make_tables_sheet(doc, output, extra, stamp))
        if wsfu is not None:
            _record(created, "wsfu", _create_wsfu_view(doc, output, *wsfu))
    return outcome


def _record(created, key, made):
    """Put the ids of what a builder made under created. made is a sheet, a
    view, or a (view, sheet) pair; None is skipped."""
    if made is None:
        return
    if isinstance(made, tuple):
        view, sheet = made
        if view is not None:
            created[key + "_view_id"] = revit_helpers.eid_int(view.Id)
        if sheet is not None:
            created[key + "_sheet_id"] = revit_helpers.eid_int(sheet.Id)
            created[key + "_sheet_number"] = sheet.SheetNumber
        return
    created[key + "_id"] = revit_helpers.eid_int(made.Id)
    number = getattr(made, "SheetNumber", None)
    if number:
        created[key + "_number"] = number


# ---------------------------------------------------------------------------
# Headless entry (pyRevit agent runtime: no dialog, no pick, no alert)
# ---------------------------------------------------------------------------

TOOL = "water_performance"
DETAILS = ("flagged", "all")


def default_settings(doc=None):
    """What the dialog prefills (checks ticked as the dialog opens: size and
    velocity on, WSFU take-off on, pressure loss and pump duty off)."""
    settings = water_run.default_settings(doc)
    settings.update({
        "velocity": True,
        "speed_tolerance": float(water_performance.SPEED_TOLERANCE),
        "wsfu_calcs": True,
        "pressure_loss": False,
        "start_psi": float(water_pressure.DEFAULT_START_PSI),
        "pump_duty": False,
        "circuit_setter_psi": float(water_pump.DEFAULT_CIRCUIT_SETTER_PSI),
        "pump_safety_factor": float(water_pump.DEFAULT_SAFETY_FACTOR),
    })
    return settings


def merge_settings(overrides, doc=None):
    """default_settings with the caller's values on top; unknown keys raise."""
    settings = default_settings(doc)
    unknown = sorted(set(overrides or {}) - set(settings))
    if unknown:
        raise ValueError("unknown Water Performance setting(s): {}".format(
            ", ".join(unknown)))
    settings.update(overrides or {})
    return settings


def _num(settings, key, problems, low=None, high=None, low_open=False):
    try:
        value = float(settings[key])
    except (TypeError, ValueError):
        problems.append("{} is not a number".format(key))
        return
    if low is not None and (value <= low if low_open else value < low):
        problems.append("{} shall be {} {}".format(
            key, "above" if low_open else "at least", low))
    if high is not None and value > high:
        problems.append("{} shall be at most {}".format(key, high))


def validate_settings(settings):
    """Problems as text, empty when the settings can run. Same limits as the
    dialog (tolerance and safety factor as fractions here)."""
    problems = water_run.validate_settings(settings)
    if not (settings.get("velocity") or settings.get("pressure_loss") or
            settings.get("pump_duty") or settings.get("wsfu_calcs")):
        problems.append("tick at least one check: velocity, pressure_loss, "
                        "pump_duty or wsfu_calcs")
    _num(settings, "speed_tolerance", problems, 0.0, 1.0)
    if settings.get("pressure_loss"):
        _num(settings, "start_psi", problems, 0.0, low_open=True)
    if settings.get("pump_duty"):
        _num(settings, "circuit_setter_psi", problems, 0.0)
        _num(settings, "pump_safety_factor", problems, 0.0, 1.0)
    return problems


def run_headless(doc, start_id, start=water_run.START_RPZ, settings=None,
                 decisions=None, make_sheets=False, view_id=None,
                 detail="flagged", include_log=False, output=None):
    """Water Performance with no UI. Returns the headless.result envelope.

    start_id / start: the RPZ ("rpz") or the tee on the existing cold main
    ("tee"). decisions: {"no_cold_connector": "continue" | "stop"}.
    make_sheets=False (default) judges and reports only, nothing is made, so
    it runs in run_query. make_sheets=True colors a copy of the plan view
    (view_id, else the active view) and makes the sheets (run_modify).
    detail "flagged" lists the pipes not green; "all" every judged pipe.
    """
    import headless
    recorder = headless.RecordingOutput(forward=output)
    state = {"prompter": None, "warnings": []}

    def done(status, data=None, errors=None, created=None, decision_exc=None):
        log = recorder.log if include_log else None
        prompter = state["prompter"]
        warnings = state["warnings"] + (list(prompter.warnings) if prompter
                                        else [])
        if decision_exc is not None:
            return headless.needs_decision(TOOL, decision_exc, warnings, log)
        return headless.result(TOOL, status, data=data, warnings=warnings,
                               errors=errors, created=created, log=log)

    try:
        state["prompter"] = headless.HeadlessPrompter(decisions)
        settings = merge_settings(settings, doc)
    except ValueError as exc:
        return done("refused", errors=[str(exc)])
    problems = validate_settings(settings)
    if start not in (water_run.START_RPZ, water_run.START_TEE):
        problems.append("start shall be 'rpz' or 'tee'")
    if detail not in DETAILS:
        problems.append("detail shall be 'flagged' or 'all'")
    if problems:
        return done("refused", errors=problems)

    view = None
    if make_sheets and settings["velocity"]:
        view = (water_run._element(doc, view_id) if view_id is not None
                else doc.ActiveView)
        stop = check_plan_view(view, settings, make_sheets)
        if stop:
            return done("refused", errors=[stop])

    recorder.print_md("# Water Performance (headless)")
    try:
        origin, upstream, stop = water_run.resolve_origin(
            doc, water_run._element(doc, start_id), start, state["prompter"],
            recorder)
        if stop == water_run._DECLINED:
            return done("cancelled")
        if stop:
            return done("refused", errors=[stop])
        outcome = perf_core(doc, settings, origin, upstream, view, recorder,
                            make_sheets=make_sheets)
    except headless.NeedsDecision as exc:
        return done("needs_decision", decision_exc=exc)
    except Exception as exc:
        return done("error", errors=[revit_helpers.exc_text(exc)])

    state["warnings"].extend(outcome["warnings"])
    if outcome["status"] != "ok":
        return done(outcome["status"], errors=[outcome["message"]])
    data = _envelope_data(outcome, origin, upstream, start, settings,
                          make_sheets, detail)
    return done("ok", data=data, created=outcome["created"])


def _envelope_data(outcome, origin, upstream, start, settings, make_sheets,
                   detail):
    """The JSON data block of the headless result."""
    results = outcome["results"]
    counts = water_performance.summarize(results.values()) if results else {}
    pipes = []
    for r in sorted(results.values(), key=lambda x: str(x.element_id)):
        if detail == "flagged" and r.label == water_performance.GREEN:
            continue
        pipes.append({
            "id": r.element_id, "label": r.label, "reason": r.reason,
            "system": r.system, "is_return": r.is_return,
            "drawn_in": r.drawn_in, "required": r.required_nominal,
            "required_in": r.required_in, "gpm": r.gpm, "fps": r.fps,
            "limit_fps": r.limit_fps, "length_ft": r.length_feet,
            "existing": r.existing,
        })
    items = outcome["items"]
    checks = outcome["checks"] or {"errors": [], "warnings": [],
                                   "ready_for_sizing": None}
    data = {
        "origin": {"id": revit_helpers.eid_int(origin.Id), "start": start,
                   "tee_reason": (upstream or {}).get("reason")},
        "make_sheets": bool(make_sheets),
        "checks_run": dict((k, bool(settings.get(k))) for k in
                           ("velocity", "pressure_loss", "pump_duty",
                            "wsfu_calcs")),
        "counts": counts,
        "pipes_listed": detail,
        "pipes": pipes,
        "attention": {"columns": list(water_flagging.CONSOLE_COLUMNS),
                      "rows": [water_flagging.row_cells(i) for i in items]},
        "system_errors": [{"severity": sev, "text": text}
                          for sev, text in outcome["system_errors"]],
        "ready_for_sizing": checks["ready_for_sizing"],
        "check_errors": [f.message for f in checks["errors"]],
        "check_warnings": [f.message for f in checks["warnings"]],
        "totals": outcome["sizing"].get("totals"),
    }
    pressure = outcome["pressure"]
    if pressure is not None:
        critical = pressure["critical"]
        data["pressure"] = {
            "start_psi": settings["start_psi"],
            "summary": water_pressure.outlet_summary(pressure),
            "critical": _critical_line(critical) if critical else None,
            "warnings": list(pressure["warnings"]),
            "columns": [c[0] for c in water_pressure.TABLE_COLUMNS],
            "failing": water_pressure.table_rows(pressure, failing_only=True),
        }
    pumps = outcome["pumps"]
    if pumps:
        data["pumps"] = {
            "columns": [c[0] for c in water_pump.TABLE_COLUMNS],
            "rows": water_pump.table_rows(pumps),
            "flow_lines": [water_pump.flow_line(p) for p in pumps],
        }
    recirc = outcome["recirc"]
    if recirc:
        data["recirc"] = {
            "total_gpm": recirc.get("total_gpm"),
            "circuit_setters": len(recirc.get("legs") or []),
        }
    return data


def _create_wsfu_view(doc, output, graph, sizing, header, recirc=None):
    """The WSFU take-off on a drafting view and a new sheet, per firm standard.

    Moved here from Size Water (Colin, 2026-10-21): the take-off is a check,
    so it lives on the check tool. Its own transaction so a failure here
    cannot roll anything else back.
    """
    transaction = Transaction(doc, "RJA Tools - Water WSFU Take-Off")
    transaction.Start()
    try:
        view, content_h, content_w = water_drafting.build_wsfu_view(
            doc, graph, header, sizing, recirc=recirc)
        if view is None:
            transaction.RollBack()
            output.print_md(
                "Could not create the drafting view. The project needs a "
                "Drafting view family type and at least one Text Note type.")
            return None

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
        return None
    return view, sheet


def _fitting_labels(graph, results):
    """Worst color of the pipes next to each fitting or accessory.

    Walks through other fittings until a pipe is reached, like the duct tool.
    A fitting with no pipe reachable is GRAY.
    """
    adjacency = {}
    for nid, node in graph.nodes.items():
        adjacency.setdefault(nid, set())
        for children in node.children_by_system.values():
            for cid in children:
                adjacency[nid].add(cid)
                adjacency.setdefault(cid, set()).add(nid)

    labels = {}
    for nid, node in graph.nodes.items():
        if node.kind not in (water_graph.KIND_FITTING,
                             water_graph.KIND_ACCESSORY):
            continue
        found = []
        seen = set([nid])
        frontier = [nid]
        while frontier and not found:
            nxt = []
            for cur in frontier:
                for neighbor in adjacency.get(cur, ()):
                    if neighbor in seen:
                        continue
                    seen.add(neighbor)
                    other = graph.nodes.get(neighbor)
                    if other is None:
                        continue
                    if neighbor in results:
                        found.append(results[neighbor].label)
                    elif other.kind in (water_graph.KIND_FITTING,
                                        water_graph.KIND_ACCESSORY):
                        nxt.append(neighbor)
            frontier = nxt
        labels[nid] = (water_performance.worst_label(found)
                       or water_performance.GRAY)
    return labels


def _symbol_names(symbol):
    """(family name, type name) of a FamilySymbol, either one "" when unread.

    Never .Name first: under IronPython it raises on Revit elements (memory
    reference_ironpython_revit_name). BuiltInParameter first, then
    ElementType.FamilyName and the static Element.Name accessor.
    """
    from Autodesk.Revit.DB import BuiltInParameter, Element
    family = ""
    type_name = ""
    try:
        param = symbol.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if param is not None:
            family = param.AsString() or ""
    except Exception:
        family = ""
    if not family:
        try:
            family = symbol.FamilyName or ""
        except Exception:
            family = ""
    try:
        param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if param is not None:
            type_name = param.AsString() or ""
    except Exception:
        type_name = ""
    if not type_name:
        try:
            type_name = Element.Name.GetValue(symbol) or ""
        except Exception:
            type_name = ""
    return family, type_name


def _keynote_symbol(doc, output=None):
    """The RJA numbered circle family symbol, or None.

    Every symbol whose names could not be read is counted, and the result is
    printed (Colin, 2026-11-02: no silent failures).
    """
    unread = []
    found = None
    for symbol in FilteredElementCollector(doc).OfClass(
            FamilySymbol).ToElements():
        family, type_name = _symbol_names(symbol)
        if not family:
            unread.append(revit_helpers.eid_int(symbol.Id))
            continue
        if "RJA - Keynote Symbol" in family and (
                "Circle" in family or "Circle" in type_name):
            found = symbol
            break
    if output is not None:
        if unread:
            output.print_md("Keynote symbol search: family name NOT FOUND on "
                            "{} symbol(s), first ids {}.".format(
                                len(unread),
                                ", ".join(str(i) for i in unread[:5])))
        if found is None:
            output.print_md("Keynote symbol 'RJA - Keynote Symbol' (Circle) "
                            "NOT FOUND in the project, so the flagged pipes "
                            "are not numbered on the plan.")
        else:
            output.print_md("Keynote symbol found: element {}.".format(
                revit_helpers.eid_int(found.Id)))
    return found


def _on_view_level(view, element):
    """True when the pipe sits on the level the plan view shows."""
    try:
        level = view.GenLevel
        return (level is None or element.ReferenceLevel is None or
                element.ReferenceLevel.Id == level.Id)
    except Exception:
        return True


def _make_colored_view(doc, output, view, graph, results, fitting_labels,
                       stamp, items=None):
    """Copy the active plan view and color the system in the copy.

    Returns the new view, or None when it could not be made.
    """
    fill_id = hvac_graph.solid_fill_pattern_id(doc)

    transaction = Transaction(doc, "RJA Tools - Water Performance")
    transaction.Start()
    try:
        new_view = doc.GetElement(
            view.Duplicate(ViewDuplicateOption.Duplicate))
        name = "Water Performance - {}".format(stamp)
        try:
            new_view.Name = name
        except Exception:
            new_view.Name = name + " (2)"

        colored = 0
        failed = 0
        everything = {}
        for eid, result in results.items():
            everything[eid] = result.label
        everything.update(fitting_labels)

        for eid, label in everything.items():
            node = graph.nodes.get(eid)
            if node is None or node.element is None:
                continue
            color = _COLORS[label]
            ogs = OverrideGraphicSettings()
            ogs.SetProjectionLineColor(color)
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            try:
                new_view.SetElementOverrides(node.element.Id, ogs)
                colored += 1
            except Exception as exc:
                failed += 1
                output.print_md("Could not color element {}: {}".format(
                    eid, str(exc)))
        # Numbered circles on the first pipe of each problem run, so row 3 of
        # the table is circle 3 in the plan.
        symbol = _keynote_symbol(doc, output)
        if symbol is not None and not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()
        for item in items or []:
            node = graph.nodes.get(item.result.element_id)
            if node is None or node.element is None:
                continue
            if not _on_view_level(view, node.element):
                continue
            try:
                point = node.element.Location.Curve.Evaluate(0.5, True)
                if symbol is not None:
                    instance = doc.Create.NewFamilyInstance(
                        point, symbol, new_view)
                    label = instance.LookupParameter("Label")
                    if label is not None and not label.IsReadOnly:
                        if label.StorageType == StorageType.String:
                            label.Set(str(item.number))
                        elif label.StorageType == StorageType.Integer:
                            label.Set(item.number)
                    else:
                        output.print_md(
                            "Keynote {} at pipe {}: 'Label' parameter NOT "
                            "FOUND or read only, the circle is blank.".format(
                                item.number, item.result.element_id))
            except Exception as exc:
                output.print_md("Keynote {} could not be placed at pipe {}: "
                                "{}".format(item.number,
                                            item.result.element_id, str(exc)))
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        output.print_md("Colored view failed, nothing was added to the "
                        "model: {}".format(str(exc)))
        return None

    output.print_md("Created view **{}**. Colored {} element(s), {} "
                    "failed.".format(name, colored, failed))
    return new_view


def _legend(tolerance):
    """Legend wording, worst first, matching the colors in the plan."""
    return [
        (water_performance.RED,
         "Red - drawn smaller than the firm table size, or load above "
         "the table (size by hand)"),
        (water_performance.YELLOW,
         "Yellow - at table size, computed speed more than {:.0f} percent "
         "over the limit".format(float(tolerance) * 100.0)),
        (water_performance.PURPLE,
         "Purple - drawn larger than the firm table size"),
        (water_performance.GREEN, "Green - at table size"),
        (water_performance.GRAY, "Gray - not judged (see Reason)"),
    ]

def _read_device_drops(graph):
    """Device_Pressure_Drop_PSI from every non-pipe element, instance first,
    then Type, onto node.device_drop_psi for water_pressure."""
    for node in graph.nodes.values():
        if node.kind in (water_graph.KIND_PIPE, water_graph.KIND_FIXTURE,
                         water_graph.KIND_ORIGIN) or node.element is None:
            continue
        value = revit_helpers.get_parameter_value(
            node.element, shared_params.PARAM_DEVICE_PRESSURE_DROP)
        if value is None:
            value = revit_helpers.get_type_parameter_value(
                node.element, shared_params.PARAM_DEVICE_PRESSURE_DROP)
        try:
            node.device_drop_psi = float(value) if value else None
        except (TypeError, ValueError):
            node.device_drop_psi = None


def _read_levels(doc, graph):
    """Level elevation of the start point and every fixture, onto
    node.level_elev_ft, so static head is measured floor to floor."""
    wanted = set(graph.fixture_ids)
    wanted.add(graph.origin_id)
    for nid in wanted:
        node = graph.nodes.get(nid)
        if node is None or node.element is None:
            continue
        try:
            level = doc.GetElement(node.element.LevelId)
            node.level_elev_ft = float(level.Elevation) if level else None
        except Exception:
            node.level_elev_ft = None


def _critical_line(critical):
    return ("Lowest pressure: {} ({} water), {:.1f} psi at the outlet "
            "against a {:.0f} psi minimum.".format(
                critical.fixture_name,
                "hot" if critical.system == water_pressure.HOT else "cold",
                critical.residual_psi, critical.required_psi))


def _print_pressure(output, pressure, start_psi):
    output.print_md("## Pressure at fixture outlets")
    critical = pressure["critical"]
    if critical is not None:
        output.print_md(_critical_line(critical))
    for warning in pressure["warnings"]:
        output.print_md(":warning: " + warning)
    output.print_md(water_pressure.outlet_summary(pressure))
    failing = water_pressure.table_rows(pressure, failing_only=True)
    if failing:
        output.print_table(
            table_data=failing,
            title="Fixture outlets below their minimum, lowest first",
            columns=[c[0] for c in water_pressure.TABLE_COLUMNS])
    notes = sorted(set(n for o in pressure["outlets"] for n in o.notes))
    if notes:
        output.print_md("Not counted or assumed:\n" +
                        "\n".join("- " + n for n in _group_notes(notes)))
    output.print_md(chr(10).join(
        "    " + line for line in water_pressure.basis_lines(start_psi)))


_NOTE_ID = re.compile(r"^(.*?) (\d{5,}): (.*)$")


def _group_notes(notes, show=4):
    """'90 elbow 123: reason' x40 becomes one line: reason, 40 (ids)."""
    groups = {}
    order = []
    plain = []
    for note in notes:
        match = _NOTE_ID.match(note)
        if match is None:
            plain.append(note)
            continue
        key = (match.group(1), match.group(3))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(match.group(2))
    lines = []
    for key in order:
        ids = groups[key]
        if len(ids) == 1:
            lines.append("{} {}: {}".format(key[0], ids[0], key[1]))
            continue
        text = ", ".join(ids[:show])
        if len(ids) > show:
            text += ", +{} more".format(len(ids) - show)
        lines.append("{} x {}: {} ({})".format(len(ids), key[0], key[1], text))
    return lines + plain


def _pressure_table(pressure, start_psi):
    """The pressure section for the sheet, or None."""
    if pressure is None:
        return None
    notes = list(pressure["warnings"])
    critical = pressure["critical"]
    if critical is not None:
        notes.insert(0, _critical_line(critical))
    notes.insert(0, water_pressure.outlet_summary(pressure))
    notes.extend(water_pressure.basis_lines(start_psi))
    # Only the outlets that do not pass (Colin, 2026-10-21). With none, the
    # summary line above the basis says so and no table is drawn.
    return ("FIXTURE OUTLETS BELOW MINIMUM PRESSURE - 2024 IPC E103.3, "
            "TABLE 604.3", water_pressure.TABLE_COLUMNS,
            water_pressure.table_rows(pressure, failing_only=True), notes)


def _print_pumps(output, pumps, settings):
    output.print_md("## Recirculation pump duty point")
    output.print_table(
        table_data=water_pump.table_rows(pumps),
        title="Pump duty point",
        columns=[c[0].title() for c in water_pump.TABLE_COLUMNS])
    for p in pumps:
        output.print_md(water_pump.flow_line(p))
        rows, equations = water_pump.head_rows(p)
        if rows:
            output.print_table(
                table_data=rows,
                title="System head, heater {}".format(
                    p.get("heater_id", p["root_id"])),
                columns=[c[0].title() for c in water_pump.HEAD_COLUMNS])
            output.print_md(chr(10).join(
                "    " + line for line in equations))
    notes = sorted(set(n for p in pumps for c in p["circuits"]
                       for n in c["notes"]))
    if notes:
        output.print_md("Not counted:" + chr(10) + chr(10).join(
            "- " + n for n in _group_notes(notes)))


def _pump_tables(pumps, settings):
    """The pump sections for the sheet: the duty point table with its flow
    line, then the system head chart for each heater."""
    if not pumps:
        return []
    notes = [water_pump.flow_line(p) for p in pumps]
    tables = [("RECIRCULATION PUMP DUTY POINT", water_pump.TABLE_COLUMNS,
               water_pump.table_rows(pumps), notes)]
    for p in pumps:
        rows, equations = water_pump.head_rows(p)
        if rows:
            tables.append(("SYSTEM HEAD, HEATER {}".format(
                p.get("heater_id", p["root_id"])), water_pump.HEAD_COLUMNS,
                rows, equations))
    return tables


def _make_tables_sheet(doc, output, tables, stamp):
    """Pressure loss and pump tables on their own sheet, when the size and
    velocity check (and so the colored plan) was not run."""
    transaction = Transaction(doc, "RJA Tools - Water Performance Sheet")
    transaction.Start()
    try:
        view, height, width = water_drafting.build_tables_view(
            doc, "Water Performance Tables - {}".format(stamp),
            "WATER PERFORMANCE - BASIS 2024 IPC", tables)
        if view is None:
            raise Exception("no drafting view or text type in the project")
        sheet, _viewport = water_drafting.place_on_sheet(
            doc, view, height, width, "WP-{}".format(stamp),
            "Water Performance Tables (Sheet)")
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        output.print_md("Sheet failed, the results above are printed only: "
                        "{}".format(str(exc)))
        return None
    if sheet is not None:
        output.print_md("Created sheet **{}**.".format(sheet.SheetNumber))
    else:
        output.print_md("Created drafting view **Water Performance Tables - "
                        "{}** (no title block for a sheet).".format(stamp))
    return view, sheet


def _make_sheet(doc, output, source_view, plan_view, results, stamp,
                tolerance, extra_tables=None, wsfu=None, items=None,
                system_errors=None):
    """Put the colored plan, the flagged table and the legend on a new sheet.

    wsfu, when given, is (graph, sizing, header, recirc) for the WSFU take-off,
    placed on the same sheet under the flagged table.

    Its own transaction, so a missing title block or text type costs the
    sheet and never the colored view already committed.
    """
    counts = water_performance.summarize(results.values())
    summary = [
        "WATER PERFORMANCE - BASIS 2024 IPC, FIRM WSFU TABLE",
        "Green {}   Yellow {}   Red {}   Purple {}   Gray {}".format(
            counts[water_performance.GREEN], counts[water_performance.YELLOW],
            counts[water_performance.RED], counts[water_performance.PURPLE],
            counts[water_performance.GRAY]),
    ]
    rows = [water_flagging.row_cells(item) for item in items or []]
    legend = [(_COLORS[key], text) for key, text in _legend(tolerance)]

    source_sheet = revit_helpers.sheet_hosting_view(doc, source_view)
    source_number = source_sheet.SheetNumber if source_sheet else "NA"

    transaction = Transaction(doc, "RJA Tools - Water Performance Sheet")
    transaction.Start()
    try:
        fill_id = hvac_graph.solid_fill_pattern_id(doc)
        # One drafting view: tables, legend and (optionally) the WSFU take-off.
        wsfu_arg = None
        if wsfu is not None:
            graph, sizing, header, recirc = wsfu
            wsfu_arg = (graph, header, sizing, recirc)
        table_only_h = []
        table_view, table_h, table_w = water_drafting.build_performance_view(
            doc, "Water Performance Table - {}".format(stamp), summary,
            water_flagging.COLUMNS, rows, legend, fill_id,
            extra_tables or None, wsfu_arg, table_only_h,
            water_system_errors.lines(system_errors or []))
        wsfu_view = table_view if wsfu is not None else None
        sheet = water_drafting.place_plan_and_table(
            doc, plan_view, table_view, table_h, table_w,
            "WP-{}-{}".format(source_number, stamp),
            "Water Performance - {} (Sheet)".format(source_number),
            None, table_only_h[0] if table_only_h else None)
        if sheet is None:
            transaction.RollBack()
            output.print_md("No title block is loaded, so no sheet was made. "
                            "The colored view is kept.")
            return None
        err = revit_helpers.file_like_source_sheet(doc, sheet, source_sheet)
        if err:
            output.print_md("Could not file the sheet in the sheet browser: "
                            "{}".format(err))
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        output.print_md("Sheet failed, the colored view is kept: {}".format(
            str(exc)))
        return None
    output.print_md("Created sheet **{}**{}.".format(
        sheet.SheetNumber,
        ", WSFU take-off included" if wsfu is not None and wsfu_view else ""))
    return sheet


def _print_report(output, results, sizing, tolerance, items=None,
                  system_errors=None):
    counts = water_performance.summarize(results.values())
    output.print_md("## Result")
    output.print_md(
        "- GREEN (at table size): **{}**\n"
        "- YELLOW (at table size, computed speed over limit): **{}**\n"
        "- RED (drawn smaller than table, or load above it): **{}**\n"
        "- PURPLE (drawn larger than table): **{}**\n"
        "- GRAY (not judged): **{}**".format(
            counts[water_performance.GREEN],
            counts[water_performance.YELLOW],
            counts[water_performance.RED],
            counts[water_performance.PURPLE],
            counts[water_performance.GRAY]))

    entries = system_errors or []
    output.print_md("## System errors")
    output.print_md(water_system_errors.summary_line(entries))
    if entries:
        output.print_md(chr(10).join(
            "- **{}**: {}".format(sev, text) for sev, text in entries))

    if items:
        output.print_table(
            table_data=[water_flagging.row_cells(i) for i in items],
            title="Pipes needing attention",
            columns=water_flagging.CONSOLE_COLUMNS)
    else:
        output.print_md("Every judged pipe is at table size.")

    output.print_md(
        "    Velocity (fps) = gpm / (2.4483 x ID^2), ID in inches, ASTM B88 "
        "Type L, at the drawn size\n"
        "    Limits: 8 fps cold, 5 fps hot and return. Yellow when V > "
        "limit x (1 + {:.2f}).\n"
        "    Pass: drawn size = firm table size for the pipe's demand "
        "(Table E103.3(3) flow).".format(float(tolerance)))
