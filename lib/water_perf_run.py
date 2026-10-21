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

from Autodesk.Revit.DB import (
    Color, ElementId, OverrideGraphicSettings, Transaction, ViewDuplicateOption,
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
    """Run the tool. Returns True if it finished."""
    output.print_md("# Water Performance")

    view = doc.ActiveView
    settings = ui_helpers.show_water_performance_dialog(
        "Water Performance - 2024 IPC", water_run.project_info(doc, output))
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False
    # Only the size and velocity check colors a copy of the plan.
    if settings["velocity"] and (not isinstance(view, ViewPlan) or
                                 view.IsTemplate):
        forms.alert("Open the floor plan to check, then run again. Pipe size "
                    "and velocity colors a copy of the active plan view.",
                    title="Water Performance")
        return False

    origin, upstream = water_run.pick_origin(uidoc, doc, output, forms)
    if origin is None:
        return False

    revit_helpers.reset_pipe_diameter_approach()
    graph = water_graph.build_water_network(origin, doc, upstream)

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

    # In memory only. write_sizes is never called.
    sizing = water_sizing_engine.size_network(
        graph, apply_minimums=True, return_system_type_ids=return_ids,
        recirc=recirc)

    if not sizing["segments"]:
        output.print_md("No pipes were found on this system. Nothing was "
                        "changed.")
        return False
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
            output.print_md(":warning: Pressure loss failed, not reported: "
                            "{}".format(str(exc)))
    pumps = None
    if settings.get("pump_duty") and recirc is not None:
        try:
            pumps = water_pump.duty_points(
                graph, recirc, settings["circuit_setter_psi"],
                settings["pump_safety_factor"])
        except Exception as exc:
            output.print_md(":warning: Pump duty point failed, not reported: "
                            "{}".format(str(exc)))

    stamp = datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")
    new_view = None
    if results:
        fitting_labels = _fitting_labels(graph, results)
        new_view = _make_colored_view(doc, output, view, graph, results,
                                      fitting_labels, stamp)
        _print_report(output, results, sizing, settings["speed_tolerance"])
    if pressure is not None:
        _print_pressure(output, pressure, settings["start_psi"])
    if pumps:
        _print_pumps(output, pumps, settings)
    elif settings.get("pump_duty"):
        output.print_md("No hot water return was found, so there is no "
                        "recirculation pump duty point.")
    extra = [t for t in (_pressure_table(pressure, settings["start_psi"]),
                         _pump_table(pumps, settings)) if t]
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
        _make_sheet(doc, output, view, new_view, results, stamp,
                    settings["speed_tolerance"], extra, wsfu)
    else:
        if extra:
            _make_tables_sheet(doc, output, extra, stamp)
        if wsfu is not None:
            _create_wsfu_view(doc, output, *wsfu)
    return True


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


def _make_colored_view(doc, output, view, graph, results, fitting_labels,
                       stamp):
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
         "Red - drawn smaller than the firm table size"),
        (water_performance.YELLOW,
         "Yellow - at table size, computed speed more than {:.0f} percent "
         "over the limit".format(float(tolerance) * 100.0)),
        (water_performance.PURPLE,
         "Purple - drawn larger than the firm table size"),
        (water_performance.GREEN, "Green - at table size"),
        (water_performance.GRAY, "Gray - not judged (see Reason)"),
    ]

# (header, width ft) for the sheet table, same cells as the pyRevit table.
_TABLE_COLUMNS = [("STATUS", 0.075), ("PIPE", 0.075), ("SYSTEM", 0.065),
                  ("DRAWN", 0.055), ("TABLE", 0.060), ("GPM", 0.050),
                  ("FPS", 0.045), ("REASON", 0.560)]


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
                        "\n".join("- " + n for n in notes[:40]))
    output.print_md("\n".join(
        "    " + line for line in water_pressure.basis_lines(start_psi)))


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
    for p in pumps:
        output.print_md(
            "Heater **{}**: **{:.1f} gpm at {:.1f} ft head** (index circuit "
            "leg {}).".format(p["root_id"], p["gpm"], p["head_ft"],
                              p["index_leg"]))
    output.print_table(
        table_data=water_pump.table_rows(pumps),
        title="Pump duty point, index circuit breakdown",
        columns=[c[0].title() for c in water_pump.TABLE_COLUMNS])
    notes = sorted(set(n for p in pumps for c in p["circuits"]
                       for n in c["notes"]))
    if notes:
        output.print_md("Not counted:\n" + "\n".join("- " + n for n in notes))
    output.print_md("\n".join("    " + line for line in water_pump.basis_lines(
        settings["circuit_setter_psi"], settings["pump_safety_factor"])))


def _pump_table(pumps, settings):
    """The pump section for the sheet, or None."""
    if not pumps:
        return None
    notes = ["HEATER {}: {:.1f} GPM AT {:.1f} FT HEAD".format(
        p["root_id"], p["gpm"], p["head_ft"]) for p in pumps]
    notes.extend(water_pump.basis_lines(settings["circuit_setter_psi"],
                                        settings["pump_safety_factor"]))
    return ("RECIRCULATION PUMP DUTY POINT", water_pump.TABLE_COLUMNS,
            water_pump.table_rows(pumps), notes)


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
        return
    if sheet is not None:
        output.print_md("Created sheet **{}**.".format(sheet.SheetNumber))
    else:
        output.print_md("Created drafting view **Water Performance Tables - "
                        "{}** (no title block for a sheet).".format(stamp))


def _make_sheet(doc, output, source_view, plan_view, results, stamp,
                tolerance, extra_tables=None, wsfu=None):
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
    rows = [[str(c) for c in row] for row in _flagged_rows(results)]
    legend = [(_COLORS[key], text) for key, text in _legend(tolerance)]

    source_sheet = revit_helpers.sheet_hosting_view(doc, source_view)
    source_number = source_sheet.SheetNumber if source_sheet else "NA"

    transaction = Transaction(doc, "RJA Tools - Water Performance Sheet")
    transaction.Start()
    try:
        fill_id = hvac_graph.solid_fill_pattern_id(doc)
        table_view, table_h, table_w = water_drafting.build_performance_view(
            doc, "Water Performance Table - {}".format(stamp), summary,
            _TABLE_COLUMNS, rows, legend, fill_id, extra_tables or None)
        wsfu_view = None
        if wsfu is not None:
            graph, sizing, header, recirc = wsfu
            wsfu_view, _h, _w = water_drafting.build_wsfu_view(
                doc, graph, header, sizing, recirc=recirc)
            if wsfu_view is None:
                output.print_md("The WSFU take-off could not be drawn: the "
                                "project needs a Drafting view family type "
                                "and a Text Note type.")
        sheet = water_drafting.place_plan_and_table(
            doc, plan_view, table_view, table_h, table_w,
            "WP-{}-{}".format(source_number, stamp),
            "Water Performance - {} (Sheet)".format(source_number),
            wsfu_view)
        if sheet is None:
            transaction.RollBack()
            output.print_md("No title block is loaded, so no sheet was made. "
                            "The colored view is kept.")
            return
        err = revit_helpers.file_like_source_sheet(doc, sheet, source_sheet)
        if err:
            output.print_md("Could not file the sheet in the sheet browser: "
                            "{}".format(err))
        transaction.Commit()
    except Exception as exc:
        transaction.RollBack()
        output.print_md("Sheet failed, the colored view is kept: {}".format(
            str(exc)))
        return
    output.print_md("Created sheet **{}**{}.".format(
        sheet.SheetNumber,
        ", WSFU take-off included" if wsfu is not None and wsfu_view else ""))


def _print_report(output, results, sizing, tolerance):
    counts = water_performance.summarize(results.values())
    output.print_md("## Result")
    output.print_md(
        "- GREEN (at table size): **{}**\n"
        "- YELLOW (at table size, computed speed over limit): **{}**\n"
        "- RED (drawn smaller than table): **{}**\n"
        "- PURPLE (drawn larger than table): **{}**\n"
        "- GRAY (not judged): **{}**".format(
            counts[water_performance.GREEN],
            counts[water_performance.YELLOW],
            counts[water_performance.RED],
            counts[water_performance.PURPLE],
            counts[water_performance.GRAY]))

    rows = _flagged_rows(results)
    if rows:
        output.print_table(
            table_data=rows,
            title="Pipes needing attention",
            columns=["Status", "Pipe", "System", "Drawn (in)", "Table size",
                     "gpm", "fps", "Reason"])
    else:
        output.print_md("Every judged pipe is at table size.")

    output.print_md("## Basis of design")
    output.print_md("\n".join(
        "    " + line for line in water_performance.basis_lines(tolerance)))


def _flagged_rows(results):
    """Every pipe that is not green, worst first. One source for the pyRevit
    table and the sheet table."""
    rows = []
    order = (water_performance.RED, water_performance.YELLOW,
             water_performance.PURPLE, water_performance.GRAY)
    flagged = [r for r in results.values() if r.label in order]
    # Graph ids are already plain integers (keyed by eid_int).
    flagged.sort(key=lambda r: (-water_performance.PRIORITY[r.label],
                                r.element_id))
    for r in flagged:
        rows.append([
            r.label,
            r.element_id,
            "Return" if r.is_return else (
                "Hot" if r.system == water_tables.SYSTEM_HOT else "Cold"),
            _fmt_size(r.drawn_in),
            r.required_nominal if r.required_nominal else "-",
            "{:.1f}".format(r.gpm) if r.gpm is not None else "-",
            "{:.1f}".format(r.fps) if r.fps is not None else "-",
            r.reason,
        ])
    return rows


def _fmt_size(value):
    if value is None:
        return "-"
    return "{:.3f}".format(float(value)).rstrip("0").rstrip(".")
