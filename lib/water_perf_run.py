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
    if not isinstance(view, ViewPlan) or view.IsTemplate:
        forms.alert("Open the floor plan to check, then run again. The result "
                    "is a colored copy of the active plan view.",
                    title="Water Performance")
        return False

    settings = ui_helpers.show_water_dialog(
        "Water Performance - 2024 IPC", water_run.project_info(doc, output),
        None)
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
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

    results = {}
    for segment in sizing["segments"]:
        node = graph.nodes.get(segment.element_id)
        if node is None:
            continue
        hot = (segment.is_return or
               node.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER)
        key = water_tables.SYSTEM_HOT if hot else water_tables.SYSTEM_COLD
        results[segment.element_id] = water_performance.classify(segment, key)

    if not results:
        output.print_md("No pipes were found on this system. Nothing was "
                        "changed.")
        return False

    fitting_labels = _fitting_labels(graph, results)
    _make_colored_view(doc, output, view, graph, results, fitting_labels)
    _print_report(output, results, sizing)
    return True


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


def _make_colored_view(doc, output, view, graph, results, fitting_labels):
    """Copy the active plan view and color the system in the copy."""
    fill_id = hvac_graph.solid_fill_pattern_id(doc)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")

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
        return

    output.print_md("Created view **{}**. Colored {} element(s), {} "
                    "failed.".format(name, colored, failed))


def _print_report(output, results, sizing):
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
        "    " + line for line in water_performance.basis_lines()))


def _fmt_size(value):
    if value is None:
        return "-"
    return "{:.3f}".format(float(value)).rstrip("0").rstrip(".")
