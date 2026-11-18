# -*- coding: ascii -*-
# One-Line.pushbutton/script.py
# Gas piping one-line diagram (rebuilt 2026-11-18, plan:
# FOR CLAUDE\Staging\Gas Test Benches\ONELINE_REBUILD_PLAN.md).
#
# A pyRevit button only, no LLM or MCP at runtime (Colin). Thin on purpose:
#   pick -> walk (pipe_graph) -> One-Line dialog -> Size Gas rerun in memory
#   -> snapshot -> gate (complete and sized, else "Run Size Gas first")
#   -> schematic -> layout -> draw (one undo) -> themed console.
# Every decision lives in lib: oneline_snapshot, oneline_gate,
# oneline_schematic, oneline_layout, oneline_draw, gas_labels, gas_console.
#
# IronPython 2.7 / pyRevit

import os
import sys
import datetime
import traceback

from pyrevit import script
from Autodesk.Revit.DB import Transaction, TransactionGroup
from Autodesk.Revit.UI.Selection import ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()
logger = script.get_logger()

_lib_dir = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

import console_report as cr        # noqa: E402
import gas_console                 # noqa: E402
import gas_labels                  # noqa: E402
import gas_tables                  # noqa: E402
import oneline_draw                # noqa: E402
import oneline_gate                # noqa: E402
import oneline_layout              # noqa: E402
import oneline_schematic           # noqa: E402
import oneline_snapshot            # noqa: E402
import pipe_graph                  # noqa: E402
import revit_helpers               # noqa: E402
import sizing_engine               # noqa: E402
import tie_in                      # noqa: E402
import ui_helpers                  # noqa: E402

TITLE = "Gas One-Line"


def _alert(text, kind="error"):
    ui_helpers.rja_alert(TITLE, text, kind)


def _emit(html, md):
    cr.emit(output, html, lambda: output.print_md(md), logger, TITLE)


def _pick():
    """The meter, or the first new gas pipe at an existing building.
    Returns (element, upstream or None) or (None, None)."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, "Select the gas meter, or the first new gas "
            "pipe at the connection to existing")
    except Exception:
        output.print_md("Selection cancelled. Nothing drawn.")
        return None, None
    el = doc.GetElement(ref.ElementId)
    if el is None:
        _alert("Could not read the picked element.")
        return None, None
    if revit_helpers.is_pipe(el):
        chk = pipe_graph.check_gas_pipe(el)
        if not chk["ok"]:
            _alert("This is not a gas pipe: {0}.\n\nPick the gas meter, or "
                   "the first new GAS pipe at the connection to "
                   "existing.".format(chk["reason"]))
            return None, None
        upstream = tie_in.resolve_new_pipe_upstream(el, doc)
        if upstream["status"] != "ok":
            _alert("Could not tell which end of the picked pipe connects to "
                   "the existing gas piping: {0}.".format(upstream["reason"]))
            return None, None
        return el, upstream
    val = revit_helpers.validate_selected_element(el)
    if not val["is_valid"]:
        _alert("Pick the gas meter, or the first new gas pipe at the "
               "connection to existing.\n\n{0}".format(val["reason"]))
        return None, None
    return el, None


def _table(material, label):
    if not label:
        return None
    return gas_tables.get_table_option_by_material_and_short_label(
        material, label)


def main():
    revit_helpers.clear_log()
    el, upstream = _pick()
    if el is None:
        return
    is_tie_in = upstream is not None
    el_id = revit_helpers.eid_int(el.Id)

    # walk
    try:
        graph = pipe_graph.build_network(el, doc, upstream)
    except Exception as ex:
        _alert("The piping walk failed:\n\n{0}".format(ex))
        logger.error(traceback.format_exc())
        return

    # dialog: the downstream table only when a mid stream PRV was found
    choice = ui_helpers.show_oneline_picker(
        "Gas One-Line - Table", bool(graph.midstream_prv_ids))
    if not choice:
        output.print_md("Cancelled. Nothing drawn.")
        return
    opt = _table(choice["pipe_material"], choice["table_label"])
    table_id = opt["table_id"]
    down = _table(choice["pipe_material"], choice.get("downstream_table_label"))
    down_id = down["table_id"] if down else None
    heat = choice["heat_content"]

    existing = None
    if is_tie_in:
        existing = ui_helpers.show_existing_gas_dialog(
            "Gas One-Line - Existing Conditions",
            gas_tables.list_pipe_sizes(table_id), tie_in.recall_existing(el_id))
        if existing is None:
            output.print_md("Cancelled at existing conditions. Nothing drawn.")
            return
        tie_in.remember_existing(el_id, existing)
        if graph.longest_run is not None:
            pipe_graph.apply_existing_conditions(
                graph, existing["size_nominal"], existing["length_ft"],
                existing["load_mbh"])

    info = {"meter": "{0} ({1})".format(
                "New pipe at existing" if is_tie_in else "Gas meter", el_id),
            "table": opt["label"].split("  [")[0], "table_id": table_id,
            "heat_content": heat,
            "downstream_table": down["label"].split("  [")[0] if down else None,
            "origin_kind": "tie_in" if is_tie_in else "meter"}
    _emit(gas_console.oneline_inputs(info), gas_console.oneline_inputs_md(info))

    # Size Gas rerun in memory (writes nothing): the gate compares sizes
    result, sizing_error = None, None
    if graph.longest_run is not None:
        try:
            result = sizing_engine.size_system(
                graph, choice["pipe_material"], opt["inlet_pressure_psi"],
                table_id, heat_content_btu_per_cf=heat,
                downstream_table_id=down_id)
        except ValueError as ex:
            sizing_error = str(ex)
    else:
        sizing_error = "no longest run (no gas fixture reached)"

    snap = oneline_snapshot.snapshot(graph, doc, table_id,
                                     sizing_result=result,
                                     downstream_table_id=down_id)
    snap["sizing_error"] = sizing_error
    if result is not None and is_tie_in:
        snap["existing_check"] = sizing_engine.check_existing_pipe(result, graph)

    res = oneline_gate.check(snap)
    _emit(gas_console.gate(res), gas_console.gate_md(res))
    if not res["ok"]:
        _alert("\n".join(oneline_gate.summary_lines(res, limit=12)), "error")
        return

    # schematic and layout (pure)
    sch = oneline_schematic.build(snap)
    terms = oneline_schematic.terminals_in_order(sch)
    new_mbh = sum(float(t["mbh"] or 0.0) for t in terms)
    notes = gas_labels.notes_lines(
        table_id, new_mbh, snap["longest_run_ft"], downstream_table_id=down_id,
        existing_mbh=existing["load_mbh"] if existing else None)
    lay = oneline_layout.layout(sch, notes=notes)
    if any(l.startswith("LAYOUT FAILED") for l in lay["log"]):
        _emit(gas_console.layout_diagnostic(sch, lay),
              gas_console.layout_diagnostic_text(sch, lay))
        _alert("The diagram could not be laid out without overlaps. Nothing "
               "was drawn. The layout diagnostic is in the console.")
        return

    # draw: one undo
    view_name = "Gas Piping One-Line - {0}".format(
        datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S"))
    tg = TransactionGroup(doc, "RJA Tools - Gas One-Line Diagram")
    tg.Start()
    try:
        t1 = Transaction(doc, "One-Line symbols")
        t1.Start()
        syms, sym_rows = oneline_draw.ensure_symbols(doc)
        t1.Commit()
        t2 = Transaction(doc, "One-Line drafting view")
        t2.Start()
        view = oneline_draw.create_view(doc, view_name)
        counts, draw_errors = oneline_draw.draw(doc, view, lay["prims"], syms)
        t2.Commit()
        tg.Assimilate()
    except Exception as ex:
        try:
            tg.RollBack()
        except Exception:
            pass
        logger.error(traceback.format_exc())
        _alert("Drawing failed, nothing was kept:\n\n{0}".format(ex))
        return

    # console
    _emit(gas_console.symbols_table(sym_rows),
          gas_console.symbols_table_md(sym_rows))
    _emit(gas_console.oneline_summary(sch, snap, view_name),
          gas_console.oneline_summary_md(sch, snap, view_name))
    _emit(gas_console.length_check(sch, snap), "")
    warns = list(sch["warnings"]) + [w["text"] for w in res["warnings"]]
    warns += draw_errors
    if counts.get("failed"):
        warns.append("{0} item(s) could not be drawn.".format(counts["failed"]))
    if warns:
        _emit(gas_console.oneline_warnings(warns),
              gas_console.oneline_warnings_md(warns))
    _emit(gas_console.layout_diagnostic(sch, lay),
          gas_console.layout_diagnostic_text(sch, lay))

    try:
        uidoc.ActiveView = view
    except Exception as ex:
        logger.warning("Could not open the new view: {0}".format(ex))


main()
