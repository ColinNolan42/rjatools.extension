# -*- coding: ascii -*-
# water_drafting.py
# Places the WSFU take-off and the Water Performance tables into Revit
# drafting views, and onto sheets. The drawing itself (bands, styled tables,
# legend swatches) is lib/rja_drafting.py, shared with Duct Performance.
#
# The columns, headers and cell values come from water_report.WSFU_COLUMNS and
# water_report.wsfu_data(). Nothing about the table is defined twice - the
# console table and this schedule cannot disagree.
#
# This is the only water module that imports the Revit API, so the rest stay
# testable in plain CPython.
#
# IronPython 2.7

from Autodesk.Revit.DB import XYZ, ViewSheet, Viewport

# The RJA table look lives in rja_drafting, shared with Duct Performance.
from rja_drafting import (
    PAD, ROW_H, TEXT_H, GAP, BAND_H, SUB_BAND_H,
    _Style, _t, _band, _info_grid, _note_panel, _outline, _styled_table,
    _legend_swatch, _wrap, _new_drafting_view, _title_block_type_id,
    _hide_viewport_title)
import water_report
import water_tables


CHART_WIDTHS = (0.090, 0.160, 0.150, 0.120)   # firm pipe sizing chart

# Drafting view versions of the take-off headers, two rows like the firm
# WFUTO template: a group row over the per fixture and extended columns.
_WSFU_GROUPS = [("", 2), ("PER FIXTURE (WSFU)", 3), ("EXTENDED (WSFU)", 3)]
_WSFU_HEADERS = {"fixture": "FIXTURE TYPE**", "qty": "QUANTITY",
                 "cold_ea": "COLD", "hot_ea": "HOT", "total_ea": "TOTAL",
                 "cold_ext": "COLD", "hot_ext": "HOT", "total_ext": "TOTAL"}

# Footnotes from the firm template (Template_WSFU CALCS.xlsx, WFUTO C60:C61).
WSFU_FOOTNOTES = [
    "* For fixtures not listed, loads should be assumed by comparing the "
    "fixture to one listed using similar quantities of water and similar "
    "rates. The assigned loads for fixtures with both hot and cold water "
    "supplies are given for separate hot and cold water loads and for total "
    "load, the separate hot and cold water loads being three-fourths of the "
    "total load for the fixture in each case.",
    "** PUBLIC = COMMERCIAL BUILDINGS.  PRIVATE = RESIDENTIAL BUILDINGS.",
]


def build_wsfu_view(doc, graph, header, sizing_result=None, view_name=None,
                    recirc=None):
    """Create a Drafting View holding the WSFU take-off in the firm WFUTO
    layout: title band, job block, the two row header take-off table with
    totals and footnotes, the design loads, the firm pipe sizing chart and,
    when the system has a return, the hot water recirculation sizing.

    Args:
        doc: the Revit Document. Must already be inside a Transaction.
        graph: the traversed WaterGraph.
        header: dict with "date", "job", "job_number", "by".
        sizing_result: optional, adds a one-line sizing summary.
        view_name: optional explicit name.
        recirc: optional, what water_recirc.analyze returned.

    Returns:
        (ViewDrafting, content_height_ft, content_width_ft), or (None, 0, 0)
        if the view could not be created.
    """
    view = _new_drafting_view(doc, view_name or _default_view_name(header))
    if view is None:
        return None, 0.0, 0.0
    st = _Style(doc)

    columns = water_report.WSFU_COLUMNS
    width = sum([c[2] for c in columns], 0.0)
    ox = 0.0
    y = 0.0

    # --- Title band and job block ------------------------------------------
    y = _band(st, view, ox, y, width, "WATER FIXTURE UNIT (WFU) TAKE OFF",
              "2024 INTERNATIONAL PLUMBING CODE", BAND_H)
    label_w = 0.090
    value_w = (width - 2.0 * label_w) / 2.0
    y = _info_grid(st, view, ox, y, [label_w, value_w, label_w, value_w], [
        ["DATE", header.get("date", ""), "JOB NO.", header.get("job_number", "")],
        ["JOB", header.get("job", ""), "BY", header.get("by", "")],
    ])
    y -= GAP

    # --- The take-off ------------------------------------------------------
    y = _band(st, view, ox, y, width, "WATER SUPPLY FIXTURE UNITS*",
              "BASIS 2024 IPC TABLE E103.3(2)", SUB_BAND_H)
    data = water_report.wsfu_data(graph)
    rows = [[row.get(c[0], "") for c in columns] for row in data["rows"]]
    if rows:
        totals = [data["totals"].get(c[0], "") for c in columns]
        y = _styled_table(st, view, ox, y,
                          [(_WSFU_HEADERS[c[0]], c[2]) for c in columns],
                          rows + [totals], groups=_WSFU_GROUPS,
                          totals_row=True)
    else:
        y = _styled_table(st, view, ox, y,
                          [(_WSFU_HEADERS[c[0]], c[2]) for c in columns],
                          [["No water fixtures found on the traversed "
                            "network."] + [""] * (len(columns) - 1)],
                          groups=_WSFU_GROUPS)
    y -= PAD
    for note in WSFU_FOOTNOTES:
        for wrapped in _wrap(note, 125):
            _t(st, view, ox, y, wrapped)
            y -= TEXT_H

    # --- Design loads ------------------------------------------------------
    y -= GAP / 2.0
    load_lines = list(water_report.wsfu_load_lines(data))
    if sizing_result is not None:
        totals = sizing_result["totals"]
        load_lines.append("Pipes sized: {} of {}.  Total length {:.1f} ft.".format(
            totals["sized_count"], totals["pipe_count"],
            float(totals["total_length_ft"])))
        if totals["unsized_count"]:
            load_lines.append("{} pipe(s) NOT sized - see the pyRevit window "
                              "for the reason on each.".format(
                                  totals["unsized_count"]))
    if load_lines:
        y = _band(st, view, ox, y, width, "DESIGN LOAD", None, SUB_BAND_H)
        y = _note_panel(st, view, ox, y, width, load_lines)

    # --- Firm pipe sizing chart --------------------------------------------
    # The same rows the sizing looked each pipe up in.
    chart = water_tables.sizing_chart()
    chart_w = sum(CHART_WIDTHS, 0.0)
    y -= GAP
    y = _band(st, view, ox, y, chart_w, "FIRM PIPE SIZING CHART", None,
              SUB_BAND_H)
    y = _styled_table(st, view, ox, y,
                      list(zip(chart["headers"], CHART_WIDTHS)),
                      chart["rows"])

    # --- Hot water recirculation sizing ------------------------------------
    if recirc is not None and recirc["legs"]:
        rdata = water_report.recirc_data(recirc)
        rcols = water_report.RECIRC_COLUMNS
        rwidth = sum([c[2] for c in rcols], 0.0)
        width = max(width, rwidth)
        y -= GAP
        y = _band(st, view, ox, y, rwidth, water_report.RECIRC_TITLE, None,
                  SUB_BAND_H)
        y = _note_panel(st, view, ox, y, rwidth, rdata["basis"])
        y = _styled_table(st, view, ox, y, [(c[1], c[2]) for c in rcols],
                          [[row[c[0]] for c in rcols] for row in rdata["rows"]])
        y -= PAD
        for line in rdata["summary"]:
            _t(st, view, ox, y, line, "bold")
            y -= TEXT_H

    return view, (0.0 - y), max(width, chart_w)


def place_on_sheet(doc, view, content_h, content_w, sheet_number, sheet_name):
    """Put the drafting view on a new sheet, left-anchored near the top left.

    Returns (ViewSheet, Viewport) or (None, None) if no title block exists.
    Must already be inside a Transaction.
    """
    title_block_id = _title_block_type_id(doc)
    if title_block_id is None:
        return None, None

    try:
        sheet = ViewSheet.Create(doc, title_block_id)
    except Exception:
        return None, None

    try:
        sheet.SheetNumber = sheet_number
    except Exception:
        pass
    # Sheets and views share one name pool in Revit, and the drafting view has
    # already claimed the base name, so this always needs to differ.
    try:
        sheet.Name = sheet_name
    except Exception:
        pass

    try:
        # Viewport.Create positions by CENTRE, so the centre is derived from
        # the wanted left/top edge and the measured content size. The table
        # width changes with the fixture list, so anchoring the edge rather
        # than the centre keeps it from drifting sideways run to run.
        left_edge_x = 0.30
        top_edge_y = 2.30
        viewport = Viewport.Create(doc, sheet.Id, view.Id, XYZ(0, 0, 0))
        doc.Regenerate()
        viewport.SetBoxCenter(XYZ(left_edge_x + content_w / 2.0,
                                  top_edge_y - content_h / 2.0, 0.0))
    except Exception:
        return sheet, None

    _hide_viewport_title(doc, viewport)
    return sheet, viewport


def build_performance_view(doc, view_name, summary_lines, columns, rows,
                           legend, fill_id, extra_tables=None):
    """Drafting view for Water Performance: summary, flagged table, legend.

    Args:
        doc: the Revit Document. Must already be inside a Transaction.
        view_name: wanted view name, suffixed if taken.
        summary_lines: plain text lines printed above the table.
        columns: [(header, width_ft)].
        rows: [[cell text, ...]], already in display order. May be empty.
        legend: [(Revit Color, text)], one swatch row each.
        fill_id: solid fill pattern id, so the swatches match the plan.
        extra_tables: optional [(title, columns, rows, note_lines)] drawn
            after the legend, e.g. the pressure loss table.

    Returns:
        (ViewDrafting, content_height_ft, content_width_ft), or (None, 0, 0).
    """
    view = _new_drafting_view(doc, view_name)
    if view is None:
        return None, 0.0, 0.0
    st = _Style(doc)

    width = sum([c[1] for c in columns], 0.0)
    for _title, extra_columns, _rows, _notes in (extra_tables or []):
        width = max(width, sum([c[1] for c in extra_columns], 0.0))
    ox = 0.0
    y = 0.0

    # First summary line is the title, the rest sit in a panel under it.
    title = summary_lines[0] if summary_lines else "WATER PERFORMANCE"
    y = _band(st, view, ox, y, width, title, None, BAND_H)
    if len(summary_lines) > 1:
        y = _note_panel(st, view, ox, y, width, summary_lines[1:])
    y -= GAP

    y = _band(st, view, ox, y, width, "PIPES NEEDING ATTENTION", None,
              SUB_BAND_H)
    if rows:
        y = _styled_table(st, view, ox, y, columns, rows,
                          left_cols=(0, len(columns) - 1))
    else:
        y = _note_panel(st, view, ox, y, width,
                        ["Every judged pipe is at table size."])
    y -= GAP

    # Legend: a swatch the plan colour, then its meaning, in an outlined box.
    legend_w = min(width, 0.62)
    y = _band(st, view, ox, y, legend_w, "COLOR LEGEND", None, SUB_BAND_H)
    legend_top = y
    swatch = 0.014
    for color, meaning in legend:
        if st.region_type_id is not None:
            _legend_swatch(doc, view, st.region_type_id, ox + PAD * 2.0,
                           y - swatch - (ROW_H - swatch) / 2.0, swatch,
                           color, fill_id)
        _t(st, view, ox + swatch + PAD * 3.0, y, meaning)
        y -= ROW_H
    _outline(st, view, ox, legend_top, ox + legend_w, y)

    for title, extra_columns, extra_rows, notes in (extra_tables or []):
        y -= GAP
        table_w = sum([c[1] for c in extra_columns], 0.0)
        y = _band(st, view, ox, y, table_w, title, None, SUB_BAND_H)
        if extra_rows:
            y = _styled_table(st, view, ox, y, extra_columns, extra_rows)
        if notes:
            y = _note_panel(st, view, ox, y, table_w, notes)
    return view, (0.0 - y), width


def build_tables_view(doc, view_name, title, tables):
    """Drafting view holding only result tables (no flagged pipes, no
    legend): a title band, then each (title, columns, rows, notes).

    Returns (ViewDrafting, height_ft, width_ft), or (None, 0, 0).
    """
    view = _new_drafting_view(doc, view_name)
    if view is None:
        return None, 0.0, 0.0
    st = _Style(doc)
    width = max([sum([c[1] for c in t[1]], 0.0) for t in tables] or [0.6])
    y = _band(st, view, 0.0, 0.0, width, title, None, BAND_H)
    for table_title, columns, rows, notes in tables:
        y -= GAP
        table_w = sum([c[1] for c in columns], 0.0)
        y = _band(st, view, 0.0, y, table_w, table_title, None, SUB_BAND_H)
        if rows:
            y = _styled_table(st, view, 0.0, y, columns, rows)
        if notes:
            y = _note_panel(st, view, 0.0, y, table_w, notes)
    return view, (0.0 - y), width


def place_plan_and_table(doc, plan_view, table_view, table_h, table_w,
                         sheet_number, sheet_name, wsfu_view=None):
    """New sheet with the colored plan and the table beside it.

    Same placement as the Duct Performance sheets (centres hand placed on the
    RJA title block and read back over Revit MCP), so the two performance
    tools produce sheets that look alike. Must be inside a Transaction.

    wsfu_view, when given, is the WSFU take-off, stacked under the table in
    the same column so the two read as one (Colin, 2026-10-21).

    Returns the ViewSheet, or None when there is no title block.
    """
    title_block_id = _title_block_type_id(doc)
    if title_block_id is None:
        return None
    sheet = ViewSheet.Create(doc, title_block_id)
    try:
        sheet.SheetNumber = sheet_number
    except Exception:
        pass
    try:
        sheet.Name = sheet_name
    except Exception:
        pass

    # Regenerate before SetBoxCenter: a box centre set before the new
    # viewport's crop has settled lands offset (found on Duct Performance).
    plan_vp = Viewport.Create(doc, sheet.Id, plan_view.Id, XYZ(0, 0, 0))
    doc.Regenerate()
    plan_vp.SetBoxCenter(XYZ(1.201, 1.212, 0))
    _hide_viewport_title(doc, plan_vp)

    if table_view is not None:
        # Left edge and bottom pinned, centre derived, because the table
        # width and height change run to run.
        left_edge_x = -1.014 - 1.360 / 2.0
        bottom_y = 1.546
        table_vp = Viewport.Create(
            doc, sheet.Id, table_view.Id,
            XYZ(left_edge_x + table_w / 2.0, bottom_y + table_h / 2.0, 0))
        _hide_viewport_title(doc, table_vp)

    if wsfu_view is not None:
        # Same left edge, top just under the table's bottom. Created at the
        # origin and moved by its measured outline, so its size need not be
        # known before Revit draws it.
        left_edge_x = -1.014 - 1.360 / 2.0
        top_y = 1.546 - GAP
        wsfu_vp = Viewport.Create(doc, sheet.Id, wsfu_view.Id, XYZ(0, 0, 0))
        doc.Regenerate()
        outline = wsfu_vp.GetBoxOutline()
        height = outline.MaximumPoint.Y - outline.MinimumPoint.Y
        width = outline.MaximumPoint.X - outline.MinimumPoint.X
        wsfu_vp.SetBoxCenter(XYZ(left_edge_x + width / 2.0,
                                 top_y - height / 2.0, 0))
        _hide_viewport_title(doc, wsfu_vp)
    return sheet


def _default_view_name(header):
    job_number = header.get("job_number", "") or ""
    date = header.get("date", "") or ""
    if job_number:
        return "WSFU Take-Off - {} - {}".format(job_number, date)
    return "WSFU Take-Off - {}".format(date)
