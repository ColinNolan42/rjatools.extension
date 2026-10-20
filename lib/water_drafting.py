# -*- coding: ascii -*-
# water_drafting.py
# Places the WSFU take-off table into a Revit Drafting View, and optionally
# onto its own sheet. Same approach as the Duct Velocity schedule: detail
# curves for the grid, TextNotes for the cells, drafting view at Scale 1 so
# model feet equal paper feet.
#
# The columns, headers and cell values come from water_report.WSFU_COLUMNS and
# water_report.wsfu_data(). Nothing about the table is defined twice - the
# console table and this schedule cannot disagree.
#
# This is the only water module that imports the Revit API, so the rest stay
# testable in plain CPython.
#
# IronPython 2.7

from Autodesk.Revit.DB import (
    XYZ, Line, TextNote, TextNoteOptions, TextNoteType,
    ViewDrafting, ViewFamily, ViewFamilyType, ViewSheet, Viewport,
    FilteredElementCollector, BuiltInCategory, CurveLoop, FilledRegion,
    FilledRegionType, OverrideGraphicSettings, ElementId, Color,
    FillPatternElement, BuiltInParameter, HorizontalTextAlignment)

import water_report
import water_tables


# Layout, in feet at the drafting view's 1:1 scale, so these are printed
# inches divided by 12. Matched to the Duct Velocity schedule so the two
# tools' sheets look like they came from the same office.
PAD = 0.004        # text inset from the cell edge, about 1/24 in
HEAD_H = 0.030     # header row height, about 3/8 in
ROW_H = 0.022      # data row height, about 1/4 in
TEXT_H = 0.018     # plain text line height, about 1/5 in
GAP = 0.020        # gap between blocks


CHART_WIDTHS = (0.090, 0.160, 0.150, 0.120)   # firm pipe sizing chart

# RJA table look (Colin, 2026-10-20: "a little bit more flair"). RJA blue is
# the theme colour sampled from the logo (#0070C0), the same one the dialogs
# use, so the sheets and the dialogs read as one product.
RJA_BLUE = (0, 112, 192)
HEAD_FILL = (221, 235, 247)     # header rows, label cells
ZEBRA_FILL = (244, 247, 251)    # every other body row
TOTALS_FILL = (189, 215, 238)   # totals row
WHITE = (255, 255, 255)
BAND_H = 0.036                  # title band, about 7/16 in
SUB_BAND_H = 0.028              # section band, about 11/32 in
OUTER_WEIGHT = 5                # outer border line weight

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


# ---------------------------------------------------------------------------
# Styled drawing
# ---------------------------------------------------------------------------

class _Style(object):
    """Text types, solid fill and filled region type for one drafting view.

    Three RJA text types are made once per project by duplicating the
    project's first text type: "RJA Table Text", "RJA Table Bold" and "RJA
    Table Title" (bold, white, 1/8 in). All have a transparent background so
    the shaded rows show through. If a type cannot be made the project's
    first text type is used instead, so the table still draws.
    """

    def __init__(self, doc):
        self.doc = doc
        base = _text_note_type_id(doc)
        self.types = {
            "text": _ensure_text_type(doc, base, "RJA Table Text"),
            "bold": _ensure_text_type(doc, base, "RJA Table Bold", bold=True),
            "title": _ensure_text_type(doc, base, "RJA Table Title",
                                       bold=True, rgb=WHITE,
                                       size_ft=1.0 / 8.0 / 12.0),
        }
        self.fill_id = _solid_fill_id(doc)
        self.region_type_id = None
        for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
            self.region_type_id = frt.Id
            break


def _t(st, view, x, y, value, kind="text", align="left"):
    """A TextNote whose origin is the TOP of the text box, flowing down.
    align "center" or "right" puts x at the middle or right of the text."""
    if not value:
        return None
    try:
        opts = TextNoteOptions(st.types[kind])
        if align == "center":
            opts.HorizontalAlignment = HorizontalTextAlignment.Center
        elif align == "right":
            opts.HorizontalAlignment = HorizontalTextAlignment.Right
        pad = PAD if align == "left" else (-PAD if align == "right" else 0.0)
        return TextNote.Create(st.doc, view.Id, XYZ(x + pad, y - PAD, 0.0),
                               str(value), opts)
    except Exception:
        return None


def _fill(st, view, x0, top, x1, bottom, rgb):
    """A solid filled rectangle drawn in rgb, edges in the same colour."""
    if st.region_type_id is None or x1 <= x0 or top <= bottom:
        return
    pts = [XYZ(x0, bottom, 0.0), XYZ(x1, bottom, 0.0),
           XYZ(x1, top, 0.0), XYZ(x0, top, 0.0)]
    loop = CurveLoop()
    for i in range(4):
        loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    try:
        region = FilledRegion.Create(st.doc, st.region_type_id, view.Id, [loop])
        color = Color(rgb[0], rgb[1], rgb[2])
        ogs = OverrideGraphicSettings()
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetProjectionLineColor(color)
        if st.fill_id != ElementId.InvalidElementId:
            ogs.SetSurfaceForegroundPatternId(st.fill_id)
        view.SetElementOverrides(region.Id, ogs)
    except Exception:
        pass


def _line(st, view, x0, y0, x1, y1, weight=None):
    try:
        curve = st.doc.Create.NewDetailCurve(
            view, Line.CreateBound(XYZ(x0, y0, 0.0), XYZ(x1, y1, 0.0)))
        if weight:
            ogs = OverrideGraphicSettings()
            ogs.SetProjectionLineWeight(weight)
            view.SetElementOverrides(curve.Id, ogs)
    except Exception:
        pass


def _band(st, view, ox, top, width, title, right_text, height):
    """RJA blue band with a white title, optional text at its right end.
    Returns the y of its bottom edge."""
    bottom = top - height
    _fill(st, view, ox, top, ox + width, bottom, RJA_BLUE)
    inset = (height - TEXT_H) / 2.0
    _t(st, view, ox, top - inset + PAD, title, "title")
    if right_text:
        _t(st, view, ox + width, top - inset + PAD, right_text, "title",
           "right")
    return bottom


def _info_grid(st, view, ox, top, widths, rows):
    """Label / value grid: even columns are shaded bold labels."""
    xs = [ox]
    for w in widths:
        xs.append(xs[-1] + w)
    bottom = top - ROW_H * len(rows)
    for r, row in enumerate(rows):
        row_top = top - ROW_H * r
        for c, value in enumerate(row):
            label = (c % 2 == 0)
            if label:
                _fill(st, view, xs[c], row_top, xs[c + 1], row_top - ROW_H,
                      HEAD_FILL)
            _t(st, view, xs[c], row_top, value, "bold" if label else "text")
    for r in range(len(rows) + 1):
        _line(st, view, ox, top - ROW_H * r, xs[-1], top - ROW_H * r)
    for x in xs:
        _line(st, view, x, top, x, bottom)
    _outline(st, view, ox, top, xs[-1], bottom)
    return bottom


def _note_panel(st, view, ox, top, width, lines, chars=None):
    """Plain lines in a lightly shaded, outlined panel. Returns its bottom."""
    wrapped = []
    for line in lines:
        wrapped.extend(_wrap(line, chars or int(width / 0.0062)))
    bottom = top - TEXT_H * len(wrapped) - 2.0 * PAD
    _fill(st, view, ox, top, ox + width, bottom, ZEBRA_FILL)
    y = top - PAD
    for line in wrapped:
        _t(st, view, ox, y, line)
        y -= TEXT_H
    _outline(st, view, ox, top, ox + width, bottom)
    return bottom


def _outline(st, view, x0, top, x1, bottom):
    for a, b in (((x0, top), (x1, top)), ((x1, top), (x1, bottom)),
                 ((x1, bottom), (x0, bottom)), ((x0, bottom), (x0, top))):
        _line(st, view, a[0], a[1], b[0], b[1], OUTER_WEIGHT)


def _styled_table(st, view, ox, top, columns, rows, groups=None,
                  totals_row=False, left_cols=(0,)):
    """A gridded table in the RJA look.

    columns: [(header, width)]. rows: [[cell, ...]]. groups: optional
    [(label, column span)] drawn as a row above the headers. With totals_row
    the last row is shaded and bold. Columns in left_cols are left aligned,
    the rest centred. Returns the y of the bottom edge.
    """
    widths = [c[1] for c in columns]
    xs = [ox]
    for w in widths:
        xs.append(xs[-1] + w)
    right = xs[-1]

    head_top = top
    if groups:
        group_bottom = top - ROW_H
        _fill(st, view, ox, top, right, group_bottom, HEAD_FILL)
        col = 0
        for label, span in groups:
            x0, x1 = xs[col], xs[col + span]
            if label:
                _t(st, view, (x0 + x1) / 2.0, top, label, "bold", "center")
                _line(st, view, x0, group_bottom, x1, group_bottom)
            _line(st, view, x0, top, x0, group_bottom)
            col += span
        head_top = group_bottom

    head_bottom = head_top - HEAD_H
    _fill(st, view, ox, head_top, right, head_bottom, HEAD_FILL)
    row_tops = [head_bottom]
    for _ in rows:
        row_tops.append(row_tops[-1] - ROW_H)
    bottom = row_tops[-1]

    last = len(rows) - 1
    for r in range(len(rows)):
        if totals_row and r == last:
            _fill(st, view, ox, row_tops[r], right, row_tops[r + 1],
                  TOTALS_FILL)
        elif r % 2 == 1:
            _fill(st, view, ox, row_tops[r], right, row_tops[r + 1],
                  ZEBRA_FILL)

    _line(st, view, ox, head_bottom, right, head_bottom)
    for y in row_tops[1:-1]:
        _line(st, view, ox, y, right, y)
    if totals_row and rows:
        _line(st, view, ox, row_tops[last], right, row_tops[last], 3)
    for x in xs[1:-1]:
        _line(st, view, x, head_top, x, bottom)
    _outline(st, view, ox, top, right, bottom)

    for c, column in enumerate(columns):
        if c in left_cols:
            _t(st, view, xs[c], head_top, column[0], "bold")
        else:
            _t(st, view, (xs[c] + xs[c + 1]) / 2.0, head_top, column[0],
               "bold", "center")
    for r, row in enumerate(rows):
        kind = "bold" if (totals_row and r == last) else "text"
        for c, value in enumerate(row):
            if not value:   # TextNote.Create rejects an empty string
                continue
            if c in left_cols:
                _t(st, view, xs[c], row_tops[r], value, kind)
            else:
                _t(st, view, (xs[c] + xs[c + 1]) / 2.0, row_tops[r], value,
                   kind, "center")
    return bottom


def _ensure_text_type(doc, base_id, name, bold=False, rgb=None, size_ft=None):
    """The TextNoteType called name, made from base_id when missing.

    I believe TEXT_STYLE_BOLD, LINE_COLOR (BGR packed int), TEXT_BACKGROUND
    (1 = transparent) and TEXT_SIZE are the right BuiltInParameters; each is
    set on its own so one that does not apply only loses that one look.
    Falls back to base_id.
    """
    if base_id is None:
        return None
    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
        try:
            p = tnt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p is not None and p.AsString() == name:
                return tnt.Id
        except Exception:
            continue
    try:
        new_type = doc.GetElement(base_id).Duplicate(name)
    except Exception:
        return base_id
    settings = [(BuiltInParameter.TEXT_BACKGROUND, 1),
                (BuiltInParameter.TEXT_STYLE_BOLD, 1 if bold else 0)]
    if rgb is not None:
        settings.append((BuiltInParameter.LINE_COLOR,
                         rgb[0] + rgb[1] * 256 + rgb[2] * 65536))
    if size_ft is not None:
        settings.append((BuiltInParameter.TEXT_SIZE, float(size_ft)))
    for bip, value in settings:
        try:
            p = new_type.get_Parameter(bip)
            if p is not None and not p.IsReadOnly:
                p.Set(value)
        except Exception:
            pass
    # Fallback for the colour: the Type Properties "Color" parameter by name,
    # in case LINE_COLOR is not the one that holds it in this Revit version.
    if rgb is not None:
        try:
            p = new_type.LookupParameter("Color")
            if p is not None and not p.IsReadOnly:
                p.Set(rgb[0] + rgb[1] * 256 + rgb[2] * 65536)
        except Exception:
            pass
    return new_type.Id


def _solid_fill_id(doc):
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except Exception:
            pass
    return ElementId.InvalidElementId


def _new_drafting_view(doc, name):
    drafting_type_id = _drafting_view_type_id(doc)
    if drafting_type_id is None or _text_note_type_id(doc) is None:
        return None
    try:
        view = ViewDrafting.Create(doc, drafting_type_id)
        view.Scale = 1
    except Exception:
        return None
    _name_view(view, name)
    return view


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


def place_plan_and_table(doc, plan_view, table_view, table_h, table_w,
                         sheet_number, sheet_name):
    """New sheet with the colored plan and the table beside it.

    Same placement as the Duct Performance sheets (centres hand placed on the
    RJA title block and read back over Revit MCP), so the two performance
    tools produce sheets that look alike. Must be inside a Transaction.

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

    if table_view is not None:
        # Left edge and bottom pinned, centre derived, because the table
        # width and height change run to run.
        left_edge_x = -1.014 - 1.360 / 2.0
        bottom_y = 1.546
        table_vp = Viewport.Create(
            doc, sheet.Id, table_view.Id,
            XYZ(left_edge_x + table_w / 2.0, bottom_y + table_h / 2.0, 0))
        _hide_viewport_title(doc, table_vp)
    return sheet


def _legend_swatch(doc, view, region_type_id, x0, y0, size, color, fill_id):
    """A small square filled with the same color the plan uses."""
    pts = [XYZ(x0, y0, 0.0), XYZ(x0 + size, y0, 0.0),
           XYZ(x0 + size, y0 + size, 0.0), XYZ(x0, y0 + size, 0.0)]
    loop = CurveLoop()
    for i in range(4):
        loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    try:
        region = FilledRegion.Create(doc, region_type_id, view.Id, [loop])
        ogs = OverrideGraphicSettings()
        ogs.SetSurfaceForegroundPatternColor(color)
        if fill_id != ElementId.InvalidElementId:
            ogs.SetSurfaceForegroundPatternId(fill_id)
        view.SetElementOverrides(region.Id, ogs)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Drawing primitives
# ---------------------------------------------------------------------------

def _wrap(text, width):
    """Wrap on spaces. The basis lines are long sentences and a drafting view
    TextNote does not wrap itself unless it is given a width."""
    words = str(text).split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        if len(current) + 1 + len(word) <= width:
            current = current + " " + word
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


# ---------------------------------------------------------------------------
# Revit lookups
# ---------------------------------------------------------------------------

def _drafting_view_type_id(doc):
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements():
        try:
            if vft.ViewFamily == ViewFamily.Drafting:
                return vft.Id
        except Exception:
            continue
    return None


def _text_note_type_id(doc):
    """The first TextNoteType in the project.

    Arbitrary, and deliberately the same arbitrary choice the Duct Velocity
    schedule makes, so both tools' tables come out in the same text style. In
    RJA's template this lands on 3/32" Arial. Change the text type in the view
    afterwards if a job wants something else.
    """
    try:
        collector = FilteredElementCollector(doc).OfClass(TextNoteType)
        for tnt in collector.ToElements():
            return tnt.Id
    except Exception:
        pass
    return None


def _title_block_type_id(doc):
    try:
        collector = (FilteredElementCollector(doc)
                     .OfCategory(BuiltInCategory.OST_TitleBlocks)
                     .WhereElementIsElementType())
        for tb in collector.ToElements():
            return tb.Id
    except Exception:
        pass
    return None


def _hide_viewport_title(doc, viewport):
    """Best effort: switch to the title-less viewport type if one exists.

    The type is named exactly "None" in RJA's template (confirmed live in
    (2022) RJA Template M&P, whose viewport types are None, RJA - Double Line
    and RJA - Single Line), so match that rather than "No Title".

    Cosmetic only. The Duct Velocity tool found that a FilteredElementCollector
    on OST_Viewports misses this system-family type, which is why the existing
    viewports are sampled instead, and that reading vp_type.Name can throw
    inside pyRevit's IronPython even when the same query works outside it. This
    must never break the run.
    """
    try:
        for vp in FilteredElementCollector(doc).OfClass(Viewport):
            vp_type = doc.GetElement(vp.GetTypeId())
            if vp_type is None:
                continue
            try:
                name = vp_type.Name
            except Exception:
                continue
            if name is not None and name.strip().lower() in ("none", "no title"):
                viewport.ChangeTypeId(vp_type.Id)
                return
    except Exception:
        pass


def _name_view(view, base_name):
    """View names must be unique in the project, so fall back with a suffix."""
    for attempt in range(1, 20):
        candidate = base_name if attempt == 1 else "{} ({})".format(
            base_name, attempt)
        try:
            view.Name = candidate
            return candidate
        except Exception:
            continue
    return None


def _default_view_name(header):
    job_number = header.get("job_number", "") or ""
    date = header.get("date", "") or ""
    if job_number:
        return "WSFU Take-Off - {} - {}".format(job_number, date)
    return "WSFU Take-Off - {}".format(date)
