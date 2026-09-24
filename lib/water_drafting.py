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
    FilteredElementCollector, BuiltInCategory)

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


def build_wsfu_view(doc, graph, header, sizing_result=None, view_name=None):
    """Create a Drafting View holding the header, basis of design and the
    WSFU take-off table.

    Args:
        doc: the Revit Document. Must already be inside a Transaction.
        graph: the traversed WaterGraph.
        header: dict with "date", "job", "job_number", "by".
        sizing_result: optional, adds a one-line sizing summary.
        view_name: optional explicit name.

    Returns:
        (ViewDrafting, content_height_ft, content_width_ft), or (None, 0, 0)
        if the view could not be created.
    """
    drafting_type_id = _drafting_view_type_id(doc)
    if drafting_type_id is None:
        return None, 0.0, 0.0

    text_type_id = _text_note_type_id(doc)
    if text_type_id is None:
        return None, 0.0, 0.0

    try:
        view = ViewDrafting.Create(doc, drafting_type_id)
        view.Scale = 1
    except Exception:
        return None, 0.0, 0.0

    _name_view(view, view_name or _default_view_name(header))

    opts = TextNoteOptions(text_type_id)
    columns = water_report.WSFU_COLUMNS
    widths = [c[2] for c in columns]
    total_w = sum(widths, 0.0)

    ox = 0.0
    y = 0.0

    # --- Title -------------------------------------------------------------
    _text(doc, view, ox, y - PAD, water_report.WSFU_TITLE, opts)
    y -= TEXT_H + GAP / 2.0

    # --- Header block ------------------------------------------------------
    for label, key in (("DATE:", "date"), ("JOB:", "job"),
                       ("JOB NO.:", "job_number"), ("BY:", "by")):
        value = header.get(key, "") or ""
        _text(doc, view, ox, y - PAD, "{}  {}".format(label, value), opts)
        y -= TEXT_H
    y -= GAP / 2.0

    # --- Basis of design ---------------------------------------------------
    # Printed from water_tables so the sheet can never state a basis that
    # differs from the one the sizing actually used.
    _text(doc, view, ox, y - PAD, "BASIS OF DESIGN", opts)
    y -= TEXT_H
    for line in water_tables.basis_of_design_lines():
        for wrapped in _wrap(line, 110):
            _text(doc, view, ox, y - PAD, wrapped, opts)
            y -= TEXT_H
    y -= GAP

    # --- The table ---------------------------------------------------------
    data = water_report.wsfu_data(graph)
    rows = data["rows"]

    col_xs = [ox]
    for w in widths:
        col_xs.append(col_xs[-1] + w)

    table_top = y
    row_count = len(rows) + (1 if rows else 0)   # data rows plus the totals row
    total_h = HEAD_H + ROW_H * max(row_count, 1)

    row_tops = [table_top, table_top - HEAD_H]
    for _ in range(max(row_count, 1)):
        row_tops.append(row_tops[-1] - ROW_H)

    for line_y in row_tops:
        _hline(doc, view, ox, ox + total_w, line_y)
    for x in col_xs:
        _vline(doc, view, x, table_top, table_top - total_h)

    for index, column in enumerate(columns):
        _text(doc, view, col_xs[index], table_top - PAD, column[1], opts)

    if rows:
        body = rows + [data["totals"]]
        for row_index, row in enumerate(body):
            cell_y = row_tops[row_index + 1] - PAD
            for col_index, column in enumerate(columns):
                value = row.get(column[0], "")
                # TextNote.Create rejects an empty string, and a blank cell is
                # a real outcome here (a water closet has no hot column), so
                # the cell is left genuinely empty.
                if not value:
                    continue
                _text(doc, view, col_xs[col_index], cell_y, value, opts)
    else:
        _text(doc, view, ox, row_tops[1] - PAD,
              "No water fixtures found on the traversed network.", opts)

    y = table_top - total_h - GAP / 2.0

    # --- Sizing loads ------------------------------------------------------
    for line in water_report.wsfu_load_lines(data):
        _text(doc, view, ox, y - PAD, line, opts)
        y -= TEXT_H

    if sizing_result is not None:
        totals = sizing_result["totals"]
        y -= GAP / 2.0
        _text(doc, view, ox, y - PAD,
              "Pipes sized: {} of {}.  Total length {:.1f} ft.".format(
                  totals["sized_count"], totals["pipe_count"],
                  float(totals["total_length_ft"])), opts)
        y -= TEXT_H
        if totals["unsized_count"]:
            _text(doc, view, ox, y - PAD,
                  "{} pipe(s) NOT sized - see the pyRevit window for the "
                  "reason on each.".format(totals["unsized_count"]), opts)
            y -= TEXT_H

    return view, (0.0 - y), total_w


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


# ---------------------------------------------------------------------------
# Drawing primitives
# ---------------------------------------------------------------------------

def _hline(doc, view, x0, x1, y):
    doc.Create.NewDetailCurve(
        view, Line.CreateBound(XYZ(x0, y, 0.0), XYZ(x1, y, 0.0)))


def _vline(doc, view, x, y0, y1):
    doc.Create.NewDetailCurve(
        view, Line.CreateBound(XYZ(x, y0, 0.0), XYZ(x, y1, 0.0)))


def _text(doc, view, x, y, value, opts):
    """A TextNote whose origin is the TOP of the text box, flowing down."""
    if not value:
        return None
    try:
        return TextNote.Create(doc, view.Id, XYZ(x + PAD, y, 0.0),
                               str(value), opts)
    except Exception:
        return None


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
