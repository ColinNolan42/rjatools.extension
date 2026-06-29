# -*- coding: utf-8 -*-
"""
Duct Velocity.pushbutton/script.py  --  HVAC Phase 1

Colors ductwork in a copied floor plan view by velocity (green/yellow/red)
per SMACNA commercial low-velocity defaults (editable via dialog).
Also places FPM text annotations on each colored duct.

Run HVAC Diagnose first to verify the network and CFM values, then run this.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction,
    BuiltInCategory, ViewSheet, ViewType,
    Viewport, ViewDuplicateOption, ElementId, XYZ,
    OverrideGraphicSettings, Color,
    TextNote, TextNoteOptions, HorizontalTextAlignment
)
from Autodesk.Revit.UI.Selection import ObjectType

from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    HorizontalAlignment, VerticalAlignment, SizeToContent
)
from System.Windows.Controls import (
    Grid, Label, TextBox, Button, StackPanel,
    ColumnDefinition, RowDefinition, Orientation
)
from System.Windows import GridLength

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import hvac_graph

# ── SMACNA colors ─────────────────────────────────────────────────────────────
GREEN  = Color(0,   200,  0)
YELLOW = Color(255, 215,  0)
RED    = Color(210,  40, 40)
GRAY   = Color(160, 160, 160)

_COLOR_MAP = {'GREEN': GREEN, 'YELLOW': YELLOW, 'RED': RED, 'GRAY': GRAY}


# ── velocity settings dialog ───────────────────────────────────────────────────
def show_velocity_settings_dialog():
    """WPF dialog to edit SMACNA velocity thresholds.
    Returns dict {sys_class: (green_fpm, yellow_fpm)} or None if cancelled.
    """
    ROWS = [
        ('Supply Air',  2000, 2500),
        ('Return Air',  1500, 2000),
        ('Exhaust Air', 1200, 1500),
        ('Outside Air', 1200, 1500),
    ]

    result = [None]  # mutable container — no nonlocal in Python 2.7
    boxes  = {}      # (row_idx, col_idx) -> TextBox  col 0=green, 1=yellow

    win = Window()
    win.Title  = 'Duct Velocity Settings'
    win.Width  = 440
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen
    win.ResizeMode = System.Windows.ResizeMode.NoResize

    outer = StackPanel()
    outer.Margin = Thickness(14)

    # Instruction label
    intro = Label()
    intro.Content = 'Edit SMACNA velocity limits (FPM) before coloring the view:'
    intro.Margin  = Thickness(0, 0, 0, 8)
    outer.Children.Add(intro)

    # Header row + data rows in a Grid
    grid = Grid()
    for w in (140, 130, 130):
        cd = ColumnDefinition()
        cd.Width = GridLength(w)
        grid.ColumnDefinitions.Add(cd)
    for _ in range(len(ROWS) + 1):   # 1 header + 4 data rows
        rd = RowDefinition()
        rd.Height = GridLength(32)
        grid.RowDefinitions.Add(rd)

    def _lbl(text, col, row):
        lb = Label()
        lb.Content = text
        lb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(lb, col)
        Grid.SetRow(lb, row)
        grid.Children.Add(lb)

    _lbl('System Type',    0, 0)
    _lbl('Green ≤ FPM',  1, 0)
    _lbl('Yellow ≤ FPM', 2, 0)

    for i, (sys_class, g_def, y_def) in enumerate(ROWS):
        r = i + 1
        _lbl(sys_class, 0, r)
        for j, val in enumerate((g_def, y_def)):
            tb = TextBox()
            tb.Text = str(val)
            tb.Width = 90
            tb.Margin = Thickness(4, 4, 4, 4)
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.HorizontalAlignment = HorizontalAlignment.Left
            Grid.SetColumn(tb, j + 1)
            Grid.SetRow(tb, r)
            grid.Children.Add(tb)
            boxes[(i, j)] = tb

    outer.Children.Add(grid)

    # OK / Cancel buttons
    btn_panel = StackPanel()
    btn_panel.Orientation = Orientation.Horizontal
    btn_panel.HorizontalAlignment = HorizontalAlignment.Right
    btn_panel.Margin = Thickness(0, 12, 0, 0)

    ok_btn = Button()
    ok_btn.Content = 'OK'
    ok_btn.Width   = 72
    ok_btn.Margin  = Thickness(0, 0, 8, 0)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width   = 72

    def on_ok(s, e):
        out = {}
        try:
            for i, (sys_class, _, _) in enumerate(ROWS):
                g = float(boxes[(i, 0)].Text)
                y = float(boxes[(i, 1)].Text)
                if g >= y:
                    forms.alert(
                        'Green limit must be less than Yellow limit for {}.'.format(sys_class),
                        title='Invalid Input')
                    return
                out[sys_class] = (g, y)
            result[0] = out
        except ValueError:
            forms.alert('Enter valid numbers for all fields.', title='Invalid Input')
            return
        win.Close()

    def on_cancel(s, e):
        win.Close()

    ok_btn.Click     += on_ok
    cancel_btn.Click += on_cancel
    btn_panel.Children.Add(ok_btn)
    btn_panel.Children.Add(cancel_btn)
    outer.Children.Add(btn_panel)

    win.Content = outer
    win.ShowDialog()
    return result[0]


# ── helpers ────────────────────────────────────────────────────────────────────
def _custom_smacna_label(fpm, sys_class, custom_limits):
    limits = custom_limits.get(sys_class, hvac_graph.SMACNA.get(sys_class, None))
    if fpm <= 0 or limits is None:
        return 'GRAY'
    lo, hi = limits
    if fpm <= lo:
        return 'GREEN'
    elif fpm <= hi:
        return 'YELLOW'
    else:
        return 'RED'


def _duct_midpoint(duct):
    try:
        return duct.Location.Curve.Evaluate(0.5, True)
    except Exception:
        return None


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(elem.Id.IntegerValue)


# ── main ───────────────────────────────────────────────────────────────────────
def main():
    output.print_md('## Duct Velocity Visualizer')
    output.print_md('_Tip: run HVAC Diagnose first to verify CFM values and network._')
    output.print_md('---')

    # 1. Validate active view
    active_view = doc.ActiveView
    if active_view.ViewType != ViewType.FloorPlan:
        forms.alert(
            'Open a floor plan view first, then run Duct Velocity.',
            title='Wrong View Type', exitscript=True
        )

    # 2. Velocity settings dialog
    custom_limits = show_velocity_settings_dialog()
    if custom_limits is None:
        output.print_md('**Cancelled.**')
        return

    # 3. Select element
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            'Select any duct, air terminal, or AHU in the system to visualize'
        )
    except Exception:
        output.print_md('**Cancelled.**')
        return

    sel_elem = doc.GetElement(ref.ElementId)
    output.print_md('Selected id: **{}**'.format(sel_elem.Id.IntegerValue))

    # 4. Build network (traversal + CFM + velocity)
    output.print_md('Traversing duct network...')
    net = hvac_graph.build_network(sel_elem, doc)

    if net.root is None or len(net.errors) > 0:
        output.print_md('**Errors found — run HVAC Diagnose for details:**')
        for e in net.errors:
            output.print_md('- {}'.format(e))
        return

    ahu_label = _elem_name(net.root)
    output.print_md('AHU: **{}**  (id {})'.format(ahu_label, net.root.Id.IntegerValue))
    output.print_md('Network: {} elements  |  {} ducts  |  {} terminals'.format(
        len(net.nodes), len(net.duct_results), len(net.terminal_cfms)))

    if net.warnings:
        for w in net.warnings:
            output.print_md(':warning: {}'.format(w))

    # 5. Find source sheet number
    source_sheet_num = 'NoSheet'
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            if vp is not None and vp.ViewId == active_view.Id:
                source_sheet_num = sheet.SheetNumber
                break

    output.print_md('Source sheet: **{}**'.format(source_sheet_num))

    # 6. Title block and solid fill
    tb_list = list(FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsElementType())
    tb_id   = tb_list[0].Id if len(tb_list) > 0 else ElementId.InvalidElementId
    fill_id = hvac_graph.solid_fill_pattern_id(doc)

    # 7. Text height: aim for 5/64" printed size at the view's print scale
    view_scale = getattr(active_view, 'Scale', 48)
    text_h_ft  = (5.0 / (64.0 * 12.0)) * float(view_scale)

    # 8. Transaction: copy view → color overrides → FPM annotations → sheet
    t = Transaction(doc, 'Duct Velocity Visualizer')
    t.Start()
    try:
        # Copy floor plan
        new_vid  = active_view.Duplicate(ViewDuplicateOption.Duplicate)
        new_view = doc.GetElement(new_vid)
        base_name = 'Ducting Velocities - ' + source_sheet_num
        try:
            new_view.Name = base_name
        except Exception:
            new_view.Name = base_name + ' (2)'

        # Color overrides — use custom limits
        counts      = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'GRAY': 0}
        duct_labels = {}   # ElementId -> label string (used below for annotations)

        for eid, dr in net.duct_results.items():
            label             = _custom_smacna_label(dr.fpm, dr.sys_class, custom_limits)
            duct_labels[eid]  = label
            color             = _COLOR_MAP.get(label, GRAY)
            ogs               = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(eid, ogs)
            counts[label] = counts.get(label, 0) + 1

        # FPM annotations — skip GRAY ducts (no reliable CFM)
        note_opts = TextNoteOptions()
        note_opts.HorizontalAlignment = HorizontalTextAlignment.Center
        annotation_count = 0
        for eid, dr in net.duct_results.items():
            if duct_labels.get(eid, 'GRAY') == 'GRAY':
                continue
            mid = _duct_midpoint(dr.elem)
            if mid is None:
                continue
            try:
                fpm_text = '{:.0f} FPM'.format(dr.fpm)
                TextNote.Create(doc, new_vid, mid, text_h_ft, fpm_text, note_opts)
                annotation_count += 1
            except Exception:
                # Alternate overload (Revit version difference)
                try:
                    TextNote.Create(doc, new_vid, mid, fpm_text, note_opts)
                    annotation_count += 1
                except Exception:
                    pass  # annotation is optional; never abort the transaction

        # Output sheet
        new_sheet             = ViewSheet.Create(doc, tb_id)
        new_sheet.SheetNumber = 'DV-' + source_sheet_num
        new_sheet.Name        = 'Ducting Velocities - ' + source_sheet_num

        # Place viewport
        Viewport.Create(doc, new_sheet.Id, new_vid, XYZ(1.1, 0.8, 0))

        t.Commit()
    except Exception as ex:
        t.RollBack()
        output.print_md('**Error — transaction rolled back:** {}'.format(str(ex)))
        return

    # 9. Summary
    sa = custom_limits.get('Supply Air',  (0, 0))
    ra = custom_limits.get('Return Air',  (0, 0))
    ea = custom_limits.get('Exhaust Air', (0, 0))
    oa = custom_limits.get('Outside Air', (0, 0))

    output.print_md('---')
    output.print_md('## Done')
    output.print_md('Sheet **DV-{}** created  |  {} FPM annotations placed.'.format(
        source_sheet_num, annotation_count))
    output.print_md('')
    output.print_md('**Limits used:**')
    output.print_md('| System | Green ≤ | Yellow ≤ | Red > |')
    output.print_md('| --- | --- | --- | --- |')
    output.print_md('| Supply Air  | {} FPM | {} FPM | {} FPM |'.format(sa[0], sa[1], sa[1]))
    output.print_md('| Return Air  | {} FPM | {} FPM | {} FPM |'.format(ra[0], ra[1], ra[1]))
    output.print_md('| Exhaust Air | {} FPM | {} FPM | {} FPM |'.format(ea[0], ea[1], ea[1]))
    output.print_md('| Outside Air | {} FPM | {} FPM | {} FPM |'.format(oa[0], oa[1], oa[1]))
    output.print_md('')
    output.print_md('| Color | Duct Count | Meaning |')
    output.print_md('| --- | --- | --- |')
    output.print_md('| Green  | {} | Within SMACNA limit |'.format(counts.get('GREEN',  0)))
    output.print_md('| Yellow | {} | Approaching limit   |'.format(counts.get('YELLOW', 0)))
    output.print_md('| Red    | {} | Exceeds limit       |'.format(counts.get('RED',    0)))
    output.print_md('| Gray   | {} | No CFM data         |'.format(counts.get('GRAY',   0)))

    uidoc.ActiveView = new_sheet


main()
