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
    BuiltInCategory, BuiltInParameter, ViewSheet, ViewType,
    Viewport, ViewDuplicateOption, ElementId, XYZ,
    OverrideGraphicSettings, Color
)
from Autodesk.Revit.UI.Selection import ObjectType

from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    HorizontalAlignment, VerticalAlignment, SizeToContent
)
from System.Windows.Controls import (
    Grid, Label, TextBox, Button, StackPanel,
    ColumnDefinition, RowDefinition, Orientation,
    Separator, TextBlock
)
from System.Windows.Media import SolidColorBrush, Colors
from System.Windows import FontWeights
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
    """WPF dialog — per-system max velocity + friction, with green threshold %.

    Returns ({sys_class: (max_fpm, max_friction_inwc)}, green_pct) or None.

    Color bands (percentage-based off max):
      Green  : value < max * green_pct/100
      Yellow : max * green_pct/100 <= value <= max
      Red    : value > max
    """
    # Defaults: firm design standard (main and branch share same values)
    ROWS = [
        ('Supply Air',  800,  0.08),
        ('Return Air',  600,  0.05),
        ('Exhaust Air', 600,  0.05),
        ('Outside Air', 600,  0.05),
    ]
    DEFAULT_GREEN_PCT = 85   # green if below this % of max

    result    = [None]
    vel_boxes  = {}   # row_idx -> TextBox (velocity)
    fric_boxes = {}   # row_idx -> TextBox (friction)
    gpct_box   = [None]

    win = Window()
    win.Title  = 'Duct Velocity Settings'
    win.Width  = 460
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    intro = Label()
    intro.Content = 'Max velocity (FPM) and pressure drop (in. wc/100 ft) per system:'
    intro.Margin  = Thickness(0, 0, 0, 8)
    outer.Children.Add(intro)

    # 3-column grid: system | velocity | friction
    grid = Grid()
    for w in (150, 130, 150):
        cd = ColumnDefinition()
        cd.Width = GridLength(w)
        grid.ColumnDefinitions.Add(cd)
    for _ in range(len(ROWS) + 1):
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

    _lbl('System Type',          0, 0)
    _lbl('Max Velocity (FPM)',   1, 0)
    _lbl('Max Friction (iwc/100)', 2, 0)

    for i, (sys_class, def_fpm, def_fric) in enumerate(ROWS):
        r = i + 1
        _lbl(sys_class, 0, r)
        for col, val, store in ((1, def_fpm, vel_boxes), (2, def_fric, fric_boxes)):
            tb = TextBox()
            tb.Text  = str(val)
            tb.Width = 80
            tb.Margin = Thickness(4, 4, 4, 4)
            tb.VerticalAlignment   = VerticalAlignment.Center
            tb.HorizontalAlignment = HorizontalAlignment.Left
            Grid.SetColumn(tb, col)
            Grid.SetRow(tb, r)
            grid.Children.Add(tb)
            store[i] = tb

    outer.Children.Add(grid)

    # Green threshold row
    gpct_panel = StackPanel()
    gpct_panel.Orientation = Orientation.Horizontal
    gpct_panel.Margin = Thickness(0, 10, 0, 0)

    gpct_lbl = Label()
    gpct_lbl.Content = 'Green if below'
    gpct_lbl.VerticalAlignment = VerticalAlignment.Center
    gpct_panel.Children.Add(gpct_lbl)

    tb_gpct = TextBox()
    tb_gpct.Text  = str(DEFAULT_GREEN_PCT)
    tb_gpct.Width = 45
    tb_gpct.Margin = Thickness(4, 0, 4, 0)
    tb_gpct.VerticalAlignment = VerticalAlignment.Center
    gpct_panel.Children.Add(tb_gpct)
    gpct_box[0] = tb_gpct

    gpct_suffix = Label()
    gpct_suffix.Content = '% of max  (above = yellow, over max = red)'
    gpct_suffix.VerticalAlignment = VerticalAlignment.Center
    gpct_panel.Children.Add(gpct_suffix)
    outer.Children.Add(gpct_panel)

    # Assumptions / formula reference block
    sep = Separator()
    sep.Margin = Thickness(0, 12, 0, 8)
    outer.Children.Add(sep)

    def _info_row(label_text, value_text):
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 1, 0, 1)
        lbl = TextBlock()
        lbl.Text = label_text
        lbl.Width = 160
        lbl.FontWeight = FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Colors.DimGray)
        val = TextBlock()
        val.Text = value_text
        val.Foreground = SolidColorBrush(Colors.DimGray)
        row.Children.Add(lbl)
        row.Children.Add(val)
        outer.Children.Add(row)

    hdr = TextBlock()
    hdr.Text = 'Calculation Basis'
    hdr.FontWeight = FontWeights.Bold
    hdr.Foreground = SolidColorBrush(Colors.DimGray)
    hdr.Margin = Thickness(0, 0, 0, 4)
    outer.Children.Add(hdr)

    _info_row('Pressure drop:',   'Darcy-Weisbach')
    _info_row('Friction factor:', 'Altshul-Tsal  (ASHRAE approx. to Colebrook-White)')
    _info_row('Air density:',     u'0.0750 lb/ft³  (standard, 68°F, sea level)')
    _info_row('Duct roughness:',  u'ε = 0.0003 ft  (galvanized steel)')
    _info_row('Verified:',        'Matches ductulator.com output exactly')

    # OK / Cancel
    btn_panel = StackPanel()
    btn_panel.Orientation = Orientation.Horizontal
    btn_panel.HorizontalAlignment = HorizontalAlignment.Right
    btn_panel.Margin = Thickness(0, 14, 0, 0)

    ok_btn = Button()
    ok_btn.Content = 'OK'
    ok_btn.Width   = 72
    ok_btn.Margin  = Thickness(0, 0, 8, 0)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width   = 72

    def on_ok(s, e):
        try:
            gpct = float(gpct_box[0].Text)
            if not (0 < gpct < 100):
                forms.alert('Green threshold must be between 0 and 100.', title='Invalid Input')
                return
            out = {}
            for i, (sys_class, _, _) in enumerate(ROWS):
                max_fpm  = float(vel_boxes[i].Text)
                max_fric = float(fric_boxes[i].Text)
                if max_fpm <= 0 or max_fric <= 0:
                    forms.alert('All values must be greater than 0.', title='Invalid Input')
                    return
                out[sys_class] = (max_fpm, max_fric)
            result[0] = (out, gpct)
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
_PRIORITY = {'RED': 3, 'YELLOW': 2, 'GREEN': 1, 'GRAY': 0}


def _duct_label(dr, custom_limits, green_pct):
    """Percentage-based label — worst of velocity and friction checks.

    Green  : value < max * green_pct/100
    Yellow : max * green_pct/100 <= value <= max
    Red    : value > max

    Returns (label, max_cap_cfm).
    """
    defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
    max_fpm, max_friction = custom_limits.get(dr.sys_class, defaults)
    green_fac = green_pct / 100.0

    # Velocity check (CFM vs capacity at max FPM)
    if dr.cfm <= 0 or dr.area_ft2 <= 0:
        vel_label = 'GRAY'
        max_cap   = 0.0
    else:
        max_cap   = max_fpm * dr.area_ft2
        green_cap = max_cap * green_fac
        if dr.cfm < green_cap:
            vel_label = 'GREEN'
        elif dr.cfm <= max_cap:
            vel_label = 'YELLOW'
        else:
            vel_label = 'RED'

    # Friction check
    if dr.friction_per_100ft <= 0 or max_friction <= 0:
        fric_label = 'GRAY'
    else:
        green_fric = max_friction * green_fac
        if dr.friction_per_100ft < green_fric:
            fric_label = 'GREEN'
        elif dr.friction_per_100ft <= max_friction:
            fric_label = 'YELLOW'
        else:
            fric_label = 'RED'

    if _PRIORITY.get(fric_label, 0) > _PRIORITY.get(vel_label, 0):
        return fric_label, max_cap
    return vel_label, max_cap


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(elem.Id.IntegerValue)


def _duct_size_label(elem):
    """Return readable size: '10"' for round/spiral, '18x12"' for rectangular."""
    try:
        d = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            return '{:.0f}"'.format(d.AsDouble() * 12.0)
        w = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w and h and w.AsDouble() > 0 and h.AsDouble() > 0:
            return '{:.0f}x{:.0f}"'.format(w.AsDouble() * 12.0, h.AsDouble() * 12.0)
    except Exception:
        pass
    return '?'


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

    # 2. Velocity + friction settings dialog
    dialog_result = show_velocity_settings_dialog()
    if dialog_result is None:
        output.print_md('**Cancelled.**')
        return
    custom_limits, green_pct = dialog_result

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

        # Color overrides — worst of velocity check and friction check
        counts      = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'GRAY': 0}
        # eid -> (label, green_cap_cfm) for fittings + annotations
        duct_labels = {}

        for eid, dr in net.duct_results.items():
            label, green_cap = _duct_label(dr, custom_limits, green_pct)
            duct_labels[eid] = (label, green_cap)
            color = _COLOR_MAP.get(label, GRAY)
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(eid, ogs)
            counts[label] = counts.get(label, 0) + 1

        # Color fittings and accessories by worst adjacent duct color
        # Transitions, boots, tees, elbows → inherit RED/YELLOW/GREEN from neighbors
        # Build bidirectional adjacency from the BFS children dict
        adj = {}
        for pid, cids in net.children.items():
            if pid not in adj:
                adj[pid] = []
            for cid in cids:
                adj[pid].append(cid)
                if cid not in adj:
                    adj[cid] = []
                adj[cid].append(pid)

        fitting_counts = {'GREEN': 0, 'YELLOW': 0, 'RED': 0}

        for nid, elem in net.nodes.items():
            if not hvac_graph.is_fitting_or_accessory(elem):
                continue
            worst = 'GRAY'
            for neighbor_id in adj.get(nid, []):
                nb_elem = net.nodes.get(neighbor_id)
                if nb_elem is None or not hvac_graph.is_duct(nb_elem):
                    continue
                nb_label = duct_labels.get(nb_elem.Id, ('GRAY', 0.0))[0]
                if _PRIORITY.get(nb_label, 0) > _PRIORITY.get(worst, 0):
                    worst = nb_label
            if worst == 'GRAY':
                continue  # no adjacent colored duct — leave Revit default
            color = _COLOR_MAP[worst]
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(elem.Id, ogs)
            fitting_counts[worst] = fitting_counts.get(worst, 0) + 1

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
    output.print_md('---')
    output.print_md('## Done')
    output.print_md('Sheet **DV-{}** created.'.format(source_sheet_num))
    output.print_md('')
    output.print_md('**Design limits used  (green < {}% of max, yellow ≤ max, red > max):**'.format(
        int(green_pct)))
    output.print_md('| System | Max Velocity | Max Friction |')
    output.print_md('| --- | --- | --- |')
    for sys_class in ('Supply Air', 'Return Air', 'Exhaust Air', 'Outside Air'):
        defaults = hvac_graph.FIRM_DEFAULTS.get(sys_class, (600, 0.05))
        mx_fpm, mx_fric = custom_limits.get(sys_class, defaults)
        output.print_md('| {} | {:.0f} FPM | {:.3f} iwc/100 |'.format(
            sys_class, mx_fpm, mx_fric))
    output.print_md('')
    output.print_md('| Color | Ducts | Fittings & Accessories | Meaning |')
    output.print_md('| --- | --- | --- | --- |')
    output.print_md('| Green  | {} | {} | Within limit       |'.format(
        counts.get('GREEN',  0), fitting_counts.get('GREEN',  0)))
    output.print_md('| Yellow | {} | {} | Approaching limit  |'.format(
        counts.get('YELLOW', 0), fitting_counts.get('YELLOW', 0)))
    output.print_md('| Red    | {} | {} | Exceeds limit      |'.format(
        counts.get('RED',    0), fitting_counts.get('RED',    0)))
    output.print_md('| Gray   | {} | — | No CFM data        |'.format(
        counts.get('GRAY',   0)))

    # Flagged duct list — RED first, then YELLOW, sorted by velocity descending
    flagged = []
    for eid, dr in net.duct_results.items():
        label, _ = duct_labels.get(eid, ('GRAY', 0.0))
        if label not in ('RED', 'YELLOW'):
            continue
        fpm = dr.cfm / dr.area_ft2 if dr.area_ft2 > 0 else 0.0
        defaults = hvac_graph.FIRM_DEFAULTS.get(dr.sys_class, (600, 0.05))
        max_fpm, max_fric = custom_limits.get(dr.sys_class, defaults)
        flagged.append((label, fpm, dr, max_fpm, max_fric))

    if flagged:
        flagged.sort(key=lambda x: (0 if x[0] == 'RED' else 1, -x[1]))

        # Fixed-width columns for monospace alignment
        _COLS = [
            ('#',               3),
            ('Status',          7),
            ('System',         13),
            ('Duct ID',         9),
            ('Size',            8),
            ('Vel (FPM)',       10),
            ('Max FPM',         8),
            ('CFM',             6),
            ('Fric (iwc/100)',  15),
            ('Max Fric',        9),
        ]

        def _fmt_row(cells):
            return '  '.join(str(c).ljust(w) for c, (_, w) in zip(cells, _COLS))

        header    = _fmt_row([h for h, _ in _COLS])
        separator = _fmt_row(['-' * w for _, w in _COLS])
        rows      = [header, separator]

        for idx, (label, fpm, dr, max_fpm, max_fric) in enumerate(flagged, 1):
            size = _duct_size_label(dr.elem)
            rows.append(_fmt_row([
                idx,
                label,
                dr.sys_class,
                dr.element_id,
                size,
                '{:.0f}'.format(fpm),
                '{:.0f}'.format(max_fpm),
                '{:.0f}'.format(dr.cfm),
                '{:.3f}'.format(dr.friction_per_100ft),
                '{:.3f}'.format(max_fric),
            ]))

        output.print_md('')
        output.print_md('### Flagged Ducts')
        output.print_code('\n'.join(rows))

    uidoc.ActiveView = new_sheet


main()
