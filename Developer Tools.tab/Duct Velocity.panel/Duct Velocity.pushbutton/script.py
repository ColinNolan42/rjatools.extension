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
    """WPF dialog — one Max Velocity per system type + tolerance band.

    Returns (limits_dict, max_friction_inwc, tol_pct) or None if cancelled.
      limits_dict  = {sys_class: (max_fpm, yellow_fpm)}
      max_friction = max in. wc / 100 ft (same for all systems)
      tol_pct      = % over max before going red (applied to both vel + friction)
    """
    # Defaults: SMACNA commercial low-velocity design velocities
    ROWS = [
        ('Supply Air',  2000),
        ('Return Air',  1500),
        ('Exhaust Air', 1200),
        ('Outside Air', 1200),
    ]
    DEFAULT_TOLERANCE = 15    # % over max before going red
    DEFAULT_FRICTION  = 0.10  # in. wc / 100 ft (SMACNA equal-friction target)

    result    = [None]  # no nonlocal in Python 2.7
    max_boxes = {}      # row_idx -> TextBox
    tol_box   = [None]  # mutable ref to tolerance TextBox
    fric_box  = [None]  # mutable ref to friction TextBox

    win = Window()
    win.Title  = 'Duct Velocity Settings'
    win.Width  = 360
    win.SizeToContent = SizeToContent.Height
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    outer = StackPanel()
    outer.Margin = Thickness(14)

    intro = Label()
    intro.Content = 'Set max design velocity per system (FPM):'
    intro.Margin  = Thickness(0, 0, 0, 8)
    outer.Children.Add(intro)

    # System rows
    grid = Grid()
    for w in (160, 120):
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

    _lbl('System Type',      0, 0)
    _lbl('Max Velocity FPM', 1, 0)

    for i, (sys_class, default_fpm) in enumerate(ROWS):
        r = i + 1
        _lbl(sys_class, 0, r)
        tb = TextBox()
        tb.Text  = str(default_fpm)
        tb.Width = 80
        tb.Margin = Thickness(4, 4, 4, 4)
        tb.VerticalAlignment   = VerticalAlignment.Center
        tb.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetColumn(tb, 1)
        Grid.SetRow(tb, r)
        grid.Children.Add(tb)
        max_boxes[i] = tb

    outer.Children.Add(grid)

    # Tolerance band row
    tol_panel = StackPanel()
    tol_panel.Orientation = Orientation.Horizontal
    tol_panel.Margin = Thickness(0, 10, 0, 0)

    tol_lbl = Label()
    tol_lbl.Content = 'Yellow tolerance:'
    tol_lbl.VerticalAlignment = VerticalAlignment.Center
    tol_panel.Children.Add(tol_lbl)

    tb_tol = TextBox()
    tb_tol.Text  = str(DEFAULT_TOLERANCE)
    tb_tol.Width = 45
    tb_tol.Margin = Thickness(4, 0, 4, 0)
    tb_tol.VerticalAlignment = VerticalAlignment.Center
    tol_panel.Children.Add(tb_tol)
    tol_box[0] = tb_tol

    tol_suffix = Label()
    tol_suffix.Content = '% over max before red'
    tol_suffix.VerticalAlignment = VerticalAlignment.Center
    tol_panel.Children.Add(tol_suffix)
    outer.Children.Add(tol_panel)

    # Friction loss row
    fric_panel = StackPanel()
    fric_panel.Orientation = Orientation.Horizontal
    fric_panel.Margin = Thickness(0, 6, 0, 0)

    fric_lbl = Label()
    fric_lbl.Content = 'Max friction loss:'
    fric_lbl.VerticalAlignment = VerticalAlignment.Center
    fric_panel.Children.Add(fric_lbl)

    tb_fric = TextBox()
    tb_fric.Text  = str(DEFAULT_FRICTION)
    tb_fric.Width = 52
    tb_fric.Margin = Thickness(4, 0, 4, 0)
    tb_fric.VerticalAlignment = VerticalAlignment.Center
    fric_panel.Children.Add(tb_fric)
    fric_box[0] = tb_fric

    fric_suffix = Label()
    fric_suffix.Content = 'in. wc / 100 ft'
    fric_suffix.VerticalAlignment = VerticalAlignment.Center
    fric_panel.Children.Add(fric_suffix)
    outer.Children.Add(fric_panel)

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
            tol  = float(tol_box[0].Text)
            fric = float(fric_box[0].Text)
            if tol < 0:
                forms.alert('Tolerance must be 0 or greater.', title='Invalid Input')
                return
            if fric <= 0:
                forms.alert('Max friction loss must be greater than 0.', title='Invalid Input')
                return
            out = {}
            for i, (sys_class, _) in enumerate(ROWS):
                max_fpm    = float(max_boxes[i].Text)
                yellow_fpm = max_fpm * (1.0 + tol / 100.0)
                out[sys_class] = (max_fpm, yellow_fpm)
            result[0] = (out, fric, tol)
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


def _cfm_label(cfm, area_ft2, sys_class, custom_limits):
    """Velocity check: actual CFM vs duct capacity at green/yellow FPM limits.
    Returns (label, green_cap_cfm).
    """
    limits = custom_limits.get(sys_class, hvac_graph.SMACNA.get(sys_class, None))
    if cfm <= 0 or area_ft2 <= 0 or limits is None:
        return 'GRAY', 0.0
    green_fpm, yellow_fpm = limits
    green_cap  = green_fpm  * area_ft2
    yellow_cap = yellow_fpm * area_ft2
    if cfm <= green_cap:
        return 'GREEN', green_cap
    elif cfm <= yellow_cap:
        return 'YELLOW', green_cap
    else:
        return 'RED', green_cap


def _friction_label(friction_per_100ft, max_friction, tol_pct):
    """Friction check: actual in. wc/100ft vs design max.
    Returns label string (GREEN / YELLOW / RED / GRAY).
    """
    if friction_per_100ft <= 0 or max_friction <= 0:
        return 'GRAY'
    yellow_friction = max_friction * (1.0 + tol_pct / 100.0)
    if friction_per_100ft <= max_friction:
        return 'GREEN'
    elif friction_per_100ft <= yellow_friction:
        return 'YELLOW'
    else:
        return 'RED'


def _duct_label(dr, custom_limits, max_friction, tol_pct):
    """Combined label: worst of velocity check and friction check.
    Returns (label, green_cap_cfm).
    """
    vel_label,  green_cap = _cfm_label(dr.cfm, dr.area_ft2, dr.sys_class, custom_limits)
    fric_label             = _friction_label(dr.friction_per_100ft, max_friction, tol_pct)
    if _PRIORITY.get(fric_label, 0) > _PRIORITY.get(vel_label, 0):
        return fric_label, green_cap
    return vel_label, green_cap


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

    # 2. Velocity + friction settings dialog
    dialog_result = show_velocity_settings_dialog()
    if dialog_result is None:
        output.print_md('**Cancelled.**')
        return
    custom_limits, max_friction, tol_pct = dialog_result

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
            label, green_cap = _duct_label(dr, custom_limits, max_friction, tol_pct)
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

        # CFM annotations: "actual / capacity CFM" — skip GRAY (no CFM data)
        note_opts = TextNoteOptions()
        note_opts.HorizontalAlignment = HorizontalTextAlignment.Center
        annotation_count = 0
        annotation_errors = []
        for eid, dr in net.duct_results.items():
            label, green_cap = duct_labels.get(eid, ('GRAY', 0.0))
            if label == 'GRAY':
                continue
            mid = _duct_midpoint(dr.elem)
            if mid is None:
                annotation_errors.append('id={} midpoint=None'.format(dr.element_id))
                continue
            try:
                # Line 1: CFM comparison  Line 2: friction rate
                ann_text = '{:.0f}/{:.0f} CFM\n{:.3f} iwc/100'.format(
                    dr.cfm, green_cap, dr.friction_per_100ft)
                TextNote.Create(doc, new_vid, mid, ann_text, note_opts)
                annotation_count += 1
            except Exception as ex:
                annotation_errors.append('id={} err={}'.format(dr.element_id, str(ex)))

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
    output.print_md('Sheet **DV-{}** created  |  {} annotations placed.'.format(
        source_sheet_num, annotation_count))
    if annotation_errors:
        output.print_md('**Annotation errors ({}):**'.format(len(annotation_errors)))
        for msg in annotation_errors[:5]:   # first 5 only
            output.print_md('- `{}`'.format(msg))
    output.print_md('')
    output.print_md('**Design limits used:**')
    output.print_md('| System | Max vel (green) | Yellow ≤ | Red > |')
    output.print_md('| --- | --- | --- | --- |')
    for sys_class in ('Supply Air', 'Return Air', 'Exhaust Air', 'Outside Air'):
        mx, yw = custom_limits.get(sys_class, (0, 0))
        output.print_md('| {} | {:.0f} FPM | {:.0f} FPM | {:.0f} FPM |'.format(
            sys_class, mx, yw, yw))
    yellow_fric = max_friction * (1.0 + tol_pct / 100.0)
    output.print_md('| Friction (all) | {:.3f} iwc/100 | {:.3f} iwc/100 | {:.3f} iwc/100 |'.format(
        max_friction, yellow_fric, yellow_fric))
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

    uidoc.ActiveView = new_sheet


main()
