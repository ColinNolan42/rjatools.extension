# -*- coding: utf-8 -*-
"""
Duct Velocity.pushbutton/script.py  --  HVAC Phase 1

Colors ductwork in a copied floor plan view by velocity (green/yellow/red)
per SMACNA commercial low-velocity defaults.

Run HVAC Diagnose first to verify the network and CFM values, then run this.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys

from pyrevit import script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction,
    BuiltInCategory, ViewSheet, ViewType,
    Viewport, ViewDuplicateOption, ElementId, XYZ,
    OverrideGraphicSettings, Color
)
from Autodesk.Revit.UI.Selection import ObjectType

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import hvac_graph

# ── SMACNA colors ────────────────────────────────────────────────────────────
GREEN  = Color(0,   200,  0)
YELLOW = Color(255, 215,  0)
RED    = Color(210,  40, 40)
GRAY   = Color(160, 160, 160)

_COLOR_MAP = {'GREEN': GREEN, 'YELLOW': YELLOW, 'RED': RED, 'GRAY': GRAY}


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(elem.Id.IntegerValue)


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

    # 2. Select element
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

    # 3. Build network (traversal + CFM + velocity)
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

    # 4. Find source sheet number
    source_sheet_num = 'NoSheet'
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            if vp is not None and vp.ViewId == active_view.Id:
                source_sheet_num = sheet.SheetNumber
                break

    output.print_md('Source sheet: **{}**'.format(source_sheet_num))

    # 5. Get title block and solid fill
    tb_list = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsElementType().ToList()
    tb_id    = tb_list[0].Id if len(tb_list) > 0 else ElementId.InvalidElementId
    fill_id  = hvac_graph.solid_fill_pattern_id(doc)

    # 6. Transaction: copy view → apply overrides → create sheet
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

        # Apply color overrides
        counts = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'GRAY': 0}
        for eid, dr in net.duct_results.items():
            color = _COLOR_MAP.get(dr.label, GRAY)
            ogs   = OverrideGraphicSettings()
            ogs.SetSurfaceForegroundPatternColor(color)
            if fill_id != ElementId.InvalidElementId:
                ogs.SetSurfaceForegroundPatternId(fill_id)
            ogs.SetProjectionLineColor(color)
            new_view.SetElementOverrides(eid, ogs)
            counts[dr.label] = counts.get(dr.label, 0) + 1

        # Create output sheet
        new_sheet = ViewSheet.Create(doc, tb_id)
        new_sheet.SheetNumber = 'DV-' + source_sheet_num
        new_sheet.Name        = 'Ducting Velocities - ' + source_sheet_num

        # Place viewport
        Viewport.Create(doc, new_sheet.Id, new_vid, XYZ(1.1, 0.8, 0))

        t.Commit()
    except Exception as ex:
        t.RollBack()
        output.print_md('**Error — transaction rolled back:** {}'.format(str(ex)))
        return

    # 7. Summary
    output.print_md('---')
    output.print_md('## Done')
    output.print_md('Sheet **DV-{}** created.'.format(source_sheet_num))
    output.print_md('| Color | Count | Meaning |')
    output.print_md('| --- | --- | --- |')
    output.print_md('| Green  | {} | Within SMACNA limit |'.format(counts.get('GREEN',  0)))
    output.print_md('| Yellow | {} | Approaching limit   |'.format(counts.get('YELLOW', 0)))
    output.print_md('| Red    | {} | Exceeds limit       |'.format(counts.get('RED',    0)))
    output.print_md('| Gray   | {} | No CFM data         |'.format(counts.get('GRAY',   0)))

    uidoc.ActiveView = new_sheet


main()
