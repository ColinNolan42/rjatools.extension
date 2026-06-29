# -*- coding: utf-8 -*-
"""
HVAC Diagnose.pushbutton/script.py

Diagnostic traversal for HVAC duct velocity visualization.
Run this BEFORE Duct Velocity to verify:
  - AHU is found from the selected element
  - All diffusers/terminals are reachable and have CFM assigned
  - Duct dimensions are present for velocity calculation
  - SMACNA thresholds are checked per segment

Paste output into a conversation so Claude can read the network state.
No Revit model changes are made.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys
import datetime

from pyrevit import script, forms
from Autodesk.Revit.UI.Selection import ObjectType

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import hvac_graph


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return '(no name)'


def _ahu_name(elem):
    try:
        return elem.Symbol.Family.Name + ' : ' + elem.Name
    except Exception:
        return _elem_name(elem)


def _divider(char='-', width=70):
    return char * width


def main():
    output.print_md('# HVAC Duct Velocity Diagnostic')
    output.print_md('---')
    ts = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
    output.print_md('Timestamp: `{}`'.format(ts))

    # ── STEP 1: Select element ─────────────────────────────────────────────
    output.print_md('\n## Step 1 — Select Element')
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            'Select any duct, air terminal, or AHU in the system to diagnose'
        )
    except Exception:
        output.print_md('**Cancelled.**')
        return

    sel_elem = doc.GetElement(ref.ElementId)
    sel_id   = sel_elem.Id.IntegerValue
    sel_cat  = ''
    try:
        sel_cat = sel_elem.Category.Name
    except Exception:
        sel_cat = 'Unknown'

    output.print_md('Selected: **{}**  |  id: `{}`  |  category: `{}`'.format(
        _elem_name(sel_elem), sel_id, sel_cat))

    # ── STEP 2: Build network ──────────────────────────────────────────────
    output.print_md('\n## Step 2 — Network Traversal')
    net = hvac_graph.build_network(sel_elem, doc)

    if net.root is None:
        output.print_md(':cross_mark: **Could not establish traversal root. Aborting.**')
        return

    ahu_label = _ahu_name(net.root) if net.root is not sel_elem else '(selected element used as root)'
    output.print_md('Root (AHU): **{}**  |  id: `{}`  |  method: `{}`'.format(
        ahu_label, net.root.Id.IntegerValue, net.ahu_method))
    output.print_md('Elements traversed: **{}**'.format(len(net.nodes)))

    # ── STEP 3: Print full diagnostic (copy/paste block) ─────────────────
    output.print_md('\n---')
    output.print_md('## Diagnostic Output  —  Copy and paste below this line')

    lines = []
    lines.append(_divider('='))
    lines.append('HVAC DUCT VELOCITY DIAGNOSTIC')
    lines.append('Timestamp : {}'.format(ts))
    lines.append('Selected  : id={}  cat={}'.format(sel_id, sel_cat))
    lines.append('AHU       : {}  (id={})  via {}'.format(
        ahu_label, net.root.Id.IntegerValue, net.ahu_method))
    lines.append(_divider('='))

    # Network summary
    lines.append('')
    lines.append('[NETWORK SUMMARY]')
    lines.append('  Total elements  : {}'.format(len(net.nodes)))
    lines.append('  Ducts           : {}'.format(len(net.ducts)))
    lines.append('  Air terminals   : {}'.format(len(net.terminals)))
    lines.append('  Equipment nodes : {}'.format(len(net.equipment_nodes)))
    lines.append('  Other (fittings): {}'.format(
        len(net.nodes) - len(net.ducts) - len(net.terminals) - len(net.equipment_nodes)))

    # Air terminals
    lines.append('')
    lines.append(_divider())
    lines.append('[AIR TERMINALS]  count={}'.format(len(net.terminals)))
    lines.append('{:<12} {:<45} {:<16} {:<12} {}'.format(
        'id', 'Family', 'System', 'Flow(raw)', 'CFM'))
    lines.append(_divider())
    for nid, cfm in net.terminal_cfms.items():
        elem = net.nodes[nid]
        fp   = elem.LookupParameter('Flow')
        raw  = fp.AsDouble() if fp is not None else 0.0
        flag = '  <-- Flow=0' if cfm <= 0 else ''
        flag = '  <-- Flow param MISSING' if nid in net.missing_flow else flag
        lines.append('{:<12} {:<45} {:<16} {:<12.4f} {:.2f} CFM{}'.format(
            nid,
            hvac_graph.terminal_family_name(elem)[:44],
            hvac_graph.terminal_sys_class(elem)[:15],
            raw,
            cfm,
            flag
        ))

    if not net.terminal_cfms:
        lines.append('  (none found)')

    # Duct segments
    lines.append('')
    lines.append(_divider())
    lines.append('[DUCT SEGMENTS]  count={}'.format(len(net.duct_results)))
    lines.append('{:<12} {:<12} {:<16} {:<10} {:<10} {}'.format(
        'id', 'Size', 'System', 'CFM', 'FPM', 'SMACNA'))
    lines.append(_divider())
    for eid, dr in net.duct_results.items():
        flag = '  <-- no dims' if dr.area_ft2 <= 0 else ''
        lines.append('{:<12} {:<12} {:<16} {:<10.1f} {:<10.0f} {}{}'.format(
            dr.element_id,
            dr.size,
            dr.sys_class[:15],
            dr.cfm,
            dr.fpm,
            dr.label,
            flag
        ))

    if not net.duct_results:
        lines.append('  (none found)')

    # SMACNA summary by system type
    lines.append('')
    lines.append(_divider())
    lines.append('[SMACNA VELOCITY CHECK]')
    sys_types = sorted(set(dr.sys_class for dr in net.duct_results.values()))
    for st in sys_types:
        limits = hvac_graph.SMACNA.get(st, None)
        limit_str = 'green<={} / yellow<={} / red>{}'.format(
            limits[0], limits[1], limits[1]) if limits else 'no threshold'
        counts = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'GRAY': 0}
        fpm_vals = []
        for dr in net.duct_results.values():
            if dr.sys_class == st:
                counts[dr.label] = counts.get(dr.label, 0) + 1
                if dr.fpm > 0:
                    fpm_vals.append(dr.fpm)
        max_fpm = max(fpm_vals) if fpm_vals else 0.0
        lines.append('  {} ({})'.format(st, limit_str))
        lines.append('    Green: {}  Yellow: {}  Red: {}  Gray: {}  max={:.0f} FPM'.format(
            counts['GREEN'], counts['YELLOW'], counts['RED'], counts['GRAY'], max_fpm))

    # Traversal log (abbreviated — first 30 + last 10 lines)
    lines.append('')
    lines.append(_divider())
    lines.append('[TRAVERSAL LOG]  ({} entries)'.format(len(net.traverse_log)))
    log_head = net.traverse_log[:30]
    log_tail = net.traverse_log[-10:] if len(net.traverse_log) > 40 else []
    for l in log_head:
        lines.append('  ' + l)
    if log_tail:
        lines.append('  ... ({} entries omitted) ...'.format(
            len(net.traverse_log) - 40))
        for l in log_tail:
            lines.append('  ' + l)

    # Errors and warnings
    lines.append('')
    lines.append(_divider())
    lines.append('[ERRORS]  count={}'.format(len(net.errors)))
    for e in net.errors:
        lines.append('  ERROR: ' + e)
    if not net.errors:
        lines.append('  None')

    lines.append('')
    lines.append('[WARNINGS]  count={}'.format(len(net.warnings)))
    for w in net.warnings:
        lines.append('  WARNING: ' + w)
    if not net.warnings:
        lines.append('  None')

    lines.append('')
    lines.append('[READY FOR VISUALIZATION]  {}'.format(
        'YES' if net.ready_for_visualization else 'NO'))
    lines.append(_divider('='))

    # Print as preformatted block
    text = '\n'.join(lines)
    output.print_html(
        "<pre style='font-family:monospace;font-size:11px;'>{}</pre>".format(
            text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')))

    # ── STEP 4: Summary tables ─────────────────────────────────────────────
    output.print_md('\n---')
    output.print_md('## Summary')
    output.print_md('| Item | Value |')
    output.print_md('| --- | --- |')
    output.print_md('| AHU | {} |'.format(ahu_label))
    output.print_md('| Total elements | {} |'.format(len(net.nodes)))
    output.print_md('| Duct segments | {} |'.format(len(net.duct_results)))
    output.print_md('| Air terminals | {} |'.format(len(net.terminal_cfms)))
    output.print_md('| Terminals with Flow = 0 | {} |'.format(len(net.zero_terminals)))
    output.print_md('| Terminals missing Flow param | {} |'.format(len(net.missing_flow)))
    output.print_md('| Ducts with no dimensions | {} |'.format(len(net.no_area_ducts)))

    if net.terminal_cfms:
        output.print_md('\n### Air Terminals')
        output.print_md('| ID | Family | System | CFM |')
        output.print_md('| --- | --- | --- | --- |')
        for nid, cfm in net.terminal_cfms.items():
            elem = net.nodes[nid]
            flag = ' :warning:' if cfm <= 0 else ''
            output.print_md('| {} | {} | {} | {:.1f}{} |'.format(
                nid,
                hvac_graph.terminal_family_name(elem),
                hvac_graph.terminal_sys_class(elem),
                cfm,
                flag
            ))

    if net.duct_results:
        output.print_md('\n### Duct Segments')
        output.print_md('| ID | Size | System | CFM | FPM | SMACNA |')
        output.print_md('| --- | --- | --- | --- | --- | --- |')
        for eid, dr in net.duct_results.items():
            flag = ' :warning:' if dr.label in ('RED', 'GRAY') else ''
            output.print_md('| {} | {} | {} | {:.1f} | {:.0f} | {}{} |'.format(
                dr.element_id, dr.size, dr.sys_class,
                dr.cfm, dr.fpm, dr.label, flag
            ))

    # ── STEP 5: Errors and warnings ────────────────────────────────────────
    if net.errors:
        output.print_md('\n---')
        output.print_md('## :cross_mark: Errors — Fix Before Running Duct Velocity')
        for e in net.errors:
            output.print_md('- {}'.format(e))

    if net.warnings:
        output.print_md('\n---')
        output.print_md('## :warning: Warnings')
        for w in net.warnings:
            output.print_md('- {}'.format(w))

    output.print_md('\n---')
    if net.ready_for_visualization:
        output.print_md(':white_check_mark: **System is ready for Duct Velocity visualization.**')
    else:
        output.print_md(':cross_mark: **System is NOT ready — resolve errors above first.**')


main()
