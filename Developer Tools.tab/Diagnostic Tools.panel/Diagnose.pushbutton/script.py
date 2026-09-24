# -*- coding: utf-8 -*-
"""
Diagnose.pushbutton/script.py

One diagnose button for all three disciplines. Pick the gas meter, the
domestic water RPZ (backflow preventer), or HVAC equipment/a duct/terminal,
then click. The discipline is auto-detected from what was picked, the
matching traversal runs, and a summary of what the model actually contains
is printed - counts, accessories, ducts, pipes, equipment, total loads,
total CFM. This is a connectivity check, not a sizer: no sizes are written,
no sizing table runs.

Replaces the old "Gas Diagnose" (gas-only) and "HVAC Diagnose" (separate
button) with one entry point. Each discipline's underlying traversal
(pipe_graph / water_graph / hvac_graph) is unchanged - this script only adds
the discipline auto-detection and the cross-discipline summary layer on top.

Discipline detection, in order:
  1. HVAC  - the pick is Mechanical Equipment, or carries a duct-domain
     connector (duct, fitting, accessory, terminal, flex duct).
  2. WATER - the pick carries a piping connector whose PipeSystemType reads
     DomesticColdWater or DomesticHotWater (verified live in Revit 2024,
     see shared_params.py).
  3. GAS   - the pick carries a piping connector that is not water. ASSUMPTION,
     not verified against a live gas meter's PipeSystemType: gas is the only
     other piping discipline this codebase supports, so any non-water piping
     connector is read as gas. If a live gas meter reports something other
     than expected here, this branch needs revisiting - see the printed
     "Detected discipline" line, which always names what was actually found.
  4. Anything else - a clear error naming what to pick instead.

Paste output into a conversation so Claude can read the network state.
No Revit model changes are made.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import os
import sys
import datetime

from pyrevit import script, forms
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import Domain, BuiltInCategory

doc    = __revit__.ActiveUIDocument.Document
uidoc  = __revit__.ActiveUIDocument
output = script.get_output()

# Add lib/ to path
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

import shared_params
import revit_helpers
import pipe_graph
import report_generator
import hvac_graph
import water_graph
import water_report
from revit_helpers import eid_int


# ============================================================================
# DISCIPLINE DETECTION
# ============================================================================
_HVAC_CATEGORIES = set([
    int(BuiltInCategory.OST_MechanicalEquipment),
    int(BuiltInCategory.OST_DuctCurves),
    int(BuiltInCategory.OST_DuctFitting),
    int(BuiltInCategory.OST_DuctAccessory),
    int(BuiltInCategory.OST_DuctTerminal),
    int(BuiltInCategory.OST_FlexDuctCurves),
])

WATER_SYSTEMS = (
    shared_params.SYSTEM_DOMESTIC_COLD_WATER,
    shared_params.SYSTEM_DOMESTIC_HOT_WATER,
)


def _element_connectors(element):
    """All connectors on element, piping or duct domain. [] if none."""
    result = []
    for path in ('ConnectorManager', 'MEPModel'):
        try:
            manager = getattr(element, path)
            if path == 'MEPModel':
                manager = manager.ConnectorManager
            if manager is None:
                continue
            for conn in manager.Connectors:
                result.append(conn)
        except Exception:
            continue
        if result:
            break
    return result


def _category_id(element):
    try:
        return eid_int(element.Category.Id)
    except Exception:
        return None


def classify_discipline(element):
    """Return ('GAS' | 'WATER' | 'HVAC', detail_string).

    Raises ValueError with a plain message when the pick cannot be
    classified into any of the three disciplines.
    """
    cat_id = _category_id(element)
    if cat_id in _HVAC_CATEGORIES:
        return 'HVAC', 'category is a duct-domain category'

    connectors = _element_connectors(element)
    if not connectors:
        raise ValueError(
            "The selected element has no connectors, so its discipline "
            "cannot be determined. Select the gas meter, the domestic "
            "water RPZ, or an HVAC equipment/duct/terminal directly.")

    has_duct = False
    piping_systems = []
    for conn in connectors:
        try:
            domain = conn.Domain
        except Exception:
            continue
        if domain == Domain.DomainHvac:
            has_duct = True
        elif domain == Domain.DomainPiping:
            try:
                piping_systems.append(str(conn.PipeSystemType))
            except Exception:
                piping_systems.append(None)

    if has_duct:
        return 'HVAC', 'has a duct-domain connector'

    if any(s in WATER_SYSTEMS for s in piping_systems):
        return 'WATER', 'piping connector system is {}'.format(
            [s for s in piping_systems if s in WATER_SYSTEMS][0])

    if piping_systems:
        return 'GAS', ('has a piping connector, system reads {} '
                        '(not a recognized water system, assumed gas)'
                        .format(piping_systems))

    raise ValueError(
        "Could not determine a discipline from the selected element "
        "(no duct or piping connector found). Select the gas meter, the "
        "domestic water RPZ, or an HVAC equipment/duct/terminal directly.")


def _divider(char='-', width=70):
    return char * width


def _print_pre(text):
    output.print_html(
        "<pre style='font-family:monospace;font-size:11px;'>{}</pre>".format(
            text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')))


# ============================================================================
# GAS SUMMARY
# ============================================================================
# Keyword classification for THIS SUMMARY ONLY, mirrored from the family-name
# keywords One-Line.pushbutton already uses to place PRV/isolation symbols
# (_PRV_KW / _ISOLATION_KW there). Not imported from One-Line - that script is
# the most actively worked, most fragile file in the codebase, so this stays
# a read-only duplicate rather than a shared dependency. If the real keyword
# list there changes, update this list to match.
_GAS_PRV_KW = ("prv", "regulator", "regulating")
_GAS_ISO_KW = ("isolation", "shutoff", "shut-off", "ball valve", "gas valve")


def _gas_accessory_kind(node):
    name = (node.family_name or "").lower()
    if any(kw in name for kw in _GAS_PRV_KW):
        return "PRV"
    if any(kw in name for kw in _GAS_ISO_KW):
        return "isolation valve"
    return None


def diagnose_gas(element):
    output.print_md('Discipline: **GAS PIPING**')

    try:
        graph = pipe_graph.build_network(element, doc)
    except Exception as e:
        forms.alert("Traversal failed:\n\n{}".format(str(e)),
                    title="Diagnose - Traversal Error")
        output.print_md(':cross_mark: Traversal ERROR: {}'.format(str(e)))
        return

    fixtures = [n for n in graph.nodes.values() if n.is_gas_fixture]
    total_load = sum(n.gas_load_mbh for n in fixtures)
    pipes = [n for n in graph.nodes.values() if n.node_type == 'pipe']
    fittings = [n for n in graph.nodes.values() if n.node_type == 'fitting']
    tees = [n for n in graph.nodes.values() if n.node_type == 'tee']
    total_length = sum(getattr(n, 'length_feet', 0.0) for n in pipes)

    accessories = []
    for n in list(fittings) + list(tees):
        kind = _gas_accessory_kind(n)
        if kind:
            accessories.append((n, kind))
    prv_count = sum(1 for n, k in accessories if k == "PRV")
    iso_count = sum(1 for n, k in accessories if k == "isolation valve")

    lr = graph.longest_run or {}
    longest_ft = lr.get('total_length_feet', 0.0)
    farthest = lr.get('farthest_fixture_name', '-')

    output.print_md('\n## Summary')
    output.print_md('| Item | Value |')
    output.print_md('| --- | --- |')
    output.print_md('| Fixtures | {} |'.format(len(fixtures)))
    output.print_md('| Total connected load | {:.1f} MBH |'.format(total_load))
    output.print_md('| Pipe segments | {} |'.format(len(pipes)))
    output.print_md('| Total pipe length | {:.1f} ft |'.format(total_length))
    output.print_md('| Tees | {} |'.format(len(tees)))
    output.print_md('| Fittings (other) | {} |'.format(
        len(fittings) - len(accessories)))
    output.print_md('| PRVs found | {} |'.format(prv_count))
    output.print_md('| Isolation valves found | {} |'.format(iso_count))
    output.print_md('| Longest developed run | {:.1f} ft |'.format(longest_ft))
    output.print_md('| Farthest fixture | {} |'.format(farthest))
    output.print_md('| Disconnected elements | {} |'.format(len(graph.disconnected)))

    if fixtures:
        output.print_md('\n### Fixtures')
        output.print_md('| Name | Load (MBH) | Element ID |')
        output.print_md('| --- | --- | --- |')
        for n in fixtures:
            output.print_md('| {} | {:.1f} | {} |'.format(
                n.fixture_name, n.gas_load_mbh, n.element_id))

    if accessories:
        output.print_md('\n### Accessories (PRVs / Isolation Valves)')
        output.print_md('| Kind | Family | Element ID |')
        output.print_md('| --- | --- | --- |')
        for n, kind in accessories:
            output.print_md('| {} | {} | {} |'.format(kind, n.family_name, n.element_id))

    errors = []
    warnings = []
    if not fixtures:
        errors.append(
            "No gas fixtures found (IS_GAS_FIXTURE = Yes). Common causes: "
            "pipe caps replaced fixture families at terminations, fixture "
            "families not physically connected to the piping, or "
            "IS_GAS_FIXTURE not set to Yes.")
    if total_load <= 0:
        errors.append("Total load is 0 MBH.")
    if graph.longest_run is None:
        errors.append("Longest run could not be determined.")
    if graph.disconnected:
        warnings.append("{} disconnected element(s) found.".format(
            len(graph.disconnected)))

    _print_errors_warnings(errors, warnings)

    output.print_md('\n---')
    output.print_md('## Full Diagnostic Output — Copy and paste below this line')
    diagnostic_text = report_generator.format_diagnostic_output(graph, element)
    _print_pre(diagnostic_text)


# ============================================================================
# WATER SUMMARY
# ============================================================================
def diagnose_water(element):
    output.print_md('Discipline: **DOMESTIC WATER**')

    try:
        graph = water_graph.build_water_network(element, doc)
    except Exception as e:
        forms.alert("Traversal failed:\n\n{}".format(str(e)),
                    title="Diagnose - Traversal Error")
        output.print_md(':cross_mark: Traversal ERROR: {}'.format(str(e)))
        return

    nodes = list(graph.nodes.values())
    pipes       = [n for n in nodes if n.kind == water_graph.KIND_PIPE]
    fixtures    = [n for n in nodes if n.kind == water_graph.KIND_FIXTURE]
    heaters     = [n for n in nodes if n.kind == water_graph.KIND_HEATER]
    pumps       = [n for n in nodes if n.kind == water_graph.KIND_PUMP]
    accessories = [n for n in nodes if n.kind == water_graph.KIND_ACCESSORY]
    fittings    = [n for n in nodes if n.kind == water_graph.KIND_FITTING]
    equipment   = [n for n in nodes if n.kind == water_graph.KIND_EQUIPMENT]
    unknown     = [n for n in nodes if n.kind == water_graph.KIND_UNKNOWN]

    cw_pipes = [n for n in pipes if n.system == shared_params.SYSTEM_DOMESTIC_COLD_WATER]
    hw_pipes = [n for n in pipes if n.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER]
    total_length = sum(n.length_feet for n in pipes)

    total_cw_wsfu = sum(n.cw_wsfu for n in fixtures)
    total_hw_wsfu = sum(n.hw_wsfu for n in fixtures)
    total_wsfu    = sum(n.total_wsfu for n in fixtures)
    hwr_fixtures  = [n for n in fixtures if n.hwr_active]

    output.print_md('\n## Summary')
    output.print_md('| Item | Value |')
    output.print_md('| --- | --- |')
    output.print_md('| Fixtures | {} |'.format(len(fixtures)))
    output.print_md('| Water heaters | {} |'.format(len(heaters)))
    output.print_md('| Recirculation pumps | {} |'.format(len(pumps)))
    output.print_md('| Accessories (valves, RPZs, etc.) | {} |'.format(len(accessories)))
    output.print_md('| Fittings | {} |'.format(len(fittings)))
    output.print_md('| Other equipment | {} |'.format(len(equipment)))
    output.print_md('| Pipe segments | {} (CW {}, HW {}) |'.format(
        len(pipes), len(cw_pipes), len(hw_pipes)))
    output.print_md('| Total pipe length | {:.1f} ft |'.format(total_length))
    output.print_md('| Total WSFU (cold / hot / total) | {:.2f} / {:.2f} / {:.2f} |'.format(
        total_cw_wsfu, total_hw_wsfu, total_wsfu))
    output.print_md('| Fixtures with HWR active | {} |'.format(len(hwr_fixtures)))
    output.print_md('| Open ends | {} |'.format(len(graph.open_ends)))
    output.print_md('| Loops (revisited elements) | {} |'.format(len(graph.loops)))
    output.print_md('| Unreadable elements | {} |'.format(len(graph.unreadable)))
    output.print_md('| Unclassified elements | {} |'.format(len(unknown)))

    if fixtures:
        output.print_md('\n### Fixtures')
        output.print_md('| Name | Type | Cold | Hot | Total | HWR | Element ID |')
        output.print_md('| --- | --- | --- | --- | --- | --- | --- |')
        for n in fixtures:
            output.print_md('| {} | {} | {:.2f} | {:.2f} | {:.2f} | {} | {} |'.format(
                n.fixture_name, n.type_name, n.cw_wsfu, n.hw_wsfu, n.total_wsfu,
                'Yes' if n.hwr_active else 'No', n.element_id))

    if heaters:
        output.print_md('\n### Water Heaters')
        output.print_md('| Family | Fixtures served | Element ID |')
        output.print_md('| --- | --- | --- |')
        for n in heaters:
            output.print_md('| {} | {} | {} |'.format(
                n.family_name, len(n.served_fixture_ids), n.element_id))

    if accessories:
        output.print_md('\n### Accessories')
        output.print_md('| Family | System | Element ID |')
        output.print_md('| --- | --- | --- |')
        for n in accessories:
            output.print_md('| {} | {} | {} |'.format(
                n.family_name, n.system or '-', n.element_id))

    # WSFU take-off, same layout as Size Water's report, built from the graph
    # alone - no sizing table, matching "check systems without running a
    # sizer".
    output.print_md('\n---')
    output.print_md('## WSFU Take-off (connectivity only, not sized)')
    wsfu_lines = water_report.wsfu_table(graph)
    _print_pre('\n'.join(wsfu_lines))

    errors = []
    warnings = []
    if not fixtures:
        errors.append(
            "No water fixtures found (IS_WATER_FIXTURE = Yes). Check the "
            "fixtures are physically connected to the piping and the "
            "parameter is set.")
    if not heaters:
        warnings.append("No water heater found on the cold tree.")
    if graph.open_ends:
        warnings.append("{} open connector end(s) found.".format(len(graph.open_ends)))
    if graph.loops:
        warnings.append(
            "{} element(s) revisited (a loop in the network). HWR loops are "
            "expected here; anything else should be checked.".format(len(graph.loops)))
    if graph.unreadable:
        warnings.append("{} element(s) had unreadable connectors.".format(
            len(graph.unreadable)))

    _print_errors_warnings(errors, warnings)

    output.print_md('\n---')
    output.print_md('## Full Diagnostic Output — Copy and paste below this line')
    lines = []
    lines.append(_divider('='))
    lines.append('DOMESTIC WATER DIAGNOSTIC')
    lines.append('Origin (RPZ): id={}'.format(graph.origin_id))
    lines.append(_divider('='))
    lines.append('')
    lines.append('[NODES]  count={}'.format(len(nodes)))
    lines.append('{:<12} {:<10} {:<40} {:<18} {}'.format(
        'id', 'kind', 'family', 'system', 'detail'))
    lines.append(_divider())
    for n in nodes:
        detail = ''
        if n.kind == water_graph.KIND_PIPE:
            detail = '{:.2f} ft, {:.3f} in'.format(n.length_feet, n.diameter_inches)
        elif n.kind == water_graph.KIND_FIXTURE:
            detail = 'cold={:.2f} hot={:.2f} total={:.2f} hwr={}'.format(
                n.cw_wsfu, n.hw_wsfu, n.total_wsfu, n.hwr_active)
        lines.append('{:<12} {:<10} {:<40} {:<18} {}'.format(
            n.element_id, n.kind, n.family_name[:39], n.system or '-', detail))
    lines.append('')
    lines.append('[OPEN ENDS]  count={}'.format(len(graph.open_ends)))
    for eid, system in graph.open_ends:
        lines.append('  id={}  system={}'.format(eid, system))
    lines.append('')
    lines.append('[TRAVERSAL LOG]  ({} entries)'.format(len(graph.log_lines)))
    for l in graph.log_lines[:60]:
        lines.append('  ' + l)
    if len(graph.log_lines) > 60:
        lines.append('  ... ({} entries omitted) ...'.format(len(graph.log_lines) - 60))
    _print_pre('\n'.join(lines))


# ============================================================================
# HVAC SUMMARY
# ============================================================================
def _ahu_name(elem):
    try:
        return elem.Symbol.Family.Name + ' : ' + elem.Name
    except Exception:
        try:
            return elem.Name
        except Exception:
            return '(no name)'


def diagnose_hvac(element):
    output.print_md('Discipline: **HVAC**')

    net = hvac_graph.build_network(element, doc)
    if net.root is None:
        output.print_md(':cross_mark: **Could not establish traversal root. Aborting.**')
        return

    ahu_label = _ahu_name(net.root) if net.root is not element else '(selected element used as root)'
    total_cfm = sum(net.terminal_cfms.values())

    output.print_md('Root (AHU/equipment): **{}**  |  id: `{}`  |  method: `{}`'.format(
        ahu_label, eid_int(net.root.Id), net.ahu_method))

    output.print_md('\n## Summary')
    output.print_md('| Item | Value |')
    output.print_md('| --- | --- |')
    output.print_md('| Total elements | {} |'.format(len(net.nodes)))
    output.print_md('| Duct segments | {} |'.format(len(net.ducts)))
    output.print_md('| Air terminals | {} |'.format(len(net.terminals)))
    output.print_md('| Equipment nodes | {} |'.format(len(net.equipment_nodes)))
    output.print_md('| Accessories (dampers, etc.) | {} |'.format(len(net.accessories)))
    output.print_md('| Fittings (other) | {} |'.format(len(net.fittings)))
    output.print_md('| Total terminal CFM | {:.0f} |'.format(total_cfm))
    output.print_md('| Terminals with Flow = 0 | {} |'.format(len(net.zero_terminals)))
    output.print_md('| Terminals missing Flow param | {} |'.format(len(net.missing_flow)))
    output.print_md('| Ducts with no dimensions | {} |'.format(len(net.no_area_ducts)))

    # SMACNA velocity summary by system type
    sys_types = sorted(set(dr.sys_class for dr in net.duct_results.values()))
    if sys_types:
        output.print_md('\n### Velocity Check by System')
        output.print_md('| System | Green | Yellow | Red | Gray | Max FPM |')
        output.print_md('| --- | --- | --- | --- | --- | --- |')
        for st in sys_types:
            counts = {'GREEN': 0, 'YELLOW': 0, 'RED': 0, 'GRAY': 0}
            fpm_vals = []
            for dr in net.duct_results.values():
                if dr.sys_class == st:
                    counts[dr.label] = counts.get(dr.label, 0) + 1
                    if dr.fpm > 0:
                        fpm_vals.append(dr.fpm)
            max_fpm = max(fpm_vals) if fpm_vals else 0.0
            output.print_md('| {} | {} | {} | {} | {} | {:.0f} |'.format(
                st, counts['GREEN'], counts['YELLOW'], counts['RED'], counts['GRAY'], max_fpm))

    if net.terminal_cfms:
        output.print_md('\n### Air Terminals')
        _T_COLS = [('ID', 12), ('Family', 46), ('System', 16), ('CFM', 10), ('', 18)]
        def _t_row(cells):
            return '  '.join(str(c).ljust(w) for c, (_, w) in zip(cells, _T_COLS))
        t_lines = [_t_row([h for h, _ in _T_COLS]),
                   _t_row(['-' * w for _, w in _T_COLS])]
        for nid, cfm in net.terminal_cfms.items():
            elem = net.nodes[nid]
            flag = '<-- Flow = 0' if cfm <= 0 else ''
            t_lines.append(_t_row([
                nid,
                hvac_graph.terminal_family_name(elem)[:45],
                hvac_graph.terminal_sys_class(elem)[:15],
                '{:.1f} CFM'.format(cfm),
                flag,
            ]))
        output.print_code('\n'.join(t_lines))

    if net.duct_results:
        output.print_md('\n### Duct Segments')
        _D_COLS = [('ID', 12), ('Size', 9), ('System', 16),
                   ('CFM', 8), ('FPM', 8), ('SMACNA', 8), ('', 14)]
        def _d_row(cells):
            return '  '.join(str(c).ljust(w) for c, (_, w) in zip(cells, _D_COLS))
        d_lines = [_d_row([h for h, _ in _D_COLS]),
                   _d_row(['-' * w for _, w in _D_COLS])]
        for eid, dr in net.duct_results.items():
            flag = '<-- check' if dr.label in ('RED', 'GRAY') else ''
            d_lines.append(_d_row([
                dr.element_id, dr.size, dr.sys_class[:15],
                '{:.1f}'.format(dr.cfm), '{:.0f}'.format(dr.fpm), dr.label, flag,
            ]))
        output.print_code('\n'.join(d_lines))

    if net.accessories:
        output.print_md('\n### Accessories')
        output.print_md('| Family | System | Element ID |')
        output.print_md('| --- | --- | --- |')
        for n in net.accessories:
            try:
                fam = n.Symbol.Family.Name
            except Exception:
                fam = hvac_graph.terminal_family_name(n)
            output.print_md('| {} | {} | {} |'.format(
                fam, hvac_graph.terminal_sys_class(n), eid_int(n.Id)))

    _print_errors_warnings(net.errors, net.warnings)

    output.print_md('\n---')
    output.print_md('## Full Diagnostic Output — Copy and paste below this line')
    lines = []
    lines.append(_divider('='))
    lines.append('HVAC DIAGNOSTIC')
    lines.append('AHU/equipment : {}  (id={})  via {}'.format(
        ahu_label, eid_int(net.root.Id), net.ahu_method))
    lines.append(_divider('='))
    lines.append('')
    lines.append('[TRAVERSAL LOG]  ({} entries)'.format(len(net.traverse_log)))
    log_head = net.traverse_log[:30]
    log_tail = net.traverse_log[-10:] if len(net.traverse_log) > 40 else []
    for l in log_head:
        lines.append('  ' + l)
    if log_tail:
        lines.append('  ... ({} entries omitted) ...'.format(len(net.traverse_log) - 40))
        for l in log_tail:
            lines.append('  ' + l)
    _print_pre('\n'.join(lines))

    output.print_md('\n---')
    if net.ready_for_visualization and not net.errors:
        output.print_md(':white_check_mark: **System is ready for Duct Velocity visualization.**')
    else:
        output.print_md(':cross_mark: **System is NOT ready — resolve errors above first.**')


# ============================================================================
# SHARED
# ============================================================================
def _print_errors_warnings(errors, warnings):
    if errors:
        output.print_md('\n---')
        output.print_md('## :cross_mark: Errors')
        for e in errors:
            output.print_md('- {}'.format(e))
    if warnings:
        output.print_md('\n---')
        output.print_md('## :warning: Warnings')
        for w in warnings:
            output.print_md('- {}'.format(w))
    if not errors:
        output.print_md('\n---')
        output.print_md(':white_check_mark: **No blocking errors found.**')


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    output.print_md('# Diagnose')
    output.print_md('Select the gas meter, the domestic water RPZ, or HVAC '
                     'equipment/a duct/terminal, then click.')
    output.print_md('---')
    ts = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
    output.print_md('Timestamp: `{}`'.format(ts))

    revit_helpers.clear_log()

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            'Select the gas meter, the water RPZ, or HVAC equipment/a duct/terminal'
        )
        selected_element = doc.GetElement(ref.ElementId)
    except Exception:
        output.print_md('**Cancelled.**')
        return

    if selected_element is None:
        forms.alert("Could not retrieve the selected element. Please try again.",
                    title="Diagnose - Selection Error")
        return

    sel_id = eid_int(selected_element.Id)
    try:
        sel_cat = selected_element.Category.Name
    except Exception:
        sel_cat = 'Unknown'

    output.print_md('Selected: id `{}`  |  category `{}`'.format(sel_id, sel_cat))

    try:
        discipline, detail = classify_discipline(selected_element)
    except ValueError as e:
        forms.alert(str(e), title="Diagnose - Cannot Determine Discipline")
        output.print_md(':cross_mark: {}'.format(str(e)))
        return

    output.print_md('Detected discipline: **{}** ({})'.format(discipline, detail))
    output.print_md('---')

    if discipline == 'GAS':
        diagnose_gas(selected_element)
    elif discipline == 'WATER':
        diagnose_water(selected_element)
    elif discipline == 'HVAC':
        diagnose_hvac(selected_element)


main()
