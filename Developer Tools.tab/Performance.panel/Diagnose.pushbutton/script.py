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
import tie_in
import hvac_graph
import water_graph
import water_checks
import water_recirc
import water_run
import water_report
import ui_helpers
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
# Accessory words for the summary counts, from the one classifier
# (shared_params.accessory_kind, also used by pipe_graph and the One-Line).
_GAS_ACCESSORY_WORDS = {"prv": "PRV", "iso": "isolation valve",
                        "ansul": "Ansul valve"}


def _gas_accessory_kind(node):
    return _GAS_ACCESSORY_WORDS.get(
        shared_params.accessory_kind(node.family_name))


def _print_log_digest(graph):
    """F5: the traversal SKIP / WARNING / TERMINAL / EXISTING lines and the
    helper WARNING / ERROR entries, collapsed, so nothing the walk dropped
    is silent (Rule 6)."""
    walk, helpers = pipe_graph.log_digest(graph, revit_helpers.get_log())
    if not walk and not helpers:
        output.print_md('\n**Traversal log:** no skips, stops or helper '
                        'warnings.')
        return
    lines = ['TRAVERSAL ({})'.format(len(walk))] + walk
    lines += ['', 'HELPERS ({})'.format(len(helpers))] + helpers
    body = '\n'.join(lines)
    output.print_html(
        "<details><summary><b>Traversal log: {} walk line(s), {} helper "
        "warning(s)</b></summary><pre style='font-family:monospace;"
        "font-size:11px;'>{}</pre></details>".format(
            len(walk), len(helpers),
            body.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;')))


def _gas_check_pipe_pick(element):
    """n=33/41: a pipe pick runs the one gas start check Size Gas and
    One-Line run (tie_in.check_gas_start). Returns the check dict, or None
    after reporting why the pick cannot be diagnosed as a tie in."""
    check = tie_in.check_gas_start(element, doc)
    if not check["ok"]:
        forms.alert(check["message"], title="Diagnose - Gas")
        output.print_md(':cross_mark: {}'.format(
            report_generator.md_cell(check["message"])))
        return None
    output.print_md(
        ':warning: A pipe was picked, not a meter. This is diagnosed as '
        'the TIE-IN PIPE, the first new pipe at the connection to existing. '
        'The existing piping, its length and its load are NOT included, and '
        'the meter checks are skipped. {}'.format(
            report_generator.md_cell(check["message"])))
    return check


def _gas_summary_table(graph, counts):
    fixtures = pipe_graph.gas_fixtures(graph)
    # Float starts: '{:.1f}' on int 0 raises in IronPython 2.7. Pipes are the
    # graph's edges, not nodes (no node is ever typed 'pipe').
    total_load = pipe_graph.total_connected_mbh(graph)
    pipes = list(graph.edges.values())
    total_length = sum((float(e.length_feet or 0.0) for e in pipes), 0.0)
    lr = graph.longest_run or {}
    longest_ft = float(lr.get('total_length_feet', 0.0) or 0.0)
    farthest = lr.get('farthest_fixture_name', '-')

    output.print_md('\n## Summary')
    output.print_md('| Item | Value |')
    output.print_md('| --- | --- |')
    output.print_md('| Fixtures | {} |'.format(len(fixtures)))
    output.print_md('| Total connected load | {:.1f} MBH |'.format(total_load))
    output.print_md('| Pipe segments | {} |'.format(len(pipes)))
    output.print_md('| Total pipe length | {:.1f} ft |'.format(total_length))
    output.print_md('| Tees | {} |'.format(counts["tees"]))
    output.print_md('| Fittings (other) | {} |'.format(counts["fittings_other"]))
    output.print_md('| PRVs found | {} |'.format(counts["prv_count"]))
    output.print_md('| Isolation valves found | {} |'.format(counts["iso_count"]))
    output.print_md('| Longest developed run (whole system) | {:.1f} ft |'.format(
        longest_ft))
    # n=62: each zone's run, the lengths Size Gas sizes on with the box ticked
    for label, run in report_generator.zone_run_rows(graph):
        output.print_md('| {} | {} |'.format(
            report_generator.md_cell(label),
            report_generator.md_cell(report_generator.zone_run_text(run))))
    output.print_md('| Farthest fixture | {} |'.format(
        report_generator.md_cell(farthest)))
    output.print_md('| Disconnected elements | {} |'.format(len(graph.disconnected)))
    return fixtures


def _gas_fixture_table(fixtures):
    if not fixtures:
        return
    output.print_md('\n### Fixtures')
    output.print_md('| Name | Load (MBH) | Element ID |')
    output.print_md('| --- | --- | --- |')
    for n in fixtures:
        # n=63: FIXTURE_NAME is free text, escaped for the table
        output.print_md('| {} | {:.1f} | {} |'.format(
            report_generator.md_cell(n.fixture_name),
            float(n.gas_load_mbh or 0.0), n.element_id))


def _gas_accessory_table(accessories):
    if not accessories:
        return
    output.print_md('\n### Accessories (PRVs / Isolation Valves)')
    output.print_md('| Kind | Family | Element ID | Pipe downstream | Sizing |')
    output.print_md('| --- | --- | --- | --- | --- |')
    for n, kind in accessories:
        if kind == "PRV":
            downstream = '{:.1f} ft'.format(float(n.downstream_pipe_ft or 0.0))
            if n.is_midstream_prv:
                role = 'mid-stream step down'
            else:
                role = 'equipment PRV, ignored (within {:g} ft)'.format(
                    shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT)
        else:
            downstream, role = '-', '-'
        output.print_md('| {} | {} | {} | {} | {} |'.format(
            kind, report_generator.md_cell(n.family_name), n.element_id,
            downstream, role))


def _gas_existing_appliance_table(graph):
    """Existing phase gas appliances a tie in run reached on new piping.
    Their load is counted and they are never resized."""
    rows = report_generator.existing_appliance_rows(graph)
    if not rows:
        return
    output.print_md('\n### Existing Appliances on New Piping (load counted)')
    output.print_md('| Name | Element ID |')
    output.print_md('| --- | --- |')
    for eid, name in rows:
        output.print_md('| {} | {} |'.format(
            report_generator.md_cell(name), eid))


def _gas_findings(graph, origin_connectors):
    """Errors, warnings and data errors for the summary, from the same
    report_generator.gas_validation() the paste block uses, so the summary
    and ready_for_sizing agree (n=42). Diagnose does not know whether Mid
    Stream PRV will be ticked, so nested PRVs are a warning here."""
    validation = report_generator.gas_validation(
        graph, origin_connectors, mid_stream_prv=None)
    errors = list(validation["error_messages"])
    warnings = list(validation["warning_messages"])
    if graph.non_gas_skipped:
        warnings.append(
            "{} pipe(s) on a system that is not gas ended the walk, nothing "
            "past them is counted: {}".format(
                len(graph.non_gas_skipped),
                ", ".join("{} ({})".format(i, c)
                          for i, c in graph.non_gas_skipped)))
    bare_end = pipe_graph.non_fixture_ends(graph)
    if bare_end:
        warnings.append(
            "{} end(s) are not gas fixtures (IS_GAS_FIXTURE not set), no "
            "load counted: {}".format(
                len(bare_end),
                ", ".join("{} ({})".format(n.element_id, n.family_name)
                          for n in bare_end)))
    # n=63: family, system and fixture names are model text
    esc = report_generator.md_cell
    return ([esc(e) for e in errors], [esc(w) for w in warnings],
            validation["data_errors"])


def diagnose_gas(element):
    output.print_md('Discipline: **GAS PIPING**')
    tie_in_check = None
    upstream = None
    if revit_helpers.is_pipe(element):
        tie_in_check = _gas_check_pipe_pick(element)
        if tie_in_check is None:
            return
        upstream = tie_in_check["upstream"]

    try:
        graph = pipe_graph.build_network(element, doc, upstream)
    except Exception as e:
        forms.alert("Traversal failed:\n\n{}".format(str(e)),
                    title="Diagnose - Traversal Error")
        output.print_md(':cross_mark: Traversal ERROR: {}'.format(str(e)))
        return

    counts = report_generator.gas_summary_counts(graph, _gas_accessory_kind)
    fixtures = _gas_summary_table(graph, counts)
    _gas_fixture_table(fixtures)
    _gas_accessory_table(counts["accessories"])
    _gas_existing_appliance_table(graph)

    # piping connectors only (n=30), as the paste block lists them
    origin_connectors = (None if tie_in_check is not None
                         else revit_helpers.get_piping_connectors(element))
    errors, warnings, data_errors = _gas_findings(graph, origin_connectors)
    _print_errors_warnings(errors, warnings, data_errors)
    _print_log_digest(graph)

    output.print_md('\n---')
    output.print_md('## Full Diagnostic Output — Copy and paste below this line')
    diagnostic_text = report_generator.format_diagnostic_output(
        graph, element, tie_in_check=tie_in_check)
    _print_pre(diagnostic_text)


# ============================================================================
# WATER SUMMARY
# ============================================================================
def diagnose_water(element):
    output.print_md('Discipline: **DOMESTIC WATER**')
    if revit_helpers.is_pipe(element):
        forms.alert("A pipe was picked. Pick the RPZ / backflow preventer, or "
                    "on an existing building the tee on the existing cold "
                    "water main.", title="Diagnose - Water")
        output.print_md(':cross_mark: A pipe was picked, not an RPZ or tee.')
        return

    # Same two start points as Size Water, told apart by category so Diagnose
    # stays zero input: a pipe fitting is the tee on an existing main,
    # anything else is the RPZ.
    upstream = None
    if _category_id(element) == int(BuiltInCategory.OST_PipeFitting):
        element, upstream = water_run._tee_origin(element, doc, output, forms)
        if element is None:
            return

    try:
        revit_helpers.reset_pipe_diameter_approach()
        graph = water_graph.build_water_network(element, doc, upstream)
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
    # Seeded with 0.0 so an empty list still gives a float: IronPython
    # raises on '{:.1f}'.format(0) (int).
    total_length = float(sum([n.length_feet or 0.0 for n in pipes], 0.0))

    total_cw_wsfu = float(sum([n.cw_wsfu or 0.0 for n in fixtures], 0.0))
    total_hw_wsfu = float(sum([n.hw_wsfu or 0.0 for n in fixtures], 0.0))
    total_wsfu    = float(sum([n.total_wsfu or 0.0 for n in fixtures], 0.0))
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
                n.fixture_name, n.type_name, float(n.cw_wsfu or 0.0),
                float(n.hw_wsfu or 0.0), float(n.total_wsfu or 0.0),
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

    # Return detection and the completeness check, in the same order Size
    # Water runs them, so Diagnose reads the model exactly as the sizer will.
    output.print_md('\n---')
    detection = water_graph.detect_return_system_types(graph)
    water_run._print_return_detection(output, detection)
    water_graph.apply_return_system_types(
        graph, detection["detected"] - detection["unconfirmed"])
    return_ids = set(detection["detected"])

    checks = water_checks.check_system(graph, return_system_type_ids=return_ids)
    output.print_md('\n## Completeness Check')
    _print_pre('\n'.join(water_checks.format_checks(checks)))

    # WSFU take-off, same layout as Size Water's report, built from the graph
    # alone - no sizing table, matching "check systems without running a
    # sizer".
    output.print_md('\n---')
    output.print_md('## WSFU Take-off (connectivity only, not sized)')
    wsfu_lines = water_report.wsfu_table(graph)
    _print_pre('\n'.join(wsfu_lines))

    # Circuit setter flows at the default heat loss and temperature drop.
    # Size Water and Water Performance take both from their dialog.
    if return_ids:
        try:
            recirc = water_recirc.analyze(graph)
            output.print_md('\n## Recirculation (default {:.0f} BTU/h per '
                            '100 ft, {:.0f} F drop)'.format(
                                shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT,
                                shared_params.DEFAULT_RECIRC_DELTA_T_F))
            _print_pre('\n'.join(water_report.recirc_table(recirc)))
        except Exception as e:
            output.print_md(':warning: Recirculation analysis failed: {}'.format(
                str(e)))

    errors = []
    warnings = []
    if graph.loops:
        warnings.append(
            "{} element(s) revisited (a loop in the network). HWR loops are "
            "expected here; anything else should be checked.".format(len(graph.loops)))
    if graph.unreadable:
        warnings.append("{} element(s) had unreadable connectors.".format(
            len(graph.unreadable)))
    if not checks["ready_for_sizing"]:
        errors.append("Completeness check: {} error(s), see above.".format(
            len(checks["errors"])))

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
        output.print_md(':white_check_mark: **System is ready for Duct Performance.**')
    else:
        output.print_md(':cross_mark: **System is NOT ready — resolve errors above first.**')


# ============================================================================
# SHARED
# ============================================================================
def _print_errors_warnings(errors, warnings, data_errors=None):
    """data_errors (gas, n=10/11): one line per 0 MBH fixture, listed under
    their own heading. Not blocking: the walk skips them."""
    if errors:
        output.print_md('\n---')
        output.print_md('## :cross_mark: Errors')
        for e in errors:
            output.print_md('- {}'.format(e))
    if data_errors:
        output.print_md('\n---')
        output.print_md('## :cross_mark: Data errors (0 MBH gas fixtures)')
        for d in data_errors:
            output.print_md('- {}'.format(report_generator.md_cell(d)))
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
    output.print_md('Tools loaded from `{}`'.format(revit_helpers.lib_folder()))

    revit_helpers.clear_log()

    # the shared pick step without the popup: Diagnose takes zero input
    selected_element = ui_helpers.pick_element(
        uidoc, doc, 'Diagnose',
        'Select the gas meter, the water RPZ, or HVAC equipment/a duct/terminal',
        popup=False)
    if selected_element is None:
        # n=38: a failed pick is reported, not shown as a cancel
        failure = ui_helpers.last_pick_error()
        if failure:
            output.print_md(':cross_mark: **The pick failed:** {}'.format(
                report_generator.md_cell(failure)))
        else:
            output.print_md('**Cancelled.**')
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
