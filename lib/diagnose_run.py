# -*- coding: utf-8 -*-
# diagnose_run.py
# The Diagnose pipeline, shared by the Diagnose button and the headless entry
# (pyRevit agent runtime).
#
# One diagnose for all three disciplines. Pick the gas meter, the domestic
# water RPZ (or the tee on an existing cold main), or HVAC equipment/a
# duct/terminal. The discipline is auto-detected from the pick, the matching
# traversal runs, and a summary of what the model contains is printed.
# Connectivity check only: no sizes are written, no sizing table runs.
#
# The step bodies moved here verbatim from Diagnose.pushbutton/script.py;
# doc and output are passed in, alerts became a returned stop message, and
# each discipline also returns its summary as data for the headless envelope.
#
# Discipline detection, in order:
#   1. HVAC  - Mechanical Equipment, or a duct-domain connector.
#   2. WATER - a piping connector whose PipeSystemType reads DomesticColdWater
#      or DomesticHotWater.
#   3. GAS   - a piping connector that is not water (ASSUMPTION: gas is the
#      only other piping discipline supported; the printed "Detected
#      discipline" line names what was found).
#   4. Anything else - a clear error naming what to pick instead.
#
# IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.

import datetime

import shared_params
import revit_helpers
import pipe_graph
import report_generator
import hvac_graph
import water_graph
import water_checks
import water_recirc
import water_run
import water_report
from revit_helpers import eid_int


TOOL = "diagnose"
QIDS = ()

WATER_SYSTEMS = (
    shared_params.SYSTEM_DOMESTIC_COLD_WATER,
    shared_params.SYSTEM_DOMESTIC_HOT_WATER,
)


def _hvac_categories():
    from Autodesk.Revit.DB import BuiltInCategory
    return set([
        int(BuiltInCategory.OST_MechanicalEquipment),
        int(BuiltInCategory.OST_DuctCurves),
        int(BuiltInCategory.OST_DuctFitting),
        int(BuiltInCategory.OST_DuctAccessory),
        int(BuiltInCategory.OST_DuctTerminal),
        int(BuiltInCategory.OST_FlexDuctCurves),
    ])


# ============================================================================
# Button path
# ============================================================================

def run(doc, uidoc, output, forms):
    """Diagnose behind the button: header, pick, then diagnose()."""
    import ui_helpers
    _header(output)
    # the shared pick step without the popup: Diagnose takes zero input
    selected_element = ui_helpers.pick_element(
        uidoc, doc, 'Diagnose',
        'Select the gas meter, the water RPZ, or HVAC equipment/a duct/terminal',
        popup=False)
    if selected_element is None:
        output.print_md('**Cancelled.**')
        return None

    outcome = diagnose(doc, selected_element, output)
    if outcome["stop"]:
        forms.alert(outcome["stop"], title=outcome["stop_title"])
    return outcome


def _header(output):
    output.print_md('# Diagnose')
    output.print_md('Select the gas meter, the domestic water RPZ, or HVAC '
                    'equipment/a duct/terminal, then click.')
    output.print_md('---')
    ts = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
    output.print_md('Timestamp: `{}`'.format(ts))
    output.print_md('Tools loaded from `{}`'.format(revit_helpers.lib_folder()))
    revit_helpers.clear_log()


def diagnose(doc, element, output):
    """Classify the pick and run its discipline's diagnose.

    Returns dict: status ("ok" | "refused" | "error"), stop (message that
    stopped the run, or None), stop_title, discipline, detail, plus the
    discipline's summary under "summary", and errors, warnings, ready.
    """
    outcome = {"status": "ok", "stop": None, "stop_title": "Diagnose",
               "discipline": None, "detail": None, "summary": {},
               "errors": [], "warnings": [], "ready": False}

    sel_id = eid_int(element.Id)
    try:
        sel_cat = element.Category.Name
    except Exception:
        sel_cat = 'Unknown'
    output.print_md('Selected: id `{}`  |  category `{}`'.format(sel_id, sel_cat))
    outcome["selected"] = {"id": sel_id, "category": sel_cat}

    try:
        discipline, detail = classify_discipline(element)
    except ValueError as e:
        output.print_md(':cross_mark: {}'.format(str(e)))
        outcome.update({"status": "refused", "stop": str(e),
                        "errors": [str(e)],
                        "stop_title": "Diagnose - Cannot Determine Discipline"})
        return outcome

    output.print_md('Detected discipline: **{}** ({})'.format(discipline, detail))
    output.print_md('---')
    outcome["discipline"] = discipline
    outcome["detail"] = detail

    if discipline == 'GAS':
        part = diagnose_gas(doc, element, output)
    elif discipline == 'WATER':
        part = diagnose_water(doc, element, output)
    else:
        part = diagnose_hvac(doc, element, output)
    outcome.update(part)
    return outcome


# ============================================================================
# DISCIPLINE DETECTION
# ============================================================================

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
    from Autodesk.Revit.DB import Domain
    cat_id = _category_id(element)
    if cat_id in _hvac_categories():
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


def _print_pre(output, text):
    output.print_html(
        "<pre style='font-family:monospace;font-size:11px;'>{}</pre>".format(
            text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')))


def _ids(seq):
    """Element ids as ints from a list of ids, elements or (id, ...) pairs."""
    out = []
    for item in seq or []:
        if isinstance(item, tuple):
            item = item[0]
        try:
            out.append(eid_int(item.Id))
            continue
        except Exception:
            pass
        try:
            out.append(int(item))
        except Exception:
            out.append(str(item))
    return out


def _stopped(text, title):
    return {"status": "refused", "stop": text, "stop_title": title,
            "summary": {}, "errors": [text], "warnings": [], "ready": False}


# ============================================================================
# GAS SUMMARY
# ============================================================================
# The PRV keywords are the one shared list in shared_params. The isolation
# valve list mirrors One-Line.pushbutton's _ISOLATION_KW (not imported: that
# script is fragile); update both together.
_GAS_PRV_KW = shared_params.PRV_FAMILY_KEYWORDS
_GAS_ISO_KW = ("isolation", "shutoff", "shut-off", "ball valve", "gas valve")


def _gas_accessory_kind(node):
    name = (node.family_name or "").lower()
    if any(kw in name for kw in _GAS_PRV_KW):
        return "PRV"
    if any(kw in name for kw in _GAS_ISO_KW):
        return "isolation valve"
    return None


def _print_log_digest(output, graph):
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


def diagnose_gas(doc, element, output):
    output.print_md('Discipline: **GAS PIPING**')
    pipe_start = revit_helpers.is_pipe(element)
    if pipe_start:
        output.print_md(
            ':warning: A pipe was picked, not a meter. This is diagnosed as '
            'an existing building connection starting at that pipe. The '
            'existing piping, its length and its load are NOT included, and '
            'the SYSTEM ORIGIN section below describes the pipe, not a meter.')

    try:
        graph = pipe_graph.build_network(element, doc)
    except Exception as e:
        output.print_md(':cross_mark: Traversal ERROR: {}'.format(str(e)))
        return _stopped("Traversal failed:\n\n{}".format(str(e)),
                        "Diagnose - Traversal Error")

    fixtures = [n for n in graph.nodes.values() if n.is_gas_fixture]
    # Float starts: '{:.1f}' on int 0 raises in IronPython 2.7. Pipes are the
    # graph's edges, not nodes.
    total_load = sum((n.gas_load_mbh for n in fixtures), 0.0)
    pipes = list(graph.edges.values())
    fittings = [n for n in graph.nodes.values() if n.node_type == 'fitting']
    tees = [n for n in graph.nodes.values() if n.node_type == 'tee']
    total_length = sum((float(e.length_feet or 0.0) for e in pipes), 0.0)

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

    acc_rows = []
    if accessories:
        output.print_md('\n### Accessories (PRVs / Isolation Valves)')
        output.print_md('| Kind | Family | Element ID | Pipe downstream | Sizing |')
        output.print_md('| --- | --- | --- | --- | --- |')
        for n, kind in accessories:
            if kind == "PRV":
                downstream = '{:.1f} ft'.format(n.downstream_pipe_ft or 0.0)
                if n.is_midstream_prv:
                    role = 'mid-stream step down'
                else:
                    role = 'equipment PRV, ignored (within {:g} ft)'.format(
                        shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT)
            else:
                downstream, role = '-', '-'
            output.print_md('| {} | {} | {} | {} | {} |'.format(
                kind, n.family_name, n.element_id, downstream, role))
            acc_rows.append({"kind": kind, "family": n.family_name,
                             "id": n.element_id,
                             "downstream_ft": (n.downstream_pipe_ft
                                               if kind == "PRV" else None),
                             "role": role})

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
    if graph.nested_prv_ids:
        errors.append(
            "Mid-stream PRV(s) {} sit downstream of another mid-stream PRV. "
            "Size Gas supports a single step down only.".format(
                ", ".join(str(i) for i in graph.nested_prv_ids)))

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

    _print_errors_warnings(output, errors, warnings)
    _print_log_digest(output, graph)

    output.print_md('\n---')
    output.print_md('## Full Diagnostic Output — Copy and paste below this line')
    diagnostic_text = report_generator.format_diagnostic_output(graph, element)
    _print_pre(output, diagnostic_text)

    summary = {
        "pipe_start": bool(pipe_start),
        "fixtures": [{"id": n.element_id, "name": n.fixture_name,
                      "mbh": float(n.gas_load_mbh)} for n in fixtures],
        "total_load_mbh": float(total_load),
        "pipe_segments": len(pipes),
        "total_pipe_ft": float(total_length),
        "tees": len(tees),
        "fittings_other": len(fittings) - len(accessories),
        "accessories": acc_rows,
        "prvs": prv_count,
        "isolation_valves": iso_count,
        "longest_run_ft": float(longest_ft or 0.0),
        "farthest_fixture": farthest,
        "disconnected": _ids(graph.disconnected),
        "nested_prv_ids": _ids(graph.nested_prv_ids),
        "existing_connections": _ids(graph.existing_connections),
    }
    return {"summary": summary, "errors": errors, "warnings": warnings,
            "ready": not errors, "report": diagnostic_text}


# ============================================================================
# WATER SUMMARY
# ============================================================================

def diagnose_water(doc, element, output):
    from Autodesk.Revit.DB import BuiltInCategory
    output.print_md('Discipline: **DOMESTIC WATER**')
    if revit_helpers.is_pipe(element):
        output.print_md(':cross_mark: A pipe was picked, not an RPZ or tee.')
        return _stopped(
            "A pipe was picked. Pick the RPZ / backflow preventer, or on an "
            "existing building the tee on the existing cold water main.",
            "Diagnose - Water")

    # Same two start points as Size Water, told apart by category so Diagnose
    # stays zero input: a pipe fitting is the tee on an existing main,
    # anything else is the RPZ.
    upstream = None
    if _category_id(element) == int(BuiltInCategory.OST_PipeFitting):
        element, upstream, stop = water_run._tee_origin(element, doc, output)
        if stop:
            output.print_md("Stopped. Nothing was changed.")
            return _stopped(stop, "Size Water")

    try:
        revit_helpers.reset_pipe_diameter_approach()
        graph = water_graph.build_water_network(element, doc, upstream)
    except Exception as e:
        output.print_md(':cross_mark: Traversal ERROR: {}'.format(str(e)))
        return _stopped("Traversal failed:\n\n{}".format(str(e)),
                        "Diagnose - Traversal Error")

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
    _print_pre(output, '\n'.join(water_checks.format_checks(checks)))

    output.print_md('\n---')
    output.print_md('## WSFU Take-off (connectivity only, not sized)')
    wsfu_lines = water_report.wsfu_table(graph)
    _print_pre(output, '\n'.join(wsfu_lines))

    recirc_data = None
    if return_ids:
        try:
            recirc = water_recirc.analyze(graph)
            output.print_md('\n## Recirculation (default {:.0f} BTU/h per '
                            '100 ft, {:.0f} F drop)'.format(
                                shared_params.DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT,
                                shared_params.DEFAULT_RECIRC_DELTA_T_F))
            _print_pre(output, '\n'.join(water_report.recirc_table(recirc)))
            recirc_data = {"total_gpm": recirc.get("total_gpm"),
                           "circuit_setters": len(recirc.get("legs") or [])}
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

    _print_errors_warnings(output, errors, warnings)

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
    report = '\n'.join(lines)
    _print_pre(output, report)

    summary = {
        "origin_id": graph.origin_id,
        "tee_start": upstream is not None,
        "fixtures": [{"id": n.element_id, "name": n.fixture_name,
                      "type": n.type_name, "cw": float(n.cw_wsfu or 0.0),
                      "hw": float(n.hw_wsfu or 0.0),
                      "total": float(n.total_wsfu or 0.0),
                      "hwr": bool(n.hwr_active)} for n in fixtures],
        "heaters": [{"id": n.element_id, "family": n.family_name,
                     "fixtures_served": len(n.served_fixture_ids)}
                    for n in heaters],
        "pumps": len(pumps), "accessories": len(accessories),
        "fittings": len(fittings), "equipment": len(equipment),
        "pipe_segments": len(pipes), "cw_pipes": len(cw_pipes),
        "hw_pipes": len(hw_pipes), "total_pipe_ft": total_length,
        "wsfu": {"cold": total_cw_wsfu, "hot": total_hw_wsfu,
                 "total": total_wsfu},
        "hwr_fixtures": len(hwr_fixtures),
        "open_ends": [{"id": e, "system": s} for e, s in graph.open_ends],
        "loops": len(graph.loops), "unreadable": len(graph.unreadable),
        "unclassified": _ids([n.element_id for n in unknown]),
        "return_system_types": [
            {"name": c.get("name"), "detected": bool(c.get("detected")),
             "pipes": c.get("pipe_count")}
            for c in (detection.get("candidates") or [])],
        "ready_for_sizing": checks["ready_for_sizing"],
        "check_errors": [f.message for f in checks["errors"]],
        "check_warnings": [f.message for f in checks["warnings"]],
        "recirc": recirc_data,
    }
    return {"summary": summary, "errors": errors, "warnings": warnings,
            "ready": not errors, "report": report}


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


def diagnose_hvac(doc, element, output):
    output.print_md('Discipline: **HVAC**')

    net = hvac_graph.build_network(element, doc)
    if net.root is None:
        output.print_md(':cross_mark: **Could not establish traversal root. Aborting.**')
        out = _stopped("Could not establish traversal root.", "Diagnose - HVAC")
        out["stop"] = None   # the button printed it and never alerted
        return out

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

    by_system = []
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
            row = dict(counts)
            row.update({"system": st, "max_fpm": float(max_fpm)})
            by_system.append(row)

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

    flagged_ducts = []
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
            if flag:
                flagged_ducts.append({"id": dr.element_id, "size": dr.size,
                                      "system": dr.sys_class,
                                      "cfm": float(dr.cfm),
                                      "fpm": float(dr.fpm),
                                      "label": dr.label})
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

    _print_errors_warnings(output, net.errors, net.warnings)

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
    report = '\n'.join(lines)
    _print_pre(output, report)

    ready = bool(net.ready_for_visualization and not net.errors)
    output.print_md('\n---')
    if ready:
        output.print_md(':white_check_mark: **System is ready for Duct Performance.**')
    else:
        output.print_md(':cross_mark: **System is NOT ready — resolve errors above first.**')

    summary = {
        "root": {"id": eid_int(net.root.Id), "name": ahu_label,
                 "method": net.ahu_method},
        "elements": len(net.nodes), "ducts": len(net.ducts),
        "terminals": len(net.terminals),
        "equipment": len(net.equipment_nodes),
        "accessories": len(net.accessories), "fittings": len(net.fittings),
        "total_cfm": float(total_cfm),
        "zero_flow_terminals": _ids(net.zero_terminals),
        "missing_flow_terminals": _ids(net.missing_flow),
        "no_area_ducts": _ids(net.no_area_ducts),
        "velocity_by_system": by_system,
        "flagged_ducts": flagged_ducts,
    }
    return {"summary": summary, "errors": list(net.errors),
            "warnings": list(net.warnings), "ready": ready, "report": report}


# ============================================================================
# SHARED
# ============================================================================

def _print_errors_warnings(output, errors, warnings):
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
# Headless entry (read only, fits run_query)
# ============================================================================

def run_headless(doc, start_id, include_report=False, include_log=False,
                 output=None):
    """Diagnose with no UI. Returns the headless.result envelope.

    start_id: the gas meter (or first new gas pipe), the water RPZ (or the tee
    on an existing cold main), or HVAC equipment / a duct / a terminal. The
    discipline is detected from it. data: discipline, detail, selected,
    ready, summary (discipline specific), and report (the "Full Diagnostic
    Output" text) when include_report. errors / warnings are the
    discipline's blocking errors and warnings.
    """
    import headless
    recorder = headless.RecordingOutput(forward=output)
    try:
        _header(recorder)
        element = revit_helpers.get_element_by_id(doc, start_id)
        if element is None:
            return headless.result(TOOL, "refused", errors=[
                "Element {} not found in the active document.".format(
                    start_id)])
        outcome = diagnose(doc, element, recorder)
    except Exception as exc:
        return headless.result(
            TOOL, "error", errors=[revit_helpers.exc_text(exc)],
            log=recorder.log if include_log else None)

    data = {
        "discipline": outcome["discipline"],
        "detail": outcome["detail"],
        "selected": outcome.get("selected"),
        "ready": outcome["ready"],
        "summary": outcome["summary"],
    }
    if include_report:
        data["report"] = outcome.get("report")
    return headless.result(
        TOOL, outcome["status"], data=data, warnings=outcome["warnings"],
        errors=outcome["errors"],
        log=recorder.log if include_log else None)
