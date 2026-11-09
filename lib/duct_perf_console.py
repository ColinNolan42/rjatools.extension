# -*- coding: utf-8 -*-
"""
duct_perf_console.py  --  Duct Performance sections for the pyRevit output window.

Pure renderers, no Revit imports. Each takes the structured records the
Duct Performance script builds alongside its drafting view text (the same
values, computed once) and returns one HTML string made from the
lib/console_report.py builders. The script prints it with output.print_html().

Kept out of script.py so the layout can be exercised under CPython
(lib/_test_console_report.py) without Revit.

Record shapes (all text already formatted unless noted):

  system summary  {'grand': [(label, value)], 'equipment': [[unit, cfm,
                  diffusers, other]]} (by_class and types may be present; the
                  console no longer shows them)

  external static {'units': [unit], 'combined': [{'title', 'text'}],
                  'method': text, 'assumptions': [text] (optional),
                  'not_included': text}
                  unit = {'label', 'classes': [cls], 'esp': text or None,
                  'one_sided': text or None, 'checks': [(ok bool, text)]
                  (passed checks; a failed one is a correction),
                  'esp_total': float or None, 'esp_sides': [class names]}
                  cls = {'sys_class', 'airflow', 'rows': [(item, inwc float,
                  detail)], 'subtotal' float, 'safety' float, 'safety_detail',
                  'total' float, 'total_detail', 'takeoff': [[item, table,
                  qty, C, in. wc]] (optional, from takeoff_rows; rendered
                  instead of 'rows' when present), 'filter': float or None}

                  Warnings, notes, fan system effect lines and unpriced
                  fittings are not in these records: they are System
                  Corrections entries (Colin, 2026-11-02).

  VAV boxes       the dict returned by script.py _vav_box_report()

  corrections     [(group, unit label or None, kind, text)] from correction(),
                  gathered by main() over the whole run (System Corrections)

IronPython 2.7: no f-strings; floats are cast with float() before '{:.3f}'.
"""

import console_report as cr


# Relative column widths per table kind. Every table of one kind is built with
# the same widths (fixed layout), so stacked tables line up column for column.
W_EQUIPMENT = [40, 18, 12, 30]
W_TAKEOFF = [24, 42, 12, 10, 12]
W_VAV = [22, 14, 16, 16, 18, 14]
W_LIMITS = [50, 25, 25]
W_LEGEND = [12, 8, 18, 62]
NARROW = '60%'


def _f3(value):
    if value is None:
        return '-'
    return '{:.3f}'.format(float(value))


def system_summary(rec):
    """Section a: System Summary. rec['caution'], when present (VAV Box
    scope with external static), is said under the equipment table."""
    out = [cr.section('System Summary')]
    out.append(cr.keyvalue(rec.get('grand') or []))
    if rec.get('equipment'):
        out.append(cr.table(
            ['Equipment', 'Supply discharge CFM', 'Diffusers', 'Other air feeding it'],
            rec['equipment'], align=['l', 'r', 'r', 'l'], widths=W_EQUIPMENT))
    if rec.get('caution'):
        out.append(cr.callout(rec['caution'], 'warn'))
    return ''.join(out)


# VAV Box scope (Colin, 2026-11-09): external static runs from the box outlet
# to the worst diffuser only. Said wherever its total prints.
VAVBOX_ESP_LABEL = 'EXTERNAL STATIC AFTER VAV BOX'
VAVBOX_ESP_CAUTION = (
    'EXTERNAL STATIC AFTER THE VAV BOX ONLY. NOT THE TOTAL SYSTEM LOSS. Low '
    'pressure ductwork downstream of the box assumed; the medium pressure duct to '
    "the box inlet, the box drop and the unit's return and external static are not "
    'included.')


TAKEOFF_HEADERS = ['Item', 'Table / basis', 'Qty', 'C', 'in. wc']

# Row order in the take-off: duct first, fittings, taps passed, components.
_TAKEOFF_RANK = (
    ('Straight duct', 0), ('Flex duct', 1), ('Flex bend', 2),
    ('Fan system effect', 3),
    ('Tap passed', 5), ('Unit side stretch', 7),
)
_KNOWN_DROP = 'known drop'


def _rank(name, table):
    for prefix, r in _TAKEOFF_RANK:
        if name.startswith(prefix):
            return r
    if table == _KNOWN_DROP:
        return 6
    if name.startswith('End of main'):
        return 4
    return 3


def takeoff_rows(items):
    """External static take-off, the way the WSFU take-off lists fixtures:
    one row per kind of item on the index run, with its table, how many, its
    C and the pressure drop they add up to (Colin, 2026-10-23).

    items: the index run items, (label, detail, fpm, loss, (name, table, C,
    feet)). Items of the same name, table and C share a row; straight duct
    rows carry their total length in Qty. Returns [[item, table, qty, C,
    in. wc]] as text, plus the summed loss as the last element of each row
    in a parallel float list: (rows, losses)."""
    groups = {}
    order = []
    for it in items:
        if len(it) < 5 or it[4] is None:
            key = (it[0], '', None)
            ft = None
        else:
            name, table, c, ft = it[4]
            key = (name, table, None if c is None else round(float(c), 2))
        if key not in groups:
            groups[key] = [0, 0.0, 0.0]
            order.append(key)
        g = groups[key]
        g[0] += 1
        g[1] += float(it[3] or 0.0)
        g[2] += float(ft or 0.0)
    first = dict((k, i) for i, k in enumerate(order))
    order = sorted(order, key=lambda k: (_rank(k[0], k[1]), first[k]))
    rows, losses = [], []
    for key in order:
        name, table, c = key
        n, loss, ft = groups[key]
        qty = '%d (%.0f ft)' % (n, ft) if ft else str(n)
        if table == _KNOWN_DROP and n:
            ctext = '%.3f each' % (loss / n)
        else:
            ctext = '-' if c is None else '%.2f' % c
        rows.append([name, table or '-', qty, ctext, _f3(loss)])
        losses.append(loss)
    return rows, losses


def _class_rows(cls, counted, sheet=False):
    """One system class inside the unit's combined table: a heading band, its
    items, Subtotal, Safety factor and the class total. counted False marks a
    side that does not go into the unit total (the lesser of return/exhaust).

    sheet True is the drafting view form: the class total row is returned as
    {'cells': [...], 'bold': True} and its long 'worst run ends at ...'
    detail is left out (the sheet puts it in the note panel under the table,
    see esp_unit_notes, because drafting view cells do not wrap)."""
    head = cls['sys_class']
    if cls.get('airflow'):
        head = '{}, {}'.format(head, cls['airflow'])
    if not counted:
        head += '  (not counted, the larger return side governs)'
    rows = [{'band': head}]
    if cls.get('takeoff'):
        rows.extend(list(r) for r in cls['takeoff'])
    else:
        rows.extend([item, det or '', '', '', _f3(val)] for item, val, det in cls['rows'])
    if cls.get('filter') is not None:
        rows.append(['Return filter', 'known drop', '1', '-', _f3(cls['filter'])])
    rows.append(['Subtotal', '', '', '', _f3(cls['subtotal'])])
    rows.append(['Safety factor', cls.get('safety_detail') or '', '', '',
                 _f3(cls['safety'])])
    total = ['{} total'.format(cls['sys_class']),
             '' if sheet else (cls.get('total_detail') or ''),
             '', '', _f3(cls['total'])]
    rows.append({'cells': total, 'bold': True} if sheet else total)
    return rows


def esp_unit_rows(unit, sheet=False):
    """The rows of one unit's external static table, shared by the console
    (_esp_unit_table) and the drafting view: (rows, total_row). rows holds
    TAKEOFF_HEADERS lists and {'band': text} class headings (sheet True:
    see _class_rows). total_row is the TOTAL EXTERNAL STATIC PRESSURE row, or
    None when only one side was traced."""
    sides = unit.get('esp_sides')
    rows = []
    for cls in unit.get('classes') or []:
        counted = sides is None or cls['sys_class'] in sides
        rows.extend(_class_rows(cls, counted, sheet))
    total_row = None
    if unit.get('esp_total') is not None:
        total_row = [VAVBOX_ESP_LABEL if unit.get('caution')
                     else 'TOTAL EXTERNAL STATIC PRESSURE', ' + '.join(sides or []),
                     '', '', _f3(unit['esp_total'])]
    return rows, total_row


def _esp_unit_table(unit):
    """Supply and return in ONE table per unit, ending in the unit's total
    external static pressure (Colin, 2026-10-23). With only one side traced the
    table ends at that side's total and the one sided callout says so."""
    rows, total_row = esp_unit_rows(unit)
    return cr.table(TAKEOFF_HEADERS, rows, align=['l', 'l', 'r', 'r', 'r'],
                    total_row=total_row, widths=W_TAKEOFF)


def esp_unit_notes(unit, sheet=False):
    """The results said under a unit's table, as [(kind, text)] with kind
    'note', 'warn', 'ok' or 'error': a one sided total and the unit max
    check. sheet True first adds where each class's index run ends ('note'),
    which the sheet table leaves out of its total row. Every correction,
    warning and note is in System Corrections instead (Colin, 2026-11-02)."""
    out = []
    if unit.get('caution'):
        # VAV Box scope: first, ahead of everything (Colin, 2026-11-09).
        out.append(('warn', unit['caution']))
    if sheet:
        for cls in unit.get('classes') or []:
            if cls.get('total_detail'):
                out.append(('note', '{}: {}.'.format(cls['sys_class'],
                                                     cls['total_detail'])))
    if unit.get('one_sided'):
        out.append(('warn', unit['one_sided']))
    if unit.get('fan_type'):
        out.append(('note', unit['fan_type']))      # exhaust fans (2026-11-03)
    for ok, text in unit.get('checks') or []:
        out.append(('ok' if ok else 'error', text))
    return out


# Text that already names what it is keeps its own lead on the sheet.
_OWN_LEADS = ('WARNING', 'UNPRICED', 'FALLBACK', 'NOTE', 'EXTERNAL STATIC', 'LOW',
              'GRAY', 'RED')


def sheet_lines(notes):
    """[(kind, text)] as plain drafting view note lines: a warning reads
    'WARNING: ...', a failed check 'CHECK FAILED: ...', a passed one
    'CHECK: ...', a note as is."""
    lines = []
    for kind, text in notes:
        if kind == 'warn' and not text.upper().startswith(_OWN_LEADS):
            text = 'WARNING: ' + text
        elif kind == 'error':
            text = 'CHECK FAILED: ' + text
        elif kind == 'ok':
            text = 'CHECK: ' + text
        lines.append(text)
    return lines


# ── System Corrections ──────────────────────────────────────────────────────
#
# Every correction, warning and note of a run is gathered into ONE section
# under the duct table, console and sheet alike (Colin, 2026-11-02), instead
# of following the section that found it. main() collects them as it runs
# with correction(); system_corrections() and corrections_sheet() render
# them in two groups, Action required first, then Notes.

ACTION = 'action'
NOTES = 'note'
GROUP_TITLES = ((ACTION, 'Action required'), (NOTES, 'Notes'))
NO_CORRECTIONS = 'No corrections. Every judged duct is within its limits.'
NO_FLAGGED = 'No flagged ducts.'

# A line that asks the engineer to change something.
_ACTION_KINDS = ('warn', 'gray', 'red', 'error')
_ACTION_LEADS = ('WARNING', 'FALLBACK', 'UNPRICED', 'GRAY', 'RED', 'LOW', 'CHECK FAILED',
                 'CLEARANCE')


def correction_group(kind, text):
    """ACTION or NOTES for one line. Its own lead decides first: NOTE is
    informational whatever its kind (a lookup read on a table edge, a duct
    pick), WARNING / FALLBACK / UNPRICED / GRAY / RED / LOW ask for a change.
    With no lead, kind 'warn', 'gray', 'red' or 'error' is an action."""
    lead = (text or '').lstrip().upper()
    if lead.startswith('NOTE'):
        return NOTES
    if lead.startswith(_ACTION_LEADS) or kind in _ACTION_KINDS:
        return ACTION
    return NOTES


def correction(kind, text, unit=None, group=None):
    """One System Corrections entry, (group, unit label or None, kind, text).
    kind: 'note', 'warn', 'gray', 'red' (a RED plan item) or 'error' (a
    failed check). group None
    classifies it (correction_group); pass ACTION to force a line whose text
    reads NOTE but asks for a change (a clearance failure)."""
    if group is None:
        group = correction_group(kind, text)
    if group == ACTION and kind == 'note':
        kind = 'warn'
    return (group, unit, kind, text)


def correction_lines(entries):
    """The entries as {ACTION: [(kind, line)], NOTES: [(kind, line)]}.

    Each line carries the sheet lead (sheet_lines: 'WARNING: ' on a warning
    without its own lead, 'CHECK FAILED: ' on a failed check) and then its
    unit, 'RTU-1 (id 5): WARNING: ...', only when the run has more than one
    unit (Colin, 2026-11-02: one unit is named in its ESP block already). A
    line already said for the same unit is dropped. Run wide lines (no unit) come first, then each unit in the
    order it first appears; within that, the order they were found."""
    units = []
    for _g, unit, _k, _t in entries:
        if unit not in units:
            units.append(unit)
    rank = dict((u, (0 if u is None else 1, i)) for i, u in enumerate(units))
    prefix = len([u for u in units if u]) > 1
    out = {ACTION: [], NOTES: []}
    seen = set()
    ordered = sorted(enumerate(entries), key=lambda it: (rank[it[1][1]], it[0]))
    for _i, (group, unit, kind, text) in ordered:
        text = (text or '').strip()
        if not text:
            continue
        key = (unit, text)
        if key in seen:
            continue
        seen.add(key)
        line = sheet_lines([(kind, text)])[0] if group == ACTION else text
        if unit and prefix:
            line = '{}: {}'.format(unit, line)
        out[group if group in out else NOTES].append((kind, line))
    return out


def _callout_kind(kind):
    return kind if kind in ('gray', 'red', 'error') else 'warn'


def system_corrections(headers, rows, entries, status_col=None, align=None,
                       widths=None, subtitle=None):
    """Section e: System Corrections. The duct table (when any duct is
    listed), then every correction line, Action required (gray, error or
    warn callouts) then Notes (muted lines). Nothing at all: one
    NO_CORRECTIONS line, so the section is always there."""
    groups = correction_lines(entries)
    if subtitle is None:
        subtitle = '{} duct(s)'.format(len(rows))
    out = [cr.section('System Corrections', subtitle)]
    if rows:
        out.append(cr.table(headers, rows, align=align, status_col=status_col,
                            widths=widths))
    elif groups[ACTION] or groups[NOTES]:
        out.append(cr.note(NO_FLAGGED))
    else:
        out.append(cr.callout(NO_CORRECTIONS, 'ok'))
        return ''.join(out)
    for group, title in GROUP_TITLES:
        if not groups[group]:
            continue
        out.append(cr.subsection(title))
        for kind, line in groups[group]:
            out.append(cr.callout(line, _callout_kind(kind)) if group == ACTION
                       else cr.note(line))
    return ''.join(out)


def corrections_text(entries, any_rows=False):
    """The correction lines as markdown lines, for the print_md fallback
    under the duct table text."""
    groups = correction_lines(entries)
    if not (groups[ACTION] or groups[NOTES]):
        return [] if any_rows else [NO_CORRECTIONS]
    out = [] if any_rows else [NO_FLAGGED]
    for group, title in GROUP_TITLES:
        if groups[group]:
            out.append('**{}**'.format(title))
            out.extend('- ' + line for _k, line in groups[group])
    return out


def corrections_sheet(entries):
    """[(heading, [line])] for the drafting view's System Corrections block,
    Action required then Notes, empty groups left out."""
    groups = correction_lines(entries)
    return [(title.upper(), [line for _k, line in groups[group]])
            for group, title in GROUP_TITLES if groups[group]]


def summary_band_text(rec):
    """The System Summary grand totals as one line for the right end of the
    sheet's SYSTEM SUMMARY band, e.g. 'Grand total airflow 5,250 CFM |
    Diffusers 14 | Ducts 60'."""
    return ' | '.join('{} {}'.format(label, value)
                      for label, value in rec.get('grand') or [])


def external_static(rec):
    """Section b: External Static Pressure, one combined table per equipment,
    then any picked duct pairing."""
    out = [cr.section('External Static Pressure',
                      'index run, per RJA SP_LOSS_WORKSHEET')]
    for unit in rec.get('units') or []:
        out.append(cr.subsection(unit['label']))
        if unit.get('classes'):
            out.append(_esp_unit_table(unit))
        for kind, text in esp_unit_notes(unit):
            out.append(cr.note(text) if kind == 'note' else cr.callout(text, kind))
    for comb in rec.get('combined') or []:
        out.append(cr.subsection(comb['title']))
        out.append(cr.callout(comb['text'], 'info'))
    if rec.get('method'):
        out.append(cr.note(rec['method']))
    # Stated assumptions, e.g. the SWSI inlet tables used for a DWDI fan
    # (Colin, 2026-10-29): one clear line each, not buried in the method.
    for a in rec.get('assumptions') or []:
        out.append(cr.callout(a, 'info'))
    if rec.get('not_included'):
        out.append(cr.note(rec['not_included']))
    return ''.join(out)


VAV_HEADERS = ['Unit', 'Box', 'MP loss to inlet', 'with safety factor',
               'Pressure left', 'Status']


def vav_rows(rep, sheet=False):
    """One row per VAV box over every unit, VAV_HEADERS columns. sheet True
    writes the two long cell texts short ('not reached', 'not checked'); the
    sheet explains them in vav_notes."""
    rows = []
    for unit in rep.get('units') or []:
        for b in unit.get('boxes') or []:
            status = b.get('status')
            if status == 'NOT CHECKED':
                left = 'not checked' if sheet else 'not checked (no unit max entered)'
            else:
                left = _f3(b.get('left'))
            if status == 'NOT REACHED':
                loss = 'not reached' if sheet else 'not reached by the index run search'
            else:
                loss = _f3(b.get('loss'))
            rows.append([unit['label'], str(b['box']), loss, _f3(b.get('loss_sf')),
                         left, status or '-'])
    return rows


def vav_notes(rep, sheet=False):
    """The table key under the VAV table as [(kind, text)]: the unit of its
    values, and (sheet True) what the short 'not reached' and 'not checked'
    cells mean, when present. A LOW box and a diffuser on an MP main are
    corrections, not table notes: see vav_corrections (Colin, 2026-11-02)."""
    out = [('note', 'Values in in. wc.')]
    if sheet:
        statuses = set(b.get('status') for u in rep.get('units') or []
                       for b in u.get('boxes') or [])
        if 'NOT REACHED' in statuses:
            out.append(('note', 'Not reached: the index run search did not reach the box.'))
        if 'NOT CHECKED' in statuses:
            out.append(('note', 'Not checked: no unit max external static entered.'))
    return out


def vav_corrections(rep):
    """The VAV lines for System Corrections, as correction() entries: one LOW
    line per box under the minimum, and each unit's diffusers on an MP main."""
    out = []
    minimum = rep.get('minimum')
    for unit in rep.get('units') or []:
        for b in unit.get('boxes') or []:
            if b.get('status') == 'LOW':
                text = 'LOW: VAV box {}, pressure left {} in. wc'.format(
                    b['box'], _f3(b.get('left')))
                if minimum is not None:
                    text += ', under the {:.2f} in. wc minimum'.format(float(minimum))
                out.append(correction('warn', text + '.', unit['label']))
        if unit.get('mp_terminal_text'):
            out.append(correction('warn', unit['mp_terminal_text'], unit['label']))
    return out


def vav_boxes(rep):
    """Section c: VAV Boxes, one table over every unit."""
    out = [cr.section('VAV Boxes', rep.get('basis'))]
    out.append(cr.table(
        VAV_HEADERS, vav_rows(rep), align=['l', 'l', 'r', 'r', 'r', 'c'],
        status_col=5, widths=W_VAV))
    for kind, text in vav_notes(rep):
        out.append(cr.note(text) if kind == 'note' else cr.callout(text, kind))
    return ''.join(out)


def limits_and_legend(rec):
    """Section f: Design Limits and Legend.

    rec = {'roles_note', 'limits_caption', 'limits': [[system, fpm, fric]],
    'vav_basis': [(label, value)], 'notes': [(bold lead, text)],
    'legend': [[key, ducts, fittings, meaning]], 'clearance': (failed bool,
    text)}. Each clearance failure's own NOTE is a System Corrections line.
    """
    out = [cr.section('Design Limits and Legend')]
    if rec.get('roles_note'):
        out.append(cr.paragraph(rec['roles_note'][1], rec['roles_note'][0]))
    out.append(cr.table(['System', 'Max Velocity', 'Max Friction'], rec.get('limits') or [],
                        align=['l', 'r', 'r'], caption=rec.get('limits_caption'),
                        widths=W_LIMITS, width=NARROW))
    if rec.get('vav_basis'):
        out.append(cr.subsection('VAV System'))
        out.append(cr.keyvalue(rec['vav_basis']))
    for lead, text in rec.get('notes') or []:
        out.append(cr.paragraph(text, lead))
    out.append(cr.table(['Color', 'Ducts', 'Fittings, Accessories & Diffusers', 'Meaning'],
                        rec.get('legend') or [], align=['l', 'r', 'r', 'l'],
                        swatch_col=0, caption='Note legend', widths=W_LEGEND))
    if rec.get('clearance'):
        out.append(cr.callout(rec['clearance'][1],
                              'error' if rec['clearance'][0] else 'ok'))
    return ''.join(out)


def sheets_created(sheet_numbers):
    """Section g: Sheets Created."""
    out = [cr.section('Sheets Created')]
    if sheet_numbers:
        out.append(cr.table(['Sheet'], [[s] for s in sheet_numbers]))
    else:
        out.append(cr.note('No sheet was created.'))
    return ''.join(out)
