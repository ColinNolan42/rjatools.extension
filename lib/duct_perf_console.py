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
                  qty, C, in. wc]] (optional, from takeoff_rows, in route
                  order from the unit; rendered instead of 'rows' when
                  present), 'filter': float or None}

                  Warnings, notes, fan system effect lines and unpriced
                  fittings are not in these records: they are System
                  Corrections entries (Colin, 2026-11-02).

  VAV boxes       the dict returned by script.py _vav_box_report()

  corrections     [(group, unit label or None, kind, text)] from correction(),
                  gathered by main() over the whole run (System Corrections)

IronPython 2.7: no f-strings; floats are cast with float() before '{:.3f}'.
"""

import re

import ashrae_fittings
import console_report as cr


# Relative column widths per table kind. Every table of one kind is built with
# the same widths (fixed layout), so stacked tables line up column for column.
W_EQUIPMENT = [40, 18, 12, 30]
W_EQUIPMENT_3 = [56, 24, 20]       # Exhaust System: no "other air" column
W_TAKEOFF = [24, 42, 12, 10, 12]
W_VAV = [22, 14, 16, 16, 18, 14]
W_LIMITS = [50, 25, 25]
W_LEGEND = [12, 8, 18, 62]
NARROW = '60%'


def _f3(value):
    if value is None:
        return '-'
    return '{:.3f}'.format(float(value))


# ── wording per Scope (Colin, 2026-11-09) ───────────────────────────────────
# The airflow words follow the Scope that was picked, so an exhaust run never
# reads "Supply". The terminal count reads "Diffusers/Grilles" on every Scope
# (Colin, 2026-11-09, replacing the per Scope count words).
_SCOPE_WORDS = {
    'rtu': {'airflow': 'Supply CFM', 'other': 'Return / other air',
            'discharge': 'supply'},
    'vavbox': {'airflow': 'Box discharge CFM', 'other': 'Other air',
               'discharge': 'box discharge'},
    'exhaust': {'airflow': 'Exhaust CFM', 'other': None, 'discharge': 'exhaust'},
}
_SCOPE_WORDS['vav'] = _SCOPE_WORDS['rtu']
_SCOPE_WORDS['fancoil'] = _SCOPE_WORDS['rtu']   # Indoor System reads as RTU (Colin, 2026-11-09)
_COUNT_WORDS = {'count': 'Diffusers/Grilles', 'types': 'Diffuser/Grille type',
                'one': 'diffuser/grille', 'many': 'diffusers/grilles'}


def scope_words(scope):
    """The label set for a Scope ('rtu', 'vav', 'vavbox', 'exhaust'); RTU's
    for anything else. 'other' is None when the Equipment table drops its
    last column (Exhaust System: it repeated the exhaust cfm). 'scope' is
    the Scope the words are for ('rtu' for an unknown one)."""
    key = scope if scope in _SCOPE_WORDS else 'rtu'
    out = dict(_SCOPE_WORDS[key])
    out.update(_COUNT_WORDS)
    out['scope'] = key
    return out


# The count line of the System Summary, the same on every Scope (Colin,
# 2026-11-09).
TOTAL_COUNT_LABEL = 'Total Diffusers/Grilles'
_CLASS_CFM_ORDER = ('Supply Air', 'Return Air', 'Exhaust Air', 'Outside Air',
                    'Transfer Air')
_CLASS_CFM_LABEL = {'Supply Air': 'Supply CFM', 'Return Air': 'Return CFM',
                    'Exhaust Air': 'Exhaust CFM', 'Outside Air': 'Outside Air CFM',
                    'Transfer Air': 'Transfer Air CFM'}


def grand_rows(class_cfm, terminal_count, duct_count):
    """The top of the System Summary as [(label, value)] (Colin, 2026-11-09):
    one CFM line per system present (Supply, Return, then Exhaust, Outside
    Air, Transfer Air, then any other), never summed, only when above 0;
    then Total Diffusers/Grilles and Ducts. class_cfm: {system class: cfm}."""
    rows = []
    rest = sorted(k for k in class_cfm if k not in _CLASS_CFM_ORDER)
    for cls in list(_CLASS_CFM_ORDER) + rest:
        cfm = float(class_cfm.get(cls) or 0.0)
        if cfm > 0:
            rows.append((_CLASS_CFM_LABEL.get(cls, '{} CFM'.format(cls)),
                         '{:,.0f}'.format(cfm)))
    rows.append((TOTAL_COUNT_LABEL, str(terminal_count)))
    rows.append(('Ducts', str(duct_count)))
    return rows


def equipment_headers(words):
    """Equipment table headers for a Scope's words; three columns when
    words['other'] is None."""
    heads = ['Equipment', words['airflow'], words['count']]
    if words.get('other'):
        heads.append(words['other'])
    return heads


_BRANCH_LEAD = ('Branch ducts are checked against the terminal sizing tables '
                '(diffusers and grilles, MEP Design Standards, section 2), not '
                'against velocity or friction.')
_BRANCH_CHECKS = ('whether the terminal\'s own neck/face is big enough for its own '
                  'CFM, and whether the branch duct is at least as big as the '
                  'terminal connects with. Either one failing is red; there is no '
                  'yellow band on a size match. Oversize is not checked on branches.')
_BRANCH_NOT_RJA = ('Branches feeding a terminal that is not an RJA auto-sizing family',
                   '(no Sizing Warning parameter) get the main velocity/friction check.')


def branch_notes(words):
    """The branch method notes of the Design Limits and Legend section, as
    [(bold lead, text)], in terminal terms (Colin, 2026-11-09: "terminals,
    it is more accurate"). Only the sentences about the auto-sizing diffuser
    and slot families name diffusers, and Exhaust System leaves them out,
    where they cannot apply."""
    if words.get('scope') == 'exhaust':
        return [(_BRANCH_LEAD, 'Two things are checked: ' + _BRANCH_CHECKS),
                _BRANCH_NOT_RJA]
    return [
        (_BRANCH_LEAD,
         'An auto-sizing diffuser is selected on static pressure and NC rather '
         'than duct velocity, so velocity-checking a run that feeds only that one '
         'diffuser is redundant. Two things are checked instead: ' + _BRANCH_CHECKS),
        ('Branches feeding a slot diffuser',
         'are judged off the auto-sized slot family: the branch duct shall be at '
         'least the slot inlet diameter, a filled Sizing Warning on the slot is '
         'red, and the slot\'s Flow over the slot table for its length and number '
         'of slots (Titus ML-38 supply, MLR-38 return) is red, stating the '
         'table capacity, or "Max CFM reached, split slot".'),
        _BRANCH_NOT_RJA,
    ]


def clearance_text(count):
    """The clearance line of the legend, in terminal terms on every Scope
    (Colin, 2026-11-09)."""
    return 'Terminal/duct height clearance issues (flagged Red): {}'.format(count)


def legend_fittings_header(words):
    return 'Fittings, Accessories & ' + words['count']


def system_summary(rec):
    """Section a: System Summary. rec['words'] is the Scope's label set
    (scope_words). rec['caution'], when present (VAV Box scope with
    external static), is said under the equipment table."""
    words = rec.get('words') or scope_words('rtu')
    out = [cr.section('System Summary')]
    out.append(cr.keyvalue(rec.get('grand') or []))
    if rec.get('equipment'):
        heads = equipment_headers(words)
        rows = [list(r[:len(heads)]) for r in rec['equipment']]
        if len(heads) == 4:
            align, widths = ['l', 'r', 'r', 'l'], W_EQUIPMENT
        else:
            align, widths = ['l', 'r', 'r'], W_EQUIPMENT_3
        out.append(cr.table(heads, rows, align=align, widths=widths))
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


# ── Ranked run (Colin, 2026-11-10) ──────────────────────────────────────────
# A sanity check for the engineer: print the take-off of the Nth most
# restrictive run instead of the first, while the most restrictive (MR) total
# still prints and stays the design ESP.
DEFAULT_RANK = 2
MR_TAG = ' (MOST RESTRICTIVE)'
MR_LABEL = 'TOTAL EXTERNAL STATIC PRESSURE' + MR_TAG
RANKED_TOTAL_LABEL = 'Ranked runs total (not the design ESP)'
RANKED_TOTAL_NOTE = 'sanity check'


def parse_rank(text):
    """(rank int, None) or (None, error text) for a Ranked run box: a whole
    number, 2 or higher, no fixed top (a rank above the terminal count is a
    run time WARNING, see pick_ranked)."""
    try:
        v = float(str(text).strip())
    except (ValueError, TypeError):
        return None, 'must be a whole number'
    if v != v or v in (float('inf'), float('-inf')) or v != int(v):
        return None, 'must be a whole number'
    if v < 2:
        return None, 'must be 2 or higher (1 is the most restrictive run, always printed)'
    return int(v), None


def pick_ranked(runs, rank):
    """(run, rank used, run count, over) from a class leaf's sorted 'runs'.
    over is True when rank is above the run count; the last (least
    restrictive) run is then used. (None, 0, 0, False) with no runs."""
    n = len(runs or ())
    if n == 0:
        return None, 0, 0, False
    over = rank > n
    used = n if over else max(1, int(rank))
    return runs[used - 1], used, n, over


def tied_runs(runs, run):
    """The runs other than run whose total prints the same at 0.001 in. wc
    (Colin, 2026-11-11: a tie names both). runs is a class leaf's sorted
    'runs' list; run one of them (or the leaf itself)."""
    key = '{:.3f}'.format(float(run['total_inwc']))
    return [r for r in (runs or ())
            if r['terminal_id'] != run['terminal_id']
            and '{:.3f}'.format(float(r['total_inwc'])) == key]


def and_join(names):
    """'A', 'A and B', 'A, B and C'."""
    names = list(names)
    if len(names) < 2:
        return ''.join(names)
    return '{} and {}'.format(', '.join(names[:-1]), names[-1])


def mark_lines(marks, one='diffuser/grille'):
    """NOTE LEGEND lines for the plan marks: marks is the set of labels
    placed ('R', or 'R1' with 'R2', 'R3' ... under Ranked run)."""
    lines = []
    if 'R' in marks:
        lines.append('(R): Most restrictive {}. It ends the index run that sets '
                     'the required external static pressure for its equipment '
                     'or system; that run is outlined in bold orange on the plan.'.format(one))
    if 'R1' in marks:
        lines.append('(R1): Most restrictive {}. It ends the index run that sets '
                     'the required external static pressure for its equipment '
                     'or system; that run is outlined in bold orange on the plan.'.format(one))
    for m in sorted((m for m in marks if m not in ('R', 'R1')),
                    key=lambda s: int(s[1:])):
        lines.append('({}): {} most restrictive {}, the ranked run shown in the '
                     'external static take-off.'.format(m, _ordinal(int(m[1:])), one))
    return lines


def _ordinal(n):
    if 10 <= n % 100 <= 20:
        suf = 'th'
    else:
        suf = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return '{}{}'.format(n, suf).capitalize()


def rank_over_text(side, rank, count):
    """The WARNING when a Ranked run box is above the terminals on a side."""
    return ('{} rank {} is above the {} diffuser(s)/grille(s) on this system; '
            'run {} (the least restrictive) is shown.'.format(side, rank, count, count))


TAKEOFF_HEADERS = ['Item', 'Table / basis', 'Qty', 'C', 'in. wc']

# The take-off follows the ROUTE of the run from the unit out to the terminal
# (Colin, 2026-11-11): fan system effect rows first, each with the rows its
# group lists at 0.000 under it, then every item in the order the run meets
# it, one row each, with straight duct between them merged into one row.
# The return filter goes right under the fan system effect rows
# (_class_rows); an inline fan's discharge follows the intake route.
_IN_FSE = ', in the fan system effect'
_KNOWN_DROP = 'known drop'
_DUCT_ROWS = ('Straight duct', 'Flex duct', 'Discharge duct')
_ITEM_ID = re.compile(r'^(?!Fan system effect).* (\d+)$')


def _is_fse(it):
    return (it[0] or '').startswith('Fan system effect')


def _in_fse(it):
    t4 = it[4] if len(it) > 4 else None
    return bool(t4 and (t4[1] or '').endswith(_IN_FSE)) or (it[1] or '').endswith(_IN_FSE)


def _is_discharge(it):
    return (it[0] or '').startswith('Discharge ')


def _is_backdraft(it):
    t4 = it[4] if len(it) > 4 else None
    return it[0] == 'Accessory (assumed)' and bool(t4) and t4[0] == 'Backdraft damper'


def _route_order(items):
    """The items in take-off order (see above). items come from the walk in
    route order already, unit first; only the fan system effect group, the
    assumed backdraft damper (at the fan), the inline fan discharge and its
    roof cap move."""
    fse = [it for it in items if _is_fse(it)]
    grp = [it for it in items if not _is_fse(it) and _in_fse(it)]
    rest = [it for it in items if not _is_fse(it) and not _in_fse(it)]
    dis = [it for it in rest if _is_discharge(it)]
    backdraft = [it for it in rest if _is_backdraft(it)]
    cap = [it for it in rest if it[0] == 'Exhaust roof cap']
    route = [it for it in rest if not _is_discharge(it) and not _is_backdraft(it)
             and it[0] != 'Exhaust roof cap']
    head = []
    for i, f in enumerate(fse):
        head.append(f)
        last = i == len(fse) - 1
        if len(fse) == 1:
            head.extend(grp)
        elif i == 0:
            head.extend(g for g in grp if not _is_discharge(g))
        elif last:
            head.extend(g for g in grp if _is_discharge(g))
    if not fse:
        head.extend(grp)
    if dis:
        return head + route + backdraft + dis + cap
    return head + backdraft + route + cap


def _with_id(table, label):
    m = _ITEM_ID.match(label or '')
    if m is None:
        return table
    if table.endswith(_IN_FSE):
        return '{}, id {}{}'.format(table[:-len(_IN_FSE)], m.group(1), _IN_FSE)
    return '{}, id {}'.format(table, m.group(1))


def takeoff_rows(items):
    """External static take-off in route order (Colin, 2026-11-11): one row
    per fitting, accessory and known drop in the order the run meets it from
    the unit, its element id in Table / basis, and the straight duct between
    them as one row (consecutive pieces merged, length and loss summed).

    items: the run items, (label, detail, fpm, loss, (name, table, C, feet)).
    Returns (rows, losses): rows [[item, table, qty, C, in. wc]] as text and
    each row's loss as a float, so sum(losses) is the sum of the items."""
    rows, losses = [], []
    duct = None          # [name, in group, feet, loss] of the open duct row

    def close():
        if duct is not None:
            basis = 'Darcy-Weisbach, %.1f ft' % duct[2]
            if duct[1]:
                basis += _IN_FSE
            rows.append([duct[0], basis, '', '-', _f3(duct[3])])
            losses.append(duct[3])

    for it in _route_order(items):
        loss = float(it[3] or 0.0)
        t4 = it[4] if len(it) > 4 else None
        if t4 is not None and t4[0] in _DUCT_ROWS and not _is_fse(it):
            grp = _in_fse(it)
            if duct is not None and duct[0] == t4[0] and duct[1] == grp:
                duct[2] += float(t4[3] or 0.0)
                duct[3] += loss
                continue
            close()
            duct = [t4[0], grp, float(t4[3] or 0.0), loss]
            continue
        close()
        duct = None
        if t4 is None:
            name, table, c = it[0], '', None
        else:
            name, table, c = t4[0], t4[1] or '', t4[2]
        if not _is_fse(it):
            if it[0] == 'Accessory (assumed)':
                table = '{}, assumed'.format(table or _KNOWN_DROP)
            elif it[0] == 'Tap passed' and (it[1] or '').startswith('on duct '):
                table = '{}, {}'.format(table, it[1].split(',')[0])
            else:
                table = _with_id(table, it[0])
        ctext = '-' if c is None else '%.2f' % float(c)
        rows.append([name, table or '-', '1', ctext, _f3(loss)])
        losses.append(loss)
    close()
    return rows, losses


def _fse_head(rows):
    """How many take-off rows lead with the fan system effect: its rows and
    the rows its group lists at 0.000 under it."""
    n = 0
    for r in rows:
        if isinstance(r, dict):
            break
        if r[0].startswith('Fan system effect') or (r[1] or '').endswith(_IN_FSE):
            n += 1
        else:
            break
    return n


def _class_rows(cls, counted, sheet=False):
    """One system class inside the unit's combined table: a heading band, its
    items and the class subtotal (no safety factor here, esp_unit_rows adds it
    once). counted False marks a side that does not go into the unit total
    (the lesser of return/exhaust).

    sheet True is the drafting view form: the class subtotal row is returned as
    {'cells': [...], 'bold': True} and its long 'worst run ends at ...'
    detail is left out (the sheet puts it in the note panel under the table,
    see esp_unit_notes, because drafting view cells do not wrap)."""
    head = cls['sys_class']
    if cls.get('airflow'):
        head = '{}, {}'.format(head, cls['airflow'])
    if cls.get('rank_text'):
        # Ranked run (Colin, 2026-11-10): which run this take-off is.
        head = '{}, {}'.format(head, cls['rank_text'])
    if not counted:
        head += '  (not counted, the larger return side governs)'
    rows = [{'band': head}]
    body = ([list(r) for r in cls['takeoff']] if cls.get('takeoff') else
            [[item, det or '', '', '', _f3(val)] for item, val, det in cls['rows']])
    if cls.get('filter') is not None:
        # At the unit end of the return, right under the fan system effect
        # rows, since the filter sits at the unit (Colin, 2026-11-11).
        at = _fse_head(body) if cls.get('takeoff') else len(body)
        body.insert(at, ['Return filter', 'known drop', '1', '-', _f3(cls['filter'])])
    rows.extend(body)
    # The class ends at its subtotal. The safety factor is applied once, at
    # the end of the unit's table (Colin, 2026-11-12).
    sub = ['{} subtotal{}'.format(cls['sys_class'],
                                  ', run #{}'.format(cls['rank']) if cls.get('rank') else ''),
           '' if sheet else (cls.get('total_detail') or ''),
           '', '', _f3(cls['subtotal'])]
    rows.append({'cells': sub, 'bold': True} if sheet else sub)
    return rows


def esp_unit_rows(unit, sheet=False):
    """The rows of one unit's external static table, shared by the console
    (_esp_unit_table) and the drafting view: (rows, total_row). rows holds
    TAKEOFF_HEADERS lists and {'band': text} class headings (sheet True:
    see _class_rows). total_row is the TOTAL EXTERNAL STATIC PRESSURE row, or
    None when only one side was traced.

    The total row is the most restrictive (MR) total, the design ESP, marked
    MR (Colin, 2026-11-10). With Ranked run on (unit['ranked']), the classes
    hold the ranked runs: their own total goes in a 'Ranked runs total' row
    just above it, and the MR row carries unit['mr_total'] and the terminal
    each MR run ends at (unit['mr_text'])."""
    sides = unit.get('esp_sides')
    rows = []
    for cls in unit.get('classes') or []:
        counted = sides is None or cls['sys_class'] in sides
        rows.extend(_class_rows(cls, counted, sheet))
    total_row = None
    counted_cls = [c for c in unit.get('classes') or []
                   if sides is None or c['sys_class'] in sides]
    esp_total = unit.get('esp_total')
    if counted_cls:
        # Safety factor once: sides' subtotals, a combined Subtotal when two
        # sides, then the one factor row. Totals are the unit's own numbers,
        # so the factor is total less subtotal, never recomputed.
        names = [c['sys_class'] for c in counted_cls]
        sub = sum(c['subtotal'] for c in counted_cls)
        tot = (float(esp_total) if esp_total is not None
               else sum(c['total'] for c in counted_cls))
        det = next((c['safety_detail'] for c in counted_cls
                    if c.get('safety_detail')), '')
        if len(counted_cls) > 1:
            rows.append(['Subtotal', ' + '.join(names), '', '', _f3(sub)])
        rows.append(['Safety factor', det, '', '', _f3(tot - sub)])
        if esp_total is None:
            # One side only, no unit total row: the side total ends the table.
            last = ['{} total'.format(' + '.join(names)), '', '', '', _f3(tot)]
            rows.append({'cells': last, 'bold': True} if sheet else last)
    if unit.get('esp_total') is not None:
        mr_label = (VAVBOX_ESP_LABEL + MR_TAG if unit.get('caution') else MR_LABEL)
        if unit.get('ranked'):
            ranked = [RANKED_TOTAL_LABEL, '{}, {}'.format(
                unit.get('ranked_text') or ' + '.join(sides or []), RANKED_TOTAL_NOTE),
                '', '', _f3(unit['esp_total'])]
            # Plain, so the MOST RESTRICTIVE row below it is the one that
            # stands out (Colin, 2026-11-11).
            rows.append(ranked)
            total_row = [mr_label, unit.get('mr_text') or ' + '.join(sides or []),
                         '', '', _f3(unit.get('mr_total'))]
        else:
            # mr_text names where each most restrictive run ends, ties
            # included (Colin, 2026-11-11).
            total_row = [mr_label, unit.get('mr_text') or ' + '.join(sides or []),
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


# ── grouped report lines (Colin, 2026-11-11) ────────────────────────────────
# One line per (table, reason) for the tap FALLBACKs of a unit, one NOTE for
# every lookup read at a table edge, one for every blank cell. Lines that
# say the same thing once.
_SEG_SPLIT = re.compile(r'; (?=(?:FALLBACK|NOTE): )')
_FB_NOT_POS = re.compile(r'^FALLBACK: tap (\d+) (?:passed on|into branch off) duct (\d+), '
                         r'(2001 \S+ Cc,[bs]) -?\d+(?:\.\d+)? is not positive; '
                         r'(RJA worksheet C \d+(?:\.\d+)? used)$')
_FB_WHY = re.compile(r'^FALLBACK: tap (\d+) passed on duct (\d+), '
                     r'(RJA worksheet C \d+(?:\.\d+)?) \((.+)\)$')
_PART_SPLIT = re.compile(r'; (?=(?:transition|elbow|tap|tee|cross|wye|split|fitting|'
                         r'offset|duct) \d)')
_LOCATED = re.compile(r'^(\w+ \d+)[^:;]*: (.+)$')
EDGE_LEAD = 'NOTE: read at the table edge: '
BLANK_LEAD = 'NOTE: read next to a blank cell, nearest printed value used: '


def _fallback_line(tail, pairs):
    """'FALLBACK: taps 1, 2 (duct 9), 3 (duct 8): <tail>'."""
    ducts, taps = [], {}
    for duct, tap in pairs:
        if duct not in ducts:
            ducts.append(duct)
        if tap not in taps.setdefault(duct, []):
            taps[duct].append(tap)
    n = sum(len(v) for v in taps.values())
    named = ', '.join('%s (duct %s)' % (', '.join(sorted(taps[d], key=int)), d)
                      for d in ducts)
    return 'FALLBACK: %s %s: %s' % ('taps' if n > 1 else 'tap', named, tail)


def group_lines(lines):
    """The report lines of one unit with repeats grouped: the tap FALLBACKs
    one line per (table, reason), every table edge read in one NOTE, every
    blank cell in one NOTE, a line said twice once. A line holding a
    FALLBACK and a NOTE is split into the two. Order: where each first
    appears."""
    slots, fb, edges, blanks = [], {}, [], []

    def _slot(key):
        if key not in slots:
            slots.append(key)
    for line in lines:
        for seg in _SEG_SPLIT.split((line or '').strip()):
            seg = seg.strip()
            if not seg:
                continue
            m = _FB_NOT_POS.match(seg)
            if m:
                key = ('fb', '%s not positive; %s' % (m.group(3), m.group(4)))
                fb.setdefault(key, []).append((m.group(2), m.group(1)))
                _slot(key)
                continue
            m = _FB_WHY.match(seg)
            if m:
                key = ('fb', '%s (%s)' % (m.group(3), m.group(4)))
                fb.setdefault(key, []).append((m.group(2), m.group(1)))
                _slot(key)
                continue
            if seg.startswith('NOTE: '):
                rest = []
                for part in _PART_SPLIT.split(seg[len('NOTE: '):]):
                    lm = _LOCATED.match(part)
                    if lm and part.endswith(ashrae_fittings.EDGE_TAG):
                        item = '%s %s' % (lm.group(1),
                                          lm.group(2)[:-len(ashrae_fittings.EDGE_TAG)])
                        if item not in edges:
                            edges.append(item)
                        _slot(('edge',))
                    elif lm and part.endswith(ashrae_fittings.BLANK_TAG):
                        item = '%s %s' % (lm.group(1),
                                          lm.group(2)[:-len(ashrae_fittings.BLANK_TAG)])
                        if item not in blanks:
                            blanks.append(item)
                        _slot(('blank',))
                    else:
                        rest.append(part)
                if not rest:
                    continue
                seg = 'NOTE: ' + '; '.join(rest)
            _slot(seg)
    out = []
    for key in slots:
        if key == ('edge',):
            out.append(EDGE_LEAD + '; '.join(edges))
        elif key == ('blank',):
            out.append(BLANK_LEAD + '; '.join(blanks))
        elif isinstance(key, tuple):
            out.append(_fallback_line(key[1], fb[key]))
        else:
            out.append(key)
    return out


def _grouped_entries(ordered):
    """correction entries, already in report order, with each unit's lines
    grouped (group_lines). An unchanged line keeps its own group and kind;
    a grouped or split one is classified again."""
    out, i = [], 0
    while i < len(ordered):
        unit = ordered[i][1]
        j = i
        while j < len(ordered) and ordered[j][1] == unit:
            j += 1
        block = ordered[i:j]
        first = {}
        for e in block:
            first.setdefault((e[3] or '').strip(), e)
        for text in group_lines([e[3] for e in block]):
            e = first.get(text)
            if e is None:
                kind = 'warn' if text.startswith(('FALLBACK', 'UNPRICED')) else 'note'
                e = (correction_group(kind, text), unit, kind, text)
            out.append(e)
        i = j
    return out


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
    ordered = [e for _i, e in sorted(enumerate(entries),
                                     key=lambda it: (rank[it[1][1]], it[0]))]
    for group, unit, kind, text in _grouped_entries(ordered):
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


# ── plain language System Corrections (Colin, 2026-11-23) ───────────────────
# The engine's lines name element ids, ASHRAE table codes and ratios: right
# for tracing, hard to act on. Each line is shown as one plain sentence
# (what is wrong, then what to do) in one of three groups:
#   Fix the model   the tool could not read something (Flow, a size, a link)
#   Fix the design  something fails the RJA standard
#   For information an estimate or assumption, no change needed
# The SHEET shows only the plain sentences, no ids or table codes (the
# colored plan shows where), repeats said once, and For information as one
# line. The CONSOLE shows each sentence with the engine's own lines under it
# as the details. The engine's text is unchanged, so its tests still hold.

MODEL = 'model'
DESIGN = 'design'
INFO = 'info'
PLAIN_TITLES = ((MODEL, 'Fix the model'), (DESIGN, 'Fix the design'),
                (INFO, 'For information'))
ESTIMATE_SHEET_LINE = ('Some fitting losses are estimated or read at the edge of the '
                       'ASHRAE tables. No change needed; see the console for each one.')
_LEAD_RE = re.compile(r'^(?:WARNING|FALLBACK|UNPRICED|NOTE|GRAY|RED|LOW|CHECK FAILED|'
                      r'CHECK|CLEARANCE|ERROR|ASSUMPTION)\s*:\s*')
_UNIT_ID_RE = re.compile(r'\s*\(id \d+\)')
_WHERE_NOISE = re.compile(r'\s+at element \d+|, [\d.]+ in\. roof curb included')
_NUM = r'([\d.]+)'


def _strip_leads(text):
    """The text without its lead words, 'UNPRICED: WARNING: x' -> 'x'."""
    t = (text or '').strip()
    while True:
        m = _LEAD_RE.match(t)
        if not m:
            return t
        t = t[m.end():]


def _count(n, one, many):
    return '1 {}'.format(one) if int(n) == 1 else '{} {}'.format(n, many)


def _gray_headline(reason):
    """(group, sentence) for a GRAY duct by its reason."""
    r = reason.strip().lower()
    if r.startswith('no airflow'):
        return MODEL, ('Some ducts could not be checked (gray) because no airflow reaches '
                       'them. Enter Flow on every diffuser and grille, check the ductwork '
                       'connects, then rerun.')
    if r.startswith('no duct size'):
        return MODEL, ('Some ducts could not be checked (gray) because they have no size. '
                       'Give them a size and rerun.')
    if r.startswith('no table for') and r.endswith(('outside air', 'transfer air')):
        return INFO, ('Outside air and transfer air terminals are not checked (gray); '
                      'there is no RJA table for them.')
    if r.startswith(('cannot classify', 'no table for')):
        return MODEL, ('Some branches could not be checked (gray) because their diffuser '
                       'or grille could not be identified. Check the terminal\'s duct '
                       'connector and System Classification.')
    if r.startswith('cannot check') and 'not in table' in r:
        return MODEL, ('Some branches could not be checked (gray) because the diffuser or '
                       'grille size is not in the RJA sizing table. Check the terminal\'s '
                       'Type size.')
    if r.startswith('cannot check') and 'unreadable' in r:
        return MODEL, ('Some branches could not be checked (gray) because a branch duct or '
                       'diffuser size could not be read. Check the branch duct and the '
                       'diffuser connection sizes.')
    if r.startswith('cannot check') and (' on round' in r or ' on rect' in r):
        return MODEL, ('Some branches could not be checked (gray) because the branch duct '
                       'shape does not match the diffuser connection. Match the branch duct '
                       'shape to the diffuser neck.')
    return MODEL, ('Some ducts could not be checked (gray). Check their connections and '
                   'sizes, then rerun.')


# (pattern on the text without its leads, group, sentence or function of the
# match, estimate). First match wins, so specific patterns come first. Every
# engine template is covered by lib/_test_plain_corrections.py.
_PLAIN_RULES = [
    # ── Fix the design ──
    (r'^branch duct .+? leaves main .+? with only [\d.]+" of height, ' + _NUM +
     r'" min\. Raise the main to ' + _NUM + r'" tall, or make the branch ' + _NUM + '"',
     DESIGN, 'Ducts need a {0} in. minimum height difference where a branch leaves a '
             'main. Increase the main to {1} in. or reduce the branch to {2} in.', False),
    (r'^diffuser .+? neck is within [\d.]+" of the height of duct .+?, ' + _NUM +
     r'" min\. Raise the duct to ' + _NUM + r'" tall, or reduce the neck to ' + _NUM + '"',
     DESIGN, 'A diffuser neck needs to be at least {0} in. smaller than the duct feeding '
             'it. Increase the duct to {1} in. or reduce the neck to {2} in.', False),
    (r'^(?:main|branch) duct .+? steps down to duct .+? by only [\d.]+" of height, ' + _NUM +
     r'" min\. Make duct \S+ ' + _NUM + r'" tall or less, or keep it ' + _NUM + '"',
     DESIGN, 'Ducts need a {0} in. minimum height change where a duct steps down. Reduce '
             'the smaller duct to {1} in. or keep it at {2} in.', False),
    (r'^(?:Inlet swirl|Close-coupled elbows) at the (.+?):', DESIGN,
     'Two elbows at the {0} turn different ways too close to the fan. Add straight duct '
     'or turning vanes.', False),
    (r'^Double bend at the (.+?):', DESIGN,
     'Two elbows in a row are too close to the fan at the {0}. Add straight duct between '
     'the fan and the elbows.', False),
    (r'^Damper in the fan clearance at the (.+?): .+? inside .+? \(' + _NUM + r' ft\)',
     DESIGN, 'A damper is too close to the fan at the {0}. Move it at least {1} ft from the '
             'fan.', False),
    (r'^Elbow \S+ is [\d.]+ ft from the ([^,;]+)[^;]*?; recommended: at least ' + _NUM +
     ' ft of straight duct', DESIGN,
     'Elbow close to the {0}. At least {1} ft of straight duct is recommended between the '
     '{0} and the elbow; the fan loss may be higher than shown.', False),
    (r'^Elbow \S+ is [\d.]+ ft from the ([^,]+),.*?Recommended: at least ' + _NUM +
     ' ft of straight duct', DESIGN,
     'Elbow too close to the {0}. Provide at least {1} ft of straight duct between the {0} '
     'and the elbow.', False),
    (r'^System Effect Pressure Loss too high at \S+ \S+, [\d.]+ ft from the ([^,]+),.*?'
     r'recommended: increase length of duct to ' + _NUM + ' ft', DESIGN,
     'Fan connection loss is too high at the {0}. Increase the straight duct to {1} ft.',
     False),
    (r'^System Effect Pressure Loss too high at \S+ \S+, [\d.]+ ft from the ([^,]+),',
     DESIGN, 'Fan connection loss is too high at the {0}, and more straight duct will not '
             'fix it. Rework the duct connection at the fan.', False),
    (r'^Tap \S+ is [\d.]+ ft from the ([^,]+),.*?Recommended: at least ' + _NUM +
     ' ft of straight duct to the tap', DESIGN,
     'A tap is too close to the {0}. Provide at least {1} ft of straight duct between the '
     '{0} and the tap.', False),
    (r'^.+?, [\d.]+ ft from the (.+?), inside .+? \(' + _NUM + r' ft\)\. A fan system '
     r'effect disruption', DESIGN,
     'A fitting is too close to the fan at the {0}; its extra loss is not in the external '
     'static. Keep fittings at least {1} ft from the fan.', False),
    (r'flex duct .+? on the exhaust discharge', DESIGN,
     'Flex duct is used on the exhaust discharge. Exhaust ductwork shall be rigid.', False),
    (r'diffuser\(s\) on an MP main', DESIGN,
     'Diffusers are connected to the medium pressure main with no VAV box. Connect them '
     'downstream of a VAV box.', False),
    (r'^VAV box .+?, pressure left', DESIGN,
     'A VAV box gets less pressure than the minimum entered. Reduce the duct losses to the '
     'box or raise the unit\'s external static.', False),
    (r'^Unit max external static .+EXCEEDS IT', DESIGN,
     'The external static pressure is above the unit\'s maximum. Reduce the duct losses or '
     'select a unit with more static.', False),
    # ── Fix the model ──
    (r'^(\d+) terminal\(s\) have Flow = 0', MODEL,
     lambda m: '{} no airflow (Flow is 0). Enter the cfm and rerun.'.format(
         _count(m.group(1), 'diffuser or grille has', 'diffusers or grilles have')), False),
    (r'^(\d+) terminal\(s\) are missing the "Flow" parameter', MODEL,
     lambda m: '{} no Flow parameter. Use a family with Flow, enter the cfm and rerun.'.format(
         _count(m.group(1), 'diffuser or grille has', 'diffusers or grilles have')), False),
    (r'^(\d+) duct\(s\) have no dimension data', MODEL,
     lambda m: '{} no size. Give them a size and rerun.'.format(
         _count(m.group(1), 'duct has', 'ducts have')), False),
    (r'^No elements found in traversal', MODEL,
     'Nothing is connected to this unit. Connect the ductwork to the unit and rerun.', False),
    (r'^Ceiling exhaust fan discharge has', MODEL,
     'The ceiling exhaust fan discharge shall be one duct run to one wall or roof cap. Fix '
     'the discharge ductwork and rerun.', False),
    (r'^No air terminals', MODEL,
     'No diffusers or grilles are connected to this unit. Connect them to the ductwork and '
     'rerun.', False),
    (r'^No duct segments found', MODEL,
     'No ductwork is connected to this unit. Connect it and rerun.', False),
    (r'inlet connector direction could not be read', MODEL,
     'A piece of equipment\'s connector direction could not be read, so it was treated as '
     'a VAV box. Check that equipment\'s duct connectors.', False),
    (r'^No base equipment', MODEL,
     'The run did not start at a unit, so the airflow direction may be wrong. Pick the unit '
     'and rerun.', False),
    (r'^VAV System is ticked but no VAV box', MODEL,
     'VAV System was chosen but no VAV box is connected to this unit. Choose Rooftop Unit, '
     'or connect the VAV boxes, and rerun.', False),
    (r'^Parallel units on one main', MODEL,
     'Two units share this ductwork, so the external static is not reliable. Check it by '
     'hand.', False),
    (r'Take-off rows add to', MODEL,
     'The external static take-off does not add up. Send this run to Colin Nolan.', False),
    (r'Path search hit its step limit', MODEL,
     'The system is too large to search fully, so the external static may be low. Check the '
     'worst run by hand.', False),
    (r'but no index run was traced', MODEL,
     'Airflow was found but no run from the unit could be traced, so there is no external '
     'static for it. Check the ductwork connects from the unit to every diffuser and '
     'grille.', False),
    (r'^no mechanical floor plan on .+? could not be created', MODEL,
     'A level has no mechanical floor plan and one could not be made, so its ductwork is not '
     'on the sheet. Create a mechanical floor plan for that level and rerun.', False),
    (r'no connection reaches a grille or diffuser', MODEL,
     'The exhaust fan\'s intake is not connected to any grille. Connect its intake ductwork '
     'and rerun.', False),
    (r'size unreadable; fan system effect not checked|no size could be read', MODEL,
     'A fitting near the unit has no readable size, so the fan connection was not checked. '
     'Check the duct and fitting sizes at the unit.', False),
    (r'^duct \S+ \(.*?\) not judged: (.+?)\.?$', None,
     lambda m: _gray_headline(m.group(1)), False),
    (r'^[^:,]+: unrecognised fitting', MODEL,
     'A fitting on the worst run was not recognised, so its loss is not in the external '
     'static. Check the fitting family.', False),
    (r'transition, size change unreadable', MODEL,
     'A transition\'s sizes could not be read, so its loss is not in the external static. '
     'Check the transition.', False),
    (r'accessory on the (?:run|discharge), not priced', MODEL,
     'An accessory on the worst run was not recognised, so its loss is not in the external '
     'static. Check the accessory family.', False),
    (r'is a single diffuser main', MODEL,
     'A unit connection feeds only one diffuser. This is rare; if it is not intended, check '
     'the model for a disconnected run.', False),
    # ── For information ──
    (r'RJA worksheet C', INFO,
     'Some fitting losses use the RJA worksheet value because the ASHRAE tables do not '
     'cover them. No change needed.', True),
    (r'^read at the table edge', INFO,
     'Some fittings are outside the ASHRAE table range and use the nearest table value. '
     'No change needed.', True),
    (r'^read next to a blank cell', INFO,
     'Some fittings fall on a blank cell of an ASHRAE table and use the nearest printed '
     'value. No change needed.', True),
    (r'tap (?:\S+ )?positions? on duct .+? unreadable', INFO,
     'Some tap positions could not be read, so the full duct airflow was used there '
     '(conservative). No change needed.', True),
    (r'unit (?:inlet|.+? connection size) unreadable|inlet elbow geometry unreadable|'
     r'is mitered: ER7-1', INFO,
     'Some unit connection or elbow sizes could not be read, so the nearest duct size was '
     'used. No change needed.', True),
    (r'rank \d+ is above the', INFO,
     'The ranked run asked for is past the number of diffusers and grilles, so the last one '
     'is shown.', False),
    (r'(?:most restrictive run|run \d+) \(', INFO,
     'How the fan connection on the worst run was priced. No change needed.', False),
    (r'not checked for fan system effect|fan system effect at the .+? not checked', INFO,
     'A unit connection that feeds equipment directly, or has no duct or airflow past it, '
     'is not checked for fan loss.', False),
    (r'fan system effect not priced at a VAV box', INFO,
     'Fan loss is not checked at a VAV box.', False),
    (r'exhaust fan read as|inline exhaust fan with|^exhaust fan \S+, connection|'
     r'read as a ceiling exhaust fan', INFO,
     'How the exhaust fan type was read from its connections. No change needed.', False),
    (r'Outside Air branch\(es\) not traversed', INFO,
     'Outside air ductwork is not checked yet.', False),
    (r'upstream Supply Air branch\(es\) not traversed', INFO,
     'Ductwork upstream of the picked unit is not checked.', False),
    (r'^oval duct', INFO,
     'Oval ducts are checked as rectangular; oval sizing is not supported.', False),
    (r'one was created from the active view', INFO,
     'A mechanical floor plan was created for a level to show its ductwork on the sheet.',
     False),
    (r'discharge splits at element', INFO,
     'The exhaust discharge splits; only its first leg is priced.', False),
    (r'duct pick|fitting\(s\) between the picked duct|More than one system was selected',
     INFO, 'A note on the duct pick. No change needed.', False),
    (r'^(?:elbow|transition|tap|tee|cross|wye|split|fitting|offset|takeoff)\b.*?: ', INFO,
     'Some fitting losses use an ASHRAE table assumption. No change needed.', True),
]
_PLAIN_RULES = [(re.compile(p), g, s, est) for p, g, s, est in _PLAIN_RULES]


def plain(kind, text):
    """(group, sentence, estimate) for one correction line. group MODEL,
    DESIGN or INFO; estimate True for an estimated loss (one sheet line for
    all of them). A line no rule knows keeps a generic sentence in the group
    its kind and lead give it."""
    body = _strip_leads(text)
    for rx, group, sentence, est in _PLAIN_RULES:
        m = rx.search(body)
        if not m:
            continue
        if callable(sentence):
            out = sentence(m)
            if isinstance(out, tuple):
                group, out = out
            return group, out, est
        # the sheet names the place, never an element: 'exhaust inlet at
        # element 123' (two connections of one class) and the roof curb
        # clause read as 'exhaust inlet'
        args = [_WHERE_NOISE.sub('', a) if a else a for a in m.groups()]
        return group, sentence.format(*args), est
    if correction_group(kind, text) == ACTION:
        return ((DESIGN if kind == 'red' else MODEL),
                'One item needs a check; see the console for the details.', False)
    return INFO, 'See the console for the details.', False


def plain_groups(entries):
    """{MODEL, DESIGN, INFO: [(callout kind, sentence, [detail lines])]}.

    The engine's lines are grouped first (group_lines: one per tap table and
    reason, one per table edge and blank cell), then each becomes a plain
    sentence; lines with the same sentence (and unit, in a run of more than
    one unit) are one item, their engine lines kept as its details. The
    sentence carries the unit's mark, without its id, when the run has more
    than one unit."""
    raw = correction_lines(entries)
    units = []
    for _g, unit, _k, _t in entries:
        if unit and unit not in units:
            units.append(unit)
    multi = len(units) > 1
    out = {MODEL: [], DESIGN: [], INFO: []}
    index = {}
    for group in (ACTION, NOTES):
        for kind, line in raw[group]:
            unit = None
            detail = line
            if multi:
                for u in units:
                    if line.startswith(u + ': '):
                        unit, detail = u, line[len(u) + 2:]
                        break
            pgroup, sentence, est = plain(kind, detail)
            if unit:
                sentence = '{}: {}'.format(_UNIT_ID_RE.sub('', unit), sentence)
            key = (pgroup, sentence)
            if key not in index:
                ckind = ('red' if pgroup == DESIGN else 'gray' if pgroup == MODEL
                         else 'info')
                index[key] = (ckind, sentence, [], est)
                out[pgroup].append(index[key])
            index[key][2].append(_strip_leads(detail) or detail)
    return out


def system_corrections(headers, rows, entries, status_col=None, align=None,
                       widths=None, subtitle=None):
    """Section e: System Corrections. The duct table (when any duct is
    listed), then each plain sentence (Fix the model, Fix the design as
    callouts; For information as lines), the engine's own lines under it as
    the details. Nothing at all: one NO_CORRECTIONS line."""
    groups = plain_groups(entries)
    if subtitle is None:
        subtitle = '{} duct(s)'.format(len(rows))
    out = [cr.section('System Corrections', subtitle)]
    anything = any(groups[g] for g in groups)
    if rows:
        out.append(cr.table(headers, rows, align=align, status_col=status_col,
                            widths=widths))
    elif anything:
        out.append(cr.note(NO_FLAGGED))
    else:
        out.append(cr.callout(NO_CORRECTIONS, 'ok'))
        return ''.join(out)
    for group, title in PLAIN_TITLES:
        if not groups[group]:
            continue
        out.append(cr.subsection(title))
        for ckind, sentence, details, _est in groups[group]:
            out.append(cr.callout(sentence, ckind))
            for d in details:
                out.append(cr.note('Details: ' + d))
    return ''.join(out)


def corrections_text(entries, any_rows=False):
    """The plain sentences and their details as markdown lines, for the
    print_md fallback under the duct table text."""
    groups = plain_groups(entries)
    if not any(groups[g] for g in groups):
        return [] if any_rows else [NO_CORRECTIONS]
    out = [] if any_rows else [NO_FLAGGED]
    for group, title in PLAIN_TITLES:
        if groups[group]:
            out.append('**{}**'.format(title))
            for _c, sentence, details, _e in groups[group]:
                out.append('- ' + sentence)
                out.extend('    - Details: ' + d for d in details)
    return out


def corrections_sheet(entries):
    """[(heading, [line])] for the drafting view's System Corrections block:
    FIX THE MODEL and FIX THE DESIGN with each plain sentence once, no ids
    or table codes; FOR INFORMATION as the one ESTIMATE_SHEET_LINE when any
    loss was estimated. Empty groups left out."""
    groups = plain_groups(entries)
    out = []
    for group, title in PLAIN_TITLES:
        if group == INFO:
            if any(item[3] for item in groups[INFO]):
                out.append((title.upper(), [ESTIMATE_SHEET_LINE]))
            continue
        if groups[group]:
            out.append((title.upper(), [item[1] for item in groups[group]]))
    return out


def summary_band_text(rec):
    """The System Summary grand totals as one line for the right end of the
    sheet's SYSTEM SUMMARY band, e.g. 'Supply CFM 5,250 | Return CFM 1,200 |
    Total Diffusers/Grilles 18 | Ducts 60'."""
    return ' | '.join('{} {}'.format(label, value)
                      for label, value in rec.get('grand') or [])


# The basis named in the External Static Pressure header (Colin, 2026-11-09):
# the 2001 tables by geometry since 2026-10-23, the worksheet only as the
# fallback.
ESP_BASIS = ('index run, fittings per 2001 ASHRAE Fundamentals Ch. 34 by geometry, '
             'RJA SP_LOSS_WORKSHEET as fallback')


def external_static(rec):
    """Section b: External Static Pressure, one combined table per equipment,
    then any picked duct pairing."""
    out = [cr.section('External Static Pressure', ESP_BASIS)]
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
            # Blank inputs skip the check (Colin, 2026-11-12); say which.
            if rep.get('unit_max') is None:
                out.append(('note', 'Not checked: no unit max external static entered.'))
            if rep.get('minimum', 0) is None:
                out.append(('note', 'Not judged: no minimum pressure at a box entered.'))
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
    words = rec.get('words') or scope_words('rtu')
    out.append(cr.table(['Color', 'Ducts', legend_fittings_header(words), 'Meaning'],
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
