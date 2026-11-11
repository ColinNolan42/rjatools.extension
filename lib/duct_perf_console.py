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
                '(diffusers and grilles, MEP Design Standards, section 1), not '
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
         'of slots (Titus ML-38 supply, MLR-38 return) is red, naming the slot '
         'that fits or "Max CFM reached, split slot".'),
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
RANKED_TOTAL_LABEL = 'Ranked runs total'
RANKED_TOTAL_NOTE = 'sanity check, not the design ESP'


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
                     'or system.'.format(one))
    if 'R1' in marks:
        lines.append('(R1): Most restrictive {}. It ends the index run that sets '
                     'the required external static pressure for its equipment '
                     'or system.'.format(one))
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
    if cls.get('rank_text'):
        # Ranked run (Colin, 2026-11-10): which run this take-off is.
        head = '{}, {}'.format(head, cls['rank_text'])
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
    total = ['{} total{}'.format(cls['sys_class'],
                                 ', run #{}'.format(cls['rank']) if cls.get('rank') else ''),
             '' if sheet else (cls.get('total_detail') or ''),
             '', '', _f3(cls['total'])]
    rows.append({'cells': total, 'bold': True} if sheet else total)
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
    if unit.get('esp_total') is not None:
        mr_label = (VAVBOX_ESP_LABEL + MR_TAG if unit.get('caution') else MR_LABEL)
        if unit.get('ranked'):
            ranked = [RANKED_TOTAL_LABEL, '{}, {}'.format(
                unit.get('ranked_text') or ' + '.join(sides or []), RANKED_TOTAL_NOTE),
                '', '', _f3(unit['esp_total'])]
            rows.append({'cells': ranked, 'bold': True} if sheet else ranked)
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
