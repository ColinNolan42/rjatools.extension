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
                  diffusers, other]], 'warnings': [text]} (by_class and types
                  may be present; the console no longer shows them)

  external static {'units': [unit], 'combined': [{'title', 'text'}],
                  'warnings': [text], 'method': text, 'assumptions': [text]
                  (optional), 'not_included': text}
                  unit = {'label', 'classes': [cls], 'untraced': [text],
                  'esp': text or None, 'one_sided': text or None,
                  'checks': [(ok bool, text)], 'warnings': [text],
                  'esp_total': float or None, 'esp_sides': [class names],
                  'fse': [(kind, text)] fan system effect notes, kind 'note',
                  'warn' or 'gray'}
                  cls = {'sys_class', 'airflow', 'rows': [(item, inwc float,
                  detail)], 'subtotal' float, 'safety' float, 'safety_detail',
                  'total' float, 'total_detail', 'notes': [text],
                  'warnings': [text], 'takeoff': [[item, table, qty, C,
                  in. wc]] (optional, from takeoff_rows; rendered instead of
                  'rows' when present), 'filter': float or None}

  VAV boxes       the dict returned by script.py _vav_box_report()

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
    """Section a: System Summary."""
    out = [cr.section('System Summary')]
    out.append(cr.keyvalue(rec.get('grand') or []))
    if rec.get('equipment'):
        out.append(cr.table(
            ['Equipment', 'Supply discharge CFM', 'Diffusers', 'Other air feeding it'],
            rec['equipment'], align=['l', 'r', 'r', 'l'], widths=W_EQUIPMENT))
    for w in rec.get('warnings') or []:
        out.append(cr.callout(w, 'warn'))
    return ''.join(out)


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
        total_row = ['TOTAL EXTERNAL STATIC PRESSURE', ' + '.join(sides or []),
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
    """Everything said under a unit's table, in console order, as
    [(kind, text)] with kind 'note', 'warn', 'ok' or 'error'. sheet True
    first adds where each class's index run ends ('note'), which the sheet
    table leaves out of its total row."""
    out = []
    for cls in unit.get('classes') or []:
        if sheet and cls.get('total_detail'):
            out.append(('note', '{}: {}.'.format(cls['sys_class'], cls['total_detail'])))
        for n in cls.get('notes') or []:
            out.append(('note', '{}: {}'.format(cls['sys_class'], n)))
        for w in cls.get('warnings') or []:
            out.append(('warn', w))
    # The unit's fan system effect, once per unit (not per class): every
    # connection's notes, kind 'note', 'warn' or 'gray' (Colin, 2026-10-29).
    for kind, text in unit.get('fse') or []:
        out.append((kind, text))
    for w in unit.get('untraced') or []:
        out.append(('warn', w))
    if unit.get('one_sided'):
        out.append(('warn', unit['one_sided']))
    for ok, text in unit.get('checks') or []:
        out.append(('ok' if ok else 'error', text))
    for w in unit.get('warnings') or []:
        out.append(('warn', w))
    return out


# Text that already names what it is keeps its own lead on the sheet.
_OWN_LEADS = ('WARNING', 'UNPRICED', 'FALLBACK', 'NOTE', 'EXTERNAL STATIC', 'LOW',
              'GRAY')


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
    for w in rec.get('warnings') or []:
        out.append(cr.callout(w, 'warn'))
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
    """The notes under the VAV table as [(kind, text)]. sheet True adds what
    the short 'not reached' and 'not checked' cells mean, when present."""
    out = []
    if rep.get('minimum') is not None:
        out.append(('note', 'LOW: pressure left under the {:.2f} in. wc minimum. '
                            'Values in in. wc.'.format(float(rep['minimum']))))
    if sheet:
        statuses = set(b.get('status') for u in rep.get('units') or []
                       for b in u.get('boxes') or [])
        if 'NOT REACHED' in statuses:
            out.append(('note', 'Not reached: the index run search did not reach the box.'))
        if 'NOT CHECKED' in statuses:
            out.append(('note', 'Not checked: no unit max external static entered.'))
    for unit in rep.get('units') or []:
        if unit.get('mp_terminal_text'):
            out.append(('warn', '{}: {}'.format(unit['label'], unit['mp_terminal_text'])))
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


def duct_table(title, headers, rows, status_col=None, align=None, widths=None):
    """Section e: Flagged Ducts / All Ducts. widths: the _COLUMN_DEFS console
    widths, so the columns keep the same proportions as the sheet table."""
    return cr.section(title, '{} duct(s)'.format(len(rows))) + cr.table(
        headers, rows, align=align, status_col=status_col, widths=widths)


def limits_and_legend(rec):
    """Section f: Design Limits and Legend.

    rec = {'roles_note', 'limits_caption', 'limits': [[system, fpm, fric]],
    'vav_basis': [(label, value)], 'notes': [(bold lead, text)],
    'legend': [[key, ducts, fittings, meaning]], 'clearance': text}
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
