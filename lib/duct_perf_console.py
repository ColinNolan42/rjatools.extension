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
                  diffusers, other]], 'by_class': [[system, cfm, diffusers]],
                  'types': [[family, count, cfm]], 'warnings': [text]}

  external static {'units': [unit], 'combined': [{'title', 'text'}],
                  'warnings': [text], 'method': text, 'not_included': text}
                  unit = {'label', 'classes': [cls], 'untraced': [text],
                  'esp': text or None, 'one_sided': text or None,
                  'checks': [(ok bool, text)], 'warnings': [text]}
                  cls = {'sys_class', 'airflow', 'rows': [(item, inwc float,
                  detail)], 'subtotal' float, 'safety' float, 'safety_detail',
                  'total' float, 'total_detail', 'notes': [text],
                  'warnings': [text], 'takeoff': [[item, table, qty, C,
                  in. wc]] (optional, from takeoff_rows; rendered instead of
                  'rows' when present), 'filter': float or None}

  VAV boxes       the dict returned by script.py _vav_box_report()

  index run       [{'title', 'items': [(item, fpm, inwc float, detail)],
                  'subtotal' float, 'subtotal_detail'}]

IronPython 2.7: no f-strings; floats are cast with float() before '{:.3f}'.
"""

import console_report as cr


# Relative column widths per table kind. Every table of one kind is built with
# the same widths (fixed layout), so stacked tables line up column for column.
W_EQUIPMENT = [40, 18, 12, 30]
W_BY_CLASS = [50, 25, 25]
W_TYPES = [60, 15, 25]
W_TAKEOFF = [24, 42, 12, 10, 12]
W_ESP = [24, 12, 64]
W_VAV = [22, 14, 16, 16, 18, 14]
W_RUN = [22, 8, 10, 60]
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
    if rec.get('by_class'):
        out.append(cr.table(['System', 'CFM', 'Diffusers'], rec['by_class'],
                            align=['l', 'r', 'r'], caption='By system class',
                            widths=W_BY_CLASS, width=NARROW))
    if rec.get('types'):
        out.append(cr.table(['Family', 'Count', 'CFM'], rec['types'],
                            align=['l', 'r', 'r'], caption='Diffuser types',
                            widths=W_TYPES, width=NARROW))
    for w in rec.get('warnings') or []:
        out.append(cr.callout(w, 'warn'))
    return ''.join(out)


TAKEOFF_HEADERS = ['Item', 'Table / basis', 'Qty', 'C', 'in. wc']

# Row order in the take-off: duct first, fittings, taps passed, components.
_TAKEOFF_RANK = (
    ('Straight duct', 0), ('Flex duct', 1), ('Flex bend', 2),
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


def _esp_takeoff_table(cls):
    """The class's take-off table with Subtotal, Safety factor and TOTAL."""
    rows = [list(r) for r in cls['takeoff']]
    if cls.get('filter') is not None:
        rows.append(['Return filter', 'known drop', '1', '-', _f3(cls['filter'])])
    rows.append(['Subtotal', '', '', '', _f3(cls['subtotal'])])
    rows.append(['Safety factor', cls.get('safety_detail') or '', '', '',
                 _f3(cls['safety'])])
    caption = cls['sys_class']
    if cls.get('airflow'):
        caption = '{}, {}'.format(caption, cls['airflow'])
    return cr.table(TAKEOFF_HEADERS, rows, align=['l', 'l', 'r', 'r', 'r'],
                    total_row=['TOTAL', cls.get('total_detail') or '', '', '',
                               _f3(cls['total'])],
                    caption=caption, widths=W_TAKEOFF)


def _esp_class_table(cls):
    if cls.get('takeoff'):
        return _esp_takeoff_table(cls)
    rows = [[item, _f3(val), det or ''] for item, val, det in cls['rows']]
    rows.append(['Subtotal', _f3(cls['subtotal']), ''])
    rows.append(['Safety factor', _f3(cls['safety']), cls.get('safety_detail') or ''])
    caption = cls['sys_class']
    if cls.get('airflow'):
        caption = '{}, {}'.format(caption, cls['airflow'])
    return cr.table(['Item', 'in. wc', 'Detail'], rows, align=['l', 'r', 'l'],
                    total_row=['TOTAL', _f3(cls['total']), cls.get('total_detail') or ''],
                    caption=caption, widths=W_ESP)


def external_static(rec):
    """Section b: External Static Pressure, per equipment then per pairing."""
    out = [cr.section('External Static Pressure',
                      'index run, per RJA SP_LOSS_WORKSHEET')]
    for unit in rec.get('units') or []:
        out.append(cr.subsection(unit['label']))
        for cls in unit.get('classes') or []:
            out.append(_esp_class_table(cls))
            for n in cls.get('notes') or []:
                out.append(cr.note(n))
            for w in cls.get('warnings') or []:
                out.append(cr.callout(w, 'warn'))
        for w in unit.get('untraced') or []:
            out.append(cr.callout(w, 'warn'))
        if unit.get('esp'):
            out.append(cr.callout(unit['esp'], 'info'))
        if unit.get('one_sided'):
            out.append(cr.callout(unit['one_sided'], 'warn'))
        for ok, text in unit.get('checks') or []:
            out.append(cr.callout(text, 'ok' if ok else 'error'))
        for w in unit.get('warnings') or []:
            out.append(cr.callout(w, 'warn'))
    for comb in rec.get('combined') or []:
        out.append(cr.subsection(comb['title']))
        out.append(cr.callout(comb['text'], 'info'))
    for w in rec.get('warnings') or []:
        out.append(cr.callout(w, 'warn'))
    if rec.get('method'):
        out.append(cr.note(rec['method']))
    if rec.get('not_included'):
        out.append(cr.note(rec['not_included']))
    return ''.join(out)


def vav_boxes(rep):
    """Section c: VAV Boxes, one table over every unit."""
    out = [cr.section('VAV Boxes', rep.get('basis'))]
    rows = []
    for unit in rep.get('units') or []:
        for b in unit.get('boxes') or []:
            status = b.get('status')
            if status == 'NOT CHECKED':
                left = 'not checked (no unit max entered)'
            else:
                left = _f3(b.get('left'))
            if status == 'NOT REACHED':
                loss = 'not reached by the index run search'
            else:
                loss = _f3(b.get('loss'))
            rows.append([unit['label'], str(b['box']), loss, _f3(b.get('loss_sf')),
                         left, status or '-'])
    out.append(cr.table(
        ['Unit', 'Box', 'MP loss to inlet', 'with safety factor', 'Pressure left', 'Status'],
        rows, align=['l', 'l', 'r', 'r', 'r', 'c'], status_col=5, widths=W_VAV))
    if rep.get('minimum') is not None:
        out.append(cr.note('LOW: pressure left under the {:.2f} in. wc minimum. '
                           'Values in in. wc.'.format(float(rep['minimum']))))
    for unit in rep.get('units') or []:
        if unit.get('mp_terminal_text'):
            out.append(cr.callout('{}: {}'.format(unit['label'], unit['mp_terminal_text']),
                                  'warn'))
    return ''.join(out)


def index_run_detail(runs):
    """Section d: Index Run Detail, item by item per run."""
    out = [cr.section('Index Run Detail')]
    for run in runs:
        out.append(cr.subsection(run['title']))
        rows = [[lbl, fpm, _f3(loss), det] for lbl, fpm, loss, det in run['items']]
        out.append(cr.table(['Item', 'FPM', 'in. wc', 'Detail'], rows,
                            align=['l', 'r', 'r', 'l'],
                            total_row=['Subtotal', '', _f3(run['subtotal']),
                                       run.get('subtotal_detail') or ''],
                            widths=W_RUN))
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
                        swatch_col=0, caption='Color legend', widths=W_LEGEND))
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
