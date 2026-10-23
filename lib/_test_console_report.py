# -*- coding: ascii -*-
"""Tests for console_report (themed HTML builders for the pyRevit output
window) and duct_perf_console (the Duct Performance sections built from them).

Both modules are pure Python, so they are tested directly under CPython.
The markup is checked by parsing it with html.parser: tag balance, row and
cell counts, and the text that reaches the reader after unescaping.

Run:  python _test_console_report.py
"""

import io
import os
import re
import sys
import unittest

try:
    from html.parser import HTMLParser
except ImportError:   # pragma: no cover, CPython 2
    from HTMLParser import HTMLParser

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import console_report as cr        # noqa: E402
import duct_perf_console as dpc    # noqa: E402


class _Shape(HTMLParser):
    """Counts tags, checks nesting, and collects cell text and styles."""

    VOID = ('br', 'img', 'hr')

    def __init__(self):
        HTMLParser.__init__(self, convert_charrefs=True)
        self.stack = []
        self.errors = []
        self.counts = {}
        self.rows = []        # list of rows, each a list of (tag, style, text)
        self._cell = None
        self.text = []

    def handle_starttag(self, tag, attrs):
        self.counts[tag] = self.counts.get(tag, 0) + 1
        if tag in self.VOID:
            return
        self.stack.append(tag)
        if tag == 'tr':
            self.rows.append([])
        if tag in ('td', 'th'):
            self._cell = [tag, dict(attrs).get('style', ''), '']

    def handle_endtag(self, tag):
        if tag in self.VOID:
            return
        if not self.stack or self.stack[-1] != tag:
            self.errors.append('unbalanced </{}>'.format(tag))
            return
        self.stack.pop()
        if tag in ('td', 'th') and self._cell is not None:
            self.rows[-1].append(tuple(self._cell))
            self._cell = None

    def handle_data(self, data):
        self.text.append(data)
        if self._cell is not None:
            self._cell[2] += data


def shape(html):
    p = _Shape()
    p.feed(html)
    p.close()
    if p.stack:
        p.errors.append('unclosed {}'.format(p.stack))
    return p


class EscapeTests(unittest.TestCase):

    def test_special_characters_escaped(self):
        self.assertEqual(cr.esc('<b>&"\''), '&lt;b&gt;&amp;&quot;&#39;')

    def test_none_and_numbers(self):
        self.assertEqual(cr.esc(None), '')
        self.assertEqual(cr.esc(12), '12')
        self.assertEqual(cr.esc(0.5), '0.5')

    def test_ampersand_escaped_once(self):
        self.assertEqual(cr.esc('&amp;'), '&amp;amp;')

    def test_text_reaches_reader_unchanged(self):
        raw = 'Slot <Linear> & "Co" 24x14'
        for html in (cr.section(raw, raw), cr.subsection(raw), cr.callout(raw, 'warn'),
                     cr.note(raw), cr.paragraph(raw, raw),
                     cr.keyvalue([(raw, raw)]), cr.table([raw], [[raw]])):
            p = shape(html)
            self.assertEqual(p.errors, [])
            self.assertIn(raw, ''.join(p.text))
            self.assertNotIn('<Linear>', html)

    def test_unicode_text(self):
        le = chr(0x2264)
        p = shape(cr.note('green ' + le + ' max'))
        self.assertIn(le, ''.join(p.text))


class TableTests(unittest.TestCase):

    def test_shape(self):
        p = shape(cr.table(['A', 'B', 'C'], [['1', '2', '3'], ['4', '5', '6']]))
        self.assertEqual(p.errors, [])
        self.assertEqual(p.counts['table'], 1)
        self.assertEqual(len(p.rows), 3)
        self.assertEqual([c[0] for c in p.rows[0]], ['th', 'th', 'th'])
        self.assertTrue(all(len(r) == 3 for r in p.rows))
        self.assertEqual([c[2] for c in p.rows[2]], ['4', '5', '6'])

    def test_short_row_padded_long_row_trimmed(self):
        p = shape(cr.table(['A', 'B'], [['1'], ['1', '2', '3']]))
        self.assertEqual([len(r) for r in p.rows], [2, 2, 2])

    def test_zebra_rows(self):
        p = shape(cr.table(['A'], [['1'], ['2'], ['3']]))
        self.assertIn('#FFFFFF', p.rows[1][0][1])
        self.assertIn(cr.ZEBRA_BG, p.rows[2][0][1])
        self.assertIn('#FFFFFF', p.rows[3][0][1])

    def test_alignment(self):
        p = shape(cr.table(['A', 'B', 'C'], [['x', '1', 'y']], align=['l', 'r', 'c']))
        self.assertIn('text-align:left', p.rows[1][0][1])
        self.assertIn('text-align:right', p.rows[1][1][1])
        self.assertIn('text-align:center', p.rows[1][2][1])

    def test_total_row(self):
        p = shape(cr.table(['Item', 'in. wc'], [['Duct', '0.100']],
                           total_row=['TOTAL', '0.110']))
        self.assertEqual(len(p.rows), 3)
        last = p.rows[-1]
        self.assertEqual([c[2] for c in last], ['TOTAL', '0.110'])
        for cell in last:
            self.assertIn('font-weight:bold', cell[1])
            self.assertIn(cr.TOTAL_BG, cell[1])

    def test_empty_rows(self):
        p = shape(cr.table(['A', 'B'], []))
        self.assertEqual(p.errors, [])
        self.assertEqual(len(p.rows), 1)

    def test_caption(self):
        html = cr.table(['A'], [['1']], caption='Supply Air, 950 CFM')
        self.assertIn('Supply Air, 950 CFM', ''.join(shape(html).text))


class StatusTests(unittest.TestCase):

    def test_status_keys(self):
        for word in ('GREEN', 'YELLOW', 'RED', 'PURPLE', 'GRAY'):
            self.assertEqual(cr.status_key(word), word)
            self.assertEqual(cr.status_key(word.lower()), word)
        self.assertEqual(cr.status_key('LOW: under the minimum'), 'RED')
        self.assertEqual(cr.status_key('OK'), 'GREEN')
        self.assertEqual(cr.status_key('NOT CHECKED'), 'GRAY')
        self.assertIsNone(cr.status_key(''))
        self.assertIsNone(cr.status_key(None))
        self.assertIsNone(cr.status_key('Main'))

    def test_plan_colors(self):
        # Duct Performance plan overrides: Color(0,200,0) (255,215,0)
        # (210,40,40) (140,40,200) (160,160,160).
        def hexof(r, g, b):
            return '#{:02X}{:02X}{:02X}'.format(r, g, b)
        self.assertEqual(cr.STATUS_COLORS['GREEN'], hexof(0, 200, 0))
        self.assertEqual(cr.STATUS_COLORS['YELLOW'], hexof(255, 215, 0))
        self.assertEqual(cr.STATUS_COLORS['RED'], hexof(210, 40, 40))
        self.assertEqual(cr.STATUS_COLORS['PURPLE'], hexof(140, 40, 200))
        self.assertEqual(cr.STATUS_COLORS['GRAY'], hexof(160, 160, 160))

    def test_status_cell_colored(self):
        rows = [['1', s] for s in ('RED', 'YELLOW', 'PURPLE', 'GREEN', 'GRAY', 'Main')]
        p = shape(cr.table(['#', 'Status'], rows, status_col=1))
        for r, key in zip(p.rows[1:], ('RED', 'YELLOW', 'PURPLE', 'GREEN', 'GRAY')):
            self.assertIn('background:' + cr.STATUS_COLORS[key], r[1][1])
            self.assertIn('color:' + cr.STATUS_TEXT[key], r[1][1])
            self.assertNotIn(cr.STATUS_COLORS[key], r[0][1])
        for key in cr.STATUS_COLORS:
            self.assertNotIn(cr.STATUS_COLORS[key], p.rows[-1][1][1])

    def test_swatch_column(self):
        html = cr.table(['Color', 'Meaning'], [['Purple', 'oversized']], swatch_col=0)
        self.assertIn('background:' + cr.STATUS_COLORS['PURPLE'], html)
        self.assertIn('Purple', ''.join(shape(html).text))


class BlockTests(unittest.TestCase):

    def test_section_is_blue_bar(self):
        html = cr.section('System Summary', 'sub')
        self.assertIn(cr.RJA_BLUE, html)
        self.assertIn('color:#FFFFFF', html)
        self.assertIn('Segoe UI', html)

    def test_subsection_rule(self):
        self.assertIn('border-left:4px solid ' + cr.RJA_BLUE, cr.subsection('RTU-1'))

    def test_callout_kinds(self):
        seen = set()
        for kind in ('warn', 'error', 'info', 'ok'):
            html = cr.callout('x', kind)
            self.assertIn(cr.CALLOUT_STYLES[kind][1], html)
            seen.add(cr.CALLOUT_STYLES[kind][1])
        self.assertEqual(len(seen), 4)
        self.assertIn(cr.CALLOUT_STYLES['info'][1], cr.callout('x', 'nonsense'))

    def test_keyvalue_shape(self):
        p = shape(cr.keyvalue([('CFM', '5,250'), ('Ducts', '37')]))
        self.assertEqual(p.errors, [])
        self.assertEqual([[c[2] for c in r] for r in p.rows],
                         [['CFM', '5,250'], ['Ducts', '37']])

    def test_markdown_table(self):
        md = cr.markdown_table(['A', 'B'], [['1', 'x|y']])
        self.assertEqual(md.split('\n'), ['| A | B |', '| --- | --- |', '| 1 | x/y |'])

    def test_no_layout_css_or_external_resources(self):
        html = (cr.section('a') + cr.subsection('b') + cr.table(['a'], [['1']]) +
                cr.callout('c') + cr.keyvalue([('a', 'b')]) + cr.note('d'))
        for bad in ('display:flex', 'display:grid', 'http', '<link', '<script', 'class='):
            self.assertNotIn(bad, html)


def _sample_esp():
    cls = {'sys_class': 'Supply Air', 'airflow': '950 CFM, 3 terminals',
           'rows': [('Duct friction', 0.212, '12 ducts, 144 ft'),
                    ('Fitting losses', 0.301, '9 fittings, 4 taps restricting the main'),
                    ('Components', 0.3, 'Diffuser / grille x1 @ 0.050')],
           'subtotal': 0.813, 'safety': 0.0813, 'safety_detail': '+10%',
           'total': 0.8943, 'total_detail': 'worst run ends at CD-1 [12]',
           'notes': ['(Includes 2 fitting(s) between the picked duct and the unit connection.)'],
           'warnings': ['UNPRICED: Fitting 777']}
    return {'units': [{'label': 'RTU-1 (id 1)', 'classes': [cls], 'untraced': [],
                       'esp': 'EXTERNAL STATIC PRESSURE = 1.257 in. wc (Supply Air + Return Air)',
                       'esp_total': 0.8943, 'esp_sides': ['Supply Air'],
                       'one_sided': None, 'checks': [(False, 'Unit max ... EXCEEDS IT.')],
                       'warnings': ['Parallel units on one main.']}],
            'combined': [{'title': 'Duct 3 + Duct 4', 'text': 'EXTERNAL STATIC PRESSURE = 0.700'}],
            'warnings': [], 'method': 'METHOD: Darcy-Weisbach', 'not_included': 'NOT INCLUDED: coil'}


class DuctPerfSectionTests(unittest.TestCase):

    def test_esp_unit_table_rows_and_total(self):
        p = shape(dpc.external_static(_sample_esp()))
        self.assertEqual(p.errors, [])
        body = [[c[2] for c in r] for r in p.rows]
        self.assertEqual(body[0], dpc.TAKEOFF_HEADERS)
        self.assertEqual([r[0] for r in body[1:]],
                         ['Supply Air, 950 CFM, 3 terminals', 'Duct friction',
                          'Fitting losses', 'Components', 'Subtotal', 'Safety factor',
                          'Supply Air total', 'TOTAL EXTERNAL STATIC PRESSURE'])
        self.assertEqual(body[-1], ['TOTAL EXTERNAL STATIC PRESSURE', 'Supply Air',
                                    '', '', '0.894'])
        html = dpc.external_static(_sample_esp())
        self.assertIn('colspan="5"', html)
        self.assertIn(cr.CALLOUT_STYLES['error'][1], html)
        self.assertIn(cr.CALLOUT_STYLES['warn'][1], html)
        self.assertIn('Duct 3 + Duct 4', html)

    def test_supply_and_return_in_one_table(self):
        """Colin, 2026-10-23: one chart per unit, supply then return, ending in
        the unit's total external static pressure; the lesser return side is
        shown but marked not counted."""
        def cls(name, total):
            return {'sys_class': name, 'rows': [('Duct friction', total / 1.1, '')],
                    'subtotal': total / 1.1, 'safety': total - total / 1.1,
                    'total': total, 'total_detail': ''}
        unit = {'label': 'RTU-1', 'classes': [cls('Supply Air', 0.5), cls('Return Air', 0.3),
                                              cls('Exhaust Air', 0.1)],
                'esp_total': 0.8, 'esp_sides': ['Supply Air', 'Return Air']}
        html = dpc.external_static({'units': [unit]})
        self.assertEqual(html.count('<table'), 1)
        self.assertIn('Exhaust Air  (not counted, the larger return side governs)', html)
        self.assertNotIn('Return Air  (not counted', html)
        body = [[c[2] for c in r] for r in shape(html).rows]
        self.assertEqual(body[-1], ['TOTAL EXTERNAL STATIC PRESSURE',
                                    'Supply Air + Return Air', '', '', '0.800'])

    def test_one_sided_unit_has_no_grand_total(self):
        unit = {'label': 'Duct 9', 'classes': [{'sys_class': 'Supply Air', 'rows': [],
                'subtotal': 0.2, 'safety': 0.02, 'total': 0.22}],
                'one_sided': 'EXTERNAL STATIC PRESSURE = 0.220 in. wc (Supply Air only)'}
        html = dpc.external_static({'units': [unit]})
        self.assertNotIn('TOTAL EXTERNAL STATIC PRESSURE</td>', html)
        self.assertIn('Supply Air only', html)

    def test_vav_status_and_low(self):
        rep = {'basis': 'MP supply 1500 FPM', 'minimum': 0.5, 'units': [
            {'label': 'RTU-1', 'mp_terminal_text': '1 diffuser(s) on an MP main',
             'boxes': [
                 {'box': 101, 'loss': 0.41, 'loss_sf': 0.451, 'left': 1.049, 'status': 'OK'},
                 {'box': 102, 'loss': 1.3, 'loss_sf': 1.43, 'left': 0.07, 'status': 'LOW'},
                 {'box': 103, 'loss': None, 'loss_sf': None, 'left': None,
                  'status': 'NOT REACHED'},
                 {'box': 104, 'loss': 0.2, 'loss_sf': 0.22, 'left': None,
                  'status': 'NOT CHECKED'}]}]}
        p = shape(dpc.vav_boxes(rep))
        self.assertEqual(p.errors, [])
        self.assertEqual(len(p.rows), 5)
        self.assertEqual([c[2] for c in p.rows[2]],
                         ['RTU-1', '102', '1.300', '1.430', '0.070', 'LOW'])
        self.assertIn(cr.STATUS_COLORS['RED'], p.rows[2][5][1])
        self.assertIn(cr.STATUS_COLORS['GREEN'], p.rows[1][5][1])
        self.assertEqual(p.rows[3][2][2], 'not reached by the index run search')
        self.assertEqual(p.rows[4][4][2], 'not checked (no unit max entered)')

    def test_vav_without_external_static(self):
        rep = {'basis': 'b', 'minimum': 0.5, 'units': [
            {'label': 'RTU-1', 'mp_terminal_text': None,
             'boxes': [{'box': 7, 'loss': None, 'loss_sf': None, 'left': None,
                        'status': None}]}]}
        p = shape(dpc.vav_boxes(rep))
        self.assertEqual([c[2] for c in p.rows[1]], ['RTU-1', '7', '-', '-', '-', '-'])

    def test_summary_is_compact(self):
        """Colin, 2026-10-23: only totals and the equipment table; the by system
        and diffuser type breakdowns stay on the sheet."""
        rec = {'grand': [('Grand total airflow', '5,250 CFM')],
               'equipment': [['RTU-1 (id 1)', '950', '3', '-']],
               'by_class': [['Supply Air', '2,700', '6']],
               'types': [['CD', '4', '1,500']]}
        html = dpc.system_summary(rec)
        self.assertNotIn('By system class', html)
        self.assertNotIn('Diffuser types', html)
        self.assertIn('RTU-1 (id 1)', html)

    def test_summary_and_legend_and_sheets(self):
        rec = {'grand': [('Grand total airflow', '5,250 CFM')],
               'equipment': [['RTU-1 (id 1)', '950', '3', '-']],
               'warnings': ['1 diffuser(s) with Flow = 0']}
        for html in (dpc.system_summary(rec),
                     dpc.limits_and_legend({'limits': [['Supply Air', '800 FPM', '0.080']],
                                            'legend': [['Red', '3', '1', 'exceeds']],
                                            'clearance': (True, 'clearance: 2')}),
                     dpc.sheets_created(['M-601']), dpc.sheets_created([]),
                     dpc.duct_table('Flagged Ducts', ['#', 'Status'], [['1', 'RED']],
                                    status_col=1)):
            self.assertEqual(shape(html).errors, [])


class TakeoffTests(unittest.TestCase):
    """External static take-off rows (Colin, 2026-10-23)."""

    def _items(self):
        return [
            ('Duct 1', '24x12, 10 ft', 800.0, 0.010, ('Straight duct', 'Darcy-Weisbach', None, 10.0)),
            ('Fitting 2', 'Round Elbow, CD3-1', 800.0, 0.004, ('Round elbow, 90 deg', 'CD3-1', 0.11, None)),
            ('Duct 3', '12 dia, 5 ft', 900.0, 0.006, ('Straight duct', 'Darcy-Weisbach', None, 5.0)),
            ('Fitting 4', 'Round Elbow, CD3-1', 900.0, 0.004, ('Round elbow, 90 deg', 'CD3-1', 0.11, None)),
            ('Terminal 5', 'SD', None, 0.05, ('Diffuser / grille', 'known drop', None, None)),
            ('Tap passed', 'on duct 1', 800.0, 0.002, ('Tap passed on main (straight through)', 'SR5-13 Cs', 0.04, None)),
        ]

    def test_groups_and_sums(self):
        rows, losses = dpc.takeoff_rows(self._items())
        names = [r[0] for r in rows]
        self.assertEqual(names[0], 'Straight duct')
        self.assertEqual(rows[0][2], '2 (15 ft)')
        self.assertEqual(names.index('Round elbow, 90 deg'), 1)
        self.assertEqual(rows[1][2], '2')
        self.assertEqual(rows[1][3], '0.11')
        self.assertEqual(names[-1], 'Diffuser / grille')
        self.assertEqual(rows[-1][3], '0.050 each')
        self.assertAlmostEqual(sum(losses), 0.076, places=9)

    def test_rendered_instead_of_rows(self):
        rows, _l = dpc.takeoff_rows(self._items())
        cls = {'sys_class': 'Supply Air', 'airflow': '1,000 CFM', 'rows': [],
               'takeoff': rows, 'filter': None, 'subtotal': 0.076,
               'safety': 0.0076, 'safety_detail': '+10%', 'total': 0.0836,
               'total_detail': 'worst run'}
        html = dpc._esp_unit_table({'classes': [cls], 'esp_total': 0.0836,
                                    'esp_sides': ['Supply Air']})
        self.assertIn('Table / basis', html)
        self.assertIn('CD3-1', html)
        self.assertIn('0.084', html)


class WidthTests(unittest.TestCase):

    def test_widths_fix_layout_and_sum_to_100(self):
        html = cr.table(['A', 'B', 'C'], [['x', '1', 'y']], align=['l', 'r', 'l'],
                        widths=[1, 1, 1])
        self.assertIn('table-layout:fixed;width:100%;', html)
        pct = [int(v) for v in re.findall(r'<col style="width:(\d+)%;" />', html)]
        self.assertEqual(len(pct), 3)
        self.assertEqual(sum(pct), 100)
        self.assertEqual(shape(html).errors, [])

    def test_number_cells_never_wrap(self):
        html = cr.table(['A', 'B'], [['long name', '0.123']], align=['l', 'r'],
                        widths=[3, 1])
        self.assertIn('text-align:right;vertical-align:top;background:#FFFFFF;'
                      'white-space:nowrap;', html)
        self.assertIn('word-wrap:break-word;', html)

    def test_no_widths_keeps_auto_layout(self):
        html = cr.table(['A'], [['x']])
        self.assertNotIn('table-layout', html)
        self.assertNotIn('<colgroup>', html)

    def test_same_kind_tables_share_widths(self):
        def unit(name, item):
            return {'classes': [{'sys_class': name, 'rows': [(item, 0.1, '')],
                                 'subtotal': 0.1, 'safety': 0.01, 'total': 0.11}]}
        a = dpc._esp_unit_table(unit('Supply Air', 'Duct friction'))
        b = dpc._esp_unit_table(unit('Return Air', 'A much longer item name'))
        cols = lambda h: re.findall(r'<col style="width:\d+%;" />', h)
        self.assertEqual(cols(a), cols(b))
        self.assertTrue(cols(a))

    def test_band_row_spans_the_table(self):
        html = cr.table(['A', 'B'], [{'band': 'Group'}, ['x', '1']])
        self.assertIn('<td colspan="2"', html)
        self.assertEqual(shape(html).errors, [])


class SourceTests(unittest.TestCase):

    def test_no_f_strings_and_no_revit(self):
        for name in ('console_report.py', 'duct_perf_console.py'):
            with io.open(os.path.join(HERE, name), encoding='utf-8') as fh:
                src = fh.read()
            self.assertIsNone(re.search(r"""(?<![\w])[fF][rR]?['"]""", src), name)
            self.assertNotIn('Autodesk', src)
            self.assertNotIn('import clr', src)


if __name__ == '__main__':
    unittest.main()
