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
           'total': 0.8943, 'total_detail': 'worst run ends at CD-1 [12]'}
    return {'units': [{'label': 'RTU-1 (id 1)', 'classes': [cls],
                       'esp': 'EXTERNAL STATIC PRESSURE = 1.257 in. wc (Supply Air + Return Air)',
                       'esp_total': 0.8943, 'esp_sides': ['Supply Air'],
                       'one_sided': None,
                       'checks': [(True, 'Unit max 1.000 in. wc: computed 0.894, within it.')]}],
            'combined': [{'title': 'Duct 3 + Duct 4', 'text': 'EXTERNAL STATIC PRESSURE = 0.700'}],
            'method': 'METHOD: Darcy-Weisbach',
            'assumptions': ['ASSUMPTION: SWSI tables used for a DWDI fan'],
            'not_included': 'NOT INCLUDED: coil'}


class DuctPerfSectionTests(unittest.TestCase):

    def test_esp_unit_table_rows_and_total(self):
        p = shape(dpc.external_static(_sample_esp()))
        self.assertEqual(p.errors, [])
        body = [[c[2] for c in r] for r in p.rows]
        self.assertEqual(body[0], dpc.TAKEOFF_HEADERS)
        self.assertEqual([r[0] for r in body[1:]],
                         ['Supply Air, 950 CFM, 3 terminals', 'Duct friction',
                          'Fitting losses', 'Components', 'Supply Air subtotal',
                          'Safety factor', dpc.MR_LABEL])
        self.assertEqual(body[-1], [dpc.MR_LABEL, 'Supply Air',
                                    '', '', '0.894'])
        html = dpc.external_static(_sample_esp())
        self.assertIn('colspan="5"', html)
        self.assertIn(cr.CALLOUT_STYLES['ok'][1], html)       # passed unit max check
        self.assertNotIn(cr.CALLOUT_STYLES['warn'][1], html)  # no correction lines here
        self.assertIn('Duct 3 + Duct 4', html)
        # The basis stays with the take-off in the console (Colin, 2026-11-02).
        for basis in ('METHOD: Darcy-Weisbach', 'ASSUMPTION: SWSI', 'NOT INCLUDED: coil'):
            self.assertIn(basis, html)

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
        self.assertEqual(body[-1], [dpc.MR_LABEL,
                                    'Supply Air + Return Air', '', '', '0.800'])

    def test_ranked_run_rows(self):
        """Ranked run (Colin, 2026-11-10): the classes are the ranked runs,
        their total sits in a 'Ranked runs total' row, and the last row is
        still the MR total, the design ESP, naming where each MR run ends."""
        def cls(name, total, rank):
            return {'sys_class': name, 'rows': [('Duct friction', total / 1.1, '')],
                    'subtotal': total / 1.1, 'safety': total - total / 1.1,
                    'total': total, 'total_detail': '', 'rank': rank,
                    'rank_text': 'run {} of 12'.format(rank)}
        unit = {'label': 'RTU-1',
                'classes': [cls('Supply Air', 0.419, 2), cls('Return Air', 0.300, 3)],
                'esp_total': 0.719, 'esp_sides': ['Supply Air', 'Return Air'],
                'ranked': True, 'ranked_text': 'Supply Air #2 + Return Air #3',
                'mr_total': 0.739,
                'mr_text': 'Supply Air 0.427 at D-4 + Return Air 0.312 at RG-1'}
        html = dpc.external_static({'units': [unit]})
        body = [[c[2] for c in r] for r in shape(html).rows]
        self.assertIn('Supply Air, run 2 of 12', html)
        self.assertIn(['Supply Air subtotal, run #2', '', '', '', '0.381'], body)
        self.assertEqual([r[0] for r in body[-4:-1]],
                         ['Subtotal', 'Safety factor', dpc.RANKED_TOTAL_LABEL])
        self.assertEqual(body[-2], [dpc.RANKED_TOTAL_LABEL,
                                    'Supply Air #2 + Return Air #3, ' + dpc.RANKED_TOTAL_NOTE,
                                    '', '', '0.719'])
        self.assertEqual(body[-1], [dpc.MR_LABEL,
                                    'Supply Air 0.427 at D-4 + Return Air 0.312 at RG-1',
                                    '', '', '0.739'])
        # the sheet form draws both totals bold, the same cells
        rows, total = dpc.esp_unit_rows(unit, sheet=True)
        self.assertEqual(rows[-1], body[-2])
        self.assertEqual(total, body[-1])

    def test_safety_factor_once_at_the_end(self):
        """Colin, 2026-11-12: sides end at a subtotal, one Safety factor row
        near the end, and the total equals the old per side sum."""
        def cls(name, sub):
            return {'sys_class': name, 'rows': [('Duct friction', sub, '')],
                    'subtotal': sub, 'safety': sub * 0.1, 'safety_detail': '+10%',
                    'total': sub * 1.1, 'total_detail': ''}
        s, r = 0.500123, 0.295456
        old = s * 1.1 + r * 1.1
        two = {'classes': [cls('Supply Air', s), cls('Return Air', r)],
               'esp_total': old, 'esp_sides': ['Supply Air', 'Return Air'],
               'mr_text': 'Supply Air 0.500 at D-1, plus 10%'}
        one = {'classes': [cls('Supply Air', s)], 'esp_total': s * 1.1,
               'esp_sides': ['Supply Air']}
        for unit, want in ((two, old), (one, s * 1.1)):
            for sheet in (False, True):
                rows, total = dpc.esp_unit_rows(unit, sheet=sheet)
                cells = [r['cells'] if 'cells' in r else r for r in rows if not (isinstance(r, dict) and 'band' in r)]
                labels = [c[0] for c in cells if len(c) > 1]
                self.assertEqual(labels.count('Safety factor'), 1)
                self.assertFalse(any(l.endswith(' total') for l in labels))
                i = labels.index('Safety factor')
                self.assertEqual(cells[i][4], '%.3f' % (want - sum(
                    c['subtotal'] for c in unit['classes'])))
                self.assertAlmostEqual(float(unit['esp_total']), want, places=9)
                self.assertEqual(total[0], dpc.MR_LABEL)
        rows, _t = dpc.esp_unit_rows(two)
        self.assertIn(['Subtotal', 'Supply Air + Return Air', '', '', '0.796'], rows)
        rows, _t = dpc.esp_unit_rows(one)
        self.assertNotIn('Subtotal', [r[0] for r in rows if isinstance(r, list)])
        # a one side unit with no unit total row ends in its side total
        lone = {'classes': [cls('Supply Air', s)], 'esp_total': None}
        rows, t = dpc.esp_unit_rows(lone)
        self.assertIsNone(t)
        self.assertEqual([r[0] for r in rows[-2:]], ['Safety factor', 'Supply Air total'])
        self.assertEqual(rows[-1][4], '0.550')

    def test_tied_runs_and_marks(self):
        """A tie is the same printed total at 0.001 in. wc (Colin, 2026-11-11):
        M05C RG-9 0.3682 and RG-10 0.3681 tie, RG-7 0.3311 does not."""
        rg9 = {'terminal_id': 9, 'total_inwc': 0.3682}
        rg10 = {'terminal_id': 10, 'total_inwc': 0.3681}
        rg7 = {'terminal_id': 7, 'total_inwc': 0.3311}
        runs = [rg9, rg10, rg7]
        self.assertEqual(dpc.tied_runs(runs, rg9), [rg10])
        self.assertEqual(dpc.tied_runs(runs, rg10), [rg9])
        self.assertEqual(dpc.tied_runs(runs, rg7), [])
        self.assertEqual(dpc.tied_runs(None, rg9), [])
        self.assertEqual(dpc.and_join(['A']), 'A')
        self.assertEqual(dpc.and_join(['A', 'B']), 'A and B')
        self.assertEqual(dpc.and_join(['A', 'B', 'C']), 'A, B and C')
        self.assertEqual(len(dpc.mark_lines({'R'})), 1)
        self.assertTrue(dpc.mark_lines({'R'})[0].startswith('(R): Most restrictive'))
        lines = dpc.mark_lines({'R1', 'R2', 'R12'}, 'diffuser/grille')
        self.assertTrue(lines[0].startswith('(R1): Most restrictive diffuser/grille'))
        self.assertTrue(lines[1].startswith('(R2): 2nd most restrictive'))
        self.assertTrue(lines[2].startswith('(R12): 12th most restrictive'))
        self.assertEqual(dpc.mark_lines(set()), [])

    def test_mr_text_without_ranked(self):
        """Ranked run off, the MOST RESTRICTIVE row still names where each
        run ends, ties included."""
        unit = {'classes': [], 'esp_sides': ['Supply Air', 'Return Air'],
                'esp_total': 1.109,
                'mr_text': 'Supply Air 0.740 at D-FR1 + Return Air 0.368 at '
                           'RG-9 and RG-10, tied'}
        rows, total = dpc.esp_unit_rows(unit)
        self.assertEqual(total, [dpc.MR_LABEL, unit['mr_text'], '', '', '1.109'])

    def test_rank_helpers(self):
        self.assertEqual(dpc.parse_rank('3'), (3, None))
        self.assertEqual(dpc.parse_rank(' 12 '), (12, None))
        for bad in ('1', '0', '-2', '2.5', 'x', ''):
            v, err = dpc.parse_rank(bad)
            self.assertIsNone(v)
            self.assertTrue(err)
        runs = ['r1', 'r2', 'r3']
        self.assertEqual(dpc.pick_ranked(runs, 2), ('r2', 2, 3, False))
        self.assertEqual(dpc.pick_ranked(runs, 3), ('r3', 3, 3, False))
        self.assertEqual(dpc.pick_ranked(runs, 7), ('r3', 3, 3, True))
        self.assertEqual(dpc.pick_ranked([], 2), (None, 0, 0, False))
        self.assertEqual(dpc.rank_over_text('Supply', 7, 3),
                         'Supply rank 7 is above the 3 diffuser(s)/grille(s) on this '
                         'system; run 3 (the least restrictive) is shown.')

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

    def test_grand_rows_per_system(self):
        """Colin, 2026-11-09: one CFM line per system present, never summed,
        then Total Diffusers/Grilles and Ducts."""
        rows = dpc.grand_rows({'Return Air': 1200.0, 'Supply Air': 5250.0,
                               'Outside Air': 0.0}, 18, 60)
        self.assertEqual(rows, [('Supply CFM', '5,250'), ('Return CFM', '1,200'),
                                ('Total Diffusers/Grilles', '18'), ('Ducts', '60')])
        self.assertEqual(dpc.grand_rows({'Exhaust Air': 700.0}, 5, 9)[0],
                         ('Exhaust CFM', '700'))
        self.assertIn(('Transfer Air CFM', '150'),
                      dpc.grand_rows({'Supply Air': 900.0, 'Transfer Air': 150.0}, 4, 6))
        self.assertEqual(dpc.summary_band_text({'grand': rows}),
                         'Supply CFM 5,250 | Return CFM 1,200 | '
                         'Total Diffusers/Grilles 18 | Ducts 60')

    def test_scope_wording(self):
        """Colin, 2026-11-09: the airflow words follow the Scope (an exhaust
        run never reads Supply, and drops the other air column); the count
        reads Diffusers/Grilles on every Scope."""
        legend = {'limits': [], 'legend': [['Red', '1', '0', 'x']]}
        count = 'Diffusers/Grilles'
        exp = {'rtu': 'Supply CFM', 'vav': 'Supply CFM', 'vavbox': 'Box discharge CFM',
               'exhaust': 'Exhaust CFM'}
        for scope, air in exp.items():
            w = dpc.scope_words(scope)
            row = ['EF-1 (id 1)', '700', '5', '700 CFM Exhaust Air']
            html = dpc.system_summary({'grand': [], 'equipment': [row], 'words': w})
            self.assertIn(air, html)
            self.assertIn(count, html)
            self.assertIn('Fittings, Accessories &amp; ' + count,
                          dpc.limits_and_legend(dict(legend, words=w)))
            self.assertEqual(dpc.scope_words(scope)['types'], 'Diffuser/Grille type')
            self.assertEqual(dpc.scope_words(scope)['one'], 'diffuser/grille')
            self.assertNotIn('tables', dpc.scope_words(scope))
            if scope == 'exhaust':
                self.assertNotIn('Supply', html)
                self.assertNotIn('700 CFM Exhaust Air', html)
                self.assertEqual(len(shape(html).rows[0]), 3)
            else:
                self.assertEqual(len(shape(html).rows[0]), 4)
        self.assertEqual(dpc.scope_words('nonsense'), dpc.scope_words('rtu'))

    def test_branch_notes_and_clearance_per_scope(self):
        """Colin, 2026-11-09 (TB-M10C rerun): terminal wording, "it is more
        accurate". Exhaust System leaves out the auto-sizing diffuser and slot
        notes; the other Scopes name diffusers only in those."""
        ex = dpc.branch_notes(dpc.scope_words('exhaust'))
        text = ' '.join(a + ' ' + b for a, b in ex)
        self.assertIn('terminal sizing tables', text)
        self.assertIn("whether the terminal's own neck/face", text)
        self.assertNotIn('Diffuser', text.replace('diffusers and grilles', ''))
        self.assertNotIn('diffuser', text.replace('diffusers and grilles', ''))
        self.assertNotIn('slot', text)
        self.assertEqual(dpc.clearance_text(0),
                         'Terminal/duct height clearance issues (flagged Red): 0')
        for scope in ('rtu', 'vav', 'vavbox'):
            notes = dpc.branch_notes(dpc.scope_words(scope))
            self.assertEqual(len(notes), 3)
            self.assertIn('terminal sizing tables', notes[0][0])
            self.assertIn("whether the terminal's own neck/face", notes[0][1])
            self.assertTrue(notes[1][0].startswith('Branches feeding a slot diffuser'))
            self.assertTrue(notes[2][0].startswith('Branches feeding a terminal'))
        html = dpc.limits_and_legend({'limits': [], 'legend': [], 'notes': ex,
                                      'words': dpc.scope_words('exhaust'),
                                      'clearance': (False, dpc.clearance_text(0))})
        self.assertIn('Terminal/duct height clearance', html)
        self.assertNotIn('slot', html)

    def test_takeoff_group_rows_read_as_covered(self):
        """Colin, 2026-11-09: the group's straight duct and elbow print as
        their own 0.000 rows marked 'in the fan system effect', apart from the
        other straight duct and elbows, and the FSE row carries the value."""
        items = [
            ('Duct 1', '?, 5.0 ft, in the fan system effect', 1000.0, 0.0,
             ('Straight duct', 'Darcy-Weisbach, in the fan system effect', None, 5.0)),
            ('Duct 3', '?, 10.0 ft', 1000.0, 0.010,
             ('Straight duct', 'Darcy-Weisbach', None, 10.0)),
            ('Fitting 2', 'x, in the fan system effect', 1000.0, 0.0,
             ('Rect radius elbow, 90 deg', 'CR3-1, in the fan system effect', 0.25, None)),
            ('Fitting 4', 'x', 1000.0, 0.0156,
             ('Rect radius elbow, 90 deg', 'CR3-1', 0.25, None)),
            ('Fan system effect 2', 'x', 1000.0, 0.316,
             ('Fan system effect, supply outlet',
              'Duct losses: straight duct and elbow (CR3-1), L 5.1 / Le 5.7 ft', None, None)),
        ]
        rows, losses = dpc.takeoff_rows(items)
        # Route order (2026-11-11): the FSE row first, its group under it.
        self.assertEqual(rows[0][0], 'Fan system effect, supply outlet')
        self.assertEqual((rows[0][3], rows[0][4]), ('-', '0.316'))
        self.assertEqual(rows[1], ['Straight duct',
                                   'Darcy-Weisbach, 5.0 ft, in the fan system effect',
                                   '', '-', '0.000'])
        self.assertEqual(rows[2], ['Rect radius elbow, 90 deg',
                                   'CR3-1, id 2, in the fan system effect',
                                   '1', '0.25', '0.000'])
        self.assertEqual(rows[3], ['Straight duct', 'Darcy-Weisbach, 10.0 ft', '', '-',
                                   '0.010'])
        self.assertEqual(rows[4], ['Rect radius elbow, 90 deg', 'CR3-1, id 4', '1', '0.25',
                                   '0.016'])
        self.assertAlmostEqual(sum(losses), 0.010 + 0.0156 + 0.316)

    def test_esp_header_basis(self):
        """Colin, 2026-11-09: the 2001 tables by geometry, the worksheet only
        as the fallback."""
        html = dpc.external_static({'units': []})
        self.assertIn('2001 ASHRAE Fundamentals Ch. 34 by geometry', html)
        self.assertIn('SP_LOSS_WORKSHEET as fallback', html)
        self.assertNotIn('per RJA SP_LOSS_WORKSHEET', html)

    def test_summary_and_legend_and_sheets(self):
        rec = {'grand': [('Grand total airflow', '5,250 CFM')],
               'equipment': [['RTU-1 (id 1)', '950', '3', '-']]}
        for html in (dpc.system_summary(rec),
                     dpc.limits_and_legend({'limits': [['Supply Air', '800 FPM', '0.080']],
                                            'legend': [['Red', '3', '1', 'exceeds']],
                                            'clearance': (True, 'clearance: 2')}),
                     dpc.sheets_created(['M-601']), dpc.sheets_created([]),
                     dpc.system_corrections(['#', 'Status'], [['1', 'RED']], [],
                                            status_col=1)):
            self.assertEqual(shape(html).errors, [])

    def test_legend_has_no_clearance_notes(self):
        html = dpc.limits_and_legend({'limits': [], 'legend': [],
                                      'clearance': (True, 'clearance: 1'),
                                      'clearance_notes': ['NOTE: main duct 5 steps down']})
        self.assertIn('clearance: 1', html)
        self.assertNotIn('steps down', html)


def _entries():
    c = dpc.correction
    return [
        c('warn', '1 terminal(s) have Flow = 0. Assign CFM values in the model.', 'RTU-2 (id 2)'),
        c('note', 'NOTE: unit 1 connection at element 9 is a single diffuser main.', 'RTU-1 (id 1)'),
        c('gray', 'GRAY: System Effect Pressure Loss too high at elbow 921.', 'RTU-1 (id 1)'),
        c('warn', 'FALLBACK: tap 5 passed on duct 4, RJA worksheet C 0.20', 'RTU-1 (id 1)'),
        c('warn', 'NOTE: transition 7: angle below the lowest table value', 'RTU-1 (id 1)'),
        c('warn', 'NOTE: main duct 5 steps down by only 1 in. of height.', 'RTU-1 (id 1)',
          dpc.ACTION),
        c('error', 'Unit max external static 0.750 in. wc: computed 0.850, EXCEEDS IT.',
          'RTU-1 (id 1)'),
        c('note', 'NOTE: oval duct 8 read as rectangular W x H.'),
        c('warn', 'More than one system was selected.'),
        # duplicates: same unit and text said twice (two classes, two passes)
        c('warn', 'FALLBACK: tap 5 passed on duct 4, RJA worksheet C 0.20', 'RTU-1 (id 1)'),
        c('gray', 'GRAY: System Effect Pressure Loss too high at elbow 921.', 'RTU-1 (id 1)'),
        # the same text for a different unit is not a duplicate
        c('warn', '1 terminal(s) have Flow = 0. Assign CFM values in the model.', 'RTU-1 (id 1)'),
    ]


class GroupLinesTests(unittest.TestCase):
    """Colin, 2026-11-11: repeated lines of a unit grouped, one per (table,
    reason) for the tap FALLBACKs, one NOTE for the table edge reads, one per
    duct for unreadable tap positions."""

    def test_fallback_taps_grouped_by_table_and_reason(self):
        lines = [
            'FALLBACK: tap 3641127 passed on duct 3640926, 2001 ED5-2 Cc,s -0.29 is not '
            'positive; RJA worksheet C 0.60 used',
            'FALLBACK: tap 3641113 passed on duct 3640926, 2001 ED5-2 Cc,s -0.17 is not '
            'positive; RJA worksheet C 0.60 used',
            'FALLBACK: tap 3641141 passed on duct 3640923, 2001 ED5-2 Cc,s -0.25 is not '
            'positive; RJA worksheet C 0.60 used',
            'FALLBACK: tap 501 into branch off duct 500, 2001 ER5-3 Cc,b -0.10 is not '
            'positive; RJA worksheet C 0.60 used',
            'FALLBACK: tap 7 passed on duct 6, RJA worksheet C 0.20 (main airflow unreadable)',
            'FALLBACK: tap 8 passed on duct 6, RJA worksheet C 0.20 (main airflow unreadable)',
        ]
        self.assertEqual(dpc.group_lines(lines), [
            'FALLBACK: taps 3641113, 3641127 (duct 3640926), 3641141 (duct 3640923): 2001 '
            'ED5-2 Cc,s not positive; RJA worksheet C 0.60 used',
            'FALLBACK: tap 501 (duct 500): 2001 ER5-3 Cc,b not positive; RJA worksheet C '
            '0.60 used',
            'FALLBACK: taps 7, 8 (duct 6): RJA worksheet C 0.20 (main airflow unreadable)'])

    def test_table_edge_and_blank_cell_one_note_each(self):
        lines = [
            'NOTE: transition 3640946 (Supply Air, 12x12" to 10x12", 7.6 deg): SR4-1 angle '
            '7.6 deg at 10.0 deg (table edge)',
            'FALLBACK: tap 9 into branch off duct 4, 2001 SR5-11 Cc,b -0.10 is not positive; '
            'RJA worksheet C 0.60 used; NOTE: tap 3764706 on main 3764127 (Supply Air 6x12"): '
            'SR5-11 Ab/Ac 1.09 at 0.90 (table edge); tap 3764706 on main 3764127 (Supply Air '
            '6x12"): ER5-3 is printed for Ab/Ac 0.50 only and is read there; this fitting is '
            'Ab/Ac 0.35',
            'NOTE: elbow 12 (Supply Air, 10"dia, 90 deg): CD3-12 next to a blank cell',
            'NOTE: transition 3640946 (Supply Air, 12x12" to 10x12", 7.6 deg): SR4-1 angle '
            '7.6 deg at 10.0 deg (table edge)',
        ]
        self.assertEqual(dpc.group_lines(lines), [
            'NOTE: read at the table edge: transition 3640946 SR4-1 angle 7.6 deg at 10.0 deg; '
            'tap 3764706 SR5-11 Ab/Ac 1.09 at 0.90',
            'FALLBACK: tap 9 (duct 4): 2001 SR5-11 Cc,b not positive; RJA worksheet C 0.60 used',
            'NOTE: tap 3764706 on main 3764127 (Supply Air 6x12"): ER5-3 is printed for Ab/Ac '
            '0.50 only and is read there; this fitting is Ab/Ac 0.35',
            'NOTE: read next to a blank cell, nearest printed value used: elbow 12 CD3-12'])

    def test_repeats_once_and_other_lines_untouched(self):
        t = ('NOTE: tap positions on duct 5 unreadable: every tap priced at the full duct '
             'flow, friction charged over its whole length at its unit end cfm')
        self.assertEqual(dpc.group_lines([t, 'WARNING: x', t]), [t, 'WARNING: x'])

    def test_corrections_group_per_unit(self):
        c = dpc.correction
        fb = ('FALLBACK: tap %d passed on duct 4, 2001 ED5-2 Cc,s -0.20 is not positive; '
              'RJA worksheet C 0.60 used')
        g = dpc.correction_lines([c('warn', fb % 1, 'A'), c('warn', fb % 2, 'A'),
                                  c('warn', fb % 3, 'B')])
        self.assertEqual([line for _k, line in g[dpc.ACTION]], [
            'A: FALLBACK: taps 1, 2 (duct 4): 2001 ED5-2 Cc,s not positive; RJA worksheet '
            'C 0.60 used',
            'B: FALLBACK: tap 3 (duct 4): 2001 ED5-2 Cc,s not positive; RJA worksheet C '
            '0.60 used'])


class SystemCorrectionsTests(unittest.TestCase):
    """Colin, 2026-11-02: every correction in one System Corrections section,
    Action required then Notes, each line under its unit, no duplicates."""

    def test_grouping(self):
        g = dpc.correction_group
        self.assertEqual(g('warn', 'anything'), dpc.ACTION)
        self.assertEqual(g('gray', 'GRAY: x'), dpc.ACTION)
        self.assertEqual(g('error', 'Unit max x'), dpc.ACTION)
        self.assertEqual(g('note', 'WARNING: x'), dpc.ACTION)
        self.assertEqual(g('note', 'FALLBACK: x'), dpc.ACTION)
        self.assertEqual(g('note', 'UNPRICED: x'), dpc.ACTION)
        self.assertEqual(g('note', 'LOW: x'), dpc.ACTION)
        self.assertEqual(g('warn', 'NOTE: x'), dpc.NOTES)    # its own lead wins
        self.assertEqual(g('note', 'plain'), dpc.NOTES)
        # forced: a clearance NOTE asks for a resize
        e = dpc.correction('note', 'NOTE: steps down', 'U', dpc.ACTION)
        self.assertEqual(e, (dpc.ACTION, 'U', 'warn', 'NOTE: steps down'))

    def test_lines_grouped_prefixed_ordered_and_deduped(self):
        g = dpc.correction_lines(_entries())
        action = [line for _k, line in g[dpc.ACTION]]
        notes = [line for _k, line in g[dpc.NOTES]]
        self.assertEqual(action, [
            'WARNING: More than one system was selected.',
            'RTU-2 (id 2): WARNING: 1 terminal(s) have Flow = 0. Assign CFM values in the model.',
            'RTU-1 (id 1): GRAY: System Effect Pressure Loss too high at elbow 921.',
            'RTU-1 (id 1): FALLBACK: tap 5 passed on duct 4, RJA worksheet C 0.20',
            'RTU-1 (id 1): NOTE: main duct 5 steps down by only 1 in. of height.',
            'RTU-1 (id 1): CHECK FAILED: Unit max external static 0.750 in. wc: computed '
            '0.850, EXCEEDS IT.',
            'RTU-1 (id 1): WARNING: 1 terminal(s) have Flow = 0. Assign CFM values in the model.',
        ])
        self.assertEqual(notes, [
            'NOTE: oval duct 8 read as rectangular W x H.',
            'RTU-1 (id 1): NOTE: unit 1 connection at element 9 is a single diffuser main.',
            'RTU-1 (id 1): NOTE: transition 7: angle below the lowest table value',
        ])

    def test_console_section(self):
        html = dpc.system_corrections(['#', 'Status'], [['1', 'RED']], _entries(),
                                      status_col=1)
        p = shape(html)
        self.assertEqual(p.errors, [])
        self.assertIn('System Corrections', html)
        self.assertNotIn('Flagged Ducts', html)
        self.assertLess(html.index('Action required'), html.index('>Notes<'))
        self.assertLess(html.index('<table'), html.index('Action required'))
        self.assertEqual(html.count('elbow 921'), 1)
        self.assertEqual(html.count('worksheet C 0.20'), 1)
        self.assertIn(cr.CALLOUT_STYLES['gray'][1], html)
        self.assertIn(cr.CALLOUT_STYLES['warn'][1], html)
        self.assertIn(cr.CALLOUT_STYLES['error'][1], html)
        # Notes are muted lines, not callouts
        i = html.index('single diffuser main')
        self.assertIn(cr.MUTED, html[html.rindex('<div', 0, i):i])

    def test_no_corrections(self):
        html = dpc.system_corrections(['#'], [], [])
        self.assertIn('System Corrections', html)
        self.assertIn(dpc.NO_CORRECTIONS, html)
        self.assertNotIn('<table', html)
        self.assertEqual(dpc.corrections_text([]), [dpc.NO_CORRECTIONS])
        self.assertEqual(dpc.corrections_sheet([]), [])

    def test_corrections_without_flagged_ducts(self):
        html = dpc.system_corrections(['#'], [], _entries()[:1])
        self.assertIn(dpc.NO_FLAGGED, html)
        self.assertNotIn(dpc.NO_CORRECTIONS, html)
        self.assertIn('Action required', html)
        self.assertNotIn('>Notes<', html)

    def test_text_fallback_and_sheet(self):
        text = dpc.corrections_text(_entries(), any_rows=True)
        self.assertEqual(text[0], '**Action required**')
        self.assertIn('**Notes**', text)
        self.assertTrue(all(t.startswith(('**', '- ')) for t in text))
        sheet = dpc.corrections_sheet(_entries())
        self.assertEqual([h for h, _l in sheet], ['ACTION REQUIRED', 'NOTES'])
        self.assertEqual(len(sheet[0][1]), 7)
        self.assertEqual(len(sheet[1][1]), 3)

    def test_vav_corrections(self):
        rep = {'minimum': 0.75, 'units': [
            {'label': 'RTU-1 (id 1)', 'mp_terminal_text': '1 diffuser(s) on an MP main.',
             'boxes': [{'box': 101, 'status': 'OK', 'left': 1.0},
                       {'box': 102, 'status': 'LOW', 'left': 0.07}]}]}
        ents = dpc.vav_corrections(rep)
        self.assertEqual([e[0] for e in ents], [dpc.ACTION, dpc.ACTION])
        self.assertEqual(ents[0][3], 'LOW: VAV box 102, pressure left 0.070 in. wc, '
                                     'under the 0.75 in. wc minimum.')
        self.assertEqual(ents[0][1], 'RTU-1 (id 1)')
        # the VAV section keeps only its table key
        html = dpc.vav_boxes(rep)
        self.assertNotIn('MP main', html)
        self.assertIn('Values in in. wc.', html)


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

    def test_route_order_one_row_each(self):
        """Colin, 2026-11-11: route order from the unit, one row per item,
        no grouping, the element id in Table / basis."""
        rows, losses = dpc.takeoff_rows(self._items())
        self.assertEqual([r[0] for r in rows], [
            'Straight duct', 'Round elbow, 90 deg', 'Straight duct',
            'Round elbow, 90 deg', 'Diffuser / grille',
            'Tap passed on main (straight through)'])
        self.assertEqual(rows[0], ['Straight duct', 'Darcy-Weisbach, 10.0 ft', '', '-',
                                   '0.010'])
        self.assertEqual(rows[1], ['Round elbow, 90 deg', 'CD3-1, id 2', '1', '0.11',
                                   '0.004'])
        self.assertEqual(rows[3][1], 'CD3-1, id 4')
        self.assertEqual(rows[4], ['Diffuser / grille', 'known drop, id 5', '1', '-',
                                   '0.050'])
        self.assertEqual(rows[5][1], 'SR5-13 Cs, on duct 1')
        self.assertAlmostEqual(sum(losses), 0.076, places=9)

    def test_consecutive_duct_merged(self):
        """Consecutive duct pieces with no fitting between are one row,
        length and loss summed; flex stays its own row."""
        items = [
            ('Duct 1', 'a', 800.0, 0.010, ('Straight duct', 'Darcy-Weisbach', None, 10.0)),
            ('Duct 2', 'b', 800.0, 0.004, ('Straight duct', 'Darcy-Weisbach', None, 2.5)),
            ('Fitting 3', 'x', 800.0, 0.003, ('Round elbow, 90 deg', 'CD3-1', 0.11, None)),
            ('Duct 4', 'c', 700.0, 0.002, ('Straight duct', 'Darcy-Weisbach', None, 3.0)),
            ('Duct 5', 'd', 700.0, 0.006, ('Flex duct', 'Darcy-Weisbach', None, 5.0)),
            ('Flex bend', 'C 3.30', 700.0, 0.100,
             ('Flex bend, one per flex run', 'DSDG Ch. 2', 3.3, None)),
            ('Duct 6', 'e', 700.0, 0.001, ('Flex duct', 'Darcy-Weisbach', None, 1.0)),
        ]
        rows, losses = dpc.takeoff_rows(items)
        self.assertEqual([r[0] for r in rows], [
            'Straight duct', 'Round elbow, 90 deg', 'Straight duct', 'Flex duct',
            'Flex bend, one per flex run', 'Flex duct'])
        self.assertEqual(rows[0][1:], ['Darcy-Weisbach, 12.5 ft', '', '-', '0.014'])
        self.assertAlmostEqual(losses[0], 0.014)
        self.assertEqual(rows[3][1], 'Darcy-Weisbach, 5.0 ft')
        self.assertAlmostEqual(sum(losses), sum(it[3] for it in items))

    def test_filter_under_fse_rows(self):
        """The return filter sits at the unit end, right under the fan
        system effect rows, and the rows still add up to the subtotal."""
        items = [
            ('Fan system effect 9', 'x', 600.0, 0.020,
             ('Fan system effect, return inlet', 'Fan system effect: ED7-2', 0.5, None)),
            ('Duct 1', 'a, in the fan system effect', 600.0, 0.0,
             ('Straight duct', 'Darcy-Weisbach, in the fan system effect', None, 2.0)),
            ('Duct 2', 'b', 600.0, 0.003, ('Straight duct', 'Darcy-Weisbach', None, 8.0)),
            ('Terminal 5', 'RG', None, 0.05, ('Diffuser / grille', 'known drop', None, None)),
        ]
        rows, losses = dpc.takeoff_rows(items)
        sub = sum(losses) + 0.14
        cls = {'sys_class': 'Return Air', 'rows': [], 'takeoff': rows, 'filter': 0.14,
               'subtotal': sub, 'safety': sub * 0.1, 'total': sub * 1.1}
        out = dpc._class_rows(cls, True)
        lists = [r for r in out if isinstance(r, list)]
        self.assertEqual([r[0] for r in lists[:4]], [
            'Fan system effect, return inlet', 'Straight duct', 'Return filter',
            'Straight duct'])
        body = lists[:5]
        self.assertEqual(lists[5][0], 'Return Air subtotal')
        self.assertAlmostEqual(sum(float(r[4]) for r in body), sub, places=2)
        self.assertEqual(lists[5][4], '%.3f' % sub)

    def test_discharge_after_intake(self):
        """An inline fan's discharge follows the intake route, the assumed
        backdraft damper leading it and the roof cap last; the discharge FSE
        row stays at the top with its own group under it."""
        items = [
            ('Fan system effect 7', 'x', 1000.0, 0.03,
             ('Fan system effect, exhaust inlet', 'Fan system effect: ED7-2', 0.4, None)),
            ('Fan system effect 8', 'y', 1000.0, 0.04,
             ('Fan system effect, exhaust outlet', 'Fan system effect: SR7-11', 0.5, None)),
            ('Duct 1', 'a, in the fan system effect', 900.0, 0.0,
             ('Straight duct', 'Darcy-Weisbach, in the fan system effect', None, 1.0)),
            ('Duct 2', 'b', 900.0, 0.002, ('Straight duct', 'Darcy-Weisbach', None, 6.0)),
            ('Discharge duct 11', 'c, in the fan system effect', 1000.0, 0.0,
             ('Discharge duct', 'Darcy-Weisbach, in the fan system effect', None, 1.5)),
            ('Discharge fitting 12', 'd', 1000.0, 0.004,
             ('Discharge round elbow, 90 deg', 'CD3-1', 0.11, None)),
            ('Accessory (assumed)', 'Backdraft damper at the fan', None, 0.1,
             ('Backdraft damper', 'known drop', None, None)),
            ('Terminal 5', 'EG', None, 0.05, ('Diffuser / grille', 'known drop', None, None)),
            ('Exhaust roof cap', 'Include exhaust roof cap', None, 0.06,
             ('Roof cap', 'known drop', None, None)),
        ]
        rows, losses = dpc.takeoff_rows(items)
        self.assertEqual([r[0] for r in rows], [
            'Fan system effect, exhaust inlet', 'Straight duct',
            'Fan system effect, exhaust outlet', 'Discharge duct',
            'Straight duct', 'Diffuser / grille', 'Backdraft damper',
            'Discharge round elbow, 90 deg', 'Roof cap'])
        self.assertEqual(rows[6][1], 'known drop, assumed')
        self.assertEqual(rows[7][1], 'CD3-1, id 12')
        self.assertAlmostEqual(sum(losses), sum(it[3] for it in items))

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
