# -*- coding: ascii -*-
"""CPython tests for the Duct Performance sheet tables (Colin, 2026-10-23).

rja_drafting imports Autodesk.Revit.DB at module scope, so a stub package is
put in sys.modules first; the drawing primitives (_t, _fill, _line) are then
replaced with recorders, so _styled_table, _note_panel and _band run their
real layout code and every text they would place is captured with its cell.

The sheet section drawers (_draw_system_summary, _draw_external_static,
_draw_vav_boxes, _sheet_columns and the _SHEET_* constants) live in the
Duct Performance pushbutton script, which imports Revit and pyRevit, so their
source is pulled out with a regex and exec()'d against the real
duct_perf_console and the patched rja_drafting, the same technique as
_test_critical_path.py.

Run:  python _test_duct_perf_sheet.py
"""

import io
import os
import re
import sys
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def _stub_revit():
    names = ('XYZ Line TextNote TextNoteOptions TextNoteType ViewDrafting '
             'ViewFamily ViewFamilyType Viewport FilteredElementCollector '
             'BuiltInCategory CurveLoop FilledRegion FilledRegionType '
             'OverrideGraphicSettings ElementId Color FillPatternElement '
             'BuiltInParameter HorizontalTextAlignment').split()
    pkg = types.ModuleType('Autodesk')
    revit = types.ModuleType('Autodesk.Revit')
    db = types.ModuleType('Autodesk.Revit.DB')
    for n in names:
        setattr(db, n, type(n, (object,), {}))
    pkg.Revit = revit
    revit.DB = db
    sys.modules.setdefault('Autodesk', pkg)
    sys.modules.setdefault('Autodesk.Revit', revit)
    sys.modules.setdefault('Autodesk.Revit.DB', db)


_stub_revit()
import rja_drafting as rd          # noqa: E402
import duct_perf_console as dpc    # noqa: E402


class Recorder(object):
    """Replaces the Revit drawing calls with lists."""

    def __init__(self):
        self.texts = []    # (x, y, value, kind, align)
        self.fills = []    # (x0, top, x1, bottom, rgb)
        self.lines = []

    def install(self):
        self._saved = (rd._t, rd._fill, rd._line)
        rec = self

        def _t(st, view, x, y, value, kind="text", align="left"):
            if value:
                rec.texts.append((x, y, str(value), kind, align))
            return None

        def _fill(st, view, x0, top, x1, bottom, rgb):
            rec.fills.append((x0, top, x1, bottom, rgb))

        def _line(st, view, x0, y0, x1, y1, weight=None):
            rec.lines.append((x0, y0, x1, y1, weight))

        rd._t, rd._fill, rd._line = _t, _fill, _line
        return self

    def restore(self):
        rd._t, rd._fill, rd._line = self._saved

    def values(self):
        return [t[2] for t in self.texts]


SCRIPT = os.path.join(os.path.dirname(HERE), 'Developer Tools.tab', 'Performance.panel',
                      'Duct Performance.pushbutton', 'script.py')


def _load_drawers():
    src = io.open(SCRIPT, encoding='utf-8').read()
    m = re.search(r'^# Sheet column proportions.*?(?=^def _build_summary_view\()',
                  src, re.S | re.M)
    if not m:
        raise SystemExit('Sheet drawers not found in script.py')
    ns = {'rja_drafting': rd, 'duct_perf_console': dpc}
    exec(m.group(0), ns)
    return ns


DRAW = _load_drawers()
TOTAL_W = 1.105    # the flagged duct table width with the default columns
LONG_TERMINAL = 'RJA - Supply Diffuser - Linear Slot (CFM Auto-Sizing) [3664128]'


def _cls(name, total, takeoff=True):
    rows = [['Straight duct', 'Darcy-Weisbach', '12 (240 ft)', '-', '0.180'],
            ['Tap passed on main (straight through)', 'SR5-13 Cs', '6', '0.04', '0.012']]
    return {'sys_class': name, 'airflow': '2,450 CFM, 14 terminals',
            'rows': [], 'takeoff': rows if takeoff else None,
            'filter': 0.14 if name == 'Return Air' else None,
            'subtotal': total / 1.1, 'safety': total - total / 1.1,
            'safety_detail': '+10%', 'total': total,
            'total_detail': 'worst run ends at ' + LONG_TERMINAL}


def _unit():
    return {'label': 'RTU-1 (id 123456)',
            'classes': [_cls('Supply Air', 0.55), _cls('Return Air', 0.30),
                        _cls('Exhaust Air', 0.10)],
            'one_sided': None,
            'checks': [(True, 'Unit max external static 1.000 in. wc: computed 0.850 '
                              'in. wc (Supply Air + Return Air), within it.')],
            'esp_total': 0.85, 'esp_sides': ['Supply Air', 'Return Air']}


def _corrections():
    c = dpc.correction
    return [c('warn', 'UNPRICED: Fitting 777', 'RTU-1 (id 123456)'),
            c('warn', 'Parallel units on one main. Its external static is unreliable.',
              'RTU-1 (id 123456)'),
            c('note', 'NOTE: oval duct 8 read as rectangular W x H.'),
            c('warn', 'UNPRICED: Fitting 777', 'RTU-1 (id 123456)')]


class FitTests(unittest.TestCase):
    """Colin, 2026-11-09: a column widens to its longest text, never clipped
    with '...' and never wrapped, however wide the table gets."""

    def test_no_clip_left(self):
        self.assertFalse(hasattr(rd, '_clip'))

    def test_text_width(self):
        self.assertAlmostEqual(rd.text_width('abcd'), 4 * rd.CHAR_W + 2 * rd.PAD)
        self.assertGreater(rd.text_width('x' * 50, True), rd.text_width('x' * 50))

    def test_fit_shares_target_when_text_fits(self):
        w = rd.fit_widths([0.05, 0.05], [0.5, 0.5], 1.0)
        self.assertAlmostEqual(w[0], 0.5)
        self.assertAlmostEqual(sum(w), 1.0)

    def test_fit_holds_need_and_keeps_target(self):
        w = rd.fit_widths([0.7, 0.05, 0.05], [0.34, 0.33, 0.33], 1.0)
        self.assertAlmostEqual(w[0], 0.7)
        self.assertAlmostEqual(sum(w), 1.0)
        self.assertAlmostEqual(w[1], w[2])

    def test_fit_widens_past_target(self):
        w = rd.fit_widths([0.8, 0.5], [0.5, 0.5], 1.0)
        self.assertEqual(w, [0.8, 0.5])

    def test_column_needs_bold_and_band(self):
        rows = [{'band': 'B' * 100}, ['abc', '1'], {'cells': ['x' * 30, '2'], 'bold': True}]
        needs, band = rd.column_needs(['Item', 'in. wc'], rows)
        self.assertAlmostEqual(needs[0], rd.text_width('x' * 30, True))
        self.assertAlmostEqual(band, rd.text_width('B' * 100, True))
        self.assertAlmostEqual(rd.table_need(['Item', 'in. wc'], rows), band)


class StyledTableTests(unittest.TestCase):

    def setUp(self):
        self.rec = Recorder().install()

    def tearDown(self):
        self.rec.restore()

    def test_plain_rows_unchanged_behaviour(self):
        bottom = rd._styled_table(None, None, 0.0, 0.0, [('A', 0.2), ('B', 0.1)],
                                  [['x', '1'], ['y', '2'], ['z', '3']])
        self.assertAlmostEqual(bottom, -(rd.HEAD_H + 3 * rd.ROW_H))
        zebra = [f for f in self.rec.fills if f[4] == rd.ZEBRA_FILL]
        self.assertEqual(len(zebra), 1)   # only the second body row
        self.assertEqual(self.rec.values(), ['A', 'B', 'x', '1', 'y', '2', 'z', '3'])

    def test_band_and_bold_rows(self):
        rows = [{'band': 'Supply Air'}, ['a', '1'], ['b', '2'],
                {'cells': ['Supply Air total', '3'], 'bold': True},
                {'band': 'Return Air'}, ['c', '4'], ['TOTAL', '7']]
        bottom = rd._styled_table(None, None, 0.0, 0.0, [('A', 0.3), ('B', 0.1)], rows,
                                  totals_row=True)
        self.assertAlmostEqual(bottom, -(rd.HEAD_H + 7 * rd.ROW_H))
        kinds = dict((t[2], t[3]) for t in self.rec.texts)
        self.assertEqual(kinds['Supply Air'], 'bold')
        self.assertEqual(kinds['Supply Air total'], 'bold')
        self.assertEqual(kinds['a'], 'text')
        self.assertEqual(kinds['TOTAL'], 'bold')
        heads = [f for f in self.rec.fills if f[4] == rd.HEAD_FILL]
        self.assertEqual(len(heads), 3)   # header + two band rows
        # banding restarts under each heading: 'b' and 'Supply Air total'
        # rows striped, 'c' (first under Return Air) not
        zebra_tops = sorted(round(f[1], 6) for f in self.rec.fills if f[4] == rd.ZEBRA_FILL)
        row_top = lambda r: round(-rd.HEAD_H - r * rd.ROW_H, 6)
        self.assertEqual(zebra_tops, sorted([row_top(2)]))
        # no column divider crosses a heading row
        x_div = 0.3
        for x0, y0, x1, y1, _w in self.rec.lines:
            if abs(x0 - x_div) < 1e-9 and abs(x1 - x_div) < 1e-9:
                lo, hi = min(y0, y1), max(y0, y1)
                for r in (0, 4):
                    mid = row_top(r) - rd.ROW_H / 2.0
                    self.assertFalse(lo < mid < hi)

    def test_fit_widens_instead_of_clipping(self):
        rd._styled_table(None, None, 0.0, 0.0, [('Item', 0.2), ('in. wc', 0.08)],
                         [[LONG_TERMINAL, '0.123'], {'band': 'Supply Air, ' + 'z' * 120}],
                         fit=True)
        self.assertIn(LONG_TERMINAL, self.rec.values())
        right = max(max(x0, x1) for x0, _y0, x1, _y1, _w in self.rec.lines)
        self.assertGreaterEqual(right + 1e-9, rd.text_width('Supply Air, ' + 'z' * 120, True))
        divider = [x0 for x0, y0, x1, y1, _w in self.rec.lines
                   if abs(x0 - x1) < 1e-12 and 0 < x0 < right - 1e-9]
        self.assertGreaterEqual(min(divider) + 1e-9, rd.text_width(LONG_TERMINAL))


class SheetRowTests(unittest.TestCase):

    def test_sheet_rows_match_console_rows(self):
        unit = _unit()
        console, total = dpc.esp_unit_rows(unit)
        sheet, total2 = dpc.esp_unit_rows(unit, sheet=True)
        self.assertEqual(total, total2)
        self.assertEqual(len(console), len(sheet))
        self.assertEqual(total[0], 'TOTAL EXTERNAL STATIC PRESSURE')
        self.assertEqual(total[4], '0.850')
        for c, s in zip(console, sheet):
            if isinstance(s, dict) and 'cells' in s:
                self.assertTrue(s['bold'])
                self.assertEqual(s['cells'][1], '')     # detail moved to the notes
                self.assertEqual(s['cells'][4], c[4])
                self.assertTrue(c[1].startswith('worst run ends at'))
            else:
                self.assertEqual(c, s)

    def test_lesser_return_marked_not_counted(self):
        bands = [r['band'] for r in dpc.esp_unit_rows(_unit(), sheet=True)[0]
                 if isinstance(r, dict) and 'band' in r]
        self.assertEqual(len(bands), 3)
        self.assertNotIn('not counted', bands[1])
        self.assertIn('not counted', bands[2])

    def test_sheet_notes_carry_run_end_and_check_only(self):
        lines = dpc.sheet_lines(dpc.esp_unit_notes(_unit(), sheet=True))
        self.assertIn('Supply Air: worst run ends at ' + LONG_TERMINAL + '.', lines)
        self.assertTrue(any(l.startswith('CHECK: Unit max') for l in lines))
        # warnings and notes are System Corrections lines (Colin, 2026-11-02)
        self.assertFalse(any('UNPRICED' in l or 'WARNING' in l for l in lines))
        console = dpc.esp_unit_notes(_unit())
        self.assertFalse(any('worst run ends' in t for _k, t in console))

    def test_vav_rows_short_on_sheet(self):
        rep = {'basis': 'MP supply 1500 FPM', 'minimum': 0.5, 'units': [
            {'label': 'RTU-1', 'mp_terminal_text': None, 'boxes': [
                {'box': 'VAV-1', 'status': 'NOT REACHED'},
                {'box': 'VAV-2', 'status': 'NOT CHECKED', 'loss': 0.2, 'loss_sf': 0.22}]}]}
        rows = dpc.vav_rows(rep, sheet=True)
        self.assertEqual(rows[0][2], 'not reached')
        self.assertEqual(rows[1][4], 'not checked')
        self.assertEqual(dpc.vav_rows(rep)[0][2], 'not reached by the index run search')
        notes = [t for _k, t in dpc.vav_notes(rep, sheet=True)]
        self.assertTrue(any(t.startswith('Not reached') for t in notes))
        self.assertTrue(any(t.startswith('Not checked') for t in notes))
        self.assertFalse(any(t.startswith('Not reached') for _k, t in dpc.vav_notes(rep)))

    def test_summary_band_text(self):
        rec = {'grand': [('Grand total airflow', '5,250 CFM'), ('Diffusers', '14'),
                         ('Ducts', '60')]}
        self.assertEqual(dpc.summary_band_text(rec),
                         'Grand total airflow 5,250 CFM | Diffusers 14 | Ducts 60')
        self.assertEqual(dpc.summary_band_text({}), '')


class SheetLayoutTests(unittest.TestCase):
    """The section drawers from script.py against the recorder."""

    def setUp(self):
        self.rec = Recorder().install()

    def tearDown(self):
        self.rec.restore()

    def test_system_summary(self):
        rep = {'grand': [('Grand total airflow', '5,250 CFM')],
               'equipment': [['RJA - Rooftop Unit - Packaged Gas Electric 25 Ton (id 123456)',
                              '5,250', '14', '1,200 CFM Return Air']],
               'by_class': [['Supply Air', '5,250', '14'], ['Return Air', '1,200', '4']],
               'types': [['RJA - Supply Diffuser - Linear Slot (CFM Auto-Sizing)', '6', '900']],
               'warnings': ['1 diffuser(s) with Flow = 0 (missing CFM data)']}
        y = DRAW['_draw_system_summary'](None, None, 0.0, TOTAL_W, rep)
        vals = self.rec.values()
        self.assertEqual(vals[0], 'SYSTEM SUMMARY')
        self.assertIn('Grand total airflow 5,250 CFM', vals)
        self.assertIn('RJA - Supply Diffuser - Linear Slot (CFM Auto-Sizing)', vals)
        # no note panel after the section (Colin, 2026-11-02)
        self.assertFalse(any('Flow = 0' in v for v in vals))
        self.assertLess(y, 0.0)
        # by class and types side by side: same top, inside the width
        # RTU wording by default (Colin, 2026-11-09): every terminal counted
        heads = [t for t in self.rec.texts if t[2] in ('System', 'Diffuser/Grille type')]
        self.assertEqual(len(heads), 2)
        self.assertAlmostEqual(heads[0][1], heads[1][1])
        for x0, _t, x1, _b, _c in self.rec.fills:
            self.assertLessEqual(x1, TOTAL_W + 1e-9)

    def test_exhaust_scope_wording_sheet(self):
        # Colin, 2026-11-09: an Exhaust System sheet reads Exhaust CFM and
        # Diffusers/Grilles, with no other air column and no Supply text.
        rep = {'grand': dpc.grand_rows({'Exhaust Air': 700.0}, 5, 9),
               'equipment': [['EF-1 (id 1)', '700', '5', '700 CFM Exhaust Air']],
               'by_class': [['Exhaust Air', '700', '5']],
               'types': [['RJA - Exhaust Grille', '5', '700']],
               'words': dpc.scope_words('exhaust')}
        DRAW['_draw_system_summary'](None, None, 0.0, TOTAL_W, rep)
        vals = self.rec.values()
        for v in ('Exhaust CFM', 'Diffusers/Grilles', 'Diffuser/Grille type'):
            self.assertIn(v, vals)
        self.assertIn('Exhaust CFM 700 | Total Diffusers/Grilles 5 | Ducts 9', vals)
        self.assertFalse(any('Supply' in v for v in vals), vals)
        self.assertNotIn('700 CFM Exhaust Air', vals)
        for x0, _t, x1, _b, _c in self.rec.fills:
            self.assertLessEqual(x1, TOTAL_W + 1e-9)
        src = io.open(SCRIPT, encoding='utf-8').read().replace('\r\n', '\n')
        m = re.search(r'^def _legend_rows\(.*?\n    \]\n', src, re.S | re.M)
        ns = {'duct_perf_console': dpc}
        exec(m.group(0), ns)
        legend = ns['_legend_rows'](dpc.scope_words('exhaust'))
        self.assertIn('Branch: diffuser/grille and duct both correctly sized.', legend[0][1])
        self.assertIn('Branch: diffuser/grille or duct undersized.', legend[3][1])

    def test_vav_box_caution_sheet(self):
        # VAV Box scope (Colin, 2026-11-09): the total is labelled as after
        # the box, and the caution prints under the ESP table and the
        # System Summary equipment table.
        box = dict(_unit(), label='VAV-1 (id 5)', classes=[_cls('Supply Air', 0.30)],
                   checks=[], esp_total=0.33, esp_sides=['Supply Air'],
                   caution=dpc.VAVBOX_ESP_CAUTION)
        DRAW['_draw_external_static'](None, None, 0.0, TOTAL_W, {'units': [box]})
        vals = self.rec.values()
        self.assertIn(dpc.VAVBOX_ESP_LABEL, vals)
        self.assertNotIn('TOTAL EXTERNAL STATIC PRESSURE', vals)
        text = ' '.join(' '.join(vals).split())
        self.assertIn('EXTERNAL STATIC AFTER THE VAV BOX ONLY. NOT THE TOTAL SYSTEM LOSS.',
                      text)
        self.assertNotIn('WARNING: EXTERNAL STATIC AFTER', text)
        self.rec.texts = []
        rep = {'grand': [], 'equipment': [['VAV-1 (id 5)', '800', '4', '-']],
               'by_class': [], 'types': [], 'caution': dpc.VAVBOX_ESP_CAUTION}
        DRAW['_draw_system_summary'](None, None, 0.0, TOTAL_W, rep)
        text = ' '.join(' '.join(self.rec.values()).split())
        self.assertIn('NOT THE TOTAL SYSTEM LOSS', text)

    def test_external_static_one_table_per_unit(self):
        esp = {'units': [_unit()],
               'combined': [{'title': 'Duct 3 + Duct 4',
                             'text': 'EXTERNAL STATIC PRESSURE = 0.700 in. wc (Supply Air + Return Air)'}],
               'warnings': [], 'method': 'METHOD: Darcy-Weisbach',
               'not_included': 'NOT INCLUDED: coil and cabinet'}
        y = DRAW['_draw_external_static'](None, None, 0.0, TOTAL_W, esp)
        vals = self.rec.values()
        self.assertEqual(vals.count('Item'), 1)          # one table header
        self.assertIn('EXTERNAL STATIC PRESSURE - RTU-1 (id 123456)', vals)
        self.assertIn('TOTAL EXTERNAL STATIC PRESSURE', vals)
        self.assertIn('Supply Air total', vals)
        self.assertNotIn('worst run ends at ' + LONG_TERMINAL, vals)
        self.assertIn('EXTERNAL STATIC PRESSURE - Duct 3 + Duct 4', vals)
        # No METHOD / NOT INCLUDED on the sheet (Colin, 2026-11-02): the
        # basis is read in the dialog.
        self.assertNotIn('METHOD: Darcy-Weisbach', vals)
        self.assertNotIn('NOT INCLUDED: coil and cabinet', vals)
        self.assertNotIn('EXTERNAL STATIC TAKE-OFF', ' '.join(vals))
        self.assertFalse(any('UNPRICED' in v or 'WARNING' in v for v in vals))
        total = [t for t in self.rec.texts if t[2] == 'TOTAL EXTERNAL STATIC PRESSURE'][0]
        self.assertEqual(total[3], 'bold')
        self.assertLess(y, 0.0)

    def test_every_sheet_table_fits_its_text(self):
        """Colin, 2026-11-09: no '...' in any drawn table text, every column
        at least its longest text estimate, and every table spans the sheet
        width _sheet_width sets."""
        long_rows = [
            ['Duct losses: straight duct, transition and elbow (SR4-1, CR3-1), '
             'L 3.4 / Le 3.5 ft', 'x', '1', '-', '0.050'],
            ['Fan system effect: SR7-11 pos C, L 1.0 / Le 4.9 ft', 'x', '1', '-', '0.100'],
            ['none: tap 3653648 at 3.3 ft inside 5 H, flagged, not priced', 'x', '1', '-', ''],
            ['Straight duct', 'Darcy-Weisbach, in the fan system effect', '12', '-', '0.000']]
        unit = _unit()
        unit['classes'][0]['takeoff'] = long_rows
        esp = {'units': [unit], 'combined': [], 'warnings': []}
        vav = {'basis': 'MP', 'minimum': 0.5, 'units': [{'label': 'RTU-1 (id 1)',
               'mp_terminal_text': None, 'boxes': [
                   {'box': 'VAV-1-01 ' + LONG_TERMINAL, 'status': 'LOW', 'loss': 0.4,
                    'loss_sf': 0.44, 'left': 0.31}]}]}
        summ = {'grand': [('Grand total airflow', '5,250 CFM')],
                'equipment': [[LONG_TERMINAL + ' ' + LONG_TERMINAL, '5,250', '14', '-']],
                'by_class': [['Supply Air', '5,250', '14']],
                'types': [[LONG_TERMINAL, '6', '900']]}
        heads = ['#', 'Status', 'Reason']
        flagged = [['1', 'RED', 'Straight duct: elbow 123 at 1.2 ft under 0.5 Le 3.5 ft, '
                                'Required Size Straight duct 3.5 ft']]
        legend = ['Gray: ' + 'w' * 260]
        width = DRAW['_sheet_width'](0.6, summ, esp, vav, heads, flagged, legend)
        self.assertGreaterEqual(width, DRAW['_LEGEND_TEXT_X'] + rd.text_width(legend[0]))

        drawn = []
        real = rd._styled_table

        def spy(st, view, ox, top, columns, rows, **kw):
            drawn.append((ox, columns, rows, kw))
            return real(st, view, ox, top, columns, rows, **kw)

        rd._styled_table = spy
        try:
            DRAW['_draw_system_summary'](None, None, 0.0, width, summ)
            DRAW['_draw_external_static'](None, None, 0.0, width, esp)
            DRAW['_draw_vav_boxes'](None, None, 0.0, width, vav)
            cols = DRAW['_sheet_columns'](heads, [0.05, 0.1, 0.5], width, flagged)
            DRAW['_draw_system_corrections'](None, None, 0.0, width, cols, flagged,
                                             (0, 2), [])
        finally:
            rd._styled_table = real
        self.assertEqual(len(drawn), 6)
        for ox, columns, rows, kw in drawn:
            self.assertTrue(kw.get('fit'))
            needs, band = rd.column_needs([c[0] for c in columns], rows,
                                          kw.get('totals_row', False))
            for (h, w), n in zip(columns, needs):
                self.assertGreaterEqual(w + 1e-9, n, h)
            right = ox + sum(c[1] for c in columns)
            self.assertLessEqual(right, width + 1e-9)
            self.assertGreaterEqual(sum(c[1] for c in columns) + 1e-9, band)
        # the full width tables span the sheet exactly
        for ox, columns, rows, kw in drawn:
            if 'System' not in [c[0] for c in columns] and ox == 0.0 and \
                    'Diffuser/Grille type' not in [c[0] for c in columns]:
                self.assertAlmostEqual(sum(c[1] for c in columns), width)
        for value in self.rec.values():
            self.assertNotIn('...', value)
        for text in [r[0] for r in long_rows] + ['Darcy-Weisbach, in the fan system effect',
                                                 LONG_TERMINAL]:
            self.assertIn(text, self.rec.values())

    def test_vav_boxes(self):
        rep = {'basis': 'MP supply 1500 FPM, 0.10 in. wc per 100 ft', 'minimum': 0.5,
               'units': [{'label': 'RTU-1 (id 1)', 'mp_terminal_text': '2 diffuser(s) on an MP main',
                          'boxes': [{'box': 'VAV-1-01', 'status': 'LOW', 'loss': 0.4,
                                     'loss_sf': 0.44, 'left': 0.31}]}]}
        y = DRAW['_draw_vav_boxes'](None, None, 0.0, TOTAL_W, rep)
        vals = self.rec.values()
        self.assertEqual(vals[0], 'VAV BOXES')
        self.assertIn('MP supply 1500 FPM, 0.10 in. wc per 100 ft', vals)
        self.assertIn('LOW', vals)
        self.assertIn('Values in in. wc.', vals)
        # the MP diffuser line is a System Corrections line now
        self.assertFalse(any('2 diffuser' in v for v in vals))
        self.assertLess(y, 0.0)

    def _corrections_block(self, rows, entries):
        cols = [('#', 0.05), ('Status', 0.1), ('Reason', TOTAL_W - 0.15)]
        y = DRAW['_draw_system_corrections'](None, None, 0.0, TOTAL_W, cols, rows,
                                             (2,), entries)
        return y, self.rec.values()

    def test_system_corrections_block(self):
        y, vals = self._corrections_block([['1', 'RED', 'Velocity']], _corrections())
        self.assertEqual(vals[0], 'SYSTEM CORRECTIONS')
        self.assertIn('1 duct(s)', vals)
        self.assertNotIn('DUCTS NEEDING ATTENTION', vals)
        self.assertIn('RED', vals)
        a, n = vals.index('ACTION REQUIRED'), vals.index('NOTES')
        self.assertLess(vals.index('Velocity'), a)
        self.assertLess(a, n)
        joined = ' '.join(vals)
        self.assertEqual(joined.count('UNPRICED: Fitting 777'), 1)       # deduped
        self.assertIn('WARNING: Parallel units', joined)
        self.assertNotIn('RTU-1 (id 123456):', joined)   # one unit, no prefix
        self.assertIn('NOTE: oval duct 8', ' '.join(vals[n:]))
        self.assertLess(y, 0.0)

    def test_system_corrections_empty(self):
        _y, vals = self._corrections_block([], [])
        self.assertEqual(vals, ['SYSTEM CORRECTIONS', dpc.NO_CORRECTIONS])

    def test_system_corrections_without_ducts(self):
        _y, vals = self._corrections_block([], _corrections()[2:3])
        self.assertIn(dpc.NO_FLAGGED, vals)
        self.assertNotIn('ACTION REQUIRED', vals)
        self.assertIn('NOTES', vals)

    def test_vav_without_boxes_draws_nothing(self):
        y = DRAW['_draw_vav_boxes'](None, None, -1.0, TOTAL_W, {'units': []})
        self.assertEqual(y, -1.0)
        self.assertEqual(self.rec.texts, [])


if __name__ == '__main__':
    unittest.main()
