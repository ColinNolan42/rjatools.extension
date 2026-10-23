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
            'total_detail': 'worst run ends at ' + LONG_TERMINAL,
            'notes': ['(Includes 2 fitting(s) between the picked duct and the unit connection.)'],
            'warnings': ['UNPRICED: Fitting 777']}


def _unit():
    return {'label': 'RTU-1 (id 123456)',
            'classes': [_cls('Supply Air', 0.55), _cls('Return Air', 0.30),
                        _cls('Exhaust Air', 0.10)],
            'untraced': [], 'one_sided': None,
            'checks': [(False, 'Unit max external static 0.750 in. wc: computed 0.850 '
                               'in. wc (Supply Air + Return Air), EXCEEDS IT.')],
            'warnings': ['Parallel units on one main. This external static is unreliable.'],
            'esp_total': 0.85, 'esp_sides': ['Supply Air', 'Return Air']}


class ClipTests(unittest.TestCase):

    def test_short_text_untouched(self):
        self.assertEqual(rd._clip('0.123', 0.1), '0.123')

    def test_long_text_cut_with_ellipsis(self):
        out = rd._clip(LONG_TERMINAL, 0.2)
        self.assertTrue(out.endswith('...'))
        self.assertLessEqual(len(out), int((0.2 - 2 * rd.PAD) / rd.CHAR_W))

    def test_bold_cuts_sooner(self):
        self.assertLess(len(rd._clip('x' * 200, 0.3, True)), len(rd._clip('x' * 200, 0.3)))

    def test_tiny_cell_left_alone(self):
        self.assertEqual(rd._clip('abcdef', 0.01), 'abcdef')


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

    def test_clip_keeps_cells_inside(self):
        rd._styled_table(None, None, 0.0, 0.0, [('Item', 0.2), ('in. wc', 0.08)],
                         [[LONG_TERMINAL, '0.123']], clip=True)
        self.assertNotIn(LONG_TERMINAL, self.rec.values())
        self.assertIn('0.123', self.rec.values())


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

    def test_sheet_notes_carry_run_end_and_leads(self):
        lines = dpc.sheet_lines(dpc.esp_unit_notes(_unit(), sheet=True))
        self.assertIn('Supply Air: worst run ends at ' + LONG_TERMINAL + '.', lines)
        self.assertIn('UNPRICED: Fitting 777', lines)          # own lead kept
        self.assertTrue(any(l.startswith('CHECK FAILED: Unit max') for l in lines))
        self.assertTrue(any(l.startswith('WARNING: Parallel units') for l in lines))
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
        self.assertIn('WARNING: 1 diffuser(s) with Flow = 0 (missing CFM data)', vals)
        self.assertLess(y, 0.0)
        # by class and types side by side: same top, inside the width
        heads = [t for t in self.rec.texts if t[2] in ('System', 'Diffuser type')]
        self.assertEqual(len(heads), 2)
        self.assertAlmostEqual(heads[0][1], heads[1][1])
        for x0, _t, x1, _b, _c in self.rec.fills:
            self.assertLessEqual(x1, TOTAL_W + 1e-9)

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
        self.assertIn('METHOD: Darcy-Weisbach', vals)
        self.assertNotIn('EXTERNAL STATIC TAKE-OFF', ' '.join(vals))
        total = [t for t in self.rec.texts if t[2] == 'TOTAL EXTERNAL STATIC PRESSURE'][0]
        self.assertEqual(total[3], 'bold')
        self.assertLess(y, 0.0)

    def test_table_cells_fit_their_columns(self):
        """Every table text is clipped to its column width (sheet columns
        at the narrowest table width)."""
        esp = {'units': [_unit()], 'combined': [], 'warnings': []}
        DRAW['_draw_external_static'](None, None, 0.0, TOTAL_W, esp)
        widths = [TOTAL_W * f for f in DRAW['_SHEET_TAKEOFF_COLS']]
        xs = [0.0]
        for w in widths:
            xs.append(xs[-1] + w)
        for x, _y, value, kind, align in self.rec.texts:
            if align != 'left' or kind == 'title':
                continue
            for c in range(len(widths)):
                if abs(x - xs[c]) < 1e-9:
                    cw = rd.BOLD_CHAR_W if kind == 'bold' else rd.CHAR_W
                    # a band heading row spans the full width
                    room = (TOTAL_W if value.startswith(('Supply Air,', 'Return Air,',
                                                         'Exhaust Air,'))
                            else widths[c]) - 2 * rd.PAD
                    if c == 0 and len(value) * cw > room:
                        # note panel text starts at x 0 too, wrapped to the width
                        self.assertLessEqual(len(value) * rd.CHAR_W, TOTAL_W, value)
                    elif c > 0:
                        self.assertLessEqual(len(value) * cw, room + 1e-9, value)

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
        self.assertTrue(any(v.startswith('WARNING: RTU-1 (id 1): 2 diffuser') for v in vals))
        self.assertLess(y, 0.0)

    def test_vav_without_boxes_draws_nothing(self):
        y = DRAW['_draw_vav_boxes'](None, None, -1.0, TOTAL_W, {'units': []})
        self.assertEqual(y, -1.0)
        self.assertEqual(self.rec.texts, [])


if __name__ == '__main__':
    unittest.main()
