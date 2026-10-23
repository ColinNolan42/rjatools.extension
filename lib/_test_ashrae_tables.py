# -*- coding: utf-8 -*-
"""Verification tests for ashrae_fitting_tables.json (plain CPython, no Revit).

Source: 2001 ASHRAE Fundamentals Handbook (SI), Ch. 34 Duct Design, fitting
loss coefficients. Every expected value below was read off the rendered PDF
page image, not taken from the parser output.

Run:  python lib\\_test_ashrae_tables.py
"""
import io
import json
import math
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(HERE, 'ashrae_fitting_tables.json')

EXPECTED_KEYS = {
    'CD3-1': '1d', 'CD3-3': '1d', 'CD3-12': '1d', 'CD6-4': 'const',
    'CR3-1': '2d', 'CR3-1_K': '1d', 'CR3-6': '2d', 'CR3-9': 'const',
    'CR3-17': '2d', 'CR3-17_Kr': '1d',
    'SR4-1': '2d', 'ER4-1': '2d', 'SD4-1': '2d', 'ED4-1': '2d',
    'SD4-2': '2d', 'SR4-3': '2d', 'ER4-3': '2d', 'ED4-2': '2d',
    'SR5-13_Cb': '2d', 'SR5-13_Cs': '2d', 'SR5-11_Cb': '2d', 'SR5-11_Cs': '2d',
    'SD5-1_Cb': '2d', 'SD5-1_Cs': '2d', 'SD5-9_Cb': '2d', 'SD5-9_Cs': '2d',
    'SD5-10_Cb': '2d', 'SD5-10_Cs': '2d', 'SD5-24_Cb1': '3d', 'SD5-24_Cs': '2d',
    'SR5-14': '1d', 'ER5-2_Cb': '1d', 'ER5-2_Cs': '1d', 'ER5-3_Cb': '1d',
    'ER5-3_Cs': '1d', 'ER5-4': '1d',
    'ED5-3_le250_Cb': '3d', 'ED5-3_le250_Cs': '3d',
    'ED5-3_gt250_Cb': '3d', 'ED5-3_gt250_Cs': '3d',
    'ED5-2_Cb': '3d', 'ED5-2_Cs': '3d', 'SR2-1': 'const',
}

# Shapes as printed: 1d length, 2d rows x cols, 3d layers x rows x cols.
EXPECTED_SHAPES = {
    'CD3-1': (8,), 'CD3-3': (8,), 'CD3-12': (4,), 'CR3-1_K': (11,), 'CR3-17_Kr': (9,),
    'SR5-14': (2,), 'ER5-4': (2,), 'ER5-2_Cb': (10,), 'ER5-2_Cs': (9,),
    'ER5-3_Cb': (10,), 'ER5-3_Cs': (9,),
    'CR3-1': (5, 11), 'CR3-6': (6, 11), 'CR3-17': (10, 14),
    'SR4-1': (9, 10), 'ER4-1': (9, 10), 'SD4-1': (9, 10), 'ED4-1': (9, 10),
    'SD4-2': (9, 10), 'SR4-3': (9, 10), 'ER4-3': (9, 10), 'ED4-2': (9, 10),
    'SR5-13_Cb': (9, 9), 'SR5-13_Cs': (9, 9), 'SR5-11_Cb': (9, 9), 'SR5-11_Cs': (9, 9),
    'SD5-1_Cb': (9, 9), 'SD5-1_Cs': (9, 9), 'SD5-9_Cb': (9, 9), 'SD5-9_Cs': (9, 9),
    'SD5-10_Cb': (9, 9), 'SD5-10_Cs': (9, 9), 'SD5-24_Cs': (9, 9),
    'SD5-24_Cb1': (5, 9, 9),
    'ED5-3_le250_Cb': (9, 9, 9), 'ED5-3_le250_Cs': (9, 9, 9),
    'ED5-3_gt250_Cb': (9, 9, 9), 'ED5-3_gt250_Cs': (9, 9, 9),
    'ED5-2_Cb': (9, 9, 9), 'ED5-2_Cs': (9, 9, 9),
}

# Blank (null) cells as printed. Each row is described by the count of filled
# cells; LEFT means values are left aligned (blanks at the right end), RIGHT
# means blanks are leading (values pushed to the right end). Any grid not
# listed here shall contain no nulls at all.
_TRI_SPLIT = [2, 5, 7, 9, 9, 9, 9, 9, 9]          # SD5-x Cs tables
_TRI_DIAG = [1, 2, 3, 4, 5, 6, 7, 8, 9]           # SR5-11 / SR5-13 Cs tables
BLANKS_LEFT = {
    'SR5-13_Cs': _TRI_DIAG,
    'SR5-11_Cs': _TRI_DIAG,
    'SD5-1_Cb': [3, 7, 9, 9, 9, 9, 9, 9, 9],
    'SD5-1_Cs': _TRI_SPLIT,
    'SD5-9_Cs': _TRI_SPLIT,
    'SD5-10_Cb': _TRI_SPLIT,
    'SD5-10_Cs': _TRI_SPLIT,
    'SD5-24_Cs': _TRI_SPLIT,
}
# SD5-24 Cb1: leading blank count per Ab1/Ac row, per As/Ac layer.
SD5_24_LEADING = {
    0.20: [0, 1, 2, 2, 3, 4, 4, 5, 6],
    0.35: [1, 2, 3, 4, 5, 6, 7, 9, 9],
    0.55: [1, 2, 3, 4, 6, 7, 8, 9, 9],
    0.80: [0] * 9,
    1.00: [0] * 9,
}

# Spot cells read off the rendered page images.
#   1d:    (x, value)
#   2d:    (row, col, value)
#   3d:    (layer, row, col, value)
#   const: value
SPOT = {
    'CD3-1': [(75, 0.30), (150, 0.14), (250, 0.11)],
    'CD3-3': [(75, 0.18), (150, 0.08), (250, 0.07)],
    'CD3-12': [(0.75, 0.54), (1.0, 0.42), (1.5, 0.34), (2.0, 0.33)],
    'CD6-4': 0.24,
    'CR3-9': 0.11,
    'SR2-1': 1.0,
    'CR3-1': [(0.50, 0.25, 1.53), (0.50, 8.00, 1.18), (2.00, 0.25, 0.20),
              (2.00, 8.00, 0.15), (1.50, 1.00, 0.17)],
    'CR3-1_K': [(0, 0.00), (45, 0.60), (180, 1.40), (110, 1.13)],
    'CR3-6': [(20, 0.25, 0.08), (90, 0.25, 1.30), (90, 8.00, 0.83), (75, 1.00, 0.81)],
    'CR3-17': [(0.25, 0.0, 0.00), (0.25, 100.0, 2.53), (8.00, 0.4, 0.43),
               (8.00, 100.0, 1.61), (1.50, 1.8, 4.01)],
    'CR3-17_Kr': [(10, 1.40), (60, 1.09), (500, 1.00)],
    'SR4-1': [(0.10, 10, 0.05), (0.10, 60, 0.08), (1.00, 180, 1.00),
              (16.00, 10, 66.56), (16.00, 180, 250.88)],
    'ER4-1': [(0.06, 10, 0.26), (0.10, 60, 0.82), (6.00, 120, 10.20), (10.00, 180, 43.10)],
    'SD4-1': [(0.10, 10, 0.05), (2.00, 10, 0.44), (16.00, 30, 153.60), (16.00, 180, 225.28)],
    'ED4-1': [(0.06, 10, 0.21), (2.00, 45, 0.22), (6.00, 120, 10.10), (10.00, 180, 43.00)],
    'SD4-2': [(0.10, 10, 0.05), (2.00, 20, 1.00), (16.00, 15, 138.24), (16.00, 180, 238.08)],
    'SR4-3': [(0.10, 10, 0.05), (0.50, 15, 0.07), (10.00, 15, 50.00), (16.00, 180, 238.08)],
    'ER4-3': [(0.06, 10, 0.30), (6.00, 120, 10.14), (10.00, 180, 43.05)],
    'ED4-2': [(0.06, 10, 0.30), (0.06, 180, 0.93), (10.00, 90, 19.06), (10.00, 180, 43.05)],
    'SR5-13_Cb': [(0.1, 0.1, 0.73), (0.1, 0.9, 0.40), (0.5, 0.5, 0.73),
                  (0.9, 0.1, 78.98), (0.9, 0.9, 0.73)],
    'SR5-13_Cs': [(0.1, 0.1, 0.04), (0.6, 0.5, 0.10), (0.9, 0.1, 51.37), (0.9, 0.9, 0.04)],
    'SR5-11_Cb': [(0.1, 0.1, 1.58), (0.1, 0.9, 0.75), (0.9, 0.1, 72.92), (0.9, 0.9, 1.58)],
    'SR5-11_Cs': [(0.1, 0.1, 0.04), (0.5, 0.4, 0.13), (0.9, 0.1, 51.37), (0.9, 0.9, 0.04)],
    'SD5-1_Cb': [(0.1, 0.1, 0.38), (0.1, 0.3, 0.48), (0.2, 0.7, 0.45),
                 (0.9, 0.1, 73.59), (0.9, 0.9, 0.38)],
    'SD5-1_Cs': [(0.1, 0.2, 0.16), (0.4, 0.9, 0.34), (0.9, 0.1, 36.45), (0.9, 0.9, 0.13)],
    'SD5-9_Cb': [(0.1, 0.1, 1.20), (0.1, 0.9, 7.02), (0.9, 0.1, 80.60), (0.9, 0.9, 1.20)],
    'SD5-9_Cs': [(0.2, 0.5, 0.28), (0.9, 0.1, 36.45), (0.9, 0.9, 0.13)],
    'SD5-10_Cb': [(0.1, 0.1, 0.65), (0.1, 0.2, 0.24), (0.4, 0.9, 0.21),
                  (0.9, 0.1, 76.41), (0.9, 0.9, 0.65)],
    'SD5-10_Cs': [(0.1, 0.1, 0.13), (0.3, 0.7, 0.20), (0.9, 0.9, 0.13)],
    'SD5-24_Cb1': [(0.20, 0.1, 0.1, 2.07), (0.20, 0.4, 0.3, 0.90), (0.20, 0.9, 0.9, 2.07),
                   (0.35, 0.7, 0.8, 0.11), (0.55, 0.4, 0.5, 0.00),
                   (1.00, 0.9, 0.1, 80.60), (1.00, 0.9, 0.9, 1.20)],
    'SD5-24_Cs': [(0.1, 0.1, 0.13), (0.9, 0.1, 36.45), (0.9, 0.9, 0.13)],
    'SR5-14': [(0.5, 0.30), (1.0, 1.00)],
    'ER5-2_Cb': [(0.1, -12.25), (0.5, 1.27), (1.0, 1.49)],
    'ER5-2_Cs': [(0.1, 2.15), (0.2, 11.91), (0.9, 0.10)],
    'ER5-3_Cb': [(0.1, -18.00), (0.5, 0.76), (1.0, 0.91)],
    'ER5-3_Cs': [(0.1, 2.15), (0.4, 3.74), (0.9, 0.10)],
    'ER5-4': [(0.5, 0.23), (1.0, 0.28)],
    'ED5-2_Cb': [(0.2, 0.2, 0.1, -25.19), (0.7, 1.0, 0.9, -0.03), (0.8, 0.2, 0.5, 1.03),
                 (1.0, 1.0, 0.9, -0.75)],
    'ED5-2_Cs': [(0.2, 0.2, 0.1, -10.16), (0.5, 1.0, 0.2, 0.00), (1.0, 0.2, 0.1, -288.39),
                 (1.0, 1.0, 0.9, -0.15)],
    'ED5-3_le250_Cb': [(0.2, 0.2, 0.1, -24.56), (0.6, 1.0, 0.1, -140.62),
                       (0.7, 0.2, 0.9, 1.49), (1.0, 1.0, 0.9, -0.80)],
    'ED5-3_le250_Cs': [(0.2, 0.2, 0.1, 18.11), (0.5, 0.5, 0.5, 1.34),
                       (1.0, 0.2, 0.1, -6.40), (1.0, 1.0, 0.9, -0.14)],
    'ED5-3_gt250_Cb': [(0.2, 0.2, 0.1, -26.08), (0.2, 1.0, 0.1, -674.72),
                       (1.0, 1.0, 0.9, 2.30)],
    'ED5-3_gt250_Cs': [(0.2, 0.2, 0.1, 20.43), (0.9, 0.3, 0.1, -21.27),
                       (1.0, 1.0, 0.1, 237.90), (1.0, 1.0, 0.9, 0.07)],
}

EXPECTED_COEF = {
    'CR3-1': 'Cp', 'CR3-1_K': 'K', 'CR3-17': 'Cp', 'CR3-17_Kr': 'Kr', 'SD5-24_Cb1': 'Cb1',
}


def _idx(values, v):
    for i, x in enumerate(values):
        if abs(x - v) < 1e-9:
            return i
    raise KeyError('%r not in %r' % (v, values))


def _shape(t):
    c = t['C']
    if t['kind'] == '1d':
        return (len(c),)
    if t['kind'] == '2d':
        return (len(c), len(c[0]))
    return (len(c), len(c[0]), len(c[0][0]))


def _rows_of(t):
    """Yield (label, row list) for every printed grid row."""
    if t['kind'] == '2d':
        for r, row in zip(t['rows']['values'], t['C']):
            yield r, row
    else:
        for lay, block in zip(t['layers']['values'], t['C']):
            for r, row in zip(t['rows']['values'], block):
                yield (lay, r), row


class AshraeTablesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with io.open(JSON_PATH, 'rb') as f:
            raw = f.read()
        cls.raw = raw
        cls.doc = json.loads(raw.decode('utf-8'))
        cls.tables = cls.doc['tables']

    def test_no_bom(self):
        self.assertFalse(self.raw.startswith(b'\xef\xbb\xbf'), 'file shall be UTF-8 without BOM')

    def test_header(self):
        self.assertIn('2001 ASHRAE Fundamentals Handbook (SI)', self.doc['source'])
        self.assertEqual(self.doc['transcribed'], '2026-10-23')

    def test_all_keys_present_with_kind(self):
        for key, kind in EXPECTED_KEYS.items():
            self.assertIn(key, self.tables, key)
            self.assertEqual(self.tables[key]['kind'], kind, key)
        self.assertEqual(set(self.tables), set(EXPECTED_KEYS))

    def test_metadata(self):
        for key, t in self.tables.items():
            for field in ('fitting', 'title', 'page', 'pdf_page', 'coef', 'conditions', 'kind', 'C'):
                self.assertIn(field, t, '%s missing %s' % (key, field))
            self.assertTrue(key.startswith(t['fitting']), key)
            self.assertTrue(t['page'].startswith('34.'), key)
            self.assertIn(t['coef'], ('Co', 'Cp', 'Cb', 'Cs', 'Cb1', 'K', 'Kr'), key)
            if key in EXPECTED_COEF:
                self.assertEqual(t['coef'], EXPECTED_COEF[key], key)
            elif key.endswith('_Cb'):
                self.assertEqual(t['coef'], 'Cb', key)
            elif key.endswith('_Cs'):
                self.assertEqual(t['coef'], 'Cs', key)

    def test_conditions(self):
        t = self.tables
        self.assertEqual(t['SR5-13_Cb']['conditions']['L'], '0.25 Wb, 75 mm min.')
        self.assertEqual(t['CR3-9']['conditions']['r_mm'], 50)
        self.assertEqual(t['CR3-9']['conditions']['s_mm'], 40)
        self.assertEqual(t['SR5-14']['conditions']['Qb/Qc'], 0.5)
        self.assertEqual(t['ER5-4']['conditions']['Qb/Qc'], 0.5)
        self.assertEqual(t['ER5-3_Cb']['conditions']['Ab/Ac'], 0.5)
        self.assertEqual(t['ED5-3_le250_Cb']['conditions']['Dc condition'], 'Dc <= 250 mm')
        self.assertEqual(t['ED5-3_gt250_Cs']['conditions']['Dc condition'], 'Dc > 250 mm')

    def test_shapes_and_axes(self):
        for key, shp in EXPECTED_SHAPES.items():
            t = self.tables[key]
            self.assertEqual(_shape(t), shp, key)
            if t['kind'] == '1d':
                self.assertEqual(len(t['x']['values']), shp[0], key)
            if t['kind'] in ('2d', '3d'):
                self.assertEqual(len(t['rows']['values']), shp[-2], key)
                self.assertEqual(len(t['cols']['values']), shp[-1], key)
                rows = t['C'] if t['kind'] == '2d' else [r for b in t['C'] for r in b]
                for row in rows:
                    self.assertEqual(len(row), shp[-1], '%s not rectangular' % key)
            if t['kind'] == '3d':
                self.assertEqual(len(t['layers']['values']), shp[0], key)
                for b in t['C']:
                    self.assertEqual(len(b), shp[1], '%s not rectangular' % key)

    def test_axis_names(self):
        t = self.tables
        self.assertEqual(t['CD3-1']['x']['name'], 'D_mm')
        self.assertEqual(t['CD3-12']['x']['name'], 'r/D')
        self.assertEqual(t['CR3-1_K']['x']['name'], 'theta_deg')
        self.assertEqual((t['CR3-1']['rows']['name'], t['CR3-1']['cols']['name']), ('r/W', 'H/W'))
        self.assertEqual((t['CR3-6']['rows']['name'], t['CR3-6']['cols']['name']), ('theta_deg', 'H/W'))
        self.assertEqual((t['CR3-17']['rows']['name'], t['CR3-17']['cols']['name']), ('H/W', 'L/W'))
        self.assertEqual((t['SR4-1']['rows']['name'], t['SR4-1']['cols']['name']), ('Ao/A1', 'theta_deg'))
        self.assertEqual((t['SR5-13_Cb']['rows']['name'], t['SR5-13_Cb']['cols']['name']), ('Ab/Ac', 'Qb/Qc'))
        self.assertEqual((t['SR5-13_Cs']['rows']['name'], t['SR5-13_Cs']['cols']['name']), ('As/Ac', 'Qs/Qc'))
        self.assertEqual(t['ER5-3_Cb']['x']['name'], 'Qb/Qc')
        self.assertEqual(t['ER5-3_Cs']['x']['name'], 'Qs/Qc')
        self.assertEqual(t['ER5-4']['x']['name'], 'Ab/Ac')
        self.assertEqual(t['CR3-17_Kr']['x']['name'], 'Re/1000')
        for k in ('ED5-2_Cb', 'ED5-3_le250_Cs', 'ED5-3_gt250_Cb'):
            self.assertEqual((t[k]['layers']['name'], t[k]['rows']['name']), ('As/Ac', 'Ab/Ac'), k)
        self.assertEqual((t['SD5-24_Cb1']['layers']['name'], t['SD5-24_Cb1']['rows']['name'],
                          t['SD5-24_Cb1']['cols']['name']), ('As/Ac', 'Ab1/Ac', 'Qb1/Qc'))
        self.assertEqual(t['SD5-24_Cb1']['layers']['values'], [0.2, 0.35, 0.55, 0.8, 1.0])
        self.assertEqual(t['ED5-2_Cb']['rows']['values'], [0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0])

    def test_all_numbers_finite(self):
        def walk(o):
            if isinstance(o, list):
                for x in o:
                    walk(x)
            elif o is not None:
                self.assertIsInstance(o, (int, float))
                self.assertTrue(math.isfinite(o))
        for t in self.tables.values():
            walk(t['C'])

    def test_nulls_only_where_printed_blank(self):
        for key, t in self.tables.items():
            if t['kind'] == 'const':
                self.assertIsNotNone(t['C'], key)
                continue
            if t['kind'] == '1d':
                self.assertNotIn(None, t['C'], key)
                continue
            rows = list(_rows_of(t))
            if key in BLANKS_LEFT:
                counts = BLANKS_LEFT[key]
                for (label, row), n in zip(rows, counts):
                    self.assertTrue(all(v is not None for v in row[:n]), (key, label))
                    self.assertTrue(all(v is None for v in row[n:]), (key, label))
            elif key == 'SD5-24_Cb1':
                for (label, row) in rows:
                    lay, r = label
                    lead = SD5_24_LEADING[lay][_idx(t['rows']['values'], r)]
                    self.assertTrue(all(v is None for v in row[:lead]), label)
                    self.assertTrue(all(v is not None for v in row[lead:]), label)
            else:
                for label, row in rows:
                    self.assertNotIn(None, row, (key, label))

    def test_spot_cells_from_page_images(self):
        for key, cells in SPOT.items():
            t = self.tables[key]
            kind = t['kind']
            if kind == 'const':
                self.assertAlmostEqual(t['C'], cells, places=6, msg=key)
                continue
            self.assertGreaterEqual(len(cells), 2, key)
            for cell in cells:
                if kind == '1d':
                    x, v = cell
                    got = t['C'][_idx(t['x']['values'], x)]
                elif kind == '2d':
                    r, c, v = cell
                    got = t['C'][_idx(t['rows']['values'], r)][_idx(t['cols']['values'], c)]
                else:
                    lay, r, c, v = cell
                    got = t['C'][_idx(t['layers']['values'], lay)][
                        _idx(t['rows']['values'], r)][_idx(t['cols']['values'], c)]
                self.assertIsNotNone(got, (key, cell))
                self.assertAlmostEqual(got, v, places=6, msg='%s %r' % (key, cell))

    def test_every_table_has_spot_checks(self):
        self.assertEqual(set(SPOT), set(EXPECTED_KEYS))
        for key, cells in SPOT.items():
            if self.tables[key]['kind'] in ('2d', '3d'):
                self.assertGreaterEqual(len(cells), 3, key)


if __name__ == '__main__':
    unittest.main(verbosity=2)
