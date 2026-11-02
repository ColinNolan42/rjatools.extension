# -*- coding: utf-8 -*-
"""CPython tests for fan_system_effect.py (Colin, 2026-10-29). No Revit.

Expected values are hand calculations from the 2001 ASHRAE Fundamentals (SI)
equations and table cells read off the rendered pages: Eq. 36 (p. 34.12),
Eq. 37 (p. 34.13), SR7-5 to SR7-12 (p. 34.66 to 34.68), ED7-2 (p. 34.48),
ER7-1 (p. 34.60). The Duct Performance dialog's XAML and _settings_result
are pulled out of script.py as text, the way _test_critical_path does.

Run:  python lib\\_test_fan_system_effect.py
"""
import io
import os
import re
import sys
import unittest
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import fan_system_effect as F  # noqa: E402
import vav_checks              # noqa: E402

SCRIPT = os.path.join(os.path.dirname(HERE), 'Developer Tools.tab', 'Performance.panel',
                      'Duct Performance.pushbutton', 'script.py')

M_PER_FT = 0.3048


class EffectiveLengthTest(unittest.TestCase):
    def test_eq37_at_or_below_13_mps(self):
        # 24 x 24 in. = 4 ft2 = 371,612 mm2, sqrt 609.6 mm; Le = 609.6 / 350 m
        le, eq = F.effective_duct_length_ft(4.0, 1000.0)
        self.assertEqual(eq, 'Eq. 37')
        self.assertAlmostEqual(le, 609.6 / 350.0 / M_PER_FT, places=6)
        self.assertAlmostEqual(le, 5.714, places=3)

    def test_eq36_above_13_mps(self):
        # 3000 fpm = 15.24 m/s; Le = 15.24 x 609.6 / 4500 m
        le, eq = F.effective_duct_length_ft(4.0, 3000.0)
        self.assertEqual(eq, 'Eq. 36')
        self.assertAlmostEqual(le, 15.24 * 609.6 / 4500.0 / M_PER_FT, places=6)

    def test_split_is_13_mps(self):
        at = 13.0 / 0.00508                     # 2559 fpm
        self.assertEqual(F.effective_duct_length_ft(4.0, at)[1], 'Eq. 37')
        self.assertEqual(F.effective_duct_length_ft(4.0, at + 1.0)[1], 'Eq. 36')


class OutletTest(unittest.TestCase):
    def test_dwdi_worst_position_at_zero_length(self):
        # Ab/Ao 0.5, L/Le 0: A 2.20, B 3.60, C 3.80, D 3.20 -> C governs
        c, pos, key, notes = F.outlet_c('DWDI', 0.5, 0.0)
        self.assertAlmostEqual(c, 3.80)
        self.assertEqual((pos, key), ('C', 'SR7-11'))
        self.assertEqual(notes, [])

    def test_swsi_tables(self):
        # SWSI Ab/Ao 0.4, L/Le 0: A 3.20, B 3.80, C 5.50, D 5.50 -> C (first)
        c, pos, key, _n = F.outlet_c('SWSI', 0.4, 0.0)
        self.assertAlmostEqual(c, 5.50)
        self.assertEqual(key, 'SR7-7')

    def test_interpolates_between_columns(self):
        # DWDI Ab/Ao 0.5 at L/Le 0.06, half way 0.00 to 0.12:
        # A 2.00, B 3.25, C 3.50, D 2.95 -> 3.50
        c, pos, _k, _n = F.outlet_c('DWDI', 0.5, 0.06)
        self.assertAlmostEqual(c, 3.50)
        self.assertEqual(pos, 'C')

    def test_zero_from_le_on_without_clamp_note(self):
        for x in (1.0, 3.0, 25.0):
            c, _p, _k, notes = F.outlet_c('DWDI', 0.5, x)
            self.assertEqual(c, 0.0, x)
            self.assertEqual(notes, [], x)

    def test_unknown_fan_type_takes_default_and_says_so(self):
        c, _p, key, notes = F.outlet_c('AXIAL', 0.5, 0.0)
        self.assertEqual(key, 'SR7-11')
        self.assertTrue(any('unknown' in n for n in notes))

    def test_outlet_effect_reasonable_loss_not_high(self):
        # Colin, 2026-11-02: judged on the loss, not C. C 3.80 at 1000 fpm is
        # 3.80 x 0.0623 = 0.237 in. wc, under 0.5: priced, not GRAY.
        r = F.outlet_effect('DWDI', 0.5, 4.0, 4000.0, 0.0)
        self.assertAlmostEqual(r['v_fpm'], 1000.0)
        self.assertAlmostEqual(r['c'], 3.80)
        self.assertAlmostEqual(r['loss'], 3.80 * (1000.0 / 4005.0) ** 2)
        self.assertFalse(r['high'])
        self.assertIsNone(r['fix_ft'])
        self.assertEqual(F.basis_text(r), 'SR7-11 pos C, L 0.0 / Le 5.7 ft')

    def test_outlet_effect_high_and_fix_length(self):
        # 2000 fpm: 3.80 x 0.2494 = 0.948 in. wc, GRAY. The fix is where the
        # worst of A to D drops under 0.5 / 0.2494 = 2.00.
        r = F.outlet_effect('DWDI', 0.5, 4.0, 8000.0, 0.0)
        self.assertTrue(r['high'])
        fx = r['fix_ft']
        self.assertIsNotNone(fx)
        self.assertLess(F.outlet_effect('DWDI', 0.5, 4.0, 8000.0, fx)['loss'], 0.5)
        self.assertGreaterEqual(F.outlet_effect('DWDI', 0.5, 4.0, 8000.0, fx - 0.1)['loss'], 0.5)
        txt = F.high_text(r, 'supply outlet', 921)
        self.assertTrue(txt.startswith('GRAY: System Effect Pressure Loss too high'))
        # names the elbow and its distance from the unit (Colin, 2026-11-02)
        self.assertIn('at elbow 921, 0.0 ft from the supply outlet, 0.948 in. wc', txt)
        self.assertIn('increase length of duct to %.1f ft' % fx, txt)
        self.assertIn('SR7-11', txt)
        self.assertIn('0.5 in. wc or more', txt)

    def test_pv_of_is_used(self):
        r = F.outlet_effect('DWDI', 0.5, 4.0, 4000.0, 0.0, pv_of=lambda v: 0.2)
        self.assertAlmostEqual(r['loss'], 0.76)
        self.assertTrue(r['high'])

    def test_outlet_effect_past_le_is_zero(self):
        r = F.outlet_effect('DWDI', 0.5, 4.0, 4000.0, 6.0)
        self.assertEqual(r['c'], 0.0)
        self.assertFalse(r['high'])
        self.assertIsNone(r['fix_ft'])

    def test_reasonable_c_not_high(self):
        # Ab/Ao 1.0, L/Le 0.5: A 0.18, B 0.23, C 0.26, D 0.22 -> 0.26
        le, _eq = F.effective_duct_length_ft(4.0, 1000.0)
        r = F.outlet_effect('DWDI', 1.0, 4.0, 4000.0, 0.5 * le)
        self.assertAlmostEqual(r['c'], 0.26)
        self.assertFalse(r['high'])


class InletTest(unittest.TestCase):
    def test_ed7_2_cells(self):
        self.assertAlmostEqual(F.inlet_c(True, 1.5, 0.0)[0], 1.10)
        self.assertAlmostEqual(F.inlet_c(True, 1.0, 2.0)[0], 0.67)
        self.assertAlmostEqual(F.inlet_c(True, 3.0, 10.0)[0], 0.22)

    def test_er7_1_cells(self):
        self.assertAlmostEqual(F.inlet_c(False, 1.0, 2.0)[0], 0.67)
        self.assertAlmostEqual(F.inlet_c(False, 0.5, 0.0)[0], 2.50)
        self.assertEqual(F.inlet_c(False, 1.0, 2.0)[1], 'ER7-1')

    def test_unreadable_radius_takes_smallest_row(self):
        c, key, r_used, notes = F.inlet_c(True, None, 0.0)
        self.assertAlmostEqual(c, 1.80)
        self.assertEqual((key, r_used), ('ED7-2', 0.5))
        self.assertTrue(any('conservative' in n for n in notes))

    def test_long_run_reads_the_floor_without_clamp_note(self):
        c, _k, _r, notes = F.inlet_c(True, 1.5, 40.0)
        self.assertAlmostEqual(c, 0.33)
        self.assertEqual(notes, [])

    def test_mitered_radius_clamps_with_note(self):
        _c, _k, _r, notes = F.inlet_c(False, 0.1, 0.0)
        self.assertTrue(any('lowest value' in n and 'was used' in n for n in notes))

    def test_floor_ratio(self):
        self.assertEqual(F.floor_ratio('ED7-2', 1.0), 5.0)
        self.assertEqual(F.floor_ratio('ED7-2', 0.6), 5.0)
        self.assertEqual(F.floor_ratio('ER7-1', 2.0), 5.0)
        self.assertEqual(F.zero_ratio('SR7-10'), 1.0)

    def test_inlet_effect_floor_length_and_reasonable_loss(self):
        # 24 in. round, r/D 1.0, elbow on the inlet: C 1.20, floor 0.33 at 5 D
        # = 10 ft. At 1000 fpm 1.20 x 0.0623 = 0.075 in. wc: not GRAY.
        r = F.inlet_effect(True, 24.0, 24.0, 3.1416, 3141.6, 0.0)
        self.assertAlmostEqual(r['c'], 1.20)
        self.assertFalse(r['high'])
        self.assertAlmostEqual(r['floor_ft'], 10.0)
        self.assertAlmostEqual(r['floor_c'], 0.33)
        self.assertAlmostEqual(r['v_fpm'], 1000.0)
        self.assertEqual(F.basis_text(r), 'ED7-2 r/D 1.00, L 0.0 ft, D 24 in')

    def test_inlet_effect_high_and_fix_length(self):
        # 3000 fpm: Pv 0.561, 1.20 x 0.561 = 0.673 in. wc, GRAY. Under 0.5
        # once C < 0.891: L/D 2 x 0.309 / 0.53 = 1.166, 2.33 ft, up to 2.4 ft.
        r = F.inlet_effect(True, 24.0, 24.0, 3.1416, 9424.8, 0.0)
        self.assertTrue(r['high'])
        self.assertAlmostEqual(r['fix_ft'], 2.4)
        txt = F.high_text(r, 'return inlet', 1002)
        self.assertIn('elbow 1002, 0.0 ft from the return inlet', txt)
        self.assertIn('increase length of duct to 2.4 ft', txt)
        self.assertIn('ED7-2', txt)

    def test_inlet_effect_at_two_diameters(self):
        r = F.inlet_effect(False, 24.0, 36.0, 4.0, 4000.0, 4.0)   # r/H 1.5, L/H 2.0
        self.assertAlmostEqual(r['c'], 0.57)
        self.assertFalse(r['high'])


class SettingsTest(unittest.TestCase):
    def test_defaults(self):
        d = F.default_settings()
        self.assertEqual(d, {'on': True, 'fan_type': 'DWDI', 'blast_ratio': 0.5})
        self.assertEqual(F.HIGH_LOSS_INWC, 0.5)
        self.assertFalse(hasattr(F, 'HIGH_C'))
        self.assertFalse(hasattr(F, 'RED_C'))
        self.assertFalse(hasattr(F, 'WINDOW_FT'))    # fixed window dropped

    def test_inlet_window_is_the_floor_distance(self):
        # every ED7-2 / ER7-1 row floors at the 5.0 column: 5 D, 5 H
        self.assertAlmostEqual(F.inlet_window_ft(True, 24.0), 10.0)
        self.assertAlmostEqual(F.inlet_window_ft(False, 18.0), 7.5)

    def test_inlet_effect_within_flag(self):
        self.assertTrue(F.inlet_effect(True, 24.0, 24.0, 3.1416, 3141.6, 9.9)['within'])
        self.assertFalse(F.inlet_effect(True, 24.0, 24.0, 3.1416, 3141.6, 10.0)['within'])

    def test_connection_role_hook(self):
        # today: by classification
        self.assertEqual(F.connection_role('Supply Air'), F.OUTLET)
        self.assertEqual(F.connection_role('Return Air'), F.INLET)
        self.assertEqual(F.connection_role('Exhaust Air'), F.INLET)
        self.assertIsNone(F.connection_role('Outside Air'))
        self.assertIsNone(F.connection_role(None))
        # the exhaust fan hook: a known direction wins
        self.assertEqual(F.connection_role('Exhaust Air', 'out'), F.OUTLET)
        self.assertEqual(F.connection_role('Exhaust Air', 'in'), F.INLET)
        self.assertEqual(F.where_text('Exhaust Air', F.OUTLET), 'exhaust outlet')
        self.assertEqual(F.where_text('Supply Air', F.OUTLET), 'supply outlet')

    def test_method_lines_and_dwdi_assumption(self):
        m, a = F.method_lines({'on': True, 'fan_type': 'DWDI', 'blast_ratio': 0.5})
        self.assertIn('SR7-9 to SR7-12', m)
        self.assertIn('not priced again', m)
        self.assertEqual(len(a), 1)
        self.assertTrue(a[0].startswith('ASSUMPTION'))
        self.assertIn('SWSI', a[0])
        self.assertIn('DWDI', a[0])
        m2, a2 = F.method_lines({'on': True, 'fan_type': 'SWSI', 'blast_ratio': 0.7})
        self.assertIn('SR7-5 to SR7-8', m2)
        self.assertIn('0.70', m2)
        self.assertEqual(a2, [])
        self.assertNotIn('adapter', m)
        self.assertIn('their length counted', m)
        self.assertIn('r/H 1.0', m)
        m3, a3 = F.method_lines({'on': False})
        self.assertIn('not included', m3)
        self.assertEqual(a3, [])

    def test_no_adapter_rule(self):
        # Colin, 2026-11-02: every transition in the window is passed, its
        # length counted, so the 1.0 ft adapter rule is gone.
        for name in ('is_adapter', 'adapter_text', 'ADAPTER_TRANSITION_FT'):
            self.assertFalse(hasattr(F, name), name)
        self.assertEqual(F.MITERED_R_OVER_H, 1.0)
        # the r/H 1.0 row of ER7-1 at L 0 is below the unreadable row (r/H 0.5)
        self.assertLess(F.inlet_c(False, F.MITERED_R_OVER_H, 0.0)[0],
                        F.inlet_c(False, None, 0.0)[0])

    def test_blast_ratio_validation(self):
        self.assertEqual(F.parse_blast_ratio('0.5'), (0.5, None))
        self.assertEqual(F.parse_blast_ratio('0.4'), (0.4, None))
        self.assertEqual(F.parse_blast_ratio('1.0'), (1.0, None))
        for bad in ('0.39', '1.01', 'abc', '', 'nan'):
            v, err = F.parse_blast_ratio(bad)
            self.assertIsNone(v, bad)
            self.assertTrue(err, bad)

    def test_high_threshold(self):
        self.assertTrue(F.is_high(0.5))
        self.assertFalse(F.is_high(0.49))
        self.assertFalse(F.is_high(None))
        self.assertAlmostEqual(F.c_limit(0.25), 2.0)

    def test_fix_length_none_when_floor_is_high(self):
        # nothing in ED7-2 / ER7-1 floors at 1.0 or more, so a fix always exists
        # (at a C limit of 1.0)
        self.assertIsNotNone(F.inlet_fix_ft(False, 0.5, 24.0, 1.0))       # ER7-1 worst row
        self.assertAlmostEqual(F.inlet_fix_ft(True, 1.5, 24.0, 1.0), 0.8)   # 1.10 -> 0.60: L/D 0.4 x 2 ft
        self.assertAlmostEqual(F.outlet_fix_ft('DWDI', 1.0, 5.0, 1.0), 0.1)   # C 1.00 at L 0
        # ED7-2 floors at 0.33 at r/D 1.0: a limit under it has no fix
        self.assertIsNone(F.inlet_fix_ft(True, 1.0, 24.0, 0.3))
        self.assertTrue(F.ER7_1_SQUARE_ASSUMPTION.startswith('ASSUMPTION: ER7-1'))

    def test_text_builders(self):
        t = F.disruption_text('tap', 982, 'supply outlet', 1.0, 5.7, 'Le')
        self.assertTrue(t.startswith('WARNING: tap 982, 1.0 ft from the supply outlet'))
        self.assertIn('not priced', t)
        self.assertIn('inside Le (5.7 ft)', t)


class ReportTest(unittest.TestCase):
    """The unit's fan system effect notes are System Corrections lines
    (Colin, 2026-11-02): said once per unit, a GRAY one under Action required
    as a gray callout in the console and as its own 'GRAY:' line on the
    sheet, a NOTE under Notes; none under the external static table."""

    def test_gray_once_per_unit_console_and_sheet(self):
        import console_report as cr
        import duct_perf_console as dpc
        gray = F.high_text(F.outlet_effect('DWDI', 0.5, 4.0, 8000.0, 0.0), 'supply outlet',
                           921)
        unit = {'label': 'RTU-1', 'classes': [
            {'sys_class': 'Supply Air', 'rows': [], 'subtotal': 0.1, 'safety': 0.01,
             'total': 0.11},
            {'sys_class': 'Return Air', 'rows': [], 'subtotal': 0.1, 'safety': 0.01,
             'total': 0.11}]}
        fse = [('gray', gray), ('note', 'NOTE: x')]
        # the walker hands every class leaf the same list: gathered twice
        entries = [dpc.correction(k, t, 'RTU-1') for k, t in fse + fse]
        groups = dpc.correction_lines(entries)
        self.assertEqual(groups[dpc.ACTION], [('gray', 'RTU-1: ' + gray)])
        self.assertEqual(groups[dpc.NOTES], [('note', 'RTU-1: NOTE: x')])
        sheet = dpc.corrections_sheet(entries)
        self.assertEqual(sheet[0], ('ACTION REQUIRED', ['RTU-1: ' + gray]))  # no 'WARNING:'
        html = dpc.system_corrections(['#'], [], entries)
        self.assertEqual(html.count('System Effect Pressure Loss too high'), 1)
        self.assertIn(cr.CALLOUT_STYLES['gray'][1], html)
        self.assertIn(cr.STATUS_COLORS['GRAY'], html)
        esp = dpc.external_static({'units': [unit]})
        self.assertNotIn('System Effect Pressure Loss', esp)
        self.assertEqual(dpc.esp_unit_notes(unit), [])


class DialogTest(unittest.TestCase):
    """The themed dialog's XAML parses and carries the new controls, and the
    legacy dialog (no fan fields) still yields the defaults."""

    @classmethod
    def setUpClass(cls):
        with io.open(SCRIPT, encoding='utf-8') as fh:
            cls.src = fh.read()

    def test_xaml_parses_with_fan_controls(self):
        m = re.search(r"^_SETTINGS_XAML = '''(.*?)'''", self.src, re.S | re.M)
        self.assertIsNotNone(m)
        xaml = m.group(1).replace('@@RESOURCES@@', '').replace('@@HEADER@@', '')
        root = ET.fromstring(xaml)
        xn = '{http://schemas.microsoft.com/winfx/2006/xaml}Name'
        names = set(e.get(xn) for e in root.iter() if e.get(xn))
        for n in ('cb_fse', 'pnl_fse', 'rb_fan_dwdi', 'rb_fan_swsi', 'tb_blast',
                  'tb_fse_hint', 'pnl_static', 'cb_static'):
            self.assertIn(n, names)
        # the fan fields sit inside the external static group, greyed with it
        pnl = [e for e in root.iter() if e.get(xn) == 'pnl_static'][0]
        inner = set(e.get(xn) for e in pnl.iter() if e.get(xn))
        self.assertTrue({'cb_fse', 'pnl_fse', 'tb_blast'} <= inner)

    def test_return_filter_starts_checked(self):
        # Colin, 2026-11-02: Include return filter is on by default, in the
        # themed dialog and the legacy fallback alike.
        self.assertEqual(len(re.findall(r'^\s*cb_filter\.IsChecked = True', self.src,
                                        re.M)), 2)
        self.assertNotIn('cb_filter.IsChecked = False', self.src)

    def test_settings_result_defaults_for_legacy_dialog(self):
        m = re.search(r"^def _settings_result\(.*?\n    }\n", self.src, re.S | re.M)
        self.assertIsNotNone(m)
        ns = {'vav_checks': vav_checks, 'fan_system_effect': F}
        exec(m.group(0), ns)
        r = ns['_settings_result']({}, 10, set(), False, True, 10, {}, {}, {}, 0.0,
                                   False, vav_checks.default_settings())
        self.assertEqual(r['fan_effect'], F.default_settings())
        r2 = ns['_settings_result']({}, 10, set(), False, True, 10, {}, {}, {}, 0.0,
                                    False, vav_checks.default_settings(),
                                    {'on': False, 'fan_type': 'SWSI', 'blast_ratio': 0.7})
        self.assertEqual(r2['fan_effect']['fan_type'], 'SWSI')
        self.assertFalse(r2['fan_effect']['on'])


if __name__ == '__main__':
    unittest.main(verbosity=2)
