# -*- coding: ascii -*-
"""
_test_water_tables.py

Standalone unittest suite for water_tables.py. Uses only the stdlib
unittest module so it can be run directly:

    python _test_water_tables.py

Also runnable under CPython 3.x. water_tables.py itself must remain
IronPython 2.7 compatible (pure ASCII, no f-strings, no type hints) --
this test file is CPython-only tooling and does not need to obey those
constraints, but is kept simple/portable anyway.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import water_tables as wt


class TestSelectSizeCold(unittest.TestCase):
    """Boundary checks for select_size(wsfu, 'cold') against the firm
    table: 1/2=2.0, 3/4=6.0, 1=19.0, 1-1/4=55.0, 1-1/2=100.0, 2=250.0,
    2-1/2=475.0, 3=750.0 (max_wsfu_8fps_cold column)."""

    def test_half_inch_exact_and_below(self):
        r = wt.select_size(2.0, "cold")
        self.assertEqual(r["nominal_size"], "1/2")
        self.assertEqual(r["status"], "ok")
        self.assertEqual(r["limit_wsfu"], 2.0)

    def test_half_inch_just_below_limit(self):
        r = wt.select_size(1.99, "cold")
        self.assertEqual(r["nominal_size"], "1/2")

    def test_three_quarter_inch_just_above_half_inch_limit(self):
        r = wt.select_size(2.01, "cold")
        self.assertEqual(r["nominal_size"], "3/4")

    def test_three_quarter_inch_exact(self):
        r = wt.select_size(6.0, "cold")
        self.assertEqual(r["nominal_size"], "3/4")

    def test_one_inch_just_above_three_quarter(self):
        r = wt.select_size(6.01, "cold")
        self.assertEqual(r["nominal_size"], "1")

    def test_one_inch_exact(self):
        r = wt.select_size(19.0, "cold")
        self.assertEqual(r["nominal_size"], "1")

    def test_one_and_quarter_just_above_one(self):
        r = wt.select_size(19.01, "cold")
        self.assertEqual(r["nominal_size"], "1-1/4")

    def test_one_and_quarter_exact(self):
        r = wt.select_size(55.0, "cold")
        self.assertEqual(r["nominal_size"], "1-1/4")

    def test_one_and_half_just_above(self):
        r = wt.select_size(55.01, "cold")
        self.assertEqual(r["nominal_size"], "1-1/2")

    def test_one_and_half_exact(self):
        r = wt.select_size(100.0, "cold")
        self.assertEqual(r["nominal_size"], "1-1/2")

    def test_two_inch_just_above(self):
        r = wt.select_size(100.01, "cold")
        self.assertEqual(r["nominal_size"], "2")

    def test_two_inch_exact(self):
        r = wt.select_size(250.0, "cold")
        self.assertEqual(r["nominal_size"], "2")

    def test_two_half_just_above(self):
        r = wt.select_size(250.01, "cold")
        self.assertEqual(r["nominal_size"], "2-1/2")

    def test_two_half_exact(self):
        r = wt.select_size(475.0, "cold")
        self.assertEqual(r["nominal_size"], "2-1/2")

    def test_three_inch_just_above(self):
        r = wt.select_size(475.01, "cold")
        self.assertEqual(r["nominal_size"], "3")

    def test_three_inch_exact(self):
        r = wt.select_size(750.0, "cold")
        self.assertEqual(r["nominal_size"], "3")
        self.assertEqual(r["status"], "ok")

    def test_exceeds_table_just_above_max(self):
        r = wt.select_size(750.01, "cold")
        self.assertIsNone(r["nominal_size"])
        self.assertEqual(r["status"], "exceeds_table")
        self.assertEqual(r["max_table_wsfu"], 750.0)

    def test_case_insensitive_system(self):
        r = wt.select_size(2.0, "COLD")
        self.assertEqual(r["nominal_size"], "1/2")
        r2 = wt.select_size(2.0, "Cold")
        self.assertEqual(r2["nominal_size"], "1/2")


class TestSelectSizeHot(unittest.TestCase):
    """Boundary checks for select_size(wsfu, 'hot') against the firm
    table: 1/2=1.0, 3/4=3.0, 1=7.0, 1-1/4=18.0, 1-1/2=40.0, 2=120.0,
    2-1/2=250.0, 3=400.0 (max_wsfu_5fps_hot column)."""

    def test_half_inch_exact(self):
        r = wt.select_size(1.0, "hot")
        self.assertEqual(r["nominal_size"], "1/2")

    def test_three_quarter_exact(self):
        r = wt.select_size(3.0, "hot")
        self.assertEqual(r["nominal_size"], "3/4")

    def test_three_inch_exact(self):
        r = wt.select_size(400.0, "hot")
        self.assertEqual(r["nominal_size"], "3")
        self.assertEqual(r["status"], "ok")

    def test_exceeds_table_just_above_max(self):
        r = wt.select_size(400.01, "hot")
        self.assertIsNone(r["nominal_size"])
        self.assertEqual(r["status"], "exceeds_table")
        self.assertEqual(r["max_table_wsfu"], 400.0)

    def test_case_insensitive_system(self):
        r = wt.select_size(1.0, "HOT")
        self.assertEqual(r["nominal_size"], "1/2")


class TestSelectSizeSpreadsheetGap(unittest.TestCase):
    """The source workbook's IF() formulas test '<=X' then '>X.01', which
    leaves a gap for a value like X.005. select_size must NOT reproduce
    that defect -- every non-negative wsfu must map to a real size."""

    def test_cold_gap_at_1_005(self):
        r = wt.select_size(1.005, "cold")
        self.assertIsNotNone(r["nominal_size"])
        self.assertEqual(r["nominal_size"], "1/2")
        self.assertEqual(r["status"], "ok")

    def test_cold_gap_at_2_005(self):
        r = wt.select_size(2.005, "cold")
        self.assertIsNotNone(r["nominal_size"])
        self.assertEqual(r["nominal_size"], "3/4")
        self.assertEqual(r["status"], "ok")

    def test_hot_gap_at_1_005(self):
        r = wt.select_size(1.005, "hot")
        self.assertIsNotNone(r["nominal_size"])
        self.assertEqual(r["status"], "ok")


class TestSelectSizeEdgesAndErrors(unittest.TestCase):

    def test_zero_wsfu_is_no_demand(self):
        r = wt.select_size(0, "cold")
        self.assertEqual(r["status"], "no_demand")
        self.assertIsNone(r["nominal_size"])
        self.assertIsNone(r["size_inches"])

    def test_zero_wsfu_hot_is_no_demand(self):
        r = wt.select_size(0.0, "hot")
        self.assertEqual(r["status"], "no_demand")

    def test_negative_wsfu_raises(self):
        with self.assertRaises(ValueError):
            wt.select_size(-1, "cold")

    def test_bad_system_raises(self):
        with self.assertRaises(ValueError):
            wt.select_size(5, "lukewarm")

    def test_none_wsfu_raises(self):
        with self.assertRaises(ValueError):
            wt.select_size(None, "cold")

    def test_non_numeric_wsfu_raises(self):
        with self.assertRaises(ValueError):
            wt.select_size("abc", "cold")


class TestWsfuToGpm(unittest.TestCase):

    def test_flush_tank_exact_rows(self):
        self.assertEqual(wt.wsfu_to_gpm(10, "flush_tank")["gpm"], 14.6)
        self.assertEqual(wt.wsfu_to_gpm(100, "flush_tank")["gpm"], 43.5)
        self.assertEqual(wt.wsfu_to_gpm(5000, "flush_tank")["gpm"], 593.0)

    def test_flushometer_exact_row(self):
        r = wt.wsfu_to_gpm(10, "flushometer_valve")
        self.assertEqual(r["gpm"], 27.0)

    def test_flush_tank_interpolated_midpoint(self):
        # Hand-computed: rows wsfu=20 gpm=19.6, wsfu=25 gpm=21.5.
        # At wsfu=22.5 (40% of the way from 20 to 25):
        #   gpm = 19.6 + 0.5 * (21.5 - 19.6) = 19.6 + 0.95 = 20.55
        r = wt.wsfu_to_gpm(22.5, "flush_tank")
        self.assertEqual(r["status"], "ok")
        self.assertAlmostEqual(r["gpm"], 20.55, places=6)

    def test_flush_tank_below_table_clamps(self):
        r = wt.wsfu_to_gpm(0.5, "flush_tank")
        self.assertEqual(r["status"], "below_table")
        self.assertEqual(r["gpm"], 3.0)  # first row (wsfu=1.0) gpm

    def test_flushometer_below_table_clamps(self):
        r = wt.wsfu_to_gpm(2, "flushometer_valve")
        self.assertEqual(r["status"], "below_table")
        self.assertEqual(r["gpm"], 15.0)  # first row (wsfu=5.0) gpm

    def test_above_table_exceeds(self):
        r = wt.wsfu_to_gpm(6000, "flush_tank")
        self.assertEqual(r["status"], "exceeds_table")
        self.assertIsNone(r["gpm"])

    def test_above_table_exceeds_flushometer(self):
        r = wt.wsfu_to_gpm(6000, "flushometer_valve")
        self.assertEqual(r["status"], "exceeds_table")
        self.assertIsNone(r["gpm"])

    def test_bad_curve_raises(self):
        with self.assertRaises(ValueError):
            wt.wsfu_to_gpm(10, "not_a_curve")

    def test_negative_wsfu_raises(self):
        with self.assertRaises(ValueError):
            wt.wsfu_to_gpm(-5, "flush_tank")

    def test_curve_case_insensitive(self):
        r = wt.wsfu_to_gpm(10, "FLUSH_TANK")
        self.assertEqual(r["gpm"], 14.6)


class TestFixtureLookups(unittest.TestCase):

    def test_min_fixture_supply_lavatory(self):
        self.assertEqual(wt.min_fixture_supply("Lavatory"), "3/8")

    def test_min_fixture_supply_case_insensitive(self):
        self.assertEqual(wt.min_fixture_supply("lavatory"), "3/8")
        self.assertEqual(wt.min_fixture_supply("LAVATORY"), "3/8")

    def test_min_fixture_supply_unknown_returns_none(self):
        self.assertIsNone(wt.min_fixture_supply("Not A Real Fixture"))

    def test_min_fixture_supply_none_input(self):
        self.assertIsNone(wt.min_fixture_supply(None))

    def test_fixture_flow_lavatory_public(self):
        row = wt.fixture_flow("Lavatory, public")
        self.assertIsNotNone(row)
        self.assertEqual(row["flow_gpm"], 0.4)
        self.assertEqual(row["flow_pressure_psi"], 8.0)

    def test_fixture_flow_unknown_returns_none(self):
        self.assertIsNone(wt.fixture_flow("Not A Real Fixture"))

    def test_fixture_flow_returns_copy(self):
        row = wt.fixture_flow("Lavatory, public")
        row["flow_gpm"] = 999.0
        row2 = wt.fixture_flow("Lavatory, public")
        self.assertEqual(row2["flow_gpm"], 0.4)


class TestRowAccessors(unittest.TestCase):

    def test_firm_sizing_rows_count_and_copy(self):
        rows = wt.firm_sizing_rows()
        self.assertEqual(len(rows), 8)
        rows[0]["nominal_size"] = "MUTATED"
        rows2 = wt.firm_sizing_rows()
        self.assertEqual(rows2[0]["nominal_size"], "1/2")

    def test_wsfu_rows_count_and_copy(self):
        rows = wt.wsfu_rows()
        self.assertEqual(len(rows), 27)
        rows[0]["fixture"] = "MUTATED"
        rows2 = wt.wsfu_rows()
        self.assertNotEqual(rows2[0]["fixture"], "MUTATED")


class TestBasisOfDesignLines(unittest.TestCase):

    def test_returns_nonempty_ascii_string_list(self):
        lines = wt.basis_of_design_lines()
        self.assertIsInstance(lines, list)
        self.assertGreater(len(lines), 0)
        for line in lines:
            self.assertIsInstance(line, str)
            line.encode("ascii")  # raises UnicodeEncodeError if not pure ASCII

    def test_extra_lines_appended(self):
        lines = wt.basis_of_design_lines(extra=["Extra note one.", "Extra note two."])
        self.assertIn("Extra note one.", lines)
        self.assertIn("Extra note two.", lines)
        self.assertEqual(lines[-2], "Extra note one.")
        self.assertEqual(lines[-1], "Extra note two.")

    def test_mentions_code_edition_and_key_values(self):
        lines = wt.basis_of_design_lines()
        joined = " ".join(lines)
        self.assertIn("2024 International Plumbing Code", joined)
        self.assertIn("E103.3(2)", joined)
        self.assertIn("E103.3(3)", joined)
        self.assertIn("8 fps", joined)
        self.assertIn("5 fps", joined)


class TestDataIntegrity(unittest.TestCase):

    def test_firm_table_ascending_both_limit_columns(self):
        rows = wt.firm_sizing_rows()
        cold_limits = [row["max_wsfu_8fps_cold"] for row in rows]
        hot_limits = [row["max_wsfu_5fps_hot"] for row in rows]
        self.assertEqual(cold_limits, sorted(cold_limits))
        self.assertEqual(hot_limits, sorted(hot_limits))
        for i in range(1, len(cold_limits)):
            self.assertGreater(cold_limits[i], cold_limits[i - 1])
        for i in range(1, len(hot_limits)):
            self.assertGreater(hot_limits[i], hot_limits[i - 1])

    def test_demand_flush_tank_strictly_increasing_wsfu(self):
        data = wt._load()["demand_flush_tank"]
        wsfu_values = [row["wsfu"] for row in data]
        for i in range(1, len(wsfu_values)):
            self.assertGreater(wsfu_values[i], wsfu_values[i - 1])

    def test_demand_flushometer_strictly_increasing_wsfu(self):
        data = wt._load()["demand_flushometer_valve"]
        wsfu_values = [row["wsfu"] for row in data]
        for i in range(1, len(wsfu_values)):
            self.assertGreater(wsfu_values[i], wsfu_values[i - 1])


def _run():
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromModule(sys.modules[__name__])
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    total = result.testsRun
    failures = len(result.failures)
    errors = len(result.errors)
    passed = total - failures - errors

    print("")
    print("=" * 70)
    print("SUMMARY: {0} run, {1} passed, {2} failed, {3} errors".format(
        total, passed, failures, errors))
    print("=" * 70)

    if failures or errors:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(_run())
