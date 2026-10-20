# -*- coding: ascii -*-
"""Tests for duct_sizing: smallest even size within the firm limits.

Uses the REAL friction function: the friction block of hvac_graph.py is
executed on its own (hvac_graph itself imports the Revit API and cannot load
in CPython), so these tests size with exactly the numbers Revit will.

Run:  python _test_duct_sizing.py
"""

import io
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import duct_sizing as ds  # noqa: E402


def _load_friction():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("STANDARD_AIR_DENSITY_LB_FT3 = ")
    end = src.index("def fitting_family_name(")
    ns = {"math": math}
    exec(src[start:end], ns)
    return ns["duct_friction_loss_per_100ft"]


FRICTION = _load_friction()
LIMITS = {"Supply Air": (800.0, 0.08), "Return Air": (600.0, 0.05)}


def passes(cfm, shape, size, sys="Supply Air"):
    fpm_max, fric_max = LIMITS[sys]
    if shape == ds.SHAPE_ROUND:
        fpm, fric = ds.performance(cfm, ds.round_area_ft2(size), size, FRICTION)
    else:
        w, h = size
        fpm, fric = ds.performance(cfm, ds.rect_area_ft2(w, h),
                                   ds.rect_dh_in(w, h), FRICTION)
    return fpm <= fpm_max + 1e-9 and fric <= fric_max + 1e-12


class TestSnap(unittest.TestCase):

    def test_even_only(self):
        self.assertEqual(ds.snap_even(17), 18)
        self.assertEqual(ds.snap_even(18), 18)
        self.assertEqual(ds.snap_even(10.000000000000002), 10)
        self.assertEqual(ds.snap_even(8.5), 10)


class TestRound(unittest.TestCase):

    def test_smallest_that_passes(self):
        for cfm in (100, 350, 800, 1500, 3000, 5000):
            d = ds.size_round(cfm, 800, 0.08, FRICTION)
            self.assertTrue(passes(cfm, ds.SHAPE_ROUND, d), cfm)
            i = ds.ROUND_SIZES.index(d)
            if i > 0:
                self.assertFalse(passes(cfm, ds.SHAPE_ROUND,
                                        ds.ROUND_SIZES[i - 1]), cfm)

    def test_minimum_six(self):
        self.assertEqual(ds.size_round(20, 800, 0.08, FRICTION), 6)

    def test_too_big(self):
        self.assertIsNone(ds.size_round(60000, 800, 0.08, FRICTION))

    def test_past_36(self):
        # 6,000 cfm at 800 fpm needs more than 36 in round.
        d = ds.size_round(6000, 800, 0.08, FRICTION)
        self.assertGreater(d, 36)
        self.assertEqual(d % 2, 0)

    def test_large_rect_main(self):
        # 16,000 cfm held at 24 in high (Colin's 60x24 example reads high).
        w, h = ds.size_rect(16000, 24, 800, 0.08, FRICTION)
        self.assertGreater(w * h, 60 * 24)
        self.assertTrue(passes(16000, ds.SHAPE_RECT, (w, h)))


class TestRect(unittest.TestCase):

    def test_holds_drawn_height(self):
        for cfm in (500, 1200, 2400):
            w, h = ds.size_rect(cfm, 12, 800, 0.08, FRICTION)
            self.assertEqual(h, 12)
            self.assertTrue(passes(cfm, ds.SHAPE_RECT, (w, h)))
            if w - 2 >= max(ds.snap_even(h / 4.0), 6):
                self.assertFalse(passes(cfm, ds.SHAPE_RECT, (w - 2, h)))
            self.assertEqual(w % 2, 0)

    def test_height_steps_up_past_four_to_one(self):
        # 8 in drawn height caps width at 32; a big flow needs more.
        w, h = ds.size_rect(4000, 8, 800, 0.08, FRICTION)
        self.assertGreater(h, 8)
        self.assertLessEqual(float(w) / h, 4.0)
        self.assertTrue(passes(4000, ds.SHAPE_RECT, (w, h)))

    def test_odd_drawn_height_is_evened(self):
        w, h = ds.size_rect(600, 11, 800, 0.08, FRICTION)
        self.assertEqual(h, 12)

    def test_never_below_six(self):
        w, h = ds.size_rect(30, 6, 800, 0.08, FRICTION)
        self.assertEqual((w, h), (6, 6))

    def test_return_limits_are_tighter(self):
        sa = ds.size_rect(1500, 12, 800, 0.08, FRICTION)
        ra = ds.size_rect(1500, 12, 600, 0.05, FRICTION)
        self.assertGreater(ra[0] * ra[1], sa[0] * sa[1])


class TestPlan(unittest.TestCase):

    def test_branch_flex_existing_never_sized(self):
        for kw in ({"is_branch": True}, {"is_flex": True},
                   {"is_existing": True}):
            p = ds.plan_duct(1, ds.SHAPE_RECT, (12, 12), 500, "Supply Air",
                             LIMITS, FRICTION, **kw)
            self.assertIsNone(p.new)
            self.assertTrue(p.skip_reason)

    def test_missing_terminal_flow_not_sized(self):
        p = ds.plan_duct(1, ds.SHAPE_RECT, (20, 12), 1500, "Supply Air",
                         LIMITS, FRICTION, flow_missing=True)
        self.assertIsNone(p.new)
        self.assertIn("no Flow", p.skip_reason)

    def test_no_flow_or_unknown_system(self):
        p = ds.plan_duct(1, ds.SHAPE_ROUND, 10, 0, "Supply Air", LIMITS,
                         FRICTION)
        self.assertIn("airflow", p.skip_reason)
        p = ds.plan_duct(1, ds.SHAPE_ROUND, 10, 500, "Hydronic", LIMITS,
                         FRICTION)
        self.assertIn("no limits", p.skip_reason)

    def test_oval_not_sized(self):
        p = ds.plan_duct(1, None, None, 500, "Supply Air", LIMITS, FRICTION)
        self.assertIn("oval", p.skip_reason)

    def test_changed_flag(self):
        d = ds.size_round(800, 800, 0.08, FRICTION)
        same = ds.plan_duct(1, ds.SHAPE_ROUND, d, 800, "Supply Air", LIMITS,
                            FRICTION)
        self.assertFalse(same.changed)
        other = ds.plan_duct(1, ds.SHAPE_ROUND, d + 4, 800, "Supply Air",
                             LIMITS, FRICTION)
        self.assertTrue(other.changed)
        self.assertEqual(other.new, d)

    def test_rect_plan_reports_performance(self):
        p = ds.plan_duct(1, ds.SHAPE_RECT, (30, 14), 1800, "Supply Air",
                         LIMITS, FRICTION)
        self.assertEqual(p.new[1], 14)
        self.assertLessEqual(p.fpm, 800.0)
        self.assertLessEqual(p.friction, 0.08)


if __name__ == "__main__":
    unittest.main()
