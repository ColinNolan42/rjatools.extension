# -*- coding: ascii -*-
"""Tests for the diffuser/duct height clearance rule (hvac_graph).

RJA rule (Colin, 2026-10-23): a duct stepping down to a smaller duct or
diffuser neck needs at least 2 in. of height difference. 12 in. to a 10 in.
neck passes; 11 in. to 10 in. fails. The pure block of hvac_graph.py is run
on its own, since hvac_graph itself imports the Revit API.

Run:  python _test_height_clearance.py
"""

import io
import math
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))


def _load():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("CLEARANCE_MARGIN_IN = ")
    end = src.index("def effective_height_in(")
    ns = {"math": math}
    exec(src[start:end], ns)
    return ns


def _load_nominal():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("def _nominal_even_in(")
    end = src.index("CLEARANCE_MARGIN_IN = ")
    ns = {"math": math}
    exec(src[start:end], ns)
    return ns["_nominal_even_in"]


NS = _load()
NOMINAL = _load_nominal()
ROUND = NS["_clearance_height_in"]
FAILS = NS["fails_height_clearance"]


def fails(duct_in, down_in):
    return FAILS(ROUND(duct_in), ROUND(down_in))


class TestHeightClearance(unittest.TestCase):

    def test_two_inch_step_passes(self):
        self.assertFalse(fails(12.0, 10.0))

    def test_one_inch_step_fails(self):
        self.assertTrue(fails(11.0, 10.0))

    def test_half_inch_step_fails(self):
        self.assertTrue(fails(10.5, 10.0))

    def test_bigger_step_passes(self):
        self.assertFalse(fails(14.0, 10.0))

    def test_equal_or_larger_downstream_is_not_checked(self):
        self.assertFalse(fails(10.0, 10.0))
        self.assertFalse(fails(10.0, 12.0))

    def test_float_noise_does_not_flag_a_two_inch_step(self):
        # A true 10 in. duct read back as 10.000000000000002 in. (2026-08-26).
        self.assertFalse(fails(12.0, 10.000000000000002))
        self.assertFalse(fails(11.999999999999998, 10.0))

    def test_unknown_height_is_not_flagged(self):
        self.assertFalse(FAILS(None, 10.0))
        self.assertFalse(FAILS(12.0, None))

    def test_margin_is_two_inches(self):
        self.assertEqual(NS["CLEARANCE_MARGIN_IN"], 2.0)


class TestNominalEven(unittest.TestCase):
    """_nominal_even_in rounds half way UP on every interpreter (Colin,
    2026-11-02): CPython's round() sends 9 to 8, IronPython's to 10."""

    def test_half_way_rounds_up(self):
        self.assertEqual(NOMINAL(9.0), 10.0)
        self.assertEqual(NOMINAL(13.0), 14.0)

    def test_nearest_and_noise(self):
        self.assertEqual(NOMINAL(10.000000000000002), 10.0)
        self.assertEqual(NOMINAL(9.9), 10.0)
        self.assertEqual(NOMINAL(10.9), 10.0)
        self.assertEqual(NOMINAL(11.1), 12.0)


class TestClearanceTexts(unittest.TestCase):
    """The clearance Reason names its cause (Colin, 2026-11-02)."""

    TEXTS = staticmethod(NS["clearance_texts"])

    def test_main_step(self):
        # TB-M01 seed S01-5: SA-A2 12x12 into SA-A3 8x11.
        reason, note = self.TEXTS('step', 'Main', 12.0, 11.0,
                                  '3650050 (12"x12")', '3650051 (8"x11")')
        self.assertEqual(reason, 'Main step 12" to 11", 2" min')
        self.assertIn('Make duct 3650051 10" tall or less, or keep it 12" tall', note)
        self.assertIn('3650050 (12"x12")', note)

    def test_branch_too_tall_for_main(self):
        reason, note = self.TEXTS('branch', 'Main', 12.0, 11.0,
                                  '100 (24"x12")', '200 (12"x11")')
        self.assertEqual(reason, 'Branch 11" too tall for 12" main')
        self.assertIn('Raise the main to 14" tall', note)
        self.assertIn('make the branch 10" tall or less', note)

    def test_neck(self):
        reason, note = self.TEXTS('neck', 'Branch', 9.0, 8.0,
                                  '300 (9")', '400 (8" neck)')
        self.assertEqual(reason, 'Neck 8" too tall for 9" branch')
        self.assertIn('Raise the duct to 10" tall', note)
        self.assertIn('reduce the neck to 7" or less', note)

    def test_branch_step_and_fraction(self):
        reason, _ = self.TEXTS('step', 'Branch', 10.5, 10.0, 'a', 'b')
        self.assertEqual(reason, 'Branch step 10.5" to 10", 2" min')


if __name__ == "__main__":
    unittest.main()
