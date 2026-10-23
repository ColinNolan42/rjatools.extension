# -*- coding: ascii -*-
"""Tests for the diffuser/duct height clearance rule (hvac_graph).

RJA rule (Colin, 2026-10-23): a duct stepping down to a smaller duct or
diffuser neck needs at least 2 in. of height difference. 12 in. to a 10 in.
neck passes; 11 in. to 10 in. fails. The pure block of hvac_graph.py is run
on its own, since hvac_graph itself imports the Revit API.

Run:  python _test_height_clearance.py
"""

import io
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))


def _load():
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("CLEARANCE_MARGIN_IN = ")
    end = src.index("def effective_height_in(")
    ns = {}
    exec(src[start:end], ns)
    return ns


NS = _load()
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


if __name__ == "__main__":
    unittest.main()
