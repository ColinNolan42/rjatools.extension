# -*- coding: ascii -*-
"""Tests for water_performance. Plain CPython.

Run:  python _test_water_performance.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import _install_revit_stubs   # noqa: E402

_install_revit_stubs()

import water_performance as wp               # noqa: E402
import water_sizing_engine as eng            # noqa: E402


def seg(drawn, req, nominal, wsfu=10.0, gpm=10.0, **kw):
    s = eng.SegmentResult(1)
    s.existing_size_in = drawn
    s.size_inches = req
    s.nominal_size = nominal
    s.demand_wsfu = wsfu
    s.demand_gpm = gpm
    for k, v in kw.items():
        setattr(s, k, v)
    return s


class T(unittest.TestCase):
    def test_id_table(self):
        self.assertAlmostEqual(wp.inside_diameter_in(0.75), 0.785)
        self.assertAlmostEqual(wp.inside_diameter_in(3.0), 2.945)
        self.assertIsNone(wp.inside_diameter_in(4.0))
        self.assertIsNone(wp.inside_diameter_in(None))

    def test_velocity_matches_note(self):
        # 35.2 gpm in 1-1/4 in is about 8.98 fps (sizing method note)
        self.assertAlmostEqual(wp.velocity_fps(35.2, 1.25), 8.98, delta=0.02)
        self.assertIsNone(wp.velocity_fps(None, 1.0))
        self.assertIsNone(wp.velocity_fps(5.0, 4.0))

    def test_equal_is_green(self):
        r = wp.classify(seg(1.0, 1.0, "1", gpm=5.0), "cold")
        self.assertEqual(r.label, wp.GREEN)

    def test_under_red_over_purple(self):
        self.assertEqual(wp.classify(seg(0.75, 1.0, "1"), "cold").label, wp.RED)
        self.assertEqual(wp.classify(seg(1.5, 1.0, "1"), "cold").label,
                         wp.PURPLE)

    def test_equal_but_fast_is_yellow(self):
        r = wp.classify(seg(1.25, 1.25, "1-1/4", gpm=35.2), "cold")
        self.assertEqual(r.label, wp.YELLOW)

    def test_table_limit_pipes_are_not_flagged(self):
        # firm table limits, computed 8.02 fps cold 2-1/2 and 5.04 fps hot
        self.assertEqual(
            wp.classify(seg(2.5, 2.5, "2-1/2", gpm=119.2), "cold").label,
            wp.GREEN)
        self.assertEqual(
            wp.classify(seg(2.5, 2.5, "2-1/2", gpm=75.0), "hot").label,
            wp.GREEN)

    def test_exceeds_is_gray(self):
        s = seg(3.0, None, None, exceeds_table=True)
        self.assertEqual(wp.classify(s, "cold").label, wp.GRAY)

    def test_zero_demand_gray(self):
        s = seg(1.0, 0.5, "1/2", wsfu=0.0)
        self.assertEqual(wp.classify(s, "cold").label, wp.GRAY)

    def test_missing_drawn_gray(self):
        self.assertEqual(wp.classify(seg(None, 1.0, "1"), "cold").label,
                         wp.GRAY)

    def test_return_uses_circ_gpm_and_hot(self):
        s = seg(0.5, 0.5, "1/2", wsfu=0.0, is_return=True, circ_gpm=1.0)
        r = wp.classify(s, "hot")
        self.assertEqual(r.label, wp.GREEN)
        self.assertEqual(r.gpm, 1.0)
        self.assertEqual(r.limit_fps, 5.0)

    def test_worst(self):
        self.assertEqual(wp.worst_label([wp.GREEN, wp.RED, wp.GRAY]), wp.RED)
        self.assertIsNone(wp.worst_label([]))


if __name__ == "__main__":
    unittest.main()
