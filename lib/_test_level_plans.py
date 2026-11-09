# -*- coding: utf-8 -*-
"""CPython tests for Duct Performance's one colored plan per level (Colin,
2026-11-09). The pure helpers are pulled out of the pushbutton script by
regex and run here; the Revit calls around them are checked by source.

Run: python lib/_test_level_plans.py
"""
import io
import os
import re
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(os.path.dirname(HERE), 'Developer Tools.tab', 'Performance.panel',
                      'Duct Performance.pushbutton', 'script.py')
SRC = io.open(SCRIPT, encoding='utf-8').read().replace('\r\n', '\n')


def _load(*names):
    ns = {}
    for name in names:
        m = re.search(r'^def %s\(.*?(?=^\S)' % re.escape(name), SRC, re.S | re.M)
        if m is None:
            raise SystemExit('%s not found in script.py' % name)
        exec(m.group(0), ns)
    return ns


NS = _load('_rect_overlap', '_plan_rank', '_crop_box_xy', '_row_centers')


class OverlapTests(unittest.TestCase):

    def test_overlap_area(self):
        f = NS['_rect_overlap']
        self.assertEqual(f((0, 0, 10, 10), (5, 5, 20, 20)), 25.0)
        self.assertEqual(f((0, 0, 10, 10), (10, 0, 20, 10)), 0.0)   # touching edge
        self.assertEqual(f((0, 0, 10, 10), (30, 30, 40, 40)), 0.0)

    def test_crop_box_with_margin(self):
        f = NS['_crop_box_xy']
        self.assertEqual(f([(0, 0), (10, 4)], 5.0), (-5.0, -5.0, 15.0, 9.0))
        self.assertIsNone(f([], 5.0))


class PlanRankTests(unittest.TestCase):
    """Colin, 2026-11-09 (TB-M12C took TB-M01C's plan): a candidate plan's
    crop shall overlap the level's system; ranked by template, then cropped
    over uncropped, then the larger overlap."""

    SYS = (0.0, 0.0, 50.0, 30.0)

    def test_other_bench_plan_is_out(self):
        self.assertIsNone(NS['_plan_rank'](2, (400.0, 0.0, 450.0, 30.0), self.SYS))

    def test_non_mechanical_plan_is_out(self):
        self.assertIsNone(NS['_plan_rank'](0, (0.0, 0.0, 50.0, 30.0), self.SYS))
        self.assertIsNone(NS['_plan_rank'](0, None, self.SYS))

    def test_template_ranks_first(self):
        r = NS['_plan_rank']
        same = r(2, (40.0, 20.0, 60.0, 40.0), self.SYS)      # small overlap
        rja = r(1, (0.0, 0.0, 50.0, 30.0), self.SYS)         # full overlap
        self.assertGreater(same, rja)

    def test_cropped_over_uncropped_then_larger_overlap(self):
        r = NS['_plan_rank']
        uncropped = r(2, None, self.SYS)
        small = r(2, (40.0, 20.0, 60.0, 40.0), self.SYS)
        big = r(2, (-10.0, -10.0, 60.0, 40.0), self.SYS)
        self.assertIsNotNone(uncropped)
        self.assertGreater(small, uncropped)
        self.assertGreater(big, small)

    def test_no_system_rect_takes_only_an_uncropped_plan(self):
        self.assertIsNone(NS['_plan_rank'](2, (0, 0, 1, 1), None))
        self.assertIsNotNone(NS['_plan_rank'](2, None, None))


class LayoutTests(unittest.TestCase):

    def test_row_left_to_right(self):
        xs = NS['_row_centers'](1.0, [2.0, 1.0, 3.0], 0.1)
        self.assertAlmostEqual(xs[0], 1.0)
        self.assertAlmostEqual(xs[1], 1.0 + 1.0 + 0.1 + 0.5)        # right of 1st + gap + half
        self.assertAlmostEqual(xs[2], xs[1] + 0.5 + 0.1 + 1.5)

    def test_one_and_none(self):
        self.assertEqual(NS['_row_centers'](1.2, [2.0], 0.1), [1.2])
        self.assertEqual(NS['_row_centers'](1.2, [], 0.1), [])


class WiringTests(unittest.TestCase):

    def test_levels_from_ducts_and_terminals_only(self):
        # A level holding only the unit gets no plan (TB-M12C LEVEL 2 - DECK).
        m = re.search(r'    level_ids = \{\}\n(.*?)    level_targets, missing_levels', SRC, re.S)
        self.assertIsNotNone(m)
        self.assertNotIn('all_root_ids', m.group(1))
        self.assertIn('for _nid in all_terminals.keys():', m.group(1))

    def test_plan_pick_uses_overlap(self):
        self.assertIn('key = _plan_rank(rank, _view_crop_rect(v), level_rects.get(lid))', SRC)
        self.assertIn('_target_plan_views(active_view, level_ids, sheets_by_view,\n'
                      '                                                       level_rects)', SRC)

    def test_unverified_api_is_guarded(self):
        # ViewPlan.Create and the CropBox read and write are wrapped and
        # logged, with a fallback (Colin, 2026-11-09: no MCP port yet).
        m = re.search(r'^def _create_level_plan\(.*?(?=^\S)', SRC, re.S | re.M).group(0)
        self.assertIn('try:\n        v = ViewPlan.Create(doc, like_view.GetTypeId(), lvl_eid)', m)
        self.assertIn('return None', m)
        for name in ('_view_crop_rect', '_crop_to_elements'):
            body = re.search(r'^def %s\(.*?(?=^\S)' % name, SRC, re.S | re.M).group(0)
            self.assertIn('except Exception as ex:', body)
        # A plan that cannot be made is an Action required line, not silent.
        self.assertIn("_correct('warn', 'WARNING: no mechanical floor plan on {} and one could '",
                      SRC)
        # Nothing at all to draw falls back to the active view alone.
        self.assertIn('if multi_level and not targets:', SRC)

    def test_view_range_opening_is_gone(self):
        self.assertNotIn('_open_view_range', SRC)
        self.assertNotIn('system_top_z', SRC)

    def test_single_level_unchanged(self):
        self.assertIn("targets = [(None, active_view, sheets_by_view.get(eid_int(active_view.Id)))]",
                      SRC)
        self.assertIn('diagram_vp.SetBoxCenter(XYZ(1.201, 1.212, 0))', SRC)


if __name__ == '__main__':
    unittest.main(verbosity=2)
