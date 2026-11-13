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
# The engine moved to lib/duct_perf_run.py (2026-11-12); read both files.
ENGINE = os.path.join(HERE, 'duct_perf_run.py')
SRC = (io.open(SCRIPT, encoding='utf-8').read() + '\n\n\n' + io.open(ENGINE, encoding='utf-8').read()).replace('\r\n', '\n')


def _load(*names):
    ns = {}
    for name in names:
        m = re.search(r'^def %s\(.*?(?=^\S)' % re.escape(name), SRC, re.S | re.M)
        if m is None:
            raise SystemExit('%s not found in script.py' % name)
        exec(m.group(0), ns)
    return ns


NS = _load('_rect_overlap', '_plan_rank', '_crop_box_xy', '_column_centers',
           '_canon_level_pick', '_fit_plan_scale')
NS['LEVEL_SAME_ELEV_FT'] = 0.5
NS['STANDARD_PLAN_SCALES'] = (24, 32, 48, 64, 96, 128, 192)
NS['LEVEL_PLAN_TARGET_IN'] = 24.0



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

    def test_column_top_to_bottom(self):
        """Each extra floor underneath the one before (Colin, 2026-11-11)."""
        ys = NS['_column_centers'](1.0, [2.0, 1.0, 3.0], 0.1)
        self.assertAlmostEqual(ys[0], 1.0)
        self.assertAlmostEqual(ys[1], 1.0 - 1.0 - 0.1 - 0.5)        # bottom of 1st - gap - half
        self.assertAlmostEqual(ys[2], ys[1] - 0.5 - 0.1 - 1.5)

    def test_one_and_none(self):
        self.assertEqual(NS['_column_centers'](1.2, [2.0], 0.1), [1.2])
        self.assertEqual(NS['_column_centers'](1.2, [], 0.1), [])


class CanonLevelTests(unittest.TestCase):
    """Colin, 2026-11-11, TB-M12: a leftover level with no plan at LEVEL 1's
    elevation; ceiling and deck levels with no plan."""
    PLANS = [(1, 0.0), (2, 14.33)]          # LEVEL 1, LEVEL 2

    def test_same_elevation(self):
        self.assertEqual(NS['_canon_level_pick'](0.0, self.PLANS), 1)

    def test_ceiling_goes_to_the_floor_below(self):
        self.assertEqual(NS['_canon_level_pick'](9.0, self.PLANS), 1)
        self.assertEqual(NS['_canon_level_pick'](23.33, self.PLANS), 2)

    def test_deck_just_under_the_next_floor(self):
        self.assertEqual(NS['_canon_level_pick'](14.0, self.PLANS), 2)   # within 0.5 ft

    def test_below_every_plan(self):
        self.assertIsNone(NS['_canon_level_pick'](-10.0, self.PLANS))


class FitScaleTests(unittest.TestCase):
    """The largest drawing whose width fits 24 in., never shrunk."""

    def test_enlarges_to_fit(self):
        # 50 ft wide: 1/4" (48) is 12.5 in., 3/8" (32) is 18.75 in., 1/2" (24) 25 in.
        self.assertEqual(NS['_fit_plan_scale'](50.0, 96), 32)

    def test_never_shrinks(self):
        self.assertEqual(NS['_fit_plan_scale'](300.0, 96), 96)
        self.assertEqual(NS['_fit_plan_scale'](10.0, 24), 24)

    def test_exact_fit(self):
        self.assertEqual(NS['_fit_plan_scale'](96.0, 192), 48)    # 24.0 in. at 1/4"


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
