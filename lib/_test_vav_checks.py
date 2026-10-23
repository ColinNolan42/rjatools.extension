# -*- coding: ascii -*-
"""Tests for vav_checks: VAV box detection, the MP set, the VAV red rules,
the parallel unit check and the shared limits lookup.

vav_checks is pure Python, so the real module is tested against small fake
networks built from ids, kinds and system classes.

Run:  python _test_vav_checks.py
"""

import io
import logging
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import duct_sizing as ds  # noqa: E402
import vav_checks as vc   # noqa: E402

# The box warnings are expected here; keep the test output readable.
logging.getLogger('vav_checks').setLevel(logging.ERROR)


def _load_friction():
    """The real Darcy-Weisbach friction block of hvac_graph.py (same trick as
    _test_duct_sizing: hvac_graph itself imports the Revit API)."""
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("STANDARD_AIR_DENSITY_LB_FT3 = ")
    end = src.index("def fitting_family_name(")
    ns = {"math": math}
    exec(src[start:end], ns)
    return ns["duct_friction_loss_per_100ft"]


FRICTION = _load_friction()
FIRM = {'Supply Air': (800, 0.08), 'Return Air': (600, 0.05)}


class Net(object):
    """A fake network: id -> (kind, own system class, entry direction)."""

    def __init__(self, root, children, info):
        self.root = root
        self.children = children
        self.info = info

    def kind_of(self, nid):
        return self.info.get(nid, (None, None, None))[0]

    def class_of(self, nid):
        return self.info.get(nid, (None, None, None))[1]

    def entry_dir_of(self, nid):
        return self.info.get(nid, (None, None, None))[2]

    def analyze(self):
        return vc.analyze_system(self.root, self.children, self.kind_of,
                                 self.class_of, self.entry_dir_of)


SA, RA = 'Supply Air', 'Return Air'


def _vav_net():
    """
    AHU 1
      +- duct 2 (SA) -- tap 3 -- duct 4 -- BOX 5 (in) -- duct 6 -- diffuser 7
      |            +-- tap 8 -- duct 9 -- diffuser 10          (on the MP main)
      |            +-- tap 11 -- flex 12 -- BOX 13 (dir ?) -- elbow 14 -- duct 15
      |                                                        -- diffuser 16
      |                                                        -- diffuser 17
      +- duct 20 (RA) -- grille 21
    """
    children = {1: [2, 20], 2: [3, 8, 11], 3: [4], 4: [5], 5: [6], 6: [7],
                8: [9], 9: [10], 11: [12], 12: [13], 13: [14], 14: [15],
                15: [16, 17], 20: [21]}
    info = {
        1: ('equipment', None, None),
        2: ('duct', SA, None), 3: ('fitting', None, None), 4: ('duct', SA, None),
        5: ('equipment', None, 'in'), 6: ('duct', SA, None), 7: ('terminal', SA, None),
        8: ('fitting', None, None), 9: ('duct', SA, None), 10: ('terminal', SA, None),
        11: ('fitting', None, None), 12: ('duct', SA, None),
        13: ('equipment', None, None), 14: ('fitting', None, None),
        15: ('duct', SA, None), 16: ('terminal', SA, None), 17: ('terminal', SA, None),
        20: ('duct', RA, None), 21: ('terminal', RA, None),
    }
    return Net(1, children, info)


class TestAnalyzeSystem(unittest.TestCase):

    def setUp(self):
        self.vs = _vav_net().analyze()

    def test_boxes_found_by_position(self):
        self.assertEqual(sorted(self.vs.boxes), [5, 13])
        self.assertTrue(self.vs.is_vav)

    def test_unreadable_direction_is_a_box_and_listed(self):
        self.assertEqual(self.vs.unknown_direction, [13])

    def test_mp_set_is_unit_to_box_inlets_only(self):
        self.assertEqual(sorted(self.vs.mp_ids), [2, 3, 4, 11, 12])

    def test_after_the_box_is_not_mp(self):
        for nid in (6, 14, 15):
            self.assertNotIn(nid, self.vs.mp_ids)

    def test_return_is_never_mp(self):
        self.assertNotIn(20, self.vs.mp_ids)

    def test_the_tap_to_a_lone_diffuser_is_not_mp_but_flagged(self):
        self.assertNotIn(9, self.vs.mp_ids)
        self.assertEqual(self.vs.mp_terminals, {10: 9})

    def test_inlet_and_outlet_ducts(self):
        self.assertEqual(self.vs.box_inlet, {5: 4, 13: 12})
        self.assertEqual(self.vs.box_outlets[5], [6])
        self.assertEqual(self.vs.box_outlets[13], [15])

    def test_no_box_means_no_vav(self):
        net = Net(1, {1: [2], 2: [3]},
                  {1: ('equipment', None, None), 2: ('duct', SA, None),
                   3: ('terminal', SA, None)})
        vs = net.analyze()
        self.assertFalse(vs.is_vav)
        self.assertEqual(vs.mp_ids, set())
        self.assertEqual(vs.mp_terminals, {})

    def test_entered_through_outlet_is_parallel_not_a_box(self):
        net = Net(1, {1: [2], 2: [3], 3: [4], 4: [5]},
                  {1: ('equipment', None, None), 2: ('duct', SA, None),
                   3: ('equipment', None, 'out'), 4: ('duct', SA, None),
                   5: ('terminal', SA, None)})
        vs = net.analyze()
        self.assertEqual(vs.reversed, [3])
        self.assertEqual(vs.boxes, [])

    def test_return_side_equipment_is_not_a_box(self):
        net = Net(1, {1: [2], 2: [3], 3: [4], 4: [5]},
                  {1: ('equipment', None, None), 2: ('duct', RA, None),
                   3: ('equipment', None, 'in'), 4: ('duct', RA, None),
                   5: ('terminal', RA, None)})
        vs = net.analyze()
        self.assertEqual(vs.boxes, [])

    def test_mp_is_by_position_not_by_name(self):
        """A custom System Type name with Supply in it still gets MP by position,
        and an LP duct named the same after the box does not."""
        name = 'Supply Air - VAV'
        net = Net(1, {1: [2], 2: [3], 3: [4], 4: [5]},
                  {1: ('equipment', None, None), 2: ('duct', name, None),
                   3: ('equipment', None, 'in'), 4: ('duct', name, None),
                   5: ('terminal', name, None)})
        vs = net.analyze()
        self.assertEqual(vs.mp_ids, set([2]))

    def test_a_cycle_does_not_hang(self):
        net = Net(1, {1: [2], 2: [3], 3: [2, 4], 4: [5]},
                  {1: ('equipment', None, None), 2: ('duct', SA, None),
                   3: ('duct', SA, None), 4: ('equipment', None, 'in'),
                   5: ('terminal', SA, None)})
        vs = net.analyze()
        self.assertEqual(vs.boxes, [4])


class TestParallel(unittest.TestCase):

    def test_shared_duct_flags_both(self):
        out = vc.overlapping_roots({1: set([1, 2, 3]), 10: set([10, 3, 11])})
        self.assertEqual(out, set([1, 10]))

    def test_disjoint_is_fine(self):
        self.assertEqual(vc.overlapping_roots({1: set([1, 2]), 10: set([10, 11])}),
                         set())

    def test_a_box_picked_with_its_ahu_is_nested_not_parallel(self):
        ahu = set([1, 2, 5, 6, 7])
        box = set([5, 6, 7])
        self.assertEqual(vc.overlapping_roots({1: ahu, 5: box}), set())


class TestLimitsAndFields(unittest.TestCase):

    def test_mp_duct_takes_mp_limits_whatever_its_name(self):
        limits = {SA: (800, 0.08), vc.MP_KEY: (1800, 0.25)}
        self.assertEqual(vc.limits_for(SA, limits, FIRM, is_mp=True), (1800, 0.25))
        self.assertEqual(vc.limits_for('Odd Name', limits, FIRM, is_mp=True),
                         (1800, 0.25))

    def test_mp_falls_back_to_the_defaults(self):
        self.assertEqual(vc.limits_for(SA, {}, FIRM, is_mp=True),
                         (vc.MP_DEFAULT_FPM, vc.MP_DEFAULT_FRICTION))

    def test_lp_lookup_is_unchanged(self):
        self.assertEqual(vc.limits_for(SA, {SA: (700, 0.07)}, FIRM), (700, 0.07))
        self.assertEqual(vc.limits_for(RA, {}, FIRM), (600, 0.05))
        self.assertEqual(vc.limits_for('Unknown', {}, FIRM), (600, 0.05))

    def test_plan_defaults(self):
        d = vc.default_settings()
        self.assertEqual(d['mp_fpm'], 2000.0)
        self.assertEqual(d['mp_fric'], 0.30)
        self.assertEqual(d['box_drop'], 0.25)
        self.assertIsNone(d['unit_max_esp'])
        self.assertEqual(d['min_box_inwc'], 0.75)

    def test_parse_field(self):
        self.assertEqual(vc.parse_field('', 'pos_or_blank'), (None, None))
        self.assertEqual(vc.parse_field(' 1.5 ', 'pos_or_blank'), (1.5, None))
        self.assertIsNotNone(vc.parse_field('', 'pos')[1])
        self.assertIsNotNone(vc.parse_field('0', 'pos')[1])
        self.assertIsNotNone(vc.parse_field('-1', 'nonneg')[1])
        self.assertEqual(vc.parse_field('0', 'nonneg'), (0.0, None))
        self.assertIsNotNone(vc.parse_field('abc', 'nonneg')[1])
        self.assertIsNotNone(vc.parse_field('nan', 'pos')[1])

    def test_reasons_fit_the_34_char_reason_column(self):
        for text in (vc.REASON_DIFFUSER_ON_MP, vc.ROLE_MP):
            self.assertLessEqual(len(text), 34, text)

    def test_basis_rows_name_sources_and_are_ascii(self):
        rows = vc.basis_rows({'unit_max_esp': 1.5})
        joined = ' '.join(a + ' ' + b for a, b in rows)
        joined.encode('ascii')
        for needle in ('EDR VAV Guide p3, p100', 'EDR p99', 'Taylor 2019 p54',
                       '2000 fpm', '0.30', '1.50'):
            self.assertIn(needle, joined)
        self.assertNotIn(' - ', joined)
        self.assertNotIn('straight', joined.lower())

    def test_pressure_left(self):
        self.assertIsNone(vc.pressure_left(None, 0.4))
        self.assertAlmostEqual(vc.pressure_left(1.5, 0.4), 1.1)
        self.assertTrue(vc.pressure_low(0.7, 0.75))
        self.assertFalse(vc.pressure_low(0.75, 0.75))
        self.assertFalse(vc.pressure_low(None, 0.75))


class TestSizeDuctMp(unittest.TestCase):
    """duct_sizing.plan_duct sizes an MP main against the MP row."""

    LIMITS = {SA: (800.0, 0.08), vc.MP_KEY: (2000.0, 0.30)}

    def test_mp_main_is_sized_smaller_than_lp(self):
        lp = ds.plan_duct(1, ds.SHAPE_ROUND, 16.0, 1500.0, SA, self.LIMITS, FRICTION)
        mp = ds.plan_duct(1, ds.SHAPE_ROUND, 16.0, 1500.0, SA, self.LIMITS, FRICTION,
                          limit_key=vc.MP_KEY)
        self.assertEqual(mp.limit_key, vc.MP_KEY)
        self.assertEqual(lp.limit_key, SA)
        self.assertLess(mp.new, lp.new)
        self.assertLessEqual(mp.fpm, 2000.0 + 1e-6)
        self.assertLessEqual(mp.friction, 0.30 + 1e-9)

    def test_missing_mp_row_is_not_sized(self):
        plan = ds.plan_duct(1, ds.SHAPE_ROUND, 10.0, 500.0, SA, {SA: (800, 0.08)},
                            FRICTION, limit_key=vc.MP_KEY)
        self.assertIsNone(plan.new)
        self.assertIn(vc.MP_KEY, plan.skip_reason)

    def test_basis_names_the_mp_rule_only_with_vav(self):
        self.assertFalse(any('VAV' in l for l in ds.basis_lines({SA: (800, 0.08)})))
        self.assertTrue(any('VAV' in l for l in ds.basis_lines(self.LIMITS)))


class TestSideHelpers(unittest.TestCase):

    def test_same_side(self):
        self.assertTrue(vc.same_side(SA, 'Supply Air - VAV'))
        self.assertFalse(vc.same_side(SA, RA))
        self.assertTrue(vc.same_side(RA, 'Exhaust Air'))
        self.assertTrue(vc.same_side(None, RA))
        self.assertTrue(vc.same_side(SA, ''))

    def test_entry_class_skips_unclassified_fittings(self):
        parents = {2: 1, 3: 2, 4: 3}
        cls = {1: None, 2: SA, 3: None, 4: None}
        self.assertEqual(vc.entry_class(4, parents, cls.get), SA)
        self.assertIsNone(vc.entry_class(2, parents, cls.get))

    def test_counted_children(self):
        side = {10: SA, 11: RA, 12: None}
        self.assertEqual(vc.counted_children([10, 11, 12], SA, side.get), [10, 12])
        self.assertEqual(vc.counted_children([10, 11, 12], None, side.get),
                         [10, 11, 12])


# -- C1: compute_cfm / compute_downstream_terminal_ids at mid tree equipment ---

class _E(object):
    def __init__(self, kind, cls=None):
        self.kind = kind
        self.cls = cls


def _load_accumulators():
    """The real _upward_children, compute_cfm and
    compute_downstream_terminal_ids out of hvac_graph.py, run against fake
    element helpers (the module itself imports the Revit API)."""
    src = io.open(os.path.join(HERE, "hvac_graph.py"), encoding="utf-8").read()
    start = src.index("def _upward_children(")
    end = src.index("def solid_fill_pattern_id(")

    def is_equipment(e):
        return e is not None and e.kind == 'equipment'

    def elem_sys_class(e):
        return e.cls if e.kind in ('duct', 'terminal') else None

    def _branch_sys_class(nid, nodes, children):
        e = nodes.get(nid)
        if e is None:
            return None
        if e.kind in ('duct', 'terminal'):
            return e.cls
        if e.kind in ('fitting', 'accessory'):
            for c in children.get(nid, []):
                got = _branch_sys_class(c, nodes, children)
                if got is not None:
                    return got
        return None

    ns = {'vav_checks': vc, 'is_equipment': is_equipment,
          'elem_sys_class': elem_sys_class, '_branch_sys_class': _branch_sys_class}
    exec(src[start:end], ns)
    return ns


ACC = _load_accumulators()


class TestMidTreeEquipmentFlow(unittest.TestCase):
    """AHU 1 -> MP duct 2 -> FPB 3 -> LP duct 4 -> 400 cfm diffuser 5
                                 `-> return duct 6 -> 300 cfm grille 7"""

    def setUp(self):
        self.nodes = {1: _E('equipment'), 2: _E('duct', SA), 3: _E('equipment'),
                      4: _E('duct', SA), 5: _E('terminal', SA),
                      6: _E('duct', RA), 7: _E('terminal', RA)}
        self.children = {1: [2], 2: [3], 3: [4, 6], 4: [5], 6: [7]}

    def test_mp_duct_reads_only_the_supply_diffusers(self):
        cfm, by_class = ACC['compute_cfm'](1, self.nodes, self.children,
                                           {5: 400.0, 7: 300.0})
        self.assertEqual(cfm[2], 400.0)
        self.assertEqual(cfm[3], 400.0)
        # the breakdown still names the return
        self.assertEqual(by_class[3], {SA: 400.0, RA: 300.0})

    def test_root_keeps_its_combined_total(self):
        children = {1: [2, 8], 2: [5], 8: [7]}
        nodes = {1: _E('equipment'), 2: _E('duct', SA), 5: _E('terminal', SA),
                 8: _E('duct', RA), 7: _E('terminal', RA)}
        cfm, _ = ACC['compute_cfm'](1, nodes, children, {5: 400.0, 7: 300.0})
        self.assertEqual(cfm[1], 700.0)

    def test_return_grille_is_not_downstream_of_the_mp_duct(self):
        down = ACC['compute_downstream_terminal_ids'](
            1, self.nodes, self.children, set([5, 7]))
        self.assertEqual(down[2], set([5]))
        self.assertEqual(down[6], set([7]))

    def test_unknown_entry_class_counts_everything(self):
        nodes = dict(self.nodes)
        nodes[2] = _E('fitting')
        cfm, _ = ACC['compute_cfm'](1, nodes, self.children, {5: 400.0, 7: 300.0})
        self.assertEqual(cfm[2], 700.0)


# -- W3 / S1: red_reasons ------------------------------------------------------

class TestRedReasons(unittest.TestCase):
    """AHU 1 -> flex 2 -> fitting 3 -> BOX 4 -> fitting 5 -> duct 6 -> T7."""

    def setUp(self):
        vs = vc.VavSystem()
        vs.boxes = [4]
        vs.box_inlet = {4: 2}
        vs.box_outlets = {4: [6]}
        self.vs = vs

    def test_nothing_to_flag(self):
        self.assertEqual(vc.red_reasons([self.vs]), {})

    def test_flex_at_inlet_is_not_flagged(self):
        # Rule removed (Colin, 2026-10-23): flex 2 at the box inlet is fine.
        self.assertFalse(hasattr(vc, 'REASON_FLEX_AT_INLET'))
        self.assertEqual(vc.red_reasons([self.vs]), {})

    def test_diffuser_on_mp(self):
        self.vs.mp_terminals = {7: 6}
        self.assertEqual(vc.red_reasons([self.vs]),
                         {6: vc.REASON_DIFFUSER_ON_MP})

    def test_terminal_with_no_feed_duct_is_skipped(self):
        self.vs.mp_terminals = {7: None}
        self.assertEqual(vc.red_reasons([self.vs]), {})

    def test_systems_may_be_a_generator(self):
        self.vs.mp_terminals = {7: 6}
        out = vc.red_reasons((v for v in [self.vs]))
        self.assertEqual(out, {6: vc.REASON_DIFFUSER_ON_MP})


# -- W5: VAV inlet runouts -------------------------------------------------------

class TestInletRunouts(unittest.TestCase):

    def test_one_box_downstream_is_a_runout(self):
        vs = _vav_net().analyze()
        # 2 feeds both boxes (MP main); 3, 4 feed box 5; 11, 12 feed box 13
        self.assertEqual(sorted(vs.inlet_runout_ids), [3, 4, 11, 12])
        self.assertNotIn(2, vs.inlet_runout_ids)
        self.assertTrue(vs.inlet_runout_ids <= vs.mp_ids)

    def test_single_box_system_is_all_runout(self):
        net = Net(1, {1: [2], 2: [3], 3: [4], 4: [5], 5: [6]},
                  {1: ('equipment', None, None), 2: ('duct', SA, None),
                   3: ('fitting', None, None), 4: ('duct', SA, None),
                   5: ('equipment', None, 'in'), 6: ('terminal', SA, None)})
        vs = net.analyze()
        self.assertEqual(vs.inlet_runout_ids, set([2, 3, 4]))

    def test_labels_fit_the_columns(self):
        self.assertLessEqual(len(vc.REASON_INLET_OK), 34)
        self.assertLessEqual(len(vc.ROLE_INLET), 18)
        self.assertEqual(vc.REASON_INLET_OK, 'VAV INLET, SIZED TO BOX')
        self.assertEqual(vc.ROLE_INLET, 'VAV Inlet')

    def test_size_duct_never_resizes_a_runout(self):
        limits = {SA: (800.0, 0.08), vc.MP_KEY: (2000.0, 0.30)}
        plan = ds.plan_duct(4, ds.SHAPE_ROUND, 6.0, 1500.0, SA, limits, FRICTION,
                            limit_key=vc.MP_KEY, is_inlet_runout=True)
        self.assertIsNone(plan.new)
        self.assertEqual(plan.skip_reason, ds.SKIP_INLET_RUNOUT)

    def test_size_duct_still_sizes_an_mp_main(self):
        limits = {SA: (800.0, 0.08), vc.MP_KEY: (2000.0, 0.30)}
        plan = ds.plan_duct(2, ds.SHAPE_ROUND, 6.0, 1500.0, SA, limits, FRICTION,
                            limit_key=vc.MP_KEY, is_inlet_runout=False)
        self.assertIsNotNone(plan.new)


def _run():
    suite = unittest.TestLoader().loadTestsFromModule(sys.modules[__name__])
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    print("=" * 70)
    print("SUMMARY: {} run, {} passed, {} failed, {} errors".format(
        result.testsRun,
        result.testsRun - len(result.failures) - len(result.errors),
        len(result.failures), len(result.errors)))
    print("=" * 70)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(_run())
