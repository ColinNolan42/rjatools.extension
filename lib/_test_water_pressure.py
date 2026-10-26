# -*- coding: ascii -*-
"""Tests for water_pressure: friction against the code's own Type L data,
fitting kind from geometry, the hot path through the heater, Table 604.3
matching and an end to end pressure budget.

Plain CPython on synthetic graphs. Proves the arithmetic and the tree logic,
not the traversal against a real model.

Run:  python _test_water_pressure.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, COLD, HOT)

import water_graph          # noqa: E402
import water_pressure       # noqa: E402


class Seg(object):
    def __init__(self, eid, gpm, is_return=False):
        self.element_id = eid
        self.demand_gpm = gpm
        self.is_return = is_return


class TestFriction(unittest.TestCase):
    """Against 2024 IPC Figure E103.3(3) (Type L), read off the vector chart
    (2026-10-20). Darcy-Weisbach reads 4 to 10 percent BELOW the chart at
    every point; this test pins that band so a change in either direction is
    seen. The worked example's own column (1.9 at 77 gpm) is not used: the
    example has known transcription errors."""

    POINTS = [  # gpm, Type L ID, chart psi per 100 ft
        (108.0, 2.465, 3.12),   # worked example flow, 2-1/2 in
        (77.0, 2.465, 1.80),
        (38.0, 1.985, 1.40),
        (28.6, 1.505, 3.12),
        (5.0, 0.545, 17.9),     # commentary p. 6-16, 1/2 in
    ]

    def test_chart_points(self):
        for gpm, d, chart in self.POINTS:
            ratio = water_pressure.friction_psi_per_100ft(gpm, d) / chart
            self.assertTrue(0.88 <= ratio <= 1.0,
                            "%s gpm in %s in: ratio %.3f to chart"
                            % (gpm, d, ratio))

    def test_worked_example_value(self):
        got = water_pressure.friction_psi_per_100ft(108.0, 2.465)
        self.assertAlmostEqual(got, 2.99, delta=0.06)

    def test_zero_and_unknown(self):
        self.assertEqual(water_pressure.friction_psi_per_100ft(0.0, 1.025), 0.0)
        self.assertIsNone(water_pressure.friction_psi_per_100ft(None, 1.025))
        self.assertIsNone(water_pressure.friction_psi_per_100ft(5.0, None))


class TestEquivalentLength(unittest.TestCase):

    def test_table_values(self):
        self.assertEqual(water_pressure.equivalent_length_ft("elbow_90", 1.0)[0], 2.5)
        self.assertEqual(water_pressure.equivalent_length_ft("tee_branch", 2.5)[0], 12.0)
        self.assertEqual(water_pressure.equivalent_length_ft("check", 3.0)[0], 14.5)

    def test_dash_is_zero(self):
        self.assertEqual(water_pressure.equivalent_length_ft("tee_run", 0.75)[0], 0.0)

    def test_between_sizes_rounds_up(self):
        self.assertEqual(water_pressure.equivalent_length_ft("elbow_90", 0.9)[0], 2.5)


def fitting(g, nid, center, ends):
    """ends: {other id: (x, y, z)} connector origins."""
    node = add_node(g, nid, water_graph.KIND_FITTING)
    node.location_xyz = list(center)
    node.connectors = [{"connected_element_id": k, "origin_xyz": list(v),
                        "diameter_in": 1.0} for k, v in ends.items()]
    node.connector_count = len(node.connectors)
    return node


class TestFittingKind(unittest.TestCase):

    def test_elbow_90(self):
        g = make_graph()
        f = fitting(g, 1, (0, 0, 0), {10: (-0.1, 0, 0), 11: (0, 0.1, 0)})
        self.assertEqual(water_pressure.classify_fitting(f, 10, 11)[0], "elbow_90")

    def test_elbow_45(self):
        g = make_graph()
        f = fitting(g, 1, (0, 0, 0), {10: (-0.1, 0, 0), 11: (0.0707, 0.0707, 0)})
        self.assertEqual(water_pressure.classify_fitting(f, 10, 11)[0], "elbow_45")

    def test_coupling(self):
        g = make_graph()
        f = fitting(g, 1, (0, 0, 0), {10: (-0.1, 0, 0), 11: (0.1, 0, 0)})
        self.assertEqual(water_pressure.classify_fitting(f, 10, 11)[0], "coupling")

    def test_tee_run_and_branch(self):
        g = make_graph()
        f = fitting(g, 1, (0, 0, 0), {10: (-0.1, 0, 0), 11: (0.1, 0, 0),
                                      12: (0, 0.1, 0)})
        self.assertEqual(water_pressure.classify_fitting(f, 10, 11)[0], "tee_run")
        self.assertEqual(water_pressure.classify_fitting(f, 10, 12)[0], "tee_branch")

    def test_no_geometry_takes_worst_case(self):
        g = make_graph()
        f = fitting(g, 1, (0, 0, 0), {10: (-0.1, 0, 0), 11: (0.1, 0, 0),
                                      12: (0, 0.1, 0)})
        f.location_xyz = None
        kind, _size, note = water_pressure.classify_fitting(f, 10, 11)
        self.assertEqual(kind, "tee_branch")
        self.assertIn("worst case", note)


class TestFixturePressure(unittest.TestCase):

    def _node(self, type_name):
        g = make_graph()
        n = add_node(g, 1, water_graph.KIND_FIXTURE)
        n.type_name = type_name
        return n

    def test_rows(self):
        cases = [("Water Closet, Flushometer Valve", 35.0),
                 ("Water Closet, Flushometer Tank", 20.0),
                 ("Water Closet, Flush Tank", 20.0),
                 ('Urinal, 3/4" Flushometer Valve', 25.0),
                 ("Lavatory", 8.0), ("Shower", 20.0), ("Bathtub", 20.0),
                 ("Service Sink", 8.0), ("Drinking Fountain", 8.0)]
        for type_name, psi in cases:
            self.assertEqual(water_pressure.required_pressure(
                self._node(type_name))[0], psi, type_name)

    def test_unknown(self):
        self.assertIsNone(water_pressure.required_pressure(
            self._node("Eyewash"))[0])

    def test_flush_tank_urinal_as_tank_water_closet(self):
        self.assertEqual(water_pressure.required_pressure(
            self._node("Urinal, Flush Tank"))[0], 20.0)


def pipe(g, nid, parent, system, length, nominal):
    node = add_node(g, nid, water_graph.KIND_PIPE, parent=parent,
                    system=system, length_feet=length)
    node.diameter_inches = nominal
    return node


def build():
    """RPZ(1) -> cold 100 ft 1 in (2) -> lav(3, 10 ft up)
       RPZ(1) -> cold (2) is also parent of heater(4) via 20 ft 3/4 (5)
       heater(4) -> hot 40 ft 3/4 (6) -> lav(3)
    """
    g = make_graph()
    origin = add_node(g, 1, water_graph.KIND_ORIGIN)
    origin.location_xyz = [0.0, 0.0, 0.0]
    g.origin_id = 1
    pipe(g, 2, 1, COLD, 100.0, 1.0)
    lav = add_node(g, 3, water_graph.KIND_FIXTURE, parent=2, system=COLD,
                   cw=1.5, hw=1.5, total=2.0, fixture_name="L-1")
    lav.type_name = "Lavatory"
    lav.location_xyz = [0.0, 0.0, 10.0]
    pipe(g, 5, 2, COLD, 20.0, 0.75)
    heater = add_node(g, 4, water_graph.KIND_HEATER, parent=5, system=COLD)
    g.heater_ids.append(4)
    pipe(g, 6, 4, HOT, 40.0, 0.75)
    lav.parent_by_system[HOT] = 6
    g.nodes[6].children_by_system.setdefault(HOT, []).append(3)
    heater.location_xyz = [0.0, 0.0, 0.0]
    segs = {"segments": [Seg(2, 10.0), Seg(5, 4.0), Seg(6, 4.0)]}
    return g, segs


class TestPathsAndBudget(unittest.TestCase):

    def test_hot_path_runs_through_heater(self):
        g, _ = build()
        self.assertEqual(water_pressure.path_to_fixture(g, 3, HOT),
                         [1, 2, 5, 4, 6, 3])
        self.assertEqual(water_pressure.path_to_fixture(g, 3, COLD),
                         [1, 2, 3])

    def test_budget(self):
        g, segs = build()
        result = water_pressure.analyze(g, segs, start_psi=50.0)
        cold = [o for o in result["outlets"] if o.system == COLD][0]
        rate = water_pressure.friction_psi_per_100ft(10.0, 1.025)
        self.assertAlmostEqual(cold.pipe_psi, rate, places=6)
        self.assertAlmostEqual(cold.static_psi, 4.33, places=6)
        self.assertAlmostEqual(cold.residual_psi, 50.0 - 4.33 - rate, places=6)
        self.assertEqual(cold.required_psi, 8.0)
        self.assertTrue(cold.passes)
        self.assertEqual(len(result["outlets"]), 2)
        self.assertIs(result["critical"], result["outlets"][0])
        self.assertTrue(any("ater heater" in n
                            for n in [x for o in result["outlets"] for x in o.notes]))

    def test_low_start_fails(self):
        g, segs = build()
        result = water_pressure.analyze(g, segs, start_psi=10.0)
        self.assertFalse(result["critical"].passes)

    def test_high_static_warns(self):
        g, segs = build()
        g.nodes[3].location_xyz = [0.0, 0.0, -20.0]
        result = water_pressure.analyze(g, segs, start_psi=75.0)
        self.assertTrue(any("604.8" in w for w in result["warnings"]))

    def test_unpriced_pipe_is_not_judged(self):
        # A pipe with no flow would read as zero loss and could pass a fail.
        g, segs = build()
        segs["segments"][0].demand_gpm = None
        result = water_pressure.analyze(g, segs, start_psi=10.0)
        cold = [o for o in result["outlets"] if o.system == COLD][0]
        self.assertIsNone(cold.passes)
        self.assertIsNone(cold.margin_psi)
        self.assertTrue(any("could not be priced" in w
                            for w in result["warnings"]))

    def test_prv_on_the_way_is_not_judged(self):
        g, segs = build()
        valve = add_node(g, 7, water_graph.KIND_ACCESSORY)
        valve.family_name = "Watts LF25AUB Pressure Reducing Valve"
        # put it between the cold pipe and the lavatory
        g.nodes[3].parent_by_system[COLD] = 7
        valve.parent_by_system[COLD] = 2
        result = water_pressure.analyze(g, segs, start_psi=50.0)
        cold = [o for o in result["outlets"] if o.system == COLD][0]
        self.assertTrue(cold.behind_prv)
        self.assertIsNone(cold.passes)

    def test_unreached_outlet_not_in_static_warning(self):
        g, segs = build()
        g.nodes[3].parent_by_system = {}
        result = water_pressure.analyze(g, segs, start_psi=85.0)
        self.assertIsNone(result["max_static_psi"])
        self.assertFalse(any("604.8" in w for w in result["warnings"]))
        self.assertTrue(any("No path" in w for w in result["warnings"]))

    def test_static_head_by_level(self):
        # Start point hung 3 ft up, fixture connectors at the floor: same
        # level, so no static head (the -1.3 psi Colin saw).
        g, segs = build()
        g.nodes[1].location_xyz = [0.0, 0.0, 3.0]
        g.nodes[3].location_xyz = [0.0, 0.0, 0.0]
        g.nodes[1].level_elev_ft = 0.0
        g.nodes[3].level_elev_ft = 0.0
        cold = [o for o in water_pressure.analyze(g, segs, 50.0)["outlets"]
                if o.system == COLD][0]
        self.assertAlmostEqual(cold.static_psi, 0.0)
        # One floor up, 12 ft: 12 x 0.433.
        g.nodes[3].level_elev_ft = 12.0
        cold = [o for o in water_pressure.analyze(g, segs, 50.0)["outlets"]
                if o.system == COLD][0]
        self.assertAlmostEqual(cold.static_psi, 12.0 * 0.433)

    def test_rows_show_pressure_and_verdict(self):
        g, segs = build()
        rows = water_pressure.table_rows(water_pressure.analyze(g, segs, 50.0))
        self.assertIn(rows[0][4], ("YES", "NO", "NOT CHECKED"))

    def test_rows_shape(self):
        g, segs = build()
        rows = water_pressure.table_rows(water_pressure.analyze(g, segs, 50.0))
        self.assertEqual(len(rows[0]), len(water_pressure.TABLE_COLUMNS))


if __name__ == "__main__":
    _install_revit_stubs()
    unittest.main()
