# -*- coding: ascii -*-
"""Tests for starting Size Gas at a new pipe on an existing building.

Runs in plain CPython on synthetic graphs. Covers the existing length joining
the developed length, the effect on the sizes, and the existing pipe check.
It does NOT test the pick, the Existing Conditions box, the phase and cap
read against a real model, or the One-Line drawing. Those need a live Revit
run.

Run:  python _test_gas_tie_in.py
"""

import sys
import unittest

from _test_prv_zones import _node, _edge, _finish, _Stub   # installs the stubs

import gas_tables      # noqa: E402
import pipe_graph      # noqa: E402
import sizing_engine   # noqa: E402
import tie_in          # noqa: E402

TABLE = "402.4(2)"     # Sched 40 steel, natural gas, under 2 psi, 0.5 in. w.c.
HEAT = 1000.0          # so CFH equals MBH and the table reads directly
ORIGIN = -100          # the stand-in node at the existing connection


def _build(prv=False):
    """connection(-100) -> pipe 100 (10 ft) -> tee 2
           tee 2 -> pipe 101 (20 ft) -> RTU-1, 120 MBH
           tee 2 -> pipe 102 (10 ft) -> GUH-1, 60 MBH

    New longest run 30 ft, no elbows. New load 180 MBH.
    """
    g = pipe_graph.NetworkGraph()
    g.origin_id = ORIGIN
    g.origin_is_tie_in = True
    _node(g, ORIGIN, "meter")
    _node(g, 2, "tee")
    _node(g, 4, "fixture", 120.0, "RTU-1")
    _node(g, 5, "fixture", 60.0, "GUH-1")
    _edge(g, 100, ORIGIN, 2, 10.0)
    _edge(g, 101, 2, 4, 20.0)
    _edge(g, 102, 2, 5, 10.0)
    _finish(g)
    return g


def _size(g):
    return sizing_engine.size_system(
        g, "Schedule 40 Steel", 0.5, TABLE, heat_content_btu_per_cf=HEAT)


class TestExistingLengthJoinsTheRun(unittest.TestCase):

    def test_new_run_alone(self):
        g = _build()
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 30.0)
        self.assertIsNone(g.existing)

    def test_existing_length_is_added(self):
        g = _build()
        pipe_graph.apply_existing_conditions(g, "1-1/4", 60.0, 200.0)
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 90.0)
        self.assertAlmostEqual(g.longest_run["existing_length_feet"], 60.0)
        self.assertAlmostEqual(g.longest_run["pipe_length_feet"], 30.0)
        self.assertEqual(g.existing["size_nominal"], "1-1/4")
        self.assertAlmostEqual(g.existing["load_mbh"], 200.0)

    def test_applying_twice_does_not_add_twice(self):
        g = _build()
        pipe_graph.apply_existing_conditions(g, "1-1/4", 60.0, 200.0)
        pipe_graph.apply_existing_conditions(g, "1-1/4", 60.0, 200.0)
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 90.0)

    def test_new_piping_loads_do_not_include_the_existing_load(self):
        g = _build()
        pipe_graph.apply_existing_conditions(g, "1-1/4", 60.0, 200.0)
        self.assertAlmostEqual(g.edges[100].cumulative_load_mbh, 180.0)
        self.assertAlmostEqual(g.nodes[ORIGIN].cumulative_load_mbh, 180.0)

    def test_no_fixture_reached_stays_at_zero_so_sizing_stops(self):
        g = pipe_graph.NetworkGraph()
        g.origin_id = ORIGIN
        _node(g, ORIGIN, "meter")
        _node(g, 2, "tee")
        _edge(g, 100, ORIGIN, 2, 10.0)
        _finish(g)
        pipe_graph.apply_existing_conditions(g, "1", 60.0, 100.0)
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 0.0)
        self.assertRaises(ValueError, _size, g)


class TestSizesUseTheCombinedLength(unittest.TestCase):

    def test_without_existing_the_first_pipe_is_three_quarter(self):
        # 180 CFH at 30 ft: 3/4 in carries 199
        self.assertEqual(gas_tables.get_capacity(TABLE, 30, "3/4"), 199)
        result = _size(_build())
        self.assertEqual(result["sizes"][100], "3/4")
        self.assertAlmostEqual(result["longest_run_ft"], 30.0)

    def test_with_existing_length_the_first_pipe_grows(self):
        # 60 ft existing + 30 ft new = 90 ft. 3/4 in carries 110, 1 in 207.
        self.assertEqual(gas_tables.get_capacity(TABLE, 90, "3/4"), 110)
        self.assertEqual(gas_tables.get_capacity(TABLE, 90, "1"), 207)
        g = _build()
        pipe_graph.apply_existing_conditions(g, "1-1/4", 60.0, 200.0)
        result = _size(g)
        self.assertAlmostEqual(result["longest_run_ft"], 90.0)
        self.assertEqual(result["sizes"][100], "1")       # 180 CFH
        self.assertEqual(result["sizes"][101], "1")       # 120 CFH > 110
        self.assertEqual(result["sizes"][102], "3/4")     # 60 CFH


class TestExistingPipeCheck(unittest.TestCase):

    def _check(self, size, existing_load):
        g = _build()
        pipe_graph.apply_existing_conditions(g, size, 60.0, existing_load)
        return sizing_engine.check_existing_pipe(_size(g), g)

    def test_none_when_the_run_started_at_a_meter(self):
        g = _build()
        self.assertIsNone(sizing_engine.check_existing_pipe(_size(g), g))
        self.assertEqual(sizing_engine.format_existing_check(None), [])

    def test_adequate(self):
        # 1-1/4 in at 90 ft carries 424. 200 existing + 180 new = 380.
        self.assertEqual(gas_tables.get_capacity(TABLE, 90, "1-1/4"), 424)
        c = self._check("1-1/4", 200.0)
        self.assertTrue(c["adequate"])
        self.assertAlmostEqual(c["new_load_mbh"], 180.0)
        self.assertAlmostEqual(c["total_load_mbh"], 380.0)
        self.assertAlmostEqual(c["demand_cfh"], 380.0)
        self.assertAlmostEqual(c["capacity_cfh"], 424.0)
        self.assertAlmostEqual(c["run_ft"], 90.0)

    def test_not_adequate(self):
        # 300 existing + 180 new = 480, over the 424 the pipe carries
        c = self._check("1-1/4", 300.0)
        self.assertFalse(c["adequate"])
        text = "\n".join(sizing_engine.format_existing_check(c))
        text.encode("ascii")
        self.assertIn("NOT ADEQUATE", text)
        self.assertIn("WAS NOT RESIZED", text)
        self.assertIn("engineer shall verify", text)

    def test_capacity_exactly_equal_passes(self):
        c = self._check("1-1/4", 424.0 - 180.0)
        self.assertTrue(c["adequate"])

    def test_heat_content_is_applied_to_the_check(self):
        g = _build()
        pipe_graph.apply_existing_conditions(g, "1-1/4", 60.0, 200.0)
        result = sizing_engine.size_system(
            g, "Schedule 40 Steel", 0.5, TABLE, heat_content_btu_per_cf=840.0)
        c = sizing_engine.check_existing_pipe(result, g)
        # 380 MBH / 840 BTU per CF = 452.4 CFH, over 424
        self.assertAlmostEqual(c["demand_cfh"], 380000.0 / 840.0)
        self.assertFalse(c["adequate"])

    def test_a_size_that_is_not_in_the_table_is_not_passed(self):
        c = self._check("EHD-25", 10.0)
        self.assertIsNone(c["capacity_cfh"])
        self.assertFalse(c["adequate"])
        self.assertIn("NOT CHECKED",
                      "\n".join(sizing_engine.format_existing_check(c)))


class TestStartAtANewPipe(unittest.TestCase):
    """The stand-in origin. The far element fetch needs Revit, so only the
    bookkeeping is checked here."""

    def _upstream(self, **kw):
        base = {"status": "ok", "connector_index": 0, "reason": "it is open",
                "upstream_xyz": [1.0, 2.0, 3.0], "upstream_neighbour_id": 55,
                "downstream_neighbour_ids": [77]}
        base.update(kw)
        return base

    def test_origin_is_a_stand_in_at_the_upstream_end(self):
        g = pipe_graph.NetworkGraph()
        visited, queue = pipe_graph._start_at_new_pipe(
            g, None, _Stub(900), self._upstream())
        self.assertEqual(g.origin_id, -900)
        self.assertTrue(g.origin_is_tie_in)
        origin = g.nodes[-900]
        self.assertEqual(origin.node_type, "meter")
        self.assertIsNone(origin.element)
        self.assertEqual(origin.location_xyz, [1.0, 2.0, 3.0])
        self.assertFalse(origin.is_elbow)
        # the picked pipe is the first edge, from the connection outward
        self.assertEqual(g.edges[900].from_node_id, -900)
        self.assertEqual(g.edges[900].to_node_id, 77)
        # the existing side is closed to the walk
        self.assertIn(55, visited)
        self.assertIn(900, visited)

    def test_unresolved_upstream_is_refused(self):
        g = pipe_graph.NetworkGraph()
        self.assertRaises(
            ValueError, pipe_graph._start_at_new_pipe, g, None, _Stub(900),
            {"status": "ambiguous", "reason": "both ends are open"})

    def test_a_tap_on_the_picked_pipe_is_reported_not_dropped(self):
        g = pipe_graph.NetworkGraph()
        pipe_graph._start_at_new_pipe(
            g, None, _Stub(900),
            self._upstream(downstream_neighbour_ids=[77, 88]))
        self.assertIn(88, g.disconnected)

    def test_nothing_downstream_is_reported(self):
        g = pipe_graph.NetworkGraph()
        pipe_graph._start_at_new_pipe(
            g, None, _Stub(900), self._upstream(downstream_neighbour_ids=[]))
        self.assertIn(900, g.disconnected)
        self.assertIsNone(g.edges[900].to_node_id)


class TestWalkStopsAtExistingPhase(unittest.TestCase):
    """A run that started at a new pipe does not follow existing phase
    piping. A run from a meter still does, so Size Gas can audit it."""

    def _traverse(self, tie_in_run, element_phase):
        import revit_helpers
        g = pipe_graph.NetworkGraph()
        g.origin_id = ORIGIN
        g.origin_is_tie_in = tie_in_run
        g.origin_phase_id = 1
        _node(g, ORIGIN, "meter")
        processed = []

        keep = (revit_helpers.get_created_phase_id, pipe_graph._is_pipe,
                pipe_graph._process_pipe)
        revit_helpers.get_created_phase_id = lambda el: element_phase
        pipe_graph._is_pipe = lambda el: True
        pipe_graph._process_pipe = (
            lambda graph, doc, pipe, parent, visited, queue:
            processed.append(pipe.Id.Value))

        class _Typed(_Stub):
            def GetType(self):
                return type("T", (), {"Name": "Pipe"})()

        try:
            pipe_graph._traverse(g, None, set([ORIGIN]),
                                 [(_Typed(700), ORIGIN)])
        finally:
            (revit_helpers.get_created_phase_id, pipe_graph._is_pipe,
             pipe_graph._process_pipe) = keep
        return g, processed

    def test_existing_phase_element_is_not_followed(self):
        g, processed = self._traverse(True, 2)
        self.assertEqual(processed, [])
        self.assertEqual(g.existing_connections, [700])

    def test_new_phase_element_is_followed(self):
        g, processed = self._traverse(True, 1)
        self.assertEqual(processed, [700])
        self.assertEqual(g.existing_connections, [])

    def test_a_meter_run_still_walks_existing_piping(self):
        g, processed = self._traverse(False, 2)
        self.assertEqual(processed, [700])
        self.assertEqual(g.existing_connections, [])


class TestRememberedBetweenTools(unittest.TestCase):

    def test_what_size_gas_entered_comes_back_for_one_line(self):
        entered = {"size_nominal": "2", "length_ft": 85.0, "load_mbh": 640.0}
        tie_in.remember_existing(900, entered)
        self.assertEqual(tie_in.recall_existing(900), entered)
        self.assertIsNone(tie_in.recall_existing(901))
        # a copy, so a caller cannot corrupt what is stored
        tie_in.recall_existing(900)["load_mbh"] = 0.0
        self.assertAlmostEqual(tie_in.recall_existing(900)["load_mbh"], 640.0)


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
