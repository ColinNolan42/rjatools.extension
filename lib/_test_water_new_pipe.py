# -*- coding: ascii -*-
"""Tests for starting Size Water at the first new cold water pipe on an
existing building, instead of at an RPZ.

Runs in plain CPython on synthetic data. It proves the upstream / downstream
decision and that a pipe origin is sized like any other pipe. It does NOT
prove the phase read (Element.CreatedPhaseId) or the cap test against a real
model. Those need a live Revit run.

Run:  python _test_water_new_pipe.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, link, COLD, HOT)

import tie_in               # noqa: E402
import water_graph          # noqa: E402
import water_checks         # noqa: E402
import water_sizing_engine  # noqa: E402
import water_report         # noqa: E402

OPEN = tie_in.END_OPEN
EXISTING = tie_in.END_EXISTING
CAP = tie_in.END_CAP
NEW = tie_in.END_NEW


def ends(*kinds):
    return [{"connector_index": i, "kind": k} for i, k in enumerate(kinds)]


def connector(index, connected, neighbour=None, system=COLD):
    return {"connector_index": index, "direction": "Bidirectional",
            "is_connected": connected, "connected_element_id": neighbour,
            "connected_element_type": None, "origin_xyz": None,
            "connected_system_element_id": None, "system_type": system,
            "diameter_in": 0.75}


class TestWhichEndIsUpstream(unittest.TestCase):
    """Check both sides of the new pipe: the side on existing is upstream,
    the side on new construction is downstream."""

    def test_existing_phase_side_is_upstream(self):
        r = tie_in.pick_upstream_end(ends(NEW, EXISTING))
        self.assertEqual(r["status"], "ok")
        self.assertEqual(r["connector_index"], 1)
        self.assertIn("existing phase", r["reason"])

    def test_open_end_is_upstream_when_existing_is_not_modeled(self):
        r = tie_in.pick_upstream_end(ends(OPEN, NEW))
        self.assertEqual(r["status"], "ok")
        self.assertEqual(r["connector_index"], 0)
        self.assertIn("not modeled", r["reason"])

    def test_capped_end_is_upstream(self):
        r = tie_in.pick_upstream_end(ends(NEW, CAP))
        self.assertEqual(r["status"], "ok")
        self.assertEqual(r["connector_index"], 1)

    def test_a_pipe_in_the_middle_of_new_work_is_refused(self):
        r = tie_in.pick_upstream_end(ends(NEW, NEW))
        self.assertEqual(r["status"], "no_upstream")
        self.assertNotIn("connector_index", r)

    def test_both_ends_on_existing_is_refused_not_guessed(self):
        for pair in ((OPEN, OPEN), (EXISTING, OPEN), (EXISTING, EXISTING),
                     (CAP, OPEN)):
            r = tie_in.pick_upstream_end(ends(*pair))
            self.assertEqual(r["status"], "ambiguous", pair)

    def test_a_tap_on_the_pipe_does_not_confuse_it(self):
        # two ends plus a tap connector, all on new work except the tie-in
        r = tie_in.pick_upstream_end(ends(EXISTING, NEW, NEW))
        self.assertEqual(r["status"], "ok")
        self.assertEqual(r["connector_index"], 0)


def build_from_new_pipe():
    """New work off an existing main, started at the first new pipe.

        (existing, not modeled) -- newPipe(1, 25 ft) -- tee(2)
              tee(2) -- pipe(3) -- WC(4)   public flush tank, 5.0
              tee(2) -- pipe(5) -- lav(6)  public, 1.5 / 1.5 / 2.0, cold only here
    """
    g = make_graph()
    g.origin_id = 1
    g.origin_is_new_pipe = True
    g.origin_upstream_connectors = set([0])
    g.origin_upstream_reason = "it is open, the existing piping is not modeled"
    add_node(g, 1, water_graph.KIND_PIPE, length_feet=25.0)
    add_node(g, 2, water_graph.KIND_FITTING, parent=1)
    add_node(g, 3, water_graph.KIND_PIPE, parent=2, length_feet=6.0)
    wc = add_node(g, 4, water_graph.KIND_FIXTURE, parent=3,
                  cw=5.0, hw=0.0, total=5.0, fixture_name="WC-1")
    add_node(g, 5, water_graph.KIND_PIPE, parent=2, length_feet=8.0)
    lav = add_node(g, 6, water_graph.KIND_FIXTURE, parent=5,
                   cw=1.5, hw=1.5, total=2.0, fixture_name="L-1")
    for node, name in ((wc, "Water Closet, Flush Tank"), (lav, "Lavatory")):
        node.type_name = name
        node.is_public = True
        node.connectors = [connector(0, True)]
    return g


class TestPipeOriginIsSized(unittest.TestCase):

    def setUp(self):
        self.g = build_from_new_pipe()
        water_graph._assign_demand(self.g)
        self.result = water_sizing_engine.size_network(self.g)
        self.by_id = dict((s.element_id, s) for s in self.result["segments"])

    def test_the_picked_pipe_carries_everything_downstream(self):
        # 5.0 + 2.0 = 7.0 total wsfu on the first new pipe
        self.assertAlmostEqual(self.g.nodes[1].demand_wsfu, 7.0)

    def test_the_picked_pipe_gets_a_size(self):
        # cold table: 6 < 7 <= 19 -> 1 inch
        self.assertEqual(self.by_id[1].nominal_size, "1")
        self.assertEqual(self.by_id[1].fixture_count, 2)

    def test_branches_size_as_before(self):
        self.assertEqual(self.by_id[3].nominal_size, "3/4")   # 5.0 wsfu
        # 2.0 wsfu is 1/2 by the table, raised to the 3/4 fixture connector
        self.assertEqual(self.by_id[5].nominal_size, "3/4")

    def test_report_states_the_existing_main_was_not_checked(self):
        header = {"date": "2026-10-13", "job": "TI", "job_number": "26.000",
                  "by": "CN"}
        text = water_report.build_report(self.g, self.result, header)
        text.encode("ascii")
        self.assertIn("sizing starts at the new construction connection "
                      "(pipe 1)", text)
        self.assertIn("assumed adequate for the added load and was NOT sized "
                      "or checked", text)

    def test_rpz_run_says_nothing_about_an_existing_main(self):
        self.g.origin_is_new_pipe = False
        self.assertEqual(water_report.origin_basis_lines(self.g), [])


class TestChecksOnAPipeOrigin(unittest.TestCase):

    def test_no_rpz_complaints_and_the_assumption_is_stated(self):
        g = build_from_new_pipe()
        g.nodes[1].connectors = [connector(0, False), connector(1, True, 2)]
        result = water_checks.check_system(g)
        found = set(f.code for f in result["findings"])
        self.assertIn("origin_new_pipe", found)
        for code in ("origin_not_on_cold_water", "origin_not_inline",
                     "origin_disconnected", "origin_new_pipe_no_upstream"):
            self.assertNotIn(code, found)
        message = [f.message for f in result["findings"]
                   if f.code == "origin_new_pipe"][0]
        self.assertIn("NOT checked", message)

    def test_unresolved_upstream_is_an_error(self):
        g = build_from_new_pipe()
        g.origin_upstream_connectors = set()
        result = water_checks.check_system(g)
        self.assertIn("origin_new_pipe_no_upstream",
                      set(f.code for f in result["findings"]))
        self.assertFalse(result["ready_for_sizing"])


class TestWalkLeavesTheExistingSideAlone(unittest.TestCase):
    """The end facing existing is not walked and is not an open end."""

    def _origin_only(self, upstream_connectors):
        g = make_graph()
        g.origin_id = 1
        g.origin_is_new_pipe = True
        g.origin_upstream_connectors = set(upstream_connectors)
        node = add_node(g, 1, water_graph.KIND_PIPE)
        # connector 0 faces the unmodeled existing main, connector 1 is open
        # too, standing in for an unfinished new run
        node.connectors = [connector(0, False), connector(1, False)]
        water_graph._walk(g, None, 1, COLD, set([1]))
        return g

    def test_the_existing_side_is_not_reported_as_an_open_end(self):
        g = self._origin_only([0])
        self.assertEqual(g.open_ends, [(1, COLD)])     # connector 1 only

    def test_without_the_marker_both_ends_would_be_reported(self):
        g = self._origin_only([])
        self.assertEqual(len(g.open_ends), 2)


class TestWalkStopsAtExistingPhase(unittest.TestCase):
    """A run that started at a new pipe does not follow existing phase
    piping. The connection to existing is assumed adequate."""

    def _walk_with_phase(self, neighbour_phase, origin_phase=1,
                         new_pipe=True):
        import revit_helpers
        g = make_graph()
        g.origin_id = 1
        g.origin_is_new_pipe = new_pipe
        g.origin_phase_id = origin_phase
        g.origin_upstream_connectors = set([0])
        node = add_node(g, 1, water_graph.KIND_PIPE)
        node.connectors = [connector(0, False), connector(1, True, 50)]

        keep = (water_graph._get_element, revit_helpers.get_created_phase_id,
                water_graph._make_node)
        made = []

        def fake_make(graph, element, element_id, system):
            made.append(element_id)
            return add_node(graph, element_id, water_graph.KIND_PIPE)

        water_graph._get_element = lambda doc, eid: object()
        revit_helpers.get_created_phase_id = lambda el: neighbour_phase
        water_graph._make_node = fake_make
        try:
            water_graph._walk(g, None, 1, COLD, set([1]))
        finally:
            (water_graph._get_element, revit_helpers.get_created_phase_id,
             water_graph._make_node) = keep
        return g, made

    def test_existing_phase_neighbour_is_not_followed(self):
        g, made = self._walk_with_phase(neighbour_phase=2)
        self.assertEqual(made, [])
        self.assertNotIn(50, g.nodes)
        self.assertEqual(g.existing_connections, [(1, 50)])
        self.assertEqual(g.open_ends, [])

    def test_same_phase_neighbour_is_followed(self):
        g, made = self._walk_with_phase(neighbour_phase=1)
        self.assertEqual(made, [50])
        self.assertEqual(g.existing_connections, [])

    def test_unreadable_phase_is_followed_not_guessed(self):
        g, made = self._walk_with_phase(neighbour_phase=None)
        self.assertEqual(made, [50])

    def test_an_rpz_run_still_walks_existing_piping(self):
        g, made = self._walk_with_phase(neighbour_phase=2, new_pipe=False)
        self.assertEqual(made, [50])
        self.assertEqual(g.existing_connections, [])

    def test_report_says_what_was_not_followed(self):
        g = build_from_new_pipe()
        g.existing_connections = [(3, 77)]
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        text = "\n".join(water_report.exceptions_block(g, result))
        self.assertIn("1 connection(s) to EXISTING phase piping were not "
                      "followed", text)
        self.assertIn("77", text)


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
