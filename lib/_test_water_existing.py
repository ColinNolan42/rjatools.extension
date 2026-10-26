# -*- coding: ascii -*-
# CPython tests for existing phase pipes left alone and the upstream RPZ
# service stub (Colin, 2026-10-26). Revit is stubbed out as in the other
# _test_water_*.py files.

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, COLD)

import water_graph           # noqa: E402
import water_sizing_engine   # noqa: E402
import water_performance     # noqa: E402


def build(existing_drawn=0.5):
    """RPZ -> service stub (dead end) and RPZ -> main -> existing branch -> WC.

    Node 2 is the service stub behind the RPZ inlet. Node 4 is an EXISTING
    pipe carrying 5 WSFU, which the table sizes at 3/4, drawn smaller.
    """
    g = make_graph()
    g.origin_id = 1
    add_node(g, 1, water_graph.KIND_ORIGIN)
    add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=1.0)
    add_node(g, 3, water_graph.KIND_PIPE, parent=1, length_feet=10.0)
    existing = add_node(g, 4, water_graph.KIND_PIPE, parent=3,
                        length_feet=6.0, is_existing=True)
    existing.diameter_inches = existing_drawn
    wc = add_node(g, 5, water_graph.KIND_FIXTURE, parent=4,
                  cw=5.0, hw=0.0, total=5.0, fixture_name="WC-1")
    wc.type_name = "Water Closet, Flush Tank"
    wc.is_public = True
    g.nodes[3].diameter_inches = 1.0
    g.nodes[2].diameter_inches = 1.0
    return g


class TestUpstreamStub(unittest.TestCase):

    def test_stub_with_no_fixture_is_marked_and_not_sized(self):
        g = build()
        water_graph._mark_upstream_service(g, COLD)
        self.assertEqual(g.upstream_ids, set([2]))
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        ids = [s.element_id for s in result["segments"]]
        self.assertNotIn(2, ids)
        self.assertIn(3, ids)

    def test_open_end_on_the_stub_is_not_reported(self):
        g = build()
        g.open_ends.append((2, COLD))
        water_graph._mark_upstream_service(g, COLD)
        self.assertEqual(g.open_ends, [])


class TestExistingPipesLeftAlone(unittest.TestCase):

    def setUp(self):
        self.g = build(existing_drawn=0.5)
        water_graph._mark_upstream_service(self.g, COLD)
        water_graph._assign_demand(self.g)
        self.result = water_sizing_engine.size_network(self.g)
        self.by_id = dict((s.element_id, s) for s in self.result["segments"])

    def test_existing_pipe_keeps_its_drawn_size(self):
        seg = self.by_id[4]
        self.assertTrue(seg.keep_existing)
        self.assertAlmostEqual(seg.size_inches, 0.5)
        self.assertEqual(seg.nominal_size, "1/2")

    def test_required_size_is_kept_and_upsize_flagged(self):
        seg = self.by_id[4]
        self.assertEqual(seg.required_nominal, "3/4")
        self.assertTrue(seg.upsize_needed)
        self.assertTrue(any("UPSIZE" in f for f in self.result["flags"]))

    def test_new_pipe_is_not_held(self):
        self.assertFalse(self.by_id[3].keep_existing)

    def test_adequate_existing_pipe_is_not_flagged(self):
        g = build(existing_drawn=1.0)
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        seg = [s for s in result["segments"] if s.element_id == 4][0]
        self.assertTrue(seg.keep_existing)
        self.assertFalse(seg.upsize_needed)

    def test_performance_judges_existing_against_the_required_size(self):
        judged = water_performance.classify(self.by_id[4], "cold")
        self.assertEqual(judged.label, water_performance.RED)


class TestTransitionIsNotAnElbow(unittest.TestCase):

    def test_zero_length_transition_reads_as_a_coupling(self):
        import water_pressure

        class Stub(object):
            family_name = "Transition - Generic"
            connector_count = 2
            location_xyz = (0.0, 0.0, 0.0)
            connectors = [
                {"connected_element_id": 1, "origin_xyz": (0.0, 0.0, 0.0),
                 "diameter_in": 0.75},
                {"connected_element_id": 2, "origin_xyz": (0.0, 0.0, 0.0),
                 "diameter_in": 0.5}]

        kind, _size, note = water_pressure.classify_fitting(Stub(), 1, 2)
        self.assertEqual(kind, "coupling")
        self.assertEqual(note, "")


class TestCompactReport(unittest.TestCase):

    def test_grouped_flags_and_size_summary(self):
        import water_report
        flags = ["Pipe %d: no fixture demand downstream" % i
                 for i in range(1, 24)] + ["EXISTING pipe 9 is 1/2: UPSIZE"]
        lines = water_report.group_flags(flags)
        self.assertEqual(len(lines), 2)
        self.assertTrue(lines[0].startswith("23 pipe(s): no fixture demand"))
        self.assertIn("+18 more", lines[0])

        g = build()
        water_graph._mark_upstream_service(g, COLD)
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        header = {"date": "d", "job": "J", "job_number": "1", "by": "b"}
        text = water_report.build_compact_report(g, result, header)
        self.assertIn("PIPES SIZED", text)
        self.assertIn("existing phase pipe(s) left", text)
        self.assertNotIn("PIPE SIZING BY SEGMENT", text)

    def test_recirc_equation_is_shown_with_numbers(self):
        import water_report
        recirc = {"heat_loss_btu_per_100ft": 5000, "delta_t_f": 20,
                  "min_gpm": 0.5, "total_gpm": 0.5, "return_length_ft": 94.9,
                  "unrooted_pipe_ids": [], "root_gpm": {},
                  "legs": [{"number": 1, "method": "LOOP",
                            "effective_length_ft": 94.9, "calc_gpm": 0.4745,
                            "gpm": 0.5}]}
        text = chr(10).join(water_report.recirc_compact(recirc))
        self.assertIn("(94.9 / 100 x 5000) / (500 x 20) = 0.47 gpm", text)


if __name__ == "__main__":
    unittest.main()
