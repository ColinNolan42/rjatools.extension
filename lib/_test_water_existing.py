# -*- coding: ascii -*-
# CPython tests for existing phase pipes left alone and the upstream RPZ
# service stub (Colin, 2026-10-26). Revit is stubbed out as in the other
# _test_water_*.py files.

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, link, COLD)

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

    def test_oversized_existing_branch_does_not_raise_the_new_main(self):
        # Bench 5 (2026-10-29): a 2 in existing branch to one fixture raised
        # the new main above it to 2 in.
        g = build(existing_drawn=2.0)
        water_graph._mark_upstream_service(g, COLD)
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        by_id = dict((s.element_id, s) for s in result["segments"])
        self.assertEqual(by_id[4].nominal_size, "2")
        self.assertEqual(by_id[3].nominal_size, "3/4")

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

class TestHotColdTie(unittest.TestCase):

    def _graph(self, type_id):
        g = build()
        node = g.nodes[3]
        node.system_type_id = type_id
        g.cross_ties.append((1, 3, water_graph.shared_params.
                             SYSTEM_DOMESTIC_HOT_WATER, COLD))
        return g

    def test_hot_pipe_on_the_cold_run_is_flagged(self):
        import water_checks
        result = water_checks.check_system(self._graph(None))
        self.assertTrue(any(f.code == "hot_cold_tie"
                            for f in result["findings"]))

    def test_the_return_tie_is_not_flagged(self):
        import water_checks
        result = water_checks.check_system(self._graph(77), [77])
        self.assertFalse(any(f.code == "hot_cold_tie"
                             for f in result["findings"]))
class TestUnloadedHeater(unittest.TestCase):

    def _graph(self, heater_serves):
        g = make_graph()
        g.origin_id = 1
        hot = water_graph.shared_params.SYSTEM_DOMESTIC_HOT_WATER
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=5.0)
        wc = add_node(g, 3, water_graph.KIND_FIXTURE, parent=2,
                      cw=5.0, hw=0.0, total=5.0)
        wc.type_name = "Water Closet, Flush Tank"
        wc.is_public = True
        add_node(g, 4, water_graph.KIND_PIPE, parent=1, length_feet=4.0)
        add_node(g, 5, water_graph.KIND_HEATER, parent=4,
                 served=list(heater_serves))
        add_node(g, 6, water_graph.KIND_PIPE, system=hot, length_feet=1.0)
        link(g, 5, 6, hot)
        g.heater_ids.append(5)
        for nid in (2, 4, 6):
            g.nodes[nid].diameter_inches = 1.0
        return g

    def test_cold_feed_and_hot_stub_are_not_sized(self):
        g = self._graph([])
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        by_id = dict((s.element_id, s) for s in result["segments"])
        self.assertIsNone(by_id[4].nominal_size)
        self.assertIsNone(by_id[6].nominal_size)
        self.assertIn("serves no hot fixtures", by_id[4].flag)
        self.assertIsNotNone(by_id[2].nominal_size)

    def test_parallel_heater_is_not_reported_as_serving_nothing(self):
        # 4C (2026-10-29): two heaters feed one hot header; the walk reaches
        # the fixtures through heater 5 only.
        import water_checks
        hot = water_graph.shared_params.SYSTEM_DOMESTIC_HOT_WATER
        g = self._graph([3])
        add_node(g, 7, water_graph.KIND_PIPE, parent=1, length_feet=4.0)
        add_node(g, 8, water_graph.KIND_HEATER, parent=7, served=[])
        add_node(g, 10, water_graph.KIND_PIPE, system=hot, length_feet=1.0)
        link(g, 8, 10, hot)
        g.heater_ids.append(8)
        g.hot_joints = {5: [6], 6: [5, 10], 10: [6, 8], 8: [10]}
        result = water_checks.check_system(g)
        self.assertFalse(any(f.code == "heater_serves_nothing"
                             for f in result["findings"]))

    def test_a_loaded_heater_is_sized_as_before(self):
        g = self._graph([3])
        water_graph._assign_demand(g)
        result = water_sizing_engine.size_network(g)
        by_id = dict((s.element_id, s) for s in result["segments"])
        self.assertFalse(by_id[4].flag and "serves no hot" in by_id[4].flag)

class TestParallelBackflow(unittest.TestCase):

    def _graph(self):
        """Pick RPZ B (1). Its outlet pipe 2 feeds a tee 3: the main (4, one
        WC) and a dead end leg 5, 7 that ends at RPZ A (8)."""
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=3.0)
        add_node(g, 3, water_graph.KIND_FITTING, parent=2)
        add_node(g, 4, water_graph.KIND_PIPE, parent=3, length_feet=20.0)
        wc = add_node(g, 9, water_graph.KIND_FIXTURE, parent=4,
                      cw=5.0, hw=0.0, total=5.0)
        wc.type_name = "Water Closet, Flush Tank"
        wc.is_public = True
        add_node(g, 5, water_graph.KIND_PIPE, parent=3, length_feet=2.0)
        add_node(g, 6, water_graph.KIND_FITTING, parent=5)
        add_node(g, 7, water_graph.KIND_PIPE, parent=6, length_feet=2.0)
        rpz = add_node(g, 8, water_graph.KIND_ACCESSORY, parent=7)
        rpz.family_name = "RPZ 1-INCH"
        return g

    def test_the_other_rpz_leg_takes_the_full_demand(self):
        g = self._graph()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[5].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[7].demand_wsfu, 5.0)
        self.assertEqual(g.parallel_backflow, [8])

    def test_leg_that_ends_in_a_loop_joint_to_the_other_rpz(self):
        # The walk reached RPZ A from its inlet side (upstream), so the outlet
        # leg 5, 7 is a dead end with a loop joint back to RPZ A.
        g = self._graph()
        g.nodes[7].children_by_system[COLD] = []
        add_node(g, 20, water_graph.KIND_PIPE, parent=1)
        g.nodes[8].parent_by_system[COLD] = 20
        g.nodes[20].children_by_system[COLD] = [8]
        g.upstream_ids = set([20, 8])
        g.loop_edges.append((7, 8, COLD))
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[5].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[7].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[20].demand_wsfu, 0.0)
        self.assertEqual(g.parallel_backflow, [8])

    def test_other_rpz_reached_from_its_inlet_is_not_service_piping(self):
        # Bench 4 and 4C (2026-10-29): pick RPZ B (1). Its inlet pipe 30 runs
        # to the service tee 31 (stub 32), then pipe 33 into RPZ A (34), whose
        # outlet leg 35, 36 meets the main tee 3 at a loop joint.
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=3.0)
        add_node(g, 3, water_graph.KIND_FITTING, parent=2)
        add_node(g, 4, water_graph.KIND_PIPE, parent=3, length_feet=20.0)
        wc = add_node(g, 9, water_graph.KIND_FIXTURE, parent=4,
                      cw=5.0, hw=0.0, total=5.0)
        wc.type_name = "Water Closet, Flush Tank"
        wc.is_public = True
        add_node(g, 30, water_graph.KIND_PIPE, parent=1, length_feet=1.0)
        add_node(g, 31, water_graph.KIND_FITTING, parent=30)
        add_node(g, 32, water_graph.KIND_PIPE, parent=31, length_feet=1.0)
        add_node(g, 33, water_graph.KIND_PIPE, parent=31, length_feet=1.0)
        rpz_a = add_node(g, 34, water_graph.KIND_ACCESSORY, parent=33)
        rpz_a.family_name = "RPZ 1-INCH"
        add_node(g, 35, water_graph.KIND_PIPE, parent=34, length_feet=2.0)
        add_node(g, 36, water_graph.KIND_PIPE, parent=35, length_feet=2.0)
        g.loop_edges.append((36, 3, COLD))
        water_graph._mark_upstream_service(g, COLD)
        self.assertEqual(g.upstream_ids, set([30, 31, 32, 33, 34]))
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[35].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[36].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[32].demand_wsfu, 0.0)
        self.assertEqual(g.parallel_backflow, [34])

    def _live_4c(self):
        """Bench 4C as walked live (2026-10-29): each RPZ outlet goes through
        a transition before its first pipe, and the other RPZ's leg hangs off
        the main tee and ends in a loop joint at that RPZ's transition.

        Pick RPZ A (1): transition 2, leg 3, elbow 4, pipe 5, main tee 6 with
        the main 7 (one WC) and RPZ B's leg 10, transition 11. Service: pipe
        20, tee 21, stub 22, pipe 23, RPZ B 24."""
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_FITTING, parent=1)
        add_node(g, 3, water_graph.KIND_PIPE, parent=2, length_feet=2.0)
        add_node(g, 4, water_graph.KIND_FITTING, parent=3)
        add_node(g, 5, water_graph.KIND_PIPE, parent=4, length_feet=2.0)
        add_node(g, 6, water_graph.KIND_FITTING, parent=5)
        add_node(g, 7, water_graph.KIND_PIPE, parent=6, length_feet=20.0)
        wc = add_node(g, 9, water_graph.KIND_FIXTURE, parent=7,
                      cw=5.0, hw=0.0, total=5.0)
        wc.type_name = "Water Closet, Flush Tank"
        wc.is_public = True
        add_node(g, 10, water_graph.KIND_PIPE, parent=6, length_feet=2.0)
        add_node(g, 11, water_graph.KIND_FITTING, parent=10)
        add_node(g, 20, water_graph.KIND_PIPE, parent=1, length_feet=1.0)
        add_node(g, 21, water_graph.KIND_FITTING, parent=20)
        add_node(g, 22, water_graph.KIND_PIPE, parent=21, length_feet=1.0)
        add_node(g, 23, water_graph.KIND_PIPE, parent=21, length_feet=1.0)
        rpz_b = add_node(g, 24, water_graph.KIND_ACCESSORY, parent=23)
        rpz_b.family_name = "RPZ 1-INCH"
        g.loop_edges += [(24, 11, COLD), (11, 24, COLD)]
        return g

    def test_live_4c_leg_under_the_main_tee_takes_the_full_demand(self):
        g = self._live_4c()
        water_graph._mark_upstream_service(g, COLD)
        self.assertEqual(g.upstream_ids, set([20, 21, 22, 23, 24]))
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[10].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[22].demand_wsfu, 0.0)
        self.assertEqual(g.parallel_backflow, [24])

    def test_live_4c_two_pipe_leg(self):
        # The pick B mirror: the other leg is two pipes and an elbow.
        g = self._live_4c()
        add_node(g, 12, water_graph.KIND_FITTING, parent=11)
        add_node(g, 13, water_graph.KIND_PIPE, parent=12, length_feet=2.0)
        add_node(g, 14, water_graph.KIND_FITTING, parent=13)
        g.loop_edges = [(24, 14, COLD), (14, 24, COLD)]
        water_graph._mark_upstream_service(g, COLD)
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[10].demand_wsfu, 5.0)
        self.assertAlmostEqual(g.nodes[13].demand_wsfu, 5.0)
        self.assertEqual(g.parallel_backflow, [24])

    def test_no_second_rpz_changes_nothing(self):
        g = self._graph()
        g.nodes[8].family_name = "BALL VALVE"
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[5].demand_wsfu, 0.0)
        self.assertEqual(g.parallel_backflow, [])


class TestHoseBibs(unittest.TestCase):

    def test_hose_bibb_is_five_cold_wsfu(self):
        # Colin, 2026-10-29: 5 cold WSFU each, not a fixed 5 gpm.
        import water_tables
        self.assertEqual(water_tables.HOSE_BIBB_WSFU, 5.0)

    def test_two_hose_bibs_size_on_wsfu(self):
        import water_tables
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=10.0)
        for nid in (3, 4):
            add_node(g, nid, water_graph.KIND_FIXTURE, parent=2,
                     cw=water_tables.HOSE_BIBB_WSFU, hw=0.0,
                     total=water_tables.HOSE_BIBB_WSFU)
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 10.0)
        self.assertAlmostEqual(g.nodes[2].demand_fixed_gpm, 0.0)

    def test_name_recognition(self):
        f = water_graph._name_is_hose_bib
        self.assertTrue(f("OUTDOOR HOSE BIB", "65C-8"))
        self.assertTrue(f("Wall Hydrant", "Sill Cock 1/2"))
        self.assertTrue(f("RJA - Hose Bibb", ""))
        self.assertFalse(f("RJA- Water Fixture Family", "Lavatory"))

    def _graph(self):
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1, length_feet=10.0)
        add_node(g, 3, water_graph.KIND_FIXTURE, parent=2, fixed_gpm=5.0)
        add_node(g, 4, water_graph.KIND_FIXTURE, parent=2, fixed_gpm=5.0)
        for nid in (3, 4):
            g.nodes[nid].fixture_name = "Hose bibb"
            g.nodes[nid].type_name = "65C-8"
        for nid in (2,):
            g.nodes[nid].diameter_inches = 1.0
        return g

    def test_pipe_to_hose_bibs_is_sized_on_the_fixed_gpm(self):
        g = self._graph()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[2].demand_fixed_gpm, 10.0)
        result = water_sizing_engine.size_network(g)
        seg = result["segments"][0]
        self.assertAlmostEqual(seg.demand_gpm, 10.0)
        self.assertEqual(seg.nominal_size, "3/4")
        self.assertIn("gpm fixed", seg.rule)
        judged = water_performance.classify(seg, "cold")
        self.assertNotEqual(judged.reason, "no fixture demand downstream")

    def test_hose_bibs_are_not_a_zero_wsfu_error_and_show_in_the_take_off(self):
        import water_checks
        import water_report
        g = self._graph()
        water_graph._assign_demand(g)
        checks = water_checks.check_system(g)
        self.assertFalse(any(f.code == "fixture_zero_wsfu"
                             for f in checks["findings"]))
        data = water_report.wsfu_data(g)
        self.assertEqual(data["ext_fixed_gpm"], 10.0)
        self.assertIn("Hose bibb", data["rows"][0]["fixture"])
        self.assertEqual(data["rows"][0]["qty"], "2")
        self.assertIn("plus 10 gpm fixed",
                      "\n".join(water_report.wsfu_load_lines(data)))

    def test_unrecognized_equipment_is_flagged(self):
        import water_system_errors
        g = self._graph()
        add_node(g, 7, water_graph.KIND_EQUIPMENT, parent=2)
        g.nodes[7].family_name = "Mystery Tank"
        g.unrecognized_ids.append(7)
        out = water_system_errors.collect(
            g, {"segments": []}, None, {"findings": []})
        self.assertTrue(any("not recognized" in t and "Mystery Tank" in t
                            for _s, t in out))


class TestExistingBranchProbe(unittest.TestCase):

    def setUp(self):
        import revit_helpers
        self.rh = revit_helpers
        self.saved = (water_graph._get_element, revit_helpers.get_connectors,
                      revit_helpers.get_created_phase_id)
        self.world = {}
        water_graph._get_element = lambda doc, nid: self.world.get(nid)
        revit_helpers.get_connectors = lambda el: el["conns"]
        revit_helpers.get_created_phase_id = lambda el: el["phase"]
        self.g = make_graph()
        self.g.origin_phase_id = 10

    def tearDown(self):
        (water_graph._get_element, self.rh.get_connectors,
         self.rh.get_created_phase_id) = self.saved

    def put(self, nid, phase, *neighbours):
        self.world[nid] = {"phase": phase, "conns": [
            {"system_type": COLD, "is_connected": True,
             "connected_element_id": n} for n in neighbours]}

    def test_a_branch_to_existing_fixtures_is_walkable(self):
        self.put(2, 5, 1, 3)
        self.put(3, 5, 2)            # the existing fixture, dead end
        members = water_graph._probe_existing_branch(
            self.g, None, 2, 1, COLD, set([1]))
        self.assertEqual(members, set([2, 3]))

    def test_a_branch_that_loops_back_is_not_walked(self):
        self.put(2, 5, 1, 3)
        self.put(3, 5, 2, 9)         # 9 was already visited
        self.assertIsNone(water_graph._probe_existing_branch(
            self.g, None, 2, 1, COLD, set([1, 9])))

    def test_a_branch_that_meets_new_work_is_not_walked(self):
        self.put(2, 5, 1, 3)
        self.put(3, 10, 2)           # new phase element, not yet visited
        self.assertIsNone(water_graph._probe_existing_branch(
            self.g, None, 2, 1, COLD, set([1])))

    def test_a_long_existing_run_is_the_main_not_a_branch(self):
        self.put(2, 5, 1, 3)
        for nid in range(3, 140):
            self.put(nid, 5, nid - 1, nid + 1)
        self.put(140, 5, 139)
        self.assertIsNone(water_graph._probe_existing_branch(
            self.g, None, 2, 1, COLD, set([1])))


class TestExistingPipeRow(unittest.TestCase):

    def test_existing_pipe_reason_says_existing(self):
        import water_flagging
        r = water_performance.PipeResult(2)
        r.label = water_performance.RED
        r.existing = True
        text = water_flagging.short_reason(r, [])
        self.assertEqual(text, "Existing pipe, undersized")


if __name__ == "__main__":
    unittest.main()
