# -*- coding: ascii -*-
"""Tests for hot water recirculation: the supply / return split, the circuit
setter flows and the return pipe sizes.

Runs in plain CPython on synthetic graphs. This proves the arithmetic and the
tree logic only. It does NOT prove the traversal against a real model, and the
pump pass-through in water_graph._walk is not exercised here at all, because
it needs live Revit connectors.

Run:  python _test_water_recirc.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (          # noqa: E402
    _install_revit_stubs, make_graph, add_node, link, COLD, HOT)

import water_graph          # noqa: E402
import water_tables         # noqa: E402
import water_recirc         # noqa: E402
import water_sizing_engine  # noqa: E402
import water_report         # noqa: E402

SUPPLY_TYPE = 100
RETURN_TYPE = 200
RET = water_graph.TREE_RETURN


def hot_pipe(g, nid, parent, length, type_id=SUPPLY_TYPE):
    node = add_node(g, nid, water_graph.KIND_PIPE, parent=parent, system=HOT,
                    length_feet=length)
    node.system_type_id = type_id
    node.system_type_name = ("Domestic Hot Water Recirculation"
                             if type_id == RETURN_TYPE else "Domestic Hot Water")
    g.system_types[type_id] = node.system_type_name
    return node


def build_loop_and_branch():
    """One loop and one branch return on the same heater.

        heater(1) -> main(10, 30 ft) -> tee(11) -> branch(12, 8 ft) -> lav(13)
                                        tee(11) -> main(14, 40 ft) -> tee(15)
                                                   tee(15) -> stub(16, 5 ft) -> lav(17)

        LOOP return:   tee(15) -> ret(20, 60 ft) -> tee(21) -> ret(22, 90 ft)
                       -> pump(23) -> ret(24, 10 ft) -> heater(1)
        BRANCH return: lav(13) -> ret(30, 12 ft) -> tee(21)

    The tree is laid out the way the hot walk really leaves it on a loop: the
    return front from the heater got to tee(15) before the supply front did,
    so tee(15), stub(16) and lav(17) hang off the RETURN side, and the joint
    main(14)-tee(15) is only a recorded loop edge. That is the case
    apply_return_system_types exists to fix.
    """
    g = make_graph()
    add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
    g.heater_ids.append(1)

    # supply front
    hot_pipe(g, 10, 1, 30.0)
    add_node(g, 11, water_graph.KIND_FITTING, parent=10, system=HOT)
    hot_pipe(g, 12, 11, 8.0)
    add_node(g, 13, water_graph.KIND_FIXTURE, parent=12, system=HOT,
             cw=1.5, hw=1.5, total=2.0, fixture_name="L-1")
    hot_pipe(g, 14, 11, 40.0)

    # return front, walked from the heater's return connection
    hot_pipe(g, 24, 1, 10.0, RETURN_TYPE)
    add_node(g, 23, water_graph.KIND_PUMP, parent=24, system=HOT)
    g.pump_ids.append(23)
    hot_pipe(g, 22, 23, 90.0, RETURN_TYPE)
    add_node(g, 21, water_graph.KIND_FITTING, parent=22, system=HOT)
    hot_pipe(g, 20, 21, 60.0, RETURN_TYPE)
    hot_pipe(g, 30, 21, 12.0, RETURN_TYPE)
    add_node(g, 15, water_graph.KIND_FITTING, parent=20, system=HOT)
    hot_pipe(g, 16, 15, 5.0)
    add_node(g, 17, water_graph.KIND_FIXTURE, parent=16, system=HOT,
             cw=1.5, hw=1.5, total=2.0, fixture_name="L-2")

    for fid in (13, 17):
        g.nodes[fid].type_name = "Lavatory"
        g.nodes[fid].is_public = True

    # the two joints the spanning tree dropped
    g.loop_edges.append((14, 15, HOT))
    g.loop_edges.append((30, 13, HOT))
    return g


class TestHeatLossFlow(unittest.TestCase):
    """The firm tab: BTU/h = L/100 x 5000, gpm = BTU/h / (500 x dT)."""

    def test_100_ft_at_defaults_is_half_a_gpm(self):
        # 100/100 x 5000 = 5000 BTU/h. 5000 / (500 x 20) = 0.50 gpm.
        self.assertAlmostEqual(water_recirc.heat_loss_gpm(100, 5000, 20), 0.5)

    def test_150_ft_loop(self):
        # 150/100 x 5000 = 7500 BTU/h. 7500 / 10000 = 0.75 gpm.
        self.assertAlmostEqual(water_recirc.heat_loss_gpm(150, 5000, 20), 0.75)

    def test_the_firm_tab_example(self):
        # The template's own numbers: 20 ft -> 1000 BTU/h -> 0.1 gpm.
        self.assertAlmostEqual(water_recirc.heat_loss_gpm(20, 5000, 20), 0.1)

    def test_a_tighter_temperature_drop_needs_more_flow(self):
        self.assertAlmostEqual(water_recirc.heat_loss_gpm(100, 5000, 10), 1.0)


class TestReturnSizeTable(unittest.TestCase):
    """The hot water table (5 fps) restated in gpm by Table E103.3(3)."""

    def test_rows_are_the_hot_limits_converted_to_gpm(self):
        rows = dict((r["nominal_size"], r) for r in water_tables.return_gpm_rows())
        # hot WSFU limit -> flush tank gpm, read off Table E103.3(3)
        for size, wsfu, gpm in (("1/2", 1, 3.0), ("3/4", 3, 6.5),
                                ("1", 7, 11.8), ("1-1/4", 18, 18.8),
                                ("1-1/2", 40, 26.3), ("2", 120, 48.0),
                                ("2-1/2", 250, 75.0), ("3", 400, 105.0)):
            self.assertAlmostEqual(rows[size]["max_wsfu"], wsfu)
            self.assertAlmostEqual(rows[size]["max_gpm"], gpm)

    def test_minimum_return_is_half_inch(self):
        self.assertEqual(water_tables.select_return_size(0.5)["nominal_size"],
                         "1/2")
        self.assertEqual(water_tables.return_gpm_rows()[0]["size_inches"], 0.5)

    def test_limits_are_inclusive(self):
        self.assertEqual(water_tables.select_return_size(3.0)["nominal_size"],
                         "1/2")
        self.assertEqual(water_tables.select_return_size(3.01)["nominal_size"],
                         "3/4")
        self.assertEqual(water_tables.select_return_size(6.5)["nominal_size"],
                         "3/4")
        self.assertEqual(water_tables.select_return_size(6.6)["nominal_size"],
                         "1")

    def test_no_flow_and_over_the_table(self):
        self.assertEqual(water_tables.select_return_size(0)["status"], "no_flow")
        over = water_tables.select_return_size(500)
        self.assertEqual(over["status"], "exceeds_table")
        self.assertIsNone(over["nominal_size"])
        self.assertAlmostEqual(over["max_table_gpm"], 105.0)

    def test_negative_flow_is_rejected(self):
        self.assertRaises(ValueError, water_tables.select_return_size, -1)


class TestSupplyReturnSplit(unittest.TestCase):

    def test_the_raw_walk_undercounts_the_supply_main(self):
        """Documents the fault: before the split, lav(17) hangs off the
        return side and the main never sees it."""
        g = build_loop_and_branch()
        water_graph._assign_demand(g)
        self.assertAlmostEqual(g.nodes[10].demand_wsfu, 1.5)

    def test_split_puts_every_hot_fixture_on_the_supply(self):
        g = build_loop_and_branch()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        self.assertAlmostEqual(g.nodes[10].demand_wsfu, 3.0)   # both lavs
        self.assertAlmostEqual(g.nodes[14].demand_wsfu, 1.5)   # the far lav
        self.assertAlmostEqual(g.nodes[12].demand_wsfu, 1.5)
        self.assertEqual(sorted(g.nodes[1].served_fixture_ids), [13, 17])

    def test_return_pipes_leave_the_supply_tree(self):
        g = build_loop_and_branch()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        below_heater = g.descendants(1, HOT)
        for pid in (20, 22, 24, 30):
            self.assertNotIn(pid, below_heater)
            self.assertNotIn(HOT, g.nodes[pid].parent_by_system)

    def test_return_tree_is_rooted_at_the_heater(self):
        g = build_loop_and_branch()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        self.assertEqual(g.return_roots, [1])
        self.assertEqual(g.return_unrooted, [])
        # 24 is the last pipe before the heater, everything drains through it
        self.assertEqual(g.nodes[24].parent_by_system[RET], 1)
        below = g.descendants(24, RET)
        for pid in (22, 20, 30):
            self.assertIn(pid, below)

    def test_no_return_type_changes_nothing(self):
        g = build_loop_and_branch()
        before = dict((nid, dict(n.parent_by_system))
                      for nid, n in g.nodes.items())
        self.assertIsNone(water_graph.apply_return_system_types(g, []))
        after = dict((nid, dict(n.parent_by_system))
                     for nid, n in g.nodes.items())
        self.assertEqual(before, after)

    def test_split_can_be_applied_twice(self):
        g = build_loop_and_branch()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        self.assertAlmostEqual(g.nodes[10].demand_wsfu, 3.0)
        self.assertEqual(g.return_roots, [1])


class TestCircuitSetterFlows(unittest.TestCase):

    def setUp(self):
        self.g = build_loop_and_branch()
        water_graph.apply_return_system_types(self.g, [RETURN_TYPE])
        self.r = water_recirc.analyze(self.g)
        self.legs = dict((l["start_pipe_id"], l) for l in self.r["legs"])

    def test_actual_drop_is_at_the_set_flow(self):
        """Loop: 160 ft x 50 BTU/h/ft = 8000 BTU/h, calc 0.80 gpm at 20 F,
        so the calculation governs and the actual drop is the design 20 F.
        Branch: 8 ft = 400 BTU/h, calc 0.04 gpm, set to the 0.5 gpm minimum,
        so the actual drop is 400 / (500 x 0.5) = 1.6 F."""
        for leg in self.r["legs"]:
            self.assertAlmostEqual(
                leg["heat_loss_btuh"],
                leg["effective_length_ft"] / 100.0 * self.r["heat_loss_btu_per_100ft"])
            self.assertAlmostEqual(
                leg["actual_delta_t_f"],
                leg["heat_loss_btuh"] / (500.0 * leg["gpm"]))
            if leg["gpm"] > leg["calc_gpm"]:
                self.assertLess(leg["actual_delta_t_f"], self.r["delta_t_f"])
            else:
                self.assertAlmostEqual(leg["actual_delta_t_f"], self.r["delta_t_f"])

    def test_total_return_length_is_every_return_pipe(self):
        # 60 + 90 + 10 + 12, the four return pipes in build_loop_and_branch.
        self.assertAlmostEqual(self.r["return_length_ft"], 172.0)

    def test_drafting_rows_carry_the_drops_and_the_total(self):
        data = water_report.recirc_data(self.r)
        self.assertEqual(len(data["rows"]), len(self.r["legs"]))
        for row in data["rows"]:
            for key, _header, _w in water_report.RECIRC_COLUMNS:
                self.assertTrue(row[key])
        self.assertIn("172.0 ft", " ".join(data["summary"]))

    def test_one_circuit_setter_per_return_leg(self):
        self.assertEqual(sorted(self.legs.keys()), [20, 30])

    def test_loop_length_is_last_fixture_back_to_the_heater(self):
        loop = self.legs[20]
        self.assertEqual(loop["method"], water_recirc.METHOD_LOOP)
        # 60 + 90 + 10 = 160 ft of return
        self.assertAlmostEqual(loop["effective_length_ft"], 160.0)
        # 160/100 x 5000 = 8000 BTU/h. 8000 / (500 x 20) = 0.80 gpm.
        self.assertAlmostEqual(loop["calc_gpm"], 0.8)
        self.assertAlmostEqual(loop["gpm"], 0.8)
        self.assertEqual(loop["start_element_id"], 15)

    def test_branch_length_is_main_to_fixture_and_takes_the_minimum(self):
        branch = self.legs[30]
        self.assertEqual(branch["method"], water_recirc.METHOD_BRANCH)
        # only pipe 12 (8 ft) lies between the main and the lavatory
        self.assertAlmostEqual(branch["effective_length_ft"], 8.0)
        self.assertAlmostEqual(branch["calc_gpm"], 0.04)
        self.assertAlmostEqual(branch["gpm"], 0.5)
        self.assertEqual(branch["start_element_id"], 13)

    def test_flows_add_at_the_junction_toward_the_heater(self):
        flow = self.r["pipe_gpm"]
        self.assertAlmostEqual(flow[20], 0.8)     # loop only
        self.assertAlmostEqual(flow[30], 0.5)     # branch only
        self.assertAlmostEqual(flow[22], 1.3)     # after they join
        self.assertAlmostEqual(flow[24], 1.3)     # at the heater

    def test_total_is_the_pump_flow(self):
        self.assertAlmostEqual(self.r["total_gpm"], 1.3)
        self.assertAlmostEqual(self.r["root_gpm"][1], 1.3)

    def test_loop_is_listed_first(self):
        self.assertEqual(self.r["legs"][0]["method"], water_recirc.METHOD_LOOP)
        self.assertEqual([l["number"] for l in self.r["legs"]], [1, 2])

    def test_inputs_change_the_loop_but_not_the_minimum(self):
        r = water_recirc.analyze(self.g, heat_loss_btu_per_100ft=5000,
                                 delta_t_f=10)
        legs = dict((l["start_pipe_id"], l) for l in r["legs"])
        self.assertAlmostEqual(legs[20]["gpm"], 1.6)
        self.assertAlmostEqual(legs[30]["gpm"], 0.5)

    def test_short_loop_takes_the_minimum(self):
        # 10 + 10 + 10 = 30 ft of loop return. The branch return is kept
        # shorter still, so the loop stays the longest run.
        for pid in (20, 22, 24):
            self.g.nodes[pid].length_feet = 10.0
        self.g.nodes[30].length_feet = 4.0
        r = water_recirc.analyze(self.g)
        legs = dict((l["start_pipe_id"], l) for l in r["legs"])
        self.assertAlmostEqual(legs[20]["calc_gpm"], 0.15)
        self.assertAlmostEqual(legs[20]["gpm"], 0.5)


class TestASecondLongWingIsItsOwnLoop(unittest.TestCase):
    """Colin, 2026-10-13: a second long recirculation wing counts as its own
    loop, and then the flows combine."""

    def setUp(self):
        self.g = build_loop_and_branch()
        # Main loop: 300 + 90 + 10 = 400 ft back to the heater.
        self.g.nodes[20].length_feet = 300.0
        # Second wing: 120 ft of its own return before it joins at tee 21.
        self.g.nodes[30].length_feet = 120.0
        water_graph.apply_return_system_types(self.g, [RETURN_TYPE])
        self.r = water_recirc.analyze(self.g)
        self.legs = dict((l["start_pipe_id"], l) for l in self.r["legs"])

    def test_main_loop_is_measured_back_to_the_heater(self):
        main = self.legs[20]
        self.assertEqual(main["method"], water_recirc.METHOD_LOOP)
        self.assertAlmostEqual(main["effective_length_ft"], 400.0)
        # 400/100 x 5000 / (500 x 20) = 2.00 gpm
        self.assertAlmostEqual(main["gpm"], 2.0)

    def test_wing_is_a_loop_measured_to_where_it_joins(self):
        wing = self.legs[30]
        self.assertEqual(wing["method"], water_recirc.METHOD_LOOP)
        self.assertAlmostEqual(wing["effective_length_ft"], 120.0)
        # 120/100 x 5000 / (500 x 20) = 0.60 gpm, over the 0.5 minimum
        self.assertAlmostEqual(wing["gpm"], 0.6)

    def test_flows_combine_where_the_wing_joins(self):
        flow = self.r["pipe_gpm"]
        self.assertAlmostEqual(flow[20], 2.0)
        self.assertAlmostEqual(flow[30], 0.6)
        self.assertAlmostEqual(flow[22], 2.6)
        self.assertAlmostEqual(flow[24], 2.6)
        self.assertAlmostEqual(self.r["total_gpm"], 2.6)

    def test_a_short_leg_is_still_a_branch_at_the_minimum(self):
        # 100 ft of own return is exactly the minimum, so it is not a loop
        self.g.nodes[30].length_feet = 100.0
        r = water_recirc.analyze(self.g)
        leg = dict((l["start_pipe_id"], l) for l in r["legs"])[30]
        self.assertEqual(leg["method"], water_recirc.METHOD_BRANCH)
        self.assertAlmostEqual(leg["gpm"], 0.5)


class TestReturnSizing(unittest.TestCase):

    def setUp(self):
        self.g = build_loop_and_branch()
        water_graph.apply_return_system_types(self.g, [RETURN_TYPE])
        self.recirc = water_recirc.analyze(self.g)
        self.result = water_sizing_engine.size_network(
            self.g, return_system_type_ids=[RETURN_TYPE], recirc=self.recirc)
        self.by_id = dict((s.element_id, s) for s in self.result["segments"])

    def test_every_return_pipe_is_sized_on_its_flow(self):
        for pid, gpm in ((20, 0.8), (30, 0.5), (22, 1.3), (24, 1.3)):
            seg = self.by_id[pid]
            self.assertTrue(seg.is_return)
            self.assertAlmostEqual(seg.circ_gpm, gpm)
            # all under the 3.0 gpm the table allows in 1/2 in
            self.assertEqual(seg.nominal_size, "1/2")
            self.assertIn("circulation", seg.rule)

    def test_supply_is_still_sized_on_fixture_units(self):
        main = self.by_id[10]
        self.assertFalse(main.is_return)
        self.assertAlmostEqual(main.demand_wsfu, 3.0)
        self.assertEqual(main.nominal_size, "3/4")     # hot: 1 < 3 <= 3

    def test_a_big_system_steps_the_return_up_toward_the_heater(self):
        # 2000 ft of loop return: 20/100... 2000/100 x 5000 / 10000 = 10 gpm
        self.g.nodes[20].length_feet = 1900.0
        recirc = water_recirc.analyze(self.g)
        result = water_sizing_engine.size_network(
            self.g, return_system_type_ids=[RETURN_TYPE], recirc=recirc)
        by_id = dict((s.element_id, s) for s in result["segments"])
        # loop leg 1900 + 90 + 10 = 2000 ft -> 10.0 gpm -> 1 in (6.5 < 10 <= 11.8)
        self.assertAlmostEqual(by_id[20].circ_gpm, 10.0)
        self.assertEqual(by_id[20].nominal_size, "1")
        # plus the 0.5 gpm branch after the junction: 10.5 gpm, still 1 in
        self.assertAlmostEqual(by_id[24].circ_gpm, 10.5)
        self.assertEqual(by_id[24].nominal_size, "1")
        self.assertEqual(by_id[30].nominal_size, "1/2")
        self.assertGreaterEqual(by_id[24].size_inches, by_id[20].size_inches)

    def test_without_an_analysis_the_return_is_left_unsized(self):
        result = water_sizing_engine.size_network(
            self.g, return_system_type_ids=[RETURN_TYPE])
        by_id = dict((s.element_id, s) for s in result["segments"])
        self.assertIsNone(by_id[20].nominal_size)
        self.assertEqual(by_id[20].rule, "NOT SIZED")

    def test_report_lists_the_circuit_setters(self):
        header = {"date": "2026-10-12", "job": "Demo", "job_number": "26.000",
                  "by": "CN"}
        text = water_report.build_report(self.g, self.result, header,
                                         self.recirc)
        text.encode("ascii")
        for needle in ("HOT WATER RECIRCULATION - CIRCUIT SETTER FLOWS",
                       "LOOP", "BRANCH", "160.0", "0.80", "0.50",
                       "TOTAL CIRCULATION FLOW, 2 circuit setter(s): 1.30 gpm",
                       "HWR",
                       "4 hot water return pipe(s)"):
            self.assertIn(needle, text)
        self.assertNotIn("still to do", text)

    def test_report_without_a_return_has_no_recirculation_block(self):
        header = {"date": "2026-10-12", "job": "", "job_number": "", "by": ""}
        text = water_report.build_report(self.g, self.result, header)
        self.assertNotIn("CIRCUIT SETTER FLOWS", text)


class TestReturnTiedIntoTheColdInlet(unittest.TestCase):
    """The return comes back to a tee on the heater's cold inlet, not to a
    connector on the heater, so it never meets the heater on the hot side."""

    def build(self):
        g = make_graph()
        g.origin_id = 1
        add_node(g, 1, water_graph.KIND_ORIGIN)
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)               # cold
        add_node(g, 3, water_graph.KIND_FITTING, parent=2)            # the tie-in tee
        add_node(g, 4, water_graph.KIND_PIPE, parent=3)               # cold
        add_node(g, 5, water_graph.KIND_HEATER, parent=4)
        g.heater_ids.append(5)

        hot_pipe(g, 10, 5, 50.0)
        add_node(g, 11, water_graph.KIND_FIXTURE, parent=10, system=HOT,
                 cw=1.5, hw=1.5, total=2.0)
        hot_pipe(g, 20, 11, 120.0, RETURN_TYPE)
        add_node(g, 21, water_graph.KIND_PUMP, parent=20, system=HOT)
        g.pump_ids.append(21)
        hot_pipe(g, 22, 21, 6.0, RETURN_TYPE)
        link(g, 22, 3, HOT)      # the hot walk reaches the cold tee from the return
        return g

    def test_return_ends_at_the_cold_tie_in(self):
        g = self.build()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        self.assertEqual(g.return_roots, [3])
        self.assertEqual(g.return_unrooted, [])
        r = water_recirc.analyze(g)
        self.assertEqual(len(r["legs"]), 1)
        leg = r["legs"][0]
        self.assertEqual(leg["method"], water_recirc.METHOD_LOOP)
        self.assertAlmostEqual(leg["effective_length_ft"], 126.0)
        # 126/100 x 5000 / 10000 = 0.63 gpm
        self.assertAlmostEqual(leg["gpm"], 0.63)
        self.assertAlmostEqual(r["pipe_gpm"][22], 0.63)

    def test_cold_demand_is_untouched(self):
        g = self.build()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        # the cold main feeds only the heater, which serves the one lavatory
        self.assertAlmostEqual(g.nodes[2].demand_wsfu, 1.5)


class TestReturnWithNoEnd(unittest.TestCase):

    def test_a_return_that_reaches_nothing_is_flagged_not_guessed(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        g.heater_ids.append(1)
        hot_pipe(g, 10, 1, 50.0)
        add_node(g, 11, water_graph.KIND_FIXTURE, parent=10, system=HOT,
                 cw=1.5, hw=1.5, total=2.0)
        hot_pipe(g, 20, 11, 80.0, RETURN_TYPE)     # runs off and stops

        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        self.assertEqual(g.return_roots, [])
        self.assertEqual(g.return_unrooted, [20])

        recirc = water_recirc.analyze(g)
        self.assertEqual(recirc["legs"], [])
        self.assertEqual(recirc["unrooted_pipe_ids"], [20])

        result = water_sizing_engine.size_network(
            g, return_system_type_ids=[RETURN_TYPE], recirc=recirc)
        seg = [s for s in result["segments"] if s.element_id == 20][0]
        self.assertIsNone(seg.nominal_size)
        self.assertIn("could not find where this return line ends", seg.flag)
        self.assertTrue(any("Pipe 20" in f for f in result["flags"]))
        # the supply still sizes
        supply = [s for s in result["segments"] if s.element_id == 10][0]
        self.assertEqual(supply.nominal_size, "3/4")


class TestAGuessIsNeverActedOn(unittest.TestCase):
    """The split and the return sizing change real pipe sizes, so they only
    run on a return that something confirmed."""

    def build_two_supply_types(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        g.heater_ids.append(1)
        for nid in range(10, 20):
            hot_pipe(g, nid, 1, 10.0, 100)
        hot_pipe(g, 30, 1, 10.0, 300)
        g.system_types[100] = "HW A"
        g.system_types[300] = "HW B"
        for nid in range(10, 20):
            g.nodes[nid].system_type_name = "HW A"
        g.nodes[30].system_type_name = "HW B"
        add_node(g, 31, water_graph.KIND_FIXTURE, parent=30, system=HOT,
                 cw=1.5, hw=1.5, total=2.0)
        return g

    def test_pipe_count_alone_is_reported_as_unconfirmed(self):
        g = self.build_two_supply_types()
        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([300]))
        self.assertEqual(found["unconfirmed"], set([300]))

    def test_unconfirmed_return_is_left_unsized_and_flagged(self):
        g = self.build_two_supply_types()
        found = water_graph.detect_return_system_types(g)
        water_graph.apply_return_system_types(
            g, found["detected"] - found["unconfirmed"])
        water_graph._assign_demand(g)
        self.assertEqual(g.return_system_type_ids, set())
        recirc = water_recirc.analyze(g)
        result = water_sizing_engine.size_network(
            g, return_system_type_ids=found["detected"], recirc=recirc)
        seg = [s for s in result["segments"] if s.element_id == 30][0]
        self.assertIsNone(seg.nominal_size)
        self.assertIn("nothing confirmed it", seg.flag)
        self.assertTrue(any("Pipe 30" in f for f in result["flags"]))

    def test_a_confirmed_return_has_no_unconfirmed_entry(self):
        g = build_loop_and_branch()
        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([RETURN_TYPE]))
        self.assertEqual(found["unconfirmed"], set())
        self.assertTrue(found["certain"])

    def test_split_refused_when_every_hot_pipe_is_called_a_return(self):
        g = build_loop_and_branch()
        self.assertIsNone(water_graph.apply_return_system_types(
            g, [SUPPLY_TYPE, RETURN_TYPE]))
        self.assertEqual(g.return_system_type_ids, set())
        self.assertIn(HOT, g.nodes[10].parent_by_system)   # tree untouched

    def test_pump_beside_the_heater_does_not_see_through_it(self):
        """pump - fitting - heater - supply pipe. The pump must not reach the
        supply main on the far side of the heater and call it the return."""
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        g.heater_ids.append(1)
        hot_pipe(g, 10, 1, 30.0)                                # supply
        add_node(g, 2, water_graph.KIND_FITTING, parent=1, system=HOT)
        add_node(g, 3, water_graph.KIND_PUMP, parent=2, system=HOT)
        g.pump_ids.append(3)
        hot_pipe(g, 20, 3, 80.0, RETURN_TYPE)
        found = water_graph.detect_return_system_types(g)
        self.assertEqual(found["detected"], set([RETURN_TYPE]))

    def test_bad_inputs_are_rejected(self):
        g = build_loop_and_branch()
        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        self.assertRaises(ValueError, water_recirc.analyze, g, 5000, 0)
        self.assertRaises(ValueError, water_recirc.analyze, g, 0, 20)


class TestTwoLoopsIntoOneHeater(unittest.TestCase):

    def test_each_line_into_the_heater_has_its_own_loop(self):
        g = make_graph()
        add_node(g, 1, water_graph.KIND_HEATER, system=HOT)
        g.heater_ids.append(1)
        for base in (10, 50):
            hot_pipe(g, base, 1, 40.0)
            add_node(g, base + 1, water_graph.KIND_FIXTURE, parent=base,
                     system=HOT, cw=1.5, hw=1.5, total=2.0)
            hot_pipe(g, base + 2, base + 1, 200.0, RETURN_TYPE)
            g.loop_edges.append((base + 2, 1, HOT))

        water_graph.apply_return_system_types(g, [RETURN_TYPE])
        r = water_recirc.analyze(g)
        self.assertEqual([l["method"] for l in r["legs"]],
                         [water_recirc.METHOD_LOOP, water_recirc.METHOD_LOOP])
        # 200 ft each -> 1.0 gpm each, 2.0 gpm at the pump
        self.assertAlmostEqual(r["total_gpm"], 2.0)


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
