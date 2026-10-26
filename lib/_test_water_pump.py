# -*- coding: ascii -*-
"""Tests for water_pump: recirculation pump flow and head on the loop and
branch network from _test_water_recirc.

Plain CPython. Proves the circuit walk and the arithmetic, not the traversal
against a real model.

Run:  python _test_water_pump.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_recirc import build_loop_and_branch, RETURN_TYPE  # noqa: E402

import water_graph      # noqa: E402
import water_recirc     # noqa: E402
import water_pressure   # noqa: E402
import water_pump       # noqa: E402


def network(size_in=0.75):
    g = build_loop_and_branch()
    for node in g.nodes.values():
        if node.kind == water_graph.KIND_PIPE:
            node.diameter_inches = size_in
    water_graph.apply_return_system_types(g, [RETURN_TYPE])
    return g, water_recirc.analyze(g)


class TestDutyPoint(unittest.TestCase):

    def test_flow_is_the_heater_total(self):
        g, recirc = network()
        points = water_pump.duty_points(g, recirc)
        self.assertEqual(len(points), 1)
        self.assertAlmostEqual(points[0]["gpm"], recirc["root_gpm"][1])

    def test_head_includes_setter_and_safety(self):
        g, recirc = network()
        p = water_pump.duty_points(g, recirc, circuit_setter_psi=1.0,
                                   safety_factor=0.10)[0]
        index = [c for c in p["circuits"] if c["number"] == p["index_leg"]][0]
        self.assertAlmostEqual(index["total_psi"],
                               index["supply_psi"] + index["return_psi"] +
                               index["fitting_psi"] + 1.0, places=9)
        self.assertAlmostEqual(p["head_ft"], index["total_psi"] * 2.31 * 1.1,
                               places=9)
        self.assertGreater(index["return_psi"], 0.0)

    def test_index_is_the_worst_circuit(self):
        g, recirc = network()
        p = water_pump.duty_points(g, recirc)[0]
        worst = max(c["total_psi"] for c in p["circuits"])
        self.assertEqual(p["circuits"][0]["total_psi"], worst)
        self.assertEqual(p["circuits"][0]["number"], p["index_leg"])

    def test_return_friction_by_hand(self):
        # The LOOP leg's return: 60 + 90 + 10 ft of 3/4 in at its flows.
        g, recirc = network()
        p = water_pump.duty_points(g, recirc)[0]
        loop = [c for c in p["circuits"] if c["method"] == "LOOP"][0]
        expected = 0.0
        for pid in (20, 22, 24):
            rate = water_pressure.friction_psi_per_100ft(
                recirc["pipe_gpm"][pid], 0.785)
            expected += g.nodes[pid].length_feet * rate / 100.0
        self.assertAlmostEqual(loop["return_psi"], expected, places=9)

    def test_smaller_pipe_more_head(self):
        g1, r1 = network(0.75)
        g2, r2 = network(0.5)
        self.assertGreater(water_pump.duty_points(g2, r2)[0]["head_ft"],
                           water_pump.duty_points(g1, r1)[0]["head_ft"])

    def test_no_return_no_pump(self):
        self.assertEqual(water_pump.duty_points(None, None), [])

    def test_rows_shape(self):
        g, recirc = network()
        rows = water_pump.table_rows(water_pump.duty_points(g, recirc))
        self.assertEqual(len(rows[0]), len(water_pump.TABLE_COLUMNS))


class TestHeadChart(unittest.TestCase):

    def test_chart_rows_add_up_to_the_total(self):
        g, recirc = network()
        point = water_pump.duty_points(g, recirc)[0]
        rows, equations = water_pump.head_rows(point)
        self.assertEqual(rows[0][0], "Supply pipe friction")
        self.assertEqual(rows[-1][0], "PUMP HEAD")
        self.assertEqual(rows[-2][0], "TOTAL")
        parts = sum(float(r[4].split()[0]) for r in rows[:-2])
        self.assertAlmostEqual(parts, float(rows[-2][4].split()[0]), delta=0.02)
        self.assertEqual(rows[-1][4], "{:.1f} ft".format(point["head_ft"]))
        self.assertEqual(len(equations), 8)
        self.assertTrue(equations[0].startswith("1. V"))
        self.assertIn("{:.1f} ft".format(point["head_ft"]), equations[-1])
        self.assertIn("psi/100 ft", rows[0][3])
        self.assertEqual(len(rows[0]), len(water_pump.HEAD_COLUMNS))

    def test_flow_line_uses_the_sigma(self):
        g, recirc = network()
        point = water_pump.duty_points(g, recirc)[0]
        line = water_pump.flow_line(point)
        self.assertIn(water_pump.SIGMA, line)
        self.assertIn("{:.1f} gpm".format(point["gpm"]), line)

    def test_point_names_the_heater(self):
        g, recirc = network()
        point = water_pump.duty_points(g, recirc)[0]
        self.assertIn("heater_id", point)

    def test_heater_found_through_a_long_cold_run(self):
        g, _recirc = network()
        # A chain of cold fittings between the return end and the heater.
        import _test_water_graph as t
        for nid in (901, 902, 903):
            t.add_node(g, nid, water_graph.KIND_FITTING)
        g.nodes[901].connectors = [
            {"system_type": t.COLD, "connected_element_id": 902}]
        g.nodes[902].connectors = [
            {"system_type": t.COLD, "connected_element_id": 901},
            {"system_type": t.COLD, "connected_element_id": 903}]
        g.nodes[903].connectors = [
            {"system_type": t.COLD, "connected_element_id": 902},
            {"system_type": t.COLD, "connected_element_id": 1}]
        self.assertEqual(water_pump._heater_for(g, 901), 1)

    def test_nearest_rate_skips_non_pipes(self):
        rates = {5: 1.5, 9: 3.0}
        path = [4, 6, 7, 5, 8, 9]
        self.assertEqual(water_pump._nearest_rate(path, rates, 4), 1.5)
        self.assertEqual(water_pump._nearest_rate(path, rates, 0), 1.5)
        self.assertEqual(water_pump._nearest_rate([1, 2], {}, 0), 0.0)


if __name__ == "__main__":
    unittest.main()
