# -*- coding: ascii -*-
"""Tests for the pre-sizing completeness checks.

Pure logic, Revit stubbed. Proves the arithmetic and the rules only, not
behaviour against a real model.

Run:  python _test_water_checks.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _test_water_graph import (      # noqa: E402
    _install_revit_stubs, make_graph, add_node, link, COLD, HOT)

import shared_params        # noqa: E402
import water_graph          # noqa: E402
import water_checks         # noqa: E402


def conn(system, connected=True):
    return {"connector_index": 0, "direction": "In", "is_connected": connected,
            "connected_element_id": 99 if connected else None,
            "connected_element_type": None, "origin_xyz": None,
            "connected_system_element_id": None,
            "system_type": system, "diameter_in": 1.0}


def codes(result):
    return set(f.code for f in result["findings"])


def build(with_heater=True, with_pump=False, hwr_fixtures=0,
          hot_fixture=True, origin_conns=None):
    """A minimal RPZ -> main -> fixture system, with optional heater and pump."""
    g = make_graph()
    g.origin_id = 1
    origin = add_node(g, 1, water_graph.KIND_ORIGIN)
    origin.connectors = origin_conns if origin_conns is not None else [
        conn(COLD), conn(COLD)]

    add_node(g, 2, water_graph.KIND_PIPE, parent=1)
    fixture = add_node(g, 3, water_graph.KIND_FIXTURE, parent=2,
                       cw=1.5, hw=1.5 if hot_fixture else 0.0, total=2.0,
                       fixture_name="L-1")
    fixture.type_name = "Lavatory"

    if with_heater:
        add_node(g, 4, water_graph.KIND_PIPE, parent=2)
        heater = add_node(g, 5, water_graph.KIND_HEATER, parent=4,
                          served=[3] if hot_fixture else [])
        heater.family_name = "WH-1"
        g.heater_ids.append(5)
        hot_pipe = add_node(g, 6, water_graph.KIND_PIPE, parent=5, system=HOT)
        hot_pipe.system_type_id = 1822
        hot_pipe.system_type_name = "Domestic Hot Water"
        if hot_fixture:
            link(g, 6, 3, HOT)

    if with_pump:
        pump = add_node(g, 7, water_graph.KIND_PUMP, parent=6 if with_heater else 2,
                        system=HOT)
        pump.family_name = "HWCP"
        g.pump_ids.append(7)

    for i in range(hwr_fixtures):
        node = add_node(g, 20 + i, water_graph.KIND_FIXTURE, parent=2,
                        cw=1.5, hw=1.5, total=2.0)
        node.type_name = "Lavatory"
        node.hwr_active = True
        if with_heater:
            g.nodes[5].served_fixture_ids.append(20 + i)
            link(g, 6, 20 + i, HOT)
    return g


class TestCompleteSystem(unittest.TestCase):

    def test_rpz_heater_and_pump_with_hwr_is_ready(self):
        g = build(with_heater=True, with_pump=True, hwr_fixtures=1)
        r = water_checks.check_system(g, return_system_type_ids=[66256])
        self.assertTrue(r["ready_for_sizing"])
        self.assertEqual(r["errors"], [])
        self.assertIn("recirc_complete", codes(r))
        self.assertIn("hot_water_available", codes(r))
        self.assertTrue(r["profile"]["has_hot"])
        self.assertTrue(r["profile"]["has_recirc"])

    def test_cold_only_system_is_ready(self):
        g = build(with_heater=False, hot_fixture=False)
        r = water_checks.check_system(g)
        self.assertTrue(r["ready_for_sizing"])
        self.assertIn("cold_only_system", codes(r))
        self.assertFalse(r["profile"]["has_hot"])


class TestMissingHeater(unittest.TestCase):

    def test_hot_demand_without_heater_is_an_error(self):
        g = build(with_heater=False, hot_fixture=True)
        r = water_checks.check_system(g)
        self.assertFalse(r["ready_for_sizing"])
        self.assertIn("hot_demand_no_heater", codes(r))

    def test_heater_serving_nothing_warns(self):
        g = build(with_heater=True, hot_fixture=False)
        r = water_checks.check_system(g)
        self.assertIn("heater_serves_nothing", codes(r))
        # a warning, not an error: a cold-only job can still be sized
        self.assertTrue(r["ready_for_sizing"])


class TestRecirculation(unittest.TestCase):

    def test_hwr_fixtures_without_pump_warns(self):
        g = build(with_heater=True, with_pump=False, hwr_fixtures=2)
        r = water_checks.check_system(g, return_system_type_ids=[66256])
        self.assertIn("hwr_wanted_no_pump", codes(r))
        self.assertTrue(r["ready_for_sizing"])   # warning only

    def test_pump_without_hwr_fixtures_warns(self):
        g = build(with_heater=True, with_pump=True, hwr_fixtures=0)
        r = water_checks.check_system(g, return_system_type_ids=[66256])
        self.assertIn("pump_no_hwr_fixtures", codes(r))

    def test_hwr_fixtures_without_return_system_warns(self):
        g = build(with_heater=True, with_pump=True, hwr_fixtures=1)
        r = water_checks.check_system(g, return_system_type_ids=[])
        self.assertIn("hwr_wanted_no_return_system", codes(r))
        self.assertFalse(r["profile"]["has_recirc"])

    def test_return_system_without_pump_warns(self):
        g = build(with_heater=True, with_pump=False, hwr_fixtures=1)
        r = water_checks.check_system(g, return_system_type_ids=[66256])
        self.assertIn("return_system_no_pump", codes(r))


class TestOrigin(unittest.TestCase):

    def test_origin_with_no_cold_connector_is_an_error(self):
        g = build(origin_conns=[conn(HOT)])
        r = water_checks.check_system(g)
        self.assertFalse(r["ready_for_sizing"])
        self.assertIn("origin_not_on_cold_water", codes(r))

    def test_single_connector_origin_warns_not_inline(self):
        g = build(origin_conns=[conn(COLD)])
        r = water_checks.check_system(g)
        self.assertIn("origin_not_inline", codes(r))

    def test_origin_with_nothing_connected_is_an_error(self):
        g = build(origin_conns=[conn(COLD, connected=False),
                                conn(COLD, connected=False)])
        r = water_checks.check_system(g)
        self.assertFalse(r["ready_for_sizing"])
        self.assertIn("origin_disconnected", codes(r))


class TestFixtureData(unittest.TestCase):

    def test_no_fixtures_is_an_error(self):
        g = make_graph()
        g.origin_id = 1
        origin = add_node(g, 1, water_graph.KIND_ORIGIN)
        origin.connectors = [conn(COLD), conn(COLD)]
        add_node(g, 2, water_graph.KIND_PIPE, parent=1)
        r = water_checks.check_system(g)
        self.assertFalse(r["ready_for_sizing"])
        self.assertIn("no_fixtures", codes(r))

    def test_zero_wsfu_fixture_is_an_error_not_zero_demand(self):
        g = build()
        node = add_node(g, 30, water_graph.KIND_FIXTURE, parent=2,
                        cw=0.0, hw=0.0, total=0.0, fixture_name="BAD-1")
        node.type_name = "Lavatory"
        r = water_checks.check_system(g)
        self.assertFalse(r["ready_for_sizing"])
        self.assertIn("fixture_zero_wsfu", codes(r))


class TestReportBlock(unittest.TestCase):

    def test_format_is_ascii_and_states_readiness(self):
        g = build(with_heater=True, with_pump=True, hwr_fixtures=1)
        r = water_checks.check_system(g, return_system_type_ids=[66256])
        text = "\n".join(water_checks.format_checks(r))
        text.encode("ascii")
        self.assertIn("SYSTEM COMPLETENESS CHECK", text)
        self.assertIn("READY FOR SIZING: YES", text)

    def test_failing_system_says_no(self):
        g = build(with_heater=False, hot_fixture=True)
        r = water_checks.check_system(g)
        text = "\n".join(water_checks.format_checks(r))
        self.assertIn("READY FOR SIZING: NO", text)


def _run():
    suite = unittest.TestLoader().loadTestsFromModule(sys.modules[__name__])
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    print("=" * 70)
    print("SUMMARY: {} run, {} passed, {} failed, {} errors".format(
        result.testsRun,
        result.testsRun - len(result.failures) - len(result.errors),
        len(result.failures), len(result.errors)))
    print("=" * 70)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    sys.exit(_run())
