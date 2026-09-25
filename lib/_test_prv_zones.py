# -*- coding: ascii -*-
"""Tests for gas PRV pressure zones and zone-aware sizing.

Runs in plain CPython by stubbing the Revit assemblies. Covers pipe_graph's
mid-stream classification (the 10 ft rule), zone assignment, per-system
longest run and sizing_engine.size_system() with a downstream table. It does
NOT test traversal against a real model, family-name detection on real
families, or the WPF dialog - those need a live Revit session, and the rigor
level of any such test must be stated separately.

Run:  python _test_prv_zones.py
"""

import os
import sys
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _install_revit_stubs():
    autodesk = types.ModuleType("Autodesk")
    revit = types.ModuleType("Autodesk.Revit")
    db = types.ModuleType("Autodesk.Revit.DB")

    class _Enum(object):
        def __init__(self, name):
            self.name = name

        def __str__(self):
            return self.name

    for name in ("BuiltInParameter", "ElementId", "FilteredElementCollector",
                 "BuiltInCategory", "FlowDirectionType", "MEPSystem",
                 "SpecTypeId"):
        setattr(db, name, _Enum(name))

    autodesk.Revit = revit
    revit.DB = db
    sys.modules["Autodesk"] = autodesk
    sys.modules["Autodesk.Revit"] = revit
    sys.modules["Autodesk.Revit.DB"] = db

    pyrevit = types.ModuleType("pyrevit")

    class _HostApp(object):
        version = "2022"

    pyrevit.HOST_APP = _HostApp()
    sys.modules["pyrevit"] = pyrevit


_install_revit_stubs()

import gas_tables      # noqa: E402
import shared_params   # noqa: E402
import pipe_graph      # noqa: E402
import sizing_engine   # noqa: E402


# ---------------------------------------------------------------------------
# Graph builders that bypass Revit entirely.
# ---------------------------------------------------------------------------
class _Id(object):
    def __init__(self, value):
        self.IntegerValue = value
        self.Value = value


class _Stub(object):
    """A stand-in element: just enough for NetworkNode's constructor."""
    def __init__(self, eid):
        self.Id = _Id(eid)
        self.Location = None


def _node(graph, eid, kind, load=0.0, name="", prv=False, elbow=False):
    n = pipe_graph.NetworkNode(eid, _Stub(eid), kind)
    n.is_gas_fixture = (kind == "fixture")
    n.gas_load_mbh = load
    n.fixture_name = name
    n.is_prv = prv
    n.is_elbow = elbow
    if prv:
        graph.prv_ids.append(eid)
    graph.add_node(n)
    return n


def _edge(graph, eid, frm, to, length):
    e = object.__new__(pipe_graph.NetworkEdge)
    e.element_id = eid
    e.pipe = None
    e.from_node_id = frm
    e.to_node_id = to
    e.length_feet = length
    e.diameter_inches = 0.0
    e.cumulative_load_mbh = 0.0
    e.zone = 0
    e.run_key = None
    graph.add_edge(e)
    return e


def _finish(graph):
    pipe_graph._calculate_cumulative_loads(graph)
    pipe_graph._find_longest_run(graph)
    pipe_graph._classify_midstream_prvs(graph)
    pipe_graph._assign_pressure_zones(graph)
    pipe_graph._find_zone_runs(graph)


def _build(with_prv=True, prv_to_fixture_ft=20.0, elbow_after_prv=False):
    """meter 1 -> tee 2 -> PRV 3 -> [elbow 7] -> RTU-1 4, tee 2 -> GUH-1 5.

    prv_to_fixture_ft is the pipe between the PRV and RTU-1 (split across two
    pipes when an elbow sits in it).
    """
    g = pipe_graph.NetworkGraph()
    g.origin_id = 1
    _node(g, 1, "meter")
    _node(g, 2, "tee")
    _node(g, 3, "fitting", prv=with_prv)
    _node(g, 4, "fixture", 800.0, "RTU-1")
    _node(g, 5, "fixture", 400.0, "GUH-1")
    _edge(g, 101, 1, 2, 50.0)
    _edge(g, 102, 2, 3, 30.0)
    _edge(g, 104, 2, 5, 40.0)
    if elbow_after_prv:
        _node(g, 7, "fitting", elbow=True)
        _edge(g, 103, 3, 7, prv_to_fixture_ft / 2.0)
        _edge(g, 106, 7, 4, prv_to_fixture_ft / 2.0)
    else:
        _edge(g, 103, 3, 4, prv_to_fixture_ft)
    _finish(g)
    return g


def _build_two(second_downstream_ft):
    """meter 1 -> tee 2 -> PRV 3 -> PRV 6 -> RTU-1 4, tee 2 -> GUH-1 5.

    PRV 3 has 20 ft of pipe to PRV 6; PRV 6 has second_downstream_ft to the
    fixture.
    """
    g = pipe_graph.NetworkGraph()
    g.origin_id = 1
    _node(g, 1, "meter")
    _node(g, 2, "tee")
    _node(g, 3, "fitting", prv=True)
    _node(g, 6, "fitting", prv=True)
    _node(g, 4, "fixture", 800.0, "RTU-1")
    _node(g, 5, "fixture", 400.0, "GUH-1")
    _edge(g, 101, 1, 2, 50.0)
    _edge(g, 102, 2, 3, 30.0)
    _edge(g, 104, 2, 5, 40.0)
    _edge(g, 103, 3, 6, 20.0)
    _edge(g, 105, 6, 4, second_downstream_ft)
    _finish(g)
    return g


UP = "402.4(5)"     # Sched 40 steel, natural, 2 psi, 1 psi drop
DOWN = "402.4(2)"   # Sched 40 steel, natural, <2 psi, 0.5 in. w.c.


def _expected_size(table_id, run_ft, mbh, heat=840.0):
    """Independent re-derivation: smallest size with capacity >= CFH, 3/4 in.
    firm minimum."""
    cfh = mbh * 1000.0 / heat
    for s in gas_tables.list_pipe_sizes(table_id):
        if sizing_engine.NOMINAL_TO_INCHES.get(s, 0.0) < sizing_engine.MIN_PIPE_INCHES:
            continue
        try:
            if gas_tables.get_capacity(table_id, run_ft, s) >= cfh:
                return s
        except ValueError:
            continue
    raise AssertionError("no size fits")


class ZoneAssignment(unittest.TestCase):
    def test_zones_and_run_keys(self):
        g = _build()
        for eid in (101, 102, 104):
            self.assertEqual(g.edges[eid].zone, 0, eid)
            self.assertEqual(g.edges[eid].run_key, 1, eid)
        self.assertEqual(g.edges[103].zone, 1)
        self.assertEqual(g.edges[103].run_key, 3)
        # The PRV is still in the zone it regulates FROM.
        self.assertEqual(g.nodes[3].zone, 0)
        self.assertEqual(g.nodes[4].zone, 1)

    def test_no_prv_leaves_everything_in_zone_zero(self):
        g = _build(with_prv=False)
        self.assertEqual(g.prv_ids, [])
        self.assertEqual(g.zone_runs, {})
        self.assertTrue(all(e.zone == 0 for e in g.edges.values()))

    def test_per_system_longest_run(self):
        g = _build()
        # Whole-system run is unchanged: through the PRV to RTU-1 = 100 ft.
        self.assertAlmostEqual(g.longest_run["total_length_feet"], 100.0)
        # Meter system ends at the farthest fixture OR PRV: GUH-1 at 90 ft.
        self.assertAlmostEqual(g.zone_runs[1]["total_length_feet"], 90.0)
        self.assertEqual(g.zone_runs[1]["farthest_fixture_name"], "GUH-1")
        # The PRV's own system: 20 ft to RTU-1.
        self.assertAlmostEqual(g.zone_runs[3]["total_length_feet"], 20.0)
        self.assertEqual(g.zone_runs[3]["farthest_fixture_name"], "RTU-1")

    def test_meter_run_stops_at_prv(self):
        g = _build()
        # If the walk went THROUGH the PRV the meter run would be 100 ft.
        self.assertLess(g.zone_runs[1]["total_length_feet"],
                        g.longest_run["total_length_feet"])

    def test_nested_midstream_prv_is_flagged(self):
        g = _build_two(15.0)
        self.assertEqual(g.midstream_prv_ids, [3, 6])
        self.assertEqual(g.nested_prv_ids, [6])
        self.assertEqual(g.edges[105].zone, 2)

    def test_equipment_prv_downstream_of_midstream_is_not_nested(self):
        g = _build_two(4.0)
        self.assertEqual(g.midstream_prv_ids, [3])
        self.assertEqual(g.nested_prv_ids, [])
        # The equipment PRV starts no zone: its outgoing pipe stays in zone 1.
        self.assertEqual(g.edges[105].zone, 1)
        self.assertEqual(g.edges[105].run_key, 3)
        self.assertEqual(sorted(g.zone_runs), [1, 3])


class MidstreamClassification(unittest.TestCase):
    """A PRV is a step down only when more than 10 ft of pipe follows it."""

    def test_threshold_is_ten_feet(self):
        self.assertEqual(shared_params.PRV_MIDSTREAM_MIN_DOWNSTREAM_FT, 10.0)

    def test_regulator_at_its_equipment_is_ignored(self):
        g = _build(prv_to_fixture_ft=4.0)
        self.assertEqual(g.prv_ids, [3])
        self.assertEqual(g.midstream_prv_ids, [])
        self.assertFalse(g.nodes[3].is_midstream_prv)
        self.assertAlmostEqual(g.nodes[3].downstream_pipe_ft, 4.0)
        self.assertEqual(g.zone_runs, {})
        self.assertTrue(all(e.zone == 0 and e.run_key == 1
                            for e in g.edges.values()))

    def test_exactly_ten_feet_is_still_equipment(self):
        g = _build(prv_to_fixture_ft=10.0)
        self.assertEqual(g.midstream_prv_ids, [])

    def test_just_over_ten_feet_is_midstream(self):
        g = _build(prv_to_fixture_ft=10.5)
        self.assertEqual(g.midstream_prv_ids, [3])
        self.assertTrue(g.nodes[3].is_midstream_prv)

    def test_elbow_equivalents_do_not_count(self):
        # 8 ft of pipe with an elbow is 13 ft developed, but it is still a
        # regulator at its equipment.
        g = _build(prv_to_fixture_ft=8.0, elbow_after_prv=True)
        self.assertEqual(g.midstream_prv_ids, [])
        self.assertAlmostEqual(g.nodes[3].downstream_pipe_ft, 8.0)

    def test_regulator_keywords_are_shared(self):
        self.assertEqual(shared_params.PRV_FAMILY_KEYWORDS,
                         ("prv", "regulator", "regulating"))
        name = "RJA - Pressure Regulating Valve".lower()
        self.assertTrue(any(k in name for k in
                            shared_params.PRV_FAMILY_KEYWORDS))
        for other in ("rja - gas meter", "rja- gas equipment cap",
                      "elbow - generic"):
            self.assertFalse(any(k in other for k in
                                 shared_params.PRV_FAMILY_KEYWORDS), other)


class ZoneSizing(unittest.TestCase):
    def test_each_side_uses_its_own_table_and_run(self):
        g = _build()
        r = sizing_engine.size_system(
            g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)
        # After the PRV: RTU-1 800 MBH on the low-pressure table, 20 ft run.
        self.assertEqual(r["sizes"][103], _expected_size(DOWN, 20.0, 800.0))
        # Before it: the trunk carries 1200 MBH on the 2 psi table, 90 ft run.
        self.assertEqual(r["sizes"][101], _expected_size(UP, 90.0, 1200.0))
        self.assertEqual(r["sizes"][104], _expected_size(UP, 90.0, 400.0))
        # The feed to the regulator carries only the regulated load.
        self.assertEqual(r["sizes"][102], _expected_size(UP, 90.0, 800.0))
        # The regulator really changes the answer: the same pipe sized on
        # the meter's 2 psi table and the whole-system run would be smaller.
        self.assertNotEqual(r["sizes"][103], _expected_size(UP, 100.0, 800.0))
        self.assertEqual(r["edge_context"][103]["table_id"], DOWN)
        self.assertEqual(r["edge_context"][101]["table_id"], UP)

    def test_zones_summary_lists_meter_first(self):
        g = _build()
        r = sizing_engine.size_system(
            g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)
        self.assertEqual(len(r["zones"]), 2)
        self.assertEqual(r["zones"][0]["table_id"], UP)
        self.assertEqual(r["zones"][1]["table_id"], DOWN)
        self.assertEqual(r["longest_run_ft"], 90.0)

    def test_without_downstream_table_matches_single_system(self):
        g = _build()
        r = sizing_engine.size_system(g, "Schedule 40 Steel", 2.0, UP)
        self.assertIsNone(r["downstream_table_id"])
        # One system, one table, the whole-system 100 ft run.
        self.assertEqual(r["longest_run_ft"], 100.0)
        for eid, size in r["sizes"].items():
            mbh = g.edges[eid].cumulative_load_mbh
            self.assertEqual(size, _expected_size(UP, 100.0, mbh), eid)
        self.assertTrue(all(c["table_id"] == UP
                            for c in r["edge_context"].values()))

    def test_downstream_table_with_no_prv_raises(self):
        g = _build(with_prv=False)
        with self.assertRaises(ValueError):
            sizing_engine.size_system(
                g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)

    def test_downstream_table_with_only_an_equipment_prv_raises(self):
        g = _build(prv_to_fixture_ft=4.0)
        with self.assertRaises(ValueError):
            sizing_engine.size_system(
                g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)

    def test_equipment_prv_is_sized_as_if_absent(self):
        with_prv = _build(prv_to_fixture_ft=4.0)
        without = _build(with_prv=False, prv_to_fixture_ft=4.0)
        a = sizing_engine.size_system(with_prv, "Schedule 40 Steel", 2.0, UP)
        b = sizing_engine.size_system(without, "Schedule 40 Steel", 2.0, UP)
        self.assertEqual(a["sizes"], b["sizes"])

    def test_nested_prv_raises(self):
        g = _build_two(15.0)
        with self.assertRaises(ValueError):
            sizing_engine.size_system(
                g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)

    def test_equipment_prv_beyond_a_midstream_one_still_sizes(self):
        g = _build_two(4.0)
        r = sizing_engine.size_system(
            g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)
        # Everything after PRV 3 is on the downstream table, including the
        # pipe past the equipment PRV.
        self.assertEqual(r["edge_context"][105]["table_id"], DOWN)
        self.assertEqual(r["edge_context"][103]["table_id"], DOWN)

    def test_format_output_names_the_zones(self):
        g = _build()
        r = sizing_engine.size_system(
            g, "Schedule 40 Steel", 2.0, UP, downstream_table_id=DOWN)
        text = sizing_engine.format_sizing_output(r, g)
        self.assertIn("Downstream:", text)
        self.assertIn(DOWN, text)
        self.assertIn("After PRV 3", text)


class DownstreamDropdown(unittest.TestCase):
    def test_downstream_dropdown_is_same_material_and_gas(self):
        labels = gas_tables.get_table_option_labels_for_material_and_gas(
            "Schedule 40 Steel", "Natural")
        self.assertTrue(labels)
        for lbl in labels:
            opt = gas_tables.get_table_option_by_material_and_short_label(
                "Schedule 40 Steel", lbl)
            self.assertEqual(opt["gas"], "Natural")
            self.assertEqual(opt["material"], "Schedule 40 Steel")
        propane = gas_tables.get_table_option_labels_for_material_and_gas(
            "Schedule 40 Steel", "Propane")
        self.assertFalse(set(labels) & set(propane))


if __name__ == "__main__":
    unittest.main()
