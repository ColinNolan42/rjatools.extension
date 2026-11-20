# -*- coding: ascii -*-
"""Tests for oneline_snapshot.py and the captured bench fixtures
(One-Line rebuild step 2, 2026-11-18).

Fixtures in _fixtures/ were captured read only from Gas Piping Example
(Revit 2027) with pipe_graph.build_network on the meter, the in memory
sizing_engine.size_system rerun, then oneline_snapshot.snapshot:
  oneline_g3c.json  meter 3618292, 402.4(2), heat 840 (sized, Size Gas PASS)
  oneline_g2.json   meter 3592842, 402.4(12), heat 840 (NOT sized yet:
                    32 pipes differ from the required size, a gate reject case)

Rigor: the snapshot() unit test runs on a stub graph in CPython; the
fixture tests check what the live capture recorded. Neither re-reads Revit.

Run:  python -I _test_oneline_snapshot.py
"""

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from _test_prv_zones import _build   # installs the Revit stubs   # noqa: E402

import oneline_snapshot   # noqa: E402

FIX = os.path.join(HERE, "_fixtures")


def _load(name):
    return oneline_snapshot.load_json(os.path.join(FIX, name))


class FixtureIntegrity(unittest.TestCase):

    def _check(self, snap):
        nodes, edges = snap["nodes"], snap["edges"]
        self.assertIn(str(snap["origin_id"]), nodes)
        for e in edges.values():
            self.assertIn(str(e["from"]), nodes, e["id"])
            if e["to"] is not None:
                self.assertIn(str(e["to"]), nodes, e["id"])
        for k, kids in snap["children"].items():
            self.assertIn(k, nodes)
            for c in kids:
                self.assertIn(str(c), nodes)
        for i in snap["longest_path"]:
            self.assertTrue(str(i) in nodes or str(i) in edges, i)
        # the length inside a fitting is no longer carried (2026-11-19)
        self.assertEqual(snap["tee_equiv_ft"], 3.0)
        self.assertEqual(snap["elbow_equiv_ft"], 5.0)
        for n in nodes.values():
            if "is_tee" in n:
                self.assertEqual(
                    n["is_tee"],
                    n["kind"] == "tee" and not n["is_gas_fixture"] and
                    not n["is_prv"], n["id"])

    def test_g3c(self):
        self._check(_load("oneline_g3c.json"))

    def test_g2(self):
        self._check(_load("oneline_g2.json"))


class G3CFacts(unittest.TestCase):
    """Handoff model facts: 1814 MBH, 12 fixtures, 124.26 ft (87.26 pipe +
    5 elbows x 5 ft + 4 tees x 3 ft), row 125."""

    @classmethod
    def setUpClass(cls):
        cls.s = _load("oneline_g3c.json")

    def test_run_and_row(self):
        self.assertAlmostEqual(self.s["longest_run_ft"], 124.26, places=2)
        self.assertEqual(self.s["sizing"]["table_length_used_ft"], 125.0)
        self.assertEqual(self.s["origin_kind"], "meter")
        self.assertEqual(self.s["table_id"], "402.4(2)")

    def test_fixtures(self):
        fx = [n for n in self.s["nodes"].values() if n["is_gas_fixture"]]
        self.assertEqual(len(fx), 12)
        self.assertAlmostEqual(sum(n["load_mbh"] for n in fx), 1814.0)
        names = sorted(n["name"] for n in fx)
        self.assertIn("SP-1", names)
        self.assertIn("FUTURE", names)

    def test_sized_to_size_gas(self):
        bad = [e["id"] for e in self.s["edges"].values()
               if e["size"] != e["required_size"]]
        self.assertEqual(bad, [])

    def test_meter_pipe(self):
        meter = self.s["origin_id"]
        first = [e for e in self.s["edges"].values() if e["from"] == meter]
        kids = self.s["children"].get(str(meter), [])
        self.assertTrue(first or kids)
        cum = max([e["cum_mbh"] for e in self.s["edges"].values()])
        self.assertAlmostEqual(cum, 1814.0)

    def test_all_new_construction(self):
        phases = set(n["phase"] for n in self.s["nodes"].values())
        self.assertEqual(phases, set(["New Construction"]))


class G2NotSized(unittest.TestCase):

    def test_g2_has_size_mismatches(self):
        s = _load("oneline_g2.json")
        bad = [e for e in s["edges"].values()
               if e["size"] != e["required_size"]]
        self.assertEqual(len(bad), 32)
        self.assertAlmostEqual(s["longest_run_ft"], 229.49, places=2)
        names = sorted(n["name"] for n in s["nodes"].values()
                       if n["is_gas_fixture"])
        for r in ("RTU-1", "RTU-2", "RTU-3", "RTU-4", "RTU-5", "RTU-6"):
            self.assertIn(r, names)


class SnapshotUnit(unittest.TestCase):

    def test_stub_graph(self):
        g = _build(with_prv=True, prv_to_fixture_ft=20.0)
        sizing = {"sizes": {101: "2", 102: "1-1/2"},
                  "table_id": "402.4(2)", "longest_run_ft": 100.0,
                  "table_length_used_ft": 100.0,
                  "heat_content_btu_per_cf": 840.0,
                  "edge_context": {}}
        s = oneline_snapshot.snapshot(g, None, "402.4(2)", sizing_result=sizing)
        self.assertEqual(sorted(s["nodes"].keys()), ["1", "2", "3", "4", "5"])
        self.assertEqual(s["edges"]["101"]["required_size"], "2")
        self.assertIsNone(s["edges"]["104"]["required_size"])
        self.assertIsNone(s["edges"]["101"]["size"])   # diameter 0 in stub
        self.assertEqual(s["edges"]["101"]["from"], 1)
        self.assertTrue(s["nodes"]["3"]["is_prv"])
        self.assertEqual(s["nodes"]["4"]["name"], "RTU-1")
        # 50 + 30 + 20 ft of pipe + 3 ft for tee 2 (the path passes it)
        self.assertAlmostEqual(s["longest_run_ft"], 103.0)
        self.assertNotIn("takeout_ft", s["nodes"]["2"])
        self.assertTrue(s["nodes"]["2"]["is_tee"])
        self.assertFalse(s["nodes"]["1"]["is_tee"])
        self.assertFalse(s["nodes"]["3"]["is_tee"])
        self.assertIsNone(s["nodes"]["1"]["phase"])
        self.assertEqual(s["sizing"]["table_length_used_ft"], 100.0)

    def test_size_gas_zones_carried(self):
        g = _build(with_prv=False)
        sizing = {"sizes": {}, "edge_context": {},
                  "zones": [{"label": "Meter to regulator(s)",
                             "table_id": "402.4(5)", "run_ft": 84.6123,
                             "row_ft": 90.0, "farthest_name": "x"},
                            {"label": "After PRV 3", "table_id": "402.4(2)",
                             "run_ft": 139.5, "row_ft": 150.0}]}
        s = oneline_snapshot.snapshot(g, None, "402.4(2)", sizing_result=sizing)
        self.assertEqual(s["sizing"]["zones"][0],
                         {"label": "Meter to regulator(s)",
                          "table_id": "402.4(5)", "run_ft": 84.6123,
                          "row_ft": 90.0})
        self.assertEqual(len(s["sizing"]["zones"]), 2)

    def test_unsized_zero_load_carried(self):
        g = _build(with_prv=False)
        sizing = {"sizes": {101: "2"}, "edge_context": {},
                  "unsized_zero_load": [{"pipe_id": 104, "fixture_ids": [5],
                                         "table_id": "402.4(2)"}]}
        s = oneline_snapshot.snapshot(g, None, "402.4(2)", sizing_result=sizing)
        self.assertEqual(s["unsized_zero_load"],
                         [{"pipe_id": 104, "fixture_ids": [5]}])

    def test_round_trip(self):
        import tempfile
        g = _build(with_prv=False)
        s = oneline_snapshot.snapshot(g, None, "402.4(2)")
        fd, path = tempfile.mkstemp(suffix=".json")
        os.close(fd)
        try:
            oneline_snapshot.dump_json(s, path)
            back = oneline_snapshot.load_json(path)
        finally:
            os.remove(path)
        self.assertEqual(back["edges"]["101"]["length_ft"], 50.0)
        self.assertEqual(back["longest_path"], s["longest_path"])


class _Id(object):
    def __init__(self, v):
        self.IntegerValue = v
        self.Value = v


class _Phases(object):
    def __init__(self, ids):
        self.l = [type("P", (), {"Id": _Id(i)})() for i in ids]
        self.Size = len(self.l)

    def get_Item(self, i):
        return self.l[i]


class _Doc(object):
    """Phases Existing (1) and New Construction (2); GetElement fails."""
    Phases = _Phases([1, 2])

    def GetElement(self, _pid):
        raise RuntimeError("no phase element")


class _El(object):
    def __init__(self, eid, phase):
        self.Id = _Id(eid)
        self.CreatedPhaseId = _Id(phase)


class _BadEl(object):
    Id = _Id(77)

    @property
    def CreatedPhaseId(self):
        raise RuntimeError("unreadable")


class PhaseRule(unittest.TestCase):
    """Review n16: new construction is the project's LAST phase, read with
    revit_helpers.is_new_construction (the Size Gas rule)."""

    def test_last_phase_is_new(self):
        self.assertTrue(oneline_snapshot._is_new(_Doc(), _El(5, 2)))
        self.assertFalse(oneline_snapshot._is_new(_Doc(), _El(5, 1)))
        self.assertIsNone(oneline_snapshot._is_new(None, _El(5, 2)))

    def test_snapshot_carries_is_new(self):
        g = _build(with_prv=False)
        for n in g.nodes.values():
            n.element = _El(n.node_id if hasattr(n, "node_id") else 1, 1)
        for e in g.edges.values():
            e.pipe = _El(e.element_id, 2)
        s = oneline_snapshot.snapshot(g, _Doc(), "402.4(2)")
        self.assertTrue(all(e["is_new"] for e in s["edges"].values()))
        self.assertTrue(all(n["is_new"] is False for n in s["nodes"].values()))


class Joints(unittest.TestCase):
    """pipe_graph joint nodes (2026-11-20): no Revit element is ever read for
    a joint id, and it takes the far end of the pipe that ends at it."""

    def _graph(self):
        import pipe_graph
        from _test_prv_zones import _node, _edge, _finish
        g = pipe_graph.NetworkGraph()
        g.origin_id = 1
        _node(g, 1, "meter").location_xyz = [0.0, 0.0, 0.0]
        j = _node(g, 101, "joint")
        j.element = None
        j.location_xyz = None
        _node(g, 4, "fixture", 90.0, "RTU-1").location_xyz = [30.0, 0.0, 0.0]
        e1 = _edge(g, 101, 1, 101, 12.0)
        e1.start_xyz, e1.end_xyz = [0.0, 0.0, 0.0], [12.0, 0.0, 0.0]
        e1.pipe = _El(101, 2)
        e2 = _edge(g, 102, 101, 4, 18.0)
        e2.pipe = _El(102, 2)
        g.data_errors = [{"element_id": 9, "name": "FUTURE", "kind": "zero_load",
                          "message": "GAS_LOAD_MBH is 0"}]
        g.tapped_pipes = [101]
        _finish(g)
        return g

    def test_joint_no_element_read(self):
        s = oneline_snapshot.snapshot(self._graph(), _Doc(), "402.4(2)")
        j = s["nodes"]["101"]
        self.assertEqual(j["kind"], "joint")
        self.assertIsNone(j["phase"])
        self.assertIsNone(j["is_new"])
        self.assertEqual(j["xyz"], [12.0, 0.0, 0.0])
        self.assertTrue(s["edges"]["101"]["is_new"])
        self.assertEqual(s["data_errors"][0]["kind"], "zero_load")
        self.assertEqual(s["tapped_pipes"], [101])
        self.assertFalse(any("element None" in w for w in s["read_warnings"]))

    def test_joint_snapshot_builds(self):
        import oneline_schematic as S
        s = oneline_snapshot.snapshot(self._graph(), None, "402.4(2)")
        sch = S.build(s)
        self.assertEqual(sch["root"]["end"]["name"], "RTU-1")
        self.assertEqual(sorted(S.eid_counts(sch)), [101, 102])


class ReadFailuresLogged(unittest.TestCase):
    """Review n46: phase and level read failures are reported (Rule 6)."""

    def test_phase_failure_warns(self):
        w = []
        self.assertIsNone(oneline_snapshot._phase_name(_Doc(), _BadEl(), w))
        self.assertEqual(len(w), 1)
        self.assertIn("Phase of element", w[0])

    def test_unreadable_phase_is_existing_and_warns(self):
        w = []
        self.assertFalse(oneline_snapshot._is_new(_Doc(), _BadEl(), w))
        self.assertTrue(any("drawn as existing" in x for x in w))

    def test_level_failure_warns(self):
        w = []
        self.assertEqual(oneline_snapshot.read_levels(_Doc(), w), [])
        self.assertEqual(len(w), 1)
        self.assertIn("Levels could not be read", w[0])

    def test_snapshot_lists_read_warnings(self):
        g = _build(with_prv=False)
        s = oneline_snapshot.snapshot(g, _Doc(), "402.4(2)")
        self.assertTrue(any("Levels could not be read" in x
                            for x in s["read_warnings"]))


if __name__ == "__main__":
    unittest.main()
