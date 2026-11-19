# -*- coding: ascii -*-
"""Tests for oneline_schematic.py (One-Line rebuild step 3, 2026-11-18).

Invariants on both bench fixtures (plan section 8):
  (a) the segments on the longest path sum to the longest run (unrounded);
  (b) every pipe id is in exactly one segment;
  (c) MBH never rises away from the meter;
  (d) one size per segment.
G3C expected labels (handoff visual pass items 1, 2, 3) and the Fraser
RTU cascade order on G2 are regression cases. Synthetic snapshots cover a
size change mid run, a bullhead tee, PRVs and isolation valves.

Run:  python -I _test_oneline_schematic.py
"""

import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import oneline_schematic as S   # noqa: E402


def _load(name):
    f = open(os.path.join(HERE, "_fixtures", name))
    try:
        return json.load(f)
    finally:
        f.close()


G3C = _load("oneline_g3c.json")
G2 = _load("oneline_g2.json")


class Invariants(unittest.TestCase):

    def _check(self, snap):
        sch = S.build(snap)
        self.assertEqual(sch["warnings"], [])
        self.assertAlmostEqual(S.trunk_sum_ft(sch), snap["longest_run_ft"],
                               places=2)
        counts = S.eid_counts(sch)
        self.assertEqual(sorted(counts.keys()),
                         sorted(int(k) for k in snap["edges"]))
        self.assertEqual([k for k, v in counts.items() if v != 1], [])
        self.assertEqual(S.mbh_violations(sch), [])
        for seg in sch["segments"].values():
            self.assertLessEqual(len(seg["sizes_seen"]), 1, seg["id"])
        return sch

    def test_g3c(self):
        self._check(G3C)

    def test_g2(self):
        self._check(G2)


class G3CLabels(unittest.TestCase):
    """Handoff visual pass, G3C (402.4(2), heat 840, row 125)."""

    @classmethod
    def setUpClass(cls):
        cls.sch = S.build(G3C)

    def _drop(self, name):
        br = S.terminal_branch(self.sch, name)
        self.assertIsNotNone(br, name)
        segs = [self.sch["segments"][s] for s in br["segments"]]
        self.assertEqual(len(segs), 1, name)
        return segs[0], br

    def test_trunk_labels(self):
        # in order from the meter out, not dict order (G2 2026-11-19)
        trunk = [s["label"][0] for s in S.trunk_segments(self.sch)]
        self.assertEqual(trunk, [
            '3"G, 25\'', '3"G, 11\'', '2"G, 34\'', '1-1/2"G, 20\'',
            '1-1/4"G, 27\''])

    def test_item1_header_spans_are_short(self):
        # kitchen header between taps reads about 3 ft, not 27 ft
        spans = [s for s in self.sch["segments"].values()
                 if s["branch_kind"] == "run" and not s["on_longest"]
                 and s["mbh"] in (755.0, 665.0, 450.0, 330.0)]
        self.assertEqual(len(spans), 4)
        for s in spans:
            self.assertEqual(s["label"][0].split(", ")[1], "3'", s["label"])

    def test_item1_drop_lengths_own_pipes(self):
        seg, _br = self._drop("CB-1")
        self.assertEqual(seg["elbows"], 2)
        self.assertAlmostEqual(seg["pipe_ft"], 5.10 + 7.22 + 1.06, places=1)
        self.assertEqual(seg["label"], ['1"G, 24\'', "100 MBH"])

    def test_item2_drop_sizes(self):
        want = {"OVEN-1": "3/4", "CB-1": "1", "GRIDDLE-1": "1",
                "RANGE-1": "1-1/4", "FRY-2": "1", "FRY-1": "1", "SP-1": "1",
                "WH-1": "1-1/4", "MAU-1": "1-1/2", "RTU-2": "1"}
        for name, size in want.items():
            seg, _br = self._drop(name)
            self.assertEqual(seg["size"], size, name)

    def test_item3_future_last_and_inline(self):
        order = [t["name"] for t in S.terminals_in_order(self.sch)]
        self.assertEqual(order[-1], "FUTURE")
        self.assertEqual(order[-2], "SP-1")
        seg, br = self._drop("FUTURE")
        self.assertEqual(br["kind"], "run")
        self.assertEqual(seg["label"], ['1"G, 6\'', "120 MBH"])
        self.assertTrue(br["end"]["inline"])
        self.assertEqual(br["end"]["kind"], "cap")
        sp, spbr = self._drop("SP-1")
        self.assertEqual(spbr["kind"], "tap")
        self.assertEqual(sp["label"], ['1"G, 24\'', "90 MBH"])

    def test_kitchen_header_order(self):
        order = [t["name"] for t in S.terminals_in_order(self.sch)]
        k = order.index("OVEN-1")
        self.assertEqual(order[k:], ["OVEN-1", "CB-1", "GRIDDLE-1", "RANGE-1",
                                     "FRY-2", "FRY-1", "SP-1", "FUTURE"])

    def test_header_mbh_sequence(self):
        mbh = []

        def _v(br, _d):
            if br["kind"] == "run" or br["kind"] == "root":
                for sid in br["segments"]:
                    mbh.append(self.sch["segments"][sid]["mbh"])
        S.walk(self.sch["root"], _v)
        for seq in ([915.0, 855.0, 755.0, 665.0, 450.0, 330.0, 210.0, 120.0],):
            i = mbh.index(seq[0])
            self.assertEqual(mbh[i:i + len(seq)], seq)

    def test_rtu1_inline_equipment_with_prv(self):
        _seg, br = self._drop("RTU-1")
        self.assertTrue(br["end"]["inline"])
        self.assertTrue(br["end"]["equipment_prv"])

    def test_every_equipment_shows_prv(self):
        for t in S.terminals_in_order(self.sch):
            self.assertEqual(t["equipment_prv"], t["kind"] == "equipment",
                             t["name"])


class Fraser(unittest.TestCase):
    """G2 RTU cascade: taps along the header, then the run (Colin)."""

    @classmethod
    def setUpClass(cls):
        cls.sch = S.build(G2)

    def test_order(self):
        order = [t["name"] for t in S.terminals_in_order(self.sch)]
        rtus = [n for n in order if n.startswith("RTU")]
        self.assertEqual(rtus, ["RTU-6", "RTU-4", "RTU-3", "RTU-1", "RTU-2",
                                "RTU-5"])

    def test_spine(self):
        spine = [s for s in S.trunk_segments(self.sch)]
        mbh = [s["mbh"] for s in spine]
        for v in (630.0, 505.0, 405.0, 315.0, 190.0, 90.0):
            self.assertIn(v, mbh)
        taps = {}
        for t in ("RTU-6", "RTU-4", "RTU-3", "RTU-1", "RTU-2"):
            br = S.terminal_branch(self.sch, t)
            taps[t] = self.sch["segments"][br["segments"][0]]["label"][0]
        hops = [s["label"][0] for s in spine if s["mbh"] in
                (505.0, 405.0, 315.0, 190.0)]
        for h in hops:
            self.assertEqual(h, '1"G, 8\'')   # K&L stripped from the label
            self.assertNotIn(h, taps.values())

    def test_equipment_prvs_fold_into_terminals(self):
        for t in S.terminals_in_order(self.sch):
            if t["name"].startswith("RTU"):
                self.assertIsNotNone(t["modeled_prv_id"], t["name"])
        for s in self.sch["segments"].values():
            self.assertEqual([v for v in s["valves"] if v["kind"] == "prv"],
                             [], s["id"])

    def test_no_collinear_fork_takes_longest(self):
        # tee 3593357: riser up, AHU-1 one way, ERV-1/AHU-2 the other
        found = []

        def _v(br, _d):
            e = br["end"]
            if e and e["type"] == "junction" and e["node_id"] == 3593357:
                found.append(e)
        S.walk(self.sch["root"], _v)
        j = found[0]
        self.assertEqual(j["run"]["end"]["type"], "junction")   # to ERV/AHU-2
        self.assertEqual(j["taps"][0]["end"]["name"], "AHU-1")


# ---------------------------------------------------------------------------
# synthetic snapshots
# ---------------------------------------------------------------------------

def _snap(nodes, edges, children=None, path=None, run=0.0):
    nd = {}
    for n in nodes:
        d = {"id": n[0], "kind": n[1], "xyz": n[2], "family": "",
             "is_elbow": False, "is_transition": False, "is_prv": False,
             "is_midstream_prv": False, "is_gas_fixture": n[1] == "fixture",
             "load_mbh": 0.0, "name": "", "cum_mbh": 0.0, "connectors": [],
             "takeout_ft": {}, "phase": None}
        d.update(n[3] if len(n) > 3 else {})
        nd[str(n[0])] = d
    ed = {}
    for e in edges:
        ed[str(e[0])] = {"id": e[0], "from": e[1], "to": e[2],
                         "length_ft": e[3], "size": e[4], "cum_mbh": e[5],
                         "start": nd[str(e[1])]["xyz"],
                         "end": nd[str(e[2])]["xyz"], "phase": None}
    return {"origin_id": nodes[0][0], "nodes": nd, "edges": ed,
            "children": children or {}, "longest_path": path or [],
            "longest_run_ft": run, "elbow_equiv_ft": 5.0,
            "origin_phase": None}


class Synthetic(unittest.TestCase):

    def test_size_change_mid_run_splits(self):
        # meter 1 -> coupling 2 -> fixture 3; pipe 2" then 1-1/2"
        s = _snap([(1, "meter", [0, 0, 0]),
                   (2, "fitting", [10, 0, 0], {"is_transition": True,
                                               "takeout_ft": {"10": 0.2, "11": 0.3}}),
                   (3, "fixture", [20, 0, 0], {"name": "RTU-1", "load_mbh": 100.0})],
                  [(10, 1, 2, 9.8, "2", 100.0), (11, 2, 3, 9.7, "1-1/2", 100.0)])
        sch = S.build(s)
        segs = [sch["segments"][x] for x in sch["root"]["segments"]]
        self.assertEqual([x["size"] for x in segs], ["2", "1-1/2"])
        self.assertAlmostEqual(segs[0]["length_ft"], 10.0)
        self.assertAlmostEqual(segs[1]["length_ft"], 10.0)

    def _bullhead(self, left_len, right_len):
        # meter 1 -> tee 2 (riser) -> left fixture 3 / right fixture 4
        return _snap([(1, "meter", [0, 0, 0]),
                      (2, "tee", [0, 0, 10]),
                      (3, "fixture", [-left_len, 0, 10], {"name": "L", "load_mbh": 50.0}),
                      (4, "fixture", [right_len, 0, 10], {"name": "R", "load_mbh": 50.0})],
                     [(10, 1, 2, 10.0, "1", 100.0),
                      (11, 2, 3, float(left_len), "3/4", 50.0),
                      (12, 2, 4, float(right_len), "3/4", 50.0)])

    def test_bullhead_longest_leg_is_run(self):
        sch = S.build(self._bullhead(5, 9))
        order = [t["name"] for t in S.terminals_in_order(sch)]
        self.assertEqual(order, ["L", "R"])
        sch = S.build(self._bullhead(9, 5))
        order = [t["name"] for t in S.terminals_in_order(sch)]
        self.assertEqual(order, ["R", "L"])

    def test_midstream_prv_stays_on_segment(self):
        s = _snap([(1, "meter", [0, 0, 0]),
                   (2, "fitting", [10, 0, 0], {"is_prv": True,
                                               "is_midstream_prv": True,
                                               "family": "RJA - Pressure Regulating Valve"}),
                   (3, "fixture", [40, 0, 0], {"name": "RTU-1", "load_mbh": 100.0})],
                  [(10, 1, 2, 10.0, "1", 100.0), (11, 2, 3, 30.0, "1", 100.0)])
        sch = S.build(s)
        seg = sch["segments"][sch["root"]["segments"][0]]
        self.assertEqual([v["kind"] for v in seg["valves"]], ["prv"])
        self.assertIsNone(sch["root"]["end"]["modeled_prv_id"])

    def test_isolation_valve_recorded(self):
        s = _snap([(1, "meter", [0, 0, 0]),
                   (2, "fitting", [10, 0, 0], {"family": "RJA - Ball Valve - Gas"}),
                   (3, "fixture", [20, 0, 0], {"name": "WH-1", "load_mbh": 100.0})],
                  [(10, 1, 2, 10.0, "1", 100.0), (11, 2, 3, 10.0, "1", 100.0)])
        sch = S.build(s)
        seg = sch["segments"][sch["root"]["segments"][0]]
        self.assertEqual(seg["valves"], [{"kind": "iso", "node_id": 2,
                                          "midstream": False}])

    def test_ansul_valve_recorded(self):
        # Colin 2026-11-19: the hood's Ansul gas shutoff, interlocked with
        # the fire suppression, ahead of the cookline
        s = _snap([(1, "meter", [0, 0, 0]),
                   (2, "fitting", [10, 0, 0],
                    {"family": "RJA - Gas Shutoff Valve (Ansul)"}),
                   (3, "fixture", [20, 0, 0], {"name": "FRY-1", "load_mbh": 100.0})],
                  [(10, 1, 2, 10.0, "1", 100.0), (11, 2, 3, 10.0, "1", 100.0)])
        sch = S.build(s)
        seg = sch["segments"][sch["root"]["segments"][0]]
        self.assertEqual([v["kind"] for v in seg["valves"]], ["ansul"])

    def test_unnamed_zero_cap_label(self):
        s = _snap([(1, "meter", [0, 0, 0]), (2, "fixture", [10, 0, 0])],
                  [(10, 1, 2, 10.0, "1", 0.0)])
        end = S.build(s)["root"]["end"]
        self.assertEqual(end["kind"], "cap")
        self.assertEqual(end["label"], ["FUTURE CONNECTION"])
        self.assertFalse(end["equipment_prv"])


if __name__ == "__main__":
    unittest.main()
