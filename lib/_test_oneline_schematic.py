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


class LabelsSumToTotal(unittest.TestCase):
    """Colin 2026-11-20: the longest run's labels showed 1 ft more than the
    notes TOTAL DEVELOPED LENGTH, each segment rounded on its own."""

    def test_g3c_numbers(self):
        # 26.17 + 13.44 + 35.60 + 22.56 + 26.50 = 124.27 -> 124; rounding
        # each gives 26 + 13 + 36 + 23 + 27 = 125. Largest remainder: floors
        # 26 13 35 22 26 = 122, the 2 ft left go to .60 and .56.
        self.assertEqual(
            S.balanced_feet([26.17, 13.44, 35.60, 22.56, 26.50], 124),
            [26, 13, 36, 23, 26])

    def test_small_and_tie(self):
        self.assertEqual(S.balanced_feet([0.3, 10.2], 11), [1, 10])
        self.assertEqual(S.balanced_feet([2.5, 2.5], 5), [3, 2])
        self.assertEqual(S.balanced_feet([], 0), [])

    def _label_ft(self, seg):
        return int(seg["label"][0].split(", ")[-1].rstrip("'"))

    def _check(self, snap):
        import gas_labels
        sch = S.build(snap)
        total = gas_labels.feet(snap["longest_run_ft"])
        labels = [self._label_ft(s) for s in S.trunk_segments(sch)]
        self.assertEqual(sum(labels), total, labels)
        for s in S.trunk_segments(sch):
            self.assertLessEqual(abs(self._label_ft(s) - s["length_ft"]), 1.0)

    def test_g3c(self):
        self._check(G3C)

    def test_g2(self):
        self._check(G2)

    def test_g3c_manifold(self):
        self._check(_load("oneline_g3c_manifold.json"))


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
        # length = pipe + 5 ft per elbow + 3 ft per tee (Colin, 2026-11-19)
        # S1 13.17 + 2x5 + 3 = 26.17; S2 10.44 + 3 = 13.44;
        # S3 22.60 + 2x5 + 3 = 35.60; S5 19.56 + 3 = 22.56;
        # S6 21.50 + 5 = 26.50 (ends at RTU-1, no tee). Sum 124.26.
        trunk = [s["label"][0] for s in S.trunk_segments(self.sch)]
        self.assertEqual(trunk, [
            '3"G, 26\'', '3"G, 13\'', '2"G, 36\'', '1-1/2"G, 23\'',
            '1-1/4"G, 26\''])

    def test_item1_header_spans_are_short(self):
        # kitchen header between taps: 2.58 ft pipe + 3 ft for the tee it
        # arrives at = 5.58, reads 6 ft (2.56 + 3 = 5.56 on the 1-1/2 in.)
        spans = [s for s in self.sch["segments"].values()
                 if s["branch_kind"] == "run" and not s["on_longest"]
                 and s["mbh"] in (755.0, 665.0, 450.0, 330.0)]
        self.assertEqual(len(spans), 4)
        for s in spans:
            self.assertEqual(s["label"][0].split(", ")[1], "6'", s["label"])

    def test_item1_drop_lengths_own_pipes(self):
        seg, _br = self._drop("CB-1")
        self.assertEqual(seg["elbows"], 2)
        self.assertAlmostEqual(seg["pipe_ft"], 5.10 + 7.22 + 1.06, places=1)
        # 13.38 ft pipe + 2 elbows x 5 ft = 23.38, no tee on a drop
        self.assertEqual(seg["label"], ['1"G, 23\'', "100 MBH"])

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
        # 13.45 ft pipe + 2 elbows x 5 ft = 23.45
        self.assertEqual(sp["label"], ['1"G, 23\'', "90 MBH"])

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
            # 7.82 ft pipe + 3 ft for the tee it arrives at = 10.82
            self.assertEqual(h, '1"G, 11\'')   # K&L stripped from the label
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
             "phase": None}
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
            "longest_run_ft": run, "elbow_equiv_ft": 5.0, "tee_equiv_ft": 3.0,
            "origin_phase": None}


class Synthetic(unittest.TestCase):

    def test_size_change_mid_run_splits(self):
        # meter 1 -> coupling 2 -> fixture 3; pipe 2" then 1-1/2"
        s = _snap([(1, "meter", [0, 0, 0]),
                   (2, "fitting", [10, 0, 0], {"is_transition": True}),
                   (3, "fixture", [20, 0, 0], {"name": "RTU-1", "load_mbh": 100.0})],
                  [(10, 1, 2, 9.8, "2", 100.0), (11, 2, 3, 9.7, "1-1/2", 100.0)])
        sch = S.build(s)
        segs = [sch["segments"][x] for x in sch["root"]["segments"]]
        self.assertEqual([x["size"] for x in segs], ["2", "1-1/2"])
        # pipe only: a transition adds no length of its own
        self.assertAlmostEqual(segs[0]["length_ft"], 9.8)
        self.assertAlmostEqual(segs[1]["length_ft"], 9.7)

    def test_tee_three_feet_on_the_arriving_segment(self):
        # meter 1 -10 ft-> tee 2 -> L (5 ft) / R (9 ft). The 3 ft sits on the
        # segment that arrives at the tee, so the trunk is 10 + 3 + 9 = 22.
        s = self._bullhead(5, 9)
        s["longest_path"] = [1, 10, 2, 12, 4]
        sch = S.build(s)
        first = sch["segments"][sch["root"]["segments"][0]]
        self.assertEqual((first["tees"], first["elbows"]), (1, 0))
        self.assertAlmostEqual(first["length_ft"], 13.0)
        legs = [sch["segments"][x] for x in
                sch["root"]["end"]["run"]["segments"]]
        self.assertEqual([x["tees"] for x in legs], [0])
        self.assertAlmostEqual(S.trunk_sum_ft(sch), 22.0)
        for seg in sch["segments"].values():
            if seg is not first:
                self.assertEqual(seg["tees"], 0, seg["id"])

    def _tee_trunk(self, branch_longest):
        # meter 1 -10-> tee 2 (kind tee, riser at [10,0,0])
        #   straight: tee 2 -12 ft-> fixture 3
        #   branch:   tee 2 -> elbow 4 -> fixture 5 (6 + 8 ft)
        s = _snap([(1, "meter", [0, 0, 0]),
                   (2, "tee", [10, 0, 0]),
                   (3, "fixture", [22, 0, 0], {"name": "S", "load_mbh": 50.0}),
                   (4, "fitting", [10, 6, 0], {"is_elbow": True}),
                   (5, "fixture", [18, 6, 0], {"name": "B", "load_mbh": 50.0})],
                  [(10, 1, 2, 10.0, "1", 100.0),
                   (11, 2, 3, 12.0 if not branch_longest else 2.0, "3/4", 50.0),
                   (12, 2, 4, 6.0, "3/4", 50.0),
                   (13, 4, 5, 8.0, "3/4", 50.0)])
        s["longest_path"] = ([1, 10, 2, 12, 4, 13, 5] if branch_longest
                             else [1, 10, 2, 11, 3])
        return s

    def test_trunk_sums_through_a_tee_straight_on(self):
        # 10 + 12 ft pipe + 3 ft tee = 25
        sch = S.build(self._tee_trunk(False))
        self.assertAlmostEqual(S.trunk_sum_ft(sch), 25.0)

    def test_trunk_sums_through_a_tee_into_the_branch(self):
        # 10 + 6 + 8 ft pipe + 5 ft elbow + 3 ft tee = 32
        sch = S.build(self._tee_trunk(True))
        self.assertAlmostEqual(S.trunk_sum_ft(sch), 32.0)
        trunk = S.trunk_segments(sch)
        self.assertEqual(sum(x["tees"] for x in trunk), 1)
        self.assertEqual(sum(x["elbows"] for x in trunk), 1)

    def test_trunk_matches_pipe_graph_longest_run(self):
        # The same tee graph the PRV zone tests use, run through the real
        # snapshot: 50 + 30 + 20 ft of pipe + 3 ft for tee 2 = 103.
        import oneline_snapshot
        from _test_prv_zones import _build as build_graph
        g = build_graph(with_prv=True, prv_to_fixture_ft=20.0)
        snap = oneline_snapshot.snapshot(g, None, "402.4(2)")
        self.assertAlmostEqual(snap["longest_run_ft"], 103.0)
        sch = S.build(snap)
        self.assertAlmostEqual(S.trunk_sum_ft(sch), snap["longest_run_ft"])

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
