# -*- coding: ascii -*-
"""Tests for oneline_gate.py (One-Line rebuild step 2b, 2026-11-18).

One-Line draws only a complete system sized to Size Gas (Colin). G3C is the
sized control and passes; G2 is not sized yet and rejects; each other case
is G3C with one defect seeded into a copy of the snapshot.

Run:  python -I _test_oneline_gate.py
"""

import copy
import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import oneline_gate as gate   # noqa: E402


def _load(name):
    f = open(os.path.join(HERE, "_fixtures", name))
    try:
        return json.load(f)
    finally:
        f.close()


G3C = _load("oneline_g3c.json")
G2 = _load("oneline_g2.json")

FRY1_DROP = "3618361"      # 1" drop to FRY-1
CB1 = "3618461"            # CB-1 fixture
FUTURE = "3618455"         # FUTURE cap, 120 MBH
TEE = "3618409"            # FRY-1 tee


def _zero(snap, nid, name):
    """A fixture at 0 MBH as the walk records it: load 0 on the node and
    one pipe_graph.data_errors entry, carried by the snapshot (item B)."""
    snap["nodes"][nid]["load_mbh"] = 0.0
    snap.setdefault("data_errors", []).append(
        {"element_id": int(nid), "name": name, "kind": "zero_load",
         "message": "GAS_LOAD_MBH is 0. Data error"})


def _codes(res):
    return sorted(set(e["code"] for e in res["errors"]))


class Gate(unittest.TestCase):

    def setUp(self):
        self.s = copy.deepcopy(G3C)

    def test_g3c_passes(self):
        res = gate.check(self.s)
        self.assertTrue(res["ok"], res["errors"])
        self.assertEqual(gate.summary_lines(res), [])

    def test_g2_not_sized_rejects(self):
        res = gate.check(G2)
        self.assertFalse(res["ok"])
        self.assertEqual(_codes(res), ["size_mismatch"])
        self.assertEqual(len(res["errors"]), 32)
        lines = gate.summary_lines(res, limit=3)
        self.assertTrue(lines[0].startswith("Run Size Gas first"))
        self.assertIn("29 more", lines[-1])

    def test_model_problem_headline_and_order(self):
        # G3 2026-11-19: a model problem says fix the model, listed first
        self.s["edges"][FRY1_DROP]["size"] = "3/4"          # sizing reason
        self.s["edges"][FRY1_DROP]["diameter_in"] = 0.75
        _zero(self.s, CB1, "CB-1")                          # model reason
        res = gate.check(self.s)
        self.assertEqual(res["headline"], gate.FIX_MODEL)
        self.assertEqual(res["errors"][0]["code"], "zero_load")
        self.assertTrue(gate.summary_lines(res)[0].startswith(
            "Fix the piping model"))

    def test_stepped_down_pipe(self):
        # the gate compares diameters (Colin 2026-11-20), so the drawn pipe
        # is stepped down in diameter, not only in name
        self.s["edges"][FRY1_DROP]["size"] = "3/4"
        self.s["edges"][FRY1_DROP]["diameter_in"] = 0.75
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["size_mismatch"])
        self.assertEqual(res["errors"][0]["id"], 3618361)
        self.assertIn("requires 1", res["errors"][0]["text"])

    def test_unreadable_size(self):
        self.s["edges"][FRY1_DROP]["size"] = None
        self.s["edges"][FRY1_DROP]["diameter_in"] = 0.9
        self.assertEqual(_codes(gate.check(self.s)), ["size_unreadable"])

    def test_unsized(self):
        self.s["edges"][FRY1_DROP]["required_size"] = None
        self.assertEqual(_codes(gate.check(self.s)), ["unsized"])

    def test_sizing_error(self):
        self.s["sizing_error"] = "demand exceeds table"
        self.assertEqual(_codes(gate.check(self.s)), ["sizing_error"])

    def test_zero_mbh_named_fixture(self):
        _zero(self.s, CB1, "CB-1")
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["zero_load"])
        self.assertIn("CB-1", res["errors"][0]["text"])

    def test_future_at_zero_is_a_data_error(self):
        # Colin 2026-11-20: a FUTURE cap shall carry its planned MBH
        _zero(self.s, FUTURE, "FUTURE")
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["zero_load"])
        self.assertEqual(res["headline"], gate.FIX_MODEL)
        self.s["nodes"][FUTURE]["name"] = ""
        self.assertEqual(_codes(gate.check(self.s)), ["zero_load"])

    def test_unsized_zero_load_pipe_is_a_data_error(self):
        # Size Gas leaves a pipe to only 0 MBH fixtures unsized (2026-11-20):
        # the gate says fix the model, not "Run Size Gas first"
        self.s["edges"][FRY1_DROP]["required_size"] = None
        self.s["unsized_zero_load"] = [{"pipe_id": int(FRY1_DROP),
                                        "fixture_ids": [77]}]
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["zero_load_pipe"])
        self.assertEqual(res["headline"], gate.FIX_MODEL)

    def test_walk_data_errors_listed_once(self):
        self.s["nodes"][FUTURE]["load_mbh"] = 0.0
        self.s["data_errors"] = [
            {"element_id": int(FUTURE), "name": "FUTURE", "kind": "zero_load",
             "message": "GAS_LOAD_MBH is 0"},
            {"element_id": 4242, "name": "B-9", "kind": "zero_load",
             "message": "GAS_LOAD_MBH text 'x' is not a number"}]
        res = gate.check(self.s)
        ids = [e["id"] for e in res["errors"] if e["code"] == "zero_load"]
        self.assertEqual(sorted(ids), sorted([int(FUTURE), 4242]))
        self.assertEqual(res["headline"], gate.FIX_MODEL)

    def test_data_errors_are_the_one_source(self):
        # item B: the gate does not test load <= 0 again; the walk's record
        # (snapshot data_errors) is what rejects
        self.s["nodes"][CB1]["load_mbh"] = 0.0
        self.assertEqual(_codes(gate.check(self.s)), [])
        _zero(self.s, CB1, "CB-1")
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["zero_load"])
        self.assertEqual(len(res["errors"]), 1)
        src = open(os.path.join(HERE, "oneline_gate.py")).read()
        self.assertNotIn("load <= 0", src)

    def test_loaded_fixture_without_name(self):
        self.s["nodes"][CB1]["name"] = ""
        self.assertEqual(_codes(gate.check(self.s)), ["no_name"])

    def test_open_pipe(self):
        self.s["edges"][FRY1_DROP]["to"] = None
        self.assertIn("open_pipe", _codes(gate.check(self.s)))

    def test_open_fitting(self):
        self.s["nodes"][TEE]["connectors"][0]["to"] = None
        self.assertEqual(_codes(gate.check(self.s)), ["open_fitting"])

    def test_meter_inlet_open_is_fine(self):
        meter = self.s["nodes"][str(self.s["origin_id"])]
        self.assertTrue(any(c["to"] is None for c in meter["connectors"]))
        self.assertTrue(gate.check(self.s)["ok"])

    def test_end_not_gas_fixture(self):
        self.s["nodes"][CB1]["is_gas_fixture"] = False
        self.assertIn("end_not_fixture", _codes(gate.check(self.s)))

    def test_disconnected(self):
        self.s["disconnected"] = [123]
        self.assertEqual(_codes(gate.check(self.s)), ["disconnected"])

    def test_nested_prv(self):
        self.s["nested_prv_ids"] = [5, 6]
        self.assertEqual(_codes(gate.check(self.s)), ["nested_prv"])

    def test_midstream_prv_box_off_warns_never_refuses(self):
        # review n15 (Colin 2026-11-20): with Mid Stream PRV unticked the
        # One-Line does what Size Gas does, checks on the meter table, warns
        # (the warning is the button's prv_notes, printed after drawing)
        self.s["midstream_prv_ids"] = [5]
        res = gate.check(self.s)
        self.assertTrue(res["ok"], res["errors"])
        self.s["downstream_table_id"] = "402.4(2)"
        self.assertTrue(gate.check(self.s)["ok"])

    # --- review n2 / n6 / n19: CSST pairs share a diameter -----------------
    def test_csst_ehd25_read_back_as_ehd23_passes(self):
        e = self.s["edges"][FRY1_DROP]
        e["table_id"] = "402.4(15)"
        e["required_size"] = "EHD-25"
        e["size"] = "EHD-23"           # nominal_for_diameter's read back
        e["diameter_in"] = 0.75
        res = gate.check(self.s)
        self.assertTrue(res["ok"], res["errors"])

    def test_csst_wrong_diameter_still_rejects(self):
        e = self.s["edges"][FRY1_DROP]
        e["table_id"] = "402.4(15)"
        e["required_size"] = "EHD-25"
        e["size"] = "EHD-18"
        e["diameter_in"] = 0.5
        self.assertEqual(_codes(gate.check(self.s)), ["size_mismatch"])

    # --- review n18: existing pipe is never refused for size ---------------
    def test_existing_pipe_larger_is_not_refused(self):
        e = self.s["edges"][FRY1_DROP]
        e["is_new"] = False
        e["size"], e["diameter_in"] = "2", 2.0
        res = gate.check(self.s)
        self.assertTrue(res["ok"], res["errors"])
        self.assertEqual([w for w in res["warnings"]
                          if w["code"].startswith("existing")], [])

    def test_existing_pipe_smaller_warns_and_draws(self):
        e = self.s["edges"][FRY1_DROP]
        e["is_new"] = False
        e["size"], e["diameter_in"] = "1/2", 0.5
        res = gate.check(self.s)
        self.assertTrue(res["ok"], res["errors"])
        w = [x for x in res["warnings"] if x["code"] == "existing_small"]
        self.assertEqual(len(w), 1)
        self.assertIn("smaller than the 1", w[0]["text"])

    def test_new_pipe_still_checked(self):
        e = self.s["edges"][FRY1_DROP]
        e["is_new"] = True
        e["size"], e["diameter_in"] = "2", 2.0
        self.assertEqual(_codes(gate.check(self.s)), ["size_mismatch"])

    # --- review n17: an edge into an element the walk did not keep --------
    def test_edge_to_unknown_node_rejects(self):
        self.s["edges"][FRY1_DROP]["to"] = 999999
        res = gate.check(self.s)
        self.assertFalse(res["ok"])
        self.assertIn("edge_to_missing", _codes(res))
        self.assertEqual(res["headline"], gate.FIX_MODEL)

    # --- review n60: more than one leg off the origin ----------------------
    def test_two_origin_legs_rejects(self):
        origin = str(self.s["origin_id"])
        kids = list(self.s["children"].get(origin) or [])
        self.assertEqual(len(gate.origin_legs(self.s)), 1)
        self.s["children"][origin] = kids + [int(TEE)]
        res = gate.check(self.s)
        self.assertIn("origin_legs", _codes(res))
        self.assertEqual(res["headline"], gate.FIX_MODEL)

    # --- review n46: model read failures reach the console -----------------
    def test_read_warnings_are_printed(self):
        self.s["read_warnings"] = ["Levels could not be read (boom)."]
        res = gate.check(self.s)
        self.assertTrue(res["ok"])
        self.assertIn("Levels could not be read (boom).",
                      [w["text"] for w in res["warnings"]])

    def test_non_gas_is_warning_only(self):
        self.s["non_gas_skipped"] = [[77, "Domestic Cold Water"]]
        res = gate.check(self.s)
        self.assertTrue(res["ok"])
        self.assertEqual(res["warnings"][0]["code"], "non_gas")


if __name__ == "__main__":
    unittest.main()
