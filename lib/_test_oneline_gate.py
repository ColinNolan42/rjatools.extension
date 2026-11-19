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
        self.s["nodes"][CB1]["load_mbh"] = 0.0              # model reason
        res = gate.check(self.s)
        self.assertEqual(res["headline"], gate.FIX_MODEL)
        self.assertEqual(res["errors"][0]["code"], "zero_load")
        self.assertTrue(gate.summary_lines(res)[0].startswith(
            "Fix the piping model"))

    def test_stepped_down_pipe(self):
        self.s["edges"][FRY1_DROP]["size"] = "3/4"
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["size_mismatch"])
        self.assertEqual(res["errors"][0]["id"], 3618361)
        self.assertIn("requires 1", res["errors"][0]["text"])

    def test_unreadable_size(self):
        self.s["edges"][FRY1_DROP]["size"] = None
        self.assertEqual(_codes(gate.check(self.s)), ["size_unreadable"])

    def test_unsized(self):
        self.s["edges"][FRY1_DROP]["required_size"] = None
        self.assertEqual(_codes(gate.check(self.s)), ["unsized"])

    def test_sizing_error(self):
        self.s["sizing_error"] = "demand exceeds table"
        self.assertEqual(_codes(gate.check(self.s)), ["sizing_error"])

    def test_zero_mbh_named_fixture(self):
        self.s["nodes"][CB1]["load_mbh"] = 0.0
        res = gate.check(self.s)
        self.assertEqual(_codes(res), ["zero_load"])
        self.assertIn("CB-1", res["errors"][0]["text"])

    def test_future_at_zero_is_fine(self):
        self.s["nodes"][FUTURE]["load_mbh"] = 0.0
        self.assertTrue(gate.check(self.s)["ok"])
        self.s["nodes"][FUTURE]["name"] = ""
        self.assertTrue(gate.check(self.s)["ok"])

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

    def test_midstream_prv_needs_downstream_table(self):
        self.s["midstream_prv_ids"] = [5]
        self.assertEqual(_codes(gate.check(self.s)), ["prv_no_table"])
        self.s["downstream_table_id"] = "402.4(2)"
        self.assertTrue(gate.check(self.s)["ok"])

    def test_non_gas_is_warning_only(self):
        self.s["non_gas_skipped"] = [[77, "Domestic Cold Water"]]
        res = gate.check(self.s)
        self.assertTrue(res["ok"])
        self.assertEqual(res["warnings"][0]["code"], "non_gas")


if __name__ == "__main__":
    unittest.main()
