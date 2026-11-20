# -*- coding: ascii -*-
"""Tests for sizing_engine.sizing_rows and serves_text (Size Gas console
table data, 2026-11-19). Plain CPython with the Revit assemblies stubbed by
_test_prv_zones. format_sizing_output must still print what it printed.

Run:  python -I _test_gas_sizing_rows.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import _test_prv_zones as Z   # noqa: E402  (installs the Revit stubs)
import sizing_engine          # noqa: E402


def _size(graph, down=None):
    return sizing_engine.size_system(
        graph, "Schedule 40 Steel", 0.5, Z.DOWN, heat_content_btu_per_cf=840.0,
        downstream_table_id=down)


class SizingRows(unittest.TestCase):

    def test_one_row_per_sized_pipe_with_the_engine_numbers(self):
        g = Z._build(with_prv=False)
        res = _size(g)
        rows = sizing_engine.sizing_rows(res, g)
        self.assertEqual(sorted(r["pipe_id"] for r in rows),
                         sorted(res["sizes"]))
        by = dict((r["pipe_id"], r) for r in rows)
        trunk = by[101]
        self.assertEqual(trunk["size"], res["sizes"][101])
        self.assertEqual(trunk["mbh"], 1200.0)
        self.assertAlmostEqual(trunk["cfh"], 1200000.0 / 840.0, 0)
        self.assertEqual(trunk["units"], "CFH")
        self.assertEqual(trunk["length_ft"], 50.0)
        self.assertEqual(trunk["status"], "OK")
        self.assertEqual(trunk["zone"], "Meter")
        self.assertTrue(trunk["serves"].startswith("trunk (2:"))
        self.assertEqual(by[104]["serves"], "GUH-1")
        self.assertEqual(trunk["table_id"], Z.DOWN)
        self.assertGreaterEqual(trunk["capacity"], trunk["cfh"])

    def test_zero_load_is_min_size(self):
        g = Z._build(with_prv=False)
        g.nodes[5].gas_load_mbh = 0.0
        Z.pipe_graph._calculate_cumulative_loads(g)
        rows = dict((r["pipe_id"], r) for r in
                    sizing_engine.sizing_rows(_size(g), g))
        self.assertEqual(rows[104]["status"], "MIN SIZE")
        self.assertEqual(rows[104]["mbh"], 0.0)
        self.assertIn("zero demand", rows[104]["note"])
        self.assertEqual(rows[101]["status"], "OK")

    def test_zones_label_each_pipe_with_its_system(self):
        g = Z._build()
        res = _size(g, down=Z.DOWN)
        rows = dict((r["pipe_id"], r) for r in
                    sizing_engine.sizing_rows(res, g))
        self.assertEqual(rows[101]["zone"], res["zones"][0]["label"])
        self.assertEqual(rows[103]["zone"], "After PRV 3")
        self.assertEqual(rows[103]["table_id"], Z.DOWN)

    def test_open_ended_pipe_has_no_row_and_unsized_is_over_table(self):
        g = Z._build(with_prv=False)
        res = _size(g)
        e = Z._edge(g, 999, 5, None, 3.0)
        self.assertNotIn(999, [r["pipe_id"] for r in
                               sizing_engine.sizing_rows(res, g)])
        # A pipe the engine did not size but that has a downstream end.
        del res["sizes"][104]
        res["segment_detail"] = [d for d in res["segment_detail"]
                                 if d["pipe_id"] != 104]
        rows = dict((r["pipe_id"], r) for r in
                    sizing_engine.sizing_rows(res, g))
        self.assertEqual(rows[104]["status"], "OVER TABLE")
        self.assertEqual(rows[104]["size"], "")
        self.assertIsNone(rows[104]["capacity"])
        del e

    def test_text_output_unchanged_by_the_refactor(self):
        g = Z._build(with_prv=False)
        text = sizing_engine.format_sizing_output(_size(g), g)
        self.assertIn("->  ", text)
        self.assertIn("  -> GUH-1", text)
        self.assertIn("  -> trunk (2: ", text)

    def test_serves_text(self):
        g = Z._build(with_prv=False)
        self.assertEqual(sizing_engine.serves_text(g.edges[104], g), "GUH-1")
        self.assertEqual(sizing_engine.serves_text(None, g), "")
        open_edge = Z._edge(g, 998, 5, None, 1.0)
        self.assertEqual(sizing_engine.serves_text(open_edge, g), "")


if __name__ == "__main__":
    unittest.main()
