# -*- coding: ascii -*-
"""Tests for oneline_layout.py (One-Line rebuild step 4, 2026-11-18).

Checks (plan section 8): no two primitives clash on either bench; nothing
but the trunk on the meter row (item 9); every Segment has one label,
centered on its last drawn line (item 4); inline terminal labels sit right
of the symbol (item 7); every equipment end has its PRV symbol; a forced
collision moves a junction or a header instead of overlapping.

Rigor: geometry is checked on estimated text boxes (3/32 in. Arial at 1:100,
CHAR_W_FT per character). The real Revit text extents are checked live in
step 8.

Run:  python -I _test_oneline_layout.py
"""

import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import gas_labels              # noqa: E402
import oneline_layout as L     # noqa: E402
import oneline_schematic as S  # noqa: E402
from _test_oneline_schematic import _snap   # noqa: E402


def _load(name):
    f = open(os.path.join(HERE, "_fixtures", name))
    try:
        return json.load(f)
    finally:
        f.close()


def _run(snap):
    sch = S.build(snap)
    tot = sum(n["load_mbh"] for n in snap["nodes"].values()
              if n["is_gas_fixture"])
    notes = gas_labels.notes_lines(snap["table_id"], tot,
                                   snap["longest_run_ft"])
    return sch, L.layout(sch, notes=notes)


G3C_SCH, G3C = _run(_load("oneline_g3c.json"))
G2_SCH, G2 = _run(_load("oneline_g2.json"))


def _of(out, **kw):
    res = []
    for p in out["prims"]:
        if all(p.get(k) == v for k, v in kw.items()):
            res.append(p)
    return res


class ClashRules(unittest.TestCase):

    def L(self, *a):
        return L._line(*a)

    def test_corner_and_tee_joints_are_fine(self):
        self.assertFalse(L.clash(self.L(0, 0, 10, 0), self.L(10, 0, 10, -9)))
        self.assertFalse(L.clash(self.L(0, 0, 10, 0), self.L(10, 0, 20, 0)))

    def test_t_touch_is_a_clash(self):
        # one line ending in the middle of another is not a drawn tee
        self.assertTrue(L.clash(self.L(0, 0, 20, 0), self.L(10, 0, 10, 9)))

    def test_crossing_and_overlap(self):
        self.assertTrue(L.clash(self.L(0, 0, 20, 0), self.L(10, -5, 10, 5)))
        self.assertTrue(L.clash(self.L(0, 0, 20, 0), self.L(10, 0, 30, 0)))

    def test_line_near_symbol(self):
        sym = L._symbol("equipment", 10.0, 0.0, 90, 1.0, 0.56)
        sym["box"] = (9.44, -1.0, 10.56, 1.0)
        own = self.L(0, 0, 9.44, 0)                 # its own pipe ends at it
        passing = self.L(9.3, -9, 9.3, 9)           # runs along its edge
        self.assertFalse(L.clash(own, sym))
        self.assertTrue(L.clash(passing, sym))

    def test_text_boxes(self):
        a = L._text(0, 0, ["ABC"], "center")
        b = L._text(0.5, 0, ["ABC"], "center")
        c = L._text(10, 0, ["ABC"], "center")
        self.assertTrue(L.clash(a, b))
        self.assertFalse(L.clash(a, c))


class Benches(unittest.TestCase):

    def _common(self, sch, out):
        prims = out["prims"]
        self.assertEqual(L.all_clashes(prims), [])
        self.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
        # independent plain box test: no two texts or symbols overlap
        boxes = [p for p in prims if p["t"] != "line"]
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                self.assertFalse(L._boxes_overlap(boxes[i]["box"],
                                                  boxes[j]["box"], 0.0),
                                 (boxes[i].get("lines"), boxes[j].get("lines")))
        # one label per segment, on its last line
        for sid, seg in sch["segments"].items():
            labels = _of(out, t="text", role="seg", seg=sid)
            self.assertEqual(len(labels), 1, sid)
            lab = labels[0]
            self.assertEqual(lab["lines"], seg["label"])
            lines = _of(out, t="line", seg=sid)
            self.assertTrue(lines, sid)
            last = lines[-1]
            if abs(last["y0"] - last["y1"]) < 1e-6:          # horizontal
                xs = [l["x0"] for l in lines if abs(l["y0"] - last["y0"]) < 1e-6
                      and abs(l["y0"] - l["y1"]) < 1e-6]
                xs += [l["x1"] for l in lines if abs(l["y0"] - last["y0"]) < 1e-6
                       and abs(l["y0"] - l["y1"]) < 1e-6]
                self.assertTrue(min(xs) <= lab["x"] <= max(xs), sid)
                self.assertGreater(lab["box"][1], last["y0"], sid)
            else:                                             # vertical
                self.assertGreater(lab["box"][0], last["x0"], sid)
        # every terminal: a symbol, a label; equipment: a PRV before it
        for t in S.terminals_in_order(sch):
            nid = t["node_id"]
            self.assertEqual(len(_of(out, t="text", role="fixture", node=nid)), 1)
            if t["kind"] == "equipment":
                self.assertEqual(len(_of(out, t="symbol", role="equipment",
                                         node=nid)), 1, t["name"])
            if t["kind"] == "cap":
                self.assertEqual(len(_of(out, t="symbol", role="cap",
                                         node=nid)), 1)
        n_eq = len([t for t in S.terminals_in_order(sch)
                    if t["kind"] == "equipment"])
        eq_prvs = [p for p in prims if p["t"] == "symbol" and
                   p["role"] == "prv" and p.get("equipment_prv")]
        self.assertEqual(len(eq_prvs), n_eq)
        # meter row: only trunk lines and the meter's utility stub
        for l in _of(out, t="line"):
            if abs(l["y0"]) < 1e-6 and abs(l["y1"]) < 1e-6:
                if l["seg"] is not None:
                    self.assertTrue(sch["segments"][l["seg"]]["on_longest"],
                                    l["seg"])
        # notes block top left, above everything else
        notes = _of(out, t="text", role="notes")[0]
        others = [p for p in prims if p is not notes]
        self.assertGreaterEqual(notes["box"][1],
                                max(p["box"][3] for p in others))
        self.assertLessEqual(notes["box"][0],
                             min(p["box"][0] for p in others) + 1e-6)

    def test_g3c(self):
        self._common(G3C_SCH, G3C)

    def test_g2(self):
        self._common(G2_SCH, G2)

    def _inline(self, sch, out, name):
        t = [x for x in S.terminals_in_order(sch) if x["name"] == name][0]
        self.assertTrue(t["inline"])
        role = "equipment" if t["kind"] == "equipment" else "cap"
        sym = _of(out, t="symbol", role=role, node=t["node_id"])[0]
        lab = _of(out, t="text", role="fixture", node=t["node_id"])[0]
        self.assertGreater(lab["box"][0], sym["box"][2])       # right of it
        mid = (lab["box"][1] + lab["box"][3]) / 2.0
        self.assertAlmostEqual(mid, sym["y"], places=6)        # centered
        return sym

    def test_item7_inline_labels_right(self):
        sym = self._inline(G3C_SCH, G3C, "RTU-1")
        self.assertEqual(sym["rot"], 90)
        self._inline(G3C_SCH, G3C, "FUTURE")
        self._inline(G2_SCH, G2, "RTU-5")

    def test_item9_kitchen_header_off_meter_row(self):
        # the kitchen header (915 down to 120 MBH) sits two rows or more
        # from the meter row; WH-1 hangs below the meter row
        segs = [s for s in G3C_SCH["segments"].values()
                if s["mbh"] in (855.0, 755.0, 665.0, 450.0, 330.0, 210.0)
                and s["branch_kind"] == "run"]
        for s in segs:
            for l in _of(G3C, t="line", seg=s["id"]):
                self.assertLessEqual(max(l["y0"], l["y1"]), -2 * L.ROW_FT + 1e-6)
        wh = [t for t in S.terminals_in_order(G3C_SCH) if t["name"] == "WH-1"][0]
        sym = _of(G3C, t="symbol", role="equipment", node=wh["node_id"])[0]
        self.assertLess(sym["box"][3], -L.ROW_FT / 2.0)

    def test_drops_hang_from_their_tee_column(self):
        # drops are vertical, one row
        for t in S.terminals_in_order(G3C_SCH):
            if t["inline"]:
                continue
            br = S.terminal_branch(G3C_SCH, t["name"])
            lines = _of(G3C, t="line", seg=br["segments"][0])
            for l in lines:
                self.assertAlmostEqual(l["x0"], l["x1"])


class ForcedCollision(unittest.TestCase):

    def test_two_header_taps_one_column_apart(self):
        # meter 1 -> tee 2 -> tee 3 -> RTU-0 9; each tee feeds an upper
        # header with two units. Both headers want the same row: the second
        # shall move out a row or its junction shall move right.
        z = 20.0
        nodes = [(1, "meter", [0, 0, 0]), (2, "tee", [10, 0, 0]),
                 (3, "tee", [12, 0, 0]),
                 (9, "fixture", [30, 0, 0], {"name": "RTU-0", "load_mbh": 10.0}),
                 (4, "tee", [10, 5, z]),
                 (5, "fixture", [10, 10, z], {"name": "A-1", "load_mbh": 10.0}),
                 (6, "fixture", [20, 5, z], {"name": "A-2", "load_mbh": 10.0}),
                 (7, "tee", [12, 5, z]),
                 (8, "fixture", [12, 10, z], {"name": "B-1", "load_mbh": 10.0}),
                 (10, "fixture", [22, 5, z], {"name": "B-2", "load_mbh": 10.0})]
        edges = [(100, 1, 2, 10.0, "2", 50.0), (101, 2, 3, 2.0, "2", 30.0),
                 (102, 3, 9, 18.0, "1", 10.0),
                 (103, 2, 4, 20.0, "1", 20.0), (104, 4, 5, 5.0, "1", 10.0),
                 (105, 4, 6, 10.0, "1", 10.0),
                 (106, 3, 7, 20.0, "1", 20.0), (107, 7, 8, 5.0, "1", 10.0),
                 (108, 7, 10, 10.0, "1", 10.0)]
        snap = _snap(nodes, edges, path=[1, 100, 2, 101, 3, 102, 9])
        sch = S.build(snap)
        out = L.layout(sch)
        self.assertEqual(L.all_clashes(out["prims"]), [])
        self.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
        self.assertTrue(out["log"], "expected a stretch or a row move")


if __name__ == "__main__":
    unittest.main()
