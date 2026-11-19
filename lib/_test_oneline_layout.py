# -*- coding: ascii -*-
"""Tests for oneline_layout.py (One-Line rebuild step 4, 2026-11-18).

Checks (plan section 8): no two primitives clash on either bench; elevation
by level band, first floor equipment on the meter row, roof units higher,
no header along the meter row (item 9, Colin 2026-11-18); every Segment has one label,
centered on its last drawn line (item 4); every unit turns up or down,
roof units up and the rest down, the last unit on a header included (Colin
2026-11-19, supersedes item 7); every equipment end has its PRV symbol; a forced
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

    def test_text_box_matches_revit(self):
        # measured on G2 2026-11-19 (3/32" Arial, 1:100): 1 line 1.39 ft,
        # 2 lines 2.65, 6 lines 7.68 below the top; the old 0.78 / 1.3
        # estimate put roof unit labels onto their symbols
        for n, ft in ((1, 1.39), (2, 2.65), (6, 7.683)):
            self.assertAlmostEqual(L._text_size(["X"] * n)[1], ft, delta=0.05)

    def test_unconnected_segments_shall_not_meet(self):
        # G2 2026-11-19: AHU-2 turning up met the boiler riser end to end
        a = self.L(0, 0, 10, 0, "S6")
        b = self.L(10, 0, 10, 9, "S19")
        a["joins"], b["joins"] = {"S6", "S4"}, {"S19", "S7"}
        self.assertTrue(L.clash(a, b))
        b["joins"] = {"S19", "S6"}
        a["joins"] = {"S6", "S19"}
        self.assertFalse(L.clash(a, b))

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
                # above the line; below it on a header whose units hang up
                # (Colin 2026-11-19)
                self.assertTrue(lab["box"][1] > last["y0"] or
                                lab["box"][3] < last["y0"], sid)
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
        # meter row: no header runs along it, only the meter outlet stub
        meter = _of(out, t="symbol", role="meter")[0]
        root_first = sch["root"]["segments"][0]
        for l in _of(out, t="line"):
            if abs(l["y0"] - meter["y"]) < 1e-6 and \
                    abs(l["y1"] - meter["y"]) < 1e-6 and l["seg"] is not None:
                self.assertEqual(l["seg"], root_first)
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

    def _hang(self, sch, out, name):
        """'up' or 'down': the label sits above or below its symbol, on its
        center line, the way every unit on a header hangs."""
        t = [x for x in S.terminals_in_order(sch) if x["name"] == name][0]
        role = "equipment" if t["kind"] == "equipment" else "cap"
        sym = _of(out, t="symbol", role=role, node=t["node_id"])[0]
        lab = _of(out, t="text", role="fixture", node=t["node_id"])[0]
        self.assertIn(sym["rot"], (0, 180), name)        # pipe vertical
        self.assertAlmostEqual(lab["x"], sym.get("ix", sym["x"]), places=6,
                               msg=name)
        if lab["box"][1] >= sym["box"][3]:
            return "up"
        self.assertLessEqual(lab["box"][3], sym["box"][1], name)
        return "down"

    def test_last_unit_on_a_header_turns(self):
        # Colin 2026-11-19 (supersedes item 7): AHU-2, B-2, RTU-5 and the
        # other header end units turn like their neighbours
        for sch, out in ((G3C_SCH, G3C), (G2_SCH, G2)):
            for t in S.terminals_in_order(sch):
                self._hang(sch, out, t["name"])

    def test_roof_hangs_up_rest_hangs_down(self):
        # Colin 2026-11-19: roof (highest level above the meter) hangs up
        for name in ("MAU-1", "RTU-1", "RTU-2"):
            self.assertEqual(self._hang(G3C_SCH, G3C, name), "up", name)
        for name in ("WH-1", "OVEN-1", "FRY-1", "SP-1", "FUTURE"):
            self.assertEqual(self._hang(G3C_SCH, G3C, name), "down", name)
        for name in ("RTU-1", "RTU-5", "AHU-2", "B-2", "ERV-1"):
            self.assertEqual(self._hang(G2_SCH, G2, name), "up", name)

    def test_g2_first_floor_boilers(self):
        # G2 2026-11-19: boilers moved to the first floor off the first
        # floor main. The layout failed (the header tried up into the roof
        # labels first); it shall drop below the main and stay on one row
        snap = _load("oneline_g2.json")
        for n in snap["nodes"].values():
            xyz = n.get("xyz")
            if xyz and 188 <= xyz[0] <= 212 and xyz[1] <= 15 and xyz[2] > 20:
                xyz[2] = 4.0
        sch = S.build(snap)
        out = L.layout(sch, notes=["x"] * 6)
        self.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
        self.assertEqual(L.all_clashes(out["prims"]), [])
        rows = set()
        for name in ("B-1", "B-2", "B-3"):
            self.assertEqual(self._hang(sch, out, name), "down", name)
            rows.add(round(self._eq_sym(sch, out, name)["y"], 6))
        self.assertEqual(len(rows), 1)                    # one header row
        for name in ("RTU-5", "AHU-2"):
            self.assertEqual(self._hang(sch, out, name), "up", name)

    def test_ansul_valve_drawn_and_named(self):
        # Colin 2026-11-19: cookline fed from one drop with the Ansul valve
        # ahead of the first appliance, manifold behind the equipment
        nodes = [(1, "meter", [0, 0, 1.0]), (2, "tee", [10, 0, 9.5]),
                 (3, "fitting", [12, 5, 6.0],
                  {"family": "RJA - Gas Shutoff Valve (Ansul)"}),
                 (4, "tee", [12, 5, 2.0]), (5, "tee", [15, 5, 2.0]),
                 (6, "fixture", [12, 4, 2.0], {"name": "FRY-1", "load_mbh": 120.0}),
                 (7, "fixture", [15, 4, 2.0], {"name": "RANGE-1", "load_mbh": 215.0}),
                 (8, "fixture", [18, 4, 2.0], {"name": "OVEN-1", "load_mbh": 60.0}),
                 (9, "fixture", [40, 0, 1.0], {"name": "WH-1", "load_mbh": 199.0})]
        edges = [(50, 1, 2, 12.0, "2", 594.0), (51, 2, 3, 8.0, "1-1/4", 395.0),
                 (52, 3, 4, 4.0, "1-1/4", 395.0), (53, 4, 6, 1.0, "3/4", 120.0),
                 (54, 4, 5, 3.0, "1-1/4", 275.0), (55, 5, 7, 1.0, "1", 215.0),
                 (56, 5, 8, 3.0, "3/4", 60.0), (57, 2, 9, 30.0, "1", 199.0)]
        snap = _snap(nodes, edges, path=[1, 50, 2, 57, 9])
        snap["levels"] = [{"name": "FIRST", "elev": 0.0},
                          {"name": "ROOF", "elev": 20.0}]
        sch = S.build(snap)
        out = L.layout(sch)
        self.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
        self.assertEqual(L.all_clashes(out["prims"]), [])
        labs = _of(out, t="text", role="valve")
        self.assertEqual([p["lines"] for p in labs], [["ANSUL VALVE"]])
        self.assertEqual(len(_of(out, t="symbol", role="ansul", node=3)), 1)
        for name in ("FRY-1", "RANGE-1", "OVEN-1", "WH-1"):
            self.assertEqual(self._hang(sch, out, name), "down", name)

    def test_one_level_building_hangs_down(self):
        # meter and every unit on one level: nothing is a roof
        nodes = [(1, "meter", [0, 0, 1.0]), (2, "tee", [10, 0, 9.0]),
                 (5, "fixture", [10, 8, 1.0], {"name": "WH-2", "load_mbh": 150.0}),
                 (6, "fixture", [30, 0, 1.0], {"name": "OVEN-9", "load_mbh": 90.0})]
        edges = [(50, 1, 2, 10.0, "1-1/4", 240.0),
                 (51, 2, 5, 15.0, "1", 150.0), (52, 2, 6, 20.0, "1", 90.0)]
        snap = _snap(nodes, edges, path=[1, 50, 2, 52, 6])
        snap["levels"] = [{"name": "FIRST", "elev": 0.0},
                          {"name": "ROOF", "elev": 20.0}]
        sch = S.build(snap)
        out = L.layout(sch)
        self.assertEqual(L.all_clashes(out["prims"]), [])
        for name in ("WH-2", "OVEN-9"):
            self.assertEqual(self._hang(sch, out, name), "down", name)

    def _eq_sym(self, sch, out, name):
        t = [x for x in S.terminals_in_order(sch) if x["name"] == name][0]
        role = "equipment" if t["kind"] == "equipment" else "cap"
        return _of(out, t="symbol", role=role, node=t["node_id"])[0]

    def test_elevation_first_floor_on_meter_row(self):
        # Colin 2026-11-18: the meter is on the first level, shown at the
        # same elevation as the kitchen equipment and WH-1
        meter = _of(G3C, t="symbol", role="meter")[0]
        for name in ("WH-1", "OVEN-1", "CB-1", "GRIDDLE-1", "RANGE-1",
                     "FRY-2", "FRY-1", "SP-1"):
            sym = self._eq_sym(G3C_SCH, G3C, name)
            self.assertAlmostEqual(sym["box"][3], meter["y"], places=6, msg=name)

    def test_item9_kitchen_header_above_meter_row(self):
        segs = [s for s in G3C_SCH["segments"].values()
                if s["mbh"] in (855.0, 755.0, 665.0, 450.0, 330.0, 210.0)
                and s["branch_kind"] == "run"]
        self.assertEqual(len(segs), 6)
        for s in segs:
            for l in _of(G3C, t="line", seg=s["id"]):
                self.assertAlmostEqual(l["y0"], L.ROW_FT)
                self.assertAlmostEqual(l["y1"], L.ROW_FT)

    def test_elevation_roof_units_higher(self):
        kitchen_header_y = L.ROW_FT
        for name in ("MAU-1", "RTU-2", "RTU-1"):
            sym = self._eq_sym(G3C_SCH, G3C, name)
            self.assertGreater(sym["box"][1], kitchen_header_y, name)
        for name in ("RTU-1", "RTU-2", "RTU-3", "RTU-4", "RTU-5", "RTU-6",
                     "AHU-1", "AHU-2", "ERV-1", "B-1", "B-2", "B-3"):
            sym = self._eq_sym(G2_SCH, G2, name)
            self.assertGreater(sym["box"][1], L.ROW_FT, name)

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
        snap["levels"] = [{"name": "FIRST", "elev": 0.0},
                          {"name": "ROOF", "elev": 18.0}]
        sch = S.build(snap)
        out = L.layout(sch)
        self.assertEqual(L.all_clashes(out["prims"]), [])
        self.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
        self.assertTrue(out["log"], "expected a stretch or a row move")


EXISTING_LABEL = '2"G, 60' + "'"


class TieIn(unittest.TestCase):
    """Start at a new pipe on an existing building (G6 shape): existing
    meter, the entered existing run (thin, labelled), the squiggle to the
    existing appliances, then the new piping."""

    def _snap(self):
        nodes = [(-50, "meter", [0, 0, 9.5]), (2, "tee", [10, 0, 9.5]),
                 (5, "fixture", [10, 8, 2.0], {"name": "WH-2", "load_mbh": 150.0}),
                 (6, "fixture", [30, 0, 9.5], {"name": "RTU-9", "load_mbh": 120.0})]
        edges = [(50, -50, 2, 10.0, "1-1/4", 270.0),
                 (51, 2, 5, 15.0, "1", 150.0), (52, 2, 6, 20.0, "1", 120.0)]
        snap = _snap(nodes, edges, path=[-50, 50, 2, 52, 6])
        snap["levels"] = [{"name": "FIRST", "elev": 0.0}]
        snap["existing"] = {"size_nominal": "2", "length_ft": 60.0,
                            "load_mbh": 400.0}
        snap["origin_kind"] = "tie_in"
        return snap

    def test_existing_run_drawn_and_counted(self):
        sch = S.build(self._snap())
        self.assertEqual(sch["existing"]["label"], [EXISTING_LABEL, "670 MBH"])
        root_len = sum(s["length_ft"] for s in S.trunk_segments(sch))
        self.assertAlmostEqual(S.trunk_sum_ft(sch), root_len + 60.0)
        out = L.layout(sch)
        self.assertEqual(L.all_clashes(out["prims"]), [])
        self.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
        existing = [p for p in out["prims"] if p["t"] == "line"
                    and p["seg"] is None and not p["new"]]
        # the existing run, its drop to the squiggle, and the existing
        # meter's service run and drop (thin, Colin 2026-11-18)
        self.assertEqual(len(existing), 4)
        self.assertEqual([p for p in out["prims"] if p["t"] == "line"
                          and p["seg"] is None and p["new"]], [])
        texts = [p["lines"] for p in out["prims"] if p["t"] == "text"
                 and p["role"] == "existing"]
        self.assertIn(["TO EXISTING", "APPLIANCES"], texts)
        self.assertIn([EXISTING_LABEL, "670 MBH"], texts)
        self.assertEqual(len(_of(out, t="symbol", role="squiggle")), 2)


class Squiggle(unittest.TestCase):
    """G1 2026-11-18: the squiggle drew flat, below its line, with a gap.
    It hangs from the line end along the line."""

    def test_across_the_line_end(self):
        # Colin 2026-11-19: turned across the line, the line runs into its
        # center
        out = L.layout(G3C_SCH)
        sq = _of(out, t="symbol", role="squiggle")
        self.assertEqual(len(sq), 1)
        sq = sq[0]
        self.assertEqual(sq["rot"], 90)
        b = sq["box"]
        self.assertGreater(b[2] - b[0], b[3] - b[1])      # wide, across
        drops = [p for p in out["prims"] if p["t"] == "line"
                 and p["seg"] is None and abs(p["x0"] - p["x1"]) < 1e-9
                 and abs(p["x0"] - sq["x"]) < 1e-9]
        self.assertEqual(len(drops), 1)
        self.assertAlmostEqual(min(drops[0]["y0"], drops[0]["y1"]), sq["y"])
        self.assertEqual(L.all_clashes(out["prims"]), [])

    def test_new_meter_service_is_new(self):
        out = L.layout(G3C_SCH)
        service = [p for p in out["prims"] if p["t"] == "line"
                   and p["seg"] is None]
        self.assertTrue(service)
        self.assertTrue(all(p["new"] for p in service))


if __name__ == "__main__":
    unittest.main()
