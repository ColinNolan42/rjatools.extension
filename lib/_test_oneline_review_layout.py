# -*- coding: ascii -*-
"""Regression tests for the 2026-11-20 review findings on oneline_layout:
n1 (drops long enough for valves and Segments), n14 (second leg the same
way while the run continues), n31 and n32 (roof rule), n43 (budget stop
message), n61 (inner transactions rolled back first, source check).

Run:  python -I _test_oneline_review_layout.py
"""

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import oneline_layout as L     # noqa: E402
import oneline_schematic as S  # noqa: E402
from _test_oneline_schematic import _snap   # noqa: E402

FIRST_ROOF = [{"name": "FIRST", "elev": 0.0}, {"name": "ROOF", "elev": 20.0}]


def _ok(test, out):
    test.assertNotIn("LAYOUT FAILED", " ".join(out["log"]))
    test.assertEqual(L.all_clashes(out["prims"]), [])


def _sym(out, role, node):
    return [p for p in out["prims"] if p["t"] == "symbol"
            and p["role"] == role and p.get("node") == node]


def _hangs_up(out, node):
    eq = _sym(out, "equipment", node)[0]
    lab = [p for p in out["prims"] if p["t"] == "text"
           and p.get("role") == "fixture" and p.get("node") == node][0]
    return lab["box"][1] >= eq["box"][3]


class DropLength(unittest.TestCase):
    """n1: a drop or the last unit's turn carrying valves or more than one
    Segment shall be long enough (both failed with LAYOUT FAILED)."""

    def _tee_drop(self, drop_nodes, drop_edges):
        nodes = [(1, "meter", [0, 0, 1.0]), (2, "tee", [10, 0, 9.0]),
                 (5, "fixture", [10, 8, 1.0], {"name": "WH-2", "load_mbh": 150.0}),
                 (6, "fixture", [30, 0, 1.0], {"name": "OVEN-9", "load_mbh": 90.0})]
        edges = [(50, 1, 2, 10.0, "1-1/4", 240.0), (52, 2, 6, 20.0, "1", 90.0)]
        snap = _snap(nodes + drop_nodes, edges + drop_edges,
                     path=[1, 50, 2, 52, 6])
        snap["levels"] = FIRST_ROOF
        return snap

    def test_ball_valve_on_a_drop(self):
        snap = self._tee_drop(
            [(3, "fitting", [10, 4, 5.0], {"family": "Ball Valve - Threaded"})],
            [(51, 2, 3, 7.0, "1", 150.0), (53, 3, 5, 8.0, "1", 150.0)])
        out = L.layout(S.build(snap))
        _ok(self, out)
        self.assertEqual(len(_sym(out, "iso", 3)), 1)

    def test_two_segments_on_a_drop(self):
        snap = self._tee_drop(
            [(3, "fitting", [10, 4, 5.0], {"is_transition": True})],
            [(51, 2, 3, 7.0, "1-1/4", 150.0), (53, 3, 5, 8.0, "1", 150.0)])
        sch = S.build(snap)
        br = S.terminal_branch(sch, "WH-2")
        self.assertEqual(len(br["segments"]), 2)
        _ok(self, L.layout(sch))

    def test_valve_on_the_last_units_turn(self):
        nodes = [(1, "meter", [0, 0, 1.0]), (2, "tee", [10, 0, 9.0]),
                 (5, "fixture", [10, 8, 1.0], {"name": "WH-2", "load_mbh": 150.0}),
                 (3, "fitting", [25, 0, 9.0], {"family": "Gas Cock Valve"}),
                 (6, "fixture", [30, 0, 1.0], {"name": "OVEN-9", "load_mbh": 90.0})]
        edges = [(50, 1, 2, 10.0, "1-1/4", 240.0), (51, 2, 5, 15.0, "1", 150.0),
                 (52, 2, 3, 15.0, "1", 90.0), (53, 3, 6, 5.0, "1", 90.0)]
        snap = _snap(nodes, edges, path=[1, 50, 2, 52, 3, 53, 6])
        snap["levels"] = FIRST_ROOF
        _ok(self, L.layout(S.build(snap)))


class SecondLegSameWay(unittest.TestCase):
    """n14: a cross with two first floor drops and the run continuing to
    another tee used to fail; the second drop leaves from the node line."""

    def test_cross_two_drops_and_run(self):
        nodes = [(1, "meter", [0, 0, 1.0]), (2, "tee", [10, 0, 9.0]),
                 (5, "fixture", [10, 8, 1.0], {"name": "WH-1", "load_mbh": 150.0}),
                 (7, "fixture", [10, -8, 1.0], {"name": "WH-2", "load_mbh": 150.0}),
                 (3, "tee", [30, 0, 9.0]),
                 (6, "fixture", [30, 8, 1.0], {"name": "F-1", "load_mbh": 90.0}),
                 (8, "fixture", [50, 0, 1.0], {"name": "F-2", "load_mbh": 90.0})]
        edges = [(50, 1, 2, 10.0, "2", 480.0), (51, 2, 5, 9.0, "1", 150.0),
                 (55, 2, 7, 9.0, "1", 150.0), (52, 2, 3, 20.0, "1-1/4", 180.0),
                 (53, 3, 6, 9.0, "1", 90.0), (54, 3, 8, 20.0, "1", 90.0)]
        snap = _snap(nodes, edges, path=[1, 50, 2, 52, 3, 54, 8])
        snap["levels"] = FIRST_ROOF
        out = L.layout(S.build(snap))
        _ok(self, out)
        node_lines = [p for p in out["prims"] if p["t"] == "line"
                      and str(p["seg"]).startswith("J")]
        self.assertTrue(node_lines)


class RoofRule(unittest.TestCase):
    """n31 / n32 (Colin 2026-11-20): roof = level name has ROOF, or units
    more than 8 ft above the top level. The meter plays no part."""

    def _two_story(self, levels, unit_z):
        nodes = [(1, "meter", [0, 0, 0.5]), (2, "tee", [10, 0, 9.0]),
                 (5, "fixture", [10, 8, 1.0], {"name": "WH-1", "load_mbh": 150.0}),
                 (3, "tee", [30, 0, unit_z - 1.0]),
                 (6, "fixture", [30, 8, unit_z], {"name": "F-1", "load_mbh": 90.0}),
                 (7, "fixture", [50, 0, unit_z], {"name": "F-2", "load_mbh": 90.0})]
        edges = [(50, 1, 2, 10.0, "2", 330.0), (51, 2, 5, 9.0, "1", 150.0),
                 (52, 2, 3, 30.0, "1-1/4", 180.0), (53, 3, 6, 9.0, "1", 90.0),
                 (54, 3, 7, 20.0, "1", 90.0)]
        snap = _snap(nodes, edges, path=[1, 50, 2, 52, 3, 54, 7])
        snap["levels"] = levels
        return snap

    def test_second_floor_hangs_down_without_roof_units(self):
        lv = [{"name": "FIRST", "elev": 0.0}, {"name": "SECOND", "elev": 14.0},
              {"name": "ROOF", "elev": 28.0}]
        out = L.layout(S.build(self._two_story(lv, 14.5)))
        _ok(self, out)
        self.assertFalse(_hangs_up(out, 6))
        self.assertFalse(_hangs_up(out, 7))

    def test_low_roof_hangs_up(self):
        lv = [{"name": "FIRST", "elev": 0.0},
              {"name": "ROOF - ONE STORY", "elev": 14.0},
              {"name": "ROOF", "elev": 28.0}]
        out = L.layout(S.build(self._two_story(lv, 15.5)))
        _ok(self, out)
        self.assertTrue(_hangs_up(out, 6))

    def test_units_8_ft_above_top_level_hang_up(self):
        lv = [{"name": "FIRST", "elev": 0.0}, {"name": "SECOND", "elev": 14.0}]
        out = L.layout(S.build(self._two_story(lv, 30.0)))
        _ok(self, out)
        self.assertTrue(_hangs_up(out, 6))

    # Colin 2026-11-20: a level named ROOF wins; the 8 ft height rule applies
    # only with no ROOF level, and each unit it lifts is listed as a note
    def test_roof_level_present_high_unit_elsewhere_hangs_down(self):
        lv = [{"name": "FIRST", "elev": 0.0},
              {"name": "LOW ROOF", "elev": 10.0},
              {"name": "SECOND", "elev": 14.0}]
        sch = S.build(self._two_story(lv, 30.0))     # 16 ft above SECOND
        out = L.layout(sch)
        _ok(self, out)
        self.assertFalse(_hangs_up(out, 6))
        self.assertFalse(_hangs_up(out, 7))
        self.assertEqual(out["roof_notes"], [])
        self.assertFalse([w for w in sch["warnings"] if "ROOF" in w])

    def test_no_roof_level_unit_12_ft_up_gets_roof_band_and_note(self):
        lv = [{"name": "FIRST", "elev": 0.0}]
        sch = S.build(self._two_story(lv, 12.0))
        out = L.layout(sch)
        _ok(self, out)
        self.assertTrue(_hangs_up(out, 6))
        self.assertTrue(_hangs_up(out, 7))
        self.assertFalse(_hangs_up(out, 5))          # WH-1 on FIRST
        notes = out["roof_notes"]
        self.assertEqual(len(notes), 2)
        joined = " ".join(notes)
        for name, eid in (("F-1", 6), ("F-2", 7)):
            self.assertIn("{0} (element {1}), 12.0 ft above the top "
                          "level (FIRST)".format(name, eid), joined)
        self.assertNotIn("WH-1", joined)
        # listed once in the warnings the One-Line button prints, even when
        # the layout runs again on the same schematic
        L.layout(sch)
        for note in notes:
            self.assertEqual(sch["warnings"].count(note), 1)

    def test_roof_tie_in(self):
        nodes = [(-50, "meter", [0, 0, 29.0]), (2, "tee", [10, 0, 29.0]),
                 (5, "fixture", [10, 8, 29.5], {"name": "RTU-1", "load_mbh": 150.0}),
                 (6, "fixture", [30, 0, 29.5], {"name": "RTU-2", "load_mbh": 120.0})]
        edges = [(50, -50, 2, 10.0, "1-1/4", 270.0),
                 (51, 2, 5, 15.0, "1", 150.0), (52, 2, 6, 20.0, "1", 120.0)]
        snap = _snap(nodes, edges, path=[-50, 50, 2, 52, 6])
        snap["levels"] = [{"name": "FIRST", "elev": 0.0},
                          {"name": "ROOF", "elev": 28.0}]
        snap["origin_kind"] = "tie_in"
        out = L.layout(S.build(snap))
        _ok(self, out)
        self.assertTrue(_hangs_up(out, 5))
        self.assertTrue(_hangs_up(out, 6))
        meter = _sym(out, "meter", -50)[0]
        for n in (5, 6):
            self.assertGreater(_sym(out, "equipment", n)[0]["y"], meter["y"])


class BudgetMessage(unittest.TestCase):
    """n43: a search stopped on the budget does not claim an overlap."""

    def test_budget_stop_says_out_of_time(self):
        msg = L.failure_message(["search stopped after 20001 placement "
                                 "checks, 20 s", "LAYOUT FAILED: no clash "
                                 "free placement found"])
        self.assertIn("ran out of time", msg)
        self.assertIn("No overlap was proven", msg)

    def test_real_failure_says_overlap(self):
        msg = L.failure_message(["LAYOUT FAILED: no clash free placement "
                                 "found"])
        self.assertIn("without overlaps", msg)
        self.assertNotIn("ran out of time", msg)


class ScriptRollback(unittest.TestCase):
    """n61: an open inner Transaction is rolled back before the group.
    Source check only; the Revit run is live tested."""

    def test_inner_rolled_back_before_group(self):
        path = os.path.join(HERE, "..", "Developer Tools.tab",
                            "Designer Tools.panel", "One-Line.pushbutton",
                            "script.py")
        src = open(path).read()
        i = src.index("except Exception as ex:\n        # an open inner")
        body = src[i:i + 900]
        self.assertLess(body.index("t.RollBack()"), body.index("tg.RollBack()"))
        self.assertNotIn("except Exception:\n            pass", body)


if __name__ == "__main__":
    unittest.main()
