# -*- coding: ascii -*-
"""Tests for the pure part of oneline_draw.py (One-Line rebuild step 6,
2026-11-18): symbol family order and paths, text type choice, fallback
geometry. The Revit calls (load, place, rotate, text, view) are checked by
the live G3C run (step 8), not here.

Run:  python -I _test_oneline_draw.py
"""

import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import oneline_draw as D        # noqa: E402
import oneline_layout as L      # noqa: E402
import oneline_schematic as S   # noqa: E402


class Symbols(unittest.TestCase):

    def test_every_role_has_families(self):
        roles = [r for r, _n in D.SYMBOL_FAMILIES]
        self.assertEqual(sorted(roles), sorted(
            ["meter", "equipment", "prv", "iso", "cap", "squiggle"]))
        self.assertEqual(dict(D.SYMBOL_FAMILIES)["meter"],
                         ["RJA - P Symbols - Meter Annotation"])
        self.assertEqual(dict(D.SYMBOL_FAMILIES)["prv"][0],
                         "RJA - P Symbols - Pressure Regulating Valve")

    def test_candidates_exact_file_only(self):
        seen = []

        def exists(p):
            seen.append(p)
            return p.endswith("Pressure Reducing Valve.rfa")
        c = D.symbol_candidates("prv", r"C:\lib", exists)
        self.assertEqual([n for n, _p in c],
                         ["RJA - P Symbols - Pressure Regulating Valve",
                          "RJA - P Symbols - Pressure Reducing Valve"])
        self.assertIsNone(c[0][1])
        self.assertTrue(c[1][1].endswith("Pressure Reducing Valve.rfa"))
        for p in seen:
            self.assertNotIn(".0001", p)

    def test_real_library_has_the_new_2022_symbols(self):
        if not os.path.isdir(D.SYMBOL_DIR):
            self.skipTest("REVIT FAM not on this machine")
        for role in ("meter", "equipment", "prv", "iso", "cap", "squiggle"):
            paths = [p for _n, p in D.symbol_candidates(role)]
            self.assertTrue(any(paths), role)


class TextType(unittest.TestCase):

    def test_pick(self):
        t = [(1, "Schedule Default", 0.125 / 12), (2, '3/32" Arial - Narrow',
                                                   0.09375 / 12),
             (3, '3/32" Arial', 0.09375 / 12)]
        self.assertEqual(D.pick_text_type(t), 3)
        self.assertEqual(D.pick_text_type(t[:2]), 2)
        self.assertEqual(D.pick_text_type(t[:1]), 1)
        self.assertIsNone(D.pick_text_type([]))


class BorderAndCentering(unittest.TestCase):

    def test_border_type_name(self):
        self.assertEqual(D.border_type_name('3/32" Arial'),
                         '3/32" Arial - Border')
        self.assertEqual(D.border_type_name('3/32" Arial - Border'),
                         '3/32" Arial - Border')
        self.assertEqual(D.border_type_name(None), '3/32" Arial - Border')

    def test_center_shift_g1_equipment(self):
        # G1 2026-11-18: inserted at (20, 19.44), the family drew down to
        # 18.327 from its top line; the shift moves its center onto 19.44
        dx, dy = D.center_shift(20.0, 19.44, (18.991, 18.327, 21.009, 19.44))
        self.assertAlmostEqual(dx, 0.0, 3)
        self.assertAlmostEqual(dy, 0.5565, 3)

    def test_center_shift_already_centered(self):
        self.assertEqual(D.center_shift(1.0, 2.0, (0.0, 1.0, 2.0, 3.0)),
                         (0.0, 0.0))

    def test_prv_and_cap_not_recentered(self):
        self.assertNotIn("prv", D.CENTER_BY_BOX)
        self.assertNotIn("cap", D.CENTER_BY_BOX)
        self.assertIn("equipment", D.CENTER_BY_BOX)


class Fallbacks(unittest.TestCase):

    def test_equipment_three_lines_rotated(self):
        segs = D.fallback_segments({"role": "equipment", "x": 10.0, "y": 5.0,
                                    "rot": 90})
        self.assertEqual(len(segs), 3)
        for s in segs:
            self.assertAlmostEqual(s[0], s[2], places=6)       # vertical
        xs = sorted(round(s[0], 3) for s in segs)
        self.assertEqual(xs, [9.44, 10.0, 10.56])

    def test_cap_on_insertion_point(self):
        segs = D.fallback_segments({"role": "cap", "x": 10.15, "y": 0.0,
                                    "ix": 10.0, "iy": 0.0, "rot": 90})
        self.assertEqual(len(segs), 1)
        self.assertAlmostEqual(segs[0][0], 10.0)
        self.assertAlmostEqual(segs[0][2], 10.0)

    def test_valves_and_meter(self):
        self.assertEqual(len(D.fallback_segments({"role": "iso", "x": 0,
                                                  "y": 0, "rot": 0})), 4)
        self.assertEqual(len(D.fallback_segments({"role": "prv", "x": 0,
                                                  "y": 0, "rot": 0})), 6)
        self.assertEqual(D.fallback_segments({"role": "meter", "x": 0,
                                              "y": 0}), [])

    def test_text_value(self):
        self.assertEqual(D.text_value({"lines": ["RTU-1", "180 MBH"]}),
                         "RTU-1\r180 MBH")


class LayoutFeedsDraw(unittest.TestCase):
    """Every primitive the layout makes is one the draw layer knows."""

    def test_g3c_prims(self):
        f = open(os.path.join(HERE, "_fixtures", "oneline_g3c.json"))
        try:
            snap = json.load(f)
        finally:
            f.close()
        out = L.layout(S.build(snap))
        roles = dict(D.SYMBOL_FAMILIES)
        for p in out["prims"]:
            self.assertIn(p["t"], ("line", "text", "symbol"))
            if p["t"] == "symbol":
                self.assertIn(p["role"], roles)
                self.assertIn(p.get("rot") or 0, (0, 90, 180, 270))
            if p["t"] == "text":
                self.assertIn(p["align"], ("center", "left"))
        caps = [p for p in out["prims"] if p["t"] == "symbol"
                and p["role"] == "cap"]
        self.assertTrue(caps and "ix" in caps[0])


if __name__ == "__main__":
    unittest.main()
