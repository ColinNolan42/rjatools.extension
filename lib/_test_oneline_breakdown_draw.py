# -*- coding: ascii -*-
"""CPython tests for the Developed Length Breakdown drawing and console
section (Colin, 2026-11-20) and for the view scale support in rja_drafting.

rja_drafting imports Autodesk.Revit.DB at module scope, so a stub package is
put in sys.modules first (the same technique as _test_duct_perf_sheet.py);
the drawing primitives (_t, _fill, _line) are replaced with recorders so the
real layout code runs and every text and fill is captured.

Run:  python _test_oneline_breakdown_draw.py
"""

import json
import os
import sys
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def _stub_revit():
    names = ('XYZ Line TextNote TextNoteOptions TextNoteType ViewDrafting '
             'ViewFamily ViewFamilyType Viewport FilteredElementCollector '
             'BuiltInCategory CurveLoop FilledRegion FilledRegionType '
             'OverrideGraphicSettings ElementId Color FillPatternElement '
             'BuiltInParameter HorizontalTextAlignment').split()
    pkg = types.ModuleType('Autodesk')
    revit = types.ModuleType('Autodesk.Revit')
    db = types.ModuleType('Autodesk.Revit.DB')
    for n in names:
        setattr(db, n, type(n, (object,), {}))
    pkg.Revit = revit
    revit.DB = db
    sys.modules.setdefault('Autodesk', pkg)
    sys.modules.setdefault('Autodesk.Revit', revit)
    sys.modules.setdefault('Autodesk.Revit.DB', db)


_stub_revit()
import rja_drafting as rd            # noqa: E402
import oneline_draw                  # noqa: E402
import oneline_schematic as S        # noqa: E402
import gas_console                   # noqa: E402


def _load(name):
    f = open(os.path.join(HERE, "_fixtures", name))
    try:
        return json.load(f)
    finally:
        f.close()


G3C = _load("oneline_g3c.json")


class Style(object):
    """Stand in for rja_drafting._Style."""

    def __init__(self, scale=1.0):
        self.scale = scale
        self.types = {}


class View(object):
    def __init__(self, scale):
        self.Scale = scale
        self.Id = 1


class Recorder(object):

    def __init__(self):
        self.texts = []    # (x, y, value, kind, align)
        self.fills = []    # (x0, top, x1, bottom, rgb)
        self.lines = []

    def install(self):
        self._saved = (rd._t, rd._fill, rd._line)
        rec = self

        def _t(st, view, x, y, value, kind="text", align="left"):
            if value:
                rec.texts.append((x, y, str(value), kind, align))

        def _fill(st, view, x0, top, x1, bottom, rgb):
            rec.fills.append((x0, top, x1, bottom, rgb))

        def _line(st, view, x0, y0, x1, y1, weight=None):
            rec.lines.append((x0, y0, x1, y1))
        rd._t, rd._fill, rd._line = _t, _fill, _line
        return self

    def remove(self):
        rd._t, rd._fill, rd._line = self._saved


class ScaleTests(unittest.TestCase):

    def test_scale_one_is_unchanged(self):
        self.assertAlmostEqual(rd.text_width("abcd"),
                               4 * rd.CHAR_W + 2 * rd.PAD)
        self.assertAlmostEqual(rd.text_width("abcd", scale=1.0),
                               rd.text_width("abcd"))
        needs, band = rd.column_needs(["A", "BB"], [["x", "y"]])
        needs1, band1 = rd.column_needs(["A", "BB"], [["x", "y"]], scale=1.0)
        self.assertEqual((needs, band), (needs1, band1))
        self.assertEqual(rd._k(None), 1.0)
        self.assertEqual(rd._k(Style()), 1.0)

    def test_scale_100_scales_text_width_and_needs(self):
        self.assertAlmostEqual(rd.text_width("abcd", False, 100),
                               100 * rd.text_width("abcd"))
        self.assertAlmostEqual(rd.text_width("abcd", True, 100),
                               100 * rd.text_width("abcd", True))
        n1, b1 = rd.column_needs(["ITEM", "LEN"], [["abc", "1"],
                                                    {"band": "xyz"}], True)
        n100, b100 = rd.column_needs(["ITEM", "LEN"], [["abc", "1"],
                                                       {"band": "xyz"}], True,
                                     scale=100)
        for a, b in zip(n1, n100):
            self.assertAlmostEqual(a * 100, b)
        self.assertAlmostEqual(b1 * 100, b100)
        self.assertAlmostEqual(rd.table_need(["A"], [["abc"]], False, 100),
                               100 * rd.table_need(["A"], [["abc"]]))

    def _rows_y(self, scale):
        rec = Recorder().install()
        try:
            rd._styled_table(Style(scale), View(1), 0.0, 0.0,
                             [("ITEM", 1.0), ("LEN", 1.0)],
                             [["r0", "1"], ["r1", "2"], ["r2", "3"]],
                             totals_row=True, fit=True)
        finally:
            rec.remove()
        ys = sorted(set(round(t[1], 9) for t in rec.texts if t[2] != "ITEM"
                        and t[2] != "LEN"), reverse=True)
        return ys

    def test_row_height_scales_with_the_view(self):
        y1 = self._rows_y(1.0)
        y100 = self._rows_y(100.0)
        self.assertEqual(len(y1), 3)
        self.assertAlmostEqual(y1[0] - y1[1], rd.ROW_H)
        self.assertAlmostEqual(y100[0] - y100[1], rd.ROW_H * 100)
        self.assertAlmostEqual(y100[0], y1[0] * 100)

    def test_band_text_inset_scales(self):
        rec = Recorder().install()
        try:
            bottom = rd._band(Style(100.0), View(100), 0.0, 0.0, 50.0,
                              "TITLE", None, rd.BAND_H * 100)
        finally:
            rec.remove()
        self.assertAlmostEqual(bottom, -rd.BAND_H * 100)
        self.assertEqual(rec.texts[0][2], "TITLE")


class DrawBreakdown(unittest.TestCase):

    def _draw(self, bd, scale=100.0, bounds=(10.0, -20.0, 400.0, 150.0)):
        rec = Recorder().install()
        try:
            bottom, width = oneline_draw.draw_breakdown(
                None, View(scale), bd, bounds, style=Style(scale))
        finally:
            rec.remove()
        return rec, bottom, width

    def _bd(self):
        sch = S.build(G3C)
        return S.developed_breakdown(sch, G3C["longest_run_ft"])

    def test_mark_band_and_table_layout(self):
        bd = self._bd()
        rec, bottom, width = self._draw(bd)
        values = [t[2] for t in rec.texts]
        self.assertEqual(values[0], "DELETE BEFORE PUTTING ON SCHEDULE")
        self.assertEqual(rec.texts[0][3], "bold")
        self.assertIn("LONGEST RUN TO RTU-1", values)
        for row in S.breakdown_table(bd)[1]:
            self.assertIn(row[0], values)
        # starts at the diagram's left edge, under the diagram bottom
        self.assertAlmostEqual(rec.texts[0][0], 10.0)
        self.assertLess(rec.texts[0][1], -20.0)
        self.assertLess(bottom, -20.0)

    def test_band_width_equals_table_width(self):
        bd = self._bd()
        rec, bottom, width = self._draw(bd)
        blue = [f for f in rec.fills if f[4] == rd.RJA_BLUE]
        self.assertEqual(len(blue), 1)
        band_w = blue[0][2] - blue[0][0]
        self.assertAlmostEqual(band_w, width)
        # the table's header fill spans the same width
        head = [f for f in rec.fills if f[4] == rd.HEAD_FILL]
        self.assertTrue(head)
        self.assertAlmostEqual(head[0][2] - head[0][0], width)
        # the width is the fitted columns, so 100x a 1:1 fit
        needs, band = rd.column_needs(*(list(S.breakdown_table(bd)) + [True]))
        self.assertGreaterEqual(width, sum(needs) * 100 - 1e-6)

    def test_long_outlet_name_widens_the_table_and_the_band(self):
        bd = self._bd()
        bd["farthest"] = "ROOFTOP UNIT NUMBER TWELVE POINT FIVE, EAST WING"
        rec, bottom, width = self._draw(bd)
        blue = [f for f in rec.fills if f[4] == rd.RJA_BLUE][0]
        self.assertAlmostEqual(blue[2] - blue[0], width)
        self.assertGreaterEqual(
            width, rd.text_width("LONGEST RUN TO " + bd["farthest"], True, 100)
            - 1e-6)

    def test_scale_one_view_draws_in_paper_units(self):
        bd = self._bd()
        rec1, b1, w1 = self._draw(bd, 1.0, (0.0, 0.0, 1.0, 1.0))
        rec100, b100, w100 = self._draw(bd, 100.0, (0.0, 0.0, 100.0, 100.0))
        self.assertAlmostEqual(w100, w1 * 100)

    def test_no_farthest_name_falls_back(self):
        bd = self._bd()
        bd["farthest"] = None
        self.assertEqual(oneline_draw.breakdown_title(bd),
                         "LONGEST RUN TO FARTHEST OUTLET")


class ConsoleSection(unittest.TestCase):

    def test_section_html_and_md(self):
        sch = S.build(G3C)
        bd = S.developed_breakdown(sch, G3C["longest_run_ft"])
        html = gas_console.developed_breakdown(bd)
        self.assertIn("Developed Length Breakdown", html)
        self.assertIn("RTU-1", html)
        self.assertIn("124.3&#x27; (124&#x27;)", html.replace("&#39;",
                                                            "&#x27;"))
        self.assertNotIn("FUTURE", html)
        md = gas_console.developed_breakdown_md(bd)
        self.assertIn("DEVELOPED LENGTH", md)
        self.assertIn("METER TO TEE 1", md)

    def test_mismatch_adds_a_warning(self):
        sch = S.build(G3C)
        bd = S.developed_breakdown(sch, G3C["longest_run_ft"] + 5.0)
        html = gas_console.developed_breakdown(bd)
        self.assertIn("different run", html)


if __name__ == "__main__":
    unittest.main()
