# -*- coding: ascii -*-
"""Tests for gas_console.py One-Line sections and console_report.emit
(One-Line rebuild step 5, 2026-11-18). Shape and text checks in the
DuctPerfSectionTests pattern, on the G3C and G2 fixtures.

Run:  python -I _test_gas_console.py
"""

import copy
import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import console_report as cr        # noqa: E402
import gas_console as gc           # noqa: E402
import oneline_gate                # noqa: E402
import oneline_layout as L         # noqa: E402
import oneline_schematic as S      # noqa: E402
from _test_console_report import shape   # noqa: E402


def _load(name):
    f = open(os.path.join(HERE, "_fixtures", name))
    try:
        return json.load(f)
    finally:
        f.close()


G3C = _load("oneline_g3c.json")
G2 = _load("oneline_g2.json")
SCH = S.build(G3C)


def _cells(html):
    p = shape(html)
    return p, [[c[2] for c in r] for r in p.rows]


class Emit(unittest.TestCase):

    class _Out(object):
        def __init__(self, fail=False):
            self.fail = fail
            self.html = []

        def print_html(self, h):
            if self.fail:
                raise RuntimeError("no browser")
            self.html.append(h)

    class _Log(object):
        def __init__(self):
            self.lines = []

        def warning(self, t):
            self.lines.append(("w", t))

        def error(self, t):
            self.lines.append(("e", t))

    def test_prints_html(self):
        out, fb = self._Out(), []
        self.assertTrue(cr.emit(out, "<b>x</b>", lambda: fb.append(1)))
        self.assertEqual(out.html, ["<b>x</b>"])
        self.assertEqual(fb, [])

    def test_fallback_logged(self):
        out, fb, log = self._Out(fail=True), [], self._Log()
        self.assertFalse(cr.emit(out, "<b>x</b>", lambda: fb.append(1), log,
                                 "One-Line"))
        self.assertEqual(fb, [1])
        self.assertEqual(log.lines[0][0], "w")
        self.assertIn("One-Line", log.lines[0][1])

    def test_fallback_failure_logged(self):
        def boom():
            raise ValueError("bad")
        log = self._Log()
        cr.emit(self._Out(fail=True), "x", boom, log)
        self.assertEqual([k for k, _t in log.lines], ["w", "e"])


class Sections(unittest.TestCase):

    def test_inputs(self):
        info = {"meter": "RJA - Gas Meter (3618292)", "table":
                "Sched 40 Steel - Natural, <2 psi, 0.5\" w.c.",
                "table_id": "402.4(2)", "heat_content": 840,
                "downstream_table": None}
        p, _rows = _cells(gc.oneline_inputs(info))
        self.assertEqual(p.errors, [])
        text = "".join(p.text)
        self.assertIn("840 BTU/CF", text)
        self.assertIn("[402.4(2)]", text)
        self.assertNotIn("Downstream", text)
        info["downstream_table"] = "402.4(2)"
        self.assertIn("Downstream of PRV", gc.oneline_inputs(info))
        self.assertIn("**IFGC table:**", gc.oneline_inputs_md(info))

    def test_gate_ok(self):
        html = gc.gate(oneline_gate.check(G3C))
        self.assertEqual(shape(html).errors, [])
        self.assertIn(cr.CALLOUT_STYLES["ok"][1], html)
        self.assertIn("complete system", gc.gate_md(oneline_gate.check(G3C)))

    def test_gate_reject_table(self):
        res = oneline_gate.check(G2)
        p, rows = _cells(gc.gate(res))
        self.assertEqual(p.errors, [])
        self.assertEqual(rows[0], ["Check", "Element", "Reason"])
        self.assertEqual(len(rows), 1 + 32)
        self.assertEqual(rows[1][0], "Size differs")
        self.assertIn("Run Size Gas first", "".join(p.text))
        self.assertIn(cr.CALLOUT_STYLES["error"][1], gc.gate(res))
        md = gc.gate_md(res)
        self.assertIn("Run Size Gas first", md)
        self.assertEqual(md.count("Size differs"), 32)

    def test_gate_warning_callout(self):
        s = copy.deepcopy(G3C)
        s["non_gas_skipped"] = [[77, "Domestic Cold Water"]]
        html = gc.gate(oneline_gate.check(s))
        self.assertIn(cr.CALLOUT_STYLES["warn"][1], html)

    def test_symbols_status_colors(self):
        rows = [{"role": "Meter", "family": "RJA - P Symbols - Meter Annotation",
                 "status": "loaded"},
                {"role": "Equipment", "family": "RJA - P Symbols - Equipment",
                 "status": "found"},
                {"role": "Cap", "family": "RJA- Equipment Cap",
                 "status": "fallback", "reason": "file not found"}]
        html = gc.symbols_table(rows)
        p, cells = _cells(html)
        self.assertEqual(p.errors, [])
        self.assertEqual([r[2] for r in cells[1:]],
                         ["OK, loaded", "OK, found", "FALLBACK, drawn"])
        styles = [r[2][1] for r in p.rows[1:]]
        self.assertIn(cr.STATUS_COLORS["GREEN"], styles[0])
        self.assertIn(cr.STATUS_COLORS["YELLOW"], styles[2])
        self.assertIn("GENERIC FAMILIES", html)
        self.assertIn("FALLBACK, drawn", gc.symbols_table_md(rows))

    def test_summary_checks_pass_on_g3c(self):
        html = gc.oneline_summary(SCH, G3C, "Gas One-Line - test")
        p, rows = _cells(html)
        self.assertEqual(p.errors, [])
        checks = [r for r in rows if len(r) == 3 and r[0] != "Check"]
        self.assertEqual([r[2] for r in checks], ["OK", "OK", "OK"])
        self.assertIn("116.96 of 116.96 ft", checks[0][1])
        self.assertIn("44 of 44 pipes", checks[1][1])
        text = "".join(p.text)
        self.assertIn("1814 MBH", text)
        self.assertIn("117'", text)
        self.assertIn("11 (11 PRV symbols)", text)

    def test_summary_flags_a_broken_tree(self):
        sch = copy.deepcopy(SCH)
        sid = [s for s in sch["segments"].values() if s["on_longest"]][0]["id"]
        sch["segments"][sid]["length_ft"] += 10.0
        sch["segments"][sid]["eids"].append(sch["segments"][sid]["eids"][0])
        rows = _cells(gc.oneline_summary(sch, G3C))[1]
        checks = [r for r in rows if len(r) == 3 and r[0] != "Check"]
        self.assertEqual([r[2] for r in checks][:2], ["FAIL", "FAIL"])

    def test_length_check_total(self):
        p, rows = _cells(gc.length_check(SCH, G3C))
        self.assertEqual(p.errors, [])
        self.assertEqual(rows[-1][0], "Total")
        self.assertEqual(rows[-1][-1], "116.96")
        self.assertEqual(len(rows), 1 + 5 + 1)

    def test_warnings(self):
        self.assertEqual(gc.oneline_warnings([]), "")
        html = gc.oneline_warnings(["Segment S3 has more than one size: 2, 1"])
        self.assertIn(cr.CALLOUT_STYLES["warn"][1], html)
        self.assertEqual(gc.oneline_warnings_md(["a"]), "- a")

    def test_layout_diagnostic(self):
        lay = L.layout(SCH)
        text = gc.layout_diagnostic_text(SCH, lay)
        self.assertIn("ONE-LINE LAYOUT DIAGNOSTIC", text)
        self.assertIn('1"G, 6\'', text)
        self.assertIn("-> FUTURE", text)
        self.assertIn("trunk sum 116.96 ft", text)
        html = gc.layout_diagnostic(SCH, lay)
        self.assertEqual(shape(html).errors, [])
        self.assertIn("<pre", html)
        self.assertIn("&quot;G", html)      # escaped inch marks


if __name__ == "__main__":
    unittest.main()
