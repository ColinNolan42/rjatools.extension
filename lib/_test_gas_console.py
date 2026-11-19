# -*- coding: ascii -*-
"""Tests for gas_console.py One-Line and Size Gas sections and
console_report.emit (One-Line rebuild step 5, 2026-11-18; Size Gas sections
2026-11-19). Shape and text checks in the
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

    def test_oneline_existing_row(self):
        info = {"meter": "New pipe at existing (1)", "table": "T",
                "table_id": "402.4(2)", "heat_content": 840.0,
                "origin_kind": "tie_in",
                "existing": "New construction. The system starts at the "
                            "picked pipe, with no existing piping upstream."}
        self.assertIn("Existing (entered)", gc.oneline_inputs(info))
        self.assertIn("New construction", gc.oneline_inputs(info))
        self.assertIn("**Existing (entered):**", gc.oneline_inputs_md(info))
        del info["existing"]
        self.assertNotIn("Existing (entered)", gc.oneline_inputs(info))

    def test_sizegas_new_construction_row(self):
        info = {"existing": "New construction. The system starts at the "
                            "picked pipe, with no existing piping upstream."}
        self.assertIn("starts at the picked pipe",
                      gc.sizegas_inputs(info))

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
        self.assertIn("never loads families", html)
        self.assertNotIn("GENERIC FAMILIES", html)
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


# ---------------------------------------------------------------------------
# Size Gas sections (2026-11-19). G3-like sample data, shared with the
# static preview script (import _test_gas_console and call the functions).
# ---------------------------------------------------------------------------

def sample_result():
    return {
        "sizes": {}, "table_id": "402.4(2)", "table_length_used_ft": 125,
        "longest_run_ft": 117.0, "pipe_material": "Schedule 40 Steel",
        "inlet_pressure_psi": 0.7, "heat_content_btu_per_cf": 840.0,
        "downstream_table_id": None,
        "zones": [{"label": "Meter", "table_id": "402.4(2)", "run_ft": 117.0,
                   "row_ft": 125, "farthest_name": "RTU-1"}],
    }


def _row(pid, serves, mbh, size, cap, length, status="OK", note="",
         units="CFH", zone="Meter", table="402.4(2)"):
    return {"pipe_id": pid, "serves": serves, "mbh": mbh,
            "cfh": None if units == "MBH" else mbh * 1000.0 / 840.0,
            "units": units, "size": size, "capacity": cap,
            "length_ft": length, "status": status, "note": note,
            "table_id": table, "zone": zone}


def sample_rows():
    return [
        _row(3634222, "trunk (7: OVEN-1, CB-1, GRIDDLE-1 ...)", 675.0, "2",
             1020.0, 3.2),
        _row(3617974, "OVEN-1", 60.0, "3/4", 92.0, 1.1),
        _row(3634234, "trunk (3: FRY-2, FRY-1, UNNAMED)", 210.0, "1-1/4",
             355.0, 2.7),
        _row(3618019, "FRY-1", 0.0, "3/4", 92.0, 1.1, "MIN SIZE",
             "zero demand - minimum size assigned"),
        _row(3618031, "trunk (3: RTU-2, RTU-1, MAU-1)", 700.0, "2", 1020.0,
             5.4),
    ]


def sample_fit():
    return {"resized": 34, "existing": 0, "no_sized_pipe": 1, "appliance": 0,
            "replaced": 0, "transitions": 11, "transitions_removed": 2,
            "no_element": 0, "failures": ["x"], "failure_rows": []}


class SizeGasSections(unittest.TestCase):

    INFO = {"lib": "C:\\ext\\lib", "material": "Schedule 40 Steel",
            "table": "Sched 40 Steel - Natural, <2 psi, 0.5\" w.c.",
            "table_id": "402.4(2)", "gas": "Natural", "heat_content": 840.0,
            "downstream_table": None, "downstream_table_id": None,
            "selected": "Element ID 3593143",
            "check": "Meter validation passed.", "existing": None}

    def test_inputs(self):
        p, _rows_ = _cells(gc.sizegas_inputs(self.INFO))
        self.assertEqual(p.errors, [])
        text = "".join(p.text)
        for want in ("Schedule 40 Steel", "Natural", "840 BTU/CF",
                     "[402.4(2)]", "Element ID 3593143",
                     "Meter validation passed."):
            self.assertIn(want, text)
        self.assertNotIn("Downstream", text)
        self.assertNotIn("Existing (entered)", text)
        info = dict(self.INFO, downstream_table="Low", downstream_table_id="X",
                    existing="2\" pipe, 40.0 ft")
        html = gc.sizegas_inputs(info)
        self.assertIn("Downstream of PRV", html)
        self.assertIn("Existing (entered)", html)
        self.assertIn("**Material:** Schedule 40 Steel",
                      gc.sizegas_inputs_md(self.INFO))

    def test_basis(self):
        res = sample_result()
        p, _r = _cells(gc.sizing_basis(res, 1694.0, 12, "CFH"))
        self.assertEqual(p.errors, [])
        text = "".join(p.text)
        self.assertIn("125 ft  (longest run 117.0 ft rounded up)", text)
        self.assertIn("840 BTU/CF", text)
        self.assertIn("1694.0 MBH, 12 fixtures", text)
        self.assertIn("0.7 psi", text)
        mbh = gc.sizing_basis(res, 100.0, 1, "MBH")
        self.assertIn("not used", mbh)
        self.assertNotIn("840 BTU/CF", mbh)
        res["downstream_table_id"] = "402.4(5)"
        res["zones"].append({"label": "After PRV 3", "table_id": "402.4(5)",
                             "run_ft": 40.0, "row_ft": 40,
                             "farthest_name": "GUH-1"})
        html = gc.sizing_basis(res, 100.0, 2, "CFH")
        self.assertIn("After PRV 3", html)
        self.assertIn("402.4(5)", gc.sizing_basis_md(res, 100.0, 2, "CFH"))

    def test_traversal(self):
        info = {"nodes": 51, "segments": 36, "existing_connections": []}
        p, _r = _cells(gc.traversal(info))
        self.assertEqual(p.errors, [])
        self.assertIn("51", "".join(p.text))
        self.assertNotIn(cr.CALLOUT_STYLES["info"][1], gc.traversal(info))
        info["existing_connections"] = [11, 12]
        html = gc.traversal(info)
        self.assertIn(cr.CALLOUT_STYLES["info"][1], html)
        self.assertIn("11, 12", html)
        self.assertIn("51 nodes", gc.traversal_md(info))

    def test_prv_table(self):
        rows = [{"id": 3, "family": "RJA - PRV", "downstream_ft": 20.0,
                 "treated": "mid-stream step down"},
                {"id": 4, "family": "Regulator", "downstream_ft": 3.0,
                 "treated": "equipment PRV, ignored (within 10 ft)"}]
        p, cells = _cells(gc.prv_table(rows))
        self.assertEqual(p.errors, [])
        self.assertEqual(cells[1], ["3", "RJA - PRV", "20.0",
                                    "mid-stream step down"])
        self.assertEqual(len(cells), 3)
        self.assertIn("| 3 | RJA - PRV |", gc.prv_table_md(rows))

    def test_sizing_table_columns_status_and_notes(self):
        html = gc.sizing_tables(sample_rows(), sample_result())
        p, cells = _cells(html)
        self.assertEqual(p.errors, [])
        self.assertEqual(cells[0], ["Pipe ID", "Serves", "MBH", "CFH", "Size",
                                    "Capacity CFH", "Length ft", "Status"])
        self.assertEqual(len(cells), 1 + 5)
        self.assertEqual(cells[1][0], "3634222")
        self.assertEqual(cells[1][2], "675.0")
        self.assertEqual(cells[1][3], "803.6")
        self.assertEqual(cells[1][4], '2"')
        self.assertEqual(cells[1][5], "1020")
        self.assertEqual(cells[1][7], "OK")
        self.assertEqual(cells[4][7], "MIN SIZE")
        text = "".join(p.text)
        self.assertIn("3618019: zero demand - minimum size assigned", text)
        self.assertIn("longest run (developed length) 117.0 ft", text)
        self.assertIn("table row 125 ft", text)
        # OK is green, MIN SIZE yellow.
        self.assertIn("background:" + cr.STATUS_COLORS["GREEN"], html)
        self.assertIn("background:" + cr.STATUS_COLORS["YELLOW"], html)
        self.assertEqual(cr.status_key("MIN SIZE"), "YELLOW")
        md = gc.sizing_tables_md(sample_rows(), sample_result())
        self.assertIn("| 3617974 | OVEN-1 |", md)

    def test_over_table_is_red_and_mbh_table_headers(self):
        rows = [_row(9, "RTU-1", 5000.0, "", None, 10.0, "OVER TABLE",
                     "no size in the table carries this load")]
        html = gc.sizing_tables(rows, sample_result())
        self.assertIn("background:" + cr.STATUS_COLORS["RED"], html)
        p, cells = _cells(html)
        self.assertEqual(cells[1][4], "-")
        self.assertEqual(cells[1][5], "-")
        self.assertEqual(cells[1][7], "OVER TABLE")
        self.assertEqual(cr.status_key("OVER TABLE"), "RED")
        mbh_rows = [_row(1, "A", 80.0, "1", 150.0, 5.0, units="MBH")]
        cells = _cells(gc.sizing_tables(mbh_rows, sample_result()))[1]
        self.assertEqual(cells[0][3], "CFH n/a")
        self.assertEqual(cells[0][5], "Capacity MBH")
        self.assertEqual(cells[1][3], "-")

    def test_one_table_per_zone(self):
        res = sample_result()
        res["downstream_table_id"] = "402.4(5)"
        res["zones"].append({"label": "After PRV 3", "table_id": "402.4(5)",
                             "run_ft": 40.0, "row_ft": 40,
                             "farthest_name": "GUH-1"})
        rows = sample_rows()[:2] + [
            _row(7, "GUH-1", 400.0, "1", 173.0, 9.0, zone="After PRV 3",
                 table="402.4(5)")]
        html = gc.sizing_tables(rows, res)
        self.assertEqual(html.count("<table"), 2)
        self.assertIn("After PRV 3: Table 402.4(5)", html)
        # A row with an unknown zone label is never dropped.
        rows[0]["zone"] = "Elsewhere"
        cells = _cells(gc.sizing_tables(rows, res))[1]
        self.assertEqual(len([c for c in cells if c[0].isdigit()]), 3)

    def test_existing_check(self):
        chk = {"size_nominal": "2", "length_ft": 40.0,
               "existing_load_mbh": 300.0, "new_load_mbh": 1694.0,
               "total_load_mbh": 1994.0, "demand_cfh": 2373.8,
               "capacity_cfh": 1020.0, "units": "CFH", "run_ft": 157.0,
               "table_id": "402.4(2)", "adequate": False}
        html = gc.existing_check(chk)
        p, _r = _cells(html)
        self.assertEqual(p.errors, [])
        self.assertIn(cr.CALLOUT_STYLES["error"][1], html)
        self.assertIn("NOT ADEQUATE", html)
        self.assertIn("1020 CFH at 157.0 ft", html)
        self.assertIn("2373.8 CFH", html)
        chk["adequate"] = True
        self.assertIn(cr.CALLOUT_STYLES["ok"][1], gc.existing_check(chk))
        chk["capacity_cfh"] = None
        self.assertIn("NOT CHECKED", gc.existing_check(chk))
        self.assertIn("NOT CHECKED", gc.existing_check_md(chk))

    def test_warning_groups_one_callout_each_with_id_tables(self):
        groups = gc.build_warning_groups(
            existing=["Existing pipe short."], prv=["PRV unchecked."],
            table_limit=["Limit 150 CFH."],
            no_load=[("FRY-1", 3618188)],
            unnamed=[(3618192, 90.0)], non_gas=[(77, "Domestic Cold")],
            bare_end=[(88, "Cap")], disconnected=[3617965], n_fixtures=12)
        self.assertEqual(len(groups), 8)
        html = gc.presize_warnings(groups)
        p, cells = _cells(html)
        self.assertEqual(p.errors, [])
        self.assertEqual(html.count("border-left:5px solid"), 8)
        self.assertEqual(html.count("<table"), 5)
        text = "".join(p.text)
        self.assertIn("1/12 fixture(s) have 0 MBH load", text)
        self.assertIn("1/12 fixture(s) have no name", text)
        self.assertIn("3618188", text)
        self.assertIn("FRY-1", text)
        self.assertIn("Domestic Cold", text)
        self.assertIn("3617965", text)
        self.assertIn("90.0", text)
        md = gc.presize_warnings_md(groups)
        self.assertIn("**0 MBH fixtures:**", md)
        self.assertIn("3618188  FRY-1", md)

    def test_no_warnings_is_empty(self):
        self.assertEqual(gc.build_warning_groups(), [])
        self.assertEqual(gc.presize_warnings([]), "")

    def test_unnamed_fixture_row_shows_unnamed(self):
        g = gc.build_warning_groups(no_load=[("", 5)], n_fixtures=2)
        self.assertEqual(g[0]["rows"], [["5", "UNNAMED"]])

    def _write_info(self, failures):
        return {"pipes_sized": 23, "stubs_sized": 12, "fit": sample_fit(),
                "failures": failures}

    def test_write_results_counts_and_failures(self):
        fails = gc.failure_rows(
            [{"id": 5, "family": "Pipe", "reason": "all API approaches "
              "failed (2\")."}],
            [{"id": 3634200, "family": "RJA - Gas Shutoff Valve (Ansul)",
              "reason": "could not set size to 2\" (FAILED)."}], [])
        html = gc.write_results(self._write_info(fails))
        p, cells = _cells(html)
        self.assertEqual(p.errors, [])
        counts = dict((r[0], r[1]) for r in cells if len(r) == 2)
        self.assertEqual(counts["Distribution pipes sized"], "23")
        self.assertEqual(counts["Fixture stub pipes sized"], "12")
        self.assertEqual(counts["Fittings and caps resized"], "34")
        self.assertEqual(counts["Transitions left to Revit"], "11")
        removed = [k for k in counts if k.startswith("Transitions removed")]
        self.assertEqual(len(removed), 1)
        self.assertIn("pipes now the same size", removed[0])
        self.assertEqual(counts[removed[0]], "2")
        self.assertIn(["Element", "Family", "Reason"], cells)
        self.assertIn("RJA - Gas Shutoff Valve (Ansul)", "".join(p.text))
        self.assertIn(cr.CALLOUT_STYLES["error"][1], html)
        self.assertIn("2 element(s) could not be written", html)
        self.assertIn("| 3634200 |", gc.write_results_md(
            self._write_info(fails)))

    def test_write_results_transitions_are_not_failures(self):
        html = gc.write_results(self._write_info([]))
        self.assertIn("No write failures.", html)
        self.assertIn(cr.CALLOUT_STYLES["ok"][1], html)
        self.assertNotIn("Element</th>", html)
        self.assertIn("Revit updates or removes them", html)

    def test_write_start(self):
        html = gc.write_start("A, B")
        self.assertIn("Writing Sizes", html)
        self.assertIn("A, B", gc.write_start_md("A, B"))

    def test_existing_audit(self):
        self.assertEqual(gc.existing_audit(0, [], []), "")
        self.assertEqual(gc.existing_audit_md(0, [], []), "")
        ok = gc.existing_audit(3, [], [])
        self.assertIn(cr.CALLOUT_STYLES["ok"][1], ok)
        over = [{"pipe_id": 5, "current_nom": "1", "demand_mbh": 300.0,
                 "current_cap": 173.0, "units": "CFH", "recommended_nom": "2"},
                {"pipe_id": 6, "current_nom": "1", "demand_mbh": 90.0,
                 "current_cap": None, "units": "CFH", "recommended_nom": "1"}]
        html = gc.existing_audit(
            4, over, [{"pipe_id": 9, "table_id": "402.4(2)"}])
        p, cells = _cells(html)
        self.assertEqual(p.errors, [])
        self.assertEqual(cells[1], ["5", '1"', "300.0", "173.0 CFH", '2"'])
        self.assertEqual(cells[2][3], "none published")
        self.assertIn("9 [Table 402.4(2)]", html)
        self.assertIn("2 undersized", gc.existing_audit_md(4, over, []))

    def _summary_info(self, fail_count=0):
        return {"success": 23, "stub_success": 12, "stub_fail": 0,
                "existing_skip": 0, "existing_overloaded": 0,
                "existing_unchecked": 0, "fit": sample_fit(),
                "fail_count": fail_count, "approach": "RBS_PIPE_DIAMETER_PARAM",
                "table_id": "402.4(2)", "heat_content": 840.0,
                "existing_check": None, "zones": sample_result()["zones"],
                "downstream": False, "longest_run": 117.0, "table_row": 125}

    def test_summary_total_row_and_values(self):
        html = gc.summary(self._summary_info())
        p, cells = _cells(html)
        self.assertEqual(p.errors, [])
        self.assertEqual(cells[-1], ["Pipes, stubs and fittings sized", "69"])
        d = dict((r[0], r[1]) for r in cells if len(r) == 2)
        self.assertEqual(d["Transitions left to Revit"], "11")
        self.assertEqual(d["Fitting failures"], "1")
        self.assertEqual(d["Pipe failures"], "0")
        self.assertEqual(d["Longest run"], "117.0 ft")
        self.assertEqual(d["Table row used"], "125 ft")
        self.assertEqual(d["API approach used"], "RBS_PIPE_DIAMETER_PARAM")
        self.assertIn(cr.CALLOUT_STYLES["ok"][1], html)
        self.assertIn("23 distribution pipes + 12 stub pipes + 34",
                      gc.summary_md(self._summary_info()))

    def test_summary_failure_and_zoned(self):
        info = self._summary_info(fail_count=2)
        info["downstream"] = True
        info["zones"] = sample_result()["zones"] + [
            {"label": "After PRV 3", "table_id": "402.4(5)", "run_ft": 40.0,
             "row_ft": 40, "farthest_name": "GUH-1"}]
        info["existing_check"] = {
            "size_nominal": "2", "length_ft": 40.0, "existing_load_mbh": 300.0,
            "adequate": True}
        html = gc.summary(info)
        self.assertIn(cr.CALLOUT_STYLES["warn"][1], html)
        self.assertIn("2/25 pipes failed", html)
        self.assertIn("After PRV 3", html)
        self.assertNotIn("Table row used", html)
        self.assertIn("adequate", html)

    def test_message(self):
        self.assertIn(cr.CALLOUT_STYLES["error"][1],
                      gc.message("Traversal ERROR: x", "error"))
        self.assertEqual(gc.message_md("plain", "warn"), "plain")

    def test_every_section_html_parses(self):
        res = sample_result()
        for html in (gc.sizegas_inputs(self.INFO),
                     gc.sizing_basis(res, 1.0, 1, "CFH"),
                     gc.sizing_tables(sample_rows(), res),
                     gc.summary(self._summary_info())):
            self.assertEqual(shape(html).errors, [])


if __name__ == "__main__":
    unittest.main()
