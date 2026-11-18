# -*- coding: utf-8 -*-
"""
gas_console.py  --  gas tool sections for the pyRevit output window.

Pure renderers, no Revit imports, built from the lib/console_report.py
builders (Plan B, TOOL_ALIGNMENT_PLAN.md). Each section returns one HTML
string; its *_md twin returns the plain markdown the button prints when
print_html fails (console_report.emit). Tested in lib/_test_gas_console.py.

This file holds the One-Line sections (rebuild step 5, 2026-11-18). The Size
Gas sections in Plan B are a separate item and come later.

IronPython 2.7: no f-strings; numbers are cast with float() before format.
"""

import console_report as cr
import gas_labels
import oneline_schematic as S


W_SYMBOLS = [18, 44, 14, 24]
W_GATE = [20, 14, 66]
W_LENGTH = [12, 30, 16, 14, 14, 14]
SYMBOL_STATUS = {"found": "OK, found", "loaded": "OK, loaded",
                 "fallback": "FALLBACK, drawn"}


def _ft(v):
    return "{0:.2f}".format(float(v or 0.0))


# ---------------------------------------------------------------------------
# inputs
# ---------------------------------------------------------------------------

def oneline_inputs(info):
    """info: {'meter', 'table', 'table_id', 'heat_content', 'downstream_table'
    (or None), 'origin_kind' ('meter' or 'tie_in')}"""
    rows = _input_rows(info)
    return cr.section("Gas One-Line", "inputs") + cr.keyvalue(rows)


def _input_rows(info):
    rows = [("Start", info.get("meter") or ""),
            ("IFGC table", "{0}  [{1}]".format(info.get("table") or "",
                                               info.get("table_id") or "")),
            ("Heat content of gas", "{0:.0f} BTU/CF".format(
                float(info.get("heat_content") or 0.0)))]
    if info.get("downstream_table"):
        rows.append(("Downstream of PRV", info["downstream_table"]))
    if info.get("origin_kind") == "tie_in":
        rows.append(("Start type", "New pipe at the existing system"))
    return rows


def oneline_inputs_md(info):
    return "\n".join("**{0}:** {1}".format(k, v) for k, v in _input_rows(info))


# ---------------------------------------------------------------------------
# gate
# ---------------------------------------------------------------------------

def gate(res):
    """oneline_gate.check() result: an ok callout, or the reasons the
    One-Line was not drawn."""
    if res["ok"]:
        out = cr.section("System Check") + cr.callout(
            "Complete system, every pipe at its Size Gas size. Drawing the "
            "One-Line.", "ok")
    else:
        out = cr.section("System Check", "not drawn") + cr.callout(
            res["headline"] + ". The One-Line was not drawn.", "error")
        rows = [[_code_text(e["code"]), str(e["id"]), e["text"]]
                for e in res["errors"]]
        out += cr.table(["Check", "Element", "Reason"], rows,
                        align=["l", "r", "l"], widths=W_GATE)
    for w in res.get("warnings") or []:
        out += cr.callout(w["text"], "warn")
    return out


def gate_md(res):
    if res["ok"]:
        return "**System Check:** complete system, sized to Size Gas."
    lines = ["**{0}. The One-Line was not drawn.**".format(res["headline"])]
    lines.append(cr.markdown_table(
        ["Check", "Element", "Reason"],
        [[_code_text(e["code"]), str(e["id"]), e["text"]] for e in res["errors"]]))
    return "\n\n".join(lines)


_CODE_TEXT = {
    "no_piping": "No piping", "disconnected": "Disconnected",
    "open_pipe": "Open pipe end", "open_fitting": "Open fitting",
    "end_not_fixture": "Not a gas fixture", "no_fixtures": "No fixtures",
    "zero_load": "0 MBH fixture", "no_name": "No FIXTURE_NAME",
    "nested_prv": "Nested PRV", "prv_no_table": "PRV table",
    "sizing_error": "Size Gas error", "unsized": "Not sized",
    "size_unreadable": "Size not in table", "size_mismatch": "Size differs",
}


def _code_text(code):
    return _CODE_TEXT.get(code, code)


# ---------------------------------------------------------------------------
# symbols
# ---------------------------------------------------------------------------

def symbols_table(rows):
    """rows: [{'role', 'family', 'status' found|loaded|fallback, 'reason'}]"""
    body = [[r["role"], r.get("family") or "",
             SYMBOL_STATUS.get(r["status"], r["status"]), r.get("reason") or ""]
            for r in rows]
    out = cr.section("Symbols")
    out += cr.table(["Symbol", "Family", "Status", "Note"], body,
                    status_col=2, widths=W_SYMBOLS)
    if any(r["status"] == "loaded" for r in rows):
        out += cr.note("Loaded from REVIT FAM\\GENERIC FAMILIES into this model.")
    return out


def symbols_table_md(rows):
    return cr.markdown_table(
        ["Symbol", "Family", "Status", "Note"],
        [[r["role"], r.get("family") or "",
          SYMBOL_STATUS.get(r["status"], r["status"]), r.get("reason") or ""]
         for r in rows])


# ---------------------------------------------------------------------------
# summary and length check
# ---------------------------------------------------------------------------

def _checks(sch, snap):
    trunk = S.trunk_sum_ft(sch)
    run = float(snap.get("longest_run_ft") or 0.0)
    counts = S.eid_counts(sch)
    edges = set(int(k) for k in (snap.get("edges") or {}))
    once = (set(counts.keys()) == edges and
            all(v == 1 for v in counts.values()))
    mbh_bad = S.mbh_violations(sch)
    return [
        ("Trunk labels sum to the longest run",
         "{0} of {1} ft".format(_ft(trunk), _ft(run)),
         "OK" if abs(trunk - run) <= 0.05 else "FAIL"),
        ("Every pipe drawn exactly once",
         "{0} of {1} pipes".format(len(counts), len(edges)),
         "OK" if once else "FAIL"),
        ("MBH never rises away from the meter",
         "{0} segment(s) rise".format(len(mbh_bad)),
         "OK" if not mbh_bad else "FAIL"),
    ]


def oneline_summary(sch, snap, view_name=None, sheet=None):
    terms = S.terminals_in_order(sch)
    eq = [t for t in terms if t["kind"] == "equipment"]
    total = sum(float(t["mbh"] or 0.0) for t in terms)
    rows = [("Drafting view", view_name or "not drawn"),
            ("Equipment", "{0} ({1} PRV symbols)".format(len(eq), len(eq))),
            ("Caps", str(len([t for t in terms if t["kind"] == "cap"]))),
            ("Total connected load", gas_labels.mbh_text(total)),
            ("Total developed length", "{0}'".format(
                gas_labels.feet(snap.get("longest_run_ft") or 0.0))),
            ("Segments drawn", str(len(sch["segments"])))]
    if sheet:
        rows.insert(1, ("Sheet", sheet))
    out = cr.section("One-Line Summary") + cr.keyvalue(rows)
    out += cr.table(["Check", "Result", "Status"],
                    [list(c) for c in _checks(sch, snap)],
                    status_col=2, widths=[50, 30, 20], width="70%")
    return out


def oneline_summary_md(sch, snap, view_name=None, sheet=None):
    lines = ["**One-Line:** {0}".format(view_name or "not drawn")]
    lines.append(cr.markdown_table(["Check", "Result", "Status"],
                                   [list(c) for c in _checks(sch, snap)]))
    return "\n\n".join(lines)


def length_check(sch, snap):
    """The trunk Segments with their unrounded parts and the total against
    the longest run (the rounded labels can differ by a foot)."""
    rows = []
    for s in S.trunk_segments(sch):
        rows.append([s["id"], s["label"][0], _ft(s["pipe_ft"]),
                     _ft(s["fitting_ft"]), str(s["elbows"]), _ft(s["length_ft"])])
    total = S.trunk_sum_ft(sch)
    out = cr.section("Trunk Length", "centerline, labels round each segment")
    out += cr.table(["Segment", "Label", "Pipe ft", "Fittings ft", "Elbows",
                     "Length ft"], rows, align=["l", "l", "r", "r", "r", "r"],
                    total_row=["Total", "", "", "", "",
                               _ft(total)], widths=W_LENGTH)
    out += cr.note("Longest run (Size Gas): {0} ft. Elbows add {1:.0f} ft each.".format(
        _ft(snap.get("longest_run_ft")), float(snap.get("elbow_equiv_ft") or 5.0)))
    return out


# ---------------------------------------------------------------------------
# warnings and the layout dump
# ---------------------------------------------------------------------------

def oneline_warnings(lines):
    if not lines:
        return ""
    out = cr.section("Warnings")
    for ln in lines:
        out += cr.callout(ln, "warn")
    return out


def oneline_warnings_md(lines):
    return "\n".join("- " + ln for ln in lines)


def layout_diagnostic_text(sch, lay):
    """Plain text dump for copy and paste into a conversation."""
    out = ["ONE-LINE LAYOUT DIAGNOSTIC"]
    segs = sch["segments"]

    def _v(br, depth):
        end = br["end"]
        tag = end["type"] == "terminal" and (end.get("name") or "?") or \
            "J{0}".format(end["node_id"])
        for sid in br["segments"]:
            s = segs[sid]
            out.append("{0}{1} {2:<5} {3} | {4} | pipe {5} fit {6} elb {7} | "
                       "eids {8}{9}".format(
                           "  " * depth, br["kind"][0].upper(), sid,
                           s["label"][0], s["label"][1], _ft(s["pipe_ft"]),
                           _ft(s["fitting_ft"]), s["elbows"],
                           ",".join(str(e) for e in s["eids"]),
                           " T" if s["on_longest"] else ""))
        out.append("{0}  -> {1}".format("  " * depth, tag))
    S.walk(sch["root"], _v)
    out.append("trunk sum {0} ft, longest run {1} ft".format(
        _ft(S.trunk_sum_ft(sch)), _ft(sch.get("longest_run_ft"))))
    if lay is not None:
        out.append("layout bounds {0}".format(
            ", ".join(_ft(v) for v in lay["bounds"])))
        for ln in lay.get("log") or []:
            out.append("layout: " + ln)
    for ln in sch.get("warnings") or []:
        out.append("warning: " + ln)
    return "\n".join(out)


def layout_diagnostic(sch, lay):
    return (cr.section("Layout Diagnostic", "copy and paste for debugging") +
            '<pre style="font-family:Consolas,monospace;font-size:11px;'
            'background:#F4F8FC;border:1px solid {0};padding:6px;">{1}</pre>'
            .format(cr.BORDER, cr.esc(layout_diagnostic_text(sch, lay))))
