# -*- coding: utf-8 -*-
"""
gas_console.py  --  gas tool sections for the pyRevit output window.

Pure renderers, no Revit imports, built from the lib/console_report.py
builders (Plan B, TOOL_ALIGNMENT_PLAN.md). Each section returns one HTML
string; its *_md twin returns the plain markdown the button prints when
print_html fails (console_report.emit). Tested in lib/_test_gas_console.py.

This file holds the One-Line sections (rebuild step 5, 2026-11-18) and, below
them, the Size Gas sections (2026-11-19).

IronPython 2.7: no f-strings; numbers are cast with float() before format.
"""

import console_report as cr
import gas_labels
import shared_params
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
    if info.get("existing"):
        rows.append(("Existing (entered)", info["existing"]))
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
    if any(r["status"] == "fallback" for r in rows):
        out += cr.note("The One-Line never loads families. Load a missing RJA "
                       "symbol family into the project to draw it next run.")
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
            ("Total developed length", "{0}' (pipe + 5 ft per elbow + 3 ft "
                                       "per tee)".format(
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
                     str(s["elbows"]), str(s["tees"]), _ft(s["length_ft"])])
    total = S.trunk_sum_ft(sch)
    out = cr.section("Trunk Length", "pipe + elbows + tees, labels round each segment")
    out += cr.table(["Segment", "Label", "Pipe ft", "Elbows", "Tees",
                     "Length ft"], rows, align=["l", "l", "r", "r", "r", "r"],
                    total_row=["Total", "", "", "", "",
                               _ft(total)], widths=W_LENGTH)
    out += cr.note("Longest run (Size Gas): {0} ft. Elbows add {1:.0f} ft each, tees {2:.0f} ft each.".format(
        _ft(snap.get("longest_run_ft")),
        float(snap.get("elbow_equiv_ft") or shared_params.ELBOW_EQUIV_LENGTH_FT),
        float(snap.get("tee_equiv_ft") or shared_params.TEE_EQUIV_LENGTH_FT)))
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
            out.append("{0}{1} {2:<5} {3} | {4} | pipe {5} elb {6} tee {7} | "
                       "eids {8}{9}".format(
                           "  " * depth, br["kind"][0].upper(), sid,
                           s["label"][0], s["label"][1], _ft(s["pipe_ft"]),
                           s["elbows"], s["tees"],
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


# ===========================================================================
# Size Gas sections (2026-11-19)
# ===========================================================================

W_SIZES = [11, 27, 9, 9, 8, 12, 10, 14]
W_PAIR = [60, 40]
W_FAIL = [14, 36, 50]


def _f1(v):
    return "{0:.1f}".format(float(v or 0.0))


def _f0(v):
    return "{0:.0f}".format(float(v or 0.0))


def _md_lines(rows):
    return "\n".join("**{0}:** {1}".format(k, v) for k, v in rows)


def message(text, kind="info"):
    """One line status or error: a callout. kind: info, ok, warn, error."""
    return cr.callout(text, kind)


def message_md(text, kind="info"):
    return text


# -- inputs, basis, traversal -----------------------------------------------

def sizegas_input_rows(info):
    """info: material, table, table_id, gas, heat_content, downstream_table,
    downstream_table_id, lib, selected (text), check (text), existing (text
    or None)."""
    rows = [("Tools loaded from", info.get("lib") or ""),
            ("Material", info.get("material") or ""),
            ("Gas", info.get("gas") or ""),
            ("IFGC table", "{0}  [{1}]".format(info.get("table") or "",
                                               info.get("table_id") or "")),
            ("Heat content of gas", "{0} BTU/CF".format(
                _f0(info.get("heat_content"))))]
    if info.get("downstream_table"):
        rows.append(("Downstream of PRV", "{0}  [{1}]".format(
            info["downstream_table"], info.get("downstream_table_id") or "")))
    rows.append(("Start", info.get("selected") or ""))
    if info.get("check"):
        rows.append(("Start check", info["check"]))
    if info.get("existing"):
        rows.append(("Existing (entered)", info["existing"]))
    return rows


def sizegas_inputs(info):
    return (cr.section("Size Gas", "IFGC Longest Run Method, inputs") +
            cr.keyvalue(sizegas_input_rows(info)))


def sizegas_inputs_md(info):
    return "**Size Gas, IFGC Longest Run Method**\n\n" + _md_lines(
        sizegas_input_rows(info))


def _basis_rows(result, total_mbh, n_fixtures, units):
    rows = [("Table", result["table_id"]),
            ("Table row", "{0} ft  (longest run {1} ft rounded up)".format(
                result["table_length_used_ft"], _f1(result["longest_run_ft"]))),
            ("Pipe material", result["pipe_material"]),
            ("Inlet pressure", "{0} psi".format(result["inlet_pressure_psi"]))]
    if result.get("downstream_table_id"):
        rows.append(("Downstream", "{0}  (after each PRV)".format(
            result["downstream_table_id"])))
        for z in result.get("zones") or []:
            rows.append(("  " + z["label"],
                         "Table {0}, longest run {1} ft (row {2} ft), "
                         "farthest {3}".format(
                             z["table_id"], _f1(z["run_ft"]), z["row_ft"],
                             z.get("farthest_name") or "-")))
    rows.append(("Capacity in", "{0}  (Table {1} as published)".format(
        units, result["table_id"])))
    if units == "MBH":
        rows.append(("Heat content", "not used. The table is in MBH, so the "
                                     "MBH load is compared directly."))
    else:
        rows.append(("Heat content", "{0} BTU/CF".format(
            _f0(result.get("heat_content_btu_per_cf")))))
    rows.append(("Total load", "{0} MBH, {1} fixtures".format(
        _f1(total_mbh), n_fixtures)))
    return rows


def sizing_basis(result, total_mbh, n_fixtures, units):
    return (cr.section("Sizing Basis") +
            cr.keyvalue(_basis_rows(result, total_mbh, n_fixtures, units)))


def sizing_basis_md(result, total_mbh, n_fixtures, units):
    return "**Sizing basis**\n\n" + _md_lines(
        _basis_rows(result, total_mbh, n_fixtures, units))


def traversal(info):
    """info: nodes, segments, existing_connections (list of ids)."""
    out = cr.section("Traversal") + cr.keyvalue(
        [("Nodes", str(info["nodes"])),
         ("Pipe segments", str(info["segments"]))])
    ex = info.get("existing_connections") or []
    if ex:
        out += cr.callout(
            "{0} connection(s) to Existing phase piping were not followed "
            "({1}). The existing piping is assumed adequate and is not "
            "sized. A new pipe that ends at one carries no load in this "
            "run.".format(len(ex), ", ".join(str(i) for i in ex[:10])), "info")
    return out


def traversal_md(info):
    out = "**Traversal:** {0} nodes, {1} pipe segments.".format(
        info["nodes"], info["segments"])
    ex = info.get("existing_connections") or []
    if ex:
        out += ("\n\n{0} connection(s) to Existing phase piping were not "
                "followed ({1}).".format(
                    len(ex), ", ".join(str(i) for i in ex[:10])))
    return out


def _prv_body(rows):
    return [[str(r["id"]), r.get("family") or "", _f1(r.get("downstream_ft")),
             r["treated"]] for r in rows]


def prv_table(rows):
    """rows: [{'id', 'family', 'downstream_ft', 'treated'}]"""
    return (cr.section("Pressure Regulating Valves") +
            cr.table(["Element ID", "Family", "Pipe downstream ft",
                      "Treated as"], _prv_body(rows),
                     align=["l", "l", "r", "l"], widths=[14, 36, 18, 32]))


def prv_table_md(rows):
    return "**Pressure Regulating Valves**\n\n" + cr.markdown_table(
        ["Element ID", "Family", "Pipe downstream ft", "Treated as"],
        _prv_body(rows))


# -- sizing tables ----------------------------------------------------------

def _size_cells(r):
    cfh = "-" if r.get("cfh") is None else _f1(r["cfh"])
    cap = "-" if r.get("capacity") is None else _f0(r["capacity"])
    length = "-" if r.get("length_ft") is None else _f1(r["length_ft"])
    size = r.get("size")
    return [str(r["pipe_id"]), r.get("serves") or "", _f1(r["mbh"]), cfh,
            '{0}"'.format(size) if size else "-", cap, length, r["status"]]


def _zone_groups(rows, result):
    zones = result.get("zones") or []
    known = set(z["label"] for z in zones)
    groups = []
    for i, z in enumerate(zones):
        zr = [r for r in rows if r.get("zone") == z["label"]]
        if i == 0:
            # A row whose zone label is unknown goes with the first zone.
            zr += [r for r in rows if r.get("zone") not in known]
        groups.append((z, zr))
    return groups


def _zone_line(z):
    return ("{0}: Table {1}, longest run (developed length) {2} ft, table "
            "row {3} ft, farthest {4}".format(
                z["label"], z["table_id"], _f1(z["run_ft"]), z["row_ft"],
                z.get("farthest_name") or "-"))


def _size_headers(units):
    return ["Pipe ID", "Serves", "MBH", "CFH" if units != "MBH" else "CFH n/a",
            "Size", "Capacity " + units, "Length ft", "Status"]


def sizing_tables(rows, result):
    """One sizing table per zone, from sizing_engine.sizing_rows."""
    out = cr.section("Pipe Sizes", "{0} segments".format(len(rows)))
    for z, zr in _zone_groups(rows, result):
        units = zr[0].get("units", "CFH") if zr else "CFH"
        out += cr.subsection(_zone_line(z))
        out += cr.table(_size_headers(units), [_size_cells(r) for r in zr],
                        align=["l", "l", "r", "r", "r", "r", "r", "l"],
                        status_col=7, widths=W_SIZES)
        for r in zr:
            if r.get("note"):
                out += cr.note("{0}: {1}".format(r["pipe_id"], r["note"]))
    return out


def sizing_tables_md(rows, result):
    lines = ["**Pipe sizes** ({0} segments)".format(len(rows))]
    for z, zr in _zone_groups(rows, result):
        units = zr[0].get("units", "CFH") if zr else "CFH"
        lines.append(_zone_line(z))
        lines.append(cr.markdown_table(_size_headers(units),
                                       [_size_cells(r) for r in zr]))
        for r in zr:
            if r.get("note"):
                lines.append("{0}: {1}".format(r["pipe_id"], r["note"]))
    return "\n\n".join(lines)


# -- existing conditions ----------------------------------------------------

def _existing_rows(check):
    units = check.get("units", "CFH")
    load = "{0} MBH existing + {1} MBH new = {2} MBH".format(
        _f1(check["existing_load_mbh"]), _f1(check["new_load_mbh"]),
        _f1(check["total_load_mbh"]))
    if units != "MBH":
        load += " ({0} {1})".format(_f1(check["demand_cfh"]), units)
    if check["capacity_cfh"] is None:
        cap = ("size {0}\" is not in Table {1} at {2} ft. NOT CHECKED."
               .format(check["size_nominal"], check["table_id"],
                       _f1(check["run_ft"])))
    else:
        cap = "{0} {1} at {2} ft developed length, Table {3}".format(
            _f0(check["capacity_cfh"]), units, _f1(check["run_ft"]),
            check["table_id"])
    return [("Existing size", '{0}"'.format(check["size_nominal"])),
            ("Existing length", "{0} ft, added to the new longest run".format(
                _f1(check["length_ft"]))),
            ("Load carried", load), ("Capacity", cap)]


_EXISTING_BASIS = ("Basis: the new piping is assumed to be the longest run. "
                   "The engineer shall verify that existing appliances "
                   "farther from the meter still receive adequate gas.")


def existing_check(check):
    out = cr.section("Existing Pipe", "meter to new connection, entered, "
                                     "not modeled")
    out += cr.keyvalue(_existing_rows(check))
    if check["adequate"]:
        out += cr.callout("ADEQUATE", "ok")
    else:
        out += cr.callout("NOT ADEQUATE. The existing pipe was not resized.",
                          "error")
    return out + cr.note(_EXISTING_BASIS)


def existing_check_md(check):
    return ("**Existing pipe, meter to new connection**\n\n" +
            _md_lines(_existing_rows(check)) + "\n\n**Result:** " +
            ("ADEQUATE" if check["adequate"] else
             "NOT ADEQUATE, the existing pipe was not resized") +
            "\n\n" + _EXISTING_BASIS)


# -- pre-sizing warnings ----------------------------------------------------

def build_warning_groups(existing=(), prv=(), table_limit=(), no_load=(),
                         unnamed=(), non_gas=(), bare_end=(), disconnected=(),
                         n_fixtures=0):
    """The Pre-Sizing Warnings as groups, one callout each. Plain data in:
    no_load [(name, id)], unnamed [(id, mbh)], non_gas [(id, class)],
    bare_end [(id, family)], disconnected [id]. Returns [{'title', 'text',
    'headers', 'rows'}], empty when nothing warns."""
    groups = []
    for w in existing:
        groups.append({"title": "Existing gas pipe", "text": w})
    for w in prv:
        groups.append({"title": "PRV", "text": w})
    for w in table_limit:
        groups.append({"title": "IFGC table limit", "text": w})
    if no_load:
        groups.append({
            "title": "0 MBH fixtures",
            "text": "{0}/{1} fixture(s) have 0 MBH load. Connected pipes "
                    "will be sized at minimum (3/4\").".format(
                        len(no_load), n_fixtures),
            "headers": ["Element ID", "Fixture"],
            "rows": [[str(i), n or "UNNAMED"] for n, i in no_load]})
    if unnamed:
        groups.append({
            "title": "Unnamed fixtures",
            "text": "{0}/{1} fixture(s) have no name (will appear as UNNAMED "
                    "on the one-line diagram).".format(len(unnamed), n_fixtures),
            "headers": ["Element ID", "MBH"],
            "rows": [[str(i), _f1(m)] for i, m in unnamed]})
    if non_gas:
        groups.append({
            "title": "Not a gas system",
            "text": "{0} pipe(s) on a system that is not gas ended the walk. "
                    "Nothing past them was sized. Check the System "
                    "Type.".format(len(non_gas)),
            "headers": ["Element ID", "System"],
            "rows": [[str(i), str(c)] for i, c in non_gas]})
    if bare_end:
        groups.append({
            "title": "Ends that are not gas fixtures",
            "text": "{0} end(s) are not gas fixtures (IS_GAS_FIXTURE not "
                    "set). Their pipes carry no load and were sized at "
                    "minimum (3/4\").".format(len(bare_end)),
            "headers": ["Element ID", "Family"],
            "rows": [[str(i), f or "unknown family"] for i, f in bare_end]})
    if disconnected:
        groups.append({
            "title": "Disconnected elements",
            "text": "{0} disconnected element(s), not included in "
                    "sizing.".format(len(disconnected)),
            "headers": ["Element ID"],
            "rows": [[str(i)] for i in disconnected]})
    return groups


def presize_warnings(groups):
    if not groups:
        return ""
    out = cr.section("Pre-Sizing Warnings", "{0} group(s)".format(len(groups)))
    for g in groups:
        out += cr.callout("{0}: {1}".format(g["title"], g["text"]), "warn")
        if g.get("rows"):
            ncol = len(g["headers"])
            out += cr.table(g["headers"], g["rows"], width="50%",
                            widths=[30, 70] if ncol == 2 else [100])
    return out


def presize_warnings_md(groups):
    lines = []
    for g in groups:
        lines.append("**{0}:** {1}".format(g["title"], g["text"]))
        for row in g.get("rows") or []:
            lines.append("  - " + "  ".join(row))
    return "\n".join(lines)


# -- write results ----------------------------------------------------------

def failure_rows(*groups):
    """[[element, family, reason]] from lists of {'id', 'family', 'reason'}
    (pipes, fittings and stubs, in the order given)."""
    out = []
    for g in groups:
        for r in g:
            out.append([str(r["id"]), r.get("family") or "", r["reason"]])
    return out


def _write_counts(info):
    fit = info["fit"]
    rows = [("Distribution pipes sized", info["pipes_sized"]),
            ("Fixture stub pipes sized", info["stubs_sized"]),
            ("Fittings and caps resized", fit.get("resized", 0)),
            ("Existing fittings left alone", fit.get("existing", 0)),
            ("Fittings with no new pipe sized", fit.get("no_sized_pipe", 0)),
            ("Appliances with no size parameter", fit.get("appliance", 0)),
            ("Transitions left to Revit", fit.get("transitions", 0))]
    if fit.get("transitions_removed"):
        rows.append(("Transitions removed by Revit (pipes now the same "
                     "size)", fit["transitions_removed"]))
    if fit.get("replaced"):
        rows.append(("Fittings replaced by Revit, replacement sized",
                     fit["replaced"]))
    return [(k, str(v)) for k, v in rows]


def write_start(approaches):
    return (cr.section("Writing Sizes") +
            cr.note("Pipe diameter: {0} are tried in order, the first that "
                    "works is used for every pipe.".format(approaches)))


def write_start_md(approaches):
    return ("**Writing sizes to the model.** API approaches in order: " +
            approaches)


def write_results(info):
    """info: pipes_sized, stubs_sized, fit (counts dict), failures (rows from
    failure_rows)."""
    out = cr.section("Write Results")
    out += cr.table(["Item", "Count"], [list(r) for r in _write_counts(info)],
                    align=["l", "r"], widths=W_PAIR, width="60%")
    fit = info["fit"]
    if fit.get("transitions") or fit.get("transitions_removed"):
        out += cr.note(
            "Transitions take their two sizes from their pipes. Revit "
            "updates or removes them when the pipes change, so Size Gas "
            "leaves them alone.")
    fails = info.get("failures") or []
    if fails:
        out += cr.callout("{0} element(s) could not be written.".format(
            len(fails)), "error")
        out += cr.table(["Element", "Family", "Reason"], fails,
                        align=["l", "l", "l"], widths=W_FAIL)
    else:
        out += cr.callout("No write failures.", "ok")
    return out


def write_results_md(info):
    lines = ["**Write results**",
             cr.markdown_table(["Item", "Count"],
                               [list(r) for r in _write_counts(info)])]
    fails = info.get("failures") or []
    if fails:
        lines.append("**{0} element(s) could not be written**".format(
            len(fails)))
        lines.append(cr.markdown_table(["Element", "Family", "Reason"], fails))
    return "\n\n".join(lines)


# -- existing audit ---------------------------------------------------------

def existing_audit(skip_count, overloaded, unchecked):
    if skip_count <= 0:
        return ""
    out = cr.section("Existing Pipes", "{0} skipped".format(skip_count))
    if overloaded:
        out += cr.callout(
            "{0} existing pipe(s) were skipped. {1} of them carry more load "
            "than their current size can handle at the system's longest run "
            "length, or the table publishes no capacity for their size at "
            "that length. These pipes were not resized. Coordinate with the "
            "project team before upsizing existing piping.".format(
                skip_count, len(overloaded)), "warn")
        rows = []
        for e in overloaded:
            cap = e["current_cap"]
            cap_text = ("none published" if cap is None else
                        "{0} {1}".format(_f1(cap), e.get("units", "CFH")))
            rows.append([str(e["pipe_id"]), '{0}"'.format(e["current_nom"]),
                         _f1(e["demand_mbh"]), cap_text,
                         '{0}"'.format(e["recommended_nom"])])
        out += cr.table(["Pipe ID", "Current Size", "New Load MBH",
                         "Current Capacity", "Recommended"], rows,
                        align=["l", "r", "r", "r", "r"])
    else:
        out += cr.callout(
            "{0} existing pipe(s) skipped, all adequately sized for the new "
            "load.".format(skip_count), "ok")
    if unchecked:
        out += cr.callout(
            "{0} existing pipe(s) could not be checked: the drawn diameter "
            "matches no size in the table ({1}).".format(
                len(unchecked), ", ".join(
                    "{0} [Table {1}]".format(u["pipe_id"], u["table_id"])
                    for u in unchecked[:10])), "warn")
    return out


def existing_audit_md(skip_count, overloaded, unchecked):
    if skip_count <= 0:
        return ""
    lines = ["**Existing pipes:** {0} skipped, {1} undersized, {2} not "
             "checked.".format(skip_count, len(overloaded), len(unchecked))]
    for e in overloaded:
        lines.append("- {0}: {1}\" carries {2} MBH, recommended {3}\"".format(
            e["pipe_id"], e["current_nom"], _f1(e["demand_mbh"]),
            e["recommended_nom"]))
    return "\n".join(lines)


# -- summary ----------------------------------------------------------------

def _summary_rows(info):
    fit = info["fit"]
    rows = [("Distribution pipes sized", info["success"]),
            ("Fixture stub pipes sized", info["stub_success"])]
    if info.get("stub_fail"):
        rows.append(("Fixture stub pipe failures", info["stub_fail"]))
    rows.append(("Existing pipes skipped", info["existing_skip"]))
    rows.append(("Existing pipes undersized", info["existing_overloaded"]))
    if info.get("existing_unchecked"):
        rows.append(("Existing pipes not checked", info["existing_unchecked"]))
    rows.append(("Fittings and caps resized", fit.get("resized", 0)))
    if fit.get("replaced"):
        rows.append(("Fittings replaced by Revit, replacement sized",
                     fit["replaced"]))
    rows.append(("Existing fittings left alone", fit.get("existing", 0)))
    rows.append(("Fittings with no new pipe sized", fit.get("no_sized_pipe", 0)))
    rows.append(("Transitions left to Revit", fit.get("transitions", 0)))
    if fit.get("transitions_removed"):
        rows.append(("Transitions removed by Revit",
                     fit["transitions_removed"]))
    rows.append(("Fitting failures", len(fit.get("failures") or [])))
    rows.append(("Pipe failures", info["fail_count"]))
    rows = [(k, str(v)) for k, v in rows]
    rows.append(("API approach used", info.get("approach") or
                 "None, all failed"))
    rows.append(("IFGC table", info["table_id"]))
    ec = info.get("existing_check")
    if ec is not None:
        rows.append(("Existing pipe (entered)",
                     '{0}", {1} ft, {2} MBH existing load, {3}'.format(
                         ec["size_nominal"], _f1(ec["length_ft"]),
                         _f1(ec["existing_load_mbh"]),
                         "adequate" if ec["adequate"] else "NOT ADEQUATE")))
    rows.append(("Heat Content of Gas", "{0} BTU/CF".format(
        _f0(info["heat_content"]))))
    if info.get("downstream"):
        for z in info.get("zones") or []:
            rows.append((z["label"], "Table {0}, run {1} ft (row {2} ft)".format(
                z["table_id"], _f1(z["run_ft"]), z["row_ft"])))
    else:
        rows.append(("Longest run", "{0} ft".format(_f1(info["longest_run"]))))
        rows.append(("Table row used", "{0} ft".format(info["table_row"])))
    return rows


def summary(info):
    """info: success, stub_success, stub_fail, existing_skip,
    existing_overloaded (count), existing_unchecked (count), fit, fail_count,
    approach, table_id, heat_content, existing_check, zones, downstream,
    longest_run, table_row."""
    rows = _summary_rows(info)
    fit = info["fit"]
    total = (int(info["success"]) + int(info["stub_success"]) +
             int(fit.get("resized", 0)))
    out = cr.section("Summary")
    out += cr.table(["Item", "Value"], [list(r) for r in rows],
                    align=["l", "r"],
                    total_row=["Pipes, stubs and fittings sized", str(total)],
                    widths=W_PAIR, width="70%")
    if info["fail_count"] == 0:
        out += cr.callout(
            "{0} distribution pipes + {1} stub pipes + {2} fittings/caps "
            "sized.".format(info["success"], info["stub_success"],
                            fit.get("resized", 0)), "ok")
    else:
        out += cr.callout(
            "{0}/{1} pipes failed. Check the failures above.".format(
                info["fail_count"],
                int(info["success"]) + int(info["fail_count"])), "warn")
    return out


def summary_md(info):
    rows = _summary_rows(info)
    out = "**Summary**\n\n" + cr.markdown_table(
        ["Item", "Value"], [list(r) for r in rows])
    if info["fail_count"] == 0:
        out += ("\n\n{0} distribution pipes + {1} stub pipes + {2} "
                "fittings/caps sized.".format(
                    info["success"], info["stub_success"],
                    info["fit"].get("resized", 0)))
    else:
        out += "\n\n{0} pipe(s) failed. Check the failures above.".format(
            info["fail_count"])
    return out
