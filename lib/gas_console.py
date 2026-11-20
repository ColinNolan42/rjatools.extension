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
    """info: {'meter', 'table', 'table_id', 'heat_content' (natural gas, as
    entered), 'propane_heat_content', 'downstream_table' (or None),
    'origin_kind' ('meter' or 'tie_in')}"""
    rows = _input_rows(info)
    return cr.section("Gas One-Line", "inputs") + cr.keyvalue(rows)


def _input_rows(info):
    rows = [("Start", info.get("meter") or ""),
            ("IFGC table", "{0}  [{1}]".format(info.get("table") or "",
                                               info.get("table_id") or "")),
            ("Heat content of gas", heat_content_text(
                info.get("table_id"), info.get("heat_content"),
                info.get("propane_heat_content")))]
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
    "zero_load": "0 MBH fixture, data error",
    "bad_load": "Load not readable, data error",
    "no_name": "No FIXTURE_NAME",
    "nested_prv": "Nested PRV",
    "sizing_error": "Size Gas error", "unsized": "Not sized",
    "size_unreadable": "Size not in table", "size_mismatch": "Size differs",
    # One-Line gate codes added 2026-11-20 (review fixes)
    "edge_to_missing": "Pipe leads to an element not read, fix the model",
    "origin_legs": "More than one main leaves the meter, fix the model",
    "zero_load_pipe": "Pipe to 0 MBH fixtures only, data error",
    # warnings (printed as callouts, named here for completeness)
    "existing_small": "Existing pipe smaller than Size Gas requires",
    "existing_unreadable": "Existing pipe size not readable",
    "model_read": "Model read", "existing_pipe": "Existing pipe entered",
    "non_gas": "Not a gas system",
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

def _label_feet(seg):
    """The whole feet a trunk label prints: label_ft when the schematic
    balanced it, else the plain rounding the label would use."""
    n = seg.get("label_ft")
    if n is None:
        n = gas_labels.feet(seg.get("length_ft") or 0.0) or 0
    return int(n)


def printed_trunk_feet(sch):
    """Sum of the PRINTED (rounded) trunk labels, the existing run included
    on a tie in."""
    total = sum(_label_feet(s) for s in S.trunk_segments(sch))
    if sch.get("existing"):
        total += _label_feet(sch["existing"])
    return total


def printed_total_feet(snap):
    """The notes block TOTAL DEVELOPED LENGTH as printed (gas_labels.feet of
    the longest run, 0 when there is none)."""
    run = float(snap.get("longest_run_ft") or 0.0)
    return gas_labels.feet(run) if run > 0 else 0


def _checks(sch, snap):
    # n=8: compare what is PRINTED. The unrounded sum always matched, so the
    # check said OK when the rounded labels did not add up to the notes.
    trunk = printed_trunk_feet(sch)
    run = printed_total_feet(snap)
    counts = S.eid_counts(sch)
    edges = set(int(k) for k in (snap.get("edges") or {}))
    once = (set(counts.keys()) == edges and
            all(v == 1 for v in counts.values()))
    mbh_bad = S.mbh_violations(sch)
    return [
        ("Trunk labels sum to the TOTAL DEVELOPED LENGTH",
         "{0}' of {1}'".format(trunk, run),
         "OK" if trunk == run else "FAIL"),
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
    ex = sch.get("existing")
    if ex:
        # n=37: the entered existing run is part of the Total, so it is
        # listed as its own row and the rows add up.
        rows.append(["Existing (entered)", (ex.get("label") or [""])[0],
                     _ft(ex.get("length_ft")), "", "",
                     _ft(ex.get("length_ft"))])
    for s in S.trunk_segments(sch):
        rows.append([s["id"], s["label"][0], _ft(s["pipe_ft"]),
                     str(s["elbows"]), str(s["tees"]), _ft(s["length_ft"])])
    total = S.trunk_sum_ft(sch)
    out = cr.section("Trunk Length", "pipe + elbows + tees, labels round each segment")
    out += cr.table(["Segment", "Label", "Pipe ft", "Elbows", "Tees",
                     "Length ft"], rows, align=["l", "l", "r", "r", "r", "r"],
                    total_row=["Total", "", "", "", "",
                               _ft(total)], widths=W_LENGTH)
    equiv = "Elbows add {0:.0f} ft each, tees {1:.0f} ft each.".format(
        float(snap.get("elbow_equiv_ft") or shared_params.ELBOW_EQUIV_LENGTH_FT),
        float(snap.get("tee_equiv_ft") or shared_params.TEE_EQUIV_LENGTH_FT))
    zones = zone_run_rows(snap)
    if not zones:
        out += cr.note("Longest run (Size Gas): {0} ft. {1}".format(
            _ft(snap.get("longest_run_ft")), equiv))
        return out
    # zoned at a mid stream PRV (Colin, 2026-11-20): Size Gas read each
    # table at its own system's run; the Total above is the whole system
    out += cr.table(["System", "Table", "Longest run ft", "Table row ft"],
                    zones, align=["l", "l", "r", "r"])
    out += cr.note("Longest run, whole system (notes TOTAL DEVELOPED "
                   "LENGTH): {0} ft. Size Gas sized each system at its own "
                   "run above. {1}".format(_ft(snap.get("longest_run_ft")),
                                           equiv))
    return out


def zone_run_rows(snap):
    """[[system, table, run ft, row ft]] from Size Gas's zones when the
    system is split at a mid stream PRV, else []."""
    zones = (snap.get("sizing") or {}).get("zones") or []
    if len(zones) < 2:
        return []
    return [[z.get("label") or "", str(z.get("table_id") or ""),
             _ft(z.get("run_ft")),
             _ft(z.get("row_ft")) if z.get("row_ft") is not None else ""]
            for z in zones]


W_BREAKDOWN = [52, 10, 18, 20]


def developed_breakdown(bd):
    """Developed Length Breakdown (Outputs tick box): the longest run's pipe
    segments, total pipe, elbows, tees and developed length. bd comes from
    oneline_schematic.developed_breakdown."""
    headers, rows = S.breakdown_table(bd)
    total = rows[-1]
    body = rows[:-1]
    farthest = bd.get("farthest") or "FARTHEST OUTLET"
    out = cr.section("Developed Length Breakdown",
                     "longest run to {0}".format(farthest))
    out += cr.table(headers, body, align=["l", "r", "r", "r"],
                    total_row=total, widths=W_BREAKDOWN, width="80%")
    out += cr.note("Developed length = total pipe + {0:.0f} ft per elbow + "
                   "{1:.0f} ft per tee. The notes block TOTAL DEVELOPED "
                   "LENGTH reads {2}'.".format(
                       float(shared_params.ELBOW_EQUIV_LENGTH_FT),
                       float(shared_params.TEE_EQUIV_LENGTH_FT),
                       gas_labels.feet(bd.get("run_ft") or 0.0)))
    if bd.get("mismatch"):
        out += cr.callout("The rows add up to a different run than Size Gas "
                          "reports ({0} ft). Check the longest run.".format(
                              _ft(bd.get("run_ft"))), "warn")
    return out


def developed_breakdown_md(bd):
    headers, rows = S.breakdown_table(bd)
    return "**Developed Length Breakdown**\n\n" + cr.markdown_table(
        headers, rows)


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
            ("Heat content of gas", heat_content_text(
                info.get("table_id"), info.get("heat_content"),
                info.get("propane_heat_content")))]
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


def heat_content_used(result, table_id):
    """The heat content a table's demand was converted with, BTU/CF, or
    None for a table in MBH (no conversion). The propane value on a propane
    CFH table (402.4(37)), else the natural gas value
    (sizing_engine.heat_content_for_table, the one rule)."""
    import sizing_engine
    try:
        return sizing_engine.heat_content_for_table(
            table_id, result.get("heat_content_btu_per_cf"),
            result.get("propane_heat_content_btu_per_cf"))
    except ValueError:
        return result.get("heat_content_btu_per_cf")


def heat_content_text(table_id, heat_content, propane_heat_content=None):
    """The heat content a table actually used, for an inputs or summary row:
    'not used (MBH table)' on 402.4(25) to (36), '2500 BTU/CF (propane)' on
    402.4(37), else the natural gas value, '840 BTU/CF'. The rule is
    heat_content_used (sizing_engine.heat_content_for_table)."""
    if _is_mbh_table(table_id):
        return "not used (MBH table)"
    heat = heat_content_used(
        {"heat_content_btu_per_cf": heat_content,
         "propane_heat_content_btu_per_cf": propane_heat_content}, table_id)
    import gas_tables
    try:
        propane = gas_tables.is_propane_table(table_id)
    except (AttributeError, ValueError):
        propane = False
    return "{0} BTU/CF{1}".format(_f0(heat), " (propane)" if propane else "")


def _is_mbh_table(table_id):
    import gas_tables
    try:
        return gas_tables.table_units(table_id) == "MBH"
    except ValueError:
        return False


def _heat_text(heat, table_id):
    if heat is None:
        return ("not used. Table {0} is in MBH, so the MBH load is compared "
                "directly.".format(table_id))
    import gas_tables
    try:
        propane = gas_tables.is_propane_table(table_id)
    except (AttributeError, ValueError):
        propane = False
    return "{0} BTU/CF{1}".format(_f0(heat), " (propane)" if propane else "")


def _zone_heat(result, z):
    """The heat content the zone's system used: the system entry's own
    heat_content_btu_per_cf when the engine gave one, else the rule."""
    if "heat_content_btu_per_cf" in z:
        return z["heat_content_btu_per_cf"]
    return heat_content_used(result, z["table_id"])


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
            heat = _zone_heat(result, z)
            rows.append(("  " + z["label"],
                         "Table {0}, longest run {1} ft (row {2} ft), "
                         "farthest {3}, heat content {4}".format(
                             z["table_id"], _f1(z["run_ft"]), z["row_ft"],
                             z.get("farthest_name") or "-",
                             "not used (MBH table)" if heat is None
                             else "{0} BTU/CF".format(_f0(heat)))))
    rows.append(("Capacity in", "{0}  (Table {1} as published)".format(
        units, result["table_id"])))
    # The heat content each system actually used (n=24): the propane value
    # on 402.4(37), none on an MBH table.
    heat = (None if units == "MBH" else
            heat_content_used(result, result["table_id"]))
    rows.append(("Heat content", _heat_text(heat, result["table_id"])))
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
                         n_fixtures=0, data_errors=(), tapped=()):
    """The Pre-Sizing Warnings as groups, one callout each. Plain data in:
    no_load [(name, id)], unnamed [(id, mbh)], non_gas [(id, class)],
    bare_end [(id, family)], disconnected [id], data_errors (the walk's
    graph.data_errors dicts: element_id, name, kind, message), tapped [pipe
    id]. A fixture in data_errors is not listed again under no_load.
    Returns [{'title', 'text', 'headers', 'rows'}], empty when nothing
    warns."""
    groups = []
    for w in existing:
        groups.append({"title": "Existing gas pipe", "text": w})
    for w in prv:
        groups.append({"title": "PRV", "text": w})
    for w in table_limit:
        groups.append({"title": "IFGC table limit", "text": w})
    if data_errors:
        # n=10/11: a 0 MBH or unreadable load is a data error
        groups.append({
            "title": "Data errors",
            "text": "{0} gas fixture(s) have a load of 0 MBH or one that "
                    "cannot be read. Each carries no load and does not set "
                    "the longest run. Enter the load and run Size Gas "
                    "again.".format(len(data_errors)),
            "headers": ["Element ID", "Fixture", "Problem"],
            "rows": [[str(d.get("element_id")), d.get("name") or "UNNAMED",
                      str(d.get("message") or d.get("kind") or "")]
                     for d in data_errors]})
        flagged = set(d.get("element_id") for d in data_errors)
        no_load = [(n, i) for n, i in no_load if i not in flagged]
    if tapped:
        groups.append({
            "title": "Tapped pipes",
            "text": "{0} pipe(s) carry a tap (more than one connection "
                    "downstream). Each branch is charged the whole pipe "
                    "length and the pipe carries every branch's load "
                    "(conservative). Break the pipe at the tap with a "
                    "tee.".format(len(tapped)),
            "headers": ["Element ID"],
            "rows": [[str(i)] for i in tapped]})
    if no_load:
        groups.append({
            "title": "0 MBH fixtures",
            "text": "{0}/{1} fixture(s) have 0 MBH load. This is a data "
                    "error: the fixture does not set the developed length "
                    "and a pipe that leads only to 0 MBH fixtures is left "
                    "unsized (status DATA ERROR). Enter the load and run "
                    "Size Gas again.".format(len(no_load), n_fixtures),
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
            widths = {1: [100], 2: [30, 70]}.get(ncol, [20, 30, 50])
            out += cr.table(g["headers"], g["rows"],
                            width="50%" if ncol < 3 else "80%",
                            widths=widths)
    return out


def existing_appliances(rows):
    """rows [(element id, name)]: existing phase gas appliances a tie in
    run reached on new piping. Their load is counted, never resized."""
    if not rows:
        return ""
    out = cr.section("Existing Appliances on New Piping",
                     "{0} found".format(len(rows)))
    out += cr.note("Existing phase gas appliances fed by the new piping. "
                   "Their load is counted and they are never resized.")
    out += cr.table(["Element ID", "Name"],
                    [[str(i), n or "-"] for i, n in rows],
                    width="50%", widths=[30, 70])
    return out


def existing_appliances_md(rows):
    if not rows:
        return ""
    return "**Existing appliances on new piping (load counted):** " + ", ".join(
        "{0} ({1})".format(i, n or "-") for i, n in rows)


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
    if "prv" in fit:
        rows.append(("Meter / PRVs left as drawn", fit.get("prv", 0)))
    if fit.get("transitions_removed"):
        rows.append(("Transitions removed by Revit (pipes now the same "
                     "size)", fit["transitions_removed"]))
    if fit.get("replaced"):
        rows.append(("Fittings replaced by Revit, replacement sized",
                     fit["replaced"]))
    return [(k, str(v)) for k, v in rows]


# -- data errors (0 MBH or unreadable fixture loads) -------------------------

def _data_error_rows(data_errors):
    return [[str(d.get("element_id")), d.get("name") or "UNNAMED",
             str(d.get("message") or d.get("kind") or "")]
            for d in data_errors or []]


def data_errors_section(data_errors=(), unsized_zero_load=(), tapped=()):
    """The Size Gas data errors section, for the script in place of its own
    _graph_notices callouts.

    data_errors: graph.data_errors dicts (element_id, name, kind, message).
    unsized_zero_load: result['unsized_zero_load'] dicts (pipe_id,
        fixture_ids), pipes left unsized because they lead only to 0 MBH
        fixtures.
    tapped: graph.tapped_pipes, pipe ids carrying a tap (a warning).
    Returns "" when there is nothing to say."""
    out = ""
    if data_errors:
        out += cr.section("Data Errors", "{0} fixture(s)".format(
            len(data_errors)))
        out += cr.callout(
            "{0} gas fixture(s) have a load of 0 MBH or one that cannot be "
            "read. Each carries no load and does not set the developed "
            "length. Enter the load and run Size Gas again.".format(
                len(data_errors)), "error")
        out += cr.table(["Element ID", "Fixture", "Problem"],
                        _data_error_rows(data_errors),
                        widths=[20, 30, 50], width="80%")
    if unsized_zero_load:
        if not data_errors:
            out += cr.section("Data Errors")
        out += cr.callout(
            "{0} pipe(s) lead only to 0 MBH fixtures and were NOT sized or "
            "written.".format(len(unsized_zero_load)), "error")
        out += cr.table(["Pipe ID", "0 MBH fixture(s)"],
                        [[str(z.get("pipe_id")),
                          ", ".join(str(i) for i in z.get("fixture_ids") or [])]
                         for z in unsized_zero_load],
                        widths=[30, 70], width="60%")
    if tapped:
        out += cr.callout(
            "{0} pipe(s) carry a tap ({1}). Each branch is charged the whole "
            "pipe length and the pipe carries every branch load "
            "(conservative). Split the pipe at the tap with a tee for an "
            "exact developed length.".format(
                len(tapped), ", ".join(str(i) for i in tapped)), "warn")
    return out


def data_errors_section_md(data_errors=(), unsized_zero_load=(), tapped=()):
    lines = []
    for r in _data_error_rows(data_errors):
        lines.append("- DATA ERROR, fixture {0} ({1}): {2}".format(*r))
    for z in unsized_zero_load or []:
        lines.append("- Pipe {0} leads only to 0 MBH fixture(s) {1}, not "
                     "sized.".format(z.get("pipe_id"), ", ".join(
                         str(i) for i in z.get("fixture_ids") or [])))
    if tapped:
        lines.append("- {0} pipe(s) carry a tap: {1}".format(
            len(tapped), ", ".join(str(i) for i in tapped)))
    return "\n".join(lines)


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

def _summary_rows(info, rolled_back=None):
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
    rolled = _rolled_back(info, rolled_back)
    if rolled:
        rows.append(("Transactions rolled back", ", ".join(rolled)))
    rows = [(k, str(v)) for k, v in rows]
    rows.append(("API approach used", _approach_text(info, rolled)))
    rows.append(("IFGC table", info["table_id"]))
    ec = info.get("existing_check")
    if ec is not None:
        rows.append(("Existing pipe (entered)",
                     '{0}", {1} ft, {2} MBH existing load, {3}'.format(
                         ec["size_nominal"], _f1(ec["length_ft"]),
                         _f1(ec["existing_load_mbh"]),
                         "adequate" if ec["adequate"] else "NOT ADEQUATE")))
    # n=24: the heat content the meter table used (none on an MBH table,
    # the propane value on 402.4(37))
    rows.append(("Heat Content of Gas", heat_content_text(
        info["table_id"], info["heat_content"],
        info.get("propane_heat_content"))))
    if info.get("downstream"):
        for z in info.get("zones") or []:
            text = "Table {0}, run {1} ft (row {2} ft)".format(
                z["table_id"], _f1(z["run_ft"]), z["row_ft"])
            if "heat_content_btu_per_cf" in z:
                heat = z["heat_content_btu_per_cf"]
                text += (", heat content not used (MBH table)" if heat is None
                         else ", {0} BTU/CF".format(_f0(heat)))
            rows.append((z["label"], text))
    else:
        rows.append(("Longest run", "{0} ft".format(_f1(info["longest_run"]))))
        rows.append(("Table row used", "{0} ft".format(info["table_row"])))
    return rows


def _approach_text(info, rolled):
    """The pipe write approach for the summary. None is named after the pipe
    step rolled back: Revit undid every write, so no approach stands."""
    if "pipe sizes" in (rolled or []):
        return "None, the pipe sizes were rolled back"
    return info.get("approach") or "None, all failed"


def _rolled_back(info, rolled_back=None):
    """Names of the write transactions that were rolled back, from the
    rolled_back argument or info["rolled_back"] (a list of names), plus
    "fixture stub pipes" when info["stub_rolled_back"] is set."""
    names = rolled_back
    if names is None:
        names = info.get("rolled_back")
    if names is None or names is False:
        names = []
    elif names is True:
        names = ["a write"]
    elif not isinstance(names, (list, tuple)):
        names = [names]
    names = [str(n) for n in names]
    if info.get("stub_rolled_back") and "fixture stub pipes" not in names:
        names.append("fixture stub pipes")
    return names


def _failure_parts(info, rolled_back=None):
    """Plain sentences naming what failed in the write (n=7/49): pipes,
    stubs, fittings and any rolled back transaction. [] when everything
    reported zero failures."""
    parts = []
    pipe_fail = int(info.get("fail_count") or 0)
    if pipe_fail:
        parts.append("{0}/{1} pipes failed".format(
            pipe_fail, int(info.get("success") or 0) + pipe_fail))
    stub_fail = int(info.get("stub_fail") or 0)
    if stub_fail:
        parts.append("{0} fixture stub pipe(s) failed".format(stub_fail))
    fit_fail = len((info.get("fit") or {}).get("failures") or [])
    if fit_fail:
        parts.append("{0} fitting(s) or cap(s) failed".format(fit_fail))
    for name in _rolled_back(info, rolled_back):
        parts.append("the {0} transaction was rolled back".format(name))
    return parts


def _sized_text(info):
    return ("{0} distribution pipes + {1} stub pipes + {2} fittings/caps "
            "sized.".format(info["success"], info["stub_success"],
                            info["fit"].get("resized", 0)))


def summary(info, rolled_back=None):
    """info: success, stub_success, stub_fail, existing_skip,
    existing_overloaded (count), existing_unchecked (count), fit, fail_count,
    approach, table_id, heat_content, existing_check, zones, downstream,
    longest_run, table_row. Optional: rolled_back (list of transaction
    names), stub_rolled_back (bool).

    The green callout shows only when pipes, stubs and fittings all report
    zero failures and nothing was rolled back; otherwise a warning callout
    names what failed (n=7/49)."""
    rows = _summary_rows(info, rolled_back)
    fit = info["fit"]
    total = (int(info["success"]) + int(info["stub_success"]) +
             int(fit.get("resized", 0)))
    out = cr.section("Summary")
    out += cr.table(["Item", "Value"], [list(r) for r in rows],
                    align=["l", "r"],
                    total_row=["Pipes, stubs and fittings sized", str(total)],
                    widths=W_PAIR, width="70%")
    parts = _failure_parts(info, rolled_back)
    if not parts:
        out += cr.callout(_sized_text(info), "ok")
    else:
        out += cr.callout(
            "Not everything was sized: {0}. Check the failures above.".format(
                ", ".join(parts)), "warn")
    return out


def summary_md(info, rolled_back=None):
    rows = _summary_rows(info, rolled_back)
    out = "**Summary**\n\n" + cr.markdown_table(
        ["Item", "Value"], [list(r) for r in rows])
    parts = _failure_parts(info, rolled_back)
    if not parts:
        out += "\n\n" + _sized_text(info)
    else:
        out += ("\n\nNot everything was sized: {0}. Check the failures "
                "above.".format(", ".join(parts)))
    return out
