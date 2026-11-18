# -*- coding: ascii -*-
"""
oneline_gate.py

One-Line rebuild step 2b (Colin, 2026-11-18): One-Line runs only after
Size Gas, on a complete system whose drawn sizes equal what Size Gas
requires. check(snap) returns every reason the diagram shall not be drawn;
the button stops with "Run Size Gas first" and lists them.

    res = check(snap)
    res["ok"]        True when nothing rejects
    res["errors"]    [{"code", "id", "text"}]  each one rejects
    res["warnings"]  [{"code", "id", "text"}]  printed, do not reject

Input is an oneline_snapshot dict (with the in memory size_system result
already folded into each edge's "required_size", and "sizing_error" set
when size_system raised). Pure Python, no Revit import. ASCII source.
"""

import gas_labels


RUN_SIZE_GAS = "Run Size Gas first"


def _err(out, code, eid, text):
    out.append({"code": code, "id": eid, "text": text})


def _name(n):
    return (n.get("name") or "").strip() or "UNNAMED"


def check(snap):
    errors = []
    warnings = []
    nodes = snap.get("nodes") or {}
    edges = snap.get("edges") or {}
    origin = snap.get("origin_id")

    # --- complete system -------------------------------------------------
    if not edges:
        _err(errors, "no_piping", origin,
             "No gas piping was found from the picked start.")
    for eid in snap.get("disconnected") or []:
        _err(errors, "disconnected", eid,
             "Element {0} could not be reached from the meter.".format(eid))
    for e in edges.values():
        if e.get("to") is None:
            _err(errors, "open_pipe", e["id"],
                 "Pipe {0} has an open end.".format(e["id"]))
    for n in nodes.values():
        nid = n["id"]
        kind = n.get("kind")
        if nid == origin:
            continue
        if kind in ("tee", "fitting"):
            for c in n.get("connectors") or []:
                if c.get("to") is None:
                    _err(errors, "open_fitting", nid,
                         "Fitting {0} ({1}) has an open connector.".format(
                             nid, n.get("family") or "fitting"))
                    break
        if kind == "fixture" and not n.get("is_gas_fixture"):
            _err(errors, "end_not_fixture", nid,
                 "Element {0} ({1}) ends the run but is not a gas fixture "
                 "(IS_GAS_FIXTURE).".format(nid, n.get("family") or "?"))

    fixtures = [n for n in nodes.values() if n.get("is_gas_fixture")]
    if not fixtures:
        _err(errors, "no_fixtures", origin, "No gas fixtures were reached.")
    for n in fixtures:
        load = float(n.get("load_mbh") or 0.0)
        name = (n.get("name") or "").strip()
        future = gas_labels.is_future_name(name)
        if load <= 0.0 and not future:
            _err(errors, "zero_load", n["id"],
                 "{0} (element {1}) has GAS_LOAD_MBH 0.".format(
                     _name(n), n["id"]))
        elif load > 0.0 and (not name or name.upper() == gas_labels.UNNAMED):
            _err(errors, "no_name", n["id"],
                 "Gas fixture {0} carries {1} but has no FIXTURE_NAME.".format(
                     n["id"], gas_labels.mbh_text(load)))

    # --- regulators --------------------------------------------------------
    nested = snap.get("nested_prv_ids") or []
    if nested:
        _err(errors, "nested_prv", nested[0],
             "More than one mid stream PRV on a path ({0}). One mid stream "
             "PRV at most.".format(", ".join(str(i) for i in nested)))
    mid = snap.get("midstream_prv_ids") or []
    if mid and not snap.get("downstream_table_id"):
        _err(errors, "prv_no_table", mid[0],
             "A mid stream PRV was found ({0}); pick the downstream "
             "table.".format(mid[0]))

    # --- sizes match Size Gas ----------------------------------------------
    if snap.get("sizing_error"):
        _err(errors, "sizing_error", origin,
             "Size Gas could not size this system: {0}".format(
                 snap["sizing_error"]))
    else:
        for e in sorted(edges.values(), key=lambda x: x["id"]):
            drawn = e.get("size")
            need = e.get("required_size")
            if need is None:
                _err(errors, "unsized", e["id"],
                     "Pipe {0} has no required size from Size Gas.".format(
                         e["id"]))
            elif drawn is None:
                _err(errors, "size_unreadable", e["id"],
                     "Pipe {0} ({1:.3f} in.) is not a size on table {2}; "
                     "required {3}.".format(e["id"],
                                            float(e.get("diameter_in") or 0.0),
                                            e.get("table_id"), need))
            elif drawn != need:
                _err(errors, "size_mismatch", e["id"],
                     "Pipe {0} is {1}, Size Gas requires {2}.".format(
                         e["id"], drawn, need))

    # --- warnings ----------------------------------------------------------
    ex = snap.get("existing_check")
    if ex and not ex.get("adequate"):
        warnings.append({"code": "existing_pipe", "id": origin,
                         "text": "Existing {0} pipe is not adequate for the "
                                 "existing plus new load on table {1} (Size "
                                 "Gas warns the same).".format(
                                     ex.get("size_nominal"), ex.get("table_id"))})
    for item in snap.get("non_gas_skipped") or []:
        warnings.append({"code": "non_gas", "id": item[0],
                         "text": "Walk stopped at element {0} on a non gas "
                                 "system ({1}).".format(item[0], item[1])})

    return {"ok": not errors, "errors": errors, "warnings": warnings,
            "headline": None if not errors else RUN_SIZE_GAS}


def summary_lines(res, limit=20):
    """Short text for the themed alert: headline plus the first reasons."""
    if res["ok"]:
        return []
    lines = [res["headline"] + ". The One-Line was not drawn:"]
    for e in res["errors"][:limit]:
        lines.append("- " + e["text"])
    extra = len(res["errors"]) - limit
    if extra > 0:
        lines.append("- and {0} more (see the console).".format(extra))
    return lines
