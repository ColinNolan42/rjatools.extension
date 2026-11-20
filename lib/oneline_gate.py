# -*- coding: ascii -*-
"""
oneline_gate.py

One-Line rebuild step 2b (Colin, 2026-11-18): One-Line runs only after
Size Gas, on a complete system whose drawn sizes equal what Size Gas
requires. check(snap) returns every reason the diagram shall not be drawn;
the button stops with "Run Size Gas first" (only sizing reasons) or "Fix
the piping model, then run Size Gas again" (any model reason) and lists
them, model reasons first.

    res = check(snap)
    res["ok"]        True when nothing rejects
    res["errors"]    [{"code", "id", "text"}]  each one rejects
    res["warnings"]  [{"code", "id", "text"}]  printed, do not reject

Input is an oneline_snapshot dict (with the in memory size_system result
already folded into each edge's "required_size", and "sizing_error" set
when size_system raised). Pure Python, no Revit import. ASCII source.
"""

import gas_labels
import sizing_engine


RUN_SIZE_GAS = "Run Size Gas first"
# G3 2026-11-19: an open stub, a 0 MBH fixture or a missing name is not
# fixed by running Size Gas again; the headline says to fix the model.
FIX_MODEL = "Fix the piping model, then run Size Gas again"
# reasons Size Gas itself clears; every other code is a model problem
SIZING_CODES = ("sizing_error", "unsized", "size_unreadable", "size_mismatch")
# a drawn pipe matches Size Gas when its diameter is within this of the
# required size's diameter (Colin, 2026-11-20). Names are not compared: the
# CSST pairs EHD-23 / EHD-25 and the like share one diameter, so a pipe Size
# Gas sized EHD-25 reads back as EHD-23. Half the smallest step between two
# sizes in any table (1/4 to 3/8 K&L, 0.125 in.) is 0.0625.
DIAMETER_TOL_IN = 0.05


def headline(errors):
    """FIX_MODEL when any reason is a model problem, else RUN_SIZE_GAS."""
    if not errors:
        return None
    if any(e["code"] not in SIZING_CODES for e in errors):
        return FIX_MODEL
    return RUN_SIZE_GAS


def _err(out, code, eid, text):
    out.append({"code": code, "id": eid, "text": text})


def _as_int(value):
    """value as an int when it reads as one (snapshot ids are ints in the
    dicts and strings as JSON keys), else value itself."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return value


def required_inches(size_name):
    """Diameter Size Gas writes for a required size name, or None."""
    if size_name is None:
        return None
    return sizing_engine.NOMINAL_TO_INCHES.get(size_name)


def diameter_matches(edge):
    """True when the drawn diameter equals the required size's diameter."""
    need_in = required_inches(edge.get("required_size"))
    drawn_in = edge.get("diameter_in")
    if need_in is None or drawn_in is None:
        return False
    return abs(float(drawn_in) - float(need_in)) <= DIAMETER_TOL_IN


def is_existing_edge(edge):
    """Existing phase pipe: not created in the project's last phase (the
    snapshot's is_new, from revit_helpers.is_new_construction)."""
    return edge.get("is_new") is False


def origin_legs(snap):
    """Pipes and fittings leaving the origin (the schematic builds one)."""
    origin = snap.get("origin_id")
    legs = [e["id"] for e in (snap.get("edges") or {}).values()
            if e.get("from") == origin and e.get("to") is not None]
    legs += list((snap.get("children") or {}).get(str(origin)) or [])
    return legs


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
        elif str(e["to"]) not in nodes:
            # a tie in pipe into an existing phase fitting or appliance the
            # walk did not keep: the schematic cannot follow it (review n17)
            _err(errors, "edge_to_missing", e["id"],
                 "Pipe {0} connects to element {1}, which is not part of the "
                 "walked gas system (an existing phase or non gas element). "
                 "Fix the connection in the model.".format(e["id"], e["to"]))
    legs = origin_legs(snap)
    if edges and len(legs) > 1:
        # only one leg can be drawn from the meter; never draw part of the
        # system (Colin, 2026-11-20, review n60)
        _err(errors, "origin_legs", origin,
             "The start element {0} has {1} connected outlets ({2}); the "
             "One-Line draws one. Fix the model so one pipe leaves the "
             "meter.".format(origin, len(legs),
                             ", ".join(str(i) for i in legs)))
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
    # The walk's data errors (pipe_graph.data_errors, carried by the
    # snapshot) are the one record of a 0, unreadable or negative
    # GAS_LOAD_MBH; the load is not tested again here (2026-11-20). A FUTURE
    # cap shall carry its planned MBH: at 0 it is a data error like any
    # other 0 MBH fixture (Colin, 2026-11-20).
    seen_load = set()
    for d in snap.get("data_errors") or []:
        did = d.get("element_id")
        if did in seen_load:
            continue
        seen_load.add(did)
        _err(errors, d.get("kind") or "zero_load", did,
             "{0} (element {1}): {2}.".format(
                 (d.get("name") or "").strip() or gas_labels.UNNAMED, did,
                 str(d.get("message") or "GAS_LOAD_MBH cannot be used")
                 .rstrip(".")))
    for n in fixtures:
        if n["id"] in seen_load or _as_int(n["id"]) in seen_load:
            continue
        load = float(n.get("load_mbh") or 0.0)
        name = (n.get("name") or "").strip()
        if not name or name.upper() == gas_labels.UNNAMED:
            _err(errors, "no_name", n["id"],
                 "Gas fixture {0} carries {1} but has no FIXTURE_NAME.".format(
                     n["id"], gas_labels.mbh_text(load)))

    # --- regulators --------------------------------------------------------
    nested = snap.get("nested_prv_ids") or []
    if nested:
        _err(errors, "nested_prv", nested[0],
             "More than one mid stream PRV on a path ({0}). One mid stream "
             "PRV at most.".format(", ".join(str(i) for i in nested)))
    # A mid stream PRV with Mid Stream PRV unticked is not refused: Size
    # Gas sizes that case on the meter table and warns, and so does the
    # One-Line (Colin, 2026-11-20, review n15). The warning text is the
    # button's prv_notes, printed with the other warnings after drawing.

    # --- sizes match Size Gas ----------------------------------------------
    if snap.get("sizing_error"):
        _err(errors, "sizing_error", origin,
             "Size Gas could not size this system: {0}".format(
                 snap["sizing_error"]))
    else:
        zero_pipes = dict((z.get("pipe_id"), z.get("fixture_ids") or [])
                          for z in snap.get("unsized_zero_load") or [])
        for e in sorted(edges.values(), key=lambda x: x["id"]):
            drawn = e.get("size")
            need = e.get("required_size")
            if is_existing_edge(e):
                # existing pipe is drawn at its modeled size, never refused;
                # Size Gas only audits it (Colin, 2026-11-20, n18)
                _existing_size_warning(warnings, e)
                continue
            if need is not None and diameter_matches(e):
                continue
            if need is None and e["id"] in zero_pipes:
                _err(errors, "zero_load_pipe", e["id"],
                     "Pipe {0} leads only to 0 MBH fixture(s) ({1}); Size Gas "
                     "leaves it unsized. Enter each fixture's "
                     "GAS_LOAD_MBH.".format(e["id"], ", ".join(
                         str(i) for i in zero_pipes[e["id"]]) or "?"))
                continue
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
    for text in snap.get("read_warnings") or []:
        warnings.append({"code": "model_read", "id": None, "text": text})
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

    # model problems first: they are what the engineer fixes
    errors = ([e for e in errors if e["code"] not in SIZING_CODES]
              + [e for e in errors if e["code"] in SIZING_CODES])
    return {"ok": not errors, "errors": errors, "warnings": warnings,
            "headline": headline(errors)}


def _existing_size_warning(warnings, e):
    need_in = required_inches(e.get("required_size"))
    drawn_in = e.get("diameter_in")
    if drawn_in is None or not float(drawn_in):
        warnings.append({"code": "existing_unreadable", "id": e["id"],
                         "text": "Existing pipe {0}: size not readable, drawn "
                                 "as modeled.".format(e["id"])})
    elif need_in is not None and             float(drawn_in) < float(need_in) - DIAMETER_TOL_IN:
        warnings.append({"code": "existing_small", "id": e["id"],
                         "text": "Existing pipe {0} is {1}, smaller than the "
                                 "{2} Size Gas requires. Drawn at its modeled "
                                 "size.".format(e["id"],
                                                e.get("size") or "{0:.3f} in.".format(float(drawn_in)),
                                                e.get("required_size"))})


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
