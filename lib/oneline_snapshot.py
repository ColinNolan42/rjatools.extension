# -*- coding: ascii -*-
"""
oneline_snapshot.py

One-Line rebuild step 2 (2026-11-18): the ONLY One-Line module that reads
the model. Turns a pipe_graph.NetworkGraph (plus the Size Gas sizing result
the gate reruns in memory) into plain, JSON safe dicts. Everything after
this (gate, schematic, layout) is pure Python and tested in CPython against
snapshots dumped from the test benches.

    snap = snapshot(graph, doc, table_id, sizing_result=None,
                    downstream_table_id=None)
    dump_json(snap, path)  /  load_json(path)

Keys are strings in JSON, so every id map here is keyed by str(element id)
and ids inside lists stay ints.

Length facts come from pipe_graph only: pipe length (edge.length_feet), the
elbow flag and the tee flag of each fitting. Developed length is pipe + 5 ft
per elbow + 3 ft per tee (Colin, 2026-11-19); the length inside a fitting is
not part of it. The pure layer never re-derives geometry.

IronPython 2.7 and CPython 3 (the module itself imports no Revit API; doc is
used only to read each element's created phase name). ASCII source.
"""

import json

import sizing_engine
import shared_params


SNAPSHOT_VERSION = 1


def _r(v, n=4):
    try:
        return round(float(v), n)
    except Exception:
        return None


def _xyz(p):
    if not p:
        return None
    try:
        return [_r(p[0]), _r(p[1]), _r(p[2])]
    except Exception:
        return None


def _log(warns, text):
    """Rule 6: a model read that failed is logged and printed as a warning
    (the gate passes snap["read_warnings"] on to the console)."""
    if warns is not None:
        warns.append(text)
    try:
        import revit_helpers
        revit_helpers._log_entry("WARNING", "oneline_snapshot", None, text)
    except Exception:
        pass        # CPython tests: no Revit, the warning list still has it


READ_WARNINGS_LIMIT = 10


def _eid(element):
    """Element id for a message, through the shared revit_helpers helper."""
    try:
        import revit_helpers
        return revit_helpers._safe_eid(element)
    except Exception:
        return None


def _short(warns, limit=READ_WARNINGS_LIMIT):
    out = []
    for w in warns:
        if w not in out:
            out.append(w)
    if len(out) > limit:
        out = out[:limit] + ["... and {0} more model read warnings (see the "
                             "log).".format(len(out) - limit)]
    return out


def _phase_name(doc, element, warns=None):
    """Created phase name of an element, or None. Never raises; a failed
    read is logged (Rule 6, review n46)."""
    if doc is None or element is None:
        return None
    try:
        pid = element.CreatedPhaseId
        ph = doc.GetElement(pid)
        if ph is None:
            return None
        # Element.Name is unreliable in pyRevit IronPython, read the param.
        try:
            from Autodesk.Revit.DB import Element as _E
            return _E.Name.GetValue(ph)
        except Exception:
            return str(getattr(ph, "Name", None))
    except Exception as ex:
        _log(warns, "Phase of element {0} could not be read ({1}).".format(
            _eid(element), ex))
        return None


def _is_new(doc, element, warns=None):
    """True when created in the project's LAST phase, the rule Size Gas
    uses (revit_helpers.is_new_construction, review n16). None without a
    doc (CPython tests). An unreadable phase is existing (that helper fails
    closed) and is reported."""
    if doc is None or element is None:
        return None
    import revit_helpers
    if revit_helpers.get_created_phase_id(element) is None:
        _log(warns, "Phase of element {0} could not be read; drawn as "
                    "existing, as Size Gas treats it.".format(_eid(element)))
    return bool(revit_helpers.is_new_construction(element, doc))


def read_levels(doc, warns=None):
    """[{"name", "elev"}] for every level in the model, lowest first. The
    One-Line draws elevation vertical by level band (Colin, 2026-11-18)."""
    out = []
    if doc is None:
        return out
    try:
        from Autodesk.Revit.DB import FilteredElementCollector, Level,             Element as _E
        for lv in FilteredElementCollector(doc).OfClass(Level):
            try:
                name = _E.Name.GetValue(lv)
            except Exception:
                name = str(getattr(lv, "Name", ""))
            out.append({"name": name, "elev": _r(lv.Elevation)})
    except Exception as ex:
        _log(warns, "Levels could not be read ({0}); the diagram is drawn "
                    "on one elevation band.".format(ex))
        return []
    out.sort(key=lambda d: d["elev"])
    return out


def _edge_table(edge, table_id, downstream_table_id, sizing_result):
    ctx = (sizing_result or {}).get("edge_context") or {}
    c = ctx.get(edge.element_id)
    if c and c.get("table_id"):
        return c["table_id"]
    if downstream_table_id and getattr(edge, "zone", 0):
        return downstream_table_id
    return table_id


def snapshot(graph, doc, table_id, sizing_result=None,
             downstream_table_id=None):
    """Plain dict snapshot of the graph for the One-Line pure layers."""
    warns = []
    # levels first: their warning is one line and shall not be cut off by
    # a long list of per element phase warnings
    levels = read_levels(doc, warns)
    nodes = {}
    for nid, n in graph.nodes.items():
        connectors = []
        for c in n.connectors or []:
            other = c.get("connected_element_id")
            connectors.append({"to": other, "xyz": _xyz(c.get("origin_xyz"))})
        nodes[str(nid)] = {
            "id": nid,
            "kind": n.node_type,
            "family": n.family_name or "",
            "xyz": _xyz(n.location_xyz),
            "is_elbow": bool(n.is_elbow),
            "is_tee": bool(n.node_type == "tee" and not n.is_gas_fixture
                           and not n.is_prv),
            "is_transition": bool(getattr(n, "is_transition", False)),
            "is_prv": bool(n.is_prv),
            "is_midstream_prv": bool(n.is_midstream_prv),
            "is_gas_fixture": bool(n.is_gas_fixture),
            "load_mbh": _r(n.gas_load_mbh, 3),
            "name": n.fixture_name or "",
            "cum_mbh": _r(n.cumulative_load_mbh, 3),
            "connector_count": int(n.connector_count or len(connectors)),
            "connectors": connectors,
            "zone": int(getattr(n, "zone", 0) or 0),
            "phase": _phase_name(doc, n.element, warns),
            "is_new": _is_new(doc, n.element, warns),
        }

    required = (sizing_result or {}).get("sizes") or {}
    edges = {}
    for eid, e in graph.edges.items():
        tid = _edge_table(e, table_id, downstream_table_id, sizing_result)
        edges[str(eid)] = {
            "id": eid,
            "from": e.from_node_id,
            "to": e.to_node_id,
            "length_ft": _r(e.length_feet),
            "diameter_in": _r(e.diameter_inches),
            "start": _xyz(getattr(e, "start_xyz", None)),
            "end": _xyz(getattr(e, "end_xyz", None)),
            "table_id": tid,
            "size": sizing_engine.nominal_for_diameter(tid, e.diameter_inches),
            "required_size": required.get(eid),
            "cum_mbh": _r(e.cumulative_load_mbh, 3),
            "zone": int(getattr(e, "zone", 0) or 0),
            "phase": _phase_name(doc, e.pipe, warns),
            "is_new": _is_new(doc, e.pipe, warns),
        }

    _place_joints(nodes, edges)

    lr = graph.longest_run or {}
    children = {}
    for k, v in (graph.node_children or {}).items():
        children[str(k)] = list(v)

    origin = graph.nodes.get(graph.origin_id)
    snap = {
        "version": SNAPSHOT_VERSION,
        "origin_id": graph.origin_id,
        "origin_kind": "tie_in" if graph.origin_is_tie_in else "meter",
        "origin_phase": (_phase_name(doc, origin.element, warns)
                         if origin else None),
        "table_id": table_id,
        "downstream_table_id": downstream_table_id,
        "elbow_equiv_ft": shared_params.ELBOW_EQUIV_LENGTH_FT,
        "tee_equiv_ft": shared_params.TEE_EQUIV_LENGTH_FT,
        "longest_run_ft": _r(lr.get("total_length_feet", 0.0)),
        "longest_path": list(lr.get("path_element_ids") or []),
        "farthest_id": lr.get("farthest_fixture_id"),
        "farthest_name": lr.get("farthest_fixture_name") or "",
        "nodes": nodes,
        "edges": edges,
        "children": children,
        "disconnected": list(graph.disconnected or []),
        "prv_ids": list(graph.prv_ids or []),
        "midstream_prv_ids": list(graph.midstream_prv_ids or []),
        "nested_prv_ids": list(graph.nested_prv_ids or []),
        "existing_connections": list(graph.existing_connections or []),
        "non_gas_skipped": [list(x) for x in (graph.non_gas_skipped or [])],
        # the walk's data errors (0 MBH fixtures, FUTURE caps included) and
        # tapped pipes (pipe_graph, 2026-11-20)
        "data_errors": [dict(d) for d in
                        (getattr(graph, "data_errors", None) or [])],
        "tapped_pipes": list(getattr(graph, "tapped_pipes", None) or []),
        "existing": graph.existing,
        "levels": levels,
        "sizing": None,
        "read_warnings": _short(warns),
    }
    if sizing_result:
        snap["sizing"] = {
            "table_id": sizing_result.get("table_id"),
            "longest_run_ft": _r(sizing_result.get("longest_run_ft")),
            "table_length_used_ft": _r(sizing_result.get("table_length_used_ft")),
            "heat_content_btu_per_cf": _r(sizing_result.get("heat_content_btu_per_cf")),
            "downstream_table_id": sizing_result.get("downstream_table_id"),
            # Size Gas's own run per system, meter first (2026-11-20): the
            # notes PRV line and the console state them when zoned
            "zones": [{"label": z.get("label"), "table_id": z.get("table_id"),
                       "run_ft": _r(z.get("run_ft")),
                       "row_ft": _r(z.get("row_ft"))}
                      for z in (sizing_result.get("zones") or [])],
        }
        # pipes Size Gas left unsized because they lead only to 0 MBH
        # fixtures: a data error, not "Run Size Gas first" (2026-11-20)
        snap["unsized_zero_load"] = [
            {"pipe_id": z.get("pipe_id"),
             "fixture_ids": list(z.get("fixture_ids") or [])}
            for z in (sizing_result.get("unsized_zero_load") or [])]
    return snap


def _place_joints(nodes, edges):
    """A joint (two pipes joined end to end, or a tapped pipe) has no Revit
    element and may have no location. It is keyed by the pipe that ends at
    it, so it takes that pipe's far end (the end away from the pipe's from
    node). Pure: reads only the snapshot dicts."""
    for key, n in nodes.items():
        if n.get("kind") != "joint" or n.get("xyz"):
            continue
        e = edges.get(key)
        if e is None:
            continue
        s, t = e.get("start"), e.get("end")
        frm = (nodes.get(str(e.get("from"))) or {}).get("xyz")
        pick = t or s
        if s and t and frm:
            ds = sum((a - b) ** 2 for a, b in zip(s, frm))
            dt = sum((a - b) ** 2 for a, b in zip(t, frm))
            pick = s if ds > dt else t
        n["xyz"] = list(pick) if pick else None


def _json_default(o):
    """IronPython hands element ids back as long / Int64; write them as int."""
    try:
        return int(o)
    except Exception:
        return str(o)


def dump_json(snap, path):
    f = open(path, "w")
    try:
        json.dump(snap, f, indent=1, sort_keys=True, default=_json_default)
    finally:
        f.close()


def load_json(path):
    f = open(path, "r")
    try:
        return json.load(f)
    finally:
        f.close()
