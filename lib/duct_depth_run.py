# -*- coding: ascii -*-
"""Duct space check for the design agents (agent_api.duct_size_in_depth).

For each row: the depth available for the duct (bottom of structure minus
the ceiling, minus the hanger allowance, duct_sizing.available_depth_in).
Both heights are ELEVATIONS in host feet (the same datum): a ceiling's
Height Offset From Level plus its level's elevation, a structure member's
bounding box minimum z from read_scoped. Then the smallest firm size meeting the system's firm velocity and friction
limits that fits that depth (duct_sizing.size_in_depth). A row with no fit
comes back with its reason and the run returns needs_decision, so the
designer asks Colin; nothing is upsized or flattened past the rules here
(Colin, 2026-11-13 and 2026-11-17).

Read only: no model access. doc is accepted because the gateway calls every
tool as agent_api.<tool>(doc, **kwargs), and is not used.

IronPython 2.7 and CPython.
"""

import duct_sizing
import headless
import vav_checks

TOOL = "duct_size_in_depth"
QIDS = ("no_fit",)  # a literal: _test_agent_api reads it from the source
QID_NO_FIT = QIDS[0]
ENGINE = "both"
MAX_ROWS = 500

# A row's system: the System Classification names, the same keys as the
# firm limits (hvac_graph.FIRM_DEFAULTS) and Duct Performance settings.
SYSTEMS = ("Supply Air", "Return Air", "Exhaust Air", "Outside Air",
           "Transfer Air")
SHAPES = (duct_sizing.SHAPE_RECT, duct_sizing.SHAPE_ROUND)
SETTING_KEYS = ("limits", "hanger_in")


def _number(v):
    """A finite float for an int or float, never a bool, text, NaN or
    infinity (json.loads accepts NaN and Infinity); else None."""
    if isinstance(v, bool):
        return None
    try:
        long_type = long  # noqa: F821 (IronPython 2.7)
    except NameError:
        long_type = int
    if not isinstance(v, (int, float, long_type)):
        return None
    f = float(v)
    if f != f or f in (float("inf"), float("-inf")):
        return None
    return f


def _firm():
    """The real friction function and firm limits, from hvac_graph (Revit).
    The calc basis is reset first so no override from an earlier run in the
    same session leaks into this one."""
    import hvac_graph
    hvac_graph.set_calc_basis()
    return hvac_graph.duct_friction_loss_per_100ft, hvac_graph.FIRM_DEFAULTS


def _check_settings(settings):
    errs = []
    limits, hanger = {}, duct_sizing.HANGER_ALLOWANCE_IN
    if settings is None:
        return limits, hanger, errs
    if not isinstance(settings, dict):
        return limits, hanger, ["settings shall be an object"]
    for k in settings:
        if k not in SETTING_KEYS:
            errs.append("unknown setting {!r} (one of {})".format(
                k, ", ".join(SETTING_KEYS)))
    lim_in = settings.get("limits") or {}
    if not isinstance(lim_in, dict):
        errs.append("limits shall be an object of system: [max fpm, max "
                    "in. wc/100 ft]")
        lim_in = {}
    for system, pair in lim_in.items():
        if system not in SYSTEMS:
            errs.append("limits for unknown system {!r}".format(system))
            continue
        ok = isinstance(pair, (list, tuple)) and len(pair) == 2
        fpm = _number(pair[0]) if ok else None
        fric = _number(pair[1]) if ok else None
        if not fpm or not fric or fpm <= 0 or fric <= 0:
            errs.append("limits for {} shall be [max fpm, max in. wc/100 ft],"
                        " both above 0".format(system))
            continue
        limits[system] = (fpm, fric)
    if "hanger_in" in settings:
        h = _number(settings["hanger_in"])
        if h is None or h < 0:
            errs.append("hanger_in shall be a number, 0 or more")
        else:
            hanger = h
    return limits, hanger, errs


def _check_row(i, row):
    """(clean row, errors) for one input row."""
    where = "row {}".format(i + 1)
    if not isinstance(row, dict):
        return None, ["{}: shall be an object".format(where)]
    errs = []
    rid = row.get("id")
    if rid is None or isinstance(rid, bool) or not str(rid).strip():
        errs.append("{}: id is missing".format(where))
    else:
        where = "row {} ({})".format(i + 1, rid)
    clean = {"id": str(rid) if rid is not None else None}
    for key in ("cfm", "structure_bottom_elev_ft", "ceiling_elev_ft"):
        v = _number(row.get(key))
        if v is None:
            errs.append("{}: {} is not a number".format(where, key))
        clean[key] = v
    if clean.get("cfm") is not None and clean["cfm"] <= 0:
        errs.append("{}: cfm shall be above 0".format(where))
    if row.get("shape") not in SHAPES:
        errs.append("{}: shape {!r} is not rect or round".format(
            where, row.get("shape")))
    if row.get("system") not in SYSTEMS:
        errs.append("{}: system {!r} is not one of {}".format(
            where, row.get("system"), ", ".join(SYSTEMS)))
    clean["shape"] = row.get("shape")
    clean["system"] = row.get("system")
    # A rect duct needs the designer's height (the design rules: 12 in mains
    # where the depth allows); without it the duct would take the whole
    # depth (a 12x44 for 1000 cfm in a 4 ft plenum, live check 2026-11-17).
    clean["preferred_h_in"] = None
    if row.get("shape") == duct_sizing.SHAPE_RECT:
        ph = _number(row.get("preferred_h_in"))
        if ph is None or ph < duct_sizing.MIN_DUCT_DIM:
            errs.append("{}: a rect row needs preferred_h_in, the height the "
                        "design rules call for (at least {} in)".format(
                            where, duct_sizing.MIN_DUCT_DIM))
        clean["preferred_h_in"] = ph
    return clean, errs


def _label(size, shape):
    if shape == duct_sizing.SHAPE_ROUND:
        return "{} in round".format(int(size))
    return "{}x{}".format(int(size[0]), int(size[1]))


def size_rows(rows, friction_fn, firm, limits=None, hanger_in=None):
    """Plain logic, no Revit: the sized rows and the ids with no fit."""
    if hanger_in is None:
        hanger_in = duct_sizing.HANGER_ALLOWANCE_IN
    out, no_fit = [], []
    for r in rows:
        sys_class = r["system"]
        max_fpm, max_fric = vav_checks.limits_for(sys_class, limits, firm)
        depth = duct_sizing.available_depth_in(
            r["structure_bottom_elev_ft"], r["ceiling_elev_ft"], hanger_in)
        size, reason = duct_sizing.size_in_depth(
            r["cfm"], depth, max_fpm, max_fric, friction_fn, shape=r["shape"],
            preferred_h_in=r.get("preferred_h_in"))
        rec = {"id": r["id"], "system": r["system"], "shape": r["shape"],
               "cfm": r["cfm"], "depth_in": round(depth, 2),
               "limits": {"max_fpm": max_fpm, "max_friction": max_fric}}
        if size is None:
            rec.update({"fits": False, "reason": reason})
            no_fit.append(r["id"])
        else:
            if r["shape"] == duct_sizing.SHAPE_ROUND:
                area = duct_sizing.round_area_ft2(size)
                dh = float(size)
                rec["size_in"] = size
            else:
                area = duct_sizing.rect_area_ft2(size[0], size[1])
                dh = duct_sizing.rect_dh_in(size[0], size[1])
                rec["size_in"] = [size[0], size[1]]
            fpm, fric = duct_sizing.performance(r["cfm"], area, dh,
                                                friction_fn)
            rec.update({"fits": True, "size": _label(size, r["shape"]),
                        "fpm": int(round(fpm)), "friction": round(fric, 4)})
        out.append(rec)
    return out, no_fit


def run_headless(doc=None, rows=None, settings=None, friction_fn=None,
                 firm=None):
    """The envelope for agent_api. friction_fn and firm are for tests; in
    Revit they come from hvac_graph."""
    limits, hanger, errs = _check_settings(settings)
    if not isinstance(rows, list) or not rows:
        errs.append("rows shall be a non empty list")
        rows = []
    elif len(rows) > MAX_ROWS:
        errs.append("{} rows; at most {} per run".format(len(rows), MAX_ROWS))
    clean = []
    for i, row in enumerate(rows[:MAX_ROWS]):
        c, e = _check_row(i, row)
        errs.extend(e)
        clean.append(c)
    ids = [c["id"] for c in clean if c and c["id"] is not None]
    if len(ids) != len(set(ids)):
        errs.append("row ids shall be unique")
    if errs:
        return headless.result(TOOL, "refused", errors=errs)
    if friction_fn is None or firm is None:
        real_fn, real_firm = _firm()
        friction_fn = friction_fn or real_fn
        firm = firm or real_firm
    out, no_fit = size_rows(clean, friction_fn, firm, limits, hanger)
    data = {"rows": out,
            "basis": {"hanger_in": hanger,
                      "limits": "firm defaults" if not limits
                      else "firm defaults, with {} from settings".format(
                          ", ".join(sorted(limits))),
                      "sizes": "even inches, 6 in minimum, 4:1 maximum"}}
    if no_fit:
        return headless.result(TOOL, "needs_decision", data=data, decision={
            "qid": QID_NO_FIT,
            "text": "No size meets the limits in the space for: {}. Ask "
                    "Colin; never upsize or flatten past the rules.".format(
                        ", ".join(no_fit)),
            "suggested": "stop"})
    return headless.result(TOOL, "ok", data=data)
