# -*- coding: ascii -*-
# duct_sizing.py
# Size Duct: the size each MAIN duct shall be, from its airflow and the firm
# velocity and friction limits.
#
# Rules (Colin, 2026-10-20):
#   - Smallest even size (2 in steps, 6 in minimum) that meets BOTH the max
#     velocity and the max friction for its system, the same limits Duct
#     Performance checks, so a sized system comes back green.
#   - Round: the firm round size list.
#   - Rectangular: hold the DRAWN height (ceiling depth) and solve the width.
#     4:1 maximum aspect ratio either way. When no width fits at that height
#     within 4:1, the height steps up 2 in at a time.
#   - Mains only. Branches (one terminal downstream), takeoff boots and
#     diffusers are never sized: the auto-sizing diffusers set the branch, and
#     Duct Performance reports a custom one that is off.
#   - Flex duct and oval duct are reported, not sized.
#
# Pure logic: the friction function is passed in (hvac_graph's Altshul-Tsal
# Darcy-Weisbach in Revit), so this module tests in plain CPython.
#
# IronPython 2.7

import math

import vav_checks

# Standard round sizes, 2 in increments (firm), and the smallest dimension
# ever used, rectangular or round. Up to 60 in (Colin, 2026-10-20: no reason
# to stop at 36; large mains carry 16,000 cfm and more).
ROUND_SIZES = list(range(6, 62, 2))
MIN_DUCT_DIM = 6
MAX_ASPECT = 4.0
MAX_RECT_DIM = 120       # search limit, inches (large mains, Colin 2026-10-20)

SHAPE_ROUND = "round"
SKIP_INLET_RUNOUT = "VAV inlet runout (sized to the box)"
SHAPE_RECT = "rect"


def snap_even(inches):
    """Round up to the next even inch. RJA never uses odd duct sizes.

    A small epsilon keeps feet to inch round-trip noise (10.000000000000002)
    from being pushed to the next size.
    """
    n = int(math.ceil(inches - 1e-6))
    if n % 2 != 0:
        n += 1
    return max(n, 2)


def rect_area_ft2(w_in, h_in):
    return float(w_in) * float(h_in) / 144.0


def rect_dh_in(w_in, h_in):
    """Hydraulic diameter 4A/P, inches."""
    return 2.0 * float(w_in) * float(h_in) / (float(w_in) + float(h_in))


def round_area_ft2(d_in):
    return math.pi * (float(d_in) / 12.0) ** 2 / 4.0


def performance(cfm, area_ft2, dh_in, friction_fn):
    """(fpm, friction in. wc per 100 ft) for a flow in a duct."""
    fpm = float(cfm) / area_ft2
    return fpm, friction_fn(fpm, dh_in)


def _passes(cfm, area_ft2, dh_in, max_fpm, max_fric, friction_fn):
    fpm, fric = performance(cfm, area_ft2, dh_in, friction_fn)
    return fpm <= max_fpm + 1e-9 and fric <= max_fric + 1e-12


def size_round(cfm, max_fpm, max_fric, friction_fn):
    """Smallest firm round size meeting both limits, or None (over 36 in)."""
    for d in ROUND_SIZES:
        if _passes(cfm, round_area_ft2(d), float(d), max_fpm, max_fric,
                   friction_fn):
            return d
    return None


def size_rect(cfm, drawn_h_in, max_fpm, max_fric, friction_fn):
    """(width, height) at the drawn height, or the next height that fits.

    Width is searched from the narrowest the 4:1 limit allows, in 2 in
    steps, up to 4 x height. Returns None when nothing up to MAX_RECT_DIM
    fits.
    """
    h = max(snap_even(drawn_h_in), MIN_DUCT_DIM)
    while h <= MAX_RECT_DIM:
        w_min = max(snap_even(h / MAX_ASPECT), MIN_DUCT_DIM)
        w_max = min(int(h * MAX_ASPECT), MAX_RECT_DIM)
        w = w_min
        while w <= w_max:
            if _passes(cfm, rect_area_ft2(w, h), rect_dh_in(w, h), max_fpm,
                       max_fric, friction_fn):
                return w, h
            w += 2
        h += 2
    return None


class DuctPlan(object):
    """What Size Duct decided for one duct."""

    def __init__(self, element_id):
        self.element_id = element_id
        self.sys_class = ""
        self.cfm = 0.0
        self.shape = None
        self.current = None      # inches: d, or (w, h)
        self.new = None          # inches: d, or (w, h); None = not sized
        self.changed = False
        self.skip_reason = ""    # why it was not sized
        self.fpm = None          # at the new size
        self.friction = None
        self.length_ft = 0.0
        self.limit_key = ""      # which limits row sized it (system, or MP)


def plan_duct(element_id, shape, current, cfm, sys_class, limits,
              friction_fn, is_branch=False, is_flex=False, is_existing=False,
              flow_missing=False, limit_key=None, is_inlet_runout=False):
    """Decide one duct.

    Args:
        shape: SHAPE_ROUND, SHAPE_RECT or None (oval / unknown).
        current: inches, d for round, (w, h) for rect.
        limits: {sys_class: (max_fpm, max_friction)}; a class not listed is
            not sized.
        limit_key: the limits row to size against when it is not the duct's
            own system: vav_checks.MP_KEY for a VAV medium pressure main.
        is_inlet_runout: a VAV inlet runout (MP, one box downstream). Sized
            to the box by the engineer, so never resized, like a branch.

    Returns DuctPlan.
    """
    plan = DuctPlan(element_id)
    plan.shape = shape
    plan.current = current
    plan.cfm = float(cfm or 0.0)
    plan.sys_class = sys_class
    plan.limit_key = limit_key or sys_class

    if is_branch:
        plan.skip_reason = "branch (auto-sizing diffuser sets it)"
        return plan
    if is_inlet_runout:
        plan.skip_reason = SKIP_INLET_RUNOUT
        return plan
    if is_flex:
        plan.skip_reason = "flex duct"
        return plan
    if is_existing:
        plan.skip_reason = "existing phase, not changed"
        return plan
    if flow_missing:
        # A terminal downstream with no Flow counts as 0 cfm, which would
        # size this main too small. Never write a size from a partial flow.
        plan.skip_reason = "a terminal downstream has no Flow value"
        return plan
    if shape not in (SHAPE_ROUND, SHAPE_RECT):
        plan.skip_reason = "oval or unknown shape"
        return plan
    if plan.cfm <= 0.0:
        plan.skip_reason = "no airflow reaches it"
        return plan
    if plan.limit_key not in limits:
        plan.skip_reason = "no limits for system '{}'".format(plan.limit_key)
        return plan

    max_fpm, max_fric = limits[plan.limit_key]
    if shape == SHAPE_ROUND:
        d = size_round(plan.cfm, max_fpm, max_fric, friction_fn)
        if d is None:
            plan.skip_reason = "needs more than {} in round".format(
                ROUND_SIZES[-1])
            return plan
        plan.new = d
        plan.fpm, plan.friction = performance(plan.cfm, round_area_ft2(d),
                                              float(d), friction_fn)
        plan.changed = (current is None or
                        abs(float(current) - d) > 0.01)
    else:
        drawn_h = current[1] if current else MIN_DUCT_DIM
        wh = size_rect(plan.cfm, drawn_h, max_fpm, max_fric, friction_fn)
        if wh is None:
            plan.skip_reason = "needs more than {} in".format(MAX_RECT_DIM)
            return plan
        plan.new = wh
        plan.fpm, plan.friction = performance(
            plan.cfm, rect_area_ft2(wh[0], wh[1]), rect_dh_in(wh[0], wh[1]),
            friction_fn)
        plan.changed = (current is None or
                        abs(float(current[0]) - wh[0]) > 0.01 or
                        abs(float(current[1]) - wh[1]) > 0.01)
    return plan


# ---------------------------------------------------------------------------
# Step-downs along one duct (Colin, 2026-10-20)
# ---------------------------------------------------------------------------
# A main steps down after any tap where the reduced airflow calls for a
# smaller size, the same test Duct Performance applies, so a sized main never
# reads oversized. The transition goes just past the tap. Live test
# 2026-10-20 (Revit 2026): BreakCurve keeps each tap on its own piece, and a
# "Rectangular Transition - Length" (15 in) is cut out of the UPSTREAM piece,
# so the break sits at tap edge + clearance + transition length.

TRANSITION_LEN_FT = 1.5       # allowance; the live rect transition was 1.25 ft
STEP_CLEARANCE_FT = 0.5       # between a tap fitting and a transition
TAP_FACE_CLEARANCE_IN = 2.0   # main face at least tap size + 2 in (AI rules)


class Tap(object):
    """A takeoff on a main, as seen along that main."""

    def __init__(self, pos_ft, half_ft, cfm, size_in, on_width_face):
        self.pos_ft = float(pos_ft)        # from the upstream end
        self.half_ft = float(half_ft)      # half its footprint along the main
        self.cfm = float(cfm or 0.0)
        self.size_in = float(size_in or 0.0)
        self.on_width_face = bool(on_width_face)   # top or bottom of a rect


class Section(object):
    def __init__(self, start_ft, cfm, size):
        self.start_ft = start_ft
        self.end_ft = None
        self.cfm = cfm
        self.size = size


def _smaller(shape, a, b):
    if shape == SHAPE_ROUND:
        return float(a) < float(b) - 0.01
    return float(a[0]) * float(a[1]) < float(b[0]) * float(b[1]) - 0.01


def _fits_taps(shape, size, taps):
    """Every tap still on the section fits the main face it sits on."""
    for tap in taps:
        if shape == SHAPE_ROUND:
            if size < tap.size_in - 0.01:
                return False
        elif tap.on_width_face and \
                size[0] < tap.size_in + TAP_FACE_CLEARANCE_IN - 0.01:
            return False
    return True


def plan_sections(shape, drawn, duct_cfm, length_ft, taps, max_fpm, max_fric,
                  friction_fn):
    """Sections along one main duct, stepping down after taps.

    Returns (sections, notes). One section means no step-down. A step that
    has no room before the next tap (or the duct end) is skipped and noted;
    the larger size then carries on to the next tap.
    """
    def size_for(cfm):
        if shape == SHAPE_ROUND:
            return size_round(cfm, max_fpm, max_fric, friction_fn)
        return size_rect(cfm, drawn[1] if drawn else MIN_DUCT_DIM, max_fpm,
                         max_fric, friction_fn)

    taps = sorted(taps, key=lambda t: t.pos_ft)
    notes = []
    first = size_for(duct_cfm)
    if first is None:
        return [], ["needs more than the largest size"]
    current = Section(0.0, float(duct_cfm), first)
    sections = [current]
    flow = float(duct_cfm)
    for i, tap in enumerate(taps):
        flow -= tap.cfm
        if flow <= 0.0:
            break
        need = size_for(flow)
        if need is None or not _smaller(shape, need, current.size):
            continue
        if not _fits_taps(shape, need, taps[i + 1:]):
            notes.append("no step after tap {}: a later tap needs the larger "
                         "face".format(i + 1))
            continue
        brk = tap.pos_ft + tap.half_ft + STEP_CLEARANCE_FT + TRANSITION_LEN_FT
        if i + 1 < len(taps):
            limit = taps[i + 1].pos_ft - taps[i + 1].half_ft - STEP_CLEARANCE_FT
        else:
            limit = length_ft - STEP_CLEARANCE_FT
        if brk > limit:
            notes.append("no room to step down after tap {}".format(i + 1))
            continue
        current.end_ft = brk
        current = Section(brk, flow, need)
        sections.append(current)
    current.end_ft = float(length_ft)
    return sections, notes


def size_label(shape, size):
    if size is None:
        return "-"
    if shape == SHAPE_ROUND:
        return '{:.0f}" dia'.format(float(size))
    return '{:.0f}x{:.0f}'.format(float(size[0]), float(size[1]))


def basis_lines(limits):
    lines = [
        "Size Duct basis: mains only (two or more terminals downstream). "
        "Smallest even size, {} in minimum, meeting BOTH the max velocity "
        "and the max friction.".format(MIN_DUCT_DIM),
        "  Round: {} to {} in, 2 in steps. Rectangular: drawn height held, "
        "width solved, {:.0f}:1 maximum aspect; the height steps up 2 in only "
        "when no width fits.".format(ROUND_SIZES[0], ROUND_SIZES[-1],
                                     MAX_ASPECT),
        "  Step-downs: after any tap where the reduced airflow calls for a "
        "smaller size; transition just past the tap ({:.1f} ft clearance), "
        "skipped when there is no room before the next tap.".format(
            STEP_CLEARANCE_FT),
        "  Friction: Darcy-Weisbach, Altshul-Tsal friction factor (same as "
        "Duct Performance).",
        "  Not sized: branches, takeoff boots and diffusers (auto-sizing "
        "diffusers), flex and oval duct, existing phase ductwork.",
    ]
    if vav_checks.MP_KEY in limits:
        lines.append(
            "  VAV: supply mains between the unit and the box inlets (MP) with "
            "two or more boxes downstream are sized to the MP limits (EDR VAV "
            "Guide p3, p100). A box inlet runout (one box downstream) is sized "
            "to the box by the engineer and never touched. After the box, as "
            "above. Boxes are never touched.")
    for name in sorted(limits):
        fpm, fric = limits[name]
        lines.append("  {}: {:.0f} fpm, {:.2f} in. wc per 100 ft.".format(
            name, float(fpm), float(fric)))
    return lines
