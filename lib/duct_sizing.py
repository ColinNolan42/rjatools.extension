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

# Standard round sizes, 2 in increments (firm), and the smallest dimension
# ever used, rectangular or round. Up to 60 in (Colin, 2026-10-20: no reason
# to stop at 36; large mains carry 16,000 cfm and more).
ROUND_SIZES = list(range(6, 62, 2))
MIN_DUCT_DIM = 6
MAX_ASPECT = 4.0
MAX_RECT_DIM = 120       # search limit, inches (large mains, Colin 2026-10-20)

SHAPE_ROUND = "round"
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


def plan_duct(element_id, shape, current, cfm, sys_class, limits,
              friction_fn, is_branch=False, is_flex=False, is_existing=False,
              flow_missing=False):
    """Decide one duct.

    Args:
        shape: SHAPE_ROUND, SHAPE_RECT or None (oval / unknown).
        current: inches, d for round, (w, h) for rect.
        limits: {sys_class: (max_fpm, max_friction)}; a class not listed is
            not sized.

    Returns DuctPlan.
    """
    plan = DuctPlan(element_id)
    plan.shape = shape
    plan.current = current
    plan.cfm = float(cfm or 0.0)
    plan.sys_class = sys_class

    if is_branch:
        plan.skip_reason = "branch (auto-sizing diffuser sets it)"
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
    if sys_class not in limits:
        plan.skip_reason = "no limits for system '{}'".format(sys_class)
        return plan

    max_fpm, max_fric = limits[sys_class]
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
        "  Friction: Darcy-Weisbach, Altshul-Tsal friction factor (same as "
        "Duct Performance).",
        "  Not sized: branches, takeoff boots and diffusers (auto-sizing "
        "diffusers), flex and oval duct, existing phase ductwork.",
    ]
    for name in sorted(limits):
        fpm, fric = limits[name]
        lines.append("  {}: {:.0f} fpm, {:.2f} in. wc per 100 ft.".format(
            name, float(fpm), float(fric)))
    return lines
