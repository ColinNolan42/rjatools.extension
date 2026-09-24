# -*- coding: utf-8 -*-
"""
fitting_tables.py  --  duct fitting loss coefficients ("C" values)

Source: RJA's own standard, `SP_LOSS_WORKSHEET(1).xls`
(FOR CLAUDE\\Design Resources\\MECHANICAL\\3 - Mechanical Calculations), which
cites 1985 ASHRAE fitting numbers. Transcribed 2026-09-24; the sum of all 16
published C values is 19.91, which matches that sheet's own AN30 check cell, so
the transcription is verified against the source.

A fitting's pressure loss is

    dP = C * Pv          Pv = (V / K)^2 in. wc, K ~ 4008 for standard air

with Pv from `hvac_graph.velocity_pressure_inwg()`. C is dimensionless and
depends on the fitting's geometry and, for a junction, WHICH WAY the air being
traced goes through it.

DO NOT size from memory here. Every number below is from the worksheet, and any
new fitting needs a real source before it gets a C. Values Colin has explicitly
ruled on are marked.

IronPython 2.7 / pyRevit -- no f-strings, no walrus, no nonlocal.
"""

import logging

log = logging.getLogger(__name__)


# ── Published C values ───────────────────────────────────────────────────────
#
# ASHRAE 1985 fitting number in the trailing comment, exactly as the worksheet
# lists it.

C_RECT_ELBOW_90          = 0.25   # 3-5/3-7  90 deg rect elbow
C_ROUND_ELBOW_90         = 0.33   # 3-1      90 deg round elbow

C_SUPPLY_TAP_BRANCH      = 0.98   # 6-29  supply dovetail take-off rect/rect, typ branch
C_SUPPLY_TAP_END         = 0.54   # 6-25  supply dovetail take-off, end of main branch
C_SUPPLY_TAP_MAIN        = 0.28   # 6-23  main duct pressure drop @ supply take-off
C_SUPPLY_TAP_CONICAL     = 1.40   # 6-27  supply conical take-off rect/round, all branches

C_EXHAUST_TAP_BRANCH     = 0.76   # 6-9   exhaust dovetail take-off rect/rect, typ branch
C_EXHAUST_TAP_END        = 3.63   # 6-9   exhaust dovetail take-off, end of main branch
C_EXHAUST_TAP_MAIN       = 0.60   # 6-3   main duct pressure drop @ exhaust take-off
C_EXHAUST_TAP_CONICAL     = 1.27  # 6-7   exhaust conical take-off rect/round, typ branch
C_EXHAUST_TAP_CONICAL_END = 5.95  # 6-7   exhaust conical take-off, end of main branch

C_TRANSITION_EXPANSION   = 0.84   # 4-3   transition expansion   (60 deg max)
C_TRANSITION_CONTRACTION = 0.08   # 5-1   transition contraction (60 deg max)

C_WYE_BULLHEAD           = 0.30   # 6-33  bull head 90 deg wye
C_OFFSET                 = 1.70   # 3-13  offset
C_ABRUPT_EXIT            = 1.00   # 2-11  abrupt exit (stacks no cone)

# The worksheet's own checksum. Asserted by the self-test, so a typo in any
# single value above cannot pass silently.
PUBLISHED_C_SUM = 19.91

_ALL_PUBLISHED_C = (
    C_RECT_ELBOW_90, C_ROUND_ELBOW_90,
    C_SUPPLY_TAP_BRANCH, C_SUPPLY_TAP_END, C_SUPPLY_TAP_CONICAL, C_SUPPLY_TAP_MAIN,
    C_EXHAUST_TAP_BRANCH, C_EXHAUST_TAP_END,
    C_EXHAUST_TAP_CONICAL, C_EXHAUST_TAP_CONICAL_END, C_EXHAUST_TAP_MAIN,
    C_TRANSITION_EXPANSION, C_TRANSITION_CONTRACTION,
    C_WYE_BULLHEAD, C_OFFSET, C_ABRUPT_EXIT,
)


# ── Colin's rulings, 2026-09-24 ──────────────────────────────────────────────
#
# 1. "assume vaned"  -> every rectangular elbow gets C_RECT_ELBOW_90 (0.25),
#    the vaned value, including the plain `Rectangular Elbow - Mitered` family.
#    A bare mitered elbow is physically worse than a vaned one; the worksheet
#    publishes a single rect elbow C and this is the firm's number.
# 2. "thats fine .33" -> every round elbow gets 0.33 regardless of radius
#    ratio, so `Round Elbow / 1 D` and `/ 1.5 D` are treated alike. 0.33 is the
#    1D value, so 1.5D elbows are conservative.
# 3. "use dovetail"  -> supply take-offs use the DOVETAIL coefficients (0.98 /
#    0.54), never the conical 1.40, whatever the branch shape. The conical
#    values stay above as data in case that is revisited.
ASSUME_VANED_RECT_ELBOWS = True
ROUND_ELBOW_C_IGNORES_RADIUS = True
USE_DOVETAIL_TAKEOFFS = True


# ── Classifying a fitting by family name ─────────────────────────────────────
#
# Colin, 2026-09-24: "for elbows and tees we have standard ones we use across
# all projects so we can use the family. just for mechanical equipment that may
# change we dont want to use that as the definitive name." So family names ARE
# the right key for FITTINGS. The no-family-names rule still holds for
# mechanical equipment and air terminals, which vary per job -- those are
# classified by connector shape and System Classification (see diffuser_tables).
#
# Names below are Revit's stock duct fitting content, confirmed against the
# Grantham 4 MP model. Matching is lowercase substring, FIRST RULE WINS, and the
# order matters:
#   - 'endcap' before any 'cap'-ish rule
#   - 'takeoff'/'tap' before 'damper', so `Round Takeoff w_Damper` is a takeoff
#   - 'wye' before 'transition', so `Round Wye with Transition - Lateral` is a wye
_ROLE_RULES = (
    ('endcap',     ('endcap', 'end cap')),
    ('takeoff',    ('takeoff', 'take-off', 'take off', 'tap ', 'tap-', ' tap')),
    ('wye',        ('wye', 'lateral')),
    ('cross',      ('cross',)),
    ('tee',        ('tee',)),
    ('transition', ('transition',)),
    ('elbow',      ('elbow', 'bend')),
    ('offset',     ('offset',)),
    ('union',      ('union',)),
)

# Roles that carry no loss at all. An endcap is a dead end (no air passes
# through it) and a union is a straight coupling.
ZERO_LOSS_ROLES = frozenset(['endcap', 'union'])


def classify_fitting(family_name):
    """Fitting role from its Revit family name.

    Returns one of: 'endcap', 'takeoff', 'wye', 'cross', 'tee', 'transition',
    'elbow', 'offset', 'union', or 'unknown'.

    'unknown' is logged and costs 0.0, and the caller must report it rather than
    let it quietly vanish -- an unpriced fitting is a hole in the total, not a
    fitting that happens to be free.
    """
    if not family_name:
        log.info('classify_fitting: empty family name -> unknown')
        return 'unknown'
    name = family_name.lower()
    for role, keywords in _ROLE_RULES:
        for kw in keywords:
            if kw in name:
                return role
    log.info('classify_fitting: NO RULE MATCHED family %r -> unknown', family_name)
    return 'unknown'


def _is_supply(sys_class):
    """Supply side? Anything else uses the exhaust/return coefficients.

    The worksheet publishes separate supply and exhaust take-off coefficients
    and they are far apart (0.98 vs 0.76 on the branch, 3.63 vs 0.54 at the end
    of a main), so getting this wrong matters.
    """
    return sys_class == 'Supply Air'


def takeoff_c(sys_class, is_end_of_main=False, conical=False):
    """C for the air that TURNS OFF through a take-off into a branch.

    is_end_of_main: the branch leaves at the terminus of the main, i.e. the
    main does not continue past it.
    conical: use the conical rect/round coefficients instead of dovetail. Off
    by default per Colin's "use dovetail".
    """
    if _is_supply(sys_class):
        if conical and not USE_DOVETAIL_TAKEOFFS:
            return C_SUPPLY_TAP_CONICAL
        return C_SUPPLY_TAP_END if is_end_of_main else C_SUPPLY_TAP_BRANCH
    if conical and not USE_DOVETAIL_TAKEOFFS:
        return C_EXHAUST_TAP_CONICAL_END if is_end_of_main else C_EXHAUST_TAP_CONICAL
    return C_EXHAUST_TAP_END if is_end_of_main else C_EXHAUST_TAP_BRANCH


def takeoff_main_c(sys_class):
    """C for the air that CONTINUES STRAIGHT past a take-off, staying in the main.

    Applied once per take-off the traced path passes BY, never together with
    takeoff_c() on the same fitting: those are two different air streams through
    one junction, and charging both to one path double-counts it.

    Revit does not model this as an element. A tap is a 2-connector in-line stub
    hanging off an unbroken main, so there is nothing on the main to find. The
    caller counts the taps attached to each main duct it traverses instead.
    """
    return C_SUPPLY_TAP_MAIN if _is_supply(sys_class) else C_EXHAUST_TAP_MAIN


def elbow_c(is_round):
    """C for a 90 degree elbow. Rect elbows assumed vaned, round ignores radius."""
    return C_ROUND_ELBOW_90 if is_round else C_RECT_ELBOW_90


def transition_c(upstream_area_ft2, downstream_area_ft2):
    """C for a transition, expansion or contraction by which way the air goes.

    Both areas come off the fitting's own two connectors, so this needs no
    look at the neighbouring ducts -- but it DOES need the flow direction, which
    only the graph knows (every Revit connector reports Bidirectional).

    Returns 0.0 when either area is unreadable: a transition that cannot be
    measured is reported as unpriced, not silently charged the wrong direction.
    """
    if not upstream_area_ft2 or not downstream_area_ft2:
        return 0.0
    if downstream_area_ft2 > upstream_area_ft2:
        return C_TRANSITION_EXPANSION
    if downstream_area_ft2 < upstream_area_ft2:
        return C_TRANSITION_CONTRACTION
    return 0.0          # same size both ends: a coupling, not a transition


def fitting_c(role, sys_class, is_round=True, upstream_area_ft2=None,
              downstream_area_ft2=None, is_end_of_main=False):
    """One dispatch point: C for a fitting the traced path passes THROUGH.

    Returns (c_value, note). `note` names what was priced, or why nothing was,
    so the caller can print it and the total is never silently short.
    """
    if role in ZERO_LOSS_ROLES:
        return 0.0, 'no loss (%s)' % role
    if role == 'elbow':
        return elbow_c(is_round), '90 deg %s elbow' % ('round' if is_round else 'rect')
    if role == 'takeoff':
        c = takeoff_c(sys_class, is_end_of_main=is_end_of_main)
        return c, 'take-off into branch%s' % (' (end of main)' if is_end_of_main else '')
    if role in ('tee', 'cross'):
        # A tee or cross the path turns through behaves as a take-off branch;
        # the worksheet publishes no separate tee coefficient.
        c = takeoff_c(sys_class, is_end_of_main=is_end_of_main)
        return c, '%s, priced as a take-off branch' % role
    if role == 'wye':
        return C_WYE_BULLHEAD, 'bull head 90 deg wye'
    if role == 'offset':
        return C_OFFSET, 'offset'
    if role == 'transition':
        c = transition_c(upstream_area_ft2, downstream_area_ft2)
        if c == C_TRANSITION_EXPANSION:
            return c, 'transition expansion'
        if c == C_TRANSITION_CONTRACTION:
            return c, 'transition contraction'
        return 0.0, 'transition, size change unreadable (UNPRICED)'
    return 0.0, 'unrecognised fitting (UNPRICED)'
