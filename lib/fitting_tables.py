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
# One row per fitting: (key, label shown in the dialog, 1985 ASHRAE fitting
# number as the worksheet lists it, published C). The label and number are here
# so the dialog can present the table without restating it, and so an override
# typed at runtime is always traceable back to the published figure it replaced.
C_TABLE = (
    ('rect_elbow_90',        '90 deg rect elbow',                        '3-5/3-7', 0.25),
    ('round_elbow_90',       '90 deg round elbow',                       '3-1',     0.33),
    ('supply_tap_branch',    'Supply take-off, typ branch',              '6-29',    0.98),
    ('supply_tap_end',       'Supply take-off, end of main',             '6-25',    0.54),
    ('supply_tap_main',      'Main duct @ supply take-off',              '6-23',    0.28),
    ('supply_tap_conical',   'Supply conical take-off (rect/round)',     '6-27',    1.40),
    ('exhaust_tap_branch',   'Exhaust take-off, typ branch',             '6-9',     0.76),
    ('exhaust_tap_end',      'Exhaust take-off, end of main',            '6-9',     3.63),
    ('exhaust_tap_main',     'Main duct @ exhaust take-off',             '6-3',     0.60),
    ('exhaust_tap_conical',     'Exhaust conical take-off, typ branch',  '6-7',     1.27),
    ('exhaust_tap_conical_end', 'Exhaust conical take-off, end of main', '6-7',     5.95),
    ('transition_expansion',   'Transition expansion (60 deg max)',      '4-3',     0.84),
    ('transition_contraction', 'Transition contraction (60 deg max)',    '5-1',     0.08),
    ('wye_bullhead',         'Bull head 90 deg wye',                     '6-33',    0.30),
    ('offset',               'Offset',                                   '3-13',    1.70),
    ('abrupt_exit',          'Abrupt exit (stacks, no cone)',            '2-11',    1.00),
)

DEFAULT_C = {}
for _k, _lbl, _no, _v in C_TABLE:
    DEFAULT_C[_k] = _v

# The worksheet's own checksum cell (AN30). Asserted by the self-test against
# the DEFAULTS, so a typo in the table above cannot pass silently. Runtime
# overrides are deliberately not checked against it - the engineer may have a
# reason to depart from the published figure, and that is their call.
PUBLISHED_C_SUM = 19.91


def c_of(key, c=None):
    """One C value: the caller's override if given, else the published default.

    Unknown keys return 0.0 rather than raising, because a missing coefficient
    must show up as an unpriced fitting in the report, not as a crash mid-run.
    """
    if c is not None and key in c:
        return c[key]
    return DEFAULT_C.get(key, 0.0)


# ── Component pressure drops ─────────────────────────────────────────────────
#
# Not C values and not geometry. These are the worksheet's column D ("INPUT
# KNOWN PRESS. DROP") entries, in in. wc, taken from the two rows of its filled
# example. The model supplies the COUNT of each; these supply the value.
#
# Treated as GIVENS with an override, per Colin 2026-09-24 ("remove diffuser and
# balancing drops inputs and have them as givens ... a drop down with input
# values for components"), rather than as required entries on the main dialog.
#
# NOT read from the Revit families, both checked against a live model: air
# terminals carry "Total Pressure = 0.09 in-wg" but on only 11 of 22 terminals
# and as a non-builtin parameter whose internal unit could not be verified, and
# `Balancing Damper - Round` carries "Pressure Drop = 1.00 in-wg" which reads as
# content boilerplate (15 of those would swamp the whole system, and the builtin
# RBS_DUCT_PRESSURE_DROP is empty).
COMPONENT_TABLE = (
    ('diffuser', 'Diffuser / grille', 0.05),
    ('damper',   'Balancing damper',  0.25),
)

DEFAULT_COMPONENTS = {}
for _k, _lbl, _v in COMPONENT_TABLE:
    DEFAULT_COMPONENTS[_k] = _v


def component_of(key, comps=None):
    """One component drop in in. wc: override if given, else the given default."""
    if comps is not None and key in comps:
        return comps[key]
    return DEFAULT_COMPONENTS.get(key, 0.0)


# ── Colin's rulings, 2026-09-24 ──────────────────────────────────────────────
#
# 1. "assume vaned"  -> every rectangular elbow gets the rect elbow C (0.25),
#    the vaned value, including the plain `Rectangular Elbow - Mitered` family.
#    A bare mitered elbow is physically worse than a vaned one; the worksheet
#    publishes a single rect elbow C and this is the firm's number.
# 2. "thats fine .33" -> every round elbow gets 0.33 regardless of radius
#    ratio, so `Round Elbow / 1 D` and `/ 1.5 D` are treated alike. 0.33 is the
#    1D value, so 1.5D elbows are conservative.
# 3. "use dovetail"  -> supply take-offs use the DOVETAIL coefficients, never
#    the conical 1.40, whatever the branch shape. The conical values stay in
#    C_TABLE as data, and are now editable, in case that is revisited.
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
# Names below are Revit's stock duct fitting content, confirmed against a live
# model. Matching is lowercase substring, FIRST RULE WINS, and order matters:
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


def takeoff_c(sys_class, is_end_of_main=False, conical=False, c=None):
    """C for the air that TURNS OFF through a take-off into a branch."""
    side = 'supply' if _is_supply(sys_class) else 'exhaust'
    if conical and not USE_DOVETAIL_TAKEOFFS:
        if side == 'supply':
            return c_of('supply_tap_conical', c)
        return c_of('exhaust_tap_conical_end' if is_end_of_main
                    else 'exhaust_tap_conical', c)
    return c_of('%s_tap_%s' % (side, 'end' if is_end_of_main else 'branch'), c)


def takeoff_main_c(sys_class, c=None):
    """C for the air that CONTINUES STRAIGHT past a take-off, staying in the main.

    Applied once per take-off the traced path passes BY, never together with
    takeoff_c() on the same fitting: those are two different air streams through
    one junction, and charging both to one path double-counts it.

    Revit does not model this as an element. A tap is a 2-connector in-line stub
    hanging off an unbroken main, so there is nothing on the main to find. The
    caller counts the taps attached to each main duct it traverses instead.
    """
    return c_of('supply_tap_main' if _is_supply(sys_class)
                else 'exhaust_tap_main', c)


def elbow_c(is_round, c=None):
    """C for a 90 degree elbow. Rect elbows assumed vaned, round ignores radius."""
    return c_of('round_elbow_90' if is_round else 'rect_elbow_90', c)


def transition_c(upstream_area_ft2, downstream_area_ft2, c=None):
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
        return c_of('transition_expansion', c)
    if downstream_area_ft2 < upstream_area_ft2:
        return c_of('transition_contraction', c)
    return 0.0          # same size both ends: a coupling, not a transition


def fitting_c(role, sys_class, is_round=True, upstream_area_ft2=None,
              downstream_area_ft2=None, is_end_of_main=False, c=None):
    """One dispatch point: C for a fitting the traced path passes THROUGH.

    Returns (c_value, note). `note` names what was priced, or why nothing was,
    so the caller can print it and the total is never silently short.
    """
    if role in ZERO_LOSS_ROLES:
        return 0.0, 'no loss (%s)' % role
    if role == 'elbow':
        return (elbow_c(is_round, c),
                '90 deg %s elbow' % ('round' if is_round else 'rect'))
    if role == 'takeoff':
        return (takeoff_c(sys_class, is_end_of_main=is_end_of_main, c=c),
                'take-off into branch%s' % (' (end of main)' if is_end_of_main else ''))
    if role in ('tee', 'cross'):
        # A tee or cross the path turns through behaves as a take-off branch;
        # the worksheet publishes no separate tee coefficient.
        return (takeoff_c(sys_class, is_end_of_main=is_end_of_main, c=c),
                '%s, priced as a take-off branch' % role)
    if role == 'wye':
        return c_of('wye_bullhead', c), 'bull head 90 deg wye'
    if role == 'offset':
        return c_of('offset', c), 'offset'
    if role == 'transition':
        val = transition_c(upstream_area_ft2, downstream_area_ft2, c)
        if val == c_of('transition_expansion', c) and val != 0.0:
            return val, 'transition expansion'
        if val == c_of('transition_contraction', c) and val != 0.0:
            return val, 'transition contraction'
        return 0.0, 'transition, size change unreadable (UNPRICED)'
    return 0.0, 'unrecognised fitting (UNPRICED)'


# ── Duct accessories: components, not fittings ───────────────────────────────
#
# A balancing damper is not a fitting and has no C value in the worksheet. RJA's
# sheet carries it as a hand-entered known pressure drop in column D (its example
# uses 0.25 in. wc for an OBD), because a damper's drop is a cutsheet number and
# a function of how far it is throttled, not of duct geometry.
#
# NOT read from the Revit family. The `Balancing Damper - Round` family in
# Grantham 4 MP carries an instance parameter "Pressure Drop = 1.00 in-wg",
# which is almost certainly Revit content boilerplate rather than a real
# selection: a balancing damper near wide open is an order of magnitude below
# that, and 15 of them at 1.00 would swamp every other loss in the system. The
# built-in RBS_DUCT_PRESSURE_DROP is empty, which is what Revit would populate
# if its own pressure-loss calc had ever been run. So the drop is a user input,
# and the count comes from the model.
_BALANCING_DAMPER_KEYWORDS = ('balancing damper', 'balance damper', 'obd',
                              'opposed blade')


def is_balancing_damper(family_name):
    """True for a balancing / opposed-blade damper accessory.

    Deliberately narrow. A fire damper, smoke damper or backdraft damper has a
    different drop and must not silently inherit the balancing-damper input, so
    anything else comes back False and the caller reports it as uncounted.
    """
    if not family_name:
        return False
    name = family_name.lower()
    for kw in _BALANCING_DAMPER_KEYWORDS:
        if kw in name:
            return True
    return False
