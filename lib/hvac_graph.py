# -*- coding: utf-8 -*-
"""
hvac_graph.py  --  HVAC duct traversal engine (shared by HVAC Diagnose + Duct Velocity)

Traversal mirrors gas pipe_graph.py:
  - Root = AHU (MechanicalEquipment / BaseEquipment)
  - BFS outward through all HVAC connectors
  - VAV boxes / FCUs treated as pass-through nodes
  - CFM is read only from OST_DuctTerminal leaf nodes via LookupParameter("Flow")
  - Each duct segment CFM = sum of all reachable downstream terminal CFMs

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import math
import logging

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter,
    FillPatternElement, ElementId, Domain, ConnectorProfileType,
    FlowDirectionType
)

from revit_helpers import eid_int

import diffuser_tables

log = logging.getLogger(__name__)

# ── Firm design defaults ─────────────────────────────────────────────────────
# {sys_class: (max_fpm, max_friction_inwc_per_100ft)}
#
# These are the MAIN-duct limits only. Branch ducts (a run feeding exactly one
# terminal) are no longer checked against them at all — see
# branch_diffuser_check() below. A correctly-selected auto-sizing diffuser is
# picked on static pressure and NC rather than on duct velocity, so velocity-
# checking a run that only ever feeds that one diffuser is redundant and
# produces false flags; branches are instead checked for size match against
# the published diffuser tables in diffuser_tables.py. The one exception is a
# slot-diffuser branch: the design standard publishes no duct/neck size
# breakpoints for slot diffusers, so those fall back to these same values.
FIRM_DEFAULTS = {
    'Supply Air':    (800,  0.08),
    'Return Air':    (600,  0.05),
    'Exhaust Air':   (600,  0.05),
    'Outside Air':   (600,  0.05),
    'Transfer Air':  (400,  0.05),
}

# SMACNA labels for Diagnose report — derived from FIRM_DEFAULTS with 85% green band
# (green_fpm = max*0.85, yellow_fpm = max)
SMACNA = {
    'Supply Air':    (680, 800),    # 800 * 0.85 = 680
    'Return Air':    (510, 600),    # 600 * 0.85 = 510
    'Exhaust Air':   (510, 600),
    'Outside Air':   (510, 600),
    'Transfer Air':  (340, 400),    # 400 * 0.85 = 340
}

# ── Flex duct sizing (SA and RA only — for future duct sizer tool) ────────────
# ⚠ NOT wired to the velocity calculator. Reference constant only.
# Source: firm standard. CFM > 750 → flag for rigid duct.
_FLEX_DUCT_TABLE = [
    (100,  6),
    (225,  8),
    (400, 10),
    (600, 12),
    (750, 14),
]
FLEX_DUCT_MAX_CFM = 750


def flex_duct_size(cfm):
    """Return minimum recommended flex duct diameter (inches) for given CFM.
    Applies to Supply Air and Return Air only.
    Returns (diameter_in, warning_or_None).
    If CFM > 750, diameter is None and warning recommends rigid duct.
    """
    if cfm > FLEX_DUCT_MAX_CFM:
        return None, 'CFM {:.0f} exceeds flex duct max ({} CFM) — use rigid duct'.format(
            cfm, FLEX_DUCT_MAX_CFM)
    for max_cfm, diam in _FLEX_DUCT_TABLE:
        if cfm <= max_cfm:
            return diam, None
    return None, 'CFM {:.0f} not matched in flex duct table'.format(cfm)

# ── Revit category IDs ───────────────────────────────────────────────────────
_CAT_TERMINAL  = int(BuiltInCategory.OST_DuctTerminal)
_CAT_DUCT      = int(BuiltInCategory.OST_DuctCurves)
_CAT_EQUIP     = int(BuiltInCategory.OST_MechanicalEquipment)
_CAT_FLEX_DUCT = int(BuiltInCategory.OST_FlexDuctCurves)
_CAT_FITTING   = int(BuiltInCategory.OST_DuctFitting)
_CAT_ACCESSORY = int(BuiltInCategory.OST_DuctAccessory)


# ── unit conversion ──────────────────────────────────────────────────────────
def to_cfm(raw, cfm_is_direct=False):
    """Convert Revit internal flow value to CFM.

    Revit stores air-flow params in ft3/s internally; display is ft3/min (CFM).
    Set cfm_is_direct=True if the Flow param is a plain Number (already in CFM).

    UnitTypeId.CubicFeetPerMinute has existed since Revit 2021, so it alone
    covers this firm's whole 2022-2026 range - no version branch needed.
    A DisplayUnitType fallback used to sit here for pre-UnitTypeId Revit;
    removed 2026-07-23 after confirming DisplayUnitType is now
    inaccessible (protection level restricted) in the live Revit 2026 API,
    so it was a dead branch that would have failed anyway if ever reached.
    """
    if cfm_is_direct:
        return raw
    from Autodesk.Revit.DB import UnitTypeId, UnitUtils
    try:
        return UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.CubicFeetPerMinute)
    except Exception:
        pass
    return raw * 60.0


# ── helpers ──────────────────────────────────────────────────────────────────
def _connector_manager(elem):
    try:
        return elem.ConnectorManager
    except Exception:
        pass
    try:
        return elem.MEPModel.ConnectorManager
    except Exception:
        pass
    return None


def _cat_id(elem):
    try:
        return eid_int(elem.Category.Id)
    except Exception:
        return -1


def is_terminal(elem):
    return _cat_id(elem) == _CAT_TERMINAL

def is_duct(elem):
    cid = _cat_id(elem)
    return cid in (_CAT_DUCT, _CAT_FLEX_DUCT)

def is_equipment(elem):
    return _cat_id(elem) == _CAT_EQUIP

def is_fitting(elem):
    return _cat_id(elem) == _CAT_FITTING

def is_accessory(elem):
    return _cat_id(elem) == _CAT_ACCESSORY

def is_fitting_or_accessory(elem):
    cid = _cat_id(elem)
    return cid in (_CAT_FITTING, _CAT_ACCESSORY)


def _nominal_even_in(inches):
    """Snap a raw dimension (already converted to inches) to the nearest 2"
    nominal duct/neck increment (6, 8, 10, 12, 14, ...) — RJA standard duct
    sizing grid, matching _ROUND_SIZES / _snap_even() in Duct Velocity's
    script.py. Rounds to NEAREST, not up — this is meant to recover the
    intended nominal value (any non-nominal reading is either Revit's own
    floating-point unit-conversion noise or a modeling slip, not a real
    odd size; lined duct's actual clear opening is detailed/keynoted
    separately rather than modeled as a different dimension) — not to
    conservatively pad a candidate suggestion the way _snap_even()'s
    ceiling behavior does for "find the smallest size that still works"
    searches.

    Confirmed live 2026-08-26: a true 10" duct can be stored internally as
    10.000000000000002" — enough to flip a strict `<` comparison at an
    exact 2" boundary (e.g. 12" vs 10"+2"=12.000000000000002") and produce
    a false diffuser/duct clearance flag on a step-down that's actually
    exactly at the safe margin. Snapping to the nominal grid kills that
    class of noise categorically, not just the one boundary it was found on.
    """
    return round(inches / 2.0) * 2.0


def effective_height_in(elem):
    """Height (rectangular) or diameter (round) in inches, used for the
    diffuser/duct height clearance check. Ducts read RBS_CURVE_* built-in
    params directly; terminals (which don't have those params) read their
    neck connector geometry instead. Returns None if not determinable.

    Snapped to the nearest 2" nominal increment via _nominal_even_in() —
    see that function's docstring for why.
    """
    if is_duct(elem):
        d = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            return _nominal_even_in(d.AsDouble() * 12.0)
        h = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if h is not None and h.AsDouble() > 0:
            return _nominal_even_in(h.AsDouble() * 12.0)
        return None
    if is_terminal(elem):
        cm = _connector_manager(elem)
        if cm is None:
            return None
        try:
            for c in cm.Connectors:
                if c.Shape == ConnectorProfileType.Round:
                    return _nominal_even_in(c.Radius * 2.0 * 12.0)
                elif c.Shape == ConnectorProfileType.Rectangular:
                    return _nominal_even_in(c.Height * 12.0)
        except Exception:
            pass
    return None


# One nominal duct/neck step, inches. Same 2" grid _nominal_even_in() snaps to.
# Used as the "clearly a size up" threshold for the informational branch
# oversize flag, so that flag can never fire on float noise alone.
NOMINAL_STEP_IN = 2.0


def terminal_connector_geometry(elem):
    """Shape and size of an air terminal's own connection, read straight off
    its connector rather than off any family/type name.

    Returns one of:
        ('round', diameter_in)
        ('oval',  None)              -- see below on why size is not read
        ('rect',  width_in, height_in)
        (None,    None)              -- no connector, or no readable shape

    Note the varying tuple length: callers should branch on element [0] and
    only unpack the rest once they know the shape. Sizes are snapped to the
    nominal 2" grid by _nominal_even_in() for the same float-noise reason
    documented on that function.

    Shape comes from Connector.Shape (ConnectorProfileType). This is the only
    classification signal used for diffusers anywhere in this tool — family
    and type names are deliberately never matched, because they have been
    renamed repeatedly across this project's history and string-matching them
    would be fragile by design.

    ConnectorProfileType members confirmed by reflection against RevitAPI.dll
    for both Revit 2022 and 2026 (2026-09-23): Round, Rectangular, Oval,
    Invalid. Oval is real and distinct from Round — RJA's slot diffuser neck
    connector is physically Oval even though the duct feeding it is round,
    which is what makes slot diffusers identifiable without a name match.

    No size is read for Oval: the design standard publishes no neck-size
    breakpoints for slot diffusers at all (capacity is given per slot length
    and open-slot count instead), so there is nothing to compare a size
    against and reading one would only invite a check that cannot be made.

    Reading Radius on a rectangular connector (or Width/Height on a round one)
    throws, so each shape is handled inside its own branch — never read one
    set of properties and hope.
    """
    eid = eid_int(elem.Id)
    cm = _connector_manager(elem)
    if cm is None:
        log.info('terminal_connector_geometry: id=%s NOT FOUND (no ConnectorManager)', eid)
        return None, None
    try:
        for c in cm.Connectors:
            try:
                shape = c.Shape
            except Exception as ex:
                log.debug('terminal_connector_geometry: id=%s connector Shape unreadable: %s',
                          eid, ex)
                continue
            if shape == ConnectorProfileType.Round:
                dia = _nominal_even_in(c.Radius * 2.0 * 12.0)
                log.info('terminal_connector_geometry: id=%s Round dia=%s in', eid, dia)
                return 'round', dia
            if shape == ConnectorProfileType.Oval:
                log.info('terminal_connector_geometry: id=%s Oval (slot diffuser, no size read)', eid)
                return 'oval', None
            if shape == ConnectorProfileType.Rectangular:
                w = _nominal_even_in(c.Width * 12.0)
                h = _nominal_even_in(c.Height * 12.0)
                log.info('terminal_connector_geometry: id=%s Rectangular %sx%s in', eid, w, h)
                return 'rect', w, h
    except Exception as ex:
        log.warning('terminal_connector_geometry: id=%s connector iteration failed: %s', eid, ex)
    log.info('terminal_connector_geometry: id=%s NOT FOUND (no connector with a readable shape)', eid)
    return None, None


_SHAPE_TO_CATEGORY = {
    'round': 'ceiling',
    'oval':  'slot',
    'rect':  'sidewall',
}


def classify_diffuser_branch(elem):
    """Which diffuser sizing table applies to this terminal: 'ceiling',
    'sidewall', 'slot', or None.

    A thin mapping over terminal_connector_geometry()'s shape — round neck is
    a ceiling diffuser/grate, oval neck is a slot diffuser, rectangular face
    is a sidewall grille. No name matching and no aspect-ratio heuristic: a
    rectangular face is a sidewall grille whatever its proportions, and the
    oval/round distinction is what separates slot from ceiling.

    None means the terminal has no readable connection geometry, so no table
    can be chosen. Callers must treat that as "cannot check", not as a pass.
    """
    kind = terminal_connector_geometry(elem)[0]
    cat  = _SHAPE_TO_CATEGORY.get(kind)
    if cat is None:
        log.info('classify_diffuser_branch: id=%s NOT CLASSIFIED (no connector geometry)',
                 eid_int(elem.Id))
    return cat


def next_real_downstream(node_id, nodes, children):
    """Return the list of real duct/terminal elements immediately downstream
    of node_id, walking through (skipping over) any fittings/accessories in
    between. A duct/terminal that branches into multiple fittings each
    leading somewhere real all get returned — the caller decides how to
    combine them (e.g. take the max height requirement).
    """
    result = []
    for cid in children.get(node_id, []):
        celem = nodes.get(cid)
        if celem is None:
            continue
        if is_fitting_or_accessory(celem):
            result.extend(next_real_downstream(cid, nodes, children))
        elif is_duct(celem) or is_terminal(celem):
            result.append(celem)
        # equipment or anything else: not a duct/terminal, ignore
    return result


def max_downstream_height_in(node_id, nodes, children):
    """Return the largest effective_height_in() among the real duct/terminal
    elements immediately downstream of node_id (skipping fittings), or None
    if no real downstream element with a determinable height is found. Used
    as the basis for the diffuser/duct height clearance check — the caller
    only needs to enforce a margin where this value is actually SMALLER than
    the duct being checked (a genuine size reduction), not wherever it's
    equal or larger (a continuous run needs no transition clearance).
    """
    best = None
    for elem in next_real_downstream(node_id, nodes, children):
        h = effective_height_in(elem)
        if h is None:
            continue
        if best is None or h > best:
            best = h
    return best


def duct_area_ft2(duct):
    """Cross-section area in ft2. Returns 0.0 if dimensions unavailable."""
    d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if d is not None and d.AsDouble() > 0:
        r = d.AsDouble() * 0.5
        return math.pi * r * r
    w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    if w is not None and h is not None and w.AsDouble() > 0 and h.AsDouble() > 0:
        return w.AsDouble() * h.AsDouble()
    return 0.0

def duct_size_label(duct):
    """Human-readable size string: '24x12' or '10dia'."""
    d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if d is not None and d.AsDouble() > 0:
        return '{:.0f}"dia'.format(d.AsDouble() * 12.0)
    w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    if w is not None and h is not None:
        return '{:.0f}x{:.0f}"'.format(w.AsDouble() * 12.0, h.AsDouble() * 12.0)
    return '?'

def duct_installed_size_in(duct):
    """Numeric installed size of a duct, for dimensional comparison against a
    diffuser's own connection size.

    Returns diameter_in (a number) for round/spiral, (width_in, height_in) for
    rectangular, or None if neither is readable. Snapped to the nominal 2"
    grid by _nominal_even_in() so it compares like-for-like against
    terminal_connector_geometry(), which snaps the same way.

    duct_size_label() already covers the display side but returns a formatted
    string; this is the numeric counterpart the branch check needs. Kept here
    rather than in the pushbutton so the nominal-grid read happens in exactly
    one place.
    """
    d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if d is not None and d.AsDouble() > 0:
        return _nominal_even_in(d.AsDouble() * 12.0)
    w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    if w is not None and h is not None and w.AsDouble() > 0 and h.AsDouble() > 0:
        return (_nominal_even_in(w.AsDouble() * 12.0),
                _nominal_even_in(h.AsDouble() * 12.0))
    log.info('duct_installed_size_in: id=%s NOT FOUND (no diameter and no width/height)',
             eid_int(duct.Id))
    return None


def duct_sys_class(duct):
    """Returns system classification string e.g. 'Supply Air'."""
    p = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
    if p is not None:
        v = p.AsValueString()
        if v:
            return v
    return 'Unknown'

def terminal_family_name(elem):
    try:
        return elem.Symbol.Family.Name
    except Exception:
        return 'Unknown'

def terminal_sys_class(elem):
    p = elem.LookupParameter('System Classification')
    if p is not None:
        v = p.AsString()
        if v:
            return v
    return 'Unknown'

def elem_sys_class(elem):
    """System classification for a duct or terminal, or None if undetermined.
    Unlike duct_sys_class()/terminal_sys_class(), returns None (not 'Unknown')
    when unresolved, so callers can tell 'no data' apart from a real class."""
    if is_duct(elem):
        v = duct_sys_class(elem)
        return v if v and v != 'Unknown' else None
    if is_terminal(elem):
        v = terminal_sys_class(elem)
        return v if v and v != 'Unknown' else None
    return None


def _branch_sys_class(node_id, nodes, children):
    """Resolve the system classification of a branch hanging off an equipment
    node, walking through fittings/accessories (which have no class of their
    own) to the nearest classified duct or terminal. Returns None if nothing
    classified is found along that branch."""
    elem = nodes.get(node_id)
    if elem is None:
        return None
    cls = elem_sys_class(elem)
    if cls is not None:
        return cls
    if is_fitting_or_accessory(elem):
        for cid in children.get(node_id, []):
            cls = _branch_sys_class(cid, nodes, children)
            if cls is not None:
                return cls
    return None

def smacna_label(fpm, sys_class):
    limits = SMACNA.get(sys_class, None)
    if fpm <= 0 or limits is None:
        return 'GRAY'
    lo, hi = limits
    if fpm <= lo:
        return 'GREEN'
    elif fpm <= hi:
        return 'YELLOW'
    else:
        return 'RED'


def duct_friction_loss_per_100ft(v_fpm, d_h_in):
    """Friction loss in in. wc per 100 ft.

    Formula: 6.82e-6 * V_fpm^1.82 / D_h_in^1.22
    Derived from ASHRAE smooth-duct correlation for standard air
    (70°F, sea level, galvanized sheet metal roughness).
    Calibration: 10" duct at 910 FPM → 0.099 in. wc/100ft (SMACNA 0.1 target).
    """
    if v_fpm <= 0 or d_h_in <= 0:
        return 0.0
    return 6.82e-6 * (v_fpm ** 1.82) / (d_h_in ** 1.22)


def _duct_length_ft(duct):
    """Duct centerline length in feet (Revit internal units). 0.0 if unavailable."""
    try:
        curve = duct.Location.Curve
        if curve is not None:
            return curve.Length
    except Exception:
        pass
    return 0.0


def _duct_d_h_in(duct):
    """Hydraulic diameter in inches from duct element parameters. 0 if unavailable."""
    d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if d is not None and d.AsDouble() > 0:
        return d.AsDouble() * 12.0   # ft → in
    w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    if w is not None and h is not None and w.AsDouble() > 0 and h.AsDouble() > 0:
        w_in = w.AsDouble() * 12.0
        h_in = h.AsDouble() * 12.0
        return 4.0 * w_in * h_in / (2.0 * (w_in + h_in))
    return 0.0


# ── branch diffuser check ────────────────────────────────────────────────────
#
# Everything a branch duct (a run feeding exactly one terminal) is judged on
# lives below, so the pushbutton only has to call branch_diffuser_check() and
# render the answer. Mains are untouched by any of this — they keep the
# FIRM_DEFAULTS velocity/friction check.

# Which published table applies, by (diffuser category, terminal system
# classification). Return and Exhaust share one published dataset.
#
# Outside Air and Transfer Air are deliberately absent: the design standard
# publishes no diffuser capacity table for either, so a terminal classified
# that way comes back "cannot check" rather than being quietly judged against
# a table that was never meant for it.
_BRANCH_TABLES = {
    ('ceiling',  'Supply Air'):   diffuser_tables.CEILING_SUPPLY_NECK_TABLE,
    ('ceiling',  'Return Air'):   diffuser_tables.CEILING_RETURN_EXHAUST_NECK_TABLE,
    ('ceiling',  'Exhaust Air'):  diffuser_tables.CEILING_RETURN_EXHAUST_NECK_TABLE,
    ('sidewall', 'Supply Air'):   diffuser_tables.SIDEWALL_SUPPLY_TABLE,
    ('sidewall', 'Return Air'):   diffuser_tables.SIDEWALL_RETURN_EXHAUST_TABLE,
    ('sidewall', 'Exhaust Air'):  diffuser_tables.SIDEWALL_RETURN_EXHAUST_TABLE,
}


class BranchDiffuserResult(object):
    """One branch duct's verdict from branch_diffuser_check().

    status is 'GREEN', 'RED' or 'GRAY' — never YELLOW. There is no tolerance
    band on a size match: a branch duct either is at least the size the
    diffuser connects with or it is not. GRAY means the published data does
    not cover this selection, which is a real third answer and must not be
    rendered as a pass.

    Field roles, since several of them look interchangeable and are not:
      diffuser_size  what the diffuser itself connects with (its own neck/face)
      installed_size what the branch duct actually is
      required_size  what the branch duct should be — the larger of the size
                     the CFM demands and the size the diffuser connects with
    """

    def __init__(self):
        self.status           = 'GRAY'
        self.reason           = ''
        self.category         = None    # 'ceiling' / 'sidewall' / 'slot' / None
        self.sys_class        = 'Unknown'
        self.cfm              = 0.0
        self.diffuser_size    = '-'
        self.installed_size   = '-'
        self.required_size    = '-'
        self.diffuser_max_cfm = None    # published capacity of the installed diffuser


def _ge_in(a, b):
    """a >= b for inch dimensions, tolerant of Revit float noise. Both sides
    are nominal-snapped before they get here, so the tolerance only has to
    absorb representation error, never a real half-size difference."""
    return float(a) >= float(b) - diffuser_tables.SIZE_TOL_IN


def branch_diffuser_check(terminal_elem, terminal_cfm, installed_branch_size):
    """Judge one branch duct against the published diffuser tables.

    Two independent parts, both of which must pass for GREEN:

      (a) Diffuser adequacy — is the diffuser's own installed neck/face big
          enough for its own Flow, per the design standard? This is checked
          against diffuser_tables rather than against the family's own sizing
          formula on purpose; see that module's docstring.

      (b) Branch duct match — is the installed duct at least as big as the
          connection the diffuser presents? Sidewall grilles are matched
          dimension-to-dimension because RJA matches the grille's rectangular
          size to the connecting ductwork directly, not through an
          area-equivalent adapter calculation.

    Either part failing is RED, with a reason naming which. Either part being
    unverifiable (a size the standard does not publish, an unreadable duct
    size) is GRAY with a reason naming what could not be read — except where
    the other part outright fails, in which case the real failure wins.

    Oversize is never checked or reported on a branch: one fed by the
    auto-sizing diffuser family is routed back to whatever size the diffuser
    computed, so it should not occur.

    Args:
        terminal_elem:  the OST_DuctTerminal element at the end of the branch.
        terminal_cfm:   that terminal's Flow in CFM (already unit-converted).
        installed_branch_size: the branch duct's size as returned by
            duct_installed_size_in() — a number for round, a (w_in, h_in)
            sequence for rectangular, or None if unreadable.

    Returns:
        BranchDiffuserResult. Never raises for bad model data; unreadable
        inputs come back GRAY with a reason.
    """
    res = BranchDiffuserResult()
    tid = eid_int(terminal_elem.Id)
    try:
        res.cfm = float(terminal_cfm)
    except (TypeError, ValueError):
        res.cfm = 0.0

    geom         = terminal_connector_geometry(terminal_elem)
    kind         = geom[0]
    res.category = _SHAPE_TO_CATEGORY.get(kind)

    if res.category not in ('ceiling', 'sidewall'):
        res.status = 'GRAY'
        res.reason = 'Cannot classify diffuser'
        log.warning('branch_diffuser_check: terminal id=%s NOT CHECKED — '
                    'connector shape %s maps to category %s', tid, kind, res.category)
        return res

    res.sys_class = terminal_sys_class(terminal_elem)
    table = _BRANCH_TABLES.get((res.category, res.sys_class))
    if table is None:
        res.status = 'GRAY'
        res.reason = 'No table for {}'.format(res.sys_class)
        log.warning('branch_diffuser_check: terminal id=%s NOT CHECKED — no published '
                    '%s table for system class "%s"', tid, res.category, res.sys_class)
        return res

    # Tri-state on purpose: True = passed, False = failed, None = could not be
    # determined. Collapsing None into either of the other two is exactly the
    # silent-failure this codebase forbids.
    diffuser_ok = None
    branch_ok   = None
    unverified  = []

    if res.category == 'ceiling':
        neck_dia          = geom[1]
        res.diffuser_size = diffuser_tables.round_size_label(neck_dia)

        # (a) is the neck big enough for its own CFM
        res.diffuser_max_cfm = diffuser_tables.max_cfm_for_diameter(table, neck_dia)
        if res.diffuser_max_cfm is None:
            unverified.append('{} neck not in table'.format(res.diffuser_size))
        else:
            diffuser_ok = res.cfm <= res.diffuser_max_cfm

        # Required size — the larger of what the CFM demands and what the
        # diffuser physically connects with. Both constraints bind the branch
        # duct at once, so showing only one of them would understate it.
        min_dia = diffuser_tables.min_diameter_for_cfm(table, res.cfm)
        if min_dia is None:
            res.required_size = '>{}'.format(diffuser_tables.round_size_label(
                diffuser_tables.largest_diameter(table)))
        else:
            res.required_size = diffuser_tables.round_size_label(
                max(min_dia, neck_dia))

        # (b) is the branch duct at least the neck size
        if installed_branch_size is None:
            unverified.append('branch duct size unreadable')
        elif isinstance(installed_branch_size, (tuple, list)):
            res.installed_size = diffuser_tables.rect_size_label(
                installed_branch_size[0], installed_branch_size[1])
            unverified.append('rect duct on round neck')
        else:
            duct_dia          = float(installed_branch_size)
            res.installed_size = diffuser_tables.round_size_label(duct_dia)
            branch_ok          = _ge_in(duct_dia, neck_dia)

    else:   # sidewall
        face_w, face_h = geom[1], geom[2]
        res.diffuser_size = diffuser_tables.rect_size_label(face_w, face_h)

        # (a) is the face big enough for its own CFM
        face_row = diffuser_tables.sidewall_row(table, face_w, face_h)
        if face_row is None:
            near, _exact = diffuser_tables.nearest_sidewall_row(table, face_w, face_h)
            unverified.append('{} face not in table (nearest {})'.format(
                res.diffuser_size, diffuser_tables.row_size_label(near)))
        else:
            res.diffuser_max_cfm = face_row[2]
            diffuser_ok = res.cfm <= res.diffuser_max_cfm

        # Required size — same "larger of the two binding constraints" rule as
        # the round case, compared by face area since the published rows span
        # three aspect ratios and are not comparable dimension-by-dimension.
        min_row = diffuser_tables.min_row_for_cfm(table, res.cfm)
        if min_row is None:
            res.required_size = '>{}'.format(diffuser_tables.row_size_label(
                diffuser_tables.largest_sidewall_row(table)))
        elif face_row is not None and (face_row[0] * face_row[1]) > (min_row[0] * min_row[1]):
            res.required_size = diffuser_tables.row_size_label(face_row)
        else:
            res.required_size = diffuser_tables.row_size_label(min_row)

        # (b) is the branch duct at least the face size
        if installed_branch_size is None:
            unverified.append('branch duct size unreadable')
        elif not isinstance(installed_branch_size, (tuple, list)):
            res.installed_size = diffuser_tables.round_size_label(
                float(installed_branch_size))
            unverified.append('round duct on rect face')
        else:
            duct_w, duct_h     = installed_branch_size[0], installed_branch_size[1]
            res.installed_size = diffuser_tables.rect_size_label(duct_w, duct_h)
            # Either orientation counts. A grille and its duct can legitimately
            # be modelled with width and height swapped relative to each other,
            # and this check is about whether the duct is big enough, not about
            # which way round it was drawn — a genuine orientation problem shows
            # up in the separate diffuser/duct height clearance check instead.
            direct  = _ge_in(duct_w, face_w) and _ge_in(duct_h, face_h)
            flipped = _ge_in(duct_w, face_h) and _ge_in(duct_h, face_w)
            branch_ok = direct or flipped

    # Compose the verdict. A real failure always beats an unverifiable half:
    # not being able to check the diffuser does not excuse an undersized duct.
    if diffuser_ok is False and branch_ok is False:
        res.status = 'RED'
        res.reason = 'Diffuser + branch undersized'
    elif diffuser_ok is False:
        res.status = 'RED'
        res.reason = 'Diffuser undersized for CFM'
    elif branch_ok is False:
        res.status = 'RED'
        res.reason = 'Branch smaller than diffuser'
    elif diffuser_ok is None or branch_ok is None:
        res.status = 'GRAY'
        res.reason = 'Cannot check: ' + '; '.join(unverified)
    else:
        res.status = 'GREEN'
        # A passing branch states WHY it passed rather than leaving Reason
        # blank: a branch is never judged on velocity or friction, so a reader
        # seeing its (real, printed) FPM next to an empty Reason would assume
        # the velocity limit is what cleared it. Colin, 2026-09-24.
        res.reason = 'BRANCH DUCTING SIZED OFF DIFFUSER'

    log.info('branch_diffuser_check: terminal id=%s cat=%s class=%s cfm=%.0f '
             'diffuser=%s installed=%s required=%s -> %s %s',
             tid, res.category, res.sys_class, res.cfm, res.diffuser_size,
             res.installed_size, res.required_size, res.status, res.reason)
    return res


# ── find AHU from any connected element ─────────────────────────────────────
def find_ahu(elem):
    """Return (ahu_element, method_used) or (None, reason_string)."""
    # Direct .MEPSystem property (works on Duct)
    try:
        sys = elem.MEPSystem
        if sys is not None:
            base = sys.BaseEquipment
            if base is not None:
                log.info('find_ahu: found via MEPSystem.BaseEquipment id=%s', eid_int(base.Id))
                return base, 'MEPSystem.BaseEquipment'
    except Exception as ex:
        log.debug('find_ahu MEPSystem attempt failed: %s', ex)

    # Via connectors (terminals, fittings, equipment)
    cm = _connector_manager(elem)
    if cm is not None:
        for conn in cm.Connectors:
            try:
                if conn.Domain != Domain.DomainHvac:
                    continue
                sys = conn.MEPSystem
                if sys is None:
                    continue
                base = sys.BaseEquipment
                if base is not None:
                    log.info('find_ahu: found via connector MEPSystem id=%s', eid_int(base.Id))
                    return base, 'connector.MEPSystem.BaseEquipment'
            except Exception:
                continue

    log.warning('find_ahu: no base equipment found on element id=%s', eid_int(elem.Id))
    return None, 'no base equipment found'


def _peek_branch_class(start_elem, exclude_id, max_hops=2):
    """Classify a branch WITHOUT adding anything to the traversal graph —
    used to decide whether to prune it before ever visiting it. Walks up to
    max_hops connector-hops from start_elem (skipping exclude_id, the node
    we're branching away from) looking for the first duct/terminal with a
    resolvable system classification. Returns None if nothing classified is
    found within the hop limit."""
    seen = set([exclude_id])
    frontier = [start_elem]
    hops = 0
    while frontier and hops <= max_hops:
        nxt = []
        for e in frontier:
            eid = eid_int(e.Id)
            if eid in seen:
                continue
            seen.add(eid)
            cls = elem_sys_class(e)
            if cls is not None:
                return cls
            cm = _connector_manager(e)
            if cm is None:
                continue
            for conn in cm.Connectors:
                try:
                    if conn.Domain != Domain.DomainHvac:
                        continue
                except Exception:
                    continue
                for ref in conn.AllRefs:
                    try:
                        owner = ref.Owner
                        oid   = eid_int(owner.Id)
                        if oid in seen:
                            continue
                        nxt.append(owner)
                    except Exception:
                        continue
        frontier = nxt
        hops += 1
    return None


# ── BFS traversal ────────────────────────────────────────────────────────────
def traverse(root, allowed_ids=None, equipment_level=False):
    """BFS outward through all HVAC connectors from root.

    allowed_ids: optional set of int element IDs.  When provided the BFS will
    only visit nodes whose ID is in this set.  Use this to re-root the tree
    after a first undirected pass without re-traversing the full model.

    equipment_level: when True, any branch found while expanding a
    mechanical equipment node (VAV, FCU, AHU, ...) is pruned — not visited,
    not sized, not reported — if either:
      (a) it classifies as 'Outside Air' (any direction), or
      (b) it classifies as 'Supply Air' AND flows INTO the equipment
          (Connector.Direction == In) — the upstream trunk connection back
          toward an AHU/source.
    Classification is resolved via a short connector-hop peek
    (_peek_branch_class) before the branch is ever added to the graph, so
    pruned branches never touch nodes/children/cfm_map. Return Air is never
    pruned by (b) even though it's also typically an 'In' connector — RA is
    part of equipment-level scope by design (a box's own local return).
    Exhaust Air is unaffected either way.

    This means: rooting directly at one or more pieces of equipment (rather
    than at a shared AHU) with equipment_level=True never walks back
    upstream past that equipment — only that equipment's own Supply
    discharge and Return branches are traversed. When equipment is instead
    reached mid-tree from an AHU root, rule (b) is naturally inert (the
    upstream connector is already in `visited` by the time it's reached, so
    it's never re-evaluated) — equipment_level only changes behavior for
    equipment used AS the root.

    Returns:
        nodes    dict  int_id -> element
        children dict  int_id -> [child_int_ids]  (directed away from root)
        log_lines list of strings for diagnostic output
    """
    nodes    = {}
    children = {}
    visited  = set()
    log_lines = []

    root_id = eid_int(root.Id)
    visited.add(root_id)
    nodes[root_id]    = root
    children[root_id] = []
    log_lines.append('ROOT  id={} cat={}'.format(root_id, _cat_id(root)))

    queue = [root]
    skipped = 0
    pruned_oa = 0
    pruned_upstream = 0

    while queue:
        elem = queue.pop(0)
        eid  = eid_int(elem.Id)
        cm   = _connector_manager(elem)
        if cm is None:
            skipped += 1
            continue
        elem_is_equip = is_equipment(elem)
        for conn in cm.Connectors:
            try:
                if conn.Domain != Domain.DomainHvac:
                    continue
            except Exception:
                continue
            for ref in conn.AllRefs:
                try:
                    owner    = ref.Owner
                    owner_id = eid_int(owner.Id)
                    if owner_id in visited:
                        continue
                    # When re-rooting, stay within the already-known node set
                    if allowed_ids is not None and owner_id not in allowed_ids:
                        continue
                    if equipment_level and elem_is_equip:
                        branch_cls = _peek_branch_class(owner, eid)
                        if branch_cls == 'Outside Air':
                            pruned_oa += 1
                            log_lines.append(
                                '  PRUNED OA branch id={} parent={} '
                                '(equipment-level mode)'.format(owner_id, eid))
                            continue
                        if branch_cls == 'Supply Air':
                            try:
                                conn_dir = conn.Direction
                            except Exception:
                                conn_dir = None
                            if conn_dir == FlowDirectionType.In:
                                pruned_upstream += 1
                                log_lines.append(
                                    '  PRUNED upstream Supply Air branch id={} parent={} '
                                    '(equipment-level mode — never goes upstream)'
                                    .format(owner_id, eid))
                                continue
                    visited.add(owner_id)
                    nodes[owner_id]    = owner
                    children[owner_id] = []
                    children[eid].append(owner_id)
                    queue.append(owner)
                    log_lines.append(
                        '  FOUND id={} cat={} parent={}'.format(owner_id, _cat_id(owner), eid))
                except Exception as ex:
                    log_lines.append('  CONNECTOR ERROR: {}'.format(str(ex)))
                    continue

    if skipped:
        log_lines.append('Skipped {} elements with no ConnectorManager.'.format(skipped))
    if pruned_oa:
        log_lines.append(
            'Pruned {} Outside Air branch(es) at equipment (equipment-level mode).'
            .format(pruned_oa))
    if pruned_upstream:
        log_lines.append(
            'Pruned {} upstream Supply Air branch(es) at equipment (equipment-level mode).'
            .format(pruned_upstream))

    return nodes, children, log_lines


# ── post-order CFM accumulation (iterative) ──────────────────────────────────
def compute_cfm(root_id, nodes, children, terminal_cfms):
    """Iterative post-order DFS. Returns (cfm_map, equip_class_cfm).

    cfm_map: dict int_id -> cfm (every node, same as before).

    equip_class_cfm: dict int_id -> {sys_class: cfm}, populated only for
    OST_MechanicalEquipment nodes (VAV boxes, FCUs, AHUs). At a normal node
    cfm_map[nid] is still the flat sum of all children (unchanged behavior —
    this is what upstream/trunk ductwork continues to use). But AT an
    equipment node, children are also grouped by branch system
    classification (Outside Air vs Supply Air vs Return/Exhaust, via
    _branch_sys_class) so callers can tell "outside air feeding into this
    box" apart from "supply air leaving this box" instead of only ever
    seeing them pre-summed into one number. cfm_map[nid] for an equipment
    node stays the same combined total either way — equip_class_cfm is the
    breakdown layered on top, not a replacement.
    """
    cfm_map = {}
    equip_class_cfm = {}
    stack   = [(root_id, False)]
    while stack:
        nid, done = stack.pop()
        if done:
            elem = nodes.get(nid)
            if nid in terminal_cfms:
                cfm_map[nid] = float(terminal_cfms[nid])
            elif elem is not None and is_equipment(elem):
                by_class = {}
                for cid in children.get(nid, []):
                    cls = _branch_sys_class(cid, nodes, children) or 'Unknown'
                    by_class[cls] = by_class.get(cls, 0.0) + cfm_map.get(cid, 0.0)
                equip_class_cfm[nid] = by_class
                # sum(..., 0.0) forces float — sum([]) returns int 0 in Python 2.7
                cfm_map[nid] = sum(by_class.values(), 0.0)
            else:
                # sum(..., 0.0) forces float — sum([]) returns int 0 in Python 2.7
                cfm_map[nid] = sum((cfm_map.get(c, 0.0) for c in children.get(nid, [])), 0.0)
        else:
            stack.append((nid, True))
            for cid in children.get(nid, []):
                stack.append((cid, False))
    return cfm_map, equip_class_cfm


def compute_downstream_terminal_ids(root_id, nodes, children, terminal_ids):
    """Iterative post-order DFS. Returns {node_id: set of terminal int-ids
    reachable at or below that node}.

    This is what separates a branch from a main: a duct is a BRANCH when
    exactly one terminal is reachable downstream of it (however many segments
    and fittings that tap-off is made of), and a MAIN when two or more are.
    Counting terminals rather than looking at duct size or position is what
    makes the split hold on real geometry — a long multi-segment tap with
    three elbows in it is still one branch, and a short stub that happens to
    split two ways is still a main.

    Iterative rather than recursive for the same reason compute_cfm() is:
    real networks are deep enough to be worth not trusting to Python's
    recursion limit.

    A node's own id is included in its set when it is itself a terminal, so
    the terminal at the end of a branch reports itself and the check stays
    consistent all the way down the run.

    terminal_ids: any container supporting `in`; a set is expected. Nodes not
    reachable from root_id simply do not appear in the result — callers should
    use .get(nid, ...) rather than indexing.
    """
    result = {}
    stack  = [(root_id, False)]
    while stack:
        nid, done = stack.pop()
        if done:
            acc = set()
            if nid in terminal_ids:
                acc.add(nid)
            for cid in children.get(nid, []):
                acc.update(result.get(cid, ()))
            result[nid] = acc
        else:
            stack.append((nid, True))
            for cid in children.get(nid, []):
                stack.append((cid, False))
    return result


# ── solid fill pattern ───────────────────────────────────────────────────────
def solid_fill_pattern_id(doc):
    from Autodesk.Revit.DB import FilteredElementCollector
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except Exception:
            pass
    return ElementId.InvalidElementId


# ── full network build (used by both Diagnose and Duct Velocity) ─────────────
class HvacNetwork(object):
    """Holds everything about one traversal run."""

    def __init__(self):
        self.root            = None      # AHU element or fallback root
        self.ahu_method      = ''        # how AHU was found
        self.nodes           = {}        # int_id -> element
        self.children        = {}        # int_id -> [child_ids]
        self.traverse_log    = []        # raw BFS log lines
        self.terminal_cfms   = {}        # int_id -> cfm
        self.zero_terminals  = []        # int_ids with Flow = 0
        self.missing_flow    = []        # int_ids where Flow param not found
        self.cfm_map         = {}        # int_id -> cfm (all nodes)
        self.equip_class_cfm = {}        # int_id -> {sys_class: cfm} (equipment nodes only)
        self.duct_results    = {}        # ElementId -> DuctResult
        self.no_area_ducts   = []        # int_ids
        self.errors          = []
        self.warnings        = []

    @property
    def terminals(self):
        return [e for e in self.nodes.values() if is_terminal(e)]

    @property
    def ducts(self):
        return [e for e in self.nodes.values() if is_duct(e)]

    @property
    def equipment_nodes(self):
        return [e for e in self.nodes.values() if is_equipment(e)]

    @property
    def accessories(self):
        return [e for e in self.nodes.values() if is_accessory(e)]

    @property
    def fittings(self):
        """Fittings only (OST_DuctFitting) - excludes accessories, which
        Diagnose reports as their own row. is_fitting_or_accessory() covers
        both categories for callers that don't need the split."""
        return [e for e in self.nodes.values() if is_fitting(e)]

    @property
    def ready_for_visualization(self):
        return len(self.errors) == 0 and len(self.ducts) > 0

    @property
    def terminal_count(self):
        return len(self.terminal_cfms)

    @property
    def duct_count(self):
        return len(self.duct_results)

    def equipment_discharge_cfm(self, equip_id):
        """Supply Air CFM leaving this equipment node (its real discharge
        total) — excludes Outside/Return/Exhaust Air branches feeding into
        it. Falls back to the old combined total (cfm_map) if no branch on
        this equipment resolved to a known classification at all (e.g. no
        System Classification data in the model), so unclassified projects
        keep working exactly as before."""
        by_class = self.equip_class_cfm.get(equip_id)
        if not by_class:
            return self.cfm_map.get(equip_id, 0.0)
        if 'Supply Air' in by_class:
            return by_class['Supply Air']
        # No Supply Air branch resolved (e.g. this equipment IS the terminal
        # end, or classification data is missing) — fall back to the combined
        # total rather than silently reporting 0.
        return self.cfm_map.get(equip_id, 0.0)

    def equipment_other_class_cfm(self, equip_id):
        """dict {sys_class: cfm} for every branch on this equipment node
        EXCEPT Supply Air (e.g. {'Outside Air': 100.0}) — the CFM that
        equipment_discharge_cfm() deliberately excludes from its total."""
        by_class = self.equip_class_cfm.get(equip_id, {})
        return dict((k, v) for k, v in by_class.items() if k != 'Supply Air')


class DuctResult(object):
    def __init__(self, elem, cfm, area, sys_class):
        self.elem              = elem
        self.element_id        = eid_int(elem.Id)
        self.cfm               = cfm
        self.area_ft2          = area
        self.sys_class         = sys_class
        self.fpm               = (cfm / area) if area > 0 and cfm > 0 else 0.0
        self.label             = smacna_label(self.fpm, sys_class)
        self.size              = duct_size_label(elem)
        self.d_h_in            = _duct_d_h_in(elem)
        self.friction_per_100ft = duct_friction_loss_per_100ft(self.fpm, self.d_h_in)
        self.length_ft          = _duct_length_ft(elem)
        self.friction_loss_inwc = self.length_ft / 100.0 * self.friction_per_100ft


def build_network(selected_elem, doc, cfm_is_direct=False, equipment_level=False):
    """Full traversal from selection → AHU → network.

    equipment_level: Supply + Return Air only, no Outside Air and no
    upstream travel past equipment used as root — see traverse(). Only
    applied to the final, rooted-at-equipment traversal; the PASS-2
    undirected discovery pass (used only to locate the equipment itself)
    always runs unpruned so equipment-finding isn't affected by it.

    Returns HvacNetwork populated with all traversal data.
    """
    net = HvacNetwork()

    # If the selected element is itself mechanical equipment (fan coil, VRF
    # indoor unit, split system — no central AHU), use it as root directly.
    if is_equipment(selected_elem):
        net.root       = selected_elem
        net.ahu_method = 'selected element is mechanical equipment'
        net.nodes, net.children, net.traverse_log = traverse(
            net.root, equipment_level=equipment_level)
    else:
        # PASS 1: try fast MEPSystem.BaseEquipment lookup
        ahu, method = find_ahu(selected_elem)
        if ahu is not None:
            net.root       = ahu
            net.ahu_method = method
            net.nodes, net.children, net.traverse_log = traverse(
                net.root, equipment_level=equipment_level)
        else:
            # PASS 2: undirected BFS then re-root at any equipment found.
            # Always unpruned — pruning here could hide the very equipment
            # this pass exists to locate.
            all_nodes, all_children, all_log = traverse(selected_elem)
            all_ids = set(all_nodes.keys())

            sel_id = eid_int(selected_elem.Id)
            equip_found = [
                elem for nid, elem in all_nodes.items()
                if is_equipment(elem) and nid != sel_id
            ]

            if equip_found:
                net.root       = equip_found[0]
                net.ahu_method = (
                    'found in traversal: OST_MechanicalEquipment id={}'
                    .format(eid_int(net.root.Id))
                )
                net.nodes, net.children, net.traverse_log = traverse(
                    net.root, allowed_ids=all_ids,
                    equipment_level=equipment_level
                )
                net.traverse_log.insert(0,
                    'NOTE: re-rooted from selection id={} to equipment id={}'
                    .format(sel_id, eid_int(net.root.Id)))
            else:
                # True fallback — no AHU found anywhere in the network
                net.warnings.append(
                    'No base equipment (AHU) found in traversal. '
                    'CFM sums are computed away from the selected element '
                    'and may not reflect actual flow direction.')
                net.root       = selected_elem
                net.ahu_method = 'fallback: selected element used as root'
                net.nodes      = all_nodes
            net.children   = all_children
            net.traverse_log = all_log

    if equipment_level:
        pruned_oa = sum(1 for ln in net.traverse_log if 'PRUNED OA branch' in ln)
        pruned_up = sum(1 for ln in net.traverse_log if 'PRUNED upstream Supply Air branch' in ln)
        if pruned_oa:
            net.warnings.append(
                '{} Outside Air branch(es) not traversed — equipment-level '
                'mode is on. OA ductwork is not sized or colored.'.format(pruned_oa))
        if pruned_up:
            net.warnings.append(
                '{} upstream Supply Air branch(es) not traversed — equipment-level '
                'mode never travels back toward an AHU/source.'.format(pruned_up))

    if len(net.nodes) == 0:
        net.errors.append('No elements found in traversal. Check that the selected element is connected to a duct system.')
        return net

    # Collect terminal CFMs — try multiple param names to handle
    # supply diffusers ("Flow"), return/exhaust grilles ("Airflow", built-in)
    _FLOW_PARAM_NAMES = ['Flow', 'Airflow', 'Air Flow', 'CFM']

    for nid, elem in net.nodes.items():
        if not is_terminal(elem):
            continue

        # Try named parameters first
        fp = None
        for pname in _FLOW_PARAM_NAMES:
            candidate = elem.LookupParameter(pname)
            if candidate is not None and candidate.AsDouble() > 0:
                fp = candidate
                break
        # Fall back to built-in RBS_DUCT_FLOW_PARAM
        if fp is None or fp.AsDouble() <= 0:
            builtin = elem.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM)
            if builtin is not None and builtin.AsDouble() > 0:
                fp = builtin

        if fp is None:
            net.terminal_cfms[nid] = 0.0
            net.missing_flow.append(nid)
        else:
            cfm = to_cfm(fp.AsDouble(), cfm_is_direct)
            net.terminal_cfms[nid] = cfm
            if cfm <= 0:
                net.zero_terminals.append(nid)

    # Post-order CFM sum
    net.cfm_map, net.equip_class_cfm = compute_cfm(
        eid_int(net.root.Id), net.nodes, net.children, net.terminal_cfms)

    # Duct results
    for nid, elem in net.nodes.items():
        if not is_duct(elem):
            continue
        cfm      = net.cfm_map.get(nid, 0.0)
        area     = duct_area_ft2(elem)
        sys_class = duct_sys_class(elem)
        if area <= 0:
            net.no_area_ducts.append(nid)
        dr = DuctResult(elem, cfm, area, sys_class)
        net.duct_results[elem.Id] = dr

    # Errors and warnings
    if len(net.terminal_cfms) == 0:
        net.errors.append(
            'No air terminals (OST_DuctTerminal) found in traversal. '
            'Check that diffusers are physically connected to the duct system.')
    if len(net.zero_terminals) > 0:
        net.warnings.append(
            '{} terminal(s) have Flow = 0. Assign CFM values in the model '
            'before running Duct Velocity.'.format(len(net.zero_terminals)))
    if len(net.missing_flow) > 0:
        net.warnings.append(
            '{} terminal(s) are missing the "Flow" parameter entirely.'.format(
                len(net.missing_flow)))
    if len(net.no_area_ducts) > 0:
        net.warnings.append(
            '{} duct(s) have no dimension data — will show gray.'.format(
                len(net.no_area_ducts)))
    if len(net.ducts) == 0:
        net.errors.append('No duct segments found in traversal.')

    return net
