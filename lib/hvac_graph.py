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
    FillPatternElement, ElementId, Domain
)

log = logging.getLogger(__name__)

# ── Firm design defaults ─────────────────────────────────────────────────────
# {sys_class: (max_fpm, max_friction_inwc_per_100ft)}
# Main and branch ducts share the same values — no split.
FIRM_DEFAULTS = {
    'Supply Air':  (800,  0.08),
    'Return Air':  (600,  0.05),
    'Exhaust Air': (600,  0.05),
    'Outside Air': (600,  0.05),
}

# SMACNA labels for Diagnose report — derived from FIRM_DEFAULTS with 85% green band
# (green_fpm = max*0.85, yellow_fpm = max)
SMACNA = {
    'Supply Air':  (680, 800),    # 800 * 0.85 = 680
    'Return Air':  (510, 600),    # 600 * 0.85 = 510
    'Exhaust Air': (510, 600),
    'Outside Air': (510, 600),
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
    """
    if cfm_is_direct:
        return raw
    try:
        from Autodesk.Revit.DB import UnitTypeId, UnitUtils
        return UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.CubicFeetPerMinute)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertFromInternalUnits(raw, DisplayUnitType.DUT_CUBIC_FEET_PER_MINUTE)
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
        return elem.Category.Id.IntegerValue
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


# ── find AHU from any connected element ─────────────────────────────────────
def find_ahu(elem):
    """Return (ahu_element, method_used) or (None, reason_string)."""
    # Direct .MEPSystem property (works on Duct)
    try:
        sys = elem.MEPSystem
        if sys is not None:
            base = sys.BaseEquipment
            if base is not None:
                log.info('find_ahu: found via MEPSystem.BaseEquipment id=%s', base.Id.IntegerValue)
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
                    log.info('find_ahu: found via connector MEPSystem id=%s', base.Id.IntegerValue)
                    return base, 'connector.MEPSystem.BaseEquipment'
            except Exception:
                continue

    log.warning('find_ahu: no base equipment found on element id=%s', elem.Id.IntegerValue)
    return None, 'no base equipment found'


# ── BFS traversal ────────────────────────────────────────────────────────────
def traverse(root, allowed_ids=None):
    """BFS outward through all HVAC connectors from root.

    allowed_ids: optional set of int element IDs.  When provided the BFS will
    only visit nodes whose ID is in this set.  Use this to re-root the tree
    after a first undirected pass without re-traversing the full model.

    Returns:
        nodes    dict  int_id -> element
        children dict  int_id -> [child_int_ids]  (directed away from root)
        log_lines list of strings for diagnostic output
    """
    nodes    = {}
    children = {}
    visited  = set()
    log_lines = []

    root_id = root.Id.IntegerValue
    visited.add(root_id)
    nodes[root_id]    = root
    children[root_id] = []
    log_lines.append('ROOT  id={} cat={}'.format(root_id, _cat_id(root)))

    queue = [root]
    skipped = 0

    while queue:
        elem = queue.pop(0)
        eid  = elem.Id.IntegerValue
        cm   = _connector_manager(elem)
        if cm is None:
            skipped += 1
            continue
        for conn in cm.Connectors:
            try:
                if conn.Domain != Domain.DomainHvac:
                    continue
            except Exception:
                continue
            for ref in conn.AllRefs:
                try:
                    owner    = ref.Owner
                    owner_id = owner.Id.IntegerValue
                    if owner_id in visited:
                        continue
                    # When re-rooting, stay within the already-known node set
                    if allowed_ids is not None and owner_id not in allowed_ids:
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

    return nodes, children, log_lines


# ── post-order CFM accumulation (iterative) ──────────────────────────────────
def compute_cfm(root_id, nodes, children, terminal_cfms):
    """Iterative post-order DFS. Returns dict int_id -> cfm."""
    cfm_map = {}
    stack   = [(root_id, False)]
    while stack:
        nid, done = stack.pop()
        if done:
            if nid in terminal_cfms:
                cfm_map[nid] = float(terminal_cfms[nid])
            else:
                # sum(..., 0.0) forces float — sum([]) returns int 0 in Python 2.7
                cfm_map[nid] = sum((cfm_map.get(c, 0.0) for c in children.get(nid, [])), 0.0)
        else:
            stack.append((nid, True))
            for cid in children.get(nid, []):
                stack.append((cid, False))
    return cfm_map


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
    def ready_for_visualization(self):
        return len(self.errors) == 0 and len(self.ducts) > 0

    @property
    def terminal_count(self):
        return len(self.terminal_cfms)

    @property
    def duct_count(self):
        return len(self.duct_results)


class DuctResult(object):
    def __init__(self, elem, cfm, area, sys_class):
        self.elem              = elem
        self.element_id        = elem.Id.IntegerValue
        self.cfm               = cfm
        self.area_ft2          = area
        self.sys_class         = sys_class
        self.fpm               = (cfm / area) if area > 0 and cfm > 0 else 0.0
        self.label             = smacna_label(self.fpm, sys_class)
        self.size              = duct_size_label(elem)
        self.d_h_in            = _duct_d_h_in(elem)
        self.friction_per_100ft = duct_friction_loss_per_100ft(self.fpm, self.d_h_in)


def build_network(selected_elem, doc, cfm_is_direct=False):
    """Full traversal from selection → AHU → network.
    Returns HvacNetwork populated with all traversal data.
    """
    net = HvacNetwork()

    # If the selected element is itself mechanical equipment (fan coil, VRF
    # indoor unit, split system — no central AHU), use it as root directly.
    if is_equipment(selected_elem):
        net.root       = selected_elem
        net.ahu_method = 'selected element is mechanical equipment'
        net.nodes, net.children, net.traverse_log = traverse(net.root)
    else:
        # PASS 1: try fast MEPSystem.BaseEquipment lookup
        ahu, method = find_ahu(selected_elem)
        if ahu is not None:
            net.root       = ahu
            net.ahu_method = method
            net.nodes, net.children, net.traverse_log = traverse(net.root)
        else:
            # PASS 2: undirected BFS then re-root at any equipment found
            all_nodes, all_children, all_log = traverse(selected_elem)
            all_ids = set(all_nodes.keys())

            sel_id = selected_elem.Id.IntegerValue
            equip_found = [
                elem for nid, elem in all_nodes.items()
                if is_equipment(elem) and nid != sel_id
            ]

            if equip_found:
                net.root       = equip_found[0]
                net.ahu_method = (
                    'found in traversal: OST_MechanicalEquipment id={}'
                    .format(net.root.Id.IntegerValue)
                )
                net.nodes, net.children, net.traverse_log = traverse(
                    net.root, allowed_ids=all_ids
                )
                net.traverse_log.insert(0,
                    'NOTE: re-rooted from selection id={} to equipment id={}'
                    .format(sel_id, net.root.Id.IntegerValue))
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
    net.cfm_map = compute_cfm(net.root.Id.IntegerValue, net.nodes, net.children, net.terminal_cfms)

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
