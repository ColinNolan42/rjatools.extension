# -*- coding: ascii -*-
"""
vav_checks.py  --  VAV mode for Duct Performance and Size Duct

Pure Python, no Revit imports, so every rule here tests in plain CPython
(lib/_test_vav_checks.py). The Revit side (connector directions) lives in
hvac_graph.py and hands this module plain ids, strings and numbers.

Approved by Colin, 2026-10-22 (.claude/docs/Duct Velocity/VAV_PLAN.md). Single
duct and fan powered boxes are treated the same. Dual duct is ignored.

What it decides:
  - which mid tree equipment is a VAV box (entered through its supply inlet),
    and which is a parallel unit (entered through its supply outlet)
  - the medium pressure (MP) set: supply ducts, fittings and accessories
    between the unit and the box inlets. Decided by POSITION in the graph,
    never by System Type name, because MP and LP supply usually share one
    System Type.
  - supply terminals with no box upstream (a diffuser on an MP main)
  - two picked units that share ductwork

Sources (VAV design values come only from these):
  FOR CLAUDE\\Design Resources\\MECHANICAL\\Air Side Systems\\VAV Systems, plus
  the ASHRAE Duct Systems Design Guide (DSDG).

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import logging

log = logging.getLogger(__name__)


# -- Defaults (VAV_PLAN.md, Sources table) -----------------------------------

# Key used for the MP limits inside a limits dict ({system: (fpm, friction)}).
# Not a Revit System Type name on purpose: MP is a position, not a system.
MP_KEY = 'Supply Air (MP)'

MP_DEFAULT_FPM = 2000.0          # EDR VAV Guide p100, ducts in ceiling plenums
MP_DEFAULT_FRICTION = 0.30       # EDR VAV Guide p3, p100 (start at 0.25 to 0.30)
DEFAULT_BOX_DROP_INWC = 0.25     # Colin: priced like a balancing damper for now
DEFAULT_MIN_BOX_PRESSURE_INWC = 0.75   # Colin's firm rule of thumb
DEFAULT_UNIT_MAX_ESP_INWC = None       # blank: no check, engineer enters the unit's value

ROLE_MP = 'Main (MP)'
# An MP duct (or fitting) with exactly ONE box downstream: the MP analog of a
# branch. The engineer sizes it to the box inlet, so it is never judged on the
# MP limits and never resized (Colin, 2026-10-22, review W5).
ROLE_INLET = 'VAV Inlet'
REASON_INLET_OK = 'VAV INLET, SIZED TO BOX'

# Red reasons. Each fits the 34 character Reason column of Duct Performance.
REASON_DIFFUSER_ON_MP = 'Diffuser on MP main'
# The "flex at the VAV inlet" RED rule was removed (Colin, 2026-10-23:
# "nobody does that"), so it is no longer checked.

# Dialog fields: (key, label, unit hint, default, kind). kind is one of
# 'pos' (> 0), 'nonneg' (>= 0) or 'pos_or_blank' (blank allowed, else > 0).
# One list for the themed dialog, the legacy dialog and the tests.
# Limit fields sit in the VAV section; static fields sit under the external
# static option and show only when VAV and external static are both ticked
# (Colin, 2026-10-22).
VAV_LIMIT_FIELDS = (
    ('mp_fpm', 'MP supply max velocity', 'FPM', MP_DEFAULT_FPM, 'pos'),
    ('mp_fric', 'MP supply max friction', 'in. wc/100 ft', MP_DEFAULT_FRICTION, 'pos'),
)
VAV_STATIC_FIELDS = (
    ('box_drop', 'VAV box pressure drop', 'in. wc', DEFAULT_BOX_DROP_INWC, 'nonneg'),
    ('unit_max_esp', 'Unit max external static', 'in. wc, blank for no check',
     DEFAULT_UNIT_MAX_ESP_INWC, 'pos_or_blank'),
    ('min_box_inwc', 'Minimum pressure available at a box', 'in. wc',
     DEFAULT_MIN_BOX_PRESSURE_INWC, 'nonneg'),
)
VAV_FIELDS = VAV_LIMIT_FIELDS + VAV_STATIC_FIELDS


def default_settings():
    """{key: default} for every VAV_FIELDS row."""
    return dict((k, d) for k, _l, _u, d, _kind in VAV_FIELDS)


def format_default(value):
    """Dialog text for a default: blank for None, else the plain number."""
    if value is None:
        return ''
    return '%g' % float(value)


def parse_field(text, kind):
    """(value, error or None) for one dialog box. Blank is None for
    'pos_or_blank'; anything else must be a finite number in range."""
    s = (text or '').strip()
    if s == '' and kind == 'pos_or_blank':
        return None, None
    try:
        v = float(s)
    except (TypeError, ValueError):
        return None, 'must be a number'
    if v != v or v in (float('inf'), float('-inf')):
        return None, 'must be a number'
    if kind in ('pos', 'pos_or_blank') and v <= 0:
        return None, 'must be greater than 0'
    if kind == 'nonneg' and v < 0:
        return None, 'cannot be negative'
    return v, None


def basis_rows(settings=None):
    """(label, text) rows naming the VAV limits and their sources, for the
    dialogs' Calculation basis, the sheet legend and the console."""
    s = default_settings()
    if settings:
        s.update(settings)
    esp = s.get('unit_max_esp')
    return [
        ('VAV, MP supply',
         'Supply mains between the unit and the box inlets with two or more boxes '
         'downstream: {:.0f} fpm and {:.2f} in. wc per 100 ft (EDR VAV Guide p3, '
         'p100). Always mains, never oversized.'.format(
             float(s['mp_fpm']), float(s['mp_fric']))),
        ('VAV, LP',
         'After the box: the Supply Air limits (EDR p99, Taylor 2019 p54).'),
        ('VAV, red',
         'A diffuser on an MP main.'),
        ('VAV, inlet',
         'An MP duct with one box downstream is the box inlet runout, sized to '
         'the box by the engineer: not judged on the MP limits, never resized.'),
        ('VAV, static',
         'Box drop {:.2f} in. wc. Pressure left at each box is the unit max '
         'external static ({}) less the MP loss, flagged under {:.2f} in. wc '
         '(RJA rule).'.format(
             float(s['box_drop']),
             'not entered, no check' if esp is None else '%.2f in. wc' % float(esp),
             float(s['min_box_inwc']))),
    ]


# -- Shared limits lookup ------------------------------------------------------

def limits_for(sys_class, limits, defaults, is_mp=False):
    """(max_fpm, max_friction) for one duct.

    The one lookup both tools use. An MP duct takes the MP limits whatever its
    System Type is called; anything else takes its own system's limits, then
    the firm default for that system, then (600, 0.05).

    limits: {system: (fpm, friction)} from the dialog, may carry MP_KEY.
    defaults: hvac_graph.FIRM_DEFAULTS (passed in so this stays pure).
    """
    limits = limits or {}
    if is_mp:
        return tuple(limits.get(MP_KEY, (MP_DEFAULT_FPM, MP_DEFAULT_FRICTION)))
    fallback = (defaults or {}).get(sys_class, (600, 0.05))
    return tuple(limits.get(sys_class, fallback))


def is_supply_class(sys_class):
    """Supply side? Matches 'Supply Air' and any System Type name with Supply
    in it, the same test Duct Performance uses to tell a picked duct's side."""
    return bool(sys_class) and 'supply' in sys_class.lower()


def same_side(a, b):
    """Two system classes on the same side (supply, or not supply)? An unknown
    (None or empty) class matches anything, so missing data never cuts flow."""
    if not a or not b:
        return True
    return is_supply_class(a) == is_supply_class(b)


def entry_class(nid, parents, class_of):
    """System class the walk entered nid from: the class of the nearest
    classified element above it, or None."""
    seen = set([nid])
    cur = parents.get(nid)
    while cur is not None and cur not in seen:
        seen.add(cur)
        cls = class_of(cur)
        if cls:
            return cls
        cur = parents.get(cur)
    return None


def counted_children(kids, entry_cls, child_class_of):
    """Children of a NON-ROOT piece of equipment whose flow counts toward the
    duct it was entered from: only those on the same side.

    A fan powered box entered from an MP supply duct passes its supply
    diffusers' flow upward, never its own ducted return (code review C1,
    2026-10-22). Not VAV only: any mid tree equipment works this way. With no
    entry class every child counts, as before.
    """
    if not entry_cls:
        return list(kids)
    return [c for c in kids if same_side(entry_cls, child_class_of(c))]


# -- Graph analysis -------------------------------------------------------------

class VavSystem(object):
    """What analyze_system() found for one unit.

    boxes              equipment ids treated as VAV boxes
    unknown_direction  boxes whose inlet connector direction could not be read
                       (treated as boxes, per the plan, and logged)
    reversed           equipment entered through its supply OUTLET: parallel units
    mp_ids             duct / fitting / accessory ids between the unit and the boxes
    inlet_runout_ids   the mp_ids with exactly ONE box downstream (VAV Inlet)
    mp_terminals       supply terminal id -> nearest duct id upstream of it, for a
                       terminal with no box upstream (a diffuser on an MP main)
    box_inlet          box id -> nearest duct id upstream (its inlet runout)
    box_outlets        box id -> [nearest duct ids downstream]
    """

    def __init__(self):
        self.boxes = []
        self.unknown_direction = []
        self.reversed = []
        self.mp_ids = set()
        self.inlet_runout_ids = set()
        self.mp_terminals = {}
        self.box_inlet = {}
        self.box_outlets = {}

    @property
    def is_vav(self):
        return bool(self.boxes)


def parents_of(children):
    """{child id: parent id} from a children map."""
    out = {}
    for pid, kids in children.items():
        for k in kids:
            out[k] = pid
    return out


def nearest_duct_up(nid, parents, kind_of):
    """Nearest duct strictly above nid, or None."""
    seen = set([nid])
    cur = parents.get(nid)
    while cur is not None and cur not in seen:
        seen.add(cur)
        if kind_of(cur) == 'duct':
            return cur
        cur = parents.get(cur)
    return None


def nearest_ducts_down(nid, children, kind_of):
    """Nearest ducts strictly below nid, walking through fittings and
    accessories only, in child order."""
    out = []
    seen = set([nid])
    frontier = list(children.get(nid, []))
    while frontier:
        nxt = []
        for c in frontier:
            if c in seen:
                continue
            seen.add(c)
            k = kind_of(c)
            if k == 'duct':
                out.append(c)
            elif k in ('fitting', 'accessory'):
                nxt.extend(children.get(c, []))
        frontier = nxt
    return out


def analyze_system(root_id, children, kind_of, class_of, entry_dir_of):
    """Find the boxes and the MP set of one unit's network.

    Args:
        root_id:  the unit (or picked duct) the network is rooted at.
        children: {id: [child ids]}, directed away from the root.
        kind_of(id):  'duct', 'fitting', 'accessory', 'terminal', 'equipment'
                      or None.
        class_of(id): the element's own system class, or None when it has none
                      (fittings, accessories, equipment). A node with none
                      inherits its parent's.
        entry_dir_of(id): for equipment, the direction of the connector the
                      walk entered it through: 'in', 'out' or None (unreadable
                      or bidirectional).

    Returns VavSystem.
    """
    vs = VavSystem()
    parents = parents_of(children)

    side = {root_id: class_of(root_id)}
    under_box = {root_id: False}
    order = [root_id]
    queue = [root_id]
    seen = set([root_id])
    while queue:
        nid = queue.pop(0)
        nid_is_box = False
        if nid != root_id and kind_of(nid) == 'equipment' and \
                is_supply_class(side.get(nid)) and not under_box.get(nid):
            d = entry_dir_of(nid)
            if d == 'out':
                vs.reversed.append(nid)
                log.warning('vav_checks: equipment id=%s entered through its supply '
                            'outlet: parallel units on one main', nid)
            else:
                vs.boxes.append(nid)
                nid_is_box = True
                if d is None:
                    vs.unknown_direction.append(nid)
                    log.warning('vav_checks: equipment id=%s inlet direction not '
                                'readable, treated as a VAV box', nid)
        for c in children.get(nid, []):
            if c in seen:
                continue
            seen.add(c)
            own = class_of(c)
            side[c] = own if own else side.get(nid)
            under_box[c] = under_box.get(nid, False) or nid_is_box
            order.append(c)
            queue.append(c)

    if not vs.boxes:
        return vs

    box_set = set(vs.boxes)
    boxes_below = {}
    for box in vs.boxes:
        vs.box_inlet[box] = nearest_duct_up(box, parents, kind_of)
        vs.box_outlets[box] = nearest_ducts_down(box, children, kind_of)
        cur = parents.get(box)
        walked = set([box])
        while cur is not None and cur not in walked and cur not in box_set:
            walked.add(cur)
            if kind_of(cur) in ('duct', 'fitting', 'accessory') and \
                    is_supply_class(side.get(cur)):
                vs.mp_ids.add(cur)
                boxes_below[cur] = boxes_below.get(cur, 0) + 1
            cur = parents.get(cur)
    vs.inlet_runout_ids = set(n for n in vs.mp_ids if boxes_below.get(n) == 1)

    for nid in order:
        if kind_of(nid) == 'terminal' and is_supply_class(side.get(nid)) and \
                not under_box.get(nid):
            vs.mp_terminals[nid] = nearest_duct_up(nid, parents, kind_of)

    log.info('vav_checks: root=%s boxes=%d mp=%d reversed=%d mp_terminals=%d',
             root_id, len(vs.boxes), len(vs.mp_ids), len(vs.reversed),
             len(vs.mp_terminals))
    return vs


def overlapping_roots(nodes_by_root):
    """Roots whose networks share any element: parallel units on one main.

    nodes_by_root: {root id: set of node ids}. A root that sits INSIDE another
    root's network (a VAV box picked alongside its AHU) is nested, not
    parallel, and is not reported.
    """
    out = set()
    roots = list(nodes_by_root.keys())
    for i, a in enumerate(roots):
        for b in roots[i + 1:]:
            na, nb = nodes_by_root[a], nodes_by_root[b]
            if a in nb or b in na:
                continue
            if na & nb:
                out.add(a)
                out.add(b)
    return out


def red_reasons(systems):
    """{duct id: RED reason} for the VAV rules, over every unit's VavSystem:
    a diffuser fed straight off an MP main. The caller applies these after the
    normal label, so they override it.

    Args:
        systems: iterable of VavSystem.
    """
    out = {}
    for vs in systems:
        for _tid, feed in sorted(vs.mp_terminals.items()):
            if feed is not None:
                out.setdefault(feed, REASON_DIFFUSER_ON_MP)
    return out


# -- External static at the boxes ---------------------------------------------

def pressure_left(unit_max_esp, mp_loss_inwc):
    """Unit max ESP less the MP loss to the box, or None when no unit max."""
    if unit_max_esp is None:
        return None
    return float(unit_max_esp) - float(mp_loss_inwc)


def pressure_low(left_inwc, minimum_inwc):
    """True when the pressure left at a box is under the minimum."""
    if left_inwc is None:
        return False
    return left_inwc < float(minimum_inwc) - 1e-9
