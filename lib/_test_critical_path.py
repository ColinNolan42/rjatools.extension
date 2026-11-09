# -*- coding: utf-8 -*-
"""CPython test harness for _critical_path_loss() in Duct Velocity's script.py.

The target function lives in a pyRevit pushbutton script that imports
Autodesk.Revit.DB (and pyrevit) at module scope, so it cannot be imported
normally under plain CPython. Technique: pull the function's source text out
of the file with a regex and exec() it into a namespace where 'hvac_graph'
and 'fitting_tables' are pre-bound. 'fitting_tables' has zero Revit imports,
so the REAL module is imported and used unmodified (its constants are also
used to compute several of the "by hand" expected numbers below, so a
production typo in fitting_tables.py would be caught here too). Only
'hvac_graph' is faked, since it's the module full of Revit connector/element
lookups.

Read-only against the source: this file never writes to the extension.
"""
import io
import logging
import os
import re
import sys
import time
import types

# Derived from this file's own location so the suite travels with the repo
# and runs on any machine: lib/_test_critical_path.py -> extension root.
EXT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT_PATH = os.path.join(
    EXT_ROOT, "Developer Tools.tab", "Performance.panel",
    "Duct Performance.pushbutton", "script.py")
LIB_DIR = os.path.join(EXT_ROOT, "lib")

# fitting_tables.py has no Revit imports -> import the real thing.
sys.path.insert(0, LIB_DIR)
import fitting_tables  # noqa: E402  (real module, pure Python)
import vav_checks      # noqa: E402  (real module, pure Python)
import ashrae_fittings # noqa: E402  (real module, pure Python, reads the JSON tables)
import fan_system_effect  # noqa: E402  (real module, pure Python)
import duct_perf_console  # noqa: E402  (real module, pure Python)
import math            # noqa: E402

src = io.open(SCRIPT_PATH, encoding="utf-8").read()
m = re.search(r"^def _critical_path_loss\(.*?\n    return result$", src, re.S | re.M)
if not m:
    raise SystemExit("Could not find _critical_path_loss() in script.py - "
                      "the function signature/body shape probably changed.")
fn_src = m.group(0)


# ── fake hvac_graph ──────────────────────────────────────────────────────────
#
# Only the 7 hvac_graph members _critical_path_loss() actually calls are
# implemented, against simple fake "node" objects we control directly.

class Elem(object):
    """Stand-in for whatever hvac_graph hands back for a fitting/duct node.

    kind: 'duct' | 'tap' | 'elbow' | 'unknown_fitting' | 'transition'
    'duct' nodes are only ever looked at through takeoff_child_ids /
    duct_continues_past (the real code finds a DuctResult first and never
    calls is_fitting() on a duct node) -- they're included here anyway so
    those two lookups can tell a continuing main from a dead-end tap.
    """
    def __init__(self, kind, family_name=None, is_round=True, down_area=None,
                 angle=None, geom=None, tgeom=None, tap_pos=None):
        self.kind = kind
        self.geom = geom          # elbow_geometry() answer, None = unreadable
        self.tgeom = tgeom        # (ends, theta) for transition_geometry()
        self.tap_pos = tap_pos or {}   # duct only: {tap id: feet}
        self.family_name = family_name
        self.is_round = is_round
        self.down_area = down_area
        self.angle = angle


_FITTING_KINDS = frozenset(['tap', 'elbow', 'unknown_fitting', 'transition'])
# Accessories are a separate Revit category and a separate code path: their drop
# is a user-entered cutsheet number, not a C value off duct geometry.
_ACCESSORY_KINDS = frozenset(['accessory'])


def _takeoff_child_ids(nid, nodes, children):
    """Taps hanging off duct nid: children of nid whose node kind == 'tap'."""
    return [cid for cid in children.get(nid, [])
            if nodes.get(cid) is not None and nodes[cid].kind == 'tap']


def _duct_continues_past(nid, nodes, children):
    """True if nid has a child that is itself another main duct."""
    return any(nodes.get(cid) is not None and nodes[cid].kind == 'duct'
               for cid in children.get(nid, []))


def _duct_is_capped(nid, nodes, children):
    """True if nid has a child end cap (node kind 'cap')."""
    return any(nodes.get(cid) is not None and nodes[cid].kind == 'cap'
               for cid in children.get(nid, []))


def _is_fitting(elem):
    return elem is not None and elem.kind in _FITTING_KINDS


def _is_accessory(elem):
    return elem is not None and elem.kind in _ACCESSORY_KINDS


def _is_equipment(elem):
    """Mid tree equipment (a VAV box, a duct heater): kind 'equipment'."""
    return elem is not None and elem.kind == 'equipment'


def _fitting_family_name(elem):
    return elem.family_name


def _fitting_is_round(elem):
    return elem.is_round


def _fitting_angle_deg(elem):
    return elem.angle


def _elbow_geometry(elem):
    return elem.geom


def _elbow_turn_normal(elem):
    """The swirl cases set Elem.normal (a unit (x, y, z)); None = unreadable."""
    return getattr(elem, 'normal', None) if elem is not None else None


def _transition_geometry(elem):
    return elem.tgeom if elem.tgeom else (None, None)


def _connector_areas_ft2(elem):
    return []      # tap openings unknown: the single branch duct's area is used


def _tap_positions_ft(duct, from_id):
    duct.tap_from = from_id        # the end the taps were measured from
    return dict(duct.tap_pos)


def _transition_areas(elem, up_area):
    return (up_area, elem.down_area)


def _velocity_pressure_inwg(v_fpm):
    return (v_fpm / 4007.7) ** 2


def _duct_friction_loss_per_100ft(v_fpm, d_h_in, eps_ft=None):
    """Flat fake: 0.05 in. wc per 100 ft whatever the duct, so the unit-side
    friction in case 12 is a number that can be checked by hand."""
    return 0.05


def _equipment_connection_geometry(equip, other_id):
    """A unit's connector toward other_id: Elem.conns {child id: geometry},
    set on the equipment Elem by the fan system effect cases. None = unreadable."""
    return (getattr(equip, 'conns', None) or {}).get(other_id)


hvac_graph = types.ModuleType('hvac_graph')
hvac_graph.log = logging.getLogger('hvac_graph_fake')
hvac_graph.equipment_connection_geometry = _equipment_connection_geometry
hvac_graph.duct_friction_loss_per_100ft = _duct_friction_loss_per_100ft
hvac_graph.takeoff_child_ids       = _takeoff_child_ids
hvac_graph.duct_continues_past     = _duct_continues_past
hvac_graph.duct_is_capped          = _duct_is_capped
hvac_graph.is_fitting              = _is_fitting
hvac_graph.is_accessory            = _is_accessory
hvac_graph.is_equipment            = _is_equipment
hvac_graph.fitting_family_name     = _fitting_family_name
hvac_graph.fitting_is_round        = _fitting_is_round
hvac_graph.fitting_angle_deg       = _fitting_angle_deg
hvac_graph.elbow_geometry          = _elbow_geometry
hvac_graph.elbow_turn_normal       = _elbow_turn_normal
hvac_graph.transition_geometry     = _transition_geometry
hvac_graph.tap_positions_ft        = _tap_positions_ft
hvac_graph.connector_areas_ft2     = _connector_areas_ft2
hvac_graph.transition_areas        = _transition_areas
hvac_graph.velocity_pressure_inwg  = _velocity_pressure_inwg
hvac_graph.CAP_LABEL              = 'Wall / roof cap'

ns = {'hvac_graph': hvac_graph, 'fitting_tables': fitting_tables,
      'vav_checks': vav_checks, 'ashrae_fittings': ashrae_fittings,
      'fan_system_effect': fan_system_effect,
      're': re, 'math': math}
exec(fn_src, ns)
crit = ns['_critical_path_loss']


# ── fake duct results / element ids ──────────────────────────────────────────

class DuctResult(object):
    def __init__(self, element_id, friction_loss_inwc, length_ft, fpm,
                 area_ft2, sys_class):
        self.element_id         = element_id
        self.friction_loss_inwc = friction_loss_inwc
        self.length_ft          = length_ft
        self.fpm                = fpm
        self.area_ft2           = area_ft2
        self.sys_class          = sys_class
        self.cfm                = fpm * area_ft2   # as the real DuctResult


class EID(object):
    """Stand-in for a Revit ElementId: deliberately NOT equal to the int key,
    which is the bug this walker re-keys against (dr_by_int_id)."""
    def __init__(self, i):
        self.i = i

    def __hash__(self):
        return hash(("eid", self.i))

    def __eq__(self, other):
        return isinstance(other, EID) and other.i == self.i

    def __ne__(self, other):
        return not self.__eq__(other)


def build(rows):
    """rows: list of (int_id, friction_inwc, length_ft, fpm, area_ft2, sys_class)."""
    return dict((EID(r[0]), DuctResult(*r)) for r in rows)


def pv(fpm):
    """Independent restatement of the documented formula Pv = (V/4007.7)^2,
    used to hand-compute expected fitting losses below (not a call into the
    code under test)."""
    return (fpm / 4007.7) ** 2


results = []  # (case_name, passed, detail)


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("  -- " + detail if detail and not cond else ""))


# ══════════════════════════════════════════════════════════════════════════
# Case 1: two branches off one trunk -- the longer/lossier one must win.
#         Verify it's the MAX path, not the SUM of both branches.
# ══════════════════════════════════════════════════════════════════════════
#   root(0) -> 1 -> 2 -> T10           (0.10 + 0.20        = 0.30)
#                \-> 3 -> 4 -> T11     (0.10 + 0.05 + 0.40 = 0.55)  <- index run
children1 = {0: [1], 1: [2, 3], 2: [10], 3: [4], 4: [11]}
ducts1 = build([
    (1, 0.10, 20.0, 900.0, 1.0, 'Supply Air'),
    (2, 0.20, 30.0, 900.0, 1.0, 'Supply Air'),
    (3, 0.05, 10.0, 900.0, 1.0, 'Supply Air'),
    (4, 0.40, 60.0, 900.0, 1.0, 'Supply Air'),
])
terms1 = {10: (400.0, 'Supply Air', 'SD-A'), 11: (250.0, 'Supply Air', 'SD-B')}
r1 = crit([0], children1, ducts1, terms1, {})
sa1 = r1.get(0, {}).get('Supply Air', {})
# hand calc: long branch friction = 0.10+0.05+0.40 = 0.55, length = 20+10+60 = 90
# (a wrong SUM-of-both-branches implementation would give 0.30+0.55 = 0.85)
check("case1 friction is the MAX branch (0.55), not a sum",
      abs(sa1.get('friction_inwc', -1) - 0.55) < 1e-9,
      "got %r" % sa1.get('friction_inwc'))
check("case1 picks the long branch's terminal (SD-B)",
      sa1.get('terminal_name') == 'SD-B', "got %r" % sa1.get('terminal_name'))
check("case1 duct_count is 3 (ducts 1,3,4)",
      sa1.get('duct_count') == 3, "got %r" % sa1.get('duct_count'))
check("case1 length_ft is 90.0 (20+10+60)",
      abs(sa1.get('length_ft', -1) - 90.0) < 1e-9, "got %r" % sa1.get('length_ft'))
_item_sum = sum(i[3] or 0.0 for i in sa1.get('items', ()))
check("case1 index run items add up to the subtotal",
      abs(_item_sum - sa1.get('subtotal_inwc', -1)) < 1e-9,
      "items %r vs subtotal %r" % (_item_sum, sa1.get('subtotal_inwc')))


# ══════════════════════════════════════════════════════════════════════════
# Case 2: Supply and Return as separate roots -> each root gets its own
#         top-level entry, and within it, its own system-class entry.
# ══════════════════════════════════════════════════════════════════════════
children2 = {0: [1], 1: [10], 100: [2], 2: [20]}
ducts2 = build([
    (1, 0.31, 50.0, 800.0, 1.0, 'Supply Air'),
    (2, 0.12, 40.0, 600.0, 1.0, 'Return Air'),
])
terms2 = {10: (400.0, 'Supply Air', 'SD-1'), 20: (400.0, 'Return Air', 'RG-1')}
r2 = crit([0, 100], children2, ducts2, terms2, {})
check("case2 both roots present as separate top-level keys",
      set(r2.keys()) == {0, 100}, "got %r" % (list(r2.keys()),))
check("case2 root 0 has only the Supply Air system class",
      set(r2.get(0, {}).keys()) == {'Supply Air'}, "got %r" % (list(r2.get(0, {}).keys()),))
check("case2 root 100 has only the Return Air system class",
      set(r2.get(100, {}).keys()) == {'Return Air'}, "got %r" % (list(r2.get(100, {}).keys()),))
check("case2 Supply Air friction == 0.31",
      abs(r2.get(0, {}).get('Supply Air', {}).get('friction_inwc', -1) - 0.31) < 1e-9)
check("case2 Return Air friction == 0.12",
      abs(r2.get(100, {}).get('Return Air', {}).get('friction_inwc', -1) - 0.12) < 1e-9)


# ══════════════════════════════════════════════════════════════════════════
# Case 3: a main with 3 taps, path leaves through ONE of them -> the
#         "main past take-off" coefficient (0.20, C_SUPPLY_TAP_MAIN) must be
#         charged for exactly 2 bypassed taps, not 3 and not 0.
# ══════════════════════════════════════════════════════════════════════════
#   main duct 1 (fric 0, fpm 1200) has children: taps 21, 22, 23, and a
#   continuing main duct 2 (so duct_continues_past(1)=True and the
#   "end of main" swing does NOT also fire -- isolates the bypass count).
#   The traced path exits through tap 22 -> terminal 30.
children3 = {1: [21, 22, 23, 2], 22: [30]}
nodes3 = {
    2:  Elem(kind='duct'),
    21: Elem(kind='tap', family_name='Rectangular Takeoff Single', is_round=False),
    22: Elem(kind='tap', family_name='Rectangular Takeoff Single', is_round=False),
    23: Elem(kind='tap', family_name='Rectangular Takeoff Single', is_round=False),
}
ducts3 = build([
    (1, 0.0, 5.0, 1200.0, 1.0, 'Supply Air'),
    (2, 0.0, 1.0, 500.0, 0.5, 'Supply Air'),   # dead-end continuation, no terminal
])
terms3 = {30: (300.0, 'Supply Air', 'SD-3')}
r3 = crit([1], children3, ducts3, terms3, nodes3)
sa3 = r3.get(1, {}).get('Supply Air', {})
# hand calc:
#   Pv(1200) = (1200/4007.7)^2
#   bypass:   n_by = len(taps)-1 = 2 taps bypassed -> 2 * C_SUPPLY_TAP_MAIN * Pv(1200)
#   take-off: the exit tap itself (not end-of-main, since duct 2 makes the
#             main continue) -> 1 * C_SUPPLY_TAP_BRANCH * Pv(1200)
expected_fit3 = (2 * fitting_tables.DEFAULT_C['supply_tap_main'] * pv(1200.0)
                 + fitting_tables.DEFAULT_C['supply_tap_branch'] * pv(1200.0))
check("case3 friction is 0 (only fitting losses in this network)",
      abs(sa3.get('friction_inwc', -1) - 0.0) < 1e-9)
check("case3 fitting_inwc == 2*0.20*Pv(1200) + 0.98*Pv(1200)",
      abs(sa3.get('fitting_inwc', -1) - expected_fit3) < 1e-9,
      "expected %.6f got %r" % (expected_fit3, sa3.get('fitting_inwc')))
check("case3 tap_bypass_count == 2 (not 3, not 0)",
      sa3.get('tap_bypass_count') == 2, "got %r" % sa3.get('tap_bypass_count'))
check("case3 fitting_count == 1 (only the exit tap itself)",
      sa3.get('fitting_count') == 1, "got %r" % sa3.get('fitting_count'))


# ══════════════════════════════════════════════════════════════════════════
# Case 4: an elbow on the path adds C*Pv using the UPSTREAM duct's velocity,
#         not the downstream duct's.
# ══════════════════════════════════════════════════════════════════════════
#   duct 100 (fpm 1000) -> elbow 101 (round) -> duct 102 (fpm 1600) -> T110
children4 = {100: [101], 101: [102], 102: [110]}
nodes4 = {101: Elem(kind='elbow', family_name='Round Elbow - Fixed L1', is_round=True)}
ducts4 = build([
    (100, 0.0,  10.0, 1000.0, 1.0, 'Supply Air'),
    (102, 0.05, 5.0,  1600.0, 0.6, 'Supply Air'),
])
terms4 = {110: (300.0, 'Supply Air', 'SD-4')}
r4 = crit([100], children4, ducts4, terms4, nodes4)
sa4 = r4.get(100, {}).get('Supply Air', {})
# hand calc: elbow priced at UPSTREAM fpm (1000, from duct 100), not the
# downstream duct 102's 1600 fpm.
expected_fit4 = fitting_tables.DEFAULT_C['round_elbow_90'] * pv(1000.0)
wrong_fit4    = fitting_tables.DEFAULT_C['round_elbow_90'] * pv(1600.0)  # if it used downstream fpm instead
check("case4 friction is 0.05 (only duct 102 has friction)",
      abs(sa4.get('friction_inwc', -1) - 0.05) < 1e-9)
check("case4 elbow fitting_inwc uses UPSTREAM fpm (1000), not downstream (1600)",
      abs(sa4.get('fitting_inwc', -1) - expected_fit4) < 1e-9
      and abs(sa4.get('fitting_inwc', -1) - wrong_fit4) > 1e-6,
      "expected(upstream) %.6f  wrong(downstream) %.6f  got %r"
      % (expected_fit4, wrong_fit4, sa4.get('fitting_inwc')))
check("case4 fitting_count == 1", sa4.get('fitting_count') == 1)


# ══════════════════════════════════════════════════════════════════════════
# Case 5: safety_pct=10 must scale total_inwc to 1.10x subtotal_inwc, and
#         subtotal_inwc must equal friction_inwc + fitting_inwc.
# ══════════════════════════════════════════════════════════════════════════
r5 = crit([100], children4, ducts4, terms4, nodes4, safety_pct=10.0)
sa5 = r5.get(100, {}).get('Supply Air', {})
check("case5 subtotal_inwc == friction + fitting + components",
      abs(sa5['subtotal_inwc'] - (sa5['friction_inwc'] + sa5['fitting_inwc']
                                  + sa5['component_inwc'])) < 1e-12,
      "subtotal=%r friction=%r fitting=%r components=%r"
      % (sa5['subtotal_inwc'], sa5['friction_inwc'], sa5['fitting_inwc'],
         sa5['component_inwc']))
check("case5 omitting comp_values falls back to the published givens, not zero",
      sa5['component_inwc'] > 0.0,
      "component_inwc=%r (DEFAULT_COMPONENTS=%r)"
      % (sa5['component_inwc'], fitting_tables.DEFAULT_COMPONENTS))
check("case5 total_inwc == subtotal_inwc * 1.10",
      abs(sa5['total_inwc'] - sa5['subtotal_inwc'] * 1.10) < 1e-12,
      "total=%r subtotal=%r" % (sa5['total_inwc'], sa5['subtotal_inwc']))


# ══════════════════════════════════════════════════════════════════════════
# Case 6: a cycle in `children` must not hang the walker.
# ══════════════════════════════════════════════════════════════════════════
children6 = {200: [201], 201: [202], 202: [201, 210]}  # 202 -> 201 is a back-edge
ducts6 = build([
    (201, 0.10, 10.0, 800.0, 1.0, 'Supply Air'),
    (202, 0.10, 10.0, 800.0, 1.0, 'Supply Air'),
])
terms6 = {210: (100.0, 'Supply Air', 'SD-X')}
t0 = time.time()
r6 = crit([200], children6, ducts6, terms6, {})
elapsed6 = time.time() - t0
check("case6 cycle terminates quickly (did not hang)",
      elapsed6 < 5.0, "elapsed %.3fs" % elapsed6)
check("case6 cycle still reaches the terminal",
      'Supply Air' in r6.get(200, {})
      and r6[200]['Supply Air'].get('terminal_name') == 'SD-X')


# ══════════════════════════════════════════════════════════════════════════
# Case 7: an unrecognised fitting family name lands in `unpriced`, not a
#         silent 0.
# ══════════════════════════════════════════════════════════════════════════
children7 = {300: [301], 301: [310]}
nodes7 = {301: Elem(kind='unknown_fitting', family_name='Zorbo Flangeless Widget',
                     is_round=True)}
ducts7 = build([(300, 0.0, 5.0, 900.0, 1.0, 'Supply Air')])
terms7 = {310: (100.0, 'Supply Air', 'SD-7')}
r7 = crit([300], children7, ducts7, terms7, nodes7)
sa7 = r7.get(300, {}).get('Supply Air', {})
unpriced7 = sa7.get('unpriced', ())
check("case7 unrecognised fitting costs 0 fitting_inwc",
      abs(sa7.get('fitting_inwc', -1) - 0.0) < 1e-9)
check("case7 unrecognised fitting is recorded in `unpriced`, not silently dropped",
      len(unpriced7) == 1
      and 'Zorbo Flangeless Widget' in unpriced7[0]
      and 'UNPRICED' in unpriced7[0],
      "got %r" % (unpriced7,))


# ══════════════════════════════════════════════════════════════════════════
# Case 8: a node reachable by two different-length paths -- does the shared
#         `visited` set cause the walker to MISS the worse (higher-loss) path?
# ══════════════════════════════════════════════════════════════════════════
#          /--- 402 (fric 0.50, expensive) ---\
#   400 --<                                     >-- 410 (fric 0.02) -- 420 (T)
#          \--- 401 (fric 0.05, cheap)     ---/
#
# children8[400] lists the EXPENSIVE branch (402) first and the CHEAP branch
# (401) second. The walker uses a Python list as a stack (append/pop from the
# end), so children are explored in REVERSE of list order -> the cheap
# branch (401) is popped and fully walked FIRST, reaching the shared duct
# 410 and marking it visited, then reaching terminal 420 and recording
# best = cheap path (0.05 + 0.02 = 0.07).
# The expensive branch (402) is then popped, adds its own 0.50, pushes 410
# as ITS child too -- but 410 is already in `visited`, so `if nid in
# visited: continue` throws that continuation away without ever re-walking
# it. The expensive branch's route to terminal 420 is never explored, so
# its true loss (0.50 + 0.02 = 0.52) is never compared against the cheap
# path at all. `best` keeps 0.07, silently missing the actually-worse 0.52
# critical path.
children8 = {400: [402, 401], 401: [410], 402: [410], 410: [420]}
ducts8 = build([
    (401, 0.05, 10.0, 800.0, 1.0, 'Supply Air'),
    (402, 0.50, 10.0, 800.0, 1.0, 'Supply Air'),
    (410, 0.02, 5.0,  800.0, 1.0, 'Supply Air'),
])
terms8 = {420: (100.0, 'Supply Air', 'SD-8')}
r8 = crit([400], children8, ducts8, terms8, {})
sa8 = r8.get(400, {}).get('Supply Air', {})
found_friction8 = sa8.get('friction_inwc')
true_worst8 = 0.50 + 0.02   # 0.52, via the 402 branch -- never discovered
cheap_result8 = 0.05 + 0.02  # 0.07, via the 401 branch -- what the walker reports
# FIXED 2026-09-24: the walker now carries a PER-PATH visited set instead of one
# shared across the whole DFS, so a node reached by two routes no longer blocks
# the second route at the merge point. This asserts the CORRECT behavior: the
# diamond's expensive branch must win. It caught a real bug that silently
# under-reported external static by 7x on a merged graph.
check("case8 diamond graph: the WORSE route through a shared downstream duct "
      "is found (0.52), not the cheap one (0.07)",
      found_friction8 is not None and abs(found_friction8 - true_worst8) < 1e-9,
      "found=%r  expected_worst=%.2f  cheap_path_would_be=%.2f"
      % (found_friction8, true_worst8, cheap_result8))


# ── case 9: component drops (diffuser, balancing damper, fire/smoke damper) ──
#
#   root(900) -> duct 901 -> BALANCING DAMPER 902 -> FIRE/SMOKE DAMPER 903
#             -> duct 904 -> BACKDRAFT DAMPER 905 -> terminal 920
#
# Each accessory kind carries its OWN value out of comp_values. The backdraft
# damper is its own key (Colin, 2026-11-02, 0.05), never anyone else's number.
COMP9 = {'diffuser': 0.05, 'damper': 0.25, 'fire_damper': 0.10, 'backdraft': 0.07}

children9 = {900: [901], 901: [902], 902: [903], 903: [904], 904: [905],
             905: [920]}
ducts9 = build([(901, 0.10, 10.0, 800, 1.0, 'Supply Air'),
                (904, 0.10, 10.0, 800, 1.0, 'Supply Air')])
nodes9 = {
    901: Elem('duct'),
    902: Elem('accessory', family_name='Balancing Damper - Round'),
    903: Elem('accessory', family_name='RJA - Fire Smoke Damper - Rectangular'),
    904: Elem('duct'),
    905: Elem('accessory', family_name='RJA - Backdraft Damper - Automatic - Round'),
}
terms9 = {920: (400.0, 'Supply Air', 'SD-9')}

r9 = crit([900], children9, ducts9, terms9, nodes9,
          safety_pct=0.0, comp_values=COMP9)
sa9 = r9[900]['Supply Air']
expected_comp9 = 0.25 + 0.10 + 0.07 + 0.05   # balancing + fire/smoke + backdraft + diffuser
check("case9 components == balancing 0.25 + fire/smoke 0.10 + backdraft 0.07 + diffuser 0.05",
      abs(sa9['component_inwc'] - expected_comp9) < 1e-9,
      "got %.4f expected %.4f" % (sa9['component_inwc'], expected_comp9))
check("case9 the fire/smoke damper got its OWN value, not the balancing one",
      abs(sa9['component_inwc'] - (0.25 * 2 + 0.07 + 0.05)) > 1e-6,
      "component_inwc=%.4f would be %.4f if it inherited 0.25"
      % (sa9['component_inwc'], 0.25 * 2 + 0.07 + 0.05))
check("case9 the backdraft damper is priced, not reported as unpriced",
      not any('Backdraft' in u for u in sa9['unpriced']),
      "unpriced=%r" % (sa9['unpriced'],))
check("case9 the backdraft default is 0.05 (Colin, 2026-11-02)",
      fitting_tables.component_of('backdraft') == 0.05)
check("case9 component_keys names exactly what was counted",
      sorted(sa9['component_keys']) == ['backdraft', 'damper', 'diffuser', 'fire_damper'],
      "got %r" % (sa9['component_keys'],))

# A component left at 0.00 contributes nothing but is still counted, so the
# report shows it as a deliberate zero rather than a miss.
r9b = crit([900], children9, ducts9, terms9, nodes9, safety_pct=0.0,
           comp_values={'diffuser': 0.05, 'damper': 0.25, 'fire_damper': 0.0,
                        'backdraft': 0.0})
check("case9 a 0.00 fire damper adds nothing but is still listed",
      abs(r9b[900]['Supply Air']['component_inwc'] - 0.30) < 1e-9
      and 'fire_damper' in r9b[900]['Supply Air']['component_keys'],
      "comp=%.4f keys=%r" % (r9b[900]['Supply Air']['component_inwc'],
                             r9b[900]['Supply Air']['component_keys']))

r9c = crit([900], children9, ducts9, terms9, nodes9, safety_pct=0.0,
           comp_values={'diffuser': 0.0, 'damper': 0.0, 'fire_damper': 0.0,
                        'backdraft': 0.0})
check("case9 zeroed component values contribute nothing",
      abs(r9c[900]['Supply Air']['component_inwc']) < 1e-12,
      "got %r" % r9c[900]['Supply Air']['component_inwc'])


# ══════════════════════════════════════════════════════════════════════════
# Case 10: TWO separate roots (two AHUs), each with its own disconnected
#          supply tree -- the regression test for the root_id -> sys_class
#          shape change. One unit's loss is clearly larger than the other's.
#          Before this change both roots' 'Supply Air' results were merged
#          into ONE dict keyed by system class, so the smaller AHU's number
#          was silently overwritten/hidden by whichever root's traversal
#          happened to win the `current is None or (...) >` comparison.
# ══════════════════════════════════════════════════════════════════════════
#   root(500, small AHU) -> duct 501 (fric 0.10) -> T510
#   root(600, big AHU)   -> duct 601 (fric 0.80) -> T610
#   (disjoint node id ranges and disjoint children/ducts entries -> the two
#   trees never touch each other)
children10 = {500: [501], 501: [510], 600: [601], 601: [610]}
ducts10 = build([
    (501, 0.10, 20.0,  800.0, 1.0, 'Supply Air'),
    (601, 0.80, 100.0, 800.0, 1.0, 'Supply Air'),
])
terms10 = {510: (300.0, 'Supply Air', 'SD-Small'),
           610: (500.0, 'Supply Air', 'SD-Big')}
r10 = crit([500, 600], children10, ducts10, terms10, {})
check("case10 exactly two top-level keys, one per root id",
      set(r10.keys()) == {500, 600}, "got %r" % (list(r10.keys()),))
small10 = r10.get(500, {}).get('Supply Air', {})
big10   = r10.get(600, {}).get('Supply Air', {})
check("case10 small AHU (root 500) Supply Air friction == 0.10 (its own tree)",
      abs(small10.get('friction_inwc', -1) - 0.10) < 1e-9,
      "got %r" % small10.get('friction_inwc'))
check("case10 big AHU (root 600) Supply Air friction == 0.80 (its own tree)",
      abs(big10.get('friction_inwc', -1) - 0.80) < 1e-9,
      "got %r" % big10.get('friction_inwc'))
check("case10 small AHU's terminal is its own (SD-Small), not the big unit's",
      small10.get('terminal_name') == 'SD-Small',
      "got %r" % small10.get('terminal_name'))
check("case10 big AHU's terminal is its own (SD-Big), not the small unit's",
      big10.get('terminal_name') == 'SD-Big',
      "got %r" % big10.get('terminal_name'))
check("case10 REGRESSION: small root's friction is NOT overwritten/hidden by "
      "the big root's value (the old merged-by-sys-class behavior would have "
      "made these equal)",
      abs(small10.get('friction_inwc', -1) - big10.get('friction_inwc', -2)) > 1e-6,
      "small=%r big=%r" % (small10.get('friction_inwc'), big10.get('friction_inwc')))


# ── case 11: dialog C overrides must reach the walker ────────────────────────
#
# The C table is editable at runtime now, so verify an override actually changes
# the answer rather than being collected and ignored. Reuses case 4's fixtures
# (root 100, one duct at 1000 FPM then a round elbow then a duct at 1600 FPM).
override_c = dict(fitting_tables.DEFAULT_C)
override_c['round_elbow_90'] = 0.66          # exactly 2x the published 0.33
r11 = crit([100], children4, ducts4, terms4, nodes4, c_values=override_c)
sa11 = r11[100]['Supply Air']
r11_default = crit([100], children4, ducts4, terms4, nodes4)
sa11_default = r11_default[100]['Supply Air']
check("case11 doubling the round-elbow C doubles the fitting loss",
      abs(sa11['fitting_inwc'] - 2.0 * sa11_default['fitting_inwc']) < 1e-12,
      "override=%.6f default=%.6f" % (sa11['fitting_inwc'],
                                      sa11_default['fitting_inwc']))
check("case11 the override did NOT disturb duct friction",
      abs(sa11['friction_inwc'] - sa11_default['friction_inwc']) < 1e-12,
      "override=%.6f default=%.6f" % (sa11['friction_inwc'],
                                      sa11_default['friction_inwc']))
# An override of a key the run never touches must change nothing at all.
untouched_c = dict(fitting_tables.DEFAULT_C)
untouched_c['abrupt_exit'] = 99.0
r11b = crit([100], children4, ducts4, terms4, nodes4, c_values=untouched_c)
check("case11 overriding a C the run never hits changes nothing",
      abs(r11b[100]['Supply Air']['fitting_inwc']
          - sa11_default['fitting_inwc']) < 1e-12,
      "got %.6f expected %.6f" % (r11b[100]['Supply Air']['fitting_inwc'],
                                  sa11_default['fitting_inwc']))


# ── case 12: fittings between a picked duct and the unit are always counted ──
#
# Colin, 2026-10-08. With no unit on the ductwork the run is rooted at the picked
# duct, and the stretch back toward the unit hangs off the root as a side branch
# that reaches no terminal. Real example: Kyrus RTU-3, picked duct 3570508, with
# a mitered elbow and a 0.2 ft stub (3570510 / 3570509) between it and the unit.
#
#   root duct 200 (800 FPM, 2.5 ft2)
#     |- 201 rect elbow -> 202 stub duct (no flow of its own, 0.2 ft)   <- unit side
#     |- 203 duct -> T210                                               <- the run
#     `- 205 tap (dead end, no terminal)                          <- NOT unit side
children12 = {200: [201, 203, 205], 201: [202], 203: [210]}
nodes12 = {201: Elem(kind='elbow', family_name='Rectangular Elbow - Mitered',
                     is_round=False),
           205: Elem(kind='tap', family_name='Round Takeoff - Shoe')}
ducts12 = build([
    (200, 0.0, 5.0, 800.0, 2.5, 'Supply Air'),
    (202, 0.0, 0.2, 0.0,   2.5, 'Supply Air'),
    (203, 0.0, 10.0, 800.0, 2.5, 'Supply Air'),
])
for _dr in ducts12.values():
    _dr.d_h_in = 18.0
terms12 = {210: (2000.0, 'Supply Air', 'SD-12')}

r12_off = crit([200], children12, ducts12, terms12, nodes12)
sa12_off = r12_off[200]['Supply Air']
r12 = crit([200], children12, ducts12, terms12, nodes12, duct_rooted_ids=set([200]))
sa12 = r12[200]['Supply Air']
exp_elbow12 = fitting_tables.DEFAULT_C['rect_elbow_90'] * pv(800.0)
# The dead-end tap on the root is still a tap the run passes by, so it carries the
# 0.20 main-duct coefficient in BOTH cases; that is not unit-side loss.
exp_bypass12 = fitting_tables.DEFAULT_C['supply_tap_main'] * pv(800.0)
check("case12 baseline (not duct-rooted): the unit-side elbow is NOT charged "
      "(only the bypassed tap is)",
      sa12_off['fitting_count'] == 0
      and abs(sa12_off['fitting_inwc'] - exp_bypass12) < 1e-9,
      "count=%r fit=%r expected bypass only %.6f"
      % (sa12_off['fitting_count'], sa12_off['fitting_inwc'], exp_bypass12))
check("case12 duct-rooted: the elbow behind the picked duct IS charged "
      "(C 0.25 x Pv(800), the root's own velocity)",
      abs(sa12['fitting_inwc'] - (exp_bypass12 + exp_elbow12)) < 1e-9,
      "expected %.6f got %r" % (exp_bypass12 + exp_elbow12, sa12['fitting_inwc']))
check("case12 duct-rooted: fitting_count == 1 and unit_side_fittings == 1",
      sa12['fitting_count'] == 1 and sa12['unit_side_fittings'] == 1,
      "count=%r unit=%r" % (sa12['fitting_count'], sa12['unit_side_fittings']))
check("case12 the stub's friction is added (0.05/100ft x 0.2 ft) and its length counted",
      abs(sa12['friction_inwc'] - 0.05 * 0.2 / 100.0) < 1e-12
      and abs(sa12['length_ft'] - (5.0 + 10.0 + 0.2)) < 1e-9,
      "fric=%r len=%r" % (sa12['friction_inwc'], sa12['length_ft']))
check("case12 a dead-end TAP off the root is NOT mistaken for the unit side "
      "(counted once as a bypassed tap, never as a unit-side fitting)",
      sa12['fitting_count'] == 1 and sa12['unit_side_fittings'] == 1
      and sa12['tap_bypass_count'] == 1,
      "count=%r unit=%r bypass=%r" % (sa12['fitting_count'],
                                       sa12['unit_side_fittings'],
                                       sa12['tap_bypass_count']))
check("case12 the unit side is added to the run once, not per branch "
      "(adding it changed the total by exactly one elbow)",
      abs((sa12['fitting_inwc'] - sa12_off['fitting_inwc']) - exp_elbow12) < 1e-9,
      "delta=%r expected %.6f" % (sa12['fitting_inwc'] - sa12_off['fitting_inwc'],
                                   exp_elbow12))


# ── case 13: Include Return Filter ───────────────────────────────────────────
#
# The filter is a known drop added once to the RETURN path total only. Reuses
# case 2's fixtures (root 0 supply, root 100 return).
r13_off = crit([0, 100], children2, ducts2, terms2, {}, safety_pct=10.0)
r13 = crit([0, 100], children2, ducts2, terms2, {}, safety_pct=10.0,
           return_filter_inwc=0.14)
ret_off = r13_off[100]['Return Air']
ret_on  = r13[100]['Return Air']
sup_on  = r13[0]['Supply Air']
check("case13 off by default: no filter on the return path",
      ret_off['filter_inwc'] == 0.0, "got %r" % ret_off['filter_inwc'])
check("case13 on: return filter_inwc == 0.14",
      abs(ret_on['filter_inwc'] - 0.14) < 1e-12, "got %r" % ret_on['filter_inwc'])
check("case13 on: return subtotal rises by exactly 0.14",
      abs(ret_on['subtotal_inwc'] - ret_off['subtotal_inwc'] - 0.14) < 1e-12,
      "on=%r off=%r" % (ret_on['subtotal_inwc'], ret_off['subtotal_inwc']))
check("case13 on: the 10% safety factor applies to the filter too",
      abs(ret_on['total_inwc'] - ret_on['subtotal_inwc'] * 1.10) < 1e-12)
check("case13 the SUPPLY path is untouched by the return filter",
      sup_on['filter_inwc'] == 0.0
      and abs(sup_on['total_inwc'] - r13_off[0]['Supply Air']['total_inwc']) < 1e-12)


# ── case 14: transitions are priced in the direction the AIR flows ───────────
#
# Colin, 2026-10-21. The walk always goes unit outward. Supply air goes the same
# way, return air the opposite way. One transition that GROWS along the walk
# (0.5 ft2 -> 1.0 ft2) is an expansion on supply but a contraction on return.
#   duct 701 (0.5 ft2) -> transition 702 (down_area 1.0) -> duct 703 -> T710
def _transition_net(sys_class):
    ch = {700: [701], 701: [702], 702: [703], 703: [710]}
    nd = {702: Elem(kind='transition', family_name='Round Transition - Angle',
                    is_round=True, down_area=1.0)}
    dc = build([(701, 0.0, 5.0, 1000.0, 0.5, sys_class),
                (703, 0.0, 5.0, 500.0, 1.0, sys_class)])
    tm = {710: (500.0, sys_class, 'T-14')}
    return crit([700], ch, dc, tm, nd)[700][sys_class]

sup14 = _transition_net('Supply Air')
ret14 = _transition_net('Return Air')
exp14 = fitting_tables.DEFAULT_C['transition_expansion'] * pv(1000.0)
con14 = fitting_tables.DEFAULT_C['transition_contraction'] * pv(1000.0)
check("case14 supply: a transition growing along the walk is an EXPANSION",
      abs(sup14['fitting_inwc'] - exp14) < 1e-9,
      "expected %.6f got %r" % (exp14, sup14['fitting_inwc']))
check("case14 return: the same transition is a CONTRACTION (air flows toward the unit)",
      abs(ret14['fitting_inwc'] - con14) < 1e-9,
      "expected %.6f got %r (expansion would be %.6f)"
      % (con14, ret14['fitting_inwc'], exp14))
ext14 = _transition_net('Exhaust Air')
check("case14 exhaust follows return (contraction)",
      abs(ext14['fitting_inwc'] - con14) < 1e-9, "got %r" % ext14['fitting_inwc'])


# ── case 15: unit-side stretch is walked TOWARD the unit, so the rule flips ──
#
#   picked duct 800 (0.5 ft2) -> transition 801 (down_area 1.0) -> stub 802  <- unit side
#                             -> duct 803 -> T810                           <- the run
# Walking from the picked duct to the unit, the area grows. For SUPPLY, air comes
# FROM the unit, so it really goes 1.0 -> 0.5: a contraction. For RETURN the air
# goes to the unit with the walk: an expansion.
def _unit_side_net(sys_class):
    ch = {800: [801, 803], 801: [802], 803: [810]}
    nd = {801: Elem(kind='transition', family_name='Round Transition - Angle',
                    is_round=True, down_area=1.0)}
    dc = build([(800, 0.0, 5.0, 1000.0, 0.5, sys_class),
                (802, 0.0, 0.5, 0.0, 1.0, sys_class),
                (803, 0.0, 5.0, 1000.0, 0.5, sys_class)])
    for _d in dc.values():
        _d.d_h_in = 10.0
    tm = {810: (500.0, sys_class, 'T-15')}
    return crit([800], ch, dc, tm, nd, duct_rooted_ids=set([800]))[800][sys_class]

sup15 = _unit_side_net('Supply Air')
ret15 = _unit_side_net('Return Air')
check("case15 unit side, supply: priced as a CONTRACTION (air flows away from the unit)",
      abs(sup15['fitting_inwc'] - con14) < 1e-9,
      "expected %.6f got %r" % (con14, sup15['fitting_inwc']))
check("case15 unit side, return: priced as an EXPANSION (air flows into the unit)",
      abs(ret15['fitting_inwc'] - exp14) < 1e-9,
      "expected %.6f got %r" % (exp14, ret15['fitting_inwc']))


# ── case 16: one flex bend per continuous flex run ───────────────────────────
#
# Colin, 2026-10-21: C 3.3 (top of the ASHRAE Duct Systems Design Guide's 0.9 to
# 3.3 range), once per flex run, at the flex's own velocity.
#   rigid 901 -> flex 902 -> flex 903 (same run, split in Revit) -> T910
def _flex_net(flex_ids):
    ch = {900: [901], 901: [902], 902: [903], 903: [910]}
    dc = build([(901, 0.0, 10.0, 900.0, 1.0, 'Supply Air'),
                (902, 0.0, 3.0, 600.0, 0.35, 'Supply Air'),
                (903, 0.0, 3.0, 600.0, 0.35, 'Supply Air')])
    for _d in dc.values():
        _d.is_flex = _d.element_id in flex_ids
    tm = {910: (210.0, 'Supply Air', 'T-16')}
    return crit([900], ch, dc, tm, {})[900]['Supply Air']

none16 = _flex_net(set())
two16  = _flex_net(set([902, 903]))
bend16 = fitting_tables.DEFAULT_C['flex_bend'] * pv(600.0)
check("case16 no flex on the path: no bend charged",
      abs(none16['fitting_inwc']) < 1e-12, "got %r" % none16['fitting_inwc'])
check("case16 a flex run split in two segments gets ONE bend, at the flex velocity",
      abs(two16['fitting_inwc'] - bend16) < 1e-9 and two16['fitting_count'] == 1,
      "expected %.6f got %r count %r" % (bend16, two16['fitting_inwc'],
                                          two16['fitting_count']))
check("case16 flex_bend default is 3.3",
      fitting_tables.DEFAULT_C['flex_bend'] == 3.3)
ovr16 = dict(fitting_tables.DEFAULT_C)
ovr16['flex_bend'] = 0.9
ch16 = {900: [901], 901: [902], 902: [903], 903: [910]}
dc16 = build([(901, 0.0, 10.0, 900.0, 1.0, 'Supply Air'),
              (902, 0.0, 3.0, 600.0, 0.35, 'Supply Air'),
              (903, 0.0, 3.0, 600.0, 0.35, 'Supply Air')])
for _d in dc16.values():
    _d.is_flex = _d.element_id in (902, 903)
r16o = crit([900], ch16, dc16, {910: (210.0, 'Supply Air', 'T-16')}, {},
            c_values=ovr16)[900]['Supply Air']
check("case16 the dialog override of flex_bend reaches the walker",
      abs(r16o['fitting_inwc'] - 0.9 * pv(600.0)) < 1e-9,
      "got %r" % r16o['fitting_inwc'])


# ── case 17: VAV box mid tree (VAV plan, Colin 2026-10-22) ───────────────────
#
#   AHU 1000 (root, equipment) -> MP duct 1001 (1800 FPM, 0.20 fric)
#     -> BOX 1002 -> elbow 1003 (round) -> LP duct 1004 (600 FPM, 0.05) -> T1010
# The box is priced at the VAV box drop, the elbow right on its discharge is
# priced at the OUTLET duct's 600 FPM (not the MP 1800), and the MP loss to the
# box inlet is reported per box.
children17 = {1000: [1001], 1001: [1002], 1002: [1003], 1003: [1004], 1004: [1010]}
nodes17 = {1000: Elem(kind='equipment'), 1002: Elem(kind='equipment'),
           1003: Elem(kind='elbow', family_name='Round Elbow', is_round=True)}
ducts17 = build([(1001, 0.20, 40.0, 1800.0, 0.5, 'Supply Air'),
                 (1004, 0.05, 15.0, 600.0, 0.6, 'Supply Air')])
terms17 = {1010: (360.0, 'Supply Air', 'SD-17')}
COMP17 = {'diffuser': 0.05, 'damper': 0.25, 'fire_damper': 0.05}

r17 = crit([1000], children17, ducts17, terms17, nodes17, safety_pct=10.0,
           comp_values=COMP17, vav_box_ids=set([1002]), vav_box_drop_inwc=0.40)
sa17 = r17[1000]['Supply Air']
check("case17 box drop is a component (0.40 box + 0.05 diffuser + 0.25 assumed damper)",
      abs(sa17['component_inwc'] - 0.70) < 1e-9, "got %r" % sa17['component_inwc'])
check("case17 component_keys name the box",
      sorted(sa17['component_keys']) == ['damper', 'diffuser', 'vav_box'],
      "got %r" % (sa17['component_keys'],))
exp_elbow17 = fitting_tables.DEFAULT_C['round_elbow_90'] * pv(600.0)
wrong_elbow17 = fitting_tables.DEFAULT_C['round_elbow_90'] * pv(1800.0)
check("case17 the elbow on the box discharge is priced at the OUTLET duct velocity",
      abs(sa17['fitting_inwc'] - exp_elbow17) < 1e-9
      and abs(sa17['fitting_inwc'] - wrong_elbow17) > 1e-6,
      "expected %.6f got %r (inlet velocity would give %.6f)"
      % (exp_elbow17, sa17['fitting_inwc'], wrong_elbow17))
check("case17 the root AHU itself is not priced as a component",
      list(sa17['component_keys']).count('vav_box') == 1
      and 'mid_equip' not in sa17['component_keys'])
box17 = sa17.get('boxes', {}).get(1002)
check("case17 MP loss to the box inlet is the MP duct only (0.20), x1.10 with SF",
      box17 is not None and abs(box17[0] - 0.20) < 1e-9
      and abs(box17[1] - 0.22) < 1e-9, "got %r" % (box17,))
_item_sum17 = sum(i[3] or 0.0 for i in sa17.get('items', ()))
check("case17 index run items still add up to the subtotal",
      abs(_item_sum17 - sa17['subtotal_inwc']) < 1e-9,
      "items %r vs subtotal %r" % (_item_sum17, sa17['subtotal_inwc']))
r17d = crit([1000], children17, ducts17, terms17, nodes17, comp_values=COMP17,
            vav_box_ids=set([1002]))
check("case17 box drop defaults to the balancing damper value",
      abs(r17d[1000]['Supply Air']['component_inwc'] - (0.25 + 0.05 + 0.25)) < 1e-9,
      "got %r" % r17d[1000]['Supply Air']['component_inwc'])

# VAV OFF: the same equipment is priced as a balancing damper, and nothing is
# reset (the elbow keeps the inlet duct's velocity, as before).
r17off = crit([1000], children17, ducts17, terms17, nodes17, comp_values=COMP17)
sa17off = r17off[1000]['Supply Air']
check("case17 VAV off: mid tree equipment has its own key, at the damper value",
      sorted(sa17off['component_keys']) == ['damper', 'diffuser', 'mid_equip']
      and abs(sa17off['component_inwc'] - 0.55) < 1e-9,
      "keys %r comp %r" % (sa17off['component_keys'], sa17off['component_inwc']))
check("case17 VAV off: no box list",
      sa17off.get('boxes') == {}, "got %r" % (sa17off.get('boxes'),))

# Worst full path across two boxes: MP to box, box, worst LP inside.
#   AHU 1100 -> MP 1101 -> tap A 1102 -> MP 1103 (0.10) -> BOX 1104 -> LP 1105 (0.30) -> T1110
#                       -> tap B 1106 -> MP 1107 (0.25) -> BOX 1108 -> LP 1109 (0.05) -> T1111
children18 = {1100: [1101], 1101: [1102, 1106], 1102: [1103], 1103: [1104],
              1104: [1105], 1105: [1110], 1106: [1107], 1107: [1108],
              1108: [1109], 1109: [1111]}
nodes18 = {1100: Elem(kind='equipment'),
           1102: Elem(kind='tap', family_name='Round Takeoff'),
           1106: Elem(kind='tap', family_name='Round Takeoff'),
           1104: Elem(kind='equipment'), 1108: Elem(kind='equipment')}
ducts18 = build([(1101, 0.0, 10.0, 0.0, 1.0, 'Supply Air'),
                 (1103, 0.10, 10.0, 0.0, 0.3, 'Supply Air'),
                 (1105, 0.30, 30.0, 0.0, 0.5, 'Supply Air'),
                 (1107, 0.25, 30.0, 0.0, 0.3, 'Supply Air'),
                 (1109, 0.05, 5.0, 0.0, 0.5, 'Supply Air')])
terms18 = {1110: (300.0, 'Supply Air', 'SD-A'), 1111: (300.0, 'Supply Air', 'SD-B')}
r18 = crit([1100], children18, ducts18, terms18, nodes18,
           comp_values=COMP17, vav_box_ids=set([1104, 1108]), vav_box_drop_inwc=0.25)
sa18 = r18[1100]['Supply Air']
check("case18 the worst FULL path wins (0.10 + 0.30 beats 0.25 + 0.05)",
      sa18['terminal_name'] == 'SD-A' and abs(sa18['friction_inwc'] - 0.40) < 1e-9,
      "terminal %r friction %r" % (sa18['terminal_name'], sa18['friction_inwc']))
check("case18 each box gets its own MP loss (0.10 and 0.25)",
      abs(sa18['boxes'][1104][0] - 0.10) < 1e-9
      and abs(sa18['boxes'][1108][0] - 0.25) < 1e-9,
      "got %r" % (sa18['boxes'],))


# ── case 19: through a fan powered box the walk stays on its own side (C2) ───
#
#   AHU 1200 -> MP 1201 (SA) -> FPB 1202 -> LP 1203 (SA, 0.05) -> diffuser T1210
#                                        `-> return duct 1204 (RA, 0.90) -> grille T1211
# Before the fix the walk went MP -> box -> ducted return, and that 0.90 return
# duct turned up as a Return Air index run of the AHU.
children19 = {1200: [1201], 1201: [1202], 1202: [1203, 1204], 1203: [1210],
              1204: [1211]}
nodes19 = {1200: Elem(kind='equipment'), 1202: Elem(kind='equipment')}
ducts19 = build([(1201, 0.20, 40.0, 1800.0, 0.5, 'Supply Air'),
                 (1203, 0.05, 10.0, 600.0, 0.6, 'Supply Air'),
                 (1204, 0.90, 10.0, 500.0, 0.6, 'Return Air')])
terms19 = {1210: (400.0, 'Supply Air', 'SD-19'), 1211: (300.0, 'Return Air', 'RG-19')}
for _box19 in (set([1202]), set()):
    r19 = crit([1200], children19, ducts19, terms19, nodes19,
               comp_values=COMP17, vav_box_ids=_box19)
    check("case19 (%s) the FPB's ducted return is NOT a Return Air run of the AHU"
          % ('VAV on' if _box19 else 'VAV off'),
          set(r19.get(1200, {}).keys()) == set(['Supply Air']),
          "got classes %r" % (list(r19.get(1200, {}).keys()),))
check("case19 the supply run through the box is still found",
      r19[1200]['Supply Air']['terminal_name'] == 'SD-19')


# ── case 20: a box that is ALSO a picked root is still priced in the AHU walk (C3)
r20 = crit([1000, 1002], children17, ducts17, terms17, nodes17, safety_pct=0.0,
           comp_values=COMP17, vav_box_ids=set([1002]), vav_box_drop_inwc=0.40)
sa20 = r20[1000]['Supply Air']
check("case20 the AHU's walk still charges the box (picked as a root too)",
      'vav_box' in sa20['component_keys']
      and abs(sa20['component_inwc'] - 0.70) < 1e-9,
      "keys %r comp %r" % (sa20['component_keys'], sa20['component_inwc']))
check("case20 and records the MP loss to it",
      1002 in sa20.get('boxes', {}), "got %r" % (sa20.get('boxes'),))
check("case20 the box's own walk does not charge itself",
      'vav_box' not in r20[1002]['Supply Air']['component_keys'],
      "got %r" % (r20[1002]['Supply Air']['component_keys'],))


# ── case 21: elbow angle factor (ASHRAE 2001 CR3-1, Colin 2026-10-23) ─────────
af = fitting_tables.elbow_angle_factor
check("case21 K table rows: 30 -> 0.45, 45 -> 0.60, 90 -> 1.00",
      abs(af(30) - 0.45) < 1e-9 and abs(af(45) - 0.60) < 1e-9
      and abs(af(90) - 1.00) < 1e-9)
check("case21 K interpolates: 35.6 deg between 0.45 and 0.60",
      abs(af(35.6) - (0.45 + 0.15 * 5.6 / 15.0)) < 1e-9, "got %r" % af(35.6))
check("case21 unreadable angle is priced as a 90 (K 1.0)", af(None) == 1.0)
check("case21 angles past the table clamp to its ends",
      af(-5) == 0.0 and af(200) == 1.40)
# Same path as case 4, elbow now 45 deg: priced at 0.60 x the 90 deg C.
nodes21 = {101: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                     angle=45.0)}
sa21 = crit([100], children4, ducts4, terms4, nodes21).get(100, {}).get('Supply Air', {})
exp21 = 0.60 * fitting_tables.DEFAULT_C['round_elbow_90'] * pv(1000.0)
check("case21 a 45 deg round elbow costs 0.60 x the 90 deg loss",
      abs(sa21.get('fitting_inwc', -1) - exp21) < 1e-9,
      "expected %.6f got %r" % (exp21, sa21.get('fitting_inwc')))
nodes21r = {101: Elem(kind='elbow', family_name='Rectangular Elbow - Mitered',
                      is_round=False, angle=30.0)}
sa21r = crit([100], children4, ducts4, terms4, nodes21r).get(100, {}).get('Supply Air', {})
exp21r = 0.45 * fitting_tables.DEFAULT_C['rect_elbow_90'] * pv(1000.0)
check("case21 a 30 deg rect elbow costs 0.45 x the 90 deg loss",
      abs(sa21r.get('fitting_inwc', -1) - exp21r) < 1e-9,
      "expected %.6f got %r" % (exp21r, sa21r.get('fitting_inwc')))



# ── case 22 to 27: ASHRAE 2001 lookups (Colin, 2026-10-23) ───────────────────
A = ashrae_fittings
def _sa(r, root):
    return r.get(root, {}).get('Supply Air', {})

# 22: round elbow, 12 in., r/D 1.5, 90 deg -> CD3-1 0.11 at the upstream Pv.
n22 = {101: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                 geom={'round': True, 'd_in': 12.0, 'w_in': None, 'h_in': None,
                       'r_in': 18.0, 'angle': 90.0})}
sa22 = _sa(crit([100], children4, ducts4, terms4, n22), 100)
check("case22 round 1.5D 90 deg elbow is CD3-1 0.11 x Pv",
      abs(sa22.get('fitting_inwc', -1) - 0.11 * pv(1000.0)) < 1e-9,
      "got %r" % sa22.get('fitting_inwc'))
check("case22 and raises no FALLBACK note", not sa22.get('unpriced'),
      "got %r" % (sa22.get('unpriced'),))

# 23: 1D round elbow at 45 deg -> CD3-1 x CD3-12(1.0)/CD3-12(1.5) x K(45).
n23 = {101: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                 geom={'round': True, 'd_in': 12.0, 'w_in': None, 'h_in': None,
                       'r_in': 12.0, 'angle': 45.0})}
sa23 = _sa(crit([100], children4, ducts4, terms4, n23), 100)
exp23 = 0.11 * (0.42 / 0.34) * 0.60 * pv(1000.0)
check("case23 1D round elbow at 45 deg: CD3-1 x 0.42/0.34 x K 0.60",
      abs(sa23.get('fitting_inwc', -1) - exp23) < 1e-9,
      "expected %.6f got %r" % (exp23, sa23.get('fitting_inwc')))

# 24: rect mitered at 30 deg (no vanes, CR3-6) and at 90 deg (vanes, CR3-9).
def _mit(angle):
    return {101: Elem(kind='elbow', family_name='Rectangular Elbow - Mitered - Single Thickness Vanes',
                      is_round=False,
                      geom={'round': False, 'd_in': None, 'w_in': 16.0, 'h_in': 16.0,
                            'r_in': None, 'angle': angle})}
sa24a = _sa(crit([100], children4, ducts4, terms4, _mit(30.0)), 100)
sa24b = _sa(crit([100], children4, ducts4, terms4, _mit(90.0)), 100)
check("case24 mitered 30 deg, H/W 1.0 is CR3-6 0.16 x Pv",
      abs(sa24a.get('fitting_inwc', -1) - 0.16 * pv(1000.0)) < 1e-9,
      "got %r" % sa24a.get('fitting_inwc'))
check("case24 mitered 90 deg is CR3-9 vaned 0.11 x Pv",
      abs(sa24b.get('fitting_inwc', -1) - 0.11 * pv(1000.0)) < 1e-9,
      "got %r" % sa24b.get('fitting_inwc'))

# 25: supply rect transition 1.0 ft2 -> 0.5 ft2 at 30 deg: SR4-1 Ao/A1 0.50,
#     30 deg = 0.05, referenced to section o (downstream, 2000 fpm).
n25 = {101: Elem(kind='transition', family_name='Rectangular Transition - Length',
                 is_round=False, down_area=0.5,
                 tgeom=([(1.0, 'rect', 12.0, 12.0, None), (0.5, 'rect', 6.0, 12.0, None)], 30.0))}
sa25 = _sa(crit([100], children4, ducts4, terms4, n25), 100)
check("case25 supply transition is SR4-1 0.05 at the section o (downstream) Pv",
      abs(sa25.get('fitting_inwc', -1) - 0.05 * pv(2000.0)) < 1e-9,
      "expected %.6f got %r" % (0.05 * pv(2000.0), sa25.get('fitting_inwc')))

# 26: two rect taps on one unbroken rect main, in order along the duct.
#   root 299 -> main 300 (1000 cfm, 1.0 ft2) -> tap 301 at 2 ft -> 311 (300 cfm) -> T1
#                                            -> tap 302 at 5 ft -> 312 (400 cfm) -> T2
ch26 = {299: [300], 300: [301, 302], 301: [311], 302: [312], 311: [321], 312: [322]}
n26 = {300: Elem(kind='duct', tap_pos={301: 2.0, 302: 5.0}),
       301: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       302: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       311: Elem(kind='duct'), 312: Elem(kind='duct')}
# 0.50 in. wc of friction on the far branch makes it the worst run, so the
# far tap's pricing (ordering, local flow) is what the result reports.
d26 = build([(300, 0.0, 10.0, 1000.0, 1.0, 'Supply Air'),
             (311, 0.0, 5.0, 600.0, 0.5, 'Supply Air'),
             (312, 0.5, 5.0, 800.0, 0.5, 'Supply Air')])
t26 = {321: (300.0, 'Supply Air', 'SD-1'), 322: (400.0, 'Supply Air', 'SD-2')}
r26 = crit([299], ch26, d26, t26, n26)
sa26 = _sa(r26, 299)
# The printed Cb (branch reference) and Cs (straight reference) are converted
# to the main (Eq. 34) and multiplied by the main Pv at each tap. On a full
# SR5-13 grid that equals Cb x branch Pv: far branch 400 cfm / 0.5 ft2 =
# 800 fpm. Cs is read on the As/Ac 0.9 row, so Vs/Vc = 0.7 / 0.9.
ccb_far, _l, _n = A.tap_branch_c(True, 'rect', 'rect', True, 0.5, 400.0 / 700.0)
ccs_near, _l, _n = A.tap_main_c(True, 'rect', 'rect', True, 1.0, 0.7)
cb_far_printed = 0.73 + (0.54 - 0.73) * ((400.0 / 700.0) - 0.5) / 0.1
check("case26 Cc,b x main Pv equals printed Cb x branch Pv on a full grid",
      abs(ccb_far * pv(700.0) - cb_far_printed * pv(800.0)) < 1e-9,
      "%r vs %r" % (ccb_far * pv(700.0), cb_far_printed * pv(800.0)))
exp26 = ccs_near * pv(1000.0) + ccb_far * pv(700.0)
check("case26 worst run leaves through the FAR tap",
      sa26.get('terminal_id') == 322, "got %r" % sa26.get('terminal_id'))
check("case26 far tap: near tap passed (Cc,s x main Pv) plus the far tap Cc,b at the LOCAL 700 cfm x main Pv",
      abs(sa26.get('fitting_inwc', -1) - exp26) < 1e-9,
      "expected %.6f got %r" % (exp26, sa26.get('fitting_inwc')))
check("case26 only the tap AHEAD of the exit is counted as passed",
      sa26.get('tap_bypass_count') == 1, "got %r" % sa26.get('tap_bypass_count'))
# and the near tap, made the worst run instead, sees the duct's full 1000 cfm
# and passes no tap.
d26b = build([(300, 0.0, 10.0, 1000.0, 1.0, 'Supply Air'),
              (311, 0.5, 5.0, 600.0, 0.5, 'Supply Air'),
              (312, 0.0, 5.0, 800.0, 0.5, 'Supply Air')])
sa26b = _sa(crit([299], ch26, d26b, t26, n26), 299)
ccb_near, _l, _n = A.tap_branch_c(True, 'rect', 'rect', True, 0.5, 0.3)
check("case26 near tap: Cb 2.07 (SR5-13 Ab/Ac 0.5, Qb/Qc 0.3) = 2.07 x branch Pv (600 fpm), no tap passed",
      sa26b.get('terminal_id') == 321 and sa26b.get('tap_bypass_count') == 0
      and abs(ccb_near - 2.07 * 0.36) < 1e-9
      and abs(sa26b.get('fitting_inwc', -1) - 2.07 * pv(600.0)) < 1e-9,
      "term %r taps %r fit %r" % (sa26b.get('terminal_id'), sa26b.get('tap_bypass_count'),
                                   sa26b.get('fitting_inwc')))

# 27: elbow with unreadable geometry falls back to the worksheet, flagged.
n27 = {101: Elem(kind='elbow', family_name='Round Elbow', is_round=True)}
sa27 = _sa(crit([100], children4, ducts4, terms4, n27), 100)
check("case27 unreadable elbow is priced at the RJA worksheet 0.33",
      abs(sa27.get('fitting_inwc', -1) - 0.33 * pv(1000.0)) < 1e-9,
      "got %r" % sa27.get('fitting_inwc'))
check("case27 and is reported as a FALLBACK",
      any(u.startswith('FALLBACK') for u in sa27.get('unpriced', ())),
      "got %r" % (sa27.get('unpriced'),))


# 28: a tap taking 900 of 1000 cfm, the run continuing on the main with the
#     last 100 cfm. Cs (As/Ac 0.9 row, Qs/Qc 0.1) is 51.37, referenced to the
#     straight-through Pv at 100 fpm: about 0.032 in. wc, not 3.2 (review C1).
ch28 = {399: [400], 400: [401, 402], 401: [411], 402: [422], 411: [431]}
n28 = {400: Elem(kind='duct', tap_pos={401: 2.0}),
       401: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       402: Elem(kind='duct'), 411: Elem(kind='duct')}
d28 = build([(400, 0.0, 10.0, 1000.0, 1.0, 'Supply Air'),
             (402, 0.5, 10.0, 100.0, 1.0, 'Supply Air'),
             (411, 0.0, 5.0, 1800.0, 0.5, 'Supply Air')])
t28 = {422: (100.0, 'Supply Air', 'SD-end'), 431: (900.0, 'Supply Air', 'SD-big')}
sa28 = _sa(crit([399], ch28, d28, t28, n28), 399)
ccs28, _l, _n = A.tap_main_c(True, 'rect', 'rect', True, 1.0, 0.1)
check("case28 low split: printed SR5-13 Cs 51.37 (As/Ac 0.9, Qs/Qc 0.1) is Cc,s 51.37 x (0.1/0.9)^2",
      abs(ccs28 - 51.37 * (0.1 / 0.9) ** 2) < 1e-9, "got %r" % ccs28)
check("case28 and costs Cc,s x main Pv (1000 fpm), about 0.04 in. wc, not 3.2",
      sa28.get('terminal_id') == 422
      and abs(sa28.get('fitting_inwc', -1) - ccs28 * pv(1000.0)) < 1e-9
      and sa28.get('fitting_inwc', 9) < 0.05,
      "term %r fit %r" % (sa28.get('terminal_id'), sa28.get('fitting_inwc')))


# 29: ER5-3 at its printed geometry (Ab/Ac 0.5), whatever the real branch,
#     and a tap C of 0 or less replaced by the RJA worksheet C (Colin,
#     2026-11-09, was floored at zero).
#   Small grille branch: 300 of 3000 cfm (Qb/Qc 0.1) through a 0.35 ft2 branch
#   on a 4 ft2 return main (real Ab/Ac 0.09). Printed Cb -18.00 at Qb/Qc 0.1;
#   at the printed Ab/Ac 0.5, Cc,b = -18 x (0.1/0.5)^2 = -0.72, so the
#   worksheet exhaust_tap_branch 0.76 x the main Pv is charged instead.
ccb29, _l, _n = A.tap_branch_c(False, 'rect', 'rect', True, 0.09, 0.1)
check("case29 ER5-3 is read at its printed Ab/Ac 0.5: Cc,b -0.72, not -18 x (0.1/0.09)^2",
      abs(ccb29 - (-18.0 * (0.1 / 0.5) ** 2)) < 1e-9, "got %r" % ccb29)
ccb29b, _l, _n = A.tap_branch_c(False, 'rect', 'rect', True, 0.09, 0.5)
check("case29 ER5-3 at Qb/Qc 0.5: Cc,b = 0.76 x (0.5/0.5)^2 = 0.76",
      abs(ccb29b - 0.76) < 1e-9, "got %r" % ccb29b)
# 29c: the last tap before an end cap, Qb/Qc 1.00 past the highest printed
#      column (0.90): the 0.90 column is read with no clamp NOTE (Colin,
#      2026-11-09, TB-M15 FCU-1 tap 3724988); the price is as before. Other
#      clamp NOTEs (Ab/Ac past its range) stay.
for _sup29, _tap29, _main29, _45 in ((True, 'round', 'rect', False),
                                     (False, 'round', 'round', True)):
    _c10, _lb10, _n10 = A.tap_branch_c(_sup29, _tap29, _main29, _45, 0.5, 1.0)
    _key29 = _lb10.split(' ')[0]
    _cb09, _nb = A._call(_key29 + '_Cb', A.table(_key29 + '_Cb'),
                         {'Ab/Ac': 0.5, 'Qb/Qc': 0.9, 'As/Ac': 1.0, 'Qs/Qc': 0.1})
    _read_ab29 = A._ratio_read_at(A.table(_key29 + '_Cb'), 'Ab/Ac', 0.5)
    check("case29c %s Qb/Qc 1.00: no Qb/Qc clamp NOTE, highest column read, price unchanged"
          % _key29,
          not any('Qb/Qc' in n and 'highest' in n for n in _n10)
          and abs(_c10 - A._to_common(_cb09, 1.0, _read_ab29)) < 1e-12,
          "C %r notes %r" % (_c10, _n10))
_cab29, _l, _nab29 = A.tap_branch_c(True, 'round', 'rect', False, 0.95, 1.0)
check("case29c an Ab/Ac clamp NOTE still prints",
      any(n.startswith('SR5-11_Cb: Ab/Ac 0.95 is above') for n in _nab29)
      and not any('Qb/Qc' in n for n in _nab29), "notes %r" % (_nab29,))
ch29 = {499: [500], 500: [501], 501: [511], 511: [521]}
n29 = {500: Elem(kind='duct', tap_pos={501: 2.0}),
       501: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       511: Elem(kind='duct')}
d29 = build([(500, 0.0, 10.0, 750.0, 4.0, 'Return Air'),
             (511, 0.0, 5.0, 857.0, 0.35, 'Return Air')])
t29 = {521: (300.0, 'Return Air', 'RG-1')}
ra29 = crit([499], ch29, d29, t29, n29).get(499, {}).get('Return Air', {})
check("case29 a negative return tap Cc,b takes the worksheet C 0.76 x the main Pv",
      abs(ra29.get('fitting_inwc', -1)
          - fitting_tables.DEFAULT_C['exhaust_tap_branch'] * pv(750.0)) < 1e-12,
      "fit %r" % (ra29.get('fitting_inwc'),))
check("case29 FALLBACK line names the tap, the table value and the worksheet C",
      any(u.startswith('FALLBACK: tap 501 into branch off duct 500, 2001 ER5-3 Cc,b '
                       '-0.72 is not positive; RJA worksheet C 0.76 used')
          for u in ra29.get('unpriced', ()))
      and not any('floored' in u for u in ra29.get('unpriced', ())),
      "notes %r" % (ra29.get('unpriced'),))

# 30: a main that ends in a tap is a turn into the branch (Colin, 2026-10-26):
#     priced as a 90 degree CR3-1 elbow (r/W 1.5, H/W as drawn) at the main's Pv,
#     not the tap's Cc,b (which was 5.96 x Pv on a return end tap).
ch30 = {599: [600], 600: [601], 601: [611], 611: [621]}
n30 = {600: Elem(kind='duct', tap_pos={601: 2.0}),
       601: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       611: Elem(kind='duct')}
for _cls30 in ('Supply Air', 'Return Air'):
    d30 = build([(600, 0.0, 10.0, 600.0, 1.0, _cls30),
                 (611, 0.0, 5.0, 764.0, 0.785, _cls30)])
    for _dr in d30.values():
        _dr.size = '12x12"' if _dr.element_id == 600 else '12"dia'
    t30 = {621: (600.0, _cls30, 'T-end')}
    r30 = crit([599], ch30, d30, t30, n30).get(599, {}).get(_cls30, {})
    ce30 = A.rect_smooth_elbow_c(1.5, 1.0, 90.0)[0]
    check("case30 %s end of main tap is a CR3-1 elbow x main Pv, not Cc,b" % _cls30,
          abs(r30.get('fitting_inwc', -1) - ce30 * pv(600.0)) < 1e-9
          and abs(ce30 - 0.17) < 1e-9,
          "got %r expected %.6f" % (r30.get('fitting_inwc'), ce30 * pv(600.0)))
    check("case30 %s: a damper is assumed on supply only, never on return" % _cls30,
          ('damper' in r30.get('component_keys', ())) == (_cls30 == 'Supply Air'),
          "keys %r" % (r30.get('component_keys'),))
# a drawn damper is not charged twice
n30b = dict(n30)
n30b[611] = Elem(kind='duct')
n30b[612] = Elem(kind='accessory', family_name='Balancing Damper - Rectangular')
ch30b = {599: [600], 600: [601], 601: [612], 612: [611], 611: [621]}
d30b = build([(600, 0.0, 10.0, 600.0, 1.0, 'Supply Air'),
              (611, 0.0, 5.0, 764.0, 0.785, 'Supply Air')])
for _dr in d30b.values():
    _dr.size = '12x12"' if _dr.element_id == 600 else '12"dia'
r30b = crit([599], ch30b, d30b, {621: (600.0, 'Supply Air', 'T-end')}, n30b)[599]['Supply Air']
check("case30 a drawn damper is charged once, not again as the assumed one",
      list(r30b['component_keys']).count('damper') == 1,
      "keys %r" % (r30b['component_keys'],))


# ── case 31 to 40: fan system effect (Colin, 2026-10-29) ─────────────────────
#   2001 ASHRAE Fundamentals: Le Eq. 37 (p. 34.13), SR7-9 to SR7-12 DWDI
#   outlet (p. 34.67 to 34.68), SR7-5 to SR7-8 SWSI, ED7-2 (p. 34.48) and
#   ER7-1 (p. 34.60) inlet. Unit outlet 24 x 24 in. (4 ft2) at 4000 cfm
#   (1000 fpm): Le = 609.6 mm / 350 = 1.742 m = 5.71 ft.
import fan_system_effect as FSE  # noqa: E402
# The cases below price at the connection velocity, the basis they were
# written for; the default fan velocity basis has its own cases (case60 on).
FSE_ON = {'on': True, 'fan_type': 'DWDI', 'blast_ratio': 0.5,
          'velocity_basis': 'connection'}
SQ24 = {'round': False, 'area_ft2': 4.0, 'd_in': None, 'w_in': 24.0, 'h_in': 24.0}
RD24 = {'round': True, 'area_ft2': math.pi, 'd_in': 24.0, 'w_in': None, 'h_in': None}
LE24 = 609.6 / 350.0 / 0.3048


def _unit(conns):
    u = Elem(kind='equipment')
    u.conns = conns
    return u


def _fse_run(nodes, children, ducts, terms, root, fan=FSE_ON, cls='Supply Air', **kw):
    r = crit([root], children, ducts, terms, nodes, 10.0, fan_effect=fan, **kw)
    return r.get(root, {}).get(cls, {})


def _items_sum(leaf):
    return sum(i[3] or 0.0 for i in leaf.get('items', ()))


def _fse_items(leaf):
    return [i for i in leaf.get('items', ())
            if i[4] and i[4][0].startswith('Fan system effect')]


# 31: nothing within Le of the outlet: no system effect, total unchanged.
n31 = {900: _unit({901: SQ24})}
c31 = {900: [901], 901: [910]}
d31 = build([(901, 0.10, 20.0, 1000.0, 4.0, 'Supply Air')])
t31 = {910: (4000.0, 'Supply Air', 'SD-31')}
on31 = _fse_run(n31, c31, d31, t31, 900)
off31 = _fse_run(n31, c31, d31, t31, 900, fan=None)
check("case31 straight 20 ft from the outlet (Le 5.71 ft): fse 0, total as with it off",
      on31.get('fse_inwc') == 0.0
      and abs(on31['total_inwc'] - off31['total_inwc']) < 1e-12,
      "on %r off %r" % (on31.get('total_inwc'), off31.get('total_inwc')))
# Outlet and inlet always read apart (Colin, 2026-11-05): an unpriced
# connection still prints its own take-off row at 0.000 with the reason.
_it31 = _fse_items(on31)
check("case31 one 0.000 take-off row 'Fan system effect, supply outlet', none: no elbow within Le",
      len(_it31) == 1 and _it31[0][3] == 0.0
      and _it31[0][4][0] == 'Fan system effect, supply outlet'
      and _it31[0][4][1].startswith('none: no elbow within Le (5.7 ft)'),
      "items %r" % (_it31,))
check("case31 fan system effect off: no fan system effect row at all",
      not _fse_items(off31), "items %r" % (_fse_items(off31),))

# 32: elbow right on the outlet (L 0): DWDI Ab/Ao 0.5, worst of A-D is
#     SR7-11 C 3.80 (A 2.20, B 3.60, D 3.20): GRAY (Colin, 2026-10-29).
n32 = {920: _unit({921: SQ24}),
       921: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c32 = {920: [921], 921: [922], 922: [930]}
d32 = build([(922, 0.0, 10.0, 1000.0, 4.0, 'Supply Air')])
t32 = {930: (4000.0, 'Supply Air', 'SD-32')}
sa32 = _fse_run(n32, c32, d32, t32, 920)
exp32 = 3.80 * pv(1000.0)
check("case32 elbow at L 0: FSE2 = 3.80 (SR7-11, worst of A-D) x Pv(1000)",
      abs(sa32.get('fse_inwc', -1) - exp32) < 1e-9,
      "got %r expected %.6f" % (sa32.get('fse_inwc'), exp32))
check("case32 C 3.80 at 1000 fpm is 0.237 in. wc, under 0.5: NOT gray (Colin, 2026-11-02)",
      not any(k == 'gray' for k, t in sa32.get('fse_notes', ())),
      "notes %r" % (sa32.get('fse_notes'),))
check("case32 added before the safety factor: total = (subtotal incl. FSE) x 1.10",
      abs(sa32['subtotal_inwc'] - (sa32['friction_inwc'] + sa32['fitting_inwc'] +
                                   sa32['component_inwc'] + exp32)) < 1e-9
      and abs(sa32['total_inwc'] - sa32['subtotal_inwc'] * 1.10) < 1e-9)
_it32 = _fse_items(sa32)
check("case32 one take-off line 'Fan system effect, supply outlet', SR7-11, C 3.80",
      len(_it32) == 1 and _it32[0][4][0] == 'Fan system effect, supply outlet'
      and _it32[0][4][1].startswith('Fan system effect: SR7-11 pos C, L 0.0 / Le 5.7')
      and abs(_it32[0][4][2] - 3.80) < 1e-9, "items %r" % (_it32,))
check("case32 the index run items still add up to the subtotal",
      abs(_items_sum(sa32) - sa32['subtotal_inwc']) < 1e-9)
check("case32 HIGH_LOSS_INWC is the named 0.5 in. wc threshold",
      FSE.HIGH_LOSS_INWC == 0.5 and not hasattr(FSE, 'HIGH_C'))
check("case32 priced in the total, not listed gray for the plan",
      sa32['fse_inwc'] > 0.0 and not sa32.get('fse_gray_ids'),
      "gray %r" % (sa32.get('fse_gray_ids'),))

# 32b: the same elbow at 2000 fpm: 3.80 x Pv(2000) = 0.948 in. wc, GRAY, with
#      the straight length at which the loss drops under 0.5 in. wc.
d32b = build([(922, 0.0, 10.0, 2000.0, 4.0, 'Supply Air')])
sa32b = _fse_run(n32, c32, d32b, {930: (8000.0, 'Supply Air', 'SD-32')}, 920)
_r32b = FSE.outlet_effect('DWDI', 0.5, 4.0, 8000.0, 0.0, pv_of=pv)
check("case32b 0.948 in. wc is GRAY with its fix length, and listed for the plan",
      _r32b['high'] and _r32b['fix_ft'] is not None
      and any(k == 'gray' and t.startswith('GRAY: System Effect Pressure Loss too high')
              and '%.3f in. wc' % _r32b['loss'] in t
              and 'increase length of duct to %.1f ft' % _r32b['fix_ft'] in t
              for k, t in sa32b.get('fse_notes', ()))
      and sa32b.get('fse_gray_ids') == {921},
      "notes %r gray %r" % (sa32b.get('fse_notes'), sa32b.get('fse_gray_ids')))
_fx = _r32b['fix_ft']
check("case32b at the fix length the loss is under 0.5, one step short it is not",
      FSE.outlet_effect('DWDI', 0.5, 4.0, 8000.0, _fx, pv_of=pv)['loss'] < 0.5
      and FSE.outlet_effect('DWDI', 0.5, 4.0, 8000.0, _fx - 0.1, pv_of=pv)['loss'] >= 0.5,
      "fix %r" % _fx)

# 33: elbow past Le (8 ft > 5.71 ft): outside the outlet window, no system effect.
n33 = {940: _unit({941: SQ24}),
       942: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c33 = {940: [941], 941: [942], 942: [943], 943: [950]}
d33 = build([(941, 0.0, 8.0, 1000.0, 4.0, 'Supply Air'),
             (943, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa33 = _fse_run(n33, c33, d33, {950: (4000.0, 'Supply Air', 'SD-33')}, 940)
check("case33 elbow 8 ft out, past Le 5.71 ft: fse 0, not gray",
      sa33.get('fse_inwc') == 0.0
      and not any(k == 'gray' for k, _t in sa33.get('fse_notes', ())),
      "fse %r notes %r" % (sa33.get('fse_inwc'), sa33.get('fse_notes')))
# at L = Le / 4 (the 0.25 column): A 1.20, B 2.00, C 2.20, D 1.90 -> 2.20
d33b = build([(941, 0.0, LE24 * 0.25, 1000.0, 4.0, 'Supply Air'),
              (943, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa33b = _fse_run(n33, c33, d33b, {950: (4000.0, 'Supply Air', 'SD-33')}, 940)
check("case33 elbow at L/Le 0.25: C 2.20 (SR7-11) x Pv(1000)",
      abs(sa33b.get('fse_inwc', -1) - 2.20 * pv(1000.0)) < 1e-9,
      "got %r" % sa33b.get('fse_inwc'))

# 34: a transition 2 ft out, inside Le, is passed and its own length counted
#     (Colin, 2026-11-02: the curb adapter is part of the straight length).
#     2 ft duct + 0.5 ft transition + 1 ft duct: the elbow is at L 3.5 ft.
class P(object):
    """A connector origin: a point on a line, feet."""
    def __init__(self, x):
        self.x = x

    def DistanceTo(self, other):
        return abs(self.x - other.x)


n34 = {960: _unit({961: SQ24}),
       962: Elem(kind='transition', family_name='Rectangular Transition - Length',
                 is_round=False,
                 tgeom=([(4.0, 'rect', 24.0, 24.0, P(0.0)),
                         (4.0, 'rect', 24.0, 24.0, P(0.5))], 0.0)),
       964: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c34 = {960: [961], 961: [962], 962: [963], 963: [964], 964: [965], 965: [970]}
d34 = build([(961, 0.0, 2.0, 1000.0, 4.0, 'Supply Air'),
             (963, 0.0, 1.0, 1000.0, 4.0, 'Supply Air'),
             (965, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa34 = _fse_run(n34, c34, d34, {970: (4000.0, 'Supply Air', 'SD-34')}, 960)
_r34 = FSE.outlet_effect('DWDI', 0.5, 4.0, 4000.0, 3.5)
check("case34 transition inside Le is passed, its 0.5 ft counted: elbow at L 3.5 ft",
      abs(sa34.get('fse_inwc', -1) - _r34['c'] * pv(1000.0)) < 1e-9
      and _fse_items(sa34) and 'L 3.5' in _fse_items(sa34)[0][4][1]
      and not any(k == 'warn' for k, _t in sa34.get('fse_notes', ())),
      "fse %r items %r notes %r" % (sa34.get('fse_inwc'), _fse_items(sa34),
                                    sa34.get('fse_notes')))
check("case34 the transition is still priced as an ordinary fitting",
      any(i[0] == 'Fitting 962' for i in sa34.get('items', ()))
      and abs(_items_sum(sa34) - sa34['subtotal_inwc']) < 1e-9)
# an unreadable transition length adds 0 (logged): elbow at L 3.0 ft
n34u = dict(n34)
n34u[962] = Elem(kind='transition', family_name='Rectangular Transition - Length',
                 is_round=False)
sa34u = _fse_run(n34u, c34, d34, {970: (4000.0, 'Supply Air', 'SD-34')}, 960)
check("case34 unreadable transition length counts 0: elbow at L 3.0 ft",
      abs(sa34u.get('fse_inwc', -1) -
          FSE.outlet_effect('DWDI', 0.5, 4.0, 4000.0, 3.0)['c'] * pv(1000.0)) < 1e-9,
      "got %r" % sa34u.get('fse_inwc'))
# two transitions back to back (curb adapter then a reducer): both passed
n34t = dict(n34)
n34t[963] = Elem(kind='transition', family_name='Rectangular Transition - Length',
                 is_round=False,
                 tgeom=([(4.0, 'rect', 24.0, 24.0, P(0.0)),
                         (4.0, 'rect', 24.0, 24.0, P(1.0))], 0.0))
d34t = build([(961, 0.0, 2.0, 1000.0, 4.0, 'Supply Air'),
              (965, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa34t = _fse_run(n34t, c34, d34t, {970: (4000.0, 'Supply Air', 'SD-34')}, 960)
check("case34 two transitions in a row are both passed: elbow at L 2 + 0.5 + 1.0 = 3.5 ft",
      abs(sa34t.get('fse_inwc', -1) - _r34['c'] * pv(1000.0)) < 1e-9,
      "got %r notes %r" % (sa34t.get('fse_inwc'), sa34t.get('fse_notes')))
# a tap 3 ft along the first duct is a disruption too
n34b = {980: _unit({981: SQ24}), 981: Elem(kind='duct', tap_pos={982: 3.0}),
        982: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
        983: Elem(kind='duct')}
c34b = {980: [981], 981: [982, 983], 982: [990], 983: [991]}
d34b = build([(981, 0.0, 10.0, 1000.0, 4.0, 'Supply Air'),
              (983, 0.0, 10.0, 750.0, 4.0, 'Supply Air')])
t34b = {990: (1000.0, 'Supply Air', 'SD-a'), 991: (3000.0, 'Supply Air', 'SD-b')}
sa34b = _fse_run(n34b, c34b, d34b, t34b, 980)
check("case34 tap 3 ft along the outlet duct: flagged as a tap, not priced",
      sa34b.get('fse_inwc') == 0.0
      and any(k == 'warn' and t.startswith('WARNING: tap 982, 3.0 ft from the supply outlet')
              for k, t in sa34b.get('fse_notes', ())),
      "notes %r" % (sa34b.get('fse_notes'),))

# 35: round return inlet elbow 1 ft from the unit, r/D 1.0: ED7-2 between
#     L/D 0 (1.20) and 2.0 (0.67) at L/D 0.5 = 1.0675, 0.067 in. wc, not GRAY;
#     inlet swirl is not reported.
n35 = {1000: _unit({1001: RD24}),
       1002: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                  geom={'round': True, 'd_in': 24.0, 'r_in': 24.0, 'w_in': None,
                        'h_in': None, 'angle': 90.0})}
c35 = {1000: [1001], 1001: [1002], 1002: [1003], 1003: [1010]}
d35 = build([(1001, 0.0, 1.0, 1000.0, math.pi, 'Return Air'),
             (1003, 0.0, 10.0, 1000.0, math.pi, 'Return Air')])
ra35 = _fse_run(n35, c35, d35, {1010: (1000.0 * math.pi, 'Return Air', 'RG-35')}, 1000,
                cls='Return Air')
c35v = 1.20 + (0.67 - 1.20) * 0.25
check("case35 return inlet elbow: ED7-2 r/D 1.0, L/D 0.5 -> C %.4f x Pv(1000)" % c35v,
      abs(ra35.get('fse_inwc', -1) - c35v * pv(1000.0)) < 1e-9,
      "got %r" % ra35.get('fse_inwc'))
_n35 = ra35.get('fse_notes', ())
check("case35 inlet swirl is not reported",
      not any('swirl' in t.lower() for k, t in _n35),
      "notes %r" % (_n35,))
check("case35 C 1.07 at 1000 fpm is 0.067 in. wc, under 0.5: not GRAY (Colin, 2026-11-02)",
      not any(k == 'gray' for k, t in _n35),
      "notes %r" % (_n35,))
check("case35 take-off line 'Fan system effect, return inlet' on ED7-2",
      [i[4][0] for i in _fse_items(ra35)] == ['Fan system effect, return inlet']
      and _fse_items(ra35)[0][4][1].startswith('Fan system effect: ED7-2'))

# 36: rect exhaust inlet elbow, ER7-1: H 24 in. (plane of the turn), r 36 in.
#     (r/H 1.5), L 4 ft (L/H 2.0) -> 0.57, not gray; 24 x 12 is not square,
#     so the ER7-1 square ASSUMPTION line is raised.
RC2412 = {'round': False, 'area_ft2': 2.0, 'd_in': None, 'w_in': 24.0, 'h_in': 12.0}
n36 = {1100: _unit({1101: RC2412}),
       1102: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False,
                  geom={'round': False, 'd_in': None, 'r_in': 36.0, 'w_in': 24.0,
                        'h_in': 12.0, 'angle': 90.0})}
c36 = {1100: [1101], 1101: [1102], 1102: [1103], 1103: [1110]}
d36 = build([(1101, 0.0, 4.0, 1000.0, 2.0, 'Exhaust Air'),
             (1103, 0.0, 10.0, 1000.0, 2.0, 'Exhaust Air')])
ea36 = _fse_run(n36, c36, d36, {1110: (2000.0, 'Exhaust Air', 'EG-36')}, 1100,
                cls='Exhaust Air')
check("case36 exhaust inlet rect elbow: ER7-1 r/H 1.5, L/H 2.0 -> 0.57 x Pv(1000)",
      abs(ea36.get('fse_inwc', -1) - 0.57 * pv(1000.0)) < 1e-9
      and not any(k == 'gray' for k, _t in ea36.get('fse_notes', ())),
      "got %r notes %r" % (ea36.get('fse_inwc'), ea36.get('fse_notes')))
check("case36 non-square ER7-1 elbow: the square ASSUMPTION line, not a per-run note",
      ea36.get('fse_assumptions') == [FSE.ER7_1_SQUARE_ASSUMPTION]
      and not any('plane of the turn' in t for _k, t in ea36.get('fse_notes', ())),
      "assume %r" % (ea36.get('fse_assumptions'),))

# 37: inlet elbow geometry unreadable: H = the LARGER side of the unit inlet,
#     radius unknown -> smallest row r/H 0.5; L 0 -> ER7-1 2.50.
n37 = {1100: _unit({1102: RC2412}),
       1102: Elem(kind='elbow', family_name='Rectangular Elbow - Mitered',
                  is_round=False, geom=None)}
c37 = {1100: [1102], 1102: [1103], 1103: [1110]}
d37 = build([(1103, 0.0, 10.0, 1000.0, 2.0, 'Return Air')])
ra37 = _fse_run(n37, c37, d37, {1110: (2000.0, 'Return Air', 'RG-37')}, 1100,
                cls='Return Air')
check("case37 unreadable inlet elbow: larger inlet side, smallest r/H row -> 2.50 x Pv, said so",
      abs(ra37.get('fse_inwc', -1) - 2.50 * pv(1000.0)) < 1e-9
      and any('larger side' in t for _k, t in ra37.get('fse_notes', ()))
      and any('conservative' in t for _k, t in ra37.get('fse_notes', ())),
      "got %r notes %r" % (ra37.get('fse_inwc'), ra37.get('fse_notes')))

# 38: duct pick (Colin, 2026-11-02): the picked duct's unit side end is the
#     unit outlet. Root duct 2000 (2 ft, 24 x 24, 4000 cfm); behind it the unit
#     side stretch (elbow 2001, stub 2002) is not walked; elbow 2003 past the
#     root sits at L 2.0 ft, the root's own length.
n38 = {2000: Elem(kind='duct'),
       2001: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False),
       2003: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c38 = {2000: [2001, 2003], 2001: [2002], 2003: [2004], 2004: [2010]}
d38 = build([(2000, 0.0, 2.0, 1000.0, 4.0, 'Supply Air'),
             (2002, 0.0, 0.5, 0.0, 4.0, 'Supply Air'),
             (2004, 0.0, 10.0, 1000.0, 4.0, 'Supply Air')])
for _dr in d38.values():
    _dr.size = '24x24"'
    _dr.d_h_in = 24.0
sa38 = _fse_run(n38, c38, d38, {2010: (4000.0, 'Supply Air', 'SD-38')}, 2000,
                duct_rooted_ids={2000})
_r38 = FSE.outlet_effect('DWDI', 0.5, 4.0, 4000.0, 2.0)
check("case38 duct pick: elbow 2003 priced at L 2.0 ft from the unit end of duct 2000",
      abs(sa38.get('fse_inwc', -1) - _r38['c'] * pv(1000.0)) < 1e-9
      and [i[0] for i in _fse_items(sa38)] == ['Fan system effect 2003'],
      "fse %r items %r notes %r" % (sa38.get('fse_inwc'), _fse_items(sa38),
                                    sa38.get('fse_notes')))
check("case38 duct pick NOTE names the duct and the outlet",
      ('note', 'NOTE: duct pick, fan system effect measured from the unit end of '
               'duct 2000, taken as the unit outlet.') in sa38.get('fse_notes', ()),
      "notes %r" % (sa38.get('fse_notes'),))
check("case38 the unit side elbow is priced in the unit side stretch, not the system effect",
      sa38.get('unit_side_fittings') == 1
      and abs(_items_sum(sa38) - sa38['subtotal_inwc']) < 1e-9,
      "unit side %r" % (sa38.get('unit_side_fittings'),))
check("case38 the elbow at 1000 fpm is under 0.5 in. wc: no gray id",
      not _r38['high'] and not sa38.get('fse_gray_ids'),
      "gray %r high %r" % (sa38.get('fse_gray_ids'), _r38['high']))
# 38c: the same duct pick at 3000 fpm: the elbow's loss is over 0.5 in. wc and
#      a GRAY elbow on a duct pick reaches the gray id set.
d38c = build([(2000, 0.0, 2.0, 3000.0, 4.0, 'Supply Air'),
              (2002, 0.0, 0.5, 0.0, 4.0, 'Supply Air'),
              (2004, 0.0, 10.0, 3000.0, 4.0, 'Supply Air')])
for _dr in d38c.values():
    _dr.size = '24x24"'
    _dr.d_h_in = 24.0
sa38c = _fse_run(n38, c38, d38c, {2010: (12000.0, 'Supply Air', 'SD-38')}, 2000,
                 duct_rooted_ids={2000})
_r38c = FSE.outlet_effect('DWDI', 0.5, 4.0, 12000.0, 2.0, pv_of=pv)
check("case38c a GRAY elbow on a duct pick reaches the gray id set",
      _r38c['high'] and sa38c.get('fse_gray_ids') == {2003},
      "gray %r high %r loss %r" % (sa38c.get('fse_gray_ids'), _r38c['high'],
                                   _r38c['loss']))

# 38b: duct pick on a return with a mitered inlet elbow 1 ft along: ER7-1 on
#      the r/H 1.0 row (Colin, 2026-11-02), and the elbow is not priced again.
n38b = {2100: Elem(kind='duct'),
        2101: Elem(kind='elbow', family_name='Rectangular Elbow - Mitered', is_round=False,
                   geom={'round': False, 'd_in': None, 'r_in': None, 'w_in': 24.0,
                         'h_in': 24.0, 'angle': 90.0})}
c38b = {2100: [2101], 2101: [2102], 2102: [2110]}
d38b = build([(2100, 0.0, 1.0, 1000.0, 4.0, 'Return Air'),
              (2102, 0.0, 10.0, 1000.0, 4.0, 'Return Air')])
for _dr in d38b.values():
    _dr.size = '24x24"'
ra38b = _fse_run(n38b, c38b, d38b, {2110: (4000.0, 'Return Air', 'RG-38')}, 2100,
                 cls='Return Air', duct_rooted_ids={2100})
_c38b = FSE.inlet_c(False, 1.0, 0.5)[0]
check("case38b mitered inlet elbow on a duct pick: ER7-1 r/H 1.0, L/H 0.5, not repriced",
      abs(ra38b.get('fse_inwc', -1) - _c38b * pv(1000.0)) < 1e-9
      and abs(ra38b.get('fitting_inwc', -1)) < 1e-12,
      "fse %r fit %r expected %r" % (ra38b.get('fse_inwc'), ra38b.get('fitting_inwc'),
                                     _c38b * pv(1000.0)))
check("case38b the mitered NOTE names the elbow and its distance",
      ('note', 'NOTE: elbow 2101, 1.0 ft from the return inlet, is mitered: ER7-1 read '
               'on the r/H 1.0 row (Colin, 2026-11-02).') in ra38b.get('fse_notes', ())
      and any('duct 2100, taken as the unit inlet' in t
              for _k, t in ra38b.get('fse_notes', ())),
      "notes %r" % (ra38b.get('fse_notes'),))

# 38c: duct pick with a tap on the picked duct: the tap is measured from the
#      unit side neighbor (elbow 2201) and flagged by id and distance.
n38c = {2200: Elem(kind='duct', tap_pos={2205: 1.5}),
        2201: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False),
        2205: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
        2203: Elem(kind='duct')}
c38c = {2200: [2201, 2205, 2203], 2201: [2202], 2205: [2210], 2203: [2211]}
d38c = build([(2200, 0.0, 10.0, 1000.0, 4.0, 'Supply Air'),
              (2202, 0.0, 0.5, 0.0, 4.0, 'Supply Air'),
              (2203, 0.0, 10.0, 750.0, 4.0, 'Supply Air')])
for _dr in d38c.values():
    _dr.size = '24x24"'
sa38c = _fse_run(n38c, c38c, d38c, {2210: (1000.0, 'Supply Air', 'SD-a'),
                                    2211: (3000.0, 'Supply Air', 'SD-b')}, 2200,
                 duct_rooted_ids={2200})
check("case38c tap on the picked duct measured from its unit side neighbor, flagged by id",
      getattr(n38c[2200], 'tap_from', None) == 2201
      and any(k == 'warn' and t.startswith('WARNING: tap 2205, 1.5 ft from the supply outlet')
              for k, t in sa38c.get('fse_notes', ())),
      "from %r notes %r" % (getattr(n38c[2200], 'tap_from', None), sa38c.get('fse_notes')))
# F2 (2026-11-02): external static orders the same taps from the same end,
# so the picked duct's friction is stepped, not charged whole with a NOTE.
check("case38c external static orders the picked duct's taps (no 'unreadable' NOTE)",
      not any('tap positions on duct 2200 unreadable' in u
              for u in sa38c.get('unpriced', ())),
      "unpriced %r" % (sa38c.get('unpriced'),))

# 39: SWSI switches the outlet tables (Ab/Ao 0.5, L 0: SR7-7 3.80); off = none.
sa39 = _fse_run(n32, c32, d32, t32, 920,
                fan={'on': True, 'fan_type': 'SWSI', 'blast_ratio': 0.5, 'velocity_basis': 'connection'})
check("case39 SWSI reads the SR7-5 to SR7-8 tables",
      bool(_fse_items(sa39)) and _fse_items(sa39)[0][4][1].startswith(
          'Fan system effect: SR7-7'),
      "items %r" % (_fse_items(sa39),))
sa39b = _fse_run(n32, c32, d32, t32, 920,
                 fan={'on': False, 'fan_type': 'DWDI', 'blast_ratio': 0.5, 'velocity_basis': 'connection'})
# Colin, 2026-11-03 (Q4): unticked, nothing is priced, but the straight duct
# check still runs: the elbow on the outlet is RED, nothing else is said.
check("case39 unticked: nothing priced, only the straight duct RED line",
      sa39b.get('fse_inwc') == 0.0 and not _fse_items(sa39b)
      and [k for k, _t in sa39b.get('fse_notes', ())] == ['red']
      and sa39b.get('fse_red_ids') == {921} and not sa39b.get('fse_gray_ids'),
      "notes %r" % (sa39b.get('fse_notes'),))

# 40: unit connection unreadable: the first duct's area is used, and said.
n40 = dict(n32)
n40[920] = _unit({})
sa40 = _fse_run(n40, c32, d32, t32, 920)
check("case40 unreadable unit outlet: first duct area used (same 3.80 x Pv here), noted",
      abs(sa40.get('fse_inwc', -1) - exp32) < 1e-9
      and any('connection size unreadable' in t for _k, t in sa40.get('fse_notes', ())),
      "got %r notes %r" % (sa40.get('fse_inwc'), sa40.get('fse_notes')))


# ── case 41 to 45: Colin's rulings on the fan system effect (2026-10-29) ─────

# 41: a fitting right on the unit connection was priced at 0 fpm (the walk
#     started with no velocity). It now starts at the first duct's cfm over
#     the unit connector area: 4000 cfm / 4 ft2 = 1000 fpm. Fan effect off.
sa41 = _fse_run(n32, c32, d32, t32, 920, fan=None)
_c41 = fitting_tables.elbow_c(False)
check("case41 elbow on the unit connection is priced (worksheet C x Pv(1000)), not 0",
      sa41['fitting_inwc'] > 0.0
      and abs(sa41['fitting_inwc'] - _c41 * pv(1000.0)) < 1e-9,
      "got %r expected %.6f" % (sa41['fitting_inwc'], _c41 * pv(1000.0)))
# the start velocity uses the unit connector area, not the first duct's:
# a 2 ft2 connector on the same 4000 cfm is 2000 fpm at the elbow
n41b = dict(n32)
n41b[920] = _unit({921: dict(SQ24, area_ft2=2.0)})
sa41b = _fse_run(n41b, c32, d32, t32, 920, fan=None)
check("case41 start velocity is the connector's (4000 cfm / 2 ft2 = 2000 fpm)",
      abs(sa41b['fitting_inwc'] - _c41 * pv(2000.0)) < 1e-9,
      "got %r" % sa41b['fitting_inwc'])
# and it is priced with the first duct's class: a return elbow on the unit
d41c = build([(922, 0.0, 10.0, 1000.0, 4.0, 'Return Air')])
ra41 = _fse_run(n32, c32, d41c, {930: (4000.0, 'Return Air', 'RG-41')}, 920, fan=None,
                cls='Return Air')
check("case41 return elbow on the unit priced at Pv(1000) too",
      abs(ra41.get('fitting_inwc', -1) - _c41 * pv(1000.0)) < 1e-9,
      "got %r" % ra41.get('fitting_inwc'))

# 42: inlet elbow priced by ED7-2 is NOT also priced as a fitting (ED7-2 / ER7-1
#     include the elbow): one take-off line, and the run still adds up.
check("case42 inlet elbow: no fitting loss on top of ED7-2 (fitting_inwc 0)",
      abs(ra35.get('fitting_inwc', -1)) < 1e-12, "got %r" % ra35.get('fitting_inwc'))
# Since 2026-11-04 the elbow is one of the group charged once with the fan
# system effect (the larger): listed, at 0 when the fan system effect wins.
_el42 = [i for i in ra35.get('items', ()) if i[0] == 'Fitting 1002']
check("case42 the take-off lists the elbow at 0, inside the fan system effect line",
      any(i[4] and i[4][0] == 'Fan system effect, return inlet' for i in ra35.get('items', ()))
      and len(_el42) == 1 and _el42[0][3] == 0.0
      and _el42[0][1].endswith('in the fan system effect'), "elbow %r" % (_el42,))
check("case42 items add up to the subtotal",
      abs(_items_sum(ra35) - ra35['subtotal_inwc']) < 1e-9)
check("case42 the report says the group is charged once",
      any('elbow 1002' in t and 'Fan system effect charged, the straight duct' in t
          for _k, t in ra35.get('fse_gov_notes', ())), "gov %r" % (ra35.get('fse_gov_notes'),))
check("case42 outlet: the elbow on the unit is in the group too (fitting 0, fan system effect wins)",
      abs(sa32['fitting_inwc']) < 1e-12 and sa32.get('fse_inwc', 0) > 0.0,
      "fit %r fse %r" % (sa32['fitting_inwc'], sa32.get('fse_inwc')))

# 43: inlet window is the table floor distance, 5 D (10 ft on a 24 in. inlet).
def _inlet_net(l_ft):
    n = {1300: _unit({1301: RD24}),
         1302: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                    geom={'round': True, 'd_in': 24.0, 'r_in': 24.0, 'w_in': None,
                          'h_in': None, 'angle': 90.0})}
    ch = {1300: [1301], 1301: [1302], 1302: [1303], 1303: [1310]}
    d = build([(1301, 0.0, l_ft, 1000.0, math.pi, 'Return Air'),
               (1303, 0.0, 10.0, 1000.0, math.pi, 'Return Air')])
    return _fse_run(n, ch, d, {1310: (1000.0 * math.pi, 'Return Air', 'RG-43')}, 1300,
                    cls='Return Air')
ra43a = _inlet_net(9.0)          # L/D 4.5: 0.67 - (0.67 - 0.33) x 2.5 / 3
c43 = 0.67 - (0.67 - 0.33) * 2.5 / 3.0
check("case43 inlet elbow at 9 ft (inside 5 D = 10 ft): ED7-2 %.4f x Pv, elbow not repriced" % c43,
      abs(ra43a.get('fse_inwc', -1) - c43 * pv(1000.0)) < 1e-9
      and abs(ra43a.get('fitting_inwc', -1)) < 1e-12,
      "fse %r fit %r" % (ra43a.get('fse_inwc'), ra43a.get('fitting_inwc')))
ra43b = _inlet_net(11.0)         # past 5 D: an ordinary elbow
check("case43 inlet elbow at 11 ft (past 5 D): no system effect, no swirl, priced as a fitting",
      ra43b.get('fse_inwc') == 0.0 and ra43b.get('fitting_inwc', 0) > 0.0
      and not any('FSE1,sw' in t for _k, t in ra43b.get('fse_notes', ())),
      "fse %r fit %r notes %r" % (ra43b.get('fse_inwc'), ra43b.get('fitting_inwc'),
                                  ra43b.get('fse_notes')))

# 44: outlet window is Le, whatever its length: 60 x 60 in. (25 ft2) at
#     3000 fpm (15.24 m/s, Eq. 36): Le = 15.24 x 1524 / 4500 m = 16.93 ft, so
#     an elbow 16 ft out (past the old fixed 15 ft) is still priced.
SQ60 = {'round': False, 'area_ft2': 25.0, 'd_in': None, 'w_in': 60.0, 'h_in': 60.0}
le44 = 15.24 * 1524.0 / 4500.0 / 0.3048
n44 = {1400: _unit({1401: SQ60}),
       1402: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c44 = {1400: [1401], 1401: [1402], 1402: [1403], 1403: [1410]}
d44 = build([(1401, 0.0, 16.0, 3000.0, 25.0, 'Supply Air'),
             (1403, 0.0, 5.0, 3000.0, 25.0, 'Supply Air')])
sa44 = _fse_run(n44, c44, d44, {1410: (75000.0, 'Supply Air', 'SD-44')}, 1400)
_r44 = 16.0 / le44                # L/Le between the 0.50 and 1.00 columns
c44v = 1.00 * (1.0 - (_r44 - 0.5) / 0.5)   # DWDI Ab/Ao 0.5: worst at 0.50 is C 1.00
# Inside Le the elbow is priced C 0.110 x Pv(3000), but as a fitting it
# prices higher, so the group wins (Colin, 2026-11-04): the fan system effect
# line stays listed at 0, the elbow is charged as a fitting.
_f44 = [i for i in sa44.get('items', ()) if i[4] and i[4][0].startswith('Fan system effect')]
_e44 = [i for i in sa44.get('items', ()) if i[0] == 'Fitting 1402']
check("case44 outlet elbow 16 ft out, inside Le %.2f ft: C %.3f x Pv(3000) loses to the elbow, group charged" % (le44, c44v),
      len(_f44) == 1 and _f44[0][3] > c44v * pv(3000.0)
      and _f44[0][4][1].startswith('Duct losses: straight duct and elbow')
      and _f44[0][4][2] is None
      and abs(sa44.get('fse_inwc', 0) - _f44[0][3]) < 1e-12
      and _e44 and _e44[0][3] == 0.0 and _e44[0][4][1].endswith('in the fan system effect')
      and abs(sa44['fitting_inwc']) < 1e-12
      and abs(_items_sum(sa44) - sa44['subtotal_inwc']) < 1e-9,
      "fse %r items %r" % (sa44.get('fse_inwc'), _f44 + _e44))
# a transition past Le is not a disruption at all (8 ft out, Le 5.71 ft)
n44b = dict(n34)
d44b = build([(961, 0.0, 8.0, 1000.0, 4.0, 'Supply Air'),
              (963, 0.0, 1.0, 1000.0, 4.0, 'Supply Air'),
              (965, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa44b = _fse_run(n44b, c34, d44b, {970: (4000.0, 'Supply Air', 'SD-34')}, 960)
check("case44 transition past Le: not flagged",
      not any('transition' in t for _k, t in sa44b.get('fse_notes', ())),
      "notes %r" % (sa44b.get('fse_notes'),))

# 45: the DWDI assumption is a report line (method_lines), not a per-run note.
check("case45 no per-run SWSI note on the DWDI inlet run",
      not any('SWSI' in t for _k, t in ra35.get('fse_notes', ())))


# ── case 46 to 50: review round 2 (2026-10-29) ───────────────────────────────

# 46: a tee right on the unit outlet, 4000 cfm unit, two 2000 cfm legs.
#     The connection carries 4000 cfm (both legs), so 1000 fpm over the 4 ft2
#     connector, not 500 fpm off one leg.
n46 = {1500: _unit({1501: SQ24}),
       1501: Elem(kind='unknown_fitting', family_name='Rectangular Tee', is_round=False),
       1502: Elem(kind='duct'), 1503: Elem(kind='duct')}
c46 = {1500: [1501], 1501: [1502, 1503], 1502: [1510], 1503: [1511]}
d46 = build([(1502, 0.0, 10.0, 1000.0, 2.0, 'Supply Air'),
             (1503, 0.0, 10.0, 1000.0, 2.0, 'Supply Air')])
t46 = {1510: (2000.0, 'Supply Air', 'SD-a'), 1511: (2000.0, 'Supply Air', 'SD-b')}
sa46 = _fse_run(n46, c46, d46, t46, 1500, fan=None)
_tee46 = [i for i in sa46.get('items', ()) if i[0] == 'Fitting 1501']
check("case46 tee on the unit priced at the connection's 1000 fpm (both legs), not 500",
      len(_tee46) == 1 and abs(_tee46[0][2] - 1000.0) < 1e-9,
      "items %r" % (_tee46,))
# an elbow on the unit with the tee after it: FSE2 at Pv(1000), not Pv(500)
n46b = {1500: _unit({1520: SQ24}),
        1520: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False),
        1501: Elem(kind='unknown_fitting', family_name='Rectangular Tee', is_round=False),
        1502: Elem(kind='duct'), 1503: Elem(kind='duct')}
c46b = {1500: [1520], 1520: [1501], 1501: [1502, 1503], 1502: [1510], 1503: [1511]}
sa46b = _fse_run(n46b, c46b, d46, t46, 1500)
check("case46 elbow on the unit before a tee: FSE2 = 3.80 x Pv(1000) on the full 4000 cfm",
      abs(sa46b.get('fse_inwc', -1) - 3.80 * pv(1000.0)) < 1e-9,
      "got %r expected %.6f" % (sa46b.get('fse_inwc'), 3.80 * pv(1000.0)))

# 47: two supply outlets. The worst path leaves through the plain one (2.0 in.
#     wc of friction), but the other outlet's elbow (C 3.80 at 2000 fpm, GRAY) is still
#     reported, once, named by its element.
n47 = {1600: _unit({1601: SQ24, 1611: SQ24}),
       1601: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c47 = {1600: [1601, 1611], 1601: [1602], 1602: [1620], 1611: [1621]}
d47 = build([(1602, 0.0, 10.0, 2000.0, 4.0, 'Supply Air'),
             (1611, 2.0, 20.0, 1000.0, 4.0, 'Supply Air')])
t47 = {1620: (8000.0, 'Supply Air', 'SD-elbow'), 1621: (4000.0, 'Supply Air', 'SD-long')}
sa47 = _fse_run(n47, c47, d47, t47, 1600)
check("case47 the worst path is the plain outlet: no FSE in its total",
      sa47.get('terminal_name') == 'SD-long' and sa47.get('fse_inwc') == 0.0,
      "terminal %r fse %r" % (sa47.get('terminal_name'), sa47.get('fse_inwc')))
_g47 = [t for k, t in sa47.get('fse_notes', ()) if k == 'gray']
check("case47 the other outlet's GRAY is still reported, once, named by element",
      len(_g47) == 1 and 'supply outlet at element 1601' in _g47[0],
      "notes %r" % (sa47.get('fse_notes'),))

# 48: a unit connection that feeds equipment directly is noted, not skipped silently.
n48 = {1700: _unit({}), 1701: Elem(kind='equipment')}
c48 = {1700: [1701], 1701: [1702], 1702: [1710]}
d48 = build([(1702, 0.0, 10.0, 800.0, 1.0, 'Supply Air')])
sa48 = _fse_run(n48, c48, d48, {1710: (800.0, 'Supply Air', 'SD-48')}, 1700)
check("case48 connection feeding equipment directly: noted as not checked",
      any(k == 'note' and 'feeds equipment directly' in t and '1701' in t
          for k, t in sa48.get('fse_notes', ())) and sa48.get('fse_inwc') == 0.0,
      "notes %r" % (sa48.get('fse_notes'),))

# 49: duct pick with nothing on its unit side and a tap on it: the taps
#     cannot be put in order from the unit end, and the report says so.
n49 = {1800: Elem(kind='duct', tap_pos={1805: 4.0}),
       1805: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False)}
c49 = {1800: [1805, 1801], 1805: [1810], 1801: [1811]}
d49 = build([(1800, 0.1, 10.0, 800.0, 1.0, 'Supply Air'),
             (1801, 0.0, 10.0, 600.0, 1.0, 'Supply Air')])
for _dr in d49.values():
    _dr.size = '12x12"'
sa49 = _fse_run(n49, c49, d49, {1810: (200.0, 'Supply Air', 'SD-49'),
                                1811: (600.0, 'Supply Air', 'SD-49b')}, 1800,
                duct_rooted_ids={1800})
check("case49 duct pick, no unit side: taps unordered, NOTE names the duct",
      any(k == 'note' and t.startswith('NOTE: duct pick, nothing found on the unit side '
                                       'of duct 1800')
          for k, t in sa49.get('fse_notes', ())),
      "notes %r" % (sa49.get('fse_notes'),))

# 50: a VAV box picked as a root: no system effect on boxes (Colin), noted.
n50 = {1900: _unit({1901: SQ24}),
       1901: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c50 = {1900: [1901], 1901: [1902], 1902: [1910]}
d50 = build([(1902, 0.0, 10.0, 1000.0, 4.0, 'Supply Air')])
sa50 = _fse_run(n50, c50, d50, {1910: (4000.0, 'Supply Air', 'SD-50')}, 1900,
                vav_box_ids={1900})
check("case50 VAV box root: fse 0 even with an elbow on it, 'VAV box' noted",
      sa50.get('fse_inwc') == 0.0
      and any('VAV box' in t for _k, t in sa50.get('fse_notes', ())),
      "notes %r" % (sa50.get('fse_notes'),))

# 50b: VAV Box scope (Colin, 2026-11-09): the picked box root is walked as a
#      box. External static is charged from the box outlet (the elbow on it
#      priced as a fitting), with NO fan checks of any kind: no fan system
#      effect, no straight duct or swirl RED, no damper window WARNING.
#      Case 32's graph: as a unit it prices FSE and draws the elbow RED.
check("case50b control: case 32 as a unit prices FSE and flags the elbow RED",
      sa32.get('fse_inwc', 0.0) > 0.0 and 921 in set(sa32.get('fse_red_ids', ())),
      "fse %r red %r" % (sa32.get('fse_inwc'), sa32.get('fse_red_ids')))
for _lbl50b, _fan50b in (('FSE ticked', FSE_ON),
                         ('checks only walk', dict(FSE_ON, on=False))):
    sa50b = _fse_run(n32, c32, d32, t32, 920, fan=_fan50b, vav_box_ids={920})
    check("case50b box root, %s: ESP charged, elbow priced, fse 0, no FSE row" % _lbl50b,
          sa50b.get('total_inwc', 0.0) > 0.0
          and sa50b.get('fse_inwc') == 0.0 and not _fse_items(sa50b)
          and any(i[0] == 'Fitting 921' for i in sa50b.get('items', ()))
          and abs(_items_sum(sa50b) - sa50b['subtotal_inwc']) < 1e-9,
          "leaf %r" % (sa50b,))
    check("case50b box root, %s: no straight duct, swirl or damper window flag" % _lbl50b,
          not sa50b.get('fse_red_ids') and not sa50b.get('fse_red_ducts')
          and not sa50b.get('fse_gray_ids')
          and not any(k in ('red', 'warn', 'gray')
                      for k, _t in sa50b.get('fse_notes', ())),
          "notes %r red %r" % (sa50b.get('fse_notes'), sa50b.get('fse_red_ids')))
    check("case50b box root, %s: never charged its own box drop" % _lbl50b,
          not any(i[0].startswith('VAV box') for i in sa50b.get('items', ())),
          "items %r" % (sa50b.get('items'),))

# 51: friction along an unbroken main with taps (Colin, 2026-11-02). The main
#     keeps one size but its flow drops at each tap, so its friction is the
#     sum of its pieces at their own flow, and a run leaving through a tap is
#     charged only up to that tap. Case 26's main (10 ft, 1.0 ft2, 1000 cfm,
#     taps at 2 ft taking 300 cfm and 5 ft taking 400 cfm), plus the main
#     carrying on past the taps with 300 cfm. Friction faked as V / 10000 in.
#     wc per 100 ft so each piece's flow shows in the number.
_flat_fric = hvac_graph.duct_friction_loss_per_100ft
hvac_graph.duct_friction_loss_per_100ft = lambda v, d_h=None, eps=None: v / 10000.0
ch51 = {299: [300], 300: [301, 302, 313], 301: [311], 302: [312], 311: [321],
        312: [322], 313: [323]}
n51 = dict(n26)
n51[313] = Elem(kind='duct')
t51 = dict(t26)
t51[323] = (300.0, 'Supply Air', 'SD-3')
def _d51(f311, f312, f313):
    return build([(300, 0.0, 10.0, 1000.0, 1.0, 'Supply Air'),
                  (311, f311, 5.0, 600.0, 0.5, 'Supply Air'),
                  (312, f312, 5.0, 800.0, 0.5, 'Supply Air'),
                  (313, f313, 5.0, 300.0, 1.0, 'Supply Air')])
far51 = _sa(crit([299], ch51, _d51(0.0, 0.5, 0.0), t51, n51), 299)
near51 = _sa(crit([299], ch51, _d51(0.5, 0.0, 0.0), t51, n51), 299)
end51 = _sa(crit([299], ch51, _d51(0.0, 0.0, 1.0), t51, n51), 299)
hvac_graph.duct_friction_loss_per_100ft = _flat_fric
exp_far = 2.0 / 100 * 0.1 + 3.0 / 100 * 0.07           # 0 to 2 ft at 1000, 2 to 5 at 700
exp_near = 2.0 / 100 * 0.1                              # 0 to 2 ft at 1000 only
exp_end = exp_far + 5.0 / 100 * 0.03                    # and 5 to 10 ft at 300
check("case51 run out the far tap: main charged 0-2 ft at 1000 fpm + 2-5 ft at 700 fpm",
      far51.get('terminal_id') == 322
      and abs(far51.get('friction_inwc', -1) - (0.5 + exp_far)) < 1e-9,
      "term %r fric %r expected %r" % (far51.get('terminal_id'), far51.get('friction_inwc'),
                                       0.5 + exp_far))
check("case51 run out the near tap: main charged only 0-2 ft at 1000 fpm",
      near51.get('terminal_id') == 321
      and abs(near51.get('friction_inwc', -1) - (0.5 + exp_near)) < 1e-9,
      "term %r fric %r" % (near51.get('terminal_id'), near51.get('friction_inwc')))
check("case51 run carrying on along the main: all three pieces, 1000 / 700 / 300 fpm",
      end51.get('terminal_id') == 323
      and abs(end51.get('friction_inwc', -1) - (1.0 + exp_end)) < 1e-9,
      "term %r fric %r expected %r" % (end51.get('terminal_id'), end51.get('friction_inwc'),
                                       1.0 + exp_end))
_it51 = [i for i in far51.get('items', ()) if i[0] == 'Duct 300']
check("case51 the take-off line for the main shows 5.0 of 10.0 ft and the stepped loss",
      len(_it51) == 1 and '5.0 of 10.0 ft' in _it51[0][1]
      and abs(_it51[0][3] - exp_far) < 1e-9 and abs(_it51[0][4][3] - 5.0) < 1e-9,
      "items %r" % (_it51,))
check("case51 the index run items still add up to the subtotal",
      abs(sum(i[3] or 0.0 for i in far51.get('items', ())) - far51['subtotal_inwc']) < 1e-9)

# 52: a damper on a MAIN does not cancel a branch's assumed balancing damper
#     (Colin, 2026-11-02: dampers on mains are fire/smoke dampers). Unit 3000
#     -> main 3001 -> balancing damper 3002 -> main 3003 with tap 3004 -> branch
#     3006 -> T3010 (the index run, 0.5 in. wc on the branch); the main carries
#     on as 3005 -> T3011.
def _d52(branch_damper, main_damper=True):
    n = {3000: Elem(kind='equipment'), 3003: Elem(kind='duct', tap_pos={3004: 2.0}),
         3004: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
         3005: Elem(kind='duct'), 3006: Elem(kind='duct'),
         3002: Elem(kind='accessory', family_name='Balancing Damper - Rectangular'),
         3007: Elem(kind='accessory', family_name='Balancing Damper - Rectangular')}
    ch = {3000: [3001], 3001: [3002] if main_damper else [3003], 3002: [3003],
          3003: [3004, 3005], 3005: [3011],
          3004: [3007] if branch_damper else [3006], 3007: [3006], 3006: [3010]}
    d = build([(3001, 0.0, 5.0, 800.0, 1.0, 'Supply Air'),
               (3003, 0.0, 10.0, 800.0, 1.0, 'Supply Air'),
               (3005, 0.0, 10.0, 600.0, 1.0, 'Supply Air'),
               (3006, 0.5, 10.0, 400.0, 0.5, 'Supply Air')])
    t = {3010: (200.0, 'Supply Air', 'SD-52'), 3011: (600.0, 'Supply Air', 'SD-52b')}
    return crit([3000], ch, d, t, n)[3000]['Supply Air']
sa52 = _d52(branch_damper=False)
check("case52 main damper only: the branch still pays the assumed damper (2 dampers)",
      sa52['terminal_id'] == 3010
      and list(sa52['component_keys']).count('damper') == 2
      and any(i[0] == 'Accessory (assumed)' for i in sa52['items']),
      "term %r keys %r" % (sa52['terminal_id'], sa52['component_keys']))
sa52b = _d52(branch_damper=True)
check("case52 a damper drawn on the branch replaces the assumed one (main + branch = 2)",
      list(sa52b['component_keys']).count('damper') == 2
      and not any(i[0] == 'Accessory (assumed)' for i in sa52b['items']),
      "keys %r" % (sa52b['component_keys'],))
sa52c = _d52(branch_damper=True, main_damper=False)
check("case52 branch damper alone: charged once, nothing assumed, items trace the total",
      list(sa52c['component_keys']).count('damper') == 1
      and abs(_items_sum(sa52c) - sa52c['subtotal_inwc']) < 1e-9,
      "keys %r" % (sa52c['component_keys'],))

# 53: a fitting past the last tap on an unbroken main is priced at the flow
#     the taps left (Colin, 2026-11-02), the same flow _seg_friction ends on.
#     Case 51's main (1000 cfm, 1.0 ft2, taps taking 300 and 400 cfm), carrying
#     on through elbow 330 to duct 313 (the index run): the elbow sees
#     (1000 - 300 - 400) / 1.0 = 300 fpm, not the unit end 1000 fpm.
ch53 = {299: [300], 300: [301, 302, 330], 301: [311], 302: [312], 311: [321],
        312: [322], 330: [313], 313: [323]}
n53 = dict(n51)
n53[330] = Elem(kind='elbow', family_name='Round Elbow - Fixed L1', is_round=True)
end53 = _sa(crit([299], ch53, _d51(0.0, 0.0, 1.0), t51, n53), 299)
_it53 = [i for i in end53.get('items', ()) if i[0] == 'Fitting 330']
check("case53 the elbow past the last tap is priced at the remaining 300 fpm",
      end53.get('terminal_id') == 323 and len(_it53) == 1
      and abs(_it53[0][2] - 300.0) < 1e-9,
      "term %r items %r" % (end53.get('terminal_id'), _it53))
# The same run with the tap positions unreadable: friction stays whole at the
# unit end cfm, the elbow at 1000 fpm, and the run says so once.
n53u = dict(n53)
n53u[300] = Elem(kind='duct')
u53 = _sa(crit([299], ch53, _d51(0.0, 0.0, 1.0), t51, n53u), 299)
_note53 = [u for u in u53.get('unpriced', ()) if 'tap positions on duct 300 unreadable' in u]
check("case53 unreadable tap positions: one NOTE on the run, elbow at the unit end 1000 fpm",
      len(_note53) == 1
      and [i[2] for i in u53.get('items', ()) if i[0] == 'Fitting 330'] == [1000.0],
      "unpriced %r" % (u53.get('unpriced'),))

# ── case 54: ceiling exhaust fan (Colin, 2026-11-02) ─────────────────────────
# A fan whose only duct connection is Exhaust Air, out: 4 in. discharge,
# 0.5 ft to an elbow, 20 ft to the open end, which is the wall or roof cap
# carrying the entered 80 cfm (917 fpm). Its ducts are Exhaust Air but move
# like supply (flow_class), so the system effect is at the fan OUTLET (SR7,
# Le 2.5 D = 0.83 ft at under 2500 fpm, the elbow inside it) and the run
# ends in the cap drop, not a diffuser, with no assumed balancing damper.
RD4 = {'round': True, 'area_ft2': math.pi * (2.0 / 12.0) ** 2, 'd_in': 4.0,
       'w_in': None, 'h_in': None}
_v54 = 80.0 / RD4['area_ft2']


def _fan54(discharge):
    n = {2000: _unit({2001: RD4}),
         2002: Elem(kind='elbow', family_name='Round Elbow', is_round=True)}
    c = {2000: [2001], 2001: [2002], 2002: [2003], 2003: []}
    d = build([(2001, 0.002, 0.5, _v54, RD4['area_ft2'], 'Exhaust Air'),
               (2003, 0.080, 20.0, _v54, RD4['area_ft2'], 'Exhaust Air')])
    if discharge:
        for dr in d.values():
            dr.flow_class = 'Supply Air'
    t = {2003: (80.0, 'Exhaust Air', 'Wall / roof cap')}
    r = crit([2000], c, d, t, n, 10.0, fan_effect=FSE_ON,
             discharge_roots={2000: 2003} if discharge else None)
    return r.get(2000, {}).get('Exhaust Air', {})


ea54 = _fan54(True)
_it54 = [i[4][0] for i in ea54.get('items', ()) if i[4]]
check("case54 ceiling fan: one Exhaust Air index run ending at the cap",
      ea54.get('terminal_id') == 2003, "leaf %r" % (ea54,))
check("case54 the run pays the cap drop, no diffuser and no assumed damper",
      'cap' in ea54.get('component_keys', ())
      and 'diffuser' not in ea54.get('component_keys', ())
      and 'damper' not in ea54.get('component_keys', ())
      and 'Wall / roof cap' in _it54, "keys %r rows %r" % (
          ea54.get('component_keys'), _it54))
check("case54 component drop is the cap value",
      abs(ea54.get('component_inwc', -1) - fitting_tables.component_of('cap')) < 1e-12)
check("case54 fan system effect at the fan OUTLET (SR7), not the inlet tables",
      'Fan system effect, exhaust outlet' in _it54
      and any(i[4][1].startswith(('Fan system effect: SR7-', 'Duct losses: '))
              and 'Le' in i[4][1] for i in _fse_items(ea54)),
      "rows %r" % (_it54,))
check("case54 the outlet elbow is charged once with the fan system effect (the larger)",
      (ea54.get('fse_inwc', 0.0) > 0.0) != any(
          i[4] and i[4][0].startswith('Fan system effect') and 'not charged' in i[1]
          for i in ea54.get('items', ())),
      "fse %r items %r" % (ea54.get('fse_inwc'), ea54.get('items')))
check("case54 items add up to the subtotal",
      abs(_items_sum(ea54) - ea54['subtotal_inwc']) < 1e-9)
inl54 = _fan54(False)
check("case54 the same run as an ordinary exhaust reads the INLET tables (ED7-2)",
      any(i[4][0] == 'Fan system effect, exhaust inlet' for i in _fse_items(inl54))
      and 'Wall / roof cap' not in [i[4][0] for i in inl54.get('items', ()) if i[4]],
      "rows %r" % ([i[4][0] for i in inl54.get('items', ()) if i[4]],))

# 54b: a transition on an exhaust discharge (Colin, 2026-11-09, TB-M10C
#      3696775): priced on the supply (SD4) tables through flow_class, but its
#      NOTE names the duct's own class, Exhaust Air, never Supply Air.
def _fan54b():
    n = {2100: _unit({2101: SQ24}),
         2102: Elem(kind='transition', family_name='Rectangular to Round Transition',
                    is_round=False,
                    tgeom=([(1.0, 'rect', 12.0, 12.0, None),
                            (math.pi * 0.25, 'round', 12.0, 12.0, None)], 0.0))}
    c = {2100: [2101], 2101: [2102], 2102: [2103], 2103: []}
    d = build([(2101, 0.01, 20.0, 400.0, 1.0, 'Exhaust Air'),
               (2103, 0.01, 20.0, 400.0 / (math.pi * 0.25), math.pi * 0.25,
                'Exhaust Air')])
    for dr in d.values():
        dr.flow_class = 'Supply Air'
    t = {2103: (400.0, 'Exhaust Air', 'Wall / roof cap')}
    r = crit([2100], c, d, t, n, 10.0, fan_effect=FSE_ON,
             discharge_roots={2100: 2103})
    return r.get(2100, {}).get('Exhaust Air', {})


ea54b = _fan54b()
_txt54b = ' | '.join([t for _k, t in ea54b.get('fse_notes', ())]
                     + [str(i) for i in ea54b.get('items', ())]
                     + list(ea54b.get('unpriced', ())))
check("case54b discharge transition priced on SD4 (supply tables)",
      'SD4-' in _txt54b, _txt54b)
check("case54b its printed system is Exhaust Air, not Supply Air",
      'transition 2102 (Exhaust Air' in _txt54b
      and 'transition 2102 (Supply Air' not in _txt54b, _txt54b)

# ── case 55: Include exhaust roof cap (Colin, 2026-11-02) ──────────────────
# An ordinary exhaust run (grille to fan) pays the cap drop once when ticked;
# supply runs never do, and a ceiling fan's own cap is not charged twice.
n55 = {2100: _unit({})}
c55 = {2100: [2101, 2102], 2101: [2110], 2102: [2120]}
d55 = build([(2101, 0.02, 10.0, 600.0, 1.0, 'Exhaust Air'),
             (2102, 0.02, 10.0, 600.0, 1.0, 'Supply Air')])
t55 = {2110: (600.0, 'Exhaust Air', 'EG-1'), 2120: (600.0, 'Supply Air', 'SD-1')}
off55 = crit([2100], c55, d55, t55, n55, 0.0)[2100]
on55 = crit([2100], c55, d55, t55, n55, 0.0, exhaust_cap=True)[2100]
_cap = fitting_tables.component_of('cap')
check("case55 ticked: the exhaust run pays the cap drop once",
      abs(on55['Exhaust Air']['component_inwc'] - off55['Exhaust Air']['component_inwc'] - _cap) < 1e-12
      and list(on55['Exhaust Air']['component_keys']).count('cap') == 1)
check("case55 ticked: the supply run is unchanged",
      abs(on55['Supply Air']['total_inwc'] - off55['Supply Air']['total_inwc']) < 1e-12)
check("case55 ticked: the exhaust take-off still adds up",
      abs(_items_sum(on55['Exhaust Air']) - on55['Exhaust Air']['subtotal_inwc']) < 1e-9)
check("case55 a ceiling fan run is not charged a second cap when ticked",
      list(crit([2000], {2000: [2001], 2001: [2002], 2002: [2003], 2003: []},
                dict((k, v) for k, v in build([(2001, 0.002, 0.5, _v54, RD4['area_ft2'], 'Exhaust Air'),
                                               (2003, 0.080, 20.0, _v54, RD4['area_ft2'], 'Exhaust Air')]).items()),
                {2003: (80.0, 'Exhaust Air', 'Wall / roof cap')},
                {2000: _unit({2001: RD4}), 2002: Elem(kind='elbow', family_name='Round Elbow', is_round=True)},
                0.0, discharge_roots={2000: 2003}, exhaust_cap=True)[2000]['Exhaust Air']['component_keys']).count('cap') == 1)

# ── case 56: elbows turning in different planes within 6 Dh (Colin, 2026-11-02)
# EA 61 at The Hub Boulder: CUE-060 6x6 inlet, 0.394 ft of duct, a mitered 90
# turning out of the vertical (normal +Y), 0.473 ft, then the riser top
# offset, two 45s turning in the other plane (normal +X), 1.157 ft apart.
# 6 Dh = 3.0 ft: the 90 (0.39 ft) and the first 45 (0.87 ft) are the pair.
SQ6 = {'round': False, 'area_ft2': 0.25, 'd_in': None, 'w_in': 6.0, 'h_in': 6.0}
_G6 = {'round': False, 'd_in': None, 'r_in': None, 'w_in': 6.0, 'h_in': 6.0}


def _swirl_nodes(n90, n45, fam='Rectangular Elbow - Mitered'):
    nodes = {2200: _unit({2201: SQ6})}
    for nid, nrm, ang in ((2202, n90, 90.0), (2204, n45, 45.0), (2206, n45, 45.0)):
        e = Elem(kind='elbow', family_name=fam, is_round=False, geom=dict(_G6, angle=ang))
        e.normal = nrm
        nodes[nid] = e
    return nodes


def _swirl_ducts(cls, d_h=6.0):
    d = build([(2201, 0.0, 0.394, 400.0, 0.25, cls), (2203, 0.0, 0.473, 400.0, 0.25, cls),
               (2205, 0.0, 1.157, 400.0, 0.25, cls), (2207, 0.0, 21.5, 400.0, 0.25, cls)])
    for _r in d.values():
        _r.d_h_in = d_h
    return d


c56 = {2200: [2201], 2201: [2202], 2202: [2203], 2203: [2204], 2204: [2205],
       2205: [2206], 2206: [2207], 2207: [2210]}
t56e = {2210: (100.0, 'Exhaust Air', 'EG-56')}
ea56 = _fse_run(_swirl_nodes((0.0, 1.0, 0.0), (1.0, 0.0, 0.0)), c56,
                _swirl_ducts('Exhaust Air'), t56e, 2200, cls='Exhaust Air')
# All three elbows (0.4, 0.9, 2.0 ft) are also under 5 D (2.5 ft): their own
# RED lines (Colin, 2026-11-03), so the swirl line is picked out by its lead.
_r56 = [t for k, t in ea56.get('fse_notes', ()) if k == 'red' and 'swirl' in t]


def _swirl_red(rec):
    return [t for k, t in rec.get('fse_notes', ()) if k == 'red' and
            ('swirl' in t.lower() or 'Close-coupled' in t)]
# Straight duct check (Colin, 2026-11-03): at the inlet 2 H = 1.0 ft, so the
# elbows at 0.39 and 0.87 ft and the ducts before them are RED as well.
RED56_IN = {2201, 2202, 2203, 2204}
check("case56 EA 61: the 90 and the first 45 turn in different planes within 6 Dh: RED, both drawn red",
      {2202, 2204} <= ea56.get('fse_red_ids', set()) and len(_r56) == 1
      and _r56[0].startswith('RED: Inlet swirl at the exhaust inlet: elbows 2202 (0.4 ft) '
                             'and 2204 (0.9 ft) turn in different planes within 6 Dh '
                             '(3.0 ft, Dh 6 in.)')
      and 'Fig. 12' in _r56[0],
      "red %r notes %r" % (ea56.get('fse_red_ids'), _r56))
check("case56 swirl is not priced: the system effect is the ER7-1 elbow alone",
      abs(ea56.get('fse_inwc', -1) -
          FSE.inlet_effect(False, 6.0, 6.0, 0.25, 100.0, 0.394, pv_of=pv)['loss']) < 1e-12,
      "fse %r" % ea56.get('fse_inwc'))
check("case56 a RED line is an Action required correction",
      duct_perf_console.correction('red', _r56[0] if _r56 else '')[0] == duct_perf_console.ACTION)
ea56s = _fse_run(_swirl_nodes((1.0, 0.0, 0.0), (1.0, 0.0, 0.0)), c56,
                 _swirl_ducts('Exhaust Air'), t56e, 2200, cls='Exhaust Air')
check("case56 three elbows turning in one plane: no swirl RED, the two under 2 H and their ducts",
      ea56s.get('fse_red_ids') == RED56_IN and not _swirl_red(ea56s),
      "red %r" % (ea56s.get('fse_red_ids'),))
ea56n = _fse_run(_swirl_nodes(None, (1.0, 0.0, 0.0)), c56,
                 _swirl_ducts('Exhaust Air'), t56e, 2200, cls='Exhaust Air')
check("case56 the 90's plane unreadable: only the two 45s compared, same plane, no swirl RED",
      ea56n.get('fse_red_ids') == RED56_IN and not _swirl_red(ea56n), "red %r" % (ea56n.get('fse_red_ids'),))
ea56w = _fse_run(_swirl_nodes((0.0, 1.0, 0.0), (1.0, 0.0, 0.0)), c56,
                 _swirl_ducts('Exhaust Air', d_h=1.0), t56e, 2200, cls='Exhaust Air')
check("case56 Dh 1 in. (6 Dh 0.5 ft): only the 90 is inside, no swirl RED",
      ea56w.get('fse_red_ids') == RED56_IN and not _swirl_red(ea56w), "red %r" % (ea56w.get('fse_red_ids'),))
ra56 = _fse_run(_swirl_nodes((0.0, 1.0, 0.0), (1.0, 0.0, 0.0)), c56,
                _swirl_ducts('Return Air'), {2210: (100.0, 'Return Air', 'RG-56')}, 2200,
                cls='Return Air')
check("case56 return inlet: the same RED inlet swirl line",
      ra56.get('fse_red_ids') == RED56_IN
      and any(k == 'red' and t.startswith('RED: Inlet swirl at the return inlet')
              for k, t in ra56.get('fse_notes', ())),
      "notes %r" % (ra56.get('fse_notes'),))
sa56 = _fse_run(_swirl_nodes((0.0, 1.0, 0.0), (1.0, 0.0, 0.0)), c56,
                _swirl_ducts('Supply Air'), {2210: (100.0, 'Supply Air', 'SD-56')}, 2200)
# supply outlet: Le 1.4 ft (0.25 ft2), RED line 0.5 Le = 0.71 ft: the elbow
# at 0.4 ft and its duct are RED with the swirl pair (Colin, 2026-11-03).
check("case56 supply outlet: RED as close-coupled elbows, not called swirl",
      sa56.get('fse_red_ids') == {2201, 2202, 2204}
      and any(k == 'red' and t.startswith('RED: Close-coupled elbows at the supply outlet')
              and 'swirl' not in t.lower() and 'p. 34.11' in t
              for k, t in sa56.get('fse_notes', ())),
      "notes %r" % (sa56.get('fse_notes'),))
ea56o = _fse_run(_swirl_nodes((0.0, 1.0, 0.0), (1.0, 0.0, 0.0)), c56,
                 _swirl_ducts('Exhaust Air'), t56e, 2200, fan=None, cls='Exhaust Air')
check("case56 fan system effect unticked (Q4): the swirl and straight duct checks still run, nothing priced",
      ea56o.get('fse_red_ids') == RED56_IN and _swirl_red(ea56o)
      and ea56o.get('fse_inwc') == 0.0
      and all(k == 'red' for k, _t in ea56o.get('fse_notes', ())),
      "red %r notes %r" % (ea56o.get('fse_red_ids'), ea56o.get('fse_notes')))

# ── case 57: Include backdraft damper (Colin, 2026-11-02) ─────────────────────
# An exhaust fan run with no backdraft damper modeled pays one 'backdraft'
# drop; one drawn on the run is not charged twice; a root outside
# backdraft_roots (a ceiling fan, built in) and a supply run never pay it.
COMP57 = {'diffuser': 0.05, 'backdraft': 0.07, 'damper': 0.25}
n57 = {2400: _unit({}), 2401: Elem('duct')}
c57 = {2400: [2401, 2402], 2401: [2410], 2402: [2420]}
d57 = build([(2401, 0.02, 10.0, 600.0, 1.0, 'Exhaust Air'),
             (2402, 0.02, 10.0, 600.0, 1.0, 'Supply Air')])
t57 = {2410: (600.0, 'Exhaust Air', 'EG-57'), 2420: (600.0, 'Supply Air', 'SD-57')}
off57 = crit([2400], c57, d57, t57, n57, 0.0, comp_values=COMP57)[2400]
on57 = crit([2400], c57, d57, t57, n57, 0.0, comp_values=COMP57,
            backdraft_roots={2400})[2400]
check("case57 ticked: the exhaust run pays one backdraft drop (0.07 here)",
      abs(on57['Exhaust Air']['component_inwc'] - off57['Exhaust Air']['component_inwc'] - 0.07) < 1e-12
      and list(on57['Exhaust Air']['component_keys']).count('backdraft') == 1,
      "keys %r" % (on57['Exhaust Air']['component_keys'],))
check("case57 ticked: a take-off line names it as assumed",
      any(i[0] == 'Accessory (assumed)' and i[1] == 'Backdraft damper at the fan'
          for i in on57['Exhaust Air']['items'])
      and abs(_items_sum(on57['Exhaust Air']) - on57['Exhaust Air']['subtotal_inwc']) < 1e-9)
check("case57 ticked: the supply run from the same unit is unchanged",
      abs(on57['Supply Air']['total_inwc'] - off57['Supply Air']['total_inwc']) < 1e-12)
check("case57 a root outside backdraft_roots (a ceiling fan) pays none",
      'backdraft' not in crit([2400], c57, d57, t57, n57, 0.0, comp_values=COMP57,
                              backdraft_roots=set())[2400]['Exhaust Air']['component_keys'])
n57m = {2400: _unit({}), 2401: Elem('duct'),
        2403: Elem('accessory', family_name='RJA - Backdraft Damper - Automatic - Round')}
c57m = {2400: [2401], 2401: [2403], 2403: [2410]}
m57 = crit([2400], c57m, build([(2401, 0.02, 10.0, 600.0, 1.0, 'Exhaust Air')]),
           {2410: (600.0, 'Exhaust Air', 'EG-57')}, n57m, 0.0, comp_values=COMP57,
           backdraft_roots={2400})[2400]['Exhaust Air']
check("case57 a backdraft damper modeled on the run: priced once, none assumed",
      list(m57['component_keys']).count('backdraft') == 1
      and not any(i[0] == 'Accessory (assumed)' for i in m57['items']),
      "keys %r" % (m57['component_keys'],))

# ── case 60 to 64: fan velocity basis and the 5 D elbow (Colin, 2026-11-03)
FSE_DEF = dict(FSE_ON, velocity_basis='default', fan_fpm=2500.0)

# 60: case 32's elbow on the outlet at 1000 fpm, default basis: priced at
#     2500 fpm, said so, GRAY, and RED under 5 D (D 24 in., 10 ft).
sa60 = _fse_run(n32, c32, d32, t32, 920, fan=FSE_DEF)
check("case60 default basis: FSE2 = 3.80 x Pv(2500), not Pv(1000)",
      abs(sa60.get('fse_inwc', -1) - 3.80 * pv(2500.0)) < 1e-9,
      "got %r" % sa60.get('fse_inwc'))
check("case60 a NOTE names the fan velocity over the connection's",
      any(k == 'note' and '2500 fpm, greater than the connection velocity of 1000 fpm' in t
          for k, t in sa60.get('fse_notes', ())), "notes %r" % (sa60.get('fse_notes'),))
check("case60 1.48 in. wc is GRAY",
      sa60.get('fse_gray_ids') == {921}, "gray %r" % (sa60.get('fse_gray_ids'),))
check("case60 elbow at 0 ft is RED under 0.5 Le (Le 5.7 ft, Eq. 37) and drawn red",
      sa60.get('fse_red_ids') == {921}
      and any(k == 'red' and t.startswith('RED: Elbow 921, 0.0 ft from the supply outlet, '
                                          'L/Le 0.00 (Le 5.7 ft, Eq. 37). Provide at least '
                                          '2.9 ft of straight duct; 5.8 ft (1.0 Le) removes '
                                          'the system effect.')
              for k, t in sa60.get('fse_notes', ())),
      "red %r notes %r" % (sa60.get('fse_red_ids'), sa60.get('fse_notes')))
check("case60 the take-off still adds up",
      abs(_items_sum(sa60) - sa60['subtotal_inwc']) < 1e-9)

# 61: a connection faster than the default governs, no velocity NOTE.
d61 = build([(922, 0.0, 10.0, 3000.0, 4.0, 'Supply Air')])
sa61 = _fse_run(n32, c32, d61, {930: (12000.0, 'Supply Air', 'SD-61')}, 920, fan=FSE_DEF)
check("case61 connection 3000 fpm over the 2500 default: priced at 3000",
      abs(sa61.get('fse_inwc', -1) - 3.80 * pv(3000.0)) < 1e-9
      and not any('default fan velocity' in t for _k, t in sa61.get('fse_notes', ())),
      "got %r" % sa61.get('fse_inwc'))

# 62: case 33's elbow 8 ft out, past Le 5.71 ft: no system effect and not
#     RED (Colin, 2026-11-03: Le at an outlet, not 5 D); at 5.5 ft it is.
sa62 = _fse_run(n33, c33, d33, {950: (4000.0, 'Supply Air', 'SD-33')}, 940)
check("case62 outlet elbow 8 ft out, past Le 5.71 ft: not priced, not RED",
      sa62.get('fse_inwc') == 0.0 and not sa62.get('fse_red_ids'),
      "fse %r red %r" % (sa62.get('fse_inwc'), sa62.get('fse_red_ids')))
d62b = build([(941, 0.0, 5.5, 1000.0, 4.0, 'Supply Air'),
              (943, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa62b = _fse_run(n33, c33, d62b, {950: (4000.0, 'Supply Air', 'SD-33')}, 940)
check("case62 outlet elbow 5.5 ft out, inside Le but past 0.5 Le (2.86 ft): priced, not RED",
      any(i[4] and i[4][0].startswith('Fan system effect') for i in sa62b.get('items', ()))
      and not sa62b.get('fse_red_ids'),
      "red %r" % (sa62b.get('fse_red_ids'),))
for _l62, _red62 in ((2.8, True), (2.9, False)):
    _d62 = build([(941, 0.0, _l62, 1000.0, 4.0, 'Supply Air'),
                  (943, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
    _s62 = _fse_run(n33, c33, _d62, {950: (4000.0, 'Supply Air', 'SD-33')}, 940)
    check("case62 outlet elbow at %.1f ft (0.5 Le 2.86 ft): %s" % (
              _l62, 'elbow and its duct RED' if _red62 else 'not RED'),
          (_s62.get('fse_red_ids') == {941, 942}) if _red62 else not _s62.get('fse_red_ids'),
          "red %r" % (_s62.get('fse_red_ids'),))

# 63: at an INLET, 2 D on the larger of the opening and the first duct: a
#     12 x 12 return opening (2 H 2 ft) into a 24 x 24 duct (2 H 4 ft),
#     elbow at 3 ft: RED with its duct.
SQ12 = {'round': False, 'area_ft2': 1.0, 'd_in': None, 'w_in': 12.0, 'h_in': 12.0}
n63 = dict(n33)
n63[940] = _unit({941: SQ12})
d63 = build([(941, 0.0, 3.0, 1000.0, 4.0, 'Return Air'),
             (943, 0.0, 5.0, 1000.0, 4.0, 'Return Air')])
d63[EID(941)].size = '24x24"'
sa63 = _fse_run(n63, c33, d63, {950: (4000.0, 'Return Air', 'RG-63')}, 940,
                cls='Return Air')
check("case63 12x12 return opening into a 24x24 duct: 2 H on 24 in., elbow at 3 ft RED with its duct",
      sa63.get('fse_red_ids') == {941, 942}
      and any(k == 'red' and t.startswith('RED: Elbow 942, 3.0 ft from the return inlet, '
                                          '1.5 H (H 24 in., the first duct, 24x24). Provide '
                                          'at least 4.0 ft of straight duct (2 H of the '
                                          'first duct); 10.0 ft (5 H) is where the loss '
                                          'stops falling.')
              for k, t in sa63.get('fse_notes', ()))
      and sa63.get('fse_red_ducts') == {941: ('Straight duct: elbow 942 at 3.0 ft, under 2 H',
                                              'Straight duct 4.0 ft')},
      "red %r notes %r" % (sa63.get('fse_red_ids'), sa63.get('fse_notes')))

# 63b: the inlet 2 D edge, rect (H) and round (D): a 24 in. inlet, 2 D 4.0 ft.
for _shape63, _conn63, _dim63 in (('rect', SQ24, 'H'), ('round', RD24, 'D')):
    for _l63, _red63 in ((3.9, True), (4.0, False)):
        _n = dict(n33)
        _n[940] = _unit({941: _conn63})
        _d = build([(941, 0.0, _l63, 1000.0, 4.0, 'Return Air'),
                    (943, 0.0, 5.0, 1000.0, 4.0, 'Return Air')])
        _r = _fse_run(_n, c33, _d, {950: (4000.0, 'Return Air', 'RG-63b')}, 940,
                      cls='Return Air')
        _txt = [t for k, t in _r.get('fse_notes', ()) if k == 'red']
        check("case63b %s inlet elbow at %.1f ft (2 %s 4.0 ft): %s" % (
                  _shape63, _l63, _dim63, 'RED with its duct' if _red63 else 'not RED'),
              ((_r.get('fse_red_ids') == {941, 942} and len(_txt) == 1
                and '(%s 24 in., the %s inlet collar)' % (
                    _dim63, '24x24' if _dim63 == 'H' else '24 in. round') in _txt[0]
                and '4.0 ft of straight duct (2 %s of the collar); 10.0 ft (5 %s)' % (
                    _dim63, _dim63) in _txt[0])
               if _red63 else not _r.get('fse_red_ids')),
              "red %r notes %r" % (_r.get('fse_red_ids'), _txt))

# 64: case 34's transition 2 ft out, default basis: priced at the fan
#     velocity (x 2.5 here), named, the walk unchanged past it.
n64 = dict(n34)
n64[962] = Elem(kind='transition', family_name='Rectangular Transition - Length',
                is_round=False,
                tgeom=([(4.0, 'rect', 24.0, 24.0, P(0.0)),
                        (6.0, 'rect', 36.0, 24.0, P(0.5))], 0.0))
d64 = build([(961, 0.0, 2.0, 1000.0, 4.0, 'Supply Air'),
             (963, 0.0, 1.0, 666.7, 6.0, 'Supply Air'),
             (965, 0.0, 5.0, 666.7, 6.0, 'Supply Air')])
c64 = _fse_run(n64, c34, d64, {970: (4000.0, 'Supply Air', 'SD-64')}, 960)
f64 = _fse_run(n64, c34, d64, {970: (4000.0, 'Supply Air', 'SD-64')}, 960, fan=FSE_DEF)
_t64c = [i for i in c64.get('items', ()) if i[0] == 'Fitting 962']
_t64f = [i for i in f64.get('items', ()) if i[0] == 'Fitting 962']
check("case64 connection basis: the transition is priced at the model velocity",
      _t64c and 'at fan velocity' not in _t64c[0][1], "items %r" % (_t64c,))
# Since 2026-11-04 the transition is in the group (straight duct, transition,
# first elbow) charged once with the fan system effect, never at the fan
# velocity: here the fan system effect wins and the transition reads 0.
check("case64 default basis: the transition is in the group, at the model velocity, not boosted",
      _t64f and _t64c and 'at fan velocity' not in _t64f[0][1]
      and _t64f[0][1].endswith('in the fan system effect') and _t64f[0][3] == 0.0
      and 'at 667 fpm' in _t64f[0][1],
      "conn %r fan %r" % (_t64c, _t64f))
check("case64 a NOTE names the group, no fan velocity transition NOTE",
      any(k == 'note' and 'Fan system effect charged, the straight duct' in t
          and 'supply outlet' in t for k, t in f64.get('fse_gov_notes', ()))
      and not any(t.startswith('NOTE: transition') for _k, t in f64.get('fse_notes', ())),
      "notes %r" % (f64.get('fse_notes'),))
check("case64 a transition in the group is named: 'straight duct, transition and elbow'",
      any(t.endswith('Fan system effect charged, the straight duct, transition and '
                     'elbow are not charged again.')
          for _k, t in f64.get('fse_gov_notes', ())),
      "gov %r" % (f64.get('fse_gov_notes'),))
check("case64 the group's own losses are left out of friction and fittings",
      abs(f64['fitting_inwc']) < 1e-12 and f64.get('fse_inwc', 0) > 0.0,
      "fit %r fse %r" % (f64['fitting_inwc'], f64.get('fse_inwc')))
check("case64 the ducts past it keep their own velocity",
      abs(f64['friction_inwc'] - c64['friction_inwc']) < 1e-12
      and abs(_items_sum(f64) - f64['subtotal_inwc']) < 1e-9)

# 65: review 2026-11-03. Every elbow inside the window is RED, past a tap on the
#     main and past the first elbow; a transition past the pricing window
#     (Le) is not priced at the fan velocity.
n65 = {3000: _unit({3001: SQ24}), 3001: Elem(kind='duct', tap_pos={3002: 1.0}),
       3002: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       3003: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False),
       3005: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c65 = {3000: [3001], 3001: [3002, 3003], 3002: [3010], 3003: [3004], 3004: [3005],
       3005: [3006], 3006: [3011]}
d65 = build([(3001, 0.0, 2.0, 1000.0, 4.0, 'Supply Air'),
             (3004, 0.0, 0.5, 875.0, 4.0, 'Supply Air'),
             (3006, 0.0, 10.0, 875.0, 4.0, 'Supply Air')])
t65 = {3010: (500.0, 'Supply Air', 'SD-a'), 3011: (3500.0, 'Supply Air', 'SD-b')}
sa65 = _fse_run(n65, c65, d65, t65, 3000)
_red65 = [t for k, t in sa65.get('fse_notes', ()) if k == 'red']
check("case65 tap at 1 ft, elbows at 2.0 and 2.5 ft (0.5 Le 2.86 ft): both RED with their ducts, past the tap",
      sa65.get('fse_red_ids') == {3001, 3003, 3004, 3005} and len(_red65) == 2
      and _red65[0].startswith('RED: Elbow 3003, 2.0 ft')
      and _red65[1].startswith('RED: Elbow 3005, 2.5 ft'),
      "red %r notes %r" % (sa65.get('fse_red_ids'), _red65))
check("case65 the tap is still the priced window's first disruption (WARNING, no fse)",
      sa65.get('fse_inwc') == 0.0
      and any(k == 'warn' and t.startswith('WARNING: tap 3002') for k, t in
              sa65.get('fse_notes', ())), "notes %r" % (sa65.get('fse_notes'),))
n65t = {3100: _unit({3101: SQ24}),
        3102: Elem(kind='transition', family_name='Rectangular Transition - Length',
                   is_round=False,
                   tgeom=([(4.0, 'rect', 24.0, 24.0, P(0.0)),
                           (6.0, 'rect', 36.0, 24.0, P(0.5))], 0.0))}
c65t = {3100: [3101], 3101: [3102], 3102: [3103], 3103: [3110]}
d65t = build([(3101, 0.0, 7.0, 1000.0, 4.0, 'Supply Air'),
              (3103, 0.0, 10.0, 666.7, 6.0, 'Supply Air')])
f65t = _fse_run(n65t, c65t, d65t, {3110: (4000.0, 'Supply Air', 'SD-65')}, 3100,
                fan=FSE_DEF)
_t65 = [i for i in f65t.get('items', ()) if i[0] == 'Fitting 3102']
check("case65 transition at 7 ft, past Le 5.71 ft: not priced at the fan velocity",
      _t65 and 'at fan velocity' not in _t65[0][1]
      and not any(t.startswith('NOTE: transition') for _k, t in f65t.get('fse_notes', ())),
      "items %r" % (_t65,))

# ── case 66: exhaust fan type from its connections (Colin, 2026-11-03) ──
# Inline fan 4000: intake 4001 (12x12, 10 ft) to grille 4010 at 500 cfm;
# discharge 4002 (12x12, 1 ft, no terminals) to a radius elbow 4003 and duct
# 4004 (6 ft). The discharge is priced at the 500 grille cfm and added to
# the exhaust run; its outlet fan system effect is added too.
SQ12E = {'round': False, 'area_ft2': 1.0, 'd_in': None, 'w_in': 12.0, 'h_in': 12.0}
FSE_SWSI = dict(FSE_ON, fan_type='SWSI')


def _inline66(with_discharge=True):
    nodes = {4000: _unit(dict([(4001, SQ12E)] + ([(4002, SQ12E)] if with_discharge else []))),
             4003: Elem(kind='elbow', family_name='Rectangular Elbow - Radius',
                        is_round=False)}
    ch = {4000: [4001] + ([4002] if with_discharge else []), 4001: [4010]}
    rows = [(4001, 0.02, 10.0, 500.0, 1.0, 'Exhaust Air')]
    if with_discharge:
        ch.update({4002: [4003], 4003: [4004], 4004: []})
        rows += [(4002, 0.0, 1.0, 0.0, 1.0, 'Exhaust Air'),
                 (4004, 0.0, 6.0, 0.0, 1.0, 'Exhaust Air')]
    d = build(rows)
    for _r in d.values():
        _r.d_h_in = 12.0
    return nodes, ch, d, {4010: (500.0, 'Exhaust Air', 'EG-66')}


n66, c66, d66, t66 = _inline66()
in66 = crit([4000], c66, d66, t66, n66, 10.0, fan_effect=FSE_SWSI,
            inline_fan_roots={4000})[4000]['Exhaust Air']
base66 = crit([4000], c66, d66, t66, n66, 10.0, fan_effect=FSE_SWSI)[4000]['Exhaust Air']
# The 1 ft duct before the discharge elbow is in the outlet group (charged
# once with the fan system effect, which wins here): only 4004's 6 ft adds.
_fr66 = hvac_graph.duct_friction_loss_per_100ft(500.0, 12.0) * 6.0 / 100.0
check("case66 inline fan: read as inline, with a NOTE naming intake and discharge",
      in66.get('exhaust_fan_kind') == 'inline'
      and any(t.startswith('NOTE: exhaust fan read as an inline fan: intake at element 4001, '
                           'discharge at element 4002. The discharge carries 500 cfm, '
                           'the sum of the exhaust terminal airflows, and is priced once '
                           'on the index run.')
              for _k, t in in66.get('fse_notes', ())),
      "kind %r notes %r" % (in66.get('exhaust_fan_kind'), in66.get('fse_notes')))
check("case66 the discharge ducts are priced at the grille cfm (6 ft past the group at 500 fpm)",
      abs(in66['friction_inwc'] - base66['friction_inwc'] - _fr66) < 1e-9
      and in66['length_ft'] - base66['length_ft'] == 7.0,
      "in %r base %r expected +%r" % (in66['friction_inwc'], base66['friction_inwc'], _fr66))
_dis66 = [i for i in in66['items'] if i[0].startswith('Discharge ')]
check("case66 the take-off shows the discharge as its own lines",
      [i[0] for i in _dis66] == ['Discharge duct 4002', 'Discharge fitting 4003',
                                 'Discharge duct 4004']
      and _dis66[1][4][0].startswith('Discharge rect radius elbow')
      and _dis66[0][3] == 0.0 and _dis66[1][3] == 0.0 and _dis66[2][3] > 0.0
      and _dis66[1][1].endswith('in the fan system effect'), "items %r" % (_dis66,))
_out66 = FSE.outlet_effect('SWSI', 0.5, 1.0, 500.0, 1.0, pv_of=pv)
check("case66 outlet fan system effect at the discharge elbow (SWSI, L 1 ft of Le 2.9) added",
      any(i[4] and i[4][0] == 'Fan system effect, exhaust outlet' for i in in66['items'])
      and _out66['loss'] > 0
      and abs(in66['fse_inwc'] - base66['fse_inwc'] - _out66['loss']) < 1e-9,
      "fse %r base %r outlet %r" % (in66['fse_inwc'], base66['fse_inwc'], _out66['loss']))
check("case66 the take-off still adds up",
      abs(_items_sum(in66) - in66['subtotal_inwc']) < 1e-9)
cap66 = crit([4000], c66, d66, t66, n66, 10.0, fan_effect=FSE_SWSI,
             inline_fan_roots={4000}, exhaust_cap=True)[4000]['Exhaust Air']
check("case66 Include roof cap ticked: one cap on the inline fan's run",
      list(cap66['component_keys']).count('cap') == 1)

# a roof fan: inlet only, no discharge walked; a ticked cap is still charged
n66r, c66r, d66r, t66r = _inline66(with_discharge=False)
roof66 = crit([4000], c66r, d66r, t66r, n66r, 10.0, fan_effect=FSE_SWSI,
              inline_fan_roots={4000}, exhaust_cap=True)[4000]['Exhaust Air']
check("case66 inlet only: read as an upblast or downblast roof fan, no discharge lines",
      roof66.get('exhaust_fan_kind') == 'roof'
      and not any(i[0].startswith('Discharge ') for i in roof66['items'])
      and any('upblast or downblast roof fan' in t for _k, t in roof66.get('fse_notes', ())),
      "kind %r" % (roof66.get('exhaust_fan_kind'),))
check("case66 roof fan: the ticked cap is still charged (the engineer decides)",
      list(roof66['component_keys']).count('cap') == 1)
# neither connection reaches a grille: fallback NOTE, nothing added
n66n, c66n, d66n, _t = _inline66()
none66 = crit([4000], c66n, d66n, {}, n66n, 10.0, fan_effect=FSE_SWSI,
              inline_fan_roots={4000})
check("case66 no connection reaches a grille: no run, nothing priced on the discharge",
      not none66.get(4000), "got %r" % (none66.get(4000),))

# 67: review 2 (2026-11-03). A discharge transition takes the fan velocity
#     like one on an intake; a second discharge gets no fan system effect;
#     a connection into equipment is no discharge; a flex discharge is
#     flagged, exhaust is never flex.
n67 = {4100: _unit({4101: SQ12E, 4102: SQ12E}),
       4103: Elem(kind='transition', family_name='Rectangular Transition - Length',
                  is_round=False,
                  tgeom=([(1.0, 'rect', 12.0, 12.0, P(0.0)),
                          (1.5, 'rect', 18.0, 12.0, P(0.5))], 0.0)),
       4105: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c67 = {4100: [4101, 4102], 4101: [4110], 4102: [4103], 4103: [4104], 4104: [4105],
       4105: [4106], 4106: []}
d67 = build([(4101, 0.02, 10.0, 500.0, 1.0, 'Exhaust Air'),
             (4102, 0.0, 0.5, 0.0, 1.0, 'Exhaust Air'),
             (4104, 0.0, 0.5, 0.0, 1.5, 'Exhaust Air'),
             (4106, 0.0, 5.0, 0.0, 1.5, 'Exhaust Air')])
for _r in d67.values():
    _r.d_h_in = 12.0
t67 = {4110: (500.0, 'Exhaust Air', 'EG-67')}
c67r = crit([4100], c67, d67, t67, n67, 10.0, fan_effect=FSE_SWSI,
            inline_fan_roots={4100})[4100]['Exhaust Air']
f67r = crit([4100], c67, d67, t67, n67, 10.0,
            fan_effect=dict(FSE_SWSI, velocity_basis='default', fan_fpm=2500.0),
            inline_fan_roots={4100})[4100]['Exhaust Air']
_t67c = [i for i in c67r['items'] if i[0] == 'Discharge fitting 4103']
_t67f = [i for i in f67r['items'] if i[0] == 'Discharge fitting 4103']
check("case67 discharge transition at the default basis: in the outlet group, never at fan velocity",
      _t67c and _t67f and 'at fan velocity' not in _t67f[0][1]
      and 'at fan velocity' not in _t67c[0][1]
      and _t67f[0][1].endswith('in the fan system effect') and _t67f[0][3] == 0.0,
      "conn %r fan %r" % (_t67c, _t67f))
check("case67 the group NOTE names the discharge elbow, no transition NOTE",
      any('elbow 4105' in t and 'Fan system effect charged, the straight duct' in t
          for _k, t in f67r.get('fse_gov_notes', ()))
      and not any(t.startswith('NOTE: transition 4103') for _k, t in f67r.get('fse_notes', ())))

# a second discharge: no fan system effect record, a NOTE that the first is priced
n67b = dict(n67)
n67b[4100] = _unit({4101: SQ12E, 4102: SQ12E, 4120: SQ12E})
n67b[4121] = Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)
c67b = dict(c67)
c67b[4100] = [4101, 4102, 4120]
c67b.update({4120: [4121], 4121: [4122], 4122: []})
d67b = dict(d67)
d67b.update(build([(4120, 0.0, 0.2, 0.0, 1.0, 'Exhaust Air'),
                   (4122, 0.0, 2.0, 0.0, 1.0, 'Exhaust Air')]))
for _r in d67b.values():
    _r.d_h_in = 12.0
b67 = crit([4100], c67b, d67b, t67, n67b, 10.0, fan_effect=FSE_SWSI,
           inline_fan_roots={4100})[4100]['Exhaust Air']
check("case67 second discharge: not priced, no GRAY or RED for it, NOTE names the first",
      4121 not in (b67.get('fse_gray_ids') or set()) and 4121 not in (b67.get('fse_red_ids') or set())
      and not any(i[0].startswith('Discharge') and '412' in i[0] for i in b67['items'])
      and any('2 discharge connections; the one at element 4102 is priced' in t
              for _k, t in b67.get('fse_notes', ())),
      "red %r notes %r" % (b67.get('fse_red_ids'), b67.get('fse_notes')))

# a connection straight into other equipment is no discharge
n67e = {4200: _unit({4201: SQ12E}), 4202: Elem(kind='equipment')}
c67e = {4200: [4201, 4202], 4201: [4210], 4202: []}
d67e = build([(4201, 0.02, 10.0, 500.0, 1.0, 'Exhaust Air')])
e67 = crit([4200], c67e, d67e, {4210: (500.0, 'Exhaust Air', 'EG-67e')}, n67e, 10.0,
           fan_effect=FSE_SWSI, inline_fan_roots={4200})[4200]['Exhaust Air']
check("case67 connection into equipment: not read as inline, a NOTE, nothing discharged",
      e67.get('exhaust_fan_kind') is None
      and not any(i[0].startswith('Discharge') for i in e67['items'])
      and any('feeds other equipment' in t for _k, t in e67.get('fse_notes', ())),
      "kind %r notes %r" % (e67.get('exhaust_fan_kind'), e67.get('fse_notes')))

# a flex discharge pays one flex bend
n66f, c66f, d66f, t66f = _inline66()
d66f[EID(4004)].is_flex = True
fx66 = crit([4000], c66f, d66f, t66f, n66f, 10.0, fan_effect=FSE_SWSI,
            inline_fan_roots={4000})[4000]['Exhaust Air']
check("case67 flex discharge: flagged (exhaust shall be rigid), no flex bend priced",
      not any('flex bend' in i[0].lower() for i in fx66['items'])
      and any('flex duct 4004 on the exhaust discharge' in u for u in fx66['unpriced'])
      and abs(_items_sum(fx66) - fx66['subtotal_inwc']) < 1e-9,
      "unpriced %r" % (fx66['unpriced'],))

# ── case 68 to 69: inlet velocity and the turn plane H (Colin, 2026-11-04) ───
# 68: the fan velocity basis is for outlets only. A 24 x 24 return inlet at
#     1000 fpm, elbow 3 ft in: the default basis (2500 fpm) prices it the
#     same as the connection basis, with no fan velocity NOTE.
n68 = dict(n33)
n68[940] = _unit({941: SQ24})
d68 = build([(941, 0.0, 3.0, 1000.0, 4.0, 'Return Air'),
             (943, 0.0, 5.0, 1000.0, 4.0, 'Return Air')])
_t68 = {950: (4000.0, 'Return Air', 'RG-68')}
c68 = _fse_run(n68, c33, d68, _t68, 940, cls='Return Air')
f68 = _fse_run(n68, c33, d68, _t68, 940, fan=FSE_DEF, cls='Return Air')
check("case68 inlet on the default basis: priced at the connection velocity, as the connection basis",
      c68.get('fse_inwc', 0) > 0
      and abs(f68.get('fse_inwc', -1) - c68.get('fse_inwc', -2)) < 1e-12
      and not any('default fan velocity' in t for _k, t in f68.get('fse_notes', ())),
      "conn %r default %r notes %r" % (c68.get('fse_inwc'), f68.get('fse_inwc'),
                                       f68.get('fse_notes')))

# 69: a 24 x 12 return inlet, 2 H on the elbow's side in the plane of its
#     turn. Turning on the 12 in. side: 2 H 2.0 ft, so 3.0 ft is not RED and
#     1.9 ft is RED "(H 12 in.)". Turning on the 24 in. side: 2 H 4.0 ft,
#     3.0 ft is RED "(H 24 in.)".
for _w69, _l69, _red69 in ((12.0, 3.0, False), (12.0, 1.9, True), (24.0, 3.0, True)):
    _n = {940: _unit({941: RC2412}),
          942: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False,
                    geom={'round': False, 'd_in': None, 'r_in': 1.5 * _w69, 'w_in': _w69,
                          'h_in': 36.0 - _w69, 'angle': 90.0})}
    _d = build([(941, 0.0, _l69, 1000.0, 2.0, 'Return Air'),
                (943, 0.0, 5.0, 1000.0, 2.0, 'Return Air')])
    _d[EID(941)].size = '24x12"'
    _r = _fse_run(_n, c33, _d, {950: (2000.0, 'Return Air', 'RG-69')}, 940, cls='Return Air')
    _txt = [t for k, t in _r.get('fse_notes', ()) if k == 'red']
    _h = '%.0f' % _w69
    check("case69 24x12 inlet, elbow turning on its %s in. side at %.1f ft: %s" % (
              _h, _l69, 'RED (H %s in.)' % _h if _red69 else 'not RED'),
          ((_r.get('fse_red_ids') == {941, 942} and len(_txt) == 1
            and "(H %s in., the elbow's side in the plane of its turn)" % _h in _txt[0]
            and '%.1f ft of straight duct (2 H of the elbow)' % (2 * _w69 / 12.0) in _txt[0])
           if _red69 else not _r.get('fse_red_ids')),
          "red %r notes %r" % (_r.get('fse_red_ids'), _txt))

# 69b: a 36 x 8 inlet (a fan coil return) into a 10 in. round elbow past a
#      transition: H is the elbow's D, 10 in. (the opening's smaller side,
#      8 in., is less), 2 H 1.7 ft. At 2.0 ft not RED (the larger side, 36
#      in., would ask 6 ft); at 1.5 ft RED "(H 10 in.)".
RC368 = {'round': False, 'area_ft2': 2.0, 'd_in': None, 'w_in': 36.0, 'h_in': 8.0}
for _l69b, _red69b in ((2.0, False), (1.5, True)):
    _n = {940: _unit({941: RC368}),
          942: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                    geom={'round': True, 'd_in': 10.0, 'r_in': 15.0, 'w_in': None,
                          'h_in': None, 'angle': 90.0})}
    _d = build([(941, 0.0, _l69b, 1000.0, 2.0, 'Return Air'),
                (943, 0.0, 5.0, 1000.0, 2.0, 'Return Air')])
    _d[EID(941)].size = '36x8"'
    _r = _fse_run(_n, c33, _d, {950: (2000.0, 'Return Air', 'RG-69b')}, 940, cls='Return Air')
    _txt = [t for k, t in _r.get('fse_notes', ()) if k == 'red']
    check("case69b 36x8 inlet, 10 in. round elbow at %.1f ft (2 H 1.7 ft): %s" % (
              _l69b, 'RED (H 10 in.)' if _red69b else 'not RED'),
          ((_r.get('fse_red_ids') == {941, 942} and len(_txt) == 1
            and '(H 10 in., the 10 in. round elbow)' in _txt[0]
            and '(2 H of the elbow)' in _txt[0])
           if _red69b else not _r.get('fse_red_ids')),
          "red %r notes %r" % (_r.get('fse_red_ids'), _txt))

# 70: the max() group (Colin, 2026-11-04) with a transition in it, where the
#     duct and fittings win: case 64's run with the elbow pushed to about
#     5.6 ft (Le 5.71 ft), so the fan system effect is small and the
#     transition and elbow, priced at their model velocity, are charged.
d70 = build([(961, 0.0, 4.5, 1000.0, 4.0, 'Supply Air'),
             (963, 0.0, 0.6, 666.7, 6.0, 'Supply Air'),
             (965, 0.0, 5.0, 666.7, 6.0, 'Supply Air')])
f70 = _fse_run(n64, c34, d70, {970: (4000.0, 'Supply Air', 'SD-70')}, 960, fan=FSE_ON)
_g70 = [i for i in f70.get('items', ()) if i[0] in ('Fitting 962', 'Fitting 964')]
_fs70 = [i for i in f70.get('items', ()) if i[4] and i[4][0].startswith('Fan system effect')]
_off70 = _fse_run(n64, c34, d70, {970: (4000.0, 'Supply Air', 'SD-70')}, 960,
                  fan=dict(FSE_ON, on=False))
check("case70 transition in the group, own losses win: the group value on the FSE row, "
      "named 'straight duct, transition and elbow (SR4-1, ...)', C '-'",
      len(_fs70) == 1 and _fs70[0][3] > 0.0
      and abs(f70.get('fse_inwc', 0) - _fs70[0][3]) < 1e-12
      and _fs70[0][4][1].startswith('Duct losses: straight duct, transition and elbow (SR4-1')
      and _fs70[0][4][2] is None, "fse %r rows %r" % (f70.get('fse_inwc'), _fs70))
check("case70 the transition and elbow rows are listed at 0, in the fan system effect, "
      "and the subtotal is the same as with the fan system effect off",
      len(_g70) == 2 and all(i[3] == 0.0 and i[4][1].endswith('in the fan system effect')
                             for i in _g70)
      and abs(f70['subtotal_inwc'] - _off70['subtotal_inwc']) < 1e-12
      and abs(_items_sum(f70) - f70['subtotal_inwc']) < 1e-9,
      "rows %r sub %r off %r" % (_g70, f70.get('subtotal_inwc'), _off70.get('subtotal_inwc')))

# ── case71: manual volume damper keywords (Colin, 2026-11-04) ───────────────
_ca = fitting_tables.classify_accessory
check("case71 a manual volume damper, a volume damper and an MVD are balancing dampers",
      _ca('RJA - Manual Volume Damper - Round') == 'damper'
      and _ca('Volume Damper - Rectangular') == 'damper'
      and _ca('MVD - Round') == 'damper' and _ca('Damper (MVD)') == 'damper',
      "%r %r %r" % (_ca('RJA - Manual Volume Damper - Round'), _ca('MVD - Round'),
                    _ca('Damper (MVD)')))
check("case71 MVD matches as a whole word only, and fire / smoke names are unchanged",
      _ca('Xmvdx Coupling') is None and _ca('RJA - Fire Damper - Round') == 'fire_damper'
      and _ca('RJA - Fire Smoke Damper - Round (Motorized)') == 'fire_damper'
      and _ca('Sound Attenuator - Rectangular') is None,
      "%r %r" % (_ca('Xmvdx Coupling'), _ca('Sound Attenuator - Rectangular')))

# ── case72: double bend WARNING (Colin, 2026-11-04) ───────────────────────────
# Same plane elbows inside 6 Dh (the case56 nodes with one normal): a WARNING
# naming the first two, not RED, nothing priced for it.
def _dbl(rec):
    return [t for k, t in rec.get('fse_notes', ()) if k == 'warn' and 'Double bend' in t]
_d56s = _dbl(ea56s)
check("case72 three same plane elbows inside 6 Dh of an exhaust inlet: one double bend WARNING",
      len(_d56s) == 1 and _d56s[0].startswith(
          'WARNING: Double bend at the exhaust inlet: elbows 2202 (0.4 ft) and 2204 '
          '(0.9 ft) follow each other within 6 Dh (3.0 ft, Dh 6 in.) of the fan.')
      and 'Best practice' in _d56s[0] and 'p. 34.11' in _d56s[0],
      "notes %r" % (ea56s.get('fse_notes'),))
check("case72 the double bend is a WARNING in Action required, not RED, and not priced",
      duct_perf_console.correction('warn', _d56s[0] if _d56s else '')[0] == duct_perf_console.ACTION
      and ea56s.get('fse_red_ids') == RED56_IN
      and abs(ea56s.get('fse_inwc', -1) - ea56.get('fse_inwc', -2)) < 1e-12,
      "red %r fse %r / %r" % (ea56s.get('fse_red_ids'), ea56s.get('fse_inwc'), ea56.get('fse_inwc')))
check("case72 unreadable plane on one elbow: still two elbows inside 6 Dh, WARNING",
      len(_dbl(ea56n)) == 1, "notes %r" % (ea56n.get('fse_notes'),))
check("case72 no WARNING with a RED swirl pair, or with one elbow inside 6 Dh",
      not _dbl(ea56) and not _dbl(ea56w) and not _dbl(ra56),
      "%r %r" % (_dbl(ea56), _dbl(ea56w)))
sa72 = _fse_run(_swirl_nodes((1.0, 0.0, 0.0), (1.0, 0.0, 0.0)), c56,
                _swirl_ducts('Supply Air'), {2210: (100.0, 'Supply Air', 'SD-72')}, 2200)
check("case72 a supply outlet double bend: the same WARNING at the supply outlet",
      len(_dbl(sa72)) == 1 and _dbl(sa72)[0].startswith('WARNING: Double bend at the supply outlet'),
      "notes %r" % (sa72.get('fse_notes'),))

# ── case73: two taps at one station (F2, Colin 2026-11-04) ───────────────────
# A 1000 cfm main ends in two facing taps at the same 8.0 ft station, 500 cfm
# each (the M05 R5 shoes). Each is a tap from its own flow (Qb/Qc 0.5 on the
# arriving 1000 cfm), neither is an "end of main turn", and Revit's connector
# order (the child order) cannot change the answer.
def _run73(order, cls, qa=500.0, qb=500.0, pos_b=8.0, fric_b=0.0):
    ch = {699: [700], 700: list(order), 701: [711], 702: [712], 711: [721], 712: [722]}
    n = {700: Elem(kind='duct', tap_pos={701: 8.0, 702: pos_b}),
         701: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
         702: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
         711: Elem(kind='duct'), 712: Elem(kind='duct')}
    d = build([(700, 0.0, 10.0, 1000.0, 1.0, cls),
               (711, 0.0, 5.0, qa / 0.35, 0.35, cls),
               (712, fric_b, 5.0, qb / 0.35, 0.35, cls)])
    for _dr in d.values():
        _dr.size = '12x12"'
    t = {721: (qa, cls, 'T-A'), 722: (qb, cls, 'T-B')}
    return crit([699], ch, d, t, n).get(699, {}).get(cls, {})
def _tap_rows(r):
    return sorted((i[0], round(i[3], 12), i[1]) for i in r.get('items', ())
                  if i[0].startswith('Fitting 70'))
for _cls73 in ('Return Air', 'Supply Air'):
    _a73 = _run73((701, 702), _cls73)
    _b73 = _run73((702, 701), _cls73)
    check("case73 %s: both child orders give the same index run, total and tap rows" % _cls73,
          _a73.get('terminal_id') == _b73.get('terminal_id')
          and abs(_a73.get('total_inwc', -1) - _b73.get('total_inwc', -2)) < 1e-12
          and _tap_rows(_a73) == _tap_rows(_b73),
          "A %r %r / B %r %r" % (_a73.get('terminal_id'), _tap_rows(_a73),
                                 _b73.get('terminal_id'), _tap_rows(_b73)))
    check("case73 %s: the coincident tap is a tap, not an end of main turn" % _cls73,
          _tap_rows(_a73) and not any('end of main' in r[2] for r in _tap_rows(_a73)),
          "rows %r" % (_tap_rows(_a73),))
# Equal taps price the same whichever is the index; the tie goes to the lower id.
_e73 = _run73((702, 701), 'Supply Air')
_l701 = [r[1] for r in _tap_rows(_e73)]
check("case73 equal taps: the index run leaves through the lower id tap 701",
      _e73.get('terminal_id') == 721, "term %r rows %r" % (_e73.get('terminal_id'), _tap_rows(_e73)))
# A tap 0.05 ft on is the same station; one 0.5 ft on is not (the main then
# carries only the second tap's flow there and it is the end of main turn).
_n73 = _run73((702, 701), 'Supply Air', qa=300.0, qb=700.0, pos_b=8.05, fric_b=0.5)
_f73 = _run73((702, 701), 'Supply Air', qa=300.0, qb=700.0, pos_b=8.5, fric_b=0.5)
check("case73 a tap 0.05 ft on shares the station: no end of main turn",
      _n73.get('terminal_id') == 722 and not any('end of main' in r[2] for r in _tap_rows(_n73)), "rows %r" % (_tap_rows(_n73),))
check("case73 a tap 0.5 ft on is its own station: the last tap is the end of main turn",
      _f73.get('terminal_id') == 722 and any('end of main' in r[2] for r in _tap_rows(_f73)), "rows %r" % (_tap_rows(_f73),))

# ── case75: a CAPPED main's last tap is a tap (Colin, 2026-11-05, TB-M10C) ───
# The case73 layout with the second tap 0.5 ft on (its own station, the last
# tap) and an end cap 703 on main 700: the main runs on past the tap to the
# cap, so the last tap prices from the tap table, never as an end of main turn.
def _run75(cap):
    ch = {699: [700], 700: [701, 702] + ([703] if cap else []),
          701: [711], 702: [712], 711: [721], 712: [722]}
    n = {700: Elem(kind='duct', tap_pos={701: 8.0, 702: 8.5}),
         701: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
         702: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
         711: Elem(kind='duct'), 712: Elem(kind='duct')}
    if cap:
        n[703] = Elem(kind='cap', family_name='Rectangular Endcap', is_round=False)
    d = build([(700, 0.0, 10.0, 1000.0, 1.0, 'Supply Air'),
               (711, 0.0, 5.0, 300.0 / 0.35, 0.35, 'Supply Air'),
               (712, 0.5, 5.0, 700.0 / 0.35, 0.35, 'Supply Air')])
    for _dr in d.values():
        _dr.size = '12x12"'
    t = {721: (300.0, 'Supply Air', 'T-A'), 722: (700.0, 'Supply Air', 'T-B')}
    return crit([699], ch, d, t, n).get(699, {}).get('Supply Air', {})
_o75, _c75 = _run75(False), _run75(True)
check("case75 uncapped: the last tap is still the end of main turn (2026-10-26 ruling)",
      any('end of main' in r[2] for r in _tap_rows(_o75)), "rows %r" % (_tap_rows(_o75),))
check("case75 capped: the last tap is a tap, not an end of main turn",
      _tap_rows(_c75) and not any('end of main' in r[2] for r in _tap_rows(_c75)),
      "rows %r" % (_tap_rows(_c75),))
check("case75 capped: the tap prices higher than the turn, so the run total rises",
      _c75.get('total_inwc', 0) > _o75.get('total_inwc', 0),
      "capped %r uncapped %r" % (_c75.get('total_inwc'), _o75.get('total_inwc')))

# ── case74: damper in the fan clearance WARNING (F1, Colin 2026-11-04) ───────
# 24 x 24 outlet at 4000 cfm (Le 5.71 ft): duct 741 (l1 ft), accessory 742,
# duct 743 (1.0 ft), radius elbow 744, duct 745 (10 ft) to the diffuser. A
# damper of any kind (balancing, fire / smoke, backdraft) inside the window
# is a WARNING; L and the price are unchanged; past the window, or an
# accessory that is not a damper, says nothing.
def _run74(fam, l1=1.0, cls='Supply Air', conn=SQ24):
    n = {740: _unit({741: conn}), 742: Elem(kind='accessory', family_name=fam),
         744: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
    c = {740: [741], 741: [742], 742: [743], 743: [744], 744: [745], 745: [750]}
    d = build([(741, 0.0, l1, 1000.0, 4.0, cls), (743, 0.0, 1.0, 1000.0, 4.0, cls),
               (745, 0.0, 10.0, 1000.0, 4.0, cls)])
    for _r in d.values():
        _r.size = '24x24"'
    return _fse_run(n, c, d, {750: (4000.0, cls, 'T-74')}, 740, cls=cls)
def _dmp(rec):
    return [t for k, t in rec.get('fse_notes', ()) if k == 'warn' and 'Damper in the fan' in t]
_fd74 = _run74('RJA - Fire Damper - Rectangular')
_sa74 = _run74('Sound Attenuator - Rectangular')
check("case74 a fire damper 1.0 ft from a supply outlet (Le 5.7 ft): one WARNING, exact text",
      len(_dmp(_fd74)) == 1 and _dmp(_fd74)[0] == (
          'WARNING: Damper in the fan clearance at the supply outlet: RJA - Fire Damper - '
          'Rectangular 742 at 1.0 ft is inside Le (%.1f ft). The fan system effect counts '
          'straight duct only, so the damper\'s length is not part of L, and the damper\'s '
          'added loss this close to the fan is not priced, so the true loss is higher than '
          'shown. Best practice: place dampers past Le from a fan connection.' % LE24),
      "notes %r" % (_fd74.get('fse_notes'),))
check("case74 the WARNING is Action required, not RED; L and the fan system effect unchanged",
      duct_perf_console.correction('warn', _dmp(_fd74)[0] if _dmp(_fd74) else '')[0]
      == duct_perf_console.ACTION
      and 742 not in _fd74.get('fse_red_ids', set())
      and _fd74.get('fse_inwc', -1) > 0.0
      and abs(_fd74.get('fse_inwc', -1) - _sa74.get('fse_inwc', -2)) < 1e-12,
      "fse %r / %r red %r" % (_fd74.get('fse_inwc'), _sa74.get('fse_inwc'),
                              _fd74.get('fse_red_ids')))
check("case74 balancing and backdraft dampers inside the window warn too",
      len(_dmp(_run74('Balancing Damper - Rectangular'))) == 1
      and len(_dmp(_run74('Backdraft Damper - Rectangular'))) == 1)
check("case74 an accessory that is not a damper says nothing",
      not _dmp(_sa74), "notes %r" % (_sa74.get('fse_notes'),))
_far74 = _run74('RJA - Fire Damper - Rectangular', l1=7.0)
check("case74 a damper 7.0 ft out, past Le 5.7 ft: no WARNING",
      not _dmp(_far74), "notes %r" % (_far74.get('fse_notes'),))
_in74 = _run74('RJA - Fire Damper - Rectangular', l1=3.0, cls='Return Air')
check("case74 a return inlet damper at 3.0 ft, inside 5 H (10.0 ft): WARNING names the inlet window",
      len(_dmp(_in74)) == 1 and _dmp(_in74)[0].startswith(
          'WARNING: Damper in the fan clearance at the return inlet: RJA - Fire Damper - '
          'Rectangular 742 at 3.0 ft is inside the inlet window, 5 H (10.0 ft).'),
      "notes %r" % (_in74.get('fse_notes'),))
check("case74 a return damper 11.0 ft out, past 5 H: no WARNING",
      not _dmp(_run74('RJA - Fire Damper - Rectangular', l1=11.0, cls='Return Air')))

# 76: near the end of Le the fan system effect falls toward 0, so the straight
#     duct and elbow at their own losses govern (Colin, 2026-11-09). Supply
#     outlet, elbow at L/Le 0.9, a straight duct pricing 0.30 in. wc (well
#     over the fan system effect there): the duct and elbow rows are charged
#     at their own values, the fan system effect row is listed at 0, not
#     charged, and a NOTE says the individual losses govern, with both values.
_v76, _gov76 = FSE.design_velocity(FSE_ON, 1000.0)
_le76, _eq76 = FSE.effective_duct_length_ft(4.0, _v76)
_l76 = 0.9 * _le76
n76 = {3500: _unit({3501: SQ24}),
       3502: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c76 = {3500: [3501], 3501: [3502], 3502: [3503], 3503: [3510]}
d76 = build([(3501, 0.30, _l76, 1000.0, 4.0, 'Supply Air'),
             (3503, 0.01, 10.0, 1000.0, 4.0, 'Supply Air')])
sa76 = _fse_run(n76, c76, d76, {3510: (4000.0, 'Supply Air', 'SD-75')}, 3500)
_f76 = _fse_items(sa76)
_d76 = [i for i in sa76.get('items', ()) if i[0] == 'Duct 3501']
_e76 = [i for i in sa76.get('items', ()) if i[0] == 'Fitting 3502']
_g76 = [t for _k, t in sa76.get('fse_gov_notes', ())]
_off76 = _fse_run(n76, c76, d76, {3510: (4000.0, 'Supply Air', 'SD-75')}, 3500,
                  fan=dict(FSE_ON, on=False))
_own76 = 0.30 + _off76['fitting_inwc']      # the duct and elbow at their own losses
check("case76 L/Le 0.9: own losses govern, their value sits on the FSE row "
      "(Duct losses basis, C '-')",
      len(_f76) == 1 and abs(_f76[0][3] - _own76) < 1e-12
      and abs(sa76.get('fse_inwc', 0) - _own76) < 1e-12
      and _f76[0][4][1] == 'Duct losses: straight duct and elbow (RJA worksheet), '
                          'L %.1f / Le %.1f ft' % (_l76, _le76)
      and _f76[0][4][2] is None, "fse %r rows %r" % (sa76.get('fse_inwc'), _f76))
check("case76 the straight duct and elbow rows are listed at 0.000, in the fan system "
      "effect, out of the straight duct and fitting totals",
      _d76 and _d76[0][3] == 0.0 and _e76 and _e76[0][3] == 0.0
      and _d76[0][4][1] == 'Darcy-Weisbach, in the fan system effect'
      and _e76[0][4][1].endswith(', in the fan system effect')
      and abs(sa76['friction_inwc'] - 0.01) < 1e-12 and abs(sa76['fitting_inwc']) < 1e-12,
      "duct %r elbow %r fric %r" % (_d76, _e76, sa76.get('friction_inwc')))
check("case76 subtotal unchanged: the same as with the fan system effect off",
      abs(sa76['subtotal_inwc'] - _off76['subtotal_inwc']) < 1e-12,
      "on %r off %r" % (sa76.get('subtotal_inwc'), _off76.get('subtotal_inwc')))
check("case76 items add up to the subtotal",
      abs(_items_sum(sa76) - sa76['subtotal_inwc']) < 1e-9)
check("case76 NOTE: individual losses charged instead, with L/Le 0.90 and both values",
      len(_g76) == 1 and _g76[0].startswith('NOTE: supply outlet, elbow 3502 at ')
      and _g76[0].endswith('Straight duct and elbow charged at their own losses, '
                           'no fan system effect.')
      and '(L/Le 0.90)' in _g76[0]
      and ('straight duct and elbow %.3f in. wc.' % _own76) in _g76[0],
      "gov %r" % (_g76,))
check("case76 control: case 32 (elbow at L 0) says the fan system effect is charged, "
      "and its row keeps its plain basis (the NOTE is enough, Colin 2026-11-09)",
      any('Fan system effect charged, the straight duct' in t and '(L/Le 0.00)' in t
          for _k, t in sa32.get('fse_gov_notes', ()))
      and 'charged' not in _fse_items(sa32)[0][1],
      "gov %r rows %r" % (sa32.get('fse_gov_notes'), _fse_items(sa32)))

# 77: the "priced at the default fan velocity" NOTE only where a fan system
#     effect was priced (Colin, 2026-11-09, TB-M10C EF-1): case 31 has no
#     elbow inside Le, so nothing is priced and the NOTE is not said; case 32
#     prices the elbow at L 0 and says it.
def _vel_notes(leaf):
    return [t for _k, t in leaf.get('fse_notes', ())
            if 'priced at the default fan velocity' in t]


_FSE77 = dict(FSE_ON, velocity_basis='default', fan_fpm=2000.0)   # 2000 over 1000 fpm
_none77 = _fse_run(n31, c31, d31, t31, 900, fan=_FSE77)
_elb77 = _fse_run(n32, c32, d32, t32, 920, fan=_FSE77)
check("case77 nothing priced (no elbow within Le): no fan velocity NOTE",
      _none77.get('fse_inwc') == 0.0 and not _vel_notes(_none77),
      "notes %r" % (_none77.get('fse_notes'),))
check("case77 nothing priced: the no fan system effect NOTE is said instead, once",
      [t for _k, t in _none77.get('fse_notes', ())
       if t.startswith('NOTE: no fan system effect at the supply outlet: no elbow within '
                       'Le (')
       and t.endswith('the straight duct and fittings are priced at their own losses.')]
      .__len__() == 1, "notes %r" % (_none77.get('fse_notes'),))
check("case77 priced: no 'no fan system effect' NOTE",
      not any(t.startswith('NOTE: no fan system effect') for _k, t in
              _elb77.get('fse_notes', ())), "notes %r" % (_elb77.get('fse_notes'),))
check("case77 elbow priced at L 0: the fan velocity NOTE is said once",
      _elb77.get('fse_inwc', 0.0) > 0.0 and len(_vel_notes(_elb77)) == 1,
      "notes %r" % (_elb77.get('fse_notes'),))

# 78: the governs NOTE's ratio is on the size the inlet table was read at
#     (Colin, 2026-11-09, TB-M10 return inlet elbow 3652084): an 18 x 16
#     opening, the elbow's side in the plane of its turn H 16 in., L 1.4 ft.
#     ER7-1 reads L/H on 16 in., the RED line says 5 H = 6.7 ft, so the NOTE
#     shall read L/5 H 1.4 / 6.67 = 0.21, not a ratio on the search window
#     (the larger of the opening and the first duct, 18 in.).
RC1816 = {'round': False, 'area_ft2': 2.0, 'd_in': None, 'w_in': 18.0, 'h_in': 16.0}
n78 = {1150: _unit({1151: RC1816}),
       1152: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False,
                  geom={'round': False, 'd_in': None, 'r_in': 24.0, 'w_in': 16.0,
                        'h_in': 18.0, 'angle': 90.0})}
c78 = {1150: [1151], 1151: [1152], 1152: [1153], 1153: [1160]}
d78 = build([(1151, 0.001, 1.4, 1000.0, 2.0, 'Return Air'),
             (1153, 0.0, 10.0, 1000.0, 2.0, 'Return Air')])
for _dr in d78.values():
    _dr.size = '18x16"'
ra78 = _fse_run(n78, c78, d78, {1160: (2000.0, 'Return Air', 'RG-78')}, 1150,
                cls='Return Air')
_g78 = [t for _k, t in ra78.get('fse_gov_notes', ())]
_red78 = [t for k, t in ra78.get('fse_notes', ()) if k == 'red']
check("case78 inlet governs NOTE: L/5 H on the elbow's own H (16 in.): 0.21",
      len(_g78) == 1 and '(L/5 H %.2f, ' % (1.4 / (5.0 * 16.0 / 12.0)) in _g78[0]
      and '(L/5 H 0.21, H 16 in. elbow)' in _g78[0], "gov %r" % (_g78,))
check("case78 the RED line names the same 5 H, 6.7 ft",
      any('6.7 ft (5 H)' in t for t in _red78), "red %r" % (_red78,))

# 78b: TB-M12 L1 inline fan 3720082 (Colin, 2026-11-09): a 12x12 exhaust
#      inlet collar, an 8 in. round elbow at 1.3 ft. The straight duct RED
#      reads H 12 in. off the collar (the opening's smaller side governs over
#      the 8 in. elbow); the fan system effect reads ED7-2 on the elbow's own
#      D 8 in. Each names its size so the two do not read as a contradiction.
RC1212 = {'round': False, 'area_ft2': 1.0, 'd_in': None, 'w_in': 12.0, 'h_in': 12.0}
n78b = {1170: _unit({1171: RC1212}),
        1172: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                   geom={'round': True, 'd_in': 8.0, 'r_in': 12.0, 'w_in': None,
                         'h_in': None, 'angle': 90.0})}
c78b = {1170: [1171], 1171: [1172], 1172: [1173], 1173: [1180]}
d78b = build([(1171, 0.0, 1.3, 400.0, 0.35, 'Exhaust Air'),
              (1173, 0.0, 10.0, 400.0, 0.35, 'Exhaust Air')])
ea78b = _fse_run(n78b, c78b, d78b, {1180: (400.0, 'Exhaust Air', 'EG-78b')}, 1170,
                 cls='Exhaust Air')
_red78b = [t for k, t in ea78b.get('fse_notes', ()) if k == 'red']
_g78b = [t for _k, t in ea78b.get('fse_gov_notes', ())]
check("case78b RED names the collar: (H 12 in., the 12x12 inlet collar), 2 H of the collar",
      len(_red78b) == 1 and '1.3 H (H 12 in., the 12x12 inlet collar)' in _red78b[0]
      and '2.0 ft of straight duct (2 H of the collar); 5.0 ft (5 H)' in _red78b[0],
      "red %r" % (_red78b,))
check("case78b governs NOTE names the elbow: (L/5 D 0.39, D 8 in. elbow)",
      len(_g78b) == 1 and '(L/5 D %.2f, D 8 in. elbow)' % (1.3 / (40.0 / 12.0)) in _g78b[0],
      "gov %r" % (_g78b,))

# 78c: the same with a lossy straight duct, so its own losses govern: the
#      FSE row's Duct losses basis names the elbow's D (Colin, 2026-11-09).
d78c = build([(1171, 0.50, 1.3, 400.0, 0.35, 'Exhaust Air'),
              (1173, 0.0, 10.0, 400.0, 0.35, 'Exhaust Air')])
ea78c = _fse_run(n78b, c78b, d78c, {1180: (400.0, 'Exhaust Air', 'EG-78c')}, 1170,
                 cls='Exhaust Air')
_f78c = _fse_items(ea78c)
check("case78c own losses govern: basis 'L 1.3 / 5 D 3.3 ft (D 8 in., the elbow)'",
      len(_f78c) == 1 and _f78c[0][4][1].startswith('Duct losses: ')
      and _f78c[0][4][1].endswith(', L 1.3 / 5 D 3.3 ft (D 8 in., the elbow)'),
      "rows %r" % (_f78c,))

# 80: exhaust fan inlets take the fan velocity basis (Colin, 2026-11-09):
#     an inline exhaust fan, 12 in. round intake connection, 500 cfm (500
#     fpm at the connection as this harness reads it), a 12 in. round elbow
#     at L/D 2 (2.0 ft). Priced
#     at Pv(2000 fpm) with the fan velocity NOTE. An RTU return inlet of the
#     same geometry stays at its connection velocity, no NOTE.
RD12 = {'round': True, 'area_ft2': math.pi / 4.0, 'd_in': 12.0, 'w_in': None, 'h_in': None}
FSE80 = dict(FSE_ON, fan_type='SWSI', velocity_basis='default', fan_fpm=2000.0)


def _net80(cls):
    n = {4600: _unit({4601: RD12, 4602: RD12}),
         4603: Elem(kind='elbow', family_name='Round Elbow', is_round=True,
                    geom={'round': True, 'd_in': 12.0, 'r_in': 18.0, 'w_in': None,
                          'h_in': None, 'angle': 90.0})}
    c = {4600: [4601, 4602], 4601: [4603], 4603: [4604], 4604: [4610], 4602: [4605],
         4605: []}
    d = build([(4601, 0.0, 2.0, 500.0, math.pi / 4.0, cls),
               (4604, 0.0, 10.0, 500.0, math.pi / 4.0, cls),
               (4602, 0.0, 10.0, 0.0, math.pi / 4.0, cls)])
    for _r in d.values():
        _r.d_h_in = 12.0
    return n, c, d, {4610: (500.0, cls, 'G-80')}


_n80, _c80, _d80, _t80 = _net80('Exhaust Air')
ef80 = crit([4600], _c80, _d80, _t80, _n80, 10.0, fan_effect=FSE80,
            inline_fan_roots={4600})[4600]['Exhaust Air']
_rows80 = _fse_items(ef80)
_vn80 = [t for _k, t in ef80.get('fse_notes', ()) if 'priced at the default fan velocity' in t]
check("case80 inline exhaust fan inlet elbow at L/D 2: priced at Pv(2000 fpm)",
      any(abs((i[2] or 0) - 2000.0) < 1e-6 for i in _rows80)
      and abs(ef80.get('fse_inwc', 0.0) -
              max(i[3] for i in _rows80)) < 1e-12,
      "rows %r fse %r" % (_rows80, ef80.get('fse_inwc')))
check("case80 inline exhaust fan inlet: the fan velocity NOTE, 2000 fpm over 500 fpm",
      len(_vn80) == 1 and _vn80[0] == 'NOTE: fan system effect at the exhaust inlet at '
      'element 4601 priced at the default fan velocity, 2000 fpm, greater than the '
      'connection velocity of 500 fpm.', "notes %r" % (_vn80,))
_nr, _cr, _dr80, _tr = _net80('Return Air')
rtu80 = _fse_run(_nr, _cr, _dr80, _tr, 4600, fan=FSE80, cls='Return Air')
_rr80 = _fse_items(rtu80)
check("case80 RTU return inlet, same geometry: connection velocity (500 fpm), no fan NOTE",
      _rr80 and all(abs((i[2] or 0) - 500.0) < 1e-6 for i in _rr80)
      and not any('priced at the default fan velocity' in t
                  for _k, t in rtu80.get('fse_notes', ())),
      "rows %r" % (_rr80,))
# The cabinet inlet NOTE says why (Colin, 2026-11-09), once, only where a
# fan system effect is priced, and not on the 'connection' basis.
_cab80 = [t for _k, t in rtu80.get('fse_notes', ()) if t.startswith('NOTE') and 'cabinet inlet' in t]
check("case80 RTU return inlet: the cabinet inlet NOTE, once",
      _cab80 == ['NOTE: fan system effect at the return inlet at element 4601 priced at '
                 'the connection velocity, 500 fpm: cabinet inlet, the fan sits past the '
                 'filter and coil.'], "notes %r" % (rtu80.get('fse_notes'),))
check("case80 inline exhaust fan inlet: no cabinet inlet NOTE",
      not any('cabinet inlet' in t for _k, t in ef80.get('fse_notes', ())))
check("case80 inline exhaust fan inlet: no 3 D WARNING",
      not any(t.startswith('WARNING: Elbow') for _k, t in ef80.get('fse_notes', ())))
_rc80 = _fse_run(_nr, _cr, _dr80, _tr, 4600, fan=dict(FSE80, velocity_basis='connection'),
                 cls='Return Air')
check("case80 'connection' basis: no cabinet inlet NOTE, same loss",
      not any(t.startswith('NOTE') and 'cabinet inlet' in t for _k, t in _rc80.get('fse_notes', ()))
      and abs(_rc80.get('fse_inwc', -1) - rtu80.get('fse_inwc', -2)) < 1e-12)
_c80r = (_rows80[0][3] / hvac_graph.velocity_pressure_inwg(2000.0)) if _rows80 else None
_c80c = (_rr80[0][3] / hvac_graph.velocity_pressure_inwg(500.0)) if _rr80 else None
check("case80 same ED7-2 C either way, only Pv changes",
      _c80r is not None and _c80c is not None and abs(_c80r - _c80c) < 1e-9,
      "C %r vs %r" % (_c80r, _c80c))

# 81: the cabinet inlet 3 D WARNING (Colin, 2026-11-09, TB-M15 FCU-2 elbow
#     3726300): a 12x12 return inlet, rect radius elbow turning on its 12 in.
#     side. Under 2 H: RED only. 2 H to 3 H: a WARNING, no color, nothing
#     priced differently; the governs NOTE points to it. Past 3 H: nothing.
#     The same layout on an exhaust fan inlet: no WARNING.
RC1212R = {'round': False, 'area_ft2': 1.0, 'd_in': None, 'w_in': 12.0, 'h_in': 12.0}


def _run81(l_ft, inline=False, fan=None, **kw):
    n = {4700: _unit({4701: RC1212R}),
         4702: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False,
                    geom={'round': False, 'd_in': None, 'r_in': 18.0, 'w_in': 12.0,
                          'h_in': 12.0, 'angle': 90.0})}
    c = {4700: [4701], 4701: [4702], 4702: [4703], 4703: [4710]}
    cls = 'Exhaust Air' if inline else 'Return Air'
    d = build([(4701, 0.0, l_ft, 800.0, 1.0, cls), (4703, 0.0, 10.0, 800.0, 1.0, cls)])
    for _r in d.values():
        _r.d_h_in = 12.0
    t = {4710: (800.0, cls, 'RG-81')}
    if inline:
        return crit([4700], c, d, t, n, 10.0, fan_effect=fan or FSE_ON,
                    inline_fan_roots={4700}, **kw)[4700][cls]
    return _fse_run(n, c, d, t, 4700, cls=cls, fan=fan or FSE_ON, **kw)


def _w81(leaf):
    return [t for k, t in leaf.get('fse_notes', ()) if k == 'warn' and t.startswith('WARNING: Elbow')]


_r81 = _run81(2.4)
check("case81 cabinet inlet elbow at 2.4 ft (2 H 2.0, 3 H 3.0): the 3 H WARNING, not RED",
      _w81(_r81) == ["WARNING: Elbow 4702 is 2.4 ft from the return inlet; provide at least "
                     "3.0 ft of straight duct (3 H, H 12 in.). Under 3 H the elbow causes a "
                     "fan system effect that is not priced (the inlet is priced at its 800 "
                     "fpm opening velocity), so the external static shown is low until the "
                     "elbow is moved."]
      and not _r81.get('fse_red_ids'), "notes %r red %r" % (_r81.get('fse_notes'),
                                                             _r81.get('fse_red_ids')))
_g81 = [t for _k, t in _r81.get('fse_gov_notes', ())]
check("case81 the governs NOTE points to the WARNING",
      len(_g81) == 1 and (
          _g81[0].endswith('charged at their own losses; the fan system effect under 3 H is '
                           'not priced, see the WARNING.')
          or _g81[0].endswith('Fan system effect charged at the opening velocity, the '
                              'straight duct and elbow are not charged again; under 3 H the '
                              'full fan system effect is not priced, see the WARNING.')),
      "gov %r" % (_g81,))
# both governs wordings, straight from the text builder
_gw = FSE.governs_text(True, 1, 2.4, 'return inlet', 0.05, 0.01, 5.0, '5 H',
                       transition=True, warn_dim='H')
_go = FSE.governs_text(False, 1, 2.4, 'return inlet', 0.01, 0.05, 5.0, '5 H',
                       transition=True, warn_dim='D')
check("case81 governs wordings with the WARNING",
      _gw.endswith('Fan system effect charged at the opening velocity, the straight duct, '
                   'transition and elbow are not charged again; under 3 H the full fan '
                   'system effect is not priced, see the WARNING.')
      and _go.endswith('Straight duct, transition and elbow charged at their own losses; the '
                       'fan system effect under 3 D is not priced, see the WARNING.')
      and FSE.governs_text(False, 1, 2.4, 'return inlet', 0.01, 0.05, 5.0, '5 H')
      .endswith('charged at their own losses, no fan system effect.'),
      "%r / %r" % (_gw, _go))
_r81r = _run81(1.9)
check("case81 elbow at 1.9 ft: RED only, no WARNING; Reason says 2 H",
      not _w81(_r81r) and 4702 in (_r81r.get('fse_red_ids') or set())
      and (_r81r.get('fse_red_ducts') or {}).get(4701, ('',))[0]
      == 'Straight duct: elbow 4702 at 1.9 ft, under 2 H',
      "notes %r ducts %r" % (_r81r.get('fse_notes'), _r81r.get('fse_red_ducts')))
_r81p = _run81(3.0)
check("case81 elbow at 3.0 ft: no WARNING, and the governs NOTE does not point to one",
      not _w81(_r81p) and not any('see the WARNING' in t for _k, t in _r81p.get('fse_gov_notes', ())),
      "notes %r" % (_r81p.get('fse_notes'),))
_r81e = _run81(2.4, inline=True)
check("case81 exhaust fan inlet, same layout: no WARNING",
      not _w81(_r81e), "notes %r" % (_r81e.get('fse_notes'),))

# 82: the roof curb (Colin, 2026-11-09): never modeled, its height counted
#     as straight duct at a unit's own connections, length only.
def _curb(inches, base=None):
    return dict(base or FSE_ON, curb_in=float(inches))


def _notes(leaf, kind=None):
    return [t for k, t in leaf.get('fse_notes', ()) if kind is None or k == kind]


# inlet, 12x12 return, rect elbow H 12 (2 H 2.0 ft, 3 H 3.0 ft)
_i0 = _run81(1.9)
_i14 = _run81(1.9, fan=_curb(14))
check("case82 inlet 1.9 ft: RED with no curb; a 14 in. curb (L 3.07 ft) clears 3 H: no RED, no WARNING",
      4702 in (_i0.get('fse_red_ids') or set())
      and not _i14.get('fse_red_ids') and not _w81(_i14),
      "red %r notes %r" % (_i14.get('fse_red_ids'), _notes(_i14)))
_i6 = _run81(1.9, fan=_curb(6))
check("case82 inlet 1.9 ft + 6 in. curb (L 2.4 ft): the 3 H WARNING names the curb and the duct past it",
      _w81(_i6) == ['WARNING: Elbow 4702 is 1.9 ft from the return inlet (plus a 6 in. roof '
                    'curb); provide at least 3.0 ft of straight duct (3 H, H 12 in.), the 6 '
                    'in. curb included (2.5 ft of duct past the curb). Under 3 H the elbow '
                    'causes a fan system effect that is not priced (the inlet is priced at '
                    'its 800 fpm opening velocity), so the external static shown is low until '
                    'the elbow is moved.'] and not _i6.get('fse_red_ids'),
      "notes %r" % (_notes(_i6),))
_i6r = _run81(1.0, fan=_curb(6))
_red6 = _notes(_i6r, 'red')
check("case82 inlet 1.0 ft + 6 in. curb (L 1.5 ft): RED, the duct past the curb asked for",
      len(_red6) == 1 and _red6[0].startswith(
          'RED: Elbow 4702, 1.0 ft from the return inlet (plus a 6 in. roof curb), 1.5 H (H 12 in.')
      and 'Provide at least 2.0 ft of straight duct (2 H of the elbow), the 6 in. curb included '
          '(1.5 ft of duct past the curb); 5.0 ft (5 H, curb included)' in _red6[0]
      and (_i6r.get('fse_red_ducts') or {}).get(4701) ==
      ('Straight duct: elbow 4702 at 1.0 ft + 6 in. curb, under 2 H',
       'Straight duct 1.5 ft past the curb'),
      "red %r ducts %r" % (_red6, _i6r.get('fse_red_ducts')))
check("case82 one curb NOTE per unit",
      _notes(_i6).count('NOTE: roof curb 6 in. counted as straight duct at the unit '
                        'connections (not modeled).') == 1, "notes %r" % (_notes(_i6),))
_i6g = [t for _k, t in _i6.get('fse_gov_notes', ())]
_f6 = _fse_items(_i6)
check("case82 the basis and governs NOTE give L as duct + curb",
      _f6 and 'L 1.9 + 0.5 ft curb' in _f6[0][1]
      and len(_i6g) == 1 and 'elbow 4702 at 1.9 + 0.5 ft curb' in _i6g[0],
      "rows %r gov %r" % (_f6, _i6g))
_c0 = _run81(1.9, fan=_curb(0))
check("case82 curb 0: the same as no curb, no curb NOTE",
      _c0.get('fse_red_ids') == _i0.get('fse_red_ids') and _notes(_c0) == _notes(_i0)
      and abs(_c0['total_inwc'] - _i0['total_inwc']) < 1e-12)
check("case82 the curb adds length only: friction and fittings unchanged",
      abs(_i6['friction_inwc'] - _run81(1.9)['friction_inwc']) < 1e-12)
_e0 = _run81(1.9, inline=True)
_e14 = _run81(1.9, inline=True, fan=_curb(14))
check("case82 exhaust fan inlet: the curb is not counted",
      _notes(_e14) == _notes(_e0) and _e14.get('fse_red_ids') == _e0.get('fse_red_ids')
      and abs(_e14['total_inwc'] - _e0['total_inwc']) < 1e-12, "notes %r" % (_notes(_e14),))
_b14 = _run81(1.9, fan=_curb(14), vav_box_ids={4700})
check("case82 VAV box root: no curb, no curb NOTE",
      not any('roof curb' in t for t in _notes(_b14)), "notes %r" % (_notes(_b14),))
# outlet: SQ24 supply outlet, elbow 1.0 ft out (case 31 layout, Le 5.71 ft)
n82 = {990: _unit({991: SQ24}),
       992: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c82 = {990: [991], 991: [992], 992: [993], 993: [995]}
d82 = build([(991, 0.0, 1.0, 1000.0, 4.0, 'Supply Air'),
             (993, 0.0, 10.0, 1000.0, 4.0, 'Supply Air')])
t82 = {995: (4000.0, 'Supply Air', 'SD-82')}
_o0 = _fse_run(n82, c82, d82, t82, 990)
_o14 = _fse_run(n82, c82, d82, t82, 990, fan=_curb(14))
_ro14 = _notes(_o14, 'red')
check("case82 outlet 1.0 ft + 14 in. curb: still RED under 0.5 Le, named with the curb, lower FSE",
      len(_ro14) == 1 and _ro14[0].startswith('RED: Elbow 992, 1.0 ft from the supply outlet '
                                              '(plus a 14 in. roof curb), L/Le ')
      and 'ft of straight duct, the 14 in. curb included (' in _ro14[0]
      and _o14.get('fse_inwc', 0) < _o0.get('fse_inwc', 0)
      and 'L 1.0 + 1.2 ft curb / Le ' in _fse_items(_o14)[0][1],
      "red %r fse %r / %r rows %r" % (_ro14, _o14.get('fse_inwc'), _o0.get('fse_inwc'),
                                      _fse_items(_o14)))
check("case82 METHOD names the curb used",
      'A roof curb (dialog, default 14 in.; 14 in. this run) is counted as straight duct '
      'at RTU connections, length only.' in FSE.method_lines(_curb(14))[0]
      and 'roof curb' not in FSE.method_lines(FSE_ON)[0])

# 79: taps never reduce the system (Colin, 2026-11-09). A return round shoe
#     tap on a 12 in. round main (ED5-2 stands in, no shoe table on a round
#     main), 100 of 1000 cfm off it. The run carrying on past it (Qs/Qc 0.9,
#     made the index run by a long duct after the tap) reads a negative
#     ED5-2 Cc,s: the worksheet exhaust_tap_main 0.60 x the main Pv is
#     charged, FALLBACK said. The branch side reads a negative Cc,b too and
#     takes exhaust_tap_branch 0.76. A supply run through the same geometry
#     reads a positive SD5-1 Cc,s, kept as read.
RD12A = math.pi * 0.25
ch79 = {799: [700], 700: [701, 702], 701: [711], 711: [721], 702: [722]}
n79 = {700: Elem(kind='duct', tap_pos={701: 2.0}),
       701: Elem(kind='tap', family_name='Round Takeoff - Shoe', is_round=True),
       711: Elem(kind='duct'), 702: Elem(kind='duct')}


def _run79(cls, branch=False):
    d = build([(700, 0.0, 10.0, 1000.0 / RD12A, RD12A, cls),
               (711, 2.0 if branch else 0.0, 5.0, 100.0 / 0.2, 0.2, cls),
               (702, 0.0 if branch else 2.0, 10.0, 900.0 / RD12A, RD12A, cls)])
    for _dr in d.values():
        _dr.size = '12"dia' if _dr.element_id != 711 else '6"dia'
    t = {721: (100.0, cls, 'G-79a'), 722: (900.0, cls, 'G-79b')}
    return crit([799], ch79, d, t, n79).get(799, {}).get(cls, {})


ra79 = _run79('Return Air')
_tp79 = [i for i in ra79.get('items', ()) if i[0] == 'Tap passed']
_fb79 = [u for u in ra79.get('unpriced', ())
         if re.match(r'FALLBACK: tap 701 passed on duct 700, 2001 ED5-2 Cc,s -\d+\.\d\d '
                     r'is not positive; RJA worksheet C 0\.60 used$', u)]
check("case79 return, run past the tap: the worksheet 0.60 x main Pv, not the credit",
      ra79.get('terminal_id') == 722 and len(_tp79) == 1
      and abs(_tp79[0][3] - fitting_tables.DEFAULT_C['exhaust_tap_main']
              * pv(1000.0 / RD12A)) < 1e-12
      and _tp79[0][4][1] == 'RJA worksheet' and abs(_tp79[0][4][2] - 0.60) < 1e-12,
      "end %r tap rows %r" % (ra79.get('terminal_id'), _tp79))
check("case79 FALLBACK: tap 701 passed on duct 700, 2001 ED5-2 Cc,s <negative> is not "
      "positive; RJA worksheet C 0.60 used", len(_fb79) == 1,
      "notes %r" % (ra79.get('unpriced'),))
rb79 = _run79('Return Air', branch=True)
check("case79 return, run through the branch: negative ED5-2 Cc,b takes the worksheet "
      "0.76, FALLBACK said",
      rb79.get('terminal_id') == 721
      and any(re.match(r'FALLBACK: tap 701 into branch off duct 700, 2001 ED5-2 Cc,b '
                       r'-\d+\.\d\d is not positive; RJA worksheet C 0\.76 used$', u)
              for u in rb79.get('unpriced', ()))
      and abs(rb79['fitting_inwc'] - fitting_tables.DEFAULT_C['exhaust_tap_branch']
              * pv(1000.0 / RD12A)) < 1e-12,
      "fit %r notes %r" % (rb79.get('fitting_inwc'), rb79.get('unpriced')))
sa79 = _run79('Supply Air')
_tps79 = [i for i in sa79.get('items', ()) if i[0] == 'Tap passed']
check("case79 supply, a positive SD5-1 Cc,s stays as read, no FALLBACK",
      len(_tps79) == 1 and _tps79[0][4][1].startswith('SD5-1')
      and _tps79[0][4][2] > 0.0
      and abs(_tps79[0][3] - _tps79[0][4][2] * pv(1000.0 / RD12A)) < 1e-12
      and not any(u.startswith('FALLBACK: tap 701') for u in sa79.get('unpriced', ())),
      "tap rows %r notes %r" % (_tps79, sa79.get('unpriced')))

# ── summary ──────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
n_pass = sum(1 for _, ok, _ in results if ok)
print("%d / %d assertions passed" % (n_pass, len(results)))
if n_pass != len(results):
    print("FAILURES:")
    for name, ok, detail in results:
        if not ok:
            print("  - %s :: %s" % (name, detail))
