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


def _transition_geometry(elem):
    return elem.tgeom if elem.tgeom else (None, None)


def _connector_areas_ft2(elem):
    return []      # tap openings unknown: the single branch duct's area is used


def _tap_positions_ft(duct, from_id):
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
hvac_graph.equipment_connection_geometry = _equipment_connection_geometry
hvac_graph.duct_friction_loss_per_100ft = _duct_friction_loss_per_100ft
hvac_graph.takeoff_child_ids       = _takeoff_child_ids
hvac_graph.duct_continues_past     = _duct_continues_past
hvac_graph.is_fitting              = _is_fitting
hvac_graph.is_accessory            = _is_accessory
hvac_graph.is_equipment            = _is_equipment
hvac_graph.fitting_family_name     = _fitting_family_name
hvac_graph.fitting_is_round        = _fitting_is_round
hvac_graph.fitting_angle_deg       = _fitting_angle_deg
hvac_graph.elbow_geometry          = _elbow_geometry
hvac_graph.transition_geometry     = _transition_geometry
hvac_graph.tap_positions_ft        = _tap_positions_ft
hvac_graph.connector_areas_ft2     = _connector_areas_ft2
hvac_graph.transition_areas        = _transition_areas
hvac_graph.velocity_pressure_inwg  = _velocity_pressure_inwg

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
# damper matches no component key, so it must be reported as not priced rather
# than inheriting anyone else's number.
COMP9 = {'diffuser': 0.05, 'damper': 0.25, 'fire_damper': 0.10}

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
expected_comp9 = 0.25 + 0.10 + 0.05     # balancing + fire/smoke + diffuser
check("case9 components == balancing 0.25 + fire/smoke 0.10 + diffuser 0.05",
      abs(sa9['component_inwc'] - expected_comp9) < 1e-9,
      "got %.4f expected %.4f" % (sa9['component_inwc'], expected_comp9))
check("case9 the fire/smoke damper got its OWN value, not the balancing one",
      abs(sa9['component_inwc'] - (0.25 * 2 + 0.05)) > 1e-6,
      "component_inwc=%.4f would be %.4f if it inherited 0.25"
      % (sa9['component_inwc'], 0.25 * 2 + 0.05))
check("case9 the backdraft damper is REPORTED as not priced",
      any('Backdraft' in u for u in sa9['unpriced']),
      "unpriced=%r" % (sa9['unpriced'],))
check("case9 component_keys names exactly what was counted",
      sorted(sa9['component_keys']) == ['damper', 'diffuser', 'fire_damper'],
      "got %r" % (sa9['component_keys'],))

# A component left at 0.00 contributes nothing but is still counted, so the
# report shows it as a deliberate zero rather than a miss.
r9b = crit([900], children9, ducts9, terms9, nodes9, safety_pct=0.0,
           comp_values={'diffuser': 0.05, 'damper': 0.25, 'fire_damper': 0.0})
check("case9 a 0.00 fire damper adds nothing but is still listed",
      abs(r9b[900]['Supply Air']['component_inwc'] - 0.30) < 1e-9
      and 'fire_damper' in r9b[900]['Supply Air']['component_keys'],
      "comp=%.4f keys=%r" % (r9b[900]['Supply Air']['component_inwc'],
                             r9b[900]['Supply Air']['component_keys']))

r9c = crit([900], children9, ducts9, terms9, nodes9, safety_pct=0.0,
           comp_values={'diffuser': 0.0, 'damper': 0.0, 'fire_damper': 0.0})
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
#     and return tap losses floored at zero (Colin, 2026-10-23).
#   Small grille branch: 300 of 3000 cfm (Qb/Qc 0.1) through a 0.35 ft2 branch
#   on a 4 ft2 return main (real Ab/Ac 0.09). Printed Cb -18.00 at Qb/Qc 0.1;
#   at the printed Ab/Ac 0.5, Cc,b = -18 x (0.1/0.5)^2 = -0.72, floored to 0.
ccb29, _l, _n = A.tap_branch_c(False, 'rect', 'rect', True, 0.09, 0.1)
check("case29 ER5-3 is read at its printed Ab/Ac 0.5: Cc,b -0.72, not -18 x (0.1/0.09)^2",
      abs(ccb29 - (-18.0 * (0.1 / 0.5) ** 2)) < 1e-9, "got %r" % ccb29)
ccb29b, _l, _n = A.tap_branch_c(False, 'rect', 'rect', True, 0.09, 0.5)
check("case29 ER5-3 at Qb/Qc 0.5: Cc,b = 0.76 x (0.5/0.5)^2 = 0.76",
      abs(ccb29b - 0.76) < 1e-9, "got %r" % ccb29b)
ch29 = {499: [500], 500: [501], 501: [511], 511: [521]}
n29 = {500: Elem(kind='duct', tap_pos={501: 2.0}),
       501: Elem(kind='tap', family_name='Rectangular Tap - Beveled', is_round=False),
       511: Elem(kind='duct')}
d29 = build([(500, 0.0, 10.0, 750.0, 4.0, 'Return Air'),
             (511, 0.0, 5.0, 857.0, 0.35, 'Return Air')])
t29 = {521: (300.0, 'Return Air', 'RG-1')}
ra29 = crit([499], ch29, d29, t29, n29).get(499, {}).get('Return Air', {})
check("case29 the negative return tap loss is floored at 0 and named in the report",
      abs(ra29.get('fitting_inwc', -1)) < 1e-12
      and any('floored at 0' in u for u in ra29.get('unpriced', ())),
      "fit %r notes %r" % (ra29.get('fitting_inwc'), ra29.get('unpriced')))

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
FSE_ON = {'on': True, 'fan_type': 'DWDI', 'blast_ratio': 0.5}
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
check("case31 straight 20 ft from the outlet (Le 5.71 ft): fse 0, no take-off line, total as with it off",
      on31.get('fse_inwc') == 0.0 and not _fse_items(on31)
      and abs(on31['total_inwc'] - off31['total_inwc']) < 1e-12,
      "on %r off %r" % (on31.get('total_inwc'), off31.get('total_inwc')))

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
check("case32 C 3.80 is GRAY: increase length of duct to 2.9 ft (C under 1.0)",
      any(k == 'gray' and t.startswith('GRAY: System Effect Pressure Loss too high')
          and 'increase length of duct to 2.9 ft' in t
          for k, t in sa32.get('fse_notes', ())), "notes %r" % (sa32.get('fse_notes'),))
check("case32 added before the safety factor: total = (subtotal incl. FSE) x 1.10",
      abs(sa32['subtotal_inwc'] - (sa32['friction_inwc'] + sa32['fitting_inwc'] +
                                   sa32['component_inwc'] + exp32)) < 1e-9
      and abs(sa32['total_inwc'] - sa32['subtotal_inwc'] * 1.10) < 1e-9)
_it32 = _fse_items(sa32)
check("case32 one take-off line 'Fan system effect, supply outlet', SR7-11, C 3.80",
      len(_it32) == 1 and _it32[0][4][0] == 'Fan system effect, supply outlet'
      and _it32[0][4][1].startswith('SR7-11 pos C, L 0.0 / Le 5.7')
      and abs(_it32[0][4][2] - 3.80) < 1e-9, "items %r" % (_it32,))
check("case32 the index run items still add up to the subtotal",
      abs(_items_sum(sa32) - sa32['subtotal_inwc']) < 1e-9)
check("case32 HIGH_C is the named 1.0 threshold", FSE.HIGH_C == 1.0)
check("case32 the GRAY elbow is still priced in the total and listed for the plan",
      sa32['fse_inwc'] > 0.0 and sa32.get('fse_gray_ids') == {921},
      "gray %r" % (sa32.get('fse_gray_ids'),))

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

# 34: a transition 2 ft out, inside Le: flagged, not priced; the elbow after
#     it is not the first disruption.
n34 = {960: _unit({961: SQ24}),
       962: Elem(kind='transition', family_name='Rectangular Transition - Length',
                 is_round=False),
       964: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c34 = {960: [961], 961: [962], 962: [963], 963: [964], 964: [965], 965: [970]}
d34 = build([(961, 0.0, 2.0, 1000.0, 4.0, 'Supply Air'),
             (963, 0.0, 1.0, 1000.0, 4.0, 'Supply Air'),
             (965, 0.0, 5.0, 1000.0, 4.0, 'Supply Air')])
sa34 = _fse_run(n34, c34, d34, {970: (4000.0, 'Supply Air', 'SD-34')}, 960)
check("case34 transition inside Le: flagged, fse 0, the elbow after it ignored",
      sa34.get('fse_inwc') == 0.0 and not _fse_items(sa34)
      and any(k == 'warn' and 'transition 2.0 ft' in t and 'not priced' in t
              for k, t in sa34.get('fse_notes', ())), "notes %r" % (sa34.get('fse_notes'),))
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
      and any(k == 'warn' and t.startswith('WARNING: tap 3.0 ft')
              for k, t in sa34b.get('fse_notes', ())),
      "notes %r" % (sa34b.get('fse_notes'),))

# 35: round return inlet elbow 1 ft from the unit, r/D 1.0: ED7-2 between
#     L/D 0 (1.20) and 2.0 (0.67) at L/D 0.5 = 1.0675, GRAY; plus swirl warning.
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
check("case35 inlet swirl is warned, never priced",
      any(k == 'warn' and 'FSE1,sw' in t and 'not priced' in t for k, t in _n35),
      "notes %r" % (_n35,))
check("case35 C 1.07 is GRAY: increase length of duct to 1.6 ft (C under 1.0 from L/D 0.755)",
      any(k == 'gray' and 'increase length of duct to 1.6 ft' in t for k, t in _n35),
      "notes %r" % (_n35,))
check("case35 take-off line 'Fan system effect, return inlet' on ED7-2",
      [i[4][0] for i in _fse_items(ra35)] == ['Fan system effect, return inlet']
      and _fse_items(ra35)[0][4][1].startswith('ED7-2'))

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

# 38: duct-rooted pick: no fan connection, skipped with a note.
sa38 = _fse_run(n32, c32, d32, t32, 920, duct_rooted_ids={920})
check("case38 duct-rooted pick: fse 0 and the skip is noted",
      sa38.get('fse_inwc') == 0.0
      and any('rooted at a picked duct' in t for _k, t in sa38.get('fse_notes', ())),
      "notes %r" % (sa38.get('fse_notes'),))

# 39: SWSI switches the outlet tables (Ab/Ao 0.5, L 0: SR7-7 3.80); off = none.
sa39 = _fse_run(n32, c32, d32, t32, 920,
                fan={'on': True, 'fan_type': 'SWSI', 'blast_ratio': 0.5})
check("case39 SWSI reads the SR7-5 to SR7-8 tables",
      bool(_fse_items(sa39)) and _fse_items(sa39)[0][4][1].startswith('SR7-7'),
      "items %r" % (_fse_items(sa39),))
sa39b = _fse_run(n32, c32, d32, t32, 920,
                 fan={'on': False, 'fan_type': 'DWDI', 'blast_ratio': 0.5})
check("case39 unticked: no system effect at all",
      sa39b.get('fse_inwc') == 0.0 and not sa39b.get('fse_notes'))

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
_rows42 = [i[4][0] for i in ra35.get('items', ()) if i[4]]
check("case42 the take-off shows the elbow once, as the fan system effect line",
      'Fan system effect, return inlet' in _rows42
      and not any(r.startswith('Round elbow') for r in _rows42), "rows %r" % (_rows42,))
check("case42 items add up to the subtotal",
      abs(_items_sum(ra35) - ra35['subtotal_inwc']) < 1e-9)
check("case42 the report says the elbow is not priced again",
      any('not priced again' in t for _k, t in ra35.get('fse_notes', ())))
check("case42 outlet SR7 stays system effect only: the outlet elbow is still a fitting",
      sa32['fitting_inwc'] > 0.0)

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
check("case44 outlet elbow 16 ft out, inside Le %.2f ft: priced (C %.3f x Pv(3000))" % (le44, c44v),
      abs(sa44.get('fse_inwc', -1) - c44v * pv(3000.0)) < 1e-9,
      "got %r" % sa44.get('fse_inwc'))
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
#     wc of friction), but the other outlet's elbow (C 3.80, GRAY) is still
#     reported, once, named by its element.
n47 = {1600: _unit({1601: SQ24, 1611: SQ24}),
       1601: Elem(kind='elbow', family_name='Rectangular Elbow - Radius', is_round=False)}
c47 = {1600: [1601, 1611], 1601: [1602], 1602: [1620], 1611: [1621]}
d47 = build([(1602, 0.0, 10.0, 1000.0, 4.0, 'Supply Air'),
             (1611, 2.0, 20.0, 1000.0, 4.0, 'Supply Air')])
t47 = {1620: (4000.0, 'Supply Air', 'SD-elbow'), 1621: (4000.0, 'Supply Air', 'SD-long')}
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

# 49: walk-level duct-rooted pick: the root IS a duct (no unit on the run).
c49 = {1800: [1810]}
d49 = build([(1800, 0.1, 10.0, 800.0, 1.0, 'Supply Air')])
sa49 = _fse_run({}, c49, d49, {1810: (800.0, 'Supply Air', 'SD-49')}, 1800,
                duct_rooted_ids={1800})
check("case49 duct root: fse 0 and 'rooted at a picked duct' noted",
      sa49.get('fse_inwc') == 0.0
      and any('rooted at a picked duct' in t for _k, t in sa49.get('fse_notes', ())),
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

# ── summary ──────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
n_pass = sum(1 for _, ok, _ in results if ok)
print("%d / %d assertions passed" % (n_pass, len(results)))
if n_pass != len(results):
    print("FAILURES:")
    for name, ok, detail in results:
        if not ok:
            print("  - %s :: %s" % (name, detail))
