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
    EXT_ROOT, "Developer Tools.tab", "HVAC Tools.panel",
    "Duct Velocity.pushbutton", "script.py")
LIB_DIR = os.path.join(EXT_ROOT, "lib")

# fitting_tables.py has no Revit imports -> import the real thing.
sys.path.insert(0, LIB_DIR)
import fitting_tables  # noqa: E402  (real module, pure Python)

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
    def __init__(self, kind, family_name=None, is_round=True, down_area=None):
        self.kind = kind
        self.family_name = family_name
        self.is_round = is_round
        self.down_area = down_area


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


def _fitting_family_name(elem):
    return elem.family_name


def _fitting_is_round(elem):
    return elem.is_round


def _transition_areas(elem, up_area):
    return (up_area, elem.down_area)


def _velocity_pressure_inwg(v_fpm):
    return (v_fpm / 4007.7) ** 2


hvac_graph = types.ModuleType('hvac_graph')
hvac_graph.takeoff_child_ids       = _takeoff_child_ids
hvac_graph.duct_continues_past     = _duct_continues_past
hvac_graph.is_fitting              = _is_fitting
hvac_graph.is_accessory            = _is_accessory
hvac_graph.fitting_family_name     = _fitting_family_name
hvac_graph.fitting_is_round        = _fitting_is_round
hvac_graph.transition_areas        = _transition_areas
hvac_graph.velocity_pressure_inwg  = _velocity_pressure_inwg

ns = {'hvac_graph': hvac_graph, 'fitting_tables': fitting_tables}
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
#         "main past take-off" coefficient (0.28, C_SUPPLY_TAP_MAIN) must be
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
check("case3 fitting_inwc == 2*0.28*Pv(1200) + 0.98*Pv(1200)",
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
check("case5 subtotal_inwc == friction_inwc + fitting_inwc",
      abs(sa5.get('subtotal_inwc', -1)
          - (sa5.get('friction_inwc', 0) + sa5.get('fitting_inwc', 0))) < 1e-9,
      "got subtotal=%r friction=%r fitting=%r"
      % (sa5.get('subtotal_inwc'), sa5.get('friction_inwc'), sa5.get('fitting_inwc')))
check("case5 total_inwc == subtotal_inwc * 1.10",
      abs(sa5.get('total_inwc', -1) - sa5.get('subtotal_inwc', -1) * 1.10) < 1e-9,
      "got total=%r subtotal=%r" % (sa5.get('total_inwc'), sa5.get('subtotal_inwc')))


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


# ── case 9: component drops (diffuser + balancing dampers) ──────────────────
#
#   root(900) -> duct 901 -> BALANCING DAMPER 902 -> FIRE DAMPER 903
#                         -> duct 904 -> terminal 920
#
# diffuser_drop = 0.05 applied ONCE at the terminal.
# damper_drop   = 0.25 applied ONCE, for the balancing damper only; the fire
#                 damper must NOT inherit it and must be reported as uncounted.
children9 = {900: [901], 901: [902], 902: [903], 903: [904], 904: [920]}
ducts9 = build([(901, 0.10, 10.0, 800, 1.0, 'Supply Air'),
                (904, 0.10, 10.0, 800, 1.0, 'Supply Air')])
nodes9 = {
    901: Elem('duct'),
    902: Elem('accessory', family_name='Balancing Damper - Round'),
    903: Elem('accessory', family_name='RJA - Fire Damper - Round'),
    904: Elem('duct'),
}
terms9 = {920: (400.0, 'Supply Air', 'SD-9')}

r9 = crit([900], children9, ducts9, terms9, nodes9,
          safety_pct=0.0, diffuser_drop=0.05, damper_drop=0.25)
sa9 = r9[900]['Supply Air']
expected_comp9 = 0.25 + 0.05          # one balancing damper + one diffuser
check("case9 component_inwc == 1 balancing damper (0.25) + 1 diffuser (0.05)",
      abs(sa9['component_inwc'] - expected_comp9) < 1e-9,
      "got %.4f expected %.4f" % (sa9['component_inwc'], expected_comp9))
check("case9 the FIRE damper did not silently inherit the balancing-damper drop",
      abs(sa9['component_inwc'] - (0.25 * 2 + 0.05)) > 1e-6,
      "component_inwc=%.4f would be %.4f if both dampers counted"
      % (sa9['component_inwc'], 0.25 * 2 + 0.05))
check("case9 the uncounted fire damper is REPORTED, not silently dropped",
      any('Fire Damper' in u for u in sa9['unpriced']),
      "unpriced=%r" % (sa9['unpriced'],))
check("case9 subtotal now includes components",
      abs(sa9['subtotal_inwc'] -
          (sa9['friction_inwc'] + sa9['fitting_inwc'] + sa9['component_inwc'])) < 1e-9,
      "subtotal=%.4f" % sa9['subtotal_inwc'])
# with both drops at 0 the components must vanish entirely
r9b = crit([900], children9, ducts9, terms9, nodes9,
           safety_pct=0.0, diffuser_drop=0.0, damper_drop=0.0)
check("case9 zeroed component drops contribute nothing",
      abs(r9b[900]['Supply Air']['component_inwc']) < 1e-12,
      "got %r" % r9b[900]['Supply Air']['component_inwc'])


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


# ── summary ──────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
n_pass = sum(1 for _, ok, _ in results if ok)
print("%d / %d assertions passed" % (n_pass, len(results)))
if n_pass != len(results):
    print("FAILURES:")
    for name, ok, detail in results:
        if not ok:
            print("  - %s :: %s" % (name, detail))
