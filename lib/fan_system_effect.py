# -*- coding: utf-8 -*-
"""
fan_system_effect.py  --  fan system effect (FSE) at a unit's duct connections,
from the 2001 ASHRAE Fundamentals Handbook (SI), Ch. 34.

Pure math, no Revit imports, so it runs under CPython
(lib/_test_fan_system_effect.py). The graph walk that finds the first
disruption past a unit connection lives in the Duct Performance script
(_critical_path_loss), next to the walk it shares helpers with.

THE BOOK:
  Eq. 38 (p. 34.14)  Pt = (ps,5 + pv,5) + dp2-5 + FSE2 + (ps,4 + pv,4)
                          + dp4-1 + FSE1 + FSE1,sw
                     subscripts 1 fan inlet, 2 fan outlet, sw swirl.
  Eq. 36 (p. 34.12)  Vo > 13 m/s:   Le = Vo x Ao^0.5 / 4500
  Eq. 37 (p. 34.13)  Vo <= 13 m/s:  Le = Ao^0.5 / 350
                     Le effective duct length, m; Vo duct velocity, m/s;
                     Ao duct area, mm2 (100% recovery, Fig. 11).
  SR7-5 to SR7-8     centrifugal SWSI fan outlet with elbow, positions A to D
  SR7-9 to SR7-12    centrifugal DWDI fan outlet with elbow, positions A to D
                     (p. 34.66 to 34.68), Co by Ab/Ao and L/Le, zero from
                     L/Le 1.00 on.
  ED7-2              fan inlet, centrifugal SWSI, 4 gore round elbow, Co by
                     r/Do and L/Do (p. 34.48).
  ER7-1              fan inlet, centrifugal SWSI, 90 deg smooth radius square
                     elbow, Co by r/H and L/H (p. 34.60).
  The two inlet tables never reach zero: the L/D 5.0 and 10.0 columns are
  equal, so C floors at the 5.0 column.

RJA DECISIONS (Colin, 2026-10-29):
  - Only the FIRST disruption past each unit connection counts, and only
    inside that connection's window. L is the straight duct from the unit
    connection to it. Outlet window: Le (Eq. 36 / 37) for that connection,
    whatever its length. Inlet window: the distance at which the inlet table
    reaches its floor (5 D or 5 H), since the inlet tables never reach zero;
    past it the elbow is an ordinary fitting (Colin, 2026-10-29, second
    ruling; replaces a fixed 15 ft window).
  - Outlet elbow: Co is the WORST of positions A to D (the turn direction
    against the fan rotation is not in the model), read at the blast area
    ratio Ab/Ao (dialog, default 0.5) and L/Le. FSE2 = Co x Pv at the outlet.
  - Inlet elbow: ED7-2 (round) or ER7-1 (rect), at r/D and L/D. FSE1 = Co x
    Pv at the inlet. These Co INCLUDE the elbow's own loss (Colin), so the
    walk does not price that elbow again; the outlet SR7 Co is the system
    effect only and the elbow is priced separately.
  - ED7-2 and ER7-1 are printed for SWSI fans. With DWDI selected they are
    used as a stated ASSUMPTION (method_lines). Inlet swirl, FSE1,sw, has no coefficient in the book
    (p. 34.13: inlet spin changes the fan curve itself), so it is never
    priced, only warned about.
  - A transition or tap inside the window is flagged, never priced (no table).
  - C at or above HIGH_C is GRAY (Colin, 2026-10-29, was RED): "System
    Effect Pressure Loss too high, increase length of duct to X ft", X the
    straight length from the unit connection at which C first drops BELOW
    HIGH_C. It is still priced in the total.
  - ER7-1 is printed for SQUARE elbows (figure, p. 34.60). A non-square
    rectangular inlet elbow is read with H = the side in the plane of the
    turn, stated as an ASSUMPTION (ER7_1_SQUARE_ASSUMPTION).

IronPython 2.7 / pyRevit -- no f-strings, no nonlocal.
"""

import logging
import math

import ashrae_fittings

log = logging.getLogger(__name__)

HIGH_C = 1.0                 # a fan system effect C at or above this is GRAY
LENGTH_STEP_FT = 0.1         # the "increase length of duct to" figure, rounded up
ER7_1_SQUARE_ASSUMPTION = ('ASSUMPTION: ER7-1 is printed for square elbows; the '
                           'rectangular elbow is read with H = the side in the plane '
                           'of the turn.')
FAN_TYPES = ('SWSI', 'DWDI')
DEFAULT_FAN_TYPE = 'DWDI'
DEFAULT_BLAST_RATIO = 0.5
BLAST_RATIO_MIN = 0.4        # the SR7 tables' printed Ab/Ao range
BLAST_RATIO_MAX = 1.0
LE_SPLIT_MPS = 13.0          # Eq. 36 above this outlet velocity, Eq. 37 at or below

# Positions A to D, one table each, per fan type (p. 34.66 to 34.68).
OUTLET_TABLES = {
    'SWSI': (('A', 'SR7-5'), ('B', 'SR7-6'), ('C', 'SR7-7'), ('D', 'SR7-8')),
    'DWDI': (('A', 'SR7-9'), ('B', 'SR7-10'), ('C', 'SR7-11'), ('D', 'SR7-12')),
}
INLET_ROUND = 'ED7-2'
INLET_RECT = 'ER7-1'

_MM2_PER_FT2 = 92903.04
_M_PER_FT = 0.3048
_MPS_PER_FPM = 0.00508


# Connection roles. Today a unit's role comes from its system classification
# (supply = outlet, return / exhaust = inlet). connection_role also takes the
# connector's air direction, for exhaust fans (ceiling and inline) where BOTH
# sides are Exhaust Air: the duct drawing air to the fan is the inlet and the
# discharge duct the outlet. Not wired yet (Colin, 2026-10-29).
OUTLET = 'outlet'
INLET = 'inlet'
_CLASS_ROLE = {'Supply Air': OUTLET, 'Return Air': INLET, 'Exhaust Air': INLET}
_CLASS_WORD = {'Supply Air': 'supply', 'Return Air': 'return', 'Exhaust Air': 'exhaust'}


def connection_role(sys_class, direction=None):
    """OUTLET, INLET or None (not a fan connection this prices).

    direction: the unit connector's air direction as seen from the fan,
    'out' (air leaves the fan through it) or 'in', None when unknown. A known
    direction wins over the classification: that is the hook for exhaust
    fans. The Duct Performance script passes None today, so a unit's role is
    its classification. Outside air and transfer are never priced."""
    if sys_class not in _CLASS_ROLE:
        return None
    if direction == 'out':
        return OUTLET
    if direction == 'in':
        return INLET
    return _CLASS_ROLE[sys_class]


def where_text(sys_class, role):
    """'supply outlet', 'return inlet', 'exhaust inlet' (and, for an exhaust
    fan later, 'exhaust outlet')."""
    return '%s %s' % (_CLASS_WORD.get(sys_class, (sys_class or '').lower()), role)


def default_settings():
    """The dialog defaults: on, DWDI, blast area ratio 0.5."""
    return {'on': True, 'fan_type': DEFAULT_FAN_TYPE,
            'blast_ratio': DEFAULT_BLAST_RATIO}


def parse_blast_ratio(text):
    """(value, None) or (None, error text) for the dialog's Ab/Ao box."""
    try:
        v = float(text)
        if math.isnan(v) or math.isinf(v):
            raise ValueError('not finite')
    except (ValueError, TypeError):
        return None, 'must be a number'
    if not (BLAST_RATIO_MIN <= v <= BLAST_RATIO_MAX):
        return None, 'must be between %.1f and %.1f' % (BLAST_RATIO_MIN, BLAST_RATIO_MAX)
    return v, None


def is_high(c):
    """True when a fan system effect C reaches HIGH_C (reported GRAY)."""
    return c is not None and float(c) >= HIGH_C - 1e-9


def _ceil_step(ft):
    """ft rounded UP to LENGTH_STEP_FT, so the figure given is past the edge.
    Any length above zero is at least one step: a C of exactly HIGH_C at
    L 0 needs some straight duct, never '0.0 ft'."""
    n = math.ceil(float(ft) / LENGTH_STEP_FT - 1e-9)
    if float(ft) > 0.0:
        n = max(n, 1)
    return round(n * LENGTH_STEP_FT, 6)


def _first_below(c_of, x_max, iters=60):
    """Smallest x in [0, x_max] where c_of(x) < HIGH_C, or None when even
    x_max does not get under it. c_of is a table read, which falls (does not
    rise) with straight length, so bisection finds the edge."""
    if c_of(0.0) < HIGH_C:
        return 0.0
    if c_of(x_max) >= HIGH_C:
        return None
    lo, hi = 0.0, float(x_max)
    for _ in range(iters):
        mid = 0.5 * (lo + hi)
        if c_of(mid) < HIGH_C:
            hi = mid
        else:
            lo = mid
    return hi


def outlet_fix_ft(fan_type, blast_ratio, le_ft):
    """Straight length from the outlet at which the worst of positions A to
    D first drops below HIGH_C, rounded up; 0.0 when it never reaches it."""
    x = _first_below(lambda r: outlet_c(fan_type, blast_ratio, r)[0], 1.0)
    return _ceil_step(x * float(le_ft)) if x is not None else None


def inlet_fix_ft(is_round, r_over, size_in):
    """Straight length from the inlet at which ED7-2 / ER7-1 first drops below
    HIGH_C, rounded up; None if the table's floor is itself that high."""
    key = INLET_ROUND if is_round else INLET_RECT
    x = _first_below(lambda y: inlet_c(is_round, r_over, y)[0], _cols(key)[-1])
    return _ceil_step(x * float(size_in) / 12.0) if x is not None else None


# ── effective duct length ────────────────────────────────────────────────────

def effective_duct_length_ft(area_ft2, v_fpm):
    """(Le ft, 'Eq. 36' or 'Eq. 37') for an outlet of area_ft2 at v_fpm.

    The book works in SI: Ao in mm2, Vo in m/s, Le in m."""
    ao_mm2 = float(area_ft2) * _MM2_PER_FT2
    vo = float(v_fpm) * _MPS_PER_FPM
    if vo > LE_SPLIT_MPS:
        le_m, eq = vo * math.sqrt(ao_mm2) / 4500.0, 'Eq. 36'
    else:
        le_m, eq = math.sqrt(ao_mm2) / 350.0, 'Eq. 37'
    return le_m / _M_PER_FT, eq


# ── table helpers ────────────────────────────────────────────────────────────

def _cols(key):
    return ashrae_fittings.table(key)['cols']['values']


def _rows(key):
    return ashrae_fittings.table(key)['rows']['values']


def zero_ratio(key):
    """Smallest printed column (L/Le) from which every row reads 0, or None."""
    t = ashrae_fittings.table(key)
    cols = t['cols']['values']
    for j, x in enumerate(cols):
        if all(all(row[k] == 0.0 for k in range(j, len(cols))) for row in t['C']):
            return x
    return None


def inlet_window_ft(is_round, size_in):
    """How far past an inlet a disruption is looked for: the largest floor
    column over every row of its table (5 D on ED7-2, 5 H on ER7-1) x the
    connection size. Past it no row of the table is still falling."""
    key = INLET_ROUND if is_round else INLET_RECT
    worst = max(floor_ratio(key, r) for r in _rows(key))
    return worst * float(size_in) / 12.0


def floor_ratio(key, row_value):
    """Smallest printed column (L/D) from which the row at row_value stops
    falling: its value equals the last column's. The rows either side of an
    interpolated row_value are both checked, the later column governs."""
    t = ashrae_fittings.table(key)
    rows, cols = t['rows']['values'], t['cols']['values']
    i0, i1, _f, _cl = ashrae_fittings._bracket(rows, float(row_value))
    first = 0
    for i in (i0, i1):
        row = t['C'][i]
        j = len(cols) - 1
        while j > 0 and row[j - 1] == row[-1]:
            j -= 1
        first = max(first, j)
    return cols[first]


# ── coefficients ─────────────────────────────────────────────────────────────

def outlet_c(fan_type, blast_ratio, l_over_le):
    """Co at the fan outlet: the worst of positions A to D.

    Returns (Co, position, table key, notes). An unknown fan type takes the
    default (DWDI) and says so. L/Le past the last printed column reads the
    last column (0), not a clamp."""
    notes = []
    if fan_type not in OUTLET_TABLES:
        notes.append('fan type %r unknown, %s used' % (fan_type, DEFAULT_FAN_TYPE))
        fan_type = DEFAULT_FAN_TYPE
    worst = None
    for pos, key in OUTLET_TABLES[fan_type]:
        x = min(max(float(l_over_le), 0.0), _cols(key)[-1])
        c, n = ashrae_fittings.lookup(key, float(blast_ratio), x)
        if worst is None or c > worst[0]:
            worst = (c, pos, key, n)
    c, pos, key, n = worst
    return c, pos, key, notes + list(n)


def inlet_c(is_round, r_over, l_over):
    """Co at the fan inlet from ED7-2 (round) or ER7-1 (rect).

    r_over: r/D or r/H, None when the elbow radius is unreadable, which takes
    the smallest printed ratio (the highest C, conservative). L/D past the
    last printed column reads the last column (the floor), not a clamp.
    Returns (Co, table key, r ratio used, notes)."""
    key = INLET_ROUND if is_round else INLET_RECT
    notes = []
    rows = _rows(key)
    if r_over is None:
        r_over = rows[0]
        notes.append('%s: elbow radius unreadable, smallest printed r/%s %.2f used '
                     '(highest C, conservative)' % (key, 'D' if is_round else 'H', r_over))
    x = min(max(float(l_over), 0.0), _cols(key)[-1])
    c, n = ashrae_fittings.lookup(key, float(r_over), x)
    return c, key, float(r_over), notes + list(n)


# ── one connection ───────────────────────────────────────────────────────────

def outlet_effect(fan_type, blast_ratio, area_ft2, cfm, l_ft):
    """The fan outlet record for an elbow l_ft from the outlet of area_ft2.

    Returns {'side': 'outlet', 'c', 'table', 'position', 'v_fpm', 'le_ft',
    'eq', 'l_ft', 'ratio' (L/Le), 'blast_ratio', 'high', 'fix_ft', 'notes'}.
    Pv is the caller's (hvac_graph's standard air), so the loss is c x
    Pv(v_fpm). fix_ft: the straight length from the outlet at which C first
    drops below HIGH_C (outlet_fix_ft)."""
    v = float(cfm) / float(area_ft2)
    le, eq = effective_duct_length_ft(area_ft2, v)
    ratio = float(l_ft) / le if le > 0 else 0.0
    c, pos, key, notes = outlet_c(fan_type, blast_ratio, ratio)
    rec = {'side': 'outlet', 'c': c, 'table': key, 'position': pos,
           'v_fpm': v, 'le_ft': le, 'eq': eq, 'l_ft': float(l_ft),
           'ratio': ratio, 'blast_ratio': float(blast_ratio), 'high': is_high(c),
           'fix_ft': outlet_fix_ft(fan_type, blast_ratio, le) if is_high(c) else None,
           'notes': notes}
    log.info('fan_system_effect outlet: %s pos %s Ab/Ao %.2f L %.2f ft Le %.2f ft '
             '(%s) C %.2f at %.0f fpm', key, pos, float(blast_ratio), float(l_ft),
             le, eq, c, v)
    return rec


def inlet_effect(is_round, size_in, r_in, area_ft2, cfm, l_ft):
    """The fan inlet record for an elbow l_ft from the inlet.

    size_in: the elbow's D (round) or H (rect, the side in the plane of the
    turn). r_in: centerline radius, None when unreadable (conservative row).
    area_ft2: the inlet connection area, for the velocity pressure.
    Returns {'side': 'inlet', 'c', 'table', 'r_ratio', 'size_in', 'v_fpm',
    'l_ft', 'ratio' (L/D), 'high', 'fix_ft', 'floor_ft', 'floor_ratio',
    'floor_c', 'within', 'notes'}. floor_ft is the straight length at which
    C reaches the table's floor, floor_c that floor (never 0 on these
    tables). within False: the elbow is at or past floor_ft, outside the
    inlet window, and is an ordinary fitting; the caller does not price it
    here. fix_ft: where C first drops below HIGH_C (inlet_fix_ft)."""
    size_in = float(size_in)
    r_over = (float(r_in) / size_in) if r_in else None
    l_over = float(l_ft) * 12.0 / size_in
    c, key, r_used, notes = inlet_c(is_round, r_over, l_over)
    fr = floor_ratio(key, r_used)
    floor_c, _n = ashrae_fittings.lookup(key, r_used, fr)
    v = float(cfm) / float(area_ft2)
    floor_ft = fr * size_in / 12.0
    rec = {'side': 'inlet', 'c': c, 'table': key, 'r_ratio': r_used,
           'size_in': size_in, 'v_fpm': v, 'l_ft': float(l_ft), 'ratio': l_over,
           'high': is_high(c),
           'fix_ft': inlet_fix_ft(is_round, r_used, size_in) if is_high(c) else None,
           'floor_ft': floor_ft, 'floor_ratio': fr, 'floor_c': floor_c,
           'within': float(l_ft) < floor_ft - 1e-9, 'notes': notes}
    log.info('fan_system_effect inlet: %s r/%s %.2f L %.2f ft (%.2f x %.1f in.) '
             'C %.2f at %.0f fpm', key, 'D' if is_round else 'H', r_used,
             float(l_ft), l_over, size_in, c, v)
    return rec


# ── report text ──────────────────────────────────────────────────────────────

def basis_text(rec):
    """The take-off's Table / basis cell, kept short for the sheet's clipped
    cell (Ab/Ao is in the METHOD line): 'SR7-10 pos B, L 2.0 / Le 5.7 ft' or
    'ED7-2 r/D 1.50, L 1.0 ft, D 24 in'."""
    if rec['side'] == 'outlet':
        return '%s pos %s, L %.1f / Le %.1f ft' % (
            rec['table'], rec['position'], rec['l_ft'], rec['le_ft'])
    dim = 'D' if rec['table'] == INLET_ROUND else 'H'
    return '%s r/%s %.2f, L %.1f ft, %s %.0f in' % (
        rec['table'], dim, rec['r_ratio'], rec['l_ft'], dim, rec['size_in'])


def high_text(rec, where):
    """The GRAY line for the engineer (Colin, 2026-10-29), in the Duct
    Performance reason style. where: 'supply outlet', 'return inlet'."""
    if rec.get('fix_ft') is None:
        fix = 'no straight length in the table brings C under %.1f' % HIGH_C
    else:
        fix = 'increase length of duct to %.1f ft' % rec['fix_ft']
    if rec['side'] == 'outlet':
        basis = '%s, worst of positions A to D, Le %.1f ft' % (rec['table'], rec['le_ft'])
    else:
        basis = '%s r/%s %.2f' % (rec['table'], 'D' if rec['table'] == INLET_ROUND
                                  else 'H', rec['r_ratio'])
    return ('GRAY: System Effect Pressure Loss too high at the %s, %s (C %.2f at '
            'L %.1f ft, %s; C %.1f or more).' % (
                where, fix, rec['c'], rec['l_ft'], basis, HIGH_C))


def swirl_text(where, l_ft):
    """The inlet swirl warning: never priced, the book has no coefficient."""
    return ('WARNING: elbow %.1f ft from the %s. Inlet swirl (FSE1,sw) is not '
            'priced: the 2001 Fundamentals gives no coefficient for it (p. 34.13, '
            'inlet spin changes the fan curve itself).' % (float(l_ft), where))


def disruption_text(kind, where, l_ft, limit_ft, limit_name):
    """A transition, tap or other fitting inside the window: flagged, not
    priced. limit_name: 'Le' (outlet) or 'the inlet window'."""
    return ('WARNING: %s %.1f ft from the %s, inside %s (%.1f ft). A fan system '
            'effect disruption with no 2001 table, not priced.' % (
                kind, float(l_ft), where, limit_name, float(limit_ft)))


def method_lines(settings):
    """(METHOD sentence, [ASSUMPTION lines]) for the external static report,
    console and drafting view alike. Off: a not-included sentence, no
    assumptions."""
    if not settings or not settings.get('on'):
        return 'Fan system effect not included (unticked).', []
    fan = settings.get('fan_type') or DEFAULT_FAN_TYPE
    tables = OUTLET_TABLES.get(fan, OUTLET_TABLES[DEFAULT_FAN_TYPE])
    method = ('Fan system effect (2001 Fundamentals Eq. 38): the first disruption '
              'past each unit connection. Outlet: within Le (Eq. 36 or 37), an elbow '
              'is priced from the %s tables %s to %s, worst of positions A to D, at '
              'Ab/Ao %.2f; the elbow is also priced as a fitting. Inlet: within the '
              'table floor distance (5 D or 5 H), an elbow is priced from ED7-2 '
              '(round) or ER7-1 (rect), whose C includes the elbow, so it is not '
              'priced again. Added before the safety factor. C %.1f or more is GRAY, '
              'with the straight length that brings it under. '
              'Inlet swirl not priced (no 2001 coefficient); transitions and taps '
              'flagged, not priced.' % (
                  fan, tables[0][1], tables[-1][1],
                  float(settings.get('blast_ratio', DEFAULT_BLAST_RATIO)), HIGH_C))
    assumptions = []
    if fan != 'SWSI':
        assumptions.append(
            'ASSUMPTION: the fan inlet tables ED7-2 and ER7-1 are printed for SWSI '
            'fans only; they are used for the %s fan selected.' % fan)
    return method, assumptions
