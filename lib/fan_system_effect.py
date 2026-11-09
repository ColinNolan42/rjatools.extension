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
    Pv at the inlet.
  - Charged once, outlets AND inlets (Colin, 2026-11-04, c0333e9; replaces
    the 10-29 outlet ruling that priced the elbow separately, which Colin
    ruled wrong 2026-11-09): the straight duct, any transition on it and the
    first elbow are one group, charged as the LARGER of the fan system effect
    and their own duct and fitting losses. Own losses generally govern only
    near Le, where the fan system effect is small. A NOTE says which governed
    (governs_text).
  - ED7-2 and ER7-1 are printed for SWSI fans. With DWDI selected they are
    used as a stated ASSUMPTION (method_lines). Inlet swirl, FSE1,sw, has no
    coefficient in the book (p. 34.13: inlet spin changes the fan curve
    itself), so it is never priced. Two elbows turning in different planes
    within SWIRL_DH hydraulic diameters of a unit connection (the
    close-coupled limit, p. 34.11; the Fig. 12 spin connections at an inlet)
    are reported RED and both elbows drawn red, supply, return and exhaust
    alike (Colin, 2026-11-02, replaces "not reported").
  - L counts from the unit through the curb, straight duct and any
    transitions, each transition's own length included, to the FIRST elbow
    (Colin, 2026-11-02: the system effect checks that the straight duct at
    the connection is long enough, and a curb adapter is part of it).
    Transitions are priced as ordinary fittings by the walk. A tap, split
    or other fitting inside the window is flagged, never priced (no table).
  - A rect mitered inlet elbow is read on the ER7-1 r/H 1.0 row (Colin,
    2026-11-02), not the conservative unreadable radius row.
  - A duct pick with no unit takes the picked duct's unit side end as the
    unit outlet or inlet (Colin, 2026-11-02). Duct Performance no longer
    offers a duct pick (Colin, 2026-11-03); the walker keeps the case.
  - A loss (C x Pv) at or above HIGH_LOSS_INWC is GRAY: "System Effect
    Pressure Loss too high, increase length of duct to X ft", X the straight
    length from the unit connection at which the loss first drops BELOW it.
    It is still priced in the total. Colin, 2026-11-02: judged on the loss,
    not on C (was C 1.0 from 2026-10-29); 0.058 in. wc is reasonable, 0.5 is
    not, which takes a very short straight into a small, fast elbow.
  - ER7-1 is printed for SQUARE elbows (figure, p. 34.60). A non-square
    rectangular inlet elbow is read with H = the side in the plane of the
    turn, stated as an ASSUMPTION (ER7_1_SQUARE_ASSUMPTION).

RJA DECISIONS (Colin, 2026-11-03):
  - Velocity basis, a dialog choice. 'default' (the default, conservative):
    the velocity is the LARGER of a default fan velocity (DEFAULT_FAN_FPM)
    and the connection velocity, cfm / the unit connector area. A unit's
    opening is often smaller than the duct it feeds, and the fan discharge
    faster than the duct. Default 2000 fpm (Colin, 2026-11-03; Greenheck's
    simulator is fixed at 2500 fpm).
    'connection': the connection velocity alone. The velocity sets Pv, the
    loss, the GRAY check and the outlet Le. The transitions passed between
    the unit and the first disruption are priced at the same velocity.
    Outlets only (Colin, 2026-11-04): an inlet always uses its connection
    velocity.
  - Straight duct check, RED for the engineer, still priced as above
    (Colin, 2026-11-03, .claude/docs/Duct Velocity/FAN_SYSTEM_EFFECT_PLAN.md
    section C): an elbow under OUTLET_RED_RATIO x Le of a unit OUTLET (Le
    Eq. 36 / 37; Eq. 37 is the 2.5 D rule up to about 2500 fpm), or under
    INLET_RED_D x D of a unit INLET, where the loss is steep. D is the larger
    of the unit opening and the first duct (round D). On a rect inlet elbow
    H is its side in the plane of the turn, not less than the opening's
    smaller side (Colin, 2026-11-04). The straight duct between the unit
    and the elbow is RED with it. The
    note gives the RED line length and the no loss length: 1.0 Le at an
    outlet, INLET_FLOOR_D x D at an inlet (where the tables stop falling).
    Elbows past the RED line are still priced. Replaces 5 D at both sides.

IronPython 2.7 / pyRevit -- no f-strings, no nonlocal.
"""

import logging
import math

import ashrae_fittings

log = logging.getLogger(__name__)

HIGH_LOSS_INWC = 0.5         # a fan system effect loss at or above this is GRAY
LENGTH_STEP_FT = 0.1         # the "increase length of duct to" figure, rounded up
MITERED_R_OVER_H = 1.0       # ER7-1 row a rect mitered inlet elbow is read on
ER7_1_SQUARE_ASSUMPTION = ('ASSUMPTION: ER7-1 is printed for square elbows; the '
                           'rectangular elbow is read with H = the side in the plane '
                           'of the turn.')
FAN_TYPES = ('SWSI', 'DWDI')
DEFAULT_FAN_TYPE = 'DWDI'
DEFAULT_BLAST_RATIO = 0.5
BLAST_RATIO_MIN = 0.4        # the SR7 tables' printed Ab/Ao range
BLAST_RATIO_MAX = 1.0
VELOCITY_BASES = ('default', 'connection')
DEFAULT_VELOCITY_BASIS = 'default'
DEFAULT_FAN_FPM = 2000.0     # default fan velocity, firm value (Colin, 2026-11-03)
OUTLET_RED_RATIO = 0.5       # an elbow under 0.5 Le of a unit outlet is RED
INLET_RED_D = 2.0            # an elbow under 2 D of a unit inlet is RED
INLET_FLOOR_D = 5.0          # where the inlet tables stop falling (the note)
LE_SPLIT_MPS = 13.0          # Eq. 36 above this outlet velocity, Eq. 37 at or below
SWIRL_DH = 6.0               # inlet swirl window, hydraulic diameters (p. 34.11)
SWIRL_SAME_PLANE = 0.1       # |n1 x n2| under this: two turn planes are the same

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
    """The dialog defaults: on, DWDI, blast area ratio 0.5, the default fan
    velocity basis at DEFAULT_FAN_FPM."""
    return {'on': True, 'fan_type': DEFAULT_FAN_TYPE,
            'blast_ratio': DEFAULT_BLAST_RATIO,
            'velocity_basis': DEFAULT_VELOCITY_BASIS, 'fan_fpm': DEFAULT_FAN_FPM}


def parse_fan_fpm(text):
    """(value, None) or (None, error text) for the dialog's fan velocity box."""
    try:
        v = float(text)
        if math.isnan(v) or math.isinf(v):
            raise ValueError('not finite')
    except (ValueError, TypeError):
        return None, 'must be a number'
    if v <= 0:
        return None, 'must be greater than 0'
    return v, None


def design_velocity(settings, v_conn_fpm):
    """(fpm used, True when the default fan velocity governs).

    'default' basis: the larger of the settings' fan_fpm and the connection
    velocity. 'connection' basis, or no settings: the connection velocity."""
    v = float(v_conn_fpm)
    if settings and settings.get('velocity_basis', DEFAULT_VELOCITY_BASIS) == 'default':
        f = float(settings.get('fan_fpm', DEFAULT_FAN_FPM))
        if f > v:
            return f, True
    return v, False


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


def std_pv_inwc(v_fpm):
    """Velocity pressure of standard air, in. wc. The default when the caller
    passes no pv_of (Duct Performance passes hvac_graph's, which follows the
    dialog's air density)."""
    return (float(v_fpm) / 4005.0) ** 2


def is_high(loss_inwc):
    """True when a fan system effect loss reaches HIGH_LOSS_INWC (GRAY)."""
    return loss_inwc is not None and float(loss_inwc) >= HIGH_LOSS_INWC - 1e-9


def c_limit(pv_inwc):
    """The C at which the loss reaches HIGH_LOSS_INWC at this Pv."""
    return HIGH_LOSS_INWC / pv_inwc if pv_inwc > 0 else float('inf')


def _ceil_step(ft):
    """ft rounded UP to LENGTH_STEP_FT, so the figure given is past the edge.
    Any length above zero is at least one step: a loss of exactly
    HIGH_LOSS_INWC at L 0 needs some straight duct, never '0.0 ft'."""
    n = math.ceil(float(ft) / LENGTH_STEP_FT - 1e-9)
    if float(ft) > 0.0:
        n = max(n, 1)
    return round(n * LENGTH_STEP_FT, 6)


def _first_below(c_of, x_max, c_lim, iters=60):
    """Smallest x in [0, x_max] where c_of(x) < c_lim, or None when even
    x_max does not get under it. c_of is a table read, which falls (does not
    rise) with straight length, so bisection finds the edge."""
    if c_of(0.0) < c_lim:
        return 0.0
    if c_of(x_max) >= c_lim:
        return None
    lo, hi = 0.0, float(x_max)
    for _ in range(iters):
        mid = 0.5 * (lo + hi)
        if c_of(mid) < c_lim:
            hi = mid
        else:
            lo = mid
    return hi


def outlet_fix_ft(fan_type, blast_ratio, le_ft, c_lim):
    """Straight length from the outlet at which the worst of positions A to
    D first drops below c_lim, rounded up; None when it never does."""
    x = _first_below(lambda r: outlet_c(fan_type, blast_ratio, r)[0], 1.0, c_lim)
    return _ceil_step(x * float(le_ft)) if x is not None else None


def inlet_fix_ft(is_round, r_over, size_in, c_lim):
    """Straight length from the inlet at which ED7-2 / ER7-1 first drops below
    c_lim, rounded up; None if the table's floor is itself that high."""
    key = INLET_ROUND if is_round else INLET_RECT
    x = _first_below(lambda y: inlet_c(is_round, r_over, y)[0], _cols(key)[-1], c_lim)
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

def outlet_effect(fan_type, blast_ratio, area_ft2, cfm, l_ft, pv_of=None, v_fpm=None):
    """The fan outlet record for an elbow l_ft from the outlet of area_ft2.

    Returns {'side': 'outlet', 'c', 'table', 'position', 'v_fpm', 'pv',
    'loss', 'le_ft', 'eq', 'l_ft', 'ratio' (L/Le), 'blast_ratio', 'high',
    'fix_ft', 'notes'}. pv_of(v_fpm) gives the velocity pressure (standard
    air when None); loss = c x pv. fix_ft: the straight length from the
    outlet at which the loss first drops below HIGH_LOSS_INWC.
    v_fpm: the velocity used (design_velocity), None for cfm / area_ft2."""
    v = float(v_fpm) if v_fpm else float(cfm) / float(area_ft2)
    pv = (pv_of or std_pv_inwc)(v)
    le, eq = effective_duct_length_ft(area_ft2, v)
    ratio = float(l_ft) / le if le > 0 else 0.0
    c, pos, key, notes = outlet_c(fan_type, blast_ratio, ratio)
    high = is_high(c * pv)
    rec = {'side': 'outlet', 'c': c, 'table': key, 'position': pos,
           'v_fpm': v, 'pv': pv, 'loss': c * pv, 'le_ft': le, 'eq': eq,
           'l_ft': float(l_ft), 'ratio': ratio, 'blast_ratio': float(blast_ratio),
           'high': high,
           'fix_ft': outlet_fix_ft(fan_type, blast_ratio, le, c_limit(pv)) if high else None,
           'notes': notes}
    log.info('fan_system_effect outlet: %s pos %s Ab/Ao %.2f L %.2f ft Le %.2f ft '
             '(%s) C %.2f at %.0f fpm', key, pos, float(blast_ratio), float(l_ft),
             le, eq, c, v)
    return rec


def inlet_effect(is_round, size_in, r_in, area_ft2, cfm, l_ft, pv_of=None, v_fpm=None):
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
    here. fix_ft: where the loss first drops below HIGH_LOSS_INWC. pv_of, pv,
    loss and v_fpm as outlet_effect."""
    size_in = float(size_in)
    r_over = (float(r_in) / size_in) if r_in else None
    l_over = float(l_ft) * 12.0 / size_in
    c, key, r_used, notes = inlet_c(is_round, r_over, l_over)
    fr = floor_ratio(key, r_used)
    floor_c, _n = ashrae_fittings.lookup(key, r_used, fr)
    v = float(v_fpm) if v_fpm else float(cfm) / float(area_ft2)
    pv = (pv_of or std_pv_inwc)(v)
    high = is_high(c * pv)
    floor_ft = fr * size_in / 12.0
    rec = {'side': 'inlet', 'c': c, 'table': key, 'r_ratio': r_used,
           'size_in': size_in, 'v_fpm': v, 'pv': pv, 'loss': c * pv,
           'l_ft': float(l_ft), 'ratio': l_over, 'high': high,
           'fix_ft': inlet_fix_ft(is_round, r_used, size_in, c_limit(pv)) if high else None,
           'floor_ft': floor_ft, 'floor_ratio': fr, 'floor_c': floor_c,
           'within': float(l_ft) < floor_ft - 1e-9, 'notes': notes}
    log.info('fan_system_effect inlet: %s r/%s %.2f L %.2f ft (%.2f x %.1f in.) '
             'C %.2f at %.0f fpm', key, 'D' if is_round else 'H', r_used,
             float(l_ft), l_over, size_in, c, v)
    return rec


# ── inlet swirl ──────────────────────────────────────────────────────────────

def swirl_window_ft(d_h_in):
    """How far ahead of an inlet two elbows are checked for swirl:
    SWIRL_DH x the first duct's hydraulic diameter, ft."""
    return SWIRL_DH * float(d_h_in) / 12.0


def planes_differ(n1, n2):
    """True when two elbows' turn plane normals (unit (x, y, z), the cross
    product of each elbow's connector axes) are not parallel."""
    cx = n1[1] * n2[2] - n1[2] * n2[1]
    cy = n1[2] * n2[0] - n1[0] * n2[2]
    cz = n1[0] * n2[1] - n1[1] * n2[0]
    return math.sqrt(cx * cx + cy * cy + cz * cz) > SWIRL_SAME_PLANE


def swirl_pair(elbows):
    """The first two elbows that turn in different planes, or None.

    elbows: [(element id, L ft from the inlet, normal or None)] in order
    from the inlet. An elbow whose normal is unreadable is skipped."""
    known = [e for e in elbows if e[2] is not None]
    for i, a in enumerate(known):
        for b in known[i + 1:]:
            if planes_differ(a[2], b[2]):
                return a, b
    return None


def swirl_text(role, where, pair, window_ft, d_h_in):
    """The RED line for an elbow pair turning in different planes near a fan,
    in the Duct Performance reason style: each elbow named with its distance
    from the unit. INLET: inlet swirl (p. 34.13, Fig. 12). OUTLET: the book
    describes spin at the inlet only, so a supply outlet pair is reported as
    close-coupled elbows (p. 34.11), not as swirl (Colin, 2026-11-02: same
    6 Dh check on supply, return and exhaust)."""
    (a, la, _na), (b, lb, _nb) = pair
    head = ('RED: Inlet swirl at the %s' if role == INLET else
            'RED: Close-coupled elbows at the %s') % where
    tail = ('Swirl is not priced (no 2001 coefficient for FSE1,sw); provide straight '
            'duct or turning vanes (2001 Fundamentals p. 34.13, Fig. 12).'
            if role == INLET else
            'The loss is not priced past the first elbow (no 2001 data for '
            'close-coupled fittings, p. 34.11); provide straight duct or turning vanes.')
    return ('%s: elbows %s (%.1f ft) and %s (%.1f ft) turn in different planes within '
            '%.0f Dh (%.1f ft, Dh %.0f in.) of the fan. %s'
            % (head, a, float(la), b, float(lb), SWIRL_DH, float(window_ft),
               float(d_h_in), tail))


def double_elbow_text(where, elbows, window_ft, d_h_in):
    """The WARNING line for a double bend: two or more elbows within 6 Dh of
    a unit connection that do not turn in different planes (the RED swirl /
    close-coupled pair covers those). The first elbow is priced in the fan
    system effect and the second at its own C, but the added loss of two
    elbows this close is not, so the true loss is higher than shown (Colin,
    2026-11-04: best practice, untested in numbers). elbows: [(element id,
    L ft, normal or None)] in order from the unit; the first two are named."""
    (a, la, _na), (b, lb, _nb) = elbows[0], elbows[1]
    return ('WARNING: Double bend at the %s: elbows %s (%.1f ft) and %s (%.1f ft) '
            'follow each other within %.0f Dh (%.1f ft, Dh %.0f in.) of the fan. The '
            'fan system effect prices the first elbow only; the added loss of the '
            'second elbow this close is not priced (no 2001 data for close-coupled '
            'fittings, p. 34.11), so the true loss is higher than shown. Best '
            'practice: avoid a double bend at a fan connection.'
            % (where, a, float(la), b, float(lb), SWIRL_DH, float(window_ft),
               float(d_h_in)))


# Accessory kinds (fitting_tables.classify_accessory keys) that count as a
# damper for the damper in the fan clearance WARNING (Colin, 2026-11-04,
# backdraft dampers included).
DAMPER_KINDS = ('damper', 'fire_damper', 'backdraft')


def damper_window_text(where, acc_id, family, l_ft, win_ft, win_name):
    """The WARNING line for a damper inside the fan system effect window of
    a unit connection (1.0 Le at an outlet, 5 D or 5 H at an inlet; Colin,
    2026-11-04, F1). The walk passes an accessory at zero length, so the
    damper's own length is not part of L, and its added loss this close to
    the fan is not priced. A WARNING only: L and the price are unchanged."""
    return ('WARNING: Damper in the fan clearance at the %s: %s %s at %.1f ft is '
            'inside %s (%.1f ft). The fan system effect counts straight duct only, '
            'so the damper\'s length is not part of L, and the damper\'s added loss '
            'this close to the fan is not priced, so the true loss is higher than '
            'shown. Best practice: place dampers past %s from a fan connection.'
            % (where, family or '?', acc_id, float(l_ft), win_name, float(win_ft),
               win_name))


# ── straight duct RED (Colin, 2026-11-03) ────────────────────────────────────

def outlet_red_ft(le_ft):
    """OUTLET_RED_RATIO x Le, ft: an outlet elbow closer than this is RED."""
    return OUTLET_RED_RATIO * float(le_ft)


def inlet_red_ft(size_in):
    """INLET_RED_D x D, ft: an inlet elbow closer than this is RED."""
    return INLET_RED_D * float(size_in) / 12.0


def inlet_floor_ft(size_in):
    """INLET_FLOOR_D x D, ft: where the inlet tables stop falling."""
    return INLET_FLOOR_D * float(size_in) / 12.0


def le_text(where, eid, l_ft, le_ft, eq):
    """The RED line for an elbow under OUTLET_RED_RATIO x Le of a unit
    outlet, in the Duct Performance reason style. The lengths given are
    rounded UP to LENGTH_STEP_FT, so they are past the edge."""
    ratio = float(l_ft) / float(le_ft) if le_ft else 0.0
    return ('RED: Elbow %s, %.1f ft from the %s, L/Le %.2f (Le %.1f ft, %s). Provide at '
            'least %.1f ft of straight duct; %.1f ft (1.0 Le) removes the system '
            'effect.' % (eid, float(l_ft), where, ratio, float(le_ft), eq,
                         _ceil_step(outlet_red_ft(le_ft)), _ceil_step(le_ft)))


def inlet_red_text(where, eid, l_ft, size_in, is_round=True, source=None,
                   source_short=None):
    """The RED line for an elbow under INLET_RED_D x D of a unit inlet, in
    the Duct Performance reason style. H for a rectangular inlet. source
    names what D or H was read from, e.g. 'the 12x12 inlet collar' (short
    'the collar'), so it does not read against the fan system effect row,
    which reads the elbow's own size (Colin, 2026-11-09, TB-M12 3720096)."""
    dim = 'D' if is_round else 'H'
    of = (', ' + source) if source else ''
    of_short = (' of ' + source_short) if source_short else ''
    return ('RED: Elbow %s, %.1f ft from the %s, %.1f %s (%s %.0f in.%s). Provide at least '
            '%.1f ft of straight duct (%.0f %s%s); %.1f ft (%.0f %s) is where the loss '
            'stops falling.' % (
                eid, float(l_ft), where, float(l_ft) * 12.0 / float(size_in), dim, dim,
                float(size_in), of, _ceil_step(inlet_red_ft(size_in)), INLET_RED_D, dim,
                of_short, _ceil_step(inlet_floor_ft(size_in)), INLET_FLOOR_D, dim))


def straight_duct_reason(eid, l_ft, role):
    """The flagged table Reason for a straight duct RED with its elbow."""
    return 'Straight duct: elbow %s at %.1f ft, under %s' % (
        eid, float(l_ft), '%.1f Le' % OUTLET_RED_RATIO if role == OUTLET
        else '%.0f D' % INLET_RED_D)


def velocity_text(where, v_used, v_conn):
    """The NOTE when the default fan velocity governs at a connection."""
    return ('NOTE: fan system effect at the %s priced at the default fan velocity, '
            '%.0f fpm, greater than the connection velocity of %.0f fpm.' % (
                where, float(v_used), float(v_conn)))


def transition_text(eid, l_ft, where, v_used):
    """The NOTE for a transition priced at the fan velocity. Not printed
    since 2026-11-04 (the group charged once replaced the fan velocity
    transition; its NOTE is governs_text)."""
    return ('NOTE: transition %s, %.1f ft from the %s, priced at the default fan '
            'velocity, %.0f fpm at the unit end.' % (eid, float(l_ft), where, float(v_used)))


def no_effect_text(where, why):
    """The NOTE when no fan system effect is used at a connection (Colin,
    2026-11-09): no elbow inside the window, so the straight duct and
    fittings are priced at their own losses. Said in place of the fan
    velocity NOTE."""
    return ('NOTE: no fan system effect at the %s: %s; the straight duct and fittings '
            'are priced at their own losses.' % (where, why))


def governs_text(fse_wins, eid, l_ft, where, fse_inwc, own_inwc, win_ft, win_short,
                 transition=False, size_note=None):
    """The NOTE saying which was charged for the group at one connection
    (Colin, 2026-11-09): the two values side by side, then the outcome.
    Near the end of the window (L close to Le) the fan system effect falls
    toward 0 and the group's own losses govern. transition: a transition is
    in the group, so it is named. size_note: at an inlet, the size the
    window was read at, e.g. 'D 8 in. elbow' (Colin, 2026-11-09)."""
    group = 'straight duct, transition and elbow' if transition else 'straight duct and elbow'
    ratio = (' (L/%s %.2f%s)' % (win_short, float(l_ft) / float(win_ft),
                                 (', ' + size_note) if size_note else '')
             if win_ft else '')
    head = ('NOTE: %s, elbow %s at %.1f ft%s: fan system effect %.3f in. wc, %s %.3f '
            'in. wc. ' % (where, eid, float(l_ft), ratio, float(fse_inwc), group,
                          float(own_inwc)))
    if fse_wins:
        return head + 'Fan system effect charged, the %s are not charged again.' % group
    return head + ('%s charged at their own losses, no fan system effect.'
                   % (group[0].upper() + group[1:]))


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


def high_text(rec, where, eid):
    """The GRAY line for the engineer (Colin, 2026-10-29), in the Duct
    Performance reason style. where: 'supply outlet', 'return inlet'. eid:
    the elbow's element id, named with its distance from the unit (Colin,
    2026-11-02: every note names its element and where it sits)."""
    if rec.get('fix_ft') is None:
        fix = ('no straight length in the table brings it under %.1f in. wc'
               % HIGH_LOSS_INWC)
    else:
        fix = 'increase length of duct to %.1f ft' % rec['fix_ft']
    if rec['side'] == 'outlet':
        basis = '%s, worst of positions A to D, Le %.1f ft' % (rec['table'], rec['le_ft'])
    else:
        basis = '%s r/%s %.2f' % (rec['table'], 'D' if rec['table'] == INLET_ROUND
                                  else 'H', rec['r_ratio'])
    return ('GRAY: System Effect Pressure Loss too high at elbow %s, %.1f ft from '
            'the %s, %.3f in. wc, %s (C %.2f at %.0f fpm, %s; %.1f in. wc or more).' % (
                eid, rec['l_ft'], where, rec['loss'], fix, rec['c'], rec['v_fpm'],
                basis, HIGH_LOSS_INWC))


def disruption_text(kind, eid, where, l_ft, limit_ft, limit_name):
    """A tap, split or other fitting inside the window: flagged, not priced.
    eid: the tap, split node or fitting id. limit_name: 'Le' (outlet) or
    'the inlet window'."""
    return ('WARNING: %s %s, %.1f ft from the %s, inside %s (%.1f ft). A fan system '
            'effect disruption with no 2001 table, not priced.' % (
                kind, eid, float(l_ft), where, limit_name, float(limit_ft)))


def method_lines(settings):
    """(METHOD sentence, [ASSUMPTION lines]) for the external static report,
    console and drafting view alike. Off: a not-included sentence, no
    assumptions."""
    if not settings or not settings.get('on'):
        return ('Fan system effect not included (unticked); the straight duct and '
                'swirl checks still run.', [])
    fan = settings.get('fan_type') or DEFAULT_FAN_TYPE
    tables = OUTLET_TABLES.get(fan, OUTLET_TABLES[DEFAULT_FAN_TYPE])
    method = ('Fan system effect (2001 Fundamentals Eq. 38): the first disruption '
              'past each unit connection. Outlet: within Le (Eq. 36 or 37), an elbow '
              'is priced from the %s tables %s to %s, worst of positions A to D, at '
              'Ab/Ao %.2f. Inlet: within the table floor distance (5 D or 5 H), an '
              'elbow is priced from ED7-2 (round) or ER7-1 (rect). The straight '
              'duct, any transition on it and the first elbow are charged once, as '
              'the larger of the fan system effect and their own duct and fitting '
              'losses. Added before the safety factor. A loss of %.1f in. wc '
              'or more is GRAY, with the straight length that brings it under. '
              'L runs from the unit through straight duct and any transitions, '
              'their length counted, to the first elbow. Taps and splits inside '
              'the window are flagged, not '
              'priced. A rect mitered inlet elbow is read on the ER7-1 r/H %.1f '
              'row. An elbow under 0.5 Le of a unit outlet, or under 2 D of an inlet '
              '(D the larger of the unit opening and the first duct; on a rect '
              'elbow H, its side in the plane of the turn), is RED with the '
              'straight duct before it. Two elbows within 6 Dh of a unit '
              'connection: turning in different planes is RED (swirl or close-'
              'coupled), in the same plane a double bend WARNING; the added loss '
              'is not priced. A damper within the fan system effect window of a '
              'unit connection is a WARNING; its length is not counted in L and '
              'its added loss is not priced.' % (
                  fan, tables[0][1], tables[-1][1],
                  float(settings.get('blast_ratio', DEFAULT_BLAST_RATIO)), HIGH_LOSS_INWC,
                  MITERED_R_OVER_H))
    if settings.get('velocity_basis', DEFAULT_VELOCITY_BASIS) == 'default':
        method += (' Velocity at an outlet: the larger of the default fan velocity, '
                   '%.0f fpm, and the connection velocity (cfm over the unit connector '
                   'area). At an inlet: the connection velocity.'
                   % float(settings.get('fan_fpm', DEFAULT_FAN_FPM)))
    else:
        method += (' Velocity: the connection velocity, cfm over the unit connector '
                   'area (the first duct when unreadable).')
    assumptions = []
    if fan != 'SWSI':
        assumptions.append(
            'ASSUMPTION: the fan inlet tables ED7-2 and ER7-1 are printed for SWSI '
            'fans only; they are used for the %s fan selected.' % fan)
    return method, assumptions
