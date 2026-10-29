# -*- coding: utf-8 -*-
"""
ashrae_fittings.py  --  duct fitting loss coefficients looked up from the 2001
ASHRAE Fundamentals Handbook (SI), Ch. 34, fitting loss coefficient tables.

Colin, 2026-10-23: "that 2001 hvac fundamentals hand book is what i want to
use". Every C is looked up per fitting from that fitting's own geometry and
flows (angle, r/W, H/W, area and flow ratios), not one number per fitting type.
The tables are data in ashrae_fitting_tables.json, transcribed from the book
and spot checked against it (_test_ashrae_tables.py). Marked up copy of the
pages used: FOR CLAUDE\\Staging\\Design Standards\\
ASHRAE 2001 Fundamentals - RJA Fitting C Value Tables.pdf.

REFERENCE VELOCITY PRESSURE (the book, not a choice):
  - Fittings other than junctions: Co is referenced to section o (Eq. 30).
    Supply transitions run 1 -> o, so o is DOWNSTREAM; return/exhaust
    transitions run o -> 1 (toward the fan), so o is UPSTREAM in the air
    direction (SD4-1 and ED4-1 figures). Equal area elbows: either section.
  - Junctions (taps, tees, wyes): the PRINTED Cb is referenced to the branch
    section Pv and the printed Cs to the straight (downstream) section Pv.
    Example 8 (p. 34.24 to 34.25) sums each tee's Cb into its branch section
    and Cs into its straight section, and Eq. 35 prices every coefficient at
    its own section's Pv. Eq. 32 and 33 (p. 34.11) use Cc,b and Cc,s, the
    common section form, which is NOT what the tables print (Eq. 34 converts).
    The table ratios (Ab/Ac, Qb/Qc, As/Ac, Qs/Qc) are still against the common
    section c: the main on the unit side of the tap, supply or return.
    tap_branch_c / tap_main_c return Cc,b / Cc,s: the printed value converted
    to the common (main) section by Eq. 34 at the area ratio the table was
    READ at, so a table printed for one geometry (ER5-2, ER5-3 at Ab/Ac 0.5)
    keeps it and a small real branch cannot scale it up. The caller
    multiplies by the main's Pv (Colin, 2026-10-23; RJA mains are designed at
    600 fpm return and 800 fpm supply, low pressure).

Lookups interpolate linearly between printed rows and columns and clamp at the
table edges. Every lookup returns (C, notes); a note says when a value was
clamped or read next to a blank cell, so the report can show it.

RJA standards applied here (Colin, 2026-10-23):
  - Round elbows are die stamped smooth radius, r/D 1.5 (CD3-1). An elbow
    under r/D 1.5 (a 1 D type) takes CD3-1 x CD3-12(r/D) / CD3-12(1.5), CD3-12
    being the only 2001 round table that varies with r/D.
  - Smooth radius elbows at other angles: C x the CR3-1 angle factor K.
  - Rect mitered elbows: no vanes up to 45 deg (CR3-6 by angle and H/W),
    single thickness vanes above 45 deg (CR3-9, published for 90 deg only).

IronPython 2.7 / pyRevit -- no f-strings, no nonlocal.
"""

import io
import json
import logging
import os

log = logging.getLogger(__name__)

_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     'ashrae_fitting_tables.json')
_DATA = None

MM_PER_IN = 25.4
RJA_ROUND_ELBOW_R_D = 1.5      # company standard (Colin, 2026-10-23)
MITERED_NO_VANES_MAX_DEG = 45.0  # firm rule: vanes above this angle


def data():
    """The transcribed tables, loaded once."""
    global _DATA
    if _DATA is None:
        with io.open(_JSON, 'r', encoding='utf-8') as fh:
            _DATA = json.load(fh)
    return _DATA


def table(key):
    """One table dict, or None (logged) when the key is not transcribed."""
    t = data()['tables'].get(key)
    if t is None:
        log.info('ashrae_fittings: table %s NOT FOUND', key)
    return t


def table_label(key):
    """'SR5-13 Tee, 45 Degree Entry Branch, Diverging (p. 34.65)'."""
    t = table(key)
    if t is None:
        return key
    return '%s %s (p. %s)' % (t.get('fitting', key), t.get('title', ''), t.get('page', '?'))


# ── interpolation ────────────────────────────────────────────────────────────

def _bracket(xs, x):
    """(i0, i1, frac, clamped) for x on the ascending axis xs."""
    if x <= xs[0]:
        return 0, 0, 0.0, x < xs[0]
    if x >= xs[-1]:
        n = len(xs) - 1
        return n, n, 0.0, x > xs[-1]
    for i in range(1, len(xs)):
        if x <= xs[i]:
            a, b = xs[i - 1], xs[i]
            return i - 1, i, (x - a) / float(b - a), False
    n = len(xs) - 1
    return n, n, 0.0, True


def _nearest_filled(row, j):
    """Value in row at column j, or the nearest non-blank one beside it."""
    if row[j] is not None:
        return row[j], False
    for d in range(1, len(row)):
        for k in (j - d, j + d):
            if 0 <= k < len(row) and row[k] is not None:
                return row[k], True
    return None, True


def _lerp(a, b, f):
    return a + (b - a) * f


def _grid2(grid, rows, cols, r, c):
    """Bilinear value at (r, c) on a printed grid. Returns (value, clamped, blank)."""
    i0, i1, fr, cl_r = _bracket(rows, r)
    j0, j1, fc, cl_c = _bracket(cols, c)
    blank = False
    corner = {}
    for i in (i0, i1):
        for j in (j0, j1):
            v, b = _nearest_filled(grid[i], j)
            if v is None:
                # whole printed row blank: borrow the nearest filled row
                for d in range(1, len(grid)):
                    for k in (i - d, i + d):
                        if 0 <= k < len(grid):
                            v, b2 = _nearest_filled(grid[k], j)
                            if v is not None:
                                break
                    if v is not None:
                        break
                b = True
            corner[(i, j)] = v
            blank = blank or b
    top = _lerp(corner[(i0, j0)], corner[(i0, j1)], fc)
    bot = _lerp(corner[(i1, j0)], corner[(i1, j1)], fc)
    return _lerp(top, bot, fr), (cl_r or cl_c), blank


def lookup(key, *args):
    """C from table `key` at the given independent variables, in the order the
    table lists them (x; rows, cols; layers, rows, cols).

    Returns (C or None, notes). None means the table is missing."""
    t = table(key)
    if t is None:
        return None, ['%s not transcribed' % key]
    kind = t['kind']
    notes = []
    if kind == 'const':
        return float(t['C']), notes
    if kind == '1d':
        xs = t['x']['values']
        i0, i1, f, cl = _bracket(xs, float(args[0]))
        a, b = t['C'][i0], t['C'][i1]
        if a is None or b is None:
            a = b = (a if a is not None else b)
            notes.append('%s: next to a blank cell' % key)
        if cl:
            notes.append('%s: %s %.2f outside the table, edge value used'
                         % (key, t['x']['name'], float(args[0])))
        return _lerp(a, b, f), notes
    if kind == '2d':
        v, cl, blank = _grid2(t['C'], t['rows']['values'], t['cols']['values'],
                              float(args[0]), float(args[1]))
    elif kind == '3d':
        lv = t['layers']['values']
        l0, l1, fl, cll = _bracket(lv, float(args[0]))
        v0, cl0, b0 = _grid2(t['C'][l0], t['rows']['values'], t['cols']['values'],
                             float(args[1]), float(args[2]))
        v1, cl1, b1 = _grid2(t['C'][l1], t['rows']['values'], t['cols']['values'],
                             float(args[1]), float(args[2]))
        v, cl, blank = _lerp(v0, v1, fl), (cll or cl0 or cl1), (b0 or b1)
    else:
        return None, ['%s: unknown table kind %r' % (key, kind)]
    if cl:
        notes.append('%s: outside the table at %s, edge value used'
                     % (key, ', '.join('%.2f' % float(a) for a in args)))
    if blank:
        notes.append('%s: blank cell in the book, nearest printed value used' % key)
    return v, notes


# ── elbows ───────────────────────────────────────────────────────────────────

def angle_factor(angle_deg):
    """CR3-1 angle factor K. None (angle unreadable) is a 90: K 1.0."""
    if angle_deg is None:
        return 1.0, []
    return lookup('CR3-1_K', float(angle_deg))


def round_elbow_c(d_in, r_over_d, angle_deg):
    """Round smooth radius elbow: CD3-1 (die stamped, r/D 1.5) by diameter,
    corrected for r/D by CD3-12 and for angle by the CR3-1 K.

    Returns (C, label, notes)."""
    c90, notes = lookup('CD3-1', float(d_in) * MM_PER_IN)
    # The last CD3-1 column is printed "250 mm and up": not a clamp.
    notes = [n for n in notes if not n.startswith('CD3-1:')
             or float(d_in) * MM_PER_IN < 75.0]
    label = 'CD3-1 r/D 1.5, D %.0f in.' % float(d_in)
    if r_over_d and abs(float(r_over_d) - RJA_ROUND_ELBOW_R_D) > 0.05:
        at_rd, n1 = lookup('CD3-12', float(r_over_d))
        at_std, n2 = lookup('CD3-12', RJA_ROUND_ELBOW_R_D)
        notes += n1 + n2
        if at_std:
            c90 = c90 * at_rd / at_std
            label += ' x CD3-12 r/D %.2f (x%.2f)' % (float(r_over_d), at_rd / at_std)
    k, n3 = angle_factor(angle_deg)
    notes += n3
    if angle_deg is not None and abs(float(angle_deg) - 90.0) > 0.5:
        label += ', %.0f deg K %.2f' % (float(angle_deg), k)
    return c90 * k, label, notes


def rect_smooth_elbow_c(r_over_w, h_over_w, angle_deg):
    """Rect smooth radius elbow without vanes: CR3-1 Cp by r/W and H/W, x K."""
    cp, notes = lookup('CR3-1', float(r_over_w), float(h_over_w))
    k, n2 = angle_factor(angle_deg)
    label = 'CR3-1 r/W %.2f, H/W %.2f' % (float(r_over_w), float(h_over_w))
    if angle_deg is not None and abs(float(angle_deg) - 90.0) > 0.5:
        label += ', %.0f deg K %.2f' % (float(angle_deg), k)
    return cp * k, label, notes + n2


def rect_mitered_elbow_c(h_over_w, angle_deg):
    """Rect mitered elbow. Up to 45 deg, no vanes: CR3-6 by angle and H/W.
    Above 45 deg (and an unreadable angle), single thickness vanes: CR3-9."""
    if angle_deg is not None and float(angle_deg) <= MITERED_NO_VANES_MAX_DEG:
        c, notes = lookup('CR3-6', float(angle_deg), float(h_over_w))
        return c, 'CR3-6 no vanes, %.0f deg, H/W %.2f' % (float(angle_deg),
                                                         float(h_over_w)), notes
    c, notes = lookup('CR3-9')
    return c, 'CR3-9 single thickness vanes', notes


# ── transitions ──────────────────────────────────────────────────────────────

# (supply?, shape at section 1, shape at section o) -> table. Section 1 is
# upstream for supply and downstream (toward the fan) for return.
_TRANSITION_TABLES = {
    (True,  'rect',  'rect'):  'SR4-1',
    (True,  'round', 'round'): 'SD4-1',
    (True,  'rect',  'round'): 'SD4-2',
    (True,  'round', 'rect'):  'SR4-3',
    (False, 'rect',  'rect'):  'ER4-1',
    (False, 'round', 'round'): 'ED4-1',
    # Return figures: ER4-3 (rect to round) has the rect end at o and the
    # round end at 1 toward the fan; ED4-2 (round to rect) the reverse.
    (False, 'round', 'rect'):  'ER4-3',
    (False, 'rect',  'round'): 'ED4-2',
}


def transition_c(is_supply, shape_1, shape_o, area_1, area_o, theta_deg):
    """Co for a transition, referenced to section o.

    is_supply: air runs 1 -> o (supply) or o -> 1 (return/exhaust).
    shape_*: 'rect' or 'round'. area_*: ft2. theta_deg: included angle, the
    larger of the two side pairs on a rectangular transition.
    Returns (Co, label, notes), Co None when no table fits."""
    key = _TRANSITION_TABLES.get((bool(is_supply), shape_1, shape_o))
    if key is None or not area_1 or not area_o or theta_deg is None:
        return None, 'transition geometry unreadable', []
    ratio = float(area_o) / float(area_1)
    if abs(ratio - 1.0) < 1e-6:
        return 0.0, '%s Ao/A1 1.00 (no size change)' % key, []
    c, notes = lookup(key, ratio, float(theta_deg))
    return c, '%s Ao/A1 %.2f, %.0f deg' % (key, ratio, float(theta_deg)), notes


# ── junctions ────────────────────────────────────────────────────────────────

def _junction_keys(is_supply, tap_shape, main_shape, is_45_entry, main_d_in=None):
    """(Cb table, Cs table) for a tap of this shape on a main of this shape."""
    if is_supply:
        if main_shape == 'rect':
            return (('SR5-13_Cb', 'SR5-13_Cs') if tap_shape == 'rect'
                    else ('SR5-11_Cb', 'SR5-11_Cs'))
        if tap_shape == 'round':
            return (('SD5-1_Cb', 'SD5-1_Cs') if is_45_entry
                    else ('SD5-9_Cb', 'SD5-9_Cs'))
        return None, None
    if main_shape == 'rect':
        return (('ER5-3_Cb', 'ER5-3_Cs') if tap_shape == 'rect'
                else ('ER5-2_Cb', 'ER5-2_Cs'))
    if tap_shape == 'round':
        return (('ED5-2_Cb', 'ED5-2_Cs') if is_45_entry
                else (_ed53_key('Cb', main_d_in), _ed53_key('Cs', main_d_in)))
    return None, None


def _ed53_key(coef='Cb', main_d_in=None):
    """ED5-3 is printed in two parts by main diameter, Dc 250 mm or less and
    over 250 mm. An unknown diameter takes the larger main's table."""
    small = main_d_in is not None and float(main_d_in) * MM_PER_IN <= 250.0
    return 'ED5-3_%s_%s' % ('le250' if small else 'gt250', coef)


def _call(key, t, ratios):
    """Look a junction table up with the ratios it is printed against.

    ratios: {'Ab/Ac':, 'Qb/Qc':, 'As/Ac':, 'Qs/Qc':}. The table's own axis
    names pick which ratios go in, so a 1d Cb-by-Qb/Qc table and a 2d
    Ab/Ac x Qb/Qc table are called the same way."""
    names = []
    if t['kind'] == '1d':
        names = [t['x']['name']]
    elif t['kind'] == '2d':
        names = [t['rows']['name'], t['cols']['name']]
    elif t['kind'] == '3d':
        names = [t['layers']['name'], t['rows']['name'], t['cols']['name']]
    args = []
    for n in names:
        if n not in ratios:
            return None, ['%s: needs %s' % (key, n)]
        args.append(ratios[n])
    return lookup(key, *args)


def _ratio_read_at(t, name, actual):
    """The area ratio a junction table was actually read at.

    A table printed for one geometry (ER5-2 and ER5-3: Ab/Ac 0.5, As = Ac)
    is read at that geometry whatever the model's areas are (Colin,
    2026-10-23: "keep them at the geometry they were printed for"). A table
    printed across the ratio is read at the model's value, clamped to the
    printed range."""
    cond = t.get('conditions') or {}
    if name in cond:
        return float(cond[name])
    if name == 'As/Ac' and cond.get('As = Ac'):
        return 1.0
    for axis in ('layers', 'rows', 'cols', 'x'):
        if axis in t and t[axis]['name'] == name:
            vals = t[axis]['values']
            return min(max(float(actual), vals[0]), vals[-1])
    return float(actual)


def _to_common(c, q_ratio, area_ratio):
    """Eq. 34 turned around: the printed C, referenced to its own section,
    re-referenced to the common (main) section velocity pressure.
    Cc,i = Ci x (Vi/Vc)^2, with Vi/Vc = (Qi/Qc) / (Ai/Ac)."""
    v_ratio = float(q_ratio) / float(area_ratio) if area_ratio else 0.0
    return c * v_ratio * v_ratio


def tap_branch_c(is_supply, tap_shape, main_shape, is_45_entry,
                 ab_ac, qb_qc, as_ac=1.0, main_d_in=None):
    """Cc,b for the air turning off through a tap: referenced to the COMMON
    (main) section velocity pressure, so the caller multiplies by the main's
    Pv. The printed Cb (branch section reference) is converted at the Ab/Ac
    the table was read at, so a table printed for one geometry keeps it.
    Returns (Cc,b, label, notes); None when no table fits (the caller falls
    back to the RJA worksheet)."""
    kb, _ks = _junction_keys(is_supply, tap_shape, main_shape, is_45_entry, main_d_in)
    t = table(kb) if kb else None
    if t is None:
        return None, 'no 2001 table for this tap', []
    ratios = {'Ab/Ac': ab_ac, 'Qb/Qc': qb_qc, 'As/Ac': as_ac,
              'Qs/Qc': 1.0 - qb_qc}
    cb, notes = _call(kb, t, ratios)
    if cb is None:
        return None, 'no 2001 table for this tap', notes
    read_ab = _ratio_read_at(t, 'Ab/Ac', ab_ac)
    q_read = min(max(float(qb_qc), 0.1), 1.0)
    ccb = _to_common(cb, q_read, read_ab)
    label = '%s Ab/Ac %.2f, Qb/Qc %.2f: Cb %.2f, Cc,b %.2f' % (
        kb.split('_')[0], read_ab, q_read, cb, ccb)
    return ccb, label, notes


def tap_main_c(is_supply, tap_shape, main_shape, is_45_entry,
               as_ac, qs_qc, ab_ac=None, main_d_in=None):
    """Cc,s for the air continuing past a tap on the main: referenced to the
    COMMON (main) section velocity pressure, converted from the printed Cs
    (straight section reference) at the As/Ac the table was read at.
    Returns (Cc,s, label, notes)."""
    _kb, ks = _junction_keys(is_supply, tap_shape, main_shape, is_45_entry, main_d_in)
    t = table(ks) if ks else None
    if t is None:
        return None, 'no 2001 table for this tap', []
    ratios = {'As/Ac': as_ac, 'Qs/Qc': qs_qc, 'Qb/Qc': 1.0 - qs_qc}
    if ab_ac is not None:
        # A table printed against Ab/Ac (ED5-2, ED5-3 Cs) needs the real
        # branch area; without it _call answers None and the caller falls back.
        ratios['Ab/Ac'] = ab_ac
    cs, notes = _call(ks, t, ratios)
    if cs is None:
        return None, 'no 2001 table for this tap', notes
    read_as = _ratio_read_at(t, 'As/Ac', as_ac)
    q_read = min(max(float(qs_qc), 0.1), 1.0)
    ccs = _to_common(cs, q_read, read_as)
    label = '%s As/Ac %.2f, Qs/Qc %.2f: Cs %.2f, Cc,s %.2f' % (
        ks.split('_')[0], read_as, q_read, cs, ccs)
    return ccs, label, notes


def bullhead_wye_c(is_supply, ab_ac):
    """Symmetrical dovetail wye at Qb/Qc 0.5: SR5-14 / ER5-4 Cb by Ab/Ac."""
    key = 'SR5-14' if is_supply else 'ER5-4'
    c, notes = lookup(key, float(ab_ac))
    return c, '%s Ab/Ac %.2f' % (key, float(ab_ac)), notes


# ── what the Advanced settings dialog lists ──────────────────────────────────
#
# One row per fitting the tool prices: (label, model family it covers, the
# table keys it reads, the RJA worksheet key used as the fallback when the
# fitting's geometry cannot be read).
FITTING_ROWS = (
    ('Round elbow', 'Round Elbow (1.5D standard, 1D corrected)',
     ('CD3-1', 'CD3-12', 'CR3-1_K'), 'round_elbow_90'),
    ('Rect smooth radius elbow', 'Rectangular Elbow - Radius',
     ('CR3-1', 'CR3-1_K'), 'rect_elbow_90'),
    ('Rect mitered elbow, up to 45 deg', 'Rectangular Elbow - Mitered',
     ('CR3-6',), 'rect_elbow_90'),
    ('Rect mitered elbow, over 45 deg (vanes)', 'Rectangular Elbow - Mitered - Single Thickness Vanes',
     ('CR3-9',), 'rect_elbow_90'),
    ('Rect transition, supply', 'Rectangular Transition - Length',
     ('SR4-1',), 'transition_contraction'),
    ('Rect transition, return / exhaust', 'Rectangular Transition - Length',
     ('ER4-1',), 'transition_expansion'),
    ('Round transition, supply', 'Round Transition - Angle',
     ('SD4-1',), 'transition_contraction'),
    ('Round transition, return / exhaust', 'Round Transition - Angle',
     ('ED4-1',), 'transition_expansion'),
    ('Rect / round transition, supply', 'Rectangular to Round Transition',
     ('SD4-2', 'SR4-3'), 'transition_contraction'),
    ('Rect / round transition, return / exhaust', 'Rectangular to Round Transition',
     ('ER4-3', 'ED4-2'), 'transition_expansion'),
    ('Rect tap on rect main, supply', 'Rectangular Tap - Beveled',
     ('SR5-13_Cb', 'SR5-13_Cs'), 'supply_tap_branch'),
    ('Round tap on rect main, supply', 'Round Takeoff - Shoe / w_Damper',
     ('SR5-11_Cb', 'SR5-11_Cs'), 'supply_tap_branch'),
    ('Round 45 deg tap on round main, supply', 'Round Takeoff - Shoe',
     ('SD5-1_Cb', 'SD5-1_Cs'), 'supply_tap_branch'),
    ('Round straight tap on round main, supply', 'Round Takeoff w_Damper, tees',
     ('SD5-9_Cb', 'SD5-9_Cs'), 'supply_tap_branch'),
    ('Rect tap on rect main, return / exhaust (table fixed at Ab/Ac 0.5)',
     'Rectangular Tap - Beveled', ('ER5-3_Cb', 'ER5-3_Cs'), 'exhaust_tap_branch'),
    ('Round tap on rect main, return / exhaust (table fixed at Ab/Ac 0.5)',
     'Round Takeoff', ('ER5-2_Cb', 'ER5-2_Cs'), 'exhaust_tap_branch'),
    ('Round 45 deg tap on round main, return / exhaust', 'Round Takeoff - Shoe',
     ('ED5-2_Cb', 'ED5-2_Cs'), 'exhaust_tap_branch'),
    ('Round straight tap on round main, return / exhaust', 'Round Takeoff w_Damper, tees',
     ('ED5-3_le250_Cb', 'ED5-3_le250_Cs', 'ED5-3_gt250_Cb', 'ED5-3_gt250_Cs'),
     'exhaust_tap_branch'),
    # Fan system effect (Colin, 2026-10-29), read by fan_system_effect.py.
    # No worksheet fallback: an unreadable connection is reported, not priced.
    ('Fan outlet with elbow, DWDI (fan system effect)', 'Unit supply connection',
     ('SR7-9', 'SR7-10', 'SR7-11', 'SR7-12'), None),
    ('Fan outlet with elbow, SWSI (fan system effect)', 'Unit supply connection',
     ('SR7-5', 'SR7-6', 'SR7-7', 'SR7-8'), None),
    ('Fan inlet with round elbow (fan system effect)', 'Unit return / exhaust connection',
     ('ED7-2',), None),
    ('Fan inlet with rect elbow (fan system effect)', 'Unit return / exhaust connection',
     ('ER7-1',), None),
)
# Not yet looked up from the 2001 tables, still the RJA worksheet value:
# tees, crosses, wyes (incl. bullhead, SR5-14 / ER5-4 are transcribed) and
# offsets (CR3-17 / CD6-4 transcribed).


def table_keys_like(prefix):
    """Every transcribed key starting with prefix (ED5-3 has several parts)."""
    return [k for k in sorted(data()['tables'].keys()) if k.startswith(prefix)]
