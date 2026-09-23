# -*- coding: ascii -*-
"""
diffuser_tables.py

Diffuser / grille sizing tables from RJA's MEP Design Standards, section
"1. Diffuser and Branch Ducting Standards" (discussion draft effective
2026-09-17). Pure data + lookup module, same shape as gas_tables.py:
plain module-level data, free functions, no classes, no Revit API, no I/O.

Why this module exists at all
-----------------------------
A correctly-selected auto-sizing diffuser is picked on static pressure and
NC, not on duct velocity. Velocity-checking a duct run that only ever feeds
one such diffuser is therefore redundant and produces false flags. The real
risk left on that run is a plain size mismatch: is the duct actually the
size the diffuser needs?

These capacities are transcribed here independently of the diffuser
families' own live sizing formulas on purpose. An independent lookup still
catches a branch whose family formula has drifted from its own catalog data
- that has actually happened on this project (CONFIRMED BUG 2026-09-16, the
live Return Grate family disagreeing with its own source table). Reading the
family's formula instead would just agree with the bug.

Units and snapping contract
---------------------------
All dimensions are inches. Every lookup here expects dimensions that the
CALLER has already snapped to the nominal 2" grid via
hvac_graph._nominal_even_in(). This module only does tolerant equality
(_TOL_IN) against the published values, so the nominal-snapping rule lives
in exactly one place (hvac_graph) instead of being re-implemented here -
and this module stays free of any Revit dependency, which is also what
keeps hvac_graph -> diffuser_tables a one-way import with no cycle.

What "not found" means
----------------------
Every lookup returns None rather than a default when a size or a CFM falls
outside the published rows. None means "the design standard does not cover
this, so the selection cannot be verified" - which is a third answer,
distinct from pass and from fail. Callers must surface it as such (gray /
"cannot check"), never collapse it into a pass.

IronPython 2.7 compatible. Pure ASCII source.
"""


# ---------------------------------------------------------------------------
# Ceiling diffusers / grates - round neck, (diameter_in, max_cfm)
# ---------------------------------------------------------------------------
#
# Both ceiling tables are MERGED across face modules (12x12 / 24x12 / 24x24).
# Verified against the source doc: every module publishes identical CFM
# limits for every neck size it shares with the others - the modules differ
# only in how far up the neck range they go. A single diameter -> max-CFM
# curve is therefore valid regardless of face size, and no module/face-size
# detection is needed anywhere in the branch check.
#
# Ascending by diameter. Rows are the MAXIMUM CFM for that neck; the source
# doc also publishes a Min CFM per row (the bottom of each selection band),
# deliberately not carried here - the branch check is a max-capacity check
# only, per the approved plan.

# Basis of design: flex duct manufacturer maximum recommended airflow per
# neck diameter. Even neck sizes only, capped at 14".
# NOTE: the 4" row is flagged "residential only" in the source doc. It is
# kept here so a residential 4" neck can still be checked rather than
# reported as an unpublished size; it is not a commercial selection.
CEILING_SUPPLY_NECK_TABLE = [
    (4,   35),
    (6,  110),
    (8,  200),
    (10, 375),
    (12, 580),
    (14, 750),
]

# Basis of design: Titus PAR perforated face grate. Return and Exhaust share
# identical published data, hence one table for both.
CEILING_RETURN_EXHAUST_NECK_TABLE = [
    (6,   85),
    (8,  185),
    (10, 325),
    (12, 470),
    (14, 640),
    (16, 840),
]


# ---------------------------------------------------------------------------
# Sidewall grilles - rectangular face, (width_in, height_in, max_cfm, nc)
# ---------------------------------------------------------------------------
#
# Flattened from the source doc's three separate aspect-ratio tables
# (Square / 2:1 / 3:1) into one list of real published W x H points. They are
# flattened because the branch check matches an installed face size against
# whatever row actually publishes it - which aspect-ratio table that row came
# from is not something the Revit model can tell us, and does not change the
# capacity.
#
# Rows are NOT sorted here; they are kept grouped Square, then 2:1, then 3:1
# in source-doc order so they can be diffed against the doc by eye. Every
# lookup that needs an order imposes its own.
#
# NC is carried for reference / future reporting only. Nothing in this module
# or in the branch check reads it - the check is capacity-only.

SIDEWALL_SUPPLY_TABLE = [
    # --- Square type ---
    (6,   6,  125, 20),
    (8,   8,  245, 23),
    (10, 10,  390, 25),
    (12, 12,  585, 27),
    (14, 14,  810, 28),
    (16, 16, 1075, 29),
    (18, 18, 1370, 30),
    (20, 20, 1705, 31),
    (22, 22, 2080, 32),
    (24, 24, 2485, 33),
    # --- 2:1 type ---
    (12,  6,  270, 24),
    (18,  8,  585, 27),
    (18, 10,  735, 28),
    (22, 10,  910, 28),
    (24, 12, 1205, 30),
    (30, 16, 2060, 32),
    (36, 16, 2485, 33),
    # --- 3:1 type ---
    (16,  6,  380, 25),
    (18,  6,  415, 26),
    (20,  6,  475, 26),
    (36, 12, 1820, 32),
    (48, 10, 2060, 32),
    (42, 12, 2135, 32),
    (48, 12, 2485, 33),
]

SIDEWALL_RETURN_EXHAUST_TABLE = [
    # --- Square type ---
    (6,   6,   95, 10),
    (8,   8,  185, 10),
    (10, 10,  295, 10),
    (12, 12,  440, 11),
    (14, 14,  610, 12),
    (16, 16,  810, 14),
    (18, 18, 1035, 14),
    (20, 20, 1285, 16),
    (22, 22, 1570, 17),
    (24, 24, 1875, 18),
    # --- 2:1 type ---
    (12,  6,  205, 10),
    (18, 10,  555, 12),
    (22, 10,  685, 13),
    (24, 12,  910, 14),
    (30, 16, 1555, 17),
    # --- 3:1 type ---
    (16,  6,  285, 10),
    (18,  6,  315, 10),
    (20,  6,  360, 11),
    (36, 12, 1375, 16),
    (42, 12, 1610, 17),
    (48, 12, 1875, 18),
]


# Dimensional match tolerance, inches. Callers snap to the nominal 2" grid
# before calling in, so anything reaching a lookup should already be an exact
# even number; this only absorbs float representation noise from the Revit
# feet <-> inch round trip. Kept well under 1" so it can never bridge two
# adjacent nominal sizes.
#
# Public because hvac_graph's duct-vs-neck comparisons need the same
# tolerance, and two independently-tuned copies of it would be free to drift
# apart.
SIZE_TOL_IN = 0.25


def _same_in(a, b):
    """True if two inch dimensions are the same nominal size within SIZE_TOL_IN."""
    return abs(float(a) - float(b)) <= SIZE_TOL_IN


# ---------------------------------------------------------------------------
# Display helpers
# ---------------------------------------------------------------------------
#
# Formatting lives here so the branch check and the output tables can never
# render the same size two different ways. Formats deliberately match
# Duct Velocity script.py's _duct_size_label(): '10"' and '18"x12"'.

def round_size_label(diameter_in):
    """Display string for a round neck / round duct: '10"'."""
    if diameter_in is None:
        return '-'
    return '{:.0f}"'.format(diameter_in)


def rect_size_label(w_in, h_in):
    """Display string for a rectangular face / rectangular duct: '18"x12"'."""
    if w_in is None or h_in is None:
        return '-'
    return '{:.0f}"x{:.0f}"'.format(w_in, h_in)


def row_size_label(row):
    """Display string for a sidewall table row tuple (w, h, max_cfm, nc)."""
    if row is None:
        return '-'
    return rect_size_label(row[0], row[1])


# ---------------------------------------------------------------------------
# Round-neck lookups (ceiling supply, ceiling return/exhaust)
# ---------------------------------------------------------------------------

def max_cfm_for_diameter(table, diameter_in):
    """Max CFM the design standard publishes for a given round neck size.

    This is the "is the diffuser's own neck adequate for its own CFM" half of
    the branch check: compare the terminal's actual Flow against this.

    table:       CEILING_SUPPLY_NECK_TABLE or CEILING_RETURN_EXHAUST_NECK_TABLE
    diameter_in: neck diameter in inches, already nominal-snapped by the caller

    Returns the row's max CFM as a number, or None when that diameter is not a
    published row (including any diameter above the largest published neck).
    None means unverifiable, not unlimited - see the module docstring.
    """
    if diameter_in is None:
        return None
    for dia, max_cfm in table:
        if _same_in(dia, diameter_in):
            return max_cfm
    return None


def min_diameter_for_cfm(table, cfm):
    """Smallest published round neck that carries the given CFM.

    Used for the "Required Size" display value, which is wanted even when the
    installed neck already passes - the point of that column is to show what
    the standard actually calls for, not only to explain failures.

    Returns the diameter in inches, or None when no published neck carries
    that CFM (i.e. the load is past the top of the table and needs a
    selection the standard does not cover).
    """
    if cfm is None:
        return None
    best = None
    for dia, max_cfm in sorted(table):
        if cfm <= max_cfm:
            best = dia
            break
    return best


def largest_diameter(table):
    """Largest published neck diameter in a round table - used only to word
    an over-the-top-of-the-table message ('>14"') rather than print nothing."""
    return max(dia for dia, _max_cfm in table)


# ---------------------------------------------------------------------------
# Sidewall (rectangular face) lookups
# ---------------------------------------------------------------------------

def _face_area(row):
    """Face area in square inches for a sidewall row, used as the ordering
    key wherever 'smallest that works' or 'bigger than' is needed. Area is the
    right ordering here because the published rows span three aspect ratios
    and are not comparable dimension-by-dimension (a 48x10 is not 'smaller'
    than a 24x24 on either single dimension, but it is a larger face)."""
    return float(row[0]) * float(row[1])


def sidewall_row(table, w_in, h_in):
    """The published catalog row for an installed face size, or None.

    Matching is orientation-insensitive: a face published as 18x10 is matched
    by an installed 10x18 as well. Capacity is a function of the face's free
    area, not of which dimension the family happens to call Width, and a
    sidewall grille legitimately gets placed either way up. Being strict about
    orientation here would produce false "size not in table" results on real,
    correct selections.

    w_in / h_in: inches, already nominal-snapped by the caller.

    Returns the (w, h, max_cfm, nc) tuple as published, or None if no row
    matches - meaning the installed face is not a standard selection at all,
    which the caller must report as unverifiable rather than as a pass.
    """
    if w_in is None or h_in is None:
        return None
    for row in table:
        if _same_in(row[0], w_in) and _same_in(row[1], h_in):
            return row
        if _same_in(row[0], h_in) and _same_in(row[1], w_in):
            return row
    return None


def nearest_sidewall_row(table, w_in, h_in):
    """Closest published row to an installed face size, as (row, is_exact).

    Only for wording a message when sidewall_row() found nothing - it lets the
    output say which real size the modeled face is nearly, instead of just
    "not in table". The returned row is NOT a substitute verdict: when
    is_exact is False the caller must still treat the selection as
    unverifiable, because a face that is not a published size has no published
    capacity.

    Distance is summed absolute dimension error, tie-broken by the smaller
    face area so the answer is deterministic. Returns (None, False) for an
    empty table or missing dimensions.
    """
    if w_in is None or h_in is None or not table:
        return None, False
    exact = sidewall_row(table, w_in, h_in)
    if exact is not None:
        return exact, True
    best = None
    best_key = None
    for row in table:
        direct   = abs(row[0] - w_in) + abs(row[1] - h_in)
        flipped  = abs(row[0] - h_in) + abs(row[1] - w_in)
        dist     = direct if direct <= flipped else flipped
        key      = (dist, _face_area(row))
        if best_key is None or key < best_key:
            best     = row
            best_key = key
    return best, False


def min_row_for_cfm(table, cfm):
    """Smallest published face (by area) that carries the given CFM.

    Counterpart of min_diameter_for_cfm() for sidewall grilles - same purpose,
    the "Required Size" display value.

    Returns the (w, h, max_cfm, nc) tuple, or None when no published face
    carries that CFM.
    """
    if cfm is None:
        return None
    best = None
    for row in table:
        if cfm > row[2]:
            continue
        if best is None or _face_area(row) < _face_area(best):
            best = row
    return best


def largest_sidewall_row(table):
    """Largest published face by area - used only to word an
    over-the-top-of-the-table message rather than print nothing."""
    best = None
    for row in table:
        if best is None or _face_area(row) > _face_area(best):
            best = row
    return best


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

def _self_test():
    """Non-Revit self-test. Runs under plain CPython as well as IronPython -
    this module imports nothing from Revit specifically so that it can be
    sanity-checked outside a Revit session."""
    print('diffuser_tables.py self-test')

    # Row counts as transcribed from the source doc.
    assert len(CEILING_SUPPLY_NECK_TABLE) == 6, 'ceiling supply row count'
    assert len(CEILING_RETURN_EXHAUST_NECK_TABLE) == 6, 'ceiling ret/exh row count'
    assert len(SIDEWALL_SUPPLY_TABLE) == 24, 'sidewall supply row count'
    assert len(SIDEWALL_RETURN_EXHAUST_TABLE) == 21, 'sidewall ret/exh row count'

    # Round lookups
    assert max_cfm_for_diameter(CEILING_SUPPLY_NECK_TABLE, 10) == 375
    assert max_cfm_for_diameter(CEILING_SUPPLY_NECK_TABLE, 10.000000000000002) == 375
    assert max_cfm_for_diameter(CEILING_SUPPLY_NECK_TABLE, 18) is None
    assert max_cfm_for_diameter(CEILING_RETURN_EXHAUST_NECK_TABLE, 4) is None
    assert max_cfm_for_diameter(CEILING_RETURN_EXHAUST_NECK_TABLE, 16) == 840

    assert min_diameter_for_cfm(CEILING_SUPPLY_NECK_TABLE, 200) == 8
    assert min_diameter_for_cfm(CEILING_SUPPLY_NECK_TABLE, 201) == 10
    assert min_diameter_for_cfm(CEILING_SUPPLY_NECK_TABLE, 751) is None
    assert largest_diameter(CEILING_SUPPLY_NECK_TABLE) == 14

    # Sidewall lookups
    assert sidewall_row(SIDEWALL_SUPPLY_TABLE, 18, 10) == (18, 10, 735, 28)
    assert sidewall_row(SIDEWALL_SUPPLY_TABLE, 10, 18) == (18, 10, 735, 28), \
        'orientation-insensitive match'
    assert sidewall_row(SIDEWALL_SUPPLY_TABLE, 14, 6) is None, 'unpublished face'

    row, exact = nearest_sidewall_row(SIDEWALL_SUPPLY_TABLE, 12, 12)
    assert exact and row == (12, 12, 585, 27)
    row, exact = nearest_sidewall_row(SIDEWALL_SUPPLY_TABLE, 14, 6)
    assert not exact and row is not None, 'nearest fallback returns something'

    assert min_row_for_cfm(SIDEWALL_RETURN_EXHAUST_TABLE, 440) == (12, 12, 440, 11)
    assert min_row_for_cfm(SIDEWALL_RETURN_EXHAUST_TABLE, 441) is not None
    assert min_row_for_cfm(SIDEWALL_RETURN_EXHAUST_TABLE, 441)[2] >= 441
    assert min_row_for_cfm(SIDEWALL_RETURN_EXHAUST_TABLE, 99999) is None
    assert largest_sidewall_row(SIDEWALL_SUPPLY_TABLE)[2] == 2485

    # Labels
    assert round_size_label(10) == '10"'
    assert rect_size_label(18, 12) == '18"x12"'
    assert row_size_label((18, 10, 735, 28)) == '18"x10"'
    assert round_size_label(None) == '-'

    print('All assertions passed.')


if __name__ == '__main__':
    _self_test()
