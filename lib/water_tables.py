# -*- coding: ascii -*-
"""
water_tables.py

2024 IPC domestic water sizing table loader and lookup.

Loads ipc_water_tables.json (co-located in lib/) and provides lookup
functions used by the Domestic Water Tools (Size Water, WSFU schedule,
Diagnose Water, etc.).

Pure data module. No Revit API calls. Never raises for ordinary
out-of-range data (the caller logs and flags it); raises ValueError only
for genuinely invalid input, such as a negative fixture-unit load or an
unrecognized system/curve name.

Tables loaded from the JSON (see _source in the JSON for provenance):
    wsfu_by_fixture         Table E103.3(2)  fixture -> WSFU load
    demand_flush_tank       Table E103.3(3)  WSFU -> gpm, flush tank column
    demand_flushometer_valve Table E103.3(3) WSFU -> gpm, flushometer column
    fixture_flow_604_3      Table 604.3      fixture -> flow rate/pressure
    fixture_supply_604_5    Table 604.5      fixture -> min supply pipe size
    firm_sizing_table       RJA firm WSFU-per-pipe-size table (Copper Type L,
                             8 fps cold / 5 fps hot), from the firm's Excel
                             sizing calculator, used as data (not regenerated)

IronPython 2.7 compatible. Pure ASCII source.
"""

import os
import re
import json


# ---------------------------------------------------------------------------
# Module-level data load
# ---------------------------------------------------------------------------

_HERE = os.path.dirname(os.path.abspath(__file__))
_JSON_PATH = os.path.join(_HERE, "ipc_water_tables.json")

_DATA = None  # Populated by _load() on first access

_REQUIRED_KEYS = (
    "_source",
    "wsfu_by_fixture",
    "demand_flush_tank",
    "demand_flushometer_valve",
    "fixture_flow_604_3",
    "fixture_supply_604_5",
    "firm_sizing_table",
)


def _load():
    """Load the JSON file once and cache it at module level."""
    global _DATA
    if _DATA is not None:
        return _DATA

    if not os.path.isfile(_JSON_PATH):
        raise IOError(
            "IPC water sizing tables JSON not found at: " + _JSON_PATH +
            ". Expected ipc_water_tables.json in the lib/ folder next to "
            "water_tables.py."
        )

    f = open(_JSON_PATH, "r")
    try:
        data = json.load(f)
    finally:
        f.close()

    for key in _REQUIRED_KEYS:
        if key not in data:
            raise ValueError(
                "IPC water sizing tables JSON is malformed: missing '" +
                key + "' key."
            )

    _DATA = data
    return _DATA


# ---------------------------------------------------------------------------
# Public constants
# ---------------------------------------------------------------------------

SYSTEM_COLD = "cold"
SYSTEM_HOT = "hot"

CURVE_FLUSH_TANK = "flush_tank"
CURVE_FLUSHOMETER = "flushometer_valve"

_CURVE_DATA_KEYS = {
    CURVE_FLUSH_TANK: "demand_flush_tank",
    CURVE_FLUSHOMETER: "demand_flushometer_valve",
}

_SYSTEM_LIMIT_FIELDS = {
    SYSTEM_COLD: "max_wsfu_8fps_cold",
    SYSTEM_HOT: "max_wsfu_5fps_hot",
}


def _system_key(system):
    """Normalize a 'cold'/'hot' system string. Raises ValueError otherwise."""
    if system is None:
        raise ValueError("system must not be None. Expected 'cold' or 'hot'.")
    key = str(system).strip().lower()
    if key not in _SYSTEM_LIMIT_FIELDS:
        raise ValueError(
            "Unknown system: '" + str(system) + "'. Expected '" +
            SYSTEM_COLD + "' or '" + SYSTEM_HOT + "'."
        )
    return key


def _curve_key(curve):
    """Normalize a demand curve name. Raises ValueError otherwise."""
    if curve is None:
        raise ValueError(
            "curve must not be None. Expected '" + CURVE_FLUSH_TANK +
            "' or '" + CURVE_FLUSHOMETER + "'."
        )
    key = str(curve).strip().lower()
    if key not in _CURVE_DATA_KEYS:
        raise ValueError(
            "Unknown curve: '" + str(curve) + "'. Expected '" +
            CURVE_FLUSH_TANK + "' or '" + CURVE_FLUSHOMETER + "'."
        )
    return key


def _to_nonneg_float(value, label):
    """Cast to float and require >= 0. Raises ValueError otherwise."""
    if value is None:
        raise ValueError(label + " must not be None.")
    try:
        v = float(value)
    except (TypeError, ValueError):
        raise ValueError(label + " must be numeric, got: " + repr(value))
    if v < 0:
        raise ValueError(
            label + " must not be negative, got: " + str(v) + "."
        )
    return v


# ---------------------------------------------------------------------------
# Firm pipe sizing table lookup
# ---------------------------------------------------------------------------

def select_size(wsfu, system):
    """Select the firm pipe size for a fixture-unit load on one system.

    Rule: pick the SMALLEST firm_sizing_table row whose limit column is
    >= wsfu. The table's limits are upper bounds, inclusive.

    NOTE: the source Excel workbook's own IF() formulas have a boundary
    defect -- they test "<=X" then ">X.01", which leaves a gap (e.g. a
    load of 1.005 falls through every branch and the formula evaluates
    to FALSE). This function does NOT reproduce that defect: every
    non-negative wsfu maps to a real size, or to "exceeds_table" only
    when it is genuinely above the largest published limit.

    Args:
        wsfu: Fixture-unit load (float or int), >= 0.
        system: "cold" (uses max_wsfu_8fps_cold) or "hot" (uses
            max_wsfu_5fps_hot). Case-insensitive.

    Returns:
        Dict with keys: nominal_size, size_inches, wsfu, system,
        limit_wsfu, status. status is one of:
            "ok"             - a size was found; nominal_size/size_inches
                                 set, limit_wsfu is the matched row's limit.
            "no_demand"      - wsfu == 0; nominal_size/size_inches/
                                 limit_wsfu are None.
            "exceeds_table"  - wsfu is above the largest published limit;
                                 nominal_size/size_inches/limit_wsfu are
                                 None, and "max_table_wsfu" is included.

    Raises:
        ValueError: If wsfu is negative/non-numeric, or system is not
            "cold" or "hot".
    """
    w = _to_nonneg_float(wsfu, "wsfu")
    sys_key = _system_key(system)
    limit_field = _SYSTEM_LIMIT_FIELDS[sys_key]

    if w == 0:
        return {
            "nominal_size": None,
            "size_inches": None,
            "wsfu": w,
            "system": sys_key,
            "limit_wsfu": None,
            "status": "no_demand",
        }

    rows = _load()["firm_sizing_table"]

    for row in rows:
        limit = row[limit_field]
        if limit >= w:
            return {
                "nominal_size": row["nominal_size"],
                "size_inches": row["size_inches"],
                "wsfu": w,
                "system": sys_key,
                "limit_wsfu": limit,
                "status": "ok",
            }

    max_limit = rows[-1][limit_field]
    return {
        "nominal_size": None,
        "size_inches": None,
        "wsfu": w,
        "system": sys_key,
        "limit_wsfu": None,
        "status": "exceeds_table",
        "max_table_wsfu": max_limit,
    }


def firm_sizing_rows():
    """Return a mutation-safe copy of the 8 firm_sizing_table rows."""
    return [dict(row) for row in _load()["firm_sizing_table"]]


# ---------------------------------------------------------------------------
# Table E103.3(2) - WSFU by fixture
# ---------------------------------------------------------------------------

def wsfu_rows():
    """Return a mutation-safe copy of the 27 wsfu_by_fixture rows.

    cold_wsfu / hot_wsfu may be None (printed as a dash on Table
    E103.3(2) for fixtures that don't use that system).
    """
    return [dict(row) for row in _load()["wsfu_by_fixture"]]


# ---------------------------------------------------------------------------
# Table E103.3(3) - WSFU -> gpm demand
# ---------------------------------------------------------------------------

def wsfu_to_gpm(wsfu, curve=CURVE_FLUSH_TANK):
    """Convert a WSFU load to gpm demand via linear interpolation.

    Args:
        wsfu: Fixture-unit load (float or int), >= 0.
        curve: "flush_tank" or "flushometer_valve". Case-insensitive.

    Returns:
        Dict with keys: wsfu, gpm, curve, status. status is one of:
            "ok"          - exact row match or interpolated between rows.
            "below_table" - wsfu is below the curve's first row; gpm is
                             clamped to the first row's gpm.
            "exceeds_table" - wsfu is above the curve's last row (5000);
                             gpm is None.

    Raises:
        ValueError: If wsfu is negative/non-numeric, or curve is not
            "flush_tank" or "flushometer_valve".
    """
    w = _to_nonneg_float(wsfu, "wsfu")
    curve_key = _curve_key(curve)
    rows = _load()[_CURVE_DATA_KEYS[curve_key]]

    first = rows[0]
    last = rows[-1]

    if w < first["wsfu"]:
        return {
            "wsfu": w,
            "gpm": first["gpm"],
            "curve": curve_key,
            "status": "below_table",
        }

    if w > last["wsfu"]:
        return {
            "wsfu": w,
            "gpm": None,
            "curve": curve_key,
            "status": "exceeds_table",
        }

    prev_row = None
    for row in rows:
        if row["wsfu"] == w:
            return {
                "wsfu": w,
                "gpm": row["gpm"],
                "curve": curve_key,
                "status": "ok",
            }
        if row["wsfu"] > w:
            # prev_row is guaranteed set here: w >= first["wsfu"] and did
            # not exact-match any row seen so far.
            lo, hi = prev_row, row
            span = float(hi["wsfu"] - lo["wsfu"])
            frac = float(w - lo["wsfu"]) / span
            gpm = lo["gpm"] + frac * (hi["gpm"] - lo["gpm"])
            return {
                "wsfu": w,
                "gpm": gpm,
                "curve": curve_key,
                "status": "ok",
            }
        prev_row = row

    # Unreachable given the bounds checks above; kept as a safety net.
    return {
        "wsfu": w,
        "gpm": last["gpm"],
        "curve": curve_key,
        "status": "ok",
    }


# ---------------------------------------------------------------------------
# Table 604.3 / Table 604.5 - fixture flow and minimum supply size
# ---------------------------------------------------------------------------

def min_fixture_supply(fixture_name):
    """Return the Table 604.5 minimum supply pipe size string for a fixture.

    Args:
        fixture_name: Fixture name, matched case-insensitively against the
            "fixture" field of fixture_supply_604_5.

    Returns:
        The "min_supply_in" string (e.g. "3/8"), or None if no exact
        case-insensitive match is found.
    """
    if fixture_name is None:
        return None
    target = str(fixture_name).strip().lower()
    for row in _load()["fixture_supply_604_5"]:
        if str(row["fixture"]).strip().lower() == target:
            return row["min_supply_in"]
    return None


def fixture_flow(fixture_name):
    """Return the Table 604.3 flow rate/pressure dict for a fixture.

    Args:
        fixture_name: Fixture name, matched case-insensitively against the
            "fixture" field of fixture_flow_604_3.

    Returns:
        A mutation-safe copy of the matching row dict (keys: fixture,
        flow_gpm, flow_pressure_psi), or None if no exact case-insensitive
        match is found.
    """
    if fixture_name is None:
        return None
    target = str(fixture_name).strip().lower()
    for row in _load()["fixture_flow_604_3"]:
        if str(row["fixture"]).strip().lower() == target:
            return dict(row)
    return None


# ---------------------------------------------------------------------------
# Basis of design
# ---------------------------------------------------------------------------

_FPS_RE = re.compile(r"(\d+)fps")


def _fps_from_limit_field(field_name):
    """Extract the numeric fps value embedded in a firm_sizing_table limit
    column name, e.g. 'max_wsfu_8fps_cold' -> 8.0. Returns None if the
    field name doesn't contain a '<N>fps' token.
    """
    m = _FPS_RE.search(field_name)
    if m is None:
        return None
    return float(m.group(1))


def _fmt_fps(value):
    """Format a fps float with no trailing .0 (IronPython 2.7 safe: cast
    to float before '{:.0f}'.format(), since formatting a bare int with
    '{:.0f}' raises ValueError under IronPython 2.7)."""
    if value is None:
        return "?"
    return "{:.0f}".format(float(value))


def basis_of_design_lines(extra=None):
    """Build the basis-of-design text block every Domestic Water tool
    output shall print, generated from the loaded tables so it cannot
    drift from the data actually in use.

    Args:
        extra: Optional list of additional ASCII strings to append
            (e.g. project-specific notes).

    Returns:
        List of plain ASCII strings, no trailing newlines.
    """
    data = _load()
    source = data["_source"]
    firm_rows = data["firm_sizing_table"]

    code_edition = source.get("code", "2024 International Plumbing Code")

    wsfu_table_id = str(source.get("wsfu_table", "")).split(",")[0].strip()
    if not wsfu_table_id:
        wsfu_table_id = "Table E103.3(2)"

    demand_table_id = str(source.get("demand_table", "")).split(",")[0].strip()
    if not demand_table_id:
        demand_table_id = "Table E103.3(3)"

    firm_source_parts = [
        p.strip() for p in str(source.get("firm_sizing_table", "")).split(",")
    ]
    pipe_material = firm_source_parts[-1] if firm_source_parts[-1] else "Copper Type L"

    cold_fps = None
    hot_fps = None
    for key in firm_rows[0].keys():
        if key.endswith("_cold"):
            cold_fps = _fps_from_limit_field(key)
        elif key.endswith("_hot"):
            hot_fps = _fps_from_limit_field(key)

    size_min = firm_rows[0]["nominal_size"]
    size_max = firm_rows[-1]["nominal_size"]

    lines = []
    lines.append("Basis of design: " + str(code_edition) + ".")
    lines.append(
        "Fixture loads per " + wsfu_table_id + " (Water Supply Fixture Units)."
    )
    lines.append(
        "Demand (gpm) per " + demand_table_id + ", flush tank column."
    )
    lines.append(
        "Pipe sizing per the firm WSFU-per-size table (" + pipe_material +
        "), nominal sizes " + str(size_min) + "\" through " + str(size_max) +
        "\", maximum velocity " + _fmt_fps(cold_fps) + " fps cold water and " +
        _fmt_fps(hot_fps) + " fps hot water."
    )
    lines.append(
        "Cold water piping is sized on TOTAL fixture units; hot water "
        "piping is sized on HOT fixture units."
    )
    lines.append(
        "A cold water branch serving only a water heater is sized on the "
        "TOTAL HOT fixture units that heater serves."
    )
    lines.append("Pressure loss is NOT evaluated in this build.")

    if extra:
        for line in extra:
            lines.append(str(line))

    return lines


def basis_of_design_rows():
    """The same basis of design, as (label, value) pairs for the dialog.

    The printed report wants sentences, the Size Water dialog wants the
    label / value layout the Duct Velocity dialog uses for its Calculation
    Basis block. Both are generated here from the loaded tables so the
    dialog can never show a basis different from the one the sizing used.

    Returns:
        List of (label, value) ASCII string pairs.
    """
    data = _load()
    source = data["_source"]
    firm_rows = data["firm_sizing_table"]

    code_edition = source.get("code", "2024 International Plumbing Code")

    wsfu_table_id = str(source.get("wsfu_table", "")).split(",")[0].strip()
    if not wsfu_table_id:
        wsfu_table_id = "Table E103.3(2)"

    demand_table_id = str(source.get("demand_table", "")).split(",")[0].strip()
    if not demand_table_id:
        demand_table_id = "Table E103.3(3)"

    firm_source_parts = [
        p.strip() for p in str(source.get("firm_sizing_table", "")).split(",")
    ]
    pipe_material = (firm_source_parts[-1] if firm_source_parts[-1]
                     else "Copper Type L")

    cold_fps = None
    hot_fps = None
    for key in firm_rows[0].keys():
        if key.endswith("_cold"):
            cold_fps = _fps_from_limit_field(key)
        elif key.endswith("_hot"):
            hot_fps = _fps_from_limit_field(key)

    size_min = firm_rows[0]["nominal_size"]
    size_max = firm_rows[-1]["nominal_size"]

    return [
        ("Code:", str(code_edition)),
        ("Fixture loads:", wsfu_table_id + ", Water Supply Fixture Units"),
        ("Demand (gpm):", demand_table_id + ", flush tank column"),
        ("Pipe sizing:",
         "firm WSFU-per-size table (" + pipe_material + "), nominal " +
         str(size_min) + "\" through " + str(size_max) + "\""),
        ("Max velocity:",
         _fmt_fps(cold_fps) + " fps cold water, " + _fmt_fps(hot_fps) +
         " fps hot water"),
        ("Sized on:",
         "cold water on TOTAL fixture units, hot water on HOT fixture "
         "units. A cold branch serving only a water heater is sized on the "
         "TOTAL HOT fixture units that heater serves."),
        ("Not included:",
         "pressure loss and velocity are NOT evaluated in this build. The "
         "hot water return is not sized, it is sized on circulation flow "
         "rather than on fixture units."),
    ]


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

def _self_test():
    """Run a non-Revit self-test against known table values."""
    print("water_tables.py self-test")
    print("JSON path: " + _JSON_PATH)

    r = select_size(2.0, "cold")
    print("select_size(2.0, 'cold') -> " + str(r["nominal_size"]))
    assert r["nominal_size"] == "1/2", "Expected 1/2"

    r2 = select_size(1.005, "cold")
    print("select_size(1.005, 'cold') -> " + str(r2["nominal_size"]))
    assert r2["nominal_size"] == "1/2", "Spreadsheet gap case failed"

    r3 = wsfu_to_gpm(10, "flush_tank")
    print("wsfu_to_gpm(10, 'flush_tank') -> " + str(r3["gpm"]))
    assert r3["gpm"] == 14.6, "Expected 14.6 gpm"

    supply = min_fixture_supply("Lavatory")
    print("min_fixture_supply('Lavatory') -> " + str(supply))
    assert supply == "3/8", "Expected 3/8"

    flow = fixture_flow("Lavatory, public")
    print("fixture_flow('Lavatory, public') -> " + str(flow))
    assert flow["flow_gpm"] == 0.4, "Expected 0.4 gpm"

    lines = basis_of_design_lines()
    print("basis_of_design_lines() -> " + str(len(lines)) + " lines")
    for line in lines:
        print("  " + line)

    print("All assertions passed.")


if __name__ == "__main__":
    _self_test()
