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

import shared_params


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
# Hot water return sizing, by circulation flow
# ---------------------------------------------------------------------------

def return_gpm_rows():
    """The firm hot water table restated in gpm, for sizing the return.

    Return piping carries circulation flow, not fixture units, so it cannot be
    looked up in a WSFU table directly. Each row's hot water limit (WSFU at
    5 fps) is converted to the gpm Table E103.3(3) gives for that load, flush
    tank column, which is the flow the firm table allows in that size
    (Colin, 2026-10-12). Derived at call time, never stored, so it cannot
    drift from the two tables it comes from.

    Returns:
        List of dicts: nominal_size, size_inches, max_wsfu, max_gpm. The
        first row is 1/2 inch, which is the minimum return size.
    """
    return firm_gpm_rows(SYSTEM_HOT)


def firm_gpm_rows(system):
    """The firm table for one system restated in gpm.

    Each row's WSFU limit is converted to the gpm Table E103.3(3) gives for
    that load, flush tank column: the flow the firm table allows in that size
    at its velocity limit. Used to size on gpm where WSFU cannot be looked up
    directly (hot water return, and cold water on a flush valve system).
    """
    rows = []
    limit_field = _SYSTEM_LIMIT_FIELDS[_system_key(system)]
    for row in _load()["firm_sizing_table"]:
        demand = wsfu_to_gpm(row[limit_field])
        rows.append({
            "nominal_size": row["nominal_size"],
            "size_inches": row["size_inches"],
            "max_wsfu": row[limit_field],
            "max_gpm": demand["gpm"],
        })
    return rows


def select_size_by_gpm(gpm, system):
    """Smallest firm size whose gpm capacity (firm_gpm_rows) covers gpm.

    Returns the select_size() dict shape: nominal_size, size_inches, gpm,
    system, limit_gpm, status ("ok", "no_demand", "exceeds_table" with
    max_table_gpm).
    """
    flow = _to_nonneg_float(gpm, "gpm")
    sys_key = _system_key(system)
    result = {"nominal_size": None, "size_inches": None, "gpm": flow,
              "system": sys_key, "limit_gpm": None, "status": "no_demand"}
    if flow == 0:
        return result
    rows = firm_gpm_rows(sys_key)
    for row in rows:
        if row["max_gpm"] is not None and row["max_gpm"] >= flow:
            result.update(nominal_size=row["nominal_size"],
                          size_inches=row["size_inches"],
                          limit_gpm=row["max_gpm"], status="ok")
            return result
    result["status"] = "exceeds_table"
    result["max_table_gpm"] = rows[-1]["max_gpm"]
    return result


def flush_valve_min_wsfu():
    """The first load the flush valve column of Table E103.3(3) lists."""
    return float(_load()[_CURVE_DATA_KEYS[CURVE_FLUSHOMETER]][0]["wsfu"])


def demand_gpm(wsfu, flush_valve_system=False):
    """Probable demand for a load, choosing the Table E103.3(3) column.

    On a flush valve system the flush valve column applies; a load below
    that column's first row (it gives no value there) takes the flush tank
    column. Returns dict: gpm, curve (the column used), status.
    """
    w = _to_nonneg_float(wsfu, "wsfu")
    if flush_valve_system and w >= flush_valve_min_wsfu():
        return wsfu_to_gpm(w, CURVE_FLUSHOMETER)
    return wsfu_to_gpm(w, CURVE_FLUSH_TANK)


def select_return_size(gpm):
    """Select the hot water return size for a circulation flow in gpm.

    Rule: the SMALLEST row of return_gpm_rows() whose max_gpm is >= gpm. The
    smallest row is 1/2 inch, so the 1/2 inch minimum return size needs no
    separate step.

    Returns:
        Dict with keys: nominal_size, size_inches, gpm, limit_gpm, status.
        status is "ok", "no_flow" (gpm == 0, sizes are None) or
        "exceeds_table" (sizes are None, "max_table_gpm" is included).

    Raises:
        ValueError: If gpm is negative or non-numeric.
    """
    flow = _to_nonneg_float(gpm, "gpm")
    result = {"nominal_size": None, "size_inches": None, "gpm": flow,
              "limit_gpm": None, "status": "no_flow"}
    if flow == 0:
        return result

    rows = return_gpm_rows()
    for row in rows:
        if row["max_gpm"] is not None and row["max_gpm"] >= flow:
            result["nominal_size"] = row["nominal_size"]
            result["size_inches"] = row["size_inches"]
            result["limit_gpm"] = row["max_gpm"]
            result["status"] = "ok"
            return result

    result["status"] = "exceeds_table"
    result["max_table_gpm"] = rows[-1]["max_gpm"]
    return result


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


def velocity_limit_fps(system):
    """The firm velocity limit for 'cold' or 'hot', in fps.

    Parsed from the same limit column the sizing uses, so the number cannot
    drift from the table in use. Hot water return uses the hot limit.
    """
    return _fps_from_limit_field(_SYSTEM_LIMIT_FIELDS[_system_key(system)])


def _fmt_fps(value):
    """Format a fps float with no trailing .0 (IronPython 2.7 safe: cast
    to float before '{:.0f}'.format(), since formatting a bare int with
    '{:.0f}' raises ValueError under IronPython 2.7)."""
    if value is None:
        return "?"
    return "{:.0f}".format(float(value))


# One wording for the return method, shared by the printed basis and the
# dialog rows so the two cannot drift. The heat loss and temperature drop are
# entered in the dialog, so their values are printed with the circuit setter
# table rather than here.
_RETURN_BASIS = (
    "One circuit setter per return leg. Its flow makes up the heat loss over "
    "the leg's effective length: gpm = BTU/h / (500 x temperature drop), " +
    "{:.1f}".format(float(shared_params.MIN_CIRCUIT_SETTER_GPM)) +
    " gpm minimum. Flows add at each junction toward the water heater.")


def _demand_sentence(demand_table_id):
    return (demand_table_id + ", flush tank column. When the system serves "
            "a flush valve fixture, cold water uses the flush valve column; "
            "a load below " + _fmt_fps(flush_valve_min_wsfu()) + " WSFU, the "
            "first row of that column, uses the flush tank column.")


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
    lines.append("Fixture loads: " + wsfu_table_id +
                 ", water supply fixture units (WSFU).")
    lines.append("Demand: " + _demand_sentence(demand_table_id))
    lines.append(
        "Pipe sizing: the firm WSFU per size table (" + pipe_material +
        "), " + str(size_min) + "\" through " + str(size_max) + "\", at a "
        "maximum velocity of " + _fmt_fps(cold_fps) + " fps for cold water "
        "and " + _fmt_fps(hot_fps) + " fps for hot water. On a flush valve "
        "system, cold water piping is sized on gpm against the same table "
        "restated in gpm.")
    lines.append(
        "Cold water piping is sized on total fixture units and hot water "
        "piping on hot fixture units. A cold water branch that serves only a "
        "water heater is sized on the hot fixture units that heater serves.")
    lines.append(
        "Hot water return: sized on circulation flow. " + _RETURN_BASIS +
        " The return is sized against the hot water table restated in gpm "
        "per " + demand_table_id + " (" + _fmt_fps(hot_fps) + " fps), " +
        str(size_min) + "\" minimum.")
    lines.append(
        "Velocity: computed from ASTM B88 Type L inside diameters and "
        "checked in Water Performance.")
    lines.append(
        "Pressure loss: 2024 IPC Section E103.3 segmented loss method, "
        "checked against Table 604.3 in Water Performance.")

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
        ("Code:", str(code_edition) + "."),
        ("Fixture loads:", wsfu_table_id +
         ", water supply fixture units (WSFU)."),
        ("Demand:", _demand_sentence(demand_table_id)),
        ("Pipe sizing:",
         "Firm WSFU per size table (" + pipe_material + "), " +
         str(size_min) + "\" through " + str(size_max) + "\". On a flush "
         "valve system, cold water piping is sized on gpm against the same "
         "table restated in gpm."),
        ("Max velocity:",
         _fmt_fps(cold_fps) + " fps for cold water, " + _fmt_fps(hot_fps) +
         " fps for hot water."),
        ("Sized on:",
         "Cold water piping on total fixture units, hot water piping on hot "
         "fixture units. A cold water branch that serves only a water heater "
         "is sized on the hot fixture units that heater serves."),
        ("Recirculation:",
         _RETURN_BASIS + " The return is sized against the hot water table "
         "restated in gpm per " + demand_table_id + " (" +
         _fmt_fps(hot_fps) + " fps), " + str(size_min) + "\" minimum."),
        ("Velocity:",
         "Computed from ASTM B88 Type L inside diameters and checked in "
         "Water Performance."),
        ("Pressure loss:",
         "2024 IPC Section E103.3 segmented loss method, checked against "
         "Table 604.3 in Water Performance."),
    ]


def sizing_chart():
    """The firm WSFU-to-pipe-size chart, as it is shown to the user.

    Built from the same rows select_size() and select_return_size() read, so
    the chart in the dialog and on the drafting view is the table the sizing
    actually used. One row per pipe size: the largest cold and hot fixture
    unit load that size carries, and the largest circulation flow it carries
    as a return.

    Returns:
        Dict with "headers" (4 strings) and "rows" (lists of 4 strings).
    """
    firm_rows = _load()["firm_sizing_table"]
    cold_field = _SYSTEM_LIMIT_FIELDS[SYSTEM_COLD]
    hot_field = _SYSTEM_LIMIT_FIELDS[SYSTEM_HOT]
    returns = return_gpm_rows()

    headers = [
        "PIPE SIZE",
        "COLD WSFU (" + _fmt_fps(_fps_from_limit_field(cold_field)) + " FPS)",
        "HOT WSFU (" + _fmt_fps(_fps_from_limit_field(hot_field)) + " FPS)",
        "RETURN GPM",
    ]
    rows = []
    for row, ret in zip(firm_rows, returns):
        rows.append([
            str(row["nominal_size"]) + "\"",
            "up to " + fmt_number(row[cold_field]),
            "up to " + fmt_number(row[hot_field]),
            ("-" if ret["max_gpm"] is None
             else "up to " + "{:.1f}".format(float(ret["max_gpm"]))),
        ])
    return {"headers": headers, "rows": rows}


def fmt_number(value):
    """A table limit without a trailing .0 for whole numbers. Cast to float
    first: IronPython 2.7 rejects '{:.0f}'.format(an_int)."""
    number = float(value)
    if abs(number - round(number)) < 1e-9:
        return "{:.0f}".format(number)
    return "{:g}".format(number)


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
