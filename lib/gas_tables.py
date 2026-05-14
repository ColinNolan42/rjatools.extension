# -*- coding: ascii -*-
"""
gas_tables.py

IFGC gas pipe sizing table loader and lookup.

Loads ifgc_gas_sizing_tables.json (co-located in lib/) and provides
lookup functions used by sizing_engine.py.

Pure data module. No Revit API calls. Raises ValueError on bad inputs
so the caller can log and surface the error.

Per 2024 IFGC Appendix A, Section A103.1 Step 5: when the actual pipe
length does not match a row in the table, the NEXT LONGER row is used.
This is conservative (lower capacity for the same pipe size).

Supported materials and tables:
    Schedule 40 Steel   402.4(1)-(7) natural, 402.4(25)-(28) propane
    Semirigid Copper    402.4(8)-(14) natural, 402.4(29)-(31) propane
    CSST                402.4(15)-(19) natural, 402.4(32)-(34) propane
    PE Plastic Pipe     402.4(20)-(24) natural, 402.4(35)-(37) propane

IronPython 2.7 compatible. Pure ASCII source.
"""

import os
import json


# ---------------------------------------------------------------------------
# Module-level data load
# ---------------------------------------------------------------------------

_HERE = os.path.dirname(os.path.abspath(__file__))
_JSON_PATH = os.path.join(_HERE, "ifgc_gas_sizing_tables.json")

_DATA = None  # Populated by _load() on first access


def _load():
    """Load the JSON file once and cache it at module level."""
    global _DATA
    if _DATA is not None:
        return _DATA

    if not os.path.isfile(_JSON_PATH):
        raise IOError(
            "IFGC sizing tables JSON not found at: " + _JSON_PATH +
            ". Expected ifgc_gas_sizing_tables.json in the lib/ folder "
            "next to gas_tables.py."
        )

    f = open(_JSON_PATH, "r")
    try:
        _DATA = json.load(f)
    finally:
        f.close()

    if "tables" not in _DATA:
        raise ValueError(
            "IFGC sizing tables JSON is malformed: missing 'tables' key."
        )

    return _DATA


def _tables():
    """Return the tables dict from the loaded JSON."""
    return _load()["tables"]


# ---------------------------------------------------------------------------
# Material support
# ---------------------------------------------------------------------------

SUPPORTED_MATERIALS = (
    "Schedule 40 Steel",
    "Semirigid Copper Tubing",
    "CSST",
    "Polyethylene Plastic Pipe",
)


# ---------------------------------------------------------------------------
# Table options - used by Size Gas dialog to populate the sizing table list.
# Each entry maps a human-readable label to a specific IFGC table.
# "inlet_pressure_psi" is None for <2 psi tables (pressure is below threshold).
# ---------------------------------------------------------------------------

TABLE_OPTIONS = [
    # --- Schedule 40 Steel - Natural gas ---
    {
        "label":   "Sched 40 Steel - Natural, <2 psi, 0.3\" w.c.  [402.4(1)]",
        "table_id": "402.4(1)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Sched 40 Steel - Natural, <2 psi, 0.5\" w.c.  [402.4(2)]",
        "table_id": "402.4(2)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Sched 40 Steel - Natural, <2 psi, 1.0\" w.c.  [402.4(3)]",
        "table_id": "402.4(3)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Sched 40 Steel - Natural, 2 psi, 1 psi drop  [402.4(4)]",
        "table_id": "402.4(4)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 2.0,
    },
    {
        "label":   "Sched 40 Steel - Natural, 2 psi, 2 psi drop  [402.4(5)]",
        "table_id": "402.4(5)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 2.0,
    },
    {
        "label":   "Sched 40 Steel - Natural, 5 psi, 3.5 psi drop  [402.4(6)]",
        "table_id": "402.4(6)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 5.0,
    },
    {
        "label":   "Sched 40 Steel - Natural, 10 psi, 1 psi drop  [402.4(7)]",
        "table_id": "402.4(7)",
        "material": "Schedule 40 Steel",
        "gas":      "Natural",
        "inlet_pressure_psi": 10.0,
    },
    # --- Schedule 40 Steel - Propane ---
    {
        "label":   "Sched 40 Steel - Propane, <2 psi, 0.3\" w.c.  [402.4(25)]",
        "table_id": "402.4(25)",
        "material": "Schedule 40 Steel",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Sched 40 Steel - Propane, <2 psi, 0.5\" w.c.  [402.4(26)]",
        "table_id": "402.4(26)",
        "material": "Schedule 40 Steel",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Sched 40 Steel - Propane, <2 psi, 1.0\" w.c.  [402.4(27)]",
        "table_id": "402.4(27)",
        "material": "Schedule 40 Steel",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Sched 40 Steel - Propane, 2 psi, 1 psi drop  [402.4(28)]",
        "table_id": "402.4(28)",
        "material": "Schedule 40 Steel",
        "gas":      "Propane",
        "inlet_pressure_psi": 2.0,
    },
    # --- Semirigid Copper Tubing - Natural gas ---
    {
        "label":   "Copper Tubing - Natural, <2 psi, 0.3\" w.c.  [402.4(8)]",
        "table_id": "402.4(8)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Copper Tubing - Natural, <2 psi, 0.5\" w.c.  [402.4(9)]",
        "table_id": "402.4(9)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Copper Tubing - Natural, <2 psi, 1.0\" w.c.  [402.4(10)]",
        "table_id": "402.4(10)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Copper Tubing - Natural, 2 psi, 1 psi drop  [402.4(11)]",
        "table_id": "402.4(11)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 2.0,
    },
    {
        "label":   "Copper Tubing - Natural, 2 psi, 2 psi drop  [402.4(12)]",
        "table_id": "402.4(12)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 2.0,
    },
    {
        "label":   "Copper Tubing - Natural, 5 psi, 3.5 psi drop  [402.4(13)]",
        "table_id": "402.4(13)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 5.0,
    },
    {
        "label":   "Copper Tubing - Natural, 10 psi, 1 psi drop  [402.4(14)]",
        "table_id": "402.4(14)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Natural",
        "inlet_pressure_psi": 10.0,
    },
    # --- Semirigid Copper Tubing - Propane ---
    {
        "label":   "Copper Tubing - Propane, <2 psi, 0.3\" w.c.  [402.4(29)]",
        "table_id": "402.4(29)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Copper Tubing - Propane, <2 psi, 0.5\" w.c.  [402.4(30)]",
        "table_id": "402.4(30)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "Copper Tubing - Propane, <2 psi, 1.0\" w.c.  [402.4(31)]",
        "table_id": "402.4(31)",
        "material": "Semirigid Copper Tubing",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    # --- CSST - Natural gas ---
    {
        "label":   "CSST - Natural, <2 psi, 0.5\" w.c.  [402.4(15)]",
        "table_id": "402.4(15)",
        "material": "CSST",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "CSST - Natural, <2 psi, 0.3\" w.c.  [402.4(16)]",
        "table_id": "402.4(16)",
        "material": "CSST",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "CSST - Natural, <2 psi, 1.0\" w.c.  [402.4(17)]",
        "table_id": "402.4(17)",
        "material": "CSST",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "CSST - Natural, 2 psi, 1 psi drop  [402.4(18)]",
        "table_id": "402.4(18)",
        "material": "CSST",
        "gas":      "Natural",
        "inlet_pressure_psi": 2.0,
    },
    {
        "label":   "CSST - Natural, 5 psi, 3.5 psi drop  [402.4(19)]",
        "table_id": "402.4(19)",
        "material": "CSST",
        "gas":      "Natural",
        "inlet_pressure_psi": 5.0,
    },
    # --- CSST - Propane ---
    {
        "label":   "CSST - Propane, <2 psi, 0.3\" w.c.  [402.4(32)]",
        "table_id": "402.4(32)",
        "material": "CSST",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "CSST - Propane, <2 psi, 0.5\" w.c.  [402.4(33)]",
        "table_id": "402.4(33)",
        "material": "CSST",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "CSST - Propane, 2 psi, 1 psi drop  [402.4(34)]",
        "table_id": "402.4(34)",
        "material": "CSST",
        "gas":      "Propane",
        "inlet_pressure_psi": 2.0,
    },
    # --- Polyethylene Plastic Pipe - Natural gas ---
    {
        "label":   "PE Plastic Pipe - Natural, <2 psi, 0.3\" w.c.  [402.4(20)]",
        "table_id": "402.4(20)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "PE Plastic Pipe - Natural, <2 psi, 0.5\" w.c.  [402.4(21)]",
        "table_id": "402.4(21)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "PE Plastic Pipe - Natural, <2 psi, 1.0\" w.c.  [402.4(22)]",
        "table_id": "402.4(22)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Natural",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "PE Plastic Pipe - Natural, 2 psi, 1 psi drop  [402.4(23)]",
        "table_id": "402.4(23)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Natural",
        "inlet_pressure_psi": 2.0,
    },
    {
        "label":   "PE Plastic Pipe - Natural, 5 psi, 3.5 psi drop  [402.4(24)]",
        "table_id": "402.4(24)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Natural",
        "inlet_pressure_psi": 5.0,
    },
    # --- Polyethylene Plastic Pipe - Propane ---
    {
        "label":   "PE Plastic Pipe - Propane, <2 psi, 0.3\" w.c.  [402.4(35)]",
        "table_id": "402.4(35)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "PE Plastic Pipe - Propane, <2 psi, 0.5\" w.c.  [402.4(36)]",
        "table_id": "402.4(36)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Propane",
        "inlet_pressure_psi": 0.7,
    },
    {
        "label":   "PE Plastic Pipe - Propane, 2 psi, 1 psi drop  [402.4(37)]",
        "table_id": "402.4(37)",
        "material": "Polyethylene Plastic Pipe",
        "gas":      "Propane",
        "inlet_pressure_psi": 2.0,
    },
]

# Label -> option dict for fast lookup by label string
_LABEL_INDEX = {opt["label"]: opt for opt in TABLE_OPTIONS}


def get_table_options():
    """Return the full list of TABLE_OPTIONS dicts."""
    return list(TABLE_OPTIONS)


def get_table_option_labels():
    """Return just the label strings, in order."""
    return [opt["label"] for opt in TABLE_OPTIONS]


def get_table_option_by_label(label):
    """Return the TABLE_OPTIONS entry for a given label string.

    Args:
        label: Label string matching an entry in TABLE_OPTIONS.

    Returns:
        Dict with keys: label, table_id, material, gas, inlet_pressure_psi.

    Raises:
        ValueError: If label is not found.
    """
    opt = _LABEL_INDEX.get(label)
    if opt is None:
        raise ValueError(
            "Unknown sizing table option: '" + str(label) + "'."
        )
    return opt


# ---------------------------------------------------------------------------
# Table selection - legacy helper (Schedule 40 Steel default mapping)
# ---------------------------------------------------------------------------

def select_table(inlet_pressure_psi, pipe_material="Schedule 40 Steel"):
    """Select the default IFGC sizing table based on inlet pressure.

    Returns the most commonly used table for the given material and pressure.
    For precise table selection (specific pressure drop), use TABLE_OPTIONS
    and get_table_option_by_label() instead.

    Schedule 40 Steel:
        inlet_pressure < 2.0    -> 402.4(2)   0.5 in. w.c. drop (natural gas)
        inlet_pressure == 2.0   -> 402.4(4)   1.0 PSI drop
        inlet_pressure == 5.0   -> 402.4(6)   3.5 PSI drop
        inlet_pressure == 10.0  -> 402.4(7)   1.0 PSI drop

    Semirigid Copper Tubing:
        inlet_pressure < 2.0    -> 402.4(9)   0.5 in. w.c. drop
        inlet_pressure == 2.0   -> 402.4(11)  1.0 PSI drop
        inlet_pressure == 5.0   -> 402.4(13)  3.5 PSI drop
        inlet_pressure == 10.0  -> 402.4(14)  1.0 PSI drop

    CSST:
        inlet_pressure < 2.0    -> 402.4(15)  0.5 in. w.c. drop
        inlet_pressure == 2.0   -> 402.4(18)  1.0 PSI drop
        inlet_pressure == 5.0   -> 402.4(19)  3.5 PSI drop

    Polyethylene Plastic Pipe:
        inlet_pressure < 2.0    -> 402.4(20)  0.3 in. w.c. drop
        inlet_pressure == 2.0   -> 402.4(23)  1.0 PSI drop
        inlet_pressure == 5.0   -> 402.4(24)  3.5 PSI drop

    Args:
        inlet_pressure_psi: Gas supply pressure at the meter, in PSI.
        pipe_material: Pipe material string from SUPPORTED_MATERIALS.

    Returns:
        Table ID string, e.g. "402.4(2)".

    Raises:
        ValueError: If pipe material is unsupported or pressure not supported.
    """
    if pipe_material not in SUPPORTED_MATERIALS:
        raise ValueError(
            "Unsupported pipe material: '" + str(pipe_material) + "'. "
            "Supported: " + ", ".join(SUPPORTED_MATERIALS) + "."
        )

    if inlet_pressure_psi is None:
        raise ValueError("Inlet pressure must not be None.")

    try:
        p = float(inlet_pressure_psi)
    except (TypeError, ValueError):
        raise ValueError(
            "Inlet pressure must be numeric, got: " + repr(inlet_pressure_psi)
        )

    if p <= 0:
        raise ValueError(
            "Inlet pressure must be positive, got: " + str(p) + " PSI."
        )

    if pipe_material == "Schedule 40 Steel":
        if p < 2.0:
            table_id = "402.4(2)"
        elif p <= 2.0:
            table_id = "402.4(4)"
        elif p <= 5.0:
            table_id = "402.4(6)"
        elif p <= 10.0:
            table_id = "402.4(7)"
        else:
            raise ValueError(
                "No IFGC table for Schedule 40 Steel at " + str(p) + " PSI. "
                "Max supported: 10 PSI."
            )

    elif pipe_material == "Semirigid Copper Tubing":
        if p < 2.0:
            table_id = "402.4(9)"
        elif p <= 2.0:
            table_id = "402.4(11)"
        elif p <= 5.0:
            table_id = "402.4(13)"
        elif p <= 10.0:
            table_id = "402.4(14)"
        else:
            raise ValueError(
                "No IFGC table for Copper Tubing at " + str(p) + " PSI. "
                "Max supported: 10 PSI."
            )

    elif pipe_material == "CSST":
        if p < 2.0:
            table_id = "402.4(15)"
        elif p <= 2.0:
            table_id = "402.4(18)"
        elif p <= 5.0:
            table_id = "402.4(19)"
        else:
            raise ValueError(
                "No IFGC table for CSST at " + str(p) + " PSI. "
                "Max supported: 5 PSI."
            )

    elif pipe_material == "Polyethylene Plastic Pipe":
        if p < 2.0:
            table_id = "402.4(20)"
        elif p <= 2.0:
            table_id = "402.4(23)"
        elif p <= 5.0:
            table_id = "402.4(24)"
        else:
            raise ValueError(
                "No IFGC table for PE Pipe at " + str(p) + " PSI. "
                "Max supported: 5 PSI."
            )

    else:
        raise ValueError("Unsupported material: " + str(pipe_material))

    # Sanity check: the selected table must exist in the JSON.
    tables = _tables()
    if table_id not in tables:
        raise ValueError(
            "Selected table '" + table_id + "' is not present in the loaded "
            "IFGC tables JSON."
        )

    return table_id


# ---------------------------------------------------------------------------
# Table accessors
# ---------------------------------------------------------------------------

def get_table(table_id):
    """Return the raw table dict for the given table ID.

    Args:
        table_id: Table identifier string, e.g. "402.4(2)".

    Returns:
        Dict with keys: table_id, description, gas, inlet_pressure,
        pressure_drop, specific_gravity, units, pipe_sizes_nominal,
        pipe_sizes_actual_id, data.

    Raises:
        ValueError: If table_id is not present in the JSON.
    """
    tables = _tables()
    if table_id not in tables:
        raise ValueError(
            "Unknown table ID: '" + str(table_id) + "'. "
            "Available tables: " + ", ".join(sorted(tables.keys())) + "."
        )
    return tables[table_id]


def list_pipe_sizes(table_id):
    """Return the list of nominal pipe size strings for a table.

    Order matches the JSON: ascending from smallest to largest.
    sizing_engine.py iterates this list to find the smallest pipe whose
    capacity meets the demand.

    Args:
        table_id: Table identifier string.

    Returns:
        List of strings, e.g. ["1/2", "3/4", "1", "1-1/4", ...].

    Raises:
        ValueError: If table_id is not present.
    """
    return list(get_table(table_id)["pipe_sizes_nominal"])


# ---------------------------------------------------------------------------
# Length row lookup (IFGC A103.1 Step 5)
# ---------------------------------------------------------------------------

def get_length_row(table_id, length_ft):
    """Return the table row to use for a given pipe length.

    Per IFGC A103.1 Step 5: if the actual length does not match a row,
    use the NEXT LONGER row. This is conservative because longer rows
    have lower capacities for the same pipe size (more friction loss).

    Args:
        table_id: Table identifier string.
        length_ft: Actual pipe length (typically the longest run), in feet.

    Returns:
        Tuple of (table_length_used_ft, capacities_list).
        capacities_list is the raw row from the JSON: a list of CFH
        values (or None for unavailable sizes), in the same order as
        pipe_sizes_nominal for that table.

    Raises:
        ValueError: If table_id is unknown, length is invalid, or length
            exceeds the maximum row in the table.
    """
    if length_ft is None:
        raise ValueError("Length must not be None.")

    try:
        L = float(length_ft)
    except (TypeError, ValueError):
        raise ValueError("Length must be numeric, got: " + repr(length_ft))

    if L <= 0:
        raise ValueError("Length must be positive, got: " + str(L) + " ft.")

    table = get_table(table_id)
    data = table["data"]

    # Available lengths are JSON keys (strings). Convert to ints, sort.
    available_lengths = sorted(int(k) for k in data.keys())

    # IFGC Step 5: use the smallest table row >= actual length.
    selected = None
    for row_len in available_lengths:
        if row_len >= L:
            selected = row_len
            break

    if selected is None:
        max_len = available_lengths[-1]
        raise ValueError(
            "Pipe length " + str(L) + " ft exceeds the maximum row of " +
            str(max_len) + " ft in Table " + table_id + ". "
            "Engineering calculations per IFGC A104 are required."
        )

    return selected, list(data[str(selected)])


# ---------------------------------------------------------------------------
# Capacity lookup
# ---------------------------------------------------------------------------

def get_capacity(table_id, length_ft, pipe_size_nominal):
    """Return the CFH capacity of a pipe size at a given run length.

    Performs IFGC A103.1 Step 5 length rounding internally.

    Args:
        table_id: Table identifier string.
        length_ft: Actual pipe length (typically the longest run), in feet.
        pipe_size_nominal: Nominal pipe size string matching the JSON,
            e.g. "1/2", "3/4", "1", "1-1/4", "2-1/2", "EHD-25", "1 K&L".

    Returns:
        Integer capacity in CFH.

    Raises:
        ValueError: If table_id is unknown, length is invalid, length
            exceeds the table maximum, pipe size is not in the table,
            or the table entry for this size/length is null (size not
            available at that length).
    """
    table = get_table(table_id)
    sizes = table["pipe_sizes_nominal"]

    if pipe_size_nominal not in sizes:
        raise ValueError(
            "Pipe size '" + str(pipe_size_nominal) + "' is not in Table " +
            table_id + ". Available sizes: " + ", ".join(sizes) + "."
        )

    _table_length_used, row = get_length_row(table_id, length_ft)
    idx = sizes.index(pipe_size_nominal)
    capacity = row[idx]

    if capacity is None:
        raise ValueError(
            "Pipe size '" + str(pipe_size_nominal) + "' has no published "
            "capacity at length " + str(length_ft) + " ft in Table " +
            table_id + " (table cell is null)."
        )

    return capacity


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

def _self_test():
    """Run a non-Revit self-test against the IFGC Appendix A Example."""
    print("gas_tables.py self-test")
    print("JSON path: " + _JSON_PATH)

    table_id = select_table(1.5, "Schedule 40 Steel")
    print("select_table(1.5, 'Schedule 40 Steel') -> " + table_id)
    assert table_id == "402.4(2)", "Expected 402.4(2), got " + table_id

    table_id_4 = select_table(2.0, "Schedule 40 Steel")
    print("select_table(2.0, 'Schedule 40 Steel') -> " + table_id_4)
    assert table_id_4 == "402.4(4)", "Expected 402.4(4), got " + table_id_4

    sizes = list_pipe_sizes("402.4(2)")
    print("list_pipe_sizes('402.4(2)') first 5 -> " + ", ".join(sizes[:5]))

    used_len, _row = get_length_row("402.4(2)", 60)
    print("get_length_row('402.4(2)', 60) used row -> " + str(used_len) + " ft")
    assert used_len == 60, "Expected 60 ft row"

    used_len_2, _row2 = get_length_row("402.4(2)", 63.41)
    print("get_length_row('402.4(2)', 63.41) used row -> " + str(used_len_2) + " ft")
    assert used_len_2 == 70, "Expected 70 ft row (Step 5 rounding)"

    cap_1in_60ft = get_capacity("402.4(2)", 60, "1")
    print("get_capacity('402.4(2)', 60, '1') -> " + str(cap_1in_60ft) + " CFH")
    assert cap_1in_60ft == 257, "Expected 257 CFH"

    # CSST table check
    cap_csst = get_capacity("402.4(15)", 10, "EHD-13")
    print("get_capacity('402.4(15)', 10, 'EHD-13') -> " + str(cap_csst) + " CFH")
    assert cap_csst == 32, "Expected 32 CFH"

    # Copper table check
    cap_cu = get_capacity("402.4(8)", 10, "1/2 K&L")
    print("get_capacity('402.4(8)', 10, '1/2 K&L') -> " + str(cap_cu) + " CFH")
    assert cap_cu == 20, "Expected 20 CFH"

    # TABLE_OPTIONS check
    opts = get_table_option_labels()
    print("TABLE_OPTIONS count: " + str(len(opts)))
    assert len(opts) == 37, "Expected 37 table options"

    opt = get_table_option_by_label(
        "Sched 40 Steel - Natural, <2 psi, 0.5\" w.c.  [402.4(2)]")
    assert opt["table_id"] == "402.4(2)", "Label lookup failed"
    print("Label lookup -> " + opt["table_id"])

    print("All assertions passed.")


if __name__ == "__main__":
    _self_test()
