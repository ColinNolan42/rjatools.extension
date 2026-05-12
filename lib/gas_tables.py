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

Supported tables (Schedule 40 Steel only at this time):
    402.4(2) - Less than 2 PSI, 0.5 in. w.c. drop
    402.4(5) - 2.0 PSI inlet, 1.0 PSI drop
    402.4(6) - 3.0 PSI inlet, 2.0 PSI drop
    402.4(7) - 5.0 PSI inlet, 3.5 PSI drop

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
    """Load the JSON file once and cache it at module level.

    Raises:
        IOError: If the JSON file cannot be found at the expected path.
        ValueError: If the JSON file is malformed.
    """
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

SUPPORTED_MATERIALS = ("Schedule 40 Steel",)


def _validate_material(pipe_material):
    """Raise ValueError if material is not supported.

    Other materials (Type L Copper, CSST, etc.) will be added when their
    sizing tables are added to the JSON.
    """
    if pipe_material not in SUPPORTED_MATERIALS:
        raise ValueError(
            "Unsupported pipe material: '" + str(pipe_material) + "'. "
            "Supported materials: " + ", ".join(SUPPORTED_MATERIALS) + "."
        )


# ---------------------------------------------------------------------------
# Table selection (handoff Task 2.1)
# ---------------------------------------------------------------------------

def select_table(inlet_pressure_psi, pipe_material="Schedule 40 Steel"):
    """Select the appropriate IFGC sizing table based on inlet pressure.

    Selection logic per project handoff (Schedule 40 Steel):
        inlet_pressure < 2.0    -> 402.4(2)   0.5 in. w.c. drop
        inlet_pressure == 2.0   -> 402.4(5)   1.0 PSI drop
        inlet_pressure == 3.0   -> 402.4(6)   2.0 PSI drop
        inlet_pressure == 5.0   -> 402.4(7)   3.5 PSI drop

    Pressures between the discrete higher values (e.g. 2.5 PSI, 4.0 PSI)
    are not covered by IFGC tables and raise ValueError. The user shall
    select a system pressure that matches a published table.

    Args:
        inlet_pressure_psi: Gas supply pressure at the meter, in PSI.
        pipe_material: Pipe material string. Currently only
            "Schedule 40 Steel" is supported.

    Returns:
        Table ID string, e.g. "402.4(2)".

    Raises:
        ValueError: If pipe material is unsupported, inlet pressure is
            invalid, or no published table matches the inlet pressure.
    """
    _validate_material(pipe_material)

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

    if p < 2.0:
        table_id = "402.4(2)"
    elif p == 2.0:
        table_id = "402.4(5)"
    elif p == 3.0:
        table_id = "402.4(6)"
    elif p == 5.0:
        table_id = "402.4(7)"
    else:
        raise ValueError(
            "No IFGC table available for inlet pressure " + str(p) + " PSI. "
            "Supported pressures for Schedule 40 Steel: <2.0, 2.0, 3.0, 5.0 PSI."
        )

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
# Capacity lookup (handoff Task 2.1)
# ---------------------------------------------------------------------------

def get_capacity(table_id, length_ft, pipe_size_nominal):
    """Return the CFH capacity of a pipe size at a given run length.

    Performs IFGC A103.1 Step 5 length rounding internally: the caller
    passes the raw longest run length and this function selects the
    next-longer table row.

    Args:
        table_id: Table identifier string.
        length_ft: Actual pipe length (typically the longest run), in feet.
        pipe_size_nominal: Nominal pipe size string matching the JSON,
            e.g. "1/2", "3/4", "1", "1-1/4", "2-1/2".

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
# Self-test (run this file standalone outside Revit to verify the loader)
# ---------------------------------------------------------------------------

def _self_test():
    """Run a non-Revit self-test against the IFGC Appendix A Example 1.

    Example 1 conditions:
        Table 402.4(2), longest run 60 ft.
        Section 3 (all four appliances) demand = 245 CFH.
        Expected size: 1 inch (capacity 257 CFH at 60 ft).

    This is the same example validated by the reference
    gas_sizing_engine.py file.
    """
    print("gas_tables.py self-test")
    print("JSON path: " + _JSON_PATH)

    table_id = select_table(1.5, "Schedule 40 Steel")
    print("select_table(1.5, 'Schedule 40 Steel') -> " + table_id)
    assert table_id == "402.4(2)", "Expected 402.4(2)"

    table_id_2 = select_table(2.0)
    print("select_table(2.0) -> " + table_id_2)
    assert table_id_2 == "402.4(5)", "Expected 402.4(5)"

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
    assert cap_1in_60ft == 257, "Expected 257 CFH (IFGC Example 1)"

    cap_3q_60ft = get_capacity("402.4(2)", 60, "3/4")
    print("get_capacity('402.4(2)', 60, '3/4') -> " + str(cap_3q_60ft) + " CFH")
    assert cap_3q_60ft == 137, "Expected 137 CFH"

    # Step 5 rounding: 63.41 ft should use the 70 ft row.
    cap_1in_63 = get_capacity("402.4(2)", 63.41, "1")
    print("get_capacity('402.4(2)', 63.41, '1') -> " + str(cap_1in_63) + " CFH")
    assert cap_1in_63 == 237, "Expected 237 CFH (70 ft row, 1 inch)"

    # Error cases
    try:
        select_table(2.5)
    except ValueError as e:
        print("select_table(2.5) correctly raised: " + str(e))

    try:
        get_capacity("402.4(2)", 60, "1-3/4")
    except ValueError as e:
        print("Bad pipe size correctly raised: " + str(e))

    try:
        get_length_row("402.4(2)", 5000)
    except ValueError as e:
        print("Excessive length correctly raised: " + str(e))

    print("All assertions passed.")


if __name__ == "__main__":
    _self_test()
