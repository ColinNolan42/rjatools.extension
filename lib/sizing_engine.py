# -*- coding: ascii -*-
# sizing_engine.py
# IFGC Longest Run Method gas pipe sizing engine.
# Pure Python - no Revit API calls. All table lookups via gas_tables.py.
#
# IronPython 2.7

import gas_tables
import shared_params


# ---------------------------------------------------------------------------
# Nominal size -> decimal inches mapping
# Used when writing sizes back to Revit (Size Gas script converts to feet).
# Keys must match pipe_sizes_nominal strings in ifgc_gas_sizing_tables.json.
# ---------------------------------------------------------------------------

MIN_PIPE_INCHES = 0.75  # 3/4" firm minimum regardless of table result

NOMINAL_TO_INCHES = {
    # Schedule 40 Steel and PE Plastic Pipe (standard nominal sizes)
    "1/2":   0.5,
    "3/4":   0.75,
    "1":     1.0,
    "1-1/4": 1.25,
    "1-1/2": 1.5,
    "2":     2.0,
    "2-1/2": 2.5,
    "3":     3.0,
    "4":     4.0,
    "5":     5.0,
    "6":     6.0,
    "8":     8.0,
    "10":    10.0,
    "12":    12.0,
    # Semirigid Copper Tubing (K&L nominal sizes per IFGC Table 402.4(8)-(14))
    "3/8 K&L":   0.375,
    "1/2 K&L":   0.5,
    "5/8 K&L":   0.625,
    "3/4 K&L":   0.75,
    "1 K&L":     1.0,
    "1-1/4 K&L": 1.25,
    "1-1/2 K&L": 1.5,
    "1-5/8 ACR": 1.625,
    "2 K&L":     2.0,
    # CSST (approximate nominal equivalent per EHD designation)
    "EHD-13": 0.25,
    "EHD-15": 0.375,
    "EHD-18": 0.5,
    "EHD-19": 0.5,
    "EHD-23": 0.75,
    "EHD-25": 0.75,
    "EHD-30": 1.0,
    "EHD-31": 1.0,
    "EHD-37": 1.25,
    "EHD-39": 1.25,
    "EHD-46": 1.5,
    "EHD-48": 1.5,
    "EHD-60": 2.0,
    "EHD-62": 2.0,
}


# ---------------------------------------------------------------------------
# Heat Content of Gas -> CFH conversion (RJA standard)
# ---------------------------------------------------------------------------

def mbh_to_cfh(demand_mbh, heat_content_btu_per_cf):
    """Convert an MBH gas load to the CFH figure compared against IFGC
    Table 402.4 capacities.

    RJA standard (2026-09-23): CFH = BTUH / Heat Content of Gas. Heat
    Content of Gas is project/location-specific (get it from the utility),
    not a fixed altitude derate - this is the ONLY adjustment applied to
    gas pipe sizing. Matches RJA's real "Low Pressure Gas Size.xls" calc
    template. At sea level (1000 BTU/CF), CFH = MBH exactly.

    Args:
        demand_mbh:               float  cumulative downstream load, MBH.
        heat_content_btu_per_cf:  float  project/location Heat Content of
                                  Gas, BTU per actual cubic foot.

    Returns:
        float  CFH.
    """
    if heat_content_btu_per_cf <= 0:
        return float("inf")
    return (demand_mbh * 1000.0) / heat_content_btu_per_cf


# ---------------------------------------------------------------------------
# Low/High Pressure pipe capacity formulas (RJA standard, 2026-09-23)
#
# Ported directly from RJA's real calc templates ("Template Low Pressure Gas
# Size.xls" and "Template - High Pressure Gas Size.xlsx",
# FOR CLAUDE\Calc Tools\Plumbing\Gas Piping) via their live Excel formulas
# (read through COM, not the flattened values), then numerically verified
# to reproduce those workbooks' exact computed capacities. This is a
# DIFFERENT methodology from the discrete IFGC Appendix A capacity tables
# in ifgc_gas_sizing_tables.json (which are fixed at 0.60 specific gravity
# with no adjustment mechanism) - these formulas are genuinely Specific
# Gravity - sensitive and are RJA's own standard for both low and high
# pressure sizing, superseding the discrete table lookup.
# ---------------------------------------------------------------------------

LOW_PRESSURE_K_BINS = [
    # (K constant, max diameter this bin covers, in.) - from the Low
    # Pressure workbook's "Table" sheet Q1:T6 (its own named categories:
    # "3/4"-1"", "1-1/2"", "2"", "3"", "4""). No category exists above 4" -
    # per Colin, a low-pressure system needing more than 4" should be
    # stepped up to a higher design pressure (High Pressure/Weymouth-Cox
    # method) instead, not extrapolated.
    (1000.0, 1.0),
    (1100.0, 1.5),
    (1200.0, 2.0),
    (1300.0, 3.0),
    (1400.0, 4.0),
]

LOW_PRESSURE_MAX_DIAMETER_IN = 4.0


def _low_pressure_required_diameter(demand_cfh, k, specific_gravity,
                                     length_ft, pressure_loss_inwc):
    """Inverse of the Low Pressure workbook's capacity formula - the
    diameter (in.) that would exactly carry demand_cfh at the given K."""
    length_factor = length_ft / 3.59
    return ((demand_cfh / k) ** 2 * specific_gravity * length_factor
             / pressure_loss_inwc) ** 0.2


def low_pressure_required_diameter(demand_cfh, specific_gravity, length_ft,
                                    pressure_loss_inwc):
    """Required pipe diameter (in.) for demand_cfh under RJA's Low Pressure
    formula, resolving the workbook's manual "guess a size category, check
    for self-consistency" step automatically instead of requiring a human
    guess.

    The workbook's capacity formula (Capacity = K * sqrt(D^5 * PressureLoss
    / (SpecificGravity * (Length_ft/3.59)))) uses one K constant, chosen by
    which of 5 named pipe-size categories the answer is expected to fall
    in (LOW_PRESSURE_K_BINS) - the same K then applies to every diameter
    checked against it. Rather than guess, this tries each category's own
    K, computes what diameter would be required at that K, and returns the
    one that is self-consistent (falls within that same category's own
    diameter range) - the same condition a person manually iterating the
    spreadsheet would be checking for by eye.

    Raises:
        ValueError: if no category is self-consistent within 4" - per
                    Colin (2026-09-23), a real design would step up to a
                    higher pressure class rather than use a pipe this
                    large at low pressure; this is not extrapolated.
    """
    prev_max = 0.0
    for k, max_diameter in LOW_PRESSURE_K_BINS:
        d = _low_pressure_required_diameter(
            demand_cfh, k, specific_gravity, length_ft, pressure_loss_inwc)
        if prev_max < d <= max_diameter or (prev_max == 0.0 and d <= max_diameter):
            return d
        prev_max = max_diameter
    raise ValueError(
        "Demand {:.1f} CFH requires more than {:.0f}\" pipe under the Low "
        "Pressure method - step up to a higher design pressure (High "
        "Pressure / Weymouth-Cox sizing) instead.".format(
            demand_cfh, LOW_PRESSURE_MAX_DIAMETER_IN))


def low_pressure_capacity_cfh(diameter_in, specific_gravity, length_ft,
                               pressure_loss_inwc):
    """Capacity (CFH) of a given pipe diameter under RJA's Low Pressure
    formula, using the K constant for whichever named category
    diameter_in itself falls into (LOW_PRESSURE_K_BINS)."""
    k = None
    for bin_k, max_diameter in LOW_PRESSURE_K_BINS:
        if diameter_in <= max_diameter + 1e-9:
            k = bin_k
            break
    if k is None:
        raise ValueError(
            "{:.3f}\" exceeds the Low Pressure method's {:.0f}\" scope."
            .format(diameter_in, LOW_PRESSURE_MAX_DIAMETER_IN))
    length_factor = length_ft / 3.59
    return k * ((diameter_in ** 5) * pressure_loss_inwc
                / (specific_gravity * length_factor)) ** 0.5


def weymouth_capacity_cfh(diameter_in, inlet_absolute_psi, outlet_absolute_psi,
                           specific_gravity, length_ft, atm_pressure_psi):
    """High Pressure capacity (CFH) via the Weymouth Formula, for pipe 3"
    and larger. Ported verbatim from the High Pressure workbook:
    Q = 18.062 x (T/P) x sqrt((P1^2 - P2^2) x D^(16/3) / (G x T1 x L))
    where T = 520 (absolute temp, deg R, fixed), T1 = T+60 = 580 (flowing
    temp), P = atm_pressure_psi (base pressure), P1/P2 = inlet/outlet
    ABSOLUTE pressure (gauge + atm_pressure_psi), L = length in miles.
    """
    T = 520.0
    T1 = T + 60.0
    length_miles = length_ft / 5280.0
    return (18.062 * (T / atm_pressure_psi)
            * ((inlet_absolute_psi ** 2 - outlet_absolute_psi ** 2)
               * diameter_in ** (16.0 / 3.0)
               / (specific_gravity * T1 * length_miles)) ** 0.5)


def cox_capacity_cfh(diameter_in, inlet_absolute_psi, outlet_absolute_psi,
                      specific_gravity, length_ft):
    """High Pressure capacity (CFH) via the Cox Formula, for pipe under 3".
    Ported verbatim from the High Pressure workbook:
    Q = 33.3 x sqrt((P1^2 - P2^2) x D^5 / (G x L)), L = length in miles.
    """
    length_miles = length_ft / 5280.0
    return 33.3 * ((inlet_absolute_psi ** 2 - outlet_absolute_psi ** 2)
                    * (diameter_in ** 5) / (specific_gravity * length_miles)) ** 0.5


def high_pressure_capacity_cfh(diameter_in, inlet_absolute_psi,
                                outlet_absolute_psi, specific_gravity,
                                length_ft, atm_pressure_psi):
    """Dispatches to the Weymouth Formula (>=3") or Cox Formula (<3"),
    matching the High Pressure workbook's own two-section split."""
    if diameter_in >= 3.0:
        return weymouth_capacity_cfh(
            diameter_in, inlet_absolute_psi, outlet_absolute_psi,
            specific_gravity, length_ft, atm_pressure_psi)
    return cox_capacity_cfh(
        diameter_in, inlet_absolute_psi, outlet_absolute_psi,
        specific_gravity, length_ft)


# ---------------------------------------------------------------------------
# Main sizing function
# ---------------------------------------------------------------------------

def size_system(graph, pipe_material, inlet_pressure_psi, table_id=None,
                 heat_content_btu_per_cf=None):
    """Size every pipe segment using the IFGC Longest Run Method.

    Per IFGC A103.1:
      - One longest run length is used for ALL segments.
      - Each segment is sized for its cumulative downstream demand (CFH).
      - Smallest nominal size whose table capacity >= demand is selected.
      - CFH = BTUH / Heat Content of Gas (see mbh_to_cfh()) - the sole
        adjustment applied when converting MBH demand to CFH for table
        lookup. Per Colin (2026-09-23): this Heat Content conversion is
        the only piece of the real Low/High Pressure Excel calc templates
        that belongs in this tool - pipe capacity/sizing itself stays on
        the discrete IFGC Table 402.4 lookup (gas_tables.get_capacity()),
        not the templates' own K-constant/Weymouth/Cox capacity formulas
        (see low_pressure_capacity_cfh()/high_pressure_capacity_cfh()
        above - verified working, kept for possible future use, but not
        called from here).

    Args:
        graph:               NetworkGraph from pipe_graph.build_network()
        pipe_material:       str  e.g. "Schedule 40 Steel"
        inlet_pressure_psi:  float  supply pressure at the meter (used if
                             table_id is None)
        table_id:            str  optional - IFGC table ID to use directly,
                             e.g. "402.4(2)". When supplied, inlet_pressure_psi
                             is stored in the result but not used for lookup.
        heat_content_btu_per_cf: float  project/location Heat Content of
                             Gas, BTU per actual cubic foot. Defaults to
                             shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF
                             (Denver) if not supplied.

    Returns:
        dict with keys:
            sizes                 {pipe_element_id (int): nominal_size (str)}
            table_id              str  e.g. "402.4(2)"
            table_length_used_ft  float
            longest_run_ft        float
            pipe_material         str
            inlet_pressure_psi    float
            heat_content_btu_per_cf  float
            segment_detail        list of dicts - one per sized segment

    Raises:
        ValueError: If longest run is missing, table not available, or any
                    pipe demand exceeds the maximum table capacity.
    """
    if heat_content_btu_per_cf is None:
        heat_content_btu_per_cf = shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF

    if graph.longest_run is None:
        raise ValueError(
            "Longest run not found. Run Diagnose and verify the system "
            "has at least one reachable fixture before sizing.")

    longest_run_ft = graph.longest_run["total_length_feet"]
    if longest_run_ft <= 0:
        raise ValueError(
            "Longest run is 0 ft. Verify meter connection and traversal.")

    if table_id is None:
        table_id = gas_tables.select_table(inlet_pressure_psi, pipe_material)
    else:
        gas_tables.get_table(table_id)  # validate table exists
    pipe_sizes = gas_tables.list_pipe_sizes(table_id)
    table_length_used, _ = gas_tables.get_length_row(table_id, longest_run_ft)

    sizes = {}
    segment_detail = []
    sizing_errors = []

    def _apply_minimum(nom, pipe_sizes_list):
        """If nom is smaller than MIN_PIPE_INCHES, return the first size in
        pipe_sizes_list whose decimal equivalent meets the minimum."""
        if NOMINAL_TO_INCHES.get(nom, 0.0) >= MIN_PIPE_INCHES:
            return nom, False
        for s in pipe_sizes_list:
            if NOMINAL_TO_INCHES.get(s, 0.0) >= MIN_PIPE_INCHES:
                return s, True
        return nom, False  # fallback: no larger size available

    for edge in graph.edges.values():
        # Skip open-ended pipes - no downstream node to size for
        if edge.to_node_id is None:
            continue

        demand_mbh = edge.cumulative_load_mbh

        # CFH = BTUH / Heat Content of Gas (see mbh_to_cfh() docstring).
        demand_cfh_effective = mbh_to_cfh(demand_mbh, heat_content_btu_per_cf)

        # Zero demand: assign minimum available pipe size
        if demand_mbh <= 0:
            selected = pipe_sizes[0]
            selected, upsized = _apply_minimum(selected, pipe_sizes)
            capacity_at_size = gas_tables.get_capacity(
                table_id, longest_run_ft, selected)
            note = "zero demand - minimum size assigned"
            if upsized:
                note += " (upsized to 3/4\" firm minimum)"
            segment_detail.append({
                "pipe_id":             edge.element_id,
                "demand_mbh":          0.0,
                "demand_cfh":          0.0,
                "selected_size":       selected,
                "capacity_mbh":        capacity_at_size,
                "note":                note
            })
            sizes[edge.element_id] = selected
            continue

        # Find smallest size whose capacity >= demand
        selected = None
        selected_capacity = None
        for size in pipe_sizes:
            try:
                capacity = gas_tables.get_capacity(
                    table_id, longest_run_ft, size)
                if capacity >= demand_cfh_effective:
                    selected = size
                    selected_capacity = capacity
                    break
            except ValueError:
                # Size not available at this length - try the next size up
                continue

        if selected is None:
            sizing_errors.append(
                "Pipe {}: demand {:.1f} MBH ({:.1f} CFH) exceeds max table "
                "capacity at {:.0f} ft in Table {}.".format(
                    edge.element_id, demand_mbh, demand_cfh_effective,
                    table_length_used, table_id))
            continue

        selected, upsized = _apply_minimum(selected, pipe_sizes)
        if upsized:
            selected_capacity = gas_tables.get_capacity(
                table_id, longest_run_ft, selected)

        sizes[edge.element_id] = selected
        segment_detail.append({
            "pipe_id":             edge.element_id,
            "demand_mbh":          round(demand_mbh, 1),
            "demand_cfh":          round(demand_cfh_effective, 1),
            "selected_size":       selected,
            "capacity_mbh":        selected_capacity,
            "note":                "upsized to 3/4\" firm minimum" if upsized else ""
        })

    if sizing_errors:
        raise ValueError(
            "Sizing failed for {} pipe(s):\n{}".format(
                len(sizing_errors), "\n".join(sizing_errors)))

    return {
        "sizes":                  sizes,
        "table_id":               table_id,
        "table_length_used_ft":   table_length_used,
        "longest_run_ft":         longest_run_ft,
        "pipe_material":          pipe_material,
        "inlet_pressure_psi":     inlet_pressure_psi,
        "heat_content_btu_per_cf": heat_content_btu_per_cf,
        "segment_detail":         segment_detail,
    }


# ---------------------------------------------------------------------------
# Downstream fixture helper
# ---------------------------------------------------------------------------

def _downstream_fixtures(start_node_id, graph):
    """Return list of fixture names reachable downstream from start_node_id."""
    names = []
    visited = set()
    stack = [start_node_id]

    while stack:
        nid = stack.pop()
        if nid is None or nid in visited:
            continue
        visited.add(nid)

        node = graph.nodes.get(nid)
        if node is None:
            continue

        if node.is_gas_fixture:
            names.append(node.fixture_name or "UNNAMED")
            continue

        for edge in graph.edges.values():
            if edge.from_node_id == nid and edge.to_node_id is not None:
                stack.append(edge.to_node_id)

        for child_id in graph.node_children.get(nid, []):
            stack.append(child_id)

    return names


# ---------------------------------------------------------------------------
# Diagnostic formatter
# ---------------------------------------------------------------------------

def format_sizing_output(sizing_result, graph):
    """Return a formatted string summarizing sizing results for the output window.

    Args:
        sizing_result: dict returned by size_system()
        graph:         NetworkGraph used for sizing

    Returns:
        str
    """
    lines = []

    # Total load and fixture count
    fixture_nodes = [n for n in graph.nodes.values() if n.is_gas_fixture]
    total_mbh = sum(n.gas_load_mbh for n in fixture_nodes)

    lines.append("=== SIZING RESULTS ===")
    lines.append("Table:         {}".format(sizing_result["table_id"]))
    lines.append("Table row:     {} ft  (longest run {:.1f} ft rounded up)".format(
        sizing_result["table_length_used_ft"],
        sizing_result["longest_run_ft"]))
    lines.append("Pipe material: {}".format(sizing_result["pipe_material"]))
    lines.append("Inlet PSI:     {}".format(sizing_result["inlet_pressure_psi"]))
    lines.append("Heat Content:  {:.0f} BTU/CF".format(
        sizing_result.get("heat_content_btu_per_cf",
                           shared_params.DEFAULT_HEAT_CONTENT_BTU_PER_CF)))
    lines.append("Total load:    {:.1f} MBH  |  {} fixtures".format(
        total_mbh, len(fixture_nodes)))
    lines.append("")

    sizes = sizing_result["sizes"]
    lines.append("=== PIPE SIZES ({} segments) ===".format(len(sizes)))

    for detail in sizing_result["segment_detail"]:
        edge = graph.edges.get(detail["pipe_id"])
        length_str = "{:.1f}'".format(edge.length_feet) if edge else "?"

        # Downstream fixture label
        fixture_label = ""
        if edge and edge.to_node_id is not None:
            fixtures = _downstream_fixtures(edge.to_node_id, graph)
            if len(fixtures) == 1:
                fixture_label = "  -> {}".format(fixtures[0])
            elif len(fixtures) > 1:
                fixture_label = "  -> trunk ({}: {})".format(
                    len(fixtures), ", ".join(fixtures[:3]) +
                    (" ..." if len(fixtures) > 3 else ""))

        note = "  ({})".format(detail["note"]) if detail["note"] else ""
        cfh_note = ""
        if abs(detail.get("demand_cfh", detail["demand_mbh"]) - detail["demand_mbh"]) > 0.05:
            cfh_note = "  [{:.1f} CFH]".format(detail.get("demand_cfh", 0.0))
        lines.append(
            "  [{}]  {:.1f} MBH{}  ->  {}\"  "
            "(cap {} MBH)  {}{}{}".format(
                detail["pipe_id"],
                detail["demand_mbh"],
                cfh_note,
                detail["selected_size"],
                int(detail["capacity_mbh"]),
                length_str,
                fixture_label,
                note))

    return "\n".join(lines)


def format_stub_report(skipped_stubs):
    """Return a formatted section listing fixture stub pipes that were skipped.

    Args:
        skipped_stubs: list of dicts with keys:
            pipe_id, fixture_name, demand_mbh, recommended_size

    Returns:
        str
    """
    if not skipped_stubs:
        return ""

    lines = ["=== FIXTURE STUB PIPES (not written - manually resize these) ==="]
    for s in skipped_stubs:
        lines.append(
            "  [{}]  {}  {:.1f} MBH  ->  recommend {}\"".format(
                s["pipe_id"],
                s["fixture_name"],
                s["demand_mbh"],
                s["recommended_size"]))
    return "\n".join(lines)
