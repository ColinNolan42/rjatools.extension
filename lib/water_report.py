# -*- coding: ascii -*-
# water_report.py
# Formats the Size Water output for the pyRevit window.
#
# Three blocks, in this order:
#   1. Basis of design (from water_tables, so it cannot drift from the tables
#      actually used)
#   2. The WSFU take-off table, laid out like the firm's
#      Template_WSFU CALCS.xlsx sheet WFUTO
#   3. The segment sizing table, plus everything that was flagged or skipped
#
# Returns plain strings. No Revit imports, no pyRevit imports, so the layout
# can be checked without opening Revit.
#
# IronPython 2.7

import shared_params
import water_tables
import water_graph


def build_report(graph, sizing_result, header):
    """Return the whole report as a single string.

    Args:
        graph: the traversed WaterGraph.
        sizing_result: what water_sizing_engine.size_network returned.
        header: dict with "date", "job", "job_number", "by".
    """
    lines = []
    lines.extend(basis_block(header))
    lines.append("")
    lines.extend(wsfu_table(graph))
    lines.append("")
    lines.extend(segment_table(graph, sizing_result))
    lines.append("")
    lines.extend(exceptions_block(graph, sizing_result))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 1. Basis of design
# ---------------------------------------------------------------------------

def basis_block(header):
    lines = ["=" * 78,
             "SIZE WATER - DOMESTIC WATER PIPE SIZING",
             "=" * 78,
             "DATE:     {}".format(header.get("date", "")),
             "JOB:      {}".format(header.get("job", "")),
             "JOB NO.:  {}".format(header.get("job_number", "")),
             "BY:       {}".format(header.get("by", "")),
             "",
             "BASIS OF DESIGN"]
    for line in water_tables.basis_of_design_lines():
        lines.append("  " + line)
    return lines


# ---------------------------------------------------------------------------
# 2. WSFU take-off table (WFUTO layout)
# ---------------------------------------------------------------------------

# ONE definition drives BOTH renderings of this table: the console text below
# and the drafting-view schedule in water_drafting.py. They are drawn by
# completely different code (padded monospace vs. detail lines and TextNotes),
# and the HVAC tool's equivalent tables drifted apart when each renderer owned
# its own column list, so the columns, headers and cell values all come from
# here.
#   view_w     column width in feet at the drafting view's 1:1 scale
#              (= printed inches / 12)
#   console_w  column width in monospace characters
WSFU_COLUMNS = [
    # key          header           view_w  console_w
    ("fixture",   "FIXTURE TYPE",    0.230,   34),
    ("qty",       "QTY",             0.055,    4),
    ("cold_ea",   "COLD",            0.070,    7),
    ("hot_ea",    "HOT",             0.070,    6),
    ("total_ea",  "TOTAL",           0.070,    7),
    ("cold_ext",  "COLD",            0.070,    7),
    ("hot_ext",   "HOT",             0.070,    6),
    ("total_ext", "TOTAL",           0.070,    7),
]

WSFU_TITLE = "WATER SUPPLY FIXTURE UNIT TAKE-OFF"


def wsfu_data(graph):
    """Group the traversed fixtures into WFUTO rows.

    Mirrors Template_WSFU CALCS.xlsx sheet WFUTO: the fixed 28-row fixture list
    is replaced by the fixture Types actually found in the model. Public and
    private are separate rows because they carry different WSFU.

    Returns dict with "rows" (list of cell dicts keyed by WSFU_COLUMNS keys),
    "totals" (a cell dict), and the raw "ext_cold"/"ext_hot"/"ext_total"
    numbers the sizing loads are read from.
    """
    groups = {}
    for fid in graph.fixture_ids:
        node = graph.nodes.get(fid)
        if node is None:
            continue
        occupancy = "PUBLIC" if node.is_public else "PRIVATE"
        key = (node.type_name, occupancy)
        if key not in groups:
            groups[key] = {"count": 0, "cold": node.cw_wsfu,
                           "hot": node.hw_wsfu, "total": node.total_wsfu}
        groups[key]["count"] += 1

    rows = []
    ext_cold = 0.0
    ext_hot = 0.0
    ext_total = 0.0

    for key in sorted(groups.keys()):
        type_name, occupancy = key
        g = groups[key]
        row_cold = g["count"] * g["cold"]
        row_hot = g["count"] * g["hot"]
        row_total = g["count"] * g["total"]
        ext_cold += row_cold
        ext_hot += row_hot
        ext_total += row_total
        rows.append({
            "fixture": "{}, {}".format(type_name, occupancy),
            "qty": str(g["count"]),
            "cold_ea": _num(g["cold"]),
            "hot_ea": _num(g["hot"]),
            "total_ea": _num(g["total"]),
            "cold_ext": _num(row_cold),
            "hot_ext": _num(row_hot),
            "total_ext": _num(row_total),
        })

    totals = {
        "fixture": "TOTALS",
        "qty": str(sum(g["count"] for g in groups.values())),
        "cold_ea": "", "hot_ea": "", "total_ea": "",
        "cold_ext": _num(ext_cold),
        "hot_ext": _num(ext_hot),
        "total_ext": _num(ext_total),
    }

    return {"rows": rows, "totals": totals, "ext_cold": ext_cold,
            "ext_hot": ext_hot, "ext_total": ext_total}


def wsfu_load_lines(data):
    """The two sizing-load lines printed under the take-off table.

    These tie the take-off to the sizing: the TOTAL here is the load the
    building main carries in the segment table.
    """
    return [
        "Cold water sizing load (TOTAL wsfu): {}  ->  {} gpm".format(
            _num(data["ext_total"]),
            _gpm(water_tables.wsfu_to_gpm(data["ext_total"]))),
        "Hot water sizing load (HOT wsfu): {}  ->  {} gpm".format(
            _num(data["ext_hot"]),
            _gpm(water_tables.wsfu_to_gpm(data["ext_hot"]))),
    ]


def wsfu_table(graph):
    """Console rendering of the take-off table."""
    data = wsfu_data(graph)

    fmt = "{:<34}{:>4}{:>7}{:>6}{:>7}{:>7}{:>6}{:>7}"
    lines = [WSFU_TITLE,
             "-" * 78,
             fmt.format(*[c[1] for c in WSFU_COLUMNS]),
             "{:<34}{:>4}{:>20}{:>20}".format(
                 "", "", "--- PER FIXTURE ---", "---- EXTENDED ----"),
             "-" * 78]

    if not data["rows"]:
        lines.append("  (no water fixtures found on the traversed network)")
        lines.append("-" * 78)
        return lines

    for row in data["rows"]:
        cells = [row[c[0]] for c in WSFU_COLUMNS]
        cells[0] = _clip(cells[0], 34)
        lines.append(fmt.format(*cells))

    lines.append("-" * 78)
    lines.append(fmt.format(*[data["totals"][c[0]] for c in WSFU_COLUMNS]))
    lines.append("-" * 78)
    lines.extend(wsfu_load_lines(data))
    return lines


# ---------------------------------------------------------------------------
# 3. Segment sizing table
# ---------------------------------------------------------------------------

def segment_table(graph, sizing_result):
    lines = ["PIPE SIZING BY SEGMENT",
             "-" * 78,
             "{:<10}{:<7}{:>8}{:>8}{:>8}{:>9}{:>9}  {}".format(
                 "PIPE ID", "SYS", "WSFU", "GPM", "LEN FT",
                 "EXISTING", "NEW", "RULE"),
             "-" * 78]

    segments = sizing_result["segments"]
    if not segments:
        lines.append("  (no pipes were traversed)")
        lines.append("-" * 78)
        return lines

    for s in segments:
        lines.append("{:<10}{:<7}{:>8}{:>8}{:>8}{:>9}{:>9}  {}".format(
            s.element_id,
            _system_label(s.system),
            _num(s.demand_wsfu),
            _gpm({"gpm": s.demand_gpm}),
            "{:.1f}".format(float(s.length_feet)),
            _size(s.existing_size_in),
            s.nominal_size if s.nominal_size else "-",
            s.rule))

    totals = sizing_result["totals"]
    lines.append("-" * 78)
    lines.append("{} pipes, {} sized, {} not sized, {:.1f} ft total".format(
        totals["pipe_count"], totals["sized_count"],
        totals["unsized_count"], float(totals["total_length_ft"])))
    lines.append("-" * 78)
    return lines


# ---------------------------------------------------------------------------
# 4. Everything the run could not do
# ---------------------------------------------------------------------------

def exceptions_block(graph, sizing_result):
    lines = ["FLAGS, SKIPPED AND NOT EVALUATED", "-" * 78]
    any_flag = False

    if not graph.heater_ids:
        lines.append("  NO WATER HEATER found on the cold network. Hot water "
                     "demand is therefore NOT included in any cold pipe.")
        any_flag = True
    else:
        for hid in graph.heater_ids:
            node = graph.nodes[hid]
            lines.append("  Water heater {} serves {} fixture(s).".format(
                hid, len(node.served_fixture_ids)))

    for flag in sizing_result["flags"]:
        lines.append("  " + flag)
        any_flag = True

    if graph.open_ends:
        lines.append("  {} open connector(s) - the network may be "
                     "incomplete:".format(len(graph.open_ends)))
        for element_id, system in graph.open_ends[:15]:
            lines.append("      element {} ({})".format(
                element_id, _system_label(system)))
        if len(graph.open_ends) > 15:
            lines.append("      ... and {} more".format(
                len(graph.open_ends) - 15))
        any_flag = True

    if graph.loops:
        lines.append("  {} loop connection(s) detected. A recirculation loop "
                     "is expected here; any other loop is a modelling "
                     "error.".format(len(graph.loops)))
        any_flag = True

    if graph.unreadable:
        lines.append("  {} element(s) had no readable connectors.".format(
            len(graph.unreadable)))
        any_flag = True

    fixtures_without_load = [
        fid for fid in graph.fixture_ids
        if graph.nodes[fid].total_wsfu <= 0]
    if fixtures_without_load:
        lines.append("  {} fixture(s) carry zero WSFU - check the Type and "
                     "the public/private setting: {}".format(
                         len(fixtures_without_load),
                         ", ".join(str(f) for f in fixtures_without_load[:10])))
        any_flag = True

    if len(graph.system_types) > 1:
        lines.append("  Piping system types found on the network:")
        for type_id in sorted(graph.system_types.keys()):
            lines.append("      {} (id {})".format(
                graph.system_types[type_id], type_id))
        lines.append("      A hot water RETURN type also classifies as "
                     "Domestic Hot Water, so it cannot be told apart "
                     "automatically. Any type ticked as a return in the "
                     "dialog was reported but not sized.")

    # Stated positively first, because "hot water recirculation is not sized"
    # on its own reads as "hot water is not sized", which is wrong.
    cw_count = len([s for s in sizing_result["segments"]
                    if s.system == shared_params.SYSTEM_DOMESTIC_COLD_WATER
                    and s.nominal_size is not None])
    hw_count = len([s for s in sizing_result["segments"]
                    if s.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER
                    and s.nominal_size is not None])
    lines.append("  SIZED BY THIS RUN: {} cold water pipe(s) and {} hot water "
                 "pipe(s).".format(cw_count, hw_count))
    lines.append("  NOT SIZED: the hot water RETURN line only, which is sized "
                 "on circulation flow rather than fixture units.")
    lines.append("  NOT EVALUATED: pressure loss and velocity.")

    if not any_flag:
        lines.insert(2, "  No blocking issues found.")
    lines.append("=" * 78)
    return lines


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _system_label(system):
    if system == shared_params.SYSTEM_DOMESTIC_COLD_WATER:
        return "CW"
    if system == shared_params.SYSTEM_DOMESTIC_HOT_WATER:
        return "HW"
    return "?"


def _num(value):
    """Whole numbers print without a decimal tail.

    float() is required before '{:.0f}' - IronPython 2.7 raises ValueError when
    formatting an int with that spec.
    """
    number = float(value)
    if abs(number - round(number)) < 0.005:
        return "{:.0f}".format(number)
    return "{:.2f}".format(number)


def _gpm(demand):
    if demand is None or demand.get("gpm") is None:
        return "-"
    return "{:.1f}".format(float(demand["gpm"]))


def _size(inches):
    if not inches:
        return "-"
    return "{:.2f}".format(float(inches))


def _clip(text, width):
    if len(text) <= width:
        return text
    return text[:width - 1] + "."
