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

import re

import shared_params
import water_tables
import water_graph


def build_report(graph, sizing_result, header, recirc=None):
    """Return the whole report as a single string.

    Args:
        graph: the traversed WaterGraph.
        sizing_result: what water_sizing_engine.size_network returned.
        header: dict with "date", "job", "job_number", "by".
        recirc: what water_recirc.analyze returned, or None when the system
            has no hot water return.
    """
    lines = []
    lines.extend(basis_block(header))
    for line in origin_basis_lines(graph):
        lines.append("  " + line)
    lines.append("")
    lines.extend(wsfu_table(graph))
    lines.append("")
    lines.extend(segment_table(graph, sizing_result))
    lines.append("")
    if recirc is not None and (recirc["legs"] or recirc["unrooted_pipe_ids"]):
        lines.extend(recirc_table(recirc))
        lines.append("")
    lines.extend(exceptions_block(graph, sizing_result))
    return "\n".join(lines)


def build_compact_report(graph, sizing_result, header, recirc=None):
    """The window report: results, the equations with their numbers, and what
    needs a look. No pipe by pipe listing, no basis paragraphs (the full pipe
    table and the basis of design go on the drafting view and in
    build_report). Colin, 2026-10-26.
    """
    lines = ["=" * 78,
             "SIZE WATER   JOB {} {}   {}   BY {}".format(
                 header.get("job_number", ""), header.get("job", ""),
                 header.get("date", ""), header.get("by", "")),
             "=" * 78,
             "Basis: 2024 IPC Appendix E, Table E103.3(2) WSFU, "
             "Table E103.3(3) demand, firm WSFU per size table (Copper L, "
             "8 fps cold, 5 fps hot)."]
    for line in origin_basis_lines(graph):
        lines.append(line)
    lines.append("")
    lines.extend(wsfu_table(graph))
    lines.append("")
    lines.extend(size_summary(sizing_result))
    lines.append("")
    if recirc is not None and (recirc["legs"] or recirc["unrooted_pipe_ids"]):
        lines.extend(recirc_compact(recirc))
        lines.append("")
    lines.extend(exceptions_compact(graph, sizing_result))
    return "\n".join(lines)


def size_summary(sizing_result):
    """Pipe counts by nominal size and system, in place of the pipe table."""
    order = [row["nominal_size"] for row in water_tables.firm_sizing_rows()]
    counts = {}
    total_ft = 0.0
    held = 0
    for s in sizing_result["segments"]:
        if s.keep_existing:
            held += 1
            continue
        if s.nominal_size is None:
            continue
        key = "HWR" if s.is_return else _system_label(s.system)
        row = counts.setdefault(s.nominal_size, {"CW": 0, "HW": 0, "HWR": 0})
        row[key] = row.get(key, 0) + 1
        total_ft += float(s.length_feet or 0.0)
    lines = ["PIPES SIZED (count by size)", "-" * 78,
             "{:<8}{:>8}{:>8}{:>8}".format("SIZE", "CW", "HW", "HWR")]
    totals = {"CW": 0, "HW": 0, "HWR": 0}
    for nominal in order:
        row = counts.get(nominal)
        if row is None:
            continue
        lines.append("{:<8}{:>8}{:>8}{:>8}".format(
            nominal, row["CW"], row["HW"], row["HWR"]))
        for key in totals:
            totals[key] += row[key]
    lines.append("{:<8}{:>8}{:>8}{:>8}   {:.1f} ft".format(
        "TOTAL", totals["CW"], totals["HW"], totals["HWR"], total_ft))
    if held:
        lines.append("{} existing phase pipe(s) left at their drawn size."
                     .format(held))
    return lines


def recirc_compact(recirc):
    """Circuit setter flows as the equation with its numbers."""
    heat_loss = recirc["heat_loss_btu_per_100ft"]
    drop = recirc["delta_t_f"]
    legs = recirc["legs"]
    lines = ["HOT WATER RECIRCULATION", "-" * 78,
             "gpm = (L / 100 x {} BTU/h) / (500 x {} F), {:.1f} gpm minimum "
             "per circuit setter".format(_num(heat_loss), _num(drop),
                                         float(recirc["min_gpm"]))]
    if legs:
        leg = legs[0]
        lines.append(
            "Leg {}: ({:.1f} / 100 x {}) / (500 x {}) = {:.2f} gpm, set "
            "{:.2f}".format(leg["number"], float(leg["effective_length_ft"]),
                            _num(heat_loss), _num(drop),
                            float(leg["calc_gpm"]), float(leg["gpm"])))
    branches = [l for l in legs if l["method"] != "LOOP"]
    shown = [l for l in legs if l["method"] == "LOOP"]
    if len(branches) <= 6:
        shown = legs
        branches = []
    lines.append("{:<5}{:<9}{:>10}{:>10}{:>9}".format(
        "NO.", "METHOD", "L FT", "CALC", "SET GPM"))
    for leg in shown:
        lines.append("{:<5}{:<9}{:>10.1f}{:>10.2f}{:>9.2f}".format(
            leg["number"], leg["method"],
            float(leg["effective_length_ft"]), float(leg["calc_gpm"]),
            float(leg["gpm"])))
    if branches:
        lengths = [float(l["effective_length_ft"]) for l in branches]
        sets = sorted(set("{:.2f}".format(float(l["gpm"])) for l in branches))
        lines.append("{} BRANCH legs, L {:.1f} to {:.1f} ft, set {} gpm".format(
            len(branches), min(lengths), max(lengths), ", ".join(sets)))
    lines.append("TOTAL CIRCULATION {:.2f} gpm, {} circuit setter(s), "
                 "{:.1f} ft of return".format(
                     float(recirc["total_gpm"]), len(legs),
                     float(recirc.get("return_length_ft", 0.0))))
    if recirc["unrooted_pipe_ids"]:
        lines.append("{} return pipe(s) NOT SIZED, the end of their return "
                     "line was not found: {}".format(
                         len(recirc["unrooted_pipe_ids"]),
                         _id_list(recirc["unrooted_pipe_ids"])))
    return lines


_FLAG_PIPE = re.compile(r"^Pipe (\d+): (.*)$")


def group_flags(flags):
    """One line per distinct message, with a count and the first ids."""
    order = []
    ids = {}
    plain = []
    for flag in flags:
        match = _FLAG_PIPE.match(flag)
        if match is None:
            if flag not in plain:
                plain.append(flag)
            continue
        message = match.group(2)
        if message not in ids:
            ids[message] = []
            order.append(message)
        ids[message].append(int(match.group(1)))
    lines = []
    for message in order:
        lines.append("{} pipe(s): {} ({})".format(
            len(ids[message]), message, _id_list(ids[message])))
    lines.extend(plain)
    return lines


def _id_list(ids, show=5):
    ids = list(ids)
    text = ", ".join(str(i) for i in ids[:show])
    if len(ids) > show:
        text += ", +{} more".format(len(ids) - show)
    return text


def exceptions_compact(graph, sizing_result):
    lines = ["FLAGS", "-" * 78]
    if not graph.heater_ids:
        lines.append("NO WATER HEATER found on the cold network. Hot water "
                     "demand is NOT included in any cold pipe.")
    else:
        reported = set()
        groups = getattr(graph, "heater_groups", {}) or {}
        for hid in graph.heater_ids:
            group = groups.get(hid, (hid,))
            if len(group) > 1:
                if group in reported:
                    continue
                reported.add(group)
                fixtures = set()
                for member in group:
                    fixtures.update(graph.nodes[member].served_fixture_ids)
                lines.append(
                    "Heaters {} in parallel serve {} fixture(s), each carries "
                    "1/{} of the hot load.".format(
                        ", ".join(str(m) for m in group), len(fixtures),
                        len(group)))
            else:
                lines.append("Heater {} serves {} fixture(s).".format(
                    hid, len(graph.nodes[hid].served_fixture_ids)))

    lines.extend(group_flags(sizing_result["flags"]))

    if graph.open_ends:
        lines.append("{} open connector(s), the network may be incomplete: "
                     "{}".format(len(graph.open_ends),
                                 _id_list([e for e, _s in graph.open_ends])))
    if graph.existing_connections:
        lines.append("{} connection(s) to EXISTING piping not followed: {}"
                     .format(len(graph.existing_connections),
                             _id_list([far for _n, far in
                                       graph.existing_connections])))
    passthrough = getattr(graph, "existing_passthrough", [])
    if passthrough:
        lines.append("{} EXISTING device(s) inside the new run, walked "
                     "through: {}".format(len(passthrough),
                                          _id_list(passthrough)))
    if getattr(graph, "parallel_backflow", None):
        lines.append("Backflow preventer(s) {} are in parallel with the "
                     "picked one. The leg behind each is sized for the full "
                     "demand.".format(
                         _id_list(graph.parallel_backflow)))
    if graph.upstream_ids:
        lines.append("{} element(s) behind the RPZ inlet are outside the "
                     "scope and were not sized.".format(
                         len(graph.upstream_ids)))
    if graph.loops:
        lines.append("{} loop joint(s), expected on a recirculating system, "
                     "otherwise a modeling error.".format(len(graph.loops)))
    if graph.unreadable:
        lines.append("{} element(s) had no readable connectors.".format(
            len(graph.unreadable)))
    zero = [f for f in graph.fixture_ids
            if graph.nodes[f].total_wsfu <= 0 and
            not float(graph.nodes[f].fixed_gpm or 0.0)]
    if zero:
        lines.append("{} fixture(s) carry zero WSFU, check Type and "
                     "public/private: {}".format(len(zero), _id_list(zero)))
    if len(graph.system_types) > 1:
        lines.append("System Types: " + ", ".join(
            "{} ({})".format(graph.system_types[t], t)
            for t in sorted(graph.system_types.keys())))
    if sizing_result.get("flush_valve_system"):
        lines.append("Flush valve system: cold water sized on gpm, flush "
                     "valve column of Table E103.3(3).")
    if len(lines) == 2:
        lines.append("None.")
    lines.append("=" * 78)
    return lines


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
    lines.append("")
    lines.extend(sizing_chart_lines())
    return lines


def sizing_chart_lines():
    """The firm WSFU-to-pipe-size chart as fixed-width console lines."""
    chart = water_tables.sizing_chart()
    fmt = "  {:<11}{:<20}{:<19}{:<14}"
    lines = ["FIRM PIPE SIZING CHART", fmt.format(*chart["headers"])]
    for row in chart["rows"]:
        lines.append(fmt.format(*row))
    return lines


def origin_basis_lines(graph):
    """Basis lines that depend on where the run started.

    Shared by the printed report and the drafting view, so the sheet says the
    same thing the window does. Empty for a run that started at the RPZ.
    """
    if not graph.origin_is_new_pipe:
        return []
    return [
        "Existing building: sizing starts at the new construction "
        "connection ({} {}), not at the backflow preventer.".format(
            "the branch of tee" if getattr(graph, "origin_is_tee", False)
            else "pipe", graph.origin_id),
        "The existing cold water main upstream of that connection is "
        "assumed adequate for the added load and was NOT sized or checked.",
    ]


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
    # The extended columns are quantity x per fixture, so they are named as
    # totals. Plain COLD / HOT / TOTAL twice read as the same numbers (Colin).
    ("cold_ext",  "TOTAL COLD",      0.100,   11),
    ("hot_ext",   "TOTAL HOT",       0.095,   10),
    ("total_ext", "TOTAL WSFU",      0.100,   11),
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
    # The extended columns SUM each fixture's own values. They are deliberately
    # not count x (the first member's value): if the Type name ever fails to
    # read, every fixture collapses into one group, and multiplying the first
    # member's load by the count then silently reports a wrong total. That
    # happened live on 2026-09-24, where 12 water closets and 6 lavatories
    # printed as 18 x 5.0 = 90 cold and 0 hot instead of 69 cold and 9 hot.
    groups = {}
    fixed_loads = {}      # fixture name -> [count, gpm each]
    for fid in graph.fixture_ids:
        node = graph.nodes.get(fid)
        if node is None:
            continue
        if float(node.fixed_gpm or 0.0) > 0.0:
            entry = fixed_loads.setdefault(
                node.fixture_name or node.type_name,
                [0, float(node.fixed_gpm)])
            entry[0] += 1
            continue
        occupancy = "PUBLIC" if node.is_public else "PRIVATE"
        key = (node.type_name, occupancy)
        if key not in groups:
            groups[key] = {"count": 0, "cold": 0.0, "hot": 0.0, "total": 0.0,
                           "per_fixture": set()}
        g = groups[key]
        g["count"] += 1
        g["cold"] += node.cw_wsfu
        g["hot"] += node.hw_wsfu
        g["total"] += node.total_wsfu
        g["per_fixture"].add((node.cw_wsfu, node.hw_wsfu, node.total_wsfu))

    rows = []
    ext_cold = 0.0
    ext_hot = 0.0
    ext_total = 0.0

    for key in sorted(groups.keys()):
        type_name, occupancy = key
        g = groups[key]
        ext_cold += g["cold"]
        ext_hot += g["hot"]
        ext_total += g["total"]

        # The per-fixture columns only mean something when every fixture in the
        # row carries the same load. If they differ, say so rather than print
        # one member's value as if it applied to all of them.
        if len(g["per_fixture"]) == 1:
            each = list(g["per_fixture"])[0]
            cold_ea, hot_ea, total_ea = (_num(each[0]), _num(each[1]),
                                         _num(each[2]))
        else:
            cold_ea = hot_ea = total_ea = "varies"

        rows.append({
            "fixture": "{}, {}".format(type_name, occupancy),
            "qty": str(g["count"]),
            "cold_ea": cold_ea,
            "hot_ea": hot_ea,
            "total_ea": total_ea,
            "cold_ext": _num(g["cold"]),
            "hot_ext": _num(g["hot"]),
            "total_ext": _num(g["total"]),
        })

    ext_fixed_gpm = 0.0
    for name in sorted(fixed_loads.keys()):
        count, each = fixed_loads[name]
        ext_fixed_gpm += count * each
        rows.append({
            "fixture": "{}, {} gpm each (Table 604.3)".format(
                name, _num(each)),
            "qty": str(count),
            "cold_ea": "-", "hot_ea": "-", "total_ea": "-",
            "cold_ext": "-", "hot_ext": "-", "total_ext": "-",
        })

    totals = {
        "fixture": "TOTALS",
        "qty": str(sum(g["count"] for g in groups.values()) +
                   sum(v[0] for v in fixed_loads.values())),
        "cold_ea": "", "hot_ea": "", "total_ea": "",
        "cold_ext": _num(ext_cold),
        "hot_ext": _num(ext_hot),
        "total_ext": _num(ext_total),
    }

    return {"rows": rows, "totals": totals, "ext_cold": ext_cold,
            "ext_hot": ext_hot, "ext_total": ext_total,
            "ext_fixed_gpm": ext_fixed_gpm,
            "fixed_loads": fixed_loads,
            "flush_valve_system": water_graph.has_flush_valves(graph)}


def wsfu_load_lines(data):
    """The sizing-load lines printed under the take-off table.

    These tie the take-off to the sizing: the TOTAL here is the load the
    building main carries in the segment table.

    The footnote is not decoration. COLD + HOT does not equal TOTAL, and every
    reader checks that sum. Note a to 2024 IPC Table E103.3(2) sets the
    separate hot and cold loads at three-fourths of the total for any fixture
    served by both, as a diversity allowance, so the columns are deliberately
    not additive.

    The gpm follows the Table E103.3(3) column the sizing used, and says
    which (Colin, 2026-10-21): flushometer valve for cold water on a system
    serving a flush valve fixture, flush tank otherwise and for hot water.
    """
    cold = water_tables.demand_gpm(data["ext_total"],
                                   data.get("flush_valve_system", False))
    hot = water_tables.wsfu_to_gpm(data["ext_hot"])
    if cold["curve"] == water_tables.CURVE_FLUSHOMETER:
        basis = ("WSFU to gpm: 2024 IPC Table E103.3(3). Cold water uses the "
                 "flushometer valve column (the system serves a flush valve "
                 "fixture), hot water the flush tank column.")
    elif data.get("flush_valve_system"):
        basis = ("WSFU to gpm: 2024 IPC Table E103.3(3), flush tank column. "
                 "The system serves a flush valve fixture, but the cold load "
                 "is below the first row of the flushometer valve column.")
    else:
        basis = ("WSFU to gpm: 2024 IPC Table E103.3(3), flush tank column "
                 "(no flush valve fixtures on the system).")
    fixed = float(data.get("ext_fixed_gpm", 0.0) or 0.0)
    if fixed > 0.0 and cold.get("gpm") is not None:
        cold_line = ("Cold water design load = {} total WSFU = {} gpm, plus "
                     "{} gpm fixed = {:.1f} gpm".format(
                         _num(data["ext_total"]), _gpm(cold), _num(fixed),
                         float(cold["gpm"]) + fixed))
    elif fixed > 0.0:
        cold_line = ("Cold water design load = {} total WSFU, plus {} gpm "
                     "fixed".format(_num(data["ext_total"]), _num(fixed)))
    else:
        cold_line = "Cold water design load = {} total WSFU = {} gpm".format(
            _num(data["ext_total"]), _gpm(cold))
    return [
        cold_line,
        "Hot water design load = {} hot WSFU = {} gpm".format(
            _num(data["ext_hot"]), _gpm(hot)),
        basis,
        "Note: COLD + HOT does not equal TOTAL. Per note a to IPC Table "
        "E103.3(2), a fixture served by both hot and cold water has its "
        "separate hot and cold loads set at three-fourths of its total load.",
    ]


def wsfu_table(graph):
    """Console rendering of the take-off table."""
    data = wsfu_data(graph)

    fmt = "".join(["{:<%d}" % c[3] if i == 0 else "{:>%d}" % c[3]
                   for i, c in enumerate(WSFU_COLUMNS)])
    lines = [WSFU_TITLE,
             "-" * 86,
             "{:<38}{:>20}{:>28}".format(
                 "", "--- PER FIXTURE ---", "--- QTY x PER FIXTURE ---"),
             fmt.format(*[c[1] for c in WSFU_COLUMNS]),
             "-" * 86]

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
        if s.is_return:
            # Return piping carries circulation flow, not fixture units.
            lines.append("{:<10}{:<7}{:>8}{:>8}{:>8}{:>9}{:>9}  {}".format(
                s.element_id,
                "HWR",
                "-",
                "-" if s.circ_gpm is None else "{:.2f}".format(
                    float(s.circ_gpm)),
                "{:.1f}".format(float(s.length_feet)),
                _size(s.existing_size_in),
                s.nominal_size if s.nominal_size else "-",
                s.rule))
            continue
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
# 4. Hot water recirculation: the circuit setter settings
# ---------------------------------------------------------------------------

RECIRC_TITLE = "HOT WATER RECIRCULATION SIZING"

RECIRC_COLUMNS = [
    # key          header             view_w
    ("number",    "NO.",              0.045),
    ("method",    "METHOD",           0.075),
    ("element",   "AT ELEMENT",       0.095),
    ("length",    "EFF LENGTH FT",    0.105),
    ("btuh",      "HEAT LOSS BTU/H",  0.120),
    ("design_dt", "DESIGN DROP F",    0.105),
    ("calc_gpm",  "CALC GPM",         0.080),
    ("gpm",       "SET GPM",          0.075),
    ("actual_dt", "ACTUAL DROP F",    0.105),
]


def recirc_data(recirc):
    """Rows for the recirculation sizing table, keyed by RECIRC_COLUMNS.

    ACTUAL DROP is the temperature drop at the flow the circuit setter is set
    to: heat loss / (500 x set gpm). It equals the design drop where the
    calculated flow governs, and is smaller where the minimum flow governs.

    Returns dict with "rows", "summary" (lines printed under the table) and
    "basis" (lines printed above it).
    """
    design_dt = float(recirc["delta_t_f"])
    rows = []
    for leg in recirc["legs"]:
        rows.append({
            "number": str(leg["number"]),
            "method": leg["method"],
            "element": ("-" if leg["start_element_id"] is None
                        else str(leg["start_element_id"])),
            "length": "{:.1f}".format(float(leg["effective_length_ft"])),
            "btuh": "{:.0f}".format(float(leg.get("heat_loss_btuh", 0.0))),
            "design_dt": _num(design_dt),
            "calc_gpm": "{:.2f}".format(float(leg["calc_gpm"])),
            "gpm": "{:.2f}".format(float(leg["gpm"])),
            "actual_dt": "{:.1f}".format(float(leg.get("actual_delta_t_f", 0.0))),
        })

    basis = [
        "Heat loss {} BTU/h per 100 ft, design temperature drop {} F, {} gpm "
        "minimum per circuit setter.".format(
            _num(recirc["heat_loss_btu_per_100ft"]), _num(design_dt),
            "{:.1f}".format(float(recirc["min_gpm"]))),
        "gpm = (effective length / 100 x heat loss) / (500 x temperature "
        "drop). Actual drop = heat loss / (500 x set gpm).",
        "LOOP length is the return from the last fixture back to the water "
        "heater, or to where a second loop joins the main return. BRANCH "
        "length is the hot supply from the main to the fixture.",
    ]
    summary = [
        "Total hot water return piping: {:.1f} ft.".format(
            float(recirc.get("return_length_ft", 0.0))),
        "Total circulation flow, {} circuit setter(s): {:.2f} gpm.".format(
            len(recirc["legs"]), float(recirc["total_gpm"])),
    ]
    return {"rows": rows, "summary": summary, "basis": basis}


def recirc_table(recirc):
    """One row per circuit setter, with the flow to set it to."""
    lines = ["HOT WATER RECIRCULATION - CIRCUIT SETTER FLOWS",
             "-" * 78,
             "  Heat loss {} BTU/h per 100 ft, temperature drop {} F, "
             "{} gpm minimum".format(
                 _num(recirc["heat_loss_btu_per_100ft"]),
                 _num(recirc["delta_t_f"]),
                 "{:.1f}".format(float(recirc["min_gpm"]))),
             "  per circuit setter. One circuit setter is assumed on each "
             "return leg.",
             "  LOOP length is the return from the last fixture back to the "
             "water heater.",
             "  A second loop that joins the main return is measured to where "
             "it joins.",
             "  BRANCH length is the hot supply from the main to the fixture.",
             "-" * 78,
             "{:<5}{:<9}{:>12}{:>11}{:>10}{:>10}{:>9}{:>12}".format(
                 "NO.", "METHOD", "AT ELEMENT", "EFF LEN FT", "BTU/H",
                 "CALC GPM", "SET GPM", "ACT DROP F"),
             "-" * 78]

    for leg in recirc["legs"]:
        lines.append("{:<5}{:<9}{:>12}{:>11}{:>10}{:>10}{:>9}{:>12}".format(
            leg["number"],
            leg["method"],
            "-" if leg["start_element_id"] is None else leg["start_element_id"],
            "{:.1f}".format(float(leg["effective_length_ft"])),
            "{:.0f}".format(float(leg.get("heat_loss_btuh", 0.0))),
            "{:.2f}".format(float(leg["calc_gpm"])),
            "{:.2f}".format(float(leg["gpm"])),
            "{:.1f}".format(float(leg.get("actual_delta_t_f", 0.0)))))
    if not recirc["legs"]:
        lines.append("  (no return leg could be laid out)")

    lines.append("-" * 78)
    for root_id in sorted(recirc["root_gpm"].keys()):
        lines.append("  Circulation flow returning at element {}: {:.2f} "
                     "gpm".format(root_id, float(recirc["root_gpm"][root_id])))
    lines.append("  TOTAL CIRCULATION FLOW, {} circuit setter(s): {:.2f} "
                 "gpm".format(len(recirc["legs"]), float(recirc["total_gpm"])))
    lines.append("  TOTAL HOT WATER RETURN PIPING: {:.1f} ft".format(
        float(recirc.get("return_length_ft", 0.0))))
    if recirc["unrooted_pipe_ids"]:
        lines.append("  {} return pipe(s) NOT SIZED, the end of their return "
                     "line was not found: {}".format(
                         len(recirc["unrooted_pipe_ids"]),
                         ", ".join(str(p) for p in
                                   recirc["unrooted_pipe_ids"][:10])))
    lines.append("-" * 78)
    return lines


# ---------------------------------------------------------------------------
# 5. Everything the run could not do
# ---------------------------------------------------------------------------

def exceptions_block(graph, sizing_result):
    lines = ["FLAGS, SKIPPED AND NOT EVALUATED", "-" * 78]
    any_flag = False

    if not graph.heater_ids:
        lines.append("  NO WATER HEATER found on the cold network. Hot water "
                     "demand is therefore NOT included in any cold pipe.")
        any_flag = True
    else:
        reported = set()
        groups = getattr(graph, "heater_groups", {}) or {}
        for hid in graph.heater_ids:
            group = groups.get(hid, (hid,))
            if len(group) > 1:
                if group in reported:
                    continue
                reported.add(group)
                fixtures = set()
                for member in group:
                    fixtures.update(graph.nodes[member].served_fixture_ids)
                lines.append(
                    "  Water heaters {} are in parallel: they serve {} "
                    "fixture(s) together and split the hot load equally, so "
                    "each cold feed and hot outlet carries 1/{} of it, "
                    "combined at the tee where the outlets join.".format(
                        ", ".join(str(m) for m in group), len(fixtures),
                        len(group)))
                continue
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

    if graph.existing_connections:
        lines.append("  {} connection(s) to EXISTING phase piping were not "
                     "followed. The existing piping is assumed adequate and "
                     "is not sized: {}".format(
                         len(graph.existing_connections),
                         ", ".join(str(far) for _near, far in
                                   graph.existing_connections[:10])))

    passthrough = getattr(graph, "existing_passthrough", [])
    if passthrough:
        lines.append("  {} EXISTING phase valve(s) or fitting(s) sit inside "
                     "the new construction run, with new work on every other "
                     "side, so the run was walked through them. Check their "
                     "phase: {}".format(
                         len(passthrough),
                         ", ".join(str(e) for e in passthrough[:10])))

    if graph.loops:
        lines.append("  {} loop connection(s) detected. A recirculation loop "
                     "is expected here; any other loop is a modeling "
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
        lines.append("      Hot supply and hot water recirculation share one "
                     "system CLASSIFICATION, so the RETURN is identified by "
                     "System Type instead. A system identified as the return "
                     "is sized on circulation flow, not fixture units.")

    segments = sizing_result["segments"]
    cw_count = len([s for s in segments
                    if s.system == shared_params.SYSTEM_DOMESTIC_COLD_WATER
                    and s.nominal_size is not None])
    hw_count = len([s for s in segments
                    if s.system == shared_params.SYSTEM_DOMESTIC_HOT_WATER
                    and not s.is_return and s.nominal_size is not None])
    hwr_count = len([s for s in segments
                     if s.is_return and s.nominal_size is not None])
    hwr_unsized = len([s for s in segments
                       if s.is_return and s.nominal_size is None])
    lines.append("  SIZED BY THIS RUN: {} cold water pipe(s), {} hot water "
                 "pipe(s) and {} hot water return pipe(s).".format(
                     cw_count, hw_count, hwr_count))
    if hwr_unsized:
        lines.append("  NOT SIZED: {} hot water return pipe(s), see the "
                     "flags above.".format(hwr_unsized))
    if sizing_result.get("flush_valve_system"):
        lines.append("  DEMAND: the system serves a flush valve fixture, so "
                     "cold water uses the flush valve column of Table "
                     "E103.3(3) and is sized on gpm.")
    lines.append("  VELOCITY AND PRESSURE LOSS: checked by Water "
                 "Performance.")

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
