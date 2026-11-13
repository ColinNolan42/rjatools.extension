# -*- coding: utf-8 -*-
"""
duct_perf_run.py  --  Duct Performance engine (judging, external static,
fittings, labels, sheet and level plan drawing).

Moved verbatim out of Duct Performance.pushbutton/script.py (2026-11-12) so an
agent can import it through the pyRevit agent runtime, the same way the other
tools' pipelines live in lib. The button keeps its dialogs, picker and main()
and imports from here. No behavior change.

The engine reads the module globals `doc` and `output`, as it did inside the
button script. Call bind(doc, output) once before using it; the button does.

IronPython 2.7 / pyRevit  --  no f-strings, no walrus, no nonlocal.
"""

import math
import re

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter, ViewSheet, ViewType,
    ViewDuplicateOption, ElementId, XYZ, ViewPlan, BoundingBoxXYZ,
    Color, TextNote, TextNoteOptions,
    ConnectorProfileType, ConnectorType
)
from pyrevit import script

import hvac_graph
import fitting_tables
import ashrae_fittings
import fan_system_effect
import revit_helpers
import rja_drafting
import vav_checks
import duct_perf_console
from revit_helpers import eid_int

doc = None
output = None


def bind(document, out):
    """Point the engine at a document and an output window (pyRevit
    output, or anything with print_md / print_html / print_code)."""
    global doc, output
    doc = document
    output = out

# ── SMACNA colors ─────────────────────────────────────────────────────────────
GREEN  = Color(0,   200,  0)
YELLOW = Color(255, 215,  0)
RED    = Color(210,  40, 40)
PURPLE = Color(140,  40, 200)
GRAY   = Color(160, 160, 160)

_COLOR_MAP = {'GREEN': GREEN, 'YELLOW': YELLOW, 'RED': RED, 'PURPLE': PURPLE, 'GRAY': GRAY}


# ── output table columns ───────────────────────────────────────────────────────
# ONE definition list drives BOTH output tables: the drafting-view schedule
# placed on the sheet and the console table. They are rendered by different
# code (detail lines + TextNotes vs. an HTML table) and have
# drifted apart before, so the column set, the headers and the cell values all
# come from here and from _row_cells() rather than being written out twice.
#
#   view_w     column width in feet at the schedule view's 1:1 scale
#              (= printed inches / 12)
#   console_w  column width in monospace characters (plain text console
#              fallback, used only when output.print_html fails)
#   default_on starting checkbox state in the settings dialog (selectable
#              columns only - see _FIXED_COLUMN_KEYS)
#
# Some columns are always in the tables and have no checkbox; the rest are
# optional. Defaults for the optional ones are the firm's chosen starting
# state, not a stored preference.
_COLUMN_DEFS = [
    # key          header                              view_w  console_w  default_on
    ('num',       '#',                                  0.050,    3,  True),
    ('status',    'Status',                             0.120,    7,  True),
    ('reason',    'Reason',                             0.385,   46,  True),
    ('role',      'Duct Role',                          0.150,   18,  True),
    ('size',      'Size',                               0.100,    9,  True),
    ('cfm',       'Total CFM through Duct',             0.200,   22,  False),
    ('fpm',       'Actual FPM',                         0.120,   11,  True),
    ('fric',      'Actual Fric (iwc/100)',              0.180,   21,  True),
    ('length',    'Length (ft)',                        0.090,   11,  False),
]

# Always shown, no checkbox. '#' is also the number printed in the keynote
# circles placed in the view, so it stays in the table to keep that
# correlation.
_FIXED_COLUMN_KEYS = frozenset(
    ['num', 'status', 'reason', 'role', 'size', 'fpm', 'fric'])
# No Required Size column (Colin, 2026-11-12): Duct Performance is a check
# tool. It says high or low against the standard; the engineer picks the size
# (space, noise). Size Duct still sizes.


def _optional_columns():
    """The _COLUMN_DEFS entries that get a checkbox, in definition order."""
    return [c for c in _COLUMN_DEFS if c[0] not in _FIXED_COLUMN_KEYS]


def _selected_columns(selected_keys):
    """The _COLUMN_DEFS entries to render: every fixed column plus the
    optional ones the user checked, in definition order.

    Both renderers call this so column ORDER is defined once too, not just
    the set - a dialog that returned an unordered set would otherwise let the
    two tables order themselves differently.
    """
    return [c for c in _COLUMN_DEFS
            if c[0] in _FIXED_COLUMN_KEYS or c[0] in selected_keys]


DUCT_REDUCTION_REASON = 'Reduced Duct Size Available'   # PURPLE reason (Colin, 2026-11-12; was 'Duct reduction size available', 'Oversized')
MEETS_STANDARD_REASON = 'Meets Design Standard'         # GREEN main (Colin, 2026-11-12)
# RED / YELLOW main reasons (Colin, 2026-11-12): a check, never a size.
_MAIN_REASONS = {
    ('RED', 'vel'):     'High Velocity',
    ('RED', 'fric'):    'High Pressure Loss',
    ('RED', 'both'):    'High Velocity and Pressure Loss',
    ('YELLOW', 'vel'):  'Velocity Above Max Standard',
    ('YELLOW', 'fric'): 'Pressure Loss Above Max Standard',
    ('YELLOW', 'both'): 'Velocity and Pressure Loss Above Max Standard',
}
MR_LINE_WEIGHT = 8   # projection line weight of the most restrictive run (1 to 16)
MR_LINE_COLOR = Color(255, 128, 0)   # its outline: orange, apart from every label color
_PRIORITY = {'RED': 4, 'YELLOW': 3, 'PURPLE': 2, 'GREEN': 1, 'GRAY': 0}


# Character budget for one line of the System Summary block. The block is drawn
# at the flagged-duct table's width, and TextNote.Create with no width
# constraint does NOT wrap, so a longer line runs off the right edge of the chart
# instead of folding. Sized just under the shortest lines already known to fit.
SUMMARY_WRAP_CHARS = 92


def _wrap_summary_lines(lines, width=SUMMARY_WRAP_CHARS):
    """Fold over-long summary lines at word boundaries, keeping their indent.

    Colin, 2026-09-24: the METHOD and NOT INCLUDED lines "are writing too far out
    and need to be condensed (entered another line) down so it fit within the
    chart."

    A continuation gets its parent's indent plus two, so a folded sentence still
    reads as one item rather than as a new entry. Applied to the whole list
    rather than just the two long lines, so this cannot silently regress the next
    time someone adds a wordy line.
    """
    out = []
    for line in lines:
        if len(line) <= width:
            out.append(line)
            continue
        indent = len(line) - len(line.lstrip(' '))
        cont   = ' ' * (indent + 2)
        cur    = ' ' * indent
        for word in line.split():
            add = word if not cur.strip() else ' ' + word
            if cur.strip() and len(cur) + len(add) > width:
                out.append(cur)
                cur = cont + word
            else:
                cur = cur + add
        if cur.strip():
            out.append(cur)
    return out


def _esp_text_line(item, inwc, detail=None):
    """One external static line of the drafting view summary.

    Item padded to 19 so every value starts in the same column, e.g.
    '      Duct friction      0.123 in. wc   (12 ducts, 140 ft)'.
    """
    line = '      {:<19}{:.3f} in. wc'.format(item, inwc)
    if detail:
        line += '   ({})'.format(detail)
    return line


def _esp_console_text(total_inwc, sides_text):
    """The highlighted external static total in the console."""
    return 'EXTERNAL STATIC PRESSURE = {:.3f} in. wc ({})'.format(
        float(total_inwc), sides_text)


def _emit_html(html, fallback):
    """Print one themed console section, or its plain text when print_html
    is missing or throws. The failure is logged, never silent."""
    try:
        output.print_html(html)
        return True
    except Exception as ex:
        script.get_logger().warning(
            'Duct Performance: output.print_html failed ({}: {}), plain text '
            'printed instead.'.format(type(ex).__name__, ex))
    try:
        fallback()
    except Exception as ex:
        script.get_logger().error(
            'Duct Performance: plain text fallback failed too ({}: {}).'.format(
                type(ex).__name__, ex))
    return False


def _emit_corrections(corrections, headers=(), rows=(), status_col=None,
                      align=None, widths=None, subtitle=None, rows_text=None):
    """Print the System Corrections section: the duct table (rows may be
    empty) and every gathered correction line, Action required then Notes.
    rows_text: the padded monospace duct table lines for the plain text
    fallback, or None."""
    rows = list(rows)

    def _text():
        output.print_md('### System Corrections')
        if rows_text:
            output.print_code('\n'.join(rows_text))
        for line in duct_perf_console.corrections_text(corrections, bool(rows)):
            output.print_md(line)

    _emit_html(duct_perf_console.system_corrections(
        list(headers), rows, corrections, status_col=status_col, align=align,
        widths=widths, subtitle=subtitle), _text)


def _root_class_airflow(root_id, all_children, all_terminals):
    """Total CFM and terminal count per system class, downstream of ONE root.

    Per equipment, deliberately: the external static report is grouped by unit,
    and an airflow total that silently spanned two RTUs would not match the
    static pressure printed beside it.

    Counted per REACHABLE TERMINAL, the same source the duct CFMs are built
    from (hvac_graph reads Flow only at OST_DuctTerminal leaves), so this total
    and the duct velocities cannot disagree. A terminal reachable from two roots
    is counted under both, which is correct: both fans move that air.

    Returns {sys_class: (total_cfm, terminal_count)}.
    """
    out = {}
    stack = [root_id]
    seen = set([root_id])
    while stack:
        nid = stack.pop()
        term = all_terminals.get(nid)
        if term is not None:
            cfm, sys_class, _family = term
            key = sys_class or 'Unknown'
            tot, cnt = out.get(key, (0.0, 0))
            out[key] = (tot + cfm, cnt + 1)
        for cid in all_children.get(nid, []):
            if cid not in seen:
                seen.add(cid)
                stack.append(cid)
    return out


def _critical_path_loss(all_root_ids, all_children, all_duct_results,
                        all_terminals, all_nodes, safety_pct=0.0,
                        c_values=None, comp_values=None,
                        duct_rooted_ids=None, return_filter_inwc=0.0,
                        vav_box_ids=None, vav_box_drop_inwc=None,
                        fan_effect=None, discharge_roots=None, exhaust_cap=False,
                        backdraft_roots=None, inline_fan_roots=None,
                        keep_runs=False):
    """Worst fan-to-terminal path per system class: the index run, with fittings.

    keep_runs (Ranked run, Colin 2026-11-10): every class leaf also carries
    'runs', each terminal's own worst run priced the same way, sorted by
    total high to low, ties by terminal id, so runs[0] is the index run.

    Total external static pressure is a PATH, not a sum. Air leaving the fan
    takes one route to one terminal and the fan has to beat the worst route;
    shorter runs get dampered back. Colin, 2026-09-24: "we do not need to
    account for every branch just the most restrictive run."

    Follows RJA's SP_LOSS_WORKSHEET method:
      duct friction   Darcy-Weisbach + Altshul-Tsal, per segment
      fittings        C * Pv, C from fitting_tables (1985 ASHRAE numbers)
      safety factor   a flat % on the subtotal, the sheet's own last row

    TWO coefficients apply at a take-off and they are never both charged to the
    same air, because they are two different streams through one junction:
      - the run TURNS OFF through a tap        -> 0.98  (supply dovetail branch)
      - the run CONTINUES past a tap on a main -> 0.20  (main duct @ take-off; worksheet 0.28, lowered 2026-10-09 per ASHRAE)
    Over a whole index run you therefore accumulate 0.20 for every OTHER tap
    hanging off the mains it traverses, plus one 0.98 where it finally leaves.
    Colin (2026-09, when it was .28): "for each branch you should add the .28 as it restricts the main duct
    which may be the most restrictive."

    Revit models a tap as a 2-connector in-line stub on an UNBROKEN main, so
    there is no element on the main representing that 0.20 - which is exactly
    why that column sits at 0 in the filled worksheet. The taps hanging off each
    main are counted instead (hvac_graph.takeoff_child_ids).

    Components on the run are included too: a balancing damper's drop for each
    one found, and the diffuser's at the end. Those two are COUNTED from the
    model but VALUED in the dialog, because a damper's drop depends on how far it
    is throttled and a diffuser's is a cutsheet figure - neither is derivable
    from duct geometry.

    NOT INCLUDED, and by definition rather than as a gap: filter, coil and
    cabinet losses. Those are inside the unit and already deducted from the
    manufacturer's published ESP, so counting them here would double them. The one
    exception is the return filter, added as a known drop when the engineer ticks
    "Include return filter" (return_filter_inwc), for a filter the published ESP
    does not already cover.

    duct_rooted_ids: roots that are a picked duct because no unit was found. For
    those, the duct and fittings between the picked duct and the unit connection
    are priced and added to every path (see _unit_side_chain).
    Duct Performance no longer offers a duct pick (Colin, 2026-11-03), so main()
    passes an empty set; the case is kept for the walker and its tests.
    Fire / smoke and backdraft dampers are priced at their OWN known drops, never
    the balancing damper figure; any other accessory on the run is reported as
    uncounted. exhaust_cap: every Exhaust Air index run not from a ceiling
    exhaust fan also pays the 'cap' drop once (an inline fan's roof cap).
    inline_fan_roots: inline exhaust fan roots (Colin, 2026-11-03). The
    connection reaching the grilles is the inlet, the other the discharge; the
    discharge's friction, fittings, accessories and outlet fan system effect,
    at the intake cfm, are added to every Exhaust Air run of the fan.
    backdraft_roots: exhaust fan roots whose Exhaust Air run pays one
    'backdraft' drop when no backdraft damper is modeled on it (Include
    backdraft damper; a ceiling fan's root is left out, its damper is built in).

    all_children / all_terminals / all_nodes are keyed by int element id.
    all_duct_results is keyed by the real Revit ElementId, so it is re-keyed by
    its own DuctResult.element_id here rather than assumed to match.

    Returns {root_id: {sys_class: {...}}}, i.e. keyed by EQUIPMENT FIRST.
    Each leaf carries friction_inwc, fitting_inwc, component_inwc,
    subtotal_inwc, total_inwc, duct_count, fitting_count, tap_bypass_count,
    length_ft, terminal_id, terminal_name, unpriced, truncated.

    Per equipment, NOT merged across equipment. Colin, 2026-09-24: "if i
    selected two equipments i need to total external pressure drops ... split per
    equipment." Every unit gets its own fan, so its own index run and its own
    external static; merging two AHUs would let the bigger one's run hide the
    smaller one's entirely.

    Mid tree equipment (VAV plan, Colin 2026-10-22): equipment the run passes
    THROUGH (anything but this walk's own root, so a box that is also a picked
    root is still priced in its AHU's walk) is a component. One in vav_box_ids
    is a VAV box, priced at vav_box_drop_inwc (default: the balancing damper
    value) under component key 'vav_box', and the velocity the next fitting
    sees is reset to the box's outlet duct. Any other mid tree equipment (VAV
    off, or a duct heater and the like) is charged under 'mid_equip' at the
    balancing damper value. Through any mid tree equipment the walk only
    continues on its own side (supply stays supply).
    Each leaf also carries 'boxes': {box id: (MP loss to the box inlet,
    same with the safety factor)}, the worst path to each box of that root.

    Fan system effect (Colin, 2026-10-29; lib/fan_system_effect.py):
    fan_effect is the dialog's {'on', 'fan_type', 'blast_ratio'}, None or
    off for none. For each unit root, each unit connection is walked to its
    first disruption inside its window (outlet: Le; inlet: the table's floor
    distance, 5 D or 5 H). An elbow is priced (outlet: SR7, worst of
    positions A to D; inlet: ED7-2 / ER7-1) and added to the subtotal BEFORE
    the safety factor. At outlets and inlets alike the straight duct, any
    transition on it and that elbow are charged once, as the larger of the
    fan system effect and their own losses (Colin, 2026-11-04,
    _fse_group_max); a tap or other fitting inside the window is flagged,
    not priced.

    The walk leaves a unit at the unit connection's velocity (first duct cfm
    over the connector area) and the first duct's class, so a fitting right
    on the unit is priced; it started at 0 fpm before 2026-10-29. Each leaf carries 'fse_inwc' (0.0 when none) and
    'fse_notes' [(kind, text)], kind 'note' or 'warn'. A duct-rooted pick
    takes the picked duct's unit side end as the unit outlet or inlet (Colin,
    2026-11-02), with a note naming the duct.
    """
    dr_by_int_id = {}
    for dr in all_duct_results.values():
        dr_by_int_id[dr.element_id] = dr

    # Ceiling exhaust fans (Colin, 2026-11-02): {fan root id: cap node id}.
    # The run ends at the wall or roof cap, priced as the 'cap' component in
    # place of a diffuser, and the fan system effect is at the fan OUTLET.
    discharge_roots = discharge_roots or {}
    inline_fan_roots = set(inline_fan_roots or ())
    inline_cache = {}       # root id -> _inline_sides(root id)
    fan_kinds = {}          # root id -> 'inline', 'roof' or None (exhaust fans)
    cap_ids = set(discharge_roots.values())

    box_ids  = vav_box_ids or set()
    box_drop = (vav_box_drop_inwc if vav_box_drop_inwc is not None
                else fitting_tables.component_of('damper', comp_values))

    def _comp(key):
        """One component's drop: the VAV box value, mid tree equipment at the
        balancing damper value, else the component table."""
        if key == 'vav_box':
            return box_drop
        if key == 'mid_equip':
            return fitting_tables.component_of('damper', comp_values)
        return fitting_tables.component_of(key, comp_values)

    def _kind(n):
        """Node kind for vav_checks' graph helpers."""
        if n in dr_by_int_id:
            return 'duct'
        if n in all_terminals:
            return 'terminal'
        e = all_nodes.get(n)
        if e is None:
            return None
        if hvac_graph.is_equipment(e):
            return 'equipment'
        if hvac_graph.is_fitting(e):
            return 'fitting'
        if hvac_graph.is_accessory(e):
            return 'accessory'
        return None

    first_duct_cache = {}
    def _first_duct(start):
        """DuctResult at start, or the nearest one below it through fittings
        and accessories (vav_checks.nearest_ducts_down), or None."""
        if start not in first_duct_cache:
            if start in dr_by_int_id:
                found = dr_by_int_id[start]
            else:
                below = vav_checks.nearest_ducts_down(start, all_children, _kind)
                found = dr_by_int_id.get(below[0]) if below else None
            first_duct_cache[start] = found
        return first_duct_cache[start]

    def _side_below(start):
        """System class of the branch starting at start, or None."""
        if start in all_terminals:
            return all_terminals[start][1]
        d = _first_duct(start)
        return d.sys_class if d is not None else None

    # Cache per-duct take-off children and end-of-main, so a duct reached by
    # several candidate paths is not re-walked for each one.
    tap_cache = {}
    def _taps(nid):
        if nid not in tap_cache:
            tap_cache[nid] = hvac_graph.takeoff_child_ids(nid, all_nodes, all_children)
        return tap_cache[nid]

    cont_cache = {}
    def _continues(nid):
        if nid not in cont_cache:
            cont_cache[nid] = hvac_graph.duct_continues_past(nid, all_nodes, all_children)
        return cont_cache[nid]

    cap_cache = {}
    def _capped(nid):
        if nid not in cap_cache:
            cap_cache[nid] = hvac_graph.duct_is_capped(nid, all_nodes, all_children)
        return cap_cache[nid]

    # ── ASHRAE 2001 lookups (Colin, 2026-10-23) ─────────────────────────────
    # Every elbow, transition and tap is priced from the 2001 Fundamentals
    # table for its own geometry and flows (ashrae_fittings). When the geometry
    # cannot be read, or no 2001 table fits, the RJA worksheet C is used and a
    # FALLBACK note is reported, so the total is never silently short.
    parent_of = {}
    for _p, _cs in all_children.items():
        for _c in _cs:
            parent_of.setdefault(_c, _p)

    def _is_round_duct(dr):
        return 'dia' in (getattr(dr, 'size', '') or '')

    def _round_d_in(dr):
        """Diameter of a round main in inches, or None."""
        if not _is_round_duct(dr) or not dr.area_ft2:
            return None
        return math.sqrt(4.0 * dr.area_ft2 / math.pi) * 12.0

    def _is_45_tap(fam, main_round=True):
        # Every tap on a rect main is a 45 deg entry: RJA taps branches and
        # sub mains at 45 deg, so a rect supply main reads SR5-13 for any tap
        # (Colin, 2026-11-12). Return tables on a rect main have no 45 entry
        # version, so this changes supply only. Round mains go by family name.
        if not main_round:
            return True
        f = (fam or '').lower()
        return 'shoe' in f or 'bevel' in f or '45' in f

    def _tap_branch_area(tap_id):
        """Area ft2 of a tap's branch opening: the smaller of its connectors
        (a beveled tap widens toward the main), or None."""
        elem = all_nodes.get(tap_id)
        if elem is None:
            return None
        try:
            areas = [a for a, _r in hvac_graph.connector_areas_ft2(elem) if a > 0]
        except Exception:
            areas = []
        return min(areas) if areas else None

    def _branch_flow(tap_id):
        """(cfm, area ft2) of the branch a tap feeds, or (None, None).

        The cfm is every duct the branch splits into below the tap (a tap
        feeding a tee counts both legs), else the terminals right on it. The
        area is the tap's own branch opening, else the one branch duct's."""
        below = vav_checks.nearest_ducts_down(tap_id, all_children, _kind)
        ducts = [dr_by_int_id[d] for d in below if d in dr_by_int_id]
        q = sum([d.cfm for d in ducts], 0.0)
        if q <= 0.0:
            q = sum([all_terminals[ch][0] for ch in all_children.get(tap_id, [])
                     if ch in all_terminals], 0.0)
        if q <= 0.0:
            return None, None
        ab = _tap_branch_area(tap_id)
        if ab is None and len(ducts) == 1:
            ab = ducts[0].area_ft2
        return q, ab

    tapflow_cache = {}
    tapafter_cache = {}     # duct id -> {tap id: main cfm left past its station}
    _TAP_STATION_FT = 0.1   # taps this close share one station (F2, 2026-11-04)
    def _tap_flows(duct_id):
        """({tap id: (local common cfm, branch cfm, branch area, feet)}, ordered).

        On an unbroken Revit main the duct's cfm is the flow at its unit end;
        each tap sees that less the branches taken off ahead of it, so taps are
        put in order along the duct (hvac_graph.tap_positions_ft). When the
        order cannot be read every tap sees the duct's own cfm (NOTE).

        Taps within _TAP_STATION_FT of each other are one station (Colin,
        2026-11-04, F2): each sees the flow arriving there and is priced from
        its own geometry and flow, feet is the station's, and the main loses
        the whole station's branch flow past it. Ties sort by element id, so
        Revit's connector order can never change the answer."""
        if duct_id in tapflow_cache:
            return tapflow_cache[duct_id]
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        pos = {}
        elem = all_nodes.get(duct_id)
        par = parent_of.get(duct_id)
        if par is None and duct_rooted_ids and duct_id in duct_rooted_ids and taps:
            # A picked duct has no parent in the walk: its taps are measured
            # from its unit side neighbor, the end the fan system effect walk
            # also uses (F2, Colin 2026-11-02).
            side = _unit_side_firsts(duct_id)
            par = side[0][0] if side else None
            fan_system_effect.log.info('tap order: picked duct %s measured from %s',
                                       duct_id, par if par is not None else 'NOT FOUND')
        if elem is not None and par is not None and taps:
            try:
                pos = hvac_graph.tap_positions_ft(elem, par)
            except Exception:
                pos = {}
        ordered = bool(taps) and all(t in pos for t in taps)
        order = (sorted(taps, key=lambda t: (pos[t], t)) if ordered
                 else sorted(taps))
        stations = []
        for t in order:
            if (ordered and stations and
                    pos[t] - pos[stations[-1][0]] <= _TAP_STATION_FT + 1e-9):
                stations[-1].append(t)
            else:
                stations.append([t])
        q = dr.cfm
        out, after = {}, {}
        for group in stations:
            q_in, taken = q, 0.0
            for t in group:
                qb, ab = _branch_flow(t)
                out[t] = (q_in, qb, ab, pos.get(group[0]))
                taken += qb or 0.0
            if ordered:
                q = max(q_in - taken, 0.0)
            for t in group:
                after[t] = q
            if len(group) > 1:
                fan_system_effect.log.info('tap station: duct %s taps %r at %s ft share '
                                           '%.0f cfm', duct_id, group,
                                           pos.get(group[0]), q_in)
        tapflow_cache[duct_id] = (out, ordered)
        tapafter_cache[duct_id] = after
        return tapflow_cache[duct_id]

    def _tap_c_positive(c_table, c_ws, table, what):
        """(C used, FALLBACK text or None). A tap never reduces the system
        and always carries some loss (Colin, 2026-11-09): a 2001 table C of
        0 or less (ASHRAE prints negative junction C for the slower stream,
        supply or return) is replaced by the RJA worksheet C for the same
        case, applied at the same Pv. A positive table C stays as read.
        Replaces the 2026-10-23 floor at 0 on return and exhaust."""
        if c_table > 0.0:
            return c_table, None
        return c_ws, ('FALLBACK: %s, 2001 %s %.2f is not positive; RJA worksheet C '
                      '%.2f used' % (what, table, c_table, c_ws))

    tapprice_cache = {}
    def _tap_turn(tap_id):
        """(loss in. wc, label, notes) for the air turning off through tap_id,
        or None when no 2001 table fits.

        The printed Cb is referenced to the branch section (2001 Example 8,
        Eq. 35). ashrae_fittings converts it to Cc,b, referenced to the main,
        at the Ab/Ac the table was read at (Eq. 34), so ER5-2 / ER5-3 keep
        their printed Ab/Ac 0.5 geometry; the loss is Cc,b x the main's Pv
        (Colin, 2026-10-23: mains are designed at 600 fpm return and 800 fpm
        supply, low pressure). A Cc,b of 0 or less takes the RJA worksheet C
        instead (_tap_c_positive, Colin, 2026-11-09)."""
        if tap_id in tapprice_cache:
            return tapprice_cache[tap_id]
        res = None
        main = parent_of.get(tap_id)
        dr = dr_by_int_id.get(main)
        elem = all_nodes.get(tap_id)
        if dr is not None and elem is not None and dr.area_ft2:
            flows, _ord = _tap_flows(main)
            qc, qb, ab, _ft = flows.get(tap_id, (dr.cfm, None, None, None))
            if qc and qb and ab:
                fam = hvac_graph.fitting_family_name(elem)
                # A main that ends in a tap is a turn into the branch, not a
                # junction (Colin, 2026-10-26): all of its air turns (Qb/Qc at
                # 1.0), so it is priced as a 90 degree elbow at the main's own
                # Pv. Round main: die stamped r/D 1.5, the company standard.
                # Rect main: smooth radius elbow, r/W 1.5, H/W as drawn.
                # Unreadable size falls through to the tap table below.
                # A CAPPED main runs on past its last tap to the cap: that
                # tap is a real side tap, priced from the tap table (Colin,
                # 2026-11-05, TB-M10C).
                if (min(qb / qc, 1.0) >= 0.99 and not _continues(main)
                        and not _capped(main)):
                    ec, elabel, enotes = None, None, []
                    if _is_round_duct(dr):
                        _d = _round_d_in(dr)
                        if _d:
                            ec, elabel, enotes = ashrae_fittings.round_elbow_c(
                                _d, 1.5, 90.0)
                    else:
                        _wh = re.match(r'\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)',
                                       getattr(dr, 'size', '') or '')
                        if _wh and float(_wh.group(1)) > 0:
                            ec, elabel, enotes = ashrae_fittings.rect_smooth_elbow_c(
                                1.5, float(_wh.group(2)) / float(_wh.group(1)), 90.0)
                    if ec is not None:
                        vc = qc / dr.area_ft2
                        res = (ec * hvac_graph.velocity_pressure_inwg(vc),
                               '%s, end of main turn, C %.2f at main %.0f fpm'
                               % (elabel, ec, vc),
                               _located('tap %s, end of main %s (%s %s)' % (
                                   tap_id, main, dr.sys_class, getattr(dr, 'size', '')), enotes))
                        tapprice_cache[tap_id] = res
                        return res
                cb, label, notes = ashrae_fittings.tap_branch_c(
                    vav_checks.is_supply_class(dr.sys_class),
                    'round' if hvac_graph.fitting_is_round(elem) else 'rect',
                    'round' if _is_round_duct(dr) else 'rect',
                    _is_45_tap(fam, _is_round_duct(dr)), ab / dr.area_ft2, min(qb / qc, 1.0),
                    main_d_in=_round_d_in(dr))
                if cb is not None:
                    vc = qc / dr.area_ft2
                    c_use, fb = _tap_c_positive(
                        cb, fitting_tables.takeoff_c(dr.sys_class, c=c_values),
                        label.split(' ')[0] + ' Cc,b',
                        'tap %s into branch off duct %s' % (tap_id, main))
                    if fb is None:
                        detail = '%s at main %.0f fpm' % (label, vc)
                    else:
                        # The worksheet C is the one the take-off reads
                        # (_takeoff_key: 'RJA worksheet', C from 'C x.xx').
                        detail = ('RJA worksheet tap into branch, C %.2f at main %.0f fpm'
                                  % (c_use, vc))
                    res = (c_use * hvac_graph.velocity_pressure_inwg(vc), detail,
                           _located('tap %s on main %s (%s %s)' % (
                               tap_id, main, dr.sys_class, getattr(dr, 'size', '')), notes)
                           + ([fb] if fb else []))
        tapprice_cache[tap_id] = res
        return res

    passprice_cache = {}
    def _tap_pass(duct_id, t):
        """(loss in. wc, table, C, notes, fallback) for the air continuing past
        tap t on main duct_id: Cc,s (converted from the printed Cs at the
        As/Ac the table was read at) x the main's Pv at the tap. On an
        unbroken main As = Ac; the Cs tables stop at As/Ac 0.9 and are read on
        that row (stated in the METHOD)."""
        key = (duct_id, t)
        if key in passprice_cache:
            return passprice_cache[key]
        dr = dr_by_int_id[duct_id]
        flows, _ord = _tap_flows(duct_id)
        qc, qb, ab, _ft = flows.get(t, (dr.cfm, None, None, None))
        elem = all_nodes.get(t)
        res = None
        # Why the table could not be read, for the FALLBACK line (Colin,
        # 2026-11-02: say why, not just "unreadable").
        if not qb:
            why = 'the branch carries no airflow, see the Flow = 0 warning'
        elif not qc:
            why = 'main airflow unreadable'
        elif not dr.area_ft2 or elem is None:
            why = 'main size or tap unreadable'
        else:
            why = 'outside the table'
        if qc and qb and dr.area_ft2 and elem is not None:
            fam = hvac_graph.fitting_family_name(elem)
            qs = max(qc - qb, 0.0)
            cs, lbl, _n = ashrae_fittings.tap_main_c(
                vav_checks.is_supply_class(dr.sys_class),
                'round' if hvac_graph.fitting_is_round(elem) else 'rect',
                'round' if _is_round_duct(dr) else 'rect',
                _is_45_tap(fam, _is_round_duct(dr)), 1.0, qs / qc,
                ab / dr.area_ft2 if ab else None, main_d_in=_round_d_in(dr))
            if cs is not None:
                vc = qc / dr.area_ft2
                c_use, fb = _tap_c_positive(
                    cs, fitting_tables.takeoff_main_c(dr.sys_class, c_values),
                    lbl.split(' ')[0] + ' Cc,s', 'tap %s passed on duct %s' % (t, duct_id))
                res = (c_use * hvac_graph.velocity_pressure_inwg(vc),
                       'RJA worksheet' if fb else lbl.split(' ')[0] + ' Cc,s', c_use, [], fb)
        if res is None:
            c_ws = fitting_tables.takeoff_main_c(dr.sys_class, c_values)
            one = c_ws * hvac_graph.velocity_pressure_inwg(dr.fpm)
            res = (one, 'RJA worksheet', c_ws, [],
                   'FALLBACK: tap %s passed on duct %s, RJA worksheet C %.2f '
                   '(%s)' % (t, duct_id, c_ws, why))
        passprice_cache[key] = res
        return res

    def _taps_passed(duct_id, child_id):
        """(loss in. wc, count, [(table, C, loss)], report notes) for the taps
        on main duct_id the run passes before leaving through child_id. Only
        taps AHEAD of the exit are passed when the order is known; a run
        continuing on the main passes all."""
        taps = _taps(duct_id)
        flows, ordered = _tap_flows(duct_id)
        notes = []
        if child_id in taps:
            if ordered:
                mine = flows[child_id][3]
                passed = [t for t in taps if t != child_id and flows[t][3] < mine]
            else:
                passed = [t for t in taps if t != child_id]
                notes.append(_tap_order_text(duct_id))
        else:
            passed = list(taps)
        loss, parts = 0.0, []
        for t in passed:
            one, table, c, _n, fb = _tap_pass(duct_id, t)
            loss += one
            parts.append((table, c, one))
            if fb:
                notes.append(fb)
        return loss, len(passed), parts, notes

    def _remaining_cfm(duct_id):
        """Flow left on an unbroken main past its last tap: the duct's unit
        end cfm less every tap's branch cfm, the same flow _seg_friction ends
        on. None when the duct has no taps or their order cannot be read."""
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        if not taps:
            return None
        flows, ordered = _tap_flows(duct_id)
        if not ordered:
            return None
        return max(dr.cfm - sum([(flows[t][1] or 0.0) for t in taps], 0.0), 0.0)

    def _tap_order_note(duct_id):
        """The NOTE for a main whose tap positions cannot be read, or None.
        Without the order, _seg_friction keeps the whole duct at its unit end
        cfm (Colin, 2026-11-02: say so in the run's notes)."""
        dr = dr_by_int_id[duct_id]
        if not _taps(duct_id) or not dr.area_ft2 or _tap_flows(duct_id)[1]:
            return None
        return _tap_order_text(duct_id)

    def _tap_order_text(duct_id):
        """One NOTE per duct whose tap positions cannot be read (Colin,
        2026-11-11): the tap pricing and the friction said together, so the
        two cases print as one line (System Corrections drops repeats)."""
        return ('NOTE: tap positions on duct %s unreadable: every tap priced at the '
                'full duct flow, friction charged over its whole length at its unit '
                'end cfm' % duct_id)

    def _seg_friction(duct_id, child_id):
        """(friction in. wc, feet charged) on an unbroken main with taps, or
        None to keep the duct's own whole length figure.

        The main keeps one size but its flow drops at every tap, so the
        friction is the sum of its pieces, each at its own flow (Colin,
        2026-11-02): unit end to tap 1 at the duct cfm, tap 1 to tap 2 at
        that less tap 1's branch, and so on. A run leaving through a tap is
        charged only up to that tap. Needs the tap order (positions read);
        without it the whole duct is charged at its unit end cfm, as before."""
        dr = dr_by_int_id[duct_id]
        taps = _taps(duct_id)
        if not taps or not dr.area_ft2:
            return None
        flows, ordered = _tap_flows(duct_id)
        if not ordered:
            return None
        length = float(dr.length_ft or 0.0)
        stop = (min(float(flows[child_id][3]), length) if child_id in taps
                else length)
        d_h = getattr(dr, 'd_h_in', None)
        eps = getattr(dr, 'roughness_ft', None)
        q, x0, total = dr.cfm, 0.0, 0.0
        after = tapafter_cache.get(duct_id, {})
        for t in sorted(taps, key=lambda t: (flows[t][3], t)):
            x1 = min(max(float(flows[t][3]), x0), stop)
            if x1 > x0 and q > 0:
                total += (x1 - x0) / 100.0 * hvac_graph.duct_friction_loss_per_100ft(
                    q / dr.area_ft2, d_h, eps)
            x0 = x1
            if x0 >= stop:
                break
            q = after.get(t, max(flows[t][0] - (flows[t][1] or 0.0), 0.0))
        if stop > x0 and q > 0:
            total += (stop - x0) / 100.0 * hvac_graph.duct_friction_loss_per_100ft(
                q / dr.area_ft2, d_h, eps)
        hvac_graph.log.info('_seg_friction: duct %s exit %s, %.2f of %.2f ft, %.5f in. wc '
                            '(whole duct at unit end cfm %.5f)', duct_id, child_id, stop,
                            length, total, dr.friction_loss_inwc)
        return total, stop

    def _shown_class(nid, sys_class):
        """The system name printed for a fitting. A discharge duct carries
        flow_class 'Supply Air' so its fittings read the supply (away from the
        fan) tables (hvac_graph.build_network); the report names the duct's own
        classification, Exhaust Air on an exhaust fan discharge (Colin,
        2026-11-09, TB-M10C transition 3696775)."""
        if nid is None:
            return sys_class
        for d in vav_checks.nearest_ducts_down(nid, all_children, _kind):
            dr = dr_by_int_id.get(d)
            if (dr is not None and getattr(dr, 'flow_class', None) == sys_class
                    and dr.sys_class):
                return dr.sys_class
        return sys_class

    # ── tees, crosses and wyes (Colin, 2026-11-11) ──────────────────────────
    # Priced on the leg the run LEAVES through, so in the child loop of the
    # walk, not when the fitting is reached: the branch pays Cc,b, the
    # straight Cc,s, x the common (unit side) leg Pv, the tap convention.
    # Legs, areas and axes from the fitting's own connectors; each leg's cfm
    # is the ducts below it. Unreadable legs or flows: the RJA worksheet, as
    # before, flagged FALLBACK.
    _JUNCTION_ROLES = ('tee', 'cross', 'wye')
    _UNREADABLE = 'geometry or flows unreadable'
    junction_cache = {}
    # Why a junction, or one exit leg of it, is on the worksheet: the reason
    # in its FALLBACK line (review, 2026-11-11).
    junction_why = {}
    junction_exit_why = {}

    def _leg_cfm(leg_id):
        """cfm through one leg of a junction: the duct on it, the terminal on
        it, else every duct below it."""
        d = dr_by_int_id.get(leg_id)
        if d is not None:
            return float(d.cfm or 0.0)
        if leg_id in all_terminals:
            return float(all_terminals[leg_id][0] or 0.0)
        below = vav_checks.nearest_ducts_down(leg_id, all_children, _kind)
        q = sum([float(dr_by_int_id[x].cfm or 0.0) for x in below if x in dr_by_int_id], 0.0)
        if q <= 0.0:
            q = sum([float(all_terminals[ch][0] or 0.0) for ch in all_children.get(leg_id, [])
                     if ch in all_terminals], 0.0)
        return q

    def _junction(nid):
        """{'role', 'legs', 'q', 'qc', 'common', 'layout'} for a tee, cross or
        wye whose every leg reads, else None (the worksheet prices it)."""
        if nid in junction_cache:
            return junction_cache[nid]
        res = None
        why = _UNREADABLE
        elem = all_nodes.get(nid)
        par = parent_of.get(nid)
        kids = list(all_children.get(nid, []))
        role = (fitting_tables.classify_fitting(hvac_graph.fitting_family_name(elem))
                if elem is not None else None)
        if role in _JUNCTION_ROLES and par is not None and len(kids) >= 2:
            legs = hvac_graph.junction_legs(elem, [par] + kids)
            if par in legs and all([k in legs for k in kids]):
                q = dict((k, _leg_cfm(k)) for k in kids)
                qc = sum(q.values(), 0.0)
                # A leg with no airflow (terminals without Flow) would read a
                # wrong split; only an end cap carries none (review,
                # 2026-11-11).
                dry = sorted([k for k in kids if q[k] <= 0.0 and not _is_endcap_node(k)])
                lay = ashrae_fittings.junction_layout(
                    legs[par]['dir'], dict((k, legs[k]['dir']) for k in kids))
                if dry:
                    why = ('%s: leg %s carries no airflow (a terminal without Flow '
                           'below it)' % (_UNREADABLE, ', '.join([str(k) for k in dry])))
                elif lay is None:
                    why = '%s: connector axes unreadable' % _UNREADABLE
                elif qc > 0.0:
                    res = {'role': role, 'legs': legs, 'q': q, 'qc': qc,
                           'common': legs[par], 'layout': lay}
                    # Only reached on the worksheet where the main walk does
                    # not price it per leg (between the unit and the first
                    # disruption, at the fan velocity).
                    why = 'read, but priced on the worksheet at the fan velocity'
                else:
                    why = '%s: no airflow through it' % _UNREADABLE
            else:
                why = '%s: connector to %s not read' % (_UNREADABLE, ', '.join(
                    [str(k) for k in [par] + kids if k not in legs]))
        elif role in _JUNCTION_ROLES:
            why = '%s: fewer than two legs leave it on this run' % _UNREADABLE
        if role in _JUNCTION_ROLES:
            hvac_graph.log.info('_junction: %s %s %s', role, nid,
                                'read' if res else 'NOT READ, worksheet')
        junction_cache[nid] = res
        junction_why[nid] = why
        return res

    junction_exit_cache = {}
    def _junction_exit(nid, child_id, sys_class):
        """(loss in. wc, detail, notes, leg fpm, leg area ft2) for the run
        leaving junction nid through child_id, or None (worksheet)."""
        key = (nid, child_id, sys_class)
        if key in junction_exit_cache:
            return junction_exit_cache[key]
        res = None
        j = _junction(nid)
        if j is None:
            junction_exit_why[key] = junction_why.get(nid, _UNREADABLE)
        elif child_id not in j['legs'] or not j['q'].get(child_id):
            junction_exit_why[key] = 'leg %s carries no airflow' % child_id
        if j is not None and child_id in j['legs'] and j['q'].get(child_id):
            legs, q, qc, lay, role = j['legs'], j['q'], j['qc'], j['layout'], j['role']
            com = j['common']
            ac = com['area_ft2']
            leg = legs[child_id]
            vc = qc / ac
            sup = vav_checks.is_supply_class(sys_class)
            straight, branches = lay['straight'], lay['branches']
            c, label, n = None, '', []
            legname = 'branch'
            if straight is None:
                # No leg in line with the common: a symmetrical wye, or a
                # bullhead tee fed through its branch (SR5-14 / ER5-4).
                stands = {'tee': 'bullhead tee', 'cross': 'cross'}.get(role)
                c, label, n = ashrae_fittings.split_c(
                    sup, leg['area_ft2'] / ac, min(q[child_id] / qc, 1.0), stands_for=stands)
                if c is None:
                    junction_exit_why[key] = label
                ws = fitting_tables.c_of('wye_bullhead', c_values)
            else:
                if child_id == straight:
                    legname = 'straight'
                    # Cs is read against the largest branch's geometry.
                    b = (sorted(branches, key=lambda k: (-q.get(k, 0.0), k))[0]
                         if branches else None)
                else:
                    b = child_id
                if b is not None:
                    bl = legs[b]
                    entry_45 = branches.get(b, 90.0) < ashrae_fittings.ENTRY_45_MAX_DEG
                    fn = (ashrae_fittings.cross_c if (role == 'cross' or len(branches) >= 2)
                          else ashrae_fittings.tee_c)
                    c, label, n = fn(sup, com['shape'], bl['shape'], entry_45, legname,
                                     bl['area_ft2'] / ac, legs[straight]['area_ft2'] / ac,
                                     min(q.get(b, 0.0) / qc, 1.0),
                                     min(q.get(straight, 0.0) / qc, 1.0),
                                     main_d_in=com['d_in'])
                    if c is None:
                        if ashrae_fittings._junction_keys(
                                sup, bl['shape'], com['shape'], entry_45,
                                com['d_in']) == (None, None):
                            junction_exit_why[key] = (
                                'no 2001 table for a %s branch on a %s main, %s'
                                % (bl['shape'], com['shape'],
                                   'supply' if sup else 'return / exhaust'))
                        else:
                            junction_exit_why[key] = '%s: %s' % (
                                _UNREADABLE, '; '.join(n or []) or label)
                ws = (fitting_tables.takeoff_c(sys_class, c=c_values) if legname == 'branch'
                      else fitting_tables.takeoff_main_c(sys_class, c_values))
            if c is not None:
                c_use, fb = _tap_c_positive(
                    c, ws, label.split(' ')[0] + (' Cc,s' if legname == 'straight' else ' Cc,b'),
                    '%s %s leaving through %s' % (role, nid, child_id))
                if fb is None:
                    detail = '%s at common %.0f fpm' % (label, vc)
                else:
                    detail = ('RJA worksheet %s %s, C %.2f at common %.0f fpm'
                              % (role, legname, c_use, vc))
                where = '%s %s (%s, %s, leaving through %s %s)' % (
                    role, nid, _shown_class(nid, sys_class),
                    _size_text(com['shape'] == 'round', com['d_in'], com['w_in'], com['h_in']),
                    legname, child_id)
                res = (c_use * hvac_graph.velocity_pressure_inwg(vc), detail,
                       _located(where, n) + ([fb] if fb else []),
                       q[child_id] / leg['area_ft2'], leg['area_ft2'])
        junction_exit_cache[key] = res
        return res

    def _ashrae_fitting(elem, role, fam, sys_class, up_fpm, up_area, loc_id=None):
        """(loss in. wc, label, notes, downstream area) from the 2001 tables,
        or None to fall back to the worksheet. sys_class picks the table; the
        printed name is _shown_class."""
        eid = loc_id if loc_id is not None else '?'
        shown = _shown_class(loc_id, sys_class)
        if role == 'elbow':
            g = hvac_graph.elbow_geometry(elem)
            if g is None:
                return None
            where = ('elbow %s (%s, %s, %.0f deg)'
                     % (eid, shown, _size_text(g['round'], g['d_in'], g['w_in'],
                                                   g['h_in']), g['angle'] or 90.0))
            f = (fam or '').lower()
            if g['round']:
                if not g['d_in']:
                    return None
                rd = (g['r_in'] / g['d_in']) if g['r_in'] else None
                c, label, notes = ashrae_fittings.round_elbow_c(g['d_in'], rd, g['angle'])
            else:
                if not g['w_in'] or not g['h_in']:
                    return None
                hw = g['h_in'] / g['w_in']
                if 'radius' in f and 'miter' not in f:
                    if not g['r_in']:
                        return None
                    c, label, notes = ashrae_fittings.rect_smooth_elbow_c(
                        g['r_in'] / g['w_in'], hw, g['angle'])
                else:
                    c, label, notes = ashrae_fittings.rect_mitered_elbow_c(hw, g['angle'])
            if c is None:
                return None
            return (c * hvac_graph.velocity_pressure_inwg(up_fpm),
                    '%s, C %.2f' % (label, c), _located(where, notes), None)
        if role == 'transition':
            ends, theta = hvac_graph.transition_geometry(elem)
            if not ends or theta is None or not up_area or not up_fpm:
                return None
            # Section 1 is the walk-upstream (unit side) end and o the far end,
            # for supply (air 1 -> o) and return (air o -> 1, 1 toward the fan)
            # alike, so the walk order IS the table order.
            e1, eo = sorted(ends, key=lambda e: abs(e[0] - up_area))
            c, label, notes = ashrae_fittings.transition_c(
                vav_checks.is_supply_class(sys_class), e1[1], eo[1], e1[0], eo[0], theta)
            if c is None:
                return None
            vo = up_fpm * up_area / eo[0]
            where = ('transition %s (%s, %s to %s, %.1f deg)'
                     % (eid, shown,
                        _size_text(e1[1] == 'round', e1[2], e1[2], e1[3]),
                        _size_text(eo[1] == 'round', eo[2], eo[2], eo[3]), theta))
            return (c * hvac_graph.velocity_pressure_inwg(vo),
                    '%s, Co %.2f at %.0f fpm' % (label, c, vo),
                    _located(where, notes), eo[0])
        return None

    def _terminal_mark(nid):
        """The terminal's Mark, or None (logged)."""
        mark = None
        try:
            p = all_nodes.get(nid).LookupParameter('Mark')
            mark = p.AsString() if p is not None else None
        except Exception:
            fan_system_effect.log.info('terminal %s Mark NOT FOUND', nid)
        return mark

    def _terminal_label(nid):
        """'D-6 (id 3639741)': the terminal's Mark and id, so the end of the
        index run can be found on the plan. Id alone when there is no Mark."""
        mark = _terminal_mark(nid)
        return '%s (id %s)' % (mark, nid) if mark else 'id %s' % nid

    def _run_note(rec, facts, nid):
        """The one NOTE for the unit connection a run leaves through (Colin,
        2026-11-11), or None. rec: the connection's record (or the tap sub
        record the run leaves through); facts: what _fse_group_max found
        (priced values, or own losses with none priced); nid: the run's
        terminal. A connection not checked keeps its own NOTE instead."""
        if rec is None or not rec.get('where_item'):
            return None
        facts = facts or {}
        priced = facts.get('priced') if rec.get('item') is not None else None
        none = rec.get('none')
        if priced is None and (not none or none.startswith('not checked')):
            return None
        via = rec.get('via')
        if via is not None:
            via = dict(via)
            if priced is not None and facts.get('win_ft'):
                # the window the table was read at (an inlet: 5 x the
                # elbow's own D or H)
                via['win_ft'] = facts['win_ft']
                via['win_short'] = facts.get('win_short') or via.get('win_short')
                via['size'] = facts.get('size')
            elif priced is None and none and str(via.get('id')) in none:
                via['bare'] = True
        mark = _terminal_mark(nid)
        term = '%s, id %s' % (mark, nid) if mark else 'id %s' % nid
        return fan_system_effect.run_note(
            rec['where'], term, passed=rec.get('passed') or (), via=via,
            priced=priced, none=none, own=facts.get('own') if priced is None else None,
            curb_ft=rec.get('curb_ft') or 0.0)

    def _size_text(is_round, d_in, w_in, h_in):
        """'16x12"' or '10"dia' for a report note."""
        if is_round:
            return '%.0f"dia' % float(d_in or 0)
        return '%.0fx%.0f"' % (float(w_in or 0), float(h_in or 0))

    def _located(where, notes):
        """Lookup notes led by the fitting they belong to, so the reader can
        find it: 'transition 3639637 (Supply Air, 18x12" to 16x12", 7.6 deg):
        SR4-1: angle 7.6 deg is below ...'."""
        return ['%s: %s' % (where, n) for n in (notes or []) if n]

    _COMP_LABEL = dict((k, lbl) for k, lbl, _v in fitting_tables.COMPONENT_TABLE)
    _C_IN_DETAIL = re.compile(r'\b(?:C|Cb|Co|Cs)\s+(-?\d+\.\d+)')
    _CC_IN_DETAIL = re.compile(r'Cc,[bs]\s+(-?\d+\.\d+)')
    _TABLE_ID = re.compile(r'^([A-Z]{2}\d+-\d+)')

    def _takeoff_key(elem, fam, detail, counted, note):
        """(row name, table, C, feet) of one fitting's row in the external
        static take-off table (route order, one row each, 2026-11-11)."""
        role = fitting_tables.classify_fitting(fam)
        f = (fam or '').lower()
        if role == 'elbow':
            ang = hvac_graph.fitting_angle_deg(elem)
            ang = 90.0 if ang is None else ang
            if hvac_graph.fitting_is_round(elem):
                kind = 'Round elbow'
            elif 'radius' in f and 'miter' not in f:
                kind = 'Rect radius elbow'
            else:
                kind = 'Rect mitered elbow'
            name = '%s, %.0f deg' % (kind, ang)
        elif role == 'transition':
            name = 'Transition'
        elif role == 'takeoff':
            name = 'Tap, into branch'
        elif role in ('tee', 'cross', 'wye'):
            straight = ('Cc,s' in (detail or '') or ' straight,' in (detail or '')
                        or 'straight through' in (detail or ''))
            name = '%s, %s' % (role.capitalize(),
                               'straight through' if straight else 'into branch')
        elif role == 'offset':
            name = role.capitalize()
        else:
            name = fam or 'Fitting'
        m = _TABLE_ID.match(detail or '')
        if note and 'UNPRICED' in note and not note.startswith(('FALLBACK', 'NOTE')):
            table = 'UNPRICED'
        elif m:
            table = m.group(1)
        else:
            table = 'RJA worksheet'
        # A tap's take-off C is Cc,b: the value that multiplies the main Pv.
        cm = _CC_IN_DETAIL.search(detail or '') or _C_IN_DETAIL.search(detail or '')
        return (name, table, float(cm.group(1)) if cm else None, None)

    def _notes_text(notes):
        """Lookup notes (clamped read, blank cell, negative credit) as one
        NOTE line for the report, or None."""
        notes = [n for n in (notes or []) if n]
        return ('NOTE: ' + '; '.join(notes)) if notes else None

    def _price_fitting(elem, sys_class, up_fpm, up_area, against_flow=False,
                       nid=None, loc_id=None, junction_leg=None, fb_why=None):
        """(loss in. wc, counted 0/1, unpriced note or None, downstream area,
        detail text).

        One place that prices a fitting the path passes THROUGH, used by the main
        walk and by the unit-side chain below so the two can never disagree.
        Elbows, transitions and taps come from the ASHRAE 2001 tables first;
        the RJA worksheet prices anything else and anything unreadable (the
        note then starts FALLBACK).

        against_flow: the walk is going the opposite way to the air (the unit
        side chain of a picked supply duct). The 2001 transition tables are
        keyed to the walk order for the main walk only, so a reversed walk
        prices transitions on the worksheet, with its areas swapped so a return
        transition that grows toward the unit is priced as the contraction the
        air sees (Colin, 2026-10-21). The returned downstream area stays in
        WALK order, because the caller tracks area along the walk.

        junction_leg: a tee, cross or wye on the worksheet whose exit leg is
        known, 'straight' (worksheet tap main C) or 'branch'; None prices it
        as the branch, conservative. fb_why: the reason in its FALLBACK line
        (review, 2026-11-11).
        """
        fam  = hvac_graph.fitting_family_name(elem)
        role = fitting_tables.classify_fitting(fam)
        if role == 'takeoff' and nid is not None:
            t = _tap_turn(nid)
            if t is not None:
                # A FALLBACK line leads, so it is an Action required line
                # (duct_perf_console.correction_group reads the lead).
                fbs = [n for n in t[2] if n and n.startswith('FALLBACK')]
                rest = _notes_text([n for n in t[2] if n and not n.startswith('FALLBACK')])
                text = '; '.join(fbs + ([rest] if rest else [])) or None
                return t[0], 1, text, None, t[1]
        reversed_walk = against_flow != fitting_tables.flows_toward_unit(sys_class)
        if role in ('elbow', 'transition') and not (role == 'transition' and reversed_walk):
            a = _ashrae_fitting(elem, role, fam, sys_class, up_fpm, up_area,
                                loc_id=nid if loc_id is None else loc_id)
            if a is not None:
                return a[0], 1, _notes_text(a[2]), a[3], a[1]
        up_a, down_a = (None, None)
        flow_up, flow_down = (None, None)
        if role == 'transition':
            up_a, down_a = hvac_graph.transition_areas(elem, up_area)
            flow_up, flow_down = (down_a, up_a) if against_flow else (up_a, down_a)
        angle = hvac_graph.fitting_angle_deg(elem) if role == 'elbow' else None
        if role in _JUNCTION_ROLES and junction_leg == 'straight':
            # Straight through a tee, cross or wye: the worksheet's tap main
            # C, not the branch's (review, 2026-11-11).
            c, note = (fitting_tables.takeoff_main_c(sys_class, c_values),
                       '%s, straight through' % role)
        else:
            c, note = fitting_tables.fitting_c(
                role, sys_class,
                is_round=hvac_graph.fitting_is_round(elem),
                upstream_area_ft2=flow_up, downstream_area_ft2=flow_down,
                is_end_of_main=False, c=c_values, angle_deg=angle)
        fb = None
        if role in _JUNCTION_ROLES:
            why = (fb_why or (junction_why.get(nid) if nid is not None else None)
                   or _UNREADABLE)
            if junction_leg is None:
                why = '%s; exit leg not identified, priced as the branch (conservative)' % why
            fb = 'FALLBACK: %s: %s, RJA worksheet C %.2f (%s)' % (fam or '?', note, c, why)
        elif (role in ('elbow', 'takeoff')
                or (role == 'transition' and not reversed_walk)):
            fb = 'FALLBACK: %s: %s, RJA worksheet C %.2f (geometry or flows unreadable)' % (
                fam or '?', note, c)
        if c > 0.0:
            return (c * hvac_graph.velocity_pressure_inwg(up_fpm), 1, fb, down_a,
                    '%s, C %.2f worksheet' % (note, c))
        if 'UNPRICED' in note:
            return 0.0, 0, (fam or '?') + ': ' + note, down_a, 'UNPRICED'
        return 0.0, 0, None, down_a, note

    def _unit_side_firsts(root_id):
        """[(first node, nodes in its subtree)] for the children of a picked
        duct root that lead back toward the unit: subtrees with no terminal
        that do not start at a take-off. Shared by _unit_side_chain and the
        duct pick fan system effect, so the two agree on the unit side."""
        out = []
        for first in all_children.get(root_id, []):
            order = []
            todo = [first]
            has_term = False
            while todo:
                n = todo.pop()
                order.append(n)
                if n in all_terminals:
                    has_term = True
                    break
                todo.extend(all_children.get(n, []))
            if has_term:
                continue
            first_elem = all_nodes.get(first)
            if (first_elem is not None and hvac_graph.is_fitting(first_elem) and
                    fitting_tables.classify_fitting(
                        hvac_graph.fitting_family_name(first_elem)) == 'takeoff'):
                continue
            out.append((first, order))
        return out

    def _unit_side_chain(root_id):
        """Duct and fittings BEHIND a picked duct, on the unit side.

        With no unit on the ductwork the run is rooted at the picked duct, and the
        undirected walk hangs the stretch back toward the unit off the root as a
        side branch that reaches no terminal, so the path search never charged it.
        Colin, 2026-10-08: the fittings between the picked duct and the unit
        connection are ALWAYS part of the external static, so they are priced here
        and added to every path from this root.

        Only subtrees with no terminal and that do not start at a take-off count,
        so a tap hanging off the root is never mistaken for the unit side. The
        air in this stretch is the root's own flow, so velocity comes from the
        root's CFM over each duct's own area.

        Returns (friction, fitting loss, fittings priced, ducts, feet, unpriced).
        """
        root_dr = dr_by_int_id.get(root_id)
        if root_dr is None or not root_dr.fpm or root_dr.fpm <= 0:
            return (0.0, 0.0, 0, 0, 0.0, ())
        flow_cfm = root_dr.fpm * root_dr.area_ft2
        fric = fit = dlen = 0.0
        n_fit = n_duct = 0
        unpriced = ()
        for _first, order in _unit_side_firsts(root_id):
            up_area = root_dr.area_ft2
            for n in order:
                elem = all_nodes.get(n)
                dr   = dr_by_int_id.get(n)
                if dr is not None:
                    fpm = flow_cfm / dr.area_ft2 if dr.area_ft2 else 0.0
                    d_h = getattr(dr, 'd_h_in', 0.0)
                    if d_h > 0 and fpm > 0:
                        fric += (hvac_graph.duct_friction_loss_per_100ft(
                                     fpm, d_h, getattr(dr, 'roughness_ft', None))
                                 * dr.length_ft / 100.0)
                    dlen   += dr.length_ft
                    n_duct += 1
                    up_area = dr.area_ft2
                elif elem is not None and hvac_graph.is_fitting(elem):
                    up_fpm = flow_cfm / up_area if up_area else 0.0
                    # This stretch is walked from the picked duct back TOWARD the
                    # unit, the reverse of the main walk, so supply air runs
                    # against it and return air runs with it.
                    loss, counted, note, down_a, _detail = _price_fitting(
                        elem, root_dr.sys_class, up_fpm, up_area,
                        against_flow=not fitting_tables.flows_toward_unit(
                            root_dr.sys_class), loc_id=n,
                        fb_why='priced on the worksheet on the unit side chain, '
                               'which has no 2001 junction lookup')
                    fit   += loss
                    n_fit += counted
                    if note:
                        unpriced = unpriced + (note,)
                    if down_a:
                        up_area = down_a
        return (fric, fit, n_fit, n_duct, dlen, unpriced)

    # ── unit connections (Colin, 2026-10-29) ────────────────────────────────
    # The connector on the unit joined to each of its first nodes: its size
    # sets the velocity the walk starts with and the fan system effect's Ao.
    unit_conn_cache = {}
    def _unit_conn(root_id, first):
        """hvac_graph.equipment_connection_geometry for the unit root_id
        toward node first, or None (not equipment, or unreadable: logged)."""
        key = (root_id, first)
        if key not in unit_conn_cache:
            root_elem = all_nodes.get(root_id)
            geo = None
            if root_elem is not None and hvac_graph.is_equipment(root_elem):
                geo = hvac_graph.equipment_connection_geometry(root_elem, first)
            unit_conn_cache[key] = geo
        return unit_conn_cache[key]

    def _connection_cfm(first):
        """Air through the unit connection at node first: a duct's own cfm,
        else the sum over the nearest ducts below it. nearest_ducts_down
        reaches each duct once (it keeps a seen set and stops at the first
        duct on every path), so a tee right on the unit counts both legs once
        each (review, 2026-10-29). 0.0 when no duct below carries flow."""
        dr = dr_by_int_id.get(first)
        if dr is not None:
            return float(dr.cfm or 0.0)
        below = vav_checks.nearest_ducts_down(first, all_children, _kind)
        return sum([float(dr_by_int_id[d].cfm or 0.0)
                    for d in below if d in dr_by_int_id], 0.0)

    def _unit_start(root_id, first):
        """(fpm, area ft2, sys_class) the walk carries from the unit into node
        first, or None.

        A fitting right on the unit connection used to be priced at 0 fpm,
        because the walk starts with no upstream velocity (Colin, 2026-10-29:
        fix it). The flow is the whole connection's (_connection_cfm, both
        legs of a tee on the unit), the area the unit connector's (the
        section the air passes at the unit, and the end a transition there is
        oriented against), else the first duct's own.
        The class is the first duct's, so a fitting before any duct is priced
        as supply or return correctly."""
        root_elem = all_nodes.get(root_id)
        if root_elem is None or not hvac_graph.is_equipment(root_elem):
            return None
        flow = _first_duct(first)
        cfm = _connection_cfm(first)
        if flow is None or cfm <= 0:
            return None
        conn = _unit_conn(root_id, first)
        area = conn['area_ft2'] if conn is not None and conn.get('area_ft2') else flow.area_ft2
        if not area:
            return None
        return cfm / area, area, getattr(flow, 'flow_class', None) or flow.sys_class

    # ── fan system effect (Colin, 2026-10-29) ───────────────────────────────
    # The 2001 Fundamentals Eq. 38 (p. 34.14) adds FSE2 at the fan outlet and
    # FSE1 at the inlet. Priced from the first disruption inside each unit
    # connection's window; the math is fan_system_effect's.
    # The straight duct and swirl checks run every time; the fan system
    # effect is priced only when ticked (Colin, 2026-11-03, Q4). Unticked,
    # each connection keeps its RED lines and drops the rest (_checks_only).
    fse_price = bool(fan_effect and fan_effect.get('on'))
    fse_set = dict(fan_effect) if fan_effect else fan_system_effect.default_settings()
    _FSE_KIND_TEXT = {'takeoff': 'tap', 'split': 'branch', 'unknown': 'unrecognised fitting'}

    def _transition_length_ft(elem, nid):
        """A transition's own length in feet, the distance between its two
        connector origins (hvac_graph.transition_geometry). 0.0 when it
        cannot be read, logged."""
        ends, _theta = hvac_graph.transition_geometry(elem)
        try:
            length = float(ends[0][4].DistanceTo(ends[1][4]))
        except Exception as ex:
            fan_system_effect.log.info('fan system effect: transition %s length '
                                       'NOT FOUND (%s), 0.0 ft counted', nid, ex)
            return 0.0
        fan_system_effect.log.info('fan system effect: transition %s length %.2f ft',
                                   nid, length)
        return length

    def _is_tap_node(nid):
        elem = all_nodes.get(nid) if nid is not None else None
        return (elem is not None and hvac_graph.is_fitting(elem) and
                fitting_tables.classify_fitting(
                    hvac_graph.fitting_family_name(elem)) == 'takeoff')

    def _is_endcap_node(nid):
        elem = all_nodes.get(nid)
        return (elem is not None and fitting_tables.classify_fitting(
            hvac_graph.fitting_family_name(elem)) == 'endcap')

    def _fse_first_disruption(root_id, first, win, where='unit connection',
                              prev=None, seen=None, collect=None, passed=None,
                              through_taps=False, ducts=None, accessories=None,
                              l0=0.0, tap_log=None, drop=None):
        """(role or None, L ft, node id, notes): the first disruption within
        win feet of the unit connection at node first.

        l0: straight length counted before the first node, a roof curb that
        is never modeled (Colin, 2026-11-09). Every L returned counts it.

        L is the straight duct from the unit connection to it: duct lengths
        and transition lengths summed (Colin, 2026-11-02: a curb adapter is
        part of the straight length), unions and accessories passed at zero
        length. Every transition is passed; the first elbow is the one the
        system effect prices. A tap on an unbroken duct sits at its own
        distance along that duct (hvac_graph.tap_positions_ft from the unit
        end). role is a fitting_tables.classify_fitting role, 'takeoff' for
        a tap, 'split' for a duct that divides; None when nothing is within
        the window or the run ends inside it (a terminal or equipment).

        prev, seen: a duct pick starts at the picked duct itself (first ==
        root_id), prev its unit side neighbor (None when there is none) and
        seen the unit side nodes, which are not walked. A unit root leaves
        both None: prev and seen are the unit.

        collect: a list for the swirl check. Each elbow inside win is
        appended as (id, L ft) and the walk goes on past it, to the end of
        the window or the first disruption that is not an elbow.

        passed: a dict filled with {transition id: L ft} for every
        transition passed, for the fan velocity basis (Colin, 2026-11-03).

        through_taps: walk past taps on the main instead of stopping at
        them, for the straight duct check.

        ducts: a list filled with (duct id, L ft at its unit end) for every
        duct passed, so the straight duct before a RED elbow is RED too.

        accessories: a list filled with (id, L ft, family) for every damper
        passed (fan_system_effect.DAMPER_KINDS), for the damper in the fan
        clearance WARNING (Colin, 2026-11-04, F1). Still passed at zero
        length.

        Taps near the unit (Colin, 2026-11-11). tap_log: the pricing walk
        passes taps off a continuing main, filling (tap id, L ft, duct id)
        for each one inside win. drop: the pricing walk's dict; a duct whose
        air all leaves through its taps (capped, or a dead end) is a drop:
        returns ('drop', L at the duct, duct id) and fills drop with
        {'duct', 'taps': [(tap id, L ft)]}. A tap fitting ON the walk (the
        unit duct tees into the side of a main): the pricing walk returns
        'tap_turn' there; every other walk goes on along the main, a capped
        dead leg skipped, 'split' when the main runs both ways."""
        notes = []
        if seen is None:
            prev, seen = root_id, set([root_id])
        nid, l_ft = first, float(l0)
        while nid is not None and nid not in seen and l_ft < win:
            seen.add(nid)
            dr = dr_by_int_id.get(nid)
            elem = all_nodes.get(nid)
            kids = [c for c in all_children.get(nid, []) if c not in seen]
            if dr is not None:
                # Listed before its taps are read, so a tap returned below
                # names the duct it sits on as the last entry (the supply
                # outlet tap check reads that duct's De, Colin 2026-11-09).
                if ducts is not None:
                    ducts.append((nid, l_ft))
                if _is_tap_node(prev):
                    # Entered through a tap on its side: on along the main
                    # the way the air goes, a capped leg skipped.
                    own = set(_taps(nid))
                    live = [c for c in kids if c not in own and not _is_endcap_node(c)]
                    if not live:
                        return None, l_ft, None, notes   # all of it leaves through taps
                    if len(live) != 1:
                        return 'split', l_ft, nid, notes
                    try:
                        d_in = hvac_graph.tap_positions_ft(elem, live[0]).get(prev)
                    except Exception:
                        d_in = None
                    if d_in is None:
                        d_in = float(dr.length_ft or 0.0)
                        notes.append(('note', 'NOTE: tap %s position on duct %s '
                                      'unreadable, the whole duct length counted '
                                      'for the fan system effect' % (prev, nid)))
                    l_ft += d_in
                    prev, nid = nid, live[0]
                    continue
                taps = _taps(nid)
                pos = {}
                if taps and elem is not None and prev is not None:
                    try:
                        pos = hvac_graph.tap_positions_ft(elem, prev)
                    except Exception:
                        pos = {}
                end_tap = ([tp for tp in taps if hvac_graph.tap_joins_duct_end(elem, tp)]
                           if taps and elem is not None else [])
                if (end_tap
                        and (drop is not None or tap_log is not None or through_taps
                             or collect is not None)
                        and not [c for c in kids if c not in taps
                                 and not _is_endcap_node(c)]):
                    # A tap joined at this duct's END, not on its side: the
                    # duct runs into it (the unit duct tees into a main,
                    # Colin 2026-11-11). Walked as a fitting on the run.
                    l_ft += float(dr.length_ft or 0.0)
                    prev, nid = nid, end_tap[0]
                    continue
                if taps:
                    if all(t in pos for t in taps):
                        tid = min(taps, key=lambda t: pos[t])
                        at = pos[tid]
                    elif prev is None:
                        tid, at = taps[0], 0.0
                        notes.append(('note', 'NOTE: duct pick, nothing found on the '
                                      'unit side of duct %s, so its taps cannot be put '
                                      'in order from the unit end; tap %s taken at that '
                                      'end for the fan system effect' % (nid, tid)))
                    else:
                        tid, at = taps[0], 0.0
                        notes.append(('note', 'NOTE: tap position on duct %s unreadable, '
                                      'taken at its unit end for the fan system effect'
                                      % nid))
                    fan_system_effect.log.info('fan system effect: tap %s on duct %s '
                                               'at %.2f ft', tid, nid, l_ft + at)
                    if drop is not None and not [c for c in kids if c not in taps
                                                 and not _is_endcap_node(c)]:
                        # A drop: every bit of air leaves through its taps.
                        drop['duct'] = nid
                        drop['taps'] = sorted(
                            [(t, l_ft + pos.get(t, at if t == tid else 0.0)) for t in taps],
                            key=lambda x: (x[1], x[0]))
                        fan_system_effect.log.info('fan system effect: duct %s is a '
                                                   'drop, taps %r', nid, drop['taps'])
                        return 'drop', l_ft, nid, notes
                    if through_taps:
                        pass                # the main goes on past the tap
                    elif tap_log is not None:
                        # A continuing main: walked past (Colin, 2026-11-11).
                        for t in sorted(taps, key=lambda x: pos.get(x, 0.0)):
                            if l_ft + pos.get(t, 0.0) < win:
                                tap_log.append((t, l_ft + pos.get(t, 0.0), nid))
                    elif l_ft + at < win:
                        return 'takeoff', l_ft + at, tid, notes
                    else:
                        return None, l_ft + at, None, notes
                l_ft += float(dr.length_ft or 0.0)
                nxt = [c for c in kids if c not in taps]
            elif elem is not None and hvac_graph.is_fitting(elem):
                role = fitting_tables.classify_fitting(hvac_graph.fitting_family_name(elem))
                if role == 'transition':
                    # Passed, its own length counted in L; the main walk
                    # prices it as an ordinary fitting.
                    fan_system_effect.log.info('fan system effect: transition %s at '
                                               '%.2f ft passed', nid, l_ft)
                    if passed is not None:
                        passed[nid] = l_ft
                    l_ft += _transition_length_ft(elem, nid)
                elif collect is not None and role == 'elbow':
                    collect.append((nid, l_ft))
                elif role == 'takeoff' and drop is not None:
                    # The unit duct tees into a main: the air turns here.
                    fan_system_effect.log.info('fan system effect: tap %s on the walk at '
                                               '%.2f ft, a turn', nid, l_ft)
                    return 'tap_turn', l_ft, nid, notes
                elif role == 'takeoff' and (collect is not None or through_taps):
                    pass                    # on along the main past it
                elif role not in fitting_tables.ZERO_LOSS_ROLES:
                    fan_system_effect.log.info('fan system effect: %s %s at %.2f ft',
                                               role, nid, l_ft)
                    return role, l_ft, nid, notes
                nxt = kids
            elif elem is not None and hvac_graph.is_accessory(elem):
                if accessories is not None:
                    fam = hvac_graph.fitting_family_name(elem)
                    akey = fitting_tables.classify_accessory(fam)
                    fan_system_effect.log.info('fan system effect: accessory %s (%s, %s) '
                                               'at %.2f ft', nid, fam, akey, l_ft)
                    if akey in fan_system_effect.DAMPER_KINDS:
                        accessories.append((nid, l_ft, fam))
                nxt = kids
            else:
                return None, l_ft, None, notes      # a terminal or equipment
            if len(nxt) > 1:
                return 'split', l_ft, nid, notes
            if not nxt:
                return None, l_ft, None, notes
            prev, nid = nid, nxt[0]
        fan_system_effect.log.info('fan system effect: nothing within %.1f ft of '
                                   'connection %s (%.1f ft walked)', win, first, l_ft)
        return None, l_ft, None, notes

    def _fse_conn_size(conn, flow, notes, what):
        """(is_round, D or H in.) of an inlet from the unit connector, its
        LARGER side if rectangular (the higher C, and the longer window),
        else the first duct; each fallback is named. (None, None) if nothing
        is readable."""
        if conn is not None:
            if conn.get('round') and conn.get('d_in'):
                return True, conn['d_in']
            if conn.get('w_in') and conn.get('h_in'):
                return False, max(conn['w_in'], conn['h_in'])
        d = _round_d_in(flow)
        if d:
            notes.append(('note', 'NOTE: unit inlet unreadable; %s D taken from the '
                          'first duct (%.0f in.)' % (what, d)))
            return True, d
        if flow.area_ft2:
            side = math.sqrt(flow.area_ft2) * 12.0
            notes.append(('note', 'NOTE: unit inlet unreadable; %s H taken as the side '
                          'of a square of the first duct\'s area (%.0f in.)' % (what, side)))
            return False, side
        return None, None

    def _fse_window_size(conn, flow, notes):
        """(is_round, size in.) the inlet walks REACH out on: every side of
        the unit opening and the first duct, so no elbow is missed. A search
        reach only, never a judging size and never printed (Colin,
        2026-11-09): each elbow is judged on its own H or D (_fse_elbow_h).
        (None, None) if nothing is readable."""
        is_round, size = _fse_conn_size(conn, flow, notes, 'window')
        if size is None:
            return None, None
        dc = _duct_conn(flow)
        d_duct = None
        if dc is not None:
            if dc.get('d_in'):
                d_duct = dc['d_in']
            elif dc.get('w_in') and dc.get('h_in'):
                d_duct = max(dc['w_in'], dc['h_in'])
        fan_system_effect.log.info('fan system effect: 5 D size, opening %.1f in., '
                                   'first duct %s in.', size,
                                   '%.1f' % d_duct if d_duct else 'NOT FOUND')
        if d_duct and d_duct > size:
            return bool(dc.get('d_in')), d_duct
        return is_round, size

    def _fse_elbow_h(eid):
        """(is_round, size in.) an inlet elbow is judged on (Colin,
        2026-11-09): a round elbow's D, a rect elbow's H, its side in the
        plane of its turn (hvac_graph.elbow_geometry). No floor from the
        opening or the first duct. (None, None) when unreadable."""
        g = hvac_graph.elbow_geometry(all_nodes.get(eid)) if eid is not None else None
        if g is not None and g.get('round') and g.get('d_in'):
            return True, float(g['d_in'])
        if g is not None and not g.get('round') and g.get('w_in'):
            fan_system_effect.log.info('inlet H: elbow %s, %.1f in. (turn plane)',
                                       eid, g['w_in'])
            return False, float(g['w_in'])
        fan_system_effect.log.info('inlet H: elbow %s NOT FOUND', eid)
        return None, None

    def _inlet_fallback_size(conn, flow):
        """(is_round, size in., long, short) when the inlet elbow is
        unreadable or none is found: the first duct's D or Height, else the
        unit inlet's D or Height, each named. (None, None, None, None) when
        nothing is readable."""
        dc = _duct_conn(flow)
        for c, what in ((dc, 'first duct'), (conn, 'inlet collar')):
            if c is None:
                continue
            if c.get('round') and c.get('d_in'):
                lab = ('the first duct, %.0f in. round' % c['d_in'] if what == 'first duct'
                       else 'the %.0f in. round inlet collar' % c['d_in'])
                return True, float(c['d_in']), lab, 'the ' + what
            if c.get('w_in') and c.get('h_in'):
                lab = ("the first duct's height, %.0fx%.0f" % (c['w_in'], c['h_in'])
                       if what == 'first duct' else
                       "the %.0fx%.0f inlet collar's height" % (c['w_in'], c['h_in']))
                return False, float(c['h_in']), lab, 'the ' + what
        return None, None, None, None

    def _inlet_size_source(eid, is_round):
        """(long, short) naming the elbow an inlet H or D was read from."""
        if is_round:
            g = hvac_graph.elbow_geometry(all_nodes.get(eid))
            return ('the %.0f in. round elbow' % g['d_in'], 'the elbow')
        return ("the elbow's side in the plane of its turn", 'the elbow')

    def _fse_inlet_size(elem, conn, flow, notes, did=None, l_ft=0.0, where='inlet'):
        """(is_round, D or H in., r in. or None) for the inlet elbow. H is
        the side in the plane of the turn (hvac_graph.elbow_geometry swaps
        Width and Height to put it there). Unreadable elbow: the first
        duct's D or Height, else the unit inlet's (_inlet_fallback_size),
        radius unknown.
        A readable rect mitered elbow (family name, as _ashrae_fitting reads
        it) has no radius and is read on the ER7-1 r/H 1.0 row (Colin,
        2026-11-02), not the conservative unreadable radius row."""
        g = hvac_graph.elbow_geometry(elem) if elem is not None else None
        if g is not None and g.get('round') and g.get('d_in'):
            return True, g['d_in'], g.get('r_in')
        if g is not None and not g.get('round') and g.get('w_in'):
            # ER7-1 is printed for square elbows only (figure, p. 34.60): a
            # non-square one is read on H and stated as an ASSUMPTION line
            # in the report (Colin, 2026-10-29).
            if g.get('h_in') and abs(g['w_in'] - g['h_in']) > 0.01:
                notes.append(('assume', fan_system_effect.ER7_1_SQUARE_ASSUMPTION))
            r_in = g.get('r_in')
            fam = (hvac_graph.fitting_family_name(elem) or '').lower()
            if not r_in and 'miter' in fam and 'radius' not in fam:
                r_in = fan_system_effect.MITERED_R_OVER_H * g['w_in']
                notes.append(('note', 'NOTE: elbow %s, %.1f ft from the %s, is mitered: '
                              'ER7-1 read on the r/H %.1f row.'
                              % (did, float(l_ft), where,
                                 fan_system_effect.MITERED_R_OVER_H)))
            return False, g['w_in'], r_in
        is_round, size, src_l, _src_s = _inlet_fallback_size(conn, flow)
        if size is not None:
            notes.append(('note', 'NOTE: inlet elbow geometry unreadable; %s %.0f in. '
                          'taken from %s, radius unknown' % (
                              'D' if is_round else 'H', size, src_l)))
        return is_round, size, None

    def _tap_inlet_size(tid):
        """(is_round, D or H in., r in. or None) for a tap priced as an inlet
        elbow (Colin, 2026-11-11): read on the duct past the tap, its LARGER
        side if rectangular, on the mitered r/H row (ER7-1); round on ED7-2
        with the radius unknown. (None, None, None) when unreadable."""
        dc = _duct_conn(_first_duct(tid))
        if dc is None:
            return None, None, None
        if dc.get('d_in'):
            return True, dc['d_in'], None
        if dc.get('w_in') and dc.get('h_in'):
            h = max(dc['w_in'], dc['h_in'])
            return False, h, fan_system_effect.MITERED_R_OVER_H * h
        return None, None, None

    def _duct_conn(dr):
        """A picked duct's own section in the shape of
        hvac_graph.equipment_connection_geometry, for a duct pick's fan
        system effect (Colin, 2026-11-02), or None when its size is
        unreadable (logged)."""
        if dr is None or not dr.area_ft2:
            fan_system_effect.log.info('fan system effect: picked duct size NOT FOUND')
            return None
        d = _round_d_in(dr)
        if d:
            return {'round': True, 'area_ft2': dr.area_ft2, 'd_in': d,
                    'w_in': None, 'h_in': None}
        wh = re.match(r'\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)',
                      getattr(dr, 'size', '') or '')
        if wh is None:
            fan_system_effect.log.info('fan system effect: duct %s size %r unreadable, '
                                       'NOT FOUND', dr.element_id, getattr(dr, 'size', None))
            return {'round': False, 'area_ft2': dr.area_ft2, 'd_in': None,
                    'w_in': None, 'h_in': None}
        return {'round': False, 'area_ft2': dr.area_ft2, 'd_in': None,
                'w_in': float(wh.group(1)), 'h_in': float(wh.group(2))}

    def _fse_connection(root_id, first, tag=False, pick=None, direction=None,
                        cfm_in=None):
        """The fan system effect at one unit connection, or None when it is
        not a fan connection this prices (outside air, transfer).

        tag: the unit has more than one connection of this class, so the
        report names this one by its first element.
        pick: (prev, seen) for a duct pick, where first is the picked duct
        itself and its unit side end is taken as the unit connection: its
        own size and cfm are Ao and the flow, and the window walk starts at
        that end (_fse_first_disruption).
        direction: 'in' or 'out' for an inline exhaust fan connection, from
        its terminals (_inline_sides); cfm_in: the airflow for a discharge,
        which carries no terminal of its own.
        Returns {'loss', 'item' (index run item or None), 'where', 'notes',
        'skip': node ids whose own fitting price the system effect replaces
        (an inlet elbow: ED7-2 / ER7-1 include it), 'gray': elbow ids whose
        loss is HIGH_LOSS_INWC or more, 'red': the two elbow ids that turn
        in different planes within 6 Dh (swirl, not priced) and an elbow
        and the straight duct before an elbow under the straight duct line,
        'red_ducts': {duct id: (table reason, required)} for those ducts,
        'boost': {} (the fan velocity transition boost, retired 2026-11-04),
        'group': the straight duct, transition and first elbow ids charged
        once with the fan system effect (_fse_group_max), 'assumptions':
        report ASSUMPTION lines}."""
        rec = {'loss': 0.0, 'item': None, 'where': 'unit connection', 'notes': [],
               'skip': set(), 'gray': set(), 'red': set(), 'boost': {},
               'red_ducts': {}, 'assumptions': [], 'group': set(),
               # Every fan connection prints its own take-off row, outlet and
               # inlet apart (Colin, 2026-11-05): 'none' is the reason a
               # connection carries no fan system effect, shown at 0.000.
               'where_item': None, 'none': None}
        if _kind(first) == 'equipment':
            rec['notes'].append(('note', 'NOTE: unit connection to element %s not '
                                 'checked for fan system effect: it feeds equipment '
                                 'directly.' % first))
            return rec
        side = _side_below(first)
        if side is None:
            rec['notes'].append(('note', 'NOTE: unit connection to element %s not '
                                 'checked for fan system effect: no duct found past '
                                 'it.' % first))
            return rec
        # The role hook (Colin, 2026-10-29): a known connector direction would
        # go in here for exhaust fans, where both sides are Exhaust Air. A
        # unit's role is its classification today, so None is passed.
        role_here = fan_system_effect.connection_role(
            side, 'out' if root_id in discharge_roots else direction)
        if role_here is None:
            fan_system_effect.log.info('fan system effect: connection %s of %s is %r, '
                                       'not checked', first, root_id, side)
            return None
        where_item = fan_system_effect.where_text(side, role_here)
        where = where_item + (' at element %s' % first if tag else '')
        rec['where'] = where
        rec['where_item'] = where_item
        notes = rec['notes']
        # A roof curb (Colin, 2026-11-09): never modeled, counted as straight
        # duct at a unit's own connections, so every L below starts at its
        # height. Not at an exhaust fan, a VAV box root or a duct pick.
        curb_ft = 0.0
        if (pick is None and root_id not in inline_fan_roots
                and root_id not in discharge_roots and root_id not in box_ids):
            curb_ft = float(fse_set.get('curb_in') or 0.0) / 12.0
        rec['curb_ft'] = curb_ft
        # One total length with the curb in it (Colin, 2026-11-09).
        where_c = fan_system_effect.curb_where(where, curb_ft)
        flow = _first_duct(first)
        cfm = cfm_in if cfm_in else _connection_cfm(first)
        if flow is None or cfm <= 0 or not flow.area_ft2:
            notes.append(('note', 'NOTE: fan system effect at the %s not checked, no '
                          'airflow on its first duct' % where))
            rec['none'] = 'not checked, no airflow on its first duct'
            return rec
        conn = _duct_conn(flow) if pick is not None else _unit_conn(root_id, first)
        if conn is not None and conn.get('area_ft2'):
            area = conn['area_ft2']
        else:
            area = flow.area_ft2
            notes.append(('note', 'NOTE: unit %s connection size unreadable; the first '
                          'duct (%s) is used for Ao and the velocity pressure'
                          % (where, getattr(flow, 'size', '?'))))
        # Velocity basis (Colin, 2026-11-03): the larger of the default fan
        # velocity and the connection's, or the connection's alone. It sets
        # Pv, the outlet Le and the transitions passed below. Outlets, and
        # the inlet of an inline or roof exhaust fan, whose duct right at the
        # inlet runs at about the fan's speed (Colin, 2026-11-09). A cabinet
        # unit's inlet (RTU, AHU, furnace, fan coil: filter and coil before
        # the fan) is priced at its own connection velocity, cfm over the
        # connection area (Colin, 2026-11-04).
        v_conn = cfm / area
        if (role_here == fan_system_effect.OUTLET
                or root_id in inline_fan_roots):
            v, fan_gov = fan_system_effect.design_velocity(fse_set, v_conn)
        else:
            v, fan_gov = v_conn, False
        # The velocity basis is said in the run NOTE, only on a run that
        # carries a fan system effect (Colin, 2026-11-11). A cabinet unit
        # inlet says why it keeps the connection velocity (the fan sits past
        # the filter and coil); not on the 'connection' basis, where every
        # connection uses it, and not at a VAV box root, which has no fan.
        vel_clause = fan_system_effect.velocity_clause(
            v, v_conn, fan_gov,
            cabinet=(not fan_gov and role_here != fan_system_effect.OUTLET
                     and root_id not in inline_fan_roots and root_id not in box_ids
                     and (fse_set or {}).get('velocity_basis',
                                             fan_system_effect.DEFAULT_VELOCITY_BASIS)
                     == 'default'))
        # The inlet search reach (internal, never printed): every side of
        # the opening and the first duct. An outlet already has its fallback
        # NOTE above, so its size notes are not repeated.
        is_round_c, size_c = _fse_window_size(
            conn, flow, notes if role_here == fan_system_effect.INLET else [])
        if role_here == fan_system_effect.OUTLET:
            win, _eq = fan_system_effect.effective_duct_length_ft(area, v)
            win_name = 'Le'
            win_short = 'Le'
        else:
            if size_c is None:
                notes.append(('warn', 'WARNING: %s size unreadable; fan system effect '
                              'not checked' % where))
                rec['none'] = 'not checked, connection size unreadable'
                return rec
            # The inlet window is 5 H (5 D) of the FIRST elbow, its own H
            # or D (Colin, 2026-11-09): found on the search reach, past taps.
            # No elbow, or an unreadable one: the first duct's D or Height.
            reach = fan_system_effect.inlet_window_ft(is_round_c, size_c)
            elbows_w = []
            if pick is not None:
                _fse_first_disruption(root_id, first, reach, where, prev=pick[0],
                                      seen=set(pick[1]), collect=elbows_w,
                                      through_taps=True, l0=curb_ft)
            else:
                _fse_first_disruption(root_id, first, reach, where, collect=elbows_w,
                                      through_taps=True, l0=curb_ft)
            first_w = min(elbows_w, key=lambda x: x[1])[0] if elbows_w else None
            round_w, h_w = _fse_elbow_h(first_w)
            if h_w is None:
                round_w, h_w, _l, _s = _inlet_fallback_size(conn, flow)
            if h_w is None:
                round_w, h_w = is_round_c, size_c
            win = fan_system_effect.inlet_window_ft(round_w, h_w)
            win_short = '5 %s' % ('D' if round_w else 'H')
            win_name = 'the inlet window, ' + win_short
            fan_system_effect.log.info('inlet window: %s, first elbow %s, %s %.1f in., '
                                       '%.2f ft (search reach %.2f ft)', where, first_w,
                                       'D' if round_w else 'H', h_w, win, reach)
        # The straight duct check (Colin, 2026-11-03, FAN_SYSTEM_EFFECT_PLAN.md
        # section C): an elbow under 0.5 Le of an outlet, or under 2 D of an
        # inlet, is RED with the straight duct before it, still priced below.
        # Elbows only, every one inside the line, walked past taps on the
        # main and past the other elbows, up to a split or another fitting.
        # A cabinet unit inlet (RTU, AHU, furnace, fan coil) also draws a
        # WARNING for the first elbow under 3 D and past the 2 D RED (Colin,
        # 2026-11-09): its connection velocity pricing holds only with 3 D
        # of straight duct. The walk reaches out to 3 D for it.
        cabinet_in = (role_here != fan_system_effect.OUTLET
                      and root_id not in inline_fan_roots and root_id not in box_ids)
        if role_here == fan_system_effect.OUTLET or size_c is not None:
            if role_here == fan_system_effect.OUTLET:
                red_ft = fan_system_effect.outlet_red_ft(win)
            else:
                red_ft = fan_system_effect.inlet_red_ft(size_c)
            walk_ft = (max(red_ft, fan_system_effect.cabinet_warn_ft(size_c))
                       if cabinet_in else red_ft)
            elbows_r, ducts_r = [], []
            if pick is not None:
                _fse_first_disruption(root_id, first, walk_ft, where, prev=pick[0],
                                      seen=set(pick[1]), collect=elbows_r,
                                      through_taps=True, ducts=ducts_r, l0=curb_ft)
            else:
                _fse_first_disruption(root_id, first, walk_ft, where, collect=elbows_r,
                                      through_taps=True, ducts=ducts_r, l0=curb_ft)
            first_elbow = min(elbows_r, key=lambda x: x[1])[0] if elbows_r else None
            fan_system_effect.log.info('straight duct check: %s, RED line %.2f ft, '
                                       'elbows %r, ducts %r', where, red_ft, elbows_r,
                                       ducts_r)
            for e_r, l_r in elbows_r:
                if role_here == fan_system_effect.OUTLET:
                    notes.append(('red', fan_system_effect.le_text(where, e_r, l_r, win,
                                                                   _eq, curb_ft)))
                    need = fan_system_effect.outlet_red_ft(win)
                else:
                    # The walk reached out on the search reach; each elbow
                    # is judged on its own H (rect, the side in the plane of
                    # its turn) or D (round), no floor (Colin, 2026-11-09).
                    # Unreadable: the first duct's D or Height, named.
                    round_e, size_e = _fse_elbow_h(e_r)
                    if size_e:
                        src_l, src_s = _inlet_size_source(e_r, round_e)
                    else:
                        round_e, size_e, src_l, src_s = _inlet_fallback_size(conn, flow)
                    if not size_e:
                        notes.append(('warn', 'WARNING: elbow %s, %s: no size could be '
                                      'read; straight duct not checked.' % (
                                          e_r, fan_system_effect.from_text(where, l_r,
                                                                           curb_ft))))
                        continue
                    dim_e = 'D' if round_e else 'H'
                    if l_r >= fan_system_effect.inlet_red_ft(size_e) - 1e-9:
                        fan_system_effect.log.info(
                            'straight duct check: elbow %s at %.2f ft is past 2 %s '
                            '(%s %.1f in.), not RED', e_r, l_r, dim_e, dim_e, size_e)
                        if (cabinet_in and e_r == first_elbow
                                and l_r < fan_system_effect.cabinet_warn_ft(size_e) - 1e-9):
                            notes.append(('warn', fan_system_effect.cabinet_warn_text(
                                where, e_r, l_r, size_e, round_e, v_conn, curb_ft)))
                            rec['warn3'] = (e_r, dim_e)
                        continue
                    notes.append(('red', fan_system_effect.inlet_red_text(
                        where, e_r, l_r, size_e, round_e, src_l, src_s, curb_ft)))
                    need = fan_system_effect.inlet_red_ft(size_e)
                rec['red'].add(e_r)
                for d_r, at in ducts_r:
                    if at < l_r - 1e-9 and d_r not in rec['red_ducts']:
                        rec['red'].add(d_r)
                        rec['red_ducts'][d_r] = (
                            fan_system_effect.straight_duct_reason(
                                e_r, l_r, role_here,
                                'D' if role_here == fan_system_effect.OUTLET else dim_e,
                                curb_ft),
                            ('Straight duct %.1f ft incl. curb' % fan_system_effect._ceil_step(
                                need)) if curb_ft else
                            ('Straight duct %.1f ft' % fan_system_effect._ceil_step(need)))
        passed = {}     # transitions inside the pricing window (win)
        # Two elbows turning in different planes within 6 Dh of the
        # connection: RED, never priced (Colin, 2026-11-02; supply, return
        # and exhaust alike). Walked on its own window, past the first elbow.
        d_h = getattr(flow, 'd_h_in', None)
        if d_h:
            swin = fan_system_effect.swirl_window_ft(d_h)
            elbows = []
            if pick is not None:
                _fse_first_disruption(root_id, first, swin, where, prev=pick[0],
                                      seen=set(pick[1]), collect=elbows, l0=curb_ft)
            else:
                _fse_first_disruption(root_id, first, swin, where, collect=elbows, l0=curb_ft)
            pair = fan_system_effect.swirl_pair(
                [(e, l, hvac_graph.elbow_turn_normal(all_nodes.get(e))) for e, l in elbows])
            fan_system_effect.log.info('swirl check: %s, %d elbows within %.2f ft, pair %r',
                                       where, len(elbows), swin, pair)
            if pair is not None:
                notes.append(('red', fan_system_effect.swirl_text(
                    role_here, where_c,
                    tuple((e, l, nv) for e, l, nv in pair), swin, d_h)))
                rec['red'].update((pair[0][0], pair[1][0]))
            elif len(elbows) >= 2:
                # A double bend (Colin, 2026-11-04): two elbows inside 6 Dh
                # that are not a different-plane pair. Only the first is in
                # the fan system effect, so the real loss is higher. A
                # WARNING, not RED, and nothing extra is priced.
                notes.append(('warn', fan_system_effect.double_elbow_text(
                    where_c, [(e, l, None) for e, l in elbows], swin, d_h)))
        else:
            fan_system_effect.log.info('swirl check: %s first duct Dh NOT FOUND, '
                                       'not checked', where)
        # A damper inside the pricing window (1.0 Le at an outlet, 5 D / 5 H
        # at an inlet): a WARNING only (Colin, 2026-11-04, F1). Walked past
        # taps and elbows to the end of the window; L and the price below
        # are unchanged.
        acc = []
        if pick is not None:
            _fse_first_disruption(root_id, first, win, where, prev=pick[0],
                                  seen=set(pick[1]), collect=[], through_taps=True,
                                  accessories=acc, l0=curb_ft)
        else:
            _fse_first_disruption(root_id, first, win, where, collect=[],
                                  through_taps=True, accessories=acc, l0=curb_ft)
        fan_system_effect.log.info('damper check: %s, window %.2f ft, dampers %r',
                                   where, win, acc)
        for a_id, a_l, a_fam in acc:
            notes.append(('warn', fan_system_effect.damper_window_text(
                where_c, a_id, a_fam, a_l, win, win_name)))
        grp_ducts = []
        taps_in = []        # taps off a continuing main inside the window
        drop_info = {}      # a drop: its duct and taps
        if pick is not None:
            role, l_ft, did, wnotes = _fse_first_disruption(
                root_id, first, win, where, prev=pick[0], seen=set(pick[1]),
                passed=passed, ducts=grp_ducts, l0=curb_ft, tap_log=taps_in,
                drop=drop_info)
        else:
            role, l_ft, did, wnotes = _fse_first_disruption(
                root_id, first, win, where, passed=passed, ducts=grp_ducts, l0=curb_ft,
                tap_log=taps_in, drop=drop_info)
        notes.extend(wnotes)
        # Taps at a supply outlet follow Le only (Colin, 2026-11-11): inside Le
        # priced as an elbow on a run through the tap (NOTE; WARNING under
        # 0.5 Le), past Le nothing. The 3 D outlet tap rule (2026-11-09) is
        # gone; it contradicted Le.
        tap_warned = False
        def _price_at(rec, notes, did, l_ft, kind='elbow'):
            """Price the fan system effect at did, L l_ft from the unit, into
            rec: an elbow, or a tap acting as one (kind 'tap', Colin
            2026-11-11). Returns rec."""
            if role_here == fan_system_effect.OUTLET:
                r = fan_system_effect.outlet_effect(
                    fse_set.get('fan_type'), fse_set.get('blast_ratio'), area, cfm, l_ft,
                    pv_of=hvac_graph.velocity_pressure_inwg, v_fpm=v)
            else:
                enotes = []
                if kind == 'tap':
                    is_round, size, r_in = _tap_inlet_size(did)
                else:
                    is_round, size, r_in = _fse_inlet_size(all_nodes.get(did), conn, flow,
                                                           enotes, did, l_ft, where_c)
                if size is None:
                    notes.append(('warn', 'WARNING: %s %s, %s, but no size '
                                  'could be read; fan system effect not priced'
                                  % (kind, did, fan_system_effect.from_text(where, l_ft, curb_ft))))
                    rec['none'] = '%s %s size unreadable, not priced' % (kind, did)
                    return rec
                r = fan_system_effect.inlet_effect(is_round, size, r_in, area, cfm, l_ft,
                                                   pv_of=hvac_graph.velocity_pressure_inwg,
                                                   v_fpm=v)
                if not r['within']:
                    # Past the table's floor distance for its own size: an
                    # ordinary elbow, priced by the walk like any other.
                    fan_system_effect.log.info('fan system effect: inlet elbow %s at %.2f ft '
                                               'is past %.2f ft, priced as a fitting',
                                               did, l_ft, r['floor_ft'])
                    _lt = fan_system_effect.l_text(l_ft, curb_ft, unit=True)
                    # Said in the run NOTE (Colin, 2026-11-11).
                    rec['none'] = ('%s %s at %s, past %.1f ft, priced as a fitting'
                                   % (kind, did, _lt, r['floor_ft']))
                    return rec
                for _k, _t in enotes:
                    if _k == 'assume':
                        rec['assumptions'].append(_t)
                    else:
                        notes.append((_k, _t))
            r['curb_ft'] = curb_ft     # the basis and GRAY text name it
            loss = r['c'] * hvac_graph.velocity_pressure_inwg(r['v_fpm'])
            # The straight duct, any transition (a curb adapter) and the first
            # elbow are one group, charged once as the larger of the fan system
            # effect and their own losses (Colin, 2026-11-04), outlet and inlet.
            # The walk prices them as usual; each path's end picks the larger.
            rec['group'] = set(d for d, _at in grp_ducts) | set(passed) | set([did])
            # Which one governs is said once the path is known (_fse_group_max,
            # Colin, 2026-11-09), not here.
            # The ratio said in the NOTE and the Duct losses basis is on the same
            # size the table was read at (Colin, 2026-11-09, TB-M10 3652084): an
            # outlet's Le; an inlet's 5 x the elbow's own D or H (the side in the
            # plane of its turn, as ER7-1 / ED7-2 and the RED line read it).
            g_size = None
            if role_here == fan_system_effect.OUTLET:
                g_win, g_short = win, win_short
            else:
                g_win = fan_system_effect.inlet_floor_ft(r['size_in'])
                g_dim = 'D' if r.get('table') == fan_system_effect.INLET_ROUND else 'H'
                g_short = '%g %s' % (fan_system_effect.INLET_FLOOR_D, g_dim)
                # Named, so it reads with the RED line (Colin, 2026-11-09,
                # TB-M12 3720096). An unreadable elbow took the first duct's D or
                # Height (its own NOTE says so).
                g_geo = hvac_graph.elbow_geometry(all_nodes.get(did))
                g_own = g_geo is not None and bool(g_geo.get('d_in') or g_geo.get('w_in'))
                g_size = ('%s %.0f in.' % (g_dim, r['size_in']),
                          'elbow' if g_own else 'first duct')
            rec['gov'] = {'eid': did, 'l_ft': l_ft, 'win': g_win, 'win_short': g_short,
                          'transition': bool(passed), 'transitions': sorted(passed),
                          'size': g_size,
                          'warn3': (rec.get('warn3') or (None, None))[1]
                          if (rec.get('warn3') or (None,))[0] == did else None,
                          'curb': curb_ft, 'kind': kind,
                          # the run NOTE (Colin, 2026-11-11)
                          'table': r.get('table') or '-', 'vel': vel_clause}
            notes.extend(('note', 'NOTE: %s %s, %s: %s' % (
                kind, did, fan_system_effect.from_text(where, l_ft, curb_ft), n)) for n in r['notes'])
            if r['high']:
                # GRAY, not RED (Colin, 2026-10-29): still priced in the total,
                # reported with the straight length that brings the loss under
                # HIGH_LOSS_INWC (0.5 in. wc, Colin 2026-11-02, was C 1.0), and
                # the elbow is drawn gray on the plan.
                notes.append(('gray', fan_system_effect.high_text(r, where, did, kind)))
                rec['gray'].add(did)
            name = 'Fan system effect, ' + where_item
            basis = fan_system_effect.basis_text(r)
            if kind == 'tap':
                basis += ', tap as elbow'
            rec['loss'] = loss
            rec['item'] = ('Fan system effect {}'.format(did), basis, r['v_fpm'], loss,
                           (name, basis, r['c'], None))
            return rec

        def _tap_close(notes, tid, lt):
            """The WARNING for a tap priced as an elbow under the straight
            duct line (0.5 Le outlet, 2 H inlet). Never RED (Colin,
            2026-11-11)."""
            if role_here == fan_system_effect.OUTLET:
                line, best = fan_system_effect.outlet_red_ft(win), win
                line_name = '%.1f Le' % fan_system_effect.OUTLET_RED_RATIO
                best_name = '1.0 Le'
            else:
                _r, size_t, _ri = _tap_inlet_size(tid)
                if not size_t:
                    return
                dim = 'D' if _r else 'H'
                line = fan_system_effect.inlet_red_ft(size_t)
                best = fan_system_effect.inlet_floor_ft(size_t)
                line_name = '%.0f %s (%s %.0f in.)' % (
                    fan_system_effect.INLET_RED_D, dim, dim, size_t)
                best_name = '%.0f %s' % (fan_system_effect.INLET_FLOOR_D, dim)
            if lt < line - 1e-9:
                notes.append(('warn', fan_system_effect.tap_close_text(
                    where, tid, lt, line, line_name, best, best_name, curb_ft)))


        def _sub_rec(rec, via):
            """An empty record for a run leaving through one tap."""
            return {'loss': 0.0, 'item': None, 'where': rec['where'], 'notes': [],
                    'skip': set(), 'gray': set(), 'red': set(), 'boost': {},
                    'red_ducts': {}, 'assumptions': [], 'group': set(),
                    'where_item': where_item, 'none': None, 'curb_ft': curb_ft,
                    'via': dict(via, win_short=win_short, win_ft=win)}

        def _tap_subs(rec):
            """Taps off a continuing main inside the window (Colin,
            2026-11-11): a run that leaves through one turns there, so it
            carries that tap as an elbow at its own L; a run that continues
            on the main keeps rec (the first elbow, or none)."""
            if not taps_in:
                return rec
            # The run that stays on the main names the taps it passes in its
            # run NOTE (Colin, 2026-11-11).
            rec['passed'] = [(tid, lt) for tid, lt, _d in taps_in]
            subs = rec.get('by_tap') or {}
            for tid, lt, _d in taps_in:
                sub = _sub_rec(rec, {'kind': 'tap', 'id': tid, 'l_ft': lt})
                _tap_close(notes, tid, lt)
                _price_at(sub, sub['notes'], tid, lt, 'tap')
                for n in sub['notes']:
                    if n not in notes:
                        notes.append(n)
                rec['gray'] |= sub['gray']
                for _a in sub['assumptions']:
                    if _a not in rec['assumptions']:
                        rec['assumptions'].append(_a)
                subs[tid] = sub
            rec['by_tap'] = subs
            return rec

        if role is None:
            rec['none'] = 'no elbow within %s (%.1f ft)' % (
                win_short, fan_system_effect._ceil_step(win))
            # The straight duct and any transition inside the window still
            # sit on the fan system effect row, at their own losses (Colin,
            # 2026-11-11), as they do when an elbow is priced.
            # What lies inside the window: a duct running past it moves only
            # its share inside the window (friction pro rata on length); the
            # rest stays on its own row.
            _gd, _frac = set(), {}
            for _d, _at in grp_ducts:
                _dr = dr_by_int_id.get(_d)
                _len = float(getattr(_dr, 'length_ft', 0.0) or 0.0)
                if _dr is None or _at >= win - 1e-6:
                    break
                _gd.add(_d)
                if _len > 1e-9 and _at + _len > win + 1e-6:
                    _frac[_d] = (win - _at) / _len
                    break
            _gt = set(_t for _t, _at in passed.items() if _gd and _at < win)
            rec['group'] = _gd | _gt
            rec['group_frac'] = _frac
            rec['win_short'] = win_short
            rec['none_group'] = bool(_gd)
            rec['none_trans'] = bool(_gt)
            # Said in the run NOTE, for the index run only (Colin, 2026-11-11).
            return _tap_subs(rec)
        if role not in ('elbow', 'tap_turn', 'drop'):
            if not tap_warned:
                notes.append(('warn', fan_system_effect.disruption_text(
                    _FSE_KIND_TEXT.get(role, role), did, where_c, l_ft, win,
                    win_name)))
            rec['none'] = '%s %s at %s inside %s, flagged, not priced' % (
                _FSE_KIND_TEXT.get(role, role), did,
                fan_system_effect.l_text(l_ft, curb_ft, unit=True), win_short)
            return _tap_subs(rec)

        if role == 'tap_turn':
            rec['via'] = {'kind': 'tap_turn', 'id': did, 'l_ft': l_ft,
                          'win_short': win_short, 'win_ft': win}
            _tap_close(notes, did, l_ft)
            return _tap_subs(_price_at(rec, notes, did, l_ft, 'tap'))
        if role == 'drop':
            # Each tap off the drop acts as an elbow, per run (Colin,
            # 2026-11-11): a run picks the record of the tap it passes.
            rec['by_tap'] = rec.get('by_tap') or {}
            for tid, lt in drop_info['taps']:
                sub = _sub_rec(rec, {'kind': 'drop', 'id': tid, 'l_ft': lt,
                                     'drop_duct': drop_info['duct']})
                if lt >= win - 1e-9:
                    sub['none'] = 'tap %s at %s, past %s (%.1f ft)' % (
                        tid, fan_system_effect.l_text(lt, curb_ft, unit=True),
                        win_short, fan_system_effect._ceil_step(win))
                else:
                    _tap_close(notes, tid, lt)
                    _price_at(sub, sub['notes'], tid, lt, 'tap')
                for n in sub['notes']:
                    if n not in notes:
                        notes.append(n)
                rec['gray'] |= sub['gray']
                for _a in sub['assumptions']:
                    if _a not in rec['assumptions']:
                        rec['assumptions'].append(_a)
                rec['by_tap'][tid] = sub
                rec['loss'] = max(rec['loss'], sub['loss'])
            rec['none'] = 'drop %s, priced per tap' % drop_info['duct']
            return _tap_subs(rec)
        rec['via'] = {'kind': 'elbow', 'id': did, 'l_ft': l_ft,
                      'win_short': win_short, 'win_ft': win}
        return _tap_subs(_price_at(rec, notes, did, l_ft, 'elbow'))

    def _checks_only(rec):
        """A connection record with the fan system effect pricing taken out,
        the straight duct and swirl RED kept (fan system effect unticked)."""
        rec['loss'] = 0.0
        rec['item'] = None
        rec['by_tap'] = {}
        rec['group'] = set()
        rec['skip'] = set()
        rec['gray'] = set()
        rec['boost'] = {}
        rec['assumptions'] = []
        rec['none'] = None
        rec['none_group'] = False
        rec['notes'] = [n for n in rec['notes'] if n[0] == 'red' or (
            n[0] == 'warn' and fan_system_effect.is_check_warning(n[1]))]
        return rec

    def _fse_root(root_id):
        """({unit connection node: record}, notes about the root itself).
        Unticked fan system effect: the RED checks only (_checks_only)."""
        recs, rnotes = _fse_root_all(root_id)
        if not fse_price:
            for _r in recs.values():
                _checks_only(_r)
            rnotes = []
        return recs, rnotes

    def _fse_root_all(root_id):
        recs, rnotes = {}, []
        root_elem = all_nodes.get(root_id)
        if (duct_rooted_ids and root_id in duct_rooted_ids
                and root_id not in dr_by_int_id):
            fan_system_effect.log.info('fan system effect: picked root %s is not a '
                                       'duct, NOT FOUND', root_id)
        elif duct_rooted_ids and root_id in duct_rooted_ids:
            # No unit on the run: the picked duct's unit side end is taken as
            # the unit outlet or inlet (Colin, 2026-11-02). The unit side
            # stretch behind it (_unit_side_firsts) is not walked; its first
            # node is the end the taps on the picked duct are measured from.
            side_firsts = _unit_side_firsts(root_id)
            prev = side_firsts[0][0] if side_firsts else None
            seen = set()
            for _f, _order in side_firsts:
                seen.update(_order)
            fan_system_effect.log.info('fan system effect: duct pick %s, unit side '
                                       'neighbor %s', root_id,
                                       prev if prev is not None else 'NOT FOUND')
            rec = _fse_connection(root_id, root_id, pick=(prev, seen))
            if rec is not None:
                role_here = fan_system_effect.connection_role(_side_below(root_id), None)
                rnotes.append(('note', 'NOTE: duct pick, fan system effect measured from '
                               'the unit end of duct %s, taken as the unit %s.'
                               % (root_id, role_here or 'connection')))
                recs[root_id] = rec
        elif root_id in box_ids:
            rnotes.append(('note', 'NOTE: fan system effect not priced at a VAV box.'))
        elif root_elem is None or not hvac_graph.is_equipment(root_elem):
            fan_system_effect.log.info('fan system effect: root %s is not equipment, '
                                       'NOT FOUND', root_id)
        else:
            # A unit connector also references its logical Duct System
            # element; that is no connection to price (OLMC, 2026-10-30).
            firsts = []
            for f in all_children.get(root_id, []):
                if _kind(f) is None:
                    fan_system_effect.log.info('fan system effect: node %s under %s '
                                               'is not ductwork, skipped', f, root_id)
                else:
                    firsts.append(f)
            classes = [_side_below(f) if _kind(f) != 'equipment' else None
                       for f in firsts]
            dirs, dis_cfm, _n, _kind_fan = inline_cache.get(root_id, ({}, 0.0, [], None))
            # Only the discharge that is priced gets a fan system effect
            # record, so a second one never shows a loss in no total (review,
            # 2026-11-03).
            outs = [f for f in all_children.get(root_id, []) if dirs.get(f) == 'out']
            for first, cls in zip(firsts, classes):
                tag = cls is not None and classes.count(cls) > 1
                d = dirs.get(first)
                if d == 'out' and first != outs[0]:
                    fan_system_effect.log.info('fan system effect: discharge %s of %s not '
                                               'priced (second discharge)', first, root_id)
                    continue
                rec = _fse_connection(root_id, first, tag, direction=d,
                                      cfm_in=dis_cfm if d == 'out' else None)
                if rec is not None:
                    recs[first] = rec
        return recs, rnotes

    # ── inline exhaust fans (Colin, 2026-11-03) ──────────────────────────────
    # Both connections are Exhaust Air. The one whose ductwork reaches the
    # grilles is the INLET, the other, with no terminals, the OUTLET
    # (discharge). No FlowDirectionType: the terminals decide.
    def _reaches_terminal(first):
        """[terminal ids] below node first (its subtree), walked once."""
        found, todo, seen = [], [first], set()
        while todo:
            n = todo.pop()
            if n in seen:
                continue
            seen.add(n)
            if n in all_terminals:
                found.append(n)
            todo.extend(all_children.get(n, []))
        return found

    def _inline_sides(root_id):
        """({first: 'in' or 'out'}, discharge cfm, notes, kind) for an
        exhaust fan that is not a ceiling fan, from its connections, read
        every run (Colin, 2026-11-03):
          'inline': an inlet (reaches the grilles) and an outlet (no
            grilles). The discharge is priced at the total grille cfm.
          'roof': inlet only, an upblast or downblast roof fan that
            discharges itself: the intake alone, no discharge walked. The
            fan type sets only the fan system effect and the discharge walk;
            the cap and backdraft boxes are charged as ticked (Colin,
            2026-11-03).
          None: no connection reaches the grilles; read on the
            classification, with a NOTE."""
        firsts = [f for f in all_children.get(root_id, []) if _kind(f) is not None]
        terms = dict((f, _reaches_terminal(f)) for f in firsts)
        intake = [f for f in firsts if terms[f]]
        # A connection straight into other equipment is no discharge (review,
        # 2026-11-03).
        equip = [f for f in firsts if not terms[f] and _kind(f) == 'equipment']
        dis = [f for f in firsts if not terms[f] and _kind(f) != 'equipment']
        for f in firsts:
            fan_system_effect.log.info('inline exhaust fan %s: connection %s (%s) reaches '
                                       '%d terminal(s)', root_id, f, _kind(f), len(terms[f]))
        if intake and not dis and equip:
            return {}, 0.0, [('note', 'NOTE: exhaust fan %s, connection %s feeds other '
                              'equipment; fan type not read, no discharge priced.'
                              % (root_id, ', '.join(str(f) for f in equip)))], None
        if not intake:
            # Colin, 2026-11-03: an error, not computed. build_network stops a
            # fan with no terminal at all first (main() reports it as an
            # error for an exhaust fan); this covers the rest.
            return {}, 0.0, [('error', 'ERROR: exhaust fan %s, no connection reaches a '
                              'grille or diffuser; the fan is not computed. Connect '
                              'its intake ductwork to the grilles.' % root_id)], 'error'
        if not dis:
            return (dict((f, 'in') for f in intake), 0.0,
                    [('note', 'NOTE: exhaust fan read as an upblast or downblast roof '
                      'fan (inlet only, at element %s): the worst intake path is priced, '
                      'no discharge walked.' % ', '.join(str(f) for f in intake))],
                    'roof')
        # The air picked up walking back from every grille to the fan, the
        # same post-order sum as every duct (Colin, 2026-11-03).
        cfm = sum([_connection_cfm(f) for f in intake], 0.0)
        dirs = dict((f, 'in') for f in intake)
        dirs.update((f, 'out') for f in dis)
        # Worded as the sum of the terminals (Colin, 2026-11-09).
        return dirs, cfm, [('note', 'NOTE: exhaust fan read as an inline fan: intake at '
                            'element %s, discharge at element %s. The discharge carries '
                            '%.0f cfm, the sum of the exhaust terminal airflows, and is '
                            'priced once on the index run.' % (
                                ', '.join(str(f) for f in intake),
                                ', '.join(str(f) for f in dis), cfm))], 'inline'

    def _discharge_chain(first, cfm, boost=None):
        """The discharge duct and fittings from the fan out, priced at the
        intake cfm: no path ends down there, so they are added to every
        exhaust run of the fan, like the roof cap (Colin, 2026-11-03). The air
        moves with the walk, away from the fan, so fittings are priced on the
        supply tables in walk order. boost: the discharge connection's
        {transition id: velocity factor} (fan velocity basis), as the main
        walk applies it. A flex discharge duct is flagged: exhaust shall be
        rigid (Colin, 2026-11-03).
        Returns {'fric', 'fit', 'n_duct', 'n_fit', 'feet', 'comp_keys',
        'items', 'unpriced'}."""
        out = {'fric': 0.0, 'fit': 0.0, 'n_duct': 0, 'n_fit': 0, 'feet': 0.0,
               'comp_keys': (), 'items': (), 'unpriced': ()}
        boost = boost or {}
        nid, seen = first, set()
        up_fpm, up_area = 0.0, None
        while nid is not None and nid not in seen:
            seen.add(nid)
            elem = all_nodes.get(nid)
            dr = dr_by_int_id.get(nid)
            if dr is not None:
                fpm = cfm / dr.area_ft2 if dr.area_ft2 else 0.0
                d_h = getattr(dr, 'd_h_in', 0.0)
                fr = 0.0
                if d_h > 0 and fpm > 0:
                    fr = (hvac_graph.duct_friction_loss_per_100ft(
                        fpm, d_h, getattr(dr, 'roughness_ft', None))
                          * float(dr.length_ft) / 100.0)
                out['fric'] += fr
                out['n_duct'] += 1
                out['feet'] += float(dr.length_ft)
                up_fpm, up_area = fpm, dr.area_ft2
                out['items'] += (('Discharge duct {}'.format(nid), '{}, {:.1f} ft'.format(
                    getattr(dr, 'size', '?'), float(dr.length_ft)), fpm, fr,
                    ('Discharge duct', 'Darcy-Weisbach', None, float(dr.length_ft))),)
                # Exhaust is never flex: the low pressure collapses it
                # (Colin, 2026-11-03). Flagged, not priced as a flex run.
                if getattr(dr, 'is_flex', False):
                    out['unpriced'] += ('WARNING: flex duct %s on the exhaust discharge; '
                                        'exhaust ductwork shall be rigid' % nid,)
            elif elem is not None and hvac_graph.is_fitting(elem):
                if up_area is None:
                    flow = _first_duct(nid)
                    if flow is not None and flow.area_ft2:
                        up_area = flow.area_ft2
                        up_fpm = cfm / up_area
                _k = boost.get(nid)
                price_fpm = up_fpm * _k if (_k and up_fpm) else up_fpm
                loss, counted, note, down_a, detail = _price_fitting(
                    elem, 'Supply Air', price_fpm, up_area, against_flow=False, loc_id=nid,
                    fb_why='priced on the worksheet on the exhaust discharge, '
                           'which has no 2001 junction lookup')
                if _k and up_fpm:
                    detail = '{}, at fan velocity'.format(detail)
                fam = hvac_graph.fitting_family_name(elem) or '?'
                k = _takeoff_key(elem, fam, detail, counted, note)
                out['fit'] += loss
                out['n_fit'] += counted
                if note:
                    out['unpriced'] += (note,)
                out['items'] += (('Discharge fitting {}'.format(nid),
                                  '{}, {}'.format(fam, detail), up_fpm, loss,
                                  ('Discharge ' + k[0][:1].lower() + k[0][1:],) + k[1:]),)
                if down_a and up_area and up_fpm:
                    up_fpm, up_area = up_fpm * up_area / down_a, down_a
            elif elem is not None and hvac_graph.is_accessory(elem):
                fam = hvac_graph.fitting_family_name(elem)
                akey = fitting_tables.classify_accessory(fam)
                if akey is not None:
                    out['comp_keys'] += (akey,)
                    out['items'] += (('Discharge accessory {}'.format(nid), fam or '?', None,
                                      fitting_tables.component_of(akey, comp_values),
                                      (_COMP_LABEL.get(akey, akey), 'known drop', None,
                                       None)),)
                else:
                    out['unpriced'] += ((fam or '?') + ': accessory on the discharge, '
                                        'not priced',)
            else:
                break
            kids = [c for c in all_children.get(nid, []) if c not in seen]
            if len(kids) > 1:
                out['unpriced'] += ('discharge splits at element %s: priced along '
                                    'its first leg only' % nid,)
            nid = kids[0] if kids else None
        fan_system_effect.log.info('inline exhaust fan discharge from %s: %d ducts, %d '
                                   'fittings, %.1f ft, friction %.4f, fittings %.4f at '
                                   '%.0f cfm', first, out['n_duct'], out['n_fit'],
                                   out['feet'], out['fric'], out['fit'], cfm)
        return out

    _GROUP_ITEM = re.compile(r'^(?:Discharge )?(duct|fitting) (\d+)$', re.I)

    def _fse_group_max(rec, its):
        """One unit connection's group (rec['group']: the straight duct, any
        transition and the first elbow) against its fan system effect line,
        charged once as the larger (Colin, 2026-11-04).

        its: the take-off items of the path, or of a discharge chain.
        Returns (items, fse charged, friction removed, fittings removed,
        fse item or None, run NOTE facts or None: _run_note). The fan system effect wins a
        tie. Whichever governs, its value is charged on the fan system effect
        row and the group items are listed at 0.000 'in the fan system
        effect' (Colin, 2026-11-09); the row's Table / basis reads 'Fan system
        effect: <table>' or 'Duct losses: straight duct and elbow (<table>),
        L / Le' with C '-'. 'fse charged' is that row's value, so the
        subtotal is the same either way. The run NOTE says which governed."""
        if (rec is not None and rec.get('item') is None and rec.get('none_group')
                and rec.get('where_item') and rec.get('none')):
            return _none_group(rec, its)
        if rec is None or rec.get('item') is None or not rec.get('group'):
            item = (rec or {}).get('item')
            if item is None and rec and rec.get('none') and rec.get('where_item'):
                # No fan system effect here: its own row at 0.000 with the
                # reason, so outlet and inlet always read apart (2026-11-05).
                item = ('Fan system effect (none)', rec['none'], None, 0.0,
                        ('Fan system effect, ' + rec['where_item'],
                         'none: ' + rec['none'], None, None))
            return its, (rec or {}).get('loss', 0.0) or 0.0, 0.0, 0.0, item, None
        g_fric, g_fit, hits = 0.0, 0.0, []
        for i, it in enumerate(its):
            m = _GROUP_ITEM.match(it[0])
            if m is None or int(m.group(2)) not in rec['group']:
                continue
            hits.append(i)
            if m.group(1).lower() == 'duct':
                g_fric += it[3] or 0.0
            else:
                g_fit += it[3] or 0.0
        group = g_fric + g_fit
        hit_items = [its[i] for i in hits]     # before they are listed at 0
        fse = rec['loss']
        fan_system_effect.log.info('fan system effect group: %s, fse %.4f vs duct and '
                                   'fittings %.4f (%d items)', rec.get('where'), fse,
                                   group, len(hits))
        item = rec['item']
        gov = rec.get('gov') or {}
        # What the run NOTE says (_run_note, Colin, 2026-11-11).
        note = None
        if gov:
            note = {'priced': {'table': gov.get('table') or '-', 'vel': gov.get('vel') or '',
                               'fse': fse, 'own': group, 'fse_wins': fse >= group,
                               'transition': gov.get('transition', False),
                               'warn_dim': gov.get('warn3')},
                    'win_ft': gov.get('win'), 'win_short': gov.get('win_short'),
                    'size': gov.get('size')}
        # The group's charge always sits on the fan system effect row, whichever
        # governs (Colin, 2026-11-09); its own rows stay listed at 0.000 with
        # ', in the fan system effect' on their Table / basis, so the take-off
        # lists them under that row and the 0.000 reads as covered, not free. The
        # subtotal is the same either way: the group's friction and fittings
        # come out, the governing value goes in on this row.
        fse_wins = fse >= group
        its = tuple(_in_group(it) if i in hits else it for i, it in enumerate(its))
        if fse_wins:
            charged = fse
            name, basis, c, ft = item[4]
            row = (name, 'Fan system effect: ' + basis, c, ft)
            detail = item[1]
        else:
            charged = group
            name, _basis, _c, ft = item[4]
            row = (name, _own_loss_basis(rec, hit_items), None, ft)
            detail = '{}; {} at own losses ({:.4f} in. wc)'.format(
                item[1], _group_words(gov), group)
        item = (item[0], detail, item[2], charged, row)
        return its, charged, g_fric, g_fit, item, note

    def _none_group(rec, its):
        """No fan system effect at a connection, but the straight duct (and any
        transition) inside the window is listed on the fan system effect row at
        its own losses, its own rows at 0.000 'in the fan system effect'
        (Colin, 2026-11-11). Same return as _fse_group_max."""
        g_fric, g_fit, hits, parts = 0.0, 0.0, [], {}
        frac = rec.get('group_frac') or {}
        for i, it in enumerate(its):
            m = _GROUP_ITEM.match(it[0])
            if m is None or int(m.group(2)) not in rec['group']:
                continue
            f = frac.get(int(m.group(2)))
            if f is not None and m.group(1).lower() == 'duct':
                share = (it[3] or 0.0) * f
                g_fric += share
                parts[i] = share
                continue
            hits.append(i)
            if m.group(1).lower() == 'duct':
                g_fric += it[3] or 0.0
            else:
                g_fit += it[3] or 0.0
        words = ('straight duct and transition' if rec.get('none_trans') and g_fit
                 else 'straight duct')
        name = 'Fan system effect, ' + rec['where_item']
        if not hits and not parts:
            return its, 0.0, 0.0, 0.0, ('Fan system effect (none)', rec['none'], None, 0.0,
                                        (name, 'none: ' + rec['none'], None, None)), None
        group = g_fric + g_fit
        its = tuple(_in_group(it) if i in hits
                    else (it[0], it[1], it[2], (it[3] or 0.0) - parts[i], it[4]) if i in parts
                    else it for i, it in enumerate(its))
        if parts:
            words += ' inside %s' % rec.get('win_short', 'the window')
        basis = 'none: %s; duct losses: %s' % (rec['none'], words)
        item = ('Fan system effect (none)', '%s at own losses (%.4f in. wc)' % (basis, group),
                None, group, (name, basis, None, None))
        return its, group, g_fric, g_fit, item, {'own': (words, group)}

    def _in_group(it):
        """A group item listed at 0.000, its basis marked as in the fan system
        effect row (the take-off lists it under that row)."""
        t4 = it[4]
        if t4 is not None:
            t4 = (t4[0], '{}, in the fan system effect'.format(t4[1] or '-'), t4[2], t4[3])
        return (it[0], '{}, in the fan system effect'.format(it[1]), it[2], 0.0, t4)

    def _group_words(gov):
        kind = (gov or {}).get('kind') or 'elbow'
        return ('straight duct, transition and %s' % kind if (gov or {}).get('transition')
                else 'straight duct and %s' % kind)

    def _table_of(it):
        """The ASHRAE table id of a priced fitting item ('CR3-1'), else its
        own Table / basis ('RJA worksheet' for a fallback), or None."""
        t4 = (it[4] or (None, None))[1]
        for text in (t4, it[1]):
            m = _TABLE_ID.match(text or '')
            if m:
                return m.group(1)
        return t4 or None

    def _own_loss_basis(rec, group_items):
        """The FSE row's Table / basis when the group's own losses govern
        (Colin, 2026-11-09): 'Duct losses: straight duct and elbow (CR3-1),
        L 3.4 / Le 3.5 ft', the transition and its table named when one is in
        the group."""
        gov = rec.get('gov') or {}
        by_id = {}
        for it in group_items:
            m = _GROUP_ITEM.match(it[0])
            if m is not None:
                by_id[int(m.group(2))] = it
        tables = [_table_of(by_id[t]) for t in gov.get('transitions', ()) if t in by_id]
        elbow = by_id.get(gov.get('eid'))
        tables.append(_table_of(elbow) if elbow is not None else None)
        tables = [t for t in tables if t]
        text = 'Duct losses: ' + _group_words(gov)
        if tables:
            text += ' (%s)' % ', '.join(tables)
        if gov.get('win'):
            text += ', L %s / %s %.1f ft' % (
                fan_system_effect.l_text(gov['l_ft'], gov.get('curb') or 0.0),
                gov.get('win_short', 'Le'), float(gov['win']))
            if gov.get('size'):
                # inlet: the size the window was read at (Colin, 2026-11-09)
                text += ' (%s, the %s)' % gov['size']
        return text

    best = {}              # root_id -> {sys_class: {...}}
    runs = {}              # root_id -> {sys_class: {terminal id: {...}}} (keep_runs)
    box_reach = {}         # root_id -> {box id: worst MP loss to the box inlet}
    truncated_roots = set()  # PER ROOT: one unit truncating must not label the
                             # others, now that results print per equipment.

    # Bound on total path steps. A simple-path search is exponential in the
    # worst case, and while duct networks are near-trees, a badly modelled one
    # with many merge points could in principle blow up. Hitting this is
    # reported, never silently swallowed.
    MAX_PATH_STEPS = 200000

    for root_id in all_root_ids:
        best_here = {}     # this equipment's own per-system-class winners
        runs_here = {}     # sys_class -> {terminal id: its own worst run} (keep_runs)
        boxes_here = {}    # box id -> worst MP loss reaching its inlet
        # Fan system effect per unit connection, charged to the paths that
        # leave the unit through that connection.
        # An inline exhaust fan: its sides from the terminals, and the
        # discharge priced once, added to each of its exhaust runs below.
        dis_rec = None
        inline_notes = []
        if root_id in inline_fan_roots:
            inline_cache[root_id] = _inline_sides(root_id)
            dirs_here, dis_cfm_here, inline_notes, fan_kind = inline_cache[root_id]
            fan_kinds[root_id] = fan_kind
        fse_recs, fse_root_notes = _fse_root(root_id)
        if root_id in inline_fan_roots:
            # After the fan system effect, so the discharge transitions take
            # the fan velocity its record sets.
            outs = [f for f in all_children.get(root_id, []) if dirs_here.get(f) == 'out']
            if outs and dis_cfm_here > 0:
                dis_rec = _discharge_chain(outs[0], dis_cfm_here,
                                           (fse_recs.get(outs[0]) or {}).get('boost'))
                dis_rec['first'] = outs[0]
                if len(outs) > 1:
                    inline_notes = inline_notes + [(
                        'note', 'NOTE: inline exhaust fan with %d discharge connections; '
                        'the one at element %s is priced.' % (len(outs), outs[0]))]
        fse_root_notes = list(inline_notes) + list(fse_root_notes)
        # A duct pick's record is keyed by the picked duct itself, on every path.
        fse_firsts = [c for c in list(all_children.get(root_id, [])) + [root_id]
                      if c in fse_recs]
        fse_skip = set()
        fse_gray = set()
        fse_red = set()
        fse_boost = {}     # transition id -> velocity factor (fan velocity basis)
        fse_red_ducts = {}  # straight duct id -> (reason, required), RED with its elbow
        fse_unit_notes = list(fse_root_notes)
        fse_unit_assume = []
        # Every connection's notes, not only the one the worst path leaves
        # through (review, 2026-10-29): a second outlet's GRAY is still shown.
        for _f in fse_firsts:
            _r = fse_recs[_f]
            fse_skip |= _r['skip']
            fse_gray |= _r['gray']
            fse_red |= _r['red']
            fse_boost.update(_r.get('boost') or {})
            for _d, _v in (_r.get('red_ducts') or {}).items():
                fse_red_ducts.setdefault(_d, _v)
            fse_unit_notes.extend(_r['notes'])
            for _a in _r['assumptions']:
                if _a not in fse_unit_assume:
                    fse_unit_assume.append(_a)
        # (node, friction, fitting loss, ducts, fittings, taps bypassed, feet,
        #  upstream FPM, upstream area ft2, sys_class, unpriced notes,
        #  nodes already on THIS path)
        # A duct-rooted pick starts with the unit-side stretch already priced.
        pre = (0.0, 0.0, 0, 0, 0.0, ())
        if duct_rooted_ids and root_id in duct_rooted_ids:
            pre = _unit_side_chain(root_id)
        # pre = (friction, fitting loss, fittings, ducts, feet, unpriced)
        # Last field: was the previous duct on this path flex? One flex bend is
        # charged when a path ENTERS flex, so a flex run split into several
        # Revit segments still gets one bend, not one per segment.
        # Last field: the run item by item, (item, detail, fpm, loss in. wc),
        # printed under the summary so every number in the total can be traced.
        # Last field: a balancing damper drawn on THIS branch, below the last
        # split or take-off the path passed (Colin, 2026-11-02: dampers on
        # mains are fire/smoke dampers, so only a branch damper replaces the
        # assumed one).
        pre_items = ()
        if pre[1] or pre[0]:
            pre_items = (('Unit side stretch', '{} fitting(s), {} duct(s), {:.0f} ft'.format(
                pre[2], pre[3], pre[4]), None, pre[0] + pre[1],
                ('Unit side stretch (picked duct to unit)', 'mixed', None, pre[4])),)
        stack = [(root_id, pre[0], pre[1], (), pre[3], pre[2], 0, pre[4], 0.0,
                  0.0, None, pre[5], frozenset([root_id]), False, pre_items, False)]
        steps = 0
        while stack:
            steps = steps + 1
            if steps > MAX_PATH_STEPS:
                hvac_graph.log.warning('_critical_path_loss: path search truncated at %d '
                            'steps from root %s', MAX_PATH_STEPS, root_id)
                truncated_roots.add(root_id)
                break
            (nid, fric, fit, comp_keys, dcount, fcount, tcount, dlen,
             up_fpm, up_area, sys_class, unpriced, path, in_flex,
             items, br_damper) = stack.pop()

            elem = all_nodes.get(nid)
            dr   = dr_by_int_id.get(nid)
            jn_here = False

            if dr is not None:
                fric      = fric + dr.friction_loss_inwc
                dlen      = dlen + dr.length_ft
                dcount    = dcount + 1
                up_fpm    = dr.fpm
                up_area   = dr.area_ft2
                sys_class = getattr(dr, 'flow_class', None) or dr.sys_class
                items     = items + ((
                    'Duct {}'.format(nid),
                    '{}, {:.1f} ft'.format(getattr(dr, 'size', '?'),
                                           float(dr.length_ft)),
                    dr.fpm, dr.friction_loss_inwc,
                    ('Flex duct' if getattr(dr, 'is_flex', False) else 'Straight duct',
                     'Darcy-Weisbach', None, float(dr.length_ft))),)
                # Straight flex friction alone covers far less than one real
                # bend (ASHRAE Duct Systems Design Guide Ch. 2: flex bend C 0.9 to
                # 3.3). Colin, 2026-10-21: charge one bend per continuous flex
                # run at the duct's own velocity.
                is_flex_here = bool(getattr(dr, 'is_flex', False))
                if is_flex_here and not in_flex and dr.fpm:
                    bend   = (fitting_tables.flex_bend_c(c_values) *
                              hvac_graph.velocity_pressure_inwg(dr.fpm))
                    fit    = fit + bend
                    fcount = fcount + 1
                    items  = items + ((
                        'Flex bend', 'C {:.2f}'.format(
                            fitting_tables.flex_bend_c(c_values)),
                        dr.fpm, bend,
                        ('Flex bend, one per flex run', 'DSDG Ch. 2',
                         fitting_tables.flex_bend_c(c_values), None)),)
                in_flex = is_flex_here

            elif elem is not None and nid in fse_skip:
                # An inlet elbow the fan system effect priced: ED7-2 / ER7-1
                # include the elbow's own loss (Colin, 2026-10-29), so it is
                # one line in the take-off, the fan system effect one.
                pass

            elif (elem is not None and hvac_graph.is_fitting(elem)
                  and nid not in fse_boost and _junction(nid) is not None):
                # A tee, cross or wye whose legs read: priced in the child
                # loop below, on the leg the run leaves through (Colin,
                # 2026-11-11).
                fcount = fcount + 1
                jn_here = True

            elif elem is not None and hvac_graph.is_fitting(elem):
                # A transition between the unit and the first disruption,
                # under the fan velocity basis: priced with the air at the fan
                # velocity (Colin, 2026-11-03). The walk carries on at the
                # model's own velocity past it.
                _k = fse_boost.get(nid)
                price_fpm = up_fpm * _k if (_k and up_fpm) else up_fpm
                loss, counted, note, _down_a, detail = _price_fitting(
                    elem, sys_class, price_fpm, up_area,
                    against_flow=fitting_tables.flows_toward_unit(sys_class),
                    nid=nid)
                if _k and up_fpm:
                    detail = '{}, at fan velocity'.format(detail)
                fit    = fit + loss
                fcount = fcount + counted
                if note:
                    unpriced = unpriced + (note,)
                _fam = hvac_graph.fitting_family_name(elem) or '?'
                if fitting_tables.classify_fitting(_fam) == 'takeoff':
                    br_damper = False      # a new branch starts at the take-off
                items = items + ((
                    'Fitting {}'.format(nid),
                    '{}, {}'.format(_fam, detail),
                    up_fpm, loss,
                    _takeoff_key(elem, _fam, detail, counted, note)),)
                # Past a transition the air moves at its far end's velocity, so
                # the next fitting (or transition) is priced and oriented there.
                # Same airflow, new area.
                if _down_a and up_area and up_fpm:
                    up_fpm = up_fpm * up_area / _down_a
                    up_area = _down_a

            elif elem is not None and hvac_graph.is_accessory(elem):
                # Accessories are components, not fittings: the drop is a
                # cutsheet number the engineer enters and the model supplies only
                # the COUNT. Each kind gets its OWN value - a fire or smoke damper
                # must never inherit the balancing damper's, and anything
                # matching none of the kinds is reported as uncounted rather
                # than priced wrong.
                fam  = hvac_graph.fitting_family_name(elem)
                akey = fitting_tables.classify_accessory(fam)
                if akey is not None:
                    comp_keys = comp_keys + (akey,)
                    if akey == 'damper':
                        br_damper = True
                    items = items + ((
                        'Accessory {}'.format(nid), fam or '?', None,
                        fitting_tables.component_of(akey, comp_values),
                        (_COMP_LABEL.get(akey, akey), 'known drop', None, None)),)
                else:
                    unpriced = unpriced + (
                        (fam or '?') + ': accessory on the run, not priced',)

            box_here = False
            mid_equip = False
            # nid != root_id, not "not a picked root": a box picked alongside
            # its AHU is still priced inside the AHU's walk (review C3).
            if (dr is None and elem is not None and nid != root_id
                    and hvac_graph.is_equipment(elem)):
                mid_equip = True
                if nid in box_ids:
                    # What the MP supply costs to reach this box's inlet, before
                    # the box's own drop: the unit's max ESP less this is the
                    # pressure left at the box.
                    # Fan system effect is deliberately not part of the box
                    # check (Colin, 2026-10-29: no system effect on VAV boxes).
                    _reach = fric + fit + sum([_comp(_k) for _k in comp_keys], 0.0)
                    if _reach > boxes_here.get(nid, -1.0):
                        boxes_here[nid] = _reach
                    comp_keys = comp_keys + ('vav_box',)
                    items = items + ((
                        'VAV box {}'.format(nid), 'VAV box drop', None, box_drop,
                        ('VAV box', 'known drop', None, None)),)
                    box_here = True
                else:
                    # VAV off, or not a box: its own key, at the balancing
                    # damper value for now (VAV plan item 9, review S4).
                    comp_keys = comp_keys + ('mid_equip',)
                    items = items + ((
                        'Equipment {}'.format(nid),
                        'Mid tree equipment (as damper)', None, _comp('mid_equip'),
                        ('Mid tree equipment (as damper)', 'known drop', None, None)),)

            term = all_terminals.get(nid)
            if term is not None:
                # The diffuser the run ends at is itself a component drop; a
                # ceiling exhaust fan run ends at its wall or roof cap.
                is_cap = nid in cap_ids
                keys_here = comp_keys + ('cap' if is_cap else 'diffuser',)
                # Include exhaust roof cap: an inline fan's discharge cap,
                # charged once on each exhaust run (Colin, 2026-11-02).
                roof_cap = (exhaust_cap and not is_cap
                            and (term[1] or sys_class) == 'Exhaust Air')
                if roof_cap:
                    keys_here = keys_here + ('cap',)
                # Every supply diffuser is assumed to have a balancing damper
                # whether or not one is drawn (Colin, 2026-10-26), so a branch
                # with no damper accessory (a w_Damper take-off, or none shown)
                # still pays one. A damper drawn on the branch itself is never
                # charged twice; one on a main upstream of the branch is a
                # fire/smoke damper and does not cancel it (Colin, 2026-11-02).
                _assumed_damper = (
                        vav_checks.is_supply_class(term[1] or sys_class)
                        and not br_damper)
                if _assumed_damper:
                    keys_here = keys_here + ('damper',)
                # Include backdraft damper: an exhaust fan's run with none
                # modeled pays one (Colin, 2026-11-02); one drawn on the run
                # is already in comp_keys and is never charged twice.
                # The discharge of an inline exhaust fan: on every exhaust
                # run of the fan (Colin, 2026-11-03), like the roof cap.
                dis_here = (dis_rec if dis_rec is not None
                            and (term[1] or sys_class) == 'Exhaust Air' else None)
                if dis_here is not None:
                    keys_here = keys_here + dis_here['comp_keys']
                _assumed_backdraft = (
                        backdraft_roots is not None and root_id in backdraft_roots
                        and (term[1] or sys_class) == 'Exhaust Air'
                        and 'backdraft' not in comp_keys
                        and not (dis_here and 'backdraft' in dis_here['comp_keys']))
                if _assumed_backdraft:
                    keys_here = keys_here + ('backdraft',)
                comp_here = 0.0
                for _k in keys_here:
                    comp_here += _comp(_k)
                _cfm, term_class, _family = term
                key     = term_class or sys_class or 'Unknown'
                # Known drop added once to the RETURN path only (the "Include
                # Return Filter" option). Same on every return path, so it never
                # changes which run is the worst, only the total.
                filt    = return_filter_inwc if key == 'Return Air' else 0.0
                # The unit connection this path leaves through, and its fan
                # system effect (0 when none, or not checked).
                fse_rec = None
                for _first in fse_firsts:
                    if _first not in path:
                        continue
                    _fr = fse_recs[_first]
                    if _fr.get('by_tap'):
                        # A drop (Colin, 2026-11-11): the record of the tap
                        # this path leaves the drop through.
                        for _t, _sub in _fr['by_tap'].items():
                            if _t in path:
                                _fr = _sub
                                break
                    if fse_rec is None or _fr['loss'] > fse_rec['loss']:
                        fse_rec = _fr
                # The group at the connection (straight duct, transition, first
                # elbow) is charged once with its fan system effect, the larger.
                items_l, fse_loss, _gfr, _gft, _fitem, _gnote = _fse_group_max(
                    fse_rec, items)
                fse_items = (_fitem,) if _fitem is not None else ()
                # The run NOTE per connection on this run, written below only
                # for a run that is kept (Colin, 2026-11-11).
                _gov_src = ((fse_rec, _gnote),)
                fric_l, fit_l, dis_items = fric - _gfr, fit - _gft, ()
                n_d, n_f, len_l, unpriced_l = dcount, fcount, dlen, unpriced
                if dis_here is not None:
                    _drec = fse_recs.get(dis_here['first'])
                    if _drec is not None and _drec.get('by_tap'):
                        _dnames = set(_it[0] for _it in dis_here['items'])
                        for _t, _sub in _drec['by_tap'].items():
                            if 'Fitting {}'.format(_t) in _dnames:
                                _drec = _sub
                                break
                    dis_items, _dl, _dfr, _dft, _ditem, _dnote = _fse_group_max(
                        _drec, dis_here['items'])
                    _gov_src = _gov_src + ((_drec, _dnote),)
                    fric_l += dis_here['fric'] - _dfr
                    fit_l += dis_here['fit'] - _dft
                    n_d += dis_here['n_duct']
                    n_f += dis_here['n_fit']
                    len_l += dis_here['feet']
                    unpriced_l = unpriced_l + dis_here['unpriced']
                    if _drec is not None and (_drec['loss'] or _drec.get('none_group')):
                        fse_loss += _dl
                    if _drec is not None and _ditem is not None:
                        fse_items = fse_items + (_ditem,)
                current = best_here.get(key)
                # Equal runs (two taps at one station, F2 2026-11-04): the
                # lower terminal id is the index run, so Revit's connector
                # order cannot change which terminal is named.
                _tot_here = fric_l + fit_l + comp_here + filt + fse_loss
                _tot_cur = None if current is None else (
                    current['friction_inwc'] + current['fitting_inwc'] +
                    current['component_inwc'] + current['filter_inwc'] +
                    current['fse_inwc'])
                _better_best = (current is None or _tot_here > _tot_cur + 1e-12 or
                                (abs(_tot_here - _tot_cur) <= 1e-12 and
                                 nid < current['terminal_id']))
                _better_term = False
                if keep_runs:
                    # This terminal's own worst run (a terminal reached by
                    # two routes keeps the lossier one).
                    _tr = runs_here.setdefault(key, {}).get(nid)
                    _better_term = _tr is None or _tot_here > (
                        _tr['friction_inwc'] + _tr['fitting_inwc'] +
                        _tr['component_inwc'] + _tr['filter_inwc'] +
                        _tr['fse_inwc']) + 1e-12
                if _better_best or _better_term:
                    gov_notes = tuple(('note', _t) for _t in (
                        _run_note(_gr, _gf, nid) for _gr, _gf in _gov_src) if _t)
                    _rec = {
                        'friction_inwc':    fric_l,
                        'fitting_inwc':     fit_l,
                        'component_inwc':   comp_here,
                        'filter_inwc':      filt,
                        'fse_inwc':         fse_loss,
                        # The UNIT's fan system effect notes, the same on
                        # every class of this root: print them once per unit.
                        'fse_notes':        fse_unit_notes,
                        # Which governed at this run's connection(s): fan
                        # system effect or the individual losses (2026-11-09).
                        'fse_gov_notes':    gov_notes,
                        'fse_gray_ids':     fse_gray,
                        'fse_red_ids':      fse_red,
                        'fse_red_ducts':    fse_red_ducts,
                        'fse_assumptions':  fse_unit_assume,
                        'unit_side_fittings': pre[2],
                        'exhaust_fan_kind': ('ceiling' if root_id in discharge_roots
                                             else fan_kinds.get(root_id)),
                        'component_keys':   keys_here,
                        'duct_count':       n_d,
                        'fitting_count':    n_f,
                        'tap_bypass_count': tcount,
                        'length_ft':        len_l,
                        'terminal_id':      nid,
                        'terminal_name':    _family or '-',
                        'terminal_label':   _terminal_label(nid),
                        'unpriced':         unpriced_l,
                        'items':            fse_items + items_l + dis_items + (((
                            'Accessory (assumed)', 'Balancing damper at ' + (_family or '-'),
                            None, fitting_tables.component_of('damper', comp_values),
                            ('Balancing damper', 'known drop', None, None)),)
                            if _assumed_damper else ()) + (((
                            'Accessory (assumed)', 'Backdraft damper at the fan',
                            None, fitting_tables.component_of('backdraft', comp_values),
                            ('Backdraft damper', 'known drop', None, None)),)
                            if _assumed_backdraft else ()) + ((
                            'Terminal {}'.format(nid), _family or '-', None,
                            fitting_tables.component_of(
                                'cap' if is_cap else 'diffuser', comp_values),
                            (hvac_graph.CAP_LABEL if is_cap else 'Diffuser / grille',
                             'known drop', None, None)),) + (((
                            'Exhaust roof cap', 'Include exhaust roof cap', None,
                            fitting_tables.component_of('cap', comp_values),
                            (hvac_graph.CAP_LABEL, 'known drop', None, None)),)
                            if roof_cap else ()),
                    }
                    if _better_best:
                        best_here[key] = _rec
                    if _better_term:
                        runs_here[key][nid] = _rec

            for child_id in all_children.get(nid, []):
                # Per-PATH loop guard, not a global visited set. A single shared
                # visited set was a real bug: it marked a node seen on whichever
                # path reached it first, so where two routes merge (two AHUs
                # feeding a common duct, a ring main, or a mis-modelled shared
                # segment) the second route was thrown away at the merge and its
                # losses were never compared. Since this function exists to find
                # the WORST path, that silently under-reported external static.
                # Verified: a diamond graph reported 0.07 in. wc when the real
                # critical path was 0.52.
                if child_id in path:
                    continue
                # Through mid tree equipment the walk only goes on along its
                # own side: MP supply into a fan powered box never carries on
                # out through the box's ducted return (review C2).
                if mid_equip and not vav_checks.same_side(
                        sys_class, _side_below(child_id)):
                    continue
                extra_fit = fit
                extra_t   = tcount
                extra_items = items
                extra_unpriced = unpriced
                c_fric, c_dlen = fric, dlen
                if dr is not None:
                    # Friction along a main with taps: each piece at its own
                    # flow, only as far as this run goes (_seg_friction).
                    seg = _seg_friction(nid, child_id)
                    _on = _tap_order_note(nid)
                    if _on and _on not in extra_unpriced:
                        extra_unpriced = extra_unpriced + (_on,)
                    if seg is not None:
                        _sf, _ft = seg
                        c_fric = fric - dr.friction_loss_inwc + _sf
                        c_dlen = dlen - float(dr.length_ft) + _ft
                        _name = 'Duct {}'.format(nid)
                        for _i in range(len(extra_items) - 1, -1, -1):
                            if extra_items[_i][0] == _name:
                                _old = extra_items[_i]
                                extra_items = extra_items[:_i] + ((
                                    _name,
                                    '{}, {:.1f} of {:.1f} ft, flow stepped at taps'.format(
                                        getattr(dr, 'size', '?'), _ft, float(dr.length_ft)),
                                    _old[2], _sf,
                                    (_old[4][0], _old[4][1], _old[4][2], _ft)),) + \
                                    extra_items[_i + 1:]
                                break
                    # Every tap on this main that the run passes before it
                    # leaves restricts it: Cs from the tap's 2001 table at its
                    # own flow split (worksheet C where flows are unreadable).
                    # A run leaving through a tap passes only the taps AHEAD
                    # of it on the duct when their order is known.
                    taps = _taps(nid)
                    _tl, n_by, _parts, _fbs = _taps_passed(nid, child_id)
                    if n_by > 0:
                        extra_fit = extra_fit + _tl
                        extra_t = extra_t + n_by
                        extra_unpriced = extra_unpriced + tuple(_fbs)
                        for _tbl, _c, _one in _parts:
                            extra_items = extra_items + ((
                                'Tap passed', 'on duct {}, {} {:.2f}'.format(nid, _tbl, _c),
                                dr.fpm, _one,
                                ('Tap passed on main (straight through)', _tbl, _c, None)),)
                    # A tap at the terminus of a main is the worksheet's "end of
                    # main branch" case and carries a different C, so the swing
                    # is applied here where the parent duct is known. A tap
                    # priced from its 2001 table already reads the end of a
                    # main through its flow split (Qb/Qc at 1.0), so it gets
                    # no swing.
                    if (child_id in taps and not _continues(nid)
                            and _tap_turn(child_id) is None):
                        ch = all_nodes.get(child_id)
                        if ch is not None:
                            end_c = fitting_tables.takeoff_c(
                                dr.sys_class, is_end_of_main=True, c=c_values)
                            typ_c = fitting_tables.takeoff_c(
                                dr.sys_class, c=c_values)
                            _sw = ((end_c - typ_c) *
                                   hvac_graph.velocity_pressure_inwg(dr.fpm))
                            extra_fit = extra_fit + _sw
                            extra_items = extra_items + ((
                                'End of main tap', 'C {:.2f} in place of {:.2f}'.format(
                                    end_c, typ_c),
                                dr.fpm, _sw,
                                ('End of main tap adjustment', 'RJA worksheet', None, None)),)
                c_fpm, c_area = up_fpm, up_area
                c_class = sys_class
                if jn_here:
                    # The junction's loss for THIS exit leg: Cc,b into a
                    # branch, Cc,s straight through (_junction_exit), the
                    # next fitting priced at this leg's own velocity. A leg
                    # that cannot be read takes the worksheet, FALLBACK.
                    _jfam = hvac_graph.fitting_family_name(elem) or '?'
                    _jx = _junction_exit(nid, child_id, sys_class)
                    if _jx is not None:
                        _jl, _jd, _jnotes, _jfpm, _jarea = _jx
                        _jfb = [x for x in _jnotes if x and x.startswith('FALLBACK')]
                        _jrest = _notes_text([x for x in _jnotes
                                              if x and not x.startswith('FALLBACK')])
                        _jtext = '; '.join(_jfb + ([_jrest] if _jrest else [])) or None
                        c_fpm, c_area = _jfpm, _jarea
                    else:
                        # The exit leg is known here (_junction read it), so
                        # straight through takes the worksheet's tap main C,
                        # a branch the take-off C (review, 2026-11-11).
                        _jj = _junction(nid)
                        _jleg = None
                        if _jj is not None:
                            _jleg = ('straight' if _jj['layout']['straight'] == child_id
                                     else 'branch')
                        _jl, _jc, _jtext, _ja, _jd = _price_fitting(
                            elem, sys_class, up_fpm, up_area,
                            against_flow=fitting_tables.flows_toward_unit(sys_class),
                            nid=nid, junction_leg=_jleg,
                            fb_why=junction_exit_why.get((nid, child_id, sys_class)))
                    extra_fit = extra_fit + _jl
                    if _jtext:
                        extra_unpriced = extra_unpriced + (_jtext,)
                    extra_items = extra_items + ((
                        'Fitting {}'.format(nid), '{}, {}'.format(_jfam, _jd),
                        up_fpm, _jl, _takeoff_key(elem, _jfam, _jd, 1, _jtext)),)
                if dr is not None and dr.area_ft2 and child_id not in _taps(nid):
                    # Past the last tap on an unbroken main only the flow the
                    # taps left is still moving, so a fitting there is priced
                    # at that velocity, not the unit end cfm (Colin,
                    # 2026-11-02, _remaining_cfm).
                    _rem = _remaining_cfm(nid)
                    if _rem is not None:
                        c_fpm = _rem / dr.area_ft2
                if nid == root_id and dr is None:
                    # Leaving the unit: start at the unit connection's
                    # velocity, not 0, so a fitting right on the unit is
                    # priced (Colin, 2026-10-29).
                    _st = _unit_start(root_id, child_id)
                    if _st is not None:
                        c_fpm, c_area, c_class = _st
                if box_here:
                    # Past a box the air moves at the box OUTLET duct's
                    # velocity, not the inlet runout's, so a fitting right on
                    # the box discharge is priced at the right Pv.
                    _od = _first_duct(child_id)
                    if _od is not None:
                        c_fpm, c_area = _od.fpm, _od.area_ft2
                stack.append((child_id, c_fric, extra_fit, comp_keys, dcount,
                              fcount, extra_t, c_dlen, c_fpm, c_area,
                              c_class, extra_unpriced,
                              path | frozenset([child_id]), in_flex,
                              extra_items,
                              # past a split the path is on a new branch
                              br_damper and len(all_children.get(nid, [])) < 2))

        if best_here:
            best[root_id] = best_here
        if runs_here:
            runs[root_id] = runs_here
        if boxes_here:
            box_reach[root_id] = boxes_here

    result = {}
    sf = 1.0 + safety_pct / 100.0
    for root_id, per_class in best.items():
        out = {}
        boxes = dict((bid, (v, v * sf))
                     for bid, v in box_reach.get(root_id, {}).items())
        def _finish(b):
            # Fan system effect is in the subtotal, so the safety factor
            # applies to it like every other fitting (Colin, 2026-10-29).
            sub = (b['friction_inwc'] + b['fitting_inwc'] + b['component_inwc'] +
                   b['filter_inwc'] + b['fse_inwc'])
            b['subtotal_inwc'] = sub
            b['total_inwc']    = sub * sf
            b['truncated']     = root_id in truncated_roots
            b['boxes']         = boxes
        for sys_class, b in per_class.items():
            _finish(b)
            if keep_runs:
                ranked = list(runs.get(root_id, {}).get(sys_class, {}).values())
                for r in ranked:
                    if r is not b:
                        _finish(r)
                ranked.sort(key=lambda r: (-r['total_inwc'], r['terminal_id']))
                b['runs'] = ranked
            out[sys_class]     = b
        result[root_id] = out
    return result


def _vav_box_report(root_ids, vav_by_root, root_labels, critical, vav_set):
    """The VAV BOXES block of the System Summary, as structured data: per unit,
    each box and (external static on) the MP loss to its inlet and the
    pressure left at it.

    Pressure left = unit max external static less the MP loss WITH the safety
    factor, the same basis as the external static total it is compared to.
    Only shown when the engineer entered the unit's max (VAV plan item 8).

    Every value is worked out once here. _vav_summary_lines() flattens the
    'lines' of each record into the console's plain text fallback; the
    console and the drafting view render the other fields as a table, and
    duct_perf_console.vav_corrections() turns LOW boxes and diffusers on an
    MP main into System Corrections lines.

    Returns {'title', 'basis', 'minimum', 'lines' (block header lines), 'units': [{'label',
    'lines', 'boxes': [{'box', 'lines', 'loss', 'loss_sf', 'left',
    'status'}], 'mp_terminals', 'mp_terminal_text'}]}. status is None
    (external static off), 'NOT REACHED', 'NOT CHECKED', 'OK' or 'LOW'.
    """
    umax = vav_set.get('unit_max_esp')
    minimum = vav_set.get('min_box_inwc')   # None: not entered, not judged
    minimum = None if minimum is None else float(minimum)
    basis = 'MP supply {:.0f} FPM, {:.2f} in. wc per 100 ft'.format(
        float(vav_set['mp_fpm']), float(vav_set['mp_fric']))
    title = 'VAV BOXES ({})'.format(basis)
    report = {'title': title, 'basis': basis, 'lines': ['', title], 'units': [],
              'minimum': minimum, 'unit_max': umax}

    for rid in root_ids:
        vs = vav_by_root.get(rid)
        if vs is None or not vs.boxes:
            continue
        label = root_labels.get(rid, 'id {}'.format(rid))
        unit = {'label': label, 'lines': ['  === {} ==='.format(label)],
                'boxes': [], 'mp_terminals': 0, 'mp_terminal_text': None}
        report['units'].append(unit)
        boxes = {}
        if critical is not None:
            for leaf in (critical.get(rid) or {}).values():
                boxes = leaf.get('boxes') or {}
                break
        for box in vs.boxes:
            rec = {'box': box, 'lines': ['    Box {}'.format(box)],
                   'loss': None, 'loss_sf': None, 'left': None, 'status': None}
            unit['boxes'].append(rec)
            if critical is None:
                continue
            reach = boxes.get(box)
            if reach is None:
                rec['status'] = 'NOT REACHED'
                rec['lines'].append(
                    '      MP loss to inlet: not reached by the index run search')
                continue
            rec['loss'], rec['loss_sf'] = reach[0], reach[1]
            text = '      MP loss to inlet {:.3f} in. wc ({:.3f} with safety factor)'.format(
                reach[0], reach[1])
            left = vav_checks.pressure_left(umax, reach[1])
            rec['left'] = left
            if left is None:
                rec['status'] = 'NOT CHECKED'
                text += ', pressure left not checked (no unit max entered)'
            else:
                text += ', pressure left {:.3f} in. wc'.format(left)
                if minimum is None:
                    rec['status'] = 'NOT CHECKED'
                    text += ', not judged (no minimum at a box entered)'
                elif vav_checks.pressure_low(left, minimum):
                    rec['status'] = 'LOW'
                    text += ', LOW: under the {:.2f} in. wc minimum'.format(minimum)
                else:
                    rec['status'] = 'OK'
            rec['lines'].append(text)
        if vs.mp_terminals:
            unit['mp_terminals'] = len(vs.mp_terminals)
            unit['mp_terminal_text'] = (
                '{} diffuser(s) on an MP main with no box upstream '
                '(flagged red on the duct feeding each).'.format(len(vs.mp_terminals)))
    return report


def _vav_summary_lines(report):
    """Flatten a _vav_box_report() into the drafting view text lines."""
    lines = list(report['lines'])
    for unit in report['units']:
        lines.extend(unit['lines'])
        for rec in unit['boxes']:
            lines.extend(rec['lines'])
    # The MP diffuser line is a System Corrections line
    # (duct_perf_console.vav_corrections), not repeated here.
    return lines


# Minimum transition clearance between a duct and a genuinely smaller real
# downstream neighbour, inches.
_CLEARANCE_MARGIN_IN = 2.0


def _fails_height_clearance(dr, downstream_height_in):
    """True if this duct fails the diffuser/duct height clearance rule.

    Only fires where the real downstream neighbour is actually SMALLER than
    this duct (a genuine size reduction) and the step down is too shallow to
    build a transition into — an equal or larger downstream connection is a
    continuous run and needs no margin at all.

    Shared by the main and the branch labelling paths deliberately: this is a
    hard installability problem, orthogonal to whichever sizing method judged
    the duct, and it overrides both of them. Keeping it in one function is
    what stops the two paths' clearance behaviour drifting apart.
    """
    return hvac_graph.fails_height_clearance(
        hvac_graph.effective_height_in(dr.elem), downstream_height_in,
        _CLEARANCE_MARGIN_IN)


def _clearance_reason(dr, down_h, down_elem, is_branch, branch_duct_ids, notes):
    """Specific Reason text for a clearance failure, and its NOTE appended to
    notes. What sits downstream decides the wording: the next duct of the same
    run, a branch duct leaving a main, or a diffuser neck."""
    duct_h = hvac_graph.effective_height_in(dr.elem)
    if duct_h is None or down_h is None or down_elem is None:
        return 'Diffuser/Duct Clearance'
    role = 'Branch' if is_branch else 'Main'
    down_id = eid_int(down_elem.Id)
    if hvac_graph.is_terminal(down_elem):
        kind = 'neck'
        down_label = '{} ({}" neck)'.format(down_id, hvac_graph._fmt_in(down_h))
    else:
        kind = ('branch' if (not is_branch and down_id in branch_duct_ids)
                else 'step')
        down_label = '{} ({})'.format(down_id, _duct_size_label(down_elem))
    up_label = '{} ({})'.format(dr.element_id, _duct_size_label(dr.elem))
    reason, note = hvac_graph.clearance_texts(kind, role, duct_h, down_h,
                                              up_label, down_label,
                                              _CLEARANCE_MARGIN_IN)
    notes.append(note)
    return reason


def _limits_of(dr, custom_limits):
    """(max_fpm, max_friction) for one duct: the one lookup every check uses
    (vav_checks.limits_for), MP limits for a VAV medium pressure duct."""
    return vav_checks.limits_for(dr.sys_class, custom_limits,
                                 hvac_graph.FIRM_DEFAULTS,
                                 getattr(dr, 'is_mp', False))


def _duct_label(dr, custom_limits, tol_pct, downstream_height_in=None):
    """Worst of velocity and friction checks (one-sided, unchanged), then
    downgraded to PURPLE if the duct is GREEN but oversized (a smaller
    standard size would still satisfy max FPM and max friction), then
    overridden to RED regardless of the above if the duct fails the
    diffuser/duct height clearance rule (its height doesn't clear its real
    downstream neighbor's height by the required 2" margin, wherever that
    neighbor is actually smaller — a hard installability problem, checked
    independent of velocity/friction).

    MAIN DUCTS ONLY as of the branch/diffuser split, with one exception: a
    branch feeding a round sidewall diffuser also comes through here (no table
    yet). Slot branches are judged off the slot family since 2026-10-22.
    Behaviour is unchanged from before the split — see _branch_duct_label()
    for what every other branch gets instead, and _label_duct() for the
    dispatch between them.

      Green  : cfm <= max_cap
      Yellow : max_cap < cfm <= max_cap * (1 + tol_pct/100)
      Red    : cfm > max_cap * (1 + tol_pct/100), OR fails height clearance
      Purple : GREEN AND a smaller standard size exists (see _suggest_size_info)

    Returns (label, max_cap_cfm, reason) — reason is one of _MAIN_REASONS,
    DUCT_REDUCTION_REASON, MEETS_STANDARD_REASON, 'Diffuser/Duct Clearance', or '' (GRAY).

    A VAV medium pressure duct (dr.is_mp) is judged against the MP limits,
    PURPLE included: the shrink candidate is tested on the MP limits too
    (Colin, 2026-11-02; it was never PURPLE before).
    """
    max_fpm, max_friction = _limits_of(dr, custom_limits)
    tol_fac = tol_pct / 100.0
    # At the limit is green (Colin, 2026-11-12, TB-M01C: 8,000 cfm in a 64x30
    # at exactly 600 fpm read YELLOW because 600 x 13.333 ft2 stores as
    # 7,999.99999). One part in a million absorbs the float error only.
    _eps = 1.0 + 1e-6

    # Velocity check (CFM vs capacity at max FPM)
    if dr.cfm <= 0 or dr.area_ft2 <= 0:
        vel_label = 'GRAY'
        max_cap   = 0.0
    else:
        max_cap = max_fpm * dr.area_ft2
        red_cap = max_cap * (1.0 + tol_fac)
        if dr.cfm <= max_cap * _eps:
            vel_label = 'GREEN'
        elif dr.cfm <= red_cap * _eps:
            vel_label = 'YELLOW'
        else:
            vel_label = 'RED'

    # Friction check
    if dr.friction_per_100ft <= 0 or max_friction <= 0:
        fric_label = 'GRAY'
    else:
        red_fric = max_friction * (1.0 + tol_fac)
        if dr.friction_per_100ft <= max_friction * _eps:
            fric_label = 'GREEN'
        elif dr.friction_per_100ft <= red_fric * _eps:
            fric_label = 'YELLOW'
        else:
            fric_label = 'RED'

    fric_wins = _PRIORITY.get(fric_label, 0) > _PRIORITY.get(vel_label, 0)
    combined  = fric_label if fric_wins else vel_label
    if combined in ('YELLOW', 'RED'):
        which = ('both' if fric_label == vel_label else ('fric' if fric_wins else 'vel'))
        reason = _MAIN_REASONS[(combined, which)]
    elif combined == 'GREEN':
        reason = MEETS_STANDARD_REASON
    else:
        reason = ''

    # GREEN ducts get downgraded to PURPLE if a smaller standard size — both
    # width and height at least one nominal size down — would still satisfy
    # both max FPM and max friction. i.e. it's oversized.
    if combined == 'GREEN':
        _, sugg_area = _suggest_shrink_info(dr, custom_limits, downstream_height_in)
        if sugg_area is not None:
            combined = 'PURPLE'
            reason   = DUCT_REDUCTION_REASON

    # Diffuser/duct height clearance — hard override, independent of
    # everything above.
    if _fails_height_clearance(dr, downstream_height_in):
        combined = 'RED'
        reason   = 'Diffuser/Duct Clearance'

    return combined, max_cap, reason


# ── branch labelling (diffuser-table method) ──────────────────────────────────

_ROLE_LABELS = {
    'ceiling':  'Branch (Ceiling)',
    'sidewall': 'Branch (Sidewall)',
    'slot':     'Branch (Slot)',
    'round_sidewall': 'Branch (Round Sidewall)',
    'rect_ceiling':   'Branch (Rect Ceiling)',
}


def _duct_role_label(is_branch, branch_kind):
    """Display string for the Duct Role column.

    The diffuser category is carried into the label rather than just
    'Main'/'Branch' because it is the single most useful thing to see when a
    branch verdict looks wrong: it says which table the row was judged
    against, or that no table could be chosen at all.
    """
    if not is_branch:
        return 'Main'
    return _ROLE_LABELS.get(branch_kind, 'Branch (Unknown)')


def _branch_duct_label(dr, terminal_elem, terminal_cfm, downstream_height_in=None):
    """Label a branch duct by diffuser-table size match instead of velocity.

    All of the actual sizing logic lives in hvac_graph.branch_diffuser_check();
    this only adds the height-clearance override that applies to every duct
    regardless of how it was judged.

    Returns (label, max_cap_cfm, reason, BranchDiffuserResult). max_cap_cfm is
    always 0.0 — a branch has no velocity capacity figure, and the callers
    that carry that slot only ever read the label out of it. Status can be
    GREEN, RED or GRAY (converted to RED by _label_duct) but never YELLOW or PURPLE: there is no tolerance band
    on a size match, and a branch is never checked for oversize (the
    auto-sizing diffuser family sets the branch size).
    """
    res = hvac_graph.branch_diffuser_check(
        terminal_elem, terminal_cfm, hvac_graph.duct_installed_size_in(dr.elem))
    label  = res.status
    reason = res.reason
    if _fails_height_clearance(dr, downstream_height_in):
        label  = 'RED'
        reason = 'Diffuser/Duct Clearance'
    return label, 0.0, reason, res


def _label_duct(dr, custom_limits, tol_pct, downstream_height_in, branch_ctx):
    """Label one duct GREEN / YELLOW / RED / PURPLE, or GRAY.

    GRAY means the duct could not be judged (no airflow reaching it, no size,
    unreadable diffuser, no table). That is nearly always a broken or
    disconnected system, so it is shown gray in the view and explained in the
    legend rather than passed or failed. Gray ducts are not in the flagged
    table.
    """
    label, cap, reason, branch_res = _label_duct_checked(
        dr, custom_limits, tol_pct, downstream_height_in, branch_ctx)
    if label == 'GRAY' and not reason:
        if dr.cfm <= 0:
            reason = 'No airflow data'
        elif dr.area_ft2 <= 0:
            reason = 'No duct size data'
        else:
            reason = 'Cannot check'
    return label, cap, reason, branch_res


def _label_duct_checked(dr, custom_limits, tol_pct, downstream_height_in, branch_ctx):
    """Send one duct to whichever check applies to it.

    branch_ctx is None for a main, or (diffuser_category, terminal_elem,
    terminal_cfm) for a branch — built once per duct in main() from the
    downstream-terminal map, so the graph is only walked once.

    Returns (label, max_cap_cfm, reason, branch_result_or_None). The fourth
    element is None for anything judged the ductulator way, and is what tells
    the output tables where the Required Size came from. It does NOT change
    whether velocity/friction print: every duct prints its real numbers.
    May return GRAY.
    """
    if branch_ctx is None:
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    kind, terminal_elem, terminal_cfm = branch_ctx

    if dr.cfm <= 0:
        # A branch to a diffuser with Flow 0 cannot be judged: GRAY like a
        # main with no airflow, never a table pass (Colin, 2026-11-02, TB-M01
        # D-5 drew green). Only the clearance rule still overrides it.
        label, reason = 'GRAY', 'No airflow data'
        if _fails_height_clearance(dr, downstream_height_in):
            label, reason = 'RED', 'Diffuser/Duct Clearance'
        return label, 0.0, reason, None

    if (terminal_elem is not None and kind not in ('ceiling', 'sidewall')
            and not hvac_graph.is_auto_sizing_diffuser(terminal_elem)):
        # Not an RJA auto-sizing diffuser and no neck or face table to judge
        # it by (a slot from another family, whose length and slot count are
        # not read): the branch gets the main duct check. Ceiling and sidewall
        # diffusers from any family are judged on the RJA diffuser tables
        # (Colin, 2026-11-02; until then every non-RJA one got the main check).
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    if kind in ('round_sidewall', 'rect_ceiling'):
        # Round sidewall diffusers have no table yet; Colin plans to add round
        # diffuser data later (2026-10-16). A rectangular neck pointing up or
        # down is a ceiling diffuser, never judged on the sidewall tables, and
        # has no table of its own yet (Colin, 2026-11-02). Both fall back to
        # the main check.
        # Slots are judged off the auto-sized family instead, in
        # hvac_graph.branch_diffuser_check (Colin, 2026-10-22).
        label, cap, reason = _duct_label(dr, custom_limits, tol_pct, downstream_height_in)
        return label, cap, reason, None

    if kind is None:
        # Connector geometry unreadable — no table can be chosen. Gray, not a
        # pass: an unverifiable branch must never look like a clean one.
        label  = 'GRAY'
        reason = 'Cannot classify diffuser'
        if _fails_height_clearance(dr, downstream_height_in):
            label  = 'RED'
            reason = 'Diffuser/Duct Clearance'
        return label, 0.0, reason, None

    return _branch_duct_label(dr, terminal_elem, terminal_cfm, downstream_height_in)


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(eid_int(elem.Id))


def _duct_size_label(elem):
    """Return readable size: '10"' for round/spiral, '18"x12"' for rectangular."""
    try:
        d = elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            return '{:.0f}"'.format(d.AsDouble() * 12.0)
        w = elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h = elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w and h and w.AsDouble() > 0 and h.AsDouble() > 0:
            return '{:.0f}"x{:.0f}"'.format(w.AsDouble() * 12.0, h.AsDouble() * 12.0)
    except Exception:
        pass
    return '?'


def _duct_is_oval(elem, eid):
    """True when the duct's end connector is ConnectorProfileType.Oval. An
    oval duct is read as W x H for area, velocity and friction (hvac_graph),
    and the report says so (Colin, 2026-11-02). False when unreadable,
    logged."""
    try:
        for c in elem.ConnectorManager.Connectors:
            if c.ConnectorType == ConnectorType.End:
                if c.Shape == ConnectorProfileType.Oval:
                    hvac_graph.log.info('_duct_is_oval: duct %s Oval, read as W x H', eid)
                    return True
                return False
    except Exception as ex:
        hvac_graph.log.info('_duct_is_oval: duct %s connector shape unreadable: %s',
                            eid, ex)
        return False
    hvac_graph.log.info('_duct_is_oval: duct %s end connector NOT FOUND', eid)
    return False


# Standard round sizes, the minimum dimension and the even-inch rounding are
# shared with Size Duct (lib/duct_sizing.py), so the Required Size this tool
# suggests is the size Size Duct would write.
from duct_sizing import ROUND_SIZES as _ROUND_SIZES          # noqa: E402
from duct_sizing import MIN_DUCT_DIM as _MIN_DUCT_DIM        # noqa: E402
from duct_sizing import snap_even as _snap_even              # noqa: E402


def _suggest_shrink_info(dr, custom_limits, downstream_height_in=None):
    """Return (label, area_ft2) for a SMALLER standard duct size than what's
    installed, satisfying both velocity AND friction limits — no tolerance
    band, straight against max FPM / max friction. Returns ('-', None) if no
    strictly-smaller standard size works.

    Used for the oversized/PURPLE suggestion only. Round/spiral: smaller
    standard diameter only. Rectangular: BOTH width and height must drop by
    at least one nominal (2") step from installed — a candidate that only
    shrinks one dimension (e.g. 12x12 -> 12x10) is not a valid suggestion,
    since that's a marginal height-only trim, not a genuine downsize. Only
    a candidate like 12x12 -> 10x10, where both dimensions are a nominal
    size smaller, qualifies. All dimensions snapped to 2" intervals.

    Never suggests below the firm's 6" absolute minimum (_ROUND_SIZES/
    _MIN_DUCT_DIM). If downstream_height_in is given (the effective height of
    whatever this duct's real downstream neighbor is, from
    hvac_graph.max_downstream_height_in), also never suggests a candidate
    that would leave less than the required 2" clearance margin wherever
    downstream_height_in is actually smaller than the candidate — a
    same-size or larger downstream connection needs no transition margin.
    """
    max_fpm, max_friction = _limits_of(dr, custom_limits)
    if dr.cfm <= 0 or max_fpm <= 0:
        return '-', None

    def _clears(candidate_in):
        if downstream_height_in is None:
            return True
        if downstream_height_in >= candidate_in:
            return True  # no reduction at this candidate, no margin needed
        return candidate_in >= downstream_height_in + 2.0

    try:
        d = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            installed_d = _snap_even(d.AsDouble() * 12.0)
            for std_d in _ROUND_SIZES:
                if std_d >= installed_d:
                    break
                if not _clears(std_d):
                    continue
                area_ft2 = math.pi * (std_d / 24.0) ** 2
                vel      = dr.cfm / area_ft2
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d), dr.roughness_ft)
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"'.format(std_d), area_ft2
            return '-', None

        w_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w_param and h_param and w_param.AsDouble() > 0 and h_param.AsDouble() > 0:
            w_in = _snap_even(w_param.AsDouble() * 12.0)
            h_in = _snap_even(h_param.AsDouble() * 12.0)
            # Both width and height must be at least one nominal (2") step
            # below what's installed — a candidate that only shrinks one
            # dimension isn't a genuine downsize suggestion.
            best = None
            for new_w in range(_MIN_DUCT_DIM, w_in - 2 + 1, 2):
                min_h = max(_snap_even(new_w / 4.0), _MIN_DUCT_DIM)
                max_h = min(h_in - 2, new_w * 4)
                for new_h in range(min_h, max_h + 1, 2):
                    if not _clears(new_h):
                        continue
                    area_ft2 = new_w * new_h / 144.0
                    vel      = dr.cfm / area_ft2
                    d_h      = 4.0 * new_w * new_h / (2.0 * (new_w + new_h))
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h, dr.roughness_ft)
                    if vel <= max_fpm and fric <= max_friction:
                        if best is None or area_ft2 < best[1]:
                            best = ('{}"x{}"'.format(new_w, new_h), area_ft2)
            if best is not None:
                return best
    except Exception:
        pass
    return '-', None


def _suggest_size_info(dr, custom_limits, downstream_height_in=None):
    """Return (label, area_ft2) for the smallest standard duct size satisfying
    both velocity AND friction limits — no tolerance band, straight against
    max FPM / max friction. area_ft2 is None when no standard size could be
    found (duct needs to grow past the largest round size, or past what
    width-expansion covers for rectangular).

    Used for the undersized/YELLOW/RED suggestion (grow direction) — a truly
    undersized duct can never satisfy at a smaller height than installed
    either (smaller height only makes velocity/friction worse), so searching
    the full range from the AR floor upward is safe and never suggests a
    shrink here.
    Round/spiral: iterates standard diameters smallest-first.
    Rectangular: keeps the installed width fixed, iterates height from the
    smallest AR-valid value upward. Expands width only if no height at the
    installed width satisfies both limits within AR 4:1.
    All suggested dimensions are snapped to 2" intervals (2, 4, 6, 8...) —
    odd sizes are not stocked/installed.
    Both constraints must be satisfied — takes the binding (larger) of the two requirements.

    Never suggests below the firm's 6" absolute minimum (_ROUND_SIZES/
    _MIN_DUCT_DIM), and (same as _suggest_shrink_info) never suggests a
    candidate that would leave less than the required 2" clearance margin
    wherever downstream_height_in is actually smaller than the candidate.
    """
    max_fpm, max_friction = _limits_of(dr, custom_limits)
    if dr.cfm <= 0 or max_fpm <= 0:
        return '-', None

    def _clears(candidate_in):
        if downstream_height_in is None:
            return True
        if downstream_height_in >= candidate_in:
            return True
        return candidate_in >= downstream_height_in + 2.0

    try:
        d = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if d is not None and d.AsDouble() > 0:
            # Round / spiral — first standard diameter satisfying both vel and friction
            for std_d in _ROUND_SIZES:
                if not _clears(std_d):
                    continue
                area_ft2 = math.pi * (std_d / 24.0) ** 2
                vel      = dr.cfm / area_ft2
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, float(std_d), dr.roughness_ft)
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"'.format(std_d), area_ft2
            return '>{}"'.format(_ROUND_SIZES[-1]), None

        # Rectangular — keep width fixed, find smallest satisfying height
        w_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        h_param = dr.elem.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if w_param and h_param and w_param.AsDouble() > 0 and h_param.AsDouble() > 0:
            w_in = _snap_even(w_param.AsDouble() * 12.0)
            h_in = _snap_even(h_param.AsDouble() * 12.0)

            # AR 4:1 both ways bounds the height search at this fixed width:
            # too-flat (w > 4h) below, too-tall (h > 4w) above.
            min_h = max(_snap_even(w_in / 4.0), _MIN_DUCT_DIM)
            max_h = w_in * 4
            for new_h in range(min_h, max_h + 2, 2):
                if not _clears(new_h):
                    continue
                area_ft2 = w_in * new_h / 144.0
                vel      = dr.cfm / area_ft2
                d_h      = 4.0 * w_in * new_h / (2.0 * (w_in + new_h))
                fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h, dr.roughness_ft)
                if vel <= max_fpm and fric <= max_friction:
                    return '{}"x{}"'.format(w_in, new_h), area_ft2

            # No height at the installed width works within AR — expand width
            for new_w in range(int(w_in) + 2, int(w_in) + 60, 2):
                min_h2 = max(_snap_even(new_w / 4.0), _MIN_DUCT_DIM)
                max_h2 = new_w * 4
                for new_h in range(min_h2, max_h2 + 2, 2):
                    if not _clears(new_h):
                        continue
                    area_ft2 = new_w * new_h / 144.0
                    vel      = dr.cfm / area_ft2
                    d_h      = 4.0 * new_w * new_h / (2.0 * (new_w + new_h))
                    fric     = hvac_graph.duct_friction_loss_per_100ft(vel, d_h, dr.roughness_ft)
                    if vel <= max_fpm and fric <= max_friction:
                        return '{}"x{}"'.format(new_w, new_h), area_ft2
    except Exception:
        pass
    return '-', None


def _suggest_size(dr, custom_limits, tol_pct, duct_label, downstream_height_in=None):
    """Formatted-string wrapper for schedule/table display. PURPLE (oversized)
    ducts get a shrink-only suggestion (both width and height at least one
    nominal size down); YELLOW/RED (undersized) ducts get the grow-capable
    suggestion (may need a wider duct)."""
    if duct_label == 'PURPLE':
        size, _ = _suggest_shrink_info(dr, custom_limits, downstream_height_in)
    else:
        size, _ = _suggest_size_info(dr, custom_limits, downstream_height_in)
    return size


def _row_cells(label, dr, reason, role, branch_res,
               custom_limits, tol_pct, downstream_height_in):
    """Every cell value for one flagged duct, keyed by _COLUMN_DEFS key.

    Built once per duct and handed to both renderers, so the drafting-view
    schedule and the console table cannot report different numbers for the
    same duct no matter which columns each is showing.

    '#' is deliberately NOT produced here. The two tables number their rows
    differently on purpose — the schedule's numbers match the keynote circles
    placed in the view (element-id order), while the console list is re-sorted
    worst-first — so each renderer supplies its own row index.

    Every duct prints its REAL velocity and friction, mains and branches
    alike. Nothing is ever N/A, and the limits are NOT repeated per row: the
    engineer set them in the dialog, and Status + Reason already say whether
    velocity or friction is what failed.

    A branch shows its numbers WITHOUT being judged on them. A branch's verdict
    comes from the diffuser tables only (Colin, 2026-09-24: "a branch at 1100
    FPM with a correct diffuser goes to green as the diffuser is the limiting
    factor ... there should never be a time that a diffuser works but the duct
    fails greatly").
    """
    if branch_res is not None:
        required = branch_res.required_size
    else:
        required = _suggest_size(dr, custom_limits, tol_pct, label, downstream_height_in)

    fpm_cell  = '{:.0f}'.format(float(dr.fpm))
    fric_cell = '{:.3f}'.format(float(dr.friction_per_100ft))

    return {
        'status':    label,
        'reason':    reason,
        'role':      role,
        'size':      _duct_size_label(dr.elem),
        'required':  required,
        'cfm':       '{:.0f}'.format(dr.cfm),
        'fpm':       fpm_cell,
        'fric':      fric_cell,
        'length':    '{:.1f}'.format(dr.length_ft),
    }


# ── Schedule table in a Drafting View ─────────────────────────────────────────
# Drawn in the shared RJA table look (lib/rja_drafting.py, Colin 2026-10-21:
# "update the hvac duct ones to use the new styling"), the same bands, shaded
# headers and banded rows as the Water Performance and WSFU sheets.

# Color key + meaning shown in the legend, in display order.
# Wording covers both check methods, since one view mixes mains and branches.
def _legend_rows(words=None):
    """[(color key, meaning)] for the sheet legend, the branch wording
    from the Scope's words (Colin, 2026-11-09). The clearance wording is the
    Reason column's name and stays."""
    one = (words or duct_perf_console.scope_words('rtu'))['one']
    return [
        ('GREEN',  'Main: meets the design standard. Branch: {} and duct both correctly sized.'.format(one)),
        ('PURPLE', 'Main only: reduced duct size available, a smaller standard size still meets the design standard.'),
        ('YELLOW', 'Main only: above the max standard, within the tolerance.'),
        ('RED',    'Main: high velocity or pressure loss, over the max standard plus tolerance. Branch: {} or duct undersized. '
                   'Either: fails the diffuser/duct height clearance (see Reason).'.format(one)),
        ('GRAY',   'Not judged: no CFM reaches it (broken or disconnected system), '
                   'or a check could not be read (a NOTE names the duct and why). Or '
                   'an elbow at a unit connection whose fan system effect is too high '
                   '(see External Static Pressure).'),
    ]

# Columns read best left aligned; the rest are centred.
_LEFT_COLUMNS = ('reason', 'role')


# Sheet column proportions of the table width (the flagged duct table's).
_SHEET_EQUIP_COLS = (0.34, 0.16, 0.12, 0.38)
_SHEET_EQUIP_COLS_3 = (0.56, 0.24, 0.20)    # Exhaust System, no other air column
_SHEET_CLASS_W = 0.40       # by system class table, left of the types table
_SHEET_TYPES_W = 0.56       # diffuser types table, right of it
_SHEET_CLASS_COLS = (0.50, 0.30, 0.20)
_SHEET_TYPES_COLS = (0.64, 0.14, 0.22)
_SHEET_TAKEOFF_COLS = (0.34, 0.24, 0.14, 0.12, 0.16)
_SHEET_VAV_COLS = (0.22, 0.14, 0.16, 0.16, 0.18, 0.14)


_SHEET_PAIR_W = _SHEET_CLASS_W + _SHEET_TYPES_W   # the rest is the gap
_LEGEND_SWATCH = 0.014
_LEGEND_TEXT_X = _LEGEND_SWATCH + rja_drafting.PAD * 3.0


def _sheet_columns(headers, fractions, width, rows=(), totals_row=False):
    """[(header, width)] sharing width by fractions, every column at least
    as wide as its longest text (rja_drafting.fit_widths). No cell on the
    sheet is clipped or wrapped: a long text widens its column, and the
    sheet width is set by _sheet_width so the table still spans it (Colin,
    2026-11-09: always widen)."""
    needs, band_need = rja_drafting.column_needs(headers, rows, totals_row)
    widths = rja_drafting.fit_widths(needs, fractions, width)
    short = band_need - sum(widths, 0.0)
    if short > 0 and widths:
        widths[-1] += short
    return list(zip(headers, widths))


def _pair_columns(rep, words, width):
    """(class columns, types columns) for the by system class and diffuser
    types tables drawn side by side in width, fitted as one set so the two
    fill their shares and each column holds its text."""
    cls_h = ['System', 'CFM', words['count']]
    typ_h = [words['types'], 'Count', 'CFM']
    cls_rows = rep.get('by_class') or []
    typ_rows = rep.get('types') or []
    if cls_rows and typ_rows:
        cn, _ = rja_drafting.column_needs(cls_h, cls_rows)
        tn, _ = rja_drafting.column_needs(typ_h, typ_rows)
        fr = ([_SHEET_CLASS_W * f for f in _SHEET_CLASS_COLS] +
              [_SHEET_TYPES_W * f for f in _SHEET_TYPES_COLS])
        w = rja_drafting.fit_widths(cn + tn, fr, width * _SHEET_PAIR_W)
        return list(zip(cls_h, w[:3])), list(zip(typ_h, w[3:]))
    cls = _sheet_columns(cls_h, _SHEET_CLASS_COLS, width * _SHEET_CLASS_W, cls_rows)
    typ = _sheet_columns(typ_h, _SHEET_TYPES_COLS, width * _SHEET_TYPES_W, typ_rows)
    return cls, typ


def _esp_table_rows(unit):
    """(rows, has totals row) of one unit's External Static Pressure table."""
    rows, total_row = duct_perf_console.esp_unit_rows(unit, sheet=True)
    if total_row is not None:
        rows.append(total_row)
    return rows, total_row is not None


def _legend_lines(rep, show_r_marker_note=False):
    """The NOTE LEGEND text lines, in order: one per colour, then the plan
    mark lines. show_r_marker_note is the set of marks placed ('R', or 'R1',
    'R2' ...); True means the plain (R)."""
    words = rep.get('words')
    lines = ['{}: {}'.format(k.title(), m) for k, m in _legend_rows(words)]
    if show_r_marker_note:
        marks = (set(['R']) if show_r_marker_note is True
                 else set(show_r_marker_note))
        lines.extend(duct_perf_console.mark_lines(
            marks, (words or duct_perf_console.scope_words('rtu'))['one']))
    return lines


def _sheet_width(base_w, rep, esp_rec, vav_rep, headers, rows, legend_lines):
    """Sheet table width: the flagged duct table's own width, widened until
    every table and legend line holds its longest text (Colin, 2026-11-09:
    always widen, never clip, never wrap)."""
    need = rja_drafting.table_need
    w = base_w
    words = rep.get('words') or duct_perf_console.scope_words('rtu')
    if rep.get('equipment'):
        heads = duct_perf_console.equipment_headers(words)
        w = max(w, need(heads, [list(r[:len(heads)]) for r in rep['equipment']]))
    pair = 0.0
    if rep.get('by_class'):
        pair += need(['System', 'CFM', words['count']], rep['by_class'])
    if rep.get('types'):
        pair += need([words['types'], 'Count', 'CFM'], rep['types'])
    if rep.get('by_class') and rep.get('types'):
        w = max(w, pair / _SHEET_PAIR_W)
    elif pair:
        w = max(w, pair / (_SHEET_CLASS_W if rep.get('by_class') else _SHEET_TYPES_W))
    for unit in (esp_rec or {}).get('units') or []:
        if unit.get('classes'):
            t_rows, totals = _esp_table_rows(unit)
            w = max(w, need(duct_perf_console.TAKEOFF_HEADERS, t_rows, totals))
    if vav_rep is not None:
        v_rows = duct_perf_console.vav_rows(vav_rep, sheet=True)
        if v_rows:
            w = max(w, need(duct_perf_console.VAV_HEADERS, v_rows))
    if rows:
        w = max(w, need(headers, rows))
    for line in legend_lines:
        w = max(w, _LEGEND_TEXT_X + rja_drafting.text_width(line))
    return w


def _draw_system_summary(st, view, y, total_w, rep):
    """SYSTEM SUMMARY: equipment table, by system class and diffuser types
    side by side. Its warnings are in SYSTEM CORRECTIONS. Returns the y under
    it."""
    draft = rja_drafting
    y = draft._band(st, view, 0.0, y, total_w, 'SYSTEM SUMMARY',
                    duct_perf_console.summary_band_text(rep), draft.SUB_BAND_H)
    words = rep.get('words') or duct_perf_console.scope_words('rtu')
    if rep.get('equipment'):
        heads = duct_perf_console.equipment_headers(words)
        four = len(heads) == 4
        eq_rows = [list(r[:len(heads)]) for r in rep['equipment']]
        y = draft._styled_table(
            st, view, 0.0, y, _sheet_columns(
                heads, _SHEET_EQUIP_COLS if four else _SHEET_EQUIP_COLS_3, total_w,
                eq_rows),
            eq_rows, left_cols=(0, 3) if four else (0,), fit=True)
        y -= draft.GAP
    if rep.get('caution'):
        # VAV Box scope with external static (Colin, 2026-11-09).
        y = draft._note_panel(st, view, 0.0, y, total_w, [rep['caution']])
        y -= draft.GAP
    bottoms = []
    cls_cols, typ_cols = _pair_columns(rep, words, total_w)
    if rep.get('by_class'):
        bottoms.append(draft._styled_table(
            st, view, 0.0, y, cls_cols, rep['by_class'], fit=True))
    if rep.get('types'):
        w = sum([c[1] for c in typ_cols], 0.0)
        ox = total_w - w if rep.get('by_class') else 0.0
        bottoms.append(draft._styled_table(
            st, view, ox, y, typ_cols, rep['types'], fit=True))
    if bottoms:
        y = min(bottoms) - draft.GAP
    return y


def _draw_external_static(st, view, y, total_w, esp_rec):
    """EXTERNAL STATIC PRESSURE: per unit one take-off table (supply then
    return, ending in the unit total) with where each index run ends and the
    unit max check, then picked duct pairings. Its warnings and notes are in
    SYSTEM CORRECTIONS. The rows come from duct_perf_console.esp_unit_rows,
    the same rows the console table shows. Returns the y under it."""
    draft = rja_drafting
    for unit in esp_rec.get('units') or []:
        y = draft._band(st, view, 0.0, y, total_w,
                        'EXTERNAL STATIC PRESSURE - ' + unit['label'], None,
                        draft.SUB_BAND_H)
        if unit.get('classes'):
            rows, totals = _esp_table_rows(unit)
            cols = _sheet_columns(duct_perf_console.TAKEOFF_HEADERS,
                                  _SHEET_TAKEOFF_COLS, total_w, rows, totals)
            y = draft._styled_table(st, view, 0.0, y, cols, rows,
                                    totals_row=totals,
                                    left_cols=(0, 1), fit=True)
        notes = duct_perf_console.sheet_lines(
            duct_perf_console.esp_unit_notes(unit, sheet=True))
        if notes:
            y = draft._note_panel(st, view, 0.0, y, total_w, notes)
        y -= draft.GAP
    for comb in esp_rec.get('combined') or []:
        y = draft._band(st, view, 0.0, y, total_w,
                        'EXTERNAL STATIC PRESSURE - ' + comb['title'], None,
                        draft.SUB_BAND_H)
        y = draft._note_panel(st, view, 0.0, y, total_w, [comb['text']])
        y -= draft.GAP
    # No METHOD / ASSUMPTION / NOT INCLUDED block on the sheet (Colin,
    # 2026-11-02): the engineer reads the basis in the dialog's Calculation
    # basis. The console still prints it.
    return y


def _draw_vav_boxes(st, view, y, total_w, vav_rep):
    """VAV BOXES table, one row per box, with its table key (values in in.
    wc, the short cell texts). LOW boxes are in SYSTEM CORRECTIONS. Returns
    the y under it."""
    draft = rja_drafting
    rows = duct_perf_console.vav_rows(vav_rep, sheet=True)
    if not rows:
        return y
    y = draft._band(st, view, 0.0, y, total_w, 'VAV BOXES', vav_rep.get('basis'),
                    draft.SUB_BAND_H)
    y = draft._styled_table(
        st, view, 0.0, y, _sheet_columns(duct_perf_console.VAV_HEADERS,
                                         _SHEET_VAV_COLS, total_w, rows),
        rows, left_cols=(0, 1), fit=True)
    notes = duct_perf_console.sheet_lines(duct_perf_console.vav_notes(vav_rep, sheet=True))
    if notes:
        y = draft._note_panel(st, view, 0.0, y, total_w, notes)
    return y - draft.GAP


def _draw_system_corrections(st, view, y, total_w, columns, rows, left_cols,
                             corrections):
    """SYSTEM CORRECTIONS: the flagged duct table (columns [(header,
    width)], rows already in keynote order), then every correction of the
    run in one block per group, ACTION REQUIRED then NOTES (Colin,
    2026-11-02). Nothing at all: one 'No corrections' line. Returns the y
    under it."""
    draft = rja_drafting
    y = draft._band(st, view, 0.0, y, total_w, 'SYSTEM CORRECTIONS',
                    '{} duct(s)'.format(len(rows)) if rows else None,
                    draft.SUB_BAND_H)
    groups = duct_perf_console.corrections_sheet(corrections)
    if rows:
        y = draft._styled_table(st, view, 0.0, y, columns, rows, left_cols=left_cols,
                                fit=True)
    else:
        y = draft._note_panel(st, view, 0.0, y, total_w, [
            duct_perf_console.NO_FLAGGED if groups
            else duct_perf_console.NO_CORRECTIONS])
    for heading, lines in groups:
        y -= draft.GAP / 2.0
        draft._t(st, view, 0.0, y, heading, 'bold')
        y -= draft.ROW_H
        y = draft._note_panel(st, view, 0.0, y, total_w, lines)
    return y - draft.GAP


def _build_summary_view(doc, rep, esp_rec, vav_rep, flagged_rows, selected_cols,
                        source_sheet_num, tn_type_id, ts, fill_id,
                        show_r_marker_note=False, corrections=()):
    """Create a Drafting View in the RJA table look: the System Summary,
    External Static Pressure and VAV Boxes sections drawn as tables from the
    same records the console renders (rep, esp_rec, vav_rep; the last two
    None when not computed), then SYSTEM CORRECTIONS (the flagged-duct table
    and every correction line of the run, corrections from main()) and the
    note legend (Colin, 2026-10-23: the summary looks as good as the rest of
    the sheet; 2026-11-02: one corrections block, no notes after a section).

    flagged_rows: list of per-duct cell dicts from _row_cells(), already in
    the order they should appear, the same order the numbered keynote circles
    were placed in the plan view, so row 3 here is circle 3 there.
    selected_cols: set of _COLUMN_DEFS keys the user kept in the dialog.
    tn_type_id: kept for the caller; the RJA table text types are used.

    Returns (ViewDrafting, total_content_height_ft, table_width_ft), or
    (None, 0.0, 0.0) on failure. The width is returned because the column
    selection changes it on every run, and the viewport is positioned by its
    centre: the caller needs it to keep the table's left edge on its anchor.
    """
    try:
        sched_view = rja_drafting._new_drafting_view(
            doc, 'Duct Schedule - DV-{} - {}'.format(source_sheet_num, ts))
        if sched_view is None:
            return None, 0.0, 0.0
        st = rja_drafting._Style(doc)

        # Columns come from the shared _COLUMN_DEFS, filtered to the user's
        # selection: same list, same order, same headers as the console table.
        cols = _selected_columns(selected_cols)
        # sum(..., 0.0) forces float: sum([]) returns int 0 in Python 2.7
        base_w = sum([c[2] for c in cols], 0.0)
        left = tuple(i for i, c in enumerate(cols) if c[0] in _LEFT_COLUMNS)
        # '#' is this table's own row index: it matches the numbered
        # keynote circle placed on the same duct in the plan view.
        rows = [[str(ri + 1) if c[0] == 'num' else cells.get(c[0], '')
                 for c in cols] for ri, cells in enumerate(flagged_rows)]
        legend_lines = _legend_lines(rep, show_r_marker_note)
        # Every table widens to its longest text and the sheet with it, no
        # cell is clipped (Colin, 2026-11-09).
        headers = [c[1] for c in cols]
        total_w = _sheet_width(base_w, rep, esp_rec, vav_rep, headers, rows,
                               legend_lines)
        columns = _sheet_columns(headers, [c[2] for c in cols], total_w, rows)

        y = rja_drafting._band(
            st, sched_view, 0.0, 0.0, total_w,
            'DUCT PERFORMANCE - DV-{}'.format(source_sheet_num), None,
            rja_drafting.BAND_H)
        y -= rja_drafting.GAP

        # Each section fails soft on its own: a record shape this code did
        # not expect loses that section, never the sheet.
        sections = [(_draw_system_summary, rep)]
        if esp_rec is not None:
            sections.append((_draw_external_static, esp_rec))
        if vav_rep is not None:
            sections.append((_draw_vav_boxes, vav_rep))
        for draw, record in sections:
            try:
                y = draw(st, sched_view, y, total_w, record)
            except Exception as ex:
                hvac_graph.log.warning('Duct Performance sheet: %s not drawn: %s',
                                       draw.__name__, ex)
                y -= rja_drafting.GAP

        y = _draw_system_corrections(st, sched_view, y, total_w, columns, rows,
                                     left, corrections)

        # Legend: the plan colour as a swatch, then its meaning, boxed.
        # Named NOTE LEGEND and kept to the plan marks only (Colin, 2026-10-23).
        y = rja_drafting._band(st, sched_view, 0.0, y, total_w,
                               'NOTE LEGEND', None, rja_drafting.SUB_BAND_H)
        legend_top = y
        row_h = rja_drafting.ROW_H
        pad = rja_drafting.PAD
        swatch = _LEGEND_SWATCH
        legend = _legend_rows(rep.get('words'))
        for (color_key, _m), line in zip(legend, legend_lines):
            if st.region_type_id is not None:
                rja_drafting._legend_swatch(
                    doc, sched_view, st.region_type_id, pad * 2.0,
                    y - swatch - (row_h - swatch) / 2.0, swatch,
                    _COLOR_MAP[color_key], fill_id)
            rja_drafting._t(st, sched_view, _LEGEND_TEXT_X, y, line)
            y -= row_h
        for line in legend_lines[len(legend):]:
            # The (R) and MR lines start in the swatch column so they line
            # up with the colour squares above them.
            rja_drafting._t(st, sched_view, pad * 2.0, y, line)
            y -= row_h
        rja_drafting._outline(st, sched_view, 0.0, legend_top, total_w, y)

        return sched_view, (0.0 - y), total_w

    except Exception:
        return None, 0.0, 0.0


# ── main ───────────────────────────────────────────────────────────────────────
def _view_name(view):
    """View name via its built-in parameter (never .Name in IronPython)."""
    try:
        p = view.get_Parameter(BuiltInParameter.VIEW_NAME)
        return (p.AsString() or '') if p is not None else ''
    except Exception:
        return ''


def _level_name(level):
    try:
        p = level.get_Parameter(BuiltInParameter.DATUM_TEXT)
        return (p.AsString() or '') if p is not None else ''
    except Exception:
        return ''


def _elem_level_id(elem):
    """Level an element is drawn on: a duct's Reference Level, else LevelId."""
    try:
        lvl = elem.ReferenceLevel
        if lvl is not None:
            return lvl.Id
    except Exception:
        pass
    try:
        lid = elem.LevelId
        if lid is not None and lid != ElementId.InvalidElementId:
            return lid
    except Exception:
        pass
    return None


LEVEL_SAME_ELEV_FT = 0.5       # two levels this close are the same floor


def _canon_level_pick(elev, plan_levels):
    """The floor an element on a level with no plan is drawn on (Colin,
    2026-11-11, TB-M12: 85 ducts on a leftover 'LEVEL 0' at LEVEL 1's
    elevation): the plan level at the same elevation, else the nearest plan
    level below, else None. plan_levels: [(int level id, elevation ft)].
    Pure, for the CPython test."""
    same = [(abs(e - elev), lid) for lid, e in plan_levels
            if abs(e - elev) <= LEVEL_SAME_ELEV_FT]
    if same:
        return min(same)[1]
    below = [(e, lid) for lid, e in plan_levels if e <= elev + LEVEL_SAME_ELEV_FT]
    if below:
        return max(below)[1]
    return None


STANDARD_PLAN_SCALES = (24, 32, 48, 64, 96, 128, 192)   # 1/2" to 1/16"
LEVEL_PLAN_TARGET_IN = 24.0    # sheet inches a multi floor plan's width fills


def _fit_plan_scale(width_ft, current, target_in=None, scales=None):
    """The largest drawing (smallest scale number) whose width fits target_in
    sheet inches, never smaller than the current scale (Colin, 2026-11-11:
    the two story plans read too small). Pure, for the CPython test."""
    target_in = LEVEL_PLAN_TARGET_IN if target_in is None else target_in
    scales = STANDARD_PLAN_SCALES if scales is None else scales
    for s in sorted(scales):
        if s >= current:
            break
        if float(width_ft) * 12.0 / s <= target_in:
            return s
    return current


def _fit_level_view_scale(view):
    """Enlarge a multi floor plan to fill the sheet (_fit_plan_scale). A
    view template that locks the scale is applied to the view and detached
    first, so the copy keeps its graphics. Logged; never raises."""
    try:
        cb = view.CropBox
        width = float(cb.Max.X - cb.Min.X)
        now = int(view.Scale)
    except Exception as ex:
        hvac_graph.log.warning('scale: view %s crop or scale NOT READ: %s',
                               eid_int(view.Id), ex)
        return
    fit = _fit_plan_scale(width, now)
    if fit >= now:
        hvac_graph.log.info('scale: view %s kept at 1:%d (%.1f ft wide)',
                            eid_int(view.Id), now, width)
        return
    try:
        view.Scale = fit
    except Exception:
        try:
            tid = view.ViewTemplateId
            if tid is not None and tid != ElementId.InvalidElementId:
                view.ApplyViewTemplateParameters(doc.GetElement(tid))
                view.ViewTemplateId = ElementId.InvalidElementId
            view.Scale = fit
        except Exception as ex:
            hvac_graph.log.warning('scale: view %s NOT SET to 1:%d: %s',
                                   eid_int(view.Id), fit, ex)
            return
    hvac_graph.log.info('scale: view %s 1:%d -> 1:%d (%.1f ft wide)',
                        eid_int(view.Id), now, fit, width)


def _plan_levels(sheets_by_view):
    """([(int level id, elevation ft)], {int level id: ElementId}) of the
    levels with a floor plan placed on a sheet (not a template, a demolition
    plan or a Duct Performance output)."""
    out = {}
    for v in FilteredElementCollector(doc).OfClass(ViewPlan):
        try:
            if v.IsTemplate or v.ViewType != ViewType.FloorPlan or v.GenLevel is None:
                continue
            if sheets_by_view.get(eid_int(v.Id)) is None:
                continue
            name = _view_name(v).upper()
            if 'DEMO' in name or name.startswith('DUCT PERFORMANCE'):
                continue
            out[eid_int(v.GenLevel.Id)] = (float(v.GenLevel.Elevation), v.GenLevel.Id)
        except Exception:
            continue
    return [(k, e) for k, (e, _i) in out.items()], dict((k, i) for k, (_e, i) in out.items())


def _sheets_by_view():
    """int view id -> the ViewSheet it is placed on."""
    out = {}
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for vpid in sheet.GetAllViewports():
            vp = doc.GetElement(vpid)
            if vp is not None:
                out[eid_int(vp.ViewId)] = sheet
    return out


def _rect_overlap(a, b):
    """Overlap area of two (min x, min y, max x, max y) rectangles, 0.0 when
    they do not overlap. Pure, for the CPython test."""
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


def _plan_rank(template_rank, crop_rect, sys_rect):
    """Sort key for a candidate plan on one level, or None when it does not
    qualify (Colin, 2026-11-09, TB-M12C took another bench's plan): its crop
    shall overlap that level's system. Ranked by template (2 same as the
    active view, 1 RJA Mechanical), then a cropped plan over an uncropped
    one (crop_rect None: covers everything), then the larger overlap. Pure,
    for the CPython test."""
    if template_rank <= 0:
        return None
    if crop_rect is None:
        return (template_rank, 0, 0.0)
    if sys_rect is None:
        return None
    area = _rect_overlap(crop_rect, sys_rect)
    if area <= 0.0:
        return None
    return (template_rank, 1, area)


def _view_crop_rect(view):
    """The plan's crop region as a model XY rectangle, or None when its crop
    is off (covers everything). The CropBox corners go through
    CropBox.Transform to model coordinates, as understood here, not yet
    verified live. A crop that cannot be read is logged and treated as off."""
    try:
        if not view.CropBoxActive:
            return None
        bb = view.CropBox
        tr = bb.Transform
        pts = [tr.OfPoint(XYZ(x, y, bb.Min.Z))
               for x in (bb.Min.X, bb.Max.X) for y in (bb.Min.Y, bb.Max.Y)]
    except Exception as ex:
        hvac_graph.log.warning('plan pick: view %s crop NOT READ, taken as uncropped: %s',
                               eid_int(view.Id), ex)
        return None
    return _crop_box_xy([(p.X, p.Y) for p in pts], 0.0)


def _target_plan_views(active_view, level_ids, sheets_by_view, level_rects=None):
    """One floor plan per level the traversed ductwork and terminals sit on.

    The active view covers its own level when the system is on it. Every
    other level gets a mechanical floor plan on that level that is placed on
    a sheet and whose crop overlaps that level's system (_plan_rank: same
    view template as the active view, then any 'RJA Mechanical' template,
    then cropped over uncropped, then the larger overlap). Demolition plans,
    dependent views and earlier Duct Performance outputs are never picked
    (Colin, 2026-10-16: a roof plan run stopped at the RTU boot because the
    main level was never drawn; 2026-11-09: TB-M12C took TB-M01C's plan).

    level_ids: int level id -> ElementId. level_rects: int level id -> model
    XY rectangle of that level's system (None: no overlap test).
    Returns ([(level ElementId, view, source sheet or None)], [missing level
    ElementIds]); a missing level gets a plan made in main().
    """
    level_rects = level_rects or {}
    targets = []
    active_lvl = None
    try:
        active_lvl = active_view.GenLevel.Id
    except Exception:
        pass
    if active_lvl is not None and eid_int(active_lvl) in level_ids:
        sheet = sheets_by_view.get(eid_int(active_view.Id))
        targets.append((active_lvl, active_view, sheet))

    candidates = {}
    for v in FilteredElementCollector(doc).OfClass(ViewPlan):
        try:
            if v.IsTemplate or v.ViewType != ViewType.FloorPlan or v.GenLevel is None:
                continue
        except Exception:
            continue
        lid = eid_int(v.GenLevel.Id)
        if lid not in level_ids:
            continue
        sh = sheets_by_view.get(eid_int(v.Id))
        if sh is None:
            continue
        name = _view_name(v).upper()
        if 'DEMO' in name or name.startswith('DUCT PERFORMANCE'):
            continue
        # Dependent views and anything Revit will not duplicate would throw
        # inside the transaction and roll back the whole run.
        try:
            if v.GetPrimaryViewId() != ElementId.InvalidElementId:
                continue
            if not v.CanViewBeDuplicated(ViewDuplicateOption.Duplicate):
                continue
        except Exception:
            continue
        # Mechanical plans only: a level with nothing but a plumbing or
        # lighting plan is reported as missing, never colored on that plan.
        rank = 0
        try:
            tid = v.ViewTemplateId
            if tid == active_view.ViewTemplateId and tid != ElementId.InvalidElementId:
                rank = 2
            elif tid != ElementId.InvalidElementId and \
                    _view_name(doc.GetElement(tid)).upper().startswith('RJA MECHANICAL'):
                rank = 1
        except Exception:
            pass
        key = _plan_rank(rank, _view_crop_rect(v), level_rects.get(lid))
        hvac_graph.log.info('plan pick: level %s view %s (%s) key %r', lid, eid_int(v.Id),
                            name, key)
        if key is None:
            continue
        best = candidates.get(lid)
        if best is None or key > best[0]:
            candidates[lid] = (key, v, sh)

    missing = []
    for lid, lvl_eid in sorted(level_ids.items()):
        if any(t[0] is not None and eid_int(t[0]) == lid for t in targets):
            continue
        found = candidates.get(lid)
        if found is None:
            missing.append(lvl_eid)
        else:
            targets.append((lvl_eid, found[1], found[2]))
    return targets, missing


def _create_level_plan(like_view, lvl_eid):
    """A new floor plan on level lvl_eid with like_view's view family type
    and view template, for a level the system reaches that has no
    mechanical plan (Colin, 2026-11-09). None when Revit refuses (logged).
    ViewPlan.Create(doc, view family type id, level id) is the API as
    understood here; not yet verified live."""
    try:
        v = ViewPlan.Create(doc, like_view.GetTypeId(), lvl_eid)
    except Exception as ex:
        hvac_graph.log.warning('level plan: ViewPlan.Create on level %s FAILED: %s',
                               eid_int(lvl_eid), ex)
        return None
    try:
        tid = like_view.ViewTemplateId
        if tid != ElementId.InvalidElementId:
            v.ViewTemplateId = tid
    except Exception as ex:
        hvac_graph.log.warning('level plan: template on view %s NOT SET: %s',
                               eid_int(v.Id), ex)
    hvac_graph.log.info('level plan: view %s made on level %s', eid_int(v.Id),
                        eid_int(lvl_eid))
    return v


def _elem_box_points(elems):
    """Model XY corners of each element's bounding box, skipping any that
    has none."""
    pts = []
    for el in elems or ():
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is None:
            continue
        for x in (bb.Min.X, bb.Max.X):
            for y in (bb.Min.Y, bb.Max.Y):
                pts.append((x, y))
    return pts


def _crop_box_xy(points, margin_ft):
    """(min x, min y, max x, max y) around (x, y) points, margin_ft out on
    every side, or None for no points. Pure, for the CPython test."""
    pts = list(points)
    if not pts:
        return None
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return (min(xs) - margin_ft, min(ys) - margin_ft,
            max(xs) + margin_ft, max(ys) + margin_ft)


def _crop_to_elements(view, elems, margin_ft):
    """Crop a copied plan to its level's system, margin_ft out (Colin,
    2026-11-09: 5 ft). The corners of each element's bounding box are taken
    into the view's own crop coordinates (CropBox.Transform), and the crop's
    Z is kept. A scope box on the copy is cleared first, since I believe it
    holds the crop. Any failure is logged and the source crop is kept."""
    try:
        old = view.CropBox
        inv = old.Transform.Inverse
    except Exception as ex:
        hvac_graph.log.warning('crop: view %s CropBox NOT READ: %s', eid_int(view.Id), ex)
        return False
    pts = []
    for el in elems:
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is None:
            continue
        for x in (bb.Min.X, bb.Max.X):
            for y in (bb.Min.Y, bb.Max.Y):
                q = inv.OfPoint(XYZ(x, y, bb.Min.Z))
                pts.append((q.X, q.Y))
    box = _crop_box_xy(pts, margin_ft)
    if box is None:
        hvac_graph.log.info('crop: view %s, no element boxes, crop kept', eid_int(view.Id))
        return False
    try:
        p = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if p is not None and not p.IsReadOnly and p.AsElementId() != ElementId.InvalidElementId:
            p.Set(ElementId.InvalidElementId)
            hvac_graph.log.info('crop: view %s scope box cleared', eid_int(view.Id))
    except Exception as ex:
        hvac_graph.log.warning('crop: view %s scope box NOT CLEARED: %s', eid_int(view.Id), ex)
    try:
        nb = BoundingBoxXYZ()
        nb.Transform = old.Transform
        nb.Min = XYZ(box[0], box[1], old.Min.Z)
        nb.Max = XYZ(box[2], box[3], old.Max.Z)
        view.CropBoxActive = True
        view.CropBox = nb
        hvac_graph.log.info('crop: view %s cropped to %r (%d elements)', eid_int(view.Id),
                            box, len(elems))
        return True
    except Exception as ex:
        hvac_graph.log.warning('crop: view %s crop NOT SET, source crop kept: %s',
                               eid_int(view.Id), ex)
        return False


LEVEL_VP_GAP_FT = 0.1           # sheet feet between level plans
LEVEL_LABEL_UP_FT = 0.03        # sheet feet from a plan's top to its level label


def _column_centers(first_center_y, heights, gap_ft):
    """Viewport center Y for boxes stacked top to bottom, the first centered
    at first_center_y, each gap_ft under the one before (Colin, 2026-11-11:
    each extra floor goes underneath). Pure, for the CPython test."""
    out = []
    bottom = None
    for i, h in enumerate(heights):
        if i == 0:
            cy = first_center_y
        else:
            cy = bottom - gap_ft - h / 2.0
        out.append(cy)
        bottom = cy - h / 2.0
    return out


def _layout_level_viewports(sheet, vps, tn_type_id):
    """Stack the level plans top to bottom from the first one's fixed
    center, each extra floor underneath the one before (Colin, 2026-11-11),
    and put each level's name above its plan. Overlap with the title block
    or the table is accepted: the sheet is not printed. vps: [(Viewport,
    level name)] lowest level first, so Floor 1 is on top."""
    doc.Regenerate()
    sizes = []
    for vp, _name in vps:
        try:
            o = vp.GetBoxOutline()
            sizes.append((o.MaximumPoint.X - o.MinimumPoint.X,
                          o.MaximumPoint.Y - o.MinimumPoint.Y))
        except Exception as ex:
            hvac_graph.log.warning('layout: viewport %s outline NOT READ: %s',
                                   eid_int(vp.Id), ex)
            sizes.append((1.0, 1.0))
    try:
        c0 = vps[0][0].GetBoxCenter()
    except Exception:
        c0 = XYZ(1.201, 1.212, 0)
    # The label above each plan sits in the gap: leave it room.
    ys = _column_centers(c0.Y, [h for _w, h in sizes],
                         LEVEL_VP_GAP_FT + LEVEL_LABEL_UP_FT)
    for (vp, name), cy, (w, h) in zip(vps, ys, sizes):
        cx = c0.X
        try:
            vp.SetBoxCenter(XYZ(cx, cy, 0))
        except Exception as ex:
            hvac_graph.log.warning('layout: viewport %s NOT MOVED: %s', eid_int(vp.Id), ex)
        if tn_type_id is not None and name:
            try:
                TextNote.Create(doc, sheet.Id,
                                XYZ(cx - w / 2.0, cy + h / 2.0 + LEVEL_LABEL_UP_FT, 0),
                                name.upper(), TextNoteOptions(tn_type_id))
            except Exception as ex:
                hvac_graph.log.warning('layout: level label %r NOT PLACED: %s', name, ex)


def _file_like_source_sheet(new_sheet, source_sheet):
    """File the new sheet beside its source sheet. See
    revit_helpers.file_like_source_sheet."""
    err = revit_helpers.file_like_source_sheet(doc, new_sheet, source_sheet)
    if err:
        output.print_md(':warning: Could not file sheet {} in the sheet browser: {}'.format(
            new_sheet.SheetNumber, err))
