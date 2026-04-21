# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

VERSION 18.1.0 — diagnostic logging embedded to trace sub-grid collision
behaviour at runtime. All logic identical to v18; output panel will print
bubble positions, collision pairs, leader additions and nudge results for
every view so the root cause of the 3.1 mis-movement can be identified.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "18.1.0"
__doc__     = ("Separates colliding grid bubbles using leader elbow nudging. "
               "Works for any grid orientation in Revit 2022-2025.")

import re
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Grid,
    View,
    ViewSheet,
    ViewType,
    XYZ,
    Transaction,
    DatumExtentType,
    DatumEnds,
)

from pyrevit import forms, script, revit

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()

PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

DEFAULT_BUBBLE_DIAMETER_FT = 2.0
MIN_GRID_LENGTH_FT         = 0.01
MAX_ITERATIONS             = 50
MAX_PASSES                 = 20


# =============================================================================
# Name sort helpers
# =============================================================================
def name_sort_key(name):
    """Universal sort key handling sub-grid dot notation.
    3 < 3.1 < 3.3 < 4,  H < H.1 < H.2 < I
    Each dot-segment parsed as (alpha_prefix, numeric_suffix).
    """
    key = []
    segments = str(name).split(".")
    for seg in segments:
        m = re.match(r'^([A-Za-z]*)(\d*)$', seg.strip())
        if m:
            alpha  = m.group(1).upper()
            digits = int(m.group(2)) if m.group(2) else 0
        else:
            alpha  = seg.upper()
            digits = 0
        key.append((alpha, digits))
    return key


def higher_name(name_a, name_b):
    return name_sort_key(name_a) > name_sort_key(name_b)


# =============================================================================
# Pick a reference grid to calibrate bubble size
# =============================================================================
def pick_reference_grid():
    try:
        from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

        class GridFilter(ISelectionFilter):
            def AllowElement(self, element):
                return isinstance(element, Grid)
            def AllowReference(self, reference, point):
                return False

        forms.alert(
            "Click any grid line to calibrate bubble size.\n"
            "The script will then process all plan views on sheets.",
            title="Separate Grid Bubbles — Pick a Grid",
            ok=True,
        )
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, GridFilter(), "Click any grid line")
        element = doc.GetElement(ref.ElementId)
        if isinstance(element, Grid):
            return element
        forms.alert("Selected element is not a grid. Cancelled.",
                    title="Invalid Selection")
        return None
    except Exception:
        return None


def read_bubble_diameter_ft(grid):
    try:
        grid_type = doc.GetElement(grid.GetTypeId())
        if grid_type is not None:
            for param_name in ("End 1 Default Grid Head",
                               "End 2 Default Grid Head",
                               "Default Grid Head"):
                p = grid_type.LookupParameter(param_name)
                if p is not None and p.HasValue:
                    head_sym = doc.GetElement(p.AsElementId())
                    if head_sym is None:
                        continue
                    for radius_name in ("Circle Radius", "Head Radius",
                                        "Radius", "Bubble Radius"):
                        rp = head_sym.LookupParameter(radius_name)
                        if rp is not None and rp.HasValue:
                            diameter_ft = rp.AsDouble() * 2.0
                            if 0.01 < diameter_ft < 10.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} ft**".format(diameter_ft))
                                return diameter_ft
    except Exception as ex:
        logger.debug("read_bubble_diameter_ft: {}".format(ex))

    output.print_md("Could not read bubble diameter from annotation family.")
    try:
        raw = forms.ask_for_string(
            default="2.0",
            prompt=("Enter the grid bubble diameter in MODEL SPACE FEET.\n"
                    "Common value: 2.0 ft (displays as 1/4\" at 1/8\" scale)."),
            title="Grid Bubble Diameter (Model Space Feet)",
        )
        if raw:
            val = float(raw)
            if 0.01 < val < 100.0:
                output.print_md("User-entered: **{} ft**".format(val))
                return val
    except Exception:
        pass

    output.print_md(
        "Using default: **{} ft**".format(DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# View collection
# =============================================================================
def get_sheet_view_ids(document):
    placed_ids = set()
    for sheet in (FilteredElementCollector(document)
                  .OfClass(ViewSheet).ToElements()):
        try:
            for vid in sheet.GetAllPlacedViews():
                placed_ids.add(vid.IntegerValue)
        except Exception:
            pass
    return placed_ids


def collect_plan_views_on_sheets(document):
    sheet_view_ids = get_sheet_view_ids(document)
    result = []
    for v in FilteredElementCollector(document).OfClass(View):
        if v.IsTemplate:
            continue
        if v.ViewType not in PLAN_VIEW_TYPES:
            continue
        if v.Id.IntegerValue not in sheet_view_ids:
            continue
        result.append(v)
    return result


# =============================================================================
# Grid curve helpers
# =============================================================================
def get_grid_curve_in_view(grid, view):
    for extent_type in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(extent_type, view)
            if curves:
                return curves[0]
        except Exception:
            continue
    return None


def grid_has_bubble_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True


def grid_has_leader_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.GetLeader(end, view) is not None
    except Exception:
        return False


# =============================================================================
# Canonicalized nudge direction
# =============================================================================
def get_nudge_direction(grid, view):
    try:
        curve = get_grid_curve_in_view(grid, view)
        if curve is None:
            return XYZ(1.0, 0.0, 0.0)

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        length = (dx * dx + dy * dy) ** 0.5
        if length < 1e-9:
            return XYZ(1.0, 0.0, 0.0)

        tan_x = dx / length
        tan_y = dy / length

        if abs(tan_y) > abs(tan_x):
            if tan_y < 0.0:
                tan_x = -tan_x
                tan_y = -tan_y
        else:
            if tan_x < 0.0:
                tan_x = -tan_x
                tan_y = -tan_y

        return XYZ(tan_y, -tan_x, 0.0)

    except Exception:
        return XYZ(1.0, 0.0, 0.0)


# =============================================================================
# Bubble position collection
# =============================================================================
def collect_bubble_positions(grids, view):
    positions = []
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        for end_index in (0, 1):
            datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, end_index):
                continue
            try:
                leader = g.GetLeader(datum_end, view)
                if leader and leader.Anchor:
                    pt = leader.Anchor
                else:
                    pt = curve.GetEndPoint(end_index)
                positions.append((g, datum_end, end_index, pt))
            except Exception:
                continue
    return positions


# =============================================================================
# Collision detection
# =============================================================================
def find_colliding_pairs(positions, threshold):
    pairs = []
    threshold_sq = threshold * threshold
    n = len(positions)
    for i in range(n):
        for j in range(i + 1, n):
            if positions[i][0].Id == positions[j][0].Id:
                continue
            p1 = positions[i][3]
            p2 = positions[j][3]
            dx = p1.X - p2.X
            dy = p1.Y - p2.Y
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((positions[i], positions[j]))
    return pairs


# =============================================================================
# Elbow clamping
# =============================================================================
def clamp_elbow_to_segment(proposed_x, proposed_y, anchor, end):
    ax, ay = anchor.X, anchor.Y
    ex, ey = end.X,    end.Y
    seg_x  = ex - ax
    seg_y  = ey - ay
    seg_len_sq = seg_x * seg_x + seg_y * seg_y

    if seg_len_sq < 1e-12:
        return XYZ((ax + ex) * 0.5, (ay + ey) * 0.5, anchor.Z)

    t = ((proposed_x - ax) * seg_x +
         (proposed_y - ay) * seg_y) / seg_len_sq

    margin = 0.05
    if t < margin:
        t = margin
    elif t > 1.0 - margin:
        t = 1.0 - margin

    return XYZ(ax + t * seg_x, ay + t * seg_y, anchor.Z)


# =============================================================================
# Diagnostic printer — called once per view before any changes
# =============================================================================
def print_view_diagnostic(grids, view, threshold):
    """Print full bubble state for every grid in this view so we can trace
    exactly what positions, ends and collisions the script sees."""
    output.print_md("#### DIAGNOSTIC — View: `{}`  threshold={:.3f} ft".format(
        view.Name, threshold))

    grids_sorted = sorted(grids, key=lambda g: g.Name)
    for g in grids_sorted:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            output.print_md("  `{}` — no curve in view".format(g.Name))
            continue

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        nudge = get_nudge_direction(g, view)

        output.print_md(
            "  `{}` | End0({:.2f},{:.2f}) End1({:.2f},{:.2f}) "
            "nudge({:.0f},{:.0f})".format(
                g.Name,
                p0.X, p0.Y, p1.X, p1.Y,
                nudge.X, nudge.Y))

        for end_index, datum_end in [(0, DatumEnds.End0), (1, DatumEnds.End1)]:
            try:
                visible = g.IsBubbleVisibleInView(datum_end, view)
            except Exception:
                visible = False
            if not visible:
                continue

            try:
                leader = g.GetLeader(datum_end, view)
                if leader and leader.Anchor:
                    pt = leader.Anchor
                    src = "Anchor"
                else:
                    pt = curve.GetEndPoint(end_index)
                    src = "CurveEP"
                output.print_md(
                    "    End{} VISIBLE | {} pos({:.2f},{:.2f})".format(
                        end_index, src, pt.X, pt.Y))
            except Exception as ex:
                output.print_md(
                    "    End{} VISIBLE | pos ERROR: {}".format(end_index, ex))

    # Print collision pairs
    positions = collect_bubble_positions(grids, view)
    pairs     = find_colliding_pairs(positions, threshold)
    if pairs:
        output.print_md("  **Collisions detected: {}**".format(len(pairs)))
        for pos_a, pos_b in pairs:
            g_a, _, idx_a, pt_a = pos_a
            g_b, _, idx_b, pt_b = pos_b
            name_a = g_a.Name
            name_b = g_b.Name
            dist = ((pt_a.X - pt_b.X)**2 + (pt_a.Y - pt_b.Y)**2) ** 0.5
            winner = name_a if higher_name(name_a, name_b) else name_b
            output.print_md(
                "    `{}` End{} vs `{}` End{} | "
                "dist={:.3f} | mover=`{}`".format(
                    name_a, idx_a, name_b, idx_b, dist, winner))
    else:
        output.print_md("  No collisions detected.")

    output.print_md("")


# =============================================================================
# Single pass: detect new collisions, add leaders, nudge to resolve
# =============================================================================
def run_pass(grids, view, threshold, nudge_step, existing_leader_keys,
             pass_num, diag):
    new_leader_keys = set()
    errors          = []
    name_map        = {g.Id.IntegerValue: g.Name for g in grids}

    # Phase A: detect collisions, add leaders only to higher-named grid
    positions = collect_bubble_positions(grids, view)
    pairs     = find_colliding_pairs(positions, threshold)

    if not pairs:
        return new_leader_keys, errors, False

    if diag:
        output.print_md("  **Pass {} — Phase A: {} collision(s)**".format(
            pass_num + 1, len(pairs)))

    pos_lookup = {}
    for g, datum_end, end_index, _ in positions:
        pos_lookup[(g.Id.IntegerValue, end_index)] = (g, datum_end)

    needs_leader = set()
    for pos_a, pos_b in pairs:
        g_a, end_a, idx_a, pt_a = pos_a
        g_b, end_b, idx_b, pt_b = pos_b
        name_a = name_map.get(g_a.Id.IntegerValue, "")
        name_b = name_map.get(g_b.Id.IntegerValue, "")

        if higher_name(name_a, name_b):
            mover_key  = (g_a.Id.IntegerValue, idx_a)
            mover_name = name_a
        else:
            mover_key  = (g_b.Id.IntegerValue, idx_b)
            mover_name = name_b

        needs_leader.add(mover_key)

        if diag:
            output.print_md(
                "    collision `{}` vs `{}` → mover=`{}`".format(
                    name_a, name_b, mover_name))

    for key in needs_leader:
        if key in existing_leader_keys:
            if diag:
                output.print_md(
                    "    grid_id={} end={} already has leader — skip AddLeader".format(
                        key[0], key[1]))
            continue
        if key not in pos_lookup:
            continue
        g, datum_end = pos_lookup[key]
        end_index    = key[1]
        if grid_has_leader_at_end(g, view, end_index):
            if diag:
                output.print_md(
                    "    `{}` End{} leader exists in model — skip AddLeader".format(
                        g.Name, end_index))
            continue
        try:
            g.AddLeader(datum_end, view)
            new_leader_keys.add(key)
            if diag:
                output.print_md(
                    "    AddLeader → `{}` End{}".format(g.Name, end_index))
        except Exception as ex:
            logger.debug("AddLeader grid {} end {}: {}".format(
                g.Id.IntegerValue, end_index, ex))

    doc.Regenerate()

    # Phase B: iteratively nudge
    for iteration in range(MAX_ITERATIONS):
        positions = collect_bubble_positions(grids, view)
        pairs     = find_colliding_pairs(positions, threshold)

        if not pairs:
            if diag:
                output.print_md(
                    "  Pass {} resolved after {} nudge iteration(s)".format(
                        pass_num + 1, iteration))
            break

        targets = {}
        for pos_a, pos_b in pairs:
            g_a, end_a, idx_a, _ = pos_a
            g_b, end_b, idx_b, _ = pos_b
            name_a = name_map.get(g_a.Id.IntegerValue, "")
            name_b = name_map.get(g_b.Id.IntegerValue, "")

            if higher_name(name_a, name_b):
                move_g, move_end, move_idx, move_name = (
                    g_a, end_a, idx_a, name_a)
            else:
                move_g, move_end, move_idx, move_name = (
                    g_b, end_b, idx_b, name_b)

            nudge_dir = get_nudge_direction(move_g, view)
            key = (move_g.Id.IntegerValue, move_idx)
            if key not in targets:
                targets[key] = [move_g, move_end, 0.0, 0.0, move_name]
            targets[key][2] += nudge_dir.X
            targets[key][3] += nudge_dir.Y

        sorted_targets = sorted(
            targets.items(),
            key=lambda item: name_sort_key(item[1][4]),
            reverse=True,
        )

        for key, target_data in sorted_targets:
            move_grid = target_data[0]
            move_end  = target_data[1]
            net_x     = target_data[2]
            net_y     = target_data[3]
            move_name = target_data[4]

            net_len = (net_x * net_x + net_y * net_y) ** 0.5
            if net_len < 1e-9:
                continue
            nx = (net_x / net_len) * nudge_step
            ny = (net_y / net_len) * nudge_step

            try:
                leader = move_grid.GetLeader(move_end, view)
                if leader is None:
                    if diag:
                        output.print_md(
                            "    iter {} `{}` End{} — GetLeader returned None".format(
                                iteration, move_name, move_end))
                    continue

                anchor = leader.Anchor
                elbow  = leader.Elbow
                end    = leader.End

                prop_x = elbow.X + nx
                prop_y = elbow.Y + ny

                clamped = clamp_elbow_to_segment(
                    prop_x, prop_y, anchor, end)

                if (abs(clamped.X - elbow.X) < 1e-9 and
                        abs(clamped.Y - elbow.Y) < 1e-9):
                    if diag:
                        output.print_md(
                            "    iter {} `{}` End{} — clamped, no room to move".format(
                                iteration, move_name, move_end))
                    continue

                if diag and iteration == 0:
                    output.print_md(
                        "    iter {} nudge `{}` End{} | "
                        "Anchor({:.2f},{:.2f}) "
                        "Elbow({:.2f},{:.2f})→({:.2f},{:.2f}) "
                        "End({:.2f},{:.2f}) "
                        "step({:.3f},{:.3f})".format(
                            iteration, move_name, move_end,
                            anchor.X, anchor.Y,
                            elbow.X, elbow.Y,
                            clamped.X, clamped.Y,
                            end.X, end.Y,
                            nx, ny))

                leader.Elbow = clamped
                move_grid.SetLeader(move_end, view, leader)

            except Exception as ex:
                errors.append("Nudge grid {} iter {}: {}".format(
                    move_grid.Id.IntegerValue, iteration, ex))
                logger.debug(traceback.format_exc())

    return new_leader_keys, errors, True


# =============================================================================
# Per-view processing — outer pass loop
# =============================================================================
def process_view(view, bubble_diam_ft, threshold):
    total_leaders_added  = 0
    all_errors           = []

    try:
        grids = list(FilteredElementCollector(doc, view.Id)
                     .OfClass(Grid).ToElements())
    except Exception as ex:
        all_errors.append("Collect grids: {}".format(ex))
        return total_leaders_added, all_errors

    if len(grids) < 2:
        return total_leaders_added, all_errors

    nudge_step           = threshold / 8.0
    all_leader_keys_seen = set()

    # Print full diagnostic snapshot before any changes
    print_view_diagnostic(grids, view, threshold)

    for pass_num in range(MAX_PASSES):
        new_keys, errors, had_collisions = run_pass(
            grids, view, threshold, nudge_step, all_leader_keys_seen,
            pass_num, diag=True)

        all_errors.extend(errors)
        total_leaders_added += len(new_keys)
        all_leader_keys_seen.update(new_keys)

        if not had_collisions:
            output.print_md(
                "View **{}**: clean after {} pass(es)".format(
                    view.Name, pass_num + 1))
            break

        if not new_keys:
            output.print_md(
                "View **{}**: resolved after {} pass(es)".format(
                    view.Name, pass_num + 1))
            break
    else:
        output.print_md(
            "View **{}**: hit MAX_PASSES ({}) — "
            "some collisions may remain".format(view.Name, MAX_PASSES))

    return total_leaders_added, all_errors


# =============================================================================
# Main
# =============================================================================
def main():
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation — v18.1 Diagnostic Run")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)

    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    views_processed = 0
    total_leaders   = 0
    all_errors      = []

    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                threshold = bubble_diam_ft
                added, errors = process_view(view, bubble_diam_ft, threshold)
                total_leaders += added
                for err in errors:
                    all_errors.append((view.Name, err))
                views_processed += 1
            except Exception as ex:
                all_errors.append((view.Name, str(ex)))
                logger.debug(traceback.format_exc())
                continue

        t.Commit()

    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Transaction failed and was rolled back.\n\n{}".format(ex),
            title="Error",
        )
        logger.debug(traceback.format_exc())
        script.exit()

    summary = "\n".join([
        "Views processed: {}".format(views_processed),
        "Leaders added:   {}".format(total_leaders),
        "Errors:          {}".format(len(all_errors)),
    ])

    output.print_md("### Results\n```\n{}\n```".format(summary))

    if all_errors:
        output.print_md("### Errors")
        for vname, err in all_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee pyRevit output for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


if __name__ == "__main__":
    main()