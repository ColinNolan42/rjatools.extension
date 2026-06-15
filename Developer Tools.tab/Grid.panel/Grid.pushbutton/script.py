# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

VERSION 20.0.0 — root cause fix:
  clamp_elbow_to_segment() now decomposes the proposed Elbow into
  ALONG-segment and PERPENDICULAR components separately.
  Only the along-segment component is clamped to t in [0.05, 0.95].
  The perpendicular component (the actual nudge direction) is preserved
  completely. This means a DOWN nudge on a horizontal grid correctly
  moves the Elbow DOWN rather than being projected to zero movement.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "20.0.0"
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
MAX_ITERATIONS             = 50
MAX_PASSES                 = 20
DEGENERATE_THRESHOLD       = 1e-4


# =============================================================================
# Name sort helpers
# =============================================================================
def name_sort_key(name):
    """3 < 3.1 < 3.3 < 4,  H < H.1 < H.2 < I — universal, no hardcoding."""
    key = []
    for seg in str(name).split("."):
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
# Reference grid / bubble diameter
# =============================================================================
def pick_reference_grid():
    try:
        from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

        class GridFilter(ISelectionFilter):
            def AllowElement(self, e):
                return isinstance(e, Grid)
            def AllowReference(self, r, p):
                return False

        forms.alert(
            "Click any grid line to calibrate bubble size.\n"
            "The script will then process all plan views on sheets.",
            title="Separate Grid Bubbles — Pick a Grid",
            ok=True,
        )
        ref  = uidoc.Selection.PickObject(
            ObjectType.Element, GridFilter(), "Click any grid line")
        elem = doc.GetElement(ref.ElementId)
        if isinstance(elem, Grid):
            return elem
        forms.alert("Selected element is not a grid. Cancelled.",
                    title="Invalid Selection")
        return None
    except Exception:
        return None


def read_bubble_diameter_ft(grid):
    try:
        gt = doc.GetElement(grid.GetTypeId())
        if gt is not None:
            for pname in ("End 1 Default Grid Head",
                          "End 2 Default Grid Head",
                          "Default Grid Head"):
                p = gt.LookupParameter(pname)
                if p and p.HasValue:
                    sym = doc.GetElement(p.AsElementId())
                    if sym is None:
                        continue
                    for rname in ("Circle Radius", "Head Radius",
                                  "Radius", "Bubble Radius"):
                        rp = sym.LookupParameter(rname)
                        if rp and rp.HasValue:
                            d = rp.AsDouble() * 2.0
                            if 0.01 < d < 10.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} ft**".format(d))
                                return d
    except Exception as ex:
        logger.debug("read_bubble_diameter_ft: {}".format(ex))

    output.print_md("Could not read bubble diameter from annotation family.")
    try:
        raw = forms.ask_for_string(
            default="2.0",
            prompt=("Enter the grid bubble diameter in MODEL SPACE FEET.\n"
                    "Common value: 2.0 ft (1/4\" at 1/8\" scale)."),
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
    ids = set()
    for sheet in (FilteredElementCollector(document)
                  .OfClass(ViewSheet).ToElements()):
        try:
            for vid in sheet.GetAllPlacedViews():
                ids.add(vid.IntegerValue)
        except Exception:
            pass
    return ids


def collect_plan_views_on_sheets(document):
    sheet_ids = get_sheet_view_ids(document)
    result = []
    for v in FilteredElementCollector(document).OfClass(View):
        if v.IsTemplate:
            continue
        if v.ViewType not in PLAN_VIEW_TYPES:
            continue
        if v.Id.IntegerValue not in sheet_ids:
            continue
        result.append(v)
    return result


# =============================================================================
# Grid curve helpers
# =============================================================================
def get_grid_curve_in_view(grid, view):
    for et in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(et, view)
            if curves:
                return curves[0]
        except Exception:
            continue
    return None


def grid_has_bubble_at_end(grid, view, end_index):
    try:
        de = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(de, view)
    except Exception:
        return True


def grid_has_leader_at_end(grid, view, end_index):
    try:
        de = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.GetLeader(de, view) is not None
    except Exception:
        return False


# =============================================================================
# Canonicalized nudge direction
# =============================================================================
def get_nudge_direction(grid, view):
    """nudge = (tan_y, -tan_x, 0) with canonicalized tangent.
    vertical → RIGHT (+X),  horizontal → DOWN (-Y).
    """
    try:
        curve = get_grid_curve_in_view(grid, view)
        if curve is None:
            return XYZ(1.0, 0.0, 0.0)
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx, dy = p1.X - p0.X, p1.Y - p0.Y
        L = (dx*dx + dy*dy) ** 0.5
        if L < 1e-9:
            return XYZ(1.0, 0.0, 0.0)
        tx, ty = dx/L, dy/L
        if abs(ty) > abs(tx):
            if ty < 0.0: tx, ty = -tx, -ty
        else:
            if tx < 0.0: tx, ty = -tx, -ty
        return XYZ(ty, -tx, 0.0)
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
        for ei in (0, 1):
            de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, ei):
                continue
            try:
                ldr = g.GetLeader(de, view)
                pt  = ldr.Anchor if (ldr and ldr.Anchor) else curve.GetEndPoint(ei)
                positions.append((g, de, ei, pt))
            except Exception:
                continue
    return positions


# =============================================================================
# Collision detection
# =============================================================================
def find_colliding_pairs(positions, threshold):
    pairs = []
    thr2  = threshold * threshold
    n     = len(positions)
    for i in range(n):
        for j in range(i+1, n):
            if positions[i][0].Id == positions[j][0].Id:
                continue
            p1, p2 = positions[i][3], positions[j][3]
            dx, dy = p1.X-p2.X, p1.Y-p2.Y
            if dx*dx + dy*dy <= thr2:
                pairs.append((positions[i], positions[j]))
    return pairs


# =============================================================================
# Elbow placement — decomposed along/perp, clamp only along component
# =============================================================================
def place_elbow(proposed_x, proposed_y, anchor, end):
    """Compute a valid Elbow position from a proposed XY point.

    Decomposes the proposed point into:
      - ALONG component: projection onto the Anchor→End segment (clamped
        to t in [0.05, 0.95] to satisfy Revit's "between" constraint)
      - PERPENDICULAR component: offset from the segment line (preserved
        exactly — this is the actual nudge movement)

    Result: segment_point_at_t_clamped + perpendicular_offset

    For a horizontal grid nudged DOWN:
      - Along component stays at t=0.05 (just inside Anchor end)
      - Perpendicular component = full DOWN offset
      → Elbow moves correctly DOWN, Revit constraint satisfied.

    For a vertical grid nudged RIGHT:
      - Along component stays at t=0.05 (just inside Anchor end)
      - Perpendicular component = full RIGHT offset
      → Elbow moves correctly RIGHT, Revit constraint satisfied.
    """
    ax, ay = anchor.X, anchor.Y
    ex, ey = end.X,    end.Y
    sx, sy = ex - ax,  ey - ay
    ss     = sx*sx + sy*sy

    if ss < 1e-12:
        # Degenerate segment — just return proposed point at anchor Z
        return XYZ(proposed_x, proposed_y, anchor.Z)

    # Parameter of proposed point projected onto Anchor→End line
    t_raw = ((proposed_x - ax) * sx +
             (proposed_y - ay) * sy) / ss

    # Clamp t strictly inside the segment
    t_clamped = max(0.05, min(0.95, t_raw))

    # Point on segment at clamped t
    seg_x = ax + t_clamped * sx
    seg_y = ay + t_clamped * sy

    # Perpendicular offset = proposed point minus its projection on the line
    # (uses raw t, not clamped, to get true perp offset)
    proj_x  = ax + t_raw * sx
    proj_y  = ay + t_raw * sy
    perp_x  = proposed_x - proj_x
    perp_y  = proposed_y - proj_y

    return XYZ(seg_x + perp_x, seg_y + perp_y, anchor.Z)


# =============================================================================
# Leader repair: remove stale/degenerate leader and re-add fresh
# =============================================================================
def repair_leader(grid, datum_end, end_index, view):
    try:
        grid.RemoveLeader(datum_end, view)
    except Exception as ex:
        logger.debug("RemoveLeader failed grid {} end {}: {}".format(
            grid.Id.IntegerValue, end_index, ex))
        return False
    try:
        grid.AddLeader(datum_end, view)
        return True
    except Exception as ex:
        logger.debug("Re-AddLeader failed grid {} end {}: {}".format(
            grid.Id.IntegerValue, end_index, ex))
        return False


# =============================================================================
# Single pass
# =============================================================================
def run_pass(grids, view, threshold, nudge_step, existing_leader_keys, pass_num):
    new_leader_keys = set()
    errors          = []
    name_map        = {g.Id.IntegerValue: g.Name for g in grids}

    # Phase A: detect collisions, add leaders to higher-named movers only
    positions = collect_bubble_positions(grids, view)
    pairs     = find_colliding_pairs(positions, threshold)

    if not pairs:
        return new_leader_keys, errors, False

    output.print_md("  **Pass {} — {} collision(s)**".format(
        pass_num+1, len(pairs)))

    pos_lookup   = {(g.Id.IntegerValue, ei): (g, de)
                    for g, de, ei, _ in positions}
    needs_leader = set()

    for pos_a, pos_b in pairs:
        g_a, end_a, idx_a, _ = pos_a
        g_b, end_b, idx_b, _ = pos_b
        na = name_map.get(g_a.Id.IntegerValue, "")
        nb = name_map.get(g_b.Id.IntegerValue, "")
        if higher_name(na, nb):
            needs_leader.add((g_a.Id.IntegerValue, idx_a))
            output.print_md("    `{}` vs `{}` → mover `{}`".format(na, nb, na))
        else:
            needs_leader.add((g_b.Id.IntegerValue, idx_b))
            output.print_md("    `{}` vs `{}` → mover `{}`".format(na, nb, nb))

    for key in needs_leader:
        if key in existing_leader_keys:
            continue
        if key not in pos_lookup:
            continue
        g, de = pos_lookup[key]
        ei    = key[1]
        if grid_has_leader_at_end(g, view, ei):
            continue
        try:
            g.AddLeader(de, view)
            new_leader_keys.add(key)
            output.print_md(
                "    AddLeader → `{}` End{}".format(g.Name, ei))
        except Exception as ex:
            logger.debug("AddLeader grid {} end {}: {}".format(
                g.Id.IntegerValue, ei, ex))

    doc.Regenerate()

    # Phase B: iterative nudge
    for iteration in range(MAX_ITERATIONS):
        positions = collect_bubble_positions(grids, view)
        pairs     = find_colliding_pairs(positions, threshold)
        if not pairs:
            output.print_md(
                "  Pass {} resolved after {} iteration(s)".format(
                    pass_num+1, iteration))
            break

        targets = {}
        for pos_a, pos_b in pairs:
            g_a, end_a, idx_a, _ = pos_a
            g_b, end_b, idx_b, _ = pos_b
            na = name_map.get(g_a.Id.IntegerValue, "")
            nb = name_map.get(g_b.Id.IntegerValue, "")
            if higher_name(na, nb):
                mg, me, mi, mn = g_a, end_a, idx_a, na
            else:
                mg, me, mi, mn = g_b, end_b, idx_b, nb
            nd  = get_nudge_direction(mg, view)
            key = (mg.Id.IntegerValue, mi)
            if key not in targets:
                targets[key] = [mg, me, 0.0, 0.0, mn]
            targets[key][2] += nd.X
            targets[key][3] += nd.Y

        sorted_targets = sorted(
            targets.items(),
            key=lambda item: name_sort_key(item[1][4]),
            reverse=True,
        )

        for key, td in sorted_targets:
            move_grid = td[0]
            move_end  = td[1]
            move_idx  = key[1]
            move_name = td[4]
            net_x, net_y = td[2], td[3]

            net_len = (net_x*net_x + net_y*net_y) ** 0.5
            if net_len < 1e-9:
                continue
            nx = (net_x / net_len) * nudge_step
            ny = (net_y / net_len) * nudge_step

            try:
                leader = move_grid.GetLeader(move_end, view)
                if leader is None:
                    continue

                anchor = leader.Anchor
                elbow  = leader.Elbow
                end    = leader.End

                # Detect degenerate leader (Anchor == Elbow)
                ae_dist = ((anchor.X-elbow.X)**2 +
                           (anchor.Y-elbow.Y)**2) ** 0.5
                if ae_dist < DEGENERATE_THRESHOLD:
                    repaired = repair_leader(move_grid, move_end,
                                            move_idx, view)
                    if not repaired:
                        continue
                    doc.Regenerate()
                    leader = move_grid.GetLeader(move_end, view)
                    if leader is None:
                        continue
                    anchor = leader.Anchor
                    elbow  = leader.Elbow
                    end    = leader.End

                # Proposed new Elbow = current Elbow + nudge step
                prop_x = elbow.X + nx
                prop_y = elbow.Y + ny

                # Place Elbow: preserve perpendicular nudge, clamp
                # only the along-segment component to stay in (0.05, 0.95)
                new_elbow = place_elbow(prop_x, prop_y, anchor, end)

                # Skip if no meaningful movement
                if (abs(new_elbow.X - elbow.X) < 1e-6 and
                        abs(new_elbow.Y - elbow.Y) < 1e-6):
                    continue

                leader.Elbow = new_elbow

                try:
                    move_grid.SetLeader(move_end, view, leader)

                except Exception:
                    # Stale leader — End off axis. Repair and retry.
                    repaired = repair_leader(move_grid, move_end,
                                            move_idx, view)
                    if not repaired:
                        continue
                    doc.Regenerate()
                    leader2 = move_grid.GetLeader(move_end, view)
                    if leader2 is None:
                        continue
                    anchor2 = leader2.Anchor
                    elbow2  = leader2.Elbow
                    end2    = leader2.End
                    new_elbow2 = place_elbow(
                        elbow2.X + nx, elbow2.Y + ny, anchor2, end2)
                    if (abs(new_elbow2.X - elbow2.X) < 1e-6 and
                            abs(new_elbow2.Y - elbow2.Y) < 1e-6):
                        continue
                    try:
                        leader2.Elbow = new_elbow2
                        move_grid.SetLeader(move_end, view, leader2)
                    except Exception as ex2:
                        errors.append(
                            "Retry SetLeader grid {} iter {}: {}".format(
                                move_grid.Id.IntegerValue, iteration, ex2))

            except Exception as ex:
                errors.append("Nudge grid {} iter {}: {}".format(
                    move_grid.Id.IntegerValue, iteration, ex))
                logger.debug(traceback.format_exc())

    return new_leader_keys, errors, True


# =============================================================================
# Per-view processing — outer pass loop
# =============================================================================
def process_view(view, bubble_diam_ft, threshold):
    total_added = 0
    all_errors  = []

    try:
        grids = list(FilteredElementCollector(doc, view.Id)
                     .OfClass(Grid).ToElements())
    except Exception as ex:
        all_errors.append("Collect grids: {}".format(ex))
        return total_added, all_errors

    if len(grids) < 2:
        return total_added, all_errors

    nudge_step       = threshold / 8.0
    seen_leader_keys = set()

    output.print_md("#### View: `{}`".format(view.Name))

    for pass_num in range(MAX_PASSES):
        new_keys, errors, had = run_pass(
            grids, view, threshold, nudge_step, seen_leader_keys, pass_num)

        all_errors.extend(errors)
        total_added      += len(new_keys)
        seen_leader_keys.update(new_keys)

        if not had:
            output.print_md(
                "  → clean after {} pass(es)".format(pass_num+1))
            break
        if not new_keys:
            output.print_md(
                "  → resolved after {} pass(es)".format(pass_num+1))
            break
    else:
        output.print_md(
            "  → hit MAX_PASSES ({}) — some collisions may remain".format(
                MAX_PASSES))

    return total_added, all_errors


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

    output.print_md("## Grid Bubble Separation — v20.0")
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
                added, errs = process_view(view, bubble_diam_ft, bubble_diam_ft)
                total_leaders   += added
                views_processed += 1
                for e in errs:
                    all_errors.append((view.Name, e))
            except Exception as ex:
                all_errors.append((view.Name, str(ex)))
                logger.debug(traceback.format_exc())

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