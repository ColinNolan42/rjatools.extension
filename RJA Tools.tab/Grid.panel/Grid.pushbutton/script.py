# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

HOW IT WORKS
============

Collision detection (per view):
  1. Collect all Grids visible in the view.
  2. Add a leader to every grid bubble that doesn't already have one, so
     that leader.Anchor exists and gives us a reliable 2D bubble center.
  3. doc.Regenerate() so the new Anchor/Elbow/End values are readable.
  4. Build a position list: (grid, end, end_index, anchor_xy) per visible
     bubble.
  5. Pairwise squared-distance check in 2D (Z ignored). Two bubbles on
     different grids collide if their Anchors are within the threshold
     (bubble diameter, 2.0 ft in model space).

Fixing a 2-grid collision (e.g. grids 4 and 5):
  - Higher-named grid moves (5 > 4, E > D). Lower stays on its axis.
  - Nudge direction: perpendicular to the mover's own grid axis, pointing
    from Anchor TOWARD End. This is the only direction that pulls the
    bubble back onto its own axis, which is how leader-based separation
    actually works in Revit.
  - Move Elbow by 0.25 ft (threshold / 8), CLAMPED so Elbow stays strictly
    between Anchor and End on both X and Y, inset 1/16" from the boundary.
    This clamp is the fix for "Elbow is between End and Anchor".
  - SetLeader, Regenerate, re-check. Repeat until the pair clears.

Fixing a triple collision (e.g. grids 4, 5, 6):
  - Pair check produces (4,5), (4,6), (5,6). Movers: 5, 6, 6.
    So 4 never moves, 5 moves once, 6 moves once per iteration.
  - Process HIGHEST-NAMED FIRST: 6 nudges before 5, so 6 clears space
    before 5 tries to move.
  - After both movers, Regenerate and re-check. The loop continues on
    whoever is still colliding. Extends naturally to 4+ grid collisions.

Early exit:
  - If every mover is clamped at its far Elbow boundary (no room left),
    the iteration loop exits early instead of spinning up to MAX_ITERATIONS.

KEY REVIT API FACTS (proven by diagnostic runs):
  - leader.Anchor is READ-ONLY (computed from Elbow and End).
  - leader.End must stay ON the grid's infinite axis line.
  - leader.Elbow must be geometrically BETWEEN Anchor and End.
  - AddLeader signature: grid.AddLeader(DatumEnds, View)
  - SetLeader signature: grid.SetLeader(DatumEnds, View, Leader)
  - After AddLeader, MUST call doc.Regenerate() before GetLeader/SetLeader.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "16.0.0"
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

# Bubble diameter: 2.0 ft directly in Revit model space (no view.Scale multiply)
DEFAULT_BUBBLE_DIAMETER_FT = 2.0

MIN_GRID_LENGTH_FT = 0.01
MAX_ITERATIONS     = 50   # nudge step = threshold/8 = 0.25 ft, 50 steps = 12.5 ft max

# Safety margin when clamping Elbow into the Anchor-End bounding box.
# ~1/16" in feet. Keeps Elbow strictly inside the valid region because Revit's
# "between" check can reject exact-boundary values.
CLAMP_MARGIN_FT = 1.0 / 12.0 / 16.0


# =============================================================================
# Name sort — higher number/letter = further in sequence = higher sort value
# =============================================================================
def name_sort_key(name):
    """Sort key: 4<5<6<10, A<B<C<D<E. Further in sequence = higher value."""
    parts = re.split(r'(\d+)', str(name))
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.upper()))
    return key


def higher_name(name_a, name_b):
    """Return True if name_a is further in the counting/alphabet sequence."""
    return name_sort_key(name_a) > name_sort_key(name_b)


# =============================================================================
# Pick a grid — calibrate bubble size
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
    """Read bubble diameter from grid head annotation family.
    Falls back to user input, then default 2.0 ft."""
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
                output.print_md(
                    "User-entered: **{} ft**".format(val))
                return val
    except Exception:
        pass

    output.print_md("Using default: **{} ft**".format(DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# Collision threshold — model space, no scaling
# =============================================================================
def collision_threshold(view, bubble_diameter_ft):
    return bubble_diameter_ft


# =============================================================================
# View collection
# =============================================================================
def get_sheet_view_ids(document):
    placed_ids = set()
    for sheet in FilteredElementCollector(document).OfClass(ViewSheet).ToElements():
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
# Grid curve helper
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


# =============================================================================
# Bubble helpers
# =============================================================================
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
# Nudge direction — perpendicular to axis, pointing from Anchor TOWARD End.
#
# Why this direction:
#   The default Elbow after AddLeader+Regenerate sits ON the Anchor-side
#   boundary of the valid [Anchor..End] region (Elbow.X == Anchor.X for
#   vertical grids, Elbow.Y == Anchor.Y for horizontal grids). Any direction
#   that doesn't point toward End immediately pushes Elbow past the Anchor
#   boundary, violating "Elbow is between End and Anchor" on the first
#   SetLeader call.
#
#   Deriving the direction from the live Anchor->End vector (rather than a
#   fixed compass heading based on axis tangent) guarantees the step always
#   points INTO the valid region. It also naturally produces the correct
#   visual effect: the bubble moves back toward its own grid axis, which is
#   the only way leader-based elbow nudging actually separates bubbles.
# =============================================================================
def compute_nudge_direction(leader, grid, view):
    """Unit XY vector perpendicular to the axis, pointing Anchor->End side."""
    curve = get_grid_curve_in_view(grid, view)
    if curve is None:
        return XYZ(0.0, 0.0, 0.0)

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    ax_dx = p1.X - p0.X
    ax_dy = p1.Y - p0.Y
    ax_len = (ax_dx * ax_dx + ax_dy * ax_dy) ** 0.5
    if ax_len < 1e-9:
        return XYZ(0.0, 0.0, 0.0)

    ax_x = ax_dx / ax_len
    ax_y = ax_dy / ax_len

    # Perpendicular candidate (90 deg CCW of axis tangent).
    perp_x = -ax_y
    perp_y = ax_x

    # Flip if it doesn't point from Anchor toward End.
    try:
        anchor = leader.Anchor
        end    = leader.End
    except Exception:
        return XYZ(perp_x, perp_y, 0.0)

    ae_x = end.X - anchor.X
    ae_y = end.Y - anchor.Y
    if ae_x * perp_x + ae_y * perp_y < 0.0:
        perp_x = -perp_x
        perp_y = -perp_y

    # Renormalize (perp was already unit-length, but cheap insurance).
    p_len = (perp_x * perp_x + perp_y * perp_y) ** 0.5
    if p_len < 1e-9:
        return XYZ(0.0, 0.0, 0.0)
    return XYZ(perp_x / p_len, perp_y / p_len, 0.0)


# =============================================================================
# Safe Elbow nudge — clamps Elbow into the Anchor-End bounding box.
#
# Returns True if the Elbow actually moved and SetLeader succeeded.
# Returns False if the Elbow is pinned at its boundary (no room to move in
# the requested direction) or SetLeader threw despite clamping. A False
# return from every mover in an iteration is the signal to stop iterating.
# =============================================================================
def nudge_elbow_safe(grid, datum_end, view, nudge_dir, nudge_step, errors):
    try:
        leader = grid.GetLeader(datum_end, view)
    except Exception as ex:
        errors.append("GetLeader grid {}: {}".format(
            grid.Id.IntegerValue, ex))
        return False
    if leader is None:
        return False

    try:
        anchor = leader.Anchor
        elbow  = leader.Elbow
        end    = leader.End
    except Exception as ex:
        errors.append("Read leader grid {}: {}".format(
            grid.Id.IntegerValue, ex))
        return False

    # Proposed new Elbow (before clamping).
    prop_x = elbow.X + nudge_dir.X * nudge_step
    prop_y = elbow.Y + nudge_dir.Y * nudge_step

    # Valid region = axis-aligned bounding box of Anchor and End, inset by
    # CLAMP_MARGIN_FT on each side so we never land exactly on the boundary
    # (Revit's "between" check rejects exact equality in some cases).
    min_x = min(anchor.X, end.X) + CLAMP_MARGIN_FT
    max_x = max(anchor.X, end.X) - CLAMP_MARGIN_FT
    min_y = min(anchor.Y, end.Y) + CLAMP_MARGIN_FT
    max_y = max(anchor.Y, end.Y) - CLAMP_MARGIN_FT

    # If Anchor and End are coincident on one axis (margin collapsed the
    # range), pin to the midpoint so we don't flip the ordering.
    if max_x < min_x:
        min_x = max_x = (anchor.X + end.X) / 2.0
    if max_y < min_y:
        min_y = max_y = (anchor.Y + end.Y) / 2.0

    # Clamp proposal into the valid region.
    if   prop_x < min_x: prop_x = min_x
    elif prop_x > max_x: prop_x = max_x
    if   prop_y < min_y: prop_y = min_y
    elif prop_y > max_y: prop_y = max_y

    # If clamping left the Elbow where it was, this grid has no room left
    # in the requested direction.
    if abs(prop_x - elbow.X) < 1e-9 and abs(prop_y - elbow.Y) < 1e-9:
        return False

    leader.Elbow = XYZ(prop_x, prop_y, elbow.Z)
    try:
        grid.SetLeader(datum_end, view, leader)
        return True
    except Exception as ex:
        errors.append(
            "SetLeader grid {} elbow=({:.4f},{:.4f}) anchor=({:.4f},{:.4f}) "
            "end=({:.4f},{:.4f}): {}".format(
                grid.Id.IntegerValue,
                prop_x, prop_y,
                anchor.X, anchor.Y,
                end.X, end.Y,
                ex))
        return False


# =============================================================================
# Bubble position collection — uses Anchor after leaders exist
# =============================================================================
def collect_bubble_positions(grids, view):
    """(grid, datum_end, end_index, anchor_pt) for all visible bubbles.
    Uses leader.Anchor if available, else curve endpoint."""
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
                if leader is not None and leader.Anchor is not None:
                    pt = leader.Anchor
                else:
                    pt = curve.GetEndPoint(end_index)
                positions.append((g, datum_end, end_index, pt))
            except Exception:
                continue
    return positions


# =============================================================================
# Collision detection — pairwise squared distance, 2D only.
# =============================================================================
def find_colliding_anchor_pairs(positions, threshold):
    """Colliding (pos_a, pos_b) pairs using Anchor XY distance. Pure 2D."""
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
# Per-view processing
# =============================================================================
def process_view(view, bubble_diam_ft, threshold):
    leaders_added = 0
    errors        = []

    try:
        grids = list(FilteredElementCollector(doc, view.Id)
                     .OfClass(Grid).ToElements())
    except Exception as ex:
        errors.append("Collect grids: {}".format(ex))
        return leaders_added, errors

    if len(grids) < 2:
        return leaders_added, errors

    # --- Step 1: AddLeader on every grid bubble that doesn't have one -------
    for g in grids:
        for end_index in (0, 1):
            datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, end_index):
                continue
            if grid_has_leader_at_end(g, view, end_index):
                continue
            try:
                g.AddLeader(datum_end, view)
                leaders_added += 1
            except Exception as ex:
                logger.debug("AddLeader grid {} end {}: {}".format(
                    g.Id.IntegerValue, end_index, ex))

    # --- Step 2: REQUIRED Regenerate before any Anchor/Elbow/End read -------
    doc.Regenerate()

    # --- Step 3: Iteratively nudge colliding bubbles ------------------------
    nudge_step = threshold / 8.0
    name_map   = {g.Id.IntegerValue: g.Name for g in grids}

    for iteration in range(MAX_ITERATIONS):
        positions = collect_bubble_positions(grids, view)
        pairs     = find_colliding_anchor_pairs(positions, threshold)

        if not pairs:
            break

        # For each colliding pair, mark the HIGHER-named grid as the mover.
        # Lower-named stays on axis. In a triple collision (4,5,6) this
        # means 4 never moves; 5 and 6 each move once per iteration.
        # A grid that appears in multiple pairs only moves once per
        # iteration (keyed by grid id + end index).
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

            key = (move_g.Id.IntegerValue, move_idx)
            if key not in targets:
                targets[key] = (move_g, move_end, move_name)

        # Process HIGHEST-named FIRST so the furthest grid clears space
        # before lower-named grids attempt to move.
        sorted_targets = sorted(
            targets.items(),
            key=lambda item: name_sort_key(item[1][2]),
            reverse=True
        )

        any_progress = False
        for _key, target_data in sorted_targets:
            move_g, move_end, _move_name = target_data

            try:
                leader = move_g.GetLeader(move_end, view)
            except Exception as ex:
                errors.append("GetLeader iter {} grid {}: {}".format(
                    iteration, move_g.Id.IntegerValue, ex))
                continue
            if leader is None:
                continue

            nudge_dir = compute_nudge_direction(leader, move_g, view)
            if abs(nudge_dir.X) < 1e-9 and abs(nudge_dir.Y) < 1e-9:
                continue

            if nudge_elbow_safe(move_g, move_end, view,
                                nudge_dir, nudge_step, errors):
                any_progress = True

        # Refresh Anchor/Elbow/End before the next collision check.
        doc.Regenerate()

        # If every mover is pinned at its far boundary, further iterations
        # will keep flagging the same collisions with no way to resolve
        # them. Stop early.
        if not any_progress:
            break

    return leaders_added, errors


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

    output.print_md("## Grid Bubble Separation")
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
                threshold = collision_threshold(view, bubble_diam_ft)
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