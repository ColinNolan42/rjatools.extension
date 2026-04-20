# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

VERSION 16.0.0 — fixes:
  - AddLeader only added to the higher-named grid in a colliding pair
  - In triple/N-way collisions only the single highest-named grid moves per
    iteration; lower grids are always stationary
  - Tangent vector canonicalized so direction never depends on which endpoint
    Revit stores as index 0 vs index 1 on the curve
  - Elbow clamped to the true parametric segment [Anchor..End] (not bounding
    box) so the "Elbow is between End and Anchor" constraint is always
    satisfied
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

DEFAULT_BUBBLE_DIAMETER_FT = 2.0
MIN_GRID_LENGTH_FT         = 0.01
MAX_ITERATIONS             = 50   # 50 x 0.25 ft = 12.5 ft max travel


# =============================================================================
# Name sort helpers
# =============================================================================
def name_sort_key(name):
    """4 < 5 < 6 < 10,  A < B < C — further in sequence = higher value."""
    parts = re.split(r'(\d+)', str(name))
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.upper()))
    return key


def higher_name(name_a, name_b):
    """True if name_a is further in the counting / alphabet sequence."""
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
    """Read bubble diameter from the annotation family; fall back to user
    input; fall back to default 2.0 ft."""
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
    """Return the unit vector that higher-named grids should move toward.

    Formula: nudge = (tan_y, -tan_x, 0)
      vertical grid   tan≈(0,±1,0) → canonical (0,+1,0) → nudge=(+1, 0,0) RIGHT
      horizontal grid tan≈(±1,0,0) → canonical (+1,0,0) → nudge=(0, -1,0) DOWN

    The tangent is CANONICALIZED before the formula is applied so that the
    result never depends on which endpoint Revit labels index 0 vs index 1.
    Canonicalization rule: flip the sign of the whole tangent if its primary
    component is negative, so the primary component is always positive.
    """
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

        # --- Canonicalize so primary direction is always positive ---
        # Vertical grid: |tan_y| > |tan_x|  → primary axis is Y
        #   If tan_y < 0 the curve was stored bottom→top; flip both components.
        # Horizontal grid: |tan_x| >= |tan_y| → primary axis is X
        #   If tan_x < 0 the curve was stored right→left; flip both components.
        if abs(tan_y) > abs(tan_x):
            # Vertical
            if tan_y < 0.0:
                tan_x = -tan_x
                tan_y = -tan_y
        else:
            # Horizontal
            if tan_x < 0.0:
                tan_x = -tan_x
                tan_y = -tan_y

        # nudge = perpendicular rotated 90° clockwise from canonical tangent
        # (tan_y, -tan_x, 0)
        #   vertical   (0,+1) → (+1, 0) = RIGHT (+X) ✓
        #   horizontal (+1,0) → ( 0,-1) = DOWN  (-Y) ✓
        return XYZ(tan_y, -tan_x, 0.0)

    except Exception:
        return XYZ(1.0, 0.0, 0.0)


# =============================================================================
# Bubble position collection
# =============================================================================
def collect_bubble_positions(grids, view):
    """Return list of (grid, datum_end, end_index, anchor_pt).

    Uses leader.Anchor when a leader exists (post-Regenerate accurate
    position), otherwise falls back to the curve endpoint.
    """
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
    """Return list of (pos_a, pos_b) whose Anchor XY distance ≤ threshold."""
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
# Elbow clamping — true parametric segment, not bounding box
# =============================================================================
def clamp_elbow_to_segment(proposed_x, proposed_y, anchor, end):
    """Project the proposed Elbow XY onto the segment [anchor..end] and
    return a point that is strictly BETWEEN them (not at either endpoint).

    The Revit constraint is: Elbow must lie between Anchor and End on the
    leader path.  The safest geometric interpretation is that the Elbow,
    when projected onto the line through Anchor and End, must have a
    parameter t strictly in (0, 1).

    We project the proposed point, clamp t to [margin, 1-margin], and
    return the clamped point in 3-D (Z inherited from Anchor).
    """
    # Vector from anchor to end
    ax, ay = anchor.X, anchor.Y
    ex, ey = end.X,    end.Y
    seg_x  = ex - ax
    seg_y  = ey - ay
    seg_len_sq = seg_x * seg_x + seg_y * seg_y

    if seg_len_sq < 1e-12:
        # Degenerate segment — return midpoint
        return XYZ((ax + ex) * 0.5, (ay + ey) * 0.5, anchor.Z)

    # Parameter t of the proposed point projected onto anchor→end
    t = ((proposed_x - ax) * seg_x +
         (proposed_y - ay) * seg_y) / seg_len_sq

    # Keep t strictly inside — 5% inset from each end
    margin = 0.05
    if t < margin:
        t = margin
    elif t > 1.0 - margin:
        t = 1.0 - margin

    return XYZ(ax + t * seg_x, ay + t * seg_y, anchor.Z)


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

    # ------------------------------------------------------------------
    # Step 1: Identify colliding pairs and add leaders ONLY to the
    # higher-named grid of each colliding pair.
    #
    # Approach:
    #   a) Collect bubble positions using curve endpoints (leaders not yet
    #      present so Anchor is not available yet).
    #   b) Find all colliding pairs.
    #   c) For each pair the higher-named grid is the "mover"; add a leader
    #      to that grid/end if it doesn't already have one.
    #   d) Lower-named grids never receive a leader from this script.
    # ------------------------------------------------------------------
    initial_positions = collect_bubble_positions(grids, view)
    initial_pairs     = find_colliding_pairs(initial_positions, threshold)

    if not initial_pairs:
        return leaders_added, errors

    # Collect (grid_id, end_index) keys that need a leader
    needs_leader = set()
    name_map     = {g.Id.IntegerValue: g.Name for g in grids}

    for pos_a, pos_b in initial_pairs:
        g_a, end_a, idx_a, _ = pos_a
        g_b, end_b, idx_b, _ = pos_b
        name_a = name_map.get(g_a.Id.IntegerValue, "")
        name_b = name_map.get(g_b.Id.IntegerValue, "")

        if higher_name(name_a, name_b):
            needs_leader.add((g_a.Id.IntegerValue, idx_a))
        else:
            needs_leader.add((g_b.Id.IntegerValue, idx_b))

    # Build a lookup from (grid_id, end_index) → (grid, datum_end)
    pos_lookup = {}
    for g, datum_end, end_index, _ in initial_positions:
        pos_lookup[(g.Id.IntegerValue, end_index)] = (g, datum_end)

    for key in needs_leader:
        if key not in pos_lookup:
            continue
        g, datum_end = pos_lookup[key]
        end_index    = key[1]
        if grid_has_leader_at_end(g, view, end_index):
            continue   # already has one from a previous run
        try:
            g.AddLeader(datum_end, view)
            leaders_added += 1
        except Exception as ex:
            logger.debug("AddLeader grid {} end {}: {}".format(
                g.Id.IntegerValue, end_index, ex))

    # ------------------------------------------------------------------
    # Step 2: Regenerate — REQUIRED before any GetLeader/SetLeader call
    # ------------------------------------------------------------------
    doc.Regenerate()

    # ------------------------------------------------------------------
    # Step 3: Iteratively nudge colliding bubbles apart.
    #
    # Rules:
    #   - Only the highest-named grid in any collision group moves per step.
    #   - Grid 4 never moves if grid 5 is in the same collision group.
    #   - Nudge step = threshold / 8 = 0.25 ft per iteration.
    #   - Elbow is clamped to the parametric segment [Anchor..End] after
    #     every nudge so SetLeader never throws a constraint error.
    # ------------------------------------------------------------------
    nudge_step = threshold / 8.0

    for iteration in range(MAX_ITERATIONS):
        positions = collect_bubble_positions(grids, view)
        pairs     = find_colliding_pairs(positions, threshold)

        if not pairs:
            break

        # Determine which grids to move this iteration.
        # For each colliding pair only the higher-named grid is a candidate.
        # Accumulate net nudge vectors per (grid_id, end_index) key.
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

        # Process highest-named first so it clears space before lower grids
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

            net_len = (net_x * net_x + net_y * net_y) ** 0.5
            if net_len < 1e-9:
                continue
            # Normalized direction × step size
            nx = (net_x / net_len) * nudge_step
            ny = (net_y / net_len) * nudge_step

            try:
                leader = move_grid.GetLeader(move_end, view)
                if leader is None:
                    continue

                anchor = leader.Anchor   # read-only, computed by Revit
                elbow  = leader.Elbow
                end    = leader.End      # must never be moved

                # Proposed new Elbow position
                prop_x = elbow.X + nx
                prop_y = elbow.Y + ny

                # Clamp to the true parametric segment [anchor..end]
                # so the "Elbow is between End and Anchor" constraint
                # is always satisfied.
                clamped = clamp_elbow_to_segment(
                    prop_x, prop_y, anchor, end)

                # Skip the API call if clamping produced no movement —
                # the grid has no remaining room in this direction.
                if (abs(clamped.X - elbow.X) < 1e-9 and
                        abs(clamped.Y - elbow.Y) < 1e-9):
                    continue

                leader.Elbow = clamped
                move_grid.SetLeader(move_end, view, leader)

            except Exception as ex:
                errors.append("Nudge grid {} iter {}: {}".format(
                    move_grid.Id.IntegerValue, iteration, ex))
                logger.debug(traceback.format_exc())

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
                threshold = bubble_diam_ft   # model space, no scale factor
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