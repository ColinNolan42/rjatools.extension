# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Approach — incremental Elbow nudge:
  1. Detect colliding bubble pairs (2D, pure XY)
  2. Choose which grid to move (lowest alphanumeric name)
  3. Call AddLeader on that grid's bubble end
  4. Read the Elbow position from the new leader
  5. Nudge Elbow one bubble diameter per step in the separation direction
  6. Call SetLeader after each nudge (Elbow-only move is always valid)
  7. Re-check 2D distance between this bubble's End and all neighbours
  8. Stop when clear, or after MAX_NUDGES steps

Why Elbow-only works:
  SetLeader only rejects moves where End leaves the grid axis.
  Elbow is the bend point — it has no axis constraint and can be
  moved freely. End is never touched.

Nudge directions (from diagnostic — matches Revit's default geometry):
  Vertical grids:
    Default Anchor is 2ft LEFT of grid line.
    Nudge Elbow further LEFT (-X) to increase separation.

  Horizontal grids:
    Default Anchor is 2ft BELOW grid line.
    Nudge Elbow further DOWN (-Y) to increase separation.

Movement rule — lowest alphanumeric name moves:
  4 < 5 < 10,  A < B,  D < E
  No names hardcoded — computed from grid.Name at runtime.

Bubble position (bottom vs side) from Y fraction in crop box.
Host grids only. Linked grids excluded.
FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "14.0.0"
__doc__     = ("Separates colliding grid bubbles by nudging the leader "
               "elbow incrementally until clear. Lowest name moves.")

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

# Bubble diameter: 1/2" paper space = 0.04167 ft
# Multiplied by view.Scale for model-space collision threshold
DEFAULT_BUBBLE_DIAMETER_FT = 0.5 / 12.0

MIN_GRID_LENGTH_FT = 0.01
MAX_NUDGES         = 10   # max incremental steps per bubble


# =============================================================================
# Alphanumeric sort — lowest name moves
# =============================================================================
def name_sort_key(name):
    """Integers by value, letters alphabetically. 4<5<10, A<B, D<E."""
    parts = re.split(r'(\d+)', str(name))
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.upper()))
    return key


def entry_is_lower_name(entry_a, entry_b):
    return name_sort_key(entry_a['name']) < name_sort_key(entry_b['name'])


# =============================================================================
# Bubble position — bottom vs side
# =============================================================================
def get_view_bounds(view):
    try:
        if view.CropBoxActive and view.CropBox is not None:
            bb = view.CropBox
            height = bb.Max.Y - bb.Min.Y
            if height > 0.01:
                return bb.Min.Y, bb.Max.Y, height
    except Exception:
        pass
    return None, None, None


def bubble_is_vertical(entry, view):
    """True if bubble is near top/bottom of view (vertical grid).
    False if bubble is on the side (horizontal grid).
    Falls back to grid orientation from curve if crop box unavailable.
    """
    min_y, max_y, height = get_view_bounds(view)
    if height is None:
        return entry['is_vertical']
    frac_y = (entry['y'] - min_y) / height
    return frac_y < 0.35 or frac_y > 0.65


# =============================================================================
# Nudge direction
# =============================================================================
def get_nudge_direction(entry, view):
    """Return XYZ unit vector for incremental Elbow nudge.

    Vertical grid (bottom/top bubble):
      Default Anchor is LEFT of grid line. Nudge Elbow further LEFT (-X)
      to pull the bubble away from its neighbour.

    Horizontal grid (side bubble):
      Default Anchor is BELOW grid line. Nudge Elbow further DOWN (-Y)
      to pull the bubble away from its neighbour.
    """
    if bubble_is_vertical(entry, view):
        return XYZ(-1.0, 0.0, 0.0)   # LEFT
    else:
        return XYZ(0.0, -1.0, 0.0)   # DOWN


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

    output.print_md("Using default: **1/2 in ({:.5f} ft)**".format(
        DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# Collision threshold
# =============================================================================
def collision_threshold(view, bubble_diameter_ft):
    try:
        scale = float(view.Scale)
        if scale <= 0:
            scale = 96.0
    except Exception:
        scale = 96.0
    return bubble_diameter_ft * scale


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
# Bubble and leader state
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


def get_leader_end_position(grid, view, end_index):
    """Return the XY position of the leader's End point (where bubble sits).
    Falls back to the curve endpoint if no leader exists.
    """
    try:
        datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        leader = grid.GetLeader(datum_end, view)
        if leader and leader.End:
            return leader.End.X, leader.End.Y
    except Exception:
        pass
    # Fallback to curve endpoint
    curve = get_grid_curve_in_view(grid, view)
    if curve:
        pt = curve.GetEndPoint(end_index)
        return pt.X, pt.Y
    return None, None


# =============================================================================
# Entry collection — HOST GRIDS ONLY, pure 2D
# =============================================================================
def collect_bubble_entries(document, view):
    """One entry per visible bubble. Reads leader.End if leader exists
    so positions reflect current nudged state."""
    entries = []
    seen_keys = set()

    try:
        host_grids = (FilteredElementCollector(document, view.Id)
                      .OfClass(Grid).ToElements())
    except Exception:
        return entries

    for g in host_grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        length_2d = (dx * dx + dy * dy) ** 0.5
        if length_2d < MIN_GRID_LENGTH_FT:
            continue

        is_vertical = abs(dy) >= abs(dx)

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            key = (g.Id.IntegerValue, end_index)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            # Use leader.End position if leader exists (reflects nudges)
            # otherwise use curve endpoint
            bx, by = get_leader_end_position(g, view, end_index)
            if bx is None:
                continue

            # Z from curve for clamping
            pt = curve.GetEndPoint(end_index)

            entries.append({
                'grid':        g,
                'grid_id':     g.Id.IntegerValue,
                'name':        g.Name,
                'end_index':   end_index,
                'x':           bx,
                'y':           by,
                'z':           pt.Z,
                'is_vertical': is_vertical,
            })

    return entries


# =============================================================================
# Collision detection — pure 2D
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """All colliding pairs. Pure 2D XY. No dedup."""
    pairs = []
    n = len(entries)
    threshold_sq = threshold * threshold

    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue
            dx = entries[i]['x'] - entries[j]['x']
            dy = entries[i]['y'] - entries[j]['y']
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((entries[i], entries[j]))

    return pairs


def bubbles_still_colliding(entry, all_entries, threshold):
    """Check if this entry's bubble still collides with any other bubble."""
    threshold_sq = threshold * threshold
    for other in all_entries:
        if other['grid_id'] == entry['grid_id']:
            continue
        dx = entry['x'] - other['x']
        dy = entry['y'] - other['y']
        if (dx * dx + dy * dy) <= threshold_sq:
            return True
    return False


def choose_entry_to_move(entry_a, entry_b):
    """Lowest alphanumeric name moves."""
    if entry_is_lower_name(entry_a, entry_b):
        return entry_a
    elif entry_is_lower_name(entry_b, entry_a):
        return entry_b
    return entry_a if entry_a['grid_id'] < entry_b['grid_id'] else entry_b


# =============================================================================
# Nudge application
# =============================================================================
def nudge_bubble_clear(target_entry, all_entries, view,
                       threshold, bubble_step_ft, already_done_keys):
    """Add a leader then nudge Elbow incrementally until bubble clears.

    Step 1: AddLeader if not already present.
    Step 2: Read current Elbow from leader.
    Step 3: Loop up to MAX_NUDGES:
              - Move Elbow one step in nudge direction
              - SetLeader with new Elbow only (End unchanged)
              - Update entry's x,y from new leader.End position
              - Check if still colliding — if not, done
    Step 4: Mark as done regardless (prevents re-processing).

    Returns True if leader was added/nudged, False if skipped.
    """
    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
    z         = target_entry['z']

    key = (target_entry['grid_id'], end_index)
    if key in already_done_keys:
        return False

    nudge_dir = get_nudge_direction(target_entry, view)

    # Add leader if not already present
    if not grid_has_leader_at_end(grid, view, end_index):
        grid.AddLeader(datum_end, view)

    leader = grid.GetLeader(datum_end, view)
    if leader is None:
        raise Exception("GetLeader returned None after AddLeader")

    # Nudge Elbow incrementally until bubble clears or max steps reached
    for step in range(MAX_NUDGES):
        elbow = leader.Elbow
        new_elbow = XYZ(
            elbow.X + nudge_dir.X * bubble_step_ft,
            elbow.Y + nudge_dir.Y * bubble_step_ft,
            z,
        )
        leader.Elbow = new_elbow

        try:
            grid.SetLeader(datum_end, view, leader)
        except Exception as ex:
            logger.debug("SetLeader step {}: {}".format(step, ex))
            break

        # Update position in entry to reflect new leader.End
        updated_leader = grid.GetLeader(datum_end, view)
        if updated_leader and updated_leader.End:
            target_entry['x'] = updated_leader.End.X
            target_entry['y'] = updated_leader.End.Y
            leader = updated_leader

        # Check if still colliding — stop early if clear
        if not bubbles_still_colliding(target_entry, all_entries, threshold):
            break

    already_done_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Check active view ----------------------------------------------
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    # ---- 2. Pick a grid ----------------------------------------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)

    # ---- 3. Collect views --------------------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    # ---- 4. Stats ----------------------------------------------------------
    views_processed = 0
    total_leaders   = 0
    per_view_errors = []

    # ---- 5. Transaction ----------------------------------------------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                threshold    = collision_threshold(view, bubble_diam_ft)
                # Nudge step = one full bubble diameter in model space
                bubble_step  = threshold
                done_keys    = set()

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                if not pairs:
                    views_processed += 1
                    continue

                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    key = (target['grid_id'], target['end_index'])
                    if key in done_keys:
                        continue
                    try:
                        if nudge_bubble_clear(target, entries, view,
                                              threshold, bubble_step,
                                              done_keys):
                            total_leaders += 1
                    except Exception as ex:
                        per_view_errors.append((
                            view.Name,
                            "Grid {}: {}".format(target['grid_id'], ex),
                        ))
                        logger.debug(traceback.format_exc())

                views_processed += 1

            except Exception as ex:
                per_view_errors.append((view.Name, str(ex)))
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

    # ---- 6. Results --------------------------------------------------------
    summary = "\n".join([
        "Views processed: {}".format(views_processed),
        "Bubbles nudged:  {}".format(total_leaders),
        "Errors:          {}".format(len(per_view_errors)),
    ])

    output.print_md("### Results\n```\n{}\n```".format(summary))

    if per_view_errors:
        output.print_md("### Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee pyRevit output for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


if __name__ == "__main__":
    main()