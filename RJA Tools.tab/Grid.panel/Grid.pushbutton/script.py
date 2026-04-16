# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Correct leader geometry (proven by diagnostic):
  Revit's AddLeader places Anchor, Elbow, and End automatically in a
  valid default position on the grid axis. The only valid modification
  is to extend both End AND Elbow further along the grid axis together.
  Moving End alone fails. Moving perpendicular fails.

  Strategy:
    1. Call AddLeader (or skip if leader already exists)
    2. Read the default Anchor, Elbow, End from GetLeader
    3. Compute the grid axis unit vector from the curve
    4. Extend End by the offset distance along the axis
    5. Extend Elbow by half the offset distance along the axis
       (keeps Elbow between Anchor and End)
    6. Call SetLeader with the updated positions

  This is the exact pattern that matches what Revit does manually
  when you drag the elbow/break handle on a grid bubble.

Host grids only:
  Linked grids are excluded from collision detection entirely.
  They cannot have leaders added and were causing 25 false skips.

Idempotent:
  Grids that already have a leader get their leader repositioned
  via SetLeader only — AddLeader is not called again.

Scope:
  FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "9.0.0"
__doc__     = ("Pick a grid, then separates all colliding grid bubbles on "
               "plan views on sheets using the leader elbow feature. "
               "Host grids only. Grid lines are never moved.")

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------
# Handles
# -----------------------------------------------------------------------------
doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

DEFAULT_BUBBLE_DIAMETER_INCHES = 0.375   # 3/8" standard Revit grid head
OFFSET_MULTIPLIER              = 1.25    # 1x clears + 0.25x visible gap
MIN_GRID_LENGTH_FT             = 0.01


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


def read_bubble_diameter_inches(grid):
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
                            diameter_in = rp.AsDouble() * 2.0 * 12.0
                            if 0.1 < diameter_in < 2.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} in**".format(diameter_in))
                                return diameter_in
    except Exception as ex:
        logger.debug("read_bubble_diameter_inches: {}".format(ex))

    output.print_md("Using default bubble diameter: "
                    "**{} in**".format(DEFAULT_BUBBLE_DIAMETER_INCHES))
    return DEFAULT_BUBBLE_DIAMETER_INCHES


# =============================================================================
# Scale helpers
# =============================================================================
def bubble_diameter_model_units(view, bubble_inches):
    try:
        scale = float(view.Scale)
        if scale <= 0:
            scale = 96.0
    except Exception:
        scale = 96.0
    return (bubble_inches / 12.0) * scale


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


# =============================================================================
# Entry collection — HOST GRIDS ONLY, pure 2D
# =============================================================================
def collect_bubble_entries(document, view):
    """One deduplicated entry per visible bubble on host grids only.

    Linked grids excluded entirely — they cannot have leaders added and
    were causing false skips in collision pairing.

    Stores the outward axis unit vector (out_x, out_y) for each entry
    so the leader extension direction is pre-computed from the real curve.
    """
    entries = []
    seen_keys = set()

    # HOST GRIDS ONLY — no linked model lookup
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

        # Unit vector along the grid axis from End0 to End1
        ux = dx / length_2d
        uy = dy / length_2d

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            key = (g.Id.IntegerValue, end_index)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            pt = curve.GetEndPoint(end_index)

            # Outward direction: away from the fixed end toward this bubble end
            if end_index == 1:
                out_x, out_y = ux, uy     # End0 -> End1 direction
            else:
                out_x, out_y = -ux, -uy   # End1 -> End0 direction

            entries.append({
                'grid':      g,
                'grid_id':   g.Id.IntegerValue,
                'end_index': end_index,
                'x':         pt.X,        # 2D only — Z stripped
                'y':         pt.Y,
                'z':         pt.Z,        # kept for Z-clamping only
                'out_x':     out_x,       # axis direction outward
                'out_y':     out_y,
            })

    return entries


# =============================================================================
# Collision detection — pure 2D, deduplicated by grid pair
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """(entry_a, entry_b) pairs within threshold feet. Pure 2D. One per pair."""
    pairs = []
    seen_pairs = set()
    n = len(entries)
    threshold_sq = threshold * threshold

    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue

            pair_key = frozenset([entries[i]['grid_id'],
                                   entries[j]['grid_id']])
            if pair_key in seen_pairs:
                continue

            dx = entries[i]['x'] - entries[j]['x']
            dy = entries[i]['y'] - entries[j]['y']
            if (dx * dx + dy * dy) <= threshold_sq:
                seen_pairs.add(pair_key)
                pairs.append((entries[i], entries[j]))

    return pairs


def choose_entry_to_move(entry_a, entry_b):
    """Move the entry with the higher grid_id — deterministic."""
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


# =============================================================================
# Leader application — extend default geometry along axis
# =============================================================================
def apply_leader(target_entry, view, offset_distance, already_done_keys):
    """Add or reposition a leader on the target grid's bubble end.

    Proven correct geometry (from diagnostic):
      - Call AddLeader to get Revit's valid default Anchor/Elbow/End
      - Read back the default leader with GetLeader
      - Extend End by offset_distance along the grid axis
      - Extend Elbow by offset_distance * 0.5 along the grid axis
        (this keeps Elbow between Anchor and End — required by Revit)
      - Write back with SetLeader

    If the grid already has a leader (from a previous run), skip
    AddLeader and go straight to repositioning via SetLeader.

    Returns True if leader was added/updated, False if skipped.
    """
    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1

    key = (target_entry['grid_id'], end_index)
    if key in already_done_keys:
        return False

    # Axis direction for this bubble end (pre-computed, 2D)
    out_x = target_entry['out_x']
    out_y = target_entry['out_y']
    z     = target_entry['z']

    has_leader = grid_has_leader_at_end(grid, view, end_index)

    if not has_leader:
        # AddLeader creates valid default Anchor/Elbow/End on the axis
        grid.AddLeader(datum_end, view)

    # Read the current leader (default after AddLeader, or existing)
    leader = grid.GetLeader(datum_end, view)
    if leader is None:
        raise Exception("GetLeader returned None")

    # Extend End and Elbow further along the axis.
    # Both must move together — moving End alone fails Revit validation.
    # Elbow moves half as far to stay between Anchor and End.
    current_end   = leader.End
    current_elbow = leader.Elbow

    new_end = XYZ(
        current_end.X   + out_x * offset_distance,
        current_end.Y   + out_y * offset_distance,
        z,
    )
    new_elbow = XYZ(
        current_elbow.X + out_x * (offset_distance * 0.5),
        current_elbow.Y + out_y * (offset_distance * 0.5),
        z,
    )

    leader.End   = new_end
    leader.Elbow = new_elbow
    grid.SetLeader(datum_end, view, leader)

    already_done_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Check active view type -----------------------------------------
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    # ---- 2. Pick a grid to calibrate bubble size ---------------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_inches = read_bubble_diameter_inches(ref_grid)

    # ---- 3. Collect views --------------------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    # ---- 4. Stats ----------------------------------------------------------
    views_processed  = 0
    collisions_found = 0
    leaders_added    = 0
    per_view_errors  = []

    # ---- 5. Transaction ----------------------------------------------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                bubble_diam_model = bubble_diameter_model_units(
                    view, bubble_inches)
                threshold = bubble_diam_model

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                done_keys = set()
                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    try:
                        if apply_leader(target, view,
                                        bubble_diam_model, done_keys):
                            leaders_added += 1
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
        "Views processed:  {}".format(views_processed),
        "Collisions found: {}".format(collisions_found),
        "Leaders added:    {}".format(leaders_added),
        "Errors:           {}".format(len(per_view_errors)),
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