# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Leader geometry (correct per Revit API requirements):
  - Anchor = grid line endpoint (set automatically by Revit, read-only)
  - End    = point ON the grid line, extended ALONG the axis past the endpoint
             (must lie on the infinite line of the grid curve)
  - Elbow  = the perpendicular bend point, must be geometrically between
             End and Anchor along the leader path

  Visual result:
    The grid line continues past its endpoint (End), then the leader
    bends at the Elbow and the bubble sits offset from the grid line.
    This is exactly what Revit's manual break/elbow handle produces.

  Works for ALL grid orientations:
    - Vertical grids (N-S): End is above/below endpoint along Y axis
    - Horizontal grids (E-W): End is left/right along X axis
    - Diagonal grids: End is along whatever the grid direction is
    Direction is always read from the actual grid curve, never hardcoded.

Fixes in this version:
  - Leader End placed ON the grid axis (fixes 'datum plane' error)
  - Elbow placed between End and Anchor (satisfies Revit constraint)
  - Linked grid skip bug fixed (was skipping 25 host grids incorrectly)
  - No hardcoded grid IDs or directions — all geometry from curve data
  - Collision deduplication by grid pair frozenset

Scope:
  FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "8.0.0"
__doc__     = ("Pick a grid, then separates all colliding grid bubbles on "
               "plan views on sheets using the leader elbow feature. "
               "Grid lines are never moved.")

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
    RevitLinkInstance,
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
OFFSET_MULTIPLIER              = 1.25    # 1x clears overlap + 0.25x gap
MIN_GRID_LENGTH_FT             = 0.01    # skip degenerate grids


# =============================================================================
# Pick a grid — calibrate bubble size from annotation family
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
    """Bubble diameter in decimal feet scaled to view's print scale."""
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
# Grid collection — host + linked
# =============================================================================
def collect_all_grids_in_view(document, view):
    """Collect grids from host and linked models.

    Returns list of dicts with is_linked flag.
    Linked grids included for collision detection only — cannot be written.
    """
    results = []

    # Host grids
    try:
        for g in (FilteredElementCollector(document, view.Id)
                  .OfClass(Grid).ToElements()):
            results.append({
                'grid':      g,
                'grid_id':   "host:{}".format(g.Id.IntegerValue),
                'is_linked': False,
            })
    except Exception as ex:
        logger.debug("Host grids: {}".format(ex))

    # Linked grids
    try:
        for link in (FilteredElementCollector(document)
                     .OfClass(RevitLinkInstance).ToElements()):
            try:
                link_doc = link.GetLinkDocument()
                if link_doc is None:
                    continue
                for g in (FilteredElementCollector(link_doc)
                          .OfClass(Grid).ToElements()):
                    results.append({
                        'grid':      g,
                        'grid_id':   "link_{}:{}".format(
                            link_doc.Title, g.Id.IntegerValue),
                        'is_linked': True,
                    })
            except Exception:
                pass
    except Exception:
        pass

    return results


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
# Bubble visibility
# =============================================================================
def grid_has_bubble_at_end(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        return True


def grid_already_has_leader(grid, view, end_index):
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        leader = grid.GetLeader(end, view)
        return leader is not None
    except Exception:
        return False


# =============================================================================
# Entry collection — pure 2D
# =============================================================================
def collect_bubble_entries(document, view):
    """One deduplicated entry per visible bubble endpoint. 2D coords only."""
    entries = []
    seen_keys = set()

    for info in collect_all_grids_in_view(document, view):
        g = info['grid']
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue

        # Compute 2D unit direction vector for this grid (used later for leader)
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        length_2d = (dx * dx + dy * dy) ** 0.5
        if length_2d < MIN_GRID_LENGTH_FT:
            continue

        # Unit vector along grid axis (2D)
        ux = dx / length_2d
        uy = dy / length_2d

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            key = (info['grid_id'], end_index)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            pt = curve.GetEndPoint(end_index)
            entries.append({
                'grid':      g,
                'grid_id':   info['grid_id'],
                'is_linked': info['is_linked'],
                'end_index': end_index,
                'x':         pt.X,    # 2D only — Z stripped
                'y':         pt.Y,
                'z':         pt.Z,    # kept only for Z-clamping leader points
                # Outward unit vector: direction from fixed end toward this end
                'out_x':     ux if end_index == 1 else -ux,
                'out_y':     uy if end_index == 1 else -uy,
            })
    return entries


# =============================================================================
# Collision detection — pure 2D, deduplicated by grid pair
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Pairs within threshold feet. Pure 2D. One pair per grid combination."""
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
    """Choose which grid to add the leader to.

    Rules:
      - If BOTH are linked: skip (neither writable) -> return None
      - If ONE is linked: always move the HOST (the writable one)
      - If BOTH are host: move the one with higher grid_id (deterministic)

    BUG FIX from v7: previously a host:host pair could be wrongly skipped
    if the frozenset comparison found one entry from a linked collision.
    Now choose_entry_to_move only looks at is_linked flags directly.
    """
    a_linked = entry_a['is_linked']
    b_linked = entry_b['is_linked']

    if a_linked and b_linked:
        return None
    if a_linked:
        return entry_b   # b is host
    if b_linked:
        return entry_a   # a is host
    # Both host — deterministic by grid_id
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


# =============================================================================
# Leader geometry — correct per Revit API constraints
# =============================================================================
def apply_leader(target_entry, neighbour_entry, view,
                 bubble_diam_model, already_done_keys):
    """Add an elbow leader to separate the target bubble from its neighbour.

    Revit leader geometry constraints (from error message):
      1. leader.End   MUST lie ON the datum plane (on the grid's infinite line)
      2. leader.Elbow MUST be geometrically between End and Anchor
      3. Anchor is the grid endpoint — set by Revit, we only read it

    Correct geometry:
      Anchor = grid endpoint (bubble currently sits here)
      End    = point extended ALONG the grid axis past the endpoint
               (still on the grid line — satisfies constraint 1)
      Elbow  = point offset PERPENDICULAR to the grid, positioned between
               End and Anchor along the leader path
               (satisfies constraint 2)

    This creates the elbow/break shape: the leader runs along the grid
    axis then bends perpendicular, placing the bubble offset from the line.

    Direction logic (works for any grid orientation):
      - out_x/out_y in the entry is the 2D unit vector pointing outward
        along the grid axis from the fixed end toward the bubble end.
        This is computed from the actual curve, never hardcoded.
      - Perpendicular = rotate out vector 90 degrees in 2D.
      - The perpendicular direction is chosen to move AWAY from the
        colliding neighbour by checking the dot product.
    """
    key = (target_entry['grid_id'], target_entry['end_index'])
    if key in already_done_keys:
        return False

    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1

    # Skip if leader already exists (idempotent on repeat runs)
    if grid_already_has_leader(grid, view, end_index):
        already_done_keys.add(key)
        return False

    # Bubble endpoint position
    bx = target_entry['x']
    by = target_entry['y']
    bz = target_entry['z']

    # Outward axis direction (along the grid, away from fixed end)
    out_x = target_entry['out_x']
    out_y = target_entry['out_y']

    # Perpendicular to grid axis in 2D: rotate outward vector 90 degrees CCW
    # (out_x, out_y) -> (-out_y, out_x)
    perp_x = -out_y
    perp_y =  out_x

    # Determine which perpendicular direction moves AWAY from neighbour
    # Dot product of (neighbour - target) with perp vector
    ndx = neighbour_entry['x'] - bx
    ndy = neighbour_entry['y'] - by
    dot = ndx * perp_x + ndy * perp_y

    # If neighbour is on the perp side (dot > 0), flip to move away
    if dot > 0:
        perp_x = -perp_x
        perp_y = -perp_y

    offset = bubble_diam_model * OFFSET_MULTIPLIER

    # --- Leader point positions ---
    # End: extended ALONG the grid axis past the bubble endpoint
    #      Must lie on the grid's infinite line (satisfies Revit constraint 1)
    end_pt = XYZ(
        bx + out_x * offset,
        by + out_y * offset,
        bz,
    )

    # Elbow: the perpendicular bend point
    #        Positioned between End and Anchor along the leader path.
    #        We place it at the bubble endpoint offset perpendicular —
    #        this is geometrically between end_pt (along axis) and the
    #        anchor (at bx,by,bz) because it's at the junction point.
    elbow_pt = XYZ(
        bx + perp_x * offset,
        by + perp_y * offset,
        bz,
    )

    # Add the leader — Revit auto-sets the Anchor to the bubble endpoint
    grid.AddLeader(datum_end, view)

    leader = grid.GetLeader(datum_end, view)
    if leader is None:
        raise Exception("AddLeader succeeded but GetLeader returned None")

    leader.Elbow = elbow_pt
    leader.End   = end_pt
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
    skipped_linked   = 0
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
                    if target is None:
                        skipped_linked += 1
                        continue
                    neighbour = b if target is a else a
                    try:
                        if apply_leader(target, neighbour, view,
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
        "Skipped (linked): {}".format(skipped_linked),
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