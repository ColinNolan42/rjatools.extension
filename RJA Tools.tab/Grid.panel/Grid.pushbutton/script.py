# -*- coding: utf-8 -*-
"""Detects colliding grid bubbles on top-down plan views that are placed on
sheets and automatically offsets one bubble of each colliding pair
perpendicular to the gridline so the annotations no longer overlap.

Scope:
  - Only FloorPlan, CeilingPlan, AreaPlan, and EngineeringPlan view types.
  - Only views that are actually placed on at least one sheet.
  - Sections, elevations, details, 3D views, and unplaced views are skipped.

Collision detection:
  - Threshold is computed automatically per view from view.Scale and the
    standard 3/8-inch Revit grid bubble diameter — no user input required.
  - Offset distance is also computed from view scale (1.5x bubble diameter)
    so bubbles land clearly clear of each other at any sheet scale.

Curve strategy (fixes zero-collision bug):
  - GetCurvesInView(ViewSpecific) is tried first.
  - If it returns nothing (grid has never been manually adjusted in that view,
    which is the default state for most grids), falls back to
    GetCurvesInView(Model).
  - This was the root cause of the previously reported zero-collision result.

Other behaviour:
  - Only view-specific 2D extents are written back — model geometry untouched.
  - For each colliding pair the grid with the higher ElementId is moved so
    the operation is deterministic across repeated runs.
  - Single Transaction wraps everything — one Ctrl+Z undoes the whole run.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "2.0.0"
__doc__     = ("Automatically separates colliding grid bubbles on all plan "
               "views placed on sheets. No user input required — threshold "
               "and offset are calculated from each view's print scale.")

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
    Line,
    Transaction,
    DatumExtentType,
    DatumEnds,
)

from System.Collections.Generic import List
from Autodesk.Revit.DB import Curve as RvtCurve

from pyrevit import forms, script, revit

# -----------------------------------------------------------------------------
# Document handles
# -----------------------------------------------------------------------------
doc    = revit.doc
logger = script.get_logger()
output = script.get_output()

# -----------------------------------------------------------------------------
# View-type filter — top-down plan views only
# Sections, elevations, details, and 3D views are intentionally excluded.
# -----------------------------------------------------------------------------
PLAN_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
}

# -----------------------------------------------------------------------------
# Scale constants
# Revit grid bubble = 3/8" diameter at plot scale (standard annotation circle).
# Offset multiplier: 1.5x bubble diameter gives clear daylight between bubbles
# after separation at any sheet scale.
# -----------------------------------------------------------------------------
BUBBLE_DIAMETER_INCHES = 0.375   # 3/8" standard Revit grid head circle
OFFSET_MULTIPLIER      = 1.5     # offset = 1.5 × bubble diameter


# =============================================================================
# Scale helpers
# =============================================================================
def bubble_diameter_model_units(view):
    """Return the bubble diameter in Revit internal units (decimal feet).

    Revit internal units are decimal feet. view.Scale is the ratio such that
    1 paper inch = Scale model inches (e.g. Scale=96 means 1/8"=1'-0").

        model diameter = (BUBBLE_DIAMETER_INCHES / 12) × view.Scale

    Examples:
        1/8"=1'-0"  Scale=96   -> 3.0 ft in model space
        1/4"=1'-0"  Scale=48   -> 1.5 ft in model space
        1"=1'-0"    Scale=12   -> 0.375 ft in model space
    """
    try:
        scale = float(view.Scale)
    except Exception:
        scale = 96.0  # safe default — 1/8" scale
    return (BUBBLE_DIAMETER_INCHES / 12.0) * scale


def offset_distance_model_units(view):
    """Return the perpendicular offset distance in model feet.

    Set to OFFSET_MULTIPLIER × bubble_diameter so separated bubbles have
    clear visible space between them after the move.
    """
    return OFFSET_MULTIPLIER * bubble_diameter_model_units(view)


# =============================================================================
# View collection — plans on sheets only
# =============================================================================
def get_sheet_view_ids(document):
    """Return a set of ElementId integer values for every view placed on a sheet.

    Iterates all ViewSheet elements and calls GetAllPlacedViews() on each,
    which returns the ElementIds of the views in all viewports on that sheet.
    """
    placed_ids = set()
    sheets = FilteredElementCollector(document).OfClass(ViewSheet).ToElements()
    for sheet in sheets:
        try:
            for vid in sheet.GetAllPlacedViews():
                placed_ids.add(vid.IntegerValue)
        except Exception as ex:
            logger.debug("Could not read sheet {}: {}".format(sheet.Name, ex))
    return placed_ids


def collect_plan_views_on_sheets(document):
    """Return all non-template plan views that are placed on at least one sheet.

    Filters to FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan only.
    Views not placed on any sheet are excluded — they would never be printed
    and separating bubbles on them would be wasted work.
    """
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
# Grid curve retrieval — ViewSpecific with Model fallback
# =============================================================================
def get_grid_curve_in_view(grid, view):
    """Return (curve, extent_type) for a grid in a view, or (None, None).

    ROOT CAUSE FIX for the zero-collision bug:
    GetCurvesInView(DatumExtentType.ViewSpecific, view) returns an empty list
    for any grid whose 2D extents have never been manually adjusted in that
    view — which is the default state for virtually every grid in a new model.
    The previous version only tried ViewSpecific and silently skipped all
    those grids, resulting in zero entries and zero collisions detected.

    Fix: fall back to DatumExtentType.Model when ViewSpecific is empty.
    The Model curve is always present if the grid is visible in the view.
    We still write back as ViewSpecific so we only move the annotation in
    this view — the model grid line is never altered.
    """
    for extent_type in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(extent_type, view)
            if curves:
                return curves[0], extent_type
        except Exception:
            continue
    return None, None


# =============================================================================
# Bubble visibility
# =============================================================================
def grid_has_bubble_at_end(grid, view, end_index):
    """Return True if the bubble at end_index (0=start, 1=end) is visible."""
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        # On failure assume bubble is present so the grid is not silently skipped.
        return True


# =============================================================================
# Entry collection
# =============================================================================
def collect_bubble_entries(document, view):
    """Return one dict per visible bubble endpoint in this view.

    Each dict contains:
        grid        — Grid element
        grid_id     — ElementId integer value
        curve       — the line used to read the endpoint position
        end_index   — 0 or 1 (which end has the bubble)
        point       — XYZ of the bubble endpoint in model space
    """
    entries = []
    grids = FilteredElementCollector(document, view.Id).OfClass(Grid).ToElements()
    for g in grids:
        curve, _ = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue
            entries.append({
                'grid':      g,
                'grid_id':   g.Id.IntegerValue,
                'curve':     curve,
                'end_index': end_index,
                'point':     curve.GetEndPoint(end_index),
            })
    return entries


# =============================================================================
# Collision detection — view-scale-aware, 2D only
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Return list of (entry_a, entry_b) pairs whose bubble endpoints are
    within threshold model-feet of each other.

    Z component is intentionally ignored — grid bubbles only collide in the
    2D view plane. Comparing squared distance avoids a sqrt call.

    Pairs where both entries belong to the same grid are excluded — that is
    just the two ends of one grid, not a real collision.
    """
    pairs = []
    n = len(entries)
    threshold_sq = threshold * threshold
    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue
            p1 = entries[i]['point']
            p2 = entries[j]['point']
            dx = p1.X - p2.X
            dy = p1.Y - p2.Y
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((entries[i], entries[j]))
    return pairs


def choose_entry_to_move(entry_a, entry_b):
    """Always move the entry with the higher ElementId — deterministic."""
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


# =============================================================================
# Offset application
# =============================================================================
def perpendicular_offset_vector(curve, view, distance):
    """Return a vector perpendicular to the grid in the view plane, length=distance.

    view.ViewDirection × grid_direction gives a vector that is:
      - 90 degrees to the grid line
      - lying in the view's work plane (correct for plan views)
    """
    try:
        direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    except Exception:
        direction = view.RightDirection
    try:
        perp = view.ViewDirection.CrossProduct(direction).Normalize()
    except Exception:
        perp = view.RightDirection
    return perp.Multiply(distance)


def apply_offset(entry, view, offset_distance, already_moved_keys):
    """Move the bubble endpoint perpendicular to the grid and write it back
    as a ViewSpecific curve so only this view's annotation is changed.

    already_moved_keys (set of (grid_id, end_index)) prevents moving the
    same endpoint twice when a grid collides with more than one neighbour.

    Returns True if a change was written, False if skipped.
    """
    key = (entry['grid_id'], entry['end_index'])
    if key in already_moved_keys:
        return False

    grid      = entry['grid']
    curve     = entry['curve']
    end_index = entry['end_index']

    offset_vec = perpendicular_offset_vector(curve, view, offset_distance)
    bubble_pt  = curve.GetEndPoint(end_index)
    fixed_pt   = curve.GetEndPoint(1 - end_index)
    new_bubble = bubble_pt + offset_vec

    if end_index == 0:
        new_curve = Line.CreateBound(new_bubble, fixed_pt)
    else:
        new_curve = Line.CreateBound(fixed_pt, new_bubble)

    curve_list = List[RvtCurve]()
    curve_list.Add(new_curve)

    # Always write ViewSpecific — never moves the model grid line.
    grid.SetCurvesInView(DatumExtentType.ViewSpecific, view, curve_list)
    already_moved_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Collect qualifying views ----------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert(
            "No floor plan, ceiling plan, area plan, or engineering plan "
            "views placed on sheets were found in this project.",
            title="Nothing to do",
        )
        script.exit()

    # ---- 2. Stats trackers --------------------------------------------------
    views_processed  = 0
    collisions_found = 0
    collisions_fixed = 0
    per_view_errors  = []

    # ---- 3. Single transaction — one Ctrl+Z undoes everything ---------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                # Threshold and offset scale automatically with this view's
                # print scale — no user input needed.
                threshold   = bubble_diameter_model_units(view)
                offset_dist = offset_distance_model_units(view)

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                moved_keys = set()
                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    try:
                        if apply_offset(target, view, offset_dist, moved_keys):
                            collisions_fixed += 1
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
            title="Error — Grid Bubble Separation",
        )
        logger.debug(traceback.format_exc())
        script.exit()

    # ---- 4. Results summary -------------------------------------------------
    summary = "\n".join([
        "Views processed:  {}".format(views_processed),
        "Collisions found: {}".format(collisions_found),
        "Collisions fixed: {}".format(collisions_fixed),
        "Errors:           {}".format(len(per_view_errors)),
    ])

    if per_view_errors:
        output.print_md("### Grid Bubble Separation — Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee the pyRevit output window for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()