# -*- coding: utf-8 -*-
"""Detects colliding grid bubbles across all plan/elevation/section views in the
active Revit project and automatically offsets one bubble of each colliding
pair perpendicular to the gridline so the annotations no longer overlap.

Only view-specific 2D grid extents are modified (GetCurvesInView /
SetCurvesInView). Model geometry of the grid is never altered. A single
Transaction wraps the entire run so the user can undo the operation with one
Ctrl+Z.

For each colliding pair the grid with the higher ElementId is moved, which
makes the operation deterministic across repeated runs.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "1.0.0"
__doc__     = ("Finds grid bubbles that overlap in plan, elevation and section "
               "views and offsets the bubble end of the 2D grid curve "
               "perpendicular to the gridline so the annotations are legible. "
               "User supplies the collision threshold and offset distance.")

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
import math
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Grid,
    View,
    ViewType,
    XYZ,
    Line,
    Transaction,
    DatumExtentType,
    DatumEnds,
    UnitUtils,
)

# Unit handling changed in Revit 2021 (ForgeTypeId replaced DisplayUnitType).
# Try the modern API first, fall back to legacy for Revit 2020 and earlier.
try:
    from Autodesk.Revit.DB import UnitTypeId
    _USE_FORGE_UNITS = True
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType
    _USE_FORGE_UNITS = False

from pyrevit import forms, script, revit

# -----------------------------------------------------------------------------
# Document handles
# -----------------------------------------------------------------------------
doc     = revit.doc
uidoc   = revit.uidoc
logger  = script.get_logger()
output  = script.get_output()

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
# View types we will process. 3D views, schedules, legends, drafting views,
# and browser-only views are excluded because grid bubbles do not render on
# them in the same 2D coordinate space.
PROCESSABLE_VIEW_TYPES = {
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.AreaPlan,
    ViewType.EngineeringPlan,
    ViewType.Elevation,
    ViewType.Section,
    ViewType.Detail,
}

# Tolerance used for treating two direction vectors as parallel (dot product).
PARALLEL_TOL = 1.0e-6


# =============================================================================
# Unit helpers
# =============================================================================
def is_metric_project(document):
    """Return True if the project's length display unit is metric.

    Used only to pick sensible default values for the input dialog. All
    internal math stays in Revit's internal units (decimal feet).
    """
    try:
        fmt_options = document.GetUnits().GetFormatOptions(
            __import__("Autodesk.Revit.DB", fromlist=["SpecTypeId"]).SpecTypeId.Length
            if _USE_FORGE_UNITS else
            __import__("Autodesk.Revit.DB", fromlist=["UnitType"]).UnitType.UT_Length
        )
        if _USE_FORGE_UNITS:
            unit_id = fmt_options.GetUnitTypeId()
            return unit_id != UnitTypeId.FeetFractionalInches \
               and unit_id != UnitTypeId.Feet \
               and unit_id != UnitTypeId.FeetAndFractionalInches \
               and unit_id != UnitTypeId.Inches \
               and unit_id != UnitTypeId.FractionalInches
        else:
            dut = fmt_options.DisplayUnits
            return dut not in (
                DisplayUnitType.DUT_DECIMAL_FEET,
                DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES,
                DisplayUnitType.DUT_DECIMAL_INCHES,
                DisplayUnitType.DUT_FRACTIONAL_INCHES,
            )
    except Exception:
        # On any failure assume imperial - it is the Revit internal default.
        return False


def to_internal_length(value, metric):
    """Convert a user-entered length to Revit internal units (decimal feet).

    If ``metric`` is True the value is treated as millimeters; otherwise it is
    treated as feet.
    """
    if _USE_FORGE_UNITS:
        unit = UnitTypeId.Millimeters if metric else UnitTypeId.Feet
        return UnitUtils.ConvertToInternalUnits(value, unit)
    else:
        unit = DisplayUnitType.DUT_MILLIMETERS if metric else DisplayUnitType.DUT_DECIMAL_FEET
        return UnitUtils.ConvertToInternalUnits(value, unit)


# =============================================================================
# Geometry helpers
# =============================================================================
def points_within_threshold(p1, p2, threshold):
    """Return True if two XYZ points are within ``threshold`` feet of each
    other in the view plane.

    Grid curves in a view are always flat in the view's work plane so a simple
    3D distance works and reduces correctly to 2D for plans, elevations and
    sections alike.
    """
    dx = p1.X - p2.X
    dy = p1.Y - p2.Y
    dz = p1.Z - p2.Z
    return (dx * dx + dy * dy + dz * dz) <= (threshold * threshold)


def perpendicular_offset_vector(curve, view, offset_distance):
    """Compute a vector perpendicular to the grid's 2D direction, lying in the
    view's work plane, with magnitude ``offset_distance``.

    Handles vertical (N-S), horizontal (E-W) and diagonal grids uniformly.
    """
    try:
        direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    except Exception:
        # Degenerate curve - fall back to view right direction.
        direction = view.RightDirection

    # View normal is the out-of-page direction for that view. Perpendicular in
    # the view plane = view_normal x grid_direction.
    view_normal = view.ViewDirection
    try:
        perp = view_normal.CrossProduct(direction).Normalize()
    except Exception:
        # Cross product failed (vectors parallel) - use view's right direction.
        perp = view.RightDirection

    return perp.Multiply(offset_distance)


def grid_has_bubble_at_end(grid, view, end_index):
    """Return True if the bubble at the given end (0 or 1) is visible for this
    grid in this view.

    Grids can have bubbles at one end, both, or neither, and the visibility is
    per-view. Uses the DatumEnds API which is available on DatumPlane.
    """
    try:
        end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(end, view)
    except Exception:
        # If the API call fails, assume a bubble is present so the grid is
        # still eligible for offset rather than silently skipped.
        return True


# =============================================================================
# Data collection
# =============================================================================
def collect_processable_views(document):
    """Return all non-template views whose ViewType is in
    PROCESSABLE_VIEW_TYPES.
    """
    views = []
    collector = FilteredElementCollector(document).OfClass(View)
    for v in collector:
        if v.IsTemplate:
            continue
        if v.ViewType in PROCESSABLE_VIEW_TYPES:
            views.append(v)
    return views


def collect_grid_bubble_entries(document, view):
    """For a single view, return a list of dicts describing each grid's
    bubble-bearing endpoint(s) that are visible in that view.

    Each entry contains::

        {
            'grid':       <Grid element>,
            'grid_id':    <ElementId int value>,
            'curve':      <the view-specific 2D curve>,
            'end_index':  0 or 1,      # which end the bubble is on
            'point':      <XYZ of that endpoint>,
        }

    A grid can appear twice (once per bubble end) if both bubbles are visible.
    """
    entries = []
    grids = FilteredElementCollector(document, view.Id).OfClass(Grid).ToElements()
    for g in grids:
        try:
            curves = g.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        except Exception as ex:
            logger.debug("Could not get curves for grid {} in view {}: {}"
                         .format(g.Id, view.Name, ex))
            continue
        if not curves:
            continue
        # A grid's view-specific representation is a single line in nearly all
        # cases; loop anyway for safety (multi-segment grids).
        for curve in curves:
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
# Collision resolution
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Find every pair of bubble entries within ``threshold`` feet.

    Returns a list of (entry_a, entry_b) tuples. Order inside the tuple is
    arbitrary - the caller decides which to move.
    """
    pairs = []
    n = len(entries)
    for i in range(n):
        for j in range(i + 1, n):
            if points_within_threshold(entries[i]['point'],
                                       entries[j]['point'],
                                       threshold):
                # Do not pair a grid's two bubbles with each other (same grid,
                # opposite ends). That is not a collision - it is the same
                # grid drawn across the view.
                if entries[i]['grid_id'] == entries[j]['grid_id']:
                    continue
                pairs.append((entries[i], entries[j]))
    return pairs


def choose_entry_to_move(entry_a, entry_b):
    """Return the entry of the pair with the higher ElementId.

    Deterministic so repeated runs of the script touch the same grid.
    """
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


def apply_offset(entry, view, offset_distance, already_moved_keys):
    """Offset the bubble endpoint of a grid curve perpendicular to the grid
    direction and write it back with SetCurvesInView.

    ``already_moved_keys`` is a set of (grid_id, end_index) tuples already
    offset in this view - prevents moving the same endpoint twice when a grid
    collides with more than one neighbour.
    """
    key = (entry['grid_id'], entry['end_index'])
    if key in already_moved_keys:
        return False

    grid       = entry['grid']
    curve      = entry['curve']
    end_index  = entry['end_index']

    # Compute the new endpoint.
    offset_vec = perpendicular_offset_vector(curve, view, offset_distance)
    old_point  = curve.GetEndPoint(end_index)
    new_point  = old_point + offset_vec

    # Build the replacement line. The non-bubble end stays put.
    fixed_end_index = 1 - end_index
    fixed_point     = curve.GetEndPoint(fixed_end_index)

    if end_index == 0:
        new_curve = Line.CreateBound(new_point, fixed_point)
    else:
        new_curve = Line.CreateBound(fixed_point, new_point)

    # Wrap in a .NET-compatible IList[Curve].
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import Curve
    curve_list = List[Curve]()
    curve_list.Add(new_curve)

    grid.SetCurvesInView(DatumExtentType.ViewSpecific, view, curve_list)
    already_moved_keys.add(key)
    return True


# =============================================================================
# User input
# =============================================================================
def prompt_user_for_distances(metric):
    """Show a pyrevit.forms input dialog for threshold and offset values.

    Defaults:
        Imperial: 1.5 ft threshold, 2.0 ft offset
        Metric:   450 mm threshold, 600 mm offset
    """
    if metric:
        default_threshold = "450"
        default_offset    = "600"
        unit_label        = "mm"
    else:
        default_threshold = "1.5"
        default_offset    = "2.0"
        unit_label        = "ft"

    components = [
        forms.Label("Collision threshold ({0}):".format(unit_label)),
        forms.TextBox("threshold_tb", Text=default_threshold),
        forms.Label("Offset distance ({0}):".format(unit_label)),
        forms.TextBox("offset_tb", Text=default_offset),
        forms.Button("Run"),
    ]

    try:
        form = forms.FlexForm("Separate Grid Bubbles", components)
        form.show()
        values = form.values
        if not values:
            return None, None
        threshold_val = float(values['threshold_tb'])
        offset_val    = float(values['offset_tb'])
    except Exception:
        # Fallback to a simple TaskDialog-style prompt if FlexForm misbehaves
        # in the host pyRevit build.
        threshold_raw = forms.ask_for_string(
            default=default_threshold,
            prompt="Collision threshold ({0}):".format(unit_label),
            title="Separate Grid Bubbles",
        )
        if threshold_raw is None:
            return None, None
        offset_raw = forms.ask_for_string(
            default=default_offset,
            prompt="Offset distance ({0}):".format(unit_label),
            title="Separate Grid Bubbles",
        )
        if offset_raw is None:
            return None, None
        try:
            threshold_val = float(threshold_raw)
            offset_val    = float(offset_raw)
        except ValueError:
            forms.alert("Threshold and offset must be numeric.",
                        title="Invalid input")
            return None, None

    if threshold_val <= 0 or offset_val <= 0:
        forms.alert("Threshold and offset must be positive.",
                    title="Invalid input")
        return None, None

    # Convert to Revit internal units (decimal feet).
    return (to_internal_length(threshold_val, metric),
            to_internal_length(offset_val,    metric))


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Units & user input ---------------------------------------------
    metric = is_metric_project(doc)
    threshold_ft, offset_ft = prompt_user_for_distances(metric)
    if threshold_ft is None:
        # User cancelled.
        script.exit()

    # ---- 2. Gather views ---------------------------------------------------
    views = collect_processable_views(doc)
    if not views:
        forms.alert("No plan, elevation, or section views were found.",
                    title="Nothing to do")
        script.exit()

    # ---- 3. Stats trackers -------------------------------------------------
    views_processed     = 0
    collisions_found    = 0
    collisions_fixed    = 0
    per_view_errors     = []  # list of (view_name, error_text)

    # ---- 4. Transaction ----------------------------------------------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                entries = collect_grid_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold_ft)
                collisions_found += len(pairs)

                # Track which (grid_id, end_index) we have already moved in
                # this view so a chain of colliding bubbles only gets bumped
                # once per grid.
                moved_keys = set()

                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    try:
                        if apply_offset(target, view, offset_ft, moved_keys):
                            collisions_fixed += 1
                    except Exception as ex:
                        per_view_errors.append((
                            view.Name,
                            "Grid {0}: {1}".format(target['grid_id'], ex)
                        ))
                        logger.debug(traceback.format_exc())

                views_processed += 1

            except Exception as ex:
                # One bad view must not kill the whole run.
                per_view_errors.append((view.Name, str(ex)))
                logger.debug(traceback.format_exc())
                continue

        t.Commit()

    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(
            "Transaction failed and was rolled back.\n\n{0}".format(ex),
            title="Error",
        )
        logger.debug(traceback.format_exc())
        script.exit()

    # ---- 5. Results dialog -------------------------------------------------
    summary_lines = [
        "Views processed:  {0}".format(views_processed),
        "Collisions found: {0}".format(collisions_found),
        "Collisions fixed: {0}".format(collisions_fixed),
        "Errors:           {0}".format(len(per_view_errors)),
    ]
    summary = "\n".join(summary_lines)

    if per_view_errors:
        # Route full error detail to the pyRevit output window so the user
        # can click through without a giant modal dialog.
        output.print_md("### Grid Bubble Separation - Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{0}**: {1}".format(vname, err))

        forms.alert(
            summary + "\n\nSee the pyRevit output window for error details.",
            title="Separate Grid Bubbles - Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles - Complete")


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()