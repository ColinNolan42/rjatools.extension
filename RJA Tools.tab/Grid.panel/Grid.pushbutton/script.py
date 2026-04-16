# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Workflow:
  1. User picks any one grid in the active view.
  2. Script reads the bubble diameter from that grid's annotation family.
  3. Script reads the level Z from that grid's curve (used for Z clamping).
  4. Script processes all plan views on sheets using those values.

Collision detection:
  - Pure 2D (X, Y only). Z is completely stripped before any distance
    calculation. Grid curves in a plan view all project to the same
    horizontal plane so Z carries no useful information and including it
    inflates distances, causing missed collisions.

Write strategy:
  - Promotes grid to ViewSpecific extents before writing.
  - Re-reads curve AFTER promotion (post-promotion coordinates differ).
  - Clamps each endpoint Z independently to its original value.
  - Moves strictly in XY along the grid axis — no Z drift possible.
  - SetCurveInView (singular) — correct IronPython Revit API method.
  - Linked grids detected for collision but skipped for writing (read-only).
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "6.0.0"
__doc__     = ("Pick a grid, then automatically separates all colliding grid "
               "bubbles on plan views placed on sheets. Bubble size is read "
               "from the picked grid's annotation family.")

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
    XYZ,
    Transaction,
    DatumExtentType,
    DatumEnds,
    RevitLinkInstance,
    Reference,
)

from pyrevit import forms, script, revit, HOST_APP

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

# Fallback bubble diameter if annotation family has no readable radius param
DEFAULT_BUBBLE_DIAMETER_INCHES = 0.375   # 3/8" standard Revit grid head

# Offset = 1.25x bubble diameter (1x clears overlap + 0.25x visible gap)
OFFSET_MULTIPLIER = 1.25

# Minimum 2D grid length — grids shorter than this are skipped as degenerate
MIN_GRID_LENGTH_FT = 0.01


# =============================================================================
# Step 1 — User picks a grid to calibrate bubble size and level Z
# =============================================================================
def pick_reference_grid():
    """Prompt user to click a grid in the active view.

    Returns the selected Grid element, or None if cancelled/invalid.
    Reading from the picked grid gives us:
      - The actual annotation bubble diameter for this project
      - The level Z elevation for correct Z clamping across all views
    """
    try:
        from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

        class GridFilter(ISelectionFilter):
            def AllowElement(self, element):
                return isinstance(element, Grid)
            def AllowReference(self, reference, point):
                return False

        forms.alert(
            "Click any grid line in the active view to calibrate bubble size,\n"
            "then the script will process all plan views on sheets.",
            title="Separate Grid Bubbles — Pick a Grid",
            ok=True,
        )

        ref = uidoc.Selection.PickObject(ObjectType.Element, GridFilter(),
                                         "Click any grid line")
        element = doc.GetElement(ref.ElementId)
        if isinstance(element, Grid):
            return element
        forms.alert("Selected element is not a grid. Script cancelled.",
                    title="Invalid Selection")
        return None

    except Exception:
        # User pressed Escape or selection failed
        return None


def read_bubble_diameter_inches(grid):
    """Read the grid head annotation bubble diameter from the picked grid.

    Looks up the grid's head annotation symbol and checks common parameter
    names for the circle radius. Returns diameter in inches.

    Falls back to DEFAULT_BUBBLE_DIAMETER_INCHES if not found.
    """
    try:
        # Grid head type is stored in the grid's type parameters
        grid_type = doc.GetElement(grid.GetTypeId())
        if grid_type is None:
            return DEFAULT_BUBBLE_DIAMETER_INCHES

        # The grid head family symbol is referenced by the type
        for param_name in ("End 1 Default Grid Head",
                           "End 2 Default Grid Head",
                           "Default Grid Head"):
            p = grid_type.LookupParameter(param_name)
            if p is not None and p.HasValue:
                head_id = p.AsElementId()
                head_sym = doc.GetElement(head_id)
                if head_sym is None:
                    continue
                # Look for radius parameter on the annotation symbol
                for radius_name in ("Circle Radius", "Head Radius",
                                    "Radius", "Bubble Radius",
                                    "Grid Head Radius"):
                    rp = head_sym.LookupParameter(radius_name)
                    if rp is not None and rp.HasValue:
                        radius_ft = rp.AsDouble()
                        diameter_in = radius_ft * 2.0 * 12.0
                        if 0.1 < diameter_in < 2.0:
                            output.print_md(
                                "Bubble diameter from annotation family: "
                                "**{:.4f} inches**".format(diameter_in))
                            return diameter_in
    except Exception as ex:
        logger.debug("read_bubble_diameter_inches: {}".format(ex))

    output.print_md(
        "Could not read bubble diameter from annotation family — "
        "using default **{} inches**.".format(DEFAULT_BUBBLE_DIAMETER_INCHES))
    return DEFAULT_BUBBLE_DIAMETER_INCHES


def read_reference_z(grid, view):
    """Read the Z elevation of the picked grid's curve in the active view.

    All grids in a project share the same level structure so this Z value
    is used to clamp replacement curve endpoints across all plan views at
    the same level.

    Returns the Z value in decimal feet, or 0.0 as a safe fallback.
    """
    try:
        curve = get_grid_curve_in_view(grid, view)
        if curve:
            return curve.GetEndPoint(0).Z
    except Exception as ex:
        logger.debug("read_reference_z: {}".format(ex))
    return 0.0


# =============================================================================
# Scale helpers
# =============================================================================
def bubble_diameter_model_units(view, bubble_inches):
    """Bubble diameter in decimal feet at this view's print scale.

    model_diameter = (bubble_inches / 12) x view.Scale

    view.Scale is the print scale denominator (96 = 1/8"=1'-0").
    Multiplying converts paper-space size into model-space feet.
    """
    try:
        scale = float(view.Scale)
        if scale <= 0:
            scale = 96.0
    except Exception:
        scale = 96.0
    return (bubble_inches / 12.0) * scale


def offset_distance_model_units(view, bubble_inches):
    return OFFSET_MULTIPLIER * bubble_diameter_model_units(view, bubble_inches)


# =============================================================================
# View collection — plans on sheets only
# =============================================================================
def get_sheet_view_ids(document):
    """Set of ElementId integers for every view placed on any sheet."""
    placed_ids = set()
    sheets = FilteredElementCollector(document).OfClass(ViewSheet).ToElements()
    for sheet in sheets:
        try:
            for vid in sheet.GetAllPlacedViews():
                placed_ids.add(vid.IntegerValue)
        except Exception as ex:
            logger.debug("Sheet {}: {}".format(sheet.Name, ex))
    return placed_ids


def collect_plan_views_on_sheets(document):
    """Non-template plan views placed on at least one sheet."""
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
# Grid collection — host + linked models
# =============================================================================
def collect_all_grids_in_view(document, view):
    """Grids from host model in this view, plus linked grids for detection.

    Linked grids are flagged is_linked=True and skipped during the write
    step — SetCurveInView cannot modify elements owned by a linked file.
    They are still included in collision detection so host grids move to
    clear linked ones.
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

    # Linked grids (read-only for writing, included for detection)
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
            except Exception as ex:
                logger.debug("Link grids: {}".format(ex))
    except Exception as ex:
        logger.debug("RevitLinkInstance: {}".format(ex))

    return results


# =============================================================================
# Grid curve helpers
# =============================================================================
def get_grid_curve_in_view(grid, view):
    """ViewSpecific first, Model fallback. Returns None if unavailable."""
    for extent_type in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(extent_type, view)
            if curves:
                return curves[0]
        except Exception:
            continue
    return None


def promote_and_reread(grid, view):
    """Promote to ViewSpecific extents then re-read the promoted curve.

    Must re-read AFTER SetDatumExtentType — Revit creates a new internal
    ViewSpecific curve during promotion and its endpoint coordinates are
    what we must use when building the replacement line.
    """
    try:
        grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
        grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
    except Exception as ex:
        logger.debug("SetDatumExtentType {}: {}".format(
            grid.Id.IntegerValue, ex))
    try:
        curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        if curves:
            return curves[0]
    except Exception as ex:
        logger.debug("Re-read after promote {}: {}".format(
            grid.Id.IntegerValue, ex))
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


# =============================================================================
# Entry collection
# =============================================================================
def collect_bubble_entries(document, view):
    """One dict per visible bubble endpoint in this view.

    Coordinates stored as 2D tuples (x, y) — Z is stripped here and never
    used again. All collision math is pure 2D.
    """
    entries = []
    grid_infos = collect_all_grids_in_view(document, view)

    for info in grid_infos:
        g = info['grid']
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            pt = curve.GetEndPoint(end_index)
            entries.append({
                'grid':      g,
                'grid_id':   info['grid_id'],
                'is_linked': info['is_linked'],
                'end_index': end_index,
                # 2D only — Z stripped here permanently
                'x':         pt.X,
                'y':         pt.Y,
            })
    return entries


# =============================================================================
# Collision detection — pure 2D
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Return (entry_a, entry_b) pairs whose bubble endpoints are within
    threshold feet of each other.

    PURE 2D — only X and Y are used. Z is never referenced.
    This is correct because in a plan view all grid bubbles project onto
    the same horizontal plane; Z carries no positional information and
    including it would inflate distances and cause missed detections.

    Pairs where both entries belong to the same grid_id are excluded.
    """
    pairs = []
    n = len(entries)
    threshold_sq = threshold * threshold
    for i in range(n):
        for j in range(i + 1, n):
            if entries[i]['grid_id'] == entries[j]['grid_id']:
                continue
            dx = entries[i]['x'] - entries[j]['x']
            dy = entries[i]['y'] - entries[j]['y']
            # Pure 2D squared distance — no Z component anywhere
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((entries[i], entries[j]))
    return pairs


def choose_entry_to_move(entry_a, entry_b):
    """Move the host grid with the higher grid_id string — deterministic.

    If one entry is a linked grid (read-only), always move the host grid.
    If both are linked, skip the pair (neither can be written).
    """
    a_linked = entry_a['is_linked']
    b_linked = entry_b['is_linked']

    if a_linked and b_linked:
        return None  # both linked — cannot move either
    if a_linked:
        return entry_b  # a is linked, move b (host)
    if b_linked:
        return entry_a  # b is linked, move a (host)

    # Both host — move the one with the higher grid_id (deterministic)
    return entry_a if entry_a['grid_id'] > entry_b['grid_id'] else entry_b


# =============================================================================
# Offset application
# =============================================================================
def apply_offset(entry, view, offset_distance, already_moved_keys):
    """Extend the bubble endpoint outward along the grid axis.

    XY-only direction vector:
      The movement direction is computed from the curve's XY components only,
      normalised in 2D. This eliminates any Z component that would cause the
      new endpoint to drift off the datum plane.

    Z clamping:
      Each endpoint's Z is clamped independently to its original value from
      the post-promotion curve. This guarantees the replacement line lies
      exactly on the datum plane regardless of floating-point behaviour.

    Returns True if written, False if skipped.
    """
    key = (entry['grid_id'], entry['end_index'])
    if key in already_moved_keys:
        return False

    grid      = entry['grid']
    end_index = entry['end_index']

    # Promote to ViewSpecific and re-read the fresh promoted curve
    curve = promote_and_reread(grid, view)
    if curve is None:
        raise Exception("Could not read ViewSpecific curve after promotion")

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    # 2D length check — skip degenerate grids
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    length_2d = (dx * dx + dy * dy) ** 0.5
    if length_2d < MIN_GRID_LENGTH_FT:
        raise Exception("Degenerate grid curve (2D length {:.6f} ft)".format(
            length_2d))

    # Unit direction vector in XY only — no Z component
    if end_index == 1:
        ux = dx / length_2d   # End0 -> End1 direction
        uy = dy / length_2d
    else:
        ux = -dx / length_2d  # End1 -> End0 direction (outward from end 0)
        uy = -dy / length_2d

    # New bubble point: move in XY, clamp Z to original endpoint's Z
    bubble_pt = curve.GetEndPoint(end_index)
    fixed_pt  = curve.GetEndPoint(1 - end_index)

    new_bubble = XYZ(
        bubble_pt.X + ux * offset_distance,
        bubble_pt.Y + uy * offset_distance,
        bubble_pt.Z,   # Z clamped to this endpoint's original elevation
    )
    # Fixed end: preserve its own Z independently (handles sloped grids)
    fixed_clamped = XYZ(fixed_pt.X, fixed_pt.Y, fixed_pt.Z)

    if end_index == 0:
        new_curve = Line.CreateBound(new_bubble, fixed_clamped)
    else:
        new_curve = Line.CreateBound(fixed_clamped, new_bubble)

    # SetCurveInView — singular, correct IronPython Revit API method
    grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_curve)

    already_moved_keys.add(key)
    return True


# =============================================================================
# Main
# =============================================================================
def main():
    # ---- 1. Verify active view is a plan view -------------------------------
    active_view = uidoc.ActiveView
    if active_view.ViewType not in PLAN_VIEW_TYPES:
        forms.alert(
            "Please open a floor plan, ceiling plan, area plan, or "
            "engineering plan view before running this tool.",
            title="Wrong View Type",
        )
        script.exit()

    # ---- 2. User picks a grid to calibrate bubble size ---------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_inches = read_bubble_diameter_inches(ref_grid)
    output.print_md("Bubble diameter in use: **{:.4f} inches**".format(
        bubble_inches))

    # ---- 3. Collect plan views on sheets ------------------------------------
    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert(
            "No floor plan, ceiling plan, area plan, or engineering plan "
            "views placed on sheets were found in this project.",
            title="Nothing to do",
        )
        script.exit()

    output.print_md("Plan views on sheets found: **{}**".format(len(views)))

    # ---- 4. Stats ----------------------------------------------------------
    views_processed  = 0
    collisions_found = 0
    collisions_fixed = 0
    skipped_linked   = 0
    per_view_errors  = []

    # ---- 5. Single transaction — one Ctrl+Z undoes everything --------------
    t = Transaction(doc, "Separate Grid Bubbles")
    try:
        t.Start()

        for view in views:
            try:
                threshold   = bubble_diameter_model_units(view, bubble_inches)
                offset_dist = offset_distance_model_units(view, bubble_inches)

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                moved_keys = set()
                for a, b in pairs:
                    target = choose_entry_to_move(a, b)
                    if target is None:
                        skipped_linked += 1
                        continue
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

    # ---- 6. Results --------------------------------------------------------
    summary = "\n".join([
        "Views processed:    {}".format(views_processed),
        "Collisions found:   {}".format(collisions_found),
        "Collisions fixed:   {}".format(collisions_fixed),
        "Skipped (linked):   {}".format(skipped_linked),
        "Errors:             {}".format(len(per_view_errors)),
    ])

    output.print_md("### Results\n```\n{}\n```".format(summary))

    if per_view_errors:
        output.print_md("### Errors")
        for vname, err in per_view_errors:
            output.print_md("- **{}**: {}".format(vname, err))
        forms.alert(
            summary + "\n\nSee the pyRevit output window for error details.",
            title="Separate Grid Bubbles — Complete",
        )
    else:
        forms.alert(summary, title="Separate Grid Bubbles — Complete")


if __name__ == "__main__":
    main()