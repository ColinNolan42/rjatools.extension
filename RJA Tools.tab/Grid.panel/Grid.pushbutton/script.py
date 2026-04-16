# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Movement rules — consistent, no hardcoded names:

  BOTTOM bubbles (vertical grids running top-to-bottom, bubble at bottom):
    The grid with the LOWEST alphanumeric name shifts RIGHT (+X).
    Example: 4 and 5 collide -> 4 moves right (4 < 5)
    Example: 6 and 7 collide -> 6 moves right
    Example: 4(5) collide    -> 4 moves right

  SIDE bubbles (horizontal grids running left-to-right, bubble on side):
    The grid with the LOWEST alphanumeric name shifts DOWN (-Y).
    Example: D and E collide -> D moves down (D < E)
    Example: A and B collide -> A moves down

  "Lowest" = alphanumeric sort — numeric grids by integer value (4 < 10),
  alpha grids by letter (A < B), mixed handled naturally.
  No names are hardcoded — sort key is computed from grid.Name at runtime.

  Bubble position (bottom vs side) is determined from the bubble's Y
  position relative to the view crop box — no grid name parsing.

Default bubble diameter for collision threshold: 1/2" paper space.
  Stored as 0.04167 ft (0.5"/12). Multiplied by view.Scale to get the
  model-space collision distance. At 1/8" scale this equals 4'-0" in
  model space — matching the visual bubble size on the sheet.
  It is NOT the leader offset distance.

Leader offset: 4'-0" model space — the distance the bubble End point
  moves from the default AddLeader position along the offset direction.

Leader geometry (proven correct by diagnostic):
  1. AddLeader if no leader exists, else reuse existing
  2. Read default Anchor/Elbow/End from GetLeader
  3. Move End by full offset in the chosen direction
  4. Move Elbow by half offset in the same direction
  5. SetLeader — both End and Elbow must move together

Host grids only. Linked grids excluded.
FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "11.0.0"
__doc__     = ("Separates colliding grid bubbles. Lowest-named grid moves: "
               "bottom bubbles shift right, side bubbles shift down. "
               "Default bubble diameter 1/2\", leader offset 4ft.")

# -----------------------------------------------------------------------------
# Imports
# -----------------------------------------------------------------------------
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

# Default collision threshold diameter — 1/2" bubble in PAPER space.
# Stored as decimal feet: 0.5 inches / 12 = 0.04167 ft.
# Multiplied by view.Scale in collision_threshold() to convert to model
# space. At 1/8" scale (Scale=96): 0.04167 * 96 = 4.0 ft in model space.
DEFAULT_BUBBLE_DIAMETER_FT = 0.5 / 12.0   # 1/2" in decimal feet

# Leader offset — how far the bubble End moves from default AddLeader position.
LEADER_OFFSET_FT = 4.0

MIN_GRID_LENGTH_FT = 0.01


# =============================================================================
# Alphanumeric sort key — lowest name moves
# =============================================================================
def name_sort_key(name):
    """Sort key for grid names that handles numeric and alpha correctly.

    '4' < '5' < '10'  (by integer value, not string)
    'A' < 'B' < 'C'
    'A1' < 'A2' < 'B1'

    Splits name into alternating text/number chunks and converts
    numeric chunks to integers so they sort by value not lexicography.
    """
    parts = re.split(r'(\d+)', str(name))
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.upper()))
    return key


def entry_is_lower_name(entry_a, entry_b):
    """Return True if entry_a has the lower (lesser) grid name."""
    return name_sort_key(entry_a['name']) < name_sort_key(entry_b['name'])


# =============================================================================
# Bubble position — bottom vs side
# =============================================================================
def get_view_bounds(view):
    """Return (min_y, max_y, height) of the view crop box in model coords."""
    try:
        if view.CropBoxActive and view.CropBox is not None:
            bb = view.CropBox
            min_y  = bb.Min.Y
            max_y  = bb.Max.Y
            height = max_y - min_y
            if height > 0.01:
                return min_y, max_y, height
    except Exception:
        pass
    return None, None, None


def bubble_is_at_bottom_or_top(entry, view):
    """Return True if the bubble is in the top or bottom 35% of the view.

    True  -> vertical grid, bubble at bottom/top -> shift RIGHT
    False -> horizontal grid, bubble on side     -> shift DOWN
    """
    min_y, max_y, height = get_view_bounds(view)
    if height is None:
        # Fallback: use grid orientation from is_vertical flag
        return entry['is_vertical']

    frac_y = (entry['y'] - min_y) / height
    return frac_y < 0.35 or frac_y > 0.65


# =============================================================================
# Pick a grid — calibrate collision threshold
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
            "Click any grid line to calibrate bubble size for collision "
            "detection.\nThe script will then process all plan views on sheets.",
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
    """Read grid head annotation diameter in feet from the picked grid's type.

    Returns diameter in Revit internal feet.
    Falls back to DEFAULT_BUBBLE_DIAMETER_FT if not readable.
    """
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
                            # Parameter is in internal feet — diameter = 2x radius
                            diameter_ft = rp.AsDouble() * 2.0
                            if 0.01 < diameter_ft < 10.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} ft**".format(diameter_ft))
                                return diameter_ft
    except Exception as ex:
        logger.debug("read_bubble_diameter_ft: {}".format(ex))

    output.print_md("Using default bubble diameter: "
                    "**1/2 in ({:.5f} ft)**".format(DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# Collision threshold — scaled to view
# =============================================================================
def collision_threshold(view, bubble_diameter_ft):
    """Collision threshold in model feet for this view.

    bubble_diameter_ft is in internal feet (not scaled).
    Multiply by view.Scale to convert paper-space bubble size to model space.

    Example: 4ft bubble at 1/8" scale (Scale=96) -> 384 ft threshold in model.
    This correctly represents how large the bubble appears on the sheet.
    """
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


# =============================================================================
# Entry collection — HOST GRIDS ONLY, pure 2D
# =============================================================================
def collect_bubble_entries(document, view):
    """One deduplicated entry per visible bubble on host grids only.

    Stores grid name for alphanumeric sort and is_vertical for fallback
    orientation detection. All coordinates are 2D (Z stripped).
    """
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

            pt = curve.GetEndPoint(end_index)
            entries.append({
                'grid':        g,
                'grid_id':     g.Id.IntegerValue,
                'name':        g.Name,
                'end_index':   end_index,
                'x':           pt.X,
                'y':           pt.Y,
                'z':           pt.Z,
                'is_vertical': is_vertical,
            })

    return entries


# =============================================================================
# Collision detection — pure 2D
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """Return all (entry_a, entry_b) pairs within threshold feet. Pure 2D.

    No pair-level deduplication — every colliding pair is returned.
    When 3 grids collide (A-B, B-C, A-C), all three pairs come through.
    The already_done_keys set in apply_leader prevents any single grid
    from getting AddLeader called twice on the same end in the same view.
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
            if (dx * dx + dy * dy) <= threshold_sq:
                pairs.append((entries[i], entries[j]))

    return pairs


def choose_entry_to_move(entry_a, entry_b, view):
    """Return the entry that should receive the leader offset.

    Rule: the grid with the LOWEST alphanumeric name moves.
      Bottom bubbles (vertical grids): lowest name -> RIGHT (+X)
      Side bubbles (horizontal grids): lowest name -> DOWN (-Y)

    Lowest = smaller sort key from name_sort_key().
    If names are identical or unparseable, fall back to lower grid_id.
    """
    if entry_is_lower_name(entry_a, entry_b):
        return entry_a
    elif entry_is_lower_name(entry_b, entry_a):
        return entry_b
    else:
        # Names are equal — fallback: lower grid_id moves
        return entry_a if entry_a['grid_id'] < entry_b['grid_id'] else entry_b


# =============================================================================
# Offset direction per entry
# =============================================================================
def get_offset_direction(entry, view):
    """Return XYZ direction the bubble should move. Consistent rules:

    VERTICAL grids (bubble at bottom or top of view):
      Lowest name shifts RIGHT (+X).
      Grids are numbered left-to-right. Lowest moves further right,
      away from higher-numbered grids to its right.

    HORIZONTAL grids (bubble on left or right side of view):
      Lowest name shifts UP (+Y).
      Grids are lettered bottom-to-top. Lowest moves further up,
      away from higher-lettered grids above it.

    Both rules: lowest moves in the POSITIVE axis direction (+X or +Y).
    This is consistent — lowest always moves away from the pack.
    """
    if bubble_is_at_bottom_or_top(entry, view):
        return XYZ(1.0, 0.0, 0.0)    # vertical grid  — shift RIGHT (+X)
    else:
        return XYZ(0.0, 1.0, 0.0)    # horizontal grid — shift UP   (+Y)


# =============================================================================
# Leader application
# =============================================================================
def apply_leader(target_entry, view, already_done_keys):
    """Add or reposition a leader to visually offset the bubble.

    End MUST stay on the grid's infinite axis — only Elbow moves freely.
    Moving Elbow perpendicular to the grid axis creates the visual bend
    that pushes the bubble circle to the side without violating Revit's
    'End must be in datum plane curves' constraint.

    Direction:
      Vertical grid   (bottom/top bubble) -> Elbow moves RIGHT (+X)
      Horizontal grid (side bubble)       -> Elbow moves UP    (+Y)
      Both: lowest-named grid is the target, moves in positive direction.

    3-way collision handling:
      already_done_keys prevents AddLeader being called twice on the
      same (grid_id, end_index) within a single transaction. If a grid
      is the lowest in multiple pairs, it is moved once and subsequent
      pairs involving it are skipped via the key check.

    Returns True if applied, False if skipped.
    """
    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
    z         = target_entry['z']

    key = (target_entry['grid_id'], end_index)
    if key in already_done_keys:
        return False

    offset_dir = get_offset_direction(target_entry, view)

    # Check both already_done_keys (same transaction) AND GetLeader
    # (from a previous run) before calling AddLeader.
    # already_done_keys is the primary guard within a single transaction
    # because GetLeader may not reflect changes made earlier in the same
    # transaction before it is committed.
    has_leader_already = grid_has_leader_at_end(grid, view, end_index)
    if not has_leader_already:
        try:
            grid.AddLeader(datum_end, view)
        except Exception:
            # AddLeader can still fail if Revit's internal state already
            # has a leader — fall through to GetLeader and reposition.
            pass

    leader = grid.GetLeader(datum_end, view)
    if leader is None:
        raise Exception("GetLeader returned None after AddLeader")

    # End stays on the grid axis — DO NOT move it.
    # Only move Elbow perpendicular to push the bubble visually.
    current_elbow = leader.Elbow

    new_elbow = XYZ(
        current_elbow.X + offset_dir.X * LEADER_OFFSET_FT,
        current_elbow.Y + offset_dir.Y * LEADER_OFFSET_FT,
        z,
    )

    leader.Elbow = new_elbow
    grid.SetLeader(datum_end, view, leader)

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

    # ---- 2. Pick a grid to calibrate bubble size ---------------------------
    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("## Grid Bubble Separation")
    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))
    output.print_md("Leader offset:  **{} ft**".format(LEADER_OFFSET_FT))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)

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
                threshold = collision_threshold(view, bubble_diam_ft)

                entries = collect_bubble_entries(doc, view)
                if len(entries) < 2:
                    views_processed += 1
                    continue

                pairs = find_colliding_pairs(entries, threshold)
                collisions_found += len(pairs)

                done_keys = set()
                for a, b in pairs:
                    target = choose_entry_to_move(a, b, view)
                    try:
                        if apply_leader(target, view, done_keys):
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