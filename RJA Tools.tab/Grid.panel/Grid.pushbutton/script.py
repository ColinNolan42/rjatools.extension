# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on plan views placed on sheets.

Key insight from diagnostic:
  Revit's AddLeader already places the bubble in a valid separated
  position automatically. The default geometry it creates IS the
  correct elbow/break separation. We do NOT need SetLeader at all.

  All previous SetLeader errors were caused by trying to reposition
  a leader that AddLeader had already placed correctly.

Strategy:
  1. Detect colliding bubble pairs (2D distance check)
  2. Choose which grid to move (lowest alphanumeric name)
  3. If that grid already has a leader on that end -> skip (already fixed)
  4. If not -> call AddLeader only. Done. Revit handles the geometry.

  AddLeader default geometry (confirmed by diagnostic):
    Vertical grids:   Anchor 2ft left of line, Elbow 2ft up, End on axis
    Horizontal grids: Anchor 2ft below line, Elbow 2ft left, End on axis
  This is exactly the correct separation visual.

Movement rules:
  Vertical grids (bubble at bottom/top): lowest name -> AddLeader on End1
  Horizontal grids (bubble on side):     lowest name -> AddLeader on End1
  "Lowest" = alphanumeric sort (4<5, A<B, D<E)

Host grids only. Linked grids excluded.
FloorPlan, CeilingPlan, AreaPlan, EngineeringPlan on sheets only.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "12.0.0"
__doc__     = ("Separates colliding grid bubbles using AddLeader only. "
               "Lowest-named grid gets the leader. No SetLeader calls.")

import re
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Grid,
    View,
    ViewSheet,
    ViewType,
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

# Bubble diameter in paper-space feet: 0.5" / 12 = 0.04167 ft
# Multiplied by view.Scale to get model-space collision threshold
DEFAULT_BUBBLE_DIAMETER_FT = 0.5 / 12.0

MIN_GRID_LENGTH_FT = 0.01


# =============================================================================
# Alphanumeric sort — lowest name moves
# =============================================================================
def name_sort_key(name):
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

    output.print_md("Using default bubble diameter: **1/2 in ({:.5f} ft)**".format(
        DEFAULT_BUBBLE_DIAMETER_FT))
    return DEFAULT_BUBBLE_DIAMETER_FT


# =============================================================================
# Collision threshold — scaled to view
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


# =============================================================================
# Entry collection — HOST GRIDS ONLY, pure 2D
# =============================================================================
def collect_bubble_entries(document, view):
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

        for end_index in (0, 1):
            if not grid_has_bubble_at_end(g, view, end_index):
                continue

            key = (g.Id.IntegerValue, end_index)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            pt = curve.GetEndPoint(end_index)
            entries.append({
                'grid':      g,
                'grid_id':   g.Id.IntegerValue,
                'name':      g.Name,
                'end_index': end_index,
                'x':         pt.X,
                'y':         pt.Y,
                'z':         pt.Z,
            })

    return entries


# =============================================================================
# Collision detection — pure 2D
# =============================================================================
def find_colliding_pairs(entries, threshold):
    """All colliding pairs. No dedup — already_done_keys handles multi-collisions."""
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


def choose_entry_to_move(entry_a, entry_b):
    """Lowest alphanumeric name moves. Fallback: lower grid_id."""
    if entry_is_lower_name(entry_a, entry_b):
        return entry_a
    elif entry_is_lower_name(entry_b, entry_a):
        return entry_b
    return entry_a if entry_a['grid_id'] < entry_b['grid_id'] else entry_b


# =============================================================================
# Leader application — AddLeader ONLY, no SetLeader
# =============================================================================
def apply_leader(target_entry, view, already_done_keys):
    """Add a leader to the target grid's bubble end if not already present.

    We call AddLeader only. Revit places the bubble in a valid separated
    position automatically — no SetLeader needed or called.

    The default AddLeader geometry (confirmed by diagnostic):
      Vertical grids:   Anchor offset left, Elbow up, End on axis above
      Horizontal grids: Anchor offset down, Elbow left, End on axis left
    This IS the correct elbow/break visual separation.

    Grids that already have a leader are skipped — they were either
    fixed in a previous run or earlier in this transaction.

    Returns True if leader was added, False if skipped.
    """
    grid      = target_entry['grid']
    end_index = target_entry['end_index']
    datum_end = DatumEnds.End0 if end_index == 0 else DatumEnds.End1

    key = (target_entry['grid_id'], end_index)
    if key in already_done_keys:
        return False

    # Skip if leader already exists — idempotent across repeated runs
    if grid_has_leader_at_end(grid, view, end_index):
        already_done_keys.add(key)
        return False

    # AddLeader only — Revit handles all geometry automatically
    grid.AddLeader(datum_end, view)
    already_done_keys.add(key)
    return True


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

    views_processed  = 0
    collisions_found = 0
    leaders_added    = 0
    already_had      = 0
    per_view_errors  = []

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
                    target = choose_entry_to_move(a, b)

                    # Track grids that already had leaders
                    end = DatumEnds.End0 if target['end_index'] == 0 else DatumEnds.End1
                    if grid_has_leader_at_end(target['grid'], view, target['end_index']):
                        already_had += 1
                        done_keys.add((target['grid_id'], target['end_index']))
                        continue

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

    summary = "\n".join([
        "Views processed:     {}".format(views_processed),
        "Collisions found:    {}".format(collisions_found),
        "Leaders added:       {}".format(leaders_added),
        "Already had leader:  {}".format(already_had),
        "Errors:              {}".format(len(per_view_errors)),
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