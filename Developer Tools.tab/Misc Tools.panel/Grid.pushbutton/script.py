# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on all plan views placed on sheets.

VERSION 21.3.0 — fix ldr.End at opposite datum end (Revit default placement).

Crash root cause (journals 0256-0259):
  Background rendering threads (FullUpdateGraphicCacheUpdater,
  ElementsGraphicCacheUpdater) run continuously and update graphic caches for
  all views showing grid elements — including linked model views (157+ elements)
  and the 3D project view (912 elements). Any long transaction holding grid
  elements in a modified state collides with these background threads and crashes
  Revit. This is true even with zero explicit doc.Regenerate() calls (v20.3.1).

Fix — read/write phase separation:
  1. ALL collision detection and nudge/compaction math runs in-memory,
     OUTSIDE any transaction. In-memory state: mem_anchor, mem_elbow,
     mem_end dicts. Zero GetLeader/SetLeader calls during computation.
  2. Transaction 1 (short): AddLeader only for grids that need one,
     across ALL views. Flat loop, no computation.
  3. Re-read fresh state for all views after Transaction 1 commits.
  4. Compute all separation math for all views in-memory.
  5. Transaction 2 (short): SetLeader only — flat loop across all views,
     applying pre-computed final elbow positions. No loops, no collision
     detection inside the transaction.
  Transactions are open for milliseconds, not seconds.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "21.3.0"
__doc__     = ("Separates colliding grid bubbles on all plan views placed "
               "on sheets.")

import re
import traceback

from Autodesk.Revit.DB import (
    ElementId,
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
MAX_ITERATIONS             = 50
MAX_PASSES                 = 20
DEGENERATE_THRESHOLD       = 1e-4
COMPACT_PULL_STEP          = 0.05
MAX_COMPACT_ROUNDS         = 400


# =============================================================================
# Name sort helpers
# =============================================================================
def name_sort_key(name):
    key = []
    for seg in str(name).split("."):
        m = re.match(r'^([A-Za-z]*)(\d*)$', seg.strip())
        if m:
            alpha  = m.group(1).upper()
            digits = int(m.group(2)) if m.group(2) else 0
        else:
            alpha  = seg.upper()
            digits = 0
        key.append((alpha, digits))
    return key


def higher_name(name_a, name_b):
    return name_sort_key(name_a) > name_sort_key(name_b)


# =============================================================================
# Bubble diameter
# =============================================================================
def pick_reference_grid():
    try:
        from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

        class GridFilter(ISelectionFilter):
            def AllowElement(self, e):
                return isinstance(e, Grid)
            def AllowReference(self, r, p):
                return False

        forms.alert(
            "Click any grid line to calibrate bubble size.\n"
            "The script will then process all plan views on sheets.",
            title="Separate Grid Bubbles",
            ok=True,
        )
        ref  = uidoc.Selection.PickObject(
            ObjectType.Element, GridFilter(), "Click any grid line")
        elem = doc.GetElement(ref.ElementId)
        if isinstance(elem, Grid):
            return elem
        forms.alert("Selected element is not a grid. Cancelled.",
                    title="Invalid Selection")
        return None
    except Exception:
        return None


def read_bubble_diameter_ft(grid):
    try:
        gt = doc.GetElement(grid.GetTypeId())
        if gt is not None:
            for pname in ("End 1 Default Grid Head",
                          "End 2 Default Grid Head",
                          "Default Grid Head"):
                p = gt.LookupParameter(pname)
                if p and p.HasValue:
                    sym = doc.GetElement(p.AsElementId())
                    if sym is None:
                        continue
                    for rname in ("Circle Radius", "Head Radius",
                                  "Radius", "Bubble Radius"):
                        rp = sym.LookupParameter(rname)
                        if rp and rp.HasValue:
                            d = rp.AsDouble() * 2.0
                            if 0.01 < d < 10.0:
                                output.print_md(
                                    "Bubble diameter from family: "
                                    "**{:.4f} ft**".format(d))
                                return d
    except Exception as ex:
        logger.debug("read_bubble_diameter_ft: {}".format(ex))

    output.print_md("Could not read bubble diameter from annotation family.")
    try:
        raw = forms.ask_for_string(
            default="2.0",
            prompt=("Enter the grid bubble diameter in MODEL SPACE FEET.\n"
                    "Common value: 2.0 ft (1/4\" at 1/8\" scale)."),
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
    ids = set()
    for sheet in (FilteredElementCollector(document)
                  .OfClass(ViewSheet).ToElements()):
        try:
            for vid in sheet.GetAllPlacedViews():
                ids.add(vid.IntegerValue)
        except Exception:
            pass
    return ids


def collect_plan_views_on_sheets(document):
    sheet_ids = get_sheet_view_ids(document)
    result = []
    for v in FilteredElementCollector(document).OfClass(View):
        if v.IsTemplate:
            continue
        if v.ViewType not in PLAN_VIEW_TYPES:
            continue
        if v.Id.IntegerValue not in sheet_ids:
            continue
        result.append(v)
    return result


# =============================================================================
# Grid helpers
# =============================================================================
def get_grid_curve_in_view(grid, view):
    for et in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(et, view)
            if curves:
                return curves[0]
        except Exception:
            continue
    return None


def grid_has_bubble_at_end(grid, view, end_index):
    try:
        de = DatumEnds.End0 if end_index == 0 else DatumEnds.End1
        return grid.IsBubbleVisibleInView(de, view)
    except Exception:
        return True


def get_grids_for_view(document, view):
    """Return grids visible in view. Falls back to all-document collector
    if scoped collector returns zero (known Revit limitation for new views)."""
    try:
        grids = list(FilteredElementCollector(document, view.Id)
                     .OfClass(Grid).ToElements())
        if grids:
            return grids
    except Exception:
        pass
    return list(FilteredElementCollector(document)
                .OfClass(Grid).ToElements())


# =============================================================================
# Nudge direction
# =============================================================================
def get_nudge_direction(grid, view):
    try:
        curve = get_grid_curve_in_view(grid, view)
        if curve is None:
            return XYZ(1.0, 0.0, 0.0)
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx, dy = p1.X - p0.X, p1.Y - p0.Y
        L = (dx*dx + dy*dy) ** 0.5
        if L < 1e-9:
            return XYZ(1.0, 0.0, 0.0)
        tx, ty = dx/L, dy/L
        if abs(ty) > abs(tx):
            if ty < 0.0: tx, ty = -tx, -ty
        else:
            if tx < 0.0: tx, ty = -tx, -ty
        return XYZ(ty, -tx, 0.0)
    except Exception:
        return XYZ(1.0, 0.0, 0.0)


# =============================================================================
# In-memory state — read once from Revit, then track in dicts
# =============================================================================
def init_mem_state(grids, view):
    """Read anchor, elbow, end for all visible grid leaders.
    Call outside any transaction so values are guaranteed fresh.
    Returns (mem_anchor, mem_elbow, mem_end) dicts keyed by (gid_int, ei).

    IMPORTANT — ldr.End and ldr.Elbow are NOT used directly.
    Revit's default AddLeader geometry places ldr.End at the OPPOSITE datum
    end of the grid (e.g. End0 when the bubble is at End1).  Storing that
    far-end point in mem_end causes place_elbow to project nudges onto a
    200-ft grid-spanning segment instead of the short local leader stub,
    making the crossing guard reject every nudge and producing the long
    diagonal arm that visually crosses adjacent grid lines.

    Fix: always seed mem_elbow and mem_end from curve.GetEndPoint(ei) — the
    grid's own endpoint at the bubble side.  mem_anchor uses ldr.Anchor only
    when it is closer to that endpoint than to the opposite one (i.e. it has
    been genuinely displaced by a prior tool run); otherwise it is also reset
    to the grid endpoint so the algorithm starts from a clean local state."""
    mem_anchor = {}
    mem_elbow  = {}
    mem_end    = {}
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        for ei in (0, 1):
            de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, ei):
                continue
            key      = (g.Id.IntegerValue, ei)
            ep       = curve.GetEndPoint(ei)     if curve else None
            ep_other = curve.GetEndPoint(1 - ei) if curve else None
            try:
                ldr = g.GetLeader(de, view)
                if ldr is not None:
                    anc = ldr.Anchor
                    # Accept ldr.Anchor only if it is on the correct side
                    # of the grid (closer to ep than to ep_other).  An anchor
                    # at the far end is Revit's default placement artifact.
                    if (anc is not None and
                            ep is not None and ep_other is not None):
                        da = ((anc.X - ep.X)**2 +
                              (anc.Y - ep.Y)**2) ** 0.5
                        db = ((anc.X - ep_other.X)**2 +
                              (anc.Y - ep_other.Y)**2) ** 0.5
                        if db < da:
                            anc = ep  # anchor at wrong end — reset
                    if anc is None:
                        anc = ep
                    mem_anchor[key] = anc if anc else (ep or XYZ.Zero)
                    # Always seed from the grid's own endpoint so all
                    # subsequent calculations stay in the short local arm.
                    mem_elbow[key] = ep if ep is not None else ldr.Elbow
                    mem_end[key]   = ep if ep is not None else ldr.End
                elif ep is not None:
                    mem_anchor[key] = ep
                    mem_elbow[key]  = ep
                    mem_end[key]    = ep
            except Exception:
                if ep is not None:
                    mem_anchor[key] = ep
                    mem_elbow[key]  = ep
                    mem_end[key]    = ep
    return mem_anchor, mem_elbow, mem_end


# =============================================================================
# Collision detection
# =============================================================================
def build_bubble_list(grids, view, mem_anchor):
    """(grid, datum_end, end_index, bubble_pos) for every visible bubble."""
    result = []
    for g in grids:
        for ei in (0, 1):
            de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, ei):
                continue
            pt = mem_anchor.get((g.Id.IntegerValue, ei))
            if pt is None:
                continue
            result.append((g, de, ei, pt))
    return result


def find_colliding_pairs(positions, threshold):
    pairs = []
    thr2  = threshold * threshold
    n     = len(positions)
    for i in range(n):
        for j in range(i+1, n):
            if positions[i][0].Id == positions[j][0].Id:
                continue
            p1, p2 = positions[i][3], positions[j][3]
            dx, dy = p1.X - p2.X, p1.Y - p2.Y
            if dx*dx + dy*dy <= thr2:
                pairs.append((positions[i], positions[j]))
    return pairs


# =============================================================================
# Elbow placement — clamp along-segment component only
# =============================================================================
def place_elbow(proposed_x, proposed_y, anchor, end):
    ax, ay = anchor.X, anchor.Y
    ex, ey = end.X,    end.Y
    sx, sy = ex - ax,  ey - ay
    ss     = sx*sx + sy*sy

    if ss < 1e-12:
        return XYZ(proposed_x, proposed_y, anchor.Z)

    t_raw = ((proposed_x - ax) * sx +
             (proposed_y - ay) * sy) / ss

    t_clamped = max(0.05, min(0.95, t_raw))

    seg_x = ax + t_clamped * sx
    seg_y = ay + t_clamped * sy

    proj_x = ax + t_raw * sx
    proj_y = ay + t_raw * sy
    perp_x = proposed_x - proj_x
    perp_y = proposed_y - proj_y

    return XYZ(seg_x + perp_x, seg_y + perp_y, anchor.Z)


# =============================================================================
# Phase A: detect which grids need AddLeader (pure read, no Revit writes)
# =============================================================================
def detect_grids_needing_leaders(grids, view, mem_anchor, threshold):
    """Return set of (gid_int, ei) that collide AND have no leader yet."""
    positions = build_bubble_list(grids, view, mem_anchor)
    pairs     = find_colliding_pairs(positions, threshold)
    if not pairs:
        return set()

    name_map = {g.Id.IntegerValue: g.Name for g in grids}
    grid_map = {g.Id.IntegerValue: g       for g in grids}
    movers   = set()

    for pos_a, pos_b in pairs:
        g_a, _, idx_a, _ = pos_a
        g_b, _, idx_b, _ = pos_b
        na = name_map.get(g_a.Id.IntegerValue, "")
        nb = name_map.get(g_b.Id.IntegerValue, "")
        if higher_name(na, nb):
            movers.add((g_a.Id.IntegerValue, idx_a))
        else:
            movers.add((g_b.Id.IntegerValue, idx_b))

    needs_leader = set()
    for gid, ei in movers:
        g  = grid_map.get(gid)
        de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
        if g and g.GetLeader(de, view) is None:
            needs_leader.add((gid, ei))

    return needs_leader


# =============================================================================
# Grid-line crossing guard
# =============================================================================
def build_grid_dirs(grids, view):
    """Return (dirs, ref_pts):
    - dirs: gid_int → (tx, ty) unit vector along grid line
    - ref_pts: gid_int → XYZ reference point on that grid line (End0)
    Used for own-grid-line crossing guard and adjacent-grid leader path checks."""
    dirs    = {}
    ref_pts = {}
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx, dy = p1.X - p0.X, p1.Y - p0.Y
        L = (dx*dx + dy*dy) ** 0.5
        if L > 1e-9:
            gid = g.Id.IntegerValue
            dirs[gid]    = (dx/L, dy/L)
            ref_pts[gid] = p0
    return dirs, ref_pts


def would_cross_grid_line(old_anchor, new_anchor, end_pt, grid_tx, grid_ty):
    """True if the anchor would cross the grid line from old to new position.
    The grid line passes through end_pt in direction (grid_tx, grid_ty).
    Side is determined by the sign of the cross product (pt - end) × (tx, ty)."""
    def side(pt):
        return (pt.X - end_pt.X) * grid_ty - (pt.Y - end_pt.Y) * grid_tx
    old_s = side(old_anchor)
    new_s = side(new_anchor)
    return old_s * new_s < 0.0


def _seg_crosses_line(ax, ay, bx, by, rx, ry, tx, ty):
    """True if segment A→B strictly crosses the infinite line through (rx,ry)
    with direction (tx,ty).  Touch (one endpoint on line) is NOT a crossing."""
    sa = (ax - rx) * ty - (ay - ry) * tx
    sb = (bx - rx) * ty - (by - ry) * tx
    return sa * sb < -1e-12


def leader_path_crosses_grid(anchor, elbow, end, own_gid, grid_dirs, ref_pts):
    """True if the two-segment leader path end→elbow→anchor crosses the
    infinite line of any grid OTHER than own_gid."""
    ex, ey = end.X,   end.Y
    lx, ly = elbow.X, elbow.Y
    ax, ay = anchor.X, anchor.Y
    for gid, (tx, ty) in grid_dirs.items():
        if gid == own_gid:
            continue
        rp = ref_pts.get(gid)
        if rp is None:
            continue
        rx, ry = rp.X, rp.Y
        if (_seg_crosses_line(ex, ey, lx, ly, rx, ry, tx, ty) or
                _seg_crosses_line(lx, ly, ax, ay, rx, ry, tx, ty)):
            return True
    return False


# =============================================================================
# Phase B: separation math — pure in-memory, zero Revit API calls
# =============================================================================
def compute_separation_in_memory(grids, view, mem_anchor, mem_elbow, mem_end,
                                 threshold):
    """Run separation passes entirely in-memory.
    Modifies mem_anchor and mem_elbow in place."""
    nudge_step      = threshold / 8.0
    name_map        = {g.Id.IntegerValue: g.Name for g in grids}
    grid_dirs, grid_ref_pts = build_grid_dirs(grids, view)

    for pass_num in range(MAX_PASSES):
        positions = build_bubble_list(grids, view, mem_anchor)
        pairs     = find_colliding_pairs(positions, threshold)
        if not pairs:
            output.print_md(
                "  clean after {} pass(es)".format(pass_num + 1))
            return

        output.print_md("  **Pass {} — {} collision(s)**".format(
            pass_num + 1, len(pairs)))

        for iteration in range(MAX_ITERATIONS):
            positions = build_bubble_list(grids, view, mem_anchor)
            pairs     = find_colliding_pairs(positions, threshold)
            if not pairs:
                output.print_md(
                    "  Pass {} resolved in {} iteration(s)".format(
                        pass_num + 1, iteration))
                break

            # Build per-mover nudge vectors from ACTUAL separation directions,
            # not the canonical grid-perpendicular. This avoids pushing bubbles
            # through their own grid lines when the canonical direction is wrong.
            targets = {}
            for pos_a, pos_b in pairs:
                g_a, end_a, idx_a, pt_a = pos_a
                g_b, end_b, idx_b, pt_b = pos_b
                na = name_map.get(g_a.Id.IntegerValue, "")
                nb = name_map.get(g_b.Id.IntegerValue, "")
                if higher_name(na, nb):
                    mg, me, mi, mn = g_a, end_a, idx_a, na
                    pt_mover, pt_other = pt_a, pt_b
                else:
                    mg, me, mi, mn = g_b, end_b, idx_b, nb
                    pt_mover, pt_other = pt_b, pt_a

                # Direction: from the other bubble toward the mover
                sep_x = pt_mover.X - pt_other.X
                sep_y = pt_mover.Y - pt_other.Y
                sep_l = (sep_x*sep_x + sep_y*sep_y) ** 0.5
                if sep_l > 1e-9:
                    nd_x, nd_y = sep_x/sep_l, sep_y/sep_l
                else:
                    # Bubbles on same point — fall back to canonical perpendicular
                    nd = get_nudge_direction(mg, view)
                    nd_x, nd_y = nd.X, nd.Y

                key = (mg.Id.IntegerValue, mi)
                if key not in targets:
                    targets[key] = [mg, me, 0.0, 0.0, mn]
                targets[key][2] += nd_x
                targets[key][3] += nd_y

            sorted_targets = sorted(
                targets.items(),
                key=lambda item: name_sort_key(item[1][4]),
                reverse=True,
            )

            any_moved = False
            for key, td in sorted_targets:
                gid, ei      = key
                net_x, net_y = td[2], td[3]
                net_len      = (net_x*net_x + net_y*net_y) ** 0.5
                if net_len < 1e-9:
                    continue
                nx = (net_x / net_len) * nudge_step
                ny = (net_y / net_len) * nudge_step

                anchor = mem_anchor.get(key)
                elbow  = mem_elbow.get(key)
                end    = mem_end.get(key)
                if anchor is None or elbow is None or end is None:
                    continue

                ae_dist = ((anchor.X - elbow.X)**2 +
                           (anchor.Y - elbow.Y)**2) ** 0.5
                if ae_dist < DEGENERATE_THRESHOLD:
                    nd_dir = get_nudge_direction(td[0], view)
                    elbow  = XYZ(anchor.X + nd_dir.X * nudge_step * 4,
                                 anchor.Y + nd_dir.Y * nudge_step * 4,
                                 anchor.Z)
                    mem_elbow[key] = elbow

                new_elbow = place_elbow(
                    elbow.X + nx, elbow.Y + ny, anchor, end)

                if (abs(new_elbow.X - elbow.X) < 1e-6 and
                        abs(new_elbow.Y - elbow.Y) < 1e-6):
                    continue

                delta_x = new_elbow.X - elbow.X
                delta_y = new_elbow.Y - elbow.Y
                new_anchor = XYZ(
                    anchor.X + delta_x, anchor.Y + delta_y, anchor.Z)

                # Guard: reject if this nudge would push the bubble through
                # its own grid line.
                gtx, gty = grid_dirs.get(gid, (1.0, 0.0))
                if would_cross_grid_line(anchor, new_anchor, end, gtx, gty):
                    continue

                # Guard: reject if the new leader path (end→elbow→anchor)
                # would cross any OTHER grid's line.
                if leader_path_crosses_grid(
                        new_anchor, new_elbow, end,
                        gid, grid_dirs, grid_ref_pts):
                    continue

                mem_elbow[key]  = new_elbow
                mem_anchor[key] = new_anchor
                any_moved = True

            if not any_moved:
                break

    else:
        output.print_md(
            "  hit MAX_PASSES ({}) — some collisions may remain".format(
                MAX_PASSES))


# =============================================================================
# Compaction — pure in-memory
# =============================================================================
def compact_in_memory(grids, view, mem_anchor, mem_elbow, mem_end, threshold):
    """Pull leaders back toward grid line endpoints in-memory.
    Returns total steps moved."""
    thr2            = threshold * threshold
    total_moves     = 0
    grid_dirs, grid_ref_pts = build_grid_dirs(grids, view)

    grid_eps = {}
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        for ei in (0, 1):
            if grid_has_bubble_at_end(g, view, ei):
                grid_eps[(g.Id.IntegerValue, ei)] = curve.GetEndPoint(ei)

    for _ in range(MAX_COMPACT_ROUNDS):
        # Current bubble positions for this round
        all_pts = {}
        for g in grids:
            for ei in (0, 1):
                key = (g.Id.IntegerValue, ei)
                if not grid_has_bubble_at_end(g, view, ei):
                    continue
                pt = mem_anchor.get(key)
                if pt is not None:
                    all_pts[key] = pt

        candidates = []
        for key, pt in all_pts.items():
            grid_ep = grid_eps.get(key)
            if grid_ep is None:
                continue
            dx = grid_ep.X - pt.X
            dy = grid_ep.Y - pt.Y
            dist = (dx * dx + dy * dy) ** 0.5
            if dist >= COMPACT_PULL_STEP:
                candidates.append((dist, key, grid_ep))

        if not candidates:
            break

        candidates.sort(key=lambda c: -c[0])
        moved_any = False

        for _, key, grid_ep in candidates:
            cur_pt = mem_anchor.get(key)
            if cur_pt is None:
                continue

            dx = grid_ep.X - cur_pt.X
            dy = grid_ep.Y - cur_pt.Y
            dist = (dx * dx + dy * dy) ** 0.5
            if dist < COMPACT_PULL_STEP:
                continue

            ux, uy = dx / dist, dy / dist
            cand = XYZ(cur_pt.X + ux * COMPACT_PULL_STEP,
                       cur_pt.Y + uy * COMPACT_PULL_STEP,
                       cur_pt.Z)

            safe = True
            for key2, pt2 in all_pts.items():
                if key2 == key:
                    continue
                pt2_cur = mem_anchor.get(key2, pt2)
                cdx = cand.X - pt2_cur.X
                cdy = cand.Y - pt2_cur.Y
                if cdx * cdx + cdy * cdy <= thr2:
                    safe = False
                    break
            if not safe:
                continue

            anchor = mem_anchor.get(key)
            elbow  = mem_elbow.get(key)
            end    = mem_end.get(key)
            if anchor is None or elbow is None or end is None:
                continue

            new_elbow = place_elbow(
                elbow.X + ux * COMPACT_PULL_STEP,
                elbow.Y + uy * COMPACT_PULL_STEP,
                anchor, end,
            )
            if (abs(new_elbow.X - elbow.X) < 1e-6 and
                    abs(new_elbow.Y - elbow.Y) < 1e-6):
                continue

            delta_x = new_elbow.X - elbow.X
            delta_y = new_elbow.Y - elbow.Y
            new_anchor = XYZ(
                anchor.X + delta_x, anchor.Y + delta_y, anchor.Z)

            # Guard: compaction must not push the bubble through its own
            # grid line (would flip the leader to the wrong side).
            gid_int = key[0]
            gtx, gty = grid_dirs.get(gid_int, (1.0, 0.0))
            if would_cross_grid_line(anchor, new_anchor, end, gtx, gty):
                continue

            # Guard: compaction must not move the leader path across any
            # other grid's line.
            if leader_path_crosses_grid(
                    new_anchor, new_elbow, end,
                    gid_int, grid_dirs, grid_ref_pts):
                continue

            mem_elbow[key]  = new_elbow
            mem_anchor[key] = new_anchor
            total_moves += 1
            moved_any    = True

        if not moved_any:
            break

    return total_moves


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

    output.print_md("## Grid Bubble Separation — v21.3.0")

    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, ref_grid.Id.IntegerValue))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)
    threshold      = bubble_diam_ft

    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    # ── Read Phase 1: initial state for all views (outside any transaction) ──
    view_data = {}  # view_id_int → (view, grids, mem_anchor, mem_elbow, mem_end)
    for v in views:
        grids = get_grids_for_view(doc, v)
        if len(grids) < 2:
            continue
        ma, me, mn = init_mem_state(grids, v)
        view_data[v.Id.IntegerValue] = (v, grids, ma, me, mn)

    if not view_data:
        forms.alert("No views with 2+ grids found.", title="Nothing to do")
        script.exit()

    # ── Read Phase 2: detect missing leaders for all views ──────────────────
    # needs_by_view: view_id_int → set of (gid_int, ei)
    needs_by_view = {}
    total_needs   = 0
    for vid, (v, grids, ma, me, mn) in view_data.items():
        nl = detect_grids_needing_leaders(grids, v, ma, threshold)
        if nl:
            needs_by_view[vid] = nl
            total_needs += len(nl)

    # ── Write Transaction 1: AddLeader only (short flat loop) ───────────────
    if needs_by_view:
        output.print_md(
            "Adding {} leader(s) across {} view(s)...".format(
                total_needs, len(needs_by_view)))
        t1 = Transaction(doc, "Add Grid Leaders")
        try:
            t1.Start()
            for vid, nl in needs_by_view.items():
                v, grids, _, _, _ = view_data[vid]
                gmap = {g.Id.IntegerValue: g for g in grids}
                for gid, ei in nl:
                    g  = gmap.get(gid)
                    de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
                    if g:
                        try:
                            g.AddLeader(de, v)
                        except Exception as ex:
                            logger.debug(
                                "AddLeader gid={} ei={}: {}".format(
                                    gid, ei, ex))
            t1.Commit()
        except Exception as ex:
            if t1.HasStarted() and not t1.HasEnded():
                t1.RollBack()
            forms.alert("AddLeader failed:\n{}".format(ex), title="Error")
            script.exit()

        # Re-read fresh state for all views (Revit regenerated on commit)
        for vid in list(view_data.keys()):
            v, grids, _, _, _ = view_data[vid]
            ma, me, mn = init_mem_state(grids, v)
            view_data[vid] = (v, grids, ma, me, mn)

    # ── Compute Phase: all separation math in-memory (no transactions) ──────
    output.print_md("### Separation")

    # initial_elbows_by_view: view_id_int → dict of (gid, ei) → elbow XYZ
    initial_elbows_by_view = {}
    for vid, (v, grids, ma, me, mn) in view_data.items():
        initial_elbows_by_view[vid] = dict(me)

    for vid, (v, grids, ma, me, mn) in view_data.items():
        output.print_md("#### View: `{}`".format(v.Name))
        compute_separation_in_memory(grids, v, ma, me, mn, threshold)
        c = compact_in_memory(grids, v, ma, me, mn, threshold)
        if c:
            output.print_md("  Compaction: {} step(s)".format(c))

    # ── Collect changes across all views ────────────────────────────────────
    # changes_by_view: view_id_int → [(gid_int, ei)]
    changes_by_view = {}
    total_changes   = 0
    for vid, (v, grids, ma, me, mn) in view_data.items():
        init_el = initial_elbows_by_view[vid]
        changed = []
        for key, final_elbow in me.items():
            init_elbow = init_el.get(key)
            if init_elbow is None:
                changed.append(key)
                continue
            if (abs(final_elbow.X - init_elbow.X) > 1e-6 or
                    abs(final_elbow.Y - init_elbow.Y) > 1e-6):
                changed.append(key)
        if changed:
            changes_by_view[vid] = changed
            total_changes += len(changed)

    if not changes_by_view:
        output.print_md("No elbow positions changed.")
        forms.alert("No grid bubbles needed separation.",
                    title="Separate Grid Bubbles")
        script.exit()

    output.print_md(
        "Applying **{}** change(s) across **{}** view(s)...".format(
            total_changes, len(changes_by_view)))

    # ── Write Transaction 2: SetLeader only (short flat loop) ───────────────
    t2 = Transaction(doc, "Separate Grid Bubbles")
    applied = 0
    try:
        t2.Start()
        for vid, changed_keys in changes_by_view.items():
            v, grids, ma, me, mn = view_data[vid]
            gmap = {g.Id.IntegerValue: g for g in grids}
            for gid, ei in changed_keys:
                g  = gmap.get(gid)
                de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
                if g is None:
                    continue
                try:
                    ldr = g.GetLeader(de, v)
                    if ldr is None:
                        continue
                    new_elbow = me.get((gid, ei))
                    if new_elbow is None:
                        continue
                    ldr.Elbow = new_elbow
                    g.SetLeader(de, v, ldr)
                    applied += 1
                except Exception as ex:
                    logger.debug(
                        "SetLeader gid={} ei={}: {}".format(gid, ei, ex))
        t2.Commit()
        output.print_md("Applied {} change(s).".format(applied))
    except Exception as ex:
        if t2.HasStarted() and not t2.HasEnded():
            t2.RollBack()
        forms.alert(
            "SetLeader transaction failed and was rolled back.\n\n{}".format(
                ex),
            title="Error",
        )
        script.exit()

    summary = "\n".join([
        "Views processed: {}".format(len(view_data)),
        "Leaders added:   {}".format(total_needs),
        "Changes applied: {}".format(applied),
    ])
    output.print_md("### Results\n```\n{}\n```".format(summary))
    forms.alert(summary, title="Separate Grid Bubbles — Done")


if __name__ == "__main__":
    main()
