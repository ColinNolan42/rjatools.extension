# -*- coding: utf-8 -*-
"""Separates colliding grid bubbles on all plan views placed on sheets.

VERSION 22.0.0 — perp_offset redesign.

Root causes fixed vs v21.4.0:
  1. init_mem_state reset elbow to ep (degenerate, zero offset) — now reads
     actual ldr.Elbow so the initial offset reflects Revit's current state.
  2. get_nudge_direction canonicalized to one side — replaced with sep_dot
     (dot of mover-ep minus other-ep onto mover perp_vec) which gives the
     correct side regardless of grid orientation.
  3. leader_path_crosses_grid used mem_end = ep (wrong) — replaced with an
     anchor-position crossing guard in perp_offset space.
  4. place_elbow projected onto anchor-end segment — eliminated; elbow is now
     always ep + perp_offset * perp_vec (clean single degree of freedom).

Crash-safe structure retained from v21.4.0:
  T1 (short): AddLeader only — flat loop, no computation.
  T2 (short): SetLeader(Elbow) only — flat loop, apply pre-computed offsets.
  All math runs outside transactions in-memory.
"""

__title__   = "Separate\nGrid Bubbles"
__author__  = "MEP Tools"
__version__ = "22.1.0"
__doc__     = ("Separates colliding grid bubbles on all plan views placed "
               "on sheets.")

import re

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

def _eid_int(element_id):
    """Version-safe ElementId -> int/long.

    Revit 2024+ replaced ElementId.IntegerValue (int) with ElementId.Value
    (long); Revit 2025/2026 removed IntegerValue entirely. Revit 2022/2023
    only have IntegerValue.
    """
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


DEFAULT_BUBBLE_DIAMETER_FT = 2.0
MAX_ITERATIONS             = 50
MAX_PASSES                 = 20
COMPACT_PULL_STEP          = 0.05
MAX_COMPACT_ROUNDS         = 400
# How far inside the grid (along-grid direction from the bubble ep) to place
# the leader elbow.  Must be > 0 so Elbow.Y != Anchor.Y — otherwise Revit's
# SetLeader silently rejects the call and leaves the default AddLeader position.
LEADER_SHOULDER_FT         = 1.0


# =============================================================================
# Name sort helpers (unchanged)
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
# Bubble diameter (unchanged)
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
# View collection (unchanged)
# =============================================================================
def get_sheet_view_ids(document):
    ids = set()
    for sheet in (FilteredElementCollector(document)
                  .OfClass(ViewSheet).ToElements()):
        try:
            for vid in sheet.GetAllPlacedViews():
                ids.add(_eid_int(vid))
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
        if _eid_int(v.Id) not in sheet_ids:
            continue
        result.append(v)
    return result


# =============================================================================
# Grid helpers (unchanged)
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
# perp_vec — NEW
# =============================================================================
def compute_perp_vec(grid, view):
    """Return perpendicular unit vector (XYZ) to this grid in-plane (Z=0).
    Defined as (-ty, tx, 0) where (tx, ty) is the along-grid unit vector."""
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
    return XYZ(-ty, tx, 0.0)


# =============================================================================
# In-memory state — NEW: perp_offset space
#
# mem_offset   : (gid_int, ei) → float   signed offset from ep along perp_vec
# mem_perp_vec : gid_int       → XYZ     perpendicular unit vector for grid
# mem_ep       : (gid_int, ei) → XYZ     grid endpoint for this end
# =============================================================================
def init_offset_state(grids, view):
    """Seed offset state by reading actual ldr.Elbow from Revit.

    Called AFTER T1 so leaders exist.  For a bubble with no leader (natural
    endpoint), perp_offset = 0.0 (bubble sits at ep)."""
    mem_offset   = {}
    mem_perp_vec = {}
    mem_ep       = {}

    for g in grids:
        pv    = compute_perp_vec(g, view)
        curve = get_grid_curve_in_view(g, view)
        gid   = _eid_int(g.Id)
        mem_perp_vec[gid] = pv

        for ei in (0, 1):
            de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
            if not grid_has_bubble_at_end(g, view, ei):
                continue
            if curve is None:
                continue
            ep  = curve.GetEndPoint(ei)
            key = (gid, ei)
            mem_ep[key] = ep

            try:
                ldr = g.GetLeader(de, view)
            except Exception:
                ldr = None

            if ldr is not None:
                elbow   = ldr.Elbow
                dot_val = ((elbow.X - ep.X) * pv.X +
                           (elbow.Y - ep.Y) * pv.Y)
                mem_offset[key] = dot_val
            else:
                mem_offset[key] = 0.0

    return mem_offset, mem_perp_vec, mem_ep


# =============================================================================
# Bubble position from offset
# =============================================================================
def anchor_xy(ep, pv, offset):
    """Return (ax, ay) for bubble anchor: ep + offset * perp_vec."""
    return ep.X + offset * pv.X, ep.Y + offset * pv.Y


# =============================================================================
# Collision detection (offset space)
# =============================================================================
def build_bubble_list(grids, view, mem_offset, mem_perp_vec, mem_ep):
    """Return list of (grid, datum_end, end_index, ax, ay)."""
    result = []
    for g in grids:
        gid = _eid_int(g.Id)
        pv  = mem_perp_vec.get(gid)
        if pv is None:
            continue
        for ei in (0, 1):
            key = (gid, ei)
            if not grid_has_bubble_at_end(g, view, ei):
                continue
            ep     = mem_ep.get(key)
            offset = mem_offset.get(key)
            if ep is None or offset is None:
                continue
            de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
            ax, ay = anchor_xy(ep, pv, offset)
            result.append((g, de, ei, ax, ay))
    return result


def find_colliding_pairs(positions, threshold):
    pairs = []
    thr2  = threshold * threshold
    n     = len(positions)
    for i in range(n):
        for j in range(i+1, n):
            if positions[i][0].Id == positions[j][0].Id:
                continue
            dx = positions[i][3] - positions[j][3]
            dy = positions[i][4] - positions[j][4]
            if dx*dx + dy*dy <= thr2:
                pairs.append((positions[i], positions[j]))
    return pairs


# =============================================================================
# Crossing guard — anchor-position based, offset space
# =============================================================================
def build_grid_lines(grids, view):
    """Return dict: gid_int → (ref_x, ref_y, tx, ty) along-grid unit vector."""
    lines = {}
    for g in grids:
        curve = get_grid_curve_in_view(g, view)
        if curve is None:
            continue
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx, dy = p1.X - p0.X, p1.Y - p0.Y
        L = (dx*dx + dy*dy) ** 0.5
        if L > 1e-9:
            lines[_eid_int(g.Id)] = (p0.X, p0.Y, dx/L, dy/L)
    return lines


def _side(ax, ay, ref_x, ref_y, tx, ty):
    """Signed side of (ax,ay) relative to line through (ref_x,ref_y) dir (tx,ty)."""
    return (ax - ref_x) * ty - (ay - ref_y) * tx


def anchor_crosses_line(cur_ax, cur_ay, prop_ax, prop_ay,
                        ref_x, ref_y, tx, ty):
    """True if anchor position moves across the infinite grid line."""
    s_cur  = _side(cur_ax,  cur_ay,  ref_x, ref_y, tx, ty)
    s_prop = _side(prop_ax, prop_ay, ref_x, ref_y, tx, ty)
    return s_cur * s_prop < 0.0


# =============================================================================
# Phase A: detect which grids need AddLeader
# =============================================================================
def detect_grids_needing_leaders(grids, view, mem_offset, mem_perp_vec,
                                 mem_ep, threshold):
    """Return set of (gid_int, ei) that collide and have no leader yet."""
    positions = build_bubble_list(grids, view, mem_offset, mem_perp_vec,
                                  mem_ep)
    pairs     = find_colliding_pairs(positions, threshold)
    if not pairs:
        return set()

    name_map = {_eid_int(g.Id): g.Name for g in grids}
    grid_map = {_eid_int(g.Id): g       for g in grids}
    movers   = set()

    for pos_a, pos_b in pairs:
        g_a, _, idx_a, _, _ = pos_a
        g_b, _, idx_b, _, _ = pos_b
        na = name_map.get(_eid_int(g_a.Id), "")
        nb = name_map.get(_eid_int(g_b.Id), "")
        if higher_name(na, nb):
            movers.add((_eid_int(g_a.Id), idx_a))
        else:
            movers.add((_eid_int(g_b.Id), idx_b))

    needs_leader = set()
    for gid, ei in movers:
        g  = grid_map.get(gid)
        de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
        if g:
            try:
                has_ldr = g.GetLeader(de, view) is not None
            except Exception:
                has_ldr = False
            if not has_ldr:
                needs_leader.add((gid, ei))

    return needs_leader


# =============================================================================
# Phase B: separation — pure in-memory
# =============================================================================
def compute_separation_in_memory(grids, view, mem_offset, mem_perp_vec,
                                 mem_ep, threshold):
    """Nudge colliding bubbles apart in perp_offset space.
    Modifies mem_offset in place."""
    nudge_step = threshold / 8.0
    name_map   = {_eid_int(g.Id): g.Name for g in grids}
    grid_lines = build_grid_lines(grids, view)

    for pass_num in range(MAX_PASSES):
        positions = build_bubble_list(grids, view, mem_offset,
                                      mem_perp_vec, mem_ep)
        pairs     = find_colliding_pairs(positions, threshold)
        if not pairs:
            output.print_md("  clean after {} pass(es)".format(pass_num + 1))
            return

        output.print_md("  **Pass {} — {} collision(s)**".format(
            pass_num + 1, len(pairs)))

        for iteration in range(MAX_ITERATIONS):
            positions = build_bubble_list(grids, view, mem_offset,
                                          mem_perp_vec, mem_ep)
            pairs     = find_colliding_pairs(positions, threshold)
            if not pairs:
                output.print_md(
                    "  Pass {} resolved in {} iteration(s)".format(
                        pass_num + 1, iteration))
                break

            # Accumulate per-mover direction votes from each collision
            targets = {}  # key → [grid_obj, ei, direction_sum]
            for pos_a, pos_b in pairs:
                g_a, _, idx_a, _, _ = pos_a
                g_b, _, idx_b, _, _ = pos_b
                na = name_map.get(_eid_int(g_a.Id), "")
                nb = name_map.get(_eid_int(g_b.Id), "")
                if higher_name(na, nb):
                    mg, mi = g_a, idx_a
                    key_m  = (_eid_int(g_a.Id), idx_a)
                    key_o  = (_eid_int(g_b.Id), idx_b)
                else:
                    mg, mi = g_b, idx_b
                    key_m  = (_eid_int(g_b.Id), idx_b)
                    key_o  = (_eid_int(g_a.Id), idx_a)

                ep_m = mem_ep.get(key_m)
                ep_o = mem_ep.get(key_o)
                pv   = mem_perp_vec.get(_eid_int(mg.Id))
                if ep_m is None or ep_o is None or pv is None:
                    continue

                # Which perpendicular side is the mover relative to the other?
                sep_dot   = ((ep_m.X - ep_o.X) * pv.X +
                             (ep_m.Y - ep_o.Y) * pv.Y)
                direction = 1 if sep_dot >= 0 else -1

                if key_m not in targets:
                    targets[key_m] = [mg, mi, 0]
                targets[key_m][2] += direction

            sorted_targets = sorted(
                targets.items(),
                key=lambda item: name_sort_key(
                    name_map.get(_eid_int(item[1][0].Id), "")),
                reverse=True,
            )

            any_moved = False
            for key, td in sorted_targets:
                gid, ei   = key
                direction = 1 if td[2] >= 0 else -1

                ep  = mem_ep.get(key)
                pv  = mem_perp_vec.get(gid)
                cur = mem_offset.get(key)
                if ep is None or pv is None or cur is None:
                    continue

                cur_ax, cur_ay = anchor_xy(ep, pv, cur)

                # Try preferred direction; if blocked by a crossing, try opposite
                for attempt_dir in (direction, -direction):
                    new_off = cur + attempt_dir * nudge_step

                    # Own-grid guard: don't cross the bubble through its own line
                    if cur > 0 and new_off < 0:
                        continue
                    if cur < 0 and new_off > 0:
                        continue

                    prop_ax, prop_ay = anchor_xy(ep, pv, new_off)

                    # Adjacent-grid crossing guard
                    crossed = False
                    for g2id, (rx, ry, tx, ty) in grid_lines.items():
                        if g2id == gid:
                            continue
                        if anchor_crosses_line(cur_ax, cur_ay,
                                               prop_ax, prop_ay,
                                               rx, ry, tx, ty):
                            crossed = True
                            break

                    if not crossed:
                        mem_offset[key] = new_off
                        any_moved = True
                        break

            if not any_moved:
                break

    else:
        output.print_md(
            "  hit MAX_PASSES ({}) — some collisions may remain".format(
                MAX_PASSES))


# =============================================================================
# Compaction — pull offsets back toward 0 (pure in-memory)
# =============================================================================
def compact_in_memory(grids, view, mem_offset, mem_perp_vec, mem_ep,
                      threshold):
    """Pull leader offsets back toward 0 without reintroducing collisions."""
    thr2        = threshold * threshold
    total_moves = 0
    grid_lines  = build_grid_lines(grids, view)

    for _ in range(MAX_COMPACT_ROUNDS):
        # Snapshot current anchor positions for this round
        all_anchors = {}
        for g in grids:
            gid = _eid_int(g.Id)
            pv  = mem_perp_vec.get(gid)
            if pv is None:
                continue
            for ei in (0, 1):
                key = (gid, ei)
                if not grid_has_bubble_at_end(g, view, ei):
                    continue
                ep     = mem_ep.get(key)
                offset = mem_offset.get(key)
                if ep is None or offset is None:
                    continue
                all_anchors[key] = anchor_xy(ep, pv, offset)

        # Candidates: any bubble with non-zero offset
        candidates = []
        for key in all_anchors:
            off = mem_offset.get(key)
            if off is not None and abs(off) >= COMPACT_PULL_STEP:
                candidates.append((abs(off), key))

        if not candidates:
            break

        candidates.sort(key=lambda c: -c[0])
        moved_any = False

        for _, key in candidates:
            gid, ei  = key
            ep       = mem_ep.get(key)
            pv       = mem_perp_vec.get(gid)
            cur      = mem_offset.get(key)
            if ep is None or pv is None or cur is None:
                continue

            step    = -COMPACT_PULL_STEP if cur > 0 else COMPACT_PULL_STEP
            new_off = cur + step
            # Clamp: don't overshoot zero
            if cur > 0 and new_off < 0:
                new_off = 0.0
            elif cur < 0 and new_off > 0:
                new_off = 0.0

            cur_ax,  cur_ay  = anchor_xy(ep, pv, cur)
            prop_ax, prop_ay = anchor_xy(ep, pv, new_off)

            # Collision check against all other current anchors
            safe = True
            for key2, (ax2, ay2) in all_anchors.items():
                if key2 == key:
                    continue
                dx = prop_ax - ax2
                dy = prop_ay - ay2
                if dx*dx + dy*dy <= thr2:
                    safe = False
                    break
            if not safe:
                continue

            # Adjacent-grid crossing guard
            crossed = False
            for g2id, (rx, ry, tx, ty) in grid_lines.items():
                if g2id == gid:
                    continue
                if anchor_crosses_line(cur_ax, cur_ay,
                                       prop_ax, prop_ay,
                                       rx, ry, tx, ty):
                    crossed = True
                    break
            if crossed:
                continue

            mem_offset[key] = new_off
            total_moves    += 1
            moved_any       = True

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

    output.print_md("## Grid Bubble Separation — v22.0.0")

    ref_grid = pick_reference_grid()
    if ref_grid is None:
        script.exit()

    output.print_md("Reference grid: **{}** (ID {})".format(
        ref_grid.Name, _eid_int(ref_grid.Id)))

    bubble_diam_ft = read_bubble_diameter_ft(ref_grid)
    threshold      = bubble_diam_ft

    views = collect_plan_views_on_sheets(doc)
    if not views:
        forms.alert("No plan views on sheets found.", title="Nothing to do")
        script.exit()

    output.print_md("Plan views on sheets: **{}**".format(len(views)))

    # ── Read Phase 1: initial offset state for all views (outside transaction)
    view_data = {}  # vid_int → (view, grids, mem_offset, mem_perp_vec, mem_ep)
    for v in views:
        grids = get_grids_for_view(doc, v)
        if len(grids) < 2:
            continue
        mo, mpv, mep = init_offset_state(grids, v)
        view_data[_eid_int(v.Id)] = (v, grids, mo, mpv, mep)

    if not view_data:
        forms.alert("No views with 2+ grids found.", title="Nothing to do")
        script.exit()

    # ── Read Phase 2: detect missing leaders for all views ───────────────────
    needs_by_view = {}
    total_needs   = 0
    for vid, (v, grids, mo, mpv, mep) in view_data.items():
        nl = detect_grids_needing_leaders(grids, v, mo, mpv, mep, threshold)
        if nl:
            needs_by_view[vid] = nl
            total_needs += len(nl)

    # ── Write T1: AddLeader only (short flat loop) ───────────────────────────
    if needs_by_view:
        output.print_md(
            "Adding {} leader(s) across {} view(s)...".format(
                total_needs, len(needs_by_view)))
        t1 = Transaction(doc, "Add Grid Leaders")
        try:
            t1.Start()
            for vid, nl in needs_by_view.items():
                v, grids, _, _, _ = view_data[vid]
                gmap = {_eid_int(g.Id): g for g in grids}
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

        # Re-read fresh offset state after T1 (now leaders exist at ep)
        for vid in list(view_data.keys()):
            v, grids, _, _, _ = view_data[vid]
            mo, mpv, mep = init_offset_state(grids, v)
            view_data[vid] = (v, grids, mo, mpv, mep)

        # Revit places AddLeader's initial Elbow at a small arbitrary offset
        # from ep.  If that offset is non-zero, the own-grid guard locks the
        # bubble to the wrong side for ~50% of cases.  Force newly-added
        # leaders to offset=0 so separation always starts from a clean slate.
        for vid, nl in needs_by_view.items():
            mo = view_data[vid][2]
            for gid_int, ei in nl:
                mo[(gid_int, ei)] = 0.0

    # ── Compute Phase: separation math in-memory (no transactions) ───────────
    output.print_md("### Separation")

    initial_offsets_by_view = {}
    for vid, (v, grids, mo, mpv, mep) in view_data.items():
        initial_offsets_by_view[vid] = dict(mo)

    for vid, (v, grids, mo, mpv, mep) in view_data.items():
        output.print_md("#### View: `{}`".format(v.Name))
        compute_separation_in_memory(grids, v, mo, mpv, mep, threshold)
        c = compact_in_memory(grids, v, mo, mpv, mep, threshold)
        if c:
            output.print_md("  Compaction: {} step(s)".format(c))

    # ── Collect changed keys ─────────────────────────────────────────────────
    changes_by_view = {}
    total_changes   = 0
    for vid, (v, grids, mo, mpv, mep) in view_data.items():
        init_off = initial_offsets_by_view[vid]
        changed  = []
        for key, final_off in mo.items():
            init_o = init_off.get(key)
            if init_o is None or abs(final_off - init_o) > 1e-6:
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

    # ── Write T2: SetLeader(Elbow) only (short flat loop) ────────────────────
    t2 = Transaction(doc, "Separate Grid Bubbles")
    applied = 0
    try:
        t2.Start()
        for vid, changed_keys in changes_by_view.items():
            v, grids, mo, mpv, mep = view_data[vid]
            gmap = {_eid_int(g.Id): g for g in grids}
            for gid, ei in changed_keys:
                g  = gmap.get(gid)
                de = DatumEnds.End0 if ei == 0 else DatumEnds.End1
                if g is None:
                    continue
                try:
                    ldr = g.GetLeader(de, v)
                    if ldr is None:
                        continue
                    ep     = mep.get((gid, ei))
                    pv     = mpv.get(gid)
                    offset = mo.get((gid, ei))
                    if ep is None or pv is None or offset is None:
                        continue
                    # Inner (along-grid) direction from ep into the grid:
                    #   pv = (-ty, tx)  →  along_grid = (tx, ty) = (pv.Y, -pv.X)
                    # For ei=1: inner points toward End0  = -(tx,ty) = (-pv.Y, pv.X)
                    # For ei=0: inner points toward End1  = +(tx,ty) = ( pv.Y,-pv.X)
                    if ei == 1:
                        ix, iy = -pv.Y, pv.X
                    else:
                        ix, iy =  pv.Y, -pv.X
                    new_elbow = XYZ(
                        ep.X + offset * pv.X + LEADER_SHOULDER_FT * ix,
                        ep.Y + offset * pv.Y + LEADER_SHOULDER_FT * iy,
                        ep.Z,
                    )
                    ldr.Elbow = new_elbow
                    g.SetLeader(de, v, ldr)
                    applied += 1
                except Exception as ex:
                    output.print_md(
                        "**SetLeader failed** gid={} ei={}: {}".format(
                            gid, ei, ex))
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
