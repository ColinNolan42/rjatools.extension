# -*- coding: ascii -*-
# duct_size_run.py
# The Size Duct pipeline. WRITES duct and fitting sizes into the model.
#
#   dialog -> pick the unit or a duct -> traverse (hvac_graph, the Duct
#   Performance engine) -> main / branch split -> size every main
#   (duct_sizing) -> write ducts -> match the fittings between two mains ->
#   audit every joint -> report
#
# Mains only (Colin, 2026-10-20). Branches, takeoff boots and diffusers are
# never touched: the auto-sizing diffusers set the branch, and Duct
# Performance reports one that is off. A fitting is only resized when every
# duct it joins is a main, so a tap (main on one side, branch on the other)
# is left alone without needing its family name.
#
# Each duct and each fitting is its own transaction inside one
# TransactionGroup (one Undo). A failure handler deletes warnings and rolls
# back only the element Revit cannot resize (for example a larger transition
# that no longer fits), so one bad joint never costs the rest and never opens
# a modal Revit dialog.
#
# IronPython 2.7 / pyRevit

from Autodesk.Revit.DB.Mechanical import MechanicalUtils
from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, ConnectorProfileType, ConnectorType,
    FailureProcessingResult, FailureSeverity, FilteredElementCollector,
    IFailuresPreprocessor, MEPSystem, Transaction, TransactionGroup,
    TransactionStatus)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

import duct_sizing
import hvac_graph
import revit_helpers
import ui_helpers
from revit_helpers import eid_int

IN_PER_FT = 12.0
SIZE_TOL_IN = 0.05


def run(doc, uidoc, output, forms):
    """Run the tool. Returns True if it finished."""
    output.print_md("# Size Duct")

    settings = ui_helpers.show_size_duct_dialog("Size Duct - RJA Standards")
    if settings is None:
        output.print_md("Cancelled. Nothing was changed.")
        return False
    # Standard calculation basis every run, so nothing set by an earlier
    # Duct Performance run in this session leaks in.
    hvac_graph.set_calc_basis()

    origin = _pick(uidoc, doc, output)
    if origin is None:
        return False

    net = hvac_graph.build_network(origin, doc, equipment_level=True)
    for warning in net.warnings:
        output.print_md(":warning: " + warning)
    if net.errors:
        for error in net.errors:
            output.print_md(":cross_mark: " + error)
        forms.alert("The traversal failed:\n\n" + "\n".join(net.errors),
                    title="Size Duct")
        return False
    output.print_md("Root: **{}** ({}). {} ducts, {} terminals.".format(
        eid_int(net.root.Id), net.ahu_method, len(net.duct_results),
        len(net.terminal_cfms)))

    plans = _plan(doc, net, settings["limits"])
    _print_plan(output, plans)

    to_write = [p for p in plans if p.new is not None and p.changed]
    if not to_write:
        if any(p.new is not None for p in plans):
            output.print_md("Every main that could be sized is already at its "
                            "size. Nothing was changed.")
        else:
            output.print_md("No main could be sized (see Not sized above). "
                            "Nothing was changed.")
        _print_basis(output, settings["limits"])
        return True

    group = TransactionGroup(doc, "RJA Tools - Size Duct")
    group.Start()
    try:
        before = _fitting_ids(doc)
        pieces = set()
        duct_failures = _write_ducts(doc, net, to_write, pieces)
        # Only ducts this run sized: a fitting is touched only when every
        # duct it joins is in here (never one next to a branch, an existing
        # or an unsized duct). Pieces made by a step-down count too.
        main_ids = set(p.element_id for p in plans if p.new is not None)
        main_ids |= pieces
        resized_ids = (set(p.element_id for p in to_write) -
                       set(duct_failures.keys())) | pieces
        fit_done, fit_failures, fit_skipped = _write_fittings(
            doc, net, resized_ids, main_ids)
        added = sorted(_fitting_ids(doc) - before)
        # The step-down transitions this run made join two sizes on purpose.
        fit_skipped = dict((k, v) for k, v in fit_skipped.items()
                           if k not in set(added))
        status = group.Assimilate()
        if status != TransactionStatus.Committed:
            raise Exception("the changes could not be committed "
                            "({})".format(status))
    except Exception as exc:
        group.RollBack()
        forms.alert("Size Duct failed and nothing was changed:\n\n" + str(exc),
                    title="Size Duct")
        output.print_md(":cross_mark: Failed, rolled back: {}".format(
            str(exc)))
        return False

    _report_writes(output, to_write, duct_failures, fit_done, fit_failures,
                   fit_skipped, added)
    _audit(doc, output, resized_ids | set(fit_done))
    _print_basis(output, settings["limits"])
    output.print_md("Run **Duct Performance** on this system to check the "
                    "result.")
    return True


# ---------------------------------------------------------------------------
# Pick
# ---------------------------------------------------------------------------

class _UnitOrDuctFilter(ISelectionFilter):
    """Only mechanical equipment and duct are clickable."""

    def AllowElement(self, elem):
        try:
            return hvac_graph.is_equipment(elem) or hvac_graph.is_duct(elem)
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


def _pick(uidoc, doc, output):
    prompt = "Select the unit (RTU, AHU, fan) or a supply or return duct"
    try:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element,
                                             _UnitOrDuctFilter(), prompt)
        except TypeError:
            ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
    except Exception:
        output.print_md("Cancelled. Nothing was changed.")
        return None
    elem = doc.GetElement(ref.ElementId)
    if elem is not None:
        output.print_md("Picked element **{}**.".format(eid_int(elem.Id)))
    return elem


# ---------------------------------------------------------------------------
# Plan
# ---------------------------------------------------------------------------

def _plan(doc, net, limits):
    root_id = eid_int(net.root.Id)
    downstream = hvac_graph.compute_downstream_terminal_ids(
        root_id, net.nodes, net.children, set(net.terminal_cfms.keys()))
    no_flow = set(net.zero_terminals) | set(net.missing_flow)
    parents = _parents(net)

    plans = []
    for eid, dr in net.duct_results.items():
        nid = eid_int(eid)
        shape, current = duct_shape_and_size(dr.elem)
        below = downstream.get(nid, set())
        plan = duct_sizing.plan_duct(
            nid, shape, current, dr.cfm, dr.sys_class, limits,
            hvac_graph.duct_friction_loss_per_100ft,
            is_branch=(len(below) == 1),
            is_flex=dr.is_flex,
            is_existing=not is_new_construction(doc, dr.elem),
            flow_missing=bool(below & no_flow))
        plan.length_ft = dr.length_ft
        plan.sections = None
        plan.step_notes = []
        if plan.new is not None:
            _plan_steps(net, dr.elem, plan, limits, parents)
        plans.append(plan)
    plans.sort(key=lambda p: (p.sys_class, -p.cfm))
    return plans


def _plan_steps(net, duct, plan, limits, parents):
    """Step-downs along one main, after taps where the airflow allows."""
    ends = _flow_ends(duct, parents.get(eid_int(duct.Id)))
    if ends is None:
        plan.step_notes.append("flow direction not read, no step-downs")
        return
    up_pt, direction = ends
    taps = _taps_on(net, duct, up_pt, direction)
    if not taps:
        return
    max_fpm, max_fric = limits[plan.sys_class]
    sections, notes = duct_sizing.plan_sections(
        plan.shape, plan.current, plan.cfm, plan.length_ft, taps, max_fpm,
        max_fric, hvac_graph.duct_friction_loss_per_100ft)
    plan.step_notes.extend(notes)
    if len(sections) > 1:
        plan.sections = sections
        plan.changed = True
        plan.up_pt = up_pt
        plan.direction = direction
        plan.taps = taps


def _flow_ends(duct, parent_id):
    """(upstream end point, unit direction downstream), or None."""
    if parent_id is None:
        return None
    ends = [c for c in _connectors(duct)
            if c.ConnectorType == ConnectorType.End]
    if len(ends) != 2:
        return None
    for i, conn in enumerate(ends):
        try:
            for ref in conn.AllRefs:
                if ref.Owner is not None and eid_int(ref.Owner.Id) == parent_id:
                    other = ends[1 - i].Origin
                    return conn.Origin, (other - conn.Origin).Normalize()
        except Exception:
            continue
    return None


def _taps_on(net, duct, up_pt, direction):
    """Every takeoff on the duct's side (curve connectors), as Taps."""
    taps = []
    own = eid_int(duct.Id)
    for conn in _connectors(duct):
        if conn.ConnectorType != ConnectorType.Curve:
            continue
        try:
            refs = list(conn.AllRefs)
        except Exception:
            continue
        for ref in refs:
            tap = ref.Owner
            if tap is None or isinstance(tap, MEPSystem) or \
                    eid_int(tap.Id) == own:
                continue
            pos = (ref.Origin - up_pt).DotProduct(direction)
            main_size = _conn_size_in(ref)
            branch_size = 0.0
            for other in _connectors(tap):
                if other.Origin.DistanceTo(ref.Origin) > 0.01:
                    branch_size = max(branch_size, _conn_size_in(other))
            try:
                on_top = abs(ref.CoordinateSystem.BasisZ.Z) > 0.7
            except Exception:
                on_top = True
            taps.append(duct_sizing.Tap(
                pos, (main_size / 2.0 + 2.0) / IN_PER_FT,
                net.cfm_map.get(eid_int(tap.Id), 0.0),
                branch_size or main_size, on_top))
    return taps


def _conn_size_in(conn):
    """Largest dimension of a connector, inches."""
    try:
        if conn.Shape == ConnectorProfileType.Round:
            return conn.Radius * 2.0 * IN_PER_FT
        return max(conn.Width, conn.Height) * IN_PER_FT
    except Exception:
        return 0.0


def _parents(net):
    parents = {}
    for parent_id, kids in net.children.items():
        for kid in kids:
            parents[kid] = parent_id
    return parents


def is_new_construction(doc, elem):
    """True only when the element was created in the project's last phase.
    Fails closed: an unreadable phase is treated as existing, never sized."""
    last_phase = revit_helpers.get_last_phase_id(doc)
    created = revit_helpers.get_created_phase_id(elem)
    return (last_phase is not None and created is not None and
            created == last_phase)


def duct_shape_and_size(duct):
    """(shape, inches) as drawn: round d, rect (w, h), or (None, None)."""
    try:
        for conn in duct.ConnectorManager.Connectors:
            if conn.Shape == ConnectorProfileType.Oval:
                return None, None
            break
    except Exception:
        pass
    dia = _param_ft(duct, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if dia:
        return duct_sizing.SHAPE_ROUND, dia * IN_PER_FT
    w = _param_ft(duct, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    h = _param_ft(duct, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    if w and h:
        return duct_sizing.SHAPE_RECT, (w * IN_PER_FT, h * IN_PER_FT)
    return None, None


def _param_ft(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p is not None and p.HasValue:
            value = p.AsDouble()
            return value if value > 0 else None
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Write
# ---------------------------------------------------------------------------

class _Failures(IFailuresPreprocessor):
    """Delete warnings; on an error, roll back this one element and keep the
    message so it can be reported. Never shows a Revit dialog."""

    def __init__(self):
        self.errors = []

    def PreprocessFailures(self, accessor):
        rollback = False
        for message in accessor.GetFailureMessages():
            if message.GetSeverity() == FailureSeverity.Warning:
                accessor.DeleteWarning(message)
            else:
                rollback = True
                try:
                    self.errors.append(message.GetDescriptionText())
                except Exception:
                    self.errors.append("Revit error")
        if rollback:
            return FailureProcessingResult.ProceedWithRollBack
        return FailureProcessingResult.Continue


def _one(doc, name, action):
    """Run action() in its own transaction. Returns None or the error text."""
    t = Transaction(doc, name)
    handler = _Failures()
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(handler)
    options.SetClearAfterRollback(True)
    t.SetFailureHandlingOptions(options)
    t.Start()
    try:
        action()
        doc.Regenerate()
    except Exception as exc:
        t.RollBack()
        return _text(exc)
    try:
        status = t.Commit()
    except Exception as exc:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        return _text(exc)
    if status != TransactionStatus.Committed:
        return "; ".join(handler.errors) or "rolled back by Revit"
    return None


def _text(exc):
    try:
        return str(exc)
    except Exception:
        return repr(exc)


def _write_ducts(doc, net, plans, pieces):
    """Resize each planned duct, splitting it at its step-downs.

    A duct whose step-downs fail (no room, a transition that will not fit,
    Revit refusing the split) is rolled back and sized whole at its upstream
    size instead, so the main is never left undersized. New piece ids go into
    `pieces`. Returns {duct id: error}.
    """
    failures = {}
    for plan in plans:
        duct = net.nodes.get(plan.element_id)
        if duct is None:
            failures[plan.element_id] = "element not found"
            continue

        if plan.sections:
            made = []

            def split(duct=duct, plan=plan, made=made):
                _split_and_size(doc, duct, plan, made)

            error = _one(doc, "Size Duct steps {}".format(plan.element_id),
                         split)
            if not error:
                pieces.update(made)
                continue
            plan.step_notes.append("step-downs rolled back ({}), sized "
                                   "whole".format(error))
            plan.sections = None

        def action(duct=duct, plan=plan):
            _apply_size(duct, plan.shape, plan.new)

        error = _one(doc, "Size Duct {}".format(plan.element_id), action)
        if error:
            failures[plan.element_id] = error
    return failures


def _apply_size(duct, shape, size):
    if shape == duct_sizing.SHAPE_ROUND:
        _set(duct, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
             size / IN_PER_FT)
    else:
        _set(duct, BuiltInParameter.RBS_CURVE_WIDTH_PARAM,
             size[0] / IN_PER_FT)
        _set(duct, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
             size[1] / IN_PER_FT)


def _split_and_size(doc, duct, plan, made):
    """Break the duct at each step-down, size each piece, and put a
    transition between neighbours. Raises on anything that does not come out
    clean, so the caller's transaction rolls the whole duct back.

    Live-tested technique (2026-10-20): MechanicalUtils.BreakCurve leaves the
    two new ends unconnected and keeps each tap on its own piece;
    NewTransitionFitting then joins them and sizes itself from the ducts,
    taking its length out of the upstream piece.
    """
    up_pt, direction = plan.up_pt, plan.direction

    def far(piece):
        curve = piece.Location.Curve
        return max((curve.GetEndPoint(i) - up_pt).DotProduct(direction)
                   for i in (0, 1))

    ordered = [duct]
    for section in plan.sections[1:]:
        point = up_pt + direction * section.start_ft
        current = ordered[-1]
        new_id = MechanicalUtils.BreakCurve(doc, current.Id, point)
        new = doc.GetElement(new_id)
        if new is None:
            raise Exception("BreakCurve made no new piece")
        if far(new) > far(current):
            ordered.append(new)
        else:
            ordered[-1] = new
            ordered.append(current)
    doc.Regenerate()

    for piece, section in zip(ordered, plan.sections):
        _apply_size(piece, plan.shape, section.size)
    doc.Regenerate()

    transitions = []
    for a, b in zip(ordered, ordered[1:]):
        ca, cb = _facing_ends(a, b)
        if ca is None:
            raise Exception("the ends at a step-down could not be found")
        fitting = doc.Create.NewTransitionFitting(ca, cb)
        if fitting is None:
            raise Exception("Revit made no transition")
        transitions.append(fitting)
    doc.Regenerate()

    # The transition is cut out of the upstream piece: it shall not reach a
    # tap fitting.
    for fitting in transitions:
        span = [(c.Origin - up_pt).DotProduct(direction)
                for c in _connectors(fitting)]
        lo, hi = min(span), max(span)
        for tap in plan.taps:
            if tap.pos_ft + tap.half_ft > lo - 0.05 and \
                    tap.pos_ft - tap.half_ft < hi + 0.05:
                raise Exception("a transition would overlap a tap")

    made.extend(eid_int(p.Id) for p in ordered)


def _facing_ends(a, b):
    """The unconnected end connectors where pieces a and b meet."""
    best = None
    for ca in _connectors(a):
        if ca.ConnectorType != ConnectorType.End or ca.IsConnected:
            continue
        for cb in _connectors(b):
            if cb.ConnectorType != ConnectorType.End or cb.IsConnected:
                continue
            d = ca.Origin.DistanceTo(cb.Origin)
            if best is None or d < best[0]:
                best = (d, ca, cb)
    if best is None or best[0] > 0.05:
        return None, None
    return best[1], best[2]


def _set(elem, bip, value_ft):
    p = elem.get_Parameter(bip)
    if p is None or p.IsReadOnly:
        raise Exception("parameter {} not writable".format(bip))
    p.Set(value_ft)


# Instance size parameters on the standard Revit duct fitting and accessory
# families (elbows, endcaps, dampers), confirmed live on 2026-10-16/19.
_RECT_PARAMS = ("Duct Width", "Duct Height")
_ROUND_DIAMETER = "Duct Diameter"
_ROUND_RADIUS = "Duct Radius"


def _write_fittings(doc, net, resized_ids, main_ids):
    """Match each fitting or accessory joined to a resized main.

    Only fittings whose every connected duct is a main (elbows, couplings,
    dampers, endcaps, tees along the main). A fitting that joins a branch is
    a tap and is never touched. Fittings that join two different sizes
    (transitions) are left for Revit and reported by the audit.

    Returns (done ids, {id: error}, {id: reason skipped}).
    """
    done, failures, skipped = [], {}, {}
    seen = set()
    for duct_id in resized_ids:
        duct = net.nodes.get(duct_id)
        if duct is None:
            continue
        for neighbour in _neighbours(duct):
            nid = eid_int(neighbour.Id)
            if nid in seen or hvac_graph.is_duct(neighbour):
                continue
            seen.add(nid)
            if not (hvac_graph.is_fitting(neighbour) or
                    hvac_graph.is_accessory(neighbour)):
                continue
            if not is_new_construction(doc, neighbour):
                skipped[nid] = "existing phase"
                continue
            ducts = [d for d in _neighbours(neighbour)
                     if hvac_graph.is_duct(d)]
            if any(eid_int(d.Id) not in main_ids for d in ducts):
                continue        # a tap or anything touching a branch
            sizes = set()
            for d in ducts:
                shape, size = duct_shape_and_size(d)
                sizes.add((shape, _key(size)))
            if len(sizes) != 1:
                skipped[nid] = "joins different sizes (transition)"
                continue
            shape, size = list(sizes)[0]
            if shape is None:
                skipped[nid] = "oval or unknown"
                continue
            if _fitting_matches(neighbour, shape, size):
                continue

            def action(fit=neighbour, shape=shape, size=size):
                _size_fitting(fit, shape, size, doc)

            error = _one(doc, "Size Duct fitting {}".format(nid), action)
            if error:
                failures[nid] = error
            else:
                done.append(nid)
    return done, failures, skipped


def _key(size):
    if size is None:
        return None
    if isinstance(size, tuple):
        return (round(size[0], 2), round(size[1], 2))
    return round(size, 2)


def _size_fitting(fit, shape, size, doc=None):
    """Set a fitting's instance size parameters to the duct size, then check
    its connectors took it. A rectangular fitting whose Width and Height
    parameters run the other way round is tried once swapped; one that still
    does not match raises, so its transaction rolls back."""
    if shape == duct_sizing.SHAPE_RECT:
        for order in (size, (size[1], size[0])):
            for name, value in zip(_RECT_PARAMS, order):
                p = fit.LookupParameter(name)
                if p is None or p.IsReadOnly:
                    raise Exception("no writable '{}' parameter".format(name))
                p.Set(value / IN_PER_FT)
            if doc is None:
                return
            doc.Regenerate()
            if _fitting_matches(fit, shape, size):
                return
        raise Exception("connectors did not take {:.0f}x{:.0f}".format(
            float(size[0]), float(size[1])))
    p = fit.LookupParameter(_ROUND_DIAMETER)
    if p is not None and not p.IsReadOnly:
        p.Set(size / IN_PER_FT)
        return
    p = fit.LookupParameter(_ROUND_RADIUS)
    if p is not None and not p.IsReadOnly:
        p.Set(size / 2.0 / IN_PER_FT)
        return
    raise Exception("no writable Duct Diameter or Duct Radius parameter")


def _fitting_matches(fit, shape, size):
    for conn in _connectors(fit):
        if not _connector_matches(conn, shape, size):
            return False
    return True


def _connector_matches(conn, shape, size):
    try:
        if shape == duct_sizing.SHAPE_ROUND:
            return abs(conn.Radius * 2.0 * IN_PER_FT - size) <= SIZE_TOL_IN
        return (abs(conn.Width * IN_PER_FT - size[0]) <= SIZE_TOL_IN and
                abs(conn.Height * IN_PER_FT - size[1]) <= SIZE_TOL_IN)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Connectors
# ---------------------------------------------------------------------------

def _connectors(elem):
    try:
        manager = elem.ConnectorManager
    except Exception:
        manager = None
    if manager is None:
        try:
            manager = elem.MEPModel.ConnectorManager
        except Exception:
            return []
    try:
        return list(manager.Connectors)
    except Exception:
        return []


def _neighbours(elem):
    """Elements physically joined to elem (MEP systems excluded)."""
    found = []
    own = eid_int(elem.Id)
    for conn in _connectors(elem):
        try:
            refs = list(conn.AllRefs)
        except Exception:
            continue
        for ref in refs:
            owner = ref.Owner
            if owner is None or isinstance(owner, MEPSystem):
                continue
            if eid_int(owner.Id) == own:
                continue
            found.append(owner)
    return found


def _fitting_ids(doc):
    ids = set()
    for bic in (BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory):
        for e in (FilteredElementCollector(doc).OfCategory(bic)
                  .WhereElementIsNotElementType()):
            ids.add(eid_int(e.Id))
    return ids


# ---------------------------------------------------------------------------
# Audit and report
# ---------------------------------------------------------------------------

def _audit(doc, output, resized_ids):
    """Every joint of every resized duct: connected? sizes agree?"""
    problems = []
    for duct_id in sorted(resized_ids):
        duct = revit_helpers.get_element_by_id(doc, duct_id)
        if duct is None:
            problems.append([str(duct_id), "-", "element no longer exists"])
            continue
        for conn in _connectors(duct):
            try:
                if not conn.IsConnected:
                    problems.append([str(duct_id), "-", "open end"])
                    continue
                for ref in conn.AllRefs:
                    owner = ref.Owner
                    if owner is None or isinstance(owner, MEPSystem) or \
                            eid_int(owner.Id) == duct_id:
                        continue
                    if not _same_size(conn, ref):
                        problems.append([str(duct_id), str(eid_int(owner.Id)),
                                         "size mismatch at the joint"])
            except Exception as exc:
                problems.append([str(duct_id), "-", str(exc)])
    output.print_md("## Joint audit")
    if problems:
        output.print_table(table_data=problems,
                           title="Joints to check by hand",
                           columns=["Element", "Joined to", "Problem"])
    else:
        output.print_md("Every joint on the resized ducts and fittings is "
                        "connected and the sizes agree.")


def _same_size(a, b):
    try:
        if a.Shape != b.Shape:
            return False
        if a.Shape == ConnectorProfileType.Round:
            return abs(a.Radius - b.Radius) * 2.0 * IN_PER_FT <= SIZE_TOL_IN
        return (abs(a.Width - b.Width) * IN_PER_FT <= SIZE_TOL_IN and
                abs(a.Height - b.Height) * IN_PER_FT <= SIZE_TOL_IN)
    except Exception:
        return False


def _print_plan(output, plans):
    sized = [p for p in plans if p.new is not None]
    changed = [p for p in sized if p.changed]
    output.print_md("## Mains sized: {} ({} to change)".format(
        len(sized), len(changed)))
    rows = []
    for p in sized:
        size = duct_sizing.size_label(p.shape, p.new)
        if p.sections:
            size = " > ".join(duct_sizing.size_label(p.shape, s.size)
                              for s in p.sections)
        rows.append([str(p.element_id), p.sys_class, "{:.0f}".format(p.cfm),
                     duct_sizing.size_label(p.shape, p.current), size,
                     "{:.0f}".format(p.fpm), "{:.3f}".format(p.friction),
                     "CHANGE" if p.changed else "ok"])
    notes = ["- duct {}: {}".format(p.element_id, n) for p in sized
             for n in getattr(p, "step_notes", [])]
    if notes:
        output.print_md("Step-down notes:\n" + "\n".join(notes))
    if rows:
        output.print_table(table_data=rows, title="Mains",
                           columns=["Duct", "System", "CFM", "Drawn", "Size",
                                    "FPM", "in/100ft", ""])
    skipped = [p for p in plans if p.new is None]
    counts = {}
    for p in skipped:
        counts[p.skip_reason] = counts.get(p.skip_reason, 0) + 1
    if counts:
        output.print_md("Not sized: " + "; ".join(
            "{} {}".format(n, reason) for reason, n in sorted(counts.items())))


def _report_writes(output, to_write, duct_failures, fit_done, fit_failures,
                   fit_skipped, added):
    output.print_md("## Written")
    output.print_md("- Ducts resized: **{}** of {}".format(
        len(to_write) - len(duct_failures), len(to_write)))
    output.print_md("- Fittings and accessories matched: **{}**".format(
        len(fit_done)))
    rows = []
    for eid, err in sorted(duct_failures.items()):
        rows.append(["duct", str(eid), err])
    for eid, err in sorted(fit_failures.items()):
        rows.append(["fitting", str(eid), err])
    for eid, why in sorted(fit_skipped.items()):
        rows.append(["fitting", str(eid), "left for Revit: " + why])
    if rows:
        output.print_table(table_data=rows, title="Not resized",
                           columns=["Kind", "Element", "Why"])
    if added:
        output.print_md(":warning: {} fitting(s) were added: the step-down "
                        "transitions, plus any Revit inserted at elbows: {}. "
                        "Check each one fits.".format(len(added), ", ".join(
                            str(i) for i in added[:30])))


def _print_basis(output, limits):
    output.print_md("## Basis of design")
    output.print_md("\n".join("    " + line
                              for line in duct_sizing.basis_lines(limits)))
