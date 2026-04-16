# -*- coding: utf-8 -*-
__title__   = "Leader\nDiagnostic"
__author__  = "MEP Tools"
__version__ = "diag2"
__doc__     = "Prints AddLeader default geometry and grid 4/5 details. Rolls back."

import traceback
from Autodesk.Revit.DB import (
    FilteredElementCollector, Grid, View, ViewSheet, ViewType,
    XYZ, Transaction, DatumExtentType, DatumEnds, ElementId,
    RevitLinkInstance,
)
from pyrevit import forms, script, revit

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()

PLAN_VIEW_TYPES = {
    ViewType.FloorPlan, ViewType.CeilingPlan,
    ViewType.AreaPlan,  ViewType.EngineeringPlan,
}

def get_sheet_view_ids():
    ids = set()
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements():
        try:
            for vid in sheet.GetAllPlacedViews():
                ids.add(vid.IntegerValue)
        except Exception:
            pass
    return ids

def get_curve(grid, view):
    for et in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        try:
            curves = grid.GetCurvesInView(et, view)
            if curves:
                return curves[0], et
        except Exception:
            continue
    return None, None

def fmt(pt):
    if pt is None:
        return "None"
    return "({:.4f}, {:.4f}, {:.4f})".format(pt.X, pt.Y, pt.Z)

def main():
    # Get first plan view on a sheet
    sheet_ids = get_sheet_view_ids()
    test_view = None
    for v in FilteredElementCollector(doc).OfClass(View):
        if not v.IsTemplate and v.ViewType in PLAN_VIEW_TYPES:
            if v.Id.IntegerValue in sheet_ids:
                test_view = v
                break

    if test_view is None:
        forms.alert("No plan views on sheets found.")
        script.exit()

    output.print_md("## Testing in view: **{}**  (Scale: {})".format(
        test_view.Name, test_view.Scale))

    # Get grids 4 and 5 specifically
    g4 = doc.GetElement(ElementId(3707752))  # grid 4
    g5 = doc.GetElement(ElementId(3707753))  # grid 5

    output.print_md("\n## Grid 4 (3707752) geometry")
    if g4:
        c4, et4 = get_curve(g4, test_view)
        if c4:
            output.print_md("- Extent type: {}".format(et4))
            output.print_md("- End0: {}".format(fmt(c4.GetEndPoint(0))))
            output.print_md("- End1: {}".format(fmt(c4.GetEndPoint(1))))
            output.print_md("- Bubble End0 visible: {}".format(
                g4.IsBubbleVisibleInView(DatumEnds.End0, test_view)))
            output.print_md("- Bubble End1 visible: {}".format(
                g4.IsBubbleVisibleInView(DatumEnds.End1, test_view)))
            output.print_md("- Has leader End0: {}".format(
                g4.GetLeader(DatumEnds.End0, test_view) is not None))
            output.print_md("- Has leader End1: {}".format(
                g4.GetLeader(DatumEnds.End1, test_view) is not None))

    output.print_md("\n## Grid 5 (3707753) geometry")
    if g5:
        c5, et5 = get_curve(g5, test_view)
        if c5:
            output.print_md("- Extent type: {}".format(et5))
            output.print_md("- End0: {}".format(fmt(c5.GetEndPoint(0))))
            output.print_md("- End1: {}".format(fmt(c5.GetEndPoint(1))))
            output.print_md("- Bubble End0 visible: {}".format(
                g5.IsBubbleVisibleInView(DatumEnds.End0, test_view)))
            output.print_md("- Bubble End1 visible: {}".format(
                g5.IsBubbleVisibleInView(DatumEnds.End1, test_view)))
            output.print_md("- Has leader End0: {}".format(
                g5.GetLeader(DatumEnds.End0, test_view) is not None))
            output.print_md("- Has leader End1: {}".format(
                g5.GetLeader(DatumEnds.End1, test_view) is not None))

    output.print_md("\n## Linked models in project")
    try:
        links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        output.print_md("- Link count: **{}**".format(len(links)))
        for lnk in links:
            try:
                ld = lnk.GetLinkDocument()
                name = ld.Title if ld else "NOT LOADED"
                output.print_md("  - {}".format(name))
                if ld:
                    lg = FilteredElementCollector(ld).OfClass(Grid).ToElements()
                    output.print_md("    - Grids in link: {}".format(len(lg)))
            except Exception as ex:
                output.print_md("  - Error: {}".format(ex))
    except Exception as ex:
        output.print_md("- Link error: {}".format(ex))

    output.print_md("\n## AddLeader default geometry test (ROLLED BACK)")
    t = Transaction(doc, "Leader Diag — ROLLBACK")
    try:
        t.Start()
        if g5 and c5:
            try:
                g5.AddLeader(DatumEnds.End1, test_view)
                leader = g5.GetLeader(DatumEnds.End1, test_view)
                if leader:
                    output.print_md("AddLeader SUCCESS — default values:")
                    output.print_md("- Anchor: {}".format(fmt(leader.Anchor)))
                    output.print_md("- Elbow:  {}".format(fmt(leader.Elbow)))
                    output.print_md("- End:    {}".format(fmt(leader.End)))

                    # Now try minimal SetLeader — just shift End along axis
                    p0 = c5.GetEndPoint(0)
                    p1 = c5.GetEndPoint(1)
                    dx = p1.X - p0.X
                    dy = p1.Y - p0.Y
                    ln = (dx*dx + dy*dy)**0.5
                    ux = dx/ln; uy = dy/ln
                    # End1 is the bubble — extend further along End1 direction
                    bubble = c5.GetEndPoint(1)
                    new_end = XYZ(bubble.X + ux*3.0,
                                  bubble.Y + uy*3.0,
                                  bubble.Z)
                    # Use default Elbow from AddLeader, only change End
                    leader.End = new_end
                    output.print_md("\nTrying SetLeader with End only moved along axis:")
                    output.print_md("- New End: {}".format(fmt(new_end)))
                    try:
                        g5.SetLeader(DatumEnds.End1, test_view, leader)
                        output.print_md("- SetLeader (axis only): **SUCCESS**")
                    except Exception as ex:
                        output.print_md("- SetLeader (axis only): **FAILED** — {}".format(ex))

                    # Try with both End and Elbow on axis
                    new_elbow = XYZ(bubble.X + ux*1.5,
                                    bubble.Y + uy*1.5,
                                    bubble.Z)
                    leader.Elbow = new_elbow
                    leader.End   = new_end
                    output.print_md("\nTrying SetLeader with both End+Elbow on axis:")
                    try:
                        g5.SetLeader(DatumEnds.End1, test_view, leader)
                        output.print_md("- SetLeader (both on axis): **SUCCESS**")
                    except Exception as ex:
                        output.print_md("- SetLeader (both on axis): **FAILED** — {}".format(ex))

                else:
                    output.print_md("GetLeader returned None after AddLeader")
            except Exception as ex:
                output.print_md("AddLeader FAILED: {}".format(ex))
                output.print_md(traceback.format_exc())
        t.RollBack()
        output.print_md("\nTransaction rolled back — no permanent changes.")
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        output.print_md("Transaction error: {}".format(ex))

if __name__ == "__main__":
    main()