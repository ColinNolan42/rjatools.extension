# encoding: utf-8
"""Pipe Takeoffs - Toggle tool for automated domestic water stub-outs.

Automates branch pipe takeoff placement from domestic water mains.
Click 1: Select existing CW/HW/HWC main pipe
Click 2: Pick destination point (fixture/wall location)
Script builds: tee, 6in rise, elbow, horizontal run, elbow, drop to AFF,
               elbow turning toward main, 6in stub-in.

Assembly: 4 pipe segments + 1 tee + 3 elbows
All branch pipe properties copied from the clicked main pipe.

Press ESC or click the button again to deactivate.
"""

# ============================================================================
# IMPORTS
# ============================================================================
import clr
import math
from collections import OrderedDict

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')

from System import Array, Type
from System.Collections.Generic import List

# Revit API - explicit imports only, no wildcards
from Autodesk.Revit.DB import (
    XYZ,
    Line,
    ElementId,
    BuiltInParameter,
    BuiltInCategory,
    FilteredElementCollector,
    FamilySymbol,
    Transaction,
    StructuralType
)
from Autodesk.Revit.DB.Plumbing import (
    Pipe,
    PipingSystemType,
    PipeType,
    RoutingPreferenceRuleGroupType
)
from Autodesk.Revit.DB.Mechanical import MechanicalUtils
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import (
    OperationCanceledException,
    InvalidOperationException
)

# pyRevit imports
from pyrevit import revit, DB, UI, script, forms

# ============================================================================
# CONSTANTS
# ============================================================================
ENVVAR_ACTIVE = "PIPE_TAKEOFFS_ACTIVE"
ENVVAR_SIZE = "PIPE_TAKEOFFS_SIZE"
ENVVAR_AFF = "PIPE_TAKEOFFS_AFF"

RISE_HEIGHT = 0.5       # 6 inches in feet
STUB_LENGTH = 0.5       # 6 inches in feet
DEFAULT_AFF = 3.0       # 36 inches AFF in feet
DIAGONAL_WARN_DEG = 5.0 # warn if main is more than 5 deg off axis

# Pipe sizes: display label to nominal diameter in inches
PIPE_SIZES = OrderedDict([
    ('1/4"',   0.25),
    ('3/8"',   0.375),
    ('1/2"',   0.5),
    ('3/4"',   0.75),
    ('1"',     1.0),
    ('1-1/4"', 1.25),
    ('1-1/2"', 1.5),
    ('2"',     2.0),
    ('2-1/2"', 2.5),
    ('3"',     3.0),
])

# AFF height presets in inches
AFF_PRESETS = OrderedDict([
    ('18"',  18.0),
    ('24"',  24.0),
    ('30"',  30.0),
    ('36"',  36.0),
    ('42"',  42.0),
    ('48"',  48.0),
    ('54"',  54.0),
    ('60"',  60.0),
])

# Valid domestic water system abbreviations
VALID_SYSTEMS = ["CW", "HW", "HWC"]

# Document references
doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()


# ============================================================================
# SELECTION FILTER - CW / HW / HWC pipes only
# ============================================================================
class WaterPipeFilter(ISelectionFilter):
    """Allows selection of domestic water pipes only (CW, HW, HWC)."""

    def AllowElement(self, element):
        # Must be in pipe category
        cat = element.Category
        if cat is None:
            return False
        if cat.Id.IntegerValue != int(BuiltInCategory.OST_PipeCurves):
            return False

        # Read system type abbreviation
        sys_param = element.get_Parameter(
            BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
        )
        if sys_param is None:
            return False

        sys_type_id = sys_param.AsElementId()
        if sys_type_id == ElementId.InvalidElementId:
            return False

        sys_type = doc.GetElement(sys_type_id)
        if sys_type is None:
            return False

        # Try abbreviation first, then type name
        sys_name = ""
        abbr_param = sys_type.get_Parameter(
            BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM
        )
        if abbr_param and abbr_param.AsString():
            sys_name = abbr_param.AsString().strip().upper()
        else:
            name_param = sys_type.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME
            )
            if name_param and name_param.AsString():
                sys_name = name_param.AsString().strip().upper()

        for valid in VALID_SYSTEMS:
            if valid in sys_name:
                return True
        return False

    def AllowReference(self, reference, position):
        return False


# ============================================================================
# GEOMETRY HELPERS
# ============================================================================
def project_point_onto_line(point, line_start, line_end):
    """Project a point onto a line segment. Returns closest XYZ on line."""
    line_vec = line_end - line_start
    point_vec = point - line_start
    line_length_sq = line_vec.DotProduct(line_vec)

    if line_length_sq < 1e-10:
        return line_start

    t = point_vec.DotProduct(line_vec) / line_length_sq
    t = max(0.0, min(1.0, t))

    return XYZ(
        line_start.X + t * line_vec.X,
        line_start.Y + t * line_vec.Y,
        line_start.Z + t * line_vec.Z
    )


def get_perpendicular_toward_target(main_dir, tee_point, target_point):
    """Return normalized XY perpendicular from tee toward target."""
    perp_a = XYZ(-main_dir.Y, main_dir.X, 0)
    perp_b = XYZ(main_dir.Y, -main_dir.X, 0)

    to_target = XYZ(
        target_point.X - tee_point.X,
        target_point.Y - tee_point.Y,
        0
    )

    if perp_a.DotProduct(to_target) >= 0:
        return perp_a.Normalize()
    else:
        return perp_b.Normalize()


def check_diagonal_main(main_dir):
    """Return warning string if main is more than 5 deg off-axis, else None."""
    abs_x = abs(main_dir.X)
    abs_y = abs(main_dir.Y)

    if abs_x >= abs_y:
        off_axis_rad = math.atan2(abs_y, abs_x)
    else:
        off_axis_rad = math.atan2(abs_x, abs_y)

    off_axis_deg = math.degrees(off_axis_rad)

    if off_axis_deg > DIAGONAL_WARN_DEG:
        return (
            "Main pipe is {:.1f} deg off-axis.\n"
            "Branch will run perpendicular to the main, "
            "which may not be parallel to walls.\n\n"
            "Continue anyway?"
        ).format(off_axis_deg)
    return None


def get_open_connector_closest_to(element, target_point):
    """Find the nearest unconnected connector on element to target point."""
    best = None
    best_dist = float('inf')
    conn_set = element.ConnectorManager.Connectors
    for conn in conn_set:
        if conn.IsConnected:
            continue
        d = conn.Origin.DistanceTo(target_point)
        if d < best_dist:
            best_dist = d
            best = conn
    return best


# ============================================================================
# PROPERTY COPY FROM MAIN PIPE
# ============================================================================
def copy_main_properties(pipe):
    """Read and return all properties needed to create branch pipes.

    Copies pipe type, system type, and level directly from the clicked
    main. Branch pipes inherit the main's routing preferences, so
    fittings (tee, elbow) resolve from the same family library.

    Args:
        pipe: the clicked main Pipe element

    Returns:
        dict with keys: pipe_type_id, system_type_id, level_id,
                        start, end, direction, centerline_z,
                        diameter, radius, pipe_type_name, element
    """
    location = pipe.Location
    if location is None:
        raise ValueError("Selected pipe has no location curve.")

    curve = location.Curve
    start = curve.GetEndPoint(0)
    end_pt = curve.GetEndPoint(1)

    # Direction vector - flatten to XY plane and normalize
    raw_dir = end_pt - start
    main_dir = XYZ(raw_dir.X, raw_dir.Y, 0).Normalize()

    # Verify horizontal
    z_diff = abs(end_pt.Z - start.Z)
    if z_diff > 0.1:
        raise ValueError(
            "Selected pipe is not horizontal. "
            "Z difference: {:.2f} ft. "
            "Select a horizontal main.".format(z_diff)
        )

    # ---- Copy these three IDs directly ----
    pipe_type_id = pipe.GetTypeId()
    
    sys_param = pipe.get_Parameter(
        BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
    )
    system_type_id = sys_param.AsElementId() if sys_param else None

    lvl_param = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM)
    level_id = lvl_param.AsElementId() if lvl_param else None
    if level_id is None or level_id == ElementId.InvalidElementId:
        level_id = pipe.ReferenceLevel.Id

    # Diameter in feet
    dia_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    if dia_param is None:
        raise ValueError("Cannot read pipe diameter.")
    diameter = dia_param.AsDouble()
    radius = diameter / 2.0

    # Get pipe type name for error messages
    pipe_type_elem = doc.GetElement(pipe_type_id)
    pipe_type_name = "Unknown"
    if pipe_type_elem:
        name_param = pipe_type_elem.get_Parameter(
            BuiltInParameter.ALL_MODEL_TYPE_NAME
        )
        if name_param and name_param.AsString():
            pipe_type_name = name_param.AsString()

    return {
        'pipe_type_id': pipe_type_id,
        'system_type_id': system_type_id,
        'level_id': level_id,
        'start': start,
        'end': end_pt,
        'direction': main_dir,
        'centerline_z': start.Z,
        'diameter': diameter,
        'radius': radius,
        'pipe_type_name': pipe_type_name,
        'element': pipe,
    }


# ============================================================================
# ROUTING PREFERENCE PRE-CHECK
# ============================================================================
def check_routing_preferences(pipe_type_id, branch_dia_ft, pipe_type_name):
    """Verify the main's pipe type has fittings defined at branch diameter.

    Checks routing preferences for tee and elbow families at the
    user-selected branch size BEFORE creating any elements. Returns
    a list of warnings (empty list = all good).

    Args:
        pipe_type_id: ElementId of the pipe type (copied from main)
        branch_dia_ft: branch diameter in feet
        pipe_type_name: display name for error messages

    Returns:
        list of warning strings (empty if all checks pass)
    """
    warnings = []

    pipe_type_elem = doc.GetElement(pipe_type_id)
    if pipe_type_elem is None:
        warnings.append("Cannot read pipe type from main pipe.")
        return warnings

    # Get RoutingPreferenceManager
    rpm = None
    try:
        rpm = pipe_type_elem.RoutingPreferenceManager
    except Exception:
        pass

    if rpm is None:
        warnings.append(
            "Pipe type '{}' has no routing preferences defined. "
            "Fittings may not place correctly.".format(pipe_type_name)
        )
        return warnings

    # Check for tee rule
    try:
        tee_rule_count = rpm.GetNumberOfRules(
            RoutingPreferenceRuleGroupType.Junctions
        )
        if tee_rule_count == 0:
            warnings.append(
                "No tee fitting rule in routing preferences for '{}'. "
                "Tee placement will fail.".format(pipe_type_name)
            )
    except Exception:
        warnings.append(
            "Cannot read tee rules from routing preferences for '{}'."
            .format(pipe_type_name)
        )

    # Check for elbow rule
    try:
        elbow_rule_count = rpm.GetNumberOfRules(
            RoutingPreferenceRuleGroupType.Elbows
        )
        if elbow_rule_count == 0:
            warnings.append(
                "No elbow fitting rule in routing preferences for '{}'. "
                "Elbow placement will fail.".format(pipe_type_name)
            )
    except Exception:
        warnings.append(
            "Cannot read elbow rules from routing preferences for '{}'."
            .format(pipe_type_name)
        )

    # Check for pipe segment rule at branch diameter
    try:
        seg_rule_count = rpm.GetNumberOfRules(
            RoutingPreferenceRuleGroupType.Segments
        )
        if seg_rule_count == 0:
            warnings.append(
                "No pipe segment rule in routing preferences for '{}'. "
                "Branch pipes may use wrong segment type."
                .format(pipe_type_name)
            )
    except Exception:
        pass  # Non-critical, pipes will still create

    return warnings


# ============================================================================
# GEOMETRY CALCULATION
# ============================================================================
def calculate_takeoff_geometry(props, click1, click2, aff_height):
    """Calculate all key points for the 4-segment takeoff assembly.

    Args:
        props: dict from copy_main_properties()
        click1: XYZ from PickObject (click on main)
        click2: XYZ from PickPoint (destination)
        aff_height: drop terminus elevation in feet AFF

    Returns:
        dict of XYZ points and direction vectors
    """
    cl_z = props['centerline_z']
    main_radius = props['radius']
    main_dir = props['direction']

    # Tee point - project click onto main centerline
    click1_on_cl = XYZ(click1.X, click1.Y, cl_z)
    tee_center = project_point_onto_line(
        click1_on_cl, props['start'], props['end']
    )

    # Rise points
    rise_start = XYZ(tee_center.X, tee_center.Y, cl_z)
    rise_end = XYZ(
        tee_center.X,
        tee_center.Y,
        cl_z + main_radius + RISE_HEIGHT
    )

    # Perpendicular direction toward click 2
    perp_dir = get_perpendicular_toward_target(
        main_dir, tee_center, click2
    )

    # Horizontal run distance - project target onto perp axis
    to_target_xy = XYZ(
        click2.X - tee_center.X,
        click2.Y - tee_center.Y,
        0
    )
    horiz_distance = abs(to_target_xy.DotProduct(perp_dir))

    # Horizontal run end (directly above drop location)
    horiz_end = XYZ(
        tee_center.X + perp_dir.X * horiz_distance,
        tee_center.Y + perp_dir.Y * horiz_distance,
        rise_end.Z
    )

    # Drop points
    drop_start = XYZ(horiz_end.X, horiz_end.Y, horiz_end.Z)
    drop_end = XYZ(horiz_end.X, horiz_end.Y, aff_height)

    # Stub-in points - 6in horizontal back toward main
    stub_dir = XYZ(-perp_dir.X, -perp_dir.Y, 0)
    stub_start = XYZ(drop_end.X, drop_end.Y, drop_end.Z)
    stub_end = XYZ(
        drop_end.X + stub_dir.X * STUB_LENGTH,
        drop_end.Y + stub_dir.Y * STUB_LENGTH,
        drop_end.Z
    )

    # --- Validation ---
    if aff_height >= rise_end.Z:
        raise ValueError(
            "Drop terminus ({:.0f} in AFF) is at or above the horizontal "
            "run ({:.1f} ft). Main pipe is too low or AFF is too high."
            .format(aff_height * 12.0, rise_end.Z)
        )

    if horiz_distance < 0.083:  # less than 1 inch
        raise ValueError(
            "Destination point is too close to the main pipe. "
            "Pick a point further away."
        )

    if (rise_end.Z - rise_start.Z) < 0.01:
        raise ValueError("Rise pipe has zero length. Check main pipe geometry.")

    if (drop_start.Z - drop_end.Z) < 0.01:
        raise ValueError("Drop pipe has zero length. Check AFF height.")

    return {
        'tee_center': tee_center,
        'rise_start': rise_start,
        'rise_end': rise_end,
        'horiz_start': rise_end,
        'horiz_end': horiz_end,
        'drop_start': drop_start,
        'drop_end': drop_end,
        'stub_start': stub_start,
        'stub_end': stub_end,
        'perp_direction': perp_dir,
        'stub_direction': stub_dir,
        'horiz_distance': horiz_distance,
    }


# ============================================================================
# PIPE AND FITTING CREATION
# ============================================================================
def create_pipe_segment(sys_id, type_id, lvl_id,
                        start_pt, end_pt, diameter_ft):
    """Create a pipe and set its diameter. Uses copied main properties."""
    new_pipe = Pipe.Create(
        doc, sys_id, type_id, lvl_id, start_pt, end_pt
    )

    dia_param = new_pipe.get_Parameter(
        BuiltInParameter.RBS_PIPE_DIAMETER_PARAM
    )
    if dia_param:
        dia_param.Set(diameter_ft)

    return new_pipe


def split_main_and_place_tee(main_pipe, tee_center, rise_pipe):
    """Split main at tee point, place tee fitting with 3 connectors."""
    new_segment_id = MechanicalUtils.BreakCurve(
        doc, main_pipe.Id, tee_center
    )
    new_segment = doc.GetElement(new_segment_id)

    if new_segment is None:
        raise ValueError(
            "BreakCurve failed at tee point. "
            "Click may be too close to an existing fitting or pipe end."
        )

    doc.Regenerate()

    conn_main_a = get_open_connector_closest_to(main_pipe, tee_center)
    conn_main_b = get_open_connector_closest_to(new_segment, tee_center)
    conn_branch = get_open_connector_closest_to(rise_pipe, tee_center)

    if conn_main_a is None:
        raise ValueError(
            "No open connector on main segment A at split point."
        )
    if conn_main_b is None:
        raise ValueError(
            "No open connector on main segment B at split point."
        )
    if conn_branch is None:
        raise ValueError(
            "No open connector on rise pipe at tee point."
        )

    tee = doc.Create.NewTeeFitting(conn_main_a, conn_main_b, conn_branch)
    return tee


def place_elbow(pipe_a, point_a, pipe_b, point_b):
    """Place elbow fitting between two pipe ends."""
    conn_a = get_open_connector_closest_to(pipe_a, point_a)
    conn_b = get_open_connector_closest_to(pipe_b, point_b)

    if conn_a is None:
        raise ValueError(
            "No open connector on pipe near {}".format(point_a)
        )
    if conn_b is None:
        raise ValueError(
            "No open connector on pipe near {}".format(point_b)
        )

    elbow = doc.Create.NewElbowFitting(conn_a, conn_b)
    return elbow


# ============================================================================
# MAIN BUILD FUNCTION
# ============================================================================
def build_takeoff(main_pipe, click1, click2, branch_dia_ft, aff_height):
    """Build the complete takeoff: tee + 4 pipes + 3 elbows.

    Properties are copied fresh from the clicked main pipe each time.
    Routing preferences ride along with the copied pipe type ID.

    Args:
        main_pipe: existing Pipe element (the main)
        click1: XYZ click location on main
        click2: XYZ destination point
        branch_dia_ft: branch pipe diameter in feet
        aff_height: drop terminus elevation in feet AFF
    """
    # ---- Step 1: Copy properties from main ----
    props = copy_main_properties(main_pipe)

    # ---- Step 2: Pre-check routing preferences ----
    warnings = check_routing_preferences(
        props['pipe_type_id'],
        branch_dia_ft,
        props['pipe_type_name']
    )
    if warnings:
        msg = "Routing preference warnings for '{}':\n\n".format(
            props['pipe_type_name']
        )
        for w in warnings:
            msg += "- {}\n".format(w)
        msg += "\nContinue anyway? Fittings may fail to place."
        if not forms.alert(msg, yes=True, no=True):
            return

    # ---- Step 3: Check diagonal ----
    diag_warning = check_diagonal_main(props['direction'])
    if diag_warning:
        if not forms.alert(diag_warning, yes=True, no=True):
            return

    # ---- Step 4: Calculate geometry ----
    geo = calculate_takeoff_geometry(props, click1, click2, aff_height)

    # Shorthand for copied IDs
    sys_id = props['system_type_id']
    type_id = props['pipe_type_id']
    lvl_id = props['level_id']

    # ---- Step 5: Create rise pipe (vertical) ----
    rise_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id,
        geo['rise_start'], geo['rise_end'],
        branch_dia_ft
    )

    # ---- Step 6: Split main and place tee ----
    doc.Regenerate()
    tee = split_main_and_place_tee(
        props['element'], geo['tee_center'], rise_pipe
    )

    # ---- Step 7: Create horizontal run pipe ----
    horiz_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id,
        geo['horiz_start'], geo['horiz_end'],
        branch_dia_ft
    )

    # ---- Step 8: Elbow 1 - rise top to horizontal start ----
    doc.Regenerate()
    elbow1 = place_elbow(
        rise_pipe, geo['rise_end'],
        horiz_pipe, geo['horiz_start']
    )

    # ---- Step 9: Create vertical drop pipe ----
    drop_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id,
        geo['drop_start'], geo['drop_end'],
        branch_dia_ft
    )

    # ---- Step 10: Elbow 2 - horizontal end to drop top ----
    doc.Regenerate()
    elbow2 = place_elbow(
        horiz_pipe, geo['horiz_end'],
        drop_pipe, geo['drop_start']
    )

    # ---- Step 11: Create stub-in pipe (6in toward main) ----
    stub_pipe = create_pipe_segment(
        sys_id, type_id, lvl_id,
        geo['stub_start'], geo['stub_end'],
        branch_dia_ft
    )

    # ---- Step 12: Elbow 3 - drop bottom to stub start ----
    doc.Regenerate()
    elbow3 = place_elbow(
        drop_pipe, geo['drop_end'],
        stub_pipe, geo['stub_start']
    )

    logger.debug(
        "Takeoff complete: tee + 4 pipes + 3 elbows "
        "({} branch on {} main)"
        .format(branch_dia_ft * 12.0, props['pipe_type_name'])
    )


# ============================================================================
# SETTINGS DIALOG
# ============================================================================
def show_settings_dialog(current_size, current_aff):
    """Combined pipe size and AFF picker. Returns (size, aff) or None."""
    try:
        from pyrevit.forms import FlexForm, Label, ComboBox, Separator, Button

        components = [
            Label("Branch Pipe Size:"),
            ComboBox("pipe_size", list(PIPE_SIZES.keys()),
                     default=current_size),
            Separator(),
            Label("Drop Height (AFF):"),
            ComboBox("aff_height", list(AFF_PRESETS.keys()),
                     default=current_aff),
            Separator(),
            Button("Start")
        ]

        form = FlexForm("Pipe Takeoff Settings", components)
        if form.show():
            return (
                form.values.get("pipe_size", current_size),
                form.values.get("aff_height", current_aff)
            )
        return None

    except Exception:
        # Fallback for older pyRevit versions
        size_pick = forms.SelectFromList.show(
            list(PIPE_SIZES.keys()),
            title="Branch Pipe Size",
            default=current_size,
            button_name="Select Size"
        )
        if size_pick is None:
            return None

        aff_pick = forms.SelectFromList.show(
            list(AFF_PRESETS.keys()),
            title="Drop Height (AFF)",
            default=current_aff,
            button_name="Select Height"
        )
        if aff_pick is None:
            return None

        return (size_pick, aff_pick)


# ============================================================================
# VIEW VALIDATION
# ============================================================================
def validate_view():
    """Verify active view is a floor plan or engineering plan."""
    view = doc.ActiveView
    valid_types = [DB.ViewType.FloorPlan, DB.ViewType.EngineeringPlan]

    if view.ViewType not in valid_types:
        forms.alert(
            "Pipe Takeoffs requires a Floor Plan or Engineering Plan view.\n\n"
            "Current view type: {}\n\n"
            "Switch to a plan view and try again.".format(view.ViewType),
            title="Wrong View Type"
        )
        return False
    return True


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    """Toggle on/off and run the pipe takeoff click loop."""

    # Check toggle state
    is_active = script.get_envvar(ENVVAR_ACTIVE)

    if is_active:
        # Already ON - turn OFF
        script.set_envvar(ENVVAR_ACTIVE, False)
        script.toggle_icon(False)
        return

    # Turning ON - validate view first
    if not validate_view():
        return

    # Retrieve persisted settings or defaults
    saved_size = script.get_envvar(ENVVAR_SIZE)
    saved_aff = script.get_envvar(ENVVAR_AFF)

    if not saved_size or saved_size not in PIPE_SIZES:
        saved_size = '1/2"'
    if not saved_aff or saved_aff not in AFF_PRESETS:
        saved_aff = '36"'

    # Show settings dialog
    result = show_settings_dialog(saved_size, saved_aff)
    if result is None:
        return  # User cancelled

    selected_size, selected_aff = result

    # Convert to feet
    branch_dia_ft = PIPE_SIZES[selected_size] / 12.0
    aff_height_ft = AFF_PRESETS[selected_aff] / 12.0

    # Persist settings for next activation
    script.set_envvar(ENVVAR_SIZE, selected_size)
    script.set_envvar(ENVVAR_AFF, selected_aff)

    # Toggle ON
    script.set_envvar(ENVVAR_ACTIVE, True)
    script.toggle_icon(True)

    pipe_filter = WaterPipeFilter()

    try:
        while True:
            try:
                # ---- Click 1: Pick main pipe ----
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    pipe_filter,
                    "Pick a CW/HW/HWC main pipe (ESC to exit)"
                )
                main_pipe = doc.GetElement(ref.ElementId)
                click1 = ref.GlobalPoint

                if main_pipe is None:
                    forms.alert("Could not read selected pipe. Try again.")
                    continue

                # ---- Click 2: Pick destination ----
                click2 = uidoc.Selection.PickPoint(
                    "Pick destination point for branch takeoff (ESC to cancel)"
                )

                # ---- Build takeoff in single transaction ----
                with revit.Transaction("Pipe Takeoff"):
                    build_takeoff(
                        main_pipe, click1, click2,
                        branch_dia_ft, aff_height_ft
                    )

            except OperationCanceledException:
                # ESC pressed - exit loop
                break

            except InvalidOperationException as ex:
                forms.alert(
                    "Operation interrupted:\n{}\n\n"
                    "Tool will deactivate.".format(str(ex)),
                    title="Pipe Takeoff Error"
                )
                break

            except ValueError as ex:
                # Geometry validation - show and continue
                forms.alert(str(ex), title="Pipe Takeoff - Invalid Geometry")
                continue

            except Exception as ex:
                # Unexpected - show and continue
                forms.alert(
                    "Unexpected error:\n{}\n\n"
                    "Takeoff rolled back. You can try again."
                    .format(str(ex)),
                    title="Pipe Takeoff Error"
                )
                logger.error("Pipe Takeoff error: {}".format(str(ex)))
                continue

    finally:
        # Always reset state on exit
        script.set_envvar(ENVVAR_ACTIVE, False)
        script.toggle_icon(False)


# ============================================================================
# RUN
# ============================================================================
if __name__ == "__main__" or True:
    main()