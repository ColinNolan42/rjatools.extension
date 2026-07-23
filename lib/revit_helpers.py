# revit_helpers.py
# Safe Revit API wrapper functions with null handling and logging.
# ALL Revit API calls in this codebase shall go through these functions.
# No module shall call the Revit API directly  -  use these wrappers instead.
#
# Every function:
#   - Handles null/missing values without crashing
#   - Logs what it found or did not find
#   - Returns None on failure rather than raising unhandled exceptions
#
# IronPython 2.7  -  Revit API via Autodesk.Revit.DB

import sys
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    BuiltInCategory,
)

# Connector flow direction  -  MEP specific
# Note: FlowDirectionType lives in Autodesk.Revit.DB (not a sub-namespace in all versions).
# If this import fails in your Revit version, please report the exact error so we can
# adjust the namespace. Verified against Revit 2024 API.
try:
    from Autodesk.Revit.DB import FlowDirectionType
    FLOW_DIRECTION_AVAILABLE = True
except ImportError:
    FLOW_DIRECTION_AVAILABLE = False

import shared_params

from pyrevit import HOST_APP

# Checked once at module load. Drives every version-dependent API choice
# below instead of relying purely on try/except at each call site, so any
# new version branch (e.g. a future API break in Revit 2027+) has one place
# to update. HOST_APP.version is a string like "2026"; if it can't be
# parsed, REVIT_VERSION stays None and callers fall back to attribute
# detection instead of guessing a version number.
try:
    REVIT_VERSION = int(HOST_APP.version)
except Exception:
    REVIT_VERSION = None


def eid_int(element_id):
    """Version-safe ElementId -> int/long.

    Revit 2024+ replaced ElementId.IntegerValue (int) with ElementId.Value
    (long); Revit 2025/2026 removed IntegerValue entirely. Revit 2022/2023
    only have IntegerValue. Always call this instead of .IntegerValue or
    .Value directly so the same code works across every Revit version this
    firm uses (2022 through 2026+).
    """
    if REVIT_VERSION is not None:
        return element_id.Value if REVIT_VERSION >= 2024 else element_id.IntegerValue
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def param_is_yes_no(param):
    """Version-safe YesNo/Boolean parameter detection.

    Revit 2022+ deprecated Definition.ParameterType in favor of
    Definition.GetDataType() -> ForgeTypeId, compared against
    SpecTypeId.Boolean.YesNo. Revit 2025/2026 removed ParameterType
    entirely (same deprecation timeline as ElementId.IntegerValue, and
    confirmed the same way: it fails to even compile against the live
    2026 API). REVIT_VERSION (detected once at module load, see above)
    picks the branch explicitly instead of guessing from a bare
    try/except, matching the eid_int() pattern above. The old
    Revit-2022-era behavior (this firm's oldest supported version) is
    preserved unchanged in the else branch, not replaced by a shim that
    could subtly behave differently on the version that already worked.
    """
    from Autodesk.Revit.DB import SpecTypeId

    if REVIT_VERSION is not None:
        if REVIT_VERSION >= 2022:
            try:
                return param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo
            except Exception:
                return False
        param_type = str(param.Definition.ParameterType)
        return "YesNo" in param_type or param_type == "Invalid"

    # REVIT_VERSION undetermined - try modern API first, then the old enum.
    try:
        return param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo
    except Exception:
        pass
    try:
        param_type = str(param.Definition.ParameterType)
        return "YesNo" in param_type or param_type == "Invalid"
    except Exception:
        return False


# =============================================================================
# MODULE-LEVEL LOG
# All functions append to this list. The diagnostic report reads it.
# Call get_log() to retrieve all entries. Call clear_log() before each run.
# =============================================================================

_log = []

def get_log():
    """Return all log entries accumulated during this run."""
    return list(_log)

def clear_log():
    """Clear the log. Call this at the start of each Diagnose run."""
    global _log
    _log = []

def _log_entry(level, function_name, element_id, message):
    """Append a structured log entry.
    
    Args:
        level: "INFO", "WARNING", or "ERROR"
        function_name: Name of the calling function
        element_id: Revit element ID (int) or None
        message: Description of what was found or not found
    """
    _log.append({
        "level": level,
        "function": function_name,
        "element_id": element_id,
        "message": message
    })


# =============================================================================
# PARAMETER READING
# =============================================================================

def get_parameter_value(element, param_name):
    """Read a shared or family parameter value from an element.
    
    Attempts LookupParameter by name. Handles all Revit storage types:
    Double, Integer, String, ElementId, YesNo (stored as Integer 0/1).
    
    Args:
        element: A Revit Element object.
        param_name: Parameter name string  -  use constants from shared_params.py.
        
    Returns:
        The parameter value in a Python-native type, or None if not found.
        - Double parameters -> float
        - Integer parameters -> int
        - YesNo parameters -> bool (True/False)
        - String parameters -> str
        - ElementId parameters -> int (the element ID integer value)
    """
    fn = "get_parameter_value"
    
    if element is None:
        _log_entry("ERROR", fn, None, 
                   "Element is None. Cannot read parameter '{}'.".format(param_name))
        return None
    
    try:
        eid = eid_int(element.Id)
    except Exception:
        eid = None

    try:
        param = element.LookupParameter(param_name)
    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "LookupParameter('{}') threw exception: {}".format(param_name, str(e)))
        return None

    if param is None:
        _log_entry("WARNING", fn, eid,
                   "Parameter '{}' NOT FOUND on element.".format(param_name))
        return None

    if not param.HasValue:
        _log_entry("WARNING", fn, eid,
                   "Parameter '{}' found but has no value.".format(param_name))
        return None

    try:
        storage_type = str(param.StorageType)

        if storage_type == "Double":
            value = param.AsDouble()
            _log_entry("INFO", fn, eid,
                       "Parameter '{}' = {} (Double/float).".format(param_name, value))
            return value

        elif storage_type == "Integer":
            value = param.AsInteger()
            # Yes/No parameters are stored as Integer (1 = Yes, 0 = No)
            if param_is_yes_no(param):
                bool_value = (value == 1)
                _log_entry("INFO", fn, eid,
                           "Parameter '{}' = {} (YesNo -> bool {}).".format(
                               param_name, value, bool_value))
                return bool_value
            _log_entry("INFO", fn, eid,
                       "Parameter '{}' = {} (Integer).".format(param_name, value))
            return value

        elif storage_type == "String":
            value = param.AsString()
            _log_entry("INFO", fn, eid,
                       "Parameter '{}' = '{}' (String).".format(param_name, value))
            return value

        elif storage_type == "ElementId":
            value = eid_int(param.AsElementId())
            _log_entry("INFO", fn, eid,
                       "Parameter '{}' = {} (ElementId).".format(param_name, value))
            return value

        else:
            _log_entry("WARNING", fn, eid,
                       "Parameter '{}' has unhandled StorageType: {}.".format(
                           param_name, storage_type))
            return None

    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "Failed to read value of parameter '{}': {}".format(param_name, str(e)))
        return None


# =============================================================================
# PIPE GEOMETRY
# =============================================================================

def get_pipe_length_feet(pipe):
    """Read pipe length from Location.Curve.Length.
    
    Revit internal units are decimal feet. No conversion needed.
    
    Args:
        pipe: A Revit Pipe element.
        
    Returns:
        Length in feet as float, or None on failure.
    """
    fn = "get_pipe_length_feet"

    if pipe is None:
        _log_entry("ERROR", fn, None, "Pipe element is None.")
        return None

    eid = eid_int(pipe.Id)

    try:
        length = pipe.Location.Curve.Length
        _log_entry("INFO", fn, eid,
                   "Pipe length = {:.4f} ft (from Location.Curve.Length).".format(length))
        return length
    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "Failed to read pipe length: {}".format(str(e)))
        return None


def get_pipe_diameter_inches(pipe):
    """Read pipe nominal diameter and convert to inches.
    
    Reads RBS_PIPE_DIAMETER_PARAM which returns decimal feet in Revit internal units.
    Multiplies by 12 to convert to inches.
    
    Args:
        pipe: A Revit Pipe element.
        
    Returns:
        Diameter in inches as float, or None on failure.
    """
    fn = "get_pipe_diameter_inches"

    if pipe is None:
        _log_entry("ERROR", fn, None, "Pipe element is None.")
        return None

    eid = eid_int(pipe.Id)

    try:
        param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if param is None:
            _log_entry("ERROR", fn, eid,
                       "RBS_PIPE_DIAMETER_PARAM not found on pipe.")
            return None

        diameter_feet = param.AsDouble()
        diameter_inches = diameter_feet * shared_params.INCHES_PER_FOOT
        _log_entry("INFO", fn, eid,
                   "Pipe diameter = {:.4f} ft = {:.4f} in.".format(
                       diameter_feet, diameter_inches))
        return diameter_inches

    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "Failed to read pipe diameter: {}".format(str(e)))
        return None


# =============================================================================
# CONNECTOR INSPECTION
# =============================================================================

def _get_connector_manager(element):
    """Return the ConnectorManager for any MEP element.

    Tries two access paths:
      1. element.ConnectorManager         - pipes, fittings, most elements
      2. element.MEPModel.ConnectorManager - mechanical/plumbing equipment families

    Returns ConnectorManager or None.
    """
    if element is None:
        return None

    eid = eid_int(element.Id)

    # Path 1 - direct access (pipes, fittings, most elements)
    try:
        cm = element.ConnectorManager
        if cm is not None:
            _log_entry("INFO", "_get_connector_manager", eid,
                       "ConnectorManager found via element.ConnectorManager.")
            return cm
    except Exception:
        pass

    # Path 2 - via MEPModel (mechanical equipment, plumbing fixture families)
    try:
        cm = element.MEPModel.ConnectorManager
        if cm is not None:
            _log_entry("INFO", "_get_connector_manager", eid,
                       "ConnectorManager found via element.MEPModel.ConnectorManager.")
            return cm
    except Exception:
        pass

    _log_entry("WARNING", "_get_connector_manager", eid,
               "No ConnectorManager found via any access path.")
    return None


def get_connectors(element):
    """Return all connectors on an element as a list of dicts.
    
    Each dict contains:
        - connector_index: int
        - direction: str ("In", "Out", "Bidirectional", or "Unknown")
        - is_connected: bool
        - connected_element_id: int or None
        - connected_element_type: str or None (e.g. "Pipe", "FamilyInstance")
        - origin_xyz: [x, y, z] in decimal feet
    
    Args:
        element: Any Revit element with a ConnectorManager.
        
    Returns:
        List of connector dicts. Empty list if element has no connectors.
    """
    fn = "get_connectors"

    if element is None:
        _log_entry("ERROR", fn, None, "Element is None.")
        return []

    eid = eid_int(element.Id)

    connector_manager = _get_connector_manager(element)

    if connector_manager is None:
        _log_entry("WARNING", fn, eid,
                   "No ConnectorManager found on element via any access path.")
        return []

    results = []

    try:
        connectors = connector_manager.Connectors
    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "Failed to access Connectors collection: {}".format(str(e)))
        return []

    for i, connector in enumerate(connectors):
        entry = {
            "connector_index": i,
            "direction": "Unknown",
            "is_connected": False,
            "connected_element_id": None,
            "connected_element_type": None,
            "origin_xyz": None
        }

        # --- Flow direction ---
        try:
            if FLOW_DIRECTION_AVAILABLE:
                dir_val = connector.Direction
                if dir_val == FlowDirectionType.Out:
                    entry["direction"] = "Out"
                elif dir_val == FlowDirectionType.In:
                    entry["direction"] = "In"
                elif dir_val == FlowDirectionType.Bidirectional:
                    entry["direction"] = "Bidirectional"
                else:
                    entry["direction"] = str(dir_val)
            else:
                # FlowDirectionType not available  -  read as string
                entry["direction"] = str(connector.Direction)
        except Exception as e:
            _log_entry("WARNING", fn, eid,
                       "Connector {}: could not read Direction: {}".format(i, str(e)))

        # --- Connection status and connected element ---
        try:
            entry["is_connected"] = connector.IsConnected
            if connector.IsConnected:
                refs = connector.AllRefs
                for ref in refs:
                    try:
                        owner = ref.Owner
                        owner_eid = eid_int(owner.Id)
                        if owner_eid != eid:
                            entry["connected_element_id"] = owner_eid
                            entry["connected_element_type"] = owner.GetType().Name
                            break
                    except Exception:
                        continue
        except Exception as e:
            _log_entry("WARNING", fn, eid,
                       "Connector {}: could not read connection refs: {}".format(i, str(e)))

        # --- Connector origin (location) ---
        try:
            origin = connector.Origin
            entry["origin_xyz"] = [
                round(origin.X, 4),
                round(origin.Y, 4),
                round(origin.Z, 4)
            ]
        except Exception as e:
            _log_entry("WARNING", fn, eid,
                       "Connector {}: could not read Origin: {}".format(i, str(e)))

        results.append(entry)
        _log_entry("INFO", fn, eid,
                   "Connector {}: direction={}, is_connected={}, connected_to={} ({})".format(
                       i,
                       entry["direction"],
                       entry["is_connected"],
                       entry["connected_element_id"],
                       entry["connected_element_type"]
                   ))

    return results


# =============================================================================
# ELEMENT LOCATION
# =============================================================================

def get_element_location(element):
    """Return the XYZ location of an element as a [x, y, z] list in decimal feet.
    
    For point-located elements (families, fittings): reads Location.Point.
    For curve-located elements (pipes): reads Location.Curve.GetEndPoint(0).
    
    Args:
        element: Any Revit element with a Location.
        
    Returns:
        [x, y, z] list of floats in decimal feet, or None on failure.
    """
    fn = "get_element_location"

    if element is None:
        _log_entry("ERROR", fn, None, "Element is None.")
        return None

    eid = eid_int(element.Id)

    try:
        location = element.Location

        # Try point location first (families, fittings, equipment)
        try:
            point = location.Point
            xyz = [round(point.X, 4), round(point.Y, 4), round(point.Z, 4)]
            _log_entry("INFO", fn, eid,
                       "Location (Point) = [{}, {}, {}] ft.".format(*xyz))
            return xyz
        except Exception:
            pass

        # Try curve location (pipes)
        try:
            pt = location.Curve.GetEndPoint(0)
            xyz = [round(pt.X, 4), round(pt.Y, 4), round(pt.Z, 4)]
            _log_entry("INFO", fn, eid,
                       "Location (Curve start) = [{}, {}, {}] ft.".format(*xyz))
            return xyz
        except Exception:
            pass

        _log_entry("WARNING", fn, eid,
                   "Could not read location  -  not a point or curve element.")
        return None

    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "Failed to read element location: {}".format(str(e)))
        return None


# =============================================================================
# METER VALIDATION
# =============================================================================

def validate_selected_element(element):
    """Validate that the selected element is suitable as a gas meter start point.
    
    Checks:
        1. Element is not None
        2. Element has a ConnectorManager (has MEP connectors)
        3. Element has at least one connector with Direction = Out (building side)
        4. The Out connector is connected to something (not floating)
    
    Args:
        element: The user-selected Revit element.
        
    Returns:
        Dict with keys:
            - is_valid: bool
            - reason: str describing pass or first failure found
            - connector_summary: list of connector direction strings
    """
    fn = "validate_selected_element"

    if element is None:
        _log_entry("ERROR", fn, None, "Selected element is None.")
        return {
            "is_valid": False,
            "reason": "No element selected. Please select the gas meter.",
            "connector_summary": []
        }

    eid = eid_int(element.Id)

    # Check ConnectorManager exists - tries element.ConnectorManager
    # and element.MEPModel.ConnectorManager for equipment families
    cm = _get_connector_manager(element)
    if cm is None:
        _log_entry("ERROR", fn, eid,
                   "No ConnectorManager found via any access path.")
        return {
            "is_valid": False,
            "reason": "Selected element has no MEP connectors. Please select the gas meter.",
            "connector_summary": []
        }

    # Read all connectors
    connectors = get_connectors(element)
    connector_summary = [c["direction"] for c in connectors]

    if not connectors:
        _log_entry("ERROR", fn, eid, "No connectors found on selected element.")
        return {
            "is_valid": False,
            "reason": "Selected element has no connectors.",
            "connector_summary": connector_summary
        }

    # Check for at least one Out connector
    out_connectors = [c for c in connectors if c["direction"] == "Out"]
    if not out_connectors:
        _log_entry("ERROR", fn, eid,
                   "No Out connector found. Directions found: {}".format(connector_summary))
        return {
            "is_valid": False,
            "reason": (
                "Selected element has no Out connector. "
                "The meter family shall have its building-side connector set to Flow = Out. "
                "Connector directions found: {}".format(connector_summary)
            ),
            "connector_summary": connector_summary
        }

    # Check that the Out connector is actually connected to piping
    out_connector = out_connectors[0]
    if not out_connector["is_connected"]:
        _log_entry("ERROR", fn, eid,
                   "Out connector is not connected to any piping.")
        return {
            "is_valid": False,
            "reason": (
                "The meter's Out connector is not connected to any piping. "
                "Please connect the meter to the gas distribution piping."
            ),
            "connector_summary": connector_summary
        }

    _log_entry("INFO", fn, eid,
               "Validation PASSED. Out connector connected to element {}.".format(
                   out_connector["connected_element_id"]))

    return {
        "is_valid": True,
        "reason": "PASS",
        "connector_summary": connector_summary
    }
