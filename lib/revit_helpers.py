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
        eid = element.Id.IntegerValue
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
            # Detect by checking if the parameter definition type is YesNo
            try:
                param_type = str(param.Definition.ParameterType)
                if "YesNo" in param_type or param_type == "Invalid":
                    bool_value = (value == 1)
                    _log_entry("INFO", fn, eid,
                               "Parameter '{}' = {} (YesNo -> bool {}).".format(
                                   param_name, value, bool_value))
                    return bool_value
            except Exception:
                pass
            _log_entry("INFO", fn, eid,
                       "Parameter '{}' = {} (Integer).".format(param_name, value))
            return value

        elif storage_type == "String":
            value = param.AsString()
            _log_entry("INFO", fn, eid,
                       "Parameter '{}' = '{}' (String).".format(param_name, value))
            return value

        elif storage_type == "ElementId":
            value = param.AsElementId().IntegerValue
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

    eid = pipe.Id.IntegerValue

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

    eid = pipe.Id.IntegerValue

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

    eid = element.Id.IntegerValue

    try:
        connector_manager = element.ConnectorManager
    except Exception as e:
        _log_entry("WARNING", fn, eid,
                   "Element has no ConnectorManager: {}".format(str(e)))
        return []

    if connector_manager is None:
        _log_entry("WARNING", fn, eid, "ConnectorManager is None.")
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
                        if owner.Id.IntegerValue != eid:
                            entry["connected_element_id"] = owner.Id.IntegerValue
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

    eid = element.Id.IntegerValue

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

    eid = element.Id.IntegerValue

    # Check ConnectorManager exists
    try:
        cm = element.ConnectorManager
        if cm is None:
            _log_entry("ERROR", fn, eid, "ConnectorManager is None.")
            return {
                "is_valid": False,
                "reason": "Selected element has no MEP connectors. Please select the gas meter.",
                "connector_summary": []
            }
    except Exception as e:
        _log_entry("ERROR", fn, eid,
                   "No ConnectorManager on selected element: {}".format(str(e)))
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
