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

# MEPSystem is the base class of PipingSystem / MechanicalSystem. Connector.AllRefs
# returns these SYSTEM elements alongside the physical neighbour (verified live in
# Revit 2024: a water heater's inlet reported "3608712:Pipes | 3608700:Piping
# Systems"). They are not physical neighbours and must never be traversed into.
try:
    from Autodesk.Revit.DB import MEPSystem
    MEP_SYSTEM_AVAILABLE = True
except ImportError:
    MEP_SYSTEM_AVAILABLE = False

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


def get_type_parameter_value(element, param_name):
    """Read a TYPE parameter from a family instance.

    LookupParameter on an instance only reaches instance parameters, so Type
    parameters (for example the water fixture family's 'Has CW' / 'Has HW')
    have to be read off the instance's Symbol. Falls back to the instance
    itself, because some families expose the same name in both places.

    Returns the value in a Python-native type, or None if not found.
    """
    fn = "get_type_parameter_value"

    if element is None:
        _log_entry("ERROR", fn, None,
                   "Element is None. Cannot read type parameter '{}'.".format(param_name))
        return None

    symbol = None
    try:
        symbol = element.Symbol
    except Exception:
        try:
            type_id = element.GetTypeId()
            if type_id is not None and eid_int(type_id) > 0:
                symbol = element.Document.GetElement(type_id)
        except Exception:
            symbol = None

    if symbol is not None:
        value = get_parameter_value(symbol, param_name)
        if value is not None:
            return value

    return get_parameter_value(element, param_name)


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

def _is_mep_system(element):
    """True if the element is an MEPSystem (a Piping/Mechanical System object).

    Connector.AllRefs returns the owning system element alongside the physical
    neighbour, so every caller that walks AllRefs must skip these. Falls back to
    a class-name test if the MEPSystem type could not be imported.
    """
    if element is None:
        return False
    if MEP_SYSTEM_AVAILABLE:
        try:
            return isinstance(element, MEPSystem)
        except Exception:
            pass
    try:
        return "System" in element.GetType().Name
    except Exception:
        return False


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
            "origin_xyz": None,
            # Added for the water tools. Gas callers ignore these; existing keys
            # above are unchanged so pipe_graph.py behaves exactly as before.
            "connected_system_element_id": None,   # MEPSystem ref, never traversed
            "system_type": None,                   # e.g. "DomesticColdWater"
            "diameter_in": None                    # connector size, inches
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
        # AllRefs mixes the physical neighbour with the MEPSystem element that
        # owns the run, and the order is not guaranteed. Prefer a physical
        # neighbour; never hand an MEPSystem back as connected_element_id,
        # because traversal would then try to walk into the system object.
        try:
            entry["is_connected"] = connector.IsConnected
            if connector.IsConnected:
                refs = connector.AllRefs
                for ref in refs:
                    try:
                        owner = ref.Owner
                        owner_eid = eid_int(owner.Id)
                        if owner_eid == eid:
                            continue
                        if _is_mep_system(owner):
                            if entry["connected_system_element_id"] is None:
                                entry["connected_system_element_id"] = owner_eid
                            continue
                        entry["connected_element_id"] = owner_eid
                        entry["connected_element_type"] = owner.GetType().Name
                        break
                    except Exception:
                        continue
        except Exception as e:
            _log_entry("WARNING", fn, eid,
                       "Connector {}: could not read connection refs: {}".format(i, str(e)))

        # --- Piping system type (water traversal needs this; gas ignores it) ---
        # Verified live in Revit 2024: returns DomesticColdWater / DomesticHotWater /
        # Sanitary / Vent / OtherPipe. NOTE a hot-water RECIRC connector also reports
        # DomesticHotWater, so this enum alone cannot separate supply from return -
        # callers must compare the owning PipingSystemType element for that.
        try:
            entry["system_type"] = str(connector.PipeSystemType)
        except Exception as e:
            _log_entry("INFO", fn, eid,
                       "Connector {}: no PipeSystemType ({}).".format(i, str(e)))

        # --- Connector diameter in inches ---
        try:
            entry["diameter_in"] = connector.Radius * 2.0 * shared_params.INCHES_PER_FOOT
        except Exception as e:
            _log_entry("INFO", fn, eid,
                       "Connector {}: no Radius ({}).".format(i, str(e)))

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
# PIPE DIAMETER WRITE-BACK
# Shared by Size Gas and Size Water. Lives here, not in a pushbutton, so both
# tools inherit the same fixes - a local re-definition of a shared helper is
# exactly what broke the takeoff tools on Revit 2026.
# =============================================================================

# Which API approach worked, remembered after the first success so the other
# approaches are not retried for every pipe in the run.
_confirmed_diameter_approach = [None]


def reset_pipe_diameter_approach():
    """Forget the cached approach. Call once at the start of a sizing run."""
    _confirmed_diameter_approach[0] = None


def _apply_diameter_approach(approach_name, pipe, nominal_feet):
    """Apply one specific write approach. Returns (success, approach_name)."""
    try:
        if approach_name == "RBS_PIPE_NOMINAL_DIAMETER":
            param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_NOMINAL_DIAMETER)
        elif approach_name == "RBS_PIPE_DIAMETER_PARAM":
            param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        elif approach_name == "LookupParameter":
            param = pipe.LookupParameter("Diameter")
        else:
            return False, approach_name

        if param is None:
            return False, approach_name
        if param.IsReadOnly:
            return False, approach_name

        param.Set(nominal_feet)
        return True, approach_name

    except Exception:
        return False, approach_name


def set_pipe_diameter(pipe, nominal_inches):
    """Set a pipe's nominal diameter, trying three API approaches in order.

    Revit exposes pipe diameter differently depending on version and pipe type,
    so the three known-working parameters are tried in order and the first one
    that succeeds is cached for the rest of the run.

    Args:
        pipe:           Revit Pipe element.
        nominal_inches: float, nominal diameter in inches.

    Returns:
        (success: bool, approach_name: str). approach_name is "FAILED" when
        every approach was rejected.
    """
    fn = "set_pipe_diameter"
    nominal_feet = nominal_inches / shared_params.INCHES_PER_FOOT

    try:
        pipe_id = eid_int(pipe.Id)
    except Exception:
        pipe_id = None

    # Try the remembered approach first. If it fails, fall through and retry
    # all of them rather than giving up: this module stays loaded between
    # pyRevit runs, so the cached approach can be stale (different Revit
    # version, different pipe type) and must never become a dead end.
    if _confirmed_diameter_approach[0] is not None:
        ok, name = _apply_diameter_approach(
            _confirmed_diameter_approach[0], pipe, nominal_feet)
        if ok:
            return True, name
        _log_entry("WARNING", fn, pipe_id,
                   "Cached approach '{}' failed, retrying all approaches.".format(name))
        _confirmed_diameter_approach[0] = None

    for approach in ("RBS_PIPE_NOMINAL_DIAMETER",
                     "RBS_PIPE_DIAMETER_PARAM",
                     "LookupParameter"):
        ok, name = _apply_diameter_approach(approach, pipe, nominal_feet)
        if ok:
            _confirmed_diameter_approach[0] = name
            _log_entry("INFO", fn, pipe_id,
                       "API approach confirmed: {}.".format(name))
            return True, name

    _log_entry("ERROR", fn, pipe_id,
               "All three diameter API approaches failed.")
    return False, "FAILED"


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
