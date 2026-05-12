# shared_params.py
# Central repository for all shared parameter name strings and hardcoded constants.
# ALL modules shall import from here. No parameter name strings shall be hardcoded
# anywhere else in the codebase.

# =============================================================================
# GAS FIXTURE SHARED PARAMETERS
# These must match exactly the shared parameter names defined in MEP_SharedParams.txt
# and loaded into the fixture/cap families.
# =============================================================================

PARAM_GAS_LOAD_MBH      = "GAS_LOAD_MBH"       # Number. Gas load in MBH. Instance parameter.
PARAM_IS_GAS_FIXTURE    = "IS_GAS_FIXTURE"      # Yes/No. Identifies load-bearing terminal nodes.
PARAM_FIXTURE_NAME      = "FIXTURE_NAME"        # Text. Equipment tag for one-line diagram labels.

# =============================================================================
# REVIT BUILT-IN PARAMETER REFERENCES
# Used by revit_helpers.py to read pipe geometry from the Revit model.
# These reference Autodesk.Revit.DB.BuiltInParameter enum values.
# Import BuiltInParameter from Autodesk.Revit.DB before using these.
# =============================================================================

# Pipe nominal diameter — returns value in Revit internal units (decimal feet).
# Must be converted to inches: value * 12
BUILTIN_PIPE_DIAMETER   = "RBS_PIPE_DIAMETER_PARAM"

# Pipe length — accessed via pipe.Location.Curve.Length
# Returns value in Revit internal units (decimal feet).
BUILTIN_PIPE_LENGTH     = "CURVE_ELEM_LENGTH"

# =============================================================================
# HARDCODED ENGINEERING CONSTANTS
# =============================================================================

SPECIFIC_GRAVITY        = 0.60      # Natural gas specific gravity. Matches all IFGC tables.
                                    # Do not expose as user input — hardcoded per project spec.

BTU_PER_CFH             = 1000      # Natural gas: 1 cubic foot = 1000 BTU (approx).
                                    # CFH = BTU_hr / BTU_PER_CFH

MBH_PER_BTU             = 0.001     # Convenience multiplier: BTU/hr * MBH_PER_BTU = MBH

INCHES_PER_FOOT         = 12.0      # Used when converting pipe diameter from feet to inches.

# =============================================================================
# CONNECTOR FLOW DIRECTION
# Used by pipe_graph.py to identify the building-side connector on the meter.
# The meter family sets street-side connector to Flow = In, building-side to Flow = Out.
# =============================================================================

CONNECTOR_FLOW_OUT      = "Out"         # Building side — traversal walks this direction.
CONNECTOR_FLOW_IN       = "In"          # Street side — traversal skips this direction.
CONNECTOR_FLOW_BIDI     = "Bidirectional"   # All non-meter elements expected to be bidirectional.

# =============================================================================
# MEP SYSTEM TYPE NAME
# Must match the system type name exactly as defined in the Revit model.
# If the project uses a different name (e.g. "GAS" or "Natural Gas"),
# update this constant — do not hardcode the string elsewhere.
# =============================================================================

SYSTEM_TYPE_GAS         = "Natural Gas"

# =============================================================================
# DIAGNOSTIC REPORT METADATA
# =============================================================================

REPORT_SCHEMA_VERSION   = "1.0"
TOOL_NAME               = "RevitPipeSizing"
PHASE                   = "Phase 1 - Diagnostic"
