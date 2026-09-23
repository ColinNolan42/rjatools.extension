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

# Pipe nominal diameter  -  returns value in Revit internal units (decimal feet).
# Must be converted to inches: value * 12
BUILTIN_PIPE_DIAMETER   = "RBS_PIPE_DIAMETER_PARAM"

# Pipe length  -  accessed via pipe.Location.Curve.Length
# Returns value in Revit internal units (decimal feet).
BUILTIN_PIPE_LENGTH     = "CURVE_ELEM_LENGTH"

# =============================================================================
# HARDCODED ENGINEERING CONSTANTS
# =============================================================================

SPECIFIC_GRAVITY        = 0.60      # Natural gas specific gravity. Matches all IFGC tables.
                                    # Do not expose as user input  -  hardcoded per project spec.

INCHES_PER_FOOT         = 12.0      # Used when converting pipe diameter from feet to inches.

# Real Low Pressure (K-constant) / High Pressure (Weymouth/Cox) capacity
# formulas were ported from RJA's actual calc templates ("Template Low
# Pressure Gas Size.xls" / "Template - High Pressure Gas Size.xlsx") and
# verified working (see sizing_engine.low_pressure_capacity_cfh() /
# weymouth_capacity_cfh() / cox_capacity_cfh()), but per Colin (2026-09-23)
# those formulas are NOT used for actual pipe capacity/sizing in this tool
# - only the templates' MBH->CFH Heat Content conversion is (see
# mbh_to_cfh() below). Pipe capacity/sizing stays on the discrete IFGC
# Table 402.4 lookup (SPECIFIC_GRAVITY above, fixed at 0.60). The formula
# functions are kept in sizing_engine.py, unused, in case this is
# revisited later - do not wire them back into size_system() without
# Colin explicitly asking again.

# RJA standard (2026-09-23): CFH = BTUH / Heat Content of Gas. Heat Content
# of Gas is project/location-specific (utility-provided), not a fixed
# altitude derate - this is the ONLY adjustment applied when converting a
# fixture's MBH load to the CFH figure compared against IFGC Table 402.4
# capacities. Matches RJA's real "Low Pressure Gas Size.xls" calc template
# (Table sheet: "CF/BTUH (Sea Level)" = 1000, "CF/BTUH (Denver)" = 840).
# Supersedes the old fixed 4%/1,000ft derate, and supersedes "GAS SIZING
# WHITEPAPER.docx"'s "1 CFH = 1 MBH, no adjustment" position - both are now
# outdated per Colin, do not re-cite either as authoritative for this.
SEA_LEVEL_HEAT_CONTENT_BTU_PER_CF = 1000.0

DEFAULT_HEAT_CONTENT_BTU_PER_CF   = 840.0   # Denver, CO - startup dialog default.

# =============================================================================
# CONNECTOR FLOW DIRECTION
# Used by pipe_graph.py to identify the building-side connector on the meter.
# The meter family sets street-side connector to Flow = In, building-side to Flow = Out.
# =============================================================================

CONNECTOR_FLOW_OUT      = "Out"         # Building side  -  traversal walks this direction.
CONNECTOR_FLOW_IN       = "In"          # Street side  -  traversal skips this direction.
CONNECTOR_FLOW_BIDI     = "Bidirectional"   # All non-meter elements expected to be bidirectional.

# =============================================================================
# MEP SYSTEM TYPE NAME
# Must match the system type name exactly as defined in the Revit model.
# If the project uses a different name (e.g. "GAS" or "Natural Gas"),
# update this constant  -  do not hardcode the string elsewhere.
# =============================================================================

SYSTEM_TYPE_GAS         = "Natural Gas"

# =============================================================================
# DIAGNOSTIC REPORT METADATA
# =============================================================================

REPORT_SCHEMA_VERSION   = "1.0"
TOOL_NAME               = "RevitPipeSizing"
PHASE                   = "Phase 1 - Diagnostic"
