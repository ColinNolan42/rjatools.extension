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

# Pressure regulating valves carry NO custom parameters. A PRV is auto-detected
# during traversal by a case-insensitive match of these words in its family
# name. One list, used by pipe_graph, One-Line and Diagnose. It is checked
# BEFORE ISOLATION_VALVE_KEYWORDS, so a "Gas Pressure Reducing Valve" is a PRV,
# not an isolation valve (Colin, 2026-11-20). Bare "reducing" is not a keyword:
# it would catch reducing tees and couplings.
PRV_FAMILY_KEYWORDS = ("prv", "regulator", "regulating", "pressure reducing",
                       "pressure reducer", "reducing valve")

# A PRV only counts as a MID-STREAM step down when the pipe run downstream of
# it, on its branch with the most pipe, is LONGER than this (feet of actual pipe, no
# elbow equivalents). A regulator closer to its equipment than this is that
# equipment's own regulator: it is ignored for sizing.
PRV_MIDSTREAM_MIN_DOWNSTREAM_FT = 10.0

# Isolation valves on the One-Line: a case-insensitive match of these words in
# a fitting's family name, checked only after the PRV words (a "Pressure
# Regulating Valve" is a PRV). Moved here from the One-Line script.
ISOLATION_VALVE_KEYWORDS = ("valve", "ball", "gate", "check", "shutoff")

# The hood's Ansul gas shutoff valve (interlocked with the fire suppression,
# ahead of every appliance under the hood). Checked before the isolation
# words; drawn as a valve with an ANSUL VALVE label (Colin, 2026-11-19).
# Family: RJA - Gas Shutoff Valve (Ansul).
ANSUL_VALVE_KEYWORDS = ("ansul",)


def accessory_kind(family_name):
    """The one gas accessory classifier, by family name (case-insensitive):
    "prv" (PRV_FAMILY_KEYWORDS), "ansul" (ANSUL_VALVE_KEYWORDS), "iso"
    (ISOLATION_VALVE_KEYWORDS) or None. Checked in that order, so a "Gas
    Pressure Reducing Valve" is a PRV and an "Ansul Gas Valve" is the Ansul
    shutoff. Read by pipe_graph (PRV detection), Diagnose (valve counts) and
    oneline_schematic (valve symbols). Lives here, beside its keywords,
    because oneline_schematic is pure Python and cannot import pipe_graph
    (which imports the Revit API)."""
    name = (family_name or "").lower()
    if not name:
        return None
    for kind, words in (("prv", PRV_FAMILY_KEYWORDS),
                        ("ansul", ANSUL_VALVE_KEYWORDS),
                        ("iso", ISOLATION_VALVE_KEYWORDS)):
        if any(kw in name for kw in words):
            return kind
    return None


def prv_keywords_text():
    """The PRV keywords for a message: "'prv', 'regulator', ... or
    'reducing valve'", built from PRV_FAMILY_KEYWORDS."""
    words = ["'{0}'".format(k) for k in PRV_FAMILY_KEYWORDS]
    if len(words) < 2:
        return "".join(words)
    return ", ".join(words[:-1]) + " or " + words[-1]

# Equivalent length added to the developed length for each elbow on a path,
# feet (IFGC Appendix A). Read by pipe_graph, the sizing basis text and the
# One-Line (snapshot elbow_equiv_ft, tee_equiv_ft).
ELBOW_EQUIV_LENGTH_FT = 5.0
# Equivalent length of each tee a path passes, straight through or into its
# branch (Colin, 2026-11-19). Developed length = pipe + 5 ft per elbow + 3 ft
# per tee; the length inside fittings is not added.
TEE_EQUIV_LENGTH_FT = 3.0

# Elbow detection when a fitting's part type cannot be read (the part type,
# MEPModel.PartType or the family's FAMILY_CONTENT_PART_TYPE, is read first,
# revit_helpers.is_elbow). Substrings of the family name, any case, plus whole
# words ("ell" alone would match "shell").
ELBOW_FAMILY_KEYWORDS = ("elbow", "bend")
ELBOW_FAMILY_WORDS = ("ell",)

# =============================================================================
# DOMESTIC WATER FIXTURE SHARED PARAMETERS
# Group "Water Fixture Data" in MEP_SharedParams.txt. Names are UPPERCASE to
# match the gas group. Carried by RJA - Water Fixture (IPC Auto-Sizing).
# FIXTURE_NAME above is reused, it is not duplicated here.
# =============================================================================

PARAM_IS_WATER_FIXTURE   = "IS_WATER_FIXTURE"    # Yes/No. Marks a load-bearing terminal.
PARAM_CW_FIXTURE_UNITS   = "CW_FIXTURE_UNITS"    # Number, instance. Cold WSFU.
PARAM_HW_FIXTURE_UNITS   = "HW_FIXTURE_UNITS"    # Number, instance. Hot WSFU.
PARAM_TOTAL_FIXTURE_UNITS = "TOTAL_FIXTURE_UNITS"  # Number, instance. Total WSFU.
PARAM_HWR_ACTIVE         = "HWR_ACTIVE"          # Yes/No, instance, read-only formula.

# Family-only parameters (NOT shared). Read defensively - a third-party or
# one-off fixture family may not carry them, which is not an error.
PARAM_HAS_CW             = "Has CW"              # Type Yes/No. False => ignore the CW connector.
PARAM_HAS_HW             = "Has HW"              # Type Yes/No. False => ignore the HW connector.
PARAM_IS_PUBLIC_OCCUPANCY = "Is Public Occupancy"  # Instance Yes/No. Swaps the WSFU set.
PARAM_CW_SUPPLY_SIZE     = "CW Supply Size"      # Type Length. Min fixture supply.
PARAM_HW_SUPPLY_SIZE     = "HW Supply Size"      # Type Length. Min fixture supply.

# Connector size overrides on the RJA water fixture family (2026-10-13). The
# stub and connector read `if(Override > 0, Override, Type supply size)`, so 0
# means "follow the Type". The tools write the connected pipe size here when it
# is larger than the Type size. The *_CONNECTION_SIZE formula parameters are
# read only, never write them.
PARAM_CW_SIZE_OVERRIDE     = "CW Size Override"      # Instance Length.
PARAM_HW_IN_SIZE_OVERRIDE  = "HW In Size Override"   # Instance Length.
PARAM_HW_OUT_SIZE_OVERRIDE = "HW Out Size Override"  # Instance Length.

# Equipment / device parameters approved 2026-09-23 but NOT yet on any family.
# Read defensively; absent means "not applicable", never an error.
PARAM_FLUSH_TYPE          = "Flush_Type"          # Type text: Tank, Valve, Flushometer Tank, None.
PARAM_DEVICE_PRESSURE_DROP = "Device_Pressure_Drop_PSI"  # Number. Pressure-loss stage only.

# =============================================================================
# PIPING SYSTEM CLASSIFICATION (Connector.PipeSystemType)
# Verified live in Revit 2024 against a real model. These are Revit enum names,
# NOT project system-type names, so they are safe to compare against in code.
#
# WARNING: a hot water RECIRCULATION connector also reports DomesticHotWater.
# The enum cannot separate supply from return. To tell them apart, compare the
# owning PipingSystemType ELEMENT (project types "Domestic Hot Water" and
# "Domestic Hot Water Recirc" both classify as DomesticHotWater). Never
# hardcode a project system-type name - present the names found in the model
# and let the user map them.
# =============================================================================

SYSTEM_DOMESTIC_COLD_WATER = "DomesticColdWater"
SYSTEM_DOMESTIC_HOT_WATER  = "DomesticHotWater"
SYSTEM_OTHER_PIPE          = "OtherPipe"   # what a circulation pump's own connectors report

# =============================================================================
# HOT WATER RECIRCULATION (Colin, 2026-10-12)
# From the firm calc "Template - Circ Size Calc.xlsx", tab HOT WATER RECIRC
# HEAD LOSS: BTU/h = effective length / 100 x heat loss, gpm = BTU/h /
# (500 x temperature drop). Heat loss and temperature drop are editable in the
# Size Water dialog; the temperature drop may be lowered but never raised.
# =============================================================================

DEFAULT_RECIRC_HEAT_LOSS_BTU_PER_100FT = 5000.0
DEFAULT_RECIRC_DELTA_T_F               = 20.0
MAX_RECIRC_DELTA_T_F                   = 20.0
MIN_CIRCUIT_SETTER_GPM                 = 0.5    # a circuit setter is not set lower

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
DEFAULT_PROPANE_HEAT_CONTENT_BTU_PER_CF = 2500.0  # propane CFH Table 402.4(37) only (Colin, 2026-11-20).

# =============================================================================
# CONNECTOR FLOW DIRECTION
# Used by pipe_graph.py to identify the building-side connector on the meter.
# The meter family sets street-side connector to Flow = In, building-side to Flow = Out.
# =============================================================================

CONNECTOR_FLOW_OUT      = "Out"         # Building side  -  traversal walks this direction.
CONNECTOR_FLOW_IN       = "In"          # Street side  -  traversal skips this direction.
CONNECTOR_FLOW_BIDI     = "Bidirectional"   # All non-meter elements expected to be bidirectional.

# =============================================================================
# GAS SYSTEM TYPE NAME
# The tools never match one exact System Type name (the old SYSTEM_TYPE_GAS =
# "Natural Gas" constant was unused and is removed, 2026-11-20). A pipe is
# gas by the keyword rule below, read by pipe_graph.not_gas_system_name().
# =============================================================================

# A pipe counts as gas only when its piping System Type NAME contains one of
# GAS_SYSTEM_TYPE_KEYWORDS (substring, any case) or has one of
# GAS_SYSTEM_TYPE_WORDS as a whole word, any case (Colin, 2026-11-20). Revit
# has no gas classification (gas reads OtherPipe), so compressed air,
# condensate or fuel oil on OtherPipe would otherwise pass. "Natural Gas",
# "LP Gas", "LP", "Propane" pass; "Help" and "Pumped Discharge" do not ("lp"
# is a whole word only).
GAS_SYSTEM_TYPE_KEYWORDS = ("gas", "propane")
GAS_SYSTEM_TYPE_WORDS = ("lp",)

# =============================================================================
# DIAGNOSTIC REPORT METADATA
# =============================================================================

REPORT_SCHEMA_VERSION   = "1.0"
TOOL_NAME               = "RevitPipeSizing"
PHASE                   = "Phase 1 - Diagnostic"
