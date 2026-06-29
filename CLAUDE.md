# RevitPipeSizing — Claude Code Project Context

## Active Status
**Active Phase: 2**
**Current Sub-Task: One-Line Diagram — implemented, pending Revit test**

Update this block manually as phases and sub-tasks complete.

---

## Project Summary

PyRevit extension for automated MEP pipe sizing in Revit. Three phases:
- **Phase 1:** Traversal engine + diagnostic report + one-line diagram data
- **Phase 2:** IFGC gas pipe sizing written back to Revit model (LOCKED)
- **Phase 3:** IPC water pipe sizing — DCW, DHW, HWR + velocity visualizer in plan view (LOCKED)

Duct sizing and duct velocity visualization live in a **separate project** — see "HVAC Duct Sizing Project" section below.

---

## Phase Gating — MANDATORY

Claude shall NOT write implementation code for a future phase.

- Phase 1 is active now.
- Phase 2 begins ONLY when user says: "Phase 1 is complete. Move to Phase 2."
- Phase 3 begins ONLY when user says: "Phase 2 is complete. Move to Phase 3."

Claude may answer questions about future phases but shall not write code for them. If unclear which phase is active, ask.

---

## Runtime UX

- **Phase 1:** User clicks Diagnose → Revit prompts to pick gas meter → user picks element → traversal runs → output printed to PyRevit window. One pick is the only input.
- **Phase 2:** User selects gas meter → clicks Size Gas → SelectFromList of 37 IFGC table options (material + pressure + pressure drop + table ID) → fully automatic. Specific gravity hardcoded 0.60. One-Line button generates a schematic DraftingView diagram after sizing.
- **Phase 3:** User selects water meter → clicks Size Water → one startup dialog (pipe material, street supply pressure, system type) → fully automatic. Sizing per IPC: WSFUs → GPM → velocity/pressure. Three system types: DCW, DHW, HWR. Velocity Visual button colors water pipes in a selected floor plan view by velocity range.

Single meter per system. User selects it. No auto-detection of meter.

---

## Revit Model Parameters

**Gas meter family:** Zero custom shared parameters. User selects it at runtime. Script starts traversal from the selected element.

**Gas fixture families:**
| Parameter | Type | Example |
|---|---|---|
| `GAS_LOAD_MBH` | Number | 64.8 |
| `FIXTURE_NAME` | Text | "RTU-1" |
| `IS_GAS_FIXTURE` | Yes/No | Yes |

**PRV families:** No custom shared parameters. Auto-detected during traversal by family name or Revit category.

**Water fixture families (Phase 3 — define constants now, implement later):**
| Parameter | Type | Example |
|---|---|---|
| `CW_Fixture_Units` | Number | 2.0 |
| `HW_Fixture_Units` | Number | 1.5 |
| `Fixture_Name` | Text | "LAV-1" |
| `Is_Water_Fixture` | Yes/No | Yes |

---

## Phase 1 Sub-Tasks (build in exact order)

Complete and validate each before starting the next.

### 1.1 — shared_params.py
Single file with all parameter name strings as constants. No hardcoded strings anywhere else in the codebase.

Contains:
- Gas fixture params: `Gas_BTU_Input`, `Fixture_Name`, `Is_Gas_Fixture`
- Water fixture params (constants only, no implementation): `CW_Fixture_Units`, `HW_Fixture_Units`, `Is_Water_Fixture`
- Revit built-in parameter refs for pipe diameter, pipe length, system type
- `SPECIFIC_GRAVITY = 0.60`

Acceptance: All param names defined as constants. File is importable. Zero hardcoded strings elsewhere.

### 1.2 — revit_helpers.py
Safe Revit API wrappers with null handling and logging on every call.

Functions:
- `get_parameter_value(element, param_name)` → value or None + log
- `get_element_location(element)` → XYZ tuple or None
- `get_connectors(element)` → list of connector info dicts
- `get_pipes_by_system(system_type_name)` → filtered pipe collector
- `get_all_gas_fixtures()` → elements where Is_Gas_Fixture = Yes
- `validate_selected_element(element)` → True/False with reason
- `get_pipe_length_feet(pipe)` → length in feet
- `get_pipe_diameter_inches(pipe)` → nominal diameter in inches

Acceptance: No function crashes on null. Every function logs element ID + param name + value or "NOT FOUND". User runs test and output matches Revit model properties.

### 1.3 — pipe_graph.py
Connector traversal engine. Walks from meter through every pipe, fitting, and fixture. Builds network graph.

Contains:
- `build_network(origin_element)` → graph data structure
- Graph: nodes (meter, tees, elbows, fixtures) + edges (pipes with length and diameter)
- Branch detection: tee fittings with 3+ connectors
- Loop detection: prevents infinite traversal
- Developed length per path
- Longest run: meter to farthest fixture
- Cumulative BTU load at every node

Acceptance: Every fixture reached. No pipes double-counted. Branch points correct. Longest run verified against manual measurement. Cumulative loads verified against hand calc. Disconnected elements flagged, not silently skipped.

### 1.4 — report_generator.py
Takes graph from pipe_graph.py. Produces two outputs.

Contains:
- `format_diagnostic_output(graph)` → formatted string printed to PyRevit output window for copy/paste debugging. No file saved.
- `generate_one_line_data(graph)` → structured one-line diagram data
- Validation summary: missing params, disconnected pipes, missing fixtures, ready_for_sizing flag

Pipe label format: `2-1/2"G, 25'` with cumulative MBH on line 2, decreasing away from meter.
Fixture label format: tag on line 1, MBH on line 2. Example: `RTU-1 / 64.8 MBH`

Acceptance: Printed output contains all key sections (meter, fixtures, pipes, fittings, graph, longest run, validation). All element IDs and values match Revit model properties. Output is copy/pasteable into a conversation for Claude to read. One-line labels match firm format exactly.

### 1.5 — Diagnose.pushbutton/script.py
Single-click entry point.

Flow:
1. `PickObject()` — Revit prompts "Select the gas meter element". Escape cancels cleanly.
2. `validate_selected_element()` — if invalid, show error and exit
3. `build_network(selected_element)`
4. `format_diagnostic_output(graph)` — print to PyRevit output window
5. `generate_one_line_data(graph)`
6. Print summary: fixture count, total MBH, longest run, warnings/errors
7. If errors, list them clearly

Click button → pick meter → results. One pick, no other dialogs.

Acceptance: Runs to completion on test model. Report JSON saved. Summary correct. Non-meter selection shows clear error. Disconnected pipes flagged without crash.

---

## Diagnostic Report Purpose

The diagnostic output is a **debugging tool**, not a data pipeline. The Diagnose script prints formatted output to the PyRevit output window. User copies it and pastes it into a conversation so Claude can see exactly how the code is reading the Revit model — what it found, what it missed, how the graph is structured. No file is saved. Phase 2 does NOT use diagnostic output as input. Phase 2 runs `pipe_graph.py` and `revit_helpers.py` directly against the Revit model.

Diagnostic report key sections:
- `report_metadata` — timestamp, versions, execution time
- `system_origin` — selected meter: ID, family name, location XYZ, connector count, validation
- `fixtures_found` — IDs, names, BTU, locations, connections, validation
- `pipes_found` — IDs, diameters, lengths, connectors, connection status
- `fittings_found` — IDs, connector counts, connected elements
- `network_graph` — adjacency list (nodes + edges)
- `one_line_data` — formatted segments and fixtures
- `longest_run` — total length, path, farthest fixture
- `system_summary` — totals, sizing table reference
- `disconnected_elements` — broken connections
- `validation_summary` — pass/fail checks, warnings, errors, `ready_for_sizing` flag

---

## One-Line Diagram Standard

Based on firm standard (GAS_PIPING_ONE_LINE_EXAMPLE.pdf).

**Layout:** Left to right. Meter on far left. Elevation vertical. No scale. Revit drafting view.

**Symbols:**
- Gas meter: circle with "M"
- Isolation valve: bowtie (two triangles meeting at points)
- PRV: bowtie with circle at junction (auto-detected by family name/category)
- Equipment/fixture: three horizontal lines

**Pipe segment label:**
```
2-1/2"G, 25'
4762 MBH
```
Size with "G" suffix, comma, length in feet. Cumulative MBH on line 2. MBH decreases away from meter.

**Fixture label:**
```
RTU-1
64.8 MBH
```

**Notes block (top-left of diagram):**
```
CONTRACTOR SHALL SUBMIT APPLICATIONS TO UTILITY AND COORDINATE NEW METER SERVICE
GAS PIPING SIZED FOR [X] PSI
MAX PRESSURE LOSS OF [X] PSI PER IFGC TABLE 402.4([X])
TOTAL CONNECTED LOAD: [X] MBH
TOTAL DEVELOPED LENGTH: [X]'
```

---

## Absolute Rules

**1. NEVER GUESS AT REVIT API.**
If unsure whether a property, method, or parameter exists, say: "I believe this works as follows, but please verify: [description]." Never assume.

**2. USE THE DIAGNOSTIC REPORT.**
When user pastes JSON, read it carefully. Base all analysis on report data, not assumptions about the model.

**3. NEVER SIZE FROM MEMORY.**
All sizing shall reference IFGC/IPC tables in `ifgc_gas_sizing_tables.json` or uploaded code PDFs only.

**4. MEP ENGINEERING LANGUAGE.**
Use "shall" for requirements. Use MBH, CFH, GPM, WSFU. Professional engineering documentation standard.

**5. PYREVIT PYTHON.**
IronPython 2.7. Use `rpw` where available, `Autodesk.Revit.DB` otherwise. UI via `pyrevit.forms` for error messages only. No unnecessary dialogs.

**6. LOG EVERYTHING.**
Every Revit API call result shall be captured. No silent failures. Every function logs element ID + what it found or "NOT FOUND."

**7. MODULAR CODE.**
Modules: `shared_params.py`, `revit_helpers.py`, `pipe_graph.py`, `report_generator.py`, `gas_tables.py`, `sizing_engine.py`. No logic duplication across modules.

**8. MINIMIZE USER INPUT.**
Phase 1: zero input. Phase 2: one startup dialog only, then automatic. No runtime prompts after startup.

---

## Technical Reference

```
# Pipe length (Revit internal units = feet in Revit 2022+, convert if needed)
pipe.Location.Curve.Length

# Pipe diameter
pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)

# Connectors
element.ConnectorManager.Connectors  →  connector.AllRefs

# Branch point detection
tee fittings with 3 connectors

# System filter
MEPSystemType name == "Gas"

# Shared parameters
element.LookupParameter("param_name")

# Unit conversions
BTU ÷ 1000 = MBH
BTU ÷ 1000 = CFH  (natural gas, 1000 BTU/cf)

# Specific gravity
SPECIFIC_GRAVITY = 0.60  (hardcoded, matches all IFGC tables)

# Longest Run Method (2024 IFGC Appendix A)
#
# 1. Identify the single longest DEVELOPED LENGTH from meter to farthest outlet.
#    Developed length = actual pipe run length + elbow equivalent lengths.
#    Elbow equivalent = 5 ft per elbow (per IFGC).
#    This one length is used for ALL segment table lookups in the system.
#
# 2. For each pipe segment, determine its MBH demand:
#    - Terminal segment (feeds one fixture): that fixture's MBH load
#    - Trunk segment (feeds multiple fixtures): sum of all downstream MBH loads
#
# 3. Look up the IFGC sizing table at the longest developed length row.
#    Per IFGC A103.1 Step 5: if the actual length does not match a table row,
#    use the NEXT LONGER row (conservative - lower capacity for same pipe size).
#
# 4. For each segment, select the SMALLEST nominal pipe size whose table
#    capacity (CFH) >= that segment's MBH demand.
#    1 MBH = 1 CFH for natural gas at 1000 BTU/cf.
#
# Key rule: Individual run lengths to each outlet are NOT used for sizing.
# The single longest developed length drives ALL table lookups.
#
# Example:
#   Outlet A: 100 MBH at 150 ft developed (farthest — sets system length)
#   Outlet B: 250 MBH at 120 ft developed (closer, but still sized at 150 ft)
#   Main trunk before first branch: 350 MBH, sized at 150 ft table row
```

---

## File Structure

```
Comcheck.extension/
├── CLAUDE.md
├── RJA Tools.tab/
│   ├── Gas Sizing.panel/
│   │   ├── Diagnose.pushbutton/
│   │   │   └── script.py
│   │   ├── Size Gas.pushbutton/
│   │   │   └── script.py
│   │   └── One-Line.pushbutton/
│   │       └── script.py
│   ├── Water Sizing.panel/        <- Phase 3
│   │   ├── Size Water.pushbutton/
│   │   │   └── script.py
│   │   └── Velocity Visual.pushbutton/
│   │       └── script.py
│   └── Duct Velocity.panel/       <- HVAC Duct Sizing Project Phase 1
│       └── Duct Velocity.pushbutton/
│           └── script.py
└── lib/
    ├── shared_params.py
    ├── revit_helpers.py
    ├── pipe_graph.py
    ├── report_generator.py
    ├── gas_tables.py
    ├── sizing_engine.py
    └── ifgc_gas_sizing_tables.json  (37 tables, 4 materials)
```

---

## Debugging Workflow

1. User runs Diagnose in Revit (select meter → click Diagnose button)
2. User pastes diagnostic JSON into conversation
3. Claude reads report — what was found, what was missed, how the graph looks
4. Claude explains reasoning, proposes fix
5. User applies fix, reruns, pastes updated report
6. Repeat until validation summary shows `ready_for_sizing: true` with no errors

---

## Phase 2 Summary (complete — do not re-implement core)

Gas pipe sizing per IFGC Longest Run Method. Built and working:
- `gas_tables.py` — 37 IFGC Table 402.4 tables, 4 materials (Schedule 40 Steel,
  Semirigid Copper Tubing, CSST, Polyethylene Plastic Pipe), natural gas and propane.
  `TABLE_OPTIONS` list drives the startup SelectFromList in Size Gas and One-Line.
- `sizing_engine.py` — Longest Run Method engine, writes sizes back to Revit model
  via three transactions (pipes, fittings, fixture stub pipes).
- `Size Gas.pushbutton` — full flow: pick meter, pick table, traverse, size, write to model.
- `One-Line.pushbutton` — traverses graph, computes elevation-aware schematic layout,
  draws DraftingView with meter symbol, pipe lines, fixture symbols, valve bowties, labels,
  and notes block.
- `Diagnose.pushbutton` — diagnostic traversal and report. Gas Report button removed
  (redundant with Diagnose + Size Gas terminal output).

Remaining Phase 2 deferred items (do not implement without user request):
- Fixture Editor: WPF DataGrid for assigning IS_GAS_FIXTURE / Name / MBH to unknown
  equipment families (RTUs, water heaters without the custom params).
- Minor fitting resize: transitions, ball valves, PRVs skip Nominal Radius approach.

## Phase 3 Reference (LOCKED — do not implement)

Water pipe sizing per IPC. Three system types: DCW, DHW, HWR. WSFUs → GPM via IPC tables. Velocity-based sizing. Pressure loss calculations. Water heater as hot/cold split point. Recirculation loop handling required (circular graph, unlike gas tree).

Velocity visualizer: After sizing, a separate Velocity Visual button colors each water pipe segment in a user-selected floor plan view using Revit `OverrideGraphicSettings`. Color bands: green (≤ 4 fps DCW / ≤ 3 fps DHW + HWR), yellow (approaching limit), red (exceeds limit). Existing overrides are cleared before applying new ones.

---

## HVAC Duct Sizing Project

**Separate project** — independent of the pipe sizing phases above. Lives in the same Comcheck.extension but has its own panel and phase gating.

**Active Phase: 1 — Duct Velocity Visualizer**

### Phase 1 Overview

Color ductwork in a copied floor plan view by velocity (green / yellow / red) using firm design defaults. Activated by clicking an AHU or fan unit.

**Current status:** Duct velocity coloring displays correctly. Pending further testing on real project scenarios to catch outlier errors and velocity bugs.

**UX flow:**
1. User clicks "Duct Velocity" button in the Duct Velocity panel
2. Revit prompts: select AHU/fan/equipment element
3. Settings dialog opens — per-system max velocity (FPM) + max friction (iwc/100ft) + green threshold %
4. Script traverses all duct systems connected to that equipment (SA, RA, EA, OA)
5. Copies the relevant floor plan view → places on a new sheet named "Ducting Velocities - [Source Sheet Number]"
6. Applies `OverrideGraphicSettings` color fills to each duct in the copied view; fittings/accessories inherit color from worst adjacent duct
7. Ducts with no CFM data → gray (flagged, not skipped silently)
8. Terminal output: aligned flagged ducts table (RED first, then YELLOW, sorted by velocity)

**Velocity color logic — percentage-based off firm defaults:**
| System | Default Max Vel | Default Max Friction | Green | Yellow | Red |
|--------|----------------|---------------------|-------|--------|-----|
| Supply Air (SA) | 800 FPM | 0.08 iwc/100ft | < 85% of max | 85–100% of max | > max |
| Return Air (RA) | 600 FPM | 0.05 iwc/100ft | < 85% of max | 85–100% of max | > max |
| Exhaust Air (EA) | 600 FPM | 0.05 iwc/100ft | < 85% of max | 85–100% of max | > max |
| Outside Air (OA) | 600 FPM | 0.05 iwc/100ft | < 85% of max | 85–100% of max | > max |

All defaults are editable in the settings dialog before each run.

**Friction formula:** `ΔP/100ft = 6.82×10⁻⁶ × V^1.82 / Dh^1.22`
Basis: standard air (0.075 lb/ft³, 70°F, sea level), ε = 0.0003 ft (galvanized steel). Matches ductulator.com output.

**CFM source:** `OST_DuctTerminal` leaf nodes (diffusers, grilles, registers). Parameter lookup order: `Flow` → `Airflow` → `Air Flow` → `CFM` → `RBS_DUCT_FLOW_PARAM`. VAV boxes are pass-through nodes. Each duct segment CFM = sum of all downstream terminal CFMs.

**Design decisions (resolved 2026-06-29):**
1. CFM from terminal families only — NOT from duct segments or VAV box parameters.
2. Sheet naming — `"Ducting Velocities - [Source Sheet Number]"`.
3. Multi-system — all connected systems (SA, RA, EA, OA) colored in one pass.
4. Multi-floor — traversal follows connectors across floors; copies all affected floor plan views.

**Phase gating:**
- Phase 1 (velocity visualizer) is **ACTIVE**
- Phase 2 (duct sizing — writes sizes back to model) begins ONLY when user says: "Move to HVAC Phase 2."
- Never write Phase 2 code without that trigger.

**Standalone tool:** `ductulator.py` at `C:\Users\Colin Nolan\Developer Tools\ductulator.py` — CLI duct sizing calculator with color-coded velocity output. Already built and working.

---

## Verification

Run this to confirm context loaded correctly:

```
Confirm project status:
1. What phase is active and what is the current sub-task?
2. If I ask you to build the sizing engine, what do you do?
3. What happens when the user clicks Diagnose?
4. What is the diagnostic report for?
5. Does Phase 2 read the diagnostic JSON as input?
6. What parameters does the gas meter family need?
7. What is the pipe label format?
8. If unsure about a Revit API call, what do you do?
```

Expected answers:
1. Phase 2. Current sub-task: One-Line Diagram — implemented, pending further testing.
2. Decline. Phase 3 is locked. Can answer questions about it, will not write implementation code. Duct sizing is a separate project — see HVAC Duct Sizing Project section.
3. User selects meter → clicks button → traversal → diagnostic JSON saved → summary shown. No dialogs.
4. Debugging tool. Pasted into conversation so Claude can see how the code is reading the model. Not a data pipeline.
5. No. Phase 2 runs pipe_graph.py and revit_helpers.py directly against the Revit model.
6. Zero. No custom shared parameters on the meter family.
7. `2-1/2"G, 25'` — size with G suffix, comma, length. Cumulative MBH on line 2.
8. State the assumption explicitly. Ask user to verify before writing code.
