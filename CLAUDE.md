# RevitPipeSizing — Claude Code Project Context

## Active Status
**Active Phase: 1**
**Current Sub-Task: 1.5 — Diagnose.pushbutton/script.py**

Update this block manually as phases and sub-tasks complete.

---

## Project Summary

PyRevit extension for automated MEP pipe sizing in Revit. Three phases:
- **Phase 1:** Traversal engine + diagnostic report + one-line diagram data
- **Phase 2:** IFGC gas pipe sizing written back to Revit model (LOCKED)
- **Phase 3:** IPC water pipe sizing — DCW, DHW, HWR (LOCKED)

---

## Phase Gating — MANDATORY

Claude shall NOT write implementation code for a future phase.

- Phase 1 is active now.
- Phase 2 begins ONLY when user says: "Phase 1 is complete. Move to Phase 2."
- Phase 3 begins ONLY when user says: "Phase 2 is complete. Move to Phase 3."

Claude may answer questions about future phases but shall not write code for them. If unclear which phase is active, ask.

---

## Runtime UX

- **Phase 1:** User selects gas meter → clicks Diagnose. Script traverses from selected element. Zero dialogs. Zero prompts. One selection is the only input.
- **Phase 2:** User selects gas meter → clicks Size Gas → one startup dialog (pipe material, inlet pressure, pressure drop) → fully automatic. Specific gravity hardcoded 0.60.
- **Phase 3:** User selects water meter → clicks Size Water → one startup dialog (pipe material, street supply pressure, system type) → fully automatic. Sizing per IPC: WSFUs → GPM → velocity/pressure. Three system types: DCW, DHW, HWR.

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
1. Get currently selected element
2. `validate_selected_element()` — if invalid, show error: "Please select the gas meter element"
3. `build_network(selected_element)`
4. `format_diagnostic_output(graph)` — print to PyRevit output window
5. `generate_one_line_data(graph)`
6. Print summary: fixture count, total MBH, longest run, warnings/errors
7. If errors, list them clearly

Select meter → click button → results. No dialogs.

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

**0. ASK BEFORE MODIFYING CLAUDE.md.**
Always ask the user for explicit permission before making any edit to this file. State what change is proposed and why. Do not modify CLAUDE.md without confirmation, even for small corrections.

**1. NEVER GUESS AT REVIT API.**
If unsure whether a property, method, or parameter exists, say: "I believe this works as follows, but please verify: [description]." Never assume.

**2. VERIFY BEFORE CODING.**
State assumptions explicitly. Ask user to confirm before writing code based on assumptions.

**3. USE THE DIAGNOSTIC REPORT.**
When user pastes JSON, read it carefully. Base all analysis on report data, not assumptions about the model.

**4. NEVER SIZE FROM MEMORY.**
All sizing shall reference IFGC/IPC tables in `ifgc_gas_sizing_tables.json` or uploaded code PDFs only.

**5. MEP ENGINEERING LANGUAGE.**
Use "shall" for requirements. Use MBH, CFH, GPM, WSFU. Professional engineering documentation standard.

**6. PYREVIT PYTHON.**
IronPython 2.7. Use `rpw` where available, `Autodesk.Revit.DB` otherwise. UI via `pyrevit.forms` for error messages only. No unnecessary dialogs.

**7. LOG EVERYTHING.**
Every Revit API call result shall be captured. No silent failures. Every function logs element ID + what it found or "NOT FOUND."

**8. MODULAR CODE.**
Modules: `shared_params.py`, `revit_helpers.py`, `pipe_graph.py`, `report_generator.py`, `gas_tables.py`, `sizing_engine.py`. No logic duplication across modules.

**9. EXPLAIN BEFORE FIXING.**
Never propose "try this." Explain WHY a fix works, state the reasoning, then propose the code.

**10. MINIMIZE USER INPUT.**
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

# Longest Run Method
ONE longest run length sizes ALL segments in the system
```

---

## File Structure

```
RevitPipeSizing.extension/
├── CLAUDE.md
├── RevitPipeSizing.tab/
│   ├── Gas Sizing.panel/
│   │   ├── Diagnose.pushbutton/
│   │   │   └── script.py
│   │   ├── Size Gas.pushbutton/
│   │   │   └── script.py
│   │   ├── Gas Report.pushbutton/
│   │   │   └── script.py
│   │   └── One-Line.pushbutton/
│   │       └── script.py
│   └── Water Sizing.panel/        ← Phase 3
└── lib/
    ├── shared_params.py
    ├── revit_helpers.py
    ├── pipe_graph.py
    ├── report_generator.py
    ├── gas_tables.py              ← Phase 2
    ├── sizing_engine.py           ← Phase 2
    └── ifgc_gas_sizing_tables.json
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

## Phase 2 Reference (LOCKED — do not implement)

Gas pipe sizing per IFGC Longest Run Method. One startup dialog: pipe material, inlet pressure, pressure drop. Sizes every segment. Writes sizes back to Revit model via Transaction. Table selection logic is in `gas_sizing_engine.py`. Tables are in `ifgc_gas_sizing_tables.json`.

## Phase 3 Reference (LOCKED — do not implement)

Water pipe sizing per IPC. Three system types: DCW, DHW, HWR. WSFUs → GPM via IPC tables. Velocity-based sizing. Pressure loss calculations. Water heater as hot/cold split point. Recirculation loop handling required (circular graph, unlike gas tree).

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
1. Phase 1. Sub-task 1.1 — shared_params.py.
2. Decline. Phase 2 is locked. Can answer questions, will not write implementation code.
3. User selects meter → clicks button → traversal → diagnostic JSON saved → summary shown. No dialogs.
4. Debugging tool. Pasted into conversation so Claude can see how the code is reading the model. Not a data pipeline.
5. No. Phase 2 runs pipe_graph.py and revit_helpers.py directly against the Revit model.
6. Zero. No custom shared parameters on the meter family.
7. `2-1/2"G, 25'` — size with G suffix, comma, length. Cumulative MBH on line 2.
8. State the assumption explicitly. Ask user to verify before writing code.
