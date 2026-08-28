# GRD Auto-Sizing Families — Cross-Examination Plan

**Purpose:** before redoing all Supply/Return/Exhaust diffuser families in Revit 2022,
audit the current (Revit 2026) state of all 6 built families + the 1 unbuilt one against
each other and against the new friction requirement, so the rebuild is a clean,
consistent re-execution — not a copy of unresolved inconsistencies.

**Naming note (2026-08-26):** Families are now referred to by what they physically are
(mount type + function), not by GRD schedule tag (SD-1, RG-2, EG-1, etc.). The schedule
tags still exist and still matter for drawing/schedule purposes — they're preserved as
`schedule_tags` in each lookup data file's `_meta` — but they're no longer the primary
way this project's own files and documentation refer to a family. Mapping, for reference:

| Descriptive name (used below) | Schedule tag(s) | Mount |
|---|---|---|
| Ceiling Supply Diffuser (Round Neck) | SD-1, SD-2 | Ceiling, round neck |
| Sidewall Supply Diffuser | SD-4, SD-5 | Sidewall |
| Ceiling Return Grate (Round Neck) | RG-1 | Ceiling, round neck |
| Sidewall Return Grate | RG-2 | Sidewall |
| Ceiling Exhaust Grate (Round Neck) | EG-1 | Ceiling, round neck |
| Sidewall Exhaust Grate | EG-2 | Sidewall |
| Slot Diffuser (Supply/Return) | SD-3, RG-3 | Slot |

**How to use this doc:** this is written for a fresh Claude Code instance with no
memory of the build history. Full history/reasoning lives in Claude Code memory
(`project_smart_diffuser_family.md`, ~1900 lines) — this file is the condensed,
actionable version. Where this doc says "verify live," don't trust the number here;
pull it from the actual open Family Editor session via Revit MCP first.

**Context:** Colin hand-edited the Ceiling Supply Diffuser family (24x24, Rectangular -
Round Neck) in Revit 2026 as a worked reference before this rebuild started. Its final
parameter names, formula style, and Sizing Warning pattern are the template for the
other families.

---

## 1. Current state — all families (Revit 2026, all confirmed same product-version metadata issue, all need Revit 2022 rebuild from scratch — Revit cannot open a newer-version family in an older release)

| Family | Family name (2026) | Sizing basis | Target | Cutsheet | Ceiling(s) | Ratio caps | Verified |
|---|---|---|---|---|---|---|---|
| Ceiling Supply Diffuser (Round Neck) | `RJA - Supply Diffuser - Rectangular - Round Neck (CFM Auto-Sizing)` | NC (TMS only, OMNI dropped) | NC<15 | TMS (`TMSperf_diffusers_2017.pdf`) | 209/327/535 cfm (12x12/18x18/24x24) | n/a (single diameter per Flow) | 16-18 cases + friction-corrected 2026-08-26 (see Section 2, Item 1) |
| Sidewall Supply Diffuser | `RJA - Supply Diffuser - Sidewall (CFM Auto-Sizing)` | NC, +7 deflection correction | NC<20 (raw<13) | 272RL (`aero_supplyperf_aeroblade_2017.pdf`) | 1500 cfm (24x24/48x12 tier) | Thin ≤4:1, Regular best-available (2:1 not always achievable) | 34-36 cases; friction-checked 2026-08-26, no changes needed (all 32 tiers × Regular/Thin already comply with 0.08 in.wc/100ft) |
| Ceiling Return Grate (Round Neck) | `RJA - Return Grate - Rectangular - Round Neck (CFM Auto-Sizing)` | **Pure 600 FPM velocity — NC abandoned entirely** | n/a | Perforated Return (PAR/PXP/PMR), velocity math cross-validated against it | 327/838/1584 cfm (12x12/18x18/24x24, caps = largest diameter that physically fits each face) | n/a | 17-18 cases; friction check 2026-08-26 found 6"/8" fail 0.05 limit — fix identified, not yet applied (Section 2, Item 1) |
| Sidewall Return Grate | `RJA - Return Grate - Sidewall (CFM Auto-Sizing)` | NC, no deflection correction (fixed-blade) | NC<15 | 350R/350F/350R-SS (`350.355Rperf_300_350_2017.pdf`) | 1500 cfm (same tier as supply) | Thin ≤4:1, Regular best-available | 36 cases; friction-checked 2026-08-26, one tier fails (6x6 @ 114 cfm, 0.053 vs 0.05 limit) — fix identified, not yet applied |
| Ceiling Exhaust Grate (Round Neck) | `RJA - Exhaust Grate - Rectangular - Round Neck (CFM Auto-Sizing)` | Identical to Ceiling Return Grate (byte-identical formulas) | n/a | Same as Ceiling Return Grate | Same as Ceiling Return Grate | n/a | 11 cases + re-confirmed live 2026-07-24; same friction fix as Ceiling Return Grate applies here, not yet applied |
| Sidewall Exhaust Grate | `RJA - Exhaust Grate - Sidewall (CFM Auto-Sizing)` | NC, reused Sidewall Return Grate's breakpoint array verbatim (23R rejected — half the capacity of 350R at same NC) | NC<15 (inherited) | 350R/355R (same as Sidewall Return Grate) | 1500 cfm | 4:1/2:1 ratio-cap pass CONFIRMED applied, verified live 2026-08-26 against Sidewall Return Grate's post-cap array (all 32 tiers, width+height, byte match) | 38 cases; same one-tier friction fix as Sidewall Return Grate applies here, not yet applied |
| Slot Diffuser (Supply/Return) | Shared family, name TBD | **Not built** | Not decided | ML-38 (`MLperf_diffusers_2017.pdf`, section starts pdftotext line ~59, not the ML-37 section at top) | n/a | n/a | 0 |

All 7 schedule tags trace back to `Downloads\Schedules_GRD.xls` (GRD-TITUS SCHEDULE sheet).
Confirmed live against that schedule 2026-08-26: Sidewall Supply Diffuser's schedule tags
are SD-4 (Titus S301FL, mounts to spiral duct) and SD-5 (Titus 272RL, wall mount) — Colin
confirmed these can be treated as one family/data set (S301FL uses an air scoop transition
to the same type of diffuser, doesn't change performance data). The schedule also lists
Sidewall Return/Exhaust Grate's model as "Titus 23 RL," which conflicts with the as-built
350R data — resolved as: 23R was evaluated and explicitly rejected (half the capacity of
350R at the same NC), 350R is correct, the schedule's model column is stale and should be
corrected to 350R when the schedule is next touched (not urgent, doesn't block the rebuild).

---

## 2. Open items — resolved 2026-08-26, kept here for the rebuild record

### Item 1 — Friction/pressure-drop cross-check (RESOLVED: applies to ALL families, including sidewall)
Confirmed 2026-08-26: a family can pick a neck size that satisfies its own NC/capacity
data while failing RJA's own duct friction standard used by the Duct Velocity tool
(`lib/hvac_graph.py`, `FIRM_DEFAULTS`: Supply Air 0.08 in.wc/100ft, Return/Exhaust Air
0.05 in.wc/100ft). This is the real-world bug Colin hit in the field (diffuser "auto
sized" but the neck was still undersized). Colin's decision: check ALL families,
including sidewall (derive hydraulic diameter from the grille's Duct Width/Duct Height
rectangle: `d_h = 2*w*h/(w+h)`), not just the round-neck ceiling families.

**Audit result (all 6 built families, every breakpoint, both sides of every boundary):**

| Family | Result |
|---|---|
| Ceiling Supply Diffuser (Round Neck) | **Fixed & verified live 2026-08-26.** 6" cap 157→**112** CFM, 8" cap 244→**242** CFM (18x18/24x24 tiers only; 12x12 tier's 8" cap of 209 was already compliant). 10"/12"/14" needed no change. `Duct Diameter` formula updated via `FamilyManager.SetFormula`, 14 boundary cases regenerate-and-read, all pass. |
| Sidewall Supply Diffuser | **No changes needed.** All 32 tiers × Regular/Thin (64 checks) already comply with 0.08 — grille face area is large enough relative to CFM that velocity/friction never approaches the limit. |
| Ceiling Return Grate (Round Neck) | **Fix identified, not yet applied.** 6" cap 118→**86** CFM, 8" cap 209→**187** CFM. 10"+ unaffected. |
| Sidewall Return Grate | **Fix identified, not yet applied.** One tier only: 6x6 cap 114→**110** CFM (barely over: 0.053 vs 0.05 limit). Every other tier of 32 already compliant. |
| Ceiling Exhaust Grate (Round Neck) | Same fix as Ceiling Return Grate applies (identical formulas) — not yet applied. |
| Sidewall Exhaust Grate | Same one-tier fix as Sidewall Return Grate applies (identical breakpoint array) — not yet applied. |

Formula for the check (matches `duct_friction_loss_per_100ft()`):
```
area_ft2 = pi * (diameter_in / 24)^2                      # round
area_ft2 = (w_in * h_in) / 144                             # rectangular
d_h_in   = diameter_in                                     # round
d_h_in   = 2 * w_in * h_in / (w_in + h_in)                  # rectangular, hydraulic diameter
fpm      = cfm / area_ft2
friction = 6.82e-6 * fpm^1.82 / d_h_in^1.22
```
Where a breakpoint fails, the cascade steps up to the next size earlier — this makes that
size's effective capacity SMALLER than what's currently built, not bigger. All new
breakpoints above were verified on both sides of the boundary (regenerate-and-read
standard, Section 4) before being called final.

### Item 2 — Sidewall Supply Diffuser's SD-4/S301FL question (RESOLVED)
The GRD schedule lists SD-4 as Titus S301FL (mounts to spiral duct) vs SD-5 as Titus
272RL (wall mount) — different mount types. Colin confirmed both can be treated as one
family/data set: S301FL uses an air scoop duct-transition fitting to essentially the same
type of diffuser as 272RL, so the mount-type difference doesn't change the performance
data. No separate S301FL cutsheet pull needed. Sidewall Supply Diffuser stays as one
family/data set for both schedule tags.

### Item 3 — Ceiling families use pure velocity, sidewall families use NC (not revisited, not blocking)
Ceiling Return/Exhaust Grate use pure 600 FPM velocity sizing (NC data was reviewed and
explicitly rejected by Colin: "revisit the cutsheet use the 600 fpm disregard the nc
values"). Ceiling/Sidewall Supply Diffuser and Sidewall Return/Exhaust Grate all still use
NC-driven cascades. This is a real methodological split between mounting types, made
through genuine back-and-forth decisions — worth a single explicit confirmation before
locking it into the rebuild, but not raised again during this session; carry forward as-is
unless Colin says otherwise.

### Item 4 — Sidewall Return Grate's max-CFM ceiling increase (still open, not blocking)
Colin asked for Sidewall Return Grate's max CFM to be raised at some point mid-project;
this got sidetracked into the NC-retarget work and was never given a target number. Still
capped at 1500 CFM (24x24/48x12 tier) like every other sidewall family. Get an actual
number before or during the rebuild, or confirm it's fine to leave at 1500.

### Item 5 — Sidewall Exhaust Grate's ratio-cap status (RESOLVED, verified live)
Confirmed 2026-08-26 by pulling Sidewall Exhaust Grate's live `Auto Width`/`Auto Height`
formulas via Revit MCP and comparing all 32 breakpoints (Thin and Regular, width and
height) against Sidewall Return Grate's post-cap lookup file. Every value matches
byte-for-byte, including the three Thin-collapse cases from the 4:1 ratio cap (555, 675,
795 CFM tiers). **Sidewall Exhaust Grate has the post-cap, NC<15 array — safe to use as a
copy source**, no stale pre-cap data.

### Item 6 — Slot Diffuser (RESOLVED/BUILT 2026-08-26 for supply; return still to do)
Was never started. Now built. Decisions and findings from this session:

**Source model / data.** ML-38 (3/4" slot spacing — NOT 1"; ML-39 is the 1" model, easy
mix-up since four model tables stack on the same cutsheet pages). Supply data from
`MLperf_diffusers_2017.pdf` page 2. **Return does NOT reuse supply data** — MLR-38 has its
own dedicated table on page 4, structured differently (organized by static-pressure step
rather than the supply table's columns). That answers the old "does return reuse supply"
question for this model. Two separate families per Colin, so the return family gets built
from the MLR-38 table when its turn comes.

**Target NC<20** (Colin's call), matching the Sidewall Supply Diffuser target.

**Architecture — differs from the other 6 on purpose.** Types = **Length**
(24"/36"/48"/60"/72", Titus's standard nominal lengths, driven by `Select a Plenum Nominal
Size` = 1..5). The auto-sized output is **Number of Slots (1-8)**, not a duct dimension:
total CFM = cfm/ft × active length, so slot count is what scales capacity. Length is the
Type because it's the coordination/architectural constraint the engineer picks first;
slot count is what CFM drives. Same principle as the other 6 (Type = the constraint,
auto-output = whatever CFM drives), different physical quantity.

**Length correction is folded into the breakpoints.** The cutsheet's NC table assumes a
4-ft length; other lengths get a correction (2ft −3, 4ft 0, 6ft +2, 8ft +3, 10ft +5).
Our types are 2/3/4/5/6 ft, and 3ft/5ft fall between published points — handled the
project-standard conservative way: round UP to the next published point (36" uses 4ft's
correction, 60" uses 6ft's), overestimating NC rather than under.

**Max CFM per length (NC<20), and Colin's ≥350 CFM goal:**
| Type | Max CFM | Meets ≥350? |
|---|---|---|
| 24" | 250.5 | no (physical limit) |
| 36" | 289.8 | no (physical limit) |
| 48" | 388.8 | yes |
| 60" | 487.8 | yes |
| 72" | 586.8 | yes |

24" and 36" cannot reach 350 CFM at NC<20 even at 8 slots — that's a real Titus catalog
limit, not a sizing choice. 48" and up clear it comfortably.

**Non-monotonic table ties are expected, not bugs.** At some length/slot combinations an
extra slot buys zero extra capacity at NC<20, so that slot count is unreachable and the
cascade correctly skips it (24": slot 8 ties slot 7; 36"/48": slot 4 ties slot 3;
60"/72": slot 3 ties slot 2). Breakpoint arrays are stored as running-max (monotonic) for
exactly this reason. A verification script that assumes all 8 slot counts are reachable
will report false failures — filter to distinct transitions first.

**`Max CFM` readout parameter** (instance, read-only via formula) shows the NC<20 ceiling
for the currently-selected Length type, so an engineer never has to guess-and-check: pick
the Length the architecture allows, read Max CFM, compare to load. `Sizing Warning` fires
above it ("OVER CAPACITY - INCREASE SLOT COUNT OR LENGTH").

**Round inlet (replaces Titus's flat-oval).** Done at the connector level — see gotcha 21.
Auto-sized 6"/8"/10"/12" from flow, checked against BOTH the Supply Air friction limit
(0.08 in.wc/100ft) and the 800 FPM velocity limit, whichever governs: 6" ≤112 CFM
(friction-governed), 8" ≤242 (friction), 10" ≤436 (**velocity**-governed — 439 would pass
friction but exceed 800 FPM, so 436 is the correct cap), 12" above that. The inlet is
never the binding constraint — every length's Max CFM is within its inlet's capacity.

**Connector classification had to be corrected — Titus shipped it wrong for supply.** The
live connector came configured as `System Classification = Return Air` with
`Flow Direction = Out`, i.e. a RETURN configuration, because this is Titus's *generic*
ML(R) family covering both the ML (supply) and MLR (return) model lines. Corrected to
`Supply Air` / `In` (air enters the family from the duct; a return grille would be `Out`).
The integer mapping was probed empirically rather than guessed — Supply Air = 1,
Return Air = 2, Exhaust Air = 3, on the connector's "System Classification" parameter.
**When the RETURN family is built from this same Titus base, its connector must be set to
`Return Air` / `Out`** — do not assume the copy inherits the right one.

**Visual collar geometry left as flat-oval, deliberately.** Colin's call 2026-08-26: not
noticeable at floor-plan scale, and since the connector is now round, the duct Revit draws
into it is round too — the collar is a small graphic beneath the prominent element. See
gotcha 22 for what it would take if that ever changes.

**Still open on this family:** the leftover disabled rectangular connector has not been
deleted (decide whether to — it is the same "two connectors, ambiguous which is live"
trap described in gotcha 23), and the family has not yet been saved to the RJA library
under an RJA name (still pointing at the Titus download path under Downloads).

### Item 6b — Slot Diffuser (RETURN) — not started
Build from the MLR-38 table (`MLperf_diffusers_2017.pdf` page 4), same Length-type /
auto-slot architecture as the supply family above. Note the MLR length-correction table is
DIFFERENT from the supply one (MLR: 2ft −7, 4ft −4, 10ft 0, 15ft +2, 20ft +3, and it is
referenced to a 10-ft active section, not 4-ft) — do not reuse the supply corrections.
Return/exhaust friction limit is 0.05 in.wc/100ft and 600 FPM, not supply's 0.08/800.

### Item 7 — NC target summary (not revisited, not blocking)
| Family | Target |
|---|---|
| Ceiling/Sidewall Supply Diffuser | NC<15 (ceiling) / NC<20 effective, raw<13 (sidewall) |
| Ceiling Return/Exhaust Grate | No NC — pure 600 FPM |
| Sidewall Return/Exhaust Grate | NC<15 |

---

## 3. Reference data (copy verbatim into new families — do not re-derive)

Full breakpoint arrays, formulas, and lookup files already exist and are correct as of
their last-verified state — pull them from (renamed 2026-08-26, see Section 1 mapping):
- `GRD Sizing/lookup_data/Ceiling_Supply_Diffuser_TMS.csv`
- `GRD Sizing/lookup_data/Sidewall_Supply_Diffuser_272RL.json` / `.csv`
- `GRD Sizing/lookup_data/Ceiling_Return_Grate_Perforated.json` / `.csv`
- `GRD Sizing/lookup_data/Sidewall_Return_Grate_350R.json` / `.csv`
- `GRD Sizing/lookup_data/Ceiling_Exhaust_Grate.json` / `.csv`
- `GRD Sizing/lookup_data/Sidewall_Exhaust_Grate_350R.json` / `.csv`

Each file's `_meta.schedule_tags` preserves the original GRD schedule tag(s) for
drawing/schedule traceability. **These files do NOT yet include the friction-corrected
breakpoints from Section 2, Item 1 for the four families still pending** (Ceiling Return
Grate, Ceiling Exhaust Grate, Sidewall Return Grate, Sidewall Exhaust Grate) — apply the
corrected values from the table in Section 2 on top of these files' data, don't read them
as final without that adjustment. Ceiling Supply Diffuser and Sidewall Supply Diffuser are
both already final/correct as documented in Section 1.

---

## 4. Standard procedure + gotchas checklist (apply to every family in the rebuild)

Consolidated from the full build history — every one of these was independently
discovered the hard way once already; don't re-discover them.

**Revit API / formula gotchas:**
1. `<=` / `>=` are not valid in Revit family formulas — use strict `<`/`>` only
   (breakpoints are "max CFM for this size," so `<` rounds ties up to the next size,
   which is the conservative/safe direction — no correction needed).
2. Formula STRING literals compared against a non-length parameter (e.g. `Flow < 196`)
   use that parameter's DISPLAY unit (CFM here) — do not convert to internal CFS.
3. Direct `FamilyManager.Set(param, value)` calls (no formula, a raw double) ALWAYS use
   Revit's INTERNAL unit — DO convert via `UnitUtils.ConvertToInternalUnits` first.
   (Rules 2 and 3 are opposite conventions — this is the single easiest mistake to make
   on this project, verify by regenerate-and-read, not by "it compiled.")
4. `FamilyType.Name` becomes invalid immediately after that type is passed to
   `FamilyManager.DeleteCurrentType()` — capture the name to a local string BEFORE
   deleting, never reference the live `FamilyType` object after.
5. `Document.EditFamily()` and `Document.SaveAs()` cannot be called via
   `send_code_to_revit` in any document — the MCP tool always wraps sent code in an open
   transaction, and both of these Revit API calls require zero open transactions. Family
   Editor session opening and every Save/Save As must be done manually by Colin in the
   Revit UI.
6. `WorksharingUtils.CheckoutElements()` does not make a workset editable — it only
   checks out the specific elements. There is no reliable API path to claim
   workset-level ownership; that's a UI-only action.
7. Sending a full family rebuild (types + params + 6 formulas + per-type values) in ONE
   `send_code_to_revit` call risks a generic, undiagnosable
   `"Exception has been thrown by the target of an invocation"` failure. Break it into
   ~6 smaller sequential calls (types → delete legacy → param setup → formulas →
   per-type values + Regenerate) instead — isolates which step actually fails.
8. `FamilyType.AsDouble(param)` returns a nullable `double?` — cast/coalesce
   (`(double)(x ?? 0.0)`) before using it in arithmetic or string formatting, or the
   code fails to compile.

**Data verification discipline:**
9. When re-deriving breakpoints for a new NC target, the sort order can change (a
   corrected/retargeted value can move to a different position in the ascending
   cascade) — always re-sort and rebuild the full cascade, never patch values in place.
10. Compute expected test values PROGRAMMATICALLY from the same breakpoint array used to
    build the formula, never hand-type them — this eliminated an entire class of false
    verification failures once adopted (proven across 3+ separate rebuild sessions).
11. A quick visual re-read of a dense multi-column Titus table that conflicts with a
    previously-verified number is not, by itself, evidence the old number is wrong —
    the "miscounted leading dash cells" failure mode is real and repeats. Re-verify via
    precise `page.get_text('words')` coordinate extraction before changing anything.
12. Check each cutsheet's own notes page for its dash convention (has been both "<10"
    and "<15" on different sheets in this project) and deflection-correction table (272RL
    needs +7 NC for 45° deflection; 23R's table already states 45° directly, no
    correction; confirm per-sheet, never assume it carries over from a different sheet).
13. Validate any hand-edited JSON with `python -c "import json; json.load(open(...))"`
    immediately after editing — a partial Edit call has corrupted a JSON file at least
    once on this project by not closing an object before an array bracket.
14. Never assume two schedule tags sharing a Titus model also share one Revit family, or
    that a "shared family" from an earlier note is still accurate — confirm live via
    Family Editor inspection every time. Ceiling Return/Exhaust Grate share a model but
    are separate families; Sidewall Supply Diffuser's two schedule tags (SD-4/SD-5) were
    assumed to share one family and turned out to have different mount types on the
    schedule (spiral duct vs. wall) — resolved as intentional (Item 2), but confirm live
    rather than carrying an old assumption forward on the next project like this.
15. `effective_height_in()` in `lib/hvac_graph.py` snaps duct/neck dimensions to the
    nearest 2" nominal increment before comparison — if any NEW code in this project
    reads a raw `RBS_CURVE_*` dimension directly for a comparison (not through that
    function), apply the same snap; Revit's internal feet↔inch round-trip can store a
    true 10" as `10.000000000000002"`, which is exactly what caused a false Duct
    Velocity flag earlier this session (see `hvac_graph._nominal_even_in()`).
16. When checking whether a family's schedule tag/model number is still accurate, pull
    the actual live spreadsheet (`Downloads\Schedules_GRD.xls`) rather than trusting a
    prior session's summary of it — a summary written before a later correction (or a
    correction that was flagged but never actually applied) will read as settled fact
    when it isn't. This caught the Sidewall Supply Diffuser mount-type mismatch (Item 2)
    and the Sidewall Return/Exhaust Grate 23R-vs-350R schedule discrepancy in Section 1.
17. **`FamilyManager.SetFormula()` rejects a unit suffix on the input string for a
    non-length comparison, even though Revit adds that same suffix back when it stores
    and later reads the formula.** Confirmed on Revit 2022 rebuilding the Ceiling Supply
    Diffuser: `if(Flow < 112 CFM, 6", 8")` throws `ArgumentException: It is an invalid
    formula string.` on `SetFormula`, while the length literal `Diffuser Width < 15"`
    compiles fine. The fix is to omit the unit suffix on the SET call — `if(Flow < 112,
    6", 8")` — Revit appends `CFM` (or the parameter's current display unit)
    automatically when the formula is read back afterward. Isolated by testing
    progressively simpler formula strings against the same parameter (bare literal →
    length comparison → non-length comparison) rather than trusting the generic
    `TargetInvocationException` from a full nested formula, per gotcha 7's isolation
    approach.
18. Hand-typing a deeply nested `if(...)` chain (8-9 levels, one per breakpoint) is where
    paren-count mistakes actually happen on this project — confirmed rebuilding Ceiling
    Return Grate's `Duct Diameter` formula: a hand-typed version of the exact same logic
    that worked for Ceiling Supply Diffuser (shallower nesting) threw the same generic
    `ArgumentException: It is an invalid formula string.` as gotcha 17's unit-suffix bug,
    but the real cause here was a missing/extra parenthesis, not a unit suffix. Fix:
    build the nested-if string programmatically in C# from the same breakpoint array
    used for testing (loop from the terminal/else value backward, wrapping one more
    `if(Flow < X, size, ...)` per iteration) rather than typing the chain by hand — this
    is the formula-construction analog of gotcha 9/10's "compute test values
    programmatically" rule, and should be the default approach for any breakpoint count
    beyond a handful of tiers.
19. **Editing a family in its own Family Editor window does not update instances already
    placed in an open project — the family must be explicitly reloaded** (Family Editor
    → Modify → Load into Project, targeting the project) before a live boundary test on
    a placed instance means anything. Confirmed on Ceiling Return Grate: a live test
    immediately after finishing the formulas showed `Duct Diameter` stuck at 6" for
    every Flow value tried — not a formula bug, the placed instance was still running the
    pre-rebuild family definition. After Colin reloaded, the identical test passed
    cleanly at every breakpoint.
20. A placed Air Terminal instance's `el.Parameters` collection can enumerate "Flow"
    TWICE — once resolved via `BuiltInParameter.RBS_DUCT_FLOW_PARAM` and once as a
    plain-name match with `BuiltInParameter.INVALID` on its `InternalDefinition` — but
    these are NOT two independent values. Confirmed by setting only one and reading the
    other back afterward: both report the identical value every time. They're two
    `Parameter` object wrappers onto the same underlying storage, not a duplicate
    parameter or a family-authoring mistake. `LookupParameter("Flow")` may return either
    wrapper depending on enumeration order, which is why an earlier session-in-progress
    reading looked like "only one of the two drives the formula" — it doesn't matter
    which wrapper is used to set the value, both read/write the same thing. (This
    superseded an earlier, incorrect version of this note that treated the two wrappers
    as independent parameters requiring per-family investigation — they don't.)
21. **A connector's shape can be changed in place — no delete-and-recreate needed.**
    `ConnectorElement.Shape` is read-only in the API, but the connector's
    `BuiltInParameter.CONNECTOR_PROFILE_TYPE` parameter is NOT read-only and maps directly
    to the `ConnectorProfileType` enum (Round=0, Rectangular=1, Oval=2). Setting it to 0
    converted the Titus slot diffuser's flat-oval plenum inlet to round in one call. After
    the change the `Diameter` parameter becomes writable and unassociated (Width/Height go
    read-only), so associate `Diameter` to a family parameter with
    `FamilyManager.AssociateElementParameterToFamilyParameter()` to drive it by formula.
    Far cheaper than the delete-and-recreate path in the family-editing recipe, which
    needs `ComputeReferences=true` geometry and a face reference.
22. **A manufacturer family's real geometry usually lives in NESTED families whose
    parameters are Type-level, which makes them undrivable from the host.** On the Titus ML
    family, the plenum (including the inlet collar) is a nested `ML-Plenum` family exposing
    `Oval Length`, `Oval Height`, `Oval Radius`, `Standard Inlet Size` — but all as TYPE
    parameters, and only INSTANCE parameters of a nested family can be associated to host
    family parameters. The nested instance had exactly one association available
    (`MP-Visibility`). So the host can flip the connector to round (gotcha 21) but CANNOT
    make the visible collar round, because that needs the nested family edited — and
    `Document.EditFamily()` cannot run through `send_code_to_revit` (gotcha 5). Fixing the
    visual requires manual UI work: open `ML-Plenum`, convert `Oval Length` / `Oval Height`
    (and likely `Oval Radius` / `Oval Inner Radius`) to instance parameters, reload, then
    associate and drive them from the host. Useful trick if that is ever done: **a flat
    oval whose Length equals its Height is geometrically a circle**, so the collar can be
    made round by driving `Oval Length` = `Oval Height` = `Inlet Diameter` rather than by
    rebuilding any geometry.
23. **Check `Connector Description` formulas before assuming which connector of a
    multi-connector family is the live one.** The Titus slot family has two connectors
    whose descriptions evaluate to "Disabled" vs. the real name depending on plenum mode —
    with Standard Plenum selected, the OVAL connector (`Unconditioned Air`) is live and the
    rectangular one (`Conditioned Air`) is disabled. Titus's naming is counterintuitive:
    `Unconditioned Air` is the actual supply inlet. A sizing formula was initially built
    against `Conditioned Air` purely because the name sounded right, wiring it to the dead
    connector — caught only by reading the description formulas, not by any error. Also
    note `Unconditioned Air` is a SHARED parameter and cannot be renamed
    (`FamilyManager.RenameParameter` throws "Cannot rename the shared parameter").
24. **A manufacturer family can carry TWO independent "length" concepts — setting the one
    that sounds authoritative may drive nothing visible.** On the Titus slot family,
    `Select a Plenum Nominal Size` → `Plenum Nominal Size` → `Actual Diffuser Length`
    drives the PLENUM, while a completely separate static type parameter
    `Input the ML(R)'s Nominal Length` → `Length` → `Length1` → `Diffuser Width` drives the
    actual diffuser BODY geometry. Setting only the first made every Type report a
    different `Actual Diffuser Length` while the placed instance stayed physically 60" long
    on all five types. **Caught only by measuring a placed instance's bounding box in the
    project** — all 135 family-editor checks passed because they exercised the sizing
    formulas (slot count, inlet diameter, Max CFM, warning), none of which touch body
    length. Lesson: after building a family, measure the actual placed geometry, not just
    the parameters you wrote.
25. **A formula on a Type parameter that another Type parameter's formula references can
    silently fail to propagate.** Setting `Input the ML(R)'s Nominal Length = Plenum
    Nominal Size` made `Input` read correctly (72" on the 72" type) while `Length`, whose
    Titus-authored formula is literally `[Input the ML(R)'s Nominal Length]`, stayed stuck
    at its old 60". Not a reporting parameter (`IsReporting=False`), `CanAssignFormula=True`,
    and repeated `document.Regenerate()` did not resolve it. Fix that worked reliably:
    clear the formula and set explicit per-type VALUES instead (24/36/48/60/72), after
    which `Length`/`Length1`/`Diffuser Width` all tracked correctly. Prefer per-type static
    values over formula indirection when driving a manufacturer family's existing geometry
    chain, and always read back the DOWNSTREAM parameter to confirm propagation — reading
    back only the parameter you set would have shown a false success here.

**Verification standard:** every rebuilt formula needs regenerate-and-read boundary
testing (`FamilyManager.Set()` → `document.Regenerate()` → read back via
`FamilyType.AsDouble()`) at every breakpoint, both sides. Spot-checking a handful of
entries is not sufficient — a real 4-of-32-entry data bug on the 272RL table was only
caught by a full systematic re-derivation, and passed 6 earlier spot-checks by chance.

---

## 5. Suggested order of operations

1. ~~Colin finishes the Ceiling Supply Diffuser worked example~~ — DONE. Its final
   parameter/formula shape is the template for every other family below.
2. ~~Resolve Open Items 1-2~~ — DONE 2026-08-26 (Section 2).
3. ~~Verify Item 5 (Sidewall Exhaust Grate ratio-cap status)~~ — DONE 2026-08-26,
   confirmed safe as a copy source (Section 2).
4. Apply the friction fix (Section 2, Item 1) to the three remaining families that need
   it: Ceiling Return Grate, Ceiling Exhaust Grate, Sidewall Return Grate, Sidewall
   Exhaust Grate. Ceiling Supply Diffuser is already fixed; Sidewall Supply Diffuser
   needs no fix.
5. Rebuild in Revit 2022, recommended order: ~~Ceiling Supply Diffuser~~ (DONE
   2026-08-26 — `Duct Diameter`/`Sizing Warning` formulas added with friction-corrected
   breakpoints, 20/20 boundary tests pass; live-verified in a real placed instance) →
   ~~Ceiling Return Grate~~ (DONE 2026-08-26 — actual .rfa name on disk is "RJA - Return
   Register - Rectangular - Round Neck"; friction-corrected 6"→86/8"→187 CFM breakpoints,
   36/36 boundary + Sizing Warning tests pass; first hand-typed formula attempt hit a
   paren-count `ArgumentException`, fixed by building the nested-if string
   programmatically from the breakpoint array instead — see gotcha 18) → ~~Sidewall
   Supply Diffuser~~ (DONE 2026-08-26 — full Regular/Thin/Custom rebuild, no friction fix
   needed, 128/128 checks pass, live-verified with all 3 types in a real floor plan
   instance) → ~~Ceiling Exhaust Grate~~ (DONE 2026-08-26 — actual .rfa name on disk is
   "RJA - Exhaust Grille - Rectangular - Round Neck"; byte-identical formulas to Ceiling
   Return Grate pushed directly, 36/36 tests pass; round-neck ceiling trio now complete)
   → ~~Sidewall Return Grate~~ (DONE 2026-08-26 — actual .rfa name on disk is "RJA -
   Return Register - Sidewall"; friction fix applied, 6x6 cap 114→110 CFM; full
   Regular/Thin/Custom rebuild, 128/128 checks pass; live-verified across all 3 types in
   the floor plan) → ~~Sidewall Exhaust Grate~~ (DONE 2026-08-26 — actual .rfa name on
   disk is "RJA - Exhaust Grille - Sidewall"; byte-identical formulas to Sidewall Return
   Grate, 128/128 checks pass — **sidewall trio now complete**) → Slot Diffuser (last
   remaining family, most involved, no existing template to copy from).

   **Slot Diffuser (SUPPLY) built 2026-08-26** from the Titus ML native family (stripped
   down rather than built blank, per Colin). 135/135 verification checks pass. See Item 6
   for the full design record. Two things still outstanding on it: the visual collar
   geometry is still flat-oval (gotcha 22 — needs manual nested-family work), and it has
   not been saved to the RJA library under an RJA name yet.

   **Status as of 2026-08-26: 6 of 7 families fully rebuilt in Revit 2022 and verified
   (all 6 confirmed live in a real project). Slot Diffuser supply built and verified in
   the family editor. Slot Diffuser RETURN remains — see Item 6b.**
6. Apply the full gotchas checklist (Section 4) to every family, not just the first one.
7. Confirm Open Items 3, 4, 7 with Colin at whatever point they become blocking, not
   necessarily before starting — they don't change the rebuild mechanics, just the final
   numbers.
