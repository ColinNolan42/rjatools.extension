# GRD Auto-Sizing Families — Cross-Examination Plan

**Purpose:** before redoing all Supply/Return/Exhaust diffuser families in Revit 2022,
audit the current (Revit 2026) state of all 6 built families + the 1 unbuilt one against
each other and against the new friction requirement, so the rebuild is a clean,
consistent re-execution — not a copy of unresolved inconsistencies.

**How to use this doc:** this is written for a fresh Claude Code instance with no
memory of the build history. Full history/reasoning lives in Claude Code memory
(`project_smart_diffuser_family.md`, ~1900 lines) — this file is the condensed,
actionable version. Where this doc says "verify live," don't trust the number here;
pull it from the actual open Family Editor session via Revit MCP first.

**Context:** Colin is hand-editing ONE family (24x24 Supply Diffuser, ceiling,
Rectangular - Round Neck) in Revit 2026 as a worked reference before this rebuild
starts. Use that edit as the template for exactly what "done" looks like — parameter
names, formula style, Sizing Warning pattern — before touching the other 5-6 families.

---

## 1. Current state — all families (Revit 2026, all confirmed same product-version metadata issue, all need Revit 2022 rebuild from scratch — Revit cannot open a newer-version family in an older release)

| Tag(s) | Mount | Family name (2026) | Sizing basis | Target | Cutsheet | Ceiling(s) | Ratio caps | Verified |
|---|---|---|---|---|---|---|---|---|
| SD-1/SD-2 | Ceiling, round neck | `RJA - Supply Diffuser - Rectangular - Round Neck (CFM Auto-Sizing)` | NC (TMS only, OMNI dropped) | NC<15 | TMS (`TMSperf_diffusers_2017.pdf`) | 209/327/535 cfm (12x12/18x18/24x24) | n/a (single diameter per Flow) | 16-18 cases |
| SD-4/SD-5 | Sidewall | `RJA - Supply Diffuser - Sidewall (CFM Auto-Sizing)` | NC, +7 deflection correction | NC<20 (raw<13) | 272RL (`aero_supplyperf_aeroblade_2017.pdf`) | 1500 cfm (24x24/48x12 tier) | Thin ≤4:1, Regular best-available (2:1 not always achievable) | 34-36 cases |
| RG-1 | Ceiling, round neck | `RJA - Return Grate - Rectangular - Round Neck (CFM Auto-Sizing)` | **Pure 600 FPM velocity — NC abandoned entirely** | n/a | Perforated Return (PAR/PXP/PMR), velocity math cross-validated against it | 327/838/1584 cfm (12x12/18x18/24x24, caps = largest diameter that physically fits each face) | n/a | 17-18 cases |
| RG-2 | Sidewall | `RJA - Return Grate - Sidewall (CFM Auto-Sizing)` | NC, no deflection correction (fixed-blade) | NC<15 | 350R/350F/350R-SS (`350.355Rperf_300_350_2017.pdf`) | 1500 cfm (same tier as supply) | Thin ≤4:1, Regular best-available | 36 cases |
| EG-1 | Ceiling, round neck | `RJA - Exhaust Grille - Rectangular - Round Neck` | Identical to RG-1 (byte-identical formulas) | n/a | Same as RG-1 | Same as RG-1 | n/a | 11 cases + re-confirmed live 2026-07-24 |
| EG-2 | Sidewall | `RJA - Exhaust Grille - Sidewall` | NC, reused RG-2's breakpoint array verbatim (23R rejected — half the capacity of 350R at same NC) | NC<15 (inherited from RG-2) | 350R/355R (same as RG-2) | 1500 cfm | **UNCONFIRMED whether the 4:1/2:1 ratio-cap pass was applied — see Open Item 5** | 38 cases |
| SD-3 / RG-3 | Slot | Shared family, name TBD | **Not built** | Not decided | ML-38 (`MLperf_diffusers_2017.pdf`, section starts pdftotext line ~59, not the ML-37 section at top) | n/a | n/a | 0 |

All 7 tags trace back to `Downloads\Schedules_GRD.xls` (GRD-TITUS SCHEDULE sheet).

---

## 2. Open items to resolve BEFORE or DURING the rebuild

### 1. Friction/pressure-drop is not checked anywhere — this is the reason this cross-examination exists
Confirmed 2026-08-26: a family can pick a neck size that satisfies its own NC/capacity
data (e.g. 6" at 125 CFM, well under any NC breakpoint) while failing RJA's own duct
friction standard used by the Duct Velocity tool (`lib/hvac_graph.py`,
`FIRM_DEFAULTS`: Supply Air 0.08 in.wc/100ft, Return/Exhaust Air 0.05 in.wc/100ft).
None of the 6 built families cross-check against this at all.

**What to do:** for every round-neck breakpoint in SD-1/SD-2, RG-1, EG-1 (the three
families where `Duct Diameter` unambiguously maps to a real duct segment), recompute:
```
area_ft2 = pi * (diameter_in / 24)^2
fpm = cfm / area_ft2
friction = 6.82e-6 * fpm^1.82 / diameter_in^1.22      # matches duct_friction_loss_per_100ft()
```
and confirm `friction <= 0.08` (Supply) or `<= 0.05` (Return/Exhaust) at every existing
breakpoint. Where a breakpoint fails, the cascade needs to step up to the next diameter
at that CFM — this will make some families' effective capacity per size SMALLER than
what's currently built, not bigger. Re-verify with the same regenerate-and-read boundary
testing standard as everything else in this project (see Section 4).

**Open question, don't assume an answer:** do the SIDEWALL families (SD-4/SD-5, RG-2,
EG-2) need this too? Their `Duct Width`/`Duct Height` is the grille FACE opening, not
necessarily a 1:1 stand-in for an actual branch duct cross-section the way a round neck
is. Ask Colin whether sidewall grilles have an associated branch duct size that should
also get this check, or whether friction only applies to the round-neck ceiling families.

### 2. SD-4's model was corrected to S301FL on 2026-07-23 but the correction was never actually implemented
The authoritative GRD schedule table (pasted by Colin 2026-07-23) shows **SD-4 = Titus
S301FL**, a different model from SD-5's 272RL. This was flagged explicitly at the time
("SD-4 needs its own S301FL cutsheet and treatment — do not keep assuming SD-4 = SD-5's
family"). Every session after that correction still built and refers to "SD-4/SD-5" as
one combined item sharing 272RL data with no S301FL cutsheet ever pulled.

**What to do:** ask Colin directly whether SD-4 sharing SD-5's family/data was an
intentional final decision (maybe S301FL turned out equivalent enough, or SD-4 got
dropped/merged on purpose) or whether this is a genuine unresolved gap that needs an
actual S301FL cutsheet pull before the Revit 2022 rebuild. Don't silently carry
"SD-4/SD-5 together" forward as settled fact — the memory record shows it was flagged
and then never revisited.

### 3. Ceiling families abandoned NC entirely; sidewall families still use it — confirm this split is intentional
RG-1/EG-1 use pure 600 FPM velocity sizing (NC data was reviewed and explicitly
rejected by Colin: "revisit the cutsheet use the 600 fpm disregard the nc values").
SD-1/SD-2, SD-4/SD-5, RG-2, EG-2 all still use NC-driven cascades. This is a real
methodological split between mounting types, made through genuine back-and-forth
decisions, not obviously a mistake — but it's worth a single explicit confirmation
before locking it into 6 rebuilt families rather than assuming the reasoning that led
to RG-1/EG-1's velocity-only approach shouldn't also apply elsewhere (e.g. should
SD-1/SD-2 also drop NC in favor of velocity, now that there's a precedent?).

### 4. RG-2's max-CFM ceiling increase is still open — oldest unresolved item in the project
Colin asked for RG-2's max CFM to be raised at some point mid-project; this got
sidetracked into the NC-retarget work and was never given a target number. Still capped
at 1500 CFM (24x24/48x12 tier) like every other sidewall family. Get an actual number
before or during the rebuild, or confirm it's fine to leave at 1500.

### 5. EG-2's ratio-cap status is ambiguous — verify live, don't assume
The 4:1 Thin / 2:1-best-effort Regular ratio caps were explicitly applied to RG-2 and
SD-4/SD-5 on 2026-07-24 ("Applied to RG-2 (return) live formula — DONE. Applied to
SD-4/SD-5 (supply) live formula — DONE"). EG-2's build session is filed AFTER that in
the memory doc but reused RG-2's breakpoint array "verbatim" — it's not clear whether
that means before-cap or after-cap RG-2 data. Before treating EG-2 as a reference during
the rebuild, pull its actual live `Auto Width`/`Auto Height` formulas via Revit MCP and
compare the Thin-side picks against RG-2's post-cap array (Section 3 breakpoint tables
below) to confirm which version it actually has.

### 6. Slot (SD-3/RG-3) is completely unbuilt — the largest remaining scope item
Never started. Needs: (a) the ML-38 cutsheet's supply table AND confirmation of whether
a return-mode table exists on the same sheet or if return reuses supply data (same
question already answered once for RG-2/EG-2 — check first, don't assume), (b) more
slot widths/sizes than the current family has, (c) supply vs. return switching logic
that no other family in this project needs (every other tag is single-purpose). Flagged
repeatedly as "the most architecturally involved remaining piece" — budget accordingly,
likely not a fast mechanical rebuild like the other 6.

### 7. NC target summary — confirm this table is still what Colin wants before rebuilding all 6
| Family | Target |
|---|---|
| SD-1/SD-2 (Supply Ceiling) | NC<15 |
| SD-4/SD-5 (Supply Sidewall) | NC<20 effective (raw<13 after +7 correction) |
| RG-1/EG-1 (Return/Exhaust Ceiling) | No NC — pure 600 FPM |
| RG-2/EG-2 (Return/Exhaust Sidewall) | NC<15 |

---

## 3. Reference data (copy verbatim into new families — do not re-derive)

Full breakpoint arrays, formulas, and lookup files already exist and are correct as of
their last-verified state — pull them from:
- `GRD Sizing/lookup_data/SD-1_TMS_supply_ceiling.csv`
- `GRD Sizing/lookup_data/SD-4_SD-5_272RL_sidewall.json` / `.csv`
- `GRD Sizing/lookup_data/RG-1_perforated_ceiling_return.json` / `.csv`
- `GRD Sizing/lookup_data/RG-2_350R_sidewall_return.json` / `.csv`
- `GRD Sizing/lookup_data/EG-1_ceiling_exhaust_round_neck.json` / `.csv`
- `GRD Sizing/lookup_data/EG-2_350R_sidewall_exhaust.json` / `.csv`

These are the current source of truth for breakpoints. They do NOT yet include the
friction cross-check from Open Item 1 — that has to be layered on top before the Revit
2022 build, not read as final without it.

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

**Data verification discipline:**
8. When re-deriving breakpoints for a new NC target, the sort order can change (a
   corrected/retargeted value can move to a different position in the ascending
   cascade) — always re-sort and rebuild the full cascade, never patch values in place.
9. Compute expected test values PROGRAMMATICALLY from the same breakpoint array used to
   build the formula, never hand-type them — this eliminated an entire class of false
   verification failures once adopted (proven across 3+ separate rebuild sessions).
10. A quick visual re-read of a dense multi-column Titus table that conflicts with a
    previously-verified number is not, by itself, evidence the old number is wrong —
    the "miscounted leading dash cells" failure mode is real and repeats. Re-verify via
    precise `page.get_text('words')` coordinate extraction before changing anything.
11. Check each cutsheet's own notes page for its dash convention (has been both "<10"
    and "<15" on different sheets in this project) and deflection-correction table (272RL
    needs +7 NC for 45° deflection; 23R's table already states 45° directly, no
    correction; confirm per-sheet, never assume it carries over from a different sheet).
12. Validate any hand-edited JSON with `python -c "import json; json.load(open(...))"`
    immediately after editing — a partial Edit call has corrupted a JSON file at least
    once on this project by not closing an object before an array bracket.
13. Never assume two tags sharing a Titus model also share one Revit family, or that a
    "shared family" from an earlier note is still accurate — confirm live via Family
    Editor inspection every time. RG-1/EG-1 share a model but are separate families;
    SD-4 was assumed to share SD-5's family and may not (Open Item 2).
14. `effective_height_in()` in `lib/hvac_graph.py` snaps duct/neck dimensions to the
    nearest 2" nominal increment before comparison — if any NEW code in this project
    reads a raw `RBS_CURVE_*` dimension directly for a comparison (not through that
    function), apply the same snap; Revit's internal feet↔inch round-trip can store a
    true 10" as `10.000000000000002"`, which is exactly what caused a false Duct
    Velocity flag earlier this session (see `hvac_graph._nominal_even_in()`).

**Verification standard:** every rebuilt formula needs regenerate-and-read boundary
testing (`FamilyManager.Set()` → `document.Regenerate()` → read back via
`FamilyType.AsDouble()`) at every breakpoint, both sides. Spot-checking a handful of
entries is not sufficient — a real 4-of-32-entry data bug on the 272RL table was only
caught by a full systematic re-derivation, and passed 6 earlier spot-checks by chance.

---

## 5. Suggested order of operations

1. Colin finishes the 24x24 Supply Diffuser worked example in Revit 2026 — use its
   final parameter/formula shape as the literal template for every other family below.
2. Resolve Open Items 1-2 (friction check scope, SD-4/S301FL) with Colin before building
   anything — both change what data goes into the rebuild, not just how it's verified.
3. Verify Open Item 5 (EG-2 ratio-cap status) live against the current 2026 family before
   it gets used as a copy-source.
4. Rebuild in Revit 2022, recommended order: SD-1/SD-2 → RG-1 → EG-1 (round-neck ceiling
   trio, share the same architecture and now the same friction-check addition) → SD-4/SD-5
   → RG-2 → EG-2 (sidewall trio, share the Regular/Thin/Custom architecture) → SD-3/RG-3
   slot (last, most involved, no existing template to copy from).
5. Apply the full gotchas checklist (Section 4) to every family, not just the first one.
6. Confirm Open Items 3, 4, 7 with Colin at whatever point they become blocking, not
   necessarily before starting — they don't change the rebuild mechanics, just the final
   numbers.
