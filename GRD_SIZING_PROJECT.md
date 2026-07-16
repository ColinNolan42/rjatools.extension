# Smart Diffuser Family — Project Explainer

**Status: ongoing, planning stage.** No pushbutton/script exists yet — this
is base-of-design and family-mapping work in progress. Not gated under the
Phase 1/2/3 pipe sizing rules in `CLAUDE.md`; this is a separate, unrelated
HVAC tool (like the Duct Velocity project) and has no phase lock.

## Goal

An "active" loadable Revit diffuser family (supply, return, exhaust) that,
given a CFM instance parameter, auto-sizes to the correct nominal/neck size
per RJA's base of design — using a family lookup table (`LOOKUP()` +
CSV keyed by CFM) rather than nested IF formulas, so the table stays easy
to update if the base of design changes.

## Scope — 7 active tags

Per RJA's standard GRD (Grille/Register/Diffuser) schedule
(`C:\Users\Colin Nolan\Downloads\Schedules_GRD.xls`, sheets
"GRD-TITUS SCHEDULE" and "GRD-PRICE SCHEDULE" — same tag/type structure,
Titus model numbers populated, Price column blank):

| Tag | Type | Titus Model | Max NC | Max ADP (in.wc) | RJA Revit Family (loaded, confirmed) |
|---|---|---|---|---|---|
| SD-1 | Ceiling | TMS | 30 | 0.08 | RJA - Supply Diffuser - Rectangular - Round Neck |
| SD-2 | Ceiling | OMNI | 30 | 0.08 | RJA - Supply Diffuser - Rectangular - Round Neck |
| SD-3 | Slot | ML-38 | 30 | 0.08 | RJA - Supply Diffuser - Linear Slot - Round Neck |
| SD-4 | Spiral duct (boot-connected, shares SD-5's family) | 272 RL | 30 | 0.08 | RJA - Supply Diffuser - Sidewall |
| SD-5 | Wall | 272 RL | 30 | 0.08 | RJA - Supply Diffuser - Sidewall |
| RG-1 | Ceiling | 50F | 25 | 0.06 | RJA - Return Register - Rectangular - Round Neck |
| RG-2 | Wall | 23 RL | 25 | 0.06 | RJA - Return Register - Sidewall |
| EG-1 | Ceiling | 50F | 25 | 0.06 | RJA - Exhaust Grille - Rectangular - Round Neck |
| EG-2 | Wall | 23 RL | 25 | 0.06 | RJA - Exhaust Grille - Sidewall |

**Explicitly out of scope** — SD-6 (ceiling/VAV) and TG-1/TG-2 (transfer,
ceiling and wall) were dropped by the user. Transfer grilles specifically
have no CFM value in the schedule, and this whole family's premise is
auto-sizing off an input CFM — a tag with no CFM has nothing for the
lookup table to key off.

## Base of design: Titus (not Price)

Base of design was originally scoped against the Price 500/600/700
Louvered Grille cutsheet, then switched to Titus per the schedule the user
actually provided. All 6 distinct Titus performance-data cutsheets (several
tags share one model) have been pulled and content-verified — confirmed
each PDF actually contains CFM/NC tables, not just dimensional data (a few
Titus "submittal" PDFs with plausible filenames turned out to be
dimensions-only decoys, e.g. `272r.pdf` and `23r.pdf`):

| Tag(s) | Titus Model | Cutsheet URL |
|---|---|---|
| SD-1 | TMS | https://www.titus-hvac.com/file/871/TMSperf_diffusers_2017.pdf |
| SD-2 | OMNI | https://www.titus-hvac.com/file/890/OMNIperf_diffusers_2017.pdf |
| SD-3 | ML-38 | https://www.titus-hvac.com/file/803/MLperf_diffusers_2017.pdf |
| SD-4/SD-5 | 272 RL | https://www.titus-hvac.com/file/981/aero_supplyperf_aeroblade_2017.pdf (covers 271/272/111/112/121/122/131/132 models together) |
| RG-1/EG-1 | 50F | https://www.titus-hvac.com/file/791/50F_50Rrperf_specialized_2017.pdf |
| RG-2/EG-2 | 23 RL | https://www.titus-hvac.com/file/994/23Rperf_aeroblade_2017.pdf |

Direct `WebFetch` to titus-hvac.com returns 403 — download via `curl` with
a browser `User-Agent` header instead (same issue affected the earlier
Price cutsheet fetch).

## Design criteria — one item still open

Blade config: **45° deflection for both supply and return**, confirmed.

NC target: **conflict, not yet resolved.** User initially said "NC 30
across the board" before seeing the schedule. The schedule itself sets
**NC 30 for supply (SD tags) but NC 25 for return/exhaust
(RG/EG tags)**. Needs a decision before the RG/EG lookup tables can be
built: honor the schedule's NC 25, or keep NC 30 uniformly per the user's
earlier instruction.

Max ADP (pressure drop): the schedule lists 0.08 in.wc (supply) / 0.06
in.wc (return/exhaust) as a second constraint alongside NC — not yet
decided whether this should be a hard cap in the sizing logic (i.e. reject
a size if it fails ADP even though NC passes).

ML-38 slot configuration: "38" likely denotes 3/8" slot width in Titus's
naming — not yet cross-checked against the schedule's neck-size annotation
format (`LENGTH-#SLOTS-WIDTH`, e.g. `48-1`) to confirm which slot-width
column of the ML cutsheet applies.

## Next steps

1. Resolve the NC 25 vs NC 30 conflict for RG/EG with the user.
2. Resolve whether Max ADP is a hard sizing cap.
3. Confirm ML-38 slot width against the schedule's neck-size format.
4. Build lookup table CSVs per tag from the 6 Titus cutsheets above.
5. Build/verify family formula structure (`LOOKUP()`) against each RJA
   family's existing parameter set — don't invent new parameter names if
   the loaded families already have equivalents.
6. Test in Revit (family editor, then a live model) before rollout.

## Related memory

Full working notes (including troubleshooting details, e.g. the Revit MCP
`say_hello` false-alarm timeout caused by a blocking dialog, not a
connection failure) live in this session's auto-memory under
`project_smart_diffuser_family.md`.
