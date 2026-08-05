# R23 Codex Handover

**Date:** 2026-08-01
**Handing off from:** Claude Code (this session)
**Repo:** `solidworks-vba-3d-to-2d` (github.com/cochilocovt/solidworks-vba-3d-to-2d)
**HEAD at handover:** `e9783e5` (local main == origin/main, in sync)

## Read these first, in order

1. [`Agents.md`](../Agents.md) — binding operating contract: fixture
   authorization, ask-before-live-SOLIDWORKS rule, VBA engineering rules,
   acceptance criteria. Non-negotiable.
2. [`docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md`](R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md)
   — the standing implementation handoff (evidence ladder, claim language,
   commands). This document supplements it; it does not replace it.
3. [`docs/R23_IMPLEMENTATION_PLAN.md`](R23_IMPLEMENTATION_PLAN.md) — the
   full Phase 0–12 task list (R23-xxx numbering) with per-task status.
4. [`docs/CURRENT_STATUS.md`](CURRENT_STATUS.md) — latest state; **read the
   actual `MACRO_SOURCE_REVISION`** in
   `src/target-spec-hybrid-v2/Module1_Main.bas`, never trust a doc's
   revision label over the source.
5. [`skills/solidworks-api-lookup/SKILL.md`](../skills/solidworks-api-lookup/SKILL.md)
   — mandatory before touching any `sw*` constant or API call. The project
   has an MCP server (`solidworks-api`) serving a SOLIDWORKS 2025 API corpus;
   query it, don't recall values from memory. Constants have been silently
   wrong before (see skill file for the incident).

## What this project is

A VBA macro (`src/target-spec-hybrid-v2/`) that automates 3D-to-2D drawing
generation in SOLIDWORKS 2025 SP1.2: imports model dimensions/annotations
into a drawing, builds ordinate dimension schemes, creates manufacturing
callouts, cuts section views, dimensions sections, and lays out the sheet —
all against a single reference part, `P-0251-14A-001.SLDPRT`, validated
against a hand-built reference drawing the designer already approved.

R23 is a ground-up rebuild (Phases 0–12) replacing an earlier
"active-ordinate" implementation that had unprovable behavior. The
governing idea across all of R23: **every claim must be backed by evidence
from a live SOLIDWORKS run** (Immediate Window `Debug.Print` output), not by
static source review or compile success alone. See the evidence-ladder
section of the implementation handoff doc for the exact vocabulary
(`static` / `API evidence` / `SWP deployment` / `VBE compilation` /
`live execution` / `screenshot review` / `manufacturing acceptance`).

## Standing constraints (from the user, still in force — do not relax)

- Do not modify `src/baseline-model-dims/`.
- Do not modify or save any fixture model.
- Do not overwrite manual reference drawings.
- Only the three fixtures authorized by `Agents.md` may be used.
- **The user performs all live SOLIDWORKS and VBA-editor work.** Never
  attempt Computer Use against SOLIDWORKS. Give exact manual deployment,
  compilation, and execution steps and wait for the user to paste back the
  Immediate Window output.
- Before relying on any SOLIDWORKS API member, enum, selection contract,
  coordinate transform, or return code: use the `solidworks-api-lookup`
  skill / MCP first.
- Preserve all existing work — never discard, overwrite, rename, or
  silently supersede uncommitted changes.
- Distinguish static verification from live proof in every report. Static
  success (compiles, contract tests pass, preflight passes) is never
  reported as runtime proof.
- **Tolerance policy (verbatim from the user):** "For now ignore the added
  tolerance in general for all the drawings the designer made as those
  values were manually added to represent that some tolerance is fine
  rather than the part having it. I shall get back to you after discussing
  this with my designer sometime later regarding what info from the part
  can we use to determine whether a tolerance is to be added or not." — one
  exception is already resolved: see Phase 8 H7 note below.

## Architecture — active component map

All source lives in `src/target-spec-hybrid-v2/`. Every phase below is a
`ModuleNN_Name.bas` (+ optional `C*.cls` typed record) with **exactly one
mutation boundary**: a `Public Function`/`Sub` that refuses unless called
with an explicit `allowMutation:=True` argument, plus a strictly read-only
`R23_Probe*` entry point that a Debug > Run can call directly and that never
calls the production QA gate (`Module6_QAEngine.EmitRunEvidence`).

| Module | Phase | Purpose | Mutation entry point | Probe entry point |
|---|---|---|---|---|
| `Module11_GeometryIdentity.bas` | 1 | Canonical geometry identity, COM Boolean normalization (`NormalizeSwBoolean`) | — (pure functions) | — |
| `Module12_FeatureQualification.bas` | 2 | Builds the feature catalog / location graph from the model | — (read-only module) | `BuildFeatureCatalog` (called by others) |
| `Module13_ProjectionResolution.bas` | 3 | Resolves model features to per-view page-frame projections | — | `BuildViewProjections` |
| `Module14_AnnotationImport.bas` | 4 | Imports/reconciles drawing annotations against model locations | — | (probe TBD, see plan) |
| `Module15_OrdinateScheme.bas` + `COrdinateScheme.cls`/`COrdinateBucket.cls` | 5 | Ordinate dimension grouping/crediting, coincident-projection handling | ordinate creation (gated) | `R23_Probe*` |
| `Module16_CalloutDefinition.bas` + `CCalloutDefinition.cls` | 6 | Manufacturing callout definitions (hole depth, thread, counterbore) | native/fallback callout creation (gated) | `R23_Probe*` |
| `Module17_SectionPath.bas` + `CSectionPath.cls` | 7 | J-J section-line path from model intent (bore + face-hole grid) | `CreateSectionFromPath` (gated) | `R23_ProbeSectionPath` |
| `Module10_SectionDimensionEngine.bas` + `CSectionRequirement.cls` | 8 | Semantic section-dimension engine (7 P-0251 requirement keys, H7 reference-fit application) | `CreateSectionDimension`, `ApplyReferenceFit` (both gated) | `R23_ProbeSectionDimensions` |
| `Module18_ContentEnvelope.bas` + `CContentEnvelope.cls` | 9 | Content-envelope-aware final layout (replaces fixed upward-bias placement) | `ApplyPlacementPlan` (gated) | `R23_ProbeContentEnvelope` |

Legacy modules still on the **reachable production path** (not yet
replaced — R23 modules are written but not wired into `main`):

- `Module2_DrawingPipeline.bas` — still contains the disproved
  outline-percentage J-J placement (`extension`, `topY + extension`,
  lines ~1525-1556). R23-704/R23-609 will remove this once Module17/18 are
  wired in.
- `Module7_TitleBlockEngine.bas` (lines 359-361) — still writes the legacy
  free-text `"...47 H7 (+0.025/+0.000)"` bore callout. R23-810 detects this
  (`Module10_SectionDimensionEngine.DetectLegacyBoreCallout`) but does not
  remove it yet — removal is blocked on real section dimensions existing
  first.
- `Module9_LayoutEngine.bas` (lines 442-446) — the fixed "bias the P-0251
  source row upward" placement that Phase 9's `Module18_ContentEnvelope`
  is designed to replace, not yet switched over.

Deployment/verification tooling (`tools/`):

- `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1` — guarded deployment into
  `Fable.swp`: preflight, nine gates, readback verification.
  `deployment-manifest.json` currently lists **35 managed components**.
  `deployment-request.txt` is **regenerated from source**, never hand-edit.
- `tools/solidworks-automation-companion/` — separately versioned,
  gitignored offline verification layer. `unittest` suite, currently
  **334 tests** (5 known-stale R20 failures — pre-existing, unrelated to
  R23, documented in the implementation handoff §5.2). Contains all the
  `test_r23_*_contracts.py` static source-contract suites, one per phase.

## Progress to date

Phases 0–9 have committed VBA source with passing static contract suites.
**Phases 0–7 have live-run evidence from the user's SOLIDWORKS session and
their read-only gates are SATISFIED.** Phases 8 and 9 have source complete
and contracts passing, but their first live run (just received) surfaced
real defects — see "Unresolved bugs" below, **not yet fixed**.

| Phase | Status | Evidence |
|---|---|---|
| 0 | Closed | Probe findings recorded in `docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md` |
| 1–4 | Read-only gates satisfied (live) | Recorded in plan/status docs |
| 5 (ordinate scheme) | Read-only gate SATISFIED (live) | `creditedLocations=10\|expectedLocations=10\|coverageFailures=None` |
| 6 (callout definition) | Read-only gate SATISFIED (live) | `definitionFailures=None\|counterboredFamilies=1\|threadedFamilies=1` |
| 7 (section path) | Read-only gate SATISFIED (live) | `resolvedPaths=1\|crossingsProven=4\|sectionFailures=None`; frame transform cross-checked against model geometry (bore Plücker moment) |
| 8 (section dimensions) | **Source complete, first live run FAILED gate** | See bug 1/2 below — `satisfied=0\|missing=7` |
| 9 (content envelope) | **Source complete, probe crashed before producing evidence** | See bug 3 below — `R23_ENVELOPE_FATAL\|reason=SheetRegionsUnmeasured` |

Commits (most recent first, all on `main`, pushed to `origin/main`):

```
e9783e5 Merge origin/main (trivial PR merge commit, no content changes)
3a18667 r23: Phase 9 content-envelope-aware final layout
828e61e r23: Phase 8 semantic section-dimension engine
9d54f49 r23: Phase 7 read-only gate satisfied on P-0251
60723c8 r23: Phase 7 J-J section path from model intent (source)
d2cb203 r23: Phase 6 read-only gate satisfied on P-0251
```

Local `main` and `origin/main` are fully in sync (0 ahead / 0 behind) as of
this handover. One uncommitted local change exists:
`tools/swp-deploy/deployment-request.txt` — this file is machine-regenerated
from `Module1_Main.bas`'s `MACRO_SOURCE_REVISION`; never hand-edit it,
regenerate it via the deploy script if it looks stale.

## Unresolved bugs (found in the just-received live run, NOT yet fixed)

These came from the user pasting Immediate Window output for
`Module10_SectionDimensionEngine.R23_ProbeSectionDimensions` and
`Module18_ContentEnvelope.R23_ProbeContentEnvelope` run back-to-back against
the live P-0251 drawing. **The user has not yet authorized fixes — confirm
before changing source.**

### Bug 1 — Phase 8: nominal read route fails for drawing-authored reference dimensions (blocks the whole gate)

All 7 section dimensions in `Section View J-J` (`RD1`..`RD7@Drawing View6`)
report `nominalAvailable=False`, so every one of the 7 P-0251 requirements
reports `matched=False|missing`. The dimension objects themselves are not
Nothing — `ReadDimensionTolerance` successfully read real
`toleranceType`/`fitType` values off the same `IDimension` object in the
same iteration — so this is not a null-object bug.

**Root cause hypothesis:** `TryReadNominal` in
`Module10_SectionDimensionEngine.bas` calls
`dimension.GetSystemValue3(swThisConfiguration, Empty)`. `RD1`..`RD7` are
**drawing-authored reference dimensions** (name prefix `RD`), which may not
belong to a part configuration the way the Phase 0 probe's target
(`D1@Sketch4`, an *imported model* dimension) did. The config-scoped read
route that worked for Phase 0's evidence may not apply here.

**Suggested fix direction:** Add a fallback read route —
`IDimension.SystemValue` (property) or `GetSystemValue2("")` — attempted
after `GetSystemValue3` fails, with whichever route actually succeeds
named explicitly in the evidence string (do not silently prefer one). Needs
an MCP lookup (`solidworks-api-lookup` skill) on `IDimension.SystemValue`
and `GetSystemValue2` before implementing — do not guess the contract.

### Bug 2 — Phase 8: R23-804's "type 6 only" diameter assumption is stale for this drawing

All 7 section dimensions report `type2=2` (`swLinearDimension`). None are
`type2=6` (`swDiameterDimension`) or `type2=15`
(`swDiametricLinearDimension`). Phase 0's evidence (17 dimensions in
`Drawing View3`, including type-6 `Ø47`/`Ø40`) was from a *different* view
(`Drawing View3`, whole-model import) than what Phase 8 actually reads now
(`Section View J-J`, 7 dimensions) — the drawing content differs by view.
`BuildSectionRequirements`'s two diameter requirements
(`REQ_INNER_BORE`, `REQ_FIT_BORE`) only accept types `6;15`, so they can
never match here even once Bug 1 is fixed.

**Corroborating positive finding (keep this — it's real evidence, not a
bug):** `RD4@Drawing View6` reports
`toleranceType=8|fitType=0|holeFit=H7|minimumStatus=0|minimumM=0.000000|maximumStatus=0|maximumM=0.000025`
— i.e., **H7 +0.025/0.000 already exists on this drawing reference
dimension**, exactly matching the reference specification, with 2 attached
edges (associative). This corroborates the earlier finding (see
`docs/R23_IMPLEMENTATION_PLAN.md` R23-806, resolved 2026-07-31) that H7 is
model-absent and drawing-authored — consistent with the standing tolerance
policy above.

**Suggested fix direction:** Use `IDisplayDimension.Diametric` to detect a
diameter-displaying *linear* dimension (a linear dimension can still show a
diameter value) rather than widening `AcceptedTypes` blindly to include
type 2. Needs an MCP lookup on `IDisplayDimension.Diametric` before
implementing.

### Bug 3 — Phase 8: evidence lines duplicated in Immediate Window output

Every `QA INFO: SECTION_REQUIREMENT` line appears twice in the pasted
output — once without a `view=` field, once with it. Root cause:
`CRunEvidence.AddInfo` already prefixes and prints `QA INFO:` internally,
and `Module10_SectionDimensionEngine.R23_ProbeSectionDimensions`'s explicit
`Debug.Print "QA INFO: SECTION_REQUIREMENT|..."` loop duplicates it.
Cosmetic only, but doubles log volume — low priority, fix alongside bugs 1/2.

### Bug 4 — Phase 9: probe crashes because it depends on a production fail-closed gate

```
QA FAILURE: Controlled sheet has neither an ITitleBlock definition nor a
proved legacy title-block rectangle.
R23_ENVELOPE_FATAL|reason=SheetRegionsUnmeasured
```

`Module18_ContentEnvelope.R23_ProbeContentEnvelope` calls
`Module8_RuntimeSupport.MeasureControlledSheetRegions`, which is a
**production fail-closed gate** written for a sheet the macro itself
creates from the controlled template — it correctly refuses when there's no
`ITitleBlock`/title-block rectangle. The user's live reference drawing has
its sheet format hidden and no `ITitleBlock` (it's a hand-built reference,
not macro-generated), so the gate correctly refuses — and takes the
**read-only probe** down with it. Same failure shape as the earlier Phase 5
`EmitRunEvidence` mistake (a probe should never depend on a stage gate meant
for production mutation runs).

**Suggested fix direction:** Phase 9's envelope-building only needs
`ISheet.GetSize` (width/height) to bound the sheet — it does not need the
title-block/content-border measurements to build view content envelopes.
Restructure `R23_ProbeContentEnvelope` to measure only what's provable
without the full gate, and report each protected-region type as
`NotMeasured` (with a reason) rather than aborting the whole probe when one
region can't be determined.

**Related latent defect (same root cause, not yet triggered but real):**
`Module18_ContentEnvelope.BuildProtectedRegions` reads
`evidence.ContentBorder*`/`TitleBlock*` fields **unconditionally**. If those
stay at their zero-value default (as they do whenever
`MeasureControlledSheetRegions` didn't run or partially failed), it would
silently emit degenerate protected rectangles anchored at the sheet origin
and could report false `ProtectedIntrusion` clearance violations. It
already has one existence gate (`PartIdentificationBoundsProven`) for the
part-ID band — the content-border and title-block regions need the same
explicit "was this actually measured" gating, not an implicit
zero-defaults-to-valid-rectangle assumption.

**Net effect:** Phase 9 has produced **zero live evidence** yet on any of
its actual design questions (display-data frame consistency, inverse
transform round-trip, section-line grammar disambiguation, envelope sizes,
clearance checks, placement plan). The probe needs to actually run to
completion before any of that can be assessed.

## Immediate next steps

1. **Get user authorization**, then fix bugs 1–4 in
   `Module10_SectionDimensionEngine.bas` and `Module18_ContentEnvelope.bas`.
   For bugs 1 and 2, use the `solidworks-api-lookup` skill / MCP first —
   don't guess `IDimension.SystemValue`/`GetSystemValue2` or
   `IDisplayDimension.Diametric` contracts.
2. Add/extend the corresponding contract tests
   (`tools/solidworks-automation-companion/tests/test_r23_section_dimension_contracts.py`
   and `test_r23_content_envelope_contracts.py`) to lock in the fixed
   behavior (e.g., assert the fallback nominal-read route is tried and
   named; assert `Diametric` is consulted; assert unmeasured protected
   regions don't produce degenerate rectangles).
3. Run the full companion suite (`python -m unittest discover -s tests -q`
   from `tools/solidworks-automation-companion/`) — expect the same 5
   known-stale R20 failures and nothing new.
4. Run `Deploy-TargetSpecHybrid.ps1 -PreflightOnly`, then the full deploy.
5. Give the user the manual re-run steps for both probes (VBE compile, run
   `Module10_SectionDimensionEngine.R23_ProbeSectionDimensions` then
   `Module18_ContentEnvelope.R23_ProbeContentEnvelope`), and wait for pasted
   Immediate Window output before claiming either gate satisfied.
6. Once Phase 8 gate is satisfied and Phase 9 actually produces
   `R23_ENVELOPE_END` evidence (not a fatal abort), update
   `docs/R23_IMPLEMENTATION_PLAN.md`, `docs/CURRENT_STATUS.md`,
   `docs/Changelog.md`, and `docs/SOLIDWORKS_API_VALIDATION.md` per the
   per-phase pattern already established in those files, then commit.
7. After Phase 9 closes: the accumulated mutating work table (ordinate
   creation, native callout creation, section-line creation, section
   dimension creation, layout application, plus the two legacy-literal
   removals R23-609/R23-704/R23-810) is still blocked on explicit user
   authorization to mutate the reference drawing plus a non-protected
   target drawing. Do not attempt any of that without asking first.
8. Phase 10 ("Replace count-based QA") has not been started — read its
   task list in `docs/R23_IMPLEMENTATION_PLAN.md` before beginning it.

## Things that will bite you if you skip them

- **Every VBA source file in `src/target-spec-hybrid-v2/` must be
  cp1252/ANSI (all bytes < 0x80), CRLF line endings, no BOM, no `Attribute`
  metadata, `Option Explicit`, max 79 columns.** A UTF-8 em-dash silently
  broke SWP readback once already (see Changelog). Always normalize new/
  edited files the same way the existing ones were normalized before
  committing.
- **Manifest `kind` values are exactly `StdModule` and `ClassModule`** —
  `"StandardModule"` is rejected by the deploy gate and has caused a failed
  deployment before.
- **`R23_Probe*` entry points must never call
  `Module6_QAEngine.EmitRunEvidence`** — it's the production gate and
  demands 14 pipeline stages a read-only probe never runs. This has caused
  a "Failed Closed" dialog before.
- **`ISldWorks.IsSame` returns `swObjectEquality` (0/1/2), never a Boolean**
  — compare `= 1` exactly, never truthy-test the return value.
- **Only `(CDbl(raw) <> 0#)` via `Module11_GeometryIdentity.NormalizeSwBoolean`
  is a representation-independent COM Boolean check.**
- **`IModelDocExtension.GetCorrespondingEntity2` returns Nothing with error
  0 in part drawings** — it declines rather than fails; don't build logic
  that assumes a non-Nothing return is achievable there.
- Never report static/compile/preflight success as runtime proof. The user
  has repeatedly emphasized this distinction and corrected claims that
  blurred it.
