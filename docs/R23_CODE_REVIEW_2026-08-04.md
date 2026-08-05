# R23 code review — changes after 06:16, 2026-08-04

Review date: 2026-08-04
Reviewer: Claude Code (`/code-review high`)
Base: working tree vs `6c00371` (last commit, 2026-08-02 18:27)

## Scope

Files whose mtime is later than 2026-08-04 06:16. Roughly 3,400 changed
lines in tracked sources plus 1,218 lines of new untracked files.

**In scope**

- `src/target-spec-hybrid-v2/` — every `.bas` and `.cls` **except**
  `Module3_ModelAudit.bas`
- `src/target-spec-hybrid-v2/Module20_ProbeRunner.bas` (new, 679 lines)
- `src/target-spec-hybrid-v2/Module21_EvidenceSink.bas` (new, 103 lines)
- `tools/swp-deploy/Invoke-SolidWorksMacro.ps1` (new, 164 lines)
- `tools/swp-deploy/SolidWorksDocumentOpener.cs` (new, 272 lines)

**Out of scope** (modified before 06:16)

- `src/target-spec-hybrid-v2/Module3_ModelAudit.bas` (04:51)
- `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1` (05:01)
- `tools/swp-deploy/deployment-manifest.json` (05:03)
- `tools/probe-runner/Run-R23Probes.ps1` (05:28)
- `docs/R23_CODEX_HANDOVER.md`, `docs/R23_CODEX_INIT_PROMPT.md` (08-01)

Docs and generated files (`deployment-request.txt`, `Changelog.md`,
`CURRENT_STATUS.md`, `SOLIDWORKS_API_VALIDATION.md`) were read for context
but not reviewed for defects.

## Findings

Ten findings, most severe first. Nine CONFIRMED by reading the code; one
PLAUSIBLE, needing a live run to settle.

---

### 1. Import-coverage gate self-satisfies from scheme construction

**File:** [`Module19_SemanticQA.bas:1208`](../src/target-spec-hybrid-v2/Module19_SemanticQA.bas)
**Verdict:** CONFIRMED · **Category:** correctness

`EvaluateSemanticDrawing` builds ordinate schemes (line 1205) before it
reads `CoveredX`/`CoveredY` (line 1208), and the new
`COrdinateBucket.RecordDirectionalCoverage` sets those flags during scheme
construction. The `ViewImportedNothing` gate therefore passes on a view
where nothing was imported and no ordinate was created.

**Failure scenario.** `InsertModelAnnotations4` inserts 0 annotations into
Drawing View2 and `CreateOrdinateGroup` fails for every scheme in it.
`CollectOrdinateSchemes` still calls `BuildSchemesForView`,
`PopulateSchemeLedger` credits buckets, and `COrdinateBucket.cls:60` sets
`projection.CoveredX`/`CoveredY`. `EvaluateModelImportCoverage` then counts
`coveredX > 0` for that view, skips the `ViewImportedNothing` failure, and
marks `STAGE_MODEL_IMPORT_COVERAGE` proved on an empty view.

The probe's previous order ran `EvaluateModelImportCoverage` *before*
`CollectOrdinateSchemes` and did not have this leak. The ordering change
and `RecordDirectionalCoverage` were introduced together in this window.

---

### 2. Section views added twice to the import order list

**File:** [`Module2_DrawingPipeline.bas:689`](../src/target-spec-hybrid-v2/Module2_DrawingPipeline.bas)
**Verdict:** CONFIRMED · **Category:** correctness

`CollectR23ImportViews` adds every eligible section view in loop one
(`swView.Type = 2`, lines 673-681) and adds it again in loop two, which
only excludes the primary view.

**Failure scenario.** Section View J-J is eligible and is not the primary
view, so it lands in `result` twice. `ImportModelAnnotations`
(`Module14_AnnotationImport.bas:283`) has no de-duplication: it activates
J-J and calls `InsertModelAnnotations4` twice, records two
`InsertModelAnnotations4:J-J` mutations, and folds the second call's result
into `evidence.ImportedAnnotations`.

---

### 3. Rescale overwrites approved isometric and detail scales

**File:** [`Module18_ContentEnvelope.bas:1782`](../src/target-spec-hybrid-v2/Module18_ContentEnvelope.bas)
**Verdict:** CONFIRMED · **Category:** correctness

`ApplyScaleToFit` multiplies `ScaleDecimal` on every non-parent-scaled view
in the layout set. That set includes the isometric pinned at its approved
1:2 and the detail views at their approved 3:1. No check confirms those
approved scales survive.

**Failure scenario.** Content does not fit, so `ApplyPlacementPlan`
requests factor 0.8. The isometric was created and verified at exactly
`ScaleDecimal` 0.5 (`Module2_DrawingPipeline.bas:1301` hard-fails
otherwise) and Details C/D at 3.0 (`Module9_LayoutEngine.bas:1475`
`ValidateLayout` hard-fails otherwise). `ApplyScaleToFit` sets them to 0.4
and 2.4. No rescale-policy failure is recorded; the run reports
`rescale=Applied` and `layout=Applied` while the drawing now violates both
approved-scale rules the pipeline enforces elsewhere.

---

### 4. Outline datum accepts arcs and hidden edges

**File:** [`Module15_OrdinateScheme.bas:861`](../src/target-spec-hybrid-v2/Module15_OrdinateScheme.bas)
**Verdict:** CONFIRMED · **Category:** correctness

`OutlineDatumForModelEdge` decides an edge is a horizontal bottom-outline
datum from its two transformed endpoint Y values alone.
`ResolveOutlineDatum` iterates every edge of every solid body returned by
`GetBodies2` (line 690) with no view-visibility filter.

**Failure scenario, arc.** P-0251's bottom face is arc-bounded. A
semicircular edge whose two endpoints share a Y passes the
`Abs(startY - endY)` test at line 861 and reports `horizontalSpanM` equal
to its chord, so it is accepted as "horizontal". The datum is placed at the
endpoint Y even though the arc's true low point sits below it.

**Failure scenario, visibility.** A hidden back-face edge that maps into
the view projects lower than the real silhouette and wins the bottom datum,
so every vertical ordinate is measured from geometry the drawing does not
show.

Secondary concern: `ResolveOutlineDatum` calls
`MapModelEntityToDrawingForDatum` per candidate edge, which can invoke
Route D (`ClearSelection2` + `SelectEntity` + `ClearSelection2`). On a part
with hundreds of edges that is hundreds of selection round trips per
scheme, per view.

---

### 5. Curve-order gate hides an observed mismatch

**File:** [`Module12_FeatureQualification.bas:1319`](../src/target-spec-hybrid-v2/Module12_FeatureQualification.bas)
**Verdict:** CONFIRMED · **Category:** correctness

`ProbeCurveRole` treats a `False` return from `ProbeFeatureCircularEdge` as
"try the next feature". But `False` means the two read orders actually
disagreed on the one circular edge examined —
`ProbeFeatureCircularEdge` exits on the first circular edge it finds either
way.

**Failure scenario.** Two counterbore features exist. The first yields a
circular edge where the R22 order reads radius 0.0055 and the retained
order reads 0.0033, so `CurveOrderResultsMatch` returns `False`,
`ProbeFeatureCircularEdge` logs `status=Fail` and returns `False`. The loop
continues to the second counterbore, which matches, so `ProbeCurveRole`
exits with an empty failure string and `R23_CURVE_ORDER_END` reports
`failures=None`. The gate passes despite a logged live disagreement. When
no feature matches at all, the reported reason is
`NoCircularRepresentative`, which misattributes a real order disagreement
as a missing sample.

---

### 6. `Or` is not short-circuit in VBA; `Nothing` is dereferenced

**File:** [`Module2_DrawingPipeline.bas:564`](../src/target-spec-hybrid-v2/Module2_DrawingPipeline.bas)
**Verdict:** CONFIRMED · **Category:** correctness

```vb
If path Is Nothing Or Not path.Resolved Then
```

VBA evaluates both operands unconditionally, so the guard that exists to
catch `Nothing` is the thing that raises on `Nothing`.

**Failure scenario.** `ISheet.GetViews` returns an array whose bounds make
the `For` loop body never execute, or `ResolveSectionPath` raises before
assigning its return value, leaving `path` as `Nothing`. Line 564 evaluates
`Not path.Resolved`, raises runtime error 91, and control jumps to the
`Failed` handler, which records `R23 semantic section error 91: Object
variable or With block variable not set` instead of the intended
`R23 semantic section path was not resolved.`

---

### 7. Route D clears the user's selection without restoring it

**File:** [`Module13_ProjectionResolution.bas:1006`](../src/target-spec-hybrid-v2/Module13_ProjectionResolution.bas)
**Verdict:** CONFIRMED · **Category:** correctness

`SelectModelEntityInView` calls `swDraw.ClearSelection2` on entry (line 954)
and again on every exit path (lines 1006, 1012), but never restores the
caller's prior selection. This contradicts the module header, which states
that selection happens only inside `SelectAnchorInView`, "which restores the
prior selection state before returning."

**Failure scenario.** The user has entities selected and runs
`R23_ProbeViewProjections` against the manual reference drawing.
`BuildViewProjections` records `initialSelectionCount=3`, then
`ResolveProjection` reaches Route D for a part drawing with no
`Component2`, and the selection is wiped. The probe finishes reporting
`finalSelectionCount=0` and `selectionClean=True` — the value it uses as
proof of non-interference — while the user's selection has been destroyed.

Route D is also reached from `MapModelEntityToDrawingForDatum`, so
`R23_ProbeOrdinateScheme` inherits the same behaviour.

---

### 8. View/cell mismatch is refused only after the sheet is rescaled

**File:** [`Module18_ContentEnvelope.bas:1945`](../src/target-spec-hybrid-v2/Module18_ContentEnvelope.bas)
**Verdict:** CONFIRMED · **Category:** correctness

The `views.Count <> targetCentres.Count` guard moved from the top of
`ApplyPlacementPlan` to after the rescale loop.

**Failure scenario.** `PlanPlacement` returns a replan whose cell count
differs from `views.Count`. The `Do While` loop first calls
`ApplyScaleToFit`, which mutates every root view's `ScaleDecimal` and clears
`UseSheetScale`. Line 1945 then detects the mismatch and returns
`layout=PlanMismatch` without moving a single view. The drawing is left
permanently rescaled by a call that reported it did nothing.

---

### 9. `On Error` cannot trap VBA compile errors

**File:** [`Module20_ProbeRunner.bas:391`](../src/target-spec-hybrid-v2/Module20_ProbeRunner.bas)
**Verdict:** PLAUSIBLE · **Category:** correctness

`R23_TouchAllModules` relies on a chain of `On Error GoTo M<n>Failed` to
localise which module fails to load. VBA compile errors are not trappable
runtime errors, so the handler chain cannot fire.

**Failure scenario.** `Module15_OrdinateScheme` has a syntax or type error.
`R23_RunAllProbes` gets verdict `NotClean` and calls `R23_TouchAllModules`
to localise it. VBA raises a compile error when resolving the call chain
rather than a runtime error, so control never reaches `M15Failed`; the Sub
aborts into the `Fatal` handler or dies outright. The log line documented
in `CLAUDE.md` as naming "the first module that would not load" either
never appears or appears as `firstFailedModule=None`.

Settle this with a deliberate live run: introduce a syntax error in one
module and check what `R23_RUN_TOUCH` actually reports.

---

### 10. Zero nominal now accepted, killing the `SystemValue` fallback

**File:** [`Module10_SectionDimensionEngine.bas:532`](../src/target-spec-hybrid-v2/Module10_SectionDimensionEngine.bas)
**Verdict:** CONFIRMED · **Category:** correctness

Removing `And candidate <> 0#` makes `TryNominalFromObsoleteMembers` accept
a `0.0` return from `GetSystemValue2` as a successful nominal read.

**Failure scenario.** A drawing reference dimension whose
`GetSystemValue2` returns `0.0` without raising. Lines 532-537 set
`nominalM = 0`, `route = nominalRoute=Obsolete.GetSystemValue2`, return
`True` and exit, so `dimension.SystemValue` — kept expressly as the final
fallback — is skipped. The 47 H7 requirement then fails to match on
nominal, and the evidence line claims a successful nominal read of
`0.000000000` rather than a declined route.

---

## Checked and clean

- `swView.Type = 2` is correct. `swDrawingViewTypes_e.swDrawingSectionView = 2`,
  verified against the SOLIDWORKS 2025 corpus via the `solidworks-api` MCP.
  `VIEW_TYPE_SECTION = 2` and `VIEW_TYPE_DETAIL = 3` in Modules 10/14/15/18
  match the enum.
- Every new cross-module reference resolves to a `Public` member:
  `Module19_SemanticQA.STAGE_MANUFACTURING_DEFINITION`,
  `Module15_OrdinateScheme.ORD_HORIZONTAL`/`ORD_VERTICAL`,
  `Module10_SectionDimensionEngine.CollectSectionViews` /
  `InventorySectionDimensions` / `BuildSectionRequirements` /
  `ReconcileSectionDimensions` / `VerifySectionDimensions`,
  `Module14_AnnotationImport.BuildAnnotationInventory` /
  `ReconcileWithLocationGraph` / `VerifyRequiredCoverage` /
  `IsModelImportEligibleView`,
  `Module18_ContentEnvelope.IsTemplateOrientationView` /
  `BuildProtectedRegions` / `BuildViewEnvelope`,
  `CLocationGraph.ClearImportedAnnotations` / `AnnotationCount` /
  `ProjectionCount` / `Projections`,
  `CRunEvidence.SealRequiredStages` / `StageIsProved`.
- `SolidWorksDocumentOpener.cs` out-parameter marshalling is correct in both
  the early-bound and `ParameterModifier` dispatch paths. The
  `swFileWithSameTitleAlreadyOpen = 65536` title-conflict branch does not
  close the user's document, as its comment promises.
- `Module7_TitleBlockEngine.AddRequiredManufacturingDefinitions` was deleted
  and its call site in `Module2_DrawingPipeline.RunDrawingPipeline` was
  removed in the same change. Consistent.
- `evidence.UsableLeft`/`UsableRight`/`UsableTop`/`UsableBottom` are
  populated by the sheet preflight and the rough layout before
  `CreateSemanticPrimarySection` reads them at
  `Module2_DrawingPipeline.bas:571`. The provisional section placement is
  not computed from zeroed bounds.

## Notes not raised as findings

- `Module2_DrawingPipeline.bas:9` hardcodes an absolute user-specific path
  in `R23_LAYOUT_SCRATCH_PATH`. It is a fail-closed authorization guard, so
  another machine is refused rather than misdirected.
- `Module21_EvidenceSink.LogLine` swallows every file error under
  `On Error Resume Next`. Total silent loss of the evidence file is possible,
  but degrading to `Debug.Print` alone is the module's documented design.
- `Module19_SemanticQA.EvaluateSemanticDrawing` returns a `Boolean` verdict
  that `Module2_DrawingPipeline.bas:414` discards. Stage failures still
  reach the evidence object, so the discard is cosmetic today. It becomes a
  defect the moment the return value is meant to gate the run.
