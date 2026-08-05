# R23 defect review and implementation plan — after the 18:45 production run

Review date: 2026-08-04
Base: working tree vs `6c00371`, plus the live run
`test_assets/iteration_evidence/macro_qa/20260804_184514_P-0251-14A-001/QA_REPORT.txt`
Managed revision under review: `target-spec-hybrid-v2-2026-08-04-r37`

Objective for this work package, per the user: **P-0251 must generate from
`Module1_Main.main` with no errors other than scale/resize ones.** Scale and
content-envelope repositioning are explicitly out of scope; the 2026-08-04
`UserAcceptedAsIs` policy stays in force.

## Method

Findings below are anchored to the failed production run, not to static
reading alone. Every claim cites a line from that QA report or from the
managed source. Nothing here has been re-run live; no source has been changed.

## Failure dependency map

```
F1 ActivateView return trusted ──► 0 annotations imported ──► 0 display
                                   dimensions ──► DIMENSION_ARRANGE fail,
                                   ANNOTATION_EXTENTS fail, FINAL_QA fail

F3/F4 central bore has no accepted projection ──► no singleton bore ──►
      SECTION_GEOMETRY fail ──► SECTION_DIMENSIONS fail ──►
      SECTION_CLEARANCE fail ──► "requires exactly one J-J section"

F2 rough layout runs before section+isometric exist ──► LAYOUT fail,
   0 view moves, drawing never laid out at all

F7 Hole Wizard params read as zero ──► NATIVE_CALLOUT_COVERAGE and
   MANUFACTURING_DEFINITION fail

F8 bottom-outline datum unmapped ──► ORDINATE_SCHEME fail
```

Ten of the twenty-three required stages failed. Those ten reduce to the five
independent roots F1, F2, F3/F4, F7, F8.

---

## F1 — CRITICAL. `IDrawingDoc.ActivateView`'s return value is trusted; nothing is imported

**File:** `src/target-spec-hybrid-v2/Module14_AnnotationImport.bas:305`

```vb
If Not swDrawing.ActivateView(viewName) Then
    EmitWarning evidence, "IMPORT_VIEW_NOT_ACTIVATED|view=" & viewName
    GoTo ContinueView
End If
```

On this build `ActivateView` returns False even when the view does become
active. The same run proves it, through `Module8_RuntimeSupport`'s verified
helper:

```text
ACTIVATE_VIEW|operation=ResolveOutlineDatum|view=Drawing View1|setterResult=False|readbackMatched=True
```

and the unverified call site immediately below it:

```text
IMPORT_VIEW_NOT_ACTIVATED|view=Drawing View2
IMPORT_VIEW_NOT_ACTIVATED|view=Drawing View1
```

`InsertModelAnnotations4` was therefore never called for any view. Downstream:
`Annotations imported: 0`, `Visible display dimensions: 0`,
`R23_IMPORT_RECONCILE|imported=0|reconciled=0|coverage=NoNativeHoleCallout;NoTolerancedDimension;NoOrdinateDimension`.

The same unguarded pattern exists at three more sites, each of which refuses
its whole transaction on the same false negative. They have not fired yet only
because upstream stages fail first:

| File | Line | Refusal it would emit |
|---|---|---|
| `Module15_OrdinateScheme.bas` | 1224 | `ORDINATE_CREATE_REFUSED\|reason=ViewActivationFailed` |
| `Module16_CalloutDefinition.bas` | 777 | `CALLOUT_CREATE_REFUSED\|reason=ViewActivationFailed` |
| `Module17_SectionPath.bas` | 645 | `SECTION_CREATE_REFUSED\|reason=ViewActivationFailed` |

`Module13_ProjectionResolution.bas:147` uses the raw return too, but only
records it and does not branch on it — that one is correct as written.

`Module8_RuntimeSupport.ActivateDrawingView` already implements the right
contract: call the setter, read the active view back by name, accept on
readback match, warn when the setter disagreed.

---

## F2 — CRITICAL. The only layout pass runs before the views it requires exist

**Files:** `Module2_DrawingPipeline.bas:200`, `Module9_LayoutEngine.bas:182`,
`Module9_LayoutEngine.bas:373`

`ArrangeViewsInMeasuredGrid` is called exactly once, at pipeline step 3, with
the pass name `"R23Rough"`. For P-0251 it dispatches to
`ArrangeP0251ReferenceZones`, which hard-fails unless a primary, a side, a
J-J section, and an isometric view are all present:

```vb
If primaryView Is Nothing Or sideView Is Nothing Or _
   sectionView Is Nothing Or isometricView Is Nothing Then
    evidence.AddFailure "P-0251 reference layout requires exactly one " & _
        "primary, side, J-J section, and isometric view."
```

At step 3 the sheet holds two views. The section is created at step 5 and the
isometric at step 11. The failure is therefore structural and unconditional,
independent of the section defect:

```text
- P-0251 reference layout requires exactly one primary, side, J-J section, and isometric view.
STAGE|name=LAYOUT|status=FAILED|detail=boundary, collision, or scale validation failed
Layout moves: 0
```

`Layout moves: 0` means not one view was ever positioned. The drawing is
currently laid out entirely by SOLIDWORKS defaults. This is not a scaling
issue and is not covered by the `UserAcceptedAsIs` waiver — that waiver
suspends the *final content-envelope* pass and explicitly preserves "the
initial structural grid".

---

## F3 — CRITICAL. Section fails on the principal bore, and production never says so

**Files:** `Module17_SectionPath.bas:186`, `Module2_DrawingPipeline.bas:470`

`ResolveBoreProjection` requires an **accepted** projection whose location is
alone in its semantic family. In `Drawing View1` the six accepted projections
all belong to one six-member CBORE family (`family.Count <> 1`, skipped). The
intended stepped bore is the singleton at `moment 0,0.062,0`,
`interval 0.000500000..0.016000000`, and it is not accepted:

```text
PROJECTION|view=Drawing View1|...|accepted=False|reason=ProjectionAnchorUnavailable
PROJECTION_ANCHOR|...|sourceFaces=2|facesProjected=2|boundaryEdges=4|circularEdges=4|mappedEdges=0|inventoryConfirmed=0|firstReject=NoRouteMappedThisEdge:A:Nothing:err0;B:Nothing:err0|chosenTier=None|anchorProof=None
```

So `path.RejectionReason` is `NoAcceptedSingletonBoreProjection`. Production
discards it:

```vb
If Not path.Resolved Then
    evidence.AddFailure "R23 semantic section path was not resolved."
```

The QA report contains no `SECTION_PATH`, `SECTION_PATH_CANDIDATE`,
`SECTION_CROSSING` or rejection-reason line at all — `ResolveSectionPath`
emits evidence only on success. The read-only probe logs all of this; the
production path is blind. Also, only the *last* view's path survives the loop
(`Module2:462` reassigns `path` per view), so a per-view diagnosis is lost.

---

## F4 — HIGH. Route D is suppressed in exactly the case that needs it

**File:** `Module13_ProjectionResolution.bas:665`

```vb
If mapped Is Nothing And Not visibleInventoryAvailable Then
    Set mapped = SelectModelEntityInView(...)
```

`Drawing View1` returns 39 edges from `GetVisibleEntities2`, so
`visibleInventoryAvailable` is True and the selection route is never tried —
even though Route A (`IView.GetCorrespondingEntity`) demonstrably declines for
real, visible geometry in that same view (`A:Nothing:err0` on all four circular
edges of the bore, while the six CBOREs map two edges each).

The gate is keyed on the wrong condition. It should be "Route A declined",
with acceptance still requiring **both** proofs that already exist:
`ISelectionMgr.GetSelectedObjectsDrawingView2` ownership *and*
`FindVisibleEntityIndex` membership of the visible inventory. That is strictly
stronger than the current inventory-less Route D path, not weaker, and it does
not touch the `MapVisibleDatumEntity` fail-closed guard the r37 handoff
protects.

Note the same site is where `Drawing View2`'s eleven failures come from. There
the failure is semantically correct — a `*Left` view shows no circles for
axes `(0,0,1)` or `(0,1,0)` — but the reported reason is
`ProjectionAnchorUnavailable` rather than `AxisNotNormalToView`, because
`QualificationFailureReason` tests the anchor before the axis
(`CViewHoleProjection.cls:144` before `:156`). Misleading, not wrong.

---

## F5 — HIGH. The error handler destroys the error it reports

**File:** `Module4_ModelItemImporter.bas:1029`

```vb
Failed:
    evidence.AddFailure "Dimension arrange API error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & _
        CStr(Err.Number) & ": " & Err.Description
```

VBA evaluates the concatenation left to right. `GetViewName` contains
`On Error Resume Next`, which resets the global `Err` object before
`Err.Number` is read. Live output:

```text
- Dimension arrange API error in 'Drawing View1': 0: 
```

An error number of 0 with an empty description is impossible from a real
raise. `Module2_DrawingPipeline.bas:337` already does this correctly by
capturing into locals first.

---

## F6 — HIGH. `MODEL_IMPORT_COVERAGE` proved itself on a drawing with zero annotations

**File:** `Module19_SemanticQA.bas` (`EvaluateModelImportCoverage`)

```text
QA_IMPORT_COVERAGE|view=Drawing View1|projections=11|accepted=6|coveredX=0|coveredY=0|annotated=1
QA_IMPORT_COVERAGE|view=Drawing View2|projections=11|accepted=0|coveredX=0|coveredY=0|annotated=0
STAGE|name=MODEL_IMPORT_COVERAGE|required=True|status=PROVED|detail=views=2/viewsWithAcceptedProjections=1
```

Zero annotations were imported into the drawing and the import-coverage stage
passed. The gate keys on accepted **projections**, which are a property of the
model-to-drawing map and say nothing about whether an import happened. This is
the same class of defect the earlier review recorded as finding 1
(`docs/R23_CODE_REVIEW_2026-08-04.md`); the live run now confirms a stage can
report PROVED while the thing it names did not occur.

---

## F7 — MEDIUM. Hole Wizard definitions read as all zeros but report `readStatus=Read`

**File:** `Module12_FeatureQualification.bas:451-503`

```text
FEATURE_ACCEPTED|name=CBORE for M6 Socket Head Cap Screw1|...|readStatus=Read|operation=HoleWizard|diameterM=0.000000000|depthM=0.000000000|cboreDia...
FEATURE_ACCEPTED|name=M5x0.8 Tapped Hole1|...|readStatus=Read|operation=HoleWizard|diameterM=0.000000000|depthM=0.000000000|...
```

Consequences:

```text
STAGE|name=NATIVE_CALLOUT_COVERAGE|status=FAILED|detail=definitions=3/native=1/controlledFallback=2/incomplete=2
STAGE|name=MANUFACTURING_DEFINITION|status=FAILED|detail=definitions=3/complete=1/attachmentProven=1
```

The 2025 Help confirms `IWizardHoleFeatureData2.HoleDiameter` is the right
property for these hole types (it is documented as not relevant only for
`swTapered`/`swTaperedDrilled`) and that it requires the "Accessing Selections
that Define Features" pattern, which the code does follow. So either
`AccessSelections` declined, or the definition is being read from the wrong
feature node, or these standard-driven holes expose their size through the
sibling `IWizardHoleFeatureData2.Diameter`/thread members instead. This needs a
probe before any code change.

Separately: a read that returns 0 for every dimensional member should not be
recorded as `DefinitionReadStatus = "Read"`. The zero-value case is currently
indistinguishable from a genuine read and only surfaces two stages later.

---

## F8 — MEDIUM. Vertical ordinate datum: keep the guard, fix the routing

**File:** `Module15_OrdinateScheme.bas` (`ResolveOutlineDatum`,
`OutlineDatumForModelEdge`)

```text
ORDINATE_DATUM|view=Drawing View1|...|datumPolicy=BottomOutlineGeometry|direction=Vertical|resolved=False|reason=NoBucketAvailable;outline=NoMappedBottomEdge(edges:126,curve:57,vertices:0,points:0,transform:0,notHorizontal:42,span:11,map:16)|mapSample=VisibleMapUnavailable:DatumMap:PolylineVisibilityUnavailable/source:IView.GetPolylines7/modelEdges:39/unsupported:0/error:0
```

126 model edges, 57 rejected on curve type, 42 not horizontal, 11 on span, and
**16 reached mapping and none matched**. The r37 `MapVisibleDatumEntity` guard
is behaving as designed and must stay fail-closed. The defect is in identity
matching against the non-empty `GetPolylines7` result, which is the same
mapping weakness as F4.

Carried forward unfixed from the earlier review: `OutlineDatumForModelEdge`
decides "horizontal" from two transformed endpoint Y values alone, so an
arc whose endpoints share a Y passes. P-0251's bottom face is arc-bounded.
Fixing F4/F8 mapping without also fixing this risks a datum placed at an arc
endpoint rather than at the true low point.

---

## Deferred by the user's scope decision

These are real but sit in code production no longer calls
(`RecordR23UserAcceptedLayout` reports `scaleChanges=0|viewMoves=0`):

- `Module18_ContentEnvelope.bas:1782` `ApplyScaleToFit` overwrites the approved
  isometric 1:2 and detail 3:1 scales.
- `Module18_ContentEnvelope.bas:1945` view/cell mismatch refused only *after*
  the rescale loop has already mutated every root view's scale.

Do not restore either path. Re-review them if automatic layout is ever
re-authorized.

## Confirmed fixed since the earlier review

- `Module2_DrawingPipeline.bas:469/474` — the non-short-circuit
  `If path Is Nothing Or Not path.Resolved` is now two separate tests.
- `CollectR23ImportViews` — loop two now excludes `Type = 2` and
  `AddR23ImportViewOnce` de-duplicates by view name.

---

# Implementation plan

Sequenced so each phase's evidence is readable before the next begins. Phases
1-3 are the ones that stand between the current state and a drawing that
generates its required content. Every phase ends with the full companion
`unittest` suite plus a deploy/readback, and the live gates named in
`Agents.md`.

## Phase 1 — Unblock annotation import and the layout pass

Smallest change with the largest reach. No new API surface.

1. `Module14_AnnotationImport.bas:305` — replace the raw `ActivateView` test
   with `Module8_RuntimeSupport.ActivateDrawingView`. Keep
   `IMPORT_VIEW_NOT_ACTIVATED` as the failure line for a genuine readback
   mismatch.
2. Same substitution at `Module15_OrdinateScheme.bas:1224`,
   `Module16_CalloutDefinition.bas:777`, `Module17_SectionPath.bas:645`.
   `Module13:147` stays as-is (records, does not branch).
3. `Module2_DrawingPipeline.bas:200` — make the step-3 pass a genuinely rough
   one. Two options; recommend (b):
   - (a) pass a flag so `ArrangeP0251ReferenceZones` tolerates a missing
     section/isometric at rough time;
   - (b) **recommended** — call `ArrangeViewsInMeasuredGrid` a second time
     after the isometric exists (between step 11 and the `"R23 final content"`
     rebuild at `Module2:302`), and let the step-3 call use the generic
     zone-aware path. This keeps one strict validator, run once, when the
     reference view set is actually complete, and stays inside the approved
     "initial structural grid" boundary.
4. Contract tests: an `ActivateDrawingView` fake returning setter-False /
   readback-match must still import; a rough-layout pass with two views must
   not fail the P-0251 reference-set assertion.

**Gate:** full suite green, deploy + 38/38 readback, manual VBE compile,
then one authorized production run. Expect `Annotations imported > 0`,
`Layout moves > 0`, `LAYOUT` PROVED. Section stages will still fail.

## Phase 2 — Make the section failure legible, then fix its root

1. Diagnosis first, no behaviour change: in
   `Module2_DrawingPipeline.CreateSemanticPrimarySection`, emit a per-view
   `SECTION_PATH_CANDIDATE|` line carrying `path.Summary()` and
   `path.RejectionReason` for **every** view tried, and include the reason in
   the final failure text. Stop overwriting `path` across loop iterations.
2. `Module13_ProjectionResolution.bas:665` — re-key the Route D gate on
   "Route A and Route B both declined" instead of
   `Not visibleInventoryAvailable`. When the inventory *is* available, require
   both `GetSelectedObjectsDrawingView2` ownership **and** a positive
   `FindVisibleEntityIndex` before accepting the mapped entity, and record
   `identity=SelectionOwnership+ISldWorks.IsSame`. Consult the
   `solidworks-api-lookup` skill for `IView.SelectEntity`,
   `ISelectionMgr.GetSelectedObjectsDrawingView2` and
   `IView.GetCorrespondingEntity` before writing it.
3. `CViewHoleProjection.QualificationFailureReason` — test `AxisNormalToView`
   before `HasSelectableAnchor` so a genuinely edge-on hole reports
   `AxisNotNormalToView` rather than `ProjectionAnchorUnavailable`. Reporting
   only; acceptance is unchanged.
4. Re-run the read-only probe runner on the scratch drawing before any
   production run and read
   `test_assets/iteration_evidence/probe_runs/<ts>/probe_log.txt` directly.

**Gate:** `R23_SECTION_END|resolvedPaths=1` in the probe log; then a
production run showing `SECTION_GEOMETRY` and `SECTION_DIMENSIONS` PROVED and
section count 1.

## Phase 3 — Manufacturing definitions

1. Probe-only first. Extend the read-only feature probe to report, per Hole
   Wizard feature: `AccessSelections` result, `HoleDiameter`, `Diameter`,
   `HoleDepth`, `ThreadDiameter`, `Standard`, `FastenerType`, `Size`,
   `EndCondition`. Look every one of those up through the
   `solidworks-api-lookup` skill first; do not assume any of them exists.
2. Only after that evidence: correct `ReadHoleWizardDefinition` to the route
   that actually carries the values on this build, and record the route in
   `DiameterProofSource`.
3. `DefinitionReadStatus` must distinguish `Read` from
   `ReadAllZeroValues` so the defect surfaces at the catalog, not two stages
   later.

**Gate:** `MANUFACTURING_DEFINITION` and `NATIVE_CALLOUT_COVERAGE` PROVED with
`complete=3`.

## Phase 4 — Ordinate datum

1. Reuse Phase 2's mapping fix in `MapModelEntityToDrawingForDatum`. The
   `MapVisibleDatumEntity` guard stays fail-closed; only identity matching
   improves.
2. Fix `OutlineDatumForModelEdge` to reject a curve whose parameterisation is
   an arc rather than a line, instead of inferring horizontality from two
   endpoint Y values. Filter candidates to view-visible edges before mapping —
   this also removes the per-edge selection round trips the earlier review
   flagged as a performance concern.

**Gate:** `ORDINATE_SCHEME` PROVED with `datumResolved=2`.

## Phase 5 — Close the self-satisfying gates

1. `Module19_SemanticQA.EvaluateModelImportCoverage` — key
   `ViewImportedNothing` on imported/reconciled annotation counts held on the
   projection, not on accepted projections or on
   `COrdinateBucket.RecordDirectionalCoverage` output. Assert ordering: read
   coverage before `CollectOrdinateSchemes` runs, or read a field the scheme
   builder cannot write.
2. `Module4_ModelItemImporter.bas:1029` and any sibling handler — capture
   `Err.Number`/`Err.Description` into locals as the first statements of the
   handler.
3. `Module12_FeatureQualification.ProbeCurveRole` — a `False` return from
   `ProbeFeatureCircularEdge` is an observed disagreement, not "try the next
   feature"; record it as a failure. (Earlier review, finding 5.)
4. `Module10_SectionDimensionEngine.bas:532` — restore the `candidate <> 0#`
   test so a `0.0` from `GetSystemValue2` declines the route instead of
   consuming the `SystemValue` fallback. (Earlier review, finding 10.)

**Gate:** re-run Phase 1-4's production run and confirm no stage reports PROVED
for something the counts show did not happen.

## Phase 6 — Acceptance

Full production run, QA report plus full-sheet screenshot, visual comparison
against the manual reference. Only scale/placement deltas may remain open.

## Standing constraints for every phase

- Use `tools/probe-runner/Run-R23Probes.ps1 -Deploy -DrawingPath <fixture>` for
  read-only evidence; read the written log, never ask for a paste.
- Every mutating production run needs a fresh manual VBE full compile after
  deployment and explicit per-run user authorization.
- Do not disable `DIAGNOSTIC_DRAWING_MODE` to make a run look clean.
- Do not restore automatic rescaling or content-envelope repositioning.
- Do not save or modify the P-0251 model or any reference drawing.
- Bump `MACRO_SOURCE_REVISION` and update `Changelog.md`,
  `CURRENT_STATUS.md`, `SOLIDWORKS_API_VALIDATION.md` and `README_IMPORT.md`
  on every deployable change.
