# R23 Phase 0 probes

These probes exist because `docs/R23_IMPLEMENTATION_PLAN.md` deliberately
blocks production feature classification until the load-bearing behavior is
proved in the installed SOLIDWORKS 2025 build.

## Prepared probe

`Module_R23Phase0FeatureProbe.bas` is a read-only VBA module. Its public entry
point is:

```text
R23_ProbeActivePartFeaturesAndCurves
```

The procedure:

- refuses to run on any model outside the three authorized fixtures;
- records `GetTypeName2`, `GetTypeName`, and the normalized effective type;
- records `GetDefinition` runtime type and configuration-specific suppression;
- probes exact `CUT`/`CUTTHIN` and base `EXTRUSION` definitions through
  `IExtrudeFeatureData2`;
- records extrusion contour count/state and owned profile-subfeature count;
- probes `HoleWzd` definitions through `IWizardHoleFeatureData2`;
- pairs every successful `AccessSelections` with
  `ReleaseSelectionAccess`;
- inventories feature-owned faces and cylindrical faces;
- logs pattern/mirror seed-feature evidence; and
- compares the two required circular-curve read orders on owned cylindrical
  face edges.

It does not call `ModifyDefinition`, rebuild, save, or change the part.
It mirrors the Immediate Window records to a timestamped
`R23_FEATURE_*.log` file in the live-probe evidence directory, so a long
transcript is not lost to the Immediate Window buffer.

## Operator instructions

The current disposable macro is:

```text
test_assets/iteration_evidence/r23/20260730-075811/live-probes/R23_Phase0FeatureProbe.swp
```

Its embedded `Module1` contains the first probe revision. Before the next run,
replace the complete `Module1` code with the current contents of:

```text
tools/r23-probes/Module_R23Phase0FeatureProbe.bas
```

Then:

1. choose **Debug > Compile Feature_export**;
2. activate the authorized workspace P-0251 part below;
3. clear the Immediate Window;
4. run `R23_ProbeActivePartFeaturesAndCurves`;
5. copy the complete Immediate Window transcript, from `R23_PROBE_BEGIN`
   through `R23_PROBE_END`;
6. share the timestamped `R23_FEATURE_*.log` path printed in
   `R23_PROBE_BEGIN`; this file is the authoritative untruncated transcript;
7. save only the disposable probe macro if you want to retain the corrected
   source;
8. do not save the part.

The initial live target is:

```text
test_assets/models/P-0251-14A-001.SLDPRT
```

Do not save the fixture after the probe.

## Prepared import-transaction harness

The expanded-mask comparison is isolated in a copy of `Fable.swp`:

```text
test_assets/iteration_evidence/r23/20260730-075811/live-probes/R23_Phase0ImportProbe.swp
```

The guarded deployment manifest changes only `Module1_Main`,
`Module2_DrawingPipeline`, and `Module4_ModelItemImporter` in that disposable
copy. It never targets production `Fable.swp`.

The two public entry points are:

```text
R23_ProbeImportAllViews
R23_ProbeImportSelectedViews
```

Each entry point creates a fresh unsaved P-0251 drawing from the controlled
template, stops the normal pipeline before production import/fallback/title/QA,
and runs exactly one comparison variant:

- `R23_ProbeImportAllViews`: selected primary anchor,
  `AllViews=True`, `DuplicateDims=True`;
- `R23_ProbeImportSelectedViews`: section, side, then primary,
  `AllViews=False`, `DuplicateDims=True`.

Both use mask `18055274`, including `swInsertDimensions=8`,
Hole Wizard profile/location dimensions, native hole callouts, and
toleranced dimensions. They record returned annotations, final owner views,
source dimension identity, nominal/tolerance/fit data, callout variables,
attachments, duplicate identities, and raw display geometry. Records are
mirrored to timestamped `R23_IMPORT_*.log` files.

Before the live import comparison, close this disposable macro in the VBA
editor, leave SOLIDWORKS running, and execute from workspace PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1 -ManifestPath .\tools\r23-probes\import-transaction-manifest.json
```

Then:

1. open the disposable import macro with **Tools > Macro > Edit**;
2. choose **Debug > Compile Fable** and stop if the command selects an error;
3. activate the authorized workspace P-0251 part;
4. clear the Immediate Window;
5. run `R23_ProbeImportAllViews`;
6. capture an uncropped full-sheet screenshot and retain the printed
   `R23_IMPORT_*.log`;
7. close only the generated drawing without saving and reactivate P-0251;
8. clear the Immediate Window and run `R23_ProbeImportSelectedViews`;
9. capture the second full-sheet screenshot and log;
10. close the second generated drawing without saving.

The explicit ordinate, section-dimension, and section-line probes remain
required after the feature and import evidence identifies stable owned
projection anchors and the import transaction to retain.

## Final drawing-contract probes

The final Phase 0 drawing-contract module is:

```text
tools/r23-probes/import-transaction-source/Module_R23Phase0DrawingProbes.bas
```

It exposes two public entry points:

```text
R23_ProbeDatumFirstXYOrdinates
R23_ProbeSectionDimensionsAndJJGeometry
```

Each entry point:

- requires the authorized workspace P-0251 part to be active;
- creates a fresh unsaved drawing;
- uses the proved section/side/primary `AllViews=False`,
  `DuplicateDims=True` import transaction;
- writes an untruncated timestamped log under the R23 live-probe evidence
  folder;
- restores normal pick mode and clears all selections; and
- compares the model save flag before and after without rebuilding, changing,
  or saving the model.

The ordinate probe:

- obtains the exact visible `Component2` from `IView.GetVisibleComponents`;
- maps feature-owned complete circular edges into the primary drawing view;
- resolves six M6 counterbore locations into two X and three Y coordinates;
- uses the stepped-bore centre as the X zero and a mapped bottom-left model
  vertex as the Y zero;
- selects each datum first and appends every feature entity individually with
  `IEntity.Select4`;
- records selection order/count/type and the exact
  `IModelDocExtension.AddOrdinateDimension` return code; and
- calls `IModelDoc2.SetPickMode`, clears selection, and records final
  selection/readback state after each group.

The section probe:

- inventories every imported section display dimension;
- records nominal, dimension type, tolerance type, fit type/style, hole/shaft
  fit, min/max status/value, annotation display primitives, and exact
  selection identity for the 47 mm, 40 mm, and first linear dimension;
- reports whether the installed run proves H7 and whether SOLIDWORKS exposes a
  true `swDiametricLinearDimension`; and
- records `IDrSection` state plus the raw and parsed J-J segment, arrow, and
  label geometry from `GetSectionLineCount2` and `GetSectionLineInfo2`.

Close the disposable macro in the VBA editor, keep SOLIDWORKS running, and
deploy the four-module probe overlay from workspace PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1 -ManifestPath .\tools\r23-probes\drawing-contract-manifest.json
```

Then:

1. open `R23_Phase0ImportProbe.swp` through **Tools > Macro > Edit**;
2. choose **Debug > Compile Fable** and stop if VBA selects an error;
3. activate `P-0251-14A-001.SLDPRT`;
4. run `R23_ProbeDatumFirstXYOrdinates`;
5. retain the printed `R23_ORDINATE_*.log` and an uncropped full-sheet
   screenshot;
6. close only the generated drawing without saving and reactivate P-0251;
7. run `R23_ProbeSectionDimensionsAndJJGeometry`;
8. retain the printed `R23_SECTION_*.log` and a second uncropped screenshot;
9. close the generated drawing without saving; and
10. do not save the fixture model.

## 2026-07-31 corrected drawing-contract probes (build 20260731.2)

The two final-probe defect sets identified below have been corrected in
`import-transaction-source/Module_R23Phase0DrawingProbes.bas` (probe build
`20260731.2-mapping-frame-h7-contracts`) and, for the section-line
construction capture only, in the disposable
`import-transaction-source/Module2_DrawingPipeline.bas` overlay. Production
`src/target-spec-hybrid-v2/` is unchanged. The corrected sources passed
structural static checks and the read-only disposable-manifest preflight;
they have not been deployed, compiled, or executed.

### Ordinate probe corrections

- Every owned face and edge of the target feature now logs a qualification
  record: `R23_ORDINATE_FEATURE` (owned-face count), `R23_ORDINATE_FACE`
  (cylinder test, radius, edge count, verdict), `R23_ORDINATE_FACE_TRANSFORM`
  (page-frame centre transform with proof), and `R23_ORDINATE_EDGE`
  (runtime type, `IsCircle`, `GetCurveParams3` availability/range/closure
  gap, per-route mapping results, verdict and reason).
- Three mapping routes are compared per complete circular edge:
  - route A: active-part model edge -> `IView.GetCorrespondingEntity`;
  - route B: `IComponent2.GetCorrespondingEntity` -> view correspondence;
  - route C: `IView.GetVisibleEntities2(component, Edge)` inventory with
    `ISldWorks.IsSame` identity cross-checks (`R23_ORDINATE_VISIBLE_EDGE`/
    `R23_ORDINATE_VISIBLE_EDGES`).
- Qualification is unchanged: only ownership-backed, complete circular edges
  with a route A or route B mapped drawing entity are accepted. Visible
  circles alone are never accepted.
- The bottom-left datum search logs route counters
  (`R23_ORDINATE_VERTEX_SUMMARY`) and the chosen vertex
  (`R23_ORDINATE_VERTEX_CHOSEN`).
- `AddOrdinateDimension` remains gated on exactly six counterbore locations
  resolving to two unique X and three unique Y coordinates; the selection
  transaction still records datum-first order, every `Select4` result,
  selected types/counts, decoded result codes, display-dimension deltas,
  `SetPickMode`, and zero-selection cleanup.

### Section probe corrections

- Every per-dimension value (`nominalM`, `nominalAvailable`, `targetName`,
  `dimensionH7Proven`, `toleranceSummary`) is reset at the top of each
  display-dimension iteration, eliminating the stale labels of the first run.
- Diameter targets now require the live-proven `swDiameterDimension = 6`
  type plus the exact nominal; type 15 is recorded but not required. The
  summary reports `diameter47Count`, `diameter40Count`, and
  `exactTargetCounts`, and any second match is labelled `*_DUPLICATE`.
- The original part-source dimensions are read directly through
  `IModelDoc2.Parameter` (`R23_SECTION_SOURCE_TOLERANCE` for `D1@Sketch4`
  and `D1@Sketch6`): full name, nominal, tolerance type, fit type/style,
  hole/shaft fit, `GetToleranceFitValues`, min/max status/value, and active
  configuration. `R23_H7_AUTHORITY` states whether H7 exists in the source
  and, if absent, that the user must choose between model-authoritative
  fail-closed and controlled target-spec/reference authority. The probe
  never invents or applies H7.

### J-J geometry corrections

- The Module2 overlay captures each section segment's original page-frame
  construction points and the converted view-sketch values passed to
  `CreateLine` (`R23_JJ_CONSTRUCTION_POINT`, labelled
  `frame=PageAtConstruction` and `frame=ViewSketch`).
- The parsed 49-item payload is compared point-by-point with the captured
  view-sketch path (`R23_JJ_FRAME_PROOF`), proving the segment frame instead
  of assuming it.
- Segment endpoints are converted to the current page frame exactly once
  (`R23_JJ_PAGE_TRANSFORM`, `R23_JJ_SEGMENT_POINT`) using the exact inverse
  of the construction conversion (`IView.GetXform` origin/scale plus
  `IView.Angle`); arrow and label payload values are page-frame directly.
- Page-frame reserved regions are logged with provenance
  (`R23_JJ_REGION`): sheet size, content border from zone margins, title
  block, the measured `*P-0251-14A-001*` sheet-format note extent via
  `INote.GetExtent`, and every view outline.
- Each segment, arrow, and label box receives an explicit page-frame
  clearance verdict (`R23_JJ_CLEARANCE`, `R23_JJ_CLEARANCE_SUMMARY`).
  These are truthful measurements, not acceptance claims.

### 2026-07-31 corrected-probe run result

Section and J-J gates closed. Ordinate did not: all 18 owned counterbore
faces were rejected as `NotCylindrical` because the probe used a bare
`If Not surface.IsCylinder Then`. A raw SOLIDWORKS `VARIANT_BOOL` of `1`
makes `Not value` evaluate to `-2`, which VBA treats as True. Build
`20260731.2` normalizes with `CBool`, logs `isCylinderRaw`, and applies the
same normalization to `IDrawingDoc.ActivateView`. The three mapping routes
have still never run.

Measured J-J clearance found three violations, all at the content-border top
(segment 1 start, upper arrow, upper label). The lower arrow and label clear
the measured part-identification note extent by about 7.7 mm, correcting the
earlier screenshot-derived claim of a lower-band intrusion.

The direct part-source readback proved H7 is absent from `D1@Sketch4`.

### 2026-07-31 second corrected run

The `CBool` fix worked: twelve owned counterbore faces read as cylinders with
correct radii, all passed the page transform, and the six page centres form
the required two-X by three-Y grid at 30 mm and 40 mm spacing. Ownership,
cylinder qualification and the model-to-page transform are proved.

The run then stopped at the edge-closure gate, which returned False for every
edge in the document including about thirty perfect circles. That gate chain
ended in a single-line `If ... Then _` continuation immediately before its
success assignment. It is now block-structured, reports a per-edge
`rejectGate`, and logs the `closureToleranceM` in force. No `Then _`
construct remains in the module.

### 2026-07-31 third run — SOLIDWORKS Boolean contract

`rejectGate` settled the cause in one run: a single call logged both
`isCircle=True` and `rejectGate=IsCircleFalse`.

In this VBA host, `If value Then`, `If value = False Then` and `CStr(value)`
are safe on a SOLIDWORKS COM Boolean; `If Not value Then` is not, yielding
`-2` which VBA treats as True. `CBool` is not dependable —
`CBool(rawVariant)` worked for `ISurface.IsCylinder` while
`CBool(curve.IsCircle)` did not. The probe now uses `NormalizeSwBoolean`,
an explicit `(CDbl(rawValue) <> 0#)` comparison, for `IsCircle`,
`IsCylinder`, `ActivateView`, both `Select4` calls and `GetSaveFlag`, and
logs raw values alongside normalized ones.

The `ICurve.CircleParams` exclusion is withdrawn as unproved: the
`SkippedNotCircle` sentinel came from the same defect in
`Module_R23Phase0FeatureProbe.ReadCircleState`, so the API was never called.
Each edge record now carries a non-load-bearing `circleParams=` field.

### 2026-07-31 fourth run — mapping settled, error 91 at ISelectData.View

All three routes ran for the first time. Route A
(`IView.GetCorrespondingEntity(modelEdge)`) works; route B
(`IComponent2.GetCorrespondingEntity`) returns `Nothing` every time; route C
proves route A's output is genuine drawing context via `ISldWorks.IsSame`.
Mapping is per-edge, so every owned edge must be attempted. Six unique
locations, two unique X, three unique Y, and both datums resolved.
`ICurve.CircleParams` works, closing the R23-006 exclusion.

Both ordinate groups then failed with runtime error 91 before any selection.
`CreateSelectData` returned a live object, so the failing statement is
`Set selectData.View = swView` — the same defect
`Module2_DrawingPipeline.CreatePrimarySection` already documents. The probe
now guards that assignment, reports `viewBinding=Bound` or
`UnboundAfterError:91`, proves each selection's owning view through
`ISelectionMgr.GetSelectedObjectsDrawingView2`, and tags every group record
with the exact `lastStep` reached.

### 2026-07-31 fifth run — PHASE 0 CLOSED

Both ordinate groups returned `AddOrdinateDimension = 0 Success` with exact
selection counts (datum + 2 = 3, datum + 3 = 4), display-dimension deltas +2
and +3, `ownerView=Drawing View1` on every selection, `SetPickMode` called,
zero selections remaining, and an unchanged fixture.

Values: `+15.00`/`-15.00` about the stepped-bore centre and
`10.00`/`50.00`/`90.00` from the bottom-left vertex datum.

`viewBinding=UnboundAfterError:91` on both groups confirms the
`ISelectData.View` defect, and both completed normally with unbound selection
data. Created ordinates report `Type2 = 1` and `Type2 = 7`, so QA must accept
`1`, `7` and `8`.

These probes have served their purpose. Do not rerun them for production
work; the contracts they established are recorded in
`docs/R23_IMPLEMENTATION_PLAN.md` and `docs/SOLIDWORKS_API_VALIDATION.md`.

### Rerun evidence expectations

A corrected ordinate run is diagnostic-complete when either every mapping
route is proved unusable with per-edge reasons, or six locations resolve and
both ordinate groups complete with decoded results and clean cleanup. A
corrected section run must show exactly one `DIAMETER_47`, one
`DIAMETER_40`, one `FIRST_LINEAR`, non-stale labels elsewhere, the direct
part-source tolerance records, a successful `R23_JJ_FRAME_PROOF`, page-frame
segment conversions, and clearance verdicts for every J-J item. Screenshots
remain required for visual confirmation.

## 2026-07-31 final-probe results — do not rerun unchanged

Both entry points above were compiled and run by the user. Their evidence and
the exact correction checklist are indexed in
[`docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md`](../../docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md).

- The ordinate run found the visible component but stopped with
  `CounterboreLocationsUnavailable` before selecting a datum or calling
  `AddOrdinateDimension`. It requires per-face/per-edge qualification logs and
  side-by-side direct, component-mediated, and visible-entity mapping
  diagnostics. This result is not an ordinate API failure.
- The section run found the exact 47 mm and 40 mm imports as
  `swDiameterDimension = 6`. Neither had H7 or a nonzero tolerance. Its loop
  failed to reset per-item diagnostic fields, so later target labels and
  nominals can be stale.
- The 49-item J-J payload is structurally complete, but segment coordinates are
  in the source-view sketch frame while arrow and label positions are in the
  page frame. A corrected probe must transform every segment endpoint to page
  coordinates exactly once before any clearance comparison.
- The captured logs and screenshots are diagnostic evidence, not manufacturing
  acceptance. Production R23 source remains unchanged and blocked until the
  corrected probes close these contracts.
