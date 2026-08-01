# Changelog

## 2026-08-01 - R23 Phases 8 and 9 first live run: four defects fixed

Both probes ran read-only against the P-0251 reference drawing. Neither gate
is satisfied; every cause was in my code, not the drawing.

### Phase 8 defect 1: the nominal never read

All seven section dimensions returned nominalAvailable=False while the same
dimension object answered toleranceType, fitType and the fit strings on the
next line - so GetSystemValue3(swThisConfiguration, Empty) declined
specifically. The seven are RD1..RD7@Drawing View6: drawing-authored
REFERENCE dimensions, which have no configuration to ask about. Phase 0 read
D1@Sketch4, an imported model dimension, where that route works.

Matching needs the nominal, so all seven requirements reported Missing while
their dimensions sat in the view. TryReadNominal now tries
swThisConfiguration, then swAllConfiguration, then the obsolete
GetSystemValue2("") and SystemValue as labelled last resorts, and NAMES the
route that answered so a later run can drop whichever proved unnecessary.
When every route declines it reports the raw shape of the GetSystemValue3
result, because "no nominal" and "an empty SafeArray" are different
problems.

### Phase 8 defect 2: the type rule rejected the real drawing

Every section dimension is swLinearDimension=2, including the one carrying
H7. Phase 0's type-6 D1@Sketch4/D1@Sketch6 evidence describes an earlier
state of the same fixture. Both states are real, so a diameter requirement
now accepts type 6, type 15 and the linear types, and
IDisplayDimension.Diametric is recorded to say which the drawing actually
displays as a diameter. It is reported rather than used to reject: the
nominals are 5.5 mm apart at the closest, so type corroborates and does not
discriminate. An unproved diameter display is recorded as
NotDisplayedAsDiameter or DiameterDisplayUnreadable, and any requirement
carrying its own failure is kept out of the satisfied count.

### Phase 8 defect 3: every log line printed twice

CRunEvidence.AddInfo prints what it records; the probe printed the same
requirement lines again.

### Phase 8 result worth keeping

RD4 carries toleranceType=8, fitType=0, holeFit=H7, minimumM=0.000000,
maximumM=0.000025, both statuses 0, two attached edges. H7 +0.025/0.000,
live, on a drawing reference dimension - independent corroboration that the
fit is drawing-authored and absent from the model.

### Phase 9 defect: a probe leaning on a production gate

The run aborted with "Controlled sheet has neither an ITitleBlock definition
nor a proved legacy title-block rectangle" before building a single
envelope. R23_ProbeContentEnvelope called
Module8_RuntimeSupport.MeasureControlledSheetRegions - a fail-closed gate
for a sheet the macro CREATES from the controlled template, which the
designer's reference drawing is not. Worse, that procedure SETS
ISheet.SheetFormatVisible, so a run promising mutations=0 had already
attempted one. Same shape as the Phase 5 EmitRunEvidence mistake.

MeasureSheetRegions now measures read-only - ISheet.GetSize plus
ISheet.GetZoneMargin - and REPORTS what it cannot measure
(titleBlock=Absent, contentBorder=Unmeasured,
usableSource=SheetExtentNoBorder) instead of aborting. Only an unusable
sheet size stops the probe.

BuildProtectedRegions now gates every rectangle on measured bounds. Unset
evidence fields would have produced a degenerate rectangle at the sheet
origin and reported false ProtectedIntrusion violations against it: a
boundary that does not exist is not a boundary at zero.

### Verification

Suite 345 tests with the five known-stale R20 failures. Preflight 35
managed components. MACRO_SOURCE_REVISION unchanged at
target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 9: content-envelope-aware final layout

`Module18_ContentEnvelope.bas` (33 procedures), `CContentEnvelope.cls`.
Statically verified only; no live run yet.

### What an envelope is

Everything that travels with a view: model outline, display-dimension lines
and text boxes, note extents, leader points, section segments, arrow
geometry, and both J-labels with their text heights. Each source is counted
separately, and HasAnnotationContent refuses to call an outline-only
rectangle a content envelope - that is the old behaviour wearing a new name.

### Frames, which is the real work

Four sources document their frame and one does not. GetOutline is page
frame; IAnnotation.GetPosition is sheet-relative in drawings;
INote.GetExtent is sheet space; GetSectionLineInfo2 is VIEW-SKETCH frame,
proved by Phase 0's payloadSegmentFrame=ViewSketchProved; and the
IDisplayData Remarks state no frame at all.

Section geometry is converted through ViewSketchToPage, the exact inverse of
Module17's PageToViewSketch, and ProveInverseTransform round-trips a real
page point through both before anything is contributed. Two functions that
claim to be inverses either agree to floating-point noise or one is wrong.

Display-data points are contributed and their agreement with the view's own
documented outline is COUNTED, not asserted. Claiming a frame the Help does
not state would be the same confident guess that cost this project the
swInsertDimensionsMarkedForDrawing bug.

### Three specific traps

GetTextPositionAtIndex is an OFFSET from the display-data origin, not a
coordinate; used absolutely it drags every envelope towards the sheet
origin. Leader points are consumed as XYZ triples from the returned array
rather than derived from GetLeaderStyle, whose value is OR-ed with
attachment bitmask flags that the corpus returns with mangled values. And
GetSectionLineInfo2's grammar is ambiguous between its Remarks and
GetSectionLineCount2's - one layer double or one per section line - so both
readings are walked in a dry run and the one whose consumption matches the
array length exactly is the one used.

### The fixed upward bias is replaced, not adjusted

Module9_LayoutEngine lines 442-446 pin the P-0251 source row to the top
boundary. PlanPlacement packs rows from the envelopes' own widths and
centres the block in the usable rectangle; contracts assert topBoundary -,
Bias and rowCenterY are all absent. A row pinned to a boundary has nowhere
to put the annotations that hang above it.

Placement and movement are separate procedures. PlanPlacement returns target
centres and touches nothing, so the whole plan is inspectable before a view
moves.

### Clearance and the things layout may not do

Every view-view and view-protected pair is checked with a separating-axis
measure, so touching rectangles score zero rather than passing, and the
check count is reported so an empty loop cannot read as a pass. Section
views get 2 mm from protected regions. The content border is protected as
four strips, not one rectangle - the drawable area is inside it.

No ScaleDecimal assignment, no ScaleRatio, no SetScale anywhere: R23-907
forbids shrinking a view to force a fit, so content that does not fit is
LargerSheetRequired with the required and available sizes stated. SealLayout
photographs the mutation ledger when layout completes so R23-909 can prove
nothing was created afterwards.

### Verification

Procedure blocks balanced 33/33 and 11/11. ANSI-only bytes, CRLF, no BOM,
max line 76. 32 Phase 9 contracts. Suite 334 tests with the five known-stale
R20 failures. Preflight 35 managed components. MACRO_SOURCE_REVISION
unchanged at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 8: semantic section-dimension engine

`Module10_SectionDimensionEngine.bas` (28 procedures),
`CSectionRequirement.cls`. Statically verified only; no live run yet.

### Reconcile before create

The section already carries imported dimensions - Phase 0 counted seventeen -
so `ReconcileSectionDimensions` runs before any creation path and
`CreateSectionDimension` refuses outright for a requirement that already
matched. Each requirement records six independent observations about what it
matched: source dimension identity, attached geometry, semantic role,
nominal, type and tolerance. Nominal and accepted type decide the match; the
other four are recorded so the match can be audited.

Every match is counted, not just the first, so R23-811 can fail on a
duplicate rather than quietly dimensioning the same thing twice.

### REQUIRED and OBSERVED never touch

CSectionRequirement splits what the specification demands from what was read
back, and nothing writes an OBSERVED field from a REQUIRED one. A
requirement that reports its own nominal back as the observed nominal proves
nothing; a contract asserts the assignment never appears.

### The obsolete tolerance route is gone

All four IDimension tolerance members - GetToleranceValues,
SetToleranceValues, GetToleranceFitValues, SetToleranceFitValues - are
marked obsolete by the 2025 Help, each superseded by an IDimensionTolerance
member. The Phase 0 probe used them; production reads
IDimension.Tolerance instead.

GetMinValue2 and GetMaxValue2 return a STATUS and hand the value back by
reference, so the status is reported beside the value it qualifies. A zero
value with a failed status is not a zero tolerance.

### H7 provenance is enforced, not just documented

REFERENCE_HOLE_FIT, the two deviations and
REFERENCE_FIT_AUTHORITY = "TargetSpecReferenceAuthority.NotModelData" are
each stated once. ApplyReferenceFit sets swTolFITWITHTOL BEFORE the values,
because SetValues2 refuses while the type is swTolNONE by its own Remarks
and FitType is only available for the fit types; it then reads the result
back rather than trusting the two return values, normalizing both COM
booleans.

EvaluateTolerance will not claim model provenance for a tolerance it merely
found on the drawing. Present-on-drawing is recorded as PresentOnDrawing
plus the same reference authority, because Phase 0 read the part source
directly and proved it carries none.

### No free text, ever

No InsertNote, no CreateText, no SetText anywhere in the module, and every
failure exit from CreateSectionDimension carries
policy=NoFreeTextSubstitute. A note is not a dimension: it does not move
with the geometry and cannot be inspected as one.

### Per-dimension locals reset every iteration

VBA block-scoped locals live for the whole procedure. That is exactly how
the Phase 0 section inventory mislabelled eleven of its seventeen
dimensions - index 6 set DIAMETER_40 and indices 7 to 17 kept the label.
Both loops here reset every field they report.

### Deferred, all for the same reason

R23-803's creation half needs live entity selection. R23-808 assigns lane
NAMES and leaves coordinates to Phase 9, which owns the annotation
envelope - the same call Phase 7 made about the section view's placement.
R23-809's predicate exists but Module9_LayoutEngine does not consult it, and
R23-810 detects the Module7_TitleBlockEngine free-text bore callout without
removing it, because removing it before real dimensions exist would leave
the bore undefined.

### Verification

Procedure blocks balanced 28/28 and 5/5. ANSI-only bytes, CRLF, no BOM, max
line 76. 37 Phase 8 contracts. Suite 302 tests with the five known-stale R20
failures. Preflight 33 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 7 read-only gate satisfied

`resolvedPaths=1|segments=3|columnHoles=3|crossingsProven=4|
sectionFailures=None|creations=0|drawingUnchanged=True`.

### The path resolved as the reference approves

In Drawing View4: w1=0.207331779,0.237414746 (bore centre);
w2=0.207331779,0.167414746 (same X, highest row); w3=0.192331779,0.167414746
(minimum-X column, that row); w4=0.192331779,0.087414746 (same column,
lowest row). distinctColumns=2, distinctRows=3, and crossingsProven=4 - the
bore plus all three holes on the chosen column.

The other three views report NoAcceptedSingletonBoreProjection, which is
correct: the bore is not accepted in them, so no path is invented.

### The frame transform cross-checks against the model

The arithmetic is exact (0.207331779 - 0.229331779 = -0.022), but that only
proves the code does what it says. The corroboration is independent: the
bore's Plucker moment is (0, 0.062, 0) and its viewY is 0.062 exactly. Every
counterbore behaves identically - viewY equals its moment's Y (-0.008,
-0.048, -0.088) and viewX equals its moment's X minus a constant 0.022, the
view's own centring offset, the same across all seven holes.

A wrong transform does not produce one shared offset across seven
independent points. This matters because Phase 8 onward depends on this
conversion and mixed frames have been a real defect in this project before.

### One evidence defect the run exposed

Views rejected before crossings could be tested - no bore, too few columns -
were also reporting NotAttempted in sectionFailures. That is the crossing
proof's initial STATE, not a failure of it, and listing it beside reasons
that are real dilutes them. Now excluded.

### Still unrun

R23-705's creation half, R23-706 and R23-707 all require mutation. The
selection-order verification and the GetSectionLineInfo2 read-back are
written and contract-tested but have never executed.

### Verification

23 Phase 7 contracts. Suite 265 tests with the five known-stale R20
failures. Preflight 31 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 7: J-J section path from model intent

`Module17_SectionPath.bas` (21 procedures), `CSectionPath.cls`. Statically
verified only; no live run yet.

### The path

Four waypoints, three segments, every coordinate a proved projection's page
coordinate: bore centre; same X at the highest face-hole row; minimum-X
column at that row; same column at the lowest row. R23-703 then proves the
path actually crosses the bore and every hole on the chosen column, naming
each one it misses.

### What is deliberately absent

No `extension`, no `topY`/`bottomY`, no `leftX`/`rightX`, and none of the
fractions 15/72, 90/196, 15.84/24 or 0.1x. A percentage of a view outline
knows nothing about where the holes are, which is exactly why the old upper
label landed in the zone region and the lower arrow in the
part-identification band. Contracts assert each token is gone.

### Contracts worth carrying

- The bore is the largest SINGLETON-family location, read from the graph.
  No radius threshold, so a different part is not misclassified.
- The grid is proved: fewer than two columns or two rows is a stated
  rejection, not an array index that happens to work.
- Crossings are judged against each hole's own projected radius, and the
  point-to-segment distance is CLAMPED to the finite segment. Unclamped, a
  circle beyond an endpoint reports as crossed because the infinite line
  passes through it.
- The page-to-view-sketch conversion happens exactly once per waypoint,
  immediately before CreateLine. Nothing upstream holds view coordinates, so
  there is nothing to convert twice.
- Segment selection order is verified before CreateSectionViewAt5, whose
  Remarks require the section line to be selected first. SOLIDWORKS reads
  the segments in selection order, so an unverified order cuts a different
  shape.

### Two defects caught before compiling

CreateSectionViewAt5 was being passed an empty label instead of the resolved
one. And its X/Y - the CENTRE of the new view - were being defaulted to
waypoint 3, a point inside the source view, which would have stacked the
section on top of the view it was cut from. Placement is now a caller
argument: choosing where a view sits is layout, and layout needs the full
annotation envelopes this module cannot see.

### R23-704 is half met, deliberately

Same shape as R23-609. The new path is clean; the legacy literals stay in
Module2_DrawingPipeline.bas (1525-1556) because it is the reachable
production path and Module17 is not wired into main.

### Verification

Procedure blocks balanced 21/21 and 4/4. ANSI-only bytes, CRLF, no BOM, max
line 77. 22 Phase 7 contracts. Suite 264 tests with the five known-stale R20
failures. Preflight 31 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 6 read-only gate satisfied

`definitions=3|definitionFailures=None|counterboredFamilies=1|
threadedFamilies=1|shapeFailures=None|nativeCallouts=2|creations=0|
drawingUnchanged=True`.

### Corroboration, not just self-consistency

The M5 family's depth resolved to 12.4 mm from
swCalloutVariable_Tap_Drill_Depth (28), and its thread depth to 10 mm. The
legacy hardcoded string in Module7_TitleBlockEngine.bas reads
"4.2 x 12.4 DEEP" and "TAP M5x0.8-6H x 10 DEEP". The same numbers arrived by
a completely different route - typed callout variables instead of a human
typing them - which is independent evidence that the derivation is correct.

### Both retention branches exercised live

The M5 family kept its native callout
(reason=CompleteAssociativeDefinitionAvailable). The counterbore and
stepped-bore families retained ControlledFallback with
reason=NoNativeCalloutAttributedToFamily. An earlier run also produced
reason=NativeIncomplete|nativeMissing=Depth, so R23-605 is proved in both
directions rather than only the one the fixture happened to take.

### One more mislabel the passing run exposed

threadedFamilies came back as 2. A thread DESCRIPTION was being treated as a
thread, and the counterbored clearance-hole family carries the fastener size
of the screw it clears, with threadDepthM=0. A hole that is actually tapped
has a thread depth. Now 1, which is the truth.

This one mattered less than the others - the gate would have passed either
way, because the M5 family is genuinely tapped. It was still wrong, and a
shape classifier that miscounts is a shape classifier that will eventually
pass something it should not.

### Verification

27 Phase 6 contracts. Suite 242 tests with the five known-stale R20
failures. Preflight 29 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 6: callout reconciliation and controlled fallback

`Module16_CalloutDefinition.bas` (20 procedures), `CCalloutDefinition.cls`.
Statically verified only; no live run yet.

### The shape of a definition

Either a NATIVE SOLIDWORKS hole callout carrying the Hole Wizard's own data,
or a CONTROLLED FALLBACK assembled field by field from the typed feature
data Phase 2 proved. Never free text. Every field carries its proof source,
so a definition that looks complete can still be shown to be unproven, and
R23-610 reports which field failed rather than that something did.

### Contracts worth carrying

- **R23-601: IsHoleCallout is the only classifier.** A native callout
  reports Type2 = 6 (swDiameterDimension) and so does an ordinary diameter
  dimension. No dimension-type constant is declared in the module, so none
  can be reached for.
- **R23-603: fields come from IDisplayDimension.GetHoleCalloutVariables**,
  reading HoleFit, ShaftFit, ToleranceType, ToleranceMin and ToleranceMax
  per ICalloutVariable. Parsing the rendered string would give something
  that cannot be validated field by field, which is the point of the task.
- **R23-602: a callout resolving to two families is rejected**, not
  tie-broken. Attribution is COM identity against every drawing entity a
  projection owns - not the anchor alone, because Phase 4 showed the anchor
  tier prefers the through hole while a counterbore callout attaches to the
  wider mouth.
- **R23-606: quantity is unique physical locations.** Not a feature count:
  one Hole Wizard feature plus a mirror produces many holes. Not an edge
  count: a counterbore contributes several edges per hole.
- **Depth is required only for a blind hole.** Demanding it from a through
  hole would fail every through hole. The stored end-condition code decides,
  and an unproven end condition fails on its own terms first.
- **R23-611 is stated as shapes rather than part numbers**: one multi-hole
  counterbored family and one multi-hole threaded family. P-0251 satisfies
  it; the rule does not name it.

### R23-609 is half met, and the remaining half is deliberate

The new path contains no part number, no 6X, no M5x0.8, no H7, no diameter
literal, and no scoring by feature name or by proximity to an expected
radius. A contract asserts each of those strings is absent.

The legacy literals are still in Module7_TitleBlockEngine.bas - callout text
at 359-371, name/radius scoring at 405-435 - because Module7 is the
reachable production path and Module16 is not yet wired into main. Deleting
them now would degrade the deployable macro while its replacement is
disconnected. They come out in the phase that switches the pipeline over.

### Mutation boundary

CreateNativeCalloutForFamily is the only procedure that creates anything. It
refuses without allowMutation, and refuses again without a proven anchor:
IDrawingDoc.AddHoleCallout2 attaches to whatever edge is selected, so an
unproven selection would produce an associative callout pointing at the
wrong hole and looking correct on the sheet.

### Verification

Procedure blocks balanced 20/20 and 6/6. ANSI-only bytes, CRLF, no BOM, max
line 78. A hygiene contract now asserts every byte is below 0x80, after a
UTF-8 em dash in Module14 broke SWP readback earlier today. 23 Phase 6
contracts. Suite 238 tests with the five known-stale R20 failures. Preflight
29 managed components. MACRO_SOURCE_REVISION unchanged at
target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 5 read-only gate satisfied

Final run: `schemes=4|horizontalSchemes=2|verticalSchemes=2|`
`creditedLocations=10|expectedLocations=10|coverageFailures=None|creations=0|`
`initialSelectionCount=0|finalSelectionCount=0|drawingUnchanged=True`.

### Proven live

- **R23-500** four schemes, keyed by view role + machining face + datum
  policy + direction. The top-face and side-face families separated without
  any view name being read, because machining face comes from the location's
  sign-normalized axis.
- **R23-501** `Drawing View4` horizontal datum is the stepped bore's
  projected centre, chosen by the CentreBoreProjectedCentre policy rather
  than by fallback, and proved selectable with ownership confirmed after the
  fact through ISelectionMgr.GetSelectedObjectsDrawingView2.
- **R23-503 / R23-504** two X buckets and three Y buckets in the primary
  view, every bucket anchored.
- **R23-505** all four side holes credited across two page positions.
- **R23-507** profileEntries=1; the stepped bore never enters a bucket.
- **R23-509** ten small-hole locations, ten credited, no coverage failures.

### Open, and stated as open

- **R23-502 is NOT met.** The vertical datum resolves to the lowest
  projected hole and is recorded as datumKind=ProjectionDerived. The task
  asks for bottom outline geometry, which is a different kind of entity.
  Recording the kind separately from the policy is what keeps this visible
  instead of letting a weaker datum satisfy the requirement quietly.
- **R23-506 half met**: four side locations resolved and credited, none
  dimensioned.
- **R23-508 unrun**: it mutates.

### Two defects the runs found

The probe called Module6_QAEngine.EmitRunEvidence, the production gate whose
RequireCoreStages demands fourteen pipeline stages a read-only probe never
performs. It failed closed and reported RESULT: FAIL for a run that had
resolved every scheme and datum it found.

Then the coverage gate reported credited=8 of an expected 10.
MarkCoincidentProjections sets CoincidentWithAnchoredKey on the UNANCHORED
projection - it explicitly skips anything that already has an anchor - and
the ledger read that field off the projection it was bucketing, which is by
definition the anchored one. The unanchored twin holding the link had
already been filtered out by the Accepted guard, so it was unreachable from
either side. Two of P-0251's four side holes went silently uncredited.

Both are now pinned by contracts, including one that asserts which end of
the coincidence link Module13 writes, so an inversion fails statically
rather than surfacing as a coverage shortfall.

### Verification

25 Phase 5 contracts. Suite 215 tests with the five known-stale R20
failures. Preflight 27 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 5: ordinate schemes and the transaction (source)

`Module15_OrdinateScheme.bas` (36 procedures), `COrdinateScheme.cls`,
`COrdinateBucket.cls`. Statically verified only; no live run yet.

### R23-500, the scheme key

view role + machining face + datum policy + direction. A family says what a
hole IS and nothing about which ordinate group it belongs to: two holes of
one family machined from opposite faces belong to different groups, and two
holes of different families machined from the same face in one view belong
to the same group. Every part of the key is measured - machining face from
the location's sign-normalized axis, view role from Phase 3's axis-normal
measurement - so a renamed or reoriented view still lands correctly.

`swOrdinate` (1) is deliberately never used. Letting SOLIDWORKS infer the
direction from the selected points would make the created dimension depend
on selection order rather than on the scheme.

### R23-505 and R23-509, counted the way the geometry allows

A bucket holds ONE selectable drawing entity and the LIST of physical
locations it represents. Phase 3 proved live that two coaxial holes viewed
along their shared axis produce exactly one drawing entity; dimensioning it
twice is a duplicate, and crediting only one hole silently drops the other.
Coverage is therefore counted per distinct page position and credited to
locations - the finding carried forward from Phase 3.

### R23-507, small-hole membership

Family size, read from the graph. P-0251's stepped bore is excluded because
it is a singleton, not because it is large. A radius threshold would look
equivalent here and misclassify a different part.

### R23-508, the transaction

Activate and verify the view, bind `ISelectData.View` (guarded - it is
documented get/set but raises error 91 on this build), select the datum
first, append in ascending coordinate order, verify the selection count
advanced by exactly one at every append, call `AddOrdinateDimension`, decode
all eleven `swCreateOrdDimError_e` members by name, call `SetPickMode`
whatever the result was, clear selections on every exit, then read back.

### Three defects caught before compiling

- `IsOrdinateEligibleView` and `IsDeferredCreationView` take
  `(graph, swView)`, not `(swView, graph)`.
- **`IView.GetFirstDisplayDimension5` is obsolete and its own Remarks say
  the `GetNext5` walk covers the drawing SHEET.** A read-back built on it
  would credit other views' dimensions to this scheme. Replaced with
  view-scoped `IView.GetDisplayDimensions`.
- A count-difference read-back would be inflated by any unrelated dimension.
  Replaced with a before/after snapshot diffed by `ISldWorks.IsSame`, exact
  `= 1`.

### Mutation boundary

`CreateOrdinateGroup` is the only procedure that creates anything. It
refuses without `allowMutation`, and refuses again when the datum is
unproven - the datum is the first selection and everything is measured from
it. `R23_ProbeOrdinateScheme` contains no `AddOrdinateDimension` call.

### Verification

36/36 and 9/9 and 7/7 procedure blocks balanced, ANSI/CRLF, no BOM, max line
78. 22 Phase 5 contracts. Suite 212 tests with the same five known-stale R20
failures. Preflight 27 managed components. `MACRO_SOURCE_REVISION` unchanged
at `target-spec-hybrid-v2-2026-07-29-r22`.

## 2026-08-01 - R23 Phase 4 gate satisfied; reverse route ruled out on evidence

### What the instrumented run proved

`IModelDocExtension.GetCorrespondingEntity2` returned Nothing for **all 38
annotations and every attachment, with error 0**:
`outcomes=draw1:unresolved:err0`, `resolved=0`, `modelEdgesTested=0`,
`eqMax=-1`. The call declines rather than fails, and because no comparison
ever executed, `eqMax=-1` also rules out the `swObjectEquality` Unsupported
hypothesis raised in the previous entry. That hypothesis is now closed.

The member is documented as returning the entity "in the underlying part or
subassembly". This is a **part** drawing with
`componentContext=DrawingContextOnly` — there is no component to descend
into. The reverse route is therefore unavailable here by construction, not
misused. It is retained (it is the documented direction and should work in
assembly drawings) and now reports
`reverse=UnavailableNoModelCounterpart` instead of an unqualified no-match.

### Reconciliation reclassified, not force-fitted

Unmatched annotations now record
`AuthoredDrawingEntityNoModelCounterpart` rather than
`NoAttachedProjection`, distinguishing "this drawing entity has no model
counterpart at all" from "it has one and no location owns it". Only the
second would implicate the ownership model.

**No positional or dimensional fallback was added.** Matching a dimension to
the nearest hole on the page is the precise failure the physical-location
model exists to prevent, and R23-407 forbids it explicitly.

### Gate

`annotations=38`, `coverageFailures=None`,
`COVERAGE|holeCallouts=2|ordinates=10|diameters=0|toleranced=1|withFit=1`.
R23-412 is required-category coverage — hole callout, toleranced dimension,
ordinate — and all three are present. **Phase 4 gate satisfied.**
`reconciled=1` is `RD1@Drawing View7`, which proves the R23-407 identity
mechanism; the remaining 37 are the designers' own reference dimensions.
Read-only boundary held on all six runs: `mutations=0`,
`drawingUnchanged=True`.

### Verification

33 procedures balanced, ANSI/CRLF, no BOM, max line 79. 36 Phase 4 contracts.
Full suite 190 tests, same five known-stale R20 failures. Preflight 24
managed components. `MACRO_SOURCE_REVISION` unchanged at
`target-spec-hybrid-v2-2026-07-29-r22`.

## 2026-08-01 - R23 Phase 4: reverse route instrumented after it changed nothing

### What the live run showed

The reverse-correspondence route ran on all 38 annotations and matched none.
`reconciled` stayed at 1 of 38, still the single `ForwardAlias` match on
`RD1@Drawing View7`. Every unmatched line carried
`routesTried=ForwardAlias,ReverseCorrespondence`, confirming the route
executed rather than being skipped.

**The evidence could not say why.** Two causes remained live and the log
could not separate them:

1. `GetCorrespondingEntity2` returned Nothing for every attachment.
2. It resolved, and `LocationOwnsModelEntity` rejected every result.

Predicting between them would repeat the mistake that cost the previous
iteration, so this change adds evidence rather than a fix.

### Added instrumentation

`MatchByReverseCorrespondence` now takes `ByRef diagnostics` and reports, per
attachment, the drawing entity's `swSelectType_e`, whether the reverse call
resolved, the trapped error number when it did not, and the resolved model
entity's type. It also totals `resolved`, `projectionsInView`,
`modelEdgesTested`, and `eqMax`.

`eqMax` is the load-bearing addition. **`ISldWorks.IsSame` returns
`swObjectEquality` — 0 NotSame, 1 Same, 2 Unsupported — and the existing
`ObjectsAreSame` wrapper collapses 0 and 2 into `False`.** A comparison that
*cannot be performed* across two documents is therefore indistinguishable
from a genuine non-match, which is exactly the shape of a silent zero.
`RecordEquality` keeps the raw code and reports the strongest one seen.

### Standing observations from the same run, not yet acted on

- `Section View J-J` reports its drawing component as
  `P-0251-14A-001-SectionAssembly-3-1/P-0251-14A-001-1`. Section views are
  built on a temporary section assembly, and section-cut edges have no
  counterpart in the part body. If the H7's two attached edges are cut
  geometry, no reverse mapping can exist for them. **Unverified** — the new
  `outcomes` field will show it directly.
- Several ordinate attachments report entity type `0` (`swSelNOTHING`), which
  is not valid input to `GetCorrespondingEntity2` (it documents vertex, face
  or edge). `SafeEntityType` returns `-1` on failure, so `0` is a genuine
  reading, not a trapped error.

### Verification

33 procedures balanced, ANSI/CRLF, no BOM, max line 79. 35 Phase 4 contracts
pass (2 new). Full suite 189 tests with the same five known-stale R20
failures. Preflight resolves 24 managed components. `MACRO_SOURCE_REVISION`
unchanged at `target-spec-hybrid-v2-2026-07-29-r22` — Phase 4 is not yet
wired into the deployable path.

## 2026-08-01 - R23 Phase 4: reverse correspondence route added (R23-302)

The stage instrumentation isolated the reconciliation failure in one run.

### Cause

**The counterbore hole callout attaches to a drawing edge of type
`swSelEDGES` that is none of the 18 aliases `IView.GetCorrespondingEntity`
produced for that view.** The forward model-to-drawing map is partial — it
returns 2 of each location's 4 boundary edges — so no amount of alias
comparison could ever reach the callout's edge. `Section View J-J` reported
`anchoredProjections=0|aliasesAvailable=0`, meaning its seven dimensions,
the H7 among them, were structurally unreachable by the forward route.

This was a design gap rather than a coding slip: **R23-302 asks for
drawing-to-model mapping and only the forward direction had been built.**

### Added

`MatchByReverseCorrespondence` uses
`IModelDocExtension.GetCorrespondingEntity2`, which the 2025 Help names for
resolving a drawing entity to its part entity. Each attached drawing entity
is mapped back to the model, then tested against the geometry each physical
location owns — its `SourceFaces` and those faces' boundary edges, via
`LocationOwnsModelEntity`.

The reverse route requires no projection anchor, so it can reach views where
the forward map produced nothing, including the section view. It remains
identity-only throughout (`ISldWorks.IsSame`), with no positional or
dimensional fallback. Successful matches record
`matchRoute=ForwardAlias` or `ReverseCorrespondence`, and unmatched
annotations report `routesTried=ForwardAlias,ReverseCorrespondence`, so the
two routes are never conflated in evidence.

### API contract

`IModelDocExtension.GetCorrespondingEntity2(Entity)` takes a vertex, face or
edge in a drawing view or assembly and returns the corresponding entity in
the underlying part, or Nothing. It is the documented counterpart to
`IView.GetCorrespondingEntity`.

## 2026-08-01 - R23 Phase 4 third live run: tolerance policy confirmed

Read-only boundary held: `mutations=0`, `finalSelectionCount=0`,
`drawingUnchanged=True`.

### Confirmed

The tolerance policy is in force and visible. The H7 reports
`toleranceAuthority=DrawingAuthoredNonAuthoritative`; every other dimension
reports `NoTolerance`. `dowelConvention` is gone from the evidence, and the
±10 µm constants are gone from the source.

### Not fixed: reconciliation is still 1 of 38

Matching every drawing entity a projection owns rather than only its anchor
was expected to raise the reconciled count to 3-4. **It did not move it.**
The prediction was wrong and the cause is still unknown.

Rather than attempt a second guess, `ANNOTATION_UNMATCHED` now reports what
each attachment actually is: `attachmentTypes` as `swSelectType_e` codes
(1 edge, 2 face, 3 vertex, 10 sketch segment, 11 sketch point, 28 centre
mark, 46 silhouette), together with `anchoredProjections` and
`aliasesAvailable` for the view. One run will now distinguish "the callout
attaches to an unmapped edge", "it attaches to a sketch segment", and "it
attaches to something that is not an edge at all".

Note the counterbore callout `RD3@Drawing View4` has `attachments=1` while
its projection has two mapped aliases, so the attachment is a single entity
that is not either mapped edge. That is the specific thing the next run
identifies.

## 2026-08-01 - R23 Phase 4 second live run; drawing tolerances demoted

Both fixes confirmed. The read-only boundary held again: `mutations=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

**`IDimension.SystemValue` is definitively correct for drawing dimensions.**
All 30 display dimensions now report
`nominalSource=IDimension.SystemValue=<value>|GetSystemValue3=0.000000000`.
Every value matches the reference drawing: 47.0 H7, 11.0 and 4.2 callouts,
R36, 173.6, 104.8, 80, 40, 25, 18, 12, 11.5, 6, and the ordinate chains.
`attachments` now reads 1-3 instead of 0. The H7 reads in full: nominal
`0.047`, `toleranceType=8`, `tolMaxM=0.000025`, `tolMinM=0.000000`,
`holeFit=H7`.

### Changed: drawing tolerances are not authoritative

Standing instruction from the user: the tolerances in the designers'
existing drawings were added manually to signal that *some* tolerance is
acceptable, not to state that the part holds them. They are evidence that a
designer typed a number, nothing more.

Every tolerance is now labelled
`toleranceAuthority=DrawingAuthoredNonAuthoritative`.
`ClassifyToleranceAuthority` deliberately has **no branch returning a
stronger authority**: nothing R23 can read distinguishes a binding tolerance
from an indicative one, and inventing a distinction would manufacture
authority that does not exist. The dowel ±10 µm rule and its constants are
**removed rather than parked** — the user is establishing with their designer
what part information should drive the decision to add a tolerance, and
nothing here guesses at it.

### Fixed

**Reconciliation matched only the anchor, so 1 of 38 annotations
reconciled.** A counterbore maps two edges; the native hole callout attaches
to the 11 mm mouth while the R23-308 anchor tier deliberately prefers the
6.6 mm through hole. They are different drawing entities of the same physical
location, so identity against the anchor alone failed. `ProjectionOwnsEntity`
now tests the anchor **and** every mapped alias the projection recorded.
Ownership is what reconciliation needs; the anchor is only the preferred
attachment point for new annotations. Identity only — still no positional
fallback anywhere.

**`alias` is a VBA reserved word** (`Declare ... Alias`) and would not have
compiled; caught by a new contract that also guards `name` and `type`.

## 2026-08-01 - R23 Phase 4 first live run: H7 provenance settled

The read-only reconciliation run held its boundary exactly: `mutations=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`. The manual reference
drawing was not altered. 38 annotations inventoried across four views,
`coverageFailures=None`, both hole callouts found.

### The H7 question is answered

`Section View J-J` carries `RD4@Drawing View6@P-0251-14A-001.Drawing` with
`toleranceType=8` (`swTolFITWITHTOL`), `nonZeroTolerance=True` and
`fitData=True`. **The Ø47 H7 fit is authored in the drawing, not the model.**
Phase 0 read `D1@Sketch4` directly and found no H7 and no nonzero tolerance
there; the fit exists only as a drawing-level reference dimension. R23 must
not expect model import to supply it, and must not record it with model
provenance when it does supply it.

### Dowel tolerance convention recorded

The user supplied the matching domain rule: **a dowel hole receives a
designer-added ±10 µm tolerance in the 2D drawing**, never in the model. Same
pattern as the H7 — precision tolerance is a drafting-stage act on this
drawing set. `MatchesDowelToleranceConvention` recognizes the ±0.010 mm
symmetric signature so provenance can follow from it. Recognition is
deliberately **not** treated as proof that a hole is a dowel: nothing in the
model says so, and P-0251's holes are counterbores, tapped holes and a
stepped bore, none proven to be dowels.

### Fixed

- **`nominalM` read 0 for every dimension** while `nominalAvailable` was
  True. `IDimension.GetSystemValue3` is configuration-scoped and these are
  *drawing* dimensions (`RD4@Drawing View6@....Drawing`); a drawing document
  has no configurations, so the read had nothing to resolve. Now taken from
  the configuration-free `IDimension.SystemValue`, with both readings kept in
  `NominalSource` so the disagreement stays visible.
- **`attachments=0` on every ANNOTATION line**, including the one that went
  on to reconcile successfully — the count was populated during the later
  reconciliation pass, after the line had printed. Attachments are now read
  during inventory. Evidence must never be printed before the field it
  reports is populated.

### Established by the run

- **The reference drawing authors its section diameters, H7 included, as
  `swLinearDimension` (2)**, not `swDiameterDimension`. Dimension type alone
  can never decide whether something is a diameter. Phase 0 separately showed
  *imported* section diameters arrive as type 6, so the two authoring routes
  differ.
- Annotation type 13 (`swCenterMarkSym`) accounts for the six `Other:13`
  entries in `Drawing View4`; now named.
- 38 annotations: 10 ordinates, 2 native hole callouts, 1 toleranced with
  fit, 5 cosmetic threads, 8 center marks, 1 note.

## 2026-08-01 - R23 Phase 4: annotation import and reconciliation

### Added

`src/target-spec-hybrid-v2/Module14_AnnotationImport.bas` (26 procedures)
implements R23-400 to R23-412.

**Mutation boundary.** This is the first R23 phase whose operations change a
drawing. Only `ImportModelAnnotations` and `RemoveR23CreatedAnnotations`
mutate, and both refuse unless passed an explicit `allowMutation` argument.
`R23_ProbeAnnotationReconciliation` never passes it and contains no insert,
delete or save call, so it is safe against the manual reference drawing. The
read-only path still exercises R23-406 to R23-409, R23-411 and R23-412 on
real data, because the reference drawing already carries the manufacturing
intent R23 must reproduce.

Reconciliation is by **attached-entity COM identity** against each
projection's proven anchor. Page proximity is used nowhere: it would attach a
dimension to whichever hole happened to be nearest, which is the failure the
physical-location model exists to prevent. R23-410 likewise deletes only this
run's own recorded annotation objects, never matching by name, position or
appearance.

Ordinate eligibility is decided from the graph's measured `AxisNormalToView`
data rather than from view names, so the isometric is excluded because no
hole axis is normal to it — the fact Phase 3 established — and a renamed or
reoriented view is still classified correctly.

### API contracts established

- **`swInsertAnnotation_e` mask `18055274` fully decomposed**, no unaccounted
  bit: datums 2, dimensions 8, GTols 32, notes 64, marked-for-drawing 32768,
  hole-wizard profile 65536, hole-wizard location 131072, hole callout
  1048576, toleranced dims 16777216. The callout member really is spelled
  `swInsertholeCallout`.
- **`InsertModelAnnotations4` is on `IDrawingDoc`, takes eight arguments and
  returns an ARRAY of `IAnnotation`**, not a count. There is no
  `InsertModelAnnotations3` on `IModelDocExtension`. Returning the objects is
  what makes the cleanup path safe.
- **`GetMinValue2`/`GetMaxValue2` return the STATUS**
  (`swDimensionToleranceWarning_e`) and deliver the value through an out
  parameter, which must be a local — the Phase 3 ByRef trap again.
- **`swTolFIT` and `swTolMETRIC` are both 7**, so a 7 is recorded as
  ambiguous rather than reported as one of them.
- `IAnnotation`, not `IDisplayDimension`, carries `GetAttachedEntities3`.
- `IDimension.GetSystemValue2` is obsolete; `GetSystemValue3` is used.

### Changed

- The deployment manifest now manages 24 components; the companion inventory
  lock moved from 23 to 24.

### Tests

Added `tests/test_r23_annotation_import_contracts.py`: 22 contracts, led by
the mutation-safety ones — the probe contains no insert, delete or save call;
both mutating procedures refuse without authorization; cleanup never matches
by name, position or appearance. Also asserts identity-based reconciliation
with no proximity fallback, tolerance out-parameters bound to locals, and the
mask decomposition. The suite is now 176 tests with the same five stale R20
failures.

### Not done

No live run. `MACRO_SOURCE_REVISION` remains `r22`: still no pipeline caller.

## 2026-08-01 - R23 Phase 3 gate SATISFIED; coincident projections explained

The fifth live run settled the remaining question, and disproved the
hypothesis recorded after the fourth.

**The hidden-lines explanation was wrong.** `Drawing View7` reports
`displayMode=HiddenLinesVisible`, so the far-face holes ARE drawn. The real
reason is geometric: that view looks along model Y, and the four tapped holes
lie on two axes with two holes each, so they project onto **two** page
points. The six counterbores likewise collapse to **three**. The observed
`mappedEdges` counts — 2 and 3 — equal the number of distinct page positions
exactly. Two coaxial holes seen along their axis are ONE circle on the sheet;
SOLIDWORKS holds a single drawing entity and `GetCorrespondingEntity` maps
only one of the two model edges to it. No search strategy can produce more
anchors than the drawing has entities.

**R23-310's second clause is therefore unsatisfiable as written, and the
implementation is correct.** Two side anchors is the complete answer. The
individual identity of all four tapped holes survives in `Drawing View4`,
which resolves them to four distinct page positions edge-on, and the
reference drawing calls them out once as `4x Ø4.2 ▼12.4 / M5x0.8 - 6H ▼10`.

### Added

`MarkCoincidentProjections` attributes every unanchored projection that
shares a page point with an anchored one, emitting `PROJECTION_COINCIDENT`
with `sharesPagePointWith` and
`reason=OneDrawingEntityForTwoCoaxialHoles`, and counting them as
`coincidentUnanchored` in the per-view summary. The page centre is now
recorded from the first face that projects whether or not an anchor is
found, since an unanchored location still has a provable position and
coincidence cannot otherwise be distinguished from a defect. Acceptance is
unaffected — the anchor test runs before the coordinate-frame test.

### Confirmed

`projectedAxis` now carries real values and corroborates the transform
independently: `0,0,1` for Z-axis holes in the front view, `0,0,-1` for
Y-axis holes in the side view, `1,0,0` and `-1,0,0` for the in-plane cases,
and `0.707107,0.408204,-0.577382` in the isometric — exactly
(1/√2, 1/√6, −1/√3), magnitude 1.000.

`IView.GetDisplayMode2` per view on this drawing: `Drawing View4` and
`Drawing View7` are `HiddenLinesVisible`, `Drawing View2` and
`Section View J-J` are `HiddenLinesRemoved`, and all ten sheet placeholders
report `Shaded`.

### Phase 3 gate

**Satisfied.** R23-300 through R23-310 are all proved live on P-0251, with
`selectionProved=9`, `finalSelectionCount=0`, `selectionClean=True`,
`drawingUnchanged=True`, `partUnchanged=True`. `MACRO_SOURCE_REVISION`
remains `r22`: the engine still has no pipeline caller.

**Carry into Phase 5:** required coverage must be counted per *distinct page
position per view*, not per physical location.

## 2026-08-01 - R23-310 fourth live run: primary-view clause proved

The axis gate behaved exactly as designed. Per-view acceptance:

| view | type | projections | axisNormal | anchored | accepted |
|---|---|---|---|---|---|
| `Drawing View4` (primary) | 4 | 11 | 7 | 11 | 7 |
| `Drawing View7` (side) | 4 | 11 | 4 | 6 | 2 |
| `Section View J-J` | 2 | 11 | 4 | 0 | 0 |
| `Drawing View2` (isometric) | 4 | 11 | 0 | 9 | 0 |

**R23-310's first clause is proved:** `Drawing View4` accepted all six
counterbores plus the stepped bore, every one anchored at
`anchorTier=PrimaryTypedHoleDiameter` on the Ø6.6 through hole rather than
the Ø11 mouth, and all seven proved selectable with `ownershipProven=True`,
`selectionClean=True`, both documents unchanged.

The gate's own correctness is corroborated by the axisNormal counts: 7 for
the Z-axis holes in the front view, 4 for the Y-axis tapped holes in both the
side and section views, and **0 in the isometric**, where no hole axis is
normal to the sheet. The isometric had 9 anchors and correctly accepted none.

### Fixed

**ByRef out-parameters were class fields, so the write-back was lost.**
`projection.ProjectedAxisX/Y/Z` were passed directly to
`AxisIsNormalToView`. A class Public variable is exposed as a property, so
VBA hands the callee a temporary that is discarded on return: the run logged
`projectedAxis=0.000000,0.000000,0.000000` on every line while `axisNormal`
was correct, because the function's own locals were sound and only the
write-back vanished. The values now go into locals and are assigned
explicitly. An audit found no other genuine out-parameter of this shape; the
other class-field arguments in Module12 and Module13 are input-only.

### Added

`IView.GetDisplayMode2` is now recorded per view as `displayMode=`. Under
`swHIDDEN` (Hidden Lines Removed) a far-side hole is never drawn, so
`GetCorrespondingEntity` has nothing to return and the location cannot anchor
in that view however the search is written. Recording it means an
unanchorable location can be attributed to the drawing's display setting
rather than mistaken for a mapping defect.

### Not done

**R23-310's second clause is not met: only two of the four tapped holes have
side projections.** The axis test is right — all four are `axisNormal` in
both the side and section views — but the two far-face holes return Nothing
from route A, and the section view maps nothing at all. Whether that is a
display-mode consequence or a genuine limit is what the next run settles.

## 2026-08-01 - R23-310 third live run: anchors resolve; axis gate added

The `GetCurveParams3` fix worked. `circularEdges` went from 0 to 4 (or 2) on
every location, `uMin=0`, `uMax=6.283185307`, `closureM=0.000000000` — the
Phase 0 circle proof reproduced exactly. **26 projections anchored**, every
one through route A with route C identity confirmation, and all 26 proved
selectable: `selection=True`, `ownershipProven=True`, `selectedCount=1` on
each, `initialSelectionCount=0`, `finalSelectionCount=0`,
`selectionClean=True`, `drawingUnchanged=True`, `partUnchanged=True`.

`ISelectData.View` raised error 91 on every attempt, as this repository
already recorded. The guarded binding plus after-the-fact ownership proof
handled it: `viewBinding=UnboundAfterError:91` with
`ownershipProven=True` 26 times out of 26.

### Fixed

**Normal-axis compatibility was computed but never enforced.** R23-306
requires it; `AxisNormalToView` was stored on the projection and ignored by
`QualificationFailureReason`. The run accepted all four M5 tapped holes in
`Drawing View4`, whose page coordinates prove that view's normal is model Z:
the counterbores land 40 mm and 30 mm apart, matching model X and Y, while
the M5 holes land at their model **Y** spread horizontally — their axis lies
in the page plane and the "circular" edge is seen edge-on. Such a projection
is not a usable circular anchor. It now fails as `AxisNotNormalToView`,
distinct from `ProjectionAnchorUnavailable` because the mapped anchor is
real; the verifier counts it separately rather than folding it into the
unanchored total.

**R23-310 could not be answered by a single number.** The requirement is
about which holes are usable in *which* view. Added `ViewAcceptanceSummary`
and an `R23_PROJECTION_VIEW_SUMMARY` line per view carrying `projections`,
`axisNormal`, `anchored` and `accepted`. `axisNormal` and `projectedAxis` are
now in the per-location evidence too, so the gate's own correctness is
visible rather than assumed.

### Established by the run

- Route A works from a model reached through `IView.ReferencedDocument` with
  the drawing active, settling the open question from the previous run. The
  earlier hypothesis about needing the part as active document was wrong.
- Mapping remains per-edge and partial: `mappedEdges` ranged 0 to 4 out of
  4 candidate edges on the same location in different views.
- `Section View J-J` mapped **nothing** (`mappedEdges=0` on all 11
  locations). Its visible component is a synthetic section assembly, so the
  original part's model edges have no correspondent. A section view needs a
  different acquisition route.
- Hidden geometry does not map: in the side view only the near column of
  counterbores (`py=-0.015`) and the near pair of tapped holes resolved.

### Not done

R23-310 stays open. With the axis gate applied the counts will change, and
whether "side projections for all four tapped holes" is achievable at all is
now in doubt — two of the four are on the far face and did not map.

## 2026-08-01 - R23-310 second live run: root cause found and fixed

The stage counters isolated the failure in one run. Every location reported
`sourceFaces=2, facesProjected=2, boundaryEdges=4, circularEdges=0` with
`firstReject=circle=Reject|reason=ReadError:438`, identically across all
three real views. Face retention, the page transform and edge enumeration
all work; every edge died at the same call.

### Fixed

**`IEdge.GetCurveParams3` returns an `ICurveParamData` object, not an array
of doubles.** `Module13_ProjectionResolution` assigned it into a Variant with
a `Let`, which asks the object for a default property it does not have and
raises error 438 — rejecting every candidate edge before any mapping was
attempted. The correct pattern was already in this repository:
`Module12_FeatureQualification.ComputeFaceAxialInterval` binds it with `Set`
and reads `.StartPoint` / `.EndPoint`, which is why the Phase 2 axial
intervals were always sound. Module13 now matches it and additionally
records `UMinValue` / `UMaxValue` in the circle proof.

**Unmeasurable closure failed open.** The new `PointDistance` helper reports
failure as a negative value, which would have passed the
`closureM > tolerance` test and read as a perfectly closed curve. It is now
rejected explicitly as `ClosureNotMeasurable` ahead of the tolerance test.

**Selection cleanliness had no baseline.** The run reported
`finalSelectionCount=1` while proving no anchors and selecting nothing —
the count was the operator's own pre-existing selection. The probe now
captures `initialSelectionCount` and reports `selectionAttempted` plus
`selectionClean`, so a non-zero count is only attributed to this pass when
this pass actually selected something.

An audit of every other `Let` assignment from an API member in Module13
found no further object-returning members; `GetCurveParams3` was the only
one.

### Established by the run

- `swDrawingViewTypes_e` on this build: the ten sheet placeholders report
  type `7` (`swDrawingNamedView`) with zero visible entities; the real
  projected views report `4` and the section view `2`.
- `ISheet.GetViews` returned 14 views, ten of them placeholders. The
  visible-entity skip removed 110 dead projections, taking the run from 154
  to 44.
- `Drawing View2` (65 edges) is a fourth real view, not seen in the first
  run's truncated log.

### Not done

No anchor has yet been proved. R23-310 stays open pending the rerun.

## 2026-08-01 - R23-310 first live run: no anchors; chain instrumented

The first `R23_ProbeViewProjections` run compiled and executed read-only
against the P-0251 drawing. The safety envelope held exactly:
`drawingUnchanged=True`, `partUnchanged=True`, `finalSelectionCount=0`, no
warnings, no failures. **No anchor resolved** — all 154 projections failed
`ProjectionAnchorUnavailable` with `candidates=0`.

Established by the run:

- the drawing side is sound — `GetVisibleEntities2` returned 64, 68 and 53
  edges for `Drawing View4`, `Section View J-J` and `Drawing View7`, each
  with a drawing component;
- `componentContext=DrawingContextOnly` on every real view, confirming the
  predicted part-drawing behaviour of `GetVisibleDrawingComponents`;
- `ISheet.GetViews` returns the sheet's six standard-view placeholders
  alongside real views, and they hold no drawing geometry.

### Fixed

**The evidence could not isolate the failure.** `candidates=0` is a single
number covering "no faces retained", "face centre not projected", "no
circular edge" and "nothing mapped" alike. This is the defect class Phase 0
solved with `rejectGate`, and it should not have recurred.
`ResolveProjection` now counts each stage — `sourceFaces`,
`facesProjected`, `boundaryEdges`, `circularEdges`, `mappedEdges`,
`inventoryConfirmed` — and carries the transform's own proof string into
`firstReject` rather than discarding it.

**Placeholder views generated dead projections.** A view with no visible
entities cannot supply an anchor, so it is now skipped with a single
`PROJECTION_VIEW_SKIPPED` line. 132 of the 154 projections were noise of
this kind. The test is the visible-entity count, which is a fact; the view
type is logged for evidence but deliberately not used as the filter, because
which `swDrawingViewTypes_e` code the placeholders carry is not established
on this build.

### Not done

The root cause of the zero anchors is **not yet diagnosed** and is not
guessed at here. R23-310 stays open pending the instrumented rerun.

## 2026-08-01 - R23 Phase 3: drawing-projection resolution

### Added

`src/target-spec-hybrid-v2/Module13_ProjectionResolution.bas` (27
procedures) turns each physical location into a `CViewHoleProjection` for one
drawing view: a selectable drawing-context anchor, a page-frame centre, and
the proofs behind both. It uses the routes Phase 0 settled on this build:

- route A, `IView.GetCorrespondingEntity`, is the acquisition path;
- route B, `IComponent2.GetCorrespondingEntity`, is attempted and recorded so
  a future build that fixes it is detected, but nothing depends on it;
- route C, `GetVisibleEntities2` plus `ISldWorks.IsSame`, is not an
  acquisition route at all — it is the independent confirmation that route
  A's result really is a drawing-context entity of this view.

Mapping is per-edge, so every boundary edge of every contributing face is
tried. `CPhysicalHoleLocation` gained `SourceFaces` (merged on consolidation)
so a projection can reach those edges without re-walking the feature tree.

Anchor choice follows the R23-308 priority order and ranks every candidate
rather than stopping at the first mappable edge.

`SelectAnchorInView` is the only code that selects anything; it clears the
selection before and after, guards the error-91 `ISelectData.View` binding,
and proves view ownership afterwards rather than assuming it.

### Contracts established

- **`ISldWorks.IsSame` is not a Boolean.** It returns `swObjectEquality`
  = {0 not same, 1 same, 2 unable to determine}. `NormalizeSwBoolean` would
  have accepted 2 as identity. Only an exact 1 is accepted; unreadable
  comparisons default to not-same and are counted in the evidence.
- **`GetVisibleDrawingComponents` is documented for assembly drawings**, so a
  part drawing yielding only the `GetVisibleComponents` handle is recorded as
  `DrawingContextOnly` rather than treated as a failure.
- **Direction vectors are differenced, not transformed** — applying a point
  transform to a direction folds in the view translation.

### Changed

- The deployment manifest now manages 23 components; the companion inventory
  lock moved from 22 to 23.

### Tests

Added `tests/test_r23_projection_resolution_contracts.py`: 20 static
contracts covering R23-300 through R23-310, including that the identity test
never runs through the Boolean normalizer, that the anchor loop has no early
exit, that unanchored locations are still added to the graph, and that the
pass contains no mutating call. The suite is now 142 tests with the same five
stale R20 failures.

### Not done

No live run. R23-310 is open: `R23_ProbeViewProjections` must confirm usable
primary projections for the six face holes and side projections for the four
tapped holes. `MACRO_SOURCE_REVISION` remains `r22` — still no pipeline
caller, drawing output unchanged.

## 2026-08-01 - R23-213 closed: P-0251 catalog proved live

The second `R23_ProbeFeatureCatalog` run returned `catalogFailures=None` with
all four R23-213 expectations met, 0 warnings, 0 failures and
`modelUnchanged=True`. `Mirror1` was accepted with
`seed=M5x0.8 Tapped Hole1`, `readStatus=SeedSemanticsInherited:Read` and
`thread=M5x0.8`, and the four M5 locations formed one family. The M6
counterbore reads `fit=Normal`; the tapped holes carry no fit. **The Phase 2
gate is satisfied.**

### Fixed

**Traversal was not exact-once.** Comparing the two runs, `visitedFeatures`
went 47 → 46 on an unchanged model, and which sketches were visited twice
differed between runs (`Sketch1`/`Sketch6`/`Sketch7` doubled in the first,
`Sketch4`/`Sketch7` in the second). `ObjPtr` was part of the traversal key,
so a feature reached both as a tree entry and as a consuming feature's
subfeature arrived through two COM wrappers at different addresses and
occupied two keys. The guard was too permissive, which is the opposite of the
failure it was written to prevent.

Neither catalog was affected — accepted features, locations, consolidations
and families were identical across both runs — but the defect is not
cosmetic. `CLocationGraph.AddFeatureDefinition` appends to `mFeatures`
unconditionally, and a repeated visit to an accepted feature would call
`AddStackMember` again, adding a duplicate radius and a duplicate stack
member so a counterbore would report `stackMembers=4`.

The key is now name plus type, which is exact-once because feature names are
unique within a part document. The address survives only as a fallback for a
feature whose name cannot be read, where a repeat visit is preferable to
collapsing several unnamed features into one key and skipping them.

### Not done

The traversal fix is static-only; it needs no dedicated run and can be
confirmed by a stable, duplicate-free `visitedFeatures` on the next live run.
`MACRO_SOURCE_REVISION` remains `r22`: the engine still has no pipeline
caller and drawing output is unchanged.

## 2026-08-01 - R23-213 first live catalog run; three defects found and fixed

The first read-only `R23_ProbeFeatureCatalog` run on P-0251 compiled and
executed cleanly (47 features visited, 0 warnings, 0 failures,
`modelUnchanged=True`) and reported `catalogFailures=NoFourLocationFamily`.
Three defects were found: the reported one, plus two the log exposed on
inspection.

### Fixed

**Pattern instances did not inherit their seed's semantics.** `Mirror1` was
accepted with `operation=PatternInstance` and every manufacturing field
blank, so its two mirrored M5 locations formed their own empty family. The
four identical M5x0.8 tapped holes were catalogued as two families of two and
the expected four-location family never formed. `ResolvePatternSeed` now
reads the seed through the same typed readers that qualify it in its own
right, and `CopySeedSemantics` carries the typed values across. Proof sources
are prefixed `SeedInherited(<seed>):` rather than reasserted, and a field the
seed could not prove stays unproven on the instance. A seed that is itself a
pattern is refused as `SeedIsPatternChainUnsupported`: following the chain
needs recursion whose termination this build gives no evidence for, and
P-0251 has no nested pattern.

**Seed-chain rejections were computed but never enforced.** `ResolvePatternSeed`
set `SeedChainUnresolved`, `SeedChainMultiplyResolved` or `SeedChainCircular`
and returned, but `QualifyFeature` never checked the result and the accept
path then cleared `RejectionReason`. Every one of those codes was dead: a
pattern with an unresolvable or circular seed would have been accepted
silently. It is now a `Function` whose result gates the accept path. The
existing tests asserted only that the reason strings appeared in the source,
which is why they passed; the new test asserts the check precedes the clear.

**`HoleFit` published a not-applicable code as a value.** The M5 tapped hole
returned `-1`, which was stringified into `FitDescription` and given a proof
source, putting a bogus clearance on a tapped-hole callout. The 2025 Help
limits the property to counterbore and countersink features and defines the
return as `swWzdHoleScrewClearanceTypes_e` = {0 close, 1 normal, 2 loose}.
`ScrewClearanceText` now maps those three and returns empty for anything
else, which the caller records as absent. The M6 counterbore's `1` is
genuine and now reads `Normal`.

`FEATURE_ACCEPTED` gained `seed=` and `readStatus=` so seed inheritance is
visible in the next log rather than inferred.

### Tests

Five new contracts (28 in the Phase 2 file, 122 in the suite, same five stale
R20 failures): seed rejections enforced ahead of the accept path, seed
semantics copied field by field, inherited proof sources carried and never
invented, seed chains refused rather than followed, and out-of-enum
`HoleFit` codes treated as absent.

### Not done

R23-213 stays open until a second live run confirms the corrected catalog.
`MACRO_SOURCE_REVISION` remains `r22`: the engine still has no pipeline
caller and drawing output is unchanged.

## 2026-07-31 - R23 Phase 2: feature qualification engine

### Added

`src/target-spec-hybrid-v2/Module12_FeatureQualification.bas` (28
procedures) walks the feature tree and populates a `CLocationGraph`:

- three-field type resolution, normalizing `GetTypeName2 = "ICE"` through
  `GetTypeName` and failing with a field-specific reason when unresolved;
- traversal of features and subfeatures behind a composite cycle guard
  (name plus type plus pointer), because pointer-only guards collapsed
  distinct features in an earlier probe;
- suppression proved against the drawing view's exact
  `ReferencedConfiguration`, with an active-configuration fallback permitted
  only when the two names are equal;
- typed readers for Hole Wizard, advanced holes, simple holes, exact
  `CUT`/`CUTTHIN` extrudes, cosmetic threads and pattern/mirror seeds, each
  pairing `AccessSelections` with `ReleaseSelectionAccess` through one shared
  release path;
- ownership from `IFeature.GetFaces`, with `IFace2.GetFeature` and
  `FaceInSurfaceSense` both absent from the executable path;
- physical locations built from each owned cylindrical face, with the axial
  interval measured from the face's own boundary edges so consolidation and
  separation are decided on real geometry;
- explicit reason codes for every rejection; and
- `VerifyExpectedCatalog` plus a read-only, fixture-guarded
  `R23_ProbeFeatureCatalog` evidence entry point.

`Module11_GeometryIdentity` gained `NormalizeSwBoolean`, the shared
`(CDbl(raw) <> 0#)` rule, now used for every SOLIDWORKS Boolean in the
production path.

### Changed

- The deployment manifest now manages 22 components.
- The companion inventory lock moved from 21 to 22.

### Tests

Added `tests/test_r23_feature_qualification_contracts.py`: 23 static
contracts covering R23-200 through R23-213, including that the module never
references `ModifyDefinition`, `FaceInSurfaceSense` or `IFace2.GetFeature` in
executable code, that consolidation is delegated to the graph rather than
reimplemented, and that the pipeline is not yet rewired. The suite is now
117 tests with the same five stale R20 failures.

### Fixed

`R23_ProbeFeatureCatalog` failed to compile in the VBE with "Method or data
member not found" on `evidence.InfoCount`. `CRunEvidence` keeps `mInfo`,
`mWarnings` and `mFailures` Private and exposes no item accessor, so the
replay loop could never have compiled. It was also redundant:
`CRunEvidence.AddInfo`/`AddWarning`/`AddFailure` each `Debug.Print` at the
moment of emission, and `EmitInfo`/`EmitWarning`/`EmitFailure` print again
under `mEmitDiagnostics`. The loop is replaced by an
`R23_CATALOG_EVIDENCE|warnings=…|failures=…` tally line built from the
public `WarningCount` and `FailureCount` properties. No qualification logic
changed. Every other cross-module reference in the seven new Phase 1/2
components was then checked against the target's public surface, and every
call site checked against its signature; no further mismatches exist.

### Not done

`MACRO_SOURCE_REVISION` remains `r22`: the engine has no pipeline caller yet,
so drawing output is unchanged. R23-213 stays open until one live
`R23_ProbeFeatureCatalog` run confirms the P-0251 catalog.

## 2026-07-31 - R23 Phase 1: location-graph model added (additive, no behaviour change)

### Added

Six new managed components under `src/target-spec-hybrid-v2/`:

- `CFeatureDefinition.cls` — one resolved feature, with a proof-source
  string beside every semantic manufacturing field.
- `CPhysicalHoleLocation.cls` — one physical location, identified by
  sign-normalized axis, line moment and axial interval.
- `CViewHoleProjection.cls` — one location projected into one view, with an
  explicit page-frame proof and the anchor route that produced it.
- `CImportedAnnotation.cls` — imported annotation identity and provenance,
  keeping reference-supplied tolerances distinguishable from model data.
- `CLocationGraph.cls` — nine indexes; the single consolidation point.
- `Module11_GeometryIdentity.bas` — canonical numeric normalization, with no
  SOLIDWORKS API calls so its invariants are checkable offline.

The three physical-identity invariants are enforced in one place,
`CLocationGraph.ResolveOrCreatePhysicalLocation`: coaxial steps consolidate
when their axial intervals meet; opposite blind holes on one infinite line
stay separate because theirs do not; and feature names appear in neither the
physical nor the family key.

### Changed

- `tools/swp-deploy/deployment-manifest.json` now manages 21 components,
  up from 15. `ThisLibrary` and both UserForms remain excluded.

### Tests

Added `tests/test_r23_location_graph_contracts.py` to the companion suite —
20 static source-contract tests for R23-106. The existing
`test_manifest_manages_only_replaceable_components` inventory lock was
updated from 15 to 21 components; all of its other guarantees are unchanged.
The suite is now 94 tests with the same five stale R20 failures.

### Not done

No deployment, no VBE compilation, no live run. `MACRO_SOURCE_REVISION`
stays at `r22` because no deployable behaviour changed — the new classes have
no callers, which R23-107 requires until Phase 2 migrates them. Runtime
output is identical to r22.

## 2026-07-31 - Phase 0 closed: datum-first ordinate transaction proved

### Live evidence

Both ordinate groups completed on the authorized P-0251 fixture, closing the
last outstanding Phase 0 gate.

- `AddOrdinateDimension` returned `0 Success` for both directions.
- Exact selection counts: datum + 2 = 3 (X), datum + 3 = 4 (Y). The X datum
  selected as an edge (type 1), the Y datum as a vertex (type 3).
- Display-dimension deltas +2 and +3; `SetPickMode` called and zero
  selections remaining on both exits; fixture unchanged and drawing unsaved.
- Every selection reported `ownerView=Drawing View1`.
- Ordinate values are semantically correct: `+15.00`/`-15.00` about the
  stepped-bore centre and `10.00`/`50.00`/`90.00` from the bottom-left vertex
  datum — the reference scheme of symmetric centre-zero X and bottom-zero Y.
- Five created dimensions for six locations is the designed deduplication:
  six locations collapse to two unique X and three unique Y coordinates.

Two contracts were settled for production:

- `ISelectData.View` assignment raises error 91 on this build
  (`viewBinding=UnboundAfterError:91`), and both groups completed normally
  with unbound selection data after activating the view and verifying
  ownership through `GetSelectedObjectsDrawingView2`.
- Created ordinates report `Type2 = 1` (horizontal request) and `Type2 = 7`
  (vertical request), not `7` and `8`. QA must accept `1`, `7` and `8`.

### Status

The Phase 0 gate in `docs/R23_IMPLEMENTATION_PLAN.md` is satisfied and
production Phase 1 is unblocked. R23-006, R23-009 and R23-010 are closed.
The remaining J-J top-border violations are production work for R23-704, now
measured truthfully in the page frame rather than unknown.

Production R23 source, `Fable.swp`, fixtures, the protected baseline, and
manual references remain unchanged.

## 2026-07-31 - Entity-correspondence route settled; ISelectData.View defect reproduced

### Live evidence

The Boolean normalization cleared every qualification gate and the run
exercised all three mapping routes for the first time.

- Route A, `IView.GetCorrespondingEntity(modelEdge)`, works: 12 of 24
  counterbore edges and 114 of 154 body vertices mapped with `error=0`.
- Route B, `IComponent2.GetCorrespondingEntity`, returned `Nothing` on every
  attempt and is not usable for part drawing views on this build.
- Route C confirms route A returns genuine drawing-context entities: each
  mapped entity identity-matches a visible-edge entry via `ISldWorks.IsSame`.
- Mapping is per-edge, not per-face — only one of each counterbore's two
  owned circular edges maps, so every owned edge must be attempted.
- The required scheme resolved end to end: six unique locations, two unique X
  and three unique Y, plus a stepped-bore X datum (r `0.0235`) and a
  bottom-left vertex Y datum.
- `ICurve.CircleParams` works, returning radii matching the owning cylinders
  exactly. The R23-006 exclusion was entirely an artifact of the guard defect.

Both ordinate groups then failed with runtime error 91 before any selection.
`CreateSelectData` returned a live object, so the failing statement is
`Set selectData.View = swView` — reproducing the behaviour already recorded
in `Module2_DrawingPipeline.CreatePrimarySection`. The MCP documents the
property as `get; set`, so this is an installed-build deviation, not misuse.

### Changed

- The probe now attempts the `ISelectData.View` binding inside a guarded
  helper and records `viewBinding=Bound` or `UnboundAfterError:91`, continuing
  with unbound selection data when it fails. Each selection's owning view is
  proved through `ISelectionMgr.GetSelectedObjectsDrawingView2`, and every
  group record carries the exact `lastStep` reached.

`AddOrdinateDimension` has still never been called. Production R23 source,
`Fable.swp`, fixtures, the protected baseline, and manual references remain
unchanged.

## 2026-07-31 - SOLIDWORKS Boolean contract established; CircleParams exclusion withdrawn

### Live evidence

The `rejectGate` instrumentation isolated the mechanism exactly. One call
produced both `isCircle=True` and `rejectGate=IsCircleFalse`: `CStr` rendered
the value `True` while `If Not value` fired on the same variable.

Established contract for installed SOLIDWORKS 2025 SP1.2 COM Booleans in this
VBA host: `If value Then`, `If value = False Then` and `CStr(value)` are
safe; `If Not value Then` is not, yielding `-2`, which VBA treats as True.
`CBool` is not a dependable normalization — `CBool(rawVariant)` worked for
`ISurface.IsCylinder` while `CBool(curve.IsCircle)` did not. Only an explicit
numeric comparison, `(CDbl(rawValue) <> 0#)`, is representation independent.

### Changed

- Added `NormalizeSwBoolean` to the disposable probe and applied it to
  `ICurve.IsCircle`, `ISurface.IsCylinder`, `IDrawingDoc.ActivateView`, both
  `IEntity.Select4` calls, and `IModelDoc2.GetSaveFlag`. Every remaining
  `If Not` in the module is an object `Is Nothing` test, a VBA built-in, or a
  normalized value. Raw returned values are logged alongside normalized ones.

### Withdrawn

- The `ICurve.CircleParams` exclusion in R23-006 is withdrawn as unproved.
  The `SkippedNotCircle` result came from
  `Module_R23Phase0FeatureProbe.ReadCircleState` line 611, which guards the
  call with `If Not isCircle Then` — the same defect. `CircleParams` was
  never invoked and no anomaly was ever observed. Its behaviour remains
  untested; the corrected probe now reads it as non-load-bearing evidence so
  the next run settles the question. Production still must not depend on it
  until a run exercises it.

## 2026-07-31 - Counterbore ownership proved; edge-closure gate corrected

### Live evidence

The Boolean-normalization fix worked. Twelve of the eighteen faces owned by
`CBORE for M6 Socket Head Cap Screw1` now read as cylinders with correct
radii (`0.0055` and `0.0033`), all passed `TransformPointToView` with
outline-contained page proof, and the six page centres form exactly the
required grid: two unique X (`0.080932`, `0.110932`) and three unique Y
(`0.077060`, `0.117060`, `0.157060`), spaced 30 mm and 40 mm to match the
drawing's `30.00` and `40.00`.

Feature ownership, cylinder qualification and the model-to-page transform are
proved for all six counterbore locations.

### Fixed

- Corrected the edge-closure gate. Every edge in the run returned
  `completeCircle=False` — all 24 feature-owned and all 64 visible drawing
  edges — including about thirty reporting `isCircle=True`, `uMin=0`,
  `uMax=6.283185307` and a zero endpoint gap, which satisfy every gate. With
  `GEOMETRY_TOLERANCE_M` at `0.0000001`, the success assignment should have
  been reached; the only construct able to swallow it was the single-line
  `If ... Then _` continuation immediately preceding it.
  `ReadEdgeCircleEvidence` now uses block `If` statements, reports the exact
  `rejectGate` per edge, and logs the `closureToleranceM` in force. No
  `Then _` construct remains in the module.

Mapping routes A, B and C have still never executed. Production R23 source,
`Fable.swp`, fixtures, the protected baseline, and manual references remain
unchanged.

## 2026-07-31 - Corrected-probe live results and Boolean-normalization fix

### Live evidence

The user compiled and ran probe build `20260731.2` on P-0251. The section and
J-J gates closed; the ordinate gate did not.

- Section targets are exact and non-stale: one `DIAMETER_47`, one
  `DIAMETER_40`, one `FIRST_LINEAR`, `exactTargetCounts=True`, both diameters
  `type2=6`.
- Direct part-source readback proved H7 is **absent** from `D1@Sketch4`
  (`toleranceType=0`, `fitType=-1`). The user selected controlled
  target-spec/reference authority, closing R23-806: R23 will apply
  `H7 +0.025/0.000` to an associative `Ø47` dimension with provenance
  recorded as target-spec authority and stated as such in QA.
- The J-J payload segment frame is proved, not inferred: all six endpoints
  matched the captured `CreateLine` view-sketch inputs at `deltaM=0`. The
  single inverse conversion is cross-validated by the independently supplied
  page-frame arrow endpoints, which agree exactly.
- Page-frame clearance measured three violations, all at the top:
  segment 1 start, upper arrow, and upper label exceed the content-border top
  of `0.287`. The lower arrow and label **clear** the measured
  part-identification note extent by about 7.7 mm, correcting the earlier
  screenshot-derived claim that they intruded into that band.

### Fixed

- Corrected a Boolean-normalization defect introduced in the previous probe
  revision: `If Not surface.IsCylinder Then` rejected all 18 owned counterbore
  faces because a raw SOLIDWORKS `VARIANT_BOOL` of `1` makes `Not value`
  evaluate to `-2`, which VBA treats as True. The probe now uses
  `CBool(ISurface.IsCylinder)`, matching the accepted feature probe, and logs
  the raw returned value. The same normalization was applied to
  `IDrawingDoc.ActivateView` on the ordinate transaction path, and the visible
  drawing-edge inventory now logs its curve-parameter fields.

The three mapping routes remain unexercised; the next run is the first that
can test them. Production R23 source, `Fable.swp`, fixtures, the protected
baseline, and manual references remain unchanged.

## 2026-07-31 - Corrected disposable Phase 0 drawing-contract probes

### Changed

- Corrected the disposable ordinate probe in
  `tools/r23-probes/import-transaction-source/Module_R23Phase0DrawingProbes.bas`
  (build `20260731.2-mapping-frame-h7-contracts`): every owned face and edge
  of the target feature now logs a qualification record, and each complete
  circular edge compares the direct active-part view mapping, the
  component-mediated mapping, and the `GetVisibleEntities2` inventory with
  `ISldWorks.IsSame` identity results. Ownership, cylindrical-radius,
  `IsCircle`, and closed `GetCurveParams3` qualification is unchanged, and
  `AddOrdinateDimension` remains gated on six counterbore locations
  resolving to two unique X and three unique Y coordinates.
- Corrected the disposable section probe: every per-dimension diagnostic
  value is reset at the top of each display-dimension iteration; diameter
  targets require the live-proven `swDiameterDimension = 6` type plus the
  exact nominal, with exact-count enforcement and `*_DUPLICATE` labelling;
  and the original part-source `D1@Sketch4`/`D1@Sketch6` tolerance and fit
  data are read directly through `IModelDoc2.Parameter`, feeding an
  explicit `R23_H7_AUTHORITY` record. The probe never invents H7.
- Corrected the disposable J-J diagnostic: the Module2 overlay captures the
  original page-frame construction points and the converted view-sketch
  `CreateLine` inputs; the parsed payload's segment frame is proved by
  point-by-point comparison; segment endpoints are converted to the current
  page frame exactly once via the exact inverse of the
  `IView.GetXform`/`IView.Angle` conversion; and page-frame clearance
  verdicts are logged against the content border, title block, measured
  part-identification note extent, and view outlines.
- Updated `tools/r23-probes/README.md` with the corrected-probe record
  types and rerun evidence expectations, and recorded the session's MCP
  contract lookups in `docs/SOLIDWORKS_API_VALIDATION.md`.

### Verification

Static verification only: structural source checks pass, both edited
sources remain ANSI/CRLF without BOM or `Attribute` metadata, the complete
offline suite reproduces the known 69-pass/5-stale-R20-failure baseline,
and the disposable drawing-contract manifest passes read-only preflight
against the copied probe SWP. The corrected probes have not been deployed,
compiled, or executed; production R23 source, `Fable.swp`, fixtures, the
protected baseline, and manual references are unchanged.

## 2026-07-31 - R23 Claude Code handoff after final Phase 0 probes

### Added

- Added `docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md` as the durable
  same-workspace handoff for Claude Code. It records the goal, source of truth,
  protected assets, accepted evidence, unresolved gates, R23 architecture, and
  exact next work package.
- Added `docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md` as the evidence index
  and disposable-probe correction specification.
- Preserved the two uncropped final-probe screenshots in the timestamped R23
  `live-probes` evidence folder without modifying their originals.

### Findings

- The ordinate probe stopped before `AddOrdinateDimension` because
  feature-owned counterbore edges did not survive the opaque mapping helper.
  The corrected diagnostic must log edge qualification and compare direct,
  component-mediated, and `GetVisibleEntities2` mapping paths.
- The section probe found the exact 47 mm and 40 mm imports as
  `swDiameterDimension = 6`, with no imported H7 or nonzero tolerance. It also
  retained stale per-item state across loop iterations, so only the exact
  identity records are accepted from that transcript.
- The J-J payload is structurally complete, but its segment records and
  arrow/label records use different coordinate frames. The screenshot confirms
  both top-border and lower part-identification-band collisions.

### Scope

This handoff checkpoint changes documentation and retained diagnostic evidence
only. Production R23 VBA, `Fable.swp`, the protected baseline, fixture models,
and manual reference drawings remain unchanged.

## 2026-07-30 - R23 Phase 0 implementation checkpoint

R23 implementation has started at the plan's mandatory runtime-proof gate.
This checkpoint does not change the production VBA revision or drawing
behavior.

### Preserved baseline

- Backed up `Fable.swp` before any material embedded-project change and proved
  its SHA-256 matches the production file.
- Verified 15/15 managed components and the exact R22 revision between the
  embedded project and exported source.
- Recorded the pre-change repository, SWP, and managed-source hashes under
  `test_assets/iteration_evidence/r23/20260730-075811/prechange/`.
- Re-ran the guarded deployment preflight successfully.

### Phase 0 preparation

- Added `tools/r23-probes/Module_R23Phase0FeatureProbe.bas`, a standalone,
  read-only probe for raw/effective feature types, configuration-specific
  suppression, typed Hole Wizard/extrude definitions, owned cylindrical
  geometry, pattern seed evidence, and both circular-curve read orders.
- The probe refuses any model outside the three authorized fixtures and pairs
  every successful `AccessSelections` with `ReleaseSelectionAccess`.
- Added usage and operator-gate guidance in `tools/r23-probes/README.md`.
- Updated the implementation plan with completed provenance/probe-preparation
  tasks and corrected its drawing-component context contract.

### API evidence

- Re-queried the local `solidworks-api` compatibility MCP and reflected
  installed SOLIDWORKS 2025 SP1.2 interop `33.1.2.4`.
- Confirmed `ICE -> GetTypeName`, `GetDefinition`, feature-owned faces,
  typed selection-access rollback/release, `ICurveParamData`, and the
  circle-parameter contracts used by the probe.
- Corrected the plan's earlier description of `GetVisibleComponents`: its
  returned component is limited and is suitable for `GetVisibleEntities2`;
  full component behavior requires
  `GetVisibleDrawingComponents -> IDrawingComponent.Component`.

Production R23 source remains intentionally unchanged until the required live
Phase 0 evidence exists.

### Phase 0 probe correction

- Preserved the fail-closed refusal from the first attempted run, which found
  an unauthorized `V:\VEEMAP\SW_data\...` document active.
- Preserved the first authorized-fixture transcript but rejected it as
  incomplete evidence: an `ObjPtr`-only recursion key visited 15 features and
  skipped the relevant ICE, Hole Wizard, extrusion, and mirror entries.
- Replaced the transient-wrapper-only recursion key with a composite feature
  diagnostic key.
- Updated suppression readback to accept and label both scalar and array
  `IsSuppressed2` return shapes. The installed VBA run returned a scalar
  Boolean even though older source assumed only an array.
- Added base `EXTRUSION` to the typed `IExtrudeFeatureData2` probe routes.
- Added extrusion contour count/state and owned profile-subfeature readback so
  the cut probe satisfies the plan's sketch/contour evidence requirement.
- Mirrored every feature-probe record to a timestamped evidence log to prevent
  Immediate Window buffer limits from truncating the contract proof.
- Prepared a copied `Fable.swp`, a three-module disposable overlay, and a
  guarded custom deployment manifest for the expanded import comparison. The
  two entry points create separate fresh drawings for selected-primary
  `AllViews=True` and deterministic section/side/primary `AllViews=False`,
  always with `DuplicateDims=True` and mask `18055274`.
- Added owner, source-identity, nominal, tolerance, fit, callout-variable,
  attachment, duplicate, view-delta, and raw display-geometry logging for the
  import variants. The custom manifest passed read-only deployment preflight;
  no VBA full-project compile or live import run is claimed.
- Kept the production R22 source and `Fable.swp` unchanged. The corrected
  disposable probe still requires user-operated compilation and runtime
  evidence before any production R23 classification work begins.

### Phase 0 feature evidence â€” 2026-07-31

- Compiled and ran the corrected read-only feature probe on the authorized
  P-0251 fixture. The retained `R23_FEATURE_20260731_040539.log` records 47
  visited features, all three `ICE -> Cut` resolutions, typed definition
  access/release, owned geometry, native CBORE and M5 Hole Wizard data, and
  `Mirror1` seed ownership.
- The configuration-specific `IsSuppressed2` result was Empty for every
  feature. The probe therefore records the separately labelled active-document
  `IsSuppressed` fallback; this is live evidence for the active `Defualt`
  configuration only, not a substitute for a drawing-view configuration proof.
- Both curve sequences retained stable `IsCircle=True` and complete
  `GetCurveParams3` evidence. Because the `CircleParams` helper still produced
  `SkippedNotCircle` despite true before/after predicates, R23 will not depend
  on `CircleParams`; it will qualify circles from the live-proven cylindrical
  face, `IsCircle`, and complete-boundary parameter evidence.
- Production R23 source and `Fable.swp` remain unchanged. The next live gate is
  the disposable expanded-import transaction comparison.

### Phase 0 import evidence â€” 2026-07-31

- Ran both expanded-mask import variants in fresh disposable P-0251 drawings.
  Each completed successfully with 25 unique imported source identities, no
  fixture save, and a clean isometric view.
- Selected `AllViews=False` as the R23 import transaction: explicit
  section/side/primary selection yielded 17/2/6 dimensions. The selected
  anchor with `AllViews=True` yielded 17/0/8, so it cannot satisfy the required
  side-view coverage.
- Proved native M5 Hole Wizard callout import and its tap-drill, thread, and
  depth variables. The only native callout landed in the section. Neither
  variant imported the required six-hole M6 counterbore callout, H7 fit, or a
  nonzero tolerance; R23 must therefore use the approved controlled fallback
  only when typed feature/model data proves its content.
- The generated probe sheets visually confirm that imported dimensions can
  overlap the section label/title area and do not create ordinate coverage.
  Production layout and explicit ordinate stages remain mandatory.

### Final Phase 0 probe preparation â€” 2026-07-31

- Added the disposable `Module_R23Phase0DrawingProbes.bas` overlay with public
  entry points for datum-first X/Y ordinate creation and targeted section
  dimension/J-J geometry evidence.
- The ordinate probe maps feature-owned complete circles without depending on
  `CircleParams`, requires six M6 locations resolved to two X and three Y
  coordinates, selects the datum first, appends each entity individually,
  decodes the installed ordinate result enum, and records `SetPickMode` plus
  zero-selection cleanup.
- The section probe records every imported section dimension, targeted 47/40
  tolerance/fit and display data, `IDrSection` state, and bounded parsing of
  J-J segments, arrows, and label positions.
- Added a guarded four-module deployment manifest targeting only the existing
  disposable `R23_Phase0ImportProbe.swp` and documented the exact user
  compile/run sequence. Production source, production `Fable.swp`, fixture
  models, the protected baseline, and manual references remain unchanged.
- Tied H7 proof specifically to the imported 47 mm dimension so an unrelated
  fit cannot satisfy the section contract, and added an exact reported-versus-
  flattened section-line count check before parsing any geometry.
- Reflected the installed SOLIDWORKS 2025 SP1.2 interop for the ordinate,
  selection, dimension/tolerance, display-data, `IDrSection`, and section-line
  members used by the probes.
- The four-source disposable manifest passed guarded read-only preflight.
  Source hygiene and structural checks passed for all overlay modules. The
  complete offline suite reproduced its known baseline: 74 tests run, 69
  passed, and only the five documented stale R20 contract failures remained.
- No full-project VBA compile or installed-build execution is claimed; those
  two gates remain user-operated.

## 2026-07-29 - R22 verified review resolution

Source identity moves to `target-spec-hybrid-v2-2026-07-29-r22`. R22 contains
the complete r20/r21 line, retains the corrections from `1698b7d` that satisfy
their contracts, and replaces the latest commit's unsafe or incomplete parts.

### Correctness

- `Module5`: removed the `ISurface.CylinderParams` trimming-edge fallback.
  `ICurve.IsCircle=False` now fails as `ClosedCircleCurveNotCircular`; an
  internal cylindrical face no longer proves that a closed trim is a circle.
  This prevents oblique elliptical trims from being assigned a false centre.
  The retained P-0251 contract probe shows every relevant owned edge already
  supplies `IsCircle=True` and seven-value `CircleParams`.
- `Module3`: replaced guessed feature aliases with the exact 2025
  `GetTypeName2` literals: `APattern`, `LocalChainPattern`, `DimPattern`,
  `DerivedHolePattern`, `SketchPattern`, and `LocalSketchPattern`, while
  retaining the already-correct pattern and mirror names.
- `Module4`: removed the module-level per-side ordinate-lane cache. If
  `AlignDimensions` returns False, ordinate types 1, 7, 8, and 16 now call and
  check `IDisplayDimension.AutoJogOrdinate`, then prove their readback
  positions. Deterministic `SetPosition2` lanes remain only for non-ordinate
  dimensions. This avoids both merging independent chains and splitting one
  type-1 chain by member proximity.

### Retained review fixes

R22 retains r21's projected-origin containment correction, pattern-seed
routing, FACE/SIDE callout separation, annotation-name self-identity fallback,
controlled static general-note proof, section-line parser robustness,
selection-readback tolerance, title-border window, redundant-check removal,
and reduced rejected-vertex logging.

### Verification

- Queried the local `solidworks-api` MCP for the load-bearing circle, feature,
  seed, ordinate, annotation, and enum contracts.
- Reflected the installed SOLIDWORKS 2025 SP1.2 interop assemblies, version
  `33.1.2.4`, for all newly used members and `swDimensionType_e`.
- Updated source/docs and prepared the guarded r22 SWP deployment.
- Full VBA Editor compilation, fixture runtime, QA, and visual/manufacturing
  acceptance remain separate gates.

See `docs/R22_REVIEW_RESOLUTION.md`.

## 2026-07-29 - R21 code-review remediation

Source identity moves to `target-spec-hybrid-v2-2026-07-29-r21`. R21 is r20
plus fixes for defects found in a full review of the r20 diff. No new features.

### Correctness

- `Module5`: the `ISurface.CylinderParams` fallback used the cylinder axis
  origin as the hole centre. The 2025 docs define that array as
  `origin/axis/radius` where origin is an arbitrary point on the axis, so an
  oblique bore was dimensioned at the wrong place and both circular edges of one
  cylinder collapsed to the same centre. The origin is now projected along the
  axis into the plane of the edge using a point known to lie on that edge.
- `Module5`/`Module8`: `TransformPointToView` always required the transformed
  point to lie inside `IView.GetOutline`. The model origin routinely projects
  off the solid, so `ProveCenterDatum` failed before examining any vertex and
  the Center datum became unprovable. Containment is now an explicit argument:
  required for hole centres and mapped vertices, not for the projected origin.
- `Module4`: the deterministic lane fallback allocated a new lane per
  dimension, giving an ordinate chain one baseline per member. Ordinates on a
  side now share one lane. `swOrdinateDimension` (1) - documented as "base
  ordinate and its subordinates" and already counted as an ordinate by
  `Module6` - was absent from the type switch and fell through to proximity
  routing, splitting a chain across two sides; types 1 and 16 are now handled.
- `Module3`/`Module5`: hole-seed resolution was wired only for `MirrorPattern`
  although `IFace2::GetSeedFeature` covers patterned, mirrored and copied
  bodies. Linear, circular, curve, table and fill patterns produced no
  candidate and no rejection record. All pattern families now share one proof.
- `Module7`: the FACE and SIDE manufacturing callouts shared one target
  position, so two callouts resolving to the same view were placed on top of
  each other and the collision check then failed the stage. SIDE has its own
  lane.
- `Module7`: the callout was excluded from its own collision scan using `Is`
  against a re-fetched `GetAnnotation` pointer, which may be a different
  wrapper for the same annotation - the note could collide with itself.
  Identity now also matches on annotation name.
- `Module7`: general-notes verification compared against a value read from a
  part custom property, but the controlled format carries that note as static
  text no property write can change. Any part carrying a `GeneralNotes`
  property made the stage permanently unprovable. The controlled reference
  constant is now also accepted, and whichever text proved the note is what the
  containment check uses.

### Robustness and diagnostics

- `Module6`: removed the exact-equality assertion between the section-line
  array length and the `Size` out-param of `GetSectionLineCount2`. The API docs
  describe `Size` only as "size, which includes an extra double per section
  line containing the layer ID" and never define it as the raw array length, so
  a convention difference would reject every drawing. The per-element cursor
  guards already bound every read; the pair is now recorded, not enforced.
- `Module6`: every malformed-array exit in `ValidateSectionLineInfo` now
  reports what failed and at which index. Previously they returned silently and
  the caller reused the same stage message as a real clearance violation.
- `Module4`: selection readback no longer has to equal the `Select3` success
  count. Re-selecting an already-selected annotation returns True without
  growing the list, so one duplicate let a cosmetic arrange step reject the
  whole run. An empty selection still fails; a shortfall is warned.
- `Module8`: widened the legacy title-block bottom-edge window, which accepted
  only 8.9-17.8 mm on A3 and excluded any format drawn to a 20 mm border.
  Replaying the recorded P-0251 template sketch gives byte-identical bounds
  before and after.
- `Module8`: replaced two rectangle-contract checks that compared the measured
  bounds against the same limits the candidate filter had already enforced, so
  they could never fire.
- `Module8`: removed `ViewToSheetCoordinates`, reduced in r20 to an identity
  copy with no callers, which would have silently mis-placed any future caller
  holding genuinely view-local coordinates.
- `Module5`: the per-vertex transform proof is logged only on rejection. It ran
  twice per edge over every face of every visible component.

### Verified statically only

API contracts were confirmed against SOLIDWORKS 2025 documentation through the
project `solidworks-api` MCP. The `GetSectionLineInfo2` cursor strides
(7/9/9/7) were checked against the documented layout and are correct as r20
wrote them; no change was made there. Procedure-block balance, duplicate `Dim`,
and signature/call-site arity were script-checked across every module.
**R21 has not been compiled in the VBA editor or run against a fixture.**
Full-project **Compile Project** and a focused P-0251 run remain mandatory.

## 2026-07-28 - R20 GetSectionLineCount2 compile hotfix

- A user-run full-project VBA compile stopped in
  `Module6_QAEngine.CheckSectionLineClearance` with `Argument not optional`.
- The project `solidworks-api` MCP and installed SOLIDWORKS 2025 interop both
  confirm `IView.GetSectionLineCount2(ByRef Size As Long) As Long`. R20 had
  omitted that mandatory output argument in Module6 even though Module2 already
  used the correct signature.
- Added `sectionLineInfoSize`, passed it to `GetSectionLineCount2`, and made the
  section-clearance gate reject nonzero counts with invalid sizes or
  `GetSectionLineInfo2` arrays whose item count does not match the API-provided
  size.
- Added a regression assertion that rejects any parameterless Module6 call.
  The complete project-local suite remains 74/74 passing.
- Corrected the deployment evidence interpretation:
  `COMPILE_PROBE|status=SUCCESS` only proves that the bootstrap procedure could
  execute. It does not invoke VBA Editor **Compile Project**, and therefore is
  not E5 full-project compilation.
- The first automatic deployment attempt correctly stopped because `Fable.swp`
  was open/locked. After the user closed it, deployment
  `20260728_142300` completed with a 15/15 managed-source match and exact r20
  revision. Its new `compile-probe-scope.txt` records that manual VBA Editor
  **Compile Project** remains mandatory.

## 2026-07-28 - R20 r19 functional-failure repair

### Diagnosed

- Reviewed the synchronized r19 P-0251 evidence at
  `test_assets/iteration_evidence/macro_qa/20260728_091302_P-0251-14A-001/`.
  R19 created four views and ten imported display dimensions, but reported zero
  circular edges, zero ownership candidates, zero canonical physical
  locations, and zero ordinate groups.
- Every audited internal-cylinder boundary stopped at
  `ClosedCircleIsCircleFalse`. SOLIDWORKS can expose a complete trimmed
  cylinder-boundary edge whose underlying `ICurve.IsCircle` predicate is
  false. Therefore that predicate cannot be the sole circular-boundary gate.
- `ModelToViewTransform` already supplies drawing-page coordinates for
  dimension placement. Adding `IView.Position` translated candidate and datum
  evidence twice and would have broken the newly activated center-datum path.
- The P-0251 side and section views were placed with a 6 mm requested gap while
  collision validation also required 6 mm. API readback drift therefore
  classified the threshold placement as a collision.
- Annotation QA used `UsableBottom`, the view-placement boundary above the
  title block, as a full-width lower-sheet exclusion. That rejected legal
  lower-left dimensions and section content even though the title block is only
  a lower-right rectangle.
- `AlignDimensions` was called without view-scoped `ISelectData`, per-annotation
  selection results, or selected-count readback. A one-dimension side view was
  also treated as an arrange failure although no multi-dimension arrangement
  was possible.
- R19 still lacked the controlled P-0251 stepped-bore, six-hole counterbore,
  and four-hole tapped manufacturing definitions required by the target
  specification.

### Fixed

- `TryReadClosedCircularEdge` now calls `IEdge.GetCurve` followed by
  `IEdge.GetCurveParams3`, requires a nondegenerate parameter span and
  coincident parameter endpoints, and retains the topology closure check.
  `ICurve.CircleParams` is used when available; otherwise the already
  ownership-proven internal face supplies center, axis, and radius through
  `ISurface.CylinderParams`.
- Normalized audited SOLIDWORKS Boolean returns with typed `Boolean` variables
  before applying negation or compound logic, including circle/cylinder
  predicates, `Select4`, `Select3`, detail label/outline setters, and
  `AlignDimensions`.
- Removed the second view-position translation. Candidate/datum sheet
  coordinates now equal the page coordinates returned by
  `ModelToViewTransform`; center-datum selection compares entities with the
  transformed model origin rather than numeric page origin `(0,0)`. Candidate
  centres, the projected origin, and mapped vertices now also require a runtime
  `IView.GetOutline` page-frame invariant and emit `TRANSFORM_PAGE_PROOF` or
  `TRANSFORM_PAGE_REJECT`.
- Increased the shared requested inter-view gap to 12 mm, added a 1 micrometre
  comparison tolerance, raised the P-0251 primary/side/section row to clear the
  lower J-J marker, and added requested-center, outline, and pair-clearance
  readback records with `Initial`/`Final` layout attribution.
- Added zoned content-border bounds from `ISheet.GetZoneMargin` to
  `CRunEvidence`. Annotation origins, leader points, and measurable note extents
  now use the actual border plus the measured title-block rectangle; the
  lower-left region is no longer rejected wholesale. The part-identification
  note extent is retained, and P-0251 section segments, arrows, and both label
  positions are parsed from `GetSectionLineInfo2` and checked against it.
- Dimension arrangement now uses `ISelectData.View`, checks every `Select3`
  return, verifies selected-count readback, and skips fewer-than-two selections
  without a false warning. `DIMENSION_ARRANGE` is required; when
  `AlignDimensions` returns `False`, positionable dimensions use deterministic
  6 mm lanes with exact readback and border proof, while unsupported
  radial/diametric positions must already be safe.
- Replaced the free-standing manufacturing summary with three associative
  P-0251 callouts. Each selects a retained ownership-proven drawing edge,
  creates its note with a leader, uses documented `<MOD-DIAM>` symbol syntax,
  proves nonzero attachment/leader readback, and checks note extents against
  borders, title block, part ID, model views, other measurable notes, and other
  annotation origins. Reuse now requires the complete normalized controlled
  definition rather than a short phrase that an imported Hole Wizard note
  could also contain.
- Advanced the source identity to
  `target-spec-hybrid-v2-2026-07-28-r20`.

### API and CodeStack evidence

- The project `solidworks-api` MCP verified the SW2025 contracts for
  `IEdge.GetCurveParams3`, `ICurveParamData`, `ICurve.IsCircle`,
  `ISurface.CylinderParams`, `IView.GetCorrespondingEntity`,
  `IAnnotation.GetPosition`, `IAnnotation.Select3`,
  `IAnnotation.GetAttachedEntities3`, `IAnnotation.GetLeaderCount`,
  `IModelDoc2.InsertNote`, `IView.GetSectionLineInfo2`,
  `IModelDocExtension.AlignDimensions`, `ISheet.GetZoneMargin`,
  `IModelDocExtension.AddOrdinateDimension`, and the section-view transaction.
- The CodeStack drawing examples established the page-coordinate placement
  pattern, real-component requirement for `GetVisibleEntities2`, direct
  selection of drawing-context entities, and view-scoped selection-data
  pattern. The macro retains stricter ownership, cleanup, and truthful QA gates
  than the examples.

### Verification

- The complete project-local suite passes **74 tests and 13,608 structural
  subtests**.
- Guarded deployment evidence is retained under
  `test_assets/iteration_evidence/swp_deployment/20260728_113559/`.
  `COMPILE_PROBE|status=SUCCESS`, candidate and promoted readbacks match all
  15 managed components, and both report embedded revision
  `target-spec-hybrid-v2-2026-07-28-r20`.
- `Fable.swp` is synchronized. The latest deployment folder retains the
  immediately preceding macro; the first r20 deployment evidence at
  `20260728_105950/` retains the pre-r20 binary.
- The macro has not yet been executed on P-0251 after this repair. Nonzero
  candidate/ordinate behavior, associative callout attachment/placement,
  deterministic arrangement, J-J clearance, and designer-level
  visual/manufacturing acceptance remain E6/E7 user-run gates.

## 2026-07-26 - R18 model-first ownership and truthful view QA

### Diagnosed

- Preserved and reviewed the synchronized r17 P-0251 run at
  `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/`.
  R17 retained the J-J section, imported 9 + 1 dimensions into the two
  orthographic views, created the isometric before import, executed eight
  reference-led layout moves, and calculated active-configuration mass as
  `1.296824 kg` / `1.30`.
- All 111 reverse correspondence attempts still returned `Nothing`, leaving
  zero owned locations and zero ordinate groups. R16 had already disproved
  treating the drawing proxies as model edges. Together the two runs disprove
  both drawing-edge traversal variants.
- The importer's immediate arrays reported per-view dimension counts
  `9, 1, 0, 0`, while final QA reported 39 total and 10 on the isometric.
  `GetFirstDisplayDimension5` followed by `GetNext5` advanced across the
  drawing sheet instead of remaining scoped to the starting view.
- The correct drawing property `Mass=1.30` did not update the visible
  `MASS(KG)` value of `1296.82`, proving the note remains linked to a different
  model/system property.
- The visual layout is substantially improved. Its repeated note failures came
  from invalid note extents while the controlled boundaries were already
  diagnostic and unproved.
- `IDrawingDoc.AddHoleCallout2` was rejected as an unattended fallback because
  the documented method requires a user click in its generated callout dialog.

### Fixed

- Passed the existing audited model hole-feature collection into the fallback
  engine. It now enumerates owned internal cylindrical faces and circular model
  edges, then maps each known model edge into one drawing view with
  `IView.GetCorrespondingEntity`.
- Preserved full-circle, radius-matched cylinder, feature-face membership,
  active referenced configuration, normal-axis, physical-instance,
  imported-coverage, datum-first selection, result-code, and cleanup gates.
- Replaced datum reverse lookup with visible-solid-body edge/vertex traversal
  and model-vertex-to-view mapping through `IView.GetCorrespondingEntity`.
- Replaced final QA's sheet iterator with the exact
  `IView.GetDisplayDimensions` array for each view.
- Added unique mass-note link repair. Exactly one property-linked note
  containing `MASS` is required before changing it to `$PRP:"Mass"`; rebuild
  must preserve that unresolved link and render the computed two-decimal
  kilogram value.
- In diagnostic mode only, and only while controlled boundaries are already
  unproved, invalid note extents are skipped with an explicit
  `acceptance=False` warning. Production behavior remains fail closed.
- Advanced the source identity to
  `target-spec-hybrid-v2-2026-07-18-r18`.

### Verification

- The local 2025 API corpus and installed interop `33.1.2.4` confirmed
  `IView.GetCorrespondingEntity`, `IFeature.GetFaces`, `IFace2.GetEdges`,
  `IPartDoc.GetBodies2`, `IBody2.GetEdges`, vertex accessors,
  `IView.GetDisplayDimensions`, and the get/set
  `INote.PropertyLinkedText` contract.
- R18 source-contract coverage includes the model-first candidate and datum
  graphs, exact per-view QA, unique mass-link readback, and diagnostic-only note
  skip.
- The complete project-local suite passes **66 tests**; all 15 managed
  components pass the Windows-1252/CRLF/no-BOM/no-metadata format gate; and the
  guarded read-only deployment preflight resolves r18 with the bootstrap
  present.
- Embedded VBA compilation, synchronized authorized execution, ordinate
  creation, and visual/manufacturing acceptance remain required.
- The missing structural `ITitleBlock`, approved exact title links, grouped
  callout completeness, and stepped-bore manufacturing definition are not
  claimed fixed.

## 2026-07-26 - R17 third drawing-output repair

### Diagnosed

- Retained and reviewed the guarded r16 P-0251 run at
  `test_assets/iteration_evidence/macro_qa/20260726_145625_P-0251-14A-001/`.
  Its deployment evidence proves embedded compilation, revision readback
  `target-spec-hybrid-v2-2026-07-18-r16`, and 15/15 managed-component source
  synchronization.
- R16 genuinely created the J-J section and imported 10 model annotations into
  the two registered orthographic views. The description also resolved from
  `model-configuration:Part Name`.
- All 111 component-qualified visible drawing edges were rejected as
  `NotCircular`. The runtime result disproves r16's assumption that an entity
  returned by `IView.GetVisibleEntities2` is already the underlying model edge.
- The isometric view was created after model import but nevertheless displayed
  the 10 imported dimensions. Creation order, not only `AllViews=False`, must
  isolate the orientation aid.
- The generic layout forced all non-primary views into the shallow band above
  the title block. The tall side and section views could not fit, so no layout
  moves occurred.
- The visible `MASS(KG)` value was `1296.82`, while the manual reference is
  `1.30`. Copying an unverified `Mass` or `Weight` string cannot establish its
  unit.
- `selectedByID=True` was followed by a false warning because VBA's bitwise
  `Not` was applied to the COM Boolean readback.
- The structural `ITitleBlock` and its exact controlled links remain absent.
  This remains a controlled-template blocker rather than a drawing-generation
  defect.

### Fixed

- Mapped every visible drawing edge through
  `IModelDocExtension.GetCorrespondingEntity2` on the drawing document before
  applying the existing full-circle, matched-cylinder, feature-ownership,
  configuration, axis, and deduplication gates. No semantic gate was relaxed.
- Created the isometric view before selected-view model import. Import still
  executes only for registered eligible orthographic views with
  `AllViews=False` and `DuplicateDims=True`.
- Added a P-0251-only reference layout profile that packs the primary, narrow
  side, and J-J section views across the title-clear left zone and places the
  1:2 isometric in the upper-right zone. It moves outline centres without
  changing view scale, then runs the existing boundary/title/note/collision
  validation.
- Replaced the raw mass-property copy with
  `IModelDocExtension.GetMassProperties2(Higher, status, False)`, requires
  status `OK`, reads the documented metric mass at array index 5, and writes a
  two-decimal kg value for the active configuration.
- Replaced the ambiguous `If Not selectedView` test with a direct Boolean
  branch.
- Advanced the source identity to
  `target-spec-hybrid-v2-2026-07-18-r17`.

### Verification

- Local SOLIDWORKS 2025 API MCP data confirmed the drawing-to-model direction of
  `IModelDocExtension.GetCorrespondingEntity2`, selected-view semantics of
  `InsertModelAnnotations4(AllViews=False)`, and metric/system-unit mass
  behavior. Installed interop reflection confirmed
  `swMassPropertyAccuracyLevel_Higher=2`,
  `swMassPropertiesStatus_OK=0`, and the `GetMassProperties2` signature.
- The complete project-local offline suite passes **62 tests**.
- R17 embedded VBA compilation, the next authorized P-0251 run, and visual/
  manufacturing acceptance remain required.

## 2026-07-26 - R16 second drawing-output repair

### Diagnosed

- Retained and reviewed the synchronized r15 P-0251 run at
  `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/`.
  R15 proved the requested portrait primary orientation, 1:1 scale,
  orthographic-only dimension policy, sheet restoration, and final selection
  cleanup.
- The J-J section stopped at
  `SECTION_STEP=SelectData.ViewAssignment.Before` with VBA error 91 even though
  the source drawing view and three section-line segments already existed.
- Model import logged a matching active view and `selectedByID=True` but then
  emitted the mutually inconsistent legacy selection-gate failure. That
  redundant gate prevented `InsertModelAnnotations4` from executing.
- All 111 visible component edges were rejected as `NotCircular`. The r15 path
  applied a second corresponding-entity conversion to entities already returned
  for the requested visible component.
- Layout treated the title-block top as a full-width bottom boundary. The
  corrected portrait primary therefore could not use valid lower-left sheet
  area even though the title block occupies only the lower-right rectangle.
- The model exposes the controlled description under `Part Name`, while the
  reader checked only `Description` and `PartName`.
- The sheet still has visible title graphics but no structural `ITitleBlock`.
  This is a controlled-template input failure, not a macro path or drawing-view
  defect.

### Fixed

- Kept the section source view activated and the three sketch segments selected,
  but removed the failing `ISelectData.View` assignment. Selection count,
  order, owning drawing view, and mark are still verified before
  `CreateSectionViewAt5`.
- Made exact `ActiveDrawingView` readback the required selected-view import
  proof. `SelectByID2=False` is now diagnostic only; the importer records
  `MODEL_IMPORT_EXECUTE` immediately before
  `InsertModelAnnotations4(AllViews=False, DuplicateDims=True)`.
- Passed `IView.GetVisibleEntities2(component, Edge)` results directly into the
  strict circularity, cylindrical-face, feature-ownership, configuration, axis,
  and deduplication gates.
- Replaced the full-width lower reserve with zone-aware placement. The primary
  can use the safe lower-left zone; remaining views use the band above the title
  block. Border, title-block, note, and view collisions still fail closed, and
  no view scale is silently changed.
- Added the exact `Part Name` alias after `Description` and `PartName`, retaining
  configuration-first then document-level lookup and never inventing a value.
- Advanced the source identity to
  `target-spec-hybrid-v2-2026-07-18-r16`.

### Verification

- Local SOLIDWORKS 2025 API MCP contracts were used for
  `ISelectData.View`, `CreateSectionViewAt5`,
  `InsertModelAnnotations4`, `IView.GetVisibleEntities2`, corresponding-entity
  direction, and structural title-block behavior. No internet source was used.
- The complete project-local offline suite passes **61 tests**.
- Embedded VBA compilation, the next authorized P-0251 run, and visual/
  manufacturing acceptance remain required.

## 2026-07-26 - R15 first drawing-output repair

### Diagnosed

- Compared the first complete r14 P-0251 diagnostic drawing and its QA report
  with the protected manual reference. The generated primary view was rotated
  90 degrees from the reference, the J-J section was absent, model annotations
  were present on the isometric view, all 111 circular drawing entities failed
  model correspondence, no ordinate candidates were accepted, and the template
  exposed visible title graphics without an actual `ITitleBlock` definition.
- Confirmed through the local `solidworks-api` MCP corpus that
  `IView.Angle` is a read/write rotation angle in radians,
  `IDrawingDoc.InsertModelAnnotations4(AllViews=False)` operates on the selected
  drawing view, `IModelDocExtension.GetCorrespondingEntity2` is called on the
  underlying referenced document, `CreateSectionViewAt5` requires selected
  section-line entities, and `ISheet.TitleBlock` returns `Nothing` when the
  sheet has no structural title block.

### Fixed

- Added a P-0251-only clockwise 90-degree primary-view rotation with immediate
  `IView.Angle` readback. Other fixtures and named views are unchanged.
- Replaced drawing-wide annotation insertion with explicit, checked selection
  and import for each registered orthographic view. Section, detail, sheet, and
  isometric views are excluded, and `DuplicateDims=True` remains enforced.
- Changed fallback ownership correspondence to use
  `swView.ReferencedDocument.Extension.GetCorrespondingEntity2(drawingEdge)`.
  Rejection evidence now distinguishes an unavailable referenced document,
  unavailable extension, and no correspondence from that document.
- Added exact `SECTION_STEP` before/after evidence around selection-manager
  acquisition, selection-data creation, view assignment, every segment
  selection, selection verification, and `CreateSectionViewAt5`. The next live
  error now identifies the failing section operation instead of reporting only
  VBA error 91.
- Advanced the source identity to
  `target-spec-hybrid-v2-2026-07-18-r15`.

### Verification

- The complete project-local offline suite passes **60 tests**.
- The changed deployable VBA files contain no `Attribute` metadata and no UTF-8
  BOM. Full embedded VBA compilation, the next P-0251 run, and visual acceptance
  remain required.

## 2026-07-26 - Host-project candidate save

### Fixed

- Replaced `VBProject.SaveAs`, which raises VBA error 748 for a SOLIDWORKS
  host-managed macro project, with the VBE built-in Save command. PowerShell
  copies the saved candidate input to the output slot before compile and source
  verification.
- Added failure-stage and component identity to Module0 deployment results.

### Live validation

- Refreshed and compiled the bootstrap embedded in `Fable.swp`, then completed
  the guarded deployment through the running SOLIDWORKS 2025 SP1.2 instance.
- The candidate compile probe succeeded, all 15 managed components matched the
  exported source, and both candidate and post-promotion verification reported
  revision `target-spec-hybrid-v2-2026-07-18-r14`.
- SHA-256 readback proved the promoted `Fable.swp` is byte-for-byte identical
  to the verified candidate. The distinct pre-deployment macro is retained in
  `test_assets/iteration_evidence/swp_deployment/20260726_100426/`.

## 2026-07-26 - RunMacro2 IDispatch fallback

### Fixed

- Added a late-bound `IDispatch` fallback to the compiled macro invoker for
  running SOLIDWORKS instances whose ROT proxy rejects `QueryInterface` for
  `ISldWorks` with `E_NOINTERFACE`. The fallback explicitly marks the fifth
  `RunMacro2` argument as by-reference so the real `swRunMacroError_e` value is
  retained.

## 2026-07-22 - Explicit VBA component-type creation

### Fixed

- Replaced metadata-dependent `VBComponents.Import` with explicit
  `VBComponents.Add` calls using standard-module type 1 and class-module type
  2, followed by `CodeModule.AddFromString`. Metadata-free handler sources now
  remain class modules, so their `WithEvents` declarations compile correctly.
- Clarified that `swRunMacroError_e` value 11 after stopping the debugger is a
  user-interrupt result, not the originating compile defect.

## 2026-07-22 - Metadata-free deployable VBA source

### Fixed

- Removed VBA `Attribute` records and UTF-8 BOMs from deployable standard and
  ordinary class modules. Deployment now assigns component names explicitly,
  rejects metadata-bearing inputs, and ignores VBE-generated metadata during
  post-import source verification.

## 2026-07-22 - Compiled RunMacro2 bridge

### Fixed

- Moved the `ISldWorks.RunMacro2` invocation into a small runtime-compiled C#
  bridge. This bypasses PowerShell's COM method binder and emits the final
  `out int Error` argument exactly as the installed SOLIDWORKS 2025 interop
  assembly requires.

## 2026-07-22 - SOLIDWORKS COM identity conversion

### Fixed

- Replaced PowerShell's unsupported `-as ISldWorks` conversion with
  `Marshal.GetTypedObjectForIUnknown`. The active generic COM wrapper can now
  be bound to the installed SOLIDWORKS 2025 `ISldWorks` interface before the
  strongly typed `RunMacro2` call.

## 2026-07-22 - Strongly typed RunMacro2 deployment invocation

### Fixed

- Replaced PowerShell's late-bound `RunMacro2` call with the installed
  SOLIDWORKS 2025 `ISldWorks` interop binding. This correctly marshals the
  method's final `out Int32` error argument and avoids the pre-execution
  `TYPE_E_ELEMENTNOTFOUND` COM failure.

## 2026-07-22 - SWP deployer default-path fix

### Fixed

- Corrected `Deploy-TargetSpecHybrid.ps1` so its default manifest path is
  resolved after `$PSScriptRoot` is available. Running the documented command
  from the workspace no longer resolves the manifest to `\deployment-manifest.json`.

## 2026-07-22 - Guarded SWP source deployment automation

### Added

- Added `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1` to promote the exported
  target-spec source into `Fable.swp` without module-by-module copy/paste.
- Added a fixed 15-component manifest covering the nine production standard
  modules and six ordinary class modules. `ThisLibrary` and both UserForms stay
  outside automated replacement because they are host/designer components.
- Added a stable `Module0_SourceDeployment` bootstrap that imports native
  `.bas` and `.cls` files through `VBComponents.Import`, so `Attribute VB_Name`
  metadata is consumed by the VBA importer instead of pasted as code.
- Added candidate-only mutation, timestamped original backups, a `RunMacro2`
  compile probe, `olevba` source/revision readback, atomic promotion, and
  post-promotion verification with automatic restoration on failure.
- Added focused deployment-tooling tests and increased the offline suite from
  52 to 56 passing tests.

### Current gate

- `Fable.swp` does not yet contain the stable deployment bootstrap. It requires
  one final one-time import plus live compile/save validation before automated
  deployment can be exercised end to end.

## 2026-07-22 - R14 preflight rollback and verified view-activation retry

### Fixed

- Removed the r13 `EditSheet2`/`GetEditSheet` hard preflight. In the user's
  installed build it stopped the pipeline before sheet measurement or view
  creation, despite r12 having already created three drawing views from the
  same template.
- Kept `IDrawingDoc.ActivateView` as the primary view-context operation and
  retained `ActiveDrawingView` as the acceptance readback.
- Added one bounded recovery path: when the primary activation does not produce
  the requested active view, select the named object as `DRAWINGVIEW`, retry
  `ActivateView`, and accept only an exact `ActiveDrawingView` name match.
- Updated all section, detail, model-annotation, arrange, candidate-collection,
  and ordinate-creation callers to pass the drawing `ModelDoc2` needed for the
  selection-assisted retry.
- Restored an explicit named `DRAWINGVIEW` selection for the per-view
  `InsertModelAnnotations4` retry. With `AllViews=False`, the 2025 contract
  targets the selected view; the selection is now checked and retained until
  the import call completes.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r14`.

### Evidence

- The r13 report's first and only primary failure was `Drawing remained in
  sheet-format edit mode after IDrawingDoc.EditSheet2`; all 14 stage-gate
  failures were downstream consequences and the SOLIDWORKS mutation sequence
  remained at drawing creation.
- The r12 output had already proved that the template could create and display
  Front, Left, and isometric views. The r13 preflight was therefore a regression,
  not evidence of a broken template.

## 2026-07-22 - R13 drawing-context and packed-layout correction

### Fixed

- Added an explicit `IDrawingDoc.EditSheet2` normalization and
  `GetEditSheet=True` readback immediately after drawing creation. A template
  saved in sheet-format edit mode can no longer leave later view operations in
  the wrong drawing context.
- Changed `ActivateDrawingView` to verify `IDrawingDoc.ActiveDrawingView`.
  `ActivateView=False` is accepted only when the requested named view is the
  active-view readback; a missing or different active view remains fatal.
- Replaced the diagnostic equal-cell grid with measured two-row packing that
  preserves orthographic and section scales and rejects real boundary or
  collision violations.
- Assigned the P-0251 undimensioned isometric orientation aid an independent
  standard 1:2 scale so it can coexist with the required 1:1 manufacturing
  views on A3.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r13`.

### Evidence

- The user's r12 run created Front, Left, and isometric views, but section
  creation, whole-drawing model import, and ordinate collection all stopped at
  the same `ActivateView returned False` helper.
- The same run proved that the three created 1:1 views could not fit the
  diagnostic equal-cell grid; the full-size isometric visibly overlapped both
  orthographic views and the title area.

## 2026-07-22 - R12 inspection-pipeline decoupling

### Fixed

- Separated diagnostic drawing generation from production QA acceptance so a
  recoverable verification failure no longer prevents later independent stages
  from producing inspectable output.
- Added readback verification for sheet scale, document rebuild state, and final
  sheet context instead of treating every `False` command result as proof that
  the requested state was not achieved.
- Added a conservative diagnostic-only drawing reserve when visible title-block
  graphics exist without a defined `ITitleBlock`. This fallback is explicitly
  recorded as non-acceptance and cannot prove the `LAYOUT` stage.
- Diagnostic runs now continue past failed section, detail, layout, isometric,
  and rebuild stages where later work is independent and safe.
- When diagnostic mode creates at least one view, completion now presents a
  concise `Diagnostic Output Ready` warning instead of the full fatal-looking
  QA report dialog; the complete failure evidence remains in the Immediate
  Window and atomic QA report.
- Annotation-boundary QA now calls `IAnnotation.GetPosition` only for the
  annotation types that the 2025 API documents as supported. Center marks,
  centerlines, cosmetic threads, and other unsupported annotation types no
  longer create false `position is unavailable` failures.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r12`.

### Evidence

- The user's r11 run created both required orthographic views, but stopped at
  `EditRebuild3=False` before section, dimensions, isometric view, title content,
  and final layout were attempted.
- The visible sheet showed 1:1 scale and an active sheet even though earlier
  setter-only checks reported scale and final sheet-context failures.

## 2026-07-22 - R11 display-mode readback gate

### Fixed

- Replaced the fatal setter-only `IView.SetDisplayMode4` check with a verified
  setter-and-readback transaction for orthographic and isometric views.
- A `False` setter result is accepted only if `IView.GetDisplayMode2` equals the
  requested `swDisplayMode_e` value; the run records a warning in that case.
- A readback mismatch remains a hard view-configuration failure.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r11`.

### Evidence

- The user's r10 run created `Drawing View1` and visibly displayed hidden lines,
  but stopped because `SetDisplayMode4` returned `False` before the view was
  counted or the configured Left view was attempted.

## 2026-07-22 - R10 diagnostic scale bypass

### Changed

- Extended `DIAGNOSTIC_DRAWING_MODE` past `ISheet.SetScale` failure. Diagnostic
  runs now retain the failed `SHEET_SCALE` evidence and continue using the
  template's existing scale so view generation can be inspected.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r10`.

### Evidence

- The user's r9 run continued past missing `ITitleBlock`, then stopped before
  view creation because `ISheet.SetScale` returned `False` for the already-1:1
  diagnostic sheet.


## 2026-07-22 - R9 diagnostic drawing mode

### Added

- Added `DIAGNOSTIC_DRAWING_MODE`, defaulting to `True` temporarily, so a
  controlled-sheet preflight failure such as missing `ITitleBlock` no longer
  prevents the macro from attempting view creation for inspection.
- Diagnostic runs retain all failed stage evidence and are explicitly marked as
  non-acceptance results. Set the constant to `False` before production use.

### Evidence

- The user's r8 run proved the template and sheet-format paths and dimensions,
  but stopped before any view because `ISheet.TitleBlock` was `Nothing`.


## 2026-07-22 - R8 structural controlled-sheet acceptance

### Fixed

- Changed `SheetFormatVisible = False` from a hard controlled-sheet failure to
  a warning after the r7 setter/readback remained false in SOLIDWORKS 33.1.2.
- Controlled-sheet acceptance now proceeds to the stronger template-path,
  format-name, title-block extents, zone-margin, and usable-area checks; the
  visibility flag remains recorded for diagnosis.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r8`.

### Evidence

- The user's fresh-drawing screenshots visibly show the border and title block.
- The r7 run proved both V-drive template paths and valid A3 dimensions but
  still returned `SheetFormatVisible = False` after an explicit setter.


## 2026-07-22 - R7 sheet-format visibility normalization

### Fixed

- Added an explicit `ISheet.SheetFormatVisible = True` normalization and
  readback during controlled-sheet preflight. The macro now records the
  visibility mutation and fails only if SOLIDWORKS refuses the setter or the
  subsequent readback remains false.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r7`.

### Evidence

- The user's r6 run proved both controlled paths were now correct:
  `VEEMAP DRAWING.DRWDOT` and `VEEMAP DRAWING.SLDDRT`.
- SOLIDWORKS still returned `SheetFormatVisible = False` while the same
  template rendered its border/title block in the user's fresh-drawing view.


## 2026-07-18 - R6 sheet-size output binding correction

### Fixed

- Changed `MeasureControlledSheetRegions` to receive `ISheet.GetSize` output in
  local `Double` variables and then copy the values into `CRunEvidence`.
- Removed the direct use of object members as COM `ByRef` output arguments,
  which left `SheetWidth` and `SheetHeight` at zero after a drawing had been
  created successfully.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r6`.

### Evidence

- The user's r5 run accepted the corrected `.DRWDOT`, created a drawing
  document, and reached controlled-sheet measurement. It then failed with
  `invalid physical dimensions`, zero recorded width/height, and no downstream
  view creation.
- The SOLIDWORKS 2025 `ISheet.GetSize` contract defines width and height as
  `ByRef Double` outputs; all other production call sites already use local
  variables.

## 2026-07-18 - R5 controlled-template path correction

### Fixed

- Corrected `CONTROLLED_TEMPLATE_PATH` from
  `V:\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT` to the path shown by
  the user's File Explorer breadcrumb and proved by a read-only existence
  check:
  `V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`.
- Advanced the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r5` so new QA evidence cannot be confused
  with the frozen r4 package.

### Evidence

- The user's r4 project compiled successfully in the SOLIDWORKS VBA editor.
- The first run failed before any SOLIDWORKS mutation because the configured
  path omitted the `VEEMAP` directory level; all later stage-gate failures were
  downstream `NOT_STARTED` consequences.

## 2026-07-18 - R4 target-spec hybrid source completion

### Added

- Completed the coherent replacement source under
  `src/target-spec-hybrid-v2/` with identity
  `target-spec-hybrid-v2-2026-07-18-r4`.
- Added fixture-locked acceptance profiles for all three authorized parts,
  including exact view roles, deterministic J-J/B-B section policy, and
  direction-specific P-0251 datum behavior.
- Added mandatory Pump Holder Details C and D from the exact `*Bottom` source
  view at independent 3:1 scale. The detail transaction now verifies active
  source-view identity, the single complete circular selected profile, created-view and
  parent identity, exact scale ratio, initial placement, profile ownership,
  standard/circle style, and full smooth outline before structural success.
- Added sticky requirement-stage evidence, separate physical/projection
  location ledgers, exact title source/link/rendered/extent records, actual
  sheet-scale evidence, final-cleanup invalidation, and atomic report writing.
- Added r4 source-contract tests covering fixture profiles, fixed-hybrid order,
  nonrecursive finalization, view/section/detail postconditions, truthful UI
  behavior, import guidance, and duplicate VBA declarations.

### Changed

- Removed unconditional 90-degree orthographic rotation and silent independent
  scale reduction.
- Reworked visible-circle qualification into a fail-closed, matched-face,
  configuration-aware ownership path with stable physical/family/scope keys.
- Reworked imported-dimension reconciliation to require attachment-backed,
  datum/family/view/direction coverage before suppressing fallback ordinates.
- Reworked ordinate creation around separate X/Y datums, checked selection
  order/counts, decoded return values, successful-group-only suppression, and
  unconditional pick-mode/selection cleanup.
- Reworked title, layout, QA, and finalization so mandatory unknown or
  uninspectable state fails instead of being reported as success.
- Corrected form lifecycle, list validation, custom-scale parsing, fixture
  locking, and host/UserForm import instructions.
- Implemented D-03 as the r4 first-acceptance policy, pending user override or
  confirmation: Pump Holder Details C and D are mandatory because the
  reference's `7 x 4` and `C0.5` regions are ambiguous at the main-view scale.
  E4/E6 coordinate/profile transaction behavior and E7 visual/content proof
  remain required.

### Verified offline

- Reflected the used SOLIDWORKS 2025 interfaces and enum values from installed
  interop file version `33.1.2.4`, including the new detail-view contracts.
- Ran the complete project-local companion suite: **49 tests passed**.
- Passed structural VBA procedure, continuation, caller/signature, inventory,
  and duplicate-declaration checks.
- Frozen the r4 source archive, API/test checkpoint, and 35-entry SHA-256
  manifest under
  `test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/`.
  The older checkpoint remains historical r3/28-test evidence.

### Remaining gates and scope

- R4 is E2/E3 evidence only; it has not been embedded, compiled, run, or
  visually accepted.
- D-04 remains open because the controlled `.drwdot`/`.slddrt` and approved
  property/cell mapping are absent.
- No live SOLIDWORKS control was used and no `.swp`, authorized part, protected
  baseline, or manual reference drawing was changed.

## 2026-07-18 - Planning-only target-spec completion audit

### Changed

- Rewrote `docs/TARGET_SPEC_HYBRID_V2_IMPLEMENTATION_PLAN.md` as a complete,
  resume-ready plan with explicit source, compile, live-probe, fixture, visual,
  and release gates.
- Corrected the resumption path: the current r3 export is retained as an
  architectural baseline, but a coherent offline source-completion revision is
  required before the recommended user import/compile handoff.
- Updated `docs/CURRENT_STATUS.md` to match that corrected critical path.

### Added to the plan

- Native form and `ThisLibrary` import strategy;
- installed-interface compile risks for section and datum selection;
- fixture-aware view planning and removal of unconditional rotation;
- unique physical feature identity, matched-face ownership, approved datums,
  family-scoped coverage reconciliation, and recorded ordinate transactions;
- feature-led sections, controlled title/property/mass/scale proof, complete
  annotation/leader layout, and requirement-level QA;
- exact user decisions, regression matrix, evidence packet, and definition of
  done.

### Verified

- Installed reflection reconfirmed
  `swCreateOrdDimErr_GenFailure = 1` and
  `swCreateOrdDimErr_OrdFailure = 7`.
- All 18 checkpointed files under `src/target-spec-hybrid-v2/` still match the
  retained r3 SHA-256 manifest.

### Scope

- Planning and documentation only.
- No VBA source, `.swp`, authorized model, protected baseline, reference
  drawing, or live SOLIDWORKS state was changed.

## 2026-07-18 - R3 offline stabilization and compile handoff

### Fixed

- Narrowed the source-contract test that prohibited `InsertNote` so it detects
  an actual method call instead of falsely matching the legitimate
  `InsertNotes` configuration field.
- Changed candidate configuration proof to use
  `IView.ReferencedConfiguration` and
  `IFeature.IsSuppressed2(swSpecifyConfiguration, configurationNames)`.
  Candidates now fail closed when the drawing view's referenced configuration
  or suppression result cannot be proved.
- Updated the macro/export identity to
  `target-spec-hybrid-v2-2026-07-18-r3`.

### Added

- Added static tests for balanced VBA procedure blocks, line length,
  continuation limits, installed ordinate-result decoding, and referenced-
  configuration usage.
- Added the complete r3 SHA-256 manifest and installed-interop checkpoint under
  `test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_offline/`.

### Verified offline

- Reflected the newly used SOLIDWORKS 2025 sheet, title-block, annotation,
  custom-property, section, view-configuration, and active-sheet members from
  installed interop file version `33.1.2.4`.
- Cross-checked `FaceInSurfaceSense` and the newly used contracts against the
  project-local SOLIDWORKS API MCP corpus.
- Ran the complete project-local companion suite: 28 tests passed.

### Documentation

- Reconciled the target-spec gap ledger with the clean replacement instead of
  the older broken active macro.
- Corrected the first focused P-0251 handoff to require one horizontal J-J
  section; a sectionless run is now explicitly diagnostic and expected to fail
  fixture QA.
- Made the clean-project import and full-project compile the next gate.

### Scope

- No live SOLIDWORKS control was used.
- No `.swp`, authorized part, manual reference drawing, or protected baseline
  file was changed.

## 2026-07-17 - Target-spec hybrid V2 clean replacement

### Added

- Added the complete importable replacement source set under
  `src/target-spec-hybrid-v2/` without changing the protected baseline or the
  existing broken macro.
- Added `CHoleCandidate`, `CDatumProof`, and `CRunEvidence` class modules for
  feature ownership, selectable datum proof, and structured fail-closed QA.
- Added explicit runtime support for drawing-document activation, real sheet
  scale, model/view transforms, measured layout, and normal sheet-context
  restoration.
- Added companion source-contract tests for the complete module set,
  authorization, annotation import, feature ownership, datum proof, ordinate
  cleanup, and QA behavior.
- Added `src/target-spec-hybrid-v2/README_IMPORT.md` with import order and the
  focused P-0251 compile/run handoff.

### Changed

- Retired the incremental repair backlog in the target specification in favor
  of the clean replacement strategy.
- Replaced the misleading form preview text with a semantic
  hole-producing-feature count.
- Updated current status to make VBA compilation and runtime/visual acceptance
  the next unresolved gates.

### Offline verification

- Reflected the installed SOLIDWORKS 2025 type library for the methods,
  properties, and enums used by the replacement.
- Ran the complete project-local hybrid companion suite: 25 tests passed.
- Confirmed every supplied code component uses `Option Explicit` and that the
  replacement contains no blank-view activation or
  `GetVisibleEntities2(Nothing, ...)` path.

### Scope

- No live SOLIDWORKS control was used.
- No `.swp`, authorized part, manual reference drawing, or protected baseline
  file was changed.

## 2026-07-17 - Global collaboration skill and feature-ownership design

### Added

- Installed the global
  `C:\\Users\\V.T\\.agents\\skills\\solidworks-collaborative-verification`
  skill. It makes VBA editing/compilation/execution user-owned and requires an
  explicit user-versus-Codex operator choice before every distinct live
  SOLIDWORKS task.
- Added
  `test_assets/iteration_evidence/2026-07-17_feature_ownership/OFFLINE_FEATURE_OWNERSHIP_DESIGN.md`
  with the fail-closed edge-to-feature candidate pipeline, rejection taxonomy,
  affected modules, and required proof gates.

### Offline API verification

- Verified through the SOLIDWORKS API MCP the contracts for
  `GetCorrespondingEntity2`, `GetCurveParams3`, adjacent faces, cylindrical
  surfaces, pattern seed features, face ownership/comparison, exact feature
  type mappings, Hole Wizard centre points, configuration suppression, and
  model-to-view point/vector transforms.
- Corrected the target specification's API ledger and implementation backlog to
  distinguish documented contracts from installed-build and fixture proof.

### Findings

- The current ordinate export still qualifies raw visible circles without
  model-feature ownership; the separate model audit is not connected to that
  path.
- The retained P-0251 probe counted 14 raw sketch points for the six-hole
  counterbore family, proving `ISketch.GetSketchPoints2` is not an admissible
  hole-instance count.
- The companion's simple-hole classifier misses the documented `SketchHole`
  type token, which maps to `ISimpleHoleFeatureData2`.

### Scope

- No VBA source, embedded macro, model, drawing, or manual reference was
  changed. No live SOLIDWORKS access was used.

## 2026-07-17 - Post-save source-authority and API checkpoint

### Changed

- Corrected the target specification and current-status record after detecting
  that `active_ordinate.swp` was saved at `16:28:35 +05:30`, later than the
  exported Modules 4 and 5. The latest embedded source is now recorded as
  unknown pending a user-operated export instead of being assumed stale.
- Made current embedded-source export and comparison the first gate before any
  replacement, compilation, or runtime recommendation.

### Offline API verification

- Requeried the linked SOLIDWORKS API MCP without accessing live SOLIDWORKS.
- Confirmed the documented contracts used by the exports and hybrid companion:
  `InsertModelAnnotations4`, component-qualified `GetVisibleEntities2`,
  `GetVisibleComponents`, datum-first `AddOrdinateDimension`, `MultiSelect2`,
  `SetPickMode`, `ActivateView`, and Boolean-returning `AlignDimensions`.
- Confirmed the annotation, ordinate-type, ordinate-error, and drawing-view enum
  values currently used by the project.

### Scope

- No VBA source, embedded macro, model, drawing, or manual reference was changed.

## 2026-07-17 - Runtime failure repair after Modules 4/5/6 test

### Fixed

- Corrected `Module4_ModelItemImporter.bas` so a context-dependent
  `SelectByID2=False` result no longer branches into an error handler and
  triggers VBA error 20 (`Resume without error`). The importer now follows the
  SOLIDWORKS 2025 example order, requires successful view activation, clears
  selection, and continues in the active drawing-view context.
- Replaced `GetVisibleEntities2(Nothing, edge)` in
  `Module5_FallbackDimensionEngine.bas`. The module now obtains the
  `Component2` objects returned by `IView.GetVisibleComponents` and requests
  visible edges separately for each component.
- Added explicit view-activation validation before ordinate processing.

### Evidence and validation

- The disposable probe inserted 11 annotations for `P-0251-14A-001` on
  SOLIDWORKS revision `33.1.2`.
- Live API evidence proved that the null-component call fails with a type
  mismatch, while the single visible component returned 39 visible edges.
- Added two VBA source-regression tests; the complete companion suite now
  passes 19 tests.
- Preserved the reported QA summary, Immediate Window, and settings screenshots
  under `test_assets/companion_evidence/2026-07-17_runtime_failure`.

### Still requires user-project validation

- The corrected `.bas` files are newer than `active_ordinate.swp`; replace the
  embedded Module 4 and Module 5 with the complete corrected files and compile
  the whole VBA project before rerunning the macro.
- The disposable Python probe validates the API contracts but does not compile
  or execute the embedded VBA project.

## 2026-07-17 - Hybrid SOLIDWORKS 2025 companion, phase 1

### Added

- Added a project-local, pinned fork of `wzyn20051216/solidworks-automation-skill`
  under `tools/solidworks-automation-companion` at upstream commit
  `de0bce46999ba96f996398aa5a588da14abd382b`.
- Added SOLIDWORKS 2025 symbolic constants, annotation import, feature audit,
  ordinate diagnostics, view policy, layout checks, title-block validation, and
  drawing-specific fail-closed QA modules.
- Added a disposable live API probe and 16 project-specific fake-COM/pure-logic
  tests. The upstream sketch-selection regression test remains in the suite.
- Added JSON evidence for the first live probe under
  `test_assets/companion_evidence`.

### Changed

- Replaced the fork's `InsertModelAnnotations3` drawing wrapper with
  `InsertModelAnnotations4`, a verified SOLIDWORKS 2025 annotation mask,
  `DuplicateDims=True`, anchor-view selection, whole-drawing import, selected-
  view retry, structured results, and cleanup.
- Corrected upstream template preference values from `24/25/26` to the
  SOLIDWORKS 2025 `swDefaultTemplatePart/Assembly/Drawing` values `8/9/10`.
- Repaired `src/active-ordinate/Module4_ModelItemImporter.bas` with the verified
  annotation mask, conditional hole-callout bit, duplicate handling, and
  baseline-style per-view fallback.
- Hardened `src/active-ordinate/Module5_FallbackDimensionEngine.bas` with strict
  view eligibility, explicit datum-first selection, selection-count validation,
  and `SetPickMode` cleanup.
- Changed `src/active-ordinate/Module6_QAEngine.bas` so dimension count alone can
  no longer produce an unconditional `PASS`; unresolved semantic/layout checks
  are listed explicitly.

### Not yet accepted

- The protected `src/baseline-model-dims` snapshot was not changed.
- The active macro has not yet been compiled in the VBA editor.
- The phase-1 live probe rerun was stopped at the user's request. A later
  runtime-failure investigation was explicitly authorized and is documented in
  the newer changelog entry above.
- Edge-to-feature ownership, proven model datum selection, sheet/layout bounds,
  title-block linkage, sections, and the three-part regression matrix still need
  live SOLIDWORKS validation before the macro is production-ready.
