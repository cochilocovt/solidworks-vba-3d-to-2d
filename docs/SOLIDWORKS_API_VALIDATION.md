# SOLIDWORKS API Validation of the Two Source Snapshots

## Scope and evidence

This report combines the repository analysis in `docs/ORDINATE_GAP_ANALYSIS.md` with API evidence returned by the configured `solidworks-api` MCP server.

Source abbreviations used throughout:

- **B**: `src/baseline-model-dims/`
- **A**: `src/active-ordinate/`

The MCP corpus identifies itself as **SOLIDWORKS 2026**, while this project targets **SOLIDWORKS 2025**. Values below are exact values returned by the MCP corpus; they are not inferred. Any item marked **verify in SW2025** must still be checked in the installed SOLIDWORKS 2025 Object Browser/type library before code is changed.

No VBA source file was edited, renamed, moved, or deleted during this analysis. The protected baseline remains unchanged.

## Three live findings from the first trunk ordinate runs (2026-08-06, r4-r7)

All three are **live evidence** on installed SOLIDWORKS 2025 SP1.2, obtained
by staged instrumentation in `CreateHoleOrdinateDims` across revisions r4-r7.
Together they took the ordinate engine from "throws on every view" to 8 chains
created.

### `IView.GetVisibleEntities2(Nothing, ...)` returns nothing usable

The baseline passed `Nothing` as `LpViewComponent`. MCP types that parameter
as `Component2`. Resolving the view's component through
`IView.GetVisibleComponents` and passing it produced **349 edges** where the
`Nothing` route produced none.

r7 reports `Ordinate edge route: ViaComponent`, so the component route is the
one actually answering. The `Nothing` fallback is retained but has never
succeeded on this build.

### `ISelectData.View` assignment raises error 91 with both operands valid

`Set swSelData.View = swView` raises runtime error 91 ("Object variable or
With block variable not set") even though:

- `swSelData` passed an explicit `Is Nothing` check immediately before;
- `swView` passed an explicit `Is Nothing` check immediately before, and
  `swView.Name` renders correctly in the error handler;
- `IDrawingDoc.ActivateView(swView.Name)` returned `True` first;
- `ModelDoc2.ClearSelection2 True` preceded the activation, matching the
  ordering in `src/active-ordinate/Module5_FallbackDimensionEngine.bas`,
  which carries an explicit comment that this ordering was the fix.

MCP documents the member as `View View {get; set;}` on `ISelectData`, so the
property exists and is settable. The failure is raised by SOLIDWORKS, not by
VBA binding.

**Worked around, not solved.** r6 made the assignment non-fatal: the view is
already activated, so an unscoped `SelectData` still selects into it. r7
created 8 ordinate chains with `Selection scope: Unscoped(err=91)`. Scoping
remains a correctness guard that is currently absent — **unresolved**, and a
candidate cause for any cross-view selection defect.

### `If Not <ICurve.IsCircle>` rejects every edge — confirmed in the trunk

The Boolean contract recorded under "2026-07-31 third run" is not historical
trivia; it silently disabled the whole ordinate engine.

r6: `Ordinate edges seen: 349 (circular: 0)` on a part with 12 holes.
The test was `If Not swCurve.IsCircle Then GoTo NextEdge`.

r7, after changing that single line to `If swCurve.IsCircle = False Then`:
`Ordinate edges seen: 349 (circular: 127)`, and chain creation began working.

This is direct live confirmation of the table in that earlier section: `Not`
on a SOLIDWORKS COM Boolean yields `-2`, which VBA treats as True. Nothing
errors; the branch simply always taken. **Treat any `If Not <comBooleanCall>`
in this codebase as a defect.** `Is Nothing` comparisons and genuine VBA
Boolean functions are unaffected.

### Still open after r7

`swCreateOrdDimErr_OrdFailure` (code 1) on 4 of 12 attempted chains, last seen
in `Section View J-J`. Not yet diagnosed.

## Two baseline defects located by contract reading (2026-08-05, planning)

Both were found by reading MCP contracts against `src/baseline-model-dims/`
during the 2026-08-05 reference-drawing planning session. Neither has been
reproduced live yet; both are contract-level, not runtime, evidence.

### `AddOrdinateDimension` leaves the document in ordinate pick mode

**MCP evidence**, `IModelDocExtension.AddOrdinateDimension` Remarks, quoted:

> Selections made immediately after calling this method continue to add
> ordinate dimensions to the group of ordinate dimensions. When you finish
> adding ordinate dimensions to the group, use `IModelDoc2::SetPickMode` to
> return to the default selection mode.

**Baseline use**: `B/Module5_FallbackDimensionEngine.bas:238-252`
(`CreateOneOrdinateChain`) calls `AddOrdinateDimension`, then `ClearSelection2`,
and returns. `SetPickMode` appears nowhere in `src/baseline-model-dims/`.

**Assessment**: the horizontal chain is created first
(`B/Module5_FallbackDimensionEngine.bas:86`), leaving the group open. The
vertical chain's `MultiSelect2` at `B:239` is therefore a selection made after
the call, which per the contract appends to the still-open horizontal group
rather than starting a vertical one. Every later selection in the run
(auto-arrange at `B/Module4_ModelItemImporter.bas:157`, title-block work) is
made in the same polluted mode. This is the leading candidate for the reported
"ordinate dims setting breaks the macro". `ClearSelection2` clears the
selection list; the contract does not say it closes the group.

`src/active-ordinate/Module5_FallbackDimensionEngine.bas` already calls
`SetPickMode` — the finding is specific to the baseline snapshot. Correcting
the note in this document's finding 5, which predates that change.

### `InsertModelAnnotations4` is being told to allow duplicates

**MCP evidence**, `IDrawingDoc.InsertModelAnnotations4` parameter 4:

- `DuplicateDims` (System.Boolean): "True to eliminate duplicate dimensions,
  false to allow duplicate dimensions"

The parameter name reads as the opposite of its effect; the description is the
contract.

**Baseline use**: `B/Module4_ModelItemImporter.bas:36-44` passes
`AllViews=True` and `DuplicateDims=False`, i.e. insert into every view and
allow duplicates. `B:90-98` (per-view fallback) passes the same `False`.

**Assessment**: this is a direct explanation for repeated hole callouts and
repeated dimensions across views. **Verify in SW2025** whether `True` alone
produces the reference drawing's single consolidated `6x`/`4x` callouts, or
whether per-view callout targeting is also required.

### Baseline import mask constants: all confirmed

Every `swInsertAnnotation_e` member declared at
`B/Module4_ModelItemImporter.bas:6-12` matches the MCP enum table exactly:
`swInsertDimensions=8`, `swInsertGTols=32`,
`swInsertDimensionsMarkedForDrawing=32768`,
`swInsertHoleWizardProfileDimensions=65536`,
`swInsertHoleWizardLocationDimensions=131072`,
`swInsertDimensionsNotMarkedForDrawing=524288`,
`swInsertholeCallout=1048576`. `swImportModelItemsFromEntireModel=0` also
matches `swImportModelItemsSource_e`. The baseline ordinate constants at
`B/Module5_FallbackDimensionEngine.bas:4-6` likewise match
`swAddOrdinateDims_e` (`swHorizontalOrdinate=3`, `swVerticalOrdinate=2`) and
`swCreateOrdDimError_e` (`swCreateOrdDimErr_Success=0`).

The gap analysis was right to demand verification of these; they pass. The
baseline's mask needs no numeric correction. **Verify in SW2025** still applies
to the values themselves.

### API-backed view classification exists

`B/Module2_DrawingPipeline.bas:283-285` classifies views by
`InStr(viewName, "ISO")`. Two API-backed alternatives exist: `IView.Type`
(`swDrawingViewTypes_e`: Sheet=1, Section=2, Detail=3, Projected=4,
Auxiliary=5, Standard=6, Named=7, Relative=8, Detached=9,
AlternatePosition=10) and `IView.GetOrientationName` ("Gets the predefined name
of this view"). Note that views created by `CreateDrawViewFromModelView3` are
expected to share one `Type`, so `Type` alone will not separate the front view
from the isometric — pair it with `GetOrientationName`. **Verify in SW2025**
what `Type` and `GetOrientationName` actually return for each view this
pipeline creates; this has not been probed.

## `IAnnotation.SetPosition2` shares `GetPosition`'s frame (2026-08-05, r62)

**MCP evidence**, `IAnnotation.SetPosition2` (not deprecated; `SetPosition`
is obsolete and superseded by it):

| item | value |
| --- | --- |
| signature | `Boolean SetPosition2(Double X, Double Y, Double Z)` |
| return | *"True if the position of the annotation is successfully set, false if not"* |
| frame | *"In a drawing, the x, y, z origin is relative to the origin of the drawing sheet (the lower-left corner of the sheet)"* |
| display-dimension origin | *"Point of leader attachment centered on a text box border / center point of bottom border of text box"* |

`IAnnotation.GetPosition` states the **same frame and the same
display-dimension origin**, and returns *"an array of 3 doubles"*. The two
are therefore directly comparable: a delta computed from `GetPosition` can
be passed to `SetPosition2` without a transform. This is what R23-823's
post-layout clamp relies on.

Two documented refusals that both look like a call that returned, and which
the clamp must therefore verify by readback rather than by return value:

- *"If this method attempts to set a position of an annotation that violates
  any restrictions, the annotation is placed as near as possible to the
  specified position."*
- *"Because radial and diametric dimensions are already attached to the end
  of a leader, this property is not available for these types of
  dimensions."* This matters directly here: the created `INNER_BORE_D40` and
  `FIT_BORE_D47_H7` dimensions are radial. Neither was outside the usable
  box in r61, so the case has **not** yet been exercised live.

`GetPosition` additionally warns that on failure *"the VARIANT SafeArray is
empty"*, so the array is checked before it is indexed.

Not yet live-proven. This is corpus evidence plus the r61 layout arithmetic;
the first mutating r62 run is what turns it into runtime evidence.

## `CreateSectionViewAt5` with `Options=0` ignores a jogged line (2026-08-05, r52)

**MCP evidence**, `swCreateSectionViewAtOptions_e`:

| member | value | meaning |
| --- | --- | --- |
| `swCreateSectionView_NotAligned` | 1 | section does not snap into alignment with the parent view |
| `swCreateSectionView_OffsetSection` | 2 | *"If set, then an aligned section view is created (two lines at an angle); if not set, a normal projection section view is created"* |
| `swCreateSectionView_ChangeDirection` | 4 | reverse the view direction |
| `swCreateSectionView_ScaleWithModel` | 8 | |
| `swCreateSectionView_Partial` | 16 | partial rather than complete section |
| `swCreateSectionView_DisplaySurfaceCut` | 32 | only surfaces cut by the line appear |
| `swCreateSectionView_ExcludeFasteners` | 64 | |
| `swCreateSectionView_CutSurfaceBodies` | 128 | |

`Module17_SectionPath.CreateSectionFromPath` passes `Options = 0`, so
SOLIDWORKS builds a **normal projection section** from a three-segment
jogged line.

**Measured**, run `macro_qa/20260805_050411_P-0251-14A-001`, decoded from
`IView.GetSectionLineInfo2` on Drawing View1:

```text
SECTION_LINE_SEGMENT|index=1|start=-0.102000000,0.000000000|end=0.008000000,0.000000000
SECTION_LINE_SEGMENT|index=2|start=0.008000000,0.000000000|end=0.008000000,-0.015000000
SECTION_LINE_SEGMENT|index=3|start=0.008000000,-0.015000000|end=0.088000000,-0.015000000
```

Segment lengths 0.110, 0.015, 0.080 match the path's waypoint spacing
exactly, so the drawing holds precisely the line that was asked for,
including the r51 overshoot (segment 1 runs 0.040 past the bore centre at
-0.062). The bore sits at transverse 0.000 and the six counterbores at
-0.015; the resulting section view contains the counterbore-column features
and **no bore geometry at all**. The cut is taken at one offset only, and
it is segment 3's.

**Consequence.** Lengthening segment 1 cannot change the section, which is
exactly what r51 measured: the line moved 40 mm and the section view was
byte-for-byte identical. A jogged section line needs
`swCreateSectionView_OffsetSection`. **Verify in SW2025** whether that flag
alone is sufficient or whether `swCreateSectionView_NotAligned` is also
required.

## `IView.GetSectionLineInfo2` on this build (2026-08-05, r52)

**Documented layout**, SOLIDWORKS 2025 Help: `[numSectionLines, layer,
numSegments, per segment (lineType, startPt[3], endPt[3]), arrowStart1[3],
arrowEnd1[3], arrowWidth1, arrowHeight1, arrowStyle1, arrowStart2[3],
arrowEnd2[3], arrowWidth2, arrowHeight2, arrowStyle2, textPt1[3],
textPt2[3], textHeight]`.

**Measured**, same run:
`numSectionLines=1|layer=-1|numSegments=3|segmentsDecoded=3|count=49|documentedTotal=53|tailMatchesDocumented=False`.

Two findings:

1. **The array is four doubles shorter than the documented layout** for
   three segments. The header and segment block at the front decode
   correctly and are confirmed by the waypoint match above; the arrow and
   text tail does not fit the documentation on this build. Do not index
   into the tail by the documented offsets without re-deriving them.
2. **The array mixes coordinate frames.** Segment endpoints are in the
   view's own frame - segment 1 starts at `-0.102`, which is not a sheet
   coordinate - while the tail holds sheet coordinates such as
   `0.107932223, 0.265060000`, inside Drawing View1's sheet outline
   `0.053992..0.137872, 0.061120..0.269000`. One array, two frames.

## `IView.UseSheetScale = 0` does NOT mean a different scale (2026-08-05, r50)

**MCP evidence**, `IView::UseSheetScale`, Remarks, verbatim:

> If the property is 0, then it is possible that the view scale is the same
> as the sheet scale.

Return value is documented as "1 if the view scale is the same as the sheet
scale, 0 if the view scale is independent of the sheet scale" - but the
Remarks contradict the strong reading of "independent", and the same page
names `IView::UseParentScale` as the separate member for tying a view to its
parent's scale. A section view created by `CreateSectionViewAt5` inherits
its parent view's scale, which is not the sheet flag.

**Consequence.** `Module9_LayoutEngine.bas:718` fails the required LAYOUT
stage whenever a non-isometric view reports `UseSheetScale <> 1`, and the
r49 run failed on exactly that one line for `Section View J-J`. That check
reads a flag as if it were a ratio. `IView.ScaleDecimal` is the ratio.

**Acted on at r50.** `ValidateLayout` accepts `UseSheetScale = 1` as before
and otherwise compares `IView.ScaleDecimal` against the sheet ratio proved
by `ISheet.GetProperties2`. The comparison fails closed - an unproved sheet
scale or an unreadable view scale is not a match - so a view genuinely drawn
at an unapproved scale still fails the stage. `VIEW_SCALE_READBACK|view=
|type=|isIsometric=|useSheetScale=|scaleDecimal=` is emitted for every view,
so the flag and the ratio can never be confused for one another again.

**Verify in SW2025.** The Remarks above are corpus documentation. The live
confirmation will be a run in which `Section View J-J` reports
`useSheetScale=0` with a `scaleDecimal` equal to the sheet's and LAYOUT no
longer fails on it.

## `IView.GetVisibleEntities2` really does exclude completely-obscured edges (2026-08-05, r48)

**Documented contract**, SOLIDWORKS 2025 Help: returns entities "not
completely obscured by other entities in the view".

**Measured behaviour**, run `macro_qa/20260805_033146_P-0251-14A-001`,
Drawing View1 (Front), the stepped bore of `P-0251-14A-001`:

```text
sourceFaces=2|facesProjected=2|boundaryEdges=4|circularEdges=4
|mappedEdges=4|inventoryConfirmed=0
|firstReject=MappedEntityNotInVisibleInventory
```

All four circular edges map through Route D: `IView.SelectEntity` accepts
each one and `ISelectionMgr.GetSelectedObjectsDrawingView2` returns Drawing
View1, so the entities exist in that view and belong to it. **Zero of the
four appear in `IView.GetVisibleEntities2(component, Edge)`.** The two
routes disagree in exactly the direction the Help predicts, and the same
view returns the counterbore edges through both routes in the same pass
(`mappedEdges=2|inventoryConfirmed=2`), which is the positive control.

**Consequences.**

1. The exclusion clause is real and load-bearing, not defensive wording.
   `GetVisibleEntities2` is a *visibility* inventory, not an ownership test;
   use `GetSelectedObjectsDrawingView2` when the question is ownership.
2. Membership of the inventory is a sound obscured/not-obscured
   discriminator for an entity already proved to belong to the view. This
   supersedes the r40-r41 attempt, whose model-space comparison matched
   nothing and therefore measured nothing.
3. `Module13_ProjectionResolution` correctly refuses an obscured edge as a
   circular dimension anchor. The stepped bore is hidden in every
   orthographic view of this part, so no orthographic anchor exists for it
   at all - it is a section-view feature, as the reference drawing shows.

## `IView.GetPolylines7` return array holds DRAWING entities (2026-08-04, r42)

**Documented contract**, SOLIDWORKS 2025 Help:

> Return Value: Array of modeling edges and silhouette edges corresponding to
> polylines in the view.

**Measured behaviour**, run
`macro_qa/20260804_235542_P-0251-14A-001`. The identical comparison helper
(`ISldWorks.IsSame` via `SafeObjectEquality`) was run twice against the same
`GetPolylines7` result, on twelve counterbore edges that
`IView.GetVisibleEntities2` had independently confirmed present in the same
view (`inventoryConfirmed=2` per location):

```text
mappedVisibleEdges=0          <- the MODEL edge from IFace2.GetEdges
mappedVisibleDrawingSpace=2   <- the DRAWING entity IView.GetCorrespondingEntity returned
```

Twelve of twelve in drawing space, zero of twelve in model space. The array
does not compare equal to the model edges the polylines were derived from; it
compares equal to the drawing-context entities of the same view.

**Consequence.** `Module13_ProjectionResolution.MapVisibleDatumEntity` tested
the **model** entity against this array from r33 until r42. It could never
match, so `visibleIndex` was always `-1`, and every vertical ordinate datum
failed closed as `PolylineVisibilityUnavailable` regardless of whether its
edge was visible. This also explains the long-standing scratch-versus-
production divergence: the scratch view returns `status=NoEdges` and takes
the documented scoped-selection fallback, so its datum resolves, while a
production view returns 39 entries that never match and fail closed.

**Adopted contract (r43).** Map first, then prove the **mapped drawing
entity** is present in the `GetPolylines7` array. The r37 fail-closed guard is
retained unchanged: a mapped entity absent from a *non-empty* array still
refuses (`DatumMap:MappedEntityNotVisible`), and only the documented
empty-array case falls back to `IView.SelectEntity` plus
`ISelectionMgr.GetSelectedObjectsDrawingView2` ownership proof.

**Still open.** This does not classify an edge that no route maps, because
drawing-space testing needs a drawing entity to test. The P-0251 stepped
bore remains unmapped and unclassified; whether its edges are obscured is
**not** established by this evidence.

## `IView.GetPolylines7` record layout, as used by the r50 inventory (2026-08-05)

**MCP evidence**, `IView::GetPolylines7`. The `Polylines` out-parameter is a
flat array of doubles, one record per polyline:

```text
[Type, GeomDataSize, GeomData[], LineColor, LineStyle, LineFont,
 LineWeight, LayerID, LayerOverride, NumPolyPoints, [x,y,z] * NumPolyPoints]
```

- `Type` 0 = polyline, 1 = arc or circle.
- `GeomData` for Type 1 is 12 doubles:
  `[cx,cy,cz, sx,sy,sz, ex,ey,ez, nx,ny,nz]`. Radius is not a field; it is
  the distance from centre to start.
- Six style scalars sit between `GeomData` and `NumPolyPoints`.
- The return value (entity array) is **positionally paired** with these
  records, and carries `Null` where a polyline renders a silhouette edge
  that no edge backs.
- Returns nothing at all when the view display mode is Shaded, Shaded With
  Edges, Draft Quality or Fast HLR/HLV.

**Live-confirmed on SW2025**, run
`macro_qa/20260805_041027_P-0251-14A-001`, Section View J-J:

```text
SECTION_GEOM_SUMMARY|decodeStatus=Complete|records=38|entities=38
  |recordsMatchEntities=True|doubles=1056|arcs=6|polylines=32
  |points=214|error=0
```

A walk that assumes the layout above consumed 1056 doubles exactly and
produced one record per returned entity, with no range guard tripping. The
documented layout holds on this build.

## `IView.GetPolylines7` returns VIEW-space coordinates (2026-08-05, r50)

**Measured**, same run, Section View J-J:

```text
SECTION_GEOM_FRAME|polylineBox=-0.009000,-0.098000,0.009000,0.098000
  |sheetOutline=0.289060,0.044385,0.318940,0.252265
```

The polyline points are centred on the view origin; the view sits at
sheet X 0.289 to 0.319. The two frames do not overlap at all, so the
points are in the **view's own coordinate system**, not sheet space. The
box measures 18 mm by 196 mm, which is the part's thickness and height.

**Consequence.** Anything that compares a `GetPolylines7` point against a
page coordinate - a projection's `PageX`/`PageY`, a view outline, a section
waypoint - must transform first. Mixing the two frames is the defect class
that already produced false clearance results in the section-line payload,
and it would be silent here because both frames are metres.

## `IDrawingDoc.ActivateView` return value (2026-08-04, r38)

**Documented contract**, `solidworks-api` MCP, SOLIDWORKS 2025 Help,
`IDrawingDoc.ActivateView(ViewName As String) As Boolean`:

> Return Value: True if successful, false if not.
> Remarks: This method returns false when trying to activate a drawing sheet.
> To activate a drawing sheet, use `IDrawingDoc::ActivateSheet`.

**Observed behaviour on this build**, run
`macro_qa/20260804_184514_P-0251-14A-001`. The setter returned False for a
named drawing view that did in fact become active, proved by reading
`IDrawingDoc.ActiveDrawingView` back and comparing `IView.GetName2`:

```text
ACTIVATE_VIEW|operation=ResolveOutlineDatum|view=Drawing View1|setterResult=False|readbackMatched=True
```

The same run shows the consequence at the call sites that trusted the raw
return instead:

```text
IMPORT_VIEW_NOT_ACTIVATED|view=Drawing View1
IMPORT_VIEW_NOT_ACTIVATED|view=Drawing View2
```

**Adopted contract.** The return value of `IDrawingDoc.ActivateView` is
recorded but is never a verdict. Activation is proved by active-view readback
in `Module8_RuntimeSupport.ActivateDrawingView`, which additionally retries
once through `IModelDocExtension.SelectByID2` with `"DRAWINGVIEW"` before
failing. This joins the existing `setterResult=False|readbackMatched=True`
family already adopted for sheet scale, display mode, rebuild, and sheet
activation on this build. r38 routes annotation import
(`Module14`), ordinate-group creation (`Module15`), native callout creation
(`Module16`), and the section cut (`Module17`) through it.
`Module13_ProjectionResolution` deliberately keeps the raw read: it reports
`PROJECTION_VIEW_ACTIVATION` and does not gate on it.

## `IWizardHoleFeatureData2.HoleDiameter` (2026-08-04, unresolved)

**Documented contract**, SOLIDWORKS 2025 Help: "Gets or sets the Hole Wizard
feature hole diameter." Remarks: "This property is not relevant for swTapered
and swTaperedDrilled holes. See Accessing Selections that Define Features."

**Observed behaviour.** After `IWizardHoleFeatureData2.AccessSelections`,
both P-0251 Hole Wizard features report every dimensional member as zero while
`DefinitionReadStatus` is `Read`:

```text
FEATURE_ACCEPTED|name=CBORE for M6 Socket Head Cap Screw1|...|readStatus=Read|diameterM=0.000000000|depthM=0.000000000|cboreDiameterM=0.000000000
```

Neither hole is tapered, so the documented exclusion does not apply. This
contract is **not resolved**; it needs a read-only probe of
`AccessSelections`' result and of the sibling `IWizardHoleFeatureData2.Diameter`
and thread members before `Module12_FeatureQualification.ReadHoleWizardDefinition`
is changed. No code was changed on this contract at r38.

## R23 Phase 11 production wiring (2026-08-04)

Phase 11 introduced no new SOLIDWORKS API contract. It composes only the
already validated R23 operations: feature cataloguing, projection Route D
(`IView.SelectEntity` plus selection-manager ownership readback), selected-view
annotation import, section creation, associative callout creation, and
content-envelope layout. The 2026-08-04 r26 runner proved guarded deployment,
full-project VBE compilation, and nine read-only P-0251 probes. This is not a
mutating production run, visual review, or manufacturing acceptance.

## Historical R23 Phase 9 scratch-layout entrypoint (2026-08-04, r27-r30)

No new SOLIDWORKS API contract was introduced. The controlled entrypoint in
`Module2_DrawingPipeline` reuses the existing, live-proved drawing guard:
`ISldWorks.ActiveDoc`, `IModelDoc2.GetType`, `IModelDoc2.GetPathName`,
`IDrawingDoc.GetCurrentSheet`, `ISheet.GetViews`, and
`IView.ReferencedDocument`. It accepted one exact disposable drawing path and
then applied the `ApplyR23ContentLayout` transaction. r31 retires that public
mutation route under the user-accepted layout policy; this section is retained
only as history of the earlier API-backed scratch work.

The r27 runner proved deployment/readback and full-project programmatic VBE
compilation on the scratch drawing. It did not invoke the entrypoint, so its
mutation behavior, visual result, and manufacturing suitability remain
unproven until the required manual VBE compile and authorized layout run.

The r28 retry changes only control flow around the existing, locally proved
`IView.ScaleDecimal` readback, `IView.Position`, and `EditRebuild3` calls. It
adds no API member, enum, coordinate transformation, or return-code contract.
It allows two measured scale passes because the first scratch run proved the
geometric estimate leaves text extents unchanged; a third request fails
closed as `LargerSheetRequired`.

## R23 accepted scratch persistence (2026-08-04)

The installed SOLIDWORKS 2025 SP1.2 interop assembly reports
`IModelDoc2.Save3(Int32, Int32&, Int32&) -> Boolean`. Its installed
`swSaveAsOptions_e` defines `swSaveAsOptions_Silent = 1`; the external,
compiled dispatch helper saved only the exact isolated P-0251 scratch with
that option. Live result: `succeeded=True|errors=0|warnings=0`. This helper
is not managed VBA source and does not widen the authorized drawing scope.

## R23 Phase 0 live fixture evidence (2026-07-31)

This section records installed-build behavior separately from the MCP corpus.
It is from the read-only P-0251 probe log
`test_assets/iteration_evidence/r23/20260730-075811/live-probes/R23_FEATURE_20260731_040539.log`
on SOLIDWORKS 2025 SP1.2 (`RevisionNumber=33.1.2`).

- `GetTypeName2="ICE"` plus `GetTypeName="Cut"` was observed for
  `Cut-Extrude1`, `Cut-Extrude3`, and `Cut-Extrude4`. Each successfully exposed
  typed `IExtrudeFeatureData2` selection access and release.
- The CBORE Hole Wizard yielded six sketch points, M6, an 11 mm counterbore,
  and a 6 mm counterbore depth. The M5 Hole Wizard yielded two source points,
  M5x0.8, 5 mm thread diameter, and 10 mm thread depth. `Mirror1` owned
  cylindrical faces whose seed feature was the M5 Hole Wizard.
- `IsSuppressed2(swSpecifyConfiguration, Array("Defualt"))` returned Empty for
  every feature in this VBA run. `IFeature.IsSuppressed` then reported active
  state for the active configuration. This fallback is sufficient only for
  active-model diagnostics; R23 must continue to fail closed if a drawing
  view's referenced configuration cannot be proved.
- On complete counterbore, tapped, mirrored, and cut-cylinder boundaries, both
  tested orders preserved `Curve.IsCircle=True` and complete
  `Edge.GetCurveParams3` evidence (`UMin=0`, `UMax=2pi`, coincident endpoints).
  The probe's `CircleParams` helper nevertheless emitted `SkippedNotCircle`
  despite true predicates before and after. R23 therefore must not make
  `CircleParams` a load-bearing production dependency.
- `InsertModelAnnotations4` with mask `18055274` and `DuplicateDims=True`
  returned 25 unique identities in each tested transaction. A selected-primary
  `AllViews=True` call placed 8 in the primary and 17 in the section, but 0 in
  the side. Explicit section, side, then primary calls with `AllViews=False`
  placed 17, 2, and 6 respectively, with 0 in the isometric. This installed
  build evidence selects the explicit per-view transaction for R23.
- The import produced one M5 Hole Wizard callout in the section, including
  4.2 mm tap-drill diameter, 12.4 mm drill depth, `M5x0.8`, and 10 mm thread
  depth. No M6 counterbore callout, H7 fit, or nonzero tolerance was imported.
  Native import is therefore a first-priority request, not proof that every
  required manufacturing callout will be present.

## R23 Phase 2 live catalog evidence (2026-08-01)

Read-only `R23_ProbeFeatureCatalog` run on P-0251, configuration `Defualt`
(the model's own spelling), SOLIDWORKS 2025 SP1.2. `modelUnchanged=True`;
47 features visited, 6 accepted, 0 warnings, 0 failures.

- `IWizardHoleFeatureData2.HoleFit` returned `1` for the M6 counterbore and
  `-1` for the M5x0.8 tapped hole. The 2025 Help states the property returns
  `swWzdHoleScrewClearanceTypes_e` and "applies to counterbore and
  countersink Hole Wizard features only"; the MCP corpus gives that enum as
  exactly close `0`, normal `1`, loose `2`. `-1` is therefore outside the
  enum and means not applicable, **not** a fit value. R23 records
  out-of-enum codes as absent.
- `IFace2.GetSeedFeature` resolved every mirrored face to the M5 Hole Wizard
  seed. It supplies the seed **identity only** — no manufacturing values —
  so a pattern instance's semantics must be read from the seed through the
  same typed readers.
- The axial interval separated two coaxial pairs of M5 tap-drill cylinders
  (radius `0.0021`, i.e. the 4.2 mm tap drill) that share a line key but
  occupy `-0.036..-0.0236` and `0.0236..0.036`. These are opposite blind
  holes on one axis and correctly resolved to four distinct physical
  locations rather than two. This is the case the Plücker line key alone
  cannot decide.
- 18 qualifying cylindrical faces consolidated to 11 physical locations
  through 7 merges: six M6 counterbores (radii `0.0055`/`0.0033`, two faces
  each), one Ø47/Ø40 stepped bore (radii `0.0235`/`0.0200`, two faces), and
  four single-face M5 tap drills.
- `ICosmeticThreadFeatureData` features `Hole Thread1`/`Hole Thread2` own no
  faces and were rejected `NoOwnedGeometry`. Their thread data is already
  carried by the owning Hole Wizard feature, so nothing is lost, but R23
  must not expect cosmetic threads to contribute geometry.
- **`ObjPtr` is not a stable feature identity across a traversal.** Two runs
  over the unchanged part visited 47 then 46 features, and the sketches
  visited twice differed between runs. A feature reached both as a tree
  entry and as a consuming feature's subfeature can arrive through two COM
  wrappers at different addresses. Combined with the earlier finding that
  wrapper addresses are also *reused* across distinct objects, `ObjPtr` is
  unusable as an identity component in either direction. `IFeature.Name`
  plus `GetTypeName2` is the exact-once key within a part document.

## R23 Phase 3 first live drawing run (2026-08-01)

Read-only `R23_ProbeViewProjections` on the P-0251 drawing. Installed-build
behaviour, recorded separately from the MCP corpus.

- `ISheet.GetViews` returned **14 views**, six of which are the sheet's
  standard-view placeholders (`*Left`, `*Bottom`, `*Current`, `*Isometric`,
  `*Dimetric`, `*Trimetric`). All six returned zero entities from
  `GetVisibleEntities2`. Any per-view pass must expect them.
- `IView.GetVisibleComponents` returned exactly one component for every real
  view of this **part** drawing, named `P-0251-14A-001-67` for
  `Drawing View4` and `P-0251-14A-001-69` for `Drawing View7`.
  `IView.GetVisibleDrawingComponents` returned nothing on every view, which
  matches the 2025 Help limiting it to assembly drawings.
- `Section View J-J` reported its visible component as
  `P-0251-14A-001-SectionAssembly-3-1/P-0251-14A-001-1`. A section view is
  backed by a synthetic section assembly, consistent with the Help's note
  that section views have no `ReferencedDocument`.
- `GetVisibleEntities2(component, swViewEntityType_Edge)` returned 64, 68 and
  53 edges for `Drawing View4`, `Section View J-J` and `Drawing View7`.
- `IView.ReferencedConfiguration` returned `Defualt` with
  `configurationProven=True` on all 14 views, including the placeholders.
- **`IEdge.GetCurveParams3` returns an `ICurveParamData` OBJECT**, not an
  array of doubles — unlike `ICurve.CircleParams` and
  `ISurface.CylinderParams`, which do return 7-element double arrays. It
  must be bound with `Set`; a `Let` assignment into a Variant raises **error
  438** because `ICurveParamData` has no default member. Its seven members
  are `CurveTag`, `CurveType`, `EndPoint`, `Sense`, `StartPoint`,
  `UMaxValue`, `UMinValue`, and the 2025 Help requires `IEdge::GetCurve` to
  be called first. Complete-circle proof therefore comes from
  `.StartPoint`/`.EndPoint` coincidence plus `.UMinValue`/`.UMaxValue`, which
  is what Phase 0 recorded as `UMin=0, UMax=2pi`.
- `swDrawingViewTypes_e` observed on this build: the ten sheet placeholders
  report `7` (`swDrawingNamedView`) with zero visible entities; real
  projected views report `4` (`swDrawingProjectedView`) and the section view
  `2` (`swDrawingSectionView`). The sheet has four real views —
  `Drawing View2` (65 edges), `Drawing View4` (64), `Section View J-J` (68)
  and `Drawing View7` (53).
### Third run: correspondence and selection proved (2026-08-01)

- **`IView.GetCorrespondingEntity` works on a model reached through
  `IView.ReferencedDocument` with the drawing active.** 26 anchors resolved
  through route A, each identity-confirmed against the view's
  `GetVisibleEntities2` inventory by `ISldWorks.IsSame`. Route B was never
  needed. The earlier hypothesis that correspondence required the part as
  the active document is disproved.
- Mapping is per-edge **and view-dependent**: the same location yielded
  `mappedEdges` of 0 to 4 out of the same 4 candidate edges in different
  views. Production must try every edge of every contributing face in every
  view independently, never caching a per-location result across views.
- **Hidden geometry does not map.** In the side view only the near column of
  counterbores (`py = -0.015`) and the near pair of tapped holes resolved;
  the far ones returned Nothing from route A rather than being filtered by
  route C.
- **A section view maps nothing.** All 11 locations returned
  `mappedEdges=0` in `Section View J-J`, whose visible component is the
  synthetic `…-SectionAssembly-3-1/…`. Model edges of the original part have
  no correspondent there; section-view annotation needs a different route.
- **`ISelectData.View` raised error 91 on all 26 selection attempts**,
  confirming the deviation this repository recorded for
  `Module2_DrawingPipeline.CreatePrimarySection` on a fresh code path. With
  the binding skipped, `IEntity.Select4(False, selectData)` still selected
  into the correct view every time:
  `ISelectionMgr.GetSelectedObjectsDrawingView2(1, -1)` returned the
  expected view for 26 of 26. Proving ownership after the fact is therefore
  a sufficient substitute for the unavailable binding.
- `ICurve.IsCircle` plus `IEdge.GetCurveParams3` reproduced the Phase 0
  proof exactly on drawing-mapped model edges: `uMin=0`,
  `uMax=6.283185307`, endpoint gap `0.000000000`.

### Fourth run: axis normality and a VBA out-parameter trap (2026-08-01)

- **Normal-axis classification confirmed against four views.**
  `IView.ModelToViewTransform`, applied to two model points and differenced,
  classified 7 Z-axis holes as normal in the front view, 4 Y-axis tapped
  holes as normal in both the side and section views, and **0 in the
  isometric**. The isometric result is the strongest check available: an
  isometric has no hole axis normal to the sheet, and the test agreed.
- **VBA trap, not a SOLIDWORKS one, but it corrupts API evidence.** A class
  module's `Public` variable is exposed as a property, so passing
  `projection.ProjectedAxisX` as a `ByRef` out-parameter gives the callee a
  temporary that is discarded on return. The run logged
  `projectedAxis=0,0,0` on every line while the function's return value was
  correct. Any R23 code taking geometry out of a helper must use locals and
  assign the class field explicitly.
- `IView.GetDisplayMode2` returns `swDisplayMode_e`, where `swHIDDEN=2` is
  Hidden Lines Removed and `swHIDDEN_GREYED=1` is Hidden Lines Visible.
  Under HLR a far-side hole has no drawing entity, so
  `GetCorrespondingEntity` returns Nothing and no search strategy can anchor
  it. This must be read before attributing an unanchored location to a
  mapping defect.
- Anchor selection preferred the Ø6.6 through hole (`edgeRadiusM=0.0033`)
  over the Ø11 counterbore mouth on all six counterbores, confirming that
  the smallest coaxial radius is the right primary-diameter rule for a
  manufacturing callout.

### Fifth run: coincident projections are one drawing entity (2026-08-01)

- **`IView.GetCorrespondingEntity` maps at most one model edge per drawing
  entity.** Where two coaxial holes are viewed along their common axis they
  occupy ONE page point and SOLIDWORKS holds ONE edge, so only one of the
  two model edges maps. On P-0251's side view the six counterbores collapse
  to three page points and the four tapped holes to two; the observed
  mapped counts were exactly 3 and 2. **No search strategy can produce more
  anchors than the drawing has entities**, so a coverage requirement must be
  expressed per distinct page position, not per physical location.
- **This is not a hidden-line effect.** `IView.GetDisplayMode2` reported
  `HiddenLinesVisible` for both `Drawing View4` and `Drawing View7`, so the
  far-face holes are drawn and still do not map. An earlier entry in this
  document hypothesised Hidden Lines Removed as the cause; that hypothesis
  is **disproved**. Observed modes on this drawing: the two projected views
  `HiddenLinesVisible`, `Drawing View2` and `Section View J-J`
  `HiddenLinesRemoved`, and all ten sheet placeholders `Shaded`.
- **`IView.ModelToViewTransform` verified against a known-answer case.** The
  isometric view returned a projected axis of
  `0.707107, 0.408204, -0.577382` for a model Z axis — exactly
  (1/√2, 1/√6, −1/√3) with magnitude 1.000, the standard isometric direction
  cosines. Orthographic views returned clean `0,0,±1` and `±1,0,0`.

## R23 Phase 3 contract checks (2026-08-01)

Corpus evidence, checked before the projection engine was written. Not
runtime proof.

- **`ISldWorks.IsSame` returns `System.Int32` as `swObjectEquality`, not a
  Boolean.** Members are exactly `swObjectNotSame=0`, `swObjectSame=1`,
  `swObjectUnsupported=2` ("unable to determine if the specified objects are
  the same object"). Any Boolean coercion makes `2` truthy and turns an
  undecidable comparison into a false identity match. R23 compares to `1`
  explicitly and defaults unreadable comparisons to not-same.
  **Amended 2026-08-01:** defaulting `2` to not-same is correct behaviour but
  silently discards a distinction that matters in diagnosis. Reconciliation
  compares an entity resolved out of a drawing view against faces and edges
  held from the part document; if that comparison is *undecidable* rather
  than negative, every match fails for a reason no evidence line was
  reporting. `RecordEquality` in `Module14_AnnotationImport.bas` keeps the
  raw code and publishes the strongest one seen as `eqMax`. Whether `2`
  actually occurs on this build is **not yet established** — the next live
  run decides it.
- **`IView.GetVisibleDrawingComponents` is documented as "unobscured drawing
  components in this drawing view of an *assembly* drawing"**, returning
  `IDrawingComponent`, whose `.Component` gives "a component object that
  fully supports all of the `IComponent2` methods and properties". A part
  drawing is therefore expected to supply the `GetVisibleComponents` handle
  and not this one; R23 records that as `DrawingContextOnly` context rather
  than a failure.
- `swViewEntityType_e` is exactly edge `1`, vertex `2`, face `3`, silhouette
  edge `4`.
- `IView.GetCorrespondingEntity(Object)` accepts "a vertex, face, or edge
  entity in the part or assembly" and returns "null or Nothing if none
  found". The Phase 0 per-edge finding stands: accepting the documented face
  route without trying edges would fail every counterbore.
- `ISelectionMgr.GetSelectedObjectsDrawingView2(Index, Mark)` is 1-based with
  `Mark = -1` meaning all selections regardless of marks. Index `-1` is
  reserved for the dynamically highlighted view and is not used here.

## R23 Phase 4 first live reconciliation run (2026-08-01)

Read-only run on the manual reference drawing. `mutations=0`,
`drawingUnchanged=True`.

- **`IDimension.GetSystemValue3` returns 0 for a DRAWING dimension.** All 38
  inventoried annotations reported a nominal of 0 when read with
  `GetSystemValue3(swThisConfiguration, Empty)`. These are drawing-owned
  reference dimensions (`RD4@Drawing View6@P-0251-14A-001.Drawing`) and a
  drawing document has no configurations, so a configuration-scoped read has
  nothing to resolve. `IDimension.SystemValue`, which takes no configuration,
  is the correct member for drawing dimensions.
- **The Ø47 H7 fit exists only in the drawing.** `RD4@Drawing View6` in
  `Section View J-J` returned `swTolFITWITHTOL` (8) with a nonzero tolerance
  and fit data, while Phase 0's direct read of the model's `D1@Sketch4` found
  neither. Fits on this drawing set are drafting-stage, not model-derived.
- **The reference drawing authors section diameters as `swLinearDimension`
  (2)**, including the H7 one. Phase 0 showed *imported* section diameters
  arrive as `swDiameterDimension` (6). The two authoring routes produce
  different types for the same physical measurement, so dimension type alone
  must never be used to decide that something is a diameter.
- `IView.GetAnnotations` per view on this drawing: 2, 20, 8, 8. Types seen:
  display dimension 4, cosmetic thread 1, note 6, and center-mark symbol 13.
- `IAnnotation.GetAttachedEntities3` works on drawing-owned reference
  dimensions: one hole callout reconciled to its physical location by
  attached-entity identity against the projection anchor.

## R23 Phase 4 second run: value and tolerance reads confirmed (2026-08-01)

- **`IDimension.SystemValue` is the correct member for a drawing dimension;
  `GetSystemValue3` returns 0.** Confirmed across all 30 display dimensions
  in one run, each reporting both readings side by side. A drawing document
  has no configurations, so any configuration-scoped value read is empty.
- **`IDimensionTolerance` reads a named ISO fit correctly.** The reference
  drawing's `47 H7 +0.025/0` returned `Type=8` (`swTolFITWITHTOL`),
  `GetMaxValue2` `0.000025000` with status 0, `GetMinValue2` `0.000000000`
  with status 0, and `GetHoleFitValue` the literal string `H7`. Dimensions
  without a tolerance returned `Type=0` and status `1/1` on both bounds, so
  status 1 is the "not applicable" marker.
- **`IAnnotation.GetAttachedEntities3` returns 1 to 3 entities** for drawing
  reference dimensions, center marks and cosmetic threads on this drawing;
  only the free-standing note returned 0.
- **A native hole callout does not attach to the same drawing entity R23
  prefers as an anchor.** The counterbore callout attaches to the 11 mm mouth
  edge; R23's anchor tier picks the 6.6 mm through hole. Both are mapped
  aliases of one physical location, so reconciliation must test every mapped
  entity rather than the anchor alone.
- Ordinate chains read as `swOrdinateDimension` (1) with the chain's base and
  one `swVertOrdinateDimension` (8); the zero members return a nominal of 0,
  which is a genuine value rather than a failed read.

## R23 Phase 4: forward correspondence is partial (2026-08-01)

- **`IView.GetCorrespondingEntity` returns a SUBSET of a location's drawing
  edges.** For each P-0251 counterbore it mapped 2 of the 4 model boundary
  edges. The drawing edge that SOLIDWORKS' own native hole callout attaches
  to — confirmed as `swSelEDGES` by `IEntity.GetType` — was **not** among the
  18 aliases the forward route produced for `Drawing View4`. Forward
  correspondence therefore cannot be used to reconcile pre-existing
  annotations to geometry.
- In `Section View J-J` the forward route produced **no** anchors or aliases
  at all, so every annotation there is unreachable that way.
- **`IModelDocExtension.GetCorrespondingEntity2` is the documented reverse
  member**: it takes a vertex, face or edge in a drawing view or assembly
  and returns the corresponding entity in the underlying part, or Nothing.
  The Help for `IView.GetCorrespondingEntity` explicitly points to it for
  this direction. It requires no drawing-view anchor, so it remains
  available where the forward route fails.
  **Live result, 2026-08-01 — the reverse route does not work in a part
  drawing.** Across all 38 annotations of the P-0251 reference drawing and
  every attached entity, the call returned Nothing with **error 0**:
  `outcomes=draw1:unresolved:err0`, `resolved=0`. It declines rather than
  raising. This holds for clean `swSelEDGES` attachments in ordinary
  projected views, not only in the section view, so it is not a section-cut
  artefact. Reading the contract again against that result: the member
  returns the entity "in the **underlying part or subassembly**", i.e. it
  descends a component level. A part drawing has no component to descend
  into — matching the `componentContext=DrawingContextOnly` this project
  already records — so there is nothing for it to return. Retained in
  `Module14_AnnotationImport.bas` because it is the documented direction and
  is expected to work in assembly drawings; it now reports
  `reverse=UnavailableNoModelCounterpart` so the distinction is explicit in
  evidence.
  **Corollary:** `eqMax=-1` on every annotation means no `IsSame` comparison
  ever executed, which closes the `swObjectEquality` Unsupported (2)
  hypothesis raised earlier the same day. Cross-document comparison was
  never reached and is not implicated.
- `IEntity.GetType` returns `swSelectType_e`; values observed on attached
  entities were 1 (`swSelEDGES`) and 0 (`swSelNOTHING`, for attachments that
  do not resolve to geometry).

## R23 Phase 4 contract checks (2026-08-01)

Corpus evidence, checked before the import engine was written. Not runtime
proof.

- **The Phase 0 mask `18055274` decomposes with no unaccounted bit** against
  `swInsertAnnotation_e`: `swInsertDatums` 2, `swInsertDimensions` 8,
  `swInsertGTols` 32, `swInsertNotes` 64,
  `swInsertDimensionsMarkedForDrawing` 32768,
  `swInsertHoleWizardProfileDimensions` 65536,
  `swInsertHoleWizardLocationDimensions` 131072, `swInsertholeCallout`
  1048576, `swInsertTolerancedDims` 16777216. The callout member is spelled
  with a lowercase h.
- **`InsertModelAnnotations4` is on `IDrawingDoc`, not
  `IModelDocExtension`.** No `InsertModelAnnotations3` exists on the
  extension. It takes eight arguments — Option, Types, AllViews,
  DuplicateDims, HiddenFeatureDims, UsePlacementInSketch,
  InsertAllAnnotations, InsertAllReferenceGeometry — and returns an **array
  of inserted `IAnnotation` objects**, not a count. `DuplicateDims=True`
  means *eliminate* duplicates.
- **`IDimensionTolerance.GetMinValue2` and `GetMaxValue2` return the status**
  (`swDimensionToleranceWarning_e`) and deliver the value through an out
  parameter. The out parameter must be a local variable: a class `Public`
  field is exposed as a property and the write-back would be discarded, the
  same trap recorded for the fourth Phase 3 run.
- **`swTolType_e` gives `swTolFIT` and `swTolMETRIC` the SAME value, 7.** A
  tolerance type of 7 therefore cannot be reported as one rather than the
  other. Fit-bearing types are 7, `swTolFITWITHTOL` 8 and `swTolFITTOLONLY`
  9.
- `GetAttachedEntities3` is a member of `IAnnotation`, not
  `IDisplayDimension`; reach it via `IDisplayDimension.GetAnnotation` or
  from an inventory that already holds the annotation.
- `IDimension.GetSystemValue2` is marked **obsolete**; `GetSystemValue3`
  takes `swInConfigurationOpts_e` (`swThisConfiguration` = 1) and returns a
  Variant.
- `swAnnotationType_e` values used: display dimension 4, GTol 5, note 6,
  datum tag 2, datum origin 16, cosmetic thread 1.

### Installed-contract checks for the final R23 probes

The final Phase 0 probe source was checked against installed SOLIDWORKS 2025
SP1.2 interop `33.1.2.4`. This is installed-interface evidence, not runtime
proof:

- `IModelDocExtension.AddOrdinateDimension(Int32, Double, Double, Double)`
  returns `Int32`; installed `swAddOrdinateDims_e` values are vertical `2` and
  horizontal `3`.
- Installed `swCreateOrdDimError_e` values are success `0`, general failure
  `1`, no internal dimensions `2`, bad selection `3`, model not loaded `4`,
  same-part-only `5`, extra selection `6`, ordinate failure `7`, duplicate
  `8`, bad direction `9`, and undefined `-1`.
- `IEntity.Select4(Boolean, ISelectData)` returns `Boolean`;
  `IModelDoc2.SetPickMode()` returns `Void`; and
  `ISelectionMgr.GetSelectedObjectCount2(Int32)` plus
  `GetSelectedObjectType3(Int32, Int32)` expose cleanup/selection readback.
- `IView.GetVisibleComponents()`, `GetCorrespondingEntity(Object)`,
  `GetDisplayDimensions()`, `GetSection()`,
  `GetSectionLineCount2(Int32 ByRef)`, and `GetSectionLineInfo2()` are present.
- Installed dimension types include linear `2`, diameter `6`, horizontal
  ordinate `7`, vertical ordinate `8`, and diametric-linear `15`.
  `IDimension.Tolerance` exposes type, fit type/style, hole/shaft fit, and
  min/max status/value; `IDimension.GetToleranceFitValues()` returns `String`.
  `IDisplayDimension.GetDisplayData()` exposes the line/arrow/text primitives
  used by the probe.
- `IDrSection.GetLabel()`, `GetName()`, `IsAligned()`, `SectionDepth`,
  `GetLineInfo()`, `GetArrowInfo()`, and `GetTextInfo()` are present.
- `IDisplayData.GetLineAtIndex3(Int32)`,
  `GetArrowHeadAtIndex2(Int32)`, `GetTextAtIndex(Int32)`,
  `GetTextPositionAtIndex(Int32)`, `GetTextHeightAtIndex(Int32)`, and
  `GetTextAngleAtIndex(Int32)` match the probe's readback calls.

The local `solidworks-api` MCP was not callable in this session. The probe uses
the already recorded official selection-order/`SetPickMode` contract and the
installed signatures/numeric values above; live compilation and execution
remain mandatory.

### Final drawing-contract probe runtime (2026-07-31)

The user compiled and ran both disposable final-probe entry points on the
authorized P-0251 fixture. These observations are installed-build runtime
evidence; they do not replace official API contracts or production acceptance.

- The ordinate entry point obtained the primary view's visible component and
  then stopped with `CounterboreLocationsUnavailable`. No datum/entity
  selection and no `AddOrdinateDimension` call occurred. Consequently,
  `IView.GetCorrespondingEntity` behavior and the ordinate transaction remain
  unproved. The next read-only diagnostic must expose each face/edge rejection
  and compare direct active-part mapping, component-mediated mapping, and
  `IView.GetVisibleEntities2(component, Edge)`.
- The imported 47 mm item is `D1@Sketch4` and the imported 40 mm item is
  `D1@Sketch6`; both report installed dimension type `6`
  (`swDiameterDimension`). Neither reports H7 or a nonzero tolerance. The
  source model's `D1@Sketch4` tolerance must be read directly before choosing
  whether model data or an explicitly approved target/reference specification
  is authoritative.
- `GetSectionLineCount2` reported one line with size 49 and
  `GetSectionLineInfo2` returned 49 items, so the flattened payload passed the
  structural count/size check. The probe's per-dimension loop did not reset all
  target and nominal fields, however, and its later derived labels are not
  accepted evidence.
- The live values show that the three J-J segment endpoints are expressed in
  source-view sketch coordinates, while the arrow and label positions are in
  page coordinates. This coordinate-frame conclusion is an inference from the
  installed runtime payload, the section-creation path, and the full-sheet
  screenshot; the unavailable MCP did not independently confirm it. Production
  `Module6_QAEngine` currently compares the raw segment values with a
  page-coordinate part-identification extent, so that clearance check is
  invalid until the segment endpoints are transformed to page coordinates
  exactly once.

The accepted logs, hashes, screenshots, and required probe corrections are
indexed in `docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md`.

### 2026-07-31 corrected-probe MCP contract lookups

The local `solidworks-api` MCP was callable again in this session. The
corrected disposable drawing-contract probes rely on the following MCP
contracts (compatibility-snapshot evidence; numeric values remain subject to
installed-build confirmation, and live compilation/execution remain
mandatory):

- `IView.GetCorrespondingEntity(Entity)` accepts a vertex, face, or edge
  entity in the part or assembly and returns the corresponding drawing-view
  entity, or Nothing when none is found. Its Remarks point to
  `IModelDocExtension.GetCorrespondingEntity2` for the reverse direction.
- `IComponent2.GetCorrespondingEntity(Entity)` returns the corresponding
  entity in the context of the component; the corrected probe uses it as the
  component-mediated mapping route before view correspondence.
- `IView.GetVisibleEntities2(LpViewComponent, EntityType)` requires a
  `Component2` plus `swViewEntityType_e`; visible entities are entities not
  completely obscured in the view.
- `ISldWorks.IsSame(Object1, Object2)` returns `swObjectEquality`:
  `swObjectNotSame=0`, `swObjectSame=1`, `swObjectUnsupported=2`. The probe
  logs the raw value so an unsupported comparison is distinguishable from a
  negative identity.
- `IModelDoc2.Parameter(StringIn)` accepts the fully qualified dimension
  name (for example `D1@Sketch4`) and returns the `IDimension`; this is the
  read-only route for the direct part-source H7 readback.
- `IDimension.FullName` returns `<Dimension>@<Feature>@<Model>`.
- `IView.GetXform` returns three doubles: view X/Y location relative to the
  sheet origin plus the drawing-view scale. Together with `IView.Angle`
  this defines the page-to-view-sketch conversion whose exact inverse the
  corrected probe applies exactly once to returned section segments.
- `INote.GetExtent` returns six doubles describing the note's lower-left
  and upper-right box extents in sheet space, relative to the drawing
  origin at the lower-left corner. The probe measures the template's
  `*P-0251-14A-001*` part-identification note with it.
- `IDisplayDimension.GetText(WhichText)` accepts `swDimensionTextParts_e`
  (`swDimensionTextPrefix=1`, `swDimensionTextSuffix=2`, callout parts 3-8;
  `swDimensionTextAll=0` is SetText-only) and does not support hole
  callouts.
- `IFeature.GetFirstDisplayDimension` requires
  `IModelDoc2.SetUserPreferenceToggle` with `swDisplayFeatureDimensions`
  before feature display dimensions are returned, and may not return the
  same display dimension on every call. Because toggling a document
  preference would mutate fixture document state, the corrected probe does
  not traverse model display dimensions; it reads the model `IDimension`
  through `IModelDoc2.Parameter` instead.
- `swTolType_e` returned a suspect duplicated table row
  (`swTolFIT=7` alongside `swTolMETRIC=7`), so the probe logs raw
  tolerance-type numbers and treats any name decoding as informational
  only.

### 2026-07-31 fifth run: ordinate transaction proved on the installed build

Installed-build runtime evidence. Both datum-first ordinate groups completed
against the authorized P-0251 fixture.

- `IModelDocExtension.AddOrdinateDimension` returned `0`
  (`swCreateOrdDimError_Success`) for both the horizontal request
  (`swAddOrdinateDims_Horizontal = 3`) and the vertical request
  (`swAddOrdinateDims_Vertical = 2`).
- Selection counts were exact: datum + 2 = 3 for X, datum + 3 = 4 for Y.
  The X datum selected as type `1` (edge) and the Y datum as type `3`
  (vertex).
- Display-dimension counts moved 6 → 8 and 8 → 11, matching the two and
  three appended coordinates.
- `IModelDoc2.SetPickMode` plus `ClearSelection2` left zero selections after
  each group, and the fixture save flag was unchanged.

#### `ISelectData.View` is unavailable — confirmed, with a working substitute

Both groups logged `viewBinding=UnboundAfterError:91` and then completed
normally. This confirms the earlier inference: the documented get/set
assignment raises runtime error 91 on this build, while activating the view
with `IDrawingDoc.ActivateView` beforehand and verifying each selection with
`ISelectionMgr.GetSelectedObjectsDrawingView2` is sufficient. Every selection
in both groups returned `ownerView=Drawing View1`.

#### Created ordinate `Type2` values are `1` and `7`

`IDisplayDimension.Type2` readback of the five created ordinates:

| Creation enum | Observed `Type2` | Page positions |
|---|---:|---|
| `swAddOrdinateDims_Horizontal = 3` | `1` (`swOrdinateDimension`) | `(0.080932, 0.048942)`, `(0.110932, 0.048942)` |
| `swAddOrdinateDims_Vertical = 2` | `7` (`swHorOrdinateDimension`) | `(0.041996, 0.077060 / 0.117060 / 0.157060)` |

The observed values do not follow the enum names, and no interpretation is
offered here beyond the measurement. The operational rule for production QA
is to accept `1`, `7` and `8` when classifying ordinate display dimensions;
a filter limited to `7` and `8` would have missed both X ordinates.

Geometry is correct in both cases: the X ordinates sit on a horizontal
baseline below the view at the two column coordinates, and the Y ordinates on
a vertical baseline left of the view at the three row coordinates.

### 2026-07-31 fourth run: entity-correspondence route settled

Installed-build runtime evidence. The Boolean normalization cleared the
qualification gates and the run exercised all three mapping routes for the
first time.

| Route | Result |
|---|---|
| A: `IView.GetCorrespondingEntity(modelEdge)` | **works** — returns a drawing entity, `error=0` |
| B: `IComponent2.GetCorrespondingEntity(modelEdge)` | returns `Nothing` on every attempt, `error=0` |
| C: `IView.GetVisibleEntities2(component, Edge)` | model edge never identity-matches (`-1`), as expected |

Counts: route A mapped 12 of 24 counterbore edges and 114 of 154 body
vertices; route B mapped 0 of either.

Route C nevertheless provides the decisive cross-check. Each entity returned
by route A **does** identity-match an entry in the visible-edge inventory
(`chosenVisibleIndex` 7, 19, 29-38, 44 via `ISldWorks.IsSame`), proving route
A returns genuine drawing-context entities and not a model-context alias.

Production R23 contract: use route A. Do not use `IComponent2` mediation for
part drawing views on this build. Retain `GetVisibleEntities2` as an
independent identity cross-check rather than as an acquisition path.

Note that mapping is per-edge, not per-face: for each counterbore only one of
the two owned circular edges maps. Production must attempt every owned edge
before failing a location, exactly as the probe does.

#### `ICurve.CircleParams` works

With the guard defect removed, `CircleParams` returned seven values on every
tested edge with a radius matching the owning cylinder exactly
(`0.005500000`, `0.003300000`, `0.023500000`). The R23-006 exclusion was an
artifact of the probe's `If Not <SOLIDWORKS Boolean>` guard, and the property
itself behaves correctly on this build. It remains non-load-bearing in the
probe; production may now treat it as available evidence.

#### `ISelectData.View` cannot be assigned in this VBA host

Both ordinate groups failed with runtime error 91, "Object variable or With
block variable not set", before any selection occurred
(`datumSelected=False`, `appended=0`). `ISelectionMgr.CreateSelectData`
returned a live object, so the failing statement is the documented get/set
assignment `Set selectData.View = swView`.

This reproduces the behaviour already recorded in
`Module2_DrawingPipeline.CreatePrimarySection`, whose comment states the same
assignment raises error 91 for section sketch segments and which works around
it by activating the source view first and proving ownership after `Select4`.

The MCP documents `ISelectData.View` as `View {get; set;}`, so this is an
installed-build deviation from the published contract, not a misuse. The
supported substitute is `ISelectionMgr.GetSelectedObjectsDrawingView2(Index,
Mark)`, which returns the drawing view for a selected object with `Index`
starting at 1.

The probe now attempts the binding inside a guarded helper, records
`viewBinding=Bound` or `UnboundAfterError:91`, and logs each selection's
owning view. Production R23 must not depend on the assignment succeeding.

### 2026-07-31 third run: the exact SOLIDWORKS Boolean contract

The `rejectGate` instrumentation isolated the mechanism precisely. For the
counterbore boundary edges the probe logged, from one call:

```text
isCircle=True ... rejectGate=IsCircleFalse ... completeCircle=False
```

`CStr(circleFlag)` rendered `True` while `If Not circleFlag` fired on that
same variable in the same procedure.

Established behaviour of installed SOLIDWORKS 2025 SP1.2 COM Booleans in this
VBA host:

| Construct | Safe? |
|---|---|
| `If value Then` | yes |
| `If value = False Then` | yes |
| `CStr(value)` | yes, renders `True` |
| `If Not value Then` | **no** — yields `-2`, which VBA treats as True |
| `CBool(rawVariant)` then `Not` | worked for `ISurface.IsCylinder` |
| `CBool(comCall)` then `Not` | **failed** for `ICurve.IsCircle` |

`CBool` is therefore not a dependable normalization: applied directly to the
method-call expression it can preserve the raw representation. The only
representation-independent form is an explicit numeric comparison,
`(CDbl(rawValue) <> 0#)`, which the probe now centralizes in
`NormalizeSwBoolean` and applies to `IsCircle`, `IsCylinder`, `ActivateView`,
both `Select4` calls, and `GetSaveFlag`.

Production R23 must use the same rule. This supersedes the narrower guidance
in the preceding subsection.

#### The `CircleParams` exclusion rests on this same defect

`docs/R23_IMPLEMENTATION_PLAN.md` R23-006 and `CURRENT_STATUS.md` record that
`ICurve.CircleParams` returned the probe's `SkippedNotCircle` sentinel while
`IsCircle` was true before and after, and conclude that production must not
depend on `CircleParams`.

`tools/r23-probes/Module_R23Phase0FeatureProbe.bas` line 611 shows the cause:

```vba
If Not isCircle Then
    ReadCircleState = "SkippedNotCircle"
    Exit Function
End If
```

That is the same `If Not <SOLIDWORKS Boolean>` construct. `CircleParams` was
never called on those edges, so no anomalous behaviour was ever observed. The
exclusion is not supported by evidence.

This does not establish that `CircleParams` is reliable — its behaviour on
this fixture remains untested. The corrected probe now reads it as
non-load-bearing evidence (`circleParams=` on each edge record) so the next
run settles the question rather than leaving a disproved exclusion in force.

### 2026-07-31 corrected-probe runtime: Boolean normalization is load-bearing

Installed-build runtime evidence from probe build `20260731.2`.

`ISurface.IsCylinder` is documented as returning `System.Boolean`, but the
installed VBA/COM path returns a raw `VARIANT_BOOL` whose True value does not
survive VBA's `Not` operator. The corrected drawing probe used
`If Not surface.IsCylinder Then` and rejected all 18 faces owned by
`CBORE for M6 Socket Head Cap Screw1` as `NotCylindrical`, while the accepted
feature probe reads the same faces as cylinders (radius `0.0055` and
`0.0033`) using `CBool(swSurface.IsCylinder)`.

The same run supplies its own control: `CBool(ICurve.IsCircle)` returned a
mixed True/False distribution across the 64 visible drawing edges in the same
document.

Rule for production R23: assign a SOLIDWORKS COM Boolean to a typed `Boolean`
or wrap it in `CBool` before any negation or compound logic. Assignment to a
declared `Boolean` (as `Select4` results already are) normalizes correctly;
an inline `Not <COM call>` does not. This restates the R20 repair and is now
backed by a direct installed-build reproduction.

`IView.GetXform` and the section payload were also cross-validated at
runtime. The inverse view-sketch-to-page conversion derived from
`GetXform` (origin `0.095932223, 0.165060000`, scale `1.0`) plus
`IView.Angle` (`-1.570796327`) reproduced page X values that match the
independently returned `GetSectionLineInfo2` arrow endpoints exactly
(`0.095932223` and `0.071467223`). The returned segment endpoints matched the
values passed to `CreateLine` at `deltaM=0`, proving the payload segment frame
is the source-view sketch frame.

`IModelDoc2.Parameter("D1@Sketch4")` returned the model dimension read-only
and reported `toleranceType=0`, `fitType=-1`, empty hole/shaft fit and empty
`GetToleranceFitValues`. H7 is not present in the P-0251 model source.

## Source inventory

Both snapshots contain the same 13 exported VBA components.

| Component | B lines | A lines | Main role |
|---|---:|---:|---|
| `BtnHandler.cls` | 23 | 23 | Main form button events |
| `Module1_Main.bas` | 192 | 200 | Public configuration types, globals, entry point |
| `Module2_DrawingPipeline.bas` | 286 | 399 | Drawing, views, sections, pipeline orchestration |
| `Module3_ModelAudit.bas` | 290 | 288 | Feature audit and hole-like feature records |
| `Module4_ModelItemImporter.bas` | 200 | 195 | Model annotations and dimension arrangement |
| `Module5_FallbackDimensionEngine.bas` | 253 | 539 | Circular candidates and ordinate creation |
| `Module6_QAEngine.bas` | 58 | 60 | QA summary |
| `Module7_TitleBlockEngine.bas` | 172 | 172 | Properties, notes, barcode, title block |
| `SectionBtnHandler.cls` | 24 | 24 | Section add/remove events |
| `SectionDlgBtnHandler.cls` | 24 | 24 | Section dialog events |
| `ThisLibrary.cls` | 8 | 8 | Macro project host class |
| `UserForm1.frm` | 523 | 632 | Main configuration form |
| `UserFormSection.frm` | 126 | 126 | Section entry form |

The complete static inventory, public-surface comparison, configuration-field matrix, and original call graphs are in `docs/ORDINATE_GAP_ANALYSIS.md`. The API findings below refine and, in one case, correct that static report.

## MCP coverage

All six requested MCP tools were used.

| MCP tool | Principal queries | Contribution |
|---|---|---|
| `solidworks_lookup_method` | `IDrawingDoc.InsertModelAnnotations4`, `IView.GetVisibleEntities2`, `IModelDocExtension.MultiSelect2`, `IModelDocExtension.AddOrdinateDimension`, `IModelDoc2.SetPickMode`, `IView.ModelToViewTransform`, `IView.GetOutline`, `IView.Position`, `IView.GetOrientationName`, `IView.Type`, `IModelDocExtension.AlignDimensions`, `IView.SetDisplayMode3`, `IFeature.GetTypeName2`, `ICustomPropertyManager.Get6`, `Add3`, `Set2` | Established method signatures, parameter meanings, return types, and required cleanup |
| `solidworks_search_api` | model annotation import, ordinate creation, visible drawing entities, circular edges, view classification, feature types, section coordinates | Located relevant interfaces and supporting API concepts |
| `solidworks_get_interface_members` | `ISelectData`, `IView` | Confirmed `ISelectData.View` and the available view-classification/geometry members |
| `solidworks_get_enum_values` | annotation, model-item source, ordinate, return-code, view entity/type, sketch segment, alignment, section, display, custom-property enums | Decoded the numeric constants used by both snapshots |
| `solidworks_find_related` | `InsertModelAnnotations4`, `AddOrdinateDimension`, `AlignDimensions` | Identified adjacent APIs without assuming they should replace the current implementation |
| `solidworks_get_examples` | annotation import, ordinate creation, auto-arrangement, visible drawing entities | Identified official example topics for local comparison |

## Executive API findings

1. **The active import mask is wrong against the MCP enum.** In A, `swInsertDimensionsMarkedForDrawing = 1` and `swInsertHoleWizardCallouts = 64` (`A/Module4_ModelItemImporter.bas:4-7`). The MCP maps `1` to `swInsertCThreads` and `64` to `swInsertNotes`. The correct SW2026 values for the names the active code intends are `32768` for marked-for-drawing dimensions and `1048576` for hole callouts. This directly explains why the active version can import the wrong annotation categories or no intended model dimensions.

2. **The baseline import constants match the MCP corpus.** The values declared in B for dimensions, GTols, marked dimensions, Hole Wizard profile/location dimensions, and hole callouts (`B/Module4_ModelItemImporter.bas:6-12`) match the corresponding SW2026 enum members. This materially strengthens the evidence for why the baseline import works.

3. **Both importers explicitly allow duplicate dimensions.** The fourth argument of `InsertModelAnnotations4` is `DuplicateDims`; the MCP states `True` eliminates duplicate dimensions and `False` allows them. Both snapshots pass `False` (`B/Module4_ModelItemImporter.bas:36-44`, `B:90-98`; `A/Module4_ModelItemImporter.bas:47-55`). This conflicts with the requirement to avoid unnecessary duplication.

4. **The active anchor-view selection is API-aligned; the baseline sheet-level first call is not documented by the MCP contract.** `InsertModelAnnotations4` operates from the selected drawing view, with `AllViews=True` extending insertion to all views. A activates and selects an anchor view before the call (`A/Module4_ModelItemImporter.bas:27-55`). B clears to the sheet and makes its first call without a selected drawing view (`B/Module4_ModelItemImporter.bas:32-44`), then falls back to documented selected-view calls (`B:64-114`). A repair should preserve A's selected anchor, correct the mask, and add B's per-view retry pattern; it should not blindly copy the undocumented sheet-level first call.

5. **Both ordinate engines omit required `SetPickMode` cleanup.** The MCP remarks for `AddOrdinateDimension` say to call `IModelDoc2.SetPickMode` when the ordinate group is finished. Both implementations only call `ClearSelection2` (`B/Module5_FallbackDimensionEngine.bas:238-253`; `A/Module5_FallbackDimensionEngine.bas:469-520`). The first group can therefore leave SOLIDWORKS in ordinate continuation mode when the second group starts.

6. **Datum-first selection is not proven by the current `MultiSelect2` batch.** `AddOrdinateDimension` requires the base entity to be selected as the datum before the other entities. Both arrays place the chosen datum first, but both submit the entire array in one `MultiSelect2` call (`B/Module5_FallbackDimensionEngine.bas:202-239`; `A/Module5_FallbackDimensionEngine.bas:422-484`). The MCP documents the returned selection count, but does not guarantee that array order becomes selection order. The safe contract is an explicit datum selection followed by appended selections.

7. **`GetVisibleEntities2(Nothing, edge)` remains unresolved.** The MCP signature requires a `Component2` and an entity type. Both snapshots pass `Nothing` (`B/Module5_FallbackDimensionEngine.bas:24-29`, `B:94-100`; `A/Module5_FallbackDimensionEngine.bas:103-113`). The MCP does not document `Nothing` as a valid part-view shortcut. This is a local SW2025 test/type-library question, not proof that the current calls always fail.

8. **A circular curve is not sufficient evidence of a hole or even a complete circular edge.** `ICurve.IsCircle` confirms circular curve geometry. The MCP remarks point to edge curve parameters to distinguish a complete circle from an arc. Neither ordinate engine performs that completeness check (`B/Module5_FallbackDimensionEngine.bas:43-55`; `A/Module5_FallbackDimensionEngine.bas:168-203`), and neither connects the edge to proven feature data. Therefore bosses, fillet arcs, counterbore rings, and other circular geometry remain possible candidates. This report does not classify any circular edge as a hole without model/API evidence.

9. **The earlier concern about `ISketchArc.IsCircle = 1` is resolved.** The MCP states `ISketchArc.IsCircle()` returns `Int32`, with `1` for a complete circle and `0` for a partial circle. A's explicit comparison is valid (`A/Module3_ModelAudit.bas:189-206`); B's truth-value test also works (`B/Module3_ModelAudit.bas:190-207`). This supersedes the contrary risk statement in `docs/ORDINATE_GAP_ANALYSIS.md`.

10. **Both drawing pipelines use mismatched display-mode constants.** Both define HLR as `3` and shaded-with-edges as `6` (`B/Module2_DrawingPipeline.bas:10-13`; `A/Module2_DrawingPipeline.bas:12-16`). The MCP SW2026 enum maps HLR (`swHIDDEN`) to `2`, `swSHADED` to `3`, `swFACETED_HIDDEN` to `6`, and `swSHADED_EDGES` to `7`. Both pass those constants to `SetDisplayMode3` (`B:201-207`; `A:255-267`). Verify SW2025 before correction.

11. **Name-based isometric filtering is fragile.** A and B search the drawing-view instance name for `ISO`/`ISOMETRIC` (`B/Module2_DrawingPipeline.bas:283-285`; `A/Module5_FallbackDimensionEngine.bas:534-538`). The MCP exposes `IView.Type` and `IView.GetOrientationName`; `GetOrientationName` returns names such as `*Front` and `*Isometric`, while derived views can return an empty orientation name. Eligibility should use API properties, not only a user-visible instance name.

12. **The active section-coordinate correction is API-consistent.** `IView.GetOutline` returns a view bounding box in drawing-page coordinates, and `CreateSectionViewAt5` uses drawing-sheet placement coordinates. A builds its section line from the front-view outline (`A/Module2_DrawingPipeline.bas:269-379`). B uses an approximate model `GetPartBox` directly for drawing-sketch geometry (`B/Module2_DrawingPipeline.bas:209-267`), mixing coordinate systems. This does not imply multi-section support.

## API contract validation by method

| API member | MCP contract relevant here | Code use | Assessment |
|---|---|---|---|
| `IDrawingDoc.InsertModelAnnotations4` | Source enum, annotation mask, `AllViews`, `DuplicateDims`, hidden-feature and placement flags; returns annotation array | B `Module4:16-114`; A `Module4:11-69` | A has documented selected-view setup but wrong mask and no fallback. B has correct mask and selected-view fallback, but its initial sheet-level call is undocumented. Both allow duplicates. |
| `IModelDocExtension.AddOrdinateDimension` | Selected base entity is datum; other selected entities join the group; returns `swCreateOrdDimError_e`; finish with `SetPickMode` | B `Module5:190-253`; A `Module5:403-521` | Direction and success constants match. Selection ordering is unproven; cleanup is incomplete. |
| `IModelDoc2.SetPickMode` | Returns the document to default selection mode | No call in either snapshot | Required after each completed/failed ordinate group. |
| `IModelDocExtension.MultiSelect2` | Returns the number actually selected; ignored objects can reduce the count; `AppendFlag=False` replaces the selection | B `Module5:238-239`; A `Module5:469-484` | Both validate count. Neither proves datum-first selection order. |
| `ISelectData.View` | Gets/sets the drawing view containing the selected object | B `Module5:20-22`; A `Module5:75-90` | Correct view-scoping mechanism; A also activates the view first. |
| `IView.GetVisibleEntities2` | Requires a view component and `swViewEntityType_e`; returns entities not completely obscured | B `Module5:24-29`, `94-100`; A `Module5:103-113` | Edge type `1` is correct. `Nothing` component is not documented by the MCP and needs SW2025 proof. |
| `ICurve.IsCircle` / `CircleParams` | Circular-curve test; parameters are center XYZ, axis XYZ, radius, in meters | B `Module5:43-67`; A `Module5:168-203` | Center transform and A's `radius * 2` are consistent. Completeness and feature identity are not established. |
| `ISketchSegment.GetType` | Returns `swSketchSegments_e`; `swSketchARC=1` | B `Module3:185-210`; A `Module3:184-209` | Both numeric checks are confirmed against the MCP corpus. |
| `ISketchArc.IsCircle` | Returns `1` for complete circle, `0` for partial | B `Module3:190-207`; A `Module3:189-206` | Both predicates are acceptable. |
| `IView.ModelToViewTransform` | Maps model coordinates to drawing-view coordinates | B `Module5:31-67`; A `Module5:118-202` | Correct concept for projecting circle centers. The resulting origin relative to `IView.Position` still needs a practical SW2025 coordinate-frame test. |
| `IView.GetOutline` | `[xmin, ymin, xmax, ymax]` in meters on the drawing page | B `Module5:24-25`; A `Module5:95-100`; A `Module2:293-379` | Appropriate for placement bounds and active section sketch geometry. |
| `IView.Position` | Geometric center of the view relative to the sheet origin | B `Module5:161-188`; A `Module5:349-398` | Subtracting it from sheet-space outline targets is plausible, but does not establish a true datum entity. |
| `IView.Type` | Returns `swDrawingViewTypes_e` | Not used for ordinate eligibility | Better first filter for sheet, section, detail, projected, auxiliary, standard, and named views. |
| `IView.GetOrientationName` | Returns predefined name such as `*Isometric`; empty for several derived view types | Not used | Better evidence than instance-name substring matching, used together with `Type`. |
| `IModelDocExtension.AlignDimensions` | Alignment enum plus spacing; returns Boolean | B `Module4:139-165`; A `Module4:109-150` | Auto-arrange value `0` is correct. Both ignore the Boolean result; spacing units/behavior need SW2025 verification. |
| `IView.GetDisplayDimensions` | Returns an array of display dimensions | B `Module4:145-161`; A `Module4:120-144` | A's `IsArray` guard is safer. Both arrange all dimensions without checking final sheet bounds. |
| `IView.SetDisplayMode3` | Accepts `swDisplayMode_e`; method is obsolete in SW2026 in favor of `SetDisplayMode4`; returns Boolean | B `Module2:201-207`; A `Module2:255-267` | Both use mismatched enum values and ignore the result. Confirm availability and values in SW2025. |
| `IFeature.GetTypeName2` | Identifies feature type; examples include `HoleWzd`, `AdvHoleWzd`, `SketchHole`, and `Cut`; feature data is obtained through `GetDefinition` where applicable | B `Module3:80-134`; A `Module3:79-133` | Current audit recognizes only `HoleWzd` exactly and any string containing `CUT`. It misses documented hole types and over-includes arbitrary cuts. |
| `ICustomPropertyManager.Get6` | Returns a result code; `UseCached=False` requests current data | B/A `Module7:45-67` | `False` is appropriate; return code is ignored. |
| `ICustomPropertyManager.Add3` / `Set2` | `Add3` returns add result; option `0` means only-if-new; `Set2` updates value | B/A `Module7:69-74` | `swCustomInfoText=30` is confirmed in the MCP corpus. Both calls ignore return codes. |

## Enum values used by the macros

### Model annotation import

MCP enum: `swInsertAnnotation_e`.

| Member | MCP value | B declaration/use | A declaration/use |
|---|---:|---|---|
| `swInsertCThreads` | 1 | Not used | A calls this value `swInsertDimensionsMarkedForDrawing` (`A/Module4:6`) |
| `swInsertDimensions` | 8 | Correct (`B/Module4:6`) | Not declared |
| `swInsertGTols` | 32 | Correct (`B/Module4:7`) | Not declared |
| `swInsertNotes` | 64 | Not used | A calls this value `swInsertHoleWizardCallouts` (`A/Module4:7`) |
| `swInsertDimensionsMarkedForDrawing` | 32768 | Correct (`B/Module4:8`) | Intended but not used |
| `swInsertHoleWizardProfileDimensions` | 65536 | Correct (`B/Module4:9`) | Not declared |
| `swInsertHoleWizardLocationDimensions` | 131072 | Correct (`B/Module4:10`) | Not declared |
| `swInsertDimensionsNotMarkedForDrawing` | 524288 | Declared, not included (`B/Module4:11`, `175-182`) | Not declared |
| `swInsertholeCallout` | 1048576 | Correct (`B/Module4:12`) | Intended but not used |

MCP enum `swImportModelItemsSource_e` confirms `swImportModelItemsFromEntireModel=0`, used by both snapshots (`B/Module4:4`; `A/Module4:4`).

### Ordinate creation

MCP enum `swAddOrdinateDims_e`:

| Member | Value | Source use |
|---|---:|---|
| `swAddOrdinateDims_Ordinate` | 1 | Not used directly |
| `swAddOrdinateDims_Vertical` | 2 | B/A vertical constant is correct |
| `swAddOrdinateDims_Horizontal` | 3 | B/A horizontal constant is correct |
| `swAddOrdinateDims_Angular` | 4 | Not used |

MCP enum `swCreateOrdDimError_e`:

| Return | Value | Meaning |
|---|---:|---|
| `swCreateOrdDimError_Undefined` | -1 | Undefined |
| `swCreateOrdDimError_Success` | 0 | Success |
| `swCreateOrdDimError_GenFailure` | 1 | General failure |
| `swCreateOrdDimError_GenNoInternalDims` | 2 | No internal dimensions |
| `swCreateOrdDimError_GenBadSel` | 3 | Bad selection |
| `swCreateOrdDimError_GenNeedModelLoaded` | 4 | Model must be loaded |
| `swCreateOrdDimError_GenSamePartOnly` | 5 | Selections must be from same part |
| `swCreateOrdDimError_GenExtraSelection` | 6 | Extra selection |
| `swCreateOrdDimError_OrdFailure` | 7 | Ordinate failure |
| `swCreateOrdDimError_OrdDupInGroup` | 8 | Duplicate in ordinate group |
| `swCreateOrdDimError_OrdBadDir` | 9 | Invalid ordinate direction |

Both snapshots log only the number (`B/Module5:248-250`; `A/Module5:505-511`). Decoding it in QA would make failures actionable.

### Views and display

MCP enum `swViewEntityType_e` confirms `swViewEntityType_Edge=1`, matching A's named constant and B's literal.

MCP enum `swDrawingViewTypes_e` includes:

| Type | Value |
|---|---:|
| Sheet | 1 |
| Section | 2 |
| Detail | 3 |
| Projected | 4 |
| Auxiliary | 5 |
| Standard | 6 |
| Named | 7 |
| Relative | 8 |
| Detached | 9 |
| Alternate | 10 |

These values are reported for SW2026 only. The repair should use the referenced enum names in VBA where the local library supports them, not copied numeric literals.

MCP enum `swDisplayMode_e`:

| Member | Value | Current source label |
|---|---:|---|
| `swWIREFRAME` | 0 | None |
| `swHIDDEN_GREYED` (HLV) | 1 | Correct in B/A |
| `swHIDDEN` (HLR) | 2 | B/A incorrectly use 3 |
| `swSHADED` | 3 | B/A label 3 as HLR |
| `swFACETED_WIREFRAME` | 4 | None |
| `swFACETED_HIDDEN_GREYED` | 5 | None |
| `swFACETED_HIDDEN` | 6 | B/A label 6 as shaded-with-edges |
| `swSHADED_EDGES` | 7 | Intended by B/A but not used |

### Other confirmed constants

- `swAlignDimensionType_AutoArrange=0`, matching B/A `Module4`.
- `swCreateSectionView_NotAligned=1`, matching B/A `Module2`.
- `swSketchARC=1`, matching B/A `Module3`.
- `swCustomInfoText=30`, matching B/A `Module7`.
- `swCustomPropertyOnlyIfNew=0`, which explains the option passed to `Add3` in B/A `Module7:72`.

All remain subject to SW2025 type-library confirmation before edits.

## API-annotated call graphs

### Model-dimension import

```text
B/Module2.RunDrawingPipeline [B/Module2:24-78]
  -> B/Module4.ImportModelItemsAcrossDrawing [B/Module4:16-62]
     -> GetModelItemMask [B:175-182]
        -> API-confirmed annotation values
     -> InsertModelAnnotations4(AllViews=True, DuplicateDims=False) [B:32-46]
        -> no selected view: not documented by MCP
     -> if zero: ImportModelItemsPerView [B:48-50, 64-114]
        -> activate and select each drawing view
        -> InsertModelAnnotations4(AllViews=False, DuplicateDims=False)

A/Module2.RunDrawingPipeline [A/Module2:27-107]
  -> A/Module4.ImportModelItemsAcrossDrawing [A/Module4:11-69]
     -> activate/select anchor view [A:27-43]
        -> documented selected-view context
     -> GetModelItemMask [A:170-177]
        -> values resolve to cosmetic threads and notes, not intended categories
     -> InsertModelAnnotations4(AllViews=True, DuplicateDims=False) [A:47-55]
     -> no retry when returned count is zero
```

### Ordinate candidate collection

```text
Pipeline -> Module5.CreateHoleOrdinateDims
  -> activate view and set ISelectData.View
  -> IView.GetVisibleEntities2(Nothing, Edge)
     -> Edge value confirmed; Nothing component unresolved
  -> IEdge.GetCurve
  -> ICurve.IsCircle
     -> circular geometry only; not complete-circle or hole proof
  -> ICurve.CircleParams
  -> center point * IView.ModelToViewTransform
  -> projected-center deduplication
     -> A retains the smallest concentric diameter [A/Module5:206-215]
```

`Module3.GetAllHoleLikeFeatures` is called before drawing creation (`B/Module2:34-35`; `A/Module2:35-36`), but neither candidate collector consumes its records. The audit itself also needs stricter API-backed feature typing (`B/Module3:98-134`; `A/Module3:97-133`).

### Ordinate dimension creation

```text
CreateHoleOrdinateDims
  -> ResolveDatumIndex
  -> CreateOneOrdinateChain(horizontal)
     -> batch MultiSelect2(datum-first array, Append=False)
     -> AddOrdinateDimension(horizontal)
     -> ClearSelection2 only
  -> CreateOneOrdinateChain(vertical)
     -> batch MultiSelect2(datum-first array, Append=False)
     -> AddOrdinateDimension(vertical)
     -> ClearSelection2 only
  -> missing IModelDoc2.SetPickMode after each group
```

Exact code ranges: B `Module5:161-253`; A `Module5:349-521`.

### Datum-origin resolution

```text
UserForm1.cmbDatum
  -> GlobalConfig.DatumOrigin
  -> Module5.ResolveDatumIndex
     -> IView.GetOutline (sheet coordinates)
     -> IView.Position (sheet-relative geometric center)
     -> derive target relative to view position
     -> choose nearest circular candidate
     -> use that edge as intended ordinate datum
```

Exact code ranges: B `UserForm1:119-130`, `451-466`, `Module5:161-188`; A `UserForm1:114-124`, `530-546`, `Module5:349-398`.

The labels Bottom-Left, Center, and Top-Left do not currently select a view corner/origin datum. They select the circular candidate nearest the requested target. The MCP does not establish that this is the intended datum convention.

### Dimension arrangement

```text
GlobalConfig.AutoArrange
  -> Module4.AutoArrangeAllDrawingDimensions
  -> for each real view: IView.GetDisplayDimensions
  -> select every returned dimension annotation
  -> IModelDocExtension.AlignDimensions(AutoArrange=0, SpaceValue=0.06)
  -> ignore Boolean return
  -> ClearSelection2
```

Exact code ranges: B `Module2:62-64`, `Module4:128-165`; A `Module2:75-79`, `Module4:88-150`.

## Module-by-module API assessment

| Component | API-backed assessment |
|---|---|
| `Module1_Main.bas` | No API enum defect found in its public types. Active defaults and form submission support fixed hybrid behavior (`B:4-42`, `117-121`; `A:4-42`, `122-126`). |
| `Module2_DrawingPipeline.bas` | Both snapshots use incorrect display-mode values against the MCP corpus. A's section geometry uses correct drawing-page coordinates; B mixes approximate model-box data with drawing sketch coordinates. |
| `Module3_ModelAudit.bas` | `SketchSegment.GetType=1` and A's `SketchArc.IsCircle=1` are confirmed. Feature classification is incomplete: only `HoleWzd` is recognized explicitly, while documented `AdvHoleWzd` and `SketchHole` are missed; any type containing `CUT` is treated as a candidate. |
| `Module4_ModelItemImporter.bas` | B's annotation constants match the MCP corpus; A's two mask constants do not. Both pass `DuplicateDims=False`, ignore `AlignDimensions` return, and use fixed spacing without a verified unit/bounds policy. |
| `Module5_FallbackDimensionEngine.bas` | Direction/entity/success values are confirmed. Both omit `SetPickMode`, batch-select an intended datum without an ordering guarantee, pass undocumented `Nothing` to `GetVisibleEntities2`, and accept circular geometry without complete-circle/feature evidence. A adds valuable activation, validation, center deduplication, and return-code logging. |
| `Module6_QAEngine.bas` | No critical API contract issue found. It does not receive or decode ordinate result codes, so imported dimensions can hide ordinate failure in the summary (`B:5-56`; `A:5-59`). |
| `Module7_TitleBlockEngine.bas` | `swCustomInfoText=30` and `Add3` option `0` are confirmed. `Get6`, `Add3`, and `Set2` return codes are ignored (`B/A:45-74`). |
| `UserForm1.frm` | A correctly forces both dimension stages on (`A:530-562`). B exposes mutually exclusive modes (`B:109-131`, `451-479`). The active callout checkbox reaches the mask builder, but the mask bit currently represents notes, not callouts. |
| `UserFormSection.frm` and three handler classes | No relevant API contract difference. No event interface mismatch was demonstrated. |
| `ThisLibrary.cls` | No executable API use. |

## Why the baseline imports model dimensions correctly

The baseline's observed success now has direct API support in addition to repository history:

1. Its mask constants match the MCP's annotation enum (`B/Module4:6-12`, `175-182`). In particular, `32768` is marked-for-drawing dimensions, `65536` and `131072` are Hole Wizard profile/location dimensions, and `1048576` is the hole-callout bit.
2. Views are created and rebuilt before import (`B/Module2:32-55`).
3. If the first insertion returns zero, B activates and selects each real drawing view and retries with `AllViews=False` (`B/Module4:48-50`, `64-114`). That fallback conforms more closely to the MCP's selected-view contract.
4. It counts the returned annotation array and clears selection/view state on normal and error paths (`B/Module4:46-62`, `100-114`).

The baseline is not a perfect implementation: its initial sheet-level call lacks the selected drawing view described by the MCP, `DuplicateDims=False` allows duplicates, the callout checkbox does not control its mask, and its UI is not the required fixed hybrid workflow. The baseline should remain protected and be used as behavioral evidence, not copied wholesale.

## Why the active ordinate version may fail

Ranked by immediacy:

1. **Wrong annotation bits:** A asks for cosmetic threads and optionally notes while naming them marked dimensions and Hole Wizard callouts (`A/Module4:4-7`, `170-177`).
2. **No import fallback:** a zero-result all-view call ends the import (`A/Module4:47-69`).
3. **Ordinate mode is not exited:** both chains omit the API-required `SetPickMode` (`A/Module5:469-520`). This can corrupt or extend the next group.
4. **Datum selection order is not guaranteed:** batching all objects through `MultiSelect2` does not prove the intended base edge became the datum (`A/Module5:422-503`).
5. **Visible-edge retrieval contract is unresolved:** `Nothing` is supplied where the MCP signature requires a component (`A/Module5:103-113`).
6. **Candidates lack proof:** `ICurve.IsCircle` accepts circular geometry, but the code neither rejects partial circular edges nor maps candidates to proven hole-producing features (`A/Module5:168-215`).
7. **Datum semantics are heuristic:** the requested origin becomes the nearest circular candidate, not an independent datum (`A/Module5:349-398`).
8. **View eligibility is name-based:** renamed isometric and unsupported derived views can be mishandled (`A/Module5:260-289`, `534-538`).
9. **Display mode values are mislabeled:** views may be configured in unintended modes (`A/Module2:12-16`, `255-267`).
10. **Diagnostics are not propagated:** result codes go only to the Immediate Window; QA can pass because model dimensions exist even when ordinate creation fails (`A/Module5:505-511`; `A/Module6:5-59`).

## Public surface and configuration implications

The MCP evidence does not change the earlier public-surface finding:

- `SectionConfig` and `DrawingConfig` are identical in B and A (`B/Module1:4-42`; `A/Module1:4-42`).
- Neither snapshot declares a public enum or public constant.
- A adds public `AddOrdinateDimensionsToAllViews` (`A/Module5:260-289`).
- B alone exposes unused public `InsertHoleCalloutsForView` (`B/Module5:94-149`). Its `AddHoleCallout2` path uses raw model circle coordinates without the model-to-view transformation and can display an interactive prompt, so it is not a safe automation fallback.
- B exposes `CountAllViewDimensions`; A makes the equivalent helper private (`B/Module6:36-45`; `A/Module6:45-59`).

Configuration conclusions remain:

- A's form sets `UseModelDimensions=True` and `UseOrdinateDims=True`, matching the fixed hybrid requirement (`A/UserForm1:530-562`).
- `ImportHoleCallouts` is consumed in A but currently controls the wrong enum bit (`A/UserForm1:545`; `A/Module4:170-177`). It is set but unused in B (`B/UserForm1:465`; `B/Module4:175-182`).
- `RunHybridStrategy`, `CustomScaleText`, `ShowLayoutPreview`, and compatibility `SectionCount` are set/declared but unused downstream in both snapshots.
- `GenerateQAReport` is consumed only in A; B runs QA unconditionally.
- `CreateFront` is stored but the front view is created unconditionally in both pipelines.

The complete field-by-field declaration/set/consume/unused matrix remains in `docs/ORDINATE_GAP_ANALYSIS.md` section 5.

## Ranked repair plan

No repair is implemented in this task. Multi-section generation is not proposed.

| Rank | Smallest safe correction | Dependencies/callers affected | Required focused test | SW2025 fact to verify |
|---:|---|---|---|---|
| 1 | Replace A's import mask declarations with locally verified `swInsertAnnotation_e` members; make the callout bit conditional; pass `DuplicateDims=True`; keep the selected anchor and add B-style selected per-view retry when the returned count is zero. | `A/Module4:4-69`, `170-194`; caller remains `A/Module2:58-66`. | Saved part with marked dimensions; front/right views; callouts off/on; prove nonzero imports and no duplicate model dimensions. Force a zero first result and prove retry. | Exact SW2025 values/names for marked dimensions, Hole Wizard profile/location dimensions, and hole callouts; `InsertModelAnnotations4` signature. |
| 2 | End every ordinate group with `swModel.SetPickMode`, including failure and early-exit paths, before starting the other direction. | `A/Module5:403-521`; invoked from `CreateHoleOrdinateDims` at `236-242`. Apply the same rule if B is ever reused. | Two or more proven locations; create horizontal then vertical groups; confirm two independent chains and normal selection mode afterward. | `IModelDoc2.SetPickMode` availability/behavior in SW2025. |
| 3 | Select the intended datum edge explicitly first, then append each remaining selected edge; retain exact selection-count validation and decode return codes. | `A/Module5:422-511`; datum source `349-398`. | Bottom-Left, Center, Top-Left runs; verify which entity is zero and validate coordinates. Exercise bad selection and duplicate-in-group errors. | Datum entity requirements, selection-order behavior, and `swCreateOrdDimError_e` in SW2025. |
| 4 | Prove the visible-entity acquisition contract before changing candidate logic: supply the documented component if required, reject partial circular edges, and require model/feature evidence before treating a center as a qualifying hole location. Preserve projected-center deduplication. | `A/Module5:103-221`; likely `A/Module3:17-286`; coordination at `A/Module2:35-36`, `68-73`. Update all changed public callers together. | Plate with genuine holes plus a circular boss, fillet arc, and counterbore. Only evidence-backed locations may be dimensioned; do not label geometry as a hole from shape alone. | Valid component for part drawing views; `IEdge.GetCurveParams3` completeness data; drawing edge-to-model feature mapping; feature-data interfaces for `HoleWzd`, `AdvHoleWzd`, and `SketchHole`. |
| 5 | Replace instance-name isometric checks with an eligibility helper using `IView.Type` and `GetOrientationName`; explicitly allow only supported orthographic source views. | `A/Module5:260-289`, `534-538`; possibly pipeline view policy. | Renamed isometric, standard front/top/right, projected, section, and detail views. Confirm only approved orthographic views receive ordinates. | SW2025 enum members and orientation-name behavior for projected/derived views. |
| 6 | Correct display-mode constants from the local enum and check the Boolean result of the supported display method. | `A/Module2:12-16`, `255-267`; equivalent defect exists in protected B but B must not be modified. | HLR on/off and isometric shaded-with-edges; visually confirm each mode and record API result. | SW2025 `swDisplayMode_e`; whether `SetDisplayMode3` or `SetDisplayMode4` is the supported method. |
| 7 | Make arrangement and QA report API outcomes: check `AlignDimensions` Boolean, bound placement against sheet/view extents, decode ordinate result codes, and report candidate/group counts per view. | `A/Module4:88-150`; `A/Module5:486-521`; `A/Module6:5-59`; `A/Module2:75-99`. | Dense hole pattern on small and large templates; verify no dimensions leave the sheet/title-block area and QA distinguishes import success from ordinate failure. | `AlignDimensions` spacing units; annotation extents; whether view outline includes annotations. |

## Required SW2025 local type-library verification

Before any patch is accepted, capture Object Browser/type-library evidence for:

1. `IDrawingDoc.InsertModelAnnotations4` and `swInsertAnnotation_e`.
2. `IModelDocExtension.AddOrdinateDimension`, `swAddOrdinateDims_e`, and `swCreateOrdDimError_e`.
3. `IModelDoc2.SetPickMode`.
4. `IModelDocExtension.MultiSelect2` and datum selection order.
5. `IView.GetVisibleEntities2`, `swViewEntityType_e`, and the correct component for a saved part drawing view.
6. `ICurve.IsCircle`, `ICurve.CircleParams`, and `IEdge.GetCurveParams3` or the SW2025 equivalent.
7. Drawing edge-to-model entity/feature ownership and the feature data interfaces for Hole Wizard, Advanced Hole, and simple hole features.
8. `IView.Type`, `swDrawingViewTypes_e`, and `IView.GetOrientationName`.
9. `IView.ModelToViewTransform`, `IView.GetOutline`, and `IView.Position` coordinate frames.
10. `IModelDocExtension.AlignDimensions`, its Boolean return, and spacing units.
11. `swDisplayMode_e` plus the supported `SetDisplayMode3`/`SetDisplayMode4` method.

## MCP examples and related API references

The MCP example catalog returned these official example topics for comparison during implementation/testing:

- `Insert Model Annotations (VBA)`
- `Create Ordinate Dimensions (VBA)`
- `Auto-arrange Dimensions (VBA)`
- `Get Visible Components and Entities in Drawing View (VBA)`

Related-member searches returned:

- For `InsertModelAnnotations4`: `InsertModelDimensions`, `InsertModelInPredefinedView`.
- For `AddOrdinateDimension`: `AlignOrdinate`, `CreateOrdinateDim4`, `EditOrdinate`, `InsertHorizontalOrdinate`, `InsertVerticalOrdinate`, `ReattachOrdinate`, `JogDimension`.
- For `AlignDimensions`: `AlignOrdinate`, `AlignParallelDimensions`, `BreakDimensionAlignment`.

These are reference points only. The current evidence supports narrow corrections to the active workflow, not a speculative API rewrite.

## 2026-07-28 CodeStack drawing-corpus addendum

The complete CodeStack
[`solidworks-api/document/drawing`](https://www.codestack.net/solidworks-api/document/drawing/)
corpus was reviewed at source commit
`0cde3849a184cdbbface61238ef431a6ebb9d530`.

- 33 of 33 article pages and 35 of 35 adjacent VBA/C# examples were read.
- All 33 derived public article URLs returned HTTP 200.
- CodeStack is retained as practical secondary guidance; official SOLIDWORKS
  2025 Help, the installed type library, and local runtime evidence remain
  authoritative.
- The corpus does not contain a complete manufacturing-drawing pipeline and
  does not cover the project's full section/detail, model-item import,
  ordinate, title-block, layout, and fail-closed QA requirements.

The complete ledger and the project-specific synthesis are:

- [`CODESTACK_DRAWING_API_COVERAGE.md`](CODESTACK_DRAWING_API_COVERAGE.md)
- [`3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md`](3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md)

The most important confirmed alignment with current project evidence is:

1. real drawing views must be distinguished from the sheet pseudo-view;
2. `GetVisibleEntities2` should receive the actual `Component2` returned by
   `GetVisibleComponents`;
3. model-context entities and drawing-context visible entities use different
   correspondence/selection routes;
4. model annotation import is selected-view state;
5. ordinate creation is a datum-first stateful transaction ending in
   `SetPickMode`;
6. model, view-sketch, and sheet coordinate systems must remain explicit; and
7. CodeStack examples require production hardening rather than direct copying.

This addendum does not revise historical snapshot assessments above. Where an
older table or copied numeric value conflicts with newer installed-type-library
evidence, the newer evidence and current source take precedence.

## 2026-07-28 R20 final-contract addendum

The final R20 repair also queried the project `solidworks-api` MCP for the
following load-bearing contracts:

- `IModelDoc2.InsertNote`: selections made before insertion supply leader
  attachment points; no selection produces a free-standing leaderless note.
  SOLIDWORKS symbols must use `<LibraryName-SymbolName>` syntax, so diameter is
  emitted as `<MOD-DIAM>`.
- `IAnnotation.GetAttachedEntities3`: an unassociated note returns an empty
  array. R20 therefore requires a nonempty attachment readback for each
  controlled manufacturing callout.
- `IAnnotation.GetLeaderCount`: R20 requires at least one visible leader on
  each controlled callout.
- `IView.GetSectionLineInfo2`: the flattened result contains segment endpoints,
  arrow geometry, label positions, and text height. R20 parses those fields and
  rejects J-J geometry that enters the measured part-identification extent.
- `IEntity.Select4`: the callout uses the drawing-document `IEntity` obtained
  from the ownership-proven corresponding drawing edge, with
  `ISelectData.View` bound to its drawing view.
- `IView.ModelToViewTransform` returns drawing-view-space coordinates.
  CodeStack's dimension examples use the transformed values directly for
  drawing placement; R20 now adds a runtime invariant requiring each such point
  to fall within the current `IView.GetOutline` before treating it as page
  evidence.

The installed SOLIDWORKS 2025 interop was also checked for the arrangement
signatures used by the fail-closed fallback:
`IAnnotation.Select3`, `IAnnotation.GetPosition`,
`IAnnotation.SetPosition2`, `IModelDocExtension.AlignDimensions`, and
`IView.GetOutline`. Radial and diametric dimensions are not moved by the
fallback because `SetPosition2` does not support them; their existing origins
must already pass the content-border check.

## 2026-07-28 R20 GetSectionLineCount2 correction

The user's full-project compile highlighted `Argument not optional` at the
parameterless Module6 call to `IView.GetSectionLineCount2`.

- MCP contract:
  `System.Int32 GetSectionLineCount2(out System.Int32 Size)`. `Size` is the
  returned section-line data size and includes the layer-ID double for each
  section line.
- Installed SOLIDWORKS 2025 interop:
  `Int32 GetSectionLineCount2(Int32 ByRef)` with `IsOut=True` and
  `IsOptional=False`.
- Adjacent contract:
  `IView.GetSectionLineInfo2()` is parameterless and returns the flattened
  double array.
- Source correction:
  `Module6_QAEngine.CheckSectionLineClearance` now passes a `Long` by reference
  and rejects an invalid size or a returned array whose item count differs from
  that size.

This also invalidated the earlier interpretation of the deployment
`COMPILE_PROBE`: it executes only the bootstrap procedure and does not perform
VBA Editor **Compile Project**. Manual full-project compilation remains the E5
gate.

## 2026-07-29 R22 review-resolution addendum

R22 re-queried the project `solidworks-api` MCP and reflected the installed
SOLIDWORKS 2025 SP1.2 interop assemblies (`33.1.2.4`) before replacing the
three low-confidence r21 implementations.

### Circular trimming edges

- MCP `ICurve.IsCircle`: returns True only when the curve is a circle; False
  means another curve type. Its Remarks direct callers to edge curve
  parameters to distinguish a complete circle from an arc.
- MCP `ICurve.CircleParams`: the seven-value array is
  `[center xyz, axis xyz, radius]`.
- MCP `ISurface.CylinderParams`: the seven-value array describes the
  cylindrical surface as `[axis origin xyz, axis xyz, radius]`; it does not
  state that every closed trimming edge is circular.
- Installed interop confirms `Boolean IsCircle()`, `Object CircleParams`, and
  `Object CylinderParams`.
- Source decision: `TryReadClosedCircularEdge` now requires both complete-edge
  topology/parameter proof and `ICurve.IsCircle=True`. A closed trim on a
  cylindrical face fails as `ClosedCircleCurveNotCircular` when that predicate
  is False. `CylinderParams` remains valid for independently checking the
  already-proved circle against its owning internal cylindrical face, but no
  longer manufactures circle data.

The retained
`20260728_083008_P-0251-14A-001/R19_CURVE_CONTRACT_PROBE.json` records
`is_circle=true`, complete `0..2*pi` spans, and seven-value circle parameters
for every relevant P-0251 owned edge inspected by that probe.

### Pattern type strings

The MCP and official 2025 `IFeature.GetTypeName2` table identify these exact
pattern strings used by r22:

- `APattern` -> `IFillPatternFeatureData`;
- `CirPattern`, `CurvePattern`, `LPattern`, and `TablePattern`;
- `DerivedCirPattern`, `DerivedHolePattern`, and `DerivedLPattern`;
- `LocalChainPattern`, `LocalCirPattern`, `LocalCurvePattern`,
  `LocalLPattern`, and `LocalSketchPattern`;
- `DimPattern`;
- `SketchPattern`; and
- `MirrorPattern` and `MirrorSolid`.

The guessed r21 strings `FillPattern`, `ChainPattern`, and `VariablePattern`
are not the documented return literals and have been removed.

MCP `IFace2.GetSeedFeature` returns the seed feature of a patterned, mirrored,
or copied body. Installed interop confirms
`Feature GetSeedFeature()` and the compatibility fallback
`Object GetPatternSeedFeature()`.

### Ordinate arrangement

- MCP `swDimensionType_e` and installed `swconst` agree:
  `swOrdinateDimension=1`, `swHorOrdinateDimension=7`,
  `swVertOrdinateDimension=8`, and `swAngularOrdinateDimension=16`.
- MCP `IDisplayDimension.AutoJogOrdinate` is parameterless and returns True
  only when auto-jog succeeds.
- MCP `IDisplayDimension.DisplayAsChain` applies to every dimension in one
  ordinate or angular-running set, but the 141-member interface inventory
  exposes no API for obtaining a stable ordinate-set identity.
- Installed interop confirms `Boolean AutoJogOrdinate()` and
  `Boolean DisplayAsChain`.

Source decision: when `AlignDimensions` returns False, r22 calls and checks
`AutoJogOrdinate` for ordinate types and validates annotation-position
readback. It does not construct a group key from annotation proximity or merge
separate sets into one side lane. Manual `SetPosition2` lanes remain limited to
non-ordinate dimensions.

### Annotation identity

MCP `IAnnotation.SetName` requires a new unique name and verifies uniqueness
before setting it. Installed interop confirms `String GetName()` and
`Boolean SetName(String)`. R22 therefore retains the r21 self-collision
fallback: pointer identity first, then equality of nonempty annotation names.

These checks establish E3 API/type-library compatibility only. The guarded SWP
deployment, full VBA Editor compile, authorized fixture run, QA, and
visual/manufacturing comparison remain separate gates.

## 2026-07-30 R23 Phase 0 preparation

R23 re-queried the local `solidworks-api` compatibility MCP and reflected the
installed SOLIDWORKS 2025 SP1.2 interop assemblies (`33.1.2.4`) before
preparing the read-only Phase 0 feature/curve probe. No live SOLIDWORKS
behavior is claimed by this subsection.

### Instant3D feature normalization and typed definitions

- `IFeature.GetTypeName2` identifies an Instant3D wrapper as `ICE`.
- `IFeature.GetTypeName` is the required source for the underlying type when
  `GetTypeName2="ICE"`.
- `IFeature.GetDefinition` returns a feature-data object when that feature
  type is supported; callers must reject `Nothing` or an unexpected interface.
- `IFeature.GetFaces` returns all faces owned by the feature. Its Remarks also
  state that `IFace2.GetFeature` returns only the oldest owner, so that method
  cannot be the sole ownership proof.
- `IExtrudeFeatureData2.AccessSelections(TopDoc, Component)` returns Boolean
  and places the model in rollback. A successful read-only access must end
  with `ReleaseSelectionAccess`; `ModifyDefinition` is not appropriate.
- Installed interop confirms `IsBossFeature() As Boolean`,
  `GetDepth(Forward) As Double`, and
  `GetEndCondition(Forward) As Long`.
- The same access/release contract applies to
  `IWizardHoleFeatureData2`. Its installed interface exposes
  `GetSketchPointCount` and `GetSketchPoints`.
- MCP and installed `swconst` agree that
  `swSpecifyConfiguration=3` for configuration-specific
  `IFeature.IsSuppressed2` calls.

The prepared probe therefore records both raw type values, normalizes `ICE`
before typed routing, proves the exact feature-data interface, records the
specified-configuration suppression result, and releases every successful
selection access. Feature names remain diagnostics only.

### Circular-edge read-order probe

- `IEdge.GetCurveParams3` returns `ICurveParamData`, not a Variant array.
- Its Remarks require `IEdge.GetCurve` first.
- `ICurve.IsCircle` remains the curve-type predicate.
- `ICurve.CircleParams` returns seven doubles:
  `[center xyz, axis xyz, radius]`.

`tools/r23-probes/Module_R23Phase0FeatureProbe.bas` reacquires the edge curve
and compares these two installed-build transactions:

1. `GetCurve -> GetCurveParams3 -> IsCircle -> CircleParams`;
2. `GetCurve -> IsCircle -> CircleParams -> GetCurveParams3`.

Source preparation does not select either transaction for production. The
P-0251 live output remains required.

### Drawing-component context correction

The R23 plan originally described `IView.GetVisibleComponents` as returning a
real/full `Component2`. The MCP Remarks explicitly contradict that wording:
the returned object is incomplete and, for example, does not support
`IComponent2.GetBodies3`.

The corrected R23 contract is:

- pass the exact limited `Component2` returned by
  `IView.GetVisibleComponents` to `IView.GetVisibleEntities2`;
- obtain visible `IDrawingComponent` objects through
  `IView.GetVisibleDrawingComponents`;
- use `IDrawingComponent.Component` when a component that fully supports
  `IComponent2` methods/properties is required; and
- runtime-prove that the limited drawing-context handle and full component
  converge on the same represented component and referenced configuration.

Installed interop confirms all four members. This is E3 interface evidence
only; the identity/convergence behavior remains a Phase 0 live probe.

### First installed-build feature-probe attempt

The first authorized P-0251 run reached
`R23_PROBE_END|visitedFeatures=15|status=COMPLETE`, but it is not accepted as
R23-004 through R23-006 evidence. The traversal guard used only
`ObjPtr(feature)`. In the installed VBA/COM runtime, transient feature wrappers
reused that address, so distinct tree features were incorrectly treated as
already visited. The transcript skipped the Hole Wizard, ICE, extrusion,
mirror, and most other relevant features even though they were present in the
active model tree.

That rejected transcript is preserved at:

`test_assets/iteration_evidence/r23/20260730-075811/live-probes/R23_FEATURE_PROBE_ATTEMPT2_BROKEN_TRAVERSAL.log`

The same run also established a binding detail that the probe must preserve:
`IFeature.IsSuppressed2(swSpecifyConfiguration, Array(configurationName))`
returned a scalar Boolean through SOLIDWORKS 2025 VBA rather than a Variant
array. Installed interop declares the return as `System.Object`, so the
correct VBA reader must accept both a nonempty scalar Boolean and an array.
The current R23 probe source now reports the shape explicitly as
`Active:Scalar`, `Suppressed:Scalar`, `Active:Array`, or
`Suppressed:Array`.

The corrected probe additionally:

- combines feature name, `GetTypeName2`, and wrapper address in its recursion
  key so wrapper-address reuse cannot suppress differently named features;
- includes base `EXTRUSION` features in the typed
  `IExtrudeFeatureData2` probe; and
- records installed `IExtrudeFeatureData2.GetContoursCount` plus the
  `Contours` return shape and owned `ProfileFeature` subfeature count; and
- mirrors every structured record to a timestamped evidence log so the
  Immediate Window buffer is not the sole transcript;
- remains read-only and pairs successful selection access with release.

The corrected exported source still requires a user-operated full-project
compile and P-0251 rerun. Until that transcript includes the three ICE
features, both Hole Wizard families, mirror evidence, and the target cylinder
families, the production Phase 0 gate remains closed.

### Expanded import-transaction harness

The local API MCP was not callable after the workstation outage. The prepared
import harness therefore uses the earlier recorded MCP contract together with
fresh installed-interop reflection; it does not present the compatibility
snapshot as installed-build runtime proof.

Installed SOLIDWORKS 2025 SP1.2 interop `33.1.2.4` confirms:

- `IDrawingDoc.InsertModelAnnotations4(Int32, Int32, Boolean, Boolean,
  Boolean, Boolean, Boolean, Boolean) -> Object`;
- `IView.GetAnnotations() -> Object` and
  `IView.GetDisplayDimensions() -> Object`;
- `IDisplayDimension.GetDimension2(Int32) -> IDimension`;
- `IDisplayDimension.IsHoleCallout() -> Boolean` and
  `GetHoleCalloutVariables() -> Object`;
- `IAnnotation.GetAttachedEntities3() -> Object`,
  `GetAttachedEntityTypes() -> Object`, and `GetPosition() -> Object`;
- `IDimension.GetSystemValue3(Int32, Object) -> Object`;
- `IDimensionTolerance.Type`, `FitType`, `FitDisplayStyle`,
  `GetHoleFitValue`, `GetShaftFitValue`, `GetMinValue2`, and
  `GetMaxValue2`; and
- `IDisplayData` line, arrow, text, arc, ellipse, and polyline readback
  members used for raw display-geometry evidence.

Installed `swconst` confirms `swInsertDimensions=8`,
`swInsertDimensionsMarkedForDrawing=32768`,
`swInsertHoleWizardProfileDimensions=65536`,
`swInsertHoleWizardLocationDimensions=131072`,
`swInsertHoleCallout=1048576`, and
`swInsertTolerancedDimensions=16777216`. With datums, GTols, and notes, the
prepared full comparison mask is `18055274` (`0x113806A`).

The guarded disposable harness has two independent fresh-drawing entry points:

1. selected primary anchor, `AllViews=True`, `DuplicateDims=True`;
2. deterministic section, side, and primary calls,
   `AllViews=False`, `DuplicateDims=True`.

It records per-view before/after deltas, returned-to-final annotation
ownership, source dimension identities, nominal/tolerance/fit fields, Hole
Wizard callout variables, attachments, exact duplicate identities, and raw
display primitives. The three-source custom manifest passed read-only
deployment preflight against the copied macro. Full-project VBA compilation,
both installed-build import runs, screenshots, and visual-overlap judgment
remain pending, so R23-007 and R23-008 are still open.

## 2026-08-01 R23 Phase 8 section-dimension contracts

MCP corpus lookups for the section-dimension engine. Enum values are corpus
evidence and still require SW2025 Object Browser confirmation; the method
contracts below come from each member's own Remarks.

### The IDimension tolerance members are all obsolete

Four members the Phase 0 probe used are marked `deprecated: true`:

| Obsolete member | Superseded by |
|---|---|
| `IDimension.GetToleranceValues` | `IDimensionTolerance.GetMaxValue`/`GetMinValue` |
| `IDimension.SetToleranceValues` | `IDimensionTolerance.SetValues` |
| `IDimension.GetToleranceFitValues` | `IDimensionTolerance.GetHoleFitValue`/`GetShaftFitValue` |
| `IDimension.SetToleranceFitValues` | `IDimensionTolerance.SetFitValues` |

`IDimension.Tolerance` is the supported entry point and returns
`IDimensionTolerance`. R23 Phase 8 uses only that route; a contract asserts
none of the four obsolete names appears in the module.

`GetMinValue2` and `GetMaxValue2` return a STATUS and pass the value back
through a ByRef argument. A zero value paired with a failed status is not a
zero tolerance, so both are reported together.

### Tolerance write order is constrained by the Remarks

`IDimensionTolerance.SetValues2(MinValue, MaxValue, WhichConfigurations,
Config_names)` states that values cannot be set while the tolerance type is
`swTolType_e.swTolNONE`. `IDimensionTolerance.FitType` states it is only
available for `swTolFIT`, `swTolFITTOLONLY` and `swTolFITWITHTOL`.

The type is therefore set first, then `SetFitValues(HoleFit, ShaftFit)`,
then `SetValues2`. `swSetValueInConfiguration_e.swSetValue_NoConfiguration =
-1` is documented as "ignore configurations in drawing sketches", which is
what a drawing dimension needs.

`IDimensionTolerance.Type` also carries a caution worth recording: in
SOLIDWORKS 2016 and later, setting the tolerance type for a hole's display
dimension with multiple values in a callout must go through
`ICalloutVariable.ToleranceType`, and `IDimensionTolerance.Type` does not
override it. Phase 8 sets the type on a plain diameter dimension, not on a
callout variable, so the caution does not bite - but Phase 6's callout work
is the place it would.

### Dimension creation requires selection by location

`IModelDoc2.AddDimension2(X, Y, Z)` creates a display dimension for the
already-selected entities. Its Remarks are explicit that selections must be
made by LOCATION and not by name: if a name is passed to
`IModelDocExtension.SelectByID2`, the selection routines use the name and
ignore the coordinates, and the dimensioning routines then pick a line
endpoint at random.

`IModelDocExtension.AddDimension(X, Y, Z, Direction)` is the variant to use
only when the pre-selected entities do NOT unambiguously define what to
dimension and an extension line is needed; `Direction` is
`swSmartDimensionDirection_e` (Right 0, Up 1, Left 2, Down 3). Phase 8 uses
`AddDimension2`, because a section requirement names the two entities whose
distance it is.

Both note that `swUserPreferenceToggle_e.swInputDimValOnCreate` suppresses
the value-entry dialog, which any unattended run needs.

### Enum values used

| Enum | Members used |
|---|---|
| `swDimensionType_e` | ordinate base 1, linear 2, diameter 6, horizontal ordinate 7, vertical ordinate 8, horizontal linear 11, vertical linear 12, diametric linear 15, angular ordinate 16 |
| `swTolType_e` | `swTolNONE` 0, `swTolFIT` 7, `swTolFITWITHTOL` 8, `swTolFITTOLONLY` 9 |
| `swFitType_e` | user 0, clearance 1, transitional 2, press 3 |
| `swSetValueInConfiguration_e` | `swSetValue_NoConfiguration` -1 |
| `swDrawingViewTypes_e` | `swDrawingSectionView` 2 |

Note `swTolType_e` lists both `swTolFIT` and `swTolMETRIC` as 7; they are
aliases, not a parse defect.

### View-scoped readers

`IView.GetDisplayDimensions` and `IView.GetNotes` both return the whole
array for ONE view. Their obsolete predecessors,
`GetFirstDisplayDimension5` with `IDisplayDimension.GetNext5` and
`GetFirstNote` with `INote.GetNext`, are what the Remarks recommend
replacing. `GetFirstDisplayDimension5` additionally walks the sheet rather
than the view, which would attribute another view's dimensions to the
section.

`IAnnotation.GetAttachedEntityTypes` returns `swSelectType_e` values
positionally matching `GetAttachedEntities3`, and returns EMPTY arrays when
the annotation is attached to nothing. A dangling attachment appears as
`swSelNOTHING` with a Nothing in the matching entity slot - which is how
"associative" is proved rather than assumed.

## 2026-08-04 R23-502 outline datum contract

The local SOLIDWORKS MCP had no callable tool in this session. Official Help
confirms that [IPartDoc.GetBodies2](https://help.solidworks.com/2024/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IPartDoc~GetBodies2.html)
returns visible bodies, [IBody2.GetEdges](https://help.solidworks.com/2025/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IBody2~GetEdges.html)
returns their edge array, [IView.GetCorrespondingEntity](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IView~GetCorrespondingEntity.html)
accepts a part edge and returns its corresponding drawing entity or Nothing,
and [IView.SelectEntity](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IView~SelectEntity.html)
selects an entity in a drawing view. The implementation stays inside the
already live-proven Route A / Route D mapping boundary.

Live P-0251 r29 evidence settles the coordinate rule: `IView.GetOutline` is
an enclosing page-frame view bound, not part-outline geometry. None of the 27
horizontal two-vertex model edges exactly coincided with its lower boundary;
the selected lowest mapped edge lies 5.94 mm above it. R23-502 therefore
chooses the lowest mapped horizontal model edge and proves target-view
ownership, rather than selecting the view rectangle or a lowest hole centre.
`Drawing View4` selects `SolidBodyEdge_0_49` and `Drawing View7` selects
`SolidBodyEdge_0_73`, both at page Y `0.137763153`.

## 2026-08-01 R23 Phase 9 envelope and layout contracts

### Which sources state their frame, and which do not

| Member | Frame stated by the Remarks |
|---|---|
| `IView.GetOutline` | Yes - bounding box in metres on the drawing page, `[Xmin, Ymin, Xmax, Ymax]` |
| `IAnnotation.GetPosition` | Yes - in a drawing, relative to the sheet origin at the lower-left corner |
| `INote.GetExtent` | Yes - six doubles, lower-left and upper-right, in sheet space |
| `IAnnotation.GetLeaderPointsAtIndex` | No, but computed from the text and attachment points |
| `IDisplayData.GetLineAtIndex3` | **No** |
| `IDisplayData.GetTextPositionAtIndex` | Offset only - see below |
| `IView.GetSectionLineInfo2` | **No**; Phase 0 proved VIEW-SKETCH frame |

R23 Phase 9 converts section geometry through the inverse of Module17's
transform and COUNTS whether display-data points agree with the view's own
outline rather than asserting a frame the corpus does not state.

### GetTextPositionAtIndex is an offset, not a coordinate

Its Remarks say the returned values "are actually offset values from the
origin of this display data". Treated as absolute coordinates they pull
every envelope towards the sheet origin, which looks like a plausible
rectangle and is wrong. Phase 9 adds them to the annotation position and
skips the text contribution entirely when that position is unavailable.

### GetLeaderStyle carries bitmask flags the corpus mangles

`swLeaderStyle_e` returns its base members correctly - `swNO_LEADER` 0,
`swSTRAIGHT` 1, `swBENT` 2, `swUNDERLINED` 3, `swSPLINE` 4, `swVDA` 8 - but
every attachment flag comes back with `value: "0"` and a description reading
`x100 or 256`, `x200 or 512`, `x400 or 1024`, `x800 or 2048`, `x1004 or
4100`. That is the same leading-`0x` parse defect recorded for
`swAutoInsertCenterMarkTypes_e`: the real values are 256, 512, 1024, 2048
and 4100.

Because `GetLeaderStyle` returns the base style OR-ed with those flags, the
documented point-count rule (0 for no leader, 2 for straight/underlined, 3
for bent) cannot be applied without masking. Phase 9 sidesteps it entirely
and consumes whole XYZ triples from the array `GetLeaderPointsAtIndex`
actually returns.

### GetSectionLineInfo2's grammar is ambiguous

Its Remarks give
`[numSectionLines, layer, {numSegments, {lineType, startPt[3], endPt[3]},
arrowStart1[3], arrowEnd1[3], arrowWidth1, arrowHeight1, arrowStyle1,
arrowStart2[3], arrowEnd2[3], arrowWidth2, arrowHeight2, arrowStyle2,
textPt1[3], textPt2[3], textHeight}]` - one `layer` double overall.
`IView.GetSectionLineCount2`'s Remarks say its `Size` "includes the layer-ID
double for each section line".

Both readings are walked in a dry run and the one whose consumption matches
the array length exactly is used; neither is assumed. `textHeight` is in
metres and `IView.GetSectionLineStrings` supplies the actual label text.

### Layout members

`IView.Position` gets or sets the X and Y of the view's geometric centre
relative to the sheet origin. Its Remarks carry two constraints worth
recording: alignments are honoured exactly as in the UI, so an aligned view
moves only along its alignment vector and drags its aligned children with
it; and `IView::SetViewPosition` is the member for moving a view
independently of its children. `IModelDoc2.EditRebuild3` is needed after
view changes for the graphics to reflect them.

`IView.ScaleDecimal` gets or sets the view scale as a decimal
(`ScaleRatio` 3:2 is `ScaleDecimal` 1.5). Phase 9 only ever READS it -
R23-907 forbids reducing an approved scale to force a fit - and a contract
asserts no assignment exists in the module.

## 2026-08-01 live findings: nominal routes, Diametric, and a probe trap

### GetSystemValue3 declines for drawing-authored reference dimensions

Live, on `P-0251-14A-001.SLDDRW`, `Section View J-J`: seven display
dimensions named `RD1..RD7@Drawing View6`, every one returning no value from
`IDimension.GetSystemValue3(swThisConfiguration, Empty)`. The dimension
objects were valid - `IDimension.Tolerance` answered on the same objects,
returning real `Type`, `FitType`, `GetHoleFitValue`, `GetMinValue2` and
`GetMaxValue2` results.

The distinction that matters: Phase 0 read `D1@Sketch4`, a MODEL dimension
imported into a drawing view, and the configuration route worked. A drawing
REFERENCE dimension has no configuration, and the supported route can
decline rather than fail.

R23 now tries `swThisConfiguration`, then `swAllConfiguration`, then the
obsolete `GetSystemValue2("")` and `IDimension.SystemValue`, and records
which route answered. Both obsolete members are documented as returning
system units. Which one a reference dimension actually answers is still
open; the next run settles it, and the routes that prove unnecessary can
then be removed.

### The same fixture has been seen with two dimension types

| Run | View | Records |
|---|---|---|
| Phase 0, 2026-07-31 | `Drawing View3` | 17 dimensions, `D1@Sketch4`/`D1@Sketch6` at `swDiameterDimension = 6` |
| Phase 8, 2026-08-01 | `Section View J-J` | 7 dimensions, all `swLinearDimension = 2`, named `RD1..RD7` |

Neither reading is wrong; the drawing changed. Any rule that requires a
single `swDimensionType_e` value for a diameter rejects one of the two real
states.

`IDisplayDimension.Diametric` is the member that separates them: it gets or
sets whether the dimension displays as diameter/doubled-distance rather than
radial/single-distance, and its Remarks say it toggles radial vs diameter
and radial-linear vs diametric-linear, affecting no other type. R23 records
it for every section dimension.

### Live H7 readback on a drawing reference dimension

`RD4@Drawing View6`: `toleranceType=8` (`swTolFITWITHTOL`), `fitType=0`
(`swFitUSER`), `holeFit=H7`, `minimumStatus=0|minimumM=0.000000`,
`maximumStatus=0|maximumM=0.000025`, two attached entities of type
`swSelEDGES = 1`.

This is independent corroboration of the R23-806 finding: the H7 fit exists
on the DRAWING and not in the model, whose `D1@Sketch4` reported
`toleranceType=0` and `fitType=-1`.

### MeasureControlledSheetRegions is not safe to call from a probe

`Module8_RuntimeSupport.MeasureControlledSheetRegions` is a production
fail-closed gate. Two properties make it unusable in a read-only run:

- it requires an `ITitleBlock` or a proved legacy title-block rectangle, and
  fails the whole stage without one - the designer's reference drawing has
  neither; and
- it SETS `ISheet.SheetFormatVisible`, recording the mutation on the
  evidence ledger.

A read-only probe must measure for itself. `ISheet.GetSize` and
`ISheet.GetZoneMargin` are both read-only.
`swZoneMargin_e` is `swZoneTopMargin` 0, `swZoneBottomMargin` 1,
`swZoneRightMargin` 2, `swZoneLeftMargin` 3.

## 2026-08-01 second live run: three contracts settled

### Which nominal route answers a drawing reference dimension

Live, on all seven `RD1..RD7@Drawing View6` dimensions in
`Section View J-J`:

| Route | Result |
|---|---|
| `GetSystemValue3(swThisConfiguration, Empty)` | declined |
| `GetSystemValue3(swAllConfiguration, Empty)` | declined |
| `GetSystemValue2("")` (obsolete) | **answered**, exact metres |

Nominals returned: 0.018, 0.012, 0.0115, 0.040, 0.047, 0.1736, 0.1048.

The supported `GetSystemValue3` route is what answers for an imported model
dimension - Phase 0 proved that on `D1@Sketch4`. For a drawing-authored
reference dimension on this build, the obsolete `GetSystemValue2` is the
only route that returns a value. R23 keeps both and names which answered;
`swAllConfiguration` was removed after declining on all seven.

### IDisplayDimension.Diametric is False for the drawing's bore dimensions

All seven section dimensions returned `Diametric = False` with the read
succeeding. That includes the 47 carrying H7 and the 40. So a diameter on
this drawing is not necessarily a diametric record, and
`IDisplayDimension.Type2` is not necessarily 6 or 15 either.

Three independent sources can establish that a dimension reads as a
diameter, and R23 records which one applied:

1. `Type2` is `swDiameterDimension` 6 or `swDiametricLinearDimension` 15;
2. `IDisplayDimension.Diametric` is True; or
3. the text prefix carries the symbol -
   `GetText(swDimensionTextPrefix)` for what is drawn and
   `GetText(swDimensionTextPrefixDefinition)` for the authored form, where
   SOLIDWORKS writes the `<MOD-DIAM>` token.

`swDimensionTextParts_e`: `swDimensionTextAll` 0 (SetText only),
`swDimensionTextPrefix` 1, `swDimensionTextSuffix` 2,
`swDimensionTextCalloutAbove` 3, `swDimensionTextCalloutBelow` 4,
`swDimensionTextPrefixDefinition` 5, `swDimensionTextSuffixDefinition` 6,
`swDimensionTextCalloutAboveDefinition` 7,
`swDimensionTextCalloutBelowDefinition` 8. `GetText` does not support hole
callouts, and `swDimensionTextAll` is not valid for it.

### GetSectionLineInfo2 arrow block: 9 doubles, proved by item count

`Drawing View4` returned 49 items for a section line with three segments.
The grammar that fits is:

```
2                       numSectionLines, layer
1                       numSegments
7 x 3                   lineType + startPt[3] + endPt[3], per segment
9                       arrow 1: start[3] + end[3] + width + height + style
9                       arrow 2: same
7                       textPt1[3] + textPt2[3] + textHeight
--
49
```

Reading the arrow block as 11 doubles - the mistake in R23's first Phase 9
source - matches no segment count at all, which is how it was caught. A
parser that consumes the wrong stride still returns plausible-looking
coordinates, so the item-count check is not optional.

The one-`layer`-overall reading is the one that fits; the alternative from
`GetSectionLineCount2`'s Remarks was tried and did not.

## 2026-08-04 R23 probe-automation tool: VBE control resolution

Live-verified. First successful `R23_RunAllProbes` run completed this
date; full narrative and evidence path in
[R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md](R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md)
section 12. This entry keeps the API-contract details.

**Resolved live:** `id=578`, `caption=Compile Fable` (raw
`CommandBarControl.Caption` is `"Compi&le Fable"` - see the accelerator-
marker finding below), under `Menu Bar` → `Debug`.
`R23_COMPILE_VERDICT|...|enabledBefore=True|enabledAfter=False`
`|verdict=Clean` on the accepted run - the enabled-state flip this
design relies on is confirmed real, not just documented behaviour.

### The VBIDE Compile control ID is not in this MCP corpus

Confirmed by direct search, not just absence in one query.
`solidworks_search_api` for `"CommandBars VBE compile project control"`
and `"VBA project compile macro editor"` returns only SOLIDWORKS geometry
and macro-path records - zero VBIDE hits. The corpus root is the
SOLIDWORKS API Help (`https://help.solidworks.com/2025/english/api`);
VBE `CommandBars` control IDs belong to the Microsoft Office/VBA
extensibility model, a different product's documentation entirely. No
amount of re-querying this MCP will produce that number - it has to come
from a live enumeration or the Object Browser, never from recall.

### How the tool avoids guessing it anyway

`Module20_ProbeRunner` never hardcodes a control ID. It walks
`VBE.CommandBars` and every control's own `.Controls` collection
recursively (`WalkVbeControls`), matching `.Caption` for a
case-insensitive `"compile"` substring, and logs the resolved `.Id`
alongside the `.Caption` it matched - so the ID becomes run evidence
instead of an assumption. `R23_EnumerateVbeControls` dumps the full walk
as a fallback if no caption matches, per Agents.md's "never guess a
control ID" rule. This mirrors the one VBIDE access pattern already
proved in this repo:
`tools/swp-deploy/Module0_SourceDeployment.bas:214` reaches
`targetProject.VBE.CommandBars.FindControl(1, 3, "", False)` for the Save
command, with the same `Nothing`/`.Enabled` guards before `.Execute`.

### Document-switch calls: MCP-looked-up, now locally confirmed working

The plan's design section did not say how `R23_RunAllProbes` gets the
correct document active for each probe - `R23_ProbeFeatureCatalog` needs
the authorized part as `ActiveDoc`, the other eight need the drawing
(confirmed by reading each probe's own `swDraw.GetType <>` guard). Three
members were looked up before writing the fix:

- `ISldWorks.GetFirstDocument` / `IModelDoc2.GetNext` - documented pattern
  for enumerating every open document in the session; used to find the
  open authorized part and the open drawing without assuming which one is
  foreground.
- `ISldWorks.ActivateDoc3(Name, UseUserPreferences, Option, out Errors)` -
  Remarks warn that an extension-less `Name` can collide between a part
  and an assembly of the same base name, so the runner passes
  `IModelDoc2.GetTitle` (filename with extension), matching the pattern
  already proved by `Module8_RuntimeSupport.ActivateDrawingDocument`.
  `Option = swRebuildActiveDoc (2)` was cross-checked against
  `swRebuildOnActivation_e` and matches that existing proved usage.
- `ISldWorks.OpenDoc6` - Remarks state explicitly that it "does not
  activate and display the document", which is *why* the explicit
  `ActivateDoc3` switch between probe 1 and probes 2-9 is necessary rather
  than optional. `Options = swOpenDocOptions_ReadOnly (2)` was added to
  both `OpenDoc6` calls in `Run-R23Probes.ps1` as an extra safety layer on
  top of the fixture guard - Agents.md forbids saving a fixture, and a
  read-only handle makes that impossible regardless of macro behaviour.

All four numeric values came from `solidworks_get_enum_values`/
`solidworks_lookup_method` against the 2025 corpus snapshot.
`swDocPART=1`, `swDocDRAWING=3`, `swRebuildActiveDoc=2`, and
`swOpenDocOptions_ReadOnly=2` are now confirmed against the installed
SW2025 SP1.2 build too: the 2026-08-04 live run opened both documents
read-only, switched the active document twice via `ActivateDoc3`
(`R23_RUN_ACTIVATE|title=...|errors=0|succeeded=True` for both), and
every probe's `part=` field resolved to the correct fixture.

### PowerShell cannot call `OpenDoc6` directly (live finding)

Two independent PowerShell-side failures, both live, neither an API
contract problem:

1. `$solidWorks.OpenDoc6($path, $type, $options, '', [ref]$e, [ref]$w)`
   through plain COM late binding raises
   `TYPE_E_ELEMENTNOTFOUND (0x8002802B)`. PowerShell's automatic method
   resolution cannot match the `out`-parameter overload.
2. `[SolidWorks.Interop.sldworks.ISldWorks] $rawComObject` - a direct
   bracket-cast of the `GetActiveObject()` proxy, tried after loading the
   interop assembly with `Add-Type` - raises `Cannot convert the
   "System.__ComObject" value... to type "...ISldWorks"`. This is a
   PowerShell type-conversion limitation, not a version mismatch: the
   identical cast, done inside compiled C# (`(ISldWorks)application`),
   already works and is what `SolidWorksMacroInvoker.cs` has used
   successfully for `RunMacro2` since the deploy tooling was built.

Fix: `tools/swp-deploy/SolidWorksDocumentOpener.cs`, same shape as
`SolidWorksMacroInvoker.cs` (early-bound cast, `InvalidCastException`/
`COMException`-triggered late-bound `InvokeMember` fallback), exposed
from PowerShell as `Open-SolidWorksDocument` in
`tools/swp-deploy/Invoke-SolidWorksMacro.ps1`. Anyone adding a new
PowerShell-to-SOLIDWORKS call with `out`/`ref` parameters should expect
to need the same compiled-helper treatment rather than a raw COM call.

### Same-title open conflicts are fail-closed (2026-08-04)

The installed SOLIDWORKS 2025 SP1.2 `SolidWorks.Interop.swconst` type
library confirms `swFileLoadError_e.swFileWithSameTitleAlreadyOpen = 65536`.
The runner received exactly that bit when it requested local
`test_assets\models\P-0251-14A-001.SLDPRT`; live enumeration through the
already documented `ISldWorks.GetFirstDocument` / `IModelDoc2.GetNext`
pattern identified the conflicting V: sibling.

`SolidWorksDocumentOpener.cs` now returns title/path pairs and
`Open-SolidWorksDocument` reports them on this error. It deliberately never
calls `ISldWorks.CloseDoc`: closing the same-title document could discard a
user's unsaved manual drawing. This is runner diagnostic evidence only; no
VBE compilation, probe execution, or drawing mutation occurred.

### VBIDE captions carry the raw accelerator marker (live finding)

`CommandBarControl.Caption` returns the *authoring* string, including
the `&` that marks the keyboard-accelerator letter - live value for the
Compile command was `"Compi&le Fable"`, not `"Compile Fable"`. A caption
search that does not strip `&` before matching will never match whatever
word contains the marked letter, silently, with no error raised -
`InStr` just returns 0 every time. Anything that matches or displays a
VBIDE `.Caption` in this repo must run it through `CleanControlText` (or
equivalent) first for both purposes, not just for display.

## 2026-08-04 r30 review-correction API contracts

The local `solidworks-api` MCP was available for r30. Its corpus is the
SOLIDWORKS 2025 Help compatibility snapshot; the contracts below are MCP
evidence and remain distinct from installed-SP1.2 runtime proof.

- `ICurve.IsLine()` returns a Boolean: true for a line and false for every
  other curve type. `OutlineDatumForModelEdge` now normalizes that COM Boolean
  before accepting equal endpoint Y values, so a horizontal arc chord cannot
  become an outline datum.
- `IView.GetVisibleEntities2(Component2, swViewEntityType_e)` returns only
  entities not completely obscured in that drawing view. The edge inventory is
  therefore the visibility contract; `IView.GetCorrespondingEntity` proves
  correspondence, not visibility. r30 maps a datum only when exact identity
  occurs in that inventory. A part drawing with no usable `Component2` fails
  closed rather than using selection ownership as a substitute.
- `IView.UseParentScale` means the view matches its parent scale;
  `UseSheetScale` means it matches the sheet scale; and `ScaleDecimal` reads
  or writes the decimal scale. All three state that a rebuild is needed after
  related view changes. r30 preflights those relationships and refuses an
  automatic rescale of an approved isometric or `swDrawingDetailView = 3`
  view. `IView.GetOrientationName` returns predefined names such as
  `*Isometric` and returns empty for detail views, so both orientation and
  type are used.
- VBA compile errors happen before a procedure-level `On Error` handler can
  execute. `R23_TouchAllModules` is no longer presented as compile-error
  localization. The runner stops and requires the VBE dialog text plus its
  highlighted line.

### r30 installed-SP1.2 probe evidence

The focused P-0251 scratch runner at
`test_assets/iteration_evidence/probe_runs/20260804_154251/probe_log.txt`
compiled cleanly and completed all nine read-only probes. In the two
ordinate-bearing views, candidate straight edges were found and mapped, but no
candidate survived the visible-entity identity test; both vertical schemes
reported `NoMappedBottomEdge`. This confirms the fail-closed behavior on the
installed build. It is not a manufacturing or production-mutation acceptance.

## 2026-08-04 r31 user-accepted layout policy

No new SOLIDWORKS API member or enum is introduced. The production route no
longer calls the existing `IView.Position` or `IView.ScaleDecimal` mutation
path after final content creation. Initial structural placement remains
separate. `FINAL_LAYOUT` records
`UserAcceptedLayoutAsIs|automaticClearance=DeferredByUser`; this is a user
policy record, not API or installed-build evidence that the measured layout is
clear.

## 2026-08-04 r37 outline-datum visibility fallback

The local `solidworks-api` MCP was used before this change. Its 2025-help
compatibility snapshot documents `IView.GetPolylines7(CrossHatchOption,
out Polylines)` as returning visible model and silhouette edges. It also
documents `IView.SelectEntity` as selecting an entity in the specified drawing
view.

Installed-SP1.2 runner evidence at
`test_assets/iteration_evidence/probe_runs/20260804_164014/probe_log.txt`
shows both P-0251 HLV ordinate-bearing views returning
`GetPolylines7|status=NoEdges|error=0`. The code therefore does not treat a
missing edge array as an identity mismatch. It falls back only in that explicit
state: `IView.SelectEntity` must
select the candidate model edge and `ISelectionMgr.GetSelectedObjectsDrawingView2`
must return the requested view. The runner proved that contract for the two
straight lower outline edges (`Drawing View4` and `Drawing View7`).

This is installed-build selection and mapping evidence, not proof that an
ordinate has been created. Creation still requires the separate mutating
acceptance run and dimension readback.

## 2026-08-06 r8 view classification characterised on the installed build

The §6 caveat in `BASELINE_TO_REFERENCE_DRAWING_GAP.md` — "what each actually
returns for this pipeline's views has **not** been probed" — is now closed. The
r8 QA report at
`test_assets/iteration_evidence/macro_qa/20260806_051241_P-0251-14A-001/QA_REPORT.txt`
carries the raw returns for every view on the sheet:

```
Drawing View1   | Type=7 | Orientation=*Front
Drawing View2   | Type=7 | Orientation=*Bottom
Drawing View3   | Type=7 | Orientation=*Right
Drawing View4   | Type=7 | Orientation=*Left
Drawing View5   | Type=7 | Orientation=*Isometric
Section View J-J| Type=2 | Orientation=(empty)
```

Three contracts confirmed on SOLIDWORKS 2025 SP1.2:

- **`IView.Type` cannot separate orientations.** Every view created by
  `CreateDrawViewFromModelView3` returns `7` (`swDrawingNamedView`) regardless
  of orientation. The caveat recorded before the probe was correct: `Type`
  alone is useless for telling a front view from an isometric one.
- **`IView.GetOrientationName` round-trips exactly.** It returns the same
  string that was passed to `CreateDrawViewFromModelView3`, leading `*`
  included. This is the member that carries the classification.
- **Section views match the documented empty-string case.** `Type=2`
  (`swDrawingSectionView`) with `GetOrientationName` returning `""`, as the
  2025 help states for section, detail, projected and unfolded views.

`Module8_ViewClassifier` therefore keys on `Type` for the structural cases
(sheet, section, detail, projected) and on `GetOrientationName` for the
orientation cases. The name-substring fallback is retained but was not
exercised on this build.

This is view-classification evidence only. It says nothing about datum
selection or ordinate values.

## 2026-08-06 r9/r10 SOLIDWORKS COM Boolean: only "= False" is reliable

The rule recorded on 2026-07-31 said `If Not <comBooleanCall>` is a defect
because `Not` yields `-2`, which VBA treats as True. r9 shows the rule was too
narrow. **`= True` is not the mirror image of `= False`.**

Same test, same view, same 64 visible edges, three forms:

| form | circular edges found |
|---|---|
| `If Not swCurve.IsCircle Then GoTo NextEdge` | 0 (r6, over 349 edges) |
| `If swCurve.IsCircle = True Then` | **0** (r9) |
| `If swCurve.IsCircle = False Then GoTo NextEdge` | 35 (r8, r10) |

VBA `True` is `-1`. The value `ICurve.IsCircle` returns is truthy but is not
`-1`, so an equality test against `True` never matches. r9 produced a drawing
with zero dimensions from a rewrite that looked like a no-op refactor of a
proven line.

**The safe form is `Not (<comBooleanCall> = False)`**: compare against `False`,
which works, then negate the resulting VBA Boolean, which is a real VBA
Boolean and safe to negate. `Module5_FallbackDimensionEngine` uses this for
both `ICurve.IsCircle` and `ICurve.IsLine`.

Widen the standing rule: any comparison other than `= False` against a
SOLIDWORKS COM Boolean is a defect, including `= True` and bare truthiness.

## 2026-08-06 r10 a linear edge is accepted as an ordinate datum

Open question from r9: `IModelDocExtension.AddOrdinateDimension` documents only
"select the base entity to act as the datum point", and nothing said whether a
straight edge qualifies. The engine was written to retry with holes-only
candidates if the richer set was rejected, and to count the retries.

The r10 run at
`test_assets/iteration_evidence/macro_qa/20260806_054312_P-0251-14A-001/QA_REPORT.txt`
created both chains with **zero holes-only retries**, with the Y chain's datum
reported as `Edge:offsetFromTarget=0.00mm`:

```
Ordinate edges seen: 64 (circular: 35, linear: 29)
X: 9 stations (holes=6, edges=3), datum Hole:offsetFromTarget=0.00mm
Y: 7 stations (holes=5, edges=2), datum Edge:offsetFromTarget=0.00mm
Ordinate chains created: 2 of 2 attempted
```

So on this build a straight model edge is valid both as the base entity and as
a member of an ordinate group. Gap A4's premise is confirmed: restricting
candidates to circles was a self-imposed limit, not an API constraint.

**Not established:** whether every axis-parallel edge yields a *stable*
attachment. Three stations on the r10 sheet render `0.00` in the dangling
colour while the report shows their true offsets, so some attachments do not
survive. See CURRENT_STATUS.md.

## 2026-08-06 r14/r15 IAnnotation.IsDangling is a post-rebuild property

A freshly created ordinate reports `IsDangling = False`. The flag only becomes
True after the drawing rebuilds.

Measured, not inferred. r14 ran the dangling prune inside the ordinate stage,
immediately after `AddOrdinateDimension`, and reported
`DanglingFound = 0` - the line did not print at all. The QA readback, running
later in the same run against the same view and after
`IModelDoc2.ForceRebuild3`, reported `readback: 14 dims, 2 dangling`. Neither
`IAnnotation.Select3` nor `IModelDocExtension.DeleteSelection2` had refused;
there was nothing to select.

r15 moved the prune downstream of `ForceRebuild3` in
`Module2_DrawingPipeline` and it worked first time:

```
Dangling ordinates: 2 found, 2 deleted (0 select-refused, 0 delete-refused)
readback: 12 dims, 0 dangling
```

**Consequence for any future attachment check:** it must run after a rebuild.
Placed earlier it reads clean on broken geometry, which is the silent-pass
failure mode this engine has produced repeatedly.

`IAnnotation.Select3(False, Nothing)` and
`IModelDocExtension.DeleteSelection2(0)` both succeed on an ordinate group
member, with pick mode reset and the owning view activated first. Deleting one
member does not disturb the rest of the group.

## 2026-08-06 r16 constant provenance gate, and three wrong display modes

`tools/solidworks-automation-companion/tests/test_api_constant_provenance.py`
now fails the suite when a `sw*` constant is compiled into the trunk without a
provenance record in this document. Running it for the first time listed eight
undocumented constants. Checking them found three wrong values.

### swDisplayMode_e — three defects, none of which ever raised

`IView.SetDisplayMode3` Remarks (MCP, 2026-08-06): *"To display a drawing view
shaded with edges, set swDrawingsDefaultDisplayTypeHLREdgesWhenShaded to True
and set Mode to swSHADED"*, and any `swFACETED_*` value passed in `Mode` *"are
treated the same as swWIREFRAME, sw_HIDDEN_GREY, and sw_HIDDEN,
respectively"*.

| constant | was | is | enum member |
|---|---|---|---|
| `swDisplayMode_HiddenLinesRemoved` | 3 | **2** | 3 is `swSHADED`; HLR is `swHIDDEN` = 2 |
| `swDisplayMode_HiddenLinesVisible` | 1 | 1 | `swHIDDEN_GREYED` = 1, described as "Hidden Lines Visible (HLV)" — correct |
| `swDisplayMode_ShadedWithEdges` | 6 | **3** | 6 is `swFACETED_HIDDEN`, silently treated as HLR |

Consequences on every run up to r15: `UseHLR = True` rendered views **shaded**
rather than hidden-lines-removed, and the isometric view was rendered
**hidden-lines-removed** rather than shaded with edges. Fifteen live runs, no
error, no QA signal.

Shaded-with-edges now uses `swSHADED` with the method's `Edges` argument, which
`ConfigureView` already passes `True`. The user-preference route in the Remarks
is deliberately not used: mutating an operator's installation setting to render
one view is a side effect outside this macro's remit.

**Unverified, follow-up:** `IView.SetDisplayMode3` is marked obsolete,
superseded by `IView.SetDisplayMode4`. The trunk still calls the obsolete form.
No signature comparison has been made.

### The remaining five constants

MCP evidence, 2026-08-06. Verify in the SW2025 Object Browser before
acceptance.

| constant | value | source enum |
|---|---|---|
| `swDrawingSheetType` | 1 | `swDrawingViewTypes_e.swDrawingSheet` |
| `swDrawingSectionViewType` | 2 | `swDrawingViewTypes_e.swDrawingSectionView` |
| `swDrawingDetailViewType` | 3 | `swDrawingViewTypes_e.swDrawingDetailView` |
| `swDrawingProjectedViewType` | 4 | `swDrawingViewTypes_e.swDrawingProjectedView` |
| `swDefaultTemplateDrawing` | 10 | `swUserPreferenceStringValue_e`, consumed by `ISldWorks.GetUserPreferenceStringValue` — **value not re-queried; inherited from the baseline snapshot and never checked** |

`swDefaultTemplateDrawing` is the one entry here that is still an inherited
guess. It only affects the fallback path when the controlled VEEMAP template is
missing, so a wrong value degrades to "no template found" rather than a silently
wrong drawing — but it has not been verified.

## 2026-08-06 r16 ISelectData.View error 91 was VBA syntax, not the API

Open since r5 and recorded here twice as an unexplained API refusal. It was
neither unexplained nor the API.

`ISelectData.View` is documented as `View {get; set;}`, *"Gets or sets the
drawing view that contains the selected object"*, with **no error condition**
(MCP, 2026-08-06). A refusal with both operands proved non-Nothing, after a
successful `ActivateView`, was therefore not SOLIDWORKS declining.

The property is exposed as `propertyput`, not `propertyputref`. VBA requires a
plain assignment for that; `Set` against it raises 91.

```vba
swSelData.View = swView       ' works
Set swSelData.View = swView   ' raises 91
```

r16 tries the let-assignment first and keeps `Set` as a fallback, recording
which the build accepted. The run reported **`Selection scope:
ScopedToView(Let)`** — first scoped selection since r5.

Two lessons worth keeping:

- Every selection this macro made between r5 and r15 was **unscoped**. It
  happened to work because the owning view was activated first, but it was not
  the pattern `docs/CODESTACK_DRAWING_API_COVERAGE.md` rows 17 and 31
  prescribe, and it was carried as a known-unresolved defect for eleven
  revisions.
- An error raised at a COM boundary is not evidence the callee refused. Check
  the VBA binding before recording an API limitation.

**This did not fix the dangling stations.** X=55.00 and X=135.00 still dangle
with selection correctly scoped, which rules scoping out as their cause
alongside entity sharing.

### Next hypothesis for the dangling stations, with a zero-code test

Both failing stations sit on features drawn in **hidden-line font** on the
sheet. Runs to date used `UseHLR = False`, which passes `swHIDDEN_GREYED`
(Hidden Lines Visible), so obscured features are drawn and are returned by
`GetVisibleEntities2` — the Help defines that call as excluding only entities
*completely* obscured.

If an ordinate cannot hold an attachment to an edge that is not truly visible,
running with `UseHLR = True` should drop the X station count from 9 to 7 and
produce zero danglers. That test needs no code change — only the form option —
and it is now meaningful for the first time, because before r16
`swDisplayMode_HiddenLinesRemoved` was 3 (`swSHADED`) and the HLR option never
actually selected HLR.

## 2026-08-06 r16 HLR test: hidden-line edges cannot hold an ordinate

The dangling-station root cause, open since r10 and wrong three times before
this. Confirmed by a zero-code experiment: the same deployed r16 binary, run
twice, differing only in the form's HLR setting.

| | HLV (`swHIDDEN_GREYED`) | HLR (`swHIDDEN`) |
|---|---|---|
| `GetVisibleEntities2` edges | 64 (circular 35, linear 29) | **39 (circular 22, linear 17)** |
| X stations | 9 | **5** |
| Y stations | 7 | **5** |
| Dangling ordinates | 2 | **0** |

`IView.GetVisibleEntities2` returns edges that the view draws in hidden-line
font when the display mode is Hidden Lines Visible. The Help defines the call
as excluding entities *completely* obscured, and an edge rendered as a dashed
hidden line is not completely obscured — so it is returned, it selects
successfully, `MultiSelect2` counts it, `AddOrdinateDimension` returns
`swCreateOrdDimErr_Success`, and the resulting ordinate is **dangling**. Every
gate the engine had said yes.

Under HLR those edges are not in the view at all, so they never become
stations and no dangling ordinate is created.

**Consequences.**

- `GetVisibleEntities2` visibility is not the same question as "can this
  entity carry a dimension". Do not treat the former as evidence of the
  latter.
- The three earlier diagnoses were all wrong and all plausible: shared
  entities (r11), select/delete refusal (r14), unscoped selection (r16). Each
  was eliminated by a measurement, none by reasoning.
- The station counts under HLR match the reference drawing exactly — five per
  axis — and the cross-axis chain reads `36, 15, 0, 15, 35` against the
  reference's `36, 15, 0, 15, 36`.

The post-rebuild dangling prune stays. It is now a safety net rather than the
mitigation, and it correctly reports zero under HLR.

## 2026-08-06 r18/r19 ICurve.IsCircle is true for arcs; the datum must be an edge

`ICurve.IsCircle` Remarks: *"Use IEdge::GetCurveParams2 or
IEdge::IGetCurveParams2 to determine if a circular edge is a complete circle
or an arc."* The trunk never made that distinction, so P-0251's rounded end
was treated as a hole and its arc **centre** became the long-axis datum - the
chain started 60 mm inside the part.

`IsFullCircle` now distinguishes them from `IEdge.GetStartVertex` /
`GetEndVertex`, whose Remarks state the two "return distinct vertices, unless
the edge is closed". It fails closed to "full circle": misreading an arc as a
hole is the defect being prevented, misreading a hole as an arc only loses a
station.

### r18 excluded arcs, and that was an over-correction

| | X chain | Y chain |
|---|---|---|
| r17 (arcs kept, no edge rule) | 0, 70, 110, 150, 159 | 36, 15, 0, 15, 35 |
| r18 (arcs excluded + edge rule) | **89, 49, 9, 0** | **21, 0, 30, 50** broken |
| r19 (arcs kept + edge rule) | **159, 89, 49, 9, 0** | **36, 15, 0, 15, 35** |

The rounded end's arc centre lies exactly on the part's axis of symmetry and
was the only entity on it. Excluding arcs removed the Y centreline datum and
left it 14.5 mm off centre.

**An arc's centre is a true, dimensionable coordinate** - it is what an
ordinate attached to that edge reads, and for a rounded end it is the axis of
symmetry. The r17 defect was not that the centre is wrong, but that it was
mistaken for the part's *extreme*. Requiring an end datum to be a straight
edge fixes that on its own.

`ResolveOneDatum` therefore restricts an **end** datum to straight-edge
candidates when any exist, and leaves a **centreline** datum unrestricted.

### Result against the reference drawing

| | trunk r19 | P-0251-14A-001 |
|---|---|---|
| Long axis | 159, 89, 49, 9, 0 | 160, 90, 50, 10, 0 |
| Cross axis | 36, 15, 0, 15, 35 | 36, 15, 0, 15, 36 |

Five stations per axis on both chains, zero dangling ordinates.

**Unexplained, and not to be assumed:** every long-axis station is exactly
1 mm below its reference counterpart, and the cross-axis chain matches on one
side (36) and is 1 mm short on the other (35). One candidate is the drawing's
own general note - *"All corners are chamfered 0.5 x 45 deg"* - which would put
the selected straight edge on the chamfer rather than the true face extreme.
That has **not** been tested.

## 2026-08-06 r20 outer-edge rule closes the 1 mm systematic offset

Not an API finding. A **drawing convention**, stated by the user: a dimension
always goes to the OUTER edge.

A chamfered corner presents two parallel axis-parallel straight edges - the
true outer extreme and the chamfer's inner boundary. Both are legitimate
stations, both fall well inside `COORD_DEDUP_TOL_M`, so they merge into one
station. Which entity survived was decided by array order out of
`GetVisibleEntities2`, not geometry, and it was landing on the inner edge.

`PreferOuterCandidates` promotes each station's retained runner-up wherever it
lies further from the axis midpoint. It runs before datum resolution, because
the datum is chosen from these coordinates. Holes are unaffected: concentric
circles of one hole share a centre exactly.

**Result: both chains match the reference exactly.**

| | trunk r20 | P-0251-14A-001 |
|---|---|---|
| Long axis | 160, 90, 50, 10, 0 | 160, 90, 50, 10, 0 |
| Cross axis | 36, 15, 0, 15, 36 | 36, 15, 0, 15, 36 |

`Outer-edge promotions: 2`, one per axis, and the centreline datum's
`offsetFromTarget` went from 0.50 mm to 0.00 mm.

My arithmetic before the run predicted a 0.5 mm chamfer could only account for
half the discrepancy. It accounted for all of it - both the datum edge and the
opposite extreme moved, and the cross-axis midpoint moved with them. Recorded
because the reasoning was wrong and the run was right.

## IView display-mode members (r23, MCP-checked 2026-08-06)

Used by `Module2_DrawingPipeline.ForceHlrForHarvest` / `RestoreDisplayMode` to
make HLR a precondition of the ordinate harvest rather than a form checkbox.

| Member | Contract | Note |
|---|---|---|
| `IView.GetDisplayMode2()` | returns `Int32`, a `swDisplayMode_e` | not deprecated. `GetDisplayMode` (no suffix) is obsolete and its documented return value is literally "Unknown" - do not use it |
| `IView.SetDisplayMode3(UseParent, Mode, Facetted, Edges)` | returns `Boolean`, True on success | marked obsolete, superseded by `SetDisplayMode4`. Still the member the trunk uses; migration not attempted |
| `IView.UpdateViewDisplayGeometry()` | returns void | Remarks name this exact case: after switching HLR/HLV it gives immediate access to the new geometry without waiting for Windows to repaint |

`SetDisplayMode3` argument choices in the trunk: `UseParent=False` so the view
keeps its own local setting, `Facetted=False` to keep precision quality. Per the
Remarks, faceted display is controlled by the `Facetted` argument, and passing
any `swFACETED_*` value in `Mode` is silently downgraded to its non-faceted
equivalent - which is how three wrong display-mode constants survived fifteen
runs before r16.

`swDisplayMode_e` values re-confirmed 2026-08-06: `swHIDDEN=2` is Hidden Lines
Removed (HLR), `swHIDDEN_GREYED=1` is Hidden Lines Visible (HLV), `swSHADED=3`.

**Verified 2026-08-06** (`macro_qa/20260806_152553`, operator unticked HLR):
an ordinate created under forced HLR survives the restore to HLV still
attached. The post-rebuild prune, which runs after `ForceRebuild3` when
`IAnnotation.IsDangling` is meaningful, reported 0 dangling. The same run
proves the forcing path itself: `Harvest display mode: HLR forced (was 1)`,
39 candidate edges rather than HLV's 64, and the X datum on the end face at
160 mm.

**Still not verified:** whether `SetDisplayMode3` succeeds on the restore call.
`RestoreDisplayMode` runs under `On Error Resume Next` and discards the return
value.

## CreateSectionViewAt5 with a multi-segment stepped cut (r24, live 2026-08-06)

`macro_qa/20260806_165529`. Confirms two things left open when the member was
first logged.

**"the section line or lines" means literally that.** Three separately
created `SketchLine` segments, each `Select4 True, Nothing` (additive, not
replacing prior selection), then one `CreateSectionViewAt5` call: returned a
non-Nothing `View`. A stepped/jogged section cut does not need
`swCreateSectionView_OffsetSection` - that flag's own description ("two lines
at an angle") is the revolved-section case, confirmed distinct by the enum
text alone, now also confirmed by exclusion: the trunk uses only
`swCreateSectionView_NotAligned` for the stepped cut and it worked.

**Drawing-view sketch geometry is stored at model scale, not sheet scale.**
CodeStack row 9 flagged this as unverified-on-SW2025. `ISketchLine.GetStartPoint2`
read back after each segment's creation matched the requested model-mm
coordinates exactly, with no 0.6667 view-scale factor applied: requested bore
leg at Y=0, row leg at Y=15mm, jog at X=-27mm; readback `-127.4,0`, `-27,0`,
`-27,15`. The view's scale is applied at display time, not to the sketch's
stored coordinates.

**Not verified:** whether the resulting section view's cutting plane visually
follows the intended path through the model, or merely accepted the sketch
without error. `CreateSectionViewAt5` returning non-Nothing is evidence the
call succeeded, not that the cut is geometrically correct. The operator
screenshot confirms the section now sits inside the sheet border, which the
r23 version did not.
