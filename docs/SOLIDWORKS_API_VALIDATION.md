# SOLIDWORKS API Validation of the Two Source Snapshots

## Scope and evidence

This report combines the repository analysis in `docs/ORDINATE_GAP_ANALYSIS.md` with API evidence returned by the configured `solidworks-api` MCP server.

Source abbreviations used throughout:

- **B**: `src/baseline-model-dims/`
- **A**: `src/active-ordinate/`

The MCP corpus identifies itself as **SOLIDWORKS 2026**, while this project targets **SOLIDWORKS 2025**. Values below are exact values returned by the MCP corpus; they are not inferred. Any item marked **verify in SW2025** must still be checked in the installed SOLIDWORKS 2025 Object Browser/type library before code is changed.

No VBA source file was edited, renamed, moved, or deleted during this analysis. The protected baseline remains unchanged.

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
| `swCreateOrdDimError_OrdFailure` | 1 | Ordinate failure |
| `swCreateOrdDimError_GenNoInternalDims` | 2 | No internal dimensions |
| `swCreateOrdDimError_GenBadSel` | 3 | Bad selection |
| `swCreateOrdDimError_GenNeedModelLoaded` | 4 | Model must be loaded |
| `swCreateOrdDimError_GenSamePartOnly` | 5 | Selections must be from same part |
| `swCreateOrdDimError_GenExtraSelection` | 6 | Extra selection |
| `swCreateOrdDimError_GenFailure` | 7 | General failure |
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
