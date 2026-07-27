# Ordinate Gap Analysis

## Scope and evidence

This is a static comparison of the current exported files in:

- `src/baseline-model-dims/` (abbreviated **B**)
- `src/active-ordinate/` (abbreviated **A**)

The analysis follows the current repository instructions and `docs/PROJECT_HANDOFF.MD`. No VBA source file was edited. No SOLIDWORKS session, local Object Browser, or type-library inspection was available during this static pass, so every API constant or behavior that needs local confirmation is explicitly marked **verify locally**. Numeric enum values are reported only where the source already declares them; this report does not infer which value is correct.

## Executive findings

1. The snapshots contain the same 13 component filenames, but `Module2_DrawingPipeline`, `Module4_ModelItemImporter`, `Module5_FallbackDimensionEngine`, `Module6_QAEngine`, and `UserForm1` have material behavioral differences.
2. The active form correctly enforces the required fixed hybrid workflow by setting both dimension flags to `True` (`A/UserForm1.frm:530-562`). The baseline form instead exposes mutually exclusive model/ordinate modes (`B/UserForm1.frm:109-131`, `B/UserForm1.frm:451-479`).
3. The baseline model-item importer uses a whole-drawing attempt followed by a per-view fallback (`B/Module4_ModelItemImporter.bas:16-114`). The active importer selects one anchor view, makes one insertion call, and has no fallback (`A/Module4_ModelItemImporter.bas:11-69`).
4. The two importers declare materially different annotation-mask constants (`B/Module4_ModelItemImporter.bas:4-14`; `A/Module4_ModelItemImporter.bas:4-9`). Their exact SOLIDWORKS 2025 meanings must be verified in the installed type library before any repair.
5. The active ordinate engine has useful selection, deduplication, return-code, and cleanup safeguards (`A/Module5_FallbackDimensionEngine.bas:28-255`, `A/Module5_FallbackDimensionEngine.bas:403-521`), but it still treats every visible circular curve as a candidate. It never consumes the feature evidence collected by `Module3_ModelAudit` (`A/Module2_DrawingPipeline.bas:35-36`; `A/Module5_FallbackDimensionEngine.bas:103-221`). Therefore it cannot establish that a candidate is a hole.
6. The selected datum option currently chooses the circular candidate nearest a target derived from the view outline; it does not create or select a distinct Bottom-Left, Center, or Top-Left datum entity (`A/Module5_FallbackDimensionEngine.bas:349-398`). Whether that is a valid interpretation of the intended datum must be resolved against the local API and the user's drawing convention.
7. No active code suppresses ordinate locations already represented by imported model dimensions, and deduplication is only within one view and one chain (`A/Module5_FallbackDimensionEngine.bas:294-313`, `A/Module5_FallbackDimensionEngine.bas:403-460`).

## 1. Complete file inventory

The inventory below includes every file found under both snapshots. Line and byte counts are from the current files. There are no `.frx` files in either snapshot; both forms build their controls dynamically.

| Component file | Type | B lines / bytes | A lines / bytes | Primary responsibility |
|---|---:|---:|---:|---|
| `BtnHandler.cls` | Class | 23 / 603 | 23 / 606 | Create/Cancel button relay (`B:11-22`; `A:11-22`) |
| `Module1_Main.bas` | Standard module | 192 / 5,395 | 200 / 5,418 | Public configuration types, globals, entry point, resets, template lookup (`B:4-192`; `A:4-200`) |
| `Module2_DrawingPipeline.bas` | Standard module | 286 / 9,311 | 399 / 10,948 | Drawing creation, views, dimension workflow, section, title, QA (`B:15-285`; `A:18-398`) |
| `Module3_ModelAudit.bas` | Standard module | 290 / 9,033 | 288 / 9,009 | Feature-tree audit and hole-like feature payload (`B:18-288`; `A:17-286`) |
| `Module4_ModelItemImporter.bas` | Standard module | 200 / 6,016 | 195 / 4,815 | Model-item import, dimension count, auto-arrangement (`B:16-199`; `A:11-194`) |
| `Module5_FallbackDimensionEngine.bas` | Standard module | 253 / 7,970 | 539 / 16,921 | Circular-candidate collection and ordinate-chain creation (`B:8-253`; `A:28-538`) |
| `Module6_QAEngine.bas` | Standard module | 58 / 2,263 | 60 / 2,032 | QA summary and dimension counts (`B:5-56`; `A:5-59`) |
| `Module7_TitleBlockEngine.bas` | Standard module | 172 / 5,967 | 172 / 5,970 | Properties, title block, notes, barcode (`B:6-171`; `A:6-171`) |
| `SectionBtnHandler.cls` | Class | 24 / 617 | 24 / 620 | Add/Remove section relay (`B:12-23`; `A:12-23`) |
| `SectionDlgBtnHandler.cls` | Class | 24 / 603 | 24 / 606 | Section dialog OK/Cancel relay (`B:12-23`; `A:12-23`) |
| `ThisLibrary.cls` | Document class | 8 / 304 | 8 / 307 | Empty SOLIDWORKS macro-project host class (`B:1-8`; `A:1-8`) |
| `UserForm1.frm` | UserForm | 523 / 18,939 | 632 / 19,036 | Dynamic main configuration form (`B:11-523`; `A:11-631`) |
| `UserFormSection.frm` | UserForm | 126 / 3,508 | 126 / 3,511 | Section label/direction dialog (`B:11-126`; `A:11-126`) |

The three-byte size increases in otherwise equivalent active files are consistent with the byte-order marker visible at the start of the active exports; they are not behavioral VBA changes.

## 2. Module-by-module difference table

| Component | Material differences | Assessment |
|---|---|---|
| `BtnHandler.cls` | No code-flow or public-member difference (`B:11-22`; `A:11-22`). | Semantically equivalent. |
| `Module1_Main.bas` | Types and public signatures are unchanged (`B:4-42`, `B:52-192`; `A:4-42`, `A:53-200`). Default `UseOrdinateDims` changes from `False` to `True`, and default datum from `Center` to `Bottom-Left` (`B:117-121`; `A:122-126`). | Active defaults support the fixed hybrid workflow; form submission is still the authoritative setter. |
| `Module2_DrawingPipeline.bas` | Active calls `AddOrdinateDimensionsToAllViews` directly (`A:68-73`) instead of the baseline-local `AddFallbackOrdinateDimensions` wrapper (`B:58-60`, `B:80-95`). Active guards model import on a non-`Nothing` front view (`A:58-66`), honors `GenerateQAReport` (`A:91-99`), adds sheet/view validation, and creates section sketch geometry from `View.GetOutline` rather than the model bounding box (`B:209-267`; `A:269-379`). | Dimension sequencing remains import → ordinate → arrange → title → QA. The active section-coordinate change is outside the ordinate repair, but is safer in concept and must not be mistaken for multi-section support. |
| `Module3_ModelAudit.bas` | Active changes `If swArc.IsCircle Then` to `If swArc.IsCircle = 1 Then` (`B:190-207`; `A:189-206`) and makes only case/formatting changes elsewhere (`B:265-288`; `A:264-286`). | The explicit comparison is risky until the local property return type is verified. Currently this audit feeds the form count and QA, not ordinate filtering. |
| `Module4_ModelItemImporter.bas` | Baseline declares a broad set of import-mask constants and combines dimensions, GTols, marked dimensions, Hole Wizard profile/location dimensions, and callouts (`B:4-14`, `B:175-182`). It attempts whole-drawing insertion and falls back to selected per-view insertion (`B:16-114`). Active declares only two different mask constants, conditionally adds callouts, selects one anchor view, and calls insertion once (`A:4-9`, `A:11-69`, `A:170-177`). | This is the primary regression surface for model-dimension import. Constant names/values and argument semantics require local verification before patching. |
| `Module5_FallbackDimensionEngine.bas` | Active adds view activation, object validation, `GetSwApp`, diagnostics, projected-centre deduplication with smallest concentric diameter retention, all-view traversal, selection-count checking, and return-code reporting (`A:28-289`, `A:294-521`). Baseline also deduplicates centres and chain coordinates but has less validation (`B:8-92`, `B:151-253`) and includes an unused public hole-callout fallback (`B:94-149`). | Active is structurally more defensive, but still uses visible circular geometry without model/feature proof, uses name-based isometric exclusion, and does not prevent model/ordinate or cross-view duplication. |
| `Module6_QAEngine.bas` | Baseline exposes `CountAllViewDimensions` publicly and includes per-view names plus more warnings (`B:5-56`). Active makes the count function private, labels views by index, and reports only pass/fail (`A:5-59`). | Active pipeline correctly gates QA display, but diagnostics do not include ordinate candidates, selection failures, or API return codes. |
| `Module7_TitleBlockEngine.bas` | Only export/casing differences (`Extension` versus `extension`); procedure bodies and public signature are otherwise equivalent (`B:6-171`; `A:6-171`). | Semantically equivalent. |
| `SectionBtnHandler.cls` | No code-flow or public-member difference (`B:12-23`; `A:12-23`). | Semantically equivalent. |
| `SectionDlgBtnHandler.cls` | No code-flow or public-member difference (`B:12-23`; `A:12-23`). | Semantically equivalent. |
| `ThisLibrary.cls` | No executable code in either snapshot (`B:1-8`; `A:1-8`). | Semantically equivalent. |
| `UserForm1.frm` | Baseline offers mutually exclusive model/ordinate option buttons and persists `DimStyle` (`B:109-131`, `B:243-274`, `B:310-344`, `B:451-479`). Active removes the mode controls and explicitly sets both `UseModelDimensions` and `UseOrdinateDims` to `True` (`A:107-140`, `A:282-313`, `A:369-404`, `A:530-562`). Active defaults datum to Bottom-Left rather than Center. | Active matches the required fixed hybrid product behavior. `RunHybridStrategy`, preview, and custom-scale text remain stored but unused downstream. |
| `UserFormSection.frm` | Form-class GUID and member-name casing differ; public fields and procedures are unchanged (`B:20-126`; `A:20-126`). | Semantically equivalent. |

## 3. Call graphs

### 3.1 Model-dimension import

**Baseline**

```text
Module1_Main.main [B/Module1_Main.bas:52-100]
  -> Module2.CreateDrawing [B/Module2_DrawingPipeline.bas:15-22]
    -> Module2.RunDrawingPipeline [B/Module2_DrawingPipeline.bas:24-78]
      -> if GlobalConfig.UseModelDimensions [B:54-56]
        -> Module4.ImportModelItemsAcrossDrawing [B/Module4_ModelItemImporter.bas:16-62]
          -> GetModelItemMask [B:175-182]
          -> DrawingDoc.InsertModelAnnotations4 (whole drawing) [B:32-46]
          -> CountVariantItems [B:184-199]
          -> if returned count = 0 [B:48-50]
            -> ImportModelItemsPerView [B:64-114]
              -> activate/select each real drawing view [B:74-88]
              -> InsertModelAnnotations4 (selected view) [B:89-100]
```

**Active**

```text
Module1_Main.main [A/Module1_Main.bas:53-105]
  -> Module2.CreateDrawing [A/Module2_DrawingPipeline.bas:18-25]
    -> Module2.RunDrawingPipeline [A/Module2_DrawingPipeline.bas:27-107]
      -> if GlobalConfig.UseModelDimensions and front view exists [A:58-66]
        -> Module4.ImportModelItemsAcrossDrawing [A/Module4_ModelItemImporter.bas:11-69]
          -> optional GetFirstRealViewName [A:21-25, A:152-168]
          -> activate/select only the anchor view [A:27-43]
          -> GetModelItemMask [A:170-177]
          -> DrawingDoc.InsertModelAnnotations4 once [A:45-57]
          -> CountVariantItems [A:179-194]
```

The active graph has no equivalent of `ImportModelItemsPerView`.

### 3.2 Ordinate candidate collection

**Baseline**

```text
RunDrawingPipeline [B/Module2_DrawingPipeline.bas:58-60]
  -> AddFallbackOrdinateDimensions [B:80-95]
    -> skip view when its name contains ISO/ISOMETRIC [B:87-90, B:283-285]
    -> Module5.CreateHoleOrdinateDims [B/Module5_FallbackDimensionEngine.bas:8-92]
      -> View.GetVisibleEntities2(Nothing, 1) [B:24-29]
      -> Edge.GetCurve -> Curve.IsCircle [B:43-53]
      -> Curve.CircleParams -> MathPoint -> ModelToViewTransform [B:54-67]
      -> HasNearbyPoint projected-centre dedupe [B:69-77, B:151-159]
```

**Active**

```text
RunDrawingPipeline [A/Module2_DrawingPipeline.bas:68-73]
  -> Module5.AddOrdinateDimensionsToAllViews [A/Module5_FallbackDimensionEngine.bas:260-289]
    -> skip view when its name contains ISO/ISOMETRIC [A:276-282, A:534-538]
    -> CreateHoleOrdinateDims [A:28-255]
      -> activate view and create view-scoped SelectData [A:62-90]
      -> View.GetVisibleEntities2(Nothing, swViewEntityTypeEdge) [A:95-116]
      -> Edge.GetCurve -> Curve.IsCircle -> CircleParams [A:168-203]
      -> MathPoint.MultiplyTransform(ModelToViewTransform) [A:188-201]
      -> FindHoleCentre [A:206-215, A:294-313]
      -> AddHoleLocation or retain smaller concentric diameter [A:208-215, A:315-344]
```

`Module3_ModelAudit.GetAllHoleLikeFeatures` is called before drawing creation (`A/Module2_DrawingPipeline.bas:35-36`) but its result is passed only to QA (`A:91-99`). It is not in the active candidate-collection graph.

### 3.3 Ordinate dimension creation

**Baseline**

```text
CreateHoleOrdinateDims: require at least two unique centres [B/Module5_FallbackDimensionEngine.bas:81-87]
  -> ResolveDatumIndex [B:161-188]
  -> CreateOneOrdinateChain(horizontal X coordinates) [B:86, B:190-253]
  -> CreateOneOrdinateChain(vertical Y coordinates) [B:87, B:190-253]
      -> suppress repeated coordinate values [B:202-236]
      -> MultiSelect2; require exact selection count [B:238-239]
      -> AddOrdinateDimension [B:241-246]
      -> print non-success return code [B:248-250]
      -> clear selection [B:252]
```

**Active**

```text
CreateHoleOrdinateDims: require at least two unique centres [A/Module5_FallbackDimensionEngine.bas:223-242]
  -> ResolveDatumIndex [A:349-398]
  -> CreateOneOrdinateChain(horizontal X coordinates) [A:236-238, A:403-521]
  -> CreateOneOrdinateChain(vertical Y coordinates) [A:240-242, A:403-521]
      -> require at least two unique coordinates [A:422-467]
      -> MultiSelect2; require exact selection count [A:469-484]
      -> AddOrdinateDimension [A:486-503]
      -> print return code/non-success [A:505-511]
      -> clear selection [A:513-520]
```

### 3.4 Datum-origin resolution

**Baseline**

```text
UserForm1.cmbDatum [B/UserForm1.frm:119-130]
  -> GlobalConfig.DatumOrigin [B:451-466]
    -> AddFallbackOrdinateDimensions [B/Module2_DrawingPipeline.bas:80-95]
      -> CreateHoleOrdinateDims [B/Module5_FallbackDimensionEngine.bas:8-92]
        -> ResolveDatumIndex [B:161-188]
          -> derive target from GetOutline and View.Position [B:164-173]
          -> choose nearest circular candidate [B:175-187]
```

**Active**

```text
UserForm1.cmbDatum [A/UserForm1.frm:114-124]
  -> GlobalConfig.DatumOrigin [A:530-546]
    -> AddOrdinateDimensionsToAllViews [A/Module2_DrawingPipeline.bas:68-73]
      -> CreateHoleOrdinateDims [A/Module5_FallbackDimensionEngine.bas:260-282]
        -> ResolveDatumIndex [A:349-398]
          -> derive target from GetOutline and View.Position [A:357-379]
          -> choose nearest circular candidate [A:381-397]
```

Both implementations resolve the setting to an index in the circular-candidate array. Thus the first selected ordinate entity is a candidate circle nearest the requested target, not a separately proven view origin or datum point.

### 3.5 Dimension arrangement

**Baseline**

```text
UserForm1.chkAutoArrange -> GlobalConfig.AutoArrange [B/UserForm1.frm:451-471]
  -> RunDrawingPipeline after import and ordinate creation [B/Module2_DrawingPipeline.bas:51-64]
    -> Module4.AutoArrangeAllDrawingDimensions [B/Module4_ModelItemImporter.bas:128-137]
      -> AutoArrangeDimensionsInView for every real view [B:133-135, B:139-165]
        -> View.GetDisplayDimensions [B:145-147]
        -> select each annotation [B:149-159]
        -> ModelDocExtension.AlignDimensions [B:161]
```

**Active**

```text
UserForm1.chkAutoArrange -> GlobalConfig.AutoArrange [A/UserForm1.frm:530-551]
  -> RunDrawingPipeline after import and ordinate creation [A/Module2_DrawingPipeline.bas:55-79]
    -> Module4.AutoArrangeAllDrawingDimensions [A/Module4_ModelItemImporter.bas:88-107]
      -> AutoArrangeDimensionsInView for every real view [A:103-105, A:109-150]
        -> View.GetDisplayDimensions; require an array [A:120-125]
        -> select each annotation [A:126-140]
        -> ModelDocExtension.AlignDimensions [A:142-144]
```

## 4. Public declarations that differ

### Public types

No public type differs. Both snapshots declare the same:

- `Public Type SectionConfig` with `Label As String` and `Vertical As Boolean` (`B/Module1_Main.bas:4-7`; `A/Module1_Main.bas:4-7`).
- `Public Type DrawingConfig` with the same fields and field types (`B/Module1_Main.bas:9-42`; `A/Module1_Main.bas:9-42`).

### Public enums

Neither snapshot declares a public enum.

### Public constants

Neither snapshot declares a public constant. All source-defined constants are `Private`.

The critical **private** constant divergence is nevertheless part of the repair risk:

| Source declaration | B | A | Required treatment |
|---|---:|---:|---|
| Model-item source | `swImportModelItemsFromEntireModel = 0` (`B/Module4_ModelItemImporter.bas:4`) | Same source declaration (`A:4`) | Verify name/value and `InsertModelAnnotations4` source semantics locally. |
| Marked-for-drawing mask | `swInsertDimensionsMarkedForDrawing = 32768` (`B:8`) | `swInsertDimensionsMarkedForDrawing = 1` (`A:6`) | Contradictory source values; do not choose by memory. Verify installed SOLIDWORKS 2025 enum. |
| Hole Wizard/callout masks | Profile `65536`, location `131072`, callout `1048576` (`B:9-12`) | One `swInsertHoleWizardCallouts = 64` (`A:7`) | Verify exact available enum members, values, and whether profile/location/callout are separate bits. |
| Other baseline mask members | Dimensions `8`, GTols `32`; not-marked `524288` is declared but not combined (`B:6-12`, `B:175-182`) | Not declared | Verify which annotations the product should import and which enum bits represent them. |
| Ordinate direction/success | Horizontal `3`, vertical `2`, success `0` (`B/Module5_FallbackDimensionEngine.bas:4-6`) | Same numeric declarations under slightly different names (`A:16-19`) | Verify the local ordinate type and return-code enums before retaining these values. |

### Public procedure signatures

Only these public API differences exist:

| Difference | Baseline | Active | Caller/dependency impact |
|---|---|---|---|
| All-view ordinate entry point | Absent; traversal is private `Module2.AddFallbackOrdinateDimensions` (`B/Module2_DrawingPipeline.bas:80-95`). | `Public Sub AddOrdinateDimensionsToAllViews(ByRef swDrawModel As SldWorks.ModelDoc2, ByRef swDraw As SldWorks.DrawingDoc, ByVal datumType As String)` (`A/Module5_FallbackDimensionEngine.bas:260-263`). | Called by active `RunDrawingPipeline` (`A/Module2_DrawingPipeline.bas:68-73`). Its `swDrawModel` parameter is not read in the procedure body (`A/Module5_FallbackDimensionEngine.bas:265-289`). |
| Hole-callout fallback entry point | `Public Sub InsertHoleCalloutsForView(ByRef swDraw As SldWorks.DrawingDoc, ByRef swView As SldWorks.View)` (`B/Module5_FallbackDimensionEngine.bas:94-149`). | Absent. | No baseline caller was found. Active routes callouts through the model-item mask instead (`A/Module4_ModelItemImporter.bas:170-177`). |
| QA count visibility | `Public Function CountAllViewDimensions(ByRef swDraw As SldWorks.DrawingDoc) As Long` (`B/Module6_QAEngine.bas:36-45`). | Same effective signature is `Private` (`A/Module6_QAEngine.bas:45-59`). | No external caller was found in either snapshot; `BuildRunSummary` calls it internally. |

All other public procedure signatures are effectively identical despite formatting differences. Public module variables, form result fields, and handler fields are also unchanged.

## 5. Configuration-field lifecycle

All `DrawingConfig` fields are declared in the same positions in both snapshots (`B/Module1_Main.bas:9-42`; `A/Module1_Main.bas:9-42`). “Form set” includes assignments made by `UserForm1.DoCreate` and the form-triggered section synchronization. “Consumed” means a later workflow read affects execution or output.

| Field (declaration line in B/A) | Form set | B consumer | A consumer | Status |
|---|---|---|---|---|
| `CreateFront` (10/10) | `B/UserForm1.frm:454`; `A:533` | None; front is unconditional (`B/Module2_DrawingPipeline.bas:140-145`). | None; front is unconditional (`A/Module2_DrawingPipeline.bas:162-173`). | **Unused field** in both. |
| `CreateTop` (11/11) | `B:455`; `A:534` | `B/Module2_DrawingPipeline.bas:146-148` | `A:175-180` | Used. |
| `CreateBottom` (12/12) | `B:456`; `A:535` | `B:150-152` | `A:182-187` | Used. |
| `CreateRight` (13/13) | `B:457`; `A:536` | `B:154-156` | `A:189-194` | Used. |
| `CreateLeft` (14/14) | `B:458`; `A:537` | `B:158-160` | `A:196-201` | Used. |
| `CreateBack` (15/15) | `B:459`; `A:538` | `B:162-164` | `A:203-208` | Used. |
| `CreateIso` (16/16) | `B:460`; `A:539` | `B:166-168` | `A:210-216` | Used. |
| `CreateSection` (18/18) | Indirect: `CopySectionsToGlobalState` → `SyncSectionCompatibilityFields` (`B/UserForm1.frm:437-481`, `B/Module1_Main.bas:150-160`; `A/UserForm1.frm:516-564`, `A/Module1_Main.bas:155-165`) | `B/Module2_DrawingPipeline.bas:174-176` | `A:222-224` | Used for one primary section only. |
| `SectionCount` (19/19) | Indirect sync (`B/Module1_Main.bas:151`; `A:156`) | None; pipeline reads `GlobalSectionCount` instead (`B/Module2_DrawingPipeline.bas:273`). | None; pipeline reads `GlobalSectionCount` instead (`A:388`). | **Unused compatibility field** in both. |
| `SectionLabel` (20/20) | Indirect sync (`B/Module1_Main.bas:154-159`; `A:159-164`) | Fallback only (`B/Module2_DrawingPipeline.bas:269-280`) | Fallback only (`A:381-397`) | Consumed, but normal form path uses `GlobalSections(1)`. |
| `SectionVertical` (21/21) | Indirect sync (`B/Module1_Main.bas:154-159`; `A:159-164`) | Fallback only (`B/Module2_DrawingPipeline.bas:269-280`) | Fallback only (`A:381-397`) | Consumed, but normal form path uses `GlobalSections(1)`. |
| `UseModelDimensions` (23/23) | Radio value in B (`B/UserForm1.frm:462`); forced `True` in A (`A:541`) | `B/Module2_DrawingPipeline.bas:54-56` | `A:58-66` | Used; active implements fixed hybrid behavior. |
| `UseOrdinateDims` (24/24) | Radio value in B (`B/UserForm1.frm:463`); forced `True` in A (`A:542`) | `B/Module2_DrawingPipeline.bas:58-60` | `A:68-73` | Used; active implements fixed hybrid behavior. |
| `RunHybridStrategy` (25/25) | Forced `True` (`B/UserForm1.frm:464`; `A:543`) | None | None | **Unused field** in both. |
| `ImportHoleCallouts` (26/26) | Checkbox (`B/UserForm1.frm:465`; `A:545`) | None; baseline mask always includes its callout-related members (`B/Module4_ModelItemImporter.bas:175-182`). | `A/Module4_ModelItemImporter.bas:170-177` | **Unused in B**; nominally used in A, subject to enum verification. |
| `DatumOrigin` (27/27) | Combo value (`B/UserForm1.frm:466`; `A:546`) | Passed to M5 (`B/Module2_DrawingPipeline.bas:80-90`) | Passed to M5 (`A:68-73`) | Used, but current meaning is nearest candidate to target. |
| `PopulateTitle` (29/29) | Checkbox (`B/UserForm1.frm:468`; `A:548`) | `B/Module2_DrawingPipeline.bas:66-68` | `A:81-86` | Used. |
| `InsertBarcode` (30/30) | Checkbox (`B/UserForm1.frm:469`; `A:549`) | `B/Module7_TitleBlockEngine.bas:31-32` | `A:31-32` | Used when `PopulateTitle` is true. |
| `InsertNotes` (31/31) | Checkbox (`B/UserForm1.frm:470`; `A:550`) | `B/Module7_TitleBlockEngine.bas:31-32` | `A:31-32` | Used when `PopulateTitle` is true. |
| `AutoArrange` (32/32) | Checkbox (`B/UserForm1.frm:471`; `A:551`) | `B/Module2_DrawingPipeline.bas:62-64` | `A:75-79` | Used. |
| `SheetScale` (34/34) | Computed from form (`B/UserForm1.frm:473`; `A:553-554`) | `B/Module2_DrawingPipeline.bas:124-126` | `A:143-146` | Used. |
| `CustomScaleText` (35/35) | Raw form text (`B/UserForm1.frm:474`; `A:556`) | None | None | **Unused after `SheetScale` is resolved**. |
| `UseHLR` (36/36) | Checkbox (`B/UserForm1.frm:475`; `A:557`) | `B/Module2_DrawingPipeline.bas:128-133` | `A:148-154` | Used. |
| `ShowLayoutPreview` (37/37) | Checkbox (`B/UserForm1.frm:476`; `A:558`) | None | None | **Unused / not implemented** in both. |
| `TotalCostManual` (38/38) | Textbox (`B/UserForm1.frm:477`; `A:559`) | `B/Module7_TitleBlockEngine.bas:27-29` | `A:27-29` | Used when `PopulateTitle` is true. |
| `GenerateQAReport` (40/40) | Forced `True`; no user control (`B/UserForm1.frm:478`; `A:561`) | None; QA is unconditional (`B/Module2_DrawingPipeline.bas:70-74`). | `A/Module2_DrawingPipeline.bas:91-99` | **Unused in B**; used in A but not user-configurable. |
| `Cancelled` (41/41) | Create/Cancel/unload paths (`B/UserForm1.frm:479`, `B:516-523`; `A:562`, `A:619-630`) | Entry gate (`B/Module1_Main.bas:84-86`) | Entry gate (`A:88-90`) | Used before pipeline execution. |

`SectionConfig` has two fields and both are live:

| Field | Declared | Set by form | Consumed |
|---|---|---|---|
| `Label` | `B/Module1_Main.bas:4-7`; `A:4-7` | Local section records and global copy (`B/UserForm1.frm:376-400`, `B:437-448`; `A:444-474`, `A:516-527`) | Primary label resolution (`B/Module1_Main.bas:163-173`, `B/Module2_DrawingPipeline.bas:269-280`; `A/Module1_Main.bas:168-179`, `A/Module2_DrawingPipeline.bas:381-397`) |
| `Vertical` | `B/Module1_Main.bas:4-7`; `A:4-7` | Local section records and global copy (`B/UserForm1.frm:376-400`, `B:437-448`; `A:444-474`, `A:516-527`) | Primary cut direction (`B/Module2_DrawingPipeline.bas:269-280`; `A:381-397`) |

## 6. Why the baseline imports model dimensions correctly

The repository states that the baseline is the proven working reference. The static code provides four concrete reasons its import path is more resilient:

1. **The drawing views exist and are rebuilt before import.** `RunDrawingPipeline` creates the views, forces a rebuild, processes events, and waits before calling the importer (`B/Module2_DrawingPipeline.bas:32-55`).
2. **It begins with a drawing-wide insertion attempt in a cleared, sheet-level selection context.** `ImportModelItemsAcrossDrawing` clears selection, activates `""`, and calls `InsertModelAnnotations4` with the source mask and its whole-drawing boolean choice (`B/Module4_ModelItemImporter.bas:16-46`). This avoids making success depend solely on selecting one anchor view.
3. **It has a deterministic fallback.** When the returned item count is zero, it walks every real drawing view, activates and selects it, and repeats insertion in the view-scoped mode (`B/Module4_ModelItemImporter.bas:48-50`, `B:64-114`). The active snapshot removed this fallback.
4. **It supplies the broad mask used by the known working snapshot.** The mask includes general dimensions, GTols, marked-for-drawing dimensions, Hole Wizard profile/location dimensions, and a callout member (`B/Module4_ModelItemImporter.bas:6-12`, `B:175-182`). This is evidence of the working snapshot's behavior, not proof that the numeric declarations match the installed 2025 type library. The exact members and values still require local verification.
5. **It cleans selection state on normal and failure paths.** Selection and active-view state are cleared before and after import (`B/Module4_ModelItemImporter.bas:32-33`, `B:52-60`, `B:82-104`).

The baseline is not fully compliant with the current product requirements: its form can select model dimensions *or* ordinates (`B/UserForm1.frm:116-130`, `B:462-463`), and its Hole Wizard checkbox does not control the baseline mask (`B/UserForm1.frm:465`; `B/Module4_ModelItemImporter.bas:175-182`). Those issues do not negate the baseline's value as the working import-control-flow reference.

## 7. Why the active ordinate version may fail

The following failure modes are ranked by directness and expected impact.

### 1. Model-item import contract changed in two places at once

The active importer changed both the mask declarations and the selection/control flow (`A/Module4_ModelItemImporter.bas:4-69`). It selects one anchor view, requests insertion once, and returns zero on any selection/API error. It neither retries the sheet-level operation nor walks the remaining views. This makes a view-selection mismatch, an incorrect mask, or a zero-result first call terminal. The contradictory private constants cannot be adjudicated without the local 2025 type library.

### 2. Circular geometry is treated as an ordinate candidate without feature proof

The active collector accepts a visible edge whenever `Curve.IsCircle` is true (`A/Module5_FallbackDimensionEngine.bas:168-203`). Projected-centre deduplication then retains the smallest concentric diameter (`A:206-215`). This can reduce duplicate counterbore rings, but it does not distinguish a hole from a boss, fillet, counterbore step, or unrelated circular geometry. `Module3_ModelAudit` is not connected to this path (`A/Module2_DrawingPipeline.bas:35-36`, `A:91-99`).

### 3. Datum choice is a nearest-candidate heuristic, not a proven datum origin

`ResolveDatumIndex` converts the chosen label to a target near the view outline, then picks the nearest candidate centre (`A/Module5_FallbackDimensionEngine.bas:349-398`). `CreateOneOrdinateChain` places that candidate first in `selectedObjects` (`A:422-431`). Thus “Bottom-Left,” “Center,” and “Top-Left” currently mean “use the candidate circle nearest that target,” unless the API supplies a different interpretation of selection order. This can zero the wrong location and cannot represent an outline corner/centre datum when no candidate exists there.

### 4. Deduplication tolerances can merge distinct locations

Both projected-centre and repeated-coordinate tolerances are `0.0015` in API length units (`A/Module5_FallbackDimensionEngine.bas:21-23`). The module itself documents the API unit as metre (`A:21`), so the configured threshold is 1.5 mm. Two genuine centres or coordinates closer than or equal to that threshold collapse into one (`A:294-313`, `A:437-460`). The appropriate tolerances must be derived from drawing/model resolution and test geometry rather than assumed.

### 5. View eligibility is name-based and traversal is over-broad

The active wrapper processes every real view whose name does not contain `ISO` or `ISOMETRIC` (`A/Module5_FallbackDimensionEngine.bas:260-282`, `A:534-538`). A renamed isometric view can pass; a non-isometric view with those letters can be skipped; section or other derived views can be processed. There is no API-backed check that a view is a supported orthographic source, and no policy limits ordinates to the minimum useful views.

### 6. Deduplication is local, not drawing-wide or model-aware

`FindHoleCentre` deduplicates only within the current view (`A/Module5_FallbackDimensionEngine.bas:294-313`). `CreateOneOrdinateChain` suppresses repeated coordinates only within one horizontal or vertical chain (`A:403-460`). The engine does not inspect imported model dimensions before adding locations and does not prevent equivalent chains in several orthographic/section views. Therefore it cannot meet the requirement to avoid unnecessary model/ordinate duplication or duplicate chains across views.

### 7. Fixed placement and arrangement may push dimensions outside usable sheet space

Ordinate placement is hard-coded to a `0.015` offset from the view outline (`A/Module5_FallbackDimensionEngine.bas:486-503`). Auto-arrangement passes a fixed `0.06` spacing argument (`A/Module4_ModelItemImporter.bas:142-144`). Neither path checks sheet bounds, title-block occupancy, neighboring view outlines, or final annotation extents. The units and exact behavior of the arrangement spacing argument must be verified locally.

### 8. Feature audit can under-report simple cuts

The active audit compares `swArc.IsCircle = 1` (`A/Module3_ModelAudit.bas:189-206`) where the baseline tests the Boolean directly (`B/Module3_ModelAudit.bas:190-207`). If the installed type library exposes a Boolean whose true representation does not equal numeric `1`, simple circular cuts will be omitted from the audit. Today that affects the form count and QA; it becomes a candidate-filtering defect if Module 3 is later wired into Module 5.

### 9. Failures are printed but not returned to QA

The active engine correctly records selection counts and `AddOrdinateDimension` return codes in the Immediate Window (`A/Module5_FallbackDimensionEngine.bas:471-511`), but the public procedures return no status. `BuildRunSummary` sees only the total displayed-dimension count (`A/Module6_QAEngine.bas:5-59`), so model dimensions can produce a QA “PASS” even if every ordinate chain failed.

## 8. Ranked repair plan

No source change is made by this report. The plan deliberately excludes multi-section implementation.

| Rank | Smallest safe correction | Dependencies/callers affected | Required focused test | SOLIDWORKS 2025 API fact to verify locally |
|---:|---|---|---|---|
| 1 | **Restore the baseline import control flow in active Module 4:** whole-drawing attempt, returned-item count, then selected per-view fallback. Keep `ImportHoleCallouts` conditional, but rebuild the mask only from verified enum members. | Primary file: `A/Module4_ModelItemImporter.bas:4-69`, `A:170-194`. `Module2.RunDrawingPipeline` can retain the current public call signature (`A/Module2_DrawingPipeline.bas:58-66`). | Saved part with marked drawing dimensions; at least front/right views; prove nonzero import and visible dimensions. Repeat Hole Wizard checkbox off/on and record the exact annotation difference. Force/observe a zero-result first attempt and prove per-view fallback. | Exact `swInsertAnnotation_e` member names/values; `InsertModelAnnotations4` parameter order and meanings, especially source, all-views behavior, selected-view behavior, and return value. |
| 2 | **Normalize the Module 3 circle predicate before using its evidence:** use the property as the type library defines it, without a numeric truth comparison. | `A/Module3_ModelAudit.bas:173-212`; current callers are `A/Module2_DrawingPipeline.bas:35-36`, `A/UserForm1.frm:422-442`, and `A/Module6_QAEngine.bas:15`. | Feature tree containing one Hole Wizard feature, one simple circular cut, and one non-hole cut; compare audit count/labels with the feature tree. | Return type/true representation of `ISketchArc.IsCircle`; reliable feature type names for Hole Wizard and simple cuts in the installed build. |
| 3 | **Filter ordinate candidates with model/feature evidence.** Preserve projected-centre deduplication, but do not accept a circle solely because its curve is circular. Establish an evidence-backed mapping from audited feature locations or edge ownership to visible drawing edges. | `A/Module5_FallbackDimensionEngine.bas:103-221`; likely `A/Module3_ModelAudit.bas:17-286`; `A/Module2_DrawingPipeline.bas:35-36`, `A:68-73`. Public M5 signatures and every caller must be updated together if the hole collection is passed in. | Plate with two genuine holes plus a circular boss, fillet/arc, and counterbore. Only proven hole centres may enter ordinate chains. Save feature tree, Immediate Window candidate evidence, and drawing screenshot. | How a drawing-view `Edge` maps to its model entity/feature; valid use of `GetVisibleEntities2` for part views; `swViewEntityType_e` edge member; transform needed to compare audited model points with projected view centres. |
| 4 | **Define and implement the datum contract.** If the selected option means a view-outline or view-origin datum, create/select the API-supported datum entity rather than substituting the nearest hole. If the intended meaning is “nearest hole,” rename/document it and make the selection deterministic. | `A/UserForm1.frm:114-124`, `A:530-546`; `A/Module2_DrawingPipeline.bas:68-73`; `A/Module5_FallbackDimensionEngine.bas:349-398`, `A:403-503`. | Same two-hole plate with Bottom-Left, Center, and Top-Left. For each run, prove which entity is zero, verify coordinate values manually, and retain screenshots. | Selection ordering and datum requirements of `IModelDocExtension.AddOrdinateDimension`; permitted datum entity types; coordinate frames of `IView.ModelToViewTransform`, `IView.GetOutline`, and `IView.Position`; ordinate placement coordinate system. |
| 5 | **Separate and tighten tolerances; restrict eligible views.** Use a small projected-centre tolerance justified by model precision, a separately justified repeated-coordinate tolerance, and API-backed orthographic-view eligibility. Prevent redundant chains across views. | `A/Module5_FallbackDimensionEngine.bas:21-23`, `A:260-313`, `A:403-460`, `A:534-538`; possibly view-selection policy in `A/Module2_DrawingPipeline.bas:68-73`. | Two centres less than 1.5 mm apart; concentric counterbore rings; two holes sharing X; two sharing Y; renamed isometric view; one section view. Confirm no merged genuine holes, repeated coordinates, isometric chain, or redundant view chain. | Supported method/property for distinguishing standard orthographic, isometric, projected, detail, and section views; numeric/model resolution appropriate for point comparison. |
| 6 | **Add model/ordinate overlap suppression.** Before adding a fallback location, inspect existing imported display dimensions/annotations or record which feature/location dimensions were imported. Add an ordinate only where location information is still missing. | `A/Module4_ModelItemImporter.bas:11-69`; `A/Module5_FallbackDimensionEngine.bas:28-255`; coordination in `A/Module2_DrawingPipeline.bas:55-79`. | Part with marked size and location dimensions plus two holes. Confirm sizes/tolerances remain from model items and no equivalent ordinate location is duplicated. Then remove the marked location dimension and confirm the fallback ordinate appears. | API route from `DisplayDimension`/annotation back to referenced entities and whether ordinate dimensions can be distinguished reliably before creation. |
| 7 | **Make placement and auto-arrangement sheet-aware.** Bound offsets/spacing by sheet size, view outlines, neighboring views, and title-block exclusion; retain selection cleanup. | `A/Module5_FallbackDimensionEngine.bas:486-503`; `A/Module4_ModelItemImporter.bas:88-150`; sheet/view data in `A/Module2_DrawingPipeline.bas:130-216`. | Small and large templates, dense hole pattern, views near each sheet edge. Confirm no dimensions cross sheet borders, title block, or adjacent views. | Units and behavior of `ModelDocExtension.AlignDimensions`; annotation bounding-box/position API; whether `GetOutline` includes annotations or only view geometry. |
| 8 | **Return structured ordinate diagnostics to QA.** Count eligible candidates, created chains, selection failures, and nonzero creation codes per view; do not let imported model dimensions mask an ordinate failure. | Public result contract in `A/Module5_FallbackDimensionEngine.bas:28-289`, call site `A/Module2_DrawingPipeline.bas:68-99`, report `A/Module6_QAEngine.bas:5-59`. Update all callers together. | Deliberately produce: no candidates, one candidate, duplicate coordinates, failed selection, and nonzero creation return. QA must distinguish each outcome. | Exact `swCreateOrdDimError_e` members and values; whether `AddOrdinateDimension` can partially create annotations when returning an error. |

## Local verification checklist before any code patch

- In the SOLIDWORKS 2025 Object Browser/type library, capture the exact declaration of `DrawingDoc.InsertModelAnnotations4` and the relevant annotation/source enums.
- Capture the exact declaration of `View.GetVisibleEntities2`, its entity-type enum, and the allowed component argument for part drawing views.
- Capture `Curve.IsCircle`, `Curve.CircleParams`, and `SketchArc.IsCircle` declarations.
- Capture `ModelDocExtension.MultiSelect2`, `SelectData.View`, and the supported selection ordering for ordinate creation.
- Capture `ModelDocExtension.AddOrdinateDimension`, ordinate direction enum, return-code enum, supported datum entities, and coordinate units.
- Capture `View.ModelToViewTransform`, `View.GetOutline`, and `View.Position` coordinate-frame documentation.
- Capture `ModelDocExtension.AlignDimensions`, its alignment enum, and spacing units.
- Identify an API-backed method to classify standard orthographic versus isometric/derived/section views; do not rely only on view names.

Until those facts are recorded, the baseline's observed behavior is the comparison reference, but its numeric declarations must not be copied into the active snapshot as if they were independently verified SOLIDWORKS 2025 enum values.
