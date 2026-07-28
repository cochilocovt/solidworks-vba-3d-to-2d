# CodeStack Drawing API Corpus Coverage

Review date: 2026-07-28

## Result

The CodeStack drawing corpus at
[`solidworks-api/document/drawing`](https://github.com/xarial/codestack/tree/master/solidworks-api/document/drawing)
was exhaustively inventoried from Git rather than from the website's generated
navigation.

- Source snapshot: `xarial/codestack`
  commit `0cde3849a184cdbbface61238ef431a6ebb9d530`
- Source commit date: 2026-05-07
- Article pages reviewed: **33 of 33**
- Adjacent examples reviewed: **35 of 35** VBA/C# files
- Derived public page URLs checked: **33 of 33 returned HTTP 200**
- Macro, model, protected baseline, and reference-drawing changes: **none**
- Live SOLIDWORKS or Computer Use activity: **none**

This document is the completeness ledger. The durable implementation guidance is
in [`3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md`](3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md).

## Evidence policy

CodeStack is a practical secondary source. Its examples are useful for learning
object relationships and call sequences, but they are not authoritative for
SOLIDWORKS 2025 signatures, enum values, side effects, or manufacturing
acceptance. Load-bearing behavior must still be checked against:

1. official SOLIDWORKS 2025 API Help;
2. the installed SOLIDWORKS 2025 type library;
3. a narrow authorized runtime probe when documentation is insufficient; and
4. the current macro's compile, runtime, QA, and visual evidence.

The review deliberately read the source snippets as well as the article prose.
Several important prerequisites and defects exist only in the code.

## Complete page ledger

| # | CodeStack page | Main lesson retained | Project use or caution |
|---:|---|---|---|
| 1 | [Drawing automation index](https://www.codestack.net/solidworks-api/document/drawing/) | `IDrawingDoc` is the drawing-specific entry point; document-common operations remain on `IModelDoc2`. | A taxonomy only. It does not define an end-to-end 3D-to-2D or manufacturing-acceptance workflow. |
| 2 | [Update BOM referenced configuration](https://www.codestack.net/solidworks-api/document/drawing/bom-tables-update-referenced-configuration/) | BOM configuration is independent of a drawing view's configuration; use an explicit sheet property/configuration owner. | Future BOM work only. Guard missing views/tables and verify `GetConfigurations`/`SetConfigurations`. |
| 3 | [Change face-hatch layer](https://www.codestack.net/solidworks-api/document/drawing/change-face-hatch-layer/) | `IView.GetFaceHatches` and `IFaceHatch.Layer` can standardize section hatches. | Do not copy suppress/unsuppress as an unchecked refresh transaction. Restore view state on every exit. |
| 4 | [Change selected objects' layers](https://www.codestack.net/solidworks-api/document/drawing/change-layer-selected-drawing-objects/) | Layer access is interface-specific: sketch segment, sketch point, note annotation, display-dimension annotation, and so on. | Useful for generated-content standards. The sample's global `On Error Resume Next` hides failures and is prohibited here. |
| 5 | [Clear revision table and add revision](https://www.codestack.net/solidworks-api/document/drawing/clear-revision-table-new-revision/) | Delete table items in reverse order and resolve revision IDs back to row indices before writing cells. | Peripheral and destructive. The sample supports only the first sheet and has no rollback. |
| 6 | [Copy associated drawing paths](https://www.codestack.net/solidworks-api/document/drawing/copy-drawing-paths/) | Dependency-based discovery is stronger than assuming matching filenames. | Useful for locating references; verify `GetDocumentDependencies2` flags/pair layout. The sample contains an undeclared-function-name typo. |
| 7 | [Copy view properties to drawing](https://www.codestack.net/solidworks-api/document/drawing/copy-view-properties/) | Read configuration-specific properties first, then file-level properties; drawing properties can be a title-block indirection layer. | Highly relevant to title metadata, but the sample can choose the wrong sheet/view and uses obsolete property calls without checking results. |
| 8 | [Draw border on a layer](https://www.codestack.net/solidworks-api/document/drawing/draw-border-on-layer/) | Sheet-sketch creation has its own coordinate/scale contract; `AddToDB` should be restored after batch sketch creation. | Diagnostic overlays only. A four-line sketch is not a native controlled border, and `AddToDB=True` must never leak after an error. |
| 9 | [Create drawing-view sketch geometry](https://www.codestack.net/solidworks-api/document/drawing/drawing-view-sketch/) | View-owned sketch entities move, scale, and rotate with the view; activate the view and transform into its sketch coordinates. | Directly relevant to section/detail profiles. Verify the exact `ModelToSketchTransform` direction in SW2025. |
| 10 | [Export drawing dimensions](https://www.codestack.net/solidworks-api/document/drawing/export-dimensions/) | A useful QA record includes dimension owner, type, position, zone, value, tolerance type, and tolerance limits. | Strong QA architecture. Do not copy unescaped CSV, unchecked units, or `GetNext5` traversal without proving view ownership. |
| 11 | [Export flat-pattern views](https://www.codestack.net/solidworks-api/document/drawing/export-sheet-metal-views/) | Copy a view to a disposable drawing, normalize it, strip unwanted objects, resize, export, and close. | Good export isolation pattern. Exact-fit `GetOutline` sizing has no safety margin and filenames can collide. |
| 12 | [Find a specific edge in a view](https://www.codestack.net/solidworks-api/document/drawing/find-specific-edge-in-drawing-view/) | Assembly entity conversion is a chain: part entity → component/assembly context → drawing-view selection. | Important context lesson. Named entities are brittle and the authorized fixture models must not be modified to add names. |
| 13 | [Get all visible components](https://www.codestack.net/solidworks-api/document/drawing/get-all-visible-components/) | `GetVisibleComponents` returns drawing-context leaf components; model-context ancestry may need reconstruction. | Current part fixtures only need the returned `Component2` for `GetVisibleEntities2`. The sample's hierarchy reconstruction has guards and identity defects. |
| 14 | [Get view bodies and materials](https://www.codestack.net/solidworks-api/document/drawing/get-view-bodies/) | `IView.Bodies` may not cover flat patterns; visible faces can identify a represented body. | Possible material validation. The first-face flat-pattern heuristic is not production proof. |
| 15 | [Get view-body mass](https://www.codestack.net/solidworks-api/document/drawing/get-view-bodies-mass/) | Drawing-context body pointers are not valid mass-property inputs; map to part-context bodies in the view configuration and restore prior configuration. | Directly useful for independent mass QA. Body-name matching and configuration switching need stronger proof and guaranteed restoration. |
| 16 | [Import and export layers](https://www.codestack.net/solidworks-api/document/drawing/import-export-layers/) | Layer definitions can be serialized and applied before generated objects are assigned to layers. | Useful for standards, but the sample is non-transactional, overwrites files, lacks `Option Explicit`, and trusts unvalidated enum numbers. |
| 17 | [Insert BOM balloons](https://www.codestack.net/solidworks-api/document/drawing/insert-bom-balloons/) | Enumerate visible components, pass each exact component to `GetVisibleEntities2`, scope `Select4` with `ISelectData.View`, then create an annotation. | The selection pattern is valuable. First-edge selection and a fixed 10 mm diagonal offset are not designer-quality placement. |
| 18 | [Insert location label](https://www.codestack.net/solidworks-api/document/drawing/insert-location-label/) | Demonstrates a Win32 command-injection fallback when a public API was believed absent. | Do not adopt unless no SW2025 public API exists and the installed command is explicitly proven. The hard-coded command ID and 32-bit declarations are unsafe. |
| 19 | [Insert predefined views](https://www.codestack.net/solidworks-api/document/drawing/insert-predefined-views/) | `InsertModelInPredefinedView` can populate controlled template placeholders. | A possible controlled-template strategy, not a replacement for dynamic view/section/detail generation or post-insertion QA. |
| 20 | [Insert a sheet](https://www.codestack.net/solidworks-api/document/drawing/insert-sheet/) | Decode sheet properties once, reproduce scale/size/projection/zones, and reorder using a complete sheet-name list. | `NewSheet4` copies properties, not a completed drawing sheet. Verify current members and handle unique names/empty arrays. |
| 21 | [Lock sheet-format editing](https://www.codestack.net/solidworks-api/document/drawing/lock-sheet-format/) | Application events can cancel a UI command while a retained `WithEvents` handler is alive. | Governance utility only. Hard-coded command IDs and plaintext VBA passwords are not security boundaries and can interfere with automation. |
| 22 | [Open associated drawings](https://www.codestack.net/solidworks-api/document/drawing/open-associated-drawing/) | Find drawings by dependencies, then use `DocumentSpecification` and `OpenDoc7`, optionally in detailing mode. | Useful evidence-discovery tool, not production generation. It changes active-document state and leaves drawings open. |
| 23 | [Open a view's referenced model](https://www.codestack.net/solidworks-api/document/drawing/open-referenced-model/) | A view owns a referenced document, configuration, and display state that can differ from the model's current state. | Useful diagnostics. Production code must capture, check, and restore configuration/display state and never save fixture changes. |
| 24 | [Propagate configurations to sheets](https://www.codestack.net/solidworks-api/document/drawing/propagate-configurations-sheets/) | Copy/paste a sheet, change view configurations, and preserve view center by comparing `GetOutline` before/after. | The center-preservation pattern is useful. The sample can create flat-pattern configurations, which violates this project's no-model-modification boundary. |
| 25 | [Rename flat-pattern views](https://www.codestack.net/solidworks-api/document/drawing/rename-sheet-metal-views/) | Resolve a flat-pattern body to a cut-list feature and use `SetName2`. | Future sheet-metal support only. The sample uses obsolete visible-entity access and contains an undeclared index defect. |
| 26 | [Rename sheets from custom properties](https://www.codestack.net/solidworks-api/document/drawing/rename-sheets-custom-properties-values/) | Resolve the sheet's custom-property view, read its configuration property, then fall back to file scope. | Relevant metadata pattern. Sanitize and uniquify names; inspect current property return codes and inheritance settings. |
| 27 | [Rename views after sheets](https://www.codestack.net/solidworks-api/document/drawing/rename-views-after-sheets/) | `DrawingDoc.GetViews` yields per-sheet arrays whose first entry is the sheet pseudo-view; use stable view types and `SetName2`. | Stable names improve logs and QA. Preserve system-managed section/detail identities and check collisions. |
| 28 | [Replace sheet formats](https://www.codestack.net/solidworks-api/document/drawing/replace-sheet-format/) | Apply ordered size/format matching rules, activate each sheet, replace/reload its format, then restore the original sheet. | Direct title-block relevance but high mutation risk. Verify note-preservation flags, target existence, and current numbered methods. |
| 29 | [Set projected/true dimension type](https://www.codestack.net/solidworks-api/document/drawing/set-view-dimension-type/) | `IView.ProjectedDimensions` makes the dimension interpretation explicit. | Use a view-aware policy; a blanket value can be wrong for auxiliary, rotated, or specialized views. |
| 30 | [Create sheet-context sketch geometry](https://www.codestack.net/solidworks-api/document/drawing/sheet-context-sketch/) | Compose model/view and sheet-scale transforms, activate sheet context with `ActivateView("")`, then create sheet-owned geometry. | Useful for fixed QA overlays. Do not use approximate model bounding boxes as manufacturing geometry. |
| 31 | [Dimension visible drawing entities](https://www.codestack.net/solidworks-api/document/drawing/view-dimension-drawing-entities/) | Entities returned by `GetVisibleEntities2` are drawing-context objects and can be selected with `IEntity.Select4`; placement uses projected geometry. | Directly relevant to fallback dimensions. Visibility alone does not establish hole semantics, and the sample has no collision/cleanup discipline. |
| 32 | [Dimension named model entities](https://www.codestack.net/solidworks-api/document/drawing/view-dimension-model-entities/) | A model-context entity can be selected in a view with `IView.SelectEntity`; transform geometry for a page-coordinate dimension location. | Useful correspondence pattern. Named finite edges do not generalize to circular holes or arbitrary supplied parts. |
| 33 | [Scale views from size](https://www.codestack.net/solidworks-api/document/drawing/views-size-based-scale/) | Use size-to-scale rules, preserve parent relationships, set `ScaleRatio`, and rebuild. | Treat mainly as a warning: `GetOutline` can include annotations; the 2% border deduction is undocumented; overlapping rules overwrite each other. |

## Corpus-wide findings

### What the corpus teaches well

- Drawing object hierarchy: document, sheet pseudo-view, real drawing view,
  referenced document/configuration, component, entity, annotation, and table.
- The difference between model-context and drawing-context entities.
- The need for a real `Component2` when calling `GetVisibleEntities2`.
- View-owned versus sheet-owned sketch geometry.
- Model/view/sheet coordinate transformations and metre-based drawing values.
- Configuration-specific custom-property lookup before file-level fallback.
- View iteration, sheet ownership, layer assignment, disposable export drawings,
  and state restoration as recurring concerns.

### What it does not provide

The corpus contains no single example that creates and accepts a complete
manufacturing drawing from a saved 3D part. In particular, this subtree does not
provide a complete, current sequence for:

- dynamic base, projected, isometric, section, and detail view creation;
- model-item import with `InsertModelAnnotations4`;
- datum-first ordinate groups with `AddOrdinateDimension`;
- semantic hole/feature qualification;
- structural title-block definition and linked-property verification;
- annotation-aware collision avoidance;
- reference-drawing comparison and manufacturing completeness; or
- truthful fail-closed QA.

Those gaps are covered by the project architecture, official SOLIDWORKS 2025
contracts, local type-library evidence, and runtime/visual acceptance—not by
extrapolating from unrelated CodeStack utilities.

### Production-readiness audit of the examples

Across the 35 adjacent source files:

- `Option Explicit`: **0 files**
- any `On Error` handling: **9 files**
- explicit `ClearSelection` call: **1 file**
- explicit `SetPickMode` call: **0 files**

The examples also commonly assume non-empty Variant arrays, omit active-document
type checks, ignore Boolean/object return values, use older numbered members,
hard-code UI command IDs or enum values, and fail to restore sheet, selection,
configuration, suppression, sketch, file-handle, or temporary-document state on
every exit.

The examples are therefore learning material, not drop-in project modules.

## Completeness reproduction

The reviewed page set can be reproduced from the CodeStack Git tree by selecting
every path that:

1. begins with `solidworks-api/document/drawing`; and
2. ends with `index.md`.

At the recorded commit this yields one root page and 32 child pages. Every child
directory in that subtree contains exactly one `index.md`; there are no
additional Markdown article files in the subtree.
