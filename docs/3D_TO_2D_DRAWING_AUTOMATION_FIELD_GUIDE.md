# 3D-to-2D SOLIDWORKS Drawing Automation Field Guide

Review basis: complete CodeStack drawing corpus review, official SOLIDWORKS 2025
API Help checks, existing installed-type-library evidence, current project
source, and retained runtime evidence.

This is a durable engineering guide for the `VBA 3D TO 2D` project. It is not a
claim that the current macro has passed runtime or manufacturing acceptance.

## Source authority

Use evidence in this order:

1. installed SOLIDWORKS 2025 behavior and type library;
2. official SOLIDWORKS 2025 API Help, including Remarks and return/error enums;
3. current embedded VBA source, then its synchronized export;
4. retained runtime, QA, screenshot, and reference-comparison evidence;
5. CodeStack examples as practical secondary guidance.

Never copy a numeric enum, numbered API method, internal command ID, or
coordinate assumption from CodeStack without checking the first two sources.

The exhaustive page ledger is
[`CODESTACK_DRAWING_API_COVERAGE.md`](CODESTACK_DRAWING_API_COVERAGE.md).

## What an end-to-end drawing generator must do

CodeStack provides isolated techniques. The production pipeline must integrate
them into a stateful, validated transaction:

| Stage | Required result | Principal API/context | Current project owner |
|---|---|---|---|
| 1. Source guard | Saved, authorized part and intended configuration are proved; model remains read-only in effect. | `IModelDoc2`, path/type/configuration checks | `Module1_Main`, `Module3_ModelAudit` |
| 2. Drawing creation | Controlled drawing template opens as a drawing document. | `ISldWorks.NewDocument`, `IDrawingDoc` | `Module2_DrawingPipeline` |
| 3. Sheet contract | Size, scale, format, zones, usable bounds, and structural title region are measured. | `ISheet.GetSize`, `GetProperties2`, title block/format members | `Module8_RuntimeSupport` |
| 4. Base views | Required orthographic views reference the correct model/configuration and have verified display/scale. | `CreateDrawViewFromModelView3`, `IView` properties/readback | `Module2_DrawingPipeline` |
| 5. Derived views | Required section/detail/isometric views are created from proved source geometry and ownership. | section/detail creation APIs, view sketches, selection context | `Module2_DrawingPipeline` |
| 6. Initial layout | Views fit the measured usable regions without breaking parent/scale relationships. | `IView.GetOutline`, `Position`, scale and parent flags | `Module9_LayoutEngine` |
| 7. Model items | Applicable model-marked dimensions, datums, tolerances, notes, and callouts are imported into the intended view only. | selected-view `InsertModelAnnotations4` | `Module4_ModelItemImporter` |
| 8. Missing locations | Only feature-proven, view-owned hole/feature locations absent from imported coverage become fallback candidates. | model audit, model/view correspondence, visible entities | `Module3`, `Module5` |
| 9. Dimensions | Datum-first ordinate groups are created, decoded, and cleaned up; duplicates and unsupported views are rejected. | view-scoped selection, `AddOrdinateDimension`, `SetPickMode` | `Module5_FallbackDimensionEngine` |
| 10. Metadata | Title properties, material/mass, notes, and part identification are linked and read back. | custom properties, notes, title block | `Module7_TitleBlockEngine` |
| 11. Final layout | Annotation-aware extents clear sheet/title/note boundaries and each other. | per-view outlines and annotation extents | `Module9_LayoutEngine` |
| 12. QA and handoff | View/dimension ownership, coverage, cleanup, layout, links, and manufacturing gaps are reported truthfully. | per-view collections, readbacks, atomic evidence | `Module6_QAEngine` |

API success at any one stage is not evidence that the drawing is correct.

## The four contexts that must not be mixed

### 1. Model context

Edges, faces, vertices, bodies, features, and model coordinates obtained from a
part or assembly belong to the model and its active/referenced configuration.
They are not automatically selectable drawing objects.

Examples:

- `PartDoc.GetEntityByName`
- feature/body/face traversal
- mass-property input bodies
- model bounding boxes

### 2. Component or assembly context

An entity from a component part may need conversion through
`Component2.GetCorresponding` before it represents that entity in the owning
assembly. CodeStack's named-edge example uses:

```text
part entity -> component correspondence -> drawing view selection
```

This extra step is not needed for a direct part view, but it becomes essential
when assembly support is added.

### 3. Drawing-view context

`IView.GetVisibleEntities2(component, type)` returns entities belonging to that
drawing view. CodeStack selects those entities directly through
`IEntity.Select4`; it uses `IView.SelectEntity` when starting from a model
entity.

Official SOLIDWORKS 2025 Help also provides
`IView.GetCorrespondingEntity(modelEntity)`, which returns the corresponding
view edge/face/vertex or `Nothing` when the entity is not represented in that
view.

The safe distinction is:

| Starting pointer | Conversion/selection route |
|---|---|
| Model edge/face/vertex | `IView.GetCorrespondingEntity`, or `IView.SelectEntity` when selection is the goal |
| Assembly component's part entity | Convert into component/assembly context, then map/select in the view |
| Entity returned by `GetVisibleEntities2` | `IEntity.Select4`, preferably with `ISelectData.View` |

Do not pass `Nothing` as the `GetVisibleEntities2` component merely because the
view references a part. The retained SW2025 project probe produced a type
mismatch for that shortcut and succeeded after using the `Component2` returned
by `GetVisibleComponents`.

### 4. Sheet and sketch context

The same `SketchManager.CreateLine` call can create view-owned or sheet-owned
geometry depending on the active context.

- Activate a named view and use its sketch transform for geometry that must move
  with that view.
- Activate `""` for sheet context when geometry must remain fixed on the sheet.
- Restore the intended sheet context after view-specific work.

Context is part of the API input even when it is not a formal method parameter.

## Hardened traversal patterns

### Traverse real views without treating the sheet as a model view

Official SOLIDWORKS Help states that `IDrawingDoc.GetFirstView` returns the sheet
pseudo-view; `IView.GetNextView` returns the first real view after it.

CodeStack frequently uses `IDrawingDoc.GetViews`, which returns one array per
sheet. In each sheet array, entry zero is the sheet pseudo-view and real views
start at entry one.

Production code must:

- check `IsEmpty`/`IsArray` before `LBound` or `UBound`;
- match the sheet pseudo-view name to the intended sheet;
- skip the pseudo-view for model-view operations;
- not assume `ISheet.GetViews` has identical membership; CodeStack warns that it
  can also expose View Palette content;
- use `GetName2`/`SetName2` where supported rather than the obsolete `Name`
  property; and
- keep section/detail view names and types stable for QA.

### Enumerate visible entities with their real component

Conceptual VBA pattern:

```vb
visibleComponents = swView.GetVisibleComponents
If IsEmpty(visibleComponents) Then
    ' Record an explicit no-component result.
Else
    For componentIndex = LBound(visibleComponents) To UBound(visibleComponents)
        Set viewComponent = visibleComponents(componentIndex)
        visibleEdges = swView.GetVisibleEntities2( _
            viewComponent, swViewEntityType_e.swViewEntityType_Edge)
        ' Guard Empty/array shape before iteration.
    Next componentIndex
End If
```

This is enumeration, not semantic qualification. A visible circular edge can be
a boss, counterbore ring, cosmetic circle, silhouette, recess, or unrelated
round feature. It becomes a hole candidate only after model feature/face
ownership, configuration, closed-circle, internal-cylinder, physical-instance,
and view-ownership proof.

## Coordinate and transform rules

### Stable facts

- SOLIDWORKS drawing/model length values are normally metres.
- `IView.GetOutline` returns `[xmin, ymin, xmax, ymax]` on the drawing page.
- `IView.Position` is the view geometric-center position relative to the sheet
  origin.
- `IView.ModelToViewTransform` maps model coordinates into drawing-view space.
- A view's sketch has its own `ModelToSketchTransform`.
- `ISheet.GetProperties2` returns eight values:
  paper size, template index, scale numerator, scale denominator,
  first-angle flag, width, height, and same-custom-property flag.

### Context-specific recipes

**View-owned sketch geometry**

1. Obtain the target view and `IView.GetSketch`.
2. Activate the view and prove `ActiveDrawingView`.
3. Convert the intended page/model point into the view-sketch coordinate system.
4. Create the sketch entity.
5. Verify the entity belongs to that view.

**Sheet-owned sketch geometry**

1. Start with model coordinates.
2. Compose the model-to-view transform with the required sheet-scale transform.
3. Activate sheet context using `ActivateView("")`.
4. Clear view selections.
5. Create and verify the sheet-owned entity.

**Dimension/annotation placement**

Do not automatically reuse the sheet-sketch scale workaround. CodeStack's
dimension examples transform model geometry with `ModelToViewTransform` and use
the resulting page position without an additional sheet-scale multiplication.
The exact API receiving the coordinates determines the contract.

### What must be probed, not assumed

- Whether a particular transform already includes view scale and translation.
- Whether input entity coordinates are part, component, view, or sheet values.
- Transform multiplication order.
- The exact coordinate contract of a section/detail sketch profile.
- How view rotation and non-sheet scale affect the result.

Record a known model point, transformed point, `GetOutline`, `Position`, scale,
and visible result in the same probe. A screenshot alone cannot prove the frame.

## View creation, scale, and layout

### Direct views versus predefined placeholders

CodeStack's `InsertModelInPredefinedView` is useful when a controlled template
already contains all required placeholders. The current project instead creates
fixture-specific views dynamically with `CreateDrawViewFromModelView3`, then
creates required section/detail/isometric views.

Predefined views do not remove the need to validate:

- referenced model and configuration;
- expected number and type of populated views;
- display mode and orientation;
- sheet/parent scale flags;
- view bounds and title-block clearance; and
- required derived views and dimensions.

### Base and derived view scale

CodeStack's size-based scaling page exposes several traps:

- its 2% view-border deduction is explicitly undocumented;
- `GetOutline` can change as annotations and sketches are added;
- overlapping map ranges can match more than one rule;
- assigning a scale to a derived view can disconnect it from its parent; and
- a last-match-wins loop can silently overwrite an earlier scale.

Project policy:

1. choose a supported standard scale from stable model/view geometry and
   measured usable sheet space;
2. set `UseSheetScale`, `UseParentScale`, or independent `ScaleRatio`
   intentionally for each view type;
3. read the value back;
4. preserve section/detail parent relationships unless the target specification
   explicitly requires an independent scale;
5. perform an initial layout before annotations; and
6. perform a final annotation-aware layout and QA pass afterward.

`GetOutline` is appropriate for collision checks, but it is not proof of exact
manufacturing geometry and may include an implementation-dependent view border.

## Model annotations and fallback dimensions

### Model-item import

Official SOLIDWORKS 2025 Help defines `InsertModelAnnotations4` as operating on
the currently selected drawing view when `AllViews=False`. Its return value is an
array of inserted annotations.

Required sequence:

1. activate the drawing document;
2. establish and verify the target view;
3. select/activate that view as required by the installed build;
4. clear stale selections;
5. call `InsertModelAnnotations4` with named/verified masks and flags;
6. validate the returned annotation array and per-view display-dimension delta;
7. record annotation types and ownership; and
8. clear selection and restore context.

Do not infer failure solely from one context-sensitive `SelectByID2=False`
result if active-view readback proves the intended view. Do not infer success
solely from a nonempty returned array.

### Model-context versus drawing-context dimension targets

CodeStack provides complementary examples:

- model-named edges → `IView.SelectEntity`;
- visible view edges → `IEntity.Select4`; and
- a view-scoped `ISelectData.View` for annotation/BOM selection.

For this project, candidate selection must prove:

- candidate and datum belong to the same intended view and part;
- candidate represents a supported feature location;
- datum is a real selectable view entity, not only a coordinate;
- existing imported coverage does not already define the direction/location;
- selection count and order are exactly as expected; and
- the isometric and unsupported derived views are excluded.

### Ordinate dimensions are a stateful transaction

Official SOLIDWORKS 2025 Help for
`IModelDocExtension.AddOrdinateDimension` requires the datum/base entity first
and the other group entities afterward. The method returns a
`swCreateOrdDimError_e` result, and selections made immediately afterward can
continue the same group until `IModelDoc2.SetPickMode` restores normal mode.

Every group therefore needs one cleanup path:

```text
activate view
clear stale selection
select datum first
append intended targets
validate exact selection count/order
call AddOrdinateDimension
decode the returned enum
verify created per-view dimensions
clear selection
SetPickMode
restore sheet/view context
```

Cleanup is mandatory on both success and error.

### Dimension placement is not dimension completeness

CodeStack's “longest edge” example offsets a dimension by 20% of edge length.
Its balloon example adds a fixed 10 mm diagonal offset. These demonstrate API
mechanics, not readable manufacturing layout.

Production placement must account for:

- projected direction and datum side;
- view scale and actual text height;
- existing dimensions, leaders, notes, and section labels;
- sheet/title boundaries;
- duplicate chains;
- group alignment and jogs; and
- designer/reference conventions.

## Properties, title block, material, and mass

### Property source

The recurring CodeStack pattern is:

1. use the sheet's declared `CustomPropertyView` when valid;
2. resolve that view's referenced document and configuration;
3. read configuration-specific property first;
4. fall back to file-level property only under an explicit policy; and
5. write a drawing-level property when title notes should link through the
   drawing.

Harden this by recording:

- source sheet/view/model/configuration;
- property name;
- raw and resolved values;
- API return code and cache mode;
- fallback reason; and
- link text plus rendered readback.

Do not copy CodeStack's old `Get3`, `Add2`, or `Set` calls without checking the
current SW2025 custom-property methods and result enums.

### Material and mass

CodeStack's mass example contains a durable warning: bodies obtained in drawing
context are not automatically valid input to a model mass-property object and
can yield zero mass. It maps them back to part-context bodies in
`view.ReferencedConfiguration`, computes mass, and restores the original
configuration.

For this project:

- prefer reading/calculating mass directly from the authorized source model in
  the intended configuration;
- if view-body filtering is needed, map bodies to model context explicitly;
- do not rely only on body-name equality;
- check configuration-switch returns;
- restore the original configuration in guaranteed cleanup; and
- never save the fixture model.

## Layers, tables, and export

### Layers

Annotations do not share one universal `Layer` property. Assign the layer on the
specific supported interface, commonly the annotation owned by a note or
display dimension.

Validate the layer exists before assignment, check results, and never hide the
operation behind `On Error Resume Next`.

### Tables and BOMs

BOM configuration does not automatically track a changed view configuration.
If BOM support is added, use one proved owner view/configuration per sheet and
update the BOM feature explicitly.

Revision and other table mutations need bounds checks, reverse-order deletion,
return-code checks, and rollback/recovery evidence.

### Export and QA

The dimension-export page suggests a useful evidence schema:

- stable dimension name;
- owning view/sheet;
- type;
- page X/Y and drawing zone;
- nominal value and unit;
- tolerance type and limits; and
- attachment/candidate identity for coverage.

Use per-view collections such as `IView.GetDisplayDimensions` for ownership
counts. Do not let `GetNext5` traversal cross into another view/sheet and then
attribute those dimensions to the starting view.

When exporting through a disposable drawing:

- use a unique temporary document and output path;
- close it in guaranteed cleanup;
- check SaveAs Boolean, error code, and warnings;
- add margins rather than fitting exactly to `GetOutline`;
- prevent silent overwrite; and
- preserve original drawing selection/active-document state.

## Patterns not to copy from the CodeStack examples

The corpus is educational, not production hardened. Do not copy:

- missing `Option Explicit`;
- `UBound` without an `IsEmpty`/array guard;
- direct `ActiveDoc` casts without document-type validation;
- numeric enums or UI command IDs from old pages;
- obsolete calls merely because they still compile;
- global `On Error Resume Next`;
- unchecked selection, activation, rebuild, property, rename, save, or layer
  returns;
- configuration changes without a guaranteed restore;
- suppress/unsuppress as an unguarded refresh trick;
- raw COM-object `Is` identity as the only deduplication key;
- first-component/first-face/first-edge heuristics;
- hard-coded annotation offsets without collision checks;
- exact-fit sheet resizing without margins;
- unescaped CSV; or
- abrupt `End` statements.

## Project-specific diagnostic implications

### Model-to-view correspondence failures

For a current run that accepts model hole features but produces zero mapped
locations, inspect in this order:

1. Does `view.ReferencedDocument` identify the exact source model?
2. Does `view.ReferencedConfiguration` match the configuration from which the
   entity was obtained?
3. Was the edge/vertex obtained from a body active in that configuration?
4. For an assembly, was the part entity converted through the correct component
   instance first?
5. Does `GetCorrespondingEntity` return `Nothing` for all entity types or only a
   particular edge?
6. As a narrow diagnostic, does `IView.SelectEntity(modelEntity, False)` succeed
   and does SelectionManager readback prove the intended drawing view/entity?
7. Independently, does `GetVisibleComponents` return a component and does
   `GetVisibleEntities2(component, edge)` enumerate the expected view edges?

These tests separate a correspondence/context failure from a semantic
hole-detection failure. They do not authorize accepting every visible circle.

### Section/detail profile coordinates

CodeStack distinguishes view-sketch and sheet-sketch transforms. Retain runtime
evidence for:

- source view;
- active view readback;
- profile point coordinates before/after transform;
- created sketch owner;
- selection count;
- created section/detail owner; and
- final outline/placement.

A successfully created section proves the installed sequence worked for that
case; it does not prove the coordinate helper is general across view rotation,
scale, or every fixture.

### Title property failures

When displayed title text is stale or wrong:

1. prove the sheet property view;
2. prove referenced model/configuration;
3. read the intended property with current API result codes;
4. write the drawing property;
5. rebuild;
6. read the note link expression; and
7. read the rendered note text and extent.

Writing a correct drawing property without proving the visible note's link is
not acceptance.

### Layout failures

CodeStack reinforces that bounds and scales are dynamic. Keep separate evidence
for:

- physical sheet and zone margins;
- structural title-block bounds;
- initial view geometry/outlines;
- post-annotation outlines;
- note and leader extents;
- final view positions/scales; and
- every collision or out-of-bounds rejection.

Do not derive controlled title bounds from a percentage reserve and then call
the result production-safe.

## Official SOLIDWORKS 2025 contracts confirmed for this guide

- [`IDrawingDoc`](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc.html)
  exposes drawing-specific operations.
- [`IDrawingDoc.GetFirstView`](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc~GetFirstView.html)
  returns the sheet pseudo-view first.
- [`IDrawingDoc.ActiveDrawingView`](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc~ActiveDrawingView.html)
  is `Nothing` when sheet context is active.
- [`IView.SelectEntity`](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IView~SelectEntity.html)
  selects a supplied entity in that view.
- [`IView.GetCorrespondingEntity`](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IView~GetCorrespondingEntity.html)
  maps a part/assembly edge, face, or vertex to a view entity or returns
  `Nothing`.
- [`IView.GetVisibleEntities`](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IView~GetVisibleEntities.html)
  is obsolete and superseded by `GetVisibleEntities2`.
- [`IDrawingDoc.InsertModelAnnotations4`](https://help.solidworks.com/2025/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.idrawingdoc~insertmodelannotations4.html)
  inserts into the currently selected view when `AllViews=False`.
- [`IModelDocExtension.AddOrdinateDimension`](https://help.solidworks.com/2025/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~AddOrdinateDimension.html)
  requires datum-first selection, returns `swCreateOrdDimError_e`, and requires
  `SetPickMode` when the group is complete.
- [`ISheet.GetProperties2`](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISheet~GetProperties2.html)
  returns the documented eight-value sheet-property array.

## Open SW2025 verification list

Before adapting the remaining CodeStack techniques, verify:

1. exact `ModelToViewTransform` composition for part and assembly drawing
   entities;
2. `GetSketch.ModelToSketchTransform` direction in a drawing view;
3. section/detail profile coordinate and selection contracts;
4. `DrawingDoc.GetViews` versus `Sheet.GetViews`, including View Palette items;
5. exact content and padding of `GetOutline` after annotations;
6. current preferred view, sheet, custom-property, mass-property, and SaveAs
   method versions;
7. parent/derived view scale behavior;
8. configuration/display-state mutation and dirty-state side effects;
9. layer, hatch, BOM, revision-table, and title-block member returns;
10. whether a supported location-label API replaces CodeStack's internal
    command injection; and
11. every enum and numeric flag against the installed 2025 library.

## Acceptance boundary

This guide improves implementation discipline. It cannot replace:

- compilation of the synchronized embedded VBA project;
- runs only against the three authorized fixtures;
- complete Immediate Window and QA evidence;
- comparison with the manual designer drawings; and
- manufacturing review for missing sizes, locations, datums, sections,
  tolerances, callouts, overlaps, and title information.
