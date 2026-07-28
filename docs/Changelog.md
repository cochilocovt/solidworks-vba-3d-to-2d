# Changelog

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
