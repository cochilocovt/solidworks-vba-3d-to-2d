# Current Status

Date: 2026-07-26

## R18 model-first drawing-output repair

The guarded r17 deployment compiled, synchronized all 15 managed components,
and executed on authorized fixture P-0251. Its retained evidence is:

- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/QA_REPORT.txt`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/diagnostic-drawing.png`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/R17_VISUAL_AND_QA_DIAGNOSIS.md`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/R18_API_AND_REPAIR_CHECKPOINT.md`

R17 preserved the runtime-proven J-J transaction and imported nine display
dimensions into Drawing View1 plus one into Drawing View2. The section and
isometric immediate per-view arrays were zero. It created the isometric before
import, executed eight view moves, and produced a visibly coherent four-view
layout. The metric mass calculation also returned `1.296824 kg` and wrote the
drawing property `Mass=1.30`.

R17 still failed manufacturing acceptance:

- all 64 + 47 drawing-to-model correspondence calls returned `Nothing`;
- accepted candidates, canonical locations, and ordinate groups stayed zero;
- final QA falsely reported 39 display dimensions and 10 on the isometric
  because its `GetNext5` iteration crossed the drawing sheet;
- the visible mass stayed `1296.82`, proving its note did not use the corrected
  drawing property;
- invalid note extents generated repeated diagnostic layout failures; and
- the controlled sheet still has no structural `ITitleBlock`.

The managed source is now
`target-spec-hybrid-v2-2026-07-18-r18`. R18 starts with the model audit, walks
owned internal cylindrical faces and circular model edges, and maps each known
model edge into the target view through `IView.GetCorrespondingEntity`.
Selectable datum vertices use the same model-first direction. Exact per-view QA
uses `GetDisplayDimensions`. A visible mass note is changed to `$PRP:"Mass"`
only when a single linked mass candidate exists and link/rendered-value readback
passes. Diagnostic note skips are allowed only while controlled boundaries are
already unproved and explicitly record `acceptance=False`.

`AddHoleCallout2` was not added as a fallback because the documented API
requires a user confirmation dialog. Hole callouts remain on the verified model
annotation path; model marking or another approved non-modal strategy is needed
for missing grouped callouts.

The missing structural `ITitleBlock`, approved exact title links, controlled
regions, grouped callout completeness, and stepped-bore manufacturing definition
remain deliberately fail closed. R18 is E2/E3 until embedded compilation and an
authorized synchronized run prove the new mapping and drawing result. The
complete project-local suite passes 66 tests, all 15 managed components pass the
Windows-1252/CRLF/no-BOM/no-metadata gate, and the guarded read-only deployment
preflight resolves r18 with the bootstrap present.

## Earlier source-completion and deployment outcome

The coherent replacement VBA source is written under
`src/target-spec-hybrid-v2/`. The current source identity is
`target-spec-hybrid-v2-2026-07-18-r18`; r14 removed the disproved sheet-edit
preflight, verified active drawing-view context with a selection-assisted retry,
packed diagnostic views in two rows, and used a standard independent 1:2 scale
for the P-0251 orientation aid. R15 adds the first output-driven corrections
and r16-r18 continue the evidence-driven repairs summarized above.

A guarded deployment tool now exists under `tools/swp-deploy/`. Its read-only
preflight confirms that `Fable.swp` contains 19 components including the
deployment bootstrap, while exported production source is r14. The tool manages
15 replaceable modules/classes through a disposable candidate project, compile
probe, source readback, and atomic promotion. The first candidate import exposed
UTF-8 BOM bytes plus VBA export metadata as uncompilable visible code. All
deployable `.bas` and ordinary `.cls` inputs are now Windows-1252 without a BOM
or `Attribute` records; the bootstrap assigns imported component names itself.
The first metadata-free candidate revealed that `VBComponents.Import` recreated
ordinary handler `.cls` files as standard modules, making `WithEvents` illegal.
The bootstrap now creates manifest `StdModule` and `ClassModule` components with
explicit VBIDE type values and injects the cleaned source with
`CodeModule.AddFromString`.

On 2026-07-26, the active-object proxy returned from the Running Object Table
rejected the installed interop's `ISldWorks` IID with `E_NOINTERFACE` before
Module0 could execute. The compiled invoker now retains the early-bound path but
falls back to `IDispatch` with an explicit by-reference error argument. This
fallback is live-verified.

The next live attempt reached Module0 but failed with VBA error 748 because
`VBProject.SaveAs` is not valid for the SOLIDWORKS host-managed project type.
Module0 now invokes the VBE Save command for the already-named candidate, and
PowerShell creates the separate output copy afterward. Failure evidence now
records the current stage/component.

The guarded deployment completed successfully on 2026-07-26. The compile probe
reported success, all 15 managed components matched the source, and both
candidate and post-promotion verification read back revision
`target-spec-hybrid-v2-2026-07-18-r14`. The promoted `Fable.swp` SHA-256 equals
the verified candidate SHA-256. The previous macro, verified candidate, compile
result, and source-verification evidence are retained under
`test_assets/iteration_evidence/swp_deployment/20260726_100426/`.

The r4 package contains nine standard modules, three domain/evidence classes,
three handler classes, two UserForm code snapshots, the `ThisLibrary` host-code
snapshot, and an import guide. It remains separate from both the protected
baseline and `src/active-ordinate/active_ordinate.swp`.

`Fable.swp` is now synchronized with the r14 managed source. No authorized
model, protected baseline, or manual reference drawing was changed.

## Implemented in the r4 export

- strict authorization and fixed hybrid behavior for the three fixtures;
- fixture-locked view, section, detail, datum, scale, title, note, part-ID, and
  QA acceptance profiles;
- reference-led orthographic plans without unconditional rotation;
- one deterministic primary section: P-0251 J-J, no Base Plate section, and
  Pump Holder B-B;
- mandatory Pump Holder Details C and D from the exact `*Bottom` view at an
  independent 3:1 scale, with COM-identity, selected-profile, circular-geometry,
  source/detail ownership, scale-ratio, placement, and outline readback;
- fixed model-annotation import followed by evidence-backed ordinate fallback;
- component-qualified visible-entity enumeration, matched-face feature
  ownership, configuration proof, physical-instance identity, and conservative
  rejection of unproved circles;
- typed direction-specific datum proof and family/view/datum/direction-scoped
  coverage reconciliation;
- datum-first ordinate transactions with checked selection counts, decoded
  return codes, cleanup, and physical/projection evidence ledgers;
- controlled-sheet and actual-scale readback, standard-scale layout, linked
  title-property evidence, read-only final QA, final cleanup proof, and atomic
  evidence writing; and
- deterministic import guidance for form code snapshots and the special
  `ThisLibrary` host component.

## Verified offline evidence

- The complete project-local companion suite passes **49 tests**.
- Installed SOLIDWORKS 2025 interop file version `33.1.2.4` was reflected for
  the used interfaces and enums.
- Primary annotation, visible-entity, corresponding-entity, feature,
  configuration, transform, ordinate, sheet, title-block, property, section,
  and detail-view contracts were checked at E3.
- Detail-view reflection confirms `CreateDetailViewAt4`, exact enum values,
  `UseParentScale`, `ScaleRatio`, `IView.GetDetail`, `IDetailCircle` ownership,
  profile, style/display, and outline members.
- Reflection reconfirmed `swCreateOrdDimErr_GenFailure = 1` and
  `swCreateOrdDimErr_OrdFailure = 7`.
- Structural tests cover balanced procedures, continuation limits, component
  inventory, caller/signature contracts, and duplicate local/parameter names.
- The r4 checkpoint contains a 35-entry SHA-256 manifest and source archive at
  `test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/`.
  The older similarly named checkpoint without `_r4_` is historical r3/28-test
  evidence and must not be used to identify the current source.

The user has embedded and compiled the r4 source successfully in the
SOLIDWORKS VBA editor. Its first run reached macro preflight but performed zero
SOLIDWORKS mutations because the template constant omitted the `VEEMAP`
directory level. R5 corrected that path and reached sheet measurement. R6 fixed
the `ISheet.GetSize` width/height output binding. R7 attempted to normalize
`SheetFormatVisible`; R8 records that flag as a warning and relies on structural
title-block/margin/usable-area proof. R9 permits diagnostic continuation after
those checks fail; R10 also continues after failed scale setting using the
template's existing scale. The user's r10 run created the primary view and
visually showed the requested hidden-lines-visible mode, but `SetDisplayMode4`
returned `False`. R11 replaces that setter-only gate with setter-plus-readback
verification. The r11 run then created both orthographic views but stopped on a
setter-only rebuild check. R12 permits independent diagnostic stages to continue
and provides a clearly non-acceptance layout reserve when `ITitleBlock` is not
defined. Production acceptance remains blocked until the controlled sheet
contract is proven. The r12 run reached all independent stages and exposed one
shared view-context failure blocking section, model annotation, and ordinate
operations. R13's attempted sheet-edit normalization was disproved by its next
run: it stopped before sheet measurement or view creation. R14 removes that
regression and limits recovery to a named `DRAWINGVIEW` selection plus one
`ActivateView` retry, accepted only when `ActiveDrawingView` matches.

## Remaining gates

1. Deploy and compile r18, then run the synchronized macro on P-0251 and retain
   its complete runtime and drawing evidence (E5/E6).
2. Confirm nonzero model-to-view mapped edges, ten canonical P-0251 physical
   locations, and the required X/Y ordinate coverage.
3. Prove the model-first center and bottom-left datum entities are selectable in
   the intended drawing views and that all ordinate transactions clean up.
4. Confirm exact per-view dimension counts remain `9, 1, 0, 0` before new
   ordinates and the isometric stays undimensioned.
5. Confirm the unique mass-note link reads `$PRP:"Mass"` and visibly renders
   `1.30`.
6. Resolve the authoritative template, sheet-format, structural `ITitleBlock`,
   property-name, and title/note/part-ID cell contract (D-04).
7. Resolve grouped hole-callout and stepped-bore definition without modal
   automation or invented manufacturing intent.
8. Run only the three authorized fixtures and retain complete E6 evidence.
9. Compare every output with its manual reference and pass E7 manufacturing,
   coverage, and layout acceptance.

## Controlled-template status

The controlled drawing template exists at:

- `V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`

The r4 constant incorrectly omitted the `VEEMAP` directory level, which caused
the truthful fail-closed first-run result. R5 corrects the fixed path. The
template-linked sheet-format and title/property contract still require live
evidence from the next authorized run.

`GetValidDrawingTemplatePath` performs one fixed-path existence check; it is not
template discovery. Module7's current property/link/cell map is an unapproved
D-04 candidate, not evidence that the controlled title contract is settled.

## Collaboration operating mode

The user owns VBA import/editing, compilation, saving, and macro execution by
default. Codex owns offline source work, API validation, output diagnosis,
reference comparison, test design, and evidence review.

Before every distinct live SOLIDWORKS task, Codex must ask whether the user or
Codex will perform it.

## Exact handoff point

The coherent r18 managed source passes the focused source-contract gates but is
not yet embedded. The next task is to deploy r18 into `Fable.swp`, compile it,
and run the same
focused P-0251 settings. The user has already chosen to perform SOLIDWORKS runs
and share the resulting QA report and screenshots; no Computer Use is required.
