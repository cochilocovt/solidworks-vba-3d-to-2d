# Target-Spec Hybrid V2 — Complete Implementation Plan

> **Prepared:** 2026-07-18  
> **Status:** Offline source-completion handoff; WP1-WP14 are implemented in the r4 export, while E4-E7 remain  
> **Current source baseline:** `src/target-spec-hybrid-v2/`, identity `target-spec-hybrid-v2-2026-07-18-r4`  
> **Primary specification:** `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`  
> **Runtime:** SOLIDWORKS 2025, installed interop file version `33.1.2.4`  
> **Protected artifacts:** `src/baseline-model-dims/`, the three authorized `.SLDPRT` fixtures, and all manual reference drawings

## 1. Purpose and end goal

This is the implementation and acceptance plan for the new target-spec hybrid macro. It now records the completed r4 offline source phase and the remaining compile, live-contract, fixture, visual, and release gates.

The end product is a new SOLIDWORKS 2025 VBA macro that creates a manufacturing-useful drawing for each authorized fixture and approaches the corresponding manual reference in:

- semantic dimension and callout coverage;
- datum and feature-location correctness;
- view and section usefulness;
- title-block truthfulness;
- readability and layout;
- deterministic execution and cleanup; and
- retained, fail-closed QA evidence.

The production workflow is fixed:

1. Import applicable model-marked dimensions and annotations.
2. Prove visible feature ownership from model evidence.
3. Reconcile existing feature-location coverage.
4. Add only missing, family-scoped ordinate locations from approved selectable datums.
5. Create the required orthographic, section, detail, and isometric views without dimensioning unsupported views.
6. Use measured sheet, border, title, note, and annotation regions to produce a readable layout.
7. Populate controlled linked title data, general notes, and truthful part identification.
8. Produce a complete evidence record and fail whenever a mandatory requirement is missing or unproved.

The macro is not complete merely because it compiles, runs without an error, creates a nonzero number of dimensions, or passes static tests. Completion requires E5 compilation, E6 fixture execution, and E7 visual and semantic acceptance as defined below.

## 2. Non-negotiable scope and safety boundaries

Only these models may be opened, queried, or used for macro execution:

- `test_assets/models/P-0251-14A-001.SLDPRT`
- `test_assets/models/P-0252-01-001.SLDPRT`
- `test_assets/models/P-0252-01-013.SLDPRT`

The models may be inspected but must never be modified or saved. The macro may create disposable drawings from them.

The following boundaries remain in force:

- Never modify `src/baseline-model-dims/`.
- Do not overwrite or promote `src/active-ordinate/active_ordinate.swp` while the replacement is unaccepted.
- Do not edit the manual drawings under `test_assets/reference_drawings/`.
- Preserve failed outputs as regression evidence.
- Back up the canonical replacement `.swp` before each material embedded-project change.
- Keep the embedded VBA and exported source synchronized after every accepted compile iteration.
- Do not invent dimensions, fits, tolerances, GD&T, datum intent, material data, treatments, or manufacturing intent absent from an authoritative source.
- Do not describe asterisk-delimited part-number text as a machine-readable barcode until a symbology and scan test are approved.

## 3. Collaboration and operator model

| Activity | Default owner |
|---|---|
| Edit/import VBA, compile the project, save the `.swp`, and run the macro | User |
| Offline source audit, API verification, test design, output diagnosis, and reference comparison | Codex |
| Inspect or operate live SOLIDWORKS | User or Codex, explicitly chosen before each distinct live task |
| Modify or save an authorized part | Prohibited |

Before any distinct live SOLIDWORKS task, Codex must ask whether the user or Codex will perform that task. Offline repository analysis does not require that choice.

The intended handoff pattern is complete-module replacement, not line-by-line editing. When a public signature, class field, shared type, or cross-module data contract changes, all dependent components must be imported and compiled together.

## 4. Evidence levels and release vocabulary

| Level | Evidence | What it proves |
|---|---|---|
| E1 | Manual references and generated-output images | Visible intent and layout only |
| E2 | Exported source, structural checks, fake-COM tests | Offline source invariants only |
| E3 | SOLIDWORKS 2025 MCP, official documentation, and installed interop reflection | Signatures, enums, and documented semantics |
| E4 | Narrow live probe on a named build, fixture, and view | Only the probed binding and context |
| E5 | Full embedded VBA-project compile | Compile correctness of that embedded project |
| E6 | Embedded macro execution on an authorized fixture | Runtime behavior of that run and configuration |
| E7 | Visual and semantic comparison with the reference/specification | Manufacturing-drawing acceptance |

Use these status terms precisely:

- **Implemented in export:** source exists in the exported files.
- **Embedded:** that exact source exists in the canonical `.swp`.
- **Compiled:** the complete embedded project passes **Debug > Compile VBAProject**.
- **Runtime-proven:** the embedded macro ran with retained evidence on an authorized fixture.
- **Accepted:** runtime output passed the complete visual and semantic gate.

The macro report should separate automated runtime QA from release acceptance. An automated `PASS` does not become a production acceptance claim while `E7 Acceptance = PENDING`.

## 5. Current implementation point

### 5.1 What exists in r4

`src/target-spec-hybrid-v2/` currently contains:

- nine standard modules;
- `CHoleCandidate`, `CDatumProof`, and `CRunEvidence`;
- handler classes;
- `UserForm1` and `UserFormSection` code snapshots;
- a `ThisLibrary` code snapshot; and
- import guidance.

The r4 export is the coherent offline source-completion candidate. It includes
fixture-locked view/section/detail plans, fixed hybrid import plus ordinate
fallback, matched-face ownership, typed directional datums, attachment-backed
coverage reconciliation, recorded ordinate transactions, standard-scale layout,
controlled linked-title evidence, requirement-level read-only QA, final cleanup,
and atomic evidence output.

Pump Holder Details C and D are now mandatory. Each transaction uses the exact
`*Bottom` source, a single circular profile, `CreateDetailViewAt4`, independent
3:1 scale, and structural source/detail/profile/style/outline readback. The
profile coordinates and resulting `7 x 4`/`C0.5` legibility remain E4/E7 gates.

The project-local offline suite passes **49 tests**. Installed interop reflection
has verified the primary SOLIDWORKS 2025 members used by r4, including annotation
import, visible/corresponding entities, curve/face/configuration ownership,
transforms, ordinate creation, sheet/title/property APIs, sections, details,
scale inheritance/ratio, and detail-circle profile/outline readback.

### 5.2 What is not proved

There is no retained proof yet for:

- successful execution of the documented blank-form/host-component import
  workflow in the SOLIDWORKS VBA editor;
- a clean embedded VBA compile;
- an authoritative canonical replacement `.swp`;
- embedded/exported source identity after import;
- a controlled template and sheet-format contract;
- replacement-macro execution;
- fixture-specific ownership, datum, coverage, section, title, or collision behavior; or
- any fixture's visual and manufacturing acceptance.

The corrected configured path is `V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`. The user's first-run screenshot and a read-only path check confirmed that template; the r4 path omitted the `VEEMAP` directory level. R5 applies the correction, while linked sheet-format/title-contract proof remains a runtime gate.

### 5.3 Remaining blockers after offline source completion

| Priority | Gap | Principal gate |
|---:|---|---|
| P0 | Complete component set has not been imported and compiled in SOLIDWORKS | WP15 / E5 |
| P0 | Canonical replacement `.swp` name/path and promotion policy remain unapproved | D-00 / WP15 |
| P0 | Controlled `.drwdot`/`.slddrt`, property names, and title/note/ID cell map are absent | D-04 / WP16 |
| P0 | Section/detail coordinate transforms, selectable datums, selection ordering, returned arrays, and annotation attachment behavior need installed-build/fixture proof | WP16 / E4 |
| P0 | No authorized fixture has an r4 runtime evidence packet | WP17-WP18 / E6 |
| P0 | No r4 output has passed reference-led manufacturing, coverage, or layout review | WP17-WP19 / E7 |

Installed reflection reconfirmed `swCreateOrdDimErr_GenFailure = 1` and
`swCreateOrdDimErr_OrdFailure = 7`; the r4 decoder uses that mapping and the
source-contract suite protects it.

## 6. Target architecture and immutable invariants

### 6.1 Recommended final stage order

The production run should use this one-way pipeline:

```text
Authorize and initialize
  -> load fixture-specific requirement and view plan
  -> resolve controlled template, sheet format, and standard scale plan
  -> create drawing and dimension-bearing views
  -> create the required feature-led section/detail views
  -> import approved model annotations
  -> prove feature ownership and semantic instances
  -> prove approved selectable datums
  -> reconcile per-family X/Y location coverage
  -> create only missing ordinate groups on eligible orthographic views
  -> create undimensioned isometric orientation views
  -> arrange dimensions and populate linked title/notes/part identification
  -> solve layout, rebuild, and remeasure until stable
  -> perform read-only semantic and collision QA
  -> zoom for evidence
  -> perform and verify final cleanup as the last SOLIDWORKS mutation
  -> build and retain the final report/evidence
```

Creating isometric views after annotation import is the preferred way to guarantee they remain undimensioned while preserving the required whole-drawing import attempt. Section and detail views may contain authoritative model dimensions where needed, but they must never receive fallback ordinates. If live behavior requires a different creation order, the replacement must still prove these same view-policy outcomes.

### 6.2 Core invariants

1. The hybrid importer and ordinate fallback are always enabled together in production.
2. A visible circle is never sufficient hole evidence.
3. A datum is a separate proved object, never the candidate nearest a corner.
4. Size, thread, fit, tolerance, depth, and location coverage are separate facts.
5. Equal-coordinate suppression is scoped to an approved family/view/datum group.
6. Isometric views contain no dimensions; section/detail/isometric views contain no ordinates.
7. One deterministic primary section is supported; extra section UI entries must not be silently accepted.
8. Sheet-scale orthographic views remain on the actual sheet scale unless a deliberate, approved standard-scale policy says otherwise.
9. No view or annotation movement occurs after the final layout validation without invalidating and rerunning that validation.
10. Missing or uninspectable mandatory evidence is a failure, not a warning.
11. Final cleanup is the last state-changing SOLIDWORKS action.
12. A report-write failure makes the run unverifiable and therefore failed.

## 7. Critical path and gate order

| Phase | Main output | Gate |
|---|---|---|
| A | Decisions, canonical artifact, template/property contract | Inputs are explicit; no invented policy |
| B | Source-complete offline revision, tests, API ledger, source manifest | E2/E3 pass |
| C | User-owned clean import and full VBA compile | E5 pass and embedded/export identity |
| D | Narrow live binding/template probes where static authority is insufficient | Required E4 probes pass |
| E | Focused P-0251 integration and acceptance | P-0251 E6/E7 pass |
| F | Base Plate and Pump Holder integration | All three fixture E6/E7 gates pass |
| G | Fresh-session regression and release evidence | Release candidate accepted |

Template discovery can occur in parallel with offline source work, and the macro can be compiled before the template is available. No macro run should proceed until the controlled template gate passes.

## 8. Detailed implementation work packages

### WP0 — Freeze and preserve the r3 baseline

**Objective:** Preserve a truthful point of comparison before future implementation resumes.

Tasks:

1. Retain the existing r3 source and its SHA-256 manifest.
2. Record that the 28-test result is E2 only.
3. Preserve the existing broken `active_ordinate.swp` as historical runtime evidence.
4. Do not import r3 into the future canonical macro as the recommended next action.
5. Create a new revision only after a coherent source package and its tests pass.

**Exit:** r3 remains recoverable and all later evidence names its exact successor revision.

### WP1 — Designate the canonical replacement artifact and make the package importable

**Objective:** Remove import and compile ambiguity before the user's first full-project handoff.

Tasks:

1. User chooses the canonical replacement `.swp` path and whether it will eventually replace `active_ordinate.swp`.
2. Use a blank/recoverable SOLIDWORKS macro project that already contains its host `ThisLibrary` component.
3. Treat `ThisLibrary.cls` as code to paste into the existing host component, not as an imported class.
4. Produce genuine native form exports, including the required designer structure and any `.frx` sidecars, or document the supported fallback: create blank forms named `UserForm1` and `UserFormSection`, then paste their code.
5. Fix the form-result lifecycle: show a local instance, `Hide` on OK/Cancel, read values, and unload only after the caller has copied them.
6. Correct the section sketch object to the installed return interface: `ISketchManager.CreateLine` returns `SketchSegment`, whose `Select4` method is available.
7. Update the import order and component-type checklist.
8. Add package tests that reject non-native form exports and incorrect host-component instructions.

**Exit:** every supplied component has a deterministic import/paste path and no known interface-level compile blocker remains.

### WP2 — Expand the evidence and domain data contracts first

**Objective:** Give later stages enough structured state to prove requirements rather than infer them from totals.

`CHoleCandidate` or successor records must contain:

- fixture, configuration, component, owning feature, pattern/mirror instance, and stable semantic identity;
- matched drawing edge, corresponding model edge, matched cylindrical face, and ownership proof type;
- model centre and normalized model axis;
- view and sheet centre/axis coordinates;
- feature family, machining side, size, depth, end condition, thread, counterbore/countersink, and semantic-completeness status;
- selectable drawing entity and view role;
- rejection reason or acceptance proof;
- imported/created X/Y coverage state without mixing size coverage.

`CDatumProof` or its successor must contain:

- approved datum identity and scope: global, family, or view;
- component/entity identity;
- an explicitly typed selectable `SldWorks.Entity` interface;
- model, view, and sheet coordinates;
- source/proof type;
- selection return value and owning view;
- failure code such as `DatumNotSelectable`.

`CRunEvidence` should be expanded, or split into typed record classes, to retain:

- required-stage status: `NOT_STARTED`, `PROVED`, `FAILED`, or `NOT_APPLICABLE`;
- planned/created view records;
- unique physical feature instances and projected candidates separately;
- per-family semantic and directional coverage records;
- every datum/selection/ordinate transaction;
- section/detail evidence;
- requested and actual sheet/view scales;
- property source, target link, rendered value, and cell extent;
- layout items and collision findings;
- cleanup and evidence-write results; and
- automated QA result separately from E7 acceptance state.

Use canonical model-axis/location keys so the same physical hole projected into two views is counted once. A useful canonical line identity removes the axial component from the centre and normalizes axis sign, but the formula and tolerance must be tested before use.

**Exit:** no fixture requirement depends on an aggregate count that can double-count projections or mask a missing family.

### WP3 — Harden authorization, fixed workflow, template preflight, and state reset

**Objective:** Make every production run deterministic and fail before drawing creation when its environment is uncontrolled.

Tasks:

1. Require an active, saved part whose resolved path matches one of the three fixtures.
2. Record model path, active configuration, SOLIDWORKS build, macro revision, template, format, and settings before mutation.
3. Reset global collections, view registries, selections, and ordinate pick mode.
4. Enforce the fixed hybrid workflow at the pipeline boundary even if stale registry values or form fields say otherwise.
5. Define an acceptance profile that requires title block, notes, part identification, QA, importer, and ordinate fallback. If the user disables a mandatory stage for diagnosis, the run must be labeled diagnostic and cannot produce acceptance `PASS`.
6. Resolve the controlled template before `NewDocument`; never fall back to an uncontrolled default.
7. Make missing template name, format name, invisible format, zero margins, missing title block, or unusable bounds return failure immediately.
8. Reset per-fixture form defaults instead of inheriting incompatible view choices from the prior model.
9. Enforce zero or one section with the current data model. More than one configured entry must fail preflight or be blocked by a minimal approved UI correction.
10. Check every required rebuild, activation, selection, arrangement, sheet restore, and cleanup return value.

**Exit:** unauthorized, unsaved, untemplated, or unsupported configurations stop before content generation and cannot be reported as successful.

### WP4 — Implement deterministic fixture-aware view plans

**Objective:** Replace generic checkboxes and unconditional rotation with manufacturing-purpose roles.

Create an explicit plan per fixture. Each plan item must define:

- role and information purpose;
- SOLIDWORKS orientation name or other proved orientation source;
- view type and angle;
- referenced configuration;
- initial standard scale policy and reserved layout cell;
- model-annotation eligibility;
- ordinate eligibility; and
- expected dimensions or callouts.

Minimum view intent:

| Fixture | Required information roles |
|---|---|
| P-0251-14A-001 | Principal face/profile, narrow side view for side holes, J-J through the stepped functional bore, one clean undimensioned isometric |
| P-0252-01-001 | Principal face, thin edge view, at least one undimensioned isometric exposing the necessary machining side; no accepted-default section |
| P-0252-01-013 | Profile/lower/side views sufficient for the geometry, B-B through the threaded holes, at least one undimensioned isometric, details C/D if required by the approved legibility decision |

Remove the unconditional `PI / 2` rotation. The correct angle must be determined from the fixture plan and proved against the reference. Registry-selected redundant views must not be created merely because they were selected in a previous run.

For every created view, record the role, actual type, actual orientation, actual angle, configuration, position, scale, `GetOutline`, and eligibility flags. Independently prove that isometrics have zero dimensions and that no prohibited view has an ordinate.

**Exit:** the historical 90-degree mismatch is gone, every view adds required information, and the role registry agrees with the drawing.

### WP5 — Complete the model-annotation importer

**Objective:** Import authoritative marked annotations before generating fallback location dimensions.

Required behavior:

1. Use the verified `InsertModelAnnotations4` mask for marked dimensions, GTols, intended notes, Hole Wizard profile/location dimensions, and conditional hole callouts.
2. Activate and establish the required real-view context without relying on a blank view name.
3. Make the required whole-drawing attempt while only intended dimension-bearing views exist; create isometrics afterward.
4. If the result is zero, retry approved real views with an E4/E6-proven activation/selection context.
5. Record every returned annotation by type and owning view.
6. Rebuild, then count visible display dimensions per view.
7. Prove the off/on hole-callout mask and visible-output delta on every fixture.
8. Preserve fits, tolerances, GTols, datum symbols, threads, depths, and semantic size data.
9. Never import unmarked dimensions merely to increase counts.
10. Route every success, failure, and early exit through common selection and sheet-context cleanup.

Model dimensions may be appropriate on a section/detail view when they define the feature shown there. Isometric views must remain undimensioned. The importer and QA must distinguish this from the stricter rule that fallback ordinates are allowed only on approved orthographic views.

**Exit:** imported content is repeatable, view-owned, semantically preserved, and cleanly separated from fallback location coverage.

### WP6 — Complete model ownership and semantic feature proof

**Objective:** Accept only real, active, owned hole locations with sufficient manufacturing semantics.

For each eligible view and visible component:

1. Enumerate component-qualified visible edges with `GetVisibleEntities2`.
2. Map each drawing edge to its model edge using `GetCorrespondingEntity2`.
3. Require a complete circular body edge using `GetCurve` and `GetCurveParams3`; reject arcs and degenerate closures.
4. Inspect adjacent faces and identify the radius/axis-matching cylinder.
5. Prove that the matched face belongs to the accepted feature using pattern-seed evidence or a verified `IFeature.GetFaces` plus `IFace2.IsSame` membership path. Do not accept `Face.GetFeature` alone when ownership is ambiguous.
6. Check suppression in the drawing view's `ReferencedConfiguration` using the verified configuration-specific contract.
7. Accept `HoleWzd`, `AdvHoleWzd`, `SketchHole`, or a generic cut only when the matched face itself has independently proved internal cylindrical-void semantics.
8. Resolve pattern and mirror instance identity without counting the seed and projection multiple times.
9. Use typed feature data where available to capture diameter, thread, depth, end condition, counterbore/countersink, and machining side.
10. Transform centre and axis to view space and require an E4-calibrated normal-axis tolerance.
11. Merge concentric rings into one location while retaining the complete semantic stack.
12. Log bosses, fillets, external cylinders, arcs, suppressed features, unresolved owners, missing semantics, and ambiguous machining side.

Raw sketch-point counts remain inadmissible as physical hole-instance counts. Hole Wizard centre points may be used as a semantic cross-check; the visible owned-edge path establishes displayed instances.

**Exit:** every accepted candidate has a matched-face ownership chain, configuration proof, unique physical instance identity, projected location, and explicit semantic completeness state.

### WP7 — Define and prove coherent datum identities

**Objective:** Ensure every coordinate zero represents approved manufacturing intent.

Tasks:

1. Resolve D-02 and D-06 before final datum acceptance.
2. Define whether each origin is global, per family, or per view; do not call different physical vertices the same global datum.
3. Replace projected-extrema inference with an approved entity/feature definition.
4. For `Center`, require the real origin, centreline intersection, datum point, or another explicitly approved selectable centre entity.
5. For `Bottom-Left` and `Top-Left`, require an actual model/drawing corner or datum geometry tied to a coherent model-space identity.
6. Store the entity through `SldWorks.Entity` before calling `Select4`; do not late-bind `Select4` directly on `Edge` or `Vertex`.
7. Record stable entity context plus model, view, and sheet coordinates.
8. Record the selection return value and final owning view.
9. Return `DatumNotSelectable`, skip the group, and fail required QA when proof fails.
10. Preserve formal datum A independently from any ordinate zero.

**Exit:** each exposed datum choice either resolves to a coherent, selectable, retained identity or fails closed.

### WP8 — Replace coverage reconciliation with feature/family/direction records

**Objective:** Add only genuinely missing X and Y locations without suppressing size information.

Define a coverage key containing at least:

- fixture and configuration;
- component and semantic feature instance;
- view role;
- approved datum identity;
- family/group identity;
- X or Y direction;
- coordinate and documented tolerance; and
- source: imported linear, imported ordinate, fallback-created ordinate, or unresolved.

Implementation tasks:

1. Enumerate existing display dimensions and annotations by owning view.
2. Prove imported linear location coverage only when the same display dimension can be tied to both the approved datum and the candidate/feature location through an E4-proven attachment pattern.
3. Continue recognizing existing ordinate coverage, but retain its datum, direction, and group identity.
4. Keep size/thread/depth/fit/tolerance coverage separate from X/Y location coverage.
5. Treat failure to inspect attachments as an unknown mandatory state: stop that ordinate group instead of risking duplicates.
6. Build X and Y plans per semantic family/view/datum, not one global group for every hole in a view.
7. Suppress equal coordinates only within that approved group. Equal coordinates in different families must not hide either family's coverage.
8. Preserve X-only, Y-only, X/Y, and fully uncovered states explicitly.
9. Detect contradictory dimensions or incompatible tolerances as failures.

Required offline cases:

- imported X only;
- imported Y only;
- imported X and Y;
- no imported location;
- imported size but no location;
- equal coordinate in one family;
- equal coordinate across two families;
- concentric semantic stack;
- attachment inspection failure; and
- contradictory existing location.

**Exit:** every required location direction has an explainable imported or created source, or a named failure.

### WP9 — Make ordinate creation a fully recorded transaction

**Objective:** Execute each family-scoped ordinate group with verified selection order and cleanup.

For each group:

1. Activate the approved orthographic view and verify the actual active view.
2. Clear existing selections.
3. Create `SelectData` and select the typed datum entity first with append disabled.
4. Append only the planned unique candidates.
5. Check every `Select4`/`Select3` result.
6. If `MultiSelect2` is used, verify its return and E4-prove order preservation. Otherwise use a proven sequential selection path.
7. Verify the final selection count equals datum plus planned locations.
8. Require at least one non-datum location.
9. Call `AddOrdinateDimension` with the verified direction enum.
10. Record the raw result and decoded `swCreateOrdDimError_e` name. Preserve the installed mapping `GenFailure=1`, `OrdFailure=7`, duplicate=8, and bad direction=9.
11. Identify and record the created group/annotations so QA can verify view ownership and coverage.
12. Call `SetPickMode`, clear selection, and restore sheet/view context on every path.

**Exit:** every ordinate has a proved datum, family, direction, candidate set, return result, final view, and cleanup record.

### WP10 — Replace midpoint sections and decide detail-view scope

**Objective:** Create the one supported primary section through the intended functional feature.

Tasks:

1. Store source-view role, target feature/axis/centre, cut direction, label, destination cell, and expected information.
2. Transform the feature-led cut into the correct drawing/sheet coordinate system.
3. Create the cut line as a `SketchSegment`, check selection, and call the verified section-view API.
4. Check every sketch, selection, view-creation, rebuild, and placement result.
5. Verify source view, cut intersection, arrows, label, hatch, view bounds, and ordinate exclusion after rebuild.
6. P-0251 requires one J-J section through the stepped functional bore.
7. P-0252-01-001 has no accepted-default section.
8. P-0252-01-013 requires one B-B section through the upper ear/threaded-hole intent.
9. Reject more than one configured section until an approved data-model/UI redesign exists.
10. The r4 D-03 policy, pending user override/confirmation, makes Details C and D mandatory at 3:1. Define and prove the exact `*Bottom` source, reference-led centre/profile, label, destination, scale independence, source/detail ownership, and intended small-feature content; do not infer success from label/position alone.

**Exit:** required sections are feature-led, correctly labeled, readable, in bounds, and free of fallback ordinates; required details are explicitly defined and proved.

### WP11 — Establish the controlled title, property, mass, notes, and scale contract

**Objective:** Make every visible title value traceable to an approved source and cell.

First create an approved template/property/cell manifest containing:

- canonical `.drwdot` and `.slddrt` identity/version;
- border margins, projection convention, title bounds, general-note region, and part-ID region;
- drawing property key;
- ordered configuration-level model keys;
- ordered document-level fallback keys;
- approved computed source, if any;
- mandatory/optional/manual policy;
- expected fixture value;
- exact linked-note token and target title note/cell; and
- missing-data behavior.

At minimum cover part number, description, material, mass, quantity, actual displayed scale, project, customer code, unit, surface treatment, heat treatment, revision, drafter/designer, checker, approver, and dates.

Implementation tasks:

1. Read active-configuration properties before approved document fallbacks.
2. Record source scope, raw value, resolved value, and `Get6`/`Add3` result for every field.
3. Never change the model to repair missing metadata.
4. Derive mass only through an installed-build-proven active-configuration path with recorded units. If that path or an authoritative property is unavailable, fail or use an explicitly approved controlled placeholder.
5. Read the actual sheet scale numerator/denominator from `ISheet.GetProperties2` after `SetScale`; do not write the requested UI value as though it were confirmed actual scale.
6. Record and visibly label any approved custom view scale.
7. Enumerate intended title notes through `ITitleBlock.GetNotes` and verify `INote.PropertyLinkedText`, rendered text, and `INote.GetExtent` for the exact target note. A matching string elsewhere is not proof.
8. Verify general notes and part-identification text against their approved link/token and rectangle.
9. Keep the approved three general notes unless authoritative template/model text supersedes them.
10. Keep the current output labeled part identification until D-05 defines barcode symbology, font, scanner, and acceptance standard.

**Exit:** every mandatory field has an approved source, exact visible linked cell, actual value, and containment proof; missing data cannot silently disappear.

### WP12 — Build a standard-scale, collision-aware layout pass

**Objective:** Turn a geometrically valid drawing into a readable manufacturing drawing.

The current view-grid code is only a starting point. The final layout system must:

1. Select an approved standard sheet scale and sheet size from geometry plus estimated annotation load.
2. Keep orthographic sheet-scale views on the actual sheet scale.
3. Prohibit arbitrary independent shrinking. If content does not fit, try the next approved standard scale or sheet size and report the choice.
4. Define layout items for view outlines, title block, general-note cell, part-ID cell, note/text extents, symbols, ordinate lanes, and leader polylines.
5. Use verified APIs such as `INote.GetExtent` and `IAnnotation.GetLeaderCount`/`GetLeaderPointsAtIndex`; research the correct type-specific extent path for dimensions and symbols before implementing it.
6. Treat unavailable mandatory extents as unproved and failed, not visually presumed safe.
7. Allocate dimension lanes around the owning view and keep family callouts near their features.
8. Check every annotation selection before `AlignDimensions`; use its Boolean result only as a first-pass arrangement result.
9. Rebuild and remeasure after each move.
10. Detect usable-border, title/note/ID intrusion; text/text; text/geometry; annotation/unrelated-view; and leader/unrelated-view conflicts.
11. Preserve unambiguous leader-to-feature association.
12. Use a bounded deterministic solver. If no readable result exists, fail with a larger-sheet or approved-scale recommendation.
13. Run a final complete collision pass after title rebuild. Any later view or annotation move invalidates that result and forces revalidation.

Recommended final content order:

1. finalize view roles and standard scales;
2. import/create dimensions;
3. arrange dimensions;
4. populate/rebuild title and notes;
5. resolve complete layout;
6. rebuild and remeasure;
7. perform collision QA;
8. make no further layout mutation.

**Exit:** no material clipping, overlap, title/border intrusion, ambiguous leader, or unreadable scale remains.

### WP13 — Replace count-based QA with requirement-level proof

**Objective:** Make the report explain exactly why each global and fixture requirement passed or failed.

Mandatory QA records:

- run ID, fixture, configuration, build, source revision, embedded/export hash, template, format, and settings;
- required-stage completion ledger;
- planned and created views with roles, types, orientations, scales, outlines, and eligibility;
- imported annotations by type and view;
- unique model feature instances separately from projected candidates;
- accepted/rejected ownership evidence and rejection taxonomy;
- semantic size/callout and X/Y coverage by feature family;
- datum identity and coordinate/selection proof per group;
- every selection, `MultiSelect2`, final count, and ordinate result;
- duplicate and equal-coordinate suppression;
- section/detail definition and observed result;
- title property sources, exact linked-note proof, and cell containment;
- layout items and collision findings;
- actual scale and custom scales;
- final document, sheet, pick-mode, and selection state;
- evidence-directory and report-write result; and
- automated QA result plus separate E7 acceptance state.

Mandatory failures include:

- required model annotations returning zero;
- zero total dimensions;
- unproved required ownership or semantics;
- unproved required datum;
- failure to inspect existing coverage;
- unresolved required size/callout/X/Y location;
- failed required selection or ordinate/API operation;
- duplicate or contradictory coverage;
- any dimension on an isometric view;
- any fallback ordinate on section, detail, isometric, or another prohibited view;
- wrong/missing section intent;
- title, border, note, or part-ID intrusion;
- material collision or uninspectable mandatory extent;
- missing or unresolved mandatory title data;
- failed final cleanup; or
- failed evidence retention.

Fixture checks must use unique semantic instances, not aggregate projected counts. In particular, P-0251's ten reference hole locations are six face counterbore locations plus four side tapped-hole locations; the same hole visible in another view must not increment that count.

Evidence output must use a collision-proof run ID, recursively create its directory, write the final report atomically, and ensure the text written to disk exactly matches the final reported summary.

Perform zoom before final cleanup. After cleanup, verify the active drawing document, active sheet, zero selections, and no cleanup error. The QA/evidence writer must not mutate SOLIDWORKS state after that point.

**Exit:** a missing or unknown mandatory state cannot become `PASS`, and a reviewer can reconstruct every result from the retained run directory.

### WP14 — Add offline tests and freeze the next source identity

**Objective:** Catch every source/API defect that can be detected before asking the user to import.

Required source/structure tests:

- complete component inventory and `Option Explicit`;
- balanced procedures, valid line lengths, and continuation limits;
- public-name/signature/caller consistency;
- `CreateLine` assigned to `SketchSegment` and selected through the supported interface;
- datum selected through `IEntity.Select4`;
- genuine native form-export structure or an explicitly tested blank-form paste workflow;
- host `ThisLibrary` handoff instructions;
- no unconditional 90-degree view rotation;
- no null-component `GetVisibleEntities2` call;
- no unsupported-view ordinate path;
- no arbitrary sheet-scale view conversion;
- actual sheet-scale readback;
- immediate failure for missing template/format/bounds;
- reflection-backed ordinate error mapping;
- required-stage and evidence-write failure semantics.

Required fake-COM/pure-logic tests:

- annotation import whole-drawing and zero-result retry cleanup;
- callout off/on mask delta;
- ownership face-set match, pattern/mirror instance, configuration suppression, and rejection taxonomy;
- canonical physical-location deduplication across projections;
- approved datum success/failure and typed selection;
- every coverage case listed in WP8;
- family-scoped ordinate grouping and selection-count mismatch;
- section feature intersection and label checks;
- requested/actual scale mismatch;
- title token/cell verification, including a matching string in the wrong note;
- missing/wrong-configuration mandatory property;
- note extent outside usable bounds;
- title/note/ID intrusion;
- text-box overlap;
- leader crossing an unrelated view;
- unavailable mandatory extents;
- report-directory collision and write failure; and
- post-cleanup state mutation.

For every new or uncertain COM member, cross-check the SOLIDWORKS 2025 MCP/official documentation and installed interop. Use a narrow live probe only for binding, selection, returned-array, coordinate-frame, or template behavior that static authority cannot settle.

After all tests pass:

1. increment the source identity;
2. generate hashes for every importable component;
3. update the API-contract ledger;
4. update import instructions; and
5. create a user compile handoff naming exact components and order.

**Exit:** the next source revision is one coherent, source-complete offline candidate. This proves E2/E3 only.

### WP15 — User-owned clean import and full VBA compile

**Objective:** Achieve E5 with an exact, recoverable embedded project.

User procedure after the r4 WP14 checkpoint:

1. Create or copy the designated canonical `.swp` and preserve a backup.
2. Confirm the SOLIDWORKS 2025 Type Library and Microsoft Forms 2.0 Object Library references.
3. Retain the existing host `ThisLibrary` and paste the supplied code into it.
4. Import handler/data classes and standard modules in the documented order.
5. Import genuine native forms, or create the two named blank forms and paste their code according to the tested fallback.
6. Run **Debug > Compile VBAProject** for the entire project.
7. If compilation fails, capture the exact message, highlighted line, procedure, screenshot, project name, and references list. Export the affected embedded component.
8. Fix one coherent compile defect at a time and recompile the whole project.
9. After a clean compile, export every embedded component.
10. Compare the embedded exports with the approved source manifest using normalized source hashes where native form metadata is expected to differ.
11. Save and back up the compiled canonical macro.

Do not run the macro until the clean compile and embedded/export identity have been reviewed.

**Exit:** complete project compile succeeds, the source identity is proved, and the canonical binary is backed up. This is E5 only.

### WP16 — Resolve live-only contracts and the controlled template

**Objective:** Close only those contracts that E2/E3 cannot prove.

Potential narrow E4 probes, performed by the user or Codex after an explicit operator choice:

- native form/project behavior if import still differs from the tested package;
- template and sheet-format identity, visibility, border margins, title block, linked notes, and region extents;
- `IsSuppressed2` returned-array behavior for drawing-view configurations;
- drawing-edge corresponding-entity and face/feature ownership on each fixture/view;
- pattern/mirror seed and instance behavior;
- model-to-view centre/axis tolerances;
- approved datum entity selection and coordinate transforms;
- imported linear-dimension attachment patterns;
- selection ordering for the ordinate transaction;
- section-line coordinate frame and post-creation arrows/hatch/label;
- dimension/symbol extents and permitted movement;
- `AlignDimensions` spacing units;
- active-configuration mass calculation and units; and
- title-note `PropertyLinkedText`, rendered text, and extent behavior.

Each probe must record build, fixture, configuration, view, exact source revision, input, raw result, conclusion, and cleanup. A probe must not modify or save the model.

**Exit:** every runtime-blocking uncertainty is either live-proven or causes the production path to fail closed.

### WP17 — Focused P-0251 integration and acceptance

**Objective:** Prove the whole architecture on the smallest fixture with a clear ten-location and section requirement.

Recommended initial acceptance profile, subject to the final role/orientation plan:

- P-0251-14A-001;
- Bottom-Left approved datum;
- Hole Wizard callouts on;
- auto-arrange on;
- actual 1:1 sheet scale;
- required title, notes, part identification, and QA on;
- principal and required side roles plus one isometric; and
- exactly one horizontal section labeled J through the stepped bore.

Required acceptance evidence:

- manufacturing-led principal and side views with no historical rotation mismatch;
- one J-J section through the stepped functional bore;
- one clean undimensioned isometric;
- six unique face counterbore locations and four unique side tapped-hole locations;
- grouped `6X Ø6.6 THRU`/counterbore semantics and `4X Ø4.2`/M5 thread/depth semantics from authoritative data;
- `Ø47 H7`, `Ø40`, `R36`, and required thickness/step coverage;
- approved centre/bottom location logic;
- no duplicates and no ordinates on section/isometric views;
- complete title values and notes in their approved cells; and
- no material collision or clipping.

Use the smallest focused rerun after each evidence-backed defect. Do not proceed to Base Plate acceptance until P-0251 has one retained golden E6/E7 run.

### WP18 — Base Plate and Pump Holder acceptance

#### P-0252-01-001 — Base Plate

Required proof:

- principal face and thin edge roles with clean orientation aid(s);
- no section in the accepted default;
- through, blind, tapped, precision-fit, counterbored, radius, chamfer, and back-face families remain distinct;
- front/back machining intent is explicit;
- every family has size and required X/Y location coverage;
- repeated equal coordinates are suppressed only within the correct group;
- the approved global or multi-origin policy is applied without invention;
- no compressed edge view, leader forest, title intrusion, or copied reference congestion; and
- AL6082, mass 11.13, quantity 1, 1:6 scale, BASE PLATE, and Anodizing are confirmed from the approved data contract.

#### P-0252-01-013 — Pump Holder

Required proof:

- manufacturing-useful profile/lower/side roles;
- one B-B section through the threaded-hole intent;
- clean undimensioned orientation aid(s);
- `Ø3.3` drill/depth, M4 thread/depth, and location coverage;
- datum A and perpendicularity controls preserved;
- `6.4 g6`, radii, angles, and profile dimensions preserved from authoritative annotations;
- `7 × 4` end/tab geometry and `C0.5` chamfers are readable, with Details C/D when required;
- no ordinates on section, detail, or isometric views;
- no collision or clipping; and
- AL6082, mass 0.07, quantity 1, 1.5:1 scale, PUMP HOLDER, and Anodizing are confirmed from the approved data contract.

**Exit:** each fixture has one signed-off golden run directory and explicit comparison notes against its manual reference and target requirements.

### WP19 — Complete regression matrix and release candidate

Use focused pairwise coverage rather than an unnecessary full Cartesian product, but prove every exposed behavior:

| Coverage dimension | Minimum proof |
|---|---|
| Fixtures | All three accepted golden configurations |
| Datum choices | Bottom-Left, Center, and Top-Left wherever exposed; unsupported choices fail closed |
| Hole callouts | Off/on for all fixtures with mask, returned-object, and visible-output delta |
| Auto-arrange | Off/on, including the crowded Base Plate |
| Sections | No-section diagnostic; required P-0251 J-J; required Pump B-B; one vertical behavior case where meaningful |
| View policy | Eligible orthographic/projected views; no isometric dimensions; no section/detail/isometric ordinates |
| Ownership | Hole Wizard, Advanced Hole/Simple Hole where present, generic-cut rejection/acceptance evidence, pattern/mirror, suppressed configuration |
| Coverage | Imported X, Y, X/Y, none, size-only, equal coordinate, cross-family equality, concentric stack, contradiction |
| Title | Every required field, missing-field failure, actual scale, unit-aware mass, link/cell proof, part-ID placement |
| Layout | Standard scales, dense annotations, title/border intrusion, text overlap, leader conflict, cannot-fit failure |
| Cleanup | Active drawing/sheet, normal pick mode, zero selection, no model dirty/save |
| Evidence | Unique directory, atomic report, exact final summary, source/build/template/settings identity |

Before release:

1. Start a fresh SOLIDWORKS session.
2. Compile the canonical macro again.
3. Run the three golden configurations without stale registry state.
4. Compare embedded exports to the approved source manifest.
5. Confirm the parts were not dirtied or saved.
6. Review full sheets and dense regions against the manual references.
7. Record E7 acceptance or named deviations.
8. Update `CURRENT_STATUS`, `Changelog`, target specification, import guide, source manifest, and release evidence.

**Exit:** all required matrix cells pass and the accepted embedded binary matches the released source identity.

## 9. Decisions requiring user input

These decisions should be requested just before they block their dependent work. Safe provisional rules allow unrelated offline work to continue.

| ID | Priority | Decision | Blocks | Safe provisional rule |
|---|---:|---|---|---|
| D-00 | 1 | Canonical new `.swp` name/path, and whether it eventually replaces `active_ordinate.swp` | WP15/source synchronization | Keep the replacement separate and do not overwrite the broken macro |
| D-04 | 1 | Authoritative `.drwdot`/`.slddrt`, version, property names, and title/note/ID cell map | Runtime, title, layout acceptance | Fail rather than use an uncontrolled template or invented mapping |
| D-02 | 2 | Whether Base Plate may use multiple approved feature-family origins | Coverage data model and Base Plate acceptance | Support one approved global origin first; do not invent more |
| D-06 | 2 | Exact selectable Bottom-Left, Center, and Top-Left entities per fixture/view/family | Datum and ordinate acceptance | Prove geometry; fail closed when unresolved |
| D-03 | Policy set | Confirm or override r4's mandatory Pump Holder Details C/D rule | Pump section/detail plan | Require both at 3:1; retain E4/E6 transaction proof and E7 content/legibility proof |
| D-01 | 4 | Whether Base Plate and Pump Holder require both reference isometrics | Final view plans | Require at least one; add a second only when it adds necessary opposite-side understanding and fits |
| D-07 | 4 | Approved standard sheet sizes/scales and when custom isometric scale is permitted | Layout acceptance | Use reference scale as a starting point; fail instead of arbitrary shrinking |
| D-08 | 4 | Missing metadata policy: failure versus controlled drawing-level placeholder | Title acceptance | Never modify the model; fail mandatory missing data |
| D-05 | 5 | Barcode symbology, font, scanner, and validation standard | Barcode claim only | Continue truthful part-identification text |

The first useful user-supplied dependency is D-04. It does not block offline structural work or compilation, but it blocks the first macro run.

## 10. Required evidence packet for every compile/run iteration

Each retained run directory should contain:

- run manifest with run ID and timestamp;
- fixture path and active configuration;
- SOLIDWORKS build;
- canonical macro path and revision;
- embedded/export source hashes;
- template and sheet-format identity;
- complete settings and fixture view plan;
- compile result or exact compiler error;
- complete Immediate Window log;
- final `QA_REPORT.txt` or structured JSON plus exact automated/E7 status;
- full-sheet screenshot;
- close-ups of dense, ambiguous, or failed regions;
- PDF/JPG/BMP export where useful;
- property/link/cell evidence;
- relevant feature ownership/datum/section probe evidence;
- reference-comparison notes;
- model dirty/save-state check; and
- remediation decision for every failure or accepted deviation.

Use a new directory for every attempt. Never overwrite a failed run or manual reference.

## 11. Iteration protocol

For every compile/runtime/visual defect:

1. Preserve the failed output and exact source identity.
2. State the observed symptom and evidence level.
3. Identify the likely root cause, affected procedure, and all callers/data contracts.
4. Verify uncertain methods, enums, return types, and coordinate systems against the 2025 MCP/official docs/installed interop.
5. Use a narrow live probe only if E3 cannot settle the binding or fixture behavior.
6. Prepare the smallest coherent complete-module replacement.
7. User imports it and compiles the whole project.
8. User runs the smallest affected authorized case.
9. Retain logs, QA, images, source identity, settings, and model state.
10. Compare with the previous failure, specification, and reference.
11. Promote the change only after its focused gate passes.

Do not combine unrelated ownership, title, section, and layout changes in one evidence iteration unless a shared public data contract makes them inseparable.

## 12. Principal risks and mitigations

| Risk | Consequence | Mitigation |
|---|---|---|
| Export package is not native-importable | User cannot reach E5 | Fix forms/host-component workflow and test the import package before handoff |
| Offline code still has VBA binding errors | Repeated compile cycles | Reflect return/interface types, run structure tests, then compile one coherent revision |
| Controlled template is absent or wrong | Runtime/title/layout cannot be judged | Resolve D-04; fail before `NewDocument`; no default fallback |
| Generic cuts or external cylinders are classified as holes | Incorrect ordinates | Require matched-face ownership, internal void proof, semantics, and retained rejection reasons |
| Pattern/projection counts are conflated | False fixture completeness | Keep unique model instances separate from projected candidates |
| Per-view extrema are treated as one datum | Entire coordinate scheme is wrong | Resolve approved model-space datum identities and record every transform/selection |
| Imported linear coverage is ignored | Duplicate fallback ordinates | Use attachment-proven family/direction coverage and stop on inspection failure |
| Equal coordinates are suppressed across families | Required locations disappear | Scope suppression to family/view/datum/direction |
| Midpoint section misses the functional feature | Visually valid but useless section | Bind the cut to the approved feature/axis and verify post-build evidence |
| View outlines pass while text/leaders collide | Drawing is unreadable | Build complete layout items plus mandatory visual E7 review |
| Arbitrary independent scaling changes drawing truth | Title scale and readability are false | Use approved standard scales and actual sheet readback |
| Missing model properties produce plausible blanks/defaults | False title block | Approved property manifest, source tracing, and fail-closed policy |
| QA totals mask missing family coverage | False `PASS` | Requirement-level stage/feature records and mandatory unknown-state failure |
| Evidence write or cleanup fails after apparent success | Run cannot be trusted | Make both mandatory final gates and perform cleanup last |
| Documentation overstates progress | False release confidence | Preserve E1-E7 vocabulary and update status only from retained evidence |

## 13. Definition of done

The new target-spec hybrid macro is complete only when all of the following are true:

1. The source-complete offline revision passes all source, fake-COM, regression, and reflection-backed tests.
2. Every used API contract is verified through E3 or a retained, narrowly scoped E4 probe.
3. The canonical clean macro project compiles in SOLIDWORKS 2025.
4. Embedded source exports match the approved release manifest.
5. All runs are confined to the three fixtures and no fixture is modified, dirtied, or saved.
6. The controlled template/property/cell contract is approved and runtime-proven.
7. Every required view has the correct role, orientation, scale, information purpose, and bounds.
8. Required sections/details intersect the intended features and remain readable.
9. Every required semantic feature family has authoritative size/callout and complete location coverage.
10. Every ordinate uses an approved selectable datum and eligible orthographic view.
11. No duplicate, contradictory, or cross-family-suppressed required dimension remains.
12. Isometrics are undimensioned and prohibited views contain no fallback ordinates.
13. Required title data, notes, actual scale, and part identification resolve in their approved cells.
14. No material view, annotation, text, leader, border, title, note, or part-ID collision remains.
15. Final cleanup and evidence retention pass.
16. P-0251, Base Plate, and Pump Holder each have an accepted golden E6/E7 run.
17. The focused regression matrix passes from a fresh SOLIDWORKS session.
18. Release documentation and evidence accurately name the final binary and source identity.

## 14. Exact resumption point

The r4 offline source phase is complete. Resume in this order:

1. Use the frozen r4 source archive and 35-entry SHA-256 manifest under
   `test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/`.
2. Resolve D-00 for the canonical recoverable `.swp` copy; request D-04
   materials in parallel.
3. Perform WP15 clean import and compile the whole project without running it.
4. Review the exact compiler result and export/manifest identity.
5. Resolve D-04 and run only the narrow WP16 live checks that E2/E3 cannot
   prove.
6. Run the focused P-0251 acceptance profile and retain the complete evidence
   packet.
7. Continue to Base Plate, Pump Holder, and the fresh-session regression only
   after each preceding gate passes.

The immediate next implementation task is therefore **WP15: clean import and
full VBA-project compile**. It is an E5 gate, not permission to run the macro.
