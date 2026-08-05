# R23 Target-Spec Hybrid Drawing Implementation Plan

**Status:** Phases 0-8 were satisfied at r29. r37 passed focused read-only
execution and proves both vertical outline datums with a scoped drawing-view
selection fallback when SOLIDWORKS returns no polyline-edge array. r31 preserves the user-accepted layout as-is:
automatic envelope repositioning/rescaling is retired from production and its
semantic stage is an explicit waiver, not a clearance pass. Production R23
remains blocked on the non-layout semantic gates, full production mutation
proof, and the three-fixture regression.
**Prepared:** 2026-07-30; updated 2026-08-04

## Remaining work at a glance (2026-08-04)

Read this before planning any R23 work. Everything below is open; everything
not listed here is gated SATISFIED.

| Item | State |
|---|---|
| Phase 9 `R23-903`, `R23-904` | closed by user decision: preserve approved layout; no automatic move/rescale |
| r37 non-layout corrections | 435/435 static; 38/38 deployment/readback; programmatic VBE clean; P-0251 runner complete |
| Phase 9 layout | diagnostics retained; automatic clearance is waived, never reported as passed |
| Phase 10 gate | r37 scratch runner: 9 of 10; only creation-dependent model-import coverage remains open |
| R23-502 outline datum | satisfied: straight lower model edges map by scoped target-view selection when `GetPolylines7` returns no edges |
| P-0252 semantic section | no graph-backed path defined; production fails closed |
| P-0252 fixture drawings | no disposable `.SLDDRW` exists under `test_assets/`; three-fixture runtime cannot start |
| Phase 11 `R23-1100` to `R23-1103` | source wired; VBE compiled and P-0251 read-only evidence retained; production mutation unproved |
| Phase 12 `R23-1208` to `R23-1211` | r37 focused runner retained; production artifacts remain open |
| Section 7 acceptance checklist | production drawing and visual acceptance remain open |
| Accumulated mutating work, Phases 5-9 | only P-0251 layout scale was authorized and saved on scratch |

**Latest read-only runtime proof (r37):**
`test_assets/iteration_evidence/probe_runs/20260804_164014/probe_log.txt`
records guarded deployment/readback 38/38, `R23_RUN_COMPILE|verdict=Clean`,
and all nine probes `Completed` against the saved P-0251 scratch drawing. It
resolved 44 projections and proved 11/11 projection selections. Every probe
remained read-only with selection count 0 to 0.

r37 calls `IView.GetPolylines7` first. The installed build returns no
polyline-edge array for both valid HLV target views. Only this explicit
no-edge state permits Route D: the model
edge is selected by `IView.SelectEntity` and the selection manager must read
its owner back as the requested drawing view. Both vertical schemes are now
`OutlineDerived`; semantic QA proves 9/10. The remaining failure is
model-import coverage in `Drawing View4`, which is creation-dependent and
cannot be passed by a read-only copy of a manual drawing. The r31 user decision
retires automatic envelope movement/rescaling and treats `FINAL_LAYOUT` as
`UserAcceptedLayoutAsIs|automaticClearance=DeferredByUser`.
The scratch began as an exact SHA-256-matched copy at
`test_assets/scratch_drawings/P-0251-14A-001-R23-layout-scratch.SLDDRW`.
`Module2_DrawingPipeline.R23_ApplyContentLayoutToScratch` refuses every other
path, proves the referenced local fixture, records evidence, and never saves.
Its r28 rescale loop made two measured scale passes, then rejected a third
request. The user accepted that scratch result; it was saved through the
installed SW2025 `IModelDoc2.Save3` contract with `errors=0|warnings=0`.
That human acceptance does not convert `FINAL_LAYOUT=False` into an automatic
or production gate pass.

**Target revision:** `target-spec-hybrid-v2-2026-08-04-r37`
**Production source:** `src/target-spec-hybrid-v2/`
**Protected source:** `src/baseline-model-dims/` — read-only

### Implementation checkpoint — 2026-07-30

- Pre-change `Fable.swp` SHA-256:
  `C8076D713DD8F64AC75F93871C1EA4A2D1F01EE2BC691323736013BFEB2803F2`.
- Recoverable backup:
  `test_assets/iteration_evidence/r23/20260730-075811/prechange/Fable.pre-r23.swp`.
- The backup hash matches the production SWP.
- Read-only embedded-source verification passed: 15/15 managed components,
  with both embedded and exported source at
  `target-spec-hybrid-v2-2026-07-29-r22`.
- The offline guarded deployment preflight passed with 19 embedded components,
  15 managed components, and the deployment bootstrap present.
- The read-only feature probe completed against the authorized P-0251 fixture:
  `R23_FEATURE_20260731_040539.log` records 47 visited features, all three
  `ICE -> Cut` resolutions, typed definition access/release, owned geometry,
  native Hole Wizard values, mirror seed evidence, and active-configuration
  fallback suppression results.
- Both tested `IsCircle`/`GetCurveParams3` call orders remained stable on the
  counterbore, tapped, mirrored, and cut-cylinder boundaries. `CircleParams`
  nevertheless returned the probe's `SkippedNotCircle` sentinel while both
  before/after predicates were true. R23 must not call or rely on
  `CircleParams`; it will use the live-proven cylindrical-face radius,
  `IsCircle`, and complete-boundary curve-parameter evidence instead.
- A guarded three-module overlay and copied `Fable.swp` are prepared under
  `tools/r23-probes/` and the R23 live-evidence folder for the
  `AllViews=True` versus deterministic `AllViews=False` comparison. The
  disposable manifest passed offline preflight and then completed both
  user-operated runs. `AllViews=True` returned 25 annotations but placed them
  only in the primary (8) and section (17) views. The explicit section/side/
  primary sequence returned the same 25 exactly once, distributed as primary
  (6), side (2), section (17), and isometric (0). R23 will retain the latter.
  The M5 native callout was imported only in the section; no CBORE/M6 native
  callout was imported in either variant.
- No production `.bas`, `.cls`, `.frm`, deployment manifest, or SWP has been
  changed. Production implementation remains blocked by the Phase 0 runtime
  gate below.
- `Module_R23Phase0DrawingProbes.bas` and its guarded disposable deployment
  manifest received user-operated compilation and runtime execution. The
  ordinate entry point stopped before `AddOrdinateDimension` because its opaque
  feature-edge mapper returned no six-location counterbore set. The section
  entry point found the exact 47/40 diameter dimensions and a complete
  section-line payload, but also exposed stale loop-state logging and mixed
  coordinate frames. The probes must be corrected and rerun before production.
- The four-module overlay passed guarded read-only preflight. All overlay
  sources satisfy the deployer encoding/metadata contract; the new probe has
  balanced procedure blocks, safe line/continuation lengths, and no duplicate
  procedure-local declarations. Installed SOLIDWORKS 2025 interop reflection
  confirms the load-bearing probe signatures and enum values. The complete
  offline suite reproduced the known 69-pass/5-stale-R20-failure baseline.

## 1. Purpose

R23 will replace the count-oriented R22 feature/dimension path with a
model-intent, physical-location, view-projection, and semantic-coverage
workflow. The result must satisfy the drawing target specification rather than
merely finish the macro or create a nonzero number of dimensions.

This plan covers:

- importing general, marked, toleranced, and Hole Wizard model annotations;
- recognizing holes and manufacturing features created both with and without
  Hole Wizard;
- creating the reference-style location scheme without treating imported
  linear dimensions as ordinate coverage;
- producing complete face-hole, side-hole, and stepped-bore definitions;
- creating a safe and semantically located J-J cutting line;
- adding real section dimensions and tolerance readback;
- laying out views using complete annotation envelopes;
- failing QA when manufacturing intent or geometry cannot be proved.

This document does not authorize a production-macro edit by itself. No
`.swp`, `.bas`, `.cls`, `.frm`, deployment manifest, model, generated drawing,
or protected baseline file was changed while preparing it.

## 2. Fixed boundaries

- Only these fixtures may be used:
  - `test_assets/models/P-0251-14A-001.SLDPRT`
  - `test_assets/models/P-0252-01-001.SLDPRT`
  - `test_assets/models/P-0252-01-013.SLDPRT`
- The fixtures may be opened and inspected but never modified or saved.
- `src/baseline-model-dims/` must remain unchanged.
- Manual reference drawings must remain unchanged.
- VBA remains the production drawing generator.
- Before any live SOLIDWORKS, VBA Editor, or Computer Use task, ask whether the
  user will perform the exact task and share its result or wants Codex to use
  Computer Use.
- Make a recoverable backup of the production SWP before the first material
  embedded-project change.
- Static verification, MCP/API evidence, installed-interop reflection,
  embedded VBA compilation, runtime execution, and visual/manufacturing
  acceptance are separate gates.

## 3. Evidence incorporated

### 3.1 R22 runtime evidence

The P-0251 R22 run established:

- four intended views were created;
- selected-view model import produced nine display dimensions in the primary
  view and one in the side view;
- the section and isometric contained zero display dimensions;
- 32 model cylinder boundaries were rejected in each eligible orthographic
  view as `ClosedCircleCurveNotCircular`;
- zero circular proofs, mapped edges, ownership candidates, canonical physical
  locations, ordinate groups, or manufacturing callouts were produced;
- the J-J cutting line and upper arrow entered the border/zone region;
- the lower arrow/label entered the part-identification band.

### 3.2 Reference drawing requirements

P-0251 requires:

- one primary face view, one side view, Section J-J, and one clean
  undimensioned isometric;
- six face counterbores with one grouped definition;
- four tapped side holes with one grouped definition;
- symmetric centre-X and bottom-Y location logic;
- a stepped-bore section with `Ø47 H7 +0.025/0.000`, `Ø40`, depths, wall/step
  sizes, and overall section dimensions;
- no ordinates on the section or isometric;
- no border, title, part-identification, annotation, or view collision.

### 3.3 Feature-list output

The supplied diagnostic macro reported these manufacturing-relevant features:

| Feature name | `GetTypeName2` result | Initial role |
|---|---:|---|
| `Boss-Extrude1` | `Extrusion` | base material |
| `CBORE for M6 Socket Head Cap Screw1` | `HoleWzd` | six-hole counterbore candidate |
| `Cut-Extrude1` | `ICE` | underlying type unresolved |
| `Cut-Extrude3` | `ICE` | underlying type unresolved |
| `Cut-Extrude4` | `ICE` | underlying type unresolved |
| `M5x0.8 Tapped Hole1` | `HoleWzd` | tapped-hole family candidate |
| `Mirror1` | `MirrorPattern` | derived-location candidate |

Feature names are diagnostic labels only. They must not supply manufacturing
semantics.

### 3.4 Meaning of `ICE`

Official SOLIDWORKS 2025 API Help states:

- `IFeature.GetTypeName2` returns `ICE` for an Instant3D feature;
- when `GetTypeName2="ICE"`, call `IFeature.GetTypeName` to obtain the
  underlying feature type;
- use `IFeature.GetDefinition` to obtain the matching feature-data interface
  when one exists.

Therefore `ICE` does not mean “imported cut,” and it is not a usable
manufacturing classification.

The current R22 resolver returns `GetTypeName2` whenever it is nonempty.
Consequently the three tree entries named `Cut-Extrude1`, `Cut-Extrude3`, and
`Cut-Extrude4` remain classified as raw type `ICE` and never reach the branch
that currently looks for `"CUT"`. The supplied listing alone does not prove
their underlying type; Phase 0 must capture `GetTypeName`. This classification
defect nevertheless explains why the R22 run accepted only three semantic
features: the two `HoleWzd` features and the `MirrorPattern`.

The protected baseline already contains the correct special case:

```text
If GetTypeName2 = "ICE" Then use GetTypeName
```

R23 must restore that API-required behavior without copying the baseline's
unsafe geometry, selection, duplicate, or unattended-callout paths.

## 4. Decisions fixed for R23

### 4.1 Import mask and duplicate policy

- Add `swInsertDimensions = 8`.
- Add `swInsertTolerancedDims = 16777216`.
- Retain:
  - `swInsertDatums = 2`
  - `swInsertGTols = 32`
  - `swInsertNotes = 64`
  - `swInsertDimensionsMarkedForDrawing = 32768`
  - `swInsertHoleWizardProfileDimensions = 65536`
  - `swInsertHoleWizardLocationDimensions = 131072`
  - conditional `swInsertholeCallout = 1048576`
- Retain `DuplicateDims=True`. For `InsertModelAnnotations4`, `True` eliminates
  duplicate model dimensions.
- Keep `InsertAllAnnotations=False` and
  `InsertAllReferenceGeometry=False`, so the `Types` mask remains effective.
- The callout-enabled combined mask is `18055274` (`0x113806A`).

`swInsertDimensions=8` requests general model dimensions. It does not replace
the native hole-callout bit.

### 4.2 Callout policy

Use this fixed priority:

1. complete native callout imported by `InsertModelAnnotations4`;
2. native cosmetic/model-thread callout after its selection contract is
   runtime-proved;
3. controlled associative note generated only from typed model feature data
   and imported model tolerance data;
4. fail closed if required manufacturing intent is absent.

Do not use `IDrawingDoc.AddHoleCallout2` in unattended production because its
documented workflow requires user confirmation.

### 4.3 Manufacturing-intent policy

- Hole Wizard and modeled/cosmetic thread data may provide typed size, depth,
  end condition, standard, thread, class, fit, and callout fields.
- A proved extruded cut may provide diameter, depth, direction, and end
  condition.
- A generic cut cannot establish thread class, tap size, fit, tolerance, or
  machining side from geometry alone.
- Fit and tolerance must come from an imported/source model dimension or other
  approved model data.
- The manual reference defines acceptance requirements. It is not a silent
  source from which the macro may invent model intent.

### 4.4 Location policy

- Imported linear dimensions do not satisfy a target that requires ordinates.
- Only an imported ordinate attached to the approved datum and exact physical
  location may suppress a generated ordinate in that direction.
- P-0251 must resolve ten canonical small-hole locations:
  - six face counterbores;
  - four side tapped holes.
- The stepped bore is a separate definition stack and does not count as an
  eleventh small-hole location.

## 5. R23 production data model

The current `CHoleCandidate` combines model identity, feature semantics,
physical location, drawing projection, selectable geometry, import coverage,
and ordinate state. R23 will split these responsibilities.

### 5.1 `CFeatureDefinition`

Store:

- model feature and resolved seed feature;
- feature name for diagnostics only;
- raw `GetTypeName2`;
- raw `GetTypeName`;
- normalized effective feature type;
- referenced configuration and suppression proof;
- typed feature-data interface and read status;
- owned faces;
- operation kind;
- typed diameter, depth, counterbore, thread, fit, and tolerance fields;
- proof source for every semantic field.

### 5.2 `CPhysicalHoleLocation`

Store:

- model/configuration/component identity;
- sign-normalized axis;
- line moment;
- axial interval;
- opening/access side;
- all coaxial cylindrical stack members;
- aggregate manufacturing definition;
- semantic family signature;
- unique physical-instance key.

Feature names must not appear in the physical key. Coaxial cylinders created by
separate Cut-Extrudes must be able to form one stepped-bore stack. Opposite
blind holes on the same infinite axis must remain separate.

### 5.3 `CViewHoleProjection`

Store:

- reference to one physical location;
- view and component;
- model-context aliases;
- drawing-context selectable aliases;
- projected centre and transformed axis;
- selected primary anchor;
- X/Y location-coverage state;
- imported annotation/callout attachments.

### 5.4 `CLocationGraph`

Maintain indexes by:

- feature identity;
- owned face/entity;
- physical-location key;
- semantic-family key;
- view projection;
- attached imported annotation;
- model-dimension identity.

`CRunEvidence` will record results but will no longer serve as the production
object store.

## 6. Ordered implementation task list

Tasks are checked only when their stated evidence exists. Static preparation,
live runtime evidence, embedded compilation, and drawing acceptance remain
separate gates.

### Phase 0 — Preserve and probe

- [x] **R23-000:** Record `git status`, current branch, current production
  revision, SWP hash, managed-source hashes, and existing uncommitted files.
- [x] **R23-001:** Back up the production SWP to a timestamped recoverable
  evidence folder before changing any embedded project.
- [x] **R23-002:** Export/read back current embedded VBA and compare it with
  `src/target-spec-hybrid-v2/`. Treat the current VBA Editor project as
  authoritative if a mismatch exists.
- [x] **R23-003:** Prepare a read-only feature-type probe that prints, for every
  feature:
  - name;
  - `GetTypeName2`;
  - `GetTypeName`;
  - normalized effective type;
  - `TypeName(GetDefinition)`;
  - suppression inputs/results for the referenced configuration.
- [x] **R23-004:** For each P-0251 `ICE` feature, prove that
  `GetTypeName` returns a supported underlying type before treating it as an
  extruded cut.
- [x] **R23-005:** For each resolved cut, probe:
  - `GetDefinition` result;
  - whether it supports `IExtrudeFeatureData2`;
  - `AccessSelections` result;
  - boss/cut state;
  - forward/reverse depth;
  - end condition;
  - sketch/contour state;
  - owned face count and cylindrical owned-face count.
- [x] **R23-006:** Compared both circular-curve read orders on one
  counterbore, tapped hole, mirrored hole, and resolved cut cylinder:
  - R22 order: `GetCurve → GetCurveParams3 → IsCircle → CircleParams`;
  - retained-probe order:
    `GetCurve → IsCircle → CircleParams → GetCurveParams3`.
  - `probe_runs/20260804_125350/` records all four `status=Pass`, with
    `IsCircle=True`, seven `CircleParams` values, equal radius, and zero
    closure for each order: 5.5 mm counterbore, 2.1 mm M5 tapped, 2.1 mm
    mirrored, and 20 mm extruded-cut cylinder.
  - `R23_CURVE_ORDER_END|failures=None`, `catalogFailures=None`,
    `modelUnchanged=True`, and the nine-probe runner remained read-only.
- [x] **R23-007:** Compare expanded-mask import on a disposable P-0251 drawing:
  - one selected view with `AllViews=True`;
  - deterministic selected-view calls with `AllViews=False`;
  - always `DuplicateDims=True`.
- [x] **R23-008:** For each import variant, record owner view, source dimension
  identity, `IsHoleCallout`, callout variables, attachments, nominal,
  tolerance, fit, duplicate state, extent, and visual overlap.
  - Both transactions returned 25 unique source identities and left the
    isometric clean. `AllViews=True` produced primary/section counts of 8/17
    and no side dimensions. Deterministic `AllViews=False` produced 6/2/17
    across primary/side/section, respectively, without multiplying duplicates.
  - The only imported native hole callout is the M5 callout, in the section;
    it exposes tap-drill diameter/depth, `M5x0.8`, and thread depth variables.
    No imported counterbore/M6 callout, H7 fit, or nonzero tolerance was found.
- [x] **R23-009:** Probe one explicit datum-first X group and one Y group using
  deterministic appended selections and verify result/readback.
  - **Closed 2026-07-31.** Both groups returned
    `AddOrdinateDimension = 0 Success` with exact selection counts
    (datum + 2 = 3, datum + 3 = 4), display-dimension deltas +2 and +3,
    `ownerView=Drawing View1` on every selection, `SetPickMode` called, and
    zero selections remaining. Values are `+15.00`/`-15.00` about the
    stepped-bore centre and `10.00`/`50.00`/`90.00` from the bottom-left
    vertex datum.
  - Entity correspondence: use route A,
    `IView.GetCorrespondingEntity(modelEdge)`. `IComponent2` mediation
    returns Nothing on this build. Mapping is per-edge, so every owned edge
    must be attempted before failing a location.
  - `ISelectData.View` assignment raises runtime error 91 on this build;
    activate the view first and prove ownership with
    `ISelectionMgr.GetSelectedObjectsDrawingView2`.
  - Created ordinates report `Type2 = 1` (horizontal request) and
    `Type2 = 7` (vertical request). QA must accept `1`, `7` and `8`.
  - The first run stopped fail-closed with
    `CounterboreLocationsUnavailable` after finding the visible component. It
    never selected a datum or called `AddOrdinateDimension`; this is not an
    ordinate API failure.
  - Correct the disposable mapper to log every owned face/edge qualification
    and compare three mapping routes: direct active-part edge to
    `IView.GetCorrespondingEntity`, active-part edge through
    `IComponent2.GetCorrespondingEntity` before view mapping, and
    `IView.GetVisibleEntities2(component, Edge)`.
  - Retain ownership, cylindrical-radius, `IsCircle`, and closed
    `GetCurveParams3` proof. Do not weaken the contract to arbitrary visible
    circles.
  - On the corrected run, require six mapped M6 counterbore locations, two
    unique X coordinates, three unique Y coordinates, a mapped stepped-bore
    centre datum, and a mapped bottom-left vertex datum before calling the API.
    Record every `Select4`, result code, display-dimension delta, `SetPickMode`,
    and zero-selection cleanup.
- [x] **R23-010:** Complete corrected section-dimension, tolerance-authority,
  and J-J page-geometry proof. **Closed 2026-07-31:** exactly one
  `DIAMETER_47`, one `DIAMETER_40` and one linear target with no stale
  labels; direct part-source readback proved H7 absent and the authority
  decision is recorded in R23-806; the J-J payload segment frame is proved
  against the captured `CreateLine` inputs at `deltaM=0`, converted to page
  coordinates exactly once, and clearance measured against every reserved
  region. Three top-border violations remain as production work for R23-704.
  - The first run found the exact 47 mm `D1@Sketch4` and 40 mm `D1@Sketch6`
    imports as `swDiameterDimension = 6`, not type 15. Both had tolerance type
    zero and no H7 evidence.
  - Reset `nominalM`, `targetName`, `dimensionH7Proven`, and every other
    per-item diagnostic value at the start of each display-dimension loop
    iteration. Require exactly one 47 target, one 40 target, and one selected
    linear target.
  - Add direct part-source tolerance readback for `D1@Sketch4`. If H7 is absent
    there too, obtain the user's explicit authority policy before production:
    model-authoritative fail-closed, or controlled target-spec/reference
    authority with recorded provenance.
  - Preserve the structurally valid 49-item J-J payload, but identify every
    coordinate frame and transform segment endpoints back to page coordinates
    exactly once before clearance comparisons. Compare only page-space bounds
    to page-space border, zone, title, and part-identification extents.
- [x] **R23-011:** Record probe findings in
  `docs/SOLIDWORKS_API_VALIDATION.md`, clearly separating MCP evidence,
  installed-interop presence, and live behavior.

**Phase 0 gate: SATISFIED 2026-07-31.** The corrected probes proved
feature-to-view entity mapping, datum-first X/Y ordinate creation with
success return codes and clean cleanup, non-stale 47/40 section readback, the
chosen H7 authority, and page-coordinate J-J clearance. The already accepted
`ICE -> underlying type`, typed-definition, circular-boundary, and selected
import-transaction evidence remains valid. Production Phase 1 is unblocked.

Contracts carried into production from Phase 0:

- entity correspondence uses route A only;
- attempt every feature-owned edge, because mapping is per-edge;
- never assign `ISelectData.View`; activate the view and verify ownership;
- normalize every SOLIDWORKS COM Boolean with `(CDbl(raw) <> 0#)` before any
  negation or compound logic;
- accept ordinate `Type2` values `1`, `7` and `8`;
- `ICurve.CircleParams` is available and returns correct radii; and
- section-line payload segments are view-sketch coordinates and need exactly
  one conversion before page-frame comparison.

### Phase 1 — Add the graph model

- [x] **R23-100:** Add `CFeatureDefinition.cls`. Every semantic
  manufacturing field is paired with a proof-source string, and
  `UnprovenSemanticFields` lists any field carrying a value without proof.
- [x] **R23-101:** Add `CPhysicalHoleLocation.cls`. Identity is the
  sign-normalized axis plus line moment plus axial interval;
  `CanConsolidateWith` requires both the same infinite line and meeting
  intervals.
- [x] **R23-102:** Add `CViewHoleProjection.cls`. Records the page-frame
  centre with an explicit frame proof, the anchor route that produced the
  selectable entity, and fails closed with `ProjectionAnchorUnavailable`.
- [x] **R23-103:** Add `CLocationGraph.cls`. Nine indexes;
  `ResolveOrCreatePhysicalLocation` is the single consolidation point and
  reindexes on merge so no location is reachable under two keys.
- [x] **R23-104:** Add `CImportedAnnotation.cls` with source/view/
  attachment/tolerance/callout provenance, and a `ProvenanceSource` field
  that keeps a reference-supplied tolerance distinguishable from model data.
- [x] **R23-105:** Add `Module11_GeometryIdentity.bas` with canonical
  normalization for axes, line moments, axial intervals, locations, radii,
  and tolerance comparison. No SOLIDWORKS API calls, so its invariants are
  checkable without a live session.
- [x] **R23-106:** Added `tests/test_r23_location_graph_contracts.py` in the
  companion suite: 20 static source-contract tests covering coaxial
  consolidation, interval-based separation of opposite blind holes,
  moment-based distinctness of six counterbores, seed/mirror separation,
  and exclusion of feature names from both the physical and family keys.
  These are static contracts, not runtime proofs.
- [x] **R23-107:** The graph is additive. A test asserts that no existing
  production module references `CLocationGraph`, `CPhysicalHoleLocation` or
  `CViewHoleProjection`, so the old path is untouched until Phase 2 migrates
  callers.

**Phase 1 status:** source-complete and statically verified. The offline
suite is at 94 tests with the same five stale R20 failures; the production
preflight resolves 21 managed components read-only. Nothing is deployed,
`MACRO_SOURCE_REVISION` remains `r22` because no deployable behaviour
changed, and runtime output is identical to r22.

### Phase 2 — Rewrite feature qualification

All Phase 2 tasks are implemented in `Module12_FeatureQualification.bas`
(source-complete and statically verified; the P-0251 catalog in R23-213 still
needs one live run through `R23_ProbeFeatureCatalog`).

- [x] **R23-200:** Replace `ResolveFeatureType` with a three-field resolver:
  - `rawType2 = GetTypeName2`;
  - if `rawType2="ICE"`, `effectiveType = GetTypeName`;
  - otherwise `effectiveType = rawType2`;
  - if unresolved, fail with a field-specific reason.
- [x] **R23-201:** Log `rawType2`, `rawType1`, and `effectiveType` for every
  considered feature.
- [x] **R23-202:** Traverse relevant top-level features and subfeatures with a
  cycle guard.
- [x] **R23-203:** Evaluate suppression using each drawing view's exact
  `ReferencedConfiguration`; do not assume the active part configuration.
- [x] **R23-204:** Add typed readers for:
  - `HoleWzd`;
  - advanced holes;
  - simple/sketch holes;
  - resolved exact `Cut` and `CutThin` extrusions through
    `IExtrudeFeatureData2`;
  - modeled threads;
  - cosmetic threads;
  - supported pattern/mirror seeds.
- [x] **R23-205:** Follow each feature-data interface's access contract. Use
  `AccessSelections` before reading selection-backed data such as contours,
  directions, or end-condition references. Pair every successful access with
  `ReleaseSelectionAccess` on every success, rejection, error, and early exit.
  Do not put the model into rollback state solely to read scalar data when the
  interface contract does not require it.
- [x] **R23-206:** Never call `ModifyDefinition` and never save the fixture.
- [x] **R23-207:** Build feature ownership from `IFeature.GetFaces`; do not use
  `IFace2.GetFeature` as sole proof because it can report the oldest owner.
- [x] **R23-208:** Remove `FaceInSurfaceSense` as the internal-hole classifier.
  Retain it only as supplementary orientation evidence.
- [x] **R23-209:** Qualify only exact effective `CUT` or `CUTTHIN` types for the
  extrude-data route; do not use a broad substring match. Accept them as
  manufacturing cuts only when `IExtrudeFeatureData2`, `IsBossFeature=False`,
  owned
  cylindrical geometry, closed/circular projection, depth/end condition,
  configuration, and opening/access side are proved.
- [x] **R23-210:** Reject bosses, open/partial cuts, slots, fillets, chamfers,
  ambiguous cylinders, unsupported definitions, and unproved configuration
  states with explicit reason codes.
- [x] **R23-211:** Resolve patterns/mirrors through seed features and reject
  missing, circular, multiply resolved, or unsupported seed chains. The first
  live run showed the instance must also **inherit the seed's semantics** —
  `GetSeedFeature` gives identity only — or the instances form their own
  empty family. Seed chains are refused rather than followed recursively:
  termination is unproved on this build and P-0251 has no nested pattern.
- [x] **R23-212:** Consolidate coaxial owned cylinders from separate resolved
  cuts into a stepped-bore physical stack.
- [x] **R23-213:** Verify the expected P-0251 catalog (`VerifyExpectedCatalog`).
  Proved live 2026-08-01, `catalogFailures=None`:
  - [x] one six-location counterbore family — radii `0.0055`/`0.0033`, two
    unique X and three unique Y, matching the Phase 0 ordinate evidence;
  - [x] one four-location M5x0.8 tapped family — failed as
    `NoFourLocationFamily` on the first run because pattern instances did
    not inherit seed semantics (fixed under R23-211), passing on the second;
  - [x] one separate stepped-bore stack — radii `0.0235`/`0.0200`, one
    location, two stack members;
  - [x] no fillet, chamfer, boss, sketch, plane, or folder accepted as a
    hole — 40 rejections, every one with an explicit reason code.

**Phase 2 gate: SATISFIED (2026-08-01).** The graph contains source-backed
feature definitions and physical identities, proved live on P-0251 with
`catalogFailures=None`, before any drawing projection, ordinate, callout, or
section decision is made.

**Phase 2 status:** deployed, compiled, and run twice against P-0251. 28
contract tests cover R23-200 to R23-213; the offline suite is at 122 tests
with the same five stale R20 failures; the production preflight resolves 22
managed components. One static-only follow-up is outstanding: the traversal
key dropped `ObjPtr` after the two runs disagreed on `visitedFeatures`
(47 → 46) on an unchanged model. It changed no catalog output and needs no
dedicated run — confirm on the next live run that `visitedFeatures` is
stable and free of repeated names.

Three Phase 2 decisions worth carrying forward:

- **Suppression fallback.** This build returned Empty from
  `IFeature.IsSuppressed2` for every feature, so a fallback is unavoidable.
  It is permitted only when the requested configuration name equals the
  active one, in which case the two are equivalent; otherwise the feature
  fails closed as `ConfigurationSuppressionUnproven`. The state string always
  records which route produced it.
- **Advanced holes are read but not claimed.** `IAdvancedHoleFeatureData`
  exposes its sizes through element collections whose contract is not yet
  runtime-proved on this build, so the typed object is retained and the
  status is `ReadPendingElementContract` with no semantic field asserted.
  P-0251 does not use advanced holes.
- **A not-applicable code is not a value.** `IWizardHoleFeatureData2.HoleFit`
  returns `-1` on a tapped hole because the 2025 Help limits the property to
  counterbore and countersink features. Only the three
  `swWzdHoleScrewClearanceTypes_e` members are treated as data; anything else
  is recorded as absent. The same rule applies to any future field whose
  interface exposes it unconditionally but only defines it for some types.

The gate remains open until a second live `R23_ProbeFeatureCatalog` run
confirms the corrected catalog with `catalogFailures=None`.

### Phase 3 — Resolve drawing projections and selectable anchors

Implemented in `Module13_ProjectionResolution.bas` (27 procedures).

- [x] **R23-300:** Keep the two drawing-component contexts explicit:
  - use the limited drawing-context `Component2` returned by
    `GetVisibleComponents` only where required by `GetVisibleEntities2`;
  - obtain a full model-capable component through
    `GetVisibleDrawingComponents` and `IDrawingComponent.Component` for
    ownership/body operations;
  - runtime-prove how the two handles converge on the same represented
    component/configuration. **Implemented, live proof pending** — the
    convergence state (`Converged`, `Diverged`, `DrawingContextOnly`,
    `ModelContextOnly`, `NoComponentContext`) is logged per view. The 2025
    Help documents `GetVisibleDrawingComponents` for *assembly* drawings, so
    `DrawingContextOnly` is the expected result on a part drawing and is
    recorded as context, not failure. More than one visible component is
    refused rather than arbitrarily narrowed.
- [x] **R23-301:** Enumerate drawing-context edges with
  `GetVisibleEntities2(component, entityType)`.
- [x] **R23-302:** Map drawing entities back to model/configuration context for
  ownership proof — candidate edges come only from the location's own
  retained `SourceFaces`, never from the visible inventory at large.
- [x] **R23-303:** Use model-owned forward correspondence as the completion path
  for required hidden-line entities (route A). Route B is still attempted and
  recorded so a future build that fixes it is detected, never depended on.
- [x] **R23-304:** Require both paths to converge on the same physical
  location/feature ownership — route A's result is accepted only once route C
  identity-matches it into the view's visible inventory via `ISldWorks.IsSame`.
- [x] **R23-305:** Prove circular projection using the Phase 0-selected live
  call sequence: `ICurve.IsCircle` through the shared Boolean normalizer plus
  a `GetCurveParams3` range whose endpoints coincide.
- [x] **R23-306:** Prove model cylinder axis/radius, drawing projected
  centre/radius, referenced configuration, and normal-axis compatibility.
- [x] **R23-307:** Select only drawing-document entities in the drawing using
  `IEntity.Select4` and `ISelectData.View`, with the error-91 binding guarded
  and ownership proved afterwards through
  `ISelectionMgr.GetSelectedObjectsDrawingView2`.
- [x] **R23-308:** Select the best anchor in this order:
  - imported native-callout attachment — **tier defined, unavailable until
    Phase 4** attaches annotations; recorded rather than collapsed away;
  - primary typed hole diameter — the location's smallest coaxial radius;
  - smallest stable complete-circle drawing alias.

  Every mappable candidate is ranked; the pass does not stop at the first.
- [x] **R23-309:** Keep a physical location even when no anchor exists, but fail
  its required view projection as `ProjectionAnchorUnavailable`.
- [~] **R23-310:** Prove P-0251 has usable primary projections for all six face
  holes and usable side projections for all four tapped holes. Four live runs
  of `R23_ProbeViewProjections`; per-view results from the fourth:

  | view | type | projections | axisNormal | anchored | accepted |
  |---|---|---|---|---|---|
  | `Drawing View4` (primary) | 4 | 11 | 7 | 11 | **7** |
  | `Drawing View7` (side) | 4 | 11 | 4 | 6 | **2** |
  | `Section View J-J` | 2 | 11 | 4 | 0 | 0 |
  | `Drawing View2` (isometric) | 4 | 11 | 0 | 9 | 0 |

  - [x] **primary projections for all six face holes** — `Drawing View4`
    accepted all six counterbores plus the stepped bore, every one at
    `anchorTier=PrimaryTypedHoleDiameter` on the Ø6.6 through hole rather
    than the Ø11 mouth, and all seven proved selectable with
    `ownershipProven=True`.
  - [x] **side projections for the tapped holes — two, which is all the
    drawing contains.** The task text asked for four. That is not
    achievable, and not because of any mapping limitation: `Drawing View7`
    looks along model Y, and the four tapped holes lie on two axes with two
    holes each, so they project onto **two** page points, not four. The six
    counterbores likewise collapse from six to three. The observed
    `mappedEdges` counts (3 and 2) equal the number of distinct page
    positions exactly. Two coaxial holes seen along their axis are ONE
    circle on the sheet; SOLIDWORKS holds a single drawing entity and no
    search strategy can yield more anchors than the drawing has entities.
    `MarkCoincidentProjections` now attributes each such location to the
    anchored one it shares a page point with.

    The individual identity of all four is not lost: `Drawing View4`
    resolves them to four distinct page positions, edge-on. The reference
    drawing correspondingly calls them out once as
    `4x Ø4.2 ▼12.4 / M5x0.8 - 6H ▼10`.

**Phase 3 gate: SATISFIED (2026-08-01).** Every task R23-300 to R23-310 is
proved live on P-0251.

**Carry into Phase 5:** a required-coverage scheme must count *distinct page
positions per view*, not physical locations. Requiring one annotation per
physical location in a side view is unsatisfiable by construction.

Three Phase 3 decisions worth carrying forward:

- **`ISldWorks.IsSame` is not a Boolean.** It returns `swObjectEquality`
  = {0 not same, 1 same, **2 unable to determine**}. Reading it through
  `NormalizeSwBoolean` would accept "unable to determine" as proof of
  identity. Only an exact 1 is accepted, and unreadable comparisons default
  to not-same and are counted in the evidence.
- **Direction vectors are differenced, not transformed.** The axis is
  transformed as two points and subtracted, because applying a point
  transform to a direction folds in the view translation and would make
  every axis read as oblique.
- **The anchor must be chosen, not encountered.** Taking the first mappable
  edge would make the anchor depend on face and edge iteration order, so a
  counterbore could be dimensioned on its 11 mm mouth rather than its 6.6 mm
  through hole purely by traversal accident.

### Phase 4 — Rebuild model-item import and reconciliation

Implemented in `Module14_AnnotationImport.bas` (26 procedures).

**Mutation boundary.** This is the first phase whose operations change a
drawing. Only two procedures mutate — `ImportModelAnnotations` and
`RemoveR23CreatedAnnotations` — and both refuse unless passed an explicit
`allowMutation` argument. The evidence entry point
`R23_ProbeAnnotationReconciliation` never passes it and contains no insert,
delete or save call, so it can be run against the manual reference drawing
without altering it. That read-only path is not a reduced form of the work:
the reference drawing already carries the manufacturing intent R23 must
reproduce, so reconciling against it exercises R23-406 to R23-409, R23-411
and R23-412 on real data.

- [x] **R23-400:** Import constants verified member by member against the
  2025 `swInsertAnnotation_e` table. The Phase 0 mask `18055274` decomposes
  with **no unaccounted bit**: datums 2, dimensions 8, GTols 32, notes 64,
  marked-for-drawing 32768, hole-wizard profile 65536, hole-wizard location
  131072, hole callout 1048576, toleranced dims 16777216.
- [x] **R23-401:** `IsModelImportEligibleView` and `IsOrdinateEligibleView`.
- [x] **R23-402:** Primary, side and section are import-eligible; the sheet's
  standard-view placeholders are not, because they hold no entities.
- [x] **R23-403:** Section excluded by view type; isometric excluded because
  **no hole axis is normal to it**, read from the graph's own Phase 3
  measurement rather than from a view name. A renamed or reoriented view is
  therefore still classified correctly.
- [x] **R23-404:** `IsDeferredCreationView`, using the same measured test.
- [x] **R23-405:** Section first with `IMPORT_MASK_SECTION` (general, marked
  and toleranced dimensions, datums, GTols — deliberately no hole-wizard or
  callout bits), then the remaining views with `IMPORT_MASK_FULL`,
  `AllViews=False`, `DuplicateDims=True`.
- [x] **R23-406:** The insert's returned annotations are recorded as
  R23-created, **and** every view is independently traversed with
  `IView.GetAnnotations`, so coverage never depends on what an insert
  reported.
- [x] **R23-407:** Reconciliation is by attached-entity COM identity against
  each projection's proven anchor. Page proximity is not used anywhere:
  it would attach a dimension to whichever hole happened to be nearest,
  which is the exact failure the physical-location model exists to prevent.
- [x] **R23-408:** Categories by `swAnnotationType_e`, with hole callouts
  separated from ordinary dimensions via `IDisplayDimension.IsHoleCallout`,
  and ordinates accepting the live-proven type codes 1, 7 and 8.
- [x] **R23-409:** Tolerance type, bounds, fit type, hole fit and shaft fit.
- [x] **R23-410:** Deletion matches only this run's own recorded annotation
  objects. Nothing is matched by name, position or appearance, so
  pre-existing manual content cannot be selected.
- [x] **R23-411:** Deduplication by model-dimension identity plus owning
  view, on top of the API's `DuplicateDims` behaviour.
- [x] **R23-412:** Success is required-category coverage — native hole
  callout, toleranced dimension, ordinate — never a count.

**Phase 4 gate:** SATISFIED — six read-only live runs against the P-0251
reference drawing, `mutations=0`, `initialSelectionCount=0`,
`finalSelectionCount=0`, `drawingUnchanged=True` every time.

Closing evidence, final run: `annotations=38`, `coverageFailures=None`, and
`COVERAGE|holeCallouts=2|ordinates=10|diameters=0|toleranced=1|withFit=1`.
R23-412 is defined as required-**category** coverage — native hole callout,
toleranced dimension, ordinate — and all three categories are present and
read. That is the gate, and it is met.

**`reconciled=1` of 38 is not a gate failure, and chasing it further would
have been chasing the wrong number.** R23-407 requires reconciliation *by
attached-entity COM identity*; `RD1@Drawing View7` proves that mechanism
end to end. The other 37 are hand-authored reference dimensions
(`RD*@Drawing View*`, `D*@Sketch88`) whose attached drawing entities have no
reachable model counterpart:

- The forward map is partial — `IView.GetCorrespondingEntity` returns 2 of
  each location's 4 boundary edges — so those entities are not aliases.
- The reverse map is **empirically unavailable in a part drawing**:
  `IModelDocExtension.GetCorrespondingEntity2` returned Nothing for all 38
  annotations and every attachment, **with error 0**. It declines rather
  than fails. The member is documented as returning the entity "in the
  underlying part or subassembly"; a part drawing has no component to
  descend into, matching the `componentContext=DrawingContextOnly` this
  probe already records.

Unmatched annotations are therefore reported as
`AuthoredDrawingEntityNoModelCounterpart`, which is a statement about the
drawing rather than a failure of the ownership model. Reconciling R23's
**own** imported annotations — the mutating path R23-406/407/410 exist for —
is untouched by this and is Phase 5+ work.

Three Phase 4 contracts worth carrying forward:

- **`IDrawingDoc.InsertModelAnnotations4` takes eight arguments and returns
  an ARRAY of inserted `IAnnotation` objects**, not a count. There is no
  `InsertModelAnnotations3` on `IModelDocExtension`. Returning the objects is
  what makes R23-410 safe: identity is the only sound way to know what this
  run created.
- **`IDimensionTolerance.GetMinValue2`/`GetMaxValue2` return the STATUS** and
  deliver the value through an out parameter, which must be a local. Passing
  a class `Public` field would discard it — the same trap that produced
  `projectedAxis=0,0,0` in Phase 3.
- **`swTolFIT` and `swTolMETRIC` are both 7** in the 2025 enum, so a value of
  7 cannot be reported as one rather than the other and is recorded as
  ambiguous.

### Phase 5 — Create the required location schemes

Implemented in `Module15_OrdinateScheme.bas` (36 procedures) with two typed
records, `COrdinateScheme.cls` and `COrdinateBucket.cls`.

**Mutation boundary.** Exactly one procedure creates anything:
`CreateOrdinateGroup`, which refuses unless passed an explicit
`allowMutation`. `R23_ProbeOrdinateScheme` never passes it, contains no
`AddOrdinateDimension` call at all, and reports selection counts before and
after, so it can be run against the manual reference drawing.

**Status: read-only gate SATISFIED (three live runs).** Final run:
`schemes=4|horizontalSchemes=2|verticalSchemes=2|creditedLocations=10|`
`expectedLocations=10|coverageFailures=None|creations=0|`
`initialSelectionCount=0|finalSelectionCount=0|drawingUnchanged=True`.

Everything the probe can prove without creating a dimension is proven. What
remains open is exactly what requires mutation - R23-506 and R23-508.

- [x] **R23-500:** Scheme key is `view role + machining face + datum policy
  + direction`, every part measured. Four schemes resolved live, two
  horizontal and two vertical, across `Drawing View4` and `Drawing View7`.
  Machining face comes from the location's sign-normalized axis, so the
  top-face and side-face families landed in separate schemes without any
  view name being read.
- [x] **R23-501:** `Drawing View4` horizontal datum resolved to the stepped
  bore's projected centre (`x=0.207332`, source moment `0,0.062,0`) by the
  `CentreBoreProjectedCentre` policy rather than by fallback, and was proved
  selectable: `selection=True|selectedCount=1|ownershipProven=True` via
  `ISelectionMgr.GetSelectedObjectsDrawingView2`.
- [ ] **R23-502:** The vertical datum must be the lowest visibility-qualified
  horizontal model edge in each ordinate-bearing view, not a hole centre or
  the view rectangle. r29 Route D evidence is historical only: r30 live
  evidence found no mapped bottom edge in the visible-entity inventory, so no
  vertical datum is presently accepted. `IView.GetOutline` remains an
  enclosing view bound, not silhouette geometry.
- [x] **R23-503:** `Drawing View4` horizontal scheme resolved `buckets=2`,
  `anchoredBuckets=2`, at `x=0.222332` and `x=0.192332`, crediting three
  counterbores each - the two symmetric columns. Creation is R23-508.
- [x] **R23-504:** `Drawing View4` vertical scheme resolved `buckets=3`,
  `anchoredBuckets=3`, at `y=0.087415`, `y=0.127415` and `y=0.167415`,
  crediting two counterbores each - the three rows. Creation is R23-508.
- [x] **R23-505:** Proven live, and the coverage gate caught a real defect
  doing it. The four side holes occupy two page positions in
  `Drawing View7`; `ORDINATE_COINCIDENT_CREDIT` now credits all four while
  the horizontal scheme holds a single bucket. First run reported
  `credited=8, expected=10` because the coincidence link was read from the
  anchored end, where it is never set.
- [ ] **R23-506: half met.** All four side tapped-hole locations are
  **resolved and credited** - `creditedLocations=4` in both `Drawing View7`
  schemes, with a proven selectable anchor. They are not yet **dimensioned**;
  that is the mutating half and is unrun.
- [x] **R23-507:** `profileEntries=1` in both `Drawing View4` schemes - the
  stepped bore is held apart from the small-hole ledger and never enters a
  bucket. Membership is family size read from the graph
  (`rule=FamilySize>=2`), not a radius threshold: the bore is excluded for
  being a singleton, not for being large.
- [ ] **R23-508:** For each ordinate transaction:
  - activate and verify the view;
  - bind `ISelectData.View`;
  - select the datum first;
  - append each entity explicitly in deterministic order;
  - verify selection order/count;
  - call `AddOrdinateDimension`;
  - decode the complete result;
  - call `SetPickMode`;
  - clear selections on every exit;
  - enumerate and verify created display dimensions.
- [x] **R23-509:** `smallHoleLocations=10|totalLocations=11|families=3`, and
  `creditedLocations=10|expectedLocations=10|coverageFailures=None` with two
  horizontal and two vertical schemes. Coverage is counted per distinct page
  position and credited to locations, per the Phase 3 finding.

### Phase 6 — Reconcile native callouts and controlled fallback

Implemented in `Module16_CalloutDefinition.bas` (20 procedures) with the
typed record `CCalloutDefinition.cls`.

**Mutation boundary.** One procedure creates anything:
`CreateNativeCalloutForFamily`, which refuses without `allowMutation` and
again without a proven anchor. `IDrawingDoc.AddHoleCallout2` attaches to
whatever edge is selected, so selecting an unproven entity would produce an
associative callout pointing at the wrong hole and looking correct on the
sheet. `R23_ProbeCalloutDefinition` contains no `AddHoleCallout2` call.

**Phase 6 read-only gate SATISFIED** (three live runs). Final run:
`definitions=3|definitionFailures=None|counterboredFamilies=1|`
`threadedFamilies=1|shapeFailures=None|nativeCallouts=2|creations=0|`
`initialSelectionCount=0|finalSelectionCount=0|drawingUnchanged=True`.

Everything provable without creating an annotation is proved. R23-604's
creation half is unrun because it mutates.

- [x] **R23-600:** Always request native callouts when the configuration enables
  them.
- [x] **R23-601:** `nativeCallouts=2` classified by
  `IDisplayDimension.IsHoleCallout` alone. A native callout reports
  `Type2 = 6` and so does an ordinary diameter dimension, so no
  dimension-type constant is declared in the module at all.
- [x] **R23-602:** One of the two callouts attributed by COM identity;
  the other reports `NoOwningProjection` because its edge is none of the
  forward aliases - the Phase 4 limitation, recorded rather than worked
  around. A callout resolving to two families is rejected, not tie-broken.
- [x] **R23-603:** Read from `GetHoleCalloutVariables` per
  `ICalloutVariable`, never from rendered text. The live run printed all
  four M5 variables with their `swCalloutVariable_e` types: Tap Drill
  Diameter 29, **Tap Drill Depth 28 = 0.0124**, Thread Description 41,
  Thread Depth 32 = 0.010.
- [ ] **R23-604: read-only half met.** The M5 family retained its native
  callout with `reason=CompleteAssociativeDefinitionAvailable`. Creating a
  native callout for a family that lacks one is `CreateNativeCalloutForFamily`
  and is unrun, because it mutates.
- [x] **R23-605:** Proved in both directions live. The M5 family kept its
  native callout; the other two retained `ControlledFallback` with
  `reason=NoNativeCalloutAttributedToFamily`. An earlier run also exercised
  `reason=NativeIncomplete|nativeMissing=Depth`.
- [x] **R23-606:** `quantity=6`, `4` and `1` from
  `CLocationGraph.LocationsForFamily.UniquePhysicalLocations`. Six
  counterbores come from one Hole Wizard feature and four side holes from
  one feature plus a mirror, so neither a feature count nor an edge count
  would have produced these.
- [x] **R23-607:** `cboreDiameterM=0.011`, `cboreDepthM=0.006`,
  `thread=M6` and `thread=M5x0.8` all carried from `CFeatureDefinition`
  with their proof sources. Nothing parsed from a feature name.
- [x] **R23-608:** The counterbore family's `holeFit=Normal` carries
  `toleranceSource=IWizardHoleFeatureData2.HoleFit=1`. Families with no
  source fit report `Unproven` rather than an invented value, and no
  drawing-authored tolerance is promoted here.
- [x] **R23-609:** The R23 production route uses Module16 family definitions;
  `Module7_TitleBlockEngine.PopulateTitleBlock` remains title-only. The old
  fixed P-0251 callout strings and feature-name/radius selection are absent
  from the reachable route. The r28 scratch runner reads
  `legacyBoreCallout=0|removalStatus=PipelineSwitched`.
- [x] **R23-610:** Proved by the failures it produced before the fixes
  landed - `NominalDiameter`, `Depth`, `CounterBoreDepth`, `Attachment`,
  each named individually. Depth is demanded only from a blind hole
  (`swEndCondBlind = 0`), so a through hole is never asked for one.
- [x] **R23-611:** `counterboredFamilies=1|threadedFamilies=1|`
  `shapeFailures=None` - the six-hole counterbored family and the four-hole
  tapped family, both complete and both attachment-proved. The rule is
  stated as shapes rather than part numbers.

  The first passing run counted **two** threaded families: a thread
  description alone was being treated as a thread, and the counterbored
  clearance-hole family carries the fastener size of the screw it clears
  with `threadDepthM=0`. A hole that is actually tapped has a thread depth.

### Phase 7 — Rebuild J-J from model intent

Implemented in `Module17_SectionPath.bas` (21 procedures) with the typed
record `CSectionPath.cls`.

**Mutation boundary.** One procedure creates anything:
`CreateSectionFromPath`, which refuses without `allowMutation` and refuses
again unless the path resolved with its crossings proved (R23-708). The
section view's placement is a caller argument, not a default: defaulting it
to a point on the source view would stack the section on the view it was cut
from, and choosing where a view sits is layout, a later phase.

**Phase 7 read-only gate SATISFIED** (first live run). Final run:
`resolvedPaths=1|segments=3|columnHoles=3|crossingsProven=4|`
`sectionFailures=None|creations=0|initialSelectionCount=0|`
`finalSelectionCount=0|drawingUnchanged=True`.

Everything provable without cutting a section is proved. R23-705 to R23-707
are proved to the limit a read-only run allows: the frame conversion ran and
was checked, but no line was drawn and no view was cut.

- [x] **R23-700:** Split graph/projection discovery from ordinate creation so
  section construction can consume proved locations first.
- [x] **R23-701:** In `Drawing View4`, `distinctColumns=2` and
  `distinctRows=3` - the 2x3 face-hole family - with the stepped-bore
  centre resolved as the largest singleton-family location. The other three
  views correctly report `NoAcceptedSingletonBoreProjection`, because the
  bore is not accepted there.
- [x] **R23-702:** Built live, in the approved order:
  `w1=0.207331779,0.237414746` (bore centre);
  `w2=0.207331779,0.167414746` (same X, highest row);
  `w3=0.192331779,0.167414746` (minimum-X column, that row);
  `w4=0.192331779,0.087414746` (same column, lowest row).
- [x] **R23-703:** `crossingsProven=4|columnHoles=3|failures=None` - the
  bore plus all three holes on the chosen column. Each is judged against
  its own projected radius, and the point-to-segment distance is clamped to
  the finite segment.
- [~] **R23-704:** P-0251's R23 route resolves J-J through Module17; the
  retired P-0251 section path fails named rather than creating a section.
  The older generic fallback remains for fixtures without a graph-backed
  semantic section and is not production-proven.
- [ ] **R23-705: conversion proved, creation unrun.** The read-only frame
  probe returned `pageX=0.207331779 -> viewX=-0.022000000` and
  `pageY=0.237414746 -> viewY=0.062000000` against
  `originX=0.229331779|originY=0.175414746|scale=1|angle=0`, which is exact
  arithmetic.

  It also cross-checks against the model, which matters more than the
  arithmetic: the bore's Plucker moment is `(0, 0.062, 0)` and its `viewY`
  is 0.062 exactly. Every counterbore behaves the same - `viewY` equals its
  moment's Y (-0.008, -0.048, -0.088), and `viewX` equals its moment's X
  minus a constant 0.022, the view's own centring offset, identical across
  all seven holes. A wrong transform does not produce one shared offset
  across seven independent points.
- [ ] **R23-706: unrun.** It creates sketch segments. The selection-order
  verification is written and contract-tested but has never executed.
- [ ] **R23-707: unrun.** There is nothing created to read back yet.
- [x] **R23-708:** Proved by the three views that failed. Each reported
  its own specific reasons - `NoAcceptedSingletonBoreProjection`,
  `NoBoreProjection`, `NoColumnHoles` - and no path was approximated into
  existence.

  The first run also appended `NotAttempted`, the crossing proof's initial
  STATE, alongside those reasons. A state is not a failure and reporting it
  as one dilutes the reasons that are real, so it is now excluded.

### Phase 8 — Add a dedicated section-dimension engine

Implemented in `Module10_SectionDimensionEngine.bas` (28 procedures) with
the typed record `CSectionRequirement.cls`.

**Mutation boundary.** Two procedures change anything: `ApplyReferenceFit`
and `CreateSectionDimension`. Both refuse without `allowMutation`, both
record the mutation on the evidence ledger, and both read the result back
rather than trusting a return value. `R23_ProbeSectionDimensions` calls
neither and contains no `AddDimension2`, `SetFitValues` or `SetValues2`.

**The record keeps REQUIRED and OBSERVED apart.** `CSectionRequirement`
splits what the specification demands from what was read back from the
drawing, and nothing writes an OBSERVED field from a REQUIRED one. A
requirement that reports its own nominal back as the observed nominal proves
nothing, and a contract asserts the assignment never happens.

**Phase 8 gate SATISFIED** (third live run, 2026-08-02):
`satisfied=7|missing=0|duplicated=0|sectionDimensions=7|sectionOrdinates=0|`
`requirementFailures=None|creations=0|mutations=0|drawingUnchanged=True`,
selection 0 before and after.

Both bore requirements resolved through `diameterDisplaySource=TextPrefix`
with `prefix=<MOD-DIAM>`. The drawing does show the diameter symbol; it
carries it in the dimension's text prefix rather than in the diametric
flag, which is exactly the case the third reading of R23-804 was added to
distinguish. Worth recording for later work: BOTH `GetText` forms returned
the literal `<MOD-DIAM>` token rather than a rendered glyph, so the
codepage-216 comparison never fires on this build and the token match is
what decides.

Historical status of the run before it:
The second run returned `satisfied=5|missing=0|duplicated=0` with every
nominal exact - 0.018, 0.012, 0.0115, 0.040, 0.047, 0.1736, 0.1048 - so all
seven requirement keys exist in the drawing and are matched.

The two flagged are the bore diameters, on `NotDisplayedAsDiameter:2`. Every
section dimension returned `diametric=False` with `diametricKnown=True`, a
real answer rather than a read failure. Before that can be called a defect
in the drawing, the dimension's TEXT PREFIX has to be read: a drawing can
carry the diameter symbol there while the diametric flag stays False, and
then the sheet reads correctly even though the record does not. That read is
now in place. Awaiting the third run.

- [x] **R23-800:** Added `Module10_SectionDimensionEngine.bas`. The plan
  named the free `Module10` slot and it is taken as named; the other R23
  modules run 11 to 17.
- [x] **R23-801:** All seven are defined in `BuildSectionRequirements`,
  each carrying its nominal in metres, its accepted dimension types, and its
  R23-808 lane. Only the fit bore carries a tolerance - a contract asserts
  `ToleranceRequired = True` appears exactly once, because inventing a
  tolerance for the other six is exactly what the standing policy forbids:
  - overall thickness `18.00`;
  - bore-step depth `12.00`;
  - inner bore `Ø40.00`;
  - fit bore `Ø47 H7 +0.025/0.000`;
  - lower wall/step `11.50`;
  - long vertical reference `173.60`;
  - lower vertical reference `104.80`.
- [x] **R23-802:** `ReconcileSectionDimensions` runs before any creation
  path and records all six observations on the requirement: source dimension
  identity (`IDimension.FullName`), attached geometry
  (`IAnnotation.GetAttachedEntityTypes`), semantic role, nominal
  (`IDimension.GetSystemValue3`), type (`IDisplayDimension.Type2`) and
  tolerance (`IDimension.Tolerance`). Nominal and accepted type together
  decide the match; the rest are recorded so the match can be audited.

  Every match is counted rather than only the first, because R23-811 has to
  be able to fail on a duplicate instead of quietly dimensioning the same
  thing twice.

  Per-dimension locals are reset on every iteration. VBA block-scoped locals
  live for the whole procedure, and that is precisely how the Phase 0
  section inventory mislabelled eleven of its seventeen dimensions.
- [ ] **R23-803: written, unrun.** `CreateSectionDimension` refuses when
  the requirement already matched an imported dimension, refuses when the
  caller has selected no entities, and verifies the created dimension's
  nominal and type by read-back before calling the requirement satisfied. A
  dimension that exists but measures the wrong thing is recorded as
  `CreatedButRejected`, because a wrong dimension looks finished.

  Entity selection is the caller's job by API contract, not by preference:
  `IModelDoc2.AddDimension2`'s Remarks state that entities are selected by
  LOCATION and never by name, since passing a name makes the dimensioning
  routines pick a line endpoint at random.
- [x] **R23-804: revised by the live run.** Dimensions are read through
  `IView.GetDisplayDimensions`, which is view-scoped; the obsolete
  `GetFirstDisplayDimension5` walks the whole sheet by its own Remarks and
  would attribute another view's dimensions to the section.

  The type rule changed. Phase 0 saw type-6 `D1@Sketch4`/`D1@Sketch6` in
  this fixture; the 2026-08-01 run found a different drawing state - seven
  dimensions in `Section View J-J`, every one `swLinearDimension = 2` and
  named `RD1..RD7`, drawing-authored reference dimensions rather than
  imported model dimensions. Both states are real, so a diameter
  requirement now accepts type 6, type 15 **and** the linear types.
  Requiring any single type would reject one of the two real states, which
  is exactly what R23-804 exists to prevent.

  Which of them the drawing DISPLAYS as a diameter is a separate question
  with three possible answers, and `DiameterDisplaySource` names the one
  that applies: `DiametricRecord` (the type says so), `DiametricFlag`
  (`IDisplayDimension.Diametric` says so), or `TextPrefix` (the symbol is
  in the dimension's text, read through
  `GetText(swDimensionTextPrefix)` and its `...PrefixDefinition` form,
  where SOLIDWORKS writes `<MOD-DIAM>`). Only when all three decline is
  `NotDisplayedAsDiameter` recorded.
- [x] **R23-805: corrected after the live run.** The nominal read was the
  defect that blocked the whole phase. `GetSystemValue3` with
  `swThisConfiguration` returned nothing for all seven dimensions - a
  drawing-authored reference dimension has no configuration to ask about -
  and without a nominal nothing can match, so every requirement reported
  Missing while its dimension sat in the view.

  `TryReadNominal` names the route that answered, and the second run
  settled it: all seven returned `nominalRoute=Obsolete.GetSystemValue2`
  with exact nominals. `swAllConfiguration` was reached and declined on
  every one, so it has been REMOVED - a route with live evidence against it
  is not kept "just in case". Two remain: `GetSystemValue3` with
  `swThisConfiguration`, which is the supported call and what answers for
  imported model dimensions like `D1@Sketch4`, and the obsolete
  `GetSystemValue2`, labelled obsolete in the route name, which is the only
  thing that answers for a drawing-authored reference dimension on this
  build. When both decline the raw shape of the `GetSystemValue3` result is
  reported, because "no nominal" and "an empty SafeArray" are different
  problems.

  Everything else is read through supported members only. All four
  `IDimension` tolerance members - `GetToleranceValues`,
  `SetToleranceValues`, `GetToleranceFitValues`, `SetToleranceFitValues` -
  are marked obsolete by the 2025 Help, each superseded by an
  `IDimensionTolerance` member, and a contract asserts none of them appears.

  `GetMinValue2` and `GetMaxValue2` return a STATUS and hand the value back
  through a ByRef argument, so the status is reported beside the value it
  qualifies. A zero value with a failed status is not a zero tolerance.
- [x] **R23-806:** H7 authority is **resolved**. The 2026-07-31 corrected-probe
  run read the part source directly and proved H7 is absent:
  `D1@Sketch4@P-0251-14A-001.Part`, nominal `0.047 m`, `toleranceType=0`,
  `fitType=-1`, empty hole/shaft fit and empty `GetToleranceFitValues`.
  The user selected **controlled target-spec/reference authority**.
  Implementation rules:
  - create or reuse an associative `Ø47` section dimension and apply
    `H7 +0.025/0.000` from the approved reference specification;
  - record provenance explicitly as target-spec/reference authority, never as
    model data;
  - QA must state that the tolerance did not originate in the model; and
  - R23-807 still applies: never substitute free text for a failed real
    dimension.

  Implemented as stated. `REFERENCE_HOLE_FIT = "H7"`,
  `REFERENCE_FIT_MIN_M = 0`, `REFERENCE_FIT_MAX_M = 0.000025` and
  `REFERENCE_FIT_AUTHORITY = "TargetSpecReferenceAuthority.NotModelData"`
  are stated once each. `ApplyReferenceFit` sets `swTolFITWITHTOL = 8`
  BEFORE the values, because `SetValues2` refuses while the type is
  `swTolNONE` by its own Remarks and `FitType` is only available for the fit
  types. It then reads the tolerance back rather than trusting the two
  return values, and normalizes both COM booleans.

  `EvaluateTolerance` will not claim model provenance for a tolerance it
  merely found on the drawing: present-on-drawing is recorded as
  `PresentOnDrawing.TargetSpecReferenceAuthority.NotModelData`, because
  Phase 0 proved the part-source dimension carries none.
- [x] **R23-807:** The module contains no `InsertNote`, no `CreateText` and
  no `SetText`, and every failure exit from `CreateSectionDimension` carries
  `policy=NoFreeTextSubstitute`. A note is not a dimension: it does not move
  with the geometry and cannot be inspected as one.
- [ ] **R23-808: lanes assigned, placement deferred.** Each requirement
  carries its lane NAME; a lane is not a coordinate. Turning a lane into a
  page position needs the finished section's annotation envelope, which is
  Phase 9's job - deciding it here would repeat the Phase 7 mistake of
  letting a module that cannot see the layout choose placement. The
  assignment is:
  - 18 and 12 above;
  - 11.5 below;
  - Ø40 and Ø47 on opposite bore sides;
  - 173.6 and 104.8 on separate exterior vertical lanes.
- [x] **R23-809:** `Module4_ModelItemImporter.AutoArrangeAllDrawingDimensions`
  consults `IsExcludedFromGenericArrangement`; section views are skipped from
  generic arrangement. The r28 section probe reports that exclusion.
- [x] **R23-810:** `DetectLegacyBoreCallout` remains evidence-only and the r28
  scratch runner reads `legacyBoreCallout=0|removalStatus=PipelineSwitched`.
  It does not remove notes during a read-only probe.
- [x] **R23-811:** `VerifySectionDimensions` fails on a missing key, on a
  duplicate, on an unsatisfied tolerance, on an unattached dimension, and on
  any requirement whose own failure list is non-empty - a requirement that
  recorded a problem is not satisfied, whatever the counts say. It also
  counts ordinate types 1, 7, 8 and 16 in the section separately. An
  ordinate in a section shares no datum with the Phase 5 groups and reads as
  a coordinate from an origin the section does not have.

### Phase 9 — Make final layout content-envelope aware

Implemented in `Module18_ContentEnvelope.bas` (33 procedures) with the typed
record `CContentEnvelope.cls`.

**Mutation boundary.** One procedure changes anything: `ApplyPlacementPlan`,
which refuses without `allowMutation`. `PlanPlacement` computes target
centres and moves nothing, so the decision and the action are separable and
the plan is fully inspectable before a view has been touched.

**Frames.** Every coordinate in a `CContentEnvelope` is PAGE frame, and the
sources do not agree on that by default:

| Source | Frame |
|---|---|
| `IView.GetOutline` | page, documented |
| `IAnnotation.GetPosition` | sheet-relative in drawings, documented |
| `INote.GetExtent` | sheet space, documented |
| `IView.GetSectionLineInfo2` | VIEW-SKETCH, proved by Phase 0 |
| `IDisplayData` points | **not stated by the Remarks** |

Section geometry is converted through `ViewSketchToPage`, the exact inverse
of Module17's `PageToViewSketch`, and the pairing is round-trip checked
before any point is contributed. Display-data points are contributed and
their agreement with the view's own documented outline is COUNTED rather
than asserted, because a contract the Help does not make is not one this
project states.

**Status after the third live run (2026-08-04, under the probe runner):**
the five defects below are fixed and confirmed - the section grammar now
parses (`arrow` and `section` sources contribute), no line printed twice,
and rejections carry a coordinate. The probe reported `envelopes=4|`
`annotationEnvelopes=3|protectedRegions=4|clearanceChecks=22` and
`plan=RescaleRequired|suggestedScaleFactor=0.527974|`
`requiredHeightM=0.479190|usableHeightM=0.253000` - the same arithmetic
that justified reversing R23-907, now measured rather than argued.

**Superseded by user decision (2026-08-04).** R23-903/R23-904 automatic
envelope movement and rescaling are complete for planning purposes. r31 keeps
the initial structural grid but cannot call `ApplyR23ContentLayout` from the
production route or the retired scratch command. `FINAL_LAYOUT` is reported as
`UserAcceptedLayoutAsIs|automaticClearance=DeferredByUser`, never as a
clearance pass. The diagnostic probe remains available to measure a layout,
but its result does not block the remaining semantic work.

**Historical automatic result.** Ten clearance failures stood
(`ViewOverlap` on four pairs, `ProtectedIntrusion` against the content
border on all four views). The r28 scratch macro allows two measured rescale
passes before shared placement. It has exact-path and referenced-part guards,
but `R23-903`/`R23-904` remain unrun pending the manual VBE compile and
mutating execution. The three
high-severity review findings in that path are source-corrected: placement
refuses an unusable plan, scaling refuses sheet/parent-dependent views, and
sheet measurement reads title-block extents when present. Their mutating
runtime evidence remains open.

Historical status of the run before it:

**Second live run completed end to end; five defects fixed.** The
sheet measured (`A3 0.420 x 0.297`, `contentBorder=Measured`,
`titleBlock=Absent`), four envelopes built, clearances checked, a plan
produced. Five defects the run exposed are fixed and it awaits a third run:

1. **The section-line arrow block is 9 doubles, not 11.** `Drawing View4`
   returned `items=49`, and `49 = 2 header + 1 numSegments + 7x3 segments +
   9 + 9 arrows + 7 text` - three segments, the J-J path exactly. With 11
   nothing matched, which is why `arrow=0|section=0` on every envelope.
2. **A view with no section line is not a failed parse.** An empty array was
   reported as `Unmatched` and appended to `sourceFailures` on every
   envelope; it is now `sectionGrammar=NoSectionLine`.
3. **Every envelope line printed twice** - the same `AddInfo`-already-prints
   defect fixed in Phase 8 and missed here.
4. **The display-data frame check allowed 120 mm of slack** and tested only
   line start points, so its 26/28 "consistent" counts were weaker evidence
   than they looked. Now 24 mm, both endpoints.
5. **Rejected points were counted without being sampled.** 34 rejections on
   the section view with no coordinate makes a frame error and genuinely
   off-sheet geometry indistinguishable; the first rejection is now kept
   with its source.

The eleven `*Front`/`*Top`/`*Isometric` template entries `GetViews` returns
are now skipped by name and logged as skipped.

- [x] **R23-900:** All eight sources contribute, each counted separately so
  an outline-only rectangle cannot pass as a content envelope
  (`HasAnnotationContent`). Two traps handled explicitly:
  `IDisplayData.GetTextPositionAtIndex` is an OFFSET from the display-data
  origin, not a coordinate - used absolutely it drags every envelope towards
  the sheet origin - and leader points are consumed as XYZ triples from the
  returned array rather than from `GetLeaderStyle`, whose value is OR-ed
  with attachment bitmask flags that the corpus returns mangled. The
  envelope sources are:
  - model outline;
  - display-dimension primitives and text;
  - note extents;
  - leader points;
  - callouts;
  - section segments;
  - arrow geometry;
  - J-label positions and text heights.
- [x] **R23-901: corrected after the live run.** `MeasureSheetRegions`
  measures the sheet read-only: `ISheet.GetSize` for the size, and
  `ISheet.GetZoneMargin` for the content border. An absent title block, an
  unread one and unmeasurable margins are each REPORTED - `titleBlock=Absent`,
  `titleBlock=PresentBoundsUnread`, `contentBorder=Unmeasured`,
  `usableSource=SheetExtentNoBorder` - and only an unusable sheet SIZE stops
  the probe.

  `BuildProtectedRegions` gates every rectangle on bounds that were actually
  measured. Emitting one from unset evidence fields would put a degenerate
  rectangle at the sheet origin and report false `ProtectedIntrusion`
  violations against it: a boundary that does not exist is not a boundary at
  zero.

  The content border is protected as four STRIPS rather than one rectangle,
  because the drawable area is inside it and a single rectangle would
  declare every view a violation.
- [x] **R23-902:** `PlanPlacement` packs rows from the envelopes' own
  widths and centres the whole block in the usable rectangle. A contract
  asserts none of `topBoundary -`, `Bias`, or `rowCenterY` survives - the
  tokens `Module9_LayoutEngine.bas:442-446` uses to pin the source row to
  the top boundary. A row pinned to a boundary has nowhere to put the
  annotations that hang above it, which is the whole defect.
- [~] **R23-903: r28 scratch entrypoint executed.** Two scale mutations were
  applied and saved; no position delta was applied because the re-measured
  plan never became acceptable. The delta is
  `cellCentre - envelope.Centre`, applied to `IView.Position`. A contract
  asserts `GetOutline` does not appear in that procedure at all, because
  moving by the OUTLINE centre is what `Module9_LayoutEngine` does and is
  exactly why annotations end up outside the region the layout believed it
  was filling.
- [~] **R23-904: r28 scratch entrypoint rebuilt and re-measured after each
  scale pass.** The saved drawing was then re-verified by the read-only
  runner; five conservative clearance failures remain. `EditRebuild3`, then
  every envelope is rebuilt from scratch, then clearances are re-verified;
  `MAX_CORRECTION_PASSES = 1`. A view move relocates the section line and
  every attached annotation, so nothing read before the move still describes
  the sheet. A failed rebuild is reported rather than swallowed.
- [x] **R23-905:** `VerifyClearances` checks every view-view pair and
  every view-protected pair, naming each violation, and reports the check
  count so a silently empty loop cannot read as a pass. Clearance is a
  separating-axis measure on both axes, so touching rectangles score zero
  rather than passing.
- [x] **R23-906:** `SECTION_CLEARANCE_M = 0.002`, applied by
  `RequiredClearance` when the view is `swDrawingSectionView = 2` and the
  other rectangle is protected. The section envelope already includes the
  arrows and both label points with their text height, so the 2 mm is
  measured from the geometry that actually overshot.
- [x] **R23-907: REVERSED BY THE USER on 2026-08-01** - "The views are
  allowed to rescaled as per need". The prohibition is replaced by a gate
  and a record, not by silence:

  - the only `ScaleDecimal` assignment in the module is inside
    `ApplyScaleToFit`, which refuses without `allowMutation`, records the
    mutation, and reads each new scale back rather than assuming the set
    took;
  - the factor is an ESTIMATE and says so
    (`factorIs=GeometricEstimateTextDoesNotScale`). Annotation text height
    does not scale with the view, so a view at half scale does not have half
    the envelope. The factor is applied, the drawing is rebuilt, and the
    envelopes are re-measured; and
  - `CaptureViewScales` photographs every approved scale first and
    `ReportScaleChanges` names every view whose scale changed, with the
    before and after values.

  The reversal was needed because the accepted reference drawing itself
  cannot satisfy the old rule: its four view envelopes need 0.479 m of
  height in the 0.253 m available.
- [x] **R23-908: survives the R23-907 reversal.** `PlanPlacement` now
  returns `plan=RescaleRequired` with the suggested factor and the required
  and available width and height. Rescaling is a permitted remedy, not an
  unlimited one: it happens once, the envelopes are re-measured, and if the
  content still does not fit, `ApplyPlacementPlan` returns
  `layout=Reject|reason=LargerSheetRequired`. The sheet being too small is
  still an answer this engine is allowed to give.
- [x] **R23-909:** `SealLayout` photographs the evidence ledger's
  mutation sequence when layout completes and
  `VerifyNothingCreatedAfterLayout` compares it afterwards, reporting the
  sealed and current sequence and the last operation on each side. Anything
  created after the seal changes the envelopes the layout was proved
  against.

### Phase 10 — Replace count-based QA

Implemented in `Module19_SemanticQA.bas` (25 procedures). No typed record:
every artefact it judges already has one.

**Strictly read-only, and unusually so.** This module creates, moves,
deletes and selects nothing at all - not even behind an `allowMutation`
gate. A QA engine that repairs what it is judging cannot report on it, and
a contract asserts the absence of every mutating call the other phases own.

**Why counts were not enough.** The checks this replaces ask "did anything
get imported?" and "does the note contain the expected text?". A nonzero
import count is satisfied by importing every dimension into one view and
none into the others. A note-token check is satisfied by free text that has
drifted from the geometry beside it, and fails on a correct drawing whose
wording differs. Neither says anything about the part.

**Pipeline wiring is Phase 11 work and is now present.** `Module6_QAEngine`
requires the shared semantic stages on the reachable production path.

**Status: r26 read-only runner, 2026-08-04.**
`R23_SEMANTICQA_END|stages=10|proved=9|failures=1|warnings=8`, with
`mutations=0` and `drawingUnchanged=True`. The sole failing stage is
`FINAL_LAYOUT`, whose ten named clearance/protected-region failures require
the separately authorized Phase 9 mutation path.

The run initially exposed a semantic-QA order defect: directional coverage was
recorded while building ordinate schemes, but `MODEL_IMPORT_COVERAGE` was
judged before those schemes. `COrdinateBucket` now records coverage by
direction, and Phase 10 rebuilds schemes before judging that stage. The r26
log proves `Drawing View4` at `coveredX=7|coveredY=6` and `Drawing View7` at
`coveredX=4|coveredY=4`; authored dimensions did not count as import proof.

- [x] **R23-1000:** Required stage `MODEL_INTENT_CATALOG`.
- [x] **R23-1001:** Required stage `MODEL_IMPORT_COVERAGE`.
- [x] **R23-1002:** Required stage `NATIVE_CALLOUT_COVERAGE`.
- [x] **R23-1003:** Required stage `PHYSICAL_LOCATION_GRAPH`.
- [x] **R23-1004:** Required stage `VIEW_PROJECTION`.
- [x] **R23-1005:** Required stage `ORDINATE_SCHEME`. Proved when
  every scheme resolved its datum and has buckets; a count of created
  dimensions proves nothing, because two groups can be created against
  the wrong datum and still count as two.
- [x] **R23-1006:** Required stage `SECTION_GEOMETRY`.
- [x] **R23-1007:** Required stage `SECTION_DIMENSIONS`.
- [x] **R23-1008:** Retained and strengthened: a definition counts
  only when it is complete AND attached to real drawing geometry. Text
  that describes the part correctly but hangs off nothing is a caption,
  not a manufacturing definition.
- [x] **R23-1009:** Final stage `FINAL_LAYOUT`.
- [x] **R23-1010:** `EvaluateModelImportCoverage` reports each view's own
  accepted projections and how many are covered in X, covered in Y and carry
  attached annotations. A view with accepted projections and none of the
  three fails by name (`ViewImportedNothing`), and a sheet where no view has
  accepted projections fails outright.
- [x] **R23-1011:** `EvaluateSectionDimensions` calls
  `Module10_SectionDimensionEngine.VerifySectionDimensions`, which judges
  type, nominal, attachment and tolerance. A contract asserts `GetText`,
  `GetNotes` and note-text searching appear nowhere in the module.
- [x] **R23-1012:** Every field is emitted beside its proof source, and
  a value with no source behind it fails `NoProvenance`. Blank, `None` and
  `Unproven` are treated as the same thing: nobody can say where the number
  came from. "6.6" is correct or wrong depending on whether it was read from
  the feature, read from a callout variable, or assumed.
- [x] **R23-1013:** `QA_FEATURE_TYPE` carries `RawTypeName2`,
  `RawTypeName`, `EffectiveType`, `TypeResolutionSource`, `OperationKind`
  and the rejection reason for EVERY audited feature. Every feature rather
  than every accepted one, deliberately: a rejection with no type recorded
  cannot be reviewed, and "it was rejected" is not a reason.
- [x] **R23-1014:** `TypeResolutionFailed` fails an empty effective type
  and the three resolution outcomes `Module12_FeatureQualification` actually
  writes - `IceUnresolved`, `Unresolved` and `ReadError`. A contract checks
  those literals against Module12 rather than trusting the copy here.
- [x] **R23-1015:** `EvaluateViewProjection` names every
  identity-proven location with no accepted projection anywhere
  (`NoProvedProjection`). Reporting the projection COUNT would hide it
  behind the locations that did project.
- [x] **R23-1016:** `DuplicateKeyReport` names every repeated key and
  how often, applied to physical location keys, family definition keys and
  section requirement keys. It returns `"None"` rather than an empty string,
  so "no duplicates" and "the check did not run" cannot be confused at a
  call site.

  Two locations sharing one physical key means the graph has lost track of
  which hole is which, and everything downstream that looks a location up by
  key silently gets the wrong one.
- [x] **R23-1017:** `EvaluateSectionGeometry` fails a missing path, an
  unresolved path and any crossing failure; `EvaluateFinalLayout` fails an
  envelope that could not be seeded (`EnvelopeUnavailable`) and any
  clearance failure. An envelope that could not be built is an unsafe
  layout, not a missing number.

### Phase 11 — Reorder the production pipeline

- [x] **R23-1100:** Replace the initial `modelHoleFeatures` collection with the
  location graph.
- [x] **R23-1101:** Implement this order:
  1. audit referenced configuration and build feature definitions;
  2. create primary and side views;
  3. perform rough geometry placement;
  4. resolve view projections;
  5. create semantic J-J and its section;
  6. import dimensions, tolerances, and native callouts;
  7. reconcile import coverage;
  8. create missing ordinates and side locations;
  9. complete section dimensions;
  10. add only missing source-backed family definitions;
  11. create the clean isometric;
  12. populate title, notes, and part identification;
  13. arrange eligible annotations;
  14. rebuild and perform final content-envelope layout;
  15. rebuild, clean selection state, and run read-only QA;
  16. emit complete evidence.
- [x] **R23-1102:** Static contract proves no final content is added after the
  final layout pass. r26 also proved the shared wiring compiles and its
  read-only P-0251 evidence path runs; production creation remains open.
- [x] **R23-1103:** Keep UI forms unchanged; no independent
  event/interface defect is demonstrated.

### Phase 12 — Static verification, deployment, and runtime acceptance

- [x] **R23-1200:** Update deployment manifest and import documentation for all
  added classes/modules.
- [x] **R23-1201:** The 2026-08-04 source inventory is clean for ASCII bytes,
  CRLF, no BOM, `Option Explicit`, trailing whitespace, and the 79-character
  limit across all 38 manifest-managed `.bas` and `.cls` files. r37 passed the
  full 435-test suite and programmatic VBE compilation.
- [x] **R23-1202:** Static contracts check procedure structure, public surface,
  and production call ordering. r37 programmatic VBE compilation is clean.
  public declarations, unresolved references, and call-site arity.
- [x] **R23-1203:** Static assertions reject:
  - unresolved `ICE` classification;
  - `FaceInSurfaceSense` as the hole decision;
  - old J-J extension math;
  - section import through ordinate eligibility;
  - hardcoded P-0251 manufacturing text;
  - feature-name/radius scoring;
  - final layout before definitions;
  - count-only QA.
- [x] **R23-1204:** Update:
  - `docs/SOLIDWORKS_API_VALIDATION.md`;
  - `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`;
  - `docs/CHANGELOG.md`;
  - `docs/CURRENT_STATUS.md`;
  - `src/target-spec-hybrid-v2/README_IMPORT.md`.
- [x] **R23-1205:** r37 synchronized exported source and embedded SWP through
  guarded deployment (`38/38`, candidate and target `VERIFY: PASS`).
- [x] **R23-1206:** Exact r37 managed-source readback is retained with the
  deployment evidence.
- [x] **R23-1207:** r37 full-project programmatic VBE compilation returned
  `verdict=Clean` with `firstFailedModule=None` in the focused scratch runner.
- [~] **R23-1208:** Fresh focused P-0251 r37 read-only runner evidence is
  retained at `probe_runs/20260804_164014/`,
  with `verdict=Clean`, `firstFailedModule=None`, and all nine probes
  completed. Production QA output and uncropped full-sheet screenshot remain
  open:
  - settings;
  - complete Immediate Window output;
  - QA report;
  - uncropped full-sheet screenshot;
  - relevant feature/type/definition probe evidence.
- [~] **R23-1209:** The user accepted the isolated P-0251 scratch layout after
  the r28 measured-rescale run. Semantic comparison against the manual
  reference and production acceptance remain open.
- [~] **R23-1210:** Corrected the live VBE dictionary-scope failure and the
  semantic coverage ordering defect. The r28 Phase 9 loop completed on the
  scratch, but its five automatic layout failures remain separately recorded
  from the user's visual acceptance.
- [ ] **R23-1211:** After P-0251 passes, compile and run the complete
  three-fixture regression. As of 2026-08-04, only the P-0251 disposable
  scratch drawing exists under `test_assets/`; P-0252 disposable drawings
  must be supplied before their authorized fixtures can be run.
- [x] **R23-1212:** Failed and successful runner logs are retained under
  `test_assets/iteration_evidence/`; manual references remain untouched.

## 7. P-0251 acceptance checklist

R23 is not accepted until all items are proved:

- [ ] Raw and effective type evidence exists for all three `ICE` features.
- [ ] Every accepted cut resolves through a supported typed definition.
- [ ] Six counterbore and four tapped physical locations are canonicalized.
- [ ] The stepped bore is a separate coaxial stack with source-backed intent.
- [ ] Primary X and Y ordinate schemes use proved centre/bottom datums.
- [ ] Imported linear locations have not suppressed required ordinates.
- [ ] One complete six-hole counterbore definition is visible and attached.
- [ ] One complete four-hole tapped definition is visible and attached.
- [ ] Section J-J contains all seven required semantic dimensions.
- [ ] Ø47 reads back as H7 with `+0.025/0.000`.
- [ ] No free-text note substitutes for a required section dimension.
- [ ] Side view contains complete tapped-hole location and definition coverage.
- [ ] No dimension appears on the isometric.
- [ ] No ordinate appears on the section or isometric.
- [ ] J-J has exactly three intended semantic segments.
- [ ] Both arrows and both J labels clear the border, zones, part-ID, notes,
  title block, and other view envelopes.
- [ ] No duplicate source dimensions, ordinate coordinates, callouts, section
  requirement keys, or physical instances remain.
- [ ] Full VBA compilation succeeds.
- [ ] Runtime QA passes truthfully.
- [ ] The uncropped drawing is manufacturing-useful and visually acceptable
  against the reference.

## 8. Stop conditions

R23 must fail closed instead of guessing when:

- an `ICE` feature has no supported underlying `GetTypeName`;
- `GetDefinition` does not yield the expected typed feature data;
- configuration or suppression state cannot be proved;
- owned feature geometry and drawing projection do not agree;
- a required drawing-context entity is not selectable;
- thread, fit, tolerance, depth, or machining-side intent is missing;
- a required section dimension cannot be created and read back;
- a native or fallback callout cannot be associated with one physical family;
- complete content cannot fit the approved sheet/scale;
- annotation, section-line, leader, or extent readback is unavailable.

## 9. Primary API references

- SOLIDWORKS 2025 `IFeature.GetTypeName2`:
  <https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IFeature~GetTypeName2.html>
- SOLIDWORKS 2025 `IFeature.GetTypeName`:
  <https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IFeature~GetTypeName.html>
- SOLIDWORKS 2025 `IFeature.GetDefinition`:
  <https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IFeature~GetDefinition.html>
- SOLIDWORKS 2025 `IExtrudeFeatureData2`:
  <https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2.html>
- Project CodeStack coverage:
  `docs/CODESTACK_DRAWING_API_COVERAGE.md`
- Project field guide:
  `docs/3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md`
- Drawing target:
  `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`
- API evidence ledger:
  `docs/SOLIDWORKS_API_VALIDATION.md`

## 10. Definition of done

R23 is done only when:

1. the embedded project matches the reviewed exported source;
2. the whole VBA project compiles in SOLIDWORKS 2025;
3. all three authorized fixtures complete the required regression;
4. P-0251 satisfies every semantic acceptance item above;
5. all required API operations have checked return/readback evidence;
6. no model, protected baseline, or manual reference was modified;
7. complete evidence is retained under `test_assets/`;
8. final drawings pass visual and manufacturing comparison, not merely a macro
   completion or QA-count check.
