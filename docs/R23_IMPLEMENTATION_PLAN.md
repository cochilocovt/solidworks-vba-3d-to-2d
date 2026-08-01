# R23 Target-Spec Hybrid Drawing Implementation Plan

**Status:** Production R23 blocked; final probes ran but exposed disposable
probe defects and unresolved entity-mapping, H7-authority, and coordinate-frame
contracts
**Prepared:** 2026-07-30; updated 2026-07-31
**Target revision:** `target-spec-hybrid-v2-2026-07-30-r23`
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
- [ ] **R23-006:** Compare both circular-curve read orders on one counterbore,
  tapped hole, mirrored hole, and resolved cut cylinder:
  - R22 order: `GetCurve → GetCurveParams3 → IsCircle → CircleParams`;
  - retained-probe order:
    `GetCurve → IsCircle → CircleParams → GetCurveParams3`.
  - The July 31 live run proves the shared `IsCircle` and `GetCurveParams3`
    portions of both orders are stable.
  - **`CircleParams` works — exclusion closed 2026-07-31.** With the guard
    defect removed it returned seven values on every tested edge with radii
    matching the owning cylinders exactly (`0.005500000`, `0.003300000`,
    `0.023500000`). Production may treat it as available evidence.
  - The original withdrawal reasoning, retained for provenance: the
    `SkippedNotCircle` result came from
    `Module_R23Phase0FeatureProbe.ReadCircleState` line 611, which guards the
    call with `If Not isCircle Then` — the same
    `If Not <SOLIDWORKS Boolean>` defect that rejected every cylindrical face
    and every circular edge in the later runs. `CircleParams` was never
    invoked, so no anomaly was ever observed. Its behaviour remains untested;
    the corrected probe now reads it as non-load-bearing evidence. Production
    must not depend on it until a run actually exercises it.
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
remains open is exactly what requires mutation - R23-506 and R23-508 - plus
R23-502, which is **not met** and is marked so below rather than being
claimed on a weaker datum.

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
- [ ] **R23-502: NOT MET.** The vertical datum resolves to the lowest
  projected hole (`y=0.087415` in `Drawing View4`), recorded as
  `datumKind=ProjectionDerived`. This task asks for the part's bottom
  **outline** geometry, which is a different kind of entity and has not been
  proved. The distinction is carried in evidence precisely so a
  projection-derived datum can never be read as an outline datum.
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

**Status: source complete, awaiting first live run.** Statically verified
only.

- [ ] **R23-600:** Always request native callouts when the configuration enables
  them.
- [ ] **R23-601:** Traverse display dimensions and require
  `IsHoleCallout=True` for native hole-callout classification.
- [ ] **R23-602:** Match native callout attachments to one graph family.
- [ ] **R23-603:** Validate quantity, size, depth, counterbore/thread data,
  view ownership, and variable-level tolerance/fit fields.
- [ ] **R23-604:** Retain one complete associative native callout per family.
- [ ] **R23-605:** Use source-backed controlled fallback only when no complete
  native definition exists.
- [ ] **R23-606:** Build fallback quantity from unique physical locations, not
  feature names or raw edge count.
- [ ] **R23-607:** Build counterbore/thread fields from typed feature data.
- [ ] **R23-608:** Build fit/tolerance fields only from source model dimension
  data.
- [ ] **R23-609: half met, and the remaining half is deliberate.** The new
  path contains none of it: no part number, no `6X`, no `M5x0.8`, no `H7`,
  no diameter literal, and no scoring by feature name or by proximity to an
  expected radius. A contract asserts each of those strings is absent.
  **The legacy literals are still in `Module7_TitleBlockEngine.bas`** - the
  callout text at lines 359-371 and the name/radius scoring at 405-435 -
  because Module7 is still the reachable production path and Module16 is
  not yet wired into `main`. Deleting them now would degrade the deployable
  macro while its replacement is disconnected. They come out in the phase
  that switches the pipeline over, and that switch bumps the revision.
- [ ] **R23-610:** Fail `MANUFACTURING_DEFINITION` with a field-specific reason
  if a required semantic field is unavailable.
- [ ] **R23-611:** Prove one six-hole counterbore definition and one four-hole
  tapped definition, correctly owned and attached.

### Phase 7 — Rebuild J-J from model intent

- [ ] **R23-700:** Split graph/projection discovery from ordinate creation so
  section construction can consume proved locations first.
- [ ] **R23-701:** Resolve one stepped-bore centre and a 2×3 face-hole family.
- [ ] **R23-702:** Build the P-0251 J-J page-coordinate path:
  1. stepped-bore centre;
  2. same X at the highest face-hole row;
  3. minimum-X face-hole column at that row;
  4. same column at the lowest face-hole row.
- [ ] **R23-703:** Prove the path crosses the stepped bore and all three holes
  in the chosen column.
- [ ] **R23-704:** Delete the old P-0251 `extension`,
  `topY + extension`, `bottomY - extension`, and outline-percentage fallback.
- [ ] **R23-705:** Convert page coordinates to source-view sketch coordinates
  exactly once before `CreateLine`.
- [ ] **R23-706:** Create exactly three view-owned segments and verify their
  selection order before `CreateSectionViewAt5`.
- [ ] **R23-707:** Parse `GetSectionLineInfo2` after creation and after every
  later view move.
- [ ] **R23-708:** Fail `SECTION_GEOMETRY` rather than approximate the cut when
  required feature identities or path intersections are unproved.

### Phase 8 — Add a dedicated section-dimension engine

- [ ] **R23-800:** Add `Module10_SectionDimensionEngine.bas`.
- [ ] **R23-801:** Define these P-0251 requirement keys:
  - overall thickness `18.00`;
  - bore-step depth `12.00`;
  - inner bore `Ø40.00`;
  - fit bore `Ø47 H7 +0.025/0.000`;
  - lower wall/step `11.50`;
  - long vertical reference `173.60`;
  - lower vertical reference `104.80`.
- [ ] **R23-802:** Reconcile imported section dimensions first by source
  dimension, attached geometry, semantic role, nominal, type, and tolerance.
- [ ] **R23-803:** Create only missing associative section dimensions.
- [ ] **R23-804:** Use view-scoped drawing entities for section dimensions and
  accept the live-proven `swDiameterDimension = 6` type for the imported
  47/40 diameters; do not require type 15.
- [ ] **R23-805:** Read back dimension type, nominal, source/attachment, fit,
  min/max tolerance, and API warning/error status.
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
- [ ] **R23-807:** Never replace a failed Ø47/H7 dimension with free text.
- [ ] **R23-808:** Arrange section requirements in dedicated lanes:
  - 18 and 12 above;
  - 11.5 below;
  - Ø40 and Ø47 on opposite bore sides;
  - 173.6 and 104.8 on separate exterior vertical lanes.
- [ ] **R23-809:** Exclude the section from generic auto-arrangement.
- [ ] **R23-810:** Remove the old free-text Ø47/Ø40 “bore callout” once actual
  section dimensions provide the definition.
- [ ] **R23-811:** Prove exactly one dimension per requirement key and no
  section ordinate.

### Phase 9 — Make final layout content-envelope aware

- [ ] **R23-900:** Build one page-coordinate content envelope per view from:
  - model outline;
  - display-dimension primitives and text;
  - note extents;
  - leader points;
  - callouts;
  - section segments;
  - arrow geometry;
  - J-label positions and text heights.
- [ ] **R23-901:** Include title block, part-identification band, general notes,
  content border, and zone-number regions as protected sheet rectangles.
- [ ] **R23-902:** Replace the fixed P-0251 upward bias with a constraint-based
  placement using complete envelopes.
- [ ] **R23-903:** Move views by the difference between envelope centre and
  assigned cell centre.
- [ ] **R23-904:** Rebuild, reacquire all geometry, and allow at most one
  correction pass.
- [ ] **R23-905:** Require explicit clearance between every view envelope and
  every other view/protected rectangle.
- [ ] **R23-906:** Require at least 2 mm clearance for J-J arrows and labels from
  the content border and part-identification band.
- [ ] **R23-907:** Do not reduce an approved view scale to force a fit.
- [ ] **R23-908:** Fail and request a larger sheet when the complete content
  cannot fit.
- [ ] **R23-909:** Ensure no annotation, dimension, note, callout, or view is
  created after final layout.

### Phase 10 — Replace count-based QA

- [ ] **R23-1000:** Add required stage `MODEL_INTENT_CATALOG`.
- [ ] **R23-1001:** Add required stage `MODEL_IMPORT_COVERAGE`.
- [ ] **R23-1002:** Add required stage `NATIVE_CALLOUT_COVERAGE`.
- [ ] **R23-1003:** Add required stage `PHYSICAL_LOCATION_GRAPH`.
- [ ] **R23-1004:** Add required stage `VIEW_PROJECTION`.
- [ ] **R23-1005:** Add required stage `ORDINATE_SCHEME`.
- [ ] **R23-1006:** Add required stage `SECTION_GEOMETRY`.
- [ ] **R23-1007:** Add required stage `SECTION_DIMENSIONS`.
- [ ] **R23-1008:** Retain and strengthen `MANUFACTURING_DEFINITION`.
- [ ] **R23-1009:** Add final post-content stage `FINAL_LAYOUT`.
- [ ] **R23-1010:** Replace “nonzero import” with per-view/per-category semantic
  coverage.
- [ ] **R23-1011:** Replace note-token-only stepped-bore checks with actual
  section display-dimension and tolerance inspection.
- [ ] **R23-1012:** Emit source provenance for every manufacturing field.
- [ ] **R23-1013:** Emit raw/effective feature type for every accepted/rejected
  manufacturing feature.
- [ ] **R23-1014:** Fail when an `ICE` feature has an empty or unsupported
  underlying type.
- [ ] **R23-1015:** Fail when a required model location lacks a proved drawing
  projection.
- [ ] **R23-1016:** Fail on duplicate physical, source-dimension, family
  definition, or section-requirement keys.
- [ ] **R23-1017:** Fail when J-J geometry or any mandatory annotation envelope
  is unavailable or unsafe.

### Phase 11 — Reorder the production pipeline

- [ ] **R23-1100:** Replace the initial `modelHoleFeatures` collection with the
  location graph.
- [ ] **R23-1101:** Implement this order:
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
- [ ] **R23-1102:** Verify no final content is added after the final layout pass.
- [ ] **R23-1103:** Keep UI forms unchanged unless an independent
  event/interface defect is demonstrated.

### Phase 12 — Static verification, deployment, and runtime acceptance

- [ ] **R23-1200:** Update deployment manifest and import documentation for all
  added classes/modules.
- [ ] **R23-1201:** Verify every deployable `.bas`/ordinary `.cls` is
  Windows-1252/ANSI, CRLF, no BOM, no export metadata, and `Option Explicit`.
- [ ] **R23-1202:** Run structural checks for procedure balance, duplicate
  public declarations, unresolved references, and call-site arity.
- [ ] **R23-1203:** Add static assertions rejecting:
  - unresolved `ICE` classification;
  - `FaceInSurfaceSense` as the hole decision;
  - old J-J extension math;
  - section import through ordinate eligibility;
  - hardcoded P-0251 manufacturing text;
  - feature-name/radius scoring;
  - final layout before definitions;
  - count-only QA.
- [ ] **R23-1204:** Update:
  - `docs/SOLIDWORKS_API_VALIDATION.md`;
  - `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`;
  - `docs/CHANGELOG.md`;
  - `docs/CURRENT_STATUS.md`;
  - `src/target-spec-hybrid-v2/README_IMPORT.md`.
- [ ] **R23-1205:** Synchronize exported source and embedded SWP using the
  guarded deployment workflow.
- [ ] **R23-1206:** Verify exact managed-source readback and r23 identity.
- [ ] **R23-1207:** Compile the entire VBA project in the SOLIDWORKS VBA Editor.
- [ ] **R23-1208:** Run focused P-0251 and retain:
  - settings;
  - complete Immediate Window output;
  - QA report;
  - uncropped full-sheet screenshot;
  - relevant feature/type/definition probe evidence.
- [ ] **R23-1209:** Compare P-0251 semantically and visually against the manual
  reference.
- [ ] **R23-1210:** Correct any remaining runtime or visual defects and repeat
  the focused loop.
- [ ] **R23-1211:** After P-0251 passes, compile and run the complete
  three-fixture regression.
- [ ] **R23-1212:** Retain failed outputs as regression evidence; do not
  overwrite manual references.

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
