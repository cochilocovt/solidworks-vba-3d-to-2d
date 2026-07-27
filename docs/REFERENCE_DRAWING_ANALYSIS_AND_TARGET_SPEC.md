# Reference Drawing Analysis and Target Macro Specification

> **Status:** Living production specification  
> **Updated:** 2026-07-26  
> **Applies to:** clean replacement source `src/target-spec-hybrid-v2/`; the existing `src/active-ordinate/active_ordinate.swp` is retained as runtime evidence until the replacement is compiled and validated  
> **Protected source:** `src/baseline-model-dims/` must not be changed

> **2026-07-18 r4 implementation update:** the coherent source-completion phase
> is finished offline under identity `target-spec-hybrid-v2-2026-07-18-r4`.
> Static/API evidence is E2/E3 only; clean embedded compilation, authorized
> execution, and visual/manufacturing acceptance remain E5-E7 gates.

> **2026-07-26 r15 runtime update:** r14 was deployed, compiled, and executed on
> P-0251. The retained drawing proves view creation, 1:1 scale, model-item
> insertion, property writes, cleanup, and evidence output, but fails
> manufacturing acceptance. R15 corrects the proven primary-view orientation,
> restricts model-item insertion to explicitly selected eligible orthographic
> views, resolves drawing-edge ownership through the view's referenced document,
> and adds exact section-step diagnostics. R15 is E2/E3 only until its next
> embedded compile and authorized run.
>
> **2026-07-26 r16 runtime response:** r15 was guarded-deployed and executed on
> P-0251. It proved the corrected portrait primary, 1:1 scale, isometric
> dimension exclusion, cleanup, and evidence writing. R16 removes the exact
> failing section-view assignment, eliminates the redundant model-import
> selection stop, consumes visible component entities directly for semantic
> ownership, uses title-block-rectangle-aware layout, and adds `Part Name` as a
> description alias. The template still lacks a structural `ITitleBlock`;
> production title/link acceptance remains blocked until that controlled input
> is defined.
>
> **2026-07-26 r17 runtime response:** r16 was guarded-deployed, compiled, and
> executed on P-0251. It runtime-proved J-J creation, selected-view
> orthographic model import, the portrait primary, 1:1 scale, configuration-
> first description, cleanup, and evidence writing. It also disproved direct
> use of visible drawing edges as model edges, exposed isometric dimension
> inheritance when the view is created after import, showed that the generic
> upper band cannot fit the tall side/section views, and displayed an
> unverified-unit mass. R17 restores drawing-to-model correspondence before the
> unchanged semantic gates, creates the isometric before selected-view import,
> uses a P-0251 reference-led four-view layout, and calculates active-
> configuration mass in metric kilograms. The structural `ITitleBlock` gate
> remains unchanged.
>
> **2026-07-26 r18 runtime response:** r17 was guarded-deployed, compiled, and
> executed on P-0251. It preserved the J-J section, `9 + 1` orthographic model
> dimensions, isometric isolation, eight layout moves, and the correct metric
> mass property. It also disproved drawing-document reverse correspondence for
> the visible edge proxies, exposed sheet-wide final-QA dimension iteration, and
> proved that the visible mass note still points elsewhere. R18 starts from the
> audited model features, maps their owned model edges and datum vertices into
> each view through `IView.GetCorrespondingEntity`, counts exact per-view
> arrays, and performs a unique mass-note link/readback transaction. Structural
> title, grouped callout, and bore-definition gates remain open.

## 1. Purpose, authority, and product decision

This document defines the manufacturing-drawing target for the SOLIDWORKS 2025 VBA macro. It replaces the earlier image-led implementation proposal with a current, evidence-ranked specification.

The production goal is not merely a macro that finishes or creates a nonzero number of dimensions. The result must approach the supplied manual references in manufacturing usefulness, semantic coverage, correctness, readability, and layout, while reporting unresolved intent honestly.

Only these three fixtures are authorized for live macro execution:

- `test_assets/models/P-0251-14A-001.SLDPRT`
- `test_assets/models/P-0252-01-001.SLDPRT`
- `test_assets/models/P-0252-01-013.SLDPRT`

The parts may be opened, measured, and queried, but their features, configurations, properties, and saved model state must not be changed.

### 1.1 Fixed hybrid workflow

The production macro must use one fixed workflow:

1. Import applicable model-marked dimensions and annotations.
2. Add qualified feature-location dimensions only where required location coverage is missing.
3. Use ordinate dimensions only for proven hole centres or other explicitly supported feature locations.
4. Avoid duplicate size, callout, and location information.
5. Arrange views and annotations within measured sheet and title-block limits.
6. Populate title-block data, notes, and the approved part-identification/barcode treatment.
7. Produce fail-closed QA and retained run evidence.

“Hybrid” has two distinct meanings in this project:

- **Hybrid dimensioning workflow:** model annotation import followed by evidence-backed ordinate fallback inside VBA.
- **Hybrid companion architecture:** VBA remains the only production drawing generator. The project-local Python companion performs probes, regression checks, diagnostics, exports, and review. It is not a second production drawing engine and shares no mutable runtime state with the VBA macro.

Separate user-selectable “model dimension” and “ordinate dimension” modes must not be reintroduced unless the user changes this requirement.

### 1.2 Release claim

The target is a shop-usable drawing requiring minimal designer cleanup. A drawing must not be described as release-ready when model metadata, feature ownership, datum intent, tolerance intent, or title-block data is unresolved.

## 2. Evidence hierarchy and status vocabulary

The following evidence levels must not be conflated:

| Level | Evidence | What it proves |
|---|---|---|
| E1 | Manual reference raster | Visible content, layout, and apparent manufacturing intent only |
| E2 | Exported VBA/static tests | Source text and static invariants only |
| E3 | SOLIDWORKS 2025 MCP, installed type library, and official API contract | Method signatures, enum values, documented semantics |
| E4 | Narrow disposable live probe | The probed API behavior on the named build and fixture only |
| E5 | Full embedded VBA compile | The current `.swp` project compiles |
| E6 | Embedded macro execution on an authorized fixture | Runtime behavior and captured diagnostics for that configuration |
| E7 | Visual and semantic comparison with the reference and this specification | Manufacturing-drawing acceptance |

For implementation status:

- **Implemented in export** means code exists in the exported module.
- **Embedded** means the exact frozen r4 component set is present in the
  designated recoverable replacement `.swp` and a retained post-import/re-export
  identity check matches its manifest.
- **Compiled** means the whole embedded VBA project compiled in SOLIDWORKS.
- **Runtime-proven** means the embedded macro executed successfully on an authorized fixture with retained evidence.
- **Accepted** means the output passed visual and semantic review, not merely API or count checks.

### 2.1 Historical source-completion state on 2026-07-18

| Area | Current evidence | Status |
|---|---|---|
| Reference drawings | All three 4960 × 3507 single-sheet rasters were inspected | E1 complete |
| Historical generated drawings | All six baseline/active rasters were inspected | Historical E1 complete |
| Clean replacement source | Nine standard modules, three data/evidence classes, forms, handlers, import guide, and source identity `target-spec-hybrid-v2-2026-07-18-r4` | Source-complete offline export |
| Exported model importer | Verified mask, conditional hole callout bit, duplicate elimination, activation context, zero-result retry, type/view counts, and cleanup | Implemented in export |
| Exported ownership path | Component-qualified visible edges, corresponding model edge, full-circle/internal-cylinder/matched-face/feature/seed/configuration/axis proof, stable physical/family identity, and fail-closed rejection | Implemented in export; fixture behavior needs E4/E6 |
| Exported datum/ordinate path | Typed selectable X/Y datum proofs, attachment-backed linear/ordinate coverage, datum-first selection, selected-count checks, result decoding, scoped suppression, and cleanup | Implemented in export; fixture behavior needs E4/E6 |
| Sections and details | Reference-led J-J/B-B section paths plus mandatory Pump Holder C/D detail transactions from exact `*Bottom` at 3:1 | Structural E2/E3 implementation; E4/E6 must prove coordinate, selection, creation, and readback behavior; E7 must prove feature content, hatch/arrows, and legibility |
| Controlled sheet/title path | Fail-closed template/format identity, actual scale, zone/title extents, configuration-first properties, exact link/rendered-value/extent evidence | Implemented in export; D-04/template runtime not proved |
| Exported QA | Sticky requirement stages, physical/projection ledgers, view/ordinate policy, annotation/leader/note/title regions, exact failure result, final cleanup, and atomic evidence write | Implemented in export; E6/E7 remains |
| Offline verification | Source-contract, fake-COM, structure, and regression suite | 49 tests passed; E2 only |
| R4 checkpoint | Source archive plus 35-entry SHA-256 manifest under `test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/` | Frozen E2/E3 compile-handoff evidence; historical r3 manifest superseded for r4 identity |
| Installed API verification | Interop `33.1.2.4` reflection plus local SOLIDWORKS API MCP corpus | E3 complete for currently used members; binding behavior still needs E4/E6 where identified |
| Embedded replacement macro | The clean replacement has not yet been imported into a `.swp` project | Not proved |
| Full VBA compile | No retained compile evidence for the clean replacement | Not proved |
| Controlled template availability | The candidate `VEEMAP DRAWING.DRWDOT` is confirmed at `V:\VEEMAP\SW_data\Custom Templates\`; the r4 constant omitted the `VEEMAP` directory level | R5 corrects the path; linked sheet-format/title-contract proof remains a live gate |
| Live API probe | One disposable P-0251-14A-001 front-view probe inserted 11 annotations and obtained 39 edges from a visible component | Narrow E4 only |
| Embedded macro run | No designated r4 `.swp` has been compiled or executed; `active-ordinate` history is not r4 runtime evidence | Not proved |
| Three-part visual acceptance | No current output has passed the full matrix | Not proved |

The companion’s passing tests and a probe JSON marked `PASS` do not satisfy E5, E6, or E7.

### 2.2 Current live checkpoint on 2026-07-26

The retained r14 evidence is:

- `test_assets/iteration_evidence/macro_qa/20260726_102348_P-0251-14A-001/QA_REPORT.txt`
- `test_assets/iteration_evidence/macro_qa/20260726_102348_P-0251-14A-001/diagnostic-drawing.png`

| Requirement | r14 evidence | r15 response | Acceptance state |
|---|---|---|---|
| Reference-led primary orientation | P-0251 primary was landscape, opposite the portrait manual view | P-0251 primary `IView.Angle` set to clockwise 90 degrees with immediate readback | Needs E5/E6/E7 |
| J-J section | No section; VBA error 91 after the section plan was recorded | Step-level diagnostics bracket every selection and `CreateSectionViewAt5` operation | Root step needs next E6 result |
| Orthographic-only dimensions | Isometric view contained 10 display dimensions | `AllViews=False`; each eligible orthographic view is explicitly selected and verified | Needs E6/E7 |
| Ownership-proven holes | 111 `NoModelCorrespondence` rejections; zero candidates | `ReferencedDocument.Extension.GetCorrespondingEntity2(drawingEdge)` | Needs E6 on ten canonical locations |
| Ordinate coverage | Zero horizontal/vertical groups | Downstream of the ownership correction; existing datum-first transactions remain fail-closed | Needs E6/E7 |
| Controlled title block | Visible title graphics, but `ISheet.TitleBlock Is Nothing` | No QA relaxation; a genuine structural title block and approved links remain D-04 input | Blocked by template/input |
| Description | Required property resolved empty | No value is invented | Blocked pending exact approved source/property |
| Layout | Diagnostic packed rows did not fit and the isometric overlapped the working area | Orientation and dimension-policy corrections reduce false extents; production layout still requires structural title bounds | Needs D-04 and E6/E7 |
| Scale and cleanup | 1:1 readback, sheet restoration, zero selection, and final cleanup proved | Preserved | Runtime-proven for r14 only |

This checkpoint deliberately distinguishes code defects from controlled-input
gates. A visually drawn title table is not equivalent to `ITitleBlock`, and the
macro must not invent the missing description or silently pass title/link QA.

### 2.3 R15 live checkpoint and r16 response

The retained r15 evidence is:

- `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/QA_REPORT.txt`
- `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/diagnostic-drawing-with-dialog.png`
- `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/macro-settings.png`
- `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/deployment-r15.png`

| Requirement | r15 evidence | r16 response | Acceptance state |
|---|---|---|---|
| Primary orientation | Requested and actual angle both `-1.570796327`; portrait main view visibly matches the reference direction | Preserve the proved P-0251-only rotation | Runtime-proven for r15; preserve in r16 |
| J-J section | Error 91 at `SelectData.ViewAssignment.Before`; three section segments were already created | Keep active source view and verified segment ownership/selection; skip the failing assignment before `CreateSectionViewAt5` | Needs E5/E6/E7 |
| Orthographic model annotations | Matching active view and `selectedByID=True`, followed by a contradictory legacy selection failure; import count zero | Remove the redundant fatal selection gate and record execution immediately before `InsertModelAnnotations4(False, True)` | Needs E6/E7 |
| Ownership-proven locations | 111 visible edges rejected as `NotCircular` after double conversion | Use `GetVisibleEntities2(component, Edge)` entities directly, retaining every semantic ownership gate | Needs E6 on ten canonical locations |
| Layout | Portrait main plus side/isometric could not fit above a full-width 100.98 mm lower reserve; zero layout moves | Reserve only the title-block rectangle, use safe lower-left space, validate notes/border/title/view collisions, never rescale | Needs E6/E7 |
| Description | `Description` and `PartName` were empty while the controlled title table displays the part name | Add exact alias `Part Name`; still no invented fallback | Needs E6 property evidence |
| Controlled title block | `ISheet.TitleBlock Is Nothing`; visible title graphics only | Remains fail-closed; define the title block and approved note links in the controlled template | Blocked by controlled input |
| Scale, iso policy, cleanup | 1:1 readback, zero iso dimensions, sheet restored, selection count zero | Preserved | Runtime-proven for r15 |

At the prior checkpoint, R16 was source/static evidence only. The live
checkpoint below now supersedes that state; no r17 runtime or manufacturing
result is claimed from offline tests alone.

### 2.4 R16 live checkpoint and r17 response

The retained r16 evidence is:

- `test_assets/iteration_evidence/macro_qa/20260726_145625_P-0251-14A-001/QA_REPORT.txt`
- `test_assets/iteration_evidence/macro_qa/20260726_145625_P-0251-14A-001/diagnostic-drawing.png`
- `test_assets/iteration_evidence/swp_deployment/20260726_145546/`

| Requirement | r16 evidence | r17 response | Acceptance state |
|---|---|---|---|
| Embedded identity | Candidate compile succeeded; 15/15 managed components matched; promoted revision read back as r16 | Preserve guarded candidate/compile/readback/promotion workflow | Runtime-proven for r16 |
| J-J section | Three ordered source-owned segments selected; `CreateSectionViewAt5` returned non-Nothing; label and section-line info read back | Preserve the proved transaction; improve placement only | Structurally runtime-proven; visual intersection/content still needs E7 |
| Orthographic model annotations | 9 display dimensions on Drawing View1 and 1 on Drawing View2; section skipped | Preserve selected eligible-view import | Runtime-proven; coverage/readability still needs E7 |
| Isometric dimension policy | Drawing View4 contained 10 display dimensions after the isometric was created following import | Create the isometric before import, then select only registered orthographic views with `AllViews=False` | Needs E5/E6/E7 |
| Ownership-proven locations | 64 + 47 visible entities all rejected `NotCircular` under `mapping=DirectComponentEntity` | Use drawing-document `GetCorrespondingEntity2(drawingEdge)` before unchanged semantic gates | Needs E6 on ten canonical locations |
| Ordinate coverage | Zero accepted locations and zero horizontal/vertical groups | Downstream of corrected model-edge mapping; datum/coverage/selection gates remain fail-closed | Needs E6/E7 |
| Layout | Generic remaining-view band did not fit; zero layout moves; section/title region visibly crowded | Pack primary, side, and J-J across the title-clear left zone; place 1:2 iso upper-right; no scale changes | Needs E6/E7 |
| Mass | Visible `1296.82` under `MASS(KG)` conflicts with reference `1.30` | Calculate active-configuration mass through `GetMassProperties2`, require status OK, use metric kg index 5, display two decimals | Needs E5/E6/title-link evidence |
| Description | Resolved from `model-configuration:Part Name` | Preserve | Runtime-proven property source; exact title link remains blocked |
| Controlled title block | `ISheet.TitleBlock Is Nothing`; exact links/extents unproved | No QA relaxation | Blocked by controlled input |
| Scale and cleanup | 1:1 readback, sheet restored, selection count zero, evidence write proved | Preserve | Runtime-proven for r16 |

R17 passes 62 offline source/structure/regression checks. Those checks are E2;
the drawing-edge correspondence, isometric isolation, placement, mass readback,
and final drawing still require the next embedded compile and authorized run.

### 2.5 R17 live checkpoint and r18 response

The retained r17 evidence is:

- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/QA_REPORT.txt`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/diagnostic-drawing.png`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/R17_VISUAL_AND_QA_DIAGNOSIS.md`

| Requirement | r17 evidence | r18 response | Acceptance state |
|---|---|---|---|
| J-J and four-view layout | J-J retained; eight moves; portrait primary, narrow side, section, and upper-right isometric visibly separated | Preserve the proved structure and reference-led layout | Runtime-proven for r17; preserve and recheck after ordinates |
| Orthographic import / iso policy | Immediate arrays `9, 1, 0, 0`; final QA later misreported 10 on iso | Count each exact `IView.GetDisplayDimensions` array; never traverse `GetNext5` across the sheet | Import/iso isolation runtime-proven; final QA repair needs E5/E6 |
| Ownership-proven locations | 64 + 47 `GetCorrespondingEntity2` calls returned `Nothing`; zero candidates | Traverse audited model features, internal cylindrical faces, and circular model edges; map model edge into view with `IView.GetCorrespondingEntity` | Needs E5/E6 on ten canonical locations |
| Datum proof | Never reached because candidate collection was zero; drawing-to-model vertex path shares the failed direction | Traverse visible solid-body vertices model-first and map each into the view | Needs E5/E6 selection proof |
| Ordinate coverage | Zero X/Y groups | Preserve datum-first transaction and coverage gates downstream of model-first mapping | Needs E6/E7 |
| Per-view QA | Final total 39 contradicted immediate total 10 | Exact array iteration per view | Needs E5/E6 |
| Mass | API value `1.296824 kg`; drawing property `1.30`; visible title `1296.82` | Require exactly one linked mass note, set `$PRP:"Mass"`, rebuild, verify unresolved link and rendered `1.30` | Needs E5/E6/title-note evidence |
| Layout note checks | Visually coherent; invalid note extent repeated for all views while boundaries already diagnostic | Diagnostic-only skip when boundaries are already unproved; record `acceptance=False`; production remains fail closed | E2 only; no acceptance relaxation |
| Hole callouts | No grouped callout equivalent to reference | Keep model-annotation path; do not use modal `AddHoleCallout2` in unattended automation | Open manufacturing gate |
| Controlled title | No structural `ITitleBlock` | No structural QA relaxation | Blocked by controlled input |

R18 is source/API evidence only until the managed source is embedded, compiled,
run on an authorized fixture, and visually compared. The complete project-local
suite passes 66 tests, all 15 managed components pass the required source-format
gate, and the guarded read-only deployment preflight resolves the r18 identity.

## 3. Reference-drawing evidence

The reference folder contains exactly:

| Part | Reference |
|---|---|
| P-0251-14A-001 | [P-0251-14A-001.PNG](../test_assets/reference_drawings/P-0251-14A-001.PNG) |
| P-0252-01-001 | [P-0252-01-001.JPG](../test_assets/reference_drawings/P-0252-01-001.JPG) |
| P-0252-01-013 | [P-0252-01-013.JPG](../test_assets/reference_drawings/P-0252-01-013.JPG) |

The rasters show substantially more complete, shop-oriented drawings than the generated snapshots. They do not prove that every visible dimension is semantically correct, that an annotation came from the model rather than the drawing, or that a circular edge belongs to a particular model feature.

### 3.1 Cross-part drawing language

The references establish these common expectations:

- A zoned border and controlled title block define the usable sheet.
- Manufacturing dimensions belong on suitable orthographic, section, or detail views.
- Isometric views are orientation aids and remain undimensioned.
- Feature size and manufacturing definition are separated from feature location.
- Equivalent holes use grouped quantity/size/thread/depth/counterbore callouts.
- Location schemes use deliberate part-specific origins, centre axes, or datums.
- Fits, tolerances, machining side, through/blind state, and depths are preserved.
- Centre marks, centre lines, hidden lines, hatching, cut arrows, and detail boundaries are used where they clarify intent.
- Title data and general notes are kept separate from working drawing geometry.
- Semantic coverage alone is insufficient: crowding, leader forests, ambiguity, and clipping are acceptance failures.

The references do **not** establish one universal datum strategy or one universal barcode location.

### 3.2 P-0251-14A-001 — Conveyor Bracket

#### Visible reference truth

- Main front view with the semicircular `R36` head, large bore, and 2 × 3 face-hole pattern.
- Narrow left side view with hidden side-hole geometry.
- `SECTION J-J` through the functional bore.
- One undimensioned isometric view.
- Large bore definition includes `Ø47 H7 +0.025/0.000` and `Ø40`.
- Six-hole family: `6X Ø6.6 THRU` with counterbore `Ø11` depth `6`.
- Side-hole family: `4X Ø4.2` depth `12.4`, tapped `M5x0.8 - 6H` depth `10`.
- Face-hole X locations are symmetric about a centre zero; Y locations use a bottom zero.
- No formal lettered GD&T datum is visible.
- Title data includes `SUS304`, mass `1.30`, quantity `1`, scale `1:1`, part name `Conveyor Bracket`, and surface treatment `Polishing`.

#### Target implication

The final drawing must preserve the symmetric centre/bottom location logic, grouped face-hole and side-hole definitions, stepped-bore section, fit, and readable main profile. A nearest-hole “datum” is prohibited.

### 3.3 P-0252-01-001 — Base Plate

#### Visible reference truth

- One large principal face view, one thin edge view, and two undimensioned isometric views exposing opposite faces.
- No section or detail view.
- Multiple coordinate bands and origins distinguish feature groups and front/back operations.
- Visible families include through openings `Ø40`, `Ø45`, `Ø50`, and `Ø18 THRU ALL`.
- Threaded, blind, through, precision-fit, and counterbored families are separately identified, including M4, M5, and M6 groups.
- Back-face intent is explicit in callouts such as `FROM BACK` and `190 FROM BACK`.
- Profile requirements include `4X R10` and `4X C10`.
- Title data includes `AL6082`, mass `11.13`, quantity `1`, scale `1:6`, part name `BASE PLATE`, and surface treatment `Anodizing`.

#### Layout warning

This is the most crowded reference. Several long leaders cross the part, coordinate clusters are congested, and the edge view is compressed. The target must reproduce the manufacturing meaning while improving association and spacing; the reference’s clutter is evidence of required semantic density, not a layout pattern to copy.

#### Target implication

The macro must distinguish feature families, machining face, through/blind state, fits, and depth. A single global zero may be insufficient to reproduce the reference’s several feature-group coordinate schemes; this is an explicit decision item in Section 11.

### 3.4 P-0252-01-013 — Pump Holder

#### Visible reference truth

- Dimensioned profile view, lower orthographic view, left and right end/side views, `SECTION B-B`, two enlarged details, and two undimensioned isometric views.
- `DETAIL C` and `DETAIL D` are each shown at 3:1 and define small `7 × 4` tab/end geometry with `C0.5` chamfers.
- Section B-B exposes the upper ear and two threaded holes: `2X Ø3.3` depth `10.1`, tapped `M4x0.7 - 6H` depth `8`.
- Profile evidence includes `R5`, `R2`, `R3.2`, `117.82°`, `30°`, and linear/ordinate values.
- Formal datum feature `A` is attached to a principal surface.
- Two perpendicularity frames show `⟂ 0.01 | A`.
- A `6.4 g6` fit is visible; the exact displayed deviations must be confirmed from native annotation data.
- Title data includes `AL6082`, mass `0.07`, quantity `1`, scale `1.5:1`, part name `PUMP HOLDER`, and surface treatment `Anodizing`.

#### Target implication

The final drawing must preserve datum A, applicable GD&T/fit data, the threaded-hole section, profile geometry, and readable small-feature definition. Details C and D are mandatory at 3:1 for first acceptance; the second isometric remains a separate information-value decision in Section 11.

### 3.5 General notes and part identification

All three references show these general notes:

- all dimensions are in millimetres;
- all corners are chamfered `0.5 × 45°`;
- remove all sharp edges and burrs.

The part-identification treatment is not uniform:

- P-0252-01-013 shows `BAR CODE:- *P-0252-01-013*` above the bottom-right title block.
- P-0252-01-001 and P-0251-14A-001 show large asterisk-wrapped part-number text at lower left.

The rasters do not prove machine-readable barcode bars or a validated symbology. The macro must not claim barcode validity until the required font/symbology and scan test are defined.

## 4. Historical generated-output comparison

The images under `test_assets/generated_output/` are dated regression evidence, not current embedded-macro acceptance.

| Capability | Manual references | `baseline-model-dims` snapshots | `active-ordinate` snapshots |
|---|---|---|---|
| Primary orientation | Manufacturing-led | All three primary views are rotated roughly 90° from the references | Same orientation mismatch |
| Model annotations | Curated and grouped | Many visible values/callouts, with severe overlap and weak hierarchy | No visible dimensions or callouts |
| Ordinate chains | Deliberate, part-specific | No coherent final scheme | No visible chains |
| Sections/details | Part-specific | Missing | Missing |
| Isometric views | One or two, clean, undimensioned | Present but poorly scaled/placed | P-0252-01-001 has no usable isometric; only fragments near/off the upper-right boundary |
| Layout | Uses the sheet and protects title data | Crowding and clipping coexist with blank space | Clean mainly because annotations are absent; sparse and sometimes off-frame |
| Border/title block | Visible and populated | Not visible | Not visible |
| Notes/part ID | Controlled but reference-specific | Free text, sometimes clipped | Free text, sometimes clipped |

Image evidence cannot prove why the title block is absent, which API calls failed, whether visible annotations are model-marked, or whether any circular candidate is a real hole.

## 5. Mandatory target behavior

### 5.1 Preflight and configuration

The macro must:

1. Require an active, saved `.SLDPRT`.
2. Refuse unapproved model paths.
3. Resolve and validate the drawing template and sheet format before drawing creation.
4. Record active configuration, SOLIDWORKS build, macro/export revision, template, and settings.
5. Reset global collections, selections, and pick mode deterministically.
6. Force the fixed hybrid workflow.
7. Fail an ordinate group when its configured datum cannot be represented by a proven selectable entity.

### 5.2 View policy and sheet plan

- Select a manufacturing-useful primary orientation; do not preserve the historical 90° orientation mismatch merely because geometry fits.
- Create only views that add required information.
- Keep all isometric views undimensioned.
- Exclude sheet, isometric, section, detail, auxiliary, alternate-position, detached, and unproved named/relative views from ordinate processing.
- Preserve one deterministic primary section with the current `SectionConfig` data model.
- Do not claim genuine multi-section support.
- Use optional details only where the small feature cannot be defined legibly at the main scale.
- Measure actual sheet size, border, title-block reserve, and each created `IView.GetOutline`; fixed coordinates alone are not acceptable.

### 5.3 Model annotation import

The import allowlist is:

- dimensions marked for drawing;
- geometric tolerances;
- notes intended for drawing;
- Hole Wizard profile dimensions;
- Hole Wizard location dimensions;
- hole callouts only when `ImportHoleCallouts=True`.

The importer must:

1. Establish a verified active/selected drawing-view context.
2. Use `InsertModelAnnotations4` with `DuplicateDims=True`.
3. Attempt whole-drawing import and retry selected real views after a zero result.
4. Count returned annotations and confirm displayed dimensions by view.
5. Clear selections and restore sheet/normal pick state on success and failure.
6. Preserve fits, tolerances, GTols, datum symbols, and semantic size information.
7. Never import not-marked dimensions merely to increase a completeness count.

### 5.4 Feature evidence and hole ownership

An ordinate candidate must pass all of these gates:

1. The source model contains a recognized hole-producing feature or other explicitly supported location feature.
2. Pattern/mirror instances are expanded or otherwise accounted for.
3. The feature’s centre, axis, size, depth, end condition, thread/counterbore/countersink data, configuration, and machining side are captured when available.
4. A visible drawing edge is matched to that feature instance by model ownership and projected geometry.
5. The edge is a complete qualifying circle in the current projection.
6. The feature axis is sufficiently normal to the view for a circular-centre location.
7. Concentric rings at one feature location collapse to one location candidate while retaining the semantic callout.

Bosses, fillet arcs, partial arcs, unrelated rounds, construction geometry, and unresolved generic cuts must be rejected and logged. A visible circle alone is not proof of a hole.

The required ownership bridge is drawing-edge led, not sketch-point led:

1. Retain the `Component2` that produced each visible drawing edge.
2. Map the drawing edge to the underlying model edge with `IModelDocExtension.GetCorrespondingEntity2`.
3. Require a complete circular body edge through `GetCurve` plus `GetCurveParams3`.
4. Inspect adjacent cylindrical faces and resolve owning or pattern-seed features.
5. Accept only active-configuration `HoleWzd`, `AdvHoleWzd`, `SketchHole`, or a generic cut with independently proved internal cylindrical-void semantics.
6. Transform both centre and axis into view space; reject axes not normal to the view.
7. Deduplicate by component, proven feature instance, projected centre, and axis while preserving one semantic counterbore/countersink/thread stack.

Raw `ISketch.GetSketchPoints2` results are not an admissible feature-instance count. The retained P-0251 probe reported 14 points for the reference's six-hole counterbore family and therefore overcounted that family. For Hole Wizard semantic cross-checking, use `IWizardHoleFeatureData2.GetSketchPoints`/`GetSketchPointCount`, which are documented to return centre-hole points; visible owned edges still enumerate the displayed instances.

### 5.5 Datum policy

- `Center` must resolve to a real origin, centreline intersection, datum point, or other proven selectable centre entity.
- `Bottom-Left` and `Top-Left` must resolve to actual corner/datum geometry in the drawing/model coordinate relationship.
- The selected zero must not be the hole nearest a target corner.
- The datum entity, model/view/sheet coordinates, selection result, and owning view must be recorded.
- If no valid selectable datum exists, report `DatumNotSelectable`, skip the group, and fail QA for required coverage.

Datum A and a coordinate zero are different concepts. The Pump Holder’s formal datum A must be preserved when present even if an ordinate group uses a separate validated coordinate origin.

### 5.6 Coverage reconciliation and ordinate creation

Before adding an ordinate, the macro must index existing imported location coverage by feature, view, direction, datum, and tolerance.

- Do not suppress size/thread/depth information because a location is covered.
- Add only missing X and/or Y location directions.
- Suppress repeated equal coordinates within a group using a documented model/view tolerance.
- Select the datum first with append disabled.
- Append remaining entities through a proven order-preserving path.
- Check every selection result, `MultiSelect2` result, and final selection count.
- Require the datum plus at least one unique non-datum location.
- Call `AddOrdinateDimension` with the verified horizontal or vertical enum.
- Decode the full `swCreateOrdDimError_e` result into QA.
- Call `SetPickMode` and clear selections on every success, failure, and early exit.
- Create no more than the deliberately planned groups for the feature family/view.

### 5.7 Sections and details

The current scope supports one deterministic primary section:

- use a supported orthographic source view;
- compute the cut in drawing-view/sheet coordinates;
- pass through the intended functional feature;
- check selection/sketch/view-creation return values;
- retain readable cut arrows, label, hatch, and section label;
- place the section inside a reserved layout cell.

The five UI entries contain only `Label` and `Vertical`; they do not constitute genuine five-section support. Do not change forms or handler classes without a demonstrated interface defect or explicit UI redesign approval.

Pump Holder Details C and D are mandatory. Their internal definition must name the exact `*Bottom` source view, reference-led centre and circular boundary, label, destination, independent 3:1 scale, source/detail ownership, and the small geometry each view must expose. Structural API readback is not visual proof that the `7 x 4` and `C0.5` features are correctly enclosed or legible.

### 5.8 Annotation layout

`AlignDimensions` is a checked first pass, not a complete layout engine.

Acceptance requires:

- no view or annotation intersection with the title-block reserve or outside the usable border;
- no text-to-text overlap;
- no text placed over unrelated geometry;
- unambiguous leader-to-feature association;
- no leader routed through an unrelated view;
- readable dimension lanes around the owning view;
- grouped repeated-hole callouts instead of repeated identical leaders;
- a standard scale chosen for geometry plus annotation load;
- remeasurement after every view/annotation move and rebuild.

If content cannot fit readably on the selected sheet, QA must fail and recommend a larger sheet or an explicitly approved view reduction.

### 5.9 Title block, notes, and part identification

The macro must use a controlled sheet format with a zoned border and known title-block cells.

Required fields, where authoritative source data exists:

- part number and part name/description;
- material, mass, quantity, and actual displayed scale;
- project, unit, and customer code;
- surface treatment and heat treatment;
- revision;
- drafter/designer, checker, approver, and dates.

Requirements:

1. Populate linked notes/custom properties rather than free notes at fixed unscaled coordinates.
2. Read configuration-specific properties before documented fallbacks.
3. Check custom-property API return codes.
4. Do not replace approved data with blanks or uncontrolled defaults.
5. Derive mass only from the active configuration with verified units.
6. Keep general notes and part identification inside their assigned cells/regions.
7. Treat missing mandatory fields as explicit QA failures or controlled placeholders.
8. Do not claim a barcode until symbology/font and scan validation are defined.

### 5.10 Truthful QA

Minimum retained QA fields:

- source path/configuration and SOLIDWORKS build;
- embedded macro and exported-source identity;
- template, sheet dimensions, border/title-block reserve, and scale;
- planned/created views with type, orientation, outline, and in-bounds state;
- imported annotation count by type and view;
- recognized feature instances and unresolved features;
- accepted/rejected projected candidates with reasons;
- callout and ordinate coverage by feature/direction;
- datum entity and coordinates for each group;
- `MultiSelect2`, selection-count, and `AddOrdinateDimension` results;
- duplicate suppression;
- section/detail results;
- collisions/out-of-bounds conditions;
- title property sources and missing fields;
- final selection/pick-mode state;
- `PASS`, `PASS WITH WARNINGS`, or `FAIL`.

Mandatory failures include:

- zero imported model items when required model annotations exist;
- zero total dimensions;
- an unproved required datum;
- a failed required ordinate/API operation;
- an unresolved required feature size or location;
- a duplicate required group or contradictory dimension;
- an unsupported view receiving ordinates;
- title-block/sheet intrusion;
- unreadable collision;
- missing mandatory title data.

Nonzero dimensions or a finished macro run are never sufficient for `PASS`.

## 6. Current implementation gap ledger

| Capability | Clean replacement export | Embedded/runtime state | Remaining requirement |
|---|---|---|---|
| Annotation mask and callout toggle | Implemented | Not compiled/runtime-proven | Prove off/on mask and visible result delta on all three fixtures |
| Whole-drawing import plus zero retry | Implemented | Not compiled/runtime-proven | Prove counts, view context, and cleanup on all three fixtures |
| Component-qualified visible edges | Implemented | Narrow earlier P-0251 API path only | Prove every intended replacement view |
| Datum-first selection and cleanup | Implemented with separate corner/origin proof | Not compiled/runtime-proven | Prove the selectable zero for every exposed datum/view |
| API-based view exclusion | Implemented through type/orientation plus explicit run registry | Not compiled/runtime-proven | Prove all intended inclusions and exclusions |
| Circle/semantic-stack de-duplication | Implemented by owned feature, projected centre, axis, and radii stack | Not compiled/runtime-proven | Prove counterbores, repeated coordinates, patterns, and mirrors |
| Feature ownership | Connected to the ordinate engine | API contracts verified; runtime not proved | Prove correspondence, owner/seed, configuration suppression, and tolerances on the fixtures |
| Coverage reconciliation | Implemented with attachment-backed linear/ordinate, family, view, datum, and direction proof | Not compiled/runtime-proven | Prove imported X-only, Y-only, X/Y, and uncovered cases |
| Grouped semantic callouts | Model import plus semantic-stack evidence | Not proved | Verify family/quantity/face/depth treatment; add only evidence-backed fallback behavior if import is insufficient |
| Primary section | Deterministic reference-led J-J/B-B paths with typed sketch segments, contextual selection, label and source-line readback | Not current-run proven | E4/E6-prove coordinate/selection/creation behavior and actual feature intersection; E7-prove arrows, hatch, placement, and legibility |
| Details | Mandatory C/D implementation from exact `*Bottom`; single complete circular profile, ownership, standard/circle style, independent 3:1 scale, and full smooth outline are read back | Not current-run proven | E4/E6-prove profile-coordinate and creation/readback behavior; E7-prove intended `7 x 4`/`C0.5` content, dimensions, and legibility |
| View/orientation planning | Fixture-led view plans without unconditional rotation plus measured grid | Not compiled/runtime-proven | Visually prove orientation and fixture-specific information value |
| Collision-aware layout | Standard-scale measured view grid plus view, note, annotation, leader, border, title, note-band, and part-ID checks | Not visually proved | E4/E6-prove installed-build extent/readback behavior; E7-prove glyph placement, collision resolution, and readability |
| Title block | Controlled template/format requirements; configuration-first sources; exact property link, rendered value, note extent, and actual scale evidence | Template/runtime not proved | Resolve D-04 and prove every required visible link/cell |
| QA | Sticky requirement stages, semantic ledgers, API/layout/title/cleanup proof, read-only final pass, and atomic evidence output | Static only | Prove evidence completeness and visual acceptance on retained runs |

## 7. SOLIDWORKS 2025 API contract ledger

The project-local `solidworks-api` MCP reports a SOLIDWORKS 2025 target using a compatibility-snapshot corpus. On 2026-07-17 its outputs were cross-checked with installed type-library/live-probe evidence where available.

### 7.1 Verified/documented

- `IDrawingDoc.InsertModelAnnotations4` accepts `Option`, `Types`, `AllViews`, `DuplicateDims`, `HiddenFeatureDims`, `UsePlacementInSketch`, `InsertAllAnnotations`, and `InsertAllReferenceGeometry`; `DuplicateDims=True` eliminates duplicates.
- Verified `swInsertAnnotation_e` values used by the project:
  - `swInsertDimensionsMarkedForDrawing = 32768`
  - `swInsertGTols = 32`
  - `swInsertNotes = 64`
  - `swInsertHoleWizardProfileDimensions = 65536`
  - `swInsertHoleWizardLocationDimensions = 131072`
  - `swInsertholeCallout = 1048576`
- `IModelDocExtension.AddOrdinateDimension(DimType, LocX, LocY, LocZ)` returns `swCreateOrdDimError_e`; the base entity is selected first, remaining entities are then selected, and `IModelDoc2.SetPickMode` ends the operation.
- `swHorizontalOrdinate = 3` and `swVerticalOrdinate = 2`.
- `IView.GetVisibleEntities2` requires a `Component2` plus `swViewEntityType_e`.
- `IView.GetVisibleComponents` supplies visible drawing-view components suitable for the proven part-drawing edge call, although those objects are intentionally incomplete for some other `IComponent2` methods.
- `IModelDocExtension.GetCorrespondingEntity2` maps a drawing-view edge, face, or vertex to the corresponding entity in the underlying part or subassembly and returns `Nothing` when none is found.
- `IEdge.GetCurveParams3`, called after `IEdge.GetCurve`, supplies `ICurveParamData` that can distinguish a complete circular edge from an arc; `ICurve.IsCircle` alone is insufficient.
- `IEdge.GetTwoAdjacentFaces2` returns the one or two body faces adjacent to an edge. `IFace2.GetSurface` plus `ISurface.IsCylinder` identifies cylindrical surface geometry, while `IFace2.FaceInSurfaceSense` is supplementary rather than manufacturing-intent proof.
- `IFace2.GetPatternSeedFeature` returns the seed feature for a patterned face. `IFace2.GetFeature` returns only the oldest owning feature, so ambiguous ownership requires `IFeature.GetFaces` plus `IFace2.IsSame` or equivalent verified face-set evidence.
- Verified `IFeature.GetTypeName2` mappings include `HoleWzd` to `IWizardHoleFeatureData2`, `AdvHoleWzd` to `IAdvancedHoleFeatureData`, `SketchHole` to `ISimpleHoleFeatureData2`, `Cut` to `IExtrudeFeatureData2`, and the documented circular/linear/curve/mirror pattern interfaces.
- `IWizardHoleFeatureData2.GetSketchPoints` returns centre-hole sketch points; this is semantically narrower than raw `ISketch.GetSketchPoints2`.
- `IFeature.IsSuppressed2` returns a Boolean array for the specified configuration option; verified `swThisConfiguration = 1`.
- Installed interop verifies `swSpecifyConfiguration = 3`; the r4 replacement passes the drawing view's `ReferencedConfiguration` as a one-element name array and fails closed if that state is not returned.
- `IView.ModelToViewTransform` can transform both a model-space centre through `IMathPoint.MultiplyTransform` and its axis through `IMathVector.MultiplyTransform`.
- `IModelDocExtension.AlignDimensions` returns Boolean and takes an alignment enum plus a distance.
- Verified drawing-view type values distinguish sheet, section, detail, projected, auxiliary, standard, named, relative, detached, and alternate-position views.
- Installed interop verifies `ISheet.GetTemplateName`, `GetSheetFormatName`, `SheetFormatVisible`, `TitleBlock`, `GetZoneMargin`, and `SetScale`, plus `ITitleBlock.GetExtents`.
- Installed interop verifies `IDrawingDoc.ActiveDrawingView` returns null when the sheet is active, `IAnnotation.GetType`/`GetAttachedEntities3`, and `ICustomPropertyManager.Get6`/`Add3`.
- Installed interop verifies `IDrawingDoc.CreateDetailViewAt4`, `IView.UseParentScale`, `ScaleRatio`, `Position`, and `GetDetail`, plus `IDetailCircle.GetView`, `GetDetailView`, `GetStyle`, `GetDisplay`, `GetProfileItemsCount`, `GetProfileItems`, `NoOutline`, `HasFullOutline`, and `JaggedOutline`. The r4 transaction requests standard circular details at 3:1 and fails when structural readback differs.

### 7.2 Live-proven narrowly

On the disposable P-0251-14A-001 front-view probe:

- `InsertModelAnnotations4` returned 11 annotations.
- `GetVisibleEntities2(Nothing, edge)` failed with a type mismatch.
- `RootDrawingComponent.Component` was unavailable for the attempted path.
- A component from `GetVisibleComponents` returned 39 edges.

This proves the corrected component-qualified call path on that view/build. It does not prove feature ownership, correct ordinate coverage, layout, title block, or the embedded VBA macro.

### 7.3 Installed-build behavior still requiring E4/E6 after E3

- installed-build/VBA proof of drawing-edge correspondence and selectable entity behavior on every intended fixture/view;
- installed-build tolerances for complete-circle closure, cylindrical agreement, and axis-normal qualification;
- returned suppression-array shape and pattern/mirror seed behavior on the three fixtures;
- actual selectable datum entities and model/view/sheet transforms;
- imported-dimension-to-feature/datum coverage reconciliation;
- section/detail source activation, segment/profile selection, exact coordinate
  frames, creation/readback, scale/label/outline, and cleanup behavior;
- annotation/text/leader extents and supported movement/rerouting;
- `AlignDimensions` spacing units in the installed build;
- template-specific title-block content, linked-note resolution, and approved property scope;
- barcode font/symbology and scan verification.

Section arrows/hatch and detail feature content/legibility remain E7 visual and
manufacturing-semantic gates, not E4 API-contract claims.

No new numeric enum or COM binding assumption may be added without checking this ledger against the MCP and, when binding behavior is uncertain, the installed 2025 build.

## 8. Acceptance gates

### 8.1 Global gates

Production acceptance requires all of the following:

1. The designated replacement `.swp` contains the complete frozen r4 set and
   its post-import re-export matches the r4 manifest.
2. The whole embedded VBA project compiles in SOLIDWORKS 2025.
3. Only the three authorized fixtures are executed.
4. Complete Immediate Window output, settings, QA, screenshots, and exports are retained.
5. Every API success claim is tied to a checked return/result.
6. Every required feature family has size/callout and location coverage or an explicit `FAIL`.
7. All ordinate groups use a recorded proven datum and eligible view.
8. No duplicates, collisions, clipping, sheet intrusion, or title-block intrusion remain.
9. Title data and general notes are present and sourced.
10. Each fixture passes visual and semantic comparison with this specification and its manual reference.

### 8.2 Fixture-specific acceptance

#### P-0251-14A-001

- Front, side, section J-J, and at least one clean undimensioned isometric view.
- Six-hole face family has one size/counterbore definition and complete symmetric location coverage.
- Side tapped-hole family is defined in the appropriate view/section.
- Stepped bore, `Ø47 H7` fit, `Ø40`, profile radius, and relevant thickness/step dimensions are present.
- Centre/bottom origin entities are proven and recorded.
- No ordinate appears on the section or isometric.

#### P-0252-01-001

- Manufacturing-useful face orientation, edge view, and clean orientation aid(s); no section by default.
- All required opening, tapped, blind, through, precision-fit, counterbore, radius, chamfer, and back-face families are distinguishable.
- Feature locations are complete without repeated equal coordinates or ambiguous leader association.
- Back-face operations remain explicit.
- The plan/edge annotation load fits inside the usable sheet without copying the reference’s congestion.

#### P-0252-01-013

- Manufacturing-useful orthographic set, section B-B, and clean undimensioned orientation aid(s).
- Threaded-hole size/drill depth/thread depth and location are complete.
- Datum A, applicable perpendicularity controls, `6.4 g6` fit, radii, angles, and profile dimensions are preserved when present in authoritative model annotations.
- Details C and D are both present at 3:1 and legibly expose the reference's small `7 x 4` tab/end geometry and `C0.5` chamfers.
- No ordinate appears on section, detail, or isometric views.

### 8.3 Required regression matrix

After each narrow fix, run the smallest focused authorized case. Before acceptance, cover:

- all three fixtures;
- model annotation import;
- Hole Wizard callouts off and on, with recorded mask and result delta;
- Bottom-Left, Center, and Top-Left datum choices, each with actual zero entity/coordinates;
- auto-arrange off/on where exposed by configuration;
- no section, one horizontal primary section, and one vertical primary section where meaningful for the fixture;
- supported orthographic/projected views and excluded sheet/isometric/section/detail views;
- duplicate model/ordinate coverage;
- title-block, note, barcode/part-ID, and sheet-boundary checks;
- final normal pick mode and empty selection state.

The earlier generic “simple part” and synthetic-plate tests are not authorized macro-run fixtures. They may be replaced by pure/fake-COM tests or require new user authorization.

## 9. Evidence retention

Each retained validation run must have a unique run directory under `test_assets/` containing, as applicable:

- manifest with timestamp, fixture, configuration, SOLIDWORKS build, template, macro identity, and settings;
- embedded/export synchronization evidence;
- compile result and highlighted error line if compilation fails;
- complete Immediate Window log;
- structured QA report;
- source/model feature and API-probe evidence;
- full-sheet screenshot plus close-ups of dense/failed regions;
- exported PDF/JPG/BMP when useful;
- comparison notes against the matching manual reference;
- retained failed output when it provides regression evidence.

Manual reference drawings must never be overwritten.

## 10. Clean-replacement implementation and remaining gates

The incremental repair backlog was retired after the current macro repeatedly failed to activate its drawing views and produced a dimensionless P-0251 sheet. The implementation strategy is now a complete replacement source set rather than further patching of the broken binary.

`src/target-spec-hybrid-v2/` implements the target architecture in nine standard modules plus three data/evidence classes. The forms remain the configuration front end, but fixture-controlled fields are locked and the unimplemented layout preview is disabled instead of making a misleading claim. The current source identity is `target-spec-hybrid-v2-2026-07-18-r18`; the original r4 handoff remains a frozen historical checkpoint. The replacement includes:

1. authorized-fixture, saved-part, template, scale, and drawing-document preflight;
2. explicit drawing-document activation before every view-scoped stage;
3. fixture-led orthographic orientation, including the runtime-proved P-0251-only rotation, title-block-rectangle-aware layout without silent scale changes, and one deterministic primary section;
4. the fixed annotation-import allowlist, duplicate elimination, and exact active-view selected-view transaction;
5. the complete retained ownership bridge: component-qualified visible model edge, full-circle proof, radius-matched adjacent cylinder, pattern seed or owning feature, exact supported feature type, normal-axis projection, and semantic-instance deduplication;
6. selectable projected corner/origin datum proof with `DatumNotSelectable` failure instead of nearest-hole substitution;
7. attached-entity X/Y coverage reconciliation and creation of only missing unique ordinate coordinates;
8. datum-first selection, checked `MultiSelect2` and final selection counts, decoded `AddOrdinateDimension` results, and normal pick-mode cleanup;
9. mandatory Pump Holder Details C and D from exact `*Bottom` profiles at independent 3:1 scale with structural postconditions;
10. configuration-first title properties, truthful asterisk-delimited part identification, structured fail-closed QA, and retained atomic text evidence; and
11. import guidance that preserves the special host component and treats the two form files as code snapshots rather than native designer exports.

The source-contract and hybrid-companion suite passes **61 offline tests** at
E2; separate installed-interop/MCP checks establish E3. Together they make r16 a
coherent compile-handoff candidate, not a compiled or accepted macro. The r4
source archive and 35-entry manifest remain frozen historical evidence. The
remaining critical path is:

1. deploy every dependent r16 component together into a recoverable candidate
   macro, verify readback, and compile the complete VBA project;
2. resolve D-04 and the named live-only coordinate, selection, attachment,
   template, title-link, and extent contracts;
3. run the focused P-0251 acceptance profile and retain full E6/E7 evidence;
4. advance to Base Plate and Pump Holder only after the preceding gate passes;
   and
5. pass the fresh-session three-fixture regression and reference comparison
   before any production claim.

`src/baseline-model-dims/` and the existing broken macro remain unchanged. Static tests and API reflection are not VBA compilation or visual acceptance.

## 11. Decision and policy log

This log includes both implemented provisional policies and unresolved user
inputs. Only unresolved rows require user action before their dependent gate.

| ID | Decision | Current provisional rule |
|---|---|---|
| D-01 | Must P-0252-01-001 and P-0252-01-013 reproduce both reference isometric views? | At least one is mandatory; add a second only when it materially exposes opposite-side machining and fits cleanly |
| D-02 | May one production drawing use multiple proven coordinate origins/feature-group datums, especially for P-0252-01-001? | Support the selected global origin first; do not invent additional origins without approved behavior |
| D-03 | Are Pump Holder Details C and D mandatory for first acceptance? | **Implemented r4 policy, pending user override/confirmation:** both are mandatory at 3:1. E4/E6 must prove profile-coordinate/API behavior; E7 must prove correct `7 x 4`/`C0.5` content and legibility |
| D-04 | What controlled `.drwdot`/`.slddrt` and title-property mapping is authoritative? | **Partially resolved:** the controlled `.DRWDOT` exists at `V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`, and r5 uses that fixed path. The template-linked sheet-format and exact linked-note/property/cell map still require live proof; no discovery or uncontrolled fallback is permitted |
| D-05 | What barcode symbology/font and scan standard are required? | Treat current asterisk-wrapped text as part identification, not a proven barcode |
| D-06 | What exact selectable entity represents Bottom-Left, Center, and Top-Left for each fixture/view? | Prove from model/drawing geometry; fail closed when unresolved |

Unresolved policies must be closed from the authorized models, controlled
template, and user input before their dependent capability can receive
production acceptance.

## Appendix A. Superseded assumptions

The following earlier statements are retained only as historical context and must not be used as current status:

- that the active importer still uses the old two-bit mask;
- that the exported active path still lacks per-view retry, component-qualified visible-edge retrieval, datum-first selection, API-based view filtering, or `SetPickMode`;
- that the latest embedded macro is necessarily older than the exports; the 16:28 binary save makes its exact source state unknown until re-exported;
- that a passing companion test or disposable probe proves VBA success;
- that all three references use one isometric or one barcode location;
- that the Base Plate’s crowded reference should be copied literally;
- that a generic unapproved part may be used for live macro regression.

Current embedded/runtime status and newer verified evidence always take precedence.
