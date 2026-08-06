# What `src/baseline-model-dims/` needs to produce the reference drawing

Planning session, 2026-08-05. No source file was edited.

> **Live status 2026-08-06 (r20).** Phases 1 and 2 are complete and Phase 3
> item 9 is complete. Gaps **A1-A6, A8, A10, B1, B2** are closed; **A7, A9**
> and the whole of section 7 remain open. Both ordinate chains now match the
> reference structurally:
>
> | | trunk r19 | reference |
> |---|---|---|
> | Long axis | 160, 90, 50, 10, 0 | 160, 90, 50, 10, 0 |
> | Cross axis | 36, 15, 0, 15, 36 | 36, 15, 0, 15, 36 |
>
> **Exact match at r20**, once the outer-edge drawing convention was
> applied. The ordinate engine is done against this fixture; what remains
> is everything that is not an ordinate, chiefly section 7.
>
> The per-gap table below carries individual status. Live API contracts are in
> `SOLIDWORKS_API_VALIDATION.md`; current open work is in `CURRENT_STATUS.md`.
> Section 12 records what the live runs taught that no static analysis could.

Target: `test_assets/reference_drawings/P-0251-14A-001.PNG`.
Subject: `src/baseline-model-dims/` (abbreviated **B** below).
Inputs reviewed: `docs/ORDINATE_GAP_ANALYSIS.md`,
`docs/CODESTACK_DRAWING_API_COVERAGE.md`, the baseline source, and MCP
contract lookups recorded in `docs/SOLIDWORKS_API_VALIDATION.md`
(section "Two baseline defects located by contract reading").

## 0. What the reference drawing actually contains

Read off the PNG, because the change list only makes sense against it.

**Four views**, left to right:

| # | View | Dimensions it carries | Style |
|---|---|---|---|
| 1 | Left (side profile) | `80`, `25`, `6`; leader callout `4x Ø4.2 ↧12.4 / M5x0.8-6H ↧10` | conventional linear + hole callout |
| 2 | `SECTION J-J` | `18`, `12`, `11.5`, `173.6`, `104.8`, `Ø40`, `Ø47 H7 +0.025/0.000` | conventional linear + diameter + fit tolerance |
| 3 | Front (main) | `R36`; leader callout `6x Ø6.6 THRU / ⌴Ø11 ↧6`; **two ordinate chains**; section line J–J | ordinate + radius + hole callout |
| 4 | Isometric | none | shaded, undimensioned |

**The two ordinate chains on the front view are the load-bearing detail:**

- **Vertical chain, right side**: `0`, `10`, `50`, `90`, `160`.
  Datum `0` sits on the **bottom edge of the part**. Values ascend to the bore
  centre at 160.
- **Horizontal chain, bottom**: `36`, `15`, `0`, `15`, `36`.
  Datum `0` sits on the **vertical centreline**. Values are symmetric —
  `±15` are the two hole columns, `±36` are the two silhouette edges.

Three consequences follow immediately, and they invalidate the current datum
design:

1. The X datum and the Y datum are **different entities**, chosen by different
   rules. One `DatumOrigin` setting cannot express this.
2. **Neither datum is a hole.** One is a part edge, one is a centreline.
3. The X chain **dimensions silhouette edges** (`±36`), not just holes.

**Title block fields present**: company, customer code, DRA / DGN / CHD / APPD
(name + date each), PROJECT, UNIT, PART NAME, PART NO, MATERIAL, MASS, QTY,
HEAT TREATMENT, SURFACE TREATMENT, REVISION, SCALE, "DO NOT SCALE DRAWING",
first-angle projection symbol.

**Notes block** (above title block, right): three unnumbered lines, no
"GENERAL NOTES" heading.

**Barcode**: `*P-0251-14A-001*`, bottom left, in a barcode typeface.

### 0.1 The 173.6 question is answered

`LONG_VERTICAL_REF_173_6` — the requirement the target-spec probe could not
satisfy — **is on the reference drawing**, as a vertical linear dimension in
`SECTION J-J`. It is not a spec error. The 2026-08-05 planning decision
(item 5 in `docs/R23_SCOPE_AND_GENERALIZATION_PLANNING.md` §4a) recorded
"check the reference drawing first"; that check is now done and the answer is
**code gap**.

What remains unexplained: the probe collected 19 distinct Y values in the
section view and found no pair 173.6 mm apart, with a total span of 201 mm.
So at least one endpoint of the reference's 173.6 dimension is not in the
collected Y set. Most likely the decoder is missing a silhouette or tangent
point that is not a collected edge endpoint. That is the thing to investigate
— not whether the dimension exists.

## 1. The two defects that most likely explain the reported symptoms

Both are contract-level findings, verified against MCP this session, recorded
in `docs/SOLIDWORKS_API_VALIDATION.md`. Neither is reproduced live yet.

### 1.1 "Ordinate dims setting breaks the macro" — missing `SetPickMode`

`IModelDocExtension.AddOrdinateDimension` Remarks, quoted:

> Selections made immediately after calling this method continue to add
> ordinate dimensions to the group of ordinate dimensions. When you finish
> adding ordinate dimensions to the group, use `IModelDoc2::SetPickMode` to
> return to the default selection mode.

`SetPickMode` appears nowhere in `src/baseline-model-dims/`.
`CreateOneOrdinateChain` (`B/Module5_FallbackDimensionEngine.bas:238-252`)
calls `AddOrdinateDimension`, calls `ClearSelection2`, and returns with the
group still open.

The horizontal chain runs first (`B/Module5_FallbackDimensionEngine.bas:86`).
The vertical chain's `MultiSelect2` (`B:239`) is then a selection made after
the call — per the contract it **appends to the open horizontal group**
instead of starting a vertical one. Everything selected later in the run is
made in the same mode, including auto-arrange
(`B/Module4_ModelItemImporter.bas:157`) and title-block work.

`ClearSelection2` clears the selection list. The contract does not say it
closes the group.

**Fix**: call `SetPickMode` after each `AddOrdinateDimension`, on both the
success and failure paths. Two lines. This is the single highest-value change
in this document.

### 1.2 "Hole callouts are repeated" — `DuplicateDims = False`

`IDrawingDoc.InsertModelAnnotations4` parameter 4:

- `DuplicateDims` (Boolean): "True to eliminate duplicate dimensions, false to
  allow duplicate dimensions"

The parameter name reads backwards from its effect. The description is the
contract.

`B/Module4_ModelItemImporter.bas:36-44` passes `AllViews=True` **and**
`DuplicateDims=False` — insert into every view, allow duplicates. The per-view
fallback (`B:90-98`) passes the same `False`.

**Fix**: pass `True`. Then verify against the fixture whether that alone
yields the reference's single consolidated `6x` / `4x` callouts, or whether
per-view callout targeting is also needed (§3.4).

### 1.3 Good news: the mask constants are correct

Every constant at `B/Module4_ModelItemImporter.bas:6-12` matches the
`swInsertAnnotation_e` MCP table exactly, as do the ordinate constants at
`B/Module5_FallbackDimensionEngine.bas:4-6`. The gap analysis was right to
refuse to trust them; they pass. No numeric correction is needed there.

## 2. Ordinate engine — `Module5_FallbackDimensionEngine`

| # | Gap | Evidence | Reference demands |
|---|---|---|---|
| A1 | ~~No `SetPickMode`~~ **CLOSED r7** | called unconditionally after every `AddOrdinateDimension`, success and failure paths | see §1.1 |
| A2 | ~~One datum for both axes~~ **CLOSED r10** | independent `AxisCandidates` per axis, independent datums in `ResolveAxisDatums` | X datum = centreline, Y datum = bottom edge — different entities, different rules |
| A3 | ~~Datum must be a circle~~ **CLOSED r10/r19** | datum kind is free; an **end** datum is further restricted to a straight edge (r19), a centreline datum is unrestricted | neither reference datum is a hole |
| A4 | ~~Only circular edges are candidates~~ **CLOSED r10** | straight model edges admitted as stations; the cross chain now reaches `36` | X chain dimensions the `±36` silhouette edges |
| A5 | ~~Ordinates applied to every non-ISO view~~ **CLOSED r8** | now `Module8_ViewClassifier.AllowsOrdinateDimensions`, front-only | ordinates on the front view **only**; section and left views use conventional dims |
| A6 | ~~View classification by name string~~ **CLOSED r8** | `IsIsoView` deleted; `Module8_ViewClassifier.ClassifyView` uses `IView.Type` + `IView.GetOrientationName` | see §6, now resolved with live evidence |
| A7 | Fixed 15 mm placement offset, no bounds check | unchanged | chains sit clear of the view, inside the frame, clear of the title block. **Still open** — the section view has run off the sheet in several runs |
| A8 | ~~2-D proximity test used to dedupe a 1-D chain~~ **CLOSED r10** | dedup is per axis and 1-D, in `AddAxisCandidate` | a horizontal chain cares only about X; a Y difference must not keep two points that share an X |
| A9 | 1.5 mm tolerance is an unjustified magic number | `COORD_DEDUP_TOL_M`. **Still open**, but no longer scale-dependent: r13 normalises coordinates by `IView.ScaleDecimal`, which fixed the macro emitting different dimensions at different scales | derive from model/drawing resolution |
| A10 | ~~Chain failures only reach `Debug.Print`~~ **CLOSED r7-r15** | `OrdinateRunStatus` threaded to QA; plus a post-rebuild `IAnnotation.IsDangling` readback that reports what was actually created, not what was intended | QA must see them (§7) |

**A8 is a real correctness bug independent of the reference drawing.**
`CreateOneOrdinateChain` dedupes correctly on the single coordinate array
(`B:220-225`), but the *candidate* set was already filtered by a 2-D test at
`B:69`. Two holes that share an X but differ in Y both survive
`HasNearbyPoint`, which is right for the vertical chain and wrong for the
horizontal one — the horizontal chain then relies on `B:221` to drop the
second. That works, but it means the two chains have inconsistent notions of
"distinct", and the datum index resolved once at `B:84` is used for both.

## 3. Model-item import — `Module4_ModelItemImporter`

| # | Gap | Evidence | Note |
|---|---|---|---|
| B1 | `DuplicateDims=False` | `B:39`, `B:93` | see §1.2 |
| B2 | `ImportHoleCallouts` form field is dead | `GetModelItemMask` `B:175-182` ORs `swInsertholeCallout` unconditionally | the checkbox never reaches the mask |
| B3 | No per-view callout targeting | `AllViews=True` `B:38` | reference puts `6x Ø6.6` on the front view and `4x Ø4.2/M5` on the left view — each where that hole axis reads as a circle |
| B4 | Auto-arrange selects without view scoping | `swAnn.Select3 True, Nothing` `B:157` | `Nothing` SelectData = unscoped. CodeStack's recurring lesson is to scope with `ISelectData.View` |
| B5 | `AlignDimensions` spacing `0.06` unverified | `B:161` | 60 mm if metres. Units and behaviour unconfirmed |
| B6 | Per-view fallback triggers only on a zero total | `B:48-50` | a partial whole-drawing result leaves the remaining views bare and unnoticed |

On **B3**: consolidation into `6x` / `4x` comes from the Hole Wizard feature
carrying instances, with the callout inserted once. Fixing B1 may be
sufficient. Treat that as a hypothesis to test on the fixture, not a
conclusion.

## 4. Views and section — `Module2_DrawingPipeline`

| # | Gap | Evidence |
|---|---|---|
| C1 | Section line placed at the model bbox mid-plane | `midX = (bbox(0)+bbox(3))/2`, `B:225` — coincides with the reference's J–J by luck, not intent; the reference cut is placed to pass through the hole columns |
| C2 | Section sketch created in raw model coordinates | `B:239-241` uses `bbox` values directly with `SketchManager.CreateLine`. CodeStack pages 9 and 30 both say the model→view (or model→view→sheet) transform must be composed. **Verify which context `CreateSectionViewAt5` expects.** |
| C3 | Only one section is ever created | `GetPrimarySectionSettings` reads `GlobalSections(1)` only, `B:273-275` — the form promises more |
| C4 | View placement is fixed sheet fractions | `sheetW*0.42`, `partH*scaleVal*0.75`, `B:137-167` — no collision logic, no title-block exclusion |
| C5 | `ConfigureView` is `On Error Resume Next` with every return ignored | `B:201-207` |
| C6 | No `IView.ProjectedDimensions` policy | CodeStack page 29 — affects whether section dims read true or projected |
| C7 | Section arrow direction unmanaged | reference J–J arrows point left; nothing in the pipeline controls this |

## 5. Title block and annotations — `Module7_TitleBlockEngine`

| # | Gap | Evidence |
|---|---|---|
| D1 | Seven title-block fields are never written | reference has UNIT, PART NAME, DRA/DGN/CHD/APPD, HEAT TREATMENT, SURFACE TREATMENT, REVISION, SCALE; `B:17-25` writes Description, PartNo, Material, CustomerCode, Project, Qty, Mass, DrawnDate |
| D2 | Mass is copied from a property, not computed | `B:23`. Reference shows `1.30`. CodeStack page 15 gives the mass-properties route |
| D3 | Notes text and placement do not match | `B:87-98` emits a `GENERAL NOTES` heading with numbered lines at hardcoded `(0.22, 0.03)`; reference has three unnumbered lines, no heading, above the title block |
| D4 | Barcode is plain text | `B:119` — reference uses a barcode typeface; needs `ITextFormat` font assignment |
| D5 | No read-back verification | `WriteDrawingProperty` `B:69-74` is `On Error Resume Next` with both return values ignored |
| D6 | Property→title-block linkage unverified | whether the sheet format's fields reference `$PRP:"…"` is template-dependent and has never been checked |

## 6. View classification — RESOLVED r8 (2026-08-06)

`IView.Type` returns `swDrawingViewTypes_e` (Sheet=1, Section=2, Detail=3,
Projected=4, Auxiliary=5, Standard=6, Named=7, Relative=8, Detached=9,
AlternatePosition=10). But every view this pipeline creates comes from
`CreateDrawViewFromModelView3`, so front / top / right / isometric were
expected to share one `Type`. **`Type` alone will not separate the isometric
from the front view.** Pair it with `IView.GetOrientationName` ("Gets the
predefined name of this view").

That expectation was confirmed, not assumed. The r8 run printed the raw
returns for every view: all five model views came back `Type=7`
(`swDrawingNamedView`), and `GetOrientationName` round-tripped the exact
string passed to `CreateDrawViewFromModelView3` (`*Front`, `*Bottom`,
`*Right`, `*Left`, `*Isometric`). `Section View J-J` returned `Type=2` with an
empty orientation, matching the documented empty-string case.

Implemented in `Module8_ViewClassifier`. Full evidence in
`SOLIDWORKS_API_VALIDATION.md`, "r8 view classification characterised".

The roster is emitted on **every** run by `Module6_QAEngine.BuildPerViewSummary`
and carries the raw `Type=` / `Orientation=` values, not just the derived role.
A build where these contracts differ falsifies itself in the run record rather
than failing silently.

## 7. The structural gap that no reordering of the above fixes

Count the producers. The baseline has exactly two:

1. `InsertModelAnnotations4` — emits whatever the model marked for drawing.
2. The ordinate fallback — emits chains through hole centres.

Now count what the reference needs that neither can produce:

- `Ø47 H7 +0.025/0.000` — memory `h7-fit-is-drawing-authored` records that
  P-0251's model carries **no** tolerance; the H7 lives on a drawing reference
  dimension. `InsertModelAnnotations4` can never emit it. Needs a drawing-side
  tolerance rule.
- `R36`, `Ø40` — radius and diameter dims; only arrive if marked in the model.
- `173.6`, `104.8`, `18`, `12`, `11.5`, `80`, `25`, `6` — conventional linear
  dims in the section and left views. Same condition.

So: **everything on the reference drawing except the two ordinate chains and
the two hole callouts is a conventional dimension that appears only if the
model happens to have marked it.** If it is not marked, nothing in the
baseline creates it, and QA counts dimensions rather than checking which ones.

That is the real completeness gap, and it is exactly what the Tier C
"feature-tree derived requirements" decision (§4a item 3 of
`R23_SCOPE_AND_GENERALIZATION_PLANNING.md`) was reaching for. It cannot be
closed by patching Module 4 or Module 5.

## 8. Recommended sequence

Rationale for the ordering: phases 1 and 2 are contract-verified or
mechanical, need no new architecture, and each maps to a symptom the user
reported. Phase 3 needs a decision from the user, not more API lookups.
Phase 4 is the Tier C rearchitecture already agreed.

### Phase 1 — contract-verified fixes (no new design) — **DONE r7**

1. `SetPickMode` after every `AddOrdinateDimension`, success and failure paths
   (A1).
2. `DuplicateDims = True` in both `InsertModelAnnotations4` calls (B1).
3. Wire `GlobalConfig.ImportHoleCallouts` into `GetModelItemMask` (B2).
4. Scope auto-arrange selection with `ISelectData.View` (B4).
5. Return a status from the ordinate entry points; surface it in QA (A10).

Then run the probe runner and judge against the fixture. Items 1 and 2 are the
ones expected to change observable behaviour.

### Phase 2 — read-only probe, then classification — **DONE r8**

6. Probe what `IView.Type` and `IView.GetOrientationName` return for each view
   this pipeline creates (§6).
7. Replace `IsIsoView` with the API-backed test (A6).

### Phase 3 — needs a user decision first — **items 8 and 9 DONE (r8, r10-r19); item 10 open**

8. **Per-view dimension-style policy** (A5). Which views get ordinates, which
   get conventional dims. The reference says front-only, but that is one
   drawing.
9. **Per-axis datum contract** (A2, A3, A4). X and Y resolved separately,
   against edges and centrelines as well as holes.
10. Sheet-aware placement (A7), tolerance derivation (A8, A9).

### Phase 4 — Tier C

11. Feature-tree-derived requirements (§7).
12. Structured placement report as the acceptance artefact (§4a item 6).
13. Fixture-allowlist replacement (§4a item 2).

## 9. Open questions for the user

1. **Which snapshot is the trunk?** Three partial implementations now exist:
   `baseline-model-dims` (cleanest control flow, correct import mask, missing
   `SetPickMode`), `active-ordinate` (defensive scaffolding, *has*
   `SetPickMode`), and `target-spec-hybrid-v2` (fixture-hardcoded literals
   like the 173.6 requirement). Recommendation: **baseline as trunk**, port
   the defensive scaffolding from active, drop the hardcoded spec from
   hybrid-v2. Not yet decided.
2. **Is front-view-only ordinates a rule or an instance?** Drives A5.
3. **Where do the drawing-authored tolerances come from** (the `H7`), given
   the model has none? Drives §7.
4. **Is the reference's view layout a fixed template or a computed one?**
   Drives C4.


## 12. What the live runs taught, r7 to r19

This document was written statically. Every defect that actually blocked the
engine needed a run to find, and three of them were invisible to any amount of
source reading. Recorded here because the same traps will recur.

### The engine failed by succeeding

Not one of these raised an error. Each produced a plausible drawing.

| Defect | How it presented | Found by |
|---|---|---|
| `If Not <comBool>` yields -2 | 349 edges, **0 circles**, on a part with 12 holes | counting edges in the report |
| `<comBool> = True` never matches | drawing with **zero** dimensions | a run after an apparently equivalent refactor |
| `GetVisibleEntities2(Nothing, …)` | no edges at all | trying the component route |
| Wrong `swDisplayMode_e` values | views rendered in the wrong mode for 15 runs | the constant-provenance test |
| `ModelToViewTransform` carries scale | **different dimensions at different scales** | the operator running at 1:2 |
| Hidden-line edges undimensionable | two ordinates reading `0.00` | an HLR/HLV A-B run |
| `ICurve.IsCircle` true for arcs | datum 60 mm inside the part | reading the Remarks |
| `Set` on a `propertyput` property | error 91, carried as an "API refusal" for 11 revisions | reading the Remarks |

### Two rules that came out of it

1. **An error at a COM boundary is not evidence the callee refused.** Check
   the VBA binding before recording an API limitation.
2. **`GetVisibleEntities2` visibility is not the same question as "can this
   entity carry a dimension".** An edge drawn in hidden-line font is returned,
   selects, and yields `swCreateOrdDimErr_Success` — and the ordinate dangles.

### The instruments that actually resolved things

Three diagnoses of the dangling ordinates were inferred from station counts
and sheet positions. All three were wrong. What settled it was measurement:

- the intended station list, in model mm, printed every run;
- `IAnnotation.IsDangling` + `IDimension.GetValue3` read back **after the
  rebuild** — the flag is not set before one;
- the per-view roster carrying raw `IView.Type` / `GetOrientationName`.

Each cost one function. Each mechanism guess cost a full deploy-run cycle.

### A gap this document did not anticipate

`ResetGlobalConfig` is **not** the user-visible default. `UserForm1` seeds its
controls from saved registry settings and writes them back over `GlobalConfig`
on OK, so the form wins on every run that shows it. The forms are outside
`deployment-manifest.json` and cannot be imported, so operator-visible
defaults can only be changed by hand-editing in the VBE, and no test covers
them. See `Architecture.md`.
