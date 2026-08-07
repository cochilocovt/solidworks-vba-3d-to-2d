# What `src/baseline-model-dims/` needs to produce the reference drawing

Planning session, 2026-08-05. No source file was edited.

> **Live status 2026-08-06 (r24).** Phases 1 and 2 are complete; Phase 3 item
> 9 is complete. Gaps **A1-A6, A8, A10, B1, B2, C1, C2** are closed; **A7,
> A9, C3, C4 (partial), C5, C6, C7, D1-D6** and the whole of section 7 remain
> open, plus new item **C8** below. Both ordinate chains match the reference
> exactly and have since r20:
>
> | | trunk (r20-r24) | reference |
> |---|---|---|
> | Long axis | 160, 90, 50, 10, 0 | 160, 90, 50, 10, 0 |
> | Cross axis | 36, 15, 0, 15, 36 | 36, 15, 0, 15, 36 |
>
> The ordinate engine is done against this fixture. Two things closed since
> r20, both live-only findings with no static equivalent:
>
> - **HLR became an engine precondition, not a checkbox (r23).** r17's
>   `ResetGlobalConfig` default was never reached by an operator running the
>   form — `UserForm1.frm:265` seeds from a saved registry setting, and the
>   form always wins. `Module2_DrawingPipeline.ForceHlrForHarvest` /
>   `RestoreDisplayMode` now force `swHIDDEN` for the ordinate harvest
>   regardless of the operator's setting, and restore it after. Proven live
>   both ways: `HLR (already)` when ticked, `HLR forced (was 1)` when not,
>   byte-identical ordinate output either way.
> - **The stepped section cut (C1, C2) is closed, r24.** `PlanSteppedCut`
>   derives the cut path from the model's hole geometry instead of the
>   bounding-box midpoint, and the coordinate-frame question C2 raised is
>   answered: `ISketchLine.GetStartPoint2` readback matches the requested
>   model-mm coordinates with no view-scale factor applied. Full detail in
>   `SOLIDWORKS_API_VALIDATION.md`.
>
> What remains is everything that is not an ordinate or the section cut's
> geometry: hole callouts, the H7 tolerance, title block content, and which
> views get created at all (**C8**, new).
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
| A7 | Fixed 15 mm placement offset, no bounds check | unchanged through r24 | chains sit clear of the view, inside the frame, clear of the title block. **Still open** — the section view ran off the sheet as late as r23 (fixed by C1/C2, not A7: that fix changed the section's own placement formula, not general collision logic). Front-view dimension count fell 22 to 14 across r22-r24, which reduces how often A7 bites without closing it |
| A8 | ~~2-D proximity test used to dedupe a 1-D chain~~ **CLOSED r10** | dedup is per axis and 1-D, in `AddAxisCandidate` | a horizontal chain cares only about X; a Y difference must not keep two points that share an X |
| A9 | 1.5 mm tolerance is an unjustified magic number | `COORD_DEDUP_TOL_M`. **Still open**, unchanged through r24, but no longer scale-dependent: r13 normalises coordinates by `IView.ScaleDecimal`, which fixed the macro emitting different dimensions at different scales | derive from model/drawing resolution |
| A10 | ~~Chain failures only reach `Debug.Print`~~ **CLOSED r7-r15** | `OrdinateRunStatus` threaded to QA; plus a post-rebuild `IAnnotation.IsDangling` readback that reports what was actually created, not what was intended | QA must see them (§7) |
| A11 | ~~HLV silently degrades the ordinate engine~~ **CLOSED r23** | `ForceHlrForHarvest` reads `IView.GetDisplayMode2`, forces `swHIDDEN` before the harvest, `RestoreDisplayMode` after. `OrdinateRunStatus.HarvestDisplayMode` reports which path ran. Proven live both ticked and unticked (`macro_qa/20260806_151955`, `.../20260806_152553`) | HLV was never a valid harvest state — under it, hidden-line edges select and dangle (r17), or a wider candidate pool misroutes the end datum 43 mm inside the part (r22). r17's `ResetGlobalConfig` default never reached the operator; the form always wins |

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
| B1 | ~~`DuplicateDims=False`~~ **CLOSED, pre-r20 (exact revision not logged)** | `B:39`, `B:93` now pass `True`. Table text corrected 2026-08-06 - it still described the pre-fix state though the document's own top summary already listed B1 closed | see §1.2 |
| B2 | ~~`ImportHoleCallouts` form field is dead~~ **CLOSED, pre-r20** | `GetModelItemMask` (`B:190-205`) now gates `swInsertholeCallout` on `GlobalConfig.ImportHoleCallouts` (form default `True`). Table text corrected 2026-08-06 for the same reason as B1 | the checkbox now reaches the mask. Whether the checkbox itself reaches the operator reliably (the `UserForm1`/registry pattern that made A11 necessary) is untested for this control specifically |
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
| C1 | ~~Section line placed at the model bbox mid-plane~~ **CLOSED r24** | `PlanSteppedCut` (`Module2_DrawingPipeline.bas`) reads `Module3_ModelAudit.GetAllHoleLikeFeatures`, takes the widest hole-like feature as the bore and the across-axis coordinate shared by the most non-bore holes as the row, and cuts a 3-segment stepped line (bore leg, jog, row leg) through both. Intent, not luck. Live: `macro_qa/20260806_165529`, `Section cut: stepped, 3 segments` |
| C2 | ~~Section sketch created in raw model coordinates~~ **CLOSED r24** | Verified rather than assumed: `AddCutSegment` reads each segment's endpoints back via `ISketchLine.GetStartPoint2` after creation. Requested bore leg at Y=0, row leg at Y=15mm, jog at X=-27mm; readback `-127.4,0`, `-27,0`, `-27,15` — matches exactly, no 0.6667 view-scale factor applied anywhere. Drawing-view sketch geometry is stored at model scale; the view's own scale is a display-time transform only. Closes the CodeStack pages 9/30 question for SW2025 |
| C3 | Only one section is ever created | unchanged through r24 | `GetPrimarySectionSettings` reads `GlobalSections(1)` only, `B:273-275` — the form promises more |
| C4 | View placement is fixed sheet fractions | **partially addressed, r24**: the section view's own sheet placement was fixed — was `frontPos(0) + 0.18`, a blind 180mm offset that ran the section past the sheet border on r23; now `halfFront + halfSection + 15mm gutter`, derived from the part bounding box and sheet scale, confirmed by screenshot to sit inside the border. Every other view (front, top, bottom, left, right, back, iso) is still placed by fixed sheet fractions, `sheetW*0.42` etc., `B:137-167` — no collision logic, no title-block exclusion. **Still open** for everything but the section |
| C5 | `ConfigureView` is `On Error Resume Next` with every return ignored | unchanged through r24, `B:201-207` |
| C6 | No `IView.ProjectedDimensions` policy | unchanged through r24. CodeStack page 29 — affects whether section dims read true or projected |
| C7 | Section arrow direction unmanaged | unchanged through r24. reference J–J arrows point left; nothing in the pipeline controls this |
| C8 | View set doesn't match the reference's, and the standard-view naming doesn't match the reference's layout | **new, found in review 2026-08-06**. The reference has exactly four views: Left, Section J-J, Front, Isometric. The operator's r24 run enabled Front/Top/Bottom/Right/Left/Back/Isometric/Section — eight. Worse, the reference's side-elevation dimensions (`80`, `25`, `4.20`, `6.00`) landed on the view SOLIDWORKS calls **Bottom** in that run, not **Left** — SW's Top/Bottom/Left/Right naming is a standard projection relative to Front, not a guarantee that the same physical face lands in the same named bucket the reference sheet uses. Evidence: `macro_qa/20260806_170652`. Fix is two-part: determine which SW standard view is geometrically the reference's Left face, and stop creating views the reference doesn't have |

## 5. Title block and annotations — `Module7_TitleBlockEngine`

| # | Gap | Evidence |
|---|---|---|
| D1 | Seven title-block fields, written by code or not? | **Downgraded from "never written" on re-check, 2026-08-06.** Recent screenshots (r22-r24) show UNIT, PART NAME, DRA/DGN/CHD/APPD, HEAT TREATMENT, SURFACE TREATMENT, REVISION, SCALE all populated, matching the reference. `Module7_TitleBlockEngine` (`B:17-25`) still only *writes* Description, PartNo, Material, CustomerCode, Project, Qty, Mass, DrawnDate — so either the rest are template/part custom properties already set before the macro runs, or the sheet format resolves them independently. **Not verified which** — see D6, never actually checked. Do not credit the code for these fields without checking first |
| D2 | Mass is copied from a property, not computed | unchanged through r24, `B:23`. Reference shows `1.30`; live output shows `1296.82`. Open decision, deferred by the user at r7 (units mismatch, most likely g-vs-kg on the source property) — not defaulted. CodeStack page 15 gives the mass-properties route if a computed value is wanted instead |
| D3 | Notes text and placement do not match | unchanged through r24, confirmed present in every screenshot this session. `B:87-98` emits a `GENERAL NOTES` heading with numbered lines at hardcoded `(0.22, 0.03)`, **in addition to** the title-block info box's own three-line note text — the reference has only the latter, no heading, no second block. Open decision, deferred by the user at r7 — not defaulted |
| D4 | Barcode is plain text | unchanged through r24, confirmed in every screenshot (plain bold `*P-0251-14A-001*`, not barcode-typeface). `B:119` — reference uses a barcode typeface; needs `ITextFormat` font assignment |
| D5 | No read-back verification | unchanged through r24. `WriteDrawingProperty` `B:69-74` is `On Error Resume Next` with both return values ignored |
| D6 | Property→title-block linkage unverified | unchanged through r24 — still never checked. This is now the open question D1 depends on: whether the sheet format's fields reference `$PRP:"…"` is template-dependent |

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

**Investigated 2026-08-06/08, resolved: not usable, removed.**
`Module5_FallbackDimensionEngine.InsertHoleCalloutsForView` called
`IDrawingDoc.AddHoleCallout2` once per circular edge with no display-mode
guard (the same hidden-edge exposure A11 closed for ordinates, never applied
here) and no consolidation - it could only ever emit one callout per edge,
never the reference's grouped `6x Ø6.6` / `4x Ø4.2`. More decisively,
`AddHoleCallout2`'s MCP Remarks state the call requires a user to click OK
in a system dialog per invocation - not viable inside this project's
unattended `tools/production-runner` path. Removed, not fixed. Full
writeup: `SOLIDWORKS_API_VALIDATION.md`, "AddHoleCallout2 is not viable for
unattended automation".

**New live evidence, r25 (`macro_qa/20260808_041847`)**: the *other*,
correct, non-modal path was already firing and still produces nothing.
`GetModelItemMask` requests `swInsertholeCallout` by default, 10 of the
fixture's 12 detected holes are genuine Hole-Wizard features (both counts
newly instrumented this pass), and `IDisplayDimension.IsHoleCallout` across
every dimension in the drawing reads **0 of 20**. This is now measured
directly, not inferred from a screenshot. Cause not yet isolated - see
`SOLIDWORKS_API_VALIDATION.md` for the ranked candidates and the cheapest
next diagnostic (report whether `ImportHoleCallouts` actually reached the
mask for a given run, which nothing currently surfaces).

By r24, the ordinate chains and the section's stepped-cut geometry are the
two pieces of this document that are actually solid. Everything else
listed here — hole callouts, the H7 tolerance, five of six title-block
gaps, and now C8's view-set mismatch — is exactly as open as when this
document was written on 2026-08-05, or newly found. The live-run discipline
that closed A1-A11 and C1-C2 has not yet been pointed at any of it.

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
10. Sheet-aware placement (A7), tolerance derivation (A9). (A8 closed r10;
    stale reference to it here removed 2026-08-06.)

### Phase 4 — Tier C — **not started, 2026-08-06 re-check**

11. Feature-tree-derived requirements (§7). Not started. Every heuristic
    added r10-r24 (bore-as-widest-hole, row-as-shared-Y, per-axis datum
    rules) is tuned against this one fixture, not derived from the feature
    tree.
12. Structured placement report as the acceptance artefact (§4a item 6). Not
    started. Every "matches reference" judgement through r24, including the
    ones in this document, is a human visual read of a screenshot.
13. Fixture-allowlist replacement (§4a item 2). Not started; not designed.
    Still the only runtime safety guarantee against an unauthorized part.

Also not done: a live run against either of the other two authorized
fixtures (`P-0252-01-001`, `P-0252-01-013`). Every run referenced in this
document through r24 is `P-0251-14A-001` only — there is no evidence yet
that any of Phase 1-3's work generalizes even to Tier B.

## 9. Open questions for the user

1. ~~Which snapshot is the trunk?~~ **ANSWERED.** `baseline-model-dims`, per
   `CLAUDE.md`: "Trunk: `src/baseline-model-dims/`". `active-ordinate` is
   history; `target-spec-hybrid-v2` is archived.
2. **Is front-view-only ordinates a rule or an instance?** Operationally
   decided for this fixture — `Module8_ViewClassifier` implements
   front-only (r8) and the reference confirms it for `P-0251`. Whether it
   generalizes is unverified: no run yet against `P-0252-01-001` or
   `P-0252-01-013`. Still drives A5's general form.
3. **Where do the drawing-authored tolerances come from** (the `H7`), given
   the model has none? Still open. Memory `h7-fit-is-drawing-authored`
   confirms the model carries no tolerance and the H7 lives on a drawing
   reference dimension, but no drawing-side tolerance rule has been
   implemented. Drives §7.
4. **Is the reference's view layout a fixed template or a computed one?**
   Still open, now sharpened by C8: the reference has exactly four views
   (Left, Section J-J, Front, Isometric) and the macro's form can produce
   up to eight. Whatever the answer, "create every standard view the form
   offers" is not it.


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

Re-verified true 2026-08-06: `UserForm1` still absent from
`deployment-manifest.json`, still no `VERSION 5.00` designer block. This is
why r23 made HLR an engine precondition (`ForceHlrForHarvest`, A11) instead
of trying to fix the form's checkbox default — a form default cannot be
deployed or tested, so a correctness requirement cannot depend on one.

### r20 to r24, in brief

The outer-edge drawing convention (user-stated, 2026-08-06) was the last
piece the ordinate engine needed — applying it produced the exact reference
match at r20 and it has held through every run since, including under every
view-count and view-mode combination tried.

Two more live-only lessons, beyond the ones above:

- **A default that is never exercised is not a default.** r17 set
  `UseHLR = True` in `ResetGlobalConfig` and the document above called HLR
  "the default" for five revisions. It was never once the value an operator
  actually got, because the form always overrides it. A11's fix — force the
  precondition inside the engine, independent of any config value — is the
  general lesson: **a correctness requirement that can be silently
  overridden by an unverified, undeployed form is not actually enforced.**
- **Proving a coordinate transform needs a nonzero test case.** Every
  section cut before r24 sat at an offset of 0 on the axis in question,
  which reads back as 0 under any scale factor and any coordinate frame —
  it proved nothing about C2 for four revisions. The transform was only
  actually tested once a cut was placed away from zero and its readback
  checked against the request.
