# Changelog

## 2026-08-06 (54) - r20: both ordinate chains match the reference exactly

| | trunk r20 | P-0251-14A-001 |
|---|---|---|
| Long axis | 160, 90, 50, 10, 0 | 160, 90, 50, 10, 0 |
| Cross axis | 36, 15, 0, 15, 36 | 36, 15, 0, 15, 36 |

Not an API finding. A drawing convention, stated by the user: **a dimension
always goes to the outer edge.**

A chamfered corner presents two parallel axis-parallel straight edges - the
true outer extreme and the chamfer's inner boundary. Both are legitimate
stations and they fall well inside `COORD_DEDUP_TOL_M`, so they merge. Which
one survived was decided by array order out of `GetVisibleEntities2`, not by
geometry, and it was landing on the inner edge.

`PreferOuterCandidates` promotes each station's retained runner-up wherever it
lies further from the axis midpoint, before datum resolution. Holes are
unaffected - concentric circles of one hole share a centre exactly.

`Outer-edge promotions: 2`, one per axis, and the centreline datum's
`offsetFromTarget` went from 0.50 mm to 0.00 mm.

**My arithmetic was wrong and the run was right.** Before deploying I reasoned
that a 0.5 mm chamfer could only explain half the 1 mm discrepancy. It
explained all of it: both the datum edge and the opposite extreme moved
outward, and the cross-axis midpoint moved with them.

### Ordinate engine status

Gaps A1-A6, A8, A10 closed. A7 (placement bounds) and A9 (tolerance
derivation) remain. The engine now produces the reference drawing's two
ordinate chains exactly, on the front view only, with zero dangling
dimensions.


## 2026-08-06 (53) - r18/r19: both ordinate chains match the reference

| | trunk r19 | P-0251-14A-001 |
|---|---|---|
| Long axis | 159, 89, 49, 9, 0 | 160, 90, 50, 10, 0 |
| Cross axis | 36, 15, 0, 15, 35 | 36, 15, 0, 15, 36 |

Five stations per axis, zero dangling.

**`ICurve.IsCircle` is true for arcs.** Its Remarks say to use
`IEdge::GetCurveParams2` to tell a full circle from an arc; the trunk never
did, so P-0251's rounded end was treated as a hole and its arc *centre* became
the long-axis datum - the chain started 60 mm inside the part.

**r18 over-corrected and r19 fixed it.** r18 made two changes; only one was
needed:

| change | effect |
|---|---|
| end datum must be a straight edge | fixed X: 89, 49, 9, 0 |
| exclude arcs as stations | broke Y: 21, 0, 30, 50 |

The rounded end's arc centre lies exactly on the part's axis of symmetry and
was the only entity there. An arc's centre is a true, dimensionable
coordinate - it is what an ordinate on that edge reads. The r17 defect was not
that the centre is wrong but that it was mistaken for the part's *extreme*,
which the straight-edge rule fixes on its own. r19 keeps arcs as stations.

### Correction to entry 52

Entry 52 said "HLR becomes the default". That is wrong as written.
`UserForm1` seeds its checkbox from a saved registry setting
(`ReadBoolSetting("UseHLR", False)`) and writes it back over `GlobalConfig` on
OK, so the form wins on every run that shows it. `ResetGlobalConfig` is the
no-form fallback, not the user-visible default, and the same applies to its
other fields. The form lives outside the deployment manifest, so changing what
the operator sees is a separate job. Corrected in the code comment.

### Open, unexplained

Every long-axis station is exactly 1 mm below its reference counterpart, and
the cross-axis chain matches on one side (36) and is 1 mm short on the other
(35). One candidate is the drawing's own note - "All corners are chamfered
0.5 x 45 deg" - which would put the selected straight edge on the chamfer
rather than the true face extreme. **Not tested.**


## 2026-08-06 (52) - r16/r17: dangling root cause found; HLR is the default

Two findings, both from following the API-lookup contract the new gates
enforce.

### ISelectData.View error 91 was VBA syntax

Open since r5, recorded twice in `SOLIDWORKS_API_VALIDATION.md` as an
unexplained API refusal, and worked around for eleven revisions. The property
is `propertyput`, not `propertyputref`: VBA needs a plain assignment and
raises 91 on `Set`.

```vba
swSelData.View = swView       ' works
Set swSelData.View = swView   ' raises 91
```

Every selection between r5 and r15 was unscoped. **An error raised at a COM
boundary is not evidence the callee refused.**

### Hidden-line edges cannot hold an ordinate

The dangling-station root cause, wrong three times before this. Confirmed by a
zero-code experiment - the same deployed r16 binary run twice, differing only
in the form's HLR setting:

| | HLV | HLR |
|---|---|---|
| Edges | 64 (35 circular) | **39 (22 circular)** |
| X stations | 9 | **5** |
| Y stations | 7 | **5** |
| Dangling | 2 | **0** |

An edge drawn in hidden-line font is not *completely* obscured, so
`GetVisibleEntities2` returns it. It selects, `MultiSelect2` counts it,
`AddOrdinateDimension` returns `swCreateOrdDimErr_Success` - and the ordinate
dangles. Every gate the engine had said yes.

**`GetVisibleEntities2` visibility is not the same question as "can this
entity carry a dimension".**

### Station counts now match the reference

| | ours (HLR) | reference |
|---|---|---|
| Cross axis | 36, 15, 0, 15, **35** | 36, 15, 0, 15, 36 |
| Long axis | 0, 70, 110, 150, 159 | 0, 10, 50, 90, 160 |

Five stations per axis, as the reference has. The cross-axis chain is
structurally identical bar the 35/36 asymmetry. The long axis has the right
count and the right 40 mm pitch between intermediate stations, offset by 60 mm
because the datum is still the bore centre rather than the part end - the left
end is a rounded arc and arcs are not admitted as stations.

### r17

`UseHLR` now defaults True, with the tradeoff recorded in the code: HLR also
removes the stepped bore from the front view, which is what the reference
drawing does - it dimensions that bore in SECTION J-J. One line to revert, and
the form option is unchanged.

Also fixed at r16: three `swDisplayMode_e` constants (see entry 51). The
isometric view now renders shaded with edges for the first time.


## 2026-08-06 (51) - r15: the sheet carries no wrong dimensions

`readback: 12 dims, 0 dangling`. Every ordinate on the front view is correct.

**The dangling flag is a post-rebuild property.** r14 pruned inside the
ordinate stage and found zero dangling annotations; the QA readback, later in
the same run and after `ForceRebuild3`, found two in the same view. Neither
`Select3` nor `DeleteSelection2` had refused - there was nothing yet to
select. r15 moved the prune downstream of the rebuild and it worked first
time: `2 found, 2 deleted (0 select-refused, 0 delete-refused)`.

Any future attachment check must sit after a rebuild, or it reads clean on
broken geometry.

Three revisions went into this one defect and each of my proposed mechanisms
was wrong: shared entities (r11, refuted by the readback), select/delete being
refused (r14, refuted by the counters), then rebuild timing. Each refutation
came from a measurement that cost one function. The mechanism guesses cost a
run apiece.

### Still open

The two stations at X=55.00 and X=135.00 are **deleted, not explained**. The
report says so rather than quietly dropping them.


## 2026-08-06 (50) - r11: one entity, one chain

The r10 sheet rendered three ordinates as `0.00` in the dangling colour while
the report listed their true offsets, so creation was placing them correctly
and the attachment was failing.

**Cause.** A circular edge is a station on both axes. The X chain runs first
and consumes the entity; the Y chain then re-selects an entity an open
ordinate group already owns. Worked back from the r10 data: Y's two edges are
its extremes (0.00 and 71.00), so Y=36.00 is a hole - the bore - and X=0 sits
on the bore centre. One entity, two chains.

**Fix.** An entity belongs to one chain only. Each station keeps an
`Alt` entity at the same coordinate, which a stepped bore or a counterbore
supplies naturally through its concentric circles, and the second chain
attaches to that instead. The station survives; only the entity changes.
`HolesOnlySubset` carries alternates through, and a failed chain releases its
claims before the retry so substitutions are not made for stations that were
never drawn.

**The prediction was numeric, and it matched.** The diagnosis was inferred
from station counts and sheet positions, not from reading entity identities,
so the run reports the count:

```
Shared entities: 3 substituted, 0 still shared
```

Three substitutions against three dangling ordinates, and nothing left
sharing. Sheet confirmation that the `0.00` values are gone is the remaining
check.

Unchanged and still open: the X datum sits on the bore centre rather than the
part extreme (a rounded end is an arc, and arcs are not admitted as stations);
station counts are too high against the reference (X 9 vs 5, Y 7 vs 5); the
default `Center` datum contract has still not executed live; and
`GetDisplayDimensions` reports 14 against 16 ordinate labels on the sheet.


## 2026-08-06 (49) - r9/r10: per-axis datum contract; a Boolean rule widened

Gaps A2, A3, A4 and A8 in `BASELINE_TO_REFERENCE_DRAWING_GAP.md`.

- **A2** - X and Y carried one shared 2-D datum index. They are independent
  problems with different answers, so each axis now has its own candidate set
  and its own datum (`AxisCandidates`).
- **A3** - the datum no longer has to be a circle.
- **A4** - straight model edges are admitted as stations. A line whose two
  view-space endpoints share an X is a station on the X chain, and vice versa.
  Endpoints come from `IEdge.GetStartVertex`/`GetEndVertex` + `IVertex.GetPoint`
  and are transformed as points, which avoids pushing a direction vector
  through `IView.ModelToViewTransform` - a contract this repo has not
  established.
- **A8** - deduplication was a 2-D proximity test. Two holes sharing an X but
  40 mm apart in Y are two points to that test and one station to the X chain.
  Now per-axis and 1-D.

### r9 was a regression, and it widened a standing rule

r9 produced a drawing with **zero dimensions**. While restructuring, the
proven `If swCurve.IsCircle = False Then GoTo NextEdge` was rewritten as
`If swCurve.IsCircle = True Then`. Those are not mirror images:

| form | circular edges found in the same view |
|---|---|
| `If Not swCurve.IsCircle` | 0 (r6, over 349 edges) |
| `If swCurve.IsCircle = True` | **0** (r9) |
| `If swCurve.IsCircle = False` | 35 (r8, r10) |

VBA `True` is `-1`; the returned value is truthy but not `-1`. The safe form is
`Not (<comBooleanCall> = False)`. The 2026-07-31 rule named only `If Not
<comBooleanCall>`; it now covers **any** comparison other than `= False`.

### r10 result

```
Ordinate edges seen: 64 (circular: 35, linear: 29)
Datum contract: CornerBottomLeft
  X: 9 stations (holes=6, edges=3), datum Hole:offsetFromTarget=0.00mm
    offsets(mm): 0.00, 43.00, 45.00, 55.00, 70.00, 110.00, 135.00, 150.00, 159.00
  Y: 7 stations (holes=5, edges=2), datum Edge:offsetFromTarget=0.00mm
    offsets(mm): 0.00, 12.40, 21.00, 36.00, 51.00, 59.60, 71.00
Ordinate chains created: 2 of 2 attempted
```

**A linear edge is a valid ordinate base entity on this build.** The engine was
written to retry with holes-only candidates if the richer set was rejected and
to count the retries; there were none, and the Y datum is an edge. Restricting
candidates to circles was self-imposed, not an API constraint.

### The r8 "three zeros" defect is located, not fixed

Printing the intended stations was the point of the `offsets(mm)` line, and it
worked. The sheet and the report agree on every station *position* and
disagree on three *values*: X=55.00 and X=135.00 render `0.00`, as does
Y=36.00, all three in the dangling-dimension colour. Creation places them
correctly; the attachment fails for a subset of entities. Narrow question now.

### Also open

- The X datum sits on the bore centre while the part extends ~56 mm further
  left. A rounded end is an arc, and arcs are not admitted as stations, so the
  left outline is not a candidate.
- Station counts are too high: X 9 against the reference's 5, Y 7 against 5.
  Admitting every axis-parallel edge over-corrected; stations need a selection
  rule, not just a wider net.
- The default `Center` datum contract did not execute - the operator chose
  Bottom-Left on the form - so the long-axis/centreline rule that matches the
  reference is still unproven live.


## 2026-08-06 (48) - r8: view classification is API-backed and characterised

`IsIsoView` is deleted. It tested `InStr(viewName, "ISO")` against names
SOLIDWORKS generates as `Drawing View1`..`View5`, so it never matched and r7
put ordinate chains on the isometric view. Replaced by
`Module8_ViewClassifier`, which keys on `IView.Type` for the structural cases
and `IView.GetOrientationName` for the orientation cases.

### The run is the probe

Rather than a separate read-only probe round-trip, `BuildPerViewSummary` now
prints the raw API returns next to the derived role, so every run records the
evidence for its own classification:

```
Drawing View1 | Type=7 | Orientation=*Front | Role=Front | ordinates=allowed | 9 dims
Drawing View5 | Type=7 | Orientation=*Isometric | Role=Pictorial | ordinates=skipped | 0 dims
Section View J-J | Type=2 | Orientation=(empty) | Role=Section | ordinates=skipped | 0 dims
```

This closes the §6 caveat in the gap doc. All five model views return `Type=7`
(`swDrawingNamedView`), confirming the pre-probe expectation that `Type` alone
cannot separate a front view from an isometric. `GetOrientationName`
round-trips the exact string handed to `CreateDrawViewFromModelView3`.

### Result

| | r7 | r8 |
|---|---|---|
| Views ordinated | 4 (incl. isometric) | 1 (front) |
| Chains | 8 of 12 | **2 of 2** |
| Chain failures | 4 (`swCreateOrdDimErr_OrdFailure`) | 0 |
| Front-view dims | — | 9 |

Gaps A5 and A6 closed. `AllowsOrdinateDimensions` refuses pictorial views
unconditionally but **allows** `VIEW_ROLE_UNKNOWN`: if classification breaks,
extra ordinates are visible and recoverable, whereas zero ordinates resembles a
working run that quietly did nothing — the exact failure mode the `Not IsCircle`
bug had.

### Not fixed, and not claimed fixed

- The four `swCreateOrdDimErr_OrdFailure` chains from r7 were on non-front
  views. They are **not diagnosed** — this run simply never attempted them.
- `Set ISelectData.View` still raises error 91. Report still reads
  `Selection scope: Unscoped(err=91)`.

### New defects, observed on the r8 sheet

1. The horizontal chain carries two spurious `0.00` ordinates alongside the
   real `0` datum (`0, 0.00, 70.00, 110.00, 0.00, 150.00`). A chain should
   have one zero.
2. `Section View J-J` is placed off the right edge of the sheet frame.
3. The `*Right` and `*Left` views are on the sheet per the roster but not
   visible inside the frame — consistent with the same placement defect.
4. The vertical chain reads `23.60 / 15.00 / 15.00 / 23.60` where the
   reference reads `36 / 15 / 0 / 15 / 36`. The `15` values match. `23.60`
   against `36` is consistent with gap A4 — only circular edges are
   candidates, so the chain terminates on the outermost hole rather than the
   silhouette edge. Consistent with, not proof of.
5. Both deferred decisions are now visually confirmed on the sheet: the title
   block shows `MASS(KG): 1296.82`, and the general notes appear twice — once
   in the sheet-format box, once from `InsertNotes`.


## 2026-08-06 (47) - r7 live: ordinate engine works, 8 chains created

First working ordinate run on the trunk. Four revisions of live diagnosis,
each one blocked by a different defect underneath the previous one.

| Rev | Result |
|---|---|
| r4 | 6 views threw. Staged capture named it: `err=91` at `CreateSelectData` |
| r5 | Sub-staged. Narrowed to `SetSelectDataView`, both operands non-Nothing |
| r6 | Component route + non-fatal fallback. 349 edges read, **0 circular** |
| r7 | Boolean fix. 349 edges, **127 circular**, **8 chains created**, 33 dims, PASS |

Three live API findings, recorded in `SOLIDWORKS_API_VALIDATION.md`:

- `GetVisibleEntities2(Nothing, ...)` yields nothing on this build. Resolving
  the component via `IView.GetVisibleComponents` yields 349 edges.
- `Set ISelectData.View = <view>` raises error 91 with both operands proven
  non-Nothing, after a successful `ActivateView`. Worked around by making it
  non-fatal and relying on view activation for scoping. **Unresolved.**
- `If Not <ICurve.IsCircle>` rejected all 349 edges. The Boolean contract
  recorded on 2026-07-31 (`Not` yields -2, treated as True) silently disabled
  the entire engine. One line changed to `= False` and chain creation began.
  Treat any `If Not <comBooleanCall>` in this codebase as a defect.

`SetPickMode` is finally exercised: 8 chains built in succession without the
append cascade the Phase 1 fix was written to prevent.

### Confirmed live, not yet fixed

- **The isometric view received ordinate chains.** `IsIsoView` matches on the
  name containing "ISO", but SOLIDWORKS auto-names these views
  `Drawing View1`..`View5`. Gap A6 in the gap doc, now proven rather than
  predicted. The API-backed replacement (`IView.Type` +
  `GetOrientationName`) is the fix.
- Ordinates are applied to every view; the reference uses them on the front
  view only (gap A5).
- Chain values do not match the reference. Datum resolution and the
  candidate set are both still wrong (gaps A2, A3, A4).
- 4 of 12 chains failed with `swCreateOrdDimErr_OrdFailure`, last in
  `Section View J-J`. Not diagnosed.
- Mass reads 1296.82 against the reference's 1.30 (grams, not kg).
- General notes duplicated over the title block.

## 2026-08-05 (46) - trunk moved to baseline, hybrid-v2 archived, Phase 1 landed

User decision: `src/baseline-model-dims/` becomes the trunk;
`target-spec-hybrid-v2` is archived. Revision reset to
`trunk-2026-08-05-r2`. **No live run yet — everything below is static and
contract evidence only.**

### r2: the baseline's template path was wrong

`TEMPLATE_PATH` was `V:\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`,
missing the `VEEMAP` segment. User confirmed the archived form
(`V:\VEEMAP\SW_data\...`) is correct. `GetValidDrawingTemplatePath` checks
`Dir$()` and falls through on a miss, so the baseline never raised an error —
it silently built every drawing on the SOLIDWORKS default template instead of
the controlled one. That alone would explain a missing title block and
sheet format on any baseline-produced drawing.

### Two defects found by reading MCP contracts against the baseline

- **Ordinate chains never closed their group.**
  `IModelDocExtension.AddOrdinateDimension` Remarks: selections made after
  the call keep appending to the open group until `IModelDoc2::SetPickMode`
  is called. `SetPickMode` did not appear anywhere in the baseline. The
  horizontal chain ran first, so the vertical chain's `MultiSelect2` appended
  to the horizontal group instead of starting its own, and every later
  selection in the run inherited the mode. Leading candidate for the reported
  "ordinate dims setting breaks the macro". Fixed in
  `Module5_FallbackDimensionEngine.CreateOneOrdinateChain`, unconditionally,
  on the failure path too.
- **Model-item import was told to allow duplicates.**
  `InsertModelAnnotations4` parameter 4 is `DuplicateDims`, documented "True
  to eliminate duplicate dimensions, false to allow" — the name reads
  backwards from the effect. The baseline passed `False` alongside
  `AllViews=True`, inserting the same annotation into every view. Leading
  candidate for the repeated hole callouts. Both call sites now pass `True`.

Confirmed correct and left alone: every `swInsertAnnotation_e` mask member
and every ordinate constant in the baseline matches the MCP enum tables.

### Also in Phase 1

- `ImportHoleCallouts` was a dead checkbox — `GetModelItemMask` OR-ed
  `swInsertholeCallout` unconditionally. Now gated on the form field.
- Auto-arrange selected annotations with `Select3 True, Nothing`, i.e.
  unscoped. Now scoped with `SelectData.View`.
- Ordinate failures reached only `Debug.Print`, so QA could report `PASS`
  with every chain failed. Added `OrdinateRunStatus`, a decoded
  `swCreateOrdDimError_e` message, and a fail-closed QA verdict.

### Trunk plumbing

- Stripped `Attribute` lines from all 10 deployable files. The baseline was
  **not deployable as it stood** — the deployer's hygiene contract forbids
  them. Files are pure ASCII, no BOM.
- Ported `Module21_EvidenceSink` verbatim and a trimmed `Module20_ProbeRunner`
  (compile gate and production pre-flight only). The nine `R23_Probe*` entry
  points called archived Module10–19 and were not ported, so **the probe
  runner's probe stage is inert**; its deploy and compile stages work.
- Added `MACRO_SOURCE_REVISION`, the three fixture constants,
  `IsAuthorizedFixture`, and a fail-closed authorization gate in `main` —
  the baseline had none, and both runners mirror that list.
- Manifest retargeted: 38 components to 12, `revisionPattern` to `trunk-*`.
  Same pattern updated in the two probe manifests.

### Test suite

Live suite 34 tests, passing. The 31 files testing archived code moved to
`archive/target-spec-hybrid-v2-tests/` (585 tests; 22 fail by design — see
`archive/README.md`). Two hygiene tests and the deployment-tooling test now
read the manifest instead of hardcoding hybrid-v2's file count.

### Reference-drawing findings

`LONG_VERTICAL_REF_173_6` **is on the reference drawing**, in SECTION J-J.
Code gap, not spec error. The front view's ordinate datums are per-axis — X
on the centreline, Y on the bottom edge — and neither is a hole, which
invalidates the single-`DatumOrigin` design. Full gap list in
`docs/BASELINE_TO_REFERENCE_DRAWING_GAP.md`.

## 2026-08-05 (45) - r62 live: ANNOTATION_EXTENTS proved, three placement defects opened

Run `macro_qa/20260805_071309_P-0251-14A-001`. Deploy verified at
`embedded=target-spec-hybrid-v2-2026-08-05-r62`, pre-flight
`ready=True|verdict=Clean`. Required failing stages **6 to 5**.

- **`ANNOTATION_EXTENTS` is PROVED.** The post-layout clamp did what it was
  built for: `SECTION_ANNOTATION_CLAMP_PASS|annotations=6|moved=3|`
  `stillOutside=0`, and RD2 went `fromY=0.289235` to `readbackY=0.275000`
  with `nowInside=True`. RD3/RD4/RD5 were recorded `AlreadyInside` and left
  untouched, so the radial/diametric `SetPosition2` refusal documented in
  the API notes was never exercised. It remains untested live.
- **The stage passing did not make the sheet right.** The user's screenshot
  showed what the gate cannot: three placement defects, all introduced by
  this work.
- **Defect 1. RD1 and RD2 now sit on the same point.**
  `ANNOTATION_GEOMETRY` reports `x=0.206692|y=0.275000` for both, so 12.00
  is hidden underneath 18.00. The clamp drives every violator to the
  identical boundary point instead of separating them.
- **Defect 2. RD3 and RD5 also share a point**, `x=0.179752|y=0.160060`.
  This one predates r62 and is an r61 regression: `LaneTextPoint` handles
  `LANE_BORE_SIDE_A` and `LANE_EXTERIOR_VERTICAL_INNER` in one `Case`, so
  giving each lane its own ordinal made them identical. The r60 global
  counter had been separating them by accident.
- **Defect 3. The SECTION J-J label was moved 52 mm onto the view for no
  reason.** `DetailItem927` went `y=0.036285` to `0.088749`. It was never in
  violation: `CONTROLLED_REGIONS|content=0.010000,0.010000,0.410000,`
  `0.287000`. The clamp uses the **view-usable** box, whose bottom is
  `TitleBlockTop + margin` - a constraint on views, not on annotations.
  `ANNOTATION_EXTENTS` enforces the content border and the title-block
  rectangle, and the label was inside both.
- Fix direction for r63: clamp against the regions the stage actually
  enforces (content border inset by `LAYOUT_MARGIN_M`, plus title-block
  avoidance); step a second annotation along the free axis instead of
  stacking it on the first; and give `LANE_EXTERIOR_VERTICAL_INNER` its own
  offset from `LANE_BORE_SIDE_A`.
- Unchanged: `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`,
  `SECTION_DIMENSIONS` (`satisfied:5/missing:2`),
  `MANUFACTURING_DEFINITION`, `FINAL_QA`.

## 2026-08-05 (44) - r62 clamps the section annotations after the view is placed

Offline only. 619 offline contracts pass; eleven mutations against the new
guards were each applied and each failed the suite before being reverted.
Not yet deployed or run.

- **The r61 clamp was applied at the wrong time, not computed wrongly.**
  `ClampSectionAnnotationsIntoUsableArea` is a new pass that runs after the
  structural grid, after auto-arrange and after the title block, and pulls
  the section view's annotation origins back inside the proved usable box.
  It is the fourth and last mutating procedure in Module10, refuses without
  `allowMutation`, and refuses again without `LayoutBoundariesProven`.
- **It moves annotation origins only.** No view move, no rescale; the
  content-envelope repositioning and rescaling the 2026-08-04 user decision
  retired stay uncalled. Scoped to the one section view on the P-0251
  fixture, whose only annotations are the five this run creates.
- **The readback is the verdict, not the setter result.**
  `IAnnotation::SetPosition2` returns sheet coordinates from the lower-left
  corner - the same frame `GetPosition` reports, so the two compare
  directly - but the Help says a constrained annotation is placed "as near
  as possible", and that radial and diametric dimensions cannot be
  positioned this way at all. Both look like a call that returned. Each
  annotation records `setterResult`, `readbackX/Y` and `nowInside`, and the
  pass counts `moved` against `stillOutside`.
- An annotation already inside is recorded `action=AlreadyInside` and left
  alone; a needless mutation hides the ones that mattered. The z coordinate
  is carried through from the read position rather than zeroed.
- The creation-time clamp stays as a starting bound. It is no longer relied
  on for the guarantee.
- One real defect caught offline by an existing guard: the new error
  handler read `Err.Number` after `SafeAnnotationName`, whose
  `On Error Resume Next` resets `Err`, so every clamp failure would have
  reported as error 0. Captured before the name helper now.

## 2026-08-05 (43) - r61 live: the verdict is honest; the clamp box was stale

Run `macro_qa/20260805_070039_P-0251-14A-001`. Deploy verified at
`embedded=target-spec-hybrid-v2-2026-08-05-r61`, pre-flight
`ready=True|verdict=Clean`.

- **The verdict defect is fixed and proved.**
  `R23_SECTION_DIMENSIONS|requirements=7|satisfied=5|missing=2` now agrees
  with the QA stage exactly. No `RequirementFlagged:...:NoImportedDimension`.
- **The per-lane ordinal works.** `BoreSideA`, `BoreSideB` and
  `ExteriorVerticalInner` all report `laneOrdinal=1`, where r60 gave them
  3, 4 and 5. RD2 recorded `clamped=True` at `textY=0.275000`, which is
  `UsableTop` to the micron.
- **`ANNOTATION_EXTENTS` failed anyway, and the r60 diagnosis was wrong.**
  Auto-arrange never touched this view: `ACTIVATE_VIEW|operation=Dimension
  arrange` names Drawing View1, Drawing View2 and Drawing View4 only. The
  mover is `ArrangeViewsInMeasuredGrid`, which runs *after*
  section-dimension creation:

  ```
  creation-time outline  0.289060,0.039385,0.318940,0.252265
  LAYOUT_MOVE readback   0.191752,0.053620,0.221632,0.266500
  delta                 -0.097308,          +0.014235
  ```

  RD2 requested `(0.304000, 0.275000)` plus that delta is
  `(0.206692, 0.289235)` - the violation line to six decimals. RD1 lands at
  `0.278500`, under `ContentBorderTop` 0.287, which is why only RD2 trips.
  A usable box measured before the view is placed is stale by construction.
- Required failing stages unchanged from r60: `NATIVE_CALLOUT_COVERAGE`,
  `VIEW_PROJECTION`, `SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`,
  `ANNOTATION_EXTENTS`, `FINAL_QA`. `SECTION_DIMENSIONS` holds at
  `satisfied:5`.

## 2026-08-05 (42) - r61 fixes the two defects r60's own evidence exposed

Offline only. 606 offline contracts pass; the eight mutations covering the
new guards were each applied and each failed the suite before being
reverted. Not yet deployed or run.

- **Placement no longer stacks off the drawing.** `LaneTextPoint` counted a
  *global* ordinal, so the gap grew 12/24/36/48/60 mm across five different
  lanes and the second entry in `LANE_ABOVE` started above the usable area.
  Auto-arrange finished the move and RD2 landed at `y=0.290500`,
  `region=ZonedBorder`. The ordinal is now per lane
  (`NextLaneOrdinal`, keyed by lane name), and the resulting point is
  clamped into the proved `evidence.Usable*` box before it is reported.
  `SECTION_DIM_PLACEMENT` gains `laneOrdinal`, `appliedOrdinal`, `usable=`
  and `clamped=`.
- Fails closed: without `LayoutBoundariesProven` the ordinal stops at 1
  rather than stacking against a boundary nothing measured, and the proof
  string says `usable=Unproved`.
- **Module2's verdict line was reading stale requirement state.** It said
  `R23_SECTION_DIMENSIONS|satisfied=0` with `RequirementFlagged:...:`
  `NoImportedDimension` while the QA stage independently proved
  `satisfied:5` on the same drawing. Reconciliation writes
  `NoImportedDimension` into `Failures` *before* creation runs, creation
  sets `Matched` without clearing it, and `VerifySectionDimensions` treats
  a non-empty `Failures` list as unsatisfied whatever the counts say.
  Re-reading the dimensions alone did not fix it. Module2 now rebuilds the
  requirements and re-reconciles them against the finished drawing, which
  is what Module19 already did - hence the two disagreeing numbers.
- No API contract changed; no new SOLIDWORKS call was introduced.

## 2026-08-05 (41) - r59/r60 live: five section dimensions created and verified

Runs `macro_qa/20260805_062318_P-0251-14A-001` (r59, read-only inventory)
and `macro_qa/20260805_063139_P-0251-14A-001` (r60, creation).

- **`SECTION_DIMENSIONS` goes `satisfied:0/missing:7` to
  `satisfied:5/missing:2`.** Five `SECTION_DIM_CREATED` rows, each with the
  nominal read back and matched: RD1 0.018000, RD2 0.012000, RD3 0.040000,
  RD4 0.047000, RD5 0.104800. No `CreatedButRejected`.
- **The r58 arc defect is fixed and it cost exactly what was predicted.**
  Axis-parallel classification now runs on `Type 0` polyline records only;
  an arc's tessellated chords are counted as `arcTessellationSegments` and
  no longer offered as straight-edge coordinates. `LOWER_WALL_STEP_11_5`
  lost its pair as a result, which is the correct answer: r58 would have
  attached it to the bore's tessellation and measured the wrong thing while
  looking right on the sheet.
- The two missing requirements are honest absences, skipped by name:
  `LOWER_WALL_STEP_11_5` (no pair once the tessellation stopped counting)
  and `LONG_VERTICAL_REF_173_6` (no pair of the nineteen distinct Y values
  differs by 0.1736). Five right dimensions beat seven including a wrong
  one.
- **Two defects, both introduced by this work**, each found in the run's
  own evidence rather than by inspection: the unbounded placement stack
  that regressed `ANNOTATION_EXTENTS` from PROVED to FAILED, and Module2's
  stale-state verdict line contradicting the QA stage. Both fixed in r61,
  entry (42).
- Required failing stages: `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`,
  `SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`, `ANNOTATION_EXTENTS`,
  `FINAL_QA`.

## 2026-08-05 (40) - r58 live: all ten candidates selectable; an arc is masquerading as a wall edge

Run `macro_qa/20260805_061804_P-0251-14A-001`. Five required stages still
fail, as expected - r58 creates nothing.

- **Every one of the ten candidate entities selected and proved its
  owner**: `selectable=True|route=Selectable|owner=Section View J-J` on all
  ten rows, no `EntityIsNothing`. The entity array is live for these
  records in a cut view, so the `Null` case the Help warns about did not
  arise. **Dimension creation is reachable.**
- **One mapping defect, found by reading the record indices.** Record 25
  supplies the Ø47 radius (0.023500) *and* the `x=0.008500` side of
  `LOWER_WALL_STEP_11_5`. Both cannot be true of one entity. The cause is
  in `InventorySectionGeometry`: an arc record carries a tessellated point
  array as well as its GeomData, and the axis-parallel classifier walks
  those points, so a chord of the bore's tessellation that runs parallel to
  an axis is recorded as a straight-edge coordinate. A linear dimension
  attached there would dimension the bore, not the wall.
- Fix identified, not yet made: classify axis-parallel segments from Type 0
  records only. An arc's tessellated chords approximate a curve; they are
  not edges. `LOWER_WALL_STEP_11_5` may lose its pair as a result, which
  would be the correct answer rather than a regression.
- Record 2 is shared by `OVERALL_THICKNESS_18` and `BORE_STEP_DEPTH_12`,
  which is plausible for a single long edge bounding both and is not
  evidence of a fault.

## 2026-08-05 (39) - r58 a measured section coordinate now reaches a selectable entity

Read-only, nothing created; deployed and run the same day, live result in
entry (40).

- **The blocker is selection, not geometry.** r57 proved six of seven
  requirements have their measuring geometry in Section View J-J, all
  `geometryTrusted=True`. `CreateSectionDimension` refuses with
  `reason=NoEntitiesSelected`, and nothing joined a measured coordinate to
  a selectable drawing entity.
- **The join was already in the data.** `GetPolylines7`'s return value is
  the entity array, positionally paired with the polyline records - r57
  confirmed it exactly (`records=79|entities=79|recordsMatchEntities=True`).
  Each distinct coordinate now stores the record that produced it, and that
  index is the entity index. `AddDistinctValue` appends the value and its
  record together or not at all; `FindValueNear` and `FindPairWithSpan`
  return the records they matched.
- `SECTION_REQ_CANDIDATE` gains `|recordA=|recordB=`, and a new
  `SECTION_ENTITY_SELECT|view=|key=|side=|record=|selectable=|route=`
  reports whether each candidate can actually be selected.
- **`ProveSectionEntitySelection` is Module13's Route D, unchanged in
  substance**: `IView.SelectEntity(entity, False)`, then
  `ISelectionMgr.GetSelectedObjectsDrawingView2` to prove the owning view
  before the object is trusted, `WrongOwner:` when it is not. It refuses a
  pre-existing selection with `RefusedPreexistingSelection:count`, bounds
  the record index against the entity array, and clears the selection on
  every path including the error handler.
- Attempted only when the requirement's geometry was `found` under a clean
  decode, so at most twelve selections per run.
- **`route=EntityIsNothing` is a measurement, not a failure.** The Help
  says the entity array carries `Null` where a polyline renders something
  no edge backs, which is expected for section faces in a cut view. Whether
  the six found requirements land on live entities is the question this
  iteration exists to answer.
- The pipeline now passes the drawing `ModelDoc2` - the selection manager
  lives on that interface, not on `DrawingDoc`.
- Static verification: **587/587**. Mutation-verified: dropping the record
  alongside the value, appending instead of replacing the selection,
  removing the pre-existing-selection refusal, attempting selection on
  untrusted geometry, removing the entity-array bound, and removing the
  document hand-off each fail the suite. One test bug was found and fixed
  in the writing: splitting the selector on `Failed:` matched
  `SelectEntityFailed:` first, so the error-handler assertions were reading
  the wrong text.

## 2026-08-05 (38) - r57 live: clean decode, and six of seven requirements have proved geometry

Run `macro_qa/20260805_055822_P-0251-14A-001`, invoked by the user. Stage
table unchanged: five required stages still fail.

- **First clean decode since the inventory existed.**

  ```text
  SECTION_GEOM_SUMMARY|decodeStatus=Complete|records=79|entities=79
    |recordsMatchEntities=True|doubles=2799|consumed=2289|trailing=510
    |trailingAllZero=True|arcs=18|polylines=61|points=454|error=0
  ```

  The padding hypothesis is confirmed by measurement rather than assumed:
  `trailingAllZero=True`. `consumed=2289` matches `9*79 + 12*18 + 3*454`
  exactly, and the record count matches the entity count.
- **`geometryTrusted=True` on every requirement row for the first time.**
  Six of seven have their measuring geometry present in Section View J-J:

  | requirement | nominal | evidence |
  | --- | --- | --- |
  | `OVERALL_THICKNESS_18` | 0.018 | X pair -0.009 / 0.009 |
  | `BORE_STEP_DEPTH_12` | 0.012 | X pair 0.009 / -0.003 |
  | `LOWER_WALL_STEP_11_5` | 0.0115 | X pair -0.003 / 0.0085 |
  | `INNER_BORE_D40` | 0.040 | radius 0.020 |
  | `FIT_BORE_D47_H7` | 0.047 | radius 0.0235 |
  | `LOWER_VERTICAL_REF_104_8` | 0.1048 | Y pair 0.041 / -0.0638 |
  | `LONG_VERTICAL_REF_173_6` | 0.1736 | **absent** |

  The radii `0.020000` and `0.023500` are now trusted evidence rather than
  the output of a walk that had already failed.
- `LONG_VERTICAL_REF_173_6` has no candidate: the section's extreme Y
  values are 0.1005 and -0.1005, a span of 0.201, and no pair of the
  nineteen distinct Y values differs by 0.1736. Open.
- `SECTION_DIMENSIONS` remains `requirements:7/satisfied:0` with
  `sectionDimensions:0` - nothing creates them. That is now the blocking
  work, and its inputs are proved for the first time.

## 2026-08-05 (37) - r57 the array is zero-padded; the r56 reversal was my error

Run `macro_qa/20260805_054951_P-0251-14A-001` settled it, and it reverses
entry (36). **The r54 entity-count bound was correct and I removed it on a
bad inference.**

- The r56 window shows **every double from index 2289 to the end of the
  array reads `0.000000000`**. The unbounded walk parsed 510 zeros as 56
  phantom records of stride 9 - type 0, GeomDataSize 0, six zero style
  fields, NumPolyPoints 0 - reaching the reported `records=135`, with 6
  doubles left over that cannot complete a record. **That is the whole of
  `Desynchronized:StyleAt2801`**, unexplained since r53.
- So 79 real records consume exactly 2289 doubles
  (`9*79 + 12*18 + 3*454`), `entities=79` is the true record count, and the
  Help's positional pairing holds. The 510 remaining doubles are padding.
- **My error at r56**: the r55 run stopped at 79 records with 510 doubles
  left and I read that as proof more records existed, without checking what
  those doubles contained. The r55 window already printed zeros at index
  2289 and I did not look at them. The arithmetic I cited was correct and
  the conclusion drawn from it was not.
- r57 restores the bound and **verifies the padding instead of assuming
  it**: `trailingAllZero=` is computed, and a non-zero tail is reported as
  `TrailingDataAfterEntities` rather than silently trimmed. Assuming the
  tail is what produced the wrong reversal; a bound that discards real
  records must be visible.
- `consumed=` is now reported alongside `trailing=`, so the walk's own
  arithmetic is in the log rather than reconstructed afterwards.
- Consequence: `decodeStatus=Complete` should be reachable for the first
  time, which would make the section coordinates trustworthy -
  `geometryTrusted=True` - including the radii `0.020000` and `0.023500`
  that have carried `geometryTrusted=False` since r53.
- Static verification: **573/573**.

## 2026-08-05 (36) - r55/r56 the decode was never misaligned; my control was wrong

**r55 not deployed as intended - it failed to compile.** `lower` was used
in `InventorySectionGeometry` without being declared; the pattern was
copied from `EmitSectionLineDecode`, which does declare it, and the
original walk here used only `i` and `upper`. `Option Explicit` rejected it,
`R23_PREFLIGHT_END|ready=False|reason=CompileNotClean` was logged and `main`
was never invoked. The revision was bumped rather than reused because r54's
source had already been written into `Fable.swp`. Fixed and run as r55:
`macro_qa/20260805_054524_P-0251-14A-001`.

**r56 removes the entity-count bound added at r54.** The r55 run disproved
the reasoning behind it.

- `SECTION_GEOM_DESYNC|status=StoppedAtEntityCount|stoppedAt=2289|recordsDecoded=79`
  with `trailing=510`, and the arithmetic closes to the double:
  `9*79 + 12*18 + 3*454 = 2289`. A misaligned walk does not produce a
  consumption that matches its own record, arc and point counts exactly.
- The window confirmed it independently. The record at index 2244 reads
  `type 0, GeomDataSize 0, six style scalars, NumPolyPoints 2`, a stride of
  15, and the next record begins at 2259 - exactly the reported
  `lastRecordStarts=2274,2259,2244`.
- So **alignment was never lost**. A section view simply holds more
  polyline records than entity-array entries, which is what one expects
  when cut edges have no model edge behind them. The Help's
  positional-pairing statement does not extend to a cut view.
- **The entity count was the wrong control and truncated a correct
  decode.** It is now reported as context (`recordsExceedEntities=`) and
  gates nothing. The control is exact consumption: `Complete` requires
  `trailing=0`, and every field the walk advances over is still
  range-checked.
- The r53 `Desynchronized:StyleAt2801` remains unexplained by this and is
  the thing the next run's window will settle - 2801 is two past the end,
  so either the final record is truncated in the array or a small slip
  occurs somewhere after index 2289.
- Static verification: **572/572**.

## 2026-08-05 (35) - r54 the polyline decoder stops lying and says where it broke

**Not deployed at time of writing.** Read-only. The r53 decode reported
`Desynchronized:StyleAt2801` with `records=135` against `entities=79`, and
still printed seven radii and six `found=True` verdicts derived from that
walk.

- **The walk is bounded by the entity count.** The Help states the polyline
  data and the returned entity array are positionally paired, so 135
  records from 79 entities cannot be true. The loop now stops at the entity
  count and reports `StoppedAtEntityCount` instead of running to the end of
  the array producing a number that is impossible by construction.
- **A clean decode must consume the array exactly.** `trailing=` is
  reported, and `Complete` is downgraded to `RecordCountMismatch` or
  `TrailingDoubles` when the record count or the consumed length disagrees.
  Previously only a range guard could contradict the decode.
- **A failed decode now prints the raw values around the failure**:
  `SECTION_GEOM_DESYNC|status=|stoppedAt=|recordsDecoded=|lastRecordStarts=`
  plus two `SECTION_GEOM_WINDOW` dumps of 36 doubles each - one at the
  third-most-recent record header, one at the stop point. 2799 doubles
  cannot be dumped whole, and without a window the next iteration would be
  guessing at the layout again.
- **Untrusted coordinates are labelled at every consumer.**
  `SECTION_GEOM_X/Y/R` carry `|decodeStatus=`, and every
  `SECTION_REQ_CANDIDATE` carries `|geometryTrusted=|decodeStatus=`. In r53
  only the summary line said the walk had failed, so six `found=True` rows
  read as findings when they were not.
- **The layout itself is not in doubt and was not changed.** Against the
  r50 and r52 arrays it closes to the double: 38 records x 9 fixed fields +
  6 arcs x 12 GeomData + 3 x 214 points = 1056, exactly the array length.
  Something in the richer r53 section deviates from it; the window is what
  will identify that.
- Static verification: **573/573**. Mutation-verified: removing the
  entity-count bound, removing the trailing-doubles downgrade, dropping
  either raw window, and dropping the `geometryTrusted` label each fail the
  suite.

## 2026-08-05 (34) - r53 live: the bore is in the section; two stages regressed

Run `macro_qa/20260805_051116_P-0251-14A-001`. Failed required stages
**6 to 5**. `ORDINATE_SCHEME` FAILED to PROVED.

- **The offset option did what the enum said it would.**
  `SECTION_CREATE_OPTIONS|options=2|offsetSection=True|segments=3`, and the
  section view roughly doubled: `entities=38` to **79**, `doubles=1056` to
  **2799**, `arcs=6` to **18**. Those three counts come straight from the
  API return, not from any decode of ours. The section is now hatched and
  shows the bore step and the counterbore pockets.
- **The section's radii now include `0.020000` and `0.023500`** - Ø40 and
  Ø47 - which never appeared in any previous run, and six of the seven
  `SECTION_REQ_CANDIDATE` rows report `found=True` against two before.
  **These numbers are indicative, not proved** - see the decode failure
  below.
- **The polyline decoder tripped its own guard**:
  `decodeStatus=Desynchronized:StyleAt2801|records=135|entities=79|recordsMatchEntities=False|doubles=2799`.
  It walked off the record layout and stopped rather than inventing
  geometry, which is the control working as designed, but every coordinate
  list from this run comes from that walk and cannot be trusted until the
  desynchronization is diagnosed. The entity and double counts are
  unaffected.
- **Two stages regressed and the cause is the same change.**
  `VIEW_PROJECTION` accepted projections 8 to **6**, locations without 3 to
  **5**; `NATIVE_CALLOUT_COVERAGE` incomplete 1 to **2**; and
  `MANUFACTURING_DEFINITION` complete 2 to **1**. The M5 family lost the
  attachment it gained at r49: the two near-side M5 holes projected in the
  old projection-section view and do not project in the offset one.
- `SECTION_DIMENSIONS` is unchanged at `requirements:7/satisfied:0` with
  `sectionDimensions:0`. Nothing creates them - `ReconcileR23SectionDimensions`
  only reconciles, and `CreateSectionDimension` is never called from the
  production route. That is now the largest remaining piece of work, and
  for the first time the geometry it needs appears to be present.

## 2026-08-05 (33) - r53 the section is cut as an offset section

One argument changed, on the r52 evidence and the user's instruction;
deployed and run the same day, live result in entry (34).

- `CreateSectionViewAt5` now receives
  `swCreateSectionView_OffsetSection` (MCP corpus value 2, **verify in the
  SW2025 Object Browser**) instead of `0`. Documented as: *"If set, then an
  aligned section view is created (two lines at an angle); if not set, a
  normal projection section view is created."*
- The r52 decode is the reason and it is recorded at the constant:
  `GetSectionLineInfo2` returned exactly the requested three-segment path,
  including the r51 overshoot, while the section view held the
  counterbore-column features and no bore. Bore at transverse 0.000,
  counterbores at -0.015, so the cut in use was segment 3's alone.
- `SECTION_CREATE_OPTIONS|view=|options=|offsetSection=True|segments=|source=swCreateSectionViewAtOptions_e`
  is emitted at creation. Which option was used decides what the cut
  contains and nothing else in the report would show it.
- **No other option bit was added.** `Partial`, `DisplaySurfaceCut`,
  `ChangeDirection` and `ScaleWithModel` each change what the section
  shows; a test asserts only the offset bit is set and that
  `sectionOptions` is assigned exactly once.
- The mutation boundary is untouched: still one procedure that can change a
  drawing, still refusing without explicit authorization and a resolved
  path.
- **Unverified**: whether the offset bit alone is enough or
  `swCreateSectionView_NotAligned` is also needed on this build. The run
  answers that.
- Static verification: **569/569**. Mutation-verified: setting the constant
  to 0, restoring the literal `0` argument, and dropping the evidence line
  each fail the suite.

## 2026-08-05 (32) - r52 live: the line is exactly right; `Options=0` throws the jog away

Run `macro_qa/20260805_050411_P-0251-14A-001`. Stage table unchanged (six
failing) as expected - r52 changed no behaviour. **The decode answered the
question the last two iterations could not.**

- **The drawing holds precisely the path we asked for.** Decoded from
  `IView.GetSectionLineInfo2`:

  ```text
  index=1|start=-0.102000000,0.000000000|end=0.008000000,0.000000000
  index=2|start=0.008000000,0.000000000|end=0.008000000,-0.015000000
  index=3|start=0.008000000,-0.015000000|end=0.088000000,-0.015000000
  ```

  Segment lengths 0.110, 0.015, 0.080 match the waypoint spacing exactly,
  and segment 1 runs 0.040 past the bore centre at -0.062 - the r51
  overshoot, present in the drawing. **The path was never the defect.**
- **`CreateSectionViewAt5` is called with `Options = 0`.** MCP corpus,
  `swCreateSectionViewAtOptions_e`: `swCreateSectionView_OffsetSection = 2`
  - *"If set, then an aligned section view is created (two lines at an
  angle); if not set, a normal projection section view is created."* So
  SOLIDWORKS builds a normal projection section from a three-segment jogged
  line and cuts at ONE offset. The bore sits at transverse 0.000 and the
  counterbores at -0.015; the section contains the counterbore-column
  features and no bore, so the cut being used is segment 3's.
- That explains r51 exactly: lengthening segment 1 cannot change a section
  that is not cut along segment 1. The line moved 40 mm and the view was
  byte-for-byte identical, and now there is a reason rather than a
  hypothesis.
- **Two new API facts recorded**: the returned array is 49 doubles where
  the documented layout wants 53 for three segments, and it **mixes
  frames** - segment endpoints in view space (`-0.102`), arrow and text
  points in sheet space (`0.107932223, 0.265060000`, inside Drawing View1's
  sheet outline). Both in `SOLIDWORKS_API_VALIDATION.md`.
- Not yet changed: the `Options` argument. It is a semantic change to how
  the section is cut and is the user's call.

## 2026-08-05 (31) - r52 decode what the drawing did, not what the path asked for

Read-only instrumentation, no behaviour change; deployed and run the same
day, live result in entry (32).

- `Module17_SectionPath.EmitSectionLineDecode` decodes
  `IView.GetSectionLineInfo2` for the view that carries the cut, called
  from `RecordSectionLineAfterLayout` after the layout has settled.
- **Why it did not exist before is the point.** Every run since the section
  first appeared read that array and logged only its element count. r50 and
  r51 both recorded `values=49`. r51 moved waypoint 1 by 40 mm, produced a
  visibly longer line on the sheet, satisfied the new full-crossing
  predicate, and yielded a byte-identical section view. Intent was verified
  on every run; the drawing's result never was.
- Documented layout, SOLIDWORKS 2025 Help: `[numSectionLines, layer,
  numSegments, per segment (lineType, startPt[3], endPt[3]), arrow and text
  tail]`. Three segments account for 53 doubles under that layout and the
  live array holds **49**, so the tail does not match on this build. The
  segment block sits at the front and is unaffected by the tail, so
  segments are decoded and the discrepancy is reported as
  `tailMatchesDocumented=` rather than hidden.
- **The raw array is dumped as well** - `SECTION_LINE_RAW|view=|from=|values=`,
  six doubles per line. 49 doubles is small, and a dump cannot be wrong
  about the thing a structured decode might be wrong about. The layout
  becomes decidable from the log instead of from documentation that already
  disagrees with the array length.
- `SECTION_LINE_SEGMENT|view=|index=|lineType=|start=|end=|frame=AsReturned`.
  The frame is **not** claimed: which coordinate system these points use has
  not been established, and asserting "Page" would be the exact
  mixed-frame defect this project has already paid for twice.
- Bounds: the walk refuses an array shorter than three, a segment count
  outside 1..63, and any segment whose seven doubles would read past the
  end.
- Static verification: **562/562**. Mutation-verified: removing the
  read-past-the-end guard, shifting the per-segment stride from 7 to 6,
  dropping the documentation comparison, removing the pipeline call, and
  moving the Err capture after `SafeViewName` each fail the suite.

## 2026-08-05 (30) - r51 live: the line moved, the section did not. Prediction wrong.

Run `macro_qa/20260805_043613_P-0251-14A-001`. Stage table **unchanged from
r50**: six required stages still fail. **The predicted result did not
happen and the reason is not yet known.**

- **The waypoint change worked exactly as designed.**
  `SECTION_PATH_BORE_OVERSHOOT|view=Drawing View1|centreY=0.210324890|projectedRadiusM=0.020000000|overshootM=0.040000000|w1Y=0.250324890|direction=AwayFromRows:PositiveY|reason=CutMustCrossWholeBore`,
  and the stronger predicate passed on the new geometry -
  `SECTION_CROSSING|proven=4|columnHoles=3|failures=None`. The section line
  is visibly longer on the sheet, now starting above the part.
- **The section view is byte-for-byte identical to r50.** Same
  `records=38|entities=38|doubles=1056|arcs=6|polylines=32|points=214`,
  the same three radii `0.002100;0.035000;0.036000`, the same seven X and
  seven Y coordinates, and the same view outline
  `0.191752,0.061120,0.221632,0.269000`. Not similar - identical.
- So `INNER_BORE_D40` and `FIT_BORE_D47_H7` still report `found=False`, and
  `SECTION_DIMENSIONS` is unchanged at 7 missing.
- **What this rules out.** The failure is not the crossing predicate and
  not the waypoint, both of which now do what they were built to do. A
  drawn section line's extent evidently does not determine what the cut
  contains - which is consistent with a SOLIDWORKS section cutting through
  the whole part regardless of the line's drawn length, but that is a
  hypothesis, not a measurement.
- **What is missing to answer it.** `IView.GetSectionLineInfo2` is read on
  every run and only its element count is logged (`values=49`, both runs).
  The coordinates it returns have never been decoded, so nothing in
  evidence says where SOLIDWORKS actually placed the cut relative to the
  four waypoints handed to it. That is the next measurement.
- The r51 reporting fixes did work as intended: `found=False` rows now read
  `kind=NoneOfRadiusYPairXPair|a=0.000000|b=0.000000` instead of carrying
  the previous requirement's coordinates.

## 2026-08-05 (29) - r51 the cut now passes through the whole bore, and has to prove it

Source and static verification when written; deployed and run the same day,
live result in entry (30). Two defects, both exposed by the r50 geometry
inventory (`macro_qa/20260805_041027_P-0251-14A-001`).

- **Waypoint 1 sat ON the bore centre**, so the cut removed half the bore
  and the section contained no bore opening to dimension - measured, not
  inferred: the section's only arc radii were 0.002100, 0.035 and 0.036,
  with no 0.020 or 0.023500 anywhere and no 0.040 or 0.047 span on either
  axis. Waypoint 1 now sits one full radius beyond the far wall:
  `path.BoreOvershootM = 2# * path.BoreProjection.ProjectedRadiusM`. The
  overshoot is the bore's **own** size, the same principle as the R23-703
  crossing slack, and no view outline is read - a contract test asserts
  `GetOutline`, `outline`, `UsableTop` and `UsableBottom` appear nowhere in
  the resolver.
- **The direction is read from the geometry**, not assumed: away from the
  face-hole rows, whichever side of them the bore is on.
- **The crossing predicate could not tell the two cases apart, and that is
  the deeper defect.** `PathCrossesCircle` asks only whether a segment
  comes *within* the radius, which a segment starting at the centre
  satisfies trivially. It reported
  `crossingsProven=4|crossingFailures=None` for the very cut that produced
  a section with no bore in it. A predicate that passes the case it exists
  to reject is worse than no predicate. The bore is now held to
  `PathFullyCrossesCircle`: both intersections of the line with the circle
  must lie inside the same segment, tested as
  `footDistance -/+ halfChord` fitting within the segment length.
  Failure reason `BoreNotCrossed` becomes `BoreNotFullyCrossed`.
- **Column holes deliberately keep the weaker test.** The first and last
  hole on the chosen column sit AT the segment's endpoints, so requiring a
  full crossing there would refuse the correct path. They need to be on the
  cut; only the bore needs its whole opening shown.
- Arithmetic check on the real r50 coordinates, bore centre
  `(0.137812223, 0.210324890)`, `R=0.020`, segment to
  `(0.137812223, 0.140324890)`: the old waypoint fails the new predicate
  and the new waypoint passes it. The fix and the guard agree about the
  case that was wrong.
- **The two r50 reporting defects are fixed.** `SECTION_REQ_CANDIDATE`
  resets `a=`/`b=` per requirement instead of printing the previous
  requirement's coordinates on a `found=False` row, and a diameter
  requirement is now searched as a radius **and then** as a linear span on
  both axes - a bore that has been cut open is the gap between two walls,
  not an arc.
- Static verification: **552/552**. Mutation-verified: replacing the
  radius-derived overshoot with a literal, reverting either direction
  branch, using the weaker predicate for the bore, and loosening the
  span arithmetic each fail the suite.

## 2026-08-05 (28) - r50 live: LAYOUT proved, and the section does not contain the bore

Run `macro_qa/20260805_041027_P-0251-14A-001`. Failed required stages
**7 to 6**. `LAYOUT` FAILED to PROVED.

- **The scale fix is confirmed live, exactly as the Help predicted.**
  `VIEW_SCALE_READBACK|view=Section View J-J|type=2|isIsometric=False|useSheetScale=0|scaleDecimal=1.000000`
  against `Drawing View1 ... useSheetScale=1|scaleDecimal=1.000000`. The
  section reports the flag as 0 while drawn at exactly the sheet ratio, so
  the old check was refusing a correctly scaled view. The isometric view
  reads `useSheetScale=0|scaleDecimal=0.500000` and remains exempt.
- **The polyline decode proved itself.**
  `SECTION_GEOM_SUMMARY|decodeStatus=Complete|records=38|entities=38|recordsMatchEntities=True|doubles=1056|arcs=6|polylines=32|points=214|error=0`.
  The documented record layout is now live-confirmed on SW2025.
- **The coordinate frame is answered.**
  `polylineBox=-0.009000,-0.098000,0.009000,0.098000` against
  `sheetOutline=0.289060,0.044385,0.318940,0.252265`. `GetPolylines7`
  returns **view-space** coordinates centred on the view origin, not sheet
  coordinates. The box is 18 mm by 196 mm - the part's thickness and
  height.
- **Five of the seven section requirements cannot be satisfied from this
  view, because the geometry is not in it.** Measured distinct
  coordinates: `X=0.008;0.002;-0.008;0.009;0.004;0.003;-0.009`,
  `Y=0.019;-0.097;0.062;0.098;0.018;0.017;-0.098`,
  `R=0.002100;0.035000;0.036000`.

  | requirement | nominal | present |
  | --- | --- | --- |
  | `OVERALL_THICKNESS_18` | 0.018 | yes, X pair -0.009/0.009 |
  | `BORE_STEP_DEPTH_12` | 0.012 | yes, X pair -0.008/0.004 |
  | `LOWER_WALL_STEP_11_5` | 0.0115 | no |
  | `INNER_BORE_D40` | 0.040 | no |
  | `FIT_BORE_D47_H7` | 0.047 | no |
  | `LONG_VERTICAL_REF_173_6` | 0.1736 | no |
  | `LOWER_VERTICAL_REF_104_8` | 0.1048 | no |

- **The bore is not in the section view at all.** Only three arc radii
  exist: 0.0021 (the M5 tap drill), 0.035 and 0.036 (the plate's rounded
  top profile). Neither 0.020 nor 0.0235 appears, and no linear span of
  0.040 or 0.047 exists in either axis. The bore's walls would sit at
  Y = 0.062 +/- 0.0235; the only Y near it is 0.062 itself, the bore
  centre, which is where the section path's first waypoint starts. **The
  cut begins at the bore centre instead of passing through the whole
  bore**, so the section shows no bore opening to dimension. This is a
  section-path waypoint question, not a dimension-engine question.
- Two reporting defects in this iteration's own new code, found by reading
  its own output: `SECTION_REQ_CANDIDATE` prints stale `a=`/`b=` values on
  a `found=False` row because they are not reset per requirement, and the
  two diameter requirements are searched only as arc radii - in a section a
  bore diameter is a linear span between the two cut walls, so the pair
  searches should apply to them as well. Neither changes the finding above
  (no 0.040 or 0.047 span exists either way), and both are fixed next
  iteration.

## 2026-08-05 (27) - r50 evidence pass: what does the section view actually offer

Source and static verification when written; deployed and run the same day,
live result in entry (28). Read-only instrumentation plus one behaviour
change, the LAYOUT scale check.

- **`Module10_SectionDimensionEngine.InventorySectionGeometry`** (new,
  read-only) decodes `IView.GetPolylines7` for the section view and reports
  what geometry exists: arcs with their radii, axis-parallel segments with
  their distinct X and Y coordinates, the coordinate box, and the view's
  scale flags. Section View J-J has existed only since r49; nothing has ever
  measured it, so `SECTION_DIMENSIONS requirements:7/satisfied:0` currently
  says nothing about whether the geometry to satisfy those seven is even
  present.
- **The decode is self-checking.** The record layout is the documented one
  (SOLIDWORKS 2025 Help, `IView::GetPolylines7`):
  `[Type, GeomDataSize, GeomData[], LineColor, LineStyle, LineFont,
  LineWeight, LayerID, LayerOverride, NumPolyPoints, [x,y,z]...]`, with
  Type 1 carrying `[cx,cy,cz, sx,sy,sz, ex,ey,ez, nx,ny,nz]`. Four
  range guards stop the walk and report `Desynchronized:<field>` rather
  than emitting invented coordinates, and `recordsMatchEntities=` compares
  the decoded record count against the returned entity array, which the
  Help states is positionally paired. That comparison is the control - the
  r40/r41 visibility classifiers shipped without one and measured nothing.
- **`SECTION_REQ_CANDIDATE|key=|nominalM=|kind=|found=|a=|b=`** answers, per
  requirement, whether the geometry that would measure it exists in the
  view: a curve of half the nominal radius for a diameter requirement, a
  pair of parallel curves the nominal apart for a linear one. Match window
  is 0.01 mm against nominals at least 5.5 mm apart.
- **The coordinate frame is reported, not assumed.** `SECTION_GEOM_FRAME`
  prints the polyline bounding box beside `IView.GetOutline` in sheet
  space, so which frame a section view's polylines live in becomes
  decidable from the log. It has never been established.
- **`VIEW_SCALE_READBACK|view=|type=|isIsometric=|useSheetScale=|scaleDecimal=`**
  is now emitted for every view in the layout validation loop.
- **The LAYOUT scale check was reading a flag as if it were a ratio, and is
  fixed** (user instruction, after the MCP evidence below). SOLIDWORKS 2025
  Help, `IView::UseSheetScale`: *"If the property is 0, then it is possible
  that the view scale is the same as the sheet scale"* - and
  `IView::UseParentScale` is the separate member a section view uses, so a
  section reads 0 while being drawn at exactly the sheet ratio. That single
  line failed the whole LAYOUT stage in r49.
  `Module9_LayoutEngine.ValidateLayout` now accepts `UseSheetScale = 1` as
  before and, only when that does not settle it, compares
  `IView.ScaleDecimal` against the proved sheet ratio from
  `ISheet.GetProperties2`. Failure text carries the numbers:
  `useSheetScale=/viewScale=/sheetScale=/sheetScaleProven=`.
- **The widened check fails closed.** `ViewScaleMatchesSheet` returns False
  when the sheet scale was never proved, when the denominator is zero, or
  on any read error - accepting a view whose scale nobody measured would be
  the opposite of what the stage is for. A view genuinely drawn at another
  scale still fails. The detail-view 3:1 rule and the isometric exemption
  are untouched.
- A defect in this iteration's own code was caught by the existing r45
  Err-capture contract: the new failure handler read `Err.Number` after
  calling `SafeViewName`, whose `On Error Resume Next` resets it. Captured
  first, as the guard requires.
- Static verification: **542/542**. The load-bearing new tests were
  mutation-verified: dropping the entity-count control, removing any one of
  the four decode guards, adding a mutating call to the inventory, removing
  the pipeline call, reverting the scale check to the flag, and removing
  either fail-closed guard from the ratio comparison each fail the suite.

## 2026-08-05 (26) - r49 the section path takes a position-proved bore; the anchor gate is untouched

Deployed and run at the user's instruction. Run
`macro_qa/20260805_034637_P-0251-14A-001`. User decision on Root 1, taken
after the r48 evidence in entry (25): "let the section path accept
position-proved projections, keep anchor gate for dimensions".

**The section exists. Root 1 is closed.**
`SECTION_PATH|view=Drawing View1|label=J|resolved=True|reason=None|segments=3|boreBasis=PositionProved:ProjectionAnchorUnavailable|distinctColumns=2|distinctRows=3|columnHoles=3|crossingsProven=4|crossingFailures=None`
then
`SECTION_CREATED|...|sectionView=Section View J-J|segments=3|selectionsVerified=3|sectionLine=Read|values=49`.
The prediction in entry (25) held exactly: the bore qualified on the weaker
proof in Drawing View1 and was still refused in Drawing View2.

Failed required stages **9 to 7**. `SECTION_GEOMETRY`,
`SECTION_CLEARANCE` and `ANNOTATION_EXTENTS` went FAILED to PROVED.
`VIEW_PROJECTION` improved but still fails: projections 22 to 33, accepted
6 to 8, locations without 5 to 3 - the two near-side M5 holes now project in
the section view, which is the only place they can. `ORDINATE_SCHEME`
regressed FAILED from PROVED, for a new reason belonging to the new view.
Remaining: `LAYOUT`, `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`,
`ORDINATE_SCHEME`, `SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`,
`FINAL_QA`.

- **The conflated requirement is now two requirements.**
  `CViewHoleProjection.PositionFailureReason()` /`HasProvedPosition()` prove
  the physical identity, the referenced configuration, the page coordinate
  frame and that the axis is normal to the view - everything
  `QualificationFailureReason` proves **except** the selectable anchor.
  `QualificationFailureReason` itself is byte-for-byte unchanged, and
  `Module13_ProjectionResolution` still decides `Accepted` through it alone.
- **`Module17_SectionPath.ResolveBoreProjection` accepts a candidate that is
  `Accepted` OR position-proved.** The path reads `PageX`, `PageY` and
  `ProjectedRadiusM` and never selects the bore - `Module17` contains no
  `PrimaryAnchor` and no `SelectEntity` reference, and a test holds that
  true - so a proved page position is the requirement that matches the use.
- **Largest radius still beats acceptance state.** "The principal bore" is a
  fact about the part; preferring a smaller accepted bore would cut the
  section through the wrong feature.
- Singleton-family selection, the face-hole grid proof and the crossing
  proof are all unchanged. **Face holes still require `Accepted`** - all six
  counterbores qualify in Drawing View1, so there is no evidence that they
  need loosening and they were not loosened.
- The weaker proof is **confined**: a test asserts no component other than
  `CViewHoleProjection` and `Module17_SectionPath` mentions
  `HasProvedPosition` or `PositionFailureReason`, so dimensioning, callout
  attachment and ordinate anchoring cannot reach it.
- New evidence: `SECTION_PATH_BORE_BASIS|view=|physical=|basis=|anchor=|frame=|axisNormal=|use=WaypointsOnly`,
  emitted only when the weaker proof was actually used, plus
  `boreBasis=` in every `CSectionPath.Summary()`. A position-proved section
  never passes silently as an accepted one.
- `NoAcceptedSingletonBoreProjection` renamed to
  `NoUsableSingletonBoreProjection`; the old name asserted a requirement
  that no longer holds.
- Confirmed live in Drawing View1: bore `axisNormal=True|frame=Page`,
  `pageX=0.137812223|pageY=0.210324890|projectedRadiusM=0.020000000`,
  rejected only for `ProjectionAnchorUnavailable`, six counterbores giving 2
  columns and 3 rows. Drawing View2 refused the same bore on
  `axisNormal=False`, as intended.
- **What the live run newly exposed, none of it caused by this change:**
  - `LAYOUT` fails on one line only:
    *"Non-isometric view is not using the proved sheet scale: 'Section View
    J-J'."* `Module9_LayoutEngine.bas:718` requires
    `IView.UseSheetScale = 1`, and a SOLIDWORKS section view inherits its
    PARENT view's scale, which is a different flag value even when the
    rendered ratio is identical. No per-view scale is logged, so the actual
    `UseSheetScale`/`ScaleDecimal` of that view is not yet in evidence.
  - `ORDINATE_SCHEME` regressed because the new view brought two new
    schemes. The horizontal one resolved its datum and created nothing
    (`apiResultName=swCreateOrdDimErr_Success|createdReadBack=0`); the
    vertical one was refused with
    `NoBucketAvailable;outline=NoMappedBottomEdge(edges:126,curve:57,notHorizontal:42,span:8,map:19)`.
    A section view's outline has no single mapped bottom edge.
  - `SECTION_DIMENSIONS` is now reachable for the first time and reports
    `requirements:7/satisfied:0/missing:7`. The seven named requirements
    (`OVERALL_THICKNESS_18`, `BORE_STEP_DEPTH_12`, `LOWER_WALL_STEP_11_5`,
    `INNER_BORE_D40`, `FIT_BORE_D47_H7`, `LONG_VERTICAL_REF_173_6`,
    `LOWER_VERTICAL_REF_104_8`) have never had a view to live in until now.
  - `NATIVE_CALLOUT_COVERAGE` and `MANUFACTURING_DEFINITION` both fail on
    the same single family, `op:EXTRUDEDCUT`, `missing=Attachment` - the
    stepped bore, which has no anchor in any view including the section
    (its axis lies in the section's page plane). The M5 and counterbore
    callouts were both created.
- Static verification: **526/526**. The four load-bearing new tests were
  mutation-verified: pointing `PositionFailureReason` at the anchor,
  reverting the bore resolver, loosening the face holes, and naming
  `HasProvedPosition` in `Module15_OrdinateScheme` each fail the suite.

## 2026-08-05 (25) - r48 live: Root 1 answered, the stepped bore is obscured in every orthographic view

Deployed at the user's instruction and run. Run
`macro_qa/20260805_033146_P-0251-14A-001`. Nine required stages still fail:
`LAYOUT`, `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`, `SECTION_GEOMETRY`,
`SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`, `ANNOTATION_EXTENTS`,
`FINAL_QA`, `SECTION_CLEARANCE`.

- **The selection clear worked and Route D ran for real.**
  `R23_PROJECTION_SELECTION_PRECONDITION|preexisting=1|cleared=True|remaining=0`.
- **The stepped bore's anchor line moved decisively:**

  ```text
  before r48:  mappedEdges=0  firstUnmappedRoute=A:Nothing:err0;B:Nothing:err0;D:RefusedPreexistingSelection
  r48:         mappedEdges=4  inventoryConfirmed=0  firstReject=MappedEntityNotInVisibleInventory
  ```

  All four circular edges map. `IView.SelectEntity` accepts them and
  `ISelectionMgr.GetSelectedObjectsDrawingView2` proves the owning view, so
  they exist in Drawing View1. None is in `IView.GetVisibleEntities2`, which
  the 2025 Help defines as entities "not completely obscured by other
  entities in the view". **The bore's circular edges are completely obscured
  in the Front view** - the first evidence-backed answer to the question open
  since the post-1845 review. The r40/r41 "obscured" counts were void; this
  is a different instrument with a working control. The fail-closed guard is
  behaving correctly in refusing an obscured edge as a circular anchor.
- **Root 1 is therefore a design conflict, not a defect.**
  `Module17_SectionPath.ResolveBoreProjection` requires an **accepted**
  projection, but the section path consumes only `BoreProjection.PageX`,
  `PageY` and `ProjectedRadiusM` - a proved page **position**, not a
  selectable anchor. `Module13_ProjectionResolution` already records the page
  centre for an unanchored location ("An unanchored location still has a
  provable position"). A bore hidden in every orthographic view can never
  satisfy the stricter requirement, which is exactly why the reference
  drawing puts that feature in the J-J section - and that section cannot be
  created because it demands the projection the hidden bore cannot give.
  Circular. **Referred to the user as a decision, not patched.**
- `VIEW_PROJECTION` reads
  `locations=11/projections=22/acceptedProjections=6/locationsWithProjection=6/locationsWithout=5`;
  the section path still reports `NoAcceptedSingletonBoreProjection` for both
  views.
- Noted, not fixed: the six counterbores now make two redundant Route D calls
  each, and `firstReject=MappedEntityNotInVisibleInventory` is recorded on
  locations that were nonetheless accepted, which reads as a failure when it
  is not.

## 2026-08-05 (24) - r48 stale selection was blocking Route D; every failure now traces to Root 1

Source and static verification when written; deployed and run later the same
day, live result in entry (25). Run
`macro_qa/20260805_032817_P-0251-14A-001` is the r47 evidence behind it.

- **Route D ran for the first time and refused every location.**
  `firstUnmappedRoute=A:Nothing:err0;B:Nothing:err0;D:RefusedPreexistingSelection:count1`
  on all eleven. View creation leaves one object selected, and Route D
  refuses to select over an existing selection because `IView.SelectEntity`
  with `AppendFlag=False` would destroy an interactive caller's state. One
  stale selection blocked the whole sheet. The r47 gate fix was correct; it
  was simply never reached.
- `BuildAllViewProjections` now clears a stale selection before the mapping
  loop and records
  `R23_PROJECTION_SELECTION_PRECONDITION|preexisting=|cleared=|remaining=`.
  Same correction as the ordinate stage at r44, same rationale: the
  production route owns the document it just created, and the Module13 guard
  stays intact for the read-only probes. One clear suffices because a
  successful Route D clears its own temporary selection.
- **The completeness reason settled the callout question and corrected an
  earlier claim of mine.** Both families report `missing=Attachment` and
  nothing else. The r46 changelog recorded `op:EXTRUDEDCUT` reporting
  `dia:0.000000000` as a second, separate defect needing a geometry-derived
  diameter; the completeness check does not flag `NominalDiameter` for it, so
  that zero is in the rendered family key only and the definition itself is
  complete. **No separate extruded-cut diameter defect exists.**
- Consequence: **every one of the nine remaining stage failures now traces to
  Root 1.** Attachment needs a proved projection; the section needs the bore
  projection; `LAYOUT` needs the section; the M5 side holes have axis
  `(0,1,0)` and can only project in the J-J section view, which is why the
  reference drawing puts them there.
- Static verification: **510/510**.

## 2026-08-05 (23) - r47 Route D gate re-keyed (Root 1); callout completeness made legible

**Not deployed.** Source and static verification only, at the user's
instruction. No r47 live evidence exists.

- **Callout completeness reason is now emitted.**
  `CCalloutDefinition.CompletenessFailureReason()` already named every
  missing field, but `CreateMissingR23Callouts` consulted only
  `IsComplete()`, so the r46 report said the M5 family was incomplete
  without saying which field was missing. The reason now reaches both the
  failure text and a new `R23_CALLOUT_INCOMPLETE|family=|missing=|` line.
- **Root 1: the Route D gate was keyed on the wrong condition.**
  `ResolveProjection` ran Route D only when `Not visibleInventoryAvailable`.
  Drawing View1 HAS an inventory (39 edges), so Route D was never attempted
  there - even though Route A demonstrably declines for real geometry in that
  same view: six counterbores map two edges each while the stepped bore maps
  none of its four, `A:Nothing:err0` on every one. With no second route that
  bore has no accepted projection, `ResolveBoreProjection` finds no singleton
  bore, and no section path resolves. The gate is now "Route A and B both
  declined".
- This is not a weakening. When an inventory exists, a Route-D entity must
  still clear the Route C membership check immediately below, so it is proved
  **twice** - ownership through
  `ISelectionMgr.GetSelectedObjectsDrawingView2` **and** membership of
  `IView.GetVisibleEntities2` - where the inventory-less path proves only
  ownership. A genuinely obscured edge cannot be in the visible inventory and
  is still rejected as `MappedEntityNotInVisibleInventory`. No coordinate
  search was introduced, and a contract test asserts none appears.
- **The r40-r42 visibility classifiers were removed, not kept.** They were an
  instrument for one question - what `GetPolylines7` returns - and r42
  settled it. The model-space variant matched nothing and so measured
  nothing; leaving `unmappedObscuredEdges` in evidence would keep asserting a
  fact never established. `unmappedAllRoutes` replaces it and says only what
  is true: every route declined. The drawing-space control also cost a
  `GetPolylines7` read per mapped edge. The measurement itself is preserved
  in `SOLIDWORKS_API_VALIDATION.md`, and a contract test now asserts it is
  recorded there rather than in code.
- **Still open, unchanged:** `op:EXTRUDEDCUT` reports `dia:0.000000000`. That
  is a different reader - an extruded cut has no feature-data diameter and
  its size must come from geometry. Not attempted at r47.
- Static verification: **507/507**.

## 2026-08-05 (22) - r45/r46 Hole Wizard nominal comes from the type-specific member

- **r45 probe, run `macro_qa/20260805_001154`.** A read-only dump of 22
  candidate `IWizardHoleFeatureData2` members settled Root 2 without a guess:

  | Feature | `Type` | Member holding the value |
  |---|---|---|
  | CBORE for M6 | 14 `swCounterBoreThru` | `ThruHoleDiameter=0.0066`, `ThruHoleDepth=0.018` |
  | M5x0.8 Tapped | 46 `swTapBlindCosmeticThread` | `ThreadDiameter=0.005`, `ThreadDepth=0.010` |

  `accessGranted=True` on both and `CounterBore*` read correctly, so
  `AccessSelections` was never the problem. `HoleDiameter` and `Diameter`
  simply do not apply to a standard-driven Hole Wizard feature. 6.60 THRU and
  11.00 by 6.00 match the reference callout exactly.
- **r46** replaces the direct `HoleDiameter`/`HoleDepth` reads with
  `FirstNonZeroHoleMember`, which walks a candidate chain and names the
  supplying member in the proof string. The two observed types are disjoint
  on the first two entries, so the chain is deterministic rather than a
  priority guess. An all-zero read now reports `ReadAllZeroValues:<routes>`
  instead of `Read`.
- **Confirmed live, run `macro_qa/20260805_001521`:**
  `diameterM=0.006600000|depthM=0.018000000` and
  `diameterM=0.005000000|depthM=0.010000000`.
- **`NATIVE_CALLOUT_COVERAGE` and `MANUFACTURING_DEFINITION` did not flip.**
  Failure count stays 9. Two distinct causes remain and neither is the Hole
  Wizard read:
  1. `op:EXTRUDEDCUT` still reports `dia:0.000000000`. That is a different
     reader; an extruded cut's diameter comes from geometry, not feature data.
  2. The M5 family now carries good values but is still judged `Incomplete`
     **and** `Attachment`. Attachment needs a proved projection, and the M5
     side holes have axis `(0,1,0)` - normal to neither the Front nor the
     Left view - which is Root 1, the same reason the reference drawing puts
     them in the J-J section.
- **Evidence gap:** production emits the completeness token but not
  `CCalloutDefinition.CompletenessFailureReason()`, so which field the M5
  family is still missing is not visible in the QA report.
- Two of my own defects were caught by existing gates before touching a
  drawing, and both now have offline checks: a non-ASCII byte written into
  cp1252 managed source (`VERIFY: FAIL`, nothing promoted), and a module-level
  `Const` placed between procedures (`ready=False|reason=CompileNotClean`,
  `main` never invoked).
- Static verification: **503/503**.

## 2026-08-05 (21) - r44 ordinates created; stale selection was blocking the datum

Run `macro_qa/20260805_000138_P-0251-14A-001`.

- **Ordinates exist for the first time.** `Horizontal ordinate groups: 1`,
  `Vertical ordinate groups: 1`, `Ordinate feature selections: 12`,
  `display dimension count=8, ordinate dimension count=4`. Every one of those
  was 0 in every prior run.
- Cause of the last block: the datum proof refuses to select while anything
  is already selected, because `IView.SelectEntity` with `AppendFlag=False`
  would destroy an interactive caller's selection. Annotation import runs
  immediately before the ordinate stage and leaves its last inserted
  annotation selected. r43 evidence showed the same datum, same anchor
  (`SolidBodyEdge_0_49`, `visibleEntityIndex:23`), refused at creation
  (`selection=Reject|reason=PreexistingSelection|initialSelectionCount=1`)
  and resolved in the QA pass where nothing was selected.
- `CreateR23OrdinateGroups` now clears a stale selection before building
  schemes and records
  `R23_ORDINATE_SELECTION_PRECONDITION|preexisting=1|cleared=True|remaining=0`.
  The production pipeline owns the document it just created. The Module13
  guard is untouched and still fails closed for the read-only probes.
- **Open discrepancy, not yet explained:** the vertical group reported
  `selectionsAppended=4|expectedSelections=3` and was still accepted as
  `swCreateOrdDimErr_Success`. The count mismatch is not currently gated.
- Cumulative against the 18:45 r37 baseline: failed stages **11 -> 9**;
  `DIMENSION_ARRANGE` and `ORDINATE_SCHEME` both FAILED -> PROVED.
- Static verification: **491/491**.

## 2026-08-04 (20) - r42/r43 GetPolylines7 holds drawing entities; datum comparison corrected

- **r42 two-way control settled it.** Same comparison, same array, twelve
  edges `GetVisibleEntities2` had confirmed present:
  `mappedVisibleEdges=0` (model edge) against
  `mappedVisibleDrawingSpace=2` per location (mapped drawing entity). Twelve
  of twelve in drawing space, zero of twelve in model space. Recorded in
  `SOLIDWORKS_API_VALIDATION.md`.
- **r43 corrects `MapVisibleDatumEntity`.** It tested the model entity from
  r33 onward, so `visibleIndex` was always `-1` and every vertical datum
  failed closed as `PolylineVisibilityUnavailable` regardless of real
  visibility. Order is now map first, then prove the mapped drawing entity is
  in the visible array. This also explains the scratch-versus-production
  divergence: scratch returns `status=NoEdges` and takes the documented
  fallback; production returns 39 entries that never matched.
- The r37 fail-closed guard is retained unchanged. A mapped entity absent
  from a non-empty array still refuses (`DatumMap:MappedEntityNotVisible`);
  only the documented empty-array case reaches scoped selection.
- **Not established:** whether the P-0251 stepped bore's edges are obscured.
  Drawing-space testing needs a drawing entity, and no route maps that bore,
  so it stays unclassified. The r40/r41 `unmappedObscuredEdges` counts remain
  void.
- Static verification: **486/486**. r43 live evidence pending.

## 2026-08-04 (19) - r41 SelectData.View binding; visibility control added

- **`ISelectData.View` raises 91 in this drawing context.** Modules 13, 15
  and 16 each wrap that assignment in a private `TryBindSelectDataView`;
  `Module4_ModelItemImporter` assigned it raw, so **every** dimension-arrange
  attempt in the 18:45, 23:24 and 23:43 runs died in its error handler. Module4
  now uses the same wrapper and records
  `DIMENSION_ARRANGE_BINDING|viewBinding=...`. Binding scopes the selection;
  it is not a precondition, and the view is already activated and read back,
  so an unbound view continues.
- **The r40 visibility classification shipped without a positive control**,
  which made its own output unreadable: if `FindVisibleModelEdgeIndex` can
  never match, every edge reports obscured and the count means nothing. The
  r40 run produced exactly that shape. `mappedVisibleEdges` is the control -
  an edge Route A already mapped is known to be in the view. **No conclusion
  about hidden geometry may be drawn until that control is read.**
- Static verification: **480/480**.
- **r41 live run `macro_qa/20260804_234951_P-0251-14A-001`.**
  `DIMENSION_ARRANGE` flipped **FAILED -> PROVED** (`api=1, noAction=2`);
  `AlignDimensions` returned True on Drawing View2 for the first time.
  `viewBinding=UnboundAfterError:91` is still recorded on both views, proving
  the binding genuinely fails and that it was never a precondition.
- **The visibility control read zero, and it invalidates the r40 reading.**
  Twelve counterbore edges were mapped by Route A *and* confirmed present in
  the `GetVisibleEntities2` inventory (`inventoryConfirmed=2` each), yet
  `mappedVisibleEdges=0` on every one. `FindVisibleModelEdgeIndex` therefore
  matches nothing at all, so every `unmappedObscuredEdges` count - including
  the bore's 4 - measures nothing. The hidden-geometry hypothesis is
  **untested**, not confirmed.
- Same broken comparison gates the vertical datum:
  `mapSample=VisibleMapUnavailable:DatumMap:PolylineVisibilityUnavailable`.
  It also explains why the scratch drawing resolved its datum and production
  does not - the scratch view returned `status=NoEdges` and took the
  selection fallback, while production returns 39 edges that never match.
  One defect, two symptoms.
- Note for the next investigation: `GetPolylines7` returns **39** edges for
  Drawing View1, the same count `GetVisibleEntities2` returns. The comparison
  logic is shared and demonstrably works in drawing space
  (`visibleIndex=33..38`); only the model-space comparison fails. Test
  whether `GetPolylines7`'s array holds drawing-context entities rather than
  model edges before changing any route.

## 2026-08-04 (18) - r40 Phase 2 diagnosis; three blind failures made legible

Evidence-only revision; no acceptance behaviour changed. Run
`macro_qa/20260804_234310_P-0251-14A-001`.

- Section failure now names its cause. Every candidate view reports
  `R23_SECTION_PATH_CANDIDATE|` and the reasons reach the failure text:
  `R23 semantic section path was not resolved: Drawing
  View1:NoAcceptedSingletonBoreProjection;Drawing
  View2:NoAcceptedSingletonBoreProjection (viewsWithoutProjections=0)`. This
  confirms the review's F3 diagnosis from production evidence.
  `CreateSemanticPrimarySection` also stopped overwriting `path` each
  iteration, which had been discarding every view's reason but the last.
- **Five** error handlers were destroying the error they reported, not one.
  Each built its message by calling a name helper containing
  `On Error Resume Next`, which resets the global `Err`; VBA evaluates the
  concatenation left to right. Fixed in Modules 4, 5, 10, 13 and 14, with a
  regex contract test guarding the shape. This immediately turned
  `Dimension arrange API error in 'Drawing View1': 0: ` into
  `91: Object variable or With block variable not set`.
- Read-only visibility classification added to `ResolveProjection`
  (`unmappedVisibleEdges` / `unmappedObscuredEdges`), and the post-layout
  section-line filter corrected: an uncut view answers
  `sectionLine=Read|values=0`, not `NoGeometryReturned` as r38 assumed.
- Static verification: **475/475**.

## 2026-08-04 (17) - r39 automated mutating production run; Phase 1 confirmed live

- User authorized replacing the manual **Debug > Compile Project** gate and the
  per-run chat authorization with automation. New
  `tools/production-runner/Run-R23Production.ps1` deploys, runs the new
  read-only `Module20_ProbeRunner.R23_PrepareProductionRun` (programmatic VBE
  compile plus part activation), and **refuses to invoke `Module1_Main.main`
  unless the pre-flight logged `R23_PREFLIGHT_END|ready=True`**. That refusal
  is what the manual gate was actually for: VBA compiles lazily, so a module
  that only fails when first called would otherwise abort a run after several
  views already exist.
- What did not move: `-AllowMutation` is mandatory with no default, only the
  three authorized fixtures are accepted, the part is opened read-only,
  nothing is ever saved, `UserForm1` is still driven by the operator, and
  visual/manufacturing acceptance remains the user's judgement. Recorded as
  the "Automated mutating-run exception" in `Agents.md`.
- **First r39 production run, `macro_qa/20260804_232440_P-0251-14A-001`.**
  Both r38 fixes confirmed live against the 18:45 baseline:
  `Annotations imported` 0 -> **5**, `Layout moves` 0 -> **2**, with
  `ACTIVATE_VIEW|operation=Annotation import into 'Drawing View2'|setterResult=False|readbackMatched=True`
  showing the exact false negative that used to abort the import.
- `LAYOUT` still fails, but now for a true reason rather than a sequencing
  bug: the final structural pass ran, dispatched to the P-0251 reference
  zones, and found no J-J section because the section genuinely was not
  created. It cannot pass until the section defect is fixed.
- Still open and unchanged: section path, projection Route D suppression,
  all-zero Hole Wizard reads, bottom-outline datum, and the self-satisfying
  `MODEL_IMPORT_COVERAGE` gate, which again reported PROVED while
  `Drawing View1` imported nothing.
- Static verification: **464/464** (17 new production-runner contracts).

## 2026-08-04 (16) - r38 proved view activation and a two-pass layout

Both defects were diagnosed from the failed 18:45 production run
(`macro_qa/20260804_184514_P-0251-14A-001/QA_REPORT.txt`), not from static
reading, and each was the sole cause of several stage failures.

- `IDrawingDoc.ActivateView` returns False on this build even when the view
  does become active; the same run proves it
  (`ACTIVATE_VIEW|operation=ResolveOutlineDatum|setterResult=False|readbackMatched=True`).
  Annotation import branched on that raw result, logged
  `IMPORT_VIEW_NOT_ACTIVATED` for both views, never called
  `InsertModelAnnotations4`, and produced `Annotations imported: 0`.
  Annotation import, ordinate-group creation, native callout creation, and the
  section cut now activate through
  `Module8_RuntimeSupport.ActivateDrawingView`, which proves activation by
  active-view readback and records a disagreeing setter as a warning.
  `Module13_ProjectionResolution` keeps its raw read: it reports activation
  and does not gate on it.
- The only layout pass ran at pipeline step 3, before the section and
  isometric views existed, so `ArrangeP0251ReferenceZones` failed
  unconditionally on "requires exactly one primary, side, J-J section, and
  isometric view" and the run recorded `Layout moves: 0` - no view was ever
  positioned. `ArrangeViewsInMeasuredGrid` now takes `isFinalStructuralPass`.
  The step-3 rough pass pre-places views with the generic zone-aware path and
  defers the LAYOUT verdict; a new step-11b pass runs the fixture reference
  zones once every required view exists, and is the only pass that proves or
  fails the stage. `MarkStageFailed` is permanent, which is why the rough pass
  must not write a verdict.
- The section line belongs to its source view, so the structural pass is
  followed by a read-only `RecordSectionLineAfterLayout` re-read
  (`R23_SECTION_LINE_POSTLAYOUT`).
- The retired automatic content-envelope repositioning and rescaling stay
  uncalled. `FINAL_LAYOUT` remains
  `UserAcceptedAsIs/automaticClearance=DeferredByUser`.
- Static verification: **447/447** (12 new Phase 1 contracts). r38 deployment,
  readback, VBE compilation, and live production evidence are pending.

## 2026-08-04 (15) - r37 scoped-outline evidence

- `IView.GetPolylines7` returns no edge array on both P-0251 HLV ordinate
  views. A fallback now exists only for that named API state:
  `IView.SelectEntity` must return
  an entity owned by the intended drawing view through the selection manager.
- Both vertical ordinate datums are now proven from mapped straight lower
  outline edges, never from hole centres or view bounds. The focused P-0251
  runner proves `ORDINATE_SCHEME=True`; semantic QA is **9/10**, with only
  creation-dependent `MODEL_IMPORT_COVERAGE` remaining false.
- Static verification: **435/435**. Deployment/readback: **38/38**.
  Programmatic VBE compilation: clean. All nine read-only probes completed
  without drawing or model mutation. Evidence: `probe_runs/20260804_164014/`.

## 2026-08-04 (14) - r31 user-accepted layout policy

- User accepted the layout as-is. Production retains initial structural-grid
  placement and normal dimension arrangement, but cannot invoke automatic
  content-envelope view movement or rescaling.
- The retired scratch-layout command now reports
  `UserAcceptedLayoutAsIs|mutations=0`. `FINAL_LAYOUT` remains visible as a
  user waiver with `automaticClearance=DeferredByUser`, not a clearance pass.
- Static verification: **435/435**. r31 live deployment, VBE compilation,
  read-only runner evidence, and all production-mutation evidence are pending.

## 2026-08-04 (13) - r30 independent review corrections

- Separated planned ordinate coverage from created-and-read-back coverage and
  moved semantic import coverage before read-only scheme construction.
- Deduplicated production import views; stopped hiding a curve-order mismatch;
  guarded every reviewed selection path against a pre-existing selection;
  rejected zero nominal fallbacks; and made compile failures request VBE dialog
  evidence rather than falsely naming a module.
- Layout now rejects a malformed view/cell plan before rescaling and refuses
  to override approved isometric or detail scales. Outline datums now require
  a straight edge in the documented visible-entity inventory; without that
  inventory they fail closed.
- Static verification: **433/433**. r30 guarded deployment/readback passed
  **38/38**, programmatic VBE compilation was clean, and all nine focused
  P-0251 read-only probes completed. The newly strict visibility guard leaves
  both vertical datums unresolved; semantic QA correctly reports 7/10 rather
  than promoting selection ownership to visibility proof.

## 2026-08-04 (12) - r29 outline-derived vertical ordinate datum

- Replaced the vertical datum's lowest-hole fallback with the lowest mapped
  horizontal model edge in the target view. Route A remains preferred and
  Route D is the existing guarded part-drawing fallback; target-view ownership
  is read back before the datum is accepted.
- Runner evidence at `probe_runs/20260804_132049/` proves both P-0251 vertical
  datums as `OutlineDerived` at `pageY=0.137763153`, with
  `coverageFailures=None`, no creation or mutation, and semantic QA 9/10.
  Only the pre-existing automatic `FINAL_LAYOUT` clearance gate fails.
- `IView.GetOutline` is recorded as an enclosing view bound, not silhouette
  geometry: the selected model edges sit 5.94 mm above its lower bound.
  Managed-source hygiene and the companion suite pass at **428/428**;
  revision advances to r29.

## 2026-08-04 (11) - R23-006 curve-order proof

- Added a read-only Phase 0 comparison inside the existing feature-catalog
  probe. It reads one owned circular edge for counterbore, M5 tapped, mirror,
  and extruded-cut roles in both historical API orders.
- Fresh runner evidence at `probe_runs/20260804_125350/` passed all four:
  `IsCircle=True`, seven `CircleParams` entries, equal radius, zero closure,
  `R23_CURVE_ORDER_END|failures=None`, and `modelUnchanged=True`.
- Static suite grew to **426/426 passed**; guarded deployment/readback and
  programmatic VBE compile were clean. Production behavior is unchanged, so
  source revision remains r28.

## 2026-08-04 (10) - r28 source hygiene and fresh runner evidence

- Reflowed managed VBA source without behavior changes: all 39 `.bas`/`.cls`
  files now meet ASCII, CRLF, no-BOM, no-trailing-whitespace, `Option Explicit`,
  and 79-character contracts. `MACRO_SOURCE_REVISION` remains r28.
- Full companion suite: **425/425 passed**. Guarded deployment/readback passed
  **38/38** for both candidate and target.
- Fresh P-0251 scratch runner evidence is retained at
  `probe_runs/20260804_124325/`: VBE `verdict=Clean`, module touch
  `firstFailedModule=None`, and all nine read-only probes completed. Semantic
  QA remains 9/10 because the automatic final-layout gate still names five
  clearance failures; the prior user visual acceptance remains separate.

## 2026-08-04 (9) - r28 accepted scratch-layout evidence

- The controlled layout entrypoint ran only against
  `P-0251-14A-001-R23-layout-scratch.SLDDRW`. It applied two measured scale
  passes, reached scale `0.384200`, and correctly made no position move after
  the re-measured plan remained `LargerSheetRequired`.
- The user accepted the disposable scratch result. It was saved through the
  installed SW2025 `IModelDoc2.Save3` contract with `errors=0|warnings=0`;
  neither fixture model nor reference drawing was saved or changed.
- The r28 probe runner then compiled the loaded VBA project cleanly, touched
  every module, and completed all nine probes. Semantic QA remains 9/10:
  `FINAL_LAYOUT` conservatively reports five clearance failures. Human visual
  acceptance and the automatic production gate are recorded separately.

## 2026-08-04 (8) - r28 bounded measured-rescale retry

- The first scratch layout run applied factor `0.527974`, rebuilt, then
  truthfully rejected a still-required factor `0.727687`. Text extents do not
  shrink with the view, so one geometric estimate was insufficient.
- `Module18_ContentEnvelope.ApplyPlacementPlan` now permits two
  scale/readback/rebuild/re-measure passes and records each pass. A third
  request still returns `LargerSheetRequired`; no unbounded shrink loop exists.
- Static suite: **425/425**. r28 guarded deployment/readback passed (`38/38`).
  The later accepted scratch run and r28 runner evidence are recorded above.

## 2026-08-04 (7) - r27 isolated Phase 9 scratch-layout route

- Created the byte-identical disposable P-0251 drawing at
  `test_assets/scratch_drawings/P-0251-14A-001-R23-layout-scratch.SLDDRW`.
  The V: drawing and fixture part remain unchanged.
- Added `Module2_DrawingPipeline.R23_ApplyContentLayoutToScratch`. It accepts
  only that exact path, proves the referenced fixture through a drawing view,
  writes direct evidence, invokes only the shared final-layout helper, and
  never saves the scratch drawing.
- r27 guarded deployment/readback passed (`38/38`); programmatic VBE
  compilation returned `verdict=Clean`; all nine read-only scratch probes
  completed. Static suite: **425/425**.
- The entrypoint has not mutated the scratch drawing. Manual VBE compilation,
  the one authorized layout run, screenshot review, and visual acceptance
  remain open.

## 2026-08-04 (6) - r26 semantic coverage and clean runner regression

- Fixed the VBE compile break in `CLocationGraph.ClearImportedAnnotations`:
  the class now uses its own `Scripting.Dictionary` construction contract.
- Fixed semantic QA ordering: ordinate buckets record directional scheme
  coverage, and `MODEL_IMPORT_COVERAGE` runs after scheme reconstruction.
  This does not promote authored drawing dimensions into imported evidence.
- r26 guarded deployment/readback passed (`38/38`), VBE compilation returned
  `verdict=Clean`, module touch returned `firstFailedModule=None`, and all nine
  probes completed against authorized P-0251 with `mutations=0`.
- Semantic QA improved from `proved=8|failures=2` to
  `proved=9|failures=1`; only the real protected-reference layout failures
  remain. Static suite: **424/424**.

## 2026-08-04 (5) - runner title-conflict diagnostics

- The runner now decodes locally confirmed
  `swFileWithSameTitleAlreadyOpen = 65536`, enumerates the existing
  document path, and fails closed without calling `CloseDoc`.
- Live proof: the restarted SOLIDWORKS session had
  `V:\VEEMAP\SW_data\P-0251-14A-001.SLDPRT` open with its drawing, blocking
  the authorized local fixture part before compilation or probes began.
- Corrected stale Phase 9 review wording: its three high-severity source
  fixes exist; their mutating runtime evidence remains open.
- Static verification after the runner change: **424/424** companion tests.
- Fixed the VBE compile failure in `CLocationGraph.ClearImportedAnnotations`;
  `MACRO_SOURCE_REVISION` is now `target-spec-hybrid-v2-2026-08-04-r24`.

## 2026-08-04 (4) - R23 Phase 11 production wiring, static only

- Replaced the production feature-list route with `CLocationGraph` and the
  R23 feature, projection, import, ordinate, section, callout, envelope, and
  semantic-QA stages in the required order.
- Retired the reachable hardcoded P-0251 callout and count-only QA paths;
  static contracts now reject their return.
- Generic dimension arrangement now leaves semantic section lanes unchanged.
- `MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-04-r23`.
- Static verification passed: 422 companion tests and deployment preflight
  with 38 managed components. No VBE compilation, live run, screenshot review,
  or manufacturing acceptance is claimed.

## 2026-08-04 (3) - probe runner made the required test path; batch run assessed

### Documentation

CLAUDE.md, Agents.md and R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md now
require tools/probe-runner/Run-R23Probes.ps1 for read-only testing and for
compiling the VBA project. The agent reads the run's log file; it does not
ask for a pasted Immediate Window unless the evidence is something the
command cannot produce - a mutating run, a compile-error dialog, or a
screenshot. Agents.md also records that the programmatic compile route is
built and proved, and that a verdict=Clean is VBE compilation evidence
without being runtime or manufacturing acceptance.

Stale figures corrected: the offline suite is 401 tests, not 74.

### Phase 9 and Phase 10 status updated from the batch run

Phase 9's third live run confirmed the five defects from the second run are
fixed and produced plan=RescaleRequired with suggestedScaleFactor=0.527974,
requiredHeightM=0.479190 against usableHeightM=0.253000. Gate not satisfied:
ten clearance failures, R23-903 and R23-904 still unrun, and three
high-severity review findings unfixed in that same mutating path.

Phase 10's first live run: stages=10, proved=3, failures=7, mutations=0.

### GetVisibleEntities2 returns zero under the runner

Every view in the batch run reported count=0 from IView.GetVisibleEntities2
and was skipped as NoVisibleDrawingEntities, so Phases 3-7 read zero
downstream and six of Phase 10's seven failures follow from it. The same
probes returned nonzero standalone from the VBA editor.

Phase 8 ran identically in both (satisfied=7, missing=0) and does not read
visible entities, so the drawing was open, correct, and bound to the
authorized test_assets part. The fault is specific to visible-entity
enumeration under the runner's OpenDoc6/ActivateDoc3 sequence. Recorded at
the top of R23_IMPLEMENTATION_PLAN.md with the experiment that separates the
candidates.

## 2026-08-04 (2) - R23 probe-automation tool: first live run, PA-103/110/112

Completes the tool from the entry below. User opened a live SOLIDWORKS
session; ran the full plan through to a successful
`R23_RunAllProbes` execution. Evidence:
`test_assets/iteration_evidence/probe_runs/20260804_054014/probe_log.txt`.

- Compile control resolved live by caption:
  `id=578|caption=Compile Fable`, under Menu Bar > Debug. Verdict
  `Clean` after `.Execute` (`enabledBefore=True`, `enabledAfter=False`).
- All 21 standard modules touched clean
  (`R23_RUN_TOUCH|firstFailedModule=None`).
- All nine probes ran in dependency order with `status=Completed`,
  `mutations=0`/`creations=0`, `drawingUnchanged=True` throughout, and
  every `part=` field resolving to the authorized `test_assets\models\`
  copy (the 2026-08-02 network-sibling rebind failure did not recur).
- Fixed two live-only bugs found en route: `CommandBarControl.Caption`
  carries the raw `&` accelerator marker (`"Compi&le Fable"`), which was
  breaking the caption match until it compared against the cleaned
  string instead of the raw one; and PowerShell cannot call `OpenDoc6`
  directly (fails via late binding and via a direct interop
  bracket-cast), fixed with a new `SolidWorksDocumentOpener.cs` mirroring
  the existing `SolidWorksMacroInvoker.cs` pattern, exposed as
  `Open-SolidWorksDocument` alongside `Invoke-SolidWorksMacro` in
  `tools/swp-deploy/Invoke-SolidWorksMacro.ps1`.
- Also fixed: `R23_RunAllProbes` never switched the active document
  between probe 1 (needs the part) and probes 2-9 (need the drawing) -
  `OpenDoc6` documents it does not activate on open. Added explicit
  `ActivateDoc3` calls, matching the `GetTitle`-based pattern
  `Module8_RuntimeSupport.ActivateDrawingDocument` already used.
- 5 new regression tests lock in both live-only bugs and the refactor
  (`SolidWorksDocumentOpener.cs` hygiene, the `&`-stripped caption
  match). Full suite: 401 tests, same five known-stale R20 failures.
- Full narrative, live evidence, and API-contract detail:
  `docs/R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md` section 12 and
  `docs/SOLIDWORKS_API_VALIDATION.md` (2026-08-04 entry).

## 2026-08-04 - R23 probe-automation tool built (PA-100..109, static)

Implements the tool authorized 2026-08-02:
[R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md](R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md).
Everything offline-verifiable is done; PA-103 (live VBE enumeration),
PA-110 (recording what it proved) and PA-112 (first live run) still need
the user's SOLIDWORKS session and are not claimed here.

- `Module21_EvidenceSink.bas` (new): `OpenLog`/`LogLine`/`CloseLog`/
  `IsOpen`, teeing every diagnostic line to
  `test_assets/iteration_evidence/probe_runs/<timestamp>/probe_log.txt`
  in addition to `Debug.Print`, so an external process can read a probe
  run without a manual Immediate Window paste.
- `CRunEvidence.AddInfo`/`AddWarning`/`AddFailure` now route through
  `LogLine` instead of calling `Debug.Print` directly. The
  `QA INFO: `/`QA WARNING: `/`QA FAILURE: ` prefixes are unchanged.
- All 147 `Debug.Print` sites across the nine `R23_Probe*` modules
  (Module10, 12-19) replaced with `LogLine`, mechanically, preserving
  every log line's text.
- `Module20_ProbeRunner.bas` (new): `R23_EnumerateVbeControls`,
  `R23_CompileProject` (resolves the VBE Compile control by caption, not
  a hardcoded ID - that ID is outside the SOLIDWORKS API corpus, see
  `docs/SOLIDWORKS_API_VALIDATION.md`), `R23_TouchAllModules` (localises
  a compile failure to the first module that fails to load), and
  `R23_RunAllProbes` (compile, touch, activate the correct document, run
  all nine probes in order, isolated so one error can't abort the rest).
  Contains no `allowMutation` and calls no mutating procedure.
- `Public Sub R23_CompileTouch()` added to all 21 deployed standard
  modules (the original 19 plus the two new ones).
- `tools/probe-runner/Run-R23Probes.ps1` (new): optional `-Deploy`,
  opens the authorized part then the drawing read-only via `OpenDoc6`,
  refuses any part path outside the three `Module1_Main` fixtures, and
  invokes `R23_RunAllProbes` inside a background PowerShell runspace so a
  modal compile dialog reports a timeout instead of hanging. `-DrawingPath`
  is mandatory with no default: no authoritative drawing path exists in
  this repository, and guessing one would violate the "never guess a
  path/contract" rule.
- `tools/swp-deploy/Invoke-SolidWorksMacro.ps1` (new): the RunMacro2
  invoker extracted verbatim out of `Deploy-TargetSpecHybrid.ps1` so both
  scripts dot-source the same implementation instead of two copies
  existing. `Deploy-TargetSpecHybrid.ps1` behaviour is unchanged -
  reconfirmed with `-PreflightOnly`.
- `deployment-manifest.json`: registered both new modules (36 to 38
  managed components). Inventory-lock test and preflight output updated
  to match.
- `test_r23_probe_runner_contracts.py` (new, 16 tests): sink hygiene, no
  raw `Debug.Print` survives in any probe module, the runner calls all
  nine probes in dependency order and stops before them on an unclean
  compile, no mutating marker or `allowMutation` anywhere in the runner,
  every standard module carries `R23_CompileTouch`, source hygiene for
  both new modules. Full suite: 398 tests, the same five known-stale R20
  failures and nothing else.
- `MACRO_SOURCE_REVISION` intentionally not bumped, per the plan: these
  modules are not on the reachable production path yet.

## 2026-08-02 - Phase 8 gate satisfied; probe automation authorized

### Phase 8 SATISFIED

Third live run: satisfied=7, missing=0, duplicated=0, sectionOrdinates=0,
requirementFailures=None, mutations=0, drawingUnchanged=True, selection 0
before and after.

Both bores resolved through diameterDisplaySource=TextPrefix with the
MOD-DIAM prefix token. The drawing carries the diameter symbol in the
dimension's text prefix while Diametric stays False - the case the third
reading of R23-804 was written to distinguish. Both GetText forms return
the literal token rather than a rendered glyph, so the codepage-216
comparison is dead code on this build.

### The fixture guard fired, correctly

The preceding attempt returned
R23_SECTIONDIM_FATAL|reason=UnauthorizedFixture with the part resolved to
the V: network copy instead of the authorized test_assets copy. Opening the
local part BEFORE the drawing rebinds it. If earlier runs bound different
copies, that is an alternative explanation for the Phase 0 versus
2026-08-01 dimension-state difference, and the part= field on every _BEGIN
line settles it.

### Compile gate amended for read-only probes

User-authorized. A strictly read-only R23_Probe* entry point may now be
deployed and run without a preceding manual Debug > Compile Project: a
probe that does not compile fails at its first statement, so the gate was
buying nothing there. Mutating runs and production acceptance are
unchanged, and the exemption does not extend to anything gated behind
allowMutation.

Programmatic full-project compilation is also authorized. The mechanism is
already proved here - Module0_SourceDeployment.bas reaches the VBE object
model and executes a built-in command through CommandBars.FindControl
because VBProject.SaveAs raises 748 on a host-managed .swp.

Amended: Agents.md, CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md,
R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md.

### Probe automation planned, not built

docs/R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md: an evidence sink writing
every probe line to a file as well as the Immediate Window, a runner that
compiles and executes all nine probes, and one PowerShell command that
deploys, opens the fixture in the correct order, invokes the runner and
prints the log path. Two risks are stated rather than assumed away: the
compile control ID must be proved by enumeration before it is hard-coded,
and a compile failure raises a modal dialog no VBA-side design can avoid.

### README_IMPORT corrected

Its import instructions still described nine standard modules and three
data classes, wrong since Phase 1 and flagged by the code review. It now
defers to tools/swp-deploy/deployment-manifest.json, and the stale 74-test
figure is now 381.

## 2026-08-01 - R23 Phase 10: semantic QA replaces count-based checks

`Module19_SemanticQA.bas` (25 procedures). Statically verified only.

### Ten required stages, each with an evaluator that can fail it

MODEL_INTENT_CATALOG, MODEL_IMPORT_COVERAGE, NATIVE_CALLOUT_COVERAGE,
PHYSICAL_LOCATION_GRAPH, VIEW_PROJECTION, ORDINATE_SCHEME, SECTION_GEOMETRY,
SECTION_DIMENSIONS, the retained MANUFACTURING_DEFINITION, and FINAL_LAYOUT.
Declaring a stage is what gates the run: SealRequiredStages turns any stage
nothing proved into a named failure rather than an absence nobody notices.

### The module changes nothing at all

Not even behind an allowMutation gate, unlike every other R23 module. A QA
engine that repairs what it is judging cannot report on it, and a contract
asserts the absence of every mutating call the other phases own.

### What the counts missed

"Did anything get imported?" is satisfied by importing every dimension into
one view and none into the others. Coverage is now per view AND per
category - accepted projections, covered in X, covered in Y, annotations
attached - and a view with accepted projections and none of the three fails
by name.

"Does the note contain the expected text?" passes on free text that has
drifted from the geometry beside it and fails on a correct drawing whose
wording differs. Section dimensions are judged by type, nominal, attachment
and tolerance; GetText, GetNotes and note-text searching appear nowhere in
the module.

A projection count hides an unprojected location behind the ones that did
project, so every identity-proven location with no accepted projection is
named individually.

### Provenance, types and duplicate keys

Every manufacturing field is emitted beside its proof source, and a value
whose source is blank, "None" or "Unproven" fails as NoProvenance - 6.6 is
correct or wrong depending on whether it was read from the feature, read
from a callout variable, or assumed.

Every audited feature emits raw type, effective type, resolution source,
operation kind and rejection reason - every feature, not every accepted one,
because a rejection with no type recorded cannot be reviewed. An empty
effective type and the three outcomes Module12 actually writes -
IceUnresolved, Unresolved, ReadError - fail the catalog, and a contract
checks those literals against Module12 rather than trusting the copy.

DuplicateKeyReport names every repeated key and its count across physical,
family definition and section requirement keys, and returns "None" rather
than an empty string so "no duplicates" and "the check did not run" cannot
be confused.

### Decision logic is not duplicated

CollectRetainedDefinitions runs its own loop but takes every judgement from
Module16_CalloutDefinition's public surface - IsNativeHoleCallout,
MatchCalloutToFamily, BuildDefinitionFromTypedData, ReadNativeCalloutFields,
RetainDefinitionForFamily - so the two cannot drift on what matters.

### Deferred, and why Phase 11 was not started

Module6_QAEngine still runs the count-based checks on the reachable
production path; switching over is Phase 11, the same deferral R23-609,
R23-704 and R23-810 already carry. Phase 11 also bumps
MACRO_SOURCE_REVISION and switches the deployable macro onto Phases 5 to 10,
none of which has a green live run yet.

### Verification

Procedure blocks balanced 21/21 and 3/3. ANSI-only bytes, CRLF, no BOM, max
line 79. 24 Phase 10 contracts. Suite 381 tests with the five known-stale
R20 failures. Preflight 36 managed components. MACRO_SOURCE_REVISION
unchanged at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 second live run: R23-907 reversed, six defects fixed

### Phase 8 matched every requirement

satisfied=5|missing=0|duplicated=0, with all seven nominals exact. The
nominal route is settled: every dimension answered
nominalRoute=Obsolete.GetSystemValue2, and GetSystemValue3 declined both
configuration modes on all seven. swAllConfiguration has been removed - a
route with live evidence against it is not kept for insurance.

The two flagged requirements are the bore diameters, on
NotDisplayedAsDiameter:2. All seven returned diametric=False with
diametricKnown=True, a real answer. A drawing can carry the diameter symbol
in the dimension's TEXT PREFIX while the diametric flag stays False, and
then the sheet reads correctly even though the record does not.
ReadDiameterPrefix reads GetText(swDimensionTextPrefix) and its
...PrefixDefinition form, where SOLIDWORKS writes <MOD-DIAM>, and
DiameterDisplaySource names which of DiametricRecord, DiametricFlag or
TextPrefix answered. NotDisplayedAsDiameter is recorded only when all three
decline.

### Phase 9 ran end to end, and the arrow block was the bug

Drawing View4 returned items=49 from GetSectionLineInfo2. That is
2 header + 1 numSegments + 7x3 segments + 9 + 9 arrows + 7 text: an arrow
block is start[3] + end[3] + width + height + style = 9 doubles, not the 11
the first version counted, and three segments is the J-J path exactly. With
11 nothing matched, which is why arrow=0|section=0 on every envelope. The
dry-run grammar check did its job - it refused to parse rather than
producing plausible coordinates.

Four more from the same run: a view with no section line was reported as a
failed parse rather than as having no section line; every envelope printed
twice, the same AddInfo-already-prints defect fixed in Phase 8 and missed
here; the display-data frame check allowed 120 mm of slack and tested only
line start points, so its 26/28 consistent counts were weaker than they
looked; and 34 rejected off-sheet points were counted without one
coordinate being kept, which makes a frame error and genuinely off-sheet
geometry indistinguishable. The eleven *Front/*Top/*Isometric template
entries GetViews returns are now skipped by name.

### R23-907 reversed by the user

"The views are allowed to rescaled as per need". The accepted reference
drawing cannot satisfy the old prohibition: its four view envelopes need
0.479 m of height in the 0.253 m available.

The prohibition is replaced by a gate and a record. The only ScaleDecimal
assignment is inside ApplyScaleToFit, which refuses without allowMutation,
records the mutation, and reads each new scale back. PlanPlacement returns
plan=RescaleRequired with a suggested factor labelled
factorIs=GeometricEstimateTextDoesNotScale, because annotation text height
does not scale with the view - so the factor is applied, the drawing is
rebuilt, and the envelopes are RE-MEASURED rather than scaled
arithmetically. ReportScaleChanges names every view whose scale changed with
before and after values.

R23-908 survives: rescale happens once, and if the content still does not
fit, ApplyPlacementPlan returns layout=Reject|reason=LargerSheetRequired.

### Verification

Suite 357 tests with the five known-stale R20 failures. Preflight 35
managed components. MACRO_SOURCE_REVISION unchanged at
target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phases 8 and 9 first live run: four defects fixed

Both probes ran read-only against the P-0251 reference drawing. Neither gate
is satisfied; every cause was in my code, not the drawing.

### Phase 8 defect 1: the nominal never read

All seven section dimensions returned nominalAvailable=False while the same
dimension object answered toleranceType, fitType and the fit strings on the
next line - so GetSystemValue3(swThisConfiguration, Empty) declined
specifically. The seven are RD1..RD7@Drawing View6: drawing-authored
REFERENCE dimensions, which have no configuration to ask about. Phase 0 read
D1@Sketch4, an imported model dimension, where that route works.

Matching needs the nominal, so all seven requirements reported Missing while
their dimensions sat in the view. TryReadNominal now tries
swThisConfiguration, then swAllConfiguration, then the obsolete
GetSystemValue2("") and SystemValue as labelled last resorts, and NAMES the
route that answered so a later run can drop whichever proved unnecessary.
When every route declines it reports the raw shape of the GetSystemValue3
result, because "no nominal" and "an empty SafeArray" are different
problems.

### Phase 8 defect 2: the type rule rejected the real drawing

Every section dimension is swLinearDimension=2, including the one carrying
H7. Phase 0's type-6 D1@Sketch4/D1@Sketch6 evidence describes an earlier
state of the same fixture. Both states are real, so a diameter requirement
now accepts type 6, type 15 and the linear types, and
IDisplayDimension.Diametric is recorded to say which the drawing actually
displays as a diameter. It is reported rather than used to reject: the
nominals are 5.5 mm apart at the closest, so type corroborates and does not
discriminate. An unproved diameter display is recorded as
NotDisplayedAsDiameter or DiameterDisplayUnreadable, and any requirement
carrying its own failure is kept out of the satisfied count.

### Phase 8 defect 3: every log line printed twice

CRunEvidence.AddInfo prints what it records; the probe printed the same
requirement lines again.

### Phase 8 result worth keeping

RD4 carries toleranceType=8, fitType=0, holeFit=H7, minimumM=0.000000,
maximumM=0.000025, both statuses 0, two attached edges. H7 +0.025/0.000,
live, on a drawing reference dimension - independent corroboration that the
fit is drawing-authored and absent from the model.

### Phase 9 defect: a probe leaning on a production gate

The run aborted with "Controlled sheet has neither an ITitleBlock definition
nor a proved legacy title-block rectangle" before building a single
envelope. R23_ProbeContentEnvelope called
Module8_RuntimeSupport.MeasureControlledSheetRegions - a fail-closed gate
for a sheet the macro CREATES from the controlled template, which the
designer's reference drawing is not. Worse, that procedure SETS
ISheet.SheetFormatVisible, so a run promising mutations=0 had already
attempted one. Same shape as the Phase 5 EmitRunEvidence mistake.

MeasureSheetRegions now measures read-only - ISheet.GetSize plus
ISheet.GetZoneMargin - and REPORTS what it cannot measure
(titleBlock=Absent, contentBorder=Unmeasured,
usableSource=SheetExtentNoBorder) instead of aborting. Only an unusable
sheet size stops the probe.

BuildProtectedRegions now gates every rectangle on measured bounds. Unset
evidence fields would have produced a degenerate rectangle at the sheet
origin and reported false ProtectedIntrusion violations against it: a
boundary that does not exist is not a boundary at zero.

### Verification

Suite 345 tests with the five known-stale R20 failures. Preflight 35
managed components. MACRO_SOURCE_REVISION unchanged at
target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 9: content-envelope-aware final layout

`Module18_ContentEnvelope.bas` (33 procedures), `CContentEnvelope.cls`.
Statically verified only; no live run yet.

### What an envelope is

Everything that travels with a view: model outline, display-dimension lines
and text boxes, note extents, leader points, section segments, arrow
geometry, and both J-labels with their text heights. Each source is counted
separately, and HasAnnotationContent refuses to call an outline-only
rectangle a content envelope - that is the old behaviour wearing a new name.

### Frames, which is the real work

Four sources document their frame and one does not. GetOutline is page
frame; IAnnotation.GetPosition is sheet-relative in drawings;
INote.GetExtent is sheet space; GetSectionLineInfo2 is VIEW-SKETCH frame,
proved by Phase 0's payloadSegmentFrame=ViewSketchProved; and the
IDisplayData Remarks state no frame at all.

Section geometry is converted through ViewSketchToPage, the exact inverse of
Module17's PageToViewSketch, and ProveInverseTransform round-trips a real
page point through both before anything is contributed. Two functions that
claim to be inverses either agree to floating-point noise or one is wrong.

Display-data points are contributed and their agreement with the view's own
documented outline is COUNTED, not asserted. Claiming a frame the Help does
not state would be the same confident guess that cost this project the
swInsertDimensionsMarkedForDrawing bug.

### Three specific traps

GetTextPositionAtIndex is an OFFSET from the display-data origin, not a
coordinate; used absolutely it drags every envelope towards the sheet
origin. Leader points are consumed as XYZ triples from the returned array
rather than derived from GetLeaderStyle, whose value is OR-ed with
attachment bitmask flags that the corpus returns with mangled values. And
GetSectionLineInfo2's grammar is ambiguous between its Remarks and
GetSectionLineCount2's - one layer double or one per section line - so both
readings are walked in a dry run and the one whose consumption matches the
array length exactly is the one used.

### The fixed upward bias is replaced, not adjusted

Module9_LayoutEngine lines 442-446 pin the P-0251 source row to the top
boundary. PlanPlacement packs rows from the envelopes' own widths and
centres the block in the usable rectangle; contracts assert topBoundary -,
Bias and rowCenterY are all absent. A row pinned to a boundary has nowhere
to put the annotations that hang above it.

Placement and movement are separate procedures. PlanPlacement returns target
centres and touches nothing, so the whole plan is inspectable before a view
moves.

### Clearance and the things layout may not do

Every view-view and view-protected pair is checked with a separating-axis
measure, so touching rectangles score zero rather than passing, and the
check count is reported so an empty loop cannot read as a pass. Section
views get 2 mm from protected regions. The content border is protected as
four strips, not one rectangle - the drawable area is inside it.

No ScaleDecimal assignment, no ScaleRatio, no SetScale anywhere: R23-907
forbids shrinking a view to force a fit, so content that does not fit is
LargerSheetRequired with the required and available sizes stated. SealLayout
photographs the mutation ledger when layout completes so R23-909 can prove
nothing was created afterwards.

### Verification

Procedure blocks balanced 33/33 and 11/11. ANSI-only bytes, CRLF, no BOM,
max line 76. 32 Phase 9 contracts. Suite 334 tests with the five known-stale
R20 failures. Preflight 35 managed components. MACRO_SOURCE_REVISION
unchanged at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 8: semantic section-dimension engine

`Module10_SectionDimensionEngine.bas` (28 procedures),
`CSectionRequirement.cls`. Statically verified only; no live run yet.

### Reconcile before create

The section already carries imported dimensions - Phase 0 counted seventeen -
so `ReconcileSectionDimensions` runs before any creation path and
`CreateSectionDimension` refuses outright for a requirement that already
matched. Each requirement records six independent observations about what it
matched: source dimension identity, attached geometry, semantic role,
nominal, type and tolerance. Nominal and accepted type decide the match; the
other four are recorded so the match can be audited.

Every match is counted, not just the first, so R23-811 can fail on a
duplicate rather than quietly dimensioning the same thing twice.

### REQUIRED and OBSERVED never touch

CSectionRequirement splits what the specification demands from what was read
back, and nothing writes an OBSERVED field from a REQUIRED one. A
requirement that reports its own nominal back as the observed nominal proves
nothing; a contract asserts the assignment never appears.

### The obsolete tolerance route is gone

All four IDimension tolerance members - GetToleranceValues,
SetToleranceValues, GetToleranceFitValues, SetToleranceFitValues - are
marked obsolete by the 2025 Help, each superseded by an IDimensionTolerance
member. The Phase 0 probe used them; production reads
IDimension.Tolerance instead.

GetMinValue2 and GetMaxValue2 return a STATUS and hand the value back by
reference, so the status is reported beside the value it qualifies. A zero
value with a failed status is not a zero tolerance.

### H7 provenance is enforced, not just documented

REFERENCE_HOLE_FIT, the two deviations and
REFERENCE_FIT_AUTHORITY = "TargetSpecReferenceAuthority.NotModelData" are
each stated once. ApplyReferenceFit sets swTolFITWITHTOL BEFORE the values,
because SetValues2 refuses while the type is swTolNONE by its own Remarks
and FitType is only available for the fit types; it then reads the result
back rather than trusting the two return values, normalizing both COM
booleans.

EvaluateTolerance will not claim model provenance for a tolerance it merely
found on the drawing. Present-on-drawing is recorded as PresentOnDrawing
plus the same reference authority, because Phase 0 read the part source
directly and proved it carries none.

### No free text, ever

No InsertNote, no CreateText, no SetText anywhere in the module, and every
failure exit from CreateSectionDimension carries
policy=NoFreeTextSubstitute. A note is not a dimension: it does not move
with the geometry and cannot be inspected as one.

### Per-dimension locals reset every iteration

VBA block-scoped locals live for the whole procedure. That is exactly how
the Phase 0 section inventory mislabelled eleven of its seventeen
dimensions - index 6 set DIAMETER_40 and indices 7 to 17 kept the label.
Both loops here reset every field they report.

### Deferred, all for the same reason

R23-803's creation half needs live entity selection. R23-808 assigns lane
NAMES and leaves coordinates to Phase 9, which owns the annotation
envelope - the same call Phase 7 made about the section view's placement.
R23-809's predicate exists but Module9_LayoutEngine does not consult it, and
R23-810 detects the Module7_TitleBlockEngine free-text bore callout without
removing it, because removing it before real dimensions exist would leave
the bore undefined.

### Verification

Procedure blocks balanced 28/28 and 5/5. ANSI-only bytes, CRLF, no BOM, max
line 76. 37 Phase 8 contracts. Suite 302 tests with the five known-stale R20
failures. Preflight 33 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 7 read-only gate satisfied

`resolvedPaths=1|segments=3|columnHoles=3|crossingsProven=4|
sectionFailures=None|creations=0|drawingUnchanged=True`.

### The path resolved as the reference approves

In Drawing View4: w1=0.207331779,0.237414746 (bore centre);
w2=0.207331779,0.167414746 (same X, highest row); w3=0.192331779,0.167414746
(minimum-X column, that row); w4=0.192331779,0.087414746 (same column,
lowest row). distinctColumns=2, distinctRows=3, and crossingsProven=4 - the
bore plus all three holes on the chosen column.

The other three views report NoAcceptedSingletonBoreProjection, which is
correct: the bore is not accepted in them, so no path is invented.

### The frame transform cross-checks against the model

The arithmetic is exact (0.207331779 - 0.229331779 = -0.022), but that only
proves the code does what it says. The corroboration is independent: the
bore's Plucker moment is (0, 0.062, 0) and its viewY is 0.062 exactly. Every
counterbore behaves identically - viewY equals its moment's Y (-0.008,
-0.048, -0.088) and viewX equals its moment's X minus a constant 0.022, the
view's own centring offset, the same across all seven holes.

A wrong transform does not produce one shared offset across seven
independent points. This matters because Phase 8 onward depends on this
conversion and mixed frames have been a real defect in this project before.

### One evidence defect the run exposed

Views rejected before crossings could be tested - no bore, too few columns -
were also reporting NotAttempted in sectionFailures. That is the crossing
proof's initial STATE, not a failure of it, and listing it beside reasons
that are real dilutes them. Now excluded.

### Still unrun

R23-705's creation half, R23-706 and R23-707 all require mutation. The
selection-order verification and the GetSectionLineInfo2 read-back are
written and contract-tested but have never executed.

### Verification

23 Phase 7 contracts. Suite 265 tests with the five known-stale R20
failures. Preflight 31 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 7: J-J section path from model intent

`Module17_SectionPath.bas` (21 procedures), `CSectionPath.cls`. Statically
verified only; no live run yet.

### The path

Four waypoints, three segments, every coordinate a proved projection's page
coordinate: bore centre; same X at the highest face-hole row; minimum-X
column at that row; same column at the lowest row. R23-703 then proves the
path actually crosses the bore and every hole on the chosen column, naming
each one it misses.

### What is deliberately absent

No `extension`, no `topY`/`bottomY`, no `leftX`/`rightX`, and none of the
fractions 15/72, 90/196, 15.84/24 or 0.1x. A percentage of a view outline
knows nothing about where the holes are, which is exactly why the old upper
label landed in the zone region and the lower arrow in the
part-identification band. Contracts assert each token is gone.

### Contracts worth carrying

- The bore is the largest SINGLETON-family location, read from the graph.
  No radius threshold, so a different part is not misclassified.
- The grid is proved: fewer than two columns or two rows is a stated
  rejection, not an array index that happens to work.
- Crossings are judged against each hole's own projected radius, and the
  point-to-segment distance is CLAMPED to the finite segment. Unclamped, a
  circle beyond an endpoint reports as crossed because the infinite line
  passes through it.
- The page-to-view-sketch conversion happens exactly once per waypoint,
  immediately before CreateLine. Nothing upstream holds view coordinates, so
  there is nothing to convert twice.
- Segment selection order is verified before CreateSectionViewAt5, whose
  Remarks require the section line to be selected first. SOLIDWORKS reads
  the segments in selection order, so an unverified order cuts a different
  shape.

### Two defects caught before compiling

CreateSectionViewAt5 was being passed an empty label instead of the resolved
one. And its X/Y - the CENTRE of the new view - were being defaulted to
waypoint 3, a point inside the source view, which would have stacked the
section on top of the view it was cut from. Placement is now a caller
argument: choosing where a view sits is layout, and layout needs the full
annotation envelopes this module cannot see.

### R23-704 is half met, deliberately

Same shape as R23-609. The new path is clean; the legacy literals stay in
Module2_DrawingPipeline.bas (1525-1556) because it is the reachable
production path and Module17 is not wired into main.

### Verification

Procedure blocks balanced 21/21 and 4/4. ANSI-only bytes, CRLF, no BOM, max
line 77. 22 Phase 7 contracts. Suite 264 tests with the five known-stale R20
failures. Preflight 31 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 6 read-only gate satisfied

`definitions=3|definitionFailures=None|counterboredFamilies=1|
threadedFamilies=1|shapeFailures=None|nativeCallouts=2|creations=0|
drawingUnchanged=True`.

### Corroboration, not just self-consistency

The M5 family's depth resolved to 12.4 mm from
swCalloutVariable_Tap_Drill_Depth (28), and its thread depth to 10 mm. The
legacy hardcoded string in Module7_TitleBlockEngine.bas reads
"4.2 x 12.4 DEEP" and "TAP M5x0.8-6H x 10 DEEP". The same numbers arrived by
a completely different route - typed callout variables instead of a human
typing them - which is independent evidence that the derivation is correct.

### Both retention branches exercised live

The M5 family kept its native callout
(reason=CompleteAssociativeDefinitionAvailable). The counterbore and
stepped-bore families retained ControlledFallback with
reason=NoNativeCalloutAttributedToFamily. An earlier run also produced
reason=NativeIncomplete|nativeMissing=Depth, so R23-605 is proved in both
directions rather than only the one the fixture happened to take.

### One more mislabel the passing run exposed

threadedFamilies came back as 2. A thread DESCRIPTION was being treated as a
thread, and the counterbored clearance-hole family carries the fastener size
of the screw it clears, with threadDepthM=0. A hole that is actually tapped
has a thread depth. Now 1, which is the truth.

This one mattered less than the others - the gate would have passed either
way, because the M5 family is genuinely tapped. It was still wrong, and a
shape classifier that miscounts is a shape classifier that will eventually
pass something it should not.

### Verification

27 Phase 6 contracts. Suite 242 tests with the five known-stale R20
failures. Preflight 29 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 6: callout reconciliation and controlled fallback

`Module16_CalloutDefinition.bas` (20 procedures), `CCalloutDefinition.cls`.
Statically verified only; no live run yet.

### The shape of a definition

Either a NATIVE SOLIDWORKS hole callout carrying the Hole Wizard's own data,
or a CONTROLLED FALLBACK assembled field by field from the typed feature
data Phase 2 proved. Never free text. Every field carries its proof source,
so a definition that looks complete can still be shown to be unproven, and
R23-610 reports which field failed rather than that something did.

### Contracts worth carrying

- **R23-601: IsHoleCallout is the only classifier.** A native callout
  reports Type2 = 6 (swDiameterDimension) and so does an ordinary diameter
  dimension. No dimension-type constant is declared in the module, so none
  can be reached for.
- **R23-603: fields come from IDisplayDimension.GetHoleCalloutVariables**,
  reading HoleFit, ShaftFit, ToleranceType, ToleranceMin and ToleranceMax
  per ICalloutVariable. Parsing the rendered string would give something
  that cannot be validated field by field, which is the point of the task.
- **R23-602: a callout resolving to two families is rejected**, not
  tie-broken. Attribution is COM identity against every drawing entity a
  projection owns - not the anchor alone, because Phase 4 showed the anchor
  tier prefers the through hole while a counterbore callout attaches to the
  wider mouth.
- **R23-606: quantity is unique physical locations.** Not a feature count:
  one Hole Wizard feature plus a mirror produces many holes. Not an edge
  count: a counterbore contributes several edges per hole.
- **Depth is required only for a blind hole.** Demanding it from a through
  hole would fail every through hole. The stored end-condition code decides,
  and an unproven end condition fails on its own terms first.
- **R23-611 is stated as shapes rather than part numbers**: one multi-hole
  counterbored family and one multi-hole threaded family. P-0251 satisfies
  it; the rule does not name it.

### R23-609 is half met, and the remaining half is deliberate

The new path contains no part number, no 6X, no M5x0.8, no H7, no diameter
literal, and no scoring by feature name or by proximity to an expected
radius. A contract asserts each of those strings is absent.

The legacy literals are still in Module7_TitleBlockEngine.bas - callout text
at 359-371, name/radius scoring at 405-435 - because Module7 is the
reachable production path and Module16 is not yet wired into main. Deleting
them now would degrade the deployable macro while its replacement is
disconnected. They come out in the phase that switches the pipeline over.

### Mutation boundary

CreateNativeCalloutForFamily is the only procedure that creates anything. It
refuses without allowMutation, and refuses again without a proven anchor:
IDrawingDoc.AddHoleCallout2 attaches to whatever edge is selected, so an
unproven selection would produce an associative callout pointing at the
wrong hole and looking correct on the sheet.

### Verification

Procedure blocks balanced 20/20 and 6/6. ANSI-only bytes, CRLF, no BOM, max
line 78. A hygiene contract now asserts every byte is below 0x80, after a
UTF-8 em dash in Module14 broke SWP readback earlier today. 23 Phase 6
contracts. Suite 238 tests with the five known-stale R20 failures. Preflight
29 managed components. MACRO_SOURCE_REVISION unchanged at
target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 5 read-only gate satisfied

Final run: `schemes=4|horizontalSchemes=2|verticalSchemes=2|`
`creditedLocations=10|expectedLocations=10|coverageFailures=None|creations=0|`
`initialSelectionCount=0|finalSelectionCount=0|drawingUnchanged=True`.

### Proven live

- **R23-500** four schemes, keyed by view role + machining face + datum
  policy + direction. The top-face and side-face families separated without
  any view name being read, because machining face comes from the location's
  sign-normalized axis.
- **R23-501** `Drawing View4` horizontal datum is the stepped bore's
  projected centre, chosen by the CentreBoreProjectedCentre policy rather
  than by fallback, and proved selectable with ownership confirmed after the
  fact through ISelectionMgr.GetSelectedObjectsDrawingView2.
- **R23-503 / R23-504** two X buckets and three Y buckets in the primary
  view, every bucket anchored.
- **R23-505** all four side holes credited across two page positions.
- **R23-507** profileEntries=1; the stepped bore never enters a bucket.
- **R23-509** ten small-hole locations, ten credited, no coverage failures.

### Open, and stated as open

- **R23-502 is NOT met.** The vertical datum resolves to the lowest
  projected hole and is recorded as datumKind=ProjectionDerived. The task
  asks for bottom outline geometry, which is a different kind of entity.
  Recording the kind separately from the policy is what keeps this visible
  instead of letting a weaker datum satisfy the requirement quietly.
- **R23-506 half met**: four side locations resolved and credited, none
  dimensioned.
- **R23-508 unrun**: it mutates.

### Two defects the runs found

The probe called Module6_QAEngine.EmitRunEvidence, the production gate whose
RequireCoreStages demands fourteen pipeline stages a read-only probe never
performs. It failed closed and reported RESULT: FAIL for a run that had
resolved every scheme and datum it found.

Then the coverage gate reported credited=8 of an expected 10.
MarkCoincidentProjections sets CoincidentWithAnchoredKey on the UNANCHORED
projection - it explicitly skips anything that already has an anchor - and
the ledger read that field off the projection it was bucketing, which is by
definition the anchored one. The unanchored twin holding the link had
already been filtered out by the Accepted guard, so it was unreachable from
either side. Two of P-0251's four side holes went silently uncredited.

Both are now pinned by contracts, including one that asserts which end of
the coincidence link Module13 writes, so an inversion fails statically
rather than surfacing as a coverage shortfall.

### Verification

25 Phase 5 contracts. Suite 215 tests with the five known-stale R20
failures. Preflight 27 managed components. MACRO_SOURCE_REVISION unchanged
at target-spec-hybrid-v2-2026-07-29-r22.

## 2026-08-01 - R23 Phase 5: ordinate schemes and the transaction (source)

`Module15_OrdinateScheme.bas` (36 procedures), `COrdinateScheme.cls`,
`COrdinateBucket.cls`. Statically verified only; no live run yet.

### R23-500, the scheme key

view role + machining face + datum policy + direction. A family says what a
hole IS and nothing about which ordinate group it belongs to: two holes of
one family machined from opposite faces belong to different groups, and two
holes of different families machined from the same face in one view belong
to the same group. Every part of the key is measured - machining face from
the location's sign-normalized axis, view role from Phase 3's axis-normal
measurement - so a renamed or reoriented view still lands correctly.

`swOrdinate` (1) is deliberately never used. Letting SOLIDWORKS infer the
direction from the selected points would make the created dimension depend
on selection order rather than on the scheme.

### R23-505 and R23-509, counted the way the geometry allows

A bucket holds ONE selectable drawing entity and the LIST of physical
locations it represents. Phase 3 proved live that two coaxial holes viewed
along their shared axis produce exactly one drawing entity; dimensioning it
twice is a duplicate, and crediting only one hole silently drops the other.
Coverage is therefore counted per distinct page position and credited to
locations - the finding carried forward from Phase 3.

### R23-507, small-hole membership

Family size, read from the graph. P-0251's stepped bore is excluded because
it is a singleton, not because it is large. A radius threshold would look
equivalent here and misclassify a different part.

### R23-508, the transaction

Activate and verify the view, bind `ISelectData.View` (guarded - it is
documented get/set but raises error 91 on this build), select the datum
first, append in ascending coordinate order, verify the selection count
advanced by exactly one at every append, call `AddOrdinateDimension`, decode
all eleven `swCreateOrdDimError_e` members by name, call `SetPickMode`
whatever the result was, clear selections on every exit, then read back.

### Three defects caught before compiling

- `IsOrdinateEligibleView` and `IsDeferredCreationView` take
  `(graph, swView)`, not `(swView, graph)`.
- **`IView.GetFirstDisplayDimension5` is obsolete and its own Remarks say
  the `GetNext5` walk covers the drawing SHEET.** A read-back built on it
  would credit other views' dimensions to this scheme. Replaced with
  view-scoped `IView.GetDisplayDimensions`.
- A count-difference read-back would be inflated by any unrelated dimension.
  Replaced with a before/after snapshot diffed by `ISldWorks.IsSame`, exact
  `= 1`.

### Mutation boundary

`CreateOrdinateGroup` is the only procedure that creates anything. It
refuses without `allowMutation`, and refuses again when the datum is
unproven - the datum is the first selection and everything is measured from
it. `R23_ProbeOrdinateScheme` contains no `AddOrdinateDimension` call.

### Verification

36/36 and 9/9 and 7/7 procedure blocks balanced, ANSI/CRLF, no BOM, max line
78. 22 Phase 5 contracts. Suite 212 tests with the same five known-stale R20
failures. Preflight 27 managed components. `MACRO_SOURCE_REVISION` unchanged
at `target-spec-hybrid-v2-2026-07-29-r22`.

## 2026-08-01 - R23 Phase 4 gate satisfied; reverse route ruled out on evidence

### What the instrumented run proved

`IModelDocExtension.GetCorrespondingEntity2` returned Nothing for **all 38
annotations and every attachment, with error 0**:
`outcomes=draw1:unresolved:err0`, `resolved=0`, `modelEdgesTested=0`,
`eqMax=-1`. The call declines rather than fails, and because no comparison
ever executed, `eqMax=-1` also rules out the `swObjectEquality` Unsupported
hypothesis raised in the previous entry. That hypothesis is now closed.

The member is documented as returning the entity "in the underlying part or
subassembly". This is a **part** drawing with
`componentContext=DrawingContextOnly` — there is no component to descend
into. The reverse route is therefore unavailable here by construction, not
misused. It is retained (it is the documented direction and should work in
assembly drawings) and now reports
`reverse=UnavailableNoModelCounterpart` instead of an unqualified no-match.

### Reconciliation reclassified, not force-fitted

Unmatched annotations now record
`AuthoredDrawingEntityNoModelCounterpart` rather than
`NoAttachedProjection`, distinguishing "this drawing entity has no model
counterpart at all" from "it has one and no location owns it". Only the
second would implicate the ownership model.

**No positional or dimensional fallback was added.** Matching a dimension to
the nearest hole on the page is the precise failure the physical-location
model exists to prevent, and R23-407 forbids it explicitly.

### Gate

`annotations=38`, `coverageFailures=None`,
`COVERAGE|holeCallouts=2|ordinates=10|diameters=0|toleranced=1|withFit=1`.
R23-412 is required-category coverage — hole callout, toleranced dimension,
ordinate — and all three are present. **Phase 4 gate satisfied.**
`reconciled=1` is `RD1@Drawing View7`, which proves the R23-407 identity
mechanism; the remaining 37 are the designers' own reference dimensions.
Read-only boundary held on all six runs: `mutations=0`,
`drawingUnchanged=True`.

### Verification

33 procedures balanced, ANSI/CRLF, no BOM, max line 79. 36 Phase 4 contracts.
Full suite 190 tests, same five known-stale R20 failures. Preflight 24
managed components. `MACRO_SOURCE_REVISION` unchanged at
`target-spec-hybrid-v2-2026-07-29-r22`.

## 2026-08-01 - R23 Phase 4: reverse route instrumented after it changed nothing

### What the live run showed

The reverse-correspondence route ran on all 38 annotations and matched none.
`reconciled` stayed at 1 of 38, still the single `ForwardAlias` match on
`RD1@Drawing View7`. Every unmatched line carried
`routesTried=ForwardAlias,ReverseCorrespondence`, confirming the route
executed rather than being skipped.

**The evidence could not say why.** Two causes remained live and the log
could not separate them:

1. `GetCorrespondingEntity2` returned Nothing for every attachment.
2. It resolved, and `LocationOwnsModelEntity` rejected every result.

Predicting between them would repeat the mistake that cost the previous
iteration, so this change adds evidence rather than a fix.

### Added instrumentation

`MatchByReverseCorrespondence` now takes `ByRef diagnostics` and reports, per
attachment, the drawing entity's `swSelectType_e`, whether the reverse call
resolved, the trapped error number when it did not, and the resolved model
entity's type. It also totals `resolved`, `projectionsInView`,
`modelEdgesTested`, and `eqMax`.

`eqMax` is the load-bearing addition. **`ISldWorks.IsSame` returns
`swObjectEquality` — 0 NotSame, 1 Same, 2 Unsupported — and the existing
`ObjectsAreSame` wrapper collapses 0 and 2 into `False`.** A comparison that
*cannot be performed* across two documents is therefore indistinguishable
from a genuine non-match, which is exactly the shape of a silent zero.
`RecordEquality` keeps the raw code and reports the strongest one seen.

### Standing observations from the same run, not yet acted on

- `Section View J-J` reports its drawing component as
  `P-0251-14A-001-SectionAssembly-3-1/P-0251-14A-001-1`. Section views are
  built on a temporary section assembly, and section-cut edges have no
  counterpart in the part body. If the H7's two attached edges are cut
  geometry, no reverse mapping can exist for them. **Unverified** — the new
  `outcomes` field will show it directly.
- Several ordinate attachments report entity type `0` (`swSelNOTHING`), which
  is not valid input to `GetCorrespondingEntity2` (it documents vertex, face
  or edge). `SafeEntityType` returns `-1` on failure, so `0` is a genuine
  reading, not a trapped error.

### Verification

33 procedures balanced, ANSI/CRLF, no BOM, max line 79. 35 Phase 4 contracts
pass (2 new). Full suite 189 tests with the same five known-stale R20
failures. Preflight resolves 24 managed components. `MACRO_SOURCE_REVISION`
unchanged at `target-spec-hybrid-v2-2026-07-29-r22` — Phase 4 is not yet
wired into the deployable path.

## 2026-08-01 - R23 Phase 4: reverse correspondence route added (R23-302)

The stage instrumentation isolated the reconciliation failure in one run.

### Cause

**The counterbore hole callout attaches to a drawing edge of type
`swSelEDGES` that is none of the 18 aliases `IView.GetCorrespondingEntity`
produced for that view.** The forward model-to-drawing map is partial — it
returns 2 of each location's 4 boundary edges — so no amount of alias
comparison could ever reach the callout's edge. `Section View J-J` reported
`anchoredProjections=0|aliasesAvailable=0`, meaning its seven dimensions,
the H7 among them, were structurally unreachable by the forward route.

This was a design gap rather than a coding slip: **R23-302 asks for
drawing-to-model mapping and only the forward direction had been built.**

### Added

`MatchByReverseCorrespondence` uses
`IModelDocExtension.GetCorrespondingEntity2`, which the 2025 Help names for
resolving a drawing entity to its part entity. Each attached drawing entity
is mapped back to the model, then tested against the geometry each physical
location owns — its `SourceFaces` and those faces' boundary edges, via
`LocationOwnsModelEntity`.

The reverse route requires no projection anchor, so it can reach views where
the forward map produced nothing, including the section view. It remains
identity-only throughout (`ISldWorks.IsSame`), with no positional or
dimensional fallback. Successful matches record
`matchRoute=ForwardAlias` or `ReverseCorrespondence`, and unmatched
annotations report `routesTried=ForwardAlias,ReverseCorrespondence`, so the
two routes are never conflated in evidence.

### API contract

`IModelDocExtension.GetCorrespondingEntity2(Entity)` takes a vertex, face or
edge in a drawing view or assembly and returns the corresponding entity in
the underlying part, or Nothing. It is the documented counterpart to
`IView.GetCorrespondingEntity`.

## 2026-08-01 - R23 Phase 4 third live run: tolerance policy confirmed

Read-only boundary held: `mutations=0`, `finalSelectionCount=0`,
`drawingUnchanged=True`.

### Confirmed

The tolerance policy is in force and visible. The H7 reports
`toleranceAuthority=DrawingAuthoredNonAuthoritative`; every other dimension
reports `NoTolerance`. `dowelConvention` is gone from the evidence, and the
±10 µm constants are gone from the source.

### Not fixed: reconciliation is still 1 of 38

Matching every drawing entity a projection owns rather than only its anchor
was expected to raise the reconciled count to 3-4. **It did not move it.**
The prediction was wrong and the cause is still unknown.

Rather than attempt a second guess, `ANNOTATION_UNMATCHED` now reports what
each attachment actually is: `attachmentTypes` as `swSelectType_e` codes
(1 edge, 2 face, 3 vertex, 10 sketch segment, 11 sketch point, 28 centre
mark, 46 silhouette), together with `anchoredProjections` and
`aliasesAvailable` for the view. One run will now distinguish "the callout
attaches to an unmapped edge", "it attaches to a sketch segment", and "it
attaches to something that is not an edge at all".

Note the counterbore callout `RD3@Drawing View4` has `attachments=1` while
its projection has two mapped aliases, so the attachment is a single entity
that is not either mapped edge. That is the specific thing the next run
identifies.

## 2026-08-01 - R23 Phase 4 second live run; drawing tolerances demoted

Both fixes confirmed. The read-only boundary held again: `mutations=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

**`IDimension.SystemValue` is definitively correct for drawing dimensions.**
All 30 display dimensions now report
`nominalSource=IDimension.SystemValue=<value>|GetSystemValue3=0.000000000`.
Every value matches the reference drawing: 47.0 H7, 11.0 and 4.2 callouts,
R36, 173.6, 104.8, 80, 40, 25, 18, 12, 11.5, 6, and the ordinate chains.
`attachments` now reads 1-3 instead of 0. The H7 reads in full: nominal
`0.047`, `toleranceType=8`, `tolMaxM=0.000025`, `tolMinM=0.000000`,
`holeFit=H7`.

### Changed: drawing tolerances are not authoritative

Standing instruction from the user: the tolerances in the designers'
existing drawings were added manually to signal that *some* tolerance is
acceptable, not to state that the part holds them. They are evidence that a
designer typed a number, nothing more.

Every tolerance is now labelled
`toleranceAuthority=DrawingAuthoredNonAuthoritative`.
`ClassifyToleranceAuthority` deliberately has **no branch returning a
stronger authority**: nothing R23 can read distinguishes a binding tolerance
from an indicative one, and inventing a distinction would manufacture
authority that does not exist. The dowel ±10 µm rule and its constants are
**removed rather than parked** — the user is establishing with their designer
what part information should drive the decision to add a tolerance, and
nothing here guesses at it.

### Fixed

**Reconciliation matched only the anchor, so 1 of 38 annotations
reconciled.** A counterbore maps two edges; the native hole callout attaches
to the 11 mm mouth while the R23-308 anchor tier deliberately prefers the
6.6 mm through hole. They are different drawing entities of the same physical
location, so identity against the anchor alone failed. `ProjectionOwnsEntity`
now tests the anchor **and** every mapped alias the projection recorded.
Ownership is what reconciliation needs; the anchor is only the preferred
attachment point for new annotations. Identity only — still no positional
fallback anywhere.

**`alias` is a VBA reserved word** (`Declare ... Alias`) and would not have
compiled; caught by a new contract that also guards `name` and `type`.

## 2026-08-01 - R23 Phase 4 first live run: H7 provenance settled

The read-only reconciliation run held its boundary exactly: `mutations=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`. The manual reference
drawing was not altered. 38 annotations inventoried across four views,
`coverageFailures=None`, both hole callouts found.

### The H7 question is answered

`Section View J-J` carries `RD4@Drawing View6@P-0251-14A-001.Drawing` with
`toleranceType=8` (`swTolFITWITHTOL`), `nonZeroTolerance=True` and
`fitData=True`. **The Ø47 H7 fit is authored in the drawing, not the model.**
Phase 0 read `D1@Sketch4` directly and found no H7 and no nonzero tolerance
there; the fit exists only as a drawing-level reference dimension. R23 must
not expect model import to supply it, and must not record it with model
provenance when it does supply it.

### Dowel tolerance convention recorded

The user supplied the matching domain rule: **a dowel hole receives a
designer-added ±10 µm tolerance in the 2D drawing**, never in the model. Same
pattern as the H7 — precision tolerance is a drafting-stage act on this
drawing set. `MatchesDowelToleranceConvention` recognizes the ±0.010 mm
symmetric signature so provenance can follow from it. Recognition is
deliberately **not** treated as proof that a hole is a dowel: nothing in the
model says so, and P-0251's holes are counterbores, tapped holes and a
stepped bore, none proven to be dowels.

### Fixed

- **`nominalM` read 0 for every dimension** while `nominalAvailable` was
  True. `IDimension.GetSystemValue3` is configuration-scoped and these are
  *drawing* dimensions (`RD4@Drawing View6@....Drawing`); a drawing document
  has no configurations, so the read had nothing to resolve. Now taken from
  the configuration-free `IDimension.SystemValue`, with both readings kept in
  `NominalSource` so the disagreement stays visible.
- **`attachments=0` on every ANNOTATION line**, including the one that went
  on to reconcile successfully — the count was populated during the later
  reconciliation pass, after the line had printed. Attachments are now read
  during inventory. Evidence must never be printed before the field it
  reports is populated.

### Established by the run

- **The reference drawing authors its section diameters, H7 included, as
  `swLinearDimension` (2)**, not `swDiameterDimension`. Dimension type alone
  can never decide whether something is a diameter. Phase 0 separately showed
  *imported* section diameters arrive as type 6, so the two authoring routes
  differ.
- Annotation type 13 (`swCenterMarkSym`) accounts for the six `Other:13`
  entries in `Drawing View4`; now named.
- 38 annotations: 10 ordinates, 2 native hole callouts, 1 toleranced with
  fit, 5 cosmetic threads, 8 center marks, 1 note.

## 2026-08-01 - R23 Phase 4: annotation import and reconciliation

### Added

`src/target-spec-hybrid-v2/Module14_AnnotationImport.bas` (26 procedures)
implements R23-400 to R23-412.

**Mutation boundary.** This is the first R23 phase whose operations change a
drawing. Only `ImportModelAnnotations` and `RemoveR23CreatedAnnotations`
mutate, and both refuse unless passed an explicit `allowMutation` argument.
`R23_ProbeAnnotationReconciliation` never passes it and contains no insert,
delete or save call, so it is safe against the manual reference drawing. The
read-only path still exercises R23-406 to R23-409, R23-411 and R23-412 on
real data, because the reference drawing already carries the manufacturing
intent R23 must reproduce.

Reconciliation is by **attached-entity COM identity** against each
projection's proven anchor. Page proximity is used nowhere: it would attach a
dimension to whichever hole happened to be nearest, which is the failure the
physical-location model exists to prevent. R23-410 likewise deletes only this
run's own recorded annotation objects, never matching by name, position or
appearance.

Ordinate eligibility is decided from the graph's measured `AxisNormalToView`
data rather than from view names, so the isometric is excluded because no
hole axis is normal to it — the fact Phase 3 established — and a renamed or
reoriented view is still classified correctly.

### API contracts established

- **`swInsertAnnotation_e` mask `18055274` fully decomposed**, no unaccounted
  bit: datums 2, dimensions 8, GTols 32, notes 64, marked-for-drawing 32768,
  hole-wizard profile 65536, hole-wizard location 131072, hole callout
  1048576, toleranced dims 16777216. The callout member really is spelled
  `swInsertholeCallout`.
- **`InsertModelAnnotations4` is on `IDrawingDoc`, takes eight arguments and
  returns an ARRAY of `IAnnotation`**, not a count. There is no
  `InsertModelAnnotations3` on `IModelDocExtension`. Returning the objects is
  what makes the cleanup path safe.
- **`GetMinValue2`/`GetMaxValue2` return the STATUS**
  (`swDimensionToleranceWarning_e`) and deliver the value through an out
  parameter, which must be a local — the Phase 3 ByRef trap again.
- **`swTolFIT` and `swTolMETRIC` are both 7**, so a 7 is recorded as
  ambiguous rather than reported as one of them.
- `IAnnotation`, not `IDisplayDimension`, carries `GetAttachedEntities3`.
- `IDimension.GetSystemValue2` is obsolete; `GetSystemValue3` is used.

### Changed

- The deployment manifest now manages 24 components; the companion inventory
  lock moved from 23 to 24.

### Tests

Added `tests/test_r23_annotation_import_contracts.py`: 22 contracts, led by
the mutation-safety ones — the probe contains no insert, delete or save call;
both mutating procedures refuse without authorization; cleanup never matches
by name, position or appearance. Also asserts identity-based reconciliation
with no proximity fallback, tolerance out-parameters bound to locals, and the
mask decomposition. The suite is now 176 tests with the same five stale R20
failures.

### Not done

No live run. `MACRO_SOURCE_REVISION` remains `r22`: still no pipeline caller.

## 2026-08-01 - R23 Phase 3 gate SATISFIED; coincident projections explained

The fifth live run settled the remaining question, and disproved the
hypothesis recorded after the fourth.

**The hidden-lines explanation was wrong.** `Drawing View7` reports
`displayMode=HiddenLinesVisible`, so the far-face holes ARE drawn. The real
reason is geometric: that view looks along model Y, and the four tapped holes
lie on two axes with two holes each, so they project onto **two** page
points. The six counterbores likewise collapse to **three**. The observed
`mappedEdges` counts — 2 and 3 — equal the number of distinct page positions
exactly. Two coaxial holes seen along their axis are ONE circle on the sheet;
SOLIDWORKS holds a single drawing entity and `GetCorrespondingEntity` maps
only one of the two model edges to it. No search strategy can produce more
anchors than the drawing has entities.

**R23-310's second clause is therefore unsatisfiable as written, and the
implementation is correct.** Two side anchors is the complete answer. The
individual identity of all four tapped holes survives in `Drawing View4`,
which resolves them to four distinct page positions edge-on, and the
reference drawing calls them out once as `4x Ø4.2 ▼12.4 / M5x0.8 - 6H ▼10`.

### Added

`MarkCoincidentProjections` attributes every unanchored projection that
shares a page point with an anchored one, emitting `PROJECTION_COINCIDENT`
with `sharesPagePointWith` and
`reason=OneDrawingEntityForTwoCoaxialHoles`, and counting them as
`coincidentUnanchored` in the per-view summary. The page centre is now
recorded from the first face that projects whether or not an anchor is
found, since an unanchored location still has a provable position and
coincidence cannot otherwise be distinguished from a defect. Acceptance is
unaffected — the anchor test runs before the coordinate-frame test.

### Confirmed

`projectedAxis` now carries real values and corroborates the transform
independently: `0,0,1` for Z-axis holes in the front view, `0,0,-1` for
Y-axis holes in the side view, `1,0,0` and `-1,0,0` for the in-plane cases,
and `0.707107,0.408204,-0.577382` in the isometric — exactly
(1/√2, 1/√6, −1/√3), magnitude 1.000.

`IView.GetDisplayMode2` per view on this drawing: `Drawing View4` and
`Drawing View7` are `HiddenLinesVisible`, `Drawing View2` and
`Section View J-J` are `HiddenLinesRemoved`, and all ten sheet placeholders
report `Shaded`.

### Phase 3 gate

**Satisfied.** R23-300 through R23-310 are all proved live on P-0251, with
`selectionProved=9`, `finalSelectionCount=0`, `selectionClean=True`,
`drawingUnchanged=True`, `partUnchanged=True`. `MACRO_SOURCE_REVISION`
remains `r22`: the engine still has no pipeline caller.

**Carry into Phase 5:** required coverage must be counted per *distinct page
position per view*, not per physical location.

## 2026-08-01 - R23-310 fourth live run: primary-view clause proved

The axis gate behaved exactly as designed. Per-view acceptance:

| view | type | projections | axisNormal | anchored | accepted |
|---|---|---|---|---|---|
| `Drawing View4` (primary) | 4 | 11 | 7 | 11 | 7 |
| `Drawing View7` (side) | 4 | 11 | 4 | 6 | 2 |
| `Section View J-J` | 2 | 11 | 4 | 0 | 0 |
| `Drawing View2` (isometric) | 4 | 11 | 0 | 9 | 0 |

**R23-310's first clause is proved:** `Drawing View4` accepted all six
counterbores plus the stepped bore, every one anchored at
`anchorTier=PrimaryTypedHoleDiameter` on the Ø6.6 through hole rather than
the Ø11 mouth, and all seven proved selectable with `ownershipProven=True`,
`selectionClean=True`, both documents unchanged.

The gate's own correctness is corroborated by the axisNormal counts: 7 for
the Z-axis holes in the front view, 4 for the Y-axis tapped holes in both the
side and section views, and **0 in the isometric**, where no hole axis is
normal to the sheet. The isometric had 9 anchors and correctly accepted none.

### Fixed

**ByRef out-parameters were class fields, so the write-back was lost.**
`projection.ProjectedAxisX/Y/Z` were passed directly to
`AxisIsNormalToView`. A class Public variable is exposed as a property, so
VBA hands the callee a temporary that is discarded on return: the run logged
`projectedAxis=0.000000,0.000000,0.000000` on every line while `axisNormal`
was correct, because the function's own locals were sound and only the
write-back vanished. The values now go into locals and are assigned
explicitly. An audit found no other genuine out-parameter of this shape; the
other class-field arguments in Module12 and Module13 are input-only.

### Added

`IView.GetDisplayMode2` is now recorded per view as `displayMode=`. Under
`swHIDDEN` (Hidden Lines Removed) a far-side hole is never drawn, so
`GetCorrespondingEntity` has nothing to return and the location cannot anchor
in that view however the search is written. Recording it means an
unanchorable location can be attributed to the drawing's display setting
rather than mistaken for a mapping defect.

### Not done

**R23-310's second clause is not met: only two of the four tapped holes have
side projections.** The axis test is right — all four are `axisNormal` in
both the side and section views — but the two far-face holes return Nothing
from route A, and the section view maps nothing at all. Whether that is a
display-mode consequence or a genuine limit is what the next run settles.

## 2026-08-01 - R23-310 third live run: anchors resolve; axis gate added

The `GetCurveParams3` fix worked. `circularEdges` went from 0 to 4 (or 2) on
every location, `uMin=0`, `uMax=6.283185307`, `closureM=0.000000000` — the
Phase 0 circle proof reproduced exactly. **26 projections anchored**, every
one through route A with route C identity confirmation, and all 26 proved
selectable: `selection=True`, `ownershipProven=True`, `selectedCount=1` on
each, `initialSelectionCount=0`, `finalSelectionCount=0`,
`selectionClean=True`, `drawingUnchanged=True`, `partUnchanged=True`.

`ISelectData.View` raised error 91 on every attempt, as this repository
already recorded. The guarded binding plus after-the-fact ownership proof
handled it: `viewBinding=UnboundAfterError:91` with
`ownershipProven=True` 26 times out of 26.

### Fixed

**Normal-axis compatibility was computed but never enforced.** R23-306
requires it; `AxisNormalToView` was stored on the projection and ignored by
`QualificationFailureReason`. The run accepted all four M5 tapped holes in
`Drawing View4`, whose page coordinates prove that view's normal is model Z:
the counterbores land 40 mm and 30 mm apart, matching model X and Y, while
the M5 holes land at their model **Y** spread horizontally — their axis lies
in the page plane and the "circular" edge is seen edge-on. Such a projection
is not a usable circular anchor. It now fails as `AxisNotNormalToView`,
distinct from `ProjectionAnchorUnavailable` because the mapped anchor is
real; the verifier counts it separately rather than folding it into the
unanchored total.

**R23-310 could not be answered by a single number.** The requirement is
about which holes are usable in *which* view. Added `ViewAcceptanceSummary`
and an `R23_PROJECTION_VIEW_SUMMARY` line per view carrying `projections`,
`axisNormal`, `anchored` and `accepted`. `axisNormal` and `projectedAxis` are
now in the per-location evidence too, so the gate's own correctness is
visible rather than assumed.

### Established by the run

- Route A works from a model reached through `IView.ReferencedDocument` with
  the drawing active, settling the open question from the previous run. The
  earlier hypothesis about needing the part as active document was wrong.
- Mapping remains per-edge and partial: `mappedEdges` ranged 0 to 4 out of
  4 candidate edges on the same location in different views.
- `Section View J-J` mapped **nothing** (`mappedEdges=0` on all 11
  locations). Its visible component is a synthetic section assembly, so the
  original part's model edges have no correspondent. A section view needs a
  different acquisition route.
- Hidden geometry does not map: in the side view only the near column of
  counterbores (`py=-0.015`) and the near pair of tapped holes resolved.

### Not done

R23-310 stays open. With the axis gate applied the counts will change, and
whether "side projections for all four tapped holes" is achievable at all is
now in doubt — two of the four are on the far face and did not map.

## 2026-08-01 - R23-310 second live run: root cause found and fixed

The stage counters isolated the failure in one run. Every location reported
`sourceFaces=2, facesProjected=2, boundaryEdges=4, circularEdges=0` with
`firstReject=circle=Reject|reason=ReadError:438`, identically across all
three real views. Face retention, the page transform and edge enumeration
all work; every edge died at the same call.

### Fixed

**`IEdge.GetCurveParams3` returns an `ICurveParamData` object, not an array
of doubles.** `Module13_ProjectionResolution` assigned it into a Variant with
a `Let`, which asks the object for a default property it does not have and
raises error 438 — rejecting every candidate edge before any mapping was
attempted. The correct pattern was already in this repository:
`Module12_FeatureQualification.ComputeFaceAxialInterval` binds it with `Set`
and reads `.StartPoint` / `.EndPoint`, which is why the Phase 2 axial
intervals were always sound. Module13 now matches it and additionally
records `UMinValue` / `UMaxValue` in the circle proof.

**Unmeasurable closure failed open.** The new `PointDistance` helper reports
failure as a negative value, which would have passed the
`closureM > tolerance` test and read as a perfectly closed curve. It is now
rejected explicitly as `ClosureNotMeasurable` ahead of the tolerance test.

**Selection cleanliness had no baseline.** The run reported
`finalSelectionCount=1` while proving no anchors and selecting nothing —
the count was the operator's own pre-existing selection. The probe now
captures `initialSelectionCount` and reports `selectionAttempted` plus
`selectionClean`, so a non-zero count is only attributed to this pass when
this pass actually selected something.

An audit of every other `Let` assignment from an API member in Module13
found no further object-returning members; `GetCurveParams3` was the only
one.

### Established by the run

- `swDrawingViewTypes_e` on this build: the ten sheet placeholders report
  type `7` (`swDrawingNamedView`) with zero visible entities; the real
  projected views report `4` and the section view `2`.
- `ISheet.GetViews` returned 14 views, ten of them placeholders. The
  visible-entity skip removed 110 dead projections, taking the run from 154
  to 44.
- `Drawing View2` (65 edges) is a fourth real view, not seen in the first
  run's truncated log.

### Not done

No anchor has yet been proved. R23-310 stays open pending the rerun.

## 2026-08-01 - R23-310 first live run: no anchors; chain instrumented

The first `R23_ProbeViewProjections` run compiled and executed read-only
against the P-0251 drawing. The safety envelope held exactly:
`drawingUnchanged=True`, `partUnchanged=True`, `finalSelectionCount=0`, no
warnings, no failures. **No anchor resolved** — all 154 projections failed
`ProjectionAnchorUnavailable` with `candidates=0`.

Established by the run:

- the drawing side is sound — `GetVisibleEntities2` returned 64, 68 and 53
  edges for `Drawing View4`, `Section View J-J` and `Drawing View7`, each
  with a drawing component;
- `componentContext=DrawingContextOnly` on every real view, confirming the
  predicted part-drawing behaviour of `GetVisibleDrawingComponents`;
- `ISheet.GetViews` returns the sheet's six standard-view placeholders
  alongside real views, and they hold no drawing geometry.

### Fixed

**The evidence could not isolate the failure.** `candidates=0` is a single
number covering "no faces retained", "face centre not projected", "no
circular edge" and "nothing mapped" alike. This is the defect class Phase 0
solved with `rejectGate`, and it should not have recurred.
`ResolveProjection` now counts each stage — `sourceFaces`,
`facesProjected`, `boundaryEdges`, `circularEdges`, `mappedEdges`,
`inventoryConfirmed` — and carries the transform's own proof string into
`firstReject` rather than discarding it.

**Placeholder views generated dead projections.** A view with no visible
entities cannot supply an anchor, so it is now skipped with a single
`PROJECTION_VIEW_SKIPPED` line. 132 of the 154 projections were noise of
this kind. The test is the visible-entity count, which is a fact; the view
type is logged for evidence but deliberately not used as the filter, because
which `swDrawingViewTypes_e` code the placeholders carry is not established
on this build.

### Not done

The root cause of the zero anchors is **not yet diagnosed** and is not
guessed at here. R23-310 stays open pending the instrumented rerun.

## 2026-08-01 - R23 Phase 3: drawing-projection resolution

### Added

`src/target-spec-hybrid-v2/Module13_ProjectionResolution.bas` (27
procedures) turns each physical location into a `CViewHoleProjection` for one
drawing view: a selectable drawing-context anchor, a page-frame centre, and
the proofs behind both. It uses the routes Phase 0 settled on this build:

- route A, `IView.GetCorrespondingEntity`, is the acquisition path;
- route B, `IComponent2.GetCorrespondingEntity`, is attempted and recorded so
  a future build that fixes it is detected, but nothing depends on it;
- route C, `GetVisibleEntities2` plus `ISldWorks.IsSame`, is not an
  acquisition route at all — it is the independent confirmation that route
  A's result really is a drawing-context entity of this view.

Mapping is per-edge, so every boundary edge of every contributing face is
tried. `CPhysicalHoleLocation` gained `SourceFaces` (merged on consolidation)
so a projection can reach those edges without re-walking the feature tree.

Anchor choice follows the R23-308 priority order and ranks every candidate
rather than stopping at the first mappable edge.

`SelectAnchorInView` is the only code that selects anything; it clears the
selection before and after, guards the error-91 `ISelectData.View` binding,
and proves view ownership afterwards rather than assuming it.

### Contracts established

- **`ISldWorks.IsSame` is not a Boolean.** It returns `swObjectEquality`
  = {0 not same, 1 same, 2 unable to determine}. `NormalizeSwBoolean` would
  have accepted 2 as identity. Only an exact 1 is accepted; unreadable
  comparisons default to not-same and are counted in the evidence.
- **`GetVisibleDrawingComponents` is documented for assembly drawings**, so a
  part drawing yielding only the `GetVisibleComponents` handle is recorded as
  `DrawingContextOnly` rather than treated as a failure.
- **Direction vectors are differenced, not transformed** — applying a point
  transform to a direction folds in the view translation.

### Changed

- The deployment manifest now manages 23 components; the companion inventory
  lock moved from 22 to 23.

### Tests

Added `tests/test_r23_projection_resolution_contracts.py`: 20 static
contracts covering R23-300 through R23-310, including that the identity test
never runs through the Boolean normalizer, that the anchor loop has no early
exit, that unanchored locations are still added to the graph, and that the
pass contains no mutating call. The suite is now 142 tests with the same five
stale R20 failures.

### Not done

No live run. R23-310 is open: `R23_ProbeViewProjections` must confirm usable
primary projections for the six face holes and side projections for the four
tapped holes. `MACRO_SOURCE_REVISION` remains `r22` — still no pipeline
caller, drawing output unchanged.

## 2026-08-01 - R23-213 closed: P-0251 catalog proved live

The second `R23_ProbeFeatureCatalog` run returned `catalogFailures=None` with
all four R23-213 expectations met, 0 warnings, 0 failures and
`modelUnchanged=True`. `Mirror1` was accepted with
`seed=M5x0.8 Tapped Hole1`, `readStatus=SeedSemanticsInherited:Read` and
`thread=M5x0.8`, and the four M5 locations formed one family. The M6
counterbore reads `fit=Normal`; the tapped holes carry no fit. **The Phase 2
gate is satisfied.**

### Fixed

**Traversal was not exact-once.** Comparing the two runs, `visitedFeatures`
went 47 → 46 on an unchanged model, and which sketches were visited twice
differed between runs (`Sketch1`/`Sketch6`/`Sketch7` doubled in the first,
`Sketch4`/`Sketch7` in the second). `ObjPtr` was part of the traversal key,
so a feature reached both as a tree entry and as a consuming feature's
subfeature arrived through two COM wrappers at different addresses and
occupied two keys. The guard was too permissive, which is the opposite of the
failure it was written to prevent.

Neither catalog was affected — accepted features, locations, consolidations
and families were identical across both runs — but the defect is not
cosmetic. `CLocationGraph.AddFeatureDefinition` appends to `mFeatures`
unconditionally, and a repeated visit to an accepted feature would call
`AddStackMember` again, adding a duplicate radius and a duplicate stack
member so a counterbore would report `stackMembers=4`.

The key is now name plus type, which is exact-once because feature names are
unique within a part document. The address survives only as a fallback for a
feature whose name cannot be read, where a repeat visit is preferable to
collapsing several unnamed features into one key and skipping them.

### Not done

The traversal fix is static-only; it needs no dedicated run and can be
confirmed by a stable, duplicate-free `visitedFeatures` on the next live run.
`MACRO_SOURCE_REVISION` remains `r22`: the engine still has no pipeline
caller and drawing output is unchanged.

## 2026-08-01 - R23-213 first live catalog run; three defects found and fixed

The first read-only `R23_ProbeFeatureCatalog` run on P-0251 compiled and
executed cleanly (47 features visited, 0 warnings, 0 failures,
`modelUnchanged=True`) and reported `catalogFailures=NoFourLocationFamily`.
Three defects were found: the reported one, plus two the log exposed on
inspection.

### Fixed

**Pattern instances did not inherit their seed's semantics.** `Mirror1` was
accepted with `operation=PatternInstance` and every manufacturing field
blank, so its two mirrored M5 locations formed their own empty family. The
four identical M5x0.8 tapped holes were catalogued as two families of two and
the expected four-location family never formed. `ResolvePatternSeed` now
reads the seed through the same typed readers that qualify it in its own
right, and `CopySeedSemantics` carries the typed values across. Proof sources
are prefixed `SeedInherited(<seed>):` rather than reasserted, and a field the
seed could not prove stays unproven on the instance. A seed that is itself a
pattern is refused as `SeedIsPatternChainUnsupported`: following the chain
needs recursion whose termination this build gives no evidence for, and
P-0251 has no nested pattern.

**Seed-chain rejections were computed but never enforced.** `ResolvePatternSeed`
set `SeedChainUnresolved`, `SeedChainMultiplyResolved` or `SeedChainCircular`
and returned, but `QualifyFeature` never checked the result and the accept
path then cleared `RejectionReason`. Every one of those codes was dead: a
pattern with an unresolvable or circular seed would have been accepted
silently. It is now a `Function` whose result gates the accept path. The
existing tests asserted only that the reason strings appeared in the source,
which is why they passed; the new test asserts the check precedes the clear.

**`HoleFit` published a not-applicable code as a value.** The M5 tapped hole
returned `-1`, which was stringified into `FitDescription` and given a proof
source, putting a bogus clearance on a tapped-hole callout. The 2025 Help
limits the property to counterbore and countersink features and defines the
return as `swWzdHoleScrewClearanceTypes_e` = {0 close, 1 normal, 2 loose}.
`ScrewClearanceText` now maps those three and returns empty for anything
else, which the caller records as absent. The M6 counterbore's `1` is
genuine and now reads `Normal`.

`FEATURE_ACCEPTED` gained `seed=` and `readStatus=` so seed inheritance is
visible in the next log rather than inferred.

### Tests

Five new contracts (28 in the Phase 2 file, 122 in the suite, same five stale
R20 failures): seed rejections enforced ahead of the accept path, seed
semantics copied field by field, inherited proof sources carried and never
invented, seed chains refused rather than followed, and out-of-enum
`HoleFit` codes treated as absent.

### Not done

R23-213 stays open until a second live run confirms the corrected catalog.
`MACRO_SOURCE_REVISION` remains `r22`: the engine still has no pipeline
caller and drawing output is unchanged.

## 2026-07-31 - R23 Phase 2: feature qualification engine

### Added

`src/target-spec-hybrid-v2/Module12_FeatureQualification.bas` (28
procedures) walks the feature tree and populates a `CLocationGraph`:

- three-field type resolution, normalizing `GetTypeName2 = "ICE"` through
  `GetTypeName` and failing with a field-specific reason when unresolved;
- traversal of features and subfeatures behind a composite cycle guard
  (name plus type plus pointer), because pointer-only guards collapsed
  distinct features in an earlier probe;
- suppression proved against the drawing view's exact
  `ReferencedConfiguration`, with an active-configuration fallback permitted
  only when the two names are equal;
- typed readers for Hole Wizard, advanced holes, simple holes, exact
  `CUT`/`CUTTHIN` extrudes, cosmetic threads and pattern/mirror seeds, each
  pairing `AccessSelections` with `ReleaseSelectionAccess` through one shared
  release path;
- ownership from `IFeature.GetFaces`, with `IFace2.GetFeature` and
  `FaceInSurfaceSense` both absent from the executable path;
- physical locations built from each owned cylindrical face, with the axial
  interval measured from the face's own boundary edges so consolidation and
  separation are decided on real geometry;
- explicit reason codes for every rejection; and
- `VerifyExpectedCatalog` plus a read-only, fixture-guarded
  `R23_ProbeFeatureCatalog` evidence entry point.

`Module11_GeometryIdentity` gained `NormalizeSwBoolean`, the shared
`(CDbl(raw) <> 0#)` rule, now used for every SOLIDWORKS Boolean in the
production path.

### Changed

- The deployment manifest now manages 22 components.
- The companion inventory lock moved from 21 to 22.

### Tests

Added `tests/test_r23_feature_qualification_contracts.py`: 23 static
contracts covering R23-200 through R23-213, including that the module never
references `ModifyDefinition`, `FaceInSurfaceSense` or `IFace2.GetFeature` in
executable code, that consolidation is delegated to the graph rather than
reimplemented, and that the pipeline is not yet rewired. The suite is now
117 tests with the same five stale R20 failures.

### Fixed

`R23_ProbeFeatureCatalog` failed to compile in the VBE with "Method or data
member not found" on `evidence.InfoCount`. `CRunEvidence` keeps `mInfo`,
`mWarnings` and `mFailures` Private and exposes no item accessor, so the
replay loop could never have compiled. It was also redundant:
`CRunEvidence.AddInfo`/`AddWarning`/`AddFailure` each `Debug.Print` at the
moment of emission, and `EmitInfo`/`EmitWarning`/`EmitFailure` print again
under `mEmitDiagnostics`. The loop is replaced by an
`R23_CATALOG_EVIDENCE|warnings=…|failures=…` tally line built from the
public `WarningCount` and `FailureCount` properties. No qualification logic
changed. Every other cross-module reference in the seven new Phase 1/2
components was then checked against the target's public surface, and every
call site checked against its signature; no further mismatches exist.

### Not done

`MACRO_SOURCE_REVISION` remains `r22`: the engine has no pipeline caller yet,
so drawing output is unchanged. R23-213 stays open until one live
`R23_ProbeFeatureCatalog` run confirms the P-0251 catalog.

## 2026-07-31 - R23 Phase 1: location-graph model added (additive, no behaviour change)

### Added

Six new managed components under `src/target-spec-hybrid-v2/`:

- `CFeatureDefinition.cls` — one resolved feature, with a proof-source
  string beside every semantic manufacturing field.
- `CPhysicalHoleLocation.cls` — one physical location, identified by
  sign-normalized axis, line moment and axial interval.
- `CViewHoleProjection.cls` — one location projected into one view, with an
  explicit page-frame proof and the anchor route that produced it.
- `CImportedAnnotation.cls` — imported annotation identity and provenance,
  keeping reference-supplied tolerances distinguishable from model data.
- `CLocationGraph.cls` — nine indexes; the single consolidation point.
- `Module11_GeometryIdentity.bas` — canonical numeric normalization, with no
  SOLIDWORKS API calls so its invariants are checkable offline.

The three physical-identity invariants are enforced in one place,
`CLocationGraph.ResolveOrCreatePhysicalLocation`: coaxial steps consolidate
when their axial intervals meet; opposite blind holes on one infinite line
stay separate because theirs do not; and feature names appear in neither the
physical nor the family key.

### Changed

- `tools/swp-deploy/deployment-manifest.json` now manages 21 components,
  up from 15. `ThisLibrary` and both UserForms remain excluded.

### Tests

Added `tests/test_r23_location_graph_contracts.py` to the companion suite —
20 static source-contract tests for R23-106. The existing
`test_manifest_manages_only_replaceable_components` inventory lock was
updated from 15 to 21 components; all of its other guarantees are unchanged.
The suite is now 94 tests with the same five stale R20 failures.

### Not done

No deployment, no VBE compilation, no live run. `MACRO_SOURCE_REVISION`
stays at `r22` because no deployable behaviour changed — the new classes have
no callers, which R23-107 requires until Phase 2 migrates them. Runtime
output is identical to r22.

## 2026-07-31 - Phase 0 closed: datum-first ordinate transaction proved

### Live evidence

Both ordinate groups completed on the authorized P-0251 fixture, closing the
last outstanding Phase 0 gate.

- `AddOrdinateDimension` returned `0 Success` for both directions.
- Exact selection counts: datum + 2 = 3 (X), datum + 3 = 4 (Y). The X datum
  selected as an edge (type 1), the Y datum as a vertex (type 3).
- Display-dimension deltas +2 and +3; `SetPickMode` called and zero
  selections remaining on both exits; fixture unchanged and drawing unsaved.
- Every selection reported `ownerView=Drawing View1`.
- Ordinate values are semantically correct: `+15.00`/`-15.00` about the
  stepped-bore centre and `10.00`/`50.00`/`90.00` from the bottom-left vertex
  datum — the reference scheme of symmetric centre-zero X and bottom-zero Y.
- Five created dimensions for six locations is the designed deduplication:
  six locations collapse to two unique X and three unique Y coordinates.

Two contracts were settled for production:

- `ISelectData.View` assignment raises error 91 on this build
  (`viewBinding=UnboundAfterError:91`), and both groups completed normally
  with unbound selection data after activating the view and verifying
  ownership through `GetSelectedObjectsDrawingView2`.
- Created ordinates report `Type2 = 1` (horizontal request) and `Type2 = 7`
  (vertical request), not `7` and `8`. QA must accept `1`, `7` and `8`.

### Status

The Phase 0 gate in `docs/R23_IMPLEMENTATION_PLAN.md` is satisfied and
production Phase 1 is unblocked. R23-006, R23-009 and R23-010 are closed.
The remaining J-J top-border violations are production work for R23-704, now
measured truthfully in the page frame rather than unknown.

Production R23 source, `Fable.swp`, fixtures, the protected baseline, and
manual references remain unchanged.

## 2026-07-31 - Entity-correspondence route settled; ISelectData.View defect reproduced

### Live evidence

The Boolean normalization cleared every qualification gate and the run
exercised all three mapping routes for the first time.

- Route A, `IView.GetCorrespondingEntity(modelEdge)`, works: 12 of 24
  counterbore edges and 114 of 154 body vertices mapped with `error=0`.
- Route B, `IComponent2.GetCorrespondingEntity`, returned `Nothing` on every
  attempt and is not usable for part drawing views on this build.
- Route C confirms route A returns genuine drawing-context entities: each
  mapped entity identity-matches a visible-edge entry via `ISldWorks.IsSame`.
- Mapping is per-edge, not per-face — only one of each counterbore's two
  owned circular edges maps, so every owned edge must be attempted.
- The required scheme resolved end to end: six unique locations, two unique X
  and three unique Y, plus a stepped-bore X datum (r `0.0235`) and a
  bottom-left vertex Y datum.
- `ICurve.CircleParams` works, returning radii matching the owning cylinders
  exactly. The R23-006 exclusion was entirely an artifact of the guard defect.

Both ordinate groups then failed with runtime error 91 before any selection.
`CreateSelectData` returned a live object, so the failing statement is
`Set selectData.View = swView` — reproducing the behaviour already recorded
in `Module2_DrawingPipeline.CreatePrimarySection`. The MCP documents the
property as `get; set`, so this is an installed-build deviation, not misuse.

### Changed

- The probe now attempts the `ISelectData.View` binding inside a guarded
  helper and records `viewBinding=Bound` or `UnboundAfterError:91`, continuing
  with unbound selection data when it fails. Each selection's owning view is
  proved through `ISelectionMgr.GetSelectedObjectsDrawingView2`, and every
  group record carries the exact `lastStep` reached.

`AddOrdinateDimension` has still never been called. Production R23 source,
`Fable.swp`, fixtures, the protected baseline, and manual references remain
unchanged.

## 2026-07-31 - SOLIDWORKS Boolean contract established; CircleParams exclusion withdrawn

### Live evidence

The `rejectGate` instrumentation isolated the mechanism exactly. One call
produced both `isCircle=True` and `rejectGate=IsCircleFalse`: `CStr` rendered
the value `True` while `If Not value` fired on the same variable.

Established contract for installed SOLIDWORKS 2025 SP1.2 COM Booleans in this
VBA host: `If value Then`, `If value = False Then` and `CStr(value)` are
safe; `If Not value Then` is not, yielding `-2`, which VBA treats as True.
`CBool` is not a dependable normalization — `CBool(rawVariant)` worked for
`ISurface.IsCylinder` while `CBool(curve.IsCircle)` did not. Only an explicit
numeric comparison, `(CDbl(rawValue) <> 0#)`, is representation independent.

### Changed

- Added `NormalizeSwBoolean` to the disposable probe and applied it to
  `ICurve.IsCircle`, `ISurface.IsCylinder`, `IDrawingDoc.ActivateView`, both
  `IEntity.Select4` calls, and `IModelDoc2.GetSaveFlag`. Every remaining
  `If Not` in the module is an object `Is Nothing` test, a VBA built-in, or a
  normalized value. Raw returned values are logged alongside normalized ones.

### Withdrawn

- The `ICurve.CircleParams` exclusion in R23-006 is withdrawn as unproved.
  The `SkippedNotCircle` result came from
  `Module_R23Phase0FeatureProbe.ReadCircleState` line 611, which guards the
  call with `If Not isCircle Then` — the same defect. `CircleParams` was
  never invoked and no anomaly was ever observed. Its behaviour remains
  untested; the corrected probe now reads it as non-load-bearing evidence so
  the next run settles the question. Production still must not depend on it
  until a run exercises it.

## 2026-07-31 - Counterbore ownership proved; edge-closure gate corrected

### Live evidence

The Boolean-normalization fix worked. Twelve of the eighteen faces owned by
`CBORE for M6 Socket Head Cap Screw1` now read as cylinders with correct
radii (`0.0055` and `0.0033`), all passed `TransformPointToView` with
outline-contained page proof, and the six page centres form exactly the
required grid: two unique X (`0.080932`, `0.110932`) and three unique Y
(`0.077060`, `0.117060`, `0.157060`), spaced 30 mm and 40 mm to match the
drawing's `30.00` and `40.00`.

Feature ownership, cylinder qualification and the model-to-page transform are
proved for all six counterbore locations.

### Fixed

- Corrected the edge-closure gate. Every edge in the run returned
  `completeCircle=False` — all 24 feature-owned and all 64 visible drawing
  edges — including about thirty reporting `isCircle=True`, `uMin=0`,
  `uMax=6.283185307` and a zero endpoint gap, which satisfy every gate. With
  `GEOMETRY_TOLERANCE_M` at `0.0000001`, the success assignment should have
  been reached; the only construct able to swallow it was the single-line
  `If ... Then _` continuation immediately preceding it.
  `ReadEdgeCircleEvidence` now uses block `If` statements, reports the exact
  `rejectGate` per edge, and logs the `closureToleranceM` in force. No
  `Then _` construct remains in the module.

Mapping routes A, B and C have still never executed. Production R23 source,
`Fable.swp`, fixtures, the protected baseline, and manual references remain
unchanged.

## 2026-07-31 - Corrected-probe live results and Boolean-normalization fix

### Live evidence

The user compiled and ran probe build `20260731.2` on P-0251. The section and
J-J gates closed; the ordinate gate did not.

- Section targets are exact and non-stale: one `DIAMETER_47`, one
  `DIAMETER_40`, one `FIRST_LINEAR`, `exactTargetCounts=True`, both diameters
  `type2=6`.
- Direct part-source readback proved H7 is **absent** from `D1@Sketch4`
  (`toleranceType=0`, `fitType=-1`). The user selected controlled
  target-spec/reference authority, closing R23-806: R23 will apply
  `H7 +0.025/0.000` to an associative `Ø47` dimension with provenance
  recorded as target-spec authority and stated as such in QA.
- The J-J payload segment frame is proved, not inferred: all six endpoints
  matched the captured `CreateLine` view-sketch inputs at `deltaM=0`. The
  single inverse conversion is cross-validated by the independently supplied
  page-frame arrow endpoints, which agree exactly.
- Page-frame clearance measured three violations, all at the top:
  segment 1 start, upper arrow, and upper label exceed the content-border top
  of `0.287`. The lower arrow and label **clear** the measured
  part-identification note extent by about 7.7 mm, correcting the earlier
  screenshot-derived claim that they intruded into that band.

### Fixed

- Corrected a Boolean-normalization defect introduced in the previous probe
  revision: `If Not surface.IsCylinder Then` rejected all 18 owned counterbore
  faces because a raw SOLIDWORKS `VARIANT_BOOL` of `1` makes `Not value`
  evaluate to `-2`, which VBA treats as True. The probe now uses
  `CBool(ISurface.IsCylinder)`, matching the accepted feature probe, and logs
  the raw returned value. The same normalization was applied to
  `IDrawingDoc.ActivateView` on the ordinate transaction path, and the visible
  drawing-edge inventory now logs its curve-parameter fields.

The three mapping routes remain unexercised; the next run is the first that
can test them. Production R23 source, `Fable.swp`, fixtures, the protected
baseline, and manual references remain unchanged.

## 2026-07-31 - Corrected disposable Phase 0 drawing-contract probes

### Changed

- Corrected the disposable ordinate probe in
  `tools/r23-probes/import-transaction-source/Module_R23Phase0DrawingProbes.bas`
  (build `20260731.2-mapping-frame-h7-contracts`): every owned face and edge
  of the target feature now logs a qualification record, and each complete
  circular edge compares the direct active-part view mapping, the
  component-mediated mapping, and the `GetVisibleEntities2` inventory with
  `ISldWorks.IsSame` identity results. Ownership, cylindrical-radius,
  `IsCircle`, and closed `GetCurveParams3` qualification is unchanged, and
  `AddOrdinateDimension` remains gated on six counterbore locations
  resolving to two unique X and three unique Y coordinates.
- Corrected the disposable section probe: every per-dimension diagnostic
  value is reset at the top of each display-dimension iteration; diameter
  targets require the live-proven `swDiameterDimension = 6` type plus the
  exact nominal, with exact-count enforcement and `*_DUPLICATE` labelling;
  and the original part-source `D1@Sketch4`/`D1@Sketch6` tolerance and fit
  data are read directly through `IModelDoc2.Parameter`, feeding an
  explicit `R23_H7_AUTHORITY` record. The probe never invents H7.
- Corrected the disposable J-J diagnostic: the Module2 overlay captures the
  original page-frame construction points and the converted view-sketch
  `CreateLine` inputs; the parsed payload's segment frame is proved by
  point-by-point comparison; segment endpoints are converted to the current
  page frame exactly once via the exact inverse of the
  `IView.GetXform`/`IView.Angle` conversion; and page-frame clearance
  verdicts are logged against the content border, title block, measured
  part-identification note extent, and view outlines.
- Updated `tools/r23-probes/README.md` with the corrected-probe record
  types and rerun evidence expectations, and recorded the session's MCP
  contract lookups in `docs/SOLIDWORKS_API_VALIDATION.md`.

### Verification

Static verification only: structural source checks pass, both edited
sources remain ANSI/CRLF without BOM or `Attribute` metadata, the complete
offline suite reproduces the known 69-pass/5-stale-R20-failure baseline,
and the disposable drawing-contract manifest passes read-only preflight
against the copied probe SWP. The corrected probes have not been deployed,
compiled, or executed; production R23 source, `Fable.swp`, fixtures, the
protected baseline, and manual references are unchanged.

## 2026-07-31 - R23 Claude Code handoff after final Phase 0 probes

### Added

- Added `docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md` as the durable
  same-workspace handoff for Claude Code. It records the goal, source of truth,
  protected assets, accepted evidence, unresolved gates, R23 architecture, and
  exact next work package.
- Added `docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md` as the evidence index
  and disposable-probe correction specification.
- Preserved the two uncropped final-probe screenshots in the timestamped R23
  `live-probes` evidence folder without modifying their originals.

### Findings

- The ordinate probe stopped before `AddOrdinateDimension` because
  feature-owned counterbore edges did not survive the opaque mapping helper.
  The corrected diagnostic must log edge qualification and compare direct,
  component-mediated, and `GetVisibleEntities2` mapping paths.
- The section probe found the exact 47 mm and 40 mm imports as
  `swDiameterDimension = 6`, with no imported H7 or nonzero tolerance. It also
  retained stale per-item state across loop iterations, so only the exact
  identity records are accepted from that transcript.
- The J-J payload is structurally complete, but its segment records and
  arrow/label records use different coordinate frames. The screenshot confirms
  both top-border and lower part-identification-band collisions.

### Scope

This handoff checkpoint changes documentation and retained diagnostic evidence
only. Production R23 VBA, `Fable.swp`, the protected baseline, fixture models,
and manual reference drawings remain unchanged.

## 2026-07-30 - R23 Phase 0 implementation checkpoint

R23 implementation has started at the plan's mandatory runtime-proof gate.
This checkpoint does not change the production VBA revision or drawing
behavior.

### Preserved baseline

- Backed up `Fable.swp` before any material embedded-project change and proved
  its SHA-256 matches the production file.
- Verified 15/15 managed components and the exact R22 revision between the
  embedded project and exported source.
- Recorded the pre-change repository, SWP, and managed-source hashes under
  `test_assets/iteration_evidence/r23/20260730-075811/prechange/`.
- Re-ran the guarded deployment preflight successfully.

### Phase 0 preparation

- Added `tools/r23-probes/Module_R23Phase0FeatureProbe.bas`, a standalone,
  read-only probe for raw/effective feature types, configuration-specific
  suppression, typed Hole Wizard/extrude definitions, owned cylindrical
  geometry, pattern seed evidence, and both circular-curve read orders.
- The probe refuses any model outside the three authorized fixtures and pairs
  every successful `AccessSelections` with `ReleaseSelectionAccess`.
- Added usage and operator-gate guidance in `tools/r23-probes/README.md`.
- Updated the implementation plan with completed provenance/probe-preparation
  tasks and corrected its drawing-component context contract.

### API evidence

- Re-queried the local `solidworks-api` compatibility MCP and reflected
  installed SOLIDWORKS 2025 SP1.2 interop `33.1.2.4`.
- Confirmed `ICE -> GetTypeName`, `GetDefinition`, feature-owned faces,
  typed selection-access rollback/release, `ICurveParamData`, and the
  circle-parameter contracts used by the probe.
- Corrected the plan's earlier description of `GetVisibleComponents`: its
  returned component is limited and is suitable for `GetVisibleEntities2`;
  full component behavior requires
  `GetVisibleDrawingComponents -> IDrawingComponent.Component`.

Production R23 source remains intentionally unchanged until the required live
Phase 0 evidence exists.

### Phase 0 probe correction

- Preserved the fail-closed refusal from the first attempted run, which found
  an unauthorized `V:\VEEMAP\SW_data\...` document active.
- Preserved the first authorized-fixture transcript but rejected it as
  incomplete evidence: an `ObjPtr`-only recursion key visited 15 features and
  skipped the relevant ICE, Hole Wizard, extrusion, and mirror entries.
- Replaced the transient-wrapper-only recursion key with a composite feature
  diagnostic key.
- Updated suppression readback to accept and label both scalar and array
  `IsSuppressed2` return shapes. The installed VBA run returned a scalar
  Boolean even though older source assumed only an array.
- Added base `EXTRUSION` to the typed `IExtrudeFeatureData2` probe routes.
- Added extrusion contour count/state and owned profile-subfeature readback so
  the cut probe satisfies the plan's sketch/contour evidence requirement.
- Mirrored every feature-probe record to a timestamped evidence log to prevent
  Immediate Window buffer limits from truncating the contract proof.
- Prepared a copied `Fable.swp`, a three-module disposable overlay, and a
  guarded custom deployment manifest for the expanded import comparison. The
  two entry points create separate fresh drawings for selected-primary
  `AllViews=True` and deterministic section/side/primary `AllViews=False`,
  always with `DuplicateDims=True` and mask `18055274`.
- Added owner, source-identity, nominal, tolerance, fit, callout-variable,
  attachment, duplicate, view-delta, and raw display-geometry logging for the
  import variants. The custom manifest passed read-only deployment preflight;
  no VBA full-project compile or live import run is claimed.
- Kept the production R22 source and `Fable.swp` unchanged. The corrected
  disposable probe still requires user-operated compilation and runtime
  evidence before any production R23 classification work begins.

### Phase 0 feature evidence â€” 2026-07-31

- Compiled and ran the corrected read-only feature probe on the authorized
  P-0251 fixture. The retained `R23_FEATURE_20260731_040539.log` records 47
  visited features, all three `ICE -> Cut` resolutions, typed definition
  access/release, owned geometry, native CBORE and M5 Hole Wizard data, and
  `Mirror1` seed ownership.
- The configuration-specific `IsSuppressed2` result was Empty for every
  feature. The probe therefore records the separately labelled active-document
  `IsSuppressed` fallback; this is live evidence for the active `Defualt`
  configuration only, not a substitute for a drawing-view configuration proof.
- Both curve sequences retained stable `IsCircle=True` and complete
  `GetCurveParams3` evidence. Because the `CircleParams` helper still produced
  `SkippedNotCircle` despite true before/after predicates, R23 will not depend
  on `CircleParams`; it will qualify circles from the live-proven cylindrical
  face, `IsCircle`, and complete-boundary parameter evidence.
- Production R23 source and `Fable.swp` remain unchanged. The next live gate is
  the disposable expanded-import transaction comparison.

### Phase 0 import evidence â€” 2026-07-31

- Ran both expanded-mask import variants in fresh disposable P-0251 drawings.
  Each completed successfully with 25 unique imported source identities, no
  fixture save, and a clean isometric view.
- Selected `AllViews=False` as the R23 import transaction: explicit
  section/side/primary selection yielded 17/2/6 dimensions. The selected
  anchor with `AllViews=True` yielded 17/0/8, so it cannot satisfy the required
  side-view coverage.
- Proved native M5 Hole Wizard callout import and its tap-drill, thread, and
  depth variables. The only native callout landed in the section. Neither
  variant imported the required six-hole M6 counterbore callout, H7 fit, or a
  nonzero tolerance; R23 must therefore use the approved controlled fallback
  only when typed feature/model data proves its content.
- The generated probe sheets visually confirm that imported dimensions can
  overlap the section label/title area and do not create ordinate coverage.
  Production layout and explicit ordinate stages remain mandatory.

### Final Phase 0 probe preparation â€” 2026-07-31

- Added the disposable `Module_R23Phase0DrawingProbes.bas` overlay with public
  entry points for datum-first X/Y ordinate creation and targeted section
  dimension/J-J geometry evidence.
- The ordinate probe maps feature-owned complete circles without depending on
  `CircleParams`, requires six M6 locations resolved to two X and three Y
  coordinates, selects the datum first, appends each entity individually,
  decodes the installed ordinate result enum, and records `SetPickMode` plus
  zero-selection cleanup.
- The section probe records every imported section dimension, targeted 47/40
  tolerance/fit and display data, `IDrSection` state, and bounded parsing of
  J-J segments, arrows, and label positions.
- Added a guarded four-module deployment manifest targeting only the existing
  disposable `R23_Phase0ImportProbe.swp` and documented the exact user
  compile/run sequence. Production source, production `Fable.swp`, fixture
  models, the protected baseline, and manual references remain unchanged.
- Tied H7 proof specifically to the imported 47 mm dimension so an unrelated
  fit cannot satisfy the section contract, and added an exact reported-versus-
  flattened section-line count check before parsing any geometry.
- Reflected the installed SOLIDWORKS 2025 SP1.2 interop for the ordinate,
  selection, dimension/tolerance, display-data, `IDrSection`, and section-line
  members used by the probes.
- The four-source disposable manifest passed guarded read-only preflight.
  Source hygiene and structural checks passed for all overlay modules. The
  complete offline suite reproduced its known baseline: 74 tests run, 69
  passed, and only the five documented stale R20 contract failures remained.
- No full-project VBA compile or installed-build execution is claimed; those
  two gates remain user-operated.

## 2026-07-29 - R22 verified review resolution

Source identity moves to `target-spec-hybrid-v2-2026-07-29-r22`. R22 contains
the complete r20/r21 line, retains the corrections from `1698b7d` that satisfy
their contracts, and replaces the latest commit's unsafe or incomplete parts.

### Correctness

- `Module5`: removed the `ISurface.CylinderParams` trimming-edge fallback.
  `ICurve.IsCircle=False` now fails as `ClosedCircleCurveNotCircular`; an
  internal cylindrical face no longer proves that a closed trim is a circle.
  This prevents oblique elliptical trims from being assigned a false centre.
  The retained P-0251 contract probe shows every relevant owned edge already
  supplies `IsCircle=True` and seven-value `CircleParams`.
- `Module3`: replaced guessed feature aliases with the exact 2025
  `GetTypeName2` literals: `APattern`, `LocalChainPattern`, `DimPattern`,
  `DerivedHolePattern`, `SketchPattern`, and `LocalSketchPattern`, while
  retaining the already-correct pattern and mirror names.
- `Module4`: removed the module-level per-side ordinate-lane cache. If
  `AlignDimensions` returns False, ordinate types 1, 7, 8, and 16 now call and
  check `IDisplayDimension.AutoJogOrdinate`, then prove their readback
  positions. Deterministic `SetPosition2` lanes remain only for non-ordinate
  dimensions. This avoids both merging independent chains and splitting one
  type-1 chain by member proximity.

### Retained review fixes

R22 retains r21's projected-origin containment correction, pattern-seed
routing, FACE/SIDE callout separation, annotation-name self-identity fallback,
controlled static general-note proof, section-line parser robustness,
selection-readback tolerance, title-border window, redundant-check removal,
and reduced rejected-vertex logging.

### Verification

- Queried the local `solidworks-api` MCP for the load-bearing circle, feature,
  seed, ordinate, annotation, and enum contracts.
- Reflected the installed SOLIDWORKS 2025 SP1.2 interop assemblies, version
  `33.1.2.4`, for all newly used members and `swDimensionType_e`.
- Updated source/docs and prepared the guarded r22 SWP deployment.
- Full VBA Editor compilation, fixture runtime, QA, and visual/manufacturing
  acceptance remain separate gates.

See `docs/R22_REVIEW_RESOLUTION.md`.

## 2026-07-29 - R21 code-review remediation

Source identity moves to `target-spec-hybrid-v2-2026-07-29-r21`. R21 is r20
plus fixes for defects found in a full review of the r20 diff. No new features.

### Correctness

- `Module5`: the `ISurface.CylinderParams` fallback used the cylinder axis
  origin as the hole centre. The 2025 docs define that array as
  `origin/axis/radius` where origin is an arbitrary point on the axis, so an
  oblique bore was dimensioned at the wrong place and both circular edges of one
  cylinder collapsed to the same centre. The origin is now projected along the
  axis into the plane of the edge using a point known to lie on that edge.
- `Module5`/`Module8`: `TransformPointToView` always required the transformed
  point to lie inside `IView.GetOutline`. The model origin routinely projects
  off the solid, so `ProveCenterDatum` failed before examining any vertex and
  the Center datum became unprovable. Containment is now an explicit argument:
  required for hole centres and mapped vertices, not for the projected origin.
- `Module4`: the deterministic lane fallback allocated a new lane per
  dimension, giving an ordinate chain one baseline per member. Ordinates on a
  side now share one lane. `swOrdinateDimension` (1) - documented as "base
  ordinate and its subordinates" and already counted as an ordinate by
  `Module6` - was absent from the type switch and fell through to proximity
  routing, splitting a chain across two sides; types 1 and 16 are now handled.
- `Module3`/`Module5`: hole-seed resolution was wired only for `MirrorPattern`
  although `IFace2::GetSeedFeature` covers patterned, mirrored and copied
  bodies. Linear, circular, curve, table and fill patterns produced no
  candidate and no rejection record. All pattern families now share one proof.
- `Module7`: the FACE and SIDE manufacturing callouts shared one target
  position, so two callouts resolving to the same view were placed on top of
  each other and the collision check then failed the stage. SIDE has its own
  lane.
- `Module7`: the callout was excluded from its own collision scan using `Is`
  against a re-fetched `GetAnnotation` pointer, which may be a different
  wrapper for the same annotation - the note could collide with itself.
  Identity now also matches on annotation name.
- `Module7`: general-notes verification compared against a value read from a
  part custom property, but the controlled format carries that note as static
  text no property write can change. Any part carrying a `GeneralNotes`
  property made the stage permanently unprovable. The controlled reference
  constant is now also accepted, and whichever text proved the note is what the
  containment check uses.

### Robustness and diagnostics

- `Module6`: removed the exact-equality assertion between the section-line
  array length and the `Size` out-param of `GetSectionLineCount2`. The API docs
  describe `Size` only as "size, which includes an extra double per section
  line containing the layer ID" and never define it as the raw array length, so
  a convention difference would reject every drawing. The per-element cursor
  guards already bound every read; the pair is now recorded, not enforced.
- `Module6`: every malformed-array exit in `ValidateSectionLineInfo` now
  reports what failed and at which index. Previously they returned silently and
  the caller reused the same stage message as a real clearance violation.
- `Module4`: selection readback no longer has to equal the `Select3` success
  count. Re-selecting an already-selected annotation returns True without
  growing the list, so one duplicate let a cosmetic arrange step reject the
  whole run. An empty selection still fails; a shortfall is warned.
- `Module8`: widened the legacy title-block bottom-edge window, which accepted
  only 8.9-17.8 mm on A3 and excluded any format drawn to a 20 mm border.
  Replaying the recorded P-0251 template sketch gives byte-identical bounds
  before and after.
- `Module8`: replaced two rectangle-contract checks that compared the measured
  bounds against the same limits the candidate filter had already enforced, so
  they could never fire.
- `Module8`: removed `ViewToSheetCoordinates`, reduced in r20 to an identity
  copy with no callers, which would have silently mis-placed any future caller
  holding genuinely view-local coordinates.
- `Module5`: the per-vertex transform proof is logged only on rejection. It ran
  twice per edge over every face of every visible component.

### Verified statically only

API contracts were confirmed against SOLIDWORKS 2025 documentation through the
project `solidworks-api` MCP. The `GetSectionLineInfo2` cursor strides
(7/9/9/7) were checked against the documented layout and are correct as r20
wrote them; no change was made there. Procedure-block balance, duplicate `Dim`,
and signature/call-site arity were script-checked across every module.
**R21 has not been compiled in the VBA editor or run against a fixture.**
Full-project **Compile Project** and a focused P-0251 run remain mandatory.

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
