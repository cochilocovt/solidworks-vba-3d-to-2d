# Current Status

Date: 2026-08-01

## R23 Phase 6 source complete, awaiting first live run

**2026-08-01.** `Module16_CalloutDefinition.bas` (20 procedures) plus
`CCalloutDefinition.cls`. Statically verified only.

A callout definition is either **native** - a SOLIDWORKS hole callout
carrying the Hole Wizard's own data - or a **controlled fallback** built
field by field from typed feature data. Never free text. Every field carries
the source that proved it, so a definition that looks complete can still be
shown to be unproven.

Design points worth carrying:

- **`IsHoleCallout` is the only classifier.** A native callout reports
  `Type2 = 6`, but so does an ordinary diameter dimension. No
  dimension-type constant is declared in the module, so none can be reached
  for.
- **Fields come from `GetHoleCalloutVariables`**, not from parsed text -
  `HoleFit`, `ShaftFit`, `ToleranceType`, `ToleranceMin`, `ToleranceMax`
  per variable. A rendered string cannot be validated field by field.
- **Quantity is unique physical locations.** Not features: one Hole Wizard
  feature plus a mirror makes many holes. Not edges: a counterbore
  contributes several per hole.
- **A callout resolving to two families is rejected**, not tie-broken.
- **Depth is required only when the end condition says the hole is blind**,
  and an unproven end condition fails on its own terms first.
- **R23-611 is stated as shapes, not part numbers**: one multi-hole
  counterbored family and one multi-hole threaded family. P-0251 satisfies
  it; nothing is keyed to it.

**R23-609 is half met, deliberately.** The new path has none of the
hardcoded text or name/radius scoring, and contracts assert their absence.
The legacy literals remain in `Module7_TitleBlockEngine.bas` (callout text
at 359-371, scoring at 405-435) because Module7 is still the reachable
production path and Module16 is not yet wired into `main`. Removing them now
would degrade the deployable macro while the replacement is disconnected.

Verification: 23 Phase 6 contracts, suite 238 tests with the same five stale
R20 failures. Preflight 29 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 5 read-only gate SATISFIED

**2026-08-01.** `schemes=4|horizontalSchemes=2|verticalSchemes=2|`
`creditedLocations=10|expectedLocations=10|coverageFailures=None`.
Read-only throughout: `creations=0`, `initialSelectionCount=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

Proven live: R23-500 (four measured schemes), R23-501 (stepped-bore centre
datum, selection and ownership proved), R23-503 (two X buckets), R23-504
(three Y buckets), R23-505 (all four side holes credited across two page
positions), R23-507 (`profileEntries=1`, stepped bore held out by family
size not radius), R23-509 (10 of 10).

**Still open, and stated as open:**

- **R23-502 is NOT met.** The vertical datum is the lowest projected hole,
  recorded `datumKind=ProjectionDerived`. The task asks for the part's
  bottom outline geometry. Not claimed on a weaker datum.
- **R23-506 is half met.** All four side locations are resolved and
  credited; none is dimensioned yet.
- **R23-508 is unrun.** It creates dimensions, so it needs authorization and
  a target that is not the manual reference drawing.

Two defects were found by the run rather than by me. The QA gate popped
`RESULT: FAIL` because the probe called `Module6_QAEngine.EmitRunEvidence`,
the production gate demanding fourteen pipeline stages a probe never runs.
Then the coverage gate reported `credited=8, expected=10`: I read
`CoincidentWithAnchoredKey` from the anchored end, where
`MarkCoincidentProjections` never sets it, and the unanchored twin holding
it had already been filtered out. Two of P-0251's four side holes were
silently uncredited. Both fixed, both now pinned by contracts.

Verification: 25 Phase 5 contracts, suite 215 tests with the same five stale
R20 failures. Preflight 27 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 5 source complete

**2026-08-01.** `Module15_OrdinateScheme.bas` (36 procedures) plus typed
records `COrdinateScheme.cls` and `COrdinateBucket.cls`. Statically verified
only — nothing below is runtime-proven.

The scheme key replaces feature-family grouping with **view role + machining
face + datum policy + direction**, every part measured rather than read off a
name: machining face from the location's sign-normalized axis, view role from
Phase 3's axis-normal measurement via the existing eligibility tests.

Two Phase 3 findings are carried into the design rather than worked around:

- **Coverage is counted per distinct page position, credited to locations.**
  A bucket holds one selectable entity and the list of physical locations it
  represents. Coaxial holes collapse to one drawing entity, so demanding one
  dimension per location is unsatisfiable by construction; crediting only one
  of the pair would silently drop the other.
- **Small-hole membership is family size, not a radius threshold.** P-0251's
  stepped bore is excluded because it is a singleton family. A magic
  millimetre value would misclassify a different part.

Three defects caught before compiling: `IsOrdinateEligibleView` and
`IsDeferredCreationView` take `(graph, swView)`, not `(swView, graph)`; and
`IView.GetFirstDisplayDimension5` is obsolete **and** sheet-scoped by its own
Remarks, so a read-back built on it would credit other views' dimensions to
this scheme. Read-back now uses view-scoped `IView.GetDisplayDimensions` with
a before/after snapshot diffed by `ISldWorks.IsSame`.

Mutation boundary: `CreateOrdinateGroup` alone creates anything and refuses
without `allowMutation`; it also refuses when the datum is unproven.
`R23_ProbeOrdinateScheme` contains no `AddOrdinateDimension` call at all.

Verification: 22 Phase 5 contracts, suite 212 tests with the same five stale
R20 failures. Preflight 27 managed components. `MACRO_SOURCE_REVISION`
remains `r22`.

## Historical: R23 Phase 4 gate SATISFIED (sixth live run)

**2026-08-01.** Read-only throughout: `mutations=0`, `initialSelectionCount=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

`annotations=38`, `coverageFailures=None`,
`COVERAGE|holeCallouts=2|ordinates=10|diameters=0|toleranced=1|withFit=1`.
R23-412 is required-**category** coverage and all three required categories
are present. Gate met.

**The instrumentation settled the reconciliation question rather than
improving the number.** `IModelDocExtension.GetCorrespondingEntity2` returned
Nothing for all 38 annotations and every attachment **with error 0** —
`outcomes=draw1:unresolved:err0`, `resolved=0`, `eqMax=-1` (no comparison ever
ran, so the `swObjectEquality` Unsupported hypothesis is ruled out too). The
call declines, it does not fail. This is a part drawing and the member
resolves into an underlying part or subassembly, consistent with the
`componentContext=DrawingContextOnly` already recorded.

`reconciled=1` therefore stands, and is correct. `RD1@Drawing View7` proves
the R23-407 identity mechanism end to end. The other 37 are hand-authored
reference dimensions with no reachable model counterpart; they now report
`AuthoredDrawingEntityNoModelCounterpart` — a fact about the drawing, not a
defect in the ownership model. Reconciling R23's own imported annotations is
Phase 5+ and is unaffected.

Verification: 190 offline tests, same five stale R20 failures, 36 Phase 4
contracts. Preflight 24 components. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: Phase 4 fifth live run — reverse route matched nothing

**2026-08-01.** The reverse-correspondence route executed on all 38
annotations and reconciled none. `reconciled` is still 1 of 38 — the same
`ForwardAlias` match on `RD1@Drawing View7`. Read-only boundary held:
`mutations=0`, `finalSelectionCount=0`, `drawingUnchanged=True`.

**The run did not say why, and that is the defect being fixed here.** Two
causes were equally consistent with the log: `GetCorrespondingEntity2`
returning Nothing, or resolving to an entity no location owned. Rather than
predict, `MatchByReverseCorrespondence` now reports per attachment the
drawing entity type, whether the reverse call resolved, the trapped error
number, and the resolved model entity type — plus `resolved`,
`projectionsInView`, `modelEdgesTested` and `eqMax`.

`eqMax` carries the raw `swObjectEquality`. **`ISldWorks.IsSame` returns
0 NotSame / 1 Same / 2 Unsupported, and the `ObjectsAreSame` wrapper
collapses 0 and 2.** A cross-document comparison that cannot be performed
reads exactly like a non-match — a live suspect for a silent zero.

Two unverified observations to test against the next run: `Section View J-J`
is built on a section assembly (`P-0251-14A-001-SectionAssembly-3-1/...`)
whose cut edges may have no part counterpart at all; and several ordinate
attachments genuinely read as type `0` (`swSelNOTHING`), which is not valid
input to `GetCorrespondingEntity2`.

Verification: 189 offline tests, same five stale R20 failures, 35 Phase 4
contracts. Preflight 24 components. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 fourth run: reconciliation cause found, route added

**2026-08-01.** The instrumentation isolated the reconciliation failure in
one run, and it was a design gap, not a coding slip.

**The counterbore hole callout attaches to a drawing edge
(`attachmentTypes=1`, `swSelEDGES`) that is none of the 18 aliases
`IView.GetCorrespondingEntity` produced for that view.** The forward
model-to-drawing map is partial — it returned 2 of each location's 4 boundary
edges — so matching by forward alias could never find the callout's edge no
matter how many aliases were compared.

`Section View J-J` reported `anchoredProjections=0|aliasesAvailable=0`: the
forward map produces nothing there at all, so its seven dimensions,
including the H7, were structurally unreachable.

**R23-302 asked for the other direction and I had only built the forward
one.** `IModelDocExtension.GetCorrespondingEntity2` maps a drawing entity
back to its model entity, which the 2025 Help names for exactly this
purpose. `MatchByReverseCorrespondence` now takes each attached drawing
entity, resolves it to the model, and tests it against the geometry each
physical location owns — its `SourceFaces` and those faces' boundary edges.

The reverse route needs no projection anchor, so it can reach the section
view. Still identity only: `ISldWorks.IsSame` throughout, no positional or
dimensional fallback anywhere. Matches record `matchRoute=ForwardAlias` or
`ReverseCorrespondence` so the two are never conflated.

Read-only boundary held again: `mutations=0`, `drawingUnchanged=True`.

Verification: 187 offline tests, same five stale R20 failures, 33 Phase 4
contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 third live run

**2026-08-01.** Read-only boundary held again: `mutations=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

### Tolerance policy is in force

Per the user's standing instruction, every tolerance now reports
`toleranceAuthority`. The H7 reads
`DrawingAuthoredNonAuthoritative`; all others `NoTolerance`. The dowel ±10 µm
rule is gone from the source entirely, and `ClassifyToleranceAuthority` has
no branch that could return a stronger authority. **Tolerances in the
designers' existing drawings are evidence that a designer typed a number, not
statements that the part holds them.** The rule for when a tolerance *should*
be added is open, pending the user's discussion with their designer.

### Reconciliation is still 1 of 38, and the alias fix did not help

Matching every drawing entity the projection owns — not just the anchor — was
predicted to raise this to 3-4. It did not move. Rather than guess again, the
evidence now names what each attachment actually is: `ANNOTATION_UNMATCHED`
reports `attachmentTypes` as `swSelectType_e` codes (1 edge, 2 face,
3 vertex, 10 sketch segment, 11 sketch point, 28 centre mark, 46 silhouette)
plus `anchoredProjections` and `aliasesAvailable`, so the next run says
whether the counterbore callout attaches to an unmapped edge, a sketch
segment, or something that is not an edge at all.

Only `RD1@Drawing View7` (the 4x tap-drill callout) reconciles. Everything
else in the drawing is authored as reference dimensions whose attachments
have not yet been shown to be hole edges.

Verification: 184 offline tests, same five stale R20 failures, 30 Phase 4
contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 second live run

**2026-08-01.** Both fixes confirmed, and the read-only boundary held again:
`mutations=0`, `finalSelectionCount=0`, `drawingUnchanged=True`.

**`IDimension.SystemValue` is definitively the right member for drawing
dimensions.** Every line now reads
`nominalSource=IDimension.SystemValue=<real value>|GetSystemValue3=0.000000000`
— all 30 display dimensions. Values match the reference drawing exactly:
47.0 H7, 11.0 counterbore callout, 4.2 tap-drill callout, R36, 173.6, 104.8,
80, 40, 25, 18, 12, 11.5, 6, and the ordinate chains 10/50/90/160 and
15/36 about zero. `attachments` is now 1-3 rather than 0.

The H7 reads in full: nominal `0.047000000`, `toleranceType=8`,
`tolMaxM=0.000025000`, `tolMinM=0.000000000`, `holeFit=H7`, status `0/0` —
the drawing's `47 H7 +0.025/0` exactly, and absent from the model.

### Standing instruction: drawing tolerances are not authoritative

The user directed that the tolerances in the designers' existing drawings be
ignored: they were added manually to signal that *some* tolerance is
acceptable, not that the part holds them. Every tolerance R23 reads is now
labelled `toleranceAuthority=DrawingAuthoredNonAuthoritative`, and
`ClassifyToleranceAuthority` has **no branch that can return a stronger
authority** — nothing R23 can currently read distinguishes a binding
tolerance from an indicative one. The dowel-specific ±10 µm rule is removed
rather than parked; the user is establishing with their designer what part
information should drive the decision to add a tolerance, and no rule is
guessed at meanwhile.

### Reconciliation defect fixed

Only 1 of 38 annotations reconciled. The cause: a counterbore maps **two**
edges, and the native hole callout attaches to the 11 mm mouth while the
anchor tier deliberately prefers the 6.6 mm through hole — different
entities of the same location. Matching now tests every drawing entity the
projection owns (`ProjectionOwnsEntity`), not just the chosen anchor.
Identity only; still no positional fallback.

Also caught before compiling: `alias` is a VBA reserved word.

Verification: 183 offline tests, same five stale R20 failures, 29 Phase 4
contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 first live run

**2026-08-01.** The read-only reconciliation run held its boundary exactly:
**`mutations=0`, `finalSelectionCount=0`, `drawingUnchanged=True`** — the
manual reference drawing was not altered. 38 annotations inventoried across
four views, `coverageFailures=None`.

### The H7 question is answered

`Section View J-J` carries `RD4@Drawing View6@P-0251-14A-001.Drawing` with
`toleranceType=8` (`swTolFITWITHTOL`), `nonZeroTolerance=True`,
`fitData=True`. **The Ø47 H7 fit is authored in the DRAWING, not the model** —
Phase 0 read `D1@Sketch4` directly and found no H7 there. R23 must never
expect model import to supply it, and must never record it with model
provenance.

The user also supplied the matching domain rule: **a dowel hole receives a
designer-added ±10 µm tolerance in the 2D drawing**, never in the model. Both
facts are the same pattern — precision tolerance is a drafting-stage act on
this drawing set. `MatchesDowelToleranceConvention` recognizes the ±0.010 mm
signature so provenance can follow from it, but recognizing the pattern is
deliberately *not* treated as proof that a hole is a dowel.

### Both hole callouts found

`RD3@Drawing View4` (the 6x counterbore) and `RD1@Drawing View7` (the 4x
tapped) both classified `NativeHoleCallout`, and the second reconciled to its
physical location by attached-entity identity. Ordinates: 10.

### Two defects fixed

- **`nominalM` was 0 for every dimension.** `GetSystemValue3` is
  configuration-scoped and these are *drawing* dimensions; a drawing document
  has no configurations. Now read through the configuration-free
  `IDimension.SystemValue`, with both readings kept visible.
- **`attachments=0` on every line**, including the annotation that went on to
  reconcile — the count was populated in a later pass than the line that
  printed it. Attachments are now read during inventory.

Also: the reference drawing authors its section diameters, H7 included, as
`swLinearDimension` (2), so dimension type alone can never decide whether
something is a diameter.

Verification: 181 offline tests with the same five stale R20 failures, 27 of
them Phase 4 contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 source-complete

**2026-08-01.** `Module14_AnnotationImport.bas` implements R23-400 to
R23-412: import constants, view-eligibility policy, the Phase 0 import
strategy, independent annotation traversal, category classification,
tolerance and fit reading, identity-based reconciliation against the location
graph, deduplication, and category-based coverage.

**This is the first phase that can modify a drawing.** Only
`ImportModelAnnotations` and `RemoveR23CreatedAnnotations` mutate, and both
refuse unless passed an explicit `allowMutation` argument.
`R23_ProbeAnnotationReconciliation` never passes it and contains no insert,
delete or save call, so **it is safe to run against the manual reference
drawing** — which is the point: that drawing already carries the
manufacturing intent R23 must reproduce, so reconciling against it exercises
R23-406 to R23-409, R23-411 and R23-412 on real data.

Verification is static: 176 offline tests with the same five stale R20
failures, 22 of them new Phase 4 contracts, plus a read-only production
preflight resolving 24 managed components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 3 gate SATISFIED

**2026-08-01.** R23-300 through R23-310 are proved live on P-0251 across five
runs. Final per-view acceptance:

| view | displayMode | projections | axisNormal | anchored | accepted |
|---|---|---|---|---|---|
| `Drawing View4` (primary) | HiddenLinesVisible | 11 | 7 | 11 | **7** |
| `Drawing View7` (side) | HiddenLinesVisible | 11 | 4 | 6 | **2** |
| `Section View J-J` | HiddenLinesRemoved | 11 | 4 | 0 | 0 |
| `Drawing View2` (isometric) | HiddenLinesRemoved | 11 | 0 | 9 | 0 |

`selectionProved=9`, `finalSelectionCount=0`, `selectionClean=True`,
`drawingUnchanged=True`, `partUnchanged=True`.

**R23-310's "four side projections" is unsatisfiable as written, and the
implementation is correct.** `Drawing View7` looks along model Y, so the four
tapped holes project onto **two** page points and the six counterbores onto
**three** — and the mapped counts match those numbers exactly. Two coaxial
holes seen along their axis are one circle on the sheet; SOLIDWORKS holds a
single drawing entity. Two side anchors is the complete answer. All four
tapped holes remain individually resolved in `Drawing View4`, and the
reference drawing calls them out once as `4x`.

The earlier hidden-lines hypothesis was **wrong**: `Drawing View7` is
`HiddenLinesVisible`, so the far-face holes are drawn. The cause is
projection coincidence, not visibility.

**Carry into Phase 5:** required coverage must be counted per distinct page
position per view, not per physical location. Requiring one annotation per
physical location in a side view cannot be satisfied by construction.

`MACRO_SOURCE_REVISION` remains `r22` — the engine has no pipeline caller and
drawing output is unchanged. Phases 4 onward are unstarted.

## Historical: R23-310 fourth live run, primary-view clause proved

**2026-08-01.** The axis gate works. Per-view acceptance:

| view | type | projections | axisNormal | anchored | accepted |
|---|---|---|---|---|---|
| `Drawing View4` (primary) | 4 | 11 | 7 | 11 | **7** |
| `Drawing View7` (side) | 4 | 11 | 4 | 6 | **2** |
| `Section View J-J` | 2 | 11 | 4 | 0 | 0 |
| `Drawing View2` (isometric) | 4 | 11 | 0 | 9 | 0 |

**R23-310's first clause is proved.** All six counterbores plus the stepped
bore accepted in the primary view, each anchored at
`anchorTier=PrimaryTypedHoleDiameter` on the Ø6.6 through hole rather than
the Ø11 mouth, all seven selectable with `ownershipProven=True`,
`selectionClean=True`, both documents unchanged.

The isometric is the gate's best corroboration: 9 anchors, `axisNormal=0`,
0 accepted — no hole axis is normal to an isometric sheet.

**Second clause is not met:** only two of four tapped holes have side
projections. The axis test is correct (all four are `axisNormal` in the side
and section views); the two far-face holes return Nothing from route A and
the section view maps nothing. `IView.GetDisplayMode2` is now recorded per
view, because under Hidden Lines Removed a far-side hole is never drawn and
cannot anchor however the search is written.

One defect fixed: `projection.ProjectedAxisX/Y/Z` were passed as ByRef
out-parameters. A class Public variable is a property, so the write-back was
discarded — `projectedAxis` logged as 0,0,0 while `axisNormal` was right.

## Historical: R23-310 third live run, anchors resolve and select

**2026-08-01.** The `GetCurveParams3` fix worked. `circularEdges` went 0 → 4
on every location with `uMin=0`, `uMax=6.283185307`, `closureM=0` — the
Phase 0 circle proof reproduced exactly. **26 projections anchored**, all via
route A with route C identity confirmation, and **all 26 proved selectable**
with `ownershipProven=True`, `selectedCount=1`, `finalSelectionCount=0`,
`selectionClean=True`, `drawingUnchanged=True`, `partUnchanged=True`.

`ISelectData.View` raised error 91 on all 26 attempts, exactly as this
repository predicted; the guarded binding plus after-the-fact ownership proof
absorbed it every time.

**One defect found: normal-axis compatibility was computed but not enforced.**
`AxisNormalToView` was stored and ignored by `QualificationFailureReason`, so
`Drawing View4` accepted all four M5 tapped holes even though that view's
normal is model Z and the M5 axis lies in the page plane — proved by their
page coordinates matching the model *Y* spread. Those anchors are edge-on,
not circles. They now fail as `AxisNotNormalToView`, kept distinct from
`ProjectionAnchorUnavailable` because the anchor itself is real.

Two questions settled by this run:

- Route A **does** work from a model reached through `IView.ReferencedDocument`
  with the drawing active. The earlier hypothesis was wrong.
- `Section View J-J` maps nothing at all — its component is a synthetic
  section assembly, so a section view will need its own acquisition route.

R23-310 stays open pending the fourth run, and whether "side projections for
all four tapped holes" is achievable is now in doubt: two of the four sit on
the far face and did not map. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23-310 second live run, root cause found

**2026-08-01.** The stage counters isolated the failure in one run. Every
location, in all four real views, reported `sourceFaces=2`,
`facesProjected=2`, `boundaryEdges=4`, `circularEdges=0`,
`firstReject=circle=Reject|reason=ReadError:438`.

**Root cause: `IEdge.GetCurveParams3` returns an `ICurveParamData` object,
not an array of doubles.** `Module13` assigned it into a Variant with a
`Let`, which asks for a default property the object does not have and raises
error 438, killing every candidate edge before any mapping was tried. The
correct pattern was already in this repository —
`Module12_FeatureQualification.ComputeFaceAxialInterval` binds it with `Set`
and reads `.StartPoint` / `.EndPoint` — and Module13 now matches it.

Two further defects were fixed alongside: an unmeasurable closure distance
would have failed open as a closed curve, and selection cleanliness was
reported without a baseline (the observed `finalSelectionCount=1` was the
operator's own pre-existing selection, not this pass).

Also established: the ten sheet placeholders report
`swDrawingViewTypes_e = 7`, and the visible-entity skip cut the run from 154
projections to 44.

R23-310 stays open pending the third run. `MACRO_SOURCE_REVISION` remains
`r22` — no pipeline caller, drawing output unchanged.

## Historical: R23-310 first live run, zero anchors, chain instrumented

**2026-08-01.** `R23_ProbeViewProjections` compiled and ran read-only against
the P-0251 drawing. `drawingUnchanged=True`, `partUnchanged=True`,
`finalSelectionCount=0`, 0 warnings, 0 failures — the safety envelope held.
**No projection resolved an anchor:** all 154 failed
`ProjectionAnchorUnavailable` with `candidates=0`.

Two facts are established:

- The drawing side works. `GetVisibleEntities2` returned 64 edges for
  `Drawing View4`, 68 for `Section View J-J` and 53 for `Drawing View7`,
  and each carried a drawing component (`componentContext=DrawingContextOnly`,
  as predicted for a part drawing).
- 132 of the 154 projections were against `*Left`, `*Bottom`, `*Current`,
  `*Isometric`, `*Dimetric` and `*Trimetric` — sheet standard-view
  placeholders returned by `ISheet.GetViews` with zero visible entities.

**The root cause is not yet known**, and the evidence line could not isolate
it: `candidates=0` cannot distinguish "no faces retained" from "no circular
edge found" from "nothing mapped". That is an instrumentation failure of the
same class Phase 0 solved with `rejectGate`. `ResolveProjection` now counts
`sourceFaces`, `facesProjected`, `boundaryEdges`, `circularEdges`,
`mappedEdges` and `inventoryConfirmed`, and carries the transform's own proof
string into `firstReject`. Views with no visible entities are now skipped
with one line instead of generating dead projections.

R23-310 stays open pending the rerun. No production behaviour is affected:
the engine still has no pipeline caller and `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 3 source-complete

**2026-08-01.** `Module13_ProjectionResolution.bas` resolves each physical
location into a `CViewHoleProjection` for a drawing view: drawing-context
anchor, page-frame centre, and the proofs behind both. R23-300 through
R23-309 are implemented; **R23-310 needs one live run** of the read-only
`R23_ProbeViewProjections` to confirm usable primary projections for the six
face holes and side projections for the four tapped holes.

Verification is static: 142 offline tests with the same five stale R20
failures, 20 of them new Phase 3 contracts, plus a read-only production
preflight resolving 23 managed components. `MACRO_SOURCE_REVISION` remains
`r22` — the engine has no pipeline caller and drawing output is unchanged.

One contract worth carrying: **`ISldWorks.IsSame` is not a Boolean.** It
returns `swObjectEquality` = {0 not same, 1 same, 2 unable to determine}, so
reading it through `NormalizeSwBoolean` would accept "unable to determine" as
proof of identity. Only an exact 1 is accepted.

## Historical: R23-213 CLOSED, P-0251 catalog proved live, Phase 2 gate satisfied

**2026-08-01.** Two read-only `R23_ProbeFeatureCatalog` runs on P-0251. The
second returned **`catalogFailures=None`** with all four R23-213 expectations
met, 0 warnings, 0 failures and `modelUnchanged=True`:

- one six-location M6 counterbore family (radii `0.0055`/`0.0033`, two unique
  X and three unique Y, matching the Phase 0 ordinate evidence);
- one four-location M5x0.8 tapped family;
- one Ø47/Ø40 stepped-bore stack (radii `0.0235`/`0.0200`, two members);
- nothing spurious accepted — 40 rejections, each with an explicit reason.

18 qualifying cylindrical faces consolidated to 11 physical locations through
7 merges. The axial interval earned its place: two coaxial pairs of M5 tap
drills that share a line key were correctly held apart as four distinct
locations by their disjoint intervals, which the Plücker key alone cannot do.

The first run reported `NoFourLocationFamily`, which was real, and the log
exposed two further defects. Three fixes: pattern instances now inherit their
seed's semantics; the seed-chain rejection codes are enforced rather than
merely computed; and `HoleFit` no longer publishes the not-applicable `-1` as
a fit value. See [Changelog.md](Changelog.md).

### Open follow-up: traversal is not exact-once

Comparing the two runs, `visitedFeatures` went **47 → 46** on an unchanged
model, and the sketches visited twice differed between runs. The traversal
key included `ObjPtr`, so one feature reached through two COM wrappers split
into two keys. Fixed by keying on name plus type, with the address kept only
as an unnamed-feature fallback.

This did not affect either catalog — accepted features, locations,
consolidations and families were identical across both runs — but it is not
cosmetic: a doubled visit to an *accepted* feature would add duplicate stack
members and misreport `stackMembers`. **The fix is static-only.** It needs no
dedicated run and can be confirmed by `visitedFeatures` being stable and
duplicate-free on the next live run of any Phase 3 work.

## Historical: R23 Phase 2 source-complete

**2026-07-31.** `Module12_FeatureQualification.bas` implements the Phase 2
feature-qualification engine: `ICE` type normalization, cycle-guarded
traversal of features and subfeatures, suppression proved against the exact
referenced configuration, typed readers for six feature families with paired
selection access and release, ownership from `IFeature.GetFaces`, and
physical locations built from owned cylindrical faces with axial intervals
measured from real boundary edges.

The engine has no pipeline caller. Drawing output is unchanged and
`MACRO_SOURCE_REVISION` remains `target-spec-hybrid-v2-2026-07-29-r22`.
The manifest manages 22 components.

The user deployed the 22 managed components into `Fable.swp` and ran
**Debug > Compile Project**, which stopped at `evidence.InfoCount` in
`R23_ProbeFeatureCatalog` with "Method or data member not found" —
`CRunEvidence` exposes no accessor over its Private `mInfo` collection. The
redundant replay loop was removed in favour of a `WarningCount`/`FailureCount`
tally; every other cross-module reference and call-site arity in the seven new
Phase 1/2 components was then audited against its target and found sound.
**Phase 1/2 source has not yet passed a full VBE compile**; the corrected
source needs redeploying and recompiling.

Verification is otherwise static: 117 offline tests with the same five stale
R20 failures, 23 of them new Phase 2 contracts, plus a read-only production
preflight.

**The one open Phase 2 item is R23-213**, which needs a live run of the
read-only `R23_ProbeFeatureCatalog` against P-0251 to confirm one
six-location counterbore family, one four-location tapped family, one
stepped-bore stack, and nothing spurious accepted.

## Historical: R23 Phase 1 source-complete

**2026-07-31.** The location-graph model is added under
`src/target-spec-hybrid-v2/`: `CFeatureDefinition`, `CPhysicalHoleLocation`,
`CViewHoleProjection`, `CImportedAnnotation`, `CLocationGraph` and
`Module11_GeometryIdentity`. The deployment manifest now manages 21
components.

The change is additive by design (R23-107): the new classes have no callers,
a test asserts no existing production module references them, and runtime
output is identical to r22. `MACRO_SOURCE_REVISION` therefore remains
`target-spec-hybrid-v2-2026-07-29-r22`, and nothing has been deployed.

Physical identity is an infinite axis line — sign-normalized direction plus
line moment — together with the axial interval the material occupies.
Consolidation happens only in
`CLocationGraph.ResolveOrCreatePhysicalLocation`, which requires both the
same line and meeting intervals. That is what merges a counterbore with its
through hole while keeping opposite blind holes on one axis separate, and
neither key contains a feature name.

Verification is static only: 94 offline tests with the same five stale R20
failures, source hygiene clean on all six new files, and a read-only
production preflight resolving 21 managed components. No VBE compilation and
no live run.

## Historical: R23 Phase 0 closed; production Phase 1 unblocked

**2026-07-31.** The corrected disposable probes closed every Phase 0 gate on
the authorized P-0251 fixture. Both datum-first ordinate groups returned
`AddOrdinateDimension = 0 Success` with exact selection counts, correct
display-dimension deltas, `SetPickMode`, zero-selection cleanup, and an
unchanged fixture. Ordinate values are `+15.00`/`-15.00` about the
stepped-bore centre and `10.00`/`50.00`/`90.00` from the bottom-left vertex
datum — the reference scheme.

Contracts carried into production:

- entity correspondence uses route A,
  `IView.GetCorrespondingEntity(modelEdge)`; `IComponent2` mediation returns
  Nothing on this build;
- mapping is per-edge, so every feature-owned edge must be attempted;
- never assign `ISelectData.View` (runtime error 91); activate the view and
  verify with `ISelectionMgr.GetSelectedObjectsDrawingView2`;
- normalize every SOLIDWORKS COM Boolean with `(CDbl(raw) <> 0#)` before any
  negation or compound logic;
- accept ordinate `Type2` values `1`, `7` and `8`;
- `ICurve.CircleParams` is available and returns correct radii; and
- section-line payload segments are view-sketch coordinates requiring exactly
  one conversion before page-frame comparison.

The three J-J top-border violations are production work for R23-704, now
measured truthfully rather than unknown. Production R23 source, `Fable.swp`,
fixtures, the protected baseline and manual references remain unchanged, and
no production implementation has begun.

## Historical: R23 implementation at the Phase 0 live gate

The approved R23 plan is now an active implementation checklist. Pre-change
provenance is retained under:

`test_assets/iteration_evidence/r23/20260730-075811/prechange/`

Confirmed offline:

- production `Fable.swp` and its recoverable backup have matching SHA-256
  `C8076D713DD8F64AC75F93871C1EA4A2D1F01EE2BC691323736013BFEB2803F2`;
- embedded and exported R22 source match for all 15 managed components;
- the guarded deployment preflight passes with the bootstrap present;
- the installed SOLIDWORKS 2025 SP1.2 interop remains version `33.1.2.4`;
- the local API MCP confirms that `GetTypeName2="ICE"` must be normalized
  through `GetTypeName`, and that successful typed-definition selection access
  must be released; and
- the standalone, read-only feature/curve probe is prepared at
  `tools/r23-probes/Module_R23Phase0FeatureProbe.bas`;
- a disposable copy of `Fable.swp` plus a guarded three-module overlay is
  prepared for the expanded import-transaction comparison; and
- the disposable import manifest passed read-only preflight with all three
  overlay sources present and the deployment bootstrap installed.

The first run against the authorized workspace P-0251 fixture completed, but
its transcript was rejected as incomplete probe evidence. A recursion guard
based only on `ObjPtr(feature)` collapsed distinct COM wrappers that reused an
address, so only 15 features were visited and the manufacturing-relevant ICE,
Hole Wizard, extrusion, and mirror records were skipped. The rejected
transcript is retained under the R23 live-probes evidence folder.

The exported probe has been corrected to use a composite diagnostic recursion
key, accept both scalar and array `IsSuppressed2` returns, and include base
`EXTRUSION` definitions in the typed extrusion probe. It now records
contour/profile state and mirrors every evidence record to a timestamped log
so Immediate Window truncation cannot invalidate the rerun. Static source
structure, continuation, encoding, and deployment-preflight checks pass. The
corrected feature source has now received full-project VBA compilation and a
complete user-operated run. Its retained log is
`R23_FEATURE_20260731_040539.log`: it visited 47 features, resolved all three
`ICE` entries to `Cut`, proved typed definition access/release and owned
geometry, read both native Hole Wizard definitions, and identified `Mirror1`'s
M5 seed feature. `IsSuppressed2` returned Empty for the active configuration
on every feature, so the explicitly labelled current-configuration fallback
reported active state. Both `IsCircle`/`GetCurveParams3` call orders remained
stable; the anomalous `CircleParams=SkippedNotCircle` result means production
must not rely on `CircleParams`. The disposable import overlay and final
drawing-contract extension have since received user-operated runtime execution;
their findings are summarized below.

Production R23 VBA has not been written or deployed. The approved plan
explicitly prohibits crossing into production feature classification until the
installed SW2025 build proves the three P-0251 underlying `ICE` types, typed
definition access, owned geometry, curve read order, expanded import
transaction, explicit ordinate transaction, and representative section
dimension/geometry behavior.

The user performs the remaining live SOLIDWORKS work. Both import variants
completed in the disposable SWP without changing the fixture:
`AllViews=True` returned 25 dimensions but left the side view empty, whereas
the explicit section/side/primary calls returned the same 25 as 17 section,
2 side, and 6 primary dimensions. The isometric stayed clean in both. R23 will
use deterministic `AllViews=False`, in section/side/primary order, with
`DuplicateDims=True` and the expanded mask containing
`swInsertDimensions = 8`. The logs prove one imported M5 native callout in the
section, but no imported M6 counterbore callout and no H7/nonzero tolerance.

The two final drawing-contract probes also ran, but they did not close Phase 0:

- `R23_ProbeDatumFirstXYOrdinates` found the visible component and then stopped
  fail-closed with `CounterboreLocationsUnavailable`. It never selected a
  datum, never called `AddOrdinateDimension`, and therefore proves nothing
  about the ordinate API. Cleanup completed with zero selections and the
  fixture remained unsaved. The disposable probe must add per-face/per-edge
  qualification and direct/component/visible-entity mapping diagnostics before
  a corrected rerun.
- `R23_ProbeSectionDimensionsAndJJGeometry` found the exact imported 47 mm and
  40 mm dimensions as real `swDiameterDimension = 6` items. Neither carried H7
  or a nonzero tolerance. The probe's per-dimension target fields were not
  reset on every loop iteration, so later labels contain stale state and must
  be corrected before they can be used as acceptance evidence.
- The J-J payload is structurally complete at 49 returned items, but it mixes
  source-view sketch coordinates for segments with page coordinates for arrows
  and labels. The current production clearance comparison therefore compares
  unlike coordinate frames. The retained screenshot also confirms the upper
  arrow/label enters the top border/zone band and the lower arrow/label enters
  the part-identification band.

Production R23 remains blocked. The next permitted implementation work is only
to correct the disposable probes, run the direct part-source H7 readback, and
close the entity-mapping and page-coordinate contracts. The complete handoff
and exact correction checklist are in
`docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md` and
`docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md`.

### Corrected-probe run result (2026-07-31, build 20260731.2)

The user compiled and ran both corrected entry points on P-0251. Two of the
three gates closed.

- **Section dimensions: closed.** Exactly one `DIAMETER_47`, one
  `DIAMETER_40`, one `FIRST_LINEAR`, no stale labels, both diameters
  `type2=6`.
- **H7 authority: closed.** The direct part-source readback proved
  `D1@Sketch4` carries `toleranceType=0`, `fitType=-1`, and no fit strings,
  so H7 is not in the model. The user selected controlled
  target-spec/reference authority: R23 will apply `H7 +0.025/0.000` to an
  associative `Ø47` dimension with provenance recorded as target-spec
  authority, never as model data, and QA must state that explicitly.
- **J-J frames: closed.** All six returned segment endpoints matched the
  captured `CreateLine` view-sketch inputs at `deltaM=0`, and the single
  inverse conversion is cross-validated against the independently returned
  page-frame arrow endpoints.
- **J-J clearance: measured.** Three violations, all at the top border
  (segment 1 start, upper arrow, upper label). The lower arrow and label
  **clear** the measured part-identification note extent by about 7.7 mm,
  correcting the earlier screenshot-derived claim of a lower-band intrusion.
- **Ordinate ownership: proved.** After the `CBool` normalization fix,
  twelve of the eighteen owned counterbore faces read as cylinders with
  correct radii, all passed the page transform with outline containment, and
  the six page centres form exactly the required grid — two unique X and
  three unique Y, spaced 30 mm and 40 mm.
- **Entity correspondence: settled.** Route A
  (`IView.GetCorrespondingEntity(modelEdge)`) works — 12 of 24 counterbore
  edges and 114 of 154 vertices mapped. Route B
  (`IComponent2.GetCorrespondingEntity`) returned `Nothing` every time and is
  unusable here. Route C confirms route A returns genuine drawing entities by
  `ISldWorks.IsSame` identity. Mapping is per-edge, so every owned edge must
  be attempted. The full scheme resolved: six unique locations, two unique X,
  three unique Y, a stepped-bore X datum and a bottom-left vertex Y datum.
- **`ICurve.CircleParams` works**, returning radii matching the owning
  cylinders exactly. The R23-006 exclusion was an artifact of the guard
  defect and is closed.
- **`ISelectData.View` cannot be assigned on this build.** Both ordinate
  groups failed with runtime error 91 before any selection;
  `CreateSelectData` returned a live object, so the failing statement is the
  documented get/set assignment. This reproduces the behaviour already
  recorded in `Module2_DrawingPipeline.CreatePrimarySection`. The probe now
  guards the binding, records its outcome, and proves each selection's owning
  view through `ISelectionMgr.GetSelectedObjectsDrawingView2`.
  `AddOrdinateDimension` has still never been called.
- **Historical note — superseded.** The
  second and third runs both stopped at the edge-closure gate. The
  `rejectGate` instrumentation isolated the mechanism: one call logged both
  `isCircle=True` and `rejectGate=IsCircleFalse`, because `CStr` renders the
  value `True` while `If Not value` yields `-2`, which VBA treats as True.
  `CBool` is not a dependable normalization — it worked for
  `ISurface.IsCylinder` but not for `ICurve.IsCircle`. The probe now uses
  `NormalizeSwBoolean`, an explicit `(CDbl(rawValue) <> 0#)` comparison,
  across `IsCircle`, `IsCylinder`, `ActivateView`, both `Select4` calls and
  `GetSaveFlag`. Mapping routes A, B and C have still never executed, so
  `IView.GetCorrespondingEntity` behaviour on this scaffold remains unknown
  and one further probe rerun is required.
- **`CircleParams` exclusion withdrawn.** The `SkippedNotCircle` result that
  barred it from production came from the same `If Not <SOLIDWORKS Boolean>`
  defect in the feature probe's `ReadCircleState`; the API was never called.
  Its behaviour is untested rather than anomalous. Production must still not
  depend on it until a run exercises it.

The disposable-probe corrections are now implemented at probe build
`20260731.2-mapping-frame-h7-contracts` in
`tools/r23-probes/import-transaction-source/`: per-face/per-edge ordinate
qualification with direct, component-mediated, and visible-entity mapping
comparison; per-iteration section-state reset with exact type-6 47/40 target
enforcement; direct part-source `D1@Sketch4`/`D1@Sketch6` tolerance readback
feeding an explicit H7 authority record; and frame-proved J-J geometry with a
single sketch-to-page conversion and page-frame clearance verdicts against
the content border, title block, measured part-identification note extent,
and view outlines. This state is static verification plus read-only
disposable preflight only: the corrected probes still require user-operated
deployment, full-project VBE compilation, execution on the authorized P-0251
fixture, and returned logs/screenshots before any Phase 0 gate can close.

The full offline suite remains at its known baseline of 69 passes and five
stale R20-contract failures.

The protected baseline, fixture models, manual references, production VBA
source, and production SWP remain unchanged.

## R22 corrected review line is awaiting user compile and runtime evidence

The authoritative managed source identifies as
`target-spec-hybrid-v2-2026-07-29-r22`.

R22 combines all source changes from the r20/r21 commit line, retains the
review fixes that satisfy the verified contracts, and corrects the latest
commit's three substantive defects:

- non-circular cylinder trims now fail closed at `ICurve.IsCircle=False`
  instead of being assigned a centre from `ISurface.CylinderParams`;
- pattern ownership uses the exact SOLIDWORKS 2025 `GetTypeName2` literals,
  including `APattern`, `LocalChainPattern`, `DimPattern`,
  `DerivedHolePattern`, `SketchPattern`, and `LocalSketchPattern`; and
- the arrange fallback delegates ordinate-set jogs to
  `IDisplayDimension.AutoJogOrdinate`, using manual deterministic lanes only
  for non-ordinate dimensions.

The local MCP contracts and installed SOLIDWORKS 2025 SP1.2 interop
`33.1.2.4` confirm the used member signatures and dimension enum values. The
retained P-0251 curve probe confirms that all relevant owned model edges already
provide `IsCircle=True` and valid `CircleParams`, so removing the unsafe
cylinder-trim inference does not discard those valid candidates.

The guarded r22 deployment and exact managed-source readback are the next
offline gates. After deployment, the user must run VBA Editor **Debug > Compile
Project**, then run only an authorized fixture and return the complete
Immediate Window output, `QA_REPORT.txt`, and one uncropped full-sheet
screenshot. Static checks, interop reflection, and the deployment bootstrap
probe are not full VBA compilation or manufacturing acceptance.

## Historical R20 compile hotfix and runtime baseline

The authoritative managed source and embedded macro now identify as
`target-spec-hybrid-v2-2026-07-28-r20`.

The user's full-project VBA compile exposed `Argument not optional` at
`Module6_QAEngine.CheckSectionLineClearance` because the r20 source called
`IView.GetSectionLineCount2` without its mandatory `ByRef Size As Long`
argument. The project `solidworks-api` MCP and the installed SOLIDWORKS 2025
interop both confirm the exact signature
`GetSectionLineCount2(ByRef Size As Long) As Long`; `GetSectionLineInfo2`
remains parameterless. The exported source now supplies the size argument,
rejects invalid/nonmatching array sizes, and has a regression assertion against
the zero-argument call.

After the user closed the locked VBA project, the guarded deployment completed
under `test_assets/iteration_evidence/swp_deployment/20260728_142300/`.
Candidate and promoted readbacks match all 15 managed components and the exact
r20 revision. `compile-probe-scope.txt` explicitly records that the automated
probe covers bootstrap execution only.

R20 is based on the r19 P-0251 evidence at:

- `test_assets/iteration_evidence/macro_qa/20260728_091302_P-0251-14A-001/QA_REPORT.txt`

R19 preserved four created views, ten imported display dimensions, a clean
isometric, the J-J structure, metric mass `1.30`, title properties, general
notes, part identification, and complete QA output. It nevertheless failed
because all 32 complete internal-cylinder boundary edges in each supported view
stopped at `ClosedCircleIsCircleFalse`, leaving zero mapped circular edges,
ownership candidates, canonical physical locations, or ordinate groups. The
side/section gap sat exactly on the collision threshold; the lower J-J marker
entered the part-identification band; annotation QA incorrectly excluded the
entire lower sheet below the view-placement boundary; and auto-arrange did not
prove any view-scoped selection.

R20 repairs those source defects:

- complete boundary edges are proved with `IEdge.GetCurveParams3`; when a
  trimmed edge does not identify as an `ICurve` circle, its already
  ownership-proven internal face supplies the cylinder center, axis, and radius
  through `ISurface.CylinderParams`;
- used SOLIDWORKS Boolean results are normalized before negation/compound logic;
- `ModelToViewTransform` page coordinates are no longer translated by
  `IView.Position` a second time, and the center datum compares against the
  transformed model origin. Every candidate centre, projected origin, and
  mapped vertex now fails closed unless the transformed point lies within the
  current `IView.GetOutline`, with explicit page-frame evidence;
- the common view gap is 12 mm, layout comparison has a numeric tolerance,
  P-0251's orthographic/section row is biased upward, and all requested centers,
  actual outlines, and pair clearances are logged separately for initial and
  final layout;
- annotation QA records annotation type/name/position and validates against the
  real zoned border and rectangular title block. `UsableBottom` remains a view
  placement boundary only. The measured part-ID extent is retained and
  `GetSectionLineInfo2` segment, arrow, and J-label geometry must clear it;
- auto-arrange uses `ISelectData.View`, checks `IAnnotation.Select3`, verifies
  selection count, and logs the `AlignDimensions` result. It is now a required
  stage; `False` invokes deterministic 6 mm `SetPosition2` lanes with exact
  readback and content-border proof; and
- three controlled P-0251 callouts now define the stepped bore, six
  counterbored face holes, and four tapped side holes. Each callout uses the
  documented `<MOD-DIAM>` syntax, selects an ownership-proven drawing edge
  before `InsertNote`, and must read back nonzero attachments, a visible leader,
  safe extent, and clearance from other note extents/annotation origins.
  Existing-note reuse requires the complete normalized controlled definition,
  so a shorter imported Hole Wizard phrase cannot falsely prove the stage.

The project-local suite passes **74 tests and 13,608 structural subtests**.
The current guarded deployment proves a 15/15 managed-source match and exact
r20 revision. Its `COMPILE_PROBE|status=SUCCESS` record proves only that the
deployment bootstrap could execute; it does not perform VBA Editor **Compile
Project** and must not be cited as full-project compilation.

The required next sequence is: open `Fable.swp`, run VBA Editor **Debug >
Compile Project**, and—only if compilation succeeds—run P-0251. Runtime evidence
remains the new `QA_REPORT.txt`, complete Immediate Window output, and one
uncropped full-sheet screenshot. Computer Use was not used and remains
disallowed unless the user explicitly requests it.

## Historical R18 model-first drawing-output repair

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

1. Run synchronized r22 on P-0251 and retain its complete E6 runtime and drawing
   evidence.
2. Confirm nonzero complete circular boundaries/mapped edges, ten canonical
   P-0251 physical locations, and the required X/Y ordinate coverage.
3. Prove the model-first Center-X and Bottom-Y datum entities are selectable in
   the intended drawing views and that every ordinate transaction cleans up.
4. Confirm imported dimension counts remain `9, 1, 0, 0` before new ordinates,
   the isometric stays undimensioned, and the front-view multi-dimension
   arrange stage proves either `AlignDimensions` or its deterministic lane
   fallback with exact position readback.
5. Confirm all three manufacturing callouts render the required tokens and
   SOLIDWORKS diameter symbols, remain attached to the intended bore/face/side
   geometry with leaders, and avoid other annotations, view outlines, part ID,
   and title block. Confirm parsed J-J segment/arrow/label geometry clears the
   measured part-ID extent.
6. Confirm annotation-origin diagnostics no longer reject legal lower-left
   content and identify any genuine remaining border/title/leader violation by
   type, name, and coordinates.
7. Confirm the unique mass-note link reads `$PRP:"Mass"` and visibly renders
   `1.30`.
8. Run only the three authorized fixtures and retain the complete E6 regression
   matrix.
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

Codex owns offline source work, API validation, output diagnosis, reference
comparison, tests, and automatic synchronization through the guarded
PowerShell deployer. The user owns the actual SOLIDWORKS macro run and shares
the resulting QA output and screenshot.

Computer Use must not be used for SOLIDWORKS unless the user explicitly asks
for it. A newly written revision is deployed automatically through
`tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`, after which the user is asked to
run the macro.

## Exact handoff point

The coherent r22 managed source is embedded in `Fable.swp`, bootstrap-probed,
and verified 15/15 against exported source. The user should now compile and run r22 on
`P-0251-14A-001.SLDPRT` with the same acceptance-profile settings and share the
new QA report, complete Immediate Window output, and one uncropped full-sheet
screenshot. No Computer Use is required.
