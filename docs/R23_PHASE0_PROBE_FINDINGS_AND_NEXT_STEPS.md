# R23 Phase 0 Probe Findings and Next Steps

## 0. Corrected-probe run (build 20260731.2) — supersedes several findings below

The user compiled and ran the corrected disposable probes on the authorized
P-0251 fixture. Results, by gate:

**Section dimensions — closed.** The per-iteration reset works: indices 2, 4,
5 and 7-17 now carry an empty target instead of a stale label.
`diameter47Count=1`, `diameter40Count=1`, `linearTargetSelected=True`,
`exactTargetCounts=True`. Both diameters are `type2=6`,
`diametricLinearFound=False`.

**H7 authority — resolved as absent, decision required.** The direct
part-source readback succeeded:

```text
D1@Sketch4@P-0251-14A-001.Part  nominal 0.047 m  toleranceType=0  fitType=-1
D1@Sketch6@P-0251-14A-001.Part  nominal 0.040 m  toleranceType=0  fitType=-1
R23_H7_AUTHORITY|h7PresentInSource=False
```

H7 does not exist in the model. The user has selected **controlled
target-spec/reference authority**: R23 will create or reuse an associative
`Ø47` dimension and apply `H7 +0.025/0.000` from the approved reference
specification, recording provenance as target-spec authority. Production must
never present that tolerance as model-derived, and QA must say so explicitly.
R23-806 is closed; R23-807 (no free-text substitute) still applies.

**J-J coordinate frames — closed.** The mixed-frame conclusion is now proved
rather than inferred. All six payload segment endpoints matched the captured
`CreateLine` view-sketch inputs to `deltaM=0.000000000`
(`payloadSegmentFrame=ViewSketchProved`). The single inverse conversion is
independently cross-validated: converted segment 1 lands at page X
`0.095932223`, exactly the upper arrow's payload endpoint X, and converted
segment 3 lands at `0.071467223`, exactly the lower arrow's endpoint X.

The construction capture also proves the view moved after the section line was
built (construction page Y `0.308868` versus current `0.289788`), confirming
that section geometry must be re-read after every view move.

**J-J clearance — measured, and one earlier finding is corrected.** Against
page-frame regions (content border `0.010, 0.010` to `0.410, 0.287`;
part-identification note `0.023518, 0.014625` to `0.160787, 0.031595` from
`INote.GetExtent`):

| Item | Page geometry | Result |
|---|---|---|
| Segment 1 | starts Y `0.289788` | above content-border top |
| Upper arrow | Y `0.287788` | above content-border top |
| Upper label | box Y `0.296839`-`0.303189` | above content-border top |
| Segment 2, 3 | — | inside, clear |
| Lower arrow | Y `0.042332` | inside, clears part-ID |
| Lower label | box Y `0.039318`-`0.045668` | inside, clears part-ID |

`violations=3|unavailableChecks=0`. All three violations are the upper
overshoot produced by `topY + extension`.

The earlier statement that the lower arrow and J label enter the
part-identification band is **not supported by measurement**. The lower label
sits about 7.7 mm above the measured note extent. That earlier claim came from
reading a screenshot; the identical geometry values now measure clear. The
lower end of the J-J line is not a defect to fix.

**Ordinate mapping — still not proved; root cause was a probe defect.** The
corrected diagnostic did its job and moved the failure to a precise cause:

```text
R23_ORDINATE_FEATURE|feature=CBORE for M6 Socket Head Cap Screw1|ownedFaces=18
R23_ORDINATE_FACE|...|faceIndex=0..17|isCylinder=False|reason=NotCylindrical
```

All 18 owned faces were rejected before any mapping route ran. The accepted
feature probe reads the same feature's faces as cylinders
(`faceIndex=0` radius `0.0055`, `faceIndex=2` radius `0.0033`) using
`CBool(ISurface.IsCylinder)`. The corrected drawing probe used a bare
`If Not surface.IsCylinder Then`. A raw SOLIDWORKS `VARIANT_BOOL` of `1`
makes `Not value` evaluate to `-2`, which VBA treats as True, so every face
was rejected regardless of its surface type.

This is the Boolean-normalization defect class already recorded as an R20
repair. The same run contains its own control: `ReadEdgeCircleEvidence` uses
`CBool(ICurve.IsCircle)` and returned a healthy mix of True and False on the
64 visible drawing edges in the same document.

Probe build `20260731.2` now normalizes with `CBool`, logs the raw returned
value so the defect cannot hide again, and applies the same normalization to
`IDrawingDoc.ActivateView` on the ordinate transaction path. Routes A, B and C
remain unexercised; the rerun is the first run that can actually test them.

**Not yet proved by any run:** counterbore-to-drawing entity correspondence,
datum selection, `AddOrdinateDimension` behaviour, and ordinate readback.

### 0.1 Second corrected run — ownership proved, closure gate defective

The `CBool` fix worked. Twelve of the eighteen owned faces now read as
cylinders with the correct radii (`0.0055` counterbore, `0.0033` through
hole), and every one passed `TransformPointToView` with outline-contained
page proof. The six resulting page centres are exactly the required grid:

| Page X | Page Y |
|---:|---:|
| 0.080932 | 0.077060 |
| 0.080932 | 0.117060 |
| 0.080932 | 0.157060 |
| 0.110932 | 0.077060 |
| 0.110932 | 0.117060 |
| 0.110932 | 0.157060 |

Two unique X and three unique Y, spaced 30 mm and 40 mm, matching the
drawing's `30.00` and `40.00`. Feature ownership, cylinder qualification and
the model-to-page transform are therefore proved for all six counterbore
locations.

The run then stopped at the closure gate. Every edge in the document returned
`completeCircle=False` — all 24 feature-owned edges and all 64 visible
drawing edges — including roughly thirty reporting
`isCircle=True`, `curveParams3Available=True`, `uMin=0`,
`uMax=6.283185307`, `endpointGapM=0.000000000`. Those values satisfy every
gate, so a universal False means the last gate always exited.

`GEOMETRY_TOLERANCE_M` is `0.0000001` in r22 source, so `0 > tolerance` is
False and the success assignment should have been reached. The only construct
in that chain that can swallow the assignment is the single-line
`If ... Then _` line continuation that immediately preceded it.

`ReadEdgeCircleEvidence` now uses block `If` statements throughout, returns
the exact `rejectGate` for every edge, and logs the `closureToleranceM` value
actually in force. No `Then _` construct remains anywhere in the module. The
next run either returns `completeCircle=True` for the twelve cylindrical
counterbore boundaries, or names the gate and tolerance responsible.

Mapping routes A, B and C have still never executed.

### 0.4 Fifth run — PHASE 0 CLOSED

Both ordinate groups completed. This is the last outstanding Phase 0 gate.

| Group | Datum type | Appended | Final sel. | API | Dim delta | Cleanup |
|---|---|---:|---:|---|---|---:|
| `P0251-FACE-X` | 1 (edge) | 2 / 2 | 3 / 3 | `0 Success` | 6 → 8 | 0 |
| `P0251-FACE-Y` | 3 (vertex) | 3 / 3 | 4 / 4 | `0 Success` | 8 → 11 | 0 |

Every selection reported `ownerView=Drawing View1`, `SetPickMode` was called
on both exits, final selection count is zero, and the fixture is unchanged
(`modelUnchanged=True`, `drawingSaved=False`).

**The `ISelectData.View` prediction is confirmed.** Both groups logged
`viewBinding=UnboundAfterError:91`, and both then completed normally with
unbound selection data. Activating the view first and proving ownership
afterwards is a sufficient and sufficient-only substitute; the binding itself
is unavailable on this build.

**Ordinate values are semantically correct.** Computed from the recorded
datums:

| Direction | Datum (page) | Ordinate values |
|---|---|---|
| X | `0.095932` stepped-bore centre | `+15.00`, `-15.00` mm |
| Y | `0.067060` bottom-left vertex | `10.00`, `50.00`, `90.00` mm |

The screenshot shows exactly those values, with `0` at each datum. This is
the reference scheme: symmetric centre-zero X and bottom-zero Y.

Five created dimensions for six physical locations is the designed
deduplication (R23-505): six locations collapse to two unique X and three
unique Y coordinates, each crediting the locations it represents.

**Created ordinate types are `1` and `7`, not `7` and `8`.** Readback:

- created with `swAddOrdinateDims_Horizontal=3` → `Type2=1`
  (`swOrdinateDimension`), positions `(0.080932, 0.048942)` and
  `(0.110932, 0.048942)`;
- created with `swAddOrdinateDims_Vertical=2` → `Type2=7`
  (`swHorOrdinateDimension`), positions `(0.041996, 0.077060 / 0.117060 /
  0.157060)`.

Production QA must accept `1`, `7` and `8` when classifying ordinates. A
filter limited to `7` and `8` would have missed both X ordinates. The probe's
readback already accepted all three, which is why it reported five.

Both ordinate sets lie inside the measured content border
(`0.010, 0.010`-`0.410, 0.287`) and clear the part-identification band
(top `0.031595`).

**Phase 0 stop condition (section 9) is satisfied in full.** The remaining
J-J top-border violations are a known production defect for R23-704, not a
diagnostic gap: they are now measured truthfully in the page frame.

### 0.3 Fourth run — mapping settled, six locations and both datums resolved

The Boolean normalization cleared every qualification gate and the run
reached the ordinate transaction.

**Entity correspondence is settled.** Route A
(`IView.GetCorrespondingEntity(modelEdge)`) works: 12 of 24 counterbore edges
and 114 of 154 body vertices mapped, `error=0`. Route B
(`IComponent2.GetCorrespondingEntity`) returned `Nothing` on every attempt.
Route C confirms route A's output is genuine drawing context: each mapped
entity identity-matches a visible-edge entry via `ISldWorks.IsSame`
(`chosenVisibleIndex` 7, 19, 29-38, 44).

Mapping is per-edge: for each counterbore only one of the two owned circular
edges maps, so production must try every owned edge before failing a
location.

**The required scheme resolved end to end:**

- `acceptedUniqueLocations=6`, `uniqueProjectedLocations=6`
- `xUnique=2` (`0.080932`, `0.110932`), `yUnique=3` (`0.077060`,
  `0.117060`, `0.157060`)
- X datum: stepped-bore centre from `Cut-Extrude1`, r `0.0235`, page
  (`0.095932`, `0.227060`)
- Y datum: bottom-left mapped vertex, page (`0.060932`, `0.067060`)
- 31 of 64 visible drawing edges are complete circles

**`CircleParams` works.** Seven values per edge with radii matching the
owning cylinders exactly (`0.005500000`, `0.003300000`, `0.023500000`). The
R23-006 exclusion was entirely an artifact of the guard defect.

**Both ordinate groups failed with runtime error 91** before any selection:
`datumSelected=False`, `appended=0`, `resultCode=-9999`, cleanup zero and
`SetPickMode` called on both exits. `CreateSelectData` returned a live
object, so the failing statement is `Set selectData.View = swView`.

This reproduces a defect this repository already recorded:
`Module2_DrawingPipeline.CreatePrimarySection` carries a comment stating the
same assignment raises error 91 in this VBA host, with the workaround of
activating the source view first and proving ownership after `Select4`. The
MCP documents the property as `get; set`, so this is an installed-build
deviation rather than misuse.

The probe now attempts the binding in a guarded helper, records
`viewBinding=Bound` or `UnboundAfterError:91`, proves each selection's owning
view through `ISelectionMgr.GetSelectedObjectsDrawingView2`, and tags every
group record with the exact `lastStep` reached.

`AddOrdinateDimension` has still never been called.

### 0.2 Third run — mechanism isolated, defect class eliminated

The `rejectGate` field settled it in one run. Every counterbore boundary
edge logged, from a single call:

```text
isCircle=True ... rejectGate=IsCircleFalse ... completeCircle=False
```

`CStr(circleFlag)` rendered `True` while `If Not circleFlag` fired on that
same variable. The block-`If` restructuring was therefore not the fix; it was
the instrument that exposed the real one.

Established contract for installed SOLIDWORKS 2025 SP1.2 COM Booleans in this
VBA host:

| Construct | Safe? |
|---|---|
| `If value Then` | yes |
| `If value = False Then` | yes |
| `CStr(value)` | yes, renders `True` |
| `If Not value Then` | **no** — yields `-2`, treated as True |
| `CBool(rawVariant)` then `Not` | worked for `ISurface.IsCylinder` |
| `CBool(comCall)` then `Not` | **failed** for `ICurve.IsCircle` |

`CBool` is not dependable. The probe now centralizes
`(CDbl(rawValue) <> 0#)` in `NormalizeSwBoolean` and applies it to
`IsCircle`, `IsCylinder`, `ActivateView`, both `Select4` calls and
`GetSaveFlag`. Every remaining `If Not` in the module is an object
`Is Nothing` test, a VBA built-in, or a normalized value.

This also explains the two earlier failures as one defect: the face gate
(`Not surface.IsCylinder`) and the edge gate (`Not circleFlag`).

**The `CircleParams` exclusion is withdrawn as unproved.**
`Module_R23Phase0FeatureProbe.ReadCircleState` line 611 guards the call with
`If Not isCircle Then`, so `CircleParams` was never invoked on those edges
and the `SkippedNotCircle` sentinel recorded the guard, not the API. Its
behaviour remains untested; the corrected probe now reads it as
non-load-bearing evidence. Production must still not depend on it until a run
exercises it.

Mapping routes A, B and C have still never executed.

**Evidence date:** 2026-07-31

**Fixture:** `test_assets/models/P-0251-14A-001.SLDPRT`

**Configuration reported by SOLIDWORKS:** `Defualt`

**SOLIDWORKS runtime:** 2025 SP1.2, `RevisionNumber=33.1.2`

**Purpose:** Record what the R23 probes actually proved, identify defects in
the probes and current production QA, and define the smallest diagnostic rerun
required before production R23 implementation.

## 1. Evidence inventory

All retained evidence is under:

`test_assets/iteration_evidence/r23/20260730-075811/live-probes/`

| Evidence | SHA-256 | Role |
|---|---|---|
| `R23_FEATURE_20260731_040539.log` | `B38E098C1AC55D01612A9FC8D489EF308FA8F103A80F56D5ADAC3D66749FB699` | complete feature/curve contract |
| `R23_IMPORT_AllViewsTrue_20260731_041602.log` | `D743BCD0A106C63626AF1123A4DF1065E5B73B826567B834C06E3BF0AE3A36F7` | `AllViews=True` import |
| `R23_IMPORT_SelectedViewsFalse_20260731_041721.log` | `40D5F5672EB213B16587C8B0127159B7372A71B9FCE0E2FB1FFED54327484D54` | deterministic per-view import |
| `R23_ORDINATE_20260731_064108.log` | `4EA6D3352B6C083840CF782BC8EE0A0DA04BFB75FE89B58088D428DAD9F60871` | final ordinate attempt |
| `R23_SECTION_20260731_064231.log` | `94BB1014361BF2C1C2820855B34173AB847341571EACCB5ED7E0AD2218B8D37B` | final section readback |
| `R23_ORDINATE_20260731_064108_FULL_SHEET.png` | `A7D621913BB329FBCA4A06E6E01CE48836C31A5606951C262D9649059ECC8CAF` | ordinate-run full sheet |
| `R23_SECTION_20260731_064231_FULL_SHEET.png` | `D463E2DB6602A220C68859EC663AEE6A4DBCE626BC2D0C7FA13DE16FBF3036FE` | section-run full sheet |

The two screenshots were copied from temporary clipboard paths into the
evidence folder so the handoff does not depend on temporary files.

## 2. Confidence summary

| Contract | Result | Confidence |
|---|---|---|
| `ICE` resolves to `Cut` for three P-0251 cut features | proved live | high |
| typed Hole Wizard/extrude definition access and release | proved live | high |
| six M6 counterbore source locations exist | proved live | high |
| `CircleParams` is safe to depend on | disproved | high |
| expanded import with `swInsertDimensions=8`, `DuplicateDims=True` | proved live | high |
| `AllViews=False` section/side/primary transaction | selected by live comparison | high |
| native M5 callout import | proved live | high |
| native M6 counterbore callout import | absent | high |
| H7/nonzero tolerance import | absent | high |
| six selectable counterbore drawing projections | not proved | high |
| datum-first ordinate selection/API creation | not executed | high |
| exact imported `Ø47` and `Ø40` identities | proved live | high |
| section-dimension target labels after exact hits | probe-corrupted | high |
| J-J flattened-array structure | proved live | high |
| J-J page-coordinate clearance | failed visually / not truthfully checked by current parser | high |

## 3. Feature probe findings

The accepted feature log visited 47 features and captured:

- `CBORE for M6 Socket Head Cap Screw1 [HoleWzd]`;
- `Cut-Extrude1`, `Cut-Extrude3`, and `Cut-Extrude4` as
  `GetTypeName2="ICE"`, `GetTypeName="Cut"`;
- `M5x0.8 Tapped Hole1 [HoleWzd]`;
- `Mirror1 [MirrorPattern]` with the M5 seed; and
- relevant sketch, fillet, and chamfer context.

Typed definitions and owned cylinders were readable. Both tested
`IsCircle`/`GetCurveParams3` orders stayed stable on the relevant boundaries.
The `CircleParams` helper remained anomalous, so it is excluded from the
production contract.

Configuration caveat:

- `IsSuppressed2(swSpecifyConfiguration, Array("Defualt"))` returned Empty;
- active-model `IsSuppressed` reported active state;
- this is adequate for the active diagnostic run only;
- production drawing-view qualification must use each view's
  `ReferencedConfiguration` and fail closed when unproved.

## 4. Import probe findings

Both variants returned the same 25 unique source dimension identities.

`AllViews=True`:

```text
primary=8
side=0
section=17
isometric=0
```

Deterministic `AllViews=False`:

```text
primary=6
side=2
section=17
isometric=0
```

R23 decision:

- use explicit per-view selection;
- call section, side, then primary;
- keep `DuplicateDims=True`;
- keep the expanded mask including `swInsertDimensions=8`;
- never count imported linear dimensions as ordinate coverage.

The only imported native hole callout was the M5 section callout. Its variables
included 4.2 mm tap-drill diameter, 12.4 mm drill depth, M5x0.8 thread, and
10 mm thread depth.

No native M6 counterbore callout, H7 fit, or nonzero tolerance was found.

## 5. Ordinate probe result

Authoritative output:

```text
R23_DRAWING_PROBE_SCAFFOLD|...|displayDimensions=25
R23_ORDINATE_COMPONENT|view=Drawing View1|component=P-0251-14A-001-1
R23_ORDINATE_FATAL|reason=CounterboreLocationsUnavailable
R23_DRAWING_PROBE_END|...|cleanupSelection=0|setPickModeCalled=True
```

The run proved:

- the fresh drawing scaffold and per-view import completed;
- the primary view returned exactly one visible component;
- the model remained unchanged;
- selection cleanup was zero.

The run did not prove:

- one mapped counterbore edge;
- six physical drawing projections;
- X or Y datum selection;
- appended selection order/count;
- any `AddOrdinateDimension` return code; or
- ordinate display/readback.

The full-sheet screenshot correctly contains no ordinates.

### 5.1 Why the failure is not yet diagnosed

`CollectFeatureCylinderLocations` returns False when no accepted drawing entity
is added, but it does not log the rejection point for each owned face/edge.

`FindMappedCompleteCircle` silently returns Nothing when:

- an edge is rejected as noncircular/incomplete;
- correspondence returns Nothing; or
- a COM error occurs.

The final fatal reason therefore collapses several distinct contracts into one
message.

### 5.2 Required corrected diagnostic

For each M6 counterbore feature-owned face:

1. log face index, surface type, cylinder radius, and edge count;
2. for every edge, log `IsCircle`, curve-parameter range, endpoint distance,
   and complete-boundary result;
3. log `TransformPointToView` result and proof string;
4. log direct `IView.GetCorrespondingEntity` result/type/error;
5. log the result of component-context correspondence before view mapping;
6. enumerate `GetVisibleEntities2(component, Edge)` and log drawing edge
   geometry;
7. compare model-owned and visible drawing edges using COM identity where
   valid plus projected centre/radius/axis/closure evidence;
8. report an explicit rejection reason for every attempted route; and
9. do not call the ordinate API until exactly six physical locations resolve.

Rerun acceptance:

- six locations;
- two unique X coordinates;
- three unique Y coordinates;
- selectable stepped-bore centre X datum;
- selectable lower Y datum;
- datum first in each selection;
- exact appended counts;
- `AddOrdinateDimension=0` twice;
- nonzero ordinate readback; and
- zero selected objects after `SetPickMode`.

## 6. Section dimension findings

The section contained 17 imported display dimensions.

The exact diameter records are:

| Index | Selection identity | Type2 | Nominal | Tolerance result |
|---:|---|---:|---:|---|
| 3 | `D1@Sketch4@...@Drawing View3` | 6 | `0.047 m` | none |
| 6 | `D1@Sketch6@...@Drawing View3` | 6 | `0.040 m` | none |

For `Ø47`:

```text
toleranceType=0
fitType=-1
fitDisplayStyle=-1
holeFit=
shaftFit=
fitValues=
minimumStatus=1
minimumM=0
maximumStatus=1
maximumM=0
```

The full-sheet output renders `Ø47.00`, not
`Ø47 H7 +0.025/0.000`.

`diametricLinearFound=False` is not evidence that the diameters are invalid.
The installed build imported them as standard
`swDiameterDimension=6`. Production R23 must support this actual type.

### 6.1 Probe labelling defect

The section loop does not reset per-iteration locals. In VBA, block-declared
locals retain procedure scope and their values persist across loop iterations.

Observed corruption:

- index 1 sets `FIRST_LINEAR`, and index 2 is also labelled `FIRST_LINEAR`;
- index 3 sets `DIAMETER_47`, and indices 4-5 retain that label;
- index 6 sets `DIAMETER_40`, and indices 7-17 retain that label;
- the nominal value also remains stale when a later nominal read fails.

The exact index 3 and 6 records are valid. Later target labels and targeted
geometry dumps are not.

Required correction:

```text
nominalM = 0
nominalAvailable = False
targetName = ""
dimensionH7Proven = False
toleranceSummary = ""
```

at the start of every iteration, plus equivalent reset for any new
per-dimension field.

Rerun acceptance:

- exactly one `DIAMETER_47`;
- exactly one `DIAMETER_40`;
- one deliberately chosen linear target;
- no stale target on unrelated dimensions;
- unavailable nominal logged as unavailable and zero/not-applicable rather
  than carrying the previous value.

### 6.2 H7 authority gate

The next corrected probe must read the original part-source `D1@Sketch4`
dimension directly, not only the imported drawing dimension.

Required readback:

- full dimension identity;
- nominal;
- tolerance type;
- fit type/display style;
- hole/shaft fit strings;
- `GetToleranceFitValues`;
- min/max values and status;
- prefix/suffix/callout text if present; and
- active/reference configuration.

If the part source is also empty, the user must decide whether the approved
target specification is allowed to supply H7. Production must not claim model
provenance for a reference-supplied tolerance.

## 7. J-J geometry findings

The flattened readback was structurally complete:

```text
lineCount=1
reportedSize=49
arrayItems=49
segmentCount=3
parseStatus=SUCCESS
```

This proves only the payload structure.

### 7.1 Returned geometry

Segment values:

```text
segment 1: (-0.124728, 0.000000) -> (0.008485, 0.000000)
segment 2: ( 0.008485, 0.000000) -> (0.008485,-0.024465)
segment 3: ( 0.008485,-0.024465) -> (0.124728,-0.024465)
```

Page-positioned arrow/label values:

```text
upper arrow: (0.107932,0.287788) -> (0.095932,0.287788)
lower arrow: (0.083467,0.042332) -> (0.071467,0.042332)
upper label: (0.104161,0.296839)
lower label: (0.079696,0.039318)
text height: 0.006350
```

The screenshot confirms:

- the upper arrow reaches the inner top border;
- the upper J label enters the top border/zone-number region;
- the lower arrow and J label enter the oversized part-identification band;
- the section label/dimension area is crowded.

### 7.2 Coordinate-frame defect

The source creates section segments by:

1. choosing page coordinates;
2. converting them with `SheetPointToViewSketchPoint`; and
3. creating view-owned sketch lines.

The returned segment values correspond to that source-view sketch frame.
Arrow and label positions correspond to the drawing page.

`Module6_QAEngine.ValidateSectionLineInfo` currently sends returned segment
values directly to a function that compares them with page-coordinate
part-identification bounds. This mixes frames and can produce a false
clearance result.

### 7.3 Required corrected diagnostic

The next probe must log:

- original page path;
- source-view sketch path;
- raw flattened segment path;
- returned segment path transformed back to page coordinates;
- arrow and label page coordinates;
- sheet size, content border, zone extents, part-ID bounds, title-block bounds,
  view outlines, and annotation envelopes; and
- a clearance result for each segment, arrow, and label.

Every log field must carry `frame=Page`, `frame=ViewSketch`, or another exact
frame label.

The transformation must occur exactly once. Do not add a view position twice.

## 8. Decisions fixed versus still open

### Fixed

- use `swInsertDimensions=8` within the expanded import mask;
- keep `DuplicateDims=True`;
- use deterministic selected-view `AllViews=False`;
- include non-Hole-Wizard cut definitions;
- do not use `CircleParams` as a production dependency;
- import native Hole Wizard callouts first;
- use controlled fallback only when native content is absent and source intent
  is proved;
- imported linear dimensions are not ordinate coverage;
- standard diameter type 6 is valid for the section; and
- replace the outline-extension J-J strategy.

### Open

- exact working model-edge to selectable drawing-edge correspondence route;
- exact page conversion for returned section sketch segment data;
- whether the underlying part dimension contains H7;
- if not, whether the target specification is an approved tolerance authority;
- final safe J-J arrow/label endpoints; and
- production implementation, deployment, and full regression.

## 9. Phase 0 stop condition — SATISFIED 2026-07-31

All seven conditions below are met by the corrected-probe runs recorded in
section 0. Production R23 Phase 1 is unblocked. See section 0.4 for the
closing evidence.

Production R23 work remains blocked until corrected live evidence proves:

1. six ownership-backed counterbore projections;
2. both selectable datum entities;
3. successful X and Y ordinate transactions with cleanup;
4. exact, non-stale section target classification;
5. direct model-source tolerance readback and an approved H7 authority;
6. page-coordinate J-J segment, arrow, and label readback; and
7. truthful clearance against every reserved region.

Do not rerun the current probes unchanged. The ordinate run will stop at the
same opaque mapping failure, and the section run will repeat the stale labels.

