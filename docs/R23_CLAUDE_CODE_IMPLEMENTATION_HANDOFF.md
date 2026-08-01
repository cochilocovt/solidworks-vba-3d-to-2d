# Claude Code Handoff: R23 Drawing Automation

**Prepared:** 2026-07-31

**Repository:** `C:\Users\V.T\Documents\VBA 3D TO 2D`

**Branch / commit at handoff:** `main` / `ffe2de791f6e1587bed00bc4fa30afc2874c27ec`

**Current production source revision:** `target-spec-hybrid-v2-2026-07-29-r22`

**Purpose:** Give Claude Code the complete, durable context needed to finish
the diagnostic Phase 0 gate and then implement R23 without reconstructing the
preceding Codex session or repeating disproved approaches.

## 1. Start here

Read these files in order before changing anything:

1. [Agents.md](../Agents.md) - binding authorization, protected-asset,
   deployment, API, validation, and live-operator rules.
2. [CLAUDE.md](../CLAUDE.md) - Claude Code repository entry point.
3. This handoff.
4. [R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md](R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md)
   - exact live evidence, probe defects, and rerun acceptance criteria.
5. [R23_IMPLEMENTATION_PLAN.md](R23_IMPLEMENTATION_PLAN.md) - ordered R23
   production task list. Its unchecked tasks remain unchecked unless the
   evidence stated by the task exists.
6. [CURRENT_STATUS.md](CURRENT_STATUS.md) - current revision and validation
   state.
7. [REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md](REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md)
   - product and manufacturing-drawing acceptance requirements.
8. [CODESTACK_DRAWING_API_COVERAGE.md](CODESTACK_DRAWING_API_COVERAGE.md) and
   [3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md](3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md)
   - practical drawing API patterns. These are secondary guidance; current
   source, official API remarks, installed SOLIDWORKS 2025 behavior, and live
   evidence take precedence.
9. [SOLIDWORKS_API_VALIDATION.md](SOLIDWORKS_API_VALIDATION.md) - accumulated
   API and installed-interop evidence.
10. [CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md](CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md)
    - evidence ladder, static checks, deployment workflow, and claim language.

Do not begin by editing production VBA. First reproduce the repository state,
read the final Phase 0 evidence, and confirm that the corrected diagnostic work
described below is still outstanding.

## 2. Goal

Create R23 of the SOLIDWORKS 2025 VBA macro so that its drawings are
manufacturing-usable and match the approved target specification, rather than
merely completing with a nonzero dimension count.

For P-0251 specifically, R23 must provide:

- feature-definition coverage for Hole Wizard and non-Hole-Wizard cuts;
- the required reference-style X/Y ordinate scheme for the six counterbores;
- the six-hole M6 counterbore definition;
- complete M5 side-hole definition and locations;
- a stepped-bore section with real associative dimensions, including
  `Ø47 H7 +0.025/0.000` and `Ø40`;
- one deterministic J-J section line that crosses the intended features and
  clears the sheet border, zone labels, part-identification band, title block,
  views, and annotations;
- content-aware arrangement, title information, notes, barcode, and truthful
  fail-closed QA; and
- equivalent fixture-led completeness for the other two authorized models.

Passing compilation, returning 25 imported items, or displaying a `PASS`
dialog is not the goal. The final drawings must pass visual and semantic
comparison with the manual references.

## 3. Non-negotiable boundaries

### 3.1 Authorized fixtures

Live analysis and execution are authorized only for:

- `test_assets/models/P-0251-14A-001.SLDPRT`
- `test_assets/models/P-0252-01-001.SLDPRT`
- `test_assets/models/P-0252-01-013.SLDPRT`

Never modify, rebuild for the purpose of changing design state, or save the
design/features/dimensions/configurations/properties of these models.

### 3.2 Protected assets

Never modify:

- `src/baseline-model-dims/`
- manual drawings under `test_assets/reference_drawings/`
- `src/active-ordinate/` unless the user explicitly requests historical work

The active production source is `src/target-spec-hybrid-v2/`.

### 3.3 Live operator

The user performs live SOLIDWORKS and VBA-editor work by default. The user has
also said to stop Computer Use and provide exact manual steps instead.

Prepare and verify source offline, then tell the user exactly what to compile,
run, retain, and share. Do not use Computer Use unless the user explicitly
re-authorizes it for a defined task.

### 3.4 Source and deployment

- Checked-in managed source is authoritative for production work.
- `Fable.swp` is the production target.
- `tools/swp-deploy/deployment-manifest.json` manages 15 standard/class
  components and deliberately excludes `ThisLibrary` and the UserForms.
- Deployable `.bas` and ordinary `.cls` sources must be Windows-1252/ANSI,
  CRLF, without a UTF-8 BOM or `Attribute` metadata.
- Back up `Fable.swp` before its first material R23 change.
- Use guarded deployment and exact source readback.
- `COMPILE_PROBE|status=SUCCESS` is not a VBE full-project compile.

The production SWP SHA-256 at this handoff remains:

```text
C8076D713DD8F64AC75F93871C1EA4A2D1F01EE2BC691323736013BFEB2803F2
```

The production source tree has not yet been changed to R23.

## 4. Working-tree warning

The handoff was created in an already dirty working tree. Existing changes and
untracked files belong to the user or the preceding workflow. Do not discard,
reset, overwrite, or broadly reformat them.

At handoff, `src/target-spec-hybrid-v2/` itself still identifies as R22.
Relevant uncommitted documentation and probe artifacts exist under:

- `docs/`
- `tools/r23-probes/`
- `test_assets/iteration_evidence/r23/20260730-075811/`

Always run:

```powershell
git status --short --branch
git diff --stat
git diff --name-status
git diff --check
```

before deciding what belongs to the next change.

Do not hand-edit `tools/swp-deploy/deployment-request.txt`. It is generated by
the deployer and currently records the last disposable-probe deployment.

## 5. What R22 resolved

Commit `ffe2de7` is the current production baseline. It consolidated the
R20/R21 line and corrected the substantive review defects discussed in
`docs/R21_REVIEW_HANDOFF.md` and `docs/R22_REVIEW_RESOLUTION.md`.

Important retained lessons:

- a cylindrical face does not prove that an arbitrary trimming edge is a
  circle;
- `ICurve.IsCircle=False` must fail closed;
- exact SOLIDWORKS feature type names and underlying `ICE` resolution matter;
- ordinate transactions must select the datum first, append feature entities,
  check the return code, call `SetPickMode`, and clear selections;
- imported-dimension count is not semantic coverage;
- arrangement success requires readback, not a setter call alone; and
- QA must distinguish a root failure from downstream missing output.

R22 is the source baseline, not an accepted manufacturing result. Its P-0251
output lacked the required ordinates, callouts, section tolerances, and safe
section-line placement.

## 6. Phase 0 evidence already proved

### 6.1 Feature and curve probe

Authoritative log:

`test_assets/iteration_evidence/r23/20260730-075811/live-probes/R23_FEATURE_20260731_040539.log`

It proved:

- 47 features were visited;
- `Cut-Extrude1`, `Cut-Extrude3`, and `Cut-Extrude4` expose
  `GetTypeName2="ICE"` and underlying `GetTypeName="Cut"`;
- typed `IExtrudeFeatureData2` access/release and owned geometry are available;
- the native M6 counterbore Hole Wizard has six locations;
- the native M5x0.8 Hole Wizard and `Mirror1` seed relationship are readable;
- cylinder radius plus `Curve.IsCircle` and complete
  `Edge.GetCurveParams3` evidence are stable; and
- the probe's `CircleParams` helper was anomalous even while `IsCircle`
  remained true.

Production must not depend on `CircleParams`. Use the live-proven cylindrical
surface, `IsCircle`, and complete-boundary parameter evidence.

`IsSuppressed2` returned Empty in this VBA run. The probe's separately labelled
active-document `IsSuppressed` fallback proves only the active configuration.
Production must still prove or fail closed on each drawing view's
`ReferencedConfiguration`.

### 6.2 Import transaction probe

Authoritative logs:

- `R23_IMPORT_AllViewsTrue_20260731_041602.log`
- `R23_IMPORT_SelectedViewsFalse_20260731_041721.log`

Both are under:

`test_assets/iteration_evidence/r23/20260730-075811/live-probes/`

Fixed production decision:

- retain `DuplicateDims=True`;
- retain the expanded mask including `swInsertDimensions=8`;
- call `InsertModelAnnotations4` deterministically per selected view with
  `AllViews=False`;
- call section, side, then primary; and
- leave the isometric unselected and undimensioned.

Observed results:

| Transaction | Primary | Side | Section | Isometric | Unique total |
|---|---:|---:|---:|---:|---:|
| selected primary, `AllViews=True` | 8 | 0 | 17 | 0 | 25 |
| section/side/primary, `AllViews=False` | 6 | 2 | 17 | 0 | 25 |

The per-view transaction is required because `AllViews=True` omitted all side
dimensions.

The import also proved:

- one native M5 Hole Wizard callout in the section;
- no native M6 counterbore callout;
- no imported H7 fit;
- no nonzero imported tolerance; and
- imported linear dimensions do not satisfy the required ordinate scheme.

### 6.3 Final ordinate and section probes

These runs are documented in detail in
[R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md](R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md).

Short version:

- the ordinate probe failed before selecting anything or calling
  `AddOrdinateDimension`;
- the section probe found real imported `Ø47` and `Ø40` diameter dimensions,
  but no H7;
- the J-J raw payload parsed exactly, while its arrows/labels visibly violated
  the upper border/zone and lower part-identification areas; and
- the section probe itself retained stale loop variables and mislabeled many
  later display-geometry records.

Do not describe either final probe as accepted. The section result
`evidenceStatus=CAPTURED` means evidence was captured, not that the drawing
passed.

## 7. Probe defects that must be corrected first

This is the next allowed coding scope. Correct the disposable probes before
implementing production R23.

### 7.1 Ordinate probe: opaque mapping failure

Source:

`tools/r23-probes/import-transaction-source/Module_R23Phase0DrawingProbes.bas`

The run reached `GetVisibleComponents` and found
`P-0251-14A-001-1`, then emitted:

```text
R23_ORDINATE_FATAL|reason=CounterboreLocationsUnavailable
```

It produced no per-face/edge rejection records. Therefore the current log
cannot distinguish:

- no owned edge passing `IsCircle` plus closure;
- `TransformPointToView` rejection;
- direct `IView.GetCorrespondingEntity(modelEdge)` returning Nothing;
- a component-context entity being required before view correspondence; or
- a usable edge existing only in `GetVisibleEntities2(component, Edge)`.

Required probe correction:

1. Log every counterbore owned cylindrical face with radius and edge count.
2. For every owned edge, log:
   - runtime type;
   - `Curve.IsCircle`;
   - `GetCurveParams3` availability, parameter range, and endpoint closure;
   - centre transform result and coordinate proof;
   - direct view-correspondence result and returned runtime type.
3. Compare these mapping routes explicitly:
   - active-part model edge -> `IView.GetCorrespondingEntity`;
   - active-part model edge -> `IComponent2.GetCorrespondingEntity` ->
     `IView.GetCorrespondingEntity`;
   - drawing edges returned by
     `IView.GetVisibleEntities2(component, swViewEntityType_Edge)`.
4. Log stable geometric/COM evidence that relates the accepted drawing edge
   back to the feature-owned model edge.
5. Stop before ordinate creation unless exactly six physical counterbore
   locations resolve into two X and three Y coordinates with selectable
   drawing entities.
6. Only then run the datum-first X and Y creation transaction and prove:
   - datum selection order zero;
   - exact appended selection counts;
   - `AddOrdinateDimension=0`;
   - display-dimension readback;
   - `SetPickMode`; and
   - final selection count zero.

Do not weaken ownership qualification merely to obtain six visible circles.

### 7.2 Section probe: stale per-iteration values

In VBA, declarations inside `If` and loop blocks remain procedure-scoped.
The probe declares but does not reset:

- `nominalM`;
- `targetName`; and
- `dimensionH7Proven`.

As a result:

- index 2 remained labelled `FIRST_LINEAR`;
- records after exact `Ø47` remained labelled `DIAMETER_47`; and
- records after exact `Ø40` remained labelled `DIAMETER_40`.

Required correction:

- set `nominalM = 0#`;
- set `targetName = vbNullString`;
- set `dimensionH7Proven = False`; and
- reset other per-item values

at the start of every display-dimension iteration.

Dump targeted display geometry only for the exact semantic match. The rerun
must produce exactly one `DIAMETER_47`, exactly one `DIAMETER_40`, and one
explicitly chosen representative linear target.

### 7.3 Section probe and production QA: mixed coordinate frames

The J-J flattened payload contains:

- section sketch segment endpoints in source-view sketch coordinates; and
- arrow/label positions in drawing-page coordinates.

The existing production parser in `Module6_QAEngine.ValidateSectionLineInfo`
feeds the segment values directly into page-coordinate part-identification
rectangles. That comparison is not trustworthy.

Evidence supporting the mixed-frame conclusion:

- `CreatePrimarySection` starts from page coordinates;
- `AddSectionSegment` converts them to source-view sketch coordinates before
  `CreateLine`;
- returned segment values match the view-sketch geometry;
- returned arrow and label coordinates match their visible page positions; and
- the screenshots visibly agree with the arrow/label page coordinates.

Required correction to the diagnostic path:

1. Record the original page path before section creation.
2. Record the converted view-sketch coordinates passed to `CreateLine`.
3. Read the flattened payload.
4. transform each returned segment endpoint back to page coordinates exactly
   once;
5. log both frames explicitly; and
6. compare only page-coordinate values against page-coordinate borders,
   part-identification extents, title block, views, and annotation envelopes.

Do not change J-J production geometry until the conversion and clearance
readback are proved.

## 8. Live findings that change the R23 plan

### 8.1 Section diameter type

The imported `Ø47` and `Ø40` dimensions are valid
`swDiameterDimension=6` records:

- `D1@Sketch4`, nominal `0.047 m`;
- `D1@Sketch6`, nominal `0.040 m`.

No `swDiametricLinearDimension=15` record was returned. R23's section engine
must support the real type 6 dimensions rather than require type 15.

### 8.2 H7 provenance is unresolved

The imported `Ø47` record reports:

```text
toleranceType=0
fitType=-1
holeFit=
fitValues=
minimumM=0
maximumM=0
```

The output visibly renders `Ø47.00`, not H7.

Before production implementation, probe the underlying part-source dimension
directly to distinguish:

1. H7 exists in the model but import/display lost it; or
2. H7 exists only in the approved reference/target specification.

Then obtain the user's policy decision:

- **model-authoritative:** fail closed until the model contains H7; or
- **controlled-target-spec authority:** create/reuse an associative `Ø47`
  dimension, apply `H7 +0.025/0.000` from an explicit approved fixture
  requirement, and record provenance as target-spec/reference authority.

The previous R23-806 wording allowed H7 only when model data proved it. Do not
silently change that policy. The recommended route for meeting the approved
target is controlled-target-spec authority with explicit provenance, but it
requires the user's approval.

### 8.3 J-J construction strategy

The old outline-percentage and extension strategy is disproved. Its upper
label enters the zone region and its lower arrow/label enter the
part-identification band.

R23 must build J-J from the proved location graph:

1. stepped-bore centre;
2. same X at the highest face-hole row;
3. minimum-X face-hole column at that row;
4. same column at the lowest face-hole row.

The implementation must prove that the path crosses the stepped bore and all
three intended holes, then choose safe arrow/label endpoints within the
content area. It must not extend blindly to `topY + extension` or
`bottomY - extension`.

## 9. How to implement R23 after Phase 0 closes

Follow the phase order in [R23_IMPLEMENTATION_PLAN.md](R23_IMPLEMENTATION_PLAN.md).
The intended architecture is:

### 9.1 Feature and location graph

Add typed records for:

- feature definitions;
- physical manufacturing locations;
- per-view projections/selectable anchors;
- imported annotation identities/provenance; and
- the overall location graph.

Normalize `ICE` through `GetTypeName`. Support Hole Wizard, resolved exact cut
features, mirrored/patterned features with proved seeds, and ordinary modeled
cuts. Use `IFeature.GetFaces` ownership rather than accepting visible circles
by appearance.

Coaxial counterbore steps belong to one physical location. Mirrored or
patterned instances remain distinct physical locations.

### 9.2 Drawing projections

Keep component contexts explicit:

- drawing-limited component from `GetVisibleComponents` for
  `GetVisibleEntities2`;
- full model/component context where feature/body ownership requires it; and
- drawing-view entities used for selection and annotation attachment.

Every accepted projection needs model feature provenance, physical identity,
view identity, configuration proof, geometry proof, and a selectable drawing
entity.

### 9.3 Import and reconciliation

Use the proved per-view `InsertModelAnnotations4` transaction. Build a semantic
ledger instead of treating the returned count as coverage.

Imported linear dimensions do not cover ordinate requirements. Imported
diameters can cover section-diameter geometry if their source, nominal,
attachment, and type match. Imported Hole Wizard callouts are first priority,
but absence triggers only an evidence-backed controlled fallback.

### 9.4 Ordinates

For P-0251:

- X zero: proved stepped-bore centre;
- Y zero: proved lower reference/datum entity;
- feature coverage: six counterbore locations represented as two X and three
  Y coordinates;
- select datum first;
- append only unique feature projections;
- check exact selection counts;
- decode API return;
- read back created ordinate dimensions; and
- always finish with `SetPickMode` plus zero-selection cleanup.

### 9.5 Callouts

Use native Hole Wizard callouts when present and complete. Otherwise construct
controlled associative callouts only from typed source evidence.

P-0251 requires:

- six-hole M6 counterbore definition;
- four M5x0.8 side-hole definition; and
- stepped-bore/fit definition in the section.

Free text without geometry attachment is not manufacturing evidence.

### 9.6 Section dimensions

Add a dedicated semantic section-dimension engine. Reconcile imported
dimensions first, then create only missing associative dimensions.

Support the actual type 6 diameter records. Read back nominal, type, source,
attachment, tolerance, fit, deviations, and placement.

Do not substitute a free-text H7 label for a failed real dimension.

### 9.7 J-J and layout

Construct J-J from graph locations, convert coordinate frames exactly once,
and verify the returned page geometry after creation and after every view move.

Arrange dimensions, callouts, section labels, arrows, notes, barcode, and
title information using full display envelopes. Reserve border zones,
part-identification region, and title block as rectangles.

### 9.8 QA

Replace count-based success with requirement keys. Each required item needs:

- source/provenance;
- view/attachment;
- semantic match;
- API result/readback;
- placement/extent; and
- final visible coverage.

Fail closed on any missing or unproved requirement.

## 10. Exact next work package for Claude Code

Do this next:

1. Read the files in section 1 and inspect current source/diffs.
2. Do not change `src/target-spec-hybrid-v2/`.
3. Correct only the disposable Phase 0 probe sources and their documentation.
4. Before changing any SOLIDWORKS API call or constant, read
   `skills/solidworks-api-lookup/SKILL.md` and verify the contract through the
   local MCP plus installed SW2025 interop where load-bearing.
5. Add offline structural tests or targeted source checks for:
   - per-iteration state reset;
   - complete per-edge mapping diagnostics;
   - explicit coordinate-frame labels/conversion; and
   - no production-path mutation.
6. Run the focused offline tests, complete offline suite, source hygiene,
   `git diff --check`, and disposable-manifest `-PreflightOnly`.
7. Give the user exact deploy/compile/run steps.
8. Wait for corrected logs and uncropped screenshots.
9. Analyze the evidence and update the Phase 0 findings.
10. Ask for the H7 authority decision if the direct part-source probe confirms
    that H7 is absent.
11. Begin production Phase 1 only when:
    - six counterbore projections and both datum entities are selectable;
    - both ordinate groups return success and clean up;
    - exact `Ø47`/`Ø40` records are isolated;
    - H7 authority is settled; and
    - all J-J geometry is in a proved page frame with truthful clearance
      results.

## 11. Verification commands

### 11.1 Complete offline suite

From the repository root:

```powershell
& .\tools\solidworks-automation-companion\.venv\Scripts\python.exe `
  -m unittest discover `
  -s .\tools\solidworks-automation-companion\tests `
  -p 'test*.py' `
  -v
```

Known baseline:

- 74 tests run;
- 69 pass;
- five failures remain in stale R20 string/behavior contracts.

Do not call this green. Investigate any additional, missing, or changed
failure.

### 11.2 Disposable drawing-probe preflight

```powershell
powershell -ExecutionPolicy Bypass `
  -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1 `
  -ManifestPath .\tools\r23-probes\drawing-contract-manifest.json `
  -PreflightOnly
```

This must not target production `Fable.swp`.

### 11.3 Production preflight

Only after production R23 work is authorized and prepared:

```powershell
powershell -ExecutionPolicy Bypass `
  -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1 `
  -PreflightOnly
```

The manifest property is `sourceDirectory`, not `sourceRoot`.

## 12. Evidence and claim language

The authoritative Phase 0 evidence is indexed in
[R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md](R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md).

Always distinguish:

- source/static verification;
- API documentation or installed-interop verification;
- guarded SWP deployment/readback;
- VBE full-project compilation;
- live macro execution;
- drawing screenshot review; and
- manufacturing acceptance.

For a newly corrected disposable probe that has not yet been run, use:

> The disposable probe source passed static and deployment-preflight checks.
> The user must still compile and run it in SOLIDWORKS. Production R23 remains
> unchanged.

Good acceptance language:

> The corrected P-0251 run proved the specified requirement keys and visual
> clearances. The other authorized fixtures and full regression remain
> unproved.

Never say:

- "the macro is fixed" because source compiles;
- "the drawing passes" because 25 items were imported;
- "H7 was imported" when tolerance readback is empty;
- "ordinates failed" when `AddOrdinateDimension` was never called; or
- "section clearance passed" when only the flattened array structure parsed.

## 13. Collaboration model

Claude Code owns implementation and repository updates. The user owns live
SOLIDWORKS/VBE operation unless they explicitly delegate it.

Codex CLI may be used as an independent read-only evaluator after a coherent
change. Do not ask Codex to edit the same working tree as the implementation
agent. Use:

```powershell
codex review --uncommitted
```

or a scoped read-only `codex exec` prompt after the change and offline checks.
Claude Code and Codex have separate MCP/plugin/authentication state; verify
required MCP configuration independently.
