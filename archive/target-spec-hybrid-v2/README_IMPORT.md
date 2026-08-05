# Target-Spec Hybrid V2

Source identity: `target-spec-hybrid-v2-2026-08-05-r62`

R23 production wiring is present in source and passed static contracts and
deployment preflight only. Use `tools/probe-runner/Run-R23Probes.ps1` for
read-only deployment, VBE compilation, and probe evidence; do not replace it
with manual Immediate-Window collection. Full-project VBE compilation,
mutating-run authorization, runtime regression, screenshot review, and
manufacturing acceptance remain separate gates.

After the one-time bootstrap is installed, use
`tools/swp-deploy/Deploy-TargetSpecHybrid.ps1` to synchronize the replaceable
standard/class modules into `Fable.swp`. See `tools/swp-deploy/README.md` for
the guarded candidate, compile, readback, and promotion workflow.

R22 combines r20 with the valid review corrections from r21 and replaces the
three low-confidence implementations that did not satisfy the documented API
contracts. A closed edge is accepted only when `ICurve.IsCircle=True`; a
cylindrical owning face no longer substitutes for circular-edge proof. Pattern
recognition uses the exact `IFeature.GetTypeName2` literals from the 2025 table,
including `APattern`, `LocalChainPattern`, `DimPattern`, derived-hole, and
sketch-pattern families. If `AlignDimensions` returns False, ordinate
dimensions use the installed, set-aware
`IDisplayDimension.AutoJogOrdinate`; manual `SetPosition2` lanes are used only
for non-ordinate dimensions, so unrelated ordinate sets are neither merged nor
split by per-member proximity.

R22 retains the verified r21 corrections for projected-origin containment,
FACE/SIDE callout separation, annotation self-identity, controlled general
notes, section-line parsing robustness, selection readback, title-region
measurement, and diagnostic volume. See
`docs/R22_REVIEW_RESOLUTION.md` for the contract and review record.

R20 keeps the runtime-proven J-J section, portrait primary, selected-view model
import, isometric isolation, P-0251 reference-led layout, metric mass
calculation, and model-first entity correspondence. It repairs the r19
zero-candidate path by proving complete cylinder-boundary edges with
`IEdge.GetCurveParams3` and using authoritative `ISurface.CylinderParams` when
the trimmed edge's underlying `ICurve.IsCircle` predicate is false. COM Boolean
results are normalized before logic, model-to-view page coordinates are no
longer translated twice and must prove they lie inside the current view outline,
dimension auto-arrange is a required stage with a deterministic lane fallback,
annotation QA uses the measured zoned border rather than a full-width
title-block band, J-J geometry is checked against the measured part-ID extent,
and the P-0251 controlled definitions are rendered as three leadered callouts
attached to ownership-proven drawing edges using SOLIDWORKS symbol tokens.
Existing-note reuse requires the complete normalized controlled definition, so
an imported Hole Wizard note containing only a short shared phrase is rejected.
Set
`DIAGNOSTIC_DRAWING_MODE = False` only after the focused r22 run and full
three-fixture acceptance matrix pass.

This folder is a clean replacement VBA source set derived from
`docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`. It does not patch or
overwrite `src/active-ordinate/active_ordinate.swp`, and it does not modify the
protected baseline.

## What is implemented in the exported r22 source

The following paths exist in exported source, have E2 source-contract coverage,
and use separately validated E3 signatures/enums. The guarded deployment proves
managed-source synchronization and revision readback, but its bootstrap
execution probe is not VBA Editor **Compile Project**. Manual full-project
compilation remains the E5 gate; E6 runtime and E7 manufacturing/visual
acceptance remain pending.

- strict authorization for the three test fixtures;
- one configured controlled-template existence check plus fail-closed created-
  sheet validation; there is no template discovery or uncontrolled fallback;
- explicit drawing-document activation with `ActivateDoc3` before view-scoped
  work;
- `ISheet.SetScale` plus actual-scale readback, implemented offline and awaiting
  E5/E6;
- fixture-aware orthographic roles without an unconditional view rotation;
- measured view layout above a reserved title/note band;
- one fixture-led primary section, including the P-0251 offset J-J path and
  Pump Holder B-B intent;
- mandatory Pump Holder Details C and D from reference-led normalized circular
  profiles on the exact created `*Bottom` view at independent 3:1 scale; E4/E6
  must prove coordinate/selection/creation behavior and E7 must prove the
  intended `7 x 4`/`C0.5` content and legibility;
- fixed hybrid annotation import with whole-drawing attempt and selected-view
  retry;
- audited-feature face and internal-cylinder edge enumeration with exact
  model-to-drawing-view correspondence;
- model-feature-to-face ownership proof before any ordinate candidate;
- full-circle, cylinder-radius, active-configuration, axis-normal, and stable
  semantic-instance checks;
- complete-curve parameter proof plus explicit `ICurve.IsCircle` and
  `CircleParams` proof; non-circular cylinder trims fail closed;
- typed selectable datum proof with direction-specific P-0251 Center-X and
  Bottom-Y behavior;
- family/view/datum-scoped linear-and-ordinate coverage reconciliation;
- datum-first ordinate selection, checked selection counts, decoded return
  codes, and pick-mode cleanup;
- configuration-first title properties, truthful part identification, and
  fail-closed QA evidence;
- a required, fail-closed auto-arrange stage with view-scoped selection and
  `AlignDimensions` readback; when the API returns `False`, ordinate sets use
  `AutoJogOrdinate` while non-ordinates use border-safe deterministic lanes;
- runtime proof that transformed page points lie within the current
  `IView.GetOutline`;
- measured J-J segment, arrow, and label clearance from the part-ID note; and
- three controlled P-0251 associative callouts covering the stepped bore, six
  counterbored face holes, and four tapped side holes. Each uses
  `<MOD-DIAM>`, a visible leader, attachment readback, collision checks, and
  complete-definition matching before any existing note is reused.

## Required references

The VBA project must reference:

- SOLIDWORKS 2025 Type Library;
- Microsoft Forms 2.0 Object Library.

`Scripting.Dictionary` is late-bound, so no Microsoft Scripting Runtime
reference is required.

## Important form and host-component note

`UserForm1.frm` and `UserFormSection.frm` in this folder are code snapshots,
not native VBE designer exports. A blank SOLIDWORKS macro also already owns its
special `ThisLibrary` host component.

For a clean project:

1. retain the host project's existing `ThisLibrary` and paste the supplied code
   beginning at `Option Explicit` into that component;
2. create two blank UserForms named `UserForm1` and `UserFormSection`;
3. paste each supplied form's code beginning at `Option Explicit` into the
   corresponding form code window; and
4. do not attempt to import `ThisLibrary.cls` as an ordinary class or treat the
   two form snapshots as native designer files.

The forms create their visible controls dynamically at runtime, so no manual
control placement is required after the blank forms exist.

## Import into a copy of the current macro

1. Make a recoverable copy of the current `.swp` file.
2. In the copy, remove or replace every existing same-named standard module.
   As of R23 that is `Module1_Main` through `Module9_LayoutEngine` plus
   `Module10_SectionDimensionEngine` and `Module11_GeometryIdentity` through
   `Module19_SemanticQA`. Leave unrelated modules untouched.

   The authoritative list is `tools/swp-deploy/deployment-manifest.json`;
   this prose is a convenience copy and the manifest wins.
3. Import every class module listed in the deployment manifest. As of R23
   that is `CHoleCandidate`, `CDatumProof`, `CRunEvidence`,
   `CFeatureDefinition`, `CPhysicalHoleLocation`, `CViewHoleProjection`,
   `CImportedAnnotation`, `CLocationGraph`, `COrdinateScheme`,
   `COrdinateBucket`, `CCalloutDefinition`, `CSectionPath`,
   `CSectionRequirement` and `CContentEnvelope`.
4. Import every standard module listed in the manifest.
5. Keep the current forms and handler classes only if their code matches the
   r4 snapshots. Otherwise replace the form code using the procedure above and
   replace the handler classes.
6. Confirm the inventory against `tools/swp-deploy/deployment-manifest.json`
   - the deployment preflight prints the managed-component count - plus both
   UserForms, all three handler classes, and the project-owned `ThisLibrary`
   host component, none of which the manifest manages.
7. Use **Debug > Compile VBAProject** and stop at the first highlighted error.

## Import into a blank macro project

Import every class and standard module named in
`tools/swp-deploy/deployment-manifest.json`, plus the three handler classes.
Then use the blank-form and host-component paste procedure above.

## First focused run (E6/E7)

R5 is locked to the configured candidate template at
`V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`; it does not
discover alternatives. The user-provided first-run screenshot and a read-only
path check confirmed that `.DRWDOT` at this corrected location. Module7's
property/link/cell map remains a candidate implementation pending live
template-linked evidence.

The approved template must contain a visible zoned border, a valid
`ITitleBlock`, nonzero zone margins, and visible linked notes for the required
drawing properties. The macro fails closed instead of falling back to an
uncontrolled default template.

Use only `P-0251-14A-001.SLDPRT` with:

- fixture acceptance profile (Front, Left, and one isometric role);
- isometric enabled;
- the fixture UI datum setting, with r22 resolving P-0251 as Center-X plus
  Bottom-Left-Y; both selectable entities and transforms require E4/E6 proof;
- Hole Wizard callouts enabled;
- 1:1 requested scale;
- exactly one horizontal section labeled `J`;
- title, notes, part identification, auto-arrange, and QA enabled.

The layout-preview control is intentionally disabled in r22 because this source
does not yet implement a separate preview/continue/cancel transaction.

The run writes `QA_REPORT.txt` in a unique run folder under
`test_assets/iteration_evidence/macro_qa/`. Preserve the complete Immediate
Window output, the report, the macro settings, and a full-sheet screenshot. A
successful compile or a nonzero API count is not visual acceptance.

A sectionless P-0251 run may be used only as an explicitly labelled diagnostic
that is expected to fail fixture QA for the missing J-J section.

## Offline validation already completed

Run the complete project-local hybrid companion suite immediately before
import. It currently passes **424/424 tests**. Static success remains distinct
from VBE compilation, runtime evidence, and manufacturing acceptance.
(see the handoff, section 5.2); those
source-contract, structure, regression, and fake-COM checks establish E2 only.
Separate installed-interop
and documentation checks establish E3. The guarded deployment at
`test_assets/iteration_evidence/swp_deployment/20260728_142300/` is historical
r20 evidence; the current r22 deployment evidence is reported separately. Its
`COMPILE_PROBE|status=SUCCESS` output is only a bootstrap execution check, not a
full-project compile; `compile-probe-scope.txt` records that limitation. None of
these checks substitutes for VBA Editor **Compile Project**, a SOLIDWORKS
drawing run, and visual comparison.

The r4 API/source checkpoint and SHA-256 manifest are in
`test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/`.
The older `2026-07-18_target_spec_hybrid_v2_offline` directory is historical
r3 evidence and must not be used to identify this r4 source.
