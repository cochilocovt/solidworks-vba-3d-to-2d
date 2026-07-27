# Target-Spec Hybrid V2

Source identity: `target-spec-hybrid-v2-2026-07-18-r18`

After the one-time bootstrap is installed, use
`tools/swp-deploy/Deploy-TargetSpecHybrid.ps1` to synchronize the replaceable
standard/class modules into `Fable.swp`. See `tools/swp-deploy/README.md` for
the guarded candidate, compile, readback, and promotion workflow.

R18 keeps the runtime-proven J-J section, portrait primary, selected-view model
import, isometric isolation, P-0251 reference-led layout, and metric mass
calculation. It replaces the twice-disproved drawing-edge reverse lookup with a
model-first path: audited hole-feature faces and circular model edges are mapped
into each supported drawing view through `IView.GetCorrespondingEntity`.
Selectable datum vertices use the same model-to-view direction. Final QA now
uses each view's exact `GetDisplayDimensions` array instead of a sheet-wide
iterator. A visible mass note is relinked to `$PRP:"Mass"` only when exactly one
property-linked mass candidate exists and both link and rendered text read back
after rebuild. Set `DIAGNOSTIC_DRAWING_MODE = False` only after the controlled
template provides a real `ITitleBlock` and production validation is intended.

This folder is a clean replacement VBA source set derived from
`docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`. It does not patch or
overwrite `src/active-ordinate/active_ordinate.swp`, and it does not modify the
protected baseline.

## What is implemented in the exported r4 source (E2/E3 only)

The following paths exist in exported source and have E2 source-contract
coverage; used signatures and enums have separate E3 validation. None is yet
an E5 compile, E6 run, or E7 manufacturing/visual result.

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
- typed selectable datum proof with direction-specific P-0251 Center-X and
  Bottom-Y behavior;
- family/view/datum-scoped linear-and-ordinate coverage reconciliation;
- datum-first ordinate selection, checked selection counts, decoded return
  codes, and pick-mode cleanup;
- configuration-first title properties, truthful part identification, and
  fail-closed QA evidence.

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
2. In the copy, remove or replace every existing same-named standard module
   `Module1_Main` through `Module9_LayoutEngine`. Leave unrelated modules
   untouched.
3. Import these class modules:
   - `CHoleCandidate.cls`
   - `CDatumProof.cls`
   - `CRunEvidence.cls`
4. Import `Module1_Main.bas` through `Module9_LayoutEngine.bas`.
5. Keep the current forms and handler classes only if their code matches the
   r4 snapshots. Otherwise replace the form code using the procedure above and
   replace the handler classes.
6. Confirm the expected inventory: nine standard modules, `CHoleCandidate`,
   `CDatumProof`, `CRunEvidence`, both UserForms, all three handler classes, and
   the project-owned `ThisLibrary` host component.
7. Use **Debug > Compile VBAProject** and stop at the first highlighted error.

## Import into a blank macro project

Import the three data classes, all nine standard modules, and the three handler
classes. Then use the blank-form and host-component paste procedure above.

## First focused run (after E5 and D-04)

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
- the fixture UI datum setting, with r4 resolving P-0251 as Center-X plus
  Bottom-Left-Y; both selectable entities and transforms require E4/E6 proof;
- Hole Wizard callouts enabled;
- 1:1 requested scale;
- exactly one horizontal section labeled `J`;
- title, notes, part identification, auto-arrange, and QA enabled.

The layout-preview control is intentionally disabled in r4 because this source
does not yet implement a separate preview/continue/cancel transaction.

The run writes `QA_REPORT.txt` in a unique run folder under
`test_assets/iteration_evidence/macro_qa/`. Preserve the complete Immediate
Window output, the report, the macro settings, and a full-sheet screenshot. A
successful compile or a nonzero API count is not visual acceptance.

A sectionless P-0251 run may be used only as an explicitly labelled diagnostic
that is expected to fail fixture QA for the missing J-J section.

## Offline validation already completed

Run the complete project-local hybrid companion suite immediately before
import. It currently passes **62 tests**; those source-contract, structure,
regression, and fake-COM checks establish E2 only. Separate installed-interop
and documentation checks establish E3. Neither is a substitute for VBA
compilation or a SOLIDWORKS drawing run.

The r4 API/source checkpoint and SHA-256 manifest are in
`test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/`.
The older `2026-07-18_target_spec_hybrid_v2_offline` directory is historical
r3 evidence and must not be used to identify this r4 source.
