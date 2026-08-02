# SOLIDWORKS VBA Drawing Automation — Agent Instructions

## Purpose and Current Source of Truth

This repository contains a SOLIDWORKS 2025 VBA macro that creates manufacturing-oriented drawings from saved 3D part models. Success means a correct, readable, manufacturing-usable drawing, not merely a completed macro or a nonzero dimension count.

The authoritative managed source is `src/target-spec-hybrid-v2/`. It is deployed into `Fable.swp` through `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`; `tools/swp-deploy/deployment-manifest.json` defines the managed components. Treat `src/active-ordinate/` and `active_ordinate.swp` as historical evidence, not the current implementation or deployment target.

Use `docs/CURRENT_STATUS.md` for the current revision and validation state, `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md` for product and acceptance requirements, and `tools/swp-deploy/README.md` for deployment details.

## Authorized Fixtures and Protected Assets

The agent has continuing authority to analyze and run the macro only against:

- `test_assets/models/P-0251-14A-001.SLDPRT`
- `test_assets/models/P-0252-01-001.SLDPRT`
- `test_assets/models/P-0252-01-013.SLDPRT`

Within that boundary, the agent may inspect the models, edit the managed source, back up and deploy `Fable.swp`, compile and run the macro, diagnose output, and retain disposable drawings and evidence under `test_assets/`.

The following safeguards apply:

- Never modify or save the design, features, dimensions, configurations, or properties of an authorized model.
- Never modify `src/baseline-model-dims/` or overwrite manual drawings under `test_assets/reference_drawings/`.
- Do not modify the historical `src/active-ordinate/` snapshot unless the user explicitly requests historical repair.
- Make or verify a recoverable backup before the first material change to `Fable.swp`.
- Do not run the macro on any other model without new permission.

## Live SOLIDWORKS and Computer Use Operator Policy

The user performs live SOLIDWORKS and VBA-editor work by default. If Computer Use or another live access path is needed and the user has not already assigned the operator, ask once for the described live work package. That choice remains valid until its fixture, mutation scope, or objective changes.

If the user performs the work, provide concise exact steps and analyze the returned compile output, logs, QA report, and screenshots. If the agent is assigned the live work, remain within the authorized fixtures, never save model-design changes, and report what was inspected or changed. Stop immediately if the user withdraws live authorization.

Reading documentation through the local `solidworks-api` MCP is not live SOLIDWORKS access and does not require an operator choice.

## Source, Deployment, and API Rules

Checked-in managed source is authoritative. Before replacing embedded code, preserve, extract, and report any unexpected difference from `Fable.swp`; an incidental VBA-editor difference does not silently supersede reviewed source. After deployment, use the guarded source readback to confirm that every managed component and revision matches the repository.

Files managed by the deployment manifest must follow the deployer's source-hygiene contract: deployable `.bas` and ordinary `.cls` files are Windows-1252/ANSI without a UTF-8 BOM or `Attribute` metadata. `ThisLibrary`, UserForms, designer files, legacy exports, and other components outside the manifest are not covered by that blanket rule. Follow `tools/swp-deploy/README.md` for their handling.

Before changing or relying on a SOLIDWORKS API member, `sw*` enum or constant, COM binding, selection contract, coordinate behavior, or return-code interpretation, read and use `skills/solidworks-api-lookup/SKILL.md`. Confirm load-bearing signatures, Remarks, and numeric values against SOLIDWORKS 2025 evidence; the MCP compatibility snapshot is evidence rather than installed-build proof. Record material contract findings in `docs/SOLIDWORKS_API_VALIDATION.md`.

Pure documentation, diagnostic-text, and non-API refactoring work does not require an API lookup.

## Engineering Principles

Keep the fixed hybrid workflow: import applicable model-marked annotations, add evidence-qualified hole-centre or feature-location ordinates without unnecessary duplication, arrange dimensions when configured, populate title information, notes, and barcode, and produce truthful fail-closed QA. Do not reintroduce separate user-selectable model-dimension and ordinate modes unless requested. Support one deterministic primary section unless the user approves a data-model and UI redesign.

For VBA changes:

- Use `Option Explicit`, retain compile-safe public interfaces, and inspect all callers before changing shared declarations or signatures.
- Check API return values and errors; log enough context to identify the view, operation, return code, and VBA error.
- Keep model, view, and sheet coordinates explicit; SOLIDWORKS lengths are normally metres.
- Establish the correct active view before view-scoped selection, clear selections after scoped operations, and restore normal selection mode with `SetPickMode` after ordinate work.
- Prefer the smallest coherent correction. Do not redesign the whole macro, convert Python into a second production drawing engine, or remove required output stages to hide a defect.

## Validation and Acceptance

Use evidence appropriate to the work:

1. **Review or diagnosis:** perform read-only inspection and report evidence-backed findings. Do not imply runtime validation.
2. **Source change:** run relevant static tests, validate affected API contracts when applicable, update the source revision when deployable behavior changes, and complete guarded deployment/readback checks when the change is deployed.
3. **Production acceptance:** compile the full VBA project in the SOLIDWORKS VBA editor, run only authorized fixtures, capture complete Immediate Window output, `QA_REPORT.txt`, settings, and an uncropped full-sheet screenshot, then compare the drawing with the manual reference and target specification.

**Read-only probe exception (user-authorized 2026-08-02).** A strictly
read-only `R23_Probe*` entry point may be deployed and run without a
preceding manual **Debug > Compile Project**. A probe that fails to compile
fails loudly at its first statement, so the manual gate was buying nothing
there. The agent may also perform full-project compilation programmatically
through the VBE `CommandBars` route already used by
`tools/swp-deploy/Module0_SourceDeployment.bas`.

This exception is narrow and does not move item 3. Production acceptance,
and any run that mutates a drawing, still requires the manual full-project
compile in the VBA editor.

After a narrow correction, use the smallest focused fixture and settings matrix that can prove it. Before production acceptance, complete the full three-fixture regression and the applicable datum, Hole Wizard, arrangement, section, supported-view, duplication, layout, title, notes, barcode, and QA checks defined by the target specification.

Static tests, API reflection, deployment readback, a successful bootstrap probe, nonzero counts, or a `PASS` dialog do not prove embedded compilation, runtime correctness, or manufacturing acceptance. Preserve useful failure evidence under `test_assets/iteration_evidence/` and never overwrite manual references.

## Delivery and Evidence

For a source change, report the defect, root cause, evidence, affected modules and callers, files changed, and the verification actually completed. Distinguish static verification, API or probe evidence, full-project compilation, macro execution, and visual/manufacturing acceptance.

When files were edited in the workspace, summarize the changes and link to them; provide complete replacement source only when the user must paste it manually. Update `docs/Changelog.md` and `docs/CURRENT_STATUS.md` when behavior, source revision, deployment state, or acceptance status materially changes, not for every incidental edit.

Preserve unrelated user changes, remain within the requested scope, and state any uncompleted validation gate without presenting partial evidence as final success.
