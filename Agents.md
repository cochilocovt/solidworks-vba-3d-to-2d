# Agent Instructions

SOLIDWORKS 2025 VBA macro that generates manufacturing drawings from saved 3D
parts. Success is a correct, manufacturable drawing — not a completed run or a
nonzero dimension count.

Trunk: `src/baseline-model-dims/`. Deployed via
`tools/swp-deploy/deployment-manifest.json`. `src/active-ordinate/` and
`archive/` are history, not deployment targets.

## Hard rules

- Run the macro only against the three fixtures in
  `test_assets/models/`: `P-0251-14A-001`, `P-0252-01-001`, `P-0252-01-013`.
  Any other model needs new permission.
- Never save a fixture model or overwrite anything in
  `test_assets/reference_drawings/`. Generated drawings are disposable.
- Back up `Fable.swp` before the first material change.
- Ask before live SOLIDWORKS work if the operator has not already been
  assigned. That choice holds until the fixture, mutation scope, or objective
  changes.

## Running it

Read-only evidence:

```bash
powershell -ExecutionPolicy Bypass -File ".\tools\probe-runner\Run-R23Probes.ps1" -Deploy -DrawingPath "<drawing path>"
```

Mutating run:

```bash
powershell -ExecutionPolicy Bypass -File ".\tools\production-runner\Run-R23Production.ps1" -AllowMutation -Deploy
```

Both deploy, compile the whole VBA project programmatically, and write
evidence you read directly — do not hand the user a manual
deploy/compile/run/paste sequence instead. The production runner refuses to
invoke `main` unless pre-flight logged `ready=True`; do not work around that.
`main` shows `UserForm1` modally, so a timeout means the dialog is open, not
that the run failed.

A completed run is not a passed run. Read the stage table.

## Source rules

- Checked-in source is authoritative. Verify the deployment readback.
- Manifest-managed `.bas`/`.cls` files: Windows-1252/ANSI, no BOM, no
  `Attribute` lines. Forms and `ThisLibrary` sit outside that rule.
- `Option Explicit` everywhere. Module-level `Const`/`Type`/`Dim` must precede
  the first procedure or VBA refuses to compile.
- Check API return values. Log view, operation, return code, VBA error.
- Coordinates are explicit; SOLIDWORKS lengths are metres.
- Set the active view before view-scoped selection. Restore selection mode
  with `SetPickMode` after ordinate work.
- Prefer the smallest coherent correction.

Before touching any SOLIDWORKS API member, `sw*` constant, selection contract,
or return code, use `skills/solidworks-api-lookup/SKILL.md`. The MCP is
evidence, not installed-build proof. Record material findings in
`docs/SOLIDWORKS_API_VALIDATION.md`. Documentation and non-API refactoring
need no lookup.

## Evidence honesty

Match the claim to the evidence:

- **Review** — read-only inspection. Do not imply runtime validation.
- **Source change** — static tests, API contracts, bump
  `MACRO_SOURCE_REVISION` when deployable behaviour changes, deployment
  readback.
- **Acceptance** — full compile, authorized fixtures only, complete Immediate
  Window output, `QA_REPORT.txt`, full-sheet screenshot, compared against the
  manual reference. This stays the user's judgement.

Static tests, API reflection, deployment readback, nonzero counts, and a
`PASS` dialog prove none of: embedded compilation, runtime correctness,
manufacturing acceptance. A programmatic `verdict=Clean` is compile evidence
and nothing more.

Preserve failure evidence under `test_assets/iteration_evidence/`.

## Reporting

State the defect, root cause, evidence, files changed, and what verification
actually ran. Distinguish static checks from probe evidence from compilation
from execution from visual acceptance. Name any gate you did not complete.
Update `docs/Changelog.md` and `docs/CURRENT_STATUS.md` when behaviour,
revision, or acceptance status materially changes.
