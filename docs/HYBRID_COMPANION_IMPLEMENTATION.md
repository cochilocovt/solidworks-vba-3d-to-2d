# Hybrid Companion Implementation

## Boundary

`tools/solidworks-automation-companion` is a project-local fork used for
SOLIDWORKS 2025 probes, regression fixtures, exports, diagnostics, and QA. It is
not a second production drawing engine and does not share mutable runtime state
with the VBA macro.

The fork is pinned in `UPSTREAM.md`. No global Codex skill or MCP registration
is part of this phase.

## Modules

| Module | Contract |
|---|---|
| `sw2025_constants.py` | SW2025 enums verified from interop `33.1.2.4` |
| `sw_annotations_2025.py` | `InsertModelAnnotations4`, retry, evidence, cleanup |
| `sw_feature_audit.py` | Conservative feature-tree evidence and derived-feature proof |
| `sw_ordinate.py` | Candidate qualification, duplicate suppression, datum-first API call |
| `sw_view_policy.py` | Strict orthographic eligibility; exclude section/detail/axonometric |
| `sw_layout.py` | Sheet bounds, title-block reserve, lanes, rectangle collisions |
| `sw_title_block.py` | Required properties, linked notes, barcode validation |
| `sw_drawing_review.py` | JSON/Markdown `PASS`, `PASS WITH WARNINGS`, or `FAIL` |
| `probe_sw2025_drawing_api.py` | Disposable read-only part / unsaved drawing live probe |

## Run the pure and fake-COM tests

```powershell
cd "C:\Users\V.T\Documents\VBA 3D TO 2D\tools\solidworks-automation-companion"

```

## Run the disposable live probe

Close or restart SOLIDWORKS first if it has a hidden non-ROT process, then run:

```powershell
.\.venv\Scripts\python.exe scripts\probe_sw2025_drawing_api.py `
  --part "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0252-01-013.SLDPRT" `
  --output "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\companion_evidence\P-0252-01-013_api_probe.json"
```

The probe saves no drawing. It records the installed build, configured template,
feature audit, view roles, and annotation-import attempts.

## Fail-closed rule

The reviewer must fail when required intent, annotation import, ordinate creation,
datum proof, duplicate suppression, view eligibility, layout, title-block data,
or collision evidence is missing or unsuccessful. It must not invent datums,
fits, GD&T, tolerances, or hole intent.

