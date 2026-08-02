# Claude Code Handoff: Static Review and Offline Checks

**Prepared:** 2026-07-30

**Repository:** `C:\Users\V.T\Documents\VBA 3D TO 2D`

**Primary audience:** Claude Code working on this repository

**Purpose:** Define the repeatable, evidence-led workflow for reviewing and checking the SOLIDWORKS VBA project without claiming results that require a real VBA compile, macro run, or drawing inspection.

## 1. Operating rule

Treat each verification layer as a separate claim.

A clean source review does not prove that the embedded VBA project compiles. Passing Python tests does not prove that SOLIDWORKS accepts a COM call. A successful guarded SWP deployment does not prove a full VBA project compile. A macro that finishes does not prove that the drawing is correct or useful for manufacturing.

The normal division of work is:

- Claude Code performs repository inspection, source review, API-contract research, offline tests, source-format checks, deployment preflight, evidence review, and preparation of narrowly scoped corrections.
- The user normally performs the full VBA compile, macro execution, and evidence capture in SOLIDWORKS.
- Before any distinct task involving live SOLIDWORKS, the VBA editor, or Computer Use, ask whether the user will perform the task and return the evidence or wants the agent to perform it.
- Do not use live SOLIDWORKS merely because the three test models are authorized. The operator-choice question is still required for each distinct live task.

Read [AGENTS.md](../AGENTS.md) before doing any work. Its authorization boundaries and project rules take precedence over this handoff.

## 2. Current project entry points

At the time of this handoff, use these files as the starting set:

1. [AGENTS.md](../AGENTS.md) — binding project scope, protected assets, API lookup rule, live-collaboration rule, and acceptance criteria.
2. [CLAUDE.md](../CLAUDE.md) — Claude Code repository conventions and quick commands, if present in the working tree.
3. [R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md](R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md) — current R23 objective, accepted evidence, blocked gates, and next work package.
4. [R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md](R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md) — accepted logs, probe defects, and required corrected reruns.
5. [CURRENT_STATUS.md](CURRENT_STATUS.md) — latest implementation and verification state.
6. [R22_REVIEW_RESOLUTION.md](R22_REVIEW_RESOLUTION.md) — resolution of the most recent static review findings.
7. [SOLIDWORKS_API_VALIDATION.md](SOLIDWORKS_API_VALIDATION.md) — accumulated API-contract evidence.
8. [REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md](REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md) and [ORDINATE_GAP_ANALYSIS.md](ORDINATE_GAP_ANALYSIS.md) — drawing-quality requirements.
9. [src/target-spec-hybrid-v2](../src/target-spec-hybrid-v2/) — current deployable exported VBA source.
10. [deployment-manifest.json](../tools/swp-deploy/deployment-manifest.json) — authoritative list of exported components managed in `Fable.swp`.

The source currently identifies itself as an R22 revision. Never assume that this document's revision label remains current: read the actual `SOURCE_REVISION` value and `CURRENT_STATUS.md` at the start of each task.

The protected baseline, manual reference drawings, and three model fixtures are inputs for comparison, not general edit targets:

- Never modify `src/baseline-model-dims/`.
- Never overwrite files under `test_assets/reference_drawings/`.
- Never alter or save the design of the three authorized `.SLDPRT` fixtures.
- Do not run the macro on any other model without new permission.

## 3. Evidence ladder

Use the following ladder to prevent overclaiming:

| Layer | What it establishes | What it does not establish |
|---|---|---|
| Repository state | Exact files, diffs, source revision, and current documentation inspected | Correctness |
| Static source review | Callers and declarations are coherent; required guards and diagnostics are present; obvious compile and logic defects are identified | VBA compilation or COM runtime behavior |
| Offline companion tests | Textual contracts, fake-COM behavior, evidence invariants, deployment-tool structure, and known regressions | Real SOLIDWORKS behavior |
| API evidence | Exact documented signatures, remarks, enums, and locally verified interop facts | That the present call site works in the full macro |
| Deployment preflight | Manifest, source files, revision, target inventory, and bootstrap prerequisites are coherent | Any SWP modification, compilation, or runtime success |
| Guarded deployment and readback | The managed exported components were inserted and read back byte-semantically equivalent to the repository source | A full VBE **Debug > Compile Project**, macro success, or drawing correctness |
| VBE compile and fixture run | The embedded project compiles and executes on an authorized fixture with captured logs | Manufacturing usefulness by itself |
| Visual and semantic acceptance | Datum, dimensions, view ownership, collisions, layout, title block, and QA evidence satisfy the target | Nothing further for the tested matrix; broader models remain unproved |

Always state the highest layer actually completed and list the higher layers still unproved.

## 4. Static review workflow

### 4.1 Establish the exact review scope

Start read-only:

```powershell
git status --short --branch
git diff --stat
git diff --name-status
git diff --check
```

Then inspect the relevant patch. Do not review only the visibly changed procedure: find its callers, public declarations, shared types, constants, stage ledgers, diagnostics, tests, and documentation.

Useful searches include:

```powershell
rg -n "^(Public|Private|Friend)?\s*(Sub|Function|Property)\s+" src/target-spec-hybrid-v2
rg -n "SOURCE_REVISION|Option Explicit|Attribute VB_|SetPickMode|ClearSelection" src/target-spec-hybrid-v2
rg -n "ProcedureName|FunctionName|SharedTypeName" src/target-spec-hybrid-v2
```

Replace the final search terms with the actual procedures, functions, types, fields, and constants under review.

Record:

- the defect or review question;
- the evidence that raised it;
- the affected procedures and components;
- every dependent caller;
- the expected behavioral change;
- the acceptance evidence that will later distinguish a fix from a masked failure.

### 4.2 Confirm source authority

Repository exports are authoritative for offline work only when they are known to match the embedded macro. If the user reports that the VBA editor contains newer code, the embedded code is authoritative. Do not overwrite it from an older export.

Before proposing a material SWP update:

1. compare the current exported source with the latest deployment/readback evidence;
2. identify the embedded revision;
3. report any mismatch;
4. preserve or export the newer embedded source before replacement.

`src/target-spec-hybrid-v2` is the current active tree. References elsewhere to `src/active-ordinate` may describe an older workflow and must not silently redirect the current deployment.

### 4.3 Review the implementation, not just string contracts

For each changed path, review at least:

- public signature compatibility and all callers;
- duplicate declarations or ambiguous public constants;
- `Option Explicit`;
- VBA line-length and continuation limits;
- explicit model, view, sheet, and sketch coordinate systems;
- metres versus drawing or display units;
- selection ownership, target-view activation, selection cleanup, and `SetPickMode`;
- API return values, error codes, and readback where the API can fail silently;
- fail-closed behavior when ownership, datum, circularity, projection, arrangement, or evidence is not proved;
- distinction between physical feature identity and projected drawing geometry;
- supported orthographic views versus deliberately excluded isometric or unsupported views;
- duplicate dimensions and model-item/ordinate overlap;
- diagnostics that include the view, stage, operation, return code, and VBA error;
- QA that reports a failed requirement as failed instead of converting it to a warning or count;
- preservation of title block, notes, barcode, section behavior, and existing stage outputs;
- whether a root failure is being repaired or only a downstream cascade is being hidden.

Do not accept a change solely because it makes a source-pattern test green. Confirm that the tested pattern still represents the intended SOLIDWORKS behavior.

### 4.4 Validate every load-bearing SOLIDWORKS API assumption

Before writing or approving a change involving a SOLIDWORKS method, property, enum, selection rule, view, drawing dimension, annotation, or any `sw*` constant:

1. read `skills/solidworks-api-lookup/SKILL.md`;
2. locate the member with the local `solidworks-api` MCP;
3. retrieve the exact member contract and read its Remarks;
4. confirm enum values and return semantics;
5. separate compatibility-snapshot evidence from installed SOLIDWORKS 2025 evidence;
6. use the installed type library/Object Browser or a narrow live probe for load-bearing uncertainty;
7. record material findings in `docs/SOLIDWORKS_API_VALIDATION.md`.

Search results are locators, not sufficient evidence. Do not guess enum values or infer a COM signature from an older wrapper.

If installed-build verification requires a live probe, stop at the documented uncertainty and ask the required operator-choice question before using SOLIDWORKS.

### 4.5 Review the whole evidence contract

For a drawing-generation change, trace:

```text
feature discovery
  -> candidate qualification
  -> drawing-view ownership
  -> datum proof
  -> selection
  -> dimension or annotation API call
  -> return/readback proof
  -> placement or arrangement
  -> stage evidence
  -> final QA
```

A stage must not report success merely because a later object count is nonzero. The evidence should remain attributable to the intended operation and view.

## 5. Offline checks

Use the project-local bundled Python when ambient `python` is missing or has different packages:

```powershell
$ProjectPython = 'C:\Users\V.T\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
```

Do not reuse `$HOME`, `$home`, or `$CODEX_HOME` for task variables.

### 5.1 Run focused tests first

The companion test package is easiest to invoke from its own directory:

```powershell
Push-Location .\tools\solidworks-automation-companion
& $ProjectPython -m unittest tests.test_target_spec_hybrid_v2_source -v
& $ProjectPython -m unittest tests.test_target_spec_hybrid_v2_evidence_contracts -v
& $ProjectPython -m unittest tests.test_swp_deployment_tooling -v
Pop-Location
```

Select the smallest applicable set:

- `test_target_spec_hybrid_v2_source.py` — exported-source inventory and fail-closed source contracts;
- `test_target_spec_hybrid_v2_evidence_contracts.py` — evidence ledger, projection, layout, title, extent, and final-QA invariants;
- `test_target_spec_hybrid_v2_r4_contracts.py` — older and current textual/behavioral contracts, including several stale R20 expectations;
- `test_vba_runtime_regressions.py` — source guards for previously observed VBA/runtime regressions;
- `test_swp_deployment_tooling.py` — manifest, verifier, bootstrap, compile-probe, readback, and atomic-promotion contracts;
- `test_sw2025_drawing_companion.py` — pure logic and fake-COM drawing companion behavior;
- `test_sw_part_sketch_selection.py` — generic companion selection regression tests.

### 5.2 Run the complete offline suite

From the repository root:

```powershell
& $ProjectPython -m unittest discover `
  -s .\tools\solidworks-automation-companion\tests `
  -p 'test*.py' `
  -v
```

Current observed baseline on 2026-07-30:

- 74 tests executed;
- 69 passed;
- 5 failures are present in `test_target_spec_hybrid_v2_r4_contracts.py`;
- the failures pin superseded R20 source strings or behavior rather than the current R22 design.

The affected test methods are:

- `test_r20_auto_arrange_is_required_fail_closed_and_has_proved_fallback` — produces two reported failures;
- `test_r20_closed_circle_contract_uses_curve_params_and_cylinder_fallback`;
- `test_r20_model_to_view_coordinates_are_not_double_translated`;
- `test_r20_source_identity_and_fixed_hybrid_are_explicit`.

This is a **known failing baseline**, not a green suite. Re-run it every time and compare the exact failure list. A new failure, a missing test, or a changed failure message requires investigation. Do not suppress, skip, or relabel the five failures merely to produce a zero exit code.

The companion directory is ignored by the main repository and is maintained as a separate local checkout/fork. Do not edit its tests or implementation as an incidental part of a VBA-source fix. Obtain explicit scope for companion changes and commit them in their own repository.

### 5.3 Check deployable VBA source hygiene

Every managed `.bas` and ordinary `.cls` file must:

- be Windows-1252/ANSI without a UTF-8 BOM;
- use CRLF line endings;
- contain `Option Explicit`;
- contain no VBA export metadata such as `Attribute VB_*`;
- preserve the native source component boundary and filename in the manifest.

This read-only PowerShell check covers those conditions:

```powershell
$ManifestPath = Resolve-Path .\tools\swp-deploy\deployment-manifest.json
$Manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
$SourceRoot = Resolve-Path (Join-Path (Split-Path $ManifestPath) $Manifest.sourceDirectory)
$Cp1252 = [Text.Encoding]::GetEncoding(
    1252,
    [Text.EncoderExceptionFallback]::new(),
    [Text.DecoderExceptionFallback]::new()
)
$SourceProblems = [System.Collections.Generic.List[string]]::new()

foreach ($Component in $Manifest.components) {
    $Path = Join-Path $SourceRoot $Component.file
    $Bytes = [IO.File]::ReadAllBytes($Path)
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        $SourceProblems.Add("$Path : UTF-8 BOM")
    }
    try {
        $Text = $Cp1252.GetString($Bytes)
        $RoundTrip = $Cp1252.GetBytes($Text)
        if ([Convert]::ToBase64String($Bytes) -ne [Convert]::ToBase64String($RoundTrip)) {
            $SourceProblems.Add("$Path : not lossless Windows-1252")
        }
    }
    catch {
        $SourceProblems.Add("$Path : invalid Windows-1252")
        continue
    }
    if ($Text -notmatch '(?m)^\s*Option Explicit\s*$') {
        $SourceProblems.Add("$Path : missing Option Explicit")
    }
    if ($Text -match '(?m)^\s*Attribute\s+') {
        $SourceProblems.Add("$Path : export metadata present")
    }
    if ($Text -match '(?<!\r)\n') {
        $SourceProblems.Add("$Path : non-CRLF line ending")
    }
}

if ($SourceProblems.Count -gt 0) {
    $SourceProblems
    throw 'Deployable VBA source hygiene check failed.'
}
'Deployable VBA source hygiene check passed.'
```

Also run:

```powershell
git diff --check
```

The deployer independently repeats key manifest, `Option Explicit`, metadata, source-identity, and component checks. Keep the standalone review because it catches formatting problems before any deployment step.

## 6. `tools/swp-deploy`

### 6.1 Role

[tools/swp-deploy](../tools/swp-deploy/) is the guarded bridge from exported source to `Fable.swp`. It exists to avoid manual copy/paste of standard modules and ordinary class modules while preserving evidence and preventing partial promotion.

The deployment manifest currently manages 15 components:

- 9 standard modules;
- 6 ordinary class modules.

It deliberately does not replace:

- `ThisLibrary`;
- `UserForm1`;
- `UserFormSection`;
- the one-time bootstrap module.

Those special components remain outside automated component promotion.

### 6.2 Read-only preflight

Run this after source checks:

```powershell
powershell -ExecutionPolicy Bypass `
  -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1 `
  -PreflightOnly
```

Preflight validates the manifest, source files, revision identity, target inventory, expected managed count, and bootstrap prerequisites. It exits before modifying `Fable.swp`.

A successful preflight means only that the deployment inputs appear coherent. Report it as:

> Deployment preflight passed; no SWP was changed.

Do not report it as a compile or deployment success.

### 6.3 Full guarded deployment

Full deployment is a material binary change and uses a running SOLIDWORKS/VBE automation path. It is not part of a purely offline review. Before running it:

1. ask the user the live-task operator-choice question;
2. confirm that this deployment is in the agreed scope;
3. make sure `Fable.swp` is closed in the VBA editor;
4. preserve current embedded source if it may be newer;
5. confirm SOLIDWORKS is running and the bootstrap exists.

The command is:

```powershell
powershell -ExecutionPolicy Bypass `
  -File .\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1
```

The deployer:

1. validates the manifest and repository source;
2. inventories the current target;
3. checks exclusive access;
4. saves timestamped original and candidate evidence;
5. creates each managed component by type and inserts metadata-free source;
6. saves a named candidate through the VBA environment;
7. runs the bootstrap probe through `ISldWorks.RunMacro2`;
8. extracts the candidate and verifies all managed source against the repository;
9. atomically replaces the target only after the gates pass;
10. re-verifies the promoted SWP and restores the original if post-promotion verification fails.

Evidence is written under:

```text
test_assets/iteration_evidence/swp_deployment/<timestamp>/
```

Review at least:

- `deployment-result.txt`;
- `compile-result.txt`;
- `compile-probe-scope.txt`;
- `verification.json`;
- `post-promotion-verification.json`;
- original, pre-promotion, and candidate SWP copies.

`COMPILE_PROBE|status=SUCCESS` proves only that the bootstrap executed. The deployer explicitly records `BOOTSTRAP_EXECUTION_ONLY`; it does **not** prove a full VBE project compile.

The required next gate after a successful deployment, **for any run that
mutates a drawing and for production acceptance**, is still:

1. open the promoted macro in the SOLIDWORKS VBA editor;
2. run **Debug > Compile Project**;
3. capture any error and highlighted line;
4. only after compile success, run the agreed authorized fixture;
5. capture complete Immediate Window output, `QA_REPORT`, settings, and drawing screenshots;
6. compare the drawing with the manual reference and target specification.

**Read-only probes are exempt from steps 1 to 3** (user-authorized
2026-08-02). A strictly read-only `R23_Probe*` entry point may be deployed
and run directly: a probe that does not compile fails at its first
statement, so the manual gate proved nothing there that the run itself does
not prove. Steps 4 to 6 are unchanged, and the exemption does not extend to
any procedure gated behind `allowMutation`.

Programmatic full-project compilation is also authorized, through the VBE
`CommandBars.FindControl(...).Execute` route that
`tools/swp-deploy/Module0_SourceDeployment.bas` already uses for Save. See
[R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md](R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md).

### 6.4 Independent SWP inventory and readback

The verifier may be used directly for read-only inventory:

```powershell
& $ProjectPython .\tools\swp-deploy\verify_swp_vba.py `
  --manifest .\tools\swp-deploy\deployment-manifest.json `
  --inventory-only
```

Without `--inventory-only`, it compares managed embedded source and revision with the repository source. Use `--json-output` when evidence must be retained. Do not overwrite a prior evidence file.

## 7. `tools/solidworks-automation-companion`

### 7.1 Project role

The companion at:

```text
C:\Users\V.T\Documents\VBA 3D TO 2D\tools\solidworks-automation-companion
```

is a verification and investigation layer. VBA remains the production drawing generator.

For normal offline work, use the companion for:

- fake-COM and pure-logic tests;
- exported-VBA source contracts;
- evidence-ledger and fail-closed QA contracts;
- deployment-tool tests;
- regression tests for previously observed source/runtime defects;
- narrowly scoped data analysis that does not control SOLIDWORKS.

Do not turn it into a second production drawing engine or treat a Python-generated result as proof that the VBA macro compiles or runs.

### 7.2 Live companion scripts are a separate gate

The companion also contains scripts such as environment preflight, drawing review, annotation, feature-audit, ordinate, layout, title, and selection probes. Some of these instantiate or control SOLIDWORKS.

In a source-only review, do not casually run:

- `scripts/sw_preflight.py`;
- `scripts/sw_review.py`;
- COM-based drawing or model probes;
- scripts that register/start integrations or modify the environment.

If a narrow live probe is genuinely required:

1. state the exact unresolved API or geometry question;
2. state the exact fixture and read-only operation;
3. ask whether the user will run it or wants the agent to use Computer Use/live automation;
4. keep the probe within the three authorized parts;
5. do not save a model;
6. retain the probe command, environment, SOLIDWORKS build, raw output, and interpretation;
7. distinguish probe evidence from the full macro result.

`sw_review.py` can help produce preview images and JSON/Markdown observations after a drawing exists, but those observations supplement rather than replace manual reference comparison and manufacturing assessment.

### 7.3 Separate repository boundary

The companion is gitignored by the main repository and may have its own Git history and upstream. Before changing it:

```powershell
git -C .\tools\solidworks-automation-companion status --short --branch
git -C .\tools\solidworks-automation-companion remote -v
```

Do not include companion changes in a main-repository commit by accident. Report and version companion changes separately.

## 8. Recommended change-and-review cycle

For each repair or enhancement:

1. Read `AGENTS.md`, current status, recent resolution/handoff, and the relevant source.
2. Establish the exact defect, evidence, root cause, affected procedures, and callers.
3. Compare with the protected baseline only where it provides a known working contract.
4. Validate uncertain SOLIDWORKS API behavior before changing code.
5. Make the smallest coherent source change; do not weaken strict gates to remove an error.
6. Keep every changed component complete, metadata-free, and synchronized with its callers.
7. Update source identity when deployable behavior changes.
8. Update `docs/CHANGELOG.md`, `docs/CURRENT_STATUS.md`, and material API evidence.
9. Run focused companion tests.
10. Run the complete companion suite and compare with the known failing baseline.
11. Run source-hygiene checks and `git diff --check`.
12. Run SWP deployment preflight.
13. Perform a final static diff review across code, tests, and documentation.
14. If configured, ask Codex through the installed Claude Code plugin for an independent working-tree review; treat findings as review input, inspect every finding against source and project rules, and do not auto-apply them.
15. Stop and report the offline result unless the user has separately selected the agent for the live deployment task.
16. After an authorized guarded deployment, require full VBE compile, fixture runtime evidence, and visual/manufacturing review before acceptance.

## 9. Review report format

Return a report with these headings:

### Scope

- files and revision reviewed;
- defect/question;
- protected assets left unchanged.

### Static findings

- root cause;
- affected procedures and callers;
- actionable findings with file and line;
- non-findings or disproved concerns when relevant.

### API evidence

- exact member/enum;
- documented signature and Remarks;
- installed-build verification status;
- remaining uncertainty.

### Offline checks

- focused test commands and results;
- full suite result with pass/fail counts;
- exact known versus new failures;
- source-hygiene result;
- `git diff --check` result.

### SWP state

- preflight result;
- whether deployment was performed;
- managed source/readback result if performed;
- compile-probe scope.

### Still unproved

- full VBE compile;
- fixture execution;
- Immediate Window and QA evidence;
- visual/manufacturing acceptance;
- any part of the validation matrix not run.

### Exact next action

Give the user one concrete next gate, including the fixture, settings, command or UI action, and evidence to return.

## 10. Claim language

Use precise statements such as:

> Static review found no remaining actionable defect in the inspected paths. Focused tests passed. The full offline suite remains at the known baseline of 69 passes and 5 stale R20 contract failures. Deployment preflight passed and did not modify the SWP. Full VBE compile, fixture runtime, and visual acceptance remain unproved.

Avoid statements such as:

- “the macro compiles” after Python checks;
- “the macro works” after source readback;
- “QA passed” because counts are nonzero;
- “the drawing is correct” without inspecting the generated views and reference comparison;
- “all tests passed” while known failures remain;
- “SOLIDWORKS 2025 supports this call” based only on a search result or compatibility snapshot.

The final standard is a truthful chain of evidence, not a successful command.
