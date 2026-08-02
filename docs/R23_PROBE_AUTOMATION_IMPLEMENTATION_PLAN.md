# R23 probe-automation tool — self-contained implementation plan

**Status: planned, not built.** Authorized by the user on 2026-08-02.

## 0. How to use this document

This is written so a fresh session needs nothing but the repository and this
file. Read it top to bottom before touching anything.

Before the first edit, also read the binding contracts — they are short and
they govern everything here:

- [Agents.md](../Agents.md) — fixture authorization, protected assets,
  evidence ladder, VBA engineering rules
- [../CLAUDE.md](../CLAUDE.md) — per-task reading list and update duties
- [skills/solidworks-api-lookup/SKILL.md](../skills/solidworks-api-lookup/SKILL.md)
  — mandatory before touching any `sw*` constant or API member

Do **not** read `docs/R23_IMPLEMENTATION_PLAN.md` for this work. That is the
drawing-engine plan (Phases 0–12) and is unrelated to this tool except that
this tool runs its probes.

## 1. What this repository is, in one paragraph

A SOLIDWORKS 2025 SP1.2 VBA macro that generates 2D manufacturing drawings
from 3D parts. The deployable macro is `Fable.swp` at the repository root.
Its managed source lives as plain text in `src/target-spec-hybrid-v2/` and
is synchronised into the `.swp` by a guarded PowerShell deployer,
`tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`, driven by
`tools/swp-deploy/deployment-manifest.json`. Offline verification is a
Python `unittest` suite under
`tools/solidworks-automation-companion/tests/` (currently 381 tests, with
five known-stale R20 failures that are expected and must not be "fixed").

## 2. The problem this tool removes

An R23 iteration currently costs the user five manual steps:

1. run the deployment PowerShell command;
2. open the VBA editor and run **Debug > Compile Project**;
3. open the authorized fixture part and the drawing, in the right order;
4. select and run each probe macro by hand — there are now nine; and
5. copy the entire Immediate Window and paste it into chat.

Step 5 is the expensive one. `Debug.Print` writes only to the Immediate
Window, and nothing outside the VBA host can read that buffer. That single
fact is why the paste step exists, and removing it is the core of this work.

**Intended outcome:** one PowerShell command deploys, compiles, opens the
fixture correctly, runs all nine probes, and writes a log file the agent
reads directly. The user pastes nothing.

## 3. Authorization already granted

Committed in `d4787cd`. Do not re-litigate these; they are settled.

- The agent **may perform full-project VBA compilation programmatically.**
- **Read-only `R23_Probe*` entry points no longer require a preceding manual
  Debug > Compile Project.** A probe that does not compile fails at its
  first statement, so the gate bought nothing there.
- Production acceptance and **any mutating run still require** the manual
  full-project compile. The exemption does not extend to anything gated
  behind an `allowMutation` argument.

Amended files, already committed: `Agents.md` ("Read-only probe exception"),
`docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md`,
`docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md`.

## 4. Current repository state

`src/target-spec-hybrid-v2/` holds 36 managed components: 19 standard
modules (`Module1_Main` … `Module19_SemanticQA`) and 17 class modules. The
manifest is authoritative; `kind` is `StdModule` or `ClassModule` and no
other value is accepted.

Nine read-only probe entry points exist today:

| # | Module | Procedure | Terminal evidence line |
|---|---|---|---|
| 1 | `Module12_FeatureQualification` | `R23_ProbeFeatureCatalog` | `R23_CATALOG_END` |
| 2 | `Module13_ProjectionResolution` | `R23_ProbeViewProjections` | `R23_PROJECTION_END` |
| 3 | `Module14_AnnotationImport` | `R23_ProbeAnnotationReconciliation` | `R23_ANNOTATION_END` |
| 4 | `Module15_OrdinateScheme` | `R23_ProbeOrdinateScheme` | `R23_ORDINATE_END` |
| 5 | `Module16_CalloutDefinition` | `R23_ProbeCalloutDefinition` | `R23_CALLOUT_END` |
| 6 | `Module17_SectionPath` | `R23_ProbeSectionPath` | `R23_SECTION_END` |
| 7 | `Module10_SectionDimensionEngine` | `R23_ProbeSectionDimensions` | `R23_SECTIONDIM_END` |
| 8 | `Module18_ContentEnvelope` | `R23_ProbeContentEnvelope` | `R23_ENVELOPE_END` |
| 9 | `Module19_SemanticQA` | `R23_ProbeSemanticQA` | `R23_SEMANTICQA_END` |

That is also the dependency order the runner must use.

Every probe already reports its own read-only boundary on the terminal
line — `creations=0`, `mutations=0`, `initialSelectionCount`,
`finalSelectionCount`, `drawingUnchanged=` (Module12 says `modelUnchanged=`).
Preserve that; it is the safety evidence.

`Debug.Print` call counts to convert (PA-102): Module10 16, Module12 14,
Module13 16, Module14 15, Module15 19, Module16 18, Module17 18,
Module18 20, Module19 11. Roughly 147 sites, all mechanical.

## 5. Findings that constrain the design

Each of these cost a live run or a search to establish. Do not rediscover
them the hard way.

**5.1 The `solidworks-api` MCP cannot confirm the VBE compile control ID.**
Searches for `CommandBars VBE compile project control` and
`VBA project compile macro editor` return only SOLIDWORKS geometry and
macro-path records — zero VBIDE hits. That corpus is the SOLIDWORKS API
Help; VBE `CommandBars` control IDs belong to the Microsoft Office/VBA
extensibility model, which is not in it. **Any control ID you "remember" is
recall, not lookup.** This repo has already shipped a bug from exactly that
mistake (`swInsertDimensionsMarkedForDrawing` — see
`docs/SOLIDWORKS_API_VALIDATION.md`).

**5.2 The control can be resolved by caption instead of by ID.**
`CommandBars.FindControl` takes an ID, but nothing stops walking
`VBE.CommandBars` → `.Controls` and matching `.Caption`. Doing that removes
the magic number entirely and logs the ID as a by-product. Residual risk: a
localized SOLIDWORKS install where the caption is not English — handle it by
falling back to a full enumeration dump rather than guessing.

**5.3 The VBE route is already proved in this repository.**
`tools/swp-deploy/Module0_SourceDeployment.bas:209-228` reaches the VBE
object model from inside the SOLIDWORKS VBA host and executes a built-in
command through `CommandBars.FindControl(1, 3, "", False)` — with `Nothing`
and `.Enabled` guards — because `VBProject.SaveAs` raises error 748 on a
host-managed `.swp`. Reuse that exact shape. The
`Microsoft Visual Basic for Applications Extensibility 5.3` reference is
already present in the project.

**5.4 `RunMacro2` is synchronous.**
`tools/swp-deploy/SolidWorksMacroInvoker.cs` blocks until the macro returns,
so a modal compile dialog blocks the PowerShell call. The existing
`-TimeoutSeconds` parameter on the deploy script does **not** interrupt it —
a real timeout needs the call in a background runspace.

**5.5 A compile error cannot be read programmatically.** VBE reports it in a
modal box and highlights the offending line. The error text and line come
from the user, once per failure. Design for that, do not pretend otherwise.

**5.6 Reference binding is a live failure mode.** On 2026-08-02 a probe
aborted with
`R23_SECTIONDIM_FATAL|reason=UnauthorizedFixture|path=V:\VEEMAP\SW_data\P-0251-14A-001.SLDPRT`.
The drawing lives on `V:` and had bound its part reference to the network
sibling instead of the authorized `test_assets\models\` copy. **Opening the
local part BEFORE the drawing rebinds it.** The fixture guard
(`Module1_Main.IsAuthorizedFixture`, constants at
`src/target-spec-hybrid-v2/Module1_Main.bas:57-62`) fired correctly and must
not be widened.

## 6. Design

### A. `Module21_EvidenceSink.bas` (new, deployed)

The piece that removes the paste step.

| Procedure | Purpose |
|---|---|
| `OpenLog(rootFolder) As String` | Creates `test_assets/iteration_evidence/probe_runs/<yyyyMMdd_HHmmss>/probe_log.txt`; returns the path |
| `LogLine(text)` | Writes to the file **and** `Debug.Print`s |
| `CloseLog` | Resets state |
| `IsOpen() As Boolean` | Lets `LogLine` degrade to `Debug.Print` alone when no run is active |

Open `For Append`, `Print #`, `Close #` on **every** `LogLine` call. Holding
a handle risks losing buffered lines when a modal dialog or a crash
intervenes, and the log volume is a few hundred lines — the cost is
irrelevant next to losing the evidence.

Reuse the folder-resolution pattern already in
`Module6_QAEngine.CreateUniqueReportFolder`
(`src/target-spec-hybrid-v2/Module6_QAEngine.bas:1565-1610`): it derives the
`test_assets` root from the part path by locating the `\test_assets\models\`
marker, creates each level with `MkDir` guarded by a `DirectoryExists`
helper, and disambiguates with a numeric suffix. Do not invent a second path
resolver.

Then retrofit every `Debug.Print` in the nine probe modules and in
`CRunEvidence.AddInfo` / `AddWarning` / `AddFailure`
(`src/target-spec-hybrid-v2/CRunEvidence.cls:356-373`) to route through
`LogLine`.

**Preserve the `QA INFO: ` / `QA WARNING: ` / `QA FAILURE: ` prefixes
byte-for-byte.** Every quoted log in `docs/`, every contract test, and every
piece of accumulated evidence depends on them.

Writing a file does not break the read-only claim. `Agents.md` permits
"retain disposable drawings and evidence under `test_assets/`", and the
probes still touch no model and no drawing, so `drawingUnchanged=True`
continues to mean what it has always meant. The sink must never write inside
`src/`, `test_assets/reference_drawings/`, or `src/baseline-model-dims/`.

### B. `Module20_ProbeRunner.bas` (new, deployed)

| Procedure | Purpose |
|---|---|
| `R23_EnumerateVbeControls` | Diagnostic: logs every VBE control `.Id` and `.Caption` |
| `R23_CompileProject` | Resolves the compile control, executes it, reports the verdict |
| `R23_TouchAllModules` | Calls one no-op per deployed standard module; names the first that fails |
| `R23_RunAllProbes` | Open log → compile → touch → nine probes → close log |

**Compile resolution and verdict.** Walk `VBE.CommandBars` → `.Controls`,
match a caption containing "compile" case-insensitively, and log the
matched `.Id` and `.Caption` so the ID is recorded as evidence rather than
assumed. Then: record `.Enabled` → `.Execute` → `DoEvents` → re-resolve →
record `.Enabled` again. VBE disables the Compile control once a project is
fully compiled, so **still enabled afterwards means not clean.** If no
control matches, dump the full enumeration and fail with a named reason —
never fall back to a guessed ID.

**Compile-failure localisation.** Add a trivial
`Public Sub R23_CompileTouch()` to every deployed standard module.
`R23_TouchAllModules` invokes each in turn and reports the first that fails
to load. VBA compiles at module granularity, so a module that loads has
compiled; classes are covered transitively because every class is referenced
by some module. This yields the failing **module name** without anyone
reading a dialog.

**Compile-failure loop.** Run → if unclean, the user supplies the error text
and highlighted line → fix the source → redeploy → rerun. Repeat to a clean
compile, then continue. `R23_RunAllProbes` must **not** run the probes when
the compile is unclean; it reports and stops.

**Probe isolation.** Wrap each probe so one unhandled error cannot abort the
rest, and record `probe=<name>|status=Completed` or `status=Error:<n>` per
entry.

### C. `tools/probe-runner/Run-R23Probes.ps1` (new)

1. optionally invoke `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`
   (`-Deploy` switch; default assumes the current SWP);
2. attach to the running SOLIDWORKS instance — reuse the interop-loading and
   invoker-compilation block at
   `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1:85-118` rather than writing
   a second one;
3. **open the authorized fixture part first, then the drawing**, via
   `ISldWorks.OpenDoc6`. See finding 5.6 — this is the whole reason the
   script opens documents at all;
4. invoke `Module20_ProbeRunner.R23_RunAllProbes` through the existing
   `SolidWorksMacroInvoker`, **inside a background runspace**, so a modal
   dialog yields `compile=DialogSuspected` instead of a silent hang;
5. print the log path and the last N lines.

Fixture and drawing paths are parameters with authorized defaults, validated
against the same three-fixture list `Module1_Main` enforces. The script must
refuse an unlisted path rather than pass it through. The VBA guard is the
real gate, but a script that cheerfully opens an unauthorized model and lets
the macro reject it is worse than one that will not try.

## 7. Tasks

Each task is done when its acceptance line is true.

- [ ] **PA-100** Add `Module21_EvidenceSink.bas` per Design A.
  *Accepts:* `LogLine` writes to file and Immediate Window; degrades to
  `Debug.Print` when no log is open.
- [ ] **PA-101** Route `CRunEvidence.AddInfo`/`AddWarning`/`AddFailure`
  through `LogLine`.
  *Accepts:* prefixes byte-identical to today.
- [ ] **PA-102** Replace every `Debug.Print` in the nine probe modules with
  `LogLine` (~147 sites).
  *Accepts:* no raw `Debug.Print` inside any `R23_Probe*` procedure.
- [ ] **PA-103** Add `R23_EnumerateVbeControls`; run it once live; store the
  ID/caption listing under `test_assets/iteration_evidence/`.
  *Accepts:* a stored listing showing a control captioned for compilation,
  with its ID.
- [ ] **PA-104** Add `R23_CompileProject` per Design B, resolving by caption
  with the Module0 `Nothing`/`.Enabled` guards.
  *Accepts:* reports resolved ID, caption, enabled-before, enabled-after and
  a verdict; fails named when no control matches.
- [ ] **PA-105** Add `Public Sub R23_CompileTouch()` to all 19 deployed
  standard modules, plus `R23_TouchAllModules`.
  *Accepts:* a contract test proves every `StdModule` in the manifest has
  one.
- [ ] **PA-106** Add `R23_RunAllProbes` per Design B.
  *Accepts:* stops before the probes when the compile is unclean; per-probe
  status lines; log path printed last.
- [ ] **PA-107** Add `tools/probe-runner/Run-R23Probes.ps1` per Design C.
  *Accepts:* refuses an unauthorized fixture path; opens part before
  drawing; background runspace around the invoker.
- [ ] **PA-108** Register both new modules in
  `tools/swp-deploy/deployment-manifest.json` (36 → 38, `kind` =
  `StdModule`) and update the inventory-lock count in
  `tools/solidworks-automation-companion/tests/test_swp_deployment_tooling.py`.
  *Accepts:* preflight prints `Managed components: 38`.
- [ ] **PA-109** Add `tools/solidworks-automation-companion/tests/test_r23_probe_runner_contracts.py`.
  *Accepts:* covers sink hygiene; no raw `Debug.Print` in probes; runner
  calls all nine probes in the documented order; runner contains no
  `allowMutation` and no mutating procedure name; every standard module has
  `R23_CompileTouch`; source hygiene for both new modules.
- [ ] **PA-110** Update this document to record what the live run proved —
  in particular the resolved control ID and caption.
- [ ] **PA-111** Update `docs/CURRENT_STATUS.md`, `docs/Changelog.md`, and
  `docs/SOLIDWORKS_API_VALIDATION.md`. The API doc must record that the VBE
  control ID is **not** in the SOLIDWORKS corpus and how it was proved
  instead.
- [ ] **PA-112** First live run; iterate the compile-failure loop to a clean
  compile; confirm the log file content matches the Immediate Window.

## 8. Source hygiene — non-negotiable

Every deployable `.bas` and ordinary `.cls` in `src/target-spec-hybrid-v2/`:

- Windows-1252/ANSI, **every byte below 0x80** — no smart quotes, no em
  dashes, no `Ø`. Use `ChrW$(216)` if a diameter sign is ever needed at
  runtime;
- CRLF line endings, no UTF-8 BOM;
- no `Attribute` metadata lines;
- starts with `Option Explicit`;
- lines ≤ 79 columns;
- balanced `Function`/`End Function` and `Sub`/`End Sub`.

A UTF-8 em dash inserted by an editing script once broke SWP readback in
this project. Normalise and re-check after every scripted edit.

**Do not bump `MACRO_SOURCE_REVISION`** in `Module1_Main.bas`. It is
currently `target-spec-hybrid-v2-2026-07-29-r22`. These modules are not on
the reachable production path, and several existing contract tests assert
the revision is unchanged until behaviour is wired in.

## 9. Verification

Offline, after every source change:

```bash
cd "tools/solidworks-automation-companion" && python -m unittest discover -s tests -q
```

Expect the new contracts to pass and **exactly** the five known-stale R20
failures in `test_target_spec_hybrid_v2_r4_contracts.py`. Any sixth failure
is yours.

```bash
powershell -ExecutionPolicy Bypass -File ".\tools\swp-deploy\Deploy-TargetSpecHybrid.ps1" -PreflightOnly
```

Expect `Managed components: 38` once PA-108 lands.

Live acceptance — the tool is accepted when one invocation:

```bash
powershell -ExecutionPolicy Bypass -File ".\tools\probe-runner\Run-R23Probes.ps1" -Deploy
```

produces a log file whose content matches the Immediate Window for the same
run; the compile stage reports a verdict derived from the control's enabled
state; every probe reports `mutations=0` and `drawingUnchanged=True`; the
`part=` field on each `_BEGIN` line reads the `test_assets\models\` path;
and the agent reads that file without anything being pasted.

## 10. Prohibitions

- **Never widen `IsAuthorizedFixture`.** Three fixtures, exact paths.
- **No mutating procedure in the runner or the script.** Nothing gated
  behind `allowMutation` may appear. Running one unattended requires
  separate, per-run permission from the user.
- **Never modify** `src/baseline-model-dims/`, `test_assets/reference_drawings/`,
  or any authorized model. Never save a fixture.
- **Never call `Module6_QAEngine.EmitRunEvidence` from a probe or the
  runner.** It is the production gate and demands fourteen pipeline stages a
  read-only run never reaches. A probe that called it once produced a
  spurious `RESULT: FAIL` and a "Failed Closed" dialog.
- **Do not guess a control ID, an enum value, or an API contract.** Look it
  up per the skill, or prove it live and record the evidence path.
- **Distinguish evidence levels in every report:** static verification, API
  evidence, deployment readback, VBE compilation, live execution, screenshot
  review, manufacturing acceptance. Preflight success is not runtime proof.

## 11. Still manual after this

SOLIDWORKS must be running. A compile error still needs a human to read the
dialog and the highlighted line. Screenshots and visual acceptance stay
manual, as they must. Production acceptance keeps its manual compile gate.

The honest floor: deploy and compile happen inside one command, probes run
unattended, and the agent reads the log directly.
