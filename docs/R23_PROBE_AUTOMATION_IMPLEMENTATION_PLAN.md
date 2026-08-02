# R23 probe-automation tool — implementation plan

Status: **planned, not built**. Authorized by the user on 2026-08-02.

Read this with [Agents.md](../Agents.md) and
[docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md](CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md).
This document is self-contained enough to start a fresh session from.

## 1. The problem this removes

Every R23 iteration currently costs the user five manual steps:

1. run the deployment PowerShell command;
2. open the VBA editor and run **Debug > Compile Project**;
3. open the authorized fixture part and the drawing, in the right order;
4. select and run each probe macro by hand, one at a time; and
5. copy the entire Immediate Window and paste it into chat.

Step 5 is the expensive one — the Phase 9 log alone was ~30 lines of dense
delimited text — and step 4 is about to get worse, because there are now
nine probes.

`Debug.Print` writes only to the Immediate Window, and nothing outside the
VBA host can read that buffer. That single fact is why the paste step
exists, and removing it is the core of this work.

## 2. What the user authorized

- The agent may perform **full-project VBA compilation** programmatically.
- The **manual compile gate no longer applies to read-only R23 probes.**
  It still applies unchanged to production acceptance — see
  [Agents.md](../Agents.md) "Validation and Acceptance" item 3.
- Documentation across the repo may be amended to reflect both.

Not authorized, and deliberately out of scope: invoking any **mutating**
procedure from the runner. `allowMutation` paths stay out of it, and
running one unattended needs separate, per-run permission.

## 3. Why programmatic compilation is possible here

`tools/swp-deploy/Module0_SourceDeployment.bas:209-228` already reaches the
VBE object model from inside the SOLIDWORKS VBA host:

```vba
Set saveControl = targetProject.VBE.CommandBars.FindControl(1, 3, "", False)
If saveControl Is Nothing Then Err.Raise ...
If Not CBool(saveControl.Enabled) Then Err.Raise ...
saveControl.Execute
DoEvents
```

That exists because `VBProject.SaveAs` raises error 748 against a
SOLIDWORKS host-managed `.swp`. The important part for this plan is that it
**works**: the VBIDE 5.3 reference is present, `CommandBars.FindControl`
resolves, and `.Execute` performs the built-in command.

**Compile VBAProject** is the same kind of built-in control. The documented
Office VBE control ID is **578**, but it is treated here as a candidate to
be proved, not a fact — exactly the discipline that produced ID 3.

### 3.1 Two risks, stated up front

**The control ID must be proved on this host.** Task 4 below enumerates
every VBE control ID and caption once and writes them to evidence. Only
after that log shows a control captioned "Compile VBAProject" is the
constant hard-coded, with the evidence path cited in the comment.

**A compile failure raises a modal dialog.** VBE reports the first compile
error in a message box and highlights the offending line. A script waiting
on `RunMacro2` will block until someone dismisses it. This cannot be
engineered away from inside VBA. The mitigations are:

- treat the runner as attended for the compile step — the user is at the
  machine when they launch it;
- the runner writes its log incrementally and flushes after every line, so
  everything up to the point of the dialog survives; and
- the PowerShell wrapper takes a timeout and reports
  `compile=DialogSuspected` if `RunMacro2` has not returned in time,
  instead of hanging silently.

**Pass/fail signal.** VBE disables the Compile control once the project is
fully compiled. So: `FindControl` → record `.Enabled` → `.Execute` →
`DoEvents` → `FindControl` again → record `.Enabled`. Still enabled after
execution means the project did not compile clean. This is a real signal
and it is what the stage reports; it is not a substitute for reading the
dialog when one appears.

## 4. Components

### Component A — `Module20_ProbeRunner.bas` (new, deployed)

Public surface:

| Procedure | Purpose |
|---|---|
| `R23_CompileProject` | Executes the VBE compile control and reports the enabled-before/after verdict. |
| `R23_EnumerateVbeControls` | One-off diagnostic: logs every VBE control ID and caption so the compile ID is proved rather than assumed. |
| `R23_RunAllProbes` | Opens the log, compiles, runs every read-only probe in order, closes the log. |

`R23_RunAllProbes` calls the nine existing probes in dependency order:

1. `Module12_FeatureQualification.R23_ProbeFeatureCatalog`
2. `Module13_ProjectionResolution.R23_ProbeViewProjections`
3. `Module14_AnnotationImport.R23_ProbeAnnotationReconciliation`
4. `Module15_OrdinateScheme.R23_ProbeOrdinateScheme`
5. `Module16_CalloutDefinition.R23_ProbeCalloutDefinition`
6. `Module17_SectionPath.R23_ProbeSectionPath`
7. `Module10_SectionDimensionEngine.R23_ProbeSectionDimensions`
8. `Module18_ContentEnvelope.R23_ProbeContentEnvelope`
9. `Module19_SemanticQA.R23_ProbeSemanticQA`

Each is wrapped so one probe's unhandled error cannot abort the rest; the
runner records `probe=<name>|status=Completed|Error:<n>` per entry.

**The runner must not contain a compile-order dependency on itself.** It is
the last module deployed and calls only public probe entry points.

### Component B — `Module21_EvidenceSink.bas` (new, deployed)

The piece that removes the paste step.

| Procedure | Purpose |
|---|---|
| `OpenLog(runFolder)` | Creates `test_assets/iteration_evidence/probe_runs/<yyyyMMdd_HHmmss>/probe_log.txt` and returns its path. |
| `LogLine(text)` | Writes to the file AND `Debug.Print`s, so the Immediate Window keeps working exactly as now. |
| `CloseLog` | Flushes and closes. |
| `IsOpen` | Lets `LogLine` degrade to `Debug.Print` alone when no run is active. |

Then retrofit: every `Debug.Print` in the nine probe modules, and
`CRunEvidence.AddInfo` / `AddWarning` / `AddFailure`, route through
`LogLine`. Roughly 60 call sites. Mechanical, and contract-tested by
asserting no raw `Debug.Print` survives in a probe module.

**Writing a file does not break the read-only claim.** `Agents.md` permits
"retain disposable drawings and evidence under `test_assets/`". The probes
still touch no model and no drawing, so `drawingUnchanged=True` continues to
mean what it has always meant. The runner must never write inside
`src/`, `test_assets/reference_drawings/`, or `src/baseline-model-dims/`.

### Component C — `tools/probe-runner/Run-R23Probes.ps1` (new)

One command, end to end:

1. optionally invoke `tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`
   (`-Deploy` switch; default is to assume the current SWP);
2. attach to the running SOLIDWORKS instance;
3. **open the authorized fixture part first, then the drawing** — via
   `ISldWorks.OpenDoc6`. This order is not cosmetic: it is what binds the
   drawing to `test_assets\models\P-0251-14A-001.SLDPRT` instead of the
   `V:\VEEMAP\SW_data\` sibling, and getting it wrong is what produced
   `R23_SECTIONDIM_FATAL|reason=UnauthorizedFixture` on 2026-08-02;
4. invoke `Module20_ProbeRunner.R23_RunAllProbes` through the existing
   `tools/swp-deploy/SolidWorksMacroInvoker.cs` (`RunMacro2`), with a
   timeout;
5. print the log path and the last N lines.

The invoker already exists and is already used by the deployment script for
the compile probe, so no new interop surface is introduced.

The fixture and drawing paths are **parameters with authorized defaults**,
validated against the same three-fixture list `Module1_Main` enforces. The
script must refuse an unlisted path rather than pass it through — the VBA
guard is the real gate, but a script that cheerfully opens an unauthorized
model and lets the macro reject it is worse than one that will not try.

## 5. Task list

- [ ] **PA-100:** Add `Module21_EvidenceSink.bas` with `OpenLog`,
  `LogLine`, `CloseLog`, `IsOpen`, degrading to `Debug.Print` when closed.
- [ ] **PA-101:** Route `CRunEvidence.AddInfo`/`AddWarning`/`AddFailure`
  through `LogLine`, preserving the existing `QA INFO:` / `QA WARNING:` /
  `QA FAILURE:` prefixes exactly — downstream parsing and every doc quote
  depend on them.
- [ ] **PA-102:** Replace every `Debug.Print` in the nine probe modules
  with `LogLine`. Contract: no raw `Debug.Print` remains in a
  `R23_Probe*` procedure.
- [ ] **PA-103:** Add `Module20_ProbeRunner.R23_EnumerateVbeControls` and
  run it once. Record the control IDs and captions under
  `test_assets/iteration_evidence/`. **Gate: do not proceed to PA-104
  until the compile control ID is proved from that log.**
- [ ] **PA-104:** Add `R23_CompileProject` using the proved ID, with the
  enabled-before/after verdict and the Module0 `Nothing`/`Enabled` guards.
- [ ] **PA-105:** Add `R23_RunAllProbes`: open log, compile, run all nine
  with per-probe error isolation, close log, print the log path last.
- [ ] **PA-106:** Add `tools/probe-runner/Run-R23Probes.ps1` per
  Component C, including fixture-path validation and the open-part-then-
  drawing order.
- [ ] **PA-107:** Register both new modules in
  `tools/swp-deploy/deployment-manifest.json` (36 to 38) and update the
  inventory-lock test.
- [ ] **PA-108:** Contract tests: sink hygiene, no raw `Debug.Print`, the
  runner calls all nine probes, the runner contains no `allowMutation`
  argument and no mutating procedure name.
- [ ] **PA-109:** Update `docs/CURRENT_STATUS.md`, `docs/Changelog.md`,
  `docs/SOLIDWORKS_API_VALIDATION.md` (the VBE CommandBars contract and the
  proved control ID) and `src/target-spec-hybrid-v2/README_IMPORT.md`.
- [ ] **PA-110:** First live run of `Run-R23Probes.ps1`; confirm the log
  file contains exactly what the Immediate Window shows, and that every
  probe still reports its read-only boundary.

## 6. What this does NOT remove

- **Production acceptance still requires the manual compile.** The
  amendment covers read-only probes only.
- **SOLIDWORKS must be running.** The script attaches; it does not launch.
- **A compile error still needs a human** to read the dialog and the
  highlighted line.
- **Screenshots and visual acceptance** stay manual, as they must.

The honest floor after this work: deploy and compile happen inside one
command, probes run unattended, and the agent reads the log file directly.
The user pastes nothing.

## 7. Acceptance

The tool is accepted when a single invocation of `Run-R23Probes.ps1`
produces a log file whose content matches the Immediate Window for the same
run, every probe reports `mutations=0` and `drawingUnchanged=True`, the
compile stage reports a verdict derived from the control's enabled state,
and the agent has read that file without the user copying anything.
