# Current Status

Date: 2026-08-08

## 2026-08-08: No material status change

Read-only turn. User asked to continue this chat in a separate Codex app -
no such transfer exists between the two products, so a self-contained
handoff brief was written to the session scratchpad
(`CODEX_HANDOFF.md`, outside this repo) and sent to the user. No file in
this repository was read, edited, deployed, or run. No new evidence.

## 2026-08-08: C9 form transfer resolved - all four failures were transfer errors, not source defects

No source edit, no revision change, no run this turn. Trunk unchanged since
`523dd52`; `MACRO_SOURCE_REVISION` still `trunk-2026-08-08-r26`.

The C9 form transfer is finally clean. Operator reports the forms compile and
behave correctly. All four failed attempts are now explained, and **none was a
defect in the delivered source**:

| Attempt | Symptom | Actual cause |
|---|---|---|
| 1 | `Sub or Function not defined` (`LoadPreferences`) | short paste - tail missing |
| 2 | `Invalid outside procedure` at `Ln 1` | paste missing the head |
| 3 | compiled, but `Section Settings` opened with no main form | `UserFormSection` code pasted into `UserForm1` |
| 4 | - | corrected, works |

Attempt 3's live symptom was diagnosed here as a stale modal blocking the
SOLIDWORKS UI thread. That reading was **wrong**; the operator found the real
cause. The dialog appeared without its parent because `UserForm1` contained
`UserFormSection`'s code, so `UserForm1.Show` was showing the section dialog.
Recording the bad call rather than quietly dropping it: the `ActiveDoc:
<none>` and "no QA report" evidence was consistent with both explanations,
and nothing was done to distinguish them before asserting one.

`Me.Hide` in `UserFormSection.DoOk` (the C9 fix, committed `523dd52`) is not
implicated in any of the four failures.

### Still unverified

The fix compiles and the dialog behaves, but **no run has yet created a
section view end-to-end through the form**. `section=ON` in a QA header and a
`Section View J-J` row in the view roster remain the acceptance evidence, and
neither exists yet.

### Verification gates

| Gate | Status |
|---|---|
| C9 forms transferred and compiling | **done** - operator-confirmed |
| Section dialog behaviour | **operator-confirmed** - not yet in a QA report |
| Deployment | done earlier this turn - `VERIFY: PASS`, `trunk-2026-08-08-r26` |
| Live macro execution producing evidence | **not achieved** - last invocation emitted no QA report |
| C9 verified live (`section=ON`, section row in roster) | **not verified** |
| Section view creation end-to-end | **still never achieved through the form** |
| Hole-callout root cause | **still not isolated** - 2 candidates, both need code |
| Forms importable via deployment manifest | **not attempted** - standing proposal, would retire this whole failure class |
| Companion test suite tracked in git | **no** - flagged, still undecided |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: C9 fix compiled and saved into Fable.swp; run blocked on an SWP lock

No source edit, no revision change. The trunk is unchanged since `523dd52`;
this turn moved that already-committed fix into the deployed macro file.

**The C9 transfer finally succeeded.** Operator reported a clean compile
after a full replace of both form modules, on the fourth attempt. The two
earlier failures were both transfer defects, now understood and recorded:
a short paste (missing tail, `Sub or Function not defined`) followed by a
paste missing the head (`Invalid outside procedure` at `Ln 1`). Neither was
a defect in the delivered source.

`Fable.swp` on disk confirmed written - mtime 04:59:16, checked 54 seconds
later, so the compiled forms are persisted rather than living only in the
VBE's in-memory project.

The deploy that would have exercised it then failed closed:

```
The target SWP is open or locked. Close Fable.swp in the VBA editor, but
leave SOLIDWORKS running, and retry.
Deploy-TargetSpecHybrid.ps1:172
```

Correct behaviour by the guard - deploying into a file held open by the VBE
is exactly how a half-written SWP would be produced. Nothing was written.

**The fix is compiled and saved but still unexercised.** No run has yet
created a section view through the dialog, so C9 remains unverified in the
only way that counts.

### Verification gates

| Gate | Status |
|---|---|
| C9 fix transferred into the SWP | **done** - clean compile, `Fable.swp` written to disk |
| VBA compile | **passing** - operator-reported, first clean compile of the C9 forms |
| Deployment | **blocked** - SWP locked open in the VBA editor, nothing written |
| Live macro execution | **not run** |
| C9 fix verified live (`section=ON`, section row in the list) | **not verified** - the point of the next run |
| Section view creation end-to-end | **still never achieved through the form** |
| Hole-callout root cause | **still not isolated** - 2 candidates, both need code |
| Forms importable via deployment manifest | **not attempted** - still the standing proposal if this recurs |
| Companion test suite tracked in git | **no** - flagged, still undecided |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: No material status change - C9 truncation confirmed, head now missing instead

Third handover attempt for the same C9 one-line fix. No source edit, no
deploy, no run, no revision change this turn.

**Truncation hypothesis confirmed** by the operator: the previous paste was
short, and the missing tail explained `Sub or Function not defined` on
`LoadPreferences`. That diagnosis was correct.

The repaste produced the mirror failure: `Compile error: Invalid outside
procedure` at `Ln 1`, `swPart` highlighted, procedure dropdown reading
`(General) (Declarations)`. Line 1 is `If swPart Is Nothing Then Exit Sub` -
a statement belonging inside `RefreshHoleCount`, sitting in no procedure at
all. The module now begins mid-procedure: the tail was repaired and the head
lost. Asked for a clean full replace with **both** ends verified this time -
`Ctrl+Home` must read `Option Explicit`, `Ctrl+End` must read `Ln 518`
(`Ln 127` for `UserFormSection`).

No code change was made in response to this error - it is a transfer defect,
not a defect in the delivered source, which compiles as a whole and is
unchanged since `523dd52`.

**Decision recorded for the next failure**: stop hand-carrying these. Both
forms build every control at runtime via `Controls.Add` and embed no designer
resources, so they most likely need no `.frx`, meaning proper `.frm` files
with a `VERSION 5.00` header could be generated and both forms added to
`deployment-manifest.json` - transferred by the deploy tool like every other
component. Unverified assumption (the no-`.frx` part), but it would retire
this entire failure class rather than re-rolling the same dice a fourth time.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| VBA compile | **failing** - `Invalid outside procedure` at Ln 1, transfer defect not source defect |
| Live macro execution | none this turn |
| C9 section-dialog fix, live | **not verified** - three transfer attempts, none complete |
| Section view creation end-to-end | **never yet achieved through the form** |
| Forms importable via deployment manifest | **not attempted** - proposed, `.frx` assumption unverified |
| Hole-callout root cause | **still not isolated** - 2 candidates, both need code |
| Companion test suite tracked in git | **no** - flagged twice now, still undecided |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: No material status change - C9 paste incomplete, diagnosis pending

Second failed handover attempt for the C9 section-dialog fix. No source edit,
no deploy, no run, no revision change this turn.

Operator pasted `UserForm1` and hit `Compile error: Sub or Function not
defined` on `LoadPreferences`, called from `UserForm_Initialize`.
`LoadPreferences` is defined at line 246 of the 518-line file, so the most
likely cause is a truncated paste rather than anything wrong with the code -
the declarations and `UserForm_Initialize` visible in the operator's
screenshot match the delivered file exactly, in order, so what did arrive is
correct as far as it goes.

**Not yet confirmed.** Asked for one decisive check - `Ctrl+End` in the code
window and read the `Ln` indicator: 518 means the paste completed and the
cause is something else, materially lower means truncation. Suggested
repasting from the file opened in Notepad rather than a preview pane, since
preview panes commonly clip long files, and gave the last procedure in each
form (`UserForm_QueryUnload`) as a visual end-of-file marker.

Two failed handovers in a row for the same one-line fix. Both failure modes
belong to the manual path only - the first was prose pasted as code, this one
is a probable clipboard truncation. Neither can occur on the deploy path,
which transfers files. If a third attempt fails, adding these two forms to
`deployment-manifest.json` (they currently lack the `VERSION 5.00` designer
block that would make them importable) is likely cheaper than continuing to
hand-transfer them.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| VBA compile | **failing** - `Sub or Function not defined` on `LoadPreferences`, cause not yet confirmed |
| Live macro execution | none this turn |
| C9 section-dialog fix, live | **not verified** - not yet successfully pasted |
| Paste-truncation hypothesis | **unconfirmed** - awaiting the `Ctrl+End` line count |
| Section view creation end-to-end | **never yet achieved through the form** |
| Hole-callout root cause | **still not isolated** - 2 candidates, both need code |
| Companion test suite tracked in git | **no** - flagged previously, still undecided |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: No material status change - delivered the C9 form code for manual paste

Handover turn for the C9 section-dialog fix committed in `523dd52`. No source
edit, no deploy, no run, no revision change.

First delivery attempt was a mistake worth recording: the fix was handed over
as a prose instruction document, and it was pasted wholesale into a VBE code
window, producing `Compile error: Expected: end of statement`. The document
was the input, not the code. Nothing was damaged - the last deploy had
already written and saved r26, so closing `Fable.swp` without saving discards
only the bad paste.

Redelivered as two complete attribute-stripped VBA files generated directly
from the trunk source (`UserForm1` 518 lines, `UserFormSection` 127 lines,
both verified zero non-ASCII bytes), to be pasted whole after select-all in
each code window. `UserForm1`'s only difference from what is already embedded
is `DoAddSection`; the rest is byte-identical, so the paste cannot introduce
drift.

Lesson, same shape as the A11 finding: **anything that cannot be deployed has
to be handed over as the artifact itself, not as instructions describing it.**
The deploy path never produces this failure mode because it transfers files;
only the manual path does, and these two forms are the only components on it.

Note also recorded during the commit: `tools/solidworks-automation-companion/`
is gitignored in its entirety and **no test in that suite has ever been
tracked** - including `test_userform_state_contracts.py` written this session.
The suite CLAUDE.md instructs everyone to run exists only in working copies.
Not changed unilaterally; flagged for a decision.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| VBA compile | **not run** - the C9 fix is still unpasted, so nothing has compiled it |
| Live macro execution | none this turn |
| C9 section-dialog fix, live | **not verified** - still requires the manual VBE paste |
| Section view creation end-to-end | **never yet achieved through the form** |
| Hole-callout root cause | **still not isolated** - 2 candidates, both need code |
| Companion test suite tracked in git | **no** - newly noticed, undecided |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: View-orientation candidate eliminated; section dialog bug found and fixed in source

Two findings, one live and one static. No revision bump - the only code
changed is in two `.frm` files, which are not deployable.

### Candidate (3) for the missing hole callouts is eliminated

Live `macro_qa/20260808_044112_P-0251-14A-001`, deploy `VERIFY: PASS` at
`trunk-2026-08-08-r26`, compile `Clean`, operator ticked Left/Right/Iso.

```
Views requested: front=ON top=ON bottom=off left=ON right=ON back=off iso=ON section=off
  Drawing View3 | *Right     | 0 dims, 0 callouts
  Drawing View4 | *Left      | 1 dims, 0 callouts
  Drawing View5 | *Isometric | 0 dims, 0 callouts
Hole callouts (IsHoleCallout=True) across drawing: 0 of 20 dimensions
FINDING: 10 Hole Wizard feature(s) present and swInsertholeCallout WAS
requested, yet InsertModelAnnotations4 produced no IsHoleCallout dimension.
```

The views where `4x Ø4.2/M5` reads as a circle now exist, and carry no
callout. **View orientation does not explain the missing callouts.** Two
candidates left, both needing code: `missing=Attachment`, and `AllViews=True`
versus per-view targeting (gap B3, the cheaper of the two).

Ordinate chains unaffected again - `160,90,50,10,0` / `36,15,0,15,36`, 8
created, 0 dangling.

### New gap C9: the section dialog discarded every result

Reported by the operator mid-run: picking a section and pressing OK adds
nothing to the list. Confirmed by the QA header (`section=off`) and located
statically.

`UserFormSection.DoOk` assigned its three result properties and then called
`Unload Me`. `Unload` destroys the form instance and resets every
module-level variable to its type default, so `UserForm1.DoAddSection` read
`Cancelled = False` - the Boolean default, indistinguishable from a genuine
OK - and `SectionLabel = ""`, then exited at its own `If Len(newLabel) = 0`
guard. Silent, no error. Cancel only appeared to work because it reached the
same empty string by a different route.

This is why **no run in this project has ever created a section through the
form**; every section view in earlier evidence came from a config path that
bypassed the dialog.

Fixed in source: `Me.Hide` on both exit paths, caller copies values out then
unloads. Regression test added, `tests/test_userform_state_contracts.py` -
verified to fail against the original code before being accepted (2 failures,
both naming the real defect). Suite 37 -> 40, all passing.

**Not deployable.** `UserForm1.frm` and `UserFormSection.frm` are outside
`deployment-manifest.json` with no `VERSION 5.00` designer block, so the fix
needs a manual VBE paste; instructions were handed to the operator. Nothing
live has exercised it.

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, `trunk-2026-08-08-r26` |
| VBA compile | pre-flight `verdict=Clean` |
| Live macro execution | ran, `PASS` |
| Offline suite | 40/40, including 3 new form-contract tests |
| New test proven against the real bug | yes - reintroduced the defect, confirmed the test fails |
| Hole-callout candidate (3), view orientation | **eliminated** by live evidence |
| Hole-callout root cause | **still not isolated** - 2 candidates, both need code |
| Section dialog fix | **fixed in source, NOT deployed, NOT run** - needs manual VBE paste |
| Section view creation end-to-end | **never yet achieved through the form** |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: No material status change - side-view callout run aborted before it started

Attempted the r26 deploy+run to test candidate (3) for the missing hole
callouts (view orientation - no run so far has created a side view, and the
reference places `4x Ø4.2/M5` there). The deploy failed immediately:

```
No running SOLIDWORKS instance is available. Start SOLIDWORKS and retry.
Deploy-TargetSpecHybrid.ps1:231
```

`Get-Process SLDWORKS` showed a process, but a **new** one (PID 31996,
started 04:38:33) rather than the instance earlier runs used (PID 22496) -
17s uptime, empty `MainWindowTitle`, not yet registered in the running
object table. The user confirmed they had closed and were relaunching
SOLIDWORKS. Waited; it became COM-reachable at 32s with
`P-0251-14A-001.SLDPRT [Viewing]` open.

**Nothing was deployed, run, or changed.** No source edits this turn; the
trunk remains at `trunk-2026-08-08-r26` as committed in `7aaa419`. The
`VERIFY: PASS` and QA evidence in the r26 entry below are from that
revision's own earlier run and are untouched by this turn.

Worth keeping: the failure is a plain startup race, not a tooling defect -
the deploy script's own guard caught it and refused rather than proceeding
against a half-loaded instance, which is the correct behaviour.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | **not run** - aborted, no SOLIDWORKS COM instance at invocation |
| VBA compile | **not run** |
| Live macro execution | **not run** |
| Side-view callout test (candidate 3) | **not run** - still the open next step |
| Root cause of zero hole callouts | **still not isolated** - 3 candidates, none tested |
| Everything the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: No material status change - dropped claude-code-router gateway (outside project)

User asked to remove the router (Fix B from the prior diagnosis entry).
Backed up `~/.claude/settings.json` to `settings.json.bak`, then removed the
`apiKeyHelper` line and the entire `env` block (`ANTHROPIC_BASE_URL`,
`ANTHROPIC_API_BASE_URL`, `CLAUDE_AGENT_API_BASE_URL`, `ANTHROPIC_MODEL`,
`CCR_CLAUDE_CODE_MODEL`, `CODEXL_CLAUDE_CODE_MODEL`,
`CLAUDE_CODE_ENABLE_GATEWAY_MODEL_DISCOVERY`) so the CLI talks to
`api.anthropic.com` directly instead of the dead `127.0.0.1:3456` gateway.
Not yet verified end-to-end: current process still holds the old env vars
in memory, so the fix only takes effect after a full CLI restart + `/login`
- told to the user, not yet confirmed by them.

Files outside this repo only: `~/.claude/settings.json`,
`~/.claude/settings.json.bak` (new). No file under this repo changed except
this entry.

### Verification gates

| Gate | Status |
|---|---|
| Offline unittest suite | not run this turn |
| Deployment | none this turn |
| Static/compile check | none this turn |
| Live macro execution | none this turn |
| Visual verification | none this turn |
| Push to origin/main | none this turn |
| CLI restart + working connection after edit | not confirmed by user yet |
| Everything the r24 gap-doc refresh lists as open | unchanged, still open |

## 2026-08-08: No material status change - CLI connectivity diagnosis only

Off-project turn. Diagnosed the user's `claude` CLI
`Unable to connect to API (ConnectionRefused)` error. Evidence gathered
read-only: nothing listening on TCP 3456
(`Get-NetTCPConnection -LocalPort 3456 -State Listen` empty), `ccr` absent
from PATH, `https://api.anthropic.com/v1/models` returned 401 (network
reachable, no key), and `~/.claude/settings.json` pins
`ANTHROPIC_BASE_URL`/`ANTHROPIC_API_BASE_URL`/`CLAUDE_AGENT_API_BASE_URL`
to `http://127.0.0.1:3456` with an `apiKeyHelper` pointing at
`%APPDATA%\claude-code-router\bin\`. Cause is the claude-code-router
desktop gateway not running.

No repo file changed except this entry. No settings edit made - both fixes
were offered and left for the user to choose.

### Verification gates

| Gate | Status |
|---|---|
| Offline unittest suite | not run this turn |
| Deployment | none this turn |
| Static/compile check | none this turn |
| Live macro execution | none this turn |
| Visual verification | none this turn |
| Push to origin/main | none this turn |
| Everything the r24 gap-doc refresh lists as open | unchanged, still open |

## 2026-08-06: No material status change

User answered the push-loop question: confirm each push individually,
don't auto-push trailing status commits. No repo, code, or push action this
turn - preference noted, saved to memory. `99dd49b` remains local, unpushed,
awaiting explicit confirmation as usual.

## 2026-08-06: Second follow-up push - same loop, flagged to the user

Pushed `d6ff71e..58dc066` at user request. `origin/main` now at `58dc066`.
No source, deploy, or run. This is the third status-log-then-push turn in a
row - the Stop gate requires logging every push, and each log commit is
then itself unpushed until asked. Flagged to the user this turn rather than
assumed either way: whether to keep pushing each status-only commit
automatically for the rest of this sequence, or stop here.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| Live macro execution | none this turn |
| Push to origin/main | confirmed - `d6ff71e..58dc066` |
| Everything the r24 gap-doc refresh lists as open | unchanged, still open |

## 2026-08-06: Follow-up push - the prior status entry itself reached origin/main

User asked to also push the just-created status commit (`d6ff71e`, logging
the earlier trunk-to-main push). Pushed: `4b53855..d6ff71e
trunk-baseline-r4 -> main`. `origin/main` now at `d6ff71e`. No source,
deploy, or run. This entry itself will be one commit ahead of origin/main
again the moment it's committed - normal for a doc-logging turn, not a gap.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| Live macro execution | none this turn |
| Push to origin/main | confirmed - `4b53855..d6ff71e` |
| Everything the r24 gap-doc refresh lists as open | unchanged, still open |

## 2026-08-06: Local trunk pushed to origin/main - repo, not code, change

At the user's explicit request ("commit the current state of the
repository to the main branch on github"). No source, deploy, or run this
turn - a git-history and repo-hygiene action only.

Before pushing, found and committed three files that had been running all
session but were never in git: `.claude/hooks/record_current_status_baseline.py`,
`.claude/hooks/require_current_status_update.py`,
`.claude/hooks/test_current_status_hooks.py` - the UserPromptSubmit/Stop
hook pair CLAUDE.md already documented as installed (it fired on every turn
this session). Also committed the `.claude/settings.json` wiring, the
CLAUDE.md section describing them, and `tools/swp-deploy/deployment-request.txt`
(the regenerated artifact from the r24 deploy - CLAUDE.md says never
hand-edit it). Commit `4b53855`. `.claude/hooks/test_current_status_hooks.py`
verified passing (5/5) before committing.

Explicitly excluded from staging: `docs/personal_review.md` (private, per
the user's own instruction from earlier this turn-chain), and all
`graphify-out/*` / `tools/graphify-solidworks/*` changes (another person's
in-progress work in this directory, per standing instruction - never swept
with `-A`, only explicit paths used).

`git fetch origin` confirmed `origin/main` was a strict ancestor of local
`trunk-baseline-r4` (0 commits on origin/main not in HEAD, 33 - then 34 with
this turn's commit - not yet on origin/main) - a clean fast-forward, no
merge, no conflict. Pushed `trunk-baseline-r4` to `origin/main`:
`e9783e5..4b53855`. Local checkout branch unchanged (still
`trunk-baseline-r4`); only the remote `main` ref moved.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| Live macro execution | none this turn |
| Local commit history clean, only this session's own files staged | confirmed - explicit paths, graphify and personal_review.md excluded |
| Fast-forward safety (no divergent history on origin/main) | confirmed via `git fetch` + `git merge-base --is-ancestor` before pushing |
| Push to origin/main | confirmed - `e9783e5..4b53855` |
| Hook test suite | 5/5 pass, `.claude/hooks/test_current_status_hooks.py` |
| Everything the r24 gap-doc refresh lists as open | unchanged, still open - this turn was git/repo hygiene only |

## 2026-08-06: Gap doc refreshed to r24; private review file created

`docs/BASELINE_TO_REFERENCE_DRAWING_GAP.md` was still dated r20 and had
drifted from the live evidence gathered r20-r24. Refreshed and committed
(`f4177d1`): closed C1/C2 (section cut placement, coordinate frame - both
closed by r24), added new A11 (HLV degrading the ordinate engine, closed
r23) and new C8 (reference has 4 views, recent runs create up to 8, and
side-face dimensions landed in the wrong named view bucket - `Bottom` not
`Left`). Corrected a stale reference to A8 (closed r10) still cited as open
in §8. D1 downgraded from "never written" to "populated in output, source
unverified" - recent screenshots show most title-block fields present, but
whether the code or template defaults are responsible was never checked.
Added a note that `Module5_FallbackDimensionEngine.InsertHoleCalloutsForView`
has zero callers and is a plausible existing candidate for the missing
hole-callout producer. Phase 4 (Tier C) marked explicitly not started, with
a note that no run has ever been made against the other two authorized
fixtures.

Also created `docs/personal_review.md` at the user's request - a private
snapshot of the same review, explicitly not to be linked, indexed, or cited
by Claude in future turns (saved to memory: `personal-review-file-private`).
Left untracked, not committed, per its stated private purpose.

No source file changed, no deploy, no run this turn. Offline suite reran as
a sanity check only (37/37, unchanged) - not evidence of anything new.

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn |
| Live macro execution | none this turn |
| Documentation accuracy | refreshed and committed - `BASELINE_TO_REFERENCE_DRAWING_GAP.md` now reflects r24 state |
| Everything the refreshed doc lists as open | still open - this was a documentation pass, not a fix |

## 2026-08-08: r26 - requested-config header closes the callout ambiguity on its first run

Added `Module4_ModelItemImporter.DescribeRequestedConfig`, reported at the
top of every QA report before any produced number: resolved model-item mask
with the callout bit called out by name, producer flags, all eight view
checkboxes, and HLR. Plus a classifier in `Module6_QAEngine` that turns a
zero-callout run into either a `FINDING` (requested but not produced - a
real API result) or a `NOTE` (not requested - not evidence of anything).

Deploy `VERIFY: PASS` at `trunk-2026-08-08-r26`, compile `Clean`, offline
suite 37/37. Live `macro_qa/20260808_042707_P-0251-14A-001`:

```
Model item mask: 1277992 (holeCallout bit ON)
Views requested: front=ON top=ON bottom=off left=off right=off back=off iso=off section=off
Display: HLR=off (ordinate harvest forces HLR regardless)

FINDING: 10 Hole Wizard feature(s) present and swInsertholeCallout WAS
requested, yet InsertModelAnnotations4 produced no IsHoleCallout dimension.
Not an operator setting.
```

**The r25 ambiguity is closed.** `1277992` decomposes exactly as
`8|32|32768|65536|131072|1048576` - the callout bit reached the API call.
SOLIDWORKS was asked, on a part with 10 Hole-Wizard features, and produced
none. That is now a real API finding rather than a possible unticked box,
and the run states it itself instead of relying on anyone's reading.

It also retroactively confirms the isometric diagnosis from the previous
turn: `iso=off section=off`, sticky registry checkboxes, no code fault.

Ordinate chains unaffected again - `160,90,50,10,0` / `36,15,0,15,36`, 8
created, 0 dangling, `HLR forced (was 1)` despite `HLR=off` requested,
which is A11 working as designed and now visible as such in one report.

### Next: three candidates, cheapest first

Recorded in `SOLIDWORKS_API_VALIDATION.md`. (1) `missing=Attachment`, the
archived tree's unreproduced `NATIVE_CALLOUT_COVERAGE` finding. (2)
`AllViews=True` versus per-view targeting, gap B3. (3) **View orientation** -
the reference puts `6x Ø6.6` on the front and `4x Ø4.2/M5` on the side, each
where that hole axis reads as a circle, and neither r25 nor r26 created a
side view at all. (3) needs no code, only more view boxes ticked, and would
eliminate or confirm orientation before code is written for (1) or (2).

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, embedded `trunk-2026-08-08-r26` |
| VBA compile | pre-flight `verdict=Clean` |
| Live macro execution | ran, `PASS` |
| Offline suite | 37/37 |
| Requested-config reporting | working, proven live on first run |
| Callout "declined vs never asked" ambiguity | **closed** - mask confirmed `1277992`, bit ON |
| Root cause of zero callouts | **still not isolated** - 3 candidates, none tested |
| Side-view callout behaviour | **never exercised** - no run has created a side view |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-08: r25 - hole callouts proven absent by direct query, dead callout code removed

Started working the gap doc's open items one by one, per the user's
request, MCP-first per the solidworks-api-lookup skill. First target: the
doc's own top lead, `Module5_FallbackDimensionEngine.InsertHoleCalloutsForView`
(zero callers).

`AddHoleCallout2` MCP-checked: its Remarks require a user to click OK in a
system dialog per call - not usable in this project's unattended
`tools/production-runner` path. The function also had no display-mode guard
(same class of bug A11 closed for ordinates) and no consolidation, so it
could never have produced the reference's grouped `6x`/`4x` callouts even if
the dialog problem didn't exist. Removed, not fixed.

Added real instrumentation instead of continuing to infer from screenshots:
`Module4_ModelItemImporter.CountHoleCalloutsInView` (`IDisplayDimension.IsHoleCallout`,
MCP-checked) and `Module3_ModelAudit.CountHoleWizardItems` (existing
per-feature `IsHoleWizardItem` flag, not previously summarized), surfaced in
`Module6_QAEngine`'s report.

Deploy `VERIFY: PASS` at `trunk-2026-08-06-r25`, offline suite 37/37 (after
fixing 3 non-ASCII bytes the encoding-contract test caught - `Section 7`
and `dia` instead of the section sign and diameter sign in new comments).
Live run `macro_qa/20260808_041847_P-0251-14A-001`:

```
Detected hole-like features: 12 (Hole Wizard: 10, plain cut: 2)
Hole callouts (IsHoleCallout=True) across drawing: 0 of 20 dimensions
```

Zero callouts, confirmed by direct API query on every dimension in the
drawing, despite the mask requesting them by default and 10 of 12 holes
being genuine Hole-Wizard features. Cause not yet isolated - ranked
candidates and the next diagnostic are in `SOLIDWORKS_API_VALIDATION.md`.
Ordinate chains unaffected: still exact match, 8 created, 0 dangling.

### Two operator questions answered mid-run, both traced to evidence

**"The isometric view disappeared?"** Not a regression. `UserForm1.frm:252`
reads `chkIso.Value = ReadBoolSetting("IsoView", True)` and line 313 writes
it back on every OK - the `True` is a first-ever-run fallback, not a
per-run default, so the box is whatever it was left at last time. r25
touched no view-creation code; `Module2_DrawingPipeline`'s `CreateIso`
branch is unchanged. The QA roster corroborates: this run created only
Front and Top, and Section is absent too - both checkbox-driven. The prior
run created 8 views from the same lineage.

**"I don't think the macro has stopped working yet."** It had.
`main returned.` was already in the terminal and a complete QA report had
been written and read before the question. `Get-Process SLDWORKS` confirmed
`Responding: True`, same instance since 03:54 - SOLIDWORKS idling with the
drawing open, which the macro never closes.

Both point at the same instrumentation gap, and it is the real shape of
**C8**: QA reports what was *created*, never what was *requested*. A
silently-unticked checkbox is indistinguishable from a creation failure in
the run record - exactly the blind spot that also prevents this run from
ruling out "`ImportHoleCallouts` was never `True`" as the cause of the zero
callouts above.

### Doc drift fixed while in the gap doc

Section 3's B-series table still described B1 (`DuplicateDims=False`) and
B2 (dead `ImportHoleCallouts` field) as open, though the document's own top
summary already listed both closed and the source has passed `True` /
gated the mask for many revisions. Missed on the last refresh pass.
Corrected, with the correction itself noted in the table.

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, embedded `trunk-2026-08-06-r25` |
| VBA compile | pre-flight `verdict=Clean` |
| Live macro execution | ran, `PASS` |
| Offline suite | 37/37, after an encoding-contract fix |
| Ordinate chains | unaffected, still match reference exactly |
| Hole-callout instrumentation | working, proven live - 0 of 20 |
| Root cause of zero callouts | **not isolated** - next task |
| Whether `ImportHoleCallouts` reached the mask this run | **not knowable** - nothing reports requested config |
| Everything else the r24 gap-doc refresh listed as open | unchanged |

## 2026-08-06: No material status change

Read-only review turn: full gap review of current output vs the reference
drawing, cross-checked against `BASELINE_TO_REFERENCE_DRAWING_GAP.md`,
`R23_SCOPE_AND_GENERALIZATION_PLANNING.md`, and recent QA reports/screenshots.
No code, deploy, or run. Findings delivered to the user in-chat, not written
to docs this turn. Two new items surfaced, not previously recorded anywhere:
`UserForm1.frm` sits outside `deployment-manifest.json` (no `VERSION 5.00`
block, cannot be mechanically deployed/verified - source may drift from what
`Fable.swp` actually runs), and `InsertHoleCalloutsForView` in Module5 has
zero callers in the trunk and is plausibly the unused hole-callout producer.

## 2026-08-06: No material status change

Turn was a caveman-mode level switch only. No code, docs, or run.

## 2026-08-06: Operator run with Top/Bottom/Left/Back enabled - partial side-view match, wrong view bucket

Operator-driven run, `macro_qa/20260806_170652_P-0251-14A-001`, no redeploy -
same `trunk-2026-08-06-r24` binary, confirmed by unchanged section signature
(`HLR forced (was 1)`, stepped 3-segment cut, identical readback). The
operator ticked more of the existing view checkboxes in the form; no source
was edited this turn.

```
Front:  13 dims, 0 dangling - 10,50,90,160 / 15,36,15,36 / 30,36,1,1,45
Top:     0 dims
Bottom:  6 dims, 0 dangling - 5.00, 5.00, 4.20, 6.00, 80.00, 25.00
Right:   0 dims
Left:    1 dim,  0 dangling - 72.00
Back:    0 dims
Section: unchanged from r24's first run - 13 dims, stepped cut intact
```

Imported model items 19 -> 25, total dimensions 27 -> 33. 0 dangling
everywhere.

### Reference cross-check: real overlap, wrong view

The reference's side elevation carries `173.6, 104.8, 80, 25, 11.5, 6`, plus
callouts `4x Ø4.2 T12.4 / M5x0.8-6H T10` and `6x Ø6.6 THRU / T Ø11 T6`.
`80.00` and `25.00` now appear verbatim, alongside `4.20` and `6.00` - but on
**Bottom**, not **Left**. This macro's Top/Bottom/Left/Right naming is
SOLIDWORKS's standard projection relative to the Front view; it is not
guaranteed to land the same physical face in the same named bucket as the
reference sheet's own layout. The overlap is real evidence the geometry is
reachable; the bucket it lands in is not yet controlled.

Still absent from every view: `173.6`, `104.8`, `11.5`, and both hole
callouts in full (`4x Ø4.2...`, `6x Ø6.6 THRU...` with their leader
formatting - `Module4_ModelItemImporter` imports plain dimensions, not
callout-style annotations).

### Verification gates

| Gate | Status |
|---|---|
| Deployment | none this turn - same r24 binary as prior entry |
| Live macro execution | ran, `PASS`, operator-driven |
| Front/section regression check | unaffected, values unchanged from r24's first run |
| Side-face dimensions reachable | confirmed - `80.00`/`25.00`/`4.20`/`6.00` present |
| Side-face dimensions in the reference's expected view bucket | **not achieved** - landed on Bottom |
| Full side-face coverage (`173.6`, `104.8`, `11.5`) | **not achieved** |
| Hole callout formatting (`4x Ø.../ THRU` style) | **not achieved** - importer produces plain dimensions |
| Right view, Top view | still 0 dimensions each |
| Section portrait orientation / GENERAL NOTES overlap / mass units / duplicate notes | unaddressed, unchanged |

## 2026-08-06: No material status change

Read-only turn. The user is rerunning the macro manually with a side view
enabled; Claude did not deploy, run, or edit source this turn - only pointed
out that the reference's left view carries `80, 25, 6` and that `CreateLeft`
and `CreateRight` are distinct flags in `Module1_Main.GlobalConfig`, both
False in `ResetGlobalConfig`. No new evidence, no gate changes. Awaiting the
user's QA report or screenshot from their run.

## r24 live: stepped section cut succeeds, coordinate frame confirmed

Run `macro_qa/20260806_165529_P-0251-14A-001`. Deploy `VERIFY: PASS`, embedded
`trunk-2026-08-06-r24`, 13/13 managed components matching source. Pre-flight
`verdict=Clean`. 27 dimensions, `PASS`.

### Ordinate chains: unaffected, still exact

`Ordinate edges seen: 39 (circular: 22, of which arcs: 4, linear: 17)`,
`Harvest display mode: HLR forced (was 1)`, X `160,90,50,10,0`, Y `36,15,0,15,36`,
8 created, 0 dangling. The section-cut change touches a different code path
and did not regress the r23 result.

### Section cut: stepped, and the coordinate-frame question is closed

```
Section cut: stepped, 3 segments
  bore leg at 0.mm, row leg at 15.mm, jog at -27.mm
  sketch readback (mm): -127.4,0. -27.,0. -27.,15.
  view scale 0.6667 - readback matching the requested values proves the sketch frame is model-scale
```

The three logged points are the start point of each of the three
`SketchLine` segments (bore leg, jog, row leg), read back via
`GetStartPoint2` rather than trusted from the arguments. They match the
requested along/across values with no 0.6667 scale factor applied anywhere.
This settles the question CodeStack row 9 raised: drawing-view sketch
geometry in this SW2025 install is stored at model scale; the view's own
scale is a display-time transform, not a coordinate-system change.
Previously unverified because every prior section cut sat at an offset of 0,
which reads back as 0 under any scale and proved nothing.

`CreateSectionViewAt5` accepted three separately-created, separately-selected
`SketchLine` segments as one section line and returned a non-Nothing view.
The Remarks phrase "the section line or lines" (MCP, 2026-08-06) is now
confirmed to mean what it says.

Section view dimension count rose 11 to 13 (new `40.00` and `10.00`
present); front view fell 16 to 14. Consistent with dimensions that used to
land on the front view now landing on the section instead, now that the cut
passes through the hole row - not independently confirmed, text readback
alone cannot distinguish this from a placement change. 0 dangling in both
views.

### Operator screenshot: confirms border containment, front view matches reference

The operator's post-run screenshot shows the front view matching the
reference drawing closely: `160.00, 90.00, 50.00, 10.00, 0` across the top,
`36.00, 15.00, 0, 15.00, 36.00` down the right, `R36.00`, `1.00`, `40.00`,
`30.00`, `45deg`, six-hole pattern. Section J-J now sits inside the sheet
border - the r23 overflow (12.00/18.00/12.00 dimensions running past the
frame edge) is gone. Section still renders landscape (wide), not the
reference's portrait orientation.

**Not confirmed from the screenshot at this resolution:** whether the
section's cross-section shape itself changed versus r23 in a way traceable
to the new cut path, as opposed to the same shape with one more dimension
labelled. The sketch-readback evidence above proves the cut *line* is
correctly stepped; it does not by itself prove the resulting section *view*
renders that cut correctly. Treated as proven by the readback plus the
absence of any error, not by pixel-level comparison to the reference.

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, embedded `trunk-2026-08-06-r24` |
| VBA compile | pre-flight `verdict=Clean` |
| Live macro execution | ran, `PASS` |
| Ordinate chains | unaffected, still match reference exactly |
| Section stepped-cut creation | succeeded - 3 segments, non-Nothing view |
| Section coordinate-frame scale | proven - readback matches requested values, no scale factor |
| Section within sheet border | confirmed by screenshot |
| Section geometry correctness (precise visual match to reference) | not verified at pixel level |
| Section portrait orientation | not addressed - still landscape |
| Right view 0 dimensions / GENERAL NOTES overlap / mass units / duplicate notes | unaddressed, unchanged from prior entry |

## r24 source written, deploy interrupted before any live evidence

`src/baseline-model-dims/Module2_DrawingPipeline.bas` gained a stepped
("jogged") section-cut planner and `Module1_Main.bas` was bumped to
`MACRO_SOURCE_REVISION = "trunk-2026-08-06-r24"`. **Both files are
uncommitted. Neither has been deployed. r24 has never run against
SOLIDWORKS.** The production run was launched and the user interrupted the
tool call before it executed; no deploy evidence, no QA report, no sheet
exists for r24. Do not read anything below as a result - it is a description
of source code only.

### What the change does, and why

Reading the reference drawing (`test_assets/reference_drawings/P-0251-14A-001.PNG`)
alongside the r23 sheet showed Section J-J is a **stepped cut**: it runs
through the R36 bore centre, jogs sideways once, and continues through one row
of the counterbored holes. The trunk's section cut through r23 was a single
straight line at `midY`, which is the model's mirror line - it passes between
the two hole rows (at Y = ±15) and catches neither.

New `PlanSteppedCut` reads `Module3_ModelAudit.GetAllHoleLikeFeatures` and
picks the widest hole-like feature as "the bore," then the across-axis
coordinate shared by the most non-bore holes as "the row." `CreateSectionFromConfig`
now draws three `SketchLine` segments (bore leg, jog, row leg) and selects all
three before calling `CreateSectionViewAt5` - its Remarks say to select "the
section line **or lines**" (MCP-checked 2026-08-06), which is read as license
for a multi-segment cut under the existing `swCreateSectionView_NotAligned`
flag. `swCreateSectionView_OffsetSection` was checked and rejected: its
description is "an aligned section view ... two lines at an angle," which MCP
enum text ties to a revolved section, not a step.

Also changed: the section view's sheet placement, previously
`frontPos(0) + 0.18` - a fixed 180 mm offset blind to either view's actual
width, which is what ran the section past the sheet border on the r23 sheet.
Now `halfFront + halfSection + 15mm gutter`, derived from the part bounding
box and the sheet scale.

### The coordinate-frame question this is designed to answer

CodeStack row 9 warns that drawing-view sketch entities live in the view's own
frame and to verify the transform direction on SW2025. Every section cut before
r24 was placed at `midY`, which for this fixture is 0 - and 0 reads back as 0
under any scale, so it was never evidence of which frame the sketch used. r24
is the first cut with a nonzero offset. `AddCutSegment` reads each segment's
endpoints back via `GetStartPoint2` and reports them in the new
`Section cut:` QA block, specifically so the live run can show whether the
sketch kept model-mm coordinates or something scaled by the 0.6667 view
factor. **This has not run. The frame assumption is still unverified.**

### Verification gates

| Gate | Status |
|---|---|
| Offline unit suite | ran, 37/37 pass (unchanged count - no new offline test added for the section planner) |
| Deployment + readback | **not run** |
| VBA compile | **not run** - no pre-flight this turn |
| Live macro execution | **not run** |
| Section stepped-cut geometry vs reference | **not run** |
| Section sketch coordinate-frame scale | **not run** - this is the open question the change exists to answer |
| Section placement within sheet border | **not run** |
| Visual acceptance | **not run** |
| Right view / Section portrait orientation / GENERAL NOTES overlap / mass units / duplicate notes | unaddressed, unchanged from prior entry |

### Next step

Rerun `tools/production-runner/Run-R23Production.ps1 -AllowMutation -Deploy`
to get first live evidence for r24: deploy readback, compile verdict, the QA
report's new `Section cut:` block, and a look at the sheet.

## r23 second run: the display-mode guard is proven

Run `macro_qa/20260806_152553_P-0251-14A-001`, deploy `VERIFY: PASS` at
`trunk-2026-08-06-r23`. Deliberate experiment: the operator **unticked HLR**,
so the form requested HLV.

```
Harvest display mode: HLR forced (was 1)
Ordinate edges seen: 39 (circular: 22, of which arcs: 4, linear: 17)
X: 5 stations, datum Edge:offsetFromTarget=160.00mm  ->  10, 50, 90, 160
Y: 5 stations, datum Hole:offsetFromTarget=0.00mm    ->  15, 36, 15, 36
Ordinate display dimensions actually created: 8
readback: 16 dims, 0 dangling
```

`was 1` is `swHIDDEN_GREYED`, HLV. The engine overrode the form and harvested
under HLR anyway. Output is byte-identical to the run where the operator ticked
HLR. **The checkbox no longer reaches ordinate quality**, which was the point.

Both halves of the guard are now covered by live evidence: `HLR (already)` on
the ticked run, `HLR forced (was 1)` on the unticked one.

### Closed: ordinates survive the restore to HLV

`SOLIDWORKS_API_VALIDATION.md` recorded this as unverified when the guard was
written. The prune runs after `ForceRebuild3`, when `IAnnotation.IsDangling` is
meaningful, and reports **0 dangling**. An ordinate created against HLR
geometry stays attached when the view is restored to HLV.

### Still not instrumented

Whether `RestoreDisplayMode` succeeded. It is called with `On Error Resume
Next` and reports nothing. Inferable only from the sheet showing hidden lines.

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, embedded `trunk-2026-08-06-r23` |
| VBA compile | pre-flight `verdict=Clean` |
| Live macro execution | ran, `PASS` |
| Display-mode forcing path | **proven** - `HLR forced (was 1)`, 39 edges, 160 mm datum |
| Ordinate survival across restore | **proven** - 0 dangling post-rebuild |
| `RestoreDisplayMode` success | not instrumented |
| Section cut misses every hole | **not addressed** - next |
| Right view carries 0 dimensions | **not addressed** |
| Section J-J landscape vs reference portrait | **not addressed** |
| GENERAL NOTES overlaps the title block | **not addressed** |
| Mass units, duplicate notes | **awaiting user decision** |

## r23 live: best run to date - both chains match, retry gone, guard unproven

Run `macro_qa/20260806_151955_P-0251-14A-001`. Deploy `VERIFY: PASS`, embedded
`trunk-2026-08-06-r23`, 13/13 managed components matching source. 27
dimensions, `PASS`.

```
X: 5 stations (holes=4, edges=1), datum Edge:offsetFromTarget=160.00mm
   10.00, 50.00, 90.00, 160.00
Y: 5 stations (holes=3, edges=2), datum Hole:offsetFromTarget=0.00mm
   15.00, 36.00, 15.00, 36.00
Ordinate display dimensions actually created: 8
Ordinate edges seen: 39 (circular: 22, of which arcs: 4, linear: 17)
readback: 16 dims, 0 dangling
```

Both ordinate chains match the reference. The creation counter is exact:
X 5 stations minus datum = 4, Y 5 minus datum = 4, total 8.

### The holes-only retry is gone

r20 produced the same chains but needed `RETRY holes-only after code 1: code 0`
on the Y axis. r23 creates both chains on the first attempt. Combined with r22,
this settles it: **running ordinates before model import removes the failure,
and the imported dimensions were the cause.** The retry path stays in the code
as a fallback but no longer fires on this fixture.

### Front view is less crowded

16 dimensions against r22's 22, and the `14.10` / `94.10` mismatched values are
absent. Consistent with those being an HLV artifact - not independently proven,
since the only change that could produce them is also the display mode.

### The display-mode guard did not fire

r23 added `ForceHlrForHarvest` in `Module2_DrawingPipeline`: read
`GetDisplayMode2`, force `swHIDDEN` (2) if the view is in any other mode,
`UpdateViewDisplayGeometry`, harvest, restore. Reported as
`Harvest display mode:` in the QA report.

This run reported `HLR (already)` - the operator ticked HLR, so the forcing
path never executed. **The guard is unproven.** Proving it needs one run with
HLR unticked: a working guard yields `HLR forced (was 1)`, 39 edges, and the
160 mm datum despite the checkbox.

### Why the guard exists

`UserForm1.frm:265` reads `chkHLR.Value = ReadBoolSetting("UseHLR", False)` and
`SaveSetting` persists the operator's last answer. r17 made HLR the default only
in `ResetGlobalConfig`, which is the no-form fallback, so every form run
defaulted to HLV. Under HLV the engine degrades two ways: hidden-line edges
that select and then dangle (r17), and a candidate pool of 64 edges instead of
39 that hands `ResolveOneDatum` a nearer straight edge, putting the X datum
43 mm inside the part (r22). A geometry precondition should not be reachable
from a checkbox.

### A wasted run, recorded

The first r23 attempt invoked `Run-R23Production.ps1 -AllowMutation` without
`-Deploy`, which is opt-in. r22 executed again and returned byte-identical
output. Pre-flight `verdict=Clean` in that run compiled the deployed r22 and
said nothing about the r23 edits. Evidence dir `macro_qa/20260806_151854` is a
duplicate of r22, not an r23 result.

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, embedded `trunk-2026-08-06-r23` |
| VBA compile | pre-flight `verdict=Clean` against r23 |
| Live macro execution | ran, `PASS` |
| Creation readback | 8 created, matches station arithmetic |
| Both chains vs reference | match |
| Display-mode forcing path | **unproven** - needs a run with HLR unticked |
| Cause of the 2.1 mm value mismatch | **not proven**, only absent under HLR |
| Dimension collision / placement (gap A7) | **not addressed** |
| Section cut misses every hole | **not addressed** |

## r22 live: reorder fixes the Y chain; run was HLV so X is not comparable

Run `macro_qa/20260806_150931_P-0251-14A-001`. Deploy `VERIFY: PASS` at r22,
pre-flight `verdict=Clean`, 33 dimensions, `PASS`.

### Confirmed: running ordinates before import restores the Y chain

`15, 23.60, 36.00, 15, 23.60, 36.00` on the front view, `+/-36` included, and
the operator screenshot shows the full chain `36, 23.60, 15, 0, 15, 23.60, 36`
down the right side. r21's Y chain was absent entirely.

**The imported dimensions were the cause.** The r21 hypothesis is now
supported by an A/B with the ordering as the only change.

### The new r22 counter works

`Ordinate display dimensions actually created: 14`. X 9 stations minus datum
= 8, Y 7 minus datum = 6, total 14. Exact. `Ordinate chains created: 2 of 2`
is now corroborated by a creation count rather than a return code.

### This run was HLV, not HLR - X is not comparable to r20

`Ordinate edges seen: 64 (circular: 35, of which arcs: 4, linear: 29)` is the
documented HLV signature; HLR is `39 (22, 4, 17)`.

| | r20 (HLR) | r22 (this run, HLV) |
|---|---|---|
| Edges | 39 | 64 |
| X stations | 5 | 9 |
| X datum | `Edge:offsetFromTarget=160.00mm` | `Edge:offsetFromTarget=43.00mm` |

The extra hidden-line edges gave `ResolveOneDatum` a nearer straight edge, so
the end datum landed **43 mm inside the part** instead of on the end face.
X reads `43, 0, 1, 12, 27, 67, 92, 107, 117` against the reference's
`0, 10, 50, 90, 160`. This is the r17 finding recurring in a new form: under
HLV the ordinate engine degrades, previously by dangling, now by datum choice.

### New failure mode: rendered value differs from computed station

The sheet renders `14.10` and `94.10` where the computed stations are `12` and
`92` - a ~2.1 mm mismatch on two X ordinates, drawn in a distinct colour -
while `IAnnotation.IsDangling` reports **false for all 22** dimensions and the
r22 creation counter reports the expected 14.

So neither existing instrument detects this. It is **not** the `0.00` dangling
of r10-r15. **Cause unknown, not probed.**

### Verification gates

| Gate | Status |
|---|---|
| Deployment + readback | `VERIFY: PASS`, embedded `trunk-2026-08-06-r22` |
| VBA compile | pre-flight `verdict=Clean` |
| Live macro execution | ran, `PASS` verdict |
| Creation readback | 14 created, matches station arithmetic |
| Visual acceptance | done, **fails** - X datum wrong, two values mismatched, front view crowded |
| HLR comparison run | **not done** - this run was HLV |
| Cause of the 2.1 mm value mismatch | **not probed** |
| Dimension collision / placement (gap A7) | **not addressed** |

No source file changed this turn. `MACRO_SOURCE_REVISION` remains
`trunk-2026-08-06-r22`.

## 2026-08-06: r22 redeploy launched - no material status change yet

The operator closed `Fable.swp` in the VBA editor and a redeploy-and-run of
r22 was launched. **The run had not returned when this entry was written**, so
there is no deployment verdict, no compile verdict, no QA report and no
outcome for the Y chain.

No source file changed this turn. `MACRO_SOURCE_REVISION` remains
`trunk-2026-08-06-r22`. Every gate listed in the r22 entry below is still
outstanding; nothing here supersedes it.

What the next report should answer:

- `Ordinate display dimensions actually created: N` - the new r22 counter.
  Roughly 8 if both chains land, since datum ordinates are not returned by
  `GetDisplayDimensionCount`.
- Whether the Y chain survives now that ordinates run before import. Success
  implicates the imported `72.00` spanning the `+/-36` silhouette edges;
  failure exonerates it and the cause lies elsewhere.

## r22 source: ordinates before import, and creation is measured not inferred

Source change. Deployed and a run was launched at the end of this turn; **no
result had returned when this entry was written**, so nothing below claims a
live outcome.

### The r21 suspect was wrong, and our own counter refuted it

The previous entry blamed native import for consuming the silhouette edges,
citing shared-entity substitutions rising 2 -> 4. That rise is entirely the
retry's own arithmetic:

- first Y attempt: 5 stations, **2** substitutions - identical to r20
- `CreateChainWithFallback` rolls `usedCount` back to the post-X state, the
  holes-only retry re-claims 3 holes, the same 2 collide again -> **+2**

Candidate collection was byte-identical between r20 and r21: 39 edges
(22 circular, 4 arcs, 17 linear), X 5 stations, Y 5 stations, same datums,
same 2 outer-edge promotions. **Import changed nothing about the candidates.**

### What remains, stated as hypothesis

The only difference when the Y chain ran was that 19 imported dimensions
already existed in the view, including a `72.00` spanning exactly the two
`+/-36` silhouette edges the Y chain uses. Whether an entity that already
carries a dimension can still anchor an ordinate is **not documented** -
neither `AddOrdinateDimension` nor `InsertModelAnnotations4` mentions it
(MCP, 2026-08-06). Not asserted as the cause.

### Changes

- `Module2_DrawingPipeline`: the ordinate engine now runs **before** model
  import. Removes the interaction, and is the correct order regardless since
  import is additive and has no such precondition. This ordering is also the
  experiment - Y succeeding implicates the imported dimensions, Y still
  failing exonerates them.
- `Module5_FallbackDimensionEngine`: `IView.GetDisplayDimensionCount` is read
  either side of `AddOrdinateDimension`. A success return with a zero delta is
  now reported as `CHAIN_CREATED_NOTHING` rather than counted as a created
  chain. New report line `Ordinate display dimensions actually created: N`.
  This closes the exact hole r21 fell through: `2 of 2 created` against a
  sheet carrying no Y chain.
- `Module1_Main`: `MACRO_SOURCE_REVISION` -> `trunk-2026-08-06-r22`
  (deployable behaviour changed).

### Verification gates

| Gate | Status |
|---|---|
| Offline companion suite | 37/37 pass |
| Deployment + readback | **FAILED** - `Fable.swp` locked, open in the VBA editor after the manual form paste. Nothing was written. |
| VBA compile | **not run** - deploy aborted first |
| Live macro execution | **not run** |
| Y chain restored | **unverified** |
| Cause of the `OrdFailure` | **still unproven** - the reorder is a test, not a diagnosis |
| Visual acceptance | **not done** |
| Dimension collision / placement (gap A7) | **not addressed** |

### Correction

An earlier draft of this entry claimed deployment and compile had passed. They
had not: the deploy aborted with `The target SWP is open or locked`, because
`Fable.swp` was still open in the VBA editor from the manual `UserForm1` paste.
No bytes were written to the SWP and `main` never ran. r22 is source-only and
**entirely unexercised**.

## r21 visual: the Y ordinate chain is absent, and the report says it succeeded

Screenshot of run `macro_qa/20260806_144628`, supplied by the operator.

**The Y chain is not on the sheet.** In r20 it ran down the right side of the
front view as `36, 15, 0, 15, 36`. In this sheet that column is gone; the right
side carries only imported `1.00`, `10.00`, `40.00` and a vertical `72`.

Accounting the front view's 12 readback values: X chain `10, 50, 90, 160` (4)
plus imported `72, 40, 30, 10, 36, 1, 1, 45` (8) = 12. **None left for the Y
chain.** The `36.00` is the imported `R36.00` rendered on the sheet, not an
ordinate station.

So the holes-only retry returned `code 0` (`swCreateOrdDimErr_Success`) and
created nothing. Precedent exists in the archived tree:
`swCreateOrdDimErr_Success|createdReadBack=0`.

**`Ordinate chains created: 2 of 2 attempted` is therefore false**, and the
dimension readback did not catch it.

### The instrument gap

`DescribeDimensionReadback` counts dimensions per view and reads
`IAnnotation.IsDangling`, but **cannot attribute a dimension to a producer**.
While model import was off, subtraction identified the ordinates; with both
producers active it cannot. A success return code was trusted where a
creation readback was needed - the same class of error as counting dimensions
without checking attachment (r10-r15).

Candidate fix, **not implemented**: capture created ordinate dimension names
at creation time and report per-producer counts, so the chain count is checked
against dimensions that exist rather than against a return value.

### Also visible, not yet addressed

- Front view dimensions overlap the part outline and each other (`45` over the
  profile, `36.00`/`0.50` colliding near the section). Gap A7, no placement or
  collision handling exists.
- Section J-J dimensions cross the `SECTION J-J` label.
- Section J-J is still landscape 196 wide; the reference is portrait 173.6.

### Verification gates

| Gate | Status |
|---|---|
| Visual acceptance | done, **fails** - Y chain absent, dimensions collide |
| Cause of `OrdFailure` on the first Y attempt | **not probed** |
| Why the holes-only retry created nothing | **not probed** |
| Per-producer dimension attribution | **not implemented** |

No source file changed this turn. `MACRO_SOURCE_REVISION` remains
`trunk-2026-08-06-r21`.

## r21 live: both producers ran together; the Y ordinate chain lost its edges

Run `macro_qa/20260806_144628_P-0251-14A-001`, after the operator pasted the
updated `UserForm1` into the VBE. First drawing in this project's history with
native import and the ordinate engine both active.

```
Imported model items: 19
Total drawing view dimensions: 23
  Drawing View1  (Front)          12 dims, 0 dangling
  Section View J-J                11 dims, 0 dangling
Ordinate chains created: 2 of 2 attempted
Shared entities: 4 substituted, 0 still shared
  Y: RETRY holes-only after code 1: code 0
NOTE: 1 chain(s) fell back to holes-only candidates.
```

### New regression: the holes-only fallback fired for the first time

The Y chain failed with `swCreateOrdDimErr_OrdFailure` (code 1) and only
succeeded on the holes-only retry. That fallback has existed since r10 and had
**never fired** in any prior run; r20 built both chains with straight edges
intact.

Shared-entity substitutions rose **2 -> 4** in the same run. Working
hypothesis, **not probed**: native import now runs first and consumes model
entities, so the two silhouette edges at `+/-36` were already taken when
`AddOrdinateDimension` tried to use them.

Consequence if that holds: the `+/-36` silhouette stations are absent from the
sheet and the Y chain is holes-only. The report's `offsets(mm)` line still
shows `-36, -15, 0, 15, 36` because it is computed **before** creation and
describes intent, not what landed. **No visual confirmation yet** - the
`readback` live values cannot be attributed to a chain without a screenshot.

### Front view is crowded

Front carries 12 dimensions: the X chain (`10, 50, 90, 160`) plus imported
`72, 40, 30, 10, 36, 1, 1, 45`. The reference shows only `R36`, a hole callout
and the two chains there. `1.00` twice and `45.00` are likely chamfer sketch
dimensions. No placement or collision handling exists (gap A7).

### Verification gates for this run

| Gate | Status |
|---|---|
| Live macro execution | ran, `PASS` verdict |
| Dimension readback | 23 dims, 0 dangling |
| Offline companion suite | 37/37 pass (unchanged from r21 source) |
| Deployment + readback | **not run this turn** - form was pasted manually, no redeploy |
| Visual acceptance | **not done** - no screenshot of this sheet |
| `+/-36` station presence | **unverified** |
| Cause of the `OrdFailure` | **not probed** |

No source file changed this turn. `MACRO_SOURCE_REVISION` remains
`trunk-2026-08-06-r21`.

## r21 source: the two dimension producers can finally run together

**Not deployed, not run.** Source change only; the form half needs a manual
paste (below).

`UserForm1` carried `optDimModel` and `optDimOrdinate` as mutually exclusive
members of one `"DimMode"` option group. Native import and the ordinate engine
were therefore **structurally impossible to combine** - and the reference
drawing needs both at once. Group removed; both producers always run.

- `UserForm1.frm`: option buttons, their `DimStyle` settings read/write, and
  the `optDim*.Value` assignments deleted. `GlobalConfig.UseModelDimensions`
  and `.UseOrdinateDims` set `True` unconditionally.
- `Module1_Main.ResetGlobalConfig`: `.UseOrdinateDims` `False` -> `True`.

**The form is outside `deployment-manifest.json` and cannot be imported.** The
`.frm` in source is the record; the running form must be updated by pasting
the generated code into the VBE. Until then the deployed macro still shows the
old option group.

**Verification gates for r21:**

| Gate | Status |
|---|---|
| Offline companion suite | 37/37 pass |
| Residual-reference grep (`optDimModel`, `optDimOrdinate`, `DimStyle`) | clean, no hits |
| VBA compile in SOLIDWORKS | **not run** |
| Deployment + readback | **not run** |
| Live macro execution | **not run** |
| Visual acceptance | **not run** |
| Manual VBE paste of `UserForm1` | **not done** |

`MACRO_SOURCE_REVISION` was bumped to `trunk-2026-08-06-r21` because
deployable behaviour in `Module1_Main` changed, but nothing has been deployed.

### Native import coverage, measured 2026-08-06

First run this session with model dims actually enabled
(`macro_qa/20260806_133255`): **19 imported items**, several matching the
reference.

| Imported | Reference | View |
|---|---|---|
| 47.00, 40.00 | Ø47, Ø40 | Section J-J |
| 18.00, 12.00 | 18, 12 | Section J-J |
| 36.00 (`R36.00` on sheet) | R36 | Front |

Section 7 of the gap doc was **too pessimistic**: those values are marked in
the model and arrive free. They were missing only because every prior run had
`UseModelDimensions = False`.

Still absent, two different reasons:

- `80, 25, 6` - live on the **Left view**, which this trunk never creates
  (`CreateLeft = False`). Config, not architecture.
- `173.6, 104.8, 11.5` and the `H7` on the 47 - not reached. `H7` is proved
  drawing-authored, not model-sourced.

### Hypothesis: the section cut misses every hole

**Inference from two screenshots plus r20 ordinate data. Not probed.**

`CreateSectionFromConfig` draws a horizontal cut at
`midY = (bbox(1)+bbox(4))/2`. The r20 Y chain measured hole rows at exactly
`±15` with `0` empty between them, so the cut passes through the gap and
touches no hole. Our section is a clean hatched shaft; the reference's cuts
through a hole row (`6x` markers, two depths).

If correct, `173.6/104.8/11.5` may be marked like the others and simply belong
to geometry the cut never reaches. Test: offset the cut to `Y = ±0.015` and
rerun with model dims.

Also confirmed wrong: our J-J is landscape 196 wide; the reference is portrait
173.6 tall. `SectionVertical` defaults `False`.

### Dead code found

`Module5_FallbackDimensionEngine.InsertHoleCalloutsForView` has **no caller**
anywhere in the trunk. Hole callouts come only from native import's
`swInsertholeCallout`. Left in place pending a decision to wire or delete.

## Process: CURRENT_STATUS is gated at the end of every Claude Code turn

Claude Code now snapshots this file on `UserPromptSubmit` and checks it on
`Stop`. If its SHA-256 hash is unchanged, the Stop hook continues the turn
once with an evidence-honest update instruction. A second unchanged Stop is
allowed with a visible warning so the hook cannot create an infinite loop.

Read-only turns and turns with no material project change must add a dated
`No material status change` note rather than inventing progress. The hook does
not edit this file itself, rewrite prior evidence, or imply that static,
compile, execution, or visual-acceptance gates ran.

Verification: 5 focused hook tests pass; `.claude/settings.json` parses as
JSON. No VBA source, macro revision, deployment, compile, SOLIDWORKS execution,
or visual acceptance changed as part of this process update.

## r20 live: both ordinate chains match the reference EXACTLY

`MACRO_SOURCE_REVISION` is `trunk-2026-08-06-r20`, deployed 13/13, compiled
`verdict=Clean`, run against P-0251 with ordinate mode, Center datum, HLR.
Report `macro_qa/20260806_100240_P-0251-14A-001`.

| | trunk r20 | P-0251-14A-001 |
|---|---|---|
| Long axis | **160, 90, 50, 10, 0** | 160, 90, 50, 10, 0 |
| Cross axis | **36, 15, 0, 15, 36** | 36, 15, 0, 15, 36 |

```
readback: 8 dims, 0 dangling
Ordinate edges seen: 39 (circular: 22, of which arcs: 4, linear: 17)
Selection scope: ScopedToView(Let)
View scale: 0.6667 (stations are model mm, scale-normalised)
Outer-edge promotions: 2
Datum contract: LongAxis=X:FromMinEnd;ShortAxis=Y:FromCentreline
  X: 5 stations, datum Edge:offsetFromTarget=160.00mm
  Y: 5 stations, datum Hole:offsetFromTarget=0.00mm
```

The r19 1 mm systematic offset is closed by the user's drawing convention -
**a dimension always goes to the outer edge**. See
`SOLIDWORKS_API_VALIDATION.md`, r20 section.

**The ordinate engine is done against this fixture.** Gaps A1-A6, A8, A10
closed. What remains is everything that is not an ordinate.

### Open, in order

1. **Conventional dimensions do not exist.** Everything on the reference
   except the two ordinate chains and the hole callouts - `R36`, `Ø40`,
   `Ø47 H7`, `18`, `12`, `11.5`, `173.6`, `104.8`, `80`, `25`, `6` - is a
   conventional dimension that appears only if the model happens to have it
   marked. Nothing in the trunk creates one. This is the Tier C gap; see
   section 7 of the gap doc. **Largest remaining work package.**
2. **Hole callouts.** The reference carries `6x Ø6.6 THRU / ⌴Ø11 ↧6` and
   `4x Ø4.2 ↧12.4 / M5x0.8-6H ↧10` as consolidated leader callouts. Not
   verified against the trunk since the `DuplicateDims` fix.
3. **Section view placement and sheet layout** (A7, C-series). The section
   view has run off the sheet frame in several runs; view placement is
   arithmetic on the part bounding box with no bounds check.
4. **`swCreateOrdDimErr_OrdFailure`** on non-front views. Four failures at r7,
   never attempted since. **Undiagnosed.**
5. **The form owns the settings, and it is unmanaged.** `ResetGlobalConfig` is
   a no-form fallback; `UserForm1` seeds from saved registry settings and
   overwrites `GlobalConfig` on OK. The forms are outside the deployment
   manifest, cannot be imported, and no test covers them.
6. `A9` - `COORD_DEDUP_TOL_M` is still an unjustified 1.5 mm, though no longer
   scale-dependent.
7. `swDefaultTemplateDrawing = 10` inherited, never re-queried.
   `IView.SetDisplayMode3` is obsolete and still called.

### Waiting on a user decision - do not default

- **Mass units.** Title block reads `MASS(KG): 1296.82` against the
  reference's `1.30 kg`.
- **Duplicate general notes.** Once from the template's sheet format, once
  from `InsertNotes`. `DrawingContainsText` cannot see sheet-format text.

### Process gates added 2026-08-06

Instruction text in `CLAUDE.md` did not hold across a long session: three
`swDisplayMode_e` constants sat in the trunk with values from other enum
members and survived fifteen live runs. Two mechanical gates now enforce the
API-lookup contract - a `PreToolUse` hook that blocks `sw*` edits without a
recent MCP lookup, and a test that fails the suite when a `sw*` constant has
no provenance record. See `CLAUDE.md` and `Architecture.md`.

## r19 live: both ordinate chains match the reference structurally

`MACRO_SOURCE_REVISION` is `trunk-2026-08-06-r19`, deployed, compiled
`verdict=Clean`, run against P-0251 with ordinate mode, Center datum and HLR.

| | trunk r19 | P-0251-14A-001 |
|---|---|---|
| Long axis | 159, 89, 49, 9, **0** | 160, 90, 50, 10, **0** |
| Cross axis | 36, 15, **0**, 15, 35 | 36, 15, **0**, 15, 36 |

Five stations per axis, zero dangling ordinates, correct 40 mm pitch, datums
on a straight end face and on the centreline respectively.

```
readback: 8 dims, 0 dangling
Ordinate edges seen: 39 (circular: 22, of which arcs: 4, linear: 17)
Selection scope: ScopedToView(Let)
View scale: 0.6667 (stations below are model mm, scale-normalised)
Datum contract: LongAxis=X:FromMinEnd;ShortAxis=Y:FromCentreline
  X: 5 stations (holes=4, edges=1), datum Edge:offsetFromTarget=159.00mm
  Y: 5 stations (holes=3, edges=2), datum Hole:offsetFromTarget=0.50mm
```

### Closed since r7

Gaps **A1-A6, A8, A10** in `BASELINE_TO_REFERENCE_DRAWING_GAP.md`.

| Rev | What changed |
|---|---|
| r8 | `IsIsoView` replaced by `Module8_ViewClassifier`; ordinates confined to the front view |
| r10 | Per-axis candidate sets and datums; straight edges admitted; 1-D dedup |
| r11 | One entity belongs to one chain; alternates retained per station |
| r13 | Coordinates normalised by `IView.ScaleDecimal` |
| r15 | Dangling prune moved after `ForceRebuild3` |
| r16 | `ISelectData.View` scoped again; three display-mode constants corrected |
| r17 | HLR removes the hidden-line edges that could not hold an ordinate |
| r19 | Arcs kept as stations; an **end** datum must be a straight edge |

### Open, in order

1. **The 1 mm systematic offset.** Every long-axis station is exactly 1 mm
   below its reference counterpart; the cross axis matches on one side (36)
   and is 1 mm short on the other (35). Candidate: the drawing's own note
   "All corners are chamfered 0.5 x 45 deg" would put the selected straight
   edge on the chamfer rather than the true face extreme. **Not tested.**
2. **`swCreateOrdDimErr_OrdFailure`** on non-front views. Four failures at r7;
   never attempted since, because ordinates are now front-view only.
   **Undiagnosed.**
3. **The form owns the settings, and it is unmanaged.** `ResetGlobalConfig` is
   a no-form fallback; `UserForm1` seeds from saved registry settings and
   overwrites `GlobalConfig` on OK. The forms are outside the deployment
   manifest and cannot be imported, so operator-visible defaults can only be
   changed by hand-editing in the VBE. No test covers the forms.
4. **Section view placement and sheet layout** (C-series gaps). The section
   view has run off the sheet frame in several runs; view placement is
   arithmetic on the part bounding box with no bounds check (A7).
5. **Conventional dimensions do not exist.** Everything on the reference
   except the two ordinate chains and the hole callouts is a conventional
   dimension that appears only if the model happens to have it marked. This is
   the Tier C gap; see section 7 of the gap doc.
6. `swDefaultTemplateDrawing = 10` is an inherited value, never re-queried.
7. `IView.SetDisplayMode3` is obsolete; the trunk still calls it.

### Waiting on a user decision - do not default

Both confirmed on the sheet; the user has deferred them.

- **Mass units.** Title block reads `MASS(KG): 1296.82` against the
  reference's `1.30 kg`.
- **Duplicate general notes.** Once from the template's sheet format, once
  from `InsertNotes`. `DrawingContainsText` cannot see sheet-format text, so
  QA cannot detect it.

### Process gates added 2026-08-06

Instruction text in `CLAUDE.md` did not hold across a long session: three
`swDisplayMode_e` constants sat in the trunk with values from other enum
members and survived fifteen live runs. Two mechanical gates now enforce the
API-lookup contract - a `PreToolUse` hook that blocks `sw*` edits without a
recent MCP lookup, and a test that fails the suite when a `sw*` constant has
no provenance record. See `CLAUDE.md` and `Architecture.md`.

## r15 live: no wrong dimensions on the sheet; two stations lost

`MACRO_SOURCE_REVISION` is `trunk-2026-08-06-r15`. Report at
`macro_qa/20260806_064*_P-0251-14A-001/QA_REPORT.txt`.

```
readback: 12 dims, 0 dangling
  live values: 43.00, 45.00, 70.00, 110.00, 150.00, 159.00,
               15.00, 23.60, 36.00, 15.00, 23.60, 35.00
View scale: 0.6667 (stations below are model mm, scale-normalised)
Datum contract: LongAxis=X:FromMinEnd;ShortAxis=Y:FromCentreline
  X: 9 stations, datum Hole:offsetFromTarget=0.00mm
  Y: 7 stations, datum Hole:offsetFromTarget=0.50mm
Dangling ordinates: 2 found, 2 deleted (0 select-refused, 0 delete-refused)
```

Every dimension on the sheet is now correct. Gaps A2, A3, A4 and A8 are
closed, and the `Center` datum contract is proven live: the Y chain is
symmetric about the centreline - `36.00, 23.60, 15.00, 0, 15.00, 23.60, 35.00`
against the reference's `36, 15, 0, 15, 36`.

### Fixed since r10

- **View scale.** `ModelToViewTransform` carries the view scale, so
  `COORD_DEDUP_TOL_M` was being applied in view space and the macro emitted
  **different dimensions at different scales** - at 1:2 the 43 mm and 45 mm
  stations merged. Now normalised by `IView.ScaleDecimal`.
- **Shared entities.** An entity belonged to both chains; the second use
  dangled. Each station keeps an alternate. Fixed the Y=36.00 bore station.
- **Dangling prune.** `IAnnotation.IsDangling` is a post-rebuild property, so
  the prune had to move downstream of `ForceRebuild3`.

### Open, in order

1. **Root cause of the two dangling stations** (X=55.00, X=135.00). They are
   deleted, not explained. Not a sharing problem - the X chain runs first, so
   every substitution was necessarily in the Y chain.
2. **Too many stations.** X has 9 against the reference's 5, Y has 7 against
   5. Admitting every axis-parallel edge over-corrected A4; stations need a
   selection rule, not a wider net. The extra Y pair is +/-23.60.
3. **The X datum is not the part extreme.** X=0 sits on the bore centre while
   the part extends further left; a rounded end is an arc and arcs are not
   admitted as stations.
4. **The Y extremes are asymmetric** (-36.00 / +35.00), so the centreline
   datum sits 0.50 mm off. One silhouette station is 1 mm short of the true
   outline.
5. Diagnose `swCreateOrdDimErr_OrdFailure` - the 4 failures at r7 were on
   non-front views and have not been attempted since. **Undiagnosed.**
6. `ISelectData.View` error 91 - worked around, not solved.
7. Section view placement, view layout, and the C-series gaps.

### Waiting on a user decision - do not default

Both confirmed on the sheet; the user has deferred them.

- **Mass units.** Title block reads `MASS(KG): 1296.82` against the
  reference's `1.30 kg`.
- **Duplicate general notes.** Once in the template's sheet format, once from
  `InsertNotes`. `DrawingContainsText` cannot see sheet-format text.

## r10 live: per-axis datum contract works; three stations dangle

`MACRO_SOURCE_REVISION` is `trunk-2026-08-06-r10`, deployed, compiled
`verdict=Clean`. Report at
`macro_qa/20260806_054312_P-0251-14A-001/QA_REPORT.txt`.

```
Total drawing view dimensions: 14
Ordinate edges seen: 64 (circular: 35, linear: 29)
Datum contract: CornerBottomLeft
  X: 9 stations (holes=6, edges=3), datum Hole:offsetFromTarget=0.00mm
    offsets(mm): 0.00, 43.00, 45.00, 55.00, 70.00, 110.00, 135.00, 150.00, 159.00
  Y: 7 stations (holes=5, edges=2), datum Edge:offsetFromTarget=0.00mm
    offsets(mm): 0.00, 12.40, 21.00, 36.00, 51.00, 59.60, 71.00
Ordinate chains created: 2 of 2 attempted
```

Gaps **A2, A4 and A8 are closed**: independent candidate sets and datums per
axis, straight model edges admitted as stations, and per-axis 1-D
deduplication. **A3 is closed in code** - the datum is no longer required to be
a circle, and the Y datum resolved to an edge.

A linear edge is confirmed valid as an ordinate base entity on this build
(zero holes-only retries). See `SOLIDWORKS_API_VALIDATION.md`.

### The r8 "three zeros" defect is located

The sheet and the report agree on station *positions* and disagree on three
*values*. X=55.00 and X=135.00 render `0.00`, as does Y=36.00, all three in the
dangling-dimension colour, while the report lists their true offsets. Creation
is placing them correctly; the **attachment** is failing for a subset of
entities. That is the next thing to fix, and it is now a narrow question:
which candidate entities produce an attachment that does not survive.

### Open, in order

1. **Three dangling stations** (above). Identify the entity class that fails to
   hold an attachment and reject it at collection time.
2. **The X datum is not the part extreme.** X=0 sits on the bore centre while
   the part visibly extends ~56 mm further left, so the left-hand outline is
   not in the candidate set - a rounded end is an arc, not a line, and arcs are
   not yet admitted as stations.
3. **Too many stations.** X has 9 against the reference's 5, Y has 7 against 5.
   Admitting every axis-parallel edge over-corrected A4; stations now need a
   selection rule, not just a wider net.
4. `GetDisplayDimensions` returned 14 while 16 ordinate labels are on the
   sheet. Unexplained, minor.
5. Diagnose `swCreateOrdDimErr_OrdFailure` - the 4 failures at r7 were on
   non-front views and have not been attempted since. **Undiagnosed.**
6. `ISelectData.View` error 91 - worked around, not solved.

### Untested

The run used `CornerBottomLeft`: the operator selected Bottom-Left on the
form. The default `Center` contract - long axis measured from its minimum end,
short axis from the centreline, which is what the reference drawing does - did
**not** execute and remains unproven live.

## r8 live: ordinates land on the front view only; values still wrong

`MACRO_SOURCE_REVISION` is `trunk-2026-08-06-r8`, deployed (13/13 managed
components), compiled `verdict=Clean`, run against P-0251 with ordinate mode
selected. Report at
`macro_qa/20260806_051241_P-0251-14A-001/QA_REPORT.txt`.

```
Total drawing view dimensions: 9
View roster:
  Drawing View1 | Type=7 | Orientation=*Front | Role=Front | ordinates=allowed | 9 dims
  Drawing View2 | Type=7 | Orientation=*Bottom | Role=Bottom | ordinates=skipped | 0 dims
  Drawing View3 | Type=7 | Orientation=*Right | Role=Right | ordinates=skipped | 0 dims
  Drawing View4 | Type=7 | Orientation=*Left | Role=Left | ordinates=skipped | 0 dims
  Drawing View5 | Type=7 | Orientation=*Isometric | Role=Pictorial | ordinates=skipped | 0 dims
  Section View J-J | Type=2 | Orientation=(empty) | Role=Section | ordinates=skipped | 0 dims
Ordinate views processed: 1
Ordinate chains created: 2 of 2 attempted
Selection scope: Unscoped(err=91)
PASS: Drawing contains dimensions.
```

Gaps **A5 and A6 are closed**. `IsIsoView` is deleted;
`Module8_ViewClassifier` classifies from `IView.Type` +
`IView.GetOrientationName`, and the §6 caveat in the gap doc is resolved with
live evidence (all model views `Type=7`; orientation strings round-trip
exactly). The roster is printed on every run, so the classification carries
its own evidence.

### Open defects, observed on the r8 sheet

1. **The horizontal chain has three zeros.** It reads
   `0, 0.00, 70.00, 110.00, 0.00, 150.00`; a chain has one datum.
2. **`Section View J-J` is placed off the right edge of the sheet frame**, and
   the `*Right` / `*Left` views are on the sheet per the roster but not
   visible inside the frame.
3. **Chain values do not match the reference.** The vertical chain reads
   `23.60 / 15.00 / 15.00 / 23.60` against the reference's
   `36 / 15 / 0 / 15 / 36`. The `15`s match; `23.60` against `36` is
   consistent with gap A4 (circular-edge-only candidates terminate the chain
   on the outermost hole instead of the silhouette edge). Consistent with,
   not proof of.

### Next, in order

1. **Per-axis datum contract** (gaps A2-A4) - X on the centreline, Y on the
   bottom edge, neither of them a hole, and silhouette edges as candidates.
   This is what defects 1 and 3 above both sit on.
2. View placement / section-view position (defect 2, C-series gaps).
3. Diagnose `swCreateOrdDimErr_OrdFailure`. The 4 failures seen at r7 were on
   non-front views; r8 never attempted them, so this is **undiagnosed, not
   fixed**.
4. `ISelectData.View` error 91 - worked around, not solved.

### Waiting on a user decision - do not default

Both are now visually confirmed on the r8 sheet. Recorded so they are not
lost; the user has deferred them.

- **Mass units.** Title block reads `MASS(KG): 1296.82` against the
  reference's `1.30 kg`. One line, once it is confirmed the source property is
  always grams.
- **Duplicate general notes.** The notes appear twice - once in the template's
  sheet-format box, once from `InsertNotes`. `DrawingContainsText` walks
  per-view notes only and cannot see sheet-format text, so QA cannot detect
  this.

## Historical: trunk moved to baseline; Phase 1 landed, unrun

`MACRO_SOURCE_REVISION` is `trunk-2026-08-05-r2` in
`src/baseline-model-dims/Module1_Main.bas`. **Not deployed. Not compiled in
SOLIDWORKS. Not run.** Everything below is static and API-contract evidence.

r2 corrects `TEMPLATE_PATH` to `V:\VEEMAP\SW_data\...` (user-confirmed). The
baseline's value was missing the `VEEMAP` segment and failed silently, so it
had been building on the SOLIDWORKS default drawing template.

`target-spec-hybrid-v2` is archived. Everything in the r62 section further
down this document describes the archived implementation and is no longer the
state of the deployed macro.

### What is verified

| Gate | Result |
|---|---|
| Companion live suite | 34/34 pass |
| Managed-source hygiene (ASCII, no BOM, no `Attribute`, `Option Explicit`) | pass |
| Manifest components present on disk | pass, 12/12 |
| MCP contract lookups | `AddOrdinateDimension`, `InsertModelAnnotations4`, `swInsertAnnotation_e`, `swAddOrdinateDims_e`, `swCreateOrdDimError_e`, `swImportModelItemsSource_e`, `swDrawingViewTypes_e` |
| VBA compile in SOLIDWORKS | **not run** |
| Deployment readback | **not run** |
| Live macro run | **not run** |
| Visual acceptance | **not run** |

### Next action

Deploy and compile. Phase 1 changed cross-module signatures
(`CreateHoleOrdinateDims`, `BuildRunSummary`, `AddFallbackOrdinateDimensions`
all gained a parameter; `CreateOneOrdinateChain` became a `Function`), so a
compile is the first real gate.

```bash
powershell -ExecutionPolicy Bypass -File ".\tools\production-runner\Run-R23Production.ps1" -AllowMutation -Deploy
```

Then judge whether `SetPickMode` fixed the ordinate breakage and whether
`DuplicateDims=True` stopped the repeated callouts. Those are the two
hypotheses Phase 1 exists to test.

### Known inert

The probe runner's probe stage does nothing — the nine `R23_Probe*` entry
points lived in the archived Module10–19. Deploy and compile stages work.

## Historical: r62 live - ANNOTATION_EXTENTS proved; three placement defects open

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r62`, deployed.
Companion suite **619/619**. Run
`macro_qa/20260805_071309_P-0251-14A-001`, deploy verified at r62,
pre-flight `ready=True|verdict=Clean`.

**Required failing stages: 5** (was 6). `ANNOTATION_EXTENTS` is PROVED.

```
SECTION_ANNOTATION_CLAMP_PASS|view=Section View J-J|annotations=6|moved=3|stillOutside=0
RD1            fromY=0.278500  requestY=0.275000  readbackY=0.275000  nowInside=True
RD2            fromY=0.289235  requestY=0.275000  readbackY=0.275000  nowInside=True
DetailItem927  fromY=0.036285  requestY=0.088749  readbackY=0.088749  nowInside=True
RD3 / RD4 / RD5                action=AlreadyInside
```

RD3/RD4/RD5 were never moved, so the documented `SetPosition2` refusal for
radial and diametric dimensions is **still untested live**.

### Three placement defects for r63, all introduced by this work

The stage passing did not make the sheet right; the user's screenshot showed
what the gate cannot measure.

1. **RD1 and RD2 occupy the same point.** `ANNOTATION_GEOMETRY` gives
   `x=0.206692|y=0.275000` for both - 12.00 is hidden under 18.00. The clamp
   drives every violator to the identical boundary point.
2. **RD3 and RD5 occupy the same point**, `x=0.179752|y=0.160060`. Predates
   r62. `LaneTextPoint` handles `LANE_BORE_SIDE_A` and
   `LANE_EXTERIOR_VERTICAL_INNER` in one `Case`, so the r61 per-lane ordinal
   made them identical; the r60 global counter had separated them by
   accident.
3. **The SECTION J-J label moved 52 mm onto the view for no reason.**
   `DetailItem927` went `y=0.036285` to `0.088749` while never violating:
   `CONTROLLED_REGIONS|content=0.010000,0.010000,0.410000,0.287000`. The
   clamp bounds to the **view-usable** box, whose bottom is
   `TitleBlockTop + margin` - a view constraint. `ANNOTATION_EXTENTS`
   enforces the content border and the title-block rectangle, and the label
   was inside both.

**r63 fix direction.** Clamp against the regions the stage actually enforces
(content border inset by `LAYOUT_MARGIN_M`, plus title-block avoidance);
step a second annotation along the free axis rather than stacking it; give
`LANE_EXTERIOR_VERTICAL_INNER` its own offset from `LANE_BORE_SIDE_A`.

## r62 source - the annotation clamp moves after the layout

Eleven mutations against the new guards each caught and reverted.

`Module10.ClampSectionAnnotationsIntoUsableArea` runs from Module2 after the
structural grid, auto-arrange and the title block, and before the final
rebuild. It pulls the section view's annotation origins back inside the
proved usable box.

- Fourth and last mutating procedure in Module10. Refuses without
  `allowMutation`; refuses again without `LayoutBoundariesProven`.
- Annotation origins only. No view move, no rescale - the retired
  content-envelope paths stay uncalled. Scoped to the one section view on
  the P-0251 fixture.
- `IAnnotation::SetPosition2` shares `GetPosition`'s frame (sheet
  coordinates from the lower-left corner), so the two compare directly. The
  Help warns that a constrained annotation is placed "as near as possible"
  and that radial/diametric dimensions cannot be positioned this way at
  all; both look like a call that returned, so `nowInside` is decided by
  the readback and `stillOutside` is counted separately.
- Already-inside annotations are recorded and left alone. The z coordinate
  is carried through, not zeroed.
- The creation-time clamp remains as a starting bound only.

Whether `ANNOTATION_EXTENTS` returns to PROVED is a question for the next
live run.

## r61 live - verdict honest, clamp box stale

`SECTION_DIMENSIONS` unchanged at `satisfied:5`. Run
`macro_qa/20260805_070039_P-0251-14A-001`, deploy verified at r61,
pre-flight `ready=True|verdict=Clean`.

**Fixed and proved:** `R23_SECTION_DIMENSIONS|requirements=7|satisfied=5|`
`missing=2` now agrees with the QA stage. Per-lane ordinal works -
`BoreSideA`, `BoreSideB`, `ExteriorVerticalInner` all `laneOrdinal=1`.

**Not fixed, and the r60 diagnosis was wrong.** `ANNOTATION_EXTENTS` still
failed on RD2 at `y=0.289235`. Auto-arrange never touched this view
(`ACTIVATE_VIEW|operation=Dimension arrange` names Drawing View1, Drawing
View2, Drawing View4 only). `ArrangeViewsInMeasuredGrid` moved the section
view after its dimensions were created:

```
creation-time outline  0.289060,0.039385,0.318940,0.252265
LAYOUT_MOVE readback   0.191752,0.053620,0.221632,0.266500
delta                 -0.097308,          +0.014235
```

RD2 requested `(0.304000, 0.275000)` + delta = `(0.206692, 0.289235)`, the
violation line to six decimals. RD1 lands at `0.278500`, under
`ContentBorderTop` 0.287, so only RD2 trips. A box measured before the view
is placed is stale by construction - hence r62.

## r61 source - both r60 defects addressed

Companion suite was **606/606**. Superseded on the placement defect by r62;
the verdict fix stands as written.

1. **Placement is per lane and bounded.** `NextLaneOrdinal` counts each lane
   separately, and `LaneTextPoint` clamps the result into the proved
   `evidence.Usable*` box before reporting it. `SECTION_DIM_PLACEMENT` gains
   `laneOrdinal`, `appliedOrdinal`, `usable=` and `clamped=`. Without
   `LayoutBoundariesProven` the ordinal stops at 1 and the proof says
   `usable=Unproved`.
2. **Module2 rebuilds and re-reconciles the requirements after creation**
   rather than re-verifying the objects creation mutated. Re-reading the
   dimension inventory alone was not enough - `Failures` carried
   `NoImportedDimension` from the pre-creation reconcile, and
   `VerifySectionDimensions` counts a non-empty `Failures` as unsatisfied.

Verification: eight mutations covering the new guards (each clamp side, the
per-lane ordinal, the unproved-bounds cap, the post-creation reconcile, and
the verdict's requirement source) were applied one at a time and each failed
the suite before being reverted. Whether `ANNOTATION_EXTENTS` returns to
PROVED is a question for the next live run.

## r59/r60 - five section dimensions created and correct; two defects of mine

Superseded by r61 for the two defects below; the run results stand.
Companion suite **602/602** at the time. Runs
`macro_qa/20260805_062318` (r59, classifier) and
`macro_qa/20260805_063139` (r60, creation).

**r59 confirmed the arc-chord defect.** Restricting axis-parallel
classification to Type 0 records dropped `vertical` 333 to 19, `distinctX`
10 to 7, `arcTessellation=314`, and `LOWER_WALL_STEP_11_5` lost its pair -
exactly as predicted. Five requirements survived with selectable geometry.

**r60 created all five, with correct nominals**, verified by
`CreateSectionDimension`'s own nominal readback:

| key | name | nominal read back |
| --- | --- | --- |
| `OVERALL_THICKNESS_18` | RD1 | 0.018000 |
| `BORE_STEP_DEPTH_12` | RD2 | 0.012000 |
| `INNER_BORE_D40` | RD3 | 0.040000 |
| `FIT_BORE_D47_H7` | RD4 | 0.047000 |
| `LOWER_VERTICAL_REF_104_8` | RD5 | 0.104800 |

`SECTION_DIMENSIONS` moved from `satisfied:0/missing:7` to
**`satisfied:5/missing:2`**; the two missing are the ones with no geometry
in the view. `SECTION_DIM_PASS|created=5`.

### Two defects, both mine

1. **`ANNOTATION_EXTENTS` regressed PROVED to FAILED**, on one line:
   `Annotation origin violation in 'Section View J-J': type=DisplayDimension,
   name=RD2, x=0.206692, y=0.290500, region=ZonedBorder`. `LaneTextPoint`
   stacks a *global* ordinal, so gaps grow 12/24/36/48/60 mm and the second
   lane entry lands above the usable area; auto-arrange then moved it into
   the zoned border. Fix: per-lane ordinal, and clamp to the proved
   `evidence.UsableTop/Bottom/Left/Right` rather than stacking unbounded.
2. **The Module2 verdict line is wrong**: `R23_SECTION_DIMENSIONS|satisfied=0`
   with `RequirementFlagged:...:NoImportedDimension`, while the QA stage
   independently reports `satisfied:5`. Reconciliation appends
   `NoImportedDimension` to `Failures` before creation, and creation sets
   `Matched` without clearing it. The stage verdict is the correct one -
   Module19 re-reconciles fresh requirements against the finished drawing.
   Fix: re-reconcile after creation in Module2 too, rather than re-verifying
   stale requirement state.

Six required stages now fail: the five from before plus
`ANNOTATION_EXTENTS`.

## r58 - every candidate is selectable; one mapping defect found

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r58`, deployed.
Companion suite **587/587**. Run
`macro_qa/20260805_061804_P-0251-14A-001`. Five required stages still fail.

**All ten candidate entities selected and proved their owner.** Every
`SECTION_ENTITY_SELECT` row reads
`selectable=True|route=Selectable|owner=Section View J-J`. No
`EntityIsNothing`, so the entity array is live for these records in a cut
view - the `Null` case the Help warns about did not arise here. **Dimension
creation is reachable.**

| requirement | recordA | recordB |
| --- | --- | --- |
| `OVERALL_THICKNESS_18` | 2 | 22 |
| `BORE_STEP_DEPTH_12` | 2 | 5 |
| `LOWER_WALL_STEP_11_5` | 5 | **25** |
| `INNER_BORE_D40` | 24 | 24 |
| `FIT_BORE_D47_H7` | **25** | 25 |
| `LOWER_VERTICAL_REF_104_8` | 10 | 44 |

### The defect: record 25 is an arc and a "wall edge" at once

Record 25 supplies the Ø47 radius (0.023500) **and** the `x=0.008500` side
of `LOWER_WALL_STEP_11_5`. Both cannot describe the same thing.

The cause is in `InventorySectionGeometry`: an arc record carries a
tessellated point array as well as its GeomData, and the axis-parallel
classifier walks those points. A chord of the bore's tessellation that
happens to run parallel to an axis is being recorded as a straight edge
coordinate. Attaching a linear dimension to it would dimension the bore
arc, not the wall.

**Fix**: classify axis-parallel segments from Type 0 (polyline) records
only. An arc's tessellated chords are an approximation of a curve, not
straight edges. Expected consequence: `LOWER_WALL_STEP_11_5` may lose its
pair, which would be the correct answer rather than a regression.

Record 2 is shared by `OVERALL_THICKNESS_18` and `BORE_STEP_DEPTH_12`;
that is plausible for one long edge bounding both and is not evidence of a
fault.

Geometry is not the blocker; selection is. `CreateSectionDimension` refuses
with `reason=NoEntitiesSelected`, and nothing connected a measured
coordinate to a selectable drawing entity.

**The connection was already in the data.** `GetPolylines7`'s return value
IS the entity array, positionally paired with the polyline records - r57
confirmed that exactly (`records=79|entities=79|recordsMatchEntities=True`).
So each distinct coordinate now remembers the polyline record that produced
it, and that index reaches the entity.

- `SECTION_REQ_CANDIDATE` gains `|recordA=|recordB=`.
- New `SECTION_ENTITY_SELECT|key=|side=|record=|selectable=|route=` per
  candidate, from `ProveSectionEntitySelection` - the route Module13 proved
  for orthographic views: `IView.SelectEntity`, then
  `ISelectionMgr.GetSelectedObjectsDrawingView2` to prove the owning view
  before the object is trusted. It refuses a pre-existing selection and
  clears on every path including the error handler.
- Attempted only for a requirement whose geometry was `found` under a
  trusted decode, so at most twelve selections.

**Nothing is created.** The open question is whether those entity entries
are live in a CUT view: many section curves are cut faces with no model
edge behind them, and the Help says the array carries `Null` in their
place. `route=EntityIsNothing` is the expected measurement for those, not
an error. Creating a dimension before that is answered would be guessing
again.

## r57 - clean decode; six of seven requirements have proved geometry

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r57`, deployed.
Companion suite **573/573**. Run
`macro_qa/20260805_055822_P-0251-14A-001`. Five required stages still fail:
`NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`, `SECTION_DIMENSIONS`,
`MANUFACTURING_DEFINITION`, `FINAL_QA`.

```text
SECTION_GEOM_SUMMARY|decodeStatus=Complete|records=79|entities=79
  |recordsMatchEntities=True|doubles=2799|consumed=2289|trailing=510
  |trailingAllZero=True|arcs=18|polylines=61|points=454|error=0
```

First clean decode since the inventory existed. `consumed=2289` matches
`9*79 + 12*18 + 3*454` exactly, records match entities, and the padding is
**measured** to be zeros rather than assumed.

`geometryTrusted=True` on every requirement row for the first time:

| requirement | nominal | evidence in Section View J-J |
| --- | --- | --- |
| `OVERALL_THICKNESS_18` | 0.018 | X pair -0.009 / 0.009 |
| `BORE_STEP_DEPTH_12` | 0.012 | X pair 0.009 / -0.003 |
| `LOWER_WALL_STEP_11_5` | 0.0115 | X pair -0.003 / 0.0085 |
| `INNER_BORE_D40` | 0.040 | radius 0.020 |
| `FIT_BORE_D47_H7` | 0.047 | radius 0.0235 |
| `LOWER_VERTICAL_REF_104_8` | 0.1048 | Y pair 0.041 / -0.0638 |
| `LONG_VERTICAL_REF_173_6` | 0.1736 | **absent** |

`LONG_VERTICAL_REF_173_6` has no candidate: the section's extreme Y values
are 0.1005 and -0.1005, a span of 0.201, and no pair of the nineteen
distinct Y values differs by 0.1736. Open question, not a decode failure.

### The blocking work

`SECTION_DIMENSIONS` is `requirements:7/satisfied:0` with
`sectionDimensions:0`. **Nothing creates them.**
`ReconcileR23SectionDimensions` only reconciles;
`Module10_SectionDimensionEngine.CreateSectionDimension` exists, refuses
without explicit authorization and without a selection
(`reason=NoEntitiesSelected`), and is never called from the production
route. Its inputs are proved for the first time.

The open problem is selection: the inventory reads coordinates through
`GetPolylines7`, which returns view-frame geometry, while
`CreateSectionDimension` needs selected drawing entities. Nothing yet maps
a measured coordinate back to a selectable entity in the section view.

## r57 (source) - the array is zero-padded; the r56 reversal was my error

**Every double from index 2289 to the end of the polyline array reads
`0.000000000`.** The unbounded walk parsed 510 zeros as 56 phantom records
of stride 9 (type 0, GeomDataSize 0, six zero style fields, NumPolyPoints
0), reaching `records=135` with 6 doubles left over that cannot complete a
record. **That is the whole of `Desynchronized:StyleAt2801`**, open since
r53.

So 79 real records consume exactly 2289 doubles (`9*79 + 12*18 + 3*454`),
`entities=79` is the true record count, and the Help's positional pairing
holds. The rest is padding.

**The r54 bound was right and I removed it on a bad inference.** The r55
run stopped at 79 records with 510 doubles left and I read that as proof
more records existed, without checking what those doubles contained - the
r55 window had already printed zeros at 2289 and I did not look. The
arithmetic I cited was correct; the conclusion was not.

r57 restores the bound and **verifies the padding rather than assuming
it**: `trailingAllZero=` is computed and a non-zero tail is reported as
`TrailingDataAfterEntities`. Assuming is what produced the wrong reversal.
`consumed=` is also reported, so the walk's arithmetic is in the log
instead of being reconstructed afterwards.

Expected next run: `decodeStatus=Complete` for the first time, making the
section coordinates trustworthy - including the radii `0.020000` and
`0.023500` that have carried `geometryTrusted=False` since r53.

## r56 - the decode was never misaligned; the control was wrong. NOT DEPLOYED

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r56`.
Companion suite **572/572**. **No r56 run has happened.** The last live
evidence is r55 at `macro_qa/20260805_054524_P-0251-14A-001`.

r55 measured this, and it disproves the r54 control:

```text
SECTION_GEOM_DESYNC|status=StoppedAtEntityCount|stoppedAt=2289
  |recordsDecoded=79|lastRecordStarts=2274,2259,2244
SECTION_GEOM_SUMMARY|records=79|entities=79|doubles=2799|trailing=510
  |arcs=18|polylines=61|points=454
```

`9*79 + 12*18 + 3*454 = 2289`, exactly the consumption. A misaligned walk
does not produce a consumption matching its own record, arc and point
counts to the double. The window confirms it independently: the record at
2244 reads type 0, GeomDataSize 0, six style scalars, two points - stride
15 - and the next record starts at 2259, exactly as reported.

**Alignment was never lost.** A section view holds more polyline records
than entity-array entries, which is what cut edges with no model edge
behind them would produce. The Help's positional-pairing statement does not
extend to a cut view.

So the entity count was the wrong control and truncated a correct decode.
It is now context (`recordsExceedEntities=`) and gates nothing; the control
is exact consumption, `trailing=0`, with every advanced field still
range-checked.

Open: r53's `Desynchronized:StyleAt2801` is two past the end of the array
and is not explained by the above. The unbounded walk plus window will
settle it on the next run.

### r55 failed to compile first time

`lower` was used in `InventorySectionGeometry` without being declared -
copied from `EmitSectionLineDecode`, which declares it. `Option Explicit`
rejected it, `R23_PREFLIGHT_END|ready=False|reason=CompileNotClean`, `main`
never ran. The compile gate is the right guard for this class and it cost
one cycle; no static test duplicates it.

## r54 - the decoder stops lying and says where it broke; NOT DEPLOYED

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r54`.
Companion suite **573/573**. **No r54 deployment or run has happened.** The
last live evidence is r53 at `macro_qa/20260805_051116_P-0251-14A-001`.

Read-only. Four changes, all to the instrument rather than the drawing:

1. **Bounded by the entity count.** The polyline data and the entity array
   are positionally paired, so r53's `records=135` against `entities=79`
   was impossible by construction. The walk now stops and reports
   `StoppedAtEntityCount`.
2. **A clean decode must consume the array exactly.** `trailing=` is
   reported and `Complete` is downgraded to `RecordCountMismatch` or
   `TrailingDoubles` when the counts disagree.
3. **A failed decode prints raw values around the failure** -
   `SECTION_GEOM_DESYNC|stoppedAt=|lastRecordStarts=` plus two
   `SECTION_GEOM_WINDOW` dumps of 36 doubles. 2799 doubles cannot be dumped
   whole; a window is what identifies the deviation.
4. **Untrusted coordinates are labelled at every consumer.**
   `SECTION_GEOM_X/Y/R` carry `|decodeStatus=` and every
   `SECTION_REQ_CANDIDATE` carries `|geometryTrusted=`. In r53 only the
   summary said the walk had failed, so six `found=True` rows read as
   findings when they were not.

**The layout was not changed and is not in doubt.** Against the r50 and r52
arrays it closes to the double: 38 records x 9 fixed fields + 6 arcs x 12
GeomData + 3 x 214 points = 1056, exactly the array length. Something in
the richer r53 section deviates from it, and the next run's window says
what.

## r53 - the bore is in the section; two stages regressed

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r53`, deployed.
Companion suite **569/569**. Run
`macro_qa/20260805_051116_P-0251-14A-001`. Failed required stages **6 to
5**; `ORDINATE_SCHEME` FAILED to PROVED.

Still failing: `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`,
`SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`, `FINAL_QA`.

### The offset option did what the enum said

`SECTION_CREATE_OPTIONS|options=2|offsetSection=True|segments=3`, and the
section view roughly doubled:

| | r52 | r53 |
| --- | --- | --- |
| entities | 38 | **79** |
| doubles | 1056 | **2799** |
| arcs | 6 | **18** |

Those come straight from the API return, not from any decode of ours. The
section is hatched and shows the bore step and the counterbore pockets.

### But the decoder tripped its own guard

```text
decodeStatus=Desynchronized:StyleAt2801|records=135|entities=79
  |recordsMatchEntities=False|doubles=2799
```

It walked off the record layout and stopped instead of inventing geometry -
the control working as designed. **Every coordinate list and every
`SECTION_REQ_CANDIDATE` verdict from this run comes from that walk and is
indicative only** until the desynchronization is diagnosed. What they
suggest: radii now include `0.020000` and `0.023500` (Ø40 and Ø47), which
have never appeared before, and six of seven requirements report
`found=True` against two at r50. Promising, unproved.

Diagnosing it needs the raw values around the failure index, which the
polyline decoder does not emit - 2799 doubles is too many to dump whole, so
it needs a window, not a dump.

### Two stages regressed, same cause

| | r52 | r53 |
| --- | --- | --- |
| `VIEW_PROJECTION` accepted | 8 | **6** |
| `VIEW_PROJECTION` without | 3 | **5** |
| `NATIVE_CALLOUT_COVERAGE` incomplete | 1 | **2** |
| `MANUFACTURING_DEFINITION` complete | 2 | **1** |

The M5 family lost the attachment it gained at r49. The two near-side M5
holes projected in the old projection-section view and do not project in
the offset one.

### The largest remaining piece

`SECTION_DIMENSIONS` is unchanged at `requirements:7/satisfied:0` with
`sectionDimensions:0`. **Nothing creates them.**
`ReconcileR23SectionDimensions` only reconciles, and
`CreateSectionDimension` - which exists, refuses without explicit
authorization, and requires a selection - is never called from the
production route. For the first time the geometry it needs appears to be
present.

One argument: `CreateSectionViewAt5` receives
`swCreateSectionView_OffsetSection` (corpus value 2, **verify in the SW2025
Object Browser**) instead of `0`, on the r52 evidence below.
`SECTION_CREATE_OPTIONS|options=|offsetSection=True|segments=` is emitted at
creation, because which option was used decides what the cut contains and
nothing else in the report would show it.

No other option bit was added - `Partial`, `DisplaySurfaceCut`,
`ChangeDirection` and `ScaleWithModel` each change what the section shows,
and a test asserts only the offset bit is set. The mutation boundary is
untouched.

**Unverified**: whether the offset bit alone suffices or
`swCreateSectionView_NotAligned` is also needed on this build.

## r52 - the line is exactly right; `Options=0` throws the jog away

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r52`, deployed.
Companion suite **562/562**. Run
`macro_qa/20260805_050411_P-0251-14A-001`. Stage table unchanged at six
failing, as expected - r52 changed no behaviour.

**The decode answered it.** The drawing holds precisely the path that was
asked for:

```text
index=1|start=-0.102000000,0.000000000|end=0.008000000,0.000000000
index=2|start=0.008000000,0.000000000|end=0.008000000,-0.015000000
index=3|start=0.008000000,-0.015000000|end=0.088000000,-0.015000000
```

Segment lengths 0.110, 0.015, 0.080 match the waypoint spacing exactly and
segment 1 runs 0.040 past the bore centre at -0.062, so the r51 overshoot
is in the drawing. **The path was never the defect.**

`Module17_SectionPath.CreateSectionFromPath` calls
`CreateSectionViewAt5(..., Options:=0, ...)`. MCP corpus,
`swCreateSectionViewAtOptions_e`: `swCreateSectionView_OffsetSection = 2` -
*"If set, then an aligned section view is created (two lines at an angle);
if not set, a normal projection section view is created."* SOLIDWORKS
therefore builds a normal projection section from a three-segment jogged
line and cuts at ONE offset. The bore sits at transverse 0.000 and the
counterbores at -0.015; the section holds the counterbore-column features
and no bore, so the cut in use is segment 3's.

That also explains r51 exactly: lengthening segment 1 cannot change a
section that is not cut along segment 1.

**The candidate change is one argument** - pass
`swCreateSectionView_OffsetSection`. Whether that flag alone is enough, or
`swCreateSectionView_NotAligned` is also needed, is unverified on SW2025.
It is a semantic change to how the section is cut, so it waits for the
user.

Read-only. `EmitSectionLineDecode` decodes `IView.GetSectionLineInfo2` for
the view carrying the cut and emits:

- `SECTION_LINE_RAW|view=|from=|values=` - the whole 49-double array, six
  per line. A dump cannot be wrong about the thing a structured decode
  might be wrong about.
- `SECTION_LINE_SEGMENT|view=|index=|lineType=|start=|end=|frame=AsReturned`
  - the segment endpoints SOLIDWORKS actually holds, to compare against the
  four waypoints the path handed it. The frame is deliberately not named;
  it has never been established, and claiming "Page" would be the
  mixed-frame defect this project has already paid for twice.
- `SECTION_LINE_DECODE|...|count=|documentedTotal=|tailMatchesDocumented=`
  - three segments account for 53 doubles under the documented layout while
  the live array holds 49, so the tail is known not to match and says so.

This exists because the count alone has never distinguished intent from
result: r51 moved waypoint 1 by 40 mm and produced a byte-identical section
view, with nothing in evidence contradicting the intent.

## r51 - the line moved, the section did not

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r51`, deployed.
Companion suite **552/552**. Run
`macro_qa/20260805_043613_P-0251-14A-001`. Stage table **unchanged**: six
required stages still fail.

**The prediction in this section was wrong.** The waypoint moved, the
predicate passed, and the section view did not change at all.

```text
SECTION_PATH_BORE_OVERSHOOT|centreY=0.210324890|projectedRadiusM=0.020000000
  |overshootM=0.040000000|w1Y=0.250324890|direction=AwayFromRows:PositiveY
SECTION_CROSSING|proven=4|columnHoles=3|failures=None
```

against a section view that is byte-for-byte what r50 produced: same
`records=38|entities=38|doubles=1056`, same radii
`0.002100;0.035000;0.036000`, same seven X and seven Y values, same outline
`0.191752,0.061120,0.221632,0.269000`. `INNER_BORE_D40` and
`FIT_BORE_D47_H7` still report `found=False`.

**What this rules out.** Not the predicate and not the waypoint - both now
do what they were built to do, and the line is visibly longer on the sheet.
A drawn section line's extent evidently does not determine what the cut
contains. That is consistent with a SOLIDWORKS section cutting through the
whole part regardless of the drawn length, but it is a hypothesis, not a
measurement, and it would also mean the bore was always being cut - which
the geometry says it is not.

**The missing measurement.** `IView.GetSectionLineInfo2` is read on every
run and only its element count is logged (`values=49`, r50 and r51 alike).
Its coordinates have never been decoded, so nothing in evidence says where
SOLIDWORKS actually placed the cut relative to the four waypoints handed to
it. Decode those 49 values before changing anything else.

### What r51 contains (both changes are correct and are staying)

Two defects, both exposed by the r50 geometry inventory.

1. **Waypoint 1 sat on the bore centre.** It now sits one full radius
   beyond the far wall - `BoreOvershootM = 2 * ProjectedRadiusM` - in the
   direction away from the face-hole rows, which is read from the geometry
   rather than assumed. The overshoot is the bore's own size; no view
   outline is read, and a test asserts none is.
2. **The crossing predicate could not tell the two cases apart.**
   `PathCrossesCircle` asks only whether a segment comes *within* the
   radius, which a segment starting at the centre satisfies trivially - it
   reported `crossingsProven=4|crossingFailures=None` for the cut that
   produced a section with no bore in it. The bore now requires
   `PathFullyCrossesCircle`: both line-circle intersections inside the same
   segment. Column holes keep the weaker test on purpose, because the first
   and last hole on the column sit at the segment's endpoints.

Checked arithmetically against the real r50 coordinates: the old waypoint
fails the new predicate, the new waypoint passes it.

The two r50 reporting defects are also fixed - `SECTION_REQ_CANDIDATE`
resets `a=`/`b=` per requirement, and a diameter requirement is searched as
a radius and then as a linear span, because a bore that has been cut open
is the gap between two walls rather than an arc.

**This prediction was made and did not hold.** It said Section View J-J
would contain a bore opening and that `INNER_BORE_D40` and
`FIT_BORE_D47_H7` would report `found=True`. The run produced an identical
section view. Recorded rather than deleted: the reasoning looked sound and
was still wrong, which is why the section-line coordinates now have to be
measured instead of reasoned about.

## r50 - LAYOUT proved; the section does not contain the bore

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r50`, deployed.
Companion suite **542/542**. Run
`macro_qa/20260805_041027_P-0251-14A-001`. Failed required stages **7 to
6**; `LAYOUT` FAILED to PROVED.

Still failing: `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`,
`ORDINATE_SCHEME`, `SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`,
`FINAL_QA`.

### The decisive finding: five of seven requirements have no geometry

Measured in Section View J-J, decode self-verified
(`records=38|entities=38|recordsMatchEntities=True`):

```text
X = 0.008;0.002;-0.008;0.009;0.004;0.003;-0.009
Y = 0.019;-0.097;0.062;0.098;0.018;0.017;-0.098
R = 0.002100;0.035000;0.036000
```

| requirement | nominal | present in the view |
| --- | --- | --- |
| `OVERALL_THICKNESS_18` | 0.018 | **yes** - X pair -0.009/0.009 |
| `BORE_STEP_DEPTH_12` | 0.012 | **yes** - X pair -0.008/0.004 |
| `LOWER_WALL_STEP_11_5` | 0.0115 | no |
| `INNER_BORE_D40` | 0.040 | no |
| `FIT_BORE_D47_H7` | 0.047 | no |
| `LONG_VERTICAL_REF_173_6` | 0.1736 | no |
| `LOWER_VERTICAL_REF_104_8` | 0.1048 | no |

**The bore is not in the section view.** The only arc radii are 0.0021 (M5
tap drill) and 0.035/0.036 (the plate's rounded top). Neither 0.020 nor
0.0235 appears, and no 0.040 or 0.047 span exists on either axis. The
bore's cut walls would sit at Y = 0.062 +/- 0.0235; the only nearby Y is
0.062 itself - the bore centre, which is exactly where the section path's
first waypoint starts. **The cut begins at the bore centre rather than
passing through the whole bore.** That is a waypoint question for
`Module17_SectionPath`, not a dimension-engine question, and it is the next
thing to decide.

Two defects in r50's own reporting, found by reading its own output and
fixed next iteration: `SECTION_REQ_CANDIDATE` prints stale `a=`/`b=` on a
`found=False` row, and the diameter requirements are searched only as arc
radii when a sectioned bore is a linear span between two cut walls. Neither
changes the table above.

### The rest of r50

One behaviour change - the LAYOUT scale check, below. Everything else is
read-only instrumentation, added because three of the four remaining
problems were unanswerable:

| open question | what r50 will report |
| --- | --- |
| Is the geometry for the seven section requirements even present? | `SECTION_REQ_CANDIDATE|key=|nominalM=|kind=|found=|a=|b=` per requirement |
| What curves does Section View J-J expose at all? | `SECTION_GEOM_SUMMARY`, `SECTION_GEOM_SEGMENTS`, `SECTION_GEOM_X/Y/R` |
| Which coordinate frame do a section view's polylines use? | `SECTION_GEOM_FRAME|polylineBox=|sheetOutline=` |
| Is `Section View J-J` really off-scale, or is the check misreading a flag? | `VIEW_SCALE_READBACK|useSheetScale=|scaleDecimal=` per view |

`InventorySectionGeometry` decodes `IView.GetPolylines7` using the
documented record layout and **proves its own decode**: four range guards
report `Desynchronized:<field>` instead of emitting invented coordinates,
and `recordsMatchEntities=` compares decoded records against the returned
entity array, which the Help states is positionally paired. The r40/r41
visibility classifiers shipped without such a control and measured nothing;
this one cannot fail that way silently.

### The LAYOUT scale check is fixed

It was reading a flag as if it were a ratio. SOLIDWORKS 2025 Help,
`IView::UseSheetScale` - *"If the property is 0, then it is possible that
the view scale is the same as the sheet scale"* - and `IView::UseParentScale`
is what a section view uses, so `Section View J-J` reads 0 while being drawn
at exactly the sheet ratio. That one line failed the whole LAYOUT stage in
r49.

`ValidateLayout` now accepts `UseSheetScale = 1` as before and, only when
that does not settle it, compares `IView.ScaleDecimal` against the sheet
ratio proved by `ISheet.GetProperties2`. `ViewScaleMatchesSheet` **fails
closed**: an unproved sheet scale, a zero denominator or any read error is
not a match, because accepting a view whose scale nobody measured is the
opposite of what the stage exists for. A view genuinely drawn at another
scale still fails, and the failure now carries
`useSheetScale=/viewScale=/sheetScale=/sheetScaleProven=`. The detail-view
3:1 rule and the isometric exemption are untouched.

## r49 - Root 1 closed: the section exists

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r49`, deployed.
Companion suite **526/526**. Run `macro_qa/20260805_034637_P-0251-14A-001`.

**Section View J-J was created from the model graph.** Failed required
stages **9 to 7**; `SECTION_GEOMETRY`, `SECTION_CLEARANCE` and
`ANNOTATION_EXTENTS` all went FAILED to PROVED.

```text
SECTION_PATH|view=Drawing View1|label=J|resolved=True|reason=None|segments=3
  |boreBasis=PositionProved:ProjectionAnchorUnavailable
  |distinctColumns=2|distinctRows=3|columnHoles=3
  |crossingsProven=4|crossingFailures=None
SECTION_CREATED|...|sectionView=Section View J-J|segments=3
  |selectionsVerified=3|sectionLine=Read|values=49
```

Still failing: `LAYOUT`, `NATIVE_CALLOUT_COVERAGE`, `VIEW_PROJECTION`,
`ORDINATE_SCHEME`, `SECTION_DIMENSIONS`, `MANUFACTURING_DEFINITION`,
`FINAL_QA`. They are now four independent problems, no longer one:

1. **`LAYOUT` - scale flag on the section view.** Sole failure line:
   *"Non-isometric view is not using the proved sheet scale: 'Section View
   J-J'."* `Module9_LayoutEngine.bas:718` demands
   `IView.UseSheetScale = 1`; a section view inherits its PARENT view's
   scale instead, a different flag value at the same rendered ratio. **No
   per-view scale is logged**, so the actual `UseSheetScale` and
   `ScaleDecimal` are not in evidence yet - log them before changing the
   check.
2. **`ORDINATE_SCHEME` - new view, new schemes.** Two new schemes came with
   the section view. Horizontal resolved its datum and created nothing
   (`swCreateOrdDimErr_Success|createdReadBack=0`); vertical was refused,
   `NoBucketAvailable;outline=NoMappedBottomEdge(edges:126,curve:57,notHorizontal:42,span:8,map:19)`.
   A section outline has no single mapped bottom edge, so
   `BottomOutlineGeometry` is the wrong datum policy for a section view.
3. **`SECTION_DIMENSIONS` - reachable for the first time**, `7 requirements
   / 0 satisfied`. `OVERALL_THICKNESS_18`, `BORE_STEP_DEPTH_12`,
   `LOWER_WALL_STEP_11_5`, `INNER_BORE_D40`, `FIT_BORE_D47_H7`,
   `LONG_VERTICAL_REF_173_6`, `LOWER_VERTICAL_REF_104_8` have never had a
   view to live in until now. This is the largest remaining work package.
4. **`NATIVE_CALLOUT_COVERAGE` + `MANUFACTURING_DEFINITION` - one family.**
   `op:EXTRUDEDCUT`, `missing=Attachment`: the stepped bore, which has no
   anchor in ANY view. Its axis lies in the section's page plane, so the
   section does not give it one either. In the reference drawing that bore
   is dimensioned in section (D40, D47 H7), not called out - so the open
   question is whether a bore callout is the right requirement at all.

`VIEW_PROJECTION` improved and still fails: projections 22 to 33, accepted 6
to 8, locations without 5 to 3. The two near-side M5 holes now project in
the section view. The three left are the stepped bore and the two far-side
M5 holes.

The user's decision, 2026-08-05: *let the section path accept
position-proved projections, keep anchor gate for dimensions.* Implemented
as two separate requirements rather than one loosened requirement:

| consumer | needs | requirement |
| --- | --- | --- |
| dimensioning, callouts, ordinates | to **select** the feature | `QualificationFailureReason` - unchanged, still demands a selectable anchor |
| section-line waypoints | to know **where** it lands | `PositionFailureReason` - identity, configuration, page frame, axis normal to view |

`Module13_ProjectionResolution` still decides `Accepted` through the anchor
gate alone, and a contract test asserts no component other than
`CViewHoleProjection` and `Module17_SectionPath` can even name the weaker
proof. `Module17` contains no `PrimaryAnchor` and no `SelectEntity`, which
is what makes a page position the right requirement there.

A section built on the weaker proof announces itself:
`SECTION_PATH_BORE_BASIS|...|use=WaypointsOnly` and `boreBasis=` in every
path summary.

**What should happen on the next live run.** In Drawing View1 the bore reads
`axisNormal=True|frame=Page|pageX=0.137812223|pageY=0.210324890|projectedRadiusM=0.020000000`
and was rejected only for `ProjectionAnchorUnavailable`, so it should now
qualify; the six counterbores give 2 distinct columns and 3 distinct rows.
Drawing View2 reads `axisNormal=False` for the same bore and should still
refuse it. Once the section exists, the post-section projection pass
(`Module2_DrawingPipeline` step after section creation, already wired)
should give the M5 side holes - axis `(0,1,0)` - projections in the J-J
view, which is the only place they can project as circles, and that is what
`missing=Attachment` on both callout families is waiting for.
**All of this is prediction, not evidence, until the run happens.**

Face holes still require `Accepted`; all six qualify today, so nothing was
loosened without cause.

## r48 - Root 1 answered: the stepped bore is obscured in every orthographic view

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r48`, deployed.
Companion suite **510/510**. Run
`macro_qa/20260805_033146_P-0251-14A-001`.

Route D now runs (`R23_PROJECTION_SELECTION_PRECONDITION|preexisting=1|cleared=True|remaining=0`)
and the stepped bore's anchor line changed decisively:

```text
before r48:  mappedEdges=0  firstUnmappedRoute=A:Nothing:err0;B:Nothing:err0;D:RefusedPreexistingSelection
r48:         mappedEdges=4  inventoryConfirmed=0  firstReject=MappedEntityNotInVisibleInventory
```

All four circular edges now map. `IView.SelectEntity` accepts them and
`ISelectionMgr.GetSelectedObjectsDrawingView2` proves the owning view, so
they exist in Drawing View1. **None is in `IView.GetVisibleEntities2`**,
which the 2025 Help defines as entities "not completely obscured by other
entities in the view".

**The stepped bore's circular edges are completely obscured in the Front
view.** That is the first evidence-backed answer to the question open since
the r40 review; the r40/r41 "obscured" counts were void and this is not
them. The fail-closed guard is behaving correctly: it refuses to publish an
obscured edge as a circular anchor.

### Root 1 is a design conflict, not a defect

`Module17_SectionPath.ResolveBoreProjection` requires an **accepted**
projection. But the section path only ever consumes `BoreProjection.PageX`,
`PageY` and `ProjectedRadiusM` - a proved page **position**, not a
selectable circular anchor. `Module13_ProjectionResolution` already records
the page centre for an unanchored location and says so in its own comment:
"An unanchored location still has a provable position."

So the requirement is stricter than the use. A bore that is hidden in every
orthographic view can never satisfy it, which is precisely why the reference
drawing shows that feature in the J-J section - and the section cannot be
created because it demands the projection the hidden bore cannot provide.
Circular dependency.

**This needed a decision, not a patch.** The user took it on 2026-08-05: the
section path accepts a position-proved projection for its waypoints while
the dimensioning and callout paths keep requiring a selectable anchor.
Nothing about the anchor gate weakened. Implemented at r49, above.

Minor, noted not fixed: the six counterbores now make two extra Route D
calls each, and `firstReject=MappedEntityNotInVisibleInventory` is recorded
on locations that were nonetheless accepted, which reads as a failure when
it is not.

## r47 - Root 1 addressed in source; NOT DEPLOYED

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r47`.
Companion suite **507/507**. **No r47 deployment, compile, or live run has
happened** - the user holds the deploy decision for this iteration. The last
live evidence is r46 at `macro_qa/20260805_001521_P-0251-14A-001`.

r47 contains two source changes, both unproven live:

1. **Route D gate re-keyed** from `Not visibleInventoryAvailable` to "Route A
   and B both declined". This is Root 1: Drawing View1 has a visible
   inventory, so Route D was never tried for the stepped bore whose four
   circular edges Route A declines. Safety is preserved because the existing
   Route C membership check still applies when an inventory exists, proving a
   Route-D entity twice rather than once.
2. **Callout completeness reason emitted** as
   `R23_CALLOUT_INCOMPLETE|family=|missing=|` and in the failure text.

The r40-r42 visibility classifiers were removed. Their model-space form
matched nothing, so `unmappedObscuredEdges` measured nothing; the finding it
was built to settle is recorded in
[SOLIDWORKS_API_VALIDATION.md](SOLIDWORKS_API_VALIDATION.md) instead.
**Whether the stepped bore's edges are obscured remains unestablished.** If
Route D now maps it and Route C rejects it, that will be the first real
answer.

### What r47 is expected to move, and what it cannot

If Route D maps the bore: `VIEW_PROJECTION`, `SECTION_GEOMETRY`,
`SECTION_DIMENSIONS`, `SECTION_CLEARANCE` and `LAYOUT` should follow, since
all five fail behind the missing section. If Route D cannot map it, the run
will say so explicitly instead of never trying.

`NATIVE_CALLOUT_COVERAGE` and `MANUFACTURING_DEFINITION` need one more thing
regardless: **`op:EXTRUDEDCUT` still reports `dia:0.000000000`.** That is a
separate reader from the Hole Wizard path fixed at r46 - an extruded cut has
no feature-data diameter and its size has to come from geometry. Not
attempted at r47.

## r44 - Phase 2 landed; two roots left

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-05-r44`.
Companion suite **491/491**. Latest production run:
`macro_qa/20260805_000138_P-0251-14A-001`.

Against the 18:45 r37 baseline:

| | r37 18:45 | r44 |
|---|---|---|
| Annotations imported | 0 | 5 |
| Layout moves | 0 | 2 |
| Display dimensions | 0 | 8 |
| Ordinate dimensions | 0 | 4 |
| Ordinate groups | 0 | 2 |
| Failed required stages | 11 | 9 |

`DIMENSION_ARRANGE` and `ORDINATE_SCHEME` are PROVED. Four defects were
fixed on live evidence: the `ActivateView` false negative, the layout
sequencing, the raw `ISelectData.View` assignment that raised 91, and the
`GetPolylines7` drawing-space comparison. A fifth, the stale selection before
ordinate creation, unblocked ordinate creation itself.

**The nine remaining failures reduce to two roots.**

1. **The P-0251 stepped bore has no accepted projection.** No route maps its
   four circular edges, so `ResolveBoreProjection` finds no singleton bore,
   no section path resolves, and `SECTION_GEOMETRY`, `SECTION_DIMENSIONS`,
   `SECTION_CLEARANCE` and `LAYOUT` all fail behind it, along with part of
   `VIEW_PROJECTION`. Whether those edges are obscured is **not
   established** - drawing-space visibility testing needs a drawing entity,
   and no route produces one for this bore.
2. **Hole Wizard definitions read as all zeros**, failing
   `NATIVE_CALLOUT_COVERAGE` and `MANUFACTURING_DEFINITION`. Needs the
   read-only member probe described in
   [SOLIDWORKS_API_VALIDATION.md](SOLIDWORKS_API_VALIDATION.md); no code has
   been changed on that contract.

`ANNOTATION_EXTENTS` and `FINAL_QA` are consequences of the two above.

Open and unexplained: the vertical ordinate group reported
`selectionsAppended=4|expectedSelections=3` and was still accepted.

## r39 - Phase 1 confirmed live; mutating runs are automated

`MACRO_SOURCE_REVISION` is `target-spec-hybrid-v2-2026-08-04-r39`.
Companion suite **464/464**. Deployment/readback **38/38**, programmatic VBE
compilation `verdict=Clean`, nine read-only probes **9/10 semantic stages**
(`probe_runs/20260804_230543`, byte-identical to the r37 run - no regression).

**Mutating production runs now go through
`tools/production-runner/Run-R23Production.ps1 -AllowMutation`**, per the
"Automated mutating-run exception" in `Agents.md`. It refuses to invoke
`Module1_Main.main` unless `R23_PrepareProductionRun` logged
`R23_PREFLIGHT_END|ready=True`. Visual and manufacturing acceptance are
unchanged and remain the user's judgement.

**First r39 production run: `macro_qa/20260804_232440_P-0251-14A-001`.**
Against the 18:45 r37 baseline:

| | r37 18:45 | r39 23:24 |
|---|---|---|
| Annotations imported | 0 | **5** |
| Layout moves | 0 | **2** |
| Views created | 3 | 3 |
| Section count | 0 | 0 |

Both r38 fixes are confirmed by live evidence. `LAYOUT` still fails, but the
cause changed from a sequencing bug to a real dependency: the final
structural pass runs, dispatches to the P-0251 reference zones, and finds no
J-J section. It cannot pass until the section defect is fixed.

Remaining roots from
[R23_DEFECT_REVIEW_AND_PLAN_2026-08-04_POST_1845.md](R23_DEFECT_REVIEW_AND_PLAN_2026-08-04_POST_1845.md),
all unchanged: the section path has no accepted singleton bore projection;
Route D is suppressed when a visible inventory exists; Hole Wizard definitions
read as all zeros; the bottom-outline datum is unresolved; and
`MODEL_IMPORT_COVERAGE` again reported PROVED while `Drawing View1` imported
nothing.

`Module4_ModelItemImporter.bas:1029` is now blocking diagnosis rather than
merely untidy: the r39 run reports
`Dimension arrange API error in 'Drawing View1': 0:` for two views, and the
handler destroys the error number before printing it.

The 18:45 production run failed ten of twenty-three required stages. The
review at
[R23_DEFECT_REVIEW_AND_PLAN_2026-08-04_POST_1845.md](R23_DEFECT_REVIEW_AND_PLAN_2026-08-04_POST_1845.md)
traces those ten to five independent roots. r38 addresses the first two:

- **Zero annotations imported.** `IDrawingDoc.ActivateView` returns False on
  this build even when the view activates. Annotation import branched on that
  raw result and skipped every view. The four mutating call sites
  (Modules 14, 15, 16, 17) now activate through
  `Module8_RuntimeSupport.ActivateDrawingView`, which proves activation by
  active-view readback.
- **The drawing was never laid out.** The only layout pass ran before the
  section and isometric views existed, so the P-0251 reference-zone validator
  failed unconditionally and `Layout moves` was 0. There are now two passes:
  a rough pre-placement at step 3 that defers the verdict, and a final
  structural pass at step 11b that runs the reference zones once every
  required view exists.

Still open from that review and unchanged by r38: the section path (no
accepted singleton bore projection), Route D suppression when a visible
inventory exists, all-zero Hole Wizard definition reads, and the unresolved
bottom-outline datum. Automatic rescaling and content-envelope repositioning
remain retired.

## R23 production wiring and current runner evidence

**2026-08-04.** Phase 11 production wiring is complete in source:
`Module2_DrawingPipeline` now owns the location graph, uses the R23 import,
ordinate, section, callout, envelope-layout, and shared semantic-QA route, and
does not call the retired feature-list ordinate or hardcoded callout writers.
The generic arranger skips section semantic lanes. `Module6_QAEngine` requires
the shared semantic stages rather than count-only import/location/ordinate
stages. `MACRO_SOURCE_REVISION` is
`target-spec-hybrid-v2-2026-08-04-r37`.

Static verification for r37: companion suite **435/435 passed**. Focused
P-0251 scratch evidence is retained at
`test_assets/iteration_evidence/probe_runs/20260804_164014/probe_log.txt`:
guarded deployment/readback was **38/38**, programmatic VBE compilation was
`verdict=Clean`, and all nine read-only probes completed. No probe created or
mutated drawing or model state.

The user accepted the current layout as-is on 2026-08-04. r31 retains initial
grid placement and normal dimension arrangement, but retires automatic
content-envelope view movement and rescaling from production. `FINAL_LAYOUT`
now records `UserAcceptedLayoutAsIs|automaticClearance=DeferredByUser`; this
waives the automatic layout gate without claiming the five measured clearance
failures are resolved.

That run also closes R23-006. Both historical curve-read orders returned the
same circular evidence for counterbore, M5 tapped, mirrored, and extruded-cut
representatives: `IsCircle=True`, seven `CircleParams` values, equal radius,
and zero endpoint closure. `R23_CURVE_ORDER_END|failures=None`; the part
remained unchanged.

R23-502 is satisfied at r37. `IView.GetPolylines7` returned no edge array in
both valid HLV views. The only fallback is therefore the documented scoped
`IView.SelectEntity` route:
the selected model edge must return a drawing entity whose owner is read back
as the requested view. Both vertical datums resolved from straight lower
outline edges; no hole or view-bound substitute was used.

The r37 semantic gate proves **9/10** stages. `FINAL_LAYOUT` is the documented
user waiver and `ORDINATE_SCHEME` is proved. The sole remaining stage is
`MODEL_IMPORT_COVERAGE` for `Drawing View4`: a read-only copy of a manual
drawing has no macro-created or imported coverage there. The production path
marks coverage only after ordinate creation succeeds and is read back, so this
stage remains deliberately unproved until an authorized creation run. This is
not a production run or three-fixture manufacturing acceptance.

The three-fixture regression cannot begin yet: the workspace contains only the
P-0251 disposable `.SLDDRW`; it contains no P-0252 disposable drawing. Those
drawings and fresh authorization for the remaining creation operations are
required before production acceptance can be completed.

Source-hygiene inventory: all 38 manifest-managed `.bas`/`.cls` files are ASCII,
CRLF, BOM-free, have no trailing whitespace, declare `Option Explicit`, and
meet the 79-character limit. r37 passes 435/435 static tests, deployment
readback, programmatic VBE compilation, and the focused read-only runner.

## Historical first batch run (superseded by r26)

**2026-08-04.** The probe runner is now the required path for testing the
macro and for compiling what is loaded in the VBA editor. `CLAUDE.md`,
`Agents.md` and the R23 implementation handoff were amended to say so: the
agent runs the command and reads
`test_assets/iteration_evidence/probe_runs/<timestamp>/probe_log.txt`
instead of handing the user a deploy-compile-run-paste sequence.

```powershell
powershell -ExecutionPolicy Bypass -File `
  ".\tools\probe-runner\Run-R23Probes.ps1" -Deploy `
  -DrawingPath "<full path to the fixture drawing>"
```

Two phases got their first live evidence out of that run:

- **Phase 9** produced `envelopes=4|clearanceChecks=22` and
  `plan=RescaleRequired|suggestedScaleFactor=0.527974` -
  `requiredHeightM=0.479190` against `usableHeightM=0.253000`, which is the
  measured form of the argument that reversed R23-907. Ten clearance
  failures stand and `R23-903`/`R23-904` are still unrun; applying a plan
  needs `allowMutation` and a non-protected target drawing.
- **Phase 10** ran read-only and judged all ten stages:
  `proved=3|failures=7`.

**Historical regression.** Every view in that run
reported `PROJECTION_VISIBLE|type=Edge|count=0|`
`source=IView.GetVisibleEntities2`, so Phases 3 through 7 read zero
downstream and six of Phase 10's seven failures follow from it. These same
probes returned nonzero when run standalone from the VBA editor, and their
gates were satisfied on that evidence.

Phase 8 is the control that keeps this narrow: `satisfied=7|missing=0` in
the same batch run, identical to standalone, because it does not read
visible entities. The drawing was open, correct, and bound to the
authorized part - `part=` reads the `test_assets\models\` copy on every
`_BEGIN` line. Route D subsequently resolved the runner path; r26 produced
44 projections and 11/11 proved selection owners. Treat this section as
failure provenance, not the current state.

## R23 probe-automation tool built and live-verified

**2026-08-04.** Completed
[R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md](R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md)
PA-100 through PA-112 - the whole plan, including the first successful
live run. Evidence:
`test_assets/iteration_evidence/probe_runs/20260804_054014/probe_log.txt`.

One command (`tools/probe-runner/Run-R23Probes.ps1 -DrawingPath ...`)
now deploys, resolves and executes the VBE Compile command by caption
(`id=578|caption=Compile Fable`, resolved live - not in the SOLIDWORKS
API corpus per finding 5.1, never hardcoded), confirms all 21 standard
modules load clean, opens the authorized part then the drawing
read-only, switches the active document at the right two points, runs
all nine `R23_Probe*` entry points in dependency order, and writes a log
the agent reads directly. Every probe reported `mutations=0`/
`creations=0` and `drawingUnchanged=True`; every `part=` field resolved
to the `test_assets\models\` copy, not the `V:` network sibling
(finding 5.6 did not recur).

Two live-only bugs surfaced and were fixed during this session (full
narrative in the plan doc, section 12, and
`docs/SOLIDWORKS_API_VALIDATION.md`): VBIDE `.Caption` carries the raw
`&` accelerator marker (`"Compi&le Fable"`), which broke the caption
match until compared against the cleaned string instead of the raw one;
and PowerShell's native COM calling convention cannot call `OpenDoc6`
(fails both by late binding and by a direct interop bracket-cast), fixed
with a new compiled helper, `SolidWorksDocumentOpener.cs`, mirroring
`SolidWorksMacroInvoker.cs`'s existing pattern.

Offline suite is now 424/424 passed, including the production-wiring
contracts. Preflight reports `Managed components: 38`.

`MACRO_SOURCE_REVISION` was not bumped - these modules are not on the
production drawing path.

## R23 Phase 8 gate SATISFIED; probe automation authorized

**2026-08-02.** Third live run of `R23_ProbeSectionDimensions`:
`satisfied=7|missing=0|duplicated=0|sectionOrdinates=0|`
`requirementFailures=None`, `mutations=0`, `drawingUnchanged=True`,
selection 0 before and after. All seven P-0251 section requirements exist
in the drawing, are matched to a real dimension, and carry what they claim.

Both bores resolved via `diameterDisplaySource=TextPrefix` with
`prefix=<MOD-DIAM>`. The drawing shows the diameter symbol in the
dimension's TEXT PREFIX while `Diametric` stays False - the exact case the
third reading of R23-804 exists to distinguish. Both `GetText` forms return
the literal `<MOD-DIAM>` token rather than a rendered glyph, so the
codepage-216 comparison is dead code on this build and the token match is
what decides.

`FIT_BORE_D47_H7` reads `toleranceSatisfied=True` with
`toleranceProvenance=PresentOnDrawing.TargetSpecReferenceAuthority`
`.NotModelData` against a live `holeFit=H7|maximumM=0.000025`.

### The run before it failed for a reason worth keeping

`R23_SECTIONDIM_FATAL|reason=UnauthorizedFixture` with the part resolved to
the `V:\VEEMAP\SW_data\` copy. The fixture guard fired correctly: the
drawing had bound its part reference to the network sibling instead of the
authorized `test_assets\models\` copy. Opening the local part BEFORE the
drawing rebinds it.

This matters beyond the one failure. If earlier runs bound different
copies, that is an alternative explanation for the Phase 0 versus
2026-08-01 dimension-state difference, and the `part=` field on every
`_BEGIN` line is what settles it.

### Probe automation authorized

The user authorized programmatic full-project VBA compilation and amended
the compile gate: **read-only `R23_Probe*` entry points no longer require a
preceding manual Debug > Compile Project.** Mutating runs and production
acceptance are unchanged.

The mechanism is already proved in this repo.
`tools/swp-deploy/Module0_SourceDeployment.bas:214` reaches the VBE object
model and executes a built-in command through
`CommandBars.FindControl(1, 3, "", False)`, because `VBProject.SaveAs`
raises 748 on a host-managed `.swp`. Compile is the same kind of control.

Amended: `Agents.md`,
`docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md`,
`docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md`. The build itself is
planned, not started - see
[R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md](R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md).

`README_IMPORT.md` was also corrected: its import instructions still
described a nine-module inventory and three data classes, wrong since
Phase 1 and flagged by the code review. It now defers to the deployment
manifest.

**Still open:** Phase 9 and Phase 10 have not been re-run since their
defect fixes.

## R23 Phase 10 source complete, awaiting first live run

**2026-08-01.** `Module19_SemanticQA.bas` (25 procedures). Statically
verified only. Built while Phases 8 and 9 await their third live run.

Ten required stages, each with an evaluator that can prove or fail it:
`MODEL_INTENT_CATALOG`, `MODEL_IMPORT_COVERAGE`, `NATIVE_CALLOUT_COVERAGE`,
`PHYSICAL_LOCATION_GRAPH`, `VIEW_PROJECTION`, `ORDINATE_SCHEME`,
`SECTION_GEOMETRY`, `SECTION_DIMENSIONS`, the retained
`MANUFACTURING_DEFINITION`, and `FINAL_LAYOUT`.

**The module changes nothing at all** - not even behind an `allowMutation`
gate, unlike every other R23 module. A QA engine that repairs what it is
judging cannot report on it.

What the count-based checks missed, and what replaces them:

- **"Nonzero import" is satisfied by one view receiving everything.**
  Coverage is now reported per view and per category - accepted
  projections, covered in X, covered in Y, annotations attached - and a view
  with accepted projections and none of the three fails by name.
- **A note-token check passes on text that has drifted from the geometry**
  and fails on a correct drawing worded differently. Section dimensions are
  judged by type, nominal, attachment and tolerance; `GetText` and
  `GetNotes` appear nowhere in the module.
- **A number with no source is not evidence.** Every manufacturing field is
  emitted beside its proof source, and blank, `None` and `Unproven` all fail
  as `NoProvenance`.
- **A projection COUNT hides an unprojected location** behind the ones that
  did project, so every identity-proven location with no accepted projection
  is named.
- **Duplicate keys mean the graph has lost track of which hole is which.**
  Physical, family definition and section requirement keys are all checked,
  and `DuplicateKeyReport` returns `"None"` rather than an empty string so
  "no duplicates" and "the check did not run" stay distinguishable.

`CollectRetainedDefinitions` runs the loop locally but takes every
JUDGEMENT from `Module16_CalloutDefinition`'s public surface - what is a
native callout, which family it belongs to, which definition is retained -
so the two cannot drift on the part that matters.

**Deferred:** `Module6_QAEngine` still runs the count-based checks on the
reachable production path. Switching over is Phase 11, and doing it now
would put unproven gates in front of a deployable macro.

**Phase 11 was deliberately not started.** It reorders the production
pipeline and bumps `MACRO_SOURCE_REVISION`, which switches the deployable
macro onto Phases 5 to 10 - none of which has a green live run yet. That
sequencing needs the Phase 8 and 9 third run first.

Verification: 24 Phase 10 contracts, suite 381 tests with the same five
stale R20 failures. Preflight 36 components.

## R23 Phases 8 and 9 second live run: R23-907 reversed, six defects fixed

**2026-08-01.** Both probes ran read-only. Phase 8 matched every
requirement; Phase 9 completed end to end for the first time.

### Phase 8: `satisfied=5|missing=0|duplicated=0`

All seven nominals read exactly - 0.018, 0.012, 0.0115, 0.040, 0.047,
0.1736, 0.1048 - so every P-0251 section requirement exists in the drawing
and is matched to a dimension.

**The nominal route is settled.** Every one answered
`nominalRoute=Obsolete.GetSystemValue2`. `GetSystemValue3` declined both
configuration modes on all seven, so `swAllConfiguration` has been removed -
a route with live evidence against it is not kept for insurance. Two remain:
the supported configuration call, which is what answers for imported model
dimensions, and the obsolete `GetSystemValue2`, which is the only thing that
answers for a drawing-authored reference dimension on this build.

**The two flagged requirements are the bore diameters**, on
`NotDisplayedAsDiameter:2`. Every section dimension returned
`diametric=False` with `diametricKnown=True` - a real answer. Before that is
called a defect in the drawing, the dimension's text PREFIX has to be read:
a drawing can carry the diameter symbol there while the diametric flag stays
False, and then the sheet reads correctly even though the record does not.
`ReadDiameterPrefix` now reads both the rendered prefix and its definition
form, where SOLIDWORKS writes `<MOD-DIAM>`, and `DiameterDisplaySource`
names which of the three sources answered.

`FIT_BORE_D47_H7` reads
`toleranceSatisfied=True|toleranceProvenance=PresentOnDrawing.TargetSpecReferenceAuthority.NotModelData`
against a live `holeFit=H7`, `maximumM=0.000025`. Exactly the R23-806
position.

### Phase 9: completed, five defects, and R23-907 reversed

The sheet measured cleanly - A3 0.420 x 0.297, `contentBorder=Measured` from
`ISheet.GetZoneMargin`, `titleBlock=Absent` - and four envelopes were built
with `annotationEnvelopes=3`.

**The decisive bug: the section-line arrow block is 9 doubles, not 11.**
`Drawing View4` returned `items=49`, and 49 = 2 header + 1 numSegments +
7x3 segments + 9 + 9 arrows + 7 text. Three segments - the J-J path exactly.
Counting 11 matched nothing, which is why every envelope reported
`arrow=0|section=0`. The dry-run design caught it as designed: it refused to
parse rather than producing plausible coordinates.

Four more: an absent section line was reported as a failed parse; every
envelope line printed twice; the display-data frame check allowed 120 mm of
slack and tested only line start points; and rejected off-sheet points were
counted without a single coordinate being kept.

**R23-907 is reversed by the user** - "The views are allowed to rescaled as
per need". The reference drawing itself cannot satisfy the old rule: its
four envelopes need 0.479 m of height in the 0.253 m available. Rescaling is
now a gated, recorded remedy: one `ScaleDecimal` assignment, inside
`ApplyScaleToFit`, refusing without `allowMutation`, reading each new scale
back, and reporting every changed view by name. The factor is labelled an
estimate because annotation text does not scale with the view, so the
envelopes are re-measured rather than predicted. R23-908 survives: if the
content still does not fit after the rescale, the sheet is too small.

Verification: 357 tests with the same five stale R20 failures. Preflight 35
components. `MACRO_SOURCE_REVISION` remains `r22`.

## R23 Phases 8 and 9 first live run: four defects found and fixed

**2026-08-01.** Both probes ran read-only against the P-0251 reference
drawing. Neither gate is satisfied yet; all four causes were in my code.

### Phase 8: `satisfied=0|missing=7` with seven dimensions in the view

**The nominal never read.** Every one of the seven returned
`nominalAvailable=False`, and the dimension object was fine - the same
object answered `toleranceType`, `fitType` and the fit strings on the next
line. So `GetSystemValue3(swThisConfiguration, Empty)` specifically declined.
The seven are `RD1..RD7@Drawing View6`: drawing-authored REFERENCE
dimensions, not model dimensions imported from a sketch, and a reference
dimension has no configuration to ask about. Phase 0 read `D1@Sketch4`,
where that route works.

Without a nominal nothing can match, so all seven requirements reported
Missing while their dimensions sat in the view. `TryReadNominal` now tries
`swThisConfiguration`, `swAllConfiguration`, then the obsolete
`GetSystemValue2("")` and `SystemValue` as labelled last resorts, and names
the route that answered. When all decline it reports the shape of what came
back.

**The type rule was wrong for this drawing.** All seven are
`swLinearDimension = 2`. Phase 0's type-6 evidence describes an earlier
state of the same fixture. A diameter requirement now accepts type 6, 15
and the linear types, and `IDisplayDimension.Diametric` records which of
them the drawing displays as a diameter - reported, not used to reject,
because the nominals are 5.5 mm apart at the closest.

**The log printed everything twice**, because `CRunEvidence.AddInfo` prints
what it records and the probe printed the same lines again.

**One result worth keeping.** `RD4` carries
`toleranceType=8|fitType=0|holeFit=H7|minimumM=0.000000|maximumM=0.000025`
with both statuses 0 and two attached edges. That is H7 +0.025/0.000, live,
on a drawing reference dimension - independent corroboration of R23-806's
finding that the fit is drawing-authored and absent from the model.

### Phase 9: aborted before a single envelope was built

```
QA FAILURE: Controlled sheet has neither an ITitleBlock definition nor a
proved legacy title-block rectangle.
R23_ENVELOPE_FATAL|reason=SheetRegionsUnmeasured
```

I made a read-only probe depend on
`Module8_RuntimeSupport.MeasureControlledSheetRegions`, which is a
production fail-closed gate for a sheet the macro CREATES from the
controlled template. The designer's reference drawing has no `ITitleBlock`,
so it refused - and it had already attempted `SheetFormatVisible = True`,
meaning a run that promised `mutations=0` tried to make one. Same shape as
the Phase 5 `EmitRunEvidence` mistake: a probe leaning on a production gate.

`MeasureSheetRegions` now measures read-only - `ISheet.GetSize` plus
`ISheet.GetZoneMargin` - and reports what it cannot measure instead of
aborting. `BuildProtectedRegions` gates every rectangle on measured bounds,
because unset fields would have produced a degenerate rectangle at the
origin and false intrusion violations against it.

Nothing about the envelope logic was exercised: no evidence yet on the
display-data frame, the inverse round trip, the section-line grammar or the
envelope sizes.

Verification: 345 tests with the same five stale R20 failures. Preflight 35
components. `MACRO_SOURCE_REVISION` remains `r22`.

## R23 Phase 9 source complete, awaiting first live run

**2026-08-01.** `Module18_ContentEnvelope.bas` (33 procedures) plus
`CContentEnvelope.cls`. Statically verified only.

A content envelope is everything that travels with a view - model outline,
dimension primitives and text boxes, note extents, leader points, section
segments, arrows, and the J-labels with their text heights. Layout that
reasons about the model outline alone moves a view into a place its
annotations do not fit, which is how the old J-J label reached the zone
region in the first place.

**The frame work is the part that matters here.** Four sources document
their frame and one does not:

- `IView.GetOutline` - page frame, documented;
- `IAnnotation.GetPosition` - sheet-relative in drawings, documented;
- `INote.GetExtent` - sheet space, documented;
- `IView.GetSectionLineInfo2` - VIEW-SKETCH frame, proved by Phase 0;
- `IDisplayData` points - **the Remarks state no frame at all**.

Section geometry therefore goes through `ViewSketchToPage`, the exact
inverse of Module17's forward transform, and the two are round-trip checked
against each other before a single point is contributed. Display-data points
are contributed and their agreement with the view's own outline is COUNTED,
not asserted - `displayDataFramePageConsistent` and
`displayDataFrameInconsistent` are the fields to read on the first run.
Claiming a frame the Help does not state would be exactly the kind of
confident guess this project has paid for before.

Two more traps handled explicitly: `GetTextPositionAtIndex` is an OFFSET
from the display-data origin rather than a coordinate, and leader points are
consumed as triples from the returned array instead of via `GetLeaderStyle`,
whose value is OR-ed with attachment bitmask flags the corpus returns
mangled.

`GetSectionLineInfo2`'s grammar is ambiguous between its own Remarks and
`GetSectionLineCount2`'s - one layer double, or one per section line. Both
readings are walked in a dry run and the one whose consumption matches the
array length exactly is used. A parse that consumes the wrong number of
doubles produces plausible coordinates, which is the worst kind of wrong.

**The fixed upward bias is gone.** `PlanPlacement` packs rows from the
envelopes' own sizes and centres the block in the usable rectangle;
contracts assert `topBoundary -`, `Bias` and `rowCenterY` are all absent.
Nothing is pinned to a boundary, because a row pinned to a boundary has
nowhere to put the annotations that hang above it.

**Still unrun:** R23-903 and R23-904 both require mutation.
`ApplyPlacementPlan` is the only procedure that moves anything and refuses
without `allowMutation`.

Verification: 32 Phase 9 contracts, suite 334 tests with the same five stale
R20 failures. Preflight 35 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## R23 Phase 8 source complete, awaiting first live run

**2026-08-01.** `Module10_SectionDimensionEngine.bas` (28 procedures) plus
`CSectionRequirement.cls`. Statically verified only.

The seven P-0251 section requirements are stated once, each with its nominal,
its accepted dimension types and its lane. Reconciliation against the
dimensions already in the section runs before any creation path - Phase 0
counted seventeen imported dimensions there, and creating a second dimension
for something already dimensioned is the defect that ordering prevents.

Design points worth carrying:

- **REQUIRED and OBSERVED are separate field groups on the record**, and
  nothing writes an OBSERVED field from a REQUIRED one. A requirement that
  reports its own nominal back as the observed nominal proves nothing.
- **All four `IDimension` tolerance members are obsolete.**
  `GetToleranceValues`, `SetToleranceValues`, `GetToleranceFitValues` and
  `SetToleranceFitValues` are each superseded by an `IDimensionTolerance`
  member. The Phase 0 probe used the obsolete route; production does not.
- **`GetMinValue2`/`GetMaxValue2` return a STATUS**, with the value coming
  back by reference. A zero value with a failed status is not a zero
  tolerance, so the status is printed beside the value it qualifies.
- **Type 6 is accepted and type 15 is never required** - the imported 47/40
  records are live-proven `swDiameterDimension = 6`.
- **Per-dimension locals reset every iteration.** VBA block-scoped locals
  live for the whole procedure; that is exactly how the Phase 0 inventory
  mislabelled eleven of seventeen dimensions.

**The H7 provenance rule is enforced in code, not just documented.** The fit
is applied from the approved reference specification and recorded as
`TargetSpecReferenceAuthority.NotModelData`. A tolerance merely *found* on
the drawing is recorded as `PresentOnDrawing.` plus the same authority,
because Phase 0 read the part source directly and proved it carries none.

**Still deferred, all for the same reason the pipeline is not switched:**
R23-803's creation half, R23-808's lane-to-coordinate placement (Phase 9
owns the envelope), R23-809's call from `Module9_LayoutEngine`, and
R23-810's removal of the `Module7_TitleBlockEngine.bas:359-361` free-text
bore callout - which cannot go before real dimensions replace it.

Verification: 37 Phase 8 contracts, suite 302 tests with the same five stale
R20 failures. Preflight 33 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## R23 Phase 7 read-only gate SATISFIED (first live run)

**2026-08-01.** `resolvedPaths=1|segments=3|columnHoles=3|`
`crossingsProven=4|sectionFailures=None|creations=0|drawingUnchanged=True`.

The J-J path resolved in `Drawing View4` exactly as the reference approves:
`w1=0.207331779,0.237414746` (bore centre), `w2=0.207331779,0.167414746`,
`w3=0.192331779,0.167414746`, `w4=0.192331779,0.087414746`, with
`distinctColumns=2` and `distinctRows=3`. The other three views correctly
report `NoAcceptedSingletonBoreProjection` - the bore is not accepted there,
so no path is invented for them.

**The frame conversion cross-checks against the model, which is the part
worth trusting.** The arithmetic is exact
(`0.207331779 - 0.229331779 = -0.022`), but arithmetic only proves the code
does what it says. The corroboration is that the bore's Plucker moment is
`(0, 0.062, 0)` and its `viewY` is 0.062 exactly; every counterbore does the
same, with `viewY` equal to its moment's Y and `viewX` equal to its moment's
X minus a constant 0.022 - the view's centring offset, identical across all
seven holes. A wrong transform does not produce one shared offset across
seven independent points. Phase 8 onward depends on this transform, and it
has been wrong in this project before.

One evidence defect the run exposed: views rejected before crossings could
be tested were also reporting `NotAttempted` in `sectionFailures`. That is
the crossing proof's initial state, not a failure, and listing it beside
real reasons dilutes them. Now excluded.

**Still open:** R23-705's creation half, R23-706 and R23-707 all require
mutation and are unrun. R23-704's legacy removal waits on the pipeline
switch, the same as R23-609.

Verification: 23 Phase 7 contracts, suite 265 tests with the same five stale
R20 failures. Preflight 31 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 7 source complete

**2026-08-01.** `Module17_SectionPath.bas` (21 procedures) plus
`CSectionPath.cls`. Statically verified only.

The J-J path is four waypoints and three segments, every coordinate taken
from a proved projection: the stepped-bore centre, the same X at the highest
face-hole row, the minimum-X column at that row, then the same column at the
lowest row. The disproved strategy is entirely absent - no `extension`, no
`topY`/`bottomY`, no outline fractions - because a percentage of an outline
knows nothing about where the holes are, which is why the old upper label
landed in the zone region and the lower arrow in the part-identification
band.

Design points worth carrying:

- **The bore is a singleton family, not a radius threshold.** Family size
  comes from the graph, so a different part is not misclassified.
- **The grid is proved, not assumed.** Fewer than two columns or two rows
  is a stated rejection rather than an array index that happens to work.
- **Crossings are judged against each hole's own projected radius**, and
  the segment distance is clamped to the finite segment - an unclamped
  projection would report a circle beyond an endpoint as crossed because
  the infinite line passes through it.
- **The frame conversion happens exactly once per waypoint**, immediately
  before `CreateLine`. Nothing upstream holds view coordinates, so there is
  nothing to convert twice. Mixing frames is the defect section work has hit
  before.
- **Segment selection order is verified before the cut.** SOLIDWORKS reads
  the segments in selection order, so an unverified order produces a
  differently shaped cut.

Two defects caught before compiling: I passed an empty label instead of the
resolved one, and placed the section view's centre at waypoint 3 - a point
inside the source view, which would have stacked the section on top of it.
Placement is now a caller argument, because choosing where a view sits is
layout and belongs to a later phase.

**R23-704 is half met**, the same shape as R23-609: the new path is clean
and contracts prove it, but the legacy literals stay in
`Module2_DrawingPipeline.bas` (1525-1556) until the pipeline switches over.

Verification: 22 Phase 7 contracts, suite 264 tests with the same five stale
R20 failures. Preflight 31 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 6 read-only gate SATISFIED

**2026-08-01.** `definitions=3|definitionFailures=None|`
`counterboredFamilies=1|threadedFamilies=1|shapeFailures=None|`
`nativeCallouts=2|creations=0|drawingUnchanged=True`.

The M5 family retained its **native** callout
(`reason=CompleteAssociativeDefinitionAvailable`); the counterbore and
stepped-bore families retained controlled fallbacks with
`reason=NoNativeCalloutAttributedToFamily`. Both branches of R23-605
exercised live.

**A corroboration worth recording.** The M5 depth resolved to **12.4 mm**
from `swCalloutVariable_Tap_Drill_Depth`, and its thread depth to 10 mm.
The legacy hardcoded string in `Module7_TitleBlockEngine.bas` reads
`4.2 x 12.4 DEEP` / `TAP M5x0.8-6H x 10 DEEP`. Same numbers - now derived
from typed callout variables rather than typed by hand. That is independent
evidence the derivation is right, not just self-consistent.

Four defects the runs found, in the order they surfaced:

1. **Depth rule inverted.** `swEndCondBlind = 0`, so a *blind* hole needs a
   depth. The rule read `<> 0`, demanding one from a ThroughNext family and
   letting a blind tapped hole pass with no depth at all - the dangerous
   direction.
2. **NominalDiameter unproven everywhere.** No feature on this build
   declares one. The location knows it from its own proven cylindrical face
   radius, which is measured geometry rather than a declared parameter.
3. **Family keys destroyed the evidence format.** A family key is itself a
   delimited string, so `R23_CALLOUT_END` was unparseable. `EvidenceToken`
   escapes the rendering; the key itself is untouched.
4. **Attachment unprovable for a fallback.** Only the native path set it, so
   a fallback could never earn it and a native losing on completeness
   discarded the proof it had. Attachment is a property of the family's
   geometry. ATTACHED and ATTACHABLE stay distinguishable in the proof
   source.

Then one mislabel the passing run exposed: `threadedFamilies=2`, because a
thread *description* was treated as a thread. The counterbored family
carries the fastener size of the screw it clears with `threadDepthM=0`. A
tapped hole has a thread depth.

**Still open:** R23-604's creation half and R23-609's legacy removal. The
counterbore callout remains unattributed (`NoOwningProjection`) - the Phase 4
forward-alias limitation, not a Phase 6 defect, and no longer blocking now
that attachment is proved from geometry.

Verification: 27 Phase 6 contracts, suite 242 tests with the same five stale
R20 failures. Preflight 29 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 6 source complete

**2026-08-01.** `Module16_CalloutDefinition.bas` (20 procedures) plus
`CCalloutDefinition.cls`. Statically verified only.

A callout definition is either **native** - a SOLIDWORKS hole callout
carrying the Hole Wizard's own data - or a **controlled fallback** built
field by field from typed feature data. Never free text. Every field carries
the source that proved it, so a definition that looks complete can still be
shown to be unproven.

Design points worth carrying:

- **`IsHoleCallout` is the only classifier.** A native callout reports
  `Type2 = 6`, but so does an ordinary diameter dimension. No
  dimension-type constant is declared in the module, so none can be reached
  for.
- **Fields come from `GetHoleCalloutVariables`**, not from parsed text -
  `HoleFit`, `ShaftFit`, `ToleranceType`, `ToleranceMin`, `ToleranceMax`
  per variable. A rendered string cannot be validated field by field.
- **Quantity is unique physical locations.** Not features: one Hole Wizard
  feature plus a mirror makes many holes. Not edges: a counterbore
  contributes several per hole.
- **A callout resolving to two families is rejected**, not tie-broken.
- **Depth is required only when the end condition says the hole is blind**,
  and an unproven end condition fails on its own terms first.
- **R23-611 is stated as shapes, not part numbers**: one multi-hole
  counterbored family and one multi-hole threaded family. P-0251 satisfies
  it; nothing is keyed to it.

**R23-609 is half met, deliberately.** The new path has none of the
hardcoded text or name/radius scoring, and contracts assert their absence.
The legacy literals remain in `Module7_TitleBlockEngine.bas` (callout text
at 359-371, scoring at 405-435) because Module7 is still the reachable
production path and Module16 is not yet wired into `main`. Removing them now
would degrade the deployable macro while the replacement is disconnected.

Verification: 23 Phase 6 contracts, suite 238 tests with the same five stale
R20 failures. Preflight 29 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 5 read-only gate SATISFIED

**2026-08-01.** `schemes=4|horizontalSchemes=2|verticalSchemes=2|`
`creditedLocations=10|expectedLocations=10|coverageFailures=None`.
Read-only throughout: `creations=0`, `initialSelectionCount=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

Proven live: R23-500 (four measured schemes), R23-501 (stepped-bore centre
datum, selection and ownership proved), R23-503 (two X buckets), R23-504
(three Y buckets), R23-505 (all four side holes credited across two page
positions), R23-507 (`profileEntries=1`, stepped bore held out by family
size not radius), R23-509 (10 of 10).

**Still open, and stated as open:**

- **R23-502 is NOT met.** The vertical datum is the lowest projected hole,
  recorded `datumKind=ProjectionDerived`. The task asks for the part's
  bottom outline geometry. Not claimed on a weaker datum.
- **R23-506 is half met.** All four side locations are resolved and
  credited; none is dimensioned yet.
- **R23-508 is unrun.** It creates dimensions, so it needs authorization and
  a target that is not the manual reference drawing.

Two defects were found by the run rather than by me. The QA gate popped
`RESULT: FAIL` because the probe called `Module6_QAEngine.EmitRunEvidence`,
the production gate demanding fourteen pipeline stages a probe never runs.
Then the coverage gate reported `credited=8, expected=10`: I read
`CoincidentWithAnchoredKey` from the anchored end, where
`MarkCoincidentProjections` never sets it, and the unanchored twin holding
it had already been filtered out. Two of P-0251's four side holes were
silently uncredited. Both fixed, both now pinned by contracts.

Verification: 25 Phase 5 contracts, suite 215 tests with the same five stale
R20 failures. Preflight 27 components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 5 source complete

**2026-08-01.** `Module15_OrdinateScheme.bas` (36 procedures) plus typed
records `COrdinateScheme.cls` and `COrdinateBucket.cls`. Statically verified
only — nothing below is runtime-proven.

The scheme key replaces feature-family grouping with **view role + machining
face + datum policy + direction**, every part measured rather than read off a
name: machining face from the location's sign-normalized axis, view role from
Phase 3's axis-normal measurement via the existing eligibility tests.

Two Phase 3 findings are carried into the design rather than worked around:

- **Coverage is counted per distinct page position, credited to locations.**
  A bucket holds one selectable entity and the list of physical locations it
  represents. Coaxial holes collapse to one drawing entity, so demanding one
  dimension per location is unsatisfiable by construction; crediting only one
  of the pair would silently drop the other.
- **Small-hole membership is family size, not a radius threshold.** P-0251's
  stepped bore is excluded because it is a singleton family. A magic
  millimetre value would misclassify a different part.

Three defects caught before compiling: `IsOrdinateEligibleView` and
`IsDeferredCreationView` take `(graph, swView)`, not `(swView, graph)`; and
`IView.GetFirstDisplayDimension5` is obsolete **and** sheet-scoped by its own
Remarks, so a read-back built on it would credit other views' dimensions to
this scheme. Read-back now uses view-scoped `IView.GetDisplayDimensions` with
a before/after snapshot diffed by `ISldWorks.IsSame`.

Mutation boundary: `CreateOrdinateGroup` alone creates anything and refuses
without `allowMutation`; it also refuses when the datum is unproven.
`R23_ProbeOrdinateScheme` contains no `AddOrdinateDimension` call at all.

Verification: 22 Phase 5 contracts, suite 212 tests with the same five stale
R20 failures. Preflight 27 managed components. `MACRO_SOURCE_REVISION`
remains `r22`.

## Historical: R23 Phase 4 gate SATISFIED (sixth live run)

**2026-08-01.** Read-only throughout: `mutations=0`, `initialSelectionCount=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

`annotations=38`, `coverageFailures=None`,
`COVERAGE|holeCallouts=2|ordinates=10|diameters=0|toleranced=1|withFit=1`.
R23-412 is required-**category** coverage and all three required categories
are present. Gate met.

**The instrumentation settled the reconciliation question rather than
improving the number.** `IModelDocExtension.GetCorrespondingEntity2` returned
Nothing for all 38 annotations and every attachment **with error 0** —
`outcomes=draw1:unresolved:err0`, `resolved=0`, `eqMax=-1` (no comparison ever
ran, so the `swObjectEquality` Unsupported hypothesis is ruled out too). The
call declines, it does not fail. This is a part drawing and the member
resolves into an underlying part or subassembly, consistent with the
`componentContext=DrawingContextOnly` already recorded.

`reconciled=1` therefore stands, and is correct. `RD1@Drawing View7` proves
the R23-407 identity mechanism end to end. The other 37 are hand-authored
reference dimensions with no reachable model counterpart; they now report
`AuthoredDrawingEntityNoModelCounterpart` — a fact about the drawing, not a
defect in the ownership model. Reconciling R23's own imported annotations is
Phase 5+ and is unaffected.

Verification: 190 offline tests, same five stale R20 failures, 36 Phase 4
contracts. Preflight 24 components. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: Phase 4 fifth live run — reverse route matched nothing

**2026-08-01.** The reverse-correspondence route executed on all 38
annotations and reconciled none. `reconciled` is still 1 of 38 — the same
`ForwardAlias` match on `RD1@Drawing View7`. Read-only boundary held:
`mutations=0`, `finalSelectionCount=0`, `drawingUnchanged=True`.

**The run did not say why, and that is the defect being fixed here.** Two
causes were equally consistent with the log: `GetCorrespondingEntity2`
returning Nothing, or resolving to an entity no location owned. Rather than
predict, `MatchByReverseCorrespondence` now reports per attachment the
drawing entity type, whether the reverse call resolved, the trapped error
number, and the resolved model entity type — plus `resolved`,
`projectionsInView`, `modelEdgesTested` and `eqMax`.

`eqMax` carries the raw `swObjectEquality`. **`ISldWorks.IsSame` returns
0 NotSame / 1 Same / 2 Unsupported, and the `ObjectsAreSame` wrapper
collapses 0 and 2.** A cross-document comparison that cannot be performed
reads exactly like a non-match — a live suspect for a silent zero.

Two unverified observations to test against the next run: `Section View J-J`
is built on a section assembly (`P-0251-14A-001-SectionAssembly-3-1/...`)
whose cut edges may have no part counterpart at all; and several ordinate
attachments genuinely read as type `0` (`swSelNOTHING`), which is not valid
input to `GetCorrespondingEntity2`.

Verification: 189 offline tests, same five stale R20 failures, 35 Phase 4
contracts. Preflight 24 components. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 fourth run: reconciliation cause found, route added

**2026-08-01.** The instrumentation isolated the reconciliation failure in
one run, and it was a design gap, not a coding slip.

**The counterbore hole callout attaches to a drawing edge
(`attachmentTypes=1`, `swSelEDGES`) that is none of the 18 aliases
`IView.GetCorrespondingEntity` produced for that view.** The forward
model-to-drawing map is partial — it returned 2 of each location's 4 boundary
edges — so matching by forward alias could never find the callout's edge no
matter how many aliases were compared.

`Section View J-J` reported `anchoredProjections=0|aliasesAvailable=0`: the
forward map produces nothing there at all, so its seven dimensions,
including the H7, were structurally unreachable.

**R23-302 asked for the other direction and I had only built the forward
one.** `IModelDocExtension.GetCorrespondingEntity2` maps a drawing entity
back to its model entity, which the 2025 Help names for exactly this
purpose. `MatchByReverseCorrespondence` now takes each attached drawing
entity, resolves it to the model, and tests it against the geometry each
physical location owns — its `SourceFaces` and those faces' boundary edges.

The reverse route needs no projection anchor, so it can reach the section
view. Still identity only: `ISldWorks.IsSame` throughout, no positional or
dimensional fallback anywhere. Matches record `matchRoute=ForwardAlias` or
`ReverseCorrespondence` so the two are never conflated.

Read-only boundary held again: `mutations=0`, `drawingUnchanged=True`.

Verification: 187 offline tests, same five stale R20 failures, 33 Phase 4
contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 third live run

**2026-08-01.** Read-only boundary held again: `mutations=0`,
`finalSelectionCount=0`, `drawingUnchanged=True`.

### Tolerance policy is in force

Per the user's standing instruction, every tolerance now reports
`toleranceAuthority`. The H7 reads
`DrawingAuthoredNonAuthoritative`; all others `NoTolerance`. The dowel ±10 µm
rule is gone from the source entirely, and `ClassifyToleranceAuthority` has
no branch that could return a stronger authority. **Tolerances in the
designers' existing drawings are evidence that a designer typed a number, not
statements that the part holds them.** The rule for when a tolerance *should*
be added is open, pending the user's discussion with their designer.

### Reconciliation is still 1 of 38, and the alias fix did not help

Matching every drawing entity the projection owns — not just the anchor — was
predicted to raise this to 3-4. It did not move. Rather than guess again, the
evidence now names what each attachment actually is: `ANNOTATION_UNMATCHED`
reports `attachmentTypes` as `swSelectType_e` codes (1 edge, 2 face,
3 vertex, 10 sketch segment, 11 sketch point, 28 centre mark, 46 silhouette)
plus `anchoredProjections` and `aliasesAvailable`, so the next run says
whether the counterbore callout attaches to an unmapped edge, a sketch
segment, or something that is not an edge at all.

Only `RD1@Drawing View7` (the 4x tap-drill callout) reconciles. Everything
else in the drawing is authored as reference dimensions whose attachments
have not yet been shown to be hole edges.

Verification: 184 offline tests, same five stale R20 failures, 30 Phase 4
contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 second live run

**2026-08-01.** Both fixes confirmed, and the read-only boundary held again:
`mutations=0`, `finalSelectionCount=0`, `drawingUnchanged=True`.

**`IDimension.SystemValue` is definitively the right member for drawing
dimensions.** Every line now reads
`nominalSource=IDimension.SystemValue=<real value>|GetSystemValue3=0.000000000`
— all 30 display dimensions. Values match the reference drawing exactly:
47.0 H7, 11.0 counterbore callout, 4.2 tap-drill callout, R36, 173.6, 104.8,
80, 40, 25, 18, 12, 11.5, 6, and the ordinate chains 10/50/90/160 and
15/36 about zero. `attachments` is now 1-3 rather than 0.

The H7 reads in full: nominal `0.047000000`, `toleranceType=8`,
`tolMaxM=0.000025000`, `tolMinM=0.000000000`, `holeFit=H7`, status `0/0` —
the drawing's `47 H7 +0.025/0` exactly, and absent from the model.

### Standing instruction: drawing tolerances are not authoritative

The user directed that the tolerances in the designers' existing drawings be
ignored: they were added manually to signal that *some* tolerance is
acceptable, not that the part holds them. Every tolerance R23 reads is now
labelled `toleranceAuthority=DrawingAuthoredNonAuthoritative`, and
`ClassifyToleranceAuthority` has **no branch that can return a stronger
authority** — nothing R23 can currently read distinguishes a binding
tolerance from an indicative one. The dowel-specific ±10 µm rule is removed
rather than parked; the user is establishing with their designer what part
information should drive the decision to add a tolerance, and no rule is
guessed at meanwhile.

### Reconciliation defect fixed

Only 1 of 38 annotations reconciled. The cause: a counterbore maps **two**
edges, and the native hole callout attaches to the 11 mm mouth while the
anchor tier deliberately prefers the 6.6 mm through hole — different
entities of the same location. Matching now tests every drawing entity the
projection owns (`ProjectionOwnsEntity`), not just the chosen anchor.
Identity only; still no positional fallback.

Also caught before compiling: `alias` is a VBA reserved word.

Verification: 183 offline tests, same five stale R20 failures, 29 Phase 4
contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 first live run

**2026-08-01.** The read-only reconciliation run held its boundary exactly:
**`mutations=0`, `finalSelectionCount=0`, `drawingUnchanged=True`** — the
manual reference drawing was not altered. 38 annotations inventoried across
four views, `coverageFailures=None`.

### The H7 question is answered

`Section View J-J` carries `RD4@Drawing View6@P-0251-14A-001.Drawing` with
`toleranceType=8` (`swTolFITWITHTOL`), `nonZeroTolerance=True`,
`fitData=True`. **The Ø47 H7 fit is authored in the DRAWING, not the model** —
Phase 0 read `D1@Sketch4` directly and found no H7 there. R23 must never
expect model import to supply it, and must never record it with model
provenance.

The user also supplied the matching domain rule: **a dowel hole receives a
designer-added ±10 µm tolerance in the 2D drawing**, never in the model. Both
facts are the same pattern — precision tolerance is a drafting-stage act on
this drawing set. `MatchesDowelToleranceConvention` recognizes the ±0.010 mm
signature so provenance can follow from it, but recognizing the pattern is
deliberately *not* treated as proof that a hole is a dowel.

### Both hole callouts found

`RD3@Drawing View4` (the 6x counterbore) and `RD1@Drawing View7` (the 4x
tapped) both classified `NativeHoleCallout`, and the second reconciled to its
physical location by attached-entity identity. Ordinates: 10.

### Two defects fixed

- **`nominalM` was 0 for every dimension.** `GetSystemValue3` is
  configuration-scoped and these are *drawing* dimensions; a drawing document
  has no configurations. Now read through the configuration-free
  `IDimension.SystemValue`, with both readings kept visible.
- **`attachments=0` on every line**, including the annotation that went on to
  reconcile — the count was populated in a later pass than the line that
  printed it. Attachments are now read during inventory.

Also: the reference drawing authors its section diameters, H7 included, as
`swLinearDimension` (2), so dimension type alone can never decide whether
something is a diameter.

Verification: 181 offline tests with the same five stale R20 failures, 27 of
them Phase 4 contracts. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23 Phase 4 source-complete

**2026-08-01.** `Module14_AnnotationImport.bas` implements R23-400 to
R23-412: import constants, view-eligibility policy, the Phase 0 import
strategy, independent annotation traversal, category classification,
tolerance and fit reading, identity-based reconciliation against the location
graph, deduplication, and category-based coverage.

**This is the first phase that can modify a drawing.** Only
`ImportModelAnnotations` and `RemoveR23CreatedAnnotations` mutate, and both
refuse unless passed an explicit `allowMutation` argument.
`R23_ProbeAnnotationReconciliation` never passes it and contains no insert,
delete or save call, so **it is safe to run against the manual reference
drawing** — which is the point: that drawing already carries the
manufacturing intent R23 must reproduce, so reconciling against it exercises
R23-406 to R23-409, R23-411 and R23-412 on real data.

Verification is static: 176 offline tests with the same five stale R20
failures, 22 of them new Phase 4 contracts, plus a read-only production
preflight resolving 24 managed components. `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 3 gate SATISFIED

**2026-08-01.** R23-300 through R23-310 are proved live on P-0251 across five
runs. Final per-view acceptance:

| view | displayMode | projections | axisNormal | anchored | accepted |
|---|---|---|---|---|---|
| `Drawing View4` (primary) | HiddenLinesVisible | 11 | 7 | 11 | **7** |
| `Drawing View7` (side) | HiddenLinesVisible | 11 | 4 | 6 | **2** |
| `Section View J-J` | HiddenLinesRemoved | 11 | 4 | 0 | 0 |
| `Drawing View2` (isometric) | HiddenLinesRemoved | 11 | 0 | 9 | 0 |

`selectionProved=9`, `finalSelectionCount=0`, `selectionClean=True`,
`drawingUnchanged=True`, `partUnchanged=True`.

**R23-310's "four side projections" is unsatisfiable as written, and the
implementation is correct.** `Drawing View7` looks along model Y, so the four
tapped holes project onto **two** page points and the six counterbores onto
**three** — and the mapped counts match those numbers exactly. Two coaxial
holes seen along their axis are one circle on the sheet; SOLIDWORKS holds a
single drawing entity. Two side anchors is the complete answer. All four
tapped holes remain individually resolved in `Drawing View4`, and the
reference drawing calls them out once as `4x`.

The earlier hidden-lines hypothesis was **wrong**: `Drawing View7` is
`HiddenLinesVisible`, so the far-face holes are drawn. The cause is
projection coincidence, not visibility.

**Carry into Phase 5:** required coverage must be counted per distinct page
position per view, not per physical location. Requiring one annotation per
physical location in a side view cannot be satisfied by construction.

`MACRO_SOURCE_REVISION` remains `r22` — the engine has no pipeline caller and
drawing output is unchanged. Phases 4 onward are unstarted.

## Historical: R23-310 fourth live run, primary-view clause proved

**2026-08-01.** The axis gate works. Per-view acceptance:

| view | type | projections | axisNormal | anchored | accepted |
|---|---|---|---|---|---|
| `Drawing View4` (primary) | 4 | 11 | 7 | 11 | **7** |
| `Drawing View7` (side) | 4 | 11 | 4 | 6 | **2** |
| `Section View J-J` | 2 | 11 | 4 | 0 | 0 |
| `Drawing View2` (isometric) | 4 | 11 | 0 | 9 | 0 |

**R23-310's first clause is proved.** All six counterbores plus the stepped
bore accepted in the primary view, each anchored at
`anchorTier=PrimaryTypedHoleDiameter` on the Ø6.6 through hole rather than
the Ø11 mouth, all seven selectable with `ownershipProven=True`,
`selectionClean=True`, both documents unchanged.

The isometric is the gate's best corroboration: 9 anchors, `axisNormal=0`,
0 accepted — no hole axis is normal to an isometric sheet.

**Second clause is not met:** only two of four tapped holes have side
projections. The axis test is correct (all four are `axisNormal` in the side
and section views); the two far-face holes return Nothing from route A and
the section view maps nothing. `IView.GetDisplayMode2` is now recorded per
view, because under Hidden Lines Removed a far-side hole is never drawn and
cannot anchor however the search is written.

One defect fixed: `projection.ProjectedAxisX/Y/Z` were passed as ByRef
out-parameters. A class Public variable is a property, so the write-back was
discarded — `projectedAxis` logged as 0,0,0 while `axisNormal` was right.

## Historical: R23-310 third live run, anchors resolve and select

**2026-08-01.** The `GetCurveParams3` fix worked. `circularEdges` went 0 → 4
on every location with `uMin=0`, `uMax=6.283185307`, `closureM=0` — the
Phase 0 circle proof reproduced exactly. **26 projections anchored**, all via
route A with route C identity confirmation, and **all 26 proved selectable**
with `ownershipProven=True`, `selectedCount=1`, `finalSelectionCount=0`,
`selectionClean=True`, `drawingUnchanged=True`, `partUnchanged=True`.

`ISelectData.View` raised error 91 on all 26 attempts, exactly as this
repository predicted; the guarded binding plus after-the-fact ownership proof
absorbed it every time.

**One defect found: normal-axis compatibility was computed but not enforced.**
`AxisNormalToView` was stored and ignored by `QualificationFailureReason`, so
`Drawing View4` accepted all four M5 tapped holes even though that view's
normal is model Z and the M5 axis lies in the page plane — proved by their
page coordinates matching the model *Y* spread. Those anchors are edge-on,
not circles. They now fail as `AxisNotNormalToView`, kept distinct from
`ProjectionAnchorUnavailable` because the anchor itself is real.

Two questions settled by this run:

- Route A **does** work from a model reached through `IView.ReferencedDocument`
  with the drawing active. The earlier hypothesis was wrong.
- `Section View J-J` maps nothing at all — its component is a synthetic
  section assembly, so a section view will need its own acquisition route.

R23-310 stays open pending the fourth run, and whether "side projections for
all four tapped holes" is achievable is now in doubt: two of the four sit on
the far face and did not map. `MACRO_SOURCE_REVISION` remains `r22`.

## Historical: R23-310 second live run, root cause found

**2026-08-01.** The stage counters isolated the failure in one run. Every
location, in all four real views, reported `sourceFaces=2`,
`facesProjected=2`, `boundaryEdges=4`, `circularEdges=0`,
`firstReject=circle=Reject|reason=ReadError:438`.

**Root cause: `IEdge.GetCurveParams3` returns an `ICurveParamData` object,
not an array of doubles.** `Module13` assigned it into a Variant with a
`Let`, which asks for a default property the object does not have and raises
error 438, killing every candidate edge before any mapping was tried. The
correct pattern was already in this repository —
`Module12_FeatureQualification.ComputeFaceAxialInterval` binds it with `Set`
and reads `.StartPoint` / `.EndPoint` — and Module13 now matches it.

Two further defects were fixed alongside: an unmeasurable closure distance
would have failed open as a closed curve, and selection cleanliness was
reported without a baseline (the observed `finalSelectionCount=1` was the
operator's own pre-existing selection, not this pass).

Also established: the ten sheet placeholders report
`swDrawingViewTypes_e = 7`, and the visible-entity skip cut the run from 154
projections to 44.

R23-310 stays open pending the third run. `MACRO_SOURCE_REVISION` remains
`r22` — no pipeline caller, drawing output unchanged.

## Historical: R23-310 first live run, zero anchors, chain instrumented

**2026-08-01.** `R23_ProbeViewProjections` compiled and ran read-only against
the P-0251 drawing. `drawingUnchanged=True`, `partUnchanged=True`,
`finalSelectionCount=0`, 0 warnings, 0 failures — the safety envelope held.
**No projection resolved an anchor:** all 154 failed
`ProjectionAnchorUnavailable` with `candidates=0`.

Two facts are established:

- The drawing side works. `GetVisibleEntities2` returned 64 edges for
  `Drawing View4`, 68 for `Section View J-J` and 53 for `Drawing View7`,
  and each carried a drawing component (`componentContext=DrawingContextOnly`,
  as predicted for a part drawing).
- 132 of the 154 projections were against `*Left`, `*Bottom`, `*Current`,
  `*Isometric`, `*Dimetric` and `*Trimetric` — sheet standard-view
  placeholders returned by `ISheet.GetViews` with zero visible entities.

**The root cause is not yet known**, and the evidence line could not isolate
it: `candidates=0` cannot distinguish "no faces retained" from "no circular
edge found" from "nothing mapped". That is an instrumentation failure of the
same class Phase 0 solved with `rejectGate`. `ResolveProjection` now counts
`sourceFaces`, `facesProjected`, `boundaryEdges`, `circularEdges`,
`mappedEdges` and `inventoryConfirmed`, and carries the transform's own proof
string into `firstReject`. Views with no visible entities are now skipped
with one line instead of generating dead projections.

R23-310 stays open pending the rerun. No production behaviour is affected:
the engine still has no pipeline caller and `MACRO_SOURCE_REVISION` remains
`r22`.

## Historical: R23 Phase 3 source-complete

**2026-08-01.** `Module13_ProjectionResolution.bas` resolves each physical
location into a `CViewHoleProjection` for a drawing view: drawing-context
anchor, page-frame centre, and the proofs behind both. R23-300 through
R23-309 are implemented; **R23-310 needs one live run** of the read-only
`R23_ProbeViewProjections` to confirm usable primary projections for the six
face holes and side projections for the four tapped holes.

Verification is static: 142 offline tests with the same five stale R20
failures, 20 of them new Phase 3 contracts, plus a read-only production
preflight resolving 23 managed components. `MACRO_SOURCE_REVISION` remains
`r22` — the engine has no pipeline caller and drawing output is unchanged.

One contract worth carrying: **`ISldWorks.IsSame` is not a Boolean.** It
returns `swObjectEquality` = {0 not same, 1 same, 2 unable to determine}, so
reading it through `NormalizeSwBoolean` would accept "unable to determine" as
proof of identity. Only an exact 1 is accepted.

## Historical: R23-213 CLOSED, P-0251 catalog proved live, Phase 2 gate satisfied

**2026-08-01.** Two read-only `R23_ProbeFeatureCatalog` runs on P-0251. The
second returned **`catalogFailures=None`** with all four R23-213 expectations
met, 0 warnings, 0 failures and `modelUnchanged=True`:

- one six-location M6 counterbore family (radii `0.0055`/`0.0033`, two unique
  X and three unique Y, matching the Phase 0 ordinate evidence);
- one four-location M5x0.8 tapped family;
- one Ø47/Ø40 stepped-bore stack (radii `0.0235`/`0.0200`, two members);
- nothing spurious accepted — 40 rejections, each with an explicit reason.

18 qualifying cylindrical faces consolidated to 11 physical locations through
7 merges. The axial interval earned its place: two coaxial pairs of M5 tap
drills that share a line key were correctly held apart as four distinct
locations by their disjoint intervals, which the Plücker key alone cannot do.

The first run reported `NoFourLocationFamily`, which was real, and the log
exposed two further defects. Three fixes: pattern instances now inherit their
seed's semantics; the seed-chain rejection codes are enforced rather than
merely computed; and `HoleFit` no longer publishes the not-applicable `-1` as
a fit value. See [Changelog.md](Changelog.md).

### Open follow-up: traversal is not exact-once

Comparing the two runs, `visitedFeatures` went **47 → 46** on an unchanged
model, and the sketches visited twice differed between runs. The traversal
key included `ObjPtr`, so one feature reached through two COM wrappers split
into two keys. Fixed by keying on name plus type, with the address kept only
as an unnamed-feature fallback.

This did not affect either catalog — accepted features, locations,
consolidations and families were identical across both runs — but it is not
cosmetic: a doubled visit to an *accepted* feature would add duplicate stack
members and misreport `stackMembers`. **The fix is static-only.** It needs no
dedicated run and can be confirmed by `visitedFeatures` being stable and
duplicate-free on the next live run of any Phase 3 work.

## Historical: R23 Phase 2 source-complete

**2026-07-31.** `Module12_FeatureQualification.bas` implements the Phase 2
feature-qualification engine: `ICE` type normalization, cycle-guarded
traversal of features and subfeatures, suppression proved against the exact
referenced configuration, typed readers for six feature families with paired
selection access and release, ownership from `IFeature.GetFaces`, and
physical locations built from owned cylindrical faces with axial intervals
measured from real boundary edges.

The engine has no pipeline caller. Drawing output is unchanged and
`MACRO_SOURCE_REVISION` remains `target-spec-hybrid-v2-2026-07-29-r22`.
The manifest manages 22 components.

The user deployed the 22 managed components into `Fable.swp` and ran
**Debug > Compile Project**, which stopped at `evidence.InfoCount` in
`R23_ProbeFeatureCatalog` with "Method or data member not found" —
`CRunEvidence` exposes no accessor over its Private `mInfo` collection. The
redundant replay loop was removed in favour of a `WarningCount`/`FailureCount`
tally; every other cross-module reference and call-site arity in the seven new
Phase 1/2 components was then audited against its target and found sound.
**Phase 1/2 source has not yet passed a full VBE compile**; the corrected
source needs redeploying and recompiling.

Verification is otherwise static: 117 offline tests with the same five stale
R20 failures, 23 of them new Phase 2 contracts, plus a read-only production
preflight.

**The one open Phase 2 item is R23-213**, which needs a live run of the
read-only `R23_ProbeFeatureCatalog` against P-0251 to confirm one
six-location counterbore family, one four-location tapped family, one
stepped-bore stack, and nothing spurious accepted.

## Historical: R23 Phase 1 source-complete

**2026-07-31.** The location-graph model is added under
`src/target-spec-hybrid-v2/`: `CFeatureDefinition`, `CPhysicalHoleLocation`,
`CViewHoleProjection`, `CImportedAnnotation`, `CLocationGraph` and
`Module11_GeometryIdentity`. The deployment manifest now manages 21
components.

The change is additive by design (R23-107): the new classes have no callers,
a test asserts no existing production module references them, and runtime
output is identical to r22. `MACRO_SOURCE_REVISION` therefore remains
`target-spec-hybrid-v2-2026-07-29-r22`, and nothing has been deployed.

Physical identity is an infinite axis line — sign-normalized direction plus
line moment — together with the axial interval the material occupies.
Consolidation happens only in
`CLocationGraph.ResolveOrCreatePhysicalLocation`, which requires both the
same line and meeting intervals. That is what merges a counterbore with its
through hole while keeping opposite blind holes on one axis separate, and
neither key contains a feature name.

Verification is static only: 94 offline tests with the same five stale R20
failures, source hygiene clean on all six new files, and a read-only
production preflight resolving 21 managed components. No VBE compilation and
no live run.

## Historical: R23 Phase 0 closed; production Phase 1 unblocked

**2026-07-31.** The corrected disposable probes closed every Phase 0 gate on
the authorized P-0251 fixture. Both datum-first ordinate groups returned
`AddOrdinateDimension = 0 Success` with exact selection counts, correct
display-dimension deltas, `SetPickMode`, zero-selection cleanup, and an
unchanged fixture. Ordinate values are `+15.00`/`-15.00` about the
stepped-bore centre and `10.00`/`50.00`/`90.00` from the bottom-left vertex
datum — the reference scheme.

Contracts carried into production:

- entity correspondence uses route A,
  `IView.GetCorrespondingEntity(modelEdge)`; `IComponent2` mediation returns
  Nothing on this build;
- mapping is per-edge, so every feature-owned edge must be attempted;
- never assign `ISelectData.View` (runtime error 91); activate the view and
  verify with `ISelectionMgr.GetSelectedObjectsDrawingView2`;
- normalize every SOLIDWORKS COM Boolean with `(CDbl(raw) <> 0#)` before any
  negation or compound logic;
- accept ordinate `Type2` values `1`, `7` and `8`;
- `ICurve.CircleParams` is available and returns correct radii; and
- section-line payload segments are view-sketch coordinates requiring exactly
  one conversion before page-frame comparison.

The three J-J top-border violations are production work for R23-704, now
measured truthfully rather than unknown. Production R23 source, `Fable.swp`,
fixtures, the protected baseline and manual references remain unchanged, and
no production implementation has begun.

## Historical: R23 implementation at the Phase 0 live gate

The approved R23 plan is now an active implementation checklist. Pre-change
provenance is retained under:

`test_assets/iteration_evidence/r23/20260730-075811/prechange/`

Confirmed offline:

- production `Fable.swp` and its recoverable backup have matching SHA-256
  `C8076D713DD8F64AC75F93871C1EA4A2D1F01EE2BC691323736013BFEB2803F2`;
- embedded and exported R22 source match for all 15 managed components;
- the guarded deployment preflight passes with the bootstrap present;
- the installed SOLIDWORKS 2025 SP1.2 interop remains version `33.1.2.4`;
- the local API MCP confirms that `GetTypeName2="ICE"` must be normalized
  through `GetTypeName`, and that successful typed-definition selection access
  must be released; and
- the standalone, read-only feature/curve probe is prepared at
  `tools/r23-probes/Module_R23Phase0FeatureProbe.bas`;
- a disposable copy of `Fable.swp` plus a guarded three-module overlay is
  prepared for the expanded import-transaction comparison; and
- the disposable import manifest passed read-only preflight with all three
  overlay sources present and the deployment bootstrap installed.

The first run against the authorized workspace P-0251 fixture completed, but
its transcript was rejected as incomplete probe evidence. A recursion guard
based only on `ObjPtr(feature)` collapsed distinct COM wrappers that reused an
address, so only 15 features were visited and the manufacturing-relevant ICE,
Hole Wizard, extrusion, and mirror records were skipped. The rejected
transcript is retained under the R23 live-probes evidence folder.

The exported probe has been corrected to use a composite diagnostic recursion
key, accept both scalar and array `IsSuppressed2` returns, and include base
`EXTRUSION` definitions in the typed extrusion probe. It now records
contour/profile state and mirrors every evidence record to a timestamped log
so Immediate Window truncation cannot invalidate the rerun. Static source
structure, continuation, encoding, and deployment-preflight checks pass. The
corrected feature source has now received full-project VBA compilation and a
complete user-operated run. Its retained log is
`R23_FEATURE_20260731_040539.log`: it visited 47 features, resolved all three
`ICE` entries to `Cut`, proved typed definition access/release and owned
geometry, read both native Hole Wizard definitions, and identified `Mirror1`'s
M5 seed feature. `IsSuppressed2` returned Empty for the active configuration
on every feature, so the explicitly labelled current-configuration fallback
reported active state. Both `IsCircle`/`GetCurveParams3` call orders remained
stable; the anomalous `CircleParams=SkippedNotCircle` result means production
must not rely on `CircleParams`. The disposable import overlay and final
drawing-contract extension have since received user-operated runtime execution;
their findings are summarized below.

Production R23 VBA has not been written or deployed. The approved plan
explicitly prohibits crossing into production feature classification until the
installed SW2025 build proves the three P-0251 underlying `ICE` types, typed
definition access, owned geometry, curve read order, expanded import
transaction, explicit ordinate transaction, and representative section
dimension/geometry behavior.

The user performs the remaining live SOLIDWORKS work. Both import variants
completed in the disposable SWP without changing the fixture:
`AllViews=True` returned 25 dimensions but left the side view empty, whereas
the explicit section/side/primary calls returned the same 25 as 17 section,
2 side, and 6 primary dimensions. The isometric stayed clean in both. R23 will
use deterministic `AllViews=False`, in section/side/primary order, with
`DuplicateDims=True` and the expanded mask containing
`swInsertDimensions = 8`. The logs prove one imported M5 native callout in the
section, but no imported M6 counterbore callout and no H7/nonzero tolerance.

The two final drawing-contract probes also ran, but they did not close Phase 0:

- `R23_ProbeDatumFirstXYOrdinates` found the visible component and then stopped
  fail-closed with `CounterboreLocationsUnavailable`. It never selected a
  datum, never called `AddOrdinateDimension`, and therefore proves nothing
  about the ordinate API. Cleanup completed with zero selections and the
  fixture remained unsaved. The disposable probe must add per-face/per-edge
  qualification and direct/component/visible-entity mapping diagnostics before
  a corrected rerun.
- `R23_ProbeSectionDimensionsAndJJGeometry` found the exact imported 47 mm and
  40 mm dimensions as real `swDiameterDimension = 6` items. Neither carried H7
  or a nonzero tolerance. The probe's per-dimension target fields were not
  reset on every loop iteration, so later labels contain stale state and must
  be corrected before they can be used as acceptance evidence.
- The J-J payload is structurally complete at 49 returned items, but it mixes
  source-view sketch coordinates for segments with page coordinates for arrows
  and labels. The current production clearance comparison therefore compares
  unlike coordinate frames. The retained screenshot also confirms the upper
  arrow/label enters the top border/zone band and the lower arrow/label enters
  the part-identification band.

Production R23 remains blocked. The next permitted implementation work is only
to correct the disposable probes, run the direct part-source H7 readback, and
close the entity-mapping and page-coordinate contracts. The complete handoff
and exact correction checklist are in
`docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md` and
`docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md`.

### Corrected-probe run result (2026-07-31, build 20260731.2)

The user compiled and ran both corrected entry points on P-0251. Two of the
three gates closed.

- **Section dimensions: closed.** Exactly one `DIAMETER_47`, one
  `DIAMETER_40`, one `FIRST_LINEAR`, no stale labels, both diameters
  `type2=6`.
- **H7 authority: closed.** The direct part-source readback proved
  `D1@Sketch4` carries `toleranceType=0`, `fitType=-1`, and no fit strings,
  so H7 is not in the model. The user selected controlled
  target-spec/reference authority: R23 will apply `H7 +0.025/0.000` to an
  associative `Ø47` dimension with provenance recorded as target-spec
  authority, never as model data, and QA must state that explicitly.
- **J-J frames: closed.** All six returned segment endpoints matched the
  captured `CreateLine` view-sketch inputs at `deltaM=0`, and the single
  inverse conversion is cross-validated against the independently returned
  page-frame arrow endpoints.
- **J-J clearance: measured.** Three violations, all at the top border
  (segment 1 start, upper arrow, upper label). The lower arrow and label
  **clear** the measured part-identification note extent by about 7.7 mm,
  correcting the earlier screenshot-derived claim of a lower-band intrusion.
- **Ordinate ownership: proved.** After the `CBool` normalization fix,
  twelve of the eighteen owned counterbore faces read as cylinders with
  correct radii, all passed the page transform with outline containment, and
  the six page centres form exactly the required grid — two unique X and
  three unique Y, spaced 30 mm and 40 mm.
- **Entity correspondence: settled.** Route A
  (`IView.GetCorrespondingEntity(modelEdge)`) works — 12 of 24 counterbore
  edges and 114 of 154 vertices mapped. Route B
  (`IComponent2.GetCorrespondingEntity`) returned `Nothing` every time and is
  unusable here. Route C confirms route A returns genuine drawing entities by
  `ISldWorks.IsSame` identity. Mapping is per-edge, so every owned edge must
  be attempted. The full scheme resolved: six unique locations, two unique X,
  three unique Y, a stepped-bore X datum and a bottom-left vertex Y datum.
- **`ICurve.CircleParams` works**, returning radii matching the owning
  cylinders exactly. The R23-006 exclusion was an artifact of the guard
  defect and is closed.
- **`ISelectData.View` cannot be assigned on this build.** Both ordinate
  groups failed with runtime error 91 before any selection;
  `CreateSelectData` returned a live object, so the failing statement is the
  documented get/set assignment. This reproduces the behaviour already
  recorded in `Module2_DrawingPipeline.CreatePrimarySection`. The probe now
  guards the binding, records its outcome, and proves each selection's owning
  view through `ISelectionMgr.GetSelectedObjectsDrawingView2`.
  `AddOrdinateDimension` has still never been called.
- **Historical note — superseded.** The
  second and third runs both stopped at the edge-closure gate. The
  `rejectGate` instrumentation isolated the mechanism: one call logged both
  `isCircle=True` and `rejectGate=IsCircleFalse`, because `CStr` renders the
  value `True` while `If Not value` yields `-2`, which VBA treats as True.
  `CBool` is not a dependable normalization — it worked for
  `ISurface.IsCylinder` but not for `ICurve.IsCircle`. The probe now uses
  `NormalizeSwBoolean`, an explicit `(CDbl(rawValue) <> 0#)` comparison,
  across `IsCircle`, `IsCylinder`, `ActivateView`, both `Select4` calls and
  `GetSaveFlag`. Mapping routes A, B and C have still never executed, so
  `IView.GetCorrespondingEntity` behaviour on this scaffold remains unknown
  and one further probe rerun is required.
- **`CircleParams` exclusion withdrawn.** The `SkippedNotCircle` result that
  barred it from production came from the same `If Not <SOLIDWORKS Boolean>`
  defect in the feature probe's `ReadCircleState`; the API was never called.
  Its behaviour is untested rather than anomalous. Production must still not
  depend on it until a run exercises it.

The disposable-probe corrections are now implemented at probe build
`20260731.2-mapping-frame-h7-contracts` in
`tools/r23-probes/import-transaction-source/`: per-face/per-edge ordinate
qualification with direct, component-mediated, and visible-entity mapping
comparison; per-iteration section-state reset with exact type-6 47/40 target
enforcement; direct part-source `D1@Sketch4`/`D1@Sketch6` tolerance readback
feeding an explicit H7 authority record; and frame-proved J-J geometry with a
single sketch-to-page conversion and page-frame clearance verdicts against
the content border, title block, measured part-identification note extent,
and view outlines. This state is static verification plus read-only
disposable preflight only: the corrected probes still require user-operated
deployment, full-project VBE compilation, execution on the authorized P-0251
fixture, and returned logs/screenshots before any Phase 0 gate can close.

The full offline suite remains at its known baseline of 69 passes and five
stale R20-contract failures.

The protected baseline, fixture models, manual references, production VBA
source, and production SWP remain unchanged.

## R22 corrected review line is awaiting user compile and runtime evidence

The authoritative managed source identifies as
`target-spec-hybrid-v2-2026-07-29-r22`.

R22 combines all source changes from the r20/r21 commit line, retains the
review fixes that satisfy the verified contracts, and corrects the latest
commit's three substantive defects:

- non-circular cylinder trims now fail closed at `ICurve.IsCircle=False`
  instead of being assigned a centre from `ISurface.CylinderParams`;
- pattern ownership uses the exact SOLIDWORKS 2025 `GetTypeName2` literals,
  including `APattern`, `LocalChainPattern`, `DimPattern`,
  `DerivedHolePattern`, `SketchPattern`, and `LocalSketchPattern`; and
- the arrange fallback delegates ordinate-set jogs to
  `IDisplayDimension.AutoJogOrdinate`, using manual deterministic lanes only
  for non-ordinate dimensions.

The local MCP contracts and installed SOLIDWORKS 2025 SP1.2 interop
`33.1.2.4` confirm the used member signatures and dimension enum values. The
retained P-0251 curve probe confirms that all relevant owned model edges already
provide `IsCircle=True` and valid `CircleParams`, so removing the unsafe
cylinder-trim inference does not discard those valid candidates.

The guarded r22 deployment and exact managed-source readback are the next
offline gates. After deployment, the user must run VBA Editor **Debug > Compile
Project**, then run only an authorized fixture and return the complete
Immediate Window output, `QA_REPORT.txt`, and one uncropped full-sheet
screenshot. Static checks, interop reflection, and the deployment bootstrap
probe are not full VBA compilation or manufacturing acceptance.

## Historical R20 compile hotfix and runtime baseline

The authoritative managed source and embedded macro now identify as
`target-spec-hybrid-v2-2026-07-28-r20`.

The user's full-project VBA compile exposed `Argument not optional` at
`Module6_QAEngine.CheckSectionLineClearance` because the r20 source called
`IView.GetSectionLineCount2` without its mandatory `ByRef Size As Long`
argument. The project `solidworks-api` MCP and the installed SOLIDWORKS 2025
interop both confirm the exact signature
`GetSectionLineCount2(ByRef Size As Long) As Long`; `GetSectionLineInfo2`
remains parameterless. The exported source now supplies the size argument,
rejects invalid/nonmatching array sizes, and has a regression assertion against
the zero-argument call.

After the user closed the locked VBA project, the guarded deployment completed
under `test_assets/iteration_evidence/swp_deployment/20260728_142300/`.
Candidate and promoted readbacks match all 15 managed components and the exact
r20 revision. `compile-probe-scope.txt` explicitly records that the automated
probe covers bootstrap execution only.

R20 is based on the r19 P-0251 evidence at:

- `test_assets/iteration_evidence/macro_qa/20260728_091302_P-0251-14A-001/QA_REPORT.txt`

R19 preserved four created views, ten imported display dimensions, a clean
isometric, the J-J structure, metric mass `1.30`, title properties, general
notes, part identification, and complete QA output. It nevertheless failed
because all 32 complete internal-cylinder boundary edges in each supported view
stopped at `ClosedCircleIsCircleFalse`, leaving zero mapped circular edges,
ownership candidates, canonical physical locations, or ordinate groups. The
side/section gap sat exactly on the collision threshold; the lower J-J marker
entered the part-identification band; annotation QA incorrectly excluded the
entire lower sheet below the view-placement boundary; and auto-arrange did not
prove any view-scoped selection.

R20 repairs those source defects:

- complete boundary edges are proved with `IEdge.GetCurveParams3`; when a
  trimmed edge does not identify as an `ICurve` circle, its already
  ownership-proven internal face supplies the cylinder center, axis, and radius
  through `ISurface.CylinderParams`;
- used SOLIDWORKS Boolean results are normalized before negation/compound logic;
- `ModelToViewTransform` page coordinates are no longer translated by
  `IView.Position` a second time, and the center datum compares against the
  transformed model origin. Every candidate centre, projected origin, and
  mapped vertex now fails closed unless the transformed point lies within the
  current `IView.GetOutline`, with explicit page-frame evidence;
- the common view gap is 12 mm, layout comparison has a numeric tolerance,
  P-0251's orthographic/section row is biased upward, and all requested centers,
  actual outlines, and pair clearances are logged separately for initial and
  final layout;
- annotation QA records annotation type/name/position and validates against the
  real zoned border and rectangular title block. `UsableBottom` remains a view
  placement boundary only. The measured part-ID extent is retained and
  `GetSectionLineInfo2` segment, arrow, and J-label geometry must clear it;
- auto-arrange uses `ISelectData.View`, checks `IAnnotation.Select3`, verifies
  selection count, and logs the `AlignDimensions` result. It is now a required
  stage; `False` invokes deterministic 6 mm `SetPosition2` lanes with exact
  readback and content-border proof; and
- three controlled P-0251 callouts now define the stepped bore, six
  counterbored face holes, and four tapped side holes. Each callout uses the
  documented `<MOD-DIAM>` syntax, selects an ownership-proven drawing edge
  before `InsertNote`, and must read back nonzero attachments, a visible leader,
  safe extent, and clearance from other note extents/annotation origins.
  Existing-note reuse requires the complete normalized controlled definition,
  so a shorter imported Hole Wizard phrase cannot falsely prove the stage.

The project-local suite passes **74 tests and 13,608 structural subtests**.
The current guarded deployment proves a 15/15 managed-source match and exact
r20 revision. Its `COMPILE_PROBE|status=SUCCESS` record proves only that the
deployment bootstrap could execute; it does not perform VBA Editor **Compile
Project** and must not be cited as full-project compilation.

The required next sequence is: open `Fable.swp`, run VBA Editor **Debug >
Compile Project**, and—only if compilation succeeds—run P-0251. Runtime evidence
remains the new `QA_REPORT.txt`, complete Immediate Window output, and one
uncropped full-sheet screenshot. Computer Use was not used and remains
disallowed unless the user explicitly requests it.

## Historical R18 model-first drawing-output repair

The guarded r17 deployment compiled, synchronized all 15 managed components,
and executed on authorized fixture P-0251. Its retained evidence is:

- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/QA_REPORT.txt`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/diagnostic-drawing.png`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/R17_VISUAL_AND_QA_DIAGNOSIS.md`
- `test_assets/iteration_evidence/macro_qa/20260726_163134_P-0251-14A-001/R18_API_AND_REPAIR_CHECKPOINT.md`

R17 preserved the runtime-proven J-J transaction and imported nine display
dimensions into Drawing View1 plus one into Drawing View2. The section and
isometric immediate per-view arrays were zero. It created the isometric before
import, executed eight view moves, and produced a visibly coherent four-view
layout. The metric mass calculation also returned `1.296824 kg` and wrote the
drawing property `Mass=1.30`.

R17 still failed manufacturing acceptance:

- all 64 + 47 drawing-to-model correspondence calls returned `Nothing`;
- accepted candidates, canonical locations, and ordinate groups stayed zero;
- final QA falsely reported 39 display dimensions and 10 on the isometric
  because its `GetNext5` iteration crossed the drawing sheet;
- the visible mass stayed `1296.82`, proving its note did not use the corrected
  drawing property;
- invalid note extents generated repeated diagnostic layout failures; and
- the controlled sheet still has no structural `ITitleBlock`.

The managed source is now
`target-spec-hybrid-v2-2026-07-18-r18`. R18 starts with the model audit, walks
owned internal cylindrical faces and circular model edges, and maps each known
model edge into the target view through `IView.GetCorrespondingEntity`.
Selectable datum vertices use the same model-first direction. Exact per-view QA
uses `GetDisplayDimensions`. A visible mass note is changed to `$PRP:"Mass"`
only when a single linked mass candidate exists and link/rendered-value readback
passes. Diagnostic note skips are allowed only while controlled boundaries are
already unproved and explicitly record `acceptance=False`.

`AddHoleCallout2` was not added as a fallback because the documented API
requires a user confirmation dialog. Hole callouts remain on the verified model
annotation path; model marking or another approved non-modal strategy is needed
for missing grouped callouts.

The missing structural `ITitleBlock`, approved exact title links, controlled
regions, grouped callout completeness, and stepped-bore manufacturing definition
remain deliberately fail closed. R18 is E2/E3 until embedded compilation and an
authorized synchronized run prove the new mapping and drawing result. The
complete project-local suite passes 66 tests, all 15 managed components pass the
Windows-1252/CRLF/no-BOM/no-metadata gate, and the guarded read-only deployment
preflight resolves r18 with the bootstrap present.

## Earlier source-completion and deployment outcome

The coherent replacement VBA source is written under
`src/target-spec-hybrid-v2/`. The current source identity is
`target-spec-hybrid-v2-2026-07-18-r18`; r14 removed the disproved sheet-edit
preflight, verified active drawing-view context with a selection-assisted retry,
packed diagnostic views in two rows, and used a standard independent 1:2 scale
for the P-0251 orientation aid. R15 adds the first output-driven corrections
and r16-r18 continue the evidence-driven repairs summarized above.

A guarded deployment tool now exists under `tools/swp-deploy/`. Its read-only
preflight confirms that `Fable.swp` contains 19 components including the
deployment bootstrap, while exported production source is r14. The tool manages
15 replaceable modules/classes through a disposable candidate project, compile
probe, source readback, and atomic promotion. The first candidate import exposed
UTF-8 BOM bytes plus VBA export metadata as uncompilable visible code. All
deployable `.bas` and ordinary `.cls` inputs are now Windows-1252 without a BOM
or `Attribute` records; the bootstrap assigns imported component names itself.
The first metadata-free candidate revealed that `VBComponents.Import` recreated
ordinary handler `.cls` files as standard modules, making `WithEvents` illegal.
The bootstrap now creates manifest `StdModule` and `ClassModule` components with
explicit VBIDE type values and injects the cleaned source with
`CodeModule.AddFromString`.

On 2026-07-26, the active-object proxy returned from the Running Object Table
rejected the installed interop's `ISldWorks` IID with `E_NOINTERFACE` before
Module0 could execute. The compiled invoker now retains the early-bound path but
falls back to `IDispatch` with an explicit by-reference error argument. This
fallback is live-verified.

The next live attempt reached Module0 but failed with VBA error 748 because
`VBProject.SaveAs` is not valid for the SOLIDWORKS host-managed project type.
Module0 now invokes the VBE Save command for the already-named candidate, and
PowerShell creates the separate output copy afterward. Failure evidence now
records the current stage/component.

The guarded deployment completed successfully on 2026-07-26. The compile probe
reported success, all 15 managed components matched the source, and both
candidate and post-promotion verification read back revision
`target-spec-hybrid-v2-2026-07-18-r14`. The promoted `Fable.swp` SHA-256 equals
the verified candidate SHA-256. The previous macro, verified candidate, compile
result, and source-verification evidence are retained under
`test_assets/iteration_evidence/swp_deployment/20260726_100426/`.

The r4 package contains nine standard modules, three domain/evidence classes,
three handler classes, two UserForm code snapshots, the `ThisLibrary` host-code
snapshot, and an import guide. It remains separate from both the protected
baseline and `src/active-ordinate/active_ordinate.swp`.

`Fable.swp` is now synchronized with the r14 managed source. No authorized
model, protected baseline, or manual reference drawing was changed.

## Implemented in the r4 export

- strict authorization and fixed hybrid behavior for the three fixtures;
- fixture-locked view, section, detail, datum, scale, title, note, part-ID, and
  QA acceptance profiles;
- reference-led orthographic plans without unconditional rotation;
- one deterministic primary section: P-0251 J-J, no Base Plate section, and
  Pump Holder B-B;
- mandatory Pump Holder Details C and D from the exact `*Bottom` view at an
  independent 3:1 scale, with COM-identity, selected-profile, circular-geometry,
  source/detail ownership, scale-ratio, placement, and outline readback;
- fixed model-annotation import followed by evidence-backed ordinate fallback;
- component-qualified visible-entity enumeration, matched-face feature
  ownership, configuration proof, physical-instance identity, and conservative
  rejection of unproved circles;
- typed direction-specific datum proof and family/view/datum/direction-scoped
  coverage reconciliation;
- datum-first ordinate transactions with checked selection counts, decoded
  return codes, cleanup, and physical/projection evidence ledgers;
- controlled-sheet and actual-scale readback, standard-scale layout, linked
  title-property evidence, read-only final QA, final cleanup proof, and atomic
  evidence writing; and
- deterministic import guidance for form code snapshots and the special
  `ThisLibrary` host component.

## Verified offline evidence

- The complete project-local companion suite passes **49 tests**.
- Installed SOLIDWORKS 2025 interop file version `33.1.2.4` was reflected for
  the used interfaces and enums.
- Primary annotation, visible-entity, corresponding-entity, feature,
  configuration, transform, ordinate, sheet, title-block, property, section,
  and detail-view contracts were checked at E3.
- Detail-view reflection confirms `CreateDetailViewAt4`, exact enum values,
  `UseParentScale`, `ScaleRatio`, `IView.GetDetail`, `IDetailCircle` ownership,
  profile, style/display, and outline members.
- Reflection reconfirmed `swCreateOrdDimErr_GenFailure = 1` and
  `swCreateOrdDimErr_OrdFailure = 7`.
- Structural tests cover balanced procedures, continuation limits, component
  inventory, caller/signature contracts, and duplicate local/parameter names.
- The r4 checkpoint contains a 35-entry SHA-256 manifest and source archive at
  `test_assets/iteration_evidence/2026-07-18_target_spec_hybrid_v2_r4_offline/`.
  The older similarly named checkpoint without `_r4_` is historical r3/28-test
  evidence and must not be used to identify the current source.

The user has embedded and compiled the r4 source successfully in the
SOLIDWORKS VBA editor. Its first run reached macro preflight but performed zero
SOLIDWORKS mutations because the template constant omitted the `VEEMAP`
directory level. R5 corrected that path and reached sheet measurement. R6 fixed
the `ISheet.GetSize` width/height output binding. R7 attempted to normalize
`SheetFormatVisible`; R8 records that flag as a warning and relies on structural
title-block/margin/usable-area proof. R9 permits diagnostic continuation after
those checks fail; R10 also continues after failed scale setting using the
template's existing scale. The user's r10 run created the primary view and
visually showed the requested hidden-lines-visible mode, but `SetDisplayMode4`
returned `False`. R11 replaces that setter-only gate with setter-plus-readback
verification. The r11 run then created both orthographic views but stopped on a
setter-only rebuild check. R12 permits independent diagnostic stages to continue
and provides a clearly non-acceptance layout reserve when `ITitleBlock` is not
defined. Production acceptance remains blocked until the controlled sheet
contract is proven. The r12 run reached all independent stages and exposed one
shared view-context failure blocking section, model annotation, and ordinate
operations. R13's attempted sheet-edit normalization was disproved by its next
run: it stopped before sheet measurement or view creation. R14 removes that
regression and limits recovery to a named `DRAWINGVIEW` selection plus one
`ActivateView` retry, accepted only when `ActiveDrawingView` matches.

## Remaining gates

1. Run synchronized r22 on P-0251 and retain its complete E6 runtime and drawing
   evidence.
2. Confirm nonzero complete circular boundaries/mapped edges, ten canonical
   P-0251 physical locations, and the required X/Y ordinate coverage.
3. Prove the model-first Center-X and Bottom-Y datum entities are selectable in
   the intended drawing views and that every ordinate transaction cleans up.
4. Confirm imported dimension counts remain `9, 1, 0, 0` before new ordinates,
   the isometric stays undimensioned, and the front-view multi-dimension
   arrange stage proves either `AlignDimensions` or its deterministic lane
   fallback with exact position readback.
5. Confirm all three manufacturing callouts render the required tokens and
   SOLIDWORKS diameter symbols, remain attached to the intended bore/face/side
   geometry with leaders, and avoid other annotations, view outlines, part ID,
   and title block. Confirm parsed J-J segment/arrow/label geometry clears the
   measured part-ID extent.
6. Confirm annotation-origin diagnostics no longer reject legal lower-left
   content and identify any genuine remaining border/title/leader violation by
   type, name, and coordinates.
7. Confirm the unique mass-note link reads `$PRP:"Mass"` and visibly renders
   `1.30`.
8. Run only the three authorized fixtures and retain the complete E6 regression
   matrix.
9. Compare every output with its manual reference and pass E7 manufacturing,
   coverage, and layout acceptance.

## Controlled-template status

The controlled drawing template exists at:

- `V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`

The r4 constant incorrectly omitted the `VEEMAP` directory level, which caused
the truthful fail-closed first-run result. R5 corrects the fixed path. The
template-linked sheet-format and title/property contract still require live
evidence from the next authorized run.

`GetValidDrawingTemplatePath` performs one fixed-path existence check; it is not
template discovery. Module7's current property/link/cell map is an unapproved
D-04 candidate, not evidence that the controlled title contract is settled.

## Collaboration operating mode

Codex owns offline source work, API validation, output diagnosis, reference
comparison, tests, and automatic synchronization through the guarded
PowerShell deployer. The user owns the actual SOLIDWORKS macro run and shares
the resulting QA output and screenshot.

Computer Use must not be used for SOLIDWORKS unless the user explicitly asks
for it. A newly written revision is deployed automatically through
`tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`, after which the user is asked to
run the macro.

## Exact handoff point

The coherent r22 managed source is embedded in `Fable.swp`, bootstrap-probed,
and verified 15/15 against exported source. The user should now compile and run r22 on
`P-0251-14A-001.SLDPRT` with the same acceptance-profile settings and share the
new QA report, complete Immediate Window output, and one uncropped full-sheet
screenshot. No Computer Use is required.
