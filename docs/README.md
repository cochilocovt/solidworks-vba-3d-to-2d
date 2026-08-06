# docs index

Twenty-five documents, most of them written for source trees that are now in
`archive/`. This index says which are current. **A document's own revision
label is not evidence** — `Agents.md` settles precedence: current source wins
over any doc, export, screenshot or transcript.

Trunk: `src/baseline-model-dims/`. Read `MACRO_SOURCE_REVISION` in
`Module1_Main.bas` for the real revision.

## Current — read these

| Doc | What it is |
|---|---|
| [CURRENT_STATUS.md](CURRENT_STATUS.md) | Where the macro is now, what is open, what is waiting on a user decision. Newest section first. |
| [Architecture.md](Architecture.md) | Component map, execution order, the ordinate engine step by step, and why the form owns the settings. |
| [SOLIDWORKS_API_VALIDATION.md](SOLIDWORKS_API_VALIDATION.md) | Accumulated API-contract evidence. Append-only, newest at the bottom. The provenance record every `sw*` constant must appear in. |
| [BASELINE_TO_REFERENCE_DRAWING_GAP.md](BASELINE_TO_REFERENCE_DRAWING_GAP.md) | Per-gap status against the reference drawing, the phased plan, and section 12 on what the live runs taught. |
| [Changelog.md](Changelog.md) | One entry per revision, newest first. |
| [R23_SCOPE_AND_GENERALIZATION_PLANNING.md](R23_SCOPE_AND_GENERALIZATION_PLANNING.md) | The three tiers of "done" and the recorded product decisions (section 4a). Objective is Tier C. |
| [3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md](3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md) | Durable implementation guidance. |
| [CODESTACK_DRAWING_API_COVERAGE.md](CODESTACK_DRAWING_API_COVERAGE.md) | 33-row ledger of the CodeStack drawing corpus. Check it for a tested pattern **before** designing one; it also states plainly where the corpus has nothing. Learning material, not authoritative for SW2025. |

## Superseded — banner at the top of each

| Doc | Why |
|---|---|
| [ORDINATE_GAP_ANALYSIS.md](ORDINATE_GAP_ANALYSIS.md) | Static pre-live analysis. Started the r8-r19 work; its gap list now lives in the gap doc with current status. |
| [REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md](REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md) | Calls itself a living spec for `target-spec-hybrid-v2`, now archived. Its reading of the reference drawing is still useful. |

## Historical — written for the archived `target-spec-hybrid-v2` tree

These describe an implementation in `archive/`. They are kept as a record of
how decisions were reached. **Do not follow their instructions**, and do not
treat their revision labels (r16 … r62) as related to the trunk's `r-N`.

`CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md`,
`HYBRID_COMPANION_IMPLEMENTATION.md`,
`R16_NEW_CHAT_CONTEXT_HANDOFF.md`,
`R21_REVIEW_HANDOFF.md`,
`R22_REVIEW_RESOLUTION.md`,
`R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md`,
`R23_CODEX_HANDOVER.md`,
`R23_CODEX_INIT_PROMPT.md`,
`R23_CODE_REVIEW_2026-08-04.md`,
`R23_DEFECT_REVIEW_AND_PLAN_2026-08-04_POST_1845.md`,
`R23_HANDOFF_2026-08-04_POST_0616.md`,
`R23_IMPLEMENTATION_PLAN.md`,
`R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md`,
`R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md`,
`TARGET_SPEC_HYBRID_V2_IMPLEMENTATION_PLAN.md`

See [../archive/README.md](../archive/README.md) for why that tree was
archived and what is salvageable from it.

## Update after every iteration

`Changelog.md`, `CURRENT_STATUS.md`, `SOLIDWORKS_API_VALIDATION.md` when an
API contract was established, and `MACRO_SOURCE_REVISION` in
`Module1_Main.bas` — the single version source. `deployment-request.txt` is
regenerated from it; never hand-edit that file.
