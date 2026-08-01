# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Read [Agents.md](Agents.md) first — it is the binding operating contract
(fixture authorization, ask-before-live-SOLIDWORKS rule, VBA engineering rules,
acceptance criteria). Then follow
[docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md](docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md)
for the current R23 objective, accepted probe evidence, blocked gates, and exact
next work package. Use
[docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md](docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md)
for the full review/offline-check workflow, evidence ladder, commands, and
claim language.

## Refer at the start of every task

- [Agents.md](Agents.md)
- [docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md](docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md)
- [docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md](docs/R23_PHASE0_PROBE_FINDINGS_AND_NEXT_STEPS.md)
- [docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md](docs/CLAUDE_STATIC_REVIEW_AND_OFFLINE_CHECKS_HANDOFF.md)
- [docs/CURRENT_STATUS.md](docs/CURRENT_STATUS.md) — latest state; read the actual `MACRO_SOURCE_REVISION` in `src/target-spec-hybrid-v2/Module1_Main.bas`, never assume a doc's revision label is current
- [docs/SOLIDWORKS_API_VALIDATION.md](docs/SOLIDWORKS_API_VALIDATION.md) — accumulated API-contract evidence
- [skills/solidworks-api-lookup/SKILL.md](skills/solidworks-api-lookup/SKILL.md) — mandatory before touching any `sw*` constant or API call

## Update after every iteration

- [docs/Changelog.md](docs/Changelog.md) (note: `Agents.md` calls it `CHANGELOG.md`; the real file is `Changelog.md`)
- [docs/CURRENT_STATUS.md](docs/CURRENT_STATUS.md)
- [docs/SOLIDWORKS_API_VALIDATION.md](docs/SOLIDWORKS_API_VALIDATION.md) — when API contracts were verified
- `MACRO_SOURCE_REVISION` in `Module1_Main.bas` — bump whenever deployable behaviour changes; it is the single version source (`deployment-request.txt` is regenerated from it, never hand-edit)
- `src/target-spec-hybrid-v2/README_IMPORT.md` — source-identity line

## Tools

- [tools/swp-deploy/](tools/swp-deploy/) — guarded deployment of `src/target-spec-hybrid-v2/` into `Fable.swp` (preflight, nine gates, readback verification). Commands and evidence rules are in the handoff doc, section 6.
- [tools/solidworks-automation-companion/](tools/solidworks-automation-companion/) — offline verification layer: the 74-test `unittest` suite (5 known-stale failures — see handoff section 5.2), fake-COM tests, probes. Separately versioned, gitignored; never edit it as part of a main-repo change.

## Design practices, algorithms, coding patterns

- [docs/CODESTACK_DRAWING_API_COVERAGE.md](docs/CODESTACK_DRAWING_API_COVERAGE.md) — complete ledger of the CodeStack drawing-API corpus: object hierarchy, model- vs drawing-context entities, selection/transform/state-restoration patterns, and per-example cautions. Learning material, not authoritative for SOLIDWORKS 2025 contracts.
- [docs/3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md](docs/3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md) — the durable implementation guidance distilled from that corpus (the coverage doc defers to it).
