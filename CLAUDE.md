# CLAUDE.md

Guidance for Claude Code working in this repository.

Read [Agents.md](Agents.md) first — it is the binding operating contract
(fixture authorization, live-SOLIDWORKS rule, VBA rules, evidence honesty).

## Current state

**Trunk: `src/baseline-model-dims/`** (2026-08-05). The former trunk
`target-spec-hybrid-v2` is in [archive/](archive/README.md) — 36k lines, too
fixture-coupled to generalise. `src/active-ordinate/` is history.

Objective is Tier C: a macro that works on any drawing, not just the three
fixtures. Decisions and rationale in
[docs/R23_SCOPE_AND_GENERALIZATION_PLANNING.md](docs/R23_SCOPE_AND_GENERALIZATION_PLANNING.md)
§4a.

Read the actual `MACRO_SOURCE_REVISION` in
`src/baseline-model-dims/Module1_Main.bas`. Never trust a doc's revision
label.

## Start here

- [Agents.md](Agents.md) — binding contract
- [docs/BASELINE_TO_REFERENCE_DRAWING_GAP.md](docs/BASELINE_TO_REFERENCE_DRAWING_GAP.md)
  — what the trunk still needs to produce the reference drawing, and the
  phased plan
- [docs/R23_SCOPE_AND_GENERALIZATION_PLANNING.md](docs/R23_SCOPE_AND_GENERALIZATION_PLANNING.md)
  — the three tiers of "done" and the open product questions
- [docs/CURRENT_STATUS.md](docs/CURRENT_STATUS.md)
- [docs/SOLIDWORKS_API_VALIDATION.md](docs/SOLIDWORKS_API_VALIDATION.md) —
  accumulated API-contract evidence
- [skills/solidworks-api-lookup/SKILL.md](skills/solidworks-api-lookup/SKILL.md)
  — mandatory before touching any `sw*` constant or API call
- [docs/CODESTACK_DRAWING_API_COVERAGE.md](docs/CODESTACK_DRAWING_API_COVERAGE.md)
  — check the 33-row ledger for a tested pattern **before** designing one.
  It also states plainly where the corpus has nothing, which is just as
  useful. Rows 13, 17 and 31 cover `GetVisibleEntities2`, view-scoped
  `ISelectData`, and `IEntity.Select4`.

### These two are enforced, not requested

Instruction text here did not hold: on 2026-08-06 three `swDisplayMode_e`
constants sat in the trunk with values from other enum members and survived
fifteen live runs, because a wrong display mode renders a plausible view
instead of raising. Two gates now fire without depending on anyone's memory.

- `.claude/hooks/require_api_lookup.py` (PreToolUse) **blocks** an edit to a
  managed `.bas`/`.cls` that introduces a `sw[A-Z]…` token when no
  solidworks-api MCP lookup has been recorded in the last 30 minutes. Edits
  with no such token pass untouched.
- `tests/test_api_constant_provenance.py` **fails the suite** when a `sw*`
  constant is compiled into the trunk without a provenance record in
  `docs/SOLIDWORKS_API_VALIDATION.md`. It runs before every deployment.

Neither proves a value is right — only the installed SW2025 type library does
that. They make skipping the check impossible to do silently.

## Update after every iteration

- [docs/Changelog.md](docs/Changelog.md) (Agents.md calls it `CHANGELOG.md`;
  the real file is `Changelog.md`)
- [docs/CURRENT_STATUS.md](docs/CURRENT_STATUS.md)
- [docs/SOLIDWORKS_API_VALIDATION.md](docs/SOLIDWORKS_API_VALIDATION.md) when
  API contracts were verified
- `MACRO_SOURCE_REVISION` in `Module1_Main.bas` — the single version source.
  `deployment-request.txt` is regenerated from it; never hand-edit that file.

## Tools

- [tools/production-runner/Run-R23Production.ps1](tools/production-runner/Run-R23Production.ps1)
  — **the only way to run the real macro.** Deploys, compiles the whole VBA
  project, opens the authorized part read-only, invokes `Module1_Main.main`,
  prints the QA stage table. Requires `-AllowMutation`; refuses to invoke
  `main` unless pre-flight logged `ready=True`. Never assemble a manual
  sequence instead.
- [tools/probe-runner/Run-R23Probes.ps1](tools/probe-runner/Run-R23Probes.ps1)
  — deploy + programmatic compile. **Its probe stage is currently inert:**
  the nine `R23_Probe*` entry points lived in the archived Module10–19 and
  were not ported. Use it for deploy/compile evidence; add a probe entry
  point to the trunk when there is a read-only question worth asking.
- [tools/swp-deploy/](tools/swp-deploy/) — guarded deployment into
  `Fable.swp` (preflight, gates, readback).
- [tools/solidworks-automation-companion/](tools/solidworks-automation-companion/)
  — offline `unittest` suite. 34 live tests cover the trunk and the
  deployment tooling. The 585 tests for the archived implementation moved to
  `archive/target-spec-hybrid-v2-tests/`.

```bash
python -m unittest discover -s tools/solidworks-automation-companion/tests -q
```

## Design references

- [docs/3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md](docs/3D_TO_2D_DRAWING_AUTOMATION_FIELD_GUIDE.md)
  — durable implementation guidance
- [docs/CODESTACK_DRAWING_API_COVERAGE.md](docs/CODESTACK_DRAWING_API_COVERAGE.md)
  — CodeStack corpus ledger. Learning material; not authoritative for
  SOLIDWORKS 2025 contracts.
