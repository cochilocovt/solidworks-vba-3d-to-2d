# Codex session init prompt

Paste this as your first message to Codex in this directory.

---

You're picking up work on a SOLIDWORKS 2025 SP1.2 VBA drawing-automation
macro (R23 rebuild). Before doing anything else, read these in order:

1. `docs/R23_CODEX_HANDOVER.md` — full handover: progress, architecture,
   unresolved bugs, next steps. Start here.
2. `Agents.md` — binding operating contract. Non-negotiable constraints on
   fixture use, live-SOLIDWORKS access, VBA rules, acceptance criteria.
3. `docs/R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md` — evidence ladder and
   claim-language conventions (static vs. API vs. live-run proof).
4. `docs/R23_IMPLEMENTATION_PLAN.md` — full Phase 0-12 task list, current
   per-task status.
5. `skills/solidworks-api-lookup/SKILL.md` — read before touching any `sw*`
   constant or API call; query the `solidworks-api` MCP, never guess a
   value from memory or an older wrapper.

Standing rules carried over from the prior session, still in force:

- Never modify `src/baseline-model-dims/`, never modify or save a fixture
  model, never overwrite the manual reference drawings. Only the three
  fixtures `Agents.md` authorizes may be used.
- You do not have live SOLIDWORKS or VBA-editor access. I run all live
  SOLIDWORKS/VBE steps myself. Give me exact manual deployment, compilation,
  and execution steps, then wait for me to paste back the Immediate Window
  output before you claim anything ran successfully.
- Static verification (compiles, contract tests, preflight) is never
  reported as runtime proof. Always state which evidence tier a claim rests
  on.
- Every VBA source file must stay cp1252/ANSI (all bytes < 0x80), CRLF, no
  BOM, no `Attribute` metadata, `Option Explicit`, max 79 columns. Normalize
  any file you touch to match.
- Ignore drawing-authored tolerances in general for now — I'll define the
  rule for when one should be added once I've discussed it with my
  designer. The one resolved exception (H7 on the section bore) is
  documented in the handover.
- Preserve all existing work. Never discard, overwrite, or silently
  supersede uncommitted changes.

My immediate ask: fix the four unresolved bugs listed in
`docs/R23_CODEX_HANDOVER.md` under "Unresolved bugs" (Phase 8 nominal-read
failure, Phase 8 stale diameter-type assumption, Phase 8 duplicated
evidence lines, Phase 9 probe crashing on a production-only gate). Look up
the relevant API contracts via the `solidworks-api` MCP before changing any
`IDimension`/`IDisplayDimension` call. Then extend the matching contract
test files in `tools/solidworks-automation-companion/tests/` to lock in the
fixes, and run the full suite - expect the same 5 known-stale R20 failures
and nothing new.

Confirm you've read the handover and state your plan before editing source.

---
