# R23 Scope and Generalization Planning

**Written:** 2026-08-05, at the end of the r58 to r62 session.

**Purpose:** carry the scope conversation into a fresh planning session
without reconstructing it. This document is deliberately self-contained: it
states the current state, the evidence behind each claim, and the open
product question. Nothing here is a decision - the decisions are the point
of the next session.

Read alongside, in this order:

1. [CURRENT_STATUS.md](CURRENT_STATUS.md) - live state and the r63 package.
2. [Changelog.md](Changelog.md) - entries (38) to (45) cover this session.
3. [R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md](R23_CLAUDE_CODE_IMPLEMENTATION_HANDOFF.md)
   - durable boundaries; its sections 5 to 10 are historical.
4. [Agents.md](../Agents.md) - binding operating contract.

---

## 1. The question that prompted this document

> "How much work do you think and how many more iterations should we roughly
> perform till the time we can use the macro with any SOLIDWORKS drawing and
> expect correct outputs. I know it's an ambiguous question but I'd still
> like some scope of work left."

## 2. The answer given, in full

### 2.1 Where the macro actually is

It is a **validated pipeline for one part**, not a general tool. The
infrastructure underneath is general - geometry decode, selection proof,
evidence emission, the QA stage machinery. The *intent* layer on top is not.
Concretely, from source:

- `Module10.BuildSectionRequirements()` is a hardcoded list of seven
  requirements with literal nominals - 0.018, 0.012, 0.0115, 0.040, 0.047,
  0.1736, 0.1048 - and lane names chosen for this geometry. **Those
  requirements *are* the spec.** For another part they are meaningless.
- `Module2.ReconcileR23SectionDimensions` and
  `Module2.ClampR23SectionAnnotations` both exit unless
  `Module1_Main.GetFixtureKey(evidence.PartPath)` is `"P-0251-14A-001"`.
- Layout runs the profile `P0251ReferenceZones`.
- `Module1_Main.CONTROLLED_TEMPLATE_PATH` is a hardcoded
  `V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT`.
- The fixture allowlist is three parts
  (`P-0251-14A-001`, `P-0252-01-001`, `P-0252-01-013`) and **only the first
  has ever been run**.

### 2.2 Three tiers of "done"

**Tier A - fixture 1 clean. Roughly 4 to 8 iterations.**

r63 placement defects first. Then:

- `NATIVE_CALLOUT_COVERAGE` and `MANUFACTURING_DEFINITION` share one root
  cause - `missing=Attachment` on both `op:EXTRUDEDCUT` and
  `op:HOLEWIZARD`. That is one investigation, historically 2 to 4 runs in
  this project.
- `VIEW_PROJECTION`'s five unprojected locations are the Y-axis side holes.
  Likely needs a view that shows them, or a reasoned acceptance.
- `SECTION_DIMENSIONS`' two missing requirements
  (`LOWER_WALL_STEP_11_5`, `LONG_VERTICAL_REF_173_6`) are honest absences -
  no geometry in the view. Probably a spec decision, not code.
- `FINAL_QA` falls out when the rest do.

**Tier B - all three authorized fixtures. Roughly 8 to 15 more.**

This is where you learn how general it really is. On `P-0252-01-001` and
`P-0252-01-013` every fixture-keyed path no-ops today. Expect this tier to
force requirements to be **derived from the model** rather than authored as
literals - a redesign of the intent layer, not a sequence of bugfix runs.

**Tier C - any drawing. Not an iteration count.**

A rearchitecture, and arguably a different product. It needs:

- requirements derived from the feature tree;
- a section path chosen from geometry rather than from a known bore;
- placement that survives arbitrary view counts and lane collisions;
- template independence; and
- dropping the fixture allowlist - which today is a **safety mechanism, not
  an obstacle**. Removing it changes the risk model of every run.

Months, not sessions. Worth scoping fresh rather than estimating from here.

### 2.3 Recommendation given

Define done as **Tier A + Tier B: three fixtures clean**. Roughly **12 to 23
iterations**, a handful of sessions at the current pace. That is a defensible
deliverable. Treat "any drawing" as a separate decision made *after* Tier B
tells you how much of the intent layer generalizes.

### 2.4 Caveat on the estimate

The iteration count will not collapse much. This session found roughly one to
two real defects per run, every one from the run's own evidence rather than
from inspection. That discipline is what makes the output trustworthy and
also what sets the pace. The r61 to r62 sequence is typical: a fix lands, a
gate turns PROVED, and the same run exposes the next thing. Expect that to
continue rather than a run that fixes five stages at once.

---

## 3. State at the end of this session

`MACRO_SOURCE_REVISION` = `target-spec-hybrid-v2-2026-08-05-r62`, deployed
and run. Companion offline suite 619/619. Latest run
`test_assets/iteration_evidence/macro_qa/20260805_071309_P-0251-14A-001`.

**Five required stages fail**, down from nine at the start of the session:

| stage | detail |
| --- | --- |
| `NATIVE_CALLOUT_COVERAGE` | `incomplete=2`, both `missing=Attachment` |
| `VIEW_PROJECTION` | `locationsWithout=5`, all Y-axis |
| `SECTION_DIMENSIONS` | `satisfied:5/missing:2` |
| `MANUFACTURING_DEFINITION` | same `Attachment` root cause |
| `FINAL_QA` | aggregate of the above |

**What this session achieved:** the section view went from having no
dimensions at all to five created associative dimensions with verified
nominals (18.00, 12.00, Ø40.00, Ø47.00 H7 +0.03/0.00, 104.80), and
`ANNOTATION_EXTENTS` returned to PROVED.

**Open, and known to be wrong** - three placement defects listed in
[CURRENT_STATUS.md](CURRENT_STATUS.md), all introduced by this session's
work. Two dimension pairs render on top of each other and the SECTION J-J
label was moved 52 mm onto the view. The QA stage cannot see any of it; the
user's screenshot caught all three.

---

## 4. Questions the planning session should answer

These are open. None has been decided.

1. **Is Tier B the definition of done, or is Tier C the actual goal?** The
   answer changes whether the intent layer gets refactored now or later.
2. **What happens to the fixture allowlist?** It is the main runtime safety
   guarantee. Any generalization has to state what replaces it.
3. **Where do requirements come from in a general tool?** Feature tree,
   a per-part spec file, or a drawing template convention. This is the
   single largest design question.
4. **Are the two missing section requirements a code gap or a spec error?**
   `LONG_VERTICAL_REF_173_6` has no pair among the nineteen distinct Y
   values in the view. That may mean the requirement is wrong for this view.
5. **Does `DIAGNOSTIC_DRAWING_MODE` stay on?** It lets independent stages
   continue past a failure, which is why a run can complete with five failed
   stages. Turning it off changes what "a run completed" means.
6. **What is the acceptance artefact?** Today the user reviews a screenshot.
   Tier B across three fixtures needs something more repeatable.

---

## 4a. Decisions made, 2026-08-05 planning session

Answers to section 4, in order:

1. **Done tier: Tier C** (any drawing). Tier B is not the target; the intent
   layer gets redesigned rather than patched fixture-by-fixture. This is a
   rearchitecture decision, not an iteration-count one - expect months, per
   section 2.2.
2. **Fixture allowlist: design its replacement now**, ahead of full Tier C
   work. It is today's only runtime safety guarantee; Tier C cannot drop it
   without something concrete in its place (candidate shapes: a read-only/
   dry-run mode for unauthorized parts, or an explicit per-new-part user
   confirmation gate). Not designed yet - this is the next open task.
3. **Requirements source: feature-tree derived.** The intent layer computes
   section-dimension requirements from the model's feature tree at runtime
   instead of `Module10.BuildSectionRequirements`'s hardcoded literals. This
   is the largest single piece of Tier C work and has not been scoped into
   steps yet.
4. **`DIAGNOSTIC_DRAWING_MODE`: stays on.** Keep independent-stage-continues-
   past-failure behavior through this phase; only turn off once stages are
   mostly clean, so "run completed" still surfaces every failing stage in
   one pass.
5. **`LONG_VERTICAL_REF_173_6`: check the reference drawing before writing
   any code.** Look up whether 173.6mm appears as a dimension on P-0251's
   manual/reference drawing and which view it's in. That fact decides
   spec-error vs code-gap; nobody has checked it yet. Do this before
   further section-dimension work on this requirement.
6. **Acceptance artefact: structured placement report.** Extend the QA stage
   itself with machine-checkable geometry (bounding boxes, overlap checks,
   offsets) rather than adding a separate visual-diff layer. This is meant
   to catch defects like r62's RD1/RD2 overlap and the J-J label drift,
   which today only a human screenshot review catches. Automated visual
   diff was considered and deferred, not rejected - revisit if the
   structured report proves insufficient.

**Immediate next task implied by the above:** check the P-0251 reference
drawing for the 173.6mm dimension (#5) - it's cheap and unblocks a decision
before any Tier C design work starts.

## 5. Constraints that carry forward unchanged

- Never save the fixture model. Never overwrite a reference or manual
  drawing. The generated drawing is unsaved and disposable.
- Only the three `Module1_Main.IsAuthorizedFixture` paths.
- `-AllowMutation` is mandatory on the production runner and never implied.
- Do not restore automatic rescaling or content-envelope repositioning - the
  2026-08-04 user decision preserves the layout as-is.
- Do not disable `DIAGNOSTIC_DRAWING_MODE` to make a run look clean.
- The user holds the deploy decision.
- Visual and manufacturing acceptance stay with the user.
- Before touching any `sw*` constant or API call, use
  [skills/solidworks-api-lookup/SKILL.md](../skills/solidworks-api-lookup/SKILL.md).
- Load-bearing offline tests are mutation-verified before they are trusted.
  Two vacuous tests were caught that way this session.
