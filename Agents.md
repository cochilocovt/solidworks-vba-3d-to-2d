# SOLIDWORKS VBA Drawing Automation — Agent Instructions

## Purpose

This repository contains a SOLIDWORKS 2025 VBA macro that creates manufacturing drawings from saved 3D part models.

The objective is not merely to make the macro finish or create dimensions. Its drawings must approach the manually prepared references in manufacturing usefulness, correctness, coverage, readability, and layout.

The macro supports orthographic and isometric views, model-item import, hole-centre ordinate dimensions, optional Hole Wizard callouts, one deterministic primary section, title-block population, notes, barcode, and QA output.

## Highest-Priority Standing Authorization

This is an explicit, continuing authorization from the user and the main project rule to keep in mind. It does not need to be requested again unless the user revokes or narrows it.

The agent has full authority to open, view, inspect, measure, query through the SOLIDWORKS 2025 API, and otherwise analyze exactly these three supplied test parts:

- `test_assets/models/P-0251-14A-001.SLDPRT`
- `test_assets/models/P-0252-01-001.SLDPRT`
- `test_assets/models/P-0252-01-013.SLDPRT`

The agent also has full authority to:

- Back up, edit, save, compile, and run `src/active-ordinate/active_ordinate.swp`.
- Revise its embedded VBA modules and keep the exported modules in `src/active-ordinate/` synchronized.
- Control SOLIDWORKS and its VBA editor, including activating documents, handling expected dialogs, rebuilding, compiling, running the macro, and inspecting output.
- Run the active macro against the three authorized parts, assess the drawings, diagnose failures, make further scoped corrections, and iterate until requirements are satisfied or a genuine blocker is established.
- Use narrowly scoped SOLIDWORKS probes and the project-local Python companion when they materially help validate an API contract or inspect output.
- Create disposable drawings and save screenshots, PDF/JPG/BMP exports, QA reports, Immediate Window logs, and JSON/Markdown evidence under `test_assets/`.

This authorization has strict boundaries:

- Run the macro only against those three fixtures. Any other model requires new permission.
- The three parts are authorized for opening and analysis, not for changing or saving their design, features, dimensions, configurations, or properties.
- Never modify `src/baseline-model-dims/`.
- Never edit or overwrite manual drawings under `test_assets/reference_drawings/`.
- Make a recoverable backup of `active_ordinate.swp` before the first material binary/VBA-project change in an iteration.
- API success, nonzero dimensions, or a macro that finishes without error is not proof of a correct manufacturing drawing.

Within these boundaries, proceed autonomously through offline analysis, source
preparation, API research, evidence review, and other non-live work. Live
SOLIDWORKS and Computer Use work is governed by the collaboration rule below;
the standing authorization does not waive that required user choice.

## Computer Use and Live SOLIDWORKS Collaboration

Before every distinct task that would require Computer Use, controlling the
SOLIDWORKS application, interacting with the SOLIDWORKS VBA editor, or querying
a live SOLIDWORKS session, the agent must first ask the user whether that task
should be performed by the user or by the agent.

- State the exact proposed live task and ask whether the user wants to perform
  it and share the result, or wants the agent to use Computer Use.
- Do not open, control, inspect, or interact with live SOLIDWORKS until the user
  explicitly chooses the agent for that specific task.
- Standing authorization for the three fixtures defines the permitted scope; it
  is not permission to choose Computer Use without asking.
- A user choice applies only to the stated task. Ask again before a materially
  different live SOLIDWORKS task.
- If the user chooses to perform the task, provide exact concise steps, wait for
  the returned output, and analyze that evidence without using Computer Use.
- If the user chooses the agent, use Computer Use only for the agreed task,
  remain within the three authorized fixtures, preserve the models, and report
  what was inspected or changed.
- If the user says not to use Computer Use, do not use it again until the user
  explicitly authorizes a later specific task.

By default, the user performs changes inside the SOLIDWORKS VBA editor,
compiles the VBA project, deploys or runs the macro, and shares the resulting
logs, QA reports, and screenshots. The agent prepares and reviews exported
workspace source, diagnoses failures, compares drawings with the references,
and performs API/documentation and offline verification work.

When the user explicitly assigns a live SOLIDWORKS task to the agent, the
agent's primary live role is to assess part features, inspect the generated 2D
drawing, perform verification-level checks, and capture evidence. It must not
modify or save the design of an authorized part.

Using the local `solidworks-api` MCP to read API documentation is not live
SOLIDWORKS access and does not require this operator-choice question.

## Source Layout and Authority

- `src/baseline-model-dims/` is the protected baseline. It proves working model-dimension import behavior and must remain unchanged.
- `src/active-ordinate/` is the active development version and the only VBA snapshot authorized for repair.
- Keep `active_ordinate.swp` and its exported modules synchronized.
- The user's current VBA editor code is authoritative when it differs from an older export, screenshot, transcript, document, or extracted snapshot. Inspect or export current embedded source before overwriting it and report meaningful mismatches.
- The project-local Python skill fork is a companion verification layer. VBA remains the production drawing generator unless the user explicitly changes the architecture.

## Environment and API Authority

- SOLIDWORKS 2025 with an active, saved `.SLDPRT`
- Required drawing-template paths must exist
- The updated `solidworks-api` MCP, official SOLIDWORKS 2025 API documentation, and the installed 2025 type library are authoritative
- Validate uncertain COM/VBA binding behavior in the installed build
- Never guess enum values or copy older wrapper constants and calls without 2025 verification

## Required Project Skill: SOLIDWORKS API Lookup

Before writing, changing, or reviewing VBA that touches a SOLIDWORKS API call, an `sw*` constant or enum, drawing/view/dimension/annotation/selection behavior, or any file under `src/`, read and use `skills/solidworks-api-lookup/SKILL.md`.

- Use the local `solidworks-api` MCP to establish exact signatures, return codes, enum values, and behavioural contracts; search results alone are only a locator, so confirm the relevant member lookup and read its Remarks.
- Treat values from the MCP compatibility snapshot as evidence, not confirmed SOLIDWORKS 2025 facts. Verify load-bearing members and numeric values against the installed 2025 type library/Object Browser or a narrow live probe.
- Keep MCP evidence, local verification, compilation, execution, and visual acceptance distinct in reports. Record material findings in `docs/SOLIDWORKS_API_VALIDATION.md` and check current source before relying on older evidence.

## Core Product Requirement

Use one fixed hybrid workflow:

1. Import applicable model-marked dimensions and annotations.
2. Add ordinate dimensions only for qualifying hole centres or feature locations.
3. Avoid unnecessary duplication between imported and ordinate dimensions.
4. Arrange dimensions when configured.
5. Populate title block, notes, and barcode.
6. Produce truthful QA output.

Do not reintroduce separate user-selectable model-dimension and ordinate modes unless requested.

Preserve the baseline importer while repairing the active ordinate workflow. The result must use the selected datum correctly, reject unrelated circular geometry, avoid duplicates, skip unsupported views, remain inside sheet/title-block boundaries, stay readable, and report failed API operations honestly.

## Required Iteration Loop

1. Preserve current embedded VBA source and back up the macro.
2. Identify the defect, evidence, likely root cause, affected procedures, and callers.
3. Compare relevant behavior with the protected baseline.
4. Verify uncertain SOLIDWORKS contracts against 2025 sources or a narrow live probe.
5. Make the smallest coherent active-macro change and synchronize exported source.
6. Compile the entire VBA project in the SOLIDWORKS VBA editor.
7. Run only on authorized fixtures.
8. Capture complete Immediate Window output, QA results, settings, and drawing screenshots.
9. Compare the result with the manual reference and target specification.
10. Iterate when evidence identifies another in-scope defect.

Do not claim VBA success from static tests or a Python probe alone. The embedded project must compile, execute, and pass visual/semantic assessment.

## VBA Engineering Rules

- Always use `Option Explicit`.
- Keep deployable `.bas` and ordinary `.cls` source free of VBA export metadata.
  Do not add `Attribute VB_*`, member `Attribute ...`, or a UTF-8 BOM; save these
  files as Windows-1252/ANSI so SOLIDWORKS VBA cannot expose metadata or BOM
  bytes as uncompilable source. Preserve native metadata only in protected
  document modules and `.frm` files that require it and are not deployment
  inputs.
- Keep code compile-safe and retain public signatures unless all callers are updated.
- Inspect callers before changing a public procedure, type, enum, field, or shared constant.
- Avoid duplicate public declarations and ambiguous constants.
- Prefer explicit SOLIDWORKS types supported by installed references.
- Check API return values and error codes.
- Activate the target view before view-scoped selection.
- Clear selections after `SelectData`, `Select4`, `MultiSelect2`, annotation import, and dimension creation.
- Restore normal selection mode with `SetPickMode` after ordinate operations.
- Keep model, view, and sheet coordinate systems explicit; SOLIDWORKS lengths are normally metres.
- Prefer narrow, testable corrections over speculative rewrites.
- Add diagnostics that identify the view, operation, return code, and VBA error.


## Drawing-Quality Rules

Assess more than API success:

- Datum correctness and ordinate coverage
- Missing or duplicated dimensions/chains
- Model-dimension and ordinate overlap
- View ownership and unsupported-view ordinates
- Text, leader, and annotation collisions
- Sheet, title-block, and note intrusion
- Readability at configured scale
- Hole-callout behavior
- Missing feature-size, location, section, or reference information
- Designer-like placement and manufacturing usefulness

Use `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md` and `docs/ORDINATE_GAP_ANALYSIS.md` as requirements and gap evidence, while newer verified source and live results take precedence when conflicts are reported.

## Sections and Forms

- The UI may store five section entries, but the current data model only records `Label` and `Vertical`.
- Only one deterministic primary section is currently supported.
- Genuine multi-section support requires a user-approved data-model/UI redesign.
- Use drawing-view or sheet coordinates for section geometry.
- Do not change `UserForm1`, `UserFormSection`, `BtnHandler`, `SectionBtnHandler`, or `SectionDlgBtnHandler` unless an event/interface defect is demonstrated.

## Validation Matrix and Evidence

Before production acceptance, compile the full project and cover the three authorized parts across the relevant matrix:

- Model-dimension import
- Genuine hole locations
- Bottom-Left, Center, and Top-Left datum origins
- Hole Wizard callouts off/on
- Auto-arrange behavior
- No section, one horizontal section, and one vertical section where applicable
- Supported orthographic and excluded isometric views
- Duplicate dimensions/chains
- Sheet/title-block/note collisions
- Title-block properties, notes, barcode, and QA output

Use the smallest focused matrix after a narrow fix, then run the complete three-part regression before acceptance.

Preserve under `test_assets/`: compile/runtime errors, highlighted line, complete Immediate Window output, screenshots, relevant feature-tree/API evidence, SOLIDWORKS build, macro settings, QA output, probe evidence, exports, and comparison notes. Never overwrite manual reference drawings; retain useful failed output as regression evidence.

## Code Delivery and Documentation

When modifying code:

1. State the defect, root cause, evidence, and affected modules.
2. State dependent callers/modules that must be replaced together.
3. Keep embedded and exported active source synchronized.
4. Return complete replacement text for changed `.bas`, `.cls`, or `.frm` files unless a diff is requested.
5. Report compile and focused test results.
6. Distinguish static verification, probe verification, VBA compilation, macro execution, and visual acceptance.
7. Update `docs/CHANGELOG.md` and `docs/CURRENT_STATUS.md`.
8. Preserve unrelated user changes and stay within scope.

## Non-Goals Unless Explicitly Requested

- Do not redesign the whole macro while debugging one defect.
- Do not turn Python into a second production drawing engine.
- Do not remove title-block, QA, barcode, notes, or view generation to simplify debugging.
- Do not alter the protected baseline to make snapshots match.
- Do not run the macro on models other than the three authorized fixtures.
