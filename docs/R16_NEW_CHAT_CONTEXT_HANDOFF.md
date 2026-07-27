# Revision 16 new-chat context handoff

Date prepared: 2026-07-26  
Workspace: `C:\Users\V.T\Documents\VBA 3D TO 2D`  
Prepared source identity: `target-spec-hybrid-v2-2026-07-18-r16`

## Purpose of this handoff

This document gives a new Codex chat the complete context required to receive
and analyze a revision 16 SOLIDWORKS macro run.

The user will provide the actual deployment log, QA report, screenshots, and
other runtime evidence in a separate message after this context has been read.
Do not ask the user to complete a form or replace placeholders.

When this document is first supplied, the new chat should:

1. read the project sources and documents named below;
2. understand the r15 evidence and r16 repair intent;
3. briefly confirm that the context is loaded; and
4. wait for the user to send the r16 outputs before diagnosing or changing
   source.

Do not assume that revision 16 compiled, deployed, or worked until the user
provides that evidence.

## Project goal

The project is a SOLIDWORKS 2025 VBA macro that creates manufacturing drawings
from three authorized part models.

The objective is not merely to make the macro finish or create dimensions. The
drawings must approach the supplied manual references in:

- manufacturing usefulness;
- semantic feature and dimension coverage;
- datum and location correctness;
- section and detail content;
- readability and collision-free layout;
- controlled title data and notes; and
- truthful, fail-closed QA evidence.

The fixed production workflow is:

1. create the required reference-led drawing views;
2. import applicable model-marked dimensions and annotations;
3. add evidence-backed ordinate dimensions only for qualifying missing feature
   locations;
4. avoid duplicated size and location information;
5. arrange views and annotations without crossing sheet, title-block, note, or
   neighboring-view boundaries;
6. populate title data, notes, and part identification; and
7. produce a structured QA report that does not hide unresolved failures.

## Required project reading

Before analyzing the runtime outputs, read:

1. `Agents.md`
2. `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`
3. `docs/CURRENT_STATUS.md`
4. `docs/Changelog.md`
5. `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/QA_REPORT.txt`
6. `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/VISUAL_COMPARISON.md`
7. `test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/MCP_API_CHECKPOINT.md`
8. the current exported source under `src/target-spec-hybrid-v2/`

The matching protected manual reference for the focused fixture is:

`test_assets/reference_drawings/P-0251-14A-001.PNG`

The newest r16 runtime outputs supplied by the user take precedence over the
r15 report and all older summaries.

## Collaboration model

The user controls SOLIDWORKS, the VBA editor, deployment, compilation, and macro
execution.

Codex is responsible for:

- reading and preserving the supplied evidence;
- comparing the generated drawing with the target specification and manual
  reference;
- separating root causes from cascade failures;
- checking uncertain SOLIDWORKS calls against the local API MCP;
- modifying the exported VBA source coherently;
- updating tests and project documentation;
- performing offline validation and deployment preflight; and
- telling the user exactly what evidence or SOLIDWORKS operation is required
  next.

Do not use Computer Use. Ask the user to perform any required SOLIDWORKS action.
Do not use the internet for SOLIDWORKS API research. Use the configured local
`solidworks-api` MCP, which represents the exact SOLIDWORKS 2025 documentation
used by this project.

## Authorized scope and protected assets

Live macro execution is authorized only on:

- `test_assets/models/P-0251-14A-001.SLDPRT`
- `test_assets/models/P-0252-01-001.SLDPRT`
- `test_assets/models/P-0252-01-013.SLDPRT`

The parts may be opened, measured, and queried, but their features,
configurations, dimensions, properties, or saved design state must never be
changed.

Never:

- modify `src/baseline-model-dims/`;
- edit or overwrite files under `test_assets/reference_drawings/`;
- flatten a class or UserForm into a standard module;
- add VBA `Attribute` metadata or a UTF-8 BOM to deployable `.bas` or ordinary
  `.cls` source;
- suppress a truthful QA failure merely to produce `RESULT: PASS`; or
- claim success from static tests without an embedded compile, authorized run,
  and visual/manufacturing assessment.

## Deployment architecture

The guarded deployment target is:

`C:\Users\V.T\Documents\VBA 3D TO 2D\Fable.swp`

The deployer is:

`tools/swp-deploy/Deploy-TargetSpecHybrid.ps1`

The manifest manages 15 replaceable standard/class components. The existing
`ThisLibrary`, `UserForm1`, and `UserFormSection` components retain their special
host/designer architecture.

Metadata-free ordinary class source must be deployed as class modules through
explicit component-type creation and code injection. It must not be imported in
a way that turns a `WithEvents` class into a standard module.

Immediately before the user’s r16 deployment, offline validation proved:

- source identity `target-spec-hybrid-v2-2026-07-18-r16`;
- 61 passing project-local tests;
- deployment preflight found 19 embedded components;
- all 15 managed components were represented;
- the deployment bootstrap was present; and
- the changed deployable VBA source passed metadata, BOM, ANSI, CRLF, and
  whitespace checks.

These are static and preflight results only. They do not prove that r16 was
successfully embedded, compiled, executed, or visually accepted.

## Focused fixture and expected drawing

The next evidence concerns:

`test_assets/models/P-0251-14A-001.SLDPRT`

The reference-led P-0251 drawing target includes:

- a portrait main/front view;
- a narrow left-side view;
- exactly one J-J section through the stepped bore;
- one undimensioned isometric orientation aid;
- the large-bore size and fit definition;
- a grouped six-hole face callout;
- a grouped four-hole tapped side callout;
- ten canonical physical hole locations without duplicate physical records;
- deliberate centre-X and bottom-Y location logic;
- readable model dimensions and ordinate groups;
- no geometry or annotation intrusion into the border, title block, notes, or
  neighboring views; and
- controlled title data, general notes, and part identification.

The reference drawing, rather than a nonzero API count, is the visual and
manufacturing acceptance target.

## What revision 15 proved

The retained r15 run is:

`test_assets/iteration_evidence/macro_qa/20260726_113534_P-0251-14A-001/`

Its deployment evidence proved that the embedded macro matched r15. Its runtime
proved:

- the P-0251-only clockwise 90-degree primary-view rotation;
- exact rotation readback;
- a 1:1 sheet scale;
- two eligible orthographic views plus one independent 1:2 isometric view;
- zero isometric dimensions;
- restored sheet context;
- normal selection mode and zero selected entities at completion; and
- verified atomic QA evidence writing.

The r15 drawing still lacked the J-J section, all manufacturing dimensions and
callouts, ownership-proven locations, ordinate groups, and acceptable layout.

The r15 QA report contained 30 failure lines, but they were not 30 independent
defects. The principal roots were section creation, annotation-import control
flow, visible-edge ownership conversion, layout-zone logic, description-property
resolution, and the missing structural title-block definition. Many stage-gate
failures were downstream consequences.

## Revision 16 repair intent

Revision 16 is the source state whose runtime outputs the user will send next.
It preserves the r15-proved primary rotation, scale, isometric policy, cleanup,
and evidence writing while testing the following repairs.

### 1. J-J section transaction

R15 failed with VBA error 91 at:

`SECTION_STEP=SelectData.ViewAssignment.Before`

R16 does not assign `ISelectData.View`. It relies on the explicitly activated
source drawing view and verifies the section sketch segments, owning drawing
view, selection count, order, and marks before calling
`CreateSectionViewAt5`.

Expected r16 evidence includes:

`SECTION_STEP|step=SelectData.ViewAssignment.Skipped|reason=ActiveSourceViewOwnsNewSketchSegments`

The relevant source is:

`src/target-spec-hybrid-v2/Module2_DrawingPipeline.bas`

### 2. Orthographic model-annotation import

R15 logged both a matching active view and `selectedByID=True`, followed by a
contradictory legacy selection failure.

R16 makes exact `ActiveDrawingView` readback the hard view-context proof.
`SelectByID2=False` may be recorded as a warning but cannot override matching
active-view evidence.

Expected evidence immediately before insertion includes:

`MODEL_IMPORT_EXECUTE|view=...|allViews=False|duplicateDims=True|activeViewMatched=True`

The relevant source is:

`src/target-spec-hybrid-v2/Module4_ModelItemImporter.bas`

### 3. Ownership-proven hole locations

R15 rejected all 111 component-qualified visible edges as `NotCircular` after a
second corresponding-entity conversion.

R16 sends edge objects returned by
`IView.GetVisibleEntities2(component, Edge)` directly into the existing strict
circularity, internal-cylinder, feature-ownership, configuration, axis,
physical-identity, and deduplication gates.

Expected evidence includes:

`ENTITY_SOURCE|view=...|source=IView.GetVisibleEntities2|mapping=DirectComponentEntity`

The relevant source is:

`src/target-spec-hybrid-v2/Module5_FallbackDimensionEngine.bas`

### 4. Zone-aware drawing layout

R15 treated the title-block top as a lower boundary extending across the entire
sheet. The tall portrait primary consequently could not use valid lower-left
space.

R16 reserves the actual lower-right title-block rectangle, permits the primary
to use safe lower-left space, and places remaining views in the title-clear
upper band.

It must continue to reject:

- sheet-border violations;
- title-block intrusion;
- note intrusion;
- view collisions; and
- any layout that would require an unapproved scale change.

Expected diagnostic evidence includes:

`LAYOUT_RESULT|mode=DiagnosticZoneAware|acceptance=False`

and:

`titleBlockReservedAsRectangle=True|lowerLeftZoneUsed=True`

The relevant source is:

`src/target-spec-hybrid-v2/Module9_LayoutEngine.bas`

### 5. Description property

R16 checks the following exact property aliases in order:

1. `Description`
2. `PartName`
3. `Part Name`

Lookup remains configuration-first and document-level second. The macro must
not invent a value when all approved sources are empty.

The relevant source is:

`src/target-spec-hybrid-v2/Module7_TitleBlockEngine.bas`

## Known controlled-template blocker

The r15 sheet visibly contained a title table, but SOLIDWORKS returned:

`ISheet.TitleBlock Is Nothing`

Visible sheet-format lines and notes are not equivalent to a structural
`ITitleBlock`. Revision 16 deliberately does not guess which arbitrary notes
should define the title block.

Unless the user structurally corrected the controlled template before the r16
run, the following may still fail even when all five r16 source repairs work:

- controlled-sheet proof;
- exact linked title-property proof;
- general-note and part-identification region/extent proof;
- annotation-extents QA; and
- final production QA.

Treat these as one controlled-input dependency and its downstream failures, not
automatically as several independent VBA defects. Diagnostic drawing mode should
still produce an inspectable output.

Do not remove, bypass, or weaken this requirement merely to obtain a passing
result.

## How to process the user’s next message

After this context has been acknowledged, the user will send the actual r16
outputs. Accept the files and text directly; do not ask the user to populate a
handoff template.

The useful evidence may include:

- complete PowerShell deployment output;
- first compile error or confirmation of a clean full-project compile;
- complete `QA_REPORT.txt`;
- result-dialog screenshot;
- full-sheet screenshot after dismissing the dialog;
- close-ups of the J-J section, dimensions, ordinates, and title area;
- drawing feature-tree screenshot;
- macro-settings screenshot;
- Immediate Window output; and
- a PDF or image export of the generated drawing.

Not every item is mandatory before beginning. Start with the supplied evidence,
identify material gaps, and ask the user only for specific missing evidence that
is necessary to resolve a decision.

The full QA report is normally saved under:

`test_assets/iteration_evidence/macro_qa/<run timestamp>_<part number>/QA_REPORT.txt`

The result dialog and the `Evidence report:` line near the beginning of the
report contain its exact path.

## Required r16 analysis

When the outputs arrive:

1. confirm the deployed and executed macro identity;
2. read the complete QA report rather than relying on the dialog summary;
3. preserve the report and screenshots under their run evidence directory when
   they are not already there;
4. inspect each r16 repair independently;
5. state which r16 repairs are runtime-proven, which failed, and which remain
   unproved because execution stopped earlier;
6. compare the unobstructed drawing with the P-0251 manual reference;
7. list missing, duplicated, misplaced, or semantically incorrect views,
   sections, dimensions, callouts, ordinates, notes, and title data;
8. distinguish code defects, controlled-input blockers, and downstream cascade
   failures;
9. identify the exact affected procedures and dependent callers;
10. verify every uncertain API contract through the local `solidworks-api` MCP;
11. implement the smallest coherent source correction;
12. preserve all strict feature-ownership, datum, selection, cleanup, and QA
    gates;
13. update `docs/Changelog.md`, `docs/CURRENT_STATUS.md`, and
    `docs/REFERENCE_DRAWING_ANALYSIS_AND_TARGET_SPEC.md`;
14. run the complete offline suite, VBA metadata/BOM/encoding checks, and
    deployment preflight;
15. advance to the next source revision only after the repaired source is
    coherent; and
16. give the user the exact next deployment/run and evidence instructions.

At minimum, evaluate:

- final drawing-view, section, isometric, and eligible-orthographic counts;
- imported annotation and visible display-dimension counts by view;
- whether the isometric remains undimensioned;
- section-line selection, section creation, ownership, and label evidence;
- accepted and rejected ownership candidates by reason and view;
- canonical physical-location and projected-location records;
- horizontal and vertical ordinate groups;
- ordinate selection counts and decoded return codes;
- datum proof and intended centre-X/bottom-Y behavior;
- layout mutations and every border/title/note/view collision result;
- description source and rendered title value;
- title, notes, part-identification, annotation-extents, cleanup, evidence-write,
  and final-QA stages; and
- visible manufacturing completeness against the reference.

Nonzero counts are not sufficient. Confirm that the accepted features,
locations, attachments, datum scheme, grouping, and visible result are the
intended ones.

## Response expected from the new chat

After analyzing the r16 outputs, return:

1. the exact QA report path analyzed;
2. a concise statement of what r16 fixed successfully;
3. a root-cause list for what remains;
4. a separate list of cascade failures;
5. a visual difference list against the manual reference;
6. exact affected modules, procedures, and callers;
7. relevant local SOLIDWORKS 2025 MCP findings;
8. coherent source fixes and the new source revision;
9. offline test, source-format, and deployment-preflight results;
10. the exact evidence or SOLIDWORKS task required from the user next; and
11. an explicit distinction between static verification, embedded compilation,
    runtime execution, and final visual/manufacturing acceptance.
