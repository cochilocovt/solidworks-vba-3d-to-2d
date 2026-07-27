# SOLIDWORKS VBA 3D-to-2D Drawing Automation

This project contains a SOLIDWORKS 2025 VBA workflow for producing manufacturing-oriented 2D drawings from saved 3D part models. It combines model-annotation import, qualified hole-centre ordinate dimensions, an optional primary section, title-block data, notes, barcode generation, and QA evidence.

The active development source is in `src/active-ordinate/`. The protected comparison baseline is in `src/baseline-model-dims/`, and the current target-spec source is in `src/target-spec-hybrid-v2/`.

## Repository contents

- `Agents.md` — operating rules and acceptance criteria for coding agents.
- `docs/` — architecture, requirements, API validation, change history, and current status.
- `src/` — exported VBA modules and forms. These are the source of reviewable changes.
- `tools/swp-deploy/` — local deployment and verification helpers for synchronizing source into a `.swp` macro project.

## Requirements

- SOLIDWORKS 2025 with its VBA references available.
- A saved SOLIDWORKS part and a valid drawing template configured locally.
- Windows PowerShell for the supplied deployment helper.

## Local-only materials

The repository deliberately excludes compiled `.swp` files, part models, manual reference drawings, generated drawings, QA captures, and iteration evidence. Those items may contain proprietary engineering data or are regenerated during validation. The separately versioned `tools/solidworks-automation-companion/` checkout is also excluded; obtain it from its own upstream repository when companion validation is needed.

## Working safely

Read `Agents.md` before changing the macro. In particular, keep deployable VBA files ANSI/Windows-1252 without a BOM, preserve `src/baseline-model-dims/`, and validate any VBA change in the SOLIDWORKS editor against authorized fixtures before accepting it.
