# Architecture

Trunk: `src/baseline-model-dims/`. Twelve managed components plus two forms.
Deployed into `Fable.swp` by `tools/swp-deploy`.

Revision at time of writing: `trunk-2026-08-06-r19`. Read the actual
`MACRO_SOURCE_REVISION` in `Module1_Main.bas`; never trust a doc's label.

## Execution flow

```text
Module1_Main.main
  guard: active doc is a saved part, and IsAuthorizedFixture(path)
  -> UserForm1  (modal; see "The form owns the settings" below)
  -> Module2_DrawingPipeline.RunDrawingPipeline
       -> Module3_ModelAudit          hole-like feature inventory
       -> CreateViews                 CreateDrawViewFromModelView3 per orientation,
                                      then CreateSectionViewAt5 if requested
       -> Module4_ModelItemImporter   InsertModelAnnotations4  (if UseModelDimensions)
       -> Module5_FallbackDimensionEngine
                                      ordinate chains          (if UseOrdinateDims)
            gated per view by Module8_ViewClassifier
       -> Module4 auto-arrange
       -> Module7_TitleBlockEngine    custom properties
       -> ForceRebuild3
       -> Module5.PruneDanglingAcrossDrawing   MUST be after the rebuild
       -> Module6_QAEngine.BuildRunSummary
       -> Module21_EvidenceSink.WriteQaReport
```

`Module20_ProbeRunner` is outside this flow: it holds the programmatic VBE
compile gate and the production pre-flight, and never mutates a drawing.

## Component map

| Component | Responsibility |
|---|---|
| `Module1_Main` | Entry point, fixture authorization, `MACRO_SOURCE_REVISION`, template resolution |
| `Module2_DrawingPipeline` | View creation, section creation, stage ordering |
| `Module3_ModelAudit` | Feature-tree hole inventory |
| `Module4_ModelItemImporter` | `InsertModelAnnotations4`, auto-arrange, dimension counting |
| `Module5_FallbackDimensionEngine` | Ordinate candidate collection, datum contract, chain creation, dangling prune, dimension readback |
| `Module6_QAEngine` | Run summary, per-view roster, verdict |
| `Module7_TitleBlockEngine` | Custom-property writes |
| `Module8_ViewClassifier` | `IView.Type` + `GetOrientationName` → view role → per-view dimension policy |
| `Module20_ProbeRunner` | VBE compile gate, production pre-flight |
| `Module21_EvidenceSink` | Evidence log and QA report files |
| `BtnHandler`, `SectionBtnHandler`, `SectionDlgBtnHandler` | Form event plumbing |

## Ordinate engine, in the order it runs

1. `ClearSelection2`, then `ActivateView`.
2. `SelectData` scoped with `swSelData.View = swView` — **let-assignment, not
   `Set`**; the property is `propertyput`.
3. `GetVisibleComponents` → `GetVisibleEntities2(component, Edge)`.
4. Per edge: full circle → station on both axes; arc → station on both axes,
   counted separately; straight line → station on the one axis it is
   perpendicular to.
5. Coordinates divided by `IView.ScaleDecimal` so stations, tolerances and
   reports are in model units at any scale.
6. Per-axis 1-D dedup, with an alternate entity retained per station.
7. Datum per axis. An **end** datum must be a straight edge; a **centreline**
   datum is unrestricted.
8. Datum selected first, then the rest; `AddOrdinateDimension`; `SetPickMode`
   to close the group.
9. After the pipeline's rebuild: prune any ordinate whose `IAnnotation.
   IsDangling` is true.

## The form owns the settings

`ResetGlobalConfig` in `Module1_Main` is the **no-form fallback, not the
user-visible default.** `UserForm1` seeds its controls from saved registry
settings (`ReadBoolSetting`/`SaveSetting`) and writes them back over
`GlobalConfig` when the operator confirms, so the form wins on every run that
shows it — which is every run.

The two forms are outside `deployment-manifest.json` because they carry no
`VERSION 5.00 … Begin/End` designer block and cannot be imported; their
controls are built at runtime. Changing what the operator sees therefore means
hand-editing the form in the VBE, not editing `ResetGlobalConfig`.

## Enforcement

- `.claude/hooks/require_api_lookup.py` blocks an edit to a managed
  `.bas`/`.cls` that introduces a `sw[A-Z]…` token when no solidworks-api MCP
  lookup was recorded in the last 30 minutes.
- `tests/test_api_constant_provenance.py` fails the suite when a `sw*`
  constant has no provenance record in `SOLIDWORKS_API_VALIDATION.md`.
- `tools/swp-deploy` refuses on hygiene violations and verifies the embedded
  revision by readback.
- `tools/production-runner` refuses to invoke `main` unless the pre-flight
  logged `ready=True`.
