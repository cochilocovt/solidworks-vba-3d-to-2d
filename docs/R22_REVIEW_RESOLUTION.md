# R22 Review Resolution

Date: 2026-07-29

Source identity:
`target-spec-hybrid-v2-2026-07-29-r22`

R22 starts from the complete r20/r21 line (`f2b2b29`, `1698b7d`, and
`3bffcfc`), retains the review changes that satisfy their contracts, and
replaces the three implementations that remained unsafe or incomplete.

## Corrected findings

### Circular-edge proof

`Module5_FallbackDimensionEngine.TryReadClosedCircularEdge` now fails closed
when `ICurve.IsCircle` returns False.

The prior fallback inferred a circular edge from
`ISurface.CylinderParams` and projected the cylinder axis origin to the
`CurveParamData.StartPoint` axial coordinate. That construction is valid only
after proving a planar circle perpendicular to the cylinder axis. A closed
oblique trim can be an ellipse, so the fallback could create a false hole
centre.

The local MCP contract states that `ICurve.IsCircle=False` means another curve
type and that `CircleParams` is the seven-value circle contract. The retained
P-0251 curve probe independently records `is_circle=true` and valid
`circle_params` for all relevant owned hole edges, so the unsafe fallback is
not required for that fixture.

### Pattern feature literals

`Module3_ModelAudit.IsPatternFeatureType` now uses the exact
`IFeature.GetTypeName2` strings documented in the 2025 table:

- `APattern`, not `FillPattern`;
- `LocalChainPattern`, not `ChainPattern`;
- `DimPattern`, not `VariablePattern`;
- `DerivedHolePattern`;
- `SketchPattern`; and
- `LocalSketchPattern`.

The already-correct linear, circular, curve, table, derived linear/circular,
local linear/circular/curve, mirror-pattern, and mirror-solid strings remain.
All supported types still route through `IFace2.GetSeedFeature`, with
`GetPatternSeedFeature` retained as a compatibility fallback.

### Ordinate arrangement

`Module4_ModelItemImporter` no longer assigns ordinate annotations to shared
module-level side lanes.

The r21 approach could merge independent ordinate sets that happened to use
the same side, and generic type `swOrdinateDimension` members could still choose
different sides independently. The MCP and installed SOLIDWORKS 2025 SP1.2
interop expose no set-identity member on `IDisplayDimension`, but do expose
`AutoJogOrdinate() As Boolean`.

When `AlignDimensions` returns False, r22 therefore:

1. calls and checks `AutoJogOrdinate` for ordinate types 1, 7, 8, and 16;
2. reads back each annotation origin and requires content-border containment;
3. uses deterministic `SetPosition2` lanes only for non-ordinate dimensions;
   and
4. retains radial/diametric dimensions only when their existing origins are
   already safe.

This delegates ordinate-set ordering and jog behavior to SOLIDWORKS instead of
inventing an unproved group identity.

## Retained r21 corrections

R22 retains the changes that the review found correct:

- transformed hole centres and mapped datum vertices require view-outline
  containment, while the projected model origin does not;
- FACE and SIDE manufacturing callouts use separate placement lanes;
- annotation self-collision uses pointer identity plus nonempty unique-name
  fallback;
- the controlled general-notes constant can prove a static controlled note;
- section-line parsing is cursor-bounded without assuming the returned `Size`
  equals the raw Variant item count;
- dimension-selection readback tolerates duplicate annotation selection while
  still rejecting an empty selection;
- the title-region measurement window accepts the measured controlled border;
- redundant title checks and the identity-valued coordinate helper remain
  removed; and
- rejected per-vertex transform diagnostics remain concise.

## Verification levels

- Local MCP: exact method/enum contracts queried for `ICurve.IsCircle`,
  `IFeature.GetTypeName2`, `IFace2.GetSeedFeature`,
  `IDisplayDimension.AutoJogOrdinate`, `IDisplayDimension.DisplayAsChain`,
  `IAnnotation.SetName`, and `swDimensionType_e`.
- Installed interop: SOLIDWORKS 2025 SP1.2 assembly version `33.1.2.4`
  confirms the used member signatures and dimension enum values.
- Retained probe: P-0251 model edges used by the owned-hole path report
  `IsCircle=True`, complete `0..2*pi` spans, and seven-value circle parameters.
- Static/export verification and guarded SWP deployment are recorded in the
  r22 deployment evidence.
- Full VBA Editor compilation, fixture runtime, QA, and visual/manufacturing
  acceptance remain separate user-run gates.

