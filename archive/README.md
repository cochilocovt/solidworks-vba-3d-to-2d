# Archive

Code and tests that are no longer the deployment target, kept because they
encode work that may be worth salvaging.

## `target-spec-hybrid-v2/`

The former trunk (36,013 lines, 38 managed components). Archived 2026-08-05
when `src/baseline-model-dims/` became the trunk.

Reason: it is roughly 16x the size of the trunk and the fixture part number
`P-0251` appears on 94 lines across 14 of its modules, so it could not be
generalised without a rewrite. The baseline snapshot has the cleaner control
flow and a mask that was independently confirmed correct against
`swInsertAnnotation_e`.

Worth salvaging for Tier C, if and when it is attempted:

- `Module12_FeatureQualification` — feature-tree evidence for deciding what
  is genuinely a hole.
- `Module10_SectionDimensionEngine`, `Module17_SectionPath` — section
  dimension requirements and section-path selection.
- `Module13_ProjectionResolution` — model-entity to drawing-entity mapping.
- `Module19_SemanticQA` — semantic checks rather than count-based ones.

Do not deploy from this directory. It is not in the deployment manifest.

## `target-spec-hybrid-v2-tests/`

The 31 companion test files that assert contracts about the code above.
Moved out of `tools/solidworks-automation-companion/tests/` at the same time,
so the live suite covers only what is actually deployed.

Run them from inside this directory:

```bash
python -m unittest discover -s . -q
```

**585 tests run; 22 fail, by design.** Every one of those failures is the
same assertion in a different file — some variant of:

```python
self.assertIn("Module19_SemanticQA", manifest)
```

That assertion means "this component is managed by the deployment manifest."
It was true while target-spec-hybrid-v2 was the trunk. It is now false,
because the manifest points at `src/baseline-model-dims` and these modules
are not deployed. The assertions were left intact rather than weakened: a
suite made green by deleting the assertion that detects the change records
less than a suite that fails honestly and says why.

The other 563 tests pass and remain a usable record of what the archived
implementation was contracted to do.
