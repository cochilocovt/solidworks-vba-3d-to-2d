"""R23 Phase 10 source contracts for semantic QA.

Static source-contract tests for tasks R23-1000 through R23-1017. The
load-bearing ones are R23-1010 (coverage is per view and per category, not a
nonzero count), R23-1011 (section dimensions are judged by inspecting
dimensions, never by searching a note for text), R23-1012 (every
manufacturing field carries its source) and R23-1016 (duplicate keys fail).
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module19_SemanticQA.bas"

STAGES = (
    "MODEL_INTENT_CATALOG",
    "MODEL_IMPORT_COVERAGE",
    "NATIVE_CALLOUT_COVERAGE",
    "PHYSICAL_LOCATION_GRAPH",
    "VIEW_PROJECTION",
    "ORDINATE_SCHEME",
    "SECTION_GEOMETRY",
    "SECTION_DIMENSIONS",
    "MANUFACTURING_DEFINITION",
    "FINAL_LAYOUT",
)


class R23SemanticQaContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line
            for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def body(self, marker: str, terminator: str = "\nEnd Function") -> str:
        return self.source.split(marker)[1].split(terminator)[0]

    def setUp(self):
        self.source = self.read(MODULE)
        self.executable = self.code(MODULE)

    def test_component_exists_and_is_managed(self):
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")
        self.assertTrue((SOURCE / MODULE).exists())
        self.assertIn("Module19_SemanticQA", manifest)

    # --- read-only boundary ----------------------------------------------

    def test_the_whole_module_changes_nothing(self):
        """Phase 10 judges what the earlier phases produced. A QA engine
        that repairs what it is judging cannot report on it."""
        for forbidden in (
            "AddDimension2",
            "CreateLine",
            "CreateSectionViewAt5",
            ".Position =",
            "ScaleDecimal =",
            "SetFitValues",
            "SetValues2",
            "EditRebuild3",
            "ClearSelection2 False",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

    def test_probe_does_not_invoke_the_production_qa_gate(self):
        self.assertNotIn("EmitRunEvidence", self.executable)

    def test_native_definition_scan_skips_non_display_dimension_items(self):
        """A malformed COM array item must not abort semantic QA before
        the remaining native-callout definitions are evaluated."""
        body = self.body("Public Function CollectRetainedDefinitions(")
        self.assertIn("Set swDispDim = Nothing", body)
        self.assertIn("On Error Resume Next", body)
        self.assertIn("If swDispDim Is Nothing Then GoTo ContinueDimension",
                      body)

    def test_the_probe_reports_the_read_only_boundary(self):
        body = self.source.split("Public Sub R23_ProbeSemanticQA()")[1]
        self.assertIn("mode=ReadOnly", body)
        self.assertIn("creations=0|mutations=0", body)
        self.assertIn("drawingUnchanged=", body)
        self.assertIn("initialSelectionCount=", body)
        self.assertIn("finalSelectionCount=", body)

    # --- R23-1000 to R23-1009 --------------------------------------------

    def test_every_required_stage_is_declared(self):
        body = self.body("Public Sub RequireSemanticStages(", "\nEnd Sub")
        self.assertEqual(len(STAGES), body.count("evidence.RequireStage"))
        for stage in STAGES:
            with self.subTest(stage=stage):
                self.assertIn(stage, self.executable)

    def test_declaring_a_stage_is_what_gates_probe_and_production(self):
        """A stage nothing proves has to become a named failure, not an
        absence nobody notices. SealRequiredStages is what does that."""
        probe = self.source.split("Public Sub R23_ProbeSemanticQA()")[1]
        body = self.body("Public Function EvaluateSemanticDrawing(")
        self.assertIn("RequireSemanticStages evidence", probe)
        self.assertIn("RequireSemanticStages evidence", body)
        self.assertIn("evidence.SealRequiredStages", body)
        self.assertLess(
            body.index("RequireSemanticStages evidence"),
            body.index("evidence.SealRequiredStages"),
        )

    def test_each_stage_has_an_evaluator_that_can_fail_it(self):
        for evaluator in (
            "EvaluateModelIntentCatalog",
            "EvaluateModelImportCoverage",
            "EvaluateNativeCalloutCoverage",
            "EvaluatePhysicalLocationGraph",
            "EvaluateViewProjection",
            "EvaluateOrdinateScheme",
            "EvaluateSectionGeometry",
            "EvaluateSectionDimensions",
            "EvaluateManufacturingDefinition",
            "EvaluateFinalLayout",
        ):
            with self.subTest(evaluator=evaluator):
                body = self.body("Public Function %s(" % evaluator)
                self.assertIn("MarkStageProved", body)
                self.assertIn("MarkStageFailed", body)

    def test_every_evaluator_is_called_by_the_shared_judge(self):
        body = self.body("Public Function EvaluateSemanticDrawing(")
        for evaluator in (
            "EvaluateModelIntentCatalog graph, evidence",
            "EvaluatePhysicalLocationGraph graph, evidence",
            "EvaluateViewProjection graph, evidence",
            "EvaluateModelImportCoverage graph, evidence",
            "EvaluateNativeCalloutCoverage definitions, evidence",
            "EvaluateManufacturingDefinition definitions, evidence",
            "EvaluateOrdinateScheme schemes, evidence",
            "EvaluateSectionGeometry bestPath, evidence",
            "EvaluateExistingSectionDimensions views, evidence",
            "EvaluateCurrentLayout swSheet, views, evidence",
        ):
            with self.subTest(evaluator=evaluator):
                self.assertIn(evaluator, body)

    def test_import_coverage_is_built_and_reconciled_before_evaluation(self):
        """Projection flags are not import evidence by themselves. The
        semantic graph must first receive existing annotations and reconcile
        them to its own projection objects before directional coverage runs."""
        body = self.body("Public Function EvaluateSemanticDrawing(")
        self.assertIn("BuildAnnotationInventory(", body)
        self.assertIn("ReconcileWithLocationGraph(", body)
        self.assertIn("SEMANTIC_ANNOTATION", body)
        self.assertIn("graph.AnnotationCount()", body)
        self.assertNotIn("graph.ImportedAnnotations", body)
        self.assertLess(
            body.index("BuildAnnotationInventory("),
            body.index("EvaluateModelImportCoverage graph, evidence"),
        )
        self.assertLess(
            body.index("ReconcileWithLocationGraph("),
            body.index("EvaluateModelImportCoverage graph, evidence"),
        )
        self.assertLess(
            body.index("EvaluateModelImportCoverage graph, evidence"),
            body.index("CollectOrdinateSchemes("),
        )

    def test_production_pipeline_uses_the_same_semantic_judge(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        self.assertIn("Module19_SemanticQA.EvaluateSemanticDrawing", pipeline)
        self.assertLess(
            pipeline.index("RecordR23UserAcceptedLayout"),
            pipeline.index("Module19_SemanticQA.EvaluateSemanticDrawing"),
        )

    # --- R23-1010 --------------------------------------------------------

    def test_r23_1010_coverage_is_per_view_and_per_category(self):
        """A nonzero import count is satisfied by importing everything into
        one view and nothing into the others."""
        body = self.body("Public Function EvaluateModelImportCoverage(")
        self.assertIn("graph.ProjectionsForView(", body)
        self.assertIn("QA_IMPORT_COVERAGE|view=", body)
        for category in ("coveredX=", "coveredY=", "annotated="):
            with self.subTest(category=category):
                self.assertIn(category, body)

        self.assertIn("ViewImportedNothing:", body)
        self.assertIn("NoViewHasAcceptedProjections", body)
        self.assertNotIn("SchemeCovered", body)

    # --- R23-1011 --------------------------------------------------------

    def test_r23_1011_section_dimensions_are_inspected_not_text_matched(
        self,
    ):
        """The check this replaces searched a note for a stepped-bore token.
        That passes on free text that has drifted from the geometry and
        fails on a correct drawing whose wording differs."""
        body = self.body("Public Function EvaluateSectionDimensions(")
        self.assertIn(
            "Module10_SectionDimensionEngine.VerifySectionDimensions(",
            body,
        )
        for forbidden in ("GetText", "InStr(1, note", "GetNotes"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

    # --- R23-1012 --------------------------------------------------------

    def test_r23_1012_every_manufacturing_field_emits_its_source(self):
        body = self.body("Public Function EvaluateNativeCalloutCoverage(")
        for field in (
            "diameterFrom=",
            "depthFrom=",
            "endConditionFrom=",
            "quantityFrom=",
            "attachmentFrom=",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

        self.assertIn("NoProvenance:", body)

    def test_a_blank_or_placeholder_source_is_not_provenance(self):
        """A number whose source is empty, "None" or "Unproven" has no
        authority behind it, and those three are the same failure."""
        body = self.body("Private Function SourceIsReal(")
        self.assertIn("Len(trimmed) = 0", body)
        self.assertIn("PROOF_NONE", body)
        self.assertIn("PROOF_UNPROVEN", body)

        gaps = self.body("Private Function ProvenanceGaps(")
        for field in ("diameter", "depth", "counterBore", "thread",
                      "quantity"):
            with self.subTest(field=field):
                self.assertIn('"%s"' % field, gaps)

    # --- R23-1013 and R23-1014 -------------------------------------------

    def test_r23_1013_raw_and_effective_type_for_every_feature(self):
        """Every feature, not every accepted feature: a rejection with no
        type recorded cannot be reviewed."""
        body = self.body("Public Function EvaluateModelIntentCatalog(")
        self.assertIn("For i = 1 To features.Count", body)
        for field in (
            "rawTypeName2=",
            "rawTypeName=",
            "effectiveType=",
            "resolutionSource=",
            "rejection=",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

        # The emit is unconditional - not inside an Accepted test.
        emit = body.split("QA_FEATURE_TYPE")[0]
        self.assertNotIn("If definition.Accepted Then\n            Emit",
                         emit)

    def test_r23_1014_an_unresolved_ice_feature_fails(self):
        body = self.body("Private Function TypeResolutionFailed(")
        self.assertIn("Len(Trim$(definition.EffectiveType)) = 0", body)
        for token in (
            "RESOLUTION_ICE_UNRESOLVED",
            "RESOLUTION_UNRESOLVED",
            "RESOLUTION_READ_ERROR",
        ):
            with self.subTest(token=token):
                self.assertIn(token, body)

        self.assertIn(
            'RESOLUTION_ICE_UNRESOLVED As String = "IceUnresolved"',
            self.executable,
        )

    def test_the_ice_tokens_match_what_module12_actually_writes(self):
        module12 = self.read("Module12_FeatureQualification.bas")
        for token in ("IceUnresolved", "Unresolved", "ReadError"):
            with self.subTest(token=token):
                self.assertIn('"%s"' % token, module12)

    # --- R23-1015 --------------------------------------------------------

    def test_r23_1015_a_location_without_a_projection_fails_by_name(self):
        """Reporting the projection COUNT hides an unprojected location
        behind the ones that did project."""
        body = self.body("Public Function EvaluateViewProjection(")
        self.assertIn("NoProvedProjection:", body)
        self.assertIn("locationsWithout=", body)
        self.assertIn("projection.Accepted", body)

    # --- R23-1016 --------------------------------------------------------

    def test_r23_1016_duplicate_keys_fail_in_all_four_kinds(self):
        self.assertIn("Public Function DuplicateKeyReport(", self.source)
        for caller, kind in (
            ("Public Function EvaluatePhysicalLocationGraph(",
             '"physical"'),
            ("Public Function EvaluateNativeCalloutCoverage(",
             '"familyDefinition"'),
            ("Public Function EvaluateSectionDimensions(",
             '"sectionRequirement"'),
        ):
            with self.subTest(kind=kind):
                self.assertIn(kind, self.body(caller))

    def test_no_duplicates_and_check_did_not_run_are_different(self):
        """Returning an empty string for "no duplicates" makes the two
        indistinguishable at every call site."""
        body = self.body("Public Function DuplicateKeyReport(")
        self.assertIn('duplicates = "None"', body)
        self.assertIn("DuplicateCheckError:", body)

    # --- R23-1017 --------------------------------------------------------

    def test_r23_1017_unavailable_section_or_envelope_fails(self):
        section = self.body("Public Function EvaluateSectionGeometry(")
        self.assertIn("If path Is Nothing Then", section)
        self.assertIn("If Not path.Resolved Then", section)
        self.assertIn("sectionFailures=None", section)

        layout = self.body("Public Function EvaluateFinalLayout(")
        self.assertIn("R23_LAYOUT_USER_ACCEPTED_AS_IS", layout)
        self.assertIn("automaticClearance=DeferredByUser", layout)
        self.assertLess(
            layout.index("R23_LAYOUT_USER_ACCEPTED_AS_IS"),
            layout.index("If envelopes.Count = 0"),
        )
        self.assertIn("EnvelopeUnavailable:", layout)
        self.assertIn("clearanceFailures=None", layout)

    def test_user_accepted_layout_is_reported_as_a_waiver(self):
        body = self.body("Private Sub EvaluateCurrentLayout(", "\nEnd Sub")
        self.assertIn("R23_LAYOUT_USER_ACCEPTED_AS_IS", body)
        self.assertIn("EvaluateFinalLayout deferredEnvelopes", body)
        self.assertLess(
            body.index("R23_LAYOUT_USER_ACCEPTED_AS_IS"),
            body.index("MeasureSheetRegions("),
        )

    # --- shared decision logic -------------------------------------------

    def test_definition_decisions_come_from_module16_not_a_copy(self):
        """What counts as a native callout, which family it belongs to and
        which definition is retained are Module16's judgements. Only the
        loop is local, so the two cannot drift on the part that matters."""
        body = self.body("Public Function CollectRetainedDefinitions(")
        for call in (
            "Module16_CalloutDefinition.IsNativeHoleCallout(",
            "Module16_CalloutDefinition.MatchCalloutToFamily(",
            "Module16_CalloutDefinition.BuildDefinitionFromTypedData(",
            "Module16_CalloutDefinition.ReadNativeCalloutFields",
            "Module16_CalloutDefinition.RetainDefinitionForFamily(",
        ):
            with self.subTest(call=call):
                self.assertIn(call, body)

    def test_the_stage_report_names_every_stage_and_its_status(self):
        body = self.body("Public Function StageReport(")
        for stage in ("STAGE_MODEL_INTENT_CATALOG", "STAGE_FINAL_LAYOUT"):
            with self.subTest(stage=stage):
                self.assertIn(stage, body)
        self.assertIn("evidence.StageIsProved(stageName)", body)
        self.assertIn("proved=", body)

    def test_evidence_values_are_sanitized(self):
        """Feature and family keys carry the delimiters this format uses;
        an unsanitized one corrupts every field after it on the line."""
        self.assertIn("Public Function QaToken(", self.source)
        body = self.body("Public Function EvaluateModelIntentCatalog(")
        self.assertIn("QaToken(definition.FeatureName)", body)

    # --- hygiene ---------------------------------------------------------

    def test_source_hygiene(self):
        raw = (SOURCE / MODULE).read_bytes()
        self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))
        self.assertTrue(all(b < 0x80 for b in raw))
        text = raw.decode("cp1252")
        self.assertTrue(text.startswith("Option Explicit"))
        for line in text.split("\r\n"):
            self.assertLessEqual(len(line), 79)

    def test_source_revision_identifies_wired_r23_pipeline(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
