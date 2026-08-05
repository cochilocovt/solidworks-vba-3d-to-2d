"""R23 Phase 11 production-pipeline source contracts."""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"


class R23PipelineContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def run_body(self) -> str:
        return self.read("Module2_DrawingPipeline.bas").split(
            "Public Sub RunDrawingPipeline("
        )[1].split("\nEnd Sub")[0]

    def test_graph_replaces_legacy_feature_collection(self):
        body = self.run_body()
        self.assertIn("Dim graph As CLocationGraph", body)
        self.assertIn("BuildFeatureCatalog(", body)
        self.assertNotIn("GetAllHoleLikeFeatures", body)
        self.assertNotIn("AddMissingOrdinateDimensions", body)

    def test_scratch_layout_entrypoint_is_retired_without_mutation(self):
        """User-accepted layout may not be changed by an old entry point."""
        source = self.read("Module2_DrawingPipeline.bas")
        body = source.split("Public Sub R23_ApplyContentLayoutToScratch()")[
            1
        ].split("\nEnd Sub")[0]

        self.assertIn("R23_LAYOUT_MUTATION_SKIPPED", body)
        self.assertIn("UserAcceptedLayoutAsIs", body)
        self.assertNotIn("ApplyR23ContentLayout", body)
        self.assertNotIn("ActiveDoc", body)
        self.assertNotIn(".Save", body)

    def test_user_accepted_layout_cannot_reach_automatic_placement(self):
        source = self.read("Module2_DrawingPipeline.bas")
        body = self.run_body()

        self.assertIn("RecordR23UserAcceptedLayout(evidence)", body)
        self.assertNotIn("ApplyR23ContentLayout(swDrawModel, swDraw", body)
        self.assertEqual(source.count("ApplyR23ContentLayout("), 1)
        self.assertIn("R23_LAYOUT_USER_ACCEPTED_AS_IS", self.read(
            "Module1_Main.bas"))

    def test_production_order_matches_r23_1101(self):
        body = self.run_body()
        stages = (
            "BuildFeatureCatalog(",
            "CreateConfiguredViews(",
            "ArrangeViewsInMeasuredGrid(",
            "BuildAllViewProjections(",
            "CreateSemanticPrimarySection(",
            "ImportAndReconcileR23Annotations(",
            "CreateR23OrdinateGroups(",
            "ReconcileR23SectionDimensions(",
            "CreateMissingR23Callouts(",
            "CreateIsometricView(",
            "PopulateTitleBlock",
            "AutoArrangeAllDrawingDimensions",
            "RecordR23UserAcceptedLayout(",
            "EvaluateSemanticDrawing",
        )
        positions = [body.index(stage) for stage in stages]
        self.assertEqual(positions, sorted(positions))

    def test_final_content_is_sealed_before_semantic_qa(self):
        body = self.run_body()
        tail = body.split("RecordR23UserAcceptedLayout(", 1)[1]
        for forbidden in (
            "ImportModelAnnotations(",
            "CreateR23OrdinateGroups(",
            "CreateMissingR23Callouts(",
            "PopulateTitleBlock",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, tail)

    def test_legacy_section_and_literal_callout_writers_are_unreachable(self):
        body = self.run_body()
        self.assertNotIn("CreatePrimarySection(", body)
        self.assertNotIn("AddRequiredManufacturingDefinitions(", body)
        self.assertNotIn("Module5_FallbackDimensionEngine", body)

    def test_generic_arranger_excludes_section_dimensions(self):
        importer = self.read("Module4_ModelItemImporter.bas")
        body = importer.split("Public Sub AutoArrangeAllDrawingDimensions(")[1]
        self.assertIn("IsExcludedFromGenericArrangement", body)
        self.assertIn("reason=SectionSemanticLanes", body)

    def test_import_order_contains_each_eligible_view_once(self):
        source = self.read("Module2_DrawingPipeline.bas")
        body = source.split("Private Function CollectR23ImportViews(")[1]
        body = body.split("\nEnd Function")[0]
        self.assertIn("AddR23ImportViewOnce", body)
        self.assertIn("If swView.Type = 2 Then GoTo ContinueOther", body)
        helper = source.split("Private Sub AddR23ImportViewOnce(")[1]
        helper = helper.split("\nEnd Sub")[0]
        self.assertIn("seenViews.Exists(viewName)", helper)

    def test_section_path_guard_does_not_dereference_nothing(self):
        source = self.read("Module2_DrawingPipeline.bas")
        self.assertNotIn("path Is Nothing Or Not path.Resolved", source)

    def test_final_qa_requires_semantic_stages_not_legacy_counts(self):
        qa = self.read("Module6_QAEngine.bas")
        core = qa.split("Private Sub RequireCoreStages(")[1].split(
            "\nEnd Sub"
        )[0]
        self.assertIn("Module19_SemanticQA.RequireSemanticStages", core)
        for legacy in (
            'evidence.RequireStage "MODEL_ANNOTATIONS"',
            'evidence.RequireStage "PHYSICAL_LOCATIONS"',
            'evidence.RequireStage "ORDINATE_COVERAGE"',
        ):
            with self.subTest(legacy=legacy):
                self.assertNotIn(legacy, core)

    def test_r23_1203_rejects_retired_definition_and_section_paths(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        qa = self.read("Module6_QAEngine.bas")
        title = self.read("Module7_TitleBlockEngine.bas")

        for retired in (
            "AddRequiredManufacturingDefinitions",
            "FindManufacturingCandidate",
            "P0251ManufacturingDefinitionIsVisible",
            "AssociatedCalloutContainsTokens",
        ):
            with self.subTest(retired=retired):
                self.assertNotIn(retired, pipeline + qa + title)

        p0251_section = pipeline.split("Private Function CreatePrimarySection(")[
            1
        ].split('Case "P-0252-01-001"', 1)[0]
        self.assertNotIn("topY + extension", p0251_section)
        self.assertNotIn("bottomY - extension", p0251_section)

    def test_r23_1203_uses_one_shared_semantic_production_judge(self):
        pipeline = self.run_body()
        semantic = self.read("Module19_SemanticQA.bas")

        self.assertIn("EvaluateSemanticDrawing", pipeline)
        self.assertIn("BuildAnnotationInventory", semantic)
        self.assertIn("ReconcileWithLocationGraph", semantic)
        self.assertNotIn("ImportedAnnotations > 0", self.read("Module6_QAEngine.bas"))


if __name__ == "__main__":
    unittest.main()
