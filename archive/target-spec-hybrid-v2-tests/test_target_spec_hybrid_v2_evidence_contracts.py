from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"


class TargetSpecHybridV2EvidenceContractTests(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="utf-8-sig")

    def test_required_stage_ledger_is_fail_closed_and_machine_readable(self):
        evidence = self.read("CRunEvidence.cls")
        for contract in (
            "RequireStage",
            "MarkStageProved",
            "MarkStageFailed",
            "SealRequiredStages",
            "STAGE_GATE|stage=",
            '"- STAGE|name="',
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, evidence)
        self.assertIn('StrComp(StageStatusText(key), "FAILED"', evidence)

    def test_physical_and_projection_location_ledgers_are_distinct(self):
        evidence = self.read("CRunEvidence.cls")
        qa = self.read("Module6_QAEngine.bas")
        semantic = self.read("Module19_SemanticQA.bas")
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        for contract in (
            "RegisterPhysicalLocation",
            "PHYSICAL_LOCATION|key=",
            "AddProjectionLocationRecord",
            "PROJECTION_LOCATION|physical_key=",
            "UniquePhysicalLocationCount",
            "ProjectionLocationRecordCount",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, evidence)
        self.assertIn("EvaluatePhysicalLocationGraph graph, evidence", semantic)
        self.assertNotIn("evidence.UniquePhysicalLocationCount < 10", qa)
        self.assertIn("evidence.RegisterPhysicalLocation", ordinate)
        self.assertIn("evidence.AddProjectionLocationRecord", ordinate)
        self.assertIn("BuildCanonicalModelCentreKey", ordinate)
        self.assertIn("BuildCanonicalModelAxisKey", ordinate)

    def test_sheet_preflight_and_actual_scale_readback_fail_closed(self):
        runtime = self.read("Module8_RuntimeSupport.bas")
        for contract in (
            'RequireStage "CONTROLLED_SHEET"',
            "GetSize(sheetWidth, sheetHeight)",
            "evidence.SheetWidth = sheetWidth",
            "evidence.SheetHeight = sheetHeight",
            "swSheet.SheetFormatVisible = True",
            'EnsureSheetFormatVisible',
            "TryMeasureLegacyControlledTitleBlock",
            "GetTemplateSketch",
            "GetSketchSegments",
            "GetZoneMargin",
            "GetProperties2",
            "ActualScaleNumerator",
            "ActualScaleDenominator",
            "SheetScaleReadbackProven",
            'MarkStageFailed "SHEET_SCALE"',
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, runtime)
        self.assertNotIn(
            "GetSize(evidence.SheetWidth, evidence.SheetHeight)", runtime
        )

    def test_layout_never_silently_changes_independent_view_scale(self):
        layout = self.read("Module9_LayoutEngine.bas")
        self.assertNotIn("ScaleDecimal =", layout)
        self.assertNotIn("UseSheetScale = 0", layout)
        self.assertIn("unapproved independent scale", layout)
        self.assertIn("RebuildDocumentVerified", layout)

    def test_title_evidence_proves_source_link_rendered_value_and_extent(self):
        title = self.read("Module7_TitleBlockEngine.bas")
        evidence = self.read("CRunEvidence.cls")
        for contract in (
            "titleBlock.GetNotes",
            "PropertyLinkedText",
            "LinkedTextReferencesProperty",
            "note.GetExtent",
            "NoteExtentIsInRegion",
            "AddTitlePropertyEvidence",
            '"model-configuration"',
            '"model-document"',
            '"sheet:GetProperties2"',
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, title + evidence)

    def test_final_qa_is_read_only_and_enforces_view_and_extent_policy(self):
        qa = self.read("Module6_QAEngine.bas")
        perform = qa[
            qa.index("Public Sub PerformFinalDrawingChecks") :
            qa.index("Private Sub RequireStructuralQaStages")
        ]
        for mutation in (
            "ActivateSheet",
            "ActivateView",
            "ClearSelection",
            "SetPickMode",
            "EditRebuild",
            "ViewZoom",
            ".Position =",
        ):
            with self.subTest(mutation=mutation):
                self.assertNotIn(mutation, perform)
        for contract in (
            "isometric/axonometric view",
            "unsupported view",
            "GetLeaderCount",
            "GetLeaderPointsAtIndex",
            "note.GetExtent",
            "measured title-block",
            "ContentBorderBottom",
            "ANNOTATION_GEOMETRY",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, qa)

    def test_evidence_write_is_atomic_and_exactly_read_back(self):
        qa = self.read("Module6_QAEngine.bas")
        for contract in (
            "QA_REPORT.tmp",
            "CreateTextFile",
            "OpenTextFile",
            "ReadAll",
            "vbBinaryCompare",
            "Name temporaryPath As reportPath",
            "EvidenceWriteVerified",
            'Format$(suffix, "000")',
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, qa)
        self.assertNotIn("Timer", qa)


if __name__ == "__main__":
    unittest.main()
