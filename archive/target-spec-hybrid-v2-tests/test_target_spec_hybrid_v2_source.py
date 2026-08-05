from pathlib import Path
import re
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"


class TargetSpecHybridV2SourceTests(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="utf-8-sig")

    def test_complete_importable_source_set_exists(self):
        required = {
            *(f"Module{i}_{name}.bas" for i, name in {
                1: "Main",
                2: "DrawingPipeline",
                3: "ModelAudit",
                4: "ModelItemImporter",
                5: "FallbackDimensionEngine",
                6: "QAEngine",
                7: "TitleBlockEngine",
                8: "RuntimeSupport",
                9: "LayoutEngine",
            }.items()),
            "CHoleCandidate.cls",
            "CDatumProof.cls",
            "CRunEvidence.cls",
            "UserForm1.frm",
            "UserFormSection.frm",
        }
        self.assertEqual(set(), required - {p.name for p in SOURCE.iterdir()})

    def test_all_code_components_use_option_explicit(self):
        for path in SOURCE.iterdir():
            if path.suffix.lower() in {".bas", ".cls", ".frm"}:
                with self.subTest(path=path.name):
                    self.assertIn("Option Explicit", path.read_text(encoding="utf-8-sig"))

    def test_vba_procedure_blocks_and_continuations_are_structurally_safe(self):
        start = re.compile(
            r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
            r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",
            re.IGNORECASE,
        )
        end = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.IGNORECASE)

        for path in SOURCE.iterdir():
            if path.suffix.lower() not in {".bas", ".cls", ".frm"}:
                continue

            current = None
            continuation_count = 0
            for line_number, line in enumerate(
                path.read_text(encoding="utf-8-sig").splitlines(), 1
            ):
                with self.subTest(path=path.name, line=line_number):
                    self.assertLessEqual(len(line), 1023)

                match = start.match(line)
                if match:
                    self.assertIsNone(
                        current,
                        f"{path.name}:{line_number} starts {match.group(2)} "
                        f"before {current} is closed",
                    )
                    current = (match.group(1), match.group(2), line_number)

                match = end.match(line)
                if match:
                    self.assertIsNotNone(
                        current,
                        f"{path.name}:{line_number} has orphan {match.group(0)}",
                    )
                    expected = (
                        "Property"
                        if current[0].lower().startswith("property")
                        else current[0]
                    )
                    self.assertEqual(expected.lower(), match.group(1).lower())
                    current = None

                if line.rstrip().endswith(" _"):
                    continuation_count += 1
                    self.assertLessEqual(
                        continuation_count,
                        24,
                        f"{path.name}:{line_number} exceeds VBA continuation limit",
                    )
                else:
                    continuation_count = 0

            self.assertIsNone(current, f"{path.name} has an unclosed procedure")

    def test_no_blank_view_activation_or_nothing_visible_entity_component(self):
        combined = "\n".join(
            p.read_text(encoding="utf-8-sig")
            for p in SOURCE.glob("*.bas")
        )
        self.assertNotRegex(combined, r'ActivateView\s*\(?\s*""')
        self.assertNotIn("GetVisibleEntities2(Nothing", combined)
        self.assertIn("ActivateDoc3", self.read("Module8_RuntimeSupport.bas"))
        self.assertIn("ActivateSheet", self.read("Module8_RuntimeSupport.bas"))

    def test_importer_uses_verified_2025_contract(self):
        source = self.read("Module4_ModelItemImporter.bas")
        self.assertIn("InsertModelAnnotations4", source)
        self.assertIn("swInsertDimensionsMarkedForDrawing", source)
        self.assertIn("swInsertDatums", source)
        self.assertIn("swInsertHoleWizardProfileDimensions", source)
        self.assertIn("swInsertHoleWizardLocationDimensions", source)
        self.assertRegex(
            source,
            re.compile(
                r"InsertModelAnnotations4\([^)]*False,\s*True,",
                re.DOTALL,
            ),
        )
        self.assertNotRegex(
            source,
            re.compile(
                r"InsertModelAnnotations4\([^)]*True,\s*True,",
                re.DOTALL,
            ),
        )
        self.assertIn("IsOrdinateEligibleView", source)
        self.assertIn("Set activeView = swDraw.ActiveDrawingView", source)
        self.assertIn('"DRAWINGVIEW"', source)
        self.assertIn("selectedView = swDrawModel.Extension.SelectByID2", source)
        self.assertIn("MODEL_IMPORT_VIEW_SELECTION", source)
        self.assertIn("MODEL_IMPORT_EXECUTE", source)
        self.assertIn("If selectedView Then", source)
        self.assertNotIn("If Not selectedView Then", source)
        self.assertIn("active-view readback matched, so continuing", source)
        self.assertNotIn(
            "failed to select drawing view '", source
        )
        selection_offset = source.index(
            "selectedView = swDrawModel.Extension.SelectByID2"
        )
        import_offset = source.index("inserted = swDraw.InsertModelAnnotations4")
        cleanup_offset = source.index(
            "swDrawModel.ClearSelection2 True", import_offset
        )
        self.assertLess(selection_offset, import_offset)
        self.assertLess(import_offset, cleanup_offset)
        self.assertIn("Model import skipped unsupported or unregistered view", source)
        self.assertIn("ImportModelItemsPerView", source)

    def test_ordinate_candidates_require_model_feature_ownership(self):
        source = self.read("Module5_FallbackDimensionEngine.bas")
        audit_source = self.read("Module3_ModelAudit.bas")
        transform_source = source + self.read("Module8_RuntimeSupport.bas")
        for contract in (
            "GetVisibleComponents",
            "modelHoleFeatures",
            "ownerFeature.GetFaces",
            "cylinderFace.GetEdges",
            "swView.GetCorrespondingEntity(modelEdge)",
            "GetStartVertex",
            "GetEndVertex",
            "TryReadClosedCircularEdge",
            "CylinderParams",
            "IsOwnedHoleFeature",
            "ModelToViewTransform",
            "MultiSelect2",
            "AddOrdinateDimension",
            "SetPickMode",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, transform_source)
        self.assertNotIn("GetVisibleEntities2", source)
        self.assertNotIn("GetCorrespondingEntity2", source)
        self.assertNotIn("GetSketchPoints", source)
        self.assertNotIn("nearest", source.lower())
        self.assertIn("swView.ReferencedConfiguration", source)
        self.assertNotIn("component.ReferencedConfiguration", source)
        self.assertIn("swSpecifyConfiguration", audit_source)
        self.assertIn("IsFeatureActiveInConfiguration", audit_source)

    def test_ownership_maps_audited_model_edges_into_drawing_view(self):
        source = self.read("Module5_FallbackDimensionEngine.bas")
        pipeline = self.read("Module2_DrawingPipeline.bas")
        projection = self.read("Module13_ProjectionResolution.bas")
        start = source.index("Private Function BuildOwnedCandidate")
        end = source.index("Private Sub CaptureFeatureSemantics", start)
        candidate_builder = source[start:end]

        self.assertNotIn("Set modelEdge = drawingEdge", candidate_builder)
        self.assertIn("candidate.ModelEntityAliases.Add modelEdge", candidate_builder)
        self.assertIn(
            "Set mappedObject = swView.GetCorrespondingEntity(modelEdge)",
            source,
        )
        self.assertIn(
            "|mapping=IView.GetCorrespondingEntity",
            source,
        )
        self.assertIn("AuditedFeature.GetFaces/Face2.GetEdges", source)
        self.assertIn("Dim graph As CLocationGraph", pipeline)
        self.assertIn("BuildAllViewProjections", pipeline)
        self.assertIn("GetCorrespondingEntity", projection)
        self.assertIn(
            "ENTITY_SOURCE|view=",
            source,
        )

    def test_datum_and_qa_paths_fail_closed(self):
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        qa = self.read("Module6_QAEngine.bas")
        main = self.read("Module1_Main.bas")
        self.assertIn("DatumNotSelectable", ordinate)
        self.assertIn("simultaneous minimum-X", ordinate)
        self.assertIn("IsAuthorizedFixture", main)
        self.assertIn("RESULT: FAIL", self.read("CRunEvidence.cls"))
        self.assertIn("RequireSemanticStages", qa)
        self.assertNotIn("zero visible display dimensions", qa)

    def test_ordinate_error_decode_matches_installed_2025_enum(self):
        source = self.read("Module5_FallbackDimensionEngine.bas")
        self.assertIn('Case 1: DecodeOrdinateResult = "1 GeneralFailure"', source)
        self.assertIn('Case 7: DecodeOrdinateResult = "7 OrdinateFailure"', source)
        self.assertIn('Case 8: DecodeOrdinateResult = "8 Duplicate"', source)
        self.assertIn('Case 9: DecodeOrdinateResult = "9 BadDirection"', source)

    def test_directional_datum_is_typed_and_retains_coordinate_evidence(self):
        datum = self.read("CDatumProof.cls")
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        self.assertIn("DrawingEntity As SldWorks.Entity", datum)
        self.assertIn("ModelEntity As SldWorks.Entity", datum)
        for field in (
            "StableKey",
            "ModelX",
            "ModelY",
            "ModelZ",
            "ViewX",
            "ViewY",
            "SheetX",
            "SheetY",
            "SelectionSucceeded",
        ):
            self.assertIn(field, datum)
        self.assertIn("GetDatumOriginForDirection", ordinate)
        self.assertIn("datum.DrawingEntity.Select4", ordinate)
        self.assertIn("EVIDENCE|DATUM_PROOF", ordinate)

    def test_ownership_uses_face_set_and_stable_physical_instance(self):
        audit = self.read("Module3_ModelAudit.bas")
        candidate = self.read("CHoleCandidate.cls")
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        self.assertIn("FeatureContainsFace", audit)
        self.assertIn("targetFace.IsSame(candidateFace)", audit)
        self.assertIn("AuditedFeatureFaceSet", ordinate)
        self.assertIn(
            "AuditedFeatureDoesNotContainCylinderFace", ordinate
        )
        self.assertIn("CylinderFaceMatchesRadius", ordinate)
        self.assertIn("ModelAxisX", candidate)
        self.assertIn("FamilyKey", candidate)
        self.assertIn("PhysicalInstanceKey", candidate)
        self.assertIn("BuildPhysicalInstanceKey", ordinate)

    def test_datum_vertices_are_mapped_model_first(self):
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        self.assertIn("CollectModelFirstVertexCoordinates", ordinate)
        self.assertIn("swPart.GetBodies2(swSolidBody, True)", ordinate)
        self.assertIn("body.GetEdges", ordinate)
        self.assertIn("modelEdge.GetStartVertex", ordinate)
        self.assertIn("modelEdge.GetEndVertex", ordinate)
        self.assertIn(
            "Set mappedObject = swView.GetCorrespondingEntity(modelVertex)",
            ordinate,
        )
        self.assertNotIn("CollectVisibleVertexCoordinates", ordinate)

    def test_final_qa_counts_exact_view_dimension_arrays(self):
        qa = self.read("Module6_QAEngine.bas")
        start = qa.index("Private Sub CountViewDimensions")
        end = qa.index("Private Function IsOrdinateDimensionType", start)
        counter = qa[start:end]
        self.assertIn("swView.GetDisplayDimensions", counter)
        self.assertIn("dimensionCount = 0", counter)
        self.assertIn("ordinateCount = 0", counter)
        self.assertNotIn("GetFirstDisplayDimension", counter)
        self.assertNotIn("GetNext5", counter)

    def test_mass_note_link_repair_is_unique_and_read_back(self):
        title = self.read("Module7_TitleBlockEngine.bas")
        self.assertIn("PrepareUniqueMassLink", title)
        self.assertIn("VerifyPreparedMassLink", title)
        self.assertIn('drawingPropertyLink = "$PRP:"', title)
        self.assertIn("If matchCount <> 1 Then", title)
        self.assertIn("INote.PropertyLinkedText(Mass)", title)
        self.assertIn("status=PROVED", title)

    def test_diagnostic_note_extent_skip_cannot_claim_acceptance(self):
        qa = self.read("Module6_QAEngine.bas")
        layout = self.read("Module9_LayoutEngine.bas")
        for source in (qa, layout):
            self.assertIn("NOTE_EXTENT_SKIPPED", source)
            self.assertIn("Not evidence.LayoutBoundariesProven", source)
            self.assertIn("acceptance=False", source)

    def test_coverage_is_fail_closed_and_family_datum_scoped(self):
        importer = self.read("Module4_ModelItemImporter.bas")
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        self.assertIn(
            "Public Function ApplyImportedCoverageToCandidates", importer
        )
        self.assertIn("swDimensionType_HorLinear", importer)
        self.assertIn("swDimensionType_VertLinear", importer)
        self.assertIn("TryGetDimensionAttachments", importer)
        self.assertIn("CoverageInspectionSucceeded", importer)
        self.assertIn("EVIDENCE|COVERAGE_RESULT", importer)
        self.assertIn("BuildOrdinateGroupKey", ordinate)
        self.assertIn("candidate.FamilyKey", ordinate)
        self.assertIn("datum.StableKey", ordinate)
        self.assertIn("EVIDENCE|COORDINATE_SUPPRESSED", ordinate)
        self.assertIn("EVIDENCE|ORDINATE_GROUP", ordinate)

    def test_controlled_sheet_and_linked_title_paths_fail_closed(self):
        main = self.read("Module1_Main.bas")
        runtime = self.read("Module8_RuntimeSupport.bas")
        title = self.read("Module7_TitleBlockEngine.bas")
        self.assertNotIn("GetUserPreferenceStringValue", main)
        self.assertIn("GetSheetFormatName", runtime)
        self.assertIn("GetZoneMargin", runtime)
        self.assertIn("titleBlock.GetExtents", runtime)
        self.assertIn("TryMeasureLegacyControlledTitleBlock", runtime)
        self.assertIn(
            'Array("Description", "PartName", "Part Name")', title
        )
        populate_title = title[
            title.index("Public Sub PopulateTitleBlock") :
            title.index("Private Function EnsureAssociativeManufacturingCallout")
        ]
        self.assertNotRegex(
            populate_title,
            r"(?im)^\s*(?:Set\s+\w+\s*=\s*)?.*\.InsertNote\b",
        )
        self.assertNotIn("AddRequiredManufacturingDefinitions", title)
        self.assertNotIn("<MOD-DIAM>47 H7", title)
        self.assertIn("Module16_CalloutDefinition", self.read("Module2_DrawingPipeline.bas"))
        self.assertIn("VerifyLinkedText", title)

    def test_title_mass_is_computed_in_metric_kg_for_active_configuration(self):
        title = self.read("Module7_TitleBlockEngine.bas")
        self.assertIn("ReadActiveConfigurationMassKg", title)
        self.assertIn("GetMassProperties2", title)
        self.assertIn("swMassPropertyAccuracyLevel_Higher", title)
        self.assertIn("swMassPropertiesStatus_OK", title)
        self.assertIn("massProperties(baseIndex + 5)", title)
        self.assertIn("|useSelected=False|", title)
        self.assertIn("|unit=kg|", title)
        self.assertIn('Format$(massKg, "0.00")', title)
        self.assertNotIn(
            'drawingManager, Array("Mass", "Weight"), "Mass", evidence',
            title,
        )


if __name__ == "__main__":
    unittest.main()
