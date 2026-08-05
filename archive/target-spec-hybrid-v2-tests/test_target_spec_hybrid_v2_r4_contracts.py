from pathlib import Path
import re
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"


class TargetSpecHybridV2R20ContractTests(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="utf-8-sig")

    @staticmethod
    def logical_lines(source: str):
        pending = ""
        for physical in source.splitlines():
            stripped = physical.rstrip()
            if stripped.endswith(" _"):
                pending += stripped[:-2] + " "
            else:
                yield pending + physical.lstrip() if pending else physical
                pending = ""
        if pending:
            yield pending

    def test_r23_source_identity_and_fixed_hybrid_are_explicit(self):
        main = self.read("Module1_Main.bas")
        pipeline = self.read("Module2_DrawingPipeline.bas")

        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)
        self.assertIn("DIAGNOSTIC_DRAWING_MODE As Boolean = True", main)
        self.assertIn("controlled-sheet preflight failed", pipeline)
        self.assertIn("sheet-scale transaction failed", pipeline)
        self.assertIn(
            r"V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT",
            main,
        )
        self.assertNotIn(
            r"V:\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT",
            main,
        )
        self.assertIn("ApplyFixtureAcceptanceProfile partPath", main)
        for setting in (
            "UseModelDimensions = True",
            "UseOrdinateDims = True",
            "RunHybridStrategy = True",
            "PopulateTitle = True",
            "InsertBarcode = True",
            "InsertNotes = True",
            "GenerateQAReport = True",
        ):
            with self.subTest(setting=setting):
                self.assertIn(setting, main)

        import_offset = pipeline.index("ImportAndReconcileR23Annotations")
        ordinate_offset = pipeline.index("CreateR23OrdinateGroups")
        iso_offset = pipeline.index("CreateIsometricView")
        self.assertLess(import_offset, ordinate_offset)
        self.assertLess(ordinate_offset, iso_offset)
        self.assertIn("BuildAllViewProjections", pipeline)

    def test_display_mode_false_result_requires_matching_readback(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")

        self.assertIn("ApplyDisplayModeWithReadback", pipeline)
        self.assertIn("actualMode = swView.GetDisplayMode2", pipeline)
        self.assertIn("If actualMode <> requestedMode Then", pipeline)
        self.assertIn("setterResult=False|readbackMatched=True", pipeline)

    def test_diagnostic_generation_is_separate_from_acceptance(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        runtime = self.read("Module8_RuntimeSupport.bas")
        layout = self.read("Module9_LayoutEngine.bas")
        qa = self.read("Module6_QAEngine.bas")

        self.assertIn("ContinueDiagnosticPipeline", pipeline)
        self.assertIn("RebuildDocumentVerified", runtime)
        self.assertIn("rebuildStatus = swModel.Extension.NeedsRebuild2", runtime)
        self.assertIn("readbackFullyRebuilt=True", runtime)
        self.assertIn("SHEET_SCALE|setterResult=False|", runtime)
        self.assertIn("FINAL_SHEET_CONTEXT|ActivateSheet=False|", runtime)
        self.assertIn("LAYOUT_BOUNDS|source=DiagnosticSheetReserve", layout)
        self.assertIn("diagnostic fallback bounds were used", layout)
        self.assertIn("Diagnostic Output Ready", qa)
        self.assertIn("evidence.ViewsCreated > 0", qa)
        self.assertIn("AnnotationSupportsPosition(annotation.GetType)", qa)

    def test_view_context_and_diagnostic_layout_are_state_verified(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        importer = self.read("Module4_ModelItemImporter.bas")
        runtime = self.read("Module8_RuntimeSupport.bas")
        layout = self.read("Module9_LayoutEngine.bas")

        self.assertNotIn("EnsureDrawingSheetEditMode", pipeline)
        self.assertNotIn("EnsureDrawingSheetEditMode", runtime)
        self.assertIn('"DRAWINGVIEW"', runtime)
        self.assertIn("SelectByID2", runtime)
        self.assertIn("ACTIVATE_VIEW_RETRY", runtime)
        self.assertIn("Set activeView = swDraw.ActiveDrawingView", runtime)
        self.assertIn("setterResult=False|readbackMatched=True", runtime)
        self.assertIn("Orthographic model import", importer)
        self.assertIn("IsOrdinateEligibleView", importer)
        self.assertIn("Set activeView = swDraw.ActiveDrawingView", importer)
        self.assertIn('"DRAWINGVIEW"', importer)
        self.assertIn("MODEL_IMPORT_VIEW_SELECTION", importer)
        self.assertIn("MODEL_IMPORT_EXECUTE", importer)
        self.assertIn("active-view readback matched, so continuing", importer)
        self.assertNotIn("failed to select drawing view", importer)
        self.assertNotIn("Whole-drawing model import", importer)
        self.assertNotIn('InsertModelAnnotations4( _\n                swImportModelItemsFromEntireModel, mask, True, True,', importer)
        self.assertIn("ArrangeDiagnosticRows", layout)
        self.assertIn("LAYOUT_RESULT|mode=DiagnosticZoneAware", layout)
        self.assertIn("If selectedView Then", importer)
        self.assertNotIn("If Not selectedView Then", importer)
        self.assertIn("isometricView.ScaleDecimal = 0.5", pipeline)

    def test_layout_reserves_title_as_rectangle_and_checks_notes(self):
        layout = self.read("Module9_LayoutEngine.bas")

        for contract in (
            "ArrangeZoneAwareViews",
            "titleBlockIsReservedRectangle=True",
            "TitleBlockBottom + _",
            "titleBlockReservedAsRectangle=True",
            "lowerLeftZoneUsed=True",
            "OutlineIntersectsTitleBlock",
            "ViewClearsAllNotes",
            "note.GetExtent",
            "Layout title-block intrusion",
            "Layout note intrusion",
            "without an unapproved scale change",
            "ArrangeP0251ReferenceZones",
            "profile=P0251ReferenceZones",
            "MoveViewOutlineCenter",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, layout)

        self.assertNotIn("evidence.UsableBottom = sheetHeight * 0.34", layout)

    def test_r20_repairs_runtime_proven_failure_contracts(self):
        audit = self.read("Module3_ModelAudit.bas")
        ordinate = self.read("Module5_FallbackDimensionEngine.bas")
        qa = self.read("Module6_QAEngine.bas")
        runtime = self.read("Module8_RuntimeSupport.bas")
        layout = self.read("Module9_LayoutEngine.bas")
        title = self.read("Module7_TitleBlockEngine.bas")

        candidate_start = ordinate.index("Private Function BuildOwnedCandidate")
        candidate_end = ordinate.index(
            "Private Sub CaptureFeatureSemantics", candidate_start
        )
        candidate_builder = ordinate[candidate_start:candidate_end]

        self.assertIn('"MIRRORPATTERN"', audit)
        self.assertIn("GetOwnedHoleSeedFeature", audit)
        self.assertIn("GetSeedFeature", audit)
        self.assertIn("GetPatternSeedFeature", audit)
        self.assertIn("internalCylinderProven", candidate_builder)
        self.assertNotIn(
            "IsInternalCylindricalFace(cylinderFace)", candidate_builder
        )
        self.assertIn("TryReadClosedCircularEdge", ordinate)
        self.assertNotIn(".IsCircle", candidate_builder)
        self.assertIn("provenCircleData", candidate_builder)
        self.assertIn("dimensionCount = 0", qa)
        self.assertIn("ordinateCount = 0", qa)
        self.assertIn("TryMeasureLegacyControlledTitleBlock", runtime)
        self.assertIn("GetTemplateSketch", runtime)
        self.assertIn("GetSketchSegments", runtime)
        self.assertIn("NON_RENDERED_NOTE_SKIPPED", qa)
        self.assertIn("NON_RENDERED_NOTE_SKIPPED", layout)
        self.assertIn("NoteReferencesSemanticField", title)
        self.assertIn('"SW-Sheet Scale(Sheet Scale)"', title)

    def test_r23_projection_contract_uses_proved_model_to_drawing_routes(self):
        ordinate = self.read("Module13_ProjectionResolution.bas")

        circle_start = ordinate.index("Private Function SelectModelEntityInView")
        circle_end = ordinate.index(
            "Private Function", circle_start + 1
        )
        circle_reader = ordinate[circle_start:circle_end]

        self.assertIn("swView.SelectEntity", circle_reader)
        self.assertIn("GetSelectedObjectsDrawingView2", circle_reader)
        self.assertIn("GetSelectedObject6", circle_reader)
        self.assertIn("NormalizeSwBoolean", ordinate)

    def test_r20_com_boolean_results_are_normalized_before_negation(self):
        managed = "\n".join(
            self.read(name)
            for name in (
                "Module2_DrawingPipeline.bas",
                "Module3_ModelAudit.bas",
                "Module4_ModelItemImporter.bas",
                "Module5_FallbackDimensionEngine.bas",
            )
        )

        forbidden = (
            r"If\s+Not\s+\w+\.IsCircle",
            r"If\s+Not\s+\w+\.IsCylinder",
            r"If\s+Not\s+\w+\.Select4",
            r"If\s+Not\s+\w+\.SetLabel",
            r"If\s+Not\s+\w+\.HasFullOutline",
            r"If\s+Not\s+.+AlignDimensions",
        )
        for pattern in forbidden:
            with self.subTest(pattern=pattern):
                self.assertIsNone(re.search(pattern, managed, re.IGNORECASE))

        for contract in (
            "profileSelected = CBool(",
            "labelSet = CBool(",
            "hasFullOutline = CBool(",
            "profileIsCircular = CBool(",
            "sectionSegmentSelected = CBool(",
            "annotationSelected = CBool(",
            "arrangeSucceeded = CBool(",
            "isCylinder = CBool(",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, managed)

    def test_r23_auto_arrange_excludes_semantic_section_lanes(self):
        importer = self.read("Module4_ModelItemImporter.bas")
        self.assertIn("IsExcludedFromGenericArrangement", importer)
        self.assertIn("SectionSemanticLanes", importer)
        self.assertIn("DIMENSION_ARRANGE_SKIPPED", importer)

    def test_r23_page_and_view_coordinates_have_an_explicit_inverse(self):
        section = self.read("Module17_SectionPath.bas")
        envelope = self.read("Module18_ContentEnvelope.bas")

        self.assertIn("PageToViewSketch", section)
        self.assertIn("ViewSketchToPage", envelope)
        self.assertIn("roundTripDeltaM", envelope)
        self.assertIn("GetXform", section + envelope)

    def test_r20_layout_gap_and_annotation_regions_are_structurally_safe(self):
        layout = self.read("Module9_LayoutEngine.bas")
        qa = self.read("Module6_QAEngine.bas")
        runtime = self.read("Module8_RuntimeSupport.bas")
        evidence = self.read("CRunEvidence.cls")

        self.assertIn("VIEW_GAP_M As Double = 0.012", layout)
        self.assertIn("LAYOUT_COMPARISON_TOLERANCE_M", layout)
        self.assertIn("LAYOUT_PAIR_CLEARANCE", layout)
        self.assertIn("topBoundary - rowHeight / 2# - VIEW_GAP_M / 2#", layout)
        self.assertIn("ContentBorderLeft", evidence + runtime + qa)
        self.assertIn("Content border bounds m:", evidence)
        self.assertNotIn("y < evidence.UsableBottom", qa)
        self.assertNotIn("extentBottom < evidence.UsableBottom", qa)
        self.assertIn('violationReason = "ZonedBorder"', qa)
        self.assertIn('violationReason = "TitleBlock"', qa)

    def test_r23_p0251_manufacturing_definitions_are_semantic_and_fail_closed(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        title = self.read("Module7_TitleBlockEngine.bas")
        qa = self.read("Module6_QAEngine.bas")

        self.assertIn("CreateMissingR23Callouts", pipeline)
        self.assertIn("Module16_CalloutDefinition", pipeline)
        self.assertNotIn("AddRequiredManufacturingDefinitions", pipeline)
        for contract in (
            "CollectRetainedDefinitions",
            "CreateNativeCalloutForFamily",
            "FirstAnchoredProjectionForFamily",
            "definition.IsComplete()",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, pipeline)

        self.assertNotIn("P0251ManufacturingDefinitionIsVisible", qa)
        self.assertNotIn("AssociatedCalloutContainsTokens", qa)
        self.assertIn("RequireSemanticStages", qa)
        self.assertNotIn("<MOD-DIAM>47 H7", title)

    def test_r20_section_geometry_is_checked_against_part_identification(self):
        qa = self.read("Module6_QAEngine.bas")
        evidence = self.read("CRunEvidence.cls")
        title = self.read("Module7_TitleBlockEngine.bas")

        for contract in (
            "PartIdentificationBoundsProven",
            "PartIdentificationLeft",
            "PartIdentificationBottom",
            "PartIdentificationRight",
            "PartIdentificationTop",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, evidence + title + qa)

        for contract in (
            "CheckSectionLineClearance",
            "GetSectionLineCount2",
            "GetSectionLineInfo2",
            "ValidateSectionLineInfo",
            "SectionSegmentTouchesPartIdentification",
            "PointTouchesPartIdentification",
            'RequireStage "SECTION_CLEARANCE"',
            'MarkStageProved "SECTION_CLEARANCE"',
            'MarkStageFailed "SECTION_CLEARANCE"',
            "SECTION_LINE_GEOMETRY",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, qa)

        self.assertIn(
            "swView.GetSectionLineCount2(sectionLineInfoSize)", qa
        )
        self.assertNotRegex(qa, r"GetSectionLineCount2\s*(?:\r?\n|$)")
        self.assertIn("sectionInfoItemCount <> sectionLineInfoSize", qa)

    def test_fixture_acceptance_profiles_match_reference_led_view_plans(self):
        main = self.read("Module1_Main.bas")
        form = self.read("UserForm1.frm")

        for fixture in (
            'Case "P-0251-14A-001"',
            'Case "P-0252-01-001"',
            'Case "P-0252-01-013"',
        ):
            with self.subTest(fixture=fixture):
                self.assertIn(fixture, main)
                self.assertIn(fixture, form)

        self.assertRegex(
            main,
            re.compile(
                r'Case "P-0251-14A-001".*?CreateLeft = True.*?'
                r'ConfigureRequiredSection "J", False',
                re.DOTALL,
            ),
        )
        self.assertRegex(
            main,
            re.compile(
                r'Case "P-0252-01-001".*?CreateRight = True.*?'
                r'SyncSectionCompatibilityFields',
                re.DOTALL,
            ),
        )
        self.assertRegex(
            main,
            re.compile(
                r'Case "P-0252-01-013".*?CreateBottom = True.*?'
                r'CreateLeft = True.*?CreateRight = True.*?'
                r'ConfigureRequiredSection "B", False',
                re.DOTALL,
            ),
        )

        for scale in ('"1:1"', '"1:6"', '"1.5:1"'):
            with self.subTest(scale=scale):
                self.assertIn(f"SetComboChoice cmbScale, {scale}", form)

    def test_views_have_only_the_fixture_approved_primary_rotation(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")

        self.assertIn("ApplyFixturePrimaryViewRotation", pipeline)
        self.assertIn('fixtureKey <> "P-0251-14A-001"', pipeline)
        self.assertIn("P0251_PRIMARY_CLOCKWISE_90_RAD", pipeline)
        self.assertIn("primaryView.Angle = requestedAngle", pipeline)
        self.assertIn("actualAngle = primaryView.Angle", pipeline)
        self.assertIn("VIEW_ROTATION|fixture=P-0251-14A-001", pipeline)
        self.assertNotIn("createdView.Angle =", pipeline)
        self.assertNotRegex(pipeline, r"(?im)^\s*\w+\.ScaleDecimal\s*=\s*0\.55")
        configure_ortho = pipeline[
            pipeline.index("Private Function ConfigureOrthographicView") :
            pipeline.index("Private Function ConfigureIsometricView")
        ]
        configure_iso = pipeline[
            pipeline.index("Private Function ConfigureIsometricView") :
            pipeline.index("Private Function CreatePrimarySection")
        ]
        self.assertIn("swView.UseSheetScale = 1", configure_ortho)
        self.assertIn("swView.UseSheetScale = 1", configure_iso)

    def test_pipeline_finalization_is_single_pass_and_non_recursive(self):
        main = self.read("Module1_Main.bas")
        pipeline = self.read("Module2_DrawingPipeline.bas")

        self.assertNotIn("Resume FinishRun", pipeline)
        self.assertIn("FinalizeRunOnce", pipeline)
        self.assertIn("If finalizationStarted Then Exit Sub", pipeline)
        self.assertIn("swDrawModel.GetType <> swDocDRAWING", pipeline)
        self.assertIn("TryFinalizeDrawingState", pipeline)
        self.assertIn("TryRunReadOnlyQa", pipeline)
        self.assertIn("TryEmitEvidence", pipeline)
        self.assertIn("Set GlobalEvidence = Nothing", main)
        self.assertIn("InitializeEvidenceIdentity", main)

    def test_view_and_section_failures_propagate_with_selection_context(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")

        self.assertIn(
            "Private Function ConfigureOrthographicView", pipeline
        )
        self.assertIn("Private Function ConfigureIsometricView", pipeline)
        self.assertIn("CreateSelectData", pipeline)
        self.assertNotIn("Set selectData.View = primaryView", pipeline)
        self.assertIn("ActiveSourceViewOwnsNewSketchSegments", pipeline)
        self.assertIn("VerifySectionSelection", pipeline)
        self.assertIn("GetSelectedObjectCount2(-1)", pipeline)
        self.assertIn("GetSelectedObjectsDrawingView2", pipeline)
        self.assertIn("GetSelectedObjectMark", pipeline)
        self.assertIn("SldWorks.DrSection", pipeline)
        self.assertIn("SetLabel2(sectionLabel)", pipeline)
        self.assertIn("GetSectionLineCount2", pipeline)
        self.assertIn("GetSectionLineInfo2", pipeline)

    def test_primary_section_records_step_level_selection_diagnostics(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")

        for step in (
            "SelectionManager.Acquire.Before",
            "SelectionManager.Acquire.After",
            "CreateSelectData.Before",
            "CreateSelectData.After",
            "SelectData.ViewAssignment.Before",
            "SelectData.ViewAssignment.Skipped",
            "Select4.Before.Index=",
            "Select4.After.Index=",
            "VerifySectionSelection.Before",
            "VerifySectionSelection.After.True",
            "CreateSectionViewAt5.Before",
            "CreateSectionViewAt5.After",
            "|SECTION_STEP=",
        ):
            with self.subTest(step=step):
                self.assertIn(step, pipeline)

        self.assertIn("If sectionSegment Is Nothing Then", pipeline)
        self.assertIn("RequiredDrawingContextIsNothing", pipeline)

    def test_fixture_locked_ui_and_scale_validation_are_truthful(self):
        form = self.read("UserForm1.frm")
        readme = self.read("README_IMPORT.md")

        self.assertIn("Fixture Primary Section (locked)", form)
        self.assertIn("Layout Preview (not available)", form)
        self.assertIn("chkPreview.Value = False", form)
        self.assertIn("TryResolveScaleValue", form)
        self.assertIn("TryParseCustomScale", form)
        self.assertIn("SetComboChoice", form)
        self.assertIn("does not yet implement a separate preview", readme)

    def test_primary_section_is_single_deterministic_and_coordinate_explicit(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")

        for contract in (
            "GlobalSectionCount <> 1",
            "CreateSemanticPrimarySection",
            "Module17_SectionPath.ResolveSectionPath",
            "Module17_SectionPath.CreateSectionFromPath",
            "placement=ProvisionalBeforeFinalEnvelopeLayout",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, pipeline)

        self.assertNotIn("topY + extension", pipeline)
        self.assertNotIn("bottomY - extension", pipeline)

    def test_pump_details_c_and_d_are_reference_led_and_mandatory(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        layout = self.read("Module9_LayoutEngine.bas")
        qa = self.read("Module6_QAEngine.bas")

        for contract in (
            "CreateRequiredDetails",
            "detailSourceView",
            "sourceView.ReferencedDocument",
            "sourceView.ReferencedConfiguration",
            "sourceView.Sheet",
            'Trim$(sourceView.GetOrientationName), "*Bottom"',
            "C_reference_mm=4.00,29.75",
            "D_reference_mm=35.30,8.50",
            "CreateCircleByRadius",
            "profileSegment.GetType <> swSketchARC",
            "profileArc.IsCircle <> 1",
            "CreateDetailViewAt4",
            "swDetViewSTANDARD",
            "swDetCircleCIRCLE",
            "SldWorks.DetailCircle",
            "ActiveDrawingView",
            "ObjectsAreSame",
            "IsCircle",
            "CircleParams",
            "UseParentScale",
            "ScaleRatio",
            "GetDetailView",
            "GetStyle",
            "GetDisplay",
            "GetProfileItemsCount",
            "GetProfileItems",
            "NoOutline",
            "HasFullOutline",
            "JaggedOutline",
            "Position readback",
            "ScaleDecimal - 3#",
            "REQUIRED_DETAILS_STRUCTURE",
            "REQUIRED_DETAILS_GEOMETRY|status=PENDING",
            "REQUIRED_DETAILS_LEGIBILITY|status=PENDING",
            "DETAIL_ORPHAN_PROFILE",
        ):
            with self.subTest(contract=contract):
                self.assertIn(contract, pipeline)

        self.assertNotIn("FindViewByOrientation", pipeline)
        self.assertNotIn("GlobalConfig.ConfigurationName", pipeline)

        detail_match = re.search(
            r"(?ims)^Private Function CreateOneDetail\b.*?^End Function\s*$",
            pipeline,
        )
        self.assertIsNotNone(detail_match)
        detail_block = detail_match.group(0)
        cleanup_gate = detail_block.index(
            "If structuralPostconditionsPassed And cleanupErrorNumber = 0 Then"
        )
        result_log = detail_block.index('evidence.AddInfo "DETAIL_RESULT|label="')
        accepted_count = detail_block.index(
            "evidence.ViewsCreated = evidence.ViewsCreated + 1", result_log
        )
        self.assertLess(cleanup_gate, result_log)
        self.assertLess(result_log, accepted_count)

        self.assertIn("firstView.Type = swDrawingDetailView", layout)
        self.assertIn("detailCount <> 2", qa)
        self.assertIn(
            "P-0252-01-013 acceptance requires Details C and D", qa
        )

    def test_section_dialog_returns_values_before_it_is_unloaded(self):
        main_form = self.read("UserForm1.frm")
        section_form = self.read("UserFormSection.frm")

        self.assertIn("Set sectionForm = New UserFormSection", main_form)
        self.assertIn("sectionForm.Show", main_form)
        self.assertIn("sectionForm.SectionLabel", main_form)
        self.assertIn("sectionForm.SectionVertical", main_form)
        self.assertIn("Unload sectionForm", main_form)
        self.assertIn("Me.Hide", section_form)
        self.assertNotRegex(
            section_form,
            r"(?is)Public Sub Do(?:Ok|Cancel)\(\).*?Unload Me.*?End Sub",
        )

    def test_import_handoff_preserves_host_and_form_designers(self):
        readme = self.read("README_IMPORT.md")

        self.assertIn("retain the host project's existing `ThisLibrary`", readme)
        self.assertIn(
            "do not attempt to import `ThisLibrary.cls` as an ordinary class",
            readme,
        )
        self.assertIn("not native VBE designer exports", readme)
        self.assertIn("create two blank UserForms", readme)

    def test_procedures_do_not_redeclare_parameters_or_local_variables(self):
        start = re.compile(
            r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
            r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+"
            r"([A-Za-z_]\w*)\s*(?:\((.*?)\))?",
            re.IGNORECASE,
        )
        end = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\s*$", re.I)
        declaration = re.compile(r"^\s*(?:Dim|Static)\s+(.+)$", re.I)
        parameter = re.compile(
            r"(?:^|,)\s*(?:(?:ByVal|ByRef|Optional|ParamArray)\s+)?"
            r"([A-Za-z_]\w*)",
            re.I,
        )

        for path in SOURCE.iterdir():
            if path.suffix.lower() not in {".bas", ".cls", ".frm"}:
                continue

            current = None
            names = set()
            for line_number, line in enumerate(
                self.logical_lines(path.read_text(encoding="utf-8-sig")), 1
            ):
                match = start.match(line)
                if match:
                    current = match.group(1)
                    names = {
                        item.group(1).lower()
                        for item in parameter.finditer(match.group(2) or "")
                    }
                    continue

                if current and end.match(line):
                    current = None
                    names = set()
                    continue

                if not current:
                    continue

                match = declaration.match(line)
                if not match:
                    continue

                for item in match.group(1).split(","):
                    name_match = re.match(r"\s*([A-Za-z_]\w*)", item)
                    if not name_match:
                        continue
                    name = name_match.group(1).lower()
                    self.assertNotIn(
                        name,
                        names,
                        f"{path.name}:{line_number} redeclares {name} "
                        f"inside {current}",
                    )
                    names.add(name)


if __name__ == "__main__":
    unittest.main()
