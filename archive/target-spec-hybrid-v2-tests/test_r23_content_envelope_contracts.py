"""R23 Phase 9 source contracts for content-envelope-aware final layout.

Static source-contract tests for tasks R23-900 through R23-909. The
load-bearing ones are R23-900 (the envelope grows from annotations, not just
the model outline), the frame discipline (section geometry is view-sketch
and must be converted; display-data frame is unstated and is tested rather
than asserted), R23-902 (no fixed upward bias) and R23-907 (no view is
shrunk to force a fit).
"""

from pathlib import Path
import re
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module18_ContentEnvelope.bas"
ENVELOPE_CLS = "CContentEnvelope.cls"


class R23ContentEnvelopeContracts(unittest.TestCase):
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

    def test_components_exist_and_are_managed(self):
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")
        for name in (MODULE, ENVELOPE_CLS):
            with self.subTest(name=name):
                self.assertTrue((SOURCE / name).exists())
                self.assertIn(name.split(".")[0], manifest)

    # --- mutation safety -------------------------------------------------

    def test_evidence_entry_point_creates_and_moves_nothing(self):
        body = self.source.split("Public Sub R23_ProbeContentEnvelope()")[1]
        for forbidden in (
            "ApplyPlacementPlan",
            "EditRebuild3",
            ".Position =",
            "SealLayout",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        self.assertIn("mode=ReadOnly", body)
        self.assertIn("creations=0|mutations=0", body)
        self.assertIn("drawingUnchanged=", body)

    def test_probe_does_not_invoke_the_production_qa_gate(self):
        self.assertNotIn("EmitRunEvidence", self.executable)

    def test_only_one_procedure_moves_a_view_and_it_is_authorized(self):
        self.assertEqual(1, self.executable.count("swView.Position = target"))
        self.assertEqual(1, self.executable.count("EditRebuild3"))

        body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("swView.Position = target"),
        )
        self.assertIn("evidence.RecordSolidWorksMutation", body)

    # --- R23-900 ---------------------------------------------------------

    def test_r23_900_every_named_source_contributes(self):
        body = self.body("Public Function BuildViewEnvelope(")
        for call in (
            "AddOutline(",
            "AddAnnotationGeometry ",
            "AddDisplayDimensionGeometry ",
            "AddNoteExtents ",
            "AddSectionGeometry ",
        ):
            with self.subTest(call=call):
                self.assertIn(call, body)

    def test_r23_900_source_counts_are_kept_separately(self):
        cls = self.read(ENVELOPE_CLS)
        for counter in (
            "OutlinePoints",
            "DimensionPoints",
            "TextPoints",
            "NotePoints",
            "LeaderPoints",
            "ArrowPoints",
            "SectionPoints",
        ):
            with self.subTest(counter=counter):
                self.assertIn(counter, cls)

    def test_an_outline_only_rectangle_is_not_a_content_envelope(self):
        """R23-900 asks for a CONTENT envelope. An envelope that grew only
        from the model outline is the old behaviour wearing a new name."""
        cls = self.read(ENVELOPE_CLS)
        self.assertIn("Public Function HasAnnotationContent(", cls)
        self.assertIn("- OutlinePoints) > 0", cls)
        self.assertIn("annotationEnvelopes=", self.source)

    def test_arrow_geometry_and_label_text_height_are_included(self):
        body = self.body("Private Function ConsumeSectionInfo(")
        self.assertIn("SOURCE_ARROW", body)
        self.assertIn("textHeight", body)

        label = self.body("Private Sub ContributeSectionLabel(",
                          "\nEnd Sub")
        self.assertIn("pageY + textHeight", label)

    # --- frame discipline ------------------------------------------------

    def test_section_geometry_is_converted_from_the_view_sketch_frame(self):
        """Phase 0 proved GetSectionLineInfo2's payload is in the source
        view's sketch frame. Contributing it raw would place a page-frame
        envelope around view-frame numbers."""
        self.assertIn("Public Function ViewSketchToPage(", self.source)
        point = self.body("Private Sub ContributeSectionPoint(",
                          "\nEnd Sub")
        self.assertIn("ViewSketchToPage(swView", point)

    def test_the_inverse_transform_is_round_trip_checked(self):
        body = self.body("Public Function ProveInverseTransform(")
        self.assertIn(
            "Module17_SectionPath.PageToViewSketch(", body
        )
        self.assertIn("ViewSketchToPage(swView", body)
        self.assertIn("TRANSFORM_ROUNDTRIP_TOLERANCE_M", body)
        self.assertIn("roundTripDeltaM=", body)

    def test_the_inverse_is_the_actual_inverse_of_module17(self):
        """Module17 forward: view = R(-angle) * (page - origin) / scale.
        The inverse must be page = origin + scale * R(angle) * view."""
        body = self.body("Public Function ViewSketchToPage(")
        self.assertIn(
            "deltaX = viewX * Cos(viewAngle) - viewY * Sin(viewAngle)", body
        )
        self.assertIn(
            "deltaY = viewX * Sin(viewAngle) + viewY * Cos(viewAngle)", body
        )
        self.assertIn("pageX = originX + deltaX * viewScale", body)
        self.assertIn("pageY = originY + deltaY * viewScale", body)

    def test_display_data_frame_is_tested_not_asserted(self):
        """The IDisplayData Remarks do not state a frame. Claiming one would
        be asserting a contract the Help does not make."""
        self.assertIn("Private Sub RecordFrameAgreement(", self.source)
        self.assertIn("displayDataFramePageConsistent=", self.source)
        self.assertIn("displayDataFrameInconsistent=", self.source)

    def test_text_position_is_treated_as_an_offset_not_a_coordinate(self):
        """GetTextPositionAtIndex is explicitly an offset from the display
        data origin. Used as an absolute point it drags every envelope
        towards the sheet origin."""
        body = self.body("Private Sub AddDisplayDimensionGeometry(",
                         "\nEnd Sub")
        self.assertIn("textX = originX + CDbl(offset(", body)
        self.assertIn("textY = originY + CDbl(offset(", body)
        self.assertIn("If Not originKnown Then Exit Sub", body)

    def test_off_sheet_points_are_rejected_and_counted(self):
        body = self.body("Private Sub ContributePoint(", "\nEnd Sub")
        self.assertIn("x > sheetWidth Or y > sheetHeight", body)
        self.assertIn("envelope.RecordRejection x, y, source", body)
        cls = self.read(ENVELOPE_CLS)
        self.assertIn("RejectedOffSheet = RejectedOffSheet + 1", cls)
        self.assertIn("rejectedOffSheet=", cls)

    def test_leader_points_are_consumed_as_triples_from_the_array(self):
        """GetLeaderStyle is OR-ed with attachment bitmask flags, and the
        corpus returns those flag members with mangled values. The returned
        array length needs no such reconstruction."""
        body = self.body("Private Sub AddAnnotationGeometry(", "\nEnd Sub")
        self.assertIn("total \\ 3", body)
        self.assertNotIn("GetLeaderStyle", self.executable)

    # --- R23-901 ---------------------------------------------------------

    def test_r23_901_protected_regions_use_the_measured_bounds(self):
        body = self.body("Public Function BuildProtectedRegions(")
        for field in (
            "evidence.ContentBorderLeft",
            "evidence.ContentBorderRight",
            "evidence.ContentBorderBottom",
            "evidence.ContentBorderTop",
            "evidence.TitleBlockLeft",
            "evidence.PartIdentificationLeft",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

        self.assertIn("PartIdentificationBoundsProven", body)

    def test_no_rectangle_is_emitted_from_unmeasured_bounds(self):
        """A boundary that does not exist is not a boundary at the origin.
        Unset evidence fields would produce a degenerate rectangle at (0,0)
        and report false ProtectedIntrusion violations against it."""
        body = self.body("Public Function BuildProtectedRegions(")
        self.assertIn("If evidence.LayoutBoundariesProven Then", body)
        self.assertIn("IsRealRectangle(", body)
        self.assertLess(
            body.index("If evidence.LayoutBoundariesProven Then"),
            body.index("ContentBorderLeftStrip"),
        )

        gate = self.body("Private Function IsRealRectangle(")
        self.assertIn("(rightX > leftX) And (topY > bottomY)", gate)

    def test_the_probe_measures_the_sheet_itself_read_only(self):
        """MeasureControlledSheetRegions is a production fail-closed gate:
        it demands a title block the reference drawing does not have, and it
        SETS SheetFormatVisible. Calling it from a probe aborted the
        2026-08-01 run and attempted a mutation the run had promised not to
        make."""
        self.assertNotIn(
            "MeasureControlledSheetRegions", self.executable
        )
        self.assertNotIn("SheetFormatVisible", self.executable)

        self.assertIn("Public Function MeasureSheetRegions(", self.source)
        body = self.body("Public Function MeasureSheetRegions(")
        self.assertIn("swSheet.GetSize", body)
        self.assertIn("swSheet.GetZoneMargin(ZONE_TOP_MARGIN)", body)
        self.assertIn("evidence.LayoutBoundariesProven = False", body)

    def test_unmeasurable_regions_are_reported_not_fatal(self):
        body = self.body("Public Function MeasureSheetRegions(")
        for token in (
            "contentBorder=Unmeasured",
            "titleBlock=Absent",
            "titleBlock=PresentBoundsUnread",
            "usableSource=SheetExtentNoBorder",
            "usableSource=ContentBorder",
        ):
            with self.subTest(token=token):
                self.assertIn(token, body)

        # Only an unusable sheet SIZE stops the probe.
        probe = self.source.split(
            "Public Sub R23_ProbeContentEnvelope()"
        )[1]
        self.assertIn("reason=SheetSizeUnmeasured", probe)
        self.assertNotIn("reason=SheetRegionsUnmeasured", probe)

    def test_title_block_extents_are_read_when_they_exist(self):
        """A formal ITitleBlock must become a protected rectangle when its
        bounds are real and on-sheet; reporting it as unread leaves a live
        collision class unchecked."""
        body = self.body("Public Function MeasureSheetRegions(")
        self.assertIn("titleBlock.GetExtents", body)
        self.assertIn("ITitleBlock.GetExtents", body)
        self.assertIn("titleBlock=Measured", body)
        self.assertIn("titleBlock=PresentBoundsRejected", body)
        self.assertIn("evidence.TitleBlockLeft", body)

    def test_the_content_border_is_protected_as_strips_not_a_rectangle(
        self,
    ):
        """The drawable area is INSIDE the border. One rectangle would
        declare every view a violation."""
        body = self.body("Public Function BuildProtectedRegions(")
        for strip in (
            "ContentBorderLeftStrip",
            "ContentBorderRightStrip",
            "ContentBorderBottomStrip",
            "ContentBorderTopStrip",
        ):
            with self.subTest(strip=strip):
                self.assertIn(strip, body)

    # --- R23-902 ---------------------------------------------------------

    def test_r23_902_no_fixed_upward_bias_survives(self):
        """Module9_LayoutEngine pins the source row to the top boundary. A
        row pinned to a boundary has nowhere to put the annotations that
        hang above it."""
        for forbidden in (
            "topBoundary -",
            "Bias",
            "bias =",
            "rowCenterY",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

        body = self.body("Public Function PlanPlacement(")
        self.assertIn("bias=None", body)

    def test_placement_is_derived_from_envelope_sizes(self):
        body = self.body("Public Function PlanPlacement(")
        self.assertIn("envelope.Width()", body)
        self.assertIn("envelope.Height()", body)
        self.assertIn("evidence.UsableLeft", body)
        self.assertIn("evidence.UsableBottom", body)
        self.assertIn("VIEW_CLEARANCE_M", body)

    def test_the_plan_moves_nothing(self):
        body = self.body("Public Function PlanPlacement(")
        self.assertNotIn(".Position", body)
        self.assertNotIn("EditRebuild3", body)

    # --- R23-903 ---------------------------------------------------------

    def test_r23_903_views_move_by_envelope_centre_not_outline_centre(self):
        """Module9_LayoutEngine moves by the OUTLINE centre, which is
        exactly why annotations end up outside the region the layout
        believed it was filling."""
        body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("envelope.CentreX()", body)
        self.assertIn("envelope.CentreY()", body)
        self.assertNotIn("GetOutline", body)

    # --- R23-904 ---------------------------------------------------------

    def test_r23_904_rebuild_reacquire_and_one_correction_pass(self):
        body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("For pass = 0 To MAX_CORRECTION_PASSES", body)
        self.assertIn("RebuildVerified swDraw, evidence", body)
        self.assertIn("BuildViewEnvelope(views(i)", body)
        self.assertIn("Set viewEnvelopes = rebuiltEnvelopes", body)
        self.assertLess(
            body.index('RebuildVerified swDraw, evidence, "Pass"'),
            body.index("BuildViewEnvelope(views(i)"),
        )
        self.assertEqual(
            "Public Const MAX_CORRECTION_PASSES As Long = 1",
            [
                line
                for line in self.executable.split("\n")
                if "MAX_CORRECTION_PASSES As Long" in line
            ][0],
        )

    def test_a_failed_rebuild_is_reported_not_swallowed(self):
        body = self.body("Private Sub RebuildVerified(", "\nEnd Sub")
        self.assertIn("If Not rebuilt Then", body)
        self.assertIn("editRebuild3=False", body)
        self.assertEqual(1, self.executable.count("EditRebuild3"))

    # --- R23-905 and R23-906 ---------------------------------------------

    def test_r23_905_every_pair_is_checked_explicitly(self):
        body = self.body("Public Function VerifyClearances(")
        self.assertIn("For j = i + 1 To viewEnvelopes.Count", body)
        self.assertIn("For p = 1 To protectedRegions.Count", body)
        self.assertIn("ViewOverlap:", body)
        self.assertIn("ProtectedIntrusion:", body)
        self.assertIn("clearanceChecks=", body)

    def test_clearance_uses_a_separating_axis_measure(self):
        body = self.body("Public Function ClearanceBetween(")
        self.assertIn("first.MinX - second.MaxX", body)
        self.assertIn("second.MinX - first.MaxX", body)
        self.assertIn("first.MinY - second.MaxY", body)
        self.assertIn("second.MinY - first.MaxY", body)

    def test_r23_906_section_views_need_two_millimetres(self):
        self.assertIn(
            "SECTION_CLEARANCE_M As Double = 0.002", self.executable
        )
        body = self.body("Private Function RequiredClearance(")
        self.assertIn("VIEW_TYPE_SECTION", body)
        self.assertIn("ENVELOPE_KIND_PROTECTED", body)
        self.assertIn("RequiredClearance = SECTION_CLEARANCE_M", body)

    # --- R23-907 ---------------------------------------------------------

    def test_r23_907_rescaling_is_authorized_gated_and_reported(self):
        """R23-907 originally forbade reducing an approved view scale. The
        user reversed it on 2026-08-01: "The views are allowed to rescaled
        as per need". The prohibition is replaced by a gate and a record,
        not by silence."""
        self.assertEqual(1, self.executable.count("ScaleDecimal ="))

        body = self.body("Public Function ApplyScaleToFit(")
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("swView.ScaleDecimal ="),
        )
        self.assertIn("evidence.RecordSolidWorksMutation", body)
        # The new scale is read back, not assumed to have taken.
        self.assertIn("after = CDbl(swView.ScaleDecimal)", body)
        self.assertIn("readbackFailures", body)

    def test_rescale_preserves_view_scale_relationships(self):
        """A parent-derived view must not be directly scaled again, and a
        sheet-scaled root must be explicitly detached before ScaleDecimal is
        assigned."""
        body = self.body("Public Function ApplyScaleToFit(")
        self.assertIn("UsesParentScale", body)
        self.assertIn("GetBaseView", body)
        self.assertIn("ParentNotInLayout", body)
        self.assertIn("swView.UseSheetScale = 0", body)
        self.assertIn("detachedFromSheet=", body)
        self.assertLess(
            body.index("If inherited Then GoTo ContinueScale"),
            body.index("swView.ScaleDecimal ="),
        )

    def test_rescale_refuses_to_override_approved_presentation_scales(self):
        lock = self.body("Private Function IsScaleLockedView(")
        self.assertIn("VIEW_TYPE_DETAIL", lock)
        self.assertIn("ApprovedDetailScale", lock)
        self.assertIn("ApprovedIsometricScale", lock)

        body = self.body("Public Function ApplyScaleToFit(")
        self.assertLess(
            body.index("IsScaleLockedView(swView, lockReason)"),
            body.index("swView.ScaleDecimal ="),
        )

    def test_the_scale_factor_is_an_estimate_and_says_so(self):
        """Annotation text height does not scale with the view, so a view at
        half scale does not have half the envelope. The factor is applied
        and the envelopes are RE-MEASURED."""
        body = self.body("Public Function PlanPlacement(")
        self.assertIn("factorIs=GeometricEstimateTextDoesNotScale", body)

        apply_body = self.body("Public Function ApplyPlacementPlan(")
        self.assertLess(
            apply_body.index("ApplyScaleToFit("),
            apply_body.index("Set viewEnvelopes = rescaledEnvelopes"),
        )
        self.assertLess(
            apply_body.index("Set viewEnvelopes = rescaledEnvelopes"),
            apply_body.rindex("PlanPlacement("),
        )

    def test_every_scale_change_is_reported_by_name(self):
        body = self.body("Public Function ReportScaleChanges(")
        self.assertIn("scalesChanged=", body)
        self.assertIn("scaleChanges=", body)

        apply_body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("CaptureViewScales(views)", apply_body)
        self.assertIn("ReportScaleChanges(views, capturedScales)",
                      apply_body)

    # --- R23-908 ---------------------------------------------------------

    def test_r23_908_uses_a_bounded_measured_rescale_loop(self):
        """Rescaling is permitted, but never an unlimited shrink loop.
        Text does not scale, so each of the two allowed passes must rebuild
        and measure before the final LargerSheetRequired decision."""
        body = self.body("Public Function PlanPlacement(")
        self.assertIn("plan=RescaleRequired", body)
        self.assertIn("requiredWidthM=", body)
        self.assertIn("requiredHeightM=", body)
        self.assertIn("usableWidthM=", body)
        self.assertIn("usableHeightM=", body)

        apply_body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("MAX_RESCALE_PASSES As Long = 2", self.source)
        self.assertIn("Do While InStr(1, planProof", apply_body)
        self.assertIn("rescalePass < MAX_RESCALE_PASSES", apply_body)
        self.assertIn("rescalePass = rescalePass + 1", apply_body)
        self.assertIn("LAYOUT_RESCALE|pass=", apply_body)
        self.assertIn("rescalePasses=", apply_body)
        self.assertIn("reason=LargerSheetRequired", apply_body)
        self.assertLess(
            apply_body.index("ApplyScaleToFit("),
            apply_body.index("reason=LargerSheetRequired"),
        )

    def test_placement_refuses_an_unusable_or_stale_plan_before_move(self):
        plan = self.body("Public Function PlanPlacement(")
        self.assertIn("NoViewEnvelopes", plan)
        self.assertIn("InvalidContentExtent", plan)

        apply_body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("reason=PlanRejected", apply_body)
        self.assertIn("Set targetCentres = replan", apply_body)
        self.assertLess(
            apply_body.index("reason=PlanRejected"),
            apply_body.index("swView.Position = target"),
        )
        self.assertLess(
            apply_body.index("reason=PlanViewCountMismatch"),
            apply_body.index("ApplyScaleToFit("),
        )

    # --- 2026-08-01 second live run -------------------------------------

    def test_the_section_arrow_block_is_nine_doubles(self):
        """The live run returned 49 items for Drawing View4:
        2 header + 1 numSegments + 7x3 segments + 9 + 9 arrows + 7 text.
        Counting 11 per arrow matched nothing at all."""
        self.assertIn(
            "Private Const ARROW_BLOCK_DOUBLES As Long = 9",
            self.executable,
        )
        body = self.body("Private Function ConsumeSectionInfo(")
        self.assertIn("cursor = cursor + ARROW_BLOCK_DOUBLES", body)
        self.assertNotIn("cursor = cursor + 11", body)

    def test_a_view_without_a_section_line_is_not_a_failed_parse(self):
        """Reporting absence as SectionGrammarUnmatched put a source
        failure on every envelope in the second run."""
        body = "\n".join(
            line
            for line in self.body(
                "Private Function AddSectionGeometry("
            ).split("\n")
            if not line.lstrip().startswith("'")
        )
        self.assertIn("sectionGrammar=NoSectionLine", body)
        self.assertIn("If total < 1 Then", body)
        self.assertLess(
            body.index("If total < 1 Then"),
            body.index("SectionGrammarUnmatched"),
        )

    def test_rejected_points_are_sampled_not_only_counted(self):
        """34 rejections on the section view with no coordinate makes a
        frame error and genuinely off-sheet geometry look identical."""
        cls = self.read(ENVELOPE_CLS)
        self.assertIn("Public Sub RecordRejection(", cls)
        self.assertIn("FirstRejectedSource", cls)
        self.assertIn("firstRejected=", cls)

        body = self.body("Private Sub ContributePoint(", "\nEnd Sub")
        self.assertIn("envelope.RecordRejection x, y, source", body)

    def test_the_frame_agreement_slack_is_not_wide_enough_to_pass_anything(
        self,
    ):
        """The first version allowed ten view clearances - 120 mm - which
        would accept a view-frame point by accident."""
        self.assertIn(
            "FRAME_AGREEMENT_SLACK_M As Double = 0.024", self.executable
        )
        body = self.body("Private Sub RecordFrameAgreement(", "\nEnd Sub")
        self.assertIn("slack = FRAME_AGREEMENT_SLACK_M", body)
        self.assertNotIn("VIEW_CLEARANCE_M * 10#", self.executable)

    def test_both_line_endpoints_are_frame_checked(self):
        body = self.body("Private Sub AddDisplayDimensionGeometry(",
                         "\nEnd Sub")
        self.assertEqual(2, body.count("RecordFrameAgreement "))

    def test_template_orientation_entries_are_skipped_by_name(self):
        """GetViews returns *Front, *Top, *Isometric and the rest beside
        the real views. They seeded nothing and only added noise."""
        self.assertIn(
            "Public Function IsTemplateOrientationView(", self.source
        )
        body = self.body("Public Function IsTemplateOrientationView(")
        self.assertIn('Left$(SafeViewName(swView), 1) = "*"', body)

        probe = self.source.split(
            "Public Sub R23_ProbeContentEnvelope()"
        )[1]
        self.assertIn("reason=TemplateOrientationEntry", probe)

    def test_the_probe_does_not_print_what_the_ledger_already_prints(self):
        """CRunEvidence.AddInfo prints what it records. The second run's log
        carried every ENVELOPE line twice."""
        probe = self.source.split(
            "Public Sub R23_ProbeContentEnvelope()"
        )[1]
        self.assertNotIn('Debug.Print "QA INFO: ENVELOPE|"', probe)

    # --- R23-909 ---------------------------------------------------------

    def test_r23_909_the_layout_is_sealed_and_the_seal_is_checked(self):
        self.assertIn("Public Sub SealLayout(", self.source)
        self.assertIn(
            "Public Function VerifyNothingCreatedAfterLayout(", self.source
        )

        seal = self.body("Public Sub SealLayout(", "\nEnd Sub")
        self.assertIn("evidence.SolidWorksMutationSequence", seal)

        verify = self.body(
            "Public Function VerifyNothingCreatedAfterLayout("
        )
        self.assertIn("added = 0", verify)
        self.assertIn("sealedSequence=", verify)
        self.assertIn("currentSequence=", verify)

        apply_body = self.body("Public Function ApplyPlacementPlan(")
        self.assertIn("SealLayout evidence", apply_body)

    # --- hygiene ---------------------------------------------------------

    def test_source_hygiene(self):
        for name in (MODULE, ENVELOPE_CLS):
            with self.subTest(name=name):
                raw = (SOURCE / name).read_bytes()
                self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
                self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))
                self.assertTrue(all(b < 0x80 for b in raw))
                text = raw.decode("cp1252")
                self.assertTrue(text.startswith("Option Explicit"))
                for line in text.split("\r\n"):
                    self.assertLessEqual(len(line), 79)

    def test_no_vba_statement_keyword_is_used_as_a_variable(self):
        """`Line` is a VBA statement keyword; declaring it as a variable
        compiles differently than it reads."""
        for keyword in ("line", "name", "print"):
            with self.subTest(keyword=keyword):
                self.assertIsNone(
                    re.search(
                        r"\bDim %s\b" % keyword,
                        self.executable,
                        re.IGNORECASE,
                    )
                )

    def test_source_revision_identifies_wired_r23_pipeline(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
