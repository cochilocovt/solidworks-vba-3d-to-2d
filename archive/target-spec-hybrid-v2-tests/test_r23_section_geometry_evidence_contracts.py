"""R23-812. The section-view geometry inventory is evidence, not a change.

Section View J-J first existed in the r49 run
(macro_qa/20260805_034637_P-0251-14A-001). SECTION_DIMENSIONS reported
`requirements:7/satisfied:0/missing:7` and CreateSectionDimension refuses
without a selection (`reason=NoEntitiesSelected`), so the unknown is not how
to create a dimension - it is which drawing curves that view exposes and
whether any pair of them spans a required nominal. Nothing had ever
measured it.

Two properties make the measurement worth trusting, and these tests hold
both:

1. It changes nothing. Module10's safety boundary names exactly two
   mutating procedures, CreateSectionDimension and ApplyReferenceFit, and
   the inventory is neither.
2. It is self-checking. The Help states the polyline-data array and the
   returned entity array are positionally paired, so the number of records
   decoded must equal the number of entities; any out-of-range field stops
   the walk and reports Desynchronized instead of emitting invented
   geometry. Without that control the coordinates would be unfalsifiable -
   the r40/r41 visibility classifiers failed exactly this way.

The record layout is the documented one, SOLIDWORKS 2025 Help,
IView::GetPolylines7:

    [Type, GeomDataSize, GeomData[], LineColor, LineStyle, LineFont,
     LineWeight, LayerID, LayerOverride, NumPolyPoints, [x,y,z]...]

with Type 1 carrying [cx,cy,cz, sx,sy,sz, ex,ey,ez, nx,ny,nz].
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module10_SectionDimensionEngine.bas"
PIPELINE = "Module2_DrawingPipeline.bas"
LAYOUT = "Module9_LayoutEngine.bas"


class R23SectionGeometryEvidenceContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def inventory(self) -> str:
        return self.code(MODULE).split(
            "Public Sub InventorySectionGeometry("
        )[1].split("\nEnd Sub")[0]

    # --- it is wired in ---------------------------------------------------

    def test_the_pipeline_calls_the_inventory_for_the_section_view(self):
        body = self.code(PIPELINE).split(
            "Private Function ReconcileR23SectionDimensions("
        )[1].split("\nEnd Function")[0]
        self.assertIn(
            "Module10_SectionDimensionEngine.InventorySectionGeometry",
            body,
        )
        self.assertLess(
            body.index("InventorySectionGeometry"),
            body.index("InventorySectionDimensions"),
        )

    # --- it changes nothing ----------------------------------------------

    def test_the_inventory_is_read_only(self):
        body = self.inventory()
        for forbidden in ("AddDimension2", "SelectByID2", "SelectEntity",
                          "ClearSelection2", "RecordSolidWorksMutation",
                          "SetFitValues", "SetValues2", "EditRebuild3",
                          "Position =", "UseSheetScale ="):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_the_safety_boundary_names_every_mutator(self):
        """R23-821 added CreateResolvedSectionDimensions, which creates
        nothing itself and calls CreateSectionDimension. R23-823 added
        ClampSectionAnnotationsIntoUsableArea, which creates nothing and
        only moves annotation text. The header must name all four, and each
        must refuse without allowMutation."""
        header = self.read(MODULE).split("Option Explicit")[1][:1600]
        self.assertIn("SAFETY BOUNDARY", header)
        self.assertIn("Exactly four procedures change a drawing", header)
        self.assertIn("CreateResolvedSectionDimensions", header)
        self.assertIn("ClampSectionAnnotationsIntoUsableArea", header)

        body = self.code(MODULE)
        for mutator in (
            "Public Function CreateSectionDimension(",
            "Public Function ApplyReferenceFit(",
            "Public Function CreateResolvedSectionDimensions(",
            "Public Function ClampSectionAnnotationsIntoUsableArea(",
        ):
            with self.subTest(mutator=mutator):
                proc = body.split(mutator)[1].split("\nEnd Function")[0]
                self.assertIn("ByVal allowMutation As Boolean", proc)
                self.assertIn("If Not allowMutation Then", proc)

    def test_the_read_only_probe_names_the_new_mutator_too(self):
        """The probe must stay free of every mutating call the module can
        make, including the one R23-823 introduced."""
        body = self.code(MODULE)
        probe = body.split("Public Sub R23_ProbeSectionDimensions(")[1]
        probe = probe.split("\nEnd Sub")[0]
        for forbidden in ("AddDimension2", "SetFitValues", "SetValues2",
                          "SetPosition2"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, probe)

    # --- it is self-checking ---------------------------------------------

    def test_the_walk_is_bounded_by_the_entity_count(self):
        """R23-818. The entity count IS the record count; the array is
        zero-padded beyond it. Measured in
        macro_qa/20260805_054951_P-0251-14A-001: every double from index
        2289 to the end reads 0.000000000, and an unbounded walk parsed 510
        of them as 56 phantom records of stride 9 (type 0, GeomDataSize 0,
        six zero style fields, NumPolyPoints 0), reaching 135 records with
        6 doubles left over. That is the whole of
        Desynchronized:StyleAt2801."""
        body = self.inventory()
        self.assertIn("SafeItemCount(entities)", body)
        self.assertIn(
            "If entityCount > 0 And decoded >= entityCount Then Exit Do",
            body,
        )
        self.assertIn("recordsMatchEntities=", body)

    def test_the_padding_is_verified_to_be_zeros(self):
        """A non-zero tail would mean the bound is discarding real records.
        That has to be visible, not silently trimmed - assuming the padding
        is what made the r56 reversal wrong in the first place."""
        body = self.inventory()
        self.assertIn("trailingAllZero = True", body)
        self.assertIn("If CDbl(polylineData(tailIndex)) <> 0# Then", body)
        self.assertIn('decodeStatus = "TrailingDataAfterEntities"', body)
        self.assertIn('"|trailingAllZero="', body)

    def test_the_record_count_must_match_the_entity_count(self):
        body = self.inventory()
        self.assertIn("trailingDoubles = upper - i + 1", body)
        self.assertIn('decodeStatus = "RecordCountMismatch"', body)
        self.assertIn('"|trailing="', body)
        self.assertIn('"|consumed="', body)

    def test_a_failed_decode_prints_the_raw_values_around_the_failure(self):
        """2799 doubles cannot be dumped whole, and without a window the
        next iteration would be guessing at the layout again."""
        body = self.inventory()
        self.assertIn("SECTION_GEOM_DESYNC", body)
        self.assertIn("|lastRecordStarts=", body)
        self.assertEqual(2, body.count("EmitRawWindow evidence"))

        window = self.code(MODULE).split(
            "Private Sub EmitRawWindow(")[1].split("\nEnd Sub")[0]
        self.assertIn("SECTION_GEOM_WINDOW", window)
        self.assertIn("GEOM_WINDOW_SPAN", window)

    def test_untrusted_coordinates_are_labelled_at_every_consumer(self):
        """The r53 report carried seven radii and six found=True verdicts
        from a walk that had already desynchronized, and only the summary
        line said so."""
        body = self.inventory()
        for line in ("SECTION_GEOM_Y", "SECTION_GEOM_X", "SECTION_GEOM_R"):
            with self.subTest(line=line):
                emitted = body.split(line)[1][:120]
                self.assertIn("|decodeStatus=", emitted)

        candidates = self.code(MODULE).split(
            "Private Function ReportRequirementCandidates(")[1].split(
            "\nEnd Sub")[0]
        self.assertIn("geometryTrusted = _", candidates)
        self.assertIn('StrComp(decodeStatus, "Complete", vbBinaryCompare)',
                      candidates)
        self.assertIn("|geometryTrusted=", candidates)

    def test_a_bad_field_stops_the_walk_instead_of_inventing_geometry(self):
        """Every field the walk advances over is range-checked before it is
        trusted: the header pair, the type and GeomData size, the style
        block, and the point count. One missing guard is enough to walk off
        the end of a record and emit coordinates that were never in the
        view, so all four are asserted individually."""
        body = self.inventory()
        for guard in ("Desynchronized:HeaderAt", "Desynchronized:TypeAt",
                      "Desynchronized:StyleAt", "Desynchronized:PointsAt"):
            with self.subTest(guard=guard):
                self.assertIn('decodeStatus = "' + guard, body)

        self.assertGreaterEqual(body.count("Exit Do"), 4)

    def test_record_layout_matches_the_documented_one(self):
        source = self.code(MODULE)
        self.assertIn("Private Const GEOM_ARC_DATA_SIZE As Long = 12",
                      source)
        self.assertIn("Private Const GEOM_STYLE_FIELD_COUNT As Long = 6",
                      source)
        self.assertIn("Private Const CROSSHATCH_EXCLUDE As Long = 1",
                      source)
        self.assertIn("Private Const GEOM_TYPE_ARC As Long = 1", source)

    def test_the_coordinate_frame_is_reported_not_assumed(self):
        """Which frame a section view's polylines live in has never been
        proved. The box is reported beside the sheet-space outline so it is
        decidable from the log."""
        body = self.inventory()
        self.assertIn("SECTION_GEOM_FRAME", body)
        self.assertIn("|polylineBox=", body)
        self.assertIn("|sheetOutline=", body)

    # --- what it answers --------------------------------------------------

    def test_every_requirement_gets_a_candidate_verdict(self):
        body = self.code(MODULE).split(
            "Private Function ReportRequirementCandidates("
        )[1].split("\nEnd Sub")[0]
        self.assertIn("BuildSectionRequirements()", body)
        self.assertIn("SECTION_REQ_CANDIDATE", body)
        self.assertIn("RequiresDiameterDisplay", body)
        self.assertIn("FindPairWithSpan(", body)
        self.assertIn("FindValueNear(", body)

    def test_nominal_tolerance_cannot_confuse_two_requirements(self):
        """The seven nominals are at least 5.5 mm apart; the match window
        is 0.01 mm."""
        source = self.code(MODULE)
        self.assertIn(
            "Private Const GEOM_NOMINAL_TOLERANCE_M As Double = 0.00001",
            source,
        )

    # --- the scale question is measured, not fixed ------------------------

    def test_per_view_scale_is_now_logged(self):
        body = self.code(LAYOUT)
        self.assertIn("VIEW_SCALE_READBACK", body)
        self.assertIn("|useSheetScale=", body)
        self.assertIn("|scaleDecimal=", body)

    def test_the_flag_is_no_longer_treated_as_the_ratio(self):
        """SOLIDWORKS 2025 Help, IView::UseSheetScale: "If the property is
        0, then it is possible that the view scale is the same as the sheet
        scale." A section view is tied to its parent through
        IView::UseParentScale and reads 0 here at the sheet ratio, which
        failed the whole LAYOUT stage in the r49 run."""
        body = self.code(LAYOUT)
        self.assertIn("ViewScaleMatchesSheet(firstView, evidence)", body)
        self.assertNotIn("is not using the proved sheet ", body)
        self.assertIn("is not drawn at the proved ", body)

    def test_the_flag_is_still_the_cheap_accept(self):
        """UseSheetScale = 1 is proof on its own; the ratio comparison is
        only reached when the flag does not already settle it."""
        body = self.code(LAYOUT)
        self.assertIn("If firstView.UseSheetScale <> 1 Then", body)
        self.assertLess(
            body.index("If firstView.UseSheetScale <> 1 Then"),
            body.index("ViewScaleMatchesSheet(firstView, evidence)"),
        )

    def test_the_ratio_comparison_fails_closed(self):
        """An unproved sheet scale or an unreadable view scale must NOT
        pass. Otherwise widening the check would accept a view whose scale
        nobody measured - the opposite of what the stage is for."""
        body = self.code(LAYOUT).split(
            "Private Function ViewScaleMatchesSheet("
        )[1].split("\nEnd Function")[0]
        self.assertIn(
            "If Not evidence.SheetScaleReadbackProven Then Exit Function",
            body,
        )
        self.assertIn(
            "If evidence.ActualScaleDenominator = 0# Then Exit Function",
            body,
        )
        self.assertIn("ScaleDecimal", body)
        self.assertIn("VIEW_SCALE_TOLERANCE", body)
        self.assertIn("ViewScaleMatchesSheet = False", body)

    def test_a_genuine_scale_mismatch_still_fails_the_stage(self):
        body = self.code(LAYOUT)
        self.assertIn("valid = False", body)
        self.assertIn("DescribeViewScale(firstView, evidence)", body)

    def test_the_detail_view_rule_is_untouched(self):
        """Details keep their approved independent 3:1 scale; only the
        non-isometric branch changed."""
        body = self.code(LAYOUT)
        self.assertIn("Abs(firstView.ScaleDecimal - 3#) > 0.000001", body)
        self.assertIn("Detail view is not using its approved 3:1 scale",
                      body)

    def test_the_isometric_exemption_is_untouched(self):
        body = self.code(LAYOUT)
        self.assertIn("ElseIf Not IsIsometricView(firstView) Then", body)


if __name__ == "__main__":
    unittest.main()
