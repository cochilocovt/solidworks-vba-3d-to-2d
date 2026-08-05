"""R23 Phase 7 source contracts for the model-intent section path.

Static source-contract tests for tasks R23-700 through R23-708. The
load-bearing ones are R23-704 (none of the disproved outline-percentage
strategy in the new path), R23-705 (the frame conversion happens exactly
once), and the mutation boundary.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module17_SectionPath.bas"
PATH_CLS = "CSectionPath.cls"


class R23SectionPathContracts(unittest.TestCase):
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
        for name in (MODULE, PATH_CLS):
            with self.subTest(name=name):
                self.assertTrue((SOURCE / name).exists())
                self.assertIn(name.split(".")[0], manifest)

    # --- mutation safety -------------------------------------------------

    def test_evidence_entry_point_creates_nothing(self):
        body = self.source.split("Public Sub R23_ProbeSectionPath()")[1]
        for forbidden in (
            "CreateSectionFromPath",
            "CreateLine",
            "CreateSectionViewAt5",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        self.assertIn("mode=ReadOnly", body)
        self.assertIn("creations=0", body)
        self.assertIn("drawingUnchanged=", body)

    def test_probe_does_not_invoke_the_production_qa_gate(self):
        self.assertNotIn("EmitRunEvidence", self.executable)

    def test_only_one_procedure_can_mutate_and_it_must_be_authorized(self):
        self.assertEqual(1, self.executable.count("CreateSectionViewAt5("))
        self.assertEqual(1, self.executable.count("CreateLine("))
        body = self.body("Public Function CreateSectionFromPath(")
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("CreateLine("),
        )

    def test_r23_708_an_unproved_path_is_never_approximated(self):
        """The whole point of Phase 7 is that a cut is built from proved
        geometry or not at all."""
        body = self.body("Public Function CreateSectionFromPath(")
        self.assertIn("If Not path.Resolved Then", body)
        self.assertIn("PathUnresolved", body)
        self.assertLess(
            body.index("If Not path.Resolved Then"),
            body.index("CreateLine("),
        )

    def test_view_placement_is_a_caller_argument_not_a_default(self):
        """CreateSectionViewAt5's X/Y are the CENTRE of the new view.
        Defaulting them to a point on the source view - the obvious
        shortcut - stacks the section on top of the view it was cut from.
        Choosing where a view sits is layout, a later phase."""
        body = self.body("Public Function CreateSectionFromPath(")
        self.assertIn("ByVal placeX As Double", body)
        self.assertIn("ByVal placeY As Double", body)
        self.assertIn(
            "CreateSectionViewAt5( _\n        placeX, placeY, 0#,", body
        )
        # The label must be the resolved one, not an empty string. The
        # options argument became swCreateSectionView_OffsetSection at r53;
        # see test_r23_offset_section_option_contracts.
        self.assertIn(
            "path.SectionLabel, sectionOptions, Nothing, 0#", body)

    # --- R23-700 ---------------------------------------------------------

    def test_r23_700_section_work_is_independent_of_ordinate_creation(self):
        """Section construction consumes proved locations; it must not wait
        on or interact with the ordinate engine.

        Checked against the comment-stripped source: the header comment
        names the module precisely to record that it is deliberately absent
        from the code, and that documentation is not a reference."""
        self.assertNotIn("Module15_OrdinateScheme", self.executable)

    # --- R23-701 ---------------------------------------------------------

    def test_r23_701_bore_is_a_singleton_family_not_a_radius_threshold(self):
        body = self.body("Private Function ResolveBoreProjection(")
        self.assertIn("graph.LocationsForFamily(", body)
        self.assertIn("family.Count <> 1", body)
        self.assertIn("MaximumRadiusM", body)
        # No magic millimetre value decides what a bore is.
        for magic in ("0.015", "0.020", "0.0033"):
            with self.subTest(magic=magic):
                self.assertNotIn(magic, body)

    def test_r23_701_face_holes_share_the_bore_machining_face(self):
        body = self.body("Private Function CollectFaceHoleProjections(")
        self.assertIn("MachiningFaceToken(", body)
        self.assertIn("family.Count < 2", body)
        self.assertIn("candidate.Accepted", body)

    def test_grid_shape_is_proved_not_assumed(self):
        body = self.body("Public Function ResolveSectionPath(")
        self.assertIn("FewerThanTwoFaceHoleColumns", body)
        self.assertIn("FewerThanTwoFaceHoleRows", body)

    # --- R23-702 ---------------------------------------------------------

    def test_r23_702_waypoints_follow_the_approved_order(self):
        """1 bore centre; 2 same X at highest row; 3 minimum-X column at
        that row; 4 same column at lowest row."""
        body = self.body("Public Function ResolveSectionPath(")
        self.assertIn("path.W1X = path.BoreProjection.PageX", body)
        self.assertIn(
            "path.W1Y = path.BoreProjection.PageY + path.BoreOvershootM",
            body,
        )
        self.assertIn("path.W2X = path.BoreProjection.PageX", body)
        self.assertIn("path.W2Y = highestRowY", body)
        self.assertIn("path.W3X = minimumColumnX", body)
        self.assertIn("path.W3Y = highestRowY", body)
        self.assertIn("path.W4X = minimumColumnX", body)
        self.assertIn("path.W4Y = lowestRowY", body)

    def test_three_segments(self):
        cls = self.read(PATH_CLS)
        self.assertIn("SegmentCount = 3", cls)

    # --- R23-703 ---------------------------------------------------------

    def test_r23_703_crossings_are_proved_against_each_holes_own_radius(
        self,
    ):
        body = self.body("Private Function PathCrossesCircle(")
        self.assertIn("projection.ProjectedRadiusM", body)
        self.assertIn("SegmentReaches(", body)
        # All three segments are tested, not just the last one.
        self.assertEqual(3, body.count("SegmentReaches("))

    def test_segment_distance_is_clamped_to_the_finite_segment(self):
        """An unclamped projection would report a circle beyond an endpoint
        as crossed, because the infinite line passes through it."""
        body = self.body("Private Function SegmentReaches(")
        self.assertIn("If t < 0# Then t = 0#", body)
        self.assertIn("If t > 1# Then t = 1#", body)

    def test_unattempted_crossing_state_is_not_reported_as_a_failure(self):
        """A path rejected before crossings could be tested - no bore, too
        few columns - already reported why. Appending the initial
        "NotAttempted" state alongside those reasons dilutes them."""
        body = self.body("Public Function VerifySectionGeometry(")
        self.assertIn('"NotAttempted", _', body)

    def test_missed_holes_are_named_individually(self):
        body = self.body("Private Sub ProvePathCrossings(", "\nEnd Sub")
        self.assertIn("BoreNotFullyCrossed", body)
        self.assertIn("ColumnHoleNotCrossed:", body)
        self.assertIn("NoHolesOnChosenColumn", body)

    # --- R23-704 ---------------------------------------------------------

    def test_r23_704_no_outline_percentage_strategy_in_the_new_path(self):
        """The old strategy put the upper label in the zone region and the
        lower arrow in the part-identification band, because a fraction of
        an outline knows nothing about where the holes are."""
        for forbidden in (
            "extension",
            "topY",
            "bottomY",
            "leftX",
            "rightX",
            "outline",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

        # The specific fractions the legacy path used.
        for fraction in ("15# / 72#", "90# / 196#", "15.84 / 24#", "0.1 *"):
            with self.subTest(fraction=fraction):
                self.assertNotIn(fraction, self.executable)

    # --- R23-705 ---------------------------------------------------------

    def test_r23_705_frame_conversion_happens_exactly_once_per_waypoint(
        self,
    ):
        """Mixing frames is the defect section work has hit before. Nothing
        upstream holds view coordinates, so there is nothing to convert
        twice."""
        self.assertIn("Public Function PageToViewSketch(", self.source)
        conversion = self.body("Public Function PageToViewSketch(")
        self.assertIn("swView.GetXform", conversion)
        self.assertIn("swView.Angle", conversion)
        self.assertIn("conversions=1", conversion)

        create = self.body("Public Function CreateSectionFromPath(")
        # One conversion call, inside the waypoint loop, before CreateLine.
        self.assertEqual(1, create.count("PageToViewSketch("))
        self.assertLess(
            create.index("PageToViewSketch("), create.index("CreateLine(")
        )

    def test_waypoints_are_page_frame_until_that_conversion(self):
        cls = self.read(PATH_CLS)
        self.assertIn("frame=Page", cls)

    # --- R23-706 ---------------------------------------------------------

    def test_r23_706_segment_selection_order_is_verified(self):
        """SOLIDWORKS reads the segments in selection order, so an
        unverified order produces a differently shaped cut."""
        body = self.body("Public Function CreateSectionFromPath(")
        self.assertIn("GetSelectedObjectCount2(-1)", body)
        self.assertIn("actualCount <> expectedCount + 1", body)
        self.assertIn("SegmentSelectionOrderUnverified", body)
        self.assertIn("SegmentCountMismatch", body)
        self.assertLess(
            body.index("SegmentSelectionOrderUnverified"),
            body.index("CreateSectionViewAt5("),
        )

    # --- R23-707 ---------------------------------------------------------

    def test_r23_707_section_line_info_is_read_back(self):
        self.assertIn("Public Function ReadSectionLineInfo(", self.source)
        body = self.body("Public Function ReadSectionLineInfo(")
        self.assertIn("GetSectionLineInfo2", body)
        self.assertIn("NoGeometryReturned", body)

        create = self.body("Public Function CreateSectionFromPath(")
        self.assertIn("ReadSectionLineInfo(path.SourceView)", create)

    # --- hygiene ---------------------------------------------------------

    def test_source_hygiene(self):
        for name in (MODULE, PATH_CLS):
            with self.subTest(name=name):
                raw = (SOURCE / name).read_bytes()
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
