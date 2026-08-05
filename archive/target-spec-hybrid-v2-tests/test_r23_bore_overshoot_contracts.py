"""R23-813. The cut must pass through the whole bore, and prove it did.

Measured in run macro_qa/20260805_041027_P-0251-14A-001. The r50 geometry
inventory read Section View J-J and found no bore in it:

    R = 0.002100;0.035000;0.036000      <- M5 tap drill, plate top profile
    Y = 0.019;-0.097;0.062;0.098;0.018;0.017;-0.098
    X = 0.008;0.002;-0.008;0.009;0.004;0.003;-0.009

Neither 0.020 nor 0.023500 appears as a radius and no 0.040 or 0.047 span
exists on either axis, so INNER_BORE_D40 and FIT_BORE_D47_H7 had no
candidate geometry at all. The bore's cut walls would sit at
Y = 0.062 +/- 0.023500; the only nearby coordinate is 0.062 itself - the
centre, which is where waypoint 1 sat. A cut starting at the centre removes
half the bore and leaves no opening to dimension.

Two things were wrong and both are fixed here:

1. Waypoint 1 sat ON the bore centre. It now sits one full radius beyond
   the far wall, an overshoot derived from the bore's own projected radius,
   in the direction away from the face-hole rows.
2. The crossing predicate could not tell the difference.
   PathCrossesCircle asks only whether a segment comes within the radius,
   which a segment starting at the centre satisfies trivially - it reported
   crossingsProven=4|crossingFailures=None for the very cut that produced a
   section with no bore in it. The bore now needs a full crossing: both
   intersections of the line with the circle inside the same segment.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module17_SectionPath.bas"
PATH_CLS = "CSectionPath.cls"


class R23BoreOvershootContracts(unittest.TestCase):
    def code(self, name: str) -> str:
        text = (SOURCE / name).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def body(self, marker: str, terminator: str = "\nEnd Function") -> str:
        return self.code(MODULE).split(marker)[1].split(terminator)[0]

    # --- the overshoot ----------------------------------------------------

    def test_waypoint_one_is_no_longer_the_bare_bore_centre(self):
        body = self.body("Public Function ResolveSectionPath(")
        self.assertNotIn(
            "path.W1Y = path.BoreProjection.PageY\n", body + "\n")

    def test_the_overshoot_is_derived_from_the_bore_itself(self):
        """Twice the projected radius from the centre is one full radius
        past the far wall. Same principle as the R23-703 crossing slack: a
        feature is judged on its own size, never against a shared
        literal."""
        body = self.body("Public Function ResolveSectionPath(")
        self.assertIn(
            "path.BoreOvershootM = 2# * path.BoreProjection"
            ".ProjectedRadiusM",
            body,
        )

    def test_the_overshoot_direction_is_read_from_the_geometry(self):
        """Which way is "past the bore" depends on where the rest of the
        path goes, not on an assumption about which way up the part sits."""
        body = self.body("Public Function ResolveSectionPath(")
        self.assertIn(
            "If path.BoreProjection.PageY >= highestRowY Then", body)
        self.assertIn(
            "path.W1Y = path.BoreProjection.PageY + path.BoreOvershootM",
            body,
        )
        self.assertIn(
            "path.W1Y = path.BoreProjection.PageY - path.BoreOvershootM",
            body,
        )

    def test_the_overshoot_is_recorded_and_reportable(self):
        body = self.body("Public Function ResolveSectionPath(")
        self.assertIn("SECTION_PATH_BORE_OVERSHOOT", body)
        self.assertIn("|reason=CutMustCrossWholeBore", body)

        path_cls = self.code(PATH_CLS)
        self.assertIn("Public BoreOvershootM As Double", path_cls)
        self.assertIn('"|boreOvershootM="', path_cls)

    def test_no_outline_fraction_was_introduced_by_the_overshoot(self):
        """The disproved strategy extended waypoints by a percentage of the
        view outline. This one never reads the outline."""
        body = self.body("Public Function ResolveSectionPath(")
        for forbidden in ("GetOutline", "outline", "UsableTop",
                          "UsableBottom"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    # --- the stronger predicate -------------------------------------------

    def test_the_bore_now_requires_a_full_crossing(self):
        body = self.code(MODULE).split(
            "Private Sub ProvePathCrossings(")[1].split("\nEnd Sub")[0]
        self.assertIn(
            "If Not PathFullyCrossesCircle(path, path.BoreProjection) Then",
            body,
        )
        self.assertIn("BoreNotFullyCrossed", body)

    def test_column_holes_deliberately_keep_the_weaker_test(self):
        """The first and last hole on the chosen column sit AT the
        segment's endpoints, so demanding a full crossing of those would
        refuse the correct path."""
        body = self.code(MODULE).split(
            "Private Sub ProvePathCrossings(")[1].split("\nEnd Sub")[0]
        self.assertIn("If PathCrossesCircle(path, candidate) Then", body)
        self.assertNotIn("PathFullyCrossesCircle(path, candidate)", body)

    def test_full_crossing_tests_all_three_segments(self):
        body = self.body("Private Function PathFullyCrossesCircle(")
        self.assertEqual(3, body.count("SegmentSpansCircle("))
        self.assertIn("projection.ProjectedRadiusM + CROSSING_SLACK_M",
                      body)

    def test_full_crossing_requires_both_intersections_inside_the_segment(
        self,
    ):
        """Entry and exit both inside the finite segment. Half a chord
        either side of the perpendicular foot must fit within it - that is
        the whole difference from the reaches test."""
        body = self.body("Private Function SegmentSpansCircle(")
        self.assertIn("If perpendicular > radius Then Exit Function", body)
        self.assertIn("halfChord = Sqr((radius * radius) - "
                      "(perpendicular * perpendicular))", body)
        self.assertIn("(footDistance - halfChord >= 0#) And", body)
        self.assertIn("(footDistance + halfChord <= segmentLength)", body)

    def test_the_weaker_predicate_still_exists_unchanged(self):
        """PathCrossesCircle is not deleted or altered; the bore simply no
        longer uses it. Column holes still do."""
        body = self.body("Private Function PathCrossesCircle(")
        self.assertEqual(3, body.count("SegmentReaches("))
        self.assertIn("projection.ProjectedRadiusM + CROSSING_SLACK_M",
                      body)


if __name__ == "__main__":
    unittest.main()
