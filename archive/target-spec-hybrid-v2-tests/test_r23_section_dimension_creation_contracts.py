"""R23-821. Create the section dimensions whose geometry survived.

r59 (macro_qa/20260805_062318_P-0251-14A-001) is the first run whose
section geometry is both trusted and honest about arcs. Five requirements
have measuring geometry that proved selectable, two do not:

    OVERALL_THICKNESS_18       XPair   records 2 / 22   selectable
    BORE_STEP_DEPTH_12         XPair   records 2 / 5    selectable
    INNER_BORE_D40             Radius  record 24        selectable
    FIT_BORE_D47_H7            Radius  record 25        selectable
    LOWER_VERTICAL_REF_104_8   YPair   records 10 / 44  selectable
    LOWER_WALL_STEP_11_5       none - lost its pair when arc tessellation
                               stopped counting as straight edges (r59)
    LONG_VERTICAL_REF_173_6    none - no pair of the nineteen distinct Y
                               values differs by 0.1736

So creation runs for what is proved and refuses for the rest. Five right
dimensions and two honest absences beat seven that include a wrong one:
r58 would have attached LOWER_WALL_STEP_11_5 to the bore's tessellation,
which measures the wrong thing and looks correct on the sheet.

The created dimension is still verified by CreateSectionDimension, which
reads the nominal back and rejects a mismatch as CreatedButRejected. That
readback is what makes a wrong selection or a wrong coordinate frame
visible rather than plausible.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module10_SectionDimensionEngine.bas"
PIPELINE = "Module2_DrawingPipeline.bas"
REQUIREMENT_CLS = "CSectionRequirement.cls"


class R23SectionDimensionCreationContracts(unittest.TestCase):
    def code(self, name: str) -> str:
        text = (SOURCE / name).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def member(self, marker: str, terminator: str = "\nEnd Function") -> str:
        return self.code(MODULE).split(marker)[1].split(terminator)[0]

    def pass_body(self) -> str:
        return self.member(
            "Public Function CreateResolvedSectionDimensions(")

    # --- authorization ----------------------------------------------------

    def test_the_pass_refuses_without_explicit_authorization(self):
        body = self.pass_body()
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertIn("SECTION_DIM_PASS_REFUSED", body)
        # The refusal precedes anything that could select or create.
        self.assertLess(body.index("If Not allowMutation Then"),
                        body.index("SelectCandidateEntities("))

    # --- what it will and will not create ---------------------------------

    def test_reconciliation_wins_over_creation(self):
        """R23-802. An imported dimension that already satisfies a
        requirement must not get a second one created beside it.

        Asserted as the contiguous guard, not as two strings that happen to
        appear somewhere: `Matched` is also read at the end of the loop to
        count what was created, so a looser check passes even with the
        skip removed."""
        body = self.pass_body()
        self.assertIn(
            'If requirement.Matched Then\n'
            '            skipReason = "AlreadySatisfiedByImportedDimension"',
            body,
        )
        self.assertLess(
            body.index("AlreadySatisfiedByImportedDimension"),
            body.index("SelectCandidateEntities("),
        )

    def test_unfound_and_unselectable_candidates_are_skipped_by_name(self):
        body = self.pass_body()
        for reason in ("NoCandidateGeometry:",
                       "CandidateSideANotSelectable",
                       "CandidateSideBNotSelectable"):
            with self.subTest(reason=reason):
                self.assertIn(reason, body)
        self.assertIn("SECTION_DIM_SKIPPED", body)

    def test_the_pipeline_reconciles_before_it_creates(self):
        body = self.code(PIPELINE).split(
            "Private Function ReconcileR23SectionDimensions("
        )[1].split("\nEnd Function")[0]
        self.assertLess(
            body.index("ReconcileSectionDimensions"),
            body.index("CreateResolvedSectionDimensions"),
        )

    def test_the_verdict_is_recomputed_after_creation(self):
        """Verifying against the pre-creation inventory would report the
        drawing as it was, not as it is."""
        body = self.code(PIPELINE).split(
            "Private Function ReconcileR23SectionDimensions("
        )[1].split("\nEnd Function")[0]
        self.assertEqual(2, body.count("InventorySectionDimensions("))
        self.assertLess(
            body.index("CreateResolvedSectionDimensions"),
            body.rindex("InventorySectionDimensions("),
        )
        self.assertLess(
            body.rindex("InventorySectionDimensions("),
            body.index("VerifySectionDimensions("),
        )

    def test_the_requirement_state_is_rederived_after_creation(self):
        """R23-822. Re-reading the dimensions is not enough on its own.

        Reconciliation writes NoImportedDimension into every unmatched
        requirement's Failures BEFORE creation runs, and creation sets
        Matched without clearing it. VerifySectionDimensions treats a
        non-empty Failures list as unsatisfied whatever the counts say, so
        r60 reported satisfied=0 for five dimensions the QA stage
        independently proved satisfied. The requirements must be rebuilt and
        re-reconciled against the finished drawing, as Module19 does."""
        body = self.code(PIPELINE).split(
            "Private Function ReconcileR23SectionDimensions("
        )[1].split("\nEnd Function")[0]

        self.assertEqual(2, body.count("ReconcileSectionDimensions"))
        self.assertLess(
            body.index("CreateResolvedSectionDimensions"),
            body.rindex("ReconcileSectionDimensions"),
        )

        # Rebuilt, not reused: reusing the objects creation mutated is the
        # defect itself. (The first BuildSectionRequirements call is the
        # pre-creation fallback for an inventory that returned nothing.)
        self.assertEqual(2, body.count("BuildSectionRequirements()"))
        self.assertLess(
            body.index("CreateResolvedSectionDimensions"),
            body.rindex("BuildSectionRequirements()"),
        )
        self.assertIn(
            "Module10_SectionDimensionEngine.ReconcileSectionDimensions _\n"
            "        finalRequirements, sectionDimensions, evidence",
            body,
        )
        # And the verdict reads the rebuilt set, not the stale one.
        self.assertIn(
            "verdict = Module10_SectionDimensionEngine."
            "VerifySectionDimensions( _\n"
            "        finalRequirements, sectionDimensions)",
            body,
        )

    # --- the selection --------------------------------------------------

    def test_the_selection_count_is_verified_before_creating(self):
        """SOLIDWORKS reads the selection list. An unverified count
        produces a dimension between the wrong things."""
        body = self.member("Private Function SelectCandidateEntities(")
        self.assertIn("expected = 1", body)
        self.assertIn("expected = 2", body)
        self.assertIn("GetSelectedObjectCount2(-1)", body)
        self.assertIn("actual <> expected", body)
        self.assertIn("SECTION_DIM_SELECT_FAILED", body)

    def test_the_second_entity_is_appended_not_replacing(self):
        body = self.member("Private Function SelectCandidateEntities(")
        self.assertIn(
            "SelectOneCandidate( _\n        swView, entities, "
            "requirement.CandidateRecordA, False)",
            body,
        )
        self.assertIn("requirement.CandidateRecordB, True)", body)

    def test_the_selection_is_cleared_around_every_creation(self):
        body = self.pass_body()
        self.assertIn("swDrawModel.ClearSelection2 True", body)
        select = self.member("Private Function SelectCandidateEntities(")
        self.assertIn("swDrawModel.ClearSelection2 True", select)

    # --- placement --------------------------------------------------------

    def test_the_lane_becomes_a_coordinate_in_exactly_one_place(self):
        """R23-808. A lane is a NAME everywhere else in this module."""
        body = self.code(MODULE)
        self.assertEqual(1, body.count("Private Sub LaneTextPoint("))
        lane = self.member("Private Sub LaneTextPoint(", "\nEnd Sub")
        self.assertIn("swView.GetOutline", lane)
        self.assertIn("Module8_RuntimeSupport.LAYOUT_MARGIN_M", lane)

    def test_placement_introduces_no_new_literal_or_sheet_fraction(self):
        """The disproved Phase 7 strategy placed things at a percentage of
        an outline. This derives from the view's own measured outline plus
        the existing layout margin."""
        lane = self.member("Private Sub LaneTextPoint(", "\nEnd Sub")
        for forbidden in ("0.1 *", "* 0.75", "* 0.5", "SheetWidth",
                          "SheetHeight"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, lane)

    def test_placement_is_recorded_with_what_it_was_derived_from(self):
        body = self.pass_body()
        self.assertIn("SECTION_DIM_PLACEMENT", body)
        lane = self.member("Private Sub LaneTextPoint(", "\nEnd Sub")
        self.assertIn("placement=ViewOutlinePlusMargin", lane)
        self.assertIn("|outline=", lane)
        self.assertIn("|autoArrangeRunsAfter=True", lane)

    def test_the_ordinal_is_counted_per_lane(self):
        """R23-822. A global counter grew the gap 12/24/36/48/60 mm across
        five different lanes, so the fifth dimension started 60 mm off its
        view for no reason connected to its own lane."""
        body = self.pass_body()
        self.assertIn(
            "laneOrdinal = NextLaneOrdinal(laneOrdinals, requirement.Lane)",
            body,
        )
        self.assertNotIn("laneOrdinal = laneOrdinal + 1", body)
        self.assertIn("Set laneOrdinals = New Collection", body)

        counter = self.member("Private Function NextLaneOrdinal(")
        # Keyed by lane, and the key is replaced rather than re-Added,
        # because Collection.Add on an existing key raises 457.
        self.assertIn("laneOrdinals.Item(key)", counter)
        self.assertIn("laneOrdinals.Remove key", counter)
        self.assertIn("laneOrdinals.Add current, key", counter)
        self.assertLess(
            counter.index("laneOrdinals.Remove key"),
            counter.index("laneOrdinals.Add current, key"),
        )
        # An empty key raises 5 on Add; placement bookkeeping must not take
        # the creation pass down with it.
        self.assertIn('If Len(key) = 0 Then key = "UnnamedLane"', counter)

    def test_placement_is_clamped_into_the_proved_usable_area(self):
        """R23-822. The r60 starting point was unbounded: RD2 began above
        the usable area, auto-arrange finished the move, and it landed in
        the zoned border - ANNOTATION_EXTENTS regressed PROVED to FAILED.

        Asserted as the four contiguous clamps, so removing any one side
        fails."""
        lane = self.member("Private Sub LaneTextPoint(", "\nEnd Sub")
        for clamp in (
            "If textX < evidence.UsableLeft Then textX = evidence.UsableLeft",
            "If textX > evidence.UsableRight Then "
            "textX = evidence.UsableRight",
            "If textY < evidence.UsableBottom Then\n"
            "            textY = evidence.UsableBottom",
            "If textY > evidence.UsableTop Then textY = evidence.UsableTop",
        ):
            with self.subTest(clamp=clamp.split("\n")[0]):
                self.assertIn(clamp, lane)

        # Clamped before the point is reported, not after.
        self.assertLess(
            lane.index("evidence.UsableLeft"),
            lane.index("placement=ViewOutlinePlusMargin"),
        )
        self.assertIn("|clamped=", lane)

    def test_unproved_bounds_stop_the_stack_instead_of_guessing(self):
        """Fail closed. Stacking an ordinal against a boundary that was
        never measured is how the point left the sheet in the first place,
        so without proved bounds the gap stays at one margin."""
        lane = self.member("Private Sub LaneTextPoint(", "\nEnd Sub")
        self.assertIn("boundsProven = evidence.LayoutBoundariesProven", lane)
        self.assertIn("If Not boundsProven Then effectiveOrdinal = 1", lane)
        self.assertIn("usable=Unproved", lane)
        # The clamp itself only runs on proved bounds.
        self.assertLess(
            lane.index("If boundsProven Then"),
            lane.index("evidence.UsableLeft"),
        )
        self.assertLess(
            lane.index("If Not boundsProven Then effectiveOrdinal = 1"),
            lane.index(
                "gap = Module8_RuntimeSupport.LAYOUT_MARGIN_M"),
        )

    # --- the candidate record survives onto the requirement ---------------

    def test_the_requirement_carries_its_candidate(self):
        cls = self.code(REQUIREMENT_CLS)
        for field in ("Public CandidateFound As Boolean",
                      "Public CandidateKind As String",
                      "Public CandidateRecordA As Long",
                      "Public CandidateRecordB As Long",
                      "Public CandidateSelectableA As Boolean",
                      "Public CandidateSelectableB As Boolean"):
            with self.subTest(field=field):
                self.assertIn(field, cls)

        self.assertIn("CandidateRecordA = -1", cls)
        self.assertIn("CandidateRecordB = -1", cls)

    def test_the_created_dimension_is_still_verified(self):
        """CreateSectionDimension reads the nominal back and rejects a
        mismatch. That is what makes a wrong selection visible instead of
        plausible, so the creation pass must not bypass it."""
        body = self.pass_body()
        self.assertIn("CreateSectionDimension( _", body)
        creator = self.member("Public Function CreateSectionDimension(")
        self.assertIn("CreatedButRejected", creator)
        self.assertIn("SECTION_DIM_CREATE_REJECTED", creator)


if __name__ == "__main__":
    unittest.main()
