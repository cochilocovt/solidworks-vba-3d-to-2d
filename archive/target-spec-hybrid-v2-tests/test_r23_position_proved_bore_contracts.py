"""The section path may use a position-proved bore; nothing else may.

User decision, 2026-08-05, on Root 1 of the post-1845 defect review.

Run macro_qa/20260805_033146_P-0251-14A-001 proved the stepped bore of
P-0251 is completely obscured in every orthographic view:

    mappedEdges=4  inventoryConfirmed=0
    firstReject=MappedEntityNotInVisibleInventory

All four circular edges map through Route D and
ISelectionMgr.GetSelectedObjectsDrawingView2 proves Drawing View1 owns them,
yet none is in IView.GetVisibleEntities2, which by contract holds only
entities "not completely obscured by other entities in the view". The
counterbores in the same pass read mappedEdges=2|inventoryConfirmed=2, which
is the positive control.

So no orthographic anchor for that bore can ever exist, and the J-J section
that would finally show it demanded one. Circular. The resolution splits a
requirement that was conflated: the section path reads PageX, PageY and
ProjectedRadiusM and never selects the bore, so a proved page POSITION is
what it actually needs. Dimensioning, callout attachment and ordinate
anchoring do select, and keep requiring a selectable anchor.

These tests hold that split in place. The weaker proof must stay confined to
the section path, and the anchor gate must stay exactly as strict as it was.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
PROJECTION_CLS = "CViewHoleProjection.cls"
SECTION_MODULE = "Module17_SectionPath.bas"


class R23PositionProvedBoreContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def member(self, name: str, marker: str,
               terminator: str = "\nEnd Function") -> str:
        return self.code(name).split(marker)[1].split(terminator)[0]

    # --- the weaker proof itself ----------------------------------------

    def test_position_proof_exists_and_is_separate(self):
        source = self.code(PROJECTION_CLS)
        self.assertIn("Public Function PositionFailureReason() As String",
                      source)
        self.assertIn("Public Function HasProvedPosition() As Boolean", source)

    def test_position_proof_does_not_require_an_anchor(self):
        """This is the whole point of it. An obscured edge has no anchor and
        never will."""
        body = self.member(PROJECTION_CLS,
                           "Public Function PositionFailureReason() As String")
        self.assertNotIn("HasSelectableAnchor", body)
        self.assertNotIn("PrimaryAnchor", body)
        self.assertNotIn("ProjectionAnchorUnavailable", body)

    def test_position_proof_still_requires_everything_else(self):
        """Weaker than the anchor gate is not the same as unproved. Identity,
        configuration, page frame and circular projection all still hold."""
        body = self.member(PROJECTION_CLS,
                           "Public Function PositionFailureReason() As String")
        for required in ("PhysicalLocation Is Nothing",
                         "PhysicalInstanceKey",
                         "Not ConfigurationProven",
                         "Not AxisNormalToView",
                         "CoordinateFrameProof"):
            with self.subTest(required=required):
                self.assertIn(required, body)

    def test_axis_normal_is_required_because_a_centre_needs_a_circle(self):
        """A hole seen edge-on does not project as a circle, so its 'centre'
        would be meaningless as a waypoint. Drawing View2 rejects the same
        bore on exactly this test."""
        body = self.member(PROJECTION_CLS,
                           "Public Function PositionFailureReason() As String")
        self.assertIn("AxisNotNormalToView", body)

    # --- the anchor gate is untouched ------------------------------------

    def test_anchor_gate_still_requires_a_selectable_anchor(self):
        body = self.member(
            PROJECTION_CLS,
            "Public Function QualificationFailureReason() As String")
        self.assertIn("Not HasSelectableAnchor()", body)
        self.assertIn("ProjectionAnchorUnavailable", body)

    def test_acceptance_still_runs_through_the_anchor_gate(self):
        """Module13 decides Accepted. If it ever consulted the position proof
        instead, an obscured edge would become a dimension anchor."""
        module13 = self.code("Module13_ProjectionResolution.bas")
        self.assertIn("failureReason = projection.QualificationFailureReason()",
                      module13)
        self.assertNotIn("PositionFailureReason", module13)
        self.assertNotIn("HasProvedPosition", module13)

    # --- confinement -----------------------------------------------------

    def test_only_the_section_path_may_use_the_weaker_proof(self):
        """The confinement IS the decision. Any other caller would be reusing
        a page position where a selection is required."""
        allowed = {PROJECTION_CLS, SECTION_MODULE}
        for path in sorted(SOURCE.glob("*.bas")) + sorted(SOURCE.glob("*.cls")):
            if path.name in allowed:
                continue
            with self.subTest(component=path.name):
                text = self.code(path.name)
                self.assertNotIn("HasProvedPosition", text)
                self.assertNotIn("PositionFailureReason", text)

    def test_the_section_path_never_selects_the_bore(self):
        """The justification for the weaker proof is that this module reads
        coordinates and does not select. Hold that true."""
        section = self.code(SECTION_MODULE)
        self.assertNotIn("PrimaryAnchor", section)
        self.assertNotIn("SelectEntity", section)

    # --- the bore resolver ------------------------------------------------

    def test_bore_resolver_admits_a_position_proved_candidate(self):
        body = self.member(SECTION_MODULE,
                           "Private Function ResolveBoreProjection(")
        self.assertIn("If Not candidate.Accepted Then", body)
        self.assertIn("If Not candidate.HasProvedPosition() Then", body)

    def test_bore_resolver_still_requires_a_singleton_family(self):
        """Widening the proof must not widen WHICH location is the bore."""
        body = self.member(SECTION_MODULE,
                           "Private Function ResolveBoreProjection(")
        self.assertIn("graph.LocationsForFamily(", body)
        self.assertIn("family.Count <> 1", body)
        self.assertIn("MaximumRadiusM", body)

    def test_bore_resolver_reports_which_proof_it_used(self):
        body = self.member(SECTION_MODULE,
                           "Private Function ResolveBoreProjection(")
        self.assertIn('basis = "Accepted"', body)
        self.assertIn('basis = "PositionProved:"', body)

    def test_a_position_proved_bore_is_stated_in_evidence(self):
        body = self.member(SECTION_MODULE,
                           "Public Function ResolveSectionPath(")
        self.assertIn("SECTION_PATH_BORE_BASIS", body)
        self.assertIn("|use=WaypointsOnly", body)

    def test_the_path_record_carries_the_basis(self):
        cls = self.code("CSectionPath.cls")
        self.assertIn("Public BoreProjectionBasis As String", cls)
        self.assertIn('BoreProjectionBasis = "NotResolved"', cls)
        self.assertIn('"|boreBasis=" & BoreProjectionBasis', cls)

    def test_rejection_reason_no_longer_claims_acceptance_is_required(self):
        body = self.member(SECTION_MODULE,
                           "Public Function ResolveSectionPath(")
        self.assertIn('"NoUsableSingletonBoreProjection"', body)
        self.assertNotIn("NoAcceptedSingletonBoreProjection", body)

    # --- face holes are NOT loosened --------------------------------------

    def test_face_holes_still_require_acceptance(self):
        """All six counterbores are accepted in Drawing View1, so there is no
        evidence that this needs loosening and it stays as it is."""
        body = self.member(SECTION_MODULE,
                           "Private Function CollectFaceHoleProjections(")
        self.assertIn("If Not candidate.Accepted Then GoTo ContinueCandidate",
                      body)
        self.assertNotIn("HasProvedPosition", body)

    # --- no shortcut smuggled in ------------------------------------------

    def test_no_coordinate_search_was_introduced(self):
        """A position-proved projection is one the projector computed and
        proved, not one found by looking near a page point."""
        section = self.code(SECTION_MODULE)
        for forbidden in ("nearestEntity", "ClosestEntity", "SelectByID2",
                          "searchRadius"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, section)


if __name__ == "__main__":
    unittest.main()
