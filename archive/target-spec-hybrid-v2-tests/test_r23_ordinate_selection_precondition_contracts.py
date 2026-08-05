"""Ordinate creation must start from a clean selection.

Measured in run macro_qa/20260804_235837_P-0251-14A-001. The datum proof
refuses to select while anything is already selected, because
IView.SelectEntity with AppendFlag=False would destroy an interactive
caller's selection. Annotation import runs immediately before the ordinate
stage and leaves its last inserted annotation selected, so every scheme was
refused:

    ORDINATE_DATUM|...|resolved=False|...|selection=Reject
        |reason=PreexistingSelection|initialSelectionCount=1
    ORDINATE_CREATE_REFUSED|...|reason=DatumNotProven

The same two datums resolved in the later QA pass, where nothing was
selected, with identical geometry and an identical anchor
(anchorSource=SolidBodyEdge_0_49, visibleEntityIndex:23). The refusal was
about selection state, not geometry.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"


class R23OrdinateSelectionPreconditionContracts(unittest.TestCase):
    def code(self, name: str) -> str:
        text = (SOURCE / name).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def ordinate_stage(self) -> str:
        return self.code("Module2_DrawingPipeline.bas").split(
            "Private Function CreateR23OrdinateGroups("
        )[1].split("\nEnd Function")[0]

    def test_stage_clears_a_stale_selection_before_building_schemes(self):
        body = self.ordinate_stage()
        self.assertIn("GetSelectedObjectCount2(-1)", body)
        self.assertIn("ClearSelection2 True", body)
        self.assertLess(
            body.index("ClearSelection2 True"),
            body.index("BuildSchemesForView("),
        )

    def test_precondition_is_recorded_as_evidence(self):
        body = self.ordinate_stage()
        self.assertIn("R23_ORDINATE_SELECTION_PRECONDITION", body)
        self.assertIn("|preexisting=", body)
        self.assertIn("|remaining=", body)

    def test_clearing_is_recorded_as_a_mutation(self):
        body = self.ordinate_stage()
        self.assertIn(
            'RecordSolidWorksMutation "ClearSelection2:BeforeOrdinates"', body)

    def test_the_guard_itself_is_not_weakened(self):
        """Clearing belongs to the production pipeline, which owns the
        document it just created. The refusal must remain in Module13 so the
        read-only probes still fail closed on a user's selection."""
        module13 = self.code("Module13_ProjectionResolution.bas")
        self.assertIn("D:RefusedPreexistingSelection:count", module13)
        self.assertIn("If initialSelectionCount <> 0 Then", module13)

    def projection_stage(self) -> str:
        return self.code("Module2_DrawingPipeline.bas").split(
            "Private Function BuildAllViewProjections("
        )[1].split("\nEnd Function")[0]

    def test_projection_stage_clears_a_stale_selection_too(self):
        """Route D became eligible at r47 and was refused for every location
        in the sheet with D:RefusedPreexistingSelection:count1 - view
        creation leaves one object selected, and one stale selection blocked
        all eleven."""
        body = self.projection_stage()
        self.assertIn("GetSelectedObjectCount2(-1)", body)
        self.assertIn("ClearSelection2 True", body)
        self.assertIn("R23_PROJECTION_SELECTION_PRECONDITION", body)
        self.assertIn(
            'RecordSolidWorksMutation "ClearSelection2:BeforeProjections"',
            body,
        )

    def test_projection_clear_precedes_the_mapping_loop(self):
        body = self.projection_stage()
        self.assertLess(
            body.index("ClearSelection2 True"),
            body.index("BuildViewProjections("),
        )

    def test_route_d_leaves_the_selection_clean_after_a_success(self):
        """One clear at the top of the stage is only sufficient because each
        successful Route D clears its own temporary selection."""
        selector = self.code("Module13_ProjectionResolution.bas").split(
            "Private Function SelectModelEntityInView("
        )[1].split("\nEnd Function")[0]
        success = selector.split("Set SelectModelEntityInView = selectedEntity")[1]
        self.assertIn("ClearSelection2 True", success.split("Exit Function")[0])

    def test_probe_runner_does_not_clear_selection_for_the_caller(self):
        module15 = self.code("Module15_OrdinateScheme.bas")
        probe = module15.split("Public Sub R23_ProbeOrdinateScheme()")[
            1
        ].split("\nEnd Sub")[0]
        self.assertNotIn("ClearSelection2 True", probe.split("Failed:")[0])


if __name__ == "__main__":
    unittest.main()
