"""Route D is gated on Route A declining, not on the inventory's absence.

Root 1 of the post-1845 defect review. The stepped bore in P-0251 has no
accepted projection, so ResolveBoreProjection finds no singleton bore, no
section path resolves, and SECTION_GEOMETRY, SECTION_DIMENSIONS,
SECTION_CLEARANCE and LAYOUT all fail behind it.

Measured in every production run to date, Drawing View1:

    six counterbores  mappedEdges=2  inventoryConfirmed=2
    stepped bore      mappedEdges=0  firstReject=NoRouteMappedThisEdge:
                                     A:Nothing:err0;B:Nothing:err0

Route A declines for the bore in a view where it demonstrably works for other
geometry, and Drawing View1 HAS a visible inventory (39 edges), so the old
gate "Not visibleInventoryAvailable" meant Route D was never attempted there.

Re-keying it is not a weakening. When an inventory exists, a Route-D entity
must still clear the Route C membership check immediately below, so it is
proved twice - ownership through
ISelectionMgr.GetSelectedObjectsDrawingView2 AND membership of
IView.GetVisibleEntities2 - where the inventory-less path proves only
ownership. A genuinely obscured edge cannot be in the visible inventory and
is still rejected as MappedEntityNotInVisibleInventory.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module13_ProjectionResolution.bas"


class R23RouteDGateContracts(unittest.TestCase):
    def code(self) -> str:
        text = (SOURCE / MODULE).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def resolve_body(self) -> str:
        return self.code().split("Private Sub ResolveProjection(")[1].split(
            "\nEnd Sub"
        )[0]

    def test_route_d_is_no_longer_gated_on_a_missing_inventory(self):
        body = self.resolve_body()
        self.assertNotIn(
            "If mapped Is Nothing And Not visibleInventoryAvailable Then",
            body,
        )

    def test_route_d_runs_whenever_the_earlier_routes_decline(self):
        body = self.resolve_body()
        gate = body.split("SelectModelEntityInView(")[0]
        # The nearest preceding condition is the plain nothing-mapped test.
        self.assertTrue(
            gate.rstrip().endswith("Set mapped = ")
            or "If mapped Is Nothing Then" in gate.rsplit("\n    ", 1)[-1]
            or "If mapped Is Nothing Then" in gate,
            "Route D is not reached from a plain 'mapped Is Nothing' test",
        )

    def test_route_c_membership_check_still_follows_route_d(self):
        """This is what makes the wider gate safe: an inventory-bearing view
        proves the entity twice."""
        body = self.resolve_body()
        self.assertLess(
            body.index("SelectModelEntityInView("),
            body.index("FindVisibleEntityIndex("),
        )
        self.assertIn("MappedEntityNotInVisibleInventory", body)

    def test_membership_check_is_still_conditional_on_an_inventory(self):
        body = self.resolve_body()
        self.assertIn("If visibleInventoryAvailable Then", body)

    def test_route_d_still_proves_target_view_ownership(self):
        """The selection route is narrower than a coordinate search and must
        stay that way."""
        selector = self.code().split("Private Function SelectModelEntityInView(")[
            1
        ].split("\nEnd Function")[0]
        self.assertIn("GetSelectedObjectsDrawingView2", selector)
        self.assertIn("D:RefusedPreexistingSelection:count", selector)
        self.assertIn("ClearSelection2", selector)

    def test_no_coordinate_search_was_introduced(self):
        """A bore could be 'found' by page position; that is exactly the
        unproved shortcut this module refuses."""
        body = self.resolve_body()
        for forbidden in ("nearestEntity", "ClosestEntity", "SelectByID2",
                          "searchRadius"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)


if __name__ == "__main__":
    unittest.main()
