"""R23-819. Map a measured section coordinate back to a selectable entity.

r57 (macro_qa/20260805_055822_P-0251-14A-001) produced the first clean
decode of Section View J-J and proved six of the seven requirements have
their measuring geometry in the view - including radius 0.020 for
INNER_BORE_D40 and 0.023500 for FIT_BORE_D47_H7, both
`geometryTrusted=True`.

Geometry is not the blocker. `CreateSectionDimension` refuses with
`reason=NoEntitiesSelected`, so what it needs is a SELECTED drawing entity,
and nothing connected a measured coordinate to one.

The connection already exists in the data: GetPolylines7's return value is
the entity array, positionally paired with the polyline records, which r57
confirmed exactly (`records=79|entities=79|recordsMatchEntities=True`). So
each distinct coordinate records the index of the polyline that produced
it, and that index reaches the entity.

Whether those entries are live in a CUT view is the open question - many
section curves are cut faces with no model edge behind them, and the Help
says the array carries Null in their place. That is measured here, not
assumed, using the route Module13 already proved for orthographic views:
IView.SelectEntity, then ISelectionMgr.GetSelectedObjectsDrawingView2 to
prove the owning view before the object is trusted.

Nothing is created. This iteration answers "can it be selected", and the
answer decides whether dimension creation is even reachable.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module10_SectionDimensionEngine.bas"
PIPELINE = "Module2_DrawingPipeline.bas"


class R23SectionEntitySelectionContracts(unittest.TestCase):
    def code(self, name: str) -> str:
        text = (SOURCE / name).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def member(self, marker: str, terminator: str) -> str:
        return self.code(MODULE).split(marker)[1].split(terminator)[0]

    def selector(self) -> str:
        return self.member(
            "Private Function ProveSectionEntitySelection(", "\nEnd Function")

    # --- the record-to-entity pairing -------------------------------------

    def test_every_distinct_coordinate_remembers_its_record(self):
        body = self.member("Private Sub AddDistinctValue(", "\nEnd Sub")
        self.assertIn("ByRef records As Collection", body)
        self.assertIn("records.Add recordIndex", body)
        # The value and its record are appended together or not at all.
        self.assertLess(body.index("values.Add value"),
                        body.index("records.Add recordIndex"))

    def test_all_three_coordinate_families_carry_records(self):
        body = self.code(MODULE)
        for call in ("AddDistinctValue radii, radiusRecords",
                     "AddDistinctValue horizontalYs, horizontalYRecords",
                     "AddDistinctValue verticalXs, verticalXRecords"):
            with self.subTest(call=call):
                self.assertIn(call, body)

    def test_the_matchers_return_the_records_they_matched(self):
        near = self.member("Private Function FindValueNear(",
                           "\nEnd Function")
        self.assertIn("foundRecord = CLng(records(i))", near)

        pair = self.member("Private Function FindPairWithSpan(",
                           "\nEnd Function")
        self.assertIn("firstRecord = CLng(records(i))", pair)
        self.assertIn("secondRecord = CLng(records(j))", pair)

    def test_the_candidate_row_reports_both_records(self):
        body = self.member("Private Function ReportRequirementCandidates(",
                           "\nEnd Sub")
        self.assertIn('"|recordA="', body)
        self.assertIn('"|recordB="', body)
        self.assertIn("firstRecord = -1", body)
        self.assertIn("secondRecord = -1", body)

    def test_only_polyline_records_supply_straight_edge_coordinates(self):
        """R23-820. An arc record carries a tessellated point array as well
        as its GeomData. Walking it produced a false straight edge in r58:
        record 25 supplied BOTH the 0.023500 radius for FIT_BORE_D47_H7 and
        the x=0.008500 side of LOWER_WALL_STEP_11_5, because a chord of the
        bore's tessellation runs parallel to an axis. A linear dimension
        attached there would have measured the bore and looked correct on
        the sheet."""
        body = self.member("Public Sub InventorySectionGeometry(",
                           "\nEnd Sub")
        self.assertIn(
            "If pointIndex > 0 And recordType = GEOM_TYPE_POLYLINE Then",
            body,
        )
        self.assertIn("arcTessellationSegments = arcTessellationSegments",
                      body)
        self.assertIn('"|arcTessellation="', body)

    def test_the_bounding_box_still_sees_every_point(self):
        """Excluding arc tessellation from EDGE classification must not
        shrink the measured extent of the view."""
        body = self.member("Public Sub InventorySectionGeometry(",
                           "\nEnd Sub")
        bounds = body.split("If Not boundsSeeded Then")[1].split(
            "If pointIndex > 0")[0]
        self.assertIn("If px < minX Then minX = px", bounds)
        self.assertIn("If py > maxY Then maxY = py", bounds)

    # --- the selection proof ---------------------------------------------

    def test_the_selector_follows_the_proved_route_d_pattern(self):
        body = self.selector()
        self.assertIn("swView.SelectEntity(candidate, False)", body)
        self.assertIn("GetSelectedObjectsDrawingView2(1, -1)", body)
        self.assertIn("WrongOwner:", body)

    def test_the_selector_refuses_a_preexisting_selection(self):
        """SelectEntity with AppendFlag=False replaces whatever is
        selected. Same refusal as Module13; a caller's selection is not
        ours to discard."""
        body = self.selector()
        self.assertIn("initialCount = selectionMgr.GetSelectedObjectCount2",
                      body)
        self.assertIn("RefusedPreexistingSelection:count", body)

    def test_a_null_entity_is_a_measurement_not_an_error(self):
        """The Help says the entity array carries Null where a polyline
        renders something no edge backs. In a cut view that is expected for
        section faces."""
        body = self.selector()
        self.assertIn("If candidate Is Nothing Then", body)
        self.assertIn('route = "EntityIsNothing"', body)

    def handler(self) -> str:
        """The error handler only. Split on the line-start label: the
        route strings contain "SelectEntityFailed:", which a bare
        "Failed:" split matches first."""
        return self.selector().split("\nFailed:")[1]

    def test_the_selection_is_cleared_on_every_path(self):
        body = self.selector()
        self.assertIn("CleanUp:", body)
        self.assertIn("swDrawModel.ClearSelection2 True", body)
        self.assertIn("ClearSelection2 True", self.handler())

    def test_the_record_index_is_bounds_checked_against_the_array(self):
        body = self.selector()
        self.assertIn("If entityIndex > UBound(entities) Then", body)
        self.assertIn("RecordIndexBeyondEntityArray", body)

    def test_the_error_capture_precedes_nothing_that_resets_err(self):
        handler = self.handler()
        self.assertIn("selectionErrorNumber = Err.Number", handler)
        self.assertLess(handler.index("selectionErrorNumber = Err.Number"),
                        handler.index("ClearSelection2"))

    # --- what is attempted, and what is not -------------------------------

    def test_selection_is_only_attempted_on_trusted_found_geometry(self):
        """At most twelve selections: two per found requirement, and only
        when the decode was clean."""
        body = self.member("Private Function ReportRequirementCandidates(",
                           "\nEnd Sub")
        self.assertIn("If found And geometryTrusted Then", body)
        self.assertIn("If secondRecord <> firstRecord Then", body)

    def test_the_verdict_is_recorded_per_requirement_and_side(self):
        body = self.member("Private Function ReportCandidateSelectability(",
                           "\nEnd Sub")
        self.assertIn("SECTION_ENTITY_SELECT", body)
        self.assertIn("|key=", body)
        self.assertIn("|side=", body)
        self.assertIn("|selectable=", body)
        self.assertIn("|route=", body)

    def test_nothing_is_created_this_iteration(self):
        """The question is whether the geometry can be selected. Creating a
        dimension before that is answered would be guessing again."""
        body = self.member("Public Sub InventorySectionGeometry(",
                           "\nEnd Sub")
        for forbidden in ("AddDimension2", "CreateSectionDimension",
                          "ApplyReferenceFit", "SetFitValues"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_the_pipeline_supplies_the_drawing_document(self):
        """The selection manager lives on ModelDoc2, not DrawingDoc."""
        body = self.code(PIPELINE).split(
            "Private Function ReconcileR23SectionDimensions("
        )[1].split("\nEnd Function")[0]
        self.assertIn("Set swDrawModel = swDraw", body)
        self.assertIn(
            "InventorySectionGeometry _\n        swDrawModel, "
            "sectionViews(1), evidence",
            body,
        )


if __name__ == "__main__":
    unittest.main()
