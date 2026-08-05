"""IView.GetPolylines7 returns DRAWING-space entities on this build.

Measured in run macro_qa/20260804_235542_P-0251-14A-001. The same comparison
helper, run twice against the same GetPolylines7 result on twelve edges that
GetVisibleEntities2 had independently confirmed present:

    mappedVisibleEdges=0          (model edge tested)
    mappedVisibleDrawingSpace=2   (mapped drawing entity tested, per location)

The 2025 Help calls the return "modeling edges". On this build it compares
equal to the drawing entity IView.GetCorrespondingEntity returns, not to the
model edge it came from. MapVisibleDatumEntity tested the model entity from
r33 until r42, so it could never match, and every vertical datum failed
closed with PolylineVisibilityUnavailable regardless of real visibility.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module13_ProjectionResolution.bas"


class R23DatumVisibilitySpaceContracts(unittest.TestCase):
    def code(self) -> str:
        text = (SOURCE / MODULE).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def datum_body(self) -> str:
        return self.code().split("Public Function MapVisibleDatumEntity(")[
            1
        ].split("\nEnd Function")[0]

    def test_mapping_happens_before_the_visibility_test(self):
        body = self.datum_body()
        self.assertLess(
            body.index("MapModelEntityToDrawing("),
            body.index("FindVisibleModelEdgeIndex("),
        )

    def test_visibility_is_proved_on_the_mapped_drawing_entity(self):
        body = self.datum_body()
        primary = body.split("If Not mapped Is Nothing Then")[1].split(
            "Else"
        )[0]
        self.assertIn(
            "FindVisibleModelEdgeIndex( _\n            swApp, swView, mapped,",
            primary,
        )
        self.assertIn("space=Drawing", primary)

    def test_fail_closed_guard_is_retained(self):
        """A mapped entity absent from a NON-EMPTY visible array must still
        refuse. Only the documented empty-array case may fall back."""
        body = self.datum_body()
        self.assertIn("DatumMap:MappedEntityNotVisible", body)
        self.assertIn("DatumMap:UnmappedWithVisibleArray", body)

        # Every refusal is gated on the array not being empty.
        for marker in ("DatumMap:MappedEntityNotVisible",
                       "DatumMap:UnmappedWithVisibleArray"):
            with self.subTest(marker=marker):
                preceding = body.split(marker)[0]
                self.assertIn('"status=NoEdges"', preceding)

    def test_scoped_selection_remains_the_empty_array_route_only(self):
        body = self.datum_body()
        self.assertEqual(body.count("SelectModelEntityInView("), 1)
        before_selection = body.split("SelectModelEntityInView(")[0]
        self.assertIn('"status=NoEdges"', before_selection)
        self.assertIn("visibility=ScopedViewSelection", body)

    def test_the_measurement_is_recorded_where_it_belongs(self):
        """The two-way control was an instrument for one question, not a
        permanent metric: it cost a GetPolylines7 read per mapped edge. Once
        the question was settled the instrument was removed and the finding
        moved to the API-contract record, which is what that document is
        for."""
        contract = (WORKSPACE / "docs" / "SOLIDWORKS_API_VALIDATION.md")
        text = contract.read_text(encoding="utf-8")
        section = text.split("## `IView.GetPolylines7`")[1].split("\n## ")[0]
        self.assertIn("mappedVisibleEdges=0", section)
        self.assertIn("mappedVisibleDrawingSpace=2", section)
        self.assertIn("MapVisibleDatumEntity", section)


if __name__ == "__main__":
    unittest.main()
