"""R23-823. Clamp the section annotations AFTER the view is placed.

r61 (macro_qa/20260805_070039_P-0251-14A-001) closed one defect and
disproved the diagnosis of the other.

The creation-time clamp worked exactly as designed: RD2 was recorded
`clamped=True` at `textY=0.275000`, which is `UsableTop` to the micron.
`ANNOTATION_EXTENTS` failed anyway, on RD2 at `y=0.289235`.

The mover was not auto-arrange. Section View J-J was never arranged -
`ACTIVATE_VIEW|operation=Dimension arrange` names Drawing View1, Drawing
View2 and Drawing View4 only. It was the structural grid:

    creation-time outline  0.289060,0.039385,0.318940,0.252265
    LAYOUT_MOVE readback   0.191752,0.053620,0.221632,0.266500
    delta                 -0.097308,          +0.014235

    RD2 requested (0.304000, 0.275000) + delta = (0.206692, 0.289235)

which is the violation line to six decimals. RD1 lands at 0.278500, under
ContentBorderTop 0.287, so only RD2 trips. A box measured before the view
is placed is stale by construction, so the correction has to run after
placement - after the grid, after auto-arrange, before the final rebuild.

What the clamp may not become: it moves annotation ORIGINS. It does not
move a view and does not rescale one. The content-envelope repositioning
and rescaling the 2026-08-04 user decision retired stay uncalled.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module10_SectionDimensionEngine.bas"
PIPELINE = "Module2_DrawingPipeline.bas"

CLAMP = "Public Function ClampSectionAnnotationsIntoUsableArea("


class R23SectionAnnotationClampContracts(unittest.TestCase):
    def code(self, name: str) -> str:
        text = (SOURCE / name).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def clamp(self) -> str:
        return self.code(MODULE).split(CLAMP)[1].split("\nEnd Function")[0]

    def pipeline(self) -> str:
        return self.code(PIPELINE).split(
            "Public Sub RunDrawingPipeline("
        )[1].split("\nEnd Sub")[0]

    # --- authorization and fail-closed ------------------------------------

    def test_the_clamp_refuses_without_explicit_authorization(self):
        body = self.clamp()
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertIn("SECTION_ANNOTATION_CLAMP_REFUSED", body)
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("GetAnnotations"),
        )

    def test_unproved_bounds_refuse_rather_than_guess(self):
        """Clamping to a box nothing measured would move the text
        somewhere arbitrary, which is worse than leaving it alone."""
        body = self.clamp()
        self.assertIn(
            "If Not evidence.LayoutBoundariesProven Then", body)
        self.assertIn("reason=LayoutBoundariesUnproved", body)
        self.assertLess(
            body.index("evidence.LayoutBoundariesProven"),
            body.index("GetAnnotations"),
        )

    # --- what it is allowed to touch --------------------------------------

    def test_the_clamp_moves_annotations_and_nothing_else(self):
        """R23-819 boundary. Moving or rescaling a view here would restore
        the behaviour the 2026-08-04 user decision retired."""
        body = self.clamp()
        for forbidden in ("SetPosition(", "Position2 =", "UseSheetScale",
                          "ScaleRatio", "SetScale", "AddDimension2",
                          "SetValues2", "SetFitValues", "DeleteSelection",
                          "ArrangeViews", "RepositionView"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)
        self.assertIn("annotation.SetPosition2(targetX, targetY, "
                      "currentZ)", body)

    def test_the_pipeline_scopes_the_clamp_to_the_one_section_view(self):
        body = self.code(PIPELINE).split(
            "Private Sub ClampR23SectionAnnotations("
        )[1].split("\nEnd Sub")[0]
        self.assertIn('<> _\n        "P-0251-14A-001" Then Exit Sub', body)
        self.assertIn("If sectionViews.Count <> 1 Then Exit Sub", body)

    # --- ordering, which is the whole point -------------------------------

    def test_the_clamp_runs_after_everything_that_moves_annotations(self):
        """The r61 defect is ordering, not arithmetic. The clamp must
        follow the structural grid AND auto-arrange, or it goes stale
        exactly the way the creation-time clamp did."""
        body = self.pipeline()
        call = body.index("ClampR23SectionAnnotations swDraw, evidence")
        for earlier in ("ArrangeViewsInMeasuredGrid(",
                        "AutoArrangeAllDrawingDimensions",
                        "PopulateTitleBlock"):
            with self.subTest(earlier=earlier):
                self.assertLess(body.index(earlier), call)

    def test_the_clamp_runs_before_the_rebuild_that_verifies_it(self):
        body = self.pipeline()
        call = body.index("ClampR23SectionAnnotations swDraw, evidence")
        self.assertLess(call, body.index('"R23 final content"'))
        self.assertLess(call, body.index("EvaluateSemanticDrawing"))

    # --- the readback is the verdict --------------------------------------

    def test_the_setter_result_is_not_trusted_on_its_own(self):
        """The Help says a constrained annotation is placed "as near as
        possible", and that radial and diametric dimensions cannot be
        positioned this way at all. Both look like a call that returned."""
        body = self.clamp()
        self.assertIn("readback = annotation.GetPosition", body)
        self.assertIn("|setterResult=", body)
        self.assertIn("|readbackX=", body)
        self.assertIn("|readbackY=", body)
        self.assertIn("|nowInside=", body)
        # The count that matters is decided by the readback, not the call.
        self.assertLess(
            body.index("readback = annotation.GetPosition"),
            body.index("nowInside = _"),
        )
        self.assertLess(
            body.index("nowInside = _"),
            body.index("moved = moved + 1"),
        )

    def test_an_annotation_that_stayed_outside_is_counted_as_such(self):
        body = self.clamp()
        self.assertIn(
            "If nowInside Then\n"
            "            moved = moved + 1\n"
            "        Else\n"
            "            stillOutside = stillOutside + 1",
            body,
        )
        self.assertIn("|stillOutside=", body)

    def test_an_annotation_already_inside_is_left_alone(self):
        """Setting a position that is already correct still counts as a
        mutation, and a needless one hides the ones that mattered."""
        body = self.clamp()
        self.assertIn(
            "If targetX = currentX And targetY = currentY Then", body)
        self.assertIn("action=AlreadyInside", body)
        self.assertLess(
            body.index("action=AlreadyInside"),
            body.index("annotation.SetPosition2("),
        )

    def test_the_clamp_is_bounded_by_the_proved_usable_box(self):
        body = self.clamp()
        for bound in ("evidence.UsableLeft", "evidence.UsableRight",
                      "evidence.UsableBottom", "evidence.UsableTop"):
            with self.subTest(bound=bound):
                self.assertIn(bound, body)
        self.assertIn(
            "targetX = ClampToRange( _\n"
            "            currentX, evidence.UsableLeft, "
            "evidence.UsableRight)",
            body,
        )
        self.assertIn(
            "targetY = ClampToRange( _\n"
            "            currentY, evidence.UsableBottom, "
            "evidence.UsableTop)",
            body,
        )

    def test_the_z_coordinate_is_preserved_not_zeroed(self):
        """SetPosition2 takes three coordinates. Passing 0 for z would
        move the annotation in a direction nothing asked about."""
        body = self.clamp()
        self.assertIn("currentZ = CDbl(position(lower + 2))", body)
        self.assertIn("targetY, currentZ)", body)

    def test_the_position_array_is_checked_before_it_is_indexed(self):
        """GetPosition returns an empty SafeArray when it fails."""
        body = self.clamp()
        self.assertIn("If Not IsArray(position) Then", body)
        self.assertIn("If UBound(position) - lower + 1 < 3 Then", body)


if __name__ == "__main__":
    unittest.main()
