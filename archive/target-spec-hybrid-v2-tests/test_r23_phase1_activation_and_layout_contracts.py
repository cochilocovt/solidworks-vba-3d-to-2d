"""R23 Phase 1 contracts: proved view activation and a two-pass layout.

Both defects were found in the 2026-08-04 18:45 production run.

1. IDrawingDoc.ActivateView returns False on this build even when the view
   does become active. Every mutating transaction that branched on that raw
   result refused itself; annotation import inserted zero annotations into a
   drawing whose views were activating correctly.
2. The only layout pass ran at pipeline step 3, before the section and
   isometric views existed, so the P-0251 reference-zone validator failed
   unconditionally and the drawing was never laid out at all.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"

# Modules that activate a view in order to mutate the drawing. Each must
# prove activation by readback rather than trusting the setter's return.
MUTATING_ACTIVATION_MODULES = (
    "Module14_AnnotationImport.bas",
    "Module15_OrdinateScheme.bas",
    "Module16_CalloutDefinition.bas",
    "Module17_SectionPath.bas",
)


class R23Phase1ActivationContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def test_mutating_transactions_activate_through_the_proving_helper(self):
        for module in MUTATING_ACTIVATION_MODULES:
            with self.subTest(module=module):
                body = self.code(module)
                self.assertIn(
                    "Module8_RuntimeSupport.ActivateDrawingView(", body)
                self.assertNotIn("swDrawing.ActivateView(", body)

    def test_activation_helper_accepts_a_false_setter_with_a_matching_readback(
        self,
    ):
        body = self.code("Module8_RuntimeSupport.bas")
        helper = body.split("Public Function ActivateDrawingView(")[1].split(
            "\nEnd Function"
        )[0]

        # The verdict comes from ActiveDrawingView, not from the setter.
        self.assertIn("swDraw.ActiveDrawingView", helper)
        self.assertIn("activeViewMatches", helper)
        self.assertIn("setterResult=False|readbackMatched=True", helper)
        self.assertNotIn("If Not activateResult Then Exit Function", helper)

    def test_projection_resolution_records_activation_without_branching(self):
        """Module13 is the one legitimate raw read: it reports, never gates."""
        body = self.code("Module13_ProjectionResolution.bas")
        self.assertIn("PROJECTION_VIEW_ACTIVATION|view=", body)
        self.assertNotIn("If Not viewActivated Then Exit", body)

    def test_import_refuses_only_on_a_real_readback_mismatch(self):
        body = self.code("Module14_AnnotationImport.bas")
        loop = body.split("Public Function ImportModelAnnotations(")[1].split(
            "\nEnd Function"
        )[0]

        self.assertIn("IMPORT_VIEW_NOT_ACTIVATED", loop)
        self.assertIn("Module8_RuntimeSupport.ActivateDrawingView(", loop)
        # Activation must still gate the insert; the fix changed the evidence
        # the gate uses, not whether one exists.
        self.assertLess(
            loop.index("ActivateDrawingView("),
            loop.index("InsertModelAnnotations4("),
        )

    def test_view_object_is_null_checked_before_the_helper_dereferences_it(
        self,
    ):
        checks = {
            "Module15_OrdinateScheme.bas": "scheme.DrawingView Is Nothing",
            "Module16_CalloutDefinition.bas":
                "projection.DrawingView Is Nothing",
            "Module17_SectionPath.bas": "path.SourceView Is Nothing",
        }
        for module, guard in checks.items():
            with self.subTest(module=module):
                self.assertIn(guard, self.code(module))


class R23Phase1LayoutContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def run_body(self) -> str:
        return self.read("Module2_DrawingPipeline.bas").split(
            "Public Sub RunDrawingPipeline("
        )[1].split("\nEnd Sub")[0]

    def test_layout_entry_point_distinguishes_the_two_passes(self):
        body = self.code("Module9_LayoutEngine.bas")
        signature = body.split("Public Function ArrangeViewsInMeasuredGrid(")[
            1
        ].split(") As Boolean")[0]
        self.assertIn("isFinalStructuralPass As Boolean", signature)

    def test_fixture_reference_zones_run_only_on_the_final_pass(self):
        """The reference validator asserts a view set step 3 cannot have."""
        body = self.code("Module9_LayoutEngine.bas")
        zone_aware = body.split("Private Function ArrangeZoneAwareViews(")[
            1
        ].split("\nEnd Function")[0]

        self.assertEqual(zone_aware.count("ArrangeP0251ReferenceZones("), 1)

        # The dispatch must be reached only through a final-pass condition.
        guard = zone_aware.split("ArrangeP0251ReferenceZones(")[0].rsplit(
            "\n    If ", 1
        )[-1]
        self.assertIn("isFinalStructuralPass And", guard)
        self.assertIn('GetFixtureKey(evidence.PartPath) = "P-0251-14A-001"',
                      guard)

    def test_rough_pass_never_writes_the_layout_verdict(self):
        """MarkStageFailed is permanent, so a rough failure must not stick."""
        body = self.code("Module9_LayoutEngine.bas")
        entry = body.split("Public Function ArrangeViewsInMeasuredGrid(")[
            1
        ].split("\nEnd Function")[0]

        self.assertIn('evidence.RequireStage "LAYOUT"', entry)
        self.assertIn("LAYOUT_ROUGH|pass=", entry)
        self.assertIn("verdictDeferredTo=FinalStructuralPass", entry)

        # Every outcome verdict must sit behind the rough-pass early exit.
        # The two preconditions that fail either pass - no sheet, no views -
        # are deliberately excluded: they are fatal regardless of pass.
        precondition_details = (
            '"current sheet is Nothing"',
            '"invalid sheet size"',
            '"no model views"',
            '"controlled sheet boundaries are not proved"',
            '"API error "',
        )
        for marker in ('MarkStageProved "LAYOUT"', 'MarkStageFailed "LAYOUT"'):
            segments = entry.split(marker)
            for index in range(len(segments) - 1):
                detail = segments[index + 1].split("\n\n")[0]
                if any(text in detail for text in precondition_details):
                    continue
                prefix = marker.join(segments[: index + 1])
                self.assertIn(
                    "If Not isFinalStructuralPass Then", prefix,
                    msg="verdict write is not behind the rough-pass exit",
                )

    def test_pipeline_runs_a_rough_pass_and_a_final_structural_pass(self):
        body = self.run_body()
        self.assertIn('"R23Rough", False, evidence', body)
        self.assertIn('"R23Structural", True, evidence', body)
        self.assertEqual(body.count("ArrangeViewsInMeasuredGrid("), 2)

    def test_final_structural_pass_sees_every_required_view(self):
        """It must follow the isometric and precede title and QA content."""
        body = self.run_body()
        order = (
            "CreateSemanticPrimarySection(",
            "CreateIsometricView(",
            '"R23Structural"',
            "PopulateTitleBlock",
            "EvaluateSemanticDrawing",
        )
        positions = [body.index(token) for token in order]
        self.assertEqual(positions, sorted(positions))

    def test_section_line_geometry_is_reread_after_the_views_move(self):
        """A moved source view re-lays its section line on the sheet."""
        body = self.run_body()
        self.assertIn("RecordSectionLineAfterLayout swDraw, evidence", body)
        self.assertLess(
            body.index('"R23Structural"'),
            body.index("RecordSectionLineAfterLayout"),
        )

        helper = self.code("Module2_DrawingPipeline.bas").split(
            "Private Sub RecordSectionLineAfterLayout("
        )[1].split("\nEnd Sub")[0]
        self.assertIn("Module17_SectionPath.ReadSectionLineInfo(", helper)
        self.assertIn("R23_SECTION_LINE_POSTLAYOUT", helper)
        # Read-only: it reports geometry, it does not re-cut or reposition.
        for mutation in ("CreateLine", "CreateSectionViewAt5", ".Position"):
            with self.subTest(mutation=mutation):
                self.assertNotIn(mutation, helper)

    def test_final_layout_waiver_is_untouched_by_the_structural_pass(self):
        """The retired content-envelope repositioning stays uncalled."""
        body = self.run_body()
        self.assertIn("RecordR23UserAcceptedLayout(evidence)", body)
        self.assertNotIn("ApplyScaleToFit", body)
        self.assertNotIn("ApplyPlacementPlan", body)


if __name__ == "__main__":
    unittest.main()
