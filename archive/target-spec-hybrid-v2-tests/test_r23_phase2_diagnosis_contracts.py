"""R23 Phase 2 contracts: make the remaining failures legible.

The 23:24 r39 production run failed with two blind reports. The section
stage printed "R23 semantic section path was not resolved." with no reason,
and dimension arrange printed "Dimension arrange API error in 'Drawing
View1': 0: " with no error number. Neither could be diagnosed from its own
evidence. These tests hold the fixes for both, plus the honest reporting of
an edge that declined every mapping route.

The r40/r41 visibility classification that once lived here was an instrument
for a single question - what IView.GetPolylines7 actually returns - and r42
settled it (see test_r23_datum_visibility_space_contracts). Its model-space
variant matched nothing and measured nothing, so it was removed rather than
left in evidence to mislead.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"

# Error handlers that build their message by calling a name helper. Every
# one of those helpers contains On Error Resume Next, which resets the
# global Err object; VBA evaluates a concatenation left to right, so
# Err.Number read after the call reports 0 for a real raise.
HANDLERS_THAT_MUST_CAPTURE_ERR_FIRST = {
    "Module4_ModelItemImporter.bas": "arrangeErrorNumber",
    "Module5_FallbackDimensionEngine.bas": "groupErrorNumber",
    "Module10_SectionDimensionEngine.bas": "inventoryErrorNumber",
    "Module13_ProjectionResolution.bas": "passErrorNumber",
    "Module14_AnnotationImport.bas": "annotationErrorNumber",
}


class R23ErrorCaptureContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def test_handlers_capture_err_before_calling_a_name_helper(self):
        for module, local in HANDLERS_THAT_MUST_CAPTURE_ERR_FIRST.items():
            with self.subTest(module=module):
                body = self.code(module)
                self.assertIn(f"{local} = Err.Number", body)
                self.assertIn(f"CStr({local})", body)

    def test_captured_number_is_read_before_the_helper_runs(self):
        """Order is the whole defect: capture must precede the call."""
        checks = (
            ("Module4_ModelItemImporter.bas", "arrangeErrorNumber = Err.Number",
             'AddFailure "Dimension arrange API error'),
            ("Module5_FallbackDimensionEngine.bas",
             "groupErrorNumber = Err.Number",
             'AddFailure "Ordinate group error'),
            ("Module13_ProjectionResolution.bas",
             "passErrorNumber = Err.Number", "PROJECTION_PASS_ERROR"),
        )
        for module, capture, emit in checks:
            with self.subTest(module=module):
                body = self.code(module)
                self.assertLess(body.index(capture), body.index(emit))

    def test_no_handler_still_reads_err_number_after_a_name_helper(self):
        """Regression guard for the exact shape that produced ': 0: '."""
        import re

        pattern = re.compile(
            r"(GetViewName|SafeViewName|SectionToken)\([^)]*\)"
            r"(?:[^\n]*\n){0,2}[^\n]*CStr\(Err\.Number\)"
        )
        for module in HANDLERS_THAT_MUST_CAPTURE_ERR_FIRST:
            with self.subTest(module=module):
                self.assertIsNone(
                    pattern.search(self.code(module)),
                    msg=f"{module} still reads Err.Number after a name helper",
                )


class R23SectionDiagnosisContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def section_body(self) -> str:
        return self.code("Module2_DrawingPipeline.bas").split(
            "Private Function CreateSemanticPrimarySection("
        )[1].split("\nEnd Function")[0]

    def test_every_candidate_view_reports_its_own_rejection(self):
        body = self.section_body()
        self.assertIn("R23_SECTION_PATH_CANDIDATE|", body)
        self.assertIn("path.Summary()", body)
        self.assertIn("path.RejectionReason", body)

    def test_failure_text_carries_the_reasons(self):
        """The bare sentence alone is what made two runs undiagnosable."""
        body = self.section_body()
        self.assertIn(
            'AddFailure "R23 semantic section path was not resolved: "',
            body.replace("_\n        ", "").replace("\n", ""),
        )
        self.assertIn("rejectionSummary", body)
        self.assertIn("viewsWithoutProjections", body)

    def test_a_resolved_path_is_not_overwritten_by_a_later_view(self):
        """The old loop reassigned `path` on every iteration."""
        body = self.section_body()
        self.assertIn("Set resolvedPath = path", body)
        self.assertIn("If resolvedPath Is Nothing Then", body)
        self.assertIn("Set path = resolvedPath", body)

    def test_views_without_projections_are_summarised_not_listed(self):
        body = self.section_body()
        self.assertIn('"NoProjectionsInView"', body)
        self.assertIn("viewsWithoutProjections = viewsWithoutProjections + 1",
                      body)

    def test_post_layout_section_line_filter_excludes_an_empty_read(self):
        """A view with no cut answers sectionLine=Read|values=0, not
        NoGeometryReturned as first assumed; all three views logged it."""
        helper = self.code("Module2_DrawingPipeline.bas").split(
            "Private Sub RecordSectionLineAfterLayout("
        )[1].split("\nEnd Sub")[0]
        self.assertIn('"|values=0"', helper)
        self.assertIn("R23_SECTION_LINE_POSTLAYOUT", helper)


class R23ProjectionVisibilityDiagnosisContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def resolve_body(self) -> str:
        return self.code("Module13_ProjectionResolution.bas").split(
            "Private Sub ResolveProjection("
        )[1].split("\nEnd Sub")[0]

    def test_edges_that_decline_every_route_are_counted_honestly(self):
        """The r40/r41 "obscured" label tested the MODEL edge against the
        GetPolylines7 array. r42's control proved that comparison matches
        nothing, so the label measured nothing and was removed rather than
        left to mislead. Drawing-space classification is impossible at this
        point by construction: it needs a drawing entity, and having none is
        exactly the condition being reported."""
        body = self.resolve_body()
        self.assertIn("stageUnmappedAllRoutes", body)
        self.assertIn("|unmappedAllRoutes=", body)
        self.assertIn("|firstUnmappedRoute=", body)

    def test_void_visibility_metrics_are_not_reintroduced(self):
        body = self.resolve_body()
        for retired in ("unmappedObscuredEdges", "unmappedVisibleEdges",
                        "mappedVisibleDrawingSpace", "stageMappedVisible"):
            with self.subTest(metric=retired):
                self.assertNotIn(retired, body)

    def test_unmapped_counter_runs_only_after_every_route_declined(self):
        body = self.resolve_body()
        branch = body.split("If mapped Is Nothing Then")[-1].split(
            "GoTo ContinueEdge")[0]
        self.assertIn("stageUnmappedAllRoutes = stageUnmappedAllRoutes + 1",
                      branch)
        # Route D must have been attempted before the count is taken.
        self.assertLess(
            body.index("SelectModelEntityInView("),
            body.index("stageUnmappedAllRoutes = stageUnmappedAllRoutes + 1"),
        )


if __name__ == "__main__":
    unittest.main()
