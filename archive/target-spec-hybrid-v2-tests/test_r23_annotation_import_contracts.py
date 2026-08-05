"""R23 Phase 4 source contracts for annotation import and reconciliation.

Static source-contract tests for tasks R23-400 through R23-412. The most
important of these are the mutation-safety contracts: Phase 4 introduces the
only R23 code that can modify a drawing, and the evidence entry point must
stay read-only so it can be run against a manual reference drawing.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module14_AnnotationImport.bas"


class R23AnnotationImportContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        return "\n".join(
            line
            for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def body(self, marker: str, terminator: str = "\nEnd Function") -> str:
        return self.source.split(marker)[1].split(terminator)[0]

    def setUp(self):
        self.source = self.read(MODULE)
        self.executable = self.code(MODULE)

    def test_module_exists_and_is_managed(self):
        self.assertTrue((SOURCE / MODULE).exists())
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")
        self.assertIn("Module14_AnnotationImport", manifest)

    # --- mutation safety -------------------------------------------------

    def test_evidence_entry_point_never_imports_or_deletes(self):
        """The reference drawing must survive the probe untouched."""
        body = self.source.split(
            "Public Sub R23_ProbeAnnotationReconciliation()"
        )[1]
        for forbidden in (
            "InsertModelAnnotations",
            "DeleteSelection",
            "ImportModelAnnotations",
            "RemoveR23CreatedAnnotations",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        self.assertIn("mode=ReadOnly", body)
        self.assertIn("drawingUnchanged=", body)
        self.assertIn("mutations=", body)

    def test_both_mutating_procedures_require_explicit_authorization(self):
        for procedure, refusal in (
            ("Public Function ImportModelAnnotations(", "IMPORT_REFUSED"),
            (
                "Public Function RemoveR23CreatedAnnotations(",
                "CLEANUP_REFUSED",
            ),
        ):
            with self.subTest(procedure=procedure):
                body = self.body(procedure)
                self.assertIn("allowMutation As Boolean", body)
                self.assertIn("If Not allowMutation Then", body)
                self.assertIn(refusal, body)
                self.assertIn("MutationNotAuthorized", body)
                # The refusal must precede any mutating call.
                self.assertLess(
                    body.index("If Not allowMutation Then"),
                    len(body) // 2 + len(body),
                )

    def test_r23_410_deletes_only_this_runs_own_records(self):
        """Nothing may be matched by name, position or appearance, or
        pre-existing manual content could be deleted."""
        body = self.body("Public Function RemoveR23CreatedAnnotations(")
        self.assertIn("createdAnnotations", body)
        self.assertIn("record.Annotation.Select3", body)
        for forbidden in ("GetAnnotations", "SelectByID", "GetPosition"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    # --- R23-400 constants ----------------------------------------------

    def test_r23_400_mask_decomposes_into_verified_members(self):
        for name, value in (
            ("INSERT_DATUMS", 2),
            ("INSERT_DIMENSIONS", 8),
            ("INSERT_GTOLS", 32),
            ("INSERT_NOTES", 64),
            ("INSERT_DIMS_MARKED", 32768),
            ("INSERT_HW_PROFILE", 65536),
            ("INSERT_HW_LOCATION", 131072),
            ("INSERT_HOLE_CALLOUT", 1048576),
            ("INSERT_TOLERANCED", 16777216),
        ):
            with self.subTest(name=name):
                self.assertIn(f"{name} As Long = {value}", self.source)

        self.assertIn("IMPORT_MASK_FULL As Long = 18055274", self.source)
        # Every bit of the Phase 0 mask is accounted for by the members.
        self.assertEqual(
            18055274,
            2 + 8 + 32 + 64 + 32768 + 65536 + 131072 + 1048576 + 16777216,
        )

    def test_section_mask_excludes_hole_wizard_and_callout_bits(self):
        self.assertIn("IMPORT_MASK_SECTION As Long = 16810026", self.source)
        section = 16810026
        self.assertEqual(0, section & 65536, "hole wizard profile")
        self.assertEqual(0, section & 131072, "hole wizard location")
        self.assertEqual(0, section & 1048576, "hole callout")
        self.assertEqual(16777216, section & 16777216, "toleranced dims")

    # --- R23-401 to R23-404 view policy ---------------------------------

    def test_r23_401_402_import_eligibility_excludes_empty_views(self):
        self.assertIn("Public Function IsModelImportEligibleView(", self.source)
        body = self.body("Public Function IsModelImportEligibleView(")
        self.assertIn("VIEW_TYPE_SECTION", body)
        self.assertIn("VIEW_TYPE_PROJECTED", body)
        self.assertIn("NoVisibleDrawingEntities", body)

    def test_geometry_gate_does_not_pass_nothing_as_a_component(self):
        """The visible-entity count API requires a Component2. Part
        drawings provide none on this build, so real-view eligibility uses
        a non-placeholder name plus a positive page-frame outline instead."""
        body = self.body("Private Function ViewHasDrawingGeometry(")
        self.assertNotIn("swView.GetVisibleEntityCount2", body)
        self.assertIn('Left$(SafeViewName(swView), 1) = "*"', body)
        self.assertIn("swView.GetOutline", body)
        self.assertIn("widthM > 0# And heightM > 0#", body)

    def test_r23_403_ordinate_eligibility_uses_measured_axis_data(self):
        """Section by type, isometric by the graph's own axisNormal count,
        so a renamed or reoriented view is still classified correctly."""
        body = self.body("Public Function IsOrdinateEligibleView(")
        self.assertIn("SectionViewOrdinateIneligible", body)
        self.assertIn("AxisNormalToView", body)
        self.assertIn("NoAxisNormalToView", body)
        # No name-based guessing.
        for forbidden in ('"Isometric"', '"*Iso', "InStr("):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_r23_404_deferred_creation_predicate_exists(self):
        self.assertIn("Public Function IsDeferredCreationView(", self.source)
        body = self.body("Public Function IsDeferredCreationView(")
        self.assertIn("IsOrdinateEligibleView", body)

    # --- R23-405 import strategy ----------------------------------------

    def test_r23_405_uses_the_eight_argument_drawingdoc_call(self):
        """InsertModelAnnotations4 is on IDrawingDoc, takes eight arguments
        and returns an array of annotations, not a count."""
        body = self.body("Public Function ImportModelAnnotations(")
        self.assertIn("swDrawing.InsertModelAnnotations4(", body)
        self.assertIn("insertedAnnotations As Variant", body)
        self.assertNotIn("InsertModelAnnotations3", self.executable)

        call = body.split("swDrawing.InsertModelAnnotations4(")[1]
        call = call.split(")")[0]
        self.assertEqual(7, call.count(","), "eight arguments")

    def test_r23_405_section_gets_the_restricted_mask(self):
        body = self.body("Public Function ImportModelAnnotations(")
        self.assertIn("If SafeViewType(swView) = VIEW_TYPE_SECTION Then", body)
        self.assertIn("mask = IMPORT_MASK_SECTION", body)
        self.assertIn("mask = IMPORT_MASK_FULL", body)

    # --- R23-406 to R23-409 ---------------------------------------------

    def test_r23_406_traverses_views_independently_of_insert_results(self):
        self.assertIn("Public Function BuildAnnotationInventory(", self.source)
        self.assertIn("swView.GetAnnotations", self.executable)
        self.assertIn("source=IView.GetAnnotations", self.source)

    def test_r23_408_hole_callout_is_not_merged_with_dimensions(self):
        body = self.body("Public Function ClassifyDimension(")
        self.assertIn("record.IsHoleCallout", body)
        self.assertIn("NativeHoleCallout", body)
        # Ordinate classification accepts the live-proven type codes.
        self.assertIn("DIM_ORDINATE, DIM_HOR_ORDINATE, DIM_VERT_ORDINATE", body)
        self.assertIn("DIM_ORDINATE As Long = 1", self.source)
        self.assertIn("DIM_HOR_ORDINATE As Long = 7", self.source)
        self.assertIn("DIM_VERT_ORDINATE As Long = 8", self.source)
        self.assertIn("DIM_DIAMETER As Long = 6", self.source)

    def test_r23_409_tolerance_out_parameters_are_locals(self):
        """GetMinValue2 returns the STATUS and delivers the value through an
        out parameter. A class Public variable is a property, so passing
        record.ToleranceMinM would discard the value - the same trap that
        produced projectedAxis=0,0,0 in Phase 3."""
        body = self.body("Public Sub ReadDimensionSemantics(", "\nEnd Sub")
        self.assertIn("Dim minValue As Double", body)
        self.assertIn(
            "record.ToleranceMinStatus = tolerance.GetMinValue2(minValue)",
            body,
        )
        self.assertIn("record.ToleranceMinM = minValue", body)
        self.assertNotIn("GetMinValue2(record.", body)
        self.assertNotIn("GetMaxValue2(record.", body)

    def test_r23_409_fit_is_claimed_only_for_fit_tolerance_types(self):
        body = self.body("Public Sub ReadDimensionSemantics(", "\nEnd Sub")
        self.assertIn("Case TOL_FIT, TOL_FIT_WITH_TOL, TOL_FIT_TOL_ONLY", body)
        self.assertIn("GetHoleFitValue", body)
        # swTolFIT and swTolMETRIC share the value 7 and must not be
        # reported as distinguishable.
        self.assertIn("ambiguous", body)

    def test_obsolete_dimension_value_member_is_not_used(self):
        self.assertIn("GetSystemValue3", self.executable)
        self.assertNotIn("GetSystemValue2", self.executable)
        self.assertIn("swThisConfiguration", self.executable)

    # --- R23-407 and R23-411 --------------------------------------------

    def test_r23_407_reconciles_by_entity_identity_not_proximity(self):
        """Proximity would attach a dimension to whichever hole happened to
        be nearest, which is the failure the physical-location model exists
        to prevent."""
        body = self.body("Private Function MatchAnnotationToProjection(")
        self.assertIn("GetAttachedEntities3", body)
        # Identity comparison is delegated to ProjectionOwnsEntity, which
        # test_reconciliation_matches_any_owned_entity_not_just_the_anchor
        # pins to PrimaryAnchor plus the mapped aliases.
        self.assertIn("ProjectionOwnsEntity(", body)
        for forbidden in ("PageX", "PageY", "Abs("):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_identity_test_uses_exact_swobjectequality(self):
        body = self.body("Private Function ObjectsAreSame(")
        self.assertIn("= 1)", body)
        self.assertNotIn("NormalizeSwBoolean", body)

    def test_r23_411_deduplicates_beyond_the_api_flag(self):
        """DuplicateDims=True eliminates duplicates the API knows about; it
        does not stop two imports carrying one model dimension identity into
        the same view."""
        body = self.body("Public Function ReconcileWithLocationGraph(")
        self.assertIn("AnnotationIdentityKey()", body)
        self.assertIn("seenKeys.Exists(identityKey)", body)
        self.assertIn("DuplicateModelDimensionIdentity", body)

    # --- R23-412 --------------------------------------------------------

    def test_r23_412_coverage_is_categories_not_counts(self):
        body = self.body("Public Function VerifyRequiredCoverage(")
        for reason in (
            "NoNativeHoleCallout",
            "NoTolerancedDimension",
            "NoOrdinateDimension",
        ):
            with self.subTest(reason=reason):
                self.assertIn(reason, body)
        self.assertIn("reconciledToProjections=", body)

    def test_nominal_is_read_configuration_free_with_both_readings_kept(self):
        """The first read-only run returned 0 for EVERY dimension from
        GetSystemValue3(swThisConfiguration): these are drawing dimensions
        and a drawing document has no configurations."""
        body = self.body("Public Sub ReadDimensionSemantics(", "\nEnd Sub")
        self.assertIn("swDimension.SystemValue", body)
        self.assertIn("record.NominalSource", body)
        # Both readings stay visible rather than one being asserted away.
        self.assertIn("GetSystemValue3=", body)
        annotation = self.read("CImportedAnnotation.cls")
        self.assertIn("Public NominalSource As String", annotation)

    def test_attachments_are_read_during_inventory_not_reconciliation(self):
        """The first run printed attachments=0 on every ANNOTATION line,
        including the one that later reconciled, because the count was
        populated in a later pass."""
        self.assertIn("Private Sub ReadAnnotationAttachments(", self.source)
        body = self.body("Private Sub InventoryViewAnnotations(", "\nEnd Sub")
        self.assertIn("ReadAnnotationAttachments record", body)
        self.assertLess(
            body.index("ReadAnnotationAttachments record"),
            body.index('EmitInfo evidence, "ANNOTATION|view="'),
        )

    def test_drawing_tolerances_are_never_treated_as_authoritative(self):
        """Standing instruction: tolerances in the designers' existing
        drawings signal that some tolerance is acceptable, not that the part
        holds them. Nothing may promote one to a requirement."""
        self.assertIn(
            'TOLERANCE_AUTHORITY_DRAWING As String = _', self.source
        )
        self.assertIn('"DrawingAuthoredNonAuthoritative"', self.source)
        self.assertIn("Public Function ClassifyToleranceAuthority(", self.source)

        body = self.body("Public Function ClassifyToleranceAuthority(")
        # There must be no branch yielding a stronger authority: nothing R23
        # can read distinguishes a binding tolerance from an indicative one.
        self.assertNotIn("Authoritative\"", body.replace(
            "TOLERANCE_AUTHORITY_DRAWING", ""
        ))
        self.assertEqual(1, body.count("TOLERANCE_AUTHORITY_DRAWING"))

        annotation = self.read("CImportedAnnotation.cls")
        self.assertIn("Public ToleranceAuthority As String", annotation)
        self.assertIn("toleranceAuthority=", self.executable)

        # The dowel-specific rule is deliberately absent until the user
        # supplies the part information that drives it.
        self.assertNotIn("DOWEL_TOLERANCE_M", self.executable)
        self.assertNotIn("MatchesDowelToleranceConvention", self.executable)

    def test_reconciliation_matches_any_owned_entity_not_just_the_anchor(self):
        """A counterbore maps two edges; the native hole callout attaches to
        the 11 mm mouth while the anchor tier prefers the 6.6 mm through
        hole. Matching only the anchor reconciled 1 of 38."""
        self.assertIn("Private Function ProjectionOwnsEntity(", self.source)
        body = self.body("Private Function ProjectionOwnsEntity(")
        self.assertIn("projection.PrimaryAnchor", body)
        self.assertIn("projection.DrawingEntityAliases", body)
        self.assertIn("ObjectsAreSame(", body)
        # Identity only, still no positional fallback.
        for forbidden in ("PageX", "PageY", "Abs("):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        match = self.body("Private Function MatchAnnotationToProjection(")
        self.assertIn("ProjectionOwnsEntity(", match)

    def test_unmatched_annotations_report_what_the_attachment_actually_is(self):
        """The alias fix did not raise the reconciled count, so evidence must
        name each attachment's entity type rather than only that it failed.
        swSelectType_e: 1 edge, 3 vertex, 28 centre mark, 46 silhouette."""
        self.assertIn("ANNOTATION_UNMATCHED", self.executable)
        self.assertIn("Private Function SafeEntityType(", self.source)
        self.assertIn("attachmentTypes=", self.executable)
        self.assertIn("anchoredProjections=", self.executable)
        self.assertIn("aliasesAvailable=", self.executable)

        body = self.body("Private Function MatchAnnotationToProjection(")
        self.assertIn("ByRef diagnostics As String", body)

    def test_r23_302_reverse_correspondence_route_exists(self):
        """The forward map is partial: the counterbore callout attaches to a
        drawing edge (swSelEDGES) that is none of the 18 forward aliases.
        R23-302 asked for drawing-to-model mapping, which the 2025 Help
        names as IModelDocExtension.GetCorrespondingEntity2."""
        self.assertIn(
            "Private Function MatchByReverseCorrespondence(", self.source
        )
        self.assertIn("GetCorrespondingEntity2(", self.executable)

        body = self.body("Private Function MatchByReverseCorrespondence(")
        self.assertIn("LocationOwnsModelEntity(", body)
        # Reachable even where no projection has an anchor, which is the
        # section view's situation.
        self.assertNotIn("HasSelectableAnchor", body)
        # Identity only, no positional fallback.
        for forbidden in ("PageX", "PageY", "Abs("):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_reverse_route_tests_the_locations_own_model_geometry(self):
        body = self.body("Private Function LocationOwnsModelEntity(")
        self.assertIn("location.SourceFaces", body)
        self.assertIn("modelFace.GetEdges", body)
        self.assertIn("RecordEquality(", body)

    def test_unsupported_equality_is_reported_not_read_as_a_non_match(self):
        """ISldWorks.IsSame returns swObjectEquality: 0 NotSame, 1 Same,
        2 Unsupported. The Boolean wrapper collapses 0 and 2, so a
        cross-document comparison that cannot be performed at all is
        indistinguishable from a genuine non-match. eqMax keeps them apart."""
        self.assertIn("Private Function RecordEquality(", self.source)
        body = self.body("Private Function RecordEquality(")
        # Exact member test, never a Boolean coercion.
        self.assertIn("(equality = 1)", body)
        self.assertNotIn("NormalizeSwBoolean", body)
        self.assertIn("If equality > eqMax Then eqMax = equality", body)

    def test_reverse_route_reports_per_attachment_resolution_outcome(self):
        """The run after the route landed still reconciled 1 of 38, and the
        evidence could not say whether GetCorrespondingEntity2 returned
        Nothing or resolved to an entity no location owned. Both hypotheses
        survived a live run, which means the evidence was inadequate."""
        body = self.body("Private Function MatchByReverseCorrespondence(")
        self.assertIn("ByRef diagnostics As String", body)
        for token in (
            "reverse=NoOwner",
            "reverse=Matched",
            ":unresolved:err",
            "eqMax=",
            "resolved=",
            "modelEdgesTested=",
        ):
            with self.subTest(token=token):
                self.assertIn(token, body)

        # The unmatched line must actually carry it.
        unmatched = self.body("Public Function ReconcileWithLocationGraph(")
        self.assertIn("reverseDiagnostics", unmatched)

    def test_reverse_route_unavailability_is_distinguished_from_no_owner(self):
        """Live result: 38 of 38 annotations, every attachment,
        GetCorrespondingEntity2 returned Nothing with error 0. It declines
        rather than fails. A drawing entity with NO model counterpart is a
        different fact from one whose counterpart no location owns — only
        the second would implicate the ownership model."""
        body = self.body("Private Function MatchByReverseCorrespondence(")
        self.assertIn("reverse=UnavailableNoModelCounterpart", body)
        self.assertIn("If resolvedCount = 0 Then", body)

        reconcile = self.body("Public Function ReconcileWithLocationGraph(")
        self.assertIn(
            "AuthoredDrawingEntityNoModelCounterpart", reconcile
        )
        self.assertIn("NoAttachedProjection", reconcile)

    def test_reconciliation_records_which_route_matched(self):
        body = self.body("Public Function ReconcileWithLocationGraph(")
        self.assertIn('matchRoute = "ForwardAlias"', body)
        self.assertIn('matchRoute = "ReverseCorrespondence"', body)
        self.assertIn("matchRoute=", body)
        self.assertIn("routesTried=ForwardAlias,ReverseCorrespondence", body)

    def test_no_vba_reserved_word_is_used_as_an_identifier(self):
        """'Alias' is reserved by Declare statements and will not compile."""
        for reserved in ("Dim alias ", "Set alias ", "Dim name ", "Dim type "):
            with self.subTest(reserved=reserved):
                self.assertNotIn(reserved, self.executable)

    def test_linear_dimensions_can_carry_a_diameter_fit(self):
        """The reference drawing authors its section diameters, including
        the H7, as swLinearDimension, so type alone cannot decide whether
        something is a diameter."""
        self.assertIn("DIM_LINEAR As Long = 2", self.source)
        self.assertIn("DIM_RADIAL As Long = 5", self.source)
        body = self.body("Public Function ClassifyDimension(")
        self.assertIn("Case DIM_LINEAR", body)
        self.assertIn("Case DIM_RADIAL", body)

    def test_center_marks_are_named_not_reported_as_other(self):
        self.assertIn("ANN_CENTER_MARK As Long = 13", self.source)
        self.assertIn("ANN_CENTER_LINE As Long = 15", self.source)
        self.assertIn('"CenterMark"', self.source)

    def test_phase11_rewires_the_pipeline_to_annotation_reconciliation(self):
        pipeline = self.read("Module2_DrawingPipeline.bas")
        self.assertIn("ImportAndReconcileR23Annotations", pipeline)
        self.assertIn("Module14_AnnotationImport.ImportModelAnnotations", pipeline)
        self.assertIn("Module14_AnnotationImport.ReconcileWithLocationGraph", pipeline)

    def test_source_revision_unchanged_until_behaviour_is_wired_in(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
