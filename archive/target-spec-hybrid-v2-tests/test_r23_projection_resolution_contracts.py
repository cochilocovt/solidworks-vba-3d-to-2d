"""R23 Phase 3 source contracts for drawing-projection resolution.

Static source-contract tests asserting that the projection engine is shaped
according to tasks R23-300 through R23-310, and that the acquisition routes
it uses are the ones Phase 0 actually proved on this build rather than the
ones the documentation implies. They do not execute VBA; the P-0251
projection coverage (R23-310) requires a live run through
``R23_ProbeViewProjections``.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module13_ProjectionResolution.bas"


class R23ProjectionResolutionContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        """Source with whole-line comments removed.

        This module documents the routes it refuses to depend on, so a plain
        text search would match its own commentary.
        """
        return "\n".join(
            line
            for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def body(self, marker: str, terminator: str = "\nEnd Function") -> str:
        chunk = self.source.split(marker)[1]
        return chunk.split(terminator)[0]

    def setUp(self):
        self.source = self.read(MODULE)
        self.executable = self.code(MODULE)

    def test_module_exists_and_is_managed(self):
        self.assertTrue((SOURCE / MODULE).exists())
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")
        self.assertIn("Module13_ProjectionResolution", manifest)

    def test_r23_300_both_component_contexts_are_kept_distinct(self):
        self.assertIn("ResolveViewComponentContexts", self.source)
        self.assertIn("GetVisibleComponents", self.source)
        self.assertIn("GetVisibleDrawingComponents", self.source)
        self.assertIn("drawingComponent.Component", self.source)

        body = self.body("Private Sub ResolveViewComponentContexts(", "\nEnd Sub")
        for state in (
            "NoComponentContext",
            "DrawingContextOnly",
            "ModelContextOnly",
            "Converged",
            "Diverged",
        ):
            with self.subTest(state=state):
                self.assertIn(state, body)

    def test_r23_300_ambiguous_component_context_is_refused(self):
        """More than one visible component is an assembly context this phase
        does not claim to resolve; picking an arbitrary one would silently
        attribute geometry to the wrong component."""
        for reader in (
            "Private Function SingleVisibleComponent(",
            "Private Function SingleModelCapableComponent(",
        ):
            with self.subTest(reader=reader):
                body = self.body(reader)
                self.assertIn("If count <> 1 Then", body)

    def test_r23_301_visible_entities_use_the_edge_type_constant(self):
        self.assertIn("VIEW_ENTITY_EDGE As Long = 1", self.source)
        self.assertIn("VIEW_ENTITY_VERTEX As Long = 2", self.source)
        self.assertIn("VIEW_ENTITY_FACE As Long = 3", self.source)
        self.assertIn("VIEW_ENTITY_SILHOUETTE As Long = 4", self.source)
        self.assertIn("GetVisibleEntities2(component, entityType)", self.source)

    def test_r23_302_and_303_route_a_is_primary_and_route_b_is_not_relied_on(self):
        body = self.body("Private Function MapModelEntityToDrawing(")
        # Route A is tried first and returns immediately on success.
        self.assertLess(
            body.index("swView.GetCorrespondingEntity"),
            body.index("component.GetCorrespondingEntity"),
        )
        self.assertIn('mappedRoute = "A:Direct"', body)
        self.assertIn('mappedRoute = "B:ComponentMediated"', body)

    def test_r23_304_identity_uses_exact_swobjectequality_not_a_boolean(self):
        """ISldWorks.IsSame returns swObjectEquality {0 not same, 1 same,
        2 unable to determine}. Reading it as a Boolean would accept
        'unable to determine' as proof of identity."""
        self.assertIn("OBJECT_EQUALITY_SAME As Long = 1", self.source)
        self.assertIn("OBJECT_EQUALITY_UNSUPPORTED As Long = 2", self.source)

        body = self.body("Private Function FindVisibleEntityIndex(")
        self.assertIn("If equality = OBJECT_EQUALITY_SAME Then", body)
        self.assertNotIn("NormalizeSwBoolean", body)

        # An unreadable comparison must default to "not same".
        fallback = self.body("Private Function SafeObjectEquality(")
        self.assertIn("SafeObjectEquality = 0", fallback)

    def test_r23_304_anchor_requires_visible_inventory_agreement(self):
        """Route A's output is only accepted once route C independently
        confirms it is a drawing-context entity of this view."""
        body = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        self.assertIn("FindVisibleEntityIndex(", body)
        self.assertIn("If visibleIndex < 0 Then", body)
        self.assertIn("MappedEntityNotInVisibleInventory", body)
        # The rejection must come before the anchor is committed.
        self.assertLess(
            body.index("If visibleIndex < 0 Then"),
            body.index("Set projection.PrimaryAnchor = mapped"),
        )

    def test_part_drawing_uses_selection_owner_when_inventory_is_unavailable(
        self,
    ):
        """GetVisibleEntities2 needs a Component2. The runner proved that
        the part drawing exposes none, so selection-owner readback replaces
        only that unavailable confirmation route."""
        build = self.body("Public Function BuildViewProjections(")
        self.assertIn("PartDrawingHasNoComponent2", build)
        self.assertIn("confirmation=SelectionOwnership", build)
        self.assertIn("SelectAnchorInView(swDraw, projection", build)
        self.assertIn("SelectionOwnershipUnproven", build)

        resolve = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        self.assertIn("visibleInventoryAvailable", resolve)
        self.assertIn("If visibleInventoryAvailable Then", resolve)

    def test_correspondence_activates_its_target_view_and_names_declines(self):
        """The unattended runner has no user-selected drawing view. Route A
        must activate the target view first and distinguish Nothing/error 0
        from a COM error before treating a location as unanchorable."""
        build = self.body("Public Function BuildViewProjections(")
        self.assertIn("swDrawing.ActivateView(viewName)", build)
        self.assertIn("PROJECTION_VIEW_ACTIVATION", build)
        self.assertIn("activationError", build)

        mapper = self.body("Private Function MapModelEntityToDrawing(")
        self.assertIn("ByVal modelEntity As Object", mapper)
        self.assertIn("directError = Err.Number", mapper)
        self.assertIn("A:Nothing:err", mapper)
        self.assertIn("B:UnavailableNoComponent", mapper)

    def test_route_d_is_selection_owned_and_cleans_up(self):
        """When a part drawing has no Component2 inventory and Route A
        declines, IView.SelectEntity is allowed only with target-view
        readback and cleanup; a coordinate approximation is never accepted."""
        resolve = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        self.assertIn("SelectModelEntityInView(", resolve)
        self.assertIn("Not visibleInventoryAvailable", resolve)

        route = self.body("Private Function SelectModelEntityInView(")
        self.assertIn("swView.SelectEntity(modelEntity, False)", route)
        self.assertIn("GetSelectedObjectsDrawingView2(1, -1)", route)
        self.assertIn("GetSelectedObject6(1, -1)", route)
        self.assertIn("D:RefusedPreexistingSelection", route)
        self.assertIn("initialSelectionCount =", route)
        self.assertNotIn("swDraw.ClearSelection2 True\n\n    Dim selected", route)

    def test_r23_305_circle_proof_matches_the_phase0_accepted_proof(self):
        self.assertIn("EdgeIsCompleteCircle", self.source)
        body = self.body("Private Function EdgeIsCompleteCircle(")
        self.assertIn("curve.IsCircle", body)
        self.assertIn("CircleParams", body)
        self.assertIn("GetCurveParams3", body)
        self.assertIn("CIRCLE_CLOSURE_TOLERANCE_M", body)
        self.assertIn("NormalizeSwBoolean", body)
        # The raw Boolean is never negated directly.
        self.assertNotIn("If Not curve.IsCircle", self.executable)

    def test_curve_param_data_is_bound_with_set_not_let(self):
        """IEdge.GetCurveParams3 returns an ICurveParamData object, not an
        array. A Let assignment asks it for a default property it does not
        have and raises error 438, which rejected every edge on the first
        instrumented run."""
        self.assertIn(
            "Set curveParams = modelEdge.GetCurveParams3", self.executable
        )
        self.assertNotIn(
            "curveParams = modelEdge.GetCurveParams3", self.executable.replace(
                "Set curveParams = modelEdge.GetCurveParams3", ""
            )
        )
        self.assertIn("As SldWorks.CurveParamData", self.executable)

        # Endpoints and parameter range come from the documented members.
        body = self.body("Private Function EdgeIsCompleteCircle(")
        for member in (
            "curveParams.StartPoint",
            "curveParams.EndPoint",
            "curveParams.UMinValue",
            "curveParams.UMaxValue",
        ):
            with self.subTest(member=member):
                self.assertIn(member, body)

    def test_unmeasurable_closure_fails_closed(self):
        """PointDistance reports failure as a negative value, which would
        slip under the closure tolerance and read as a closed curve."""
        body = self.body("Private Function EdgeIsCompleteCircle(")
        self.assertIn("If closureM < 0# Then", body)
        self.assertIn("ClosureNotMeasurable", body)
        self.assertLess(
            body.index("If closureM < 0# Then"),
            body.index("If closureM > CIRCLE_CLOSURE_TOLERANCE_M Then"),
        )

    def test_selection_cleanliness_is_measured_against_a_baseline(self):
        """The operator may already have something selected; a raw final
        count of 1 said nothing about whether this pass was clean."""
        self.assertIn("initialSelectionCount", self.executable)
        self.assertIn("selectionAttempted", self.executable)
        self.assertIn("selectionClean=", self.executable)
        self.assertIn("operatorSelectionPreserved=", self.executable)

    def test_r23_306_axis_direction_is_differenced_not_transformed_as_a_point(self):
        """Transforming a direction as if it were a point folds in the
        view translation and would make every axis look oblique."""
        body = self.body("Private Function AxisIsNormalToView(")
        self.assertEqual(2, body.count("TransformPointToView("))
        self.assertIn("projectedAxisX = tipX - originX", body)
        self.assertIn("projectedAxisY = tipY - originY", body)
        self.assertIn("projectedAxisZ = tipZ - originZ", body)
        self.assertIn("AXIS_NORMAL_TOLERANCE", body)

    def test_r23_306_normal_axis_is_enforced_not_merely_computed(self):
        """The first anchored run accepted all four M5 tapped holes in the
        front view, where their Y axis lies in the page plane and the
        'circular' edge is seen edge-on. AxisNormalToView was computed and
        stored but never gated acceptance."""
        projection = self.read("CViewHoleProjection.cls")
        body = projection.split("Public Function QualificationFailureReason(")[1]
        body = body.split("\nEnd Function")[0]

        self.assertIn("If Not AxisNormalToView Then", body)
        self.assertIn("AxisNotNormalToView", body)
        # A mapped-but-oblique anchor is a distinct outcome from no anchor.
        self.assertLess(
            body.index("ProjectionAnchorUnavailable"),
            body.index("AxisNotNormalToView"),
        )

        # The verifier must count it separately, not fold it into unanchored.
        verify = self.body("Public Function VerifyExpectedProjections(")
        self.assertIn('Case "AxisNotNormalToView"', verify)

    def test_byref_out_parameters_are_locals_not_class_fields(self):
        """A class Public variable is exposed as a property, so passing
        projection.ProjectedAxisX as a ByRef argument hands the callee a
        temporary that is discarded. The fourth live run logged
        projectedAxis=0,0,0 while axisNormal was correct."""
        body = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        call = body.split("AxisIsNormalToView( _")[1].split(")")[0]
        self.assertNotIn("projection.ProjectedAxis", call)
        self.assertIn("projectedAxisX, projectedAxisY, projectedAxisZ", call)
        # Written back explicitly afterwards.
        for axis in ("X", "Y", "Z"):
            with self.subTest(axis=axis):
                self.assertIn(
                    f"projection.ProjectedAxis{axis} = projectedAxis{axis}",
                    body,
                )

    def test_view_display_mode_is_recorded(self):
        """Under Hidden Lines Removed a far-side hole is never drawn, so
        GetCorrespondingEntity has nothing to return. An unanchorable
        location must be attributable to the display setting rather than
        mistaken for a mapping defect."""
        self.assertIn("GetDisplayMode2", self.executable)
        self.assertIn("displayMode=", self.executable)
        body = self.body("Private Function DisplayModeName(")
        self.assertIn('"HiddenLinesRemoved"', body)
        self.assertIn('"HiddenLinesVisible"', body)

    def test_axis_normality_is_visible_in_the_evidence(self):
        self.assertIn("axisNormal=", self.executable)
        self.assertIn("projectedAxis=", self.executable)
        projection = self.read("CViewHoleProjection.cls")
        self.assertIn("axisNormal=", projection)

    def test_coincident_projections_are_attributed_to_geometry(self):
        """Two holes on a common axis viewed along that axis are ONE circle
        on the sheet, so SOLIDWORKS has a single drawing entity and only one
        model edge maps to it. Without this the second location's failure
        reads as a mapping defect."""
        self.assertIn("Private Sub MarkCoincidentProjections(", self.source)
        self.assertIn("PAGE_COINCIDENCE_M", self.executable)
        self.assertIn("PROJECTION_COINCIDENT", self.executable)
        self.assertIn("OneDrawingEntityForTwoCoaxialHoles", self.executable)

        projection = self.read("CViewHoleProjection.cls")
        self.assertIn("Public CoincidentWithAnchoredKey As String", projection)

        # Only unanchored projections are candidates.
        body = self.body("Private Sub MarkCoincidentProjections(", "\nEnd Sub")
        self.assertIn(
            "If candidate.HasSelectableAnchor() Then GoTo ContinueCandidate",
            body,
        )
        self.assertIn("coincidentUnanchored=", self.executable)

    def test_page_centre_is_recorded_even_without_an_anchor(self):
        """An unanchored location still has a provable page position;
        without it, coincidence cannot be distinguished from a defect."""
        body = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        self.assertIn('If projection.CoordinateFrameProof = "Unproven" Then', body)
        # Recorded before the edge loop, so it does not depend on anchoring.
        self.assertLess(
            body.index('If projection.CoordinateFrameProof = "Unproven" Then'),
            body.index("For edgeIndex ="),
        )

    def test_r23_310_reports_per_view_not_a_single_total(self):
        """R23-310 asks which holes are usable in WHICH view; one global
        count cannot express that."""
        self.assertIn("Public Function ViewAcceptanceSummary(", self.source)
        body = self.body("Public Function ViewAcceptanceSummary(")
        for field in ("projections=", "axisNormal=", "anchored=", "accepted="):
            with self.subTest(field=field):
                self.assertIn(field, body)
        self.assertIn("R23_PROJECTION_VIEW_SUMMARY", self.executable)

    def test_r23_306_page_centre_uses_the_proven_production_transform(self):
        """Phase 0 proved this exact helper produced the datum page
        coordinates; the transform is not reimplemented here."""
        self.assertIn(
            "Module8_RuntimeSupport.TransformPointToView(", self.source
        )
        body = self.body("Private Function ProjectFaceCentreToPage(")
        self.assertIn("CylinderParams", body)
        self.assertIn('frameProof = "Page|source=', body)
        # A failed transform must leave the frame unproven, not default it.
        self.assertIn('frameProof = "Unproven"', body)

    def test_r23_307_selection_is_guarded_restored_and_ownership_proved(self):
        """ISelectData.View is documented get/set but raises error 91 on this
        build, so ownership is proved after the fact, never assumed."""
        self.assertIn("TryBindSelectDataView", self.source)
        self.assertIn("UnboundAfterError:", self.source)

        body = self.body("Public Function SelectAnchorInView(")
        self.assertIn("Select4(False, selectData)", body)
        self.assertIn("GetSelectedObjectsDrawingView2", body)
        self.assertIn("ownershipProven", body)
        self.assertIn("reason=PreexistingSelection", body)
        self.assertIn("initialSelectionCount =", body)
        # The proof begins only from a known-empty selection and clears its
        # own temporary selection on normal and error exits.
        self.assertGreaterEqual(body.count("ClearSelection2 True"), 2)

    def test_r23_502_requires_visible_outline_entities(self):
        """A bottom datum must be a visible model edge, not only mappable.
        Part drawings have no Component2 inventory, so GetPolylines7 is the
        exact visible-model-edge proof."""
        body = self.body("Public Function MapVisibleDatumEntity(")
        self.assertIn("ByRef swDraw As SldWorks.ModelDoc2", body)
        self.assertIn("MapModelEntityToDrawing(", body)
        self.assertIn("FindVisibleModelEdgeIndex(", body)
        self.assertIn("PolylineVisibilityUnavailable", body)
        self.assertIn("visibility=GetPolylines7", body)
        self.assertIn("status=NoEdges", body)
        self.assertIn("visibility=ScopedViewSelection", body)
        self.assertIn("SelectModelEntityInView(", body)
        visibility = self.body("Private Function FindVisibleModelEdgeIndex(")
        self.assertIn("GetPolylines7(", visibility)
        self.assertIn("FindVisibleEntityIndex(", visibility)
        self.assertIn("CROSSHATCH_EXCLUDE", visibility)

    def test_r23_308_anchor_priority_order_is_complete_and_ranked(self):
        """All three tiers are defined even though tier 1 cannot resolve
        until Phase 4 imports annotations; collapsing the order to two tiers
        would hide that the best anchor is not being used."""
        self.assertIn("ANCHOR_TIER_NATIVE_CALLOUT As Long = 1", self.source)
        self.assertIn("ANCHOR_TIER_PRIMARY_DIAMETER As Long = 2", self.source)
        self.assertIn("ANCHOR_TIER_SMALLEST_CIRCLE As Long = 3", self.source)
        self.assertIn("ANCHOR_TIER_NONE As Long = 99", self.source)

        tier = self.body("Private Function AnchorTierFor(")
        # The primary typed diameter is the location's smallest coaxial
        # radius, which the graph already computed.
        self.assertIn("RadiiMatch(", tier)
        self.assertIn("location.PrimaryRadiusM", tier)

    def test_r23_308_every_candidate_is_ranked_not_just_the_first(self):
        """Stopping at the first mappable edge would make the anchor depend
        on face and edge order, so a counterbore could be dimensioned on its
        11 mm mouth instead of its 6.6 mm through hole."""
        body = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        self.assertIn("BetterAnchor(", body)
        # No early exit out of the edge loop once an anchor is found.
        loop = body.split("For edgeIndex =")[1].split("Next edgeIndex")[0]
        self.assertNotIn("Exit Sub", loop)
        self.assertNotIn("Exit For", loop)

        better = self.body("Private Function BetterAnchor(")
        self.assertIn("If candidateTier < bestTier Then", better)
        self.assertIn("BetterAnchor = (candidateRadiusM < bestRadiusM)", better)

    def test_r23_309_unanchored_locations_are_kept_with_an_explicit_reason(self):
        """A location that cannot be projected must survive as a failed
        projection, not vanish from the graph."""
        body = self.body("Public Function BuildViewProjections(")
        self.assertIn("QualificationFailureReason()", body)
        self.assertIn("graph.AddProjection projection", body)
        # AddProjection is unconditional: it is not inside an acceptance test.
        accepted_at = body.index("projection.Accepted = True")
        added_at = body.index("graph.AddProjection projection")
        self.assertLess(accepted_at, added_at)

    def test_anchor_evidence_isolates_which_stage_of_the_chain_broke(self):
        """The first live run reported only candidates=0, which could not
        distinguish 'no faces retained' from 'no circular edge' from
        'nothing mapped'. Each stage must be counted separately."""
        for counter in (
            "sourceFaces=",
            "facesProjected=",
            "boundaryEdges=",
            "circularEdges=",
            "mappedEdges=",
            "inventoryConfirmed=",
            "firstReject=",
        ):
            with self.subTest(counter=counter):
                self.assertIn(counter, self.source)

        body = self.body("Private Sub ResolveProjection(", "\nEnd Sub")
        # The transform's own proof string is carried into the reject reason
        # rather than discarded.
        self.assertIn('firstReject = "FaceCentreNotProjected:" & frameProof', body)

    def test_placeholders_are_skipped_but_part_views_are_not(self):
        """ISheet.GetViews returns named orientation placeholders alongside
        real views. No Component2 is a valid part-drawing limitation, not
        proof that a real view contains no entity."""
        body = self.body("Public Function BuildViewProjections(")
        self.assertIn("IsSheetOrientationPlaceholder", body)
        self.assertIn("PROJECTION_VIEW_SKIPPED", body)
        self.assertIn("SheetOrientationPlaceholder", body)
        self.assertIn("PartDrawingHasNoComponent2", body)
        # The placeholder skip must happen before any location is walked.
        self.assertLess(
            body.index("IsSheetOrientationPlaceholder"),
            body.index("Set locations = graph.Locations()"),
        )

        # View type is recorded but never used as the filter.
        self.assertIn("SafeViewType", self.source)
        self.assertIn("viewType=", self.source)

    def test_r23_310_verification_reports_reasons_not_a_bare_count(self):
        self.assertIn("VerifyExpectedProjections", self.source)
        body = self.body("Public Function VerifyExpectedProjections(")
        for reason in (
            "AcceptedProjectionCount:",
            "ProjectionAnchorUnavailable",
            "CoordinateFrameUnproven",
            "ReferencedConfigurationUnproven",
        ):
            with self.subTest(reason=reason):
                self.assertIn(reason, body)

    def test_evidence_entry_point_is_read_only_and_fixture_guarded(self):
        self.assertIn("Public Sub R23_ProbeViewProjections()", self.source)
        self.assertIn("IsAuthorizedFixture", self.source)
        self.assertIn("drawingUnchanged=", self.source)
        self.assertIn("partUnchanged=", self.source)
        self.assertIn("finalSelectionCount=", self.source)

    def test_projection_pass_never_mutates_either_document(self):
        for forbidden in (
            "ModifyDefinition",
            ".Save3",
            ".SaveAs",
            "EditRebuild3",
            "InsertModelAnnotations",
            "AddOrdinateDimension",
            ".DeleteSelection",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

    def test_locations_retain_source_faces_for_edge_reachability(self):
        """Mapping is per-edge, so a projection must be able to reach every
        contributing face's edges; consolidation must not discard them."""
        location = self.read("CPhysicalHoleLocation.cls")
        self.assertIn("Public SourceFaces As Collection", location)
        self.assertIn("Public Sub AddSourceFace(", location)

        merge = location.split("Public Sub MergeFrom(")[1].split("\nEnd Sub")[0]
        self.assertIn("other.SourceFaces", merge)

        qualification = self.read("Module12_FeatureQualification.bas")
        self.assertIn("candidate.AddSourceFace ownedFace", qualification)

    def test_phase11_rewires_the_pipeline_to_view_projections(self):
        """R23-1101 resolves view anchors before drawing creation stages."""
        pipeline = self.read("Module2_DrawingPipeline.bas")
        self.assertIn("BuildAllViewProjections", pipeline)
        self.assertIn("Module13_ProjectionResolution.BuildViewProjections", pipeline)
        self.assertLess(
            pipeline.index("BuildAllViewProjections"),
            pipeline.index("ImportAndReconcileR23Annotations"),
        )

    def test_source_revision_unchanged_until_behaviour_is_wired_in(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
