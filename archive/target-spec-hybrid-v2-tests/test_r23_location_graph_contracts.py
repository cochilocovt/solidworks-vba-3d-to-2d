"""R23-106 source contracts for the Phase 1 location-graph model.

These are static source-contract tests in the established style of
``test_target_spec_hybrid_v2_source.py``: they read the exported VBA and
assert on the constructs that carry the physical-identity invariants. They do
not execute VBA, so they prove that the required logic is present and shaped
correctly, not that a live run produces a particular drawing.

The invariants under test come from ``docs/R23_IMPLEMENTATION_PLAN.md``
section 5.2 and task R23-106:

* coaxial counterbore steps consolidate into one physical location;
* opposite blind coaxial holes stay separate;
* six counterbores remain six locations;
* four seed/mirror tapped holes remain four locations;
* feature names do not affect physical identity.
"""

from pathlib import Path
import re
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"


class R23LocationGraphSourceContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def test_phase1_graph_components_exist(self):
        required = {
            "CFeatureDefinition.cls",
            "CPhysicalHoleLocation.cls",
            "CViewHoleProjection.cls",
            "CImportedAnnotation.cls",
            "CLocationGraph.cls",
            "Module11_GeometryIdentity.bas",
        }
        self.assertEqual(set(), required - {p.name for p in SOURCE.iterdir()})

    def test_axis_identity_is_sign_normalized_and_moment_based(self):
        """A line and its reverse must produce one identity, and the moment
        must make that identity independent of the sampled axis point."""
        geometry = self.read("Module11_GeometryIdentity.bas")

        self.assertIn("NormalizeAxisDirection", geometry)
        self.assertIn("AxisRequiresSignFlip", geometry)
        self.assertIn("NormalizeAxisAndMoment", geometry)
        self.assertIn("BuildAxisLineKey", geometry)

        # The moment is the cross product of an axis point with the direction.
        self.assertIn("momentX = pointY * axisZ - pointZ * axisY", geometry)
        self.assertIn("momentY = pointZ * axisX - pointX * axisZ", geometry)
        self.assertIn("momentZ = pointX * axisY - pointY * axisX", geometry)

        # Sign normalization must negate the whole vector, not one component.
        self.assertIn("axisX = -axisX", geometry)
        self.assertIn("axisY = -axisY", geometry)
        self.assertIn("axisZ = -axisZ", geometry)

    def test_identity_numbers_are_quantized_and_negative_zero_collapses(self):
        geometry = self.read("Module11_GeometryIdentity.bas")
        self.assertIn("CanonicalNumberToken", geometry)
        self.assertIn("Int((value / effectiveQuantum) + 0.5)", geometry)
        self.assertIn("If Abs(snapped) < (effectiveQuantum * 0.5) Then", geometry)

    def test_coaxial_steps_consolidate_only_when_intervals_meet(self):
        """Consolidation requires the same infinite line AND meeting axial
        intervals. Coaxial-only would wrongly merge opposite blind holes."""
        location = self.read("CPhysicalHoleLocation.cls")

        self.assertIn("Public Function CanConsolidateWith(", location)
        self.assertIn("If Not IsCoaxialWith(other) Then Exit Function", location)
        self.assertIn(
            "If Not AxialIntervalMeets(other) Then Exit Function", location
        )
        self.assertIn("AxialIntervalsOverlap", self.read(
            "Module11_GeometryIdentity.bas"
        ))

    def test_opposite_blind_holes_on_one_axis_stay_separate(self):
        """The axial interval is what distinguishes them, so it must be part
        of the physical key and part of the consolidation decision."""
        location = self.read("CPhysicalHoleLocation.cls")
        geometry = self.read("Module11_GeometryIdentity.bas")

        self.assertIn("AxialMinM", location)
        self.assertIn("AxialMaxM", location)
        self.assertIn("ExtendAxialInterval", location)
        self.assertIn("CanonicalIntervalToken", geometry)
        self.assertIn("interval=", geometry)

    def test_physical_key_excludes_feature_names(self):
        """Renaming a feature must not split or merge a physical location."""
        geometry = self.read("Module11_GeometryIdentity.bas")

        signature = re.search(
            r"Public Function BuildPhysicalLocationKey\((.*?)\) As String",
            geometry,
            re.DOTALL,
        )
        self.assertIsNotNone(signature)
        parameters = signature.group(1).lower()

        self.assertNotIn("featurename", parameters)
        self.assertNotIn("seedfeature", parameters)

        body = geometry.split("Public Function BuildPhysicalLocationKey(")[1]
        body = body.split("End Function")[0]
        self.assertNotIn("FeatureName", body)

    def test_family_key_is_semantic_not_name_based(self):
        """Six identically defined counterbores must share one family key
        regardless of the names SOLIDWORKS gave their features."""
        geometry = self.read("Module11_GeometryIdentity.bas")

        signature = re.search(
            r"Public Function BuildSemanticFamilyKey\((.*?)\) As String",
            geometry,
            re.DOTALL,
        )
        self.assertIsNotNone(signature)
        parameters = signature.group(1).lower()

        self.assertNotIn("featurename", parameters)
        for expected in ("operationkind", "nominaldiameterm", "depthm",
                         "counterborediameterm", "threaddescription"):
            self.assertIn(expected, parameters)

    def test_six_counterbores_remain_six_locations(self):
        """Distinct parallel axes differ only by moment, so the moment must
        be in the line key; otherwise six counterbores would collapse to one."""
        geometry = self.read("Module11_GeometryIdentity.bas")

        body = geometry.split("Public Function BuildAxisLineKey(")[1]
        body = body.split("End Function")[0]
        self.assertIn("CanonicalAxisToken", body)
        self.assertIn("CanonicalMomentToken", body)

    def test_mirrored_instances_remain_distinct_locations(self):
        """Seed and mirror instances are separate physical locations; the
        feature definition tracks the seed without merging the instances."""
        definition = self.read("CFeatureDefinition.cls")
        self.assertIn("SeedFeature As SldWorks.Feature", definition)
        self.assertIn("SeedFeatureName", definition)

        location = self.read("CPhysicalHoleLocation.cls")
        self.assertNotIn("SeedFeatureName", location)

    def test_consolidation_happens_in_exactly_one_place(self):
        graph = self.read("CLocationGraph.cls")
        self.assertIn(
            "Public Function ResolveOrCreatePhysicalLocation(", graph
        )
        self.assertIn("existing.CanConsolidateWith(candidate)", graph)
        self.assertIn("existing.MergeFrom candidate", graph)
        self.assertEqual(1, graph.count("existing.MergeFrom candidate"))

    def test_merge_reindexes_so_no_location_has_two_keys(self):
        graph = self.read("CLocationGraph.cls")
        self.assertIn("ReindexLocation", graph)
        self.assertIn("mLocationsByPhysicalKey.Remove previousKey", graph)
        self.assertIn("IndexedPhysicalKey", graph)

    def test_graph_indexes_every_required_dimension(self):
        graph = self.read("CLocationGraph.cls")
        for index in (
            "mFeaturesByKey",
            "mFeatureKeyByFaceKey",
            "mLocationsByPhysicalKey",
            "mLocationsByLineKey",
            "mLocationsByFamilyKey",
            "mProjectionsByKey",
            "mProjectionsByView",
            "mAnnotationsBySourceIdentity",
            "mAnnotationsByModelDimensionIdentity",
        ):
            with self.subTest(index=index):
                self.assertIn(index, graph)

    def test_annotation_snapshot_reuses_the_graph_dictionary_contract(self):
        graph = self.read("CLocationGraph.cls")
        snapshot = graph.split("Public Sub ClearImportedAnnotations(")[1].split(
            "\nEnd Sub"
        )[0]
        self.assertNotIn("CreateTextDictionary", snapshot)
        self.assertEqual(
            2,
            snapshot.count('CreateObject("Scripting.Dictionary")'),
        )

    def test_projection_fails_closed_without_selectable_anchor(self):
        projection = self.read("CViewHoleProjection.cls")
        self.assertIn("ProjectionAnchorUnavailable", projection)
        self.assertIn("ReferencedConfigurationUnproven", projection)
        self.assertIn("CoordinateFrameUnproven", projection)
        self.assertIn("HasSelectableAnchor", projection)

    def test_projection_records_the_proven_correspondence_route(self):
        """Phase 0 proved route A works and component mediation does not, so
        the route that produced an anchor must stay visible."""
        projection = self.read("CViewHoleProjection.cls")
        self.assertIn("AnchorRoute", projection)
        self.assertIn("GetCorrespondingEntity", projection)

    def test_projection_states_its_coordinate_frame(self):
        projection = self.read("CViewHoleProjection.cls")
        self.assertIn("PageX", projection)
        self.assertIn("PageY", projection)
        self.assertIn("CoordinateFrameProof", projection)
        self.assertIn("frame=Page", projection)

    def test_imported_annotation_separates_provenance_from_value(self):
        """A reference-supplied tolerance must never be recorded as model
        provenance; H7 is absent from the P-0251 model source."""
        annotation = self.read("CImportedAnnotation.cls")
        self.assertIn("ProvenanceSource", annotation)
        self.assertIn("TargetSpecification", annotation)
        self.assertIn("HasNonZeroTolerance", annotation)
        self.assertIn("HasFitData", annotation)

    def test_feature_definition_pairs_every_semantic_field_with_proof(self):
        definition = self.read("CFeatureDefinition.cls")
        for proof in (
            "DiameterProofSource",
            "DepthProofSource",
            "EndConditionProofSource",
            "CounterBoreProofSource",
            "ThreadProofSource",
            "FitProofSource",
            "ToleranceProofSource",
        ):
            with self.subTest(proof=proof):
                self.assertIn(proof, definition)
        self.assertIn("UnprovenSemanticFields", definition)

    def test_ice_normalization_fields_are_retained(self):
        definition = self.read("CFeatureDefinition.cls")
        self.assertIn("RawTypeName2", definition)
        self.assertIn("RawTypeName", definition)
        self.assertIn("EffectiveType", definition)
        self.assertIn("ICE", definition)

    def test_graph_model_is_wired_only_through_the_r23_pipeline(self):
        """R23-1100: production owns one graph, while legacy paths stay out."""
        for name in (
            "Module1_Main.bas",
            "Module3_ModelAudit.bas",
            "Module4_ModelItemImporter.bas",
            "Module5_FallbackDimensionEngine.bas",
            "Module6_QAEngine.bas",
            "Module7_TitleBlockEngine.bas",
            "Module8_RuntimeSupport.bas",
            "Module9_LayoutEngine.bas",
        ):
            with self.subTest(module=name):
                source = self.read(name)
                self.assertNotIn("CLocationGraph", source)
                self.assertNotIn("CPhysicalHoleLocation", source)
                self.assertNotIn("CViewHoleProjection", source)

        pipeline = self.read("Module2_DrawingPipeline.bas")
        self.assertIn("Dim graph As CLocationGraph", pipeline)
        self.assertIn("BuildFeatureCatalog", pipeline)

    def test_new_components_are_in_the_deployment_manifest(self):
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")

        for component in (
            "CFeatureDefinition",
            "CPhysicalHoleLocation",
            "CViewHoleProjection",
            "CImportedAnnotation",
            "CLocationGraph",
            "Module11_GeometryIdentity",
        ):
            with self.subTest(component=component):
                self.assertIn(component, manifest)


if __name__ == "__main__":
    unittest.main()
