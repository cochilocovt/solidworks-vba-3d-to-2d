"""R23 Phase 2 source contracts for feature qualification.

Static source-contract tests in the established style. They assert that the
qualification engine is shaped according to tasks R23-200 through R23-213 and
that the boundaries stated in those tasks are actually present in the source.
They do not execute VBA; the P-0251 catalog itself (R23-213) requires a live
run through ``R23_ProbeFeatureCatalog``.
"""

from pathlib import Path
import re
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module12_FeatureQualification.bas"


class R23FeatureQualificationContracts(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def code(self, name: str) -> str:
        """Source with whole-line comments removed.

        Negative assertions must run against executable code only; this
        module documents the very constructs it forbids, so a plain text
        search would match its own commentary.
        """
        return "\n".join(
            line
            for line in self.read(name).split("\n")
            if not line.lstrip().startswith("'")
        )

    def setUp(self):
        self.source = self.read(MODULE)
        self.executable = self.code(MODULE)

    def test_module_exists_and_is_managed(self):
        self.assertTrue((SOURCE / MODULE).exists())
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")
        self.assertIn("Module12_FeatureQualification", manifest)

    def test_r23_200_effective_type_resolves_ice_through_get_type_name(self):
        self.assertIn("ResolveEffectiveType", self.source)
        self.assertIn("GetTypeName2", self.source)
        self.assertIn("GetTypeName", self.source)
        self.assertIn('TYPE_ICE As String = "ICE"', self.source)
        self.assertIn("IceUnderlyingTypeMissing", self.source)
        self.assertIn("FeatureTypeUnresolved", self.source)

    def test_r23_201_all_three_type_fields_are_recorded(self):
        self.assertIn("definition.RawTypeName2 =", self.source)
        self.assertIn("definition.RawTypeName =", self.source)
        self.assertIn("definition.EffectiveType =", self.source)
        self.assertIn("rawType2=", self.source)
        self.assertIn("rawType1=", self.source)
        self.assertIn("effectiveType=", self.source)

    def test_r23_202_traversal_guard_keys_on_name_and_type_not_address(self):
        """Two live runs over an unchanged P-0251 visited 47 then 46
        features, with different sketches doubled each time. One feature
        reached through two COM wrappers yields two addresses and so two
        keys, which defeats the guard. Names are unique within a part."""
        self.assertIn("BuildTraversalKey", self.source)
        self.assertIn("visitedKeys.Exists(traversalKey)", self.source)
        self.assertIn("GetFirstSubFeature", self.source)
        self.assertIn("GetNextSubFeature", self.source)

        body = self.source.split("Private Function BuildTraversalKey(")[1]
        body = body.split("\nEnd Function")[0]
        executable = "\n".join(
            line for line in body.split("\n")
            if not line.lstrip().startswith("'")
        )
        self.assertIn("featureName", executable)
        self.assertIn("typeName2", executable)

        # The address may appear only on the unnamed-feature fallback, where
        # collapsing distinct features into one key would skip them.
        self.assertIn("If Len(featureName) = 0 Then", executable)
        pointer_lines = [
            line for line in executable.split("\n") if "ObjPtr" in line
        ]
        self.assertEqual(1, len(pointer_lines))
        self.assertIn('"ptr="', pointer_lines[0])

        # The composed key itself must not carry the address.
        composed = executable.split("If Len(featureName) = 0 Then")[1]
        composed = composed.split("BuildTraversalKey = _")[1]
        self.assertNotIn("ObjPtr", composed)

    def test_r23_203_suppression_uses_the_referenced_configuration(self):
        self.assertIn("ProveSuppressionState", self.source)
        self.assertIn("swSpecifyConfiguration", self.source)
        self.assertIn("Array(referencedConfiguration)", self.source)
        self.assertIn("ConfigurationSuppressionUnproven", self.source)
        self.assertIn("FeatureSuppressedInConfiguration", self.source)

    def test_r23_203_active_config_fallback_is_gated_on_name_equality(self):
        """The observed build returns Empty from IsSuppressed2, so a fallback
        is needed; it is only sound when the requested configuration IS the
        active one, otherwise the feature must fail closed."""
        body = self.source.split("Private Function ProveSuppressionState(")[1]
        body = body.split("End Function")[0]

        self.assertIn("activeConfiguration", body)
        self.assertIn("ActiveConfigurationEquivalenceFallback", body)
        # The unproven exit must appear before the fallback is used.
        unproven = body.index("ConfigurationSuppressionUnproven")
        fallback = body.index("ActiveConfigurationEquivalenceFallback")
        self.assertLess(unproven, fallback)

    def test_r23_204_typed_readers_exist_for_each_supported_family(self):
        for reader in (
            "ReadHoleWizardDefinition",
            "ReadAdvancedHoleDefinition",
            "ReadSimpleHoleDefinition",
            "ReadExtrudeCutDefinition",
            "ReadCosmeticThreadDefinition",
            "ResolvePatternSeed",
        ):
            with self.subTest(reader=reader):
                self.assertIn(reader, self.source)

        for interface in (
            "IWizardHoleFeatureData2",
            "ISimpleHoleFeatureData2",
            "IExtrudeFeatureData2",
            "ICosmeticThreadFeatureData",
        ):
            with self.subTest(interface=interface):
                self.assertIn(interface, self.source)

    def test_r23_205_every_access_selections_is_released(self):
        """A successful AccessSelections must be released on success,
        rejection and error alike, through one shared release path."""
        access_calls = self.source.count(".AccessSelections(")
        self.assertGreaterEqual(access_calls, 4)

        self.assertIn("Private Sub ReleaseAccess(", self.source)
        self.assertIn("featureData.ReleaseSelectionAccess", self.source)

        # Every reader that acquires access must call the shared release.
        for reader in (
            "ReadHoleWizardDefinition",
            "ReadSimpleHoleDefinition",
            "ReadExtrudeCutDefinition",
            "ReadCosmeticThreadDefinition",
        ):
            with self.subTest(reader=reader):
                body = self.source.split(f"Private Sub {reader}(")[1]
                body = body.split("\nEnd Sub")[0]
                self.assertIn("AccessSelections", body)
                # Released on the success path and on the error path.
                self.assertGreaterEqual(body.count("ReleaseAccess "), 2)

    def test_r23_206_never_modifies_or_saves_the_model(self):
        self.assertNotIn("ModifyDefinition", self.executable)
        self.assertNotIn(".Save3", self.executable)
        self.assertNotIn(".SaveAs", self.executable)
        self.assertNotIn("EditRebuild3", self.executable)

    def test_r23_207_ownership_comes_from_feature_get_faces(self):
        self.assertIn("swFeature.GetFaces", self.source)
        self.assertIn("CollectOwnedFaces", self.source)
        # IFace2.GetFeature reports only the oldest owner and must not be the
        # ownership proof.
        self.assertNotIn(".GetFeature\n", self.executable)
        self.assertNotIn("ownedFace.GetFeature", self.executable)

    def test_r23_208_face_in_surface_sense_is_not_the_hole_classifier(self):
        self.assertNotIn("FaceInSurfaceSense", self.executable)

    def test_r23_209_only_exact_cut_types_reach_the_extrude_route(self):
        """A substring match would accept unrelated types such as CUTLIST."""
        self.assertIn('TYPE_CUT As String = "CUT"', self.source)
        self.assertIn('TYPE_CUT_THIN As String = "CUTTHIN"', self.source)
        self.assertIn("Case TYPE_CUT, TYPE_CUT_THIN", self.source)

        # No substring containment test against a cut literal.
        self.assertNotIn('InStr(1, definition.EffectiveType, "CUT"', self.executable)
        self.assertNotIn('Like "*CUT*"', self.executable)

    def test_r23_209_boss_features_are_rejected(self):
        self.assertIn("IsBossFeature", self.source)
        self.assertIn("BossFeatureRejected", self.source)
        body = self.source.split("Private Sub ReadExtrudeCutDefinition(")[1]
        body = body.split("\nEnd Sub")[0]
        self.assertIn("ReadRejectedAsBoss", body)

    def test_r23_210_rejections_carry_explicit_reason_codes(self):
        for reason in (
            "FeatureTypeUnresolved",
            "IceUnderlyingTypeMissing",
            "ConfigurationSuppressionUnproven",
            "FeatureSuppressedInConfiguration",
            "DefinitionUnavailable",
            "BossFeatureRejected",
            "NoOwnedGeometry",
            "UnsupportedFeatureType",
            "DegenerateCylinderAxis",
            "AxialIntervalUnavailable",
        ):
            with self.subTest(reason=reason):
                self.assertIn(reason, self.source)

        self.assertIn("RecordRejectedFeature", self.source)
        self.assertIn("FEATURE_REJECTED|", self.source)

    def test_r23_211_seed_chains_reject_missing_circular_and_multiple(self):
        self.assertIn("GetSeedFeature", self.source)
        for reason in (
            "SeedChainUnresolved",
            "SeedChainMultiplyResolved",
            "SeedChainCircular",
        ):
            with self.subTest(reason=reason):
                self.assertIn(reason, self.source)

    def test_r23_211_supported_pattern_types_are_exact_literals(self):
        body = self.source.split("Private Function IsSupportedPatternType(")[1]
        body = body.split("End Function")[0]
        for literal in ("MIRRORPATTERN", "LPATTERN", "CIRPATTERN",
                        "DERIVEDHOLEPATTERN", "LOCALSKETCHPATTERN"):
            with self.subTest(literal=literal):
                self.assertIn(literal, body)
        # The guessed r21 strings must not reappear.
        self.assertNotIn("FILLPATTERN", body)
        self.assertNotIn("VARIABLEPATTERN", body)

    def test_r23_212_coaxial_cylinders_route_through_the_graph(self):
        """Consolidation must stay in CLocationGraph, not be reimplemented."""
        self.assertIn("BuildPhysicalLocations", self.source)
        self.assertIn("graph.ResolveOrCreatePhysicalLocation", self.source)
        self.assertIn("candidate.SetAxisFromSample", self.source)
        self.assertIn("candidate.AddStackMember", self.source)
        self.assertNotIn("CanConsolidateWith", self.executable)

    def test_axial_interval_is_measured_from_face_boundary_edges(self):
        """The interval is what separates opposite blind holes, so it must
        come from real geometry rather than being assumed."""
        self.assertIn("ComputeFaceAxialInterval", self.source)
        self.assertIn("swFace.GetEdges", self.source)
        self.assertIn("GetCurveParams3", self.source)
        self.assertIn("AxialParameter", self.source)

    def test_solidworks_booleans_are_normalized_before_use(self):
        """Raw VARIANT_BOOL values break under Not; every read must pass
        through the shared normalizer."""
        self.assertIn("NormalizeSwBoolean", self.source)
        self.assertIn(
            "Module11_GeometryIdentity.NormalizeSwBoolean(", self.source
        )
        # No bare negation of a SOLIDWORKS Boolean call.
        self.assertNotIn("If Not surface.IsCylinder", self.executable)
        self.assertNotIn("If Not extrudeData.IsBossFeature", self.executable)

    def test_r23_211_seed_rejections_are_enforced_not_merely_computed(self):
        """First catalog run: ResolvePatternSeed set RejectionReason and
        returned, but QualifyFeature never checked it and the accept path
        cleared the field, so every seed-chain reason code was dead."""
        self.assertIn(
            "Private Function ResolvePatternSeed(", self.executable
        )
        body = self.source.split("Private Sub QualifyFeature(")[1]
        body = body.split("\nEnd Sub")[0]
        self.assertIn("If Not ResolvePatternSeed(", body)
        self.assertIn(
            "RecordRejectedFeature definition, _\n"
            "                    definition.RejectionReason",
            body,
        )
        # The rejection must precede the accept path that clears the field.
        self.assertLess(
            body.index("If Not ResolvePatternSeed("),
            body.index("definition.RejectionReason = vbNullString"),
        )

    def test_r23_211_pattern_instances_inherit_seed_semantics(self):
        """A pattern instance carries no definition of its own. Without
        inheritance the four M5 tapped holes split into two families of two
        and the four-location family never forms."""
        self.assertIn("InheritSeedSemantics", self.source)
        self.assertIn("CopySeedSemantics", self.source)

        body = self.source.split("Private Sub CopySeedSemantics(")[1]
        body = body.split("\nEnd Sub")[0]
        for field in (
            "OperationKind",
            "NominalDiameterM",
            "DepthM",
            "EndConditionCode",
            "CounterBoreDiameterM",
            "CounterBoreDepthM",
            "ThreadDescription",
            "FitDescription",
        ):
            with self.subTest(field=field):
                self.assertIn(f"definition.{field} = seedDefinition.{field}", body)

    def test_r23_211_inherited_proof_sources_are_carried_never_invented(self):
        self.assertIn("Private Function PrefixProofSource(", self.source)
        body = self.source.split("Private Function PrefixProofSource(")[1]
        body = body.split("End Function")[0]
        # An unproven seed field must stay unproven on the instance.
        self.assertIn("If Len(Trim$(proofSource)) = 0 Then Exit Function", body)
        self.assertIn("SeedInherited(", self.source)

    def test_r23_211_seed_chain_is_refused_rather_than_followed(self):
        """Recursion has no proven termination on this build."""
        self.assertIn("SeedIsPatternChainUnsupported", self.source)
        body = self.source.split("Private Function InheritSeedSemantics(")[1]
        body = body.split("\nEnd Function")[0]
        self.assertNotIn("ResolvePatternSeed", body)
        self.assertIn("SeedIsBossFeature", body)

    def test_hole_fit_outside_the_enum_is_absent_not_a_value(self):
        """swWzdHoleScrewClearanceTypes_e is {0,1,2}; the 2025 Help limits
        HoleFit to counterbore and countersink. A tapped hole returns -1,
        which must not reach a callout as a clearance."""
        self.assertIn("Private Function ScrewClearanceText(", self.source)
        body = self.source.split("Private Function ScrewClearanceText(")[1]
        body = body.split("End Function")[0]
        self.assertIn("Case 0", body)
        self.assertIn("Case 1", body)
        self.assertIn("Case 2", body)
        self.assertIn('"Close"', body)
        self.assertIn('"Normal"', body)
        self.assertIn('"Loose"', body)
        # No catch-all: an out-of-enum code falls through to empty.
        self.assertNotIn("Case Else", body)

        self.assertNotIn(
            "definition.FitDescription = Trim$(CStr(holeData.HoleFit))",
            self.executable,
        )

    def test_r23_213_catalog_verification_exists(self):
        self.assertIn("VerifyExpectedCatalog", self.source)
        self.assertIn("NoSixLocationFamily", self.source)
        self.assertIn("NoFourLocationFamily", self.source)
        self.assertIn("NoSteppedBoreStack", self.source)

    def test_r23_006_compares_both_curve_read_orders(self):
        """Phase 0 must prove order independence on typed, owned geometry.

        This is a read-only diagnostic: role selection may use feature
        diagnostics, but it must not select arbitrary visible circles or
        mutate the fixture to obtain a result.
        """
        self.assertIn("ProbeCurveReadOrders", self.source)
        self.assertIn("ReadCurveInOrder", self.source)
        self.assertIn("CurveOrderResultsMatch", self.source)
        self.assertIn("R23_CURVE_ORDER|role=", self.source)
        self.assertIn("R23_CURVE_ORDER_END|failures=", self.source)
        for role in (
            "CURVE_ROLE_COUNTERBORE",
            "CURVE_ROLE_TAPPED",
            "CURVE_ROLE_MIRRORED",
            "CURVE_ROLE_EXTRUDED_CUT",
        ):
            with self.subTest(role=role):
                self.assertIn(role, self.source)

        body = self.source.split("Private Function ReadCurveInOrder(")[1]
        body = body.split("\nEnd Function")[0]
        self.assertIn("If r22Order Then Set curveParams", body)
        self.assertIn("If Not r22Order Then Set curveParams", body)
        self.assertIn("Module11_GeometryIdentity.NormalizeSwBoolean", body)
        self.assertIn("swCurve.CircleParams", body)
        self.assertIn("swEdge.GetCurveParams3", body)

    def test_curve_order_mismatch_is_not_hidden_by_a_later_feature(self):
        role = self.source.split("Private Function ProbeCurveRole(")[1]
        role = role.split("\nEnd Function")[0]
        self.assertIn("orderMismatch", role)
        self.assertIn("CurveOrderMismatch:", role)

        edge = self.source.split("Private Function ProbeFeatureCircularEdge(")[1]
        edge = edge.split("\nEnd Function")[0]
        self.assertIn("ByRef orderMismatch As Boolean", edge)
        self.assertIn("If Not matched Then orderMismatch = True", edge)

    def test_evidence_entry_point_is_read_only_and_fixture_guarded(self):
        self.assertIn("Public Sub R23_ProbeFeatureCatalog()", self.source)
        self.assertIn("IsAuthorizedFixture", self.source)
        self.assertIn("modelUnchanged=", self.source)

    def test_phase11_rewires_the_pipeline_to_the_feature_catalog(self):
        """R23-1100 retires the feature-only ownership route."""
        pipeline = self.read("Module2_DrawingPipeline.bas")
        self.assertIn("Module12_FeatureQualification.BuildFeatureCatalog", pipeline)
        self.assertIn("Dim graph As CLocationGraph", pipeline)
        self.assertNotIn("GetAllHoleLikeFeatures(swPart)", pipeline)

    def test_source_revision_unchanged_until_behaviour_is_wired_in(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
