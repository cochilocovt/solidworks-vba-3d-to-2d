"""R23 Phase 6 source contracts for callout reconciliation and fallback.

Static source-contract tests for tasks R23-600 through R23-611. The
load-bearing ones here are R23-609 (no hardcoded callout text or name/radius
scoring in the new path) and the mutation boundary: Phase 6 can create
annotations, and the evidence entry point must not.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module16_CalloutDefinition.bas"
DEFINITION_CLS = "CCalloutDefinition.cls"


class R23CalloutDefinitionContracts(unittest.TestCase):
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

    def test_components_exist_and_are_managed(self):
        manifest = (
            WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"
        ).read_text(encoding="utf-8")
        for name in (MODULE, DEFINITION_CLS):
            with self.subTest(name=name):
                self.assertTrue((SOURCE / name).exists())
                self.assertIn(name.split(".")[0], manifest)

    # --- mutation safety -------------------------------------------------

    def test_evidence_entry_point_creates_nothing(self):
        body = self.source.split("Public Sub R23_ProbeCalloutDefinition()")[1]
        for forbidden in (
            "AddHoleCallout",
            "CreateNativeCalloutForFamily",
            "DeleteSelection",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        self.assertIn("mode=ReadOnly", body)
        self.assertIn("creations=0", body)
        self.assertIn("drawingUnchanged=", body)

    def test_probe_does_not_invoke_the_production_qa_gate(self):
        """Same trap as Phase 5: EmitRunEvidence is the production gate and
        fails closed for a probe that never runs the pipeline stages."""
        self.assertNotIn("EmitRunEvidence", self.executable)

    def test_only_one_procedure_can_mutate_and_it_must_be_authorized(self):
        self.assertEqual(1, self.executable.count("AddHoleCallout2("))
        body = self.body("Public Function CreateNativeCalloutForFamily(")
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("AddHoleCallout2("),
        )

    def test_creation_requires_a_proven_anchor(self):
        """AddHoleCallout2 attaches to whatever edge is selected. Selecting
        an unproven entity would produce an associative callout pointing at
        the wrong hole, which looks correct on the sheet."""
        body = self.body("Public Function CreateNativeCalloutForFamily(")
        self.assertIn("projection.HasSelectableAnchor()", body)
        self.assertIn("NoProvenAnchor", body)
        self.assertIn("selectedCount <> 1", body)
        self.assertLess(
            body.index("NoProvenAnchor"), body.index("AddHoleCallout2(")
        )

    # --- R23-609 ---------------------------------------------------------

    def test_r23_609_no_hardcoded_callout_text_in_the_new_path(self):
        """The legacy path scores candidates by feature name and by
        proximity to an expected radius, and emits literal P-0251 text.
        None of that may exist here."""
        for literal in (
            "6X",
            "4X",
            "THRU",
            "C'BORE",
            "CBORE",
            "M5x0.8",
            "M6",
            "H7",
            "MOD-DIAM",
            "P-0251",
            "DEEP",
        ):
            with self.subTest(literal=literal):
                self.assertNotIn(literal, self.executable)

        # No radius-proximity scoring: those magic numbers are how the
        # legacy path guessed which hole a callout belonged to.
        for magic in ("0.0033", "0.0021", "0.015", "0.011"):
            with self.subTest(magic=magic):
                self.assertNotIn(magic, self.executable)

    def test_attribution_is_identity_not_name_or_proximity(self):
        body = self.body("Public Function MatchCalloutToFamily(")
        self.assertIn("ProjectionOwnsDrawingEntity(", body)
        for forbidden in ("FeatureName", "InStr(", "PageX", "PageY", "Abs("):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        owns = self.body("Private Function ProjectionOwnsDrawingEntity(")
        self.assertIn("ObjectsAreSame(", owns)
        self.assertIn("projection.DrawingEntityAliases", owns)

    def test_identity_uses_exact_swobjectequality(self):
        body = self.body("Private Function ObjectsAreSame(")
        self.assertIn("= 1)", body)
        self.assertNotIn("NormalizeSwBoolean", body)

    # --- R23-600 and R23-601 ---------------------------------------------

    def test_r23_600_callout_bit_is_asserted_not_assumed(self):
        body = self.body("Public Function CalloutImportRequested(")
        self.assertIn("IMPORT_MASK_FULL", body)
        self.assertIn("1048576", body)

    def test_r23_601_classification_is_isholecallout_only(self):
        """A native callout reports Type2 = 6, but so does an ordinary
        diameter dimension, so type must not be the classifier."""
        body = self.body("Public Function IsNativeHoleCallout(")
        self.assertIn("IsHoleCallout", body)
        self.assertIn("NormalizeSwBoolean(", body)
        self.assertNotIn("Type2", body)
        # No dimension-type constant exists to be reached for.
        self.assertNotIn("DIM_DIAMETER", self.executable)

    # --- R23-602 ---------------------------------------------------------

    def test_r23_602_a_callout_resolving_to_two_families_is_rejected(self):
        """One family, or none. Breaking the tie would attribute a callout
        to a family it may not describe."""
        body = self.body("Public Function MatchCalloutToFamily(")
        self.assertIn("AmbiguousFamilies", body)
        self.assertIn("MatchCalloutToFamily = vbNullString", body)

    # --- R23-603 ---------------------------------------------------------

    def test_r23_603_fields_come_from_callout_variables_not_text(self):
        """GetHoleCalloutVariables exposes HoleFit, ShaftFit, ToleranceType,
        ToleranceMin and ToleranceMax per variable. Parsing the rendered
        string instead would give something that cannot be validated field
        by field."""
        body = self.body(
            "Public Sub ReadNativeCalloutFields(", "\nEnd Sub"
        )
        self.assertIn("GetHoleCalloutVariables", body)
        for field in (
            "HoleFit",
            "ShaftFit",
            "ToleranceType",
            "ToleranceMin",
            "ToleranceMax",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

        self.assertNotIn("GetText", body)

    # --- R23-604 and R23-605 ---------------------------------------------

    def test_r23_605_fallback_only_when_native_is_incomplete(self):
        body = self.body("Public Function RetainDefinitionForFamily(")
        self.assertIn("nativeDefinition.IsComplete()", body)
        self.assertIn("retained=NativeCallout", body)
        self.assertIn("retained=ControlledFallback", body)
        # The reason a native definition lost must be recorded.
        self.assertIn("nativeMissing=", body)
        self.assertIn("NoNativeCalloutAttributedToFamily", body)

    # --- R23-606 ---------------------------------------------------------

    def test_r23_606_quantity_is_unique_physical_locations(self):
        """Not a feature count: one Hole Wizard feature plus a mirror makes
        many holes. Not an edge count: a counterbore contributes several
        edges per hole."""
        body = self.body("Public Function BuildDefinitionFromTypedData(")
        self.assertIn("graph.LocationsForFamily(familyKey)", body)
        self.assertIn("UniquePhysicalLocations", body)

        # Quantity is assigned exactly once, from the location count. The
        # stack IS iterated in this procedure to pull typed fields, so the
        # assertion has to name the assignment rather than ban the token.
        assignments = [
            line.strip()
            for line in body.split("\n")
            if line.strip().startswith("definition.Quantity =")
        ]
        self.assertEqual(["definition.Quantity = locations.Count"],
                         assignments)

        for forbidden in ("GetEdges", "FeatureName"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    # --- R23-607 and R23-608 ---------------------------------------------

    def test_r23_607_counterbore_and_thread_come_from_typed_data(self):
        body = self.body(
            "Private Sub ApplyTypedFeatureFields(", "\nEnd Sub"
        )
        for field in (
            "featureDefinition.CounterBoreDiameterM",
            "featureDefinition.CounterBoreDepthM",
            "featureDefinition.ThreadDescription",
            "featureDefinition.ThreadDepthM",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

        # Each copied field carries the proof source with it.
        self.assertIn("featureDefinition.CounterBoreProofSource", body)
        self.assertIn("featureDefinition.ThreadProofSource", body)

    def test_r23_608_fit_comes_only_from_source_model_data(self):
        body = self.body(
            "Private Sub ApplyTypedFeatureFields(", "\nEnd Sub"
        )
        self.assertIn("featureDefinition.FitDescription", body)
        self.assertIn("featureDefinition.FitProofSource", body)

    def test_drawing_authored_tolerances_are_not_promoted_here(self):
        """Standing instruction: a tolerance read off one of the designers'
        drawings is evidence a designer typed a number, not a manufacturing
        requirement."""
        self.assertNotIn("DrawingAuthored", self.executable)
        self.assertNotIn("MatchesDowelToleranceConvention", self.executable)

    # --- R23-610 ---------------------------------------------------------

    def test_r23_610_failure_reasons_are_field_specific(self):
        definition = self.read(DEFINITION_CLS)
        body = definition.split(
            "Public Function CompletenessFailureReason()"
        )[1].split("End Function")[0]
        for field in (
            "Quantity",
            "NominalDiameter",
            "EndConditionUnproven",
            "Depth",
            "CounterBoreDepth",
            "ThreadUnproven",
            "Attachment",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

    def test_depth_is_only_required_for_a_blind_hole(self):
        """swEndConditions_e: swEndCondBlind = 0, and every through/up-to
        member is non-zero. A BLIND hole needs a depth; a through hole has
        none to give.

        The first live run proved this was written inverted: it demanded a
        depth from the counterbore family (EndCondition 2, ThroughNext) and
        let the M5 family pass with EndCondition 0 and no depth at all. The
        second half is the dangerous one - a genuinely missing depth on a
        blind tapped hole went unreported."""
        definition = self.read(DEFINITION_CLS)
        body = definition.split(
            "Public Function CompletenessFailureReason()"
        )[1].split("End Function")[0]
        self.assertIn("EndConditionCode = 0 And DepthM <= 0#", body)
        self.assertNotIn("EndConditionCode <> 0 And DepthM", body)
        self.assertLess(
            body.index("EndConditionUnproven"),
            body.index("EndConditionCode = 0 And DepthM"),
        )

    def test_diameter_falls_back_to_proven_location_geometry(self):
        """The first live run returned diameterM=0 with
        diameterSource=Unproven for every family: no Hole Wizard or cut
        feature on this build declared a nominal diameter. The location
        knows it anyway, from its own cylindrical face and closed circular
        boundary, which is measured geometry rather than a declared
        parameter."""
        body = self.body("Public Function BuildDefinitionFromTypedData(")
        self.assertIn("firstLocation.PrimaryRadiusM", body)
        self.assertIn("* 2#", body)
        self.assertIn("ProvenCylindricalFaceRadius", body)
        # The fallback must not overwrite a diameter a feature did prove.
        self.assertIn("definition.NominalDiameterM <= 0#", body)

    def test_attachment_is_proved_from_family_geometry(self):
        """The first live run returned attachmentProven=False on every
        retained definition, including families whose holes are anchored in
        a view. Only the native path set the flag, so a fallback could never
        earn it - and a native definition losing on completeness discarded
        the proof it had. Attachment is a property of the family's geometry,
        not of a callout."""
        self.assertIn(
            "Private Function FamilyHasAnchoredProjection(", self.source
        )
        body = self.body("Private Function FamilyHasAnchoredProjection(")
        self.assertIn("projection.HasSelectableAnchor()", body)
        self.assertIn("SemanticFamilyKey", body)

        built = self.body("Public Function BuildDefinitionFromTypedData(")
        self.assertIn("FamilyHasAnchoredProjection(graph, familyKey)", built)

        # ATTACHED and ATTACHABLE are different claims and must stay
        # distinguishable in the proof source.
        self.assertIn("AttachableNotYetAttached", self.executable)

    def test_r23_603_depth_is_read_from_typed_callout_variables(self):
        """The M5 family's Hole Wizard feature returned depthM=0 on this
        build, but its native callout carries a Tap Drill Depth variable.
        swCalloutVariable_e verified: Depth 14, Hole_Depth 21,
        Tap_Drill_Depth 28. Thread depth is a different quantity and is
        carried separately, so it is NOT among the accepted members."""
        self.assertIn("CALLOUT_VAR_TAP_DRILL_DEPTH As Long = 28", self.source)
        self.assertIn("CALLOUT_VAR_HOLE_DEPTH As Long = 21", self.source)
        self.assertIn("CALLOUT_VAR_DEPTH As Long = 14", self.source)
        self.assertNotIn("CALLOUT_VAR_THREAD_DEPTH", self.executable)

        body = self.body(
            "Public Sub ReadNativeCalloutFields(", "\nEnd Sub"
        )
        self.assertIn("calloutVariable.VariableType", body)
        self.assertIn("calloutVariable.Length", body)
        # Never overwrite a depth the feature already proved.
        self.assertIn("definition.DepthM <= 0# And lengthValue > 0#", body)

    def test_family_keys_are_escaped_before_entering_evidence(self):
        """A semantic family key is itself a delimited string - it looks
        like "op=HOLEWIZARD|dia=...|thread=M6". Emitting it raw into a
        pipe-delimited line destroys that line's structure, which the first
        live run demonstrated: R23_CALLOUT_END could not be parsed field by
        field. The key itself is unchanged; only its rendering escapes."""
        self.assertIn("Public Function EvidenceToken(", self.source)
        token = self.body("Public Function EvidenceToken(")
        self.assertIn('Replace$(result, "|"', token)
        self.assertIn('Replace$(result, "="', token)

        # Every family-key emission goes through it.
        for line in self.executable.split("\n"):
            if "FamilyKey" in line and "&" in line:
                with self.subTest(line=line.strip()):
                    self.assertIn("EvidenceToken", line)

        summary = self.read(DEFINITION_CLS).split(
            "Public Function Summary()"
        )[1]
        self.assertIn("EvidenceToken(FamilyKey)", summary)
        self.assertIn("EvidenceToken(MachiningFaceKey)", summary)

    def test_verification_reports_reasons_not_a_count(self):
        body = self.body("Public Function VerifyManufacturingDefinitions(")
        self.assertIn("definitionFailures=", body)
        self.assertIn("CompletenessFailureReason()", body)

    # --- R23-611 ---------------------------------------------------------

    def test_r23_611_required_shapes_are_not_keyed_to_a_part_number(self):
        """One multi-hole counterbored family and one multi-hole threaded
        family. P-0251 satisfies it, but the rule must not name it."""
        body = self.body("Public Function VerifyRequiredDefinitionShapes(")
        self.assertIn("NoCompleteCounterboredFamilyDefinition", body)
        self.assertIn("NoCompleteThreadedFamilyDefinition", body)
        self.assertIn("definition.Quantity < 2", body)
        self.assertIn("definition.IsComplete()", body)

        # A thread DESCRIPTION alone does not make a family threaded. The
        # counterbore family reads thread="M6" with threadDepthM=0 - the
        # fastener size of a clearance hole, not a tapped thread.
        self.assertIn("definition.ThreadDepthM > 0#", body)
        for forbidden in ("P-0251", "= 6", "= 4"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    # --- hygiene ---------------------------------------------------------

    def test_source_hygiene(self):
        for name in (MODULE, DEFINITION_CLS):
            with self.subTest(name=name):
                raw = (SOURCE / name).read_bytes()
                self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
                self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))
                # Every byte must be representable in the ANSI codepage the
                # SWP round-trips; a stray UTF-8 sequence breaks readback.
                self.assertTrue(all(b < 0x80 for b in raw))
                text = raw.decode("cp1252")
                self.assertTrue(text.startswith("Option Explicit"))
                for line in text.split("\r\n"):
                    self.assertLessEqual(len(line), 79)

    def test_source_revision_identifies_wired_r23_pipeline(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
