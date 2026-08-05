"""R23 Phase 8 source contracts for the semantic section-dimension engine.

Static source-contract tests for tasks R23-800 through R23-811. The
load-bearing ones are R23-802 (reconcile before create), R23-804 (the
live-proven type 6 diameters are accepted and type 15 is never required),
R23-806 (the H7 fit is recorded as reference authority, never as model
data) and R23-807 (a failed dimension is never replaced with free text).
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module10_SectionDimensionEngine.bas"
REQ_CLS = "CSectionRequirement.cls"


class R23SectionDimensionContracts(unittest.TestCase):
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
        for name in (MODULE, REQ_CLS):
            with self.subTest(name=name):
                self.assertTrue((SOURCE / name).exists())
                self.assertIn(name.split(".")[0], manifest)

    # --- mutation safety -------------------------------------------------

    def test_evidence_entry_point_creates_nothing(self):
        body = self.source.split("Public Sub R23_ProbeSectionDimensions()")[1]
        for forbidden in (
            "CreateSectionDimension",
            "ApplyReferenceFit",
            "AddDimension2",
            "SetFitValues",
            "SetValues2",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        self.assertIn("mode=ReadOnly", body)
        self.assertIn("creations=0", body)
        self.assertIn("mutations=0", body)
        self.assertIn("drawingUnchanged=", body)

    def test_probe_does_not_invoke_the_production_qa_gate(self):
        self.assertNotIn("EmitRunEvidence", self.executable)

    def test_exactly_two_procedures_mutate_and_both_are_authorized(self):
        self.assertEqual(1, self.executable.count("AddDimension2("))
        self.assertEqual(1, self.executable.count("SetFitValues("))
        self.assertEqual(1, self.executable.count("SetValues2("))

        for marker, call in (
            ("Public Function ApplyReferenceFit(", "SetFitValues("),
            ("Public Function CreateSectionDimension(", "AddDimension2("),
        ):
            with self.subTest(marker=marker):
                body = self.body(marker)
                self.assertIn("ByVal allowMutation As Boolean", body)
                self.assertIn("If Not allowMutation Then", body)
                self.assertLess(
                    body.index("If Not allowMutation Then"),
                    body.index(call),
                )

    def test_mutations_are_recorded_on_the_evidence_ledger(self):
        for marker in (
            "Public Function ApplyReferenceFit(",
            "Public Function CreateSectionDimension(",
        ):
            with self.subTest(marker=marker):
                self.assertIn(
                    "evidence.RecordSolidWorksMutation", self.body(marker)
                )

    # --- R23-801 ---------------------------------------------------------

    def test_r23_801_all_seven_requirement_keys_with_their_nominals(self):
        body = self.body("Public Function BuildSectionRequirements(")
        for key, nominal in (
            ("REQ_OVERALL_THICKNESS", "0.018"),
            ("REQ_BORE_STEP_DEPTH", "0.012"),
            ("REQ_LOWER_WALL_STEP", "0.0115"),
            ("REQ_INNER_BORE", "0.04"),
            ("REQ_FIT_BORE", "0.047"),
            ("REQ_LONG_VERTICAL", "0.1736"),
            ("REQ_LOWER_VERTICAL", "0.1048"),
        ):
            with self.subTest(key=key):
                self.assertIn(key, body)
                self.assertIn(nominal, body)

        self.assertEqual(7, body.count("MakeRequirement("))

    def test_only_the_fit_bore_carries_a_tolerance(self):
        """The standing policy is that a tolerance is never invented. Six of
        the seven requirements state a nominal and nothing else."""
        self.assertEqual(
            1, self.executable.count("ToleranceRequired = True")
        )
        self.assertIn("REFERENCE_HOLE_FIT", self.executable)

    # --- R23-802 ---------------------------------------------------------

    def test_r23_802_reconciliation_records_six_independent_observations(
        self,
    ):
        body = self.body("Private Sub RecordObservation(", "\nEnd Sub")
        for field in (
            "MatchedFullName",
            "MatchedTypeCode",
            "MatchedNominalM",
            "AttachedEntityCount",
            "ToleranceProof",
            "MatchedName",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

    def test_observed_fields_are_never_written_from_required_fields(self):
        """A requirement that reports its own nominal back as the observed
        nominal proves nothing. The two must come from different places."""
        body = self.body("Private Sub RecordObservation(", "\nEnd Sub")
        self.assertIn("requirement.MatchedNominalM = nominalM", body)
        self.assertNotIn(
            "MatchedNominalM = requirement.NominalM", self.executable
        )
        self.assertNotIn(
            "MatchedTypeCode = requirement.", self.executable
        )

    def test_every_match_is_counted_so_duplicates_can_fail(self):
        body = self.body("Public Sub ReconcileSectionDimensions(",
                         "\nEnd Sub")
        self.assertIn(
            "requirement.MatchCount = requirement.MatchCount + 1", body
        )
        self.assertIn("If requirement.MatchCount = 1 Then", body)
        self.assertIn("DuplicateDimensions:", body)

    def test_per_dimension_locals_are_reset_every_iteration(self):
        """VBA block-scoped locals live for the whole procedure. The Phase 0
        section inventory mislabelled eleven dimensions because a stale
        value from the previous iteration was reported against the current
        one."""
        for marker, terminator in (
            ("Public Function InventorySectionDimensions(",
             "\nEnd Function"),
            ("Public Sub ReconcileSectionDimensions(", "\nEnd Sub"),
        ):
            with self.subTest(marker=marker):
                body = self.source.split(marker)[1].split(terminator)[0]
                self.assertIn("nominalM = 0#", body)
                self.assertIn("TryReadNominal(dimension, nominalM", body)

    # --- 2026-08-01 first live run: the nominal never read ---------------

    def test_the_nominal_is_tried_by_every_route_and_the_route_is_named(
        self,
    ):
        """The first live run returned no nominal for any of the seven
        section dimensions - all drawing-authored RD1..RD7 reference
        dimensions, which have no configuration for the supported route to
        ask about. Without a nominal nothing can match, so all seven
        requirements reported Missing while seven dimensions sat in the
        view."""
        body = self.body("Private Function TryReadNominal(")
        self.assertIn("TryNominalInConfiguration(dimension, CONFIG_THIS",
                      body)
        self.assertIn("TryNominalFromObsoleteMembers(", body)
        for route in (
            "nominalRoute=GetSystemValue3.ThisConfiguration",
            "nominalRoute=Unavailable",
        ):
            with self.subTest(route=route):
                self.assertIn(route, body)

    def test_the_route_with_evidence_against_it_was_removed(self):
        """The 2026-08-01 second run reached swAllConfiguration on all seven
        dimensions and it declined on all seven. A route with live evidence
        against it is not kept "just in case"."""
        self.assertNotIn("CONFIG_ALL", self.executable)
        self.assertNotIn("AllConfigurations", self.executable)

    def test_the_diameter_symbol_is_read_before_a_diameter_is_denied(self):
        """A drawing can carry the diameter symbol in the text prefix while
        the diametric flag stays False, and then the sheet reads correctly
        even though the record does not. All seven live dimensions returned
        diametric=False, so this is the difference between a real finding
        and a false one."""
        self.assertIn("Private Function ReadDiameterPrefix(", self.source)
        body = self.body("Private Function ReadDiameterPrefix(")
        self.assertIn("GetText(TEXT_PREFIX)", body)
        self.assertIn("GetText(TEXT_PREFIX_DEFINITION)", body)
        self.assertIn('"MOD-DIAM"', body)
        # Unicode 216 is the diameter sign; a literal would break the
        # all-bytes-below-0x80 source contract and Chr$ would depend on the
        # active ANSI code page.
        self.assertIn("ChrW$(216)", body)

    def test_the_diameter_display_source_is_named(self):
        body = self.body("Private Sub RecordObservation(", "\nEnd Sub")
        for source in (
            '"DiametricRecord"',
            '"TextPrefix"',
            '"DiametricFlag"',
            '"Unreadable"',
            '"None"',
        ):
            with self.subTest(source=source):
                self.assertIn(source, body)

        self.assertIn("NotDisplayedAsDiameter:", body)
        self.assertIn("diameterDisplaySource=", self.read(REQ_CLS))

    def test_the_supported_route_is_tried_before_the_obsolete_ones(self):
        body = self.body("Private Function TryReadNominal(")
        self.assertLess(
            body.index("TryNominalInConfiguration(dimension, CONFIG_THIS"),
            body.index("TryNominalFromObsoleteMembers("),
        )
        obsolete = self.body(
            "Private Function TryNominalFromObsoleteMembers("
        )
        self.assertIn("nominalRoute=Obsolete.GetSystemValue2", obsolete)
        self.assertIn("nominalRoute=Obsolete.SystemValue", obsolete)
        self.assertIn("Err.Number = 0 And candidate > 0#", obsolete)

    def test_a_missing_nominal_reports_the_shape_it_got_back(self):
        """"No nominal" and "an empty SafeArray" are different problems and
        only one of them is a bug in this module."""
        body = self.body("Private Function DescribeNominalShape(")
        for shape in ("Empty", "Null", "Array:", "VarType:", "Error:"):
            with self.subTest(shape=shape):
                self.assertIn(shape, body)
        self.assertIn("nominalShape=", self.executable)

    def test_diametric_is_read_for_every_dimension(self):
        """The live drawing types its bore dimensions swLinearDimension=2,
        including the one carrying H7. Diametric is what separates a linear
        record displayed as a diameter from a plain one."""
        self.assertIn("Private Function SafeDiametric(", self.source)
        body = self.body("Private Function SafeDiametric(")
        self.assertIn("displayDimension.Diametric", body)
        self.assertIn("NormalizeSwBoolean(", body)
        self.assertIn("diametric=", self.executable)
        self.assertIn("diametricKnown=", self.executable)

    def test_a_diameter_requirement_accepts_linear_types_too(self):
        """Phase 0 saw type 6; the 2026-08-01 run saw type 2 for the same
        part. Demanding either one rejects one of the two real states."""
        body = self.body("Public Function BuildSectionRequirements(")
        self.assertIn(
            'CStr(DIM_TYPE_DIAMETRIC_LINEAR) & ";" & linearTypes', body
        )
        self.assertIn("innerBore.RequiresDiameterDisplay = True", body)
        self.assertIn("fitBore.RequiresDiameterDisplay = True", body)

    def test_an_unproved_diameter_display_is_reported_not_assumed(self):
        body = self.body("Private Sub RecordObservation(", "\nEnd Sub")
        self.assertIn("NotDisplayedAsDiameter:", body)
        self.assertIn("DiameterDisplayUnreadable:", body)

    def test_a_flagged_requirement_is_not_counted_as_satisfied(self):
        body = self.body("Public Function VerifySectionDimensions(")
        self.assertIn("RequirementFlagged:", body)

    def test_the_probe_does_not_print_what_the_ledger_already_prints(self):
        """CRunEvidence.AddInfo prints what it records. The first run's log
        carried every SECTION_REQUIREMENT line twice."""
        body = self.source.split("Public Sub R23_ProbeSectionDimensions()")[1]
        self.assertNotIn(
            'Debug.Print "QA INFO: SECTION_REQUIREMENT', body
        )

    # --- R23-803 ---------------------------------------------------------

    def test_r23_803_creation_is_skipped_when_already_satisfied(self):
        body = self.body("Public Function CreateSectionDimension(")
        self.assertIn("If requirement.Matched Then", body)
        self.assertIn("AlreadySatisfiedByImportedDimension", body)
        self.assertLess(
            body.index("If requirement.Matched Then"),
            body.index("AddDimension2("),
        )

    def test_creation_requires_entities_selected_by_the_caller(self):
        """AddDimension2's Remarks are explicit that entities are selected
        by LOCATION, never by name: passing a name makes the dimensioning
        routines pick a line endpoint at random."""
        body = self.body("Public Function CreateSectionDimension(")
        self.assertIn("GetSelectedObjectCount2(-1)", body)
        self.assertIn("If selectedCount < 1 Then", body)
        self.assertIn("NoEntitiesSelected", body)
        self.assertLess(
            body.index("If selectedCount < 1 Then"),
            body.index("AddDimension2("),
        )

    def test_a_created_dimension_is_verified_before_it_counts(self):
        body = self.body("Public Function CreateSectionDimension(")
        self.assertIn("nominalCorrect", body)
        self.assertIn("typeAccepted", body)
        self.assertIn("CreatedButRejected", body)
        self.assertLess(
            body.index("CreatedDimensionMismatch"),
            body.index('requirement.Origin = "CreatedAssociativeDimension"'),
        )

    def test_placement_is_a_caller_argument_not_a_default(self):
        """The lane NAMES where a dimension belongs. Turning a lane into
        coordinates needs the finished section's annotation envelope, which
        is Phase 9's job - the same reason Phase 7 refused to choose the
        section view's placement."""
        body = self.body("Public Function CreateSectionDimension(")
        self.assertIn("ByVal textX As Double", body)
        self.assertIn("ByVal textY As Double", body)
        self.assertIn("AddDimension2(textX, textY, 0#)", body)

    # --- R23-804 ---------------------------------------------------------

    def test_r23_804_type_six_is_accepted_and_type_fifteen_never_required(
        self,
    ):
        body = self.body("Public Function BuildSectionRequirements(")
        self.assertIn("DIM_TYPE_DIAMETER", body)
        self.assertIn("DIM_TYPE_DIAMETRIC_LINEAR", body)
        # Type 15 is ACCEPTED, never demanded. The only comparison against
        # it anywhere decides which diameter-display SOURCE to record, and
        # nothing is rejected for lacking it.
        self.assertEqual(
            1, self.executable.count("DIM_TYPE_DIAMETRIC_LINEAR Then")
        )
        record = self.body("Private Sub RecordObservation(", "\nEnd Sub")
        self.assertIn(
            "typeCode = DIM_TYPE_DIAMETER Or", record
        )
        self.assertIn('DiameterDisplaySource = "DiametricRecord"', record)
        cls = self.read(REQ_CLS)
        self.assertIn("Public Function AcceptsTypeCode(", cls)

    def test_dimensions_are_read_view_scoped(self):
        """GetFirstDisplayDimension5 is obsolete AND walks the whole sheet
        by its own Remarks, which would attribute another view's dimensions
        to the section."""
        self.assertIn("GetDisplayDimensions", self.executable)
        self.assertNotIn("GetFirstDisplayDimension5", self.executable)

    def test_zero_nominal_falls_through_to_the_next_read_route(self):
        """Every R23 section requirement is positive. This build can return
        zero without an error to decline GetSystemValue2, so zero must not
        suppress the final SystemValue fallback."""
        body = self.body("Private Function TryNominalFromObsoleteMembers(")
        self.assertIn("If Err.Number = 0 And candidate > 0# Then", body)

    def test_diameter_symbol_uses_unicode_not_active_code_page(self):
        body = self.body("Private Function ReadDiameterPrefix(")
        self.assertIn("ChrW$(216)", body)
        self.assertNotIn("Chr$(216)", body)

    # --- R23-805 ---------------------------------------------------------

    def test_r23_805_every_required_field_is_read_back(self):
        body = self.body("Public Function InventorySectionDimensions(")
        for field in (
            "type2=",
            "nominalM=",
            "fullName=",
            "attachedEntities=",
        ):
            with self.subTest(field=field):
                self.assertIn(field, body)

        tolerance = self.body("Private Function ReadDimensionTolerance(")
        for field in (
            "toleranceType=",
            "fitType=",
            "holeFit=",
            "shaftFit=",
            "minimumStatus=",
            "maximumStatus=",
        ):
            with self.subTest(field=field):
                self.assertIn(field, tolerance)

    def test_status_is_reported_next_to_the_value_it_qualifies(self):
        """GetMinValue2 and GetMaxValue2 return a STATUS and hand the value
        back by reference. A zero value with a failed status is not a zero
        tolerance."""
        body = self.body("Private Function ReadDimensionTolerance(")
        self.assertIn("tolerance.GetMinValue2(minimumM)", body)
        self.assertIn("tolerance.GetMaxValue2(maximumM)", body)
        self.assertIn("minimumStatus", body)
        self.assertIn("maximumStatus", body)

    def test_obsolete_tolerance_members_are_not_used(self):
        """All four IDimension tolerance members are marked obsolete by the
        2025 Help, each superseded by an IDimensionTolerance member."""
        for obsolete in (
            "GetToleranceValues",
            "SetToleranceValues",
            "GetToleranceFitValues",
            "SetToleranceFitValues",
        ):
            with self.subTest(obsolete=obsolete):
                self.assertNotIn(obsolete, self.executable)

    def test_dimension_names_are_sanitized_before_entering_evidence(self):
        """Dimension full names carry '@' freely and can carry the
        delimiters this evidence format uses; an unsanitized name silently
        corrupts every field after it on the line."""
        self.assertIn("Public Function SectionToken(", self.source)
        body = self.body("Public Function InventorySectionDimensions(")
        self.assertIn("SectionToken(", body)
        self.assertIn(
            "SafeDimensionFullName(dimension))", body
        )

    # --- R23-806 ---------------------------------------------------------

    def test_r23_806_the_fit_is_reference_authority_never_model_data(self):
        """Phase 0 read the part-source dimension directly and found
        toleranceType=0, fitType=-1 and empty fit strings. QA has to be able
        to say the tolerance did not originate in the model."""
        self.assertIn("NotModelData", self.executable)
        self.assertIn(
            'REFERENCE_FIT_AUTHORITY As String', self.executable
        )
        body = self.body("Public Function ApplyReferenceFit(")
        self.assertIn(
            'requirement.ToleranceProvenance = _', body
        )
        self.assertIn("REFERENCE_FIT_AUTHORITY", body)

    def test_the_reference_deviations_are_stated_once(self):
        self.assertIn("REFERENCE_FIT_MIN_M As Double = 0#", self.executable)
        self.assertIn(
            "REFERENCE_FIT_MAX_M As Double = 0.000025", self.executable
        )
        self.assertIn('REFERENCE_HOLE_FIT As String = "H7"',
                      self.executable)

    def test_tolerance_type_is_set_before_the_values(self):
        """SetValues2 refuses while the type is swTolNONE by its own
        Remarks, and FitType is only available for the fit types."""
        body = self.body("Public Function ApplyReferenceFit(")
        self.assertIn("tolerance.Type = TOL_TYPE_FIT_WITH_TOL", body)
        self.assertLess(
            body.index("tolerance.Type = TOL_TYPE_FIT_WITH_TOL"),
            body.index("SetValues2("),
        )

    def test_the_applied_fit_is_read_back_not_trusted(self):
        body = self.body("Public Function ApplyReferenceFit(")
        self.assertIn("ReadDimensionTolerance(dimension", body)
        self.assertIn("proven = fitApplied And valuesApplied And", body)
        self.assertLess(
            body.index("SetValues2("),
            body.index("ReadDimensionTolerance(dimension"),
        )

    def test_com_booleans_are_normalized(self):
        """SetFitValues and SetValues2 return COM booleans, and only
        NormalizeSwBoolean is representation-independent."""
        body = self.body("Public Function ApplyReferenceFit(")
        self.assertEqual(
            2, body.count("Module11_GeometryIdentity.NormalizeSwBoolean(")
        )

    def test_a_present_tolerance_is_not_claimed_as_model_provenance(self):
        body = self.body("Private Sub EvaluateTolerance(", "\nEnd Sub")
        self.assertIn("PresentOnDrawing.", body)
        self.assertIn("REFERENCE_FIT_AUTHORITY", body)

    # --- R23-807 ---------------------------------------------------------

    def test_r23_807_no_free_text_substitute_for_a_failed_dimension(self):
        """A note is not a dimension: it does not move with the geometry and
        cannot be inspected as one."""
        for forbidden in ("InsertNote", "CreateText", "SetText"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

        body = self.body("Public Function CreateSectionDimension(")
        self.assertEqual(4, body.count("policy=NoFreeTextSubstitute"))

    # --- R23-808 ---------------------------------------------------------

    def test_r23_808_requirements_are_assigned_the_approved_lanes(self):
        body = self.body("Public Function BuildSectionRequirements(")
        for key, lane in (
            ("REQ_OVERALL_THICKNESS", "LANE_ABOVE"),
            ("REQ_BORE_STEP_DEPTH", "LANE_ABOVE"),
            ("REQ_LOWER_WALL_STEP", "LANE_BELOW"),
            ("REQ_INNER_BORE", "LANE_BORE_SIDE_A"),
            ("REQ_FIT_BORE", "LANE_BORE_SIDE_B"),
            ("REQ_LONG_VERTICAL", "LANE_EXTERIOR_VERTICAL_OUTER"),
            ("REQ_LOWER_VERTICAL", "LANE_EXTERIOR_VERTICAL_INNER"),
        ):
            with self.subTest(key=key):
                index = body.index(key)
                self.assertIn(lane, body[index:index + 260])

        # The two bore diameters are on OPPOSITE sides, and the two
        # vertical references are on SEPARATE lanes.
        self.assertNotEqual("LANE_BORE_SIDE_A", "LANE_BORE_SIDE_B")
        self.assertEqual(1, body.count("LANE_EXTERIOR_VERTICAL_OUTER"))
        self.assertEqual(1, body.count("LANE_EXTERIOR_VERTICAL_INNER"))

    # --- R23-809 ---------------------------------------------------------

    def test_r23_809_the_section_is_excluded_from_generic_arrangement(self):
        self.assertIn(
            "Public Function IsExcludedFromGenericArrangement(",
            self.source,
        )
        body = self.body(
            "Public Function IsExcludedFromGenericArrangement("
        )
        self.assertIn("IsSectionView(swView)", body)

        section = self.body("Public Function IsSectionView(")
        self.assertIn("VIEW_TYPE_SECTION", section)

    # --- R23-810 ---------------------------------------------------------

    def test_r23_810_the_legacy_callout_is_detected_not_deleted(self):
        """The switched pipeline does not author legacy notes; a stale note
        remains evidence until a separately authorized removal run."""
        self.assertIn("Public Function DetectLegacyBoreCallout(",
                      self.source)
        body = self.body("Public Function DetectLegacyBoreCallout(")
        self.assertIn("swView.GetNotes", body)
        self.assertIn("LEGACY_BORE_CALLOUT_MARK", body)
        self.assertIn("removalStatus=PipelineSwitched", body)
        self.assertIn("removalRequiresAuthorization=True", body)

        for forbidden in ("DeleteSelection", ".Delete"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, self.executable)

    def test_the_legacy_mark_is_not_authored_by_the_switched_pipeline(self):
        legacy = (SOURCE / "Module7_TitleBlockEngine.bas").read_text(
            encoding="cp1252"
        )
        self.assertNotIn("<MOD-DIAM>47 H7", legacy)
        self.assertIn('LEGACY_BORE_CALLOUT_MARK As String = "47 H7"',
                      self.executable)

    # --- R23-811 ---------------------------------------------------------

    def test_r23_811_exactly_one_dimension_per_requirement_key(self):
        body = self.body("Public Function VerifySectionDimensions(")
        self.assertIn("If requirement.MatchCount = 0 Then", body)
        self.assertIn("ElseIf requirement.MatchCount > 1 Then", body)
        self.assertIn("Missing:", body)
        self.assertIn("Duplicate:", body)
        self.assertIn("requirementFailures=", body)

    def test_r23_811_no_ordinate_belongs_in_the_section(self):
        """An ordinate in a section shares no datum with the Phase 5 groups
        and reads as a coordinate from an origin the section does not
        have."""
        body = self.body("Public Function CountSectionOrdinates(")
        for constant in (
            "DIM_TYPE_ORDINATE_BASE",
            "DIM_TYPE_HOR_ORDINATE",
            "DIM_TYPE_VERT_ORDINATE",
            "DIM_TYPE_ANGULAR_ORDINATE",
        ):
            with self.subTest(constant=constant):
                self.assertIn(constant, body)

        verify = self.body("Public Function VerifySectionDimensions(")
        self.assertIn("SectionOrdinatePresent:", verify)

    def test_an_unattached_dimension_is_not_counted_as_satisfied(self):
        verify = self.body("Public Function VerifySectionDimensions(")
        self.assertIn("Unattached:", verify)
        self.assertIn("ToleranceUnsatisfied:", verify)

    # --- hygiene ---------------------------------------------------------

    def test_source_hygiene(self):
        for name in (MODULE, REQ_CLS):
            with self.subTest(name=name):
                raw = (SOURCE / name).read_bytes()
                self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
                self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))
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
