"""R23 Phase 5 source contracts for ordinate schemes and the transaction.

Static source-contract tests for tasks R23-500 through R23-509. Phase 5 is
the phase that CREATES dimensions, so the mutation-safety contracts matter
most: the evidence entry point must stay read-only so it can be run against
the manual reference drawing, and the one mutating procedure must refuse
unless explicitly authorized.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module15_OrdinateScheme.bas"
SCHEME_CLS = "COrdinateScheme.cls"
BUCKET_CLS = "COrdinateBucket.cls"


class R23OrdinateSchemeContracts(unittest.TestCase):
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
        for name in (MODULE, SCHEME_CLS, BUCKET_CLS):
            with self.subTest(name=name):
                self.assertTrue((SOURCE / name).exists())
                self.assertIn(name.split(".")[0], manifest)

    # --- mutation safety -------------------------------------------------

    def test_evidence_entry_point_creates_nothing(self):
        """Phase 5 creates dimensions. The probe must not."""
        body = self.source.split("Public Sub R23_ProbeOrdinateScheme()")[1]
        for forbidden in (
            "CreateOrdinateGroup",
            "AddOrdinateDimension",
            "DeleteSelection",
            ".Save",
        ):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

        self.assertIn("mode=ReadOnly", body)
        self.assertIn("creations=0", body)
        self.assertIn("initialSelectionCount=", body)
        self.assertIn("finalSelectionCount=", body)
        self.assertIn("drawingUnchanged=", body)

    def test_probe_does_not_invoke_the_production_qa_gate(self):
        """Module6_QAEngine.EmitRunEvidence runs RequireCoreStages, which
        demands fourteen pipeline stages a read-only probe never performs.
        Calling it makes a probe that succeeded report RESULT: FAIL and pop
        the production fail-closed dialog. The Phase 3 and Phase 4 probes do
        not call it. Evidence still reaches the Immediate window because
        CRunEvidence.AddInfo Debug.Prints every line itself."""
        self.assertNotIn("EmitRunEvidence", self.executable)

        for module in (
            "Module13_ProjectionResolution.bas",
            "Module14_AnnotationImport.bas",
        ):
            with self.subTest(module=module):
                self.assertNotIn("EmitRunEvidence", self.code(module))

    def test_only_one_procedure_can_mutate_and_it_must_be_authorized(self):
        self.assertEqual(1, self.executable.count("AddOrdinateDimension("))
        body = self.body("Public Function CreateOrdinateGroup(")
        self.assertIn("ByVal allowMutation As Boolean", body)
        self.assertIn("If Not allowMutation Then", body)
        self.assertIn("MutationNotAuthorized", body)
        # The refusal must come before anything is selected or created.
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("AddOrdinateDimension("),
        )

    def test_creation_refuses_when_the_datum_is_unproven(self):
        """The datum is the first selection and everything is measured from
        it, so an unproven datum must stop the transaction rather than
        produce a group based at an unverified entity."""
        body = self.body("Public Function CreateOrdinateGroup(")
        self.assertIn("If Not scheme.DatumResolved Then", body)
        self.assertIn("DatumNotProven", body)
        self.assertLess(
            body.index("If Not scheme.DatumResolved Then"),
            body.index("AddOrdinateDimension("),
        )

    # --- R23-500 ---------------------------------------------------------

    def test_r23_500_scheme_key_replaces_family_grouping(self):
        """view role + machining face + datum policy + direction, every
        part measured rather than read off a name."""
        scheme = self.read(SCHEME_CLS)
        key = scheme.split("Public Function SchemeKey()")[1].split(
            "End Function"
        )[0]
        for part in ("role=", "face=", "datumPolicy=", "direction="):
            with self.subTest(part=part):
                self.assertIn(part, key)

        # The machining face comes from the location's own axis.
        face = self.body("Public Function MachiningFaceKey(")
        self.assertIn("CanonicalAxisToken(", face)
        self.assertIn("location.AxisX", face)

        # The view role reuses the measured eligibility tests, not names.
        role = self.body("Public Function ClassifyViewRole(")
        self.assertIn("IsDeferredCreationView(", role)
        self.assertIn("IsOrdinateEligibleView(", role)

    def test_direction_is_never_inferred_from_selection_order(self):
        """swOrdinate (1) lets SOLIDWORKS infer direction from the selected
        points, which would make the result depend on selection order
        instead of on the scheme."""
        self.assertIn("Public Const ORD_VERTICAL As Long = 2", self.source)
        self.assertIn("Public Const ORD_HORIZONTAL As Long = 3", self.source)
        create = self.body("Public Function CreateOrdinateGroup(")
        self.assertIn("scheme.Direction", create)
        self.assertNotIn("ORD_INFERRED", create)

    # --- R23-501 and R23-502 ---------------------------------------------

    def test_datum_policies_are_part_of_the_scheme_key(self):
        self.assertIn("DATUM_POLICY_X", self.source)
        self.assertIn("DATUM_POLICY_Y", self.source)
        body = self.body("Public Function DatumPolicyIdFor(")
        self.assertIn("DATUM_POLICY_X", body)
        self.assertIn("DATUM_POLICY_Y", body)

    def test_datum_selectability_is_proved_not_assumed(self):
        body = self.body("Public Function ProveBucketSelectable(")
        self.assertIn("GetSelectedObjectsDrawingView2", body)
        self.assertIn("ownershipProven", body)
        self.assertIn("NormalizeSwBoolean(", body)
        # A proof pass leaves nothing selected.
        self.assertIn("ClearSelection2 True", body)

    def test_datum_kind_is_recorded_separately_from_datum_policy(self):
        """The policy says what was asked for; the kind says what the datum
        actually is. A vertical datum must never read a projected hole as
        an outline datum."""
        self.assertIn("ProjectionDerived", self.source)
        self.assertIn("OutlineDerived", self.source)
        scheme = self.read(SCHEME_CLS)
        self.assertIn("Public DatumKind As String", scheme)
        self.assertIn("Public DatumProof As String", scheme)

    def test_r23_502_vertical_datum_is_a_mapped_bottom_outline_edge(self):
        datum = self.body("Private Sub ResolveSchemeDatum(", "\nEnd Sub")
        self.assertIn("ResolveOutlineDatum", datum)
        self.assertIn("DATUM_KIND_OUTLINE", datum)

        outline = self.body("Private Function ResolveOutlineDatum(")
        for token in (
            "GetBodies2(SOLID_BODY_TYPE, True)",
            "swBody.GetEdges",
            "OutlineDatumForModelEdge",
            "ReadViewOutlineBottom",
        ):
            with self.subTest(token=token):
                self.assertIn(token, outline)

        candidate = self.body("Private Function OutlineDatumForModelEdge(")
        for token in (
            "GetStartVertex",
            "GetEndVertex",
            "curve.IsLine",
            "TransformPointToView(",
            "MapVisibleDatumEntity(",
            "VisibleMapUnavailable",
            "viewOutlineGapM=",
            "datumKind=OutlineDerived",
        ):
            with self.subTest(token=token):
                self.assertIn(token, candidate)

        self.assertIn("DatumEvidenceToken(chosen.AnchorProof)", self.source)
        self.assertIn("|mapSample=", self.source)

    # --- R23-505 ---------------------------------------------------------

    def test_r23_505_deduplicates_coordinates_but_credits_every_location(
        self,
    ):
        """Phase 3 proved two coaxial holes share ONE drawing entity at ONE
        page coordinate. Dimensioning it twice is a duplicate; crediting
        only one hole silently drops the other."""
        bucket = self.read(BUCKET_CLS)
        self.assertIn("Public Sub CreditLocation(", bucket)
        self.assertIn("Public Function AlreadyCredited(", bucket)
        self.assertIn("RecordSchemeDirectionalCoverage projection", bucket)

        coverage = bucket.split(
            "Private Sub RecordSchemeDirectionalCoverage("
        )[1]
        self.assertIn("Module15_OrdinateScheme.ORD_HORIZONTAL", coverage)
        self.assertIn("projection.SchemeCoveredX = True", coverage)
        self.assertIn("projection.SchemeCoveredY = True", coverage)

        created = bucket.split("Public Sub MarkCreatedDirectionalCoverage(")[1]
        self.assertIn("projection.CoveredX = True", created)
        self.assertIn("projection.CoveredY = True", created)
        self.assertIn("CreatedOrdinateBucket:", created)

        create = self.body("Public Function CreateOrdinateGroup(")
        self.assertIn("If createdCount <= 0 Then", create)
        self.assertIn("MarkSchemeCreatedCoverage scheme", create)

        body = self.body("Private Sub PopulateSchemeLedger(", "\nEnd Sub")
        self.assertIn("scheme.BucketByKey(coordinateKey)", body)
        self.assertIn("bucket.CreditLocation", body)
        self.assertIn("CreditCoincidentLocations", body)

        partner = self.body(
            "Private Sub CreditCoincidentLocations(", "\nEnd Sub"
        )
        self.assertIn("OneDrawingEntityForTwoCoaxialHoles", partner)

    def test_coincidence_link_is_read_from_the_unanchored_end(self):
        """MarkCoincidentProjections sets CoincidentWithAnchoredKey on the
        UNANCHORED projection, pointing at its anchored twin. Reading it
        from the anchored end finds nothing, because a projection that has
        an anchor is never marked - that dropped two of P-0251's four side
        holes and produced credited=8 of an expected 10."""
        producer = (SOURCE / "Module13_ProjectionResolution.bas").read_text(
            encoding="cp1252"
        )
        mark = producer.split("Private Sub MarkCoincidentProjections(")[1]
        mark = mark.split("\nEnd Sub")[0]
        # The setter guards on the candidate NOT having an anchor.
        self.assertIn("If candidate.HasSelectableAnchor() Then", mark)
        self.assertIn("candidate.CoincidentWithAnchoredKey = ", mark)

        # The consumer must therefore iterate projections itself rather than
        # read the field off an already-bucketed (anchored) projection.
        body = self.body(
            "Private Sub CreditCoincidentLocations(", "\nEnd Sub"
        )
        self.assertIn("projection.CoincidentWithAnchoredKey", body)
        self.assertIn("BucketCreditingLocation(scheme, partnerKey)", body)
        self.assertIn(
            "bucket.CreditLocation projection.PhysicalInstanceKey", body
        )

    def test_coincident_credit_runs_after_every_bucket_exists(self):
        """The partners can only be credited once the anchored projections
        have created their buckets, so the pass must be second."""
        body = self.body("Private Sub PopulateSchemeLedger(", "\nEnd Sub")
        self.assertLess(
            body.index("bucket.CreditLocation"),
            body.index("CreditCoincidentLocations"),
        )

    def test_coordinate_key_is_the_quantized_page_coordinate(self):
        body = self.body("Public Function CoordinateKeyFor(")
        self.assertIn("projection.PageX", body)
        self.assertIn("projection.PageY", body)
        self.assertIn("QuantizeCoordinate(", body)

    # --- R23-507 ---------------------------------------------------------

    def test_r23_507_profile_entries_never_enter_the_small_hole_ledger(self):
        """P-0251's stepped bore is excluded because it is a singleton
        family, not because of a magic radius threshold that would
        misclassify a different part."""
        body = self.body("Public Function IsSmallHoleLocation(")
        self.assertIn("LocationsForFamily(", body)
        self.assertIn("MIN_FAMILY_SIZE_FOR_SMALL_HOLE", body)

        ledger = self.body("Private Sub PopulateSchemeLedger(", "\nEnd Sub")
        self.assertIn("scheme.ProfileProjections.Add", ledger)

        scheme = self.read(SCHEME_CLS)
        self.assertIn("Public ProfileProjections As Collection", scheme)

    # --- R23-508 ---------------------------------------------------------

    def test_r23_508_transaction_follows_the_documented_order(self):
        body = self.body("Public Function CreateOrdinateGroup(")
        order = [
            "ActivateView",
            "TryBindSelectDataView",
            "scheme.DatumBucket",
            "BucketsInDeterministicOrder",
            "AddOrdinateDimension(",
            "SetPickMode",
            "ReadBackCreatedOrdinates(",
        ]
        positions = [body.index(token) for token in order]
        self.assertEqual(positions, sorted(positions))

        # SetPickMode must run whatever the result was: the Remarks say
        # later selections keep appending to the group until it does.
        self.assertLess(body.index("SetPickMode"), body.index("If result"))

    def test_selection_count_is_verified_at_every_append(self):
        """A silently ignored selection would produce a group short by one
        hole and still look successful."""
        body = self.body("Private Function AppendBucketSelection(")
        self.assertIn("GetSelectedObjectCount2(-1)", body)
        self.assertIn("actualCount <> expectedCount + 1", body)

    def test_append_order_is_deterministic(self):
        """Two runs over identical geometry must not produce differently
        ordered groups."""
        body = self.body("Public Function BucketsInDeterministicOrder(")
        self.assertIn("BucketOrdinateValue(", body)

    def test_every_ordinate_error_member_is_decoded_by_name(self):
        body = self.body("Public Function DecodeOrdinateError(")
        for member in (
            "swCreateOrdDimErr_Success",
            "swCreateOrdDimErr_OrdFailure",
            "swCreateOrdDimErr_GenNoInternalDims",
            "swCreateOrdDimErr_GenBadSel",
            "swCreateOrdDimErr_GenNeedModelLoaded",
            "swCreateOrdDimErr_GenSamePartOnly",
            "swCreateOrdDimErr_GenExtraSelection",
            "swCreateOrdDimErr_GenFailure",
            "swCreateOrdDimErr_OrdDupInGroup",
            "swCreateOrdDimErr_OrdBadDir",
            "swCreateOrdDimErr_Undefined",
        ):
            with self.subTest(member=member):
                self.assertIn(member, body)

    def test_readback_is_view_scoped_and_identity_based(self):
        """IView.GetFirstDisplayDimension5 is obsolete AND its own Remarks
        say the GetNext5 walk covers the drawing SHEET, so a read-back built
        on it would credit other views' dimensions to this scheme."""
        self.assertNotIn("GetFirstDisplayDimension5", self.executable)
        self.assertIn("GetDisplayDimensions", self.executable)

        body = self.body("Private Function ReadBackCreatedOrdinates(")
        self.assertIn("CollectionContainsObject(", body)

        identity = self.body("Private Function CollectionContainsObject(")
        self.assertIn("swApp.IsSame(", identity)
        self.assertIn("equality = 1", identity)
        self.assertNotIn("NormalizeSwBoolean", identity)

    def test_api_result_alone_is_never_treated_as_success(self):
        body = self.body("Public Function CreateOrdinateGroup(")
        self.assertIn("createdReadBack=", body)
        self.assertIn("apiResultName=", body)

    # --- R23-509 ---------------------------------------------------------

    def test_r23_509_coverage_counts_page_positions_and_credits_locations(
        self,
    ):
        """Carried Phase 3 finding: required coverage is counted per
        distinct page position per view. Demanding one dimension per
        physical location is unsatisfiable by construction, because coaxial
        holes collapse to one drawing entity."""
        body = self.body("Public Function VerifyDirectionalCoverage(")
        self.assertIn("coverageFailures=", body)
        self.assertIn("creditedLocations=", body)
        for reason in (
            "DatumUnproven:",
            "UnanchoredBucket:",
            "NoHorizontalScheme",
            "NoVerticalScheme",
            "SmallHoleLocationCredit:",
        ):
            with self.subTest(reason=reason):
                self.assertIn(reason, body)

    def test_expected_location_count_is_measured_not_hardcoded(self):
        """R23-509 expects ten P-0251 small-hole locations. That number is
        derived from the graph so a shortfall shows up as a shortfall
        instead of redefining the target."""
        body = self.body("Public Function CountSmallHoleLocations(")
        self.assertIn("graph.Locations()", body)
        self.assertIn("IsSmallHoleLocation(", body)
        self.assertNotIn("= 10", body)

    # --- hygiene ---------------------------------------------------------

    def test_source_hygiene(self):
        for name in (MODULE, SCHEME_CLS, BUCKET_CLS):
            with self.subTest(name=name):
                raw = (SOURCE / name).read_bytes()
                self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
                self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))
                text = raw.decode("cp1252")
                self.assertTrue(text.startswith("Option Explicit"))
                for line in text.split("\r\n"):
                    self.assertLessEqual(len(line), 79)

    def test_source_revision_identifies_wired_r23_pipeline(self):
        main = self.read("Module1_Main.bas")
        self.assertIn("target-spec-hybrid-v2-2026-08-05-r62", main)


if __name__ == "__main__":
    unittest.main()
