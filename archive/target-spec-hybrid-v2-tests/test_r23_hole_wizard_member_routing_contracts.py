"""Hole Wizard nominal size comes from the type-specific member.

Measured in run macro_qa/20260805_001154_P-0251-14A-001, via the read-only
HOLEWIZARD_MEMBER_PROBE dump of 22 candidate members:

    CBORE for M6 Socket Head Cap Screw1  Type=14 swCounterBoreThru
        HoleDiameter=0  Diameter=0  Depth=0
        ThruHoleDiameter=0.0066  ThruHoleDepth=0.018
        CounterBoreDiameter=0.011  CounterBoreDepth=0.006

    M5x0.8 Tapped Hole1                  Type=46 swTapBlindCosmeticThread
        HoleDiameter=0  Diameter=0  Depth=0
        ThreadDiameter=0.005  ThreadDepth=0.010
        TapDrillDiameter=0.0042  TapDrillDepth=0.0124

accessGranted=True on both, and CounterBore* read correctly, so
AccessSelections was never the problem. IWizardHoleFeatureData2.HoleDiameter
simply does not apply to a standard-driven Hole Wizard feature. 6.60 THRU and
11.00 by 6.00 are exactly what the reference callout shows.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module12_FeatureQualification.bas"


class R23HoleWizardMemberRoutingContracts(unittest.TestCase):
    def code(self) -> str:
        text = (SOURCE / MODULE).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def read_body(self) -> str:
        return self.code().split("Private Sub ReadHoleWizardDefinition(")[
            1
        ].split("\nEnd Sub")[0]

    def test_nominal_no_longer_comes_from_holediameter_alone(self):
        body = self.read_body()
        self.assertNotIn(
            "definition.NominalDiameterM = holeData.HoleDiameter", body)
        self.assertNotIn("definition.DepthM = holeData.HoleDepth", body)

    def test_diameter_chain_covers_both_observed_types(self):
        body = self.read_body()
        chain = body.split("FirstNonZeroHoleMember(holeData, _")[1].split(
            "diameterRoute)"
        )[0]
        # Counterbore-thru and cosmetic-thread tap are disjoint on these two.
        self.assertIn("ThruHoleDiameter", chain)
        self.assertIn("ThreadDiameter", chain)
        self.assertLess(
            chain.index("ThruHoleDiameter"), chain.index("ThreadDiameter"))

    def test_depth_chain_covers_both_observed_types(self):
        body = self.read_body()
        self.assertIn("ThruHoleDepth,ThreadDepth,TapDrillDepth", body)

    def test_supplying_member_is_named_in_the_proof(self):
        """A future hole type landing on a different member must be visible
        in evidence rather than silently accepted."""
        body = self.read_body()
        self.assertIn(
            'definition.DiameterProofSource = _\n        '
            '"IWizardHoleFeatureData2." & diameterRoute', body)
        self.assertIn(
            'definition.DepthProofSource = "IWizardHoleFeatureData2." '
            '& depthRoute', body)

    def test_all_zero_read_is_not_reported_as_success(self):
        body = self.read_body()
        self.assertIn('"ReadAllZeroValues:"', body)
        self.assertIn('definition.DefinitionReadStatus = "Read"', body)

    def test_helper_guards_each_member_read_individually(self):
        """Several members are documented as relevant only for particular
        swWzdHoleTypes_e values and may raise rather than return zero."""
        helper = self.code().split("Private Function FirstNonZeroHoleMember(")[
            1
        ].split("\nEnd Function")[0]
        self.assertIn("On Error Resume Next", helper)
        self.assertIn("CallByName(holeData, memberName, VbGet)", helper)
        self.assertIn("AllCandidatesZero:", helper)

    def test_member_probe_is_retained_as_evidence(self):
        body = self.code()
        self.assertIn("HOLEWIZARD_MEMBER_PROBE", body)
        self.assertIn("HOLE_WIZARD_PROBE_MEMBERS", body)


if __name__ == "__main__":
    unittest.main()
