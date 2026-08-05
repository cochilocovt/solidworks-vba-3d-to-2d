"""R23-815. A jogged section line needs the offset-section option.

Measured in run macro_qa/20260805_050411_P-0251-14A-001. The r52 decode of
IView.GetSectionLineInfo2 showed the drawing holding exactly the requested
path:

    seg1 (-0.102,0.000)->(0.008,0.000)     length 0.110
    seg2 (0.008,0.000)->(0.008,-0.015)     length 0.015
    seg3 (0.008,-0.015)->(0.088,-0.015)    length 0.080

matching the waypoint spacing exactly, with segment 1 running 0.040 past
the bore centre at -0.062 - the r51 overshoot, present in the drawing. Yet
the section view contained the counterbore-column features and no bore. The
bore is at transverse 0.000 and the counterbores at -0.015, so the cut in
use was segment 3's alone.

CreateSectionViewAt5 was being called with Options=0. MCP corpus,
swCreateSectionViewAtOptions_e: swCreateSectionView_OffsetSection = 2,
"If set, then an aligned section view is created (two lines at an angle);
if not set, a normal projection section view is created." So SOLIDWORKS was
building a normal projection section from a jogged line and cutting at one
offset - which is also why r51 could lengthen segment 1 by 40 mm and
produce a byte-for-byte identical section view.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module17_SectionPath.bas"


class R23OffsetSectionOptionContracts(unittest.TestCase):
    def code(self) -> str:
        text = (SOURCE / MODULE).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def creator(self) -> str:
        return self.code().split(
            "Public Function CreateSectionFromPath("
        )[1].split("\nEnd Function")[0]

    def test_the_option_constant_is_the_corpus_value(self):
        self.assertIn("Private Const SECTION_OPTION_OFFSET As Long = 2",
                      self.code())

    def test_the_option_is_passed_to_the_creator(self):
        body = self.creator()
        self.assertIn("sectionOptions = SECTION_OPTION_OFFSET", body)
        self.assertIn(
            "placeX, placeY, 0#, path.SectionLabel, sectionOptions, "
            "Nothing, 0#",
            body,
        )

    def test_the_literal_zero_option_is_gone(self):
        """Options=0 built a normal projection section from a jogged line
        and silently used one segment of three."""
        body = self.creator()
        self.assertNotIn("path.SectionLabel, 0, Nothing, 0#", body)

    def test_the_option_is_recorded_in_evidence(self):
        """Nothing else in the report would show which option was used, and
        it decides what the cut contains."""
        body = self.creator()
        self.assertIn("SECTION_CREATE_OPTIONS", body)
        self.assertIn("|options=", body)
        self.assertIn("|offsetSection=True", body)

    def test_the_constant_is_declared_in_the_declarations_section(self):
        """VBA refuses a module-level Const between procedures: "Only
        comments may appear after End Sub, End Function, or End Property".
        That cost a deploy cycle at r45."""
        source = self.code()
        declaration = source.index("Private Const SECTION_OPTION_OFFSET")
        first_procedure = min(
            source.index("\nPrivate Sub "),
            source.index("\nPublic Function "),
        )
        self.assertLess(declaration, first_procedure)

    def test_the_mutation_boundary_is_unchanged(self):
        """Still exactly one procedure that can change a drawing, and it
        still refuses without explicit authorization and an unresolved
        path."""
        source = self.code()
        self.assertEqual(1, source.count("CreateSectionViewAt5("))
        body = self.creator()
        self.assertIn("If Not allowMutation Then", body)
        self.assertIn("If Not path.Resolved Then", body)
        self.assertLess(
            body.index("If Not allowMutation Then"),
            body.index("CreateSectionViewAt5("),
        )

    def test_no_other_option_bit_was_added_silently(self):
        """Partial, DisplaySurfaceCut, ChangeDirection and ScaleWithModel
        each change what the section shows. Only the offset bit is
        intended, and any other would need its own evidence."""
        source = self.code()
        for forbidden in ("SECTION_OPTION_PARTIAL",
                          "SECTION_OPTION_SURFACE",
                          "SECTION_OPTION_CHANGE_DIRECTION",
                          "SECTION_OPTION_SCALE"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, source)

        body = self.creator()
        self.assertEqual(1, body.count("sectionOptions = "))


if __name__ == "__main__":
    unittest.main()
