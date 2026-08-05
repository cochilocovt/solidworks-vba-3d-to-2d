"""R23-814. Measure what the drawing did, not what the path asked for.

Every run since the section first existed read IView.GetSectionLineInfo2 and
logged only its element count. Both r50 and r51 recorded `values=49`.

r51 moved waypoint 1 by 40 mm - `overshootM=0.040000000|w1Y=0.250324890` -
the sheet showed a visibly longer section line, the crossing predicate
passed on the new geometry, and the resulting section view was
byte-for-byte identical to r50's: same `records=38|entities=38|doubles=1056`
and the same seven X and seven Y coordinates. Our intent was verified on
every one of those runs. The drawing's result never was.

So this decodes the array. Documented layout, SOLIDWORKS 2025 Help:

    [numSectionLines, layer,
      per line: numSegments,
        per segment: lineType, startPt[3], endPt[3],
        arrowStart1[3], arrowEnd1[3], arrowWidth1, arrowHeight1,
        arrowStyle1, arrowStart2[3], arrowEnd2[3], arrowWidth2,
        arrowHeight2, arrowStyle2, textPt1[3], textPt2[3], textHeight]

Three segments under that layout account for 53 doubles and the live array
holds 49, so the tail does not match the documentation on this build. The
segment block sits at the front and is unaffected by the tail, so segments
are decoded, the discrepancy is reported as
`tailMatchesDocumented=False`, and the raw array is dumped alongside -
49 doubles is small, and a dump cannot be wrong about the thing a
structured decode might be wrong about.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MODULE = "Module17_SectionPath.bas"
PIPELINE = "Module2_DrawingPipeline.bas"


class R23SectionLineDecodeContracts(unittest.TestCase):
    def code(self, name: str) -> str:
        text = (SOURCE / name).read_text(encoding="cp1252")
        return "\n".join(
            line for line in text.split("\n")
            if not line.lstrip().startswith("'")
        )

    def decoder(self) -> str:
        return self.code(MODULE).split(
            "Public Sub EmitSectionLineDecode("
        )[1].split("\nEnd Sub")[0]

    # --- it changes nothing ----------------------------------------------

    def test_the_decoder_is_read_only(self):
        body = self.decoder()
        for forbidden in ("CreateLine", "CreateSectionViewAt5", "Select4",
                          "ClearSelection2", "RecordSolidWorksMutation",
                          ".Save", "EditRebuild3", "SetPosition2"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_the_mutation_boundary_is_still_one_procedure(self):
        """Module17's safety boundary: exactly one procedure changes a
        drawing. Adding a reader must not add a second."""
        source = self.code(MODULE)
        self.assertEqual(1, source.count("CreateSectionViewAt5("))
        self.assertEqual(1, source.count("CreateLine("))

    # --- the raw dump -----------------------------------------------------

    def test_the_raw_array_is_dumped(self):
        """A structured decode can be wrong about the layout. A dump of 49
        doubles cannot be, and it is what makes the layout decidable from
        the log rather than from the documentation."""
        body = self.decoder()
        self.assertIn("SECTION_LINE_RAW", body)
        self.assertIn("|from=", body)

    # --- the structured decode -------------------------------------------

    def test_segments_are_decoded_at_the_documented_offsets(self):
        """numSectionLines, layer, numSegments, then seven doubles per
        segment: lineType plus a start and an end point."""
        body = self.decoder()
        self.assertIn("lineCount = CLng(info(lower))", body)
        self.assertIn("layerValue = CDbl(info(lower + 1))", body)
        self.assertIn("segmentCount = CLng(info(lower + 2))", body)
        self.assertIn("base = lower + consumed + (i * 7)", body)
        self.assertIn("SECTION_LINE_SEGMENT", body)
        self.assertIn("|start=", body)
        self.assertIn("|end=", body)

    def test_the_walk_cannot_read_past_the_array(self):
        body = self.decoder()
        self.assertIn("If base + 6 > upper Then Exit For", body)
        self.assertIn("If count < 3 Then", body)
        self.assertIn("segmentCount > 0 And segmentCount < 64", body)

    def test_the_documentation_mismatch_is_reported_not_hidden(self):
        """53 documented doubles against 49 live ones. Reporting the
        comparison is the point; silently decoding the tail anyway would
        invent numbers."""
        body = self.decoder()
        self.assertIn("documentedTotal = 3 + (segmentCount * 7) + 29", body)
        self.assertIn("|tailMatchesDocumented=", body)
        self.assertIn("|segmentsDecoded=", body)
        self.assertIn("|count=", body)

    def test_the_frame_is_not_claimed(self):
        """Which frame these coordinates are in has not been established,
        so the record says AsReturned rather than naming one."""
        body = self.decoder()
        self.assertIn("|frame=AsReturned", body)
        self.assertNotIn("|frame=Page", body)

    def test_the_error_capture_precedes_the_name_helper(self):
        """SafeViewName contains On Error Resume Next, which resets the
        global Err before a concatenation can reach Err.Number."""
        body = self.decoder()
        handler = body.split("Failed:")[1]
        self.assertLess(
            handler.index("decodeErrorNumber = Err.Number"),
            handler.index("SafeViewName(swView)"),
        )

    # --- it is wired in ---------------------------------------------------

    def test_the_pipeline_decodes_the_line_after_layout(self):
        body = self.code(PIPELINE).split(
            "Private Sub RecordSectionLineAfterLayout("
        )[1].split("\nEnd Sub")[0]
        self.assertIn(
            "Module17_SectionPath.EmitSectionLineDecode swView, evidence",
            body,
        )

    def test_only_a_view_that_has_a_line_is_decoded(self):
        """All three views answer sectionLine=Read|values=0 when they carry
        no cut, so the value count is what distinguishes a real section
        line from an empty read."""
        body = self.code(PIPELINE).split(
            "Private Sub RecordSectionLineAfterLayout("
        )[1].split("\nEnd Sub")[0]
        guard = body.split("EmitSectionLineDecode")[0]
        self.assertIn('InStr(1, proof, "|values=0", vbBinaryCompare) = 0',
                      guard)


if __name__ == "__main__":
    unittest.main()
