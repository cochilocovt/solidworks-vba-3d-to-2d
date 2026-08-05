"""Source contracts for the R23 probe-automation tool (PA-100..109).

Static checks only: the evidence sink, the probe runner, and the nine
read-only R23_Probe* modules never gained a raw Debug.Print, a mutating
call, or an allowMutation escape hatch, and every deployed standard module
carries the R23_CompileTouch marker the runner needs to localise a
compile failure. See docs/R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md.
"""

import json
from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
MANIFEST = WORKSPACE / "tools" / "swp-deploy" / "deployment-manifest.json"

SINK_MODULE = "Module21_EvidenceSink.bas"
RUNNER_MODULE = "Module20_ProbeRunner.bas"
RUN_EVIDENCE_CLASS = "CRunEvidence.cls"

# Module, probe entry point, in the dependency order the plan's section 4
# table and Module20_ProbeRunner.R23_RunAllProbes both use.
PROBE_ENTRY_POINTS = [
    ("Module12_FeatureQualification.bas", "R23_ProbeFeatureCatalog"),
    ("Module13_ProjectionResolution.bas", "R23_ProbeViewProjections"),
    ("Module14_AnnotationImport.bas", "R23_ProbeAnnotationReconciliation"),
    ("Module15_OrdinateScheme.bas", "R23_ProbeOrdinateScheme"),
    ("Module16_CalloutDefinition.bas", "R23_ProbeCalloutDefinition"),
    ("Module17_SectionPath.bas", "R23_ProbeSectionPath"),
    ("Module10_SectionDimensionEngine.bas", "R23_ProbeSectionDimensions"),
    ("Module18_ContentEnvelope.bas", "R23_ProbeContentEnvelope"),
    ("Module19_SemanticQA.bas", "R23_ProbeSemanticQA"),
]

MUTATING_MARKERS = (
    "allowMutation",
    ".Save ",
    ".SaveAs",
    "ModifyDefinition",
    "EmitRunEvidence",
    "CreateSectionViewAt5",
)


def read_source(name: str) -> str:
    return (SOURCE / name).read_text(encoding="cp1252")


def executable_lines(text: str):
    for line in text.split("\n"):
        if not line.lstrip().startswith("'"):
            yield line


class EvidenceSinkContracts(unittest.TestCase):
    def setUp(self):
        self.source = read_source(SINK_MODULE)

    def test_component_is_managed(self):
        manifest = MANIFEST.read_text(encoding="utf-8")
        self.assertIn("Module21_EvidenceSink", manifest)

    def test_public_surface(self):
        for member in ("Function OpenLog(", "Sub LogLine(", "Sub CloseLog(",
                        "Function IsOpen("):
            with self.subTest(member=member):
                self.assertIn(member, self.source)

    def test_log_line_always_debug_prints(self):
        """LogLine must degrade to Debug.Print alone when no log is open,
        per PA-100's acceptance line."""
        body = self.source.split("Public Sub LogLine(")[1].split(
            "\nEnd Sub")[0]
        self.assertIn("Debug.Print text", body)
        # The Debug.Print call must sit outside the "If mLogOpen" guard,
        # not inside it, or the degrade-when-closed behaviour breaks.
        guard_index = body.index("If mLogOpen Then")
        end_if_index = body.index("End If", guard_index)
        debug_print_index = body.index("Debug.Print text")
        self.assertGreater(debug_print_index, end_if_index)

    def test_open_log_derives_root_from_test_assets_marker(self):
        body = self.source.split("Public Function OpenLog(")[1].split(
            "\nEnd Function")[0]
        self.assertIn("\\test_assets\\models\\", body)
        self.assertIn("probe_runs", body)


class CRunEvidenceRoutesThroughSinkContracts(unittest.TestCase):
    def setUp(self):
        self.source = read_source(RUN_EVIDENCE_CLASS)

    def test_no_raw_debug_print_in_add_methods(self):
        for method, prefix in (
            ("AddInfo", "QA INFO: "),
            ("AddWarning", "QA WARNING: "),
            ("AddFailure", "QA FAILURE: "),
        ):
            with self.subTest(method=method):
                body = self.source.split(
                    f"Public Sub {method}(ByVal message As String)"
                )[1].split("\nEnd Sub")[0]
                self.assertNotIn("Debug.Print", body)
                self.assertIn("Module21_EvidenceSink.LogLine", body)
                self.assertIn(prefix, body)


class NoRawDebugPrintInProbesContracts(unittest.TestCase):
    def test_no_raw_debug_print_in_any_probe_module(self):
        for filename, _ in PROBE_ENTRY_POINTS:
            with self.subTest(module=filename):
                source = read_source(filename)
                for line in executable_lines(source):
                    self.assertNotIn(
                        "Debug.Print", line,
                        msg=f"raw Debug.Print survives in {filename}: "
                        f"{line.strip()!r}",
                    )

    def test_probe_modules_route_through_the_sink(self):
        for filename, _ in PROBE_ENTRY_POINTS:
            with self.subTest(module=filename):
                self.assertIn(
                    "Module21_EvidenceSink.LogLine", read_source(filename))


class ProbeRunnerContracts(unittest.TestCase):
    def setUp(self):
        self.source = read_source(RUNNER_MODULE)
        self.executable = "\n".join(executable_lines(self.source))

    def test_component_is_managed(self):
        manifest = MANIFEST.read_text(encoding="utf-8")
        self.assertIn("Module20_ProbeRunner", manifest)

    def test_calls_all_nine_probes_in_documented_order(self):
        body = self.source.split("Public Sub R23_RunAllProbes()")[1]
        positions = []
        for module_file, procedure in PROBE_ENTRY_POINTS:
            module_name = module_file.split(".")[0]
            call = f"{module_name}.{procedure}"
            with self.subTest(call=call):
                self.assertIn(call, body)
            positions.append(body.index(call))
        self.assertEqual(
            positions, sorted(positions),
            msg="probe calls are not in the documented dependency order",
        )

    def test_stops_before_probes_when_compile_is_not_clean(self):
        body = self.source.split("Public Sub R23_RunAllProbes()")[1]
        self.assertIn("CompileNotClean", body)
        first_probe_call = body.index(
            PROBE_ENTRY_POINTS[0][0].split(".")[0]
            + "." + PROBE_ENTRY_POINTS[0][1]
        )
        self.assertLess(body.index("CompileNotClean"), first_probe_call)

    def test_compile_failure_requests_vbe_evidence_not_touch_localization(self):
        body = self.source.split("Public Sub R23_RunAllProbes()")[1]
        failed = body.split("If StrComp(compileVerdict")[1]
        failed = failed.split("End If", 1)[0]
        self.assertIn("ReadVbeDialogAndHighlightedLine", failed)
        self.assertIn("UnavailableForCompileErrors", failed)
        self.assertNotIn("R23_TouchAllModules", failed)

    def test_active_drawing_is_preferred_and_activation_is_verified(self):
        """The runner opens the requested drawing, so a pre-existing
        unrelated drawing must not win merely because document enumeration
        reaches it first."""
        find = self.source.split("Private Function FindOpenDrawing(")[1]
        self.assertIn("Set activeDocument = swApp.ActiveDoc", find)
        self.assertLess(
            find.index("Set activeDocument = swApp.ActiveDoc"),
            find.index("Set doc = swApp.GetFirstDocument"),
        )

        activate = self.source.split(
            "Private Function ActivateDocumentByTitle("
        )[1]
        self.assertIn("activeTitle=", activate)
        self.assertIn("StrComp(activeTitle, documentTitle", activate)

    def test_no_mutating_marker_anywhere_in_the_runner(self):
        for marker in MUTATING_MARKERS:
            with self.subTest(marker=marker):
                self.assertNotIn(marker, self.executable)

    def test_never_calls_the_production_qa_gate(self):
        self.assertNotIn("EmitRunEvidence", self.executable)

    def test_compile_control_is_resolved_by_caption_not_a_literal_id(self):
        """PA-104: the VBIDE control ID is outside the SOLIDWORKS API
        corpus, so it must be discovered by caption at run time, never
        hardcoded as a bare FindControl(id) call."""
        self.assertNotIn("FindControl(", self.executable)
        self.assertIn("compile", self.source.lower())

    def test_caption_match_uses_the_ampersand_stripped_caption(self):
        """Live finding 2026-08-04: VBIDE CommandBarControl.Caption
        includes the raw accelerator-key "&" (e.g. "Compi&le Fable"),
        so InStr against the raw caption never matches. WalkVbeControls
        must compare against CleanControlText's output, not ctrlCaption
        directly, or the Compile control silently stops resolving."""
        body = self.source.split("Private Sub WalkVbeControls(")[1].split(
            "\nEnd Sub")[0]
        self.assertIn("cleanedCaption = CleanControlText(ctrlCaption)", body)
        match_line = [
            line for line in body.split("\n") if "InStr(1, " in line
        ]
        self.assertTrue(match_line, "no InStr caption match found")
        self.assertIn("cleanedCaption", match_line[0])
        self.assertNotIn("InStr(1, ctrlCaption,", body)


class CompileTouchContracts(unittest.TestCase):
    def test_every_standard_module_has_compile_touch(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        std_modules = [
            component["file"]
            for component in manifest["components"]
            if component["kind"] == "StdModule"
        ]
        self.assertGreaterEqual(len(std_modules), 21)

        for filename in std_modules:
            with self.subTest(module=filename):
                source = read_source(filename)
                self.assertIn("Public Sub R23_CompileTouch()", source)

    def test_runner_touches_every_standard_module(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        std_modules = [
            component["name"]
            for component in manifest["components"]
            if component["kind"] == "StdModule"
        ]
        body = read_source(RUNNER_MODULE).split(
            "Public Function R23_TouchAllModules()")[1]
        for module_name in std_modules:
            with self.subTest(module=module_name):
                self.assertIn(f"{module_name}.R23_CompileTouch", body)


class SourceHygieneContracts(unittest.TestCase):
    def test_new_modules_are_ansi_crlf_no_bom_no_attribute(self):
        for filename in (SINK_MODULE, RUNNER_MODULE):
            with self.subTest(module=filename):
                raw = (SOURCE / filename).read_bytes()
                self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))

                text = raw.decode("cp1252")
                self.assertTrue(text.startswith("Option Explicit"))
                self.assertNotIn("Attribute ", text)

                stripped_of_crlf = text.replace("\r\n", "")
                self.assertNotIn(
                    "\n", stripped_of_crlf,
                    msg=f"{filename} has a bare LF not part of a CRLF pair",
                )

                for line_number, line in enumerate(
                        text.split("\r\n"), start=1):
                    self.assertLessEqual(
                        len(line), 79,
                        msg=f"{filename}:{line_number} exceeds 79 columns",
                    )


if __name__ == "__main__":
    unittest.main()
