"""Contracts for the automated mutating production run.

The user authorized replacing the manual "Debug > Compile Project" gate and
the per-run chat authorization with automation on 2026-08-04. That gate
existed because VBA compiles lazily: a module that only fails when first
called can abort a run after several views and dozens of dimensions already
exist. These tests hold the replacement to the same job -- a dirty compile
must still be unable to reach a mutating run -- and hold the safety rules
that were never ceremony to begin with.
"""

from pathlib import Path
import unittest


WORKSPACE = Path(__file__).resolve().parents[2]
SOURCE = WORKSPACE / "archive" / "target-spec-hybrid-v2"
RUNNER = WORKSPACE / "tools" / "production-runner" / "Run-R23Production.ps1"
PROBE_RUNNER = WORKSPACE / "tools" / "probe-runner" / "Run-R23Probes.ps1"

# Kept identical to Module1_Main.bas FIXTURE_1..3.
AUTHORIZED_FIXTURES = (
    "P-0251-14A-001.SLDPRT",
    "P-0252-01-001.SLDPRT",
    "P-0252-01-013.SLDPRT",
)


class R23ProductionPreflightContracts(unittest.TestCase):
    """Module20_ProbeRunner.R23_PrepareProductionRun stays read-only."""

    def read(self, name: str) -> str:
        return (SOURCE / name).read_text(encoding="cp1252")

    def preflight(self) -> str:
        return self.read("Module20_ProbeRunner.bas").split(
            "Public Sub R23_PrepareProductionRun()"
        )[1].split("\nEnd Sub")[0]

    def test_preflight_exists_and_is_public(self):
        self.assertIn(
            "Public Sub R23_PrepareProductionRun()",
            self.read("Module20_ProbeRunner.bas"),
        )

    def test_preflight_never_invokes_the_production_entry_point(self):
        """The compile verdict must be readable before anything mutates."""
        body = self.preflight()
        self.assertNotIn("Module1_Main.main", body)
        self.assertNotIn("CreateDrawing", body)
        for mutation in ("allowMutation", ".Save", "InsertModelAnnotations",
                         "CreateSectionViewAt5", "CreateLine"):
            with self.subTest(mutation=mutation):
                self.assertNotIn(mutation, body)

    def test_preflight_derives_its_verdict_from_the_compile_control(self):
        body = self.preflight()
        self.assertIn("R23_CompileProject()", body)
        self.assertIn("R23_PREFLIGHT_COMPILE|verdict=", body)

    def test_preflight_reports_not_ready_on_every_refusal(self):
        """Every exit must be classifiable by the caller from one line."""
        body = self.preflight()
        self.assertIn("ready=True", body)
        for reason in (
            "reason=CompileNotClean",
            "reason=NoOpenAuthorizedPart",
            "reason=PartActivationFailed",
            "reason=UnhandledError",
        ):
            with self.subTest(reason=reason):
                self.assertIn(reason, body)

    def test_preflight_asks_for_the_vbe_dialog_on_a_dirty_compile(self):
        """VBA compile errors are untrappable; no touch can name the module."""
        body = self.preflight()
        self.assertIn("ReadVbeDialogAndHighlightedLine", body)


class R23ProductionRunnerContracts(unittest.TestCase):
    def script(self) -> str:
        return RUNNER.read_text(encoding="utf-8")

    def code(self) -> str:
        return "\n".join(
            line for line in self.script().split("\n")
            if not line.lstrip().startswith("#")
        )

    def test_runner_exists(self):
        self.assertTrue(RUNNER.is_file(), f"missing runner: {RUNNER}")

    def test_mutation_requires_an_explicit_switch_with_no_default(self):
        code = self.code()
        self.assertIn("[Parameter(Mandatory)]", code)
        self.assertIn("[switch] $AllowMutation", code)
        self.assertIn("if (-not $AllowMutation)", code)
        self.assertIn("throw", code.split("if (-not $AllowMutation)")[1][:400])

    def test_main_is_unreachable_unless_the_preflight_is_ready(self):
        """This is the manual compile gate's actual job, kept intact."""
        code = self.code()
        self.assertIn("R23_PrepareProductionRun", code)
        self.assertIn("R23_PREFLIGHT_END", code)
        self.assertIn("ready=True", code)

        # Every path that reaches the main invocation must pass the check.
        before_main = code.split("-ProcedureName 'main'")[0]
        self.assertIn("'ready=True'", before_main)
        self.assertLess(
            before_main.index("R23_PREFLIGHT_END"),
            before_main.index("-ProcedureName 'main'")
            if "-ProcedureName 'main'" in before_main else len(before_main),
        )

    def test_preflight_is_invoked_before_main(self):
        code = self.code()
        self.assertLess(
            code.index("'R23_PrepareProductionRun'"),
            code.index("-ProcedureName 'main'"),
        )

    def test_fixture_allowlist_matches_module1_main(self):
        code = self.code()
        for fixture in AUTHORIZED_FIXTURES:
            with self.subTest(fixture=fixture):
                self.assertIn(fixture, code)
        self.assertIn("Test-AuthorizedFixturePart", code)
        self.assertIn("Refusing an unauthorized fixture part", code)

    def test_part_is_opened_read_only(self):
        """A read-only handle makes saving a fixture impossible."""
        code = self.code()
        self.assertIn("$swOpenDocOptions_ReadOnly = 2", code)
        self.assertIn("-Options $swOpenDocOptions_ReadOnly", code)

    def test_runner_never_saves_or_closes_anything(self):
        code = self.code()
        for forbidden in ("SaveAs", "Save3", "CloseDoc", "QuitDoc",
                          "ExitApp", "Remove-Item", "Set-Content"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, code)

    def test_new_qa_report_is_distinguished_from_a_stale_one(self):
        """A folder from an earlier session must not be read as this run's."""
        code = self.code()
        self.assertIn("$qaBefore", code)
        self.assertIn("$qaAfter", code)
        self.assertIn("$qaAfter.Name -eq $qaBeforeName", code)

    def test_probe_runner_stays_free_of_a_mutating_path(self):
        """Read-only evidence and mutating runs keep separate entry points."""
        probe = PROBE_RUNNER.read_text(encoding="utf-8")
        self.assertNotIn("AllowMutation", probe)
        self.assertNotIn("'main'", probe)
        self.assertNotIn("R23_PrepareProductionRun", probe)


class R23ProductionDocumentationContracts(unittest.TestCase):
    """The operating contract must describe what was and was not relaxed."""

    def read(self, name: str) -> str:
        return (WORKSPACE / name).read_text(encoding="utf-8")

    def test_agents_documents_the_automated_mutating_exception(self):
        agents = self.read("Agents.md")
        self.assertIn("Automated mutating-run exception", agents)
        self.assertIn("Run-R23Production.ps1", agents)
        self.assertIn("-AllowMutation", agents)
        self.assertIn("ready=True", agents)

    def test_agents_still_reserves_visual_acceptance_for_the_user(self):
        agents = self.read("Agents.md")
        exception = agents.split("Automated mutating-run exception")[1].split(
            "\n## "
        )[0]
        self.assertIn("Item 3 is unchanged", exception)
        self.assertIn("read-only", exception)
        self.assertIn(
            "Neither the model nor the generated drawing is ever saved",
            exception,
        )
        self.assertIn("A completed\n  run is not a passed run", exception)

    def test_claude_md_points_at_the_production_runner(self):
        claude = self.read("CLAUDE.md")
        self.assertIn("Run-R23Production.ps1", claude)
        self.assertIn("-AllowMutation", claude)
        self.assertIn("R23_PREFLIGHT_END|ready=True", claude)


if __name__ == "__main__":
    unittest.main()
