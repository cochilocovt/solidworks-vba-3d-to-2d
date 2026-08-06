"""Regression tests for the per-turn CURRENT_STATUS Claude Code hooks."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


HOOKS_DIR = Path(__file__).resolve().parent
BASELINE_HOOK = HOOKS_DIR / "record_current_status_baseline.py"
STOP_HOOK = HOOKS_DIR / "require_current_status_update.py"
SETTINGS = HOOKS_DIR.parent / "settings.json"


class CurrentStatusHookTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        (self.root / "docs").mkdir()
        self.status = self.root / "docs" / "CURRENT_STATUS.md"
        self.status.write_text("# Current status\n", encoding="utf-8")
        self.environment = os.environ.copy()
        self.environment["CLAUDE_PROJECT_DIR"] = str(self.root)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def run_hook(self, script: Path, **payload: object) -> subprocess.CompletedProcess[str]:
        defaults: dict[str, object] = {
            "session_id": "test/session",
            "hook_event_name": "Stop",
            "stop_hook_active": False,
        }
        defaults.update(payload)
        return subprocess.run(
            [sys.executable, str(script)],
            input=json.dumps(defaults),
            capture_output=True,
            text=True,
            check=False,
            env=self.environment,
        )

    def record_baseline(self) -> None:
        result = self.run_hook(
            BASELINE_HOOK,
            hook_event_name="UserPromptSubmit",
            stop_hook_active=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, "")

    def test_unchanged_status_blocks_first_stop(self) -> None:
        self.record_baseline()
        result = self.run_hook(STOP_HOOK)
        output = json.loads(result.stdout)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(output["decision"], "block")
        self.assertIn("No material status change", output["reason"])

    def test_changed_status_allows_stop_and_clears_state(self) -> None:
        self.record_baseline()
        self.status.write_text("# Current status\n\nUpdated.\n", encoding="utf-8")
        result = self.run_hook(STOP_HOOK)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, "")
        state_files = list((self.root / ".claude" / "state").rglob("*.json"))
        self.assertEqual(state_files, [])

    def test_unchanged_status_allows_second_stop_with_warning(self) -> None:
        self.record_baseline()
        first = self.run_hook(STOP_HOOK)
        self.assertEqual(json.loads(first.stdout)["decision"], "block")
        second = self.run_hook(STOP_HOOK, stop_hook_active=True)
        output = json.loads(second.stdout)
        self.assertNotIn("decision", output)
        self.assertIn("infinite hook loop", output["systemMessage"])

    def test_missing_status_blocks_first_stop(self) -> None:
        self.status.unlink()
        result = self.run_hook(STOP_HOOK)
        output = json.loads(result.stdout)
        self.assertEqual(output["decision"], "block")
        self.assertIn("missing", output["reason"])

    def test_settings_register_both_turn_hooks(self) -> None:
        settings = json.loads(SETTINGS.read_text(encoding="utf-8"))
        hooks = settings["hooks"]
        self.assertIn("record_current_status_baseline.py", hooks["UserPromptSubmit"][0]["hooks"][0]["command"])
        self.assertIn("require_current_status_update.py", hooks["Stop"][0]["hooks"][0]["command"])


if __name__ == "__main__":
    unittest.main()
