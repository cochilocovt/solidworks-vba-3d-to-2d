"""Require an evidence-honest CURRENT_STATUS.md update before Claude stops."""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
from pathlib import Path


UPDATE_INSTRUCTION = (
    "Update docs/CURRENT_STATUS.md before ending this turn. Add a concise, "
    "newest-first entry that states what changed, the evidence obtained, files "
    "changed, and every verification gate not completed. If this was read-only "
    "or made no material project change, record a dated 'No material status "
    "change' note instead of inventing progress. Do not rewrite prior evidence, "
    "claim static/compile/execution/visual gates that were not run, or bump the "
    "macro revision unless deployable behavior changed."
)


def project_root() -> Path:
    configured = os.environ.get("CLAUDE_PROJECT_DIR")
    if configured:
        return Path(configured).resolve()
    return Path(__file__).resolve().parents[2]


def safe_session_id(value: object) -> str:
    session_id = str(value or "unknown")
    return re.sub(r"[^A-Za-z0-9._-]", "_", session_id)[:200] or "unknown"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def state_path(root: Path, session_id: str) -> Path:
    return root / ".claude" / "state" / "current-status" / f"{session_id}.json"


def write_state(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, sort_keys=True), encoding="utf-8")
    temporary.replace(path)


def emit(payload: dict[str, str]) -> None:
    sys.stdout.write(json.dumps(payload, separators=(",", ":")))


def block(reason: str) -> None:
    emit({"decision": "block", "reason": reason})


def warn_and_allow(message: str) -> None:
    emit({"systemMessage": message})


def remove_state(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def main() -> int:
    try:
        hook_input = json.load(sys.stdin)
    except (ValueError, TypeError):
        return 0

    root = project_root()
    status_file = root / "docs" / "CURRENT_STATUS.md"
    session_id = safe_session_id(hook_input.get("session_id"))
    turn_state = state_path(root, session_id)
    stop_hook_active = hook_input.get("stop_hook_active") is True

    if not status_file.is_file():
        if stop_hook_active:
            warn_and_allow(
                "CURRENT_STATUS turn gate could not find docs/CURRENT_STATUS.md; "
                "allowing stop after one retry to prevent a hook loop."
            )
        else:
            block("docs/CURRENT_STATUS.md is missing. Restore it, then record this turn's status.")
        return 0

    try:
        current_hash = sha256_file(status_file)
    except OSError as error:
        if stop_hook_active:
            warn_and_allow(
                f"CURRENT_STATUS turn gate could not read the status file ({error}); "
                "allowing stop after one retry to prevent a hook loop."
            )
        else:
            block(f"Could not read docs/CURRENT_STATUS.md: {error}")
        return 0

    try:
        saved_state = json.loads(turn_state.read_text(encoding="utf-8"))
        baseline_hash = str(saved_state["baseline_sha256"])
    except (OSError, ValueError, TypeError, KeyError):
        if stop_hook_active:
            remove_state(turn_state)
            warn_and_allow(
                "CURRENT_STATUS turn gate has no valid turn-start baseline; "
                "allowing stop after one retry to prevent a hook loop."
            )
            return 0
        try:
            write_state(turn_state, {"baseline_sha256": current_hash, "attempts": 1})
        except OSError:
            pass
        block("No valid turn-start CURRENT_STATUS baseline was available. " + UPDATE_INSTRUCTION)
        return 0

    if current_hash != baseline_hash:
        remove_state(turn_state)
        return 0

    if stop_hook_active:
        remove_state(turn_state)
        warn_and_allow(
            "docs/CURRENT_STATUS.md is still unchanged after the Stop-hook retry. "
            "The turn is being allowed to end to prevent an infinite hook loop."
        )
        return 0

    try:
        saved_state["attempts"] = int(saved_state.get("attempts", 0)) + 1
        write_state(turn_state, saved_state)
    except (OSError, TypeError, ValueError):
        pass
    block(UPDATE_INSTRUCTION)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
