"""Record CURRENT_STATUS.md at the start of each Claude Code turn."""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import time
from pathlib import Path


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


def main() -> int:
    try:
        hook_input = json.load(sys.stdin)
        root = project_root()
        status_file = root / "docs" / "CURRENT_STATUS.md"
        if not status_file.is_file():
            return 0

        session_id = safe_session_id(hook_input.get("session_id"))
        write_state(
            state_path(root, session_id),
            {
                "baseline_sha256": sha256_file(status_file),
                "recorded_at": int(time.time()),
                "attempts": 0,
            },
        )
    except (OSError, ValueError, TypeError):
        # Turn-start bookkeeping must never prevent Claude from receiving a prompt.
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
