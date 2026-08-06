"""PreToolUse hook: refuse VBA edits that touch sw* names without a lookup.

Why this exists
---------------
CLAUDE.md instructs that the solidworks-api MCP and
skills/solidworks-api-lookup/SKILL.md be consulted before touching any sw*
constant or API call. That instruction did not hold in practice: it is read
once at session start and loses to momentum a hundred tool calls into a
debugging session. The result on 2026-08-06 was three display-mode constants
compiled into the trunk with values belonging to entirely different enum
members, surviving fifteen live runs because a wrong display mode renders a
plausible view instead of raising.

So this fires at the moment of the edit rather than relying on memory.

The rule
--------
Block Edit/Write to a managed VBA source file when the incoming content
introduces a token matching sw[A-Z]... - a SOLIDWORKS constant, enum member, or
API member reference - and no MCP lookup has been recorded within
LOOKUP_TTL_SECONDS.

Edits that introduce no such token (comments, control flow, renames, report
strings) pass untouched, so this does not tax ordinary work.

Deliberate limits, stated so nobody mistakes this for proof:
  * It cannot tell whether the lookup was RELEVANT to the constant being
    written, only that one happened recently.
  * It cannot tell whether the result was read carefully.
  * It is a prompt to check, not evidence that a value is correct. Only the
    installed SW2025 type library settles that.
The companion test test_api_constant_provenance.py covers the other side:
whether the finding actually reached docs/SOLIDWORKS_API_VALIDATION.md.
"""

import json
import re
import sys
import time
from pathlib import Path

STATE = Path(__file__).resolve().parent.parent / "state"
MARKER = STATE / "api-lookup.marker"

LOOKUP_TTL_SECONDS = 30 * 60

WATCHED_SUFFIXES = (".bas", ".cls")
WATCHED_DIRS = ("src/", "src\\", "tools/swp-deploy", "tools\\swp-deploy")

# sw followed by an upper-case letter: swInsertDimensions, swDisplayMode_e,
# swCreateOrdDimErr_Success. Locally declared object variables in this codebase
# are swView / swModel / swDraw and also match, which is intentional - those
# lines are exactly where API calls are made.
SW_TOKEN = re.compile(r"\bsw[A-Z][A-Za-z0-9_]*")


def _is_watched(path_text: str) -> bool:
    if not path_text:
        return False
    lowered = path_text.replace("\\", "/").lower()
    if not lowered.endswith(WATCHED_SUFFIXES):
        return False
    return any(d.replace("\\", "/") in lowered for d in WATCHED_DIRS)


def _incoming_text(tool_input: dict) -> str:
    parts = []
    for key in ("content", "new_string"):
        value = tool_input.get(key)
        if isinstance(value, str):
            parts.append(value)
    edits = tool_input.get("edits")
    if isinstance(edits, list):
        for edit in edits:
            if isinstance(edit, dict) and isinstance(
                edit.get("new_string"), str
            ):
                parts.append(edit["new_string"])
    return "\n".join(parts)


def _lookup_age_seconds():
    try:
        return time.time() - MARKER.stat().st_mtime
    except OSError:
        return None


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return 0

    tool_name = payload.get("tool_name", "")
    if tool_name not in ("Edit", "Write", "MultiEdit", "NotebookEdit"):
        return 0

    tool_input = payload.get("tool_input") or {}
    if not _is_watched(str(tool_input.get("file_path", ""))):
        return 0

    text = _incoming_text(tool_input)
    tokens = sorted(set(SW_TOKEN.findall(text)))
    if not tokens:
        return 0

    age = _lookup_age_seconds()
    if age is not None and age <= LOOKUP_TTL_SECONDS:
        return 0

    shown = ", ".join(tokens[:8])
    if len(tokens) > 8:
        shown += f", and {len(tokens) - 8} more"

    when = "never in this repo" if age is None else f"{int(age / 60)} min ago"

    sys.stderr.write(
        "BLOCKED: this edit introduces SOLIDWORKS API tokens "
        f"({shown}) and the last solidworks-api MCP lookup was {when}.\n"
        "\n"
        "A wrong sw* value does not raise in this macro - it renders a "
        "plausible drawing that is wrong. Three display-mode constants "
        "survived fifteen live runs that way.\n"
        "\n"
        "Before retrying:\n"
        "  1. Invoke the solidworks-api-lookup skill "
        "(skills/solidworks-api-lookup/SKILL.md).\n"
        "  2. Query the MCP for the member AND the enum it consumes - "
        "solidworks_lookup_method plus solidworks_get_enum_values. The "
        "method Remarks override the enum table.\n"
        "  3. Check docs/CODESTACK_DRAWING_API_COVERAGE.md for a tested "
        "pattern covering this operation, and note where it says the corpus "
        "has none.\n"
        "  4. Record anything material in "
        "docs/SOLIDWORKS_API_VALIDATION.md.\n"
    )
    return 2


if __name__ == "__main__":
    sys.exit(main())
