"""PostToolUse hook: record that a solidworks-api MCP lookup happened.

Paired with require_api_lookup.py, which refuses source edits that introduce
sw* names when no recent lookup is on record. This half just stamps a marker
file; it never blocks anything and always exits 0, because a failure here must
not interrupt a tool call that already succeeded.
"""

import sys
import time
from pathlib import Path

MARKER = Path(__file__).resolve().parent.parent / "state" / "api-lookup.marker"


def main() -> int:
    try:
        MARKER.parent.mkdir(parents=True, exist_ok=True)
        MARKER.write_text(str(time.time()), encoding="utf-8")
    except Exception:
        # Deliberately silent. A broken marker degrades to the PreToolUse hook
        # asking for a lookup that already happened, which is annoying but
        # safe. Raising here would surface as a hook error on a good call.
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
