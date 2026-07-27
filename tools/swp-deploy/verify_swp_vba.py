from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any
import warnings


SCRIPT_DIR = Path(__file__).resolve().parent
WORKSPACE = SCRIPT_DIR.parents[1]
LOCAL_OLETOOLS = WORKSPACE / ".codex-tools" / "oletools"
if LOCAL_OLETOOLS.is_dir():
    sys.path.insert(0, str(LOCAL_OLETOOLS))

warnings.filterwarnings("ignore", category=DeprecationWarning, module=r".*olevba")

try:
    from oletools.olevba import VBA_Parser
except ImportError as exc:  # pragma: no cover - environment diagnostic
    raise SystemExit(
        "oletools is unavailable. Expected the project-local package at "
        f"{LOCAL_OLETOOLS}: {exc}"
    ) from exc


def normalized_text(value: str) -> str:
    value = value.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
    normalized_lines: list[str] = []
    for raw_line in value.split("\n"):
        line = raw_line.rstrip()
        if re.match(r"^\s*(?:ï»¿)?Attribute\s+", line, re.IGNORECASE):
            continue
        output: list[str] = []
        in_string = False
        index = 0
        while index < len(line):
            character = line[index]
            if character == '"':
                output.append(character)
                if in_string and index + 1 < len(line) and line[index + 1] == '"':
                    output.append('"')
                    index += 2
                    continue
                in_string = not in_string
            else:
                output.append(character if in_string else character.casefold())
            index += 1
        normalized_lines.append("".join(output))
    return "\n".join(normalized_lines).rstrip() + "\n"


def digest(value: str) -> str:
    return hashlib.sha256(normalized_text(value).encode("utf-8")).hexdigest()


def resolve_manifest_path(manifest_path: Path, configured: str) -> Path:
    return (manifest_path.parent / configured).resolve()


def load_embedded_components(swp_path: Path) -> dict[str, dict[str, str]]:
    parser = VBA_Parser(str(swp_path))
    components: dict[str, dict[str, str]] = {}
    try:
        for _, stream_path, filename, code in parser.extract_macros():
            if not filename or str(filename).startswith("VBA_P-code"):
                continue
            if isinstance(code, bytes):
                code = code.decode("cp1252", errors="replace")
            name = Path(str(filename)).stem
            components[name.casefold()] = {
                "name": name,
                "filename": str(filename),
                "stream": str(stream_path),
                "code": str(code),
            }
    finally:
        parser.close()
    return components


def main() -> int:
    argument_parser = argparse.ArgumentParser(
        description="Inventory and verify source embedded in a SOLIDWORKS .swp file."
    )
    argument_parser.add_argument("--manifest", type=Path, required=True)
    argument_parser.add_argument("--swp", type=Path)
    argument_parser.add_argument("--inventory-only", action="store_true")
    argument_parser.add_argument("--json-output", type=Path)
    args = argument_parser.parse_args()

    manifest_path = args.manifest.resolve()
    manifest: dict[str, Any] = json.loads(manifest_path.read_text(encoding="utf-8"))
    swp_path = (
        args.swp.resolve()
        if args.swp
        else resolve_manifest_path(manifest_path, manifest["targetSwp"])
    )
    source_directory = resolve_manifest_path(
        manifest_path, manifest["sourceDirectory"]
    )

    if not swp_path.is_file():
        raise SystemExit(f"SWP file does not exist: {swp_path}")

    embedded = load_embedded_components(swp_path)
    result: dict[str, Any] = {
        "swp": str(swp_path),
        "sourceDirectory": str(source_directory),
        "embeddedComponentCount": len(embedded),
        "embeddedComponents": sorted(item["name"] for item in embedded.values()),
        "checks": [],
        "revision": None,
        "success": True,
    }

    if not args.inventory_only:
        for component in manifest["components"]:
            source_path = source_directory / component["file"]
            check: dict[str, Any] = {
                "name": component["name"],
                "source": str(source_path),
                "present": False,
                "matches": False,
            }

            if not source_path.is_file():
                check["error"] = "source file is missing"
                result["success"] = False
                result["checks"].append(check)
                continue

            embedded_component = embedded.get(component["name"].casefold())
            if embedded_component is None:
                check["error"] = "component is missing from the SWP"
                result["success"] = False
                result["checks"].append(check)
                continue

            source_text = source_path.read_text(encoding="utf-8-sig")
            check["present"] = True
            check["sourceSha256"] = digest(source_text)
            check["embeddedSha256"] = digest(embedded_component["code"])
            check["matches"] = check["sourceSha256"] == check["embeddedSha256"]
            if not check["matches"]:
                check["error"] = "embedded source differs"
                result["success"] = False
            result["checks"].append(check)

        revision_source = source_directory / manifest["revisionSourceFile"]
        revision_match = re.search(
            manifest["revisionPattern"],
            revision_source.read_text(encoding="utf-8-sig"),
        )
        expected_revision = revision_match.group(0) if revision_match else None
        main_component = embedded.get("module1_main")
        embedded_revision = None
        if main_component is not None:
            embedded_match = re.search(
                manifest["revisionPattern"], main_component["code"]
            )
            embedded_revision = embedded_match.group(0) if embedded_match else None

        result["revision"] = {
            "expected": expected_revision,
            "embedded": embedded_revision,
            "matches": expected_revision == embedded_revision and expected_revision is not None,
        }
        if not result["revision"]["matches"]:
            result["success"] = False

    if args.json_output:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(
            json.dumps(result, indent=2, sort_keys=True), encoding="utf-8"
        )

    print(f"SWP: {swp_path}")
    print(f"Embedded components: {result['embeddedComponentCount']}")
    if args.inventory_only:
        for component_name in result["embeddedComponents"]:
            print(f"  {component_name}")
        return 0

    matching = sum(1 for check in result["checks"] if check["matches"])
    print(f"Managed components matching source: {matching}/{len(result['checks'])}")
    if result["revision"]:
        print(
            "Revision: embedded={embedded}; expected={expected}".format(
                **result["revision"]
            )
        )
    for check in result["checks"]:
        if not check["matches"]:
            print(f"  MISMATCH {check['name']}: {check.get('error', 'unknown')}")

    print("VERIFY: PASS" if result["success"] else "VERIFY: FAIL")
    return 0 if result["success"] else 3


if __name__ == "__main__":
    raise SystemExit(main())
