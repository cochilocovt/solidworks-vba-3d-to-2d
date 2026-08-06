#!/usr/bin/env python3
"""Compact a generated Graphify text corpus into bounded provenance bundles."""
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path


MAX_CHARS = 240_000


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--clean", action="store_true")
    args = parser.parse_args()
    source = args.source_root.resolve()
    output = args.output_root.resolve()
    if output == source or source in output.parents:
        raise SystemExit("--output-root must be separate from --source-root")
    if args.clean and output.exists():
        shutil.rmtree(output)
    output.mkdir(parents=True, exist_ok=True)

    mapping: dict[str, str] = {}
    for category in ("text", "evidence", "swp", "metadata", "indexes"):
        inputs = sorted((source / category).rglob("*.txt")) if (source / category).exists() else []
        bundle: list[str] = []
        size = 0
        index = 1
        for item in inputs:
            text = item.read_text(encoding="utf-8", errors="replace")
            if bundle and size + len(text) + 2 > MAX_CHARS:
                destination = output / category / f"bundle-{index:03}.txt"
                write(destination, "\n\n".join(bundle) + "\n")
                index += 1
                bundle, size = [], 0
            bundle.append(text)
            size += len(text) + 2
            mapping[item.relative_to(source).as_posix()] = (output / category / f"bundle-{index:03}.txt").relative_to(output).as_posix()
        if bundle:
            destination = output / category / f"bundle-{index:03}.txt"
            write(destination, "\n\n".join(bundle) + "\n")

    coverage = source / "coverage.json"
    if coverage.exists():
        write(output / "coverage-manifest.txt", coverage.read_text(encoding="utf-8", errors="replace"))
    write(output / "README.txt", "Compacted text-only Graphify input. Each bundle retains original_path frontmatter for every represented source.\n")
    (output / "compaction-map.json").write_text(json.dumps(mapping, indent=2), encoding="utf-8")
    staged = list(output.rglob("*.txt"))
    print(json.dumps({"input_text_files": sum(1 for _ in source.rglob('*.txt')), "output_text_files": len(staged), "max_chars": MAX_CHARS}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
