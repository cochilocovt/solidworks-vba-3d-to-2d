#!/usr/bin/env python3
"""Build a text-only Graphify corpus for this SOLIDWORKS/VBA checkout.

The script is deliberately read-only with respect to the project inputs.  Its
only writes are beneath --output-root, which must be outside the project root.
Every selected source gets a coverage row; proprietary files become extracted
VBA text or an explicit metadata-only/error record.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


TEXT_EXTENSIONS = {
    ".bas", ".cls", ".frm", ".ps1", ".psm1", ".psd1", ".cs", ".py",
    ".md", ".txt", ".json", ".yaml", ".yml", ".html", ".htm", ".rst",
    ".xml", ".ini", ".cfg", ".vbs", ".vb", ".sln", ".csproj", ".log",
}
SOLIDWORKS_EXTENSIONS = {".swp", ".sldprt", ".slddrw", ".sldasm", ".drwdot", ".prtdot", ".asmdot"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}
EXCLUDED_PARTS = {".git", "graphify-out", ".codex", ".codex-tools", ".tools", ".firecrawl", "node_modules", "__pycache__", ".venv", ".claude"}
EVIDENCE_PREFIX = Path("test_assets") / "iteration_evidence"
MAX_EVIDENCE_CHARS = 70_000


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_text(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            pass
    return raw.decode("cp1252", errors="replace"), "cp1252-replace"


def safe_rel(rel: Path) -> Path:
    return Path(*[re.sub(r"[^A-Za-z0-9._-]", "_", part) for part in rel.parts])


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content.replace("\r\n", "\n").replace("\r", "\n"), encoding="utf-8")


def header(rel: Path, kind: str, **fields: str) -> str:
    lines = ["---", f"original_path: {rel.as_posix()}", f"corpus_kind: {kind}"]
    lines.extend(f"{key}: {value}" for key, value in fields.items())
    return "\n".join(lines) + "\n---\n\n"


def authority(rel: Path) -> str:
    if rel.parts and rel.parts[0] == "src" and len(rel.parts) > 1 and rel.parts[1] == "baseline-model-dims":
        return "trunk"
    if rel.parts and rel.parts[0] in {"archive"}:
        return "historical"
    if rel.parts[:2] == ("src", "active-ordinate"):
        return "historical"
    if rel.parts and rel.parts[0] == "test_assets":
        return "evidence"
    if rel.parts and rel.parts[0] == "tools":
        return "tooling"
    if rel.parts and rel.parts[0] == "docs":
        return "documentation"
    return "project"


def project_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        try:
            rel = path.relative_to(root)
        except ValueError:
            continue
        if any(part in EXCLUDED_PARTS for part in rel.parts):
            continue
        yield path


def build_ole_parser(root: Path):
    local = root / ".codex-tools" / "oletools"
    if local.is_dir():
        sys.path.insert(0, str(local))
    try:
        from oletools.olevba import VBA_Parser  # type: ignore
    except ImportError:
        return None
    return VBA_Parser


def extract_swp(parser_type, path: Path) -> list[tuple[str, str]]:
    parser = parser_type(str(path))
    modules: list[tuple[str, str]] = []
    try:
        for _, stream_path, filename, code in parser.extract_macros():
            if not filename or str(filename).startswith("VBA_P-code"):
                continue
            if isinstance(code, bytes):
                code = code.decode("cp1252", errors="replace")
            modules.append((f"{filename} ({stream_path})", str(code)))
    finally:
        parser.close()
    return modules


def pack_evidence(output: Path, items: list[tuple[Path, str]], coverage: list[dict]) -> None:
    if not items:
        return
    bundle: list[str] = []
    bundle_size = 0
    index = 1
    for rel, content in items:
        section = header(rel, "evidence-text", authority="evidence") + content + "\n\n"
        if bundle and bundle_size + len(section) > MAX_EVIDENCE_CHARS:
            write_text(output / "evidence" / f"bundle-{index:03}.txt", "".join(bundle))
            index += 1
            bundle, bundle_size = [], 0
        bundle.append(section)
        bundle_size += len(section)
        coverage.append({"path": rel.as_posix(), "kind": "text", "status": "staged-in-evidence-bundle", "stage": f"evidence/bundle-{index:03}.txt"})
    if bundle:
        write_text(output / "evidence" / f"bundle-{index:03}.txt", "".join(bundle))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--clean", action="store_true")
    args = parser.parse_args()

    root = args.project_root.resolve()
    output = args.output_root.resolve()
    if output == root or root in output.parents:
        raise SystemExit("--output-root must be outside the project root")
    if args.clean and output.exists():
        shutil.rmtree(output)
    output.mkdir(parents=True, exist_ok=True)

    coverage: list[dict] = []
    text_evidence: list[tuple[Path, str]] = []
    swp_paths: list[tuple[Path, Path]] = []
    image_paths: list[tuple[Path, Path]] = []
    metadata_paths: list[tuple[Path, Path]] = []

    for path in project_files(root):
        rel = path.relative_to(root)
        suffix = path.suffix.lower()
        if suffix in TEXT_EXTENSIONS:
            try:
                content, encoding = read_text(path)
            except OSError as exc:
                coverage.append({"path": rel.as_posix(), "kind": "text", "status": "error", "error": str(exc)})
                continue
            if rel.is_relative_to(EVIDENCE_PREFIX):
                text_evidence.append((rel, content))
            else:
                target = output / "text" / safe_rel(rel).with_suffix(rel.suffix + ".txt")
                write_text(target, header(rel, "text-source", encoding=encoding, authority=authority(rel)) + content)
                coverage.append({"path": rel.as_posix(), "kind": "text", "status": "staged", "stage": target.relative_to(output).as_posix()})
        elif suffix == ".swp":
            swp_paths.append((path, rel))
        elif suffix in SOLIDWORKS_EXTENSIONS:
            metadata_paths.append((path, rel))
        elif suffix in IMAGE_EXTENSIONS and rel.parts and rel.parts[0] == "test_assets":
            image_paths.append((path, rel))

    pack_evidence(output, text_evidence, coverage)

    vba_parser = build_ole_parser(root)
    canonical: dict[str, dict] = {}
    swp_index: list[str] = ["# SOLIDWORKS macro inventory\n\n"]
    for path, rel in sorted(swp_paths, key=lambda item: item[1].as_posix().lower()):
        try:
            digest = sha256_file(path)
        except OSError as exc:
            swp_index.append(f"## {rel.as_posix()}\n\nstatus: locked-or-unreadable\nerror: {exc}\n\n")
            coverage.append({"path": rel.as_posix(), "kind": "swp", "status": "metadata-only", "error": str(exc)})
            continue
        if digest in canonical:
            canonical[digest]["paths"].append(rel)
            coverage.append({"path": rel.as_posix(), "kind": "swp", "status": "duplicate", "canonical": canonical[digest]["canonical"].as_posix(), "sha256": digest})
            continue
        record = {"canonical": rel, "paths": [rel], "sha256": digest}
        canonical[digest] = record
        if vba_parser is None:
            record["error"] = "oletools.olevba unavailable"
            coverage.append({"path": rel.as_posix(), "kind": "swp", "status": "metadata-only", "error": record["error"], "sha256": digest})
            continue
        try:
            modules = extract_swp(vba_parser, path)
        except Exception as exc:  # OLE/VBA parsing must never stop the corpus build.
            record["error"] = str(exc)
            coverage.append({"path": rel.as_posix(), "kind": "swp", "status": "metadata-only", "error": str(exc), "sha256": digest})
            continue
        body = [header(rel, "swp-vba", authority=authority(rel), sha256=digest, module_count=str(len(modules)))]
        for name, code in modules:
            body.append(f"## VBA component: {name}\n\n{code}\n\n")
        target = output / "swp" / f"{digest}.txt"
        write_text(target, "".join(body))
        record["stage"] = target.relative_to(output).as_posix()
        coverage.append({"path": rel.as_posix(), "kind": "swp", "status": "extracted", "stage": record["stage"], "sha256": digest, "module_count": len(modules)})

    for digest, record in sorted(canonical.items(), key=lambda item: item[1]["canonical"].as_posix().lower()):
        paths = "\n".join(f"- {p.as_posix()}" for p in record["paths"])
        status = "extracted" if "stage" in record else "metadata-only"
        detail = record.get("stage") or record.get("error", "unknown")
        swp_index.append(f"## {record['canonical'].as_posix()}\n\nsha256: {digest}\nstatus: {status}\ndetail: {detail}\npaths represented:\n{paths}\n\n")
    write_text(output / "indexes" / "swp-inventory.txt", "".join(swp_index))

    for path, rel in metadata_paths:
        try:
            digest = sha256_file(path)
            status = "metadata-only"
            detail = "proprietary SOLIDWORKS document; no live API extraction was authorized"
        except OSError as exc:
            digest = "unavailable"
            status = "locked-or-unreadable"
            detail = str(exc)
        target = output / "metadata" / safe_rel(rel).with_suffix(rel.suffix + ".txt")
        write_text(target, header(rel, "solidworks-binary-metadata", authority=authority(rel), sha256=digest, status=status) + detail + "\n")
        coverage.append({"path": rel.as_posix(), "kind": "solidworks-binary", "status": status, "stage": target.relative_to(output).as_posix(), "sha256": digest})

    for path, rel in image_paths:
        digest = sha256_file(path)
        target = output / "metadata" / safe_rel(rel).with_suffix(rel.suffix + ".txt")
        write_text(target, header(rel, "reference-image-metadata", authority="reference", sha256=digest, status="metadata-only") + "Visual/OCR extraction is not available in this text-only corpus builder.\n")
        coverage.append({"path": rel.as_posix(), "kind": "reference-image", "status": "metadata-only", "stage": target.relative_to(output).as_posix(), "sha256": digest})

    stage_files = [p for p in output.rglob("*.txt")]
    counts = Counter(item["status"] for item in coverage)
    report = {
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "project_root": str(root),
        "output_root": str(output),
        "coverage": coverage,
        "summary": {"records": len(coverage), "stage_text_files": len(stage_files), "status_counts": dict(counts), "unique_swp": len(canonical), "swp_paths": len(swp_paths)},
    }
    write_text(output / "coverage.json", json.dumps(report, indent=2, ensure_ascii=False) + "\n")
    write_text(output / "README.txt", "Graphify input corpus. All files here are generated text representations; see coverage.json for provenance and limitations.\n")
    print(json.dumps(report["summary"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
