# -*- coding: utf-8 -*-
"""Rebuild a DOCX by preserving the source document shell and replacing formulas.

The PaperExportRequest path is useful for service-level exports, but it reflows
the whole document. This tool keeps the original DOCX package, relationships,
paragraphs, images, headers, footers, and pagination, then replaces paired
MathType OLE binaries and their WMF previews with generated outputs.
"""

from __future__ import annotations

import argparse
import csv
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path


def read_pairs(pair_csv: Path) -> list[dict]:
    with pair_csv.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def relationship_maps(docx: Path) -> tuple[dict[str, str], dict[str, str]]:
    with zipfile.ZipFile(docx) as z:
        rel_xml = z.read("word/_rels/document.xml.rels").decode("utf-8", "replace")
    rid_to_target: dict[str, str] = {}
    target_to_rid: dict[str, str] = {}
    for match in re.finditer(r"<Relationship\b[^>]*/>", rel_xml):
        item = match.group(0)
        rid_match = re.search(r'\bId="([^"]+)"', item)
        target_match = re.search(r'\bTarget="([^"]+)"', item)
        rid = rid_match.group(1) if rid_match else None
        target = target_match.group(1) if target_match else None
        if rid and target:
            rid_to_target[rid] = target
            target_to_rid[target] = rid
    return rid_to_target, target_to_rid


def all_relationship_maps(docx: Path) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(docx) as z:
        names = set(z.namelist())
        for part in names:
            if not part.startswith("word/") or not part.endswith(".xml") or part.endswith(".rels"):
                continue
            rel_part = rels_part_for_xml(part)
            if rel_part not in names:
                continue
            rel_xml = z.read(rel_part).decode("utf-8", "replace")
            rid_to_target: dict[str, str] = {}
            for match in re.finditer(r"<Relationship\b[^>]*/>", rel_xml):
                item = match.group(0)
                rid_match = re.search(r'\bId="([^"]+)"', item)
                target_match = re.search(r'\bTarget="([^"]+)"', item)
                if not rid_match or not target_match:
                    continue
                rid_to_target[rid_match.group(1)] = resolve_part_target(part, target_match.group(1))
            out[part] = rid_to_target
    return out


def object_preview_targets(docx: Path) -> dict[str, str]:
    rels_by_part = all_relationship_maps(docx)
    out: dict[str, str] = {}
    with zipfile.ZipFile(docx) as z:
        for part, rid_to_target in rels_by_part.items():
            xml = z.read(part).decode("utf-8", "replace")
            for match in re.finditer(r"<w:object\b[\s\S]*?</w:object>", xml):
                body = match.group(0)
                image = re.search(r'<v:imagedata [^>]*r:id="([^"]+)"', body)
                ole = re.search(r'<o:OLEObject [^>]*r:id="([^"]+)"', body)
                if not image or not ole:
                    continue
                ole_target = rid_to_target.get(ole.group(1))
                image_target = rid_to_target.get(image.group(1))
                if ole_target and image_target:
                    out[word_path(ole_target)] = word_path(image_target)
    return out


def generated_preview_by_ole(generated_docx: Path) -> dict[str, str]:
    return object_preview_targets(generated_docx)


def word_path(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("word/"):
        return target
    return posixpath.normpath(posixpath.join("word", target))


def rels_part_for_xml(part: str) -> str:
    directory, name = posixpath.split(part)
    return posixpath.join(directory, "_rels", f"{name}.rels")


def resolve_part_target(part: str, target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("word/"):
        return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def extract_docx(docx: Path, directory: Path) -> None:
    with zipfile.ZipFile(docx) as z:
        z.extractall(directory)


def copy_zip_entry(source_docx: Path, entry: str, dest_file: Path) -> None:
    with zipfile.ZipFile(source_docx) as z:
        dest_file.parent.mkdir(parents=True, exist_ok=True)
        dest_file.write_bytes(z.read(entry))


def zip_dir(directory: Path, output: Path) -> None:
    if output.exists():
        output.unlink()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(directory.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(directory).as_posix())


def rebuild(source_docx: Path, generated_docx: Path, pair_csv: Path, output: Path) -> dict:
    pairs = read_pairs(pair_csv)
    source_preview = object_preview_targets(source_docx)
    generated_preview = generated_preview_by_ole(generated_docx)

    replaced_ole = 0
    replaced_preview = 0
    missing_generated = 0
    with tempfile.TemporaryDirectory(prefix="xsc-template-rebuild-") as temp:
        root = Path(temp)
        extract_docx(source_docx, root)
        for row in pairs:
            source_ole = word_path(row.get("sourceEntry", ""))
            generated_ole = word_path(row.get("generatedEntry", ""))
            if not source_ole or not generated_ole:
                continue
            source_ole_path = root / source_ole
            try:
                copy_zip_entry(generated_docx, generated_ole, source_ole_path)
                replaced_ole += 1
            except KeyError:
                missing_generated += 1
                continue

            source_image = source_preview.get(source_ole)
            generated_image = generated_preview.get(generated_ole)
            if source_image and generated_image:
                try:
                    copy_zip_entry(generated_docx, generated_image, root / source_image)
                    replaced_preview += 1
                except KeyError:
                    missing_generated += 1

        output.parent.mkdir(parents=True, exist_ok=True)
        zip_dir(root, output)
    return {
        "source": str(source_docx),
        "generated": str(generated_docx),
        "pairCsv": str(pair_csv),
        "output": str(output),
        "pairs": len(pairs),
        "replacedOle": replaced_ole,
        "replacedPreviews": replaced_preview,
        "missingGeneratedEntries": missing_generated,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("generated_docx", type=Path)
    parser.add_argument("pair_csv", type=Path)
    parser.add_argument("output_docx", type=Path)
    args = parser.parse_args()
    summary = rebuild(args.source_docx, args.generated_docx, args.pair_csv, args.output_docx)
    for key, value in summary.items():
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()
