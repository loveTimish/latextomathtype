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


def source_object_preview_targets(source_docx: Path) -> dict[str, str]:
    rid_to_target, _ = relationship_maps(source_docx)
    with zipfile.ZipFile(source_docx) as z:
        document_xml = z.read("word/document.xml").decode("utf-8", "replace")
    out: dict[str, str] = {}
    for match in re.finditer(r"<w:object\b[\s\S]*?</w:object>", document_xml):
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
    rid_to_target, _ = relationship_maps(generated_docx)
    with zipfile.ZipFile(generated_docx) as z:
        document_xml = z.read("word/document.xml").decode("utf-8", "replace")
    out: dict[str, str] = {}
    for match in re.finditer(r"<w:object\b[\s\S]*?</w:object>", document_xml):
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


def word_path(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("word/"):
        return target
    return posixpath.normpath(posixpath.join("word", target))


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
    source_preview = source_object_preview_targets(source_docx)
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
