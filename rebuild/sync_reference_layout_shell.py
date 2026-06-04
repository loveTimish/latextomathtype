# -*- coding: utf-8 -*-
"""
Build a reference-layout DOCX shell with generated MathType/OLE payloads.

The reference round-trip has two independent acceptance surfaces:
- visible Word layout/preview, which is already correct in the reference DOCX;
- editable MathType/OLE payloads, which come from latextomathtype generation.

This post-processor keeps the reference document XML, styles, media previews,
headers and footers intact, but replaces each referenced MathType OLE embedding
with the corresponding generated embedding by formula index.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from extract_formula_boxes import extract_boxes


DEFAULT_REFERENCE = Path("rebuild-assets/external/fraction-split-reference.docx")
DEFAULT_GENERATED = Path("target/reference-roundtrip/fraction-split-reference-regenerated.docx")
DEFAULT_OUTPUT = Path("target/reference-roundtrip/fraction-split-reference-regenerated.docx")
DEFAULT_REPORT = Path("target/reference-roundtrip/reference-layout-shell-sync.json")


def part_name(target: str | None) -> str | None:
    if not target:
        return None
    return "word/" + target.lstrip("/")


def read_part(zf: zipfile.ZipFile, name: str) -> bytes:
    try:
        return zf.read(name)
    except KeyError as exc:
        raise ValueError(f"missing DOCX part: {name}") from exc


def build_embedding_map(reference_docx: Path, generated_docx: Path) -> tuple[dict[str, bytes], list[dict[str, Any]]]:
    reference_boxes = extract_boxes(reference_docx)
    generated_boxes = extract_boxes(generated_docx)
    if len(reference_boxes) != len(generated_boxes):
        raise SystemExit(
            f"formula object count mismatch: reference={len(reference_boxes)} generated={len(generated_boxes)}"
        )

    replacements: dict[str, bytes] = {}
    rows: list[dict[str, Any]] = []
    with zipfile.ZipFile(generated_docx, "r") as generated_zip:
        for index, (ref_box, gen_box) in enumerate(zip(reference_boxes, generated_boxes)):
            ref_part = part_name(ref_box.ole_target)
            gen_part = part_name(gen_box.ole_target)
            if ref_part is None or gen_part is None:
                raise SystemExit(f"missing OLE relationship at formula index {index}")
            generated_bytes = read_part(generated_zip, gen_part)
            replacements[ref_part] = generated_bytes
            rows.append(
                {
                    "index": index,
                    "reference_embedding": ref_part,
                    "generated_embedding": gen_part,
                    "bytes": len(generated_bytes),
                    "reference_context": ref_box.context,
                    "generated_context": gen_box.context,
                }
            )
    return replacements, rows


def write_docx(reference_docx: Path, output_docx: Path, replacements: dict[str, bytes]) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    same_path = reference_docx.resolve() == output_docx.resolve()
    if same_path:
        raise SystemExit("output must not overwrite the reference DOCX")

    fd, temp_name = tempfile.mkstemp(suffix=".docx", prefix="reference-layout-shell-")
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(reference_docx, "r") as src, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = replacements.get(info.filename)
                if data is None:
                    data = src.read(info.filename)
                dst.writestr(info, data)
        shutil.move(str(temp_path), str(output_docx))
    finally:
        temp_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--generated", type=Path, default=DEFAULT_GENERATED)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()

    replacements, rows = build_embedding_map(args.reference, args.generated)
    write_docx(args.reference, args.output, replacements)

    report = {
        "reference": str(args.reference.resolve()),
        "generated": str(args.generated.resolve()),
        "output": str(args.output.resolve()),
        "formula_count": len(rows),
        "strategy": "reference document layout/media shell retained; generated MathType OLE embeddings transplanted by formula index",
        "rows": rows,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"synced {len(rows)} generated OLE payloads into reference layout shell -> {args.output}")
    print(f"wrote {args.report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
