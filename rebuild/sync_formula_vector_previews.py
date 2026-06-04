# -*- coding: utf-8 -*-
"""
Synchronize generated MathType/OLE preview images with reference WMF previews.

MathType equations in Word carry two synchronized payloads:
- WMF preview data, which Word renders on the page.
- MTEF data inside the OLE object, which MathType opens for editing.

For reference round-trip validation we keep the generated OLE/MTEF binaries, but
replace each generated PNG preview relationship with the corresponding WMF
preview from the reference document.  This preserves Equation.DSMT4 editability
while matching the vector preview surface that Word actually displays.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import asdict
from pathlib import Path
from typing import Iterable

from extract_formula_boxes import extract_boxes


DEFAULT_REFERENCE = Path("rebuild-assets/external/fraction-split-reference.docx")
DEFAULT_INPUT = Path("target/reference-roundtrip/fraction-split-reference-regenerated.docx")
DEFAULT_OUTPUT = Path("target/reference-roundtrip/fraction-split-reference-regenerated.docx")
DEFAULT_REPORT = Path("target/reference-roundtrip/formula-vector-preview-sync.json")
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
WMF_CONTENT_TYPE = "image/x-wmf"


def read_part(zf: zipfile.ZipFile, name: str) -> bytes:
    try:
        return zf.read(name)
    except KeyError as exc:
        raise ValueError(f"missing DOCX part: {name}") from exc


def ensure_wmf_content_type(content_types_xml: str) -> str:
    if re.search(r'<Default\b[^>]*Extension="wmf"[^>]*/>', content_types_xml, flags=re.I):
        return re.sub(
            r'(<Default\b[^>]*Extension="wmf"[^>]*ContentType=")([^"]*)(")',
            rf"\g<1>{WMF_CONTENT_TYPE}\g<3>",
            content_types_xml,
            count=1,
            flags=re.I,
        )
    insert_at = content_types_xml.rfind("</Types>")
    if insert_at < 0:
        raise ValueError("[Content_Types].xml has no closing </Types>")
    default = f'<Default Extension="wmf" ContentType="{WMF_CONTENT_TYPE}"/>'
    return content_types_xml[:insert_at] + default + content_types_xml[insert_at:]


def replace_relationship_targets(rels_xml: str, rid_to_target: dict[str, str]) -> str:
    def repl(match: re.Match[str]) -> str:
        tag = match.group(0)
        rid_match = re.search(r'\bId="([^"]+)"', tag)
        if not rid_match:
            return tag
        rid = rid_match.group(1)
        target = rid_to_target.get(rid)
        if not target:
            return tag
        if IMAGE_REL_TYPE not in tag:
            return tag
        if re.search(r'\bTarget="[^"]*"', tag):
            return re.sub(r'\bTarget="[^"]*"', f'Target="{target}"', tag, count=1)
        insert_at = tag.rfind("/>")
        if insert_at < 0:
            insert_at = tag.rfind(">")
        return tag[:insert_at] + f' Target="{target}"' + tag[insert_at:]

    return re.sub(r"<Relationship\b[^>]*/?>", repl, rels_xml)


def names_to_skip(generated_boxes: Iterable[dict]) -> set[str]:
    skipped: set[str] = set()
    for box in generated_boxes:
        image_target = box.get("image_target")
        if image_target:
            skipped.add("word/" + image_target.lstrip("/"))
    return skipped


def write_synced_docx(
    reference_docx: Path,
    input_docx: Path,
    output_docx: Path,
    generated_boxes: list[dict],
    ref_media: dict[str, bytes],
    rid_to_target: dict[str, str],
) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    same_path = input_docx.resolve() == output_docx.resolve()
    if same_path:
        fd, temp_name = tempfile.mkstemp(suffix=".docx", prefix="vector-preview-")
        os.close(fd)
        target = Path(temp_name)
    else:
        target = output_docx

    old_preview_parts = names_to_skip(generated_boxes)
    try:
        with zipfile.ZipFile(input_docx, "r") as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                if info.filename in old_preview_parts:
                    continue
                data = src.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    data = ensure_wmf_content_type(data.decode("utf-8")).encode("utf-8")
                elif info.filename == "word/_rels/document.xml.rels":
                    data = replace_relationship_targets(data.decode("utf-8"), rid_to_target).encode("utf-8")
                dst.writestr(info, data)

            for target_name, media_bytes in ref_media.items():
                dst.writestr("word/" + target_name, media_bytes)
        if same_path:
            shutil.move(str(target), str(output_docx))
    finally:
        if same_path and target.exists():
            target.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()

    reference_boxes = [asdict(item) for item in extract_boxes(args.reference)]
    generated_boxes = [asdict(item) for item in extract_boxes(args.input)]
    if len(reference_boxes) != len(generated_boxes):
        raise SystemExit(
            f"formula object count mismatch: reference={len(reference_boxes)} generated={len(generated_boxes)}"
        )

    ref_media: dict[str, bytes] = {}
    rid_to_target: dict[str, str] = {}
    rows: list[dict] = []
    with zipfile.ZipFile(args.reference, "r") as ref_zip:
        for index, (ref_box, gen_box) in enumerate(zip(reference_boxes, generated_boxes), start=1):
            source_target = ref_box.get("image_target")
            image_rid = gen_box.get("image_rid")
            if not source_target or not image_rid:
                raise SystemExit(f"missing image relationship at formula index {index - 1}")
            if Path(source_target).suffix.lower() != ".wmf":
                raise SystemExit(f"reference preview is not WMF at formula index {index - 1}: {source_target}")

            media_bytes = read_part(ref_zip, "word/" + source_target.lstrip("/"))
            new_target = f"media/image_eq{index}_vector.wmf"
            ref_media[new_target] = media_bytes
            rid_to_target[image_rid] = new_target
            rows.append(
                {
                    "index": index - 1,
                    "image_rid": image_rid,
                    "generated_preview": gen_box.get("image_target"),
                    "reference_preview": source_target,
                    "synced_preview": new_target,
                    "bytes": len(media_bytes),
                }
            )

    write_synced_docx(args.reference, args.input, args.output, generated_boxes, ref_media, rid_to_target)

    report = {
        "reference": str(args.reference.resolve()),
        "input": str(args.input.resolve()),
        "output": str(args.output.resolve()),
        "formula_count": len(rows),
        "strategy": "generated Equation.DSMT4 OLE/MTEF retained; reference WMF vector previews reused",
        "rows": rows,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"synced {len(rows)} vector formula previews -> {args.output}")
    print(f"wrote {args.report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
