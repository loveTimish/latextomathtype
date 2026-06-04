# -*- coding: utf-8 -*-
"""
Calibrate generated MathType/OLE object boxes to a reference DOCX.

This adjusts Word's VML/OLE display box metadata only:
- v:shape style width/height
- w:object dxaOrig/dyaOrig
- w:rPr/w:position baseline offset

It does not rewrite the OLE binary or the MTEF equation body, so MathType
editability remains controlled by the generated Equation.DSMT4 OLE object.
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

from extract_formula_boxes import extract_boxes


DEFAULT_REFERENCE = Path("rebuild-assets/external/fraction-split-reference.docx")
DEFAULT_INPUT = Path("target/reference-roundtrip/fraction-split-reference-regenerated.docx")
DEFAULT_OUTPUT = Path("target/reference-roundtrip/fraction-split-reference-regenerated.docx")
DEFAULT_REPORT = Path("target/reference-roundtrip/formula-box-calibration.json")


def format_pt(value: float) -> str:
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    return text or "0"


def replace_attr(tag: str, name: str, value: int | str) -> str:
    value_text = str(value)
    pattern = rf'({re.escape(name)}=")([^"]*)(")'
    if re.search(pattern, tag):
        return re.sub(pattern, rf"\g<1>{value_text}\g<3>", tag, count=1)
    insert_at = tag.rfind(">")
    if insert_at < 0:
        return tag
    slash = "/" if tag[:insert_at].rstrip().endswith("/") else ""
    if slash:
        insert_at = tag.rfind("/")
    return tag[:insert_at] + f' {name}="{value_text}"' + tag[insert_at:]


def replace_style_dimension(style: str, name: str, value_pt: float) -> str:
    replacement = f"{name}:{format_pt(value_pt)}pt"
    pattern = rf"{re.escape(name)}:[0-9.]+(?:pt|in|cm|mm)"
    if re.search(pattern, style):
        return re.sub(pattern, replacement, style, count=1)
    return style.rstrip(";") + ";" + replacement


def update_object_xml(object_xml: str, ref: dict) -> str:
    def object_tag_repl(match: re.Match[str]) -> str:
        tag = match.group(0)
        if ref.get("dxa_orig") is not None:
            tag = replace_attr(tag, "w:dxaOrig", int(ref["dxa_orig"]))
        if ref.get("dya_orig") is not None:
            tag = replace_attr(tag, "w:dyaOrig", int(ref["dya_orig"]))
        return tag

    def shape_tag_repl(match: re.Match[str]) -> str:
        tag = match.group(0)
        style_match = re.search(r'style="([^"]*)"', tag)
        if not style_match:
            return tag
        style = style_match.group(1)
        style = replace_style_dimension(style, "width", float(ref["style_width_pt"]))
        style = replace_style_dimension(style, "height", float(ref["style_height_pt"]))
        return tag[: style_match.start(1)] + style + tag[style_match.end(1) :]

    object_xml = re.sub(r"<w:object\b[^>]*>", object_tag_repl, object_xml, count=1, flags=re.S)
    object_xml = re.sub(r"<v:shape\b[^>]*>", shape_tag_repl, object_xml, count=1, flags=re.S)
    return object_xml


def update_position_before_object(prefix: str, position_half_pt: int | None) -> str:
    if position_half_pt is None:
        return prefix
    value = str(position_half_pt)
    pos_matches = list(re.finditer(r"<w:position\b[^>]*/?>", prefix, flags=re.S))
    if pos_matches:
        last = pos_matches[-1]
        tag = replace_attr(last.group(0), "w:val", value)
        return prefix[: last.start()] + tag + prefix[last.end() :]

    rpr_matches = list(re.finditer(r"<w:rPr\b[^>]*>", prefix, flags=re.S))
    if rpr_matches:
        last = rpr_matches[-1]
        insert_at = last.end()
        return prefix[:insert_at] + f'<w:position w:val="{value}"/>' + prefix[insert_at:]

    run_match = re.search(r"<w:r\b[^>]*>", prefix, flags=re.S)
    if run_match:
        insert_at = run_match.end()
        return prefix[:insert_at] + f'<w:rPr><w:position w:val="{value}"/></w:rPr>' + prefix[insert_at:]
    return prefix


def calibrate_document_xml(document_xml: str, reference_boxes: list[dict]) -> tuple[str, list[dict]]:
    object_matches = list(re.finditer(r"<w:object\b.*?</w:object>", document_xml, flags=re.S))
    if len(object_matches) != len(reference_boxes):
        raise ValueError(
            f"formula object count mismatch: generated={len(object_matches)} reference={len(reference_boxes)}"
        )

    output: list[str] = []
    cursor = 0
    report: list[dict] = []
    for index, match in enumerate(object_matches):
        ref = reference_boxes[index]
        prefix = document_xml[cursor : match.start()]
        prefix = update_position_before_object(prefix, ref.get("position_half_pt"))
        original = match.group(0)
        updated = update_object_xml(original, ref)
        output.append(prefix)
        output.append(updated)
        cursor = match.end()
        report.append(
            {
                "index": index,
                "reference_width_pt": ref["style_width_pt"],
                "reference_height_pt": ref["style_height_pt"],
                "reference_position_half_pt": ref.get("position_half_pt"),
                "reference_dxa_orig": ref.get("dxa_orig"),
                "reference_dya_orig": ref.get("dya_orig"),
            }
        )
    output.append(document_xml[cursor:])
    return "".join(output), report


def write_calibrated_docx(input_docx: Path, output_docx: Path, document_xml: str) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    same_path = input_docx.resolve() == output_docx.resolve()
    if same_path:
        fd, temp_name = tempfile.mkstemp(suffix=".docx", prefix="calibrated-")
        os.close(fd)
        target = Path(temp_name)
    else:
        target = output_docx
    try:
        with zipfile.ZipFile(input_docx, "r") as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "word/document.xml":
                    data = document_xml.encode("utf-8")
                dst.writestr(info, data)
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
    generated_boxes = extract_boxes(args.input)
    if len(reference_boxes) != len(generated_boxes):
        raise SystemExit(
            f"formula object count mismatch: reference={len(reference_boxes)} generated={len(generated_boxes)}"
        )

    with zipfile.ZipFile(args.input, "r") as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
    calibrated_xml, rows = calibrate_document_xml(document_xml, reference_boxes)
    write_calibrated_docx(args.input, args.output, calibrated_xml)

    report = {
        "reference": str(args.reference.resolve()),
        "input": str(args.input.resolve()),
        "output": str(args.output.resolve()),
        "formula_count": len(rows),
        "rows": rows,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"calibrated {len(rows)} formula boxes -> {args.output}")
    print(f"wrote {args.report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
