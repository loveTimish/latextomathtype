# -*- coding: utf-8 -*-
"""
Extract MathType/OLE formula boxes from DOCX files.

The output is intentionally OOXML-centric: it records the Word object box,
image/OLE relationship ids, paragraph locality, baseline position, and nearby
text.  This is the raw dataset used by the reference-roundtrip calibrator and
comparison scripts.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


NS_DOC_RELS = "word/_rels/document.xml.rels"
NS_DOC = "word/document.xml"


@dataclass
class FormulaBox:
    index: int
    paragraph_index: int
    paragraph_local_index: int
    style_width_pt: float
    style_height_pt: float
    dxa_orig: int | None
    dya_orig: int | None
    dxa_orig_pt: float | None
    dya_orig_pt: float | None
    position_half_pt: int | None
    shape_id: str | None
    shape_style: str | None
    ole_rid: str | None
    ole_target: str | None
    image_rid: str | None
    image_target: str | None
    image_title: str | None
    prog_id: str | None
    context: str


def read_zip_text(docx: Path, name: str) -> str:
    with zipfile.ZipFile(docx) as zf:
        try:
            return zf.read(name).decode("utf-8")
        except KeyError:
            return ""


def zip_names(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as zf:
        return zf.namelist()


def attr(xml: str, name: str) -> str | None:
    match = re.search(rf'{re.escape(name)}="([^"]*)"', xml)
    return match.group(1) if match else None


def int_attr(xml: str, name: str) -> int | None:
    value = attr(xml, name)
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def style_length_to_pt(style: str, name: str) -> float | None:
    match = re.search(rf"{re.escape(name)}:([0-9.]+)(pt|in|cm|mm)", style)
    if not match:
        return None
    value = float(match.group(1))
    unit = match.group(2)
    if unit == "in":
        return value * 72.0
    if unit == "cm":
        return value * 72.0 / 2.54
    if unit == "mm":
        return value * 72.0 / 25.4
    return value


def decode_entities(text: str) -> str:
    return (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
    )


def text_from_paragraph(p_xml: str) -> str:
    chunks = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", p_xml, flags=re.S)
    text = decode_entities("".join(chunks))
    return re.sub(r"\s+", " ", text).strip()


def trim_context(text: str, limit: int = 120) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "..."


def parse_relationships(xml: str) -> dict[str, str]:
    rows: dict[str, str] = {}
    for match in re.finditer(r"<Relationship\b[^>]*/?>", xml):
        rid = attr(match.group(0), "Id")
        target = attr(match.group(0), "Target")
        if rid and target:
            rows[rid] = target
    return rows


def extract_boxes(docx: Path) -> list[FormulaBox]:
    document_xml = read_zip_text(docx, NS_DOC)
    rels = parse_relationships(read_zip_text(docx, NS_DOC_RELS))
    paragraphs = re.findall(r"<w:p\b.*?</w:p>", document_xml, flags=re.S)
    rows: list[FormulaBox] = []
    global_index = 0

    for p_idx, p_xml in enumerate(paragraphs):
        context = trim_context(text_from_paragraph(p_xml))
        objects = list(re.finditer(r"<w:object\b.*?</w:object>", p_xml, flags=re.S))
        for local_index, obj_match in enumerate(objects):
            obj_xml = obj_match.group(0)
            before = p_xml[: obj_match.start()]
            pos_matches = list(re.finditer(r"<w:position\b[^>]*/?>", before))
            position = int_attr(pos_matches[-1].group(0), "w:val") if pos_matches else None
            shape_match = re.search(r"<v:shape\b[^>]*>", obj_xml, flags=re.S)
            if not shape_match:
                continue
            shape_tag = shape_match.group(0)
            style = attr(shape_tag, "style")
            if not style:
                continue
            width = style_length_to_pt(style, "width")
            height = style_length_to_pt(style, "height")
            if width is None or height is None:
                continue
            ole_match = re.search(r"<o:OLEObject\b[^>]*/?>", obj_xml)
            ole_tag = ole_match.group(0) if ole_match else ""
            image_match = re.search(r"<v:imagedata\b[^>]*/?>", obj_xml)
            image_tag = image_match.group(0) if image_match else ""
            ole_rid = attr(ole_tag, "r:id")
            image_rid = attr(image_tag, "r:id")
            dxa_orig = int_attr(obj_xml, "w:dxaOrig")
            dya_orig = int_attr(obj_xml, "w:dyaOrig")
            rows.append(
                FormulaBox(
                    index=global_index,
                    paragraph_index=p_idx,
                    paragraph_local_index=local_index,
                    style_width_pt=round(width, 2),
                    style_height_pt=round(height, 2),
                    dxa_orig=dxa_orig,
                    dya_orig=dya_orig,
                    dxa_orig_pt=round(dxa_orig / 20.0, 2) if dxa_orig is not None else None,
                    dya_orig_pt=round(dya_orig / 20.0, 2) if dya_orig is not None else None,
                    position_half_pt=position,
                    shape_id=attr(shape_tag, "id"),
                    shape_style=style,
                    ole_rid=ole_rid,
                    ole_target=rels.get(ole_rid or ""),
                    image_rid=image_rid,
                    image_target=rels.get(image_rid or ""),
                    image_title=attr(image_tag, "o:title"),
                    prog_id=attr(ole_tag, "ProgID"),
                    context=context,
                )
            )
            global_index += 1
    return rows


def summarize(docx: Path, boxes: list[FormulaBox]) -> dict[str, Any]:
    names = zip_names(docx)
    widths = [box.style_width_pt for box in boxes]
    heights = [box.style_height_pt for box in boxes]
    return {
        "path": str(docx.resolve()),
        "bytes": docx.stat().st_size,
        "object_count": len(boxes),
        "embedding_count": len([n for n in names if n.startswith("word/embeddings/")]),
        "media_count": len([n for n in names if n.startswith("word/media/")]),
        "width_min": round(min(widths), 2) if widths else None,
        "width_max": round(max(widths), 2) if widths else None,
        "height_min": round(min(heights), 2) if heights else None,
        "height_max": round(max(heights), 2) if heights else None,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()

    boxes = extract_boxes(args.docx)
    payload = {
        "summary": summarize(args.docx, boxes),
        "boxes": [asdict(box) for box in boxes],
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
