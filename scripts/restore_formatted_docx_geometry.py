#!/usr/bin/env python3
"""Restore project-computed Word/WMF geometry after MathType formatting.

MathType Format Equations rewrites the editable OLE and its preview. It also
recomputes the Word object box using the desktop application's current font
metrics. This tool keeps the formatted OLE and WMF drawing records, while
restoring geometry from the project-generated input document.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import struct
import tempfile
import zipfile
from pathlib import Path


OBJECT_RE = re.compile(r"<w:object\b([^>]*)>(.*?)</w:object>", re.S)
PLACEABLE_WMF_KEY = 0x9AC6CDD7


def relationship_map(archive: zipfile.ZipFile) -> dict[str, str]:
    xml = archive.read("word/_rels/document.xml.rels").decode("utf-8", "replace")
    return dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', xml))


def word_part_name(target: str) -> str:
    target = target.lstrip("/")
    return target if target.startswith("word/") else "word/" + target


def object_image_target(body: str, relationships: dict[str, str]) -> str | None:
    match = re.search(r'<v:imagedata\b[^>]*r:id="([^"]+)"', body)
    if not match:
        return None
    target = relationships.get(match.group(1))
    return word_part_name(target) if target else None


def css_dimension(style: str, name: str) -> str | None:
    match = re.search(
        rf"(?:^|;)\s*{re.escape(name)}:\s*[^;]+",
        style,
        re.I,
    )
    return match.group(0).lstrip("; ") if match else None


def replace_css_dimension(style: str, name: str, declaration: str | None) -> str:
    if not declaration:
        return style
    pattern = re.compile(rf"((?:^|;)\s*){re.escape(name)}:\s*[^;]+", re.I)
    if pattern.search(style):
        return pattern.sub(lambda match: match.group(1) + declaration, style, count=1)
    separator = "" if not style or style.endswith(";") else ";"
    return style + separator + declaration


def restore_object_geometry(before_object: str, formatted_object: str) -> str:
    before_open = re.match(r"<w:object\b([^>]*)>", before_object)
    formatted_open = re.match(r"<w:object\b([^>]*)>", formatted_object)
    if not before_open or not formatted_open:
        raise ValueError("invalid w:object XML")

    restored = formatted_object
    for attribute in ("w:dxaOrig", "w:dyaOrig"):
        value = re.search(rf'{re.escape(attribute)}="([^"]+)"', before_open.group(1))
        if not value:
            continue
        restored, count = re.subn(
            rf'({re.escape(attribute)}=")[^"]+(\")',
            rf"\g<1>{value.group(1)}\g<2>",
            restored,
            count=1,
        )
        if count == 0:
            restored = restored.replace("<w:object", f'<w:object {attribute}="{value.group(1)}"', 1)

    before_style = re.search(r'<v:shape\b[^>]*\sstyle="([^"]*)"', before_object)
    formatted_style = re.search(r'<v:shape\b[^>]*\sstyle="([^"]*)"', restored)
    if before_style and formatted_style:
        style = formatted_style.group(1)
        for name in ("width", "height"):
            style = replace_css_dimension(style, name, css_dimension(before_style.group(1), name))
        restored = restored[: formatted_style.start(1)] + style + restored[formatted_style.end(1) :]
    return restored


def preceding_position(segment: str) -> str | None:
    run_start = segment.rfind("<w:r")
    tail = segment[run_start:] if run_start >= 0 else segment
    matches = list(re.finditer(r'<w:position\b[^>]*w:val="(-?\d+)"[^>]*/>', tail))
    return matches[-1].group(1) if matches else None


def restore_preceding_position(segment: str, value: str | None) -> str:
    if value is None:
        return segment
    run_start = segment.rfind("<w:r")
    if run_start < 0:
        return segment
    prefix, tail = segment[:run_start], segment[run_start:]
    matches = list(re.finditer(r'(<w:position\b[^>]*w:val=")-?\d+("[^>]*/>)', tail))
    if not matches:
        return segment
    match = matches[-1]
    tail = tail[: match.start()] + match.group(1) + value + match.group(2) + tail[match.end() :]
    return prefix + tail


def restore_document_xml(before_xml: str, formatted_xml: str) -> tuple[str, list[tuple[str, str]]]:
    before_matches = list(OBJECT_RE.finditer(before_xml))
    formatted_matches = list(OBJECT_RE.finditer(formatted_xml))
    if len(before_matches) != len(formatted_matches):
        raise ValueError(
            f"MathType object count changed: before={len(before_matches)} formatted={len(formatted_matches)}"
        )

    media_pairs: list[tuple[str, str]] = []
    pieces: list[str] = []
    before_cursor = formatted_cursor = 0
    for before_match, formatted_match in zip(before_matches, formatted_matches):
        before_segment = before_xml[before_cursor : before_match.start()]
        formatted_segment = formatted_xml[formatted_cursor : formatted_match.start()]
        pieces.append(restore_preceding_position(formatted_segment, preceding_position(before_segment)))
        pieces.append(restore_object_geometry(before_match.group(0), formatted_match.group(0)))
        before_cursor = before_match.end()
        formatted_cursor = formatted_match.end()
    pieces.append(formatted_xml[formatted_cursor:])
    return "".join(pieces), media_pairs


def is_placeable_wmf(data: bytes) -> bool:
    return len(data) >= 22 and struct.unpack("<I", data[:4])[0] == PLACEABLE_WMF_KEY


def restore_geometry(before_docx: Path, formatted_docx: Path, output_docx: Path) -> dict[str, int]:
    with zipfile.ZipFile(before_docx) as before_zip, zipfile.ZipFile(formatted_docx) as formatted_zip:
        before_xml = before_zip.read("word/document.xml").decode("utf-8", "replace")
        formatted_xml = formatted_zip.read("word/document.xml").decode("utf-8", "replace")
        restored_xml, _ = restore_document_xml(before_xml, formatted_xml)

        before_relationships = relationship_map(before_zip)
        formatted_relationships = relationship_map(formatted_zip)
        before_objects = list(OBJECT_RE.finditer(before_xml))
        formatted_objects = list(OBJECT_RE.finditer(formatted_xml))
        restored_wmf: dict[str, bytes] = {}
        for before_object, formatted_object in zip(before_objects, formatted_objects):
            before_media = object_image_target(before_object.group(2), before_relationships)
            formatted_media = object_image_target(formatted_object.group(2), formatted_relationships)
            if not before_media or not formatted_media:
                continue
            if not before_media.lower().endswith(".wmf") or not formatted_media.lower().endswith(".wmf"):
                continue
            before_data = before_zip.read(before_media)
            formatted_data = formatted_zip.read(formatted_media)
            if is_placeable_wmf(before_data) and is_placeable_wmf(formatted_data):
                restored_wmf[formatted_media] = before_data[:22] + formatted_data[22:]

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary_name = tempfile.mkstemp(
            prefix=output_docx.stem + ".geometry-", suffix=".docx", dir=output_docx.parent
        )
        os.close(fd)
        temporary = Path(temporary_name)
        try:
            with zipfile.ZipFile(temporary, "w") as output_zip:
                for info in formatted_zip.infolist():
                    if info.filename == "word/document.xml":
                        data = restored_xml.encode("utf-8")
                    else:
                        data = restored_wmf.get(info.filename, formatted_zip.read(info.filename))
                    output_zip.writestr(info, data)
            os.replace(temporary, output_docx)
        finally:
            if temporary.exists():
                temporary.unlink()

    return {
        "objectCount": len(before_objects),
        "restoredWmfCount": len(restored_wmf),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("before_docx", type=Path)
    parser.add_argument("formatted_docx", type=Path)
    parser.add_argument("output_docx", type=Path)
    args = parser.parse_args()
    result = restore_geometry(args.before_docx, args.formatted_docx, args.output_docx)
    print(json.dumps(result))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
