# -*- coding: utf-8 -*-
"""Inspect DOCX MathType WMF formula geometry, text runs, and preview ink.

This is a diagnostic tool, not an acceptance gate.  WMF ExtTextOut dx values are
layout advances; they can differ from the visible glyph ink that Word renders.
The report keeps both layers side by side so calibration can separate record
geometry changes from real preview changes.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import asdict
from pathlib import Path
from typing import Any

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
REBUILD_DIR = ROOT / "rebuild"
if str(REBUILD_DIR) not in sys.path:
    sys.path.insert(0, str(REBUILD_DIR))

from extract_formula_boxes import extract_boxes  # noqa: E402

SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from compare_docx_pair_metrics import normalize_latex_key  # noqa: E402

try:
    from PIL import Image
except ImportError:  # pragma: no cover - dependency is present in the Codex runtime.
    Image = None


PLACEABLE_WMF_KEY = 0x9AC6CDD7
META_EOF = 0x0000
META_SETWINDOWEXT = 0x020C
META_SELECTOBJECT = 0x012D
META_DELETEOBJECT = 0x01F0
META_CREATEBRUSHINDIRECT = 0x02FC
META_CREATEPENINDIRECT = 0x02FA
META_CREATEPALETTE = 0x00F7
META_CREATEREGION = 0x06FF
META_CREATEFONTINDIRECT = 0x02FB
META_POLYLINE = 0x0325
META_TEXTOUT = 0x0521
META_EXTTEXTOUT = 0x0A32
ETO_OPAQUE = 0x0002
ETO_CLIPPED = 0x0004
DEFAULT_MAGICK = Path(r"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe")


def u16(data: bytes, offset: int) -> int:
    return struct.unpack_from("<H", data, offset)[0]


def s16(data: bytes, offset: int) -> int:
    return struct.unpack_from("<h", data, offset)[0]


def u32(data: bytes, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def round3(value: float | None) -> float | None:
    return None if value is None else round(value, 3)


def parse_index_list(value: str | None) -> set[int] | None:
    if not value:
        return None
    indexes: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        indexes.add(int(part))
    return indexes


def read_docx_target(zf: zipfile.ZipFile, target: str | None) -> bytes | None:
    if not target:
        return None
    name = target.lstrip("/")
    if not name.startswith("word/"):
        name = "word/" + name
    try:
        return zf.read(name)
    except KeyError:
        return None


def parse_placeable_header(data: bytes) -> dict[str, Any] | None:
    if len(data) < 22 or u32(data, 0) != PLACEABLE_WMF_KEY:
        return None
    left, top, right, bottom, inch = struct.unpack_from("<hhhhH", data, 6)
    if inch <= 0:
        return None
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
        "unitsPerInch": inch,
        "widthPt": round3((right - left) / inch * 72.0),
        "heightPt": round3((bottom - top) / inch * 72.0),
    }


def decode_c_string(raw: bytes) -> str:
    raw = raw.split(b"\x00", 1)[0]
    return raw.decode("ascii", "replace")


def parse_create_font(
    payload: bytes,
    object_index: int,
    x_units_per_pt: float,
    y_units_per_pt: float,
) -> dict[str, Any]:
    if len(payload) < 50:
        return {"objectIndex": object_index, "error": "short_create_font"}
    height = s16(payload, 0)
    width = s16(payload, 2)
    return {
        "objectIndex": object_index,
        "heightTwips": height,
        "heightPt": round3(abs(height) / y_units_per_pt),
        "widthTwips": width,
        "widthPt": round3(abs(width) / x_units_per_pt) if width else 0.0,
        "escapement": s16(payload, 4),
        "orientation": s16(payload, 6),
        "weight": u16(payload, 8),
        "italic": payload[10],
        "underline": payload[11],
        "strikeout": payload[12],
        "charset": payload[13],
        "face": decode_c_string(payload[18:50]),
    }


def allocate_object_index(objects: dict[int, dict[str, Any]], free_indexes: list[int]) -> int:
    if free_indexes:
        return free_indexes.pop(0)
    index = 0
    while index in objects:
        index += 1
    return index


def parse_ext_text_rect(payload: bytes, offset: int, x_units_per_pt: float, y_units_per_pt: float) -> tuple[dict[str, Any] | None, int]:
    if offset + 8 > len(payload):
        return None, offset
    left = s16(payload, offset)
    top = s16(payload, offset + 2)
    right = s16(payload, offset + 4)
    bottom = s16(payload, offset + 6)
    return (
        {
            "leftPt": round3(left / x_units_per_pt),
            "topPt": round3(top / y_units_per_pt),
            "rightPt": round3(right / x_units_per_pt),
            "bottomPt": round3(bottom / y_units_per_pt),
            "raw": [left, top, right, bottom],
        },
        offset + 8,
    )


def decode_text_bytes(raw: bytes, charset: int) -> tuple[str, str]:
    encodings = {
        0: ("windows-1252", "latin1"),
        2: ("windows-1252", "latin1"),
        134: ("gb18030", "gbk", "windows-1252"),
    }.get(charset, ("gb18030", "windows-1252", "latin1"))
    best_text = ""
    best_encoding = encodings[0]
    best_score = -1
    for encoding in encodings:
        try:
            text = raw.decode(encoding, "replace")
        except LookupError:
            continue
        score = sum(1 for ch in text if ch != "\ufffd")
        if score > best_score:
            best_text = text
            best_encoding = encoding
            best_score = score
    return best_text, best_encoding


def byte_units_for_text(text: str, raw: bytes, encoding: str) -> list[dict[str, Any]]:
    units: list[dict[str, Any]] = []
    offset = 0
    for ch in text:
        encoded = ch.encode(encoding, "replace")
        size = max(1, len(encoded))
        if offset + size > len(raw):
            size = max(1, len(raw) - offset)
        units.append({"text": ch, "byteStart": offset, "byteLength": size})
        offset += size
        if offset >= len(raw):
            break
    if len(units) == len(raw):
        return units
    if sum(unit["byteLength"] for unit in units) == len(raw):
        return units
    return [
        {"text": f"0x{byte:02x}", "byteStart": i, "byteLength": 1}
        for i, byte in enumerate(raw)
    ]


def byte_dx_rows(raw: bytes, dx_values: list[int], x: int, x_units_per_pt: float) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    cursor = 0
    for index, byte in enumerate(raw):
        dx = dx_values[index] if index < len(dx_values) else 0
        rows.append(
            {
                "byteIndex": index,
                "rawHex": f"{byte:02x}",
                "dxTwips": dx,
                "dxPt": round3(dx / x_units_per_pt),
                "leftPt": round3((x + cursor) / x_units_per_pt),
                "rightPt": round3((x + cursor + dx) / x_units_per_pt),
            }
        )
        cursor += dx
    return rows


def parse_ext_text_out(
    payload: bytes,
    run_index: int,
    selected_font: int | None,
    fonts: dict[int, dict[str, Any]],
    x_units_per_pt: float,
    y_units_per_pt: float,
) -> dict[str, Any] | None:
    if len(payload) < 8:
        return None
    y = s16(payload, 0)
    x = s16(payload, 2)
    count = u16(payload, 4)
    options = u16(payload, 6)
    text_start = 8
    rect = None
    if options & (ETO_OPAQUE | ETO_CLIPPED):
        rect, text_start = parse_ext_text_rect(payload, text_start, x_units_per_pt, y_units_per_pt)
    text_end = min(text_start + count, len(payload))
    raw_text = payload[text_start:text_end]
    padded_text_len = count + (count & 1)
    dx_start = text_start + padded_text_len
    dx_values = [
        s16(payload, offset)
        for offset in range(dx_start, len(payload) - 1, 2)
    ]
    font = fonts.get(selected_font if selected_font is not None else -1, {})
    text, encoding = decode_text_bytes(raw_text, int(font.get("charset", 0)))
    byte_units = byte_units_for_text(text, raw_text, encoding)
    bytes_out = byte_dx_rows(raw_text, dx_values, x, x_units_per_pt)
    chars = []
    cursor = 0
    for char_index, unit in enumerate(byte_units):
        dx_slice = dx_values[unit["byteStart"]: unit["byteStart"] + unit["byteLength"]]
        char_dx = sum(dx_slice)
        chars.append(
            {
                "charIndex": char_index,
                "text": unit["text"],
                "byteStart": unit["byteStart"],
                "byteLength": unit["byteLength"],
                "dxTwips": char_dx,
                "dxPt": round3(char_dx / x_units_per_pt),
                "leftPt": round3((x + cursor) / x_units_per_pt),
                "rightPt": round3((x + cursor + char_dx) / x_units_per_pt),
            }
        )
        cursor += char_dx
    advance = sum(dx_values)
    return {
        "runIndex": run_index,
        "record": "ExtTextOut",
        "selectedFontObjectIndex": selected_font,
        "fontFace": font.get("face"),
        "fontHeightPt": font.get("heightPt"),
        "fontWidthPt": font.get("widthPt"),
        "charset": font.get("charset"),
        "xTwips": x,
        "yTwips": y,
        "xPt": round3(x / x_units_per_pt),
        "baselinePt": round3(y / y_units_per_pt),
        "options": options,
        "rect": rect,
        "text": text,
        "textEncoding": encoding,
        "rawTextHex": raw_text.hex(),
        "dxTwips": dx_values,
        "dxPt": [round3(value / x_units_per_pt) for value in dx_values],
        "advanceTwips": advance,
        "advanceWidthPt": round3(advance / x_units_per_pt),
        "rightEdgePt": round3((x + advance) / x_units_per_pt),
        "charCount": len(chars),
        "byteCount": len(raw_text),
        "chars": chars,
        "bytes": bytes_out,
    }


def parse_text_out(
    payload: bytes,
    run_index: int,
    selected_font: int | None,
    fonts: dict[int, dict[str, Any]],
    x_units_per_pt: float,
    y_units_per_pt: float,
) -> dict[str, Any] | None:
    if len(payload) < 6:
        return None
    count = u16(payload, 0)
    raw_text = payload[2: 2 + count]
    coord_offset = 2 + count + (count & 1)
    if coord_offset + 4 > len(payload):
        return None
    y = s16(payload, coord_offset)
    x = s16(payload, coord_offset + 2)
    font = fonts.get(selected_font if selected_font is not None else -1, {})
    text, encoding = decode_text_bytes(raw_text, int(font.get("charset", 0)))
    return {
        "runIndex": run_index,
        "record": "TextOut",
        "selectedFontObjectIndex": selected_font,
        "fontFace": font.get("face"),
        "fontHeightPt": font.get("heightPt"),
        "fontWidthPt": font.get("widthPt"),
        "charset": font.get("charset"),
        "xTwips": x,
        "yTwips": y,
        "xPt": round3(x / x_units_per_pt),
        "baselinePt": round3(y / y_units_per_pt),
        "text": text,
        "textEncoding": encoding,
        "rawTextHex": raw_text.hex(),
        "advanceTwips": None,
        "advanceWidthPt": None,
        "rightEdgePt": None,
        "charCount": len(text),
        "byteCount": len(raw_text),
        "chars": [],
        "bytes": [
            {
                "byteIndex": index,
                "rawHex": f"{byte:02x}",
                "dxTwips": None,
                "dxPt": None,
                "leftPt": None,
                "rightPt": None,
            }
            for index, byte in enumerate(raw_text)
        ],
    }


def parse_polyline(payload: bytes, x_units_per_pt: float, y_units_per_pt: float) -> dict[str, Any] | None:
    if len(payload) < 2:
        return None
    count = u16(payload, 0)
    points = []
    offset = 2
    for _ in range(count):
        if offset + 4 > len(payload):
            break
        x = s16(payload, offset)
        y = s16(payload, offset + 2)
        points.append({"xPt": round3(x / x_units_per_pt), "yPt": round3(y / y_units_per_pt)})
        offset += 4
    return {"pointCount": count, "points": points}


def scan_window_ext(data: bytes) -> tuple[int, int] | None:
    offset = 22 if parse_placeable_header(data) else 0
    if offset + 18 > len(data):
        return None
    offset += 18
    while offset + 6 <= len(data):
        size_words = u32(data, offset)
        function = u16(data, offset + 4)
        if size_words <= 0 or function == META_EOF:
            break
        payload = data[offset + 6:offset + size_words * 2]
        if function == META_SETWINDOWEXT and len(payload) >= 4:
            return s16(payload, 2), s16(payload, 0)
        offset += size_words * 2
    return None


def wmf_record(function: int, payload: bytes = b"") -> bytes:
    return struct.pack("<IH", 3 + (len(payload) + 1) // 2, function) + payload + (b"\x00" if len(payload) & 1 else b"")


def make_test_font(face: str = "Times New Roman", charset: int = 0) -> bytes:
    face_bytes = face.encode("ascii", "replace")[:31] + b"\x00"
    face_bytes = face_bytes.ljust(32, b"\x00")
    return (
        struct.pack("<hhhhHBBBBBBBB", -240, 0, 0, 0, 400, 0, 0, 0, charset, 0, 0, 0, 0)
        + face_bytes
    )


def build_self_test_wmf() -> bytes:
    records = [
        wmf_record(META_SETWINDOWEXT, struct.pack("<hh", 240, 400)),
        wmf_record(META_CREATEPENINDIRECT, struct.pack("<hhhI", 0, 1, 0, 0)),
        wmf_record(META_CREATEFONTINDIRECT, make_test_font()),
        wmf_record(META_SELECTOBJECT, struct.pack("<H", 1)),
        wmf_record(
            META_EXTTEXTOUT,
            struct.pack("<hhhhhhhh", 100, 20, 2, ETO_CLIPPED, 0, 0, 200, 120)
            + b"AB"
            + struct.pack("<hh", 120, 80),
        ),
        wmf_record(META_DELETEOBJECT, struct.pack("<H", 0)),
        wmf_record(META_EOF),
    ]
    body = b"".join(records)
    header = struct.pack("<HHHIHIH", 1, 9, 0x0300, (18 + len(body)) // 2, 2, 0, 0)
    return header + body


def run_self_test() -> None:
    wmf = parse_wmf(build_self_test_wmf(), 20.0, 20.0, 12.0)
    runs = wmf.get("runs") or []
    assert len(runs) == 1, f"expected one run, got {len(runs)}"
    run = runs[0]
    assert run.get("fontFace") == "Times New Roman", run
    assert run.get("selectedFontObjectIndex") == 1, run
    assert run.get("text") == "AB", run
    assert run.get("rect", {}).get("rightPt") == 10.0, run
    assert run.get("dxTwips") == [120, 80], run
    assert run.get("bytes", [])[0].get("rawHex") == "41", run
    assert run.get("bytes", [])[1].get("dxPt") == 4.0, run
    assert wmf.get("summary", {}).get("recordWidthPt") == 10.0, wmf


def scale_for_axis(raw_units: int | None, physical_pt: float | None, fallback_units_per_pt: float) -> tuple[float, str]:
    if raw_units and physical_pt and physical_pt > 0:
        return abs(raw_units) / physical_pt, "windowExtPhysical"
    return fallback_units_per_pt, "fallback"


def parse_wmf(data: bytes, fallback_units_per_pt: float, shape_width_pt: float | None, shape_height_pt: float | None) -> dict[str, Any]:
    placeable = parse_placeable_header(data)
    raw_window_ext = scan_window_ext(data)
    physical_width_pt = (placeable or {}).get("widthPt") or shape_width_pt
    physical_height_pt = (placeable or {}).get("heightPt") or shape_height_pt
    x_units_per_pt, x_scale_source = scale_for_axis(
        raw_window_ext[0] if raw_window_ext else None,
        physical_width_pt,
        fallback_units_per_pt,
    )
    y_units_per_pt, y_scale_source = scale_for_axis(
        raw_window_ext[1] if raw_window_ext else None,
        physical_height_pt,
        fallback_units_per_pt,
    )
    offset = 22 if placeable else 0
    header = {}
    if offset + 18 <= len(data):
        header = {
            "fileType": u16(data, offset),
            "headerSizeWords": u16(data, offset + 2),
            "version": u16(data, offset + 4),
            "fileSizeWords": u32(data, offset + 6),
            "objectCount": u16(data, offset + 10),
            "maxRecordSizeWords": u32(data, offset + 12),
        }
        offset += 18
    objects: dict[int, dict[str, Any]] = {}
    free_indexes: list[int] = []
    selected_font: int | None = None
    runs: list[dict[str, Any]] = []
    polylines: list[dict[str, Any]] = []
    window_ext = None
    record_count = 0

    while offset + 6 <= len(data):
        size_words = u32(data, offset)
        function = u16(data, offset + 4)
        if size_words <= 0:
            break
        end = offset + size_words * 2
        payload = data[offset + 6:end]
        record_count += 1
        if function == META_EOF:
            break
        if function == META_SETWINDOWEXT and len(payload) >= 4:
            height = s16(payload, 0)
            width = s16(payload, 2)
            window_ext = {
                "widthTwips": width,
                "heightTwips": height,
                "widthPt": round3(abs(width) / x_units_per_pt),
                "heightPt": round3(abs(height) / y_units_per_pt),
                "xUnitsPerPt": round3(x_units_per_pt),
                "yUnitsPerPt": round3(y_units_per_pt),
                "xScaleSource": x_scale_source,
                "yScaleSource": y_scale_source,
            }
        elif function == META_CREATEFONTINDIRECT:
            object_index = allocate_object_index(objects, free_indexes)
            font = parse_create_font(payload, object_index, x_units_per_pt, y_units_per_pt)
            font["type"] = "font"
            objects[object_index] = font
        elif function in (META_CREATEPENINDIRECT, META_CREATEBRUSHINDIRECT, META_CREATEPALETTE, META_CREATEREGION):
            object_index = allocate_object_index(objects, free_indexes)
            objects[object_index] = {"objectIndex": object_index, "type": "gdi", "function": function}
        elif function == META_DELETEOBJECT and len(payload) >= 2:
            object_index = u16(payload, 0)
            if object_index in objects:
                del objects[object_index]
                if object_index not in free_indexes:
                    free_indexes.append(object_index)
                    free_indexes.sort()
            if selected_font == object_index:
                selected_font = None
        elif function == META_SELECTOBJECT and len(payload) >= 2:
            object_index = u16(payload, 0)
            if objects.get(object_index, {}).get("type") == "font":
                selected_font = object_index
        elif function == META_EXTTEXTOUT:
            run = parse_ext_text_out(payload, len(runs), selected_font, objects, x_units_per_pt, y_units_per_pt)
            if run:
                runs.append(run)
        elif function == META_TEXTOUT:
            run = parse_text_out(payload, len(runs), selected_font, objects, x_units_per_pt, y_units_per_pt)
            if run:
                runs.append(run)
        elif function == META_POLYLINE:
            polyline = parse_polyline(payload, x_units_per_pt, y_units_per_pt)
            if polyline:
                polylines.append(polyline)
        offset = end

    right_edges = [run["rightEdgePt"] for run in runs if run.get("rightEdgePt") is not None]
    left_edges = [run["xPt"] for run in runs if run.get("xPt") is not None]
    baselines = [run["baselinePt"] for run in runs if run.get("baselinePt") is not None]
    return {
        "placeable": placeable,
        "metafileHeader": header,
        "windowExt": window_ext,
        "units": {
            "xUnitsPerPt": round3(x_units_per_pt),
            "yUnitsPerPt": round3(y_units_per_pt),
            "xScaleSource": x_scale_source,
            "yScaleSource": y_scale_source,
            "fallbackUnitsPerPt": fallback_units_per_pt,
        },
        "recordCount": record_count,
        "objects": [objects[index] for index in sorted(objects)],
        "fonts": [objects[index] for index in sorted(objects) if objects[index].get("type") == "font"],
        "runs": runs,
        "polylines": polylines,
        "summary": {
            "runCount": len(runs),
            "charCount": sum(int(run.get("charCount") or 0) for run in runs),
            "polylineCount": len(polylines),
            "advanceWidthPt": round3(sum(float(run.get("advanceWidthPt") or 0.0) for run in runs)),
            "recordLeftPt": round3(min(left_edges)) if left_edges else None,
            "recordRightPt": round3(max(right_edges)) if right_edges else None,
            "recordWidthPt": round3(max(right_edges) - min(left_edges)) if left_edges and right_edges else None,
            "baselineMinPt": round3(min(baselines)) if baselines else None,
            "baselineMaxPt": round3(max(baselines)) if baselines else None,
        },
    }


def render_media_to_png(media: bytes, suffix: str, magick: Path, temp_dir: Path) -> Path | None:
    input_path = temp_dir / f"input{suffix}"
    output_path = temp_dir / "output.png"
    input_path.write_bytes(media)
    if suffix.lower() == ".png":
        output_path.write_bytes(media)
        return output_path
    result = subprocess.run(
        [
            str(magick),
            "-density",
            "144",
            str(input_path),
            "-background",
            "white",
            "-alpha",
            "remove",
            "-alpha",
            "off",
            str(output_path),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0 or not output_path.exists():
        return None
    return output_path


def is_ink(pixel: tuple[int, ...], white_threshold: int, alpha_threshold: int) -> bool:
    if len(pixel) == 4 and pixel[3] <= alpha_threshold:
        return False
    return any(channel < white_threshold for channel in pixel[:3])


def measure_png_ink(
    png: Path,
    box_width_pt: float,
    box_height_pt: float,
    white_threshold: int,
    alpha_threshold: int,
) -> dict[str, Any]:
    if Image is None:
        return {"error": "pillow_not_available"}
    with Image.open(png).convert("RGBA") as image:
        width, height = image.size
        pixels = image.load()
        min_x, min_y = width, height
        max_x, max_y = -1, -1
        ink_count = 0
        for y in range(height):
            for x in range(width):
                if is_ink(pixels[x, y], white_threshold, alpha_threshold):
                    ink_count += 1
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)
        result: dict[str, Any] = {
            "imageWidthPx": width,
            "imageHeightPx": height,
            "inkCount": ink_count,
        }
        if max_x < 0:
            result.update({"magickInkBBoxPx": None, "magickInkWidthPt": 0.0, "magickInkHeightPt": 0.0})
            return result
        ink_width_px = max_x - min_x + 1
        ink_height_px = max_y - min_y + 1
        left_pt = min_x / max(width, 1) * box_width_pt
        top_pt = min_y / max(height, 1) * box_height_pt
        width_pt = ink_width_px / max(width, 1) * box_width_pt
        height_pt = ink_height_px / max(height, 1) * box_height_pt
        result.update(
            {
                "magickInkBBoxPx": [min_x, min_y, max_x + 1, max_y + 1],
                "magickInkLeftPt": round3(left_pt),
                "magickInkTopPt": round3(top_pt),
                "magickInkRightPt": round3(left_pt + width_pt),
                "magickInkBottomPt": round3(top_pt + height_pt),
                "magickInkWidthPt": round3(width_pt),
                "magickInkHeightPt": round3(height_pt),
                "magickInkCenterYPt": round3(top_pt + height_pt / 2.0),
            }
        )
        return result


def measure_ink(
    media: bytes | None,
    suffix: str,
    box_width_pt: float,
    box_height_pt: float,
    magick: Path | None,
    temp_root: Path,
    white_threshold: int,
    alpha_threshold: int,
) -> dict[str, Any]:
    if not media:
        return {"error": "missing_media"}
    if not magick:
        return {"error": "magick_not_available"}
    media_dir = temp_root / f"media-{len(list(temp_root.iterdir()))}"
    media_dir.mkdir()
    png = render_media_to_png(media, suffix, magick, media_dir)
    if png is None:
        return {"error": "render_failed"}
    return measure_png_ink(png, box_width_pt, box_height_pt, white_threshold, alpha_threshold)


TRACE_PREFIX_RE = re.compile(r"^\\pw(?:metrics|style)\{[^}]*}\s*")
TRACE_ASCII_WS = r"[ \t\n\x0b\f\r]"
TRACE_METRICS_RE = re.compile(r"^\\pwmetrics\{[^}]+}" + TRACE_ASCII_WS + "*")
TRACE_STYLE_RE = re.compile(r"^\\pwstyle\{[^}]*}" + TRACE_ASCII_WS + "*")
TRACE_EDGE_SPACE_RE = re.compile(r"^" + TRACE_ASCII_WS + r"+|" + TRACE_ASCII_WS + r"+$")
TRACE_SPACE_RE = re.compile(TRACE_ASCII_WS + "+")


def formula_trace_id(latex: str) -> str:
    value = (latex or "").replace("\u00a0", " ")
    value = TRACE_METRICS_RE.sub("", value)
    value = TRACE_STYLE_RE.sub("", value)
    value = TRACE_EDGE_SPACE_RE.sub("", value)
    value = TRACE_SPACE_RE.sub(" ", value)
    return "pwf:" + hashlib.sha256(value.encode("utf-8")).hexdigest()[:16]


def math_bodies(text: str) -> list[str]:
    return [
        (match.group(1) or match.group(2) or "").strip()
        for match in re.finditer(r"\$\$(.+?)\$\$|\$(.+?)\$", text, re.S)
    ]


def strip_formula_prefixes(latex: str) -> str:
    value = latex.strip()
    previous = None
    while value != previous:
        previous = value
        value = TRACE_PREFIX_RE.sub("", value).strip()
    return value


def request_math_sequence(request: dict[str, Any]) -> list[str]:
    fields = ("content", "analyze", "solution", "correct", "difficulty", "knowledgePoint")
    sequence: list[str] = []
    for section in request.get("sections", []):
        for question in section.get("questions", []):
            ordered_fields = question.get("_mathOrder")
            if ordered_fields is not None:
                ordered_values = {str(item or "") for item in ordered_fields}
                for item in ordered_fields:
                    sequence.extend(math_bodies(str(item or "")))
                for field in fields:
                    if field == "content" or str(question.get(field) or "") in ordered_values:
                        continue
                    sequence.extend(math_bodies(str(question.get(field) or "")))
            else:
                for field in fields:
                    sequence.extend(math_bodies(str(question.get(field) or "")))
            for tag in question.get("tags") or []:
                sequence.extend(math_bodies(str(tag)))
    return sequence


def extract_request_formulas(request_path: Path | None) -> list[dict[str, Any]]:
    if not request_path:
        return []
    request = json.loads(request_path.read_text(encoding="utf-8"))
    formulas = []
    for raw in request_math_sequence(request):
        stripped = strip_formula_prefixes(raw)
        formulas.append(
            {
                "formula": stripped,
                "rawFormula": raw,
                "formulaKey": normalize_latex_key(stripped),
                "traceId": formula_trace_id(raw),
            }
        )
    return formulas


def unique_request_formula_by_trace(request_formulas: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    by_trace: dict[str, list[dict[str, Any]]] = {}
    for formula in request_formulas:
        trace = formula.get("traceId")
        if trace:
            by_trace.setdefault(str(trace), []).append(formula)
    return {
        trace: formulas[0]
        for trace, formulas in by_trace.items()
        if len(formulas) == 1
    }


def docx_trace_counts(boxes: list[Any]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for box in boxes:
        trace = str(getattr(box, "image_title", "") or "")
        if trace.startswith("pwf:"):
            counts[trace] = counts.get(trace, 0) + 1
    return counts


def inspect_docx(args: argparse.Namespace) -> dict[str, Any]:
    indexes = parse_index_list(args.indices)
    request_formulas = extract_request_formulas(args.request)
    request_formula_by_trace = unique_request_formula_by_trace(request_formulas)
    magick: Path | None = None
    if args.with_ink:
        if args.magick.exists():
            magick = args.magick
        elif shutil.which("magick"):
            magick = Path(shutil.which("magick") or "")

    boxes = extract_boxes(args.docx)
    if indexes is not None:
        boxes = [box for box in boxes if box.index + 1 in indexes]
    if args.max_items is not None:
        boxes = boxes[: args.max_items]
    generated_trace_counts = docx_trace_counts(boxes)

    formulas_out: list[dict[str, Any]] = []
    char_rows: list[dict[str, Any]] = []
    byte_rows: list[dict[str, Any]] = []
    run_rows: list[dict[str, Any]] = []
    with zipfile.ZipFile(args.docx) as zf, tempfile.TemporaryDirectory(prefix="wmf-glyphs-") as tmp:
        temp_root = Path(tmp)
        for box in boxes:
            media = read_docx_target(zf, box.image_target)
            suffix = Path(box.image_target or "").suffix.lower()
            wmf = (
                parse_wmf(media or b"", args.fallback_units_per_pt, box.style_width_pt, box.style_height_pt)
                if suffix == ".wmf" and media
                else None
            )
            ink = (
                measure_ink(
                    media,
                    suffix or ".bin",
                    box.style_width_pt,
                    box.style_height_pt,
                    magick,
                    temp_root,
                    args.white_threshold,
                    args.alpha_threshold,
                )
                if args.with_ink
                else None
            )
            image_title = str(box.image_title or "")
            if image_title.startswith("pwf:"):
                if generated_trace_counts.get(image_title) == 1 and image_title in request_formula_by_trace:
                    request_formula = request_formula_by_trace[image_title]
                    formula_match_method = "trace_unique"
                elif generated_trace_counts.get(image_title, 0) > 1:
                    request_formula = {}
                    formula_match_method = "trace_duplicate"
                else:
                    request_formula = {}
                    formula_match_method = "trace_unmatched"
            else:
                request_formula = request_formulas[box.index] if box.index < len(request_formulas) else {}
                formula_match_method = "ordinal_fallback" if request_formula else "absent"
            outer = {
                "objectIndex": box.index + 1,
                "zeroBasedIndex": box.index,
                "formula": request_formula.get("formula"),
                "rawFormula": request_formula.get("rawFormula"),
                "formulaKey": request_formula.get("formulaKey"),
                "formulaTraceId": image_title if image_title.startswith("pwf:") else request_formula.get("traceId"),
                "formulaMatchMethod": formula_match_method,
                "context": box.context,
                "imageTarget": box.image_target,
                "oleTarget": box.ole_target,
                "shapeWidthPt": box.style_width_pt,
                "shapeHeightPt": box.style_height_pt,
                "dxaOrigPt": box.dxa_orig_pt,
                "dyaOrigPt": box.dya_orig_pt,
                "wmfPlaceableWidthPt": (wmf or {}).get("placeable", {}).get("widthPt") if wmf else None,
                "wmfPlaceableHeightPt": (wmf or {}).get("placeable", {}).get("heightPt") if wmf else None,
                "wmfWindowWidthPt": (wmf or {}).get("windowExt", {}).get("widthPt") if wmf else None,
                "wmfWindowHeightPt": (wmf or {}).get("windowExt", {}).get("heightPt") if wmf else None,
                "wmfXUnitsPerPt": (wmf or {}).get("units", {}).get("xUnitsPerPt") if wmf else None,
                "wmfYUnitsPerPt": (wmf or {}).get("units", {}).get("yUnitsPerPt") if wmf else None,
                "wmfXScaleSource": (wmf or {}).get("units", {}).get("xScaleSource") if wmf else None,
                "wmfYScaleSource": (wmf or {}).get("units", {}).get("yScaleSource") if wmf else None,
                "magickInkWidthPt": (ink or {}).get("magickInkWidthPt") if ink else None,
                "magickInkHeightPt": (ink or {}).get("magickInkHeightPt") if ink else None,
                "magickInkError": (ink or {}).get("error") if ink else None,
            }
            runs = (wmf or {}).get("runs", [])
            for run in runs:
                row = {
                    **outer,
                    "runIndex": run["runIndex"],
                    "fontFace": run.get("fontFace"),
                    "fontHeightPt": run.get("fontHeightPt"),
                    "fontWidthPt": run.get("fontWidthPt"),
                    "charset": run.get("charset"),
                    "xPt": run.get("xPt"),
                    "baselinePt": run.get("baselinePt"),
                    "text": run.get("text"),
                    "advanceWidthPt": run.get("advanceWidthPt"),
                    "rightEdgePt": run.get("rightEdgePt"),
                    "byteCount": run.get("byteCount"),
                    "charCount": run.get("charCount"),
                }
                run_rows.append(row)
                for ch in run.get("chars", []):
                    char_rows.append(
                        {
                            **row,
                            "charIndex": ch.get("charIndex"),
                            "char": ch.get("text"),
                            "charDxPt": ch.get("dxPt"),
                            "charLeftPt": ch.get("leftPt"),
                            "charRightPt": ch.get("rightPt"),
                            "byteStart": ch.get("byteStart"),
                            "byteLength": ch.get("byteLength"),
                        }
                    )
                for byte in run.get("bytes", []):
                    byte_rows.append(
                        {
                            **row,
                            "byteIndex": byte.get("byteIndex"),
                            "rawHex": byte.get("rawHex"),
                            "byteDxPt": byte.get("dxPt"),
                            "byteLeftPt": byte.get("leftPt"),
                            "byteRightPt": byte.get("rightPt"),
                        }
                    )
            formulas_out.append(
                {
                    **outer,
                    "box": asdict(box),
                    "wmf": wmf,
                    "ink": ink,
                }
            )

    return {
        "docx": str(args.docx.resolve()),
        "fallbackUnitsPerPt": args.fallback_units_per_pt,
        "withInk": args.with_ink,
        "pillowAvailable": Image is not None,
        "magick": str(magick) if magick else None,
        "formulaCount": len(formulas_out),
        "runCount": len(run_rows),
        "charCount": len(char_rows),
        "byteCount": len(byte_rows),
        "formulas": formulas_out,
        "runs": run_rows,
        "chars": char_rows,
        "bytes": byte_rows,
    }


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    keys: list[str] = []
    for row in rows:
        for key in row:
            if key not in keys:
                keys.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=keys)
        writer.writeheader()
        writer.writerows(rows)


def render_text(report: dict[str, Any], limit: int) -> str:
    lines = [
        "WMF formula glyph metrics",
        f"docx: {report['docx']}",
        f"formulas: {report['formulaCount']} runs: {report['runCount']} chars: {report['charCount']} bytes: {report['byteCount']}",
        "ink: {mode} magick: {magick} pillow: {pillow}".format(
            mode="enabled" if report["withInk"] else "disabled",
            magick=report.get("magick") or "",
            pillow="available" if report.get("pillowAvailable") else "missing",
        ),
        "",
    ]
    formulas = report["formulas"][:limit]
    for item in formulas:
        summary = (item.get("wmf") or {}).get("summary") or {}
        ink_bits = ""
        if item.get("ink"):
            ink = item["ink"]
            if ink.get("error"):
                ink_bits = f" magickInk_error={ink['error']}"
            else:
                ink_bits = f" magickInk={ink.get('magickInkWidthPt')}x{ink.get('magickInkHeightPt')}pt"
        lines.append(
            "#{idx} shape={sw}x{sh}pt wmf={ww}x{wh}pt recordWidth={rw}pt units={ux}/{uy} runs={runs} match={match} trace={trace}{ink} text={text}".format(
                idx=item["objectIndex"],
                sw=item["shapeWidthPt"],
                sh=item["shapeHeightPt"],
                ww=item.get("wmfPlaceableWidthPt"),
                wh=item.get("wmfPlaceableHeightPt"),
                rw=summary.get("recordWidthPt"),
                ux=item.get("wmfXUnitsPerPt"),
                uy=item.get("wmfYUnitsPerPt"),
                runs=summary.get("runCount"),
                match=item.get("formulaMatchMethod"),
                trace=item.get("formulaTraceId"),
                ink=ink_bits,
                text=(item.get("formula") or item.get("context") or "")[:120],
            )
        )
        for run in ((item.get("wmf") or {}).get("runs") or [])[:8]:
            lines.append(
                "  run {idx}: face={face} h={height}pt x={x} base={base} adv={adv} right={right} text={text}".format(
                    idx=run.get("runIndex"),
                    face=run.get("fontFace"),
                    height=run.get("fontHeightPt"),
                    x=run.get("xPt"),
                    base=run.get("baselinePt"),
                    adv=run.get("advanceWidthPt"),
                    right=run.get("rightEdgePt"),
                    text=repr(run.get("text") or ""),
                )
            )
    if report["formulaCount"] > limit:
        lines.append(f"... {report['formulaCount'] - limit} more formulas omitted")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path, nargs="?")
    parser.add_argument("--self-test", action="store_true", help="Run parser self-checks and exit.")
    parser.add_argument("--request", type=Path)
    parser.add_argument("--indices", help="Comma-separated 1-based object indexes.")
    parser.add_argument("--max-items", type=int)
    parser.add_argument("--with-ink", action="store_true")
    parser.add_argument("--require-magick-ink", action="store_true")
    parser.add_argument("--magick", type=Path, default=DEFAULT_MAGICK)
    parser.add_argument("--white-threshold", type=int, default=245)
    parser.add_argument("--alpha-threshold", type=int, default=8)
    parser.add_argument("--fallback-units-per-pt", type=float, default=20.0)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-runs-csv", type=Path)
    parser.add_argument("--out-chars-csv", type=Path)
    parser.add_argument("--out-bytes-csv", type=Path)
    parser.add_argument("--out-text", type=Path)
    parser.add_argument("--text-limit", type=int, default=20)
    args = parser.parse_args()

    if args.self_test:
        run_self_test()
        print("measure_wmf_formula_glyphs self-test passed")
        return 0

    if args.docx is None:
        parser.error("docx is required unless --self-test is used")

    report = inspect_docx(args)
    if args.require_magick_ink:
        errors = [
            item.get("magickInkError")
            for item in report["formulas"]
            if item.get("magickInkError")
        ]
        if errors:
            raise SystemExit("magick ink measurement required but failed: " + ", ".join(sorted(set(errors))))
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_runs_csv:
        write_csv(args.out_runs_csv, report["runs"])
    if args.out_chars_csv:
        write_csv(args.out_chars_csv, report["chars"])
    if args.out_bytes_csv:
        write_csv(args.out_bytes_csv, report["bytes"])
    text = render_text(report, args.text_limit)
    if args.out_text:
        args.out_text.parent.mkdir(parents=True, exist_ok=True)
        args.out_text.write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
