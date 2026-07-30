# -*- coding: utf-8 -*-
"""
Parse MathType WMF preview records into a glyph-level layout JSON.

A genuine MathType preview WMF is vector: it carries META_CREATEFONTINDIRECT,
META_SELECTOBJECT, META_EXTTEXTOUT / META_TEXTOUT and drawing records with
exact logical coordinates.  This script walks the record stream and extracts

- glyphs:  character, font name, charset, size, anchor position, alignment
- rules:   rectangles / polygons / polylines (fraction bars, radicals, ...)
- metrics: placeable bbox, window org/ext, units-per-inch mapping

Coordinates are reported both in raw logical units and in Word points
(1 pt = 1/72 inch), so the output is directly comparable with the DOCX
display box (w:extent / w:position) and with MathJax SVG glyph positions.

Usage:
    python scripts/wmf_glyph_layout.py file.wmf            # single file -> JSON stdout
    python scripts/wmf_glyph_layout.py --docx file.docx --out out.json
    python scripts/wmf_glyph_layout.py --docx file.docx --out out.json --svg-dir svgs
"""
from __future__ import annotations

import argparse
import json
import struct
import sys
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Symbol font byte -> Unicode mapping (Adobe Symbol encoding, common subset)
# ---------------------------------------------------------------------------
SYMBOL_MAP: dict[int, str] = {
    0x22: "∀", 0x24: "∃", 0x27: "∋", 0x2A: "∗",
    0x40: "≅", 0x5C: "∴", 0x5E: "⊥", 0x60: "‾",
    0x7E: "∼",
    0xA1: "ϒ", 0xA2: "′", 0xA3: "≤", 0xA4: "⁄", 0xA5: "∞", 0xA6: "ƒ",
    0xA7: "♣", 0xA8: "♦", 0xA9: "♥", 0xAA: "♠",
    0xAB: "↔", 0xAC: "←", 0xAD: "↑", 0xAE: "→", 0xAF: "↓",
    0xB0: "°", 0xB1: "±", 0xB2: "″", 0xB3: "≥", 0xB4: "×", 0xB5: "∝",
    0xB6: "∂", 0xB7: "•", 0xB8: "÷", 0xB9: "≠", 0xBA: "≡", 0xBB: "≈",
    0xBC: "…", 0xBD: "⏐", 0xBE: "⎯", 0xBF: "↵",
    0xC0: "ℵ", 0xC1: "ℑ", 0xC2: "ℜ", 0xC3: "℘",
    0xC4: "⊗", 0xC5: "⊕", 0xC6: "∅", 0xC7: "∩", 0xC8: "∪",
    0xC9: "⊃", 0xCA: "⊇", 0xCB: "⊄", 0xCC: "⊂", 0xCD: "⊆",
    0xCE: "∈", 0xCF: "∉",
    0xD0: "∠", 0xD1: "∇", 0xD5: "∏", 0xD6: "√", 0xD7: "⋅",
    0xD8: "¬", 0xD9: "∧", 0xDA: "∨",
    0xDB: "⇔", 0xDC: "⇐", 0xDD: "⇑", 0xDE: "⇒", 0xDF: "⇓",
    0xE0: "◊", 0xE1: "〈", 0xE5: "∑",
    0xE6: "⎛", 0xE7: "⎜", 0xE8: "⎝", 0xE9: "⎡", 0xEA: "⎢", 0xEB: "⎣",
    0xEC: "⎧", 0xED: "⎨", 0xEE: "⎩", 0xEF: "⎪",
    0xF1: "〉", 0xF2: "∫", 0xF3: "⌠", 0xF4: "⎮", 0xF5: "⌡",
    0xF6: "⎞", 0xF7: "⎟", 0xF8: "⎠", 0xF9: "⎤", 0xFA: "⎥", 0xFB: "⎦",
    0xFC: "⎫", 0xFD: "⎬", 0xFE: "⎭",
}
# uppercase Greek 0x41..0x5A, lowercase Greek 0x61..0x7A
for _i, _ch in enumerate("ABXDEFGHIKLMNOPQRSTUVWYZ".replace(" ", "")):
    pass
_UPPER_GREEK = "ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ"
_LOWER_GREEK = "αβχδεφγηιϕκλμνοπθρστυϖωξψζ"
for _i in range(26):
    SYMBOL_MAP[0x41 + _i] = _UPPER_GREEK[_i]
    SYMBOL_MAP[0x61 + _i] = _LOWER_GREEK[_i]

CHARSET_NAMES = {
    0x00: "ANSI", 0x01: "DEFAULT", 0x02: "SYMBOL", 0x80: "SHIFTJIS",
    0x81: "HANGUL", 0x86: "GB2312", 0x88: "CHINESEBIG5", 0xFF: "OEM",
}

# MT Extra private encoding (reverse-engineered from corpus context)
MT_EXTRA_MAP: dict[int, str] = {
    0x4C: "⋯",  # midline horizontal ellipsis (confirmed vs MathJax \cdots)
}

# Symbol extensible-delimiter pieces -> logical delimiter
EXT_PIECE_TO_CHAR = {
    "⎛": "(", "⎜": "(", "⎝": "(",
    "⎞": ")", "⎟": ")", "⎠": ")",
    "⎡": "[", "⎢": "[", "⎣": "[",
    "⎤": "]", "⎥": "]", "⎦": "]",
    "⎧": "{", "⎨": "{", "⎩": "{", "⎪": "{",
    "⎫": "}", "⎬": "}", "⎭": "}",
    "⌠": "∫", "⎮": "∫", "⌡": "∫",
}

# WMF record function codes
FN_SETWINDOWEXT = 0x020C
FN_SETWINDOWORG = 0x020B
FN_SETVIEWPORTEXT = 0x020E
FN_SETVIEWPORTORG = 0x020D
FN_SETMAPMODE = 0x0203
FN_SETTEXTALIGN = 0x012E
FN_SETTEXTCOLOR = 0x0209
FN_SELECTOBJECT = 0x012D
FN_DELETEOBJECT = 0x01F0
FN_CREATEFONTINDIRECT = 0x02FB
FN_CREATEBRUSHINDIRECT = 0x02FC
FN_CREATEPENINDIRECT = 0x02FA
FN_EXTTEXTOUT = 0x0A32
FN_TEXTOUT = 0x0521
FN_RECTANGLE = 0x041B
FN_POLYGON = 0x0324
FN_POLYLINE = 0x0325
FN_MOVETO = 0x0214
FN_LINETO = 0x0213
FN_ESCAPE = 0x0626

TA_UPDATECP = 0x0001


def extract_mtef_from_comment(payload: bytes) -> bytes | None:
    """Extract MTEF bytes from a MathType MFCOMMENT escape payload.

    Two known layouts:
    - "AppsMFCC"(8) flag(2) size1(4) size2(4) vendor\0 MTEF[size1]
    - "MathTypeUU"(10) unk(2) MTEF[remainder]   (Unicode-era MathType)
    """
    if payload.startswith(b"AppsMFCC") and len(payload) >= 18:
        p = 8
        p += 2  # flag
        size1 = struct.unpack_from("<I", payload, p)[0]
        p += 4
        p += 4  # size2 (duplicate of size1)
        # vendor zero-terminated string, e.g. "Design Science, Inc."
        end = payload.find(b"\x00", p)
        if end < 0:
            return None
        p = end + 1
        n = size1
        if p + n > len(payload):
            n = len(payload) - p
        mtef = payload[p: p + n]
        return mtef if mtef and mtef[0] == 0x05 else None
    if payload.startswith(b"MathTypeUU") and len(payload) > 12:
        mtef = payload[12:]
        return mtef if mtef and mtef[0] == 0x05 else None
    return None

PLACEABLE_KEY = 0x9AC6CDD7


def _i16(b: bytes, o: int) -> int:
    return struct.unpack_from("<h", b, o)[0]


def _u16(b: bytes, o: int) -> int:
    return struct.unpack_from("<H", b, o)[0]


def _u32(b: bytes, o: int) -> int:
    return struct.unpack_from("<I", b, o)[0]


def decode_text(raw: bytes, charset: int, face: str = "") -> list[dict]:
    """Decode a WMF text byte string into per-glyph dicts.

    Returns a list of {ch, hex} entries.  Font face wins over charset byte:
    MathType writes Symbol with charset DEFAULT, so "Symbol" face is the
    reliable signal.  GB2312 decodes as GBK byte pairs; else cp1252.
    """
    out: list[dict] = []
    face_l = face.lower()
    use_symbol = "symbol" in face_l or charset == 0x02
    if not use_symbol and (charset == 0x86 or face_l in ("system", "宋体", "simsum")):
        # GB2312 / DBCS
        i = 0
        while i < len(raw):
            b = raw[i]
            if b < 0x80:
                out.append({"ch": chr(b), "hex": f"{b:02X}"})
                i += 1
            else:
                pair = raw[i:i + 2]
                try:
                    ch = pair.decode("gbk")
                except UnicodeDecodeError:
                    ch = ""
                out.append({"ch": ch, "hex": pair.hex().upper()})
                i += 2
        return out
    for b in raw:
        if "mt extra" in face_l:
            mapped = MT_EXTRA_MAP.get(b)
            out.append({"ch": mapped if mapped is not None else f"<MT{b:02X}>",
                        "hex": f"{b:02X}"})
        elif use_symbol:
            mapped = SYMBOL_MAP.get(b)
            if mapped is None and 0x20 <= b <= 0x7E:
                # Symbol keeps ASCII positions for digits and basic punctuation
                mapped = chr(b)
            out.append({"ch": mapped if mapped is not None else f"<0x{b:02X}>",
                        "hex": f"{b:02X}"})
        else:
            try:
                ch = bytes([b]).decode("cp1252")
            except UnicodeDecodeError:
                ch = f"<0x{b:02X}>"
            out.append({"ch": ch, "hex": f"{b:02X}"})
    return out


def parse_wmf(data: bytes) -> dict:
    result: dict = {
        "glyphs": [],
        "rules": [],
        "fonts": [],
        "warnings": [],
    }
    off = 0
    units_per_inch = None
    if len(data) >= 22 and _u32(data, 0) == PLACEABLE_KEY:
        # key(4) hmf(2) left top right bottom(2 each) inch(2) reserved(4) checksum(2)
        left, top, right, bottom = (_i16(data, 6 + 2 * i) for i in range(4))
        units_per_inch = _u16(data, 14)
        result["placeable_bbox_units"] = [left, top, right, bottom]
        result["units_per_inch"] = units_per_inch
        off = 22
    if len(data) < off + 18:
        result["warnings"].append("truncated header")
        return result
    header_size = _u16(data, off + 2)
    version = _u16(data, off + 4)
    result["wmf_version"] = hex(version)
    pos = off + header_size * 2

    window_org = [0, 0]
    window_ext = [0, 0]
    text_align = 0
    cp = [0, 0]  # current position, set by MOVETO, used when TA_UPDATECP is on
    object_table: dict[int, dict] = {}
    current_font: dict | None = None

    def alloc_slot() -> int:
        for i in range(256):
            if i not in object_table:
                return i
        return -1

    while pos + 6 <= len(data):
        rec_size_words = _u32(data, pos)
        func = _u16(data, pos + 4)
        rec_bytes = rec_size_words * 2
        if rec_size_words < 3 or pos + rec_bytes > len(data) + 2:
            result["warnings"].append(f"bad record at {pos}")
            break
        params = data[pos + 6: pos + rec_bytes]
        pos += rec_bytes
        if func == 0x0000:  # EOF
            break

        if func == FN_SETWINDOWORG and len(params) >= 4:
            window_org = [_i16(params, 2), _i16(params, 0)]  # params: y, x
        elif func == FN_SETWINDOWEXT and len(params) >= 4:
            window_ext = [_i16(params, 2), _i16(params, 0)]
        elif func == FN_SETTEXTALIGN and len(params) >= 2:
            text_align = _u16(params, 0)
        elif func == FN_MOVETO and len(params) >= 4:
            cp = [_i16(params, 2), _i16(params, 0)]  # params: y, x
        elif func == FN_LINETO and len(params) >= 4:
            to = [_i16(params, 2), _i16(params, 0)]
            result["rules"].append({"kind": "line", "x1": cp[0], "y1": cp[1],
                                    "x2": to[0], "y2": to[1]})
            cp = to
        elif func == FN_CREATEFONTINDIRECT and len(params) >= 50:
            lf_height = _i16(params, 0)
            lf_width = _i16(params, 2)
            lf_escapement = _i16(params, 4)
            lf_weight = _i16(params, 8)
            charset = params[13]
            face_raw = params[18:50].split(b"\x00", 1)[0]
            if any(b >= 0x80 for b in face_raw):
                # CJK face names (e.g. 宋体) are stored in the system codepage
                try:
                    face = face_raw.decode("gbk")
                except UnicodeDecodeError:
                    face = face_raw.decode("cp1252", errors="replace")
            else:
                face = face_raw.decode("cp1252", errors="replace")
            font = {
                "height_units": lf_height,
                "width_units": lf_width,
                "escapement": lf_escapement,
                "weight": lf_weight,
                "italic": bool(params[10]),
                "charset": charset,
                "charset_name": CHARSET_NAMES.get(charset, hex(charset)),
                "face": face,
            }
            slot = alloc_slot()
            object_table[slot] = {"kind": "font", **font}
            result["fonts"].append(font)
        elif func in (FN_CREATEBRUSHINDIRECT, FN_CREATEPENINDIRECT):
            slot = alloc_slot()
            object_table[slot] = {"kind": "brush" if func == FN_CREATEBRUSHINDIRECT else "pen"}
        elif func == FN_SELECTOBJECT and len(params) >= 2:
            idx = _u16(params, 0)
            obj = object_table.get(idx)
            if obj and obj.get("kind") == "font":
                current_font = obj
        elif func == FN_DELETEOBJECT and len(params) >= 2:
            object_table.pop(_u16(params, 0), None)
        elif func == FN_EXTTEXTOUT and len(params) >= 8:
            y = _i16(params, 0)
            x = _i16(params, 2)
            count = _u16(params, 4)
            opts = _u16(params, 6)
            p = 8
            if opts & 0x0006:  # ETO_OPAQUE | ETO_CLIPPED -> rect present
                p += 8
            raw = params[p: p + count]
            p += count + (count & 1)
            dx_array = None
            if len(params) >= p + count * 2 and count > 0:
                dx_array = [_i16(params, p + 2 * i) for i in range(count)]
            charset = current_font["charset"] if current_font else 0
            face = current_font["face"] if current_font else ""
            glyphs = decode_text(raw, charset, face)
            # MathType writes TA_UPDATECP: record x/y are ignored and the
            # string anchors at the current position set by MOVETO.
            if text_align & TA_UPDATECP:
                x, y = cp[0], cp[1]
            # dx[i] is the advance from glyph i's origin to glyph i+1's origin
            # (the last entry is the advance of the final glyph).
            cursor = x
            for gi, g in enumerate(glyphs):
                entry = {
                    "ch": g["ch"],
                    "hex": g["hex"],
                    "font": current_font["face"] if current_font else None,
                    "italic": current_font.get("italic") if current_font else None,
                    "charset": current_font["charset_name"] if current_font else None,
                    "size_units": current_font["height_units"] if current_font else None,
                    "x": cursor if dx_array else x,
                    "y": y,
                    "text_align": text_align,
                }
                if dx_array and gi < len(dx_array):
                    entry["dx"] = dx_array[gi]
                result["glyphs"].append(entry)
                if dx_array and gi < len(dx_array):
                    cursor += dx_array[gi]
            if text_align & TA_UPDATECP:
                cp = [cursor, y]
        elif func == FN_TEXTOUT and len(params) >= 2:
            count = _u16(params, 0)
            raw = params[2: 2 + count]
            p = 2 + count + (count & 1)
            if len(params) >= p + 4:
                y = _i16(params, p)
                x = _i16(params, p + 2)
                if text_align & TA_UPDATECP:
                    x, y = cp[0], cp[1]
                charset = current_font["charset"] if current_font else 0
                face = current_font["face"] if current_font else ""
                for g in decode_text(raw, charset, face):
                    result["glyphs"].append({
                        "ch": g["ch"], "hex": g["hex"],
                        "font": current_font["face"] if current_font else None,
                        "charset": current_font["charset_name"] if current_font else None,
                        "size_units": current_font["height_units"] if current_font else None,
                        "x": x, "y": y, "text_align": text_align,
                    })
        elif func == FN_ESCAPE and len(params) >= 4:
            esc_id = _u16(params, 0)
            dlen = _u16(params, 2)
            payload = params[4: 4 + dlen]
            if esc_id == 15:  # MFCOMMENT
                mtef = extract_mtef_from_comment(payload)
                if mtef:
                    result["mtef_hex"] = mtef.hex()
                    result["mtef_size"] = len(mtef)
        elif func == FN_RECTANGLE and len(params) >= 8:
            bottom, right, top, left = (_i16(params, 2 * i) for i in range(4))
            result["rules"].append({"kind": "rect", "left": left, "top": top,
                                    "right": right, "bottom": bottom})
        elif func in (FN_POLYGON, FN_POLYLINE) and len(params) >= 2:
            n = _u16(params, 0)
            pts = []
            for i in range(n):
                if 2 + 4 * i + 4 <= len(params):
                    pts.append([_i16(params, 2 + 4 * i), _i16(params, 2 + 4 * i + 2)])
            result["rules"].append({
                "kind": "polygon" if func == FN_POLYGON else "polyline",
                "points": pts,
            })

    result["window_org"] = window_org
    result["window_ext"] = window_ext

    # ---- merge extensible delimiter assemblies ------------------------------
    # MathType draws tall delimiters as stacked Symbol pieces (⎛⎜⎝ / ⎞⎟⎠ / ...)
    # sharing one x.  Collapse each stack into one logical delimiter glyph so
    # identity/geometry diffs compare delimiters, not pieces.
    pieces = [g for g in result["glyphs"] if g["ch"] in EXT_PIECE_TO_CHAR]
    if pieces:
        solid = [g for g in result["glyphs"] if g["ch"] not in EXT_PIECE_TO_CHAR]
        pieces.sort(key=lambda g: (g["x"], g["y"]))
        stacks: list[list[dict]] = []
        for g in pieces:
            tol_x = max(3, abs(g.get("size_units") or 300) * 0.02)
            if stacks and abs(stacks[-1][-1]["x"] - g["x"]) <= tol_x:
                stacks[-1].append(g)
            else:
                stacks.append([g])
        merged = []
        for stack in stacks:
            stack.sort(key=lambda g: g["y"])
            runs: list[list[dict]] = [[stack[0]]]
            for g in stack[1:]:
                gap = g["y"] - runs[-1][-1]["y"]
                max_gap = abs(g.get("size_units") or 300) * 1.5
                if gap <= max_gap and EXT_PIECE_TO_CHAR[g["ch"]] == EXT_PIECE_TO_CHAR[runs[-1][-1]["ch"]]:
                    runs[-1].append(g)
                else:
                    runs.append([g])
            for run in runs:
                top, bottom = run[0]["y"], run[-1]["y"]
                merged.append({
                    **run[0],
                    "ch": EXT_PIECE_TO_CHAR[run[0]["ch"]],
                    "x": sum(g["x"] for g in run) // len(run),
                    "y": (top + bottom) // 2,
                    "assembly": True,
                    "assembly_pieces": len(run),
                    "assembly_span": [top, bottom],
                })
        result["glyphs"] = solid + merged

    # ---- convert to points -------------------------------------------------
    scale = None
    if units_per_inch:
        scale = 72.0 / units_per_inch
    elif window_ext[0] and result.get("placeable_bbox_units"):
        bbox = result["placeable_bbox_units"]
        # fallback: assume window ext == bbox size
        scale = None
    if scale:
        result["pt_per_unit"] = scale
        bbox = result.get("placeable_bbox_units")
        if bbox:
            result["placeable_bbox_pt"] = [round(v * scale, 3) for v in bbox]
        for g in result["glyphs"]:
            g["x_pt"] = round(g["x"] * scale, 3)
            g["y_pt"] = round(g["y"] * scale, 3)
            if g.get("size_units"):
                g["size_pt"] = round(abs(g["size_units"]) * scale, 3)
            if g.get("dx") is not None:
                g["dx_pt"] = round(g["dx"] * scale, 3)
        for r in result["rules"]:
            if r["kind"] == "rect":
                for k in ("left", "top", "right", "bottom"):
                    r[k + "_pt"] = round(r[k] * scale, 3)
            elif r["kind"] == "line":
                for k in ("x1", "y1", "x2", "y2"):
                    r[k + "_pt"] = round(r[k] * scale, 3)
            else:
                r["points_pt"] = [[round(px * scale, 3), round(py * scale, 3)]
                                  for px, py in r["points"]]
    else:
        result["warnings"].append("no units-per-inch; pt conversion skipped")

    # reconstructed reading text: group by baseline y (3-unit tolerance), sort by x
    lines: dict[int, list[dict]] = {}
    for g in sorted(result["glyphs"], key=lambda t: t["y"]):
        bucket = None
        for y_key in lines:
            if abs(y_key - g["y"]) <= 3:
                bucket = y_key
                break
        if bucket is None:
            bucket = g["y"]
            lines[bucket] = []
        lines[bucket].append(g)
    text_lines = []
    for y in sorted(lines):
        chars = "".join(g["ch"] for g in sorted(lines[y], key=lambda t: t["x"]))
        text_lines.append({"y": y, "text": chars})
    result["reconstructed_lines"] = text_lines
    return result


def render_svg(parsed: dict, path: Path, scale_px_per_pt: float = 6.0) -> None:
    """Minimal SVG reconstruction from extracted glyph records (visual check)."""
    bbox = parsed.get("placeable_bbox_pt")
    if not bbox:
        return
    left, top, right, bottom = bbox
    w = max((right - left) * scale_px_per_pt, 10)
    h = max((bottom - top) * scale_px_per_pt, 10)

    def sx(x_pt: float) -> float:
        return (x_pt - left) * scale_px_per_pt

    def sy(y_pt: float) -> float:
        return (y_pt - top) * scale_px_per_pt

    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{w:.0f}" height="{h:.0f}" '
             f'style="background:#fff">']
    for r in parsed["rules"]:
        if r["kind"] == "rect" and "left_pt" in r:
            x = sx(r["left_pt"]); y = sy(r["top_pt"])
            rw = (r["right_pt"] - r["left_pt"]) * scale_px_per_pt
            rh = (r["bottom_pt"] - r["top_pt"]) * scale_px_per_pt
            parts.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{rw:.1f}" height="{rh:.1f}" fill="#000"/>')
        elif r["kind"] == "line" and "x1_pt" in r:
            parts.append(
                f'<line x1="{sx(r["x1_pt"]):.1f}" y1="{sy(r["y1_pt"]):.1f}" '
                f'x2="{sx(r["x2_pt"]):.1f}" y2="{sy(r["y2_pt"]):.1f}" '
                f'stroke="#000" stroke-width="{max(scale_px_per_pt*0.4,1):.1f}"/>')
        elif "points_pt" in r and r["points_pt"]:
            pts = " ".join(f"{sx(px):.1f},{sy(py):.1f}" for px, py in r["points_pt"])
            fill = "#000" if r["kind"] == "polygon" else "none"
            stroke = "none" if r["kind"] == "polygon" else "#000"
            tag = "polygon" if r["kind"] == "polygon" else "polyline"
            parts.append(f'<{tag} points="{pts}" fill="{fill}" stroke="{stroke}"/>')
    for g in parsed["glyphs"]:
        if "x_pt" not in g:
            continue
        size = g.get("size_pt", 10) * scale_px_per_pt
        ch = (g["ch"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        parts.append(
            f'<text x="{sx(g["x_pt"]):.1f}" y="{sy(g["y_pt"]):.1f}" '
            f'font-size="{size:.1f}" font-family="Times New Roman" fill="#000">{ch}</text>')
    parts.append("</svg>")
    path.write_text("\n".join(parts), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("wmf", nargs="?", help="single .wmf file")
    ap.add_argument("--docx", help="batch: parse every .wmf in word/media of a DOCX")
    ap.add_argument("--out", help="output JSON path (batch mode)")
    ap.add_argument("--svg-dir", help="optional dir for SVG reconstructions (batch mode)")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    if args.docx:
        z = zipfile.ZipFile(args.docx)
        wmfs = sorted(n for n in z.namelist()
                      if n.startswith("word/media/") and n.lower().endswith(".wmf"))
        if args.limit:
            wmfs = wmfs[: args.limit]
        out = []
        svg_dir = Path(args.svg_dir) if args.svg_dir else None
        if svg_dir:
            svg_dir.mkdir(parents=True, exist_ok=True)
        for name in wmfs:
            parsed = parse_wmf(z.read(name))
            parsed["source"] = name
            out.append(parsed)
            if svg_dir:
                render_svg(parsed, svg_dir / (Path(name).stem + ".svg"))
        payload = {"docx": args.docx, "count": len(out), "formulas": out}
        text = json.dumps(payload, ensure_ascii=False, indent=1)
        if args.out:
            Path(args.out).parent.mkdir(parents=True, exist_ok=True)
            Path(args.out).write_text(text, encoding="utf-8")
        else:
            print(text)
        return 0

    if not args.wmf:
        ap.error("need a .wmf path or --docx")
    parsed = parse_wmf(Path(args.wmf).read_bytes())
    print(json.dumps(parsed, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
