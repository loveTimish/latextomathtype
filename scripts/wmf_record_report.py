# -*- coding: utf-8 -*-
"""Report WMF preview record types inside a DOCX."""

from __future__ import annotations

import argparse
import json
import re
import struct
import sys
import zipfile
from collections import Counter
from pathlib import Path


PLACEABLE_WMF_KEY = 0x9AC6CDD7
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
RECORD_NAMES = {
    0x0000: "EOF",
    0x0103: "SetMapMode",
    0x0106: "SetBkMode",
    0x012D: "SelectObject",
    0x012E: "SetTextAlign",
    0x01F0: "DeleteObject",
    0x0209: "SetTextColor",
    0x020B: "SetWindowOrg",
    0x020C: "SetWindowExt",
    0x02FA: "CreatePen",
    0x02FC: "CreateBrush",
    0x02FB: "CreateFont",
    0x0325: "Polyline",
    0x041B: "Rectangle",
    0x0521: "TextOut",
    0x0538: "PolyPolygon",
    0x0A32: "ExtTextOut",
    0x0922: "BitBlt",
    0x0940: "DibBitBlt",
    0x0B23: "StretchBlt",
    0x0B41: "DibStretchBlt",
    0x0D33: "SetDIBitsToDevice",
    0x0F43: "StretchDIB",
}

BITMAP_RECORDS = {0x0922, 0x0940, 0x0B23, 0x0B41, 0x0D33, 0x0F43}
TEXT_RECORDS = {0x0521, 0x0A32}
PURE_VECTOR_RECORDS = {
    0x0000, 0x0103, 0x0106, 0x012D, 0x01F0, 0x020B, 0x020C,
    0x02FA, 0x02FC, 0x0324, 0x0538,
}
LEAK_PATTERNS = [
    r"\\(?:alpha|beta|gamma|delta|theta|lambda|mu|pi|sigma|omega|frac|sqrt|begin|end|left|right|times|div|cdot|textcolor|pwmetrics|pwstyle|overset|underset|array|cases|matrix|mathrm|mathbf|mathit|overline|underline|overrightarrow|overarc|angle|triangle|because|therefore)\b",
    r"\b(?:pwmetrics|pwstyle|begin\{(?:array|equation|cases|matrix)|end\{(?:array|equation|cases|matrix)|textcolor)\b",
    r"\$\$?\s*[^$\n]{0,200}(?:\\[A-Za-z]+|\\(?:begin|end)\{)",
]

SUSPICIOUS_SCRIPT_RE = re.compile(r"[A-Za-z0-9)]\s*[_^]\s*(?:\{|[A-Za-z0-9+\-])")
REVIEW_REQUIRED_PATTERNS = [
    re.compile(pattern)
    for pattern in LEAK_PATTERNS
]


def wmf_records(data: bytes) -> list[int]:
    offset = 22 if len(data) >= 22 and struct.unpack("<I", data[:4])[0] == PLACEABLE_WMF_KEY else 0
    if offset + 18 > len(data):
        return []
    offset += 18
    records: list[int] = []
    while offset + 6 <= len(data):
        size_words = struct.unpack("<I", data[offset:offset + 4])[0]
        function = struct.unpack("<H", data[offset + 4:offset + 6])[0]
        records.append(function)
        if function == 0 or size_words <= 0:
            break
        offset += size_words * 2
    return records


def decode_wmf_text_payload(payload: bytes) -> str:
    candidates = []
    for encoding in ("utf-16le", "gb18030", "latin1"):
        try:
            text = payload.decode(encoding, "ignore")
        except Exception:
            continue
        text = "".join(ch if ch.isprintable() or ch in "\t\r\n" else " " for ch in text)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            candidates.append(text)
    return " ".join(candidates)


def looks_like_latex_leak(text: str) -> str | None:
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return None
    for pattern in LEAK_PATTERNS:
        if re.search(pattern, compact):
            return pattern
    return None


def looks_like_suspicious_script(text: str) -> bool:
    compact = re.sub(r"\s+", " ", text).strip()
    return bool(SUSPICIOUS_SCRIPT_RE.search(compact))


def suspicious_text_requires_review(text: str) -> bool:
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return False
    if "$" in compact:
        return True
    if any(pattern.search(compact) for pattern in REVIEW_REQUIRED_PATTERNS):
        return True
    return any(token in compact for token in ("pwmetrics", "pwstyle", "textcolor"))


def wmf_text_leaks(data: bytes, name: str) -> list[dict]:
    offset = 22 if len(data) >= 22 and struct.unpack("<I", data[:4])[0] == PLACEABLE_WMF_KEY else 0
    if offset + 18 > len(data):
        return []
    offset += 18
    leaks = []
    suspicious_scripts = []
    while offset + 6 <= len(data):
        size_words = struct.unpack("<I", data[offset:offset + 4])[0]
        function = struct.unpack("<H", data[offset + 4:offset + 6])[0]
        if function in (0x0521, 0x0A32) and size_words > 3:
            payload = data[offset + 6:offset + size_words * 2]
            text = decode_wmf_text_payload(payload)
            pattern = looks_like_latex_leak(text)
            if pattern:
                leaks.append({
                    "file": name,
                    "record": RECORD_NAMES.get(function, hex(function)),
                    "pattern": pattern,
                    "text": text[:300],
                })
            elif looks_like_suspicious_script(text):
                suspicious_scripts.append({
                    "file": name,
                    "record": RECORD_NAMES.get(function, hex(function)),
                    "text": text[:300],
                    "requiresReview": suspicious_text_requires_review(text),
                })
        if function == 0 or size_words <= 0:
            break
        offset += size_words * 2
    return leaks, suspicious_scripts


def report(docx: Path) -> dict:
    files = []
    with zipfile.ZipFile(docx) as z:
        for name in z.namelist():
            if not name.startswith("word/media/") or not name.lower().endswith(".wmf"):
                continue
            data = z.read(name)
            records = wmf_records(data)
            counts = Counter(records)
            leaks, suspicious_scripts = wmf_text_leaks(data, name)
            review_required = [item for item in suspicious_scripts if item.get("requiresReview")]
            files.append(
                {
                    "name": name,
                    "bytes": len(data),
                    "records": len(records),
                    "hasVectorText": counts[0x0A32] > 0 or counts[0x0521] > 0,
                    "hasVectorContent": any(counts[fn] > 0 for fn in (0x0A32, 0x0521, 0x0325, 0x041B, 0x0538)),
                    "hasStretchDib": counts[0x0F43] > 0,
                    "hasBitmapRecord": any(counts[fn] > 0 for fn in BITMAP_RECORDS),
                    "bitmapRecordCount": sum(counts[fn] for fn in BITMAP_RECORDS),
                    "wmfLatexLeakCount": len(leaks),
                    "wmfLatexLeakSamples": leaks[:5],
                    "wmfSuspiciousScriptTextCount": len(suspicious_scripts),
                    "wmfSuspiciousScriptTextSamples": suspicious_scripts[:5],
                    "wmfSuspiciousReviewRequiredCount": len(review_required),
                    "wmfSuspiciousReviewRequiredSamples": review_required[:5],
                    "recordCounts": {RECORD_NAMES.get(k, hex(k)): v for k, v in sorted(counts.items())},
                }
            )
    return {
        "docx": str(docx),
        "wmf": len(files),
        "vectorText": sum(1 for item in files if item["hasVectorText"]),
        "vectorContent": sum(1 for item in files if item["hasVectorContent"]),
        "stretchDib": sum(1 for item in files if item["hasStretchDib"]),
        "bitmapRecords": sum(item["bitmapRecordCount"] for item in files),
        "bitmapWmf": sum(1 for item in files if item["hasBitmapRecord"]),
        "wmfLatexLeakCount": sum(item["wmfLatexLeakCount"] for item in files),
        "wmfLatexLeakSamples": [sample for item in files for sample in item["wmfLatexLeakSamples"]][:20],
        "wmfSuspiciousScriptTextCount": sum(item["wmfSuspiciousScriptTextCount"] for item in files),
        "wmfSuspiciousScriptTextSamples": [sample for item in files for sample in item["wmfSuspiciousScriptTextSamples"]][:20],
        "wmfSuspiciousReviewRequiredCount": sum(item["wmfSuspiciousReviewRequiredCount"] for item in files),
        "wmfSuspiciousReviewRequiredSamples": [
            sample for item in files for sample in item["wmfSuspiciousReviewRequiredSamples"]
        ][:20],
        "files": files,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    result = report(args.docx)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
