# -*- coding: utf-8 -*-
"""Report WMF preview record types inside a DOCX."""

from __future__ import annotations

import argparse
import json
import struct
import zipfile
from collections import Counter
from pathlib import Path


PLACEABLE_WMF_KEY = 0x9AC6CDD7
RECORD_NAMES = {
    0x0000: "EOF",
    0x0103: "SetMapMode",
    0x0106: "SetBkMode",
    0x012D: "SelectObject",
    0x012E: "SetTextAlign",
    0x0209: "SetTextColor",
    0x020B: "SetWindowOrg",
    0x020C: "SetWindowExt",
    0x02FA: "CreatePen",
    0x02FB: "CreateFont",
    0x0325: "Polyline",
    0x041B: "Rectangle",
    0x0521: "TextOut",
    0x0538: "PolyPolygon",
    0x0A32: "ExtTextOut",
    0x0F43: "StretchDIB",
}


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


def report(docx: Path) -> dict:
    files = []
    with zipfile.ZipFile(docx) as z:
        for name in z.namelist():
            if not name.startswith("word/media/") or not name.lower().endswith(".wmf"):
                continue
            records = wmf_records(z.read(name))
            counts = Counter(records)
            files.append(
                {
                    "name": name,
                    "bytes": len(z.read(name)),
                    "records": len(records),
                    "hasVectorText": counts[0x0A32] > 0 or counts[0x0521] > 0,
                    "hasStretchDib": counts[0x0F43] > 0,
                    "recordCounts": {RECORD_NAMES.get(k, hex(k)): v for k, v in sorted(counts.items())},
                }
            )
    return {
        "docx": str(docx),
        "wmf": len(files),
        "vectorText": sum(1 for item in files if item["hasVectorText"]),
        "stretchDib": sum(1 for item in files if item["hasStretchDib"]),
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
