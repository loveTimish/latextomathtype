# -*- coding: utf-8 -*-
"""Validate MathType OLE objects and pure-vector WMF previews in DOCX corpora."""

from __future__ import annotations

import argparse
import json
import struct
import sys
import zipfile
from collections import Counter
from pathlib import Path

from scan_docx_latex_leaks import validate_mathtype_ole
from wmf_record_report import (
    BITMAP_RECORDS,
    PLACEABLE_WMF_KEY,
    PURE_VECTOR_RECORDS,
    RECORD_NAMES,
    TEXT_RECORDS,
)


def validate_wmf(data: bytes) -> tuple[list[int], list[str]]:
    errors: list[str] = []
    if len(data) < 46:
        return [], ["shorter than placeable and metafile headers"]
    if struct.unpack_from("<I", data, 0)[0] != PLACEABLE_WMF_KEY:
        return [], ["missing placeable key"]
    left, top, right, bottom, units = struct.unpack_from("<hhhhH", data, 6)
    if right <= left or bottom <= top or units == 0:
        errors.append("empty physical bounds")
    checksum = 0
    for offset in range(0, 20, 2):
        checksum ^= struct.unpack_from("<H", data, offset)[0]
    if checksum != struct.unpack_from("<H", data, 20)[0]:
        errors.append("invalid placeable checksum")
    if struct.unpack_from("<H", data, 24)[0] != 9:
        errors.append("invalid metafile header size")
    if struct.unpack_from("<I", data, 28)[0] * 2 != len(data) - 22:
        errors.append("declared metafile size mismatch")

    offset = 40
    records: list[int] = []
    eof = False
    while offset + 6 <= len(data):
        size_words, function = struct.unpack_from("<IH", data, offset)
        size_bytes = size_words * 2
        if size_words < 3 or offset + size_bytes > len(data):
            errors.append(f"invalid record at byte {offset}")
            break
        records.append(function)
        offset += size_bytes
        if function == 0:
            eof = True
            break
    if not eof:
        errors.append("missing EOF record")
    elif offset != len(data):
        errors.append("trailing bytes after EOF")
    return records, errors


def validate_docx(path: Path) -> dict:
    failures: list[str] = []
    record_counts: Counter[int] = Counter()
    valid_ole = 0
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        ole_names = [name for name in names if name.startswith("word/embeddings/")]
        wmf_names = [name for name in names if name.startswith("word/media/") and name.lower().endswith(".wmf")]
        for name in ole_names:
            try:
                valid, error = validate_mathtype_ole(archive.read(name))
            except Exception as exc:
                valid, error = False, str(exc)
            if valid:
                valid_ole += 1
            else:
                failures.append(f"{name}: invalid MathType OLE: {error}")
        for name in wmf_names:
            records, errors = validate_wmf(archive.read(name))
            record_counts.update(records)
            failures.extend(f"{name}: {error}" for error in errors)
            illegal = sorted(set(records) - PURE_VECTOR_RECORDS)
            if illegal:
                labels = [RECORD_NAMES.get(value, hex(value)) for value in illegal]
                failures.append(f"{name}: forbidden WMF records: {labels}")
        if len(ole_names) != len(wmf_names):
            failures.append(f"OLE/WMF count mismatch: {len(ole_names)} != {len(wmf_names)}")
    return {
        "path": str(path),
        "oleCount": len(ole_names),
        "validOleCount": valid_ole,
        "wmfCount": len(wmf_names),
        "bitmapRecordCount": sum(record_counts[value] for value in BITMAP_RECORDS),
        "textRecordCount": sum(record_counts[value] for value in TEXT_RECORDS),
        "recordCounts": {RECORD_NAMES.get(key, hex(key)): value for key, value in sorted(record_counts.items())},
        "failureCount": len(failures),
        "failures": failures[:50],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path, help="DOCX file or directory")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--minimum-ole", type=int, default=0)
    args = parser.parse_args()

    paths = [args.input] if args.input.is_file() else sorted(args.input.glob("*.docx"))
    documents = [validate_docx(path) for path in paths]
    summary = {
        "schemaVersion": 1,
        "documentCount": len(documents),
        "oleCount": sum(item["oleCount"] for item in documents),
        "validOleCount": sum(item["validOleCount"] for item in documents),
        "wmfCount": sum(item["wmfCount"] for item in documents),
        "bitmapRecordCount": sum(item["bitmapRecordCount"] for item in documents),
        "textRecordCount": sum(item["textRecordCount"] for item in documents),
        "failedDocumentCount": sum(item["failureCount"] > 0 for item in documents),
        "failures": [
            {"path": item["path"], "failures": item["failures"]}
            for item in documents if item["failureCount"] > 0
        ],
    }
    summary["passed"] = (
        summary["documentCount"] > 0
        and summary["oleCount"] >= args.minimum_ole
        and summary["validOleCount"] == summary["oleCount"] == summary["wmfCount"]
        and summary["bitmapRecordCount"] == 0
        and summary["textRecordCount"] == 0
        and summary["failedDocumentCount"] == 0
    )
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["passed"] else 1


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
