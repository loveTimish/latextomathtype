# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import io
import json
import zipfile
from pathlib import Path
from typing import Any

import olefile

from extract_formula_boxes import extract_boxes


WATCH_MTCODES = {
    0x00D7: "times",
    0x00F7: "div",
    0x2212: "minus",
    0x22EF: "cdots",
    0x2026: "ldots",
}


def part_name(target: str | None) -> str | None:
    if not target:
        return None
    return "word/" + target.lstrip("/")


def read_equation_native(ole_bytes: bytes) -> bytes | None:
    try:
        with olefile.OleFileIO(io.BytesIO(ole_bytes)) as ole:
            if not ole.exists("Equation Native"):
                return None
            return ole.openstream("Equation Native").read()
    except Exception:
        return None


def extract_mtef(native: bytes | None) -> bytes | None:
    if native is None or len(native) <= 28:
        return None
    return native[28:]


def read_docx_formula_mtef(docx: Path) -> list[dict[str, Any]]:
    boxes = extract_boxes(docx)
    rows: list[dict[str, Any]] = []
    with zipfile.ZipFile(docx, "r") as zf:
        for index, box in enumerate(boxes):
            ole_part = part_name(box.ole_target)
            if ole_part is None:
                continue
            ole_bytes = zf.read(ole_part)
            native = read_equation_native(ole_bytes)
            mtef = extract_mtef(native)
            rows.append(
                {
                    "index": index,
                    "ole_part": ole_part,
                    "context": box.context,
                    "mtef": mtef or b"",
                }
            )
    return rows


def signed_byte(value: int) -> int:
    return value - 128


def scan_char_candidates(mtef: bytes) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    allowed_options = 0x01 | 0x02 | 0x04 | 0x08 | 0x10 | 0x20
    length = len(mtef)
    for offset in range(0, max(0, length - 4)):
        if mtef[offset] != 0x02:
            continue
        pos = offset + 1
        options = mtef[pos]
        if options & ~allowed_options:
            continue
        pos += 1
        if options & 0x08:
            if pos + 2 > length:
                continue
            if mtef[pos] == 0x80 and mtef[pos + 1] == 0x80:
                pos += 6
            else:
                pos += 2
        if pos >= length:
            continue
        raw_typeface = mtef[pos]
        if raw_typeface == 0xFF or raw_typeface < 0x80:
            continue
        typeface = signed_byte(raw_typeface)
        if typeface < 1 or typeface > 32:
            continue
        pos += 1
        mtcode = None
        if not (options & 0x20):
            if pos + 2 > length:
                continue
            mtcode = mtef[pos] | (mtef[pos + 1] << 8)
            pos += 2
        bits8 = None
        if options & 0x04:
            if pos >= length:
                continue
            bits8 = mtef[pos]
            pos += 1
        bits16 = None
        if options & 0x10:
            if pos + 2 > length:
                continue
            bits16 = mtef[pos] | (mtef[pos + 1] << 8)
            pos += 2
        rows.append(
            {
                "offset": offset,
                "options": options,
                "typeface": typeface,
                "raw_typeface": raw_typeface,
                "mtcode": mtcode,
                "bits8": bits8,
                "bits16": bits16,
                "bytes": mtef[offset : min(length, pos + 4)].hex(" "),
            }
        )
    return rows


def summarize_docx(docx: Path) -> dict[str, Any]:
    formulas = []
    total_watch_counts = {name: 0 for name in WATCH_MTCODES.values()}
    for row in read_docx_formula_mtef(docx):
        chars = scan_char_candidates(row["mtef"])
        watched = [c for c in chars if c["mtcode"] in WATCH_MTCODES]
        counts: dict[str, int] = {}
        for char in watched:
            name = WATCH_MTCODES[char["mtcode"]]
            counts[name] = counts.get(name, 0) + 1
            total_watch_counts[name] += 1
        if watched:
            formulas.append(
                {
                    "index": row["index"],
                    "ole_part": row["ole_part"],
                    "context": row["context"],
                    "counts": counts,
                    "watched": watched,
                }
            )
    return {
        "docx": str(docx.resolve()),
        "formula_count_with_watched_chars": len(formulas),
        "total_watch_counts": total_watch_counts,
        "formulas": formulas,
    }


def compare_docx(reference: Path, generated: Path) -> dict[str, Any]:
    ref = summarize_docx(reference)
    gen = summarize_docx(generated)
    paired = []
    ref_by_index = {f["index"]: f for f in ref["formulas"]}
    gen_by_index = {f["index"]: f for f in gen["formulas"]}
    for index in sorted(set(ref_by_index) | set(gen_by_index)):
        r = ref_by_index.get(index)
        g = gen_by_index.get(index)
        paired.append(
            {
                "index": index,
                "reference_counts": None if r is None else r["counts"],
                "generated_counts": None if g is None else g["counts"],
                "reference_times": [] if r is None else [c for c in r["watched"] if c["mtcode"] == 0x00D7],
                "generated_times": [] if g is None else [c for c in g["watched"] if c["mtcode"] == 0x00D7],
            }
        )
    return {"reference": ref, "generated": gen, "paired": paired}


def write_text_report(report: dict[str, Any]) -> str:
    lines: list[str] = []
    if "reference" in report:
        lines.append("MTEF watched character comparison")
        lines.append(f"reference={report['reference']['docx']}")
        lines.append(f"generated={report['generated']['docx']}")
        lines.append(f"reference totals={report['reference']['total_watch_counts']}")
        lines.append(f"generated totals={report['generated']['total_watch_counts']}")
        lines.append("")
        lines.append("First formulas containing times:")
        for pair in report["paired"]:
            if not pair["reference_times"] and not pair["generated_times"]:
                continue
            lines.append(
                f"  #{pair['index']}: ref={pair['reference_counts']} gen={pair['generated_counts']}"
            )
            for label, key in (("ref", "reference_times"), ("gen", "generated_times")):
                for char in pair[key][:8]:
                    lines.append(
                        "    "
                        + f"{label} off={char['offset']} typeface={char['typeface']} "
                        + f"mtcode=0x{char['mtcode']:04X} bits8={fmt_opt(char['bits8'])} "
                        + f"bits16={fmt_opt(char['bits16'])} opts=0x{char['options']:02X} bytes={char['bytes']}"
                    )
            if len(lines) > 80:
                break
    else:
        lines.append("MTEF watched character summary")
        lines.append(f"docx={report['docx']}")
        lines.append(f"totals={report['total_watch_counts']}")
    return "\n".join(lines) + "\n"


def fmt_opt(value: int | None) -> str:
    return "none" if value is None else f"0x{value:02X}"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", type=Path)
    parser.add_argument("--reference", type=Path)
    parser.add_argument("--generated", type=Path)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-text", type=Path)
    args = parser.parse_args()

    if args.docx:
        report = summarize_docx(args.docx)
    elif args.reference and args.generated:
        report = compare_docx(args.reference, args.generated)
    else:
        raise SystemExit("pass --docx or --reference/--generated")

    text = write_text_report(report)
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_text:
        args.out_text.parent.mkdir(parents=True, exist_ok=True)
        args.out_text.write_text(text, encoding="utf-8")
    print(text, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
