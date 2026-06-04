# -*- coding: utf-8 -*-
"""
Build a MathType/OLE formula-box dataset from DOCX files.

The dataset is intentionally based on Word's object box metadata instead of
rendered screenshots.  MathType stores editable equation data in OLE/MTEF, but
Word lays out the equation by the VML/OLE display box and run baseline.  Those
are the measurements we need to fit and validate.
"""
from __future__ import annotations

import argparse
import json
import statistics
import zipfile
from dataclasses import asdict
from pathlib import Path
from typing import Any

from extract_formula_boxes import extract_boxes


DEFAULT_SOURCE_ROOT = Path(r"E:\新加卷\新建文件夹\xsc资料")
DEFAULT_OUT_JSON = Path("target/reference-roundtrip/formula-box-dataset.json")
DEFAULT_OUT_TEXT = Path("target/reference-roundtrip/formula-box-dataset.txt")

PRIORITY_KEYWORDS = (
    "分数",
    "裂项",
    "等差",
    "公式",
    "方程",
    "计算",
    "几何",
    "相似",
    "行程",
)


def percentile(values: list[float], pct: float) -> float | None:
    if not values:
        return None
    if len(values) == 1:
        return round(values[0], 2)
    ordered = sorted(values)
    rank = (len(ordered) - 1) * pct
    lo = int(rank)
    hi = min(lo + 1, len(ordered) - 1)
    fraction = rank - lo
    return round(ordered[lo] + (ordered[hi] - ordered[lo]) * fraction, 2)


def stats(values: list[float]) -> dict[str, float | int | None]:
    return {
        "n": len(values),
        "avg": round(statistics.mean(values), 2) if values else None,
        "min": round(min(values), 2) if values else None,
        "p10": percentile(values, 0.10),
        "median": percentile(values, 0.50),
        "p90": percentile(values, 0.90),
        "max": round(max(values), 2) if values else None,
    }


def top_values(values: list[float | int], limit: int = 20) -> list[dict[str, float | int]]:
    counts: dict[float | int, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    return [
        {"value": value, "count": count}
        for value, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:limit]
    ]


def height_buckets(values: list[float]) -> dict[str, int]:
    buckets = {"<12": 0, "12-18": 0, "18-24": 0, "24-32": 0, "32-45": 0, "45+": 0}
    for value in values:
        if value < 12:
            buckets["<12"] += 1
        elif value < 18:
            buckets["12-18"] += 1
        elif value < 24:
            buckets["18-24"] += 1
        elif value < 32:
            buckets["24-32"] += 1
        elif value < 45:
            buckets["32-45"] += 1
        else:
            buckets["45+"] += 1
    return buckets


def discover_docx_files(roots: list[Path], dedupe_by_name: bool) -> list[Path]:
    seen: set[Path] = set()
    files: list[Path] = []
    for root in roots:
        if root.is_file() and root.suffix.lower() == ".docx":
            candidates = [root]
        elif root.is_dir():
            candidates = list(root.rglob("*.docx"))
        else:
            candidates = []
        for item in candidates:
            if item.name.startswith("~$"):
                continue
            resolved = item.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            files.append(resolved)
    ordered = sorted(files, key=priority_sort_key)
    if not dedupe_by_name:
        return ordered

    names: set[str] = set()
    deduped: list[Path] = []
    for item in ordered:
        key = item.name.casefold()
        if key in names:
            continue
        names.add(key)
        deduped.append(item)
    return deduped


def priority_sort_key(path: Path) -> tuple[int, int, int, str]:
    name = path.name
    priority = 0 if any(keyword in name for keyword in PRIORITY_KEYWORDS) else 1
    teacher = 0 if "教师" in name else 1
    root_copy = 0 if "教研输出" not in str(path) else 1
    return (priority, teacher, root_copy, name)


def scan_document(path: Path, doc_index: int, include_boxes: bool) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    boxes = extract_boxes(path)
    widths = [box.style_width_pt for box in boxes]
    heights = [box.style_height_pt for box in boxes]
    positions = [box.position_half_pt for box in boxes if box.position_half_pt is not None]
    summary = {
        "doc_index": doc_index,
        "path": str(path),
        "name": path.name,
        "bytes": path.stat().st_size,
        "formula_count": len(boxes),
        "width": stats(widths),
        "height": stats(heights),
        "baseline_half_pt": stats([float(value) for value in positions]),
        "top_heights": top_values(heights, 10),
        "top_baselines": top_values(positions, 10),
    }
    rows: list[dict[str, Any]] = []
    if include_boxes:
        for box in boxes:
            row = asdict(box)
            row["doc_index"] = doc_index
            row["doc_name"] = path.name
            row["doc_path"] = str(path)
            rows.append(row)
    return summary, rows


def build_dataset(args: argparse.Namespace) -> dict[str, Any]:
    files = discover_docx_files(args.source_root, args.dedupe_by_name)
    documents: list[dict[str, Any]] = []
    all_boxes: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []

    for path in files:
        if args.max_docs > 0 and len(documents) >= args.max_docs:
            break
        try:
            summary, rows = scan_document(path, len(documents), args.include_boxes)
        except (zipfile.BadZipFile, KeyError, OSError, UnicodeDecodeError) as exc:
            errors.append({"path": str(path), "error": str(exc)})
            continue
        if summary["formula_count"] < args.min_formulas:
            continue
        documents.append(summary)
        all_boxes.extend(rows)
        print(f"scanned {summary['formula_count']:4d} formulas: {path}")

    widths = [float(box["style_width_pt"]) for box in all_boxes]
    heights = [float(box["style_height_pt"]) for box in all_boxes]
    positions = [
        int(box["position_half_pt"])
        for box in all_boxes
        if box.get("position_half_pt") is not None
    ]
    return {
        "source_roots": [str(path) for path in args.source_root],
        "dedupe_by_name": args.dedupe_by_name,
        "discovered_docx_count": len(files),
        "document_count": len(documents),
        "formula_count": len(all_boxes),
        "summary": {
            "width": stats(widths),
            "height": stats(heights),
            "baseline_half_pt": stats([float(value) for value in positions]),
            "height_buckets": height_buckets(heights),
            "top_widths": top_values(widths, 20),
            "top_heights": top_values(heights, 20),
            "top_baselines": top_values(positions, 20),
        },
        "documents": documents,
        "boxes": all_boxes,
        "errors": errors,
    }


def render_text_report(dataset: dict[str, Any]) -> str:
    summary = dataset["summary"]
    lines = [
        "Formula box dataset",
        "",
        f"source_roots: {dataset['source_roots']}",
        f"discovered_docx_count: {dataset['discovered_docx_count']}",
        f"document_count: {dataset['document_count']}",
        f"formula_count: {dataset['formula_count']}",
        "",
        f"width_pt: {summary['width']}",
        f"height_pt: {summary['height']}",
        f"baseline_half_pt: {summary['baseline_half_pt']}",
        f"height_buckets: {summary['height_buckets']}",
        f"top_heights: {summary['top_heights']}",
        f"top_baselines: {summary['top_baselines']}",
        "",
        "Top documents by formula count",
    ]
    for row in sorted(dataset["documents"], key=lambda item: item["formula_count"], reverse=True)[:20]:
        lines.append(
            "  {formula_count:4d} formulas  height_median={height_median}pt  width_median={width_median}pt  {name}".format(
                formula_count=row["formula_count"],
                height_median=row["height"]["median"],
                width_median=row["width"]["median"],
                name=row["name"],
            )
        )
    if dataset["errors"]:
        lines.extend(["", "Skipped files"])
        for row in dataset["errors"][:20]:
            lines.append(f"  {row['path']}: {row['error']}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, action="append", default=None)
    parser.add_argument("--out-json", type=Path, default=DEFAULT_OUT_JSON)
    parser.add_argument("--out-text", type=Path, default=DEFAULT_OUT_TEXT)
    parser.add_argument("--max-docs", type=int, default=80, help="0 means unlimited")
    parser.add_argument("--min-formulas", type=int, default=1)
    parser.add_argument("--dedupe-by-name", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--include-boxes", action=argparse.BooleanOptionalAction, default=True)
    args = parser.parse_args()
    args.source_root = args.source_root or [DEFAULT_SOURCE_ROOT]

    dataset = build_dataset(args)
    args.out_json.parent.mkdir(parents=True, exist_ok=True)
    args.out_json.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    args.out_text.write_text(render_text_report(dataset), encoding="utf-8")
    print(render_text_report(dataset))
    print(f"wrote {args.out_json}")
    print(f"wrote {args.out_text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
