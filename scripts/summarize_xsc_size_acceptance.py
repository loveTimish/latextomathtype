#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Summarize XSC source-vs-generated WMF size acceptance."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path


def read_rows(size_dir: Path) -> list[dict]:
    rows: list[dict] = []
    for csv_path in sorted(size_dir.rglob("*_vs_*.csv")):
        with csv_path.open(encoding="utf-8-sig", newline="") as fp:
            for row in csv.DictReader(fp):
                row["_csv"] = str(csv_path)
                rows.append(row)
    return rows


def as_float(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    return float(value)


def summarize(rows: list[dict], width_field: str, height_field: str) -> dict:
    width = [as_float(row.get(width_field)) for row in rows]
    height = [as_float(row.get(height_field)) for row in rows]
    width = [value for value in width if value is not None]
    height = [value for value in height if value is not None]
    return {
        "objects": len(rows),
        "widthCompared": len(width),
        "heightCompared": len(height),
        "widthWithin1Pct": sum(abs(value - 1.0) <= 0.01 for value in width),
        "heightWithin1Pct": sum(abs(value - 1.0) <= 0.01 for value in height),
        "maxWidthErrorPct": max((abs(value - 1.0) * 100.0 for value in width), default=None),
        "maxHeightErrorPct": max((abs(value - 1.0) * 100.0 for value in height), default=None),
    }


def worst(rows: list[dict], field: str, limit: int) -> list[dict]:
    ranked = []
    for row in rows:
        value = as_float(row.get(field))
        if value is None:
            continue
        ranked.append((abs(value - 1.0), value, row))
    ranked.sort(reverse=True, key=lambda item: item[0])
    return [
        {
            "errorPct": error * 100.0,
            "ratio": value,
            "csv": row.get("_csv"),
            "index": row.get("index"),
            "generatedIndex": row.get("generated_index"),
            "sourceWmfWPt": row.get("source_wmf_w_pt"),
            "generatedWmfWPt": row.get("generated_wmf_w_pt"),
            "sourceWmfHPt": row.get("source_wmf_h_pt"),
            "generatedWmfHPt": row.get("generated_wmf_h_pt"),
            "latex": row.get("latex"),
        }
        for error, value, row in ranked[:limit]
    ]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("size_dir", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--limit", type=int, default=30)
    args = parser.parse_args()

    rows = read_rows(args.size_dir)
    out = {
        "sizeDir": str(args.size_dir),
        "sourceWmf": summarize(rows, "wmf_w_pt_ratio", "wmf_h_pt_ratio"),
        "targetWmf": summarize(rows, "target_wmf_w_pt_ratio", "target_wmf_h_pt_ratio"),
        "nonWmfGenerated": sum(1 for row in rows if row.get("generated_image_ext") != "wmf"),
        "worstSourceWidth": worst(rows, "wmf_w_pt_ratio", args.limit),
        "worstSourceHeight": worst(rows, "wmf_h_pt_ratio", args.limit),
        "worstTargetWidth": worst(rows, "target_wmf_w_pt_ratio", args.limit),
        "worstTargetHeight": worst(rows, "target_wmf_h_pt_ratio", args.limit),
    }
    text = json.dumps(out, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(text + "\n", encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
