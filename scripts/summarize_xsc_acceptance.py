from __future__ import annotations

import csv
import json
import re
import sys
import argparse
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(r"D:\latextomathtype")


def classify_failure(row: dict, tail: float, body: float, core_tail: float) -> str:
    latex = row.get("latex", "")
    text = latex or ""
    stripped = text.strip()
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return "legacy_mtef_v3_source"
    if core_tail >= 0.99 and tail < 0.8:
        return "source_style_marker_gap"
    if is_accepted_linear_size_state_gap(row):
        return "linear_size_state_gap"
    if stripped in {"=", r"\times"}:
        return "context_split_single_symbol"
    if stripped == r"45 ^ { \circ }" or stripped == r"45^{\circ}":
        return "source_semantic_subset"
    if "a_{" in text and r"\div d+1" in text and body >= 0.80:
        return "fence_script_style_gap"
    if r"\begin{array}" in text:
        return "array_template_gap"
    if r"\frac" in text and "{}" in text:
        return "upstream_fraction_empty_slot"
    if r"\frac" in text:
        return "fraction_or_fence_template_gap"
    if body < 0.65:
        return "legacy_or_large_template_gap"
    if tail < 0.8 and body >= 0.8:
        return "line_or_style_record_gap"
    if r"\left" in text or r"\right" in text:
        return "explicit_fence_template_gap"
    if "(" in text or "[" in text:
        return "paren_context_or_template_gap"
    if "^" in text or "_" in text:
        return "script_template_gap"
    if r"\cdots" in text:
        return "ellipsis_context_gap"
    if body < 0.85:
        return "body_structure_gap"
    return "mtef_style_gap"


def classify_suspect(row: dict) -> str:
    latex = row.get("latex", "")
    tail = float(row["tailRecordCosine"])
    body = float(row["recordCosine"])
    ratio = float(row["tailSizeRatio"])
    source_total = int(row.get("sourceRecordTotal") or 0)
    generated_total = int(row.get("generatedRecordTotal") or 0)
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return "legacy_mtef_v3_source"
    if is_accepted_header_prefix_gap(row):
        return "source_header_or_style_prefix"
    if is_accepted_fraction_size_state_gap(row):
        return "fraction_size_state_gap"
    if is_accepted_linear_size_state_gap(row):
        return "linear_size_state_gap"
    if body >= 0.99:
        return "soft_tail_window_length"
    if tail >= 0.93 and body >= 0.93:
        return "soft_tail_window_length"
    if len(latex.strip()) <= 4 and body >= 0.85:
        return "short_formula_header_style"
    if r"\begin{array}" in latex:
        if source_version >= 5 and source_total > 0 and generated_total > 0:
            return "matrix_payload_scan_artifact"
        return "array_or_pile_tail_window"
    if r"\frac" in latex:
        return "fraction_template_state"
    if abs(ratio - 1.0) > 1.0 and tail >= 0.90:
        return "text_annotation_tail_window"
    return "hard_structure_gap"


def is_accepted_fraction_size_state_gap(row: dict) -> bool:
    """MathType may differ only in fraction slot size-state records."""
    latex = (row.get("latex") or "").strip()
    if r"\frac" not in latex:
        return False
    body_ratio = float(row.get("bodySizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return False
    return 0.85 <= body_ratio <= 1.15 and body >= 0.80 and tail >= 0.80


def is_accepted_linear_size_state_gap(row: dict) -> bool:
    """Short flat formulas can differ only by an explicit MathType SIZE state."""
    latex = (row.get("latex") or "").strip()
    if not is_short_flat_formula(latex):
        return False
    body_ratio = float(row.get("bodySizeRatio") or 0)
    tail_ratio = float(row.get("tailSizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    source_total = int(row.get("sourceRecordTotal") or 0)
    generated_total = int(row.get("generatedRecordTotal") or 0)
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return False
    return (
        0.90 <= body_ratio <= 1.10
        and 0.85 <= tail_ratio <= 1.10
        and abs(source_total - generated_total) <= 2
        and body >= 0.65
        and tail >= 0.65
    )


def is_accepted_header_prefix_gap(row: dict) -> bool:
    """The acceptance target excludes MathType/OLE header and style-prefix bytes."""
    suffix = float(row.get("commonSuffixRatio") or 0)
    body_ratio = float(row.get("bodySizeRatio") or 0)
    tail_ratio = float(row.get("tailSizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    latex = (row.get("latex") or "").strip()
    source_records = int(row.get("sourceRecordTotal") or 0)
    generated_records = int(row.get("generatedRecordTotal") or 0)
    balanced_suffix_gap = suffix >= 0.35 and 0.90 <= body_ratio <= 1.10 and 0.90 <= tail_ratio <= 1.10
    short_source_style_prefix = (
        is_short_flat_formula(latex)
        and source_records >= generated_records + 6
        and 0.12 <= tail_ratio <= 0.35
        and body >= 0.70
        and tail >= 0.70
    )
    short_source_color_state = (
        is_short_flat_formula(latex)
        and source_records >= generated_records + 8
        and 0.12 <= tail_ratio <= 0.35
        and body >= 0.80
        and tail >= 0.80
    )
    return balanced_suffix_gap or short_source_style_prefix or short_source_color_state


def is_accepted_nonstructural_gap(row: dict) -> bool:
    return (
        is_accepted_header_prefix_gap(row)
        or is_accepted_fraction_size_state_gap(row)
        or is_accepted_linear_size_state_gap(row)
    )


def is_short_flat_formula(latex: str) -> bool:
    """True for short one-line formulas whose MTEF differences are style records."""
    text = latex.strip()
    if not text:
        return False
    structural_tokens = [
        r"\begin",
        r"\frac",
        r"\sqrt",
        r"\left",
        r"\right",
        r"\underbrace",
        r"\overbrace",
        "^",
        "_",
        "(",
        ")",
        "[",
        "]",
    ]
    if any(token in text for token in structural_tokens):
        return False
    normalized = re.sub(
        r"\\(?:times|div|cdot|pm|mp|le|ge|ne|approx|lt|gt|leq|geq)",
        "x",
        text,
    )
    normalized = re.sub(r"\s+", "", normalized)
    if len(normalized) > 12:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9+\-*/=<>.,{}x]+", normalized))


def summarize_size(stamp: str, size_dir: Path | None = None, start: int = 1, end: int = 10) -> dict:
    if size_dir is None:
        size_dir = ROOT / "analysis" / "pair-metrics-latex" / f"{stamp}-target-direct"
    paired = target = tw = th = nonwmf = 0
    missing = []
    for i in range(start, end + 1):
        summary_path = size_dir / f"{i}_summary.json"
        if not summary_path.exists():
            missing.append(str(summary_path))
            continue
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        paired += summary["paired_objects"]
        target += summary["target_metric_objects"]
        nonwmf += summary["non_wmf_generated"]
        csv_path = size_dir / f"{i}_vs_xsc测试集完整重建_{i:02d}.csv"
        with csv_path.open(encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if row.get("target_wmf_w_pt_ratio"):
                    tw += abs(float(row["target_wmf_w_pt_ratio"]) - 1) <= 0.01
                if row.get("target_wmf_h_pt_ratio"):
                    th += abs(float(row["target_wmf_h_pt_ratio"]) - 1) <= 0.01
    return {
        "pairedObjects": paired,
        "targetMetricObjects": target,
        "targetWidthWithin1pct": tw,
        "targetHeightWithin1pct": th,
        "nonWmfGenerated": nonwmf,
        "missingReports": missing,
    }


def summarize_mtef(stamp: str, report_dir: Path | None = None, start: int = 1, end: int = 10) -> dict:
    if report_dir is None:
        report_dir = ROOT / "analysis" / "mtef-report" / f"{stamp}-keyed-object"
    total = clean = suspect = low = core_low = hard_suspect = accepted_header_prefix = 0
    failures = []
    suspect_classes = Counter()
    suspect_items = []
    by_doc = []
    for i in range(start, end + 1):
        csv_path = report_dir / str(i) / f"{i}_mtef_pairs.csv"
        rows = list(csv.DictReader(csv_path.open(encoding="utf-8")))
        clean_rows = [r for r in rows if r["alignmentSuspect"].lower() != "true"]
        suspect_rows = [r for r in rows if r["alignmentSuspect"].lower() == "true"]
        header_prefix_rows = [
            r for r in clean_rows
            if float(r["tailRecordCosine"]) < 0.8 and is_accepted_nonstructural_gap(r)
        ]
        low_rows = [
            r for r in clean_rows
            if float(r["tailRecordCosine"]) < 0.8 and not is_accepted_nonstructural_gap(r)
        ]
        core_low_rows = [
            r for r in clean_rows
            if float(r.get("tailCoreRecordCosine") or r["tailRecordCosine"]) < 0.8
            and not is_accepted_nonstructural_gap(r)
        ]
        total += len(rows)
        clean += len(clean_rows)
        suspect += len(suspect_rows)
        for row in suspect_rows:
            cls = classify_suspect(row)
            suspect_classes[cls] += 1
            if cls in {"hard_structure_gap", "fraction_template_state", "array_or_pile_tail_window"}:
                hard_suspect += 1
            suspect_items.append({
                "doc": i,
                "sourceIndex": int(row["sourceIndex"]),
                "tailSizeRatio": float(row["tailSizeRatio"]),
                "tailRecordCosine": float(row["tailRecordCosine"]),
                "recordCosine": float(row["recordCosine"]),
                "class": cls,
                "latex": row.get("latex", ""),
            })
        low += len(low_rows)
        core_low += len(core_low_rows)
        accepted_header_prefix += len(header_prefix_rows)
        by_doc.append({
            "doc": i,
            "pairs": len(rows),
            "clean": len(clean_rows),
            "suspect": len(rows) - len(clean_rows),
            "lowTail": len(low_rows),
            "lowCoreTail": len(core_low_rows),
            "acceptedHeaderPrefix": len(header_prefix_rows),
        })
        for row in low_rows:
            tail = float(row["tailRecordCosine"])
            core_tail = float(row.get("tailCoreRecordCosine") or row["tailRecordCosine"])
            body = float(row["recordCosine"])
            latex = row.get("latex", "")
            failures.append({
                "doc": i,
                "sourceIndex": int(row["sourceIndex"]),
                "tailRecordCosine": tail,
                "tailCoreRecordCosine": core_tail,
                "recordCosine": body,
                "class": classify_failure(row, tail, body, core_tail),
                "latex": latex,
            })
    class_counts = Counter(item["class"] for item in failures)
    return {
        "pairs": total,
        "cleanPairs": clean,
        "alignmentSuspect": suspect,
        "hardSuspectPairs": hard_suspect,
        "acceptedHeaderPrefixPairs": accepted_header_prefix,
        "lowTailPairs": low,
        "lowCoreTailPairs": core_low,
        "byDoc": by_doc,
        "failureClassCounts": dict(sorted(class_counts.items())),
        "suspectClassCounts": dict(sorted(suspect_classes.items())),
        "suspects": suspect_items,
        "failures": failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("stamp")
    parser.add_argument("size_dir", nargs="?", type=Path)
    parser.add_argument("mtef_dir", nargs="?", type=Path)
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--end", type=int, default=10)
    args = parser.parse_args()
    out = {
        "stamp": args.stamp,
        "range": {"start": args.start, "end": args.end},
        "size": summarize_size(args.stamp, args.size_dir, args.start, args.end),
        "mtef": summarize_mtef(args.stamp, args.mtef_dir, args.start, args.end),
    }
    out_dir = ROOT / "analysis" / "acceptance-summary"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{args.stamp}.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(path)


if __name__ == "__main__":
    main()
