#!/usr/bin/env python3
"""
Validate formula-level LaTeX recovery from a docx2tex output.

The check verifies that formulas extracted from the PaperExportRequest text
fields still appear in the recovered TeX after normalizing harmless docx2tex
formatting differences. It is meant to catch symbol loss such as \times or
\cdots becoming malformed characters.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path
from typing import Any


MATH_SPAN = re.compile(r"\$\$(.+?)\$\$|\$(.+?)\$", re.DOTALL)
ARRAY_ENV = re.compile(r"\\begin\{array\}\{[^}]*\}|\\end\{array\}")
TEXT_WRAPPER = re.compile(r"\\(?:textbf|mathbf|mathrm|textit|textrm)\{([^{}]*)\}")
COLOR_WRAPPER = re.compile(r"\\textcolor\{[^{}]*\}\{([^{}]*)\}")
LEFT_RIGHT = re.compile(r"\\(?:left|right)")
SPACE_COMMAND = re.compile(r"\\[,;! ]|\\quad|\\qquad")


def iter_strings(value: Any) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        out: list[str] = []
        for item in value:
            out.extend(iter_strings(item))
        return out
    if isinstance(value, dict):
        out = []
        for item in value.values():
            out.extend(iter_strings(item))
        return out
    return []


def extract_math_spans(text: str) -> list[str]:
    spans: list[str] = []
    for match in MATH_SPAN.finditer(text):
        value = match.group(1) if match.group(1) is not None else match.group(2)
        value = value.strip()
        if value:
            spans.append(value)
    return spans


def unwrap_once(pattern: re.Pattern[str], text: str) -> str:
    while True:
        replaced = pattern.sub(r"\1", text)
        if replaced == text:
            return replaced
        text = replaced


def normalize_latex(text: str) -> str:
    text = text.replace("\ufeff", "")
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("＝", "=").replace("－", "-").replace("，", ",")
    text = text.replace("{\\ldots}{\\ldots}", r"\cdots\cdots")
    text = text.replace(r"\ldots\ldots", r"\cdots\cdots")
    text = text.replace(r"\cdot \cdot \cdot", r"\cdots")
    text = text.replace(r"\cdot\cdot\cdot", r"\cdots")
    text = unwrap_once(COLOR_WRAPPER, text)
    text = unwrap_once(TEXT_WRAPPER, text)
    text = ARRAY_ENV.sub("", text)
    text = LEFT_RIGHT.sub("", text)
    text = SPACE_COMMAND.sub("", text)
    text = text.replace("&", "")
    text = text.replace(r"\\", "")
    text = re.sub(r"\s+", "", text)
    return text


def classify(formula: str) -> str:
    if r"\times" in formula and r"\cdots" in formula:
        return "times_cdots"
    if r"\cdots" in formula:
        return "cdots"
    if r"\times" in formula:
        return "times"
    if r"\frac" in formula:
        return "fraction"
    return "other"


def count_non_overlapping(text: str, needle: str) -> int:
    if not needle:
        return 0
    count = 0
    start = 0
    while True:
        index = text.find(needle, start)
        if index < 0:
            return count
        count += 1
        start = index + len(needle)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--request-json", required=True, type=Path)
    parser.add_argument("--tex", required=True, type=Path)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-text", type=Path)
    parser.add_argument("--require-risk-match", action="store_true")
    parser.add_argument("--min-risk-coverage", type=float, default=0.98)
    parser.add_argument("--require-occurrence-match", action="store_true")
    parser.add_argument("--min-occurrence-coverage", type=float, default=0.98)
    parser.add_argument("--min-occurrence-length", type=int, default=12)
    args = parser.parse_args()

    request = json.loads(args.request_json.read_text(encoding="utf-8"))
    tex = args.tex.read_text(encoding="utf-8", errors="replace")

    bad_chars = ["\ufffd", "\u25a1"]
    bad_present = [f"U+{ord(ch):04X}" for ch in bad_chars if ch in tex]
    normalized_tex = normalize_latex(tex)

    occurrences: list[dict[str, Any]] = []
    formulas_by_normalized: dict[str, dict[str, Any]] = {}
    seen: set[str] = set()
    for source in iter_strings(request):
        for formula in extract_math_spans(source):
            normalized = normalize_latex(formula)
            if not normalized:
                continue
            category = classify(formula)
            matched = normalized in normalized_tex
            occurrence = {
                "formula": formula,
                "normalized": normalized,
                "category": category,
                "matched": matched,
            }
            occurrences.append(occurrence)
            if normalized in seen:
                continue
            seen.add(normalized)
            formulas_by_normalized[normalized] = (
                {
                    "formula": formula,
                    "normalized": normalized,
                    "category": category,
                    "matched": matched,
                }
            )
    formulas = list(formulas_by_normalized.values())

    by_category: dict[str, dict[str, int | float]] = {}
    for category in sorted({item["category"] for item in formulas}):
        rows = [item for item in formulas if item["category"] == category]
        matched = sum(1 for item in rows if item["matched"])
        by_category[category] = {
            "total": len(rows),
            "matched": matched,
            "missing": len(rows) - matched,
            "coverage": matched / len(rows) if rows else 1.0,
        }

    risk_categories = {"times", "cdots", "times_cdots"}
    risk_rows = [item for item in formulas if item["category"] in risk_categories]
    risk_matched = sum(1 for item in risk_rows if item["matched"])
    risk_coverage = risk_matched / len(risk_rows) if risk_rows else 1.0
    missing_risk = [item for item in risk_rows if not item["matched"]]
    missing_all = [item for item in formulas if not item["matched"]]

    occurrence_candidates = [
        item for item in occurrences
        if len(item["normalized"]) >= args.min_occurrence_length
    ]
    occurrence_expected = Counter(item["normalized"] for item in occurrence_candidates)
    occurrence_actual = {
        normalized: count_non_overlapping(normalized_tex, normalized)
        for normalized in occurrence_expected
    }
    occurrence_matched = sum(
        min(expected, occurrence_actual.get(normalized, 0))
        for normalized, expected in occurrence_expected.items()
    )
    occurrence_total = sum(occurrence_expected.values())
    occurrence_coverage = occurrence_matched / occurrence_total if occurrence_total else 1.0
    occurrence_missing: list[dict[str, Any]] = []
    for normalized, expected in occurrence_expected.items():
        actual = occurrence_actual.get(normalized, 0)
        if actual >= expected:
            continue
        source = formulas_by_normalized[normalized]
        occurrence_missing.append(
            {
                "formula": source["formula"],
                "normalized": normalized,
                "category": source["category"],
                "expected": expected,
                "actual": actual,
                "missing": expected - actual,
            }
        )

    symbol_counts = {
        r"\times": tex.count(r"\times"),
        r"\cdots": tex.count(r"\cdots"),
        r"\ldots": tex.count(r"\ldots"),
        r"\frac": tex.count(r"\frac"),
    }

    report = {
        "request_json": str(args.request_json),
        "tex": str(args.tex),
        "formula_count": len(formulas),
        "matched_count": sum(1 for item in formulas if item["matched"]),
        "missing_count": len(missing_all),
        "formula_occurrence_count": len(occurrences),
        "checked_occurrence_count": occurrence_total,
        "occurrence_matched_count": occurrence_matched,
        "occurrence_missing_count": occurrence_total - occurrence_matched,
        "occurrence_coverage": occurrence_coverage,
        "occurrence_min_length": args.min_occurrence_length,
        "risk_formula_count": len(risk_rows),
        "risk_matched_count": risk_matched,
        "risk_missing_count": len(missing_risk),
        "risk_coverage": risk_coverage,
        "bad_characters": bad_present,
        "symbol_counts": symbol_counts,
        "by_category": by_category,
        "missing_occurrence_preview": occurrence_missing[:25],
        "missing_risk_preview": missing_risk[:25],
        "missing_preview": missing_all[:25],
    }

    lines = [
        "docx2tex formula fragment coverage",
        f"  request={args.request_json}",
        f"  tex={args.tex}",
        f"  formulas={report['formula_count']} matched={report['matched_count']} missing={report['missing_count']}",
        "  occurrences={total} checked={checked} matched={matched} missing={missing} coverage={coverage:.4f} min_length={min_length}".format(
            total=report["formula_occurrence_count"],
            checked=report["checked_occurrence_count"],
            matched=report["occurrence_matched_count"],
            missing=report["occurrence_missing_count"],
            coverage=report["occurrence_coverage"],
            min_length=report["occurrence_min_length"],
        ),
        f"  risk_formulas={report['risk_formula_count']} matched={risk_matched} missing={len(missing_risk)} coverage={risk_coverage:.4f}",
        f"  symbols={symbol_counts}",
        f"  bad_characters={bad_present}",
        "  by_category:",
    ]
    for category, stats in by_category.items():
        lines.append(
            "    {category}: total={total} matched={matched} missing={missing} coverage={coverage:.4f}".format(
                category=category,
                **stats,
            )
        )
    if missing_risk:
        lines.append("  missing_risk_preview:")
        for item in missing_risk[:10]:
            lines.append(f"    [{item['category']}] {item['formula']}")
    if occurrence_missing:
        lines.append("  missing_occurrence_preview:")
        for item in occurrence_missing[:10]:
            lines.append(
                "    [{category}] expected={expected} actual={actual} {formula}".format(
                    **item,
                )
            )

    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_text:
        args.out_text.parent.mkdir(parents=True, exist_ok=True)
        args.out_text.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("\n".join(lines))

    if bad_present:
        raise SystemExit("docx2tex output contains malformed characters")
    if args.require_risk_match and risk_coverage < args.min_risk_coverage:
        raise SystemExit(
            f"risk formula coverage below threshold: coverage={risk_coverage:.4f} threshold={args.min_risk_coverage:.4f}"
        )
    if args.require_occurrence_match and occurrence_coverage < args.min_occurrence_coverage:
        raise SystemExit(
            "formula occurrence coverage below threshold: "
            f"coverage={occurrence_coverage:.4f} threshold={args.min_occurrence_coverage:.4f}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
