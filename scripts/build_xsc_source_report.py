#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a source formula report for fixed xsc docx2tex assets."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from make_full_batch10_requests import (
    SOURCE_FORMULA_INSERTIONS,
    mml2tex_sequence_for_tex,
    repair_docx2tex_latex,
    source_docx_equations,
    source_docx_metrics,
)


def strip_math_wrapper(value: str) -> str:
    value = (value or "").strip()
    value = re.sub(r"^\$\$\s*", "", value)
    value = re.sub(r"\s*\$\$$", "", value)
    return repair_docx2tex_latex(value.strip())


def repair_single_ole_latex(value: str) -> str:
    value = strip_math_wrapper(value)
    value = value.replace(r"\right] \right\div", r"\right] \div")
    value = value.replace(r"\right\div", r"\div")
    value = value.replace(r"\left] )", r"\right]")
    value = value.replace(r"\left (", r"\left(")
    value = value.replace(r"\right )", r"\right)")
    return repair_docx2tex_latex(value)


def repair_report_output(item: dict, pos: int, mml_values: list[str]) -> tuple[str, str]:
    status = item.get("status") or ""
    reason = item.get("reason") or ""
    output = ""
    if status == "converted":
        output = repair_single_ole_latex(item.get("output", ""))
    if output:
        return output, reason
    if reason == "non-printable-rune":
        source = item.get("source") or ""
        fallback_by_source = {
            "embeddings/oleObject167.bin": "d'=2",
            "embeddings/oleObject168.bin": "a_{n}=35",
            "embeddings/oleObject172.bin": "n=18",
            "embeddings/oleObject177.bin": "d'=2",
        }
        if source in fallback_by_source:
            return fallback_by_source[source], reason
    raise ValueError(
        f"conversion report has no trusted output for source OLE {pos}: "
        f"status={status or '<empty>'} reason={reason or '<empty>'}"
    )


def equations_from_conversion_report(tex_path: Path, source_docx: Path, report_path: Path) -> list[dict]:
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    report_equations = report.get("equations") or []
    insertions = SOURCE_FORMULA_INSERTIONS.get(int(tex_path.stem), {}) if tex_path.stem.isdigit() else {}
    preview_backed = {
        ordinal
        for ordinal, insertion in insertions.items()
        if insertion.get("allowPreviewBacked") and not insertion.get("metricsOnly") and not insertion.get("disabled")
    }
    metrics = [
        item for item in source_docx_metrics(source_docx)
        if item.get("isMathType") or item.get("docObjectIndex") in preview_backed
    ]
    if len(report_equations) != len(metrics):
        return []
    mml_values = mml2tex_sequence_for_tex(tex_path)
    out: list[dict] = []
    for pos, item in enumerate(report_equations, 1):
        output, repair_reason = repair_report_output(item, pos, mml_values)
        row = {
            "output": output,
            "metrics": metrics[pos - 1].get("metrics") or {},
            "status": "converted" if item.get("status") == "converted" else "trusted-repair",
            "source": item.get("source") or f"oleObject{pos}.bin",
            "docObjectIndex": metrics[pos - 1].get("docObjectIndex"),
            "oleTarget": metrics[pos - 1].get("oleTarget") or "",
            "wmfTarget": metrics[pos - 1].get("wmfTarget") or "",
        }
        if repair_reason or item.get("status") != "converted":
            row["sourceRepairReason"] = repair_reason or item.get("status") or "missing-output"
        out.append(row)
    return out


def source_object_is_formula_truth(metric: dict, preview_backed: set[int]) -> bool:
    return bool(metric.get("isMathType") or metric.get("docObjectIndex") in preview_backed)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--index", type=int, required=True)
    parser.add_argument("--tex", type=Path, required=True)
    parser.add_argument("--source-docx", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--diagnostic-out", type=Path)
    parser.add_argument("--conversion-report", type=Path)
    args = parser.parse_args()

    equations = []
    if args.conversion_report and args.conversion_report.exists():
        equations = equations_from_conversion_report(args.tex, args.source_docx, args.conversion_report)
    if not equations:
        equations = source_docx_equations(args.tex, args.source_docx)
    if not equations:
        raise SystemExit(f"no equations found for {args.tex}")
    source_metrics = list(source_docx_metrics(args.source_docx))
    embedded_objects = len(source_metrics)
    insertions = SOURCE_FORMULA_INSERTIONS.get(args.index, {})
    excluded_source_objects = []
    explicit_excluded_ordinals = set()
    for ordinal, insertion in sorted(insertions.items()):
        if not (insertion.get("metricsOnly") or insertion.get("disabled")):
            continue
        explicit_excluded_ordinals.add(ordinal)
        metric = source_metrics[ordinal - 1] if 0 < ordinal <= len(source_metrics) else {}
        excluded_source_objects.append(
            {
                "docObjectIndex": ordinal,
                "oleTarget": metric.get("oleTarget") or "",
                "wmfTarget": metric.get("wmfTarget") or "",
                "expectedOleTarget": insertion.get("ole") or "",
                "expectedWmfTarget": insertion.get("wmf") or "",
                "isMathType": metric.get("isMathType"),
                "mathTypeError": metric.get("mathTypeError") or "",
                "reason": insertion.get("excludedReason") or "",
                "latex": insertion.get("latex") or "",
            }
        )
    preview_backed = {
        ordinal
        for ordinal, insertion in insertions.items()
        if insertion.get("allowPreviewBacked") and not insertion.get("metricsOnly") and not insertion.get("disabled")
    }
    for metric in source_metrics:
        ordinal = metric.get("docObjectIndex")
        if ordinal in explicit_excluded_ordinals:
            continue
        if source_object_is_formula_truth(metric, preview_backed):
            continue
        excluded_source_objects.append(
            {
                "docObjectIndex": ordinal,
                "oleTarget": metric.get("oleTarget") or "",
                "wmfTarget": metric.get("wmfTarget") or "",
                "expectedOleTarget": "",
                "expectedWmfTarget": "",
                "isMathType": metric.get("isMathType"),
                "mathTypeError": metric.get("mathTypeError") or "",
                "reason": "non-mathtype-filtered-object",
                "latex": "",
            }
        )
    source_objects = sum(1 for item in source_metrics if source_object_is_formula_truth(item, preview_backed))
    if len(equations) != source_objects:
        diagnostic = {
            "source": str(args.source_docx),
            "tex": str(args.tex),
            "status": "source_formula_count_mismatch",
            "equations": len(equations),
            "sourceObjects": source_objects,
            "embeddedObjects": embedded_objects,
            "unmappedObjectCount": source_objects - len(equations),
        }
        diagnostic_out = args.diagnostic_out or args.out.with_suffix(".diagnostic.json")
        diagnostic_out.parent.mkdir(parents=True, exist_ok=True)
        diagnostic_out.write_text(json.dumps(diagnostic, ensure_ascii=False, indent=2), encoding="utf-8")
        raise SystemExit(
            f"equation/source object count mismatch for {args.tex}: "
            f"equations={len(equations)} sourceObjects={source_objects} diagnostic={diagnostic_out}"
        )

    report = {
        "source": str(args.source_docx),
        "tex": str(args.tex),
        "embeddedObjects": embedded_objects,
        "sourceObjects": source_objects,
        "excludedSourceObjects": excluded_source_objects,
        "filteredSourceObjects": excluded_source_objects,
        "equations": [
            {
                "index": i,
                "status": item.get("status") or "converted",
                "source": item.get("source") or f"oleObject{i}.bin",
                "docObjectIndex": item.get("docObjectIndex"),
                "oleTarget": item.get("oleTarget") or item.get("source") or "",
                "wmfTarget": item.get("wmfTarget") or "",
                "output": item.get("output", ""),
                "metrics": item.get("metrics") or {},
                **({"sourceRepairReason": item["sourceRepairReason"]} if item.get("sourceRepairReason") else {}),
            }
            for i, item in enumerate(equations, 1)
        ],
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"out": str(args.out), "equations": len(report["equations"]), "sourceObjects": source_objects}, ensure_ascii=False))


if __name__ == "__main__":
    main()
