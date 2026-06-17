# -*- coding: utf-8 -*-
"""Compare two WMF glyph metric reports produced by measure_wmf_formula_glyphs.py.

This is a diagnostic diff tool, not an acceptance gate.  It keeps record-level
geometry, ImageMagick preview ink, and run summaries separate so a calibration
change can be attributed before it becomes a renderer tweak.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from aggregate_wmf_structure_metrics import classify_structure  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def fnum(value: Any) -> float | None:
    if value is None:
        return None
    try:
        out = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(out):
        return None
    return out


def round3(value: float | None) -> float | None:
    return None if value is None else round(value, 3)


def delta(generated: Any, source: Any) -> float | None:
    g = fnum(generated)
    s = fnum(source)
    if g is None or s is None:
        return None
    return g - s


def abs_delta(generated: Any, source: Any) -> float | None:
    value = delta(generated, source)
    return None if value is None else abs(value)


def ratio(generated: Any, source: Any) -> float | None:
    g = fnum(generated)
    s = fnum(source)
    if g is None or s is None or s == 0:
        return None
    return g / s


def summarize(values: list[float | None]) -> dict[str, Any]:
    nums = sorted(v for v in values if v is not None and math.isfinite(v))
    if not nums:
        return {"n": 0}
    return {
        "n": len(nums),
        "avg": round3(sum(nums) / len(nums)),
        "median": round3(nums[len(nums) // 2]),
        "min": round3(nums[0]),
        "max": round3(nums[-1]),
    }


def load_report(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def formula_map(report: dict[str, Any]) -> dict[int, dict[str, Any]]:
    out: dict[int, dict[str, Any]] = {}
    for item in report.get("formulas", []):
        index = item.get("objectIndex")
        if isinstance(index, int):
            out[index] = item
    return out


def normalize_formula_key(value: Any) -> str:
    return "".join(ch for ch in str(value or "") if not ch.isspace())


def item_formula_key(item: dict[str, Any]) -> str:
    for key in ("formulaKey", "formula", "rawFormula"):
        value = item.get(key)
        if value:
            return normalize_formula_key(value)
    return ""


def formula_key_map(report: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[str, int]]:
    buckets: dict[str, list[dict[str, Any]]] = {}
    for item in report.get("formulas", []):
        key = item_formula_key(item)
        if key:
            buckets.setdefault(key, []).append(item)
    return (
        {key: items[0] for key, items in buckets.items() if len(items) == 1},
        {key: len(items) for key, items in buckets.items() if len(items) > 1},
    )


def run_key(run: dict[str, Any]) -> tuple[int, str, str]:
    return (
        int(run.get("runIndex", -1)),
        str(run.get("fontFace") or ""),
        str(run.get("text") or ""),
    )


def text_join(item: dict[str, Any]) -> str:
    runs = ((item.get("wmf") or {}).get("runs") or [])
    return "".join(str(run.get("text") or "") for run in runs)


def item_structure(item: dict[str, Any]) -> str:
    latex = item.get("formula") or item.get("rawFormula") or ""
    return classify_structure(str(latex), item)


def summary_value(item: dict[str, Any], key: str) -> Any:
    return ((item.get("wmf") or {}).get("summary") or {}).get(key)


def run_byte_count(runs: list[dict[str, Any]]) -> int:
    return sum(int(run.get("byteCount") or 0) for run in runs)


def run_advance_sum(runs: list[dict[str, Any]]) -> float | None:
    values = [fnum(run.get("advanceWidthPt")) for run in runs]
    nums = [value for value in values if value is not None]
    if not nums:
        return None
    return sum(nums)


def trust_record_width_delta(row: dict[str, Any]) -> str:
    if not row.get("runStructureComparable"):
        return "low_run_structure_mismatch"
    if fnum(row.get("magickInkWidthDeltaPt")) is None:
        return "medium_no_ink"
    record = abs(float(row.get("recordWidthDeltaPt") or 0.0))
    ink = abs(float(row.get("magickInkWidthDeltaPt") or 0.0))
    if record >= 1.0 and ink <= 0.25:
        return "low_record_only"
    return "usable_with_visual_check"


def summarize_by_structure(rows: list[dict[str, Any]], limit: int) -> dict[str, Any]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        grouped.setdefault(str(row.get("structure") or "unknown"), []).append(row)
    out: dict[str, Any] = {}
    for structure, group in sorted(grouped.items()):
        out[structure] = {
            "count": len(group),
            "recordWidthDeltaPt": summarize([row.get("recordWidthDeltaPt") for row in group]),
            "magickInkWidthDeltaPt": summarize([row.get("magickInkWidthDeltaPt") for row in group]),
            "magickInkHeightDeltaPt": summarize([row.get("magickInkHeightDeltaPt") for row in group]),
            "shapeWidthDeltaPt": summarize([row.get("shapeWidthDeltaPt") for row in group]),
            "shapeHeightDeltaPt": summarize([row.get("shapeHeightDeltaPt") for row in group]),
            "lowTrustRecordWidthCount": sum(
                1 for row in group if str(row.get("recordWidthTrust") or "").startswith("low_")
            ),
            "runStructureMismatchCount": sum(1 for row in group if not row.get("runStructureComparable")),
            "worstRecordWidth": worst(group, "recordWidthDeltaPt", limit),
            "worstMagickInkWidth": worst(group, "magickInkWidthDeltaPt", limit),
        }
    return out


def item_row(source: dict[str, Any], generated: dict[str, Any]) -> dict[str, Any]:
    source_runs = ((source.get("wmf") or {}).get("runs") or [])
    generated_runs = ((generated.get("wmf") or {}).get("runs") or [])
    source_run_keys = {run_key(run) for run in source_runs}
    generated_run_keys = {run_key(run) for run in generated_runs}
    source_byte_count = run_byte_count(source_runs)
    generated_byte_count = run_byte_count(generated_runs)
    source_advance_sum = run_advance_sum(source_runs)
    generated_advance_sum = run_advance_sum(generated_runs)
    row = {
        "objectIndex": generated.get("objectIndex") or source.get("objectIndex"),
        "pairKey": generated.get("pairKey") or source.get("pairKey"),
        "structure": item_structure(generated) if item_formula_key(generated) else item_structure(source),
        "sourceFormula": source.get("formula"),
        "generatedFormula": generated.get("formula"),
        "sourceText": text_join(source),
        "generatedText": text_join(generated),
        "textMatch": text_join(source) == text_join(generated),
        "shapeWidthDeltaPt": round3(delta(generated.get("shapeWidthPt"), source.get("shapeWidthPt"))),
        "shapeHeightDeltaPt": round3(delta(generated.get("shapeHeightPt"), source.get("shapeHeightPt"))),
        "wmfWidthDeltaPt": round3(delta(generated.get("wmfPlaceableWidthPt"), source.get("wmfPlaceableWidthPt"))),
        "wmfHeightDeltaPt": round3(delta(generated.get("wmfPlaceableHeightPt"), source.get("wmfPlaceableHeightPt"))),
        "recordWidthDeltaPt": round3(delta(summary_value(generated, "recordWidthPt"), summary_value(source, "recordWidthPt"))),
        "recordWidthAbsDeltaPt": round3(abs_delta(summary_value(generated, "recordWidthPt"), summary_value(source, "recordWidthPt"))),
        "recordWidthRatio": round3(ratio(summary_value(generated, "recordWidthPt"), summary_value(source, "recordWidthPt"))),
        "sourceByteCount": source_byte_count,
        "generatedByteCount": generated_byte_count,
        "byteCountDelta": generated_byte_count - source_byte_count,
        "sourceAdvanceSumPt": round3(source_advance_sum),
        "generatedAdvanceSumPt": round3(generated_advance_sum),
        "advanceSumDeltaPt": round3(delta(generated_advance_sum, source_advance_sum)),
        "magickInkWidthDeltaPt": round3(delta(generated.get("magickInkWidthPt"), source.get("magickInkWidthPt"))),
        "magickInkWidthAbsDeltaPt": round3(abs_delta(generated.get("magickInkWidthPt"), source.get("magickInkWidthPt"))),
        "magickInkWidthRatio": round3(ratio(generated.get("magickInkWidthPt"), source.get("magickInkWidthPt"))),
        "magickInkHeightDeltaPt": round3(delta(generated.get("magickInkHeightPt"), source.get("magickInkHeightPt"))),
        "baselineMinDeltaPt": round3(delta(summary_value(generated, "baselineMinPt"), summary_value(source, "baselineMinPt"))),
        "baselineMaxDeltaPt": round3(delta(summary_value(generated, "baselineMaxPt"), summary_value(source, "baselineMaxPt"))),
        "sourceRunCount": len(source_runs),
        "generatedRunCount": len(generated_runs),
        "runCountDelta": len(generated_runs) - len(source_runs),
        "matchedRunKeys": len(source_run_keys & generated_run_keys),
        "missingRunKeys": len(source_run_keys - generated_run_keys),
        "extraRunKeys": len(generated_run_keys - source_run_keys),
        "runStructureComparable": (
            len(source_runs) == len(generated_runs)
            and text_join(source) == text_join(generated)
            and source_run_keys == generated_run_keys
        ),
        "sourceContext": source.get("context"),
        "generatedContext": generated.get("context"),
    }
    row["recordWidthTrust"] = trust_record_width_delta(row)
    return row


def run_rows(source: dict[str, Any], generated: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    source_runs = ((source.get("wmf") or {}).get("runs") or [])
    generated_runs = ((generated.get("wmf") or {}).get("runs") or [])
    count = max(len(source_runs), len(generated_runs))
    object_index = generated.get("objectIndex") or source.get("objectIndex")
    for index in range(count):
        s = source_runs[index] if index < len(source_runs) else {}
        g = generated_runs[index] if index < len(generated_runs) else {}
        rows.append(
            {
                "objectIndex": object_index,
                "runIndex": index,
                "sourceText": s.get("text"),
                "generatedText": g.get("text"),
                "sourceFontFace": s.get("fontFace"),
                "generatedFontFace": g.get("fontFace"),
                "fontHeightDeltaPt": round3(delta(g.get("fontHeightPt"), s.get("fontHeightPt"))),
                "fontWidthDeltaPt": round3(delta(g.get("fontWidthPt"), s.get("fontWidthPt"))),
                "xDeltaPt": round3(delta(g.get("xPt"), s.get("xPt"))),
                "baselineDeltaPt": round3(delta(g.get("baselinePt"), s.get("baselinePt"))),
                "advanceDeltaPt": round3(delta(g.get("advanceWidthPt"), s.get("advanceWidthPt"))),
                "rightEdgeDeltaPt": round3(delta(g.get("rightEdgePt"), s.get("rightEdgePt"))),
                "sourceAdvancePt": s.get("advanceWidthPt"),
                "generatedAdvancePt": g.get("advanceWidthPt"),
                "sourceByteCount": s.get("byteCount"),
                "generatedByteCount": g.get("byteCount"),
                "sourceRawTextHex": s.get("rawTextHex"),
                "generatedRawTextHex": g.get("rawTextHex"),
            }
        )
    return rows


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    keys: list[str] = []
    for row in rows:
        for key in row:
            if key not in keys:
                keys.append(key)
    if not keys:
        keys = ["objectIndex"]
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=keys)
        writer.writeheader()
        writer.writerows(rows)


def worst(rows: list[dict[str, Any]], key: str, limit: int) -> list[dict[str, Any]]:
    return sorted(
        (row for row in rows if fnum(row.get(key)) is not None),
        key=lambda row: abs(float(row[key])),
        reverse=True,
    )[:limit]


def render_text(summary: dict[str, Any], limit: int) -> str:
    lines = [
        "WMF glyph metrics diff",
        f"source: {summary['source']}",
        f"generated: {summary['generated']}",
        f"pairing: {summary['pairBy']}",
        f"paired: {summary['pairedCount']} missing_source={summary['missingSourceCount']} missing_generated={summary['missingGeneratedCount']}",
        "",
        "By structure",
    ]
    for structure, stats_map in (summary.get("structures") or {}).items():
        ink_w = stats_map.get("magickInkWidthDeltaPt") or {}
        ink_h = stats_map.get("magickInkHeightDeltaPt") or {}
        record_w = stats_map.get("recordWidthDeltaPt") or {}
        lines.append(
            "  {structure}: n={count} inkW_avg={ink_w} inkH_avg={ink_h} recordW_avg={record_w} lowTrust={low} runMismatch={mismatch}".format(
                structure=structure,
                count=stats_map.get("count"),
                ink_w=ink_w.get("avg"),
                ink_h=ink_h.get("avg"),
                record_w=record_w.get("avg"),
                low=stats_map.get("lowTrustRecordWidthCount"),
                mismatch=stats_map.get("runStructureMismatchCount"),
            )
        )
    lines.extend([
        "",
        "Worst record width deltas",
    ])
    for row in summary["worstRecordWidth"][:limit]:
        lines.append(
            "  #{idx}: record={record:+.3f}pt ink={ink}pt shape={shape}pt text={text}".format(
                idx=row["objectIndex"],
                record=row["recordWidthDeltaPt"] or 0.0,
                ink=row.get("magickInkWidthDeltaPt"),
                shape=row.get("shapeWidthDeltaPt"),
                text=(row.get("generatedFormula") or row.get("generatedText") or "")[:90],
            )
        )
    lines.append("")
    lines.append("Record-only / structure mismatch risks")
    risk_rows = [
        row for row in summary["formulaRows"]
        if str(row.get("recordWidthTrust") or "").startswith("low_")
    ]
    for row in risk_rows[:limit]:
        lines.append(
            "  #{idx}: trust={trust} record={record:+.3f}pt runs={sr}->{gr} bytes={sb}->{gb} text={text}".format(
                idx=row["objectIndex"],
                trust=row.get("recordWidthTrust"),
                record=row.get("recordWidthDeltaPt") or 0.0,
                sr=row.get("sourceRunCount"),
                gr=row.get("generatedRunCount"),
                sb=row.get("sourceByteCount"),
                gb=row.get("generatedByteCount"),
                text=(row.get("generatedFormula") or row.get("generatedText") or "")[:90],
            )
        )
    lines.append("")
    lines.append("Worst magick ink width deltas")
    for row in summary["worstMagickInkWidth"][:limit]:
        lines.append(
            "  #{idx}: ink={ink:+.3f}pt record={record}pt text={text}".format(
                idx=row["objectIndex"],
                ink=row["magickInkWidthDeltaPt"] or 0.0,
                record=row.get("recordWidthDeltaPt"),
                text=(row.get("generatedFormula") or row.get("generatedText") or "")[:90],
            )
        )
    lines.append("")
    lines.append("Worst magick ink height deltas")
    for row in summary["worstMagickInkHeight"][:limit]:
        lines.append(
            "  #{idx}: inkH={ink:+.3f}pt text={text}".format(
                idx=row["objectIndex"],
                ink=row["magickInkHeightDeltaPt"] or 0.0,
                text=(row.get("generatedFormula") or row.get("generatedText") or "")[:90],
            )
        )
    return "\n".join(lines)


def paired_items(
    source: dict[str, Any],
    generated: dict[str, Any],
    pair_by: str,
) -> tuple[list[tuple[Any, dict[str, Any], dict[str, Any]]], dict[str, Any]]:
    if pair_by == "formula-key":
        source_items, source_duplicates = formula_key_map(source)
        generated_items, generated_duplicates = formula_key_map(generated)
        common = sorted(set(source_items) & set(generated_items))
        return (
            [(key, source_items[key], generated_items[key]) for key in common],
            {
                "missingSourceIndexes": sorted(set(generated_items) - set(source_items)),
                "missingGeneratedIndexes": sorted(set(source_items) - set(generated_items)),
                "sourceDuplicateFormulaKeys": source_duplicates,
                "generatedDuplicateFormulaKeys": generated_duplicates,
            },
        )
    source_items = formula_map(source)
    generated_items = formula_map(generated)
    common = sorted(set(source_items) & set(generated_items))
    return (
        [(index, source_items[index], generated_items[index]) for index in common],
        {
            "missingSourceIndexes": sorted(set(generated_items) - set(source_items)),
            "missingGeneratedIndexes": sorted(set(source_items) - set(generated_items)),
            "sourceDuplicateFormulaKeys": {},
            "generatedDuplicateFormulaKeys": {},
        },
    )


def compare(source_path: Path, generated_path: Path, limit: int, pair_by: str) -> dict[str, Any]:
    source = load_report(source_path)
    generated = load_report(generated_path)
    pairs, missing = paired_items(source, generated, pair_by)
    rows = []
    for pair_key, source_item, generated_item in pairs:
        source_item = dict(source_item)
        generated_item = dict(generated_item)
        source_item["pairKey"] = pair_key
        generated_item["pairKey"] = pair_key
        rows.append(item_row(source_item, generated_item))
    runs: list[dict[str, Any]] = []
    for pair_key, source_item, generated_item in pairs:
        source_item = dict(source_item)
        generated_item = dict(generated_item)
        source_item["pairKey"] = pair_key
        generated_item["pairKey"] = pair_key
        runs.extend(run_rows(source_item, generated_item))
    summary = {
        "source": str(source_path),
        "generated": str(generated_path),
        "pairBy": pair_by,
        "pairedCount": len(pairs),
        "sourceFormulaCount": len(source.get("formulas") or []),
        "generatedFormulaCount": len(generated.get("formulas") or []),
        "missingSourceIndexes": missing["missingSourceIndexes"],
        "missingGeneratedIndexes": missing["missingGeneratedIndexes"],
        "missingSourceCount": len(missing["missingSourceIndexes"]),
        "missingGeneratedCount": len(missing["missingGeneratedIndexes"]),
        "sourceDuplicateFormulaKeys": missing["sourceDuplicateFormulaKeys"],
        "generatedDuplicateFormulaKeys": missing["generatedDuplicateFormulaKeys"],
        "textMismatchCount": sum(1 for row in rows if not row["textMatch"]),
        "recordWidthDeltaPt": summarize([row.get("recordWidthDeltaPt") for row in rows]),
        "magickInkWidthDeltaPt": summarize([row.get("magickInkWidthDeltaPt") for row in rows]),
        "magickInkHeightDeltaPt": summarize([row.get("magickInkHeightDeltaPt") for row in rows]),
        "shapeWidthDeltaPt": summarize([row.get("shapeWidthDeltaPt") for row in rows]),
        "advanceSumDeltaPt": summarize([row.get("advanceSumDeltaPt") for row in rows]),
        "lowTrustRecordWidthCount": sum(
            1 for row in rows if str(row.get("recordWidthTrust") or "").startswith("low_")
        ),
        "runStructureMismatchCount": sum(1 for row in rows if not row.get("runStructureComparable")),
        "worstRecordWidth": worst(rows, "recordWidthDeltaPt", limit),
        "worstMagickInkWidth": worst(rows, "magickInkWidthDeltaPt", limit),
        "worstMagickInkHeight": worst(rows, "magickInkHeightDeltaPt", limit),
        "structures": summarize_by_structure(rows, limit),
        "formulaRows": rows,
        "runRows": runs,
    }
    return summary


def run_self_test() -> None:
    source = {
        "formulas": [
            {
                "objectIndex": 10,
                "formula": r"\frac{1}{2}",
                "formulaKey": normalize_formula_key(r"\frac{1}{2}"),
                "shapeWidthPt": 12,
                "shapeHeightPt": 30,
                "wmfPlaceableWidthPt": 12,
                "wmfPlaceableHeightPt": 30,
                "magickInkWidthPt": 8,
                "magickInkHeightPt": 20,
                "wmf": {
                    "runs": [
                        {
                            "runIndex": 0,
                            "fontFace": "Times New Roman",
                            "text": "1",
                            "advanceWidthPt": 4,
                            "rightEdgePt": 4,
                            "byteCount": 1,
                        }
                    ],
                    "summary": {"recordWidthPt": 8},
                },
            }
        ]
    }
    generated = {
        "formulas": [
            {
                "objectIndex": 1,
                "formula": r"\frac{1}{2}",
                "formulaKey": normalize_formula_key(r"\frac{1}{2}"),
                "shapeWidthPt": 13,
                "shapeHeightPt": 28,
                "wmfPlaceableWidthPt": 13,
                "wmfPlaceableHeightPt": 28,
                "magickInkWidthPt": 9,
                "magickInkHeightPt": 18,
                "wmf": {
                    "runs": [
                        {
                            "runIndex": 0,
                            "fontFace": "Times New Roman",
                            "text": "1",
                            "advanceWidthPt": 5,
                            "rightEdgePt": 5,
                            "byteCount": 1,
                        }
                    ],
                    "summary": {"recordWidthPt": 9},
                },
            }
        ]
    }
    pairs, missing = paired_items(source, generated, "formula-key")
    assert len(pairs) == 1
    assert missing["missingSourceIndexes"] == []
    row = item_row(pairs[0][1], pairs[0][2])
    assert row["structure"] == "fraction"
    grouped = summarize_by_structure([row], 5)
    assert grouped["fraction"]["count"] == 1
    assert grouped["fraction"]["magickInkWidthDeltaPt"]["avg"] == 1.0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("source_metrics", nargs="?", type=Path)
    parser.add_argument("generated_metrics", nargs="?", type=Path)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-formulas-csv", type=Path)
    parser.add_argument("--out-runs-csv", type=Path)
    parser.add_argument("--out-text", type=Path)
    parser.add_argument("--limit", type=int, default=12)
    parser.add_argument("--pair-by", choices=("object-index", "formula-key"), default="object-index")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    if args.self_test:
        run_self_test()
        print("compare_wmf_formula_glyphs self-test passed")
        return 0
    if args.source_metrics is None or args.generated_metrics is None:
        parser.error("source_metrics and generated_metrics are required unless --self-test is used")

    summary = compare(args.source_metrics, args.generated_metrics, args.limit, args.pair_by)
    if args.out_json:
        args.out_json.parent.mkdir(parents=True, exist_ok=True)
        args.out_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.out_formulas_csv:
        write_csv(args.out_formulas_csv, summary["formulaRows"])
    if args.out_runs_csv:
        write_csv(args.out_runs_csv, summary["runRows"])
    text = render_text(summary, args.limit)
    if args.out_text:
        args.out_text.parent.mkdir(parents=True, exist_ok=True)
        args.out_text.write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
