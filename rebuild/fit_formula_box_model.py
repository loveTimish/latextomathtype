# -*- coding: utf-8 -*-
"""
Fit and report MathType/OLE display-box corrections.

This script is not a substitute for exact reference calibration.  When a
reference DOCX is available and object order matches, exact calibration is the
stronger operation.  The fitted model is used to understand systematic drift
from the generator and to provide a reusable fallback profile for future
documents that do not have a one-to-one reference.
"""
from __future__ import annotations

import argparse
import json
import math
import statistics
from dataclasses import asdict
from pathlib import Path
from typing import Any

from compare_reference_format import extract_request_formulas
from extract_formula_boxes import extract_boxes


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_REFERENCE = ROOT / "rebuild-assets/external/fraction-split-reference.docx"
DEFAULT_GENERATED = ROOT / "target/reference-roundtrip/fraction-split-reference-regenerated.docx"
DEFAULT_UNCALIBRATED_BOXES = ROOT / "target/reference-roundtrip/formula-boxes-generated-uncalibrated.json"
DEFAULT_REQUEST = ROOT / "target/reference-roundtrip/fraction-split-reference.request.json"
DEFAULT_DATASET = ROOT / "target/reference-roundtrip/formula-box-dataset.json"
DEFAULT_OUT_JSON = ROOT / "target/reference-roundtrip/formula-box-fit-model.json"
DEFAULT_OUT_TEXT = ROOT / "target/reference-roundtrip/formula-box-fit-model.txt"


def percentile(values: list[float], pct: float) -> float | None:
    if not values:
        return None
    if len(values) == 1:
        return round(values[0], 3)
    ordered = sorted(values)
    rank = (len(ordered) - 1) * pct
    lo = int(rank)
    hi = min(lo + 1, len(ordered) - 1)
    fraction = rank - lo
    return round(ordered[lo] + (ordered[hi] - ordered[lo]) * fraction, 3)


def stats(values: list[float]) -> dict[str, float | int | None]:
    return {
        "n": len(values),
        "avg": round(statistics.mean(values), 3) if values else None,
        "min": round(min(values), 3) if values else None,
        "p10": percentile(values, 0.10),
        "median": percentile(values, 0.50),
        "p90": percentile(values, 0.90),
        "max": round(max(values), 3) if values else None,
    }


def rmse(values: list[float]) -> float | None:
    if not values:
        return None
    return round(math.sqrt(sum(value * value for value in values) / len(values)), 3)


def read_box_payload(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, dict) and isinstance(payload.get("boxes"), list):
        return payload["boxes"]
    if isinstance(payload, list):
        return payload
    raise ValueError(f"Unsupported box payload: {path}")


def load_boxes(docx: Path | None, box_json: Path | None) -> list[dict[str, Any]]:
    if box_json and box_json.exists():
        return read_box_payload(box_json)
    if docx is None:
        return []
    return [asdict(item) for item in extract_boxes(docx)]


def top_values(values: list[float | int], limit: int = 20) -> list[dict[str, float | int]]:
    counts: dict[float | int, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    return [
        {"value": value, "count": count}
        for value, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:limit]
    ]


def quantize_pt(value: float, quantum: float = 0.75) -> float:
    return round(round(value / quantum) * quantum, 2)


def fit_linear(rows: list[dict[str, Any]], raw_key: str, ref_key: str) -> dict[str, float | int | None]:
    if not rows:
        return {"n": 0, "slope": None, "intercept": None}
    xs = [float(row[raw_key]) for row in rows]
    ys = [float(row[ref_key]) for row in rows]
    mean_x = statistics.mean(xs)
    mean_y = statistics.mean(ys)
    denom = sum((x - mean_x) ** 2 for x in xs)
    slope = 0.0 if denom == 0 else sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys)) / denom
    intercept = mean_y - slope * mean_x
    return {"n": len(rows), "slope": round(slope, 6), "intercept": round(intercept, 6)}


def predict(row: dict[str, Any], model: dict[str, float | int | None], raw_key: str) -> float:
    slope = model.get("slope")
    intercept = model.get("intercept")
    if slope is None or intercept is None:
        return float(row[raw_key])
    return float(slope) * float(row[raw_key]) + float(intercept)


def classify_by_request(request_json: Path, count: int) -> list[str]:
    formulas = extract_request_formulas(request_json) if request_json.exists() else []
    classes = [item.latex_class for item in formulas[:count]]
    if len(classes) < count:
        classes.extend(["unknown"] * (count - len(classes)))
    return classes


def pair_rows(reference: list[dict[str, Any]], generated: list[dict[str, Any]], classes: list[str]) -> list[dict[str, Any]]:
    count = min(len(reference), len(generated))
    rows: list[dict[str, Any]] = []
    for index in range(count):
        ref = reference[index]
        gen = generated[index]
        rows.append(
            {
                "index": index,
                "latex_class": classes[index] if index < len(classes) else "unknown",
                "generated_width_pt": float(gen["style_width_pt"]),
                "generated_height_pt": float(gen["style_height_pt"]),
                "generated_baseline_half_pt": gen.get("position_half_pt"),
                "reference_width_pt": float(ref["style_width_pt"]),
                "reference_height_pt": float(ref["style_height_pt"]),
                "reference_baseline_half_pt": ref.get("position_half_pt"),
                "context": ref.get("context") or gen.get("context") or "",
            }
        )
    return rows


def evaluate_model(rows: list[dict[str, Any]], width_model: dict[str, Any], height_model: dict[str, Any]) -> dict[str, Any]:
    width_errors = [row["generated_width_pt"] - row["reference_width_pt"] for row in rows]
    height_errors = [row["generated_height_pt"] - row["reference_height_pt"] for row in rows]
    fitted_width_errors = [
        predict(row, width_model, "generated_width_pt") - row["reference_width_pt"]
        for row in rows
    ]
    fitted_height_errors = [
        predict(row, height_model, "generated_height_pt") - row["reference_height_pt"]
        for row in rows
    ]
    snapped_height_errors = [
        quantize_pt(predict(row, height_model, "generated_height_pt")) - row["reference_height_pt"]
        for row in rows
    ]
    return {
        "raw_width_abs_delta_pt": stats([abs(value) for value in width_errors]),
        "raw_height_abs_delta_pt": stats([abs(value) for value in height_errors]),
        "raw_width_rmse_pt": rmse(width_errors),
        "raw_height_rmse_pt": rmse(height_errors),
        "linear_width_abs_delta_pt": stats([abs(value) for value in fitted_width_errors]),
        "linear_height_abs_delta_pt": stats([abs(value) for value in fitted_height_errors]),
        "linear_width_rmse_pt": rmse(fitted_width_errors),
        "linear_height_rmse_pt": rmse(fitted_height_errors),
        "snapped_height_abs_delta_pt": stats([abs(value) for value in snapped_height_errors]),
        "snapped_height_rmse_pt": rmse(snapped_height_errors),
    }


def fit_by_class(rows: list[dict[str, Any]], min_count: int) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for latex_class in sorted({row["latex_class"] for row in rows}):
        class_rows = [row for row in rows if row["latex_class"] == latex_class]
        if len(class_rows) < min_count:
            continue
        width_model = fit_linear(class_rows, "generated_width_pt", "reference_width_pt")
        height_model = fit_linear(class_rows, "generated_height_pt", "reference_height_pt")
        result[latex_class] = {
            "count": len(class_rows),
            "width_model": width_model,
            "height_model": height_model,
            "evaluation": evaluate_model(class_rows, width_model, height_model),
        }
    return result


def dataset_profile(dataset_json: Path) -> dict[str, Any]:
    if not dataset_json.exists():
        return {"available": False}
    dataset = json.loads(dataset_json.read_text(encoding="utf-8"))
    boxes = dataset.get("boxes") or []
    widths = [float(box["style_width_pt"]) for box in boxes]
    heights = [float(box["style_height_pt"]) for box in boxes]
    baselines = [int(box["position_half_pt"]) for box in boxes if box.get("position_half_pt") is not None]
    return {
        "available": True,
        "source_roots": dataset.get("source_roots"),
        "document_count": dataset.get("document_count"),
        "formula_count": len(boxes),
        "width": stats(widths),
        "height": stats(heights),
        "baseline_half_pt": stats([float(value) for value in baselines]),
        "top_widths": top_values(widths, 20),
        "top_heights": top_values(heights, 20),
        "top_baselines": top_values(baselines, 20),
        "assumed_width_quantum_pt": 0.75,
        "assumed_height_quantum_pt": 0.75,
    }


def build_model(args: argparse.Namespace) -> dict[str, Any]:
    reference_boxes = load_boxes(args.reference, args.reference_boxes)
    generated_boxes = load_boxes(args.generated, args.generated_boxes)
    classes = classify_by_request(args.request_json, min(len(reference_boxes), len(generated_boxes)))
    rows = pair_rows(reference_boxes, generated_boxes, classes)
    width_model = fit_linear(rows, "generated_width_pt", "reference_width_pt")
    height_model = fit_linear(rows, "generated_height_pt", "reference_height_pt")
    by_class = fit_by_class(rows, args.min_class_count)
    return {
        "reference": str(args.reference) if args.reference else None,
        "generated_uncalibrated": str(args.generated_boxes if args.generated_boxes.exists() else args.generated),
        "request_json": str(args.request_json),
        "paired_count": len(rows),
        "unpaired_reference": max(len(reference_boxes) - len(rows), 0),
        "unpaired_generated": max(len(generated_boxes) - len(rows), 0),
        "global_width_model": width_model,
        "global_height_model": height_model,
        "global_evaluation": evaluate_model(rows, width_model, height_model),
        "by_class": by_class,
        "dataset_profile": dataset_profile(args.dataset_json),
        "exact_reference_calibration": {
            "available": len(reference_boxes) == len(generated_boxes) and len(rows) > 0,
            "expected_width_abs_delta_pt_after_calibration": 0.0 if len(reference_boxes) == len(generated_boxes) else None,
            "expected_height_abs_delta_pt_after_calibration": 0.0 if len(reference_boxes) == len(generated_boxes) else None,
            "expected_baseline_abs_delta_half_pt_after_calibration": 0 if len(reference_boxes) == len(generated_boxes) else None,
        },
        "rows": rows[:200],
    }


def render_text_report(model: dict[str, Any]) -> str:
    dataset = model["dataset_profile"]
    lines = [
        "Formula box fit model",
        "",
        f"paired_count: {model['paired_count']}",
        f"unpaired_reference: {model['unpaired_reference']}",
        f"unpaired_generated: {model['unpaired_generated']}",
        "",
        "Global fit",
        f"  width_model: {model['global_width_model']}",
        f"  height_model: {model['global_height_model']}",
        f"  evaluation: {model['global_evaluation']}",
        "",
        "Class fit",
    ]
    for latex_class, row in model["by_class"].items():
        evaluation = row["evaluation"]
        lines.append(
            f"  {latex_class}: count={row['count']} "
            f"width_abs={evaluation['linear_width_abs_delta_pt']} "
            f"height_abs={evaluation['linear_height_abs_delta_pt']}"
        )
    lines.extend(["", "External dataset profile"])
    if dataset.get("available"):
        lines.extend(
            [
                f"  source_roots: {dataset['source_roots']}",
                f"  document_count: {dataset['document_count']}",
                f"  formula_count: {dataset['formula_count']}",
                f"  width_pt: {dataset['width']}",
                f"  height_pt: {dataset['height']}",
                f"  baseline_half_pt: {dataset['baseline_half_pt']}",
                f"  top_heights: {dataset['top_heights']}",
                f"  top_baselines: {dataset['top_baselines']}",
            ]
        )
    else:
        lines.append("  unavailable")
    lines.extend(
        [
            "",
            "Exact reference calibration",
            f"  {model['exact_reference_calibration']}",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--generated", type=Path, default=DEFAULT_GENERATED)
    parser.add_argument("--reference-boxes", type=Path)
    parser.add_argument("--generated-boxes", type=Path, default=DEFAULT_UNCALIBRATED_BOXES)
    parser.add_argument("--request-json", type=Path, default=DEFAULT_REQUEST)
    parser.add_argument("--dataset-json", type=Path, default=DEFAULT_DATASET)
    parser.add_argument("--out-json", type=Path, default=DEFAULT_OUT_JSON)
    parser.add_argument("--out-text", type=Path, default=DEFAULT_OUT_TEXT)
    parser.add_argument("--min-class-count", type=int, default=5)
    args = parser.parse_args()

    model = build_model(args)
    args.out_json.parent.mkdir(parents=True, exist_ok=True)
    args.out_json.write_text(json.dumps(model, ensure_ascii=False, indent=2), encoding="utf-8")
    args.out_text.write_text(render_text_report(model), encoding="utf-8")
    print(render_text_report(model))
    print(f"wrote {args.out_json}")
    print(f"wrote {args.out_text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
