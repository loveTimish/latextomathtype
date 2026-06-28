#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aggregate reference MathType WMF metrics by formula structure.

This tool reads source DOCX WMF previews, enriches them with source-report
LaTeX when available, then summarizes per-structure font, advance, baseline,
and box geometry.  It is meant to build the calibration table before changing
the renderer.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import re
import statistics as st
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT / "scripts") not in sys.path:
    sys.path.insert(0, str(ROOT / "scripts"))

from measure_wmf_formula_glyphs import inspect_docx  # noqa: E402


def fnum(value: Any) -> float | None:
    try:
        out = float(value)
    except (TypeError, ValueError):
        return None
    return out if math.isfinite(out) else None


def round3(value: float | None) -> float | None:
    return None if value is None else round(value, 3)


def percentile(values: list[float], q: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    idx = min(len(ordered) - 1, max(0, round((len(ordered) - 1) * q)))
    return ordered[idx]


def stats(values: list[Any]) -> dict[str, Any]:
    nums = [float(v) for v in (fnum(value) for value in values) if v is not None]
    if not nums:
        return {"n": 0}
    return {
        "n": len(nums),
        "avg": round3(sum(nums) / len(nums)),
        "median": round3(st.median(nums)),
        "p10": round3(percentile(nums, 0.10)),
        "p90": round3(percentile(nums, 0.90)),
        "min": round3(min(nums)),
        "max": round3(max(nums)),
    }


def bucket(value: float | None, step: float = 0.25) -> float | None:
    if value is None or step <= 0:
        return None
    return round(round(value / step) * step, 3)


def mode_stats(values: list[Any], step: float = 0.25, limit: int = 5) -> list[dict[str, Any]]:
    counts: dict[float, int] = {}
    for value in values:
        num = fnum(value)
        rounded = bucket(num, step)
        if rounded is None:
            continue
        counts[rounded] = counts.get(rounded, 0) + 1
    return [
        {"value": value, "count": count}
        for value, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:limit]
    ]


def json_counter_stats(values: list[Any], limit: int = 12) -> list[dict[str, Any]]:
    counts: dict[str, int] = {}
    formula_counts: dict[str, int] = {}
    for value in values:
        if not value:
            continue
        try:
            parsed = json.loads(str(value))
        except json.JSONDecodeError:
            continue
        for key, count in parsed.items():
            counts[key] = counts.get(key, 0) + int(count or 0)
            formula_counts[key] = formula_counts.get(key, 0) + 1
    return [
        {"function": key, "recordCount": count, "formulaCount": formula_counts.get(key, 0)}
        for key, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:limit]
    ]


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def source_report_by_object(path: Path | None) -> dict[int, dict[str, Any]]:
    if not path or not path.exists():
        return {}
    data = load_json(path)
    out: dict[int, dict[str, Any]] = {}
    for ordinal, item in enumerate(data.get("equations") or [], 1):
        index = item.get("docObjectIndex")
        if index in (None, ""):
            continue
        try:
            out[int(index)] = item
        except (TypeError, ValueError):
            continue
    return out


def has_nested_command(latex: str, command: str) -> bool:
    cursor = 0
    while True:
        start = latex.find(command, cursor)
        if start < 0:
            return False
        first = latex.find("{", start + len(command))
        if first < 0:
            return False
        first_end = matching(latex, first, "{", "}")
        if first_end < 0:
            return False
        second = latex.find("{", first_end + 1)
        if second < 0:
            return False
        second_end = matching(latex, second, "{", "}")
        if second_end < 0:
            return False
        body = latex[first + 1:first_end] + "\n" + latex[second + 1:second_end]
        if command in body:
            return True
        cursor = second_end + 1


def matching(text: str, start: int, open_ch: str, close_ch: str) -> int:
    depth = 0
    for i in range(start, len(text)):
        ch = text[i]
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return i
    return -1


FRACTION_COMMANDS = (r"\frac", r"\dfrac", r"\cfrac")


def latex_command_boundary(text: str, cursor: int) -> bool:
    return cursor >= len(text) or not text[cursor].isalpha()


def next_fraction_command(text: str, cursor: int = 0) -> int:
    best = -1
    for command in FRACTION_COMMANDS:
        hit = cursor
        while True:
            hit = text.find(command, hit)
            if hit < 0:
                break
            if latex_command_boundary(text, hit + len(command)):
                if best < 0 or hit < best:
                    best = hit
                break
            hit += len(command)
    return best


def fraction_command_end(text: str, start: int) -> int:
    for command in FRACTION_COMMANDS:
        if text.startswith(command, start) and latex_command_boundary(text, start + len(command)):
            return start + len(command)
    return start + len(r"\frac")


def has_fraction_command(text: str) -> bool:
    return next_fraction_command(text or "", 0) >= 0


ARRAY_ENVIRONMENTS = (
    "array",
    "aligned",
    "alignedat",
    "gathered",
    "matrix",
    "pmatrix",
    "bmatrix",
    "cases",
)


def has_array_environment(latex: str) -> bool:
    return any(r"\begin{" + env + "}" in latex for env in ARRAY_ENVIRONMENTS)


def fraction_inside_script(latex: str) -> bool:
    text = latex or ""
    cursor = 0
    while cursor < len(text):
        op = next_top_level_script_operator(text, cursor)
        if op < 0:
            return False
        body, end = read_script_body(text, op)
        if body and has_fraction_command(body):
            return True
        cursor = max(end, op + 1)
    return False


def next_top_level_script_operator(text: str, start: int = 0) -> int:
    depth = 0
    i = max(0, start)
    while i < len(text):
        ch = text[i]
        if ch == "\\":
            i += 1
            while i < len(text) and text[i].isalpha():
                i += 1
            continue
        if ch in "{[(":
            depth += 1
        elif ch in "}])":
            depth = max(0, depth - 1)
        elif depth == 0 and ch in "^_":
            return i
        i += 1
    return -1


def read_script_body(text: str, operator_index: int) -> tuple[str, int]:
    cursor = operator_index + 1
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    if cursor >= len(text):
        return "", cursor
    if text[cursor] == "{":
        end = matching(text, cursor, "{", "}")
        return (text[cursor + 1:end], end + 1) if end >= 0 else ("", cursor + 1)
    if text[cursor] == "\\":
        command_end = cursor + 1
        while command_end < len(text) and text[command_end].isalpha():
            command_end += 1
        body_end = command_end
        if has_fraction_command(text[cursor:command_end]):
            first = text.find("{", command_end)
            if first >= 0:
                first_end = matching(text, first, "{", "}")
                second = text.find("{", first_end + 1) if first_end >= 0 else -1
                second_end = matching(text, second, "{", "}") if second >= 0 else -1
                if second_end >= 0:
                    body_end = second_end + 1
        return text[cursor:body_end], body_end
    return text[cursor:cursor + 1], cursor + 1


def visible_latex_length(text: str) -> int:
    visible = re.sub(r"\\(?:mathrm|mathbf|mathit|textit|textbf|emph|text|boldsymbol)\s*\{\s*([^{}]*)\s*}", r"\1", text)
    visible = re.sub(r"\\[A-Za-z]+", "x", visible)
    visible = re.sub(r"[{}\s]", "", visible)
    return len(visible)


def contains_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text or "")


def text_heavy_fraction(latex: str) -> bool:
    cursor = 0
    while True:
        start = next_fraction_command(latex, cursor)
        if start < 0:
            break
        first = latex.find("{", fraction_command_end(latex, start))
        if first < 0:
            break
        first_end = matching(latex, first, "{", "}")
        if first_end < 0:
            break
        second = latex.find("{", first_end + 1)
        if second < 0:
            break
        second_end = matching(latex, second, "{", "}")
        if second_end < 0:
            break
        numerator = latex[first + 1:first_end]
        denominator = latex[second + 1:second_end]
        if (contains_cjk(numerator) and visible_latex_length(numerator) >= 6) or (
            contains_cjk(denominator) and visible_latex_length(denominator) >= 6
        ):
            return True
        cursor = second_end + 1
    return False


def has_nested_fraction(latex: str) -> bool:
    cursor = 0
    while True:
        start = next_fraction_command(latex, cursor)
        if start < 0:
            return False
        first = latex.find("{", fraction_command_end(latex, start))
        if first < 0:
            return False
        first_end = matching(latex, first, "{", "}")
        if first_end < 0:
            return False
        second = latex.find("{", first_end + 1)
        if second < 0:
            return False
        second_end = matching(latex, second, "{", "}")
        if second_end < 0:
            return False
        body = latex[first + 1:first_end] + "\n" + latex[second + 1:second_end]
        if has_fraction_command(body):
            return True
        cursor = second_end + 1


def has_top_level_fraction(latex: str) -> bool:
    text = latex or ""
    depth = 0
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "\\":
            command_end = i + 1
            while command_end < len(text) and text[command_end].isalpha():
                command_end += 1
            command = text[i:command_end]
            if depth == 0 and command in FRACTION_COMMANDS and latex_command_boundary(text, command_end):
                return True
            i = command_end
            continue
        if ch in "{[(":
            depth += 1
        elif ch in "}])":
            depth = max(0, depth - 1)
        i += 1
    return False


def classify_structure(latex: str, item: dict[str, Any]) -> str:
    latex = latex or ""
    if has_array_environment(latex):
        return "array"
    has_fraction = has_fraction_command(latex)
    if has_fraction and text_heavy_fraction(latex):
        return "text_fraction"
    if r"\sqrt" in latex and has_fraction:
        return "sqrt_fraction"
    if r"\sqrt" in latex:
        return "sqrt"
    has_script_fraction = fraction_inside_script(latex)
    if has_script_fraction and not has_top_level_fraction(latex):
        return "script_fraction"
    if has_script_fraction and has_top_level_fraction(latex):
        return "script_fraction_mixed"
    if has_fraction and has_nested_fraction(latex):
        return "nested_fraction"
    if has_fraction:
        return "fraction"
    if re.search(r"(?<!\\)[_^]", latex):
        return "script"
    if any(token in latex for token in (r"\overline", r"\underline", r"\overset", r"\underset")):
        return "accent"
    runs = ((item.get("wmf") or {}).get("runs") or [])
    if any((run.get("fontFace") or "").lower() == "symbol" for run in runs):
        return "symbol_linear"
    return "linear"


def formula_metrics(item: dict[str, Any], latex: str) -> dict[str, Any]:
    wmf = item.get("wmf") or {}
    runs = wmf.get("runs") or []
    polylines = wmf.get("polylines") or []
    record_function_counts = wmf.get("recordFunctionCounts") or {}
    fonts = [fnum(run.get("fontHeightPt")) for run in runs]
    fonts = [value for value in fonts if value is not None and value > 0]
    main_font = max(fonts) if fonts else None
    script_fonts = [value for value in fonts if main_font is not None and value < main_font * 0.88]
    baselines = [fnum(run.get("baselinePt")) for run in runs]
    baselines = [value for value in baselines if value is not None]
    advances = [fnum(run.get("advanceWidthPt")) for run in runs]
    advances = [value for value in advances if value is not None]
    rights = [fnum(run.get("rightEdgePt")) for run in runs]
    rights = [value for value in rights if value is not None]
    lefts = [fnum(run.get("xPt")) for run in runs]
    lefts = [value for value in lefts if value is not None]
    shape_w = fnum(item.get("shapeWidthPt"))
    shape_h = fnum(item.get("shapeHeightPt"))
    ink = item.get("ink") or {}
    ink_count = fnum(ink.get("inkCount"))
    ink_w = fnum(ink.get("magickInkWidthPt"))
    ink_h = fnum(ink.get("magickInkHeightPt"))
    if ink_w is None:
        ink_w = fnum(item.get("magickInkWidthPt"))
    if ink_h is None:
        ink_h = fnum(item.get("magickInkHeightPt"))
    if ink_count == 0:
        ink_w = None
        ink_h = None
    ink_left = fnum(ink.get("magickInkLeftPt"))
    ink_top = fnum(ink.get("magickInkTopPt"))
    ink_right = fnum(ink.get("magickInkRightPt"))
    ink_bottom = fnum(ink.get("magickInkBottomPt"))
    ink_center_y = fnum(ink.get("magickInkCenterYPt"))
    if ink_count == 0:
        ink_left = ink_top = ink_right = ink_bottom = ink_center_y = None
    baseline_min = min(baselines) if baselines else None
    baseline_max = max(baselines) if baselines else None
    record_left = min(lefts) if lefts else None
    record_right = max(rights) if rights else None
    record_width = (record_right - record_left) if record_left is not None and record_right is not None else None
    baseline_center = (baseline_min + baseline_max) / 2 if baseline_min is not None and baseline_max is not None else None
    top_margin = baseline_min - max(fonts) if baseline_min is not None and fonts else None
    bottom_margin = shape_h - baseline_max if shape_h is not None and baseline_max is not None else None
    horizontal_lines: list[dict[str, Any]] = []
    for polyline in polylines:
        points = polyline.get("points") or []
        if len(points) < 2:
            continue
        xs = [fnum(point.get("xPt")) for point in points]
        ys = [fnum(point.get("yPt")) for point in points]
        xs = [value for value in xs if value is not None]
        ys = [value for value in ys if value is not None]
        if len(xs) < 2 or len(ys) < 2:
            continue
        y_span = max(ys) - min(ys)
        x_span = max(xs) - min(xs)
        if y_span <= 0.08 and x_span >= 1.0:
            horizontal_lines.append({"xSpanPt": x_span, "yPt": sum(ys) / len(ys)})
    line_ys = [line["yPt"] for line in horizontal_lines]
    line_spans = [line["xSpanPt"] for line in horizontal_lines]
    line_min_y = min(line_ys) if line_ys else None
    line_max_y = max(line_ys) if line_ys else None
    line_center = (line_min_y + line_max_y) / 2 if line_min_y is not None and line_max_y is not None else None
    return {
        "objectIndex": item.get("objectIndex"),
        "structure": classify_structure(latex, item),
        "latex": latex,
        "shapeWidthPt": round3(shape_w),
        "shapeHeightPt": round3(shape_h),
        "shapeHeightBucketPt": bucket(shape_h, 0.25),
        "wmfWidthPt": item.get("wmfPlaceableWidthPt"),
        "wmfHeightPt": item.get("wmfPlaceableHeightPt"),
        "magickInkWidthPt": round3(ink_w),
        "magickInkHeightPt": round3(ink_h),
        "magickInkLeftPt": round3(ink_left),
        "magickInkTopPt": round3(ink_top),
        "magickInkRightPt": round3(ink_right),
        "magickInkBottomPt": round3(ink_bottom),
        "magickInkCenterYPt": round3(ink_center_y),
        "magickInkCount": int(ink_count) if ink_count is not None else None,
        "magickInkWidthRatio": round3(ink_w / shape_w) if ink_w is not None and shape_w else None,
        "magickInkHeightRatio": round3(ink_h / shape_h) if ink_h is not None and shape_h else None,
        "magickInkCenterYRatio": round3(ink_center_y / shape_h) if ink_center_y is not None and shape_h else None,
        "magickInkError": item.get("magickInkError") or ink.get("error"),
        "recordWidthPt": round3(record_width),
        "recordFillRatio": round3(record_width / shape_w) if record_width is not None and shape_w else None,
        "mainFontPt": round3(main_font),
        "mainFontBucketPt": bucket(main_font, 0.25),
        "scriptFontPt": round3(st.median(script_fonts)) if script_fonts else None,
        "scriptFontRatio": round3(st.median(script_fonts) / main_font) if script_fonts and main_font else None,
        "baselineMinPt": round3(baseline_min),
        "baselineMaxPt": round3(baseline_max),
        "baselineSpanPt": round3((baseline_max - baseline_min) if baseline_min is not None and baseline_max is not None else None),
        "baselineCenterPt": round3(baseline_center),
        "baselineMinRatio": round3(baseline_min / shape_h) if baseline_min is not None and shape_h else None,
        "baselineMaxRatio": round3(baseline_max / shape_h) if baseline_max is not None and shape_h else None,
        "baselineCenterRatio": round3(baseline_center / shape_h) if baseline_center is not None and shape_h else None,
        "fontToHeightRatio": round3(main_font / shape_h) if main_font is not None and shape_h else None,
        "topMarginPt": round3(top_margin),
        "bottomMarginPt": round3(bottom_margin),
        "runCount": len(runs),
        "polylineCount": len(polylines),
        "recordFunctionCounts": json.dumps(record_function_counts, ensure_ascii=False, sort_keys=True),
        "horizontalLineCount": len(horizontal_lines),
        "horizontalLineMedianWidthPt": round3(st.median(line_spans)) if line_spans else None,
        "horizontalLineMinYPt": round3(line_min_y),
        "horizontalLineMaxYPt": round3(line_max_y),
        "horizontalLineSpanPt": round3((line_max_y - line_min_y) if line_min_y is not None and line_max_y is not None else None),
        "horizontalLineCenterRatio": round3(line_center / shape_h) if line_center is not None and shape_h else None,
        "advanceMedianPt": round3(st.median(advances)) if advances else None,
        "advanceSumPt": round3(sum(advances)) if advances else None,
        "text": "".join(str(run.get("text") or "") for run in runs),
    }


def summarize_group(group: list[dict[str, Any]]) -> dict[str, Any]:
    ink_samples = [row for row in group if row.get("magickInkWidthPt") is not None or row.get("magickInkHeightPt") is not None]
    ink_errors = [row for row in group if row.get("magickInkError")]
    blank_ink = [row for row in group if row.get("magickInkCount") == 0]
    return {
        "count": len(group),
        "magickInkSampleCount": len(ink_samples),
        "magickInkErrorCount": len(ink_errors),
        "magickBlankInkCount": len(blank_ink),
        "magickInkCoverageRatio": round3(len(ink_samples) / len(group)) if group else None,
        "shapeWidthPt": stats([row.get("shapeWidthPt") for row in group]),
        "shapeHeightPt": stats([row.get("shapeHeightPt") for row in group]),
        "shapeHeightModes": mode_stats([row.get("shapeHeightPt") for row in group]),
        "recordFillRatio": stats([row.get("recordFillRatio") for row in group]),
        "mainFontPt": stats([row.get("mainFontPt") for row in group]),
        "mainFontModes": mode_stats([row.get("mainFontPt") for row in group]),
        "scriptFontRatio": stats([row.get("scriptFontRatio") for row in group]),
        "scriptRatioModes": mode_stats([row.get("scriptFontRatio") for row in group], 0.01),
        "baselineSpanPt": stats([row.get("baselineSpanPt") for row in group]),
        "baselineCenterRatio": stats([row.get("baselineCenterRatio") for row in group]),
        "fontToHeightRatio": stats([row.get("fontToHeightRatio") for row in group]),
        "horizontalLineCount": stats([row.get("horizontalLineCount") for row in group]),
        "horizontalLineCenterRatio": stats([row.get("horizontalLineCenterRatio") for row in group]),
        "horizontalLineMedianWidthPt": stats([row.get("horizontalLineMedianWidthPt") for row in group]),
        "recordFunctionModes": json_counter_stats([row.get("recordFunctionCounts") for row in group]),
        "topMarginPt": stats([row.get("topMarginPt") for row in group]),
        "bottomMarginPt": stats([row.get("bottomMarginPt") for row in group]),
        "advanceMedianPt": stats([row.get("advanceMedianPt") for row in group]),
        "magickInkWidthPt": stats([row.get("magickInkWidthPt") for row in group]),
        "magickInkHeightPt": stats([row.get("magickInkHeightPt") for row in group]),
        "magickInkWidthRatio": stats([row.get("magickInkWidthRatio") for row in group]),
        "magickInkHeightRatio": stats([row.get("magickInkHeightRatio") for row in group]),
        "magickInkCenterYRatio": stats([row.get("magickInkCenterYRatio") for row in group]),
        "runCount": stats([row.get("runCount") for row in group]),
    }


def read_manifest_items(path: Path, start: int, end: int) -> list[dict[str, Any]]:
    items = load_json(path)
    selected = []
    for item in items:
        index = int(item.get("index") or 0)
        if index >= start and (end <= 0 or index <= end):
            selected.append(item)
    return selected


def find_source_report(index: int, search_roots: list[Path]) -> Path | None:
    names = [f"source-report-{index}.json", f"{index}.source-report.json"]
    for root in search_roots:
        if not root.exists():
            continue
        for name in names:
            direct = root / name
            if direct.exists():
                return direct
        matches = sorted(root.rglob(f"source-report-{index}.json"), key=lambda p: p.stat().st_mtime, reverse=True)
        if matches:
            return matches[0]
    return None


def build_source_report(index: int, source_docx: Path, tex: Path, out: Path) -> Path:
    out.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        str(ROOT / "scripts" / "build_xsc_source_report.py"),
        "--index",
        str(index),
        "--tex",
        str(tex),
        "--source-docx",
        str(source_docx),
        "--out",
        str(out),
    ]
    subprocess.run(cmd, cwd=ROOT, check=True)
    return out


def inspect_source_docx(docx: Path, with_ink: bool = False, require_magick_ink: bool = False) -> dict[str, Any]:
    args = argparse.Namespace(
        docx=docx,
        request=None,
        indices=None,
        max_items=None,
        with_ink=with_ink,
        require_magick_ink=require_magick_ink,
        magick=Path(r"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe"),
        white_threshold=245,
        alpha_threshold=8,
        fallback_units_per_pt=20.0,
    )
    report = inspect_docx(args)
    if require_magick_ink:
        errors = [
            item.get("magickInkError")
            for item in report.get("formulas") or []
            if item.get("magickInkError")
        ]
        if errors:
            raise RuntimeError("magick ink measurement required but failed: " + ", ".join(sorted(set(errors))))
    return report


def collect_rows_for_doc(
    index: int,
    docx: Path,
    source_report: Path | None,
    with_ink: bool = False,
    require_magick_ink: bool = False,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    report = inspect_source_docx(docx, with_ink, require_magick_ink)
    return collect_rows_from_inspected_report(index, docx, report, source_report)


def collect_rows_from_inspected_report(
    index: int,
    docx: Path,
    report: dict[str, Any],
    source_report: Path | None,
    source_kind: str = "source-docx",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    latex_by_object = source_report_by_object(source_report)
    rows = []
    skipped = []
    for item in report.get("formulas") or []:
        object_index = int(item.get("objectIndex") or 0)
        source_item = latex_by_object.get(object_index, {})
        latex = source_item.get("output") or ""
        if not source_report or not latex:
            skipped.append(
                {
                    "docIndex": index,
                    "sourceDocx": str(docx),
                    "sourceReport": str(source_report) if source_report else "",
                    "objectIndex": object_index,
                    "reason": "missingSourceReport" if not source_report else "missingObjectMapping",
                    "sourceKind": source_kind,
                }
            )
            continue
        row = formula_metrics(item, latex)
        row["docIndex"] = index
        row["sourceReport"] = str(source_report) if source_report else ""
        row["sourceDocx"] = str(docx)
        row["sourceKind"] = source_kind
        row["sourceStatus"] = source_item.get("status") or ""
        rows.append(row)
    return rows, skipped


def summarize_rows(
    rows: list[dict[str, Any]],
    missing: list[dict[str, Any]] | None = None,
    skipped: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        groups.setdefault(row.get("structure") or "unknown", []).append(row)
    by_structure = {}
    by_structure_height: dict[str, dict[str, Any]] = {}
    for structure, group in sorted(groups.items()):
        by_structure[structure] = summarize_group(group)
        height_groups: dict[str, list[dict[str, Any]]] = {}
        for row in group:
            key = str(row.get("shapeHeightBucketPt") or "unknown")
            height_groups.setdefault(key, []).append(row)
        by_structure_height[structure] = {
            key: summarize_group(height_group)
            for key, height_group in sorted(
                height_groups.items(),
                key=lambda item: (-len(item[1]), item[0]),
            )[:8]
        }
    return {
        "formulaCount": len(rows),
        "missingSourceCount": len(missing or []),
        "missingSources": missing or [],
        "skippedUnmappedWmfCount": len(skipped or []),
        "skippedUnmappedWmfs": skipped or [],
        "magickInkSampleCount": sum(1 for row in rows if row.get("magickInkWidthPt") is not None),
        "magickInkErrorCount": sum(1 for row in rows if row.get("magickInkError")),
        "structures": by_structure,
        "structureHeightBuckets": by_structure_height,
        "worstLowTopMargin": sorted(
            rows,
            key=lambda row: fnum(row.get("topMarginPt")) if fnum(row.get("topMarginPt")) is not None else 999,
        )[:25],
        "worstLowBottomMargin": sorted(
            rows,
            key=lambda row: fnum(row.get("bottomMarginPt")) if fnum(row.get("bottomMarginPt")) is not None else 999,
        )[:25],
    }


def source_report_paths(roots: list[Path]) -> list[Path]:
    out: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        if not root.exists():
            continue
        candidates = [root] if root.is_file() else root.rglob("source-report-*.json")
        for path in candidates:
            if not path.is_file():
                continue
            resolved = str(path.resolve())
            if resolved in seen:
                continue
            seen.add(resolved)
            out.append(path)
    return out


def scan_structure_coverage(roots: list[Path], limit_per_structure: int = 8) -> dict[str, Any]:
    probe_item = {"wmf": {"runs": []}}
    by_structure: dict[str, dict[str, Any]] = {}
    for path in source_report_paths(roots):
        try:
            data = load_json(path)
        except (OSError, json.JSONDecodeError, UnicodeDecodeError):
            continue
        equations = data.get("equations") if isinstance(data, dict) else None
        if not isinstance(equations, list):
            continue
        for item in equations:
            if not isinstance(item, dict):
                continue
            latex = item.get("output") or ""
            if not latex:
                continue
            structure = classify_structure(latex, probe_item)
            bucket = by_structure.setdefault(
                structure,
                {
                    "count": 0,
                    "reports": set(),
                    "samples": [],
                },
            )
            bucket["count"] += 1
            bucket["reports"].add(str(path))
            if len(bucket["samples"]) < limit_per_structure:
                bucket["samples"].append(
                    {
                        "report": str(path),
                        "source": data.get("source") or "",
                        "docObjectIndex": item.get("docObjectIndex"),
                        "latex": latex[:240],
                    }
                )
    return {
        structure: {
            "count": value["count"],
            "reportCount": len(value["reports"]),
            "samples": value["samples"],
        }
        for structure, value in sorted(by_structure.items())
    }


def candidate_number(stats_map: dict[str, Any], field: str = "median") -> float | None:
    value = (stats_map or {}).get(field)
    return fnum(value)


JAVA_FAMILY_BY_STRUCTURE = {
    "linear": "LINEAR",
    "symbol_linear": None,
    "script": "SCRIPT",
    "script_fraction": "SCRIPT_FRACTION",
    "script_fraction_mixed": "SCRIPT_FRACTION_MIXED",
    "fraction": "ORDINARY_FRACTION",
    "nested_fraction": "NESTED_FRACTION",
    "text_fraction": "TEXT_FRACTION",
    "sqrt": "SQRT",
    "sqrt_fraction": "SQRT_FRACTION",
    "accent": "ACCENT",
    "array": "ARRAY",
}


def java_double(value: Any) -> str:
    numeric = fnum(value)
    if numeric is None:
        return "Double.NaN"
    text = f"{numeric:.3f}".rstrip("0").rstrip(".")
    if text == "-0":
        text = "0"
    if "." not in text:
        text += ".0"
    return text + "d"


def has_stats(stats_map: dict[str, Any], field: str) -> bool:
    return bool(((stats_map or {}).get(field) or {}).get("n"))


def structure_guidance(structure: str, stats_map: dict[str, Any], sample_status: str) -> dict[str, Any]:
    usable: list[str] = []
    missing: list[str] = []
    warnings: list[str] = []
    if sample_status != "missing":
        if stats_map.get("shapeHeightModes"):
            usable.append("dominantHeightPt")
        if has_stats(stats_map, "mainFontPt"):
            usable.append("mainFontPt")
        if has_stats(stats_map, "scriptFontRatio"):
            usable.append("scriptRatio")
        if has_stats(stats_map, "magickInkWidthRatio") or has_stats(stats_map, "magickInkHeightRatio"):
            usable.append("magickInkBBox")
        if (stats_map.get("magickInkSampleCount") or 0) < (stats_map.get("count") or 0):
            warnings.append("partial or missing Magick ink coverage; do not treat ink ratios as final")
    if structure in {"script_fraction_mixed", "fraction", "nested_fraction", "text_fraction", "sqrt_fraction"}:
        if has_stats(stats_map, "horizontalLineCenterRatio"):
            usable.append("horizontalLineCenterRatio")
        else:
            missing.append("barOrRuleGeometry")
            warnings.append("source WMF line geometry is not decoded enough for fraction/radical bar placement")
    if structure in {"sqrt", "sqrt_fraction"}:
        missing.extend(["radicalCheckmarkGeometry", "rootIndexPlacement"])
    if structure in {"script", "script_fraction", "script_fraction_mixed"} and not has_stats(stats_map, "scriptFontRatio"):
        missing.append("scriptFontRatio")
    if structure == "script_fraction_mixed":
        missing.append("scriptSlotToTopLevelFractionBaseline")
    if structure == "array":
        missing.extend(["columnSpacing", "rowBaselineOffsets", "braceGeometry"])
    if structure == "text_fraction":
        missing.extend(["textNumeratorBaseline", "textDenominatorBaseline", "contentSplitPolicy"])
    if sample_status == "thin":
        warnings.append("thin sample count; confirm with ink/Word preview before treating as final")
    if sample_status == "missing":
        warnings.append("no joined source WMF rows; use coverage samples only to choose reference formulas")
    next_steps = {
        "linear": "compare standard glyph ink bbox against source rows before changing global font or dx",
        "symbol_linear": "separate Symbol-heavy labels from ordinary Times text before applying width scales",
        "script": "derive base/script glyph ink boxes and script x/y offsets from matched source samples",
        "script_fraction": "collect more script-slot fraction references and compare slash policy against Word ink",
        "script_fraction_mixed": "separate top-level fraction bar geometry from script-slot slash fraction geometry",
        "fraction": "decode META 0x0626 or rendered ink for numerator/bar/denominator offsets",
        "nested_fraction": "measure outer and inner bar/denominator baselines separately before changing split constants",
        "text_fraction": "decide content split versus true stacked CJK fraction from source Word/ink evidence",
        "sqrt": "measure radical rule/checkmark ink bbox from TeXToggle references before changing current constants",
        "sqrt_fraction": "measure root body offset and nested fraction baselines together, not as generic nested_fraction",
        "accent": "verify overline/underline y from ink because source records do not expose enough line geometry",
        "array": "derive row heights and brace/column spacing from array-like source samples",
    }
    return {
        "usableRendererFields": usable,
        "missingRendererParameters": sorted(set(missing)),
        "warnings": warnings,
        "nextCalibrationStep": next_steps.get(structure, ""),
    }


def dominant_height_bucket(summary: dict[str, Any], structure: str) -> tuple[str | None, dict[str, Any] | None]:
    buckets = summary.get("structureHeightBuckets", {}).get(structure, {})
    if not buckets:
        return None, None
    key, stats_map = next(iter(buckets.items()))
    return key, stats_map


def build_parameter_candidates(summary: dict[str, Any], coverage: dict[str, Any] | None = None) -> dict[str, Any]:
    structures = summary.get("structures") or {}
    linear = structures.get("linear") or {}
    script = structures.get("script") or {}
    fraction = structures.get("fraction") or {}
    nested = structures.get("nested_fraction") or {}
    accent = structures.get("accent") or {}
    symbol_linear = structures.get("symbol_linear") or {}
    base_font = candidate_number(linear.get("mainFontPt") or {}) or candidate_number(script.get("mainFontPt") or {})
    script_ratio = candidate_number(script.get("scriptFontRatio") or {}) or candidate_number(fraction.get("scriptFontRatio") or {})
    expected_structures = [
        "linear",
        "symbol_linear",
        "script",
        "script_fraction",
        "script_fraction_mixed",
        "fraction",
        "nested_fraction",
        "text_fraction",
        "sqrt",
        "sqrt_fraction",
        "accent",
        "array",
    ]
    renderer_actions = {
        "linear": "seed standard glyph height and ordinary text boxes",
        "symbol_linear": "seed mixed Symbol/Times linear handling",
        "script": "seed script physical box family; keep visual script ratio ink-validated",
        "script_fraction": "route to slash fraction only inside script context",
        "script_fraction_mixed": "seed mixed script-slot plus top-level fraction height separately from pure script_fraction",
        "fraction": "seed ordinary stacked fraction height family",
        "nested_fraction": "seed nested stacked fraction height family; sample count is small",
        "text_fraction": "do not force into slash/unified layout; prefer content split or separate calibration",
        "sqrt": "seed root physical height; keep radical geometry ink-validated",
        "sqrt_fraction": "seed root-with-fraction physical height; do not route through generic nested fraction",
        "accent": "seed overline/underline/accent height family",
        "array": "needs matrix/array-specific calibration outside scalar glyph ladder",
    }
    candidates: dict[str, Any] = {
        "sourceFormulaCount": summary.get("formulaCount"),
        "missingSourceCount": summary.get("missingSourceCount"),
        "skippedUnmappedWmfCount": summary.get("skippedUnmappedWmfCount"),
        "notes": [
            "Use these as renderer starting constants, not as an acceptance gate.",
            "Prefer dominant height buckets over global structure medians when choosing box families.",
            "recordFillRatio can exceed 1.0 in source MathType WMFs; do not use it alone as visible ink width.",
            "Source MathType text baselines are often anchored at zero; baseline ratios are diagnostic only.",
            "Current parser sees zero source polylines; fraction/radical bar positions need META 0x0626 study or ink bbox.",
            "coverageCount is discovery-only from source reports; use count/mapped rows for calibration strength.",
            "Magick ink bbox is closer to visible geometry than record advance, but Word rendering still needs spot checks.",
        ],
        "fieldTrust": {
            "trustedForRendererSeed": [
                "mainFontPt",
                "sourceScriptFontRatio",
                "dominantHeightPt",
                "heightModes",
                "mainFontModes",
                "magickInkWidthRatio",
                "magickInkHeightRatio",
                "magickInkCenterYRatio",
            ],
            "diagnosticOnly": [
                "recordFillRatio",
                "baselineCenterRatio",
                "horizontalLineCenterRatio",
                "horizontalLineMedianWidthPt",
                "topMarginPt",
                "bottomMarginPt",
                "coverageCount",
            ],
        },
        "standardGlyph": {
            "mainFontPt": round3(base_font),
            "sourceScriptFontRatio": round3(script_ratio),
            "scriptRatioPrimary": round3(script_ratio),
            "scriptFontPt": round3(base_font * script_ratio) if base_font and script_ratio else None,
            "visualScriptRatioRequiresInkValidation": True,
        },
        "structures": {},
        "sourceReportCoverage": coverage or {},
    }
    for structure in expected_structures:
        stats_map = structures.get(structure) or {}
        height_key, height_stats = dominant_height_bucket(summary, structure)
        count = int(stats_map.get("count") or 0)
        coverage_count = int(((coverage or {}).get(structure) or {}).get("count") or 0)
        if count <= 0:
            sample_status = "missing"
        elif count < 20:
            sample_status = "thin"
        else:
            sample_status = "usable"
        guidance = structure_guidance(structure, stats_map, sample_status)
        candidates["structures"][structure] = {
            "count": count,
            "coverageCount": coverage_count,
            "coverageSamples": ((coverage or {}).get(structure) or {}).get("samples") or [],
            "sampleStatus": sample_status,
            "rendererAction": renderer_actions.get(structure, ""),
            **guidance,
            "dominantHeightPt": fnum(height_key) if height_key not in (None, "unknown") else None,
            "dominantHeightStats": height_stats,
            "heightModes": stats_map.get("shapeHeightModes") or [],
            "mainFontPt": candidate_number(stats_map.get("mainFontPt") or {}),
            "mainFontModes": stats_map.get("mainFontModes") or [],
            "scriptRatio": candidate_number(stats_map.get("scriptFontRatio") or {}),
            "scriptRatioModes": stats_map.get("scriptRatioModes") or [],
            "baselineCenterRatio": candidate_number(stats_map.get("baselineCenterRatio") or {}),
            "horizontalLineCenterRatio": candidate_number(stats_map.get("horizontalLineCenterRatio") or {}),
            "horizontalLineMedianWidthPt": candidate_number(stats_map.get("horizontalLineMedianWidthPt") or {}),
            "recordFunctionModes": stats_map.get("recordFunctionModes") or [],
            "fontToHeightRatio": candidate_number(stats_map.get("fontToHeightRatio") or {}),
            "magickInkSampleCount": stats_map.get("magickInkSampleCount", 0),
            "magickInkErrorCount": stats_map.get("magickInkErrorCount", 0),
            "magickBlankInkCount": stats_map.get("magickBlankInkCount", 0),
            "magickInkCoverageRatio": stats_map.get("magickInkCoverageRatio"),
            "magickInkWidthRatio": candidate_number(stats_map.get("magickInkWidthRatio") or {}),
            "magickInkHeightRatio": candidate_number(stats_map.get("magickInkHeightRatio") or {}),
            "magickInkCenterYRatio": candidate_number(stats_map.get("magickInkCenterYRatio") or {}),
        }
    return candidates


def render_candidates(candidates: dict[str, Any]) -> str:
    lines = [
        "WMF renderer parameter candidates",
        "source formulas: {formulas} missing sources: {missing} skipped unmapped WMFs: {skipped}".format(
            formulas=candidates.get("sourceFormulaCount"),
            missing=candidates.get("missingSourceCount"),
            skipped=candidates.get("skippedUnmappedWmfCount"),
        ),
        "",
        "Standard glyph",
        json.dumps(candidates.get("standardGlyph"), ensure_ascii=False),
        "",
        "Field trust",
        json.dumps(candidates.get("fieldTrust"), ensure_ascii=False),
        "",
        "Structures",
    ]
    for name, item in (candidates.get("structures") or {}).items():
        lines.append(
            "  {name}: status={status} n={count} coverage={coverage} height={height} font={font} scriptRatio={script} baselineCenter={baseline} lineCenter={line} inkWH={ink_w}/{ink_h} inkCenterY={ink_y} inkSamples={ink_samples}/{count} inkErrors={ink_errors} blankInk={blank_ink}".format(
                name=name,
                status=item.get("sampleStatus"),
                count=item.get("count"),
                coverage=item.get("coverageCount"),
                height=item.get("dominantHeightPt"),
                font=item.get("mainFontPt"),
                script=item.get("scriptRatio"),
                baseline=item.get("baselineCenterRatio"),
                line=item.get("horizontalLineCenterRatio"),
                ink_w=item.get("magickInkWidthRatio"),
                ink_h=item.get("magickInkHeightRatio"),
                ink_y=item.get("magickInkCenterYRatio"),
                ink_samples=item.get("magickInkSampleCount"),
                ink_errors=item.get("magickInkErrorCount"),
                blank_ink=item.get("magickBlankInkCount"),
            )
        )
        modes = ", ".join(f"{entry['value']}({entry['count']})" for entry in (item.get("heightModes") or [])[:5])
        lines.append(f"    height modes: {modes}")
        if item.get("rendererAction"):
            lines.append(f"    renderer action: {item.get('rendererAction')}")
        if item.get("usableRendererFields"):
            lines.append("    usable fields: " + ", ".join(item.get("usableRendererFields") or []))
        if item.get("missingRendererParameters"):
            lines.append("    missing params: " + ", ".join(item.get("missingRendererParameters") or []))
        if item.get("nextCalibrationStep"):
            lines.append(f"    next calibration: {item.get('nextCalibrationStep')}")
        for warning in item.get("warnings") or []:
            lines.append(f"    warning: {warning}")
        for sample in (item.get("coverageSamples") or [])[:2]:
            lines.append(
                "    coverage sample: report={report} obj={obj} latex={latex}".format(
                    report=sample.get("report"),
                    obj=sample.get("docObjectIndex"),
                    latex=(sample.get("latex") or "")[:120],
                )
            )
        record_modes = ", ".join(
            f"{entry['function']}:{entry['formulaCount']}/{entry['recordCount']}"
            for entry in (item.get("recordFunctionModes") or [])[:6]
        )
        if record_modes:
            lines.append(f"    record functions: {record_modes}")
    lines.append("")
    lines.append("Notes")
    lines.extend(f"  - {note}" for note in candidates.get("notes") or [])
    return "\n".join(lines)


def render_candidates_java(candidates: dict[str, Any]) -> str:
    lines = [
        "// Generated from aggregate_wmf_structure_metrics.py parameter candidates.",
        "// Copy into MathTypeStructureMetrics.sourceSampleMetrics only after reviewing sampleStatus and sample counts.",
    ]
    for structure, item in (candidates.get("structures") or {}).items():
        family = JAVA_FAMILY_BY_STRUCTURE.get(structure)
        if family is None:
            continue
        height = item.get("dominantHeightPt")
        main_font = item.get("mainFontPt")
        script_ratio = item.get("scriptRatio")
        ink_w = item.get("magickInkWidthRatio")
        ink_h = item.get("magickInkHeightRatio")
        ink_y = item.get("magickInkCenterYRatio")
        ink_samples = int(item.get("magickInkSampleCount") or 0)
        status = item.get("sampleStatus") or "unknown"
        count = int(item.get("count") or 0)
        height_comment = "missing" if height is None else java_double(height)
        lines.extend(
            [
                f"// {structure}: status={status} count={count} inkSamples={ink_samples} candidateHeight={height_comment}",
                f"case {family} -> new SourceSampleMetrics(",
                "    family,",
                f"    {java_double(height)},",
                f"    {java_double(main_font)},",
                f"    {java_double(script_ratio)},",
                f"    new InkMetrics({java_double(ink_w)}, {java_double(ink_h)}, {java_double(ink_y)}, {ink_samples})",
                ");",
            ]
        )
    return "\n".join(lines) + "\n"


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    keys: list[str] = []
    for row in rows:
        for key in row:
            if key not in keys:
                keys.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=keys)
        writer.writeheader()
        writer.writerows(rows)


def render_text(summary: dict[str, Any]) -> str:
    lines = [
        "WMF structure metrics",
        f"formulas: {summary['formulaCount']}",
        f"missing sources: {summary.get('missingSourceCount', 0)}",
        f"skipped unmapped WMFs: {summary.get('skippedUnmappedWmfCount', 0)}",
        "",
        "By structure",
    ]
    for name, stats_map in summary["structures"].items():
        lines.append(
            "  {name}: n={n} mainFont={font} shapeH={height} fill={fill} scriptRatio={script}".format(
                name=name,
                n=stats_map["count"],
                font=stats_map["mainFontPt"],
                height=stats_map["shapeHeightPt"],
                fill=stats_map["recordFillRatio"],
                script=stats_map["scriptFontRatio"],
            )
        )
        modes = ", ".join(f"{item['value']}({item['count']})" for item in stats_map.get("shapeHeightModes", [])[:4])
        lines.append(f"    height modes: {modes}")
        if stats_map.get("horizontalLineCenterRatio", {}).get("n"):
            lines.append(
                "    lineCenter={line} lineWidth={width}".format(
                    line=stats_map["horizontalLineCenterRatio"],
                    width=stats_map["horizontalLineMedianWidthPt"],
                )
            )
        record_modes = ", ".join(
            f"{entry['function']}:{entry['formulaCount']}/{entry['recordCount']}"
            for entry in stats_map.get("recordFunctionModes", [])[:6]
        )
        if record_modes:
            lines.append(f"    record functions: {record_modes}")
    lines.append("")
    lines.append("Dominant height buckets")
    for structure, height_map in summary.get("structureHeightBuckets", {}).items():
        rendered = []
        for height, stats_map in list(height_map.items())[:4]:
            rendered.append(
                "{height}pt:n={count},font={font},script={script},line={line}".format(
                    height=height,
                    count=stats_map["count"],
                    font=stats_map["mainFontPt"].get("median"),
                    script=stats_map["scriptFontRatio"].get("median"),
                    line=stats_map["horizontalLineCenterRatio"].get("median"),
                )
            )
        lines.append(f"  {structure}: " + "; ".join(rendered))
    lines.append("")
    lines.append("Lowest top margins")
    for row in summary["worstLowTopMargin"][:10]:
        lines.append(
            "  doc={doc} obj={obj} {structure} top={top} bottom={bottom} font={font} latex={latex}".format(
                doc=row.get("docIndex"),
                obj=row.get("objectIndex"),
                structure=row.get("structure"),
                top=row.get("topMarginPt"),
                bottom=row.get("bottomMarginPt"),
                font=row.get("mainFontPt"),
                latex=(row.get("latex") or row.get("text") or "")[:90],
            )
        )
    return "\n".join(lines)


def resolve_source_docx(item: dict[str, Any], search_roots: list[Path]) -> Path | None:
    source = Path(item.get("sourcePath") or "")
    if source.exists():
        return source
    source_name = item.get("sourceName") or source.name
    if not source_name:
        return None
    for root in search_roots:
        if not root.exists():
            continue
        direct = root / source_name
        if direct.exists():
            return direct
        matches = sorted(root.rglob(source_name), key=lambda p: p.stat().st_mtime, reverse=True)
        if matches:
            return matches[0]
    return None


def run_self_test() -> None:
    item = {"wmf": {"runs": []}}
    assert not has_fraction_command(r"\fraction{x}")
    assert classify_structure(r"\fraction{x}", item) == "linear"
    assert classify_structure(r"x^{\frac{1}{2}}", item) == "script_fraction"
    assert classify_structure(r"x_{\frac{1}{4}}=\frac{1}{4}y", item) == "script_fraction_mixed"
    assert classify_structure(r"\frac{x^{\frac{1}{2}}}{2}", item) == "nested_fraction"
    assert classify_structure(r"\frac{1+\dfrac{a}{b}}{2}", item) == "nested_fraction"
    assert classify_structure(r"\sqrt{1+\frac{a}{b}}", item) == "sqrt_fraction"
    assert classify_structure(r"\begin{array}{c}1\\2\end{array}", item) == "array"
    assert classify_structure(r"\begin{pmatrix}1&2\\3&4\end{pmatrix}+\frac{1}{2}", item) == "array"
    assert classify_structure(r"\begin{cases}x=1\\y=\frac{1}{2}\end{cases}", item) == "array"

    candidates = build_parameter_candidates({"formulaCount": 0, "missingSourceCount": 0, "structures": {}})
    structures = candidates["structures"]
    for name in ("script_fraction", "script_fraction_mixed", "sqrt", "sqrt_fraction", "array"):
        assert structures[name]["sampleStatus"] == "missing"
        assert structures[name]["rendererAction"]
        assert structures[name]["missingRendererParameters"]
        assert structures[name]["nextCalibrationStep"]
    assert "insufficient current corpus" not in structures["sqrt"]["rendererAction"]

    seeded = build_parameter_candidates(
        {
            "formulaCount": 2,
            "missingSourceCount": 0,
            "structures": {
                "sqrt": {
                    "count": 2,
                    "shapeHeightModes": [{"value": 18.0, "count": 2}],
                    "mainFontPt": {"n": 2, "median": 12.0},
                }
            },
            "structureHeightBuckets": {"sqrt": {"18.0": {"count": 2}}},
        }
    )
    assert "dominantHeightPt" in seeded["structures"]["sqrt"]["usableRendererFields"]
    assert "radicalCheckmarkGeometry" in seeded["structures"]["sqrt"]["missingRendererParameters"]

    blank_ink = formula_metrics(
        {
            "objectIndex": 1,
            "shapeWidthPt": 20,
            "shapeHeightPt": 10,
            "wmf": {"runs": []},
            "ink": {"inkCount": 0, "magickInkWidthPt": 0, "magickInkHeightPt": 0},
            "magickInkWidthPt": 0,
            "magickInkHeightPt": 0,
        },
        r"x",
    )
    assert blank_ink["magickInkWidthPt"] is None
    assert blank_ink["magickInkHeightPt"] is None
    assert blank_ink["magickInkWidthRatio"] is None
    ink_seeded = build_parameter_candidates(
        {
            "formulaCount": 2,
            "missingSourceCount": 0,
            "structures": {
                "linear": {
                    "count": 2,
                    "magickInkSampleCount": 2,
                    "shapeHeightModes": [{"value": 15.0, "count": 2}],
                    "mainFontPt": {"n": 2, "median": 12.0},
                    "magickInkWidthRatio": {"n": 2, "median": 0.58},
                    "magickInkHeightRatio": {"n": 2, "median": 0.72},
                    "magickInkCenterYRatio": {"n": 2, "median": 0.56},
                }
            },
            "structureHeightBuckets": {"linear": {"15.0": {"count": 2}}},
        }
    )
    linear_seed = ink_seeded["structures"]["linear"]
    assert "magickInkBBox" in linear_seed["usableRendererFields"]
    assert linear_seed["magickInkWidthRatio"] == 0.58
    assert "magickInkWidthRatio" in ink_seeded["fieldTrust"]["trustedForRendererSeed"]
    java_snippet = render_candidates_java(ink_seeded)
    assert "case LINEAR -> new SourceSampleMetrics(" in java_snippet
    assert "new InkMetrics(0.58d, 0.72d, 0.56d, 2)" in java_snippet
    assert "case SYMBOL_LINEAR" not in java_snippet

    unmapped_rows, unmapped_skipped = collect_rows_from_inspected_report(
        12,
        Path("missing.docx"),
        {"formulas": [{"objectIndex": 1, "wmf": {"runs": []}, "shape": {"widthPt": 12, "heightPt": 13}}]},
        None,
        "generated-word-tex-toggle",
    )
    assert unmapped_rows == []
    assert unmapped_skipped[0]["reason"] == "missingSourceReport"
    assert unmapped_skipped[0]["sourceKind"] == "generated-word-tex-toggle"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-test", action="store_true", help="Run structure classifier self-checks and exit.")
    parser.add_argument("--docx", type=Path, help="Single source DOCX.")
    parser.add_argument("--source-report", type=Path, help="Source report JSON for --docx.")
    parser.add_argument("--manifest", type=Path, default=ROOT / "analysis" / "xsc-latex-fixed" / "manifest.json")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--end", type=int, default=0)
    parser.add_argument("--source-report-root", action="append", type=Path, default=[])
    parser.add_argument("--coverage-report-root", action="append", type=Path, default=[])
    parser.add_argument("--source-docx-root", action="append", type=Path, default=[])
    parser.add_argument(
        "--extra-docx-source-report",
        action="append",
        default=[],
        help="Append an extra reference pair as DOCX=SOURCE_REPORT, for generated official MathType samples.",
    )
    parser.add_argument("--build-source-reports", action="store_true")
    parser.add_argument("--built-source-report-dir", type=Path, default=ROOT / "analysis" / "wmf-structure-source-reports")
    parser.add_argument("--out-json", type=Path, default=ROOT / "analysis" / "wmf-structure-metrics" / "summary.json")
    parser.add_argument("--out-csv", type=Path, default=ROOT / "analysis" / "wmf-structure-metrics" / "formulas.csv")
    parser.add_argument("--out-text", type=Path, default=ROOT / "analysis" / "wmf-structure-metrics" / "summary.txt")
    parser.add_argument("--out-candidates-json", type=Path)
    parser.add_argument("--out-candidates-text", type=Path)
    parser.add_argument("--out-candidates-java", type=Path)
    parser.add_argument("--with-ink", action="store_true", help="Render WMFs through ImageMagick and aggregate ink bbox fields.")
    parser.add_argument("--require-magick-ink", action="store_true", help="Fail when --with-ink cannot measure every item.")
    args = parser.parse_args()

    if args.self_test:
        run_self_test()
        print("aggregate_wmf_structure_metrics self-test passed")
        return 0

    rows: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    if args.docx:
        doc_rows, doc_skipped = collect_rows_for_doc(
            0,
            args.docx,
            args.source_report,
            args.with_ink,
            args.require_magick_ink,
        )
        rows.extend(doc_rows)
        skipped.extend(doc_skipped)
    else:
        if args.start <= 0:
            raise SystemExit("--start is required when using --manifest")
        search_roots = args.source_report_root or [
            ROOT / "analysis",
            ROOT / "analysis" / "unattended-runs",
            args.built_source_report_dir,
        ]
        source_docx_roots = args.source_docx_root or [
            Path(r"E:\新加卷\新建文件夹\xsc资料"),
            ROOT / "analysis" / "xsc-numbered-dataset",
        ]
        for item in read_manifest_items(args.manifest, args.start, args.end):
            index = int(item["index"])
            source_docx = resolve_source_docx(item, source_docx_roots)
            if source_docx is None:
                missing.append(
                    {
                        "docIndex": index,
                        "sourcePath": item.get("sourcePath") or "",
                        "sourceName": item.get("sourceName") or "",
                    }
                )
                continue
            tex = Path(item["latexPath"])
            source_report = find_source_report(index, search_roots)
            if not source_report and args.build_source_reports:
                source_report = build_source_report(
                    index,
                    source_docx,
                    tex,
                    args.built_source_report_dir / f"source-report-{index}.json",
                )
            doc_rows, doc_skipped = collect_rows_for_doc(
                index,
                source_docx,
                source_report,
                args.with_ink,
                args.require_magick_ink,
            )
            rows.extend(doc_rows)
            skipped.extend(doc_skipped)

    for ordinal, pair in enumerate(args.extra_docx_source_report or [], 1):
        if "=" not in pair:
            raise SystemExit("--extra-docx-source-report must use DOCX=SOURCE_REPORT")
        docx_text, report_text = pair.split("=", 1)
        extra_docx = Path(docx_text)
        extra_report = Path(report_text)
        doc_rows, doc_skipped = collect_rows_from_inspected_report(
            -ordinal,
            extra_docx,
            inspect_source_docx(extra_docx, args.with_ink, args.require_magick_ink),
            extra_report,
            "generated-word-tex-toggle",
        )
        rows.extend(doc_rows)
        skipped.extend(doc_skipped)

    summary = summarize_rows(rows, missing, skipped)
    summary["withMagickInk"] = bool(args.with_ink)
    summary["requireMagickInk"] = bool(args.require_magick_ink)
    if args.coverage_report_root:
        coverage_roots = args.coverage_report_root
    elif args.docx:
        coverage_roots = []
    else:
        coverage_roots = search_roots
    coverage = scan_structure_coverage(coverage_roots) if coverage_roots else {}
    if coverage:
        summary["sourceReportCoverage"] = coverage
    args.out_json.parent.mkdir(parents=True, exist_ok=True)
    args.out_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(args.out_csv, rows)
    text = render_text(summary)
    args.out_text.parent.mkdir(parents=True, exist_ok=True)
    args.out_text.write_text(text, encoding="utf-8")
    candidates = build_parameter_candidates(summary, coverage)
    candidates_json = args.out_candidates_json or args.out_json.with_name(args.out_json.stem + "-parameter-candidates.json")
    candidates_text = args.out_candidates_text or args.out_text.with_name(args.out_text.stem + "-parameter-candidates.txt")
    candidates_json.write_text(json.dumps(candidates, ensure_ascii=False, indent=2), encoding="utf-8")
    candidates_text.write_text(render_candidates(candidates), encoding="utf-8")
    if args.out_candidates_java:
        args.out_candidates_java.parent.mkdir(parents=True, exist_ok=True)
        args.out_candidates_java.write_text(render_candidates_java(candidates), encoding="utf-8")
    print(text)
    print()
    print(render_candidates(candidates))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
