# -*- coding: utf-8 -*-
"""
Compare the reference DOCX against the regenerated DOCX at the OOXML level.

The report focuses on layout facts that are hard to judge reliably from a
screenshot: formula VML box size, run baseline position, page geometry,
paragraph spacing/indentation, and formula context.
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import statistics
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_REFERENCE = ROOT / "rebuild-assets/external/fraction-split-reference.docx"
DEFAULT_GENERATED = ROOT / "target/reference-roundtrip/fraction-split-reference-regenerated.docx"
DEFAULT_REQUEST = ROOT / "target/reference-roundtrip/fraction-split-reference.request.json"
DEFAULT_OUT_JSON = ROOT / "target/reference-roundtrip/format-comparison.json"
DEFAULT_OUT_TXT = ROOT / "target/reference-roundtrip/format-comparison.txt"


@dataclass
class FormulaBox:
    index: int
    paragraph_index: int
    width_pt: float
    height_pt: float
    position_half_pt: int | None
    context: str


@dataclass
class ParagraphInfo:
    index: int
    text: str
    formula_count: int
    spacing_before: int | None
    spacing_after: int | None
    spacing_line: int | None
    spacing_rule: str | None
    indent_left: int | None
    indent_first_line: int | None
    alignment: str | None


@dataclass
class RequestFormula:
    index: int
    field: str
    latex: str
    latex_class: str
    latex_length: int
    context: str


def read_zip_text(docx: Path, name: str) -> str:
    with zipfile.ZipFile(docx) as zf:
        try:
            return zf.read(name).decode("utf-8")
        except KeyError:
            return ""


def zip_names(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as zf:
        return zf.namelist()


def attr(xml: str, name: str) -> str | None:
    match = re.search(rf'{re.escape(name)}="([^"]*)"', xml)
    return match.group(1) if match else None


def int_attr(xml: str, name: str) -> int | None:
    value = attr(xml, name)
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def float_style(style: str, name: str) -> float | None:
    match = re.search(rf"{re.escape(name)}:([0-9.]+)(pt|in|cm|mm)", style)
    if not match:
        return None
    value = float(match.group(1))
    unit = match.group(2)
    if unit == "in":
        return value * 72.0
    if unit == "cm":
        return value * 72.0 / 2.54
    if unit == "mm":
        return value * 72.0 / 25.4
    return value


def text_from_paragraph(p_xml: str) -> str:
    chunks = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", p_xml, flags=re.S)
    text = "".join(chunks)
    text = re.sub(r"<[^>]+>", "", text)
    text = (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
    )
    return re.sub(r"\s+", " ", text).strip()


def trim_context(text: str, limit: int = 90) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "..."


def strip_tags(text: str) -> str:
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    return text


def normalize_match_text(text: str) -> str:
    text = strip_tags(text)
    text = re.sub(r"\$.*?\$", " ", text, flags=re.S)
    text = re.sub(r"\\[a-zA-Z]+\*?(?:\{[^{}]*\})?", " ", text)
    text = re.sub(r"[\s,，.。:：;；、\"“”'‘’\[\]【】()（）<>《》]+", "", text)
    return text.lower()


def classify_latex(latex: str) -> str:
    value = latex.strip()
    if "\\begin{array}" in value or "\\begin{aligned}" in value or "\\begin{matrix}" in value:
        return "array"
    if "\\longdiv" in value or "\\enclose{longdiv}" in value:
        return "longdiv"
    if "\\frac" in value or "\\dfrac" in value or "\\cfrac" in value:
        return "fraction_long" if len(value) >= 80 or "\\cdots" in value else "fraction"
    if re.fullmatch(r"[a-zA-Z]+(?:\s*[<>=]\s*[a-zA-Z]+)?", value):
        return "symbol"
    if re.fullmatch(r"=?\s*\d+(?:\\frac\{[^{}]+\}\{[^{}]+\})?", value):
        return "answer_short"
    if any(token in value for token in ["=", "+", "-", "\\times", "×", "\\cdots"]):
        return "linear_equation"
    return "other"


def line_parts(text: str) -> list[str]:
    parts = re.split(r"(?i)<br\s*/?>|(?:(?:\r?\n)\s*){2,}", text)
    return [part.strip() for part in parts if part and part.strip()]


def field_label(field: str) -> str:
    return {
        "knowledgePoint": "【考点】",
        "difficulty": "【难度】",
        "analyze": "【解析】",
        "solution": "【解答】",
        "correct": "【答案】",
    }.get(field, "")


def extract_math_from_field(index_start: int, field: str, text: str) -> list[RequestFormula]:
    rows: list[RequestFormula] = []
    label = field_label(field)
    for line_index, line in enumerate(line_parts(text)):
        context = re.sub(r"\$.*?\$", " ", strip_tags(line), flags=re.S)
        context = trim_context(f"{label if line_index == 0 else ''}{context}", 140)
        for match in re.finditer(r"\$(.*?)\$", line, flags=re.S):
            latex = match.group(1).strip()
            rows.append(
                RequestFormula(
                    index=index_start + len(rows),
                    field=field,
                    latex=latex,
                    latex_class=classify_latex(latex),
                    latex_length=len(latex),
                    context=context,
                )
            )
    return rows


def extract_request_formulas(request_json: Path) -> list[RequestFormula]:
    if not request_json.exists():
        return []
    request = json.loads(request_json.read_text(encoding="utf-8"))
    rows: list[RequestFormula] = []
    sections = request.get("sections") or []
    fields = ["content", "knowledgePoint", "difficulty", "analyze", "solution", "correct"]
    for section in sections:
        for question in section.get("questions") or []:
            for field in fields:
                value = question.get(field)
                if isinstance(value, str) and value:
                    rows.extend(extract_math_from_field(len(rows), field, value))
    return rows


def paragraph_info(index: int, p_xml: str) -> ParagraphInfo:
    ppr_match = re.search(r"<w:pPr\b.*?</w:pPr>", p_xml, flags=re.S)
    ppr = ppr_match.group(0) if ppr_match else ""
    spacing_match = re.search(r"<w:spacing\b[^>]*/?>", ppr)
    indent_match = re.search(r"<w:ind\b[^>]*/?>", ppr)
    jc_match = re.search(r"<w:jc\b[^>]*/?>", ppr)
    spacing = spacing_match.group(0) if spacing_match else ""
    indent = indent_match.group(0) if indent_match else ""
    jc = jc_match.group(0) if jc_match else ""
    return ParagraphInfo(
        index=index,
        text=trim_context(text_from_paragraph(p_xml), 140),
        formula_count=len(re.findall(r"<w:object\b", p_xml)),
        spacing_before=int_attr(spacing, "w:before"),
        spacing_after=int_attr(spacing, "w:after"),
        spacing_line=int_attr(spacing, "w:line"),
        spacing_rule=attr(spacing, "w:lineRule"),
        indent_left=int_attr(indent, "w:left"),
        indent_first_line=int_attr(indent, "w:firstLine"),
        alignment=attr(jc, "w:val"),
    )


def extract_formulas(paragraphs: list[str]) -> list[FormulaBox]:
    rows: list[FormulaBox] = []
    index = 0
    for p_idx, p_xml in enumerate(paragraphs):
        context = trim_context(text_from_paragraph(p_xml))
        object_matches = list(re.finditer(r"<w:object\b.*?</w:object>", p_xml, flags=re.S))
        for obj in object_matches:
            before = p_xml[: obj.start()]
            pos_matches = list(re.finditer(r"<w:position\b[^>]*/?>", before))
            pos = int_attr(pos_matches[-1].group(0), "w:val") if pos_matches else None
            shape_match = re.search(r'<v:shape\b[^>]*style="([^"]+)"', obj.group(0), flags=re.S)
            if not shape_match:
                continue
            style = shape_match.group(1)
            width = float_style(style, "width")
            height = float_style(style, "height")
            if width is None or height is None:
                continue
            rows.append(
                FormulaBox(
                    index=index,
                    paragraph_index=p_idx,
                    width_pt=round(width, 2),
                    height_pt=round(height, 2),
                    position_half_pt=pos,
                    context=context,
                )
            )
            index += 1
    return rows


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


def media_summary(names: list[str]) -> dict[str, int]:
    result: dict[str, int] = {}
    for name in names:
        if not name.startswith("word/media/"):
            continue
        ext = Path(name).suffix.lower().lstrip(".") or "<none>"
        result[ext] = result.get(ext, 0) + 1
    return dict(sorted(result.items()))


def page_settings(document_xml: str) -> dict[str, str | None]:
    sect_match = re.search(r"<w:sectPr\b.*?</w:sectPr>", document_xml, flags=re.S)
    sect = sect_match.group(0) if sect_match else ""
    size_match = re.search(r"<w:pgSz\b[^>]*/?>", sect)
    margin_match = re.search(r"<w:pgMar\b[^>]*/?>", sect)
    size = size_match.group(0) if size_match else ""
    margin = margin_match.group(0) if margin_match else ""
    return {
        "w": attr(size, "w:w"),
        "h": attr(size, "w:h"),
        "top": attr(margin, "w:top"),
        "bottom": attr(margin, "w:bottom"),
        "left": attr(margin, "w:left"),
        "right": attr(margin, "w:right"),
        "header": attr(margin, "w:header"),
        "footer": attr(margin, "w:footer"),
    }


def analyze(docx: Path) -> dict[str, Any]:
    names = zip_names(docx)
    document_xml = read_zip_text(docx, "word/document.xml")
    paragraphs_xml = re.findall(r"<w:p\b.*?</w:p>", document_xml, flags=re.S)
    paragraphs = [paragraph_info(i, xml) for i, xml in enumerate(paragraphs_xml)]
    formulas = extract_formulas(paragraphs_xml)
    widths = [item.width_pt for item in formulas]
    heights = [item.height_pt for item in formulas]
    positions = [float(item.position_half_pt) for item in formulas if item.position_half_pt is not None]
    return {
        "path": str(docx),
        "bytes": docx.stat().st_size,
        "page": page_settings(document_xml),
        "paragraph_count": len(paragraphs),
        "ole_count": len(re.findall(r"<o:OLEObject\b", document_xml)),
        "object_count": len(re.findall(r"<w:object\b", document_xml)),
        "embedding_count": len([n for n in names if n.startswith("word/embeddings/")]),
        "media_count": len([n for n in names if n.startswith("word/media/")]),
        "media_ext": media_summary(names),
        "formula_width": stats(widths),
        "formula_height": stats(heights),
        "formula_position": stats(positions),
        "formula_height_buckets": height_buckets(heights),
        "formulas": [asdict(item) for item in formulas],
        "paragraphs": [asdict(item) for item in paragraphs],
    }


def paired_deltas(reference: list[dict[str, Any]], generated: list[dict[str, Any]]) -> dict[str, Any]:
    count = min(len(reference), len(generated))
    rows: list[dict[str, Any]] = []
    for i in range(count):
        ref = reference[i]
        gen = generated[i]
        height_ratio = gen["height_pt"] / ref["height_pt"] if ref["height_pt"] else None
        width_ratio = gen["width_pt"] / ref["width_pt"] if ref["width_pt"] else None
        rows.append(
            {
                "index": i,
                "reference_height_pt": ref["height_pt"],
                "generated_height_pt": gen["height_pt"],
                "height_ratio": round(height_ratio, 3) if height_ratio is not None else None,
                "reference_width_pt": ref["width_pt"],
                "generated_width_pt": gen["width_pt"],
                "width_ratio": round(width_ratio, 3) if width_ratio is not None else None,
                "height_abs_delta_pt": round(abs(gen["height_pt"] - ref["height_pt"]), 2),
                "width_abs_delta_pt": round(abs(gen["width_pt"] - ref["width_pt"]), 2),
                "baseline_abs_delta_half_pt": (
                    abs(int(gen["position_half_pt"]) - int(ref["position_half_pt"]))
                    if gen["position_half_pt"] is not None and ref["position_half_pt"] is not None
                    else None
                ),
                "reference_context": ref["context"],
                "generated_context": gen["context"],
            }
        )
    height_ratios = [row["height_ratio"] for row in rows if row["height_ratio"] is not None]
    width_ratios = [row["width_ratio"] for row in rows if row["width_ratio"] is not None]
    height_abs_delta = [row["height_abs_delta_pt"] for row in rows]
    width_abs_delta = [row["width_abs_delta_pt"] for row in rows]
    baseline_abs_delta = [
        row["baseline_abs_delta_half_pt"]
        for row in rows
        if row["baseline_abs_delta_half_pt"] is not None
    ]
    return {
        "paired_count": count,
        "unpaired_reference": max(len(reference) - count, 0),
        "unpaired_generated": max(len(generated) - count, 0),
        "height_ratio": stats([float(v) for v in height_ratios]),
        "width_ratio": stats([float(v) for v in width_ratios]),
        "height_abs_delta_pt": stats([float(v) for v in height_abs_delta]),
        "width_abs_delta_pt": stats([float(v) for v in width_abs_delta]),
        "baseline_abs_delta_half_pt": stats([float(v) for v in baseline_abs_delta]),
        "worst_height": sorted(rows, key=lambda row: row["height_ratio"] if row["height_ratio"] is not None else 999)[:20],
        "worst_width_over": sorted(rows, key=lambda row: row["width_ratio"] if row["width_ratio"] is not None else -1, reverse=True)[:20],
    }


def context_similarity(a: str, b: str) -> float:
    left = normalize_match_text(a)
    right = normalize_match_text(b)
    if not left and not right:
        return 1.0
    if not left or not right:
        return 0.0
    return difflib.SequenceMatcher(None, left, right).ratio()


def align_formulas(reference: list[dict[str, Any]], request_formulas: list[RequestFormula]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    ref_start = 0
    for req in request_formulas:
        best_index: int | None = None
        best_score = -1.0
        window_end = min(len(reference), ref_start + 12)
        for ref_index in range(ref_start, window_end):
            ref = reference[ref_index]
            score = context_similarity(ref.get("context", ""), req.context)
            if req.latex_class in {"fraction", "fraction_long", "array"} and ref["height_pt"] >= 24:
                score += 0.05
            if req.latex_class in {"symbol"} and ref["height_pt"] < 20:
                score += 0.05
            if score > best_score:
                best_score = score
                best_index = ref_index
        if best_index is None:
            continue
        ref_start = best_index + 1
        rows.append(
            {
                "request_index": req.index,
                "reference_index": best_index,
                "alignment_score": round(best_score, 3),
                "field": req.field,
                "latex": req.latex,
                "latex_class": req.latex_class,
                "latex_length": req.latex_length,
                "request_context": req.context,
                "reference_context": reference[best_index].get("context", ""),
                "reference_width_pt": reference[best_index]["width_pt"],
                "reference_height_pt": reference[best_index]["height_pt"],
                "reference_position_half_pt": reference[best_index]["position_half_pt"],
            }
        )
    return rows


def formula_deltas_with_latex(
    reference: list[dict[str, Any]],
    generated: list[dict[str, Any]],
    request_formulas: list[RequestFormula],
) -> dict[str, Any]:
    count = min(len(generated), len(request_formulas))
    aligned_reference = align_formulas(reference, request_formulas[:count])
    aligned_by_request = {row["request_index"]: row for row in aligned_reference}
    rows: list[dict[str, Any]] = []
    for index in range(count):
        gen = generated[index]
        req = request_formulas[index]
        ref = aligned_by_request.get(index)
        if ref is None:
            continue
        width_delta = gen["width_pt"] - ref["reference_width_pt"]
        height_delta = gen["height_pt"] - ref["reference_height_pt"]
        width_abs = abs(width_delta)
        height_abs = abs(height_delta)
        rows.append(
            {
                **ref,
                "generated_index": index,
                "generated_width_pt": gen["width_pt"],
                "generated_height_pt": gen["height_pt"],
                "generated_position_half_pt": gen["position_half_pt"],
                "width_delta_pt": round(width_delta, 2),
                "height_delta_pt": round(height_delta, 2),
                "width_abs_delta_pt": round(width_abs, 2),
                "height_abs_delta_pt": round(height_abs, 2),
                "size_abs_delta_pt": round(width_abs + height_abs, 2),
                "generated_context": gen.get("context", ""),
            }
        )
    width_abs_values = [row["width_abs_delta_pt"] for row in rows]
    height_abs_values = [row["height_abs_delta_pt"] for row in rows]
    size_abs_values = [row["size_abs_delta_pt"] for row in rows]
    classes: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        classes.setdefault(row["latex_class"], []).append(row)
    return {
        "paired_count": len(rows),
        "unpaired_request": max(len(request_formulas) - len(rows), 0),
        "unpaired_reference": max(len(reference) - len({row["reference_index"] for row in rows}), 0),
        "width_abs_delta": stats([float(v) for v in width_abs_values]),
        "height_abs_delta": stats([float(v) for v in height_abs_values]),
        "size_abs_delta": stats([float(v) for v in size_abs_values]),
        "by_class": {
            name: {
                "count": len(values),
                "width_abs_delta": stats([float(row["width_abs_delta_pt"]) for row in values]),
                "height_abs_delta": stats([float(row["height_abs_delta_pt"]) for row in values]),
            }
            for name, values in sorted(classes.items())
        },
        "worst_size": sorted(rows, key=lambda row: row["size_abs_delta_pt"], reverse=True)[:40],
        "worst_width": sorted(rows, key=lambda row: row["width_abs_delta_pt"], reverse=True)[:40],
        "worst_height": sorted(rows, key=lambda row: row["height_abs_delta_pt"], reverse=True)[:40],
        "rows": rows,
    }


def paragraph_deltas(reference: list[dict[str, Any]], generated: list[dict[str, Any]]) -> dict[str, Any]:
    count = min(len(reference), len(generated))
    rows = []
    for i in range(count):
        ref = reference[i]
        gen = generated[i]
        changed = {
            key: {"reference": ref.get(key), "generated": gen.get(key)}
            for key in [
                "spacing_before",
                "spacing_after",
                "spacing_line",
                "spacing_rule",
                "indent_left",
                "indent_first_line",
                "alignment",
            ]
            if ref.get(key) != gen.get(key)
        }
        if changed:
            rows.append(
                {
                    "index": i,
                    "reference_text": ref.get("text"),
                    "generated_text": gen.get("text"),
                    "diff": changed,
                }
            )
    return {
        "paired_count": count,
        "unpaired_reference": max(len(reference) - count, 0),
        "unpaired_generated": max(len(generated) - count, 0),
        "first_differences": rows[:40],
    }


def build_report(reference: Path, generated: Path, request_json: Path) -> dict[str, Any]:
    ref = analyze(reference)
    gen = analyze(generated)
    request_formulas = extract_request_formulas(request_json)
    return {
        "reference": ref,
        "generated": gen,
        "request_formula_count": len(request_formulas),
        "request_formulas": [asdict(item) for item in request_formulas],
        "formula_deltas_by_order": paired_deltas(ref["formulas"], gen["formulas"]),
        "formula_deltas_by_latex": formula_deltas_with_latex(ref["formulas"], gen["formulas"], request_formulas),
        "paragraph_deltas_by_order": paragraph_deltas(ref["paragraphs"], gen["paragraphs"]),
    }


def render_summary(report: dict[str, Any]) -> str:
    ref = report["reference"]
    gen = report["generated"]
    deltas = report["formula_deltas_by_order"]
    latex_deltas = report["formula_deltas_by_latex"]
    lines = [
        "Reference format comparison",
        "",
        f"reference: {ref['path']}",
        f"generated: {gen['path']}",
        "",
        "Document counts",
        f"  paragraphs: reference={ref['paragraph_count']} generated={gen['paragraph_count']}",
        f"  OLE objects: reference={ref['ole_count']} generated={gen['ole_count']}",
        f"  media: reference={ref['media_count']} {ref['media_ext']} generated={gen['media_count']} {gen['media_ext']}",
        "",
        "Formula size summary",
        f"  width pt: reference={ref['formula_width']} generated={gen['formula_width']}",
        f"  height pt: reference={ref['formula_height']} generated={gen['formula_height']}",
        f"  baseline half-pt: reference={ref['formula_position']} generated={gen['formula_position']}",
        f"  height buckets: reference={ref['formula_height_buckets']} generated={gen['formula_height_buckets']}",
        "",
        "Paired formula deltas",
        f"  paired={deltas['paired_count']} unpaired_reference={deltas['unpaired_reference']} unpaired_generated={deltas['unpaired_generated']}",
        f"  height_ratio={deltas['height_ratio']}",
        f"  width_ratio={deltas['width_ratio']}",
        f"  height_abs_delta_pt={deltas['height_abs_delta_pt']}",
        f"  width_abs_delta_pt={deltas['width_abs_delta_pt']}",
        f"  baseline_abs_delta_half_pt={deltas['baseline_abs_delta_half_pt']}",
        "",
        "LaTeX-aware formula deltas",
        f"  request_formula_count={report['request_formula_count']} paired={latex_deltas['paired_count']} unpaired_reference={latex_deltas['unpaired_reference']} unpaired_request={latex_deltas['unpaired_request']}",
        f"  width_abs_delta_pt={latex_deltas['width_abs_delta']}",
        f"  height_abs_delta_pt={latex_deltas['height_abs_delta']}",
        f"  size_abs_delta_pt={latex_deltas['size_abs_delta']}",
        f"  by_class={latex_deltas['by_class']}",
        "",
        "Worst height ratios",
    ]
    for row in deltas["worst_height"][:10]:
        lines.append(
            "  #{index}: ref={reference_height_pt}pt gen={generated_height_pt}pt ratio={height_ratio} text={generated_context}".format(
                **row
            )
        )
    lines.extend(["", "Worst width ratios"])
    for row in deltas["worst_width_over"][:10]:
        lines.append(
            "  #{index}: ref={reference_width_pt}pt gen={generated_width_pt}pt ratio={width_ratio} text={generated_context}".format(
                **row
            )
        )
    lines.extend(["", "Worst LaTeX-aware size deltas"])
    for row in latex_deltas["worst_size"][:12]:
        latex_preview = row["latex"].replace("\n", " ")
        if len(latex_preview) > 90:
            latex_preview = latex_preview[:89] + "..."
        lines.append(
            "  req#{request_index}/ref#{reference_index}/gen#{generated_index} {latex_class}: ref={reference_width_pt}x{reference_height_pt}pt gen={generated_width_pt}x{generated_height_pt}pt delta={size_abs_delta_pt} score={alignment_score} latex={latex_preview}".format(
                **row,
                latex_preview=latex_preview,
            )
        )
    lines.extend(["", "First paragraph format differences"])
    for row in report["paragraph_deltas_by_order"]["first_differences"][:12]:
        lines.append(
            f"  p{row['index']}: ref={row['reference_text']} gen={row['generated_text']} diff={row['diff']}"
        )
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--generated", type=Path, default=DEFAULT_GENERATED)
    parser.add_argument("--request-json", type=Path, default=DEFAULT_REQUEST)
    parser.add_argument("--out-json", type=Path, default=DEFAULT_OUT_JSON)
    parser.add_argument("--out-text", type=Path, default=DEFAULT_OUT_TXT)
    args = parser.parse_args()

    report = build_report(args.reference.resolve(), args.generated.resolve(), args.request_json.resolve())
    args.out_json.parent.mkdir(parents=True, exist_ok=True)
    args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    args.out_text.write_text(render_summary(report), encoding="utf-8")
    print(render_summary(report))
    print(f"wrote {args.out_json}")
    print(f"wrote {args.out_text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
