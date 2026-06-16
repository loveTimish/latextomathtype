# -*- coding: utf-8 -*-
"""Compare source and generated DOCX formula object metrics.

This is the first hard gate for the MathType-like target mode:
- WMF preview physical size
- VML shape size
- w:dxaOrig/w:dyaOrig
- w:position baseline
- OLE object and preview media counts
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import statistics as st
import struct
import sys
import zipfile
from collections import defaultdict, deque
from pathlib import Path


PLACEABLE_WMF_KEY = 0x9AC6CDD7
EMPTY_FRAC_DEN_RE = re.compile(
    r"\\frac\s*\{\s*([^{}]+?)\s*\}\s*\{\s*\}\s*"
    r"((?:[A-Za-z0-9]+)(?:\s*\\times\s*[A-Za-z0-9]+)*)"
)


def parse_wmf_phys(data: bytes) -> tuple[float, float] | None:
    if len(data) < 22 or struct.unpack("<I", data[:4])[0] != PLACEABLE_WMF_KEY:
        return None
    left, top, right, bottom, inch = struct.unpack("<hhhhH", data[6:16])
    if inch <= 0:
        return None
    return ((right - left) / inch * 72.0, (bottom - top) / inch * 72.0)


def rel_map(zip_file: zipfile.ZipFile, rels_name: str) -> dict[str, str]:
    if rels_name not in zip_file.namelist():
        return {}
    rels = zip_file.read(rels_name).decode("utf-8", "replace")
    return dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))


def read_target(zip_file: zipfile.ZipFile, target: str) -> bytes | None:
    name = target.lstrip("/")
    if not name.startswith("word/"):
        name = "word/" + name
    if name in zip_file.namelist():
        return zip_file.read(name)
    return None


def extract_formula_objects(docx_path: Path) -> list[dict]:
    out = []
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()
        doc = z.read("word/document.xml").decode("utf-8", "replace")
        rels = rel_map(z, "word/_rels/document.xml.rels")
        for idx, match in enumerate(re.finditer(r"<w:object\b([^>]*)>(.*?)</w:object>", doc, re.S), 1):
            attrs, body = match.group(1), match.group(2)
            dxa = re.search(r'w:dxaOrig="(\d+)"', attrs)
            dya = re.search(r'w:dyaOrig="(\d+)"', attrs)
            dxa_pt = int(dxa.group(1)) / 20.0 if dxa else None
            dya_pt = int(dya.group(1)) / 20.0 if dya else None
            style = re.search(r"<v:shape\b[^>]*\sstyle=\"([^\"]*)\"", body)
            shape_w = shape_h = None
            if style:
                width = re.search(r"width:([\d.]+)pt", style.group(1))
                height = re.search(r"height:([\d.]+)pt", style.group(1))
                shape_w = float(width.group(1)) if width else None
                shape_h = float(height.group(1)) if height else None
            image_rid = re.search(r'<v:imagedata [^>]*r:id="([^"]+)"', body)
            image_target = rels.get(image_rid.group(1), "") if image_rid else ""
            image_ext = image_target.rsplit(".", 1)[-1].lower() if "." in image_target else ""
            wmf_w = wmf_h = None
            if image_ext == "wmf" and image_target:
                data = read_target(z, image_target)
                phys = parse_wmf_phys(data or b"")
                if phys:
                    wmf_w, wmf_h = phys
            ole_rid = re.search(r'<o:OLEObject [^>]*r:id="([^"]+)"', body)
            ole_target = rels.get(ole_rid.group(1), "") if ole_rid else ""
            ole_size = len(read_target(z, ole_target) or b"") if ole_target else None
            prefix = doc[max(0, match.start() - 500) : match.start()]
            pos = re.search(r'<w:position w:val="(-?\d+)"', prefix)
            out.append(
                {
                    "index": idx,
                    "shape_w_pt": shape_w,
                    "shape_h_pt": shape_h,
                    "dxa_pt": dxa_pt,
                    "dya_pt": dya_pt,
                    "position_halfpt": int(pos.group(1)) if pos else None,
                    "image_ext": image_ext,
                    "wmf_w_pt": wmf_w,
                    "wmf_h_pt": wmf_h,
                    "display_w_pt": dxa_pt,
                    "display_h_pt": dya_pt,
                    "ole_size": ole_size,
                }
            )
    return out


class ObjectList(list):
    pass


def extract_formula_objects_with_meta(docx_path: Path) -> ObjectList:
    rows = ObjectList(extract_formula_objects(docx_path))
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()
        media = [n for n in names if n.startswith("word/media/")]
        embeddings = [n for n in names if n.startswith("word/embeddings/")]
        rows.meta = {
            "objects": len(rows),
            "media_wmf": len([n for n in media if n.lower().endswith(".wmf")]),
            "media_png": len([n for n in media if n.lower().endswith(".png")]),
            "media_emf": len([n for n in media if n.lower().endswith(".emf")]),
            "embeddings_bin": len([n for n in embeddings if n.lower().endswith(".bin")]),
        }
    return rows


def pct(value: float | None) -> float | None:
    return None if value is None else value * 100.0


def ratio(got: float | None, ref: float | None) -> float | None:
    if got is None or ref in (None, 0):
        return None
    return got / ref


def normalize_latex_key(value: str) -> str:
    value = value or ""
    value = re.sub(r"^\\pwmetrics\{[^}]+}\s*", "", value)
    value = re.sub(r"^\\pwstyle\{[^}]*}\s*", "", value)
    value = EMPTY_FRAC_DEN_RE.sub(lambda m: rf"\frac {{ {m.group(1).strip()} }} {{ {m.group(2).strip()} }}", value)
    value = re.sub(r"\\(?:hline|hdashline)(?:\[[^\]]*])?", "", value)
    value = value.replace(r"\lt ", "<").replace(r"\gt ", ">")
    value = re.sub(r"\s+", " ", value.strip())
    return value


def find_group_end(text: str, start: int) -> int:
    if start >= len(text) or text[start] != "{":
        return -1
    depth = 0
    index = start
    while index < len(text):
        ch = text[index]
        if ch == "\\":
            index += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return index
        index += 1
    return -1


def skip_spaces(text: str, index: int) -> int:
    while index < len(text) and text[index].isspace():
        index += 1
    return index


STYLE_WRAPPER_COMMANDS = {
    "textit",
    "textbf",
    "textnormal",
    "textrm",
    "mathrm",
    "mathit",
    "mathbf",
    "emph",
    "textsc",
}


def strip_style_wrappers(value: str) -> str:
    out: list[str] = []
    index = 0
    while index < len(value):
        if value[index] != "\\":
            out.append(value[index])
            index += 1
            continue
        name_end = index + 1
        while name_end < len(value) and value[name_end].isalpha():
            name_end += 1
        command = value[index + 1:name_end]
        cursor = skip_spaces(value, name_end)
        if command == "textcolor":
            color_end = find_group_end(value, cursor)
            body_start = skip_spaces(value, color_end + 1) if color_end >= 0 else -1
            body_end = find_group_end(value, body_start) if body_start >= 0 else -1
            if body_end >= 0:
                out.append(strip_style_wrappers(value[body_start + 1:body_end]))
                index = body_end + 1
                continue
        elif command in STYLE_WRAPPER_COMMANDS:
            body_end = find_group_end(value, cursor)
            if body_end >= 0:
                out.append(strip_style_wrappers(value[cursor + 1:body_end]))
                index = body_end + 1
                continue
        out.append(value[index])
        index += 1
    return "".join(out)


def ordinal_latex_key(value: str) -> str:
    value = normalize_latex_key(value)
    value = value.replace("&", "")
    value = value.replace("×", r"\times")
    value = re.sub(r"\$\s*([^$]+?)\s*\$", r"\1", value)
    value = strip_style_wrappers(value)
    value = re.sub(r"\\overset\s*\{\s*\\frown\s*\}\s*\{\s*([^{}]+?)\s*\}", r"\\overarc{\1}", value)
    value = re.sub(r"\\(?:left|right)\s*\.", "", value)
    value = re.sub(r"\\(?:left|right)\s*", "", value)
    value = re.sub(r"\\(?:rm|mathrm)\s*", "", value)
    value = re.sub(r"\\(?![A-Za-z])", "", value)
    return re.sub(r"[\s{}\[\]()]+", "", value)


def parse_request_math(request_path: Path) -> list[dict]:
    request = json.loads(request_path.read_text(encoding="utf-8"))
    compact = bool((request.get("paper") or {}).get("compactLayout"))
    parts = []
    for section in request.get("sections", []):
        for question in section.get("questions", []):
            parts.append(str(question.get("content") or ""))
            parts.append(str(question.get("knowledgePoint") or ""))
            if compact:
                parts.append(str(question.get("difficulty") or ""))
                parts.extend(str(tag) for tag in (question.get("tags") or []))
            else:
                parts.extend(str(tag) for tag in (question.get("tags") or []))
                parts.append(str(question.get("difficulty") or ""))
            parts.append(str(question.get("analyze") or ""))
            parts.append(str(question.get("solution") or ""))
            parts.append(str(question.get("correct") or ""))
    text = "\n".join(parts)
    out = []
    for a, b in re.findall(r"\$\$(.+?)\$\$|\$(.+?)\$", text, re.S):
        body = (a or b).strip()
        raw_body = body
        metrics = None
        match = re.match(
            r"^\\pwmetrics\{([0-9]+(?:\.[0-9]+)?)\s*,\s*([0-9]+(?:\.[0-9]+)?)(?:\s*,\s*([0-9]+(?:\.[0-9]+)?)\s*,\s*([0-9]+(?:\.[0-9]+)?))?\}\s*",
            body,
        )
        if match:
            metrics = {
                "wmfWidthPt": float(match.group(1)),
                "wmfHeightPt": float(match.group(2)),
                "shapeWidthPt": float(match.group(3)) if match.group(3) else float(match.group(1)),
                "shapeHeightPt": float(match.group(4)) if match.group(4) else float(match.group(2)),
            }
            body = body[match.end():].strip()
        body = re.sub(r"^\\pwstyle\{[^}]*}\s*", "", body).strip()
        out.append({"latex": normalize_latex_key(body), "rawLatex": raw_body, "metrics": metrics})
    return out


def request_latex_sequence(request_path: Path) -> list[str]:
    return [item["latex"] for item in parse_request_math(request_path)]


def report_latex_sequence(report_path: Path) -> list[str]:
    report = json.loads(report_path.read_text(encoding="utf-8"))
    return [
        normalize_latex_key(item.get("output", ""))
        for item in report.get("equations", [])
        if item.get("status") == "converted" and item.get("output")
    ]


def classify_latex(latex: str) -> str:
    latex = latex or ""
    if r"\begin{array}" in latex:
        return "array"
    if r"\frac" in latex:
        return "fraction"
    if r"\sqrt" in latex:
        return "sqrt"
    if re.search(r"(?<!\\)[_^]", latex):
        return "script"
    if r"\overline" in latex or r"\underline" in latex or r"\overset" in latex or r"\underset" in latex:
        return "accent"
    return "linear"


def latex_alignment(source_latex: list[str], generated_latex: list[str]) -> list[tuple[int, int]]:
    """Return object index pairs using LaTeX text as a stable key.

    Duplicates are paired in encounter order.  This avoids false size deltas when
    full-document text parsing loses a formula and the request builder appends it
    later to preserve object count.
    """
    generated_by_key: dict[str, deque[int]] = defaultdict(deque)
    for index, latex in enumerate(generated_latex):
        generated_by_key[latex].append(index)
    pairs = []
    for source_index, latex in enumerate(source_latex):
        queue = generated_by_key.get(latex)
        if queue:
            pairs.append((source_index, queue.popleft()))
    return pairs


def summarize(values: list[float]) -> dict:
    if not values:
        return {"n": 0}
    ordered = sorted(values)
    return {
        "n": len(values),
        "median": st.median(ordered),
        "p10": ordered[len(ordered) // 10],
        "p90": ordered[min(len(ordered) - 1, 9 * len(ordered) // 10)],
        "max_abs_error_pct": max(abs(v - 1.0) * 100.0 for v in ordered),
        "within_1pct": sum(abs(v - 1.0) <= 0.01 for v in ordered),
    }


def summarize_by_class(rows: list[dict], ratio_field: str) -> dict:
    groups: dict[str, list[float]] = defaultdict(list)
    for row in rows:
        value = row.get(ratio_field)
        if value is not None:
            groups[row.get("class") or "unknown"].append(value)
    return {name: summarize(values) for name, values in sorted(groups.items())}


def plausible_for_ratio(row: dict) -> bool:
    source_width = row.get("source_wmf_w_pt")
    generated_width = row.get("generated_wmf_w_pt")
    source_height = row.get("source_wmf_h_pt")
    generated_height = row.get("generated_wmf_h_pt")
    if source_width is None or generated_width is None or source_height is None or generated_height is None:
        return False
    # Extremely narrow source previews are usually ordinal/key mismatches or legacy
    # placeholder objects, not useful physical-size calibration samples.
    if source_width < 12.0 and generated_width > 80.0:
        return False
    if source_width < 6.0 or source_height < 6.0:
        return False
    return True


def plausible_rows(rows: list[dict]) -> list[dict]:
    return [row for row in rows if plausible_for_ratio(row)]


def worst_rows(rows: list[dict], ratio_field: str, limit: int = 20) -> list[dict]:
    ranked = [
        row for row in rows
        if row.get(ratio_field) is not None
    ]
    ranked.sort(key=lambda row: abs(row[ratio_field] - 1.0), reverse=True)
    out = []
    for row in ranked[:limit]:
        out.append(
            {
                "index": row["index"],
                "generated_index": row["generated_index"],
                "class": row.get("class"),
                "ratio": row.get(ratio_field),
                "error_pct": row.get(ratio_field.replace("_ratio", "_error_pct")),
                "source_wmf_w_pt": row.get("source_wmf_w_pt"),
                "generated_wmf_w_pt": row.get("generated_wmf_w_pt"),
                "latex": row.get("latex"),
            }
        )
    return out


def report_source_indexes(source_report: Path | None) -> list[int]:
    if not source_report or not source_report.exists():
        return []
    data = json.loads(source_report.read_text(encoding="utf-8-sig"))
    indexes = []
    for item in data.get("equations") or []:
        doc_index = item.get("docObjectIndex")
        if doc_index:
            indexes.append(int(doc_index))
    return indexes


def metric_rows(
    src,
    gen,
    pairs: list[tuple[int, int]],
    source_latex_by_index: dict[int, str],
    request_math: list[dict],
) -> list[dict]:
    rows = []
    for s_index, g_index in pairs:
        s, g = src[s_index], gen[g_index]
        latex = source_latex_by_index.get(int(s.get("index") or 0), "")
        row = {
            "index": s_index + 1,
            "generated_index": g_index + 1,
            "class": classify_latex(latex),
            "latex": latex,
        }
        for field in ["shape_w_pt", "shape_h_pt", "wmf_w_pt", "wmf_h_pt", "dxa_pt", "dya_pt", "display_w_pt", "display_h_pt"]:
            row[f"source_{field}"] = s.get(field)
            row[f"generated_{field}"] = g.get(field)
            row[f"{field}_ratio"] = ratio(g.get(field), s.get(field))
            row[f"{field}_error_pct"] = pct((row[f"{field}_ratio"] or 0) - 1.0) if row[f"{field}_ratio"] else None
        row["source_position_halfpt"] = s.get("position_halfpt")
        row["generated_position_halfpt"] = g.get("position_halfpt")
        row["position_diff_halfpt"] = (
            g.get("position_halfpt") - s.get("position_halfpt")
            if g.get("position_halfpt") is not None and s.get("position_halfpt") is not None
            else None
        )
        row["source_image_ext"] = s.get("image_ext")
        row["generated_image_ext"] = g.get("image_ext")
        row["source_ole_size"] = s.get("ole_size")
        row["generated_ole_size"] = g.get("ole_size")
        if g_index < len(request_math) and request_math[g_index].get("metrics"):
            target_metrics = request_math[g_index]["metrics"]
            target_wmf_w = target_metrics["wmfWidthPt"]
            target_wmf_h = target_metrics["wmfHeightPt"]
            target_shape_w = target_metrics["shapeWidthPt"]
            target_shape_h = target_metrics["shapeHeightPt"]
            row["target_wmf_w_pt"] = target_wmf_w
            row["target_wmf_h_pt"] = target_wmf_h
            row["target_wmf_w_pt_ratio"] = ratio(g.get("wmf_w_pt"), target_wmf_w)
            row["target_wmf_h_pt_ratio"] = ratio(g.get("wmf_h_pt"), target_wmf_h)
            row["target_shape_w_pt_ratio"] = ratio(g.get("shape_w_pt"), target_shape_w)
            row["target_shape_h_pt_ratio"] = ratio(g.get("shape_h_pt"), target_shape_h)
        else:
            row["target_wmf_w_pt"] = None
            row["target_wmf_h_pt"] = None
            row["target_wmf_w_pt_ratio"] = None
            row["target_wmf_h_pt_ratio"] = None
            row["target_shape_w_pt_ratio"] = None
            row["target_shape_h_pt_ratio"] = None
        rows.append(row)
    return rows


def compare_pair(source: Path, generated: Path, out_dir: Path,
                 source_report: Path | None = None, generated_request: Path | None = None) -> dict:
    src = extract_formula_objects_with_meta(source)
    gen = extract_formula_objects_with_meta(generated)
    source_latex = report_latex_sequence(source_report) if source_report else []
    source_latex_by_index: dict[int, str] = {}
    source_indexes = report_source_indexes(source_report)
    if source_indexes:
        report_data = json.loads(source_report.read_text(encoding="utf-8"))
        source_latex_by_index = {
            int(item.get("docObjectIndex")): normalize_latex_key(item.get("output", ""))
            for item in report_data.get("equations", [])
            if item.get("status") == "converted" and item.get("output") and item.get("docObjectIndex")
        }
        source_latex = [source_latex_by_index.get(index, "") for index in source_indexes]
        allowed_indexes = set(source_indexes)
        src = ObjectList([item for item in src if item.get("index") in allowed_indexes])
        src.meta = dict(getattr(src, "meta", {}))
        src.meta["objects"] = len(src)
    request_math = parse_request_math(generated_request) if generated_request else []
    generated_latex = [item["latex"] for item in request_math]
    ordinal_mismatches = []
    if source_report and generated_request:
        ordinal_count = min(len(source_latex), len(generated_latex))
        for index in range(ordinal_count):
            if ordinal_latex_key(source_latex[index]) != ordinal_latex_key(generated_latex[index]):
                ordinal_mismatches.append(
                    {
                        "index": index + 1,
                        "source": source_latex[index],
                        "generated": generated_latex[index],
                    }
                )
        for index in range(ordinal_count, max(len(source_latex), len(generated_latex))):
            ordinal_mismatches.append(
                {
                    "index": index + 1,
                    "source": source_latex[index] if index < len(source_latex) else None,
                    "generated": generated_latex[index] if index < len(generated_latex) else None,
                }
            )
    if source_report and generated_request:
        index_pairs = [
            (s, g)
            for s, g in latex_alignment(source_latex, generated_latex)
            if s < len(src) and g < len(gen)
        ]
    else:
        index_pairs = [(i, i) for i in range(min(len(src), len(gen)))]
    paired_source_indexes = {s for s, _ in index_pairs}
    paired_generated_indexes = {g for _, g in index_pairs}
    rows = metric_rows(src, gen, index_pairs, source_latex_by_index, request_math)
    ordinal_rows = metric_rows(
        src, gen, [(i, i) for i in range(min(len(src), len(gen)))], source_latex_by_index, request_math
    )
    target_rows = []
    if request_math:
        for g_index, g in enumerate(gen):
            if g_index >= len(request_math) or not request_math[g_index].get("metrics"):
                continue
            target_metrics = request_math[g_index]["metrics"]
            target_rows.append(
                {
                    "generated_index": g_index + 1,
                    "target_wmf_w_pt": target_metrics["wmfWidthPt"],
                    "target_wmf_h_pt": target_metrics["wmfHeightPt"],
                    "target_shape_w_pt": target_metrics["shapeWidthPt"],
                    "target_shape_h_pt": target_metrics["shapeHeightPt"],
                    "target_wmf_w_pt_ratio": ratio(g.get("wmf_w_pt"), target_metrics["wmfWidthPt"]),
                    "target_wmf_h_pt_ratio": ratio(g.get("wmf_h_pt"), target_metrics["wmfHeightPt"]),
                    "target_shape_w_pt_ratio": ratio(g.get("shape_w_pt"), target_metrics["shapeWidthPt"]),
                    "target_shape_h_pt_ratio": ratio(g.get("shape_h_pt"), target_metrics["shapeHeightPt"]),
                }
            )
    out_dir.mkdir(parents=True, exist_ok=True)
    detail = out_dir / f"{source.stem}_vs_{generated.stem}.csv"
    with detail.open("w", encoding="utf-8-sig", newline="") as fp:
        writer = csv.DictWriter(fp, fieldnames=list(rows[0].keys()) if rows else ["index"])
        writer.writeheader()
        writer.writerows(rows)
    summary = {
        "source": str(source),
        "generated": str(generated),
        "source_meta": src.meta,
        "generated_meta": gen.meta,
        "paired_objects": len(index_pairs),
        "unpaired_source_objects": len(src) - len(paired_source_indexes),
        "unpaired_generated_objects": len(gen) - len(paired_generated_indexes),
        "missing_generated_objects": max(0, len(src) - len(gen)),
        "extra_generated_objects": max(0, len(gen) - len(src)),
        "ordinal_latex_mismatches": len(ordinal_mismatches),
        "ordinal_latex_mismatch_samples": ordinal_mismatches[:20],
        "alignment_mode": "latex" if source_report and generated_request else "ordinal",
        "ordinal_paired_objects": len(ordinal_rows),
        "ordinal_wmf_width_ratio": summarize([r["wmf_w_pt_ratio"] for r in ordinal_rows if r["wmf_w_pt_ratio"] is not None]),
        "ordinal_wmf_height_ratio": summarize([r["wmf_h_pt_ratio"] for r in ordinal_rows if r["wmf_h_pt_ratio"] is not None]),
        "ordinal_display_width_ratio": summarize([r["display_w_pt_ratio"] for r in ordinal_rows if r["display_w_pt_ratio"] is not None]),
        "ordinal_display_height_ratio": summarize([r["display_h_pt_ratio"] for r in ordinal_rows if r["display_h_pt_ratio"] is not None]),
        "ordinal_shape_width_ratio": summarize([r["shape_w_pt_ratio"] for r in ordinal_rows if r["shape_w_pt_ratio"] is not None]),
        "ordinal_shape_height_ratio": summarize([r["shape_h_pt_ratio"] for r in ordinal_rows if r["shape_h_pt_ratio"] is not None]),
        "wmf_width_ratio": summarize([r["wmf_w_pt_ratio"] for r in rows if r["wmf_w_pt_ratio"] is not None]),
        "wmf_height_ratio": summarize([r["wmf_h_pt_ratio"] for r in rows if r["wmf_h_pt_ratio"] is not None]),
        "wmf_width_ratio_plausible": summarize([r["wmf_w_pt_ratio"] for r in plausible_rows(rows) if r["wmf_w_pt_ratio"] is not None]),
        "wmf_height_ratio_plausible": summarize([r["wmf_h_pt_ratio"] for r in plausible_rows(rows) if r["wmf_h_pt_ratio"] is not None]),
        "shape_width_ratio": summarize([r["shape_w_pt_ratio"] for r in rows if r["shape_w_pt_ratio"] is not None]),
        "shape_height_ratio": summarize([r["shape_h_pt_ratio"] for r in rows if r["shape_h_pt_ratio"] is not None]),
        "target_metric_objects": sum(1 for item in request_math if item.get("metrics")),
        "target_wmf_width_ratio": summarize([r["target_wmf_w_pt_ratio"] for r in target_rows if r["target_wmf_w_pt_ratio"] is not None]),
        "target_wmf_height_ratio": summarize([r["target_wmf_h_pt_ratio"] for r in target_rows if r["target_wmf_h_pt_ratio"] is not None]),
        "target_shape_width_ratio": summarize([r["target_shape_w_pt_ratio"] for r in target_rows if r["target_shape_w_pt_ratio"] is not None]),
        "target_shape_height_ratio": summarize([r["target_shape_h_pt_ratio"] for r in target_rows if r["target_shape_h_pt_ratio"] is not None]),
        "wmf_width_ratio_by_class": summarize_by_class(rows, "wmf_w_pt_ratio"),
        "wmf_height_ratio_by_class": summarize_by_class(rows, "wmf_h_pt_ratio"),
        "wmf_width_ratio_by_class_plausible": summarize_by_class(plausible_rows(rows), "wmf_w_pt_ratio"),
        "wmf_height_ratio_by_class_plausible": summarize_by_class(plausible_rows(rows), "wmf_h_pt_ratio"),
        "implausible_pairs": len(rows) - len(plausible_rows(rows)),
        "shape_width_ratio_by_class": summarize_by_class(rows, "shape_w_pt_ratio"),
        "shape_height_ratio_by_class": summarize_by_class(rows, "shape_h_pt_ratio"),
        "worst_wmf_width": worst_rows(rows, "wmf_w_pt_ratio"),
        "worst_wmf_height": worst_rows(rows, "wmf_h_pt_ratio"),
        "non_wmf_generated": sum(1 for item in gen if item.get("image_ext") != "wmf"),
        "position_diff_halfpt_median": st.median([r["position_diff_halfpt"] for r in rows if r["position_diff_halfpt"] is not None])
        if any(r["position_diff_halfpt"] is not None for r in rows)
        else None,
        "detail_csv": str(detail),
    }
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source")
    parser.add_argument("generated")
    parser.add_argument("out_dir")
    parser.add_argument("legacy", nargs="*")
    parser.add_argument("--source-report")
    parser.add_argument("--generated-request")
    args = parser.parse_args()
    source_report = Path(args.source_report) if args.source_report else None
    generated_request = Path(args.generated_request) if args.generated_request else None
    if args.legacy:
        if len(args.legacy) != 2 or source_report or generated_request:
            raise SystemExit(
                "usage: compare_docx_pair_metrics.py source.docx generated.docx out_dir "
                "[source.report.json generated.request.json] [--generated-request request.json]"
            )
        source_report = Path(args.legacy[0])
        generated_request = Path(args.legacy[1])
    summary = compare_pair(Path(args.source), Path(args.generated), Path(args.out_dir), source_report, generated_request)
    out_json = Path(args.out_dir) / (Path(args.source).stem + "_summary.json")
    out_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("wrote", out_json)


if __name__ == "__main__":
    main()
