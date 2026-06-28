# -*- coding: utf-8 -*-
"""
Compare visible ink bounds of MathType preview images embedded in DOCX files.

Formula OLE display boxes can match while the preview image ink is still too
small inside the box.  This script renders the preview media (WMF/PNG/etc.) to
bitmap form, computes non-background ink bounds, and reports the visible ink
size in Word points after scaling into the object box.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import asdict
from pathlib import Path
from typing import Any

from PIL import Image

from extract_formula_boxes import extract_boxes


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from compare_docx_pair_metrics import (  # noqa: E402
    latex_alignment,
    normalize_latex_key,
    ordinal_latex_key,
    parse_request_math,
)

DEFAULT_REFERENCE = ROOT / "rebuild-assets/external/fraction-split-reference.docx"
DEFAULT_GENERATED = ROOT / "target/reference-roundtrip/fraction-split-reference-regenerated.docx"
DEFAULT_OUT_JSON = ROOT / "target/reference-roundtrip/formula-preview-ink-comparison.json"
DEFAULT_OUT_TEXT = ROOT / "target/reference-roundtrip/formula-preview-ink-comparison.txt"
DEFAULT_MAGICK = Path(r"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe")


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
        "avg": round(sum(values) / len(values), 3) if values else None,
        "min": round(min(values), 3) if values else None,
        "p10": percentile(values, 0.10),
        "median": percentile(values, 0.50),
        "p90": percentile(values, 0.90),
        "max": round(max(values), 3) if values else None,
    }


def resolve_media(docx: Path, target: str | None) -> bytes | None:
    if not target:
        return None
    name = "word/" + target if not target.startswith("word/") else target
    with zipfile.ZipFile(docx) as zf:
        try:
            return zf.read(name)
        except KeyError:
            return None


def render_media_to_png(media_bytes: bytes, suffix: str, magick: Path, temp_dir: Path) -> Path | None:
    input_path = temp_dir / f"input{suffix}"
    output_path = temp_dir / "output.png"
    input_path.write_bytes(media_bytes)
    if output_path.exists():
        output_path.unlink()
    if suffix.lower() == ".png":
        output_path.write_bytes(media_bytes)
        return output_path
    command = [str(magick), str(input_path), str(output_path)]
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0 or not output_path.exists():
        return None
    return output_path


def is_ink_pixel(pixel: tuple[int, ...], white_threshold: int, alpha_threshold: int) -> bool:
    if len(pixel) == 4 and pixel[3] <= alpha_threshold:
        return False
    rgb = pixel[:3]
    return any(channel < white_threshold for channel in rgb)


def compute_ink_bounds(png_path: Path, white_threshold: int, alpha_threshold: int) -> dict[str, Any]:
    with Image.open(png_path).convert("RGBA") as image:
        width, height = image.size
        pixels = image.load()
        min_x = width
        min_y = height
        max_x = -1
        max_y = -1
        ink_count = 0
        for y in range(height):
            for x in range(width):
                if is_ink_pixel(pixels[x, y], white_threshold, alpha_threshold):
                    ink_count += 1
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)
        if max_x < 0:
            return {
                "image_width_px": width,
                "image_height_px": height,
                "ink_count": 0,
                "ink_bbox_px": None,
                "ink_width_px": 0,
                "ink_height_px": 0,
                "ink_left_px": None,
                "ink_top_px": None,
            }
        return {
            "image_width_px": width,
            "image_height_px": height,
            "ink_count": ink_count,
            "ink_bbox_px": [min_x, min_y, max_x + 1, max_y + 1],
            "ink_width_px": max_x - min_x + 1,
            "ink_height_px": max_y - min_y + 1,
            "ink_left_px": min_x,
            "ink_top_px": min_y,
        }


def measure_docx(
    docx: Path,
    magick: Path,
    white_threshold: int,
    alpha_threshold: int,
    indexes: set[int] | None = None,
    max_items: int | None = None,
    label: str = "docx",
    progress_every: int = 0,
    existing_rows: list[dict[str, Any]] | None = None,
    checkpoint_path: Path | None = None,
    checkpoint_every: int = 0,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = filter_rows(list(existing_rows or []), indexes, max_items)
    completed_indexes = {int(row.get("index", -1)) for row in rows}
    boxes = extract_boxes(docx)
    if indexes:
        boxes = [box for box in boxes if box.index in indexes]
    if max_items is not None:
        boxes = boxes[:max_items]
    boxes = [box for box in boxes if box.index not in completed_indexes]
    with tempfile.TemporaryDirectory(prefix="formula-ink-") as tmp:
        temp_root = Path(tmp)
        total = len(boxes)
        for ordinal, box in enumerate(boxes, start=1):
            if progress_every > 0 and (ordinal == 1 or ordinal % progress_every == 0 or ordinal == total):
                print(f"[{label}] measuring {ordinal}/{total} sourceIndex={box.index + 1}", flush=True)
            row = asdict(box)
            media = resolve_media(docx, box.image_target)
            suffix = Path(box.image_target or "").suffix or ".bin"
            media_dir = temp_root / f"media-{box.index}"
            media_dir.mkdir()
            rendered = render_media_to_png(media, suffix, magick, media_dir) if media else None
            if rendered is None:
                row.update(
                    {
                        "preview_error": "render_failed",
                        "image_width_px": None,
                        "image_height_px": None,
                        "ink_bbox_px": None,
                    }
                )
                rows.append(row)
                continue
            ink = compute_ink_bounds(rendered, white_threshold, alpha_threshold)
            row.update(ink)
            image_w = max(float(ink["image_width_px"]), 1.0)
            image_h = max(float(ink["image_height_px"]), 1.0)
            row["ink_width_pt"] = round(float(ink["ink_width_px"]) / image_w * box.style_width_pt, 3)
            row["ink_height_pt"] = round(float(ink["ink_height_px"]) / image_h * box.style_height_pt, 3)
            row["ink_left_pt"] = (
                round(float(ink["ink_left_px"]) / image_w * box.style_width_pt, 3)
                if ink["ink_left_px"] is not None
                else None
            )
            row["ink_top_pt"] = (
                round(float(ink["ink_top_px"]) / image_h * box.style_height_pt, 3)
                if ink["ink_top_px"] is not None
                else None
            )
            row["ink_width_ratio"] = round(float(ink["ink_width_px"]) / image_w, 4)
            row["ink_height_ratio"] = round(float(ink["ink_height_px"]) / image_h, 4)
            rows.append(row)
            if checkpoint_path and checkpoint_every > 0 and (ordinal % checkpoint_every == 0 or ordinal == total):
                write_rows(checkpoint_path, rows, docx, indexes or set(), max_items)
    return rows


def read_rows(path: Path | None) -> list[dict[str, Any]] | None:
    if not path:
        return None
    if not path.exists():
        return None
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict) and isinstance(payload.get("rows"), list):
        return payload["rows"]
    raise ValueError(f"measurement cache does not contain rows: {path}")


def write_rows(path: Path | None, rows: list[dict[str, Any]], docx: Path, indexes: set[int], max_items: int | None) -> None:
    if not path:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "docx": str(docx.resolve()),
        "indices": sorted(indexes),
        "source_indices": sorted(index + 1 for index in indexes),
        "max_items": max_items,
        "row_count": len(rows),
        "rows": sorted(rows, key=lambda row: int(row.get("index", -1))),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def filter_rows(rows: list[dict[str, Any]], indexes: set[int] | None, max_items: int | None) -> list[dict[str, Any]]:
    out = rows
    if indexes:
        out = [row for row in out if int(row.get("index", -1)) in indexes]
    if max_items is not None:
        out = out[:max_items]
    return out


def parse_index_list(value: str) -> set[int]:
    indexes: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        indexes.add(int(part))
    return indexes


def read_source_report_records(report_path: Path) -> list[dict[str, Any]]:
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    records: list[dict[str, Any]] = []
    for item in report.get("equations") or []:
        if item.get("status") != "converted" or not item.get("output"):
            continue
        raw_index = item.get("docObjectIndex")
        doc_object_index = int(raw_index) if raw_index else None
        records.append(
            {
                "reportPosition": len(records),
                "latex": normalize_latex_key(item.get("output", "")),
                "docObjectIndex": doc_object_index,
                "referenceIndex": doc_object_index - 1 if doc_object_index else None,
            }
        )
    return records


TRACE_ASCII_WS = r"[ \t\n\x0b\f\r]"
TRACE_METRICS_RE = re.compile(r"^\\pwmetrics\{[^}]+}" + TRACE_ASCII_WS + "*")
TRACE_STYLE_RE = re.compile(r"^\\pwstyle\{[^}]*}" + TRACE_ASCII_WS + "*")
TRACE_EDGE_SPACE_RE = re.compile(r"^" + TRACE_ASCII_WS + r"+|" + TRACE_ASCII_WS + r"+$")
TRACE_SPACE_RE = re.compile(TRACE_ASCII_WS + "+")


def formula_trace_hash(latex: str) -> str:
    value = (latex or "").replace("\u00a0", " ")
    value = TRACE_METRICS_RE.sub("", value)
    value = TRACE_STYLE_RE.sub("", value)
    value = TRACE_EDGE_SPACE_RE.sub("", value)
    value = TRACE_SPACE_RE.sub(" ", value)
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:16]


def formula_trace_id(index: int, latex: str) -> str:
    return f"pwf:{index + 1}-{formula_trace_hash(latex)}"


def legacy_formula_trace_id(latex: str) -> str:
    return "pwf:" + formula_trace_hash(latex)


def pair_by_key(
    source_values: list[str],
    generated_values: list[str],
    source_candidates: set[int] | None = None,
    generated_candidates: set[int] | None = None,
) -> list[tuple[int, int]]:
    generated_by_key: dict[str, list[int]] = {}
    allowed_generated = generated_candidates if generated_candidates is not None else set(range(len(generated_values)))
    allowed_source = source_candidates if source_candidates is not None else set(range(len(source_values)))
    source_counts = Counter(source_values[index] for index in allowed_source)
    generated_counts = Counter(generated_values[index] for index in allowed_generated)
    for index, value in enumerate(generated_values):
        if index in allowed_generated and source_counts[value] == 1 and generated_counts[value] == 1:
            generated_by_key.setdefault(value, []).append(index)
    pairs: list[tuple[int, int]] = []
    used_generated: set[int] = set()
    for source_index, value in enumerate(source_values):
        if source_index not in allowed_source:
            continue
        if source_counts[value] != 1 or generated_counts[value] != 1:
            continue
        queue = generated_by_key.get(value) or []
        while queue and queue[0] in used_generated:
            queue.pop(0)
        if queue:
            generated_index = queue.pop(0)
            used_generated.add(generated_index)
            pairs.append((source_index, generated_index))
    return pairs


def build_latex_alignment(
    source_report: Path | None,
    generated_request: Path | None,
    reference_rows: list[dict[str, Any]],
    generated_rows: list[dict[str, Any]],
) -> tuple[list[tuple[int, int]] | None, dict[str, Any]]:
    if not source_report or not generated_request:
        return None, {"alignment_mode": "ordinal"}

    source_records = read_source_report_records(source_report)
    source_latex = [record["latex"] for record in source_records]
    request_math = parse_request_math(generated_request)
    generated_latex = [item["latex"] for item in request_math]
    generated_trace_count = sum(1 for row in generated_rows if str(row.get("image_title") or "").startswith("pwf:"))
    generated_position_by_trace: dict[int, int] = {}
    if generated_trace_count:
        request_trace_indexes: dict[str, list[int]] = {}
        row_trace_indexes: dict[str, list[int]] = {}
        for index, item in enumerate(request_math):
            latex = item.get("rawLatex") or item.get("latex") or ""
            request_trace_indexes.setdefault(formula_trace_id(index, latex), []).append(index)
            request_trace_indexes.setdefault(legacy_formula_trace_id(latex), []).append(index)
        for row in generated_rows:
            trace = str(row.get("image_title") or "")
            if trace.startswith("pwf:"):
                row_trace_indexes.setdefault(trace, []).append(int(row["index"]))
        for trace, request_indexes in request_trace_indexes.items():
            row_indexes = row_trace_indexes.get(trace) or []
            if len(request_indexes) == 1 and len(row_indexes) == 1:
                generated_position_by_trace[request_indexes[0]] = row_indexes[0]
    exact_pairs = latex_alignment(source_latex, generated_latex)
    exact_source_positions = {source_position for source_position, _ in exact_pairs}
    exact_generated_positions = {generated_position for _, generated_position in exact_pairs}
    ordinal_source_latex = [ordinal_latex_key(value) for value in source_latex]
    ordinal_generated_latex = [ordinal_latex_key(value) for value in generated_latex]
    fallback_pairs = pair_by_key(
        ordinal_source_latex,
        ordinal_generated_latex,
        set(range(len(source_latex))) - exact_source_positions,
        set(range(len(generated_latex))) - exact_generated_positions,
    )
    pairs = exact_pairs + fallback_pairs
    pair_methods = {pair: "exact" for pair in exact_pairs}
    pair_methods.update({pair: "ordinal_key" for pair in fallback_pairs})
    reference_indexes = {int(row["index"]) for row in reference_rows}
    generated_indexes = {int(row["index"]) for row in generated_rows}
    usable_pairs: list[tuple[int, int]] = []
    usable_pair_methods: dict[str, str] = {}
    usable_pair_trace_status: dict[str, str] = {}
    for source_position, generated_position in pairs:
        if source_position >= len(source_records):
            continue
        reference_index = source_records[source_position]["referenceIndex"]
        if reference_index is None:
            continue
        reference_index = int(reference_index)
        generated_row_index = generated_position_by_trace.get(generated_position, generated_position)
        if reference_index in reference_indexes and generated_row_index in generated_indexes:
            usable_pairs.append((reference_index, generated_row_index))
            usable_pair_methods[f"{reference_index}:{generated_row_index}"] = pair_methods.get((source_position, generated_position), "unknown")
            if generated_trace_count:
                status = "matched" if generated_position in generated_position_by_trace else "unmatched"
            else:
                status = "absent"
            usable_pair_trace_status[f"{reference_index}:{generated_row_index}"] = status
    paired_source_positions = {source_position for source_position, _ in pairs}
    paired_generated_positions = {generated_position for _, generated_position in pairs}
    known_doc_indexes = [record["docObjectIndex"] for record in source_records if record["docObjectIndex"] is not None]
    missing_doc_object_index_count = sum(1 for record in source_records if record["docObjectIndex"] is None)
    doc_index_gap_count = 0
    if known_doc_indexes:
        doc_index_gap_count = len(set(range(min(known_doc_indexes), max(known_doc_indexes) + 1)) - set(known_doc_indexes))
    return usable_pairs, {
        "alignment_mode": "latex",
        "source_latex_count": len(source_latex),
        "generated_latex_count": len(generated_latex),
        "reference_row_count": len(reference_rows),
        "generated_row_count": len(generated_rows),
        "generated_trace_count": generated_trace_count,
        "generated_trace_match_count": len(generated_position_by_trace),
        "latex_pair_count": len(pairs),
        "exact_latex_pair_count": len(exact_pairs),
        "ordinal_key_latex_pair_count": len(fallback_pairs),
        "usable_latex_pair_count": len(usable_pairs),
        "usable_pair_methods": usable_pair_methods,
        "usable_pair_trace_status": usable_pair_trace_status,
        "unpaired_source_latex_count": len(source_latex) - len(paired_source_positions),
        "unpaired_generated_latex_count": len(generated_latex) - len(paired_generated_positions),
        "source_doc_object_index_min": min(known_doc_indexes, default=None),
        "source_doc_object_index_max": max(known_doc_indexes, default=None),
        "source_doc_object_index_gaps": doc_index_gap_count,
        "missing_doc_object_index_count": missing_doc_object_index_count,
        "unpaired_source_latex_samples": [
            {
                "reportPosition": record["reportPosition"],
                "sourceIndex": record["docObjectIndex"] or (record["reportPosition"] + 1),
                "docObjectIndex": record["docObjectIndex"],
                "latex": record["latex"],
            }
            for record in source_records
            if record["reportPosition"] not in paired_source_positions
        ][:12],
        "unpaired_generated_latex_samples": [
            {
                "generatedPosition": index,
                "generatedSourceIndex": index + 1,
                "latex": generated_latex[index],
            }
            for index in range(len(generated_latex))
            if index not in paired_generated_positions
        ][:12],
    }


def normalize_context(value: Any) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", "", text)


def text_similarity(left: Any, right: Any) -> float:
    a = normalize_context(left)
    b = normalize_context(right)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return round(min(len(a), len(b)) / max(len(a), len(b)), 3)
    a_tokens = set(a)
    b_tokens = set(b)
    if not a_tokens or not b_tokens:
        return 0.0
    return round(len(a_tokens & b_tokens) / len(a_tokens | b_tokens), 3)


def compare_rows(
    reference: list[dict[str, Any]],
    generated: list[dict[str, Any]],
    index_pairs: list[tuple[int, int]] | None = None,
    alignment_info: dict[str, Any] | None = None,
) -> dict[str, Any]:
    reference_by_index = {int(row["index"]): row for row in reference}
    generated_by_index = {int(row["index"]): row for row in generated}
    alignment_info = dict(alignment_info or {})
    alignment_mode = alignment_info.get("alignment_mode") or "ordinal"
    if index_pairs is None:
        index_pairs = [(index, index) for index in sorted(set(reference_by_index) & set(generated_by_index))]
    rows: list[dict[str, Any]] = []
    paired_reference_indexes = set()
    paired_generated_indexes = set()
    paired_without_ink: list[dict[str, Any]] = []
    for source_index, generated_index in index_pairs:
        ref = reference_by_index.get(source_index)
        gen = generated_by_index.get(generated_index)
        if ref is None or gen is None:
            continue
        pair_method = (alignment_info.get("usable_pair_methods") or {}).get(f"{source_index}:{generated_index}")
        trace_status = (alignment_info.get("usable_pair_trace_status") or {}).get(f"{source_index}:{generated_index}")
        if ref.get("ink_width_pt") is None or gen.get("ink_width_pt") is None:
            paired_without_ink.append(
                {
                    "reference_index": source_index,
                    "sourceIndex": source_index + 1,
                    "generated_index": generated_index,
                    "generatedSourceIndex": generated_index + 1,
                    "pair_method": pair_method,
                    "generated_trace_status": trace_status,
                    "reference_preview_error": ref.get("preview_error"),
                    "generated_preview_error": gen.get("preview_error"),
                }
            )
            continue
        paired_reference_indexes.add(source_index)
        paired_generated_indexes.add(generated_index)
        width_delta = float(gen["ink_width_pt"]) - float(ref["ink_width_pt"])
        height_delta = float(gen["ink_height_pt"]) - float(ref["ink_height_pt"])
        width_scale = float(ref["ink_width_pt"]) / max(float(gen["ink_width_pt"]), 0.001)
        height_scale = float(ref["ink_height_pt"]) / max(float(gen["ink_height_pt"]), 0.001)
        reference_context = ref.get("context")
        generated_context = gen.get("context")
        context_similarity = text_similarity(reference_context, generated_context)
        extreme_scale = max(
            width_scale,
            1.0 / max(width_scale, 0.001),
            height_scale,
            1.0 / max(height_scale, 0.001),
        )
        alignment_suspicious = context_similarity < 0.15 or extreme_scale > 3.0 or trace_status == "unmatched"
        rows.append(
            {
                "index": source_index,
                "sourceIndex": source_index + 1,
                "reference_index": source_index,
                "generated_index": generated_index,
                "generatedSourceIndex": generated_index + 1,
                "pair_method": pair_method,
                "generated_trace_status": trace_status,
                "reference_ink_width_pt": ref["ink_width_pt"],
                "generated_ink_width_pt": gen["ink_width_pt"],
                "reference_ink_height_pt": ref["ink_height_pt"],
                "generated_ink_height_pt": gen["ink_height_pt"],
                "ink_width_delta_pt": round(width_delta, 3),
                "ink_height_delta_pt": round(height_delta, 3),
                "ink_width_abs_delta_pt": round(abs(width_delta), 3),
                "ink_height_abs_delta_pt": round(abs(height_delta), 3),
                "required_width_scale": round(width_scale, 4),
                "required_height_scale": round(height_scale, 4),
                "reference_image_size_px": [ref["image_width_px"], ref["image_height_px"]],
                "generated_image_size_px": [gen["image_width_px"], gen["image_height_px"]],
                "reference_ink_bbox_px": ref["ink_bbox_px"],
                "generated_ink_bbox_px": gen["ink_bbox_px"],
                "context_similarity": context_similarity,
                "extreme_scale": round(extreme_scale, 4),
                "alignment_suspicious": alignment_suspicious,
                "reference_context": reference_context,
                "generated_context": generated_context,
            }
        )
    aligned_rows = [row for row in rows if not row["alignment_suspicious"]]
    pairing_warnings: list[str] = []
    pairing_notes: list[str] = []
    if alignment_mode == "ordinal" and set(reference_by_index) != set(generated_by_index):
        pairing_warnings.append("DOCX-local indexes are only safe when object counts and order already match in strict acceptance.")
    if alignment_mode == "latex":
        if alignment_info.get("usable_latex_pair_count") != len(reference_by_index):
            pairing_warnings.append("LaTeX alignment covers only a subset of measured reference rows; do not treat summary stats as full-document acceptance.")
        if alignment_info.get("ordinal_key_latex_pair_count"):
            pairing_warnings.append("Ordinal-key fallback pairs are diagnostic only because style stripping is lossy.")
        if alignment_info.get("source_doc_object_index_gaps"):
            pairing_warnings.append("Source docObjectIndex has gaps; remaining unpaired rows may be non-formula or filtered source objects.")
        if alignment_info.get("missing_doc_object_index_count"):
            pairing_warnings.append("Some source report equations lack docObjectIndex and were excluded from source row mapping.")
        if alignment_info.get("generated_latex_count") == len(generated_by_index):
            if alignment_info.get("generated_trace_count") == len(generated_by_index):
                if alignment_info.get("generated_trace_match_count") == alignment_info.get("generated_latex_count"):
                    pairing_notes.append("Generated DOCX trace ids matched request formulas one-to-one.")
                else:
                    pairing_warnings.append("Generated DOCX carries formula trace ids, but some request formulas could not be matched uniquely by trace.")
                    pairing_warnings.append("Rows without a unique generated trace match are marked alignment_suspicious.")
            else:
                pairing_warnings.append("Generated request order is assumed to match generated DOCX preview order; this is not an explicit renderer object map.")
        if alignment_info.get("generated_latex_count") != len(generated_by_index):
            pairing_warnings.append("Generated request formula count differs from generated preview row count; generated positions may shift.")
    summary = {
        "alignment_mode": alignment_mode,
        "paired_count": len(rows),
        "index_pair_count": len(index_pairs),
        "unpaired_reference": sorted(set(reference_by_index) - paired_reference_indexes),
        "unpaired_generated": sorted(set(generated_by_index) - paired_generated_indexes),
        "paired_without_ink_count": len(paired_without_ink),
        "paired_without_ink": paired_without_ink[:30],
        "pairing_warning": " ".join(pairing_warnings),
        "pairing_note": " ".join(pairing_notes),
        "ink_width_abs_delta_pt": stats([row["ink_width_abs_delta_pt"] for row in rows]),
        "ink_height_abs_delta_pt": stats([row["ink_height_abs_delta_pt"] for row in rows]),
        "aligned_paired_count": len(aligned_rows),
        "alignment_suspicious_count": len(rows) - len(aligned_rows),
        "aligned_ink_width_abs_delta_pt": stats([row["ink_width_abs_delta_pt"] for row in aligned_rows]),
        "aligned_ink_height_abs_delta_pt": stats([row["ink_height_abs_delta_pt"] for row in aligned_rows]),
        "required_width_scale": stats([row["required_width_scale"] for row in rows if math.isfinite(row["required_width_scale"])]),
        "required_height_scale": stats([row["required_height_scale"] for row in rows if math.isfinite(row["required_height_scale"])]),
        "worst_width": sorted(rows, key=lambda row: row["ink_width_abs_delta_pt"], reverse=True)[:30],
        "worst_height": sorted(rows, key=lambda row: row["ink_height_abs_delta_pt"], reverse=True)[:30],
        "worst_aligned_width": sorted(aligned_rows, key=lambda row: row["ink_width_abs_delta_pt"], reverse=True)[:30],
        "worst_aligned_height": sorted(aligned_rows, key=lambda row: row["ink_height_abs_delta_pt"], reverse=True)[:30],
        "alignment_suspicious_rows": [row for row in rows if row["alignment_suspicious"]][:30],
        "rows": rows,
    }
    summary.update(alignment_info)
    return summary


def render_text(report: dict[str, Any]) -> str:
    comparison = report["comparison"]
    lines = [
        "Formula preview ink comparison",
        "",
        f"reference: {report['reference']}",
        f"generated: {report['generated']}",
        f"alignment_mode: {comparison['alignment_mode']}",
        f"paired_count: {comparison['paired_count']}",
        f"index_pair_count: {comparison['index_pair_count']}",
        f"source_latex_count: {comparison.get('source_latex_count')}",
        f"generated_latex_count: {comparison.get('generated_latex_count')}",
        f"reference_row_count: {comparison.get('reference_row_count')}",
        f"generated_row_count: {comparison.get('generated_row_count')}",
        f"generated_trace_count: {comparison.get('generated_trace_count')}",
        f"generated_trace_match_count: {comparison.get('generated_trace_match_count')}",
        f"latex_pair_count: {comparison.get('latex_pair_count')}",
        f"exact_latex_pair_count: {comparison.get('exact_latex_pair_count')}",
        f"ordinal_key_latex_pair_count: {comparison.get('ordinal_key_latex_pair_count')}",
        f"usable_latex_pair_count: {comparison.get('usable_latex_pair_count')}",
        f"source_doc_object_index_range: {comparison.get('source_doc_object_index_min')}..{comparison.get('source_doc_object_index_max')}",
        f"source_doc_object_index_gaps: {comparison.get('source_doc_object_index_gaps')}",
        f"missing_doc_object_index_count: {comparison.get('missing_doc_object_index_count')}",
        f"paired_without_ink_count: {comparison.get('paired_without_ink_count')}",
        f"pairing_warning: {comparison.get('pairing_warning')}",
        f"pairing_note: {comparison.get('pairing_note')}",
        f"aligned_paired_count: {comparison['aligned_paired_count']}",
        f"alignment_suspicious_count: {comparison['alignment_suspicious_count']}",
        "",
        f"ink_width_abs_delta_pt: {comparison['ink_width_abs_delta_pt']}",
        f"ink_height_abs_delta_pt: {comparison['ink_height_abs_delta_pt']}",
        f"aligned_ink_width_abs_delta_pt: {comparison['aligned_ink_width_abs_delta_pt']}",
        f"aligned_ink_height_abs_delta_pt: {comparison['aligned_ink_height_abs_delta_pt']}",
        f"required_width_scale: {comparison['required_width_scale']}",
        f"required_height_scale: {comparison['required_height_scale']}",
        "",
        "Worst visible width deltas",
    ]
    for row in comparison["worst_width"][:12]:
        lines.append(
            "  #{index}/sourceIndex={sourceIndex}->generatedSourceIndex={generatedSourceIndex} method={pair_method}: ref={reference_ink_width_pt}pt gen={generated_ink_width_pt}pt "
            "delta={ink_width_delta_pt}pt scale={required_width_scale} suspicious={alignment_suspicious} "
            "sim={context_similarity} text={generated_context}".format(**row)
        )
    lines.extend(["", "Worst visible height deltas"])
    for row in comparison["worst_height"][:12]:
        lines.append(
            "  #{index}/sourceIndex={sourceIndex}->generatedSourceIndex={generatedSourceIndex} method={pair_method}: ref={reference_ink_height_pt}pt gen={generated_ink_height_pt}pt "
            "delta={ink_height_delta_pt}pt scale={required_height_scale} suspicious={alignment_suspicious} "
            "sim={context_similarity} text={generated_context}".format(**row)
        )
    lines.extend(["", "Worst aligned-looking width deltas"])
    for row in comparison["worst_aligned_width"][:12]:
        lines.append(
            "  #{index}/sourceIndex={sourceIndex}->generatedSourceIndex={generatedSourceIndex} method={pair_method}: ref={reference_ink_width_pt}pt gen={generated_ink_width_pt}pt "
            "delta={ink_width_delta_pt}pt scale={required_width_scale} sim={context_similarity} text={generated_context}".format(**row)
        )
    lines.extend(["", "Worst aligned-looking height deltas"])
    for row in comparison["worst_aligned_height"][:12]:
        lines.append(
            "  #{index}/sourceIndex={sourceIndex}->generatedSourceIndex={generatedSourceIndex} method={pair_method}: ref={reference_ink_height_pt}pt gen={generated_ink_height_pt}pt "
            "delta={ink_height_delta_pt}pt scale={required_height_scale} sim={context_similarity} text={generated_context}".format(**row)
        )
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--generated", type=Path, default=DEFAULT_GENERATED)
    parser.add_argument("--source-report", type=Path, help="Source docx2tex report for LaTeX-key alignment.")
    parser.add_argument("--generated-request", type=Path, help="Generated request JSON for LaTeX-key alignment.")
    parser.add_argument("--out-json", type=Path, default=DEFAULT_OUT_JSON)
    parser.add_argument("--out-text", type=Path, default=DEFAULT_OUT_TEXT)
    parser.add_argument("--magick", type=Path, default=DEFAULT_MAGICK)
    parser.add_argument("--white-threshold", type=int, default=245)
    parser.add_argument("--alpha-threshold", type=int, default=8)
    parser.add_argument("--indices", default="", help="Comma-separated zero-based formula indexes to compare; sourceIndex is index + 1.")
    parser.add_argument("--source-indices", default="", help="Comma-separated one-based sourceIndex values to compare.")
    parser.add_argument("--max-items", type=int, help="Compare only the first N formulas after index filtering.")
    parser.add_argument("--load-reference-rows", type=Path, help="Load cached reference measurement rows instead of rendering.")
    parser.add_argument("--load-generated-rows", type=Path, help="Load cached generated measurement rows instead of rendering.")
    parser.add_argument("--resume-reference-rows", type=Path, help="Resume reference measurement from an existing partial row cache.")
    parser.add_argument("--resume-generated-rows", type=Path, help="Resume generated measurement from an existing partial row cache.")
    parser.add_argument("--save-reference-rows", type=Path, help="Write reference measurement rows for reuse.")
    parser.add_argument("--save-generated-rows", type=Path, help="Write generated measurement rows for reuse.")
    parser.add_argument("--progress-every", type=int, default=0, help="Print rendering progress every N formulas.")
    parser.add_argument("--checkpoint-every", type=int, default=0, help="Save partial measurement rows every N newly rendered formulas.")
    parser.add_argument(
        "--measure-only",
        choices=["reference", "generated"],
        help="Only measure and save one side; requires the matching --save-*-rows option.",
    )
    args = parser.parse_args()
    if args.measure_only == "reference" and not args.save_reference_rows:
        parser.error("--measure-only reference requires --save-reference-rows")
    if args.measure_only == "generated" and not args.save_generated_rows:
        parser.error("--measure-only generated requires --save-generated-rows")
    if args.measure_only == "reference" and args.load_reference_rows:
        parser.error("--measure-only reference cannot be combined with --load-reference-rows; use --resume-reference-rows")
    if args.measure_only == "generated" and args.load_generated_rows:
        parser.error("--measure-only generated cannot be combined with --load-generated-rows; use --resume-generated-rows")

    selected_indexes = parse_index_list(args.indices)
    selected_source_indexes = parse_index_list(args.source_indices)
    if selected_source_indexes:
        selected_indexes.update(index - 1 for index in selected_source_indexes)
    needs_reference = args.measure_only in (None, "reference")
    needs_generated = args.measure_only in (None, "generated")
    needs_render = (
        (needs_reference and not args.load_reference_rows)
        or (needs_generated and not args.load_generated_rows)
    )
    magick: Path | None = None
    if needs_render:
        if not args.magick.exists() and not shutil.which("magick"):
            raise SystemExit(f"ImageMagick magick.exe not found: {args.magick}")
        magick = args.magick if args.magick.exists() else Path(shutil.which("magick"))
    reference_rows = read_rows(args.load_reference_rows) if needs_reference else []
    if needs_reference and reference_rows is None:
        existing_reference_rows = read_rows(args.resume_reference_rows) or []
        assert magick is not None
        reference_rows = measure_docx(
            args.reference,
            magick,
            args.white_threshold,
            args.alpha_threshold,
            selected_indexes,
            args.max_items,
            "reference",
            args.progress_every,
            existing_reference_rows,
            args.save_reference_rows,
            args.checkpoint_every,
        )
        write_rows(args.save_reference_rows, reference_rows, args.reference, selected_indexes, args.max_items)
    else:
        reference_rows = filter_rows(reference_rows, selected_indexes, args.max_items)
    if args.measure_only == "reference":
        print(f"measured reference rows: {len(reference_rows)}")
        return 0

    generated_rows = read_rows(args.load_generated_rows) if needs_generated else []
    if needs_generated and generated_rows is None:
        existing_generated_rows = read_rows(args.resume_generated_rows) or []
        assert magick is not None
        generated_rows = measure_docx(
            args.generated,
            magick,
            args.white_threshold,
            args.alpha_threshold,
            selected_indexes,
            args.max_items,
            "generated",
            args.progress_every,
            existing_generated_rows,
            args.save_generated_rows,
            args.checkpoint_every,
        )
        write_rows(args.save_generated_rows, generated_rows, args.generated, selected_indexes, args.max_items)
    else:
        generated_rows = filter_rows(generated_rows, selected_indexes, args.max_items)
    if args.measure_only == "generated":
        print(f"measured generated rows: {len(generated_rows)}")
        return 0
    alignment_info: dict[str, Any] = {"alignment_mode": "ordinal"}
    index_pairs: list[tuple[int, int]] | None = None
    if args.source_report and args.generated_request:
        index_pairs, alignment_info = build_latex_alignment(
            args.source_report,
            args.generated_request,
            reference_rows,
            generated_rows,
        )
    report = {
        "reference": str(args.reference.resolve()),
        "generated": str(args.generated.resolve()),
        "source_report": str(args.source_report.resolve()) if args.source_report else None,
        "generated_request": str(args.generated_request.resolve()) if args.generated_request else None,
        "indices": sorted(selected_indexes),
        "source_indices": sorted(index + 1 for index in selected_indexes),
        "max_items": args.max_items,
        "load_reference_rows": str(args.load_reference_rows.resolve()) if args.load_reference_rows else None,
        "load_generated_rows": str(args.load_generated_rows.resolve()) if args.load_generated_rows else None,
        "save_reference_rows": str(args.save_reference_rows.resolve()) if args.save_reference_rows else None,
        "save_generated_rows": str(args.save_generated_rows.resolve()) if args.save_generated_rows else None,
        "comparison": compare_rows(reference_rows, generated_rows, index_pairs, alignment_info),
        "reference_rows": reference_rows,
        "generated_rows": generated_rows,
    }
    args.out_json.parent.mkdir(parents=True, exist_ok=True)
    args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    args.out_text.write_text(render_text(report), encoding="utf-8")
    print(render_text(report))
    print(f"wrote {args.out_json}")
    print(f"wrote {args.out_text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
