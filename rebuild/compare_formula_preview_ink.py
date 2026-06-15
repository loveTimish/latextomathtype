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
import json
import math
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import asdict
from pathlib import Path
from typing import Any

from PIL import Image

from extract_formula_boxes import extract_boxes


ROOT = Path(__file__).resolve().parents[1]
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
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    boxes = extract_boxes(docx)
    if indexes:
        boxes = [box for box in boxes if box.index in indexes]
    if max_items is not None:
        boxes = boxes[:max_items]
    with tempfile.TemporaryDirectory(prefix="formula-ink-") as tmp:
        temp_root = Path(tmp)
        for box in boxes:
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
    return rows


def parse_index_list(value: str) -> set[int]:
    indexes: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        indexes.add(int(part))
    return indexes


def compare_rows(reference: list[dict[str, Any]], generated: list[dict[str, Any]]) -> dict[str, Any]:
    reference_by_index = {int(row["index"]): row for row in reference}
    generated_by_index = {int(row["index"]): row for row in generated}
    paired_indexes = sorted(set(reference_by_index) & set(generated_by_index))
    rows: list[dict[str, Any]] = []
    for index in paired_indexes:
        ref = reference_by_index[index]
        gen = generated_by_index[index]
        if ref.get("ink_width_pt") is None or gen.get("ink_width_pt") is None:
            continue
        width_delta = float(gen["ink_width_pt"]) - float(ref["ink_width_pt"])
        height_delta = float(gen["ink_height_pt"]) - float(ref["ink_height_pt"])
        width_scale = float(ref["ink_width_pt"]) / max(float(gen["ink_width_pt"]), 0.001)
        height_scale = float(ref["ink_height_pt"]) / max(float(gen["ink_height_pt"]), 0.001)
        rows.append(
            {
                "index": index,
                "sourceIndex": index + 1,
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
                "reference_context": ref.get("context"),
                "generated_context": gen.get("context"),
            }
        )
    return {
        "paired_count": len(rows),
        "unpaired_reference": sorted(set(reference_by_index) - set(generated_by_index)),
        "unpaired_generated": sorted(set(generated_by_index) - set(reference_by_index)),
        "pairing_warning": (
            "DOCX-local indexes are only safe when object counts and order already match in strict acceptance."
            if set(reference_by_index) != set(generated_by_index)
            else ""
        ),
        "ink_width_abs_delta_pt": stats([row["ink_width_abs_delta_pt"] for row in rows]),
        "ink_height_abs_delta_pt": stats([row["ink_height_abs_delta_pt"] for row in rows]),
        "required_width_scale": stats([row["required_width_scale"] for row in rows if math.isfinite(row["required_width_scale"])]),
        "required_height_scale": stats([row["required_height_scale"] for row in rows if math.isfinite(row["required_height_scale"])]),
        "worst_width": sorted(rows, key=lambda row: row["ink_width_abs_delta_pt"], reverse=True)[:30],
        "worst_height": sorted(rows, key=lambda row: row["ink_height_abs_delta_pt"], reverse=True)[:30],
        "rows": rows,
    }


def render_text(report: dict[str, Any]) -> str:
    comparison = report["comparison"]
    lines = [
        "Formula preview ink comparison",
        "",
        f"reference: {report['reference']}",
        f"generated: {report['generated']}",
        f"paired_count: {comparison['paired_count']}",
        "",
        f"ink_width_abs_delta_pt: {comparison['ink_width_abs_delta_pt']}",
        f"ink_height_abs_delta_pt: {comparison['ink_height_abs_delta_pt']}",
        f"required_width_scale: {comparison['required_width_scale']}",
        f"required_height_scale: {comparison['required_height_scale']}",
        "",
        "Worst visible width deltas",
    ]
    for row in comparison["worst_width"][:12]:
        lines.append(
            "  #{index}/sourceIndex={sourceIndex}: ref={reference_ink_width_pt}pt gen={generated_ink_width_pt}pt "
            "delta={ink_width_delta_pt}pt scale={required_width_scale} text={generated_context}".format(**row)
        )
    lines.extend(["", "Worst visible height deltas"])
    for row in comparison["worst_height"][:12]:
        lines.append(
            "  #{index}/sourceIndex={sourceIndex}: ref={reference_ink_height_pt}pt gen={generated_ink_height_pt}pt "
            "delta={ink_height_delta_pt}pt scale={required_height_scale} text={generated_context}".format(**row)
        )
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--generated", type=Path, default=DEFAULT_GENERATED)
    parser.add_argument("--out-json", type=Path, default=DEFAULT_OUT_JSON)
    parser.add_argument("--out-text", type=Path, default=DEFAULT_OUT_TEXT)
    parser.add_argument("--magick", type=Path, default=DEFAULT_MAGICK)
    parser.add_argument("--white-threshold", type=int, default=245)
    parser.add_argument("--alpha-threshold", type=int, default=8)
    parser.add_argument("--indices", default="", help="Comma-separated zero-based formula indexes to compare; sourceIndex is index + 1.")
    parser.add_argument("--source-indices", default="", help="Comma-separated one-based sourceIndex values to compare.")
    parser.add_argument("--max-items", type=int, help="Compare only the first N formulas after index filtering.")
    args = parser.parse_args()

    if not args.magick.exists() and not shutil.which("magick"):
        raise SystemExit(f"ImageMagick magick.exe not found: {args.magick}")
    magick = args.magick if args.magick.exists() else Path(shutil.which("magick"))

    selected_indexes = parse_index_list(args.indices)
    selected_source_indexes = parse_index_list(args.source_indices)
    if selected_source_indexes:
        selected_indexes.update(index - 1 for index in selected_source_indexes)
    reference_rows = measure_docx(
        args.reference,
        magick,
        args.white_threshold,
        args.alpha_threshold,
        selected_indexes,
        args.max_items,
    )
    generated_rows = measure_docx(
        args.generated,
        magick,
        args.white_threshold,
        args.alpha_threshold,
        selected_indexes,
        args.max_items,
    )
    report = {
        "reference": str(args.reference.resolve()),
        "generated": str(args.generated.resolve()),
        "indices": sorted(selected_indexes),
        "source_indices": sorted(index + 1 for index in selected_indexes),
        "max_items": args.max_items,
        "comparison": compare_rows(reference_rows, generated_rows),
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
