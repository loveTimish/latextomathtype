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
    label: str = "docx",
    progress_every: int = 0,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    boxes = extract_boxes(docx)
    if indexes:
        boxes = [box for box in boxes if box.index in indexes]
    if max_items is not None:
        boxes = boxes[:max_items]
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
    return rows


def read_rows(path: Path | None) -> list[dict[str, Any]] | None:
    if not path:
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
        "rows": rows,
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
    parser.add_argument("--load-reference-rows", type=Path, help="Load cached reference measurement rows instead of rendering.")
    parser.add_argument("--load-generated-rows", type=Path, help="Load cached generated measurement rows instead of rendering.")
    parser.add_argument("--save-reference-rows", type=Path, help="Write reference measurement rows for reuse.")
    parser.add_argument("--save-generated-rows", type=Path, help="Write generated measurement rows for reuse.")
    parser.add_argument("--progress-every", type=int, default=0, help="Print rendering progress every N formulas.")
    args = parser.parse_args()

    selected_indexes = parse_index_list(args.indices)
    selected_source_indexes = parse_index_list(args.source_indices)
    if selected_source_indexes:
        selected_indexes.update(index - 1 for index in selected_source_indexes)
    needs_render = not args.load_reference_rows or not args.load_generated_rows
    magick: Path | None = None
    if needs_render:
        if not args.magick.exists() and not shutil.which("magick"):
            raise SystemExit(f"ImageMagick magick.exe not found: {args.magick}")
        magick = args.magick if args.magick.exists() else Path(shutil.which("magick"))
    reference_rows = read_rows(args.load_reference_rows)
    if reference_rows is None:
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
        )
        write_rows(args.save_reference_rows, reference_rows, args.reference, selected_indexes, args.max_items)
    else:
        reference_rows = filter_rows(reference_rows, selected_indexes, args.max_items)

    generated_rows = read_rows(args.load_generated_rows)
    if generated_rows is None:
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
        )
        write_rows(args.save_generated_rows, generated_rows, args.generated, selected_indexes, args.max_items)
    else:
        generated_rows = filter_rows(generated_rows, selected_indexes, args.max_items)
    report = {
        "reference": str(args.reference.resolve()),
        "generated": str(args.generated.resolve()),
        "indices": sorted(selected_indexes),
        "source_indices": sorted(index + 1 for index in selected_indexes),
        "max_items": args.max_items,
        "load_reference_rows": str(args.load_reference_rows.resolve()) if args.load_reference_rows else None,
        "load_generated_rows": str(args.load_generated_rows.resolve()) if args.load_generated_rows else None,
        "save_reference_rows": str(args.save_reference_rows.resolve()) if args.save_reference_rows else None,
        "save_generated_rows": str(args.save_generated_rows.resolve()) if args.save_generated_rows else None,
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
