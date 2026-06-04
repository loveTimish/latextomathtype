#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compare rendered-page ink bounds for MathType inline formulas.

The layout JSON is produced by rebuild/compare_word_formula_layout.ps1.  This
script crops each formula's Word object rectangle from page PNGs, detects the
non-white pixels inside that rendered page crop, and compares the visible ink
bbox in page coordinates.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw


DEFAULT_LAYOUT = Path("target/reference-roundtrip/word-formula-layout-comparison.json")
DEFAULT_REFERENCE_PAGES = Path("target/reference-roundtrip/visual-reference-current")
DEFAULT_GENERATED_PAGES = Path("target/reference-roundtrip/visual-generated-current")
DEFAULT_OUT_JSON = Path("target/reference-roundtrip/rendered-formula-ink-comparison.json")
DEFAULT_OUT_TEXT = Path("target/reference-roundtrip/rendered-formula-ink-comparison.txt")


def percentile(values: list[float], pct: float) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return round(ordered[0], 3)
    rank = (len(ordered) - 1) * pct
    lo = int(rank)
    hi = min(lo + 1, len(ordered) - 1)
    fraction = rank - lo
    return round(ordered[lo] + (ordered[hi] - ordered[lo]) * fraction, 3)


def stats(values: list[float]) -> dict[str, float | int | None]:
    finite = [value for value in values if math.isfinite(value)]
    return {
        "n": len(finite),
        "avg": round(sum(finite) / len(finite), 3) if finite else None,
        "min": round(min(finite), 3) if finite else None,
        "p50": percentile(finite, 0.50),
        "p90": percentile(finite, 0.90),
        "p95": percentile(finite, 0.95),
        "max": round(max(finite), 3) if finite else None,
    }


def find_page_image(page_dir: Path, page: int) -> Path | None:
    candidates = [
        page_dir / f"page-{page}.png",
        page_dir / f"page-{page:02d}.png",
        page_dir / f"page-{page:06d}.png",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    matches = sorted(page_dir.glob(f"*-{page}.png"))
    return matches[0] if matches else None


def is_ink(pixel: tuple[int, ...], white_threshold: int, alpha_threshold: int) -> bool:
    if len(pixel) == 4 and pixel[3] <= alpha_threshold:
        return False
    rgb = pixel[:3]
    return any(channel < white_threshold for channel in rgb)


def crop_rect_px(row: dict[str, Any], prefix: str, px_per_pt: float, padding_pt: float, image: Image.Image) -> tuple[int, int, int, int]:
    x = float(row[f"{prefix}XPt"])
    y = float(row[f"{prefix}YPt"])
    width = float(row[f"{prefix}WidthPt"])
    height = float(row[f"{prefix}HeightPt"])
    x0 = max(0, int(math.floor((x - padding_pt) * px_per_pt)))
    y0 = max(0, int(math.floor((y - padding_pt) * px_per_pt)))
    x1 = min(image.width, int(math.ceil((x + width + padding_pt) * px_per_pt)))
    y1 = min(image.height, int(math.ceil((y + height + padding_pt) * px_per_pt)))
    return x0, y0, max(x0 + 1, x1), max(y0 + 1, y1)


def measure_rendered_ink(
    page_dir: Path,
    row: dict[str, Any],
    prefix: str,
    dpi: int,
    padding_pt: float,
    white_threshold: int,
    alpha_threshold: int,
) -> dict[str, Any]:
    page = int(row[f"{prefix}Page"])
    page_path = find_page_image(page_dir, page)
    if page_path is None:
        return {"error": "missing_page", "page": page}

    px_per_pt = dpi / 72.0
    with Image.open(page_path).convert("RGBA") as image:
        crop_box = crop_rect_px(row, prefix, px_per_pt, padding_pt, image)
        crop = image.crop(crop_box)
        pixels = crop.load()
        min_x = crop.width
        min_y = crop.height
        max_x = -1
        max_y = -1
        ink_count = 0
        for y in range(crop.height):
            for x in range(crop.width):
                if is_ink(pixels[x, y], white_threshold, alpha_threshold):
                    ink_count += 1
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)

    if max_x < 0:
        return {
            "error": "no_ink",
            "page": page,
            "page_image": str(page_path),
            "crop_bbox_px": list(crop_box),
            "ink_count": 0,
        }

    abs_x0 = crop_box[0] + min_x
    abs_y0 = crop_box[1] + min_y
    abs_x1 = crop_box[0] + max_x + 1
    abs_y1 = crop_box[1] + max_y + 1
    ink_x_pt = abs_x0 / px_per_pt
    ink_y_pt = abs_y0 / px_per_pt
    ink_width_pt = (abs_x1 - abs_x0) / px_per_pt
    ink_height_pt = (abs_y1 - abs_y0) / px_per_pt
    return {
        "page": page,
        "page_image": str(page_path),
        "crop_bbox_px": list(crop_box),
        "ink_count": ink_count,
        "ink_bbox_px": [abs_x0, abs_y0, abs_x1, abs_y1],
        "ink_x_pt": round(ink_x_pt, 3),
        "ink_y_pt": round(ink_y_pt, 3),
        "ink_width_pt": round(ink_width_pt, 3),
        "ink_height_pt": round(ink_height_pt, 3),
        "ink_center_x_pt": round(ink_x_pt + ink_width_pt / 2.0, 3),
        "ink_center_y_pt": round(ink_y_pt + ink_height_pt / 2.0, 3),
    }


def parse_index_list(value: str) -> set[int]:
    if not value.strip():
        return set()
    indexes: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        indexes.add(int(part))
    return indexes


def draw_box(draw: ImageDraw.ImageDraw, bbox: list[int] | tuple[int, int, int, int], color: str, width: int) -> None:
    x0, y0, x1, y1 = [int(value) for value in bbox]
    for offset in range(width):
        draw.rectangle([x0 - offset, y0 - offset, x1 + offset, y1 + offset], outline=color)


def write_debug_images(
    rows: list[dict[str, Any]],
    comparison: dict[str, Any],
    reference_pages: Path,
    generated_pages: Path,
    debug_dir: Path,
    indexes: set[int],
    dpi: int,
) -> list[dict[str, Any]]:
    if not indexes:
        return []
    debug_dir.mkdir(parents=True, exist_ok=True)
    rows_by_index = {int(row["Index"]): row for row in rows}
    measured_by_index = {int(row["index"]): row for row in comparison["rows"]}
    missing_by_index = {int(row["index"]): row for row in comparison["missing"]}
    written: list[dict[str, Any]] = []

    for index in sorted(indexes):
        layout_row = rows_by_index.get(index)
        if layout_row is None:
            continue
        comparison_row = measured_by_index.get(index) or missing_by_index.get(index) or {}
        for prefix, page_dir, side_name in (
            ("Reference", reference_pages, "reference"),
            ("Generated", generated_pages, "generated"),
        ):
            measurement = comparison_row.get(side_name, {})
            page = int(layout_row[f"{prefix}Page"])
            page_path = find_page_image(page_dir, page)
            if page_path is None:
                continue
            with Image.open(page_path).convert("RGBA") as page_image:
                page_overlay = page_image.copy()
                draw = ImageDraw.Draw(page_overlay)
                px_per_pt = dpi / 72.0
                object_box = [
                    int(round(float(layout_row[f"{prefix}XPt"]) * px_per_pt)),
                    int(round(float(layout_row[f"{prefix}YPt"]) * px_per_pt)),
                    int(round((float(layout_row[f"{prefix}XPt"]) + float(layout_row[f"{prefix}WidthPt"])) * px_per_pt)),
                    int(round((float(layout_row[f"{prefix}YPt"]) + float(layout_row[f"{prefix}HeightPt"])) * px_per_pt)),
                ]
                crop_box = measurement.get("crop_bbox_px")
                if crop_box:
                    draw_box(draw, crop_box, "#ff9900", 2)
                draw_box(draw, object_box, "#ff0000", 3)
                ink_box = measurement.get("ink_bbox_px")
                if ink_box:
                    draw_box(draw, ink_box, "#0066ff", 3)

                page_out = debug_dir / f"formula-{index:03d}-{side_name}-page-{page:02d}.png"
                page_overlay.save(page_out)

                crop_source = crop_box or object_box
                pad = 20
                crop_with_pad = (
                    max(0, int(crop_source[0]) - pad),
                    max(0, int(crop_source[1]) - pad),
                    min(page_image.width, int(crop_source[2]) + pad),
                    min(page_image.height, int(crop_source[3]) + pad),
                )
                crop_image = page_overlay.crop(crop_with_pad)
                crop_out = debug_dir / f"formula-{index:03d}-{side_name}-crop.png"
                crop_image.save(crop_out)

            written.append(
                {
                    "index": index,
                    "side": side_name,
                    "page": page,
                    "page_overlay": str(page_out),
                    "crop_overlay": str(crop_out),
                }
            )
    return written


def compare_rows(
    rows: list[dict[str, Any]],
    reference_pages: Path,
    generated_pages: Path,
    dpi: int,
    padding_pt: float,
    white_threshold: int,
    alpha_threshold: int,
) -> dict[str, Any]:
    measured: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    skipped_empty_pair: list[dict[str, Any]] = []
    for row in rows:
        ref = measure_rendered_ink(reference_pages, row, "Reference", dpi, padding_pt, white_threshold, alpha_threshold)
        gen = measure_rendered_ink(generated_pages, row, "Generated", dpi, padding_pt, white_threshold, alpha_threshold)
        if ref.get("error") == "no_ink" and gen.get("error") == "no_ink":
            skipped_empty_pair.append({"index": row["Index"], "reference": ref, "generated": gen})
            continue
        if ref.get("error") or gen.get("error"):
            missing.append({"index": row["Index"], "reference": ref, "generated": gen})
            continue

        width_delta = float(gen["ink_width_pt"]) - float(ref["ink_width_pt"])
        height_delta = float(gen["ink_height_pt"]) - float(ref["ink_height_pt"])
        x_delta = float(gen["ink_x_pt"]) - float(ref["ink_x_pt"])
        y_delta = float(gen["ink_y_pt"]) - float(ref["ink_y_pt"])
        center_x_delta = float(gen["ink_center_x_pt"]) - float(ref["ink_center_x_pt"])
        center_y_delta = float(gen["ink_center_y_pt"]) - float(ref["ink_center_y_pt"])
        center_distance = math.hypot(center_x_delta, center_y_delta)
        measured.append(
            {
                "index": row["Index"],
                "reference_page": row["ReferencePage"],
                "generated_page": row["GeneratedPage"],
                "reference": ref,
                "generated": gen,
                "ink_width_delta_pt": round(width_delta, 3),
                "ink_height_delta_pt": round(height_delta, 3),
                "ink_x_delta_pt": round(x_delta, 3),
                "ink_y_delta_pt": round(y_delta, 3),
                "ink_center_x_delta_pt": round(center_x_delta, 3),
                "ink_center_y_delta_pt": round(center_y_delta, 3),
                "ink_center_distance_pt": round(center_distance, 3),
                "ink_width_abs_delta_pt": round(abs(width_delta), 3),
                "ink_height_abs_delta_pt": round(abs(height_delta), 3),
                "ink_x_abs_delta_pt": round(abs(x_delta), 3),
                "ink_y_abs_delta_pt": round(abs(y_delta), 3),
                "ink_center_abs_distance_pt": round(center_distance, 3),
                "ink_width_scale": round(float(gen["ink_width_pt"]) / max(float(ref["ink_width_pt"]), 0.001), 4),
                "ink_height_scale": round(float(gen["ink_height_pt"]) / max(float(ref["ink_height_pt"]), 0.001), 4),
            }
        )

    return {
        "measured_count": len(measured),
        "missing_count": len(missing),
        "skipped_empty_pair_count": len(skipped_empty_pair),
        "missing": missing[:50],
        "skipped_empty_pair": skipped_empty_pair[:50],
        "ink_width_abs_delta_pt": stats([row["ink_width_abs_delta_pt"] for row in measured]),
        "ink_height_abs_delta_pt": stats([row["ink_height_abs_delta_pt"] for row in measured]),
        "ink_center_distance_pt": stats([row["ink_center_distance_pt"] for row in measured]),
        "ink_width_scale": stats([row["ink_width_scale"] for row in measured]),
        "ink_height_scale": stats([row["ink_height_scale"] for row in measured]),
        "worst_width": sorted(measured, key=lambda item: item["ink_width_abs_delta_pt"], reverse=True)[:30],
        "worst_height": sorted(measured, key=lambda item: item["ink_height_abs_delta_pt"], reverse=True)[:30],
        "worst_center": sorted(measured, key=lambda item: item["ink_center_distance_pt"], reverse=True)[:30],
        "rows": measured,
    }


def fail_reasons(comparison: dict[str, Any], args: argparse.Namespace) -> list[str]:
    reasons: list[str] = []
    if comparison["missing_count"] > args.max_missing:
        reasons.append(f"missing_count={comparison['missing_count']} > {args.max_missing}")
    width_p90 = comparison["ink_width_abs_delta_pt"]["p90"]
    height_p90 = comparison["ink_height_abs_delta_pt"]["p90"]
    center_p90 = comparison["ink_center_distance_pt"]["p90"]
    if width_p90 is not None and width_p90 > args.max_p90_width_delta_pt:
        reasons.append(f"width_p90={width_p90}pt > {args.max_p90_width_delta_pt}pt")
    if height_p90 is not None and height_p90 > args.max_p90_height_delta_pt:
        reasons.append(f"height_p90={height_p90}pt > {args.max_p90_height_delta_pt}pt")
    if center_p90 is not None and center_p90 > args.max_p90_center_delta_pt:
        reasons.append(f"center_p90={center_p90}pt > {args.max_p90_center_delta_pt}pt")
    return reasons


def render_text(report: dict[str, Any]) -> str:
    comparison = report["comparison"]
    lines = [
        "Rendered formula ink comparison",
        "",
        f"layout: {report['layout_json']}",
        f"reference_pages: {report['reference_pages']}",
        f"generated_pages: {report['generated_pages']}",
        f"measured_count: {comparison['measured_count']}",
        f"missing_count: {comparison['missing_count']}",
        f"skipped_empty_pair_count: {comparison['skipped_empty_pair_count']}",
        "",
        f"ink_width_abs_delta_pt: {comparison['ink_width_abs_delta_pt']}",
        f"ink_height_abs_delta_pt: {comparison['ink_height_abs_delta_pt']}",
        f"ink_center_distance_pt: {comparison['ink_center_distance_pt']}",
        f"ink_width_scale: {comparison['ink_width_scale']}",
        f"ink_height_scale: {comparison['ink_height_scale']}",
        "",
        "Worst visible width deltas",
    ]
    for row in comparison["worst_width"][:10]:
        lines.append(
            "  #{index}: ref={rw}pt gen={gw}pt delta={delta}pt page={rp}->{gp}".format(
                index=row["index"],
                rw=row["reference"]["ink_width_pt"],
                gw=row["generated"]["ink_width_pt"],
                delta=row["ink_width_delta_pt"],
                rp=row["reference_page"],
                gp=row["generated_page"],
            )
        )
    lines.extend(["", "Worst visible height deltas"])
    for row in comparison["worst_height"][:10]:
        lines.append(
            "  #{index}: ref={rh}pt gen={gh}pt delta={delta}pt page={rp}->{gp}".format(
                index=row["index"],
                rh=row["reference"]["ink_height_pt"],
                gh=row["generated"]["ink_height_pt"],
                delta=row["ink_height_delta_pt"],
                rp=row["reference_page"],
                gp=row["generated_page"],
            )
        )
    lines.extend(["", "Worst visible center deltas"])
    for row in comparison["worst_center"][:10]:
        lines.append(
            "  #{index}: distance={dist}pt dx={dx}pt dy={dy}pt page={rp}->{gp}".format(
                index=row["index"],
                dist=row["ink_center_distance_pt"],
                dx=row["ink_center_x_delta_pt"],
                dy=row["ink_center_y_delta_pt"],
                rp=row["reference_page"],
                gp=row["generated_page"],
            )
        )
    if report["fail_reasons"]:
        lines.extend(["", "Fail reasons"])
        lines.extend(f"  {reason}" for reason in report["fail_reasons"])
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--layout-json", type=Path, default=DEFAULT_LAYOUT)
    parser.add_argument("--reference-pages", type=Path, default=DEFAULT_REFERENCE_PAGES)
    parser.add_argument("--generated-pages", type=Path, default=DEFAULT_GENERATED_PAGES)
    parser.add_argument("--out-json", type=Path, default=DEFAULT_OUT_JSON)
    parser.add_argument("--out-text", type=Path, default=DEFAULT_OUT_TEXT)
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--padding-pt", type=float, default=1.0)
    parser.add_argument("--white-threshold", type=int, default=245)
    parser.add_argument("--alpha-threshold", type=int, default=8)
    parser.add_argument("--max-missing", type=int, default=0)
    parser.add_argument("--max-p90-width-delta-pt", type=float, default=3.0)
    parser.add_argument("--max-p90-height-delta-pt", type=float, default=2.5)
    parser.add_argument("--max-p90-center-delta-pt", type=float, default=8.0)
    parser.add_argument("--debug-indices", default="")
    parser.add_argument("--debug-dir", type=Path, default=Path("target/reference-roundtrip/rendered-formula-ink-debug"))
    parser.add_argument("--fail", action="store_true")
    args = parser.parse_args()

    layout = json.loads(args.layout_json.read_text(encoding="utf-8-sig"))
    rows = layout.get("Rows") or []
    comparison = compare_rows(
        rows,
        args.reference_pages,
        args.generated_pages,
        args.dpi,
        args.padding_pt,
        args.white_threshold,
        args.alpha_threshold,
    )
    reasons = fail_reasons(comparison, args)
    debug_outputs = write_debug_images(
        rows,
        comparison,
        args.reference_pages,
        args.generated_pages,
        args.debug_dir,
        parse_index_list(args.debug_indices),
        args.dpi,
    )
    report = {
        "layout_json": str(args.layout_json.resolve()),
        "reference_pages": str(args.reference_pages.resolve()),
        "generated_pages": str(args.generated_pages.resolve()),
        "dpi": args.dpi,
        "padding_pt": args.padding_pt,
        "thresholds": {
            "max_missing": args.max_missing,
            "max_p90_width_delta_pt": args.max_p90_width_delta_pt,
            "max_p90_height_delta_pt": args.max_p90_height_delta_pt,
            "max_p90_center_delta_pt": args.max_p90_center_delta_pt,
        },
        "comparison": comparison,
        "debug_outputs": debug_outputs,
        "fail_reasons": reasons,
    }
    args.out_json.parent.mkdir(parents=True, exist_ok=True)
    args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    args.out_text.write_text(render_text(report), encoding="utf-8")
    print(render_text(report))
    print(f"wrote {args.out_json}")
    print(f"wrote {args.out_text}")

    if args.fail and reasons:
        raise SystemExit("rendered formula ink acceptance failed: " + "; ".join(reasons))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
