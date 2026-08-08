# -*- coding: utf-8 -*-
"""Build labeled PNG contact sheets from the exact official OLE WMF previews."""

from __future__ import annotations

import argparse
import json
import struct
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from extract_preview_png import dib_to_image


PAGE_WIDTH = 3000
PAGE_HEIGHT = 2200
ITEMS_PER_PAGE = 8
MARGIN = 36
ROW_GAP = 12
LABEL_HEIGHT = 96


def font(size: int) -> ImageFont.ImageFont:
    candidates = (
        Path("C:/Windows/Fonts/consola.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
    )
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def fit_preview(image: Image.Image, width: int, height: int) -> Image.Image:
    rgb = image.convert("RGB")
    rgb.thumbnail((width, height), Image.Resampling.LANCZOS)
    return rgb


def strict_stretchdib(wmf: bytes) -> bytes | None:
    """Read the renderer's fixed META_STRETCHDIB payload without header guessing."""
    offset = 22 + 18
    while offset + 6 <= len(wmf):
        record_words, function = struct.unpack_from("<IH", wmf, offset)
        record_end = offset + record_words * 2
        if record_words < 3 or record_end > len(wmf):
            return None
        if function == 0x0F43:
            dib_start = offset + 6 + 22
            if dib_start + 40 > record_end:
                return None
            header_size, width, height, planes, bit_count = struct.unpack_from("<IiiHH", wmf, dib_start)
            if header_size != 40 or width <= 0 or height == 0 or planes != 1 or bit_count != 24:
                return None
            row_stride = ((width * 3 + 3) // 4) * 4
            required = header_size + row_stride * abs(height)
            return wmf[dib_start:dib_start + required] if dib_start + required <= record_end else None
        offset = record_end
    return None


def wrap_pixels(draw: ImageDraw.ImageDraw, text: str, max_width: int,
                label_font: ImageFont.ImageFont) -> list[str]:
    """Wrap every source character; never truncate the LaTeX under review."""
    lines: list[str] = []
    remaining = text
    while remaining:
        low, high = 1, len(remaining)
        while low < high:
            middle = (low + high + 1) // 2
            if draw.textlength(remaining[:middle], font=label_font) <= max_width:
                low = middle
            else:
                high = middle - 1
        lines.append(remaining[:low])
        remaining = remaining[low:]
    return lines or [""]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--decisions", default="")
    args = parser.parse_args()

    manifest_path = Path(args.manifest).resolve()
    output = Path(args.out).resolve()
    pages_dir = output / "pages"
    previews_dir = output / "previews"
    pages_dir.mkdir(parents=True, exist_ok=True)
    previews_dir.mkdir(parents=True, exist_ok=True)

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    items = manifest["items"]
    label_font = font(28)
    id_font = font(32)
    row_height = (PAGE_HEIGHT - MARGIN * 2 - ROW_GAP * (ITEMS_PER_PAGE - 1)) // ITEMS_PER_PAGE
    rendered = []
    failures = []

    for item in items:
        wmf_path = manifest_path.parent / item["wmf"]
        dib = strict_stretchdib(wmf_path.read_bytes())
        if dib is None:
            failures.append({"id": item["id"], "latex": item["latex"], "reason": "NO_STRETCHDIB"})
            continue
        preview = dib_to_image(dib).convert("RGB")
        preview_path = previews_dir / f"{item['id']}.png"
        preview.save(preview_path)
        rendered.append({**item, "preview": str(preview_path.relative_to(output)).replace("\\", "/")})

    page_records = []
    for page_index, start in enumerate(range(0, len(rendered), ITEMS_PER_PAGE), 1):
        page_items = rendered[start:start + ITEMS_PER_PAGE]
        page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
        draw = ImageDraw.Draw(page)
        ids = []
        for row, item in enumerate(page_items):
            y = MARGIN + row * (row_height + ROW_GAP)
            draw.rounded_rectangle(
                (MARGIN, y, PAGE_WIDTH - MARGIN, y + row_height),
                radius=8,
                outline="#8a8a8a",
                width=2,
                fill="#fafafa",
            )
            draw.text((MARGIN + 16, y + 12), item["id"], fill="#111111", font=id_font)
            latex_x = MARGIN + 270
            for line_index, line in enumerate(wrap_pixels(draw, item["latex"], PAGE_WIDTH - latex_x - MARGIN - 20, label_font)):
                draw.text((latex_x, y + 12 + line_index * 34), line, fill="#222222", font=label_font)
            preview = Image.open(output / item["preview"])
            fitted = fit_preview(preview, PAGE_WIDTH - MARGIN * 2 - 40, row_height - LABEL_HEIGHT - 16)
            preview_x = MARGIN + 20
            preview_y = y + LABEL_HEIGHT + max((row_height - LABEL_HEIGHT - fitted.height) // 2, 0)
            page.paste(fitted, (preview_x, preview_y))
            ids.append(item["id"])
        page_path = pages_dir / f"page-{page_index:03d}.png"
        page.save(page_path, optimize=True)
        page_records.append({
            "page": page_index,
            "file": str(page_path.relative_to(output)).replace("\\", "/"),
            "ids": ids,
        })

    index = {
        "schemaVersion": 1,
        "sourceManifest": str(manifest_path),
        "payload": manifest.get("previewPayload"),
        "itemCount": len(items),
        "renderedCount": len(rendered),
        "pageCount": len(page_records),
        "itemsPerPage": ITEMS_PER_PAGE,
        "failures": failures,
        "pages": page_records,
        "items": rendered,
    }
    (output / "contact-sheet-index.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if args.decisions:
        decisions_path = Path(args.decisions).resolve()
        decisions = json.loads(decisions_path.read_text(encoding="utf-8"))
        issues = decisions.get("issues", [])
        issue_ids = {issue["id"] for issue in issues}
        known_ids = {item["id"] for item in rendered}
        unknown_ids = sorted(issue_ids - known_ids)
        if unknown_ids:
            raise ValueError(f"Visual decisions reference unknown ids: {unknown_ids}")
        audit_report = {
            "schemaVersion": 1,
            "auditMethod": "manual page-by-page review of exact WMF STRETCHDIB previews",
            "reviewer": decisions.get("reviewer", "human-visual-review"),
            "sourceManifest": str(manifest_path),
            "contactSheetIndex": str((output / "contact-sheet-index.json").resolve()),
            "payload": manifest.get("previewPayload"),
            "reviewedCount": len(rendered),
            "passedCount": len(rendered) - len(issue_ids),
            "failedCount": len(issue_ids),
            "verdict": "pass" if not issues and not failures else "fail",
            "issueCategoryCounts": {
                category: sum(1 for issue in issues if issue["category"] == category)
                for category in sorted({issue["category"] for issue in issues})
            },
            "issues": [
                {
                    **issue,
                    "page": (int(issue["id"].rsplit("-", 1)[1]) - 1) // ITEMS_PER_PAGE + 1,
                    "latex": next(item["latex"] for item in rendered if item["id"] == issue["id"]),
                    "preview": next(item["preview"] for item in rendered if item["id"] == issue["id"]),
                }
                for issue in issues
            ],
        }
        (output / "manual-visual-audit-report.json").write_text(
            json.dumps(audit_report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(f"rendered={len(rendered)} pages={len(page_records)} failures={len(failures)} -> {output}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
