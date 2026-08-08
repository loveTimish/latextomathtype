# -*- coding: utf-8 -*-
"""Build labeled contact sheets for Batik-reference versus vector-WMF pairs."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


PAGE_WIDTH = 3000
PAGE_HEIGHT = 2200
ITEMS_PER_PAGE = 8
MARGIN = 32
GAP = 10
LABEL_HEIGHT = 88


def load_font(size: int) -> ImageFont.ImageFont:
    for candidate in (Path("C:/Windows/Fonts/consola.ttf"), Path("C:/Windows/Fonts/arial.ttf")):
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def wrap(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, width: int) -> list[str]:
    lines: list[str] = []
    remaining = text
    while remaining:
        low, high = 1, len(remaining)
        while low < high:
            middle = (low + high + 1) // 2
            if draw.textlength(remaining[:middle], font=font) <= width:
                low = middle
            else:
                high = middle - 1
        lines.append(remaining[:low])
        remaining = remaining[low:]
    return lines or [""]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus", type=Path, required=True)
    parser.add_argument("--pairs", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()

    corpus = json.loads(args.corpus.read_text(encoding="utf-8"))
    examples: list[str] = []
    seen: set[str] = set()
    for entry in corpus["entries"]:
        for example in entry.get("examples", []):
            if example not in seen:
                seen.add(example)
                examples.append(example)
    args.out.mkdir(parents=True, exist_ok=True)
    label_font = load_font(25)
    id_font = load_font(29)
    row_height = (PAGE_HEIGHT - MARGIN * 2 - GAP * (ITEMS_PER_PAGE - 1)) // ITEMS_PER_PAGE
    pages = []
    failures = []
    for page_index, start in enumerate(range(0, len(examples), ITEMS_PER_PAGE), 1):
        page = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), "white")
        draw = ImageDraw.Draw(page)
        ids = []
        for row, latex in enumerate(examples[start:start + ITEMS_PER_PAGE]):
            index = start + row + 1
            item_id = f"official-{index:03d}"
            pair_path = args.pairs / f"{item_id}.png"
            if not pair_path.exists():
                failures.append({"id": item_id, "reason": "missing pair PNG"})
                continue
            y = MARGIN + row * (row_height + GAP)
            draw.rectangle((MARGIN, y, PAGE_WIDTH - MARGIN, y + row_height),
                           fill="#fafafa", outline="#888888", width=2)
            draw.text((MARGIN + 12, y + 8), item_id, fill="#111111", font=id_font)
            latex_x = MARGIN + 245
            for line_index, line in enumerate(wrap(
                    draw, latex, label_font, PAGE_WIDTH - latex_x - MARGIN - 12)[:2]):
                draw.text((latex_x, y + 8 + line_index * 30), line, fill="#222222", font=label_font)
            preview = Image.open(pair_path).convert("RGB")
            available_width = PAGE_WIDTH - MARGIN * 2 - 24
            available_height = row_height - LABEL_HEIGHT - 10
            preview.thumbnail((available_width, available_height), Image.Resampling.LANCZOS)
            x = MARGIN + 12
            py = y + LABEL_HEIGHT + max(0, (available_height - preview.height) // 2)
            page.paste(preview, (x, py))
            draw.text((PAGE_WIDTH - MARGIN - 390, y + row_height - 30),
                      "BATIK REFERENCE  |  VECTOR WMF", fill="#555555", font=label_font)
            ids.append(item_id)
        page_path = args.out / f"page-{page_index:03d}.png"
        page.save(page_path, optimize=True)
        pages.append({"page": page_index, "file": page_path.name, "ids": ids})

    index = {
        "schemaVersion": 1,
        "itemCount": len(examples),
        "pageCount": len(pages),
        "itemsPerPage": ITEMS_PER_PAGE,
        "layout": "Batik reference on left, independently rasterized vector WMF on right",
        "failures": failures,
        "pages": pages,
    }
    (args.out / "index.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"items={len(examples)} pages={len(pages)} failures={len(failures)}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
