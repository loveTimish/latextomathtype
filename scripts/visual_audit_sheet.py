# -*- coding: utf-8 -*-
"""
One-command visual audit sheet: for each corpus formula, render the genuine
MathType WMF (GT, top strip) and unpack our renderer's OLE preview (bottom
strip), normalize both to the same px/pt using each WMF's own placeable
bbox, and stack them into a collage PNG for eyeballing.

Usage:
    python scripts/visual_audit_sheet.py \
        --gen target/corpus-preview.docx \
        --gt-docx rebuild-assets/external/fraction-split-reference.docx \
        --out target/audit-sheet.png [--only image10,image244] [--limit 20] [--offset 0]
"""
from __future__ import annotations

import argparse
import io
import re
import struct
import sys
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from wmf_glyph_layout import parse_wmf  # noqa: E402
from render_gt_wmf_png import render as render_gt  # noqa: E402
from extract_preview_png import extract_dib, dib_to_image, paragraph_labels  # noqa: E402

PX_PER_PT = 12.0
STRIP_GAP = 4
LABEL_W = 90


def placeable_bbox_pt(wmf: bytes):
    if len(wmf) < 22:
        return None
    left, top, right, bottom, inch = struct.unpack_from("<hhhhH", wmf, 0)
    if inch == 0:
        return None
    s = 72.0 / inch
    return [left * s, top * s, right * s, bottom * s]


def render_gt_at_scale(wmf: bytes) -> Image.Image | None:
    parsed = parse_wmf(wmf)
    return render_gt(parsed, PX_PER_PT)


def preview_at_scale(wmf: bytes) -> Image.Image | None:
    bbox = placeable_bbox_pt(wmf)
    dib = extract_dib(wmf)
    if dib is None:
        return None
    im = dib_to_image(dib)
    if not bbox:
        return im
    pt_w = bbox[2] - bbox[0]
    pt_h = bbox[3] - bbox[1]
    if pt_w <= 0 or pt_h <= 0:
        return im
    return im.resize((max(int(pt_w * PX_PER_PT), 1),
                      max(int(pt_h * PX_PER_PT), 1)))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--gen", required=True)
    ap.add_argument("--gt-docx", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--only", default="")
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--offset", type=int, default=0)
    args = ap.parse_args()

    gen = zipfile.ZipFile(args.gen)
    labels = paragraph_labels(gen.read("word/document.xml").decode("utf-8"))
    media = {int(re.search(r"(\d+)", n).group(1)): n
             for n in gen.namelist()
             if re.match(r"word/media/image_eq\d+\.wmf$", n)}
    gt = zipfile.ZipFile(args.gt_docx)
    gt_map = {Path(n).stem: n for n in gt.namelist()
              if n.startswith("word/media/") and n.lower().endswith(".wmf")}

    only = set(args.only.split(",")) if args.only else None
    items = []
    for i, label in enumerate(labels):
        if only and label not in only:
            continue
        items.append((label, i + 1))
    items = items[args.offset: args.offset + args.limit]

    strips = []
    for label, idx in items:
        gt_name = gt_map.get(label)
        pv_name = media.get(idx)
        if not gt_name or not pv_name:
            continue
        gt_im = render_gt_at_scale(gt.read(gt_name))
        pv_im = preview_at_scale(gen.read(pv_name))
        if gt_im is None or pv_im is None:
            continue
        w = max(gt_im.width, pv_im.width) + LABEL_W
        h = gt_im.height + pv_im.height + STRIP_GAP
        strip = Image.new("RGB", (w, h), "white")
        strip.paste(gt_im, (LABEL_W, 0))
        strip.paste(pv_im, (LABEL_W, gt_im.height + STRIP_GAP))
        d = ImageDraw.Draw(strip)
        d.text((4, 4), label, fill="black")
        strips.append(strip)

    if not strips:
        print("nothing to audit")
        return 1
    cols = 2 if len(strips) > 1 else 1
    col_w = max(s.width for s in strips)
    rows = (len(strips) + cols - 1) // cols
    row_h = [0] * rows
    for i, s in enumerate(strips):
        row_h[i % rows] = max(row_h[i % rows], s.height)
    sheet = Image.new("RGB", (col_w * cols, sum(row_h) + STRIP_GAP * rows),
                      "#999999")
    for i, s in enumerate(strips):
        c, r = i % cols, i // cols
        y = sum(row_h[:r]) + STRIP_GAP * r
        sheet.paste(s, (c * col_w, y))
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    sheet.save(args.out)
    print(f"strips={len(strips)} sheet={sheet.size} -> {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
