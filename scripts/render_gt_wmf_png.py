# -*- coding: utf-8 -*-
"""
Rasterize genuine MathType WMF previews (vector: glyph origins + rules) to
PNG with PIL and the real system TTFs, for side-by-side visual audit against
our renderer's OLE preview output.

Coordinate model matches wmf_glyph_layout.parse_wmf: glyph x/y are baseline
origins in pt (y down), rules are lines/rects/polygons in pt.

Usage:
    python scripts/render_gt_wmf_png.py --docx rebuild-assets/external/fraction-split-reference.docx \
        --out target/gt-png --only image10,image244 [--scale 8]
"""
from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

TTF = {
    ("times new roman", False): "C:/Windows/Fonts/times.ttf",
    ("times new roman", True): "C:/Windows/Fonts/timesi.ttf",
    ("symbol", False): "C:/Windows/Fonts/symbol.ttf",
}
FALLBACK = "C:/Windows/Fonts/times.ttf"
# DejaVu (bundled with the managed Python's matplotlib) covers chars TNR
# lacks, e.g. MT Extra '⋯'
DEJAVU = str(Path(__file__).resolve().parents[1] / ".venv" /
             "Lib" / "site-packages" / "matplotlib" / "mpl-data" / "fonts" /
             "ttf" / "DejaVuSans.ttf")
if not Path(DEJAVU).exists():
    import matplotlib
    DEJAVU = str(Path(matplotlib.__file__).parent / "mpl-data" / "fonts" /
                 "ttf" / "DejaVuSans.ttf")


class FontCache:
    def __init__(self, scale):
        self.scale = scale
        self.cache = {}

    def _load(self, path, size_pt):
        px = max(int(round(size_pt * self.scale)), 4)
        ck = (path, px)
        if ck not in self.cache:
            self.cache[ck] = ImageFont.truetype(path, px)
        return self.cache[ck]

    def get(self, face, italic, size_pt):
        key = ((face or "").lower(), bool(italic))
        path = TTF.get(key) or TTF.get((key[0], False)) or FALLBACK
        return self._load(path, size_pt)

    def fallback(self, size_pt):
        return self._load(DEJAVU if Path(DEJAVU).exists() else FALLBACK,
                          size_pt)


def has_glyph(font, ch):
    try:
        from fontTools.ttLib import TTFont
        tt = TTFont(font.path, lazy=True, fontNumber=font.index)
        cmap = tt.getBestCmap() or {}
        return ord(ch) in cmap
    except Exception:
        return True  # assume present; worst case a .notdef box


def render(parsed: dict, scale: float = 8.0) -> Image.Image | None:
    bbox = parsed.get("placeable_bbox_pt")
    if not bbox:
        return None
    left, top, right, bottom = bbox
    w = max(int((right - left) * scale) + 8, 10)
    h = max(int((bottom - top) * scale) + 8, 10)
    im = Image.new("RGB", (w, h), "white")
    dr = ImageDraw.Draw(im)
    fonts = FontCache(scale)

    def sx(x_pt):
        return (x_pt - left) * scale + 4

    def sy(y_pt):
        return (y_pt - top) * scale + 4

    for r in parsed.get("rules", []):
        if r["kind"] == "line" and "x1_pt" in r:
            width = max(int(0.46 * scale), 1)
            dr.line([sx(r["x1_pt"]), sy(r["y1_pt"]),
                     sx(r["x2_pt"]), sy(r["y2_pt"])], fill="black", width=width)
        elif r["kind"] == "rect" and "left_pt" in r:
            dr.rectangle([sx(r["left_pt"]), sy(r["top_pt"]),
                          sx(r["right_pt"]), sy(r["bottom_pt"])], fill="black")
        elif r.get("points_pt"):
            pts = [(sx(px), sy(py)) for px, py in r["points_pt"]]
            if r["kind"] == "polygon":
                dr.polygon(pts, fill="black")
            else:
                dr.line(pts, fill="black", width=max(int(0.46 * scale), 1))
    for g in parsed.get("glyphs", []):
        if "x_pt" not in g:
            continue
        size = g.get("size_pt") or 10.5
        font = fonts.get(g.get("font"), g.get("italic"), size)
        ch = g["ch"]
        # Symbol/MT Extra have no usable Unicode charmap via freetype;
        # their operator/Greek shapes in TNR are visually near-identical,
        # so audit images render decoded chars with TNR instead.
        face = (g.get("font") or "").lower()
        if face != "times new roman":
            font = fonts.get("Times New Roman", False, size)
        if not has_glyph(font, ch):
            font = fonts.fallback(size)
        # y is the baseline origin; PIL anchors at ascender top by default
        ascent, _ = font.getmetrics()
        dr.text((sx(g["x_pt"]), sy(g["y_pt"]) - ascent), ch,
                fill="black", font=font)
    return im


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--only", default="")
    ap.add_argument("--scale", type=float, default=8.0)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    z = zipfile.ZipFile(args.docx)
    wmfs = sorted(n for n in z.namelist()
                  if n.startswith("word/media/") and n.lower().endswith(".wmf"))
    only = set(args.only.split(",")) if args.only else None
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from wmf_glyph_layout import parse_wmf

    count = 0
    for name in wmfs:
        stem = Path(name).stem
        if only and stem not in only:
            continue
        if args.limit and count >= args.limit:
            break
        im = render(parse_wmf(z.read(name)), args.scale)
        if im is None:
            continue
        im.save(out / f"{stem}.png")
        count += 1
    print(f"rendered={count} -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
