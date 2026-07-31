# -*- coding: utf-8 -*-
"""
Extract preview images from a generated corpus docx for visual audit.

The renderer's OLE previews are WMF files holding a single high-resolution
STRETCHDIB record (0xF43). This script unpacks the DIB, rebuilds a BMP in
memory, and writes PNGs next to an index so previews can be eyeballed or
diffed against ground truth.

STRETCHDIB param layout observed in renderer output: rop(4 bytes) +
9 words (18 bytes): ?, srcH, srcW, srcY, srcX, dstH, dstW, dstY, dstX --
the DIB starts 22 bytes after the params. We locate the BITMAPINFOHEADER
by scanning for a plausible header-size word (40 or 12) instead of
hard-coding the offset.

Usage:
    python scripts/extract_preview_png.py target/corpus-preview.docx \
        --out target/preview-png [--only image10,image244] [--limit 30]
"""
from __future__ import annotations

import argparse
import io
import re
import struct
import zipfile
from pathlib import Path

from PIL import Image


def extract_dib(wmf: bytes) -> bytes | None:
    """Return DIB bytes (BITMAPINFOHEADER + palette + bits) or None."""
    off = 22 + 18  # placeable header + standard header
    while off + 6 <= len(wmf):
        rsz, fn = struct.unpack_from("<IH", wmf, off)
        if rsz < 3 or off + rsz * 2 > len(wmf):
            break
        if fn == 0xF43:  # META_STRETCHDIB
            p = off + 6
            end = off + rsz * 2
            # scan the first 64 param bytes for the BITMAPINFOHEADER size
            for start in range(p, min(p + 64, end - 4), 2):
                hsz = struct.unpack_from("<I", wmf, start)[0]
                if hsz in (12, 40, 52, 56, 108, 124) and start + hsz <= end:
                    return wmf[start:end]
        off += rsz * 2
    return None


def dib_to_image(dib: bytes) -> Image.Image:
    hsz = struct.unpack_from("<I", dib, 0)[0]
    if hsz >= 40:
        bpp = struct.unpack_from("<H", dib, 14)[0]
        clr_used = struct.unpack_from("<I", dib, 32)[0]
        ncolors = clr_used or (1 << bpp if bpp <= 8 else 0)
        palette = ncolors * 4
    else:
        bpp = struct.unpack_from("<H", dib, 10)[0]
        palette = (1 << bpp if bpp <= 8 else 0) * 3
    bits_off = 14 + hsz + palette
    bmp = b"BM" + struct.pack("<IHHI", 14 + len(dib), 0, 0, bits_off) + dib
    return Image.open(io.BytesIO(bmp))


def paragraph_labels(document_xml: str) -> list[str]:
    """image-name labels in paragraph order (label run text 'imageNN: ')."""
    return re.findall(r"<w:t[^>]*>(image\d+): </w:t>", document_xml)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--out", required=True)
    ap.add_argument("--only", default="")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    z = zipfile.ZipFile(args.docx)
    doc = z.read("word/document.xml").decode("utf-8")
    labels = paragraph_labels(doc)
    media = sorted((n for n in z.namelist()
                    if re.match(r"word/media/image_eq\d+\.wmf$", n)),
                   key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    only = set(args.only.split(",")) if args.only else None

    rows = []
    count = 0
    for i, name in enumerate(media):
        label = labels[i] if i < len(labels) else f"eq{i + 1}"
        if only and label not in only:
            continue
        if args.limit and count >= args.limit:
            break
        dib = extract_dib(z.read(name))
        if dib is None:
            rows.append((label, "NO-DIB", 0, 0))
            continue
        im = dib_to_image(dib)
        png = out / f"{label}.png"
        im.save(png)
        rows.append((label, png.name, im.width, im.height))
        count += 1
    for r in rows:
        print("\t".join(str(x) for x in r))
    print(f"extracted={count} -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
