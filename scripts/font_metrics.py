# -*- coding: utf-8 -*-
"""
TrueType advance-width provider for WMF glyph geometry.

MathType WMF gives glyph ORIGINS and origin-to-origin advances; the final
glyph of every EXTTEXTOUT run has a dx that is the advance to the next run's
anchor, not the glyph width. To close the right edge of a glyph box we need
real font advance widths, which are a pure font property (no train/val
leakage). This module reads them from the system TTFs that MathType used:
Times New Roman (regular/italic/bold) and Symbol.

- TNR faces: keyed by unicode codepoint via the best cmap.
- Symbol: keyed by the original font-position byte (glyph 'hex' field in
  wmf_glyph_layout output); Microsoft-symbol cmaps map 0xF000+byte.
- Unknown faces (MT Extra, CJK): return None; callers fall back to a
  corpus/default ratio.
"""
from __future__ import annotations

from pathlib import Path

from fontTools.ttLib import TTFont

FONT_DIR = Path("C:/Windows/Fonts")

_FACE_FILES = {
    ("times new roman", False): "times.ttf",
    ("times new roman", True): "timesi.ttf",
    ("symbol", False): "symbol.ttf",
}


class FontMetrics:
    def __init__(self, font_dir: Path = FONT_DIR):
        self._fonts: dict[str, TTFont] = {}
        self._cmaps: dict[str, dict] = {}
        self._upem: dict[str, int] = {}
        self._hmtx: dict[str, dict] = {}
        self._font_dir = font_dir

    def _load(self, fname: str) -> bool:
        if fname in self._fonts:
            return True
        path = self._font_dir / fname
        if not path.exists():
            return False
        font = TTFont(str(path), lazy=True)
        self._fonts[fname] = font
        cmap = font.getBestCmap()
        if cmap is None and fname == "symbol.ttf":
            # Symbol has no Unicode cmap fontTools accepts: merge the
            # MacRoman table (raw-byte keys) and the Microsoft-symbol
            # table (0xF000+byte keys) into byte -> glyph name.
            cmap = {}
            ms = font["cmap"].getcmap(3, 0)
            if ms is not None:
                for cp, name in ms.cmap.items():
                    if 0xF000 <= cp <= 0xF0FF:
                        cmap[cp - 0xF000] = name
            mac = font["cmap"].getcmap(1, 0)
            if mac is not None:
                for b, name in mac.cmap.items():
                    cmap.setdefault(b, name)
        self._cmaps[fname] = cmap or {}
        self._upem[fname] = font["head"].unitsPerEm
        self._hmtx[fname] = font["hmtx"]
        return True

    def advance_ratio(self, ch: str, hex_byte: str | None = None,
                      face: str | None = None, italic: bool = False
                      ) -> float | None:
        """Advance width / font size for a glyph, or None if unknown."""
        key = ((face or "").lower(), bool(italic))
        fname = _FACE_FILES.get(key) or _FACE_FILES.get((key[0], False))
        if fname is None or not self._load(fname):
            return None
        cmap = self._cmaps[fname]
        glyph_name = None
        if fname == "symbol.ttf":
            # merged byte -> glyph-name map (see _load)
            if hex_byte:
                glyph_name = cmap.get(int(hex_byte, 16))
        else:
            if ch:
                glyph_name = cmap.get(ord(ch))
        if glyph_name is None:
            return None
        adv = self._hmtx[fname].metrics.get(glyph_name)
        if adv is None:
            return None
        return adv[0] / self._upem[fname]
