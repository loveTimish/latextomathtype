# -*- coding: utf-8 -*-
"""Build full-document PaperExportRequest JSON files from docx2tex .tex files.

Unlike make_batch10_requests.py, this keeps the whole converted text instead of
sampling formulas only.  Paragraph text is preserved, LaTeX math delimiters are
kept for MathType/OLE generation, and beginPic/endPic placeholders for PNG/JPEG
assets become question image paths.
"""

from __future__ import annotations

import json
import re
import argparse
import struct
import zipfile
from html import unescape
from collections import defaultdict, deque
from pathlib import Path
import xml.etree.ElementTree as ET

from scan_docx_latex_leaks import validate_mathtype_ole


ROOT = Path(r"J:\latextomathtype\analysis")
DEFAULT_LATEX_ROOT = ROOT / "batch10-latex"
DEFAULT_FALLBACK_LATEX_ROOT = ROOT / "batch10-latex"
DEFAULT_OUT_DIR = ROOT / "batch10-full-requests"
STYLE_HINTS_ROOT = ROOT / "mtef-style-hints"
SAFE_STYLE_HINTS = {
    "asciiFlatParens",
    "fullwidthTextParen",
    "forceExplicitFenceTemplate",
    "flatParenTemplate",
    "letterGroupObarTemplate",
    "textFeComma",
    "explicitFractionFullSize",
    "mixedAsciiFullwidthParens",
    "legacyTextFeParenContent",
}
PIC_RE = re.compile(r"beginPic\{([^}]+)\}endPic")
INCLUDEGRAPHICS_RE = re.compile(r"\\includegraphics(?:\[[^\]]*])?\{([^}]+)\}")
INCLUDEGRAPHICS_LOOSE_RE = re.compile(
    r"\\includegraphics(?:\[[^\]]*])?\s+"
    r"([^\s]+?\.(?:png|jpe?g|gif|bmp|emf|wmf|bin)(?:\?[^\s$]*)?|[^\s$]+)",
    re.IGNORECASE,
)
MATH_SPAN_RE = re.compile(r"(\$\$.*?\$\$|\$.*?\$)", re.S)
METRICS_RE = re.compile(r"^\\pwmetrics\{[^}]+}\s*")
STYLE_RE = re.compile(r"^\\pwstyle\{[^}]*}\s*")
DASHLINE_RE = re.compile(r"\\(?:hline|hdashline)(?:\[[^\]]*])?")
TAG_RE = re.compile(r"\\tag\s*\{[^{}]*}")
EMPTY_METRIC_MATH_RE = re.compile(
    r"\$\\pwmetrics\{([^}]+)}\$(\s*(?:\\(?:textcolor|textbf|textit|textnormal|textrm|mathrm|textsc)"
    r"(?:\{[^{}$]*})?\{[^{}$]*})*)\s*\$"
)
EMPTY_METRIC_BEFORE_TEXT_COMMAND_RE = re.compile(
    r"\$\\pwmetrics\{([^}]+)}\$(\s*(?:\\(?:textcolor|textbf|textit|textnormal|textrm|mathrm|textsc)"
    r"(?:\{[^{}$]*})?\{(?:[^{}$]|\{[^{}$]*})*})+)\s*\$"
)
EMPTY_METRIC_BEFORE_BARE_FORMULA_RE = re.compile(
    r"\$\\pwmetrics\{([^}]+)}\$\s*([^$\n，。；;]*?(?:\\(?:frac|times|div)|[×÷=])[^$\n，。；;]*)"
)
METRIC_ARRAY_RE = re.compile(
    r"(\$\$|\$)\\pwmetrics\{([^}]+)}\s*\\begin\{array\}\{[^}]+}(.*?)\\end\{array\}\s*\1",
    re.S,
)
OLE_OBJECT_RE = re.compile(r"oleObject(\d+)\.bin")
PLACEABLE_WMF_KEY = 0x9AC6CDD7
EMPTY_EQUATION_PLACEHOLDER = r"\square"
SOURCE_FORMULA_INSERTIONS = {
    27: {
        1: {
            # docx2tex misses source oleObject1, but the bundled WMF preview is valid.
            "latex": "路程=速度×时间",
            "ole": "embeddings/oleObject1.bin",
            "wmf": "media/image3.wmf",
        },
    },
    29: {
        49: {
            # docx2tex misses source oleObject49, but the bundled WMF preview is valid.
            "latex": r"2时10\frac{10}{11}分",
            "ole": "embeddings/oleObject49.bin",
            "wmf": "media/image51.wmf",
        },
    },
    23: {
        475: {
            # docx2tex misses source oleObject475, but the bundled WMF preview is valid.
            "latex": r"\begin{cases}x+y+z=100①\\5x+3y+\frac{1}{3}z=100②\end{cases}",
            "ole": "embeddings/oleObject475.bin",
            "wmf": "media/image457.wmf",
        },
    },
    49: {
        10: {
            # docx2tex leaves this MathType object as includegraphics; restore it at its source ordinal.
            "latex": r"4=1\times 4=2\times 2",
            "ole": "embeddings/oleObject10.bin",
            "wmf": "media/image20.wmf",
        },
        139: {
            # docx2tex leaves this MathType object as includegraphics; restore it at its source ordinal.
            "latex": r"a>b",
            "ole": "embeddings/oleObject139.bin",
            "wmf": "media/image296.wmf",
        },
    },
    53: {
        78: {
            # Source embedding is not a readable CFB OLE, but the paired WMF preview is a formula.
            "latex": "n",
            "ole": "embeddings/oleObject78.bin",
            "wmf": "media/image106.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-cfb-preview-only",
        },
        79: {
            # Source embedding is not a readable CFB OLE, but the paired WMF preview is a formula.
            "latex": "n^{2}",
            "ole": "embeddings/oleObject79.bin",
            "wmf": "media/image107.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-cfb-preview-only",
        },
        148: {
            # docx2tex drops these final repeated MathType fractions.
            "latex": r"\frac{27}{40}",
            "ole": "embeddings/oleObject148.bin",
            "wmf": "media/image199.wmf",
            "disabled": True,
            "excludedReason": "duplicate-source-object-superseded-by-oleObject150",
        },
        149: {
            "latex": r"\frac{27}{40}",
            "ole": "embeddings/oleObject149.bin",
            "wmf": "media/image200.wmf",
            "disabled": True,
            "excludedReason": "duplicate-source-object-superseded-by-oleObject150",
        },
        150: {
            "latex": r"\frac{27}{40}",
            "ole": "embeddings/oleObject150.bin",
            "wmf": "media/image201.wmf",
            "injectAfterOle": "embeddings/oleObject149.bin",
        },
    },
    54: {
        8: {
            "latex": "AB",
            "ole": "embeddings/oleObject8.bin",
            "wmf": "media/image10.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-mathtype-preview-label",
        },
        9: {
            "latex": "CD",
            "ole": "embeddings/oleObject9.bin",
            "wmf": "media/image11.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-mathtype-preview-label",
        },
    },
    55: {
        8: {
            "latex": "AB",
            "ole": "embeddings/oleObject8.bin",
            "wmf": "media/image11.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-mathtype-preview-label",
        },
        9: {
            "latex": "CD",
            "ole": "embeddings/oleObject9.bin",
            "wmf": "media/image12.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-mathtype-preview-label",
        },
    },
    56: {
        8: {
            "latex": "AB",
            "ole": "embeddings/oleObject8.bin",
            "wmf": "media/image15.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-mathtype-preview-label",
        },
        9: {
            "latex": "CD",
            "ole": "embeddings/oleObject9.bin",
            "wmf": "media/image16.wmf",
            "allowPreviewBacked": True,
            "metricsOnly": True,
            "excludedReason": "non-mathtype-preview-label",
        },
    },
}
EMPTY_FRAC_DEN_RE = re.compile(
    r"\\frac\s*\{\s*([^{}]+?)\s*\}\s*\{\s*\}\s*"
    r"((?:[A-Za-z0-9]+)(?:\s*\\times\s*[A-Za-z0-9]+)*)"
)
TRAILING_RIGHT_RE = re.compile(r"\\left\[\s*([^{}\\]+?)\s*\\right\s*$")
LEFT_ZERO_PERCENT_RE = re.compile(r"\\left\s+0\s+(.+?)\s+\\right\s+\\%\[\]")
MATHOP_SIMPLE_RE = re.compile(r"\\mathop\s*\{\s*([^{}\\]+?)\s*\}")
MATHOP_MATHRM_CHAR_RE = re.compile(r"\\mathop\s*\{\s*\\mathrm\s*\{\s*([^{}\\]+?)\s*\}\s*\}")
ACCENT_MATHRM_CHAR_RE = re.compile(
    r"(\\(?:dot|hat|bar|vec|tilde))\s+\\mathrm\s*\{\s*([^{}\\]+?)\s*\}"
)
MATHRM_GREEK_RE = re.compile(r"\\mathrm\s*\{\s*([πΠ])\s*\}")
MATHRM_COMMAND_RE = re.compile(r"\\mathrm\s*\{\s*((?:\\(?:lt|gt|le|ge|leq|geq|neq|ne|times|div))|[=+\-*/])\s*\}")
COMMAND_BEFORE_CJK_RE = re.compile(r"(\\[A-Za-z]+)(?=[\u3400-\u4DBF\u4E00-\u9FFF])")
OVERSET_ARC_RE = re.compile(r"\\overset\s*(?:\{?\s*⌢\s*\}?|\{?\s*\\frown\s*\}?)\s*\{\s*([^{}]+?)\s*\}")
OVERSET_RIGHTARROW_RE = re.compile(r"\\overset\s*\{\s*([^{}]+?)\s*\}\s*\{\s*\\rightarrow\s*\}")
UNDER_RIGHTARROW_RE = re.compile(r"\\underrightarrow\s*\{\s*([^{}]+?)\s*\}")
GREEK_COMMANDS = {
    "π": r"\pi",
    "Π": r"\Pi",
}
DING_CIRCLED = {
    "172": "①",
    "173": "②",
    "174": "③",
    "175": "④",
    "176": "⑤",
    "177": "⑥",
    "178": "⑦",
    "179": "⑧",
    "180": "⑨",
    "181": "⑩",
}


def report_equations(index: int, latex_root: Path) -> list[dict]:
    report_path = latex_root / str(index) / f"{index}.report.json"
    if not report_path.exists():
        return []
    report = json.loads(report_path.read_text(encoding="utf-8"))
    equations = []
    for item in report.get("equations", []):
        if item.get("status") != "converted":
            continue
        if item.get("sourceRepairReason") == "empty-output":
            item = dict(item)
            item["output"] = ""
            equations.append(item)
        elif item.get("output"):
            equations.append(item)
    return equations


def metric_prefix(item: dict) -> str:
    metrics = item.get("metrics") or {}
    wmf_width = metrics.get("wmfWidthPt") or metrics.get("shapeWidthPt") or metrics.get("dxaOrigPt")
    wmf_height = metrics.get("wmfHeightPt") or metrics.get("shapeHeightPt") or metrics.get("dyaOrigPt")
    shape_width = metrics.get("shapeWidthPt") or wmf_width
    shape_height = metrics.get("shapeHeightPt") or wmf_height
    if not wmf_width or not wmf_height:
        return ""
    if shape_width and shape_height:
        return (
            f"\\pwmetrics{{{float(wmf_width):.3f},{float(wmf_height):.3f},"
            f"{float(shape_width):.3f},{float(shape_height):.3f}}}"
        )
    return f"\\pwmetrics{{{float(wmf_width):.3f},{float(wmf_height):.3f}}}"


def parse_wmf_phys(data: bytes) -> tuple[float, float] | None:
    if len(data) < 22 or struct.unpack("<I", data[:4])[0] != PLACEABLE_WMF_KEY:
        return None
    left, top, right, bottom, inch = struct.unpack("<hhhhH", data[6:16])
    if inch <= 0:
        return None
    return ((right - left) / inch * 72.0, (bottom - top) / inch * 72.0)


def rel_map(zip_file: zipfile.ZipFile, rels_name: str) -> dict[str, str]:
    if rels_name not in zip_file.namelist():
        return {}
    rels = zip_file.read(rels_name).decode("utf-8", "replace")
    return dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))


def read_target(zip_file: zipfile.ZipFile, target: str) -> bytes | None:
    name = target.lstrip("/")
    if not name.startswith("word/"):
        name = "word/" + name
    if name in zip_file.namelist():
        return zip_file.read(name)
    return None


def source_docx_metrics(source_docx: Path | None) -> deque[dict]:
    if not source_docx or not source_docx.exists():
        return deque()
    metrics = deque()
    with zipfile.ZipFile(source_docx) as z:
        doc = z.read("word/document.xml").decode("utf-8", "replace")
        rels = rel_map(z, "word/_rels/document.xml.rels")
        for match in re.finditer(r"<w:object\b([^>]*)>(.*?)</w:object>", doc, re.S):
            attrs, body = match.group(1), match.group(2)
            wmf_width = wmf_height = None
            shape_width = shape_height = None
            image_rid = re.search(r'<v:imagedata [^>]*r:id="([^"]+)"', body)
            image_target = rels.get(image_rid.group(1), "") if image_rid else ""
            target = ole_target(body, rels)
            ole_data = read_target(z, target) if target else None
            try:
                is_mathtype, mathtype_error = validate_mathtype_ole(ole_data or b"")
            except Exception as exc:
                is_mathtype, mathtype_error = False, str(exc)
            style = re.search(r'<v:shape\b[^>]*\sstyle="([^"]*)"', body)
            if style:
                width_match = re.search(r"width:([\d.]+)pt", style.group(1))
                height_match = re.search(r"height:([\d.]+)pt", style.group(1))
                shape_width = float(width_match.group(1)) if width_match else None
                shape_height = float(height_match.group(1)) if height_match else None
            if image_target.lower().endswith(".wmf"):
                phys = parse_wmf_phys(read_target(z, image_target) or b"")
                if phys:
                    wmf_width, wmf_height = phys
            if wmf_width is None or wmf_height is None:
                dxa = re.search(r'w:dxaOrig="(\d+)"', attrs)
                dya = re.search(r'w:dyaOrig="(\d+)"', attrs)
                if dxa and dya:
                    wmf_width, wmf_height = int(dxa.group(1)) / 20.0, int(dya.group(1)) / 20.0
            if shape_width is None or shape_height is None:
                shape_width = wmf_width
                shape_height = wmf_height
            if wmf_width and wmf_height:
                metrics.append({
                    "docObjectIndex": len(metrics) + 1,
                    "metrics": {
                        "wmfWidthPt": wmf_width,
                        "wmfHeightPt": wmf_height,
                        "shapeWidthPt": shape_width,
                        "shapeHeightPt": shape_height,
                        "dxaOrigPt": int(re.search(r'w:dxaOrig="(\d+)"', attrs).group(1)) / 20.0 if re.search(r'w:dxaOrig="(\d+)"', attrs) else None,
                        "dyaOrigPt": int(re.search(r'w:dyaOrig="(\d+)"', attrs).group(1)) / 20.0 if re.search(r'w:dyaOrig="(\d+)"', attrs) else None,
                    },
                    "oleTarget": target,
                    "wmfTarget": image_target,
                    "isMathType": is_mathtype,
                    "mathTypeError": mathtype_error,
                })
            else:
                metrics.append({
                    "docObjectIndex": len(metrics) + 1,
                    "metrics": {
                        "shapeWidthPt": shape_width,
                        "shapeHeightPt": shape_height,
                    },
                    "missingMetrics": True,
                    "oleTarget": target,
                    "wmfTarget": image_target,
                    "isMathType": is_mathtype,
                    "mathTypeError": mathtype_error,
                })
    return metrics


def ole_target(body: str, rels: dict[str, str]) -> str:
    ole_rid = re.search(r'<o:OLEObject [^>]*r:id="([^"]+)"', body)
    return rels.get(ole_rid.group(1), "") if ole_rid else ""


def mml2tex_sequence(xml_path: Path) -> list[str]:
    if not xml_path.exists():
        return []
    text = xml_path.read_text(encoding="utf-8", errors="replace")
    values = []
    for match in re.finditer(r"<\?(mml2tex|d2t)\s+(.+?)\?>", text, re.S):
        target = match.group(1)
        data = unescape(match.group(2)).strip()
        if target == "d2t" and "empty equation object" in data:
            values.append(EMPTY_EQUATION_PLACEHOLDER)
            continue
        if target != "mml2tex":
            continue
        latex = re.sub(r"\s+", " ", data)
        if latex:
            values.append(latex)
    return values


def mml2tex_sequence_for_tex(tex_path: Path) -> list[str]:
    debug_root = tex_path.parent / f"{tex_path.stem}.debug"
    candidates = [
        debug_root / "xml2tex" / "20.mml2tex.xml",
        debug_root / "mml2tex" / "10.mml2tex-main.xml",
        tex_path.with_suffix(".xml"),
    ]
    for path in candidates:
        values = mml2tex_sequence(path)
        if values:
            return values
    return []


def mml2tex_xml_path_for_tex(tex_path: Path) -> Path | None:
    debug_root = tex_path.parent / f"{tex_path.stem}.debug"
    candidates = [
        debug_root / "xml2tex" / "20.mml2tex.xml",
        debug_root / "mml2tex" / "10.mml2tex-main.xml",
        tex_path.with_suffix(".xml"),
    ]
    return next((path for path in candidates if path.exists()), None)


def source_docx_equations(tex_path: Path, source_docx: Path | None) -> list[dict]:
    formulas = mml2tex_sequence_for_tex(tex_path)
    all_metrics = list(source_docx_metrics(source_docx))
    insertions = SOURCE_FORMULA_INSERTIONS.get(int(tex_path.stem), {}) if tex_path.stem.isdigit() else {}
    metrics_only = {
        ordinal
        for ordinal, insertion in insertions.items()
        if insertion.get("metricsOnly") or insertion.get("disabled")
    }
    preview_backed = {
        ordinal
        for ordinal, insertion in insertions.items()
        if insertion.get("allowPreviewBacked") and not insertion.get("metricsOnly") and not insertion.get("disabled")
    }
    metrics = [
        item for item in all_metrics
        if item.get("isMathType") or item.get("docObjectIndex") in preview_backed
    ]
    if metrics_only and len(formulas) == len(all_metrics):
        formulas = [
            formula for formula, metric in zip(formulas, all_metrics)
            if metric.get("docObjectIndex") not in metrics_only
        ]
    if not formulas or not metrics:
        return []
    formulas = apply_source_formula_insertions(tex_path, formulas, metrics)
    count = len(formulas)
    if count > len(metrics):
        raise ValueError(f"formula count exceeds source metrics for {tex_path}: {count} > {len(metrics)}")
    return [
        {
            "output": formulas[i],
            "metrics": metrics[i].get("metrics") or {},
            "oleTarget": metrics[i].get("oleTarget") or "",
            "wmfTarget": metrics[i].get("wmfTarget") or "",
            "source": metrics[i].get("oleTarget") or "",
            "docObjectIndex": metrics[i].get("docObjectIndex"),
            "wmfTarget": metrics[i].get("wmfTarget") or "",
        }
        for i in range(count)
    ]


def apply_source_formula_insertions(tex_path: Path, formulas: list[str], metrics: list[dict]) -> list[str]:
    insertions = SOURCE_FORMULA_INSERTIONS.get(int(tex_path.stem)) if tex_path.stem.isdigit() else None
    if not insertions:
        return formulas
    text_insertions = {
        ordinal: insertion
        for ordinal, insertion in insertions.items()
        if not insertion.get("metricsOnly") and not insertion.get("disabled")
    }
    expected_count = len(metrics)
    if len(formulas) + len(text_insertions) != expected_count:
        raise ValueError(
            f"{tex_path}: source formula insertion expected {expected_count} metrics and "
            f"{len(formulas)} extracted formulas plus {len(text_insertions)} insertions"
        )
    restored = list(formulas)
    for ordinal, insertion in sorted(text_insertions.items()):
        expected_ole = insertion.get("ole")
        expected_wmf = insertion.get("wmf")
        position = ordinal - 1
        if expected_ole:
            position = next(
                (
                    idx for idx, metric in enumerate(metrics)
                    if metric.get("oleTarget") == expected_ole
                ),
                -1,
            )
        if position < 0 or position > len(restored):
            raise ValueError(f"{tex_path}: invalid source formula insertion ordinal {ordinal}")
        metric = metrics[position] if position < len(metrics) else {}
        if expected_ole and metric.get("oleTarget") != expected_ole:
            raise ValueError(
                f"{tex_path}: source formula insertion ordinal {ordinal} expected OLE "
                f"{expected_ole}, got {metric.get('oleTarget')}"
            )
        if expected_wmf and metric.get("wmfTarget") != expected_wmf:
            raise ValueError(
                f"{tex_path}: source formula insertion ordinal {ordinal} expected WMF "
                f"{expected_wmf}, got {metric.get('wmfTarget')}"
            )
        restored.insert(position, insertion["latex"])
    return restored


def blocks_from_mml2tex_xml(
    tex_path: Path,
    equations: list[dict],
    styles: dict[int, str] | None = None,
) -> tuple[list[str], list[dict]]:
    xml_path = mml2tex_xml_path_for_tex(tex_path)
    if not xml_path or not equations:
        return [], []
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_pis=True))
    root = ET.parse(xml_path, parser=parser).getroot()
    equation_cursor = EquationCursor(equations)
    blocks: list[str] = []
    parent_map = {child: parent for parent in root.iter() for child in list(parent)}
    for node in root:
        if node.tag is ET.ProcessingInstruction:
            text = xml_pi_text_with_formula(node, equation_cursor, styles or {})
        elif local_name(node.tag) == "equation":
            text = xml_equation_node_text(node, equation_cursor, styles or {})
        elif local_name(node.tag) == "para":
            text = xml_node_text_with_formulas(node, equation_cursor, styles or {})
        elif local_name(node.tag) in {"sidebar", "variablelist", "orderedlist", "informaltable"}:
            text = xml_node_text_with_formulas(node, equation_cursor, styles or {})
        else:
            continue
        text = re.sub(r"\s{2,}", " ", text).strip()
        if text:
            blocks.append(text)
    inject_source_only_formulas(tex_path, blocks, equations, styles or {})
    return blocks, equation_cursor.multi_consumes


def inject_source_only_formulas(tex_path: Path, blocks: list[str], equations: list[dict], styles: dict[int, str]) -> None:
    if not tex_path.stem.isdigit():
        return
    insertions = SOURCE_FORMULA_INSERTIONS.get(int(tex_path.stem)) or {}
    for insertion in insertions.values():
        if not insertion.get("injectAfterOle") or insertion.get("disabled"):
            continue
        formula = next(
            (item for item in equations if item.get("oleTarget") == insertion.get("ole")),
            None,
        )
        after = next(
            (item for item in equations if item.get("oleTarget") == insertion.get("injectAfterOle")),
            None,
        )
        if not formula or not after:
            raise ValueError(f"{tex_path}: invalid source-only formula injection target {insertion}")
        after_text = normalize_latex_key(after.get("output", ""))
        target_block = -1
        for index, block in enumerate(blocks):
            if any(normalize_latex_key(body) == after_text for body in math_bodies(block)):
                target_block = index
        if target_block < 0:
            raise ValueError(f"{tex_path}: missing source-only formula injection anchor {after.get('oleTarget')}")
        blocks[target_block] = blocks[target_block] + latex_formula_text(formula, styles, "")


def xml_alignment_key(value: str) -> str:
    value = repair_docx2tex_latex(value or "")
    value = normalize_latex_key(value)
    value = value.replace("&", "")
    value = re.sub(r"\\(?:left|right)\s*\.", "", value)
    value = re.sub(r"\\(?:left|right)\s*", "", value)
    value = re.sub(r"\\(?:rm|mathrm)\s*", "", value)
    value = re.sub(r"\\(?![A-Za-z])", "", value)
    return re.sub(r"[\s{}\[\]()]+", "", value)


def is_strong_xml_hint(hint_key: str) -> bool:
    compact = "".join(ch for ch in hint_key if ch.isalnum() or ch in r"\{}_^")
    if len(compact) < 3:
        return False
    if re.fullmatch(r"[A-Z]{2,6}", compact):
        return False
    return compact not in {"A", "B", "C", "D", "E", "F", "G", "H", "a", "b", "c", "d", "x", "y", "n"}


class EquationCursor:
    def __init__(self, equations: list[dict]):
        self.equations = equations
        self.index = 0
        self.multi_consumes: list[dict] = []

    def consume_for_hint(self, hint: str) -> list[dict]:
        if self.index >= len(self.equations):
            return []
        hint_key = xml_alignment_key(hint)
        if not hint_key or not is_strong_xml_hint(hint_key):
            item = self.equations[self.index]
            self.index += 1
            return [item]
        window_end = min(len(self.equations), self.index + 8)
        for pos in range(self.index, window_end):
            if xml_alignment_key(self.equations[pos].get("output", "")) == hint_key:
                items = self.equations[self.index:pos + 1]
                if len(items) > 1:
                    self.multi_consumes.append(
                        {
                            "start": self.index + 1,
                            "end": pos + 1,
                            "count": len(items),
                            "hint": hint[:120],
                        }
                    )
                self.index = pos + 1
                return items
        item = self.equations[self.index]
        self.index += 1
        return [item]

    def consume_one(self) -> list[dict]:
        if self.index >= len(self.equations):
            return []
        item = self.equations[self.index]
        self.index += 1
        return [item]

    def consume_for_ole_target(self, target: str) -> list[dict]:
        if self.index >= len(self.equations):
            return []
        normalized = target.replace("\\", "/")
        if normalized.startswith("word/"):
            normalized = normalized[5:]
        expected = str(self.equations[self.index].get("oleTarget") or "").replace("\\", "/")
        if expected.startswith("word/"):
            expected = expected[5:]
        if expected != normalized:
            return []
        window_end = min(len(self.equations), self.index + 8)
        for pos in range(self.index, window_end):
            candidate = str(self.equations[pos].get("oleTarget") or "").replace("\\", "/")
            if candidate.startswith("word/"):
                candidate = candidate[5:]
            if candidate == normalized:
                items = self.equations[self.index:pos + 1]
                if len(items) > 1:
                    self.multi_consumes.append(
                        {
                            "start": self.index + 1,
                            "end": pos + 1,
                            "count": len(items),
                            "hint": target,
                        }
                    )
                self.index = pos + 1
                return items
        return self.consume_one()

def has_para_ancestor(node: ET.Element, parent_map: dict[ET.Element, ET.Element]) -> bool:
    parent = parent_map.get(node)
    while parent is not None:
        if local_name(parent.tag) == "para":
            return True
        parent = parent_map.get(parent)
    return False


def local_name(tag: str) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def latex_formula_text(formula: dict, styles: dict[int, str], fallback: str = "") -> str:
    latex = repair_docx2tex_latex((formula.get("output", "") or fallback).strip())
    prefix = metric_prefix(formula) + style_prefix(formula, styles)
    return "$" + prefix + latex + "$"


def xml_pi_text_with_formula(pi_node: ET.Element, equation_cursor: EquationCursor, styles: dict[int, str]) -> str:
    pi_text = pi_node.text or ""
    target, _, data = pi_text.partition(" ")
    if target == "d2t" and "empty equation object" in data:
        formulas = equation_cursor.consume_one()
        return "".join(latex_formula_text(formula, styles, "") for formula in formulas)
    if target != "mml2tex":
        return "<br/>" if target == "latex" and "\\\\" in data else ""
    formulas = equation_cursor.consume_for_hint(data)
    if not formulas:
        return ""
    return "".join(
        latex_formula_text(formula, styles, data if offset == len(formulas) - 1 else "")
        for offset, formula in enumerate(formulas)
    )


def xml_node_text_with_formulas(node: ET.Element, equation_cursor: EquationCursor, styles: dict[int, str]) -> str:
    parts: list[str] = []
    if node.text:
        parts.append(node.text)
    for child in list(node):
        if child.tag is ET.ProcessingInstruction:
            parts.append(xml_pi_text_with_formula(child, equation_cursor, styles))
        elif local_name(child.tag) in {"inlineequation", "equation"}:
            parts.append(xml_equation_node_text(child, equation_cursor, styles))
        elif local_name(child.tag) in {"mediaobject", "inlinemediaobject"}:
            if xml_node_role(child) == "OLEObject":
                parts.append(xml_ole_mediaobject_text(child, equation_cursor, styles))
            else:
                image = xml_mediaobject_path(child)
                if image:
                    parts.append(f" beginPic{{{image}}}endPic ")
        else:
            parts.append(xml_node_text_with_formulas(child, equation_cursor, styles))
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)


def xml_equation_node_text(node: ET.Element, equation_cursor: EquationCursor, styles: dict[int, str]) -> str:
    latex = ""
    for child in list(node):
        if child.tag is ET.ProcessingInstruction:
            target, _, data = (child.text or "").partition(" ")
            if not latex and target == "mml2tex" and data.strip():
                latex = data.strip()
                break
    formulas = equation_cursor.consume_for_hint(latex)
    if not formulas:
        return ""
    return "".join(
        latex_formula_text(formula, styles, latex if offset == len(formulas) - 1 else "")
        for offset, formula in enumerate(formulas)
    )


def xml_mediaobject_path(node: ET.Element) -> str | None:
    for elem in node.iter():
        if local_name(elem.tag) == "imagedata":
            fileref = elem.attrib.get("fileref")
            if fileref:
                return fileref
    return None


def xml_node_role(node: ET.Element) -> str:
    return str(node.attrib.get("role") or "")


def xml_ole_mediaobject_text(node: ET.Element, equation_cursor: EquationCursor, styles: dict[int, str]) -> str:
    target = xml_mediaobject_path(node) or ""
    formulas = equation_cursor.consume_for_ole_target(target)
    return "".join(latex_formula_text(formula, styles, "") for formula in formulas)


def style_hint_lookup(index: int, style_hints_root: Path) -> dict[int, str]:
    path = style_hints_root / f"{index}.style-hints.json"
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    lookup: dict[int, str] = {}
    for item in data:
        hints = item.get("hints") or []
        if not hints:
            continue
        object_index = int(item.get("objectIndex") or 0)
        if object_index > 0:
            lookup[object_index] = ",".join(hints)
    return lookup


def source_object_index(item: dict) -> int:
    match = OLE_OBJECT_RE.search(item.get("source") or "")
    return int(match.group(1)) if match else 0


def style_prefix(item: dict, styles: dict[int, str]) -> str:
    encoded = styles.get(source_object_index(item), "")
    latex = item.get("output") or ""
    if encoded:
        parts = [part for part in encoded.split(",") if part in SAFE_STYLE_HINTS]
        encoded = ",".join(parts)
    if "forceExplicitFenceTemplate" in encoded and "\\left" not in latex and "\\right" not in latex:
        parts = [part for part in encoded.split(",") if part != "forceExplicitFenceTemplate"]
        encoded = ",".join(parts)
    return f"\\pwstyle{{{encoded}}}" if encoded else ""


def normalize_latex_key(value: str) -> str:
    value = repair_docx2tex_latex(value or "")
    value = METRICS_RE.sub("", value).strip()
    value = STYLE_RE.sub("", value).strip()
    value = value.replace(r"\lt ", "<").replace(r"\gt ", ">")
    return re.sub(r"\s+", " ", value)


def equation_metric_lookup(equations: list[dict]) -> dict[str, deque[dict]]:
    lookup: dict[str, deque[dict]] = defaultdict(deque)
    for item in equations:
        lookup[normalize_latex_key(item.get("output", ""))].append(item)
    return lookup


def normalize_math_delimiters(text: str) -> str:
    # docx2tex can place an unresolved preview immediately before adjacent
    # MathType spans. Remove it before MATH_SPAN_RE sees the dollar run, or
    # ``image.emf$$...$$$...$`` is paired as one malformed formula.
    text = INCLUDEGRAPHICS_RE.sub("", text)
    text = INCLUDEGRAPHICS_LOOSE_RE.sub("", text)
    text = re.sub(
        r"\\\[(.+?)\\\]",
        lambda m: "$$" + unwrap_inline_math_inside_command_args(m.group(1).strip()) + "$$",
        text,
        flags=re.S,
    )
    text = re.sub(r"\\\((.+?)\\\)", lambda m: "$" + m.group(1).strip() + "$", text, flags=re.S)
    text = re.sub(
        r"\\begin\{equation\*?\}(.+?)\\end\{equation\*?\}",
        lambda m: "$$" + unwrap_inline_math_inside_command_args(m.group(1).strip()) + "$$",
        text,
        flags=re.S,
    )
    text = re.sub(
        r"\\begin\{align\*?\}(.+?)\\end\{align\*?\}",
        lambda m: "$$" + unwrap_inline_math_inside_command_args(normalize_align_environment_body(m.group(1))) + "$$",
        text,
        flags=re.S,
    )
    return text


def normalize_align_environment_body(body: str) -> str:
    lines = [line.strip() for line in re.split(r"\\\\", body) if line.strip()]
    lines = [line[1:].strip() if line.startswith("&") else line for line in lines]
    return r"\begin{array}{l} " + r"\\ ".join(lines) + r" \end{array}"


TEXT_FORMAT_COMMANDS = {
    "textcolor",
    "textbf",
    "textit",
    "textnormal",
    "textrm",
    "mathrm",
    "textsc",
    "emph",
    "uline",
}

WRAPPER_COMMANDS = {
    "fbox",
    "mbox",
    "makebox",
}


def find_balanced_group_end(text: str, start: int) -> int:
    if start >= len(text) or text[start] != "{":
        return -1
    depth = 0
    i = start
    while i < len(text):
        ch = text[i]
        if ch == "\\" and i + 1 < len(text):
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def skip_ws(text: str, index: int) -> int:
    while index < len(text) and text[index].isspace():
        index += 1
    return index


def unwrap_inline_math_inside_command_args(text: str) -> str:
    out: list[str] = []
    i = 0
    changed = False
    while i < len(text):
        if text[i] != "\\":
            out.append(text[i])
            i += 1
            continue
        j = i + 1
        while j < len(text) and text[j].isalpha():
            j += 1
        command = text[i + 1:j]
        if command not in TEXT_FORMAT_COMMANDS and command != "text":
            out.append(text[i])
            i += 1
            continue
        k = skip_ws(text, j)
        if command == "textcolor":
            color_end = find_balanced_group_end(text, k)
            if color_end < 0:
                out.append(text[i])
                i += 1
                continue
            k = skip_ws(text, color_end + 1)
        body_end = find_balanced_group_end(text, k)
        if body_end < 0:
            out.append(text[i])
            i += 1
            continue
        body = unwrap_inline_math_inside_command_args(text[k + 1:body_end])
        body = re.sub(r"\$\s*([^$]+?)\s*\$", lambda m: m.group(1).strip(), body)
        out.append(text[i:k + 1] + body + "}")
        i = body_end + 1
        changed = True
    return "".join(out)


def strip_latex_text_format_commands(text: str) -> str:
    """Drop LaTeX text styling commands while preserving their visible body."""
    out: list[str] = []
    i = 0
    changed = False
    while i < len(text):
        if text[i] != "\\":
            out.append(text[i])
            i += 1
            continue
        j = i + 1
        while j < len(text) and text[j].isalpha():
            j += 1
        command = text[i + 1:j]
        if command not in TEXT_FORMAT_COMMANDS:
            out.append(text[i])
            i += 1
            continue

        k = skip_ws(text, j)
        if command == "textcolor":
            color_start = k
            color_end = find_balanced_group_end(text, color_start)
            if color_end < 0:
                out.append(text[i])
                i += 1
                continue
            k = skip_ws(text, color_end + 1)

        body_start = k
        body_end = find_balanced_group_end(text, body_start)
        if body_end < 0:
            out.append(text[i])
            i += 1
            continue
        out.append(strip_latex_text_format_commands(text[body_start + 1:body_end]))
        i = body_end + 1
        changed = True
    result = "".join(out)
    return result if not changed else strip_latex_text_format_commands(result)


def strip_latex_wrapper_commands(text: str) -> str:
    """Drop layout wrapper commands while preserving their body."""
    out: list[str] = []
    i = 0
    changed = False
    while i < len(text):
        if text[i] != "\\":
            out.append(text[i])
            i += 1
            continue
        j = i + 1
        while j < len(text) and text[j].isalpha():
            j += 1
        command = text[i + 1:j]
        if command not in WRAPPER_COMMANDS:
            out.append(text[i])
            i += 1
            continue
        k = skip_ws(text, j)
        body_start = k
        body_end = find_balanced_group_end(text, body_start)
        if body_end < 0:
            out.append(text[i])
            i += 1
            continue
        out.append(strip_latex_wrapper_commands(text[body_start + 1:body_end]))
        i = body_end + 1
        changed = True
    result = "".join(out)
    return result if not changed else strip_latex_wrapper_commands(result)


def strip_minipage_environments(text: str) -> str:
    text = re.sub(r"\\begin\{minipage\}(?:\[[^\]]*])?\{[^{}]*}", "", text)
    text = re.sub(r"\\end\{minipage\}", "", text)
    return text


def strip_non_content_latex_commands_fragment(text: str) -> str:
    text = re.sub(
        r"\\begin\s+table\s+\\begin\s+tabularx\b.*?\\arraybackslash\s*",
        " ",
        text,
        flags=re.S,
    )
    text = re.sub(r"\\(?:begin|end)(?:\{(?:table|tabularx)\}|\s+(?:table|tabularx)\b)", " ", text)
    text = re.sub(
        r"\\(?:arraybackslash|textwidth|linewidth|tabcolsep|arrayrulewidth|dimexpr|"
        r"textsuperscript|textsubscript)\b",
        " ",
        text,
    )
    text = text.replace("&", " ")
    text = strip_minipage_environments(text)
    text = strip_latex_wrapper_commands(text)
    text = re.sub(r"\\(?:fbox|mbox|makebox)\b\s*", "", text)
    text = strip_latex_text_format_commands(text)
    text = re.sub(r"\\textcolor\s+[A-Za-z][A-Za-z0-9_-]*\s*", " ", text)
    text = re.sub(r"\\(?:textbf|textit|textnormal|textrm|mathrm|textsc|emph|uline)\b\s*", "", text)
    text = DASHLINE_RE.sub(" ", text)
    text = TAG_RE.sub(" ", text)
    text = re.sub(r"\\(?:begin|end)\{description\}(?:\[[^\]]*])?", " ", text)
    text = re.sub(r"\\(?:begin|end)\{(?:center|flushleft|flushright)\}", " ", text)
    text = text.replace(r"\centering", " ")
    text = text.replace(r"\newline", " ")
    text = text.replace(r"\textemdash{}", "-")
    text = text.replace(r"\ldots", "...")
    text = re.sub(r"\\par\b", " ", text)
    return re.sub(r"\s{2,}", " ", text).strip()


def strip_non_content_latex_commands(text: str) -> str:
    parts = MATH_SPAN_RE.split(text)
    for i in range(0, len(parts), 2):
        parts[i] = strip_non_content_latex_commands_fragment(parts[i])
    return re.sub(r"\s{2,}", " ", "".join(parts)).strip()


def strip_residual_metrics_outside_math(text: str) -> str:
    parts = MATH_SPAN_RE.split(text)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"\\pwmetrics\{[^}]+}", "", parts[i])
    return "".join(parts)


def strip_residual_group_braces_outside_math(text: str) -> str:
    parts = MATH_SPAN_RE.split(text)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"(?<!\\)[{}]", " ", parts[i])
    return re.sub(r"\s{2,}", " ", "".join(parts)).strip()


def merge_metric_only_formula_runs(text: str) -> str:
    text = re.sub(
        r"\$\\pwmetrics\{([^}]+)}\$\s*\$\{?\\div}?\$\s*([^$，。；;\n]*?)=\s*\$([^$]+)\$",
        lambda m: f"$\\pwmetrics{{{m.group(1)}}}\\div {m.group(2).strip()}={m.group(3).strip()}$",
        text,
    )
    text = re.sub(
        r"\$\\pwmetrics\{([^}]+)}\$\s*\$\{?\\times}?\$\s*([^$，。；;\n]*?)=\s*\$([^$]+)\$",
        lambda m: f"$\\pwmetrics{{{m.group(1)}}}\\times {m.group(2).strip()}={m.group(3).strip()}$",
        text,
    )
    return text


def demote_metadata_numeric_formula_runs(text: str) -> str:
    return re.sub(
        r"(【难度】)\s*\$\\pwmetrics\{[^}]+}([0-9]+)\$\s*(星)",
        r"\1\2 \3",
        text,
    )


def widen_simple_numeric_formula_metrics(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        metric_text = match.group(1)
        body = match.group(2).strip()
        width_pt, height_pt = split_metric_pair(metric_text)
        min_width_pt = estimate_formula_row_width_pt(body)
        if width_pt >= min_width_pt:
            return match.group(0)
        return f"$\\pwmetrics{{{min_width_pt:.3f},{height_pt:.3f}}}{body}$"

    return re.sub(
        r"\$\\pwmetrics\{([^}]+)}\s*([0-9]+(?:\.[0-9]+)?)\$",
        repl,
        text,
    )


def widen_too_small_formula_metrics(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        metric_text = match.group(1)
        body = match.group(2).strip()
        width_pt, height_pt = split_metric_pair(metric_text)
        min_width_pt = estimate_formula_row_width_pt(body)
        min_height_pt = 12.0
        if width_pt >= min_width_pt and height_pt >= min_height_pt:
            return match.group(0)
        return f"$\\pwmetrics{{{max(width_pt, min_width_pt):.3f},{max(height_pt, min_height_pt):.3f}}}{body}$"

    return re.sub(
        r"\$\\pwmetrics\{([^}]+)}\s*([^$]{8,})\$",
        repl,
        text,
    )


def add_estimated_formula_metrics(text: str) -> str:
    """Attach content-derived metrics when docx2tex report metrics are unavailable."""
    def repl(match: re.Match[str]) -> str:
        body = (match.group(1) or match.group(2) or "").strip()
        if not body:
            return match.group(0)
        if METRICS_RE.match(body):
            return match.group(0)
        clean_body = STYLE_RE.sub("", body).strip()
        if not should_keep_unmeasured_math_span(clean_body):
            return match.group(0)
        width_pt = estimate_formula_row_width_pt(clean_body)
        height_pt = estimate_formula_height_pt(clean_body)
        prefix = f"\\pwmetrics{{{width_pt:.3f},{height_pt:.3f}}}"
        if match.group(1) is not None:
            return "$$" + prefix + body + "$$"
        return "$" + prefix + body + "$"

    return re.sub(r"\$\$(.+?)\$\$|\$(.+?)\$", repl, text, flags=re.S)


def merge_inline_operator_math_fragments(text: str) -> str:
    """Merge docx2tex runs like 2.009${\times}$43 into one editable formula."""
    pattern = re.compile(
        r"(?<![A-Za-z])"
        r"([0-9][0-9A-Za-z.()（）+\-= ]*(?:\$\{?\\(?:times|div|cdot)}?\$[0-9A-Za-z.()（）+\-= ]*)+)"
    )

    def repl(match: re.Match[str]) -> str:
        fragment = match.group(1).strip()
        if not fragment:
            return match.group(0)
        latex = re.sub(r"\$\{?(\\(?:times|div|cdot))}?\$", r" \1 ", fragment)
        latex = re.sub(r"\s{2,}", " ", latex).strip()
        if not should_attach_metric_to_bare_formula(latex):
            return match.group(0)
        return "$" + latex + "$"

    return pattern.sub(repl, text)


def estimate_formula_height_pt(body: str) -> float:
    rows = max(
        1,
        len([row for row in re.split(r"\\\\", body) if row.strip()])
        if r"\begin{array}" in body or r"\begin{matrix}" in body
        else 1,
    )
    if rows > 1:
        return min(120.0, max(14.0, rows * 14.0))
    if r"\frac" in body or r"\sqrt" in body or "^" in body or "_" in body:
        return 16.0
    return 13.0


def strip_empty_bracket_artifacts(text: str) -> str:
    return re.sub(r"(?<!\\)\[\s*(?<!\\)\]", " ", text)


def separate_adjacent_display_formula_runs(text: str) -> str:
    """Break inline/display math runs that docx2tex collapsed onto one Word line."""
    text = re.sub(r"(?<!\$)\$(?=\$\$\\pwmetrics)", "$<br/>", text)
    text = re.sub(r"(?<!\$)\$(?=\s+\$\\pwmetrics)", "$<br/>", text)
    text = re.sub(r"(\$\$)\s+(?=【(?:答案|解析|提示|点评)】)", r"\1<br/>", text)
    return text


def split_display_array_formula_runs(text: str) -> str:
    """Restore multi-line MathType arrays as one editable equation per visual row."""
    def repl(match: re.Match[str]) -> str:
        metric_text = match.group(2)
        body = match.group(3).strip()
        width_pt, height_pt = split_metric_pair(metric_text)
        rows = []
        raw_rows = [row.strip() for row in re.split(r"\\\\", body) if row.strip()]
        row_height_pt = max(12.0, min(16.0, height_pt / max(len(raw_rows), 1)))
        for row in raw_rows:
            row = row.strip()
            row = re.sub(r"(^|[^\\])&", r"\1", row)
            row_width_pt = max(width_pt, estimate_formula_row_width_pt(row))
            rows.append(f"$\\pwmetrics{{{row_width_pt:.3f},{row_height_pt:.3f}}}" + row.strip() + "$")
        return "<br/>".join(rows) if rows else ""

    return METRIC_ARRAY_RE.sub(repl, text)


def split_metric_pair(metric_text: str) -> tuple[float, float]:
    parts = [part.strip() for part in metric_text.split(",", 1)]
    try:
        return float(parts[0]), float(parts[1])
    except (IndexError, ValueError):
        return 80.0, 13.0


def estimate_formula_row_width_pt(row: str) -> float:
    visible = re.sub(r"\\(?:times|div|left|right|cdot|,|;|!|enspace|quad)\b", "x", row)
    visible = re.sub(r"\\[A-Za-z]+", "x", visible)
    visible = re.sub(r"[{}]", "", visible)
    return min(380.0, max(12.0, len(visible.strip()) * 5.8))


BARE_MATH_COMMAND_RE = re.compile(
    r"\\(?:times|div|frac|sqrt|left|right|cdot|pm|le|ge|neq|ne|lt|gt|sum|lim|sin|cos|tan)\b"
)


def wrap_bare_latex_math(text: str) -> str:
    """Wrap obvious docx2tex math fragments that lost $...$ delimiters."""
    parts = MATH_SPAN_RE.split(text)
    for i in range(0, len(parts), 2):
        parts[i] = wrap_bare_latex_math_in_plain_text(parts[i])
    return "".join(parts)


def wrap_bare_latex_math_in_plain_text(text: str) -> str:
    if "\\" not in text and "^" not in text and "_" not in text:
        return text
    text = wrap_bare_line_formula(text)
    return text


def wrap_bare_line_formula(text: str) -> str:
    out: list[str] = []
    pos = 0
    for match in re.finditer(r"(?:(?<=^)|(?<=[\s，,。；;:：]))(?=[^$\n\u3400-\u4DBF\u4E00-\u9FFF]*(?:\\(?:times|div)|\\frac|\\sqrt|\\left|\\right|\^))", text):
        start = expand_bare_math_start(text, match.start())
        if start < pos:
            continue
        end = scan_bare_formula_end(text, start)
        if end <= start:
            continue
        fragment = text[start:end].strip()
        if not should_attach_metric_to_bare_formula(fragment):
            continue
        out.append(text[pos:start])
        out.append("$" + fragment + "$")
        pos = end
    out.append(text[pos:])
    return "".join(out)


def normalize_text_subsup_commands(text: str) -> str:
    parts = MATH_SPAN_RE.split(text)
    for i in range(0, len(parts), 2):
        parts[i] = normalize_text_subsup_commands_fragment(parts[i])
    return "".join(parts)


def normalize_text_subsup_commands_fragment(text: str) -> str:
    out: list[str] = []
    i = 0
    while i < len(text):
        if text.startswith(r"\textsubscript", i) or text.startswith(r"\textsuperscript", i):
            command = "textsubscript" if text.startswith(r"\textsubscript", i) else "textsuperscript"
            j = skip_ws(text, i + len(command) + 1)
            body_end = find_balanced_group_end(text, j)
            if body_end >= 0:
                out.append("$" + text[j + 1:body_end] + "$")
                i = body_end + 1
                continue
        out.append(text[i])
        i += 1
    return "".join(out)


def _unused_wrap_bare_latex_math_in_plain_text_scan(text: str) -> str:
    out: list[str] = []
    cursor = 0
    for match in re.finditer(r"\\[A-Za-z]+|\^|_", text):
        if match.start() < cursor:
            continue
        start = expand_bare_math_start(text, match.start())
        end = expand_bare_math_end(text, match.end())
        fragment = text[start:end].strip()
        if not should_wrap_bare_math_fragment(fragment):
            continue
        out.append(text[cursor:start])
        leading = text[start: start + len(text[start:end]) - len(text[start:end].lstrip())]
        trailing = text[end - (len(text[start:end]) - len(text[start:end].rstrip())):end]
        core = text[start:end].strip()
        out.append(leading + "$" + core + "$" + trailing)
        cursor = end
    out.append(text[cursor:])
    return "".join(out)


def should_wrap_bare_math_fragment(fragment: str) -> bool:
    if not fragment or len(fragment) < 2:
        return False
    if fragment.startswith("\\begin") or fragment.startswith("\\end"):
        return False
    return bool(BARE_MATH_COMMAND_RE.search(fragment) or re.search(r"[A-Za-z0-9）)]\s*[\^_]", fragment))


def is_bare_math_char(ch: str) -> bool:
    return (
        ch.isascii() and (ch.isalnum() or ch.isspace() or ch in "\\{}[]().,+-=*/_^%")
    ) or ch in "（）＝，、."


def expand_bare_math_start(text: str, index: int) -> int:
    i = index
    while i > 0 and is_bare_math_char(text[i - 1]):
        i -= 1
    while i < index and text[i].isspace():
        i += 1
    # Keep labels such as "计算" outside the formula, but include a leading parenthesis
    # when it belongs to the expression itself.
    if i < index and text[i] in "，,。；;:":
        i += 1
    return i


def expand_bare_math_end(text: str, index: int) -> int:
    i = index
    while i < len(text) and is_bare_math_char(text[i]):
        i += 1
    while i > index and text[i - 1].isspace():
        i -= 1
    return i


def unwrap_nested_math_in_text_commands(text: str) -> str:
    """Remove stray dollar delimiters inside LaTeX text/color command arguments."""
    command = r"\\(?:text|textcolor|textbf|textit|textnormal|textrm|mathrm|textsc)"
    pattern = re.compile(
        rf"({command}(?:\s*\{{[^{{}}]*\}})?\s*\{{[^{{}}]*)"
        r"\$([^$]+?)\$"
        r"([^{}]*\})"
    )
    previous = None
    while previous != text:
        previous = text
        text = pattern.sub(lambda m: m.group(1) + m.group(2) + m.group(3), text)
    return text


def escape_html_sensitive_math(text: str, metrics_lookup=None, styles=None, metrics_queue=None) -> str:
    """Avoid Jsoup treating inline inequalities as HTML tags."""
    def repl(match: re.Match[str]) -> str:
        body = repair_docx2tex_latex(match.group(1) or match.group(2) or "")
        existing_metrics = METRICS_RE.match(body)
        existing_prefix = existing_metrics.group(0).strip() if existing_metrics else ""
        body = METRICS_RE.sub("", body).strip()
        if metrics_lookup is not None:
            prefix = existing_prefix
            if not prefix:
                queue = metrics_lookup.get(normalize_latex_key(body))
                item = queue.popleft() if queue else None
                prefix = metric_prefix(item or {}) + style_prefix(item or {}, styles or {})
            if prefix:
                body = prefix + body
        elif metrics_queue is not None:
            item = metrics_queue.popleft() if metrics_queue else None
            prefix = metric_prefix(item or {})
            if prefix:
                body = prefix + body
        body = body.replace("<", r"\lt ").replace(">", r"\gt ")
        if match.group(1) is not None:
            return "$$" + body + "$$"
        return "$" + body + "$"

    return re.sub(r"\$\$(.+?)\$\$|\$(.+?)\$", repl, text, flags=re.S)


def attach_empty_metric_math(text: str) -> str:
    """Move an empty metric-only math span onto the following bare math span."""
    previous = None
    while previous != text:
        previous = text
        text = EMPTY_METRIC_BEFORE_TEXT_COMMAND_RE.sub(
            lambda m: f"{m.group(2)}$\\pwmetrics{{{m.group(1)}}}", text
        )
        text = EMPTY_METRIC_MATH_RE.sub(lambda m: f"{m.group(2)}$\\pwmetrics{{{m.group(1)}}}", text)
        text = attach_metric_before_bare_formula(text)
    return text


def attach_metric_before_bare_formula(text: str) -> str:
    pattern = re.compile(r"\$\\pwmetrics\{([^}]+)}\$\s*")
    out: list[str] = []
    pos = 0
    for match in pattern.finditer(text):
        out.append(text[pos:match.start()])
        formula_end = scan_bare_formula_end(text, match.end())
        if formula_end <= match.end():
            out.append(match.group(0))
            pos = match.end()
            continue
        fragment = text[match.end():formula_end].strip()
        if should_attach_metric_to_bare_formula(fragment):
            out.append(f"$\\pwmetrics{{{match.group(1)}}}{fragment}$")
        else:
            out.append(match.group(0))
            out.append(text[match.end():formula_end])
        pos = formula_end
    out.append(text[pos:])
    return "".join(out)


def scan_bare_formula_end(text: str, start: int) -> int:
    i = start
    depth = 0
    saw_math = False
    while i < len(text):
        ch = text[i]
        if depth == 0 and (ch == "$" or ch == "\n" or ch in "，,。；;"):
            break
        if ch == "\\":
            j = i + 1
            while j < len(text) and text[j].isalpha():
                j += 1
            if j > i + 1:
                saw_math = True
                i = j
                continue
        if ch in "=×÷^_<>":
            saw_math = True
        if text.startswith(r"\times", i) or text.startswith(r"\div", i):
            saw_math = True
        if ch == "{":
            depth += 1
        elif ch == "}":
            if depth == 0:
                break
            depth -= 1
        if depth == 0 and is_cjk_char(ch):
            break
        i += 1
    return i if saw_math and depth == 0 else start


def should_attach_metric_to_bare_formula(fragment: str) -> bool:
    return bool(fragment and (r"\frac" in fragment or BARE_MATH_COMMAND_RE.search(fragment) or re.search(r"[=×÷^_<>]", fragment)))


def is_cjk_char(ch: str) -> bool:
    return "\u3400" <= ch <= "\u4dbf" or "\u4e00" <= ch <= "\u9fff"


def demote_unmeasured_math_spans(text: str, require_metrics: bool = False) -> str:
    """Keep only source-measured math spans as OLE candidates."""
    def repl(match: re.Match[str]) -> str:
        body = (match.group(1) or match.group(2) or "").strip()
        if r"\pwmetrics{" in body or (not require_metrics and should_keep_unmeasured_math_span(body)):
            return ("$$" + body + "$$") if match.group(1) is not None else ("$" + body + "$")
        return demote_latex_math_to_text(body)

    return re.sub(r"\$\$(.+?)\$\$|\$(.+?)\$", repl, text, flags=re.S)


def demote_latex_math_to_text(body: str) -> str:
    """Turn a non-source formula span into readable text instead of raw LaTeX."""
    text = body or ""
    text = METRICS_RE.sub("", text).strip()
    text = STYLE_RE.sub("", text).strip()
    text = re.sub(r"\\begin\{array\}\{[^}]*}", " ", text)
    text = re.sub(r"\\end\{array}", " ", text)
    text = re.sub(r"\\\\", " ", text)
    replacements = {
        r"\times": "×",
        r"\div": "÷",
        r"\cdot": "·",
        r"\left": "",
        r"\right": "",
        r"\lt": "<",
        r"\gt": ">",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = re.sub(r"\^\{([^{}]+)}", r"^\1", text)
    text = text.replace("{", "").replace("}", "")
    return re.sub(r"\s{2,}", " ", text).strip()


def should_keep_unmeasured_math_span(body: str) -> bool:
    if not body:
        return False
    explicit_markers = (
        r"\begin{array}",
        r"\begin{matrix}",
        r"\begin{cases}",
        r"\frac",
        r"\sqrt",
        r"\left",
        r"\right",
        r"\times",
        r"\div",
        r"\pm",
        r"\cdot",
        r"\sum",
        r"\lim",
    )
    if any(marker in body for marker in explicit_markers):
        return True
    return bool(re.search(r"[_^=<>]|\\[A-Za-z]+", body))


def strip_text_latex_linebreaks(text: str) -> str:
    """Remove LaTeX paragraph linebreaks outside math spans."""
    parts = re.split(r"(\$\$.*?\$\$|\$.*?\$)", text, flags=re.S)
    for i in range(0, len(parts), 2):
        parts[i] = parts[i].replace(r"\\", " ")
    return re.sub(r"\s{2,}", " ", "".join(parts)).strip()


def repair_docx2tex_latex(text: str) -> str:
    """Repair known invalid delimiter sequences before strict TeX render."""
    text = unwrap_nested_math_in_text_commands(text)
    text = strip_latex_text_format_commands(text)
    text = re.sub(r"(?<!\\)\$", "", text)
    text = normalize_text_subsup_commands(text)
    text = normalize_ding_commands(text)
    text = DASHLINE_RE.sub(" ", text)
    text = TAG_RE.sub(" ", text)
    text = text.replace(r"\right(", "(")
    text = text.replace(r"\left] )", r"\right] ")
    text = text.replace(r"\left] ", r"] ")
    text = text.replace(r"\right \right", r"\right]")
    text = text.replace(r"\right=", "=")
    text = text.replace(r"\right .", "")
    text = text.replace(r"\right \div", r"\div")
    text = text.replace(r"\left ( \right )", "()")
    text = text.replace(r"\left[ \space \right]", r"\left[ {} \right]")
    text = text.replace(r"{\frownie}", "")
    text = text.replace(r"\frownie", "")
    text = LEFT_ZERO_PERCENT_RE.sub(r"0 \1 \\%", text)
    text = text.replace("△", r"\triangle ")
    text = text.replace("Ўч", r"\triangle ")
    text = text.replace("Δ", r"\Delta ")
    text = text.replace("¦¤", r"\Delta ")
    text = text.replace("忖", r"\Delta ")
    text = text.replace("Θ", r"\Theta ")
    text = text.replace("жи", r"\Theta ")
    text = re.sub(r"(?<![\u3400-\u9FFF])成(?![\u3400-\u9FFF])", r"\\Theta ", text)
    text = text.replace(r"\text{不合{\blacksquare}意}", r"\text{不合题意}")
    text = text.replace("Ο", r"\bigcirc ")
    text = text.replace("○", r"\bigcirc ")
    text = text.replace("●", r"\bullet ")
    text = text.replace("☆", r"\whitestar ")
    text = text.replace("★", r"\blackstar ")
    text = text.replace("◇", r"\whitediamond ")
    text = text.replace("︸", r"\underbracechar ")
    text = text.replace("∶", r"\colon ")
    text = text.replace("≤", r"\le ")
    text = text.replace("≥", r"\ge ")
    text = text.replace("≠", r"\ne ")
    text = text.replace("~", r"\sim ")
    text = text.replace("∼", r"\sim ")
    text = text.replace("⇒", r"\Rightarrow ")
    text = text.replace("⇐", r"\Leftarrow ")
    text = text.replace("⇔", r"\Leftrightarrow ")
    text = text.replace("×", r"\times ")
    text = text.replace("÷", r"\div ")
    text = text.replace("±", r"\pm ")
    text = MATHRM_GREEK_RE.sub(lambda m: GREEK_COMMANDS.get(m.group(1), m.group(1)), text)
    text = MATHRM_COMMAND_RE.sub(lambda m: m.group(1), text)
    text = COMMAND_BEFORE_CJK_RE.sub(r"\1 ", text)
    text = OVERSET_ARC_RE.sub(lambda m: rf"\overarc{{{m.group(1).strip()}}}", text)
    text = OVERSET_RIGHTARROW_RE.sub(lambda m: rf"\xrightarrow{{{m.group(1).strip()}}}", text)
    text = UNDER_RIGHTARROW_RE.sub(lambda m: rf"\xrightarrow{{{m.group(1).strip()}}}", text)
    text = MATHOP_MATHRM_CHAR_RE.sub(lambda m: rf"\mathrm{{{m.group(1).strip()}}}", text)
    text = MATHOP_SIMPLE_RE.sub(lambda m: m.group(1).strip(), text)
    text = ACCENT_MATHRM_CHAR_RE.sub(lambda m: rf"{m.group(1)}{{\mathrm{{{m.group(2).strip()}}}}}", text)
    text = TRAILING_RIGHT_RE.sub(lambda m: rf"\left[ {m.group(1).strip()} \right]", text)
    text = EMPTY_FRAC_DEN_RE.sub(lambda m: rf"\frac {{ {m.group(1).strip()} }} {{ {m.group(2).strip()} }}", text)
    return text


def normalize_ding_commands(text: str) -> str:
    return re.sub(r"\\ding\s*\{\s*(\d+)\s*}", lambda m: DING_CIRCLED.get(m.group(1), ""), text)


def strip_tex_preamble(text: str) -> str:
    start = text.find(r"\begin{document}")
    if start >= 0:
        text = text[start + len(r"\begin{document}") :]
    end = text.rfind(r"\end{document}")
    if end >= 0:
        text = text[:end]
    return text


HEADING_COMMAND_RE = re.compile(r"\\(?:sub)?section\*?\{")


def replace_heading_commands(text: str) -> str:
    """Replace section headings while preserving nested braces inside formulas."""
    out = []
    pos = 0
    while True:
        match = HEADING_COMMAND_RE.search(text, pos)
        if not match:
            out.append(text[pos:])
            break
        out.append(text[pos:match.start()])
        body_start = match.end()
        depth = 1
        i = body_start
        while i < len(text):
            ch = text[i]
            if ch == "\\" and i + 1 < len(text) and text[i + 1] in "{}":
                i += 2
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    body = text[body_start:i]
                    out.append("\n\n【" + body + "】\n\n")
                    pos = i + 1
                    break
            i += 1
        else:
            out.append(text[match.start():])
            break
    return "".join(out)


def tex_to_plain_blocks(
    text: str,
    metrics_lookup=None,
    styles=None,
    metrics_queue=None,
    preserve_source_ole_spans: bool = False,
) -> list[str]:
    text = strip_tex_preamble(text)
    text = normalize_math_delimiters(text)
    text = replace_heading_commands(text)
    text = text.replace(r"\begin{enumerate}", "\n")
    text = text.replace(r"\end{enumerate}", "\n")
    text = re.sub(r"\\item\s*", "\n", text)
    text = re.sub(r"\\begin\{(?:center|flushleft|flushright)\}", "\n", text)
    text = re.sub(r"\\end\{(?:center|flushleft|flushright)\}", "\n", text)
    text = re.sub(r"\\(?:noindent|par)\b", "\n", text)
    text = text.replace("．", ".")
    blocks = []
    for raw in re.split(r"(?:\r?\n\s*){2,}", text):
        block = " ".join(line.strip() for line in raw.splitlines() if line.strip())
        block = re.sub(r"\s{2,}", " ", block).strip()
        if block:
            block = normalize_ding_commands(block)
            block = normalize_text_subsup_commands(block)
            block = strip_text_latex_linebreaks(block)
            block = attach_empty_metric_math(block)
            block = wrap_bare_latex_math(block)
            block = escape_html_sensitive_math(block, metrics_lookup, styles, metrics_queue)
            block = attach_empty_metric_math(block)
            block = strip_non_content_latex_commands(block)
            block = strip_residual_metrics_outside_math(block)
            block = strip_residual_group_braces_outside_math(block)
            if not preserve_source_ole_spans:
                block = merge_inline_operator_math_fragments(block)
            block = strip_non_content_latex_commands(block)
            block = merge_metric_only_formula_runs(block)
            block = separate_adjacent_display_formula_runs(block)
            if not preserve_source_ole_spans:
                block = split_display_array_formula_runs(block)
            block = demote_metadata_numeric_formula_runs(block)
            block = widen_simple_numeric_formula_metrics(block)
            block = widen_too_small_formula_metrics(block)
            block = add_estimated_formula_metrics(block)
            block = strip_empty_bracket_artifacts(block)
            blocks.append(demote_unmeasured_math_spans(block, False))
    return blocks


def resolve_image(tex_dir: Path, name: str) -> str | None:
    analysis_dir = tex_dir.parent.parent if len(tex_dir.parents) >= 2 else tex_dir.parent
    normalized_name = name.replace("\\", "/")
    media_name = Path(normalized_name).name
    candidates = [
        tex_dir / "img" / normalized_name,
        tex_dir / "img" / media_name,
        tex_dir / normalized_name,
        tex_dir / media_name,
        tex_dir.parent / normalized_name,
        tex_dir.parent / media_name,
        analysis_dir / "xsc-docx-ascii-work" / normalized_name,
        analysis_dir / "xsc-docx-ascii-work" / f"{tex_dir.name}.docx.tmp" / "word" / "media" / normalized_name,
        analysis_dir / "xsc-docx-ascii-work" / f"{tex_dir.name}.docx.tmp" / "word" / "media" / media_name,
        Path(normalized_name),
    ]
    path = next((candidate for candidate in candidates if candidate.exists()), None)
    if path is None:
        return None
    if path.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
        return None
    return str(path)


def split_pictures(block: str, tex_dir: Path) -> tuple[str, list[str], list[str]]:
    images = []
    unresolved = []
    def repl(match: re.Match[str]) -> str:
        name = match.group(1)
        image = resolve_image(tex_dir, name)
        if image:
            images.append(image)
        else:
            unresolved.append(name)
        return ""

    text = PIC_RE.sub(repl, block)
    text = INCLUDEGRAPHICS_RE.sub(repl, text)
    text = INCLUDEGRAPHICS_LOOSE_RE.sub(repl, text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text, images, unresolved


def normalize_question_latex_containers(question: dict, tex_dir: Path) -> dict:
    content = question.get("content") or ""
    text, images, unresolved = split_pictures(content, tex_dir)
    text = strip_non_content_latex_commands(text)
    text = strip_residual_group_braces_outside_math(text)
    question["content"] = text
    merged_images = list(question.get("images") or [])
    for image in images:
        if image not in merged_images:
            merged_images.append(image)
    question["images"] = merged_images or None
    merged_unresolved = list(question.get("unresolvedImages") or [])
    for name in unresolved:
        if name not in merged_unresolved:
            merged_unresolved.append(name)
    if merged_unresolved:
        question["unresolvedImages"] = merged_unresolved
    else:
        question.pop("unresolvedImages", None)
    return question


def strip_leading_label(text: str, label: str) -> str:
    return re.sub(rf"^\s*{re.escape(label)}\s*", "", text or "").strip()


def split_inline_metadata(text: str) -> dict[str, str]:
    pattern = re.compile(r"(【考点】|【难度】|【题型】|【关键词】)")
    matches = list(pattern.finditer(text or ""))
    values = {}
    for idx, match in enumerate(matches):
        label = match.group(1)
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        values[label] = text[start:end].strip()
    return values


def normalize_metadata_value(text: str) -> str:
    text = strip_residual_metrics_outside_math(text or "")
    return strip_residual_group_braces_outside_math(text)


def split_metadata_tags(text: str) -> list[str]:
    spans: list[tuple[str, str]] = []

    def protect(match: re.Match[str]) -> str:
        token = f"\x00MATH{len(spans)}\x00"
        spans.append((token, match.group(0)))
        return token

    protected = MATH_SPAN_RE.sub(protect, text or "")
    tags: list[str] = []
    for raw in re.split(r"[、,，]", protected):
        tag = raw.strip()
        if not tag:
            continue
        for token, span in spans:
            tag = tag.replace(token, span)
        tags.append(tag.strip())
    return [tag for tag in tags if tag]


def merge_question_media(target: dict, source: dict) -> None:
    for key in ("images", "unresolvedImages"):
        merged = list(target.get(key) or [])
        for item in source.get(key) or []:
            if item not in merged:
                merged.append(item)
        if merged:
            target[key] = merged
        else:
            target.pop(key, None)


def append_question_field(question: dict, field: str, value: str) -> None:
    if not value:
        return
    question[field] = (question.get(field, "") + "<br/>" + value).strip("<br/>")
    if math_bodies(value):
        question.setdefault("_mathOrder", []).append(value)


def fold_labeled_blocks_into_questions(questions: list[dict], preserve_formula_order: bool = False) -> list[dict]:
    folded = []
    last_question = None
    for question in questions:
        content = (question.get("content") or "").strip()
        if preserve_formula_order and math_bodies(content):
            folded.append(question)
            if question.get("serialNumber") is not None and content:
                last_question = question
            continue
        if last_question is not None and content.startswith(("【考点】", "【难度】", "【题型】")):
            values = split_inline_metadata(content)
            if "【考点】" in values:
                last_question["knowledgePoint"] = normalize_metadata_value(values["【考点】"])
            if "【难度】" in values:
                last_question["difficulty"] = normalize_metadata_value(values["【难度】"])
            if "【关键词】" in values:
                tags = split_metadata_tags(normalize_metadata_value(values["【关键词】"]))
                if tags:
                    last_question["tags"] = tags
            merge_question_media(last_question, question)
            continue
        if last_question is not None and content.startswith("【关键词】"):
            tag_text = normalize_metadata_value(strip_leading_label(content, "【关键词】"))
            tags = split_metadata_tags(tag_text)
            if tags:
                existing = list(last_question.get("tags") or [])
                for tag in tags:
                    if tag not in existing:
                        existing.append(tag)
                last_question["tags"] = existing
            merge_question_media(last_question, question)
            continue
        if last_question is not None and content.startswith("【解析】"):
            analyze = strip_leading_label(content, "【解析】")
            append_question_field(last_question, "analyze", analyze)
            merge_question_media(last_question, question)
            continue
        if last_question is not None and content.startswith("【答案】"):
            correct = strip_leading_label(content, "【答案】")
            append_question_field(last_question, "correct", correct)
            merge_question_media(last_question, question)
            continue
        folded.append(question)
        if question.get("serialNumber") is not None and content:
            last_question = question
    for serial, question in enumerate((q for q in folded if q.get("serialNumber") is not None), 1):
        question["serialNumber"] = serial
    return folded


LOST_SPEED_SUBSCRIPT_RE = re.compile(r"V_\{(?:\\mathrm\{)?�\s*(?:\})?\}")


def repair_known_replacement_context(text: str) -> str:
    """Restore the four speed labels whose paragraph-level ratios identify them uniquely."""
    if "1：12" not in text or "1：16" not in text:
        return text
    matches = list(LOST_SPEED_SUBSCRIPT_RE.finditer(text))
    if len(matches) != 4:
        return text
    labels = iter(("甲", "车", "乙", "车"))
    return LOST_SPEED_SUBSCRIPT_RE.sub(lambda _: rf"V_{{\mathrm{{{next(labels)}}}}}", text)


def build_request(index: int, tex_path: Path, latex_root: Path, styles: dict[int, str] | None = None,
                  source_docx: Path | None = None, source_name: str | None = None) -> tuple[dict, int, list[dict]]:
    equations = report_equations(index, latex_root)
    preserve_source_ole_spans = bool(equations)
    if not equations:
        equations = source_docx_equations(tex_path, source_docx)
        preserve_source_ole_spans = bool(equations)
    styles = styles or {}
    metrics_lookup = equation_metric_lookup(equations) if equations else None
    cursor_multi_consumes: list[dict] = []
    if preserve_source_ole_spans:
        blocks, cursor_multi_consumes = blocks_from_mml2tex_xml(tex_path, equations, styles)
    else:
        blocks = []
    if not blocks:
        blocks = tex_to_plain_blocks(
            tex_path.read_text(encoding="utf-8"),
            metrics_lookup,
            styles,
            None,
            preserve_source_ole_spans,
        )
    questions = []
    serial = 1
    for block in blocks:
        block = repair_known_replacement_context(block)
        block = unwrap_nested_math_in_text_commands(block)
        text, images, unresolved = split_pictures(block, tex_path.parent)
        text = strip_non_content_latex_commands(text)
        text = strip_residual_group_braces_outside_math(text)
        if not text and images:
            picture_question = {
                "serialNumber": None,
                "questionType": None,
                "score": 0,
                "content": "",
                "images": images,
            }
            if unresolved:
                picture_question["unresolvedImages"] = unresolved
            questions.append(picture_question)
            continue
        if not text:
            continue
        question = {
            "serialNumber": serial,
            "questionType": 5,
            "score": 1,
            "content": text,
            "images": images or None,
        }
        if math_bodies(text):
            question["_mathOrder"] = [text]
        if unresolved:
            question["unresolvedImages"] = unresolved
        questions.append(normalize_question_latex_containers(question, tex_path.parent))
        serial += 1
    questions = fold_labeled_blocks_into_questions(questions, preserve_source_ole_spans)
    request = {
        "paper": {
            "name": f"测试集完整重建_{index}.docx",
            "subjectType": 2,
            "stage": 2,
            "score": len(questions),
            "suggestTime": 90,
            "sourceName": source_name or (source_docx.name if source_docx else f"{index}.docx"),
            "compactLayout": True,
            "hideQuestionTypeMetadata": True,
        },
        "sections": [
            {
                "headline": f"源文件 {index}.docx 完整转写",
                "imageMaxWidthPx": 520,
                "questions": questions,
            }
        ],
    }
    if preserve_source_ole_spans:
        appended_missing = 0
        expected = len(equations)
        actual_sequence = request_math_sequence(request)
        actual = len(actual_sequence)
        if actual < expected:
            raise ValueError(
                f"source OLE formula count mismatch for {tex_path}: expected {expected}, got {actual}"
            )
        if actual != expected:
            raise ValueError(
                f"source OLE formula count mismatch for {tex_path}: expected {expected}, got {actual}"
            )
        assert_formula_sequence_matches(tex_path, equations, actual_sequence)
    else:
        appended_missing = 0
    return request, appended_missing, cursor_multi_consumes


def count_request_math(request: dict) -> int:
    return len(request_math_sequence(request))


def append_verified_trailing_formulas(request: dict, formulas: list[dict], styles: dict[int, str]) -> int:
    raise RuntimeError("trailing formula append is disabled; preserve source location instead")
    questions = request["sections"][0]["questions"]
    if not questions:
        return 0
    last_question = questions[-1]
    appended = 0
    for formula in formulas:
        text = latex_formula_text(formula, styles, "")
        last_question["content"] = (last_question.get("content", "") + "<br/>" + text).strip("<br/>")
        last_question.setdefault("_mathOrder", []).append(text)
        appended += 1
    return appended


def request_math_sequence(request: dict) -> list[str]:
    fields = ("content", "analyze", "solution", "correct", "difficulty", "knowledgePoint")
    sequence = []
    for question in request["sections"][0]["questions"]:
        ordered_fields = question.get("_mathOrder")
        if ordered_fields is not None:
            for item in ordered_fields:
                sequence.extend(math_bodies(str(item or "")))
            for field in fields:
                if field == "content" or question.get(field) in ordered_fields:
                    continue
                sequence.extend(math_bodies(str(question.get(field) or "")))
        else:
            for field in fields:
                sequence.extend(math_bodies(str(question.get(field) or "")))
        for tag in question.get("tags") or []:
            sequence.extend(math_bodies(str(tag)))
    return sequence


def math_bodies(text: str) -> list[str]:
    return [
        (match.group(1) or match.group(2) or "").strip()
        for match in re.finditer(r"\$\$(.+?)\$\$|\$(.+?)\$", text, re.S)
    ]


def formula_sequence_key(latex: str) -> str:
    value = repair_docx2tex_latex(latex or "")
    value = METRICS_RE.sub("", value).strip()
    value = STYLE_RE.sub("", value).strip()
    return normalize_latex_key(value)


def source_formula_sequence(equations: list[dict]) -> list[str]:
    return [formula_sequence_key(item.get("output", "")) for item in equations]


def request_formula_sequence(sequence: list[str]) -> list[str]:
    return [formula_sequence_key(item) for item in sequence]


def assert_formula_sequence_matches(tex_path: Path, equations: list[dict], actual_sequence: list[str]) -> None:
    expected = source_formula_sequence(equations)
    actual = request_formula_sequence(actual_sequence)
    if expected == actual:
        return
    mismatch = next(
        (idx for idx, (left, right) in enumerate(zip(expected, actual), 1) if left != right),
        min(len(expected), len(actual)) + 1,
    )
    raise ValueError(
        f"source OLE formula sequence mismatch for {tex_path}: "
        f"first mismatch at {mismatch}, expected={expected[mismatch - 1:mismatch]}, actual={actual[mismatch - 1:mismatch]}"
    )


def strip_internal_request_fields(request: dict) -> None:
    for question in request["sections"][0]["questions"]:
        question.pop("_mathOrder", None)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--end", type=int, default=10)
    parser.add_argument("--latex-root", type=Path, default=DEFAULT_LATEX_ROOT)
    parser.add_argument("--fallback-latex-root", type=Path, default=DEFAULT_FALLBACK_LATEX_ROOT)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    parser.add_argument("--style-hints-root", type=Path, default=STYLE_HINTS_ROOT)
    parser.add_argument("--source-root", type=Path)
    parser.add_argument("--manifest", type=Path)
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    manifest_lookup = {}
    if args.manifest and not args.manifest.exists():
        raise FileNotFoundError(f"missing manifest: {args.manifest}")
    if args.manifest:
        manifest_data = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
        if isinstance(manifest_data, dict):
            manifest_data = [manifest_data]
        manifest_lookup = {
            int(item["index"]): item
            for item in manifest_data
            if item.get("index") and item.get("sourcePath")
        }
    summary = []
    for index in range(args.start, args.end + 1):
        tex_path = args.latex_root / str(index) / f"{index}.tex"
        latex_root = args.latex_root
        if not tex_path.exists():
            fallback_tex = args.fallback_latex_root / str(index) / f"{index}.tex"
            if fallback_tex.exists():
                tex_path = fallback_tex
                latex_root = args.fallback_latex_root
            else:
                raise FileNotFoundError(f"missing tex: {tex_path}")
        styles = style_hint_lookup(index, args.style_hints_root)
        manifest_item = manifest_lookup.get(index)
        source_docx = Path(manifest_item["sourcePath"]) if manifest_item else None
        source_name = manifest_item.get("sourceName") if manifest_item else None
        if source_docx is None and args.source_root:
            source_docx = args.source_root / f"{index}.docx"
        request, appended_missing, cursor_multi_consumes = build_request(
            index, tex_path, latex_root, styles, source_docx, source_name
        )
        strip_internal_request_fields(request)
        out_path = args.out_dir / f"full-{index:02d}.request.json"
        out_path.write_text(json.dumps(request, ensure_ascii=False, indent=2), encoding="utf-8")
        question_count = len(request["sections"][0]["questions"])
        image_count = len(request["sections"][0].get("images") or []) + sum(
            len(q.get("images") or []) for q in request["sections"][0]["questions"]
        )
        summary.append(
            {
                "source": f"{index}.docx",
                "request": str(out_path),
                "questions": question_count,
                "png_jpeg_images": image_count,
                "appended_missing_equations": appended_missing,
                "cursor_multi_consumes": len(cursor_multi_consumes),
                "cursor_multi_consume_samples": cursor_multi_consumes[:10],
            }
        )
    (args.out_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
