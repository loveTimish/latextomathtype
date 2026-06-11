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
from collections import defaultdict, deque
from pathlib import Path


ROOT = Path(r"D:\latextomathtype\analysis")
DEFAULT_LATEX_ROOT = ROOT / "batch10-latex"
DEFAULT_OUT_DIR = ROOT / "batch10-full-requests"
STYLE_HINTS_ROOT = ROOT / "mtef-style-hints"
PIC_RE = re.compile(r"beginPic\{([^}]+)\}endPic")
METRICS_RE = re.compile(r"^\\pwmetrics\{[^}]+}\s*")
STYLE_RE = re.compile(r"^\\pwstyle\{[^}]*}\s*")
OLE_OBJECT_RE = re.compile(r"oleObject(\d+)\.bin")
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


def report_equations(index: int, latex_root: Path) -> list[dict]:
    report_path = latex_root / str(index) / f"{index}.report.json"
    report = json.loads(report_path.read_text(encoding="utf-8"))
    return [
        item for item in report.get("equations", [])
        if item.get("status") == "converted" and item.get("output")
    ]


def metric_prefix(item: dict) -> str:
    metrics = item.get("metrics") or {}
    width = metrics.get("wmfWidthPt") or metrics.get("shapeWidthPt") or metrics.get("dxaOrigPt")
    height = metrics.get("wmfHeightPt") or metrics.get("shapeHeightPt") or metrics.get("dyaOrigPt")
    if not width or not height:
        return ""
    return f"\\pwmetrics{{{float(width):.3f},{float(height):.3f}}}"


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
    text = re.sub(r"\\\[(.+?)\\\]", lambda m: "$$" + m.group(1).strip() + "$$", text, flags=re.S)
    text = re.sub(r"\\\((.+?)\\\)", lambda m: "$" + m.group(1).strip() + "$", text, flags=re.S)
    return text


def escape_html_sensitive_math(text: str, metrics_lookup=None, styles=None) -> str:
    """Avoid Jsoup treating inline inequalities as HTML tags."""
    def repl(match: re.Match[str]) -> str:
        body = repair_docx2tex_latex(match.group(1) or match.group(2) or "")
        body = METRICS_RE.sub("", body).strip()
        if metrics_lookup is not None:
            queue = metrics_lookup.get(normalize_latex_key(body))
            item = queue.popleft() if queue else None
            prefix = metric_prefix(item or {}) + style_prefix(item or {}, styles or {})
            if prefix:
                body = prefix + body
        body = body.replace("<", r"\lt ").replace(">", r"\gt ")
        if match.group(1) is not None:
            return "$$" + body + "$$"
        return "$" + body + "$"

    return re.sub(r"\$\$(.+?)\$\$|\$(.+?)\$", repl, text, flags=re.S)


def strip_text_latex_linebreaks(text: str) -> str:
    """Remove LaTeX paragraph linebreaks outside math spans."""
    parts = re.split(r"(\$\$.*?\$\$|\$.*?\$)", text, flags=re.S)
    for i in range(0, len(parts), 2):
        parts[i] = parts[i].replace(r"\\", " ")
    return re.sub(r"\s{2,}", " ", "".join(parts)).strip()


def repair_docx2tex_latex(text: str) -> str:
    """Repair known invalid delimiter sequences before strict TeX render."""
    text = text.replace(r"\right(", "(")
    text = text.replace(r"\left] )", r"\right] ")
    text = text.replace(r"\left] ", r"] ")
    text = text.replace(r"\right \right", r"\right]")
    text = text.replace(r"\right=", "=")
    text = text.replace(r"\right .", "")
    text = text.replace(r"\right \div", r"\div")
    text = text.replace(r"\left ( \right )", "()")
    text = text.replace(r"\left[ \space \right]", r"\left[ {} \right]")
    text = LEFT_ZERO_PERCENT_RE.sub(r"0 \1 \\%", text)
    text = text.replace("△", r"\triangle ")
    text = text.replace("Ўч", r"\triangle ")
    text = text.replace("Δ", r"\Delta ")
    text = text.replace("¦¤", r"\Delta ")
    text = text.replace("忖", r"\Delta ")
    text = text.replace("Θ", r"\Theta ")
    text = text.replace("жи", r"\Theta ")
    text = text.replace("成", r"\Theta ")
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


def tex_to_plain_blocks(text: str, metrics_lookup=None, styles=None) -> list[str]:
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
            block = strip_text_latex_linebreaks(block)
            blocks.append(escape_html_sensitive_math(block, metrics_lookup, styles))
    return blocks


def resolve_image(tex_dir: Path, name: str) -> str | None:
    path = tex_dir / "img" / name
    if not path.exists():
        return None
    if path.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
        return None
    return str(path)


def split_pictures(block: str, tex_dir: Path) -> tuple[str, list[str]]:
    images = []
    def repl(match: re.Match[str]) -> str:
        image = resolve_image(tex_dir, match.group(1))
        if image:
            images.append(image)
        return ""

    text = PIC_RE.sub(repl, block)
    text = re.sub(r"\s{2,}", " ", text).strip()
    return text, images


def build_request(index: int, tex_path: Path, latex_root: Path, styles: dict[int, str] | None = None) -> dict:
    equations = report_equations(index, latex_root)
    styles = styles or {}
    blocks = tex_to_plain_blocks(tex_path.read_text(encoding="utf-8"), equation_metric_lookup(equations), styles)
    questions = []
    serial = 1
    section_images = []
    for block in blocks:
        text, images = split_pictures(block, tex_path.parent)
        if not text and images:
            section_images.extend(images)
            continue
        if not text:
            continue
        questions.append(
            {
                "serialNumber": serial,
                "questionType": 5,
                "score": 1,
                "content": text,
                "images": images or None,
            }
        )
        serial += 1
    return {
        "paper": {
            "name": f"测试集完整重建_{index}.docx",
            "subjectType": 2,
            "stage": 2,
            "score": len(questions),
            "suggestTime": 90,
            "compactLayout": True,
        },
        "sections": [
            {
                "headline": f"源文件 {index}.docx 完整转写",
                "images": section_images or None,
                "imageMaxWidthPx": 520,
                "questions": questions,
            }
        ],
    }


def count_request_math(request: dict) -> int:
    text = "\n".join(q.get("content", "") for q in request["sections"][0]["questions"])
    return len(re.findall(r"\$\$(.+?)\$\$|\$(.+?)\$", text, re.S))


def append_missing_report_equations(index: int, request: dict, latex_root: Path, styles: dict[int, str] | None = None) -> int:
    """Append equations that were converted by docx2tex but lost while parsing .tex."""
    converted = report_equations(index, latex_root)
    expected = len(converted)
    current = count_request_math(request)
    missing = max(expected - current, 0)
    if missing == 0:
        return 0

    questions = request["sections"][0]["questions"]
    serial = len(questions) + 1
    for offset, item in enumerate(converted[-missing:], 1):
        latex = repair_docx2tex_latex(item.get("output", "").strip())
        prefix = metric_prefix(item)
        prefix += style_prefix(item, styles or {})
        if prefix:
            latex = prefix + latex
        questions.append(
            {
                "serialNumber": serial,
                "questionType": 5,
                "score": 1,
                "content": f"【补入丢失公式 {offset}】$$ {latex} $$",
            }
        )
        serial += 1
    request["paper"]["score"] = len(questions)
    return missing


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--end", type=int, default=10)
    parser.add_argument("--latex-root", type=Path, default=DEFAULT_LATEX_ROOT)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    parser.add_argument("--style-hints-root", type=Path, default=STYLE_HINTS_ROOT)
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    summary = []
    for index in range(args.start, args.end + 1):
        tex_path = args.latex_root / str(index) / f"{index}.tex"
        if not tex_path.exists():
            raise FileNotFoundError(f"missing tex: {tex_path}")
        styles = style_hint_lookup(index, args.style_hints_root)
        request = build_request(index, tex_path, args.latex_root, styles)
        appended_missing = append_missing_report_equations(index, request, args.latex_root, styles)
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
            }
        )
    (args.out_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
