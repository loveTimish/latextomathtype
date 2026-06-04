# -*- coding: utf-8 -*-
"""
Build a PaperExportRequest JSON from the docx2tex output of
rebuild-assets/external/fraction-split-reference.docx.

This is intentionally reference-document specific: its job is to preserve the
section and question structure of the "分数裂项" handout closely enough for
latextomathtype to regenerate a comparable Word document.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
TEX = ROOT / "target/reference-roundtrip/docx2tex/fraction-split-reference.tex"
MEDIA_DIR = ROOT / "rebuild-assets/external/fraction-split-reference.docx.tmp/word/media"
OUT_JSON = ROOT / "target/reference-roundtrip/fraction-split-reference.request.json"
PREVIEW = ROOT / "target/reference-roundtrip/fraction-split-reference.preview.txt"
GENERATED_MEDIA_DIR = ROOT / "target/reference-roundtrip/generated-media"
TITLE_IMAGE = GENERATED_MEDIA_DIR / "fraction-title.png"
SECTION_IMAGE_PREFIX = GENERATED_MEDIA_DIR / "section"
TITLE_FONT = Path(r"C:\Windows\Fonts\simhei.ttf")


def unwrap_markup(text: str) -> str:
    previous = None
    current = text
    patterns = [
        (r"\\textbf\{([^{}]*)\}", r"\1"),
        (r"\\textcolor\{[^{}]*\}\{([^{}]*)\}", r"\1"),
        (r"\\emph\{([^{}]*)\}", r"\1"),
        (r"\\underline\{([^{}]*)\}", r"\1"),
        (r"\{\\rm\{([^{}]*)\}\}", r"\1"),
    ]
    while previous != current:
        previous = current
        for pattern, replacement in patterns:
            current = re.sub(pattern, replacement, current)
    return current


def flatten_redundant_nested_arrays(text: str) -> str:
    pattern = re.compile(
        r"\\begin\{array\}\{(?:r|c|l)?l\}\s*&?\s*"
        r"\\begin\{array\}\{l\}(.*?)\\end\{array\}\s*\\end\{array\}",
        flags=re.S,
    )
    previous = None
    while previous != text:
        previous = text
        text = pattern.sub(lambda m: r"\begin{array}{l}" + m.group(1).strip() + r"\end{array}", text)
    return text


def protect_math(text: str) -> tuple[str, list[str]]:
    text = flatten_redundant_nested_arrays(text)
    math: list[str] = []

    def store(value: str) -> str:
        placeholder = f"@@MATH{len(math)}@@"
        math.append(value)
        return placeholder

    def env_repl(match: re.Match[str]) -> str:
        env = match.group(1)
        body = match.group(2).strip()
        if env.startswith("equation"):
            return store(f"${body}$")
        if env in {"align", "align*", "aligned", "split"}:
            return store(f"${aligned_to_array(body)}$")
        return store(f"$\\begin{{{env}}}{body}\\end{{{env}}}$")

    text = re.sub(
        r"\\begin\{(equation\*?|align\*?|aligned|array|split)\}(.*?)\\end\{\1\}",
        env_repl,
        text,
        flags=re.S,
    )
    text = re.sub(r"\\\[(.*?)\\\]", lambda m: store(f"$${m.group(1).strip()}$$"), text, flags=re.S)
    text = re.sub(r"\$\$(.*?)\$\$", lambda m: store(f"$${m.group(1).strip()}$$"), text, flags=re.S)
    text = re.sub(r"\$(.*?)\$", lambda m: store(f"${m.group(1).strip()}$"), text, flags=re.S)
    return text, math


def aligned_to_array(body: str) -> str:
    rows = re.split(r"(?<!\\)\\\\", body)
    max_columns = 1
    for row in rows:
        max_columns = max(max_columns, row.count("&") + 1)
    if max_columns <= 1:
        column_spec = "l"
    elif max_columns == 2:
        column_spec = "rl"
    else:
        column_spec = "r" + ("l" * (max_columns - 1))
    return f"\\begin{{array}}{{{column_spec}}}{body}\\end{{array}}"


def split_array_rows(body: str) -> list[str]:
    rows = []
    current = []
    index = 0
    depth = 0
    while index < len(body):
        if body.startswith(r"\begin{", index):
            depth += 1
            current.append(body[index])
            index += 1
            continue
        if body.startswith(r"\end{", index):
            depth = max(depth - 1, 0)
            current.append(body[index])
            index += 1
            continue
        if body.startswith(r"\\", index) and depth == 0:
            rows.append("".join(current).strip())
            current = []
            index += 2
            continue
        current.append(body[index])
        index += 1
    tail = "".join(current).strip()
    if tail:
        rows.append(tail)
    return rows


def chunk_long_array_formula(value: str) -> str:
    match = re.fullmatch(r"\$(\\begin\{array\}\{([^{}]*)\})(.*)(\\end\{array\})\$", value, flags=re.S)
    if not match:
        return value
    begin, column_spec, body, end = match.groups()
    nested = re.fullmatch(r"&?\s*\\begin\{array\}\{l\}\s*(.*)\s*\\end\{array\}", body.strip(), flags=re.S)
    if nested and column_spec in {"rl", "cl", "ll"}:
        begin = r"\begin{array}{l}"
        column_spec = "l"
        body = nested.group(1).strip()
        end = r"\end{array}"
    rows = split_array_rows(body.strip())
    if len(rows) < 2:
        return value
    if sum(len(row) for row in rows) >= 150 or any(len(row) >= 150 for row in rows):
        return "\n\n".join(
            "$" + begin + row + end + "$"
            for row in rows
        )
    if sum(len(row) for row in rows) < 220 and all(len(row) < 150 for row in rows):
        return value
    chunks: list[list[str]] = []
    current: list[str] = []
    current_len = 0
    for row in rows:
        row_len = len(row)
        if current and (current_len + row_len > 180 or len(current) >= 2):
            chunks.append(current)
            current = []
            current_len = 0
        current.append(row)
        current_len += row_len
    if current:
        chunks.append(current)
    if len(chunks) <= 1:
        return value
    return "\n\n".join(
        "$" + begin + r" \\".join(chunk) + end + "$"
        for chunk in chunks
    )


def normalize_math_for_word(value: str) -> str:
    return chunk_long_array_formula(value)


def restore_math(text: str, math: list[str]) -> str:
    previous = None
    while previous != text:
        previous = text
        for index, value in enumerate(math):
            text = text.replace(f"@@MATH{index}@@", normalize_math_for_word(value))
    return text


def normalize_chinese_punctuation(text: str) -> str:
    cjk = r"\u3400-\u9fff"
    text = re.sub(rf"(?<=[{cjk}]),(?=[{cjk}])", "，", text)
    text = re.sub(rf"(?<=[{cjk}]),(?=\s*[{cjk}])", "，", text)
    text = re.sub(rf"(?<=[{cjk}])\.(?=[{cjk}])", "。", text)
    text = re.sub(rf"(?<=[{cjk}])\.(?=\s*[{cjk}])", "。", text)
    return text


def clean_text(text: str) -> str:
    text = re.sub(r"=\\\$\s{4,}(?=\n|$)", lambda _m: r"=$ [[ANSWER_LINE]]", text)
    text, math = protect_math(text)
    text = re.sub(r"\\begin\{description\}(?:\[[^\]]*])?", "", text)
    text = text.replace(r"\end{description}", "")
    text = re.sub(r"\\item(?:\[[^\]]*])?", "", text)
    text = re.sub(r"\\fbox\{\\begin\{minipage\}\[t]\{0\.8\\textwidth\}", "", text)
    text = text.replace(r"\end{minipage}}", "")
    text = text.replace(r"\par", "")
    text = text.replace(r"\newline", "\n\n")
    text = text.replace(r"\\", "<br/>")
    text = unwrap_markup(text)
    replacements = {
        r"\textbackslash{}": "\\",
        r"\_": "_",
        r"\{": "{",
        r"\}": "}",
        r"\^{}": "^",
        r"{\ldots}{\ldots}": "...",
        r"{\ldots}": "...",
        r"\ldots": "...",
        r"\cdots": r"\cdots",
        "``": "“",
        "''": "”",
        "＝": "=",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"(?:\{\.{3}\})+", "...", text)
    text = re.sub(r"\\(?:centering|noindent)\b", "", text)
    text = re.sub(r"\\[a-zA-Z]+\*?(?:\[[^\]]*])?(?:\{[^{}]*})?", "", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"^\{\s*\}$", "", text.strip(), flags=re.S)
    text = normalize_chinese_punctuation(text)
    return restore_math(text.strip(), math)


def extract_images(text: str) -> tuple[str, list[str]]:
    images: list[str] = []

    def repl(match: re.Match[str]) -> str:
        source = Path(match.group(1)).name
        path = MEDIA_DIR / source
        if path.exists() and path.suffix.lower() in {".png", ".jpg", ".jpeg"}:
            images.append(str(path))
        return "\n"

    text = re.sub(r"\\includegraphics(?:\[[^\]]*])?\{([^}]+)\}", repl, text)
    return text, images


def fit_font(text: str, size: int, max_width: int) -> ImageFont.FreeTypeFont:
    size = max(size, 10)
    while size >= 10:
        font = ImageFont.truetype(str(TITLE_FONT), size)
        left, top, right, bottom = ImageDraw.Draw(Image.new("RGB", (1, 1))).textbbox((0, 0), text, font=font)
        if right - left <= max_width:
            return font
        size -= 2
    return ImageFont.truetype(str(TITLE_FONT), 10)


def render_title_image(text: str) -> str:
    source = MEDIA_DIR / "image1.png"
    image = Image.open(source).convert("RGBA")
    draw = ImageDraw.Draw(image)
    font = fit_font(text, 64, 520)
    box = draw.textbbox((0, 0), text, font=font)
    x = 630 - (box[2] - box[0]) // 2
    y = 120 - (box[3] - box[1]) // 2
    draw.text((x, y), text, font=font, fill=(0, 0, 0, 255))
    GENERATED_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
    image.save(TITLE_IMAGE)
    return str(TITLE_IMAGE)


def render_section_image(text: str, index: int) -> str:
    source = MEDIA_DIR / "image2.jpeg"
    base = Image.open(source).convert("RGBA")
    scale = 0.30
    shelf = base.resize((int(base.width * scale), int(base.height * scale)))
    width = 520
    height = 120
    image = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    image.alpha_composite(shelf, (0, 46))
    font = fit_font(text, 44, 330)
    draw = ImageDraw.Draw(image)
    stroke = max(2, font.size // 12)
    draw.text((198, 34), text, font=font, fill=(0, 0, 0, 255), stroke_width=stroke, stroke_fill=(255, 255, 255, 255))
    GENERATED_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
    output = Path(f"{SECTION_IMAGE_PREFIX}-{index}.png")
    image.save(output)
    return str(output)


def section_title(match: re.Match[str]) -> str:
    return clean_text(match.group(1))


def strip_preamble(tex: str) -> str:
    start = tex.find(r"\begin{document}")
    if start >= 0:
        tex = tex[start + len(r"\begin{document}"):]
    end = tex.rfind(r"\end{document}")
    if end >= 0:
        tex = tex[:end]
    return tex


def split_sections(tex: str) -> list[tuple[str, str]]:
    pattern = re.compile(
        r"\\fbox\{\\begin\{minipage\}\[t]\{0\.8\\textwidth\}\\textbf\{(.*?)\}\\end\{minipage\}\}",
        re.S,
    )
    matches = list(pattern.finditer(tex))
    sections: list[tuple[str, str]] = []
    for index, match in enumerate(matches):
        title = section_title(match)
        body_start = match.end()
        body_end = matches[index + 1].start() if index + 1 < len(matches) else len(tex)
        sections.append((title, tex[body_start:body_end]))
    return sections


def normalize_question_markers(text: str) -> str:
    text = unwrap_markup(text)
    text = re.sub(r"\\begin\{description\}(?:\[[^\]]*])?", "\n", text)
    text = text.replace(r"\end{description}", "\n")
    text = re.sub(r"\\item\[(.*?)\]", lambda m: "\n" + clean_text(m.group(1)) + "\t", text)
    return text


def split_questions(body: str) -> tuple[str, list[tuple[int, str]]]:
    body = normalize_question_markers(body)
    marker = re.compile(r"(?:【例\s*\d+】|【巩固】)")
    matches = list(marker.finditer(body))
    parts: list[tuple[int, str]] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        part = body[match.start():end].strip()
        if part:
            parts.append((match.start(), part))
    return body, parts


def parse_question(segment: str, serial: int) -> dict:
    segment, images = extract_images(segment)
    segment = clean_text(segment)

    labels = ["【考点】", "【难度】", "【题型】", "【关键词】", "【解析】", "【答案】"]
    positions = [(label, segment.find(label)) for label in labels if segment.find(label) >= 0]
    positions.sort(key=lambda item: item[1])

    first_label = positions[0][1] if positions else len(segment)
    content = segment[:first_label].strip()
    fields: dict[str, str] = {}
    for index, (label, start) in enumerate(positions):
        value_start = start + len(label)
        value_end = positions[index + 1][1] if index + 1 < len(positions) else len(segment)
        fields[label] = segment[value_start:value_end].strip()

    content = re.sub(r"^(【例\s*\d+】|【巩固】)\s*", lambda m: m.group(1) + " ", content).strip()
    if content.endswith("=$") and "____" not in content and "[[ANSWER_LINE]]" not in content:
        content += " [[ANSWER_LINE]]"
    if not content:
        content = " "

    question = {
        "serialNumber": serial,
        "questionType": 6,
        "content": content,
    }
    if images:
        question["images"] = images
    if fields.get("【考点】"):
        question["knowledgePoint"] = fields["【考点】"]
    if fields.get("【难度】"):
        question["difficulty"] = fields["【难度】"]
    if fields.get("【关键词】"):
        question["tags"] = [item.strip() for item in re.split(r"[,，、]", fields["【关键词】"]) if item.strip()]
    if fields.get("【解析】"):
        question["analyze"] = fields["【解析】"]
    if fields.get("【答案】"):
        question["correct"] = fields["【答案】"]
    return question


def body_as_intro_question(body: str, serial: int | None = None) -> dict | None:
    body, images = extract_images(body)
    text = clean_text(body)
    if not text and not images:
        return None
    question = {
        "questionType": 6,
        "content": text if text else " ",
    }
    if serial is not None:
        question["serialNumber"] = serial
    if images:
        question["images"] = images
    return question


def build_request() -> dict:
    tex = strip_preamble(TEX.read_text(encoding="utf-8"))
    sections = split_sections(tex)
    output_sections = []
    serial = 0

    for section_index, (title, body) in enumerate(sections):
        questions = []
        section_images: list[str] = []
        normalized_body, parsed_questions = split_questions(body)
        if parsed_questions:
            before_first = normalized_body[:parsed_questions[0][0]]
            intro = body_as_intro_question(before_first)
            if intro:
                if intro.get("images"):
                    if section_index == 0:
                        section_images.append(render_title_image(title))
                    else:
                        section_images.append(render_section_image(title, section_index))
                if intro.get("content", "").strip():
                    intro["phaseLabel"] = "版式说明"
                    questions.append(intro)
            for _, part in parsed_questions:
                serial += 1
                questions.append(parse_question(part, serial))
        else:
            intro = body_as_intro_question(body)
            if intro:
                if intro.get("images"):
                    if section_index == 0:
                        section_images.append(render_title_image(title))
                    else:
                        section_images.append(render_section_image(title, section_index))
                if intro.get("content", "").strip():
                    questions.append(intro)

        if questions or section_images:
            section = {
                "headline": None if section_images else title,
                "questions": questions,
            }
            if section_images:
                section["images"] = section_images
                section["imageMaxWidthPx"] = 430 if title == "分数裂项计算" else 285
                section["imageAlignment"] = "center" if title == "分数裂项计算" else "left"
            output_sections.append(section)

    return {
        "paper": {
            "subjectType": 1,
            "stage": 2,
            "compactLayout": True,
        },
        "sections": output_sections,
    }


def main() -> None:
    request = build_request()
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(request, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = [
        f"sections={len(request['sections'])}",
        f"questions={sum(len(s['questions']) for s in request['sections'])}",
        f"numbered={sum(1 for s in request['sections'] for q in s['questions'] if q.get('serialNumber'))}",
    ]
    for section in request["sections"]:
        lines.append(f"- {section['headline']}: {len(section['questions'])}")
    PREVIEW.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"JSON -> {OUT_JSON}")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
