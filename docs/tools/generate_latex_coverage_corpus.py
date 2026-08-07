#!/usr/bin/env python3
"""Generate the tracked MathType LaTeX coverage corpus from the official page."""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import pathlib
import re
import sys
import urllib.request
from collections import defaultdict
from datetime import date


ROOT = pathlib.Path(__file__).resolve().parents[2]
REFERENCE_DIR = ROOT / "docs" / "reference" / "mathtype"
DEFAULT_HTML = REFERENCE_DIR / "raw" / "latex-coverage.html"
DEFAULT_OUTPUT = REFERENCE_DIR / "latex-coverage-corpus.json"
SOURCE_URL = "https://www.wiris.net/demo/editor/docs/latex-coverage/"

ENTRY_RE = re.compile(
    r'<h3><code><a name="([^"]+)">.*?</h3><table class="examples">(.*?)</table>',
    re.DOTALL,
)
EXAMPLE_RE = re.compile(r"<td><code>(.*?)</code></td>", re.DOTALL)
TAG_RE = re.compile(r"<[^>]+>")

STRUCTURE_COMMANDS = {
    r"\atop", r"\begin{array}", r"\begin{bmatrix}", r"\begin{displaymath}",
    r"\begin{math}", r"\begin{pmatrix}", r"\begin{vmatrix}", r"\binom",
    r"\boxed", r"\brace", r"\brack", r"\buildrel", r"\cfrac", r"\choose",
    r"\dbinom", r"\dfrac", r"\frac", r"\left", r"\over", r"\overbrace",
    r"\overbracket", r"\overleftarrow", r"\overleftrightarrow", r"\overline",
    r"\overrightarrow", r"\overset", r"\right", r"\root", r"\sideset",
    r"\sqrt", r"\stackrel", r"\tbinom", r"\tfrac", r"\underbrace",
    r"\underbracket", r"\underleftarrow", r"\underleftrightarrow",
    r"\underrightarrow", r"\underset", r"\xLeftrightarrow",
    r"\xLongleftarrow", r"\xLongleftrightarrow", r"\xLongrightarrow",
    r"\xcancel", r"\xleftarrow", r"\xleftrightarrow", r"\xlongequal",
    r"\xlongleftarrow", r"\xlongleftrightarrow", r"\xlongrightarrow",
    r"\xrightarrow",
}
STYLE_COMMANDS = {
    r"\Bbb", r"\Huge", r"\LARGE", r"\Large", r"\bf", r"\boldsymbol",
    r"\cal", r"\color", r"\displaystyle", r"\frak", r"\hspace", r"\huge",
    r"\it", r"\large", r"\mathbb", r"\mathbf", r"\mathcal", r"\mathfrak",
    r"\mathit", r"\mathrm", r"\mathsf", r"\mathtt", r"\mbox", r"\normalsize",
    r"\rm", r"\scriptscriptstyle", r"\scriptsize", r"\scriptstyle", r"\sf",
    r"\small", r"\style", r"\text", r"\textbf", r"\textit", r"\textrm",
    r"\textstyle", r"\tiny", r"\tt",
}
MODIFIER_COMMANDS = {
    r"\bar", r"\hat", r"\limits", r"\nolimits", r"\not", r"\of",
}


def fetch() -> bytes:
    request = urllib.request.Request(SOURCE_URL, headers={"User-Agent": "latextomathtype-coverage/1.0"})
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def clean_markup(value: str) -> str:
    return " ".join(html.unescape(TAG_RE.sub("", value)).split())


def classify(command: str) -> str:
    if command.startswith(r"\begin{"):
        return "environment"
    if command in STRUCTURE_COMMANDS:
        return "structure"
    if command in STYLE_COMMANDS:
        return "style"
    if command in MODIFIER_COMMANDS:
        return "modifier"
    return "symbol"


def infer_invocation(command: str, example: str) -> str:
    start = example.find(command)
    if start < 0:
        return command
    cursor = start + len(command)
    while cursor < len(example) and example[cursor].isspace():
        cursor += 1
    end = cursor
    for opening, closing in (("[", "]"), ("{", "}"), ("{", "}")):
        if end >= len(example) or example[end] != opening:
            continue
        depth = 0
        while end < len(example):
            char = example[end]
            if char == opening:
                depth += 1
            elif char == closing:
                depth -= 1
                if depth == 0:
                    end += 1
                    break
            end += 1
        while end < len(example) and example[end].isspace():
            end += 1
    return example[start:end].strip() or command


def parse_entries(raw_html: str) -> list[dict[str, object]]:
    entries: list[dict[str, object]] = []
    for index, match in enumerate(ENTRY_RE.finditer(raw_html), start=1):
        command = html.unescape(match.group(1))
        examples = [clean_markup(item) for item in EXAMPLE_RE.findall(match.group(2))]
        examples = [item for item in examples if item]
        if not examples:
            raise ValueError(f"official entry has no example: {command}")
        entries.append({
            "index": index,
            "command": command,
            "category": classify(command),
            "invocation": infer_invocation(command, examples[0]),
            "examples": examples,
        })
    if not entries:
        raise ValueError("no LaTeX coverage entries found")

    aliases: dict[tuple[str, ...], list[str]] = defaultdict(list)
    for entry in entries:
        aliases[tuple(entry["examples"])].append(str(entry["command"]))
    for entry in entries:
        entry["aliases"] = [
            alias for alias in aliases[tuple(entry["examples"])] if alias != entry["command"]
        ]
    return entries


def build_document(raw: bytes, retrieved_at: str) -> dict[str, object]:
    decoded = raw.decode("utf-8")
    entries = parse_entries(decoded)
    counts: dict[str, int] = defaultdict(int)
    for entry in entries:
        counts[str(entry["category"])] += 1
    return {
        "schemaVersion": 1,
        "sourceUrl": SOURCE_URL,
        "retrievedAt": retrieved_at,
        "sourceSha256": hashlib.sha256(raw).hexdigest(),
        "entryCount": len(entries),
        "categoryCounts": dict(sorted(counts.items())),
        "entries": entries,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--html", type=pathlib.Path, default=DEFAULT_HTML)
    parser.add_argument("--output", type=pathlib.Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--fetch", action="store_true", help="download and replace the local HTML snapshot")
    parser.add_argument("--check-live", action="store_true", help="fail when the official page differs from the snapshot")
    parser.add_argument("--retrieved-at", default=date.today().isoformat())
    args = parser.parse_args()

    if args.check_live:
        expected = json.loads(args.output.read_text(encoding="utf-8"))["sourceSha256"]
        actual = hashlib.sha256(fetch()).hexdigest()
        if actual != expected:
            print(f"official coverage drift: expected {expected}, got {actual}", file=sys.stderr)
            return 1
        print(f"official coverage unchanged: {actual}")
        return 0

    if args.fetch:
        raw = fetch()
        args.html.parent.mkdir(parents=True, exist_ok=True)
        args.html.write_bytes(raw)
    else:
        raw = args.html.read_bytes()

    document = build_document(raw, args.retrieved_at)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(document, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {document['entryCount']} entries to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
