#!/usr/bin/env python3
from __future__ import annotations

import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[2]
MTEF_RECORD = ROOT / "src/main/java/com/lz/paperword/core/mtef/MtefRecord.java"
TEMPLATE_BUILDER = ROOT / "src/main/java/com/lz/paperword/core/mtef/MtefTemplateBuilder.java"
WRITER = ROOT / "src/main/java/com/lz/paperword/core/mtef/MtefWriter.java"
PARSER = ROOT / "src/main/java/com/lz/paperword/core/latex/LaTeXParser.java"
OUT = ROOT / "docs/MathType-support-matrix.md"

record_text = MTEF_RECORD.read_text(encoding="utf-8")
template_text = TEMPLATE_BUILDER.read_text(encoding="utf-8")
writer_text = WRITER.read_text(encoding="utf-8")
parser_text = PARSER.read_text(encoding="utf-8")

TM_RE = re.compile(r"public static final int (TM_[A-Z0-9_]+)\s*=\s*(\d+);")

implemented_headers = {
    "TM_FRACT": "writeFractionHeader",
    "TM_ROOT": "writeSqrtHeader/writeNthRootHeader",
    "TM_SUP": "writeSuperscriptHeader",
    "TM_SUB": "writeSubscriptHeader",
    "TM_SUBSUP": "writeSubSuperscriptHeader",
    "TM_SUM": "writeSumHeader",
    "TM_INTEG": "writeIntegralHeader",
    "TM_PROD": "writeProductHeader",
    "TM_LDIV": "writeLongDivisionHeader",
    "TM_PAREN": "writeParenHeader",
    "TM_BRACK": "writeBracketHeader",
    "TM_BRACE": "writeBraceHeader",
    "TM_BAR": "writeBarHeader",
    "TM_OBAR": "writeOverlineHeader",
    "TM_UBAR": "writeUnderlineHeader",
    "TM_VEC": "writeVecHeader",
    "TM_HAT": "writeHatHeader",
    "TM_TILDE": "writeTildeHeader",
}

writer_usage = {
    "TM_FRACT": "writeFractionNode",
    "TM_ROOT": "writeSqrtNode",
    "TM_SUP": "writeSuperscriptNode / writeSupSubAttachment",
    "TM_SUB": "writeSubscriptNode / writeSupSubAttachment",
    "TM_SUBSUP": "writeSuperscriptNode / writeSupSubAttachment",
    "TM_SUM": "writeBigOpHeader (sum-like fallback)",
    "TM_INTEG": "writeBigOpHeader (integrals)",
    "TM_PROD": "writeBigOpHeader (prod)",
    "TM_LDIV": "writeLongDivisionNode",
    "TM_PAREN": "writeParenFence / writeCommandNode",
    "TM_BRACK": "writeParenFence / writeCommandNode",
    "TM_BRACE": "writeCommandNode",
    "TM_BAR": "writeCommandNode",
    "TM_OBAR": "writeCommandNode",
    "TM_UBAR": "writeCommandNode",
    "TM_VEC": "writeCommandNode",
    "TM_HAT": "writeCommandNode",
    "TM_TILDE": "writeCommandNode",
}

parser_signals = {
    "linear": [r"case \\\"\\\\frac", r"case \\\"\\\\sqrt", r"case \\\"\\\\sum", r"case \\\"\\\\int", r"case \\\"\\\\left"],
    "environments": [r"matrix", r"pmatrix", r"bmatrix", r"cases", r"aligned", r"split", r"longdivision"],
}

rows = []
for name, value in TM_RE.findall(record_text):
    header = implemented_headers.get(name, "")
    usage = writer_usage.get(name, "")
    if header and usage:
        status = "implemented"
    elif header:
        status = "builder-only"
    else:
        status = "declared-only"
    rows.append((name, value, status, header or "—", usage or "—"))

rows.sort(key=lambda item: int(item[1]))

summary = {
    "declared": len(rows),
    "implemented": sum(1 for _, _, status, _, _ in rows if status == "implemented"),
    "builder_only": sum(1 for _, _, status, _, _ in rows if status == "builder-only"),
    "declared_only": sum(1 for _, _, status, _, _ in rows if status == "declared-only"),
}

lines = []
lines.append("# MathType support matrix\n")
lines.append("\n")
lines.append("This file is generated from the current codebase and then manually curated as implementation progresses.\n")
lines.append("\n")
lines.append("## Snapshot\n")
lines.append(f"- Declared official TM_* templates in `MtefRecord`: **{summary['declared']}**\n")
lines.append(f"- Templates with both builder + writer path: **{summary['implemented']}**\n")
lines.append(f"- Templates with builder helper only: **{summary['builder_only']}**\n")
lines.append(f"- Templates only declared in constants: **{summary['declared_only']}**\n")
lines.append("\n")
lines.append("## Parser coverage signals\n")
for label, patterns in parser_signals.items():
    hits = [p for p in patterns if re.search(p, parser_text)]
    lines.append(f"- {label}: {', '.join(hits) if hits else 'no direct hits'}\n")
lines.append("\n")
lines.append("## Official MTEF template coverage\n")
lines.append("| Template | Selector | Status | Builder path | Writer path |\n")
lines.append("|---|---:|---|---|---|\n")
for name, value, status, header, usage in rows:
    lines.append(f"| `{name}` | {value} | {status} | `{header}` | `{usage}` |\n")
lines.append("\n")
lines.append("## Notes\n")
lines.append("- `implemented` means the repository contains both a template builder helper and an explicit writer path using that template family.\n")
lines.append("- `builder-only` means helper code exists but no explicit writer use was mapped in this first pass.\n")
lines.append("- `declared-only` means the official template constant exists in `MtefRecord`, but no builder/writer path was found yet.\n")
lines.append("- This matrix measures **code-path existence**, not semantic completeness. For example, integrals may be implemented while still missing full variation coverage (double/triple/loop).\n")

OUT.write_text(''.join(lines), encoding='utf-8')
print(f"wrote {OUT}")
