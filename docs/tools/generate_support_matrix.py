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
    "TM_ANGLE": "writeAngleHeader",
    "TM_PAREN": "writeParenHeader",
    "TM_BRACE": "writeBraceHeader",
    "TM_BRACK": "writeBracketHeader",
    "TM_BAR": "writeBarHeader",
    "TM_DBAR": "writeDoubleBarHeader",
    "TM_FLOOR": "writeFloorHeader",
    "TM_CEILING": "writeCeilingHeader",
    "TM_OBRACK": "writeOpenBracketHeader",
    "TM_INTERVAL": "writeIntervalHeader",
    "TM_ROOT": "writeSqrtHeader/writeNthRootHeader",
    "TM_FRACT": "writeFractionHeader",
    "TM_UBAR": "writeUnderlineHeader",
    "TM_OBAR": "writeOverlineHeader",
    "TM_ARROW": "writeArrowHeader",
    "TM_INTEG": "writeIntegralHeader",
    "TM_SUM": "writeSumHeader",
    "TM_PROD": "writeProductHeader",
    "TM_COPROD": "writeCoproductHeader",
    "TM_UNION": "writeUnionHeader",
    "TM_INTER": "writeIntersectionHeader",
    "TM_INTOP": "writeIntegralStyleBigOpHeader",
    "TM_SUMOP": "writeSummationStyleBigOpHeader",
    "TM_LIM": "writeLimitHeader",
    "TM_HBRACE": "writeHBraceHeader",
    "TM_HBRACK": "writeHBrackHeader",
    "TM_LDIV": "writeLongDivisionHeader",
    "TM_SUB": "writeSubscriptHeader",
    "TM_SUP": "writeSuperscriptHeader",
    "TM_SUBSUP": "writeSubSuperscriptHeader",
    "TM_DIRAC": "writeDiracHeader",
    "TM_VEC": "writeVecHeader",
    "TM_TILDE": "writeTildeHeader",
    "TM_HAT": "writeHatHeader",
    "TM_ARC": "writeArcHeader",
    "TM_JSTATUS": "writeJointStatusHeader",
    "TM_STRIKE": "writeStrikeHeader",
    "TM_BOX": "writeBoxHeader",
}

writer_usage = {
    "TM_ANGLE": "writeParenFence / writeFenceTemplate",
    "TM_PAREN": "writeParenFence / writeCommandNode",
    "TM_BRACE": "writeCommandNode / writeCasesAsBraceFence",
    "TM_BRACK": "writeParenFence / writeCommandNode",
    "TM_BAR": "writeCommandNode / writeFenceTemplate",
    "TM_DBAR": "writeFenceTemplate",
    "TM_FLOOR": "writeFenceTemplate",
    "TM_CEILING": "writeFenceTemplate",
    "TM_OBRACK": "writeFenceTemplate",
    "TM_INTERVAL": "writeFenceTemplate",
    "TM_ROOT": "writeSqrtNode",
    "TM_FRACT": "writeFractionNode",
    "TM_UBAR": "writeCommandNode",
    "TM_OBAR": "writeCommandNode",
    "TM_ARROW": "writeArrowNode",
    "TM_INTEG": "writeBigOpHeader (integrals)",
    "TM_SUM": "writeBigOpHeader (sum)",
    "TM_PROD": "writeBigOpHeader (prod)",
    "TM_COPROD": "writeBigOpHeader (coprod)",
    "TM_UNION": "writeBigOpHeader (bigcup)",
    "TM_INTER": "writeBigOpHeader (bigcap)",
    "TM_INTOP": "writeBigOpHeader (intop)",
    "TM_SUMOP": "writeBigOpHeader (sumop-like)",
    "TM_LIM": "writeLimitComplete",
    "TM_HBRACE": "writeHorizontalBrace",
    "TM_HBRACK": "writeHorizontalBracket",
    "TM_LDIV": "writeLongDivisionNode",
    "TM_SUB": "writeSubscriptNode / writeLeadingScriptAttachment",
    "TM_SUP": "writeSuperscriptNode / writeLeadingScriptAttachment",
    "TM_SUBSUP": "writeSupSubAttachment / writeLeadingScriptAttachment",
    "TM_DIRAC": "writeDiracNode",
    "TM_VEC": "writeCommandNode",
    "TM_TILDE": "writeCommandNode",
    "TM_HAT": "writeCommandNode",
    "TM_ARC": "writeArcNode / writeCommandNode",
    "TM_JSTATUS": "writeCommandNode",
    "TM_STRIKE": "writeStrikeNode / writeCommandNode",
    "TM_BOX": "writeBoxNode / writeCommandNode",
}

parser_signals = {
    "linear commands": [
        (r"\\frac", r"\frac"),
        (r"\\sqrt", r"\sqrt"),
        (r"\\sum", r"\sum"),
        (r"\\int", r"\int"),
        (r"\\xrightarrow", r"\xrightarrow"),
        (r"\\xleftarrow", r"\xleftarrow"),
        (r"\\overbrace", r"\overbrace"),
        (r"\\overarc", r"\overarc"),
    ],
    "environments": [
        (r"matrix", "matrix"),
        (r"pmatrix", "pmatrix"),
        (r"bmatrix", "bmatrix"),
        (r"cases", "cases"),
        (r"aligned", "aligned"),
        (r"align\*?", "align / align*"),
        (r"split", "split"),
        (r"longdivision", "longdivision"),
    ],
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
    hits = [display for pattern, display in patterns if re.search(pattern, parser_text)]
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
lines.append("- Left-script / prescript input such as `{}^{a}x`, `{}_{b}x`, `{}_{b}^{a}x` is emitted in **MTEF v5 form** via `TM_SUB` / `TM_SUP` / `TM_SUBSUP` plus `TV_SU_PRECEDES`, rather than the legacy `TM_LSCRIPT(44)` selector.\n")
lines.append("- `\\xrightarrow` / `\\xleftarrow` now follow the official amsmath signature `\\xrightarrow[below]{above}` / `\\xleftarrow[below]{above}` for the supported single-arrow family.\n")
lines.append("- `TM_ARC` currently covers the documented over-arc family (`\\arc`, `\\overarc`, `\\overparen`, `\\wideparen`). `\\underarc` is intentionally not claimed here because it is not part of amsmath and no official MathType mapping has been confirmed in this repository yet.\n")
lines.append("- This matrix measures **code-path existence**, not semantic completeness. For example, integrals may be implemented while still missing full variation coverage (double/triple/loop), and some behaviors still rely on Linux-side byte-level validation instead of a fresh official Windows round-trip.\n")

OUT.write_text(''.join(lines), encoding='utf-8')
print(f"wrote {OUT}")
