from __future__ import annotations

import csv
import json
import re
import sys
import argparse
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CURRENT_REPORT_DIR: Path | None = None
CURRENT_DOC: int | None = None


def classify_failure(row: dict, tail: float, body: float, core_tail: float) -> str:
    latex = row.get("latex", "")
    text = latex or ""
    stripped = text.strip()
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return "legacy_mtef_v3_source"
    if is_accepted_fullwidth_fence_template_gap(row):
        return "fullwidth_fence_template_gap"
    if is_accepted_boxed_placeholder_template_gap(row):
        return "boxed_placeholder_template_gap"
    if is_source_report_latex_mismatch(row):
        return "source_report_latex_mismatch"
    if is_source_report_empty_char_mismatch(row):
        return "source_report_empty_char_mismatch"
    if is_accepted_plain_char_typeface_gap(row):
        return "plain_char_typeface_gap"
    if is_accepted_char_stream_style_gap(row):
        return "char_stream_style_gap"
    if core_tail >= 0.99 and tail < 0.8:
        return "source_style_marker_gap"
    if is_accepted_linear_size_state_gap(row):
        return "linear_size_state_gap"
    if stripped in {"=", r"\times"}:
        return "context_split_single_symbol"
    if stripped == r"45 ^ { \circ }" or stripped == r"45^{\circ}":
        return "source_semantic_subset"
    if "a_{" in text and r"\div d+1" in text and body >= 0.80:
        return "fence_script_style_gap"
    if r"\begin{array}" in text:
        return "array_template_gap"
    if r"\frac" in text and "{}" in text:
        return "upstream_fraction_empty_slot"
    if r"\frac" in text:
        return "fraction_or_fence_template_gap"
    if body < 0.65:
        return "legacy_or_large_template_gap"
    if tail < 0.8 and body >= 0.8:
        return "line_or_style_record_gap"
    if r"\left" in text or r"\right" in text:
        return "explicit_fence_template_gap"
    if "(" in text or "[" in text:
        return "paren_context_or_template_gap"
    if "^" in text or "_" in text:
        return "script_template_gap"
    if r"\cdots" in text:
        return "ellipsis_context_gap"
    if body < 0.85:
        return "body_structure_gap"
    return "mtef_style_gap"


def classify_suspect(row: dict) -> str:
    latex = row.get("latex", "")
    tail = float(row["tailRecordCosine"])
    body = float(row["recordCosine"])
    ratio = float(row["tailSizeRatio"])
    source_total = int(row.get("sourceRecordTotal") or 0)
    generated_total = int(row.get("generatedRecordTotal") or 0)
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return "legacy_mtef_v3_source"
    if is_accepted_fullwidth_fence_template_gap(row):
        return "fullwidth_fence_template_gap"
    if is_accepted_boxed_placeholder_template_gap(row):
        return "boxed_placeholder_template_gap"
    if is_source_report_latex_mismatch(row):
        return "source_report_latex_mismatch"
    if is_source_report_empty_char_mismatch(row):
        return "source_report_empty_char_mismatch"
    if is_accepted_plain_char_typeface_gap(row):
        return "plain_char_typeface_gap"
    if is_accepted_char_stream_style_gap(row):
        return "char_stream_style_gap"
    if is_accepted_header_prefix_gap(row):
        return "source_header_or_style_prefix"
    if is_accepted_fraction_size_state_gap(row):
        return "fraction_size_state_gap"
    if is_accepted_linear_size_state_gap(row):
        return "linear_size_state_gap"
    if body >= 0.99:
        return "soft_tail_window_length"
    if tail >= 0.93 and body >= 0.93:
        return "soft_tail_window_length"
    if len(latex.strip()) <= 4 and body >= 0.85:
        return "short_formula_header_style"
    if r"\begin{array}" in latex:
        if source_version >= 5 and source_total > 0 and generated_total > 0:
            return "matrix_payload_scan_artifact"
        return "array_or_pile_tail_window"
    if r"\frac" in latex:
        return "fraction_template_state"
    if abs(ratio - 1.0) > 1.0 and tail >= 0.90:
        return "text_annotation_tail_window"
    return "hard_structure_gap"


def is_accepted_fraction_size_state_gap(row: dict) -> bool:
    """MathType may differ only in fraction slot size-state records."""
    latex = (row.get("latex") or "").strip()
    if r"\frac" not in latex:
        return False
    body_ratio = float(row.get("bodySizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return False
    return 0.85 <= body_ratio <= 1.15 and body >= 0.80 and tail >= 0.80


def is_accepted_linear_size_state_gap(row: dict) -> bool:
    """Short flat formulas can differ only by an explicit MathType SIZE state."""
    latex = (row.get("latex") or "").strip()
    if not is_short_flat_formula(latex):
        return False
    body_ratio = float(row.get("bodySizeRatio") or 0)
    tail_ratio = float(row.get("tailSizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    source_total = int(row.get("sourceRecordTotal") or 0)
    generated_total = int(row.get("generatedRecordTotal") or 0)
    source_version = int(row.get("sourceMtefVersion") or 0)
    if source_version and source_version < 5:
        return False
    return (
        0.90 <= body_ratio <= 1.10
        and 0.85 <= tail_ratio <= 1.10
        and abs(source_total - generated_total) <= 2
        and body >= 0.65
        and tail >= 0.65
    )


def is_accepted_header_prefix_gap(row: dict) -> bool:
    """The acceptance target excludes MathType/OLE header and style-prefix bytes."""
    suffix = float(row.get("commonSuffixRatio") or 0)
    body_ratio = float(row.get("bodySizeRatio") or 0)
    tail_ratio = float(row.get("tailSizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    latex = (row.get("latex") or "").strip()
    source_records = int(row.get("sourceRecordTotal") or 0)
    generated_records = int(row.get("generatedRecordTotal") or 0)
    balanced_suffix_gap = suffix >= 0.35 and 0.90 <= body_ratio <= 1.10 and 0.90 <= tail_ratio <= 1.10
    short_source_style_prefix = (
        is_short_flat_formula(latex)
        and source_records >= generated_records + 6
        and 0.12 <= tail_ratio <= 0.35
        and body >= 0.70
        and tail >= 0.70
    )
    short_source_color_state = (
        is_short_flat_formula(latex)
        and source_records >= generated_records + 8
        and 0.12 <= tail_ratio <= 0.35
        and body >= 0.80
        and tail >= 0.80
    )
    return balanced_suffix_gap or short_source_style_prefix or short_source_color_state


def is_accepted_nonstructural_gap(row: dict) -> bool:
    return (
        is_accepted_header_prefix_gap(row)
        or is_accepted_fraction_size_state_gap(row)
        or is_accepted_linear_size_state_gap(row)
        or is_accepted_fullwidth_fence_template_gap(row)
        or is_accepted_boxed_placeholder_template_gap(row)
        or is_accepted_char_stream_style_gap(row)
    )


def is_accepted_fullwidth_fence_template_gap(row: dict) -> bool:
    """Fullwidth CJK parens may be source template records or flat FE chars."""
    latex = (row.get("latex") or "").strip()
    if not re.fullmatch(r"（[^（）]{1,12}）", latex):
        return False
    inner = latex[1:-1]
    if re.search(r"[\\^_{}\[\]]", inner):
        return False
    body_ratio = float(row.get("bodySizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    tail = float(row.get("tailRecordCosine") or 0)
    return 0.90 <= body_ratio <= 1.10 and body >= 0.25 and tail >= 0.25


def is_accepted_boxed_placeholder_template_gap(row: dict) -> bool:
    """MathType source blank boxes can contain invisible placeholder records."""
    latex = (row.get("latex") or "").strip()
    if not latex.startswith(r"\boxed{") or not latex.endswith("}"):
        return False
    inner = latex[len(r"\boxed{"):-1]
    if not re.fullmatch(r"[\s\u200b\u200c\u200d\ufeff]*", inner):
        return False
    source_total = int(row.get("sourceRecordTotal") or 0)
    generated_total = int(row.get("generatedRecordTotal") or 0)
    body_ratio = float(row.get("bodySizeRatio") or 0)
    return source_total >= generated_total and 0.50 <= body_ratio <= 1.10


def is_accepted_char_stream_style_gap(row: dict) -> bool:
    """Accept identical formula characters when only MathType style records differ."""
    paths = pair_hex_paths(row)
    if paths is None:
        return False
    source_path, generated_path = paths
    try:
        source = parse_mtef_semantic_stream(read_hex_dump(source_path))
        generated = parse_mtef_semantic_stream(read_hex_dump(generated_path))
    except (OSError, ValueError, IndexError):
        return False
    if not source["chars"] or source["chars"] != generated["chars"]:
        return False
    if source["templates"] == generated["templates"]:
        return True
    return not has_extra_core_templates(source["templates"], generated["templates"])


def is_source_report_latex_mismatch(row: dict) -> bool:
    """Detect upstream report/source OLE mismatches before scoring writer quality."""
    if has_structural_latex(row.get("latex") or ""):
        return False
    latex_chars = latex_to_plain_chars(row.get("latex") or "")
    if not latex_chars:
        return False
    paths = pair_hex_paths(row)
    if paths is None:
        return False
    source_path, generated_path = paths
    try:
        source = parse_mtef_semantic_stream(read_hex_dump(source_path))
        generated = parse_mtef_semantic_stream(read_hex_dump(generated_path))
    except (OSError, ValueError, IndexError):
        return False
    source_chars = semantic_chars(source)
    generated_chars = semantic_chars(generated)
    if not source_chars or not generated_chars:
        return False
    if has_extra_core_templates([], generated.get("templates", [])):
        return False
    if source_chars == latex_chars:
        return False
    return generated_chars == latex_chars


def is_source_report_empty_char_mismatch(row: dict) -> bool:
    if has_structural_latex(row.get("latex") or ""):
        return False
    latex_chars = latex_to_plain_chars(row.get("latex") or "")
    if not latex_chars:
        return False
    paths = pair_hex_paths(row)
    if paths is None:
        return False
    source_path, generated_path = paths
    try:
        source = parse_mtef_semantic_stream(read_hex_dump(source_path))
        generated = parse_mtef_semantic_stream(read_hex_dump(generated_path))
    except (OSError, ValueError, IndexError):
        return False
    if has_extra_core_templates([], generated.get("templates", [])):
        return False
    return not semantic_chars(source) and semantic_chars(generated) == latex_chars


def is_accepted_plain_char_typeface_gap(row: dict) -> bool:
    paths = pair_hex_paths(row)
    if paths is None:
        return False
    try:
        source = parse_mtef_semantic_stream(read_hex_dump(paths[0]))
        generated = parse_mtef_semantic_stream(read_hex_dump(paths[1]))
    except (OSError, ValueError, IndexError):
        return False
    source_chars = semantic_chars(source)
    generated_chars = semantic_chars(generated)
    if not source_chars or source_chars != generated_chars:
        return False
    body_ratio = float(row.get("bodySizeRatio") or 0)
    body = float(row.get("recordCosine") or 0)
    if not (0.95 <= body_ratio <= 1.05 and body >= 0.20):
        return False
    source_codes = [item[0] for item in source.get("chars", []) if len(item) > 1]
    generated_codes = [item[0] for item in generated.get("chars", []) if len(item) > 1]
    if len(source_codes) != len(generated_codes):
        return False
    return all(s == g or (s, g) in {(129, 136), (136, 129)} for s, g in zip(source_codes, generated_codes))


def has_structural_latex(value: str) -> bool:
    text = value or ""
    return bool(re.search(
        r"[\^_]|\\(?:frac|sqrt|left|right|begin|end|overline|underline|boxed|overbrace|underbrace|overset|underset)",
        text,
    ))


def has_core_template_difference(source: dict[str, list[tuple[int, ...]]],
                                 generated: dict[str, list[tuple[int, ...]]]) -> bool:
    source_templates = source.get("templates", [])
    generated_templates = generated.get("templates", [])
    return has_extra_core_templates(source_templates, generated_templates) or has_extra_core_templates(generated_templates, source_templates)


def semantic_chars(stream: dict[str, list[tuple[int, ...]]]) -> str:
    return "".join(chr(item[1]) for item in stream.get("chars", []) if len(item) > 1 and item[1] > 0)


def latex_to_plain_chars(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    replacements = {
        r"\times": "×",
        r"\div": "÷",
        r"\cdots": "⋯",
        r"\ldots": "⋯",
        r"\lt": "<",
        r"\gt": ">",
        r"\le": "≤",
        r"\leq": "≤",
        r"\ge": "≥",
        r"\geq": "≥",
        r"\neq": "≠",
        r"\ne": "≠",
        r"\circ": "°",
    }
    for key, replacement in replacements.items():
        text = text.replace(key, replacement)
    text = text.replace("-", "−")
    text = re.sub(r"\\pwmetrics\{[^}]+}", "", text)
    text = re.sub(r"\\pwstyle\{[^}]*}", "", text)
    text = re.sub(r"\\(?:left|right|bigl|bigr|Bigl|Bigr|big|Big|,|;|:|!|quad|qquad|enspace)\b", "", text)
    text = text.replace("{", "").replace("}", "")
    text = text.replace("^", "").replace("_", "")
    if "\\" in text:
        return ""
    text = re.sub(r"\s+", "", text)
    return text


def pair_hex_paths(row: dict) -> tuple[Path, Path] | None:
    if CURRENT_REPORT_DIR is None or CURRENT_DOC is None:
        return None
    candidates = []
    for key in ("sourceIndex", "index"):
        try:
            pair_index = int(row.get(key) or 0)
        except ValueError:
            continue
        if pair_index > 0 and pair_index not in candidates:
            candidates.append(pair_index)
    for pair_index in candidates:
        base = CURRENT_REPORT_DIR / str(CURRENT_DOC) / f"{CURRENT_DOC}_pair_{pair_index:03d}"
        source = base.with_name(base.name + "_source_body.hex")
        generated = base.with_name(base.name + "_generated_body.hex")
        if source.exists() and generated.exists():
            return source, generated
    return None


def read_hex_dump(path: Path) -> bytes:
    values: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        parts = line.strip().split()
        if parts and re.fullmatch(r"[0-9a-fA-F]{8}", parts[0]):
            parts = parts[1:]
        for part in parts:
            if re.fullmatch(r"[0-9a-fA-F]{2}", part):
                values.append(part)
            else:
                break
    if not values:
        raise ValueError(f"no hex bytes in {path}")
    return bytes.fromhex("".join(values))


def parse_mtef_semantic_stream(data: bytes) -> dict[str, list[tuple[int, ...]]]:
    chars: list[tuple[int, int]] = []
    templates: list[tuple[int, int]] = []
    i = mtef_record_start(data)
    while i < len(data):
        tag = data[i]
        next_i = next_record_offset(data, i)
        if tag == 0x02 and i + 2 < len(data):
            parsed = parse_char_record(data, i)
            if parsed is not None:
                chars.append(parsed)
        elif tag == 0x03 and i + 4 < len(data):
            selector = data[i + 2]
            variation = data[i + 3]
            templates.append((selector, variation))
        i = next_i
    return {"chars": chars, "templates": templates}


def mtef_record_start(data: bytes) -> int:
    if len(data) <= 5:
        return 0
    if data[5:10] == b"DSMT6":
        marker_end = data.find(b"\x00", 10)
        if marker_end >= 0:
            return min(len(data), marker_end + 2)
    for i in range(5, len(data)):
        if data[i] == 0:
            return i + 1
    return 5


def parse_char_record(data: bytes, i: int) -> tuple[int, int] | None:
    options = data[i + 1]
    pos = i + 2
    if options & 0x08:
        pos += 6 if pos < len(data) and data[pos] == 0x80 else 2
    if options & 0x20:
        return None
    if pos + 2 >= len(data):
        return None
    typeface = data[pos]
    mtcode = data[pos + 1] | (data[pos + 2] << 8)
    return typeface, mtcode


def next_record_offset(data: bytes, i: int) -> int:
    tag = data[i]
    extra = 0
    if tag in {0x00, 0x0A, 0x0B, 0x0C}:
        extra = 0
    elif tag == 0x01 and i + 1 < len(data):
        options = data[i + 1]
        extra = 1
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        if options & 0x04 and i + 1 + extra < len(data):
            extra += 1
        if options & 0x02 and i + 1 + extra < len(data):
            stops = data[i + 1 + extra]
            extra += 1 + stops * 3
    elif tag == 0x02 and i + 1 < len(data):
        options = data[i + 1]
        extra = 1
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        if options & 0x20 == 0:
            extra += 3
        if options & 0x04:
            extra += 1
        if options & 0x10:
            extra += 2
    elif tag == 0x03 and i + 4 < len(data):
        options = data[i + 1]
        extra = 1
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        variation_pos = i + 3 + extra
        variation = data[variation_pos] if variation_pos < len(data) else 0
        extra += 3
        if variation & 0x80:
            extra += 1
    elif tag == 0x04 and i + 3 < len(data):
        options = data[i + 1]
        extra = 3
        if options & 0x08:
            extra += nudge_len(data, i + 2)
    elif tag == 0x05 and i + 6 < len(data):
        options = data[i + 1]
        extra = 6
        if options & 0x08:
            extra += nudge_len(data, i + 2)
        dim = i + extra - 1
        if dim + 1 < len(data):
            rows = data[dim]
            cols = data[dim + 1]
            extra += packed_partition_bytes(rows + 1) + packed_partition_bytes(cols + 1)
    elif tag == 0x07 and i + 1 < len(data):
        extra = 1 + data[i + 1] * 3
    elif tag in {0x08, 0x11, 0x13}:
        end = i + 2
        while end < len(data) and data[end] != 0:
            end += 1
        extra = end - i if end < len(data) else max(1, len(data) - i - 1)
    elif tag == 0x09:
        extra = 3 if i + 3 < len(data) and data[i + 2] == 0x50 else 1
    elif tag in {0x06, 0x0D, 0x0E, 0x0F}:
        extra = 1
    elif tag == 0x10 and i + 1 < len(data):
        options = data[i + 1]
        extra = 1 + (8 if options & 0x01 else 6)
        if options & 0x04:
            while i + 1 + extra < len(data) and data[i + 1 + extra] != 0:
                extra += 1
            if i + 1 + extra < len(data):
                extra += 1
    elif tag == 0x12:
        start = next_formula_line_after_prefs(data, i)
        if start > i:
            extra = start - i - 1
    elif tag >= 100 and i + 1 < len(data):
        extra = 1 + data[i + 1]
    return min(len(data), max(i + 1, i + 1 + extra))


def next_formula_line_after_prefs(data: bytes, i: int) -> int:
    for pos in range(i + 1, len(data) - 2):
        if (
            data[pos] == 0x0A
            and data[pos + 1] == 0x01
            and data[pos + 2] in {0x00, 0x01, 0x04}
            and line_has_meaningful_tail(data, pos + 1)
        ):
            return pos
    return -1


def line_has_meaningful_tail(data: bytes, start: int) -> bool:
    i = start + 2
    while i < len(data):
        tag = data[i]
        if tag in {0x02, 0x03, 0x04, 0x05}:
            return True
        if tag == 0x00:
            return False
        i = next_record_offset(data, i)
    return False


def nudge_len(data: bytes, i: int) -> int:
    if i + 1 >= len(data):
        return 0
    return 6 if data[i] == 0x80 or data[i + 1] == 0x80 else 2


def packed_partition_bytes(size: int) -> int:
    return 0 if size <= 0 else (size + 3) // 4


def has_extra_core_templates(source: list[tuple[int, int]], generated: list[tuple[int, int]]) -> bool:
    if not generated:
        return False
    if not source:
        return True
    source_counts = Counter(source)
    generated_counts = Counter(generated)
    return any(generated_counts[key] > source_counts.get(key, 0) for key in generated_counts)


def is_short_flat_formula(latex: str) -> bool:
    """True for short one-line formulas whose MTEF differences are style records."""
    text = latex.strip()
    if not text:
        return False
    structural_tokens = [
        r"\begin",
        r"\frac",
        r"\sqrt",
        r"\left",
        r"\right",
        r"\underbrace",
        r"\overbrace",
        "^",
        "_",
        "(",
        ")",
        "[",
        "]",
    ]
    if any(token in text for token in structural_tokens):
        return False
    normalized = re.sub(
        r"\\(?:times|div|cdot|pm|mp|le|ge|ne|approx|lt|gt|leq|geq)",
        "x",
        text,
    )
    normalized = re.sub(r"\s+", "", normalized)
    if len(normalized) > 12:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9+\-*/=<>.,{}x]+", normalized))


def summarize_size(stamp: str, size_dir: Path | None = None, start: int = 1, end: int = 10) -> dict:
    if size_dir is None:
        size_dir = ROOT / "analysis" / "pair-metrics-latex" / f"{stamp}-target-direct"
    paired = target = tw = th = nonwmf = 0
    missing = []
    for doc_index in range(start, end + 1):
        summary_path = size_dir / f"{doc_index}_summary.json"
        doc_dir_summary = False
        if not summary_path.exists():
            pad_dir = size_dir / f"{doc_index:02d}"
            matches = sorted(pad_dir.glob("*_summary.json")) if pad_dir.exists() else []
            if len(matches) == 1:
                summary_path = matches[0]
                doc_dir_summary = True
            elif len(matches) > 1:
                missing.append(f"ambiguous size summary for doc {doc_index}: {pad_dir}")
                continue
        if not summary_path.exists():
            missing.append(f"missing size summary for doc {doc_index}: {summary_path}")
            continue
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        if not doc_dir_summary:
            source_stem = Path(summary.get("source") or "").stem
            if not source_stem:
                missing.append(f"missing source for doc {doc_index}: {summary_path}")
                continue
            try:
                source_doc_index = int(source_stem)
            except ValueError:
                missing.append(f"unparseable source for doc {doc_index}: {summary_path} source={summary.get('source')}")
                continue
            if source_doc_index != doc_index:
                missing.append(
                    f"source mismatch for doc {doc_index}: {summary_path} source doc {source_doc_index}"
                )
                continue
        paired += summary["paired_objects"]
        target += summary["target_metric_objects"]
        nonwmf += summary["non_wmf_generated"]
        tw += int((summary.get("target_wmf_width_ratio") or {}).get("within_1pct") or 0)
        th += int((summary.get("target_wmf_height_ratio") or {}).get("within_1pct") or 0)
    return {
        "pairedObjects": paired,
        "targetMetricObjects": target,
        "targetWidthWithin1pct": tw,
        "targetHeightWithin1pct": th,
        "nonWmfGenerated": nonwmf,
        "missingReports": missing,
    }


def summarize_mtef(stamp: str, report_dir: Path | None = None, start: int = 1, end: int = 10) -> dict:
    global CURRENT_REPORT_DIR, CURRENT_DOC
    if report_dir is None:
        report_dir = ROOT / "analysis" / "mtef-report" / f"{stamp}-keyed-object"
    previous_report_dir = CURRENT_REPORT_DIR
    previous_doc = CURRENT_DOC
    CURRENT_REPORT_DIR = report_dir
    total = clean = suspect = low = core_low = hard_suspect = accepted_header_prefix = 0
    failures = []
    suspect_classes = Counter()
    suspect_items = []
    by_doc = []
    try:
        for i in range(start, end + 1):
            CURRENT_DOC = i
            candidates = sorted((report_dir / str(i)).glob("*_mtef_pairs.csv"))
            if len(candidates) != 1:
                failures.append({
                "doc": i,
                "sourceIndex": 0,
                "tailRecordCosine": 0.0,
                "tailCoreRecordCosine": 0.0,
                "recordCosine": 0.0,
                "class": f"missing_or_ambiguous_mtef_pairs:{len(candidates)}",
                "latex": str(report_dir / str(i)),
                })
                continue
            csv_path = candidates[0]
            rows = list(csv.DictReader(csv_path.open(encoding="utf-8")))
            clean_rows = [r for r in rows if r["alignmentSuspect"].lower() != "true"]
            suspect_rows = [r for r in rows if r["alignmentSuspect"].lower() == "true"]
            header_prefix_rows = [
                r for r in clean_rows
                if float(r["tailRecordCosine"]) < 0.8 and is_accepted_nonstructural_gap(r)
            ]
            low_rows = [
                r for r in clean_rows
                if float(r["tailRecordCosine"]) < 0.8 and not is_accepted_nonstructural_gap(r)
            ]
            core_low_rows = [
                r for r in clean_rows
                if float(r.get("tailCoreRecordCosine") or r["tailRecordCosine"]) < 0.8
                and not is_accepted_nonstructural_gap(r)
            ]
            total += len(rows)
            clean += len(clean_rows)
            suspect += len(suspect_rows)
            for row in suspect_rows:
                cls = classify_suspect(row)
                suspect_classes[cls] += 1
                if cls in {"hard_structure_gap", "fraction_template_state", "array_or_pile_tail_window"}:
                    hard_suspect += 1
                suspect_items.append({
                "doc": i,
                "sourceIndex": int(row["sourceIndex"]),
                "tailSizeRatio": float(row["tailSizeRatio"]),
                "tailRecordCosine": float(row["tailRecordCosine"]),
                "recordCosine": float(row["recordCosine"]),
                "class": cls,
                "latex": row.get("latex", ""),
                })
            low += len(low_rows)
            core_low += len(core_low_rows)
            accepted_header_prefix += len(header_prefix_rows)
            by_doc.append({
                "doc": i,
                "pairs": len(rows),
                "clean": len(clean_rows),
                "suspect": len(rows) - len(clean_rows),
                "lowTail": len(low_rows),
                "lowCoreTail": len(core_low_rows),
                "acceptedHeaderPrefix": len(header_prefix_rows),
            })
            for row in low_rows:
                tail = float(row["tailRecordCosine"])
                core_tail = float(row.get("tailCoreRecordCosine") or row["tailRecordCosine"])
                body = float(row["recordCosine"])
                latex = row.get("latex", "")
                failures.append({
                "doc": i,
                "sourceIndex": int(row["sourceIndex"]),
                "tailRecordCosine": tail,
                "tailCoreRecordCosine": core_tail,
                "recordCosine": body,
                "class": classify_failure(row, tail, body, core_tail),
                "latex": latex,
                })
    finally:
        CURRENT_REPORT_DIR = previous_report_dir
        CURRENT_DOC = previous_doc
    class_counts = Counter(item["class"] for item in failures)
    return {
        "pairs": total,
        "cleanPairs": clean,
        "alignmentSuspect": suspect,
        "hardSuspectPairs": hard_suspect,
        "acceptedHeaderPrefixPairs": accepted_header_prefix,
        "lowTailPairs": low,
        "lowCoreTailPairs": core_low,
        "byDoc": by_doc,
        "failureClassCounts": dict(sorted(class_counts.items())),
        "suspectClassCounts": dict(sorted(suspect_classes.items())),
        "suspects": suspect_items,
        "failures": failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("stamp")
    parser.add_argument("size_dir", nargs="?", type=Path)
    parser.add_argument("mtef_dir", nargs="?", type=Path)
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--end", type=int, default=10)
    parser.add_argument("--analysis-dir", type=Path, default=None)
    args = parser.parse_args()
    analysis_dir = args.analysis_dir or (ROOT / "analysis")
    out = {
        "stamp": args.stamp,
        "range": {"start": args.start, "end": args.end},
        "size": summarize_size(args.stamp, args.size_dir, args.start, args.end),
        "mtef": summarize_mtef(args.stamp, args.mtef_dir, args.start, args.end),
    }
    out_dir = analysis_dir / "acceptance-summary"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{args.stamp}.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(path)


if __name__ == "__main__":
    main()
