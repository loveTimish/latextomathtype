from __future__ import annotations

import json
import re
import sys
from pathlib import Path


TAG_NAMES = {
    0x00: "END",
    0x01: "LINE",
    0x02: "CHAR",
    0x03: "TMPL",
    0x04: "PILE",
    0x05: "MATRIX",
    0x06: "EMBELL",
    0x07: "RULER",
    0x08: "FONT_STYLE_DEF",
    0x09: "SIZE",
    0x0A: "FULL",
    0x0B: "SUB",
    0x0C: "SUB2",
    0x0D: "SYM",
    0x0E: "SUBSYM",
    0x0F: "COLOR",
    0x10: "COLOR_DEF",
    0x11: "FONT_DEF",
    0x12: "EQN_PREFS",
    0x13: "ENCODING_DEF",
}


def read_hex(path: Path) -> bytes:
    text = path.read_text(encoding="utf-8")
    values: list[str] = []
    for line in text.splitlines():
        parts = line.strip().split()
        if not parts:
            continue
        if re.fullmatch(r"[0-9a-fA-F]{8}", parts[0]):
            parts = parts[1:]
        for part in parts:
            if re.fullmatch(r"[0-9a-fA-F]{2}", part):
                values.append(part)
            else:
                break
    if values:
        return bytes.fromhex("".join(values))
    return bytes.fromhex(re.sub(r"[^0-9a-fA-F]", "", text))


def nudge_len(data: bytes, i: int) -> int:
    if i + 1 >= len(data):
        return 0
    return 6 if data[i] == 0x80 or data[i + 1] == 0x80 else 2


def packed_partition_bytes(size: int) -> int:
    return 0 if size <= 0 else (size + 3) // 4


def record_payload_len(data: bytes, i: int) -> int:
    token, next_i = token_at(data, i)
    return max(0, next_i - i - 1)


def mtef_record_start(data: bytes) -> int:
    if len(data) <= 5:
        return 0
    for i in range(5, len(data)):
        if data[i] == 0:
            return i + 1
    return 5


def line_has_meaningful_tail(data: bytes, start: int) -> bool:
    i = start + 2
    while i < len(data):
        tag = data[i]
        if tag in {0x02, 0x03, 0x04, 0x05}:
            return True
        if tag == 0x00:
            return False
        i += 1 + record_payload_len(data, i)
    return False


def formula_content(data: bytes) -> bytes:
    i = mtef_record_start(data)
    while i < len(data):
        if (
            i + 1 < len(data)
            and data[i] == 0x01
            and data[i + 1] in {0x00, 0x01, 0x04}
            and line_has_meaningful_tail(data, i)
        ):
            return data[i:]
        i += 1 + record_payload_len(data, i)
    return data


def token_at(data: bytes, i: int) -> tuple[dict, int]:
    tag = data[i]
    token: dict = {"offset": i, "tag": tag, "name": TAG_NAMES.get(tag, f"REC_{tag}")}
    extra = 0
    if tag in {0x00, 0x0A, 0x0B, 0x0C}:
        pass
    elif tag == 0x01 and i + 1 < len(data):
        options = data[i + 1]
        extra = 1
        token["options"] = options
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        if options & 0x04 and i + 1 + extra < len(data):
            token["lineSpace"] = data[i + 1 + extra]
            extra += 1
        if options & 0x02 and i + 1 + extra < len(data):
            stops = data[i + 1 + extra]
            token["rulerStops"] = stops
            extra += 1 + stops * 3
        token["null"] = bool(options & 0x01)
    elif tag == 0x02 and i + 2 < len(data):
        options = data[i + 1]
        extra = 1
        token["options"] = options
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        pos = i + 1 + extra
        if options & 0x20 == 0 and pos + 2 < len(data):
            token["typeface"] = data[pos]
            token["mtcode"] = data[pos + 1] | (data[pos + 2] << 8)
            extra += 3
        pos = i + 1 + extra
        if options & 0x04 and pos < len(data):
            token["bits8"] = data[pos]
            extra += 1
        pos = i + 1 + extra
        if options & 0x10 and pos + 1 < len(data):
            token["bits16"] = data[pos] | (data[pos + 1] << 8)
            extra += 2
    elif tag == 0x03 and i + 4 < len(data):
        options = data[i + 1]
        extra = 1
        token["options"] = options
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        pos = i + 1 + extra
        token["selector"] = data[pos]
        extra += 1
        pos = i + 1 + extra
        variation = data[pos] if pos < len(data) else 0
        token["variation"] = variation
        extra += 1
        if variation & 0x80:
            extra += 1
        pos = i + 1 + extra
        token["templateOptions"] = data[pos] if pos < len(data) else None
        extra += 1
    elif tag == 0x04 and i + 3 < len(data):
        options = data[i + 1]
        extra = 1
        token["options"] = options
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        pos = i + 1 + extra
        if pos + 1 < len(data):
            token["halign"] = data[pos]
            token["valign"] = data[pos + 1]
            extra += 2
    elif tag == 0x05 and i + 6 < len(data):
        options = data[i + 1]
        extra = 1
        token["options"] = options
        if options & 0x08:
            extra += nudge_len(data, i + 1 + extra)
        dim = i + 1 + extra + 3
        extra += 5
        if dim + 1 < len(data):
            rows = data[dim]
            cols = data[dim + 1]
            token["rows"] = rows
            token["cols"] = cols
            extra += packed_partition_bytes(rows + 1)
            extra += packed_partition_bytes(cols + 1)
    elif tag == 0x07 and i + 2 < len(data):
        stops = data[i + 1]
        token["stops"] = stops
        extra = 1 + stops * 3
    elif tag == 0x08 and i + 1 < len(data):
        end = i + 2
        while end < len(data) and data[end] != 0:
            end += 1
        if end < len(data):
            extra = end - i
            token["style"] = data[i + 1]
            token["font"] = data[i + 2:end].decode("latin1", errors="replace")
        else:
            extra = len(data) - i - 1
    elif tag == 0x09:
        if i + 3 < len(data) and data[i + 2] == 0x50:
            extra = 3
            token["value"] = data[i + 1]
            token["sizeBytes"] = data[i + 1:i + 4].hex()
        else:
            extra = 1 if i + 1 < len(data) else 0
            if extra:
                token["value"] = data[i + 1]
    elif tag in {0x06, 0x0D, 0x0E, 0x0F}:
        extra = 1 if i + 1 < len(data) else 0
        if extra:
            token["value"] = data[i + 1]
    elif tag == 0x10 and i + 1 < len(data):
        options = data[i + 1]
        extra = 1 + (8 if options & 0x01 else 6)
        token["options"] = options
        if options & 0x04:
            while i + 1 + extra < len(data) and data[i + 1 + extra] != 0:
                extra += 1
            if i + 1 + extra < len(data):
                extra += 1
    elif tag in {0x11, 0x13}:
        end = i + 2
        while end < len(data) and data[end] != 0:
            end += 1
        extra = end - i if end < len(data) else max(1, len(data) - i - 1)
        if i + 1 < len(data):
            token["enc"] = data[i + 1]
        if i + 2 <= end <= len(data):
            token["text"] = data[i + 2:end].decode("latin1", errors="replace")
    elif tag >= 100 and i + 1 < len(data):
        length = data[i + 1]
        extra = 1 + length
        token["length"] = length
    end = min(len(data), i + 1 + extra)
    token["hex"] = data[i:end].hex()
    return token, max(i + 1, end)


def tokenize(data: bytes) -> list[dict]:
    tokens = []
    i = 0
    while i < len(data):
        token, i = token_at(data, i)
        tokens.append(token)
    return tokens


def sig(token: dict) -> str:
    name = token["name"]
    if name == "CHAR":
        return f"CHAR:{token.get('typeface')}:{token.get('mtcode')}"
    if name == "TMPL":
        return f"TMPL:{token.get('selector')}:{token.get('variation')}"
    if name == "LINE":
        return f"LINE:null={token.get('null')}"
    return name


def main() -> None:
    if len(sys.argv) != 4:
        raise SystemExit("usage: compare_mtef_sequence.py source_body.hex generated_body.hex out.json")
    source_bytes = read_hex(Path(sys.argv[1]))
    generated_bytes = read_hex(Path(sys.argv[2]))
    source_content = formula_content(source_bytes)
    generated_content = formula_content(generated_bytes)
    source = tokenize(source_content)
    generated = tokenize(generated_content)
    source_sig = [sig(t) for t in source]
    gen_sig = [sig(t) for t in generated]
    prefix = 0
    while prefix < min(len(source_sig), len(gen_sig)) and source_sig[prefix] == gen_sig[prefix]:
        prefix += 1
    suffix = 0
    while (
        suffix < min(len(source_sig), len(gen_sig)) - prefix
        and source_sig[len(source_sig) - 1 - suffix] == gen_sig[len(gen_sig) - 1 - suffix]
    ):
        suffix += 1
    out = {
        "sourceTokenCount": len(source),
        "generatedTokenCount": len(generated),
        "sourceBodyBytes": len(source_bytes),
        "generatedBodyBytes": len(generated_bytes),
        "sourceContentOffset": len(source_bytes) - len(source_content),
        "generatedContentOffset": len(generated_bytes) - len(generated_content),
        "sourceContentBytes": len(source_content),
        "generatedContentBytes": len(generated_content),
        "commonPrefixTokens": prefix,
        "commonSuffixTokens": suffix,
        "sourceWindow": source[max(0, prefix - 8): min(len(source), prefix + 24)],
        "generatedWindow": generated[max(0, prefix - 8): min(len(generated), prefix + 24)],
    }
    path = Path(sys.argv[3])
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(path)


if __name__ == "__main__":
    main()
