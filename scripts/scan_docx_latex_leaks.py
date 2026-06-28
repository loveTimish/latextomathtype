from __future__ import annotations

import argparse
import json
import re
import struct
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
LEAK_PATTERNS = [
    r"\\textcolor",
    r"\\textbf",
    r"\\textit",
    r"\\begin\{equation\*?\}",
    r"\\end\{equation\*?\}",
    r"\\begin\{array\}",
    r"\\end\{array\}",
    r"\\pwmetrics",
    r"\\pwstyle",
    r"\\fbox",
    r"\\begin\{minipage\}",
    r"\\end\{minipage\}",
    r"\\includegraphics",
    r"\\centering",
    r"\\par\b",
    r"\\(?:frac|dfrac|tfrac|sqrt|left|right|times|div|cdot|sum|prod|int|lim|begin|end|overline|underline|overset|underset|overarc|boxed|cases|matrix|array|alpha|beta|gamma|delta|theta|lambda|mu|pi|sigma|omega|leq|geq|neq|infty|sin|cos|tan|log|ln|arcsin|arccos|arctan)\b",
    r"\$[^$\n]{1,500}\$",
    r"(?<![\w/.-])[A-Za-z][A-Za-z0-9]*[_^]\{?[-+A-Za-z0-9\\]+\}?",
]

CFB_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
END_OF_CHAIN = 0xFFFFFFFE
FREE_SECTOR = 0xFFFFFFFF


def sector_offset(sector: int, sector_size: int) -> int:
    return (sector + 1) * sector_size


def read_cfb_directory(data: bytes) -> tuple[dict[str, dict], int]:
    if len(data) < 512 or data[:8] != CFB_SIGNATURE:
        raise ValueError("not an OLE compound file")
    sector_shift = struct.unpack_from("<H", data, 30)[0]
    mini_sector_shift = struct.unpack_from("<H", data, 32)[0]
    sector_size = 1 << sector_shift
    mini_sector_size = 1 << mini_sector_shift
    dir_start = struct.unpack_from("<I", data, 48)[0]
    fat_sector_count = struct.unpack_from("<I", data, 44)[0]
    fat_sector_ids = [
        struct.unpack_from("<I", data, 76 + i * 4)[0]
        for i in range(min(fat_sector_count, 109))
    ]
    fat: list[int] = []
    for sid in fat_sector_ids:
        if sid in (FREE_SECTOR, END_OF_CHAIN):
            continue
        offset = sector_offset(sid, sector_size)
        if offset + sector_size > len(data):
            raise ValueError("FAT sector outside file")
        fat.extend(struct.unpack_from("<" + "I" * (sector_size // 4), data, offset))

    def read_chain(start: int, max_bytes: int | None = None) -> bytes:
        out = bytearray()
        seen: set[int] = set()
        sid = start
        while sid not in (END_OF_CHAIN, FREE_SECTOR):
            if sid in seen or sid >= len(fat):
                raise ValueError("invalid FAT chain")
            seen.add(sid)
            offset = sector_offset(sid, sector_size)
            if offset + sector_size > len(data):
                raise ValueError("stream sector outside file")
            out.extend(data[offset:offset + sector_size])
            if max_bytes is not None and len(out) >= max_bytes:
                break
            sid = fat[sid]
        return bytes(out[:max_bytes] if max_bytes is not None else out)

    dir_bytes = read_chain(dir_start)
    entries: dict[str, dict] = {}
    root = None
    for offset in range(0, len(dir_bytes) - 127, 128):
        entry = dir_bytes[offset:offset + 128]
        name_len = struct.unpack_from("<H", entry, 64)[0]
        entry_type = entry[66]
        if entry_type == 0 or name_len < 2 or name_len > 64:
            continue
        name = entry[:name_len - 2].decode("utf-16le", "replace")
        start = struct.unpack_from("<I", entry, 116)[0]
        size = struct.unpack_from("<Q", entry, 120)[0]
        item = {"type": entry_type, "start": start, "size": size}
        entries[name] = item
        if entry_type == 5:
            root = item

    if root is None:
        raise ValueError("missing Root Entry")
    mini_stream = read_chain(root["start"], root["size"]) if root["size"] else b""

    def read_stream(name: str) -> bytes:
        if name not in entries:
            raise ValueError(f"missing stream {name}")
        entry = entries[name]
        size = int(entry["size"])
        start = int(entry["start"])
        if size < 4096 and name != "Root Entry":
            chunks = bytearray()
            sid = start
            # Generated MathType OLE objects used here keep mini streams contiguous.
            while len(chunks) < size and sid not in (END_OF_CHAIN, FREE_SECTOR):
                offset = sid * mini_sector_size
                if offset + mini_sector_size > len(mini_stream):
                    raise ValueError("mini stream outside root stream")
                chunks.extend(mini_stream[offset:offset + mini_sector_size])
                sid += 1
            return bytes(chunks[:size])
        return read_chain(start, size)

    entries["__read_stream__"] = {"fn": read_stream}
    return entries, sector_size


def validate_mathtype_ole(data: bytes) -> tuple[bool, str]:
    entries, _sector_size = read_cfb_directory(data)
    read_stream = entries["__read_stream__"]["fn"]
    comp_obj = read_stream("\x01CompObj")
    if b"Equation.DSMT4" not in comp_obj:
        return False, "CompObj missing Equation.DSMT4"
    equation_native = read_stream("Equation Native")
    if len(equation_native) < 40:
        return False, "Equation Native too short"
    header_size = struct.unpack_from("<I", equation_native, 0)[0]
    mtef_length = struct.unpack_from("<I", equation_native, 8)[0]
    if header_size < 28 or header_size >= len(equation_native):
        return False, f"invalid Equation Native header size {header_size}"
    if header_size + mtef_length != len(equation_native):
        return False, "Equation Native cbHdr/cbObject length mismatch"
    mtef = equation_native[header_size:header_size + mtef_length]
    if len(mtef) < 12 or mtef[0] != 5 or mtef[1] != 1 or mtef[2] != 0 or b"DSMT" not in mtef[:16]:
        return False, "Equation Native MTEF header is not MathType DSMT"
    return True, ""


def visible_text(document_xml: bytes) -> str:
    root = ET.fromstring(document_xml)
    parts: list[str] = []
    for node in root.iter():
        if node.tag == f"{{{NS['w']}}}t":
            parts.append(node.text or "")
        elif node.tag == f"{{{NS['w']}}}tab":
            parts.append("\t")
        elif node.tag == f"{{{NS['w']}}}br":
            parts.append("\n")
        elif node.tag == f"{{{NS['w']}}}p":
            parts.append("\n")
    return "".join(parts)


def scan_docx(path: Path) -> dict:
    ole_validation = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        document_xml = z.read("word/document.xml")
        text = visible_text(document_xml)
        ole_entries = [n for n in names if n.startswith("word/embeddings/")]
        wmf_entries = [n for n in names if n.startswith("word/media/") and n.lower().endswith(".wmf")]
        xml_text = document_xml.decode("utf-8", "replace")
        for entry in ole_entries:
            try:
                valid, error = validate_mathtype_ole(z.read(entry))
            except Exception as exc:
                valid, error = False, str(exc)
            ole_validation.append({"entry": entry, "valid": valid, "error": error})
    leaks = []
    for pattern in LEAK_PATTERNS:
        for match in re.finditer(pattern, text):
            start = max(0, match.start() - 60)
            end = min(len(text), match.end() + 100)
            leaks.append(
                {
                    "pattern": pattern,
                    "offset": match.start(),
                    "context": re.sub(r"\s+", " ", text[start:end]).strip(),
                }
            )
    return {
        "path": str(path),
        "visibleTextChars": len(text),
        "leakCount": len(leaks),
        "leaks": leaks[:50],
        "oleEmbeddingCount": len(ole_entries),
        "wmfMediaCount": len(wmf_entries),
        "equationDsmt4Count": xml_text.count("Equation.DSMT4"),
        "oleObjectXmlCount": xml_text.count("OLEObject"),
        "hasEquationDsmt4": "Equation.DSMT4" in xml_text,
        "hasOleObject": "OLEObject" in xml_text,
        "validMathTypeOleCount": sum(1 for item in ole_validation if item["valid"]),
        "invalidMathTypeOleCount": sum(1 for item in ole_validation if not item["valid"]),
        "invalidMathTypeOleSamples": [item for item in ole_validation if not item["valid"]][:20],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--expect-ole-count", type=int)
    parser.add_argument("--expect-wmf-count", type=int)
    args = parser.parse_args()

    result = scan_docx(args.docx)
    failures = []
    if args.expect_ole_count is not None and result["oleEmbeddingCount"] != args.expect_ole_count:
        failures.append(
            f"oleEmbeddingCount {result['oleEmbeddingCount']} != expected {args.expect_ole_count}"
        )
    if args.expect_ole_count is not None and result["validMathTypeOleCount"] != args.expect_ole_count:
        failures.append(
            f"validMathTypeOleCount {result['validMathTypeOleCount']} != expected {args.expect_ole_count}"
        )
    if args.expect_wmf_count is not None and result["wmfMediaCount"] != args.expect_wmf_count:
        failures.append(f"wmfMediaCount {result['wmfMediaCount']} != expected {args.expect_wmf_count}")
    result["failures"] = failures
    data = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(data + "\n", encoding="utf-8")
    print(data)
    return 1 if result["leakCount"] or failures else 0


if __name__ == "__main__":
    sys.exit(main())
