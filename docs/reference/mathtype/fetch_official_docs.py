#!/usr/bin/env python3
from __future__ import annotations

import pathlib
import re
import subprocess

ROOT = pathlib.Path(__file__).resolve().parent
RAW_DIR = ROOT / "raw"
TEXT_DIR = ROOT / "text"
RAW_DIR.mkdir(parents=True, exist_ok=True)
TEXT_DIR.mkdir(parents=True, exist_ok=True)

DOCS = [
    (
        "mtef-v5",
        "https://docs.wiris.com/en_US/mathtype-sdk/mathtype-mtef-v5-mathtype-40-and-later",
        "MathType MTEF v5 format reference",
    ),
    (
        "mtef-storage",
        "https://docs.wiris.com/en_US/mathtype-sdk/how-mtef-is-stored-in-files-and-objects",
        "How MathType stores MTEF in OLE/files",
    ),
    (
        "latex-support",
        "https://docs.wiris.com/en_US/mathtype-web-interface-features/latex-support",
        "Official LaTeX support notes",
    ),
    (
        "mathml-coverage",
        "https://docs.wiris.com/en_US/mathtype-web-interface-features/mathml-coverage-by-mathtype",
        "Official MathML coverage notes",
    ),
    (
        "latex-coverage",
        "http://www.wiris.net/demo/editor/docs/latex-coverage/",
        "MathType web LaTeX coverage sample page",
    ),
    (
        "examples",
        "https://www.wiris.net/demo/editor/examples/",
        "MathType SDK example gallery",
    ),
    (
        "sitemap",
        "https://docs.wiris.com/sitemap.xml",
        "Wiris docs sitemap for discovery",
    ),
]

TAG_RE = re.compile(r"<script.*?</script>|<style.*?</style>|<[^>]+>", re.S | re.I)
SPACE_RE = re.compile(r"\s+")


def extract_text(html: str) -> str:
    html = re.sub(r"(?i)<br\s*/?>", "\n", html)
    html = re.sub(r"(?i)</p>", "\n\n", html)
    html = re.sub(r"(?i)</li>", "\n", html)
    text = TAG_RE.sub(" ", html)
    text = (
        text.replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&#39;", "'")
        .replace("&quot;", '"')
    )
    lines = []
    for line in text.splitlines():
        line = SPACE_RE.sub(" ", line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines) + "\n"


manifest_lines = []
for slug, url, desc in DOCS:
    result = subprocess.run(
        [
            "curl",
            "-L",
            "--insecure",
            "--retry",
            "3",
            "--retry-delay",
            "2",
            "--connect-timeout",
            "20",
            "--max-time",
            "60",
            "-A",
            "Mozilla/5.0 OpenClaw/1.0",
            url,
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not result.stdout.strip():
        manifest_lines.append(f"- FAIL {slug}: {url} :: curl exit={result.returncode}")
        print(f"failed {slug}: curl exit={result.returncode}")
        continue
    html = result.stdout
    (RAW_DIR / f"{slug}.html").write_text(html, encoding="utf-8")
    (TEXT_DIR / f"{slug}.txt").write_text(extract_text(html), encoding="utf-8")
    manifest_lines.append(f"- OK {slug}: {url} :: {desc}")
    print(f"saved {slug}: {desc}")

(ROOT / "manifest.txt").write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")
