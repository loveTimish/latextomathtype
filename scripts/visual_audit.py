# -*- coding: utf-8 -*-
"""
Render a DOCX containing MathType OLE+WMF previews and ask a vision LLM to
review the rendered Word output.

The source of truth remains the DOCX/WMF structure. PNG files are only the
visual transport format for the LLM.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

DEFAULT_API_BASE = "https://api.lingzhiai.cloud/v1"
DEFAULT_MODEL_CANDIDATES = (
    "gpt-4o",
    "gpt-4o-mini",
    "gpt-4.1",
    "gpt-4.1-mini",
    "qwen-vl-max",
    "qwen-vl-plus",
)

AUDIT_PROMPT = """\
你是一名数学排版验收员。图片来自 Word 中的 MathType OLE 公式对象，公式预览图来自 DOCX 内嵌 WMF。

请只评估视觉结果是否适合验收，不要猜测二进制结构。重点检查：
1. 根号是否完整、长短是否合理、是否压住内容。
2. 分数线是否缺失、断裂、太短、太粗或位置异常。
3. 箭头是否像箭头，是否退化为短横线、问号或乱码。
4. 上下标、括号、数组、中文混排是否拥挤或错位。
5. 整体是否接近 Word/MathType 的自然排版。

请把问题说具体，不允许只写“略拥挤”或“间距偏紧”。每个问题必须指出：
- 哪一道 case，例如 case 9、case 12；如果无法确定，写页面上的行号或邻近标签。
- 具体位置，例如“外层根号左腿和内层根号左腿之间”“分数线到分子 a 的上方间距”“根号横线右端到内容右边界”。
- 观察到的视觉事实，例如“两个竖向笔画几乎贴在一起”“分数线看起来像短横线”“分子贴近根号横线”。
- 期望的自然排版效果，例如“应有可见留白”“分数线应覆盖分子分母且左右略出头”“内外根号层次应分开”。
- 是否阻塞验收，以及建议下一步优先改哪里。

只返回 JSON，不要返回 markdown：
{
  "score": 1-10,
  "verdict": "pass|review|fail",
  "issues": [
    {
      "case": "case number or label",
      "type": "sqrt|fraction|arrow|spacing|font|baseline|other",
      "severity": "low|medium|high",
      "location": "具体视觉位置",
      "observed": "观察到的事实",
      "expected": "自然排版应达到的状态",
      "blocking": true/false,
      "suggestion": "下一步优先调整什么",
      "description": "一句完整描述"
    }
  ],
  "blocking_cases": ["..."],
  "summary": "一句话总结"
}
"""


def request_json(url: str, api_key: str, payload: dict[str, Any], timeout: int = 120) -> dict[str, Any]:
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        return {"error": f"HTTP {exc.code}", "detail": detail}
    except Exception as exc:  # noqa: BLE001 - command-line tool reports failures in JSON.
        return {"error": str(exc)}


def list_models(api_base: str, api_key: str) -> list[str]:
    req = urllib.request.Request(
        f"{api_base.rstrip('/')}/models",
        headers={"Authorization": f"Bearer {api_key}"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return []
    models = []
    for item in body.get("data", []):
        model_id = item.get("id")
        if isinstance(model_id, str):
            models.append(model_id)
    return models


def encode_image_b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    import win32com.client as win32

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.SaveAs2(str(pdf_path), FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()


def pdf_to_pngs(pdf_path: Path, out_dir: Path, dpi: int) -> list[Path]:
    import fitz

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    try:
        matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pages = []
        for index, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            path = out_dir / f"page-{index:03d}.png"
            pix.save(str(path))
            pages.append(path)
        return pages
    finally:
        doc.close()


def extract_json(text: str) -> dict[str, Any]:
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = stripped.split("```", 2)[1]
        if stripped.lstrip().startswith("json"):
            stripped = stripped.lstrip()[4:]
    start = stripped.find("{")
    end = stripped.rfind("}")
    if start >= 0 and end > start:
        stripped = stripped[start:end + 1]
    return json.loads(stripped)


def audit_page(api_base: str, api_key: str, model: str, png: Path) -> dict[str, Any]:
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": AUDIT_PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{encode_image_b64(png)}"},
                    },
                ],
            }
        ],
        "temperature": 0.1,
        "max_tokens": 2200,
    }
    response = request_json(f"{api_base.rstrip('/')}/chat/completions", api_key, payload)
    if "error" in response:
        return {"page": png.name, "error": response}
    content = response.get("choices", [{}])[0].get("message", {}).get("content", "")
    try:
        audit = extract_json(content)
    except Exception as exc:  # noqa: BLE001 - preserve model output for diagnosis.
        return {"page": png.name, "parse_error": str(exc), "raw": content}
    return {"page": png.name, "audit": audit}


def probe_model(api_base: str, api_key: str, requested_model: str | None) -> str:
    if requested_model:
        return requested_model
    available = list_models(api_base, api_key)
    for candidate in DEFAULT_MODEL_CANDIDATES:
        if not available or candidate in available:
            return candidate
    return available[0]


def main() -> int:
    parser = argparse.ArgumentParser(description="LLM visual audit for DOCX MathType OLE+WMF output")
    parser.add_argument("docx", help="DOCX generated by the project OLE+WMF path")
    parser.add_argument("--api-base", default=DEFAULT_API_BASE)
    parser.add_argument("--api-key", default=os.environ.get("LINGZHI_API_KEY"))
    parser.add_argument("--model", default=None)
    parser.add_argument("--out-dir", default=None)
    parser.add_argument("--dpi", type=int, default=180)
    parser.add_argument("--max-pages", type=int, default=0)
    parser.add_argument("--reuse-rendered", action="store_true")
    args = parser.parse_args()

    if not args.api_key:
        print("Missing API key. Pass --api-key or set LINGZHI_API_KEY.", file=sys.stderr)
        return 2

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"DOCX not found: {docx_path}", file=sys.stderr)
        return 2

    out_dir = Path(args.out_dir).resolve() if args.out_dir else docx_path.parent / "visual-audit"
    pages_dir = out_dir / "pages"
    out_dir.mkdir(parents=True, exist_ok=True)

    model = probe_model(args.api_base, args.api_key, args.model)
    pdf_path = out_dir / f"{docx_path.stem}.pdf"

    if not args.reuse_rendered or not pdf_path.exists():
        print(f"[1/4] Word renders DOCX OLE+WMF previews to PDF: {pdf_path}")
        docx_to_pdf(docx_path, pdf_path)
    else:
        print(f"[1/4] Reusing PDF: {pdf_path}")

    if not args.reuse_rendered or not any(pages_dir.glob("page-*.png")):
        print(f"[2/4] Rendering PDF pages to PNG for LLM transport: {pages_dir}")
        pages = pdf_to_pngs(pdf_path, pages_dir, args.dpi)
    else:
        pages = sorted(pages_dir.glob("page-*.png"))
        print(f"[2/4] Reusing PNG pages: {pages_dir}")

    if args.max_pages > 0:
        pages = pages[:args.max_pages]

    print(f"[3/4] Auditing {len(pages)} page(s) with model: {model}")
    results = []
    for index, page in enumerate(pages, 1):
        print(f"  page {index}/{len(pages)} {page.name}")
        results.append(audit_page(args.api_base, args.api_key, model, page))
        time.sleep(0.5)

    report = {
        "source_docx": str(docx_path),
        "pdf": str(pdf_path),
        "pages_dir": str(pages_dir),
        "api_base": args.api_base,
        "model": model,
        "results": results,
    }
    report_path = out_dir / "visual-audit-report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[4/4] Report: {report_path}")

    failures = [
        item for item in results
        if item.get("error") or item.get("parse_error")
        or item.get("audit", {}).get("verdict") in {"review", "fail"}
    ]
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
