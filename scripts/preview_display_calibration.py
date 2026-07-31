# -*- coding: utf-8 -*-
"""
Display-box calibration: measure what the Word v:shape display size SHOULD
be for every corpus formula, relative to the raw MathJax box, and emit the
per-bucket scale factors that replace the hand-tuned heuristic scales in
LaTeXImageRenderer.calibratePreviewMetrics.

Five-box discipline: GT display box = the REFERENCE docx's own Word v:shape
style width/height (not the WMF placeable bbox, not ink). MathJax box = raw
worker output at the renderer's own parameters (fontPt 9.02, exRatio 0.431,
padding 2.3pt, maxWidth 400pt). Ratios are per-formula; bucket statistics
are medians/quantiles computed by this script only.

Inputs:
    rebuild-assets/external/fraction-split-reference.docx   (GT v:shape)
    target/corpus-preview.docx                              (our v:shape)
    target/wmf-ruler/wmf-latex.tsv                          (label<TAB>latex)
    target/wmf-ruler/structure-measure.jsonl                (image -> bucket)

Outputs:
    target/wmf-ruler/display-scales.json
    docs/preview-display-calibration.md
"""
from __future__ import annotations

import base64
import json
import math
import re
import statistics
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "rebuild"))
from extract_formula_boxes import extract_boxes  # noqa: E402

REF_DOCX = ROOT / "rebuild-assets" / "external" / "fraction-split-reference.docx"
GEN_DOCX = ROOT / "target" / "corpus-preview.docx"
TSV = ROOT / "target" / "wmf-ruler" / "wmf-latex.tsv"
MEASURE = ROOT / "target" / "wmf-ruler" / "structure-measure.jsonl"
MJ_CACHE = ROOT / "target" / "wmf-ruler" / "mj-boxes.json"
OUT_JSON = ROOT / "target" / "wmf-ruler" / "display-scales.json"
OUT_MD = ROOT / "docs" / "preview-display-calibration.md"

WORKER = ROOT / "tools" / "mathjax" / "render_mathjax_svg.cjs"
MJ_PARAMS = {"fontPt": 9.02, "exRatio": 0.431, "paddingPt": 2.3,
             "maxWidthPt": 400.0}


def strip_display(latex: str) -> str:
    s = latex.strip()
    s = re.sub(r"^\$\$", "", s)
    s = re.sub(r"\$\$$", "", s)
    return s.strip()


def load_corpus():
    rows = []
    for line in TSV.read_text(encoding="utf-8").splitlines():
        parts = line.split("\t", 1)
        if len(parts) != 2:
            continue
        label = Path(parts[0]).stem
        rows.append((label, strip_display(parts[1])))
    return rows


def mj_boxes(corpus):
    """label -> (widthPt, heightPt, depthPt), cached on disk."""
    cache = json.loads(MJ_CACHE.read_text(encoding="utf-8")) \
        if MJ_CACHE.exists() else {}
    todo = [(l, x) for l, x in corpus if l not in cache]
    if todo:
        proc = subprocess.Popen(
            ["node", str(WORKER), "--worker"],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, text=True, cwd=ROOT)
        for i, (label, latex) in enumerate(todo):
            req = {"id": i + 1,
                   "latexBase64": base64.b64encode(latex.encode()).decode(),
                   **MJ_PARAMS}
            proc.stdin.write(json.dumps(req) + "\n")
            proc.stdin.flush()
            resp = json.loads(proc.stdout.readline())
            if not resp.get("ok"):
                cache[label] = {"ok": False}
                continue
            cache[label] = {"ok": True, "widthPt": resp["widthPt"],
                            "heightPt": resp["heightPt"],
                            "depthPt": resp.get("depthPt", -1)}
        proc.stdin.close()
        proc.wait(timeout=60)
        MJ_CACHE.write_text(json.dumps(cache), encoding="utf-8")
    return cache


def quantiles(vals):
    vals = sorted(vals)
    if not vals:
        return {}

    def q(p):
        k = (len(vals) - 1) * p
        f = math.floor(k)
        c = min(f + 1, len(vals) - 1)
        return round(vals[f] + (vals[c] - vals[f]) * (k - f), 3)

    return {"n": len(vals), "p25": q(0.25), "median": q(0.5),
            "p75": q(0.75), "p95": q(0.95)}


def main() -> int:
    corpus = load_corpus()
    buckets = {}
    for line in MEASURE.read_text(encoding="utf-8").splitlines():
        r = json.loads(line)
        buckets.setdefault(r["image"], r["bucket"])

    # GT display boxes: image stem -> style w/h pt + depth
    gt = {}
    for b in extract_boxes(str(REF_DOCX)):
        stem = Path(b.image_target).stem
        gt[stem] = {"w": b.style_width_pt, "h": b.style_height_pt,
                    "depth_pt": -b.position_half_pt / 2.0}

    # our current display boxes, paired by paragraph order
    ours = {}
    gen_boxes = extract_boxes(str(GEN_DOCX))
    labels_in_order = [l for l, _ in corpus]
    for b in gen_boxes:
        m = re.search(r"image_eq(\d+)", b.image_target)
        if not m:
            continue
        idx = int(m.group(1)) - 1
        if 0 <= idx < len(labels_in_order):
            ours[labels_in_order[idx]] = {"w": b.style_width_pt,
                                          "h": b.style_height_pt,
                                          "depth_pt": -b.position_half_pt / 2.0}

    mj = mj_boxes(corpus)

    rows = []
    for label, _ in corpus:
        if label not in gt or label not in ours:
            continue
        m = mj.get(label) or {}
        if not m.get("ok"):
            continue
        rows.append({
            "image": label,
            "bucket": buckets.get(label, "?"),
            "gt_w": gt[label]["w"], "gt_h": gt[label]["h"],
            "gt_depth": gt[label]["depth_pt"],
            "our_w": ours[label]["w"], "our_h": ours[label]["h"],
            "our_depth": ours[label]["depth_pt"],
            "mj_w": m["widthPt"], "mj_h": m["heightPt"],
            "mj_d": m["depthPt"],
        })

    stats = {}
    for bucket in ("single", "nested", "chain", "multi"):
        br = [r for r in rows if r["bucket"] == bucket]
        if not br:
            continue
        stats[bucket] = {
            "n": len(br),
            "need_w_scale": quantiles([r["gt_w"] / r["mj_w"] for r in br]),
            "need_h_scale": quantiles([r["gt_h"] / r["mj_h"] for r in br]),
            "current_w_scale": quantiles([r["our_w"] / r["mj_w"] for r in br]),
            "current_h_scale": quantiles([r["our_h"] / r["mj_h"] for r in br]),
            "gt_depth_ratio": quantiles(
                [r["gt_depth"] / r["gt_h"] for r in br if r["gt_h"] > 0]),
            "our_depth_ratio": quantiles(
                [r["our_depth"] / r["our_h"] for r in br if r["our_h"] > 0]),
            "gt_h_pt": quantiles([r["gt_h"] for r in br]),
            "our_h_pt": quantiles([r["our_h"] for r in br]),
        }
    OUT_JSON.write_text(json.dumps(
        {"mj_params": MJ_PARAMS, "buckets": stats},
        ensure_ascii=False, indent=1), encoding="utf-8")

    lines = ["# 预览显示盒校准报告", "",
             "本报告由 `scripts/preview_display_calibration.py` 自动生成。",
             "",
             "GT 显示盒 = 参考 docx 的 Word v:shape 样式尺寸（五框之 v:shape 层）；",
             "MJ 盒 = MathJax worker 原始输出（渲染器自有参数）。",
             "比例 = GT/MJ（需要）、当前/MJ（现状）。", ""]
    for bucket, s in stats.items():
        lines.append(f"## {bucket}（n={s['n']}）")
        lines.append("")
        lines.append("| 量 | p25 | 中位 | p75 | p95 |")
        lines.append("|---|---|---|---|---|")
        for k, name in [("need_w_scale", "需要 w 比例"), ("need_h_scale", "需要 h 比例"),
                        ("current_w_scale", "当前 w 比例"), ("current_h_scale", "当前 h 比例"),
                        ("gt_depth_ratio", "GT 深度比"), ("our_depth_ratio", "当前深度比"),
                        ("gt_h_pt", "GT 高 pt"), ("our_h_pt", "当前高 pt")]:
            q = s[k]
            lines.append(f"| {name} | {q['p25']} | {q['median']} | {q['p75']} | {q['p95']} |")
        lines.append("")
    OUT_MD.write_text("\n".join(lines), encoding="utf-8")
    print(f"rows={len(rows)} -> {OUT_JSON.relative_to(ROOT)}, {OUT_MD.relative_to(ROOT)}")
    for bucket, s in stats.items():
        print(f"  {bucket}: needW={s['need_w_scale']['median']} "
              f"needH={s['need_h_scale']['median']} "
              f"curW={s['current_w_scale']['median']} "
              f"curH={s['current_h_scale']['median']} "
              f"gtDepth={s['gt_depth_ratio']['median']} "
              f"ourDepth={s['our_depth_ratio']['median']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
