# -*- coding: utf-8 -*-
"""
First identity+geometry diff between genuine MathType WMF previews (ground
truth) and our MathJax preview pipeline, for the same LaTeX formulas.

Pipeline:
    docx ─┬─ wmf_glyph_layout.parse_wmf        -> ground-truth glyph JSON
          ├─ extract_formula_boxes             -> ole <-> wmf pairing
          └─ docxtolatex report                -> LaTeX per OLE object
    LaTeX -> tools/mathjax/render_mathjax_svg.cjs -> SVG -> glyph JSON (same schema)
    diff per formula: identity (char sequence), geometry (baselines, rules,
    matched-glyph position RMSE), box dimensions.

Usage:
    python scripts/mathjax_glyph_diff.py ^
        --docx rebuild-assets/external/fraction-split-reference.docx ^
        --report target/docxtolatex-out/docxtolatex-out.report.json ^
        --out-json target/wmf-ruler/diff-report.json ^
        --out-md   target/wmf-ruler/diff-report.md
"""
from __future__ import annotations

import argparse
import base64
import difflib
import json
import math
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "rebuild"))

from wmf_glyph_layout import parse_wmf  # noqa: E402
from extract_formula_boxes import extract_boxes  # noqa: E402

SVG_NS = "http://www.w3.org/2000/svg"

# Math alphanumeric symbols -> ASCII (MathJax uses these for math italic)
def map_math_alnum(cp: int) -> int:
    if 0x1D434 <= cp <= 0x1D44D:
        return ord("A") + (cp - 0x1D434)
    if 0x1D44E <= cp <= 0x1D467:
        return ord("a") + (cp - 0x1D44E)
    if 0x1D7CE <= cp <= 0x1D7D7:
        return ord("0") + (cp - 0x1D7CE)
    return cp

# canonical char for cross-font identity compare
CANON = {
    "-": "−", "‐": "−", "−": "−",
    "–": "−", "—": "−",
}


def canon(ch: str) -> str:
    return CANON.get(ch, ch)


# ---------------------------------------------------------------------------
# MathJax SVG -> glyph JSON
# ---------------------------------------------------------------------------
def mat_mul(m1, m2):
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (
        a1 * a2 + c1 * b2,
        b1 * a2 + d1 * b2,
        a1 * c2 + c1 * d2,
        b1 * c2 + d1 * d2,
        a1 * e2 + c1 * f2 + e1,
        b1 * e2 + d1 * f2 + f1,
    )


def parse_transform(s: str):
    m = (1, 0, 0, 1, 0, 0)
    for kind, args in re.findall(r"(translate|scale|matrix)\(([^)]*)\)", s or ""):
        vals = [float(v) for v in re.split(r"[,\s]+", args.strip()) if v]
        if kind == "translate":
            t = (1, 0, 0, 1, vals[0], vals[1] if len(vals) > 1 else 0.0)
        elif kind == "scale":
            sx = vals[0]
            sy = vals[1] if len(vals) > 1 else sx
            t = (sx, 0, 0, sy, 0, 0)
        else:
            t = tuple(vals) if len(vals) == 6 else (1, 0, 0, 1, 0, 0)
        m = mat_mul(m, t)
    return m


def parse_svg_glyphs(svg: str, width_pt: float) -> dict:
    root = ET.fromstring(svg)
    vb = [float(v) for v in root.get("viewBox").split()]
    pt_per_unit = width_pt / vb[2] if vb[2] else 1.0
    glyphs = []
    rules = []

    def walk(el, m):
        tag = el.tag.split("}")[-1]
        m2 = mat_mul(m, parse_transform(el.get("transform", "")))
        if tag == "path" and el.get("data-c"):
            cp = int(el.get("data-c"), 16)
            mapped = map_math_alnum(cp)
            x = m2[4]
            y = m2[5]
            scale = math.hypot(m2[0], m2[1])
            glyphs.append({
                "ch": chr(mapped),
                "hex": f"{cp:X}",
                "style": "mathitalic" if mapped != cp else "normal",
                "x_units": x, "y_units": y, "scale": scale,
            })
        elif tag == "rect":
            try:
                x = float(el.get("x", 0) or 0)
                y = float(el.get("y", 0) or 0)
                w = float(el.get("width", 0) or 0)
                h = float(el.get("height", 0) or 0)
            except ValueError:
                w = h = 0
            # apply matrix to origin; assume axis-aligned (MathJax rects are)
            ox = m2[0] * x + m2[2] * y + m2[4]
            oy = m2[1] * x + m2[3] * y + m2[5]
            sx = math.hypot(m2[0], m2[1])
            sy = math.hypot(m2[2], m2[3])
            if w > 0 and h > 0:
                rules.append({"kind": "rect", "x_units": ox, "y_units": oy,
                              "w_units": w * sx, "h_units": h * sy})
        for child in el:
            walk(child, m2)

    walk(root, (1, 0, 0, 1, 0, 0))

    # y-down conversion: MathJax internal y is flipped by root scale(1,-1),
    # already included in the matrix walk; convert to pt relative to viewBox.
    for g in glyphs:
        g["x_pt"] = (g["x_units"] - vb[0]) * pt_per_unit
        g["y_pt"] = (g["y_units"] - vb[1]) * pt_per_unit
        g["size_pt"] = 1000 * g["scale"] * pt_per_unit
    for r in rules:
        r["x_pt"] = (r["x_units"] - vb[0]) * pt_per_unit
        r["y_pt"] = (r["y_units"] - vb[1]) * pt_per_unit
        r["w_pt"] = r["w_units"] * pt_per_unit
        r["h_pt"] = r["h_units"] * pt_per_unit
    return {"glyphs": glyphs, "rules": rules, "viewBox": vb}


# ---------------------------------------------------------------------------
# baseline clustering / sequences
# ---------------------------------------------------------------------------
def cluster_baselines(glyphs, tol_pt=1.2):
    ys = sorted(g["y_pt"] for g in glyphs)
    clusters: list[list[float]] = []
    centers: list[float] = []
    for y in ys:
        if clusters and abs(y - centers[-1]) <= tol_pt:
            clusters[-1].append(y)
            centers[-1] = sum(clusters[-1]) / len(clusters[-1])
        else:
            clusters.append([y])
            centers.append(y)

    def bucket(y):
        return min(range(len(centers)), key=lambda i: abs(y - centers[i]))

    return centers, bucket


def line_sequence(glyphs, bucket, idx, force_main=False):
    line = [g for g in glyphs
            if bucket(g["y_pt"]) == idx
            or (force_main and g.get("assembly"))]
    line.sort(key=lambda g: g["x_pt"])
    return "".join(canon(g["ch"]) for g in line), line


# ---------------------------------------------------------------------------
# diff per formula
# ---------------------------------------------------------------------------
def diff_formula(gt: dict, mj: dict, mj_meta: dict) -> dict:
    out: dict = {}
    gt_g = [g for g in gt["glyphs"] if "x_pt" in g]
    mj_g = mj["glyphs"]
    out["gt_glyphs"] = len(gt_g)
    out["mj_glyphs"] = len(mj_g)

    # --- box dims
    bbox = gt.get("placeable_bbox_pt") or [0, 0, 0, 0]
    out["gt_box_pt"] = [round(bbox[2] - bbox[0], 2), round(bbox[3] - bbox[1], 2)]
    out["mj_box_pt"] = [round(mj_meta["widthPt"], 2), round(mj_meta["heightPt"], 2)]

    # --- baselines
    gt_centers, gt_bucket = cluster_baselines(gt_g)
    mj_centers, mj_bucket = cluster_baselines(mj_g)
    out["gt_baselines"] = len(gt_centers)
    out["mj_baselines"] = len(mj_centers)

    # main baseline, tiered:
    #   tier1: clusters with relations (= < >)   -> true main line
    #   tier2: clusters with + or −              -> operator line of frac chains
    #   tier3: widest x-span, tie -> larger avg size -> topmost cluster
    REL1 = {"=", "<", ">"}
    REL2 = {"+", "−"}

    def pick_main(glyphs, centers, bucket):
        spans = [[None, None] for _ in centers]
        sizes = [[] for _ in centers]
        t1, t2 = set(), set()
        for g in glyphs:
            i = bucket(g["y_pt"])
            sizes[i].append(g.get("size_pt") or 0)
            ch = canon(g["ch"])
            if ch in REL1:
                t1.add(i)
            elif ch in REL2:
                t2.add(i)
            lo, hi = spans[i]
            spans[i] = [g["x_pt"] if lo is None else min(lo, g["x_pt"]),
                        g["x_pt"] if hi is None else max(hi, g["x_pt"])]

        def span(i):
            lo, hi = spans[i]
            return (hi - lo) if lo is not None else 0

        def avg_size(i):
            return sum(sizes[i]) / len(sizes[i]) if sizes[i] else 0

        def best(cands):
            widest = max(span(i) for i in cands)
            near = [i for i in cands if span(i) >= widest * 0.9]
            big = max(avg_size(i) for i in near)
            near = [i for i in near if avg_size(i) >= big * 0.9]
            return min(near, key=lambda i: centers[i])  # topmost

        cands = t1 or t2 or set(range(len(centers)))
        return best(list(cands))

    gt_main = pick_main(gt_g, gt_centers, gt_bucket)
    mj_main = pick_main(mj_g, mj_centers, mj_bucket)

    # tall delimiter assemblies are logically main-line content; their vertical
    # center sits on the math axis, not the baseline, so force them in
    gt_seq, gt_line = line_sequence(gt_g, gt_bucket, gt_main, force_main=True)
    mj_seq, mj_line = line_sequence(mj_g, mj_bucket, mj_main)
    out["gt_main_seq"] = gt_seq
    out["mj_main_seq"] = mj_seq

    sm = difflib.SequenceMatcher(a=gt_seq, b=mj_seq, autojunk=False)
    out["identity_ratio"] = round(sm.ratio(), 4)

    # --- geometry: matched main-line glyphs, align left edges + baseline
    matches = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                matches.append((gt_line[i1 + k], mj_line[j1 + k]))
    if matches:
        dx0 = matches[0][0]["x_pt"] - matches[0][1]["x_pt"]
        dy0 = matches[0][0]["y_pt"] - matches[0][1]["y_pt"]
        errs_x = [abs((a["x_pt"] - b["x_pt"]) - dx0) for a, b in matches]
        errs_y = [abs((a["y_pt"] - b["y_pt"]) - dy0) for a, b in matches]
        out["matched"] = len(matches)
        out["x_rmse_pt"] = round(math.sqrt(sum(e * e for e in errs_x) / len(errs_x)), 3)
        out["x_max_pt"] = round(max(errs_x), 3)
        out["y_max_pt"] = round(max(errs_y), 3)
    else:
        out["matched"] = 0

    # --- baseline offsets relative to main baseline
    gt_off = sorted(c - gt_centers[gt_main] for c in gt_centers)
    mj_off = sorted(c - mj_centers[mj_main] for c in mj_centers)
    n = min(len(gt_off), len(mj_off))
    if n:
        out["baseline_offset_mad_pt"] = round(
            sum(abs(gt_off[i] - mj_off[i]) for i in range(n)) / n, 3)
    out["gt_baseline_offsets"] = [round(v, 2) for v in gt_off]
    out["mj_baseline_offsets"] = [round(v, 2) for v in mj_off]

    # --- rules (fraction bars)
    gt_rules = [r for r in gt["rules"] if r["kind"] == "line"]
    mj_rules = [r for r in mj["rules"] if r.get("w_pt", 0) > r.get("h_pt", 0) * 3]
    out["gt_rules"] = len(gt_rules)
    out["mj_rules"] = len(mj_rules)
    gt_ry = sorted(r["y1_pt"] - gt_centers[gt_main] for r in gt_rules)
    mj_ry = sorted(r["y_pt"] - mj_centers[mj_main] for r in mj_rules)
    n = min(len(gt_ry), len(mj_ry))
    if n:
        out["rule_y_mad_pt"] = round(sum(abs(gt_ry[i] - mj_ry[i]) for i in range(n)) / n, 3)
    return out


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--report", required=True)
    ap.add_argument("--out-json", required=True)
    ap.add_argument("--out-md", required=True)
    ap.add_argument("--font-pt", type=float, default=10.5)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    docx = Path(args.docx)
    report = json.load(open(args.report, encoding="utf-8"))
    equations = {e["source"]: e for e in report["equations"] if e.get("status") == "converted"}

    boxes = extract_boxes(docx)
    pairs = []
    for box in boxes:
        ole = box.ole_target  # e.g. "embeddings/oleObject1.bin"
        img = box.image_target  # e.g. "media/image4.wmf"
        if ole in equations and img and img.lower().endswith(".wmf"):
            pairs.append({
                "ole": ole, "wmf": "word/" + img,
                "latex": equations[ole]["output"],
                "box": box,
            })
    if args.limit:
        pairs = pairs[: args.limit]
    print(f"pairs: {len(pairs)}")

    # ---- batch render with the MathJax worker
    worker = subprocess.Popen(
        ["node", str(ROOT / "tools/mathjax/render_mathjax_svg.cjs")],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True, cwd=ROOT)
    renders = {}
    for i, p in enumerate(pairs):
        req = {"id": i, "latexBase64": base64.b64encode(p["latex"].encode()).decode(),
               "fontPt": args.font_pt}
        worker.stdin.write(json.dumps(req) + "\n")
        worker.stdin.flush()
        line = worker.stdout.readline()
        if not line:
            break
        r = json.loads(line)
        renders[r["id"]] = r
    worker.kill()

    z = zipfile.ZipFile(docx)
    results = []
    for i, p in enumerate(pairs):
        r = renders.get(i)
        if not r or not r.get("ok"):
            results.append({"wmf": p["wmf"], "ole": p["ole"], "error": r.get("error") if r else "no render"})
            continue
        gt = parse_wmf(z.read(p["wmf"]))
        svg = base64.b64decode(r["svgBase64"]).decode()
        mj = parse_svg_glyphs(svg, r["widthPt"])
        d = diff_formula(gt, mj, r)
        d.update({"wmf": p["wmf"], "ole": p["ole"], "latex": p["latex"],
                  "word_box_pt": [p["box"].style_width_pt, p["box"].style_height_pt],
                  "context": p["box"].context})
        results.append(d)

    ok = [r for r in results if "error" not in r]

    def avg(key):
        vals = [r[key] for r in ok if r.get(key) is not None]
        return round(sum(vals) / len(vals), 3) if vals else None

    summary = {
        "pairs": len(pairs),
        "diffed": len(ok),
        "render_errors": len(results) - len(ok),
        "identity_ratio_avg": avg("identity_ratio"),
        "identity_perfect": sum(1 for r in ok if r.get("identity_ratio") == 1),
        "x_rmse_pt_avg": avg("x_rmse_pt"),
        "x_max_pt_p95": None,
        "baseline_offset_mad_avg": avg("baseline_offset_mad_pt"),
        "rule_count_match": sum(1 for r in ok if r.get("gt_rules") == r.get("mj_rules")),
        "font_pt": args.font_pt,
    }
    xmax = sorted((r["x_max_pt"] for r in ok if r.get("x_max_pt") is not None))
    if xmax:
        summary["x_max_pt_p95"] = round(xmax[int(len(xmax) * 0.95)], 3)

    payload = {"summary": summary, "formulas": results}
    Path(args.out_json).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out_json).write_text(json.dumps(payload, ensure_ascii=False, indent=1),
                                   encoding="utf-8")

    # ---- markdown report
    worst = sorted((r for r in ok if r.get("x_rmse_pt") is not None),
                   key=lambda r: -(r["x_rmse_pt"]))
    lines = ["# 真值 vs MathJax 预览 diff 报告", "",
             f"- 语料: `{docx.name}`", f"- 配对公式: {summary['pairs']}，成功 diff: {summary['diffed']}",
             f"- MathJax fontPt: {summary['font_pt']}", "",
             "## 汇总", "",
             "| 指标 | 值 |", "|---|---|",
             f"| 身份层全等公式 | {summary['identity_perfect']} / {summary['diffed']} |",
             f"| 平均 identity ratio | {summary['identity_ratio_avg']} |",
             f"| 主线 Δx RMSE (pt) | {summary['x_rmse_pt_avg']} |",
             f"| 主线 Δx max P95 (pt) | {summary['x_max_pt_p95']} |",
             f"| 基线偏移 MAD (pt) | {summary['baseline_offset_mad_avg']} |",
             f"| 分数线数量一致 | {summary['rule_count_match']} / {summary['diffed']} |",
             "", "## 几何偏差最大的 10 个公式", "",
             "| wmf | LaTeX | 真值序列 | MathJax 序列 | Δx RMSE | Δx max | 基线 MAD |",
             "|---|---|---|---|---|---|---|"]
    for r in worst[:10]:
        lines.append(
            f"| {Path(r['wmf']).name} | `{r['latex'][:40]}` | {r['gt_main_seq'][:24]} | "
            f"{r['mj_main_seq'][:24]} | {r.get('x_rmse_pt')} | {r.get('x_max_pt')} | "
            f"{r.get('baseline_offset_mad_pt')} |")
    lines += ["", "## 身份层不一致的公式（前 15 个）", "",
              "| wmf | LaTeX | 真值 | MathJax | ratio |", "|---|---|---|---|---|"]
    ident_bad = sorted((r for r in ok if r.get("identity_ratio", 1) < 1),
                       key=lambda r: r["identity_ratio"])[:15]
    for r in ident_bad:
        lines.append(f"| {Path(r['wmf']).name} | `{r['latex'][:40]}` | {r['gt_main_seq'][:30]} | "
                     f"{r['mj_main_seq'][:30]} | {r['identity_ratio']} |")
    Path(args.out_md).write_text("\n".join(lines), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
