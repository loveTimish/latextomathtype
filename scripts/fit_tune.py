# -*- coding: utf-8 -*-
"""Fit-tuner: render the whole corpus through the MathJax worker with
mathTypeFit and compare per-image w/h/depth against GT v:shape boxes.

Usage:
  python scripts/fit_tune.py                # evaluate current defaults
  python scripts/fit_tune.py '{...json...}' # evaluate a param override
"""
import base64, json, re, statistics, subprocess, sys
from pathlib import Path

ROOT = Path("J:/latextomathtype")
sys.path.insert(0, str(ROOT / "rebuild"))
from extract_formula_boxes import extract_boxes

REF = ROOT / "rebuild-assets/external/fraction-split-reference.docx"
TSV = ROOT / "target/wmf-ruler/wmf-latex.tsv"
MEAS = ROOT / "target/wmf-ruler/structure-measure.jsonl"
WORKER = ROOT / "tools/mathjax/render_mathjax_svg.cjs"

BASE = {
    "fontPt": 10.495, "exRatio": 0.431, "paddingPt": 2.3, "maxWidthPt": 400.0,
    "mathTypeFit": True,
}

def strip_display(latex):
    s = latex.strip()
    s = re.sub(r"^\$\$", "", s)
    s = re.sub(r"\$\$$", "", s)
    return s.strip()

def main():
    overrides = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {}
    fit_params = overrides.pop("fitParams", {})
    params = {**BASE, **overrides}

    corpus = []
    for line in TSV.read_text(encoding="utf-8").splitlines():
        p = line.split("\t", 1)
        if len(p) == 2:
            corpus.append((Path(p[0]).stem, strip_display(p[1])))

    bucket = {}
    for line in MEAS.read_text(encoding="utf-8").splitlines():
        r = json.loads(line)
        bucket.setdefault(r["image"], r["bucket"])

    gt = {}
    for b in extract_boxes(str(REF)):
        gt[Path(b.image_target).stem] = (
            b.style_width_pt, b.style_height_pt, -b.position_half_pt / 2.0)

    proc = subprocess.Popen(
        ["node", str(WORKER), "--worker"],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL, text=True, cwd=ROOT)
    res = {}
    for i, (label, latex) in enumerate(corpus):
        req = {"id": i + 1,
               "latexBase64": base64.b64encode(latex.encode()).decode(),
               **params}
        if fit_params:
            req["fitParams"] = fit_params
        proc.stdin.write(json.dumps(req) + "\n")
        proc.stdin.flush()
        resp = json.loads(proc.stdout.readline())
        if resp.get("ok"):
            res[label] = (resp["widthPt"], resp["heightPt"], resp["depthPt"])
    proc.stdin.close()
    proc.wait(timeout=300)

    by_bucket = {}
    rows = []
    for label, _ in corpus:
        if label not in gt or label not in res:
            continue
        gw, gh, gd = gt[label]
        ow, oh, od = res[label]
        b = bucket.get(label, "?")
        rows.append((label, b, ow / gw, oh / gh, (od / oh) if oh else 0,
                     (gd / gh) if gh else 0))
        by_bucket.setdefault(b, []).append(rows[-1])

    def med(vals):
        return round(statistics.median(vals), 4) if vals else 0

    print(f"rendered={len(res)}/{len(corpus)}  params={json.dumps(overrides)} fitParams={json.dumps(fit_params)}")
    total_abs = []
    for b, rs in sorted(by_bucket.items()):
        wr = [r[2] for r in rs]; hr = [r[3] for r in rs]
        dr = [abs(r[4] - r[5]) for r in rs]
        abs_wh = [abs(r[2]-1) + abs(r[3]-1) for r in rs]
        total_abs.extend(abs_wh)
        print(f"  {b:7} n={len(rs):3}  our/GT w med={med(wr)}  h med={med(hr)}  "
              f"|depthRatio diff| med={med(dr)}  mean|w,h err|={round(statistics.mean(abs_wh),4)}")
    if total_abs:
        print(f"  ALL mean(|w-1|+|h-1|)={round(statistics.mean(total_abs),4)} "
              f"p90={round(sorted(total_abs)[int(len(total_abs)*0.9)],4)}")

if __name__ == "__main__":
    main()
