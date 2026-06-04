# -*- coding: utf-8 -*-
"""
Parse the docxtolatex-recovered .tex (authoritative LaTeX of 测试.docx) and
re-organize ALL questions BY ACTUAL CONTENT into knowledge topics, each laid out
as 课堂讲解 / 课堂练习 / 课后作业, then emit a PaperExportRequest JSON.
"""
import re, os, json, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

TEX = r"J:\d2t-run\xtol_out\xtol_out.tex"
IMG_DIR = r"J:\latextomathtype\rebuild-assets"
OUT_JSON = r"J:\latextomathtype\rebuild\reorganized-paper.json"
PREVIEW = r"J:\latextomathtype\rebuild\preview.txt"

with open(TEX, encoding="utf-8") as f:
    full = f.read()

occ = [m.start() for m in re.finditer("一、单选题", full)]
text = full[occ[1]:]

# ---------- cleaning helpers ----------
def reverse_escapes(s):
    s = s.replace(r"\textbackslash{}", "\\")
    s = s.replace(r"\^{}", "^").replace(r"\^", "^")
    s = s.replace(r"\_", "_")
    s = s.replace(r"\{", "{").replace(r"\}", "}")
    s = s.replace(r"\textasciitilde{}", "~")
    return s

def strip_rm(s):
    # MTEF roman wrappers:  { \rm{ X } } -> X  (repeat for nesting)
    for _ in range(8):
        s2 = re.sub(r"\{\s*\\rm\s*\{\s*(.*?)\s*\}\s*\}", r"\1", s)
        if s2 == s:
            break
        s = s2
    return s

def dollars(s):
    return re.sub(r"\$\$\s*(.*?)\s*\$\$", lambda m: "$" + strip_rm(m.group(1)).strip() + "$", s, flags=re.S)

# convert leftover LaTeX commands that ended up OUTSIDE math into unicode/plain
LEFTOVER = [
    (r"\\times", "×"), (r"\\div", "÷"), (r"\\cdots", "…"), (r"\\ldots", "…"),
    (r"\\therefore", "∴"), (r"\\sim", "~"), (r"\\le", "≤"), (r"\\ge", "≥"),
    (r"\\pi", "π"), (r"\\uparrow", "↑"), (r"\\left", ""), (r"\\right", ""),
]
def leftover_to_unicode(s):
    # only convert leftover commands OUTSIDE $...$ math spans; keep \times etc. inside math
    parts = re.split(r"(\$[^$]*\$)", s)
    out = []
    for p in parts:
        if p.startswith("$") and p.endswith("$") and len(p) >= 2:
            out.append("$" + strip_rm(p[1:-1]) + "$")
        else:
            q = strip_rm(p)
            for pat, rep in LEFTOVER:
                q = re.sub(pat, rep, q)
            q = re.sub(r"(C_\d+\^?\{?\d+\}?(?:\s*-\s*\d+)?\s*=\s*\d+)",
                       lambda m: "$" + m.group(1).replace(" ", "") + "$", q)
            out.append(q)
    return "".join(out)

PIC = re.compile(r"beginPic\{([^}]+)\}endPic")
def extract_pics(s):
    return PIC.sub("", s), PIC.findall(s)

def strip_struct(s):
    s = re.sub(r"\\subsection\*\{(.*?)\}", r"\1", s, flags=re.S)
    s = s.replace(r"\begin{enumerate}", " ").replace(r"\end{enumerate}", " ")
    s = s.replace(r"\item", " ").replace(r"\\", " ")
    for h in ["一、单选题", "二、填空题", "三、解答题"]:
        s = s.replace(h, "")
    return s

def clean(s):
    s, pics = extract_pics(s)
    s = strip_struct(s)
    s = reverse_escapes(s)
    s = dollars(s)
    s = leftover_to_unicode(s)
    s = re.sub(r"[ \t]*\n[ \t]*", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s, pics

def main_fig(pics):
    for p in pics:
        m = re.match(r"image(\d+)-", p)
        if m and 1 <= int(m.group(1)) <= 24:
            return os.path.join(IMG_DIR, p)
    return None

def split_solution_answer(seg):
    idxs = [m.start() for m in re.finditer("【答案】", seg)]
    last = idxs[-1]
    before = seg[:last]
    after = seg[last + len("【答案】"):]
    parts = after.split("\n", 1)
    return before, parts[0].strip(), (parts[1] if len(parts) > 1 else "")

# ---------- segment by 【知识点】 ----------
kp = [(m.start(), m.end()) for m in re.finditer("【知识点】", text)]
assert len(kp) == 56, f"expected 56, got {len(kp)}"

raw, stems_next = {}, {}
for i in range(56):
    seg = text[kp[i][0]: (kp[i + 1][0] if i + 1 < 56 else len(text))]
    kpv = (re.match(r"【知识点】\s*([^\n【]*)", seg) or [None, ""])
    kpv = kpv.group(1).strip() if hasattr(kpv, "group") else ""
    dm = re.search(r"【难度】\s*([^\n【]*)", seg)
    diff = dm.group(1).strip() if dm else ""
    before, answer, nextstem = split_solution_answer(seg)
    an = before
    for lab in ["【知识点】", "【难度】", "【分析】", "【解析】", "【解答】", "【详解】", "【点评】", "【点睛】", "【答案】"]:
        an = re.sub(re.escape(lab) + r"[^\n【]*" if lab in ("【知识点】", "【难度】") else re.escape(lab), " ", an)
    raw[i + 1] = dict(kp=kpv, diff=diff, analysis=an, answer=answer)
    stems_next[i + 2] = nextstem

stem_raw = {1: text[:kp[0][0]]}
for i in range(2, 57):
    stem_raw[i] = stems_next.get(i, "")

# clean, correct analysis overrides for questions whose recovered MTEF text was garbled
ANALYSIS_OVERRIDE = {
    8: "从 $6$ 个点中任取 $3$ 个共有 $C_6^3=20$ 种，其中 $A_1,A_2,A_3$ 三点共线不能构成三角形，"
       "去掉这 $1$ 种，所以共有 $C_6^3-1=20-1=19$ 个三角形。",
    11: "按长方形所含小格的位置分类计数：含上方“※”的长方形有 $18$ 个，含下方“※”的有 $24$ 个，"
        "其中同时含两个“※”的被重复计数 $8$ 个，应减去，故含“※”的长方形共有 $18+24-8=34$ 个。",
    18: "按组成三角形的小三角形个数（$1,2,4,6,8,12$ 个）分类统计后求和，图中共有 $64$ 个三角形。",
    19: "横向每两条平行线之间取顶点组合：三角形共 $C_5^2\\times 4=40$ 个，梯形共 $C_5^2\\times 6=60$ 个，"
        "故梯形个数与三角形个数之差为 $60-40=20$。",
    21: "$6$ 个数两两取共 $C_6^2=15$ 种。按除以 $3$ 的余数分类：余 $0$ 有 $3,9$；余 $1$ 有 $1,7$；"
        "余 $2$ 有 $5,11$。和能被 $3$ 整除的情形是“两数都余 $0$”（$1$ 种）或“一余 $1$ 一余 $2$”"
        "（$2\\times 2=4$ 种），共 $5$ 种。故和不能被 $3$ 整除的取法有 $15-5=10$ 种。",
    40: "这串数都是除以 $3$ 余 $1$ 的数，故只需考虑 $4$ 倍关系。把存在倍数关系的数归入同一抽屉，"
        "至多可构造 $507$ 个互不成倍数的抽屉，因此任取 $508$ 个数，必有一个数是另一个数的倍数。",
}

Q = {}
for n in range(1, 57):
    s_clean, s_pics = clean(stem_raw[n])
    s_clean = re.sub(r"^\d+\.\s*", "", s_clean)
    if n in ANALYSIS_OVERRIDE:
        a_clean = ANALYSIS_OVERRIDE[n]
    else:
        a_clean, _ = clean(raw[n]["analysis"])
    ans_clean, _ = clean(raw[n]["answer"])
    Q[n] = dict(stem=s_clean, kp=raw[n]["kp"], diff=raw[n]["diff"],
                analysis=a_clean, answer=ans_clean, image=main_fig(s_pics))

# ---------- supplement question ----------
SUPP = {
    "S1": dict(
        stem="证明：任意 $13$ 个人中，至少有 $2$ 个人出生在同一个月份。",
        kp="抽屉原理解决证明题", diff="★★",
        analysis="一年有 $12$ 个月，把它们看作 $12$ 个抽屉，把 $13$ 个人看作 $13$ 个物体。"
                 "因为 $13 = 12 \\times 1 + 1$，根据抽屉原理，把 $13$ 个物体放进 $12$ 个抽屉，"
                 "至少有一个抽屉里有 $1+1=2$ 个物体，即至少有 $2$ 个人出生在同一个月份。",
        answer="见解析：由 $13>12$，根据抽屉原理必有两人同月出生。", image=None),
}

# ---------- CONTENT-BASED topic / phase mapping ----------
TYPE5 = {53, 54, 55, "S1"}   # 解答题
TYPE1 = {1}                  # 单选
OPTIONS_Q1 = [("A", "$5$"), ("B", "$10$"), ("C", "$15$"), ("D", "$20$")]
CORRECT_Q1 = "B"

topics = [
    ("专题一　数线段与求交点（组合计数）", [
        ("课堂讲解", [1]), ("课堂练习", [45, 8, 52]), ("课后作业", [5, 54])]),
    ("专题二　数图形中的三角形与梯形", [
        ("课堂讲解", [3]), ("课堂练习", [23, 14, 18, 13]), ("课后作业", [2, 19, 28, 31, 49, 56])]),
    ("专题三　数长方形与正方形", [
        ("课堂讲解", [27]), ("课堂练习", [32, 30, 29]), ("课后作业", [4, 11, 33])]),
    ("专题四　简单抽屉原理（平均与保证）", [
        ("课堂讲解", [20]), ("课堂练习", [22, 16, 48, 25]), ("课后作业", [9, 17, 53])]),
    ("专题五　构造型抽屉原理（余数·配对·构造）", [
        ("课堂讲解", [21]), ("课堂练习", [12, 41, 42, 26]), ("课后作业", [6, 7, 39, 40, 50])]),
    ("专题六　最不利原则", [
        ("课堂讲解", [38]), ("课堂练习", [34, 24, 51, 10, 15]), ("课后作业", [35, 36, 37, 43, 46, 55])]),
    ("专题七　抽屉原理的证明与综合应用", [
        ("课堂讲解", ["S1"]), ("课堂练习", [44]), ("课后作业", [47])]),
]

def get(n):
    return SUPP[n] if isinstance(n, str) else Q[n]

def qtype(n):
    if n in TYPE1:
        return 1
    if n in TYPE5:
        return 5
    return 4

# ---------- build PaperExportRequest ----------
serial = 0
sections = []
for headline, phases in topics:
    qlist = []
    for phase_label, nums in phases:
        first = True
        for n in nums:
            d = get(n)
            serial += 1
            q = {
                "serialNumber": serial,
                "questionType": qtype(n),
                "content": d["stem"],
            }
            if first:
                q["phaseLabel"] = phase_label
                first = False
            if n in TYPE1:
                q["options"] = [{"prefix": p, "content": c} for p, c in OPTIONS_Q1]
            if d.get("image"):
                q["images"] = [d["image"]]
            if d.get("kp"):
                q["knowledgePoint"] = d["kp"]
            if d.get("diff"):
                q["difficulty"] = d["diff"]
            if d.get("analysis"):
                q["analyze"] = d["analysis"]
            if d.get("answer"):
                q["correct"] = d["answer"]
            qlist.append(q)
    sections.append({"headline": headline, "questions": qlist})

paper = {
    "paper": {"name": "小学奥数·计数与抽屉原理专题精讲（按知识点重排）",
              "subjectType": 1, "stage": 1, "score": 112, "suggestTime": 70},
    "sections": sections,
}
with open(OUT_JSON, "w", encoding="utf-8") as f:
    json.dump(paper, f, ensure_ascii=False, indent=2)

# ---------- readable preview ----------
with open(PREVIEW, "w", encoding="utf-8") as f:
    f.write("重排预览（按内容分组）\n" + "=" * 60 + "\n")
    sn = 0
    for headline, phases in topics:
        f.write("\n■ " + headline + "\n")
        for phase_label, nums in phases:
            f.write("  ── %s ──\n" % phase_label)
            for n in nums:
                d = get(n)
                sn += 1
                src = "新增" if isinstance(n, str) else f"原Q{n}"
                img = "[图]" if d.get("image") else "   "
                f.write("   %2d. %-5s %s 【%s|%s】 答:%s\n        %s\n" % (
                    sn, src, img, d.get("kp", ""), d.get("diff", ""),
                    d.get("answer", "")[:18], d["stem"][:46]))
    f.write("\n共 %d 题（原56 + 新增1）\n" % sn)

print("JSON ->", OUT_JSON)
print("sections:", len(sections), " questions:", serial)
