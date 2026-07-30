# 标尺第一轮 diff 报告：真 MathType vs MathJax 预览

分支：`feature/mathtype-preview-ruler`
日期：2026-07-30
语料：`fraction-split-reference.docx`（真 MathType，425 个 OLE 公式）

## 1. 管线（全部可复现）

```
docx ─┬─ docxtolatex (Go, J:\docx2latex)  → LaTeX + OLE 配对报告
      ├─ extract_formula_boxes            → OLE ↔ WMF 预览配对
      └─ wmf_glyph_layout.parse_wmf       → 真值 glyph JSON
LaTeX → tools/mathjax/render_mathjax_svg.cjs (fontPt=10.5) → SVG → glyph JSON
       → scripts/mathjax_glyph_diff.py    → 逐公式 diff + 汇总
```

复现命令：

```powershell
# 1) LaTeX 配对（需 Go）
cd J:\docx2latex\docxtolatex
go run main.go --wordDocx J:\latextomathtype\rebuild-assets\external\fraction-split-reference.docx `
  --output J:\latextomathtype\target\docxtolatex-out --report
# 2) diff
python scripts/mathjax_glyph_diff.py `
  --docx rebuild-assets/external/fraction-split-reference.docx `
  --report target/docxtolatex-out/docxtolatex-out.report.json `
  --out-json target/wmf-ruler/diff-report.json --out-md target/wmf-ruler/diff-report.md
```

## 2. 汇总结果（425 对，0 渲染错误）

| 层 | 指标 | 值 |
|---|---|---|
| 身份层 | 主线字符序列全等 | **410 / 425 (96.5%)** |
| 身份层 | 平均 identity ratio | 0.987 |
| 几何层 | 主线 Δx RMSE（中位数） | 5.34 pt |
| 几何层 | 基线偏移 MAD（中位数） | 2.45 pt |
| 几何层 | 分数线数量一致 | **425 / 425 (100%)** |
| 框级 | 宽度差 GT−MJ（中位数） | −1.35 pt（MJ 略宽） |
| 框级 | **高度差 GT−MJ（中位数）** | **+10.84 pt** |

## 3. 三个实质性发现（这是标尺的第一批产出）

### 3.1 预览框高度：真值比 MathJax 内容框系统性地高 10.8 pt

真 MathType 预览框上下留白远大于 MathJax 内容框（单分式 GT 高 ~28pt vs MJ ~17pt）。
这就是生成器不得不用 `OLE_PREVIEW_SIZE=9.02`、`paddingPt=2.3` 等拟合常数凑显示框的
根源。**10.84pt 是实测值，可以直接成为校准基准**，替代拟合。

### 3.2 长链分式的间距累积漂移

| 结构 | n | Δx RMSE 中位数 |
|---|---|---|
| 单分式 | 139 | 3.26 pt |
| 链式（≥3 分数线） | 225 | 8.78 pt |

单个分式的宽度高度吻合（image5: GT 90.0pt vs MJ 89.3pt），但 MathJax 与 MathType
在 `\times`、分数线两端的间距策略不同，每个分式差 1–3pt，长链上累积成 30–90pt 的
总宽度偏差。→ 校准 profile 应按结构分类（链式分式需要单独的间距修正项）。

### 3.3 身份层残留的 15 个差异全部可归类

| 类别 | 数量 | 性质 |
|---|---|---|
| 配对异常（预览与 OLE 内容不一致，集中在 image372–397 区段） | 5 | **数据源问题**，非生成器问题；标尺反而能用来检测语料脏数据 |
| 真值用全角括号 `（）`（宋体），LaTeX/MathJax 用半角 | 6 | 真实保真差异，需决策：是否按 LaTeX 源还原全角 |
| `\left(\right)` 嵌套结构的主线检测边界 | 2 | diff 工具自身限制 |
| 带分数主线选择歧义（`15` vs `21/32`） | 2 | diff 工具自身限制 |

## 4. 过程中固化的工具知识

- MathJax SVG：`fontCache:"none"` 时字形是 `<path data-c="...">`，数学斜体用
  数学字母数字符号区间（U+1D434–1D467），需映射回 ASCII 做身份比对；
  坐标系根节点 `scale(1,-1)`，需完整仿射矩阵累乘。
- 主基线检测三层策略（关系符 > 运算符 > 跨度/字号/位置）在 425 例上把误检压到 4 例。
- 高括号组件必须强制归入主线（纵向中心在数学轴上，不在任何基线上）。
- 基线聚类要用"到簇中心距离"而非"链式相邻距离"，否则密集阶梯会跨簇粘连。

## 5. 已知边界与下一步

1. **字宽/间距是下一个校准目标**：几何层已能量化每个分式的间距差，
   可以按结构类拟合 profile 表（这是 §3.2 的直接延伸）。
2. 配对异常区段（image372–397）需要回到语料层核查——标尺的意外用途：
   **语料质量审计**。
3. 本语料只覆盖分数/裂项类。xsc 全语料（39551 对）跑同一管线即可出全类别
   的覆盖率 + 偏差排行榜，数据源在 `E:\新加卷\新建文件夹\xsc资料`。
4. 像素层（第五层）尚未接；当前身份+几何两层已能定位大部分问题。

相关文档：[wmf-ruler-prototype-report.md](wmf-ruler-prototype-report.md)（解析器验证）
