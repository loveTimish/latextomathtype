# WMF 字形级标尺解析器 — 原型验证测试文档

分支：`feature/mathtype-preview-ruler`
日期：2026-07-30
语料：`rebuild-assets/external/fraction-split-reference.docx`（真 MathType 文档，421 个 WMF 预览 + 425 个 OLE）

## 1. 目标

验证"真 MathType 预览图做标尺"的技术可行性：从真 MathType docx 的矢量 WMF 预览中，
提取**身份层**（每个字形是什么字符、用什么字体、多大字号）和**几何层**（每个字形的
精确坐标、分数线位置），输出可直接与生成器预览做 diff 的 glyph layout JSON。

工具：`scripts/wmf_glyph_layout.py`

```powershell
# 单个 WMF -> JSON
python scripts/wmf_glyph_layout.py image5.wmf
# 批量：解析 docx 内全部 WMF，输出 JSON + SVG 重建图
python scripts/wmf_glyph_layout.py --docx rebuild-assets/external/fraction-split-reference.docx `
  --out target/wmf-ruler/glyph-layouts.json --svg-dir target/wmf-ruler/svg
```

## 2. 解析中确认的 MathType WMF 结构事实

这些结论来自对原始记录流的逐字节 dump 验证，是后续所有标尺工作的前提：

| 事实 | 证据 |
|---|---|
| WMF 是 placeable WMF，bbox 和 units-per-inch 在 22 字节头里 | 头 key = `0x9AC6CDD7` |
| `SETTEXTALIGN = 0x19`（`TA_UPDATECP \| TA_BASELINE`） | image4/image5 记录流 |
| **EXTTEXTOUT 的 x/y 参数恒为 0，真实锚点由前置 `MOVETO(0x0214)` 给出** | 全部记录 y=0，每条文本记录前有 MOVETO |
| 多字符记录的 `dx` 数组是**相邻字形 origin 间距**（最后一项是末字符步进） | 见 §3 验证 |
| 分数线/横线是 `MOVETO + LINETO(0x0213)` 对 | image5 有 4 条分数线，y=460 |
| Symbol 字体 charset 字节是 DEFAULT(1) 而非 SYMBOL(2)，**须按 face 名判断** | CREATEFONT 记录 charset=1, face=Symbol |
| MTEF 原文以 `MFCOMMENT ESCAPE(id=15)` 形式嵌在 WMF 里（`AppsMFCC...`） | image5 offset 352 的 ESCAPE 记录，本语料中为 MTEF 全文 |
| 中文字体 face 名（宋体）按 GBK 编码存储 | 修正前解码为乱码 `ËÎÌå` |
| Symbol 字体 0x20–0x7E 区间保留 ASCII 基本字符（数字、`= < > + -`） | 裂项公式中的 `=` `-` 解码验证 |

## 3. 正确性验证（几何层）

对 image5（裂项公式 `1/(a×b) = 1/(b-a) × (1/a − 1/b)`，4 个分式并排）做坐标交叉验证：
4 个分子 `1` 的 origin x 由 `MOVETO 锚点 + dx 累计`得出，应与 4 条 `LINETO` 分数线的
水平中心一一对齐：

| 分式 | 分数线 x 范围 (units) | 线中心 | 分子 1 origin | 1 的中心 (origin+84) | 偏差 |
|---|---|---|---|---|---|
| 1 | 64–719 | 391 | 307 | 391 | **0** |
| 2 | 1079–1766 | 1422 | 1338 | 1422 | **0** |
| 3 | 1941–2150 | 2045 | 1961 | 2045 | **0** |
| 4 | 2486–2680 | 2583 | 2499 | 2583 | **0** |

4/4 完全对齐，证明锚点 + dx 的坐标解码与真实排版一致。

重建文本行（按基线 y 分桶）：

```
y= 331  1111            ← 4 个分子
y= 544  =(-)            ← 主线运算符
y= 808  a×bb-aab        ← 4 个分母：a×b, b-a, a, b
分数线: y=460, x∈[64,719] [1079,1766] [1941,2150] [2486,2680]
```

与裂项公式语义完全吻合。image6 重建为 `1` / `n×(n+1)×(n+2)`，正确。

## 4. 全语料解析统计（421/421 成功，0 警告）

| 指标 | 值 |
|---|---|
| 解析成功 | 421 / 421（100%） |
| 字形总数 | 10876 |
| 图形规则数（分数线/矩形/多边形） | 1598 |
| 含分数线的公式 | 384 |
| 解析警告 | 0 |
| pt 换算缺失 | 0 |

字体使用（CREATEFONT 记录数）：

| 字体 | 记录数 | 用途 |
|---|---|---|
| Times New Roman | 1040 | 变量、数字、括号主体 |
| Symbol | 386 | 运算符、希腊字母、可伸缩括号件 |
| MT Extra | 115 | MathType 私有符号槽位 |
| System / 宋体 | 421 / 12 | 辅助记录 / 中文全角括号 |

字号分布（pt）：主线 10.5（8566 个字形）、上下标 6.06、大型运算符 13.72 ——
三层字号阶梯清晰，可直接作为标尺的 size class。

未知字形：0（全部成功映射到 Unicode 或保留 hex 标记）。

大公式抽查：image245（115 字形、12 条基线），Symbol 可伸缩括号件
`⎛⎜⎝ ⎞⎟⎠` 全部正确解码并出现在正确基线上。

## 5. SVG 重建样例（肉眼校验用）

由提取的 glyph/rule 记录重绘，可在浏览器打开与原预览对照：

- [image4.svg](wmf-ruler-samples/image4.svg) — `a<b`（单线，Symbol 字符）
- [image5.svg](wmf-ruler-samples/image5.svg) — 4 分式裂项（多基线 + 4 分数线）
- [image6.svg](wmf-ruler-samples/image6.svg) — `1/(n(n+1)(n+2))`
- [image245.svg](wmf-ruler-samples/image245.svg) — 115 字形大公式（含可伸缩括号）

注：SVG 重建只做坐标校验，字体一律用 Times New Roman 近似渲染，外观不等于原图；
精确外观对比属于后续像素层标尺。

## 6. 已知限制

1. **MT Extra 是 MathType 私有编码**，当前按 cp1252 占位解码（如 0x4C 显示为 `L`，
   实为私有符号）。需要一张 MT Extra 映射表 —— 可从语料中 MT Extra 字形的
   上下文 + MTEF 语义反推建立（十字交叉箭头已有先例）。
2. 字形的**字宽**目前由 dx 步进近似；精确 ink 宽度需要字体度量（后续接像素层时补）。
3. `System` 字体记录（每文件 1 条，h=16）是 MathType 的占位对象，无实际字形。

## 7. 下一步（接入三层 diff 的接口已经就绪）

解析输出已具备标尺比对所需全部字段：

```json
{ "ch": "1", "font": "Times New Roman", "size_pt": 10.5,
  "x_pt": 9.594, "y_pt": 10.344, "dx_pt": 32.219 }
```

- **身份层 diff**：生成器预览（MathJax SVG DOM 提取同款 JSON）↔ 本 JSON 的
  (ch, font, size) 三元组比对；
- **几何层 diff**：(x_pt, y_pt) 逐字形 RMSE + 分数线端点比对；
- **框级 diff**：`placeable_bbox_pt` ↔ DOCX 显示框（接 `extract_formula_boxes`）；
- MTEF 语义配对：WMF 内嵌 MFCOMMENT 含 MTEF 全文，可直接做 MTEF↔LaTeX 配对，
  不再依赖 docx2tex。
