# WMF 直接还原 LaTeX — 可行性验证报告

分支：`feature/mathtype-preview-ruler`
日期：2026-07-30
语料：`fraction-split-reference.docx`（421 个 WMF 预览 + 425 个 OLE）

## 结论

**可以。单个 WMF 文件即可还原 LaTeX，不需要 OLE、不需要 docx 容器。**
本语料 421 个 WMF：MTEF 提取 421/421，转换成功 421/421，
与 OLE 来源的 LaTeX 对比 **425/425 完全一致（归一化后逐字节相同）**。

## 原理

真 MathType 预览 WMF 的 `MFCOMMENT` escape（id=15）里嵌着完整 MTEF，
与 OLE `Equation Native` 里的 MTEF 是同一份数据：

```
WMF → MFCOMMENT 提取 MTEF → docxtolatex eqn 包 → LaTeX
```

## 发现的两种注释格式

| 签名 | 布局 | 语料占比 |
|---|---|---|
| `AppsMFCC` | sig(8) + flag(2) + size(4) + size 副本(4) + `Design Science, Inc.\0` + MTEF[size] | 257/421 |
| `MathTypeUU` | sig(10) + 未知(2) + MTEF[剩余全部] | 164/421 |

实现：`scripts/wmf_glyph_layout.py` 的 `extract_mtef_from_comment()`
（`parse_wmf` 结果新增 `mtef_hex` / `mtef_size` 字段）。

## 工具链

1. **提取**（Python，已入库）：`parse_wmf()` 自动提取，批量产物在
   `target/wmf-ruler/mtef/*.bin`
2. **转换**（Go，工作区副本 `target/docxtolatex-copy/`，不入库可重建）：
   - `eqn/parsebody.go`：新增 `ConvertBody()` —— 裸 MTEF 字节 → LaTeX，
     复用 eqn 包全部解析逻辑，不需要 OLE 包装
   - `tools/wmf2latex/main.go`：stdin 读 .bin 路径清单，stdout 出
     `path<TAB>latex`，421 个约 10 秒

复现：

```powershell
# 批量提取（Python 内联或复用 diff 管线）
python scripts/wmf_glyph_layout.py --docx rebuild-assets/external/fraction-split-reference.docx `
  --out target/wmf-ruler/glyph-layouts.json
# 批量转换（需 Go，模块副本在 target/docxtolatex-copy）
cd target/docxtolatex-copy
go run ./tools/wmf2latex < bins.txt > ../wmf-ruler/wmf-latex.tsv
```

## 双源对比的副产品：陈旧预览实锤

第一轮 diff 发现的"配对异常"（image372–397 区段，预览字形与 LaTeX 数字不符）
现在可以定性了：**WMF 内嵌 MTEF 与 OLE MTEF 一致，但预览画面本身没更新**——
即原作者编辑了公式而 Word/MathType 保留了旧预览图。
转换器（Go）没有问题，是语料里有陈旧预览。

这给了标尺一个新用途：**语料质量审计**——凡"预览字形层"与"内嵌 MTEF 层"
打架的样本，自动标记为陈旧预览，训练/校准时应剔除。

## 意义

1. **公式恢复**：只有图片/预览没有 OLE 的场景（扫描件重组、历史文档抢救）
   可以直接从 WMF 还原可编辑 LaTeX。
2. **配对简化**：标尺管线不再依赖 docx 容器——任意一堆 WMF 就能建立
   "LaTeX ↔ 字形真值"配对，xsc 全语料铺开时配对成本更低。
3. **转换器交叉验证**：docxtolatex 的 MTEF 解析在两个独立数据源上
   输出一致（425/425），解析器可靠性获得实证。
