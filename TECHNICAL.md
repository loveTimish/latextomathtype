# 技术设计

本文档说明 `latextomathtype` 当前的导出链路：`PaperExportRequest` 如何生成 Word，LaTeX 如何变成可编辑 MathType OLE，公式预览如何渲染，以及参考文档如何验收。

## 目标

- 为试卷和版式重建场景生成 Word `.docx`。
- 公式必须保留为 MathType 兼容 OLE 对象，而不是只有图片。
- 生成文档必须能被 `docx2tex` 回切出正确 LaTeX。
- 公式在 Word 页面中的大小、基线和位置尽量贴近参考文档。
- 核心导出链路必须能在 Linux 上运行，不依赖桌面 MathType。

非目标：

- 不追求和参考 `.docx` 字节级完全一致。
- 不把 Windows 桌面 MathType 作为 Linux 导出依赖。
- 不用 MTEF 里的固定字号记录作为 Word 版式的主要控制手段。

## 总体管线

```text
PaperExportRequest
  -> DocxBuilder
  -> LaTeXParser.splitContent()
  -> LaTeXParser.parse()
  -> MathIRConverter / MathIRLowerer
  -> MtefWriter
  -> OlePackager
  -> LaTeXImageRenderer
  -> MathTypeEmbedder
  -> DOCX package
```

核心职责：

| 模块 | 职责 |
| --- | --- |
| `DocxBuilder` | 组装页面结构、段落、标题、题目、选项、图片和公式 run。 |
| `LaTeXParser` | 将 HTML/文本拆成普通文本和公式片段，并把支持的 LaTeX 解析为 AST。 |
| `MathIRConverter` / `MathIRLowerer` | 将 AST 规范化为更适合写 MTEF 的中间数学表示。 |
| `MtefWriter` | 写 MathType MTEF v5 记录，包括模板、字符、pile、matrix 和竖式相关结构。 |
| `OlePackager` | 将 MTEF 包装为 MathType OLE2 复合对象。 |
| `LaTeXImageRenderer` | 优先通过 TeX/dvisvgm 生成公式预览图，缺失时使用 Java 回退渲染。 |
| `MathTypeEmbedder` | 把预览图、OLE 二进制、VML shape、关系 ID、尺寸和基线写入 Word。 |

## DOCX 对象结构

每个公式有两个层次。

可编辑本体在：

```text
word/embeddings/oleObjectN.bin
```

这是一个 OLE2 复合文档，内部包含：

```text
\001Ole
\001CompObj
Equation Native
\003ObjInfo
```

`Equation Native` 保存 MathType 原生数据。WIRIS 文档说明，OLE equation native stream 使用 28 字节的 `EQNOLEFILEHDR`，后面跟 MTEF 数据。这个头部包含头长度、版本、剪贴板格式、MTEF 字节长度和保留字段。

Word 页面上的显示层在：

```text
word/document.xml
word/media/imageN.png
```

显示层控制：

- 预览图 relationship
- `v:shape` 宽高
- `w:dxaOrig` 和 `w:dyaOrig`
- `w:position` 基线偏移
- run 和段落位置

这是当前实现最重要的边界：MathType 可编辑性由 OLE/MTEF 决定，Word 页面上的大小和对齐由外层对象显示框决定。

## MTEF 写入

写入器目标是 MathType MTEF v5。WIRIS 文档描述的 v5 流大致为：

```text
MTEF header
equation preferences and definitions
initial SIZE/FULL record
PILE or LINE
contents
END
```

当前实现写入一个紧凑的 MTEF 流：

- 版本、平台、产品头和 `DSMT4` app key
- 默认 typesize 标记
- 顶层 `LINE`
- 按需嵌套 `TMPL`、`LINE`、`MATRIX`、`PILE`、`RULER`、`CHAR`
- 配平的 `END` 记录

支持结构包括：

- 普通变量、数字、运算符、希腊字母和常用数学符号
- 分数、根号、上标、下标、上下标
- 大运算符、极限、积分
- 括号和 fence 模板
- 矩阵、cases、aligned 结构
- 项目里使用的 K12 竖式辅助结构

字符映射集中在 `MtefCharMap`。乘号不能只当作普通 Unicode 字符写入；MathType 编辑器和 `docx2tex` 对 MTEF typeface 与字符码敏感，所以乘号需要映射到 MathType/Symbol 兼容记录。

写入器不会为普通公式写入显式固定点数字号记录。默认初始尺寸只写 full size class。Word 版式校准通过外层 VML/OLE 显示框完成，因为改 MTEF 内部字号会影响 MathType 编辑和回切语义，风险更大。

## OLE 打包

`OlePackager` 使用 Apache POI POIFS 写 OLE2 复合文件，使对象看起来像 MathType equation object：

- `Equation Native` 中写入原生 MTEF payload
- 写入 `MathType EF` native stream header
- 写入兼容 MathType equation object 的 class metadata
- 保留 Word/MathType 识别所需的 storage 名称和对象信息

验收时直接检查 OLE 复合文件：

- `Equation Native` 必须存在。
- native header 长度必须正确。
- header 中声明的对象长度必须等于后续 MTEF 字节数。
- 复合对象能作为 POIFS 打开。
- Windows GUI 验证可用时，Word 能识别 `Equation.DSMT4` 等 OLE class 名称。

## 预览渲染

`LaTeXImageRenderer` 优先使用原生 TeX 管线：

```text
latex -> dvisvgm -> SVG -> PNG
```

这条路径在真实 TeX 表达式上尺寸更可控，是推荐的 Linux 生产配置。如果缺少 `latex` 或 `dvisvgm`，会回退到 JLaTeXMath。回退路径能保证可生成，但视觉尺寸可能和 TeX/MathType 有差异。

渲染器有两层缓存：

- JVM 进程内内存缓存
- 可选持久化磁盘缓存，按规范化公式、渲染模式、尺寸和缓存版本生成 key

相关配置：

| 配置 | 含义 |
| --- | --- |
| `paperword.latex.command` | 原生 TeX 命令路径 |
| `paperword.dvisvgm.command` | 原生 dvisvgm 命令路径 |
| `paperword.latex.timeout.seconds` | 渲染超时 |
| `paperword.render.cache.enabled` | 持久化缓存开关 |
| `paperword.render.cache.dir` | 持久化缓存目录 |

## 版式校准

参考文档匹配不能靠改公式字体解决。Word 中每个公式外面都有 OLE 显示框，因此生成文档校准的是：

```text
v:shape style width
v:shape style height
w:dxaOrig
w:dyaOrig
w:position
```

重建脚本按参考文档和生成文档的公式对象顺序做一一对比，然后同步显示框。这样可以让可见宽、高、基线贴近参考，同时保留生成的 `Equation Native` 可编辑本体。

当前参考重建还会在必要时同步公式预览媒体，用来排除渲染器噪声对最终版式对比的影响。这不会替换生成的 OLE 对象，只是让可见 shell 更接近参考，OLE/MTEF payload 仍由项目生成。

## 参考重建流程

参考文件：

```text
rebuild-assets/external/fraction-split-reference.docx
```

主命令：

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

流程：

1. `docx2tex` 将参考 Word 转为 LaTeX。
2. `generate_fraction_reference_request.py` 重建 `PaperExportRequest`。
3. `ReferenceRoundTripDocxTest` 调用 `DocxBuilder` 写出生成文档。
4. `extract_formula_boxes.py` 提取参考和生成文档的 OLE 显示框。
5. `build_formula_box_dataset.py` 可从 `E:\新加卷\新建文件夹\xsc资料` 收集外部样本。
6. `fit_formula_box_model.py` 输出尺寸拟合诊断。
7. `calibrate_formula_boxes.py` 应用一一对应的参考显示框校准。
8. `sync_reference_layout_shell.py` 同步页面级参考 shell 细节。
9. `sync_formula_vector_previews.py` 在需要排除渲染差异时同步公式预览媒体。
10. `verify-mathtype-word.ps1` 在 Windows 上检查 Word/MathType 识别。
11. `verify-docx2tex-roundtrip.ps1` 对生成文档运行 `docx2tex`。
12. `verify_docx2tex_formula_fragments.py` 检查公式 fragment 覆盖率。
13. `compare_reference_format.py` 输出最终结构和公式显示框对比报告。

生成文档路径：

```text
target/reference-roundtrip/fraction-split-reference-regenerated.docx
```

主要报告：

```text
target/reference-roundtrip/format-comparison.txt
target/reference-roundtrip/format-comparison.json
target/reference-roundtrip/formula-boxes-reference.json
target/reference-roundtrip/formula-boxes-generated-calibrated.json
target/reference-roundtrip/docx2tex-fragment-check.json
```

## 验收门槛

运行 Java 测试：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd test
```

只运行参考文档生成测试：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd `
  -q `
  -Dtest=com.lz.paperword.tools.ReferenceRoundTripDocxTest `
  test
```

检查 Word/MathType OLE 识别：

```powershell
.\scripts\verify-mathtype-word.ps1 `
  -DocxPath target\reference-roundtrip\fraction-split-reference-regenerated.docx `
  -MinimumOleCount 400
```

验证生成 DOCX 能回切 LaTeX：

```powershell
.\scripts\verify-docx2tex-roundtrip.ps1 `
  -DocxPath target\reference-roundtrip\fraction-split-reference-regenerated.docx `
  -RequestJsonPath target\reference-roundtrip\fraction-split-reference.request.json `
  -OutDir target\reference-roundtrip\regenerated-docx2tex `
  -MinimumTimesCount 1000 `
  -MinimumCdotsCount 150 `
  -MinimumFractionCount 1500
```

检查公式 fragment 覆盖：

```powershell
python rebuild\verify_docx2tex_formula_fragments.py `
  --request-json target\reference-roundtrip\fraction-split-reference.request.json `
  --tex target\reference-roundtrip\regenerated-docx2tex\fraction-split-reference-regenerated.tex `
  --out-json target\reference-roundtrip\docx2tex-fragment-check.json
```

`MathTypeAlignmentRegressionTest` 会检查生成 OLE 本体中是否出现显式点数字号记录。当前回归断言会确保已知的显式 size 字节模式不存在。

## Linux 边界

Linux 使用：

```text
mathtype.windows.enabled=false
```

该模式不会调用桌面 MathType，而是用 Java 直接写 MTEF/OLE，并通过 POIFS、生成 DOCX 结构、预览渲染和 `docx2tex` 输出做验证。

Windows + Word + MathType 仍然是最终 GUI 可编辑性抽查路径，因为只有真实桌面编辑器能证明目标环境中的双击编辑行为。

## 参考资料

实现依据这些公开资料和本地抓取资料：

- WIRIS MathType SDK, "How MTEF is stored in files and objects"：MTEF 如何存放在 OLE native stream 中，包括 28 字节 OLE native header。
- WIRIS MathType SDK, "MTEF v5"：v5 header、record types、`LINE`、`CHAR`、`TMPL`、`PILE`、`MATRIX`、`SIZE` 等记录规则。
- transpect `docx2tex` README 和模块文档：命令行转换模型、`-m ole|wmf|ole+wmf` 来源选择和 DOCX 到 LaTeX 管线。

本地说明和验收计划：

```text
docs/MathType-validation-plan.md
docs/linux-runtime.md
docs/reference/mathtype/
```

## 已知风险

- 解析器只支持当前测试和参考数据覆盖到的 LaTeX 子集；不支持结构必须显式失败或显式回退。
- 缺少 TeX/dvisvgm 时，JLaTeXMath 回退预览可能存在可见尺寸偏差。
- 参考匹配依赖稳定的公式对象顺序。文档构造顺序变化后必须重跑对比脚本。
- `docx2tex` 能验证语义可回切，但不能替代 MathType GUI 编辑抽查。
