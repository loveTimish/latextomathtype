# latextomathtype

`latextomathtype` 用来把试卷数据和 LaTeX 公式导出为 Word `.docx`，并把公式嵌入为可编辑的 MathType OLE 对象。当前项目的核心验收链路是：

```text
参考 DOCX -> docx2tex -> PaperExportRequest -> 重新生成 DOCX -> MathType/OLE/docx2tex/版式检查
```

这不是一个只生成公式图片的工具。它的主要目标是：生成的公式在 Word 中仍然是 MathType OLE，Windows 上可以用 MathType 打开编辑，同时生成文档还能被 `docx2tex` 切回正确 LaTeX。

## 能力

- 通过 `POST /api/export/word` 把 `PaperExportRequest` 导出为 Word。
- 将公式写入 MathType 兼容的 OLE2 对象，并在 `Equation Native` 流中保存 MTEF 数据。
- 将“公式可编辑本体”和“Word 页面上的显示框”分开处理，避免为了调版式而在 OLE/MTEF 里硬指定公式字号。
- 使用 TeX/dvisvgm 生成公式预览，并以 WMF 媒体嵌入 OLE 显示面；OLE 预览失败时直接报错，不回退到 PNG。
- 可重复重建 `rebuild-assets/external/fraction-split-reference.docx` 参考文档。
- 用 OLE 直接检查、Word/MathType 抽查、`docx2tex` 回切和公式框尺寸对比做验收。

## 环境要求

| 工具 | 用途 |
| --- | --- |
| JDK 21 | 运行、构建和测试 |
| Maven 3.9.x | 构建和测试，仓库内置 `.mvn/apache-maven-3.9.12` |
| TeX Live 或 MiKTeX，包含 `latex` 和 `dvisvgm` | 推荐的公式预览渲染器 |
| Python 3 | 参考文档重建、版式对比和回归脚本 |
| docx2tex | DOCX 到 LaTeX 的回切验证 |
| Microsoft Word + MathType | Windows 上的公式 OLE 可编辑性抽查 |
| Docker | Linux/容器环境验收 |

Windows 验证脚本默认使用 `J:\docx2tex\d2t.bat`。如果本机路径不同，改脚本里的路径即可。

## 构建

优先使用仓库内置 Maven，避免依赖系统 Maven 配置：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd clean package
```

Linux 下对应命令：

```bash
./.mvn/apache-maven-3.9.12/bin/mvn clean package
```

运行 Java 回归测试：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd test
```

## 运行服务

启动 Spring Boot：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd spring-boot:run
```

检查服务状态：

```powershell
Invoke-WebRequest http://127.0.0.1:8081/api/export/health
```

导出一个试卷：

```powershell
Invoke-WebRequest `
  -Method Post `
  -Uri http://127.0.0.1:8081/api/export/word `
  -ContentType 'application/json; charset=utf-8' `
  -InFile exam-template.json `
  -OutFile target\paper.docx
```

主要接口：

| 接口 | 方法 | 说明 |
| --- | --- | --- |
| `/api/export/health` | `GET` | 健康检查 |
| `/api/export/word` | `POST` | 将 `PaperExportRequest` 导出为带 MathType OLE 公式的 Word |
| `/api/export/layout-word` | `POST` | 面向 OCR/PDF 重建链路的块级版式 Word 导出 |

## Linux 和 Docker

服务使用纯 Java 的 MTEF/OLE 写入路径：

```bash
java \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render \
  -jar target/paper-to-word-1.0.0.jar
```

打包后构建并运行容器：

```bash
docker build -t latextomathtype:local .
docker run --rm -p 8081:8081 \
  -v latextomathtype-cache:/var/cache/latextomathtype/formula-render \
  latextomathtype:local
```

Linux 主机快速冒烟：

```bash
sh scripts/linux-smoke.sh
```

带 TeX 工具链的参考文档验收镜像：

```powershell
docker build -f Dockerfile.render-test -t latextomathtype:render-test .
```

更多 Linux 运行说明见 [docs/linux-runtime.md](docs/linux-runtime.md)。

## 参考文档重建

当前参考版式基准：

```text
rebuild-assets/external/fraction-split-reference.docx
```

Windows 上运行完整验收链路：

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

脚本会执行：

1. 用 `docx2tex` 将参考 DOCX 转为 LaTeX。
2. 生成 `target/reference-roundtrip/fraction-split-reference.request.json`。
3. 调用 `ReferenceRoundTripDocxTest` 生成 `target/reference-roundtrip/fraction-split-reference-regenerated.docx`。
4. 提取参考文档和生成文档的公式显示框。
5. 按参考对象顺序校准 `v:shape`、`w:dxaOrig/w:dyaOrig` 和 `w:position`。
6. 在可用时通过 Word COM 检查 MathType OLE 数量和 `Equation.DSMT4` 识别。
7. 对生成 DOCX 再跑 `docx2tex`，检查公式 LaTeX 覆盖率。
8. 在 `target/reference-roundtrip` 下输出结构和版式对比报告。

最近一次认可的参考文档验收指标：

| 检查项 | 结果 |
| --- | --- |
| OLE 对象数 | 参考 `425`，生成 `425` |
| 内联公式数 | 生成 `429` |
| docx2tex 公式覆盖 | `403/403` 个公式匹配，无缺失风险公式 |
| 生成文档 docx2tex token | `\times=1110`，`\cdots=164`，`\frac=1605` |
| 成对公式显示框差异 | 校准后宽、高、基线差异为 `0` |
| OLE 本体内显式点数字号记录 | `0` |

认可输出路径：

```text
target/reference-roundtrip/fraction-split-reference-regenerated.docx
```

## xsc 全量验收

xsc 测试集用于验证“完整 DOCX 重建”，不是只导出公式。默认源目录：

```text
F:\资料\xsc资料\word_files
```

推荐流水线：

```powershell
# 1. 先用 docxtolatex 重建 LaTeX corpus
.\scripts\run_xsc_docxtolatex.ps1 `
  -Start 1 `
  -End 155 `
  -OutRoot D:\latextomathtype\analysis\xsc-latex

# 2. 从 LaTeX corpus 生成完整 PaperExportRequest
python .\scripts\make_full_batch10_requests.py `
  --start 1 `
  --end 155 `
  --latex-root D:\latextomathtype\analysis\xsc-latex `
  --out-dir D:\latextomathtype\analysis\batch10-full-requests

# 3. 分批生成 DOCX 并验收。大 corpus 建议按 10 个文件一批跑。
.\scripts\run_xsc_acceptance.ps1 `
  -Start 141 `
  -End 155 `
  -LatexRoot D:\latextomathtype\analysis\xsc-latex

# 4. 汇总所有批次验收报告，作为 1% 尺寸门槛
python .\scripts\summarize_xsc_full_acceptance.py
```

`summarize_xsc_full_acceptance.py` 默认会从 `D:\latextomathtype\analysis\acceptance-summary` 自动选择每个分段的最新报告。需要固定某次验收证据时，可传入 `--manifest`，文件格式为 JSON 数组：

```json
[
  {"stamp": "20260612-064944", "start": 141, "end": 155}
]
```

聚合脚本会检查：

- 覆盖 `1..155`，无缺段、无重复段。
- 生成侧预览全部是 WMF。
- 成对公式的 WMF 目标物理宽高均在测试集 `1%` 误差内。
- MTEF 对比对象数与尺寸对比对象数一致。

最近一次全量验收报告：

```text
D:\latextomathtype\analysis\acceptance-summary\xsc-full-acceptance.json
```

关键指标：

| 检查项 | 结果 |
| --- | --- |
| 覆盖范围 | `1..155` |
| 成对公式对象 | `39551` |
| WMF 宽度 1% 内 | `39551/39551` |
| WMF 高度 1% 内 | `39551/39551` |
| 生成侧非 WMF | `0` |
| MTEF clean pairs | `39248/39551` |
| MTEF 允许的源头/样式前缀差异 | `176`，按公式头/样式前缀差异单独统计 |
| MTEF 剩余 hard suspects | `75` |
| MTEF 剩余 low-tail failures | `51` |
| MTEF 剩余结构缺口 | `126` |

增量 MTEF writer 回归：`20260612-074744` 的 doc41-doc43 验证中，平坦 `\div` 等式链、多字符单除法和一位数乘法等式改用 MathType `TM_BOX 0x1e` 片段后，WMF 目标尺寸为 `862/862` 在 `1%` 内且非 WMF 为 `0`。doc41 hard suspects 从 `6` 降到 `3`、low-tail failures 从 `12` 降到 `6`；doc43 大数乘法保持平坦写法，hard suspects 从误泛化时的 `7` 回到 `4`。

符号语义回归：`20260612-081037` 的 doc41-doc43 验证中，特殊平坦 `\div`/`\times` 写入路径统一走命令映射，避免把 LaTeX 命令首字符 `\` 写成变量字符。doc41 四条 `164\div82=...` / `128\div64=...` 抽查从旧生成的 `0200835c00` 变为 MathType Symbol `÷` 记录 `020486f700b8`；doc41-doc43 的 WMF 目标尺寸仍为 `862/862` 在 `1%` 内且非 WMF 为 `0`。

## 验证命令

检查生成 Word 中的 MathType/OLE 对象：

```powershell
.\scripts\verify-mathtype-word.ps1 `
  -DocxPath target\reference-roundtrip\fraction-split-reference-regenerated.docx `
  -MinimumOleCount 400
```

验证生成 DOCX 仍能被 `docx2tex` 切回 LaTeX：

```powershell
.\scripts\verify-docx2tex-roundtrip.ps1 `
  -DocxPath target\reference-roundtrip\fraction-split-reference-regenerated.docx `
  -RequestJsonPath target\reference-roundtrip\fraction-split-reference.request.json `
  -OutDir target\reference-roundtrip\regenerated-docx2tex `
  -MinimumTimesCount 1000 `
  -MinimumCdotsCount 150 `
  -MinimumFractionCount 1500
```

检查请求公式和回切 LaTeX 的 fragment 覆盖：

```powershell
python rebuild\verify_docx2tex_formula_fragments.py `
  --request-json target\reference-roundtrip\fraction-split-reference.request.json `
  --tex target\reference-roundtrip\regenerated-docx2tex\fraction-split-reference-regenerated.tex `
  --out-json target\reference-roundtrip\docx2tex-fragment-check.json
```

只运行参考文档生成测试：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd `
  -q `
  -Dtest=com.lz.paperword.tools.ReferenceRoundTripDocxTest `
  test
```

## 配置项

| 配置 | 默认值 | 说明 |
| --- | --- | --- |
| `paperword.latex.command` | `latex` | 原生 TeX 渲染命令 |
| `paperword.dvisvgm.command` | `dvisvgm` | DVI 转 SVG 命令 |
| `paperword.latex.timeout.seconds` | `15` | 原生 TeX 渲染超时 |
| `paperword.render.cache.enabled` | `true` | 是否启用持久化公式渲染缓存 |
| `paperword.render.cache.dir` | 系统临时目录 | 持久化公式渲染缓存目录 |

## 项目结构

```text
src/main/java/com/lz/paperword
  controller/        REST 接口
  service/           导出服务
  core/docx/         Word 文档构建和 MathType 嵌入
  core/latex/        LaTeX 分词、解析和内容切分
  core/mathml/       中间数学表示和降级转换
  core/mtef/         MTEF v5 写入器、字符映射和模板记录
  core/ole/          OLE2 对象打包
  core/render/       TeX/dvisvgm 和回退渲染
  model/             请求 DTO

rebuild/             参考文档重建和对比脚本
rebuild-assets/      参考 DOCX 和抽取出的视觉素材
scripts/             验证脚本和 Linux 冒烟脚本
docs/                技术说明和验收计划
```

## 文档

- [TECHNICAL.md](TECHNICAL.md)：架构、MTEF/OLE、渲染边界和参考重建流程。
- [docs/linux-runtime.md](docs/linux-runtime.md)：Linux 与 Docker 运行说明。
- [docs/MathType-validation-plan.md](docs/MathType-validation-plan.md)：解析、MTEF、预览、可编辑性和 Word 显示框的验收门槛。
- [exam-template.json](exam-template.json)：最小请求示例。

实现参考：

- [WIRIS MathType SDK: How MTEF is stored in files and objects](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/how-mtef-is-stored-in-files-and-objects)
- [WIRIS MathType SDK: MTEF v5](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/mathtype-mtef-v5-mathtype-40-and-later)
- [transpect/docx2tex README](https://github.com/transpect/docx2tex)

## 边界

生成文档追求视觉和语义等价，不追求字节级相同。参考文档与生成文档的段落数量、媒体内部结构可以不同；验收重点是 OLE 对象数、公式显示框成对尺寸、MathType 可编辑性和 `docx2tex` 语义覆盖。

Linux 上没有桌面 MathType。Linux 验证依赖纯 Java OLE/MTEF 生成、POIFS 直接检查、预览渲染和 `docx2tex` 回切。Windows + Word + MathType 仍然是最终 GUI 双击编辑抽查路径。
