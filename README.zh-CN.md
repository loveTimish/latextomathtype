# latextomathtype

[English](README.md)

将试卷数据和 LaTeX 公式导出为 Word `.docx`，并把公式保留为可编辑的 MathType OLE 对象。

这不是公式截图生成器。服务会写入 MathType 兼容 OLE 对象，生成 Word 可见预览，并通过 OLE 检查、Word/MathType 抽查和 `docx2tex` 回切验证结果。

```text
PaperExportRequest -> LaTeX 解析 -> Math IR -> MTEF v5 -> OLE2 -> MathJax/Batik POLYPOLYGON WMF 预览 -> DOCX
```

## 亮点

- Java 21 / Spring Boot 的 Word 试卷导出服务。
- `POST /api/export/word` 返回带可编辑 MathType 公式的 `.docx`。
- 纯 Java MTEF/OLE 写入路径；服务端不依赖桌面 MathType。
- 公式本体、严格矢量 WMF 预览和 Word 显示框分层处理。
- 默认预览链路由 Batik 将 MathJax SVG 全部轮廓化，再写入经典 `POLYPOLYGON` WMF；不含位图或文本记录，也不依赖播放端字体。
- 支持 Linux 运行；最终 GUI 可编辑性抽查使用 Windows + Word + MathType。

## 快速开始

本地开发时安装锁定的 MathJax 依赖：

```powershell
npm install
npm run mathjax:smoke
```

发行部署使用离线 sidecar，运行时不会执行 npm 或访问网络：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-vector-sidecar.ps1 -Platform all
```

把 `vector-sidecar-windows-x64.zip` 或 `vector-sidecar-linux-x64.tar.gz` 解压到应用 JAR
旁边，并确保目录名为 `vector-sidecar`。运行时会校验 Node `v24.9.0`、MathJax
`3.2.2` 和 worker bundle 哈希。

生成包含可执行 JAR 和离线 sidecar 的完整 Windows x64、Linux x64 发布包：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-release.ps1 -Platform all
```

发布包、版本清单和 `SHA256SUMS.txt` 会写入 `target/release/<版本号>/`。

构建、测试并启动服务：

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd clean package
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd test
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd spring-boot:run
```

导出示例试卷：

```powershell
Invoke-WebRequest `
  -Method Post `
  -Uri http://127.0.0.1:8081/api/export/word `
  -ContentType 'application/json; charset=utf-8' `
  -InFile exam-template.json `
  -OutFile target\paper.docx
```

## API

| 接口 | 方法 | 用途 |
| --- | --- | --- |
| `/api/export/health` | `GET` | 服务健康检查 |
| `/api/export/word` | `POST` | 将 `PaperExportRequest` 导出为 Word |
| `/api/export/layout-word` | `POST` | 将块级版式数据导出为 Word |

请求结构见 [exam-template.json](exam-template.json)。

## 核心设计

项目刻意把三个层次分开：

| 层次 | 控制内容 |
| --- | --- |
| MTEF/OLE 本体 | MathType 可编辑性和公式语义 |
| POLYPOLYGON WMF 矢量预览 | OLE 对象打开前 Word 绘制的内容；MathJax/Batik 轮廓写入经典 WMF 多边形，禁止位图和文本记录 |
| Word 显示框 | 页面中的可见宽度、高度和基线 |

因此，版式调校主要发生在 Word 对象外壳，而不是把固定点数字号强塞进 MTEF 本体。

## 已知限制

- 结构化长除法使用 `\begin{longdivision}{r+}{除数}{商}{被除数}...\end{longdivision}`。每一步和 `\cline{m-n}` 横线范围均由调用方显式提供；项目只负责排版，不计算或校验算术。旧 `\longdiv` 命令仅作为头部兼容输入保留。

## 验证

主参考文档回环验证：

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

该脚本会对重建出的参考 DOCX 做 OLE 检查、可用时的 Word/MathType 探测、`docx2tex` 覆盖检查和公式显示框对比。

xsc 全量验收流程见 [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md)。历史语料结果保留在该文档中；当前验收要求公式具有可编辑 OLE、平衡且可规范化的 MTEF，以及不含位图和文本记录的纯 `POLYPOLYGON` WMF。未登记的 `U+FFFD` 输入会直接失败，不会静默输出乱码。

## Linux 和 Docker

在 Linux 上打包并运行：

```bash
./.mvn/apache-maven-3.9.12/bin/mvn -DskipTests package
java \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render \
  -jar target/paper-to-word-1.0.0.jar
```

Docker 和打包细节见 [docs/linux-runtime.md](docs/linux-runtime.md)。

## 项目结构

```text
src/main/java/com/lz/paperword
  controller/ service/ model/
  core/latex/ core/mathml/ core/mtef/ core/ole/ core/render/ core/docx/

scripts/   验证和语料工具
rebuild/   参考文档重建工具
docs/      技术说明和验收计划
```

## 文档

- [TECHNICAL.md](TECHNICAL.md)：架构和实现细节。
- [docs/MathType-validation-plan.md](docs/MathType-validation-plan.md)：验证层次和阶段门槛。
- [docs/MathType-support-matrix.md](docs/MathType-support-matrix.md)：已支持公式结构和缺口。
- [docs/linux-runtime.md](docs/linux-runtime.md)：Linux 和 Docker 运行说明。
- [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md)：xsc 语料流水线说明。

## 参考

- [transpect/docx2tex](https://github.com/transpect/docx2tex)
- [MathJax](https://github.com/mathjax/MathJax-src)
- [WIRIS MathType SDK: MTEF storage](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/how-mtef-is-stored-in-files-and-objects)
- [WIRIS MathType SDK: MTEF v5](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/mathtype-mtef-v5-mathtype-40-and-later)
