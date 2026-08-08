# latextomathtype

[English](README.md)

将试卷数据和 LaTeX 公式导出为 Word `.docx`，并把公式保留为可编辑的 MathType OLE 对象。

这不是公式截图生成器。服务会写入 MathType 兼容 OLE 对象，生成 Word 可见预览，并通过 OLE 检查、Word/MathType 抽查和 `docx2tex` 回切验证结果。

```text
PaperExportRequest -> LaTeX 解析 -> Math IR -> MTEF v5 -> OLE2 -> MathJax/Batik EMF+ Dual 预览 -> DOCX
```

## 亮点

- Java 21 / Spring Boot 的 Word 试卷导出服务。
- `POST /api/export/word` 返回带可编辑 MathType 公式的 `.docx`。
- 纯 Java MTEF/OLE 写入路径；服务端不依赖桌面 MathType。
- 公式本体、严格矢量 EMF+ Dual 预览和 Word 显示框分层处理。
- 默认预览链路由 Batik 将 MathJax SVG 全部轮廓化，再写入相互匹配的 EMF+ 与经典 EMF 路径；不含位图或文本记录，也不依赖播放端字体。
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
| EMF+ Dual 矢量预览 | OLE 对象打开前 Word 绘制的内容；MathJax/Batik 轮廓同时写入 EMF+ 与经典 EMF 路径，禁止位图和文本记录 |
| Word 显示框 | 页面中的可见宽度、高度和基线 |

因此，版式调校主要发生在 Word 对象外壳，而不是把固定点数字号强塞进 MTEF 本体。

## 已知限制

- 完整长除法竖式尚未实现。`\longdiv[商]{除数}{被除数}` 当前只能保留可编辑的商、除数和被除数头部；后接 `array` 只是调用方手写的相邻结构，不能宣称为已经验收的长除法语义。逐步乘减、数字下移、横线对齐和余数定位均不受支持，生产输入应使用普通除法或显式商余关系式。

## 验证

主参考文档回环验证：

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

该脚本会对重建出的参考 DOCX 做 OLE 检查、可用时的 Word/MathType 探测、`docx2tex` 覆盖检查和公式显示框对比。

xsc 全量验收流程见 [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md)。最近一次记录覆盖全部 `551` 份重建源文档和 `93,319` 个按 trace 对齐的公式对象：`93,319` 个对象全部通过 OLE 校验、MTEF 解析/平衡/规范化结构匹配和严格 EMF+ Dual 校验，失败文档与失败公式均为 `0`。报告另行记录了 `37` 次源 `U+FFFD` 替换字符的显式恢复，没有静默输出这些乱码。

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
