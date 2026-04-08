# LaTeX to MathType

将 LaTeX 数学公式转换为 MathType OLE 对象，嵌入 Word 文档（.docx），支持双击打开 MathType 编辑器编辑。

## 功能特性

- **LaTeX → MathType OLE**：将 LaTeX 公式解析为 AST，转换为 MTEF v5 二进制格式，封装为 MathType OLE 对象
- **Word 文档生成**：接收试卷 JSON 数据（题目、选项、LaTeX 公式），输出格式规范的 .docx 文件
- **双击可编辑**：生成的公式支持在 Word 中双击打开 MathType 编辑器，与 MathType 原生创建的公式完全兼容
- **高保真预览**：初始使用高分辨率 PNG 预览，打开文档后 MathType 自动替换为 WMF 矢量预览
- **K12 公式全覆盖**：分数、根号、上下标、求和、积分、矩阵、希腊字母、三角函数等

## 技术原理

详见 [TECHNICAL.md](TECHNICAL.md)，包含：
- MTEF v5 二进制格式详解
- OLE2 复合文档结构
- 预览图生成与 VML 嵌入
- 双击编辑的 COM 激活原理

## 技术栈

- Java 21 + Spring Boot 3.3.4
- Apache POI（XWPF + POIFS）
- JLaTeXMath（公式预览图渲染）
- FreeHEP（EMF 矢量图生成）
- Jsoup（HTML 解析）

## 构建与运行

```bash
# 构建
mvn clean package

# 运行
java -jar target/paper-to-word-1.0.0.jar
```

API 接口：`POST /api/export/word`

### 试卷 JSON 模板

参见 [exam-template.json](exam-template.json)，请求体结构：

| 字段 | 类型 | 说明 |
|------|------|------|
| `paper` | object | 试卷信息 |
| `paper.name` | string | 试卷名称 |
| `paper.score` | number | 总分 |
| `paper.suggestTime` | number | 建议时长（分钟） |
| `sections` | array | 大题列表 |
| `sections[].headline` | string | 大题标题（如"一、选择题"） |
| `sections[].questions` | array | 小题列表 |
| `questions[].serialNumber` | number | 题号 |
| `questions[].questionType` | number | 1=单选 2=多选 3=判断 4=填空 5=解答 6=计算 |
| `questions[].content` | string | 题干（HTML，公式用 $...$ 包裹） |
| `questions[].options` | array | 选项（选择题） |
| `questions[].correct` | string | 参考答案 |
| `questions[].score` | number | 分值 |
| `questions[].analyze` | string | 解析说明 |

## 核心管线

```
LaTeX 字符串  →  词法分析  →  语法分析  →  AST
     ↓              ↓           ↓         ↓
 "$\frac{x}{y}$"  Tokens    Parser    LaTeXNode
                                         ↓
                                    MtefWriter
                                         ↓
                                   MTEF v5 二进制
                                         ↓
                                    OlePackager
                                         ↓
                                  OLE2 复合文档 (.bin)
                                         ↓
                                  MathTypeEmbedder
                                         ↓
                                  Word 文档 (.docx)
```

## 项目结构

```
src/main/java/com/lz/paperword/
├── core/
│   ├── latex/                    # LaTeX 解析
│   │   ├── LaTeXTokenizer.java   #   词法分析器
│   │   ├── LaTeXParser.java      #   递归下降语法分析器
│   │   └── LaTeXNode.java        #   AST 节点定义
│   ├── mtef/                     # MTEF 二进制生成
│   │   ├── MtefWriter.java       #   AST → MTEF v5 转换（核心）
│   │   ├── MtefCharMap.java      #   字符映射表
│   │   ├── MtefTemplateBuilder.java  # 模板记录构建器
│   │   └── MtefRecord.java       #   记录类型常量
│   ├── ole/
│   │   └── OlePackager.java      # MTEF → OLE2 复合文档打包
│   ├── render/
│   │   └── LaTeXImageRenderer.java  # 公式预览图渲染
│   └── docx/
│       ├── DocxBuilder.java      # Word 文档构建器
│       └── MathTypeEmbedder.java # MathType OLE 嵌入协调器
├── controller/
│   └── ExportController.java     # REST API
├── service/
│   └── PaperExportService.java   # 导出业务服务
└── model/                        # 数据模型
```

## 预览图渲染配置

默认使用 JLaTeXMath 渲染高分辨率 PNG。也支持外部渲染引擎：

### MathJax HTTP（可选）

```bash
java -Dpaperword.mathjax.endpoint="http://127.0.0.1:3000/svg?tex={latex}" \
     -jar target/paper-to-word-1.0.0.jar
```

| 属性 | 默认值 | 说明 |
|------|--------|------|
| `paperword.mathjax.enabled` | `true` | 启用 MathJax HTTP |
| `paperword.mathjax.endpoint` | `https://math.vercel.app/` | MathJax 服务地址 |
| `paperword.mathjax.timeout.seconds` | `12` | 请求超时 |

### LaTeX + dvisvgm（可选）

需安装 TeX Live / MiKTeX：

```bash
java -Dpaperword.latex.command="C:\texlive\2025\bin\windows\latex.exe" \
     -Dpaperword.dvisvgm.command="C:\texlive\2025\bin\windows\dvisvgm.exe" \
     -jar target/paper-to-word-1.0.0.jar
```

## 作为 Maven 依赖

```xml
<dependency>
    <groupId>com.lz</groupId>
    <artifactId>paper-to-word</artifactId>
    <version>1.0.0</version>
</dependency>
```

```java
// 代码调用
PaperExportService service = new PaperExportService();
byte[] docxBytes = service.export(request);
```


