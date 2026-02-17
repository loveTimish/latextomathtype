# Paper-to-Word 技术文档

## 项目概述

本模块将试卷题目（含 LaTeX 数学公式）导出为 Word 文档（.docx），其中数学公式以 **MathType OLE 对象**嵌入，支持双击打开 MathType 编辑器进行编辑。

---

## 目录

1. [整体架构与数据流](#1-整体架构与数据流)
2. [文件结构与职责](#2-文件结构与职责)
3. [LaTeX 解析管线](#3-latex-解析管线)
4. [MTEF 二进制格式详解](#4-mtef-二进制格式详解)
5. [OLE 复合文档打包](#5-ole-复合文档打包)
6. [预览图生成与 VML 嵌入](#6-预览图生成与-vml-嵌入)
7. [双击编辑原理](#7-双击编辑原理)
8. [关键设计决策](#8-关键设计决策)

---

## 1. 整体架构与数据流

```
┌─────────────────────────────────────────────────────────────┐
│                   数据流总览                                  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  题目 HTML（含 $\frac{x^2}{a^2}$ 公式）                      │
│       │                                                     │
│       ▼                                                     │
│  ┌─────────────────┐                                        │
│  │  DocxBuilder     │  提取 LaTeX 公式，组装 Word 文档        │
│  └────────┬────────┘                                        │
│           │                                                 │
│           ▼                                                 │
│  ┌─────────────────┐                                        │
│  │  LaTeXParser     │  LaTeX 字符串 → 抽象语法树 (AST)       │
│  │  LaTeXTokenizer  │  词法分析：分词                         │
│  │  LaTeXNode       │  AST 节点定义                          │
│  └────────┬────────┘                                        │
│           │                                                 │
│           ▼                                                 │
│  ┌─────────────────┐                                        │
│  │  MtefWriter      │  AST → MTEF v5 二进制数据              │
│  │  MtefCharMap     │  字符映射（LaTeX→MTEF typeface/mtcode） │
│  │  MtefTemplateBuilder│ TMPL 记录构建器                     │
│  │  MtefRecord      │  MTEF 记录类型常量                     │
│  └────────┬────────┘                                        │
│           │                                                 │
│           ▼                                                 │
│  ┌─────────────────┐                                        │
│  │  OlePackager     │  MTEF → OLE2 复合文档 (.bin)           │
│  └────────┬────────┘                                        │
│           │                                                 │
│           ▼                                                 │
│  ┌─────────────────┐    ┌──────────────────┐                │
│  │  MathTypeEmbedder│◄───│ LaTeXImageRenderer│               │
│  │  （嵌入协调器）    │    │ （预览图渲染）     │               │
│  └────────┬────────┘    └──────────────────┘                │
│           │                                                 │
│           ▼                                                 │
│       .docx 文件                                             │
│   ┌─────────────────────────────┐                           │
│   │ word/document.xml           │ VML Shape + OLEObject 引用 │
│   │ word/embeddings/oleObject.bin│ OLE2 复合文档（含 MTEF）   │
│   │ word/media/image_eq.png     │ 公式预览图                 │
│   └─────────────────────────────┘                           │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 2. 文件结构与职责

### 核心模块

| 文件 | 职责 |
|------|------|
| `core/latex/LaTeXTokenizer.java` | 词法分析：将 LaTeX 字符串拆分为 Token 序列（命令、花括号、字符等） |
| `core/latex/LaTeXParser.java` | 语法分析：递归下降解析器，将 Token 序列构建为 AST |
| `core/latex/LaTeXNode.java` | AST 节点定义：ROOT、GROUP、CHAR、COMMAND、FRACTION、SQRT、SUPERSCRIPT、SUBSCRIPT 等类型 |
| `core/mtef/MtefRecord.java` | MTEF v5 记录类型常量定义（END=0, LINE=1, CHAR=2, TMPL=3 等） |
| `core/mtef/MtefCharMap.java` | 字符映射表：LaTeX 字符/命令 → MTEF typeface + Unicode MTcode |
| `core/mtef/MtefTemplateBuilder.java` | TMPL 记录头部构建器：分数、根号、上下标、括号等模板的字节序列生成 |
| `core/mtef/MtefWriter.java` | **核心**：AST → MTEF v5 二进制转换，处理所有数学结构的编码 |
| `core/ole/OlePackager.java` | OLE2 复合文档打包：将 MTEF 数据封装为 MathType 可识别的 OLE 对象 |
| `core/render/LaTeXImageRenderer.java` | 公式预览图渲染：JLaTeXMath → PNG（3 倍分辨率） |
| `core/docx/MathTypeEmbedder.java` | 嵌入协调器：组装 OLE 对象 + 预览图 → Word 文档 |
| `core/docx/DocxBuilder.java` | 文档构建器：解析题目 HTML，组装完整 .docx |

### 辅助模块

| 文件 | 职责 |
|------|------|
| `controller/ExportController.java` | REST API 接口：接收导出请求 |
| `service/PaperExportService.java` | 业务服务：调用 DocxBuilder |
| `model/QuestionDTO.java` | 题目数据传输对象 |
| `model/SectionDTO.java` | 试卷分节数据 |
| `model/PaperExportRequest.java` | 导出请求 |

---

## 3. LaTeX 解析管线

### 3.1 词法分析（LaTeXTokenizer）

将 LaTeX 字符串逐字符扫描，生成 Token 序列：

```
输入: \frac{x^{2}}{a+b}

Token序列:
  COMMAND(\frac) → LBRACE → CHAR(x) → CARET → LBRACE → CHAR(2) → RBRACE
  → RBRACE → LBRACE → CHAR(a) → CHAR(+) → CHAR(b) → RBRACE
```

**分词规则：**
- `\` 开头：命令（贪婪匹配已知命令名）
- `{` / `}` ：花括号组
- `^` / `_` ：上标/下标运算符
- `[` / `]` ：方括号
- 空白：跳过
- 其他：普通字符

### 3.2 语法分析（LaTeXParser）

递归下降解析器，将 Token 构建为 AST：

```
输入: \frac{x^{2}}{a+b}

AST:
  FRACTION
  ├── GROUP (分子)
  │   └── SUPERSCRIPT
  │       ├── CHAR 'x'     (底数)
  │       └── GROUP         (指数)
  │           └── CHAR '2'
  └── GROUP (分母)
      ├── CHAR 'a'
      ├── CHAR '+'
      └── CHAR 'b'
```

**解析优先级（从高到低）：**
1. `parseAtom` — 单个字符、命令、花括号组
2. `parseScripts` — 处理 `^` 和 `_` 运算符（右结合）
3. `parseCommand` — 处理 `\frac{}{}`、`\sqrt{}`、`\left(` 等带参数命令
4. `parseExpression` — 顶层表达式，串联多个元素

---

## 4. MTEF 二进制格式详解

### 4.1 MTEF v5 文件结构

MTEF（MathType Equation Format）是 MathType 的原生二进制公式格式：

```
┌──────────────────────────────────────────────┐
│ MTEF v5 二进制流                               │
├──────────────────────────────────────────────┤
│ 文件头 (12 字节)                                │
│   version(1)=5, platform(1)=1(Win),           │
│   product(1)=0(MathType), prodVer(1)=7,       │
│   prodSubVer(1)=0, appKey="DSMT4\0",          │
│   eqnOptions(1)=0x01                          │
├──────────────────────────────────────────────┤
│ FONT_STYLE_DEF 记录 × N                       │
│   定义 typeface → 实际字体的映射                 │
│   例如：FN_TEXT(1) → "Times New Roman"         │
├──────────────────────────────────────────────┤
│ SIZE 记录                                      │
│   FULL(0x0A) + type(101) + 12pt               │
├──────────────────────────────────────────────┤
│ 表达式 LINE                                    │
│   LINE(tag=0x01, options=0x00)                 │
│   ├── 内容记录序列（CHAR、TMPL、LINE 等）        │
│   └── END                                      │
├──────────────────────────────────────────────┤
│ END（流结束标记）                                │
└──────────────────────────────────────────────┘
```

### 4.2 记录类型

| 标签值 | 名称 | 说明 |
|--------|------|------|
| 0x00 | END | 结束标记（关闭 LINE 或 TMPL） |
| 0x01 | LINE | 行/槽位：包含一组子记录 |
| 0x02 | CHAR | 字符：typeface(1) + MTcode(2) |
| 0x03 | TMPL | 模板：分数、根号、上下标等复合结构 |
| 0x06 | EMBELL | 修饰：上划线、帽子等 |
| 0x08 | FONT_STYLE_DEF | 字体样式定义 |
| 0x09 | SIZE | 字号设置 |
| 0x0A | FULL | 恢复完整字号 |
| 0x0B | SUB | 缩小到下标字号 |
| 0x0C | SUB2 | 缩小到二级下标字号 |
| 0x0D | SYM | 符号（类似 CHAR，用于特殊符号） |

### 4.3 CHAR 记录格式

```
CHAR 记录（可变长度）:
┌──────┬──────┬──────────┬────────┬──────────┐
│ tag  │ opts │ typeface │ MTcode │ [bits8]  │
│ 0x02 │ 1B   │ 1B       │ 2B LE  │ [1B]     │
└──────┴──────┴──────────┴────────┴──────────┘

typeface 取值（高位为编码标志）：
  0x83 = FN_VARIABLE  (变量，如 x, y)
  0x86 = FN_SYMBOL    (符号，如 +, =, ×)
  0x88 = FN_NUMBER    (数字，如 1, 2, 3)
  0x96 = FN_EXPAND    (可伸缩分隔符，如括号)

MTcode：Unicode 码点（小端序 2 字节）
  例如：'x' = 0x0078, '+' = 0x002B, '×' = 0x00D7
```

### 4.4 TMPL 记录格式（模板）

模板用于表示分数、根号、上下标等复合数学结构：

```
TMPL 记录头部:
┌──────┬──────┬──────────┬───────────┬────────────┐
│ tag  │ opts │ selector │ variation │ tmplOptions│
│ 0x03 │ 0x00 │ 1B       │ 1-2B      │ 1B (=0x00) │
└──────┴──────┴──────────┴───────────┴────────────┘

selector 取值：
  1  = TM_PAREN    (圆括号)
  3  = TM_BRACK    (方括号)
  10 = TM_ROOT     (根号)
  11 = TM_FRAC     (分数)
  16 = TM_SUM      (求和)
  27 = TM_SUB      (下标)
  28 = TM_SUP      (上标)
  29 = TM_SUBSUP   (上下标)

variation 变长编码：
  若 bit7=0: 单字节 variation
  若 bit7=1: 双字节 (第 2 字节为 variation 高位)

tmplOptions: 始终为 0x00（MathType 实际输出中所有模板类型都包含此字节）
```

### 4.5 上下标模板的特殊结构

MathType 中上下标的 MTEF 结构与直觉不同——**底数字符在模板之前**写入：

```
x² 的 MTEF 结构:
┌─────────────────────────────────┐
│ CHAR 'x' (FN_VARIABLE)         │ ← 底数在模板外部
│ TMPL TM_SUP (selector=28)      │
│   SUB                           │ ← 切换到下标字号
│   LINE [NULL]                   │ ← slot 0 = 空（底数位置）
│   LINE                          │ ← slot 1 = 上标内容
│     CHAR '2' (FN_NUMBER)       │
│   END                           │
│ END                             │
└─────────────────────────────────┘

a₁ 的 MTEF 结构（注意 slot 顺序与 TM_SUP 相反）:
┌─────────────────────────────────┐
│ CHAR 'a' (FN_VARIABLE)         │ ← 底数在模板外部
│ TMPL TM_SUB (selector=27)      │
│   SUB                           │
│   LINE                          │ ← slot 0 = 下标内容（先写！）
│     CHAR '1' (FN_NUMBER)       │
│   END                           │
│   LINE [NULL]                   │ ← slot 1 = 空（底数位置）
│ END                             │
└─────────────────────────────────┘
```

### 4.6 FULL 记录管理规则

`FULL`（0x0A）是一个单字节记录，用于恢复字号到完整大小。使用规则：

1. **模板后加 FULL**：当 TM_SUP、TM_SUB、TM_ROOT 等改变字号的模板结束后，如果同一行中还有**后续内容**，则添加 FULL
2. **不在 slot 末尾添加**：如果模板是当前行的最后一个元素，不添加 FULL
3. **分数分子分母之间**：仅当分子最后一个元素是 TM_SUP/TM_SUB/TM_ROOT 时才添加 FULL；如果最后是 TM_PAREN（内部已有 FULL），则不添加
4. **TM_PAREN 内部**：content LINE END 与 FN_EXPAND 分隔符之间，仅当内容最后一个元素是模板时才添加 FULL

```
示例：x²/a² + y²/b² = 1 中的 FULL 位置：

FRAC₁ {
  LINE (分子) → x → TM_SUP('2') → END   (无 FULL：TM_SUP 是最后元素)
  FULL                                      ← 分子/分母之间需要 FULL
  LINE (分母) → a → TM_SUP('2') → END
}
END
FULL                                        ← FRAC₁ 后有后续内容 '+'
CHAR '+'
FRAC₂ { ... }
END
FULL                                        ← FRAC₂ 后有后续内容 '='
CHAR '='
CHAR '1'
```

---

## 5. OLE 复合文档打包

### 5.1 OLE2 结构

MathType 公式在 Word 中以 OLE2 复合文档（Structured Storage）形式存储：

```
oleObjectN.bin (OLE2 Compound Document)
├── \001Ole              (20 字节: OLE 嵌入标记)
├── \001CompObj          (COM 对象标识)
│   ├── UserType: "MathType 7.0 Equation"
│   └── ProgID:   "Equation.DSMT4"
├── Equation Native      (MTEF 数据)
│   ├── EQNOLEFILEHDR (28 字节头部)
│   │   ├── cbHdr = 28
│   │   ├── version = 0x00020000
│   │   ├── cf = 0 (剪贴板格式)
│   │   └── cbObject = MTEF数据长度
│   └── MTEF v5 二进制数据
├── \003ObjInfo          (6 字节: 对象显示信息)
└── [CLSID: {0002CE03-...}]  (MathType 的 COM 类标识)
```

### 5.2 打包策略

**优先方案：模板替换法**
1. 从 classpath 加载真实 MathType OLE 模板文件
2. 仅替换 `Equation Native` 流中的 MTEF 数据
3. 删除过期的 `\002OlePres000` 预览缓存
4. 保留其他流（CompObj、ObjInfo 等）不变

**降级方案：从零构建**
如果模板不可用，手动构建所有 4 个流

模板方案的优势：保留了 MathType 生成的辅助数据和精确的字节格式，最大限度提高兼容性。

---

## 6. 预览图生成与 VML 嵌入

### 6.1 预览图

Word 文档中 OLE 对象的显示依赖一张**预览图**（嵌入在 VML Shape 中）。我们生成高分辨率 PNG 作为初始预览：

```
渲染配置：
  引擎：JLaTeXMath
  样式：STYLE_DISPLAY（显示模式，分数/根号以全尺寸渲染）
  倍率：3x 渲染 + 1x 报告（确保细节清晰）
  格式：PNG（位图，100% 保留分数线等细节）

当用户在 Word 中打开文档后：
  MathType 自动用精确的 WMF 矢量预览替换初始 PNG
  → 最终效果与 MathType 原生创建的公式完全一致
```

### 6.2 VML 嵌入结构

每个公式在 document.xml 中以 VML（Vector Markup Language）+ OLE 引用的形式存在：

```xml
<w:object>
  <!-- 形状模板：定义 OLE 对象的渲染方式 -->
  <v:shapetype id="_x0000_t75" ... />

  <!-- 预览图形状：引用 PNG/WMF 图片 -->
  <v:shape style="width:57pt;height:33pt" o:ole="">
    <v:imagedata r:id="rImgN" />     ← 指向 word/media/image_eqN.png
  </v:shape>

  <!-- OLE 对象引用 -->
  <o:OLEObject
    Type="Embed"
    ProgID="Equation.DSMT4"           ← MathType 的 COM ProgID
    r:id="rOleN"                       ← 指向 word/embeddings/oleObjectN.bin
  />
</w:object>
```

### 6.3 OPC 包结构

.docx 是一个 ZIP 包，包含以下公式相关文件：

```
docx.zip/
├── word/
│   ├── document.xml                    VML Shape + OLEObject XML
│   ├── embeddings/
│   │   ├── oleObject1.bin              OLE2 复合文档（含 MTEF）
│   │   ├── oleObject2.bin
│   │   └── ...
│   ├── media/
│   │   ├── image_eq1.png              公式预览图
│   │   ├── image_eq2.png
│   │   └── ...
│   └── _rels/
│       └── document.xml.rels          关系文件（链接 OLE 和图片）
└── [Content_Types].xml
```

---

## 7. 双击编辑原理

当用户在 Word 中双击公式时，触发以下链式机制：

```
1. 用户双击 v:shape（预览图）
       │
       ▼
2. Word 识别 o:OLEObject (Type="Embed")
   读取 ProgID="Equation.DSMT4"
       │
       ▼
3. Word 通过 Windows COM 注册表查找 "Equation.DSMT4"
   → 定位到 MathType 7 的 COM 服务器
   → 验证 CLSID = {0002CE03-0000-0000-C000-000000000046}
       │
       ▼
4. Word 激活 MathType COM 服务器（就地激活 / In-Place Activation）
   → MathType 接管 Word 窗口的一部分区域
       │
       ▼
5. MathType 读取 oleObjectN.bin 中的 "Equation Native" 流
   → 解析 EQNOLEFILEHDR（28 字节头部）
   → 提取 MTEF v5 二进制数据
   → 解码为内部公式表示
       │
       ▼
6. MathType 编辑器显示公式，用户可以编辑
       │
       ▼
7. 用户关闭 MathType 编辑器
   → MathType 将修改后的公式重新编码为 MTEF
   → 更新 "Equation Native" 流
   → 生成新的 WMF 矢量预览图
   → Word 更新 v:imagedata 引用的预览图
```

**关键条件：**
- ProgID 必须为 `"Equation.DSMT4"`（MathType 7 的标识）
- CLSID 必须为 `{0002CE03-...}`（MathType 的 COM 类标识）
- 系统需安装 MathType 7 并正确注册 COM 组件
- MTEF 数据必须是有效的 v5 格式（否则 MathType 无法解析）

---

## 8. 关键设计决策

### 8.1 MTEF 模板字节的 tmplOptions

MTEF v5 规范声称 `tmplOptions` 字节仅存在于 fence（括号）和 integral（积分）模板中。但通过分析 MathType 7 的实际输出，发现**所有模板类型都包含 tmplOptions 字节**（通常为 0x00）。本实现遵循 MathType 的实际行为而非规范描述。

### 8.2 上下标底数外置

MTEF 中上下标的底数字符写在 TMPL 记录**之前**（而非内部）。模板内部的 slot 0 是一个 NULL LINE（空行），底数在父级 LINE 中紧邻模板之前输出。

### 8.3 FULL 记录的条件性

FULL 记录不是无条件添加的。错误的 FULL 位置会导致：
- slot 内末尾多余的 FULL → 分数分母出现字号异常
- 缺少 FULL → 后续内容字号继承上标的缩小字号

本实现通过 `writeContentNodes` 共享方法统一管理 FULL 插入逻辑。

### 8.4 预览图策略

初始预览使用 JLaTeXMath 渲染的高分辨率 PNG（3 倍缩放）。虽然 EMF 矢量格式理论上更优，但 freehep EMF 库存在极细矩形（分数线）消失的 bug。PNG 方案保证所有公式元素正确显示。

当用户在 Word 中打开文档时，MathType 会自动用精确的 WMF 矢量预览替换 PNG，最终达到与 MathType 原生创建一致的效果。

### 8.5 OLE 模板优先

打包 OLE 对象时优先使用真实 MathType 模板（从 classpath 加载），仅替换 MTEF 数据。这比从零构建更可靠，因为模板保留了 MathType 的专有辅助数据。

---

## 附录：常用 LaTeX 公式 ↔ MTEF 结构对照

| LaTeX | MTEF 结构 |
|-------|-----------|
| `\frac{a}{b}` | `TMPL(TM_FRAC) → LINE(a) → LINE(b) → END` |
| `x^{2}` | `CHAR(x) → TMPL(TM_SUP) → SUB → NULL_LINE → LINE(2) → END` |
| `a_{n}` | `CHAR(a) → TMPL(TM_SUB) → SUB → LINE(n) → NULL_LINE → END` |
| `\sqrt{x}` | `TMPL(TM_ROOT,var=0) → LINE(x) → SUB → NULL_LINE → END` |
| `\sqrt[3]{x}` | `TMPL(TM_ROOT,var=1) → LINE(x) → SUB → LINE(3) → END` |
| `(a+b)` | `TMPL(TM_PAREN) → LINE(a+b) → [FULL] → CHAR_EXP(() → CHAR_EXP()) → END` |
| `\sum_{k=1}^{n}` | `TMPL(TM_SUM,var=0x30) → LINE(content) → SUB → LINE(k=1) → LINE(n) → SYM → END` |
