package com.lz.paperword.core.mtef;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutCompiler;
import com.lz.paperword.core.layout.VerticalLayoutNodeFactory;
import com.lz.paperword.core.layout.VerticalLayoutSpec;
import com.lz.paperword.core.mathml.MathIRConverter;
import com.lz.paperword.core.mathml.MathIRLowerer;
import com.lz.paperword.core.mathml.MathIRNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;

/**
 * 将 LaTeX 抽象语法树 (AST) 转换为 MTEF v5 二进制数据。
 *
 * <h2>完整转换管线 (Pipeline)</h2>
 * <pre>
 *   LaTeX 字符串
 *     → LaTeXParser 解析为语法树
 *       → LaTeXNode AST（抽象语法树，树形结构）
 *         → MtefWriter（本类）将 AST 递归遍历，输出 MTEF v5 二进制字节流
 *           → OLE Equation Native 封装（由 OleWriter 完成）
 *             → 嵌入 Word 文档
 * </pre>
 *
 * <h2>MTEF v5 二进制结构概览</h2>
 * <pre>
 *   ┌─────────────────────────────────────────────────────────────┐
 *   │ MTEF Header (12 字节)                                       │
 *   │   version(1) + platform(1) + product(1) + prodVer(1)       │
 *   │   + prodSubVer(1) + appKey("DSMT4\0", 6字节) + eqOpts(1)  │
 *   ├─────────────────────────────────────────────────────────────┤
 *   │ FONT_STYLE_DEF 记录 × N（定义字体样式：Text, Variable 等）   │
 *   ├─────────────────────────────────────────────────────────────┤
 *   │ SIZE 记录（定义默认字号，通常 12pt）                          │
 *   ├─────────────────────────────────────────────────────────────┤
 *   │ 顶层 LINE 记录（表达式行，包含所有公式内容）                   │
 *   │   ├─ CHAR / TMPL / EMBELL ... 记录（扁平化的记录流）         │
 *   │   └─ END (0x00)                                             │
 *   ├─────────────────────────────────────────────────────────────┤
 *   │ END (0x00) — 结束整个 MTEF 流                                │
 *   └─────────────────────────────────────────────────────────────┘
 * </pre>
 *
 * <h2>表达式树到 MTEF 扁平记录流的映射规则</h2>
 * <ul>
 *   <li>AST 中的每个字符节点 → 一条 CHAR 记录</li>
 *   <li>AST 中的上标/下标/分数/根号 → 先写 base（在模板外），再写 TMPL 头部，
 *       然后 SUB typesize 记录，接着各 slot（LINE 包裹），最后 END 关闭模板</li>
 *   <li>括号表达式 → TM_PAREN 模板（含 FN_EXPAND 分隔符字符）</li>
 *   <li>大算子（∑∫∏）→ 专用 TMPL 模板，内含积分/求和内容 slot + 上下限 slot + SYM 算子符</li>
 * </ul>
 *
 * <h2>模板写入模式 (Template Writing Pattern)</h2>
 * <pre>
 *   CHAR 'base'          ← base 写在模板 **外部**（父级 LINE 中）
 *   TMPL header           ← 模板头部（模板类型 + 变体 + 选项）
 *     SUB                 ← 1 字节 typesize 记录，标记后续 slot 使用缩小字号
 *     LINE slot_0 END     ← 第一个 slot（用 LINE 记录包裹）
 *     LINE slot_1 END     ← 第二个 slot（如有）
 *   END                   ← 关闭模板
 * </pre>
 *
 * <h2>FULL 记录管理规则</h2>
 * <p>
 *   FULL 记录（record type 0x08）用于在缩小字号的模板（如 TM_SUP / TM_SUB / TM_ROOT）
 *   之后恢复到全尺寸字号。关键规则：
 * </p>
 * <ul>
 *   <li>仅在缩小字号的模板后面 **还有后续兄弟节点** 时才插入 FULL</li>
 *   <li>slot 末尾（LINE END 之前）**不能** 插入 FULL，否则 MathType 解析出错</li>
 *   <li>TM_PAREN 内部自带 FULL 恢复，外部通常不需要额外 FULL</li>
 *   <li>分数模板中，分子 slot 和分母 slot 之间，仅当分子末尾是缩小字号模板时才需要 FULL</li>
 * </ul>
 *
 * <p>
 *   Converts a LaTeX AST into MTEF v5 binary data.
 *   This is the reverse of the Go mtef.go Translate() method:
 *   Go reads MTEF -> produces LaTeX; we read LaTeX AST -> produce MTEF binary.
 * </p>
 */
public class MtefWriter {

    private static final Logger log = LoggerFactory.getLogger(MtefWriter.class);
    private final MathIRConverter mathIRConverter = new MathIRConverter();
    private final MathIRLowerer mathIRLowerer = new MathIRLowerer();
    private final VerticalLayoutCompiler verticalLayoutCompiler = new VerticalLayoutCompiler();
    private final VerticalLayoutNodeFactory verticalLayoutNodeFactory = new VerticalLayoutNodeFactory();
    private final MtefPileRulerWriter pileRulerWriter = new MtefPileRulerWriter();

    /** OLE 模板资源路径 — 包含一个由 MathType 生成的已知正确的 OLE 对象，用于提取 MTEF 前缀 */
    private static final String TEMPLATE_OLE_RESOURCE = "/mathtype-template/oleObject-template.bin";

    /** 目标字体名称 — MathType 中 Text/Variable/Number/Function/Vector 样式默认使用 Times New Roman */
    private static final String TARGET_FONT = "Times New Roman";
    /** 仅恢复模板里的默认字号，不修改公式内部字体。 */
    private static final String TEMPLATE_FULL_SIZE_POINTS =
        System.getProperty("paperword.mtef.full-size-points", "20");
    /**
     * 从模板 OLE 中提取的 MTEF 前缀（header + 字体定义 + SIZE 记录 + 表达式 LINE 的起始部分）。
    * 类加载时从资源文件一次性加载；如果加载失败则为 null，回退到手工构建 header 的模式。
     */
    private static final byte[] TEMPLATE_MTEF_PREFIX = loadTemplateMtefPrefix();
    /**
     * 模板前缀不含末尾 LINE 记录的版本，用于长除法等特殊格式。
     * 长除法参考格式直接以长除法头部开始，无需外层 LINE 包装。
     */
    private static final byte[] TEMPLATE_MTEF_PREFIX_WITHOUT_LINE =
        TEMPLATE_MTEF_PREFIX != null ? trimLineFromPrefix(TEMPLATE_MTEF_PREFIX) : null;
    /**
     * 长除法专用前缀：直接对齐参考 longdivision.docx 在 tmLDIV 之前的 MTEF 结构，
     * 保留 header / font defs / eqn prefs / paren + matrix 壳，仅把矩阵单元格里的除数内容留给运行时写入。
     */
    private static final byte[] LONG_DIVISION_REFERENCE_PREFIX = new byte[] {
        (byte) 0x05, (byte) 0x01, (byte) 0x00, (byte) 0x07, (byte) 0x0A, (byte) 0x44, (byte) 0x53, (byte) 0x4D, (byte) 0x54, (byte) 0x37, (byte) 0x00, (byte) 0x00, (byte) 0x13, (byte) 0x57, (byte) 0x69, (byte) 0x6E,
        (byte) 0x41, (byte) 0x6C, (byte) 0x6C, (byte) 0x42, (byte) 0x61, (byte) 0x73, (byte) 0x69, (byte) 0x63, (byte) 0x43, (byte) 0x6F, (byte) 0x64, (byte) 0x65, (byte) 0x50, (byte) 0x61, (byte) 0x67, (byte) 0x65,
        (byte) 0x73, (byte) 0x00, (byte) 0x11, (byte) 0x05, (byte) 0x54, (byte) 0x69, (byte) 0x6D, (byte) 0x65, (byte) 0x73, (byte) 0x20, (byte) 0x4E, (byte) 0x65, (byte) 0x77, (byte) 0x20, (byte) 0x52, (byte) 0x6F,
        (byte) 0x6D, (byte) 0x61, (byte) 0x6E, (byte) 0x00, (byte) 0x11, (byte) 0x03, (byte) 0x53, (byte) 0x79, (byte) 0x6D, (byte) 0x62, (byte) 0x6F, (byte) 0x6C, (byte) 0x00, (byte) 0x11, (byte) 0x05, (byte) 0x43,
        (byte) 0x6F, (byte) 0x75, (byte) 0x72, (byte) 0x69, (byte) 0x65, (byte) 0x72, (byte) 0x20, (byte) 0x4E, (byte) 0x65, (byte) 0x77, (byte) 0x00, (byte) 0x11, (byte) 0x04, (byte) 0x4D, (byte) 0x54, (byte) 0x20,
        (byte) 0x45, (byte) 0x78, (byte) 0x74, (byte) 0x72, (byte) 0x61, (byte) 0x00, (byte) 0x13, (byte) 0x57, (byte) 0x69, (byte) 0x6E, (byte) 0x41, (byte) 0x6C, (byte) 0x6C, (byte) 0x43, (byte) 0x6F, (byte) 0x64,
        (byte) 0x65, (byte) 0x50, (byte) 0x61, (byte) 0x67, (byte) 0x65, (byte) 0x73, (byte) 0x00, (byte) 0x11, (byte) 0x06, (byte) 0xCB, (byte) 0xCE, (byte) 0xCC, (byte) 0xE5, (byte) 0x00, (byte) 0x12, (byte) 0x00,
        (byte) 0x08, (byte) 0x21, (byte) 0x2F, (byte) 0x27, (byte) 0xF2, (byte) 0x5F, (byte) 0x21, (byte) 0x8F, (byte) 0x21, (byte) 0x2F, (byte) 0x47, (byte) 0x5F, (byte) 0x41, (byte) 0x50, (byte) 0xF2, (byte) 0x1F,
        (byte) 0x1E, (byte) 0x41, (byte) 0x50, (byte) 0xF4, (byte) 0x15, (byte) 0x0F, (byte) 0x41, (byte) 0x00, (byte) 0xF4, (byte) 0x45, (byte) 0xF4, (byte) 0x25, (byte) 0xF4, (byte) 0x8F, (byte) 0x42, (byte) 0x5F,
        (byte) 0x41, (byte) 0x00, (byte) 0xF4, (byte) 0x10, (byte) 0x0F, (byte) 0x43, (byte) 0x5F, (byte) 0x41, (byte) 0x00, (byte) 0xF2, (byte) 0x1F, (byte) 0x20, (byte) 0xA5, (byte) 0xF2, (byte) 0x0A, (byte) 0x25,
        (byte) 0xF4, (byte) 0x8F, (byte) 0x21, (byte) 0xF4, (byte) 0x10, (byte) 0x0F, (byte) 0x41, (byte) 0x00, (byte) 0xF4, (byte) 0x0F, (byte) 0x48, (byte) 0xF4, (byte) 0x17, (byte) 0xF4, (byte) 0x8F, (byte) 0x41,
        (byte) 0x00, (byte) 0xF2, (byte) 0x1A, (byte) 0x5F, (byte) 0x44, (byte) 0x5F, (byte) 0x45, (byte) 0xF4, (byte) 0x5F, (byte) 0x45, (byte) 0xF4, (byte) 0x5F, (byte) 0x41, (byte) 0x0F, (byte) 0x0C, (byte) 0x01,
        (byte) 0x00, (byte) 0x01, (byte) 0x00, (byte) 0x01, (byte) 0x02, (byte) 0x02, (byte) 0x02, (byte) 0x02, (byte) 0x00, (byte) 0x02, (byte) 0x00, (byte) 0x01, (byte) 0x01, (byte) 0x01, (byte) 0x00,
        (byte) 0x03, (byte) 0x00, (byte) 0x01, (byte) 0x00, (byte) 0x04, (byte) 0x00, (byte) 0x05, (byte) 0x00, (byte) 0x0A, (byte) 0x04, (byte) 0x00, (byte) 0x01, (byte) 0x01, (byte) 0x01, (byte) 0x00,
    };

    /**
     * 求和类大算子 — 使用 TM_SUM 模板，上下限以极限（limit）形式显示在算子正上方/正下方。
     * 包括：求和 ∑、求积 ∏、余积 ∐、大并集 ⋃、大交集 ⋂、大析取 ⋁、大合取 ⋀
     */
    private static final Set<String> BIG_OP_SUM_LIKE = Set.of(
        "\\sum", "\\prod", "\\coprod", "\\bigcup", "\\bigcap", "\\bigvee", "\\bigwedge"
    );

    /** 极限类算子 — 使用 TM_LIM 模板，按 summation-style 显示上下限。 */
    private static final Set<String> BIG_OP_LIMIT_LIKE = Set.of(
        "\\lim"
    );

    /**
     * 积分类大算子 — 使用 TM_INTEGRAL 模板，上下限以上标/下标形式显示在算子右侧。
     * 包括：单重积分 ∫、二重积分 ∬、三重积分 ∭、曲线积分 ∮
     */
    private static final Set<String> BIG_OP_INT_LIKE = Set.of(
        "\\int", "\\iint", "\\iiint", "\\oint"
    );

    /**
     * 将 LaTeX AST 转换为 MTEF v5 二进制字节数组（公开入口方法）。
     *
     * <p>转换策略：</p>
     * <ul>
     *   <li><b>首选模式</b>：使用从 MathType 模板中提取的前缀（TEMPLATE_MTEF_PREFIX），
     *       这样可以复用 MathType 原生生成的 header、字体定义和 SIZE 记录，确保兼容性最佳。</li>
     *   <li><b>回退模式</b>：如果模板前缀不可用，手工构建 header + 字体定义 + SIZE 记录 +
     *       顶层 LINE 记录。</li>
     * </ul>
     *
     * <p>Convert a LaTeX AST (from LaTeXParser) to MTEF v5 binary.
     * Returns the complete MTEF byte array including header.</p>
     */
    public byte[] write(LaTeXNode root) {
        return write(mathIRConverter.convert(root));
    }

    /**
     * Phase 3 入口：先面向 MathML-aligned IR，再落回现有的稳定 AST→MTEF 发射逻辑。
     */
    public byte[] write(MathIRNode root) {
        LaTeXNode normalizedAst = mathIRLowerer.lower(root);
        return writeNormalizedAst(normalizedAst);
    }

    private byte[] writeNormalizedAst(LaTeXNode root) {
        try {
            // 使用模板前缀模式（从已知正确的 MathType OLE 中提取前缀）
            if (TEMPLATE_MTEF_PREFIX != null) {
                if (isCompositeLongDivisionRoot(root)) {
                    return writeCompositeLongDivisionByTemplate(root);
                }
                if (isLongDivisionRoot(root)) {
                    return writeLongDivisionByTemplate(root);
                }
                return writeByTemplatePrefix(root);
            }

            return writeFullStream(root);
        } catch (IOException e) {
            throw new RuntimeException("Failed to write MTEF binary", e);
        }
    }

    private byte[] writeFullStream(LaTeXNode root) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream(256);
        writeHeader(out);          // 写入 MTEF v5 文件头（12 字节）
        writeFontStyleDefs(out);   // 写入字体样式定义记录
        writeSizeRecord(out);      // 写入默认字号 SIZE 记录

        // 写入顶层表达式 LINE 记录（所有公式内容都包含在这个 LINE 中）
        out.write(MtefRecord.LINE);
        out.write(0x00); // options: 0x00 表示无偏移(nudge)、非空行(not null)

        // 递归写入 AST 根节点的所有内容
        writeNode(out, root);

        // 关闭顶层 LINE 记录
        out.write(MtefRecord.END);
        // 再追加一个 END，显式结束顶层对象列表，和模板兼容路径保持一致。
        out.write(MtefRecord.END);
        return out.toByteArray();
    }

    /**
     * 检查根节点是否只包含一个长除法节点（需要特殊格式处理）。
     */
    private boolean isLongDivisionRoot(LaTeXNode root) {
        if (root.getType() == LaTeXNode.Type.LONG_DIVISION) {
            return true;
        }
        return root.getType() == LaTeXNode.Type.ROOT
            && root.getChildren().size() == 1
            && root.getChildren().get(0).getType() == LaTeXNode.Type.LONG_DIVISION;
    }

    private boolean isCompositeLongDivisionRoot(LaTeXNode root) {
        return root != null
            && root.getType() == LaTeXNode.Type.ROOT
            && root.getChildren().size() == 2
            && root.getChildren().get(0).getType() == LaTeXNode.Type.LONG_DIVISION
            && root.getChildren().get(1).getType() == LaTeXNode.Type.ARRAY;
    }

    /**
     * 长除法专用写入方法：复用裁剪后的模板前缀，保留参考对象的 header / font / eqn prefs，
     * 并在 tmLDIV 前显式重建 reference 中存在的 paren + matrix + divisor 壳。
     */
    private byte[] writeLongDivisionByTemplate(LaTeXNode root) throws IOException {
        LaTeXNode longDivision = root;
        if (root.getType() == LaTeXNode.Type.ROOT && root.getChildren().size() == 1
            && root.getChildren().get(0).getType() == LaTeXNode.Type.LONG_DIVISION) {
            longDivision = root.getChildren().get(0);
        }

        VerticalLayoutSpec layoutSpec = verticalLayoutCompiler.compileExplicitLongDivision(longDivision);
        String originalQuotient = flattenNodeText(childAt(longDivision, 1));
        boolean shouldUseComputedLayout = layoutSpec != null
            && (!layoutSpec.longDivisionHeader().quotient().equals(originalQuotient)
            || !layoutSpec.longDivisionSteps().isEmpty());
        if (shouldUseComputedLayout) {
            return writeByTemplatePrefix(verticalLayoutNodeFactory.buildComputedLongDivisionArray(longDivision, layoutSpec));
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream(512);
        out.write(LONG_DIVISION_REFERENCE_PREFIX);
        longDivision.setMetadata("skipLongDivisionLeadingDivisor", "true");
        LaTeXNode divisor = childAt(longDivision, 0);
        if (divisor != null) {
            writeNode(out, divisor);
        }
        writeNode(out, longDivision);
        out.write(MtefRecord.END);
        return out.toByteArray();
    }

    private byte[] writeCompositeLongDivisionByTemplate(LaTeXNode root) throws IOException {
        LaTeXNode longDivision = root.getChildren().get(0);
        LaTeXNode stepArray = root.getChildren().get(1);
        VerticalLayoutSpec headerSpec = verticalLayoutCompiler.compileExplicitLongDivision(longDivision);
        List<VerticalLayoutNodeFactory.RawLongDivisionLine> stepLines = extractCompositeLongDivisionLines(stepArray);
        LaTeXNode compositeArray = verticalLayoutNodeFactory.buildCompositeLongDivisionArray(longDivision, headerSpec, stepLines);
        return writeByTemplatePrefix(compositeArray);
    }

    private List<VerticalLayoutNodeFactory.RawLongDivisionLine> extractCompositeLongDivisionLines(LaTeXNode stepArray) {
        List<VerticalLayoutNodeFactory.RawLongDivisionLine> lines = new ArrayList<>();
        if (stepArray == null || stepArray.getType() != LaTeXNode.Type.ARRAY) {
            return lines;
        }
        for (LaTeXNode rowNode : stepArray.getChildren()) {
            CompositeLongDivisionLine line = flattenCompositeLongDivisionLine(rowNode);
            if (line.text().isBlank()) {
                continue;
            }
            lines.add(new VerticalLayoutNodeFactory.RawLongDivisionLine(line.text(), line.underlined()));
        }
        return lines;
    }

    private CompositeLongDivisionLine flattenCompositeLongDivisionLine(LaTeXNode node) {
        StringBuilder builder = new StringBuilder();
        boolean underlined = appendCompositeLongDivisionText(node, builder);
        return new CompositeLongDivisionLine(builder.toString(), underlined);
    }

    private boolean appendCompositeLongDivisionText(LaTeXNode node, StringBuilder builder) {
        if (node == null) {
            return false;
        }
        boolean underlined = false;
        switch (node.getType()) {
            case CHAR -> builder.append(node.getValue() == null ? "" : node.getValue());
            case COMMAND -> {
                if ("\\underline".equals(node.getValue())) {
                    underlined = true;
                    for (LaTeXNode child : node.getChildren()) {
                        underlined = appendCompositeLongDivisionText(child, builder) || underlined;
                    }
                }
            }
            case ROOT, GROUP, TEXT, ARRAY, ROW, CELL, LONG_DIVISION, FRACTION, SQRT, SUPERSCRIPT, SUBSCRIPT -> {
                for (LaTeXNode child : node.getChildren()) {
                    underlined = appendCompositeLongDivisionText(child, builder) || underlined;
                }
            }
            default -> {
            }
        }
        return underlined;
    }

    private void writeExplicitLongDivisionByReferenceShape(ByteArrayOutputStream out,
                                                           LaTeXNode longDivision,
                                                           VerticalLayoutSpec layoutSpec) throws IOException {
        LaTeXNode structured = verticalLayoutNodeFactory.buildStructuredLongDivisionNode(longDivision, layoutSpec);
        if (structured.getChildren().isEmpty()) {
            writeNode(out, longDivision);
            return;
        }
        // 保持步骤区位于 tmLDIV 外层，避免把复杂布局对象直接塞进 tmLDIV 槽位后导致 MathType 无法编辑。
        writeNode(out, structured.getChildren().get(0));
        out.write(MtefRecord.END);

        LaTeXNode pendingUnderline = null;
        int pendingUnderlineExtraEnds = 0;
        for (int index = 1; index < structured.getChildren().size(); index++) {
            LaTeXNode child = structured.getChildren().get(index);
            boolean splitStep = "true".equals(child.getMetadata(VerticalLayoutNodeFactory.LONG_DIVISION_STEP_SPLIT));
            boolean firstSplitStep = "true".equals(child.getMetadata(VerticalLayoutNodeFactory.LONG_DIVISION_FIRST_STEP_SPLIT));
            if (index == 1) {
                writeNode(out, child);
                int endCount = firstSplitStep ? 2 : 4;
                for (int i = 0; i < endCount; i++) {
                    out.write(MtefRecord.END);
                }
                continue;
            }
            if (!splitStep) {
                boolean lastChild = index == structured.getChildren().size() - 1;
                if (pendingUnderline != null && (lastChild || index >= 5)) {
                    for (int i = 0; i < pendingUnderlineExtraEnds; i++) {
                        out.write(MtefRecord.END);
                    }
                    writeNode(out, pendingUnderline);
                    if (index == 5) {
                        for (int i = 0; i < 5; i++) {
                            out.write(MtefRecord.END);
                        }
                    }
                    pendingUnderline = null;
                    pendingUnderlineExtraEnds = 0;
                }
                if (index == 3) {
                    for (int i = 0; i < 7; i++) {
                        out.write(MtefRecord.END);
                    }
                }
                writeSlot(out, child);
                if (pendingUnderline != null) {
                    for (int i = 0; i < pendingUnderlineExtraEnds; i++) {
                        out.write(MtefRecord.END);
                    }
                    writeNode(out, pendingUnderline);
                    pendingUnderline = null;
                    pendingUnderlineExtraEnds = 0;
                }
            } else {
                if (!firstSplitStep && child.getChildren().size() >= 2) {
                    if (index == 4) {
                        for (int i = 0; i < 6; i++) {
                            out.write(MtefRecord.END);
                        }
                    }
                    writeSlot(out, child.getChildren().get(0));
                    int endCount = index == 4 ? 0 : 2;
                    for (int i = 0; i < endCount; i++) {
                        out.write(MtefRecord.END);
                    }
                    if (index < structured.getChildren().size() - 1) {
                        pendingUnderline = child.getChildren().get(1);
                        pendingUnderlineExtraEnds = index == 4 ? 2 : (index == 6 ? 2 : 0);
                    } else {
                        for (int i = 0; i < 2; i++) {
                            out.write(MtefRecord.END);
                        }
                        writeNode(out, child.getChildren().get(1));
                    }
                } else {
                    writeNode(out, child);
                }
            }
        }
        if (pendingUnderline != null) {
            for (int i = 0; i < pendingUnderlineExtraEnds; i++) {
                out.write(MtefRecord.END);
            }
            writeNode(out, pendingUnderline);
        }
    }

    /**
     * 首选模式：复用从 MathType 模板提取的前缀，仅追加生成的表达式记录。
     *
     * <p>模板前缀已包含：MTEF header + 字体定义 + SIZE + 顶层 LINE 开头。
     * 本方法只需在前缀后面追加公式内容记录，然后用两个 END 结束：
     * 第一个 END 关闭顶层 LINE，第二个 END 结束整个 MTEF 流。</p>
     *
     * <p>Preferred mode: reuse the proven-good MathType template prelude and only
     * append generated expression records, then terminate with END END.</p>
     */
    private byte[] writeByTemplatePrefix(LaTeXNode root) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream(512);
        out.write(TEMPLATE_MTEF_PREFIX);   // 写入模板前缀（header + 字体 + SIZE + LINE 开头）
        writeNode(out, root);              // 递归写入 AST 内容
        out.write(MtefRecord.END);         // 第一个 END：关闭顶层 LINE 记录
        out.write(MtefRecord.END);         // 第二个 END：结束 MTEF 流
        return out.toByteArray();
    }

    /**
     * 类加载时从模板 OLE 资源中提取 MTEF 前缀。
     *
     * <p>处理流程：</p>
     * <ol>
     *   <li>读取 OLE 模板文件（由 MathType 生成的含公式 "A" 的 oleObject）</li>
     *   <li>通过 POI 解析 OLE 结构，提取 "Equation Native" 流</li>
     *   <li>跳过 Equation Native 的 28 字节 OLE 头部，得到纯 MTEF 数据</li>
     *   <li>定位模板公式的表达式部分（CHAR VAR 'A' + END + END），
     *       截取其之前的所有字节作为前缀</li>
     *   <li>前缀包含：MTEF header + 所有 FONT_STYLE_DEF + SIZE + 顶层 LINE 的起始</li>
     * </ol>
     *
     * @return MTEF 前缀字节数组；加载失败时返回 null（触发回退模式）
     */
    private static byte[] loadTemplateMtefPrefix() {
        try (InputStream in = MtefWriter.class.getResourceAsStream(TEMPLATE_OLE_RESOURCE)) {
            if (in == null) {
                return null;
            }
            byte[] oleBytes = in.readAllBytes();

            // 使用 Apache POI 解析 OLE 复合文档结构
            org.apache.poi.poifs.filesystem.POIFSFileSystem fs =
                new org.apache.poi.poifs.filesystem.POIFSFileSystem(new java.io.ByteArrayInputStream(oleBytes));
            try (fs) {
                var root = fs.getRoot();

                // 查找 "Equation Native" 流 — MathType 公式的原生二进制数据存储在这里
                if (!root.hasEntry("Equation Native")) {
                    return null;
                }
                var entry = root.getEntry("Equation Native");
                if (!(entry instanceof org.apache.poi.poifs.filesystem.DocumentEntry docEntry)) {
                    return null;
                }
                byte[] eqNative;
                try (org.apache.poi.poifs.filesystem.DocumentInputStream dis =
                         new org.apache.poi.poifs.filesystem.DocumentInputStream(docEntry)) {
                    eqNative = dis.readAllBytes();
                }

                // Equation Native 格式：前 28 字节为 OLE header，后面才是 MTEF 数据
                if (eqNative.length < 28 + 16) {
                    return null;
                }
                // 读取 OLE header 中的 header size 字段（小端序 16 位整数）
                int hdrSize = ((eqNative[0] & 0xFF) | ((eqNative[1] & 0xFF) << 8));
                if (hdrSize < 12 || eqNative.length <= hdrSize + 8) {
                    return null;
                }
                // 跳过 OLE header，提取纯 MTEF 数据
                byte[] mtef = java.util.Arrays.copyOfRange(eqNative, hdrSize, eqNative.length);

                // 模板公式 "A" 在 MTEF 中的表达式尾部为：
                //   CHAR(0x02) options(0x00) typeface(0x83=VAR) mtcode_lo(0x41='A') mtcode_hi(0x00) + END(0x00) + END(0x00)
                // 我们截取该表达式之前的所有字节作为前缀。
                // 适用于 DSMT6（前缀以 0F 01 结束）和 DSMT7（前缀以 0A 01 00 结束）两种格式。
                byte[] exprTail = new byte[]{0x02, 0x00, (byte) 0x83, 0x41, 0x00, 0x00, 0x00};
                int idx = indexOf(mtef, exprTail);
                if (idx < 0) {
                    return null;
                }
                int prefixLen = idx; // 表达式之前的所有内容即为前缀
                byte[] prefix = java.util.Arrays.copyOfRange(mtef, 0, prefixLen);
                return patchEqnPrefsFullSize(patchFontsToTimesNewRoman(prefix));
            }
        } catch (Exception e) {
            log.warn("Failed to load template MTEF prefix; fallback to legacy writer", e);
            return null;
        }
    }

    /**
     * 从模板前缀中移除末尾的 LINE 记录 (01 00)。
     * 用于长除法等需要直接以公式头部开始的特殊格式。
     */
    private static byte[] trimLineFromPrefix(byte[] prefix) {
        if (prefix == null || prefix.length < 2) {
            return prefix;
        }
        // 检查末尾是否为 LINE 记录 (01 00)
        if (prefix[prefix.length - 2] == MtefRecord.LINE && prefix[prefix.length - 1] == 0x00) {
            return java.util.Arrays.copyOfRange(prefix, 0, prefix.length - 2);
        }
        return prefix;
    }


    /**
     * 将 MTEF 前缀中非 Symbol 字体替换为 Times New Roman。
     *
     * <p>保留 "Symbol" 和 "MT Extra" 字体不变，因为 MathType 依赖这两个字体
     * 来正确渲染希腊字母和特殊数学符号（通过与 MathType 7 输出对比确认）。</p>
     *
     * <p>Patch non-Symbol font names in MTEF prefix to Times New Roman.
     * Keep "Symbol" and "MT Extra" as-is because MathType expects them for
     * SYM/Greek typeface characters (confirmed by comparing with MathType 7 output).</p>
     */
    private static byte[] patchFontsToTimesNewRoman(byte[] prefix) {
        // 仅替换应为 TNR 但实际不是的字体；
        // Symbol 和 MT Extra 必须保持不变以确保字形正确渲染。
        return prefix;
    }

    /**
     * 只恢复模板中的默认 full size，避免 Word/MathType 打开后重新排版时缩成最早的错误版本。
     * 这个补丁不改字体定义，也不改整体 nudge。
     */
    private static byte[] patchEqnPrefsFullSize(byte[] prefix) {
        if (prefix == null || prefix.length < 4) {
            return prefix;
        }
        int recordIndex = indexOf(prefix, new byte[] {
            (byte) MtefRecord.EQN_PREFS, 0x00, 0x08
        });
        if (recordIndex < 0) {
            return prefix;
        }

        int nibbleStreamStart = recordIndex + 3;
        ParseResult parseResult = parseDimensionArray(prefix, nibbleStreamStart, 8);
        if (parseResult == null || parseResult.dimensions.isEmpty()) {
            return prefix;
        }

        parseResult.dimensions.set(0, buildPointsDimension(TEMPLATE_FULL_SIZE_POINTS));
        byte[] packed = packDimensions(parseResult.dimensions);

        ByteArrayOutputStream out = new ByteArrayOutputStream(prefix.length + 8);
        out.write(prefix, 0, nibbleStreamStart);
        out.writeBytes(packed);
        out.write(prefix, parseResult.endByteOffset, prefix.length - parseResult.endByteOffset);
        return out.toByteArray();
    }

    private static ParseResult parseDimensionArray(byte[] data, int startByteOffset, int count) {
        List<int[]> dimensions = new ArrayList<>();
        List<Integer> current = new ArrayList<>();
        int nibbleIndex = 0;
        int byteOffset = startByteOffset;
        while (byteOffset < data.length && dimensions.size() < count) {
            int value = data[byteOffset] & 0xFF;
            int[] nibbles = new int[] {(value >>> 4) & 0x0F, value & 0x0F};
            for (int nibble : nibbles) {
                current.add(nibble);
                nibbleIndex++;
                if (nibble == 0x0F) {
                    dimensions.add(current.stream().mapToInt(Integer::intValue).toArray());
                    current = new ArrayList<>();
                    if (dimensions.size() == count) {
                        int consumedBytes = (nibbleIndex + (nibbleIndex % 2)) / 2;
                        return new ParseResult(dimensions, startByteOffset + consumedBytes);
                    }
                }
            }
            byteOffset++;
        }
        return null;
    }

    private static byte[] packDimensions(List<int[]> dimensions) {
        List<Integer> nibbles = new ArrayList<>();
        for (int[] dimension : dimensions) {
            for (int nibble : dimension) {
                nibbles.add(nibble & 0x0F);
            }
        }
        if ((nibbles.size() & 1) == 1) {
            nibbles.add(0);
        }
        byte[] packed = new byte[nibbles.size() / 2];
        for (int i = 0; i < nibbles.size(); i += 2) {
            packed[i / 2] = (byte) ((nibbles.get(i) << 4) | nibbles.get(i + 1));
        }
        return packed;
    }

    private static int[] buildPointsDimension(String points) {
        String normalized = points == null || points.isBlank() ? "18" : points.trim();
        int[] nibbles = new int[normalized.length() + 2];
        nibbles[0] = 0x02;
        for (int i = 0; i < normalized.length(); i++) {
            char ch = normalized.charAt(i);
            nibbles[i + 1] = ch == '.' ? 0x0A : Character.digit(ch, 10);
        }
        nibbles[nibbles.length - 1] = 0x0F;
        return nibbles;
    }

    private record ParseResult(List<int[]> dimensions, int endByteOffset) {}

    /**
     * 在字节数组中查找并替换所有匹配的子序列（通用字节替换工具方法）。
     *
     * @param data 原始数据
     * @param from 要查找的字节序列
     * @param to   替换为的字节序列
     * @return 替换后的新字节数组
     */
    private static byte[] replaceBytes(byte[] data, byte[] from, byte[] to) {
        ByteArrayOutputStream out = new ByteArrayOutputStream(data.length + 64);
        int i = 0;
        while (i < data.length) {
            if (i + from.length <= data.length) {
                boolean match = true;
                for (int j = 0; j < from.length; j++) {
                    if (data[i + j] != from[j]) {
                        match = false;
                        break;
                    }
                }
                if (match) {
                    out.write(to, 0, to.length);
                    i += from.length;
                    continue;
                }
            }
            out.write(data[i]);
            i++;
        }
        return out.toByteArray();
    }

    /**
     * 在字节数组中查找子序列首次出现的位置（类似 String.indexOf 的字节版本）。
     *
     * @return 匹配位置的起始索引；未找到返回 -1
     */
    private static int indexOf(byte[] data, byte[] pattern) {
        outer:
        for (int i = 0; i <= data.length - pattern.length; i++) {
            for (int j = 0; j < pattern.length; j++) {
                if (data[i + j] != pattern[j]) {
                    continue outer;
                }
            }
            return i;
        }
        return -1;
    }

    /**
     * 写入 MTEF v5 文件头（仅在回退模式下使用）。
     *
     * <p>MTEF 文件头格式（共 12 字节）：</p>
     * <pre>
     *   字节 0: version        — MTEF 版本号（5）
     *   字节 1: platform       — 平台标识（1 = Windows, 0 = Mac）
     *   字节 2: product        — 产品标识（0 = MathType）
     *   字节 3: prodVersion    — 产品主版本号（7）
     *   字节 4: prodSubVersion — 产品子版本号（0）
     *   字节 5-10: appKey      — 应用标识 "DSMT4" + 0x00（null 终止符）
     *   字节 11: eqOptions     — 公式选项（0x01）
     * </pre>
     *
     * <p>Write MTEF v5 header.
     * Format: version(1) + platform(1) + product(1) + prodVer(1) + prodSubVer(1)
     *       + applicationKey(null-terminated) + equationOptions(1)
     * Total: 5 + strlen("DSMT4")+1 + 1 = 12 bytes</p>
     */
    private void writeHeader(ByteArrayOutputStream out) throws IOException {
        out.write(MtefRecord.MTEF_VERSION);           // version = 5（MTEF v5 格式）
        out.write(MtefRecord.MTEF_PLATFORM_WIN);      // platform = Windows (1)
        out.write(MtefRecord.MTEF_PRODUCT_MATHTYPE);   // product = MathType (0)
        out.write(MtefRecord.MTEF_PRODUCT_VERSION);    // product version = 7
        out.write(MtefRecord.MTEF_PRODUCT_SUBVERSION); // product subversion = 0
        // 应用标识键（Application key）：null 终止的 "DSMT4" 字符串
        out.write(MtefRecord.MTEF_APPLICATION_KEY.getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        out.write(0x00); // null 终止符
        // 公式选项：与已知正确的 MathType 模板输出对齐
        out.write(0x01);
    }

    /**
     * 写入 MathType 标准字体样式定义记录（仅在回退模式下使用）。
     *
     * <p>MathType 使用不同的字体样式（typeface）来区分数学公式中的不同元素类型：</p>
     * <ul>
     *   <li>FN_TEXT / FN_FUNCTION / FN_VARIABLE / FN_VECTOR / FN_NUMBER → Times New Roman</li>
     *   <li>FN_LC_GREEK / FN_UC_GREEK / FN_SYMBOL → Symbol 字体（希腊字母和数学符号）</li>
     * </ul>
     *
     * <p>每条 FONT_STYLE_DEF 记录格式：record_type(1) + font_index(1) + font_name(ASCII) + 0x00</p>
     */
    private void writeFontStyleDefs(ByteArrayOutputStream out) throws IOException {
        // 文本/变量/数字/函数名/向量 → Times New Roman
        writeFontDef(out, MtefRecord.FN_TEXT, TARGET_FONT);
        writeFontDef(out, MtefRecord.FN_FUNCTION, TARGET_FONT);
        writeFontDef(out, MtefRecord.FN_VARIABLE, TARGET_FONT);
        // 希腊字母/符号 → Symbol 字体（MathType 依赖此字体进行正确的字形映射）
        writeFontDef(out, MtefRecord.FN_LC_GREEK, "Symbol");
        writeFontDef(out, MtefRecord.FN_UC_GREEK, "Symbol");
        writeFontDef(out, MtefRecord.FN_SYMBOL, "Symbol");
        writeFontDef(out, MtefRecord.FN_VECTOR, TARGET_FONT);
        writeFontDef(out, MtefRecord.FN_NUMBER, TARGET_FONT);
        writeFontDef(out, MtefRecord.FN_MTEXTRA, "MT Extra");
    }

    /**
     * 写入单条字体样式定义记录。
     *
     * @param fontDefIndex 字体样式索引（如 FN_TEXT, FN_VARIABLE 等）
     * @param fontName     字体名称（ASCII 编码，null 终止）
     */
    private void writeFontDef(ByteArrayOutputStream out, int fontDefIndex, String fontName) throws IOException {
        out.write(MtefRecord.FONT_STYLE_DEF);
        out.write(fontDefIndex);
        out.write(fontName.getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        out.write(0x00); // null 终止符
    }

    /**
     * 写入 SIZE 记录：设置默认全尺寸字号（12pt）。
     *
     * <p>SIZE 记录格式：record_type(1) + lsize(1) + point_size(1)</p>
     * <p>lsize = FULL 表示使用全尺寸（而非缩小的上标/下标尺寸）。</p>
     */
    private void writeSizeRecord(ByteArrayOutputStream out) {
        out.write(MtefRecord.SIZE);
        out.write(MtefRecord.FULL); // lsize = FULL（使用全尺寸）
        out.write(0x0C); // 12pt — MathType 默认字号
    }

    /**
     * AST 节点分发器：根据节点类型递归调用对应的写入方法。
     *
     * <p>这是 AST → MTEF 转换的核心递归入口。每种节点类型对应不同的 MTEF 记录生成策略：</p>
     * <ul>
     *   <li>ROOT / GROUP → 透传，递归处理子节点列表</li>
     *   <li>CHAR → 生成一条 CHAR 记录</li>
     *   <li>COMMAND → 根据命令类型生成 CHAR 记录或 TMPL 模板</li>
     *   <li>FRACTION → 生成 TM_FRACT 分数模板</li>
     *   <li>SQRT → 生成 TM_ROOT 根号模板</li>
     *   <li>SUPERSCRIPT → 生成 TM_SUP 上标模板（base 写在模板外部）</li>
     *   <li>SUBSCRIPT → 生成 TM_SUB 下标模板（base 写在模板外部）</li>
     *   <li>TEXT → 生成 FN_TEXT 字体的 CHAR 记录序列</li>
     * </ul>
     */
    private void writeNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        switch (node.getType()) {
            case ROOT, GROUP, ROW, CELL -> writeContainerNode(out, node);
            case CHAR -> writeCharNode(out, node);
            case COMMAND -> writeCommandNode(out, node);
            case FRACTION -> writeFractionNode(out, node);
            case SQRT -> writeSqrtNode(out, node);
            case SUPERSCRIPT -> writeSuperscriptNode(out, node);
            case SUBSCRIPT -> writeSubscriptNode(out, node);
            case TEXT -> writeTextNode(out, node);
            case ARRAY -> writeArrayNode(out, node);
            case LONG_DIVISION -> writeLongDivisionNode(out, node);
        }
    }

    private void writeContainerNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        if ("true".equals(node.getMetadata(VerticalLayoutNodeFactory.RAW_LINE_CONTAINER))) {
            writeRawLineNode(out, node);
            return;
        }
        if ("true".equals(node.getMetadata(VerticalLayoutNodeFactory.RAW_PILE_CONTAINER))) {
            writeRawPileNode(out, node);
            return;
        }
        writeChildren(out, node);
    }

    /**
     * 写入节点的所有子节点（ROOT 和 GROUP 节点的透传处理）。
     * 委托给 writeContentNodes 进行带 FULL 插入、括号检测和大算子处理的迭代。
     */
    private void writeChildren(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        writeContentNodes(out, node.getChildren());
    }

    /**
     * 【核心方法】遍历子节点列表，生成 MTEF 记录流，并处理三大特殊逻辑。
     *
     * <p>这是 writeChildren 和 writeBigOpComplete 共用的核心迭代逻辑，
     * 负责将一组兄弟节点按序转换为 MTEF 记录，同时处理：</p>
     *
     * <h3>1. FULL 记录插入</h3>
     * <p>当一个生成模板的节点（上标/下标/根号/分数）后面还有更多兄弟节点时，
     * 在模板之后插入 FULL 记录以恢复全尺寸字号。
     * 关键约束：slot 末尾（LINE END 之前）绝对不能插入 FULL。</p>
     *
     * <h3>2. 括号表达式检测</h3>
     * <p>检测 '(' ... ')^{exp}' 或 '(' ... ')_{sub}' 模式。
     * MathType 要求括号内容使用 TM_PAREN fence 模板，才能正确附加上标/下标。
     * 独立的 (...) 也会被转换为 TM_PAREN 模板。</p>
     *
     * <h3>3. 大算子检测</h3>
     * <p>检测求和/积分/求积等大算子模式（如 \sum_{i=0}^{n}），
     * 一旦发现大算子，将其后的所有剩余兄弟节点作为积分/求和内容，
     * 然后提前返回（因为大算子模板会消费掉所有剩余内容）。</p>
     *
     * <p>Write a list of content nodes with proper FULL insertion, paren detection, and big-op handling.
     * This is the shared core logic used by writeChildren and writeBigOpComplete.</p>
     */
    private void writeContentNodes(ByteArrayOutputStream out, java.util.List<LaTeXNode> nodes) throws IOException {
        for (int i = 0; i < nodes.size(); i++) {
            LaTeXNode child = nodes.get(i);

            // ========== 大算子检测 ==========
            // 检测当前节点是否为大算子模式（如 \sum_{lower}^{upper}）。
            // 大算子的 AST 结构：SUPERSCRIPT(SUBSCRIPT(COMMAND(\sum), lower), upper)
            // 如果匹配，将 i+1 之后的所有兄弟节点作为积分/求和内容，写入后直接 return。
            BigOpInfo bigOp = extractBigOp(child);
            if (bigOp != null) {
                java.util.List<LaTeXNode> remaining = nodes.subList(i + 1, nodes.size());
                writeBigOpComplete(out, bigOp, remaining);
                return; // 大算子消费了所有剩余兄弟节点，直接返回
            }

            // ========== 括号表达式检测 ==========
            // 检测 '(' 或 '[' 开头的括号对，匹配对应的 ')' 或 ']'。
            // MathType 要求括号内容使用 TM_PAREN / TM_BRACK fence 模板。
            if (child.getType() == LaTeXNode.Type.CHAR && child.getValue() != null
                    && child.getValue().length() == 1) {
                char openCh = child.getValue().charAt(0);
                char closeCh = matchingClose(openCh);
                if (closeCh != 0) {
                    // 在后续兄弟中查找匹配的闭括号（支持嵌套）
                    int closeIdx = findMatchingClose(nodes, i + 1, closeCh);
                    if (closeIdx > i) {
                        LaTeXNode closeNode = nodes.get(closeIdx);
                        java.util.List<LaTeXNode> parenContent = nodes.subList(i + 1, closeIdx);

                        if (closeNode.getType() == LaTeXNode.Type.SUPERSCRIPT
                                || closeNode.getType() == LaTeXNode.Type.SUBSCRIPT) {
                            // 模式：(...)^{exp} 或 (...)_{sub}
                            // 先写括号 fence 模板，再写上标/下标附件
                            writeParenFence(out, openCh, closeCh, parenContent);
                            writeSupSubAttachment(out, closeNode);
                            // 括号+上下标组合也是模板，后面有更多内容时需要 FULL 恢复字号
                            if (closeIdx < nodes.size() - 1) {
                                out.write(MtefRecord.FULL);
                            }
                            i = closeIdx; // 跳过已处理的括号内容和闭括号节点
                            continue;
                        } else if (closeNode.getType() == LaTeXNode.Type.CHAR) {
                            // 模式：(...) 独立括号 — 同样使用 TM_PAREN 模板
                            writeParenFence(out, openCh, closeCh, parenContent);
                            i = closeIdx; // 跳过已处理的括号内容和闭括号节点
                            continue;
                        }
                    }
                }
            }

            // ========== 普通节点处理 ==========
            writeNode(out, child);

            // ========== FULL 记录插入 ==========
            // 当当前节点生成了模板（上标/下标/根号/分数）且后面还有更多兄弟节点时，
            // 插入 FULL 记录恢复全尺寸字号。
            // 关键：在 slot 末尾（即 i == nodes.size()-1 时）不能插入 FULL，
            // 否则会在 LINE END 之前产生多余的 FULL 记录导致 MathType 解析错误。
            if (i < nodes.size() - 1 && generatesTemplate(child)) {
                out.write(MtefRecord.FULL);
            }
        }
    }

    /**
     * 判断一个 AST 节点是否会生成 MTEF 模板（TMPL 记录）。
     *
     * <p>生成模板的节点类型包括：上标（TM_SUP）、下标（TM_SUB）、根号（TM_ROOT）、分数（TM_FRACT）。
     * 这些模板内部使用 SUB typesize 记录缩小字号，因此在模板后需要 FULL 恢复。</p>
     */
    private boolean generatesTemplate(LaTeXNode node) {
        return switch (node.getType()) {
            case SUPERSCRIPT, SUBSCRIPT, SQRT, FRACTION, LONG_DIVISION -> true;
            default -> false;
        };
    }

    /**
     * 根据开括号字符返回对应的闭括号字符。
     *
     * @return 匹配的闭括号；不支持的字符返回 0
     */
    private static char matchingClose(char open) {
        return switch (open) {
            case '(' -> ')';
            case '[' -> ']';
            default -> 0;
        };
    }

    /**
     * 在子节点列表中查找与开括号匹配的闭括号位置（支持嵌套）。
     *
     * <p>查找逻辑：</p>
     * <ul>
     *   <li>维护嵌套深度计数器 depth（初始为 1，因为已经遇到了一个开括号）</li>
     *   <li>遇到开括号 → depth++；遇到闭括号 → depth--；depth==0 时即为匹配位置</li>
     *   <li>闭括号可能以两种形式出现：
     *       (a) 独立的 CHAR 节点，如 ')'
     *       (b) SUPERSCRIPT/SUBSCRIPT 节点的 base 子节点，如 ')^{2}' 中的 ')'</li>
     * </ul>
     *
     * <p>Find the index of the matching close delimiter in the children list.
     * Handles nesting and checks for close char either as a standalone CHAR
     * or as the base of a SUPERSCRIPT/SUBSCRIPT node.</p>
     *
     * @return 匹配闭括号的索引；未找到返回 -1
     */
    private int findMatchingClose(java.util.List<LaTeXNode> children, int start, char closeCh) {
        char openCh = closeCh == ')' ? '(' : '[';
        int depth = 1; // 嵌套深度：已有一个开括号
        for (int j = start; j < children.size(); j++) {
            LaTeXNode c = children.get(j);
            // 情况 (a)：闭括号是独立的 CHAR 节点
            if (c.getType() == LaTeXNode.Type.CHAR && c.getValue() != null
                    && c.getValue().length() == 1) {
                char ch = c.getValue().charAt(0);
                if (ch == openCh) depth++;   // 嵌套加深
                if (ch == closeCh) {
                    depth--;                  // 嵌套回退
                    if (depth == 0) return j; // 找到匹配的闭括号
                }
            }
            // 情况 (b)：闭括号是 SUPERSCRIPT/SUBSCRIPT 的 base，如 )^{2}
            // AST 结构：SUPERSCRIPT(CHAR(')'), GROUP(CHAR('2')))
            if ((c.getType() == LaTeXNode.Type.SUPERSCRIPT
                    || c.getType() == LaTeXNode.Type.SUBSCRIPT)
                    && !c.getChildren().isEmpty()) {
                LaTeXNode base = c.getChildren().get(0);
                if (base.getType() == LaTeXNode.Type.CHAR && base.getValue() != null
                        && base.getValue().length() == 1 && base.getValue().charAt(0) == closeCh) {
                    depth--;
                    if (depth == 0) return j;
                }
            }
        }
        return -1; // 未找到匹配的闭括号
    }

    /**
     * 写入括号 fence 模板（TM_PAREN 或 TM_BRACK），带 FN_EXPAND 分隔符。
     *
     * <p>生成的 MTEF 结构：</p>
     * <pre>
     *   TMPL header（TM_PAREN 或 TM_BRACK）
     *     LINE                          ← 内容 slot
     *       [内容节点...]
     *     END                           ← 关闭内容 LINE
     *     [FULL]                        ← 仅当内容末尾是模板时才需要（恢复字号）
     *     CHAR FN_EXPAND openCh         ← 开分隔符（如 '('）使用 FN_EXPAND 字体
     *     CHAR FN_EXPAND closeCh        ← 闭分隔符（如 ')'）使用 FN_EXPAND 字体
     *   END                             ← 关闭模板
     * </pre>
     *
     * <p><b>FULL 规则</b>：内容 LINE END 与 FN_EXPAND 分隔符之间的 FULL，
     * 仅在内容的最后一个元素是生成模板的节点（导致字号缩小）时才插入。
     * 如果内容全是普通字符，不需要 FULL 恢复。</p>
     *
     * <p>Write a fence template (TM_PAREN, TM_BRACK) with content and FN_EXPAND delimiters.</p>
     */
    private void writeParenFence(ByteArrayOutputStream out, char openCh, char closeCh,
                                  java.util.List<LaTeXNode> content) throws IOException {
        // 根据开括号类型选择对应的模板头部
        if (openCh == '(') {
            MtefTemplateBuilder.writeParenHeader(out);
        } else if (openCh == '[') {
            MtefTemplateBuilder.writeBracketHeader(out);
        } else {
            MtefTemplateBuilder.writeParenHeader(out); // 默认使用圆括号模板
        }

        // 内容 slot：用 LINE 记录包裹括号内的所有内容
        out.write(MtefRecord.LINE);
        out.write(0x00); // options: 非空行
        for (int ci = 0; ci < content.size(); ci++) {
            LaTeXNode n = content.get(ci);
            writeNode(out, n);
            // 内容中的模板节点后面有更多内容时，也需要 FULL 恢复字号
            if (ci < content.size() - 1 && generatesTemplate(n)) {
                out.write(MtefRecord.FULL);
            }
        }
        out.write(MtefRecord.END); // 关闭内容 LINE

        // FULL 条件插入：仅当内容最后一个元素是模板节点（字号被缩小）时，
        // 需要在 FN_EXPAND 分隔符之前插入 FULL 恢复全尺寸。
        // 纯字符内容不会改变字号，无需 FULL。
        if (!content.isEmpty() && generatesTemplate(content.get(content.size() - 1))) {
            out.write(MtefRecord.FULL);
        }

        // FN_EXPAND 分隔符字符：MathType 使用特殊的 EXPAND 字体来绘制可拉伸的括号
        writeCharRecord(out, MtefRecord.FN_EXPAND, openCh);  // 开括号
        writeCharRecord(out, MtefRecord.FN_EXPAND, closeCh);  // 闭括号
        out.write(MtefRecord.END); // 关闭模板
    }

    /**
     * 写入括号后的上标/下标附件（base 为前面的 fence 模板，这里不写 base）。
     *
     * <p>用于处理 (...)^{exp} 或 (...)_{sub} 模式。前面的 writeParenFence 已经
     * 写入了括号模板作为 base，这里只需写入无 base 的上标/下标模板。</p>
     *
     * <p>支持三种模式：</p>
     * <ul>
     *   <li>(...)^{exp} → TM_SUP 模板（NULL base + 上标 slot）</li>
     *   <li>(...)_{sub} → TM_SUB 模板（下标 slot + NULL base）</li>
     *   <li>(...)_{sub}^{exp} → TM_SUBSUP 模板（NULL base + 下标 slot + 上标 slot）</li>
     * </ul>
     *
     * <p>Write a superscript/subscript attachment (without base, since the base
     * is the preceding fence template).</p>
     */
    private void writeSupSubAttachment(ByteArrayOutputStream out, LaTeXNode supSubNode) throws IOException {
        if (supSubNode.getType() == LaTeXNode.Type.SUPERSCRIPT) {
            LaTeXNode base = childAt(supSubNode, 0);
            LaTeXNode exp = childAt(supSubNode, 1);

            // 检查是否为同时带下标和上标的模式：)_{sub}^{exp}
            // AST 结构：SUPERSCRIPT(SUBSCRIPT(CHAR(')'), sub), exp)
            if (base != null && base.getType() == LaTeXNode.Type.SUBSCRIPT) {
                LaTeXNode sub = childAt(base, 1);
                MtefTemplateBuilder.writeSubSuperscriptHeader(out);
                out.write(MtefRecord.SUB);    // typesize 记录
                writeNullLine(out);            // slot 0: NULL base（base 是前面的 fence）
                writeSlot(out, sub);           // slot 1: 下标
                writeSlot(out, exp);           // slot 2: 上标
                out.write(MtefRecord.END);     // 关闭模板
                return;
            }

            // 仅上标模式：)^{exp}
            MtefTemplateBuilder.writeSuperscriptHeader(out);
            out.write(MtefRecord.SUB);    // typesize 记录
            writeNullLine(out);            // TM_SUP slot 顺序：先 NULL base
            writeSlot(out, exp);           // 再写上标内容
            out.write(MtefRecord.END);     // 关闭模板
        } else if (supSubNode.getType() == LaTeXNode.Type.SUBSCRIPT) {
            // 仅下标模式：)_{sub}
            LaTeXNode sub = childAt(supSubNode, 1);
            MtefTemplateBuilder.writeSubscriptHeader(out);
            out.write(MtefRecord.SUB);    // typesize 记录
            writeSlot(out, sub);           // TM_SUB slot 顺序：先写下标内容
            writeNullLine(out);            // 再 NULL base
            out.write(MtefRecord.END);     // 关闭模板
        }
    }

    /**
     * 大算子信息记录：保存从 AST 中提取的大算子命令名、下限和上限。
     *
     * @param cmd   大算子 LaTeX 命令（如 "\\sum", "\\int"）
     * @param lower 下限节点（可为 null，表示无下限）
     * @param upper 上限节点（可为 null，表示无上限）
     */
    private record BigOpInfo(String cmd, LaTeXNode lower, LaTeXNode upper) {}

    /**
     * 从 AST 节点中提取大算子（求和/积分/求积）信息。
     *
     * <p>大算子在 LaTeX AST 中有三种可能的结构：</p>
     * <ol>
     *   <li>同时有上下限：\sum_{i=0}^{n}
     *       <br>AST: SUPERSCRIPT( SUBSCRIPT( COMMAND(\sum), lower ), upper )</li>
     *   <li>仅有上限：\sum^{n}
     *       <br>AST: SUPERSCRIPT( COMMAND(\sum), upper )</li>
     *   <li>仅有下限：\sum_{i=0}
     *       <br>AST: SUBSCRIPT( COMMAND(\sum), lower )</li>
     * </ol>
     *
     * @return BigOpInfo 记录；非大算子节点返回 null
     */
    private BigOpInfo extractBigOp(LaTeXNode node) {
        // 模式 1 & 2：外层为 SUPERSCRIPT
        if (node.getType() == LaTeXNode.Type.SUPERSCRIPT) {
            LaTeXNode base = childAt(node, 0);
            LaTeXNode sup = childAt(node, 1);

            // 模式 1：SUPERSCRIPT(SUBSCRIPT(COMMAND, lower), upper) — 同时有上下限
            if (base != null && base.getType() == LaTeXNode.Type.SUBSCRIPT) {
                LaTeXNode innerBase = childAt(base, 0);
                LaTeXNode sub = childAt(base, 1);
                if (innerBase != null && isBigOperator(innerBase)) {
                    return new BigOpInfo(innerBase.getValue(), sub, sup);
                }
            }
            // 模式 2：SUPERSCRIPT(COMMAND, upper) — 仅有上限
            if (base != null && isBigOperator(base)) {
                return new BigOpInfo(base.getValue(), null, sup);
            }
        }
        // 模式 3：SUBSCRIPT(COMMAND, lower) — 仅有下限
        if (node.getType() == LaTeXNode.Type.SUBSCRIPT) {
            LaTeXNode base = childAt(node, 0);
            LaTeXNode sub = childAt(node, 1);
            if (base != null && isBigOperator(base)) {
                return new BigOpInfo(base.getValue(), sub, null);
            }
        }
        return null; // 非大算子节点
    }

    /**
     * 安全获取节点的第 idx 个子节点（越界时返回 null）。
     */
    private LaTeXNode childAt(LaTeXNode node, int idx) {
        return node.getChildren().size() > idx ? node.getChildren().get(idx) : null;
    }

    /**
     * 写入完整的大算子模板：TMPL 头部 → 内容 LINE → SUB(上下限) → SYM(算子符号) → END。
     *
     * <p>大算子模板的 MTEF 结构：</p>
     * <pre>
     *   TMPL header（TM_SUM / TM_INTEGRAL / TM_PRODUCT）
     *     LINE                          ← slot 1: 积分/求和内容（被积函数/求和表达式）
     *       [内容节点...]               ← 大算子之后的所有剩余兄弟节点
     *     END
     *     SUB                           ← typesize 记录
     *       LINE lower END              ← 下限 slot（如 i=0）
     *       LINE upper END              ← 上限 slot（如 n）
     *     SYM                           ← 算子符号记录
     *       CHAR(∑ / ∫ / ∏)
     *   END                             ← 关闭模板
     * </pre>
     *
     * <p>contentNodes 是大算子节点之后的所有剩余兄弟节点，它们作为积分/求和的内容。
     * 内容 slot 内部使用 writeContentNodes 进行 FULL 插入和括号检测。</p>
     *
     * <p>Write complete big operator: TMPL header → content LINE → SUB(limits) → SYM(operator) → END.
     * Content is the remaining sibling nodes after the big-op sub/sup node.</p>
     */
    private void writeBigOpComplete(ByteArrayOutputStream out, BigOpInfo bigOp,
                                     java.util.List<LaTeXNode> contentNodes) throws IOException {
        if (BIG_OP_LIMIT_LIKE.contains(bigOp.cmd())) {
            writeLimitComplete(out, bigOp, contentNodes);
            return;
        }

        // 写入大算子模板头部（根据命令类型选择 TM_SUM / TM_INTEGRAL / TM_PRODUCT）
        writeBigOpHeader(out, bigOp.cmd(), bigOp.lower() != null, bigOp.upper() != null);

        // Slot 1: 内容 LINE（被积函数 / 求和表达式）
        // 使用 writeContentNodes 以支持内容中的 FULL 插入和括号检测
        out.write(MtefRecord.LINE);
        out.write(0x00);
        writeContentNodes(out, contentNodes);
        out.write(MtefRecord.END);

        // 写入尾部：SUB(上下限 slots) + SYM(算子字符) + END
        writeBigOpFooter(out, bigOp.cmd(), bigOp.lower(), bigOp.upper());
    }

    private void writeLimitComplete(ByteArrayOutputStream out, BigOpInfo bigOp,
                                    java.util.List<LaTeXNode> contentNodes) throws IOException {
        MtefTemplateBuilder.writeLimitHeader(out, bigOp.lower() != null, bigOp.upper() != null);

        out.write(MtefRecord.LINE);
        out.write(0x00);
        writeContentNodes(out, contentNodes);
        out.write(MtefRecord.END);

        if (bigOp.lower() != null) {
            writeSlot(out, bigOp.lower());
        }
        if (bigOp.upper() != null) {
            writeSlot(out, bigOp.upper());
        }
        out.write(MtefRecord.END);
    }

    /**
     * 写入字符节点：将 AST 中的 CHAR 节点转换为一条 MTEF CHAR 记录。
     *
     * <p>先通过 MtefCharMap 查找字符的 MTEF 字体类型和编码。
     * 如果字符在映射表中（如希腊字母、数学符号），使用映射的字体和编码；
     * 否则默认使用 FN_VARIABLE（变量）字体，直接使用字符的 Unicode 码。</p>
     */
    private void writeCharNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        String ch = node.getValue();
        if ("\t".equals(ch)) {
            // MathType 制表符要以文本字符写入，不能落到默认变量字体。
            writeCharRecord(out, MtefRecord.FN_TEXT, '\t');
            return;
        }
        MtefCharMap.CharEntry entry = MtefCharMap.lookupChar(ch.charAt(0));
        if (entry != null) {
            // 在字符映射表中找到 — 使用指定的字体类型和 MTEF 字符编码
            writeCharRecord(out, entry.typeface(), entry.mtcode());
        } else {
            // 未在映射表中 — 默认为变量字体（FN_VARIABLE），使用原始字符码
            writeCharRecord(out, MtefRecord.FN_VARIABLE, ch.charAt(0));
        }
    }

    /**
     * 写入 LaTeX 命令节点：根据命令类型生成相应的 MTEF 记录。
     *
     * <p>命令处理优先级：</p>
     * <ol>
     *   <li>间距命令（\,  \;  \quad 等）→ 跳过（MathType 自动处理间距）</li>
     *   <li>字符映射命令（\alpha, \infty 等）→ 查表生成 CHAR 记录</li>
     *   <li>装饰命令（\overline, \hat, \vec 等）→ 生成对应 TMPL 模板</li>
     *   <li>显式分隔符命令（\left(, \left[ 等）→ 生成 fence 模板</li>
     *   <li>函数名命令（\sin, \cos 等）→ 生成 FN_FUNCTION 字体的字符序列</li>
     * </ol>
     */

    /** LaTeX 间距命令集合 — 这些命令在 MathType 中被忽略，因为 MathType 有自己的间距算法 */
    private static final Set<String> LATEX_SPACING_COMMANDS = Set.of(
        "\\,", "\\;", "\\:", "\\!", "\\quad", "\\qquad", "\\hspace", "\\hskip"
    );

    private void writeCommandNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        String cmd = node.getValue();

        // 1. 跳过 LaTeX 间距命令 — MathType 自动处理数学公式间距
        if (LATEX_SPACING_COMMANDS.contains(cmd)) {
            return;
        }

        if (cmd != null && cmd.startsWith("\\left")) {
            if (tryWriteExplicitFence(out, node)) {
                return;
            }
        }

        // 2. 查找字符映射表：将 LaTeX 命令（如 \alpha, \infty）转换为 MTEF CHAR 记录
        // 大算子（\sum, \int）在作为独立命令出现时也会命中这里
        MtefCharMap.CharEntry entry = MtefCharMap.lookup(cmd);
        if (entry != null) {
            writeCharRecord(out, entry.typeface(), entry.mtcode());
            // 如果命令有子节点（如 \overline{x} 被解析为 COMMAND 时），继续写入子节点
            writeChildren(out, node);
            return;
        }

        // 3. 带子节点的装饰/分隔符命令 — 生成对应的 MTEF 模板
        switch (cmd) {
            case "\\overline" -> {
                // 上划线装饰：生成 TM_BAR(上方) 模板
                MtefTemplateBuilder.writeOverlineHeader(out);
                writeSlot(out, node.getChildren().isEmpty() ? null : node.getChildren().get(0));
                out.write(MtefRecord.END);
            }
            case "\\underline" -> {
                // 下划线装饰：生成 TM_UBAR(下方) 模板
                MtefTemplateBuilder.writeUnderlineHeader(out);
                LaTeXNode underlineContent = node.getChildren().isEmpty() ? null : node.getChildren().get(0);
                if ("true".equals(node.getMetadata(VerticalLayoutNodeFactory.RAW_INLINE_CONTAINER))) {
                    out.write(MtefRecord.END);
                    writeSlot(out, underlineContent);
                } else {
                    writeSlot(out, underlineContent);
                    out.write(MtefRecord.END);
                }
            }
            case "\\vec" -> {
                // 向量箭头装饰：生成 TM_HAT 模板 + FN_EXPAND 组合右箭头字符 (U+20D7)
                MtefTemplateBuilder.writeVecHeader(out);
                writeSlot(out, node.getChildren().isEmpty() ? null : node.getChildren().get(0));
                writeCharRecord(out, MtefRecord.FN_EXPAND, 0x20D7);
                out.write(MtefRecord.END);
            }
            case "\\hat" -> {
                // 帽子装饰（^）：生成 TM_HAT 模板
                MtefTemplateBuilder.writeHatHeader(out);
                writeSlot(out, node.getChildren().isEmpty() ? null : node.getChildren().get(0));
                out.write(MtefRecord.END);
            }
            case "\\tilde" -> {
                // 波浪号装饰（~）：生成 TM_TILDE 模板
                MtefTemplateBuilder.writeTildeHeader(out);
                writeSlot(out, node.getChildren().isEmpty() ? null : node.getChildren().get(0));
                out.write(MtefRecord.END);
            }
            case "\\bar" -> {
                // 短上划线装饰：与 \overline 相同，使用 TM_BAR 模板
                MtefTemplateBuilder.writeOverlineHeader(out);
                writeSlot(out, node.getChildren().isEmpty() ? null : node.getChildren().get(0));
                out.write(MtefRecord.END);
            }
            case "\\dot" -> {
                // 点装饰：使用 EMBELL（修饰）机制，在字符上方添加单点
                if (!node.getChildren().isEmpty()) {
                    writeNodeWithEmbellishment(out, node.getChildren().get(0), MtefRecord.EMB_1DOT);
                }
            }
            default -> {
                // 4. 数学函数名（如 \sin, \cos, \log）→ 使用 FN_FUNCTION 字体逐字符写入
                if (cmd.startsWith("\\")) {
                    String funcName = cmd.substring(1);
                    if (isFunctionName(funcName)) {
                        writeFunctionName(out, funcName);
                    }
                }
                // 处理命令可能携带的子节点
                writeChildren(out, node);
            }
        }
    }

    private boolean tryWriteExplicitFence(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        FenceSpec spec = resolveFenceSpec(node.getMetadata("leftDelimiter"), node.getMetadata("rightDelimiter"));
        if (spec == null) {
            return false;
        }
        writeFenceTemplate(out, spec, node.getChildren().isEmpty() ? null : node.getChildren().get(0));
        return true;
    }

    private FenceSpec resolveFenceSpec(String leftDelimiter, String rightDelimiter) {
        String left = leftDelimiter == null ? "" : leftDelimiter;
        String right = rightDelimiter == null ? "" : rightDelimiter;
        if ("(".equals(left) || ")".equals(right)) {
            return new FenceSpec(MtefRecord.TM_PAREN, '(', ')', !".".equals(left), !".".equals(right));
        }
        if ("[".equals(left) || "]".equals(right)) {
            return new FenceSpec(MtefRecord.TM_BRACK, '[', ']', !".".equals(left), !".".equals(right));
        }
        if ("{".equals(left) || "\\{".equals(left) || "}".equals(right)) {
            return new FenceSpec(MtefRecord.TM_BRACE, '{', '}', !".".equals(left), !".".equals(right));
        }
        if ("|".equals(left) || "|".equals(right)) {
            return new FenceSpec(MtefRecord.TM_BAR, '|', '|', !".".equals(left), !".".equals(right));
        }
        if ("||".equals(left) || "||".equals(right)) {
            return new FenceSpec(MtefRecord.TM_DBAR, 0x2016, 0x2016, !".".equals(left), !".".equals(right));
        }
        if ("⌊".equals(left) || "⌋".equals(right)) {
            return new FenceSpec(MtefRecord.TM_FLOOR, 0x230A, 0x230B, !".".equals(left), !".".equals(right));
        }
        if ("⌈".equals(left) || "⌉".equals(right)) {
            return new FenceSpec(MtefRecord.TM_CEILING, 0x2308, 0x2309, !".".equals(left), !".".equals(right));
        }
        return null;
    }

    private void writeFenceTemplate(ByteArrayOutputStream out, FenceSpec spec, LaTeXNode content) throws IOException {
        switch (spec.selector()) {
            case MtefRecord.TM_PAREN -> MtefTemplateBuilder.writeFenceHeader(out, MtefRecord.TM_PAREN, spec.hasLeft(), spec.hasRight());
            case MtefRecord.TM_BRACK -> MtefTemplateBuilder.writeFenceHeader(out, MtefRecord.TM_BRACK, spec.hasLeft(), spec.hasRight());
            case MtefRecord.TM_BRACE -> MtefTemplateBuilder.writeFenceHeader(out, MtefRecord.TM_BRACE, spec.hasLeft(), spec.hasRight());
            case MtefRecord.TM_BAR -> MtefTemplateBuilder.writeFenceHeader(out, MtefRecord.TM_BAR, spec.hasLeft(), spec.hasRight());
            case MtefRecord.TM_DBAR -> MtefTemplateBuilder.writeDoubleBarHeader(out, spec.hasLeft(), spec.hasRight());
            case MtefRecord.TM_FLOOR -> MtefTemplateBuilder.writeFloorHeader(out, spec.hasLeft(), spec.hasRight());
            case MtefRecord.TM_CEILING -> MtefTemplateBuilder.writeCeilingHeader(out, spec.hasLeft(), spec.hasRight());
            default -> throw new IllegalArgumentException("Unsupported fence selector: " + spec.selector());
        }
        writeSlot(out, content);
        if (needsFullAfterSlot(content)) {
            out.write(MtefRecord.FULL);
        }
        if (spec.hasLeft()) {
            writeCharRecord(out, MtefRecord.FN_EXPAND, spec.leftChar());
        }
        if (spec.hasRight()) {
            writeCharRecord(out, MtefRecord.FN_EXPAND, spec.rightChar());
        }
        out.write(MtefRecord.END);
    }

    private record FenceSpec(int selector, int leftChar, int rightChar, boolean hasLeft, boolean hasRight) {}

    /**
     * 写入分数节点：生成 TM_FRACT 分数模板。
     *
     * <p>生成的 MTEF 结构：</p>
     * <pre>
     *   TMPL TM_FRACT
     *     LINE numerator END        ← slot 0: 分子
     *     [FULL]                    ← 条件性 FULL（见下文说明）
     *     LINE denominator END      ← slot 1: 分母
     *   END
     * </pre>
     *
     * <p><b>分子与分母之间的 FULL 规则（needsFullAfterSlot）：</b></p>
     * <ul>
     *   <li>如果分子的最后一个有效元素是缩小字号的模板（TM_SUP / TM_SUB / TM_ROOT），
     *       需要在分子 slot 之后插入 FULL 恢复字号，否则分母字号会不正确</li>
     *   <li>如果分子末尾是 TM_PAREN（括号模板自带内部 FULL），则不需要额外 FULL</li>
     *   <li>如果分子末尾是普通字符，也不需要 FULL</li>
     * </ul>
     *
     * <p>Write fraction: TMPL(FRACT) + numerator slot + [FULL] + denominator slot.
     * FULL between slots is only needed when the numerator's last effective element
     * leaves the size context in a reduced state (TM_SUP/TM_SUB/TM_ROOT).
     * When the numerator ends with TM_PAREN (which has internal FULL), no extra FULL is needed.</p>
     */
    private void writeFractionNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        MtefTemplateBuilder.writeFractionHeader(out);
        LaTeXNode numerator = node.getChildren().size() > 0 ? node.getChildren().get(0) : null;
        // Slot 0: 分子
        writeSlot(out, numerator);
        // 条件性 FULL：仅当分子末尾是缩小字号的模板时才插入
        if (needsFullAfterSlot(numerator)) {
            out.write(MtefRecord.FULL);
        }
        // Slot 1: 分母
        writeSlot(out, node.getChildren().size() > 1 ? node.getChildren().get(1) : null);
        out.write(MtefRecord.END); // 关闭分数模板
    }

    /**
     * 判断一个 slot 的内容是否会使字号上下文处于缩小状态，需要在 slot 之后插入 FULL 恢复。
     *
     * <p>FULL 需求规则：</p>
     * <ul>
     *   <li>TM_SUP（上标）、TM_SUB（下标）、TM_ROOT（根号）内部使用 SUB typesize，
     *       会将字号缩小；如果 slot 以这些模板结尾，字号处于缩小状态 → 需要 FULL</li>
     *   <li>TM_PAREN（括号）内部自带 FULL 恢复，不会留下缩小的字号 → 不需要 FULL</li>
     *   <li>普通字符不改变字号 → 不需要 FULL</li>
     * </ul>
     *
     * <p>Check if a slot's content leaves the size context in a reduced state,
     * requiring a FULL record after the slot to restore normal size.
     * TM_SUP, TM_SUB, TM_ROOT use SUB typesize internally and leave size reduced.
     * TM_PAREN has its own internal FULL and doesn't leave size reduced.</p>
     */
    private boolean needsFullAfterSlot(LaTeXNode slotNode) {
        LaTeXNode last = getEffectiveLastChild(slotNode);
        if (last == null) return false;
        return switch (last.getType()) {
            case SUPERSCRIPT, SUBSCRIPT, SQRT -> true;  // 缩小字号的模板 → 需要 FULL
            default -> false;                            // 字符或括号 → 不需要
        };
    }

    /**
     * 获取节点的"有效最后子节点"，递归穿透 GROUP/ROOT 包装层。
     *
     * <p>LaTeX 解析器可能生成多层嵌套的 GROUP/ROOT 节点，
     * 这个方法递归深入找到实际的最后一个内容节点。</p>
     *
     * <p>Get the effective last child of a node, recursing through GROUP/ROOT wrappers.</p>
     */
    private LaTeXNode getEffectiveLastChild(LaTeXNode node) {
        if (node == null) return null;
        if (node.getType() == LaTeXNode.Type.ROOT || node.getType() == LaTeXNode.Type.GROUP) {
            java.util.List<LaTeXNode> children = node.getChildren();
            if (children.isEmpty()) return null;
            LaTeXNode last = children.get(children.size() - 1);
            // 如果最后一个子节点仍然是 GROUP/ROOT 包装，继续递归
            if (last.getType() == LaTeXNode.Type.ROOT || last.getType() == LaTeXNode.Type.GROUP) {
                return getEffectiveLastChild(last);
            }
            return last;
        }
        return node;
    }

    /**
     * 写入根号节点：生成 TM_ROOT 根号模板。
     *
     * <p>MathType 根号模板结构（通过参考二进制确认）：</p>
     * <pre>
     *   TMPL TM_ROOT var=0x00(平方根)/0x01(n次根) tmplOpts=0x00
     *     LINE content END          ← slot 0: 被开方数（radicand）
     *     SUB                       ← 1 字节 typesize 记录（标记后续 slot 使用缩小字号）
     *     LINE [degree/NULL] END    ← slot 1: 根的次数（n 次根为 n，平方根为 NULL LINE）
     *   END
     * </pre>
     *
     * <p>注意：</p>
     * <ul>
     *   <li>平方根 \sqrt{x} → variation=0x00，slot 1 为 NULL LINE</li>
     *   <li>n 次根 \sqrt[n]{x} → variation=0x01 (TV_ROOT_NTH)，slot 1 包含次数 n</li>
     *   <li>AST 子节点顺序：[0]=次数（可选），[1]=被开方数；但 MTEF slot 顺序相反</li>
     *   <li>FULL 由 writeChildren 在有更多兄弟节点时自动添加</li>
     * </ul>
     *
     * <p>Write sqrt: TMPL(ROOT) + content slot + SUB + degree slot.</p>
     */
    private void writeSqrtNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        if (node.getChildren().size() == 2) {
            // n 次根：\sqrt[n]{content}，AST 子节点 [0]=次数 n, [1]=被开方数
            MtefTemplateBuilder.writeNthRootHeader(out);       // variation = TV_ROOT_NTH (0x01)
            writeSlot(out, node.getChildren().get(1));         // slot 0: 被开方数（content）
            out.write(MtefRecord.SUB);                         // 1 字节 typesize 记录
            writeSlot(out, node.getChildren().get(0));         // slot 1: 次数（n）
        } else if (node.getChildren().size() == 1) {
            // 平方根：\sqrt{content}，AST 子节点 [0]=被开方数
            MtefTemplateBuilder.writeSqrtHeader(out);          // variation = 0x00（标准平方根）
            writeSlot(out, node.getChildren().get(0));         // slot 0: 被开方数（content）
            out.write(MtefRecord.SUB);                         // 1 字节 typesize 记录
            writeNullLine(out);                                // slot 1: 次数为空（NULL LINE）
        }
        out.write(MtefRecord.END); // 关闭根号模板
        // FULL 由外层 writeContentNodes 在有更多兄弟节点时自动插入
    }

    /**
     * 写入上标节点：base 写在模板外部，然后生成 TM_SUP 上标模板。
     *
     * <p><b>模板写入模式</b>（MathType 的核心设计）：base 字符写在模板外部（父级 LINE 中），
     * 模板内部的 base slot 使用 NULL LINE 占位。这样 MathType 可以将上标
     * 正确地附加到前面的字符上。</p>
     *
     * <p>MathType 结构（通过参考二进制确认）：</p>
     * <pre>
     *   CHAR 'base'                   ← 写在模板 **外部**（父级 LINE 中）
     *   TMPL TM_SUP var=0 tmplOpts=0
     *     SUB                          ← 1 字节 typesize 记录
     *     LINE [NULL]                  ← slot 0: base（NULL，因为 base 已在外部写入）
     *     LINE                         ← slot 1: 上标内容
     *       CHAR 'exp'
     *     END
     *   END
     * </pre>
     *
     * <p>特殊处理：</p>
     * <ul>
     *   <li>如果节点是大算子模式（如 \sum^n），转发给 writeBigOpComplete</li>
     *   <li>如果 base 本身是下标节点（如 x_{i}^{2}），生成 TM_SUBSUP 同时上下标模板</li>
     * </ul>
     *
     * <p>Write superscript: base CHAR outside, then TMPL(SUP) with NULL base slot.</p>
     */
    private void writeSuperscriptNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        LaTeXNode base = node.getChildren().size() > 0 ? node.getChildren().get(0) : null;
        LaTeXNode sup = node.getChildren().size() > 1 ? node.getChildren().get(1) : null;

        // 优先检查是否为大算子模式（如 \sum_{i=0}^{n}）
        BigOpInfo bigOp = extractBigOp(node);
        if (bigOp != null) {
            writeBigOpComplete(out, bigOp, java.util.List.of());
            return;
        }

        // 检查同时上下标模式：base 是 SUBSCRIPT 节点，如 x_{i}^{2}
        // AST 结构：SUPERSCRIPT( SUBSCRIPT(innerBase, sub), sup )
        if (base != null && base.getType() == LaTeXNode.Type.SUBSCRIPT) {
            LaTeXNode innerBase = base.getChildren().size() > 0 ? base.getChildren().get(0) : null;
            LaTeXNode sub = base.getChildren().size() > 1 ? base.getChildren().get(1) : null;
            // base 写在模板外部
            if (innerBase != null) writeNode(out, innerBase);
            MtefTemplateBuilder.writeSubSuperscriptHeader(out);  // TM_SUBSUP 模板
            out.write(MtefRecord.SUB);    // typesize 记录
            writeNullLine(out);            // slot 0: NULL base
            writeSlot(out, sub);           // slot 1: 下标
            writeSlot(out, sup);           // slot 2: 上标
            out.write(MtefRecord.END);     // 关闭模板
            return;
        }

        // 标准上标模式：base^{sup}
        if (base != null) writeNode(out, base);    // base 写在模板外部
        MtefTemplateBuilder.writeSuperscriptHeader(out);  // TM_SUP 模板
        out.write(MtefRecord.SUB);    // typesize 记录
        writeNullLine(out);            // slot 0: NULL base（base 已在外部写入）
        writeSlot(out, sup);           // slot 1: 上标内容
        out.write(MtefRecord.END);     // 关闭模板
        // FULL 由外层 writeContentNodes 在有更多兄弟节点时自动插入
    }

    /**
     * 写入下标节点：base 写在模板外部，然后生成 TM_SUB 下标模板。
     *
     * <p><b>重要：TM_SUB 的 slot 顺序与 TM_SUP 相反！</b></p>
     * <pre>
     *   TM_SUP: SUB → NULL base（先） → content（后）
     *   TM_SUB: SUB → content（先） → NULL base（后）
     * </pre>
     *
     * <p>MathType TM_SUB 结构（通过参考二进制确认）：</p>
     * <pre>
     *   CHAR 'base'              ← 写在模板外部
     *   TMPL TM_SUB var=0 tmplOpts=0
     *     SUB                     ← 1 字节 typesize 记录
     *     LINE content END        ← slot 0: 下标内容（**先写**！）
     *     LINE [NULL]             ← slot 1: base（NULL，**后写**！）
     *   END
     * </pre>
     *
     * <p>Write subscript: base CHAR outside, then TMPL(SUB).
     * NOTE: TM_SUB slot order is OPPOSITE to TM_SUP!</p>
     */
    private void writeSubscriptNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        LaTeXNode base = node.getChildren().size() > 0 ? node.getChildren().get(0) : null;
        LaTeXNode sub = node.getChildren().size() > 1 ? node.getChildren().get(1) : null;

        // 优先检查是否为大算子模式（如 \sum_{i=0}）
        BigOpInfo bigOp = extractBigOp(node);
        if (bigOp != null) {
            writeBigOpComplete(out, bigOp, java.util.List.of());
            return;
        }

        // base 写在模板外部
        if (base != null) writeNode(out, base);
        MtefTemplateBuilder.writeSubscriptHeader(out);  // TM_SUB 模板
        out.write(MtefRecord.SUB);    // typesize 记录
        writeSlot(out, sub);           // slot 0: 下标内容（TM_SUB 先写内容！）
        writeNullLine(out);            // slot 1: NULL base（TM_SUB 后写 base！）
        out.write(MtefRecord.END);     // 关闭模板
        // FULL 由外层 writeContentNodes 在有更多兄弟节点时自动插入
    }

    /**
     * 写入文本节点：将 \text{...} 或 \mathrm{...} 中的文本内容转换为 FN_TEXT 字体的 CHAR 记录。
     *
     * <p>文本节点的子节点可能包含 GROUP 包装层，需要递归展开。
     * 每个字符使用 FN_TEXT（文本）字体，而非 FN_VARIABLE（变量）字体，
     * 这样 MathType 会以直立体（upright）而非斜体（italic）显示。</p>
     */
    private void writeTextNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        for (LaTeXNode child : node.getChildren()) {
            if (child.getType() == LaTeXNode.Type.GROUP) {
                // GROUP 包装层：展开其子节点
                for (LaTeXNode gc : child.getChildren()) {
                    writeTextChar(out, gc);
                }
            } else {
                writeTextChar(out, child);
            }
        }
    }

    /**
     * 将单个文本字符写为 FN_TEXT 字体的 CHAR 记录。
     * 支持多字符值（逐字符写入）。
     */
    private void writeTextChar(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        if (node.getType() == LaTeXNode.Type.CHAR && node.getValue() != null) {
            for (char c : node.getValue().toCharArray()) {
                writeCharRecord(out, MtefRecord.FN_TEXT, c);
            }
        }
    }

    private void writeLongDivisionNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        LaTeXNode divisor = childAt(node, 0);
        LaTeXNode quotient = childAt(node, 1);
        LaTeXNode dividend = childAt(node, 2);
        boolean hasQuotient = quotient != null && !quotient.getChildren().isEmpty();

        if (!"true".equals(node.getMetadata("skipLongDivisionLeadingDivisor")) && divisor != null) {
            writeNode(out, divisor);
        }
        MtefTemplateBuilder.writeLongDivisionHeader(out, hasQuotient);
        writeSlot(out, dividend);
        if (hasQuotient) {
            writeSlot(out, quotient);
        }
        out.write(MtefRecord.END);
    }

    private void writeArrayNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        // PRESERVE_RAW_ARRAY check must come first to avoid composite long division processing destroying structure
        if ("true".equals(node.getMetadata(VerticalLayoutNodeFactory.PRESERVE_RAW_ARRAY))) {
            writeMatrixNode(out, node);
            return;
        }
        String environment = node.getMetadata("environment");
        if ("cases".equals(environment)) {
            LaTeXNode content = shallowCloneArray(node);
            content.getMetadata().remove("environment");
            writeFenceTemplate(out, new FenceSpec(MtefRecord.TM_BRACE, '{', '}', true, false), content);
            return;
        }
        if (isAlignedRelationArray(node)) {
            writeAlignedRelationPile(out, node);
            return;
        }
        VerticalLayoutCompiler.CrossMultiplicationLayout crossLayout = verticalLayoutCompiler.compileCrossMultiplicationArray(node);
        if (crossLayout != null) {
            // 十字交叉需要保留参考里的嵌套矩阵结构，不能压平成普通算术表格。
            writeMatrixNode(out, verticalLayoutNodeFactory.buildCrossMultiplicationArray(crossLayout));
            return;
        }
        VerticalLayoutSpec layoutSpec = verticalLayoutCompiler.compileArray(node);
        if (layoutSpec != null && layoutSpec.kind() == VerticalLayoutSpec.Kind.DECIMAL) {
            // 小数加减法单独走单列 PILE，每行写完整小数，避免拆成多列。
            pileRulerWriter.writeLayout(out, layoutSpec, this::writeNode);
            return;
        }
        LaTeXNode normalized = layoutSpec != null ? verticalLayoutNodeFactory.buildArrayNode(layoutSpec) : node;
        writeMatrixNode(out, normalized);
    }

    private void writeAlignedRelationPile(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        int columnCount = resolveArrayColumnCount(node);

        out.write(MtefRecord.PILE);
        out.write(MtefRecord.OPT_LP_RULER);
        out.write(0x01);
        out.write(0x02);
        pileRulerWriter.writeRuler(out, buildAlignedRelationTabStops(columnCount));

        for (LaTeXNode row : node.getChildren()) {
            out.write(MtefRecord.LINE);
            out.write(0x00);
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                writeTabChar(out);
                if (columnIndex < row.getChildren().size()) {
                    writeNode(out, row.getChildren().get(columnIndex));
                }
            }
            writeTabChar(out);
            out.write(MtefRecord.END);
        }
        out.write(MtefRecord.END);
    }

    private List<VerticalLayoutSpec.VerticalTabStop> buildAlignedRelationTabStops(int columnCount) {
        List<VerticalLayoutSpec.VerticalTabStop> tabStops = new ArrayList<>();
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            VerticalLayoutSpec.TabStopKind kind = columnIndex % 2 == 0
                ? VerticalLayoutSpec.TabStopKind.RIGHT
                : VerticalLayoutSpec.TabStopKind.RELATION;
            tabStops.add(new VerticalLayoutSpec.VerticalTabStop(columnIndex, kind, (columnIndex + 1) * 240));
        }
        return tabStops;
    }

    private boolean isAlignedRelationArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY) {
            return false;
        }
        if (!"aligned".equals(node.getMetadata("environment"))) {
            return false;
        }
        if (!"relation-pairs".equals(node.getMetadata("alignmentMode"))) {
            return false;
        }
        int columnCount = resolveArrayColumnCount(node);
        if (columnCount <= 0 || (columnCount % 2) != 0) {
            return false;
        }
        for (LaTeXNode row : node.getChildren()) {
            if (row.getChildren().size() > columnCount) {
                return false;
            }
        }
        return true;
    }

    private void writeTabChar(ByteArrayOutputStream out) throws IOException {
        writeCharRecord(out, MtefRecord.FN_TEXT, '\t');
    }

    private void writeMatrixNode(ByteArrayOutputStream out, LaTeXNode normalized) throws IOException {
        int rows = normalized.getChildren().size();
        int cols = resolveArrayColumnCount(normalized);

        out.write(MtefRecord.MATRIX);
        out.write(0x00);
        out.write(resolveMatrixVerticalAlignment(normalized));
        out.write(resolveMatrixHorizontalAlignment(normalized));
        out.write(resolveMatrixVerticalJustification(normalized));
        writeUnsignedInt(out, rows);
        writeUnsignedInt(out, cols);
        writePartitionLineArray(out, parsePartitionArray(normalized.getMetadata("rowLines"), rows + 1));
        writePartitionLineArray(out, parsePartitionArray(normalized.getMetadata("columnLines"), cols + 1));

        for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
            LaTeXNode row = normalized.getChildren().get(rowIndex);
            for (int colIndex = 0; colIndex < cols; colIndex++) {
                LaTeXNode cell = colIndex < row.getChildren().size() ? row.getChildren().get(colIndex) : null;
                writeMatrixCell(out, cell);
            }
        }
        out.write(MtefRecord.END);
    }

    private LaTeXNode normalizeArrayNode(LaTeXNode node) {
        if (node == null || node.getChildren().isEmpty()) {
            return node;
        }
        VerticalLayoutSpec layoutSpec = verticalLayoutCompiler.compileArray(node);
        if (layoutSpec != null) {
            return verticalLayoutNodeFactory.buildArrayNode(layoutSpec);
        }
        return node;
    }

    private boolean isMultiplicationArray(LaTeXNode node) {
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                String text = flattenNodeText(cell);
                if (text.contains("\\times") || text.contains("×")) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean isDecimalArray(LaTeXNode node) {
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                if (".".equals(flattenNodeText(cell))) {
                    return true;
                }
            }
        }
        return false;
    }

    private LaTeXNode normalizeMultiplicationArray(LaTeXNode node) {
        LaTeXNode normalized = shallowCloneArray(node);
        int maxRowSize = 0;
        for (LaTeXNode row : normalized.getChildren()) {
            maxRowSize = Math.max(maxRowSize, row.getChildren().size());
        }
        int contentCols = maxRowSize;
        int targetCols = contentCols + 1; // 末尾哨兵空列，锁定右边界
        for (LaTeXNode row : normalized.getChildren()) {
            int padLeft = contentCols - row.getChildren().size();
            if (padLeft > 0) {
                List<LaTeXNode> padded = new ArrayList<>();
                for (int i = 0; i < padLeft; i++) {
                    padded.add(new LaTeXNode(LaTeXNode.Type.CELL));
                }
                padded.addAll(row.getChildren());
                row.getChildren().clear();
                row.getChildren().addAll(padded);
            }
            row.addChild(new LaTeXNode(LaTeXNode.Type.CELL));
        }
        normalized.setMetadata("columnCount", String.valueOf(targetCols));
        normalized.setMetadata("columnSpec", "r".repeat(Math.max(targetCols, 0)));
        normalized.setMetadata("columnLines", encodeZeroPartitionArray(targetCols + 1));
        return normalized;
    }

    private LaTeXNode normalizeDecimalArray(LaTeXNode node) {
        LaTeXNode normalized = shallowCloneArray(node);
        for (LaTeXNode row : normalized.getChildren()) {
            List<LaTeXNode> originalCells = new ArrayList<>(row.getChildren());
            row.getChildren().clear();

            String first = originalCells.isEmpty() ? "" : flattenNodeText(originalCells.get(0));
            String sign = "";
            String integer = first;
            if (!first.isEmpty() && (first.charAt(0) == '+' || first.charAt(0) == '-')) {
                sign = String.valueOf(first.charAt(0));
                integer = first.substring(1);
            }
            String dot = originalCells.size() > 1 ? flattenNodeText(originalCells.get(1)) : "";
            String fraction = originalCells.size() > 2 ? flattenNodeText(originalCells.get(2)) : "";

            row.addChild(buildCellFromText(sign));
            row.addChild(buildCellFromText(integer));
            row.addChild(buildCellFromText(dot));
            row.addChild(buildCellFromText(fraction));
        }
        normalized.setMetadata("columnCount", "4");
        normalized.setMetadata("columnSpec", "rcrl");
        normalized.setMetadata("columnLines", encodeZeroPartitionArray(5));
        return normalized;
    }

    private LaTeXNode shallowCloneArray(LaTeXNode node) {
        LaTeXNode copy = new LaTeXNode(node.getType(), node.getValue());
        copy.getMetadata().putAll(node.getMetadata());
        for (LaTeXNode row : node.getChildren()) {
            LaTeXNode rowCopy = new LaTeXNode(row.getType(), row.getValue());
            rowCopy.getMetadata().putAll(row.getMetadata());
            for (LaTeXNode cell : row.getChildren()) {
                rowCopy.addChild(deepCloneNode(cell));
            }
            copy.addChild(rowCopy);
        }
        return copy;
    }

    private LaTeXNode deepCloneNode(LaTeXNode node) {
        LaTeXNode copy = new LaTeXNode(node.getType(), node.getValue());
        copy.getMetadata().putAll(node.getMetadata());
        for (LaTeXNode child : node.getChildren()) {
            copy.addChild(deepCloneNode(child));
        }
        return copy;
    }

    private LaTeXNode buildCellFromText(String text) {
        LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
        if (text == null || text.isEmpty()) {
            return cell;
        }
        for (char ch : text.toCharArray()) {
            cell.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        return cell;
    }

    private String flattenNodeText(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getValue() != null && (node.getType() == LaTeXNode.Type.CHAR || node.getType() == LaTeXNode.Type.COMMAND)) {
            return node.getValue();
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flattenNodeText(child));
        }
        return builder.toString();
    }

    private String encodeZeroPartitionArray(int size) {
        int[] parts = new int[Math.max(size, 1)];
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) {
                builder.append(',');
            }
            builder.append('0');
        }
        return builder.toString();
    }

    private int resolveArrayColumnCount(LaTeXNode node) {
        String metadata = node.getMetadata("columnCount");
        if (metadata != null && !metadata.isBlank()) {
            return Integer.parseInt(metadata);
        }
        int maxColumns = 0;
        for (LaTeXNode row : node.getChildren()) {
            maxColumns = Math.max(maxColumns, row.getChildren().size());
        }
        return maxColumns;
    }

    private int resolveMatrixVerticalAlignment(LaTeXNode node) {
        String spec = node.getMetadata("columnSpec");
        if (spec != null && spec.contains("|")) {
            return 2;
        }
        return 1;
    }

    private int resolveMatrixHorizontalAlignment(LaTeXNode node) {
        String explicit = node.getMetadata(VerticalLayoutNodeFactory.RAW_MATRIX_HALIGN);
        if (explicit != null && !explicit.isBlank()) {
            return Integer.parseInt(explicit);
        }
        String spec = node.getMetadata("columnSpec");
        if (spec == null || spec.isBlank()) {
            return 2;
        }
        boolean hasLeft = spec.indexOf('l') >= 0;
        boolean hasCenter = spec.indexOf('c') >= 0;
        boolean hasRight = spec.indexOf('r') >= 0;
        if (hasRight && !hasLeft && !hasCenter) {
            return 3;
        }
        if (hasCenter && !hasLeft && !hasRight) {
            return 2;
        }
        if (hasLeft && !hasCenter && !hasRight) {
            return 1;
        }
        return hasRight ? 3 : (hasCenter ? 2 : 1);
    }

    private int resolveMatrixVerticalJustification(LaTeXNode node) {
        return 1;
    }

    private void writeMatrixCell(ByteArrayOutputStream out, LaTeXNode cell) throws IOException {
        if (cell == null || cell.getChildren().isEmpty()) {
            writeNullLine(out);
            return;
        }
        writeSlot(out, cell);
    }

    private int[] parsePartitionArray(String encoded, int expectedSize) {
        int[] parts = new int[Math.max(expectedSize, 1)];
        if (encoded == null || encoded.isBlank()) {
            return parts;
        }
        String[] values = encoded.split(",");
        for (int i = 0; i < values.length && i < parts.length; i++) {
            try {
                parts[i] = Integer.parseInt(values[i].trim());
            } catch (NumberFormatException ignored) {
                parts[i] = 0;
            }
        }
        return parts;
    }

    private void writePartitionLineArray(ByteArrayOutputStream out, int[] parts) throws IOException {
        int byteCount = (parts.length + 3) / 4;
        for (int byteIndex = 0; byteIndex < byteCount; byteIndex++) {
            int packed = 0;
            for (int offset = 0; offset < 4; offset++) {
                int partIndex = byteIndex * 4 + offset;
                int value = partIndex < parts.length ? parts[partIndex] & 0x03 : 0;
                // Matrix partition arrays are stored low-bits first in MTEF.
                // Writing them high-bits first shifts every line boundary left by one slot in MathType.
                packed |= value << (offset * 2);
            }
            out.write(packed);
        }
    }

    private void writeUnsignedInt(ByteArrayOutputStream out, int value) throws IOException {
        if (value < 255) {
            out.write(value & 0xFF);
            return;
        }
        out.write(0xFF);
        out.write(value & 0xFF);
        out.write((value >> 8) & 0xFF);
    }

    // ======================== 辅助方法 (Helper Methods) ========================

    /**
     * 写入一个 slot（用 LINE 记录包裹内容，以 END 结束）。
     *
     * <p>MTEF 模板中的每个 slot（插槽）都是一个 LINE 记录，格式：</p>
     * <pre>
     *   LINE(0x0A) + options(0x00) + [内容记录...] + END(0x00)
     * </pre>
     *
     * <p>如果 node 为 null，会生成一个空的 LINE（有 END 但无内容），
     * 这与 writeNullLine 不同 — NULL LINE 连 END 都没有。</p>
     */
    private void writeSlot(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        out.write(MtefRecord.LINE);
        out.write(0x00); // options: 0x00 = 非空行
        if (node != null) {
            writeNode(out, node);
        }
        out.write(MtefRecord.END); // 关闭 LINE（slot 结束）
    }

    private void writeRawLineNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        out.write(MtefRecord.LINE);
        out.write(0x00);
        writeChildren(out, node);
        out.write(MtefRecord.END);
    }

    private void writeRawPileNode(ByteArrayOutputStream out, LaTeXNode node) throws IOException {
        out.write(MtefRecord.PILE);
        List<VerticalLayoutSpec.VerticalTabStop> tabStops = parseRawPileTabStops(node);
        out.write(tabStops.isEmpty() ? 0x00 : MtefRecord.OPT_LP_RULER);
        out.write(parseMetadataInt(node, VerticalLayoutNodeFactory.RAW_PILE_HALIGN, 1));
        out.write(parseMetadataInt(node, VerticalLayoutNodeFactory.RAW_PILE_VALIGN, 1));
        if (!tabStops.isEmpty()) {
            pileRulerWriter.writeRuler(out, tabStops);
        }
        writeChildren(out, node);
        out.write(MtefRecord.END);
    }

    private int parseMetadataInt(LaTeXNode node, String key, int defaultValue) {
        String value = node == null ? null : node.getMetadata(key);
        if (value == null || value.isBlank()) {
            return defaultValue;
        }
        try {
            return Integer.parseInt(value.trim());
        } catch (NumberFormatException ignored) {
            return defaultValue;
        }
    }

    private List<VerticalLayoutSpec.VerticalTabStop> parseRawPileTabStops(LaTeXNode node) {
        String encoded = node == null ? null : node.getMetadata(VerticalLayoutNodeFactory.RAW_PILE_RULER);
        if (encoded == null || encoded.isBlank()) {
            return List.of();
        }

        List<VerticalLayoutSpec.VerticalTabStop> tabStops = new ArrayList<>();
        String[] entries = encoded.split(";");
        for (int index = 0; index < entries.length; index++) {
            String entry = entries[index].trim();
            if (entry.isEmpty()) {
                continue;
            }

            String[] parts = entry.split(":");
            if (parts.length != 2) {
                continue;
            }
            try {
                int kindCode = Integer.parseInt(parts[0].trim());
                int offset = Integer.parseInt(parts[1].trim());
                tabStops.add(new VerticalLayoutSpec.VerticalTabStop(index, decodeTabStopKind(kindCode), offset));
            } catch (NumberFormatException ignored) {
                // ruler 元数据损坏时忽略该制表位，避免影响其它公式输出。
            }
        }
        return tabStops;
    }

    private VerticalLayoutSpec.TabStopKind decodeTabStopKind(int kindCode) {
        return switch (kindCode) {
            case 0 -> VerticalLayoutSpec.TabStopKind.LEFT;
            case 1 -> VerticalLayoutSpec.TabStopKind.CENTER;
            case 3 -> VerticalLayoutSpec.TabStopKind.RELATION;
            case 4 -> VerticalLayoutSpec.TabStopKind.DECIMAL;
            default -> VerticalLayoutSpec.TabStopKind.RIGHT;
        };
    }

    /**
     * 写入 NULL LINE（空 slot 占位符）。
     *
     * <p>根据 MTEF 规范：当 LINE 记录的 options 设置了 mtefOPT_LINE_NULL (0x01) 位时，
     * 表示这是一个空行，对象列表完全省略（<b>没有</b> END 记录跟随）。</p>
     *
     * <p>NULL LINE 仅占 2 字节：LINE(0x0A) + options(0x01)。
     * 常用于模板中不需要内容的 base slot（如上标模板的 base）。</p>
     *
     * <p>Write a null LINE (empty slot placeholder).
     * Per MTEF spec: LINE with mtefOPT_LINE_NULL (0x01) set,
     * object list is omitted entirely (no END record).</p>
     */
    private void writeNullLine(ByteArrayOutputStream out) throws IOException {
        out.write(MtefRecord.LINE);
        out.write(MtefRecord.OPT_LINE_NULL); // 0x01 = NULL LINE（空行，无后续 END）
    }

    /**
     * 写入一条 MTEF CHAR 记录。
     *
     * <p>CHAR 记录格式：</p>
     * <pre>
     *   CHAR(0x02) + options(1字节) + typeface(1字节) + mtcode(2字节LE) + [bits8(1字节)]
     * </pre>
     *
     * <p>对于 Symbol / Greek 字体，需要额外的 bits8 字节。
     * bits8 是该字符在 Symbol 字体中的字形位置（8 位编码），
     * 因为 Symbol 字体使用自己的字符编码（非 Unicode）。
     * MathType 原生格式要求此字段才能正确渲染希腊字母和数学符号。</p>
     *
     * <p>Write a single CHAR record in MTEF binary.
     * For Symbol/Greek typefaces, includes ENC_CHAR_8 byte with the
     * Symbol font-specific encoding (bits8), matching MathType's native format.</p>
     */
    private void writeCharRecord(ByteArrayOutputStream out, int typeface, int mtcode) throws IOException {
        out.write(MtefRecord.CHAR);

        // 判断是否需要 bits8 字段（仅 Symbol 和 Greek 字体需要）
        boolean needBits8 = (typeface == MtefRecord.FN_SYMBOL
                          || typeface == MtefRecord.FN_LC_GREEK
                          || typeface == MtefRecord.FN_UC_GREEK);
        int bits8 = needBits8 ? MtefCharMap.getSymbolBits8(mtcode) : -1;

        // options: 如果有 bits8，设置 OPT_CHAR_ENC_CHAR_8 标志位
        int options = (bits8 >= 0) ? MtefRecord.OPT_CHAR_ENC_CHAR_8 : 0;
        out.write(options);
        out.write(encodeTypeface(typeface));  // 编码后的字体索引（高位置 1）
        // mtcode: 16 位小端序字符编码
        out.write(mtcode & 0xFF);        // 低字节
        out.write((mtcode >> 8) & 0xFF); // 高字节
        // bits8: Symbol 字体的字形位置（仅在需要时写入）
        if (bits8 >= 0) {
            out.write(bits8 & 0xFF);
        }
    }

    /**
     * 写入带修饰（embellishment）的字符节点。
     *
     * <p>MTEF 中的修饰（如点 ·、帽 ^、波浪 ~ 等）通过 EMBELL 记录附加在 CHAR 记录之后。
     * CHAR 记录的 options 必须设置 OPT_CHAR_EMBELL 标志位，表示后面跟有 EMBELL 记录。</p>
     *
     * <p>生成的结构：CHAR(options=EMBELL) + typeface + mtcode + EMBELL + embellType</p>
     */
    private void writeNodeWithEmbellishment(ByteArrayOutputStream out, LaTeXNode node, int embellType) throws IOException {
        if (node.getType() == LaTeXNode.Type.CHAR) {
            String ch = node.getValue();
            MtefCharMap.CharEntry entry = MtefCharMap.lookupChar(ch.charAt(0));
            int typeface = entry != null ? entry.typeface() : MtefRecord.FN_VARIABLE;
            int mtcode = entry != null ? entry.mtcode() : ch.charAt(0);

            // 写入带 EMBELL 标志的 CHAR 记录
            out.write(MtefRecord.CHAR);
            out.write(MtefRecord.OPT_CHAR_EMBELL);    // options: 标记后面跟有修饰记录
            out.write(encodeTypeface(typeface));
            out.write(mtcode & 0xFF);
            out.write((mtcode >> 8) & 0xFF);

            // 写入 EMBELL 修饰记录
            out.write(MtefRecord.EMBELL);
            out.write(0x00);          // options: 无额外选项
            out.write(embellType);     // 修饰类型（如 EMB_1DOT = 单点）
        } else {
            // 非字符节点无法添加修饰，回退为普通写入
            writeNode(out, node);
        }
    }

    /**
     * 写入数学函数名（如 "sin", "cos", "log"）为一系列 FN_FUNCTION 字体的 CHAR 记录。
     *
     * <p>MathType 中的函数名使用 FN_FUNCTION 字体（直立体/非斜体），
     * 第一个字符需要设置 OPT_CHAR_FUNC_START 标志位，
     * 告诉 MathType 这是函数名的起始（用于正确的间距和格式处理）。</p>
     */
    private void writeFunctionName(ByteArrayOutputStream out, String name) throws IOException {
        boolean first = true;
        for (char c : name.toCharArray()) {
            out.write(MtefRecord.CHAR);
            // 第一个字符标记为函数名起始，后续字符 options=0
            int options = first ? MtefRecord.OPT_CHAR_FUNC_START : 0;
            out.write(options);
            out.write(encodeTypeface(MtefRecord.FN_FUNCTION));
            out.write(c & 0xFF);       // mtcode 低字节（ASCII 字符）
            out.write(0x00);            // mtcode 高字节（0x00，因为是 ASCII）
            first = false;
        }
    }

    /**
     * 编码字体索引（typeface）：将字体索引的高位（bit 7）置 1。
     *
     * <p>MathType 原生二进制流中，字体/样式索引通常以高位置 1 的形式存储。
     * 例如：变量字体索引 3 存储为 0x83（3 | 0x80）。
     * 这是 MTEF 格式的约定，可能用于区分字体索引和其他字节。</p>
     */
    private int encodeTypeface(int typeface) {
        return (typeface & 0x7F) | 0x80;
    }

    /**
     * 写入大算子模板头部（根据命令类型选择积分/求和/求积模板）。
     *
     * <p>注意：此方法只写入 TMPL 头部，不写入内容 slot。
     * 调用者需要在此方法之后写入内容 LINE，然后调用 writeBigOpFooter 关闭模板。</p>
     *
     * <p>MathType 7 大算子结构（通过二进制对比确认）：</p>
     * <pre>
     *   TMPL header（由本方法写入）
     *     LINE: 积分/求和内容         ← 由调用者写入
     *     SUB: 上下限 LINE            ← 由 writeBigOpFooter 写入
     *     SYM: 算子字符（∫, ∑, ∏）   ← 由 writeBigOpFooter 写入
     *   END                           ← 由 writeBigOpFooter 写入
     * </pre>
     *
     * <p>Write a big operator (sum/int/prod) with optional lower and upper limits.</p>
     */
    private void writeBigOpHeader(ByteArrayOutputStream out, String cmd, boolean hasLower, boolean hasUpper) throws IOException {
        if (BIG_OP_INT_LIKE.contains(cmd)) {
            // 积分类：∫ ∬ ∭ ∮ → TM_INTEGRAL 模板
            MtefTemplateBuilder.writeIntegralHeader(out, cmd, hasLower, hasUpper);
        } else if ("\\prod".equals(cmd)) {
            // 求积：∏ → TM_PRODUCT 模板
            MtefTemplateBuilder.writeProductHeader(out, hasLower, hasUpper);
        } else {
            // 求和及其他：∑ ∐ ⋃ ⋂ ⋁ ⋀ → TM_SUM 模板
            MtefTemplateBuilder.writeSumHeader(out, hasLower, hasUpper);
        }
    }

    /**
     * 写入大算子模板的尾部：SUB(上下限) + SYM(算子符号) + END。
     *
     * <p>SUB 记录包含上下限 slot（各为一个 LINE）；
     * SYM 记录包含算子符号的 CHAR 记录（如 ∑、∫、∏）。</p>
     */
    private void writeBigOpFooter(ByteArrayOutputStream out, String cmd, LaTeXNode lower, LaTeXNode upper) throws IOException {
        boolean hasLower = lower != null;
        boolean hasUpper = upper != null;

        // SUB 记录：包含上下限 slot
        if (hasLower || hasUpper) {
            out.write(MtefRecord.SUB); // record type 11（typesize + limit slots）
            if (hasLower) writeSlot(out, lower);  // 下限 slot
            if (hasUpper) writeSlot(out, upper);  // 上限 slot
        }

        // SYM 记录：算子符号字符（通过字符映射表查找）
        MtefCharMap.CharEntry entry = MtefCharMap.lookup(cmd);
        if (entry != null) {
            out.write(MtefRecord.SYM); // record type 13
            writeCharRecord(out, entry.typeface(), entry.mtcode());
        }

        out.write(MtefRecord.END); // 关闭大算子模板
    }

    /**
     * 判断一个 AST 节点是否为大算子命令（求和/积分/求积类）。
     */
    private boolean isBigOperator(LaTeXNode node) {
        if (node.getType() != LaTeXNode.Type.COMMAND) return false;
        String cmd = node.getValue();
        return BIG_OP_SUM_LIKE.contains(cmd) || BIG_OP_INT_LIKE.contains(cmd) || BIG_OP_LIMIT_LIKE.contains(cmd);
    }

    /**
     * 判断一个名称是否为已知的数学函数名。
     *
     * <p>已知函数名包括三角函数、对数函数、极限等，
     * 在 MathType 中使用 FN_FUNCTION 字体（直立体）显示。</p>
     */
    private boolean isFunctionName(String name) {
        return Set.of("sin", "cos", "tan", "cot", "sec", "csc",
            "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
            "log", "ln", "exp", "lim", "max", "min",
            "det", "dim", "gcd", "lcm", "mod").contains(name);
    }

    private record CompositeLongDivisionLine(String text, boolean underlined) {
        private CompositeLongDivisionLine {
            text = text == null ? "" : text;
        }
    }
}
