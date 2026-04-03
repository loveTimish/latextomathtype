package com.lz.paperword.core.latex;

import com.lz.paperword.core.mathml.MathIRConverter;
import com.lz.paperword.core.mathml.MathIRNode;
import com.lz.paperword.core.latex.LaTeXTokenizer.Token;
import com.lz.paperword.core.latex.LaTeXTokenizer.TokenType;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * LaTeX 解析器：从 HTML 内容中提取 LaTeX 数学公式并构建抽象语法树（AST）。
 *
 * <p>本类是 LaTeX 解析流水线的核心组件，负责两个主要任务：</p>
 * <ol>
 *   <li><b>内容分割</b>：将包含 LaTeX 公式的文本拆分为纯文本段和数学公式段</li>
 *   <li><b>AST 构建</b>：对每个数学公式段，使用递归下降解析器（Recursive Descent Parser）
 *       将 Token 序列转换为 {@link LaTeXNode} 构成的 AST</li>
 * </ol>
 *
 * <h3>整体处理流水线：</h3>
 * <pre>
 * HTML 输入（如 "&lt;p&gt;已知 $\frac{x}{2}=3$&lt;/p&gt;"）
 *     ↓ parseHtml()：用 Jsoup 去除 HTML 标签，提取纯文本
 * 纯文本（如 "已知 $\frac{x}{2}=3$"）
 *     ↓ parseText()：用正则表达式识别 $...$ 和 $$...$$ 分隔符
 * 内容段列表 [纯文本("已知"), 数学公式("\frac{x}{2}=3")]
 *     ↓ parseLaTeX()：对数学公式段进行 Token 化 + AST 构建
 * ContentSegment 列表 [text段, math段(含AST)]
 * </pre>
 *
 * <h3>LaTeX 公式识别策略：</h3>
 * <ul>
 *   <li><b>标准分隔符</b>：$$...$$ (行间公式) 和 $...$ (行内公式)，使用联合正则确保 $$ 优先匹配</li>
 *   <li><b>替代分隔符</b>：\[...\] 和 \(...\) 会被预处理为 $$...$$ 和 $...$</li>
 *   <li><b>强约束</b>：只有被数学定界符包裹的内容才视为公式，其他裸露 LaTeX 文本一律按普通文本保留</li>
 * </ul>
 *
 * <h3>递归下降解析器设计：</h3>
 * <p>AST 构建采用自顶向下的递归下降策略，主要解析函数及其职责：</p>
 * <ul>
 *   <li>{@code parseExpression()}：解析一个表达式序列（循环调用 parseAtom + parseScripts）</li>
 *   <li>{@code parseAtom()}：解析一个原子元素（CHAR、COMMAND 或 LBRACE 开头的分组）</li>
 *   <li>{@code parseScripts()}：处理上标（^）和下标（_）运算符，支持连续 ^/_ 运算</li>
 *   <li>{@code parseCommand()}：根据命令名分发到具体的命令解析函数</li>
 *   <li>{@code parseGroup()}：解析花括号分组 {...} 中的内容</li>
 *   <li>{@code parseRequiredGroup()}：解析必需参数（花括号分组或单个原子）</li>
 * </ul>
 *
 * <p>上标和下标运算符具有右结合性，且可以链式使用。例如 {@code x^{2}_{i}} 会被解析为
 * 先创建 SUPERSCRIPT(x, 2)，再将其整体作为 SUBSCRIPT 的底数。</p>
 */
public class LaTeXParser {

    /**
     * 组合正则表达式：同时匹配行间公式 $$...$$ 和行内公式 $...$。
     * 使用非贪婪匹配（.+?），$$...$$ 分支在前以确保优先匹配，
     * 避免将 $$ 误识别为两个单独的 $ 分隔符。
     *
     * <p>匹配组说明：</p>
     * <ul>
     *   <li>group(1)：行间公式内容（$$...$$ 之间的部分）</li>
     *   <li>group(2)：行内公式内容（$...$ 之间的部分）</li>
     * </ul>
     */
    private static final Pattern LATEX_PATTERN = Pattern.compile("\\$\\$(.+?)\\$\\$|\\$(.+?)\\$", Pattern.DOTALL);

    /**
     * 数学函数命令集合。
     *
     * <p>这些命令在渲染时以直立体（罗马体）显示，而非斜体，
     * 包括三角函数（sin、cos、tan 等）、对数函数（log、ln）、极限（lim）等。
     * 解析时这些命令不需要花括号参数，直接作为 COMMAND 节点返回。</p>
     */
    private static final Set<String> FUNCTION_COMMANDS = Set.of(
        "\\sin", "\\cos", "\\tan", "\\cot", "\\sec", "\\csc",
        "\\arcsin", "\\arccos", "\\arctan",
        "\\sinh", "\\cosh", "\\tanh",
        "\\log", "\\ln", "\\exp", "\\lim", "\\max", "\\min",
        "\\det", "\\dim", "\\gcd"
    );

    /**
     * 内容段记录类：表示解析后的一个内容片段，可以是纯文本或数学公式。
     *
     * <p>在最终的 Word 文档生成中，纯文本段以普通文本写入，
     * 数学公式段则通过 MathType OLE 嵌入为可编辑的公式对象。</p>
     *
     * @param isMath  是否为数学公式段（true=数学公式，false=纯文本）
     * @param rawText 原始文本：纯文本段为文本内容，数学段为去掉 $ 分隔符后的 LaTeX 源码
     * @param ast     数学公式的 AST 根节点；纯文本段此字段为 null
     */
    public record ContentSegment(boolean isMath, String rawText, LaTeXNode ast) {}

    /** 词法分析器实例，用于将 LaTeX 字符串拆分为 Token 序列 */
    private final LaTeXTokenizer tokenizer = new LaTeXTokenizer();
    /** Phase 3 新增：将现有 AST 归一化为 MathML-aligned IR。 */
    private final MathIRConverter mathIRConverter = new MathIRConverter();

    /**
     * 解析 HTML 内容，提取其中的纯文本和 LaTeX 公式段。
     *
     * <p>处理步骤：</p>
     * <ol>
     *   <li>使用 Jsoup 解析 HTML，提取 body 中的纯文本内容（去除所有 HTML 标签）</li>
     *   <li>将提取的纯文本传递给 {@link #parseText(String)} 进行公式识别和解析</li>
     * </ol>
     *
     * @param html 包含 LaTeX 公式的 HTML 字符串（如题目内容 HTML）
     * @return 内容段列表，按出现顺序排列
     */
    public List<ContentSegment> parseHtml(String html) {
        if (html == null || html.isBlank()) {
            return List.of();
        }

        // 使用 Jsoup 去除 HTML 标签，保留纯文本内容
        Document doc = Jsoup.parse(html);
        String text = doc.body().text();

        return parseText(text);
    }

    /**
     * 解析可能包含 LaTeX 分隔符（$$...$$ 或 $...$）的纯文本。
     *
     * <p>处理流程：</p>
     * <ol>
     *   <li>先将 \[...\] 和 \(...\) 标准化为 $$...$$ 和 $...$</li>
     *   <li>使用联合正则逐个匹配公式段，$$...$$ 优先于 $...$</li>
     *   <li>公式段之间的文本整体按普通文本保留，不再尝试识别裸露命令</li>
     *   <li>每个识别出的公式段调用 {@link #parseLaTeX(String)} 构建 AST</li>
     * </ol>
     *
     * @param text 可能包含 LaTeX 公式分隔符的纯文本
     * @return 内容段列表（纯文本段 + 数学公式段）
     */
    public List<ContentSegment> parseText(String text) {
        List<ContentSegment> segments = new ArrayList<>();
        // 将 \[...\] 和 \(...\) 统一转换为 $$...$$ 和 $...$
        text = normalizeMathDelimiters(text);

        Matcher matcher = LATEX_PATTERN.matcher(text);
        int lastEnd = 0;       // 上一个匹配结束的位置
        boolean found = false;  // 是否找到过任何公式段

        while (matcher.find()) {
            found = true;
            // 处理公式前的纯文本部分
            if (matcher.start() > lastEnd) {
                addPlainTextSegment(text.substring(lastEnd, matcher.start()), segments);
            }
            // group(1) = 行间公式 $$...$$ 的内容, group(2) = 行内公式 $...$ 的内容
            String latex = matcher.group(1) != null ? matcher.group(1).trim() : matcher.group(2).trim();
            // 将 LaTeX 源码解析为 AST
            LaTeXNode ast = parseLaTeX(latex);
            segments.add(new ContentSegment(true, latex, ast));
            lastEnd = matcher.end();
        }

        if (!found) {
            // 未找到任何定界公式时，整段都按普通文本保留。
            addPlainTextSegment(text, segments);
        } else if (lastEnd < text.length()) {
            // 处理最后一个公式之后的剩余文本
            addPlainTextSegment(text.substring(lastEnd), segments);
        }

        return segments;
    }

    /**
     * 标准化数学分隔符：将 LaTeX 的 \[...\] 和 \(...\) 转换为 $$...$$ 和 $...$。
     * 这样后续只需要用一种正则就能统一匹配所有公式分隔符。
     *
     * @param text 原始文本
     * @return 分隔符标准化后的文本
     */
    private String normalizeMathDelimiters(String text) {
        return text
            .replace("\\[", "$$")
            .replace("\\]", "$$")
            .replace("\\(", "$")
            .replace("\\)", "$");
    }

    /**
     * 将一段纯文本添加为纯文本内容段。
     * 会对文本进行 trim 处理，空白文本不会被添加。
     *
     * @param text 纯文本字符串
     * @param out  输出的内容段列表
     */
    private void addPlainTextSegment(String text, List<ContentSegment> out) {
        if (text == null) return;
        String normalized = text.trim();
        if (!normalized.isBlank()) {
            out.add(new ContentSegment(false, normalized, null));
        }
    }

    // ==================== AST 构建（递归下降解析器） ====================

    /**
     * 将一段 LaTeX 数学表达式字符串解析为 AST。
     *
     * <p>这是 AST 构建的入口方法，处理流程：</p>
     * <ol>
     *   <li>调用 {@link LaTeXTokenizer#tokenize(String)} 将 LaTeX 字符串分词为 Token 列表</li>
     *   <li>将 Token 列表封装为 {@link TokenStream} 以支持前瞻（peek）和消费（next）操作</li>
     *   <li>创建 ROOT 类型的根节点</li>
     *   <li>调用 {@link #parseExpression(TokenStream, LaTeXNode)} 递归解析所有 Token</li>
     * </ol>
     *
     * @param latex LaTeX 数学表达式字符串（不含 $ 分隔符）
     * @return AST 根节点（类型为 ROOT）
     */
    public LaTeXNode parseLaTeX(String latex) {
        List<Token> tokens = tokenizer.tokenize(latex);
        TokenStream stream = new TokenStream(tokens);
        LaTeXNode root = new LaTeXNode(LaTeXNode.Type.ROOT);
        parseExpression(stream, root);
        return root;
    }

    /**
     * 将 LaTeX 直接解析为 MathML-aligned IR，供 Phase 3 之后的语义层和诊断使用。
     */
    public MathIRNode parseMathIR(String latex) {
        return mathIRConverter.convert(parseLaTeX(latex));
    }

    /**
     * 生成可读的 IR 树 dump，便于 Linux 上做探针和测试断言。
     */
    public String dumpMathIR(String latex) {
        return mathIRConverter.dump(parseMathIR(latex));
    }

    /**
     * 解析一个表达式序列，将解析出的节点作为子节点添加到 parent 中。
     *
     * <p>核心循环逻辑：</p>
     * <ol>
     *   <li>检查下一个 Token 是否为右花括号或右方括号（标志分组结束），若是则退出</li>
     *   <li>调用 {@link #parseAtom(TokenStream)} 解析一个原子元素</li>
     *   <li>调用 {@link #parseScripts(TokenStream, LaTeXNode)} 检查并处理后续的上标/下标</li>
     *   <li>将最终节点添加为 parent 的子节点</li>
     *   <li>重复以上步骤直到 Token 流耗尽或遇到分组结束标记</li>
     * </ol>
     *
     * @param stream Token 流
     * @param parent 父节点，解析出的元素将添加为其子节点
     */
    private void parseExpression(TokenStream stream, LaTeXNode parent) {
        while (stream.hasNext()) {
            stream.skipWhitespace();
            if (!stream.hasNext()) {
                break;
            }
            Token token = stream.peek();

            // 遇到右花括号 } 或右方括号 ]，表示当前分组/可选参数结束
            if (token.type() == TokenType.RBRACE || token.type() == TokenType.RBRACKET) {
                break;
            }

            // 解析一个原子元素（字符、命令或分组）
            LaTeXNode node = parseAtom(stream);
            if (node != null) {
                // 检查原子元素后面是否有上标 ^ 或下标 _，若有则包装为 SUPERSCRIPT/SUBSCRIPT 节点
                node = parseScripts(stream, node);
                parent.addChild(node);
            }
        }
    }

    /**
     * 解析一个原子元素（Atom）——表达式中不可再分的基本单位。
     *
     * <p>根据当前 Token 类型分派处理：</p>
     * <ul>
     *   <li>CHAR → 创建 CHAR 类型节点（单个字符）</li>
     *   <li>COMMAND → 调用 {@link #parseCommand(TokenStream, String)} 处理命令</li>
     *   <li>LBRACE → 调用 {@link #parseGroup(TokenStream)} 解析花括号分组</li>
     *   <li>其他类型（RBRACE、CARET 等）→ 返回 null（由调用者处理）</li>
     * </ul>
     *
     * @param stream Token 流
     * @return 解析出的 AST 节点，或 null（无法识别的 Token）
     */
    private LaTeXNode parseAtom(TokenStream stream) {
        stream.skipWhitespace();
        if (!stream.hasNext()) return null;
        Token token = stream.next();

        return switch (token.type()) {
            case CHAR -> new LaTeXNode(LaTeXNode.Type.CHAR, token.value());
            case COMMAND -> parseCommand(stream, token.value());
            case LBRACE -> parseGroup(stream);
            default -> null;
        };
    }

    /**
     * 根据命令名分派到对应的命令解析函数。
     *
     * <p>命令分类及处理方式：</p>
     * <ul>
     *   <li>\frac → {@link #parseFrac}：解析分数，读取两个必需参数（分子、分母）</li>
     *   <li>\sqrt → {@link #parseSqrt}：解析根号，可选的根次参数 + 必需的被开方数</li>
     *   <li>\left → {@link #parseLeftRight}：解析 \left...\right 定界符对</li>
     *   <li>\overline 等一元装饰命令 → {@link #parseUnaryCommand}：读取一个必需参数</li>
     *   <li>\text 等文本/字体命令 → {@link #parseTextCommand}：读取一个必需参数</li>
     *   <li>\sum 等大型运算符 → {@link #parseBigOp}：仅创建节点（上下标由外层处理）</li>
     *   <li>\lim → {@link #parseLimCommand}：类似大型运算符处理</li>
     *   <li>函数命令（sin、cos 等）→ {@link #parseFunctionCommand}：仅创建节点</li>
     *   <li>其他命令（希腊字母、符号等）→ 直接创建 COMMAND 节点</li>
     * </ul>
     *
     * @param stream Token 流
     * @param cmd    完整命令名（含反斜杠，如 "\frac"）
     * @return 解析出的 AST 节点
     */
    private LaTeXNode parseCommand(TokenStream stream, String cmd) {
        return switch (cmd) {
            case "\\begin" -> parseBeginEnvironment(stream);
            case "\\frac" -> parseFrac(stream);
            case "\\sqrt" -> parseSqrt(stream);
            case "\\longdiv" -> parseLongDiv(stream);
            case "\\left" -> parseLeftRight(stream);
            case "\\overline", "\\underline", "\\hat", "\\tilde",
                 "\\vec", "\\bar", "\\dot", "\\arc", "\\overarc",
                 "\\overbrace", "\\underbrace", "\\overbracket", "\\underbracket",
                 "\\boxed", "\\cancel", "\\bcancel", "\\xcancel" -> parseUnaryCommand(stream, cmd);
            case "\\text", "\\mathrm", "\\mathbf", "\\mathit",
                 "\\mathcal", "\\mathbb" -> parseTextCommand(stream, cmd);
            case "\\sum", "\\int", "\\iint", "\\iiint", "\\oint",
                 "\\prod", "\\coprod", "\\bigcup", "\\bigcap" -> parseBigOp(stream, cmd);
            case "\\lim" -> parseLimCommand(stream, cmd);
            default -> {
                if (FUNCTION_COMMANDS.contains(cmd)) {
                    yield parseFunctionCommand(stream, cmd);
                }
                // 希腊字母、符号等无参数命令，直接作为 COMMAND 节点返回
                yield new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
            }
        };
    }

    private LaTeXNode parseBeginEnvironment(TokenStream stream) {
        String envName = extractPlainText(parseRequiredGroup(stream));
        if ("array".equals(envName)) {
            return parseArrayEnvironment(stream, envName, extractPlainText(parseRequiredGroup(stream)));
        }
        if ("longdivision".equals(envName)) {
            // `longdivision` 环境暂按 array 兼容处理：
            // 目前已确认的 MTEF `tmLDIV` 只支持商槽和被除数槽，
            // 不能安全承载完整步骤矩阵；为保证 MathType 可识别，这里不将环境直接提升为自定义模板节点。
            return parseArrayEnvironment(stream, envName, extractPlainText(parseRequiredGroup(stream)));
        }
        if (isMatrixLikeEnvironment(envName)) {
            return wrapEnvironmentFence(envName, parseArrayEnvironment(stream, envName, null));
        }
        if (isAlignedLikeEnvironment(envName) || "cases".equals(envName)) {
            return parseArrayEnvironment(stream, envName, null);
        }
        return new LaTeXNode(LaTeXNode.Type.COMMAND, "\\begin{" + envName + "}");
    }

    private boolean isMatrixLikeEnvironment(String envName) {
        return Set.of("matrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix").contains(envName);
    }

    private boolean isAlignedLikeEnvironment(String envName) {
        return Set.of("aligned", "align", "align*", "split").contains(envName);
    }

    private LaTeXNode wrapEnvironmentFence(String envName, LaTeXNode content) {
        String command = switch (envName) {
            case "pmatrix" -> "\\left(";
            case "bmatrix" -> "\\left[";
            case "Bmatrix" -> "\\left\\{";
            case "vmatrix" -> "\\left|";
            case "Vmatrix" -> "\\left\\lVert";
            default -> null;
        };
        if (command == null) {
            return content;
        }
        LaTeXNode fenced = new LaTeXNode(LaTeXNode.Type.COMMAND, command);
        fenced.setMetadata("leftDelimiter", normalizeDelimiter(command.substring("\\left".length())));
        fenced.setMetadata("rightDelimiter", switch (envName) {
            case "pmatrix" -> ")";
            case "bmatrix" -> "]";
            case "Bmatrix" -> "}";
            case "vmatrix" -> "|";
            case "Vmatrix" -> "||";
            default -> null;
        });
        fenced.addChild(content);
        return fenced;
    }

    private LaTeXNode parseArrayEnvironment(TokenStream stream, String envName, String explicitColumnSpec) {
        LaTeXNode arrayNode = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\" + envName);
        arrayNode.setMetadata("environment", envName);

        List<Integer> rowLines = new ArrayList<>();
        rowLines.add(0);

        LaTeXNode currentRow = new LaTeXNode(LaTeXNode.Type.ROW);
        LaTeXNode currentCell = new LaTeXNode(LaTeXNode.Type.CELL);
        boolean seenContent = false;

        while (stream.hasNext()) {
            stream.skipWhitespace();
            if (!stream.hasNext()) {
                break;
            }
            if (stream.matchesEnvironmentEnd(envName)) {
                finalizeArrayCell(currentRow, currentCell);
                finalizeArrayRow(arrayNode, currentRow);
                stream.consumeEnvironmentEnd();
                break;
            }

            Token token = stream.peek();
            if (token.type() == TokenType.COMMAND && "\\hline".equals(token.value())) {
                stream.next();
                rowLines.set(rowLines.size() - 1, 1);
                continue;
            }
            if (token.type() == TokenType.CHAR && "&".equals(token.value())) {
                stream.next();
                finalizeArrayCell(currentRow, currentCell);
                currentCell = new LaTeXNode(LaTeXNode.Type.CELL);
                seenContent = true;
                continue;
            }
            if (token.type() == TokenType.COMMAND && "\\\\".equals(token.value())) {
                stream.next();
                finalizeArrayCell(currentRow, currentCell);
                finalizeArrayRow(arrayNode, currentRow);
                currentRow = new LaTeXNode(LaTeXNode.Type.ROW);
                currentCell = new LaTeXNode(LaTeXNode.Type.CELL);
                rowLines.add(0);
                seenContent = false;
                continue;
            }

            LaTeXNode child = parseAtom(stream);
            if (child != null) {
                child = parseScripts(stream, child);
                currentCell.addChild(child);
                seenContent = true;
                continue;
            }
        }

        String columnSpec = explicitColumnSpec;
        if (columnSpec == null || columnSpec.isBlank()) {
            int inferredColumns = resolveMaxColumns(arrayNode);
            columnSpec = defaultColumnSpecForEnvironment(envName, inferredColumns);
        }
        arrayNode.setMetadata("columnSpec", columnSpec);
        arrayNode.setMetadata("columnCount", String.valueOf(countArrayColumns(columnSpec)));
        arrayNode.setMetadata("columnLines", encodeColumnPartitionLines(columnSpec));
        if (usesRelationPairAlignment(envName)) {
            arrayNode.setMetadata("alignmentMode", "relation-pairs");
        }

        if (!arrayNode.getChildren().isEmpty() || seenContent || !currentCell.getChildren().isEmpty()) {
            arrayNode.setMetadata("rowLines", encodeRowPartitionLines(rowLines, arrayNode.getChildren().size()));
        } else {
            arrayNode.setMetadata("rowLines", "0");
        }
        return arrayNode;
    }

    private int resolveMaxColumns(LaTeXNode arrayNode) {
        int maxColumns = 0;
        for (LaTeXNode row : arrayNode.getChildren()) {
            maxColumns = Math.max(maxColumns, row.getChildren().size());
        }
        return Math.max(maxColumns, 1);
    }

    private String defaultColumnSpecForEnvironment(String envName, int columns) {
        int safeColumns = Math.max(columns, 1);
        if (isMatrixLikeEnvironment(envName)) {
            return "c".repeat(safeColumns);
        }
        if (isAlignedLikeEnvironment(envName)) {
            return buildAlternatingColumnSpec(safeColumns, 'r', 'l');
        }
        if ("cases".equals(envName)) {
            StringBuilder spec = new StringBuilder(safeColumns);
            for (int i = 0; i < safeColumns; i++) {
                spec.append('l');
            }
            return spec.toString();
        }
        return "c".repeat(safeColumns);
    }

    private boolean usesRelationPairAlignment(String envName) {
        return Set.of("aligned", "align", "align*", "split").contains(envName);
    }

    private String buildAlternatingColumnSpec(int columns, char even, char odd) {
        StringBuilder spec = new StringBuilder(Math.max(columns, 1));
        for (int i = 0; i < Math.max(columns, 1); i++) {
            spec.append(i % 2 == 0 ? even : odd);
        }
        return spec.toString();
    }

    private void finalizeArrayCell(LaTeXNode row, LaTeXNode cell) {
        if (isExplicitEmptyCell(cell)) {
            // 显式写出的 {} 需要保留“占位但无内容”的语义，后续竖式对齐会用到。
            cell.setMetadata("explicitEmptyCell", "true");
        }
        row.addChild(cell);
    }

    private boolean isExplicitEmptyCell(LaTeXNode cell) {
        if (cell == null || cell.getChildren().size() != 1) {
            return false;
        }
        LaTeXNode child = cell.getChildren().get(0);
        return child.getType() == LaTeXNode.Type.GROUP && child.getChildren().isEmpty();
    }

    private void finalizeArrayRow(LaTeXNode arrayNode, LaTeXNode row) {
        if (row.getChildren().isEmpty()) {
            return;
        }
        arrayNode.addChild(row);
    }

    private int countArrayColumns(String columnSpec) {
        if (columnSpec == null || columnSpec.isBlank()) {
            return 0;
        }
        int count = 0;
        for (int i = 0; i < columnSpec.length(); i++) {
            char ch = columnSpec.charAt(i);
            if (ch == 'l' || ch == 'c' || ch == 'r') {
                count++;
            }
        }
        return count;
    }

    private String encodeColumnPartitionLines(String columnSpec) {
        int columns = countArrayColumns(columnSpec);
        int[] parts = new int[columns + 1];
        int boundary = 0;
        for (int i = 0; i < columnSpec.length(); i++) {
            char ch = columnSpec.charAt(i);
            if (ch == '|') {
                parts[boundary] = 1;
            } else if (ch == 'l' || ch == 'c' || ch == 'r') {
                boundary++;
            }
        }
        return encodePartitionArray(parts);
    }

    private String encodeRowPartitionLines(List<Integer> rowLines, int rowCount) {
        int[] parts = new int[Math.max(rowCount + 1, 1)];
        for (int i = 0; i < parts.length && i < rowLines.size(); i++) {
            parts[i] = rowLines.get(i);
        }
        return encodePartitionArray(parts);
    }

    private String encodePartitionArray(int[] parts) {
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) {
                builder.append(',');
            }
            builder.append(parts[i]);
        }
        return builder.toString();
    }

    private String extractPlainText(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR) {
            return node.getValue() == null ? "" : node.getValue();
        }
        if (node.getType() == LaTeXNode.Type.COMMAND) {
            return node.getValue() == null ? "" : node.getValue().replace("\\", "");
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(extractPlainText(child));
        }
        return builder.toString();
    }

    /**
     * 解析分数命令 \frac{分子}{分母}。
     *
     * <p>创建 FRACTION 类型节点，依次读取两个必需的花括号参数：</p>
     * <ul>
     *   <li>children[0]：分子（numerator）</li>
     *   <li>children[1]：分母（denominator）</li>
     * </ul>
     *
     * @param stream Token 流（当前位置在 \frac 之后）
     * @return FRACTION 类型的 AST 节点
     */
    private LaTeXNode parseFrac(TokenStream stream) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.FRACTION, "\\frac");
        // 读取分子参数
        node.addChild(parseRequiredGroup(stream));
        // 读取分母参数
        node.addChild(parseRequiredGroup(stream));
        return node;
    }

    /**
     * 解析根号命令 \sqrt{...} 或 \sqrt[n]{...}。
     *
     * <p>处理逻辑：</p>
     * <ol>
     *   <li>检查下一个 Token 是否为左方括号 '['，若是则解析可选的根次参数</li>
     *   <li>解析必需的花括号参数（被开方数）</li>
     * </ol>
     *
     * <p>子节点结构：</p>
     * <ul>
     *   <li>无根次：children[0] = 被开方数（GROUP）</li>
     *   <li>有根次：children[0] = 根次（GROUP），children[1] = 被开方数（GROUP）</li>
     * </ul>
     *
     * @param stream Token 流（当前位置在 \sqrt 之后）
     * @return SQRT 类型的 AST 节点
     */
    private LaTeXNode parseSqrt(TokenStream stream) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.SQRT, "\\sqrt");
        // 检查可选参数 [n]（第 n 次方根）
        if (stream.hasNext() && stream.peek().type() == TokenType.LBRACKET) {
            stream.next(); // 消费 [
            // 解析方括号内的根次内容
            LaTeXNode degree = new LaTeXNode(LaTeXNode.Type.GROUP);
            while (stream.hasNext()) {
                stream.skipWhitespace();
                if (!stream.hasNext() || stream.peek().type() == TokenType.RBRACKET) {
                    break;
                }
                LaTeXNode child = parseAtom(stream);
                if (child != null) degree.addChild(child);
            }
            if (stream.hasNext()) stream.next(); // 消费 ]
            node.addChild(degree); // 第一个子节点 = 根次
        }
        // 读取必需的被开方数参数 {content}
        node.addChild(parseRequiredGroup(stream));
        return node;
    }

    /**
     * 解析长除法命令 \longdiv[quotient]{divisor}{dividend}。
     *
     * <p>创建 LONG_DIVISION 类型节点。MathType 的 LDivBoxClass 只有两个模板槽位：
     * quotient slot 和 dividend slot；除数应写在模板外部。
     * 因此本节点的子节点顺序约定为：</p>
     * <ul>
     *   <li>children[0]：除数（divisor，模板外部）</li>
     *   <li>children[1]：商（quotient，可选）</li>
     *   <li>children[2]：被除数（dividend，模板内主槽位）</li>
     * </ul>
     *
     * @param stream Token 流（当前位置在 \longdiv 之后）
     * @return LONG_DIVISION 类型的 AST 节点
     */
    private LaTeXNode parseLongDiv(TokenStream stream) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.LONG_DIVISION, "\\longdiv");
        // 不再强制使用简单模板 - 使用 computed array 路径以支持结构化步骤显示
        if (stream.hasNext() && stream.peek().type() == TokenType.LBRACKET) {
            stream.next();
            LaTeXNode quotient = new LaTeXNode(LaTeXNode.Type.GROUP);
            while (stream.hasNext()) {
                stream.skipWhitespace();
                if (!stream.hasNext() || stream.peek().type() == TokenType.RBRACKET) {
                    break;
                }
                LaTeXNode child = parseAtom(stream);
                if (child != null) {
                    child = parseScripts(stream, child);
                    quotient.addChild(child);
                }
            }
            if (stream.hasNext()) {
                stream.next();
            }
            node.addChild(parseRequiredGroup(stream)); // divisor
            node.addChild(quotient);                  // quotient
            node.addChild(parseRequiredGroup(stream)); // dividend
            return node;
        }
        node.addChild(parseRequiredGroup(stream));     // divisor
        node.addChild(new LaTeXNode(LaTeXNode.Type.GROUP)); // no quotient slot
        node.addChild(parseRequiredGroup(stream));     // dividend
        return node;
    }

    /**
     * 解析 \left...\right 定界符对。
     *
     * <p>处理 LaTeX 中的自适应大小定界符，如 \left( ... \right)、\left\{ ... \right\}。</p>
     *
     * <p>处理逻辑：</p>
     * <ol>
     *   <li>读取左定界符字符（如 '('、'\{'）</li>
     *   <li>将节点值设置为 "\left" + 定界符（如 "\left("）</li>
     *   <li>解析定界符之间的内容，直到遇到 \right 命令</li>
     *   <li>消费 \right 及其后的右定界符</li>
     * </ol>
     *
     * @param stream Token 流（当前位置在 \left 之后）
     * @return COMMAND 类型的 AST 节点，value 为 "\left" + 左定界符
     */
    private LaTeXNode parseLeftRight(TokenStream stream) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\left");
        // 读取左定界符字符
        String leftDelim = "(";
        if (stream.hasNext()) {
            Token delim = stream.next();
            leftDelim = normalizeDelimiter(delim.value());
        }
        node.setValue("\\left" + leftDelim);
        node.setMetadata("leftDelimiter", leftDelim);

        // 解析定界符之间的内容，直到遇到 \right 命令
        LaTeXNode content = new LaTeXNode(LaTeXNode.Type.GROUP);
        while (stream.hasNext()) {
            stream.skipWhitespace();
            if (!stream.hasNext()) {
                break;
            }
            Token t = stream.peek();
            if (t.type() == TokenType.COMMAND && t.value().equals("\\right")) {
                stream.next(); // 消费 \right
                if (stream.hasNext()) {
                    node.setMetadata("rightDelimiter", normalizeDelimiter(stream.next().value()));
                }
                break;
            }
            LaTeXNode child = parseAtom(stream);
            if (child != null) {
                // 定界符内的元素也可能有上下标
                child = parseScripts(stream, child);
                content.addChild(child);
            }
        }
        node.addChild(content);
        return node;
    }

    private String normalizeDelimiter(String raw) {
        if (raw == null || raw.isBlank()) {
            return "(";
        }
        return switch (raw) {
            case "\\{", "\\lbrace" -> "{";
            case "\\}", "\\rbrace" -> "}";
            case "\\lvert", "\\rvert" -> "|";
            case "\\lVert", "\\rVert", "\\Vert" -> "||";
            case "\\lfloor" -> "⌊";
            case "\\rfloor" -> "⌋";
            case "\\lceil" -> "⌈";
            case "\\rceil" -> "⌉";
            default -> raw;
        };
    }

    /**
     * 解析一元装饰命令（如 \overline{AB}、\hat{x}、\vec{v}）。
     *
     * <p>这些命令接受一个必需的花括号参数，表示被装饰的内容。
     * 创建 COMMAND 类型节点，子节点为参数内容。</p>
     *
     * @param stream Token 流
     * @param cmd    命令名（如 "\overline"）
     * @return COMMAND 类型节点，children[0] 为被装饰的内容
     */
    private LaTeXNode parseUnaryCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        node.addChild(parseRequiredGroup(stream));
        return node;
    }

    /**
     * 解析文本/字体命令（如 \text{内容}、\mathrm{ABC}、\mathbb{R}）。
     *
     * <p>创建 TEXT 类型节点（区别于 COMMAND），表示非数学的文本内容。
     * value 存储命令名，子节点为花括号内的文本内容。</p>
     *
     * @param stream Token 流
     * @param cmd    命令名（如 "\text"、"\mathrm"）
     * @return TEXT 类型节点，children[0] 为文本内容
     */
    private LaTeXNode parseTextCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.TEXT, cmd);
        node.addChild(parseTextRequiredGroup(stream));
        return node;
    }

    private LaTeXNode parseTextRequiredGroup(TokenStream stream) {
        if (!stream.hasNext()) {
            return new LaTeXNode(LaTeXNode.Type.GROUP);
        }
        Token token = stream.peek();
        if (token.type() == TokenType.LBRACE) {
            stream.next();
            return parseTextGroup(stream);
        }
        if (token.type() == TokenType.CHAR || token.type() == TokenType.WHITESPACE) {
            stream.next();
            return new LaTeXNode(LaTeXNode.Type.CHAR, token.value());
        }
        return parseRequiredGroup(stream);
    }

    /**
     * 文本命令组需要显式保留空格，因此不能复用默认的数学分组解析。
     */
    private LaTeXNode parseTextGroup(TokenStream stream) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        while (stream.hasNext() && stream.peek().type() != TokenType.RBRACE) {
            Token token = stream.next();
            switch (token.type()) {
                case CHAR, WHITESPACE -> appendLiteralText(group, token.value());
                case COMMAND -> appendLiteralText(group, decodeTextCommandLiteral(token.value()));
                case LBRACE -> appendTextGroupChildren(group, parseTextGroup(stream));
                default -> {
                }
            }
        }
        if (stream.hasNext()) {
            stream.next();
        }
        return group;
    }

    private void appendTextGroupChildren(LaTeXNode target, LaTeXNode source) {
        if (target == null || source == null) {
            return;
        }
        for (LaTeXNode child : source.getChildren()) {
            target.addChild(child);
        }
    }

    private void appendLiteralText(LaTeXNode group, String text) {
        if (group == null || text == null || text.isEmpty()) {
            return;
        }
        group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, text));
    }

    private String decodeTextCommandLiteral(String command) {
        if (command == null || command.isEmpty()) {
            return "";
        }
        if (command.length() == 2 && command.charAt(0) == '\\' && !Character.isLetter(command.charAt(1))) {
            return command.substring(1);
        }
        return command;
    }

    /**
     * 解析大型运算符命令（如 \sum、\int、\prod、\bigcup）。
     *
     * <p>大型运算符本身不消费参数——它们的上下标（求和范围、积分界限等）
     * 由外层的 {@link #parseScripts(TokenStream, LaTeXNode)} 统一处理。
     * 因此此方法仅创建一个 COMMAND 节点并返回。</p>
     *
     * @param stream Token 流
     * @param cmd    命令名（如 "\sum"）
     * @return COMMAND 类型节点
     */
    private LaTeXNode parseBigOp(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        return node;
    }

    /**
     * 解析 \lim 命令。
     *
     * <p>与大型运算符类似，\lim 的下标（如 \lim_{x \to 0}）
     * 由外层 parseScripts 处理，此处仅创建节点。</p>
     *
     * @param stream Token 流
     * @param cmd    命令名（"\lim"）
     * @return COMMAND 类型节点
     */
    private LaTeXNode parseLimCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        return node;
    }

    /**
     * 解析数学函数命令（如 \sin、\cos、\log）。
     *
     * <p>函数命令在数学排版中以直立体显示（非斜体）。
     * 它们不消费花括号参数，参数通过自然的 LaTeX 排列传递。
     * 此处仅创建 COMMAND 节点。</p>
     *
     * @param stream Token 流
     * @param cmd    命令名（如 "\sin"）
     * @return COMMAND 类型节点
     */
    private LaTeXNode parseFunctionCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        return node;
    }

    /**
     * 处理上标（^）和下标（_）运算符。
     *
     * <p>上标和下标是后缀运算符，作用于前面的原子元素（base）。
     * 本方法会循环检查后续 Token，支持连续的上下标（如 x^{2}_{i} 或 a_{n}^{k}）。</p>
     *
     * <p>处理逻辑：</p>
     * <ul>
     *   <li>遇到 ^ → 创建 SUPERSCRIPT 节点：children[0]=底数(base)，children[1]=指数内容</li>
     *   <li>遇到 _ → 创建 SUBSCRIPT 节点：children[0]=底数(base)，children[1]=下标内容</li>
     *   <li>新创建的节点替代原 base 继续参与后续上下标检测（实现链式上下标）</li>
     * </ul>
     *
     * <p>示例：{@code x^{2}_{i}} 的解析过程：</p>
     * <ol>
     *   <li>base = CHAR(x)</li>
     *   <li>遇到 ^，创建 SUPERSCRIPT(CHAR(x), GROUP(CHAR(2)))，base 更新为此节点</li>
     *   <li>遇到 _，创建 SUBSCRIPT(SUPERSCRIPT(...), GROUP(CHAR(i)))，base 更新为此节点</li>
     * </ol>
     *
     * @param stream Token 流
     * @param base   作为底数的节点（上标/下标将作用于此节点）
     * @return 处理完上下标后的最终节点（可能是原 base，也可能是包装后的 SUPERSCRIPT/SUBSCRIPT）
     */
    private LaTeXNode parseScripts(TokenStream stream, LaTeXNode base) {
        while (stream.hasNext()) {
            Token t = stream.peek();
            if (t.type() == TokenType.CARET) {
                // 上标运算符 ^
                stream.next(); // 消费 ^
                LaTeXNode sup = new LaTeXNode(LaTeXNode.Type.SUPERSCRIPT, "^");
                sup.addChild(base);                    // children[0] = 底数
                sup.addChild(parseRequiredGroup(stream)); // children[1] = 指数内容
                base = sup; // 更新 base，支持链式上下标
            } else if (t.type() == TokenType.UNDERSCORE) {
                // 下标运算符 _
                stream.next(); // 消费 _
                LaTeXNode sub = new LaTeXNode(LaTeXNode.Type.SUBSCRIPT, "_");
                sub.addChild(base);                    // children[0] = 底数
                sub.addChild(parseRequiredGroup(stream)); // children[1] = 下标内容
                base = sub; // 更新 base，支持链式上下标
            } else {
                // 既不是上标也不是下标，退出循环
                break;
            }
        }
        return base;
    }

    /**
     * 解析一个必需的参数组。
     *
     * <p>LaTeX 中很多命令要求花括号参数（如 \frac{...}{...}），但也支持单个原子作为参数
     * （如 \frac xy 等价于 \frac{x}{y}）。本方法处理两种情况：</p>
     * <ul>
     *   <li>下一个 Token 是 LBRACE → 消费 { 并调用 {@link #parseGroup(TokenStream)} 解析完整分组</li>
     *   <li>下一个 Token 不是 LBRACE → 调用 {@link #parseAtom(TokenStream)} 解析单个原子元素</li>
     *   <li>Token 流已耗尽 → 返回空 GROUP 节点（容错处理）</li>
     * </ul>
     *
     * @param stream Token 流
     * @return 参数内容的 AST 节点（通常是 GROUP 或单个原子节点）
     */
    private LaTeXNode parseRequiredGroup(TokenStream stream) {
        stream.skipWhitespace();
        if (!stream.hasNext()) {
            // Token 流已耗尽，返回空分组作为容错
            return new LaTeXNode(LaTeXNode.Type.GROUP);
        }
        Token token = stream.peek();
        if (token.type() == TokenType.LBRACE) {
            stream.next(); // 消费 {
            return parseGroup(stream);
        }
        // 非花括号参数：解析单个原子元素作为参数
        LaTeXNode atom = parseAtom(stream);
        if (atom == null) {
            return new LaTeXNode(LaTeXNode.Type.GROUP);
        }
        return atom;
    }

    /**
     * 解析花括号分组 {...} 的内容。
     *
     * <p>调用此方法时，左花括号 { 已经被消费。方法会持续解析内容直到遇到
     * 右花括号 } 或 Token 流耗尽，然后消费右花括号。</p>
     *
     * <p>分组内的每个原子元素都会检查后续上下标，确保 {x^{2}} 中的上标被正确解析。</p>
     *
     * @param stream Token 流（当前位置在 { 之后）
     * @return GROUP 类型的 AST 节点，子节点为分组内的所有元素
     */
    private LaTeXNode parseGroup(TokenStream stream) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        // 持续解析直到遇到右花括号 }
        while (stream.hasNext()) {
            stream.skipWhitespace();
            if (!stream.hasNext() || stream.peek().type() == TokenType.RBRACE) {
                break;
            }
            LaTeXNode child = parseAtom(stream);
            if (child != null) {
                // 分组内的元素也可能有上下标
                child = parseScripts(stream, child);
                group.addChild(child);
            }
        }
        // 消费右花括号 }（如果存在）
        if (stream.hasNext()) {
            stream.next();
        }
        return group;
    }

    /**
     * Token 流封装类：为 Token 列表提供顺序访问接口。
     *
     * <p>支持三个基本操作：</p>
     * <ul>
     *   <li>{@link #hasNext()}：检查是否还有未消费的 Token</li>
     *   <li>{@link #peek()}：前瞻（查看）下一个 Token，不消费</li>
     *   <li>{@link #next()}：消费并返回下一个 Token，指针后移</li>
     * </ul>
     *
     * <p>递归下降解析器通过 peek() 进行前瞻判断，通过 next() 消费 Token，
     * 这种模式使得解析器可以根据下一个 Token 的类型决定进入哪个解析分支。</p>
     */
    private static class TokenStream {
        /** Token 列表（不可变引用） */
        private final List<Token> tokens;
        /** 当前读取位置（指向下一个待消费的 Token） */
        private int pos = 0;

        /**
         * 构造 Token 流。
         *
         * @param tokens 词法分析器输出的 Token 列表
         */
        TokenStream(List<Token> tokens) {
            this.tokens = tokens;
        }

        /**
         * 检查是否还有未消费的 Token。
         *
         * @return 如果还有 Token 未被消费则返回 true
         */
        boolean hasNext() {
            return pos < tokens.size();
        }

        /**
         * 前瞻：查看下一个 Token，不移动指针。
         * 调用前应先通过 {@link #hasNext()} 确认还有 Token 可用。
         *
         * @return 下一个 Token
         */
        Token peek() {
            return tokens.get(pos);
        }

        /**
         * 消费并返回下一个 Token，指针后移一位。
         * 调用前应先通过 {@link #hasNext()} 确认还有 Token 可用。
         *
         * @return 当前位置的 Token
         */
        Token next() {
            return tokens.get(pos++);
        }

        void skipWhitespace() {
            while (hasNext() && peek().type() == TokenType.WHITESPACE) {
                pos++;
            }
        }

        boolean matchesEnvironmentEnd(String envName) {
            if (!hasNext() || peek().type() != TokenType.COMMAND || !"\\end".equals(peek().value())) {
                return false;
            }
            int idx = pos + 1;
            if (idx >= tokens.size() || tokens.get(idx).type() != TokenType.LBRACE) {
                return false;
            }
            idx++;
            StringBuilder builder = new StringBuilder();
            while (idx < tokens.size() && tokens.get(idx).type() != TokenType.RBRACE) {
                builder.append(tokens.get(idx).value());
                idx++;
            }
            return idx < tokens.size() && envName.equals(builder.toString());
        }

        void consumeEnvironmentEnd() {
            if (!hasNext()) {
                return;
            }
            next(); // \end
            if (!hasNext() || peek().type() != TokenType.LBRACE) {
                return;
            }
            next(); // {
            while (hasNext() && peek().type() != TokenType.RBRACE) {
                next();
            }
            if (hasNext()) {
                next(); // }
            }
        }
    }
}
