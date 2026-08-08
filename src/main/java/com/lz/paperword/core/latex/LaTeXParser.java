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
import java.util.LinkedHashSet;
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

    public enum DiagnosticSeverity {
        WARNING,
        ERROR
    }

    public record ParseDiagnostic(DiagnosticSeverity severity, String code, String command, String message) {
    }

    public record DetailedParseResult(
        String sourceLatex,
        String normalizedLatex,
        LaTeXNode ast,
        MathIRNode mathIR,
        List<String> consumedCommands,
        List<ParseDiagnostic> diagnostics
    ) {
        public DetailedParseResult {
            consumedCommands = List.copyOf(consumedCommands);
            diagnostics = List.copyOf(diagnostics);
        }

        public boolean isSupported() {
            return diagnostics.stream().noneMatch(diagnostic -> diagnostic.severity() == DiagnosticSeverity.ERROR);
        }
    }

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
    private static final Pattern METRICS_PATTERN = Pattern.compile(
        "^\\\\pwmetrics\\{([0-9]+(?:\\.[0-9]+)?)\\s*,\\s*([0-9]+(?:\\.[0-9]+)?)(?:\\s*,\\s*([0-9]+(?:\\.[0-9]+)?)\\s*,\\s*([0-9]+(?:\\.[0-9]+)?))?\\}\\s*",
        Pattern.DOTALL);
    private static final Pattern LOOSE_METRICS_PATTERN = Pattern.compile(
        "^\\\\pwmetrics\\s+([0-9]+(?:\\.[0-9]+)?)\\s*,\\s*([0-9]+(?:\\.[0-9]+)?)(?:\\s*,\\s*([0-9]+(?:\\.[0-9]+)?)\\s*,\\s*([0-9]+(?:\\.[0-9]+)?))?\\s+",
        Pattern.DOTALL);
    private static final Pattern STYLE_HINT_PATTERN = Pattern.compile(
        "^\\\\pwstyle\\{([^}]*)\\}\\s*",
        Pattern.DOTALL);
    private static final Pattern STYLE_WRAPPED_ALIGNMENT_MARKER = Pattern.compile(
        "(\\\\(?:mathbf|mathrm|mathit|mathsf|mathtt|mathcal|mathbb|boldsymbol)\\s*\\{\\s*)&");
    private static final Pattern NESTED_TEXT_COLOR_MATH = Pattern.compile(
        "(\\\\textcolor\\{[^{}]+}\\{)\\$([^$]*)\\$");
    private static final Pattern SHARED_DELIMITER_BEFORE_BARE_STRUCTURE = Pattern.compile(
        "}\\$(?=\\\\(?:frac|dfrac|tfrac|cfrac|sqrt|binom)\\b)");
    private static final Pattern NON_CONTENT_INCLUDE_GRAPHICS = Pattern.compile(
        "\\\\includegraphics(?:\\[[^]]*])?\\s*(?:\\{[^}]+}|\\S+)",
        Pattern.CASE_INSENSITIVE);
    private static final Pattern TEXT_TABLE_PREAMBLE = Pattern.compile(
        "\\\\begin\\s*\\{?table}?\\s+\\\\begin\\s*\\{?tabularx}?.*\\\\arrayrulewidth\\s*\\|",
        Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
    private static final Pattern TEXT_TABLE_EMPTY_FIRST_PREAMBLE = Pattern.compile(
        "\\\\begin\\s*\\{?table}?\\s+\\\\begin\\s*\\{?tabularx}?.*?\\|\\s*&",
        Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
    private static final Pattern TEXT_TABLE_END = Pattern.compile(
        "\\\\end\\s*\\{?(?:tabularx|tabular|table)}?",
        Pattern.CASE_INSENSITIVE);
    private static final Pattern LEGACY_SPEED_RATIO_LABELS = Pattern.compile(
        "V_\\{\\s*\uFFFD\\s*}\\s*\\$\\s*[：:]\\s*\\$"
            + "[^$]*V_\\{\\s*\uFFFD\\s*}\\s*\\$\\s*=\\s*1\\s*[：:]\\s*12\\s*[,，]\\s*\\$"
            + "[^$]*V_\\{\\s*\uFFFD\\s*}\\s*\\$\\s*[：:]\\s*\\$"
            + "[^$]*V_\\{\\s*\uFFFD\\s*}\\s*\\$\\s*=\\s*1\\s*[：:]\\s*16",
        Pattern.DOTALL);

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
    public record ContentSegment(boolean isMath, String rawText, LaTeXNode ast, FormulaMetrics metrics,
                                 FormulaStyleHints styleHints) {
        public ContentSegment(boolean isMath, String rawText, LaTeXNode ast) {
            this(isMath, rawText, ast, null, FormulaStyleHints.empty());
        }
    }

    public record FormulaMetrics(double wmfWidthPt, double wmfHeightPt,
                                 double shapeWidthPt, double shapeHeightPt) {
        public FormulaMetrics(double widthPt, double heightPt) {
            this(widthPt, heightPt, widthPt, heightPt);
        }

        public double widthPt() {
            return wmfWidthPt;
        }

        public double heightPt() {
            return wmfHeightPt;
        }
    }
    public record FormulaStyleHints(boolean asciiFlatParens, boolean explicitScriptFullSize,
                                    boolean explicitFractionFullSize, boolean explicitTopFullSize,
                                    boolean forceExplicitFenceTemplate, boolean explicitBlackColor,
                                    boolean flatParenTemplate, boolean letterGroupObarTemplate,
                                    boolean textFeComma, boolean fullwidthTextParen,
                                    boolean mixedAsciiFullwidthParens,
                                    boolean legacyTextFeParenContent,
                                    FormulaMetrics sourceMetrics) {
        public FormulaStyleHints(boolean asciiFlatParens, boolean explicitScriptFullSize,
                                 boolean explicitFractionFullSize, boolean explicitTopFullSize,
                                 boolean forceExplicitFenceTemplate, boolean explicitBlackColor,
                                 boolean flatParenTemplate, boolean letterGroupObarTemplate,
                                 boolean textFeComma, boolean fullwidthTextParen,
                                 FormulaMetrics sourceMetrics) {
            this(asciiFlatParens, explicitScriptFullSize, explicitFractionFullSize, explicitTopFullSize,
                forceExplicitFenceTemplate, explicitBlackColor, flatParenTemplate, letterGroupObarTemplate,
                textFeComma, fullwidthTextParen, false, false, sourceMetrics);
        }

        public static FormulaStyleHints empty() {
            return new FormulaStyleHints(false, false, false, false, false, false, false, false, false, false,
                false, false, null);
        }

        public FormulaStyleHints withSourceMetrics(FormulaMetrics metrics) {
            if (metrics == null) {
                return this;
            }
            return new FormulaStyleHints(asciiFlatParens, explicitScriptFullSize, explicitFractionFullSize,
                explicitTopFullSize, forceExplicitFenceTemplate, explicitBlackColor, flatParenTemplate,
                letterGroupObarTemplate, textFeComma, fullwidthTextParen, mixedAsciiFullwidthParens,
                legacyTextFeParenContent, metrics);
        }
    }

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
        text = NON_CONTENT_INCLUDE_GRAPHICS.matcher(text).replaceAll("");
        text = recoverLegacySpeedRatioLabels(text);
        text = recoverKnownReplacementArtifactsInText(text);
        text = normalizePlainTextTables(text);
        // 将 \[...\] 和 \(...\) 统一转换为 $$...$$ 和 $...$
        text = normalizeMathDelimiters(text);
        text = flattenNestedTextColorMath(text);
        text = splitSharedMathDelimiterBeforeBareStructure(text);

        int lastEnd = 0;       // 上一个匹配结束的位置
        boolean found = false;  // 是否找到过任何公式段
        int searchFrom = 0;

        DelimitedFormula matched;
        while ((matched = findNextDelimitedFormula(text, searchFrom)) != null) {
            found = true;
            // 处理公式前的纯文本部分
            if (matched.start() > lastEnd) {
                addPlainTextSegment(text.substring(lastEnd, matched.start()), segments);
            }
            String latex = matched.content().trim();
            latex = latex.replaceFirst("^\\$\\s*(?=\\\\pwmetrics\\b)", "");
            ParsedFormulaMetrics parsedMetrics = stripFormulaMetrics(latex);
            latex = stripEmbeddedFormulaAnnotations(parsedMetrics.latex());
            ParsedFormulaStyle parsedStyle = stripFormulaStyle(latex);
            latex = parsedStyle.latex();
            DetailedParseResult detailed = parseDetailed(latex);
            if (!detailed.isSupported()) {
                throw new IllegalArgumentException("Unsupported LaTeX formula: " + latex
                    + "; diagnostics=" + detailed.diagnostics());
            }
            segments.add(new ContentSegment(true, latex, detailed.ast(),
                parsedMetrics.metrics(), parsedStyle.styleHints()));
            lastEnd = matched.endExclusive();
            searchFrom = lastEnd;
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
     * Recovers the four Far East labels whose source MTEF only retained U+FFFD plus a
     * legacy font code. The visible source Word object establishes the sequence as
     * V_甲:V_车=1:12 and V_乙:V_车=1:16. Keeping this repair at text level is important:
     * once the formulas are split, the four identical replacement characters no longer
     * carry enough context to recover different labels safely.
     */
    private String recoverLegacySpeedRatioLabels(String text) {
        if (text == null || text.indexOf('\uFFFD') < 0) {
            return text;
        }
        Matcher matcher = LEGACY_SPEED_RATIO_LABELS.matcher(text);
        StringBuffer out = new StringBuffer(text.length());
        String[] labels = {"\\text{甲}", "\\text{车}", "\\text{乙}", "\\text{车}"};
        while (matcher.find()) {
            String fragment = matcher.group();
            StringBuilder repaired = new StringBuilder(fragment.length() + 24);
            int label = 0;
            for (int index = 0; index < fragment.length(); index++) {
                char ch = fragment.charAt(index);
                if (ch == '\uFFFD' && label < labels.length) {
                    repaired.append(labels[label++]);
                } else {
                    repaired.append(ch);
                }
            }
            if (label != labels.length) {
                throw new IllegalStateException("legacy speed-ratio label recovery count mismatch");
            }
            matcher.appendReplacement(out, Matcher.quoteReplacement(repaired.toString()));
        }
        matcher.appendTail(out);
        return out.toString();
    }

    /**
     * Applies only replacements whose missing source glyph can be recovered from the
     * surrounding prose. This runs before formulas are split so parser entry points
     * never have to guess what an isolated U+FFFD represented.
     */
    private String recoverKnownReplacementArtifactsInText(String text) {
        if (text == null || text.indexOf('\uFFFD') < 0) {
            return text;
        }
        String repaired = text.replaceAll("\uFFFD\\s*路程", "总路程");

        // These are fixed extraction artifacts in the versioned xsc corpus. Each
        // replacement is anchored by the complete arithmetic/prose context so an
        // unrelated U+FFFD can never be guessed or silently discarded.
        repaired = repaired
            .replace("(75+60)\uFFFD \\times 20=\uFFFD 2700", "(75+60) \\times 20=2700")
            .replace("80-75=\uFFFD 5", "80-75=5")
            .replace("8-5\uFFFD =3", "8-5=3")
            .replace("24\\div 1\uFFFD =24", "24\\div 1=24")
            .replace("1-0.25\uFFFD =0.75", "1-0.25=0.75")
            .replace("255\\div (45+40)\uFFFD =3", "255\\div (45+40)=3")
            .replace("80\\times 3=\uFFFD 240", "80\\times 3=240")
            .replace("=(18+9)\\div (18-9)\uFFFD", "=(18+9)\\div (18-9)")
            .replace("(54-27)\uFFFD千米", "(54-27)千米")
            .replace("(15+30\uFFFD )千米", "(15+30)千米");
        return repaired;
    }

    private DelimitedFormula findNextDelimitedFormula(String text, int searchFrom) {
        for (int start = Math.max(searchFrom, 0); start < text.length(); start++) {
            if (text.charAt(start) != '$' || isEscapedAt(text, start)) {
                continue;
            }
            int delimiterLength = start + 1 < text.length() && text.charAt(start + 1) == '$' ? 2 : 1;
            int close = findClosingMathDelimiter(text, start + delimiterLength, delimiterLength);
            if (close >= 0) {
                return new DelimitedFormula(start, close + delimiterLength,
                    text.substring(start + delimiterLength, close));
            }
        }
        return null;
    }

    private int findClosingMathDelimiter(String text, int cursor, int delimiterLength) {
        int braceDepth = 0;
        while (cursor < text.length()) {
            char current = text.charAt(cursor);
            if (current == '\\') {
                cursor += Math.min(2, text.length() - cursor);
                continue;
            }
            if (current == '{') {
                braceDepth++;
                cursor++;
                continue;
            }
            if (current == '}' && braceDepth > 0) {
                braceDepth--;
                cursor++;
                continue;
            }
            if (current == '$' && braceDepth == 0) {
                if (delimiterLength == 1) {
                    return cursor;
                }
                if (cursor + 1 < text.length() && text.charAt(cursor + 1) == '$') {
                    return cursor;
                }
            }
            cursor++;
        }
        return -1;
    }

    private boolean isEscapedAt(String text, int index) {
        int slashes = 0;
        for (int cursor = index - 1; cursor >= 0 && text.charAt(cursor) == '\\'; cursor--) {
            slashes++;
        }
        return (slashes & 1) == 1;
    }

    private record DelimitedFormula(int start, int endExclusive, String content) {
    }

    private String normalizePlainTextTables(String text) {
        if (text == null || !text.contains("\\begin")) {
            return text;
        }
        boolean hasPlainTextTablePreamble = TEXT_TABLE_PREAMBLE.matcher(text).find()
            || TEXT_TABLE_EMPTY_FIRST_PREAMBLE.matcher(text).find();
        if (!hasPlainTextTablePreamble) {
            return text;
        }
        String normalized = TEXT_TABLE_PREAMBLE.matcher(text).replaceAll("");
        normalized = TEXT_TABLE_EMPTY_FIRST_PREAMBLE.matcher(normalized).replaceAll("");
        normalized = TEXT_TABLE_END.matcher(normalized).replaceAll("");
        normalized = normalized.replace("% D2T: Empty equation removed!", "")
            .replaceAll("\\\\hline\\b", " ")
            .replace("\\_", "_")
            .replace('&', ' ');
        return normalized;
    }

    private ParsedFormulaMetrics stripFormulaMetrics(String latex) {
        Matcher matcher = METRICS_PATTERN.matcher(latex == null ? "" : latex);
        if (!matcher.find()) {
            matcher = LOOSE_METRICS_PATTERN.matcher(latex == null ? "" : latex);
            if (!matcher.find()) {
                return new ParsedFormulaMetrics(latex, null);
            }
        }
        try {
            FormulaMetrics metrics = new FormulaMetrics(
                Double.parseDouble(matcher.group(1)),
                Double.parseDouble(matcher.group(2)),
                matcher.group(3) != null ? Double.parseDouble(matcher.group(3)) : Double.parseDouble(matcher.group(1)),
                matcher.group(4) != null ? Double.parseDouble(matcher.group(4)) : Double.parseDouble(matcher.group(2))
            );
            return new ParsedFormulaMetrics(latex.substring(matcher.end()).trim(), metrics);
        } catch (NumberFormatException e) {
            return new ParsedFormulaMetrics(latex, null);
        }
    }

    private record ParsedFormulaMetrics(String latex, FormulaMetrics metrics) {}

    private String flattenNestedTextColorMath(String text) {
        String current = text;
        String previous;
        do {
            previous = current;
            current = NESTED_TEXT_COLOR_MATH.matcher(current).replaceAll("$1$2");
        } while (!current.equals(previous));
        return current;
    }

    private String splitSharedMathDelimiterBeforeBareStructure(String text) {
        Matcher matcher = SHARED_DELIMITER_BEFORE_BARE_STRUCTURE.matcher(text);
        StringBuffer out = new StringBuffer(text.length());
        while (matcher.find()) {
            matcher.appendReplacement(out, Matcher.quoteReplacement("}$ $"));
        }
        matcher.appendTail(out);
        return out.toString();
    }

    private String stripEmbeddedFormulaAnnotations(String latex) {
        if (latex == null || latex.indexOf("\\pwmetrics") < 0) {
            return latex;
        }
        return latex
            .replaceAll("\\\\pwmetrics\\{[^}]+}\\s*", "")
            .replaceAll("\\\\pwmetrics\\s+[0-9]+(?:\\.[0-9]+)?\\s*,\\s*[0-9]+(?:\\.[0-9]+)?"
                + "(?:\\s*,\\s*[0-9]+(?:\\.[0-9]+)?\\s*,\\s*[0-9]+(?:\\.[0-9]+)?)?\\s+", "")
            .trim();
    }

    private ParsedFormulaStyle stripFormulaStyle(String latex) {
        Matcher matcher = STYLE_HINT_PATTERN.matcher(latex == null ? "" : latex);
        if (!matcher.find()) {
            return new ParsedFormulaStyle(latex, FormulaStyleHints.empty());
        }
        FormulaStyleHints hints = parseStyleHints(matcher.group(1));
        return new ParsedFormulaStyle(latex.substring(matcher.end()).trim(), hints);
    }

    private FormulaStyleHints parseStyleHints(String encoded) {
        boolean asciiFlatParens = false;
        boolean explicitScriptFullSize = false;
        boolean explicitFractionFullSize = false;
        boolean explicitTopFullSize = false;
        boolean forceExplicitFenceTemplate = false;
        boolean explicitBlackColor = false;
        boolean flatParenTemplate = false;
        boolean letterGroupObarTemplate = false;
        boolean textFeComma = false;
        boolean fullwidthTextParen = false;
        boolean mixedAsciiFullwidthParens = false;
        boolean legacyTextFeParenContent = false;
        for (String part : encoded.split(",")) {
            String hint = part.trim();
            if ("asciiFlatParens".equals(hint)) {
                asciiFlatParens = true;
            } else if ("explicitScriptFullSize".equals(hint)) {
                explicitScriptFullSize = true;
            } else if ("explicitFractionFullSize".equals(hint)) {
                explicitFractionFullSize = true;
            } else if ("explicitTopFullSize".equals(hint)) {
                explicitTopFullSize = true;
            } else if ("forceExplicitFenceTemplate".equals(hint)) {
                forceExplicitFenceTemplate = true;
            } else if ("explicitBlackColor".equals(hint)) {
                explicitBlackColor = true;
            } else if ("flatParenTemplate".equals(hint)) {
                flatParenTemplate = true;
            } else if ("letterGroupObarTemplate".equals(hint)) {
                letterGroupObarTemplate = true;
            } else if ("textFeComma".equals(hint)) {
                textFeComma = true;
            } else if ("fullwidthTextParen".equals(hint)) {
                fullwidthTextParen = true;
            } else if ("mixedAsciiFullwidthParens".equals(hint)) {
                mixedAsciiFullwidthParens = true;
            } else if ("legacyTextFeParenContent".equals(hint)) {
                legacyTextFeParenContent = true;
            }
        }
        return new FormulaStyleHints(asciiFlatParens, explicitScriptFullSize, explicitFractionFullSize,
            explicitTopFullSize, forceExplicitFenceTemplate, explicitBlackColor, flatParenTemplate,
            letterGroupObarTemplate, textFeComma, fullwidthTextParen, mixedAsciiFullwidthParens,
            legacyTextFeParenContent, null);
    }

    private record ParsedFormulaStyle(String latex, FormulaStyleHints styleHints) {}

    /**
     * 标准化数学分隔符：将 LaTeX 的 \[...\] 和 \(...\) 转换为 $$...$$ 和 $...$。
     * 这样后续只需要用一种正则就能统一匹配所有公式分隔符。
     *
     * @param text 原始文本
     * @return 分隔符标准化后的文本
     */
    private String normalizeMathDelimiters(String text) {
        StringBuilder normalized = new StringBuilder(text.length());
        int index = 0;
        while (index < text.length()) {
            if (text.charAt(index) != '\\') {
                normalized.append(text.charAt(index++));
                continue;
            }
            int slashStart = index;
            while (index < text.length() && text.charAt(index) == '\\') {
                index++;
            }
            int slashCount = index - slashStart;
            char next = index < text.length() ? text.charAt(index) : '\0';
            boolean delimiter = next == '[' || next == ']' || next == '(' || next == ')';
            if (delimiter && (slashCount & 1) == 1) {
                normalized.append("\\".repeat(slashCount - 1));
                normalized.append(next == '[' || next == ']' ? "$$" : "$");
                index++;
            } else {
                normalized.append("\\".repeat(slashCount));
            }
        }
        return normalized.toString();
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
        String normalized = normalizePlainTextEscapes(text).trim();
        if (!normalized.isBlank()) {
            out.add(new ContentSegment(false, normalized, null));
        }
    }

    private String normalizePlainTextEscapes(String text) {
        return unwrapPlainTextRaiseBoxes(text)
            .replaceAll("\\\\raisebox\\s*(?:\\{\\s*)?[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)"
                + "\\s*(?:pt|px|em|ex)(?:\\s*})?", "")
            .replaceAll("\\\\raisebox\\s*[-+]?\\s*$", "")
            .replace("\\textasciitilde", "~")
            .replace("\\textasciicircum", "^")
            .replace("\\textless", "<")
            .replace("\\textgreater", ">")
            .replace("\\_", "_")
            .replace("\\%", "%")
            .replace("\\#", "#")
            .replace("\\&", "&")
            .replace("\\$", "$")
            .replace("\\{", "{")
            .replace("\\}", "}")
            .replace("\\textbackslash", "\\");
    }

    /** Removes a plain-text layout wrapper while retaining its complete payload. */
    private String unwrapPlainTextRaiseBoxes(String text) {
        if (text == null || !text.contains("\\raisebox")) {
            return text;
        }
        StringBuilder out = new StringBuilder(text.length());
        int cursor = 0;
        while (cursor < text.length()) {
            int command = text.indexOf("\\raisebox", cursor);
            if (command < 0) {
                out.append(text, cursor, text.length());
                break;
            }
            out.append(text, cursor, command);
            int afterCommand = command + "\\raisebox".length();
            if (afterCommand < text.length() && Character.isLetter(text.charAt(afterCommand))) {
                out.append("\\raisebox");
                cursor = afterCommand;
                continue;
            }
            int dimensionStart = skipWhitespace(text, afterCommand);
            if (dimensionStart >= text.length() || text.charAt(dimensionStart) != '{') {
                out.append("\\raisebox");
                cursor = afterCommand;
                continue;
            }
            int dimensionEnd = findMatching(text, dimensionStart, '{', '}');
            if (dimensionEnd < 0) {
                out.append(text, command, text.length());
                break;
            }
            int contentStart = skipWhitespace(text, dimensionEnd + 1);
            for (int optional = 0; optional < 2
                    && contentStart < text.length() && text.charAt(contentStart) == '['; optional++) {
                int optionalEnd = findMatching(text, contentStart, '[', ']');
                if (optionalEnd < 0) {
                    out.append(text, command, text.length());
                    return out.toString();
                }
                contentStart = skipWhitespace(text, optionalEnd + 1);
            }
            if (contentStart >= text.length() || text.charAt(contentStart) != '{') {
                out.append(text, command, dimensionEnd + 1);
                cursor = dimensionEnd + 1;
                continue;
            }
            int contentEnd = findMatching(text, contentStart, '{', '}');
            if (contentEnd < 0) {
                out.append(text, command, text.length());
                break;
            }
            String content = text.substring(contentStart + 1, contentEnd);
            out.append(unwrapPlainTextRaiseBoxes(content));
            cursor = contentEnd + 1;
        }
        return out.toString();
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
        String source = latex == null ? "" : latex;
        String normalized = preNormalizeLatex(source);
        rejectReplacementCharacter(source, normalized);
        LaTeXNode ast = parseNormalizedLatex(normalized).ast();
        ast.setMetadata("sourceLatex", source);
        return ast;
    }

    /**
     * Parses a formula and exposes the coverage evidence required by the official corpus gate.
     */
    public DetailedParseResult parseDetailed(String latex) {
        String source = latex == null ? "" : latex;
        String normalized = preNormalizeLatex(source);
        ParsedAst parsed = parseNormalizedLatex(normalized);
        parsed.ast().setMetadata("sourceLatex", source);
        MathIRNode mathIR = mathIRConverter.convert(parsed.ast());
        LinkedHashSet<String> commands = new LinkedHashSet<>();
        for (Token token : tokenizer.tokenize(source)) {
            if (token.type() == TokenType.COMMAND) {
                commands.add(token.value());
            }
        }

        List<ParseDiagnostic> diagnostics = new ArrayList<>();
        if (source.indexOf('\uFFFD') >= 0 || (normalized != null && normalized.indexOf('\uFFFD') >= 0)) {
            diagnostics.add(new ParseDiagnostic(
                DiagnosticSeverity.ERROR,
                "SOURCE_REPLACEMENT_CHARACTER",
                null,
                "LaTeX contains U+FFFD; the original symbol encoding is unavailable"));
        }
        if (parsed.consumedTokenCount() < parsed.tokenCount()) {
            diagnostics.add(new ParseDiagnostic(
                DiagnosticSeverity.ERROR,
                "UNCONSUMED_TOKEN",
                null,
                "Parser stopped at token " + parsed.consumedTokenCount() + " of " + parsed.tokenCount()
            ));
        }
        collectUnsupportedDiagnostics(mathIR, diagnostics);
        return new DetailedParseResult(
            source,
            normalized,
            parsed.ast(),
            mathIR,
            new ArrayList<>(commands),
            diagnostics
        );
    }

    private void rejectReplacementCharacter(String source, String normalized) {
        if ((source != null && source.indexOf('\uFFFD') >= 0)
                || (normalized != null && normalized.indexOf('\uFFFD') >= 0)) {
            throw new IllegalArgumentException(
                "SOURCE_REPLACEMENT_CHARACTER: LaTeX contains U+FFFD; the original symbol encoding is unavailable");
        }
    }

    private ParsedAst parseNormalizedLatex(String normalizedLatex) {
        List<Token> tokens = tokenizer.tokenize(normalizedLatex == null ? "" : normalizedLatex);
        TokenStream stream = new TokenStream(tokens);
        LaTeXNode root = new LaTeXNode(LaTeXNode.Type.ROOT);
        parseExpression(stream, root);
        normalizeLegacyInfixStructures(root);
        return new ParsedAst(root, stream.position(), tokens.size());
    }

    private void collectUnsupportedDiagnostics(MathIRNode node, List<ParseDiagnostic> diagnostics) {
        if (node == null) {
            return;
        }
        if (node.getType() == MathIRNode.Type.UNSUPPORTED) {
            diagnostics.add(new ParseDiagnostic(
                DiagnosticSeverity.ERROR,
                "UNSUPPORTED_COMMAND",
                node.getValue(),
                "No semantic MathIR mapping exists for " + node.getValue()
            ));
        }
        for (MathIRNode child : node.getChildren()) {
            collectUnsupportedDiagnostics(child, diagnostics);
        }
    }

    /**
     * 解析前的字符串级标准化，吸收 docxtolatex 输出里的兼容性写法：
     *
     * <ul>
     *   <li>{@code \rm}/{@code \bf}/{@code \it} 旧式字体切换：MTEF 层不支持，直接剥离，
     *       内容本身保留（如 {@code { \rm{ 2 } } } → {@code { { 2 } } }）。</li>
     *   <li>{@code \left \begin{...}}：缺失定界符，补 {@code \left.}。</li>
     *   <li>顶层（不在任何环境或花括号内）的 {@code \\} 换行：MathType 原对象是 pile，
     *       整体包一层 {@code \begin{array}{l}...\end{array}} 还原多行结构。</li>
     * </ul>
     */
    public static String preNormalizeLatex(String latex) {
        return preNormalizeLatex(latex, true);
    }

    /**
     * @param wrapTopLevelBreaks whether to wrap top-level {@code \\} breaks in
     *                           an array environment. The MathType-fit worker
     *                           splits pile lines itself, so it must pass false
     *                           to avoid MathJax mtable row-spacing collisions.
     */
    public static String preNormalizeLatex(String latex, boolean wrapTopLevelBreaks) {
        if (latex == null || latex.isBlank()) {
            return latex;
        }
        String normalized = normalizeOuterMathMode(latex)
            .replace("\\text{相遇{\\blacksquare}{\\blacksquare}}", "\\text{相遇时间}")
            .replace("\\text{追及{\\blacksquare}{\\blacksquare}}", "\\text{追及时间}")
            .replace("\\text{不合{\\blacksquare}意}", "\\text{不合题意}")
            .replace("\\text{心想事\\Theta }", "\\text{心想事成}")
            .replace("\\text{梦想\\Theta 真}", "\\text{梦想成真}")
            .replace("\\text{If $x=0$ then $y=2$.}",
                "\\text{If }x=0\\text{ then }y=2\\text{.}")
            .replaceAll("\\\\begin\\s*\\{(?:math|displaymath)\\}", "")
            .replaceAll("\\\\end\\s*\\{(?:math|displaymath)\\}", "")
            .replaceAll("\\\\lt\\b", "<")
            .replaceAll("\\\\gt\\b", ">")
            .replaceAll("\\\\dots\\b", "\\\\ldots")
            .replaceAll("\\\\iff\\b", "\\\\Leftrightarrow")
            .replaceAll("\\\\implies\\b", "\\\\Longrightarrow")
            .replaceAll("\\\\impliedby\\b", "\\\\Longleftarrow")
            .replaceAll("\\\\euro\\s*\\{\\s*}", "")
            .replaceAll("\\\\left\\s+(?=\\\\begin\\b)", "\\\\left. ");
        normalized = normalized.replaceAll("(cm\\^\\{2)\\s*$", "$1}");
        normalized = normalizeStyleWrappedAlignmentMarkers(normalized);
        normalized = stripTopLevelAlignmentMarkers(normalized);
        normalized = normalizeTensorScripts(normalized);
        normalized = normalizeFrownOverset(normalized);
        normalized = normalizeBottomLeftArtifacts(normalized);
        normalized = normalizeUnderRightArrow(normalized);
        normalized = normalizeVisualUnderbraceCounters(normalized);
        normalized = normalizeArrayLineBreakSpacing(normalized);
        normalized = normalized.replace("\\newline", "\\\\");
        if (wrapTopLevelBreaks && hasTopLevelLineBreak(normalized)) {
            normalized = "\\begin{array}{l} " + normalized + " \\end{array}";
        }
        return normalized;
    }

    private static String normalizeOuterMathMode(String latex) {
        String trimmed = latex.trim();
        if (trimmed.startsWith("$$") && trimmed.endsWith("$$") && trimmed.length() >= 4) {
            return trimmed.substring(2, trimmed.length() - 2);
        }
        if (trimmed.startsWith("$") && trimmed.endsWith("$") && trimmed.length() >= 2) {
            return trimmed.substring(1, trimmed.length() - 1);
        }
        if ((trimmed.startsWith("\\(") && trimmed.endsWith("\\)"))
                || (trimmed.startsWith("\\[") && trimmed.endsWith("\\]"))) {
            return trimmed.substring(2, trimmed.length() - 2);
        }
        return latex;
    }

    private static String normalizeStyleWrappedAlignmentMarkers(String latex) {
        if (latex == null || latex.indexOf('&') < 0) {
            return latex;
        }
        Matcher matcher = STYLE_WRAPPED_ALIGNMENT_MARKER.matcher(latex);
        StringBuffer out = new StringBuffer(latex.length());
        while (matcher.find()) {
            matcher.appendReplacement(out, Matcher.quoteReplacement("&" + matcher.group(1)));
        }
        matcher.appendTail(out);
        return out.toString();
    }

    /**
     * Removes column separators left behind when an aligned/array row is extracted as a
     * standalone formula. Separators inside an environment still belong to that environment,
     * and an escaped {@code \&} is a visible ampersand rather than an alignment marker.
     */
    private static String stripTopLevelAlignmentMarkers(String latex) {
        if (latex == null || latex.indexOf('&') < 0) {
            return latex;
        }
        StringBuilder out = new StringBuilder(latex.length());
        int environmentDepth = 0;
        for (int i = 0; i < latex.length(); i++) {
            int beginEnd = environmentDirectiveEnd(latex, i, "begin");
            if (beginEnd >= 0) {
                environmentDepth++;
            } else if (environmentDirectiveEnd(latex, i, "end") >= 0) {
                environmentDepth = Math.max(0, environmentDepth - 1);
            }

            char ch = latex.charAt(i);
            if (ch == '&' && environmentDepth == 0 && !isEscaped(latex, i)) {
                continue;
            }
            out.append(ch);
        }
        return out.toString();
    }

    private static int environmentDirectiveEnd(String latex, int offset, String directive) {
        String marker = "\\" + directive;
        if (!latex.startsWith(marker, offset)) {
            return -1;
        }
        int cursor = offset + marker.length();
        if (cursor < latex.length() && Character.isLetter(latex.charAt(cursor))) {
            return -1;
        }
        cursor = skipWhitespace(latex, cursor);
        if (cursor >= latex.length() || latex.charAt(cursor) != '{') {
            return -1;
        }
        return findMatching(latex, cursor, '{', '}');
    }

    private static boolean isEscaped(String text, int offset) {
        int slashCount = 0;
        for (int i = offset - 1; i >= 0 && text.charAt(i) == '\\'; i--) {
            slashCount++;
        }
        return (slashCount & 1) == 1;
    }

    private static String normalizeTensorScripts(String latex) {
        String marker = "\\tensor*";
        if (latex == null || latex.indexOf(marker) < 0) {
            return latex;
        }
        StringBuilder out = new StringBuilder(latex.length());
        int cursor = 0;
        while (cursor < latex.length()) {
            int start = latex.indexOf(marker, cursor);
            if (start < 0) {
                out.append(latex.substring(cursor));
                break;
            }
            out.append(latex, cursor, start);
            int firstStart = skipWhitespace(latex, start + marker.length());
            if (firstStart >= latex.length() || latex.charAt(firstStart) != '[') {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            int firstEnd = findMatching(latex, firstStart, '[', ']');
            int secondStart = firstEnd < 0 ? -1 : skipWhitespace(latex, firstEnd + 1);
            int secondEnd = secondStart >= 0 && secondStart < latex.length() && latex.charAt(secondStart) == '{'
                ? findMatching(latex, secondStart, '{', '}') : -1;
            int thirdStart = secondEnd < 0 ? -1 : skipWhitespace(latex, secondEnd + 1);
            int thirdEnd = thirdStart >= 0 && thirdStart < latex.length() && latex.charAt(thirdStart) == '{'
                ? findMatching(latex, thirdStart, '{', '}') : -1;
            int fourthStart = thirdEnd < 0 ? -1 : skipWhitespace(latex, thirdEnd + 1);
            boolean fourthIsGroup = fourthStart >= 0 && fourthStart < latex.length() && latex.charAt(fourthStart) == '{';
            int fourthEnd = fourthIsGroup ? findMatching(latex, fourthStart, '{', '}') : findTensorTrailingEnd(latex, fourthStart);
            if (firstEnd < 0 || secondEnd < 0 || thirdEnd < 0) {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            out.append(normalizeTensorScriptSpec(latex.substring(firstStart + 1, firstEnd)));
            out.append(latex, secondStart + 1, secondEnd);
            out.append(latex, thirdStart + 1, thirdEnd);
            if (fourthEnd < 0) {
                cursor = thirdEnd + 1;
                continue;
            }
            if (fourthIsGroup) {
                out.append(latex, fourthStart + 1, fourthEnd);
                cursor = fourthEnd + 1;
            } else {
                out.append(latex, fourthStart, fourthEnd);
                cursor = fourthEnd;
            }
        }
        return out.toString();
    }

    private static String normalizeUnderRightArrow(String latex) {
        String marker = "\\underrightarrow";
        if (latex == null || latex.indexOf(marker) < 0) {
            return latex;
        }
        StringBuilder out = new StringBuilder(latex.length());
        int cursor = 0;
        while (cursor < latex.length()) {
            int start = latex.indexOf(marker, cursor);
            if (start < 0) {
                out.append(latex.substring(cursor));
                break;
            }
            out.append(latex, cursor, start);
            int groupStart = skipWhitespace(latex, start + marker.length());
            if (groupStart >= latex.length() || latex.charAt(groupStart) != '{') {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            int groupEnd = findMatching(latex, groupStart, '{', '}');
            if (groupEnd < 0) {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            out.append(latex, groupStart + 1, groupEnd).append("\\rightarrow");
            cursor = groupEnd + 1;
        }
        return out.toString();
    }

    private static int findTensorTrailingEnd(String text, int start) {
        if (start < 0 || start >= text.length()) {
            return -1;
        }
        int cursor = start;
        while (cursor < text.length()) {
            char ch = text.charAt(cursor);
            if (Character.isWhitespace(ch) || ch == ',' || ch == ';' || ch == ')' || ch == ']' || ch == '}') {
                break;
            }
            if (ch == '\\') {
                break;
            }
            cursor++;
        }
        return cursor > start ? cursor : -1;
    }

    private static String normalizeFrownOverset(String latex) {
        String marker = "\\overset";
        if (latex == null || latex.indexOf(marker) < 0 || latex.indexOf("\\frown") < 0) {
            return latex;
        }
        StringBuilder out = new StringBuilder(latex.length());
        int cursor = 0;
        while (cursor < latex.length()) {
            int start = latex.indexOf(marker, cursor);
            if (start < 0) {
                out.append(latex.substring(cursor));
                break;
            }
            out.append(latex, cursor, start);
            int firstStart = skipWhitespace(latex, start + marker.length());
            if (firstStart >= latex.length() || latex.charAt(firstStart) != '{') {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            int firstEnd = findMatching(latex, firstStart, '{', '}');
            int secondStart = firstEnd < 0 ? -1 : skipWhitespace(latex, firstEnd + 1);
            int secondEnd = secondStart >= 0 && secondStart < latex.length() && latex.charAt(secondStart) == '{'
                ? findMatching(latex, secondStart, '{', '}') : -1;
            if (firstEnd < 0 || secondEnd < 0 || !"\\frown".equals(latex.substring(firstStart + 1, firstEnd).trim())) {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            out.append("\\overarc{");
            out.append(latex, secondStart + 1, secondEnd);
            out.append('}');
            cursor = secondEnd + 1;
        }
        return out.toString();
    }

    private static String normalizeBottomLeftArtifacts(String latex) {
        String marker = "\\bottom";
        if (latex == null || latex.indexOf(marker) < 0) {
            return latex;
        }
        StringBuilder out = new StringBuilder(latex.length());
        int cursor = 0;
        while (cursor < latex.length()) {
            int start = latex.indexOf(marker, cursor);
            if (start < 0) {
                out.append(latex.substring(cursor));
                break;
            }
            out.append(latex, cursor, start);
            int afterBottom = skipWhitespace(latex, start + marker.length());
            if (!latex.startsWith("left", afterBottom)) {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            int groupStart = skipWhitespace(latex, afterBottom + "left".length());
            if (groupStart >= latex.length() || latex.charAt(groupStart) != '{') {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            int groupEnd = findMatching(latex, groupStart, '{', '}');
            if (groupEnd < 0) {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            String body = latex.substring(groupStart + 1, groupEnd);
            out.append(flattenInlineArrayArtifact(body));
            cursor = groupEnd + 1;
        }
        return out.toString();
    }

    private static String flattenInlineArrayArtifact(String text) {
        if (text == null || !text.contains("\\begin{array}")) {
            return text;
        }
        return text
            .replaceAll("\\\\begin\\{array}\\{[^}]*}", "")
            .replaceAll("\\\\end\\{array}", "")
            .replace('&', ' ')
            .replaceAll("\\s+", " ")
            .trim();
    }

    private static String normalizeTensorScriptSpec(String spec) {
        if (spec == null || spec.isBlank()) {
            return "";
        }
        return spec.replaceAll("\\^\\s*\\{\\s*}", "")
            .replaceAll("_\\s*\\{\\s*}", "");
    }

    private static int skipWhitespace(String text, int index) {
        int cursor = index;
        while (cursor < text.length() && Character.isWhitespace(text.charAt(cursor))) {
            cursor++;
        }
        return cursor;
    }

    private static int findMatching(String text, int open, char openChar, char closeChar) {
        int depth = 0;
        for (int i = open; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '\\') {
                i++;
                continue;
            }
            if (ch == openChar) {
                depth++;
            } else if (ch == closeChar) {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static final Pattern VISUAL_UNDERBRACE_COUNTER =
        Pattern.compile("(?<body>(?:\\\\mathrm\\{[0-9]\\}|[0-9])+\\s*(?:\\\\cdots|\\.\\.\\.)\\s*(?:\\\\mathrm\\{[0-9]\\}|[0-9])+)\\s*(?<note>\\d+个(?:\\\\mathrm\\{[0-9]\\}|[0-9])+(?:和\\d+个(?:\\\\mathrm\\{[0-9]\\}|[0-9])+)?|\\d+个[0-9]+(?:和\\d+个[0-9]+)?)︸");

    private static String normalizeVisualUnderbraceCounters(String latex) {
        if (latex == null || latex.indexOf('︸') < 0) {
            return latex;
        }
        String joinedVisualCounters = normalizeJoinedVisualUnderbraceCounters(latex);

        Matcher matcher = VISUAL_UNDERBRACE_COUNTER.matcher(joinedVisualCounters);
        StringBuffer out = new StringBuffer(joinedVisualCounters.length());
        while (matcher.find()) {
            String body = matcher.group("body").trim();
            String note = matcher.group("note").trim();
            matcher.appendReplacement(out, Matcher.quoteReplacement("\\underbrace{" + body + "}_{" + note + "}"));
        }
        matcher.appendTail(out);
        return out.toString();
    }

    private static String normalizeJoinedVisualUnderbraceCounters(String latex) {
        StringBuilder out = new StringBuilder(latex.length());
        int cursor = 0;
        while (cursor < latex.length()) {
            int brace = latex.indexOf('︸', cursor);
            if (brace < 0) {
                out.append(latex.substring(cursor));
                break;
            }
            int ellipsis = lastEllipsisBefore(latex, cursor, brace);
            if (ellipsis < 0) {
                out.append(latex, cursor, brace + 1);
                cursor = brace + 1;
                continue;
            }
            int suffixStart = ellipsis + ellipsisLengthAt(latex, ellipsis);
            int ge = latex.indexOf('个', suffixStart);
            if (ge < 0 || ge > brace) {
                out.append(latex, cursor, brace + 1);
                cursor = brace + 1;
                continue;
            }
            while (suffixStart < ge && Character.isWhitespace(latex.charAt(suffixStart))) {
                suffixStart++;
            }
            String tail = latex.substring(suffixStart, ge);
            String unit = readCounterUnit(latex, ge + 1, brace);
            if (unit.isEmpty() || !tail.startsWith(unit) || tail.length() == unit.length()) {
                out.append(latex, cursor, brace + 1);
                cursor = brace + 1;
                continue;
            }
            String count = tail.substring(unit.length());
            if (!isCounterCount(count)) {
                out.append(latex, cursor, brace + 1);
                cursor = brace + 1;
                continue;
            }
            int bodyStart = findVisualUnderbraceBodyStart(latex, cursor, ellipsis);
            String body = latex.substring(bodyStart, suffixStart) + unit;
            String note = count + latex.substring(ge, brace);
            out.append(latex, cursor, bodyStart);
            out.append("\\underbrace{")
                .append(body)
                .append("}_{")
                .append(note)
                .append("}");
            cursor = brace + 1;
        }
        return out.toString();
    }

    private static int ellipsisLengthAt(String latex, int index) {
        if (latex.startsWith("\\cdot \\cdot \\cdot", index)) {
            return "\\cdot \\cdot \\cdot".length();
        }
        if (latex.startsWith("\\cdots", index)) {
            return "\\cdots".length();
        }
        return "...".length();
    }

    private static int findVisualUnderbraceBodyStart(String latex, int from, int ellipsis) {
        int best = from;
        String[] delimiters = {"\\times", "\\div", "=", "+", "-", "("};
        for (String delimiter : delimiters) {
            int idx = latex.lastIndexOf(delimiter, ellipsis);
            if (idx >= from) {
                best = Math.max(best, idx + delimiter.length());
            }
        }
        while (best < ellipsis && Character.isWhitespace(latex.charAt(best))) {
            best++;
        }
        return best;
    }

    private static int lastEllipsisBefore(String latex, int from, int to) {
        int cdots = latex.lastIndexOf("\\cdots", to);
        int dots = latex.lastIndexOf("...", to);
        int cdotDots = latex.lastIndexOf("\\cdot \\cdot \\cdot", to);
        int best = Math.max(Math.max(cdots, dots), cdotDots);
        return best >= from ? best : -1;
    }

    private static String readCounterUnit(String latex, int from, int to) {
        int i = from;
        if (latex.startsWith("\\mathrm{", i)) {
            int close = latex.indexOf('}', i + "\\mathrm{".length());
            if (close > i && close < to) {
                return latex.substring(i, close + 1);
            }
        }
        while (i < to && (Character.isDigit(latex.charAt(i)) || isAsciiLetter(latex.charAt(i)))) {
            i++;
        }
        return i > from ? latex.substring(from, i) : "";
    }

    private static boolean isCounterCount(String value) {
        if (value == null || value.isBlank()) {
            return false;
        }
        for (int i = 0; i < value.length(); i++) {
            char ch = value.charAt(i);
            if (!Character.isDigit(ch) && !isAsciiLetter(ch)) {
                return false;
            }
        }
        return true;
    }

    private static boolean isAsciiLetter(char ch) {
        return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
    }

    private static String normalizeArrayLineBreakSpacing(String latex) {
        if (latex == null || latex.indexOf("\\begin{array}") < 0 || latex.indexOf("\\,") < 0) {
            return latex;
        }
        StringBuilder out = new StringBuilder(latex.length());
        int envDepth = 0;
        for (int i = 0; i < latex.length(); i++) {
            if (latex.startsWith("\\begin{array}", i)) {
                envDepth++;
                out.append("\\begin{array}");
                i += "\\begin{array}".length() - 1;
                continue;
            }
            if (latex.startsWith("\\end{array}", i)) {
                envDepth = Math.max(0, envDepth - 1);
                out.append("\\end{array}");
                i += "\\end{array}".length() - 1;
                continue;
            }
            if (envDepth > 0 && latex.startsWith("\\\\,", i)) {
                out.append("\\\\ ");
                i += 2;
                continue;
            }
            out.append(latex.charAt(i));
        }
        return out.toString();
    }

    /** 检测不在花括号分组或 begin/end 环境内的 {@code \\} 换行。 */
    private static boolean hasTopLevelLineBreak(String latex) {
        int braceDepth = 0;
        int envDepth = 0;
        for (int i = 0; i < latex.length(); i++) {
            char c = latex.charAt(i);
            if (c == '\\' && i + 1 < latex.length()) {
                char next = latex.charAt(i + 1);
                if (next == '\\') {
                    if (braceDepth == 0 && envDepth == 0) {
                        return true;
                    }
                    i++;
                    continue;
                }
                if (next == '{' || next == '}') {
                    i++;
                    continue;
                }
                if (latex.startsWith("\\begin", i)) {
                    envDepth++;
                } else if (latex.startsWith("\\end", i)) {
                    envDepth--;
                }
                continue;
            }
            if (c == '{') {
                braceDepth++;
            } else if (c == '}') {
                braceDepth = Math.max(braceDepth - 1, 0);
            }
        }
        return false;
    }

    /**
     * 将 LaTeX 直接解析为 MathML-aligned IR，供 Phase 3 之后的语义层和诊断使用。
     */
    public MathIRNode parseMathIR(String latex) {
        ParsedFormulaMetrics parsedMetrics = stripFormulaMetrics(latex);
        ParsedFormulaStyle parsedStyle = stripFormulaStyle(parsedMetrics.latex());
        return mathIRConverter.convert(parseLaTeX(parsedStyle.latex()));
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
     *   <li>检查下一个 Token 是否为右花括号（标志花括号分组结束），若是则退出</li>
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

            // 遇到右花括号 }，表示当前花括号分组结束。普通方括号应保留为公式字符。
            if (token.type() == TokenType.RBRACE) {
                break;
            }

            // 解析一个原子元素（字符、命令或分组）
            LaTeXNode node = parseAtom(stream);
            if (node != null) {
                // 检查原子元素后面是否有上标 ^ 或下标 _，若有则包装为 SUPERSCRIPT/SUBSCRIPT 节点
                node = parseScripts(stream, node);
                if (isStyleDeclaration(node)) {
                    LaTeXNode content = new LaTeXNode(LaTeXNode.Type.GROUP);
                    parseExpression(stream, content);
                    node.addChild(content);
                    parent.addChild(node);
                    break;
                }
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
     *   <li>LBRACKET/RBRACKET → 普通方括号字符；可选参数由专用解析函数处理</li>
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
            case LBRACKET -> new LaTeXNode(LaTeXNode.Type.CHAR, "[");
            case RBRACKET -> new LaTeXNode(LaTeXNode.Type.CHAR, "]");
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
            case "\\frac", "\\dfrac", "\\tfrac", "\\cfrac" -> parseFrac(stream, cmd);
            case "\\nicefrac" -> parseNiceFraction(stream);
            case "\\binom", "\\dbinom", "\\tbinom" -> parseBinomial(stream, cmd);
            case "\\sqrt" -> parseSqrt(stream);
            case "\\root" -> parseRootOf(stream);
            case "\\longdiv" -> parseLongDiv(stream);
            case "\\left" -> parseLeftRight(stream);
            case "\\overline", "\\underline", "\\hat", "\\tilde",
                 "\\vec", "\\bar", "\\dot", "\\jstatus", "\\jointstatus",
                 "\\overleftarrow", "\\overleftrightarrow", "\\overrightarrow",
                 "\\underleftarrow", "\\underleftrightarrow", "\\underrightarrow",
                 "\\arc", "\\overarc", "\\overparen", "\\wideparen",
                 "\\bra", "\\ket",
                 "\\overbrace", "\\underbrace", "\\overbracket", "\\underbracket",
                 "\\boxed", "\\cancel", "\\bcancel", "\\xcancel" -> parseUnaryCommand(stream, cmd);
            case "\\xrightarrow", "\\xleftarrow", "\\xleftrightarrow", "\\xlongequal",
                 "\\xLeftrightarrow", "\\xLongleftarrow", "\\xLongleftrightarrow", "\\xLongrightarrow",
                 "\\xlongleftarrow", "\\xlongleftrightarrow", "\\xlongrightarrow" ->
                parseExtensibleArrowCommand(stream, cmd);
            case "\\overset", "\\underset" -> parseBinaryCommand(stream, cmd);
            case "\\stackrel" -> parseBinaryCommand(stream, "\\overset");
            case "\\buildrel" -> parseBuildRel(stream);
            case "\\not" -> parseUnaryCommand(stream, cmd);
            case "\\bmod", "\\mod", "\\pmod" -> parseModuloCommand(stream, cmd);
            case "\\ltr", "\\rtl" -> parseDirectionCommand(stream, cmd);
            case "\\sideset" -> parseSideSet(stream);
            case "\\style" -> parseCssStyleCommand(stream, cmd);
            case "\\color" -> parseColorDeclaration(stream, cmd);
            case "\\textcolor" -> parseTextColor(stream, cmd);
            case "\\raisebox" -> parseRaiseBox(stream, cmd);
            case "\\tiny", "\\scriptsize", "\\small", "\\normalsize", "\\large", "\\Large",
                 "\\LARGE", "\\huge", "\\Huge", "\\displaystyle", "\\textstyle",
                 "\\scriptstyle", "\\scriptscriptstyle", "\\rm", "\\bf", "\\it", "\\cal",
                 "\\sf", "\\tt" -> parseStyleDeclaration(cmd);
            case "\\braket" -> parseBraketCommand(stream, cmd);
            case "\\mathrm", "\\mathbf", "\\mathit", "\\textit", "\\textbf", "\\emph", "\\boldsymbol",
                 "\\mathcal", "\\mathbb", "\\Bbb", "\\mathfrak", "\\frak", "\\mathsf", "\\mathtt" ->
                parseScopedFontCommand(stream, cmd);
            case "\\text", "\\mbox", "\\textrm", "\\operatorname" -> parseTextCommand(stream, cmd);
            case "\\sum", "\\sumop", "\\int", "\\intop", "\\iint", "\\iiint", "\\oint",
                 "\\prod", "\\coprod", "\\bigcup", "\\bigcap", "\\bigvee", "\\bigwedge",
                 "\\biguplus", "\\bigoplus", "\\bigotimes" -> parseBigOp(stream, cmd);
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

    private LaTeXNode parseBinaryCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        node.addChild(parseRequiredGroup(stream));
        node.addChild(parseRequiredGroup(stream));
        return node;
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
        boolean endedWithRowBreak = false;

        while (stream.hasNext()) {
            stream.skipWhitespace();
            if (!stream.hasNext()) {
                break;
            }
            if (stream.matchesEnvironmentEnd(envName)) {
                if (!endedWithRowBreak) {
                    finalizeArrayCell(currentRow, currentCell);
                    finalizeArrayRow(arrayNode, currentRow);
                }
                stream.consumeEnvironmentEnd();
                break;
            }

            Token token = stream.peek();
            if (!seenContent && currentCell.getChildren().isEmpty()
                    && token.type() == TokenType.CHAR && ",".equals(token.value())) {
                stream.next();
                continue;
            }
            if (token.type() == TokenType.COMMAND
                    && ("\\hline".equals(token.value()) || "\\hdashline".equals(token.value()))) {
                stream.next();
                rowLines.set(rowLines.size() - 1, 1);
                if ("\\hdashline".equals(token.value())) {
                    arrayNode.setMetadata("rowLineStyle", "dashed");
                }
                continue;
            }
            if (token.type() == TokenType.CHAR && "&".equals(token.value())) {
                stream.next();
                finalizeArrayCell(currentRow, currentCell);
                currentCell = new LaTeXNode(LaTeXNode.Type.CELL);
                seenContent = true;
                endedWithRowBreak = false;
                continue;
            }
            if (token.type() == TokenType.COMMAND && "\\\\".equals(token.value())) {
                stream.next();
                parseOptionalBracketGroup(stream);
                finalizeArrayCell(currentRow, currentCell);
                finalizeArrayRow(arrayNode, currentRow);
                currentRow = new LaTeXNode(LaTeXNode.Type.ROW);
                currentCell = new LaTeXNode(LaTeXNode.Type.CELL);
                rowLines.add(0);
                seenContent = false;
                endedWithRowBreak = true;
                continue;
            }

            LaTeXNode child = parseAtom(stream);
            if (child != null) {
                child = parseScripts(stream, child);
                currentCell.addChild(child);
                seenContent = true;
                endedWithRowBreak = false;
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
        return parseFrac(stream, "\\frac");
    }

    private LaTeXNode parseFrac(TokenStream stream, String command) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.FRACTION, command);
        node.setMetadata("latexCommand", command);
        node.setMetadata("fractionStyle", switch (command) {
            case "\\dfrac" -> "display";
            case "\\tfrac" -> "text";
            case "\\cfrac" -> "continued";
            default -> "auto";
        });
        // 读取分子参数
        node.addChild(parseRequiredGroup(stream));
        // 读取分母参数
        node.addChild(parseRequiredGroup(stream));
        return node;
    }

    private LaTeXNode parseNiceFraction(TokenStream stream) {
        LaTeXNode node = parseFrac(stream, "\\nicefrac");
        node.setMetadata("fractionStyle", "slash");
        return node;
    }

    private LaTeXNode parseBinomial(TokenStream stream, String command) {
        LaTeXNode pile = createTwoRowPile(parseRequiredGroup(stream), parseRequiredGroup(stream));
        LaTeXNode fence = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\left(");
        fence.setMetadata("leftDelimiter", "(");
        fence.setMetadata("rightDelimiter", ")");
        fence.setMetadata("latexCommand", command);
        fence.setMetadata("binomialStyle", command.substring(1));
        fence.addChild(pile);
        return fence;
    }

    private LaTeXNode parseRootOf(TokenStream stream) {
        LaTeXNode degree = parseRequiredGroup(stream);
        stream.skipWhitespace();
        if (stream.hasNext() && stream.peek().type() == TokenType.COMMAND
                && "\\of".equals(stream.peek().value())) {
            stream.next();
        }
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.SQRT, "\\root");
        node.setMetadata("latexCommand", "\\root");
        node.addChild(degree);
        node.addChild(parseRequiredGroup(stream));
        return node;
    }

    private LaTeXNode parseBuildRel(TokenStream stream) {
        LaTeXNode annotation = parseRequiredGroup(stream);
        stream.skipWhitespace();
        if (stream.hasNext() && stream.peek().type() == TokenType.COMMAND
                && "\\over".equals(stream.peek().value())) {
            stream.next();
        }
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\overset");
        node.setMetadata("latexCommand", "\\buildrel");
        node.addChild(annotation);
        node.addChild(parseRequiredGroup(stream));
        return node;
    }

    private LaTeXNode parseModuloCommand(TokenStream stream, String command) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        group.setMetadata("latexCommand", command);
        LaTeXNode text = new LaTeXNode(LaTeXNode.Type.TEXT, "\\operatorname");
        LaTeXNode word = new LaTeXNode(LaTeXNode.Type.GROUP);
        for (char ch : "mod".toCharArray()) {
            word.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        text.addChild(word);
        group.addChild(text);
        LaTeXNode argument = parseRequiredGroup(stream);
        if ("\\pmod".equals(command)) {
            group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, "("));
        }
        group.addChild(argument);
        if ("\\pmod".equals(command)) {
            group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, ")"));
        }
        return group;
    }

    private LaTeXNode parseDirectionCommand(TokenStream stream, String command) {
        LaTeXNode content = parseRequiredGroup(stream);
        content.setMetadata("direction", "\\rtl".equals(command) ? "rtl" : "ltr");
        content.setMetadata("directionCommand", command);
        return content;
    }

    private LaTeXNode parseSideSet(TokenStream stream) {
        SideScripts left = parseSideScripts(stream);
        SideScripts right = parseSideScripts(stream);
        LaTeXNode base = parseRequiredGroup(stream);

        LaTeXNode emptyBase = new LaTeXNode(LaTeXNode.Type.GROUP);
        LaTeXNode leading = attachScripts(emptyBase, left);
        leading.setMetadata("latexCommand", "\\sideset");
        LaTeXNode trailing = attachScripts(base, right);

        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        group.setMetadata("latexCommand", "\\sideset");
        group.addChild(leading);
        group.addChild(trailing);
        return group;
    }

    private SideScripts parseSideScripts(TokenStream stream) {
        stream.skipWhitespace();
        if (!stream.hasNext() || stream.peek().type() != TokenType.LBRACE) {
            return new SideScripts(null, null);
        }
        stream.next();
        LaTeXNode lower = null;
        LaTeXNode upper = null;
        while (stream.hasNext() && stream.peek().type() != TokenType.RBRACE) {
            stream.skipWhitespace();
            if (!stream.hasNext() || stream.peek().type() == TokenType.RBRACE) {
                break;
            }
            Token token = stream.next();
            if (token.type() == TokenType.UNDERSCORE) {
                lower = parseRequiredGroup(stream);
            } else if (token.type() == TokenType.CARET) {
                upper = parseRequiredGroup(stream);
            }
        }
        if (stream.hasNext() && stream.peek().type() == TokenType.RBRACE) {
            stream.next();
        }
        return new SideScripts(lower, upper);
    }

    private LaTeXNode attachScripts(LaTeXNode base, SideScripts scripts) {
        LaTeXNode result = base;
        if (scripts.lower() != null) {
            LaTeXNode sub = new LaTeXNode(LaTeXNode.Type.SUBSCRIPT, "_");
            sub.addChild(result);
            sub.addChild(scripts.lower());
            result = sub;
        }
        if (scripts.upper() != null) {
            LaTeXNode sup = new LaTeXNode(LaTeXNode.Type.SUPERSCRIPT, "^");
            sup.addChild(result);
            sup.addChild(scripts.upper());
            result = sup;
        }
        return result;
    }

    private record SideScripts(LaTeXNode lower, LaTeXNode upper) {}

    private LaTeXNode parseStyleDeclaration(String command) {
        LaTeXNode style = new LaTeXNode(LaTeXNode.Type.STYLE, command);
        style.setMetadata("styleDeclaration", "true");
        if (Set.of("\\tiny", "\\scriptsize", "\\small", "\\normalsize", "\\large", "\\Large",
                "\\LARGE", "\\huge", "\\Huge").contains(command)) {
            style.setMetadata("styleKind", "size");
            style.setMetadata("fontSizePt", switch (command) {
                case "\\tiny" -> "6.0";
                case "\\scriptsize" -> "8.0";
                case "\\small" -> "10.0";
                case "\\large" -> "14.4";
                case "\\Large" -> "17.28";
                case "\\LARGE" -> "20.74";
                case "\\huge" -> "24.88";
                case "\\Huge" -> "29.86";
                default -> "12.0";
            });
        } else if (command.endsWith("style")) {
            style.setMetadata("styleKind", "math-style");
            style.setMetadata("mathStyle", command.substring(1));
        } else {
            style.setMetadata("styleKind", "font");
            style.setMetadata("fontVariant", legacyFontVariant(command));
        }
        return style;
    }

    private LaTeXNode parseScopedFontCommand(TokenStream stream, String command) {
        LaTeXNode style = new LaTeXNode(LaTeXNode.Type.STYLE, command);
        style.setMetadata("styleKind", "font");
        style.setMetadata("fontVariant", fontVariant(command));
        style.addChild(parseRequiredGroup(stream));
        return style;
    }

    private LaTeXNode parseCssStyleCommand(TokenStream stream, String command) {
        String css = extractPlainText(parseRequiredGroup(stream));
        LaTeXNode style = new LaTeXNode(LaTeXNode.Type.STYLE, command);
        style.setMetadata("styleKind", "css");
        style.setMetadata("css", css);
        Matcher size = Pattern.compile("(?:^|;)\\s*font-size\\s*:\\s*([0-9.]+)(px|pt)?", Pattern.CASE_INSENSITIVE)
            .matcher(css);
        if (size.find()) {
            double value = Double.parseDouble(size.group(1));
            if ("px".equalsIgnoreCase(size.group(2))) {
                value *= 0.75d;
            }
            style.setMetadata("fontSizePt", Double.toString(value));
        }
        style.addChild(parseRequiredGroup(stream));
        return style;
    }

    private LaTeXNode parseColorDeclaration(TokenStream stream, String command) {
        LaTeXNode model = parseOptionalBracketGroup(stream);
        String values = extractPlainText(parseRequiredGroup(stream));
        LaTeXNode style = new LaTeXNode(LaTeXNode.Type.STYLE, command);
        style.setMetadata("styleDeclaration", "true");
        style.setMetadata("styleKind", "color");
        configureColorMetadata(style, model == null ? "named" : extractPlainText(model), values);
        return style;
    }

    private LaTeXNode parseTextColor(TokenStream stream, String command) {
        LaTeXNode model = parseOptionalBracketGroup(stream);
        String values = extractPlainText(parseRequiredGroup(stream));
        LaTeXNode style = new LaTeXNode(LaTeXNode.Type.STYLE, command);
        style.setMetadata("styleKind", "color");
        configureColorMetadata(style, model == null ? "named" : extractPlainText(model), values);
        style.addChild(parseRequiredGroup(stream));
        return style;
    }

    private void configureColorMetadata(LaTeXNode style, String model, String values) {
        String normalizedModel = model == null ? "named" : model.trim().toLowerCase();
        String normalizedValues = values == null ? "" : values.trim();
        if ("named".equals(normalizedModel)) {
            String rgb = namedColorRgb(normalizedValues);
            if (rgb != null) {
                style.setMetadata("colorModel", "rgb");
                style.setMetadata("colorValue", rgb);
                style.setMetadata("colorName", normalizedValues.toLowerCase());
                return;
            }
        }
        style.setMetadata("colorModel", normalizedModel);
        style.setMetadata("colorValue", normalizedValues);
    }

    private String namedColorRgb(String name) {
        if (name == null) {
            return null;
        }
        return switch (name.trim().toLowerCase()) {
            case "black" -> "0,0,0";
            case "white" -> "1,1,1";
            case "red" -> "1,0,0";
            case "green" -> "0,0.5,0";
            case "blue" -> "0,0,1";
            case "cyan", "aqua" -> "0,1,1";
            case "magenta", "fuchsia" -> "1,0,1";
            case "yellow" -> "1,1,0";
            case "gray", "grey" -> "0.5,0.5,0.5";
            case "maroon" -> "0.502,0,0";
            case "olive" -> "0.5,0.5,0";
            case "navy" -> "0,0,0.5";
            case "purple" -> "0.5,0,0.5";
            case "teal" -> "0,0.5,0.5";
            case "silver" -> "0.75,0.75,0.75";
            default -> null;
        };
    }

    private LaTeXNode parseRaiseBox(TokenStream stream, String command) {
        String shift = extractPlainText(parseRequiredGroup(stream)).trim();
        LaTeXNode height = parseOptionalBracketGroup(stream);
        LaTeXNode depth = parseOptionalBracketGroup(stream);
        LaTeXNode style = new LaTeXNode(LaTeXNode.Type.STYLE, command);
        style.setMetadata("styleKind", "vertical-shift");
        style.setMetadata("verticalShift", shift);
        style.setMetadata("verticalShiftPt", Double.toString(parseTeXDimensionPt(shift)));
        if (height != null) {
            style.setMetadata("boxHeight", extractPlainText(height).trim());
        }
        if (depth != null) {
            style.setMetadata("boxDepth", extractPlainText(depth).trim());
        }
        style.addChild(parseRequiredGroup(stream));
        return style;
    }

    private double parseTeXDimensionPt(String dimension) {
        Matcher matcher = Pattern.compile("^([-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+))\\s*(pt|px|em|ex)?$",
            Pattern.CASE_INSENSITIVE).matcher(dimension == null ? "" : dimension.trim());
        if (!matcher.matches()) {
            throw new IllegalArgumentException("Invalid \\raisebox dimension: " + dimension);
        }
        double value = Double.parseDouble(matcher.group(1));
        String unit = matcher.group(2) == null ? "pt" : matcher.group(2).toLowerCase();
        return switch (unit) {
            case "px" -> value * 0.75d;
            case "em" -> value * 12.0d;
            case "ex" -> value * 6.0d;
            default -> value;
        };
    }

    private String legacyFontVariant(String command) {
        return switch (command) {
            case "\\bf" -> "bold";
            case "\\it" -> "italic";
            case "\\cal" -> "script";
            case "\\sf" -> "sans-serif";
            case "\\tt" -> "monospace";
            default -> "normal";
        };
    }

    private String fontVariant(String command) {
        return switch (command) {
            case "\\mathbf", "\\textbf", "\\boldsymbol" -> "bold";
            case "\\mathit", "\\textit", "\\emph" -> "italic";
            case "\\mathcal" -> "script";
            case "\\mathbb", "\\Bbb" -> "double-struck";
            case "\\mathfrak", "\\frak" -> "fraktur";
            case "\\mathsf" -> "sans-serif";
            case "\\mathtt" -> "monospace";
            default -> "normal";
        };
    }

    private LaTeXNode createTwoRowPile(LaTeXNode top, LaTeXNode bottom) {
        LaTeXNode array = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        array.setMetadata("environment", "array");
        array.setMetadata("columnSpec", "c");
        array.setMetadata("columnCount", "1");
        array.setMetadata("columnLines", "0,0");
        array.setMetadata("rowLines", "0,0,0");
        array.setMetadata("binomialPile", "true");
        for (LaTeXNode item : List.of(top, bottom)) {
            LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
            LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
            cell.addChild(item);
            row.addChild(cell);
            array.addChild(row);
        }
        return array;
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
        LaTeXNode quotient = parseOptionalBracketGroup(stream);
        if (quotient != null) {
            node.addChild(parseRequiredGroup(stream)); // divisor
            node.addChild(quotient);                   // quotient
            node.addChild(parseRequiredGroup(stream)); // dividend
            return node;
        }
        node.addChild(parseRequiredGroup(stream));          // divisor
        node.addChild(new LaTeXNode(LaTeXNode.Type.GROUP)); // no quotient slot
        node.addChild(parseRequiredGroup(stream));          // dividend
        return node;
    }

    private LaTeXNode parseOptionalBracketGroup(TokenStream stream) {
        if (!stream.hasNext() || stream.peek().type() != TokenType.LBRACKET) {
            return null;
        }
        stream.next();
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        while (stream.hasNext()) {
            stream.skipWhitespace();
            if (!stream.hasNext() || stream.peek().type() == TokenType.RBRACKET) {
                break;
            }
            LaTeXNode child = parseAtom(stream);
            if (child != null) {
                child = parseScripts(stream, child);
                group.addChild(child);
            }
        }
        if (stream.hasNext() && stream.peek().type() == TokenType.RBRACKET) {
            stream.next();
        }
        return group;
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
        // 读取左定界符字符；docxtolatex 常输出 "\left ("，空白不属于定界符。
        stream.skipWhitespace();
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
                stream.skipWhitespace();
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
            case "\\langle" -> "⟨";
            case "\\rangle" -> "⟩";
            case "\\llbracket" -> "⟦";
            case "\\rrbracket" -> "⟧";
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
     * 解析标准双槽 Dirac 记号 \braket{a|b}。
     *
     * <p>该命令不是普通一元命令：必须把花括号中的顶层竖线拆成左槽/右槽两部分，
     * 供后续 MathIR 和 MTEF tmDIRAC 全链路使用，而不是退化成普通字符拼接。</p>
     */
    private LaTeXNode parseBraketCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        node.setMetadata("middleDelimiter", "|");

        LaTeXNode argument = parseRequiredGroup(stream);
        DiracSlots slots = splitDiracSlots(argument);
        if (slots == null) {
            node.addChild(argument);
            return node;
        }

        node.addChild(slots.left());
        node.addChild(slots.right());
        return node;
    }

    private DiracSlots splitDiracSlots(LaTeXNode argument) {
        if (argument == null || argument.getType() != LaTeXNode.Type.GROUP) {
            return null;
        }

        int separatorIndex = -1;
        for (int index = 0; index < argument.getChildren().size(); index++) {
            LaTeXNode child = argument.getChildren().get(index);
            if (child.getType() == LaTeXNode.Type.CHAR && "|".equals(child.getValue())) {
                if (separatorIndex >= 0) {
                    return null;
                }
                separatorIndex = index;
            }
        }

        if (separatorIndex <= 0 || separatorIndex >= argument.getChildren().size() - 1) {
            return null;
        }

        LaTeXNode left = new LaTeXNode(LaTeXNode.Type.GROUP);
        LaTeXNode right = new LaTeXNode(LaTeXNode.Type.GROUP);
        for (int index = 0; index < separatorIndex; index++) {
            left.addChild(argument.getChildren().get(index));
        }
        for (int index = separatorIndex + 1; index < argument.getChildren().size(); index++) {
            right.addChild(argument.getChildren().get(index));
        }
        return new DiracSlots(left, right);
    }

    /**
     * 解析 amsmath 风格的可伸缩箭头命令（\xrightarrow / \xleftarrow）。
     *
     * <p>依据官方 {@code amsmath.sty} 定义，二者签名均为
     * {@code \xrightarrow[below]{above}} / {@code \xleftarrow[below]{above}}：
     * 方括号内可选参数对应箭头下方标注，花括号内必需参数对应箭头上方标注。</p>
     *
     * <p>当前仍保守限定在单线 left/right 两个公开变体；不在本方法中扩展双线、鱼叉等其它 arrow family。</p>
     */
    private LaTeXNode parseExtensibleArrowCommand(TokenStream stream, String cmd) {
        LaTeXNode node = new LaTeXNode(LaTeXNode.Type.COMMAND, cmd);
        node.setMetadata("templateFamily", "TM_ARROW");
        node.setMetadata("arrowDirection", switch (cmd) {
            case "\\xleftarrow", "\\xlongleftarrow", "\\xLongleftarrow" -> "left";
            case "\\xleftrightarrow", "\\xLeftrightarrow", "\\xlongleftrightarrow", "\\xLongleftrightarrow" -> "both";
            case "\\xlongequal" -> "none";
            default -> "right";
        });
        node.setMetadata("arrowVariant",
            cmd.startsWith("\\xLong") || "\\xLeftrightarrow".equals(cmd) || "\\xlongequal".equals(cmd)
                ? "double" : "single");

        LaTeXNode bottomAnnotation = parseOptionalBracketGroup(stream);
        LaTeXNode topAnnotation = parseRequiredGroup(stream);

        node.addChild(topAnnotation);
        if (bottomAnnotation != null && !bottomAnnotation.getChildren().isEmpty()) {
            node.setMetadata("annotationPlacement", "top-bottom");
            node.addChild(bottomAnnotation);
        } else {
            node.setMetadata("annotationPlacement", "top");
        }
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
        stream.skipWhitespace();
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

    private void normalizeLegacyInfixStructures(LaTeXNode container) {
        for (LaTeXNode child : List.copyOf(container.getChildren())) {
            normalizeLegacyInfixStructures(child);
        }
        List<LaTeXNode> children = container.getChildren();
        for (int index = 0; index < children.size(); index++) {
            LaTeXNode marker = children.get(index);
            if (marker.getType() != LaTeXNode.Type.COMMAND
                    || !Set.of("\\over", "\\atop", "\\choose", "\\brace", "\\brack").contains(marker.getValue())) {
                continue;
            }
            LaTeXNode numerator = new LaTeXNode(LaTeXNode.Type.GROUP);
            LaTeXNode denominator = new LaTeXNode(LaTeXNode.Type.GROUP);
            children.subList(0, index).forEach(numerator::addChild);
            children.subList(index + 1, children.size()).forEach(denominator::addChild);
            LaTeXNode replacement;
            if ("\\over".equals(marker.getValue())) {
                replacement = new LaTeXNode(LaTeXNode.Type.FRACTION, "\\over");
                replacement.addChild(numerator);
                replacement.addChild(denominator);
            } else {
                LaTeXNode pile = createTwoRowPile(numerator, denominator);
                replacement = switch (marker.getValue()) {
                    case "\\choose" -> wrapLegacyPile(pile, "(", ")", marker.getValue());
                    case "\\brace" -> wrapLegacyPile(pile, "{", "}", marker.getValue());
                    case "\\brack" -> wrapLegacyPile(pile, "[", "]", marker.getValue());
                    default -> pile;
                };
            }
            replacement.setMetadata("latexCommand", marker.getValue());
            children.clear();
            children.add(replacement);
            break;
        }
    }

    private LaTeXNode wrapLegacyPile(LaTeXNode pile, String left, String right, String command) {
        LaTeXNode fence = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\left" + left);
        fence.setMetadata("leftDelimiter", left);
        fence.setMetadata("rightDelimiter", right);
        fence.setMetadata("latexCommand", command);
        fence.addChild(pile);
        return fence;
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
            int scriptLookahead = stream.position();
            stream.skipWhitespace();
            if (!stream.hasNext()) {
                stream.setPosition(scriptLookahead);
                break;
            }
            Token t = stream.peek();
            if (t.type() == TokenType.COMMAND
                    && ("\\limits".equals(t.value()) || "\\nolimits".equals(t.value()))) {
                stream.next();
                base.setMetadata("limitPlacement", "\\limits".equals(t.value()) ? "under-over" : "scripts");
                base.setMetadata("limitCommand", t.value());
                continue;
            }
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
                stream.setPosition(scriptLookahead);
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
                if (isStyleDeclaration(child)) {
                    LaTeXNode content = new LaTeXNode(LaTeXNode.Type.GROUP);
                    parseExpression(stream, content);
                    child.addChild(content);
                    group.addChild(child);
                    break;
                }
                group.addChild(child);
            }
        }
        // 消费右花括号 }（如果存在）
        if (stream.hasNext()) {
            stream.next();
        }
        return group;
    }

    private boolean isStyleDeclaration(LaTeXNode node) {
        return node != null && node.getType() == LaTeXNode.Type.STYLE
            && "true".equals(node.getMetadata("styleDeclaration"));
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

        int position() {
            return pos;
        }

        void setPosition(int position) {
            pos = Math.max(0, Math.min(position, tokens.size()));
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

    private record DiracSlots(LaTeXNode left, LaTeXNode right) {}

    private record ParsedAst(LaTeXNode ast, int consumedTokenCount, int tokenCount) {}
}
