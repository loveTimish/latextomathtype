package com.lz.paperword.core.render;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutCompiler;
import com.lz.paperword.core.layout.VerticalLayoutSpec;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionHeader;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionStep;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RuleSpan;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.PNGTranscoder;
import org.freehep.graphicsio.emf.EMFGraphics2D;
import org.scilab.forge.jlatexmath.TeXConstants;
import org.scilab.forge.jlatexmath.TeXFormula;
import org.scilab.forge.jlatexmath.TeXIcon;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Path2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * LaTeX 公式图片渲染器 — 将 LaTeX 公式渲染为 OLE 预览图。
 *
 * <h2>核心用途</h2>
 * <p>为嵌入 Word 的 MathType OLE 对象生成预览图片。在 .docx 中，每个 OLE 公式
 * 由两部分组成：OLE 二进制数据（可双击编辑）+ 预览图（静态显示）。
 * 本类负责生成后者。当用户在 Word 中打开文档后，MathType 会自动用其内部的
 * 原生预览替换我们生成的过渡预览。</p>
 *
 * <h2>渲染策略与回退链</h2>
 * <p>{@link #renderToPng} 方法采用多级回退策略，按优先级依次尝试：</p>
 * <ol>
 *   <li><b>MathJax HTTP 服务</b> — 调用远程 MathJax 渲染端点获取 SVG，再转为 PNG。
 *       渲染质量最高，支持所有 LaTeX 符号，但依赖网络可用性</li>
 *   <li><b>dvisvgm 本地工具链</b> — 调用系统安装的 latex + dvisvgm 命令行工具，
 *       tex → dvi → svg → png。需要 TeX 发行版环境</li>
 *   <li><b>JLaTeXMath 纯 Java 渲染</b> — 使用 JLaTeXMath 库在内存中渲染，
 *       无外部依赖，作为最终兜底方案</li>
 * </ol>
 *
 * <h2>OLE 预览图渲染</h2>
 * <p>{@link #renderForOlePreview} 专用于 OLE 形状预览，优先生成矢量预览（EMF），
 * 回退到 PNG。预览采用更接近行内公式的度量方式，尽量贴近 Word 中 MathType
 * 内联对象的实际占位。</p>
 *
 * <h2>配置属性</h2>
 * <ul>
 *   <li>{@code paperword.mathjax.enabled} — 是否启用 MathJax HTTP 渲染（默认 true）</li>
 *   <li>{@code paperword.mathjax.only} — 仅使用 MathJax，失败直接退化到 JLaTeXMath（默认 false）</li>
 *   <li>{@code paperword.mathjax.endpoint} — MathJax 渲染服务 URL</li>
 *   <li>{@code paperword.latex.command} — latex 命令路径（默认 "latex"）</li>
 *   <li>{@code paperword.dvisvgm.command} — dvisvgm 命令路径（默认 "dvisvgm"）</li>
 *   <li>{@code paperword.latex.timeout.seconds} — 外部命令超时秒数（默认 20）</li>
 * </ul>
 */
public class LaTeXImageRenderer {

    private static final Logger log = LoggerFactory.getLogger(LaTeXImageRenderer.class);
    private final VerticalLayoutCompiler verticalLayoutCompiler = new VerticalLayoutCompiler();

    /** 默认公式字体大小（磅），对应 Word 中正文公式的标准尺寸 */
    private static final float DEFAULT_SIZE = 13f;

    /** 通用渲染缩放因子，提高位图预览清晰度 */
    private static final float RENDER_SCALE = 2.0f;

    // === 系统属性键（通过 -D 或 System.setProperty 配置） ===
    private static final String LATEX_CMD_PROP = "paperword.latex.command";
    private static final String DVISVGM_CMD_PROP = "paperword.dvisvgm.command";
    private static final String INKSCAPE_CMD_PROP = "paperword.inkscape.command";
    private static final String RENDER_TIMEOUT_PROP = "paperword.latex.timeout.seconds";
    private static final String MATHJAX_ENABLED_PROP = "paperword.mathjax.enabled";
    private static final String MATHJAX_ONLY_PROP = "paperword.mathjax.only";
    private static final String MATHJAX_ENDPOINT_PROP = "paperword.mathjax.endpoint";
    private static final String MATHJAX_TIMEOUT_PROP = "paperword.mathjax.timeout.seconds";
    private static final int DEFAULT_TIMEOUT_SECONDS = 20;
    private static final int DEFAULT_MATHJAX_TIMEOUT_SECONDS = 12;
    private static final float PX_PER_PT = 1.0f / 0.75f;
    private static final String WINDOWS_MIKTEX_BIN = System.getenv("LOCALAPPDATA") == null
        ? ""
        : System.getenv("LOCALAPPDATA") + "\\Programs\\MiKTeX\\miktex\\bin\\x64";

    /** 标记外部工具（latex/dvisvgm）是否不可用，避免反复尝试失败 */
    private volatile boolean externalToolUnavailable = false;
    /** 标记 SVG -> EMF 转换工具是否不可用，避免反复尝试失败 */
    private volatile boolean externalVectorToolUnavailable = false;
    /** 标记 xlop 官方渲染链不可用，避免长除法反复走失败的官方渲染链 */
    private volatile boolean xlopToolchainUnavailable = false;

    /** 用于调用 MathJax HTTP 渲染服务的 HTTP 客户端（连接超时 6 秒） */
    private final HttpClient httpClient = HttpClient.newBuilder()
        .connectTimeout(java.time.Duration.ofSeconds(6))
        .build();

    /**
     * 预览图数据记录。
     *
     * @param data        图片二进制数据（PNG 或 EMF）
     * @param widthPx     显示宽度（像素），用于 VML shape 的 style:width
     * @param heightPx    显示高度（像素），用于 VML shape 的 style:height
     * @param extension   文件扩展名（"png" 或 "emf"）
     * @param contentType MIME 类型（"image/png" 或 "image/x-emf"）
     */
    public record PreviewImage(byte[] data, int widthPx, int heightPx, String extension, String contentType) {}

    private static final float STRUCTURED_RENDER_SCALE = 3.0f;
    private static final int STRUCTURED_FONT_SIZE = 42;
    private static final int STRUCTURED_PADDING = 18;
    private static final int STRUCTURED_ROW_GAP = 8;

    /**
     * 将 LaTeX 公式渲染为 PNG 字节数组（使用默认字体大小）。
     *
     * <p>内部调用 {@link #renderToPng(String, float)}，采用多级回退策略：
     * MathJax HTTP → dvisvgm 本地工具链 → JLaTeXMath 纯 Java 渲染。</p>
     *
     * @param latex LaTeX 公式字符串（不含 $ 定界符）
     * @return PNG 图片字节数组；渲染失败时返回占位图
     */
    public byte[] renderToPng(String latex) {
        return renderToPng(latex, DEFAULT_SIZE);
    }

    /**
     * 为 OLE 形状生成预览图。
     *
     * <p><b>渲染策略：</b>优先尝试矢量预览（EMF），失败时再回退到 PNG。
     * 为了避免行内 OLE 对象在 Word 里占位过高/过宽，这里采用更接近行内公式的
     * STYLE_TEXT 度量，并收紧内边距。</p>
     *
     * @param latex LaTeX 公式字符串
     * @return 预览图数据（含 PNG 字节和 1 倍显示尺寸）；渲染失败返回 null
     */
    public PreviewImage renderForOlePreview(String latex) {
        PreviewImage remotePreview = renderPreviewViaMathJaxHttp(latex);
        if (remotePreview != null) {
            return remotePreview;
        }
        try {
            TeXFormula formula = new TeXFormula(latex);
            float renderScale = 3.0f;
            int style = usesStackedPreviewStyle(latex) ? TeXConstants.STYLE_DISPLAY : TeXConstants.STYLE_TEXT;
            TeXIcon icon = formula.createTeXIcon(style, DEFAULT_SIZE * renderScale);
            // 预览图刻意比旧实现放松一点，避免 Word 中公式显得过于紧凑。
            icon.setInsets(createOlePreviewInsets(latex, renderScale));
            icon.setForeground(Color.BLACK);

            // 确保最小尺寸，避免空图异常
            int width = Math.max(icon.getIconWidth(), 10);
            int height = Math.max(icon.getIconHeight(), 10);

            // 使用 TYPE_INT_RGB（不透明白底），避免 Word 图片插值时出现灰边
            BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2 = image.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, width, height);  // 白色背景填充
            icon.paintIcon(null, g2, 0, 0);    // 绘制公式
            g2.dispose();

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(image, "png", baos);
            byte[] pngData = baos.toByteArray();
            if (pngData.length == 0) {
                return null;
            }
            int displayWidth = Math.max((int)(width / renderScale), 10);
            int displayHeight = Math.max((int)(height / renderScale), 10);
            return new PreviewImage(pngData, displayWidth, displayHeight, "png", "image/png");
        } catch (Exception e) {
            log.debug("Preview render failed: {}", latex, e);
            return null;
        }
    }

    public PreviewImage renderForOlePreview(LaTeXNode latexAst, String latex) {
        PreviewImage structured = renderStructuredPreview(latexAst);
        if (structured != null) {
            return structured;
        }
        return renderForOlePreview(latex);
    }

    /**
     * 长除法专用图片渲染入口。
     *
     * <p>优先使用本地 xlop 文档渲染，以获得接近教材竖式的版式；若本地 TeX 工具链不可用，
     * 再回退到当前项目已有的结构化预览绘制逻辑，保证文档里至少插入稳定图片。</p>
     */
    public PreviewImage renderLongDivisionPicture(LaTeXNode latexAst, String rawLatex) {
        VerticalLayoutSpec layout = latexAst == null ? null : verticalLayoutCompiler.compile(latexAst);
        String xlopLatex = buildLongDivisionPictureLatex(layout, latexAst);
        if (!xlopLatex.isBlank()) {
            PreviewImage xlopPreview = renderXlopLongDivisionPreview(xlopLatex, DEFAULT_SIZE + 1f);
            if (xlopPreview != null) {
                return xlopPreview;
            }
            // 先尝试官方 xlop 链；一旦当前环境无法稳定产图，则明确记录后退回本地长除法图片兜底。
            log.warn("Xlop long division render failed, fallback to local preview: {}", rawLatex);
        }

        PreviewImage fallbackPreview = renderLocalLongDivisionFallback(latexAst, rawLatex, xlopLatex);
        if (fallbackPreview != null) {
            return fallbackPreview;
        }
        return null;
    }

    private PreviewImage renderStructuredPreview(LaTeXNode latexAst) {
        VerticalLayoutSpec layout = verticalLayoutCompiler.compile(latexAst);
        if (layout == null) {
            return null;
        }
        return renderLayoutPreview(layout);
    }

    private PreviewImage renderXlopLongDivisionPreview(String latex, float size) {
        byte[] png = renderViaDvisvgm(latex, size, List.of("\\usepackage{xlop}"), false);
        return buildPreviewFromPng(png);
    }

    String buildLongDivisionPictureLatex(LaTeXNode latexAst) {
        return buildLongDivisionPictureLatex(
            latexAst == null ? null : verticalLayoutCompiler.compile(latexAst),
            latexAst
        );
    }

    private String buildLongDivisionPictureLatex(VerticalLayoutSpec layout, LaTeXNode latexAst) {
        LongDivisionOperands operands = normalizeLongDivisionOperands(layout, latexAst);
        if (operands == null) {
            return "";
        }

        List<String> options = new ArrayList<>();
        if (operands.divisor().contains(".") || operands.dividend().contains(".")) {
            options.add("decimalsepsymbol={.}");
            // 小数长除法只补最基本的小数点配置，其余完全交给 xlop 官方默认排版。
            options.add("shiftdecimalsep=none");
        }

        // 整数长除法必须走 \opidiv，避免被 xlop 当成普通除法算出小数商。
        String xlopCommand = options.isEmpty() ? "\\opidiv" : "\\opdiv";

        // 长除法这里只喂给 xlop 被除数和除数，不再手写商、步骤或版式参数。
        if (options.isEmpty()) {
            return """
                %s{%s}{%s}
                """.formatted(xlopCommand, operands.dividend(), operands.divisor());
        }
        return """
            %s[%s]{%s}{%s}
            """.formatted(xlopCommand, String.join(",", options), operands.dividend(), operands.divisor());
    }

    private LongDivisionOperands normalizeLongDivisionOperands(VerticalLayoutSpec layout, LaTeXNode latexAst) {
        LaTeXNode longDivisionNode = findFirstLongDivisionNode(latexAst);
        VerticalLayoutSpec.LongDivisionHeader header = layout != null && layout.isLongDivision()
            ? layout.longDivisionHeader()
            : null;
        String divisor = sanitizeLongDivisionOperand(
            !extractLongDivisionOperand(longDivisionNode, 0).isBlank()
                ? extractLongDivisionOperand(longDivisionNode, 0)
                : header != null ? header.divisor() : ""
        );
        String dividend = sanitizeLongDivisionOperand(
            !extractLongDivisionOperand(longDivisionNode, 2).isBlank()
                ? extractLongDivisionOperand(longDivisionNode, 2)
                : header != null ? header.dividend() : ""
        );
        if (divisor.isBlank() || dividend.isBlank()) {
            return null;
        }
        return new LongDivisionOperands(divisor, dividend);
    }

    private LaTeXNode findFirstLongDivisionNode(LaTeXNode node) {
        if (node == null) {
            return null;
        }
        if (node.getType() == LaTeXNode.Type.LONG_DIVISION) {
            return node;
        }
        for (LaTeXNode child : node.getChildren()) {
            LaTeXNode found = findFirstLongDivisionNode(child);
            if (found != null) {
                return found;
            }
        }
        return null;
    }

    private String extractLongDivisionOperand(LaTeXNode node, int childIndex) {
        if (node == null || node.getType() != LaTeXNode.Type.LONG_DIVISION || node.getChildren().size() <= childIndex) {
            return "";
        }
        return flattenNodeText(node.getChildren().get(childIndex));
    }

    private String sanitizeLongDivisionOperand(String value) {
        String normalized = blankToEmpty(value).replaceAll("\\s+", "");
        if (normalized.matches("[0-9]+(?:\\.[0-9]+)?")) {
            return normalized;
        }
        return "";
    }

    private PreviewImage buildPreviewFromPng(byte[] pngData) {
        if (pngData == null || pngData.length == 0) {
            return null;
        }
        try {
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(pngData));
            if (image == null) {
                return null;
            }
            return new PreviewImage(
                pngData,
                Math.max(image.getWidth(), 10),
                Math.max(image.getHeight(), 10),
                "png",
                "image/png"
            );
        } catch (IOException e) {
            log.debug("Failed to inspect PNG dimensions", e);
            return null;
        }
    }

    private PreviewImage renderLocalLongDivisionFallback(LaTeXNode latexAst, String rawLatex, String xlopLatex) {
        VerticalLayoutSpec layout = latexAst == null ? null : verticalLayoutCompiler.compile(latexAst);
        LongDivisionOperands operands = normalizeLongDivisionOperands(layout, latexAst);
        if (xlopLatex.isBlank()) {
            // 无法正规化到官方 opdiv 命令时，允许回到本地兜底，保证旧模板仍可出图。
            PreviewImage structured = latexAst == null ? null : renderLongDivisionPreview(latexAst);
            if (structured != null) {
                return structured;
            }
            if (operands != null) {
                PreviewImage computed = renderComputedLongDivision(operands.divisor(), "", operands.dividend());
                if (computed != null) {
                    return computed;
                }
            }
            return buildPreviewFromPng(renderToPng(rawLatex, DEFAULT_SIZE + 1f));
        }
        PreviewImage structured = latexAst == null ? null : renderLongDivisionPreview(latexAst);
        if (structured != null) {
            return structured;
        }
        if (operands != null) {
            PreviewImage computed = renderComputedLongDivision(operands.divisor(), "", operands.dividend());
            if (computed != null) {
                return computed;
            }
        }
        return buildPreviewFromPng(renderToPng(rawLatex, DEFAULT_SIZE + 1f));
    }

    private boolean shouldFallbackToLocalLongDivisionPreview() {
        return externalToolUnavailable || xlopToolchainUnavailable;
    }

    private PreviewImage renderLayoutPreview(VerticalLayoutSpec layout) {
        if (layout.isLongDivision()) {
            return renderLongDivisionLayoutPreview(layout);
        }
        if (layout.kind() == VerticalLayoutSpec.Kind.DECIMAL) {
            return renderDecimalLayoutPreview(layout);
        }
        return renderGridLayoutPreview(layout);
    }

    private PreviewImage renderGridLayoutPreview(VerticalLayoutSpec layout) {
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        List<VerticalRow> rows = layout.rows();
        int maxCols = layout.columnCount();
        if (rows.isEmpty() || maxCols <= 0) {
            return null;
        }

        int digitWidth = Math.max(metrics.stringWidth("0"), metrics.stringWidth("8"));
        int operatorWidth = Math.max(metrics.stringWidth("+"), Math.max(metrics.stringWidth("-"), metrics.stringWidth("×")));
        int[] colWidths = new int[maxCols];
        for (int col = 0; col < maxCols; col++) {
            boolean operatorColumn = false;
            int maxWidth = 0;
            for (VerticalRow row : rows) {
                String text = col < row.cells().size() ? renderText(row.cells().get(col)) : "";
                if (text.contains("+") || text.contains("-") || text.contains("×")) {
                    operatorColumn = true;
                }
                maxWidth = Math.max(maxWidth, metrics.stringWidth(text));
            }
            colWidths[col] = operatorColumn
                ? Math.max(operatorWidth + 16, maxWidth + 10)
                : Math.max(digitWidth + 18, maxWidth + 12);
        }

        int previewPadding = STRUCTURED_PADDING + 8;
        int gap = 14;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP + 6;
        int totalWidth = previewPadding * 2 + sum(colWidths) + gap * Math.max(maxCols - 1, 0);
        int totalHeight = previewPadding * 2 + rows.size() * rowAdvance;

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int rightEdge = previewPadding + sum(colWidths) + gap * Math.max(maxCols - 1, 0);
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            VerticalRow row = rows.get(rowIndex);
            int baseline = previewPadding + metrics.getAscent() + rowIndex * rowAdvance;
            if (!row.segments().isEmpty()) {
                for (var segment : row.segments()) {
                    String text = renderText(segment.text());
                    int textWidth = metrics.stringWidth(text);
                    int right = previewPadding + sum(colWidths, segment.endColumn() + 1) + gap * segment.endColumn();
                    g2.drawString(text, right - textWidth, baseline);
                }
            } else {
                int x = previewPadding;
                for (int col = 0; col < maxCols; col++) {
                    String text = col < row.cells().size() ? renderText(row.cells().get(col)) : "";
                    int textWidth = metrics.stringWidth(text);
                    if (!text.isEmpty()) {
                        g2.drawString(text, x + colWidths[col] - textWidth, baseline);
                    }
                    x += colWidths[col] + gap;
                }
            }
            RuleSpan span = findRuleSpan(layout.ruleSpans(), rowIndex);
            if (span != null) {
                int lineStart = previewPadding + sum(colWidths, span.startColumn()) + gap * span.startColumn() - 4;
                int lineEnd = previewPadding + sum(colWidths, span.endColumn() + 1) + gap * span.endColumn() + 2;
                int lineY = resolveRowSeparatorY(metrics, baseline);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(lineStart, lineY, lineEnd, lineY);
            }
        }
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private PreviewImage renderDecimalLayoutPreview(VerticalLayoutSpec layout) {
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        List<String> rows = layout.rows().stream()
            .map(this::resolveDecimalRowText)
            .toList();

        FontMetrics initialMetrics = metrics;
        int maxWidth = rows.stream()
            .mapToInt(initialMetrics::stringWidth)
            .max()
            .orElse(initialMetrics.stringWidth("0")) + 4;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP;
        int totalWidth = STRUCTURED_PADDING * 2 + maxWidth;
        int totalHeight = STRUCTURED_PADDING * 2 + rows.size() * rowAdvance;

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int rightEdge = totalWidth - STRUCTURED_PADDING;
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            String row = rows.get(rowIndex);
            int baseline = STRUCTURED_PADDING + metrics.getAscent() + rowIndex * rowAdvance;
            if (!row.isEmpty()) {
                g2.drawString(row, rightEdge - metrics.stringWidth(row), baseline);
            }
            if (findRuleSpan(layout.ruleSpans(), rowIndex) != null) {
                int lineY = resolveRowSeparatorY(metrics, baseline);
                int lineStart = rightEdge - metrics.stringWidth(row);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(lineStart, lineY, rightEdge, lineY);
            }
        }
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private String resolveDecimalRowText(VerticalRow row) {
        return row.cells().stream()
            .filter(text -> text != null && !text.isBlank())
            .findFirst()
            .orElse("");
    }

    private PreviewImage renderLongDivisionLayoutPreview(VerticalLayoutSpec layout) {
        if (!layout.hasStructuredLongDivisionSteps()) {
            return renderLegacyLongDivisionLayoutPreview(layout);
        }
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        LongDivisionHeader header = layout.longDivisionHeader();
        VerticalRow dividendRow = findDividendRow(layout);
        int maxCols = Math.max(layout.columnCount() - 2, 1);
        int digitWidth = Math.max(metrics.stringWidth("0"), metrics.stringWidth("8"));
        int[] colWidths = new int[maxCols];
        for (int col = 0; col < colWidths.length; col++) {
            int maxWidth = 0;
            for (String text : trimDivisionCells(dividendRow == null ? List.of() : dividendRow.cells())) {
                maxWidth = Math.max(maxWidth, metrics.stringWidth(renderText(text)));
            }
            for (LongDivisionStep step : layout.longDivisionSteps()) {
                String product = col < trimDivisionCells(step.productRow().cells()).size()
                    ? renderText(trimDivisionCells(step.productRow().cells()).get(col)) : "";
                String remainder = col < trimDivisionCells(step.remainderRow().cells()).size()
                    ? renderText(trimDivisionCells(step.remainderRow().cells()).get(col)) : "";
                maxWidth = Math.max(maxWidth, Math.max(metrics.stringWidth(product), metrics.stringWidth(remainder)));
            }
            colWidths[col] = Math.max(digitWidth + 16, maxWidth + 10);
        }

        int gap = 8;
        int divisorWidth = metrics.stringWidth(header.divisor());
        int quotientWidth = metrics.stringWidth(header.quotient());
        int hookGap = 12;
        int hookWidth = 18;
        // 除法参考图的纵向节奏更松，这里额外拉开行距并增加上下留白。
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP + 4;
        int stepsWidth = sum(colWidths) + gap * Math.max(colWidths.length - 1, 0);
        int dividendWidth = Math.max(metrics.stringWidth(header.dividend()) + 8, stepsWidth);
        int verticalPadding = STRUCTURED_PADDING + 8;
        int totalWidth = STRUCTURED_PADDING * 2 + divisorWidth + hookGap + hookWidth
            + Math.max(dividendWidth, quotientWidth + 8) + 12;
        int totalHeight = verticalPadding * 2 + rowAdvance * (2 + layout.longDivisionSteps().size() * 2);

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int symbolX = STRUCTURED_PADDING + divisorWidth + hookGap;
        int dividendX = symbolX + hookWidth;
        int quotientBaseline = verticalPadding + metrics.getAscent();
        int dividendBaseline = quotientBaseline + rowAdvance;
        int topLineY = dividendBaseline - metrics.getAscent() - 2;
        int rightEdge = dividendX + dividendWidth;
        int bottomY = dividendBaseline + Math.max(layout.longDivisionSteps().size() * 2 - 1, 0) * rowAdvance + metrics.getDescent();

        if (!header.quotient().isBlank()) {
            drawTextAcrossColumns(g2, stepStartXForWidths(rightEdge, colWidths, gap), colWidths, gap,
                colWidths.length - 1, quotientBaseline, header.quotient(), metrics);
        }
        g2.drawString(header.divisor(), STRUCTURED_PADDING, dividendBaseline);

        int stepStartX = rightEdge - stepsWidth;
        drawRowAcrossColumns(g2, stepStartX, colWidths, gap, dividendRow, dividendBaseline, metrics);

        int baseline = dividendBaseline + rowAdvance;
        for (LongDivisionStep step : layout.longDivisionSteps()) {
            drawRowAcrossColumns(g2, stepStartX, colWidths, gap, step.productRow(), baseline, metrics);
            List<String> productCells = trimDivisionCells(step.productRow().cells());
            List<String> remainderCells = trimDivisionCells(step.remainderRow().cells());
            int lineStartCol = Math.max(Math.min(firstNonEmptyColumn(productCells), firstNonEmptyColumn(remainderCells)), 0);
            int lineEndCol = Math.max(lastNonEmptyColumn(productCells), lastNonEmptyColumn(remainderCells));
            int lineY = baseline + Math.max(metrics.getDescent() / 2, 3);
            g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g2.drawLine(
                stepStartX + sum(colWidths, lineStartCol) + gap * lineStartCol - 2,
                lineY,
                stepStartX + sum(colWidths, lineEndCol + 1) + gap * lineEndCol + 2,
                lineY
            );
            baseline += rowAdvance;
            drawRowAcrossColumns(g2, stepStartX, colWidths, gap, step.remainderRow(), baseline, metrics);
            baseline += rowAdvance;
        }

        drawLongDivisionSign(g2, symbolX, topLineY, bottomY, rightEdge + 2);
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private PreviewImage renderLegacyLongDivisionLayoutPreview(VerticalLayoutSpec layout) {
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        LongDivisionHeader header = layout.longDivisionHeader();
        List<VerticalRow> rows = layout.rows().stream()
            .filter(row -> row.kind() != VerticalLayoutSpec.RowKind.DIVIDEND)
            .toList();
        int maxCols = Math.max(layout.columnCount(), 1);
        int firstStepColumn = layout.isLongDivision() ? 2 : 0;
        int digitWidth = Math.max(metrics.stringWidth("0"), metrics.stringWidth("8"));
        int[] colWidths = new int[Math.max(maxCols - firstStepColumn, 1)];
        for (int col = 0; col < colWidths.length; col++) {
            int maxWidth = 0;
            for (VerticalRow row : rows) {
                int actualCol = col + firstStepColumn;
                String text = actualCol < row.cells().size() ? renderText(row.cells().get(actualCol)) : "";
                maxWidth = Math.max(maxWidth, metrics.stringWidth(text));
            }
            colWidths[col] = Math.max(digitWidth + 16, maxWidth + 10);
        }

        int gap = 8;
        int divisorWidth = metrics.stringWidth(header.divisor());
        int quotientWidth = metrics.stringWidth(header.quotient());
        int hookGap = 12;
        int hookWidth = 18;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP + 4;
        int stepsWidth = sum(colWidths) + gap * Math.max(colWidths.length - 1, 0);
        int dividendWidth = Math.max(metrics.stringWidth(header.dividend()) + 8, stepsWidth);
        int verticalPadding = STRUCTURED_PADDING + 8;
        int totalWidth = STRUCTURED_PADDING * 2 + divisorWidth + hookGap + hookWidth + Math.max(dividendWidth, quotientWidth + 8) + 12;
        int totalHeight = verticalPadding * 2 + rowAdvance * (2 + rows.size());

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int symbolX = STRUCTURED_PADDING + divisorWidth + hookGap;
        int dividendX = symbolX + hookWidth;
        int quotientBaseline = verticalPadding + metrics.getAscent();
        int dividendBaseline = quotientBaseline + rowAdvance;
        int stepTopBaseline = dividendBaseline + rowAdvance;
        int topLineY = dividendBaseline - metrics.getAscent() - 2;
        int rightEdge = dividendX + dividendWidth;
        int bottomY = stepTopBaseline + Math.max(rows.size() - 1, 0) * rowAdvance + metrics.getDescent();

        if (!header.quotient().isBlank()) {
            drawTextAcrossColumns(g2, stepStartXForWidths(rightEdge, colWidths, gap), colWidths, gap,
                colWidths.length - 1, quotientBaseline, header.quotient(), metrics);
        }
        g2.drawString(header.divisor(), STRUCTURED_PADDING, dividendBaseline);

        int stepStartX = rightEdge - stepsWidth;
        drawTextAcrossColumns(g2, stepStartX, colWidths, gap, colWidths.length - 1, dividendBaseline, header.dividend(), metrics);
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            VerticalRow row = rows.get(rowIndex);
            int rowBaseline = stepTopBaseline + rowIndex * rowAdvance;
            drawRowAcrossColumns(g2, stepStartX, colWidths, gap, row, rowBaseline, metrics);
            RuleSpan span = findRuleSpan(layout.ruleSpans(), rowIndex);
            if (span != null) {
                int adjustedStart = Math.max(span.startColumn() - firstStepColumn, 0);
                int adjustedEnd = Math.max(span.endColumn() - firstStepColumn, 0);
                int lineStart = stepStartX + sum(colWidths, adjustedStart) + gap * adjustedStart - 2;
                int lineEnd = stepStartX + sum(colWidths, adjustedEnd + 1) + gap * adjustedEnd + 2;
                int lineY = resolveRowSeparatorY(metrics, rowBaseline);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(lineStart, lineY, lineEnd, lineY);
            }
        }

        drawLongDivisionSign(g2, symbolX, topLineY, bottomY, rightEdge + 2);
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private boolean isCompositeLongDivisionArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY || node.getChildren().size() < 2) {
            return false;
        }
        return findFirstDescendant(node.getChildren().get(0), LaTeXNode.Type.LONG_DIVISION) != null
            && findNestedArray(node.getChildren().get(1), node) != null;
    }

    private PreviewImage renderCompositeLongDivisionPreview(LaTeXNode node) {
        LaTeXNode longDiv = findFirstDescendant(node.getChildren().get(0), LaTeXNode.Type.LONG_DIVISION);
        LaTeXNode stepsArray = findNestedArray(node.getChildren().get(1), node);
        if (longDiv == null || stepsArray == null) {
            return null;
        }

        String divisor = flattenNodeText(longDiv.getChildren().size() > 0 ? longDiv.getChildren().get(0) : null);
        String quotient = flattenNodeText(longDiv.getChildren().size() > 1 ? longDiv.getChildren().get(1) : null);
        String dividend = flattenNodeText(longDiv.getChildren().size() > 2 ? longDiv.getChildren().get(2) : null);

        List<List<String>> stepRows = extractRowTexts(stepsArray);
        if (stepRows.isEmpty()) {
            return renderComputedLongDivision(divisor, quotient, dividend);
        }
        for (List<String> row : stepRows) {
            trimRight(row);
        }

        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        int maxCols = stepRows.stream().mapToInt(List::size).max().orElse(0);
        int digitWidth = Math.max(metrics.stringWidth("0"), metrics.stringWidth("8"));
        int[] colWidths = new int[Math.max(maxCols, 1)];
        for (int col = 0; col < colWidths.length; col++) {
            int maxWidth = 0;
            for (List<String> row : stepRows) {
                if (col < row.size()) {
                    maxWidth = Math.max(maxWidth, metrics.stringWidth(renderText(row.get(col))));
                }
            }
            colWidths[col] = Math.max(digitWidth + 16, maxWidth + 10);
        }

        int gap = 8;
        int divisorWidth = metrics.stringWidth(divisor);
        int quotientWidth = metrics.stringWidth(quotient);
        int hookGap = 12;
        int hookWidth = 18;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP;
        int stepsWidth = sum(colWidths) + gap * Math.max(colWidths.length - 1, 0);
        int dividendWidth = Math.max(metrics.stringWidth(dividend) + 8, stepsWidth);
        int totalWidth = STRUCTURED_PADDING * 2 + divisorWidth + hookGap + hookWidth + Math.max(dividendWidth, quotientWidth + 8) + 12;
        int totalHeight = STRUCTURED_PADDING * 2 + rowAdvance * (2 + stepRows.size());

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int symbolX = STRUCTURED_PADDING + divisorWidth + hookGap;
        int dividendX = symbolX + hookWidth;
        int quotientBaseline = STRUCTURED_PADDING + metrics.getAscent();
        int dividendBaseline = quotientBaseline + rowAdvance;
        int stepTopBaseline = dividendBaseline + rowAdvance;
        int topLineY = dividendBaseline - metrics.getAscent() - 2;
        int rightEdge = dividendX + dividendWidth;
        int bottomY = stepTopBaseline + Math.max(stepRows.size() - 1, 0) * rowAdvance + metrics.getDescent();

        if (!quotient.isBlank()) {
            drawTextAcrossColumns(g2, rightEdge - stepsWidth, colWidths, gap, colWidths.length - 1, quotientBaseline, quotient, metrics);
        }
        g2.drawString(divisor, STRUCTURED_PADDING, dividendBaseline);
        drawTextAcrossColumns(g2, rightEdge - stepsWidth, colWidths, gap, colWidths.length - 1, dividendBaseline, dividend, metrics);

        int stepStartX = rightEdge - stepsWidth;
        int[] rowLines = parsePartitionArray(stepsArray.getMetadata("rowLines"), stepRows.size() + 1);
        for (int rowIndex = 0; rowIndex < stepRows.size(); rowIndex++) {
            List<String> row = stepRows.get(rowIndex);
            int baseline = stepTopBaseline + rowIndex * rowAdvance;
            int x = stepStartX;
            for (int col = 0; col < colWidths.length; col++) {
                String text = col < row.size() ? renderText(row.get(col)) : "";
                int textWidth = metrics.stringWidth(text);
                if (!text.isEmpty()) {
                    g2.drawString(text, x + colWidths[col] - textWidth, baseline);
                }
                x += colWidths[col] + gap;
            }
            if (rowIndex > 0 && rowLines.length > rowIndex && rowLines[rowIndex] > 0) {
                int startCol = Math.min(firstNonEmptyColumn(stepRows.get(rowIndex - 1)), firstNonEmptyColumn(stepRows.get(rowIndex)));
                int lineStart = startCol >= 0 ? stepStartX + sum(colWidths, startCol) + gap * startCol - 2 : stepStartX;
                int lineY = resolveRowSeparatorY(metrics, baseline);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(lineStart, lineY, rightEdge + 2, lineY);
            }
        }

        drawLongDivisionSign(g2, symbolX, topLineY, bottomY, rightEdge + 2);
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private LaTeXNode findStructuredNode(LaTeXNode node) {
        if (node == null) {
            return null;
        }
        if (node.getType() == LaTeXNode.Type.LONG_DIVISION) {
            return node;
        }
        if (node.getType() == LaTeXNode.Type.ARRAY
            && (isDivisionArray(node) || isDecimalArray(node) || isArithmeticArray(node))) {
            return node;
        }
        for (LaTeXNode child : node.getChildren()) {
            LaTeXNode candidate = findStructuredNode(child);
            if (candidate != null) {
                return candidate;
            }
        }
        return null;
    }

    private LaTeXNode findFirstDescendant(LaTeXNode node, LaTeXNode.Type type) {
        if (node == null) {
            return null;
        }
        if (node.getType() == type) {
            return node;
        }
        for (LaTeXNode child : node.getChildren()) {
            LaTeXNode candidate = findFirstDescendant(child, type);
            if (candidate != null) {
                return candidate;
            }
        }
        return null;
    }

    private LaTeXNode findNestedArray(LaTeXNode node, LaTeXNode exclude) {
        if (node == null) {
            return null;
        }
        if (node != exclude && node.getType() == LaTeXNode.Type.ARRAY) {
            return node;
        }
        for (LaTeXNode child : node.getChildren()) {
            LaTeXNode candidate = findNestedArray(child, exclude);
            if (candidate != null) {
                return candidate;
            }
        }
        return null;
    }

    private LaTeXNode unwrapStructuredNode(LaTeXNode latexAst) {
        if (latexAst == null) {
            return null;
        }
        if (latexAst.getType() == LaTeXNode.Type.ROOT && latexAst.getChildren().size() == 1) {
            return latexAst.getChildren().get(0);
        }
        return latexAst;
    }

    private boolean isArithmeticArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY) {
            return false;
        }
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                String text = flattenNodeText(cell);
                if (text.contains("+") || text.contains("-") || text.contains("×") || text.contains("\\times")) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean isDecimalArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY) {
            return false;
        }
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                if (".".equals(flattenNodeText(cell))) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean isDivisionArray(LaTeXNode node) {
        return node != null
            && node.getType() == LaTeXNode.Type.ARRAY
            && node.getMetadata("columnSpec") != null
            && node.getMetadata("columnSpec").contains("|");
    }

    private PreviewImage renderArithmeticArrayPreview(LaTeXNode node) {
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        List<List<String>> rows = new ArrayList<>();
        for (LaTeXNode rowNode : node.getChildren()) {
            List<String> row = new ArrayList<>();
            for (LaTeXNode cell : rowNode.getChildren()) {
                row.add(flattenNodeText(cell));
            }
            trimRight(row);
            rows.add(row);
        }
        int maxCols = rows.stream().mapToInt(List::size).max().orElse(0);
        if (maxCols == 0) {
            return null;
        }
        for (List<String> row : rows) {
            while (row.size() < maxCols) {
                row.add(0, "");
            }
        }

        int digitWidth = Math.max(metrics.stringWidth("0"), metrics.stringWidth("8"));
        int operatorWidth = Math.max(metrics.stringWidth("+"), Math.max(metrics.stringWidth("-"), metrics.stringWidth("×")));
        int[] colWidths = new int[maxCols];
        for (int col = 0; col < maxCols; col++) {
            boolean operatorColumn = false;
            int maxWidth = 0;
            for (List<String> row : rows) {
                String text = row.get(col);
                if (text.contains("+") || text.contains("-") || text.contains("×") || text.contains("\\times")) {
                    operatorColumn = true;
                }
                maxWidth = Math.max(maxWidth, metrics.stringWidth(renderText(text)));
            }
            colWidths[col] = operatorColumn
                ? Math.max(operatorWidth + 16, maxWidth + 10)
                : Math.max(digitWidth + 18, maxWidth + 12);
        }

        int gap = 10;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP;
        int totalWidth = STRUCTURED_PADDING * 2;
        for (int width : colWidths) {
            totalWidth += width;
        }
        totalWidth += gap * Math.max(maxCols - 1, 0);
        int totalHeight = STRUCTURED_PADDING * 2 + rows.size() * rowAdvance;

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int[] rowLines = parsePartitionArray(node.getMetadata("rowLines"), rows.size() + 1);
        int rightEdge = STRUCTURED_PADDING + sum(colWidths) + gap * Math.max(maxCols - 1, 0);

        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            int baseline = STRUCTURED_PADDING + metrics.getAscent() + rowIndex * rowAdvance;
            int x = STRUCTURED_PADDING;
            List<String> row = rows.get(rowIndex);
            for (int col = 0; col < maxCols; col++) {
                String text = renderText(row.get(col));
                int textWidth = metrics.stringWidth(text);
                if (!text.isEmpty()) {
                    g2.drawString(text, x + colWidths[col] - textWidth, baseline);
                }
                x += colWidths[col] + gap;
            }
            if (rowIndex > 0 && rowLines.length > rowIndex && rowLines[rowIndex] > 0) {
                int startCol = Math.min(firstNonEmptyColumn(rows.get(rowIndex - 1)), firstNonEmptyColumn(rows.get(rowIndex)));
                int lineStart = startCol >= 0 ? STRUCTURED_PADDING + sum(colWidths, startCol) + gap * startCol - 4 : STRUCTURED_PADDING;
                int lineY = resolveRowSeparatorY(metrics, baseline);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(lineStart, lineY, rightEdge + 2, lineY);
            }
        }
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private PreviewImage renderDecimalArrayPreview(LaTeXNode node) {
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        List<DecimalRow> rows = new ArrayList<>();
        for (LaTeXNode rowNode : node.getChildren()) {
            List<String> cells = new ArrayList<>();
            for (LaTeXNode cell : rowNode.getChildren()) {
                cells.add(flattenNodeText(cell));
            }
            rows.add(parseDecimalRow(cells));
        }

        FontMetrics initialMetrics = metrics;
        int signWidth = Math.max(initialMetrics.stringWidth("+"), initialMetrics.stringWidth("-")) + 10;
        int integerWidth = rows.stream().mapToInt(r -> initialMetrics.stringWidth(r.integer())).max().orElse(initialMetrics.stringWidth("0")) + 14;
        int dotWidth = Math.max(initialMetrics.stringWidth("."), 6) + 4;
        int fractionWidth = rows.stream().mapToInt(r -> initialMetrics.stringWidth(r.fraction())).max().orElse(initialMetrics.stringWidth("0")) + 14;
        int gap = 6;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP;
        int totalWidth = STRUCTURED_PADDING * 2 + signWidth + integerWidth + dotWidth + fractionWidth + gap * 3;
        int totalHeight = STRUCTURED_PADDING * 2 + rows.size() * rowAdvance;

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int[] rowLines = parsePartitionArray(node.getMetadata("rowLines"), rows.size() + 1);
        int rightEdge = totalWidth - STRUCTURED_PADDING;

        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            DecimalRow row = rows.get(rowIndex);
            int baseline = STRUCTURED_PADDING + metrics.getAscent() + rowIndex * rowAdvance;
            int x = STRUCTURED_PADDING;
            if (!row.sign().isEmpty()) {
                g2.drawString(row.sign(), x + signWidth - metrics.stringWidth(row.sign()), baseline);
            }
            x += signWidth + gap;
            if (!row.integer().isEmpty()) {
                g2.drawString(row.integer(), x + integerWidth - metrics.stringWidth(row.integer()), baseline);
            }
            x += integerWidth + gap;
            if (!row.dot().isEmpty()) {
                g2.drawString(row.dot(), x + (dotWidth - metrics.stringWidth(row.dot())) / 2, baseline);
            }
            x += dotWidth + gap;
            if (!row.fraction().isEmpty()) {
                g2.drawString(row.fraction(), x, baseline);
            }

            if (rowIndex > 0 && rowLines.length > rowIndex && rowLines[rowIndex] > 0) {
                int lineY = resolveRowSeparatorY(metrics, baseline);
                g2.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(STRUCTURED_PADDING, lineY, rightEdge, lineY);
            }
        }
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private DecimalRow parseDecimalRow(List<String> cells) {
        if (cells.size() >= 4) {
            String sign = cells.get(0);
            String integer = cells.get(1);
            String dot = cells.get(2);
            String fraction = cells.get(3);
            if ((sign == null || sign.isEmpty()) && integer != null && !integer.isEmpty()
                && (integer.charAt(0) == '+' || integer.charAt(0) == '-')) {
                sign = String.valueOf(integer.charAt(0));
                integer = integer.substring(1);
            }
            return new DecimalRow(blankToEmpty(sign), blankToEmpty(integer), blankToEmpty(dot), blankToEmpty(fraction));
        }

        String first = cells.isEmpty() ? "" : blankToEmpty(cells.get(0));
        String sign = "";
        String integer = first;
        if (!first.isEmpty() && (first.charAt(0) == '+' || first.charAt(0) == '-')) {
            sign = String.valueOf(first.charAt(0));
            integer = first.substring(1);
        }
        String dot = cells.size() > 1 ? blankToEmpty(cells.get(1)) : "";
        String fraction = cells.size() > 2 ? blankToEmpty(cells.get(2)) : "";
        return new DecimalRow(sign, integer, dot, fraction);
    }

    private String blankToEmpty(String text) {
        return text == null || text.isBlank() ? "" : text;
    }

    private RuleSpan findRuleSpan(List<RuleSpan> spans, int rowIndex) {
        for (RuleSpan span : spans) {
            if (span.boundaryIndex() == rowIndex) {
                return span;
            }
        }
        return null;
    }

    private VerticalRow findDividendRow(VerticalLayoutSpec layout) {
        for (VerticalRow row : layout.rows()) {
            if (row.kind() == VerticalLayoutSpec.RowKind.DIVIDEND) {
                return row;
            }
        }
        return null;
    }

    private List<String> trimDivisionCells(List<String> cells) {
        List<String> trimmed = new ArrayList<>();
        if (cells == null) {
            return trimmed;
        }
        for (int index = 2; index < cells.size(); index++) {
            trimmed.add(cells.get(index));
        }
        return trimmed;
    }

    private void drawRowAcrossColumns(Graphics2D g2, int startX, int[] colWidths, int gap,
                                      VerticalRow row, int baseline, FontMetrics metrics) {
        if (row == null) {
            return;
        }
        List<String> cells = trimDivisionCells(row.cells());
        int x = startX;
        for (int col = 0; col < colWidths.length; col++) {
            String text = col < cells.size() ? renderText(cells.get(col)) : "";
            int textWidth = metrics.stringWidth(text);
            if (!text.isEmpty()) {
                g2.drawString(text, x + colWidths[col] - textWidth, baseline);
            }
            x += colWidths[col] + gap;
        }
    }

    private String cellAt(List<String> cells, int index) {
        if (cells == null || index < 0 || index >= cells.size()) {
            return "";
        }
        return blankToEmpty(cells.get(index));
    }

    private PreviewImage renderDivisionArrayPreview(LaTeXNode node) {
        List<List<String>> rows = extractRowTexts(node);
        if (rows.isEmpty() || rows.get(0).size() < 2) {
            return null;
        }
        String divisor = rows.get(0).get(0);
        String dividend = rows.get(0).get(1);
        List<String> steps = new ArrayList<>();
        for (int i = 1; i < rows.size(); i++) {
            if (rows.get(i).size() > 1) {
                steps.add(rows.get(i).get(1));
            }
        }
        return renderLongDivisionLegacy(divisor, dividend, steps);
    }

    private PreviewImage renderLongDivisionPreview(LaTeXNode node) {
        VerticalLayoutSpec layout = verticalLayoutCompiler.compile(node);
        if (layout != null && layout.isLongDivision()) {
            return renderLongDivisionLayoutPreview(layout);
        }
        String divisor = flattenNodeText(node.getChildren().size() > 0 ? node.getChildren().get(0) : null);
        String quotient = flattenNodeText(node.getChildren().size() > 1 ? node.getChildren().get(1) : null);
        String dividend = flattenNodeText(node.getChildren().size() > 2 ? node.getChildren().get(2) : null);
        return renderComputedLongDivision(divisor, quotient, dividend);
    }

    private PreviewImage renderLongDivisionLegacy(String divisor, String dividend, List<String> steps) {
        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        int divisorWidth = metrics.stringWidth(divisor);
        int dividendWidth = Math.max(metrics.stringWidth(dividend), steps.stream().mapToInt(metrics::stringWidth).max().orElse(0));
        int hookGap = 12;
        int hookWidth = 18;
        int topPadding = STRUCTURED_PADDING;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP;
        int rowsBelow = Math.max(steps.size(), 1);
        int totalWidth = STRUCTURED_PADDING * 2 + divisorWidth + hookGap + hookWidth + dividendWidth + 8;
        int totalHeight = topPadding * 2 + rowAdvance * (rowsBelow + 1);

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int symbolX = STRUCTURED_PADDING + divisorWidth + hookGap;
        int dividendX = symbolX + hookWidth;
        int quotientBaseline = topPadding + metrics.getAscent();
        int dividendBaseline = quotientBaseline;
        int topLineY = dividendBaseline - metrics.getAscent() - 2;
        int bottomY = dividendBaseline + metrics.getDescent() + Math.max(steps.size(), 1) * rowAdvance - metrics.getHeight() / 2;
        int rightEdge = dividendX + dividendWidth + 4;

        g2.drawString(divisor, STRUCTURED_PADDING, dividendBaseline);
        g2.drawString(dividend, dividendX, dividendBaseline);

        int stepBaseline = dividendBaseline + rowAdvance;
        for (String step : steps) {
            g2.drawString(step, dividendX, stepBaseline);
            stepBaseline += rowAdvance;
        }

        drawLongDivisionSign(g2, symbolX, topLineY, bottomY, rightEdge);
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private PreviewImage renderComputedLongDivision(String divisor, String quotient, String dividend) {
        List<LongDivisionRow> rows = buildLongDivisionRows(divisor, dividend);
        String resolvedQuotient = blankToEmpty(quotient);
        if (resolvedQuotient.isEmpty()) {
            resolvedQuotient = computeQuotientText(divisor, dividend);
        }
        if (rows.isEmpty()) {
            return renderLongDivisionLegacy(divisor, dividend, List.of());
        }

        StructuredCanvas canvas = createStructuredCanvas();
        Graphics2D g2 = canvas.graphics();
        FontMetrics metrics = canvas.metrics();

        int slotWidth = Math.max(metrics.stringWidth("0"), metrics.stringWidth("8")) + 10;
        int columns = rows.stream().mapToInt(LongDivisionRow::endColumn).max().orElse(0) + 1;
        int divisorWidth = metrics.stringWidth(divisor);
        int quotientWidth = metrics.stringWidth(resolvedQuotient);
        int hookGap = 12;
        int hookWidth = 18;
        int rowAdvance = metrics.getHeight() + STRUCTURED_ROW_GAP;
        int dividendWidth = columns * slotWidth;
        int totalWidth = STRUCTURED_PADDING * 2 + divisorWidth + hookGap + hookWidth + Math.max(dividendWidth, quotientWidth + 8) + 12;
        int totalHeight = STRUCTURED_PADDING * 2 + rowAdvance * (rows.size() + 1);

        ensureCanvasSize(canvas, totalWidth, totalHeight);
        g2 = canvas.graphics();
        metrics = canvas.metrics();

        int symbolX = STRUCTURED_PADDING + divisorWidth + hookGap;
        int dividendX = symbolX + hookWidth;
        int quotientBaseline = STRUCTURED_PADDING + metrics.getAscent();
        int firstRowBaseline = quotientBaseline + rowAdvance;
        int topLineY = firstRowBaseline - metrics.getAscent() - 2;
        int bottomY = firstRowBaseline + (rows.size() - 1) * rowAdvance + metrics.getDescent();
        int rightEdge = dividendX + dividendWidth + 4;

        if (!resolvedQuotient.isEmpty()) {
            drawTextAcrossFixedSlots(g2, dividendX, slotWidth, columns - 1, quotientBaseline, resolvedQuotient, metrics);
        }
        g2.drawString(divisor, STRUCTURED_PADDING, firstRowBaseline);

        for (int index = 0; index < rows.size(); index++) {
            LongDivisionRow row = rows.get(index);
            int baseline = firstRowBaseline + index * rowAdvance;
            int textX = drawTextAcrossFixedSlots(g2, dividendX, slotWidth, row.endColumn(), baseline, row.text(), metrics);
            if (row.underline()) {
                int lineY = baseline + Math.max(metrics.getDescent() / 2, 3);
                g2.setStroke(new BasicStroke(2.8f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                g2.drawLine(textX - 2, lineY, textX + metrics.stringWidth(row.text()) + 2, lineY);
            }
        }

        drawLongDivisionSign(g2, symbolX, topLineY, bottomY, rightEdge);
        return toPreviewImage(canvas, totalWidth, totalHeight);
    }

    private List<LongDivisionRow> buildLongDivisionRows(String divisorText, String dividendText) {
        String divisor = blankToEmpty(divisorText);
        String dividend = blankToEmpty(dividendText);
        if (!divisor.matches("\\d+") || !dividend.matches("\\d+")) {
            return List.of();
        }

        int divisorValue;
        try {
            divisorValue = Integer.parseInt(divisor);
        } catch (NumberFormatException ignored) {
            return List.of();
        }
        if (divisorValue == 0) {
            return List.of();
        }

        List<LongDivisionRow> rows = new ArrayList<>();
        rows.add(new LongDivisionRow(dividend, dividend.length() - 1, false));

        int current = 0;
        boolean started = false;
        for (int index = 0; index < dividend.length(); index++) {
            current = current * 10 + (dividend.charAt(index) - '0');
            if (!started && current < divisorValue) {
                continue;
            }
            started = true;

            int product = (current / divisorValue) * divisorValue;
            int remainder = current - product;
            rows.add(new LongDivisionRow(String.valueOf(product), index, true));
            rows.add(new LongDivisionRow(String.valueOf(remainder), index, false));
            current = remainder;
        }
        return rows;
    }

    private String computeQuotientText(String divisorText, String dividendText) {
        String divisor = blankToEmpty(divisorText);
        String dividend = blankToEmpty(dividendText);
        if (!divisor.matches("\\d+") || !dividend.matches("\\d+")) {
            return "";
        }
        try {
            int divisorValue = Integer.parseInt(divisor);
            int dividendValue = Integer.parseInt(dividend);
            if (divisorValue == 0) {
                return "";
            }
            return String.valueOf(dividendValue / divisorValue);
        } catch (NumberFormatException ignored) {
            return "";
        }
    }

    private int alignTextToColumn(int startX, int slotWidth, int endColumn, FontMetrics metrics, String text) {
        return startX + (endColumn + 1) * slotWidth - metrics.stringWidth(text);
    }

    private int drawTextAcrossFixedSlots(Graphics2D g2, int startX, int slotWidth, int endColumn, int baseline,
                                         String text, FontMetrics metrics) {
        if (text == null || text.isEmpty()) {
            return startX;
        }
        int startColumn = Math.max(endColumn - text.length() + 1, 0);
        for (int i = 0; i < text.length(); i++) {
            String ch = String.valueOf(text.charAt(i));
            int col = startColumn + i;
            int right = startX + (col + 1) * slotWidth;
            g2.drawString(ch, right - metrics.stringWidth(ch), baseline);
        }
        return startX + (startColumn + 1) * slotWidth - metrics.stringWidth(String.valueOf(text.charAt(0)));
    }

    private void drawTextAcrossColumns(Graphics2D g2, int startX, int[] colWidths, int gap, int endColumn,
                                       int baseline, String text, FontMetrics metrics) {
        if (text == null || text.isEmpty() || colWidths.length == 0) {
            return;
        }
        int startColumn = Math.max(endColumn - text.length() + 1, 0);
        for (int i = 0; i < text.length() && startColumn + i < colWidths.length; i++) {
            String ch = String.valueOf(text.charAt(i));
            int col = startColumn + i;
            int right = startX + sum(colWidths, col + 1) + gap * col;
            g2.drawString(ch, right - metrics.stringWidth(ch), baseline);
        }
    }

    private int stepStartXForWidths(int rightEdge, int[] colWidths, int gap) {
        return rightEdge - (sum(colWidths) + gap * Math.max(colWidths.length - 1, 0));
    }

    private int resolveRowSeparatorY(FontMetrics metrics, int baseline) {
        return baseline - metrics.getAscent() - Math.max(STRUCTURED_ROW_GAP / 2, 3);
    }

    private void drawLongDivisionSign(Graphics2D g2, int symbolX, int topY, int bottomY, int rightEdge) {
        g2.setStroke(new BasicStroke(3.2f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
        g2.drawLine(symbolX, topY, rightEdge, topY);

        int hookHeight = Math.max((bottomY - topY) / 4, 14);
        Path2D.Float path = new Path2D.Float();
        path.moveTo(symbolX, topY);
        path.curveTo(symbolX - 12, topY + hookHeight / 4f, symbolX - 12, topY + hookHeight * 0.85f, symbolX, topY + hookHeight);
        path.lineTo(symbolX, bottomY);
        g2.draw(path);
    }

    private List<List<String>> extractRowTexts(LaTeXNode node) {
        List<List<String>> rows = new ArrayList<>();
        for (LaTeXNode rowNode : node.getChildren()) {
            List<String> row = new ArrayList<>();
            for (LaTeXNode cell : rowNode.getChildren()) {
                row.add(flattenNodeText(cell));
            }
            rows.add(row);
        }
        return rows;
    }

    private String flattenNodeText(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR) {
            return node.getValue() == null ? "" : node.getValue();
        }
        if (node.getType() == LaTeXNode.Type.COMMAND) {
            if ("\\times".equals(node.getValue())) {
                return "×";
            }
            return node.getValue() == null ? "" : node.getValue().replace("\\", "");
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flattenNodeText(child));
        }
        return builder.toString();
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

    private int firstNonEmptyColumn(List<String> row) {
        for (int i = 0; i < row.size(); i++) {
            if (!renderText(row.get(i)).isEmpty()) {
                return i;
            }
        }
        return -1;
    }

    private int lastNonEmptyColumn(List<String> row) {
        for (int i = row.size() - 1; i >= 0; i--) {
            if (!renderText(row.get(i)).isEmpty()) {
                return i;
            }
        }
        return 0;
    }

    private int sum(int[] values) {
        return sum(values, values.length);
    }

    private int sum(int[] values, int count) {
        int total = 0;
        for (int i = 0; i < count && i < values.length; i++) {
            total += values[i];
        }
        return total;
    }

    private void trimRight(List<String> row) {
        while (!row.isEmpty() && renderText(row.get(row.size() - 1)).isEmpty()) {
            row.remove(row.size() - 1);
        }
    }

    private String renderText(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        return "\\times".equals(text) ? "×" : text;
    }

    private StructuredCanvas createStructuredCanvas() {
        int width = 64;
        int height = 64;
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g2.setColor(Color.WHITE);
        g2.fillRect(0, 0, width, height);
        g2.setColor(Color.BLACK);
        g2.setFont(new Font("Times New Roman", Font.PLAIN, STRUCTURED_FONT_SIZE));
        return new StructuredCanvas(image, g2, g2.getFontMetrics());
    }

    private void ensureCanvasSize(StructuredCanvas canvas, int width, int height) {
        if (canvas.image().getWidth() >= width && canvas.image().getHeight() >= height) {
            canvas.graphics().setColor(Color.WHITE);
            canvas.graphics().fillRect(0, 0, canvas.image().getWidth(), canvas.image().getHeight());
            canvas.graphics().setColor(Color.BLACK);
            return;
        }
        canvas.graphics().dispose();
        BufferedImage resized = new BufferedImage(Math.max(width, 64), Math.max(height, 64), BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = resized.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g2.setColor(Color.WHITE);
        g2.fillRect(0, 0, resized.getWidth(), resized.getHeight());
        g2.setColor(Color.BLACK);
        g2.setFont(new Font("Times New Roman", Font.PLAIN, STRUCTURED_FONT_SIZE));
        canvas.reset(resized, g2, g2.getFontMetrics());
    }

    private PreviewImage toPreviewImage(StructuredCanvas canvas, int width, int height) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(canvas.image(), "png", baos);
            int displayWidth = Math.max((int) Math.ceil(width / STRUCTURED_RENDER_SCALE), 10);
            int displayHeight = Math.max((int) Math.ceil(height / STRUCTURED_RENDER_SCALE), 10);
            return new PreviewImage(baos.toByteArray(), displayWidth, displayHeight, "png", "image/png");
        } catch (IOException e) {
            return null;
        } finally {
            canvas.graphics().dispose();
        }
    }

    /**
     * 将 LaTeX 公式渲染为 PNG 字节数组（多级回退策略）。
     *
     * <p>回退链：
     * <ol>
     *   <li>MathJax HTTP 渲染（若 {@code paperword.mathjax.enabled=true}）</li>
     *   <li>dvisvgm 本地工具链（若非 mathjaxOnly 模式）</li>
     *   <li>JLaTeXMath 纯 Java 渲染（最终兜底）</li>
     * </ol>
     *
     * @param latex LaTeX 公式字符串
     * @param size  字体大小（磅）
     * @return PNG 图片字节数组
     */
    public byte[] renderToPng(String latex, float size) {
        boolean mathjaxEnabled = Boolean.parseBoolean(System.getProperty(MATHJAX_ENABLED_PROP, "true"));
        boolean mathjaxOnly = Boolean.parseBoolean(System.getProperty(MATHJAX_ONLY_PROP, "false"));

        // 第一优先级：MathJax HTTP 远程渲染（返回 SVG → 转 PNG）
        if (mathjaxEnabled) {
            byte[] mathjax = renderViaMathJaxHttp(latex);
            if (mathjax != null && mathjax.length > 0) {
                return mathjax;
            }
            // mathjaxOnly 模式下跳过 dvisvgm，直接退化到 JLaTeXMath
            if (mathjaxOnly) {
                return renderByJLatexMath(latex, size);
            }
        }

        // 第二优先级：dvisvgm 本地工具链（latex → dvi → svg → png）
        byte[] external = renderViaDvisvgm(latex, size);
        if (external != null && external.length > 0) {
            return external;
        }

        // 最终兜底：JLaTeXMath 纯 Java 渲染
        return renderByJLatexMath(latex, size);
    }

    /**
     * 使用 JLaTeXMath 库在内存中渲染 LaTeX 公式为 PNG。
     *
     * <p>JLaTeXMath 是一个纯 Java 的 LaTeX 数学公式渲染库，无需安装 TeX 发行版。
     * 以 STYLE_TEXT 模式渲染（行内公式风格），并按 {@link #RENDER_SCALE} 倍放大
     * 以获得更清晰的 Word 预览效果。</p>
     *
     * <p>使用 TYPE_INT_RGB（不透明白底）而非 TYPE_INT_ARGB（透明底），
     * 因为 Word 对 PNG 的透明度插值处理会导致公式边缘出现灰色伪影。</p>
     *
     * @param latex LaTeX 公式字符串
     * @param size  字体大小（磅）
     * @return PNG 字节数组；异常时返回占位图
     */
    private byte[] renderByJLatexMath(String latex, float size) {
        try {
            TeXFormula formula = new TeXFormula(latex);
            // 以 RENDER_SCALE 倍分辨率渲染，使公式在 Word 中缩放后仍保持清晰
            TeXIcon icon = formula.createTeXIcon(TeXConstants.STYLE_TEXT, size * RENDER_SCALE);
            icon.setInsets(new Insets(1, 1, 1, 1));
            icon.setForeground(Color.BLACK);

            int width = icon.getIconWidth();
            int height = icon.getIconHeight();
            if (width <= 0 || height <= 0) {
                log.warn("Formula produced empty image: {}", latex);
                return createPlaceholderImage(latex);
            }

            // 白色不透明背景：避免 Word 图片插值产生灰色伪影
            BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2 = image.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, width, height);
            icon.paintIcon(null, g2, 0, 0);
            g2.dispose();

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(image, "png", baos);
            return baos.toByteArray();
        } catch (Exception e) {
            log.error("Failed to render LaTeX formula: {}", latex, e);
            return createPlaceholderImage(latex);
        }
    }

    /** EMF 矢量渲染的缩放因子，适当提高可减少 Word 首次显示时的粗糙感 */
    private static final float EMF_RENDER_SCALE = 2.0f;

    /**
     * 将 LaTeX 公式渲染为 EMF（Enhanced Metafile）矢量图。
     *
     * <p><b>EMF 格式的优势：</b>EMF 是 Windows 原生支持的矢量图格式，
     * Word 能直接解析并以任意缩放比例无损显示。理论上比 PNG 预览更优。</p>
     *
     * <p><b>2 倍渲染策略：</b>以 2 倍 scale 渲染到 EMF 坐标系，确保分数线、
     * 根号上横线等细线元素在矢量输出中具有足够的坐标精度。
     * VML shape 的尺寸设为 1 倍显示大小，Word 将 EMF 内容缩小显示，
     * 从而保留所有细节。</p>
     *
     * <p><b>TEXT_AS_SHAPES：</b>必须启用此选项将文本转为矢量路径。
     * JLaTeXMath 使用 Computer Modern 等自定义字体，Windows 默认未安装这些字体。
     * 如果不转为路径，Word 会用系统字体替代，导致符号显示错误
     * （例如 π 显示为 ¼，× 显示为 £，α 显示为 ®）。</p>
     *
     * @param latex LaTeX 公式字符串
     * @param size  字体大小（磅）
     * @return EMF 预览图数据；渲染失败返回 null
     */
    private PreviewImage renderToEmf(String latex, float size) {
        try {
            TeXIcon icon = createInlinePreviewIcon(latex, size * EMF_RENDER_SCALE, EMF_RENDER_SCALE);
            icon.setForeground(Color.BLACK);

            int emfWidth = Math.max(icon.getIconWidth(), 10);
            int emfHeight = Math.max(icon.getIconHeight(), 10);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            EMFGraphics2D emf = new EMFGraphics2D(baos, new Dimension(emfWidth, emfHeight));

            // 关键配置：将文本渲染为矢量路径而非字体引用
            // JLaTeXMath 使用的 Computer Modern 字体在 Windows 上不可用，
            // 不转为路径会导致 Word 用系统字体替代，出现错误字形
            java.util.Properties emfProps = new java.util.Properties();
            emfProps.setProperty(
                "org.freehep.graphicsio.AbstractVectorGraphicsIO.TEXT_AS_SHAPES",
                "true");
            emf.setProperties(emfProps);

            // 执行 EMF 导出：白底 + 黑色公式
            emf.startExport();
            emf.setBackground(Color.WHITE);
            emf.clearRect(0, 0, emfWidth, emfHeight);
            emf.setColor(Color.BLACK);
            icon.paintIcon(null, emf, 0, 0);
            emf.endExport();

            byte[] data = baos.toByteArray();
            if (data.length == 0) {
                return null;
            }
            // 报告 1 倍显示尺寸：Word 将 2 倍精度的 EMF 缩放到此尺寸显示
            int displayWidth = Math.max((int)(emfWidth / EMF_RENDER_SCALE), 10);
            int displayHeight = Math.max((int)(emfHeight / EMF_RENDER_SCALE), 10);
            return new PreviewImage(data, displayWidth, displayHeight, "emf", "image/x-emf");
        } catch (Throwable e) {
            log.debug("EMF preview render failed, fallback to PNG: {}", latex, e);
            return null;
        }
    }

    private TeXIcon createInlinePreviewIcon(String latex, float scaledSize, float renderScale) {
        TeXFormula formula = new TeXFormula(latex);
        int style = usesStackedPreviewStyle(latex) ? TeXConstants.STYLE_DISPLAY : TeXConstants.STYLE_TEXT;
        TeXIcon icon = formula.createTeXIcon(style, scaledSize);
        int verticalInset = usesStackedPreviewStyle(latex)
            ? Math.max((int)(2 * renderScale), 2)
            : Math.max((int)(1 * renderScale), 1);
        icon.setInsets(new Insets(
            verticalInset,
            Math.max((int)(0.5f * renderScale), 0),
            verticalInset,
            Math.max((int)(0.5f * renderScale), 0)));
        return icon;
    }

    private Insets createOlePreviewInsets(String latex, float renderScale) {
        boolean stacked = usesStackedPreviewStyle(latex);
        int verticalInset = stacked
            ? Math.max((int)(3.5f * renderScale), 3)
            : Math.max((int)(2.0f * renderScale), 2);
        int horizontalInset = stacked
            ? Math.max((int)(2.5f * renderScale), 2)
            : Math.max((int)(1.5f * renderScale), 1);
        return new Insets(verticalInset, horizontalInset, verticalInset, horizontalInset);
    }

    private boolean usesStackedPreviewStyle(String latex) {
        if (latex == null) {
            return false;
        }
        return latex.contains("\\frac")
            || latex.contains("\\dfrac")
            || latex.contains("\\cfrac")
            || latex.contains("\\sqrt")
            || latex.contains("\\sum")
            || latex.contains("\\int")
            || latex.contains("\\prod");
    }

    /**
     * 通过系统安装的 latex + dvisvgm 工具链渲染公式。
     *
     * <p>渲染管线：LaTeX 源码 → .tex 文件 → latex 编译为 .dvi → dvisvgm 转为 SVG → Batik 转为 PNG。</p>
     *
     * <p>此方法需要系统安装 TeX 发行版（如 TeX Live、MiKTeX）。
     * 如果首次调用时发现外部命令不可用（IOException），则设置
     * {@link #externalToolUnavailable} 标志，后续调用直接跳过。</p>
     *
     * <p>所有临时文件在 finally 块中清理。</p>
     *
     * @param latex LaTeX 公式字符串
     * @return PNG 字节数组；工具不可用或渲染失败返回 null
     */
    private byte[] renderViaDvisvgm(String latex, float size) {
        byte[] svg = renderSvgViaDvisvgm(latex, size, List.of(), true, null);
        if (svg == null || svg.length == 0) {
            return null;
        }
        try {
            return svgToPng(svg);
        } catch (Exception e) {
            log.debug("SVG -> PNG transcode failed, fallback to next renderer: {}", latex, e);
            return null;
        }
    }

    private PreviewImage renderVectorPreviewViaExternalTools(String latex, float size) {
        if (externalToolUnavailable || externalVectorToolUnavailable) {
            return null;
        }
        Path tempDir = null;
        try {
            tempDir = Files.createTempDirectory("paperword-vector-preview-");
            byte[] svg = renderSvgViaDvisvgm(latex, size, List.of(), true, tempDir);
            if (svg == null || svg.length == 0) {
                return null;
            }

            Path svgFile = tempDir.resolve("eq.svg");
            if (!Files.exists(svgFile)) {
                Files.write(svgFile, svg);
            }
            Path emfFile = tempDir.resolve("eq.emf");
            if (!convertSvgToEmf(svgFile, emfFile, Integer.getInteger(RENDER_TIMEOUT_PROP, DEFAULT_TIMEOUT_SECONDS))) {
                return null;
            }

            byte[] emf = Files.readAllBytes(emfFile);
            if (emf.length == 0) {
                return null;
            }

            SvgDimensions dimensions = extractSvgDisplayDimensions(svg);
            int widthPx = Math.max((int) Math.ceil(dimensions.widthPt() * PX_PER_PT), 10);
            int heightPx = Math.max((int) Math.ceil(dimensions.heightPt() * PX_PER_PT), 10);
            return new PreviewImage(emf, widthPx, heightPx, "emf", "image/x-emf");
        } catch (IOException e) {
            if (isCommandUnavailable(e)) {
                externalVectorToolUnavailable = true;
            }
            log.debug("External SVG -> EMF preview failed, fallback to internal EMF: {}", latex, e);
            return null;
        } catch (Exception e) {
            log.debug("External vector preview failed, fallback to internal EMF: {}", latex, e);
            return null;
        } finally {
            if (tempDir != null) {
                deleteQuietly(tempDir);
            }
        }
    }

    private byte[] renderSvgViaDvisvgm(String latex, float size) {
        return renderSvgViaDvisvgm(latex, size, List.of(), true, null);
    }

    private byte[] renderSvgViaDvisvgm(String latex, float size, Path providedTempDir) {
        return renderSvgViaDvisvgm(latex, size, List.of(), true, providedTempDir);
    }

    private byte[] renderViaDvisvgm(String latex, float size, List<String> extraPackages, boolean wrapDisplayMath) {
        byte[] svg = renderSvgViaDvisvgm(latex, size, extraPackages, wrapDisplayMath, null);
        if (svg == null || svg.length == 0) {
            return null;
        }
        try {
            return svgToPng(svg);
        } catch (Exception e) {
            log.debug("SVG -> PNG transcode failed, fallback to next renderer: {}", latex, e);
            return null;
        }
    }

    private byte[] renderSvgViaDvisvgm(String latex, float size, List<String> extraPackages, boolean wrapDisplayMath,
                                       Path providedTempDir) {
        // 外部工具已标记为不可用，直接跳过
        if (externalToolUnavailable) {
            return null;
        }

        Path tempDir = providedTempDir;
        boolean ownsTempDir = false;
        try {
            // 创建临时目录存放编译中间文件
            if (tempDir == null) {
                tempDir = Files.createTempDirectory("paperword-latex-");
                ownsTempDir = true;
            }
            Path texFile = tempDir.resolve("eq.tex");
            Path dviFile = tempDir.resolve("eq.dvi");
            Path svgFile = tempDir.resolve("eq.svg");

            // 写入完整的 LaTeX 文档（包含 amsmath/amssymb 宏包）
            Files.writeString(texFile, buildLatexDocument(latex, size, extraPackages, wrapDisplayMath), StandardCharsets.UTF_8);

            String latexCmd = resolveLatexCommand();
            String dvisvgmCmd = resolveDvisvgmCommand();
            int timeoutSeconds = Integer.getInteger(RENDER_TIMEOUT_PROP, DEFAULT_TIMEOUT_SECONDS);

            // 步骤 1：latex 编译 .tex → .dvi
            // -interaction=nonstopmode：遇到错误不暂停等待输入
            // -halt-on-error：遇到错误立即停止
            // -no-shell-escape：禁用外部命令执行（安全考虑）
            CommandResult latexResult = runCommand(
                List.of(
                    latexCmd,
                    "-interaction=nonstopmode",
                    "-halt-on-error",
                    "-no-shell-escape",
                    "-output-directory=" + tempDir.toAbsolutePath(),
                    texFile.toAbsolutePath().toString()
                ),
                tempDir,
                timeoutSeconds
            );
            if (latexResult.exitCode != 0 || !Files.exists(dviFile)) {
                if (requiresXlopPackage(extraPackages)) {
                    // 只要最小 xlop 文档无法编译完成，就视为当前环境的官方长除法链不可用。
                    xlopToolchainUnavailable = true;
                }
                log.debug("latex render failed (code={}): {}", latexResult.exitCode, latexResult.outputText());
                return null;
            }

            // 步骤 2：dvisvgm 将 .dvi 转为 SVG
            // --no-fonts：将文本转为路径（避免字体依赖问题）
            // --exact-bbox：精确计算边界框
            // --precision=8：进一步提高坐标精度，减少 Office 首次显示时的粗糙感
            CommandResult svgResult = runCommand(
                List.of(
                    dvisvgmCmd,
                    "--verbosity=0",
                    "--exact-bbox",
                    "--no-fonts",
                    "--precision=8",
                    "-o",
                    svgFile.toAbsolutePath().toString(),
                    dviFile.toAbsolutePath().toString()
                ),
                tempDir,
                timeoutSeconds
            );
            if (svgResult.exitCode != 0 || !Files.exists(svgFile)) {
                if (requiresXlopPackage(extraPackages)) {
                    // dvisvgm 失败同样意味着当前机器无法完成 xlop -> SVG 的官方渲染链。
                    xlopToolchainUnavailable = true;
                }
                log.debug("dvisvgm render failed (code={}): {}", svgResult.exitCode, svgResult.outputText());
                return null;
            }
            return Files.readAllBytes(svgFile);
        } catch (IOException e) {
            // 命令不存在或执行环境异常：标记外部工具不可用，后续调用直接跳过
            if (isCommandUnavailable(e)) {
                externalToolUnavailable = true;
            }
            log.warn("External LaTeX/SVG tools unavailable, fallback to JLaTeXMath: {}", e.getMessage());
            return null;
        } catch (Exception e) {
            log.debug("External LaTeX->SVG render failed, fallback to JLaTeXMath: {}", latex, e);
            return null;
        } finally {
            // 清理临时目录及其中所有文件
            if (ownsTempDir && tempDir != null) {
                deleteQuietly(tempDir);
            }
        }
    }

    private boolean convertSvgToEmf(Path svgFile, Path emfFile, int timeoutSeconds) throws IOException, InterruptedException {
        String inkscapeCmd = System.getProperty(INKSCAPE_CMD_PROP, "inkscape");
        CommandResult result = runCommand(
            List.of(
                inkscapeCmd,
                svgFile.toAbsolutePath().toString(),
                "--export-filename=" + emfFile.toAbsolutePath()
            ),
            svgFile.getParent(),
            timeoutSeconds
        );
        if (result.exitCode != 0 || !Files.exists(emfFile)) {
            log.debug("Inkscape SVG -> EMF failed (code={}): {}", result.exitCode, result.outputText());
            return false;
        }
        return true;
    }

    private String resolveLatexCommand() {
        return resolveExternalCommand(LATEX_CMD_PROP, "latex.exe", "latex");
    }

    private String resolveDvisvgmCommand() {
        return resolveExternalCommand(DVISVGM_CMD_PROP, "dvisvgm.exe", "dvisvgm");
    }

    private String resolveExternalCommand(String propertyName, String windowsExecutable, String fallbackCommand) {
        String configured = System.getProperty(propertyName);
        if (configured != null && !configured.isBlank()) {
            return configured;
        }

        // Windows 下优先探测 MiKTeX 用户目录，避免新装后 PATH 尚未刷新导致找不到命令。
        if (!WINDOWS_MIKTEX_BIN.isBlank()) {
            Path candidate = Path.of(WINDOWS_MIKTEX_BIN, windowsExecutable);
            if (Files.exists(candidate)) {
                return candidate.toString();
            }
        }
        return fallbackCommand;
    }

    /**
     * 通过 MathJax HTTP 渲染服务将 LaTeX 公式转为 PNG。
     *
     * <p>调用远程 MathJax 渲染端点（默认 https://math.vercel.app/），
     * 获取 SVG 格式的渲染结果，再通过 {@link #svgToPng} 转为 PNG。</p>
     *
     * <p>端点 URL 支持两种模式：
     * <ul>
     *   <li>包含 {@code {latex}} 占位符：直接替换</li>
     *   <li>不含占位符：追加 {@code ?from=} 或 {@code &from=} 查询参数</li>
     * </ul>
     *
     * @param latex LaTeX 公式字符串
     * @return PNG 字节数组；网络异常或服务不可用时返回 null
     */
    private PreviewImage renderPreviewViaMathJaxHttp(String latex) {
        try {
            byte[] svg = fetchMathJaxSvg(latex);
            if (svg == null || svg.length == 0) {
                return null;
            }
            byte[] png = svgToPng(svg);
            SvgDimensions dimensions = extractSvgDisplayDimensions(svg);
            int widthPx = Math.max((int) Math.ceil(dimensions.widthPt() * PX_PER_PT), 10);
            int heightPx = Math.max((int) Math.ceil(dimensions.heightPt() * PX_PER_PT), 10);
            return new PreviewImage(png, widthPx, heightPx, "png", "image/png");
        } catch (Exception e) {
            log.debug("MathJax preview render failed, fallback to local preview: {}", latex, e);
            return null;
        }
    }

    private byte[] renderViaMathJaxHttp(String latex) {
        try {
            byte[] svg = fetchMathJaxSvg(latex);
            if (svg == null || svg.length == 0) {
                return null;
            }
            return svgToPng(svg);
        } catch (Exception e) {
            log.debug("MathJax HTTP render failed, fallback to next renderer: {}", latex, e);
            return null;
        }
    }

    private byte[] fetchMathJaxSvg(String latex) {
        try {
            String endpoint = System.getProperty(MATHJAX_ENDPOINT_PROP, "https://math.vercel.app/");
            String encoded = URLEncoder.encode(latex, StandardCharsets.UTF_8);
            // 根据端点 URL 格式构造请求地址
            String url;
            if (endpoint.contains("{latex}")) {
                url = endpoint.replace("{latex}", encoded);
            } else {
                String sep = endpoint.contains("?") ? "&" : "?";
                url = endpoint + sep + "from=" + encoded;
            }

            int timeoutSeconds = Integer.getInteger(MATHJAX_TIMEOUT_PROP, DEFAULT_MATHJAX_TIMEOUT_SECONDS);
            HttpRequest req = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .timeout(java.time.Duration.ofSeconds(timeoutSeconds))
                .header("Accept", "image/svg+xml,text/plain,*/*")
                .GET()
                .build();

            HttpResponse<byte[]> resp = httpClient.send(req, HttpResponse.BodyHandlers.ofByteArray());
            if (resp.statusCode() / 100 != 2 || resp.body() == null || resp.body().length == 0) {
                return null;
            }
            log.debug("Rendered by MathJax HTTP: {}", latex);
            return resp.body();
        } catch (Exception e) {
            log.debug("MathJax HTTP fetch failed, fallback to next renderer: {}", latex, e);
            return null;
        }
    }

    /**
     * 使用 Apache Batik 将 SVG 字节转为 PNG 字节。
     *
     * <p>Batik 的 PNGTranscoder 将 SVG XML 解析并光栅化为 PNG 位图，
     * 背景设为白色以匹配 Word 文档的白色页面。</p>
     *
     * @param svgBytes SVG 格式的字节数据
     * @return PNG 字节数组
     * @throws Exception SVG 解析或转码失败时抛出
     */
    private byte[] svgToPng(byte[] svgBytes) throws Exception {
        PNGTranscoder transcoder = new PNGTranscoder();
        transcoder.addTranscodingHint(PNGTranscoder.KEY_BACKGROUND_COLOR, Color.WHITE);

        TranscoderInput input = new TranscoderInput(new StringReader(new String(svgBytes, StandardCharsets.UTF_8)));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        TranscoderOutput output = new TranscoderOutput(baos);
        transcoder.transcode(input, output);
        return baos.toByteArray();
    }

    /**
     * 执行外部命令并等待完成。
     *
     * @param command        命令及参数列表
     * @param workDir        工作目录
     * @param timeoutSeconds 超时秒数，超时后强制终止进程
     * @return 命令执行结果（退出码 + stdout/stderr 输出）
     * @throws IOException          命令启动失败（如命令不存在）
     * @throws InterruptedException 等待过程被中断
     */
    private CommandResult runCommand(List<String> command, Path workDir, int timeoutSeconds)
        throws IOException, InterruptedException {
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.directory(workDir.toFile());
        pb.redirectErrorStream(true); // 合并 stdout 和 stderr

        Process process = pb.start();
        byte[] output = process.getInputStream().readAllBytes();
        boolean finished = process.waitFor(timeoutSeconds, TimeUnit.SECONDS);
        if (!finished) {
            process.destroyForcibly(); // 超时强制终止
            return new CommandResult(-1, output);
        }
        return new CommandResult(process.exitValue(), output);
    }

    /**
     * 构建完整的 LaTeX 文档源码。
     *
     * <p>将公式包裹在 article 文档类中，引入 amsmath 和 amssymb 宏包
     * 以支持常用数学符号和环境，使用 display math 模式（\[ \]）渲染。</p>
     *
     * @param latex LaTeX 公式内容
     * @return 完整的 .tex 文档字符串
     */
    private String buildLatexDocument(String latex, float size) {
        return buildLatexDocument(latex, size, List.of(), true);
    }

    private String buildLatexDocument(String latex, float size, List<String> extraPackages, boolean wrapDisplayMath) {
        String sizePt = String.format(Locale.ROOT, "%.1f", Math.max(size, 8f));
        String baselineSkipPt = String.format(Locale.ROOT, "%.1f", Math.max(size + 2f, 10f));
        StringBuilder packageBuilder = new StringBuilder("\\usepackage{amsmath,amssymb}\n");
        for (String extraPackage : extraPackages) {
            if (extraPackage != null && !extraPackage.isBlank()) {
                packageBuilder.append(extraPackage).append('\n');
            }
        }
        String body = wrapDisplayMath
            ? """
            \\[
            %s
            \\]
            """.formatted(latex)
            : latex;
        return """
            \\documentclass[12pt]{article}
            %s
            \\pagestyle{empty}
            \\begin{document}
            \\fontsize{%s}{%s}\\selectfont
            %s
            \\end{document}
            """.formatted(packageBuilder, sizePt, baselineSkipPt, body);
    }

    /**
     * 静默删除目录及其全部内容。
     *
     * <p>使用逆序遍历（文件优先于目录）确保删除顺序正确，忽略所有 IO 异常。</p>
     *
     * @param dir 要删除的目录路径
     */
    private void deleteQuietly(Path dir) {
        try (var stream = Files.walk(dir)) {
            stream.sorted(Comparator.reverseOrder())
                .forEach(path -> {
                    try {
                        Files.deleteIfExists(path);
                    } catch (IOException ignored) {
                    }
                });
        } catch (IOException ignored) {
        }
    }

    /**
     * 外部命令执行结果记录。
     *
     * @param exitCode 进程退出码（0 = 成功，-1 = 超时）
     * @param output   stdout + stderr 合并输出的字节数据
     */
    private record CommandResult(int exitCode, byte[] output) {
        /** 将输出字节转为 UTF-8 字符串（便于日志记录） */
        private String outputText() {
            return new String(output, StandardCharsets.UTF_8);
        }
    }

    static SvgDimensions extractSvgDisplayDimensions(byte[] svgBytes) {
        String svg = new String(svgBytes, StandardCharsets.UTF_8);
        Matcher matcher = Pattern.compile(
            "<svg[^>]*\\bwidth=['\"]([^'\"]+)['\"][^>]*\\bheight=['\"]([^'\"]+)['\"]",
            Pattern.CASE_INSENSITIVE | Pattern.DOTALL)
            .matcher(svg);
        if (!matcher.find()) {
            return new SvgDimensions(12.0f, 12.0f);
        }
        return new SvgDimensions(parseSvgLengthToPt(matcher.group(1)), parseSvgLengthToPt(matcher.group(2)));
    }

    private static float parseSvgLengthToPt(String value) {
        Matcher matcher = Pattern.compile("([0-9.]+)\\s*([a-zA-Z]*)").matcher(value == null ? "" : value.trim());
        if (!matcher.matches()) {
            return 12.0f;
        }
        float number = Float.parseFloat(matcher.group(1));
        String unit = matcher.group(2).toLowerCase(Locale.ROOT);
        return switch (unit) {
            case "", "px" -> number / PX_PER_PT;
            case "pt" -> number;
            case "in" -> number * 72.0f;
            case "mm" -> number * 72.0f / 25.4f;
            case "cm" -> number * 72.0f / 2.54f;
            default -> number;
        };
    }

    private boolean isCommandUnavailable(IOException exception) {
        String message = exception.getMessage();
        if (message == null) {
            return false;
        }
        String normalized = message.toLowerCase(Locale.ROOT);
        return normalized.contains("cannot run program")
            || normalized.contains("createprocess error=2")
            || normalized.contains("no such file");
    }

    private boolean requiresXlopPackage(List<String> extraPackages) {
        return extraPackages.stream()
            .filter(extraPackage -> extraPackage != null && !extraPackage.isBlank())
            .anyMatch(extraPackage -> extraPackage.contains("{xlop}"));
    }

    private record LongDivisionOperands(String divisor, String dividend) {
    }

    record SvgDimensions(float widthPt, float heightPt) {
    }

    /**
     * 将 LaTeX 公式渲染为 BufferedImage 对象。
     *
     * <p>使用 STYLE_DISPLAY 模式（行间公式风格）渲染，返回 ARGB 格式的位图。
     * 此方法主要供 EMF 转换等需要 BufferedImage 输入的场景调用。</p>
     *
     * @param latex LaTeX 公式字符串
     * @param size  字体大小（磅）
     * @return 渲染后的 BufferedImage；异常时返回 20×20 的空白图
     */
    public BufferedImage renderToImage(String latex, float size) {
        try {
            TeXFormula formula = new TeXFormula(latex);
            TeXIcon icon = formula.createTeXIcon(TeXConstants.STYLE_DISPLAY, size);

            int width = Math.max(icon.getIconWidth(), 1);
            int height = Math.max(icon.getIconHeight(), 1);

            BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2 = image.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, width, height);
            icon.paintIcon(null, g2, 0, 0);
            g2.dispose();

            return image;
        } catch (Exception e) {
            log.error("Failed to render LaTeX to image: {}", latex, e);
            return new BufferedImage(20, 20, BufferedImage.TYPE_INT_ARGB);
        }
    }

    /**
     * 创建错误占位图。
     *
     * <p>当公式渲染完全失败时，生成一个 100×30 的红色 "[formula]" 文本占位图，
     * 提示用户此处应有一个公式但渲染失败了。</p>
     *
     * @param latex 原始 LaTeX 字符串（当前未在占位图中使用，保留供扩展）
     * @return 占位 PNG 字节数组；极端情况下返回空数组
     */
    private byte[] createPlaceholderImage(String latex) {
        try {
            BufferedImage img = new BufferedImage(100, 30, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2 = img.createGraphics();
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, 100, 30);
            g2.setColor(Color.RED);
            g2.setFont(new Font("Arial", Font.PLAIN, 10));
            g2.drawString("[formula]", 5, 20);
            g2.dispose();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(img, "png", baos);
            return baos.toByteArray();
        } catch (IOException e) {
            return new byte[0];
        }
    }

    private record DecimalRow(String sign, String integer, String dot, String fraction) {
    }

    private record LongDivisionRow(String text, int endColumn, boolean underline) {
    }

    private static final class StructuredCanvas {
        private BufferedImage image;
        private Graphics2D graphics;
        private FontMetrics metrics;

        private StructuredCanvas(BufferedImage image, Graphics2D graphics, FontMetrics metrics) {
            this.image = image;
            this.graphics = graphics;
            this.metrics = metrics;
        }

        private BufferedImage image() {
            return image;
        }

        private Graphics2D graphics() {
            return graphics;
        }

        private FontMetrics metrics() {
            return metrics;
        }

        private void reset(BufferedImage image, Graphics2D graphics, FontMetrics metrics) {
            this.image = image;
            this.graphics = graphics;
            this.metrics = metrics;
        }
    }
}
