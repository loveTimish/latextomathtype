package com.lz.paperword.core.render;

import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.PNGTranscoder;
import org.scilab.forge.jlatexmath.TeXConstants;
import org.scilab.forge.jlatexmath.TeXFormula;
import org.scilab.forge.jlatexmath.TeXIcon;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * LaTeX 公式图片渲染器。
 *
 * <p>当前版本统一以原生 TeX 工具链作为预览图首选来源：
 * latex → dvi → dvisvgm → SVG → PNG。这样 OLE 预览图、普通图片模式和
 * 直接导出的 PNG 都走同一条链路，避免不同渲染器造成的字体、边距和比例漂移。</p>
 *
 * <p>仅当外部 TeX 工具不可用或渲染失败时，才回退到 JLaTeXMath 生成兜底位图。</p>
 */
public class LaTeXImageRenderer {

    private static final Logger log = LoggerFactory.getLogger(LaTeXImageRenderer.class);

    /** 默认公式字体大小（磅），对应 Word 中正文公式的标准尺寸。 */
    private static final float DEFAULT_SIZE = 13f;

    /** JLaTeXMath 兜底通道的渲染缩放因子，用于提高位图清晰度。 */
    private static final float RENDER_SCALE = 2.0f;

    /** 原生 TeX 预览图转 PNG 时的放大倍率，只提高底图分辨率，不改变文档显示尺寸。 */
    private static final float PNG_OUTPUT_SCALE = 2.0f;

    /** 原生 TeX 生成 OLE 预览图时的字号微调，尽量贴近行内 MathType 占位。 */
    private static final float OLE_PREVIEW_SIZE = 12f;

    /** 系统属性：latex 命令路径。 */
    private static final String LATEX_CMD_PROP = "paperword.latex.command";
    /** 系统属性：dvisvgm 命令路径。 */
    private static final String DVISVGM_CMD_PROP = "paperword.dvisvgm.command";
    /** 系统属性：外部命令超时秒数。 */
    private static final String RENDER_TIMEOUT_PROP = "paperword.latex.timeout.seconds";
    /** 外部命令默认超时秒数。 */
    private static final int DEFAULT_TIMEOUT_SECONDS = 20;
    /** 像素到磅的换算比例。 */
    private static final float PX_PER_PT = 1.0f / 0.75f;
    /** 显式长除法命令提取模式。 */
    private static final Pattern LONG_DIVISION_COMMAND_PATTERN =
        Pattern.compile("\\\\longdiv(?:\\[([^\\]]*)])?\\{([^{}]+)}\\{([^{}]+)}");

    /** 标记外部工具是否不可用，避免每个公式都重复探测失败。 */
    private volatile boolean externalToolUnavailable = false;

    /**
     * 预览图数据记录。
     *
     * @param data        图片字节
     * @param widthPx     显示宽度（像素）
     * @param heightPx    显示高度（像素）
     * @param extension   扩展名
     * @param contentType MIME 类型
     * @param placeholder 是否为占位图
     */
    public record PreviewImage(
        byte[] data,
        int widthPx,
        int heightPx,
        String extension,
        String contentType,
        boolean placeholder
    ) {}

    /**
     * 使用默认字号渲染 PNG。
     *
     * @param latex LaTeX 公式源码
     * @return PNG 字节数组
     */
    public byte[] renderToPng(String latex) {
        return renderToPng(latex, DEFAULT_SIZE);
    }

    /**
     * 为 OLE 对象生成预览图。
     *
     * <p>优先走原生 TeX 渲染，这样与普通图片模式保持同源。失败后再回退到
     * JLaTeXMath，避免 OLE 直接退化成原始文本。</p>
     *
     * @param latex LaTeX 公式源码
     * @return 预览图数据；失败返回 null
     */
    public PreviewImage renderForOlePreview(String latex) {
        PreviewImage preview = renderPreviewViaTeX(latex, OLE_PREVIEW_SIZE);
        if (preview != null) {
            return preview;
        }
        preview = renderPreviewViaJLatexMath(latex, OLE_PREVIEW_SIZE);
        if (preview != null) {
            return preview;
        }
        byte[] placeholder = createPlaceholderImage(latex);
        if (placeholder.length == 0) {
            return null;
        }
        return new PreviewImage(placeholder, 100, 30, "png", "image/png", true);
    }

    /**
     * 为 Word 普通图片模式生成预览图。
     *
     * @param latex LaTeX 公式源码
     * @return 预览图数据；失败返回 null
     */
    public PreviewImage renderForWordImage(String latex) {
        PreviewImage preview = renderPreviewViaTeX(latex, DEFAULT_SIZE);
        if (preview != null) {
            return preview;
        }
        return renderPreviewViaJLatexMath(latex, DEFAULT_SIZE);
    }

    /**
     * 将 LaTeX 渲染为 PNG 字节数组。
     *
     * <p>主链路固定为原生 TeX，本地工具链失败时才回退到 JLaTeXMath。</p>
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return PNG 字节数组
     */
    public byte[] renderToPng(String latex, float size) {
        String localRenderLatex = normalizeLatexForLocalRender(latex);
        byte[] external = renderViaDvisvgm(localRenderLatex, size);
        if (external != null && external.length > 0) {
            return external;
        }
        return renderByJLatexMath(localRenderLatex, size);
    }

    /**
     * 统一的原生 TeX 预览图入口。
     *
     * <p>这里直接使用 SVG 的宽高元数据计算显示尺寸，避免再依赖位图反推尺寸，
     * 从而让 OLE 预览和普通图片模式共享完全一致的度量基准。</p>
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return 预览图数据；失败返回 null
     */
    private PreviewImage renderPreviewViaTeX(String latex, float size) {
        String localRenderLatex = normalizeLatexForLocalRender(latex);
        byte[] svg = renderSvgViaDvisvgm(localRenderLatex, size);
        if (svg == null || svg.length == 0) {
            return null;
        }
        try {
            SvgDimensions dimensions = extractSvgDisplayDimensions(svg);
            int widthPx = Math.max((int) Math.ceil(dimensions.widthPt() * PX_PER_PT), 10);
            int heightPx = Math.max((int) Math.ceil(dimensions.heightPt() * PX_PER_PT), 10);
            int renderWidthPx = Math.max((int) Math.ceil(widthPx * PNG_OUTPUT_SCALE), widthPx);
            int renderHeightPx = Math.max((int) Math.ceil(heightPx * PNG_OUTPUT_SCALE), heightPx);
            byte[] pngData = svgToPng(svg, renderWidthPx, renderHeightPx);
            if (pngData == null || pngData.length == 0) {
                return null;
            }
            return new PreviewImage(pngData, widthPx, heightPx, "png", "image/png", false);
        } catch (Exception e) {
            log.debug("Native TeX preview render failed, fallback to JLaTeXMath: {}", latex, e);
            return null;
        }
    }

    /**
     * JLaTeXMath 兜底预览图入口。
     *
     * <p>这里只在原生 TeX 通道失败时使用，因此不再承担主链路渲染职责。</p>
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return 预览图数据；失败返回 null
     */
    private PreviewImage renderPreviewViaJLatexMath(String latex, float size) {
        try {
            byte[] pngData = renderByJLatexMath(normalizeLatexForLocalRender(latex), size);
            if (pngData == null || pngData.length == 0) {
                return null;
            }
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(pngData));
            if (image == null) {
                return null;
            }
            return new PreviewImage(
                pngData,
                Math.max(image.getWidth(), 10),
                Math.max(image.getHeight(), 10),
                "png",
                "image/png",
                isPlaceholderImage(image)
            );
        } catch (Exception e) {
            log.debug("JLaTeXMath preview render failed: {}", latex, e);
            return null;
        }
    }

    /**
     * 占位图尺寸固定，统一在这里识别，避免误判为真实公式。
     *
     * @param image 已解码图片
     * @return 是否为占位图
     */
    private boolean isPlaceholderImage(BufferedImage image) {
        return image != null && image.getWidth() == 100 && image.getHeight() == 30;
    }

    /**
     * 识别当前项目里会触发长除法图片分流的源码形式。
     *
     * @param latex LaTeX 公式源码
     * @return 是否包含长除法语法
     */
    private boolean containsLongDivisionLatex(String latex) {
        return latex != null
            && (latex.contains("\\enclose{longdiv}")
            || latex.contains("\\enclose{longdiv}{")
            || latex.contains("\\longdiv"));
    }

    /**
     * 对原生 TeX 通道做轻量标准化，主要处理项目里的长除法兼容写法。
     *
     * @param latex LaTeX 公式源码
     * @return 可交给本地 TeX 的公式源码
     */
    private String normalizeLatexForLocalRender(String latex) {
        if (latex == null || latex.isBlank()) {
            return latex;
        }
        String normalized = latex.replaceAll("\\\\kern\\s*[-+]?\\d*\\.?\\d+[a-zA-Z]+", "");
        String compositeLongDivision = replaceEmbeddedLongDivisionHeader(normalized);
        if (compositeLongDivision != null) {
            return compositeLongDivision;
        }
        String expandedLongDivision = expandLongDivisionPreview(normalized);
        if (expandedLongDivision != null) {
            return expandedLongDivision;
        }
        if (!containsLongDivisionLatex(normalized)) {
            return normalized;
        }
        return Pattern.compile("\\\\enclose\\{longdiv\\}\\{([^{}]+)}")
            .matcher(normalized)
            .replaceAll("\\\\big)\\\\overline{$1}");
    }

    private String replaceEmbeddedLongDivisionHeader(String latex) {
        Matcher matcher = LONG_DIVISION_COMMAND_PATTERN.matcher(latex);
        if (!matcher.find()) {
            return null;
        }
        String quotient = matcher.group(1) == null ? "" : matcher.group(1).trim();
        String divisor = matcher.group(2) == null ? "" : matcher.group(2).trim();
        String dividend = matcher.group(3) == null ? "" : matcher.group(3).trim();
        String replacement = buildLongDivisionPreviewLatex(divisor, quotient, dividend);
        return latex.substring(0, matcher.start()) + replacement + latex.substring(matcher.end());
    }

    private String expandLongDivisionPreview(String latex) {
        Matcher matcher = LONG_DIVISION_COMMAND_PATTERN.matcher(latex);
        if (!matcher.matches()) {
            return null;
        }
        String quotient = matcher.group(1) == null ? "" : matcher.group(1).trim();
        String divisor = matcher.group(2) == null ? "" : matcher.group(2).trim();
        String dividend = matcher.group(3) == null ? "" : matcher.group(3).trim();
        if (!divisor.matches("\\d+") || !dividend.matches("\\d+")) {
            return null;
        }
        return buildLongDivisionPreviewLatex(divisor, quotient, dividend);
    }

    private String buildLongDivisionPreviewLatex(String divisor, String quotient, String dividend) {
        // 预览图只保留头部，不再根据 bare longdiv 自动推导步骤区。
        String header = divisor;
        if (!quotient.isBlank()) {
            header += "\\overset{" + quotient + "}{\\overline{\\left)" + dividend + "\\right.}}";
        } else {
            header += "\\overline{\\left)" + dividend + "\\right.}";
        }
        return header;
    }

    private String buildLongDivisionUnderlineLine(int endColumn, String digits) {
        String aligned = buildLongDivisionAlignedText(endColumn, digits);
        int leadingSpaces = countLeadingSpaces(aligned);
        String visibleDigits = aligned.substring(leadingSpaces);
        StringBuilder builder = new StringBuilder();
        if (leadingSpaces > 0) {
            // 预览图不复用 MathType 的真实空格宽度，因此这里转成显式 hspace，避免 TeX 折叠连续空格。
            builder.append(buildHorizontalSpaceCommand(leadingSpaces));
        }
        builder.append("\\underline{").append(visibleDigits).append("}");
        return builder.toString();
    }

    private String buildLongDivisionTextLine(int endColumn, String digits) {
        String aligned = buildLongDivisionAlignedText(endColumn, digits);
        int leadingSpaces = countLeadingSpaces(aligned);
        String visibleDigits = aligned.substring(leadingSpaces);
        StringBuilder builder = new StringBuilder();
        if (leadingSpaces > 0) {
            builder.append(buildHorizontalSpaceCommand(leadingSpaces));
        }
        builder.append(visibleDigits);
        return builder.toString();
    }

    private String buildLongDivisionAlignedText(int endColumn, String digits) {
        if (digits == null || digits.isBlank()) {
            return "";
        }
        int leadingColumns = Math.max(endColumn - digits.length() + 1, 0);
        // 这里复用公式层已经确定下来的空格公式：n 位数 = 3*c + n - 2。
        int spaces = Math.max(leadingColumns * 3 + digits.length() - 2, 0);
        return " ".repeat(spaces) + digits;
    }

    private int countLeadingSpaces(String text) {
        int index = 0;
        while (index < text.length() && text.charAt(index) == ' ') {
            index++;
        }
        return index;
    }

    private String buildHorizontalSpaceCommand(int spaceCount) {
        double em = spaceCount * 0.33d;
        return String.format(Locale.ROOT, "\\hspace*{%.2fem}", em);
    }

    /**
     * 使用 JLaTeXMath 在内存中渲染公式，作为原生 TeX 失败时的最终兜底。
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return PNG 字节数组
     */
    private byte[] renderByJLatexMath(String latex, float size) {
        try {
            TeXFormula formula = new TeXFormula(latex);
            TeXIcon icon = formula.createTeXIcon(TeXConstants.STYLE_TEXT, size * RENDER_SCALE);
            icon.setInsets(new Insets(1, 1, 1, 1));
            icon.setForeground(Color.BLACK);

            int width = icon.getIconWidth();
            int height = icon.getIconHeight();
            if (width <= 0 || height <= 0) {
                log.warn("Formula produced empty image: {}", latex);
                return createPlaceholderImage(latex);
            }

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

    /**
     * 通过系统安装的 latex + dvisvgm 工具链渲染 PNG。
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return PNG 字节数组；失败返回 null
     */
    private byte[] renderViaDvisvgm(String latex, float size) {
        byte[] svg = renderSvgViaDvisvgm(latex, size);
        if (svg == null || svg.length == 0) {
            return null;
        }
        try {
            SvgDimensions dimensions = extractSvgDisplayDimensions(svg);
            int widthPx = Math.max((int) Math.ceil(dimensions.widthPt() * PX_PER_PT), 10);
            int heightPx = Math.max((int) Math.ceil(dimensions.heightPt() * PX_PER_PT), 10);
            return svgToPng(
                svg,
                Math.max((int) Math.ceil(widthPx * PNG_OUTPUT_SCALE), widthPx),
                Math.max((int) Math.ceil(heightPx * PNG_OUTPUT_SCALE), heightPx)
            );
        } catch (Exception e) {
            log.debug("SVG -> PNG transcode failed, fallback to JLaTeXMath: {}", latex, e);
            return null;
        }
    }

    /**
     * 通过原生 TeX 工具链生成 SVG。
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return SVG 字节数组；失败返回 null
     */
    private byte[] renderSvgViaDvisvgm(String latex, float size) {
        return renderSvgViaDvisvgm(latex, size, null);
    }

    /**
     * 通过原生 TeX 工具链生成 SVG，并允许外部传入临时目录。
     *
     * @param latex           LaTeX 公式源码
     * @param size            字号（磅）
     * @param providedTempDir 外部提供的临时目录
     * @return SVG 字节数组；失败返回 null
     */
    private byte[] renderSvgViaDvisvgm(String latex, float size, Path providedTempDir) {
        if (externalToolUnavailable) {
            return null;
        }

        Path tempDir = providedTempDir;
        boolean ownsTempDir = false;
        try {
            if (tempDir == null) {
                tempDir = Files.createTempDirectory("paperword-latex-");
                ownsTempDir = true;
            }
            Path texFile = tempDir.resolve("eq.tex");
            Path dviFile = tempDir.resolve("eq.dvi");
            Path svgFile = tempDir.resolve("eq.svg");

            Files.writeString(texFile, buildLatexDocument(latex, size), StandardCharsets.UTF_8);

            String latexCmd = System.getProperty(LATEX_CMD_PROP, "latex");
            String dvisvgmCmd = System.getProperty(DVISVGM_CMD_PROP, "dvisvgm");
            int timeoutSeconds = Integer.getInteger(RENDER_TIMEOUT_PROP, DEFAULT_TIMEOUT_SECONDS);

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
                log.debug("latex render failed (code={}): {}", latexResult.exitCode, latexResult.outputText());
                return null;
            }

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
                log.debug("dvisvgm render failed (code={}): {}", svgResult.exitCode, svgResult.outputText());
                return null;
            }
            return Files.readAllBytes(svgFile);
        } catch (IOException e) {
            if (isCommandUnavailable(e)) {
                externalToolUnavailable = true;
            }
            log.warn("External LaTeX/SVG tools unavailable, fallback to JLaTeXMath: {}", e.getMessage());
            return null;
        } catch (Exception e) {
            log.debug("External LaTeX->SVG render failed, fallback to JLaTeXMath: {}", latex, e);
            return null;
        } finally {
            if (ownsTempDir && tempDir != null) {
                deleteQuietly(tempDir);
            }
        }
    }

    /**
     * 使用 Apache Batik 将 SVG 转为 PNG。
     *
     * @param svgBytes SVG 字节数组
     * @return PNG 字节数组
     * @throws Exception 转码失败时抛出
     */
    private byte[] svgToPng(byte[] svgBytes, int widthPx, int heightPx) throws Exception {
        PNGTranscoder transcoder = new PNGTranscoder();
        transcoder.addTranscodingHint(PNGTranscoder.KEY_BACKGROUND_COLOR, Color.WHITE);
        transcoder.addTranscodingHint(PNGTranscoder.KEY_WIDTH, (float) Math.max(widthPx, 10));
        transcoder.addTranscodingHint(PNGTranscoder.KEY_HEIGHT, (float) Math.max(heightPx, 10));

        TranscoderInput input = new TranscoderInput(new StringReader(new String(svgBytes, StandardCharsets.UTF_8)));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        TranscoderOutput output = new TranscoderOutput(baos);
        transcoder.transcode(input, output);
        return baos.toByteArray();
    }

    /**
     * 执行外部命令并等待完成。
     *
     * @param command        命令及参数
     * @param workDir        工作目录
     * @param timeoutSeconds 超时秒数
     * @return 执行结果
     * @throws IOException          启动失败时抛出
     * @throws InterruptedException 等待中断时抛出
     */
    private CommandResult runCommand(List<String> command, Path workDir, int timeoutSeconds)
        throws IOException, InterruptedException {
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.directory(workDir.toFile());
        pb.redirectErrorStream(true);

        Process process = pb.start();
        byte[] output = process.getInputStream().readAllBytes();
        boolean finished = process.waitFor(timeoutSeconds, TimeUnit.SECONDS);
        if (!finished) {
            process.destroyForcibly();
            return new CommandResult(-1, output);
        }
        return new CommandResult(process.exitValue(), output);
    }

    /**
     * 构建原生 TeX 渲染使用的完整文档。
     *
     * <p>这里统一切到行内数学模式，缩小与 Word 行内公式在高度和占位上的差异；
     * 同时补充 array/cancel 宏包，覆盖项目里常见的竖式和删除线语法。</p>
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return 完整 TeX 文档
     */
    private String buildLatexDocument(String latex, float size) {
        String sizePt = String.format(Locale.ROOT, "%.1f", Math.max(size, 8f));
        String baselineSkipPt = String.format(Locale.ROOT, "%.1f", Math.max(size + 2f, 10f));
        return """
            \\documentclass[12pt]{article}
            \\usepackage{amsmath,amssymb}
            \\usepackage{array}
            \\usepackage{cancel}
            \\pagestyle{empty}
            \\setlength{\\parindent}{0pt}
            \\begin{document}
            \\fontsize{%s}{%s}\\selectfont
            $%s$
            \\end{document}
            """.formatted(sizePt, baselineSkipPt, latex);
    }

    /**
     * 静默删除目录及其全部内容。
     *
     * @param dir 临时目录
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
     * 外部命令执行结果。
     *
     * @param exitCode 退出码
     * @param output   输出内容
     */
    private record CommandResult(int exitCode, byte[] output) {
        /** 将输出字节转为 UTF-8 字符串，便于记录日志。 */
        private String outputText() {
            return new String(output, StandardCharsets.UTF_8);
        }
    }

    /**
     * 从 SVG 中提取显示尺寸。
     *
     * @param svgBytes SVG 字节数组
     * @return SVG 的磅值宽高
     */
    static SvgDimensions extractSvgDisplayDimensions(byte[] svgBytes) {
        String svg = new String(svgBytes, StandardCharsets.UTF_8);
        Matcher matcher = Pattern.compile(
            "<svg[^>]*\\bwidth=['\"]([^'\"]+)['\"][^>]*\\bheight=['\"]([^'\"]+)['\"]",
            Pattern.CASE_INSENSITIVE | Pattern.DOTALL
        ).matcher(svg);
        if (!matcher.find()) {
            return new SvgDimensions(12.0f, 12.0f);
        }
        return new SvgDimensions(parseSvgLengthToPt(matcher.group(1)), parseSvgLengthToPt(matcher.group(2)));
    }

    /**
     * 解析 SVG 长度到磅值。
     *
     * @param value SVG 长度字符串
     * @return 磅值
     */
    private static float parseSvgLengthToPt(String value) {
        Matcher matcher = Pattern.compile("([0-9.]+)\\s*([a-zA-Z]*)")
            .matcher(value == null ? "" : value.trim());
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

    /**
     * 判断外部命令是否不可用。
     *
     * @param exception 启动异常
     * @return 是否为命令缺失类错误
     */
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

    /** SVG 尺寸对象，单位为 pt。 */
    record SvgDimensions(float widthPt, float heightPt) {
    }

    /**
     * 将 LaTeX 公式渲染为 BufferedImage。
     *
     * <p>这个公共辅助方法保留原有行为，方便调试或独立图片用途，不参与主预览链路。</p>
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return 渲染后的位图
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
     * @param latex 原始 LaTeX 字符串
     * @return 占位 PNG 字节数组
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
}
