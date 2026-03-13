package com.lz.paperword.core.render;

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
import java.awt.image.BufferedImage;
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
import java.util.Comparator;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * LaTeX 公式图片渲染器 — 将 LaTeX 公式渲染为 PNG 预览图。
 *
 * <h2>核心用途</h2>
 * <p>为嵌入 Word 的 MathType OLE 对象生成预览图片。在 .docx 中，每个 OLE 公式
 * 由两部分组成：OLE 二进制数据（可双击编辑）+ 预览图（静态显示）。
 * 本类负责生成后者。当用户在 Word 中打开文档后，MathType 会自动用其内部的
 * WMF 矢量预览替换我们生成的 PNG 预览，因此这里的 PNG 主要用于首次打开时的
 * 过渡显示。</p>
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
 * <p>{@link #renderForOlePreview} 专用于 OLE 形状预览，使用 STYLE_DISPLAY 模式
 * （确保分数线、根号等元素按正确比例显示），以 3 倍分辨率渲染以获得清晰效果，
 * 但报告 1 倍尺寸用于 VML shape 的宽高定义。</p>
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

    /** 默认公式字体大小（磅），对应 Word 中正文公式的标准尺寸 */
    private static final float DEFAULT_SIZE = 13f;

    /** 通用渲染缩放因子，2 倍渲染使 PNG 在 Word 缩放时保持清晰 */
    private static final float RENDER_SCALE = 2.0f;

    // === 系统属性键（通过 -D 或 System.setProperty 配置） ===
    private static final String LATEX_CMD_PROP = "paperword.latex.command";
    private static final String DVISVGM_CMD_PROP = "paperword.dvisvgm.command";
    private static final String RENDER_TIMEOUT_PROP = "paperword.latex.timeout.seconds";
    private static final String MATHJAX_ENABLED_PROP = "paperword.mathjax.enabled";
    private static final String MATHJAX_ONLY_PROP = "paperword.mathjax.only";
    private static final String MATHJAX_ENDPOINT_PROP = "paperword.mathjax.endpoint";
    private static final String MATHJAX_TIMEOUT_PROP = "paperword.mathjax.timeout.seconds";
    private static final int DEFAULT_TIMEOUT_SECONDS = 20;
    private static final int DEFAULT_MATHJAX_TIMEOUT_SECONDS = 12;

    /** 标记外部工具（latex/dvisvgm）是否不可用，避免反复尝试失败 */
    private volatile boolean externalToolUnavailable = false;

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
     * 为 OLE 形状生成高分辨率预览图。
     *
     * <p><b>渲染策略：</b>使用 JLaTeXMath 的 STYLE_DISPLAY 模式渲染（与行内 STYLE_TEXT 不同，
     * STYLE_DISPLAY 确保分数线、根号上横线、求和号等元素按正常大小显示，
     * 与 MathType 的默认显示尺寸一致）。</p>
     *
     * <p><b>3 倍渲染 / 1 倍报告：</b>以 3 倍分辨率（renderScale = 3.0）渲染图片，
     * 保证 PNG 在 Word 中放大时仍然清晰。但返回的 widthPx/heightPx 为 1 倍尺寸，
     * 供 VML shape 的 style:width/style:height 使用——Word 会将高分辨率图片
     * 缩放到该 1 倍尺寸显示，从而实现"视网膜"效果。</p>
     *
     * <p><b>临时性：</b>此 PNG 预览仅在首次打开时使用。用户在 Word 中打开文档后，
     * MathType 会自动用自身生成的 WMF 矢量预览替换此 PNG，获得完美的矢量显示效果。</p>
     *
     * @param latex LaTeX 公式字符串
     * @return 预览图数据（含 PNG 字节和 1 倍显示尺寸）；渲染失败返回 null
     */
    public PreviewImage renderForOlePreview(String latex) {
        try {
            TeXFormula formula = new TeXFormula(latex);
            // 使用 STYLE_DISPLAY 模式：分数线、根号等元素按正确比例渲染
            // 3 倍缩放确保 PNG 在 Word 中的清晰度
            float renderScale = 3.0f;
            TeXIcon icon = formula.createTeXIcon(TeXConstants.STYLE_DISPLAY, DEFAULT_SIZE * renderScale);
            // 设置内边距（按 renderScale 缩放），为公式周围留出呼吸空间
            icon.setInsets(new Insets(
                (int)(5 * renderScale), (int)(2 * renderScale),
                (int)(5 * renderScale), (int)(2 * renderScale)));
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
            // 报告 1 倍显示尺寸：VML shape 使用此尺寸定义公式在文档中的显示大小
            // Word 会将 3 倍分辨率的 PNG 缩放到此 1 倍尺寸，实现高清显示
            int displayWidth = Math.max((int)(width / renderScale), 10);
            int displayHeight = Math.max((int)(height / renderScale), 10);
            return new PreviewImage(pngData, displayWidth, displayHeight, "png", "image/png");
        } catch (Exception e) {
            log.debug("Preview render failed: {}", latex, e);
            return null;
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
        byte[] external = renderViaDvisvgm(latex);
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

    /** EMF 矢量渲染的缩放因子 */
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
            TeXFormula formula = new TeXFormula(latex);
            // 以 2 倍 scale 渲染：确保细线元素在 EMF 坐标中有足够精度
            TeXIcon icon = formula.createTeXIcon(TeXConstants.STYLE_DISPLAY, size * EMF_RENDER_SCALE);
            icon.setInsets(new Insets(
                (int)(5 * EMF_RENDER_SCALE), (int)(2 * EMF_RENDER_SCALE),
                (int)(5 * EMF_RENDER_SCALE), (int)(2 * EMF_RENDER_SCALE)));
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
    private byte[] renderViaDvisvgm(String latex) {
        // 外部工具已标记为不可用，直接跳过
        if (externalToolUnavailable) {
            return null;
        }

        Path tempDir = null;
        try {
            // 创建临时目录存放编译中间文件
            tempDir = Files.createTempDirectory("paperword-latex-");
            Path texFile = tempDir.resolve("eq.tex");
            Path dviFile = tempDir.resolve("eq.dvi");

            // 写入完整的 LaTeX 文档（包含 amsmath/amssymb 宏包）
            Files.writeString(texFile, buildLatexDocument(latex), StandardCharsets.UTF_8);

            String latexCmd = System.getProperty(LATEX_CMD_PROP, "latex");
            String dvisvgmCmd = System.getProperty(DVISVGM_CMD_PROP, "dvisvgm");
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
                log.debug("latex render failed (code={}): {}", latexResult.exitCode, latexResult.outputText());
                return null;
            }

            // 步骤 2：dvisvgm 将 .dvi 转为 SVG（输出到 stdout）
            // --no-fonts：将文本转为路径（避免字体依赖问题）
            // --exact-bbox：精确计算边界框
            // --precision=5：坐标精度 5 位小数
            CommandResult svgResult = runCommand(
                List.of(
                    dvisvgmCmd,
                    "--verbosity=0",
                    "--exact-bbox",
                    "--no-fonts",
                    "--precision=5",
                    "--stdout",
                    dviFile.toAbsolutePath().toString()
                ),
                tempDir,
                timeoutSeconds
            );
            if (svgResult.exitCode != 0 || svgResult.output.length == 0) {
                log.debug("dvisvgm render failed (code={}): {}", svgResult.exitCode, svgResult.outputText());
                return null;
            }

            // 步骤 3：使用 Apache Batik 将 SVG 转为 PNG
            return svgToPng(svgResult.output);
        } catch (IOException e) {
            // 命令不存在或执行环境异常：标记外部工具不可用，后续调用直接跳过
            externalToolUnavailable = true;
            log.warn("External LaTeX/SVG tools unavailable, fallback to JLaTeXMath: {}", e.getMessage());
            return null;
        } catch (Exception e) {
            log.debug("External LaTeX->SVG render failed, fallback to JLaTeXMath: {}", latex, e);
            return null;
        } finally {
            // 清理临时目录及其中所有文件
            if (tempDir != null) {
                deleteQuietly(tempDir);
            }
        }
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
    private byte[] renderViaMathJaxHttp(String latex) {
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
            // 将服务返回的 SVG 转为 PNG
            return svgToPng(resp.body());
        } catch (Exception e) {
            log.debug("MathJax HTTP render failed, fallback to next renderer: {}", latex, e);
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
    private String buildLatexDocument(String latex) {
        return """
            \\documentclass[12pt]{article}
            \\usepackage{amsmath,amssymb}
            \\pagestyle{empty}
            \\begin{document}
            \\[
            %s
            \\]
            \\end{document}
            """.formatted(latex);
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
}
