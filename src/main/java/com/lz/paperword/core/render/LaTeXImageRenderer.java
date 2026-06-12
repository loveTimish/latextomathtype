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
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * LaTeX 公式图片渲染器。
 *
 * <p>当前版本统一以原生 TeX 工具链作为预览图首选来源：
 * latex → dvi → dvisvgm → SVG。OLE 预览图写成 Word 兼容的 WMF 媒体，
 * 普通图片模式和直接导出的 PNG 仍按需从同一份 SVG 派生。</p>
 *
 * <p>OLE 预览图使用严格 TeX/WMF 链路，失败即失败，不再静默回退为 JLaTeXMath 或 PNG。</p>
 */
public class LaTeXImageRenderer {

    private static final Logger log = LoggerFactory.getLogger(LaTeXImageRenderer.class);

    /** 默认公式字体大小（磅），对应 Word 中正文公式的标准尺寸。 */
    private static final float DEFAULT_SIZE = 13f;

    /** JLaTeXMath 兜底通道的渲染缩放因子，用于提高位图清晰度。 */
    private static final float RENDER_SCALE = 4.0f;

    /** 原生 TeX 预览图转 PNG 时的放大倍率，只提高底图分辨率，不改变文档显示尺寸。 */
    private static final float PNG_OUTPUT_SCALE = 4.0f;

    /** 原生 TeX 生成 OLE 预览图时的字号。MathType 默认 full size 为 12pt，跟测试集对齐。 */
    private static final float OLE_PREVIEW_SIZE = 12f;

    /** 系统属性：latex 命令路径。 */
    private static final String LATEX_CMD_PROP = "paperword.latex.command";
    /** 系统属性：xelatex 命令路径（中文公式渲染）。 */
    private static final String XELATEX_CMD_PROP = "paperword.xelatex.command";
    /** 系统属性：dvisvgm 命令路径。 */
    private static final String DVISVGM_CMD_PROP = "paperword.dvisvgm.command";
    /** 系统属性：外部命令超时秒数。 */
    private static final String RENDER_TIMEOUT_PROP = "paperword.latex.timeout.seconds";
    /** 系统属性：是否启用跨进程磁盘缓存。 */
    private static final String CACHE_ENABLED_PROP = "paperword.render.cache.enabled";
    /** 系统属性：渲染磁盘缓存目录。 */
    private static final String CACHE_DIR_PROP = "paperword.render.cache.dir";
    /** 缓存版本，公式渲染度量或图片生成逻辑变化时递增。 */
    private static final String CACHE_VERSION = "v38-xsc-vector-wmf-script-fraction";
    /** 外部命令默认超时秒数。 */
    private static final int DEFAULT_TIMEOUT_SECONDS = 20;
    /** 像素到磅的换算比例。 */
    private static final float PX_PER_PT = 1.0f / 0.75f;
    /** Keep DIB-backed WMF previews reasonably sized for extremely wide vertical-layout formulas. */
    private static final int MAX_WMF_DIB_SIDE = 4096;
    /** 显式长除法命令提取模式。 */
    private static final Pattern LONG_DIVISION_COMMAND_PATTERN =
        Pattern.compile("\\\\longdiv(?:\\[([^\\]]*)])?\\{([^{}]+)}\\{([^{}]+)}");
    /** CJK 字符检测：汉字、CJK 标点、全角形式、带圈数字。命中时走 XeLaTeX。 */
    private static final Pattern CJK_PATTERN =
        Pattern.compile("[\\u2460-\\u24FF\\u3000-\\u303F\\u3400-\\u4DBF\\u4E00-\\u9FFF\\uF900-\\uFAFF\\uFF00-\\uFFEF]");

    /** 标记外部工具是否不可用，避免每个公式都重复探测失败。 */
    private volatile boolean externalToolUnavailable = false;

    /** 公式预览图缓存。高清 TeX 渲染成本高，同一批试卷内重复公式很多，缓存能明显稳住速度。 */
    private static final Map<String, PreviewImage> PREVIEW_CACHE = new ConcurrentHashMap<>();

    /** PNG 字节缓存，供直接图片导出入口复用。 */
    private static final Map<String, byte[]> PNG_CACHE = new ConcurrentHashMap<>();

    /**
     * 预览图数据记录。
     *
     * @param data        图片字节
     * @param widthPx     显示宽度（像素）
     * @param heightPx    显示高度（像素）
     * @param extension   扩展名
     * @param contentType MIME 类型
     * @param placeholder 是否为占位图
     * @param depthPt     基线以下深度（磅）；&lt;0 表示未知
     * @param widthPt     物理宽度（磅）；&lt;0 表示按像素换算
     * @param heightPt    物理高度（磅）；&lt;0 表示按像素换算
     */
    public record PreviewImage(
        byte[] data,
        int widthPx,
        int heightPx,
        String extension,
        String contentType,
        boolean placeholder,
        double depthPt,
        double widthPt,
        double heightPt
    ) {
        public PreviewImage(byte[] data, int widthPx, int heightPx, String extension, String contentType,
                            boolean placeholder) {
            this(data, widthPx, heightPx, extension, contentType, placeholder, -1d, -1d, -1d);
        }

        public PreviewImage(byte[] data, int widthPx, int heightPx, String extension, String contentType,
                            boolean placeholder, double depthPt) {
            this(data, widthPx, heightPx, extension, contentType, placeholder, depthPt,
                widthPx * 0.75d, heightPx * 0.75d);
        }
    }

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
     * <p>严格走原生 TeX/WMF 渲染，失败后直接中止嵌入，避免悄悄退回 PNG 或占位图。</p>
     *
     * @param latex LaTeX 公式源码
     * @return 预览图数据；失败返回 null
     */
    public PreviewImage renderForOlePreview(String latex) {
        String cacheKey = cacheKey("ole", latex, OLE_PREVIEW_SIZE);
        PreviewImage cached = PREVIEW_CACHE.get(cacheKey);
        if (cached != null) {
            return cached;
        }
        cached = readPreviewFromDisk(cacheKey);
        if (cached != null) {
            PREVIEW_CACHE.put(cacheKey, cached);
            return cached;
        }
        PreviewImage preview = renderWmfPreviewViaTeX(latex, OLE_PREVIEW_SIZE);
        if (preview != null) {
            PREVIEW_CACHE.put(cacheKey, preview);
            writePreviewToDisk(cacheKey, preview);
            return preview;
        }
        throw new IllegalStateException("Native TeX/WMF OLE preview rendering failed: " + latex);
    }

    public PreviewImage renderForOlePreview(String latex, Double targetWidthPt, Double targetHeightPt) {
        if (targetWidthPt == null || targetHeightPt == null || targetWidthPt <= 0d || targetHeightPt <= 0d) {
            return renderForOlePreview(latex);
        }
        String cacheKey = cacheKey("ole-target-" + String.format(Locale.ROOT, "%.2fx%.2f", targetWidthPt, targetHeightPt),
            latex, OLE_PREVIEW_SIZE);
        PreviewImage cached = PREVIEW_CACHE.get(cacheKey);
        if (cached != null) {
            return cached;
        }
        cached = readPreviewFromDisk(cacheKey);
        if (cached != null) {
            PREVIEW_CACHE.put(cacheKey, cached);
            return cached;
        }
        PreviewImage preview = renderWmfPreviewViaTeX(latex, OLE_PREVIEW_SIZE, targetWidthPt, targetHeightPt);
        if (preview != null) {
            PREVIEW_CACHE.put(cacheKey, preview);
            writePreviewToDisk(cacheKey, preview);
            return preview;
        }
        throw new IllegalStateException("Native TeX/WMF target preview rendering failed: " + latex);
    }

    /**
     * 为 Word 普通图片模式生成预览图。
     *
     * @param latex LaTeX 公式源码
     * @return 预览图数据；失败返回 null
     */
    public PreviewImage renderForWordImage(String latex) {
        String cacheKey = cacheKey("word", latex, DEFAULT_SIZE);
        PreviewImage cached = PREVIEW_CACHE.get(cacheKey);
        if (cached != null) {
            return cached;
        }
        cached = readPreviewFromDisk(cacheKey);
        if (cached != null) {
            PREVIEW_CACHE.put(cacheKey, cached);
            return cached;
        }
        PreviewImage preview = renderPreviewViaTeX(latex, DEFAULT_SIZE);
        if (preview != null) {
            PREVIEW_CACHE.put(cacheKey, preview);
            writePreviewToDisk(cacheKey, preview);
            return preview;
        }
        preview = renderPreviewViaJLatexMath(latex, DEFAULT_SIZE);
        if (preview != null) {
            PREVIEW_CACHE.put(cacheKey, preview);
            writePreviewToDisk(cacheKey, preview);
        }
        return preview;
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
        String cacheKey = cacheKey("png", latex, size);
        byte[] cached = PNG_CACHE.get(cacheKey);
        if (cached != null) {
            return cached;
        }
        cached = readPngFromDisk(cacheKey);
        if (cached != null && cached.length > 0) {
            PNG_CACHE.put(cacheKey, cached);
            return cached;
        }
        String localRenderLatex = normalizeLatexForLocalRender(latex);
        byte[] external = renderViaDvisvgm(localRenderLatex, size);
        if (external != null && external.length > 0) {
            PNG_CACHE.put(cacheKey, external);
            writePngToDisk(cacheKey, external);
            return external;
        }
        byte[] fallback = renderByJLatexMath(localRenderLatex, size);
        if (fallback != null && fallback.length > 0) {
            PNG_CACHE.put(cacheKey, fallback);
            writePngToDisk(cacheKey, fallback);
        }
        return fallback;
    }

    private String cacheKey(String mode, String latex, float size) {
        return CACHE_VERSION + "|" + mode + "|" + size + "|" + normalizeLatexForLocalRender(latex == null ? "" : latex);
    }

    private PreviewImage readPreviewFromDisk(String cacheKey) {
        if (!diskCacheEnabled()) {
            return null;
        }
        Path base = cacheBasePath(cacheKey);
        Path metaPath = base.resolveSibling(base.getFileName() + ".properties");
        if (!Files.isRegularFile(metaPath)) {
            return null;
        }
        try {
            Properties props = new Properties();
            try (var in = Files.newInputStream(metaPath)) {
                props.load(in);
            }
            String extension = props.getProperty("extension", "png");
            Path imagePath = base.resolveSibling(base.getFileName() + "." + extension);
            if (!Files.isRegularFile(imagePath)) {
                return null;
            }
            byte[] data = Files.readAllBytes(imagePath);
            if (data.length == 0) {
                return null;
            }
            return new PreviewImage(
                data,
                Integer.parseInt(props.getProperty("widthPx", "10")),
                Integer.parseInt(props.getProperty("heightPx", "10")),
                props.getProperty("extension", "png"),
                props.getProperty("contentType", "image/png"),
                Boolean.parseBoolean(props.getProperty("placeholder", "false")),
                Double.parseDouble(props.getProperty("depthPt", "-1")),
                Double.parseDouble(props.getProperty("widthPt", "-1")),
                Double.parseDouble(props.getProperty("heightPt", "-1"))
            );
        } catch (Exception e) {
            log.debug("Formula preview disk cache read failed: {}", cacheKey, e);
            return null;
        }
    }

    private void writePreviewToDisk(String cacheKey, PreviewImage preview) {
        if (!diskCacheEnabled() || preview == null || preview.data() == null || preview.data().length == 0) {
            return;
        }
        Path base = cacheBasePath(cacheKey);
        Path imagePath = base.resolveSibling(base.getFileName() + "." + preview.extension());
        Path metaPath = base.resolveSibling(base.getFileName() + ".properties");
        try {
            Files.createDirectories(base.getParent());
            Files.write(imagePath, preview.data());
            Properties props = new Properties();
            props.setProperty("widthPx", Integer.toString(preview.widthPx()));
            props.setProperty("heightPx", Integer.toString(preview.heightPx()));
            props.setProperty("extension", preview.extension());
            props.setProperty("contentType", preview.contentType());
            props.setProperty("placeholder", Boolean.toString(preview.placeholder()));
            props.setProperty("depthPt", Double.toString(preview.depthPt()));
            props.setProperty("widthPt", Double.toString(preview.widthPt()));
            props.setProperty("heightPt", Double.toString(preview.heightPt()));
            try (var out = Files.newOutputStream(metaPath)) {
                props.store(out, "paperword formula preview cache");
            }
        } catch (Exception e) {
            log.debug("Formula preview disk cache write failed: {}", cacheKey, e);
        }
    }

    private byte[] readPngFromDisk(String cacheKey) {
        if (!diskCacheEnabled()) {
            return null;
        }
        Path pngPath = cacheBasePath(cacheKey).resolveSibling(cacheBasePath(cacheKey).getFileName() + ".png");
        try {
            return Files.isRegularFile(pngPath) ? Files.readAllBytes(pngPath) : null;
        } catch (IOException e) {
            log.debug("Formula PNG disk cache read failed: {}", cacheKey, e);
            return null;
        }
    }

    private void writePngToDisk(String cacheKey, byte[] png) {
        if (!diskCacheEnabled() || png == null || png.length == 0) {
            return;
        }
        Path pngPath = cacheBasePath(cacheKey).resolveSibling(cacheBasePath(cacheKey).getFileName() + ".png");
        try {
            Files.createDirectories(pngPath.getParent());
            Files.write(pngPath, png);
        } catch (IOException e) {
            log.debug("Formula PNG disk cache write failed: {}", cacheKey, e);
        }
    }

    private boolean diskCacheEnabled() {
        return Boolean.parseBoolean(System.getProperty(CACHE_ENABLED_PROP, "true"));
    }

    private Path cacheBasePath(String cacheKey) {
        String configuredDir = System.getProperty(CACHE_DIR_PROP, "data/cache/formula-render");
        String digest = sha256Base64Url(cacheKey);
        return Path.of(configuredDir, digest.substring(0, 2), digest.substring(2));
    }

    private String sha256Base64Url(String value) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            return Base64.getUrlEncoder().withoutPadding()
                .encodeToString(digest.digest(value.getBytes(StandardCharsets.UTF_8)));
        } catch (NoSuchAlgorithmException e) {
            throw new IllegalStateException("SHA-256 digest unavailable", e);
        }
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
     * OLE 对象预览图的严格 TeX/WMF 入口。
     *
     * <p>当前实现先生成 Word 能识别的 placeable WMF。后续可继续把 WMF 内部从
     * DIB 预览升级为 MathType 风格矢量 record，但 DOCX 媒体路线已经不再使用 PNG。</p>
     */
    private PreviewImage renderWmfPreviewViaTeX(String latex, float size) {
        return renderWmfPreviewViaTeX(latex, size, null, null);
    }

    private PreviewImage renderWmfPreviewViaTeX(String latex, float size, Double targetWidthPt, Double targetHeightPt) {
        String localRenderLatex = normalizeLatexForLocalRender(latex);
        Path tempDir = null;
        try {
            tempDir = Files.createTempDirectory("paperword-latex-");
            byte[] svg = renderSvgViaDvisvgm(localRenderLatex, size, tempDir);
            if (svg == null || svg.length == 0) {
                return null;
            }
            TexBoxMetrics metrics = readTexBoxMetrics(tempDir.resolve("eq.size"));
            SvgDimensions dimensions = extractSvgDisplayDimensions(svg);
            double widthPt = metrics != null ? metrics.widthPt() : dimensions.widthPt();
            double heightPt = metrics != null ? metrics.heightPt() + metrics.depthPt() : dimensions.heightPt();
            double depthPt = metrics != null ? metrics.depthPt() : -1d;
            PreviewMetrics calibrated = calibratePreviewMetrics(latex, widthPt, heightPt, depthPt);
            widthPt = calibrated.widthPt();
            heightPt = calibrated.heightPt();
            depthPt = calibrated.depthPt();
            if (targetWidthPt != null && targetHeightPt != null && targetWidthPt > 0d && targetHeightPt > 0d) {
                if (depthPt >= 0d && heightPt > 0d) {
                    depthPt = depthPt * targetHeightPt / heightPt;
                }
                widthPt = targetWidthPt;
                heightPt = targetHeightPt;
            }

            int widthPx = Math.max((int) Math.round(widthPt * PX_PER_PT), 4);
            int heightPx = Math.max((int) Math.round(heightPt * PX_PER_PT), 4);
            if (VectorWmfFormulaRenderer.canRender(latex)) {
                byte[] wmfData = VectorWmfFormulaRenderer.render(latex, widthPt, heightPt);
                if (wmfData != null && wmfData.length > 0) {
                    return new PreviewImage(wmfData, widthPx, heightPx, "wmf", "image/x-wmf", false,
                        depthPt, widthPt, heightPt);
                }
            }
            int renderWidthPx = Math.max((int) Math.ceil(widthPx * PNG_OUTPUT_SCALE), widthPx);
            int renderHeightPx = Math.max((int) Math.ceil(heightPx * PNG_OUTPUT_SCALE), heightPx);
            byte[] pngData = svgToPng(svg, renderWidthPx, renderHeightPx);
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(pngData));
            if (image == null) {
                return null;
            }
            byte[] wmfData = bufferedImageToPlaceableWmf(image, widthPt, heightPt);
            return new PreviewImage(wmfData, widthPx, heightPx, "wmf", "image/x-wmf", false,
                depthPt, widthPt, heightPt);
        } catch (Exception e) {
            log.error("Native TeX/WMF preview render failed: {}", latex, e);
            return null;
        } finally {
            if (tempDir != null) {
                deleteQuietly(tempDir);
            }
        }
    }

    /** TeX 盒子度量（磅）。 */
    private record TexBoxMetrics(double widthPt, double heightPt, double depthPt) {
    }

    private record PreviewMetrics(double widthPt, double heightPt, double depthPt) {
    }

    /**
     * Empirical xsc/MathType preview calibration.
     *
     * <p>The editable OLE body is produced from the original AST. These factors only tune
     * the TeX-derived WMF preview box so Word's displayed object metrics better match the
     * legacy MathType corpus.</p>
     */
    private PreviewMetrics calibratePreviewMetrics(String latex, double widthPt, double heightPt, double depthPt) {
        double widthScale = 1.0d;
        double heightScale = 1.0d;
        boolean hasArray = latex.contains("\\begin{array}");
        int arrayCount = countOccurrences(latex, "\\begin{array}");
        int lineBreaks = countOccurrences(latex, "\\\\");
        int fractionCount = countOccurrences(latex, "\\frac");
        boolean script = hasScript(latex);
        String previewClass = classifyPreviewLatex(latex, hasArray, fractionCount, script);

        if (hasArray) {
            if (arrayCount == 1 && lineBreaks >= 1) {
                widthScale *= 0.86d;
                heightScale *= fractionCount > 0 ? 1.03d : 0.94d;
                if (lineBreaks == 1 && fractionCount >= 4) {
                    widthScale *= 1.07d;
                    heightScale *= 1.12d;
                }
                if (fractionCount == 0 && !script && !latex.contains("\\cdots")
                    && countOccurrences(latex, "\\left") >= 1 && latex.length() >= 70
                    && (latex.contains("\\div") || latex.contains("\\left ("))) {
                    widthScale *= 0.95d;
                    heightScale *= 1.12d;
                }
                if (fractionCount == 0 && latex.contains("\\begin{array}{ccc}")
                    && countOccurrences(latex, "{}") >= 6) {
                    widthScale *= 3.20d;
                }
            } else if (arrayCount == 1 && lineBreaks == 0 && fractionCount > 0) {
                widthScale *= 0.78d;
                heightScale *= 1.05d;
            } else if (arrayCount > 1) {
                widthScale *= 0.94d;
                heightScale *= 1.04d;
                if (arrayCount == 3 && lineBreaks >= 6 && latex.contains("\\right.,\\left")) {
                    widthScale *= 0.71d;
                }
                if (arrayCount >= 4 && !latex.contains("{cc}") && countOccurrences(latex, "\\left") >= 4) {
                    widthScale *= 1.25d;
                    heightScale *= 1.60d;
                } else if (lineBreaks == 0) {
                    widthScale *= 0.78d;
                    heightScale *= 0.76d;
                } else if (arrayCount == 2 && lineBreaks == 2 && latex.contains("\\right.,\\left")) {
                    widthScale *= 0.73d;
                    heightScale *= 0.88d;
                }
            }
        } else if (fractionCount == 1) {
            widthScale *= 0.96d;
            heightScale *= 0.94d;
        } else if (fractionCount > 1) {
            widthScale *= 0.90d;
            heightScale *= 0.94d;
        } else if (script) {
            widthScale *= 0.96d;
            boolean cjk = containsCjk(latex);
            if (latex.contains("\\left") || latex.contains("\\right") || cjk) {
                heightScale *= 1.32d;
                if (cjk && latex.length() > 60 && countOccurrences(latex, "\\div") > 0) {
                    widthScale *= 0.74d;
                    heightScale *= 1.33d;
                }
            } else {
                heightScale *= 1.01d;
            }
        } else if (!latex.contains("\\sqrt")) {
            widthScale *= 1.08d;
            if (latex.contains("\\left") || latex.contains("\\right")) {
                widthScale *= 0.90d;
                heightScale *= 1.30d;
            }
        }

        if (latex.contains("\\sqrt")) {
            widthScale *= 0.92d;
            heightScale *= 1.16d;
        }
        widthScale *= previewWidthClassScale(previewClass);

        double adjustedDepth = depthPt >= 0d ? depthPt * heightScale : depthPt;
        return new PreviewMetrics(
            Math.max(widthPt * widthScale, 1.0d),
            Math.max(heightPt * heightScale, 1.0d),
            adjustedDepth
        );
    }

    private String classifyPreviewLatex(String latex, boolean hasArray, int fractionCount, boolean script) {
        if (hasArray) {
            return "array";
        }
        if (fractionCount > 0) {
            return "fraction";
        }
        if (latex.contains("\\sqrt")) {
            return "sqrt";
        }
        if (script) {
            return "script";
        }
        if (latex.contains("\\overline") || latex.contains("\\underline")
            || latex.contains("\\overset") || latex.contains("\\underset")) {
            return "accent";
        }
        return "linear";
    }

    private double previewWidthClassScale(String previewClass) {
        String propertyName = "paperword.preview.width.scale." + previewClass;
        double defaultScale = switch (previewClass) {
            case "array" -> 0.91d;
            case "fraction" -> 0.88d;
            case "sqrt" -> 0.90d;
            case "script" -> 0.85d;
            case "accent" -> 0.90d;
            default -> 0.86d;
        };
        return readDoubleProperty(propertyName, defaultScale);
    }

    private double readDoubleProperty(String propertyName, double defaultValue) {
        String configured = System.getProperty(propertyName);
        if (configured == null || configured.isBlank()) {
            return defaultValue;
        }
        try {
            return Double.parseDouble(configured.trim());
        } catch (NumberFormatException e) {
            log.warn("Invalid numeric property {}={}, using {}", propertyName, configured, defaultValue);
            return defaultValue;
        }
    }

    private static int countOccurrences(String text, String needle) {
        if (text == null || text.isEmpty() || needle == null || needle.isEmpty()) {
            return 0;
        }
        int count = 0;
        int index = 0;
        while ((index = text.indexOf(needle, index)) >= 0) {
            count++;
            index += needle.length();
        }
        return count;
    }

    /**
     * 读取 TeX 文档写出的盒子度量文件，内容形如 {@code 123.4pt,10.0pt,3.0pt}。
     */
    private TexBoxMetrics readTexBoxMetrics(Path sizeFile) {
        try {
            if (!Files.isRegularFile(sizeFile)) {
                return null;
            }
            String[] parts = Files.readString(sizeFile, StandardCharsets.UTF_8).trim().split(",");
            if (parts.length != 3) {
                return null;
            }
            return new TexBoxMetrics(
                parseTexPt(parts[0]),
                parseTexPt(parts[1]),
                parseTexPt(parts[2])
            );
        } catch (Exception e) {
            log.debug("Failed to read TeX box metrics: {}", sizeFile, e);
            return null;
        }
    }

    /** TeX 的 pt 是 big point 的 72.27/72；Word/WMF 使用 PostScript point，需要换算。 */
    private static double parseTexPt(String value) {
        String numeric = value.trim().replace("pt", "");
        return Double.parseDouble(numeric) * 72.0 / 72.27;
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
                Math.max(Math.round(image.getWidth() / RENDER_SCALE), 10),
                Math.max(Math.round(image.getHeight() / RENDER_SCALE), 10),
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
        String normalized = com.lz.paperword.core.latex.LaTeXParser.preNormalizeLatex(
            latex.replaceAll("\\\\kern\\s*[-+]?\\d*\\.?\\d+[a-zA-Z]+", ""));
        normalized = simplifyFlatDelimiters(normalized);
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
            boolean useXeCJK = containsCjk(latex);
            Path texFile = tempDir.resolve("eq.tex");
            Path dviFile = tempDir.resolve(useXeCJK ? "eq.xdv" : "eq.dvi");
            Path svgFile = tempDir.resolve("eq.svg");

            Files.writeString(texFile, buildLatexDocument(latex, size, useXeCJK), StandardCharsets.UTF_8);

            String dvisvgmCmd = System.getProperty(DVISVGM_CMD_PROP, "dvisvgm");
            int timeoutSeconds = Integer.getInteger(RENDER_TIMEOUT_PROP, DEFAULT_TIMEOUT_SECONDS);

            List<String> latexCommand;
            if (useXeCJK) {
                latexCommand = List.of(
                    System.getProperty(XELATEX_CMD_PROP, "xelatex"),
                    "-no-pdf",
                    "-interaction=nonstopmode",
                    "-halt-on-error",
                    "-no-shell-escape",
                    "-output-directory=" + tempDir.toAbsolutePath(),
                    texFile.toAbsolutePath().toString()
                );
            } else {
                latexCommand = List.of(
                    System.getProperty(LATEX_CMD_PROP, "latex"),
                    "-interaction=nonstopmode",
                    "-halt-on-error",
                    "-no-shell-escape",
                    "-output-directory=" + tempDir.toAbsolutePath(),
                    texFile.toAbsolutePath().toString()
                );
            }
            CommandResult latexResult = runCommand(latexCommand, tempDir, timeoutSeconds);
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
        transcoder.addTranscodingHint(PNGTranscoder.KEY_WIDTH, (float) Math.max(widthPx, 4));
        transcoder.addTranscodingHint(PNGTranscoder.KEY_HEIGHT, (float) Math.max(heightPx, 4));

        TranscoderInput input = new TranscoderInput(new StringReader(new String(svgBytes, StandardCharsets.UTF_8)));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        TranscoderOutput output = new TranscoderOutput(baos);
        transcoder.transcode(input, output);
        return baos.toByteArray();
    }

    /**
     * 将 TeX 渲染得到的预览位图封装进 placeable WMF，供 Word 的 VML OLE preview 使用。
     *
     * <p>这一步只负责把 DOCX 媒体路线从 PNG 切到 WMF；后续可以继续把 DIB 记录替换为
     * MathType 风格矢量绘制记录。</p>
     */
    private byte[] bufferedImageToPlaceableWmf(BufferedImage image, double logicalWidthPt, double logicalHeightPt)
        throws IOException {
        int srcWidth = Math.max(image.getWidth(), 1);
        int srcHeight = Math.max(image.getHeight(), 1);
        int destWidth = Math.max((int) Math.round(logicalWidthPt * PX_PER_PT), 1);
        int destHeight = Math.max((int) Math.round(logicalHeightPt * PX_PER_PT), 1);
        double dibScale = Math.min(1.0d, (double) MAX_WMF_DIB_SIDE / Math.max(destWidth, destHeight));
        int dibWidth = destWidth;
        int dibHeight = destHeight;
        if (dibScale < 1.0d) {
            dibWidth = Math.max((int) Math.round(destWidth * dibScale), 1);
            dibHeight = Math.max((int) Math.round(destHeight * dibScale), 1);
        }
        if (image.getWidth() != dibWidth || image.getHeight() != dibHeight) {
            destWidth = dibWidth;
            destHeight = dibHeight;
            image = scaleImage(image, destWidth, destHeight);
            srcWidth = Math.max(image.getWidth(), 1);
            srcHeight = Math.max(image.getHeight(), 1);
        }

        byte[] dib = createBottomUp24BitDib(image);
        final int wmfSrcWidth = srcWidth;
        final int wmfSrcHeight = srcHeight;
        final int wmfDestWidth = destWidth;
        final int wmfDestHeight = destHeight;

        ByteArrayOutputStream records = new ByteArrayOutputStream();
        int maxRecordWords = 0;
        maxRecordWords = Math.max(maxRecordWords, writeWmfRecord(records, 0x0103, wmfPayload(out -> {
            writeWord(out, 8); // MM_ANISOTROPIC
        })));
        maxRecordWords = Math.max(maxRecordWords, writeWmfRecord(records, 0x020B, wmfPayload(out -> {
            writeShort(out, 0);
            writeShort(out, 0);
        })));
        maxRecordWords = Math.max(maxRecordWords, writeWmfRecord(records, 0x020C, wmfPayload(out -> {
            writeShort(out, wmfDestHeight);
            writeShort(out, wmfDestWidth);
        })));
        maxRecordWords = Math.max(maxRecordWords, writeWmfRecord(records, 0x0F43, wmfPayload(out -> {
            writeDWord(out, 0x00CC0020L); // SRCCOPY
            writeWord(out, 0); // DIB_RGB_COLORS
            writeShort(out, wmfSrcHeight);
            writeShort(out, wmfSrcWidth);
            writeShort(out, 0);
            writeShort(out, 0);
            writeShort(out, wmfDestHeight);
            writeShort(out, wmfDestWidth);
            writeShort(out, 0);
            writeShort(out, 0);
            out.write(dib);
        })));
        maxRecordWords = Math.max(maxRecordWords, writeWmfRecord(records, 0x0000, new byte[0]));

        byte[] recordBytes = records.toByteArray();
        int fileSizeWords = (18 + recordBytes.length) / 2;

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writePlaceableHeader(out, logicalWidthPt, logicalHeightPt);
        writeWord(out, 1); // memory metafile
        writeWord(out, 9); // header size in WORDs
        writeWord(out, 0x0300);
        writeDWord(out, fileSizeWords);
        writeWord(out, 0);
        writeDWord(out, maxRecordWords);
        writeWord(out, 0);
        out.write(recordBytes);
        return out.toByteArray();
    }

    private BufferedImage scaleImage(BufferedImage source, int width, int height) {
        BufferedImage scaled = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = scaled.createGraphics();
        try {
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, width, height);
            g.drawImage(source, 0, 0, width, height, null);
        } finally {
            g.dispose();
        }
        return scaled;
    }

    private byte[] createBottomUp24BitDib(BufferedImage image) throws IOException {
        int width = image.getWidth();
        int height = image.getHeight();
        int rowStride = ((width * 3 + 3) / 4) * 4;
        int imageSize = rowStride * height;

        ByteArrayOutputStream out = new ByteArrayOutputStream(40 + imageSize);
        writeDWord(out, 40);
        writeDWord(out, width);
        writeDWord(out, height);
        writeWord(out, 1);
        writeWord(out, 24);
        writeDWord(out, 0);
        writeDWord(out, imageSize);
        writeDWord(out, 0);
        writeDWord(out, 0);
        writeDWord(out, 0);
        writeDWord(out, 0);

        byte[] padding = new byte[rowStride - width * 3];
        for (int y = height - 1; y >= 0; y--) {
            for (int x = 0; x < width; x++) {
                int rgb = image.getRGB(x, y);
                out.write(rgb & 0xFF);
                out.write((rgb >>> 8) & 0xFF);
                out.write((rgb >>> 16) & 0xFF);
            }
            out.write(padding);
        }
        return out.toByteArray();
    }

    private void writePlaceableHeader(ByteArrayOutputStream out, double widthPt, double heightPt) throws IOException {
        int inch = placeableUnitsPerInch(widthPt, heightPt);
        int right = Math.max((int) Math.round(widthPt / 72.0d * inch), 1);
        int bottom = Math.max((int) Math.round(heightPt / 72.0d * inch), 1);

        ByteArrayOutputStream header = new ByteArrayOutputStream(22);
        writeDWord(header, 0x9AC6CDD7L);
        writeWord(header, 0);
        writeShort(header, 0);
        writeShort(header, 0);
        writeShort(header, right);
        writeShort(header, bottom);
        writeWord(header, inch);
        writeDWord(header, 0);

        byte[] prefix = header.toByteArray();
        int checksum = 0;
        for (int i = 0; i < 10; i++) {
            checksum ^= Short.toUnsignedInt(ByteBuffer.wrap(prefix, i * 2, 2)
                .order(ByteOrder.LITTLE_ENDIAN).getShort());
        }
        out.write(prefix);
        writeWord(out, checksum);
    }

    private int placeableUnitsPerInch(double widthPt, double heightPt) {
        double maxInches = Math.max(widthPt, heightPt) / 72.0d;
        if (maxInches <= 0) {
            return 1440;
        }
        int maxUnits = (int) Math.floor(32760.0d / maxInches);
        int[] candidates = {1440, 720, 360, 180, 120, 96, 72};
        for (int candidate : candidates) {
            if (candidate <= maxUnits) {
                return candidate;
            }
        }
        return Math.max(maxUnits, 1);
    }

    private int writeWmfRecord(ByteArrayOutputStream out, int function, byte[] payload) throws IOException {
        int payloadLength = payload.length + (payload.length % 2);
        int sizeWords = (6 + payloadLength) / 2;
        writeDWord(out, sizeWords);
        writeWord(out, function);
        out.write(payload);
        if ((payload.length & 1) != 0) {
            out.write(0);
        }
        return sizeWords;
    }

    private byte[] wmfPayload(WmfPayloadWriter writer) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(out);
        return out.toByteArray();
    }

    private void writeWord(ByteArrayOutputStream out, int value) throws IOException {
        out.write(value & 0xFF);
        out.write((value >>> 8) & 0xFF);
    }

    private void writeShort(ByteArrayOutputStream out, int value) throws IOException {
        writeWord(out, value);
    }

    private void writeDWord(ByteArrayOutputStream out, long value) throws IOException {
        out.write((int) (value & 0xFF));
        out.write((int) ((value >>> 8) & 0xFF));
        out.write((int) ((value >>> 16) & 0xFF));
        out.write((int) ((value >>> 24) & 0xFF));
    }

    @FunctionalInterface
    private interface WmfPayloadWriter {
        void write(ByteArrayOutputStream out) throws IOException;
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
     * <p>对齐 MathType 的对象框几何（基于 word_files 测试集 4 万个对象的标定）：</p>
     * <ul>
     *   <li>每个公式行有完整行框：12pt 时 ascent 10pt + descent 3pt = 13pt；</li>
     *   <li>分数的分子分母各占一个完整行框（12pt 分数总高约 28pt）；</li>
     *   <li>array/pile 行高约 16.5pt（12pt 时），用 arraystretch 撑开；</li>
     *   <li>使用 preview/tightpage 让输出 bbox 等于 TeX 盒子（含空白行框），
     *       并把盒子的 width/height/depth 写入 eq.size 供尺寸与基线计算。</li>
     * </ul>
     *
     * @param latex LaTeX 公式源码
     * @param size  字号（磅）
     * @return 完整 TeX 文档
     */
    private String buildLatexDocument(String latex, float size, boolean useXeCJK) {
        float effective = Math.max(size, 8f);
        String sizePt = String.format(Locale.ROOT, "%.1f", effective);
        String baselineSkipPt = String.format(Locale.ROOT, "%.1f", effective + 2f);
        // MathType 行框：12pt 时 ascent 10 / descent 3，按字号等比缩放。
        // 含上下标的公式行框更大（测试集标定：ascent 11 / descent 5）。
        float scale = effective / 12f;
        float ascentPt = (hasScript(latex) ? 11f : 10f) * scale;
        float descentPt = (hasScript(latex) ? 5f : 3f) * scale;
        String strutDpPt = String.format(Locale.ROOT, "%.2f", descentPt);
        String strutTotalPt = String.format(Locale.ROOT, "%.2f", ascentPt + descentPt);
        String cjkSetup = useXeCJK
            ? """
              \\usepackage{xeCJK}
              \\IfFontExistsTF{SimSun}{\\setCJKmainfont{SimSun}}{\\IfFontExistsTF{FandolSong}{\\setCJKmainfont{FandolSong}}{\\setCJKmainfont{Noto Serif CJK SC}}}
              """
            : "";
        // MathType 上下标字号为全尺寸的 58%/42%（TeX 默认 70%/50%），对齐以贴近测试集宽度。
        String scriptSizePt = String.format(Locale.ROOT, "%.1f", effective * 0.58f);
        String scriptScriptSizePt = String.format(Locale.ROOT, "%.1f", effective * 0.42f);
        return """
            \\documentclass[12pt]{article}
            \\usepackage{amsmath,amssymb}
            \\usepackage{array}
            \\usepackage{cancel}
            \\usepackage[active,tightpage]{preview}
            %6$s\\pagestyle{empty}
            \\setlength{\\parindent}{0pt}
            \\setlength{\\PreviewBorder}{0pt}
            \\setlength{\\arraycolsep}{2.5pt}
            \\newcommand{\\mtstrut}{\\rule[-%2$spt]{0pt}{%3$spt}}
            \\let\\paperwordorigcdots\\cdots
            \\renewcommand{\\cdots}{\\mathinner{\\cdotp\\mkern-2mu\\cdotp\\mkern-2mu\\cdotp}}
            \\let\\paperwordorigldots\\ldots
            \\renewcommand{\\ldots}{\\mathinner{.\\mkern-2mu.\\mkern-2mu.}}
            \\let\\paperwordorigfrac\\frac
            \\renewcommand{\\frac}[2]{\\paperwordorigfrac{\\mtstrut #1}{\\mtstrut #2}}
            \\newcommand{\\overarc}[1]{\\overset{\\frown}{#1}}
            \\newcommand{\\arc}[1]{\\overarc{#1}}
            \\newcommand{\\wideparen}[1]{\\overarc{#1}}
            \\newcommand{\\whitestar}{\\star}
            \\newcommand{\\blackstar}{\\star}
            \\newcommand{\\whitediamond}{\\diamond}
            \\newcommand{\\underbracechar}{\\underbrace{\\hphantom{0}}}
            \\renewcommand{\\arraystretch}{1.15}
            \\DeclareMathSizes{%4$s}{%4$s}{%7$s}{%8$s}
            \\thinmuskip=2mu
            \\medmuskip=3mu plus 1mu minus 2mu
            \\thickmuskip=3.5mu plus 2mu
            \\newwrite\\paperwordsize
            \\begin{document}
            \\fontsize{%4$s}{%5$s}\\selectfont
            \\setbox0=\\hbox{$\\displaystyle\\mtstrut %1$s$}
            \\immediate\\openout\\paperwordsize=eq.size
            \\immediate\\write\\paperwordsize{\\the\\wd0,\\the\\ht0,\\the\\dp0}
            \\immediate\\closeout\\paperwordsize
            \\begin{preview}\\box0\\end{preview}
            \\end{document}
            """.formatted(latex, strutDpPt, strutTotalPt, sizePt, baselineSkipPt, cjkSetup,
                scriptSizePt, scriptScriptSizePt);
    }

    /**
     * 没有分数、根号、多行等高结构的公式里，\left/\right 自动定界符只会让
     * 括号变宽（TeX 的 \nulldelimiterspace 和内侧间距），与 MathType 的紧凑
     * 括号差距明显。这里把平坦公式里的伸缩定界符退化为普通字符。
     */
    private static String simplifyFlatDelimiters(String latex) {
        if (latex == null
            || latex.contains("\\frac") || latex.contains("\\dfrac") || latex.contains("\\cfrac")
            || latex.contains("\\sqrt") || latex.contains("\\begin")
            || latex.contains("\\sum") || latex.contains("\\int") || latex.contains("\\prod")
            || latex.contains("\\overline") || latex.contains("\\underline")) {
            return latex;
        }
        return latex
            .replaceAll("\\\\left\\s*\\.", "")
            .replaceAll("\\\\right\\s*\\.", "")
            .replaceAll("\\\\(?:left|right)\\s*(?=[()\\[\\]|])", "")
            .replaceAll("\\\\(?:left|right)\\s*(?=\\\\[{}|])", "");
    }

    /** 是否包含会撑大 MathType 行框的上下标结构。 */
    private static boolean hasScript(String latex) {
        return latex != null && (latex.indexOf('^') >= 0 || latex.indexOf('_') >= 0);
    }

    /** 是否包含需要 XeLaTeX/xeCJK 渲染的中文或全角字符。 */
    private static boolean containsCjk(String latex) {
        return latex != null && CJK_PATTERN.matcher(latex).find();
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
