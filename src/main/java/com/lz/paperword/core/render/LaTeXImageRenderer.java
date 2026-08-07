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
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.Base64;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.UUID;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * LaTeX 公式图片渲染器。
 *
 * <p>OLE 预览图使用 MathJax 排版 SVG，再由 Batik 展开为字形轮廓并编码为
 * 纯 {@code POLYPOLYGON} placeable WMF。普通图片模式仍保留原生 TeX/JLaTeXMath 通道。</p>
 *
 * <p>OLE 预览图使用严格 MathJax/WMF 链路，失败即失败，不再静默回退为纯文本占位。</p>
 */
public class LaTeXImageRenderer {

    private static final Logger log = LoggerFactory.getLogger(LaTeXImageRenderer.class);

    /** 默认公式字体大小（磅），对应 Word 中正文公式的标准尺寸。 */
    private static final float DEFAULT_SIZE = 13f;

    /** JLaTeXMath 兜底通道的渲染缩放因子，用于提高位图清晰度。 */
    private static final float RENDER_SCALE = 4.0f;

    /** 原生 TeX 预览图转 PNG 时的放大倍率，只提高底图分辨率，不改变文档显示尺寸。 */
    private static final float PNG_OUTPUT_SCALE = 4.0f;

    /** MathJax 生成 OLE 预览图时的字号，来自 xsc/MathType 尺寸拟合。 */
    private static final float OLE_PREVIEW_SIZE = 9.02f;

    /** 系统属性：MathType 几何重排开关（分式全尺寸槽 + 实测垂直间距），默认开启。 */
    private static final String MATHJAX_MATHTYPE_FIT_PROP = "paperword.mathjax.mathtypefit";
    /** 系统属性：MathType 几何重排使用的字号（MathType 全尺寸槽的实测目标字号）。 */
    private static final String MATHJAX_MATHTYPE_FIT_FONT_PT_PROP = "paperword.mathjax.mathtypefit.fontpt";
    /** MathType 全尺寸槽的实测目标字号（display-scales 拟合：9.02 × 1.163）。 */
    private static final double MATHJAX_MATHTYPE_FIT_DEFAULT_FONT_PT = 12.0d;

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
    /** 系统属性：Node.js 命令路径。 */
    private static final String MATHJAX_NODE_CMD_PROP = "paperword.mathjax.node.command";
    /** 系统属性：MathJax worker 脚本路径。 */
    private static final String MATHJAX_SCRIPT_PROP = "paperword.mathjax.script";
    /** 系统属性：MathJax ex/pt 比例。 */
    private static final String MATHJAX_EX_RATIO_PROP = "paperword.mathjax.exRatio";
    /** 系统属性：MathJax WMF 预览内边距，单位 pt。 */
    private static final String MATHJAX_PADDING_PT_PROP = "paperword.mathjax.paddingPt";
    /** 系统属性：MathJax WMF 预览最大宽度，单位 pt。 */
    private static final String MATHJAX_MAX_WIDTH_PT_PROP = "paperword.mathjax.maxWidthPt";
    private static final double MATHJAX_DEFAULT_EX_RATIO = 0.431d;
    private static final double MATHJAX_DEFAULT_PADDING_PT = 2.3d;
    private static final double MATHJAX_DEFAULT_MAX_WIDTH_PT = 400.0d;
    /** 缓存版本，公式渲染度量或图片生成逻辑变化时递增。 */
    private static final String CACHE_VERSION = "v302-batik-vector-wmf-ink";
    private static final String EXPECTED_NODE_VERSION = "v24.9.0";
    private static final String EXPECTED_MATHJAX_VERSION = "3.2.2";
    private static final String EXPECTED_MATHJAX_BUNDLE_HASH =
        "c2b620e67102608a19c6c9729960cfa96fd14c9fe1b125b23def61dd544faf37";
    /** 外部命令默认超时秒数。 */
    private static final int DEFAULT_TIMEOUT_SECONDS = 20;
    private static final List<String> ARRAY_LIKE_ENVIRONMENTS = List.of(
        "array", "aligned", "alignedat", "gathered", "matrix", "pmatrix", "bmatrix", "cases");
    /** 像素到磅的换算比例。 */
    private static final float PX_PER_PT = 1.0f / 0.75f;
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

    private static final Object MATHJAX_WORKER_LOCK = new Object();
    private static Process mathJaxWorkerProcess;
    private static BufferedWriter mathJaxWorkerInput;
    private static BufferedReader mathJaxWorkerOutput;
    private static long mathJaxRequestId = 0L;

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
        if (cached != null && isRequestedPreviewFormat(cached)) {
            return cached;
        }
        cached = readPreviewFromDisk(cacheKey);
        if (cached != null && isRequestedPreviewFormat(cached)) {
            PREVIEW_CACHE.put(cacheKey, cached);
            return cached;
        }
        PreviewImage preview = renderWmfPreviewViaTeX(latex, OLE_PREVIEW_SIZE);
        if (preview != null) {
            if (isRequestedPreviewFormat(preview)) {
                PREVIEW_CACHE.put(cacheKey, preview);
                writePreviewToDisk(cacheKey, preview);
            }
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
        if (cached != null && isRequestedPreviewFormat(cached)) {
            return cached;
        }
        cached = readPreviewFromDisk(cacheKey);
        if (cached != null && isRequestedPreviewFormat(cached)) {
            PREVIEW_CACHE.put(cacheKey, cached);
            return cached;
        }
        PreviewImage preview = renderWmfPreviewViaTeX(latex, OLE_PREVIEW_SIZE, targetWidthPt, targetHeightPt);
        if (preview != null) {
            if (isRequestedPreviewFormat(preview)) {
                PREVIEW_CACHE.put(cacheKey, preview);
                writePreviewToDisk(cacheKey, preview);
            }
            return preview;
        }
        throw new IllegalStateException("Native TeX/WMF target preview rendering failed: " + latex);
    }

    /**
     * 校验预览产物是否满足当前请求的格式。
     *
     * <p>请求 EMF 而产物是回退的 WMF 时（EMF 后端一次性失败所致），
     * 缓存会把临时故障固化为永久行为：历史污染条目按未命中处理，
     * 新的回退产物也不再写入缓存。</p>
     */
    private boolean isRequestedPreviewFormat(PreviewImage preview) {
        return preview != null && "wmf".equals(preview.extension())
            && !SvgVectorWmfRenderer.containsBitmapRecord(preview.data());
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
        return CACHE_VERSION + "|" + renderConfigKey() + "|" + mode + "|" + size + "|"
            + normalizeLatexForLocalRender(latex == null ? "" : latex);
    }

    private String renderConfigKey() {
        return "latex=" + System.getProperty(LATEX_CMD_PROP, "")
            + "|xelatex=" + System.getProperty(XELATEX_CMD_PROP, "")
            + "|dvisvgm=" + System.getProperty(DVISVGM_CMD_PROP, "")
            + "|timeout=" + System.getProperty(RENDER_TIMEOUT_PROP, String.valueOf(DEFAULT_TIMEOUT_SECONDS))
            + "|oleVectorBackend=batik-polypolygon-v1"
            + "|vectorFontSet=" + BundledVectorFonts.FONT_SET_ID
            + "|mathjaxNode=" + System.getProperty(MATHJAX_NODE_CMD_PROP, "node")
            + "|mathjaxScript=" + System.getProperty(MATHJAX_SCRIPT_PROP, "tools/mathjax/render_mathjax_svg.cjs")
            + "|mathjaxBundle=" + EXPECTED_MATHJAX_BUNDLE_HASH
            + "|mathjaxFontPt=" + OLE_PREVIEW_SIZE
            + "|mathjaxMathTypeFit=" + mathJaxMathTypeFit()
            + "|mathjaxMathTypeFitFontPt=" + mathJaxMathTypeFitFontPt()
            + "|mathjaxExRatio=" + mathJaxExRatio()
            + "|mathjaxPaddingPt=" + mathJaxPaddingPt()
            + "|mathjaxMaxWidthPt=" + mathJaxMaxWidthPt();
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
            Properties props = new Properties();
            props.setProperty("widthPx", Integer.toString(preview.widthPx()));
            props.setProperty("heightPx", Integer.toString(preview.heightPx()));
            props.setProperty("extension", preview.extension());
            props.setProperty("contentType", preview.contentType());
            props.setProperty("placeholder", Boolean.toString(preview.placeholder()));
            props.setProperty("depthPt", Double.toString(preview.depthPt()));
            props.setProperty("widthPt", Double.toString(preview.widthPt()));
            props.setProperty("heightPt", Double.toString(preview.heightPt()));
            writeAtomically(imagePath, preview.data());
            Path tempMeta = tempSibling(metaPath);
            try (var out = Files.newOutputStream(tempMeta)) {
                props.store(out, "paperword formula preview cache");
            }
            moveAtomically(tempMeta, metaPath);
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
            writeAtomically(pngPath, png);
        } catch (IOException e) {
            log.debug("Formula PNG disk cache write failed: {}", cacheKey, e);
        }
    }

    private void writeAtomically(Path target, byte[] data) throws IOException {
        Path temp = tempSibling(target);
        Files.write(temp, data);
        moveAtomically(temp, target);
    }

    private Path tempSibling(Path target) {
        return target.resolveSibling(target.getFileName() + "." + UUID.randomUUID() + ".tmp");
    }

    private void moveAtomically(Path source, Path target) throws IOException {
        try {
            Files.move(source, target, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
        } catch (AtomicMoveNotSupportedException e) {
            Files.move(source, target, StandardCopyOption.REPLACE_EXISTING);
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

    /** OLE 对象预览图的严格 MathJax/WMF 入口。 */
    private PreviewImage renderWmfPreviewViaTeX(String latex, float size) {
        return renderWmfPreviewViaTeX(latex, size, null, null);
    }

    private PreviewImage renderWmfPreviewViaTeX(String latex, float size, Double targetWidthPt, Double targetHeightPt) {
        try {
            PreviewImage preview = renderMathJaxWmfPreview(latex, targetWidthPt, targetHeightPt);
            if (preview != null) {
                return preview;
            }
            throw new IllegalStateException("MathJax WMF preview returned no image: " + latex);
        } catch (Exception e) {
            log.error("MathJax WMF preview render failed: {}", latex, e);
            throw new IllegalStateException("Strict vector WMF preview failed for LaTeX: " + latex, e);
        }
    }

    private PreviewImage renderMathJaxWmfPreview(String latex, Double targetWidthPt, Double targetHeightPt)
        throws Exception {
        String localRenderLatex = normalizeLatexForLocalRender(latex);
        MathJaxSvgResult svg = renderSvgViaMathJax(localRenderLatex);
        double widthPt = svg.widthPt();
        double heightPt = svg.heightPt();
        double depthPt = mathJaxMathTypeFit()
            ? svg.depthPt()
            : calibrateMathJaxDepthPt(latex, heightPt, svg.depthPt());
        if (targetWidthPt != null && targetHeightPt != null && targetWidthPt > 0d && targetHeightPt > 0d) {
            widthPt = targetWidthPt;
            heightPt = targetHeightPt;
            depthPt = targetDepthPt(latex, heightPt);
        } else {
            double maxWidthPt = genericVectorWidthCapPt(latex);
            double scale = Math.min(1.0d, maxWidthPt / Math.max(widthPt, 1.0d));
            widthPt *= scale;
            heightPt *= scale;
            depthPt *= scale;
            if (scale < 1.0d && isLongLinearFormula(latex)) {
                widthPt = maxWidthPt;
                heightPt = Math.rint(heightPt * 2.0d) / 2.0d;
            }
        }
        SvgVectorWmfRenderer.VectorWmfResult vector =
            SvgVectorWmfRenderer.renderDetailed(svg.svgBytes(), widthPt, heightPt);
        if (vector.recordSummary().bitmapRecords() != 0 || vector.recordSummary().textRecords() != 0) {
            throw new IOException("Vector encoder emitted forbidden bitmap/text records");
        }
        int widthPx = Math.max((int) Math.round(widthPt * PX_PER_PT), 4);
        int heightPx = Math.max((int) Math.round(heightPt * PX_PER_PT), 4);
        return new PreviewImage(vector.bytes(), widthPx, heightPx, "wmf", "image/x-wmf", false,
            depthPt, widthPt, heightPt);
    }

    private BufferedImage addSafetyBorderIfInkTouchesEdge(BufferedImage source, int borderPx) {
        if (source == null || borderPx <= 0 || !inkTouchesRasterEdge(source)) {
            return source;
        }
        BufferedImage padded = new BufferedImage(
            source.getWidth() + borderPx * 2,
            source.getHeight() + borderPx * 2,
            BufferedImage.TYPE_INT_ARGB
        );
        Graphics2D graphics = padded.createGraphics();
        try {
            graphics.setColor(Color.WHITE);
            graphics.fillRect(0, 0, padded.getWidth(), padded.getHeight());
            graphics.drawImage(source, borderPx, borderPx, null);
        } finally {
            graphics.dispose();
        }
        return padded;
    }

    private boolean inkTouchesRasterEdge(BufferedImage image) {
        int width = image.getWidth();
        int height = image.getHeight();
        if (width <= 0 || height <= 0) {
            return false;
        }
        for (int x = 0; x < width; x++) {
            if (isInkPixel(image.getRGB(x, 0)) || isInkPixel(image.getRGB(x, height - 1))) {
                return true;
            }
        }
        for (int y = 1; y < height - 1; y++) {
            if (isInkPixel(image.getRGB(0, y)) || isInkPixel(image.getRGB(width - 1, y))) {
                return true;
            }
        }
        return false;
    }

    private boolean isInkPixel(int argb) {
        int alpha = (argb >>> 24) & 0xFF;
        if (alpha == 0) {
            return false;
        }
        int red = (argb >>> 16) & 0xFF;
        int green = (argb >>> 8) & 0xFF;
        int blue = argb & 0xFF;
        return Math.min(red, Math.min(green, blue)) < 245;
    }

    private double calibrateMathJaxDepthPt(String latex, double heightPt, double depthPt) {
        if (heightPt <= 0d) {
            return depthPt;
        }
        String text = latex == null ? "" : latex;
        double adjusted = depthPt;
        if (hasFractionCommand(text)) {
            adjusted = Math.max(adjusted, heightPt * 0.75d);
        } else if (text.contains("\\sqrt") || hasArrayLikeEnvironment(text)) {
            adjusted = Math.max(adjusted, heightPt * 0.38d);
        } else if (hasScript(text)) {
            adjusted = Math.max(adjusted, heightPt * 0.26d);
        }
        return Math.min(Math.max(adjusted, 0d), Math.max(heightPt - 1.0d, 0d));
    }

    private double estimateVectorWidthPt(String latex) {
        String text = latex == null ? "" : latex.replaceAll("\\\\pwmetrics\\{[^}]+}\\s*", "");
        text = text.replaceAll("\\\\pwstyle\\{[^}]*}\\s*", "");
        int visible = Math.max(text.replaceAll("\\\\[A-Za-z]+", "x").replaceAll("[{}\\s]", "").length(), 1);
        return Math.max(12.0d, visible * 6.0d);
    }

    private double genericVectorWidthCapPt(String latex) {
        if (isLongLinearFormula(latex)) {
            return 200.0d;
        }
        if (latex == null || !latex.contains("\\begin{array}")) {
            return 430.0d;
        }
        if (latex.contains("\\searrow") || latex.contains("\\nearrow")) {
            return 140.0d;
        }
        if (hasFractionCommand(latex) || latex.contains("\\cdots")) {
            return 420.0d;
        }
        return 300.0d;
    }

    private boolean isLongLinearFormula(String latex) {
        if (latex == null || latex.isBlank()) {
            return false;
        }
        boolean structured = latex.contains("\\frac") || latex.contains("\\sqrt")
            || latex.contains("\\begin") || latex.indexOf('^') >= 0 || latex.indexOf('_') >= 0;
        String atoms = latex.replaceAll("\\\\[A-Za-z]+", "x")
            .replaceAll("[{}\\s]", "");
        return !structured && atoms.codePointCount(0, atoms.length()) >= 20;
    }

    private double estimateVectorHeightPt(String latex) {
        return classifyStructureFamily(latex).heightPt();
    }

    private double targetDepthPt(String latex, double heightPt) {
        if (latex == null || heightPt <= 0d) {
            return -1d;
        }
        String text = latex.replaceAll("\\\\pwmetrics\\{[^}]+}\\s*", "");
        if (hasFractionCommand(text)) {
            return Math.max(0.0d, heightPt * MathTypeStructureMetrics.FRACTION_DEPTH_RATIO);
        }
        if (text.contains("\\begin{array}") || text.contains("\\sqrt")) {
            return Math.max(0.0d, heightPt * 0.32d);
        }
        if (hasScript(text)) {
            return Math.max(0.0d, heightPt * 0.24d);
        }
        return Math.max(0.0d, heightPt * 0.22d);
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
        MathTypeStructureMetrics.FamilyMetrics structure = classifyStructureFamily(latex);
        boolean hasArray = hasArrayLikeEnvironment(latex);
        int arrayCount = countArrayLikeEnvironments(latex);
        int lineBreaks = countOccurrences(latex, "\\\\");
        int fractionCount = countFractionCommands(latex);
        boolean script = hasScript(latex);
        boolean textHeavyFraction = hasTextHeavyFraction(latex);
        boolean nestedFraction = hasNestedFraction(latex);
        String previewClass = structure.previewClass();
        boolean sourceSeededHeight = structure.sourceSeededHeight();

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
        } else if (textHeavyFraction) {
            widthScale *= 1.08d;
            heightScale *= 0.62d;
        } else if (fractionCount == 1) {
            widthScale *= 0.96d;
            heightScale *= 0.94d;
        } else if (nestedFraction) {
            widthScale *= 0.92d;
            heightScale *= 1.02d;
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

        widthScale *= previewWidthClassScale(latex, structure.family(), previewClass);

        if (sourceSeededHeight) {
            heightScale = 1.0d;
        }
        double adjustedDepth = depthPt >= 0d ? depthPt * heightScale : depthPt;
        return new PreviewMetrics(
            Math.max(widthPt * widthScale, 1.0d),
            Math.max(heightPt * heightScale, 1.0d),
            adjustedDepth
        );
    }

    private MathTypeStructureMetrics.FamilyMetrics classifyStructureFamily(String latex) {
        String text = latex == null ? "" : latex;
        if (hasArrayLikeEnvironment(text)) {
            long rows = countFirstArrayLikeTopLevelRows(text);
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.ARRAY, rows);
        }
        if (text.contains("\\sqrt")) {
            int sqrtDepth = sqrtCommandDepthOutsideText(text);
            if (sqrtDepth > 1 && hasFractionCommand(text)) {
                return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SQRT_NESTED);
            }
            if (sqrtDepth > 2) {
                return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SQRT_NESTED);
            }
            if (hasFractionCommand(text)) {
                return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SQRT_FRACTION);
            }
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SQRT);
        }
        if (hasOnlyScriptSlotFractions(text)) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SCRIPT_FRACTION);
        }
        if (hasTextHeavyFraction(text)) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.TEXT_FRACTION);
        }
        if (hasScriptSlotFraction(text) && hasTopLevelFraction(text)) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SCRIPT_FRACTION_MIXED);
        }
        if (hasNestedFraction(text)) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.NESTED_FRACTION);
        }
        if (hasFractionCommand(text)) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.ORDINARY_FRACTION);
        }
        if (hasScript(text)) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SCRIPT);
        }
        if (text.contains("\\overline") || text.contains("\\underline")
            || text.contains("\\overset") || text.contains("\\underset")) {
            return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.ACCENT);
        }
        return MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.LINEAR);
    }

    private double previewWidthClassScale(String latex, MathTypeStructureMetrics.Family family, String previewClass) {
        String propertyName = "paperword.preview.width.scale." + previewClass;
        double defaultScale = switch (family) {
            case LINEAR -> MathTypeStructureMetrics.LINEAR_PREVIEW_WIDTH_SCALE;
            case SQRT -> hasScript(latex)
                ? MathTypeStructureMetrics.SQRT_SCRIPT_PREVIEW_WIDTH_SCALE
                : MathTypeStructureMetrics.SQRT_PREVIEW_WIDTH_SCALE;
            case SQRT_FRACTION -> !hasTopLevelFraction(latex) && hasTopLevelTextOutsideSqrt(latex)
                ? MathTypeStructureMetrics.SQRT_FRACTION_MIXED_PREVIEW_WIDTH_SCALE
                : MathTypeStructureMetrics.SQRT_FRACTION_PREVIEW_WIDTH_SCALE;
            case SQRT_NESTED -> sqrtCommandDepthOutsideText(latex) > 1
                && hasFractionCommand(latex)
                    ? MathTypeStructureMetrics.SQRT_NESTED_FRACTION_PREVIEW_WIDTH_SCALE
                    : 0.86d;
            default -> switch (previewClass) {
            case "array" -> 0.91d;
            case "textFraction", "text_fraction" -> 1.0d;
            case "fraction", "nested_fraction" -> 0.88d;
            case "script", "script_fraction", "script_fraction_mixed" -> 0.85d;
            case "accent" -> 0.90d;
            default -> 0.86d;
            };
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

    private static boolean hasTopLevelTextOutsideSqrt(String latex) {
        if (latex == null || latex.isBlank()) {
            return false;
        }
        StringBuilder outside = new StringBuilder();
        int cursor = 0;
        while (cursor < latex.length()) {
            int start = latex.indexOf("\\sqrt", cursor);
            if (start < 0) {
                outside.append(latex.substring(cursor));
                break;
            }
            outside.append(latex, cursor, start);
            int groupStart = skipWhitespace(latex, start + "\\sqrt".length());
            if (groupStart >= latex.length() || latex.charAt(groupStart) != '{') {
                cursor = start + "\\sqrt".length();
                continue;
            }
            int groupEnd = matchingBrace(latex, groupStart);
            if (groupEnd < 0) {
                break;
            }
            cursor = groupEnd + 1;
        }
        String normalized = normalizeTextForLength(outside.toString())
            .replaceAll("\\\\(?:left|right|displaystyle|textstyle|scriptstyle|scriptscriptstyle)\\b\\s*\\.?", "")
            .replaceAll("[\\s{}\\[\\]()（）,，.。:：;；]", "");
        return normalized.codePoints().anyMatch(Character::isLetterOrDigit);
    }

    private MathJaxSvgResult renderSvgViaMathJax(String latex)
        throws IOException, InterruptedException, ExecutionException, TimeoutException {
        synchronized (MATHJAX_WORKER_LOCK) {
            ensureMathJaxWorker();
            long id = ++mathJaxRequestId;
            String latexBase64 = Base64.getEncoder().encodeToString((latex == null ? "" : latex)
                .getBytes(StandardCharsets.UTF_8));
            boolean mathTypeFit = mathJaxMathTypeFit();
            double fontPt = mathTypeFit ? mathJaxMathTypeFitFontPt() : (double) OLE_PREVIEW_SIZE;
            String request = String.format(Locale.ROOT,
                "{\"id\":%d,\"latexBase64\":\"%s\",\"fontPt\":%.6f,\"exRatio\":%.6f,\"paddingPt\":%.6f,\"maxWidthPt\":%.6f,\"mathTypeFit\":%s}",
                id, latexBase64, fontPt, mathJaxExRatio(), mathJaxPaddingPt(), mathJaxMaxWidthPt(), mathTypeFit);
            mathJaxWorkerInput.write(request);
            mathJaxWorkerInput.newLine();
            mathJaxWorkerInput.flush();
            String response = readMathJaxResponseLine();
            long responseId = (long) jsonNumber(response, "id", -1d);
            if (responseId != id) {
                stopMathJaxWorker();
                throw new IOException("MathJax worker response id mismatch: " + response);
            }
            if (!jsonBoolean(response, "ok")) {
                String error = jsonString(response, "error", "unknown MathJax error");
                throw new IOException(error);
            }
            String engine = jsonString(response, "engine", "");
            String mathJaxVersion = jsonString(response, "mathJaxVersion", "");
            String nodeVersion = jsonString(response, "nodeVersion", "");
            String bundleHash = jsonString(response, "bundleHash", "");
            if (!"mathjax-svg".equals(engine)
                    || !EXPECTED_NODE_VERSION.equals(nodeVersion)
                    || !EXPECTED_MATHJAX_VERSION.equals(mathJaxVersion)
                    || !EXPECTED_MATHJAX_BUNDLE_HASH.equals(bundleHash)) {
                stopMathJaxWorker();
                throw new IOException("MathJax worker version mismatch: engine=" + engine
                    + ", node=" + nodeVersion + ", MathJax=" + mathJaxVersion
                    + ", bundle=" + bundleHash + "; expected node=" + EXPECTED_NODE_VERSION
                    + ", MathJax=" + EXPECTED_MATHJAX_VERSION
                    + ", bundle=" + EXPECTED_MATHJAX_BUNDLE_HASH);
            }
            String svgBase64 = jsonString(response, "svgBase64", "");
            if (svgBase64.isBlank()) {
                throw new IOException("MathJax worker returned empty SVG");
            }
            byte[] svgBytes = Base64.getDecoder().decode(svgBase64);
            double widthPt = jsonNumber(response, "widthPt", 12d);
            double heightPt = jsonNumber(response, "heightPt", 12d);
            double depthPt = jsonNumber(response, "depthPt", -1d);
            return new MathJaxSvgResult(svgBytes, widthPt, heightPt, depthPt);
        }
    }

    MathJaxSvgResult renderMathJaxSvgForAcceptance(String latex)
        throws IOException, InterruptedException, ExecutionException, TimeoutException {
        return renderSvgViaMathJax(normalizeLatexForLocalRender(latex));
    }

    private void ensureMathJaxWorker() throws IOException {
        if (mathJaxWorkerProcess != null && mathJaxWorkerProcess.isAlive()
            && mathJaxWorkerInput != null && mathJaxWorkerOutput != null) {
            return;
        }
        stopMathJaxWorker();
        Path script = mathJaxScriptPath();
        if (!Files.isRegularFile(script)) {
            throw new IOException("MathJax worker script not found: " + script);
        }
        ProcessBuilder pb = new ProcessBuilder(mathJaxNodeCommand(), script.toString(), "--worker");
        pb.directory(Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().toFile());
        mathJaxWorkerProcess = pb.start();
        mathJaxWorkerInput = new BufferedWriter(new OutputStreamWriter(
            mathJaxWorkerProcess.getOutputStream(), StandardCharsets.UTF_8));
        mathJaxWorkerOutput = new BufferedReader(new InputStreamReader(
            mathJaxWorkerProcess.getInputStream(), StandardCharsets.UTF_8));
        Thread stderrDrainer = new Thread(() -> {
            try (var reader = new BufferedReader(new InputStreamReader(
                mathJaxWorkerProcess.getErrorStream(), StandardCharsets.UTF_8))) {
                while (reader.readLine() != null) {
                    // Drain only. Worker failures are reported through JSON responses.
                }
            } catch (IOException ignored) {
            }
        }, "paperword-mathjax-stderr");
        stderrDrainer.setDaemon(true);
        stderrDrainer.start();
    }

    private String readMathJaxResponseLine()
        throws InterruptedException, ExecutionException, TimeoutException, IOException {
        int timeoutSeconds = Integer.getInteger(RENDER_TIMEOUT_PROP, DEFAULT_TIMEOUT_SECONDS);
        CompletableFuture<String> responseFuture = CompletableFuture.supplyAsync(() -> {
            try {
                return mathJaxWorkerOutput.readLine();
            } catch (IOException e) {
                throw new IllegalStateException(e);
            }
        });
        String line;
        try {
            line = responseFuture.get(timeoutSeconds, TimeUnit.SECONDS);
        } catch (TimeoutException e) {
            // The readLine task stays blocked and the worker may still write the
            // late response later: the stale task would then consume the NEXT
            // request's reply and desync the id pairing. Cancel the task and
            // destroy the worker so the next request starts a fresh process.
            responseFuture.cancel(true);
            stopMathJaxWorker();
            throw e;
        }
        if (line == null) {
            stopMathJaxWorker();
            throw new IOException("MathJax worker exited without response");
        }
        return line;
    }

    private static void stopMathJaxWorker() {
        closeQuietly(mathJaxWorkerInput);
        closeQuietly(mathJaxWorkerOutput);
        if (mathJaxWorkerProcess != null) {
            mathJaxWorkerProcess.destroyForcibly();
        }
        mathJaxWorkerProcess = null;
        mathJaxWorkerInput = null;
        mathJaxWorkerOutput = null;
    }

    private static void closeQuietly(AutoCloseable closeable) {
        if (closeable == null) {
            return;
        }
        try {
            closeable.close();
        } catch (Exception ignored) {
        }
    }

    private static String mathJaxNodeCommand() {
        String configured = System.getProperty(MATHJAX_NODE_CMD_PROP, "").trim();
        if (!configured.isEmpty()) {
            return configured;
        }
        Path root = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath();
        Path bundled = java.io.File.separatorChar == '\\'
            ? root.resolve("vector-sidecar/node/node.exe")
            : root.resolve("vector-sidecar/node/bin/node");
        return Files.isRegularFile(bundled) ? bundled.toString() : "node";
    }

    private static Path mathJaxScriptPath() {
        String configuredValue = System.getProperty(MATHJAX_SCRIPT_PROP, "").trim();
        Path root = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath();
        if (configuredValue.isEmpty()) {
            Path bundled = root.resolve("vector-sidecar/tools/mathjax/render_mathjax_svg.cjs");
            if (Files.isRegularFile(bundled)) {
                return bundled.normalize();
            }
            configuredValue = "tools/mathjax/render_mathjax_svg.cjs";
        }
        Path configured = Path.of(configuredValue);
        if (configured.isAbsolute()) {
            return configured;
        }
        return root.resolve(configured).normalize();
    }

    private static double mathJaxExRatio() {
        return readDoubleSystemProperty(MATHJAX_EX_RATIO_PROP, MATHJAX_DEFAULT_EX_RATIO);
    }

    private static boolean mathJaxMathTypeFit() {
        return Boolean.parseBoolean(System.getProperty(MATHJAX_MATHTYPE_FIT_PROP, "false"));
    }

    private static double mathJaxMathTypeFitFontPt() {
        return readDoubleSystemProperty(MATHJAX_MATHTYPE_FIT_FONT_PT_PROP, MATHJAX_MATHTYPE_FIT_DEFAULT_FONT_PT);
    }

    private static double mathJaxPaddingPt() {
        return readDoubleSystemProperty(MATHJAX_PADDING_PT_PROP, MATHJAX_DEFAULT_PADDING_PT);
    }

    private static double mathJaxMaxWidthPt() {
        return readDoubleSystemProperty(MATHJAX_MAX_WIDTH_PT_PROP, MATHJAX_DEFAULT_MAX_WIDTH_PT);
    }

    private static double readDoubleSystemProperty(String name, double fallback) {
        String configured = System.getProperty(name);
        if (configured == null || configured.isBlank()) {
            return fallback;
        }
        try {
            return Double.parseDouble(configured.trim());
        } catch (NumberFormatException e) {
            return fallback;
        }
    }

    private static boolean jsonBoolean(String json, String name) {
        Matcher matcher = Pattern.compile("\"" + Pattern.quote(name) + "\"\\s*:\\s*(true|false)")
            .matcher(json == null ? "" : json);
        return matcher.find() && Boolean.parseBoolean(matcher.group(1));
    }

    private static double jsonNumber(String json, String name, double fallback) {
        Matcher matcher = Pattern.compile("\"" + Pattern.quote(name) + "\"\\s*:\\s*(-?[0-9]+(?:\\.[0-9]+)?)")
            .matcher(json == null ? "" : json);
        return matcher.find() ? Double.parseDouble(matcher.group(1)) : fallback;
    }

    private static String jsonString(String json, String name, String fallback) {
        Matcher matcher = Pattern.compile("\"" + Pattern.quote(name) + "\"\\s*:\\s*\"([^\"]*)\"")
            .matcher(json == null ? "" : json);
        if (!matcher.find()) {
            return fallback;
        }
        return matcher.group(1)
            .replace("\\\"", "\"")
            .replace("\\\\", "\\");
    }

    private static int sqrtCommandDepthOutsideText(String latex) {
        if (latex == null || latex.isBlank()) {
            return 0;
        }
        int maxDepth = 0;
        for (int i = 0; i < latex.length(); i++) {
            if (!latex.startsWith("\\sqrt", i) || isEscapedCommandCharacter(latex, i)) {
                continue;
            }
            maxDepth = Math.max(maxDepth, 1 + sqrtCommandDepthOutsideText(sqrtBody(latex, i + 5)));
        }
        return maxDepth;
    }

    private static boolean isEscapedCommandCharacter(String latex, int index) {
        return index > 0 && latex.charAt(index - 1) == '\\';
    }

    private static String sqrtBody(String latex, int cursor) {
        while (cursor < latex.length() && Character.isWhitespace(latex.charAt(cursor))) {
            cursor++;
        }
        if (cursor < latex.length() && latex.charAt(cursor) == '[') {
            int end = findBalancedEnd(latex, cursor, '[', ']');
            cursor = end > cursor ? end + 1 : cursor;
        }
        while (cursor < latex.length() && Character.isWhitespace(latex.charAt(cursor))) {
            cursor++;
        }
        if (cursor >= latex.length()) {
            return "";
        }
        if (latex.charAt(cursor) == '{') {
            int end = findBalancedEnd(latex, cursor, '{', '}');
            return end > cursor ? latex.substring(cursor + 1, end) : "";
        }
        if (latex.charAt(cursor) == '\\') {
            int end = cursor + 1;
            while (end < latex.length() && Character.isLetter(latex.charAt(end))) {
                end++;
            }
            return latex.substring(cursor, end);
        }
        return latex.substring(cursor, cursor + 1);
    }

    private static int findBalancedEnd(String text, int start, char open, char close) {
        int depth = 0;
        for (int i = start; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == open) {
                depth++;
            } else if (ch == close) {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static boolean hasArrayLikeEnvironment(String latex) {
        return countArrayLikeEnvironments(latex) > 0;
    }

    private static int countArrayLikeEnvironments(String latex) {
        if (latex == null || latex.isEmpty()) {
            return 0;
        }
        int count = 0;
        for (String env : ARRAY_LIKE_ENVIRONMENTS) {
            count += countOccurrences(latex, "\\begin{" + env + "}");
        }
        return count;
    }

    private static long countFirstArrayLikeTopLevelRows(String latex) {
        ArrayLikeBody body = firstArrayLikeBody(latex);
        if (body == null || body.body().isBlank()) {
            return 1L;
        }
        long rows = 1L;
        int braceDepth = 0;
        int nestedArrayDepth = 0;
        String text = body.body();
        for (int i = 0; i < text.length(); i++) {
            if (startsArrayLikeBegin(text, i)) {
                nestedArrayDepth++;
                continue;
            }
            if (startsArrayLikeEnd(text, i)) {
                nestedArrayDepth = Math.max(0, nestedArrayDepth - 1);
                continue;
            }
            char ch = text.charAt(i);
            if (ch == '\\') {
                if (i + 1 < text.length() && text.charAt(i + 1) == '\\' && braceDepth == 0
                    && nestedArrayDepth == 0) {
                    rows++;
                    i++;
                }
                continue;
            }
            if (ch == '{') {
                braceDepth++;
            } else if (ch == '}') {
                braceDepth = Math.max(0, braceDepth - 1);
            }
        }
        return rows;
    }

    private static ArrayLikeBody firstArrayLikeBody(String latex) {
        if (latex == null) {
            return null;
        }
        for (int i = 0; i < latex.length(); i++) {
            String env = arrayLikeEnvironmentAt(latex, i, true);
            if (env == null) {
                continue;
            }
            int bodyStart = i + "\\begin{".length() + env.length() + 1;
            if ("array".equals(env)) {
                bodyStart = skipWhitespace(latex, bodyStart);
                if (bodyStart < latex.length() && latex.charAt(bodyStart) == '{') {
                    int columnsEnd = matchingBrace(latex, bodyStart);
                    if (columnsEnd < 0) {
                        return null;
                    }
                    bodyStart = columnsEnd + 1;
                }
            }
            String endToken = "\\end{" + env + "}";
            int depth = 1;
            int cursor = bodyStart;
            while (cursor < latex.length()) {
                String nestedBegin = arrayLikeEnvironmentAt(latex, cursor, true);
                if (nestedBegin != null) {
                    depth++;
                    cursor += "\\begin{".length() + nestedBegin.length() + 1;
                    continue;
                }
                String nestedEnd = arrayLikeEnvironmentAt(latex, cursor, false);
                if (nestedEnd != null) {
                    depth--;
                    if (depth == 0 && latex.startsWith(endToken, cursor)) {
                        return new ArrayLikeBody(latex.substring(bodyStart, cursor));
                    }
                    cursor += "\\end{".length() + nestedEnd.length() + 1;
                    continue;
                }
                cursor++;
            }
        }
        return null;
    }

    private static boolean startsArrayLikeBegin(String latex, int index) {
        return arrayLikeEnvironmentAt(latex, index, true) != null;
    }

    private static boolean startsArrayLikeEnd(String latex, int index) {
        return arrayLikeEnvironmentAt(latex, index, false) != null;
    }

    private static String arrayLikeEnvironmentAt(String latex, int index, boolean begin) {
        String prefix = begin ? "\\begin{" : "\\end{";
        if (latex == null || !latex.startsWith(prefix, index)) {
            return null;
        }
        int nameStart = index + prefix.length();
        int nameEnd = latex.indexOf('}', nameStart);
        if (nameEnd < 0) {
            return null;
        }
        String env = latex.substring(nameStart, nameEnd);
        return ARRAY_LIKE_ENVIRONMENTS.contains(env) ? env : null;
    }

    private record ArrayLikeBody(String body) {
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

    private static boolean hasFractionCommand(String latex) {
        return nextFractionCommand(latex, 0) >= 0;
    }

    private static int countFractionCommands(String latex) {
        int count = 0;
        int cursor = 0;
        while ((cursor = nextFractionCommand(latex, cursor)) >= 0) {
            count++;
            cursor = fractionCommandEnd(latex, cursor);
        }
        return count;
    }

    private static int nextFractionCommand(String latex, int cursor) {
        if (latex == null) {
            return -1;
        }
        int best = -1;
        for (String command : List.of("\\frac", "\\dfrac", "\\cfrac")) {
            int hit = cursor;
            while ((hit = latex.indexOf(command, hit)) >= 0) {
                if (latexCommandBoundary(latex, hit + command.length())) {
                    if (best < 0 || hit < best) {
                        best = hit;
                    }
                    break;
                }
                hit += command.length();
            }
        }
        return best;
    }

    private static int fractionCommandEnd(String latex, int start) {
        if (latex == null || start < 0) {
            return start + "\\frac".length();
        }
        for (String command : List.of("\\frac", "\\dfrac", "\\cfrac")) {
            if (latex.startsWith(command, start) && latexCommandBoundary(latex, start + command.length())) {
                return start + command.length();
            }
        }
        return start + "\\frac".length();
    }

    private static boolean latexCommandBoundary(String latex, int cursor) {
        return cursor >= latex.length() || !Character.isLetter(latex.charAt(cursor));
    }

    private static boolean hasTextHeavyFraction(String latex) {
        if (latex == null || !hasFractionCommand(latex)) {
            return false;
        }
        int cursor = 0;
        while ((cursor = nextFractionCommand(latex, cursor)) >= 0) {
            int commandEnd = fractionCommandEnd(latex, cursor);
            int numeratorStart = skipWhitespace(latex, commandEnd);
            if (numeratorStart >= latex.length() || latex.charAt(numeratorStart) != '{') {
                cursor = commandEnd;
                continue;
            }
            int numeratorEnd = matchingBrace(latex, numeratorStart);
            if (numeratorEnd < 0) {
                return false;
            }
            int denominatorStart = skipWhitespace(latex, numeratorEnd + 1);
            if (denominatorStart >= latex.length() || latex.charAt(denominatorStart) != '{') {
                cursor = numeratorEnd + 1;
                continue;
            }
            int denominatorEnd = matchingBrace(latex, denominatorStart);
            if (denominatorEnd < 0) {
                return false;
            }
            String numerator = latex.substring(numeratorStart + 1, numeratorEnd);
            String denominator = latex.substring(denominatorStart + 1, denominatorEnd);
            if (isTextHeavyFractionPart(numerator) || isTextHeavyFractionPart(denominator)) {
                return true;
            }
            cursor = denominatorEnd + 1;
        }
        return false;
    }

    private static boolean hasNestedFraction(String latex) {
        if (latex == null || countFractionCommands(latex) < 2) {
            return false;
        }
        int cursor = 0;
        while ((cursor = nextFractionCommand(latex, cursor)) >= 0) {
            int commandEnd = fractionCommandEnd(latex, cursor);
            int numeratorStart = skipWhitespace(latex, commandEnd);
            if (numeratorStart >= latex.length() || latex.charAt(numeratorStart) != '{') {
                cursor = commandEnd;
                continue;
            }
            int numeratorEnd = matchingBrace(latex, numeratorStart);
            if (numeratorEnd < 0) {
                return false;
            }
            int denominatorStart = skipWhitespace(latex, numeratorEnd + 1);
            if (denominatorStart >= latex.length() || latex.charAt(denominatorStart) != '{') {
                cursor = numeratorEnd + 1;
                continue;
            }
            int denominatorEnd = matchingBrace(latex, denominatorStart);
            if (denominatorEnd < 0) {
                return false;
            }
            String numerator = latex.substring(numeratorStart + 1, numeratorEnd);
            String denominator = latex.substring(denominatorStart + 1, denominatorEnd);
            if (hasFractionCommand(numerator) || hasFractionCommand(denominator)) {
                return true;
            }
            cursor = denominatorEnd + 1;
        }
        return false;
    }

    private static boolean hasOnlyScriptSlotFractions(String latex) {
        return hasFractionCommand(latex) && hasScriptSlotFraction(latex) && !hasTopLevelFraction(latex);
    }

    private static boolean hasScriptSlotFraction(String latex) {
        if (latex == null || !hasFractionCommand(latex)) {
            return false;
        }
        int cursor = 0;
        while (cursor < latex.length()) {
            int op = nextScriptOperator(latex, cursor);
            if (op < 0) {
                return false;
            }
            int groupStart = skipWhitespace(latex, op + 1);
            if (groupStart >= latex.length()) {
                return false;
            }
            int groupEnd;
            String body;
            if (latex.charAt(groupStart) == '{') {
                groupEnd = matchingBrace(latex, groupStart);
                if (groupEnd < 0) {
                    return false;
                }
                body = latex.substring(groupStart + 1, groupEnd);
                cursor = groupEnd + 1;
            } else if (latex.charAt(groupStart) == '\\') {
                int commandEnd = skipCommandName(latex, groupStart + 1);
                body = latex.substring(groupStart, commandEnd);
                cursor = commandEnd;
            } else {
                body = latex.substring(groupStart, groupStart + 1);
                cursor = groupStart + 1;
            }
            if (hasFractionCommand(body)) {
                return true;
            }
        }
        return false;
    }

    private static boolean hasTopLevelFraction(String latex) {
        if (latex == null || !hasFractionCommand(latex)) {
            return false;
        }
        int depth = 0;
        for (int i = 0; i < latex.length(); i++) {
            char ch = latex.charAt(i);
            if (ch == '\\') {
                int commandEnd = skipCommandName(latex, i + 1);
                String command = latex.substring(i, commandEnd);
                if (depth == 0 && (command.equals("\\frac") || command.equals("\\dfrac")
                    || command.equals("\\cfrac"))) {
                    return true;
                }
                i = commandEnd - 1;
            } else if (ch == '{' || ch == '[' || ch == '(') {
                depth++;
            } else if (ch == '}' || ch == ']' || ch == ')') {
                depth = Math.max(0, depth - 1);
            }
        }
        return false;
    }

    private static int nextScriptOperator(String latex, int cursor) {
        int depth = 0;
        for (int i = Math.max(0, cursor); i < latex.length(); i++) {
            char ch = latex.charAt(i);
            if (ch == '\\') {
                i = skipCommandName(latex, i + 1) - 1;
                continue;
            }
            if (ch == '{' || ch == '[' || ch == '(') {
                depth++;
            } else if (ch == '}' || ch == ']' || ch == ')') {
                depth = Math.max(0, depth - 1);
            } else if (depth == 0 && (ch == '^' || ch == '_')) {
                return i;
            }
        }
        return -1;
    }

    private static int skipCommandName(String latex, int start) {
        int cursor = start;
        while (cursor < latex.length() && Character.isLetter(latex.charAt(cursor))) {
            cursor++;
        }
        return cursor;
    }

    private static boolean isTextHeavyFractionPart(String text) {
        String visible = normalizeTextForLength(text == null ? "" : text);
        return containsCjk(visible) && visible.codePointCount(0, visible.length()) >= 6;
    }

    private static String normalizeTextForLength(String text) {
        return text.replaceAll("\\\\(?:mathrm|mathbf|mathit|textit|textbf|emph|text|boldsymbol)\\s*\\{\\s*([^{}]*)\\s*}", "$1")
            .replaceAll("\\\\[A-Za-z]+", "x")
            .replaceAll("[{}\\s]", "");
    }

    private static int skipWhitespace(String text, int cursor) {
        while (cursor < text.length() && Character.isWhitespace(text.charAt(cursor))) {
            cursor++;
        }
        return cursor;
    }

    private static int matchingBrace(String text, int open) {
        int depth = 0;
        for (int i = open; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '{') {
                depth++;
            } else if (ch == '}') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }
        return -1;
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
            latex.replaceAll("\\\\kern\\s*[-+]?\\d*\\.?\\d+[a-zA-Z]+", ""),
            !mathJaxMathTypeFit());
        normalized = escapeRawUnicodeSymbolsForMathJax(normalized);
        normalized = normalizeLegacyBbbPreview(normalized);
        normalized = separateNestedRadicalDegreesForMathJax(normalized);
        normalized = normalizeGreekCapitalAliasesForPreview(normalized);
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

    private String normalizeLegacyBbbPreview(String latex) {
        return latex.replaceAll("\\\\(?:Bbb|mathbb)\\s+x(?![A-Za-z])", "x")
            .replaceAll("\\\\(?:Bbb|mathbb)\\s+([A-Za-z])", "\\\\mathbb{$1}");
    }

    private String separateNestedRadicalDegreesForMathJax(String latex) {
        StringBuilder normalized = new StringBuilder(latex.length() + 8);
        int cursor = 0;
        while (cursor < latex.length()) {
            int sqrt = latex.indexOf("\\sqrt[", cursor);
            if (sqrt < 0) {
                normalized.append(latex, cursor, latex.length());
                break;
            }
            normalized.append(latex, cursor, sqrt);
            int degreeStart = sqrt + 6;
            int bracketDepth = 1;
            boolean nestedBracket = false;
            int end = degreeStart;
            for (; end < latex.length(); end++) {
                char ch = latex.charAt(end);
                if (ch == '[') {
                    bracketDepth++;
                    nestedBracket = true;
                } else if (ch == ']') {
                    bracketDepth--;
                    if (bracketDepth == 0) {
                        break;
                    }
                }
            }
            if (end >= latex.length()) {
                normalized.append(latex, sqrt, latex.length());
                break;
            }
            String degree = latex.substring(degreeStart, end);
            int radicandStart = skipWhitespace(latex, end + 1);
            int radicandEnd = radicandStart < latex.length() && latex.charAt(radicandStart) == '{'
                ? matchingBrace(latex, radicandStart)
                : -1;
            if (nestedBracket && radicandEnd > radicandStart) {
                String radicand = latex.substring(radicandStart + 1, radicandEnd);
                normalized.append("{}^{").append(degree).append("}\\!\\sqrt{")
                    .append(radicand).append('}');
                cursor = radicandEnd + 1;
            } else {
                normalized.append(latex, sqrt, end + 1);
                cursor = end + 1;
            }
        }
        return normalized.toString();
    }

    /**
     * MathJax's SVG font tables do not expose every raw Unicode symbol to the
     * geometry fitter.  Preserve the code point while spelling symbols through
     * MathJax's supported Unicode command; letters (including CJK text) stay raw.
     */
    private String escapeRawUnicodeSymbolsForMathJax(String latex) {
        StringBuilder escaped = new StringBuilder(latex.length());
        latex.codePoints().forEach(codePoint -> {
            int type = Character.getType(codePoint);
            if (codePoint > Character.MAX_VALUE || (codePoint > 0x7F && (type == Character.CURRENCY_SYMBOL
                || type == Character.MATH_SYMBOL
                || type == Character.MODIFIER_SYMBOL
                || type == Character.OTHER_SYMBOL))) {
                escaped.append("\\unicode{x")
                    .append(Integer.toHexString(codePoint).toUpperCase(Locale.ROOT))
                    .append('}');
            } else {
                escaped.appendCodePoint(codePoint);
            }
        });
        return escaped.toString();
    }

    private String normalizeGreekCapitalAliasesForPreview(String latex) {
        Map<String, String> aliases = Map.ofEntries(
            Map.entry("Alpha", "A"), Map.entry("Beta", "B"),
            Map.entry("Epsilon", "E"), Map.entry("Zeta", "Z"),
            Map.entry("Eta", "H"), Map.entry("Iota", "I"),
            Map.entry("Kappa", "K"), Map.entry("Mu", "M"),
            Map.entry("Nu", "N"), Map.entry("Omicron", "O"),
            Map.entry("Rho", "P"), Map.entry("Tau", "T"),
            Map.entry("Chi", "X")
        );
        String normalized = latex;
        for (Map.Entry<String, String> alias : aliases.entrySet()) {
            normalized = normalized.replaceAll(
                "\\\\" + alias.getKey() + "(?![A-Za-z])",
                "\\\\mathrm{" + alias.getValue() + "}"
            );
        }
        return normalized;
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
        CompletableFuture<byte[]> outputFuture = CompletableFuture.supplyAsync(() -> {
            try {
                return process.getInputStream().readAllBytes();
            } catch (IOException e) {
                return new byte[0];
            }
        });
        boolean finished = process.waitFor(timeoutSeconds, TimeUnit.SECONDS);
        if (!finished) {
            process.destroyForcibly();
            byte[] output = outputFuture.completeOnTimeout(new byte[0], 2, TimeUnit.SECONDS).join();
            return new CommandResult(-1, output);
        }
        byte[] output = outputFuture.join();
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

    record MathJaxSvgResult(byte[] svgBytes, double widthPt, double heightPt, double depthPt) {
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
