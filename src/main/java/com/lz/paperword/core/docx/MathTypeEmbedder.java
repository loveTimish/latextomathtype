package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.latex.LaTeXParser.FormulaMetrics;
import com.lz.paperword.core.latex.LaTeXParser.FormulaStyleHints;
import com.lz.paperword.core.mtef.MtefWriter;
import com.lz.paperword.core.ole.OlePackager;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Pattern;

/**
 * MathType OLE 公式嵌入器 — 将 LaTeX 公式作为可编辑的 MathType OLE 对象嵌入 Word 文档。
 *
 * @see OlePackager   OLE2 复合文档打包
 * @see LaTeXImageRenderer  LaTeX 公式预览图渲染
 * @see MtefWriter    LaTeX AST → MTEF 二进制转换
 */
public class MathTypeEmbedder {

    private static final Logger log = LoggerFactory.getLogger(MathTypeEmbedder.class);

    private static final double PT_PER_PX = 0.75d;
    /** 版面安全上限：测试集公式最宽约 428pt，超出页面可用宽度才整体缩小。 */
    private static final double MAX_GENERIC_FORMULA_WIDTH_PT = 430.0d;
    private static final double MAX_DISPLAY_HEIGHT_PT = 230.0d;
    private static final Pattern TRACE_METRICS_PATTERN = Pattern.compile("^\\\\pwmetrics\\{[^}]+}[ \\t\\n\\x0B\\f\\r]*");
    private static final Pattern TRACE_STYLE_PATTERN = Pattern.compile("^\\\\pwstyle\\{[^}]*}[ \\t\\n\\x0B\\f\\r]*");
    private static final Pattern TRACE_EDGE_SPACE_PATTERN = Pattern.compile("^[ \\t\\n\\x0B\\f\\r]+|[ \\t\\n\\x0B\\f\\r]+$");
    private static final Pattern TRACE_SPACE_PATTERN = Pattern.compile("[ \\t\\n\\x0B\\f\\r]+");
    private final MtefWriter mtefWriter = new MtefWriter();
    private final OlePackager olePackager = new OlePackager();
    private final LaTeXImageRenderer imageRenderer = new LaTeXImageRenderer();
    private final AtomicInteger oleCounter = new AtomicInteger(1);

    /**
     * 将 MathType 公式嵌入 Word 段落的指定 run 中。
     */
    public void resetDocumentFormulaCounter() {
        oleCounter.set(1);
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex) {
        embedEquation(paragraph, run, latexAst, rawLatex, 1.0d);
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex,
                              double displayScale) {
        embedEquation(paragraph, run, latexAst, rawLatex, displayScale, Double.MAX_VALUE);
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex,
                              double displayScale, double maxWidthPt) {
        embedEquation(paragraph, run, latexAst, rawLatex, displayScale, maxWidthPt, null);
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex,
                              double displayScale, double maxWidthPt, FormulaMetrics targetMetrics) {
        embedEquation(paragraph, run, latexAst, rawLatex, displayScale, maxWidthPt, targetMetrics,
            FormulaStyleHints.empty());
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex,
                              double displayScale, double maxWidthPt, FormulaMetrics targetMetrics,
                              FormulaStyleHints styleHints) {
        try {
            FormulaStyleHints effectiveStyleHints = styleHints == null
                ? FormulaStyleHints.empty()
                : styleHints.withSourceMetrics(targetMetrics);
            byte[] mtefData = mtefWriter.write(latexAst, effectiveStyleHints);
            byte[] oleData = olePackager.packageOle(mtefData);

            LaTeXImageRenderer.PreviewImage preview = targetMetrics != null
                ? imageRenderer.renderForOlePreview(rawLatex, targetMetrics.wmfWidthPt(), targetMetrics.wmfHeightPt())
                : imageRenderer.renderForOlePreview(rawLatex);
            if (preview == null || preview.data() == null || preview.data().length == 0) {
                throw new IllegalStateException("OLE preview rendering returned no image data");
            }
            PreviewBox previewBox = targetMetrics != null
                ? new PreviewBox(preview.widthPx(), preview.heightPx())
                : constrainPreviewBox(rawLatex, preview.widthPx(), preview.heightPx(),
                    displayScale, maxWidthPt);

            OPCPackage pkg = paragraph.getDocument().getPackage();
            int idx = oleCounter.getAndIncrement();

            PackagePartName olePartName = PackagingURIHelper.createPartName(
                "/word/embeddings/oleObject" + idx + ".bin");
            PackagePart olePart = pkg.createPart(olePartName,
                "application/vnd.openxmlformats-officedocument.oleObject");
            try (var os = olePart.getOutputStream()) {
                os.write(oleData);
            }

            PackagePartName imgPartName = PackagingURIHelper.createPartName(
                "/word/media/image_eq" + idx + "." + preview.extension());
            PackagePart imgPart = pkg.createPart(imgPartName, preview.contentType());
            try (var os = imgPart.getOutputStream()) {
                os.write(preview.data());
            }

            PackagePart docPart = paragraph.getDocument().getPackagePart();
            PackageRelationship oleRel = docPart.addRelationship(
                olePartName, TargetMode.INTERNAL,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject");
            PackageRelationship imgRel = docPart.addRelationship(
                imgPartName, TargetMode.INTERNAL,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            double previewWidthPt = resolvePreviewWidthPt(preview, previewBox, targetMetrics);
            double previewHeightPt = resolvePreviewHeightPt(preview, previewBox, targetMetrics);
            double shapeWidthPt = targetMetrics != null && targetMetrics.shapeWidthPt() > 0d
                ? targetMetrics.shapeWidthPt()
                : previewWidthPt;
            double shapeHeightPt = targetMetrics != null && targetMetrics.shapeHeightPt() > 0d
                ? targetMetrics.shapeHeightPt()
                : previewHeightPt;
            double depthPt = resolveDisplayDepthPt(preview, previewHeightPt, shapeHeightPt);

            insertOleObjectXml(paragraph, run, oleRel.getId(), imgRel.getId(), idx,
                previewBox.widthPx(), previewBox.heightPx(),
                shapeWidthPt, shapeHeightPt,
                rawLatex, depthPt);

        } catch (Exception e) {
            throw new IllegalStateException("Failed to embed MathType equation: " + rawLatex, e);
        }
    }

    private void insertOleObjectXml(XWPFParagraph paragraph, XWPFRun run, String oleRelId, String imgRelId,
                                     int shapeIdx, int widthPx, int heightPx, double shapeWidthPt, double shapeHeightPt,
                                     String rawLatex, double depthPt) {
        try {
            widthPx = Math.max(widthPx, 4);
            heightPx = Math.max(heightPx, 4);

            String shapeId = "_x0000_i" + (1024 + shapeIdx);

            double targetShapeWidthPt = shapeWidthPt > 0d ? shapeWidthPt : widthPx * PT_PER_PX;
            double targetShapeHeightPt = shapeHeightPt > 0d ? shapeHeightPt : heightPx * PT_PER_PX;
            String styleWidth = String.format("%.3fpt", targetShapeWidthPt);
            String styleHeight = String.format("%.3fpt", targetShapeHeightPt);
            int dxaOrig = Math.max((int) Math.round(targetShapeWidthPt * 20), 1);
            int dyaOrig = Math.max((int) Math.round(targetShapeHeightPt * 20), 1);
            int posHalfPt = depthPt >= 0d
                ? Math.min(-(int) Math.round(depthPt * 2d), 0)
                : resolveRunPositionHalfPoints(rawLatex, targetShapeHeightPt);

            String objectId = "_" + Integer.toUnsignedString((shapeId + ":" + oleRelId).hashCode());
            String formulaTraceId = formulaTraceId(shapeIdx, rawLatex);

            // 参考文档 OLE run rPr 仅含 w:position，不含 w:rFonts。
            // w:position 负值 = 下移（半磅），用于补偿公式基线与文本基线的偏差。
            String runXml = "<w:r " +
                "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
                "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<w:rPr><w:position w:val=\"" + posHalfPt + "\"/></w:rPr>" +
                "<w:object w:dxaOrig=\"" + dxaOrig + "\" w:dyaOrig=\"" + dyaOrig + "\" " +
                "w14:anchorId=\"" + Integer.toHexString(objectId.hashCode()).toUpperCase() + "\">" +

                "<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" " +
                "path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">" +
                "<v:stroke joinstyle=\"miter\"/>" +
                "<v:formulas>" +
                "<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>" +
                "<v:f eqn=\"sum @0 1 0\"/>" +
                "<v:f eqn=\"sum 0 0 @1\"/>" +
                "<v:f eqn=\"prod @2 1 2\"/>" +
                "<v:f eqn=\"prod @3 21600 pixelWidth\"/>" +
                "<v:f eqn=\"prod @3 21600 pixelHeight\"/>" +
                "<v:f eqn=\"sum @0 0 1\"/>" +
                "<v:f eqn=\"prod @6 1 2\"/>" +
                "<v:f eqn=\"prod @7 21600 pixelWidth\"/>" +
                "<v:f eqn=\"sum @8 21600 0\"/>" +
                "<v:f eqn=\"prod @7 21600 pixelHeight\"/>" +
                "<v:f eqn=\"sum @10 21600 0\"/>" +
                "</v:formulas>" +
                "<v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>" +
                "<o:lock v:ext=\"edit\" aspectratio=\"t\"/>" +
                "</v:shapetype>" +

                // o:ole 必须为 "t"（VML 布尔真值）——空字符串是假值，
                // Word 会把形状当普通图片，双击不激活 MathType；WPS 不检查此属性
                "<v:shape id=\"" + shapeId + "\" o:spt=\"75\" type=\"#_x0000_t75\" " +
                "style=\"width:" + styleWidth + ";height:" + styleHeight + "\" " +
                "o:ole=\"t\" filled=\"f\" o:preferrelative=\"t\" stroked=\"f\" coordsize=\"21600,21600\">" +
                "<v:imagedata r:id=\"" + imgRelId + "\" o:title=\"" + xmlAttr(formulaTraceId) + "\"/>" +
                "</v:shape>" +

                "<o:OLEObject Type=\"Embed\" ProgID=\"Equation.DSMT4\" " +
                "ShapeID=\"" + shapeId + "\" DrawAspect=\"Content\" " +
                "ObjectID=\"" + objectId + "\" " +
                "r:id=\"" + oleRelId + "\" />" +
                "</w:object></w:r>";

            CTR replacement = CTR.Factory.parse(runXml);
            int runIndex = paragraph.getRuns().indexOf(run);
            if (runIndex < 0) {
                throw new IllegalStateException("Could not locate target run in paragraph");
            }
            paragraph.getCTP().setRArray(runIndex, replacement);

            XWPFRun spacerRun = paragraph.insertNewRun(runIndex + 1);
            spacerRun.setText(" ");

        } catch (Exception e) {
            log.error("Failed to insert OLE XML into run", e);
        }
    }

    private String formulaTraceId(int formulaIndex, String rawLatex) {
        String normalized = normalizeTraceLatex(rawLatex);
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            byte[] bytes = digest.digest(normalized.getBytes(StandardCharsets.UTF_8));
            StringBuilder out = new StringBuilder("pwf:");
            out.append(formulaIndex).append("-");
            for (int i = 0; i < 8 && i < bytes.length; i++) {
                out.append(String.format("%02x", bytes[i] & 0xFF));
            }
            return out.toString();
        } catch (NoSuchAlgorithmException e) {
            throw new IllegalStateException("SHA-256 is not available", e);
        }
    }

    private String normalizeTraceLatex(String rawLatex) {
        String value = rawLatex == null ? "" : rawLatex;
        value = value.replace('\u00A0', ' ');
        value = TRACE_METRICS_PATTERN.matcher(value).replaceFirst("");
        value = TRACE_STYLE_PATTERN.matcher(value).replaceFirst("");
        value = TRACE_EDGE_SPACE_PATTERN.matcher(value).replaceAll("");
        return TRACE_SPACE_PATTERN.matcher(value).replaceAll(" ");
    }

    private String xmlAttr(String value) {
        return value
            .replace("&", "&amp;")
            .replace("\"", "&quot;")
            .replace("<", "&lt;")
            .replace(">", "&gt;");
    }

    /**
     * 根据对象高度推算 w:position 值（半磅，负值=下移）。
     *
     * <p>基于测试集（word_files，4 万个 MathType 对象）的统计标定：
     * MathType 对象按垂直中心对齐数学轴（12pt 正文约在基线上 3pt），
     * 即下移量 ≈ h/2 - 3pt，换算半磅为 -(h - 6)。
     * 实测全类别（线性/上下标/分数/多行）误差不超过 1pt。</p>
     */
    private int resolveRunPositionHalfPoints(String rawLatex, double targetHeightPt) {
        int halfPt = -(int) Math.round(targetHeightPt - 6d);
        return Math.min(halfPt, -2);
    }

    private double resolveDisplayDepthPt(LaTeXImageRenderer.PreviewImage preview, double previewHeightPt,
                                         double shapeHeightPt) {
        double depthPt = preview.depthPt();
        if (depthPt < 0d) {
            return depthPt;
        }
        if (previewHeightPt > 0d && shapeHeightPt > 0d && Math.abs(previewHeightPt - shapeHeightPt) > 0.01d) {
            return depthPt * shapeHeightPt / previewHeightPt;
        }
        return depthPt;
    }

    private double resolvePreviewWidthPt(LaTeXImageRenderer.PreviewImage preview, PreviewBox previewBox,
                                         FormulaMetrics targetMetrics) {
        if (targetMetrics != null && targetMetrics.wmfWidthPt() > 0d) {
            return targetMetrics.wmfWidthPt();
        }
        if (preview.widthPt() > 0d && preview.widthPx() == previewBox.widthPx()) {
            return preview.widthPt();
        }
        return previewBox.widthPx() * PT_PER_PX;
    }

    private double resolvePreviewHeightPt(LaTeXImageRenderer.PreviewImage preview, PreviewBox previewBox,
                                          FormulaMetrics targetMetrics) {
        if (targetMetrics != null && targetMetrics.wmfHeightPt() > 0d) {
            return targetMetrics.wmfHeightPt();
        }
        if (preview.heightPt() > 0d && preview.heightPx() == previewBox.heightPx()) {
            return preview.heightPt();
        }
        return previewBox.heightPx() * PT_PER_PX;
    }

    /**
     * 预览框尺寸直通：渲染出的物理尺寸即写入 Word 的显示框尺寸。
     *
     * <p>测试集（word_files）里 MathType 显示框尺寸与 WMF 物理尺寸一比一，
     * 所以这里不再做按公式类型的人为缩放，只保留版面安全的最大宽高约束。</p>
     */
    private PreviewBox constrainPreviewBox(String rawLatex, int widthPx, int heightPx, double externalDisplayScale,
                                           double externalMaxWidthPt) {
        widthPx = Math.max(widthPx, 4);
        heightPx = Math.max(heightPx, 4);
        double displayScale = Math.max(externalDisplayScale, 0.25d);
        if (displayScale != 1.0d) {
            widthPx = Math.max((int) Math.round(widthPx * displayScale), 4);
            heightPx = Math.max((int) Math.round(heightPx * displayScale), 4);
        }
        double maxWidthPt = externalMaxWidthPt < Double.MAX_VALUE / 2.0d
            ? externalMaxWidthPt
            : MAX_GENERIC_FORMULA_WIDTH_PT;
        double scale = Math.min(1.0d, maxWidthPt / Math.max(widthPx * PT_PER_PX, 1.0d));
        scale = Math.min(scale, MAX_DISPLAY_HEIGHT_PT / Math.max(heightPx * PT_PER_PX, 1.0d));
        if (scale >= 0.999d) {
            return new PreviewBox(widthPx, heightPx);
        }
        return new PreviewBox(
            Math.max((int) Math.round(widthPx * scale), 4),
            Math.max((int) Math.round(heightPx * scale), 4)
        );
    }

    private record PreviewBox(int widthPx, int heightPx) {
    }
}
