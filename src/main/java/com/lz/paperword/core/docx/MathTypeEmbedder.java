package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefWriter;
import com.lz.paperword.core.ole.OlePackager;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

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
    private static final double MAX_GENERIC_ARRAY_WIDTH_PT = 315.0d;
    private static final double MAX_DERIVATION_ARRAY_WIDTH_PT = 420.0d;
    private static final double MAX_GENERIC_FORMULA_WIDTH_PT = 315.0d;
    private static final double MAX_CROSS_ARRAY_WIDTH_PT = 140.0d;
    private static final double MAX_LONG_DIVISION_WIDTH_PT = 150.0d;
    private static final double MAX_DISPLAY_HEIGHT_PT = 96.0d;
    private static final double FRACTION_DISPLAY_SCALE = 1.95d;
    /** Word 中的 w:position 使用半磅；参考文档多数分式对象约为 -22 到 -26。 */
    private static final double BASELINE_SHIFT_SCALE = 1.0d;
    private final MtefWriter mtefWriter = new MtefWriter();
    private final OlePackager olePackager = new OlePackager();
    private final LaTeXImageRenderer imageRenderer = new LaTeXImageRenderer();
    private final AtomicInteger oleCounter = new AtomicInteger(1);

    /**
     * 将 MathType 公式嵌入 Word 段落的指定 run 中。
     */
    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex) {
        embedEquation(paragraph, run, latexAst, rawLatex, 1.0d);
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex,
                              double displayScale) {
        embedEquation(paragraph, run, latexAst, rawLatex, displayScale, Double.MAX_VALUE);
    }

    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex,
                              double displayScale, double maxWidthPt) {
        try {
            byte[] mtefData = mtefWriter.write(latexAst);
            byte[] oleData = olePackager.packageOle(mtefData);

            LaTeXImageRenderer.PreviewImage preview = imageRenderer.renderForOlePreview(rawLatex);
            if (preview == null || preview.data() == null || preview.data().length == 0) {
                throw new IllegalStateException("OLE preview rendering returned no image data");
            }
            PreviewBox previewBox = constrainPreviewBox(rawLatex, preview.widthPx(), preview.heightPx(),
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

            insertOleObjectXml(paragraph, run, oleRel.getId(), imgRel.getId(), idx,
                previewBox.widthPx(), previewBox.heightPx(), rawLatex);

        } catch (Exception e) {
            log.error("Failed to embed MathType equation: {}", rawLatex, e);
            run.setText("[" + rawLatex + "]");
        }
    }

    private void insertOleObjectXml(XWPFParagraph paragraph, XWPFRun run, String oleRelId, String imgRelId,
                                     int shapeIdx, int widthPx, int heightPx, String rawLatex) {
        try {
            widthPx = Math.max(widthPx, 10);
            heightPx = Math.max(heightPx, 10);

            String shapeId = "_x0000_i" + (1024 + shapeIdx);

            double targetWidthPt = widthPx * PT_PER_PX;
            double targetHeightPt = heightPx * PT_PER_PX;
            String styleWidth = String.format("%.1fpt", targetWidthPt);
            String styleHeight = String.format("%.1fpt", targetHeightPt);
            int dxaOrig = Math.max((int) Math.round(targetWidthPt * 20), 1);
            int dyaOrig = Math.max((int) Math.round(targetHeightPt * 20), 1);
            int posHalfPt = resolveRunPositionHalfPoints(rawLatex, targetHeightPt);

            String objectId = "_" + Integer.toUnsignedString((shapeId + ":" + oleRelId).hashCode());

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

                "<v:shape id=\"" + shapeId + "\" type=\"#_x0000_t75\" " +
                "style=\"width:" + styleWidth + ";height:" + styleHeight + "\" o:ole=\"\">" +
                "<v:imagedata r:id=\"" + imgRelId + "\" o:title=\"\"/>" +
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

    /**
     * 根据公式类型和高度推算 w:position 值（半磅，负值=下移）。
     * 参考文档中的实测值：线性=-6，分数=-24，根号=-8。
     */
    private int resolveRunPositionHalfPoints(String rawLatex, double targetHeightPt) {
        String latex = rawLatex == null ? "" : rawLatex;

        if (latex.contains("\\longdiv") || latex.contains("\\enclose{longdiv}")) {
            // 长除法对象整体更高，保留更大的下移量，但跟着新字号同比缩小。
            return scaleHalfPoints(-112);
        }
        if (latex.contains("\\frac") || latex.contains("\\dfrac") || latex.contains("\\cfrac")) {
            return targetHeightPt >= 24d ? -24 : scaleHalfPoints(-24);
        }
        if (latex.contains("\\sqrt")) {
            return scaleHalfPoints(-8);
        }
        if (targetHeightPt >= 28d) {
            return scaleHalfPoints(-24);
        }
        if (targetHeightPt >= 24d) {
            return scaleHalfPoints(-8);
        }
        return scaleHalfPoints(-6);
    }

    private int scaleHalfPoints(int originalHalfPt) {
        return (int) Math.round(originalHalfPt * BASELINE_SHIFT_SCALE);
    }

    /**
     * 宽矩阵和十字交叉的预览图如果按原始尺寸写入，Word 中的 OLE 占位会被拉得过宽。
     * 这里仅收敛预览框尺寸，不修改内部 MTEF，可编辑性仍由原始公式决定。
     */
    private PreviewBox constrainPreviewBox(String rawLatex, int widthPx, int heightPx, double externalDisplayScale,
                                           double externalMaxWidthPt) {
        widthPx = Math.max(widthPx, 10);
        heightPx = Math.max(heightPx, 10);
        double displayScale = resolvePreviewDisplayScale(rawLatex) * Math.max(externalDisplayScale, 0.25d);
        widthPx = Math.max((int) Math.round(widthPx * displayScale), 10);
        heightPx = Math.max((int) Math.round(heightPx * displayScale), 10);
        PreviewBox normalizedBox = normalizeArrayPreviewBox(rawLatex, widthPx, heightPx);
        widthPx = normalizedBox.widthPx();
        heightPx = normalizedBox.heightPx();
        double maxWidthPt = externalMaxWidthPt < Double.MAX_VALUE / 2.0d
            ? externalMaxWidthPt
            : resolveMaxPreviewWidthPt(rawLatex);
        double scale = Math.min(1.0d, maxWidthPt / Math.max(widthPx * PT_PER_PX, 1.0d));
        scale = Math.min(scale, MAX_DISPLAY_HEIGHT_PT / Math.max(heightPx * PT_PER_PX, 1.0d));
        if (scale >= 0.999d) {
            return normalizeCompactPreviewHeight(rawLatex, new PreviewBox(widthPx, heightPx), externalMaxWidthPt);
        }
        PreviewBox constrainedBox = new PreviewBox(
            Math.max((int) Math.round(widthPx * scale), 10),
            Math.max((int) Math.round(heightPx * scale), 10)
        );
        return normalizeCompactPreviewHeight(rawLatex, constrainedBox, externalMaxWidthPt);
    }

    private PreviewBox normalizeArrayPreviewBox(String rawLatex, int widthPx, int heightPx) {
        String latex = rawLatex == null ? "" : rawLatex;
        if (!latex.contains("\\begin{array}")) {
            return new PreviewBox(widthPx, heightPx);
        }

        double widthScale = 0.94d;
        double minHeightPt = 20.3d;
        if (countOccurrences(latex, "\\begin{array}") >= 2) {
            widthScale = 0.90d;
            minHeightPt = 52.5d;
        } else if (latex.contains("\\\\") || latex.contains("\\cr")) {
            widthScale = 0.98d;
            minHeightPt = 36.8d;
        } else if (latex.length() >= 150 || latex.contains("\\cdots")) {
            widthScale = 0.94d;
            minHeightPt = 25.5d;
        }

        int normalizedWidthPx = Math.max((int) Math.round(widthPx * widthScale), 10);
        int minHeightPx = Math.max((int) Math.round(minHeightPt / PT_PER_PX), 10);
        return new PreviewBox(normalizedWidthPx, Math.max(heightPx, minHeightPx));
    }

    private PreviewBox normalizeCompactPreviewHeight(String rawLatex, PreviewBox box, double externalMaxWidthPt) {
        if (externalMaxWidthPt >= Double.MAX_VALUE / 2.0d) {
            return box;
        }
        String latex = rawLatex == null ? "" : rawLatex.trim();
        if (latex.isEmpty()
            || latex.contains("\\begin{array}")
            || latex.contains("\\begin{aligned}")
            || latex.contains("\\begin{matrix}")
            || latex.contains("\\longdiv")
            || latex.contains("\\enclose{longdiv}")) {
            return box;
        }

        double minHeightPt = resolveCompactMinimumHeightPt(latex);
        if (minHeightPt <= 0d) {
            return box;
        }
        int minHeightPx = Math.max((int) Math.round(minHeightPt / PT_PER_PX), 10);
        if (box.heightPx() >= minHeightPx) {
            return box;
        }
        return new PreviewBox(box.widthPx(), minHeightPx);
    }

    private double resolveCompactMinimumHeightPt(String latex) {
        if (latex.contains("\\frac") || latex.contains("\\dfrac") || latex.contains("\\cfrac")) {
            return 27.75d;
        }
        if (latex.matches(".*[=+\\-*/×÷].*") || latex.matches(".*\\d.*")) {
            return latex.length() <= 16 ? 24.0d : 21.0d;
        }
        if (latex.contains("<") || latex.contains(">")) {
            return 16.0d;
        }
        return 0d;
    }

    private int countOccurrences(String value, String token) {
        int count = 0;
        int offset = 0;
        while (value != null && (offset = value.indexOf(token, offset)) >= 0) {
            count++;
            offset += token.length();
        }
        return count;
    }

    private double resolveMaxPreviewWidthPt(String rawLatex) {
        String latex = rawLatex == null ? "" : rawLatex;
        if (isCrossMultiplicationArray(latex)) {
            return MAX_CROSS_ARRAY_WIDTH_PT;
        }
        if (latex.contains("\\longdiv") || latex.contains("\\enclose{longdiv}")) {
            return MAX_LONG_DIVISION_WIDTH_PT;
        }
        if (latex.contains("\\begin{array}") || latex.contains("\\begin{aligned}") || latex.contains("\\begin{matrix}")) {
            return isArithmeticArray(latex) ? MAX_GENERIC_ARRAY_WIDTH_PT : MAX_DERIVATION_ARRAY_WIDTH_PT;
        }
        return MAX_GENERIC_FORMULA_WIDTH_PT;
    }

    private double resolvePreviewDisplayScale(String rawLatex) {
        String latex = rawLatex == null ? "" : rawLatex;
        if (latex.contains("\\begin{array}")
            || latex.contains("\\begin{aligned}")
            || latex.contains("\\begin{matrix}")) {
            return 1.18d;
        }
        if (latex.contains("\\longdiv") || latex.contains("\\enclose{longdiv}")) {
            return 1.0d;
        }
        if (latex.contains("\\frac") || latex.contains("\\dfrac") || latex.contains("\\cfrac")) {
            return FRACTION_DISPLAY_SCALE;
        }
        return 1.0d;
    }

    /**
     * 十字交叉矩阵横向信息密度最高，单独使用更严格的预览宽度上限。
     */
    private boolean isCrossMultiplicationArray(String latex) {
        return latex.contains("\\begin{array}{ccccc}")
            && latex.contains("\\nearrow")
            && latex.contains("\\searrow");
    }

    private boolean isArithmeticArray(String latex) {
        java.util.regex.Matcher matcher = java.util.regex.Pattern
            .compile("\\\\begin\\{array\\}\\{([^{}]+)\\}")
            .matcher(latex == null ? "" : latex);
        if (!matcher.find()) {
            return false;
        }
        String columnSpec = matcher.group(1).replace("|", "");
        return !columnSpec.isBlank() && columnSpec.chars().allMatch(ch -> ch == 'r');
    }

    private record PreviewBox(int widthPx, int heightPx) {
    }
}
