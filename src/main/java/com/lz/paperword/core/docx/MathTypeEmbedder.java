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
    /** 公式层默认字号已从 30pt 收敛到 20pt，基线偏移也按同样比例缩小。 */
    private static final double BASELINE_SHIFT_SCALE = 20.0d / 30.0d;
    private final MtefWriter mtefWriter = new MtefWriter();
    private final OlePackager olePackager = new OlePackager();
    private final LaTeXImageRenderer imageRenderer = new LaTeXImageRenderer();
    private final AtomicInteger oleCounter = new AtomicInteger(1);

    /**
     * 将 MathType 公式嵌入 Word 段落的指定 run 中。
     */
    public void embedEquation(XWPFParagraph paragraph, XWPFRun run, LaTeXNode latexAst, String rawLatex) {
        try {
            byte[] mtefData = mtefWriter.write(latexAst);
            byte[] oleData = olePackager.packageOle(mtefData);

            LaTeXImageRenderer.PreviewImage preview = imageRenderer.renderForOlePreview(rawLatex);
            if (preview == null || preview.data() == null || preview.data().length == 0) {
                throw new IllegalStateException("OLE preview rendering returned no image data");
            }

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
                preview.widthPx(), preview.heightPx(), rawLatex);

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
            return scaleHalfPoints(-24);
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
}
