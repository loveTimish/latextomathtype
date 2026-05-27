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
 * @see OlePackager OLE2 复合文档打包
 * @see LaTeXImageRenderer LaTeX 公式预览图渲染
 * @see MtefWriter LaTeX AST → MTEF 二进制转换
 */
public class MathTypeEmbedder {

 private static final Logger log = LoggerFactory.getLogger(MathTypeEmbedder.class);

 private static final double PT_PER_PX = 0.75d;
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
 log.warn("Could not render preview for: {}", rawLatex);
 run.setText("[" + rawLatex + "]");
 return;
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
 String runXml = " " +
 " " +
 " " +

 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +
 " " +

 " " +
 " " +
 " " +

 " " +
 " ";

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

 if (latex.contains("\\frac") || latex.contains("\\dfrac") || latex.contains("\\cfrac")) {
 return -24;
 }
 if (latex.contains("\\sqrt")) {
 return -8;
 }
 if (targetHeightPt >= 28d) {
 return -24;
 }
 if (targetHeightPt >= 24d) {
 return -8;
 }
 return -6;
 }
}
