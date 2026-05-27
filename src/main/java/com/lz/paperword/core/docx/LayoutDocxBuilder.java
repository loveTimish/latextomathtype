package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.latex.LaTeXParser.ContentSegment;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 将 OCR/版面归一化后的布局模型直接写成保持分页与块结构的 DOCX。
 */
public class LayoutDocxBuilder {

    private static final Logger log = LoggerFactory.getLogger(LayoutDocxBuilder.class);
    private static final int DEFAULT_PAGE_WIDTH_TWIPS = 11906;
    private static final int DEFAULT_PAGE_HEIGHT_TWIPS = 16838;
    private static final Pattern SERIAL_PREFIX_PATTERN = Pattern.compile("^(\\s*\\(?\\d+[.)、]?)(\\s*)(.+)$");

    private final LaTeXParser latexParser = new LaTeXParser();
    private final MathTypeEmbedder mathEmbedder = new MathTypeEmbedder();
    private final boolean embedMathTypeOle;

    public LayoutDocxBuilder(boolean embedMathTypeOle) {
        this.embedMathTypeOle = embedMathTypeOle;
    }

    public byte[] build(LayoutDocumentRequest request) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
            setPageMargins(doc, request.getDocument());
            List<LayoutDocumentRequest.Page> pages = request.getPages() == null
                ? List.of()
                : request.getPages().stream()
                .sorted(Comparator.comparing(page -> safeInt(page.getPageNumber(), 0)))
                .toList();

            for (int i = 0; i < pages.size(); i++) {
                writePage(doc, pages.get(i), request.getDocument());
                if (i < pages.size() - 1) {
                    XWPFParagraph pageBreak = doc.createParagraph();
                    pageBreak.setPageBreak(true);
                }
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            doc.write(baos);
            return baos.toByteArray();
        }
    }

    private void setPageMargins(XWPFDocument doc, LayoutDocumentRequest.DocumentInfo info) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setTop(BigInteger.valueOf(safeInt(info.getMarginTopTwips(), 1440)));
        pageMar.setBottom(BigInteger.valueOf(safeInt(info.getMarginBottomTwips(), 1440)));
        pageMar.setLeft(BigInteger.valueOf(safeInt(info.getMarginLeftTwips(), 1440)));
        pageMar.setRight(BigInteger.valueOf(safeInt(info.getMarginRightTwips(), 1440)));

        CTPageSz pageSz = sectPr.addNewPgSz();
        pageSz.setW(BigInteger.valueOf(DEFAULT_PAGE_WIDTH_TWIPS));
        pageSz.setH(BigInteger.valueOf(DEFAULT_PAGE_HEIGHT_TWIPS));
    }

    private void writePage(XWPFDocument doc, LayoutDocumentRequest.Page page, LayoutDocumentRequest.DocumentInfo info) {
        List<LayoutDocumentRequest.Block> blocks = page.getBlocks();
        if ((blocks == null || blocks.isEmpty()) && page.getBackgroundImagePath() != null) {
            LayoutDocumentRequest.Block imageBlock = new LayoutDocumentRequest.Block();
            imageBlock.setType(LayoutDocumentRequest.BlockType.IMAGE);
            imageBlock.setImagePath(page.getBackgroundImagePath());
            imageBlock.setImageContentType("image/png");
            blocks = List.of(imageBlock);
        }
        if (blocks == null) {
            return;
        }

        blocks.stream()
            .sorted(Comparator.comparing((LayoutDocumentRequest.Block block) -> yOf(block.getBbox()))
                .thenComparing(block -> xOf(block.getBbox())))
            .forEach(block -> writeBlock(doc, page, info, block));
    }

    private void writeBlock(XWPFDocument doc, LayoutDocumentRequest.Page page,
                            LayoutDocumentRequest.DocumentInfo info,
                            LayoutDocumentRequest.Block block) {
        LayoutDocumentRequest.BlockType type = block.getType() == null
            ? LayoutDocumentRequest.BlockType.PARAGRAPH
            : block.getType();
        switch (type) {
            case IMAGE -> writeImageBlock(doc, page, info, block);
            case TABLE -> writeTableBlock(doc, block);
            case FORMULA -> writeFormulaBlock(doc, block);
            case TITLE, HEADING, PARAGRAPH -> writeTextBlock(doc, block, type);
        }
    }

    private void writeTextBlock(XWPFDocument doc, LayoutDocumentRequest.Block block,
                                LayoutDocumentRequest.BlockType type) {
        XWPFParagraph para = doc.createParagraph();
        applyParagraphStyle(para, block.getStyle(), type);
        writeContentWithMath(para, "", block.getText(), block.getStyle());
    }

    private void writeFormulaBlock(XWPFDocument doc, LayoutDocumentRequest.Block block) {
        XWPFParagraph para = doc.createParagraph();
        LayoutDocumentRequest.Style style = block.getStyle() == null ? new LayoutDocumentRequest.Style() : block.getStyle();
        applyParagraphStyle(para, style, LayoutDocumentRequest.BlockType.FORMULA);
        String content = block.getText();
        if ((content == null || content.isBlank()) && block.getLatex() != null) {
            content = "$" + block.getLatex() + "$";
        }
        writeContentWithMath(para, "", content, style);
    }

    private void writeTableBlock(XWPFDocument doc, LayoutDocumentRequest.Block block) {
        List<LayoutDocumentRequest.TableRow> rows = block.getRows();
        if (rows == null || rows.isEmpty()) {
            return;
        }
        int columnCount = rows.stream()
            .mapToInt(row -> row.getCells() == null ? 0 : row.getCells().size())
            .max()
            .orElse(0);
        if (columnCount <= 0) {
            return;
        }

        XWPFTable table = doc.createTable(rows.size(), columnCount);
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            List<LayoutDocumentRequest.TableCell> cells = rows.get(rowIndex).getCells();
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                XWPFTableCell cell = table.getRow(rowIndex).getCell(columnIndex);
                cell.removeParagraph(0);
                XWPFParagraph para = cell.addParagraph();
                LayoutDocumentRequest.Style style = block.getStyle() == null ? new LayoutDocumentRequest.Style() : block.getStyle();
                if (cells != null && columnIndex < cells.size()) {
                    writeContentWithMath(para, "", cells.get(columnIndex).getText(), style);
                }
            }
        }
    }

    private void writeImageBlock(XWPFDocument doc, LayoutDocumentRequest.Page page,
                                 LayoutDocumentRequest.DocumentInfo info,
                                 LayoutDocumentRequest.Block block) {
        String imagePath = block.getImagePath() != null ? block.getImagePath() : page.getBackgroundImagePath();
        if (imagePath == null || imagePath.isBlank()) {
            return;
        }
        Path path = Path.of(imagePath);
        if (!Files.exists(path)) {
            log.warn("Image block path does not exist: {}", imagePath);
            return;
        }

        XWPFParagraph para = doc.createParagraph();
        LayoutDocumentRequest.Style style = block.getStyle() == null ? new LayoutDocumentRequest.Style() : block.getStyle();
        applyParagraphStyle(para, style, LayoutDocumentRequest.BlockType.IMAGE);
        try (InputStream inputStream = Files.newInputStream(path)) {
            BufferedImage image = ImageIO.read(path.toFile());
            int widthPx = image != null ? image.getWidth() : safeInt(block.getImageWidthPx(), 0);
            int heightPx = image != null ? image.getHeight() : safeInt(block.getImageHeightPx(), 0);
            if (widthPx <= 0 || heightPx <= 0) {
                return;
            }

            int usableWidthTwips = DEFAULT_PAGE_WIDTH_TWIPS
                - safeInt(info.getMarginLeftTwips(), 1440)
                - safeInt(info.getMarginRightTwips(), 1440);
            int targetWidthTwips = deriveTargetWidthTwips(page, block, usableWidthTwips);
            int targetHeightTwips = (int) Math.round((double) heightPx * targetWidthTwips / Math.max(widthPx, 1));

            XWPFRun run = para.createRun();
            run.addPicture(inputStream, resolvePictureType(imagePath), path.getFileName().toString(),
                Units.toEMU(targetWidthTwips / 20.0), Units.toEMU(targetHeightTwips / 20.0));
        } catch (Exception e) {
            log.error("Failed to insert image {}", imagePath, e);
        }
    }

    private int deriveTargetWidthTwips(LayoutDocumentRequest.Page page,
                                       LayoutDocumentRequest.Block block,
                                       int usableWidthTwips) {
        LayoutDocumentRequest.BoundingBox bbox = block.getBbox();
        if (bbox == null || bbox.getWidth() == null || bbox.getWidth() <= 0) {
            return usableWidthTwips;
        }
        int pageWidth = safeInt(page.getWidth(), 1000);
        return Math.max(1440, usableWidthTwips * bbox.getWidth() / Math.max(pageWidth, 1));
    }

    private void applyParagraphStyle(XWPFParagraph para, LayoutDocumentRequest.Style style,
                                     LayoutDocumentRequest.BlockType type) {
        LayoutDocumentRequest.Style effective = style == null ? new LayoutDocumentRequest.Style() : style;
        para.setAlignment(mapAlignment(effective.getAlignment()));
        if (effective.getIndentLeftTwips() != null) {
            para.setIndentationLeft(effective.getIndentLeftTwips());
        }
        para.setSpacingBefore(safeInt(effective.getSpacingBeforeTwips(), defaultSpacingBefore(type)));
        para.setSpacingAfter(safeInt(effective.getSpacingAfterTwips(), defaultSpacingAfter(type)));
    }

    private void writeContentWithMath(XWPFParagraph para, String prefix, String content,
                                      LayoutDocumentRequest.Style style) {
        LayoutDocumentRequest.Style effective = style == null ? new LayoutDocumentRequest.Style() : style;
        if (prefix != null && !prefix.isEmpty()) {
            XWPFRun prefixRun = para.createRun();
            applyTextRunStyle(prefixRun, effective);
            writeTextPreserveLineBreaks(prefixRun, prefix);
        }
        if (content == null || content.isBlank()) {
            return;
        }

        String normalized = normalizeLineBreaks(content);
        if ((prefix == null || prefix.isEmpty()) && looksLikeDisplayLatexBlock(normalized)) {
            writeMathExpression(para, unwrapDisplayMathBlock(normalized.trim()), effective);
            return;
        }
        List<String> lines = splitContentLinesPreservingDisplayMath(normalized);
        for (int lineIndex = 0; lineIndex < lines.size(); lineIndex++) {
            if (lineIndex > 0) {
                para.createRun().addBreak();
            }
            if (lines.get(lineIndex).isEmpty()) {
                continue;
            }
            writeLineWithMath(para, lines.get(lineIndex), effective);
        }
    }

    /**
     * 避免把 $$...$$ 行间公式按普通换行拆散，否则多行矩阵会被拆成残缺文本。
     */
    private List<String> splitContentLinesPreservingDisplayMath(String content) {
        List<String> lines = new ArrayList<>();
        if (content == null || content.isEmpty()) {
            return lines;
        }
        StringBuilder current = new StringBuilder();
        boolean insideDisplayMath = false;
        for (int index = 0; index < content.length(); index++) {
            if (content.startsWith("$$", index)) {
                insideDisplayMath = !insideDisplayMath;
                current.append("$$");
                index++;
                continue;
            }
            char currentChar = content.charAt(index);
            if (currentChar == '\n' && !insideDisplayMath) {
                lines.add(current.toString());
                current.setLength(0);
                continue;
            }
            current.append(currentChar);
        }
        lines.add(current.toString());
        return lines;
    }

    private void writeLineWithMath(XWPFParagraph para, String line, LayoutDocumentRequest.Style style) {
        String trimmed = line.trim();
        if (trimmed.isEmpty()) {
            return;
        }
        if (containsWholeDisplayMathBlock(trimmed)) {
            writeLineWithDisplayMathBlock(para, line, style);
            return;
        }
        writeGenericSegments(para, line, style);
    }

    /**
     * 某些题目会把“【解答】”和整块 $$...$$ 放在同一行；这里优先抽出定界公式，
     * 仅把真正被 $$ 包裹的内容送入 MathType，前后缀仍按普通文本写入。
     */
    private void writeLineWithDisplayMathBlock(XWPFParagraph para, String line, LayoutDocumentRequest.Style style) {
        int beginIndex = findDisplayMathStart(line);
        if (beginIndex < 0) {
            writeGenericSegments(para, line, style);
            return;
        }
        int endIndex = findDisplayMathEnd(line, beginIndex);
        if (endIndex < 0) {
            writeGenericSegments(para, line, style);
            return;
        }
        String prefix = line.substring(0, beginIndex);
        String math = line.substring(beginIndex + 2, endIndex);
        String suffix = line.substring(endIndex);
        if (!prefix.isBlank()) {
            writeTextSegment(para, stripDisplayDelimiterNoise(prefix), style);
        }
        writeMathExpression(para, math.trim(), style);
        if (!suffix.isBlank()) {
            writeGenericSegments(para, stripDisplayDelimiterNoise(suffix), style);
        }
    }

    private boolean containsWholeDisplayMathBlock(String text) {
        int beginIndex = findDisplayMathStart(text);
        return beginIndex >= 0 && findDisplayMathEnd(text, beginIndex) > beginIndex;
    }

    private int findDisplayMathStart(String text) {
        return text.indexOf("$$");
    }

    private int findDisplayMathEnd(String text, int beginIndex) {
        if (beginIndex < 0) {
            return -1;
        }
        return text.indexOf("$$", beginIndex + 2);
    }

    private void writeEquationSequence(XWPFParagraph para, String text, LayoutDocumentRequest.Style style) {
        if (text == null || text.isBlank()) {
            return;
        }
        String normalized = text.trim();
        if (normalized.startsWith("=")) {
            writeTextSegment(para, "= ", style);
            normalized = normalized.substring(1).trim();
        }
        if (normalized.isEmpty()) {
            return;
        }
        String[] parts = normalized.split("\\s+=\\s+");
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) {
                writeTextSegment(para, " = ", style);
            }
            writeExpressionOrText(para, parts[i].trim(), style);
        }
    }

    private void writeExpressionOrText(XWPFParagraph para, String text, LayoutDocumentRequest.Style style) {
        if (text == null || text.isBlank()) {
            return;
        }
        writeGenericSegments(para, text, style);
    }

    private void writeGenericSegments(XWPFParagraph para, String text, LayoutDocumentRequest.Style style) {
        List<ContentSegment> segments = looksLikeHtml(text)
            ? latexParser.parseHtml(text)
            : latexParser.parseText(text);
        for (ContentSegment segment : segments) {
            if (segment.isMath() && segment.ast() != null) {
                writeMathAst(para, segment.rawText(), segment.ast(), style);
            } else {
                writeTextSegment(para, segment.rawText(), style);
            }
        }
    }

    private void writeMathExpression(XWPFParagraph para, String latex, LayoutDocumentRequest.Style style) {
        try {
            writeMathAst(para, latex, latexParser.parseLaTeX(latex), style);
        } catch (Exception e) {
            log.error("Failed to parse formula, falling back to text: {}", latex, e);
            writeTextSegment(para, latex, style);
        }
    }

    private void writeMathAst(XWPFParagraph para, String rawLatex, LaTeXNode ast,
                              LayoutDocumentRequest.Style style) {
        XWPFRun mathRun = para.createRun();
        applyMathRunStyle(mathRun, style);
        if (embedMathTypeOle) {
            try {
                mathEmbedder.embedEquation(para, mathRun, ast, rawLatex);
            } catch (Exception e) {
                log.error("Failed to embed formula, falling back to text: {}", rawLatex, e);
                writeTextPreserveLineBreaks(mathRun, "$" + rawLatex + "$");
            }
        } else {
            writeTextPreserveLineBreaks(mathRun, "$" + rawLatex + "$");
        }
    }

    private void writeTextSegment(XWPFParagraph para, String text, LayoutDocumentRequest.Style style) {
        if (text == null || text.isEmpty()) {
            return;
        }
        XWPFRun textRun = para.createRun();
        applyTextRunStyle(textRun, style);
        writeTextPreserveLineBreaks(textRun, text);
    }

    private void applyTextRunStyle(XWPFRun run, LayoutDocumentRequest.Style style) {
        run.setBold(style.isBold());
        run.setItalic(false);
        if (style.getFontFamily() != null && !style.getFontFamily().isBlank()) {
            run.setFontFamily(style.getFontFamily());
        }
        if (style.getFontSizePt() != null && style.getFontSizePt() > 0) {
            run.setFontSize(style.getFontSizePt());
        }
    }

    private void applyMathRunStyle(XWPFRun run, LayoutDocumentRequest.Style style) {
        run.setBold(style.isBold());
        run.setItalic(false);
    }

    private void writeTextPreserveLineBreaks(XWPFRun run, String text) {
        if (text == null) {
            return;
        }
        String normalized = normalizeLineBreaks(text);
        String[] parts = normalized.split("\n", -1);
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            if (!parts[i].isEmpty()) {
                run.setText(parts[i]);
            }
        }
    }

    private String normalizeLineBreaks(String text) {
        return text.replace("\r\n", "\n").replace('\r', '\n');
    }

    private boolean looksLikeHtml(String text) {
        return text.contains("<") && text.contains(">");
    }

    /**
     * 只有完整的 $$...$$ 才视为整块行间公式，其余裸露 LaTeX 一律按普通文本保留。
     */
    private boolean looksLikeDisplayLatexBlock(String text) {
        if (text == null) {
            return false;
        }
        String normalized = text.trim();
        if (normalized.isEmpty()) {
            return false;
        }
        return normalized.startsWith("$$")
            && normalized.endsWith("$$")
            && normalized.length() > 4
            && findDisplayMathEnd(normalized, 0) == normalized.length() - 2;
    }

    /**
     * 去掉 display math 两侧的 $$ 定界符，只把内部 LaTeX 交给 MathType。
     */
    private String unwrapDisplayMathBlock(String text) {
        return text.substring(2, text.length() - 2).trim();
    }

    /**
     * 当前后缀因为上游混排行残留裸 $$ 时，兜底去掉这些无意义文本，避免直接写进 Word。
     */
    private String stripDisplayDelimiterNoise(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        return text.replace("$$", "").trim();
    }

    private ParagraphAlignment mapAlignment(LayoutDocumentRequest.Alignment alignment) {
        if (alignment == null) {
            return ParagraphAlignment.LEFT;
        }
        return switch (alignment) {
            case CENTER -> ParagraphAlignment.CENTER;
            case RIGHT -> ParagraphAlignment.RIGHT;
            case BOTH -> ParagraphAlignment.BOTH;
            case LEFT -> ParagraphAlignment.LEFT;
        };
    }

    private int defaultSpacingBefore(LayoutDocumentRequest.BlockType type) {
        return switch (type) {
            case TITLE -> 120;
            case HEADING -> 180;
            case FORMULA, IMAGE, TABLE -> 120;
            default -> 80;
        };
    }

    private int defaultSpacingAfter(LayoutDocumentRequest.BlockType type) {
        return switch (type) {
            case TITLE -> 200;
            case HEADING -> 120;
            case FORMULA, IMAGE, TABLE -> 120;
            default -> 60;
        };
    }

    private int resolvePictureType(String fileName) {
        String lower = fileName.toLowerCase();
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        }
        if (lower.endsWith(".gif")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        }
        if (lower.endsWith(".bmp")) {
            return XWPFDocument.PICTURE_TYPE_BMP;
        }
        return XWPFDocument.PICTURE_TYPE_PNG;
    }

    private int xOf(LayoutDocumentRequest.BoundingBox bbox) {
        return bbox == null ? 0 : safeInt(bbox.getX(), 0);
    }

    private int yOf(LayoutDocumentRequest.BoundingBox bbox) {
        return bbox == null ? 0 : safeInt(bbox.getY(), 0);
    }

    private int safeInt(Integer value, int fallback) {
        return value != null ? value : fallback;
    }
}
