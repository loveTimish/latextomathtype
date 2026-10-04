package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.latex.LaTeXParser.ContentSegment;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.jsoup.Jsoup;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.xmlbeans.XmlBoolean;

/**
 * Word 文档（.docx）构建器：将试卷导出请求转换为完整的 Word 文档。
 *
 * <p>本类是试卷导出功能的顶层编排器（Orchestrator），负责将结构化的试卷数据
 * （{@link PaperExportRequest}）转换为格式规范的 .docx 文件字节流。</p>
 *
 * <h3>在整体系统中的角色：</h3>
 * <pre>
 * 前端/业务层
 *     ↓ PaperExportRequest（试卷标题、大题、小题、选项、答案、解析）
 * DocxBuilder.build()
 *     ├── 设置页面布局（A4 纸、页边距）
 *     ├── 写入试卷标题（居中、加粗、大号字体）
 *     ├── 写入试卷信息行（总分、建议时长）
 *     └── 遍历各大题（SectionDTO）
 *         ├── 写入大题标题
 *         └── 遍历各小题（QuestionDTO）
 *             ├── 写入题目内容（含 LaTeX 公式 → MathType OLE）
 *             ├── 写入选项（选择题：两列布局）
 *             ├── 写入答题空间（简答/计算题）
 *             ├── 写入正确答案
 *             └── 写入解析说明
 * </pre>
 *
 * <h3>LaTeX 公式处理流水线：</h3>
 * <p>题目内容（HTML 格式）中可能嵌入 LaTeX 数学公式（如 $\frac{x}{2}$），处理流程为：</p>
 * <ol>
 *   <li>HTML 内容 → {@link LaTeXParser#parseHtml(String)}：提取纯文本 + LaTeX 公式段</li>
 *   <li>LaTeX 公式段 → Token 化 → AST 构建（由 LaTeXParser 内部完成）</li>
 *   <li>AST → {@link MathTypeEmbedder#embedEquation}：将公式嵌入为 MathType OLE 对象</li>
 * </ol>
 *
 * <h3>文档格式规范：</h3>
 * <ul>
 *   <li>页面：A4 纸张（210mm × 297mm），四周 1 英寸页边距</li>
 *   <li>标题：宋体 18pt 加粗居中</li>
 *   <li>正文：宋体 11pt</li>
 *   <li>数学公式：Cambria Math 字体（或 MathType OLE 嵌入）</li>
 *   <li>答案和解析：蓝色标签（#0000CC）</li>
 * </ul>
 */
// Instances contain per-document layout and MTEF state. Do not share one
// instance between concurrent build() calls; use one builder per export.
public class DocxBuilder {

    private static final Logger log = LoggerFactory.getLogger(DocxBuilder.class);
    private static final String STYLE_NORMAL = "Normal";
    private static final String STYLE_TITLE = "Title";
    private static final String STYLE_HEADING_1 = "Heading1";
    private static final String STYLE_HEADING_2 = "Heading2";
    private static final String STYLE_BODY_TEXT = "BodyText";
    private static final String ANSWER_LINE_MARKER = "[[ANSWER_LINE]]";
    private static final double COMPACT_BODY_FONT_SIZE_PT = 10.5d;
    private static final double COMPACT_ANSWER_FONT_SIZE_PT = 10.5d;
    private static final double COMPACT_TITLE_FONT_SIZE_PT = 11.0d;
    private static final String COMPACT_BODY_FONT = "宋体";
    private static final String COMPACT_ANSWER_FONT = "楷体_GB2312";
    private static final Pattern EQUALS_ANSWER_BLANK_PATTERN = Pattern.compile(
        "([=＝])\\s*(?:_{3,}|([.．])(?![.．\\d]))([。．.])?");
    private static final Pattern STANDALONE_ANSWER_BLANK_PATTERN = Pattern.compile(
        "(?:(?<=[:：])\\s*|(?<=\\s))_{3,}(?=\\s|[。．.]|$)");
    private static final String FONT_TABLE_CONTENT_TYPE =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    private static final String FONT_TABLE_RELATION =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";

    /** LaTeX 解析器：负责从 HTML 中提取公式并解析为 AST */
    private final LaTeXParser latexParser = new LaTeXParser();

    /** MathType 公式嵌入器：负责将 AST 转换为 OLE 对象并嵌入 Word 文档 */
    private final MathTypeEmbedder mathEmbedder = new MathTypeEmbedder();

    /**
     * 是否启用 MathType OLE 嵌入模式。
     * <ul>
     *   <li>true：将 LaTeX 公式作为 MathType OLE 对象嵌入（可在 Word 中双击编辑）</li>
     *   <li>false：草稿模式，保留 $...$ 标记，供后续 Windows 端 MathType 批处理器转换</li>
     * </ul>
     */
    private final boolean embedMathTypeOle;
    private final ImageAssetLoader imageAssets;
    private boolean compactLayoutContext;
    private boolean hideQuestionTypeMetadataContext;
    private boolean printLayoutContext;
    private boolean examTypographyContext;
    private boolean metadataContext;
    private com.lz.paperword.core.render.LaTeXImageRenderer.PreviewBackground previewBackgroundContext;
    private final java.util.Set<XWPFParagraph> formulaParagraphs = java.util.Collections.newSetFromMap(new java.util.IdentityHashMap<>());
    private final java.util.Set<XWPFParagraph> metadataParagraphs = java.util.Collections.newSetFromMap(new java.util.IdentityHashMap<>());
    private final java.util.Set<XWPFParagraph> workingParagraphs = java.util.Collections.newSetFromMap(new java.util.IdentityHashMap<>());

    /**
     * 默认构造函数：启用 MathType OLE 嵌入。
     */
    public DocxBuilder() {
        this(true);
    }

    /**
     * 构造函数：可选择是否启用 MathType OLE 嵌入。
     *
     * @param embedMathTypeOle true=嵌入 MathType OLE 对象，false=保留 LaTeX 文本标记
     */
    public DocxBuilder(boolean embedMathTypeOle) {
        this(embedMathTypeOle, ImageAssetLoader.fromSystemProperties());
    }

    public DocxBuilder(boolean embedMathTypeOle, ImageAssetLoader imageAssets) {
        this.embedMathTypeOle = embedMathTypeOle;
        this.imageAssets = java.util.Objects.requireNonNull(imageAssets, "imageAssets");
    }

    private static void validateMode(String field, String value, String optIn) {
        if (value != null && !"legacy".equals(value) && !optIn.equals(value)) {
            throw new com.lz.paperword.model.ExportRequestValidationException(field + " must be legacy or " + optIn);
        }
    }

    private void applyExamTypography(XWPFDocument doc) {
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            String style = paragraph.getStyle();
            if (STYLE_TITLE.equals(style) || STYLE_HEADING_1.equals(style) || STYLE_HEADING_2.equals(style)) {
                for (XWPFRun run : paragraph.getRuns()) ExamTypography.styleRun(run,
                    STYLE_TITLE.equals(style) ? 18.0 : STYLE_HEADING_2.equals(style) ? 13.0 : 12.0);
                continue;
            }
            boolean metadata = metadataParagraphs.contains(paragraph);
            for (XWPFRun run : paragraph.getRuns()) {
                if (run.getCTR().sizeOfObjectArray() == 0) ExamTypography.styleRun(run,
                    metadata ? ExamTypography.METADATA_SIZE_PT : ExamTypography.BODY_SIZE_PT);
            }
            if (workingParagraphs.contains(paragraph)) ExamTypography.workingLine(paragraph);
            else ExamTypography.bodyParagraph(paragraph, metadata,
                formulaParagraphs.contains(paragraph) || paragraph.getRuns().stream().anyMatch(run -> !run.getEmbeddedPictures().isEmpty()));
        }
    }

    private void applyQuestionKeeps(List<XWPFParagraph> paragraphs, int first, int workspaceStart,
                                    int solutionStart, int end, boolean hasAnswer) {
        // Bound the opening group by its estimated physical size. Long questions and
        // long answer areas may flow rather than moving an entire question to a new page.
        int target = solutionStart >= 0 ? solutionStart : Math.min(end - 1, workspaceStart + 1);
        int estimatedLines = 0;
        for (int i = first; i < target && i < end - 1; i++) {
            XWPFParagraph paragraph = paragraphs.get(i);
            estimatedLines += Math.max(1, (paragraph.getText().length() + 55) / 56);
            if (estimatedLines > 6 || i - first >= 5) break;
            keepWithNext(paragraph);
        }
        for (int i = first; i < end - 1; i++) {
            String text = paragraphs.get(i).getText().trim();
            if (text.length() <= 60 && (text.endsWith(":") || text.endsWith("："))) keepWithNext(paragraphs.get(i));
        }
        if (solutionStart >= 0 && hasAnswer && end - 2 >= solutionStart) {
            keepWithNext(paragraphs.get(end - 2));
            // A short concluding step and the answer should not form a two-line last page.
            if (end - 3 >= solutionStart && paragraphs.get(end - 2).getText().length() <= 60
                    && paragraphs.get(end - 3).getText().length() <= 90) keepWithNext(paragraphs.get(end - 3));
        }
    }

    private void initializeStyles(XWPFDocument doc) {
        XWPFStyles styles = doc.createStyles();
        double bodyFontSize = examTypographyContext ? ExamTypography.BODY_SIZE_PT : compactLayoutContext ? COMPACT_BODY_FONT_SIZE_PT : 11.0d;
        addParagraphStyle(styles, STYLE_NORMAL, "Normal", bodyFontSize, false);
        addParagraphStyle(styles, STYLE_BODY_TEXT, "Body Text", bodyFontSize, false);
        addParagraphStyle(styles, STYLE_TITLE, "Title", 18, true);
        addParagraphStyle(styles, STYLE_HEADING_1, "Heading 1", 12, true);
        addParagraphStyle(styles, STYLE_HEADING_2, "Heading 2", 13, true);
    }

    private void addParagraphStyle(XWPFStyles styles, String styleId, String name, double fontSizePt, boolean bold) {
        CTStyle style = CTStyle.Factory.newInstance();
        style.setStyleId(styleId);
        style.setType(STStyleType.PARAGRAPH);
        if (STYLE_NORMAL.equals(styleId)) {
            style.setDefault(XmlBoolean.Factory.newValue(true));
        }

        CTString styleName = style.addNewName();
        styleName.setVal(name);

        CTOnOff qFormat = style.addNewQFormat();
        qFormat.setVal(XmlBoolean.Factory.newValue(true));

        CTPPrGeneral pPr = style.addNewPPr();
        CTSpacing spacing = pPr.addNewSpacing();
        spacing.setAfter(BigInteger.ZERO);

        CTRPr rPr = style.addNewRPr();
        CTFonts fonts = rPr.addNewRFonts();
        fonts.setAscii(examTypographyContext ? "Times New Roman" : "宋体");
        fonts.setHAnsi(examTypographyContext ? "Times New Roman" : "宋体");
        fonts.setCs(examTypographyContext ? "Times New Roman" : "宋体");
        fonts.setEastAsia("宋体");
        CTHpsMeasure size = rPr.addNewSz();
        size.setVal(halfPoints(fontSizePt));
        CTHpsMeasure csSize = rPr.addNewSzCs();
        csSize.setVal(halfPoints(fontSizePt));
        if (bold) {
            CTOnOff b = rPr.addNewB();
            b.setVal(XmlBoolean.Factory.newValue(true));
            CTOnOff bCs = rPr.addNewBCs();
            bCs.setVal(XmlBoolean.Factory.newValue(true));
        }

        styles.addStyle(new XWPFStyle(style));
    }

    /**
     * 构建完整的 .docx 文档。
     *
     * <p>这是文档生成的主入口方法，按以下顺序组装文档内容：</p>
     * <ol>
     *   <li>设置 A4 页面布局和页边距</li>
     *   <li>写入试卷标题（居中加粗）</li>
     *   <li>写入试卷信息（总分、建议时长）</li>
     *   <li>遍历所有大题（Section），逐个写入标题和包含的小题</li>
     *   <li>序列化为字节数组返回</li>
     * </ol>
     *
     * @param request 试卷导出请求，包含试卷信息、大题列表及其下的小题数据
     * @return .docx 文件的字节数组，可直接写入文件或通过 HTTP 返回
     * @throws IOException 文档序列化失败时抛出
     */
    public byte[] build(PaperExportRequest request) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
            mathEmbedder.resetDocumentFormulaCounter();
            compactLayoutContext = isCompactLayout(request);
            String typography = request != null && request.getPaper() != null ? request.getPaper().getTypography() : null;
            validateMode("paper.typography", typography, "exam");
            examTypographyContext = "exam".equals(typography);
            if (examTypographyContext && compactLayoutContext) {
                throw new com.lz.paperword.model.ExportRequestValidationException("paper.typography=exam cannot be combined with compactLayout");
            }
            String background = request != null && request.getPaper() != null ? request.getPaper().getPreviewBackground() : null;
            if (background != null && !"transparent".equals(background) && !"white".equals(background)) {
                throw new com.lz.paperword.model.ExportRequestValidationException("paper.previewBackground must be transparent or white");
            }
            if ("white".equals(background) && !examTypographyContext) {
                throw new com.lz.paperword.model.ExportRequestValidationException("paper.previewBackground=white requires paper.typography=exam");
            }
            previewBackgroundContext = "white".equals(background)
                ? com.lz.paperword.core.render.LaTeXImageRenderer.PreviewBackground.WHITE
                : com.lz.paperword.core.render.LaTeXImageRenderer.PreviewBackground.TRANSPARENT;
            metadataContext = false;
            formulaParagraphs.clear();
            metadataParagraphs.clear();
            workingParagraphs.clear();
            printLayoutContext = request != null && request.getPaper() != null
                && Boolean.TRUE.equals(request.getPaper().getPrintLayout());
            hideQuestionTypeMetadataContext = request != null
                    && request.getPaper() != null
                    && Boolean.TRUE.equals(request.getPaper().getHideQuestionTypeMetadata());
            initializeStyles(doc);
            // 设置 A4 页面和默认页边距
            setPageMargins(doc, compactLayoutContext);
            if (compactLayoutContext) {
                addCompactFooter(doc, request.getPaper());
            } else if (request != null && request.getPaper() != null
                    && Boolean.TRUE.equals(request.getPaper().getPageNumbers())) {
                addPageNumberFooter(doc);
            }

            // 1. 写入试卷标题
            writePaperTitle(doc, request.getPaper());

            // 2. 写入试卷信息行（总分 + 建议时长）
            writePaperInfo(doc, request.getPaper());

            // 3. 遍历各大题，写入大题标题和小题内容
            if (request.getSections() != null) {
                for (SectionDTO section : request.getSections()) {
                    writeSection(doc, section);
                }
            }
            ensureFontTablePart(doc);
            if (examTypographyContext) applyExamTypography(doc);
            if (printLayoutContext || examTypographyContext) {
                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    CTPPr properties = paragraph.getCTP().isSetPPr() ? paragraph.getCTP().getPPr() : paragraph.getCTP().addNewPPr();
                    (properties.isSetKeepLines() ? properties.getKeepLines() : properties.addNewKeepLines()).setVal(true);
                }
            }

            // 将文档序列化为字节数组
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            doc.write(baos);
            return baos.toByteArray();
        }
    }

    private void ensureFontTablePart(XWPFDocument doc) throws IOException {
        try {
            OPCPackage pkg = doc.getPackage();
            PackagePartName fontTableName = PackagingURIHelper.createPartName("/word/fontTable.xml");
            if (!pkg.containPart(fontTableName)) {
                PackagePart fontTablePart = pkg.createPart(fontTableName, FONT_TABLE_CONTENT_TYPE);
                try (OutputStream out = fontTablePart.getOutputStream()) {
                    out.write(minimalFontTableXml().getBytes(StandardCharsets.UTF_8));
                }
            }
            PackagePart documentPart = doc.getPackagePart();
            if (!hasRelationship(documentPart, fontTableName, FONT_TABLE_RELATION)) {
                documentPart.addRelationship(fontTableName, TargetMode.INTERNAL, FONT_TABLE_RELATION);
            }
        } catch (Exception e) {
            throw new IOException("Failed to add Word font table part", e);
        }
    }

    private boolean hasRelationship(PackagePart part, PackagePartName target, String relationshipType) throws Exception {
        for (PackageRelationship rel : part.getRelationshipsByType(relationshipType)) {
            if (target.getURI().equals(rel.getTargetURI())) {
                return true;
            }
        }
        return false;
    }

    private String minimalFontTableXml() {
        return """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:font w:name="宋体">
                <w:charset w:val="86"/>
                <w:family w:val="roman"/>
                <w:pitch w:val="variable"/>
              </w:font>
              <w:font w:name="Times New Roman">
                <w:charset w:val="00"/>
                <w:family w:val="roman"/>
                <w:pitch w:val="variable"/>
              </w:font>
              <w:font w:name="Cambria Math">
                <w:charset w:val="00"/>
                <w:family w:val="roman"/>
                <w:pitch w:val="variable"/>
              </w:font>
            </w:fonts>
            """;
    }

    /**
     * 设置文档页面布局：A4 纸张和页边距。
     *
     * <p>单位说明：Word 使用 twip 作为度量单位，1 英寸 = 1440 twips。</p>
     * <ul>
     *   <li>页边距：上下左右各 1440 twips（1 英寸 ≈ 2.54cm）</li>
     *   <li>纸张尺寸：宽 11906 twips（210mm），高 16838 twips（297mm）= 标准 A4</li>
     * </ul>
     *
     * @param doc Word 文档对象
     */
    private boolean isCompactLayout(PaperExportRequest request) {
        return request != null
            && request.getPaper() != null
            && Boolean.TRUE.equals(request.getPaper().getCompactLayout());
    }

    private void setPageMargins(XWPFDocument doc, boolean compactLayout) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        // 页边距设置（单位：twips，1440 twips = 1 英寸 ≈ 2.54cm）
        long vertical = examTypographyContext ? 1020L : compactLayout ? 851L : 1440L;
        long horizontal = examTypographyContext ? 1247L : compactLayout ? 1134L : 1440L;
        pageMar.setTop(BigInteger.valueOf(vertical));
        pageMar.setBottom(BigInteger.valueOf(vertical));
        pageMar.setLeft(BigInteger.valueOf(horizontal));
        pageMar.setRight(BigInteger.valueOf(horizontal));
        if (compactLayout) {
            pageMar.setHeader(BigInteger.valueOf(851L));
            pageMar.setFooter(BigInteger.valueOf(851L));
            pageMar.setGutter(BigInteger.ZERO);
        }

        if (examTypographyContext) pageMar.setFooter(BigInteger.valueOf(425L));

        // A4 纸张尺寸（单位：twips）
        CTPageSz pageSz = sectPr.addNewPgSz();
        pageSz.setW(BigInteger.valueOf(compactLayout ? 11907L : 11906L)); // A4 宽度：210mm
        pageSz.setH(BigInteger.valueOf(compactLayout ? 16840L : 16838L)); // A4 高度：297mm
        if (compactLayout) {
            pageSz.setCode(BigInteger.valueOf(9L));
            CTColumns cols = sectPr.isSetCols() ? sectPr.getCols() : sectPr.addNewCols();
            cols.setSpace(BigInteger.valueOf(425L));
        }
    }

    private void addCompactFooter(XWPFDocument doc, PaperExportRequest.PaperInfo paper) {
        try {
            XWPFHeaderFooterPolicy policy = doc.createHeaderFooterPolicy();
            XWPFFooter footer = policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFParagraph line = footer.createParagraph();
            line.setBorderTop(Borders.THIN_THICK_SMALL_GAP);
            CTPBdr pBdr = line.getCTP().getPPr().getPBdr();
            pBdr.getTop().setSz(BigInteger.valueOf(24));
            pBdr.getTop().setSpace(BigInteger.ONE);
            line.setSpacingBefore(0);
            line.setSpacingAfter(0);

            XWPFParagraph para = footer.createParagraph();
            para.setSpacingBefore(0);
            para.setSpacingAfter(0);
            configureFooterTabs(para);
            XWPFRun left = para.createRun();
            left.setText(compactFooterLeftText(paper));
            setFooterRunStyle(left);
            para.createRun().addTab();
            XWPFRun center = para.createRun();
            center.setText("教师版");
            setFooterRunStyle(center);
            para.createRun().addTab();
            XWPFRun pageLabel = para.createRun();
            pageLabel.setText("page ");
            setFooterRunStyle(pageLabel);
            addFieldRun(para, "PAGE");
            XWPFRun ofLabel = para.createRun();
            ofLabel.setText(" of ");
            setFooterRunStyle(ofLabel);
            addFieldRun(para, "NUMPAGES");
        } catch (Exception e) {
            log.warn("Failed to add compact footer", e);
        }
    }

    private void addPageNumberFooter(XWPFDocument doc) throws IOException {
        try {
            XWPFFooter footer = doc.createHeaderFooterPolicy().createFooter(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFParagraph paragraph = footer.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            paragraph.setSpacingBefore(0);
            paragraph.setSpacingAfter(0);
            for (String part : new String[]{"第 ", "PAGE", " 页 / 共 ", "NUMPAGES", " 页"}) {
                if (part.equals("PAGE") || part.equals("NUMPAGES")) addFieldRun(paragraph, part);
                else { XWPFRun run = paragraph.createRun(); run.setText(part); setFooterRunStyle(run); }
            }
        } catch (Exception e) {
            throw new IOException("Failed to add requested page-number footer", e);
        }
    }

    private void keepWithNext(XWPFParagraph paragraph) {
        CTPPr properties = paragraph.getCTP().isSetPPr() ? paragraph.getCTP().getPPr() : paragraph.getCTP().addNewPPr();
        (properties.isSetKeepNext() ? properties.getKeepNext() : properties.addNewKeepNext()).setVal(true);
    }

    private String compactFooterLeftText(PaperExportRequest.PaperInfo paper) {
        if (paper != null && paper.getSourceName() != null && !paper.getSourceName().isBlank()) {
            return paper.getSourceName().replaceFirst("(?i)\\.docx$", "");
        }
        return "题库";
    }

    private void configureFooterTabs(XWPFParagraph para) {
        CTPPr pPr = para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr();
        CTTabs tabs = pPr.isSetTabs() ? pPr.getTabs() : pPr.addNewTabs();
        CTTabStop center = tabs.addNewTab();
        center.setVal(STTabJc.CENTER);
        center.setPos(BigInteger.valueOf(4800));
        CTTabStop right = tabs.addNewTab();
        right.setVal(STTabJc.RIGHT);
        right.setPos(BigInteger.valueOf(9600));
    }

    private void setFooterRunStyle(XWPFRun run) {
        run.setFontFamily("宋体");
        run.setFontSize(9);
    }

    private void addFieldRun(XWPFParagraph para, String instruction) {
        XWPFRun begin = para.createRun();
        begin.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        XWPFRun instr = para.createRun();
        instr.getCTR().addNewInstrText().setStringValue(instruction);
        XWPFRun separate = para.createRun();
        separate.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        XWPFRun value = para.createRun();
        value.setText("1");
        setFooterRunStyle(value);
        XWPFRun end = para.createRun();
        end.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    /**
     * 写入试卷标题：居中、加粗、18pt 宋体。
     *
     * @param doc   Word 文档对象
     * @param paper 试卷基本信息（包含试卷名称）
     */
    private void writePaperTitle(XWPFDocument doc, PaperExportRequest.PaperInfo paper) {
        if (paper == null || paper.getName() == null) return;

        XWPFParagraph titlePara = doc.createParagraph();
        titlePara.setStyle(STYLE_TITLE);
        titlePara.setAlignment(ParagraphAlignment.CENTER);  // 居中对齐
        titlePara.setSpacingAfter(200);                      // 段后间距

        XWPFRun titleRun = titlePara.createRun();
        titleRun.setText(paper.getName());
        titleRun.setBold(true);       // 加粗
        titleRun.setFontSize(18);     // 18pt 大号字体
        titleRun.setFontFamily("宋体");
    }

    /**
     * 写入试卷信息行：总分和建议时长，居中显示。
     *
     * <p>格式示例："总分：100分    建议时长：90分钟"</p>
     *
     * @param doc   Word 文档对象
     * @param paper 试卷基本信息
     */
    private void writePaperInfo(XWPFDocument doc, PaperExportRequest.PaperInfo paper) {
        if (paper == null) return;

        // 组装信息文本
        StringBuilder info = new StringBuilder();
        if (paper.getScore() != null) {
            info.append("总分：").append(paper.getScore()).append("分");
        }
        if (paper.getSuggestTime() != null) {
            if (!info.isEmpty()) info.append("    "); // 用空格分隔
            info.append("建议时长：").append(paper.getSuggestTime()).append("分钟");
        }

        if (!info.isEmpty()) {
            XWPFParagraph infoPara = doc.createParagraph();
            infoPara.setStyle(STYLE_BODY_TEXT);
            infoPara.setAlignment(ParagraphAlignment.CENTER);
            infoPara.setSpacingAfter(200);
            if (examTypographyContext) metadataParagraphs.add(infoPara);

            XWPFRun infoRun = infoPara.createRun();
            infoRun.setText(info.toString());
            infoRun.setFontSize(11);
            infoRun.setFontFamily("宋体");
        }
    }

    /**
     * 写入一个大题（Section）：包含大题标题和其下的所有小题。
     *
     * @param doc     Word 文档对象
     * @param section 大题数据（包含标题和小题列表）
     */
    private void writeSection(XWPFDocument doc, SectionDTO section) {
        if (section == null) return;

        // 写入大题标题（如"一、选择题"）：加粗、12pt 宋体
        if (section.getHeadline() != null) {
            XWPFParagraph headPara = doc.createParagraph();
            headPara.setStyle(STYLE_HEADING_1);
            headPara.setSpacingBefore(300); // 段前间距（与上文拉开距离）
            headPara.setSpacingAfter(100);

            XWPFRun headRun = headPara.createRun();
            headRun.setText(section.getHeadline());
            headRun.setBold(true);
            headRun.setFontSize(12);
            headRun.setFontFamily("宋体");
            if (printLayoutContext || examTypographyContext) keepWithNext(headPara);
        }

        if (section.getImages() != null && !section.getImages().isEmpty()) {
            int maxWidthPx = section.getImageMaxWidthPx() != null ? section.getImageMaxWidthPx() : 560;
            ParagraphAlignment alignment = resolveImageAlignment(section.getImageAlignment());
            for (String imagePath : section.getImages()) {
                writeQuestionImage(doc, imagePath, maxWidthPx, alignment);
            }
        }

        // 逐个写入大题下的所有小题
        if (section.getQuestions() != null) {
            for (QuestionDTO q : section.getQuestions()) {
                writeQuestion(doc, q);
            }
        }
    }

    /**
     * 写入一道小题的完整内容，包括题目、选项、答题空间、答案和解析。
     *
     * <p>处理的题目类型：</p>
     * <ul>
     *   <li>type=1：单选题（写题目 + 选项）</li>
     *   <li>type=2：多选题（写题目 + 选项）</li>
     *   <li>type=3：判断题（写题目 + 选项）</li>
     *   <li>type=4：填空题（题目中含下划线，不额外添加答题空间）</li>
     *   <li>type=5/6：简答题/计算题（题目后添加空白答题区域）</li>
     * </ul>
     *
     * @param doc      Word 文档对象
     * @param question 小题数据
     */
    private void writeQuestion(XWPFDocument doc, QuestionDTO question) {
        if (question == null) return;
        int firstParagraph = doc.getParagraphs().size();
        validateMode("question.contentFlow", question.getContentFlow(), "flow");
        boolean flow = "flow".equals(question.getContentFlow());
        if (flow && question.getContent() != null && Pattern.compile("(?i)<(?:img|table)\\b").matcher(question.getContent()).find()) {
            throw new com.lz.paperword.model.ExportRequestValidationException(
                "question.contentFlow=flow does not support HTML img/table blocks; use question.images or the layout export route");
        }
        Integer requestedLines = question.getAnswerSpaceLines();
        if (requestedLines != null && (requestedLines < 0 || requestedLines > 12)) {
            throw new com.lz.paperword.model.ExportRequestValidationException("answerSpaceLines must be between 0 and 12");
        }

        // ---- 教学阶段小节标题（课堂讲解/课堂练习/课后作业等） ----
        if (question.getPhaseLabel() != null && !question.getPhaseLabel().isBlank()) {
            writePhaseHeader(doc, question.getPhaseLabel());
        }

        int stemStart = doc.getParagraphs().size();
        // ---- 题目内容行 ----
        // 题号前缀（如 "1. "、"2. "）
        String prefix = !compactLayoutContext && question.getSerialNumber() != null ?
            question.getSerialNumber() + ". " : "";

        // 题干中的 <br/> 需要拆成多个段落，否则复杂竖式会和题干挤在同一行
        List<String> contentLines = flow ? ContentFlowLayout.split(question.getContent()) : splitHtmlContentLines(question.getContent());
        if (contentLines.isEmpty()) {
            XWPFParagraph qPara = doc.createParagraph();
            qPara.setStyle(STYLE_BODY_TEXT);
            qPara.setSpacingBefore(compactSpacing(100, 0));
            qPara.setSpacingAfter(0);
            applyCompactParagraphRhythm(qPara);
            writeContentWithMath(qPara, prefix, question.getContent());
        } else {
            for (int i = 0; i < contentLines.size(); i++) {
                XWPFParagraph qPara = doc.createParagraph();
                qPara.setStyle(STYLE_BODY_TEXT);
                qPara.setSpacingBefore(i == 0 ? compactSpacing(100, 35) : 0);
                qPara.setSpacingAfter(0);
                applyCompactParagraphRhythm(qPara);

                String line = contentLines.get(i);
                if (i > 0) {
                    applyCompactBodyIndent(qPara, 315);
                    if (!examTypographyContext && isStandaloneDisplayFormula(line)) {
                        qPara.setAlignment(ParagraphAlignment.CENTER);
                        qPara.setIndentationLeft(0);
                        qPara.setIndentationFirstLine(0);
                    }
                } else if (shouldUseCompactFirstLineIndent(question, line)) {
                    applyCompactBodyIndent(qPara, isNumberedQuestion(question) ? 0 : 315);
                }

                writeContentWithMath(qPara, i == 0 ? prefix : "", line);
            }
        }

        // ---- 题目配图（渲染在题干之后、选项之前，居中显示） ----
        if (question.getImages() != null && !question.getImages().isEmpty()) {
            for (String imagePath : question.getImages()) {
                writeQuestionImage(doc, imagePath, compactLayoutContext ? 520 : 200, ParagraphAlignment.CENTER);
            }
        }
        if (question.getUnresolvedImages() != null && !question.getUnresolvedImages().isEmpty()) {
            writeInlineLabeledSection(doc, "【缺失图片】", String.join("、", question.getUnresolvedImages()));
        }

        // ---- 选项（选择题：type 1=单选, 2=多选, 3=判断） ----
        if (question.getOptions() != null && !question.getOptions().isEmpty()) {
            writeOptions(doc, question.getOptions());
        }

        int answerSpaceStart = doc.getParagraphs().size();
        // ---- 答题空间（根据题目类型决定） ----
        if (question.getQuestionType() != null) {
            int type = question.getQuestionType();
            if (type == 4) {
                // 填空题：题目内容中已包含下划线占位符，不需要额外空间
            } else if (!hideQuestionTypeMetadataContext && (type == 5 || type == 6)
                    && isNumberedQuestion(question) && !hasResolvedContent(question)) {
                // 简答题/计算题：在题目后添加 3 行空白区域供答题
                addAnswerSpace(doc, requestedLines == null ? 3 : requestedLines);
            }
        }

        if (compactLayoutContext) {
            writeCompactMetadata(doc, question);
        } else {
            if (question.getKnowledgePoint() != null && !question.getKnowledgePoint().isBlank()) {
                writeInlineLabeledSection(doc, "【知识点】", question.getKnowledgePoint());
            }
            if (question.getTags() != null && !question.getTags().isEmpty()) {
                writeInlineLabeledSection(doc, "【标签】", String.join("、", question.getTags()));
            }
            if (question.getDifficulty() != null && !question.getDifficulty().isBlank()) {
                writeInlineLabeledSection(doc, "【难度】", question.getDifficulty());
            }
        }
        if (question.getAnalyze() != null && !question.getAnalyze().isBlank()) {
            writeInlineLabeledSection(doc, "【解析】", question.getAnalyze(), usesLooseFormulaContinuation(question));
        }
        int solutionStart = -1;
        if (question.getSolution() != null && !question.getSolution().isBlank()) {
            solutionStart = doc.getParagraphs().size();
            writeMultilineLabeledSection(doc, "【解答】", question.getSolution(), usesLooseFormulaContinuation(question));
        }

        // ---- 正确答案（标签 + 答案内容） ----
        if (question.getCorrect() != null && !question.getCorrect().isBlank()) {
            writeAnswerSection(doc, question.getCorrect());
        }
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        int end = paragraphs.size();
        if (examTypographyContext) {
            for (int i = stemStart; i < end; i++) {
                ExamTypography.indent(paragraphs.get(i), i == stemStart && !prefix.isEmpty());
            }
        }
        if (Boolean.TRUE.equals(question.getPageBreakBefore()) && end > firstParagraph) {
            paragraphs.get(firstParagraph).setPageBreak(true);
        }
        if (printLayoutContext || examTypographyContext) {
            applyQuestionKeeps(paragraphs, firstParagraph, answerSpaceStart, solutionStart, end,
                question.getCorrect() != null && !question.getCorrect().isBlank());
        }
    }

    /**
     * 写入教学阶段小节标题（如“课堂讲解”），加粗稍大字号，用于在专题内区分讲练阶段。
     *
     * @param doc   Word 文档对象
     * @param label 阶段标题文本
     */
    private void writePhaseHeader(XWPFDocument doc, String label) {
        XWPFParagraph para = doc.createParagraph();
        para.setStyle(STYLE_HEADING_2);
        para.setSpacingBefore(compactSpacing(180, 0));
        para.setSpacingAfter(0);
        XWPFRun run = para.createRun();
        run.setText("【" + label + "】");
        run.setBold(true);
        run.setFontSize(13);
        run.setFontFamily("宋体");
        run.setColor("C00000");
    }

    /**
     * 在题目中插入一张本地配图，居中显示，按最大宽度等比缩放。
     *
     * @param doc       Word 文档对象
     * @param imagePath 图片本地文件路径
     */
    private ParagraphAlignment resolveImageAlignment(String alignment) {
        if (alignment == null) {
            return ParagraphAlignment.CENTER;
        }
        return switch (alignment.toLowerCase(java.util.Locale.ROOT)) {
            case "left" -> ParagraphAlignment.LEFT;
            case "right" -> ParagraphAlignment.RIGHT;
            default -> ParagraphAlignment.CENTER;
        };
    }

    private void writeQuestionImage(XWPFDocument doc, String imagePath, int maxWidthPx, ParagraphAlignment alignment) {
        ImageAssetLoader.ImageAsset image = imageAssets.load(imagePath);
        if (maxWidthPx <= 0) {
            throw new ImageAssetException("IMAGE_ASSET_INVALID_PATH", "Image display width must be positive");
        }
        int wPx = image.widthPx();
        int hPx = image.heightPx();
        // Fit both axes before EMU conversion, including extreme panoramic/tall images.
        double scale = Math.min(1.0, Math.min((double) Math.min(maxWidthPx, 600) / wPx, 900.0 / hPx));
        wPx = Math.max(1, (int) Math.round(wPx * scale));
        hPx = Math.max(1, (int) Math.round(hPx * scale));

        XWPFParagraph para = doc.createParagraph();
        para.setStyle(STYLE_BODY_TEXT);
        para.setAlignment(alignment);
        para.setSpacingBefore(compactSpacing(50, 0));
        para.setSpacingAfter(0);
        applyCompactContinuationRhythm(para);
        XWPFRun run = para.createRun();
        try (ByteArrayInputStream input = new ByteArrayInputStream(image.data())) {
            run.addPicture(input, image.pictureType(), image.fileName(),
                org.apache.poi.util.Units.pixelToEMU(wPx),
                org.apache.poi.util.Units.pixelToEMU(hPx));
        } catch (Exception e) {
            throw new ImageAssetException("IMAGE_ASSET_EMBED_FAILED", "Failed to embed a validated question or section image", e);
        }
    }

    /**
     * 写入选择题选项，采用两列布局（每行两个选项）。
     *
     * <p>选项按顺序两两一组写在同一行，用 Tab 制表符分隔。
     * 如果选项总数为奇数，最后一个选项独占一行。
     * 所有选项行有统一的左缩进（400 twips）。</p>
     *
     * @param doc     Word 文档对象
     * @param options 选项列表（通常为 A、B、C、D）
     */
    private void writeOptions(XWPFDocument doc, List<QuestionDTO.OptionDTO> options) {
        // 每次处理两个选项，写在同一行
        for (int i = 0; i < options.size(); i += 2) {
            XWPFParagraph optPara = doc.createParagraph();
            optPara.setStyle(STYLE_BODY_TEXT);
            optPara.setIndentationLeft(400); // 左缩进，使选项区域与题目区分
            optPara.setSpacingBefore(0);
            optPara.setSpacingAfter(0);

            // 写入第一个选项（如 "A. 选项内容"）
            QuestionDTO.OptionDTO opt1 = options.get(i);
            writeOptionContent(optPara, opt1);

            // 如果有第二个选项，用 Tab 间隔后写在同一行
            if (i + 1 < options.size()) {
                XWPFRun tabRun = optPara.createRun();
                tabRun.addTab();  // 添加多个 Tab 实现列对齐
                tabRun.addTab();
                tabRun.addTab();

                QuestionDTO.OptionDTO opt2 = options.get(i + 1);
                writeOptionContent(optPara, opt2);
            }
        }
    }

    /**
     * 写入单个选项的内容（前缀 + 内容文本/公式）。
     *
     * @param para   段落对象
     * @param option 选项数据（含前缀如 "A" 和内容）
     */
    private void writeOptionContent(XWPFParagraph para, QuestionDTO.OptionDTO option) {
        String optPrefix = option.getPrefix() != null ? option.getPrefix() + ". " : "";
        writeContentWithMath(para, optPrefix, option.getContent());
    }

    /**
     * 写入可能包含 LaTeX 数学公式的混合内容。
     *
     * <p>这是文档生成中最核心的方法之一，负责将 HTML 格式的内容（可能包含 $...$ 公式）
     * 正确地写入 Word 段落。处理流程：</p>
     * <ol>
     *   <li>写入纯文本前缀（如题号 "1. "、选项前缀 "A. "）</li>
     *   <li>调用 {@link LaTeXParser#parseHtml(String)} 将 HTML 内容分割为纯文本段和数学公式段</li>
     *   <li>对每个内容段：
     *     <ul>
     *       <li>纯文本段：创建 Run 直接写入文本（宋体 11pt）</li>
     *       <li>数学公式段（embedMathTypeOle=true）：通过 {@link MathTypeEmbedder} 嵌入为 OLE 对象</li>
     *       <li>数学公式段（embedMathTypeOle=false）：保留 $...$ 标记作为纯文本（草稿模式）</li>
     *     </ul>
     *   </li>
     * </ol>
     *
     * <p>公式嵌入失败时的降级策略：如果 MathType OLE 嵌入抛出异常，
     * 会自动降级为纯文本 $LaTeX$ 格式并记录错误日志。</p>
     *
     * @param para        目标段落
     * @param prefix      纯文本前缀（如题号、选项标识）
     * @param htmlContent HTML 格式的内容（可能包含 LaTeX 公式）
     */
    private void writeContentWithMath(XWPFParagraph para, String prefix, String htmlContent) {
        writeContentWithMath(para, prefix, htmlContent, COMPACT_BODY_FONT, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void writeContentWithMath(XWPFParagraph para, String prefix, String htmlContent, String textFontFamily) {
        writeContentWithMath(para, prefix, htmlContent, textFontFamily, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void writeContentWithMath(XWPFParagraph para, String prefix, String htmlContent, String textFontFamily,
                                      double textFontSizePt) {
        // 写入纯文本前缀
        if (prefix != null && !prefix.isEmpty()) {
            XWPFRun prefixRun = para.createRun();
            if (examTypographyContext && prefix.matches("\\d+\\.\\s+")) {
                prefixRun.setText(prefix.trim());
                prefixRun.addTab();
            } else prefixRun.setText(prefix);
            applyBodyRunStyle(prefixRun, textFontFamily, textFontSizePt);
        }

        if (htmlContent == null || htmlContent.isBlank()) return;

        // 解析 HTML 内容，提取纯文本段和 LaTeX 数学公式段
        List<ContentSegment> segments = latexParser.parseHtml(htmlContent);
        if (examTypographyContext && segments.stream().anyMatch(segment -> segment.metrics() != null
                || (segment.styleHints() != null && segment.styleHints().sourceMetrics() != null))) {
            throw new com.lz.paperword.model.ExportRequestValidationException(
                "paper.typography=exam cannot be combined with source formula metrics");
        }
        boolean displayParagraph = isStandaloneDisplayFormula(htmlContent)
            || (segments.stream().anyMatch(ContentSegment::isMath)
                && segments.stream().filter(segment -> !segment.isMath())
                    .allMatch(segment -> segment.rawText().matches("[\\s\\p{P}]*")));
        if (segments.stream().anyMatch(ContentSegment::isMath)) formulaParagraphs.add(para);
        if (metadataContext) metadataParagraphs.add(para);
        boolean containsAnswerLine = segments.stream()
            .filter(seg -> !seg.isMath())
            .anyMatch(seg -> normalizeAnswerLineMarkers(seg.rawText()).contains(ANSWER_LINE_MARKER));
        boolean firstSegment = true;

        for (ContentSegment seg : segments) {
            if (seg.isMath() && seg.ast() != null) {
                // ---- 数学公式段：嵌入为 MathType OLE 或保留为文本标记 ----
                XWPFRun mathRun = para.createRun();
                mathRun.setFontSize(runFontSizePt(COMPACT_BODY_FONT_SIZE_PT));
                mathRun.setFontFamily("Cambria Math"); // 数学公式使用 Cambria Math 字体

                if (embedMathTypeOle) {
                    // OLE 嵌入模式：将 AST 转换为 MathType OLE 对象
                    try {
                        double displayScale = resolveCompactFormulaDisplayScale(seg.rawText(), containsAnswerLine,
                            firstSegment);
                        double maxWidthPt = resolveCompactFormulaMaxWidth(seg.rawText(), containsAnswerLine,
                            firstSegment);
                        if (examTypographyContext) {
                            mathEmbedder.embedEquationAtSize(para, mathRun, seg.ast(), seg.rawText(),
                                metadataContext ? ExamTypography.METADATA_SIZE_PT : ExamTypography.BODY_SIZE_PT,
                                displayParagraph, ExamTypography.FORMULA_MAX_WIDTH_PT, seg.styleHints(), previewBackgroundContext);
                        } else {
                            mathEmbedder.embedEquation(para, mathRun, seg.ast(), seg.rawText(),
                                displayScale, maxWidthPt, seg.metrics(), seg.styleHints());
                        }
                    } catch (com.lz.paperword.model.ExportRequestValidationException e) {
                        throw e;
                    } catch (Exception e) {
                        throw new IllegalStateException("Failed to embed formula: " + seg.rawText(), e);
                    }
                } else {
                    // 草稿模式：保留 $...$ 标记，供 Windows 端 MathType 后处理器批量转换
                    mathRun.setText("$" + seg.rawText() + "$");
                }
            } else {
                // ---- 纯文本段：直接写入 ----
                writeTextSegment(para, normalizeAnswerLineMarkers(seg.rawText()), firstSegment,
                    textFontFamily, textFontSizePt);
            }
            firstSegment = false;
        }
    }

    private double resolveCompactFormulaDisplayScale(String rawLatex, boolean containsAnswerLine,
                                                     boolean firstSegment) {
        return 1.0d;
    }

    private double resolveCompactFormulaMaxWidth(String rawLatex, boolean containsAnswerLine,
                                                 boolean firstSegment) {
        if (!compactLayoutContext) {
            return Double.MAX_VALUE;
        }
        String latex = rawLatex == null ? "" : rawLatex;
        if (containsAnswerLine) {
            return isLongFormula(latex) ? 360.0d : 330.0d;
        }
        if (isArrayLikeFormula(latex)) {
            return isLongFormula(latex) ? (firstSegment ? 380.0d : 290.0d) : 260.0d;
        }
        if (containsFractionFormula(latex)) {
            if (isLongFormula(latex)) {
                return firstSegment ? (isDenseFractionFormula(latex) ? 360.0d : 380.0d) : 285.0d;
            }
            return 240.0d;
        }
        return 220.0d;
    }

    private boolean isArrayLikeFormula(String latex) {
        return latex.contains("\\begin{array}")
            || latex.contains("\\begin{aligned}")
            || latex.contains("\\begin{matrix}");
    }

    private boolean containsFractionFormula(String latex) {
        return latex.contains("\\frac")
            || latex.contains("\\dfrac")
            || latex.contains("\\cfrac");
    }

    private boolean isLongFormula(String latex) {
        return latex.length() >= 120 || latex.contains("\\cdots");
    }

    private boolean isDenseFractionFormula(String latex) {
        return countOccurrences(latex, "\\frac") >= 4
            || latex.contains("\\cdots")
            || latex.contains("1+2+")
            || latex.contains("\\times 100")
            || latex.contains("1993");
    }

    private int countOccurrences(String value, String needle) {
        int count = 0;
        int index = 0;
        while (value != null && (index = value.indexOf(needle, index)) >= 0) {
            count++;
            index += needle.length();
        }
        return count;
    }

    private String normalizeAnswerLineMarkers(String htmlContent) {
        if (htmlContent == null || htmlContent.isBlank() || htmlContent.contains(ANSWER_LINE_MARKER)) {
            return htmlContent;
        }
        String normalized = replaceEqualsAnswerBlanks(htmlContent);
        return STANDALONE_ANSWER_BLANK_PATTERN.matcher(normalized).replaceAll(ANSWER_LINE_MARKER);
    }

    private String replaceEqualsAnswerBlanks(String text) {
        Matcher matcher = EQUALS_ANSWER_BLANK_PATTERN.matcher(text);
        StringBuffer buffer = new StringBuffer();
        while (matcher.find()) {
            String punctuation = "";
            if (matcher.group(3) != null) {
                punctuation = matcher.group(3);
            } else if (matcher.group(2) != null) {
                punctuation = matcher.group(2);
            }
            matcher.appendReplacement(buffer, Matcher.quoteReplacement(matcher.group(1) + ANSWER_LINE_MARKER + punctuation));
        }
        matcher.appendTail(buffer);
        return buffer.toString();
    }

    private void writeTextSegment(XWPFParagraph para, String text, boolean firstSegment) {
        writeTextSegment(para, text, firstSegment, COMPACT_BODY_FONT, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void writeTextSegment(XWPFParagraph para, String text, boolean firstSegment, String textFontFamily) {
        writeTextSegment(para, text, firstSegment, textFontFamily, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void writeTextSegment(XWPFParagraph para, String text, boolean firstSegment, String textFontFamily,
                                  double textFontSizePt) {
        if (text != null && text.contains(ANSWER_LINE_MARKER)) {
            writeTextWithAnswerLine(para, text, firstSegment, textFontFamily, textFontSizePt);
            return;
        }
        if (compactLayoutContext && firstSegment && text != null) {
            java.util.regex.Matcher matcher = java.util.regex.Pattern
                .compile("^(【例\\s*\\d+】|【巩固】)(.*)$", java.util.regex.Pattern.DOTALL)
                .matcher(text);
            if (matcher.matches()) {
                XWPFRun labelRun = para.createRun();
                labelRun.setText(matcher.group(1));
                labelRun.setBold(true);
                labelRun.setFontSize(runFontSizePt(COMPACT_TITLE_FONT_SIZE_PT));
                labelRun.setFontFamily(COMPACT_BODY_FONT);
                labelRun.setColor("C00000");
                String remainder = matcher.group(2);
                if (remainder != null && !remainder.isEmpty()) {
                    XWPFRun remainderRun = para.createRun();
                    remainderRun.setText(remainder);
                    remainderRun.setFontSize(runFontSizePt(COMPACT_TITLE_FONT_SIZE_PT));
                    remainderRun.setFontFamily(COMPACT_BODY_FONT);
                    remainderRun.setBold(true);
                    remainderRun.setColor("C00000");
                }
                return;
            }
        }

        XWPFRun textRun = para.createRun();
        textRun.setText(text);
        applyBodyRunStyle(textRun, textFontFamily, textFontSizePt);
    }

    private void writeTextWithAnswerLine(XWPFParagraph para, String text, boolean firstSegment) {
        writeTextWithAnswerLine(para, text, firstSegment, COMPACT_BODY_FONT, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void writeTextWithAnswerLine(XWPFParagraph para, String text, boolean firstSegment, String textFontFamily) {
        writeTextWithAnswerLine(para, text, firstSegment, textFontFamily, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void writeTextWithAnswerLine(XWPFParagraph para, String text, boolean firstSegment,
                                         String textFontFamily, double textFontSizePt) {
        String[] parts = text.split(java.util.regex.Pattern.quote(ANSWER_LINE_MARKER), -1);
        for (int i = 0; i < parts.length; i++) {
            String part = parts[i];
            if (!part.isEmpty()) {
                writeTextSegment(para, part, firstSegment && i == 0, textFontFamily, textFontSizePt);
            }
            if (i < parts.length - 1) {
                writeAnswerLineRun(para);
            }
        }
    }

    private void writeAnswerLineRun(XWPFParagraph para) {
        XWPFRun lineRun = para.createRun();
        lineRun.setText("       ");
        lineRun.setFontSize(runFontSizePt(COMPACT_TITLE_FONT_SIZE_PT));
        lineRun.setFontFamily(COMPACT_BODY_FONT);
        lineRun.setColor("000000");
        lineRun.setUnderline(UnderlinePatterns.SINGLE);
    }

    /**
     * 写入正确答案区域：蓝色加粗的"【答案】"标签 + 答案内容。
     *
     * <p>答案内容也可能包含 LaTeX 公式（如数学题的答案），
     * 因此同样通过 {@link #writeContentWithMath} 处理。</p>
     *
     * @param doc     Word 文档对象
     * @param correct 正确答案内容（可能为 HTML 格式）
     */
    private void writeAnswerSection(XWPFDocument doc, String correct) {
        XWPFParagraph para = doc.createParagraph();
        para.setStyle(STYLE_BODY_TEXT);
        para.setSpacingBefore(compactSpacing(100, 45));
        para.setSpacingAfter(0);
        applyCompactParagraphRhythm(para);

        writeLabelRun(para, "【答案】", COMPACT_ANSWER_FONT);
        writeContentWithMath(para, "", correct, COMPACT_ANSWER_FONT, COMPACT_ANSWER_FONT_SIZE_PT);
    }

    /**
     * 写入解析说明区域：蓝色加粗的"【解析】"标签 + 多行解析内容。
     *
     * <p>解析内容可能包含 HTML 换行标签（&lt;br/&gt;），本方法会按换行标签
     * 将内容拆分为多行，每行作为独立段落写入。每行内容也支持 LaTeX 公式。</p>
     *
     * @param doc         Word 文档对象
     * @param analyzeHtml 解析说明内容（HTML 格式，可能包含 &lt;br/&gt; 换行）
     */
    private void writeInlineLabeledSection(XWPFDocument doc, String label, String content) {
        writeInlineLabeledSection(doc, label, content, false);
    }

    private void writeInlineLabeledSection(XWPFDocument doc, String label, String content,
                                           boolean looseFormulaContinuation) {
        boolean previousMetadata = metadataContext;
        metadataContext = examTypographyContext && java.util.Set.of("【知识点】", "【标签】", "【难度】").contains(label);
        List<String> lines = splitHtmlContentLines(content);
        if (lines.isEmpty()) {
            lines = List.of(content);
        }
        for (int i = 0; i < lines.size(); i++) {
            XWPFParagraph para = doc.createParagraph();
            para.setStyle(STYLE_BODY_TEXT);
            para.setSpacingBefore(i == 0 ? compactSpacing(100, 45) : 0);
            para.setSpacingAfter(0);
            if (i == 0) {
                applyCompactParagraphRhythm(para);
            } else if (looseFormulaContinuation && isFormulaDominantLine(lines.get(i))) {
                applyCompactFormulaContinuationRhythm(para);
            } else {
                applyCompactContinuationRhythm(para);
            }
            if (i == 0) {
                writeLabelRun(para, label);
            }
            writeContentWithMath(para, "", lines.get(i));
        }
        metadataContext = previousMetadata;
    }

    private void writeCompactMetadata(XWPFDocument doc, QuestionDTO question) {
        List<LabeledValue> values = new ArrayList<>();
        if (question.getKnowledgePoint() != null && !question.getKnowledgePoint().isBlank()) {
            values.add(new LabeledValue("【考点】", question.getKnowledgePoint()));
        }
        if (question.getDifficulty() != null && !question.getDifficulty().isBlank()) {
            values.add(new LabeledValue("【难度】", question.getDifficulty()));
        }
        if (!hideQuestionTypeMetadataContext && isNumberedQuestion(question) && question.getQuestionType() != null) {
            values.add(new LabeledValue("【题型】", questionTypeName(question.getQuestionType())));
        }
        if (question.getTags() != null && !question.getTags().isEmpty()) {
            values.add(new LabeledValue("【关键词】", String.join("、", question.getTags())));
        }
        if (values.isEmpty()) {
            return;
        }

        XWPFParagraph para = doc.createParagraph();
        para.setStyle(STYLE_BODY_TEXT);
        para.setSpacingBefore(0);
        para.setSpacingAfter(0);
        applyCompactParagraphRhythm(para);
        for (int i = 0; i < values.size(); i++) {
            if (i > 0) {
                XWPFRun separator = para.createRun();
                separator.setText("    ");
                applyBodyRunStyle(separator);
            }
            LabeledValue value = values.get(i);
            writeLabelRun(para, value.label());
            writeContentWithMath(para, "", value.value());
        }
    }

    private String questionTypeName(Integer type) {
        return switch (type) {
            case 1 -> "单选";
            case 2 -> "多选";
            case 3 -> "判断";
            case 4 -> "填空";
            case 5 -> "解答";
            case 6 -> "计算";
            default -> "综合";
        };
    }

    private boolean isNumberedQuestion(QuestionDTO question) {
        return question != null && question.getSerialNumber() != null;
    }

    private void writeMultilineLabeledSection(XWPFDocument doc, String label, String contentHtml) {
        writeMultilineLabeledSection(doc, label, contentHtml, false);
    }

    private void writeMultilineLabeledSection(XWPFDocument doc, String label, String contentHtml,
                                              boolean looseFormulaContinuation) {
        writeInlineLabeledSection(doc, label, normalizeHtmlLineBreaks(contentHtml), looseFormulaContinuation);
    }

    private void writeLabelRun(XWPFParagraph para, String label) {
        writeLabelRun(para, label, COMPACT_BODY_FONT);
    }

    private void writeLabelRun(XWPFParagraph para, String label, String fontFamily) {
        XWPFRun labelRun = para.createRun();
        labelRun.setText(label);
        labelRun.setBold(true);
        double sizePt = COMPACT_ANSWER_FONT.equals(fontFamily) ? COMPACT_ANSWER_FONT_SIZE_PT : COMPACT_BODY_FONT_SIZE_PT;
        applyBodyRunStyle(labelRun, fontFamily, sizePt);
    }

    private int compactSpacing(int normalTwips, int compactTwips) {
        return compactLayoutContext ? compactTwips : normalTwips;
    }

    private void applyCompactParagraphRhythm(XWPFParagraph para) {
        if (!compactLayoutContext) {
            return;
        }
        applyCompactLineHeight(para, 255);
    }

    private void applyCompactContinuationRhythm(XWPFParagraph para) {
        if (!compactLayoutContext) {
            return;
        }
        applyCompactLineHeight(para, 250);
    }

    private void applyCompactFormulaContinuationRhythm(XWPFParagraph para) {
        if (!compactLayoutContext) {
            return;
        }
        applyCompactLineHeight(para, 265);
    }

    private void applyCompactLineHeight(XWPFParagraph para, int lineTwips) {
        para.setSpacingBetween(1.0d, LineSpacingRule.AUTO);
        CTPPr pPr = para.getCTP().isSetPPr() ? para.getCTP().getPPr() : para.getCTP().addNewPPr();
        CTSpacing spacing = pPr.isSetSpacing() ? pPr.getSpacing() : pPr.addNewSpacing();
        spacing.setLine(BigInteger.valueOf(lineTwips));
        spacing.setLineRule(STLineSpacingRule.AUTO);
    }

    private void applyCompactBodyIndent(XWPFParagraph para, int firstLineTwips) {
        if (!compactLayoutContext || firstLineTwips <= 0) {
            return;
        }
        para.setIndentationLeft(0);
        para.setIndentationFirstLine(firstLineTwips);
    }

    private boolean shouldUseCompactFirstLineIndent(QuestionDTO question, String line) {
        if (!compactLayoutContext || isNumberedQuestion(question) || line == null) {
            return false;
        }
        String text = Jsoup.parse(line.replaceAll("\\$[^$]*\\$", " ")).body().text().trim();
        if (text.isEmpty()) {
            return false;
        }
        return !text.matches("^([一二三四五六七八九十]+、|\\(\\d+\\)|【[^】]+】).*");
    }

    private boolean isFormulaDominantLine(String line) {
        if (!compactLayoutContext || line == null) {
            return false;
        }
        String strippedText = Jsoup.parse(line.replaceAll("\\$[^$]*\\$", " ")).body().text().trim();
        int mathMarkers = countOccurrences(line, "$");
        return mathMarkers >= 2 && strippedText.length() <= 8;
    }

    private boolean usesLooseFormulaContinuation(QuestionDTO question) {
        if (!compactLayoutContext || question == null || question.getSerialNumber() == null) {
            return false;
        }
        return question.getSerialNumber() >= 53;
    }

    private void applyBodyRunStyle(XWPFRun run) {
        applyBodyRunStyle(run, COMPACT_BODY_FONT);
    }

    private void applyBodyRunStyle(XWPFRun run, String fontFamily) {
        applyBodyRunStyle(run, fontFamily, COMPACT_BODY_FONT_SIZE_PT);
    }

    private void applyBodyRunStyle(XWPFRun run, String fontFamily, double fontSizePt) {
        if (examTypographyContext) {
            ExamTypography.styleRun(run, metadataContext ? ExamTypography.METADATA_SIZE_PT : ExamTypography.BODY_SIZE_PT);
        } else {
            setRunFontSize(run, fontSizePt);
            run.setFontFamily(fontFamily);
        }
    }

    private int runFontSizePt(double fontSizePt) {
        return (int) Math.round(fontSizePt);
    }

    private BigInteger halfPoints(double fontSizePt) {
        return BigInteger.valueOf(Math.round(fontSizePt * 2.0d));
    }

    private void setRunFontSize(XWPFRun run, double fontSizePt) {
        CTRPr rPr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTHpsMeasure size = rPr.sizeOfSzArray() > 0 ? rPr.getSzArray(0) : rPr.addNewSz();
        size.setVal(halfPoints(fontSizePt));
        CTHpsMeasure csSize = rPr.sizeOfSzCsArray() > 0 ? rPr.getSzCsArray(0) : rPr.addNewSzCs();
        csSize.setVal(halfPoints(fontSizePt));
    }

    private record LabeledValue(String label, String value) {
    }

    private List<String> splitHtmlContentLines(String htmlContent) {
        return ContentFlowLayout.splitLegacy(htmlContent);
    }

    private boolean isStandaloneDisplayFormula(String line) {
        return ContentFlowLayout.isDisplayFormula(line);
    }

    private boolean hasResolvedContent(QuestionDTO question) {
        return (question.getCorrect() != null && !question.getCorrect().isBlank())
            || (question.getAnalyze() != null && !question.getAnalyze().isBlank())
            || (question.getSolution() != null && !question.getSolution().isBlank());
    }

    private String normalizeHtmlLineBreaks(String contentHtml) {
        if (contentHtml == null || contentHtml.isBlank()) {
            return contentHtml;
        }
        String[] lines = contentHtml.split("(?i)<br\\s*/?>\\s*");
        StringBuilder builder = new StringBuilder();
        String mathDelimiter = "";
        for (String line : lines) {
            String trimmed = line.trim();
            if (trimmed.isEmpty()) {
                continue;
            }
            String cleanText = Jsoup.parse(trimmed).body().text().trim();
            if (cleanText.isEmpty()) {
                continue;
            }
            if (!builder.isEmpty()) {
                // Keep the explicit paragraph boundary understood by
                // splitHtmlContentLines. A single newline is collapsed by the
                // subsequent HTML/text pass and used to join every solution step.
                // HTML line wrapping inside an existing math span is not a
                // paragraph boundary: preserve the complete delimited formula.
                builder.append(mathDelimiter.isEmpty() ? "<br/>" : " ");
            }
            builder.append(cleanText);
            mathDelimiter = solutionMathDelimiterAfter(cleanText, mathDelimiter);
        }
        return builder.toString();
    }

    private String solutionMathDelimiterAfter(String text, String delimiter) {
        int backslashes = 0;
        for (int i = 0; i < text.length(); i++) {
            boolean escaped = (backslashes & 1) != 0;
            backslashes = text.charAt(i) == '\\' ? backslashes + 1 : 0;
            if (escaped) continue;
            if (!delimiter.isEmpty()) {
                if (text.startsWith(delimiter, i)) { i += delimiter.length() - 1; delimiter = ""; backslashes = 0; }
            } else if (text.startsWith("$$", i)) { delimiter = "$$"; i++; }
            else if (text.charAt(i) == '$') delimiter = "$";
            else if (text.startsWith("\\[", i)) { delimiter = "\\]"; i++; backslashes = 0; }
            else if (text.startsWith("\\(", i)) { delimiter = "\\)"; i++; backslashes = 0; }
        }
        return delimiter;
    }

    /**
     * 添加空白答题区域（用于简答题/计算题）。
     *
     * <p>通过插入指定行数的空段落来为学生提供书写答案的空间。
     * 每个空段落有 200 twips 的段后间距。</p>
     *
     * @param doc   Word 文档对象
     * @param lines 空白行数
     */
    private void addAnswerSpace(XWPFDocument doc, int lines) {
        for (int i = 0; i < lines; i++) {
            XWPFParagraph blankPara = doc.createParagraph();
            blankPara.setStyle(STYLE_BODY_TEXT);
            blankPara.setSpacingAfter(200);
            if (printLayoutContext || examTypographyContext) ExamTypography.workingLine(blankPara);
            workingParagraphs.add(blankPara);
            XWPFRun blankRun = blankPara.createRun();
            blankRun.setText("");
        }
    }
}
