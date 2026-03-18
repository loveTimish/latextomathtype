package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.latex.LaTeXParser.ContentSegment;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.jsoup.Jsoup;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

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
public class DocxBuilder {

    private static final Logger log = LoggerFactory.getLogger(DocxBuilder.class);

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
        this.embedMathTypeOle = embedMathTypeOle;
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
            // 设置 A4 页面和默认页边距
            setPageMargins(doc);

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

            // 将文档序列化为字节数组
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            doc.write(baos);
            return baos.toByteArray();
        }
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
    private void setPageMargins(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        // 页边距设置（单位：twips，1440 twips = 1 英寸 ≈ 2.54cm）
        pageMar.setTop(BigInteger.valueOf(1440));
        pageMar.setBottom(BigInteger.valueOf(1440));
        pageMar.setLeft(BigInteger.valueOf(1440));
        pageMar.setRight(BigInteger.valueOf(1440));

        // A4 纸张尺寸（单位：twips）
        CTPageSz pageSz = sectPr.addNewPgSz();
        pageSz.setW(BigInteger.valueOf(11906)); // A4 宽度：210mm
        pageSz.setH(BigInteger.valueOf(16838)); // A4 高度：297mm
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

        XWPFParagraph infoPara = doc.createParagraph();
        infoPara.setAlignment(ParagraphAlignment.CENTER);
        infoPara.setSpacingAfter(200);

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
            headPara.setSpacingBefore(300); // 段前间距（与上文拉开距离）
            headPara.setSpacingAfter(100);

            XWPFRun headRun = headPara.createRun();
            headRun.setText(section.getHeadline());
            headRun.setBold(true);
            headRun.setFontSize(12);
            headRun.setFontFamily("宋体");
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

        // ---- 题目内容行 ----
        // 题号前缀（如 "1. "、"2. "）
        String prefix = question.getSerialNumber() != null ?
            question.getSerialNumber() + ". " : "";

        // 题干中的 <br/> 需要拆成多个段落，否则复杂竖式会和题干挤在同一行
        List<String> contentLines = splitHtmlContentLines(question.getContent());
        if (contentLines.isEmpty()) {
            XWPFParagraph qPara = doc.createParagraph();
            qPara.setSpacingBefore(100);
            qPara.setSpacingAfter(50);
            writeContentWithMath(qPara, prefix, question.getContent());
        } else {
            for (int i = 0; i < contentLines.size(); i++) {
                XWPFParagraph qPara = doc.createParagraph();
                qPara.setSpacingBefore(i == 0 ? 100 : 0);
                qPara.setSpacingAfter(50);

                String line = contentLines.get(i);
                if (i > 0) {
                    qPara.setIndentationLeft(360);
                    if (isStandaloneDisplayFormula(line)) {
                        qPara.setAlignment(ParagraphAlignment.CENTER);
                        qPara.setIndentationLeft(0);
                    }
                }

                writeContentWithMath(qPara, i == 0 ? prefix : "", line);
            }
        }

        // ---- 选项（选择题：type 1=单选, 2=多选, 3=判断） ----
        if (question.getOptions() != null && !question.getOptions().isEmpty()) {
            writeOptions(doc, question.getOptions());
        }

        // ---- 答题空间（根据题目类型决定） ----
        if (question.getQuestionType() != null) {
            int type = question.getQuestionType();
            if (type == 4) {
                // 填空题：题目内容中已包含下划线占位符，不需要额外空间
            } else if ((type == 5 || type == 6) && !hasResolvedContent(question)) {
                // 简答题/计算题：在题目后添加 3 行空白区域供答题
                addAnswerSpace(doc, 3);
            }
        }

        if (question.getKnowledgePoint() != null && !question.getKnowledgePoint().isBlank()) {
            writeInlineLabeledSection(doc, "【知识点】", question.getKnowledgePoint());
        }
        if (question.getTags() != null && !question.getTags().isEmpty()) {
            writeInlineLabeledSection(doc, "【标签】", String.join("、", question.getTags()));
        }
        if (question.getDifficulty() != null && !question.getDifficulty().isBlank()) {
            writeInlineLabeledSection(doc, "【难度】", question.getDifficulty());
        }
        if (question.getAnalyze() != null && !question.getAnalyze().isBlank()) {
            writeInlineLabeledSection(doc, "【分析】", question.getAnalyze());
        }
        if (question.getSolution() != null && !question.getSolution().isBlank()) {
            writeMultilineLabeledSection(doc, "【解答】", question.getSolution());
        }

        // ---- 正确答案（标签 + 答案内容） ----
        if (question.getCorrect() != null && !question.getCorrect().isBlank()) {
            writeAnswerSection(doc, question.getCorrect());
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
            optPara.setIndentationLeft(400); // 左缩进，使选项区域与题目区分

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
        // 写入纯文本前缀
        if (prefix != null && !prefix.isEmpty()) {
            XWPFRun prefixRun = para.createRun();
            prefixRun.setText(prefix);
            prefixRun.setFontSize(11);
            prefixRun.setFontFamily("宋体");
        }

        if (htmlContent == null || htmlContent.isBlank()) return;

        // 解析 HTML 内容，提取纯文本段和 LaTeX 数学公式段
        List<ContentSegment> segments = latexParser.parseHtml(htmlContent);

        for (ContentSegment seg : segments) {
            if (seg.isMath() && seg.ast() != null) {
                // ---- 数学公式段：嵌入为 MathType OLE 或保留为文本标记 ----
                XWPFRun mathRun = para.createRun();
                mathRun.setFontSize(11);
                mathRun.setFontFamily("Cambria Math"); // 数学公式使用 Cambria Math 字体

                if (embedMathTypeOle) {
                    // OLE 嵌入模式：将 AST 转换为 MathType OLE 对象
                    try {
                        mathEmbedder.embedEquation(para, mathRun, seg.ast(), seg.rawText());
                    } catch (Exception e) {
                        // 嵌入失败时降级为纯文本显示
                        log.error("Failed to embed formula, falling back to text: {}", seg.rawText(), e);
                        mathRun.setText("$" + seg.rawText() + "$");
                    }
                } else {
                    // 草稿模式：保留 $...$ 标记，供 Windows 端 MathType 后处理器批量转换
                    mathRun.setText("$" + seg.rawText() + "$");
                }
            } else {
                // ---- 纯文本段：直接写入 ----
                XWPFRun textRun = para.createRun();
                textRun.setText(seg.rawText());
                textRun.setFontSize(11);
                textRun.setFontFamily("宋体");
            }
        }
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
        para.setSpacingBefore(100);

        writeLabelRun(para, "【答案】");
        writeContentWithMath(para, "", correct);
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
        XWPFParagraph para = doc.createParagraph();
        para.setSpacingBefore(100);
        writeLabelRun(para, label);
        writeContentWithMath(para, "", content);
    }

    private void writeMultilineLabeledSection(XWPFDocument doc, String label, String contentHtml) {
        XWPFParagraph para = doc.createParagraph();
        para.setSpacingBefore(100);
        writeLabelRun(para, label);
        writeContentWithMath(para, "", normalizeHtmlLineBreaks(contentHtml));
    }

    private void writeLabelRun(XWPFParagraph para, String label) {
        XWPFRun labelRun = para.createRun();
        labelRun.setText(label);
        labelRun.setBold(true);
        labelRun.setFontSize(11);
        labelRun.setFontFamily("宋体");
    }

    private List<String> splitHtmlContentLines(String htmlContent) {
        if (htmlContent == null || htmlContent.isBlank()) {
            return List.of();
        }
        String[] parts = htmlContent.split("(?i)<br\\s*/?>");
        List<String> lines = new java.util.ArrayList<>();
        for (String part : parts) {
            String trimmed = part == null ? "" : part.trim();
            if (!trimmed.isEmpty()) {
                lines.add(trimmed);
            }
        }
        return lines;
    }

    private boolean isStandaloneDisplayFormula(String line) {
        if (line == null) {
            return false;
        }
        String trimmed = line.trim();
        return (trimmed.startsWith("\\[") && trimmed.endsWith("\\]"))
            || (trimmed.startsWith("$$") && trimmed.endsWith("$$"));
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
                builder.append('\n');
            }
            builder.append(cleanText);
        }
        return builder.toString();
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
            blankPara.setSpacingAfter(200);
            XWPFRun blankRun = blankPara.createRun();
            blankRun.setText("");
        }
    }
}
