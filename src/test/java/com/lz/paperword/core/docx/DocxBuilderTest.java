package com.lz.paperword.core.docx;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class DocxBuilderTest {

    private final DocxBuilder builder = new DocxBuilder();

    @Test
    void testBuildSimplePaper(@TempDir Path tempDir) throws IOException {
        PaperExportRequest request = createSampleRequest();
        byte[] docx = builder.build(request);

        assertNotNull(docx);
        assertTrue(docx.length > 0);

        // Verify it's a valid ZIP (docx is ZIP-based)
        assertEquals(0x50, docx[0] & 0xFF); // PK signature
        assertEquals(0x4B, docx[1] & 0xFF);

        // Save to file for manual inspection
        Path outputFile = tempDir.resolve("test_paper.docx");
        Files.write(outputFile, docx);
        assertTrue(Files.exists(outputFile));
        assertTrue(Files.size(outputFile) > 0);
    }

    @Test
    void testBuildWithMathFormulas(@TempDir Path tempDir) throws IOException {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("数学公式测试");
        paper.setScore(100);
        paper.setSuggestTime(60);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、计算题");

        QuestionDTO q1 = new QuestionDTO();
        q1.setSerialNumber(1);
        q1.setQuestionType(6);
        q1.setContent("<p>计算 $\\frac{x^{2}+1}{x-1}$ 的值</p>");
        q1.setScore(10);

        QuestionDTO q2 = new QuestionDTO();
        q2.setSerialNumber(2);
        q2.setQuestionType(6);
        q2.setContent("<p>已知 $\\sqrt{x+3}=5$，求 $x$ 的值</p>");
        q2.setScore(10);

        QuestionDTO q3 = new QuestionDTO();
        q3.setSerialNumber(3);
        q3.setQuestionType(6);
        q3.setContent("<p>求 $\\sum_{i=1}^{n}i^{2}$ 的公式</p>");
        q3.setScore(15);

        section.setQuestions(List.of(q1, q2, q3));
        request.setSections(List.of(section));

        byte[] docx = builder.build(request);

        assertNotNull(docx);
        assertTrue(docx.length > 100);

        Path outputFile = tempDir.resolve("math_test.docx");
        Files.write(outputFile, docx);
        System.out.println("Test docx saved to: " + outputFile);
    }

    @Test
    void testBuildVennDiagramQuestion() throws IOException {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("数学思维训练");
        paper.setSubjectType(2);
        paper.setScore(10);
        paper.setSuggestTime(15);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、填空题（共10分）");

        QuestionDTO q = new QuestionDTO();
        q.setSerialNumber(1);
        q.setQuestionType(4); // 填空题
        q.setScore(10);
        q.setContent("干洗店统计洗上衣、裙子和裤子的顾客人数，只洗上衣的顾客与只洗裤子和裙子两样的顾客都是9人；" +
                "只洗裤子的顾客与不洗裤子的顾客人数相同；三样全洗、只洗一样、只洗两样的顾客人数相同；只洗上" +
                "衣和裤子两样的顾客有15人，洗裙子的顾客有48人。则一共有____位顾客");
        q.setCorrect("90");
        q.setAnalyze("画出韦恩图，分析数量关系即可解答。<br/><br/>" +
                "【解答】解：<br/><br/>" +
                "根据题意可得：<br/><br/>" +
                "$$ B+C+D+9=48 $$……①<br/><br/>" +
                "$$ B+C+9=A $$……②<br/><br/>" +
                "$$ D=A+B+9=15+9+C $$……③<br/><br/>" +
                "把①②③式联立后可得：<br/><br/>" +
                "$$ A=18 $$，$$ B=3 $$，$$ C=6 $$，$$ D=30 $$。<br/><br/>" +
                "$$ 9+15+A+C+D+9+B $$<br/><br/>" +
                "$$ =9+15+18+6+30+9+3 $$<br/><br/>" +
                "$$ =90 $$（位）<br/><br/>" +
                "答：一共有90位顾客。<br/><br/>" +
                "故答案为：90。");

        section.setQuestions(List.of(q));
        request.setSections(List.of(section));

        byte[] docx = builder.build(request);

        assertNotNull(docx);
        assertTrue(docx.length > 100);
        assertEquals(0x50, docx[0] & 0xFF); // PK signature
        assertEquals(0x4B, docx[1] & 0xFF);

        // Save to a fixed output location for manual inspection
        Path outputDir = Path.of("E:/lingzhi/extensions/paper-to-word/output");
        Files.createDirectories(outputDir);
        Path outputFile = outputDir.resolve("venn_diagram_test.docx");
        Files.write(outputFile, docx);
        System.out.println("=== Venn Diagram test docx saved to: " + outputFile + " ===");
        System.out.println("File size: " + docx.length + " bytes");
    }

    @Test
    void testBuildEmptyPaper() throws IOException {
        PaperExportRequest request = new PaperExportRequest();
        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("空试卷");
        request.setPaper(paper);

        byte[] docx = builder.build(request);

        assertNotNull(docx);
        assertTrue(docx.length > 0);
    }

    @Test
    void testSymbolRendering() throws IOException {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("符号渲染测试（TNR）");
        paper.setScore(100);
        paper.setSuggestTime(30);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、符号验证");

        QuestionDTO q1 = new QuestionDTO();
        q1.setSerialNumber(1);
        q1.setQuestionType(6);
        q1.setContent("<p>圆的面积公式为 $S = \\pi r^{2}$</p>");
        q1.setScore(10);

        QuestionDTO q2 = new QuestionDTO();
        q2.setSerialNumber(2);
        q2.setQuestionType(6);
        q2.setContent("<p>计算 $3 \\times 5 = 15$</p>");
        q2.setScore(10);

        QuestionDTO q3 = new QuestionDTO();
        q3.setSerialNumber(3);
        q3.setQuestionType(6);
        q3.setContent("<p>综合: $\\alpha + \\beta = \\gamma$, $a \\div b = c$, $x \\pm y$, $A \\leq B$, $\\theta \\geq \\Omega$</p>");
        q3.setScore(10);

        QuestionDTO q4 = new QuestionDTO();
        q4.setSerialNumber(4);
        q4.setQuestionType(6);
        q4.setContent("<p>向量 $\\vec{a} \\cdot \\vec{b} = |a||b|\\cos\\theta$</p>");
        q4.setScore(10);

        QuestionDTO q5 = new QuestionDTO();
        q5.setSerialNumber(5);
        q5.setQuestionType(6);
        q5.setContent("<p>求 $\\int_{0}^{\\pi} \\sin x \\, dx$ 和 $\\sum_{i=1}^{n} i$</p>");
        q5.setScore(10);

        section.setQuestions(List.of(q1, q2, q3, q4, q5));
        request.setSections(List.of(section));

        byte[] docx = builder.build(request);
        assertNotNull(docx);

        Path outputDir = Path.of("E:/lingzhi/extensions/paper-to-word/output");
        Files.createDirectories(outputDir);
        Path outputFile = outputDir.resolve("symbol_test_tnr.docx");
        Files.write(outputFile, docx);
        System.out.println("=== Symbol test (TNR) saved to: " + outputFile + " ===");
        System.out.println("File size: " + docx.length + " bytes");
    }

    @Test
    void testK12AllFormulas() throws IOException {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("K12 全学段数学公式综合测试");
        paper.setScore(150);
        paper.setSuggestTime(120);
        request.setPaper(paper);

        // ==================== 小学阶段 ====================
        SectionDTO sec1 = new SectionDTO();
        sec1.setHeadline("一、小学基础公式");

        QuestionDTO e1 = q(1, "<p>四则运算: $3 \\times 5 = 15$, $24 \\div 6 = 4$, $7 + 8 = 15$, $20 - 13 = 7$</p>");
        QuestionDTO e2 = q(2, "<p>分数运算: $\\frac{1}{2} + \\frac{1}{3} = \\frac{5}{6}$</p>");
        QuestionDTO e3 = q(3, "<p>混合运算: $\\frac{2}{3} \\times \\frac{3}{4} = \\frac{1}{2}$</p>");
        QuestionDTO e4 = q(4, "<p>圆的面积: $S = \\pi r^{2}$, 圆的周长: $C = 2\\pi r$</p>");
        QuestionDTO e5 = q(5, "<p>长方体体积: $V = abc$, 球的体积: $V = \\frac{4}{3}\\pi r^{3}$</p>");

        sec1.setQuestions(List.of(e1, e2, e3, e4, e5));

        // ==================== 初中阶段 ====================
        SectionDTO sec2 = new SectionDTO();
        sec2.setHeadline("二、初中核心公式");

        QuestionDTO m1 = q(6, "<p>完全平方公式: $(a + b)^{2} = a^{2} + 2ab + b^{2}$</p>");
        QuestionDTO m2 = q(7, "<p>平方差公式: $(a + b)(a - b) = a^{2} - b^{2}$</p>");
        QuestionDTO m3 = q(8, "<p>一元二次方程求根公式: $x = \\frac{-b \\pm \\sqrt{b^{2} - 4ac}}{2a}$</p>");
        QuestionDTO m4 = q(9, "<p>勾股定理: $a^{2} + b^{2} = c^{2}$</p>");
        QuestionDTO m5 = q(10, "<p>绝对值方程: $|x - 3| = 5$, 则 $x = 8$ 或 $x = -2$</p>");
        QuestionDTO m6 = q(11, "<p>二次函数顶点式: $y = a(x - h)^{2} + k$</p>");
        QuestionDTO m7 = q(12, "<p>一次函数: $y = kx + b$, 斜率 $k = \\frac{y_{2} - y_{1}}{x_{2} - x_{1}}$</p>");
        QuestionDTO m8 = q(13, "<p>不等式: $2x + 1 > 5$ 解得 $x > 2$</p>");
        QuestionDTO m9 = q(14, "<p>比例: $\\frac{a}{b} = \\frac{c}{d}$, 则 $ad = bc$</p>");
        QuestionDTO m10 = q(15, "<p>根式运算: $\\sqrt{8} = 2\\sqrt{2}$, $\\sqrt[3]{27} = 3$</p>");

        sec2.setQuestions(List.of(m1, m2, m3, m4, m5, m6, m7, m8, m9, m10));

        // ==================== 高中阶段 ====================
        SectionDTO sec3 = new SectionDTO();
        sec3.setHeadline("三、高中重要公式");

        QuestionDTO h1 = q(16, "<p>三角恒等式: $\\sin^{2}\\theta + \\cos^{2}\\theta = 1$</p>");
        QuestionDTO h2 = q(17, "<p>二倍角公式: $\\sin 2\\alpha = 2\\sin\\alpha\\cos\\alpha$</p>");
        QuestionDTO h3 = q(18, "<p>余弦定理: $c^{2} = a^{2} + b^{2} - 2ab\\cos C$</p>");
        QuestionDTO h4 = q(19, "<p>正弦定理: $\\frac{a}{\\sin A} = \\frac{b}{\\sin B} = \\frac{c}{\\sin C} = 2R$</p>");
        QuestionDTO h5 = q(20, "<p>等差数列通项: $a_{n} = a_{1} + (n - 1)d$</p>");
        QuestionDTO h6 = q(21, "<p>等差数列求和: $S_{n} = \\frac{n(a_{1} + a_{n})}{2}$</p>");
        QuestionDTO h7 = q(22, "<p>等比数列求和: $S_{n} = a_{1} \\cdot \\frac{1 - q^{n}}{1 - q}$</p>");
        QuestionDTO h8 = q(23, "<p>向量点积: $\\vec{a} \\cdot \\vec{b} = |a||b|\\cos\\theta$</p>");
        QuestionDTO h9 = q(24, "<p>椭圆标准方程: $\\frac{x^{2}}{a^{2}} + \\frac{y^{2}}{b^{2}} = 1$</p>");
        QuestionDTO h10 = q(25, "<p>双曲线标准方程: $\\frac{x^{2}}{a^{2}} - \\frac{y^{2}}{b^{2}} = 1$</p>");
        QuestionDTO h11 = q(26, "<p>定积分: $\\int_{0}^{\\pi} \\sin x dx$</p>");
        QuestionDTO h12 = q(27, "<p>求和公式: $\\sum_{k=1}^{n} k^{2} = \\frac{n(n+1)(2n+1)}{6}$</p>");
        QuestionDTO h13 = q(28, "<p>指数与对数: $a^{x} = N$, 则 $x = \\log_{a} N$</p>");
        QuestionDTO h14 = q(29, "<p>排列组合: $C_{n}^{k} = \\frac{n!}{k!(n-k)!}$</p>");
        QuestionDTO h15 = q(30, "<p>复数模: $|z| = \\sqrt{a^{2} + b^{2}}$, 其中 $z = a + bi$</p>");

        sec3.setQuestions(List.of(h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12, h13, h14, h15));

        request.setSections(List.of(sec1, sec2, sec3));

        byte[] docx = builder.build(request);
        assertNotNull(docx);

        Path outputDir = Path.of("E:/lingzhi/extensions/paper-to-word/output");
        Files.createDirectories(outputDir);
        Path outputFile = outputDir.resolve("k12_formulas_v3.docx");
        Files.write(outputFile, docx);
        System.out.println("=== K12 formulas saved to: " + outputFile + " ===");
        System.out.println("File size: " + docx.length + " bytes");
        System.out.println("Total formulas: 30 questions covering elementary/middle/high school");
    }

    private QuestionDTO q(int serial, String content) {
        QuestionDTO q = new QuestionDTO();
        q.setSerialNumber(serial);
        q.setQuestionType(6);
        q.setContent(content);
        q.setScore(5);
        return q;
    }

    private PaperExportRequest createSampleRequest() {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("期中数学测试");
        paper.setSubjectType(2);
        paper.setStage(2);
        paper.setScore(100);
        paper.setSuggestTime(90);
        request.setPaper(paper);

        // Section 1: 选择题
        SectionDTO section1 = new SectionDTO();
        section1.setHeadline("一、选择题（每题3分，共15分）");

        QuestionDTO q1 = new QuestionDTO();
        q1.setSerialNumber(1);
        q1.setQuestionType(1);
        q1.setContent("<p>已知 $\\frac{x+1}{2}=3$，则 $x$ 的值为</p>");
        q1.setScore(3);

        QuestionDTO.OptionDTO optA = new QuestionDTO.OptionDTO();
        optA.setPrefix("A");
        optA.setContent("$5$");
        QuestionDTO.OptionDTO optB = new QuestionDTO.OptionDTO();
        optB.setPrefix("B");
        optB.setContent("$3$");
        QuestionDTO.OptionDTO optC = new QuestionDTO.OptionDTO();
        optC.setPrefix("C");
        optC.setContent("$7$");
        QuestionDTO.OptionDTO optD = new QuestionDTO.OptionDTO();
        optD.setPrefix("D");
        optD.setContent("$1$");
        q1.setOptions(List.of(optA, optB, optC, optD));

        section1.setQuestions(List.of(q1));

        // Section 2: 解答题
        SectionDTO section2 = new SectionDTO();
        section2.setHeadline("二、解答题（每题10分）");

        QuestionDTO q2 = new QuestionDTO();
        q2.setSerialNumber(2);
        q2.setQuestionType(5);
        q2.setContent("<p>求方程 $x^{2}-5x+6=0$ 的所有根</p>");
        q2.setScore(10);

        section2.setQuestions(List.of(q2));

        request.setSections(List.of(section1, section2));
        return request;
    }
}
