package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 生成 K12 常见公式与竖式样例文档，便于人工集中检查公式效果。
 */
class K12FormulaShowcaseDocxTest {

    @Test
    void shouldGenerateK12FormulaShowcaseDocx() throws IOException {
        DocxBuilder builder = new DocxBuilder();
        byte[] docx = builder.build(buildRequest());

        Path outputDir = Path.of("target", "generated-docs");
        Files.createDirectories(outputDir);

        // 使用时间戳避免覆盖正在打开的旧文件。
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path output = outputDir.resolve("k12-formula-showcase-" + timestamp + ".docx");
        Files.write(output, docx);

        System.out.println("Generated K12 showcase document: " + output);
        assertTrue(Files.exists(output), "生成的文档应存在");
        assertTrue(Files.size(output) > 0, "生成的文档不应为空");
    }

    private PaperExportRequest buildRequest() {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("K12 常见公式与竖式样例");
        paper.setScore(100);
        paper.setSuggestTime(60);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、K12 常见公式与竖式总览");
        section.setQuestions(List.of(
            createQuestion(1, 10, """
                小学常见公式：<br/>
                长方形周长：$C=2(a+b)$<br/>
                长方形面积：$S=ab$<br/>
                正方形周长：$C=4a$<br/>
                正方形面积：$S=a^2$<br/>
                三角形面积：$S=\\frac{ah}{2}$<br/>
                平行四边形面积：$S=ah$<br/>
                梯形面积：$S=\\frac{(a+b)h}{2}$<br/>
                圆周长：$C=2\\pi r$<br/>
                圆面积：$S=\\pi r^2$<br/>
                平均数：$\\bar{x}=\\frac{x_1+x_2+\\cdots+x_n}{n}$
                """),
            createQuestion(2, 10, """
                小学分数、小数、百分数：<br/>
                分数加法：$\\frac{3}{4}+\\frac{1}{2}=\\frac{5}{4}$<br/>
                分数乘法：$\\frac{2}{3}\\times\\frac{3}{5}=\\frac{2}{5}$<br/>
                分数除法：$\\frac{3}{4}\\div\\frac{1}{2}=\\frac{3}{2}$<br/>
                小数乘法：$1.25\\times0.4=0.5$<br/>
                百分数：$25\\%=\\frac{1}{4}=0.25$<br/>
                比例：$\\frac{a}{b}=\\frac{c}{d}\\Rightarrow ad=bc$
                """),
            createQuestion(3, 10, """
                初中常见代数公式：<br/>
                平方差：$(a+b)(a-b)=a^2-b^2$<br/>
                完全平方：$(a+b)^2=a^2+2ab+b^2$<br/>
                完全平方：$(a-b)^2=a^2-2ab+b^2$<br/>
                一元一次方程：$2x+5=17$<br/>
                一元二次方程求根公式：$x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}$<br/>
                二次函数顶点式：$y=a(x-h)^2+k$
                """),
            createQuestion(4, 10, """
                初中几何与三角：<br/>
                勾股定理：$a^2+b^2=c^2$<br/>
                直角三角形斜边：$c=\\sqrt{a^2+b^2}$<br/>
                正弦定义：$\\sin A=\\frac{a}{c}$<br/>
                余弦定义：$\\cos A=\\frac{b}{c}$<br/>
                正切定义：$\\tan A=\\frac{a}{b}$
                """),
            createQuestion(5, 10, """
                高中常见公式：<br/>
                等差数列前 $n$ 项和：$S_n=\\frac{n(a_1+a_n)}{2}$<br/>
                等比数列前 $n$ 项和：$S_n=\\frac{a_1(1-q^n)}{1-q}$<br/>
                两角和公式：$\\sin(\\alpha+\\beta)=\\sin\\alpha\\cos\\beta+\\cos\\alpha\\sin\\beta$<br/>
                导数定义：$f'(x)=\\lim_{\\Delta x\\to0}\\frac{f(x+\\Delta x)-f(x)}{\\Delta x}$<br/>
                组合数：$C_n^m=\\frac{n!}{m!(n-m)!}$
                """),
            createQuestion(6, 15, """
                整数竖式：<br/>
                加法：<br/>
                $$\\begin{array}{rrrr}{} & {1} & {2} & {3} \\\\ {+} & {4} & {5} & {6} \\\\ \\hline {} & {5} & {7} & {9}\\end{array}$$<br/>
                减法：<br/>
                $$\\begin{array}{rrrr}{} & {8} & {6} & {4} \\\\ {-} & {2} & {7} & {9} \\\\ \\hline {} & {5} & {8} & {5}\\end{array}$$<br/>
                乘法：<br/>
                $$\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {} \\\\ \\hline {} & {5} & {5} & {3} & {5} & {}\\end{array}$$
                """),
            createQuestion(7, 15, """
                小数竖式：<br/>
                小数加法：<br/>
                $$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {+} & {3} & {.} & {75} \\\\ \\hline {} & {16} & {.} & {25}\\end{array}$$<br/>
                小数减法：<br/>
                $$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {-} & {3} & {.} & {75} \\\\ \\hline {} & {8} & {.} & {75}\\end{array}$$
                """),
            createQuestion(8, 20, """
                长除法与交叉法：<br/>
                当前长除法：<br/>
                $$\\longdiv[246]{5}{1234}$$<br/>
                浓度十字交叉：<br/>
                $$\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {20\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {10\\%} & {} & {} & {} & {20\\%}\\end{array}$$
                """)
        ));

        request.setSections(List.of(section));
        return request;
    }

    /**
     * 创建一个展示题，统一题号、分值和题型，便于后续扩展样例内容。
     */
    private QuestionDTO createQuestion(int serialNumber, int score, String content) {
        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(serialNumber);
        question.setQuestionType(6);
        question.setScore(score);
        question.setContent(content);
        return question;
    }
}
