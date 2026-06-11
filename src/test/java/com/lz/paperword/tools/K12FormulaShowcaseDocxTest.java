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
            createQuestion(1, 5, """
                小学几何与计量：<br/>
                长方形周长：$C=2(a+b)$<br/>
                长方形面积：$S=ab$<br/>
                正方形周长：$C=4a$<br/>
                正方形面积：$S=a^2$<br/>
                三角形面积：$S=\\frac{ah}{2}$<br/>
                平行四边形面积：$S=ah$<br/>
                梯形面积：$S=\\frac{(a+b)h}{2}$<br/>
                圆周长：$C=2\\pi r$<br/>
                圆面积：$S=\\pi r^2$<br/>
                长方体体积：$V=abh$<br/>
                正方体体积：$V=a^3$<br/>
                圆柱体积：$V=\\pi r^2h$<br/>
                圆锥体积：$V=\\frac{1}{3}\\pi r^2h$
                """),
            createQuestion(2, 5, """
                小学分数、小数、百分数与比例：<br/>
                分数加法：$\\frac{3}{4}+\\frac{1}{2}=\\frac{5}{4}$<br/>
                分数乘法：$\\frac{2}{3}\\times\\frac{3}{5}=\\frac{2}{5}$<br/>
                分数除法：$\\frac{3}{4}\\div\\frac{1}{2}=\\frac{3}{2}$<br/>
                小数乘法：$1.25\\times0.4=0.5$<br/>
                百分数：$25\\%=\\frac{1}{4}=0.25$<br/>
                比例：$\\frac{a}{b}=\\frac{c}{d}\\Rightarrow ad=bc$<br/>
                速度路程时间：$s=vt$，$v=\\frac{s}{t}$，$t=\\frac{s}{v}$<br/>
                单价数量总价：$T=pn$
                """),
            createQuestion(3, 5, """
                初中数与式：<br/>
                同底数幂相乘：$a^m\\cdot a^n=a^{m+n}$<br/>
                幂的乘方：$(a^m)^n=a^{mn}$<br/>
                积的乘方：$(ab)^n=a^nb^n$<br/>
                同底数幂相除：$\\frac{a^m}{a^n}=a^{m-n}$<br/>
                负整数指数幂：$a^{-n}=\\frac{1}{a^n}$<br/>
                零指数幂：$a^0=1$<br/>
                根式性质：$\\sqrt{ab}=\\sqrt a\\sqrt b$，$\\sqrt{a^2}=|a|$
                """),
            createQuestion(4, 5, """
                初中整式、因式分解与方程：<br/>
                平方差：$(a+b)(a-b)=a^2-b^2$<br/>
                完全平方：$(a+b)^2=a^2+2ab+b^2$<br/>
                完全平方：$(a-b)^2=a^2-2ab+b^2$<br/>
                立方和：$a^3+b^3=(a+b)(a^2-ab+b^2)$<br/>
                立方差：$a^3-b^3=(a-b)(a^2+ab+b^2)$<br/>
                一元二次方程求根公式：$x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}$<br/>
                判别式：$\\Delta=b^2-4ac$<br/>
                韦达定理：$x_1+x_2=-\\frac{b}{a}$，$x_1x_2=\\frac{c}{a}$
                """),
            createQuestion(5, 5, """
                初中函数与坐标：<br/>
                一次函数：$y=kx+b$<br/>
                正比例函数：$y=kx$<br/>
                反比例函数：$y=\\frac{k}{x}$<br/>
                二次函数一般式：$y=ax^2+bx+c$<br/>
                二次函数顶点式：$y=a(x-h)^2+k$<br/>
                二次函数顶点：$\\left(-\\frac{b}{2a},\\frac{4ac-b^2}{4a}\\right)$<br/>
                对称轴：$x=-\\frac{b}{2a}$<br/>
                两点距离：$d=\\sqrt{(x_2-x_1)^2+(y_2-y_1)^2}$
                """),
            createQuestion(6, 5, """
                初中几何与锐角三角函数：<br/>
                勾股定理：$a^2+b^2=c^2$<br/>
                直角三角形斜边：$c=\\sqrt{a^2+b^2}$<br/>
                正弦定义：$\\sin A=\\frac{a}{c}$<br/>
                余弦定义：$\\cos A=\\frac{b}{c}$<br/>
                正切定义：$\\tan A=\\frac{a}{b}$<br/>
                圆心角弧长：$l=\\frac{n\\pi r}{180}$<br/>
                扇形面积：$S=\\frac{n\\pi r^2}{360}=\\frac{1}{2}lr$<br/>
                相似三角形：$\\frac{AB}{A'B'}=\\frac{BC}{B'C'}=\\frac{AC}{A'C'}$
                """),
            createQuestion(7, 5, """
                高中集合、逻辑与不等式：<br/>
                交集：$A\\cap B$<br/>
                并集：$A\\cup B$<br/>
                补集：$U-A$<br/>
                基本不等式：$a+b\\ge 2\\sqrt{ab}$<br/>
                绝对值不等式：$|x-a|<r$<br/>
                充分必要条件：$p\\Leftrightarrow q$
                """),
            createQuestion(8, 5, """
                高中函数与指数对数：<br/>
                指数运算：$a^{x+y}=a^xa^y$<br/>
                对数定义：$a^x=N\\Leftrightarrow x=\\log_a N$<br/>
                对数乘法：$\\log_a MN=\\log_a M+\\log_a N$<br/>
                对数除法：$\\log_a\\frac{M}{N}=\\log_a M-\\log_a N$<br/>
                对数换底：$\\log_a b=\\frac{\\log_c b}{\\log_c a}$<br/>
                复合函数：$(f\\circ g)(x)=f(g(x))$
                """),
            createQuestion(9, 5, """
                高中三角函数：<br/>
                同角关系：$\\sin^2\\alpha+\\cos^2\\alpha=1$<br/>
                正切关系：$\\tan\\alpha=\\frac{\\sin\\alpha}{\\cos\\alpha}$<br/>
                两角和正弦：$\\sin(\\alpha+\\beta)=\\sin\\alpha\\cos\\beta+\\cos\\alpha\\sin\\beta$<br/>
                两角和余弦：$\\cos(\\alpha+\\beta)=\\cos\\alpha\\cos\\beta-\\sin\\alpha\\sin\\beta$<br/>
                二倍角正弦：$\\sin2\\alpha=2\\sin\\alpha\\cos\\alpha$<br/>
                二倍角余弦：$\\cos2\\alpha=\\cos^2\\alpha-\\sin^2\\alpha$<br/>
                正弦定理：$\\frac{a}{\\sin A}=\\frac{b}{\\sin B}=\\frac{c}{\\sin C}=2R$<br/>
                余弦定理：$a^2=b^2+c^2-2bc\\cos A$
                """),
            createQuestion(10, 5, """
                高中数列：<br/>
                等差数列通项：$a_n=a_1+(n-1)d$<br/>
                等差数列前 $n$ 项和：$S_n=\\frac{n(a_1+a_n)}{2}$<br/>
                等差中项：$A=\\frac{a+b}{2}$<br/>
                等比数列通项：$a_n=a_1q^{n-1}$<br/>
                等比数列前 $n$ 项和：$S_n=\\frac{a_1(1-q^n)}{1-q}$<br/>
                等比中项：$G=\\sqrt{ab}$<br/>
                递推关系：$a_{n+1}=qa_n$
                """),
            createQuestion(11, 5, """
                高中概率统计：<br/>
                古典概型：$P(A)=\\frac{m}{n}$<br/>
                对立事件：$P(\\overline A)=1-P(A)$<br/>
                加法公式：$P(A\\cup B)=P(A)+P(B)-P(A\\cap B)$<br/>
                条件概率：$P(B|A)=\\frac{P(AB)}{P(A)}$<br/>
                期望：$E(X)=\\sum x_i p_i$<br/>
                方差：$D(X)=E(X^2)-[E(X)]^2$<br/>
                平均数：$\\bar{x}=\\frac{x_1+x_2+\\cdots+x_n}{n}$<br/>
                方差：$s^2=\\frac{1}{n}\\sum_{i=1}^{n}(x_i-\\bar{x})^2$
                """),
            createQuestion(12, 5, """
                高中排列组合与二项式：<br/>
                排列数：$A_n^m=\\frac{n!}{(n-m)!}$<br/>
                组合数：$C_n^m=\\frac{n!}{m!(n-m)!}$<br/>
                组合性质：$C_n^m=C_n^{n-m}$<br/>
                组合递推：$C_n^m=C_{n-1}^{m}+C_{n-1}^{m-1}$<br/>
                二项式定理：$(a+b)^n=\\sum_{k=0}^{n}C_n^k a^{n-k}b^k$<br/>
                通项：$T_{k+1}=C_n^k a^{n-k}b^k$
                """),
            createQuestion(13, 5, """
                高中导数与常见极限：<br/>
                导数定义：$f'(x)=\\lim_{\\Delta x\\to0}\\frac{f(x+\\Delta x)-f(x)}{\\Delta x}$<br/>
                幂函数导数：$(x^n)'=nx^{n-1}$<br/>
                指数导数：$(e^x)'=e^x$<br/>
                对数导数：$(\\ln x)'=\\frac{1}{x}$<br/>
                正弦导数：$(\\sin x)'=\\cos x$<br/>
                余弦导数：$(\\cos x)'=-\\sin x$<br/>
                乘法求导：$(uv)'=u'v+uv'$<br/>
                商法求导：$\\left(\\frac{u}{v}\\right)'=\\frac{u'v-uv'}{v^2}$
                """),
            createQuestion(14, 5, """
                高中解析几何：<br/>
                直线斜率：$k=\\frac{y_2-y_1}{x_2-x_1}$<br/>
                点斜式：$y-y_0=k(x-x_0)$<br/>
                点到直线距离：$d=\\frac{|Ax_0+By_0+C|}{\\sqrt{A^2+B^2}}$<br/>
                圆标准方程：$(x-a)^2+(y-b)^2=r^2$<br/>
                椭圆标准方程：$\\frac{x^2}{a^2}+\\frac{y^2}{b^2}=1$<br/>
                双曲线标准方程：$\\frac{x^2}{a^2}-\\frac{y^2}{b^2}=1$<br/>
                抛物线标准方程：$y^2=2px$
                """),
            createQuestion(15, 5, """
                高中立体几何与向量：<br/>
                向量数量积：$\\vec a\\cdot\\vec b=|\\vec a||\\vec b|\\cos\\theta$<br/>
                向量夹角：$\\cos\\theta=\\frac{\\vec a\\cdot\\vec b}{|\\vec a||\\vec b|}$<br/>
                空间两点距离：$d=\\sqrt{(x_2-x_1)^2+(y_2-y_1)^2+(z_2-z_1)^2}$<br/>
                球表面积：$S=4\\pi R^2$<br/>
                球体积：$V=\\frac{4}{3}\\pi R^3$<br/>
                棱柱体积：$V=Sh$<br/>
                棱锥体积：$V=\\frac{1}{3}Sh$
                """),
            createQuestion(16, 5, """
                物理常见公式：<br/>
                牛顿第二定律：$F=ma$<br/>
                重力：$G=mg$<br/>
                压强：$p=\\frac{F}{S}$<br/>
                功：$W=Fs$<br/>
                功率：$P=\\frac{W}{t}$<br/>
                动能：$E_k=\\frac{1}{2}mv^2$<br/>
                欧姆定律：$I=\\frac{U}{R}$<br/>
                电功率：$P=UI=I^2R=\\frac{U^2}{R}$
                """),
            createQuestion(17, 5, """
                化学与科学计量常见公式：<br/>
                物质的量：$n=\\frac{m}{M}$<br/>
                质量分数：$w=\\frac{m_s}{m_l}\\times100\\%$<br/>
                溶质质量：$m_s=m_lw$<br/>
                稀释前后：$c_1V_1=c_2V_2$<br/>
                密度：$\\rho=\\frac{m}{V}$<br/>
                气体状态方程：$pV=nRT$
                """),
            createQuestion(18, 5, """
                整数竖式：<br/>
                加法：<br/>
                $$\\begin{array}{rrrr}{} & {1} & {2} & {3} \\\\ {+} & {4} & {5} & {6} \\\\ \\hline {} & {5} & {7} & {9}\\end{array}$$<br/>
                减法：<br/>
                $$\\begin{array}{rrrr}{} & {8} & {6} & {4} \\\\ {-} & {2} & {7} & {9} \\\\ \\hline {} & {5} & {8} & {5}\\end{array}$$<br/>
                乘法：<br/>
                $$\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {} \\\\ \\hline {} & {5} & {5} & {3} & {5} & {}\\end{array}$$
                """),
            createQuestion(19, 5, """
                小数竖式：<br/>
                小数加法：<br/>
                $$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {+} & {3} & {.} & {75} \\\\ \\hline {} & {16} & {.} & {25}\\end{array}$$<br/>
                小数减法：<br/>
                $$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {-} & {3} & {.} & {75} \\\\ \\hline {} & {8} & {.} & {75}\\end{array}$$
                """),
            createQuestion(20, 5, """
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
