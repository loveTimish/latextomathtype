package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class VectorWmfFormulaRendererTest {

    @Test
    void linearFormulaUsesVectorTextRecordsInsteadOfStretchDib() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("2.25\\div 0.9=2.5", 68.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(""));
        assertTrue(records.contains(0x02FB), "vector WMF should create a font");
        assertTrue(records.contains(0x0A32), "vector WMF should draw formula text with ExtTextOut");
        assertFalse(records.contains(0x0F43), "linear vector WMF must not embed a DIB bitmap");
    }

    @Test
    void limitedStructuredPlaceholdersCanRenderAsVectorText() {
        assertFalse(VectorWmfFormulaRenderer.canRender("x^{\\frac{1}{2}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\sqrt{\\frac{1}{2}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\sqrt{}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\underline{ }"));
    }

    @Test
    void explicitFlatParenFenceCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\left ( { 第+十+一+届+华+杯+赛 } \\right )", 141.0d, 19.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void squarePlaceholderCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("+2=\\square", 40.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void boxedTextCanRenderAsVectorTextAndLines() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\boxed{?????????????}", 92.0d, 18.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\boxed{?????????????}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\boxed{}\\ast (19\\ast 99)=80"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\overline{bc}\\times a=\\boxed{}5\\boxed{}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("2010\\left| \\overline{abcd}-\\overline{efgh}\\right."));
        assertTrue(VectorWmfFormulaRenderer.canRender("N=\\overline{a3}^{1}\\mathbf{\\times }\\overline{b1}^{8}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("20+\\overline{\\boxed{}\\boxed{}}+8=28+\\overline{\\boxed{}\\boxed{}}"));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleArrayCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\begin{array}{ccccc} ABCD-EFGH2008 & \\end{array}", 89.0d, 49.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{array}{c}1\\\\2\\end{array}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(\\begin{array}{cc} 4A & 8 \\end{array}\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\left.\\begin{array}{c} 15| n\\\\ 17| \\left(n+1\\right) \\end{array}"
                + "\\begin{array}{c} \\rightarrow \\\\ \\rightarrow \\end{array}\\right\\}\\Rightarrow [15,17]| \\left(2n-15\\right)"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{aligned} V&=abh \\\\ V&=Sh \\end{aligned}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} =\\left(1234+8766\\right)^{2}=10000^{2}\\\\ =100000000 \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2x&=6\\\\ x&=3\\\\ \\begin{cases} x=3\\\\ y=2 \\end{cases} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} \\text{相遇{\\blacksquare}{\\blacksquare}}\\begin{cases} ? 路程=速度和\\times "
                + "\\text{相遇{\\blacksquare}{\\blacksquare}}\\\\ 速度和=? 路程\\div "
                + "\\text{相遇{\\blacksquare}{\\blacksquare}} \\end{cases} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{cccccccc} ( & \\boxed{} & + & \\boxed{} & )\\div & \\boxed{} & = & 2\\\\ "
                + "& + & & - & & \\div & & \\\\ & \\boxed{} & - & \\boxed{} & - & \\boxed{} & = & 0\\\\ "
                + "& - & & - & & - & & \\\\ & \\boxed{} & - & \\boxed{} & \\times & \\boxed{} & = & 0\\\\ "
                + "& - & & + & & \\div & & \\\\ & \\boxed{} & + & \\boxed{} & \\div & \\boxed{} & = & 8\\\\ "
                + "& | | & & | | & & | | & & \\\\ & 1 & & 2 & & 6 & & \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{cases} \\square \\times \\square =5\\square \\\\ \\boxed{12}+\\square =\\square +\\square \\end{cases}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{cases} a=4\\left(x-a\\right)+2\\\\ a+8=6\\left(x-a-8\\right) \\end{cases} \\Rightarrow x=147"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 21\\xrightarrow{红}42\\xrightarrow{红}84\\xrightarrow{红}168\\xrightarrow{红}336"
                + "\\xrightarrow{黄}33\\xrightarrow{黄}3\\\\ "
                + "21\\xrightarrow{红}42\\xrightarrow{红}84\\xrightarrow{红}168\\xrightarrow{黄}16"
                + "\\xrightarrow{红}32\\xrightarrow{黄}3\\\\ "
                + "21\\xrightarrow{红}42\\xrightarrow{黄}4\\xrightarrow{红}8\\xrightarrow{红}16"
                + "\\xrightarrow{红}32\\xrightarrow{黄}3 \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "21\\xrightarrow{黄}2\\xrightarrow{红}4\\xrightarrow{红}8\\xrightarrow{红}16"
                + "\\xrightarrow{红}32\\xrightarrow{黄}3"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\underrightarrow{\\text{A}}\\begin{cases} 6\\\\ 14\\\\ 4 \\end{cases}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "8\\begin{array}{c} \\end{array}8\\begin{array}{c} \\end{array}8=1000"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2\\bottom left{12}\\\\ \\begin{array}{l} \\end{array}2\\bottom left{6}\\\\ \\begin{array}{l} 3 \\end{array} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\xcancel{\\begin{array}{ccc} & & \\\\ & & \\end{array}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2\\bottom left{12}\\\\ \\begin{array}{l} \\end{array}2\\bottom left{6}\\\\ "
                + "\\begin{array}{l} \\end{array}\\begin{array}{l} \\end{array}\\begin{array}{l} 3 \\end{array} \\end{array}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 2\\bottom left{\\begin{array}{ll} 18 & 12 \\end{array}}\\\\ "
                + "\\begin{array}{l} \\end{array}3\\bottom left{\\begin{array}{ll} 9 & 6 \\end{array}}\\\\ "
                + "\\begin{array}{l} \\end{array}\\begin{array}{l} \\end{array}\\begin{array}{ll} 3 & 2 \\end{array} \\end{array}"
        ));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void textCommandsCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\vartriangle =\\mathrm{9}+\\cdots", 70.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\mathrm{9}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\boldsymbol{\\pi}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\square +\\Circle +\\bigtriangleup +\\mathrm{\\whitestar }="));
        assertTrue(VectorWmfFormulaRenderer.canRender("M\\ast N"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a\\& b=a+b\\div 10"));
        assertTrue(VectorWmfFormulaRenderer.canRender("10.5\\%"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\{y\\}+y=20.09"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left\\{a\\right\\}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left.25\\right| \\left(n+7\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(3x-2)\\colon (2x+3)=4\\colon 7"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x\\Delta y=\\frac{6\\times x\\times y}{x+2y}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\vartriangle"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\bigtriangleup}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\bigtriangledown}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\uppi =3.14"));
        assertTrue(VectorWmfFormulaRenderer.canRender("1\\sim 9"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\backsim}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\nmid"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a\\in \\left(10,20\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x\\notin A\\cup B"));
        assertTrue(VectorWmfFormulaRenderer.canRender("A\\subseteq B\\cap C"));
        assertTrue(VectorWmfFormulaRenderer.canRender("CN\\parallel AF"));
        assertTrue(VectorWmfFormulaRenderer.canRender("BO\\bot AE"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\angle}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\llcorner}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\equiv}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(1\\prec x\\prec \\left(n-1\\right)\\right)"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\therefore}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\textbf{\\textcolor{maroon}{=20,}}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("0.\\dot{1}+0.125+0.\\dot{3}+0.1\\dot{6}\\approx 0.736"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\bigcirc"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Circle}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\CIRCLE}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Sun}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\oplus}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\odot}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\rightarrow}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\downarrow"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\uparrow"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\blacksquare}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Diamond}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\Diamondblack}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{\\bigstar}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(1+2+3+\\cdots +9)\\div 3=15"));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleScriptsCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("C\\times D=kD ^ { 2 }", 92.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("a_{ 1 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("p_{\\max }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("C\\times D=kD ^ { 2 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(n-1\\right)^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(a\\pm b\\right)^{2}=a^{2}\\pm 2ab+b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a\\nabla n=a^{n}+a^{n-1}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=(V_{\\mathrm{顺}}+V_{\\mathrm{逆}})\\times 10\\min"));
        assertTrue(VectorWmfFormulaRenderer.canRender("{a_{1}}^{，}=1"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x=2^{m},\\,y=2^{n},"));
        assertTrue(VectorWmfFormulaRenderer.canRender("=\\frac{1}{8}S_{\\euro{}ABCD}=\\frac{1}{8}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("^{\\prime}"));
        assertTrue(records.contains(0x02FB));
        assertTrue(records.contains(0x012D));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleFractionsCanRenderAsVectorTextAndLines() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\frac { 1 } { 15 }", 24.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac { 1 } { 15 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{}{}\\:"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\left ( { 1+2+\\cdots +9+a+b+c } \\right )\\div 3=15+\\frac { a+b+c } { 3 }"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\frac{1}{9}\\times \\overline{\\underset{1997\\mathrm{个}4}{\\underbrace{444\\cdots 4} }3"
                + "\\underset{1997\\mathrm{个}5}{\\underbrace{555\\cdots 5} }6}"
        ));
        assertTrue(records.contains(0x02FA), "fraction vector WMF should create a pen");
        assertTrue(records.contains(0x0325), "fraction vector WMF should draw a fraction bar");
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void longLinearFractionSeriesCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\left(\\frac{1}{2}+\\frac{1}{3}+\\frac{1}{4}+\\cdots +\\frac{1}{20}\\right)"
            + "+\\left(\\frac{2}{3}+\\frac{2}{4}+\\frac{2}{5}+\\cdots +\\frac{2}{20}\\right)"
            + "+\\frac{19}{20}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 316.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} 5\\mathrm{个篮球}\\,\\,\\,3\\mathrm{个排球}\\,\\,\\,318\\mathrm{元}\\\\ "
                + "\\underline{\\times 2}\\\\ 10\\mathrm{个篮球}\\,\\,6\\mathrm{个排球}\\,\\,\\,636\\mathrm{元} \\end{array}"
        ));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void productFractionSeriesWithLdotsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "2008\\times \\frac{1}{2}\\times \\frac{2}{3}\\times \\frac{3}{4}"
            + "\\times \\ldots \\ldots \\times \\frac{1999}{2000}=\\frac{2008}{2000}=\\frac{251}{250}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 240.0d, 30.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void fractionsWithSimpleScriptsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac { a_{ 1 } +a_{ 2 } +a_{ 3 } } { 3 }";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 54.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender("a=\\frac { k ^ { 2 } } { b-k }+k"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "=\\frac{1}{4}\\times 20^{2}\\times 21^{2}-8\\times \\frac{1}{4}\\times 10^{2}\\times 11^{2}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "a_{n}=\\frac{\\left(n+1\\right)^{2}}{\\left(n+1+1\\right)\\left(n+1-1\\right)}=\\frac{\\left(n+1\\right)^{2}}{n\\left(n+2\\right)}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{\\angle B+\\angle C}{360{^{\\circ}}}\\times \\uppi \\times 2^{2}=4"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac{\\mathrm{P}_{n}^{n}}{n}=\\mathrm{P}_{n-1}^{n-1}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\tan \\angle 1=\\frac{DA}{DC}=\\frac{1}{3}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("P_{3}^{1}\\spot P_{5}^{1}"));
        assertTrue(records.contains(0x02FA));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void nestedMixedFractionsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac{3\\frac{3}{4}\\times 0.2}{1.38}\\times 5.84";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 92.0d, 32.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\frac{\\frac{7}{18}\\times 4.5+0.1\\dot{6}}{13\\frac{1}{3}-3.75\\times 3.2}\\times \\left(\\frac{1}{3}+\\frac{1}{15}+\\frac{1}{35}+\\frac{1}{63}\\right)="
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\frac{\\left(0.\\overset{\\cdot }{6}\\Uptheta \\frac{15}{23}\\right)+\\left(0.625\\Delta \\frac{23}{35}\\right)}{\\left(0.\\overset{\\cdot }{3}\\Delta \\frac{34}{99}\\right)+\\left(\\frac{11}{6}\\Uptheta 2.25\\right)}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "3-\\frac{2}{3-\\frac{2}{\\begin{array}{l} \\enspace \\enspace \\vdots \\\\ 3-\\frac{2}{3} \\end{array}}}"
        ));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void annotatedBraceFractionsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac{1}{19}+\\frac{202}{1919}+\\frac{30303}{191919}"
            + "+\\frac{\\overset{8\\mathrm{个}90}{\\overbrace{90\\cdots 90} }9}"
            + "{\\underset{9\\mathrm{个}19}{\\underbrace{19\\cdots 19} }}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 188.0d, 34.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void leftBraceArraysCanRenderAsVectorText() throws IOException {
        String latex = "\\left \\{ \\begin{array}{l}B=2 \\\\,s=14\\end{array} \\right.";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 42.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{cases} b=1\\\\ c=1 \\end{cases}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{cases} x=1\\\\ y=3 \\end{cases} ,\\begin{cases} x=6\\\\ y=1 \\end{cases}"
        ));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\begin{array}{l} \\begin{cases} \\frac{x}{y}=\\frac{7}{3}\\\\ "
                + "\\frac{x+70}{y+70}=\\frac{7}{4} \\end{cases} \\\\ \\end{array}"
        ));
        assertTrue(records.contains(0x02FB));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void emptyAndParenArraysCanRenderAsVectorText() throws IOException {
        String empty = "\\begin{array}{cccc} {} & \\end{array}";
        String paren = "\\left ( { \\begin{array}{cc} {} & 4, \\\\,8 & \\end{array} } \\right )";

        assertTrue(VectorWmfFormulaRenderer.canRender(empty));
        assertTrue(VectorWmfFormulaRenderer.canRender(paren));
        assertFalse(records(VectorWmfFormulaRenderer.render(empty, 18.0d, 13.0d)).contains(0x0F43));
        assertFalse(records(VectorWmfFormulaRenderer.render(paren, 34.0d, 33.0d)).contains(0x0F43));
    }

    @Test
    void adjacentArraysCanRenderAsVectorText() throws IOException {
        String adjacent = "\\begin{array}{l}cba \\\\,\\times abc\\end{array}\\begin{array}{l}c \\\\,b \\\\,b \\\\,a\\end{array}";
        String withText = "\\begin{array}{l}1b5 \\\\,\\times 5b1\\end{array}1b505\\begin{array}{l}1 \\\\,b \\\\,b \\\\,5\\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(adjacent));
        assertTrue(VectorWmfFormulaRenderer.canRender(withText));
        assertFalse(records(VectorWmfFormulaRenderer.render(adjacent, 60.0d, 50.0d)).contains(0x0F43));
        assertFalse(records(VectorWmfFormulaRenderer.render(withText, 92.0d, 50.0d)).contains(0x0F43));
    }

    @Test
    void widePuzzleArraysCanRenderAsVectorText() throws IOException {
        String latex = "\\begin{array}{cccccccc} ( & \\mathrm{7} & + & \\mathrm{9} & )\\div & \\mathrm{8} & = & \\mathrm{2} \\\\,{} & + & {} & - & {} & \\div & {} & {} \\\\,{} & \\mathrm{1}\\mathrm{1} & - & \\mathrm{1}\\mathrm{0} & - & \\mathrm{1} & = & \\mathrm{0} \\\\,{} & - & {} & - & {} & - & {} & {} \\\\,{} & \\mathrm{1}\\mathrm{2} & - & \\mathrm{3} & \\times & \\mathrm{4} & = & \\mathrm{0} \\\\,{} & - & {} & + & {} & \\div & {} & {} \\\\,{} & \\mathrm{5} & + & \\mathrm{6} & \\div & \\mathrm{2} & = & \\mathrm{8} \\\\,{} & \\|\\| & {} & \\|\\| & {} & \\|\\| & {} & {} \\\\,{} & \\mathrm{1} & {} & \\mathrm{2} & {} & \\mathrm{6} & {} & {} \\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertFalse(records(VectorWmfFormulaRenderer.render(latex, 150.0d, 110.0d)).contains(0x0F43));
    }

    @Test
    void hlineArraysCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\begin{array}{cccccccccccccccccc} & \\mathrm{和} & = & 1 & + & 2 & + & 3 & + & 4 & + & \\cdots & + & 98 & + & 99 & + & 100\\\\ + & \\mathrm{和} & = & 100 & + & 99 & + & 98 & + & 97 & + & \\cdots & + & 3 & + & 2 & + & 1\\\\ \\hline & 2\\mathrm{倍和} & = & 101 & + & 101 & + & 101 & + & 101 & + & \\cdots & + & 101 & + & 101 & + & 101 \\end{array}";
        List<Integer> records = records(VectorWmfFormulaRenderer.render(latex, 260.0d, 42.0d));

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void overarcCanRenderAsVectorTextAndArcLines() throws IOException {
        String latex = "\\overarc{ACD}=\\overarc{AC}\\times 3=240";
        List<Integer> records = records(VectorWmfFormulaRenderer.render(latex, 120.0d, 18.0d));

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\arc{AB}=60"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\wideparen{ABC}=120"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\overset{\\frown }{\\mathrm{AEB}}"));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void nestedArrayCellsCanRenderAsVectorText() throws IOException {
        String latex = "\\begin{array}{ccccc} {} & {} & 11 & {} & {} \\\\,{} & 17 & {} & 13 & {} \\\\,23 & {} & 19 & {} & 15 \\\\,{} & 25 & {} & 21 & {} \\\\,{} & {} & 27 & {} & {} \\\\,\\to & \\begin{array}{ccccc} {} & \\end{array} & {} & {} & 27 \\\\,{} & {} & {} & 17 & {} \\\\,13 & {} & 23 & {} & 19 \\\\,{} & 15 & {} & 25 & {} \\\\,21 & {} & {} & {} & 11 \\\\,{} & {} & \\to & \\begin{array}{ccccc} {} & \\end{array} & {} \\\\,{} & 27 & {} & {} & {} \\\\,17 & {} & 13 & {} & 15 \\\\,{} & 19 & {} & 23 & {} \\\\,25 & {} & 21 & {} & {} \\\\,{} & 11 & {} & {} & \\to \\\\,\\begin{array}{ccc} 172713151923251121 & \\end{array} & \\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertFalse(records(VectorWmfFormulaRenderer.render(latex, 168.0d, 205.0d)).contains(0x0F43));
    }

    @Test
    void standaloneScriptsAndEscapedUnderscoresCanRenderAsVectorText() throws IOException {
        byte[] script = VectorWmfFormulaRenderer.render("^ { \\mathrm{3} }", 12.0d, 13.0d);
        byte[] underline = VectorWmfFormulaRenderer.render("EF=\\_\\_\\_\\_\\_", 56.0d, 13.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("^ { \\mathrm{3} }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("EF=\\_\\_\\_\\_\\_"));
        assertTrue(records(script).contains(0x0A32));
        assertTrue(records(underline).contains(0x0A32));
        assertFalse(records(script).contains(0x0F43));
        assertFalse(records(underline).contains(0x0F43));
    }

    @Test
    void pairedScriptsStayInsidePreviewWidth() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("x_i^j", 18.0d, 13.0d);
        byte[] superscript = VectorWmfFormulaRenderer.render("a^2", 13.0d, 13.0d);
        byte[] subscript = VectorWmfFormulaRenderer.render("x_i", 13.0d, 13.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("x_i^j"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^2"));
        assertTrue(VectorWmfFormulaRenderer.canRender("x_i"));
        assertFalse(records(wmf).contains(0x0F43));
        assertFalse(records(superscript).contains(0x0F43));
        assertFalse(records(subscript).contains(0x0F43));
        assertTrue(maxRecordCoordinate(wmf) <= 18.0d * 20.0d);
        assertTrue(maxRecordCoordinate(superscript) <= 13.0d * 20.0d);
        assertTrue(maxRecordCoordinate(subscript) <= 13.0d * 20.0d);
    }

    @Test
    void onlySimpleShortScriptsUseReducedScriptFontHeight() throws IOException {
        byte[] shortScript = VectorWmfFormulaRenderer.render("a^2", 13.0d, 13.0d);
        byte[] widerScript = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);

        assertTrue(createFontHeights(shortScript).get(1) > -160);
        assertEquals(-163, createFontHeights(widerScript).get(1));
        assertTrue(maxRecordCoordinate(shortScript) <= 13.0d * 20.0d);
        assertTrue(maxRecordCoordinate(widerScript) <= 36.0d * 20.0d);
    }

    @Test
    void simpleShortScriptsUseCompactHorizontalAdvance() throws IOException {
        byte[] shortScript = VectorWmfFormulaRenderer.render("a^2", 13.0d, 13.0d);
        byte[] widerScript = VectorWmfFormulaRenderer.render("S=\\left(a+b\\right)^2=9", 76.0d, 18.75d);

        assertTrue(createFontWidths(shortScript).get(1) > 0);
        assertEquals(0, createFontWidths(widerScript).get(1));
        assertTrue(totalTextDx(shortScript) < totalTextDx(widerScript));
        assertTrue(firstTextDxTotal(shortScript) < firstTextDxTotal(widerScript));
        assertTrue(maxRecordCoordinate(shortScript) <= 13.0d * 20.0d);
        assertTrue(maxRecordCoordinate(widerScript) <= 76.0d * 20.0d);
    }

    @Test
    void standaloneShortScriptsCompactWholeRunWithoutChangingRatioFormulas() throws IOException {
        byte[] upperSubscript = VectorWmfFormulaRenderer.render("S_{1}", 12.0d, 15.75d);
        byte[] lowerSuperscript = VectorWmfFormulaRenderer.render("a^{2}", 12.75d, 15.0d);
        byte[] ratioFormula = VectorWmfFormulaRenderer.render("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}", 64.0d, 17.25d);
        byte[] rawColonFormula = VectorWmfFormulaRenderer.render("S_{1}:S_{3}=a^{2}:b^{2}", 64.0d, 17.25d);
        byte[] longChainFormula = VectorWmfFormulaRenderer.render(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab",
            131.25d, 17.25d);
        byte[] triangleChainFormula = VectorWmfFormulaRenderer.render(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35",
            131.25d, 17.25d);
        byte[] leftRightEquation = VectorWmfFormulaRenderer.render("S=\\left(a+b\\right)^2=9", 76.0d, 18.75d);
        byte[] repeatedLeftRightEquation = VectorWmfFormulaRenderer.render(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9", 111.0d, 18.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("S_{1}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^{2}"));
        assertTrue(VectorWmfFormulaRenderer.isStandaloneUpperSubscript("S_{1}"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneUpperSubscript("a^{2}"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneUpperSubscript("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S_{1}:S_{3}=a^{2}:b^{2}"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^{-1}b"));
        assertTrue(VectorWmfFormulaRenderer.canRender("a^{-1}+b"));
        assertTrue(maxTextRightCoordinate(upperSubscript) < 12.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(lowerSuperscript) < 12.75d * 20.0d);
        assertEquals(maxTextRightCoordinate(rawColonFormula), maxTextRightCoordinate(ratioFormula));
        assertEquals(totalTextDx(rawColonFormula), totalTextDx(ratioFormula));
        assertFalse(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("a^{-1}b"));
        assertFalse(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("a^{2}(b-c)"));
        assertTrue(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("a^{-1}+b"));
        assertTrue(VectorWmfFormulaRenderer.hasTopLevelRelationOperator("S_{1}:S_{3}=a^{2}:b^{2}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationWidthScale("a^{-1}b"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationWidthScale("a^{2}(b-c)"));
        assertEquals(0.975d, VectorWmfFormulaRenderer.scriptRelationWidthScale("a^{2}+b"));
        assertEquals(3, VectorWmfFormulaRenderer.topLevelRelationOperatorCount("S_{1}:S_{3}=a^{2}:b^{2}"));
        assertEquals(7, VectorWmfFormulaRenderer.topLevelRelationOperatorCount(
            "S_{1}:S_{3}:S_{2}:S_{4}=a^{2}:b^{2}:ab:ab"));
        assertEquals(0.970d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertEquals(0.970d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{1}:S_{3}=a^{2}:b^{2}"));
        assertEquals(0.9895d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertEquals(0.98d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertEquals(0.98d, VectorWmfFormulaRenderer.scriptRelationWidthScale(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}\\colon S_{\\bigtriangleup COD}"
                + "\\colon S_{\\bigtriangleup DOA}=a^{2}\\colon ab\\colon b^{2}\\colon ac"));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertEquals(0.887d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{2}=2"));
        assertEquals(0.975d, VectorWmfFormulaRenderer.scriptRelationWidthScale("S_{1}=a^{2}=1"));
        assertEquals(0.975d, VectorWmfFormulaRenderer.scriptRelationWidthScale("S_{3}=4=b^{2}"));
        assertEquals(0.887d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{1}=a^{2}=1"));
        assertEquals(0.890d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{3}=4=b^{2}"));
        assertEquals(0.93d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{2}=2=a\\times b"));
        assertEquals(0.93d, VectorWmfFormulaRenderer.scriptRelationFontYScale(
            "S=S_{1}+S_{2}+S_{3}+S_{4}=1+2+4+2=9"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("a^{-1}b"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("a^{2}+b"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("a^{2}-b"));
        assertEquals(1.107d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation("S_{1}=a^{2}=1"));
        assertEquals(1.025d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S_{\\bigtriangleup AOB}\\colon S_{\\bigtriangleup BOC}=a^{2}\\colon ab=25\\colon 35"));
        assertEquals(1.022d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S_{1}\\colon S_{3}\\colon S_{2}\\colon S_{4}=a^{2}\\colon b^{2}\\colon ab\\colon ab"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S=S_{1}+S_{2}+S_{3}+S_{4}=1+2+4+2=9"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation("S_{3}=4=b^{2}"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationHeightWidthCompensation("S_{2}=2=a\\times b"));
        assertFalse(VectorWmfFormulaRenderer.repeatedEquationParenPower("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.repeatedEquationParenPower(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9"));
        assertTrue(maxTextRightCoordinate(ratioFormula) < 62.5d * 20.0d);
        assertTrue(maxTextRightCoordinate(longChainFormula) > 126.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(triangleChainFormula) < 129.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(ratioFormula) > maxTextRightCoordinate(upperSubscript) * 5);
        assertTrue(totalTextDx(ratioFormula) > totalTextDx(upperSubscript) * 6);
        assertTrue(maxTextRightCoordinate(leftRightEquation) > maxTextRightCoordinate(upperSubscript) * 6);
        assertTrue(maxTextRightCoordinate(repeatedLeftRightEquation) < 2200);
        assertTrue(totalTextDx(repeatedLeftRightEquation) < totalTextDx(leftRightEquation) * 1.455d);
        assertTrue(totalTextDx(repeatedLeftRightEquation) < 2200);
        assertTrue(maxRecordCoordinate(upperSubscript) <= 12.0d * 20.0d);
        assertTrue(maxRecordCoordinate(lowerSuperscript) <= 12.75d * 20.0d);
        assertTrue(maxRecordCoordinate(ratioFormula) <= 64.0d * 20.0d);
        assertTrue(maxRecordCoordinate(rawColonFormula) <= 64.0d * 20.0d);
        assertTrue(maxRecordCoordinate(longChainFormula) <= 131.25d * 20.0d);
        assertTrue(maxRecordCoordinate(triangleChainFormula) <= 131.25d * 20.0d);
        assertTrue(maxRecordCoordinate(leftRightEquation) <= 76.0d * 20.0d);
    }

    @Test
    void shortGeometryLabelsUseCompactHorizontalAdvance() throws IOException {
        byte[] geometry = VectorWmfFormulaRenderer.render("ABCD", 36.0d, 12.75d);
        byte[] mixed = VectorWmfFormulaRenderer.render("ABCD1", 36.0d, 12.75d);

        int geometryDx = firstTextDxTotal(geometry);
        int mixedDx = firstTextDxTotal(mixed);

        assertTrue(VectorWmfFormulaRenderer.canRender("ABCD"));
        assertTrue(VectorWmfFormulaRenderer.canRender("ABCD1"));
        assertTrue(geometryDx > 0);
        assertTrue(mixedDx > 0);
        assertTrue(geometryDx < mixedDx);
        assertTrue(maxRecordCoordinate(geometry) <= 36.0d * 20.0d);
    }

    @Test
    void standaloneTwoDigitObjectsUseCompactHorizontalAdvance() throws IOException {
        byte[] twoDigits = VectorWmfFormulaRenderer.render("25", 14.25d, 12.75d);
        byte[] anotherTwoDigits = VectorWmfFormulaRenderer.render("35", 14.25d, 12.75d);
        byte[] threeDigits = VectorWmfFormulaRenderer.render("250", 22.0d, 12.75d);
        byte[] embeddedDigits = VectorWmfFormulaRenderer.render("S=25+35", 64.0d, 12.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("25"));
        assertTrue(VectorWmfFormulaRenderer.canRender("35"));
        assertTrue(VectorWmfFormulaRenderer.canRender("250"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=25+35"));
        assertEquals(firstTextAverageDx(twoDigits), firstTextAverageDx(anotherTwoDigits));
        assertTrue(firstTextAverageDx(twoDigits) < firstTextAverageDx(threeDigits));
        assertTrue(firstTextAverageDx(embeddedDigits) > firstTextAverageDx(twoDigits));
        assertTrue(maxRecordCoordinate(twoDigits) <= 14.25d * 20.0d);
        assertTrue(maxRecordCoordinate(anotherTwoDigits) <= 14.25d * 20.0d);
        assertTrue(maxRecordCoordinate(threeDigits) <= 22.0d * 20.0d);
        assertTrue(maxRecordCoordinate(embeddedDigits) <= 64.0d * 20.0d);
    }

    @Test
    void standaloneParenthesizedPowersUseCompactHorizontalAdvance() throws IOException {
        byte[] standalone = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);
        byte[] metricsStandalone = VectorWmfFormulaRenderer.render(
            "\\pwmetrics{36.000,19.000,36.000,18.750}\\left(a+b\\right)^{2}", 36.0d, 18.75d);
        byte[] plainParen = VectorWmfFormulaRenderer.render("(a+b)^2", 36.0d, 18.75d);
        byte[] longEquation = VectorWmfFormulaRenderer.render("S=\\left(a+b\\right)^2=9", 76.0d, 18.75d);
        byte[] repeatedEquation = VectorWmfFormulaRenderer.render(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9", 111.0d, 18.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(a+b\\right)^2"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(a+b)^2"));
        assertTrue(VectorWmfFormulaRenderer.canRender("S=\\left(a+b\\right)^2=9"));
        assertEquals(0, createFontWidths(standalone).get(0));
        assertEquals(0, createFontWidths(standalone).get(1));
        assertEquals(0, createFontWidths(metricsStandalone).get(0));
        assertEquals(0, createFontWidths(metricsStandalone).get(1));
        assertEquals(0, createFontWidths(plainParen).get(0));
        assertEquals(0, createFontWidths(plainParen).get(1));
        assertEquals(0, createFontWidths(longEquation).get(0));
        assertEquals(0, createFontWidths(longEquation).get(1));
        assertEquals(totalTextDx(standalone), totalTextDx(metricsStandalone));
        assertTrue(maxTextRightCoordinate(standalone) < maxTextRightCoordinate(plainParen));
        assertEquals(maxTextRightCoordinate(standalone), maxTextRightCoordinate(metricsStandalone));
        assertFalse(VectorWmfFormulaRenderer.repeatedEquationParenPower("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.repeatedEquationParenPower(
            "S=\\left(a+b\\right)^2=\\left(1+2\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("\\left(a+b\\right)^2"));
        assertTrue(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("S=\\left(a+b\\right)^2=9"));
        assertTrue(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("(a+b)^2"));
        assertFalse(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertFalse(VectorWmfFormulaRenderer.hasClosingFenceSuperscript("a^{2}+b"));
        assertTrue(maxTextRightCoordinate(repeatedEquation) < 2200);
        assertTrue(totalTextDx(repeatedEquation) < 2200);
        assertTrue(totalTextDx(standalone) < firstTextDxTotal(longEquation) * 0.8d);
        assertTrue(maxRecordCoordinate(standalone) <= 36.0d * 20.0d);
        assertTrue(maxRecordCoordinate(metricsStandalone) <= 36.0d * 20.0d);
        assertTrue(maxRecordCoordinate(plainParen) <= 36.0d * 20.0d);
        assertTrue(maxRecordCoordinate(longEquation) <= 76.0d * 20.0d);
        assertTrue(maxRecordCoordinate(repeatedEquation) <= 111.0d * 20.0d);
    }

    @Test
    void standaloneShortLabelsUseLocalWidthCorrection() throws IOException {
        byte[] singleS = VectorWmfFormulaRenderer.render("\\pwmetrics{9.983,12.977,9.750,12.750}S", 9.75d,
            12.75d);
        byte[] singleA = VectorWmfFormulaRenderer.render("A", 9.75d, 12.75d);
        byte[] scriptS = VectorWmfFormulaRenderer.render("S_{1}", 12.0d, 15.75d);
        byte[] scriptS3 = VectorWmfFormulaRenderer.render("S_{3}", 12.0d, 15.75d);
        byte[] scriptA1 = VectorWmfFormulaRenderer.render("A_{1}", 12.0d, 15.75d);
        byte[] bd = VectorWmfFormulaRenderer.render("\\pwmetrics{17.992,11.995,18.000,12.000}BD", 18.0d,
            12.0d);
        byte[] ab = VectorWmfFormulaRenderer.render("\\pwmetrics{17.992,11.995,18.000,12.000}AB", 18.0d,
            12.0d);
        byte[] abcd = VectorWmfFormulaRenderer.render("ABCD", 32.25d, 12.75d);
        byte[] twentyFive = VectorWmfFormulaRenderer.render("25", 14.25d, 12.75d);
        byte[] thirtyFive = VectorWmfFormulaRenderer.render("35", 14.25d, 12.75d);
        byte[] longNumber = VectorWmfFormulaRenderer.render("144", 20.0d, 12.75d);
        byte[] aSquared = VectorWmfFormulaRenderer.render("a^{2}", 9.75d, 15.75d);
        byte[] bSquared = VectorWmfFormulaRenderer.render("b^{2}", 9.75d, 15.75d);
        byte[] bCubed = VectorWmfFormulaRenderer.render("b^{3}", 9.75d, 15.75d);
        byte[] cSquared = VectorWmfFormulaRenderer.render("c^{2}", 9.75d, 15.75d);

        assertTrue(createFontWidths(singleS).get(0) > 0);
        assertEquals(0, createFontWidths(singleA).get(0));
        assertEquals(0, createFontWidths(scriptS).get(0));
        assertEquals(0, createFontWidths(scriptS3).get(0));
        assertTrue(firstTextDxTotal(singleS) > firstTextDxTotal(singleA));
        assertTrue(totalTextDx(scriptS) < totalTextDx(scriptA1));
        assertTrue(totalTextDx(scriptS3) > totalTextDx(scriptS));
        assertTrue(firstTextDxTotal(bd) > firstTextDxTotal(ab));
        assertTrue(firstTextDxTotal(twentyFive) < firstTextDxTotal(longNumber));
        assertEquals(firstTextDxTotal(twentyFive), firstTextDxTotal(thirtyFive));
        assertTrue(maxTextRightCoordinate(bSquared) < maxTextRightCoordinate(aSquared));
        assertTrue(maxTextRightCoordinate(bCubed) > maxTextRightCoordinate(bSquared));
        assertTrue(maxTextRightCoordinate(cSquared) > maxTextRightCoordinate(bSquared));
        assertTrue(maxRecordCoordinate(bd) <= 18.0d * 20.0d);
        assertTrue(maxRecordCoordinate(abcd) <= 32.25d * 20.0d);
        assertTrue(maxRecordCoordinate(twentyFive) <= 14.25d * 20.0d);
        assertTrue(maxRecordCoordinate(aSquared) <= 9.75d * 20.0d);
        assertTrue(maxRecordCoordinate(bSquared) <= 9.75d * 20.0d);
    }

    @Test
    void simpleLinearFormulasUseCompactFontHeightButLeftRightScriptsDoNot() throws IOException {
        byte[] simple = VectorWmfFormulaRenderer.render("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}", 64.0d, 17.25d);
        byte[] shortEquation = VectorWmfFormulaRenderer.render("S_{2}=2", 30.0d, 15.75d);
        byte[] shortSuperscriptEquation = VectorWmfFormulaRenderer.render("a^{2}=1", 30.0d, 15.75d);
        byte[] shortScript = VectorWmfFormulaRenderer.render("S_{2}", 12.0d, 15.75d);
        byte[] leftRight = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);
        byte[] geometryLabel = VectorWmfFormulaRenderer.render("\\pwmetrics{17.992,11.995,18.000,12.000}AB", 18.0d,
            12.0d);
        byte[] mixedGeometryText = VectorWmfFormulaRenderer.render("AB1", 18.0d, 12.0d);

        assertTrue(createFontHeights(simple).get(0) > -240);
        assertTrue(createFontHeights(simple).get(0) > createFontHeights(shortScript).get(0));
        assertTrue(createFontHeights(shortEquation).get(0) < createFontHeights(simple).get(0));
        assertTrue(createFontHeights(geometryLabel).get(0) < createFontHeights(mixedGeometryText).get(0));
        assertTrue(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel(
            "\\pwmetrics{17.992,11.995,18.000,12.000}AB"));
        assertTrue(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel(
            "\\pwmetrics{17.992,11.995,18.000,12.000}BD"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("AC"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("CD"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("ABCD"));
        assertFalse(VectorWmfFormulaRenderer.isStandaloneTallGeometryLabel("A1"));
        assertEquals(createFontHeights(shortEquation).get(0), createFontHeights(shortSuperscriptEquation).get(0));
        assertEquals(1.0d, VectorWmfFormulaRenderer.scriptRelationFontYScale("S_{2}=2"));
        assertTrue(VectorWmfFormulaRenderer.shortScriptEquationWidthScale("S_{2}=2") < 1.0d);
        assertTrue(VectorWmfFormulaRenderer.shortScriptEquationWidthScale("a^{2}=1") < 1.0d);
        assertEquals(0.976d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("S_{2}=2"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("S_{2}=2=a\\times b"));
        assertEquals(0.965d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("b=2"));
        assertEquals(1.025d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("a=1"));
        assertEquals(1.025d, VectorWmfFormulaRenderer.shortExactEquationWidthScale(
            "\\pwmetrics{22.980,12.989,23.250,12.750}a=1"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortExactEquationWidthScale("a\\colon b=5\\colon 7"));
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortScriptEquationWidthScale(
            "S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertTrue(maxTextRightCoordinate(shortEquation) < 30.0d * 20.0d);
        assertTrue(maxTextRightCoordinate(shortSuperscriptEquation) < 30.0d * 20.0d);
        assertEquals(1.0d, VectorWmfFormulaRenderer.shortScriptEquationWidthScale("S_{2}"));
        assertTrue(createFontHeights(simple).get(0) > createFontHeights(shortScript).get(0));
        assertEquals(-245, createFontHeights(leftRight).get(0));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("S_{2}=2"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("S^{2}=4"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("a^{2}=1"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("a_{2}=1"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("x_{1}=2"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("A^{2}=1"));
        assertTrue(VectorWmfFormulaRenderer.isShortScriptEquation("S_{2}=2.5"));
        assertFalse(VectorWmfFormulaRenderer.isShortScriptEquation("S_{1}\\colon S_{3}=a^{2}\\colon b^{2}"));
        assertFalse(VectorWmfFormulaRenderer.isShortScriptEquation("S=\\left(a+b\\right)^2=9"));
        assertFalse(VectorWmfFormulaRenderer.canRender("\\dfrac{a}{b}^2"));
        assertFalse(VectorWmfFormulaRenderer.canRender("\\cfrac{a}{b}^2"));
    }

    @Test
    void closingFenceSuperscriptsUseRaisedBaseline() throws IOException {
        byte[] leftRight = VectorWmfFormulaRenderer.render("\\left(a+b\\right)^2", 36.0d, 18.75d);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\left(a+b\\right)^2"));
        assertTrue(minTextYCoordinate(leftRight) <= 115);
        assertTrue(maxRecordCoordinate(leftRight) <= 36.0d * 20.0d);
    }

    @Test
    void symbolCommandsDoNotEncodeAsQuestionMarks() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render(
            "\\times+\\div+\\leq+\\geq+\\neq+\\pi+\\cdots+\\parallel+\\because", 180.0d, 16.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\times+\\div+\\leq+\\geq+\\neq+\\pi+\\cdots+\\parallel+\\because"));
        assertFalse(textRecords.isEmpty());
        assertFalse(textRecords.stream().anyMatch(VectorWmfFormulaRendererTest::containsQuestionMark));
    }

    @Test
    void flatAlignmentMarkersAreNotRenderedAsText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("&=8.27\\times (100+3)", 76.0d, 15.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("&=8.27\\times (100+3)"));
        assertFalse(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) '&')));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) '=')));
    }

    @Test
    void escapedAmpersandInsideArrayCellIsNotSplitAsAlignment() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\begin{array}{c}a\\&b\\end{array}", 32.0d, 15.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{array}{c}a\\&b\\end{array}"));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) '&')));
    }

    @Test
    void multiplicationAndDivisionUseSymbolGlyphBytesInWmfPreview() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("8.27\\times 100\\div 2", 92.0d, 15.0d);
        List<byte[]> textRecords = extTextOutBytes(wmf);

        assertFalse(textRecords.isEmpty());
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) 0xB4)));
        assertTrue(textRecords.stream().anyMatch(bytes -> containsByte(bytes, (byte) 0xB8)));
    }

    @Test
    void extTextOutKeepsTransparentOptionsAndCarriesDxArray() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("8.27\\times 100", 72.0d, 15.0d);

        assertTrue(hasExtTextOutDxArray(wmf));
    }

    @Test
    void cjkExtTextOutCarriesByteAlignedDxArray() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("3\\mathrm{个篮球}", 60.0d, 15.0d);

        assertTrue(hasExtTextOutDxArray(wmf));
    }

    @Test
    void scaledPreviewCoordinatesStayInsideWmfWindow() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render(
            "\\left(\\frac{1}{2}+\\frac{1}{3}+\\cdots+\\frac{1}{20}\\right)", 40.0d, 16.0d);

        assertTrue(maxRecordCoordinate(wmf) <= 40.0d * 20.0d);
    }

    @Test
    void tallArraysKeepHorizontalScaleWhenHeightIsCompressed() throws IOException {
        String latex = "\\begin{array}{ccccccccc} 1 & & & & & & & & \\\\ "
            + "2 & 3 & 4 & & & & & & \\\\ "
            + "5 & 6 & 7 & 8 & 9 & & & & \\\\ "
            + "10 & 11 & 12 & 13 & 14 & 15 & 16 & & \\\\ "
            + "17 & 18 & 19 & 20 & 21 & 22 & 23 & 24 & 25\\\\ "
            + "\\cdots & \\cdots & & & & & & & \\end{array}";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 108.0d, 14.0d);

        assertEquals((int) (108.0d * 20.0d), windowExtX(wmf));
        assertEquals((int) (14.0d * 20.0d), windowExtY(wmf));
        assertTrue(maxTextRightCoordinate(wmf) > 90.0d * 20.0d);
        assertTrue(maxRecordCoordinate(wmf) <= 108.0d * 20.0d);
        assertTrue(maxTextYCoordinate(wmf) <= 14.0d * 20.0d);
    }

    private static List<Integer> records(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            out.add(function);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static List<byte[]> extTextOutBytes(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<byte[]> out = new ArrayList<>();
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textOffset = offset + 14;
                if (count >= 0 && textOffset + count <= data.length) {
                    byte[] text = new byte[count];
                    System.arraycopy(data, textOffset, text, 0, count);
                    out.add(text);
                }
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static List<Integer> createFontHeights(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 8 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB) {
                out.add(signedWord(data, offset + 6));
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static List<Integer> createFontWidths(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 10 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x02FB) {
                out.add(signedWord(data, offset + 8));
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static boolean containsQuestionMark(byte[] bytes) {
        return containsByte(bytes, (byte) '?');
    }

    private static boolean hasExtTextOutDxArray(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int options = word(data, offset + 12);
                int textBytes = count + (count & 1);
                int payloadBytes = sizeWords * 2 - 6;
                if (options == 0 && count > 0 && payloadBytes >= 8 + textBytes + count * 2) {
                    return true;
                }
            }
            offset += sizeWords * 2;
        }
        return false;
    }

    private static boolean containsByte(byte[] bytes, byte target) {
        for (byte value : bytes) {
            if (value == target) {
                return true;
            }
        }
        return false;
    }

    private static int maxRecordCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int x = word(data, offset + 8);
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxOffset += 2;
                }
                max = Math.max(max, x + dxTotal);
            } else if (function == 0x0325 && offset + 18 <= data.length) {
                max = Math.max(max, word(data, offset + 10));
                max = Math.max(max, word(data, offset + 14));
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static int firstTextDxTotal(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxOffset += 2;
                }
                return dxTotal;
            }
            offset += sizeWords * 2;
        }
        return 0;
    }

    private static int firstTextAverageDx(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                int dxCount = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxCount++;
                    dxOffset += 2;
                }
                return dxCount == 0 ? 0 : dxTotal / dxCount;
            }
            offset += sizeWords * 2;
        }
        return 0;
    }

    private static int totalTextDx(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int total = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    total += word(data, dxOffset);
                    dxOffset += 2;
                }
            }
            offset += sizeWords * 2;
        }
        return total;
    }

    private static int maxTextRightCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                int x = word(data, offset + 8);
                int count = word(data, offset + 10);
                int textBytes = count + (count & 1);
                int dxOffset = offset + 14 + textBytes;
                int dxTotal = 0;
                while (dxOffset + 2 <= offset + sizeWords * 2) {
                    dxTotal += word(data, dxOffset);
                    dxOffset += 2;
                }
                max = Math.max(max, x + dxTotal);
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static int maxTextYCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int max = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                max = Math.max(max, word(data, offset + 6));
            }
            offset += sizeWords * 2;
        }
        return max;
    }

    private static int minTextYCoordinate(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int min = Integer.MAX_VALUE;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x0A32 && offset + 14 <= data.length) {
                min = Math.min(min, word(data, offset + 6));
            }
            offset += sizeWords * 2;
        }
        return min == Integer.MAX_VALUE ? 0 : min;
    }

    private static int windowExtX(byte[] data) {
        return windowExt(data, true);
    }

    private static int windowExtY(byte[] data) {
        return windowExt(data, false);
    }

    private static int windowExt(byte[] data, boolean x) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        while (offset + 10 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            if (function == 0x020C && offset + 10 <= data.length) {
                int yExt = word(data, offset + 6);
                int xExt = word(data, offset + 8);
                return x ? xExt : yExt;
            }
            offset += sizeWords * 2;
        }
        return -1;
    }

    private static boolean hasPlaceableHeader(byte[] data) {
        return data.length >= 22 && dword(data, 0) == 0x9AC6CDD7;
    }

    private static int word(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort() & 0xffff;
    }

    private static int signedWord(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort();
    }

    private static int dword(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
    }
}
