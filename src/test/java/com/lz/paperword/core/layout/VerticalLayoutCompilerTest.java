package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class VerticalLayoutCompilerTest {

    private final LaTeXParser parser = new LaTeXParser();
    private final VerticalLayoutCompiler compiler = new VerticalLayoutCompiler();

    @Test
    void shouldCompileAdditionArrayWithRuleSpan() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertEquals(VerticalLayoutSpec.Kind.ARITHMETIC, spec.kind());
        assertEquals(4, spec.columnCount());
        assertEquals(1, spec.ruleSpans().size());
        assertEquals(2, spec.ruleSpans().get(0).boundaryIndex());
        assertEquals(0, spec.ruleSpans().get(0).startColumn());
        assertEquals(3, spec.ruleSpans().get(0).endColumn());
    }

    @Test
    void shouldCompileDecimalArrayIntoFixedColumns() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {+} & {3} & {.} & {75} \\\\ \\hline {} & {16} & {.} & {25}\\end{array}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertEquals(VerticalLayoutSpec.Kind.DECIMAL, spec.kind());
        assertEquals(1, spec.columnCount());
        assertEquals("12.50", spec.rows().get(0).cells().get(0));
        assertEquals("+3.75", spec.rows().get(1).cells().get(0));
        assertEquals("16.25", spec.rows().get(2).cells().get(0));
    }

    @Test
    void shouldKeepExplicitLongDivisionHeaderWithoutSynthesizingSteps() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.isLongDivision());
        assertEquals("13", spec.longDivisionHeader().divisor());
        assertEquals("65", spec.longDivisionHeader().quotient());
        assertEquals("845", spec.longDivisionHeader().dividend());
        assertTrue(spec.longDivisionSteps().isEmpty(), "长除法步骤区应改由 LLM 原始结果提供，不再本地补算");
    }

    @Test
    void shouldComputeSemanticLongDivisionLinePlacementFromHeader() {
        VerticalLayoutSpec.LongDivisionHeader header =
            new VerticalLayoutSpec.LongDivisionHeader("7", "14300", "100100");

        java.util.List<VerticalLayoutSpec.LongDivisionPlacedLine> placedLines =
            compiler.computeExpectedLongDivisionLines(header);

        assertEquals(8, placedLines.size(), "整数长除法应能从头部推导完整的列位序列");
        assertEquals(1, placedLines.get(0).startColumn(), "首步单数字乘积应在当前被除数块内右对齐");
        assertEquals(1, placedLines.get(0).underlineStartColumn(), "首步下划线应与单数字乘积的右对齐起点一致");
        assertEquals(1, placedLines.get(0).underlineEndColumn(), "首步下划线应覆盖当前被除数块宽度");
        assertEquals(1, placedLines.get(1).startColumn(), "下移后的 30 应从上一轮余数的尾列继续展开");
        assertEquals(5, placedLines.get(7).startColumn(), "最终余数 0 应落在最后一位被下移数字所在列");
    }

    @Test
    void shouldKeepMultiplicationColumnCountFromExplicitPlaceholders() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {} \\\\ \\hline {} & {5} & {5} & {3} & {5} & {}\\end{array}"
        );
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertEquals(VerticalLayoutSpec.Kind.ARITHMETIC, spec.kind());
        assertEquals(6, spec.columnCount());
        assertEquals(2, spec.ruleSpans().size());
        assertEquals("×", spec.rows().get(1).cells().get(0));
        assertEquals("4", spec.rows().get(1).cells().get(3));
        assertEquals("5", spec.rows().get(1).cells().get(4));
    }

    @Test
    void shouldKeepBareLongDivisionHeaderWithoutStructuredSteps() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertEquals("5", spec.longDivisionHeader().divisor());
        assertEquals("246", spec.longDivisionHeader().quotient());
        assertEquals("1234", spec.longDivisionHeader().dividend());
        assertFalse(spec.hasStructuredLongDivisionSteps(), "bare longdiv 不应再触发本地步骤推导");
    }

    @Test
    void shouldKeepBareLongDivisionHeaderForIntegerDivision() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{7}{100100}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertEquals("7", spec.longDivisionHeader().divisor());
        assertEquals("", spec.longDivisionHeader().quotient());
        assertEquals("100100", spec.longDivisionHeader().dividend());
        assertTrue(spec.longDivisionSteps().isEmpty(), "整数除法不应仅凭头部自动补出步骤");
    }

    @Test
    void shouldKeepRawDecimalLongDivisionHeader() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{2.5}{12.5}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.isLongDivision());
        assertEquals("2.5", spec.longDivisionHeader().divisor(), "小数除法头部应保留 LLM 返回的原始除数");
        assertEquals("12.5", spec.longDivisionHeader().dividend(), "小数除法头部应保留原始被除数");
        assertEquals("", spec.longDivisionHeader().quotient(), "未显式给出商时不应本地补出");
        assertTrue(spec.longDivisionSteps().isEmpty(), "小数长除法步骤区也不应本地补算");
    }

    @Test
    void shouldKeepExplicitDecimalQuotientFromLlmHeader() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[0.44]{5}{2.2}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.isLongDivision());
        assertEquals("5", spec.longDivisionHeader().divisor());
        assertEquals("2.2", spec.longDivisionHeader().dividend());
        assertEquals("0.44", spec.longDivisionHeader().quotient(), "带小数商时应保留 LLM 返回的显式商值");
        assertTrue(spec.longDivisionSteps().isEmpty(), "即便商为小数，步骤位置仍应由 LLM 提供");
    }

    @Test
    void shouldUnescapePercentInCrossArrayCells() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {10\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {20\\%} & {} & {} & {} & {20\\%}\\end{array}"
        );
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertEquals("50%", spec.rows().get(0).cells().get(0));
        assertEquals("10%", spec.rows().get(0).cells().get(4));
        assertEquals("30%", spec.rows().get(2).cells().get(2));
        assertEquals("20%", spec.rows().get(4).cells().get(0));
        assertEquals("20%", spec.rows().get(4).cells().get(4));
    }

    @Test
    void shouldRecognizeCrossMultiplicationLayoutWithoutFlatteningStructure() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {10\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {20\\%} & {} & {} & {} & {20\\%}\\end{array}"
        );
        VerticalLayoutCompiler.CrossMultiplicationLayout layout =
            compiler.compileCrossMultiplicationArray(ast.getChildren().get(0));

        assertNotNull(layout);
        assertEquals(5, layout.rows().size());
        assertEquals("50%", layout.rows().get(0).get(0));
        assertEquals("10%", layout.rows().get(0).get(4));
        assertEquals("\\searrow", layout.rows().get(1).get(1));
        assertEquals("\\nearrow", layout.rows().get(1).get(3));
        assertEquals("30%", layout.rows().get(2).get(2));
    }

    private String joinCells(java.util.List<String> cells) {
        StringBuilder builder = new StringBuilder();
        for (String cell : cells) {
            if (cell != null && !cell.isBlank()) {
                builder.append(cell);
            }
        }
        return builder.toString();
    }
}
