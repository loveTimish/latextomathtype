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
    void shouldCompileExplicitLongDivisionHeaderWithoutSynthesizingSteps() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.isLongDivision());
        assertEquals("13", spec.longDivisionHeader().divisor());
        assertEquals("65", spec.longDivisionHeader().quotient());
        assertEquals("845", spec.longDivisionHeader().dividend());
        assertTrue(spec.longDivisionSteps().isEmpty(), "bare longdiv 头部不应再自动推导步骤区");
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
    void shouldNotCompileThreeStepLongDivisionFromHeaderAlone() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertFalse(spec.hasStructuredLongDivisionSteps(), "不能只凭 longdiv 头部自动补三步长除法");
        assertTrue(spec.longDivisionSteps().isEmpty(), "步骤区应由原图显式提供，而不是代码自动生成");
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
