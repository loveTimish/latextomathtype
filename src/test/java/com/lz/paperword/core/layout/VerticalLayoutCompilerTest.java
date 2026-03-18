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
    void shouldCompileExplicitLongDivisionHeaderAndSteps() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.isLongDivision());
        assertEquals("13", spec.longDivisionHeader().divisor());
        assertEquals("65", spec.longDivisionHeader().quotient());
        assertEquals("845", spec.longDivisionHeader().dividend());
        assertEquals(2, spec.longDivisionSteps().size(), "explicit long division should produce structured step blocks");
        assertEquals("78", joinCells(spec.longDivisionSteps().get(0).productRow().cells()));
        assertEquals("65", joinCells(spec.longDivisionSteps().get(0).remainderRow().cells()));
    }

    @Test
    void shouldCompileCompositeLongDivisionArray() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{l}{\\longdiv[65]{13}{845}} \\\\ {\\begin{array}{rrr}{7} & {8} & {} \\\\ \\hline {} & {6} & {5} \\\\ {} & {6} & {5} \\\\ \\hline {} & {} & {0}\\end{array}}\\end{array}"
        );
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.isLongDivision());
        assertEquals("65", spec.longDivisionHeader().quotient());
        assertEquals("845", spec.longDivisionHeader().dividend());
        assertEquals(5, spec.columnCount());
        assertEquals(2, spec.ruleSpans().size());
        assertEquals("", spec.rows().get(0).cells().get(0));
        assertEquals("", spec.rows().get(0).cells().get(1));
    }

    @Test
    void shouldCompileExplicitLongDivisionIntoThreeStructuredSteps() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
        VerticalLayoutSpec spec = compiler.compile(ast);

        assertNotNull(spec);
        assertTrue(spec.hasStructuredLongDivisionSteps());
        assertEquals(3, spec.longDivisionSteps().size());
        assertEquals("10", joinCells(spec.longDivisionSteps().get(0).productRow().cells()));
        assertEquals("23", joinCells(spec.longDivisionSteps().get(0).remainderRow().cells()));
        assertEquals("20", joinCells(spec.longDivisionSteps().get(1).productRow().cells()));
        assertEquals("34", joinCells(spec.longDivisionSteps().get(1).remainderRow().cells()));
        assertEquals("30", joinCells(spec.longDivisionSteps().get(2).productRow().cells()));
        assertEquals("4", joinCells(spec.longDivisionSteps().get(2).remainderRow().cells()));
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
