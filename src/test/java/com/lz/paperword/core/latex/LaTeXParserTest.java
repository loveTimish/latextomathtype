package com.lz.paperword.core.latex;

import com.lz.paperword.core.latex.LaTeXParser.ContentSegment;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class LaTeXParserTest {

    private final LaTeXParser parser = new LaTeXParser();

    @Test
    void testParseSimpleHtml() {
        List<ContentSegment> segments = parser.parseHtml("<p>已知 $x=3$，求 $y$ 的值</p>");
        // segments: "已知 ", "x=3", "，求 ", "y", " 的值"
        assertEquals(5, segments.size());
        assertFalse(segments.get(0).isMath());
        assertTrue(segments.get(0).rawText().contains("已知"));
        assertTrue(segments.get(1).isMath());
        assertEquals("x=3", segments.get(1).rawText());
        assertTrue(segments.get(3).isMath());
        assertEquals("y", segments.get(3).rawText());
        assertFalse(segments.get(4).isMath());
        assertTrue(segments.get(4).rawText().contains("的值"));
    }

    @Test
    void testParseFraction() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{x+1}{2}");
        assertNotNull(ast);
        assertEquals(LaTeXNode.Type.ROOT, ast.getType());
        assertEquals(1, ast.getChildren().size());

        LaTeXNode frac = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.FRACTION, frac.getType());
        assertEquals(2, frac.getChildren().size()); // numerator + denominator
    }

    @Test
    void testParseSuperscript() {
        LaTeXNode ast = parser.parseLaTeX("x^{2}");
        assertNotNull(ast);
        LaTeXNode sup = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, sup.getType());
        assertEquals(2, sup.getChildren().size()); // base + exponent
    }

    @Test
    void testParseSubscript() {
        LaTeXNode ast = parser.parseLaTeX("a_{n}");
        assertNotNull(ast);
        LaTeXNode sub = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, sub.getType());
    }

    @Test
    void testParseSqrt() {
        LaTeXNode ast = parser.parseLaTeX("\\sqrt{x+1}");
        assertNotNull(ast);
        LaTeXNode sqrt = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SQRT, sqrt.getType());
        assertEquals(1, sqrt.getChildren().size()); // content
    }

    @Test
    void testParseNthRoot() {
        LaTeXNode ast = parser.parseLaTeX("\\sqrt[3]{x}");
        assertNotNull(ast);
        LaTeXNode sqrt = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SQRT, sqrt.getType());
        assertEquals(2, sqrt.getChildren().size()); // degree + content
    }

    @Test
    void testParseGreekLetters() {
        LaTeXNode ast = parser.parseLaTeX("\\alpha + \\beta");
        assertNotNull(ast);
        assertTrue(ast.getChildren().size() >= 2);
        assertEquals(LaTeXNode.Type.COMMAND, ast.getChildren().get(0).getType());
        assertEquals("\\alpha", ast.getChildren().get(0).getValue());
    }

    @Test
    void testParseComplexExpression() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{-b \\pm \\sqrt{b^{2}-4ac}}{2a}");
        assertNotNull(ast);
        // Should have a fraction at root level
        LaTeXNode frac = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.FRACTION, frac.getType());
    }

    @Test
    void testParsePlainTextOnly() {
        List<ContentSegment> segments = parser.parseHtml("<p>这是纯文本内容</p>");
        assertEquals(1, segments.size());
        assertFalse(segments.get(0).isMath());
        assertEquals("这是纯文本内容", segments.get(0).rawText());
    }

    @Test
    void testParseEmptyInput() {
        List<ContentSegment> segments = parser.parseHtml("");
        assertTrue(segments.isEmpty());
    }

    @Test
    void testParseSumWithLimits() {
        LaTeXNode ast = parser.parseLaTeX("\\sum_{i=1}^{n}a_i");
        assertNotNull(ast);
        assertTrue(ast.getChildren().size() > 0);
    }

    @Test
    void testParseArrayForVerticalAddition() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rrrr", array.getMetadata("columnSpec"));
        assertEquals("4", array.getMetadata("columnCount"));
        assertEquals("0,0,0,0,0", array.getMetadata("columnLines"));
        assertEquals("0,0,1,0", array.getMetadata("rowLines"));
        assertEquals(3, array.getChildren().size());

        LaTeXNode firstRow = array.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ROW, firstRow.getType());
        assertEquals(4, firstRow.getChildren().size());
    }

    @Test
    void testParseArrayForDecimalAlignment() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcr}12 & . & 50 \\\\ +3 & . & 75 \\\\ \\hline 16 & . & 25\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rcr", array.getMetadata("columnSpec"));
        assertEquals("3", array.getMetadata("columnCount"));
        assertEquals(3, array.getChildren().size());
        assertEquals(".", array.getChildren().get(0).getChildren().get(1).getChildren().get(0).getValue());
    }

    @Test
    void testParseArrayForDivisionTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("r|l", array.getMetadata("columnSpec"));
        assertEquals("0,1,0", array.getMetadata("columnLines"));
        assertEquals(5, array.getChildren().size());
    }
}
