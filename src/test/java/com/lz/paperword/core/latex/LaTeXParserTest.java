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
    void testParseBareLatexAsPlainTextWhenMissingDelimiters() {
        List<ContentSegment> segments = parser.parseText("【解答】\\frac{1}{2}+\\frac{1}{3}");
        assertEquals(1, segments.size());
        assertFalse(segments.get(0).isMath());
        assertEquals("【解答】\\frac{1}{2}+\\frac{1}{3}", segments.get(0).rawText());
    }

    @Test
    void testParseMultilineDisplayMathKeepsWholeBlock() {
        List<ContentSegment> segments = parser.parseText("【解答】\n$$\\begin{array}{ccccc}\n{30\\%} & {} & {} & {} & {20\\%} \\\\\n{} & {\\searrow} & {} & {\\nearrow} & {}\n\\end{array}$$");
        assertEquals(2, segments.size());
        assertFalse(segments.get(0).isMath());
        assertTrue(segments.get(1).isMath());
        assertTrue(segments.get(1).rawText().contains("\\begin{array}{ccccc}"));
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
    void testParseArrayForVerticalSubtraction() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 8 & 6 & 4 \\\\ - & 2 & 7 & 9 \\\\ \\hline & 5 & 8 & 5\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rrrr", array.getMetadata("columnSpec"));
        assertEquals(3, array.getChildren().size());
    }

    @Test
    void testParseArrayForDecimalSubtraction() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcr}12 & . & 50 \\\\ -3 & . & 75 \\\\ \\hline 8 & . & 75\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rcr", array.getMetadata("columnSpec"));
        assertEquals(3, array.getChildren().size());
    }

    @Test
    void testParseExplicitLongDivisionCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode longDiv = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.LONG_DIVISION, longDiv.getType());
        assertEquals("\\longdiv", longDiv.getValue());
        assertEquals(3, longDiv.getChildren().size());
        assertEquals("13", flatten(longDiv.getChildren().get(0)));
        assertEquals("65", flatten(longDiv.getChildren().get(1)));
        assertEquals("845", flatten(longDiv.getChildren().get(2)));
    }

    @Test
    void testParseExplicitLongDivisionCommandWithoutQuotient() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{13}{845}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode longDiv = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.LONG_DIVISION, longDiv.getType());
        assertEquals(3, longDiv.getChildren().size());
        assertEquals("13", flatten(longDiv.getChildren().get(0)));
        assertEquals("", flatten(longDiv.getChildren().get(1)));
        assertEquals("845", flatten(longDiv.getChildren().get(2)));
    }

    @Test
    void testParseCompositeLongDivisionWithStepArrayInSingleBlock() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\longdiv[570]{6}{3420}\\begin{array}{l}\\text{   }\\underline{30}\\\\\\text{    }42\\\\\\text{    }\\underline{42}\\\\\\text{      }0\\end{array}"
        );
        assertNotNull(ast);
        assertEquals(2, ast.getChildren().size());
        assertEquals(LaTeXNode.Type.LONG_DIVISION, ast.getChildren().get(0).getType());
        assertEquals(LaTeXNode.Type.ARRAY, ast.getChildren().get(1).getType());
        assertEquals("l", ast.getChildren().get(1).getMetadata("columnSpec"));
        assertEquals(4, ast.getChildren().get(1).getChildren().size());
        assertTrue(flatten(ast.getChildren().get(1)).contains("    42"), "\\text{空格} 中的前导空格应被完整保留");
    }

    @Test
    void testParseConcentrationCrossArrayKeepsArrowCommandsInMathAst() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {10\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {20\\%} & {} & {} & {} & {20\\%}\\end{array}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("ccccc", array.getMetadata("columnSpec"));
        assertEquals(5, array.getChildren().size());
        assertTrue(flatten(array).contains("\\searrow"));
        assertTrue(flatten(array).contains("\\nearrow"));
    }

    @Test
    void testParseMatrixEnvironmentPromotesToArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{matrix}1&2\\\\3&4\\end{matrix}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("matrix", array.getMetadata("environment"));
        assertEquals("cc", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParsePmatrixEnvironmentWrapsArrayWithFence() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("\\left(", fence.getValue());
        assertEquals("(", fence.getMetadata("leftDelimiter"));
        assertEquals(")", fence.getMetadata("rightDelimiter"));
        assertEquals(1, fence.getChildren().size());
        assertEquals(LaTeXNode.Type.ARRAY, fence.getChildren().get(0).getType());
        assertEquals("pmatrix", fence.getChildren().get(0).getMetadata("environment"));
    }

    @Test
    void testParseLeftRightPreservesBothDelimiters() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\{x+1\\right]");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("\\left{", fence.getValue());
        assertEquals("{", fence.getMetadata("leftDelimiter"));
        assertEquals("]", fence.getMetadata("rightDelimiter"));
    }

    @Test
    void testParseBoxedCommandAsUnaryEnclosure() {
        LaTeXNode ast = parser.parseLaTeX("\\boxed{x+1}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode boxed = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, boxed.getType());
        assertEquals("\\boxed", boxed.getValue());
        assertEquals(1, boxed.getChildren().size());
        assertEquals("x+1", flatten(boxed.getChildren().get(0)));
    }

    @Test
    void testParseCancelCommandAsUnaryEnclosure() {
        LaTeXNode ast = parser.parseLaTeX("\\xcancel{x+1}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode cancel = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, cancel.getType());
        assertEquals("\\xcancel", cancel.getValue());
        assertEquals(1, cancel.getChildren().size());
        assertEquals("x+1", flatten(cancel.getChildren().get(0)));
    }

    @Test
    void testParseLeftRightNormalizesExtendedFenceDelimiters() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\lfloor x \\right\\rceil");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("⌊", fence.getMetadata("leftDelimiter"));
        assertEquals("⌉", fence.getMetadata("rightDelimiter"));
    }

    @Test
    void testParseVmatrixEnvironmentWrapsArrayWithDoubleBarFence() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{Vmatrix}1&2\\\\3&4\\end{Vmatrix}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("\\left\\lVert", fence.getValue());
        assertEquals("||", fence.getMetadata("leftDelimiter"));
        assertEquals("||", fence.getMetadata("rightDelimiter"));
        assertEquals("Vmatrix", fence.getChildren().get(0).getMetadata("environment"));
    }

    @Test
    void testParseAlignedEnvironmentPromotesToImplicitArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{aligned}a&=b\\\\c&=d\\end{aligned}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("aligned", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rl", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseAlignedEnvironmentAlternatesColumnSpecAcrossMultiplePairs() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{aligned}a&=b&c&=d\\end{aligned}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("aligned", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rlrl", array.getMetadata("columnSpec"));
        assertEquals("4", array.getMetadata("columnCount"));
    }

    @Test
    void testParseSplitEnvironmentPromotesToRelationPairArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{split}a&=b\\\\c&=d\\end{split}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("split", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rl", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseAlignStarEnvironmentPromotesToRelationPairArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{align*}a&=b\\\\c&=d\\end{align*}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("align*", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rl", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseCasesEnvironmentPromotesToImplicitArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{cases}x&1\\\\y&2\\end{cases}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("cases", array.getMetadata("environment"));
        assertEquals("ll", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseMultiplicationArrayMarksExplicitEmptyCells() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {}\\end{array}"
        );
        assertNotNull(ast);

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("true", array.getChildren().get(0).getChildren().get(0).getMetadata("explicitEmptyCell"));
        assertEquals("true", array.getChildren().get(1).getChildren().get(1).getMetadata("explicitEmptyCell"));
    }

    private String flatten(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR || node.getType() == LaTeXNode.Type.COMMAND) {
            return node.getValue() == null ? "" : node.getValue();
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flatten(child));
        }
        return builder.toString();
    }

    @Test
    void testParseOverbraceAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\overbrace{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode overbrace = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, overbrace.getType());
        assertEquals("\\overbrace", overbrace.getValue());
        assertEquals(1, overbrace.getChildren().size());
        assertEquals("a+b+c", flatten(overbrace.getChildren().get(0)));
    }

    @Test
    void testParseUnderbraceAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\underbrace{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode underbrace = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, underbrace.getType());
        assertEquals("\\underbrace", underbrace.getValue());
        assertEquals(1, underbrace.getChildren().size());
        assertEquals("a+b+c", flatten(underbrace.getChildren().get(0)));
    }

    @Test
    void testParseOverbracketAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\overbracket{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode overbracket = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, overbracket.getType());
        assertEquals("\\overbracket", overbracket.getValue());
        assertEquals(1, overbracket.getChildren().size());
        assertEquals("a+b+c", flatten(overbracket.getChildren().get(0)));
    }

    @Test
    void testParseUnderbracketAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\underbracket{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode underbracket = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, underbracket.getType());
        assertEquals("\\underbracket", underbracket.getValue());
        assertEquals(1, underbracket.getChildren().size());
        assertEquals("a+b+c", flatten(underbracket.getChildren().get(0)));
    }
}
