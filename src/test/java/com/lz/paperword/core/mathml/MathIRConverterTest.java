package com.lz.paperword.core.mathml;

import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class MathIRConverterTest {

    private final LaTeXParser parser = new LaTeXParser();

    @Test
    void testParseMathIrNormalizesRootsScriptsAndFractions() {
        MathIRNode ir = parser.parseMathIR("\\frac{1}{\\sqrt[3]{x_i}}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(1, ir.getChildren().size());

        MathIRNode fraction = ir.child(0);
        assertEquals(MathIRNode.Type.FRACTION, fraction.getType());
        assertEquals(MathIRNode.Type.SEQUENCE, fraction.child(0).getType());
        assertEquals(MathIRNode.Type.NUMBER, fraction.child(0).child(0).getType());
        assertEquals(MathIRNode.Type.SEQUENCE, fraction.child(1).getType());
        assertEquals(MathIRNode.Type.ROOT, fraction.child(1).child(0).getType());
        assertEquals(MathIRNode.Type.SEQUENCE, fraction.child(1).child(0).child(1).getType());
        assertEquals(MathIRNode.Type.SUB, fraction.child(1).child(0).child(1).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesBigOperatorLimits() {
        MathIRNode ir = parser.parseMathIR("\\sum_{i=1}^{n} a_i");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.UNDEROVER, ir.child(0).getType());
        assertEquals(MathIRNode.Type.OPERATOR, ir.child(0).child(0).getType());
        assertEquals("big-operator", ir.child(0).child(0).getMetadata("role"));
        assertEquals("underover", ir.child(0).child(0).getMetadata("limitPlacement"));
        assertEquals(MathIRNode.Type.SUB, ir.child(1).getType());
    }

    @Test
    void testParseMathIrNormalizesCoproductLimitsToDedicatedBigOperator() {
        MathIRNode ir = parser.parseMathIR("\\coprod_{i=1}^{n} A_i");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.UNDEROVER, ir.child(0).getType());
        assertEquals(MathIRNode.Type.OPERATOR, ir.child(0).child(0).getType());
        assertEquals("\\coprod", ir.child(0).child(0).getMetadata("latexCommand"));
        assertEquals("big-operator", ir.child(0).child(0).getMetadata("role"));
        assertEquals("underover", ir.child(0).child(0).getMetadata("limitPlacement"));
        assertEquals(MathIRNode.Type.SUB, ir.child(1).getType());
    }

    @Test
    void testParseMathIrNormalizesFencedMatrixEnvironment() {
        MathIRNode ir = parser.parseMathIR("\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.FENCE, ir.child(0).getType());
        assertEquals("(", ir.child(0).getMetadata("openDelimiter"));
        assertEquals(")", ir.child(0).getMetadata("closeDelimiter"));
        assertEquals(MathIRNode.Type.TABLE, ir.child(0).child(0).getType());
        assertEquals("pmatrix", ir.child(0).child(0).getMetadata("environment"));
    }

    @Test
    void testParseMathIrPreservesExtendedFenceMetadata() {
        MathIRNode ir = parser.parseMathIR("\\left\\lVert x \\right.");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.FENCE, ir.child(0).getType());
        assertEquals("||", ir.child(0).getMetadata("openDelimiter"));
        assertEquals(".", ir.child(0).getMetadata("closeDelimiter"));
    }

    @Test
    void testParseMathIrPreservesAngleFenceMetadata() {
        MathIRNode ir = parser.parseMathIR("\\left\\langle x+y \\right\\rangle");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.FENCE, ir.child(0).getType());
        assertEquals("⟨", ir.child(0).getMetadata("openDelimiter"));
        assertEquals("⟩", ir.child(0).getMetadata("closeDelimiter"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrPreservesOpenBracketFenceMetadata() {
        MathIRNode ir = parser.parseMathIR("\\left\\llbracket x+y \\right\\rrbracket");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.FENCE, ir.child(0).getType());
        assertEquals("⟦", ir.child(0).getMetadata("openDelimiter"));
        assertEquals("⟧", ir.child(0).getMetadata("closeDelimiter"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesBoxedEnclosure() {
        MathIRNode ir = parser.parseMathIR("\\boxed{x+1}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ENCLOSURE, ir.child(0).getType());
        assertEquals("box", ir.child(0).getMetadata("notation"));
        assertEquals("\\boxed", ir.child(0).getMetadata("latexCommand"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesCancelEnclosure() {
        MathIRNode ir = parser.parseMathIR("\\xcancel{x}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ENCLOSURE, ir.child(0).getType());
        assertEquals("updiagonalstrike downdiagonalstrike", ir.child(0).getMetadata("notation"));
        assertEquals("\\xcancel", ir.child(0).getMetadata("latexCommand"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrKeepsCasesStructureExplicit() {
        MathIRNode ir = parser.parseMathIR("\\begin{cases}x&1\\\\y&2\\end{cases}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.TABLE, ir.child(0).getType());
        assertEquals("cases", ir.child(0).getMetadata("environment"));
        assertEquals("{", ir.child(0).getMetadata("openDelimiter"));
        assertEquals(".", ir.child(0).getMetadata("closeDelimiter"));
        assertEquals(MathIRNode.Type.TABLE_ROW, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrKeepsAlignedRelationPairMetadata() {
        MathIRNode ir = parser.parseMathIR("\\begin{aligned}a&=b\\\\c&=d\\end{aligned}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.TABLE, ir.child(0).getType());
        assertEquals("aligned", ir.child(0).getMetadata("environment"));
        assertEquals("relation-pairs", ir.child(0).getMetadata("alignmentMode"));
        assertEquals("rl", ir.child(0).getMetadata("columnSpec"));
    }

    @Test
    void testParseMathIrKeepsAlignedMultiPairColumns() {
        MathIRNode ir = parser.parseMathIR("\\begin{aligned}a&=b&c&=d\\end{aligned}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.TABLE, ir.child(0).getType());
        assertEquals("aligned", ir.child(0).getMetadata("environment"));
        assertEquals("relation-pairs", ir.child(0).getMetadata("alignmentMode"));
        assertEquals("rlrl", ir.child(0).getMetadata("columnSpec"));
        assertEquals("4", ir.child(0).getMetadata("columnCount"));
        assertEquals(4, ir.child(0).child(0).getChildren().size());
    }

    @Test
    void testParseMathIrKeepsSplitRelationPairMetadata() {
        MathIRNode ir = parser.parseMathIR("\\begin{split}a&=b\\\\c&=d\\end{split}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.TABLE, ir.child(0).getType());
        assertEquals("split", ir.child(0).getMetadata("environment"));
        assertEquals("relation-pairs", ir.child(0).getMetadata("alignmentMode"));
        assertEquals("rl", ir.child(0).getMetadata("columnSpec"));
    }

    @Test
    void testParseMathIrKeepsAlignStarRelationPairMetadata() {
        MathIRNode ir = parser.parseMathIR("\\begin{align*}a&=b\\\\c&=d\\end{align*}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.TABLE, ir.child(0).getType());
        assertEquals("align*", ir.child(0).getMetadata("environment"));
        assertEquals("relation-pairs", ir.child(0).getMetadata("alignmentMode"));
        assertEquals("rl", ir.child(0).getMetadata("columnSpec"));
    }

    @Test
    void testDumpMathIrMarksUnsupportedCommandsExplicitly() {
        String dump = parser.dumpMathIR("\\foo{1}");

        assertTrue(dump.contains("UNSUPPORTED(\\foo)"));
        assertTrue(dump.contains("latex=\\foo"));
    }

    @Test
    void testParseMathIrNormalizesOverarcToDedicatedArcNode() {
        MathIRNode ir = parser.parseMathIR("\\overarc{AB}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ARC, ir.child(0).getType());
        assertEquals("\\overarc", ir.child(0).getMetadata("latexCommand"));
        assertEquals("top", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesArcAliasToDedicatedArcNode() {
        MathIRNode ir = parser.parseMathIR("\\arc{AB}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ARC, ir.child(0).getType());
        assertEquals("\\arc", ir.child(0).getMetadata("latexCommand"));
        assertEquals("top", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesOverparenToDedicatedArcNode() {
        MathIRNode ir = parser.parseMathIR("\\overparen{AB}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ARC, ir.child(0).getType());
        assertEquals("\\overparen", ir.child(0).getMetadata("latexCommand"));
        assertEquals("top", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesWideparenToDedicatedArcNode() {
        MathIRNode ir = parser.parseMathIR("\\wideparen{AB}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ARC, ir.child(0).getType());
        assertEquals("\\wideparen", ir.child(0).getMetadata("latexCommand"));
        assertEquals("top", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesBraToDiracLeftSlice() {
        MathIRNode ir = parser.parseMathIR("\\bra{\\psi}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.DIRAC, ir.child(0).getType());
        assertEquals("\\bra", ir.child(0).getMetadata("latexCommand"));
        assertEquals("true", ir.child(0).getMetadata("leftPresent"));
        assertEquals("false", ir.child(0).getMetadata("rightPresent"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        assertEquals(MathIRNode.Type.IDENT, ir.child(0).child(0).child(0).getType());
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(1).getType());
        assertEquals(0, ir.child(0).child(1).getChildren().size());
    }

    @Test
    void testParseMathIrNormalizesKetToDiracRightSlice() {
        MathIRNode ir = parser.parseMathIR("\\ket{\\psi}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.DIRAC, ir.child(0).getType());
        assertEquals("\\ket", ir.child(0).getMetadata("latexCommand"));
        assertEquals("false", ir.child(0).getMetadata("leftPresent"));
        assertEquals("true", ir.child(0).getMetadata("rightPresent"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        assertEquals(0, ir.child(0).child(0).getChildren().size());
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(1).getType());
        assertEquals(MathIRNode.Type.IDENT, ir.child(0).child(1).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesBraketToDualSlotDirac() {
        MathIRNode ir = parser.parseMathIR("\\braket{\\psi|\\phi}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.DIRAC, ir.child(0).getType());
        assertEquals("\\braket", ir.child(0).getMetadata("latexCommand"));
        assertEquals("true", ir.child(0).getMetadata("leftPresent"));
        assertEquals("true", ir.child(0).getMetadata("rightPresent"));
        assertEquals("|", ir.child(0).getMetadata("middleDelimiter"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        assertEquals(MathIRNode.Type.IDENT, ir.child(0).child(0).child(0).getType());
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(1).getType());
        assertEquals(MathIRNode.Type.IDENT, ir.child(0).child(1).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesOverbraceWithNoAnnotation() {
        MathIRNode ir = parser.parseMathIR("\\overbrace{x+y}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.HBRACE, ir.child(0).getType());
        assertEquals("\\overbrace", ir.child(0).getMetadata("latexCommand"));
        assertEquals("top", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        // convertArgument(null) returns empty SEQUENCE, not null
        assertNotNull(ir.child(0).child(1));
        assertEquals(0, ir.child(0).child(1).getChildren().size());
    }

    @Test
    void testParseMathIrNormalizesUnderbraceWithAnnotationSubscript() {
        MathIRNode ir = parser.parseMathIR("\\underbrace{x+y}_{i=1}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.HBRACE, ir.child(0).getType());
        assertEquals("\\underbrace", ir.child(0).getMetadata("latexCommand"));
        assertEquals("bottom", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        assertNotNull(ir.child(0).child(1));
    }

    @Test
    void testParseMathIrNormalizesOverbracketWithNoAnnotation() {
        MathIRNode ir = parser.parseMathIR("\\overbracket{x+y}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.HBRACK, ir.child(0).getType());
        assertEquals("\\overbracket", ir.child(0).getMetadata("latexCommand"));
        assertEquals("top", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        assertNotNull(ir.child(0).child(1));
        assertEquals(0, ir.child(0).child(1).getChildren().size());
    }

    @Test
    void testParseMathIrNormalizesUnderbracketWithAnnotationSubscript() {
        MathIRNode ir = parser.parseMathIR("\\underbracket{x+y}_{i=1}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.HBRACK, ir.child(0).getType());
        assertEquals("\\underbracket", ir.child(0).getMetadata("latexCommand"));
        assertEquals("bottom", ir.child(0).getMetadata("placement"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
        assertNotNull(ir.child(0).child(1));
    }

    @Test
    void testParseMathIrNormalizesXrightarrowToDedicatedArrowNode() {
        MathIRNode ir = parser.parseMathIR("\\xrightarrow{n\\to\\infty}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ARROW, ir.child(0).getType());
        assertEquals("\\xrightarrow", ir.child(0).getMetadata("latexCommand"));
        assertEquals("right", ir.child(0).getMetadata("direction"));
        assertEquals("single", ir.child(0).getMetadata("variant"));
        assertEquals("true", ir.child(0).getMetadata("topPresent"));
        assertEquals("false", ir.child(0).getMetadata("bottomPresent"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }

    @Test
    void testParseMathIrNormalizesXleftarrowToDedicatedArrowNode() {
        MathIRNode ir = parser.parseMathIR("\\xleftarrow{f}");

        assertEquals(MathIRNode.Type.MATH, ir.getType());
        assertEquals(MathIRNode.Type.ARROW, ir.child(0).getType());
        assertEquals("\\xleftarrow", ir.child(0).getMetadata("latexCommand"));
        assertEquals("left", ir.child(0).getMetadata("direction"));
        assertEquals(MathIRNode.Type.SEQUENCE, ir.child(0).child(0).getType());
    }
}
