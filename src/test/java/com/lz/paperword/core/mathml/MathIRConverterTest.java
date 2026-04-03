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
}
