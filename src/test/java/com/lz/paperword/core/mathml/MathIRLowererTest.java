package com.lz.paperword.core.mathml;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MathIRLowererTest {

    private final LaTeXParser parser = new LaTeXParser();
    private final MathIRLowerer lowerer = new MathIRLowerer();

    @Test
    void testLowerBraDiracNodeBackToUnaryCommand() {
        MathIRNode ir = parser.parseMathIR("\\bra{\\psi}");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\bra", lowered.getChildren().get(0).getValue());
        assertEquals(1, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerKetDiracNodeBackToUnaryCommand() {
        MathIRNode ir = parser.parseMathIR("\\ket{\\psi}");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\ket", lowered.getChildren().get(0).getValue());
        assertEquals(1, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerBraketDiracNodeBackToBinaryCommand() {
        MathIRNode ir = parser.parseMathIR("\\braket{\\psi|\\phi}");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\braket", lowered.getChildren().get(0).getValue());
        assertEquals("|", lowered.getChildren().get(0).getMetadata("middleDelimiter"));
        assertEquals(2, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerXrightarrowArrowNodeBackToDedicatedCommand() {
        MathIRNode ir = parser.parseMathIR("\\xrightarrow{f}");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\xrightarrow", lowered.getChildren().get(0).getValue());
        assertEquals("TM_ARROW", lowered.getChildren().get(0).getMetadata("templateFamily"));
        assertEquals(1, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerXleftarrowArrowNodeBackToDedicatedCommand() {
        MathIRNode ir = parser.parseMathIR("\\xleftarrow{g}");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\xleftarrow", lowered.getChildren().get(0).getValue());
        assertEquals("left", lowered.getChildren().get(0).getMetadata("arrowDirection"));
        assertEquals(1, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerAngleFenceNodeBackToDedicatedFenceCommand() {
        MathIRNode ir = parser.parseMathIR("\\left\\langle x+y \\right\\rangle");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\left\\langle", lowered.getChildren().get(0).getValue());
        assertEquals("⟨", lowered.getChildren().get(0).getMetadata("leftDelimiter"));
        assertEquals("⟩", lowered.getChildren().get(0).getMetadata("rightDelimiter"));
        assertEquals(1, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerOpenBracketFenceNodeBackToDedicatedFenceCommand() {
        MathIRNode ir = parser.parseMathIR("\\left\\llbracket x+y \\right\\rrbracket");

        LaTeXNode lowered = lowerer.lower(ir);

        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(1, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.COMMAND, lowered.getChildren().get(0).getType());
        assertEquals("\\left\\llbracket", lowered.getChildren().get(0).getValue());
        assertEquals("⟦", lowered.getChildren().get(0).getMetadata("leftDelimiter"));
        assertEquals("⟧", lowered.getChildren().get(0).getMetadata("rightDelimiter"));
        assertEquals(1, lowered.getChildren().get(0).getChildren().size());
    }

    @Test
    void testLowerCoproductBigOperatorBackToCoproductCommandShape() {
        MathIRNode ir = parser.parseMathIR("\\coprod_{i=1}^{n} A_i");

        LaTeXNode lowered = lowerer.lower(ir);

        assertNotNull(lowered);
        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(2, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, lowered.getChildren().get(0).getType());
        assertEquals(LaTeXNode.Type.SUBSCRIPT, lowered.getChildren().get(0).getChildren().get(0).getType());
        assertEquals("\\coprod", lowered.getChildren().get(0).getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testLowerUnionBigOperatorBackToBigcupCommandShape() {
        MathIRNode ir = parser.parseMathIR("\\bigcup_{i=1}^{n} A_i");

        LaTeXNode lowered = lowerer.lower(ir);

        assertNotNull(lowered);
        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(2, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, lowered.getChildren().get(0).getType());
        assertEquals("\\bigcup", lowered.getChildren().get(0).getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testLowerIntersectionBigOperatorBackToBigcapCommandShape() {
        MathIRNode ir = parser.parseMathIR("\\bigcap_{i=1}^{n} A_i");

        LaTeXNode lowered = lowerer.lower(ir);

        assertNotNull(lowered);
        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(2, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, lowered.getChildren().get(0).getType());
        assertEquals("\\bigcap", lowered.getChildren().get(0).getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testLowerSummationStyleGenericBigOperatorBackToBigoplusCommandShape() {
        MathIRNode ir = parser.parseMathIR("\\bigoplus_{i=1}^{n} A_i");

        LaTeXNode lowered = lowerer.lower(ir);

        assertNotNull(lowered);
        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertEquals(2, lowered.getChildren().size());
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, lowered.getChildren().get(0).getType());
        assertEquals("\\bigoplus", lowered.getChildren().get(0).getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testLowerIntegralStyleGenericBigOperatorBackToIntopCommandShape() {
        MathIRNode ir = parser.parseMathIR("\\intop_{a}^{b} f(x)");

        LaTeXNode lowered = lowerer.lower(ir);

        assertNotNull(lowered);
        assertEquals(LaTeXNode.Type.ROOT, lowered.getType());
        assertTrue(!lowered.getChildren().isEmpty());
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, lowered.getChildren().get(0).getType());
        assertEquals("\\intop", lowered.getChildren().get(0).getChildren().get(0).getChildren().get(0).getValue());
    }
}
