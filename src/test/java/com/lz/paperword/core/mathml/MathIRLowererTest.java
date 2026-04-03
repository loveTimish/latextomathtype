package com.lz.paperword.core.mathml;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

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
}
