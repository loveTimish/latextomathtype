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
}
