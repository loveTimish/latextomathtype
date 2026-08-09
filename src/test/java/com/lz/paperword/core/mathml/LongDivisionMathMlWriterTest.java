package com.lz.paperword.core.mathml;

import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LongDivisionMathMlWriterTest {
    private final LaTeXParser parser = new LaTeXParser();

    @Test
    void writesStandardMathMl3LongDivisionWithoutLatexFlattening() {
        MathIRNode ir = parser.parseDetailed(
            "\\begin{longdivision}{rrrr}{6}{570}{3420}"
                + "&\\frac{30}{1}\\\\\\cline{1-2}&&\\sqrt{42}\\\\&&42\\\\\\cline{2-3}&&&0"
                + "\\end{longdivision}").mathIR().getChildren().get(0);

        String mathml = new LongDivisionMathMlWriter().write(ir);

        assertTrue(mathml.contains("<mlongdiv longdivstyle=\"lefttop\">"));
        assertTrue(mathml.contains("<menclose notation=\"longdiv\">"));
        assertTrue(mathml.contains("<annotation-xml encoding=\"application/mathml-presentation+xml\">"));
        assertTrue(mathml.contains("<menclose notation=\"bottom\"><mtable width="));
        assertTrue(mathml.contains("columnalign=\"right\""));
        assertTrue(mathml.contains("<mn>6</mn><mn>570</mn><mn>3420</mn>"));
        assertTrue(mathml.contains("<mfrac>"));
        assertTrue(mathml.contains("<msqrt>"));
        assertTrue(mathml.contains("<msgroup position=\"2\"><msrow><mrow><mfrac>"));
        assertTrue(mathml.contains("<msline length=\"2\"/>"));
        assertTrue(mathml.contains("<msgroup position=\"1\"><msrow><mrow><msqrt>"));
        assertFalse(mathml.contains("overset"));
        assertFalse(mathml.contains("overline"));
    }

    @Test
    void keepsWholeNumbersAsDirectMstackDigits() {
        MathIRNode ir = parser.parseDetailed(
            "\\begin{longdivision}{rrrr}{6}{570}{3420}&30\\\\\\cline{1-2}&&42"
                + "\\\\&&42\\\\\\cline{2-3}&&&0\\end{longdivision}")
            .mathIR().getChildren().get(0);

        String mathml = new LongDivisionMathMlWriter().write(ir);

        assertTrue(mathml.contains("<mn>6</mn><mn>570</mn><mn>3420</mn>"));
        assertTrue(mathml.contains("<msgroup position=\"2\"><msrow><mrow><mn>30</mn></mrow></msrow>"));
        assertTrue(mathml.contains("<msgroup position=\"1\"><msrow><mrow><mn>42</mn></mrow></msrow>"));
        assertFalse(mathml.contains("<mrow><mn>3</mn><mn>4</mn><mn>2</mn><mn>0</mn></mrow>"));
    }

    @Test
    void writesEmptyQuotientWithoutMathJaxErrorGlyph() {
        MathIRNode ir = parser.parseDetailed(
            "\\begin{longdivision}{rr}{6}{}{12}&12\\end{longdivision}")
            .mathIR().getChildren().get(0);

        String mathml = new LongDivisionMathMlWriter().write(ir);

        assertTrue(mathml.contains("<mn>6</mn><mrow/><mn>12</mn>"));
        assertFalse(mathml.contains("<none/>"));
    }

    @Test
    void writesSymbolicQuotientAndDividendAsElementaryRows() {
        MathIRNode ir = parser.parseDetailed(
            "\\begin{longdivision}{rrr}{x-1}{x+1}{x^{2}-1}&x^{2}-x\\end{longdivision}")
            .mathIR().getChildren().get(0);

        String mathml = new LongDivisionMathMlWriter().write(ir);

        assertTrue(mathml.contains("<msrow><mi>x</mi><mo>+</mo><mn>1</mn></msrow>"));
        assertTrue(mathml.contains("<msrow><msup>"));
    }

    @Test
    void positionsRuleAndStepIndependentlyAndKeepsDecimalsTogether() {
        MathIRNode ir = parser.parseDetailed(
            "\\begin{longdivision}{rrrr}{2.5}{}{12.5}&125\\\\\\cline{1-3}&&&0"
                + "\\end{longdivision}").mathIR().getChildren().get(0);

        String mathml = new LongDivisionMathMlWriter().write(ir);

        assertTrue(mathml.contains("<mn>2.5</mn><mrow/><mn>12.5</mn>"));
        assertTrue(mathml.contains("<msgroup position=\"1\"><msrow position=\"1\">"));
        assertTrue(mathml.contains("<msline length=\"3\"/>"));
    }

    @Test
    void longDivisionSpecIsDeeplyImmutable() {
        MathIRNode ir = parser.parseDetailed(
            "\\begin{longdivision}{rr}{6}{}{12}&12\\\\\\cline{1-2}&0\\end{longdivision}")
            .mathIR().getChildren().get(0);
        LongDivisionSpec spec = LongDivisionSpec.from(ir);

        ir.child(0).setValue("changed");
        ir.child(3).getChildren().clear();

        assertEquals("6", spec.divisor().children().get(0).value());
        assertEquals(2, spec.steps().size());
        assertThrows(UnsupportedOperationException.class, () -> spec.steps().clear());
        assertThrows(UnsupportedOperationException.class, () -> spec.divisor().metadata().clear());
    }
}
