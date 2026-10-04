package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;
import java.nio.charset.StandardCharsets;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class GeometricBaselineRegressionTest {
    @Test
    void defaultPreviewCarriesActualMathJaxBaselineRatherThanFamilyRatios() throws Exception {
        var renderer = new LaTeXImageRenderer();
        for (String latex : List.of("x", "x_i^2+1", "\\frac{a+b}{c+d}",
                "\\sqrt{x^2+y^2}", "\\frac{1}{1+\\frac{1}{x+1}}",
                "\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}")) {
            var svg = renderer.renderMathJaxSvgForAcceptance(latex);
            var preview = renderer.renderForOlePreview(latex);
            assertEquals(svg.depthPt(), preview.depthPt(), 0.05, latex);
            assertTrue(WmfPreviewInspector.inspect(preview.data()).pureVector());
        }
    }

    @Test
    void targetBoxLetterboxingMovesBaselineBySameAffineTransformAsInk() throws Exception {
        var renderer = new LaTeXImageRenderer();
        String latex = "\\frac{x+1}{y+2}";
        var svg = renderer.renderMathJaxSvgForAcceptance(latex);
        // Width permits a 2x scale; 3x height leaves half a source height above/below.
        var preview = renderer.renderForOlePreview(latex, svg.widthPt()*2, svg.heightPt()*3);
        double expectedDepth = svg.depthPt()*2 + svg.heightPt()*0.5;
        assertEquals(expectedDepth, preview.depthPt(), 0.06);
    }

    @Test
    void unknownDepthCannotBecomeARealDescentAfterLetterboxing() {
        assertThrows(java.io.IOException.class, () -> LaTeXImageRenderer.sourceBaselineFraction(
            new LaTeXImageRenderer.MathJaxSvgResult(new byte[0],24,12,-1)));
        assertThrows(java.io.IOException.class, () -> LaTeXImageRenderer.sourceBaselineFraction(
            new LaTeXImageRenderer.MathJaxSvgResult(new byte[0],24,0,3)));
        assertThrows(java.io.IOException.class, () -> LaTeXImageRenderer.sourceBaselineFraction(
            new LaTeXImageRenderer.MathJaxSvgResult(new byte[0],24,12,Double.NaN)));
    }

    @Test
    void baselineMappingAccountsForEncoderSafetyPadding() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24pt\" height=\"12pt\" viewBox=\"0 0 32 16\">"
            + "<rect x=\"0\" y=\"0\" width=\"32\" height=\"16\"/></svg>";
        var result = SvgVectorWmfRenderer.renderDetailed(svg.getBytes(StandardCharsets.UTF_8),24,12,0.75);
        // Source viewport 32x16 px plus 2px padding on all sides, uniform fit.
        // The baseline is y=12px, mapped to (12+2)*heightUnits/20.
        assertEquals(12*(1-14d/20),result.baselineDepthPt(),0.03);
    }
}
