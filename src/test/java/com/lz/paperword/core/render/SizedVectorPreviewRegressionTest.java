package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;

import static com.lz.paperword.core.render.LaTeXImageRenderer.PreviewBackground.*;
import static org.junit.jupiter.api.Assertions.*;

class SizedVectorPreviewRegressionTest {
    private final LaTeXImageRenderer renderer = new LaTeXImageRenderer();

    @Test
    void fontSizeAndDisplayStyleReachMathJaxBeforeOutlining() throws Exception {
        String formula = "\\frac{a+2}{b+3}";
        var inline = renderer.renderMathJaxSvgForAcceptance(formula, 12, false);
        var display = renderer.renderMathJaxSvgForAcceptance(formula, 12, true);
        var large = renderer.renderMathJaxSvgForAcceptance(formula, 24, true);
        assertTrue(display.heightPt() > inline.heightPt() * 1.2, "display fractions must use real display layout");
        // MathJax gets the requested font size; its explicit 0.75pt margins do not scale with font size.
        assertEquals((display.widthPt() - 1.5) * 2, large.widthPt() - 1.5, 0.02);
        assertEquals((display.heightPt() - 1.5) * 2, large.heightPt() - 1.5, 0.02);
        var preview = renderer.renderForOlePreviewAtSize(formula, 12, true);
        assertTrue(preview.heightPt() > renderer.renderForOlePreviewAtSize(formula, 12, false).heightPt());
        assertPhysicalFrame(preview);
    }

    @Test
    void naturalLongLinearFormulaHasOnlyAnAspectPreserving440PointCap() {
        String medium = "12345+23456+34567+45678+56789+67890+78901";
        var natural = renderer.renderForOlePreviewAtSize(medium, 12, false);
        assertTrue(natural.widthPt() > 200, "sized mode must not use the old 200pt linear cap");
        assertTrue(natural.widthPt() <= 440);
        String veryLong = medium + "+" + medium + "+" + medium;
        var wide = renderer.renderForOlePreviewAtSize(veryLong, 12, false, 440);
        var narrow = renderer.renderForOlePreviewAtSize(veryLong, 12, false, 220);
        assertTrue(wide.widthPt() > 439.9 && wide.widthPt() <= 440);
        assertTrue(narrow.widthPt() > 219.9 && narrow.widthPt() <= 220);
        assertEquals(0.5, narrow.heightPt() / wide.heightPt(), 0.01, "no separate height rounding/stretch");
        assertPhysicalFrame(wide);
        assertPhysicalFrame(narrow);
    }

    @Test
    void physicalFrameIsDerivedFromActualBoundsAndNeverCropsInk() throws Exception {
        // Sparse SVG viewport: the 12x6pt painted rect, not its 60x30pt viewport, determines the frame.
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"60pt\" height=\"30pt\" viewBox=\"0 0 80 40\">"
            + "<rect x=\"8\" y=\"8\" width=\"16\" height=\"8\"/></svg>";
        var natural = SvgVectorWmfRenderer.renderNaturalDetailed(svg.getBytes(StandardCharsets.UTF_8),
            60, 0.4, 440, 0.75, TRANSPARENT);
        assertEquals(13.5, natural.widthPt(), 0.03);
        assertEquals(7.5, natural.heightPt(), 0.03);
        var inspection = WmfPreviewInspector.inspect(natural.vector().bytes());
        assertTrue(inspection.pureVector());
        assertFalse(inspection.inkTouchesEdge());
        assertEquals(natural.widthPt(), inspection.physicalWidth() * 72d / inspection.unitsPerInch(), 1e-10);
        assertEquals(natural.heightPt(), inspection.physicalHeight() * 72d / inspection.unitsPerInch(), 1e-10);
    }

    @Test
    void symbolsAndRaisedInkKeepKnownBaselineInsteadOfLegacyGuessing() {
        for (String formula : List.of("=", "+", "\\raisebox{3pt}{$x$}", "x_i^2", "\\sqrt{x+1}")) {
            var preview = renderer.renderForOlePreviewAtSize(formula, 12, false);
            assertTrue(Double.isFinite(preview.depthPt()) && preview.depthPt() >= 0, formula);
            assertFalse(WmfPreviewInspector.inspect(preview.data()).inkTouchesEdge(), formula);
            assertPhysicalFrame(preview);
        }
    }

    @Test
    void completeWhiteCanvasAddsExactlyOneUnstrokedPolygonWithoutChangingGlyphs() {
        String formula = "\\frac{a+5}{b+7}";
        var transparent = renderer.renderForOlePreviewAtSize(formula, 12, true, 440, TRANSPARENT);
        var white = renderer.renderForOlePreviewAtSize(formula, 12, true, 440, WHITE);
        assertEquals(transparent.widthPt(), white.widthPt());
        assertEquals(transparent.heightPt(), white.heightPt());
        assertEquals(transparent.depthPt(), white.depthPt());
        List<byte[]> a = polygonRecords(transparent.data());
        List<byte[]> b = polygonRecords(white.data());
        assertEquals(a.size() + 1, b.size());
        for (int i = 0; i < a.size(); i++) assertArrayEquals(a.get(i), b.get(i + 1));
        var whiteInspection = WmfPreviewInspector.inspect(white.data());
        var transparentInspection = WmfPreviewInspector.inspect(transparent.data());
        assertTrue(whiteInspection.pureVector());
        assertFalse(whiteInspection.inkTouchesEdge(), "white frame cannot count as formula ink");
        assertEquals(transparentInspection.foregroundPixels(), whiteInspection.foregroundPixels(),
            transparentInspection.foregroundPixels() * 0.04);
        var aImage = WmfPreviewInspector.rasterize(transparent.data()).image();
        var bImage = WmfPreviewInspector.rasterize(white.data()).image();
        assertEquals(0, aImage.getRGB(0, 0) >>> 24);
        assertEquals(0xFFFFFFFF, bImage.getRGB(0, 0));
        // RECTANGLE is deliberately absent: the white frame uses the same fill-only POLYPOLYGON route.
        ByteBuffer data = ByteBuffer.wrap(white.data()).order(ByteOrder.LITTLE_ENDIAN);
        for (int offset = 40; offset + 6 <= white.data().length;) {
            int function = Short.toUnsignedInt(data.getShort(offset + 4));
            if (function == 0x02FA) assertEquals(5, Short.toUnsignedInt(data.getShort(offset + 6)), "PS_NULL");
            assertNotEquals(0x041B, function);
            offset += data.getInt(offset) * 2;
        }
    }

    @Test
    void sizedInputsRejectSourceMetricsAndNonFiniteGeometryImmediately() {
        assertThrows(IllegalArgumentException.class, () -> renderer.renderForOlePreviewAtSize("\\pwmetrics{10,10}x", 12, false));
        for (double size : new double[]{Double.NaN, Double.POSITIVE_INFINITY, 0, 0.1, -12, 513}) {
            assertThrows(IllegalArgumentException.class, () -> renderer.renderForOlePreviewAtSize("x", size, false));
        }
        assertThrows(IllegalArgumentException.class, () -> renderer.renderForOlePreviewAtSize("x", 12, false, Double.NaN));
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class, () ->
            SvgVectorWmfRenderer.renderNaturalDetailed(("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"12pt\" height=\"12pt\">"
                + "<rect width=\"8\" height=\"8\"/></svg>").getBytes(StandardCharsets.UTF_8), 12, Double.NaN, 440, 0.75, TRANSPARENT));
    }

    @Test
    void sizedSingleFlightIncludesFontStyleGeometryAndBackground() throws Exception {
        String formula = "\\frac{x+" + Long.toUnsignedString(System.nanoTime()) + "}{y+1}";
        long before = requestCount();
        var gate = new CountDownLatch(1);
        try (var pool = Executors.newFixedThreadPool(6)) {
            var futures = new ArrayList<Future<LaTeXImageRenderer.PreviewImage>>();
            for (int i = 0; i < 6; i++) futures.add(pool.submit(() -> {
                gate.await();
                return renderer.renderForOlePreviewAtSize(formula, 12, true, 440, TRANSPARENT);
            }));
            gate.countDown();
            byte[] first = futures.getFirst().get(20, TimeUnit.SECONDS).data();
            for (var result : futures) assertArrayEquals(first, result.get(20, TimeUnit.SECONDS).data());
        }
        assertEquals(1, requestCount() - before);
        renderer.renderForOlePreviewAtSize(formula, 12, false, 440, TRANSPARENT);
        renderer.renderForOlePreviewAtSize(formula, 11, true, 440, TRANSPARENT);
        renderer.renderForOlePreviewAtSize(formula, 12, true, 200, TRANSPARENT);
        renderer.renderForOlePreviewAtSize(formula, 12, true, 440, WHITE);
        assertEquals(5, requestCount() - before);
        renderer.renderForOlePreviewAtSize(formula, 12, true, 440, TRANSPARENT);
        assertEquals(5, requestCount() - before);
    }

    @Test
    void legacyTargetCacheDoesNotRoundDistinctGeometryIntoOneKey() {
        var first = renderer.renderForOlePreview("x+1", 20.001, 10.001);
        var second = renderer.renderForOlePreview("x+1", 20.002, 10.002);
        assertEquals(20.001, first.widthPt());
        assertEquals(20.002, second.widthPt());
        var image = WmfPreviewInspector.rasterize(renderer.renderForOlePreview("x+1").data()).image();
        assertEquals(0, image.getRGB(0, 0) >>> 24, "legacy API stays transparent");
    }

    private static void assertPhysicalFrame(LaTeXImageRenderer.PreviewImage preview) {
        var inspection = WmfPreviewInspector.inspect(preview.data());
        assertTrue(inspection.pureVector(), inspection.error());
        assertEquals(preview.widthPt(), inspection.physicalWidth() * 72d / inspection.unitsPerInch(), 1e-10);
        assertEquals(preview.heightPt(), inspection.physicalHeight() * 72d / inspection.unitsPerInch(), 1e-10);
    }

    private static long requestCount() throws Exception {
        Field field = LaTeXImageRenderer.class.getDeclaredField("mathJaxRequestId");
        field.setAccessible(true);
        return field.getLong(null);
    }

    private static List<byte[]> polygonRecords(byte[] bytes) {
        ByteBuffer data = ByteBuffer.wrap(bytes).order(ByteOrder.LITTLE_ENDIAN);
        List<byte[]> result = new ArrayList<>();
        for (int offset = 40; offset + 6 <= bytes.length;) {
            int size = data.getInt(offset) * 2;
            if (Short.toUnsignedInt(data.getShort(offset + 4)) == 0x0538) {
                result.add(Arrays.copyOfRange(bytes, offset, offset + size));
            }
            offset += size;
        }
        return result;
    }
}
