package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LaTeXImageRendererTest {

    @Test
    void shouldParseSvgPtDimensions() {
        byte[] svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="36pt" height="18pt" viewBox="0 0 36 18">
              <path d="M0 0h36v18H0z"/>
            </svg>
            """.getBytes(StandardCharsets.UTF_8);

        LaTeXImageRenderer.SvgDimensions dimensions = LaTeXImageRenderer.extractSvgDisplayDimensions(svg);
        assertEquals(36.0f, dimensions.widthPt(), 0.01f);
        assertEquals(18.0f, dimensions.heightPt(), 0.01f);
    }

    @Test
    void shouldParseSvgPxDimensions() {
        byte[] svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="40px" height="20px" viewBox="0 0 40 20">
              <path d="M0 0h40v20H0z"/>
            </svg>
            """.getBytes(StandardCharsets.UTF_8);

        LaTeXImageRenderer.SvgDimensions dimensions = LaTeXImageRenderer.extractSvgDisplayDimensions(svg);
        assertEquals(30.0f, dimensions.widthPt(), 0.05f);
        assertEquals(15.0f, dimensions.heightPt(), 0.05f);
    }

    @Test
    void shouldKeepLongDivisionPreviewHeaderOnly() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(), "\\longdiv[570]{6}{3420}");
        assertTrue(normalized.contains("\\overset{570}{\\overline{\\left)3420\\right.}}"));
        assertFalse(normalized.contains("\\underline{30}"), "预览图不应再自动补第一步乘积");
        assertFalse(normalized.contains("\\underline{42}"), "预览图不应再自动补后续步骤");
    }

    @Test
    void shouldRemoveStandaloneAlignmentMarkerBeforeLocalRender() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(
            new LaTeXImageRenderer(), "&=20.08\\times (200.9-200.7)");

        assertEquals("=20.08\\times (200.9-200.7)", normalized);
    }

    @Test
    void shouldPlaceImplicitLimitArgumentUnderOperatorForMathTypePreview() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod(
            "normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        assertEquals("\\lim\\limits_{x\\to0}\\frac{\\sin x}{x}=1",
            method.invoke(renderer, "\\lim_{x\\to0}\\frac{\\sin x}{x}=1"));
        assertEquals("\\lim\\limits_{x\\to0}",
            method.invoke(renderer, "\\lim\\limits_{x\\to0}"));
        assertEquals("\\lim\\nolimits_{x\\to0}",
            method.invoke(renderer, "\\lim\\nolimits_{x\\to0}"));

        String svg = new String(renderer.renderMathJaxSvgForAcceptance(
            "\\lim_{x\\to0}\\frac{\\sin x}{x}=1").svgBytes(), StandardCharsets.UTF_8);
        assertTrue(svg.contains("data-mml-node=\"munderover\"")
                || svg.contains("data-mml-node=\"munder\""),
            "the serialized MathJax scene must keep x->0 below lim");
    }

    @Test
    void shouldNormalizeLegacyUnbracedBbbToReadableGlyph() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        assertEquals("x", method.invoke(new LaTeXImageRenderer(), "\\Bbb x"));
        assertEquals("\\Bbb{x}", method.invoke(new LaTeXImageRenderer(), "\\Bbb{x}"));
    }

    @Test
    void shouldSeparateNestedRadicalDegreeBeforeLocalRender() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(),
            "\\sqrt[1+\\sqrt[2]{3}+4]{5}-\\sqrt[6]{7}");

        assertEquals("{}^{1+\\sqrt[2]{3}+4}\\!\\sqrt{5}-\\sqrt[6]{7}", normalized);
    }

    @Test
    void shouldTranslateRaiseBoxToMathJaxRaiseWithoutLosingNestedContent() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(),
            "\\raisebox{-3pt}{${\\times}$6+\\raisebox{1pt}{a\\}b\\{c}}");

        assertEquals("\\raise{-3pt}{{\\times}6+\\raise{1pt}{a\\}b\\{c}}", normalized);
        LaTeXImageRenderer.PreviewImage preview =
            new LaTeXImageRenderer().renderForOlePreview("\\raisebox{-3pt}{2}");
        assertFalse(preview.placeholder());
        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(preview.data()));
    }

    @Test
    void shouldRejectRaiseBoxMetricOverridesThatMathJaxCannotPreserve() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        InvocationTargetException failure = assertThrows(InvocationTargetException.class,
            () -> method.invoke(new LaTeXImageRenderer(), "\\raisebox{-3pt}[8pt][2pt]{x}"));

        assertTrue(failure.getCause() instanceof IllegalArgumentException);
        assertTrue(failure.getCause().getMessage().contains("optional height/depth"));

        InvocationTargetException unbalanced = assertThrows(InvocationTargetException.class,
            () -> method.invoke(new LaTeXImageRenderer(), "\\raisebox{-3pt}{$x}"));
        assertTrue(unbalanced.getCause() instanceof IllegalArgumentException);
        assertTrue(unbalanced.getCause().getMessage().contains("Unbalanced $ delimiter"));
    }

    @Test
    void shouldRenderLegacyUnbracedBbbAsReadableDoubleStruckGlyph() {
        LaTeXImageRenderer.PreviewImage preview =
            new LaTeXImageRenderer().renderForOlePreview("\\Bbb x");

        assertEquals("emf", preview.extension());
        assertEquals("image/x-emf", preview.contentType());
        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(preview.data()));
    }

    @Test
    void shouldHoistStyleWrappedArrayAlignmentMarkerBeforeLocalRender() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(),
            "\\begin{array}{l}21x\\mathbf{&=}140\\\\20x\\mathbf{&=}65\\end{array}");

        assertFalse(normalized.contains("\\mathbf{&="));
        assertTrue(normalized.contains("x&\\mathbf{=}"));
    }

    @Test
    void shouldReplaceCompositeLongDivisionHeaderInsideSingleBlock() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(
            new LaTeXImageRenderer(),
            "\\longdiv[570]{6}{3420}\\begin{array}{l}\\text{   }\\underline{30}\\\\\\text{    }42\\end{array}"
        );
        assertFalse(normalized.contains("\\longdiv[570]{6}{3420}"), "单块复合长除法预览时应替换掉原始 longdiv 命令");
        assertTrue(normalized.contains("\\overset{570}{\\overline{\\left)3420\\right.}}"), "预览图应保留长除法头部");
        assertTrue(normalized.contains("\\begin{array}{l}"), "单块复合长除法的步骤区应继续保留");
        assertTrue(normalized.contains("\\underline{30}"), "显式步骤区应继续进入预览渲染");
    }

    @Test
    void shouldEscapeRawUnicodeSymbolsForMathJaxWithoutChangingText() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(),
            "\\begin{array}{r}2.30€\\\\中文𝛼\\end{array}");

        assertEquals("\\begin{array}{r}2.30\\unicode{x20AC}\\\\中文\\unicode{x1D6FC}\\end{array}", normalized);
    }

    @Test
    void shouldRepairOnlyKnownDocx2texArtifactsAndRejectUnknownReplacementCharacters() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        InvocationTargetException replacementFailure = assertThrows(InvocationTargetException.class,
            () -> method.invoke(new LaTeXImageRenderer(), "(75+60)�\\times20=�2700"));
        assertTrue(replacementFailure.getCause() instanceof IllegalArgumentException);
        assertTrue(replacementFailure.getCause().getMessage().contains("SOURCE_REPLACEMENT_CHARACTER"));
        assertEquals("\\text{相遇时间}+\\text{追及时间}", method.invoke(new LaTeXImageRenderer(),
            "\\text{相遇{\\blacksquare}{\\blacksquare}}+\\text{追及{\\blacksquare}{\\blacksquare}}"));
        assertEquals("2=64（cm^{2}", method.invoke(new LaTeXImageRenderer(), "2=64（cm^{2"));
        assertEquals("\\text{心想事成}", method.invoke(new LaTeXImageRenderer(), "\\text{心想事\\Theta }"));
        assertEquals("\\text{梦想成真}",
            method.invoke(new LaTeXImageRenderer(), "\\text{梦想\\Theta 真}"));
        assertEquals("P_{3}^{1}\\cdot P_{5}^{1}",
            method.invoke(new LaTeXImageRenderer(), "P_{3}^{1}\\spot P_{5}^{1}"));
        assertEquals("\\cdot_1+\\cdot2+\\cdot+\\spotlight+\\spot甲",
            method.invoke(new LaTeXImageRenderer(), "\\spot_1+\\spot2+\\spot+\\spotlight+\\spot甲"));
        assertEquals("\\text{不合题意}",
            method.invoke(new LaTeXImageRenderer(), "\\text{不合{\\blacksquare}意}"));
        assertEquals("x+\\blacksquare", method.invoke(new LaTeXImageRenderer(), "x+\\blacksquare"));
    }

    @Test
    void shouldRenderEuroArrayAsEditableOlePreview() {
        LaTeXImageRenderer.PreviewImage preview = new LaTeXImageRenderer().renderForOlePreview(
            "\\begin{array}{r}2.30€\\\\55.60€\\\\1001.00€\\end{array}");

        assertEquals("emf", preview.extension());
        assertFalse(preview.placeholder());
        assertTrue(preview.data().length > 0);
        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(preview.data()));
    }

    @Test
    void shouldWriteRenderedFormulaToDiskCache(@TempDir Path tempDir) throws Exception {
        String oldEnabled = System.getProperty("paperword.render.cache.enabled");
        String oldDir = System.getProperty("paperword.render.cache.dir");
        try {
            System.setProperty("paperword.render.cache.enabled", "true");
            System.setProperty("paperword.render.cache.dir", tempDir.toString());

            byte[] png = new LaTeXImageRenderer().renderToPng("x+" + UUID.randomUUID(), 13f);

            assertTrue(png.length > 0);
            try (var stream = Files.walk(tempDir)) {
                assertTrue(stream.anyMatch(path -> path.getFileName().toString().endsWith(".png")),
                    "rendered formulas should be written to the disk cache");
            }
        } finally {
            restoreProperty("paperword.render.cache.enabled", oldEnabled);
            restoreProperty("paperword.render.cache.dir", oldDir);
        }
    }

    @Test
    void cacheKeyIncludesCurrentVectorRenderProperties() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("cacheKey", String.class, String.class, float.class);
        method.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        String oldPadding = System.getProperty("paperword.mathjax.paddingPt");
        try {
            System.setProperty("paperword.mathjax.paddingPt", "2.30");
            String defaultKey = (String) method.invoke(renderer, "ole-target-10.00x10.00", "x+1", 12f);
            System.setProperty("paperword.mathjax.paddingPt", "3.25");
            String tunedKey = (String) method.invoke(renderer, "ole-target-10.00x10.00", "x+1", 12f);

            assertFalse(defaultKey.equals(tunedKey), "MathJax vector geometry must invalidate preview cache keys");
        } finally {
            restoreProperty("paperword.mathjax.paddingPt", oldPadding);
        }
    }

    @Test
    void legacyBitmapVectorToggleCannotChangeStrictVectorCacheIdentity() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("cacheKey", String.class, String.class, float.class);
        method.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        String oldVector = System.getProperty("paperword.ole.preview.vectorWmf");
        try {
            System.setProperty("paperword.ole.preview.vectorWmf", "false");
            String bitmap = (String) method.invoke(renderer, "ole", "x+1", 12f);
            System.setProperty("paperword.ole.preview.vectorWmf", "true");
            String vector = (String) method.invoke(renderer, "ole", "x+1", 12f);

            assertEquals(bitmap, vector);
        } finally {
            restoreProperty("paperword.ole.preview.vectorWmf", oldVector);
        }
    }

    @Test
    void legacyEmfToggleCannotBypassStrictVectorWmfBackend() {
        String oldEmf = System.getProperty("paperword.ole.preview.emf");
        String oldFormat = System.getProperty("paperword.ole.previewFormat");
        try {
            System.setProperty("paperword.ole.preview.emf", "true");
            System.setProperty("paperword.ole.previewFormat", "wmf");

            LaTeXImageRenderer.PreviewImage preview =
                new LaTeXImageRenderer().renderForOlePreview("\\sqrt{a^{2}+b^{2}}", 56.0d, 21.0d);

            assertEquals("wmf", preview.extension());
            assertEquals("image/x-wmf", preview.contentType());
            assertTrue(preview.data().length > 0);
            assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(preview.data()));
            assertTrue(WmfPreviewInspector.inspect(preview.data()).pureVector());
            assertEquals(56.0d, preview.widthPt(), 0.01d);
            assertEquals(21.0d, preview.heightPt(), 0.01d);
        } finally {
            restoreProperty("paperword.ole.preview.emf", oldEmf);
            restoreProperty("paperword.ole.previewFormat", oldFormat);
        }
    }

    @Test
    void vectorHeightEstimatesUseMathTypeStructureMetrics() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("estimateVectorHeightPt", String.class);
        method.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        assertEquals(MathTypeStructureMetrics.LINEAR_HEIGHT_PT, (double) method.invoke(renderer, "x+1"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SCRIPT_HEIGHT_PT, (double) method.invoke(renderer, "a^{2}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT,
            (double) method.invoke(renderer, "x^{\\frac{1}{2}}+a^{\\frac{2}{3}}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SCRIPT_FRACTION_MIXED_HEIGHT_PT,
            (double) method.invoke(renderer, "S_{\\frac{1}{4}\\mathrm{圆}}=\\frac{1}{4}\\pi r^{2}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.ORDINARY_FRACTION_HEIGHT_PT,
            (double) method.invoke(renderer, "\\frac{1}{2}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT,
            (double) method.invoke(renderer, "\\frac{1+\\frac{a}{b}}{2}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SQRT_HEIGHT_PT,
            (double) method.invoke(renderer, "\\sqrt{2}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT,
            (double) method.invoke(renderer, "\\sqrt{1+\\frac{a}{b}}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT,
            (double) method.invoke(renderer, "\\sqrt{1+\\sqrt{\\frac{a}{b}}}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT,
            (double) method.invoke(renderer, "\\sqrt{1+\\sqrt{\\frac{a}{b}+\\sqrt{z}}}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT,
            (double) method.invoke(renderer, "\\frac{\\sqrt{a^{2}+b^{2}}}{2}"), 0.01d);
        assertEquals(MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT,
            (double) method.invoke(renderer, "\\frac{三角形ABD的面积}{三角形CBD的面积}"), 0.01d);
    }

    @Test
    void sourceSeededStructureHeightsAreNotScaledTwice() throws Exception {
        Method estimate = LaTeXImageRenderer.class.getDeclaredMethod("estimateVectorHeightPt", String.class);
        estimate.setAccessible(true);
        Method calibrate = LaTeXImageRenderer.class.getDeclaredMethod(
            "calibratePreviewMetrics", String.class, double.class, double.class, double.class);
        calibrate.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        assertCalibratedHeight(renderer, estimate, calibrate, "\\frac{1}{2}",
            MathTypeStructureMetrics.ORDINARY_FRACTION_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\frac{1+\\frac{a}{b}}{2}",
            MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\frac{三角形ABD的面积}{三角形CBD的面积}",
            MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\sqrt{2}",
            MathTypeStructureMetrics.SQRT_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\sqrt{1+\\frac{a}{b}}",
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\sqrt{1+\\sqrt{\\frac{a}{b}}}",
            MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\sqrt{1+\\sqrt{\\frac{a}{b}+\\sqrt{z}}}",
            MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\frac{\\sqrt{a^{2}+b^{2}}}{2}",
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "x^{\\frac{1}{2}}+a^{\\frac{2}{3}}",
            MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate,
            "S_{\\frac{1}{4}\\mathrm{圆}}=\\frac{1}{4}\\pi r^{2}",
            MathTypeStructureMetrics.SCRIPT_FRACTION_MIXED_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\begin{array}{c}1\\\\2\\end{array}",
            MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\begin{cases}1\\\\2\\end{cases}",
            MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT);
        assertCalibratedHeight(renderer, estimate, calibrate, "\\overline{AB}",
            MathTypeStructureMetrics.ACCENT_HEIGHT_PT);
    }

    @Test
    void structureFamilyQueryIsTheSingleMetricsContract() throws Exception {
        Method family = LaTeXImageRenderer.class.getDeclaredMethod("classifyStructureFamily", String.class);
        family.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        assertFamily(renderer, family, "x+1", MathTypeStructureMetrics.Family.LINEAR,
            MathTypeStructureMetrics.LINEAR_HEIGHT_PT, "linear");
        assertFamily(renderer, family, "a^{2}", MathTypeStructureMetrics.Family.SCRIPT,
            MathTypeStructureMetrics.SCRIPT_HEIGHT_PT, "script");
        assertFamily(renderer, family, "x^{\\frac{1}{2}}+a^{\\frac{2}{3}}",
            MathTypeStructureMetrics.Family.SCRIPT_FRACTION,
            MathTypeStructureMetrics.SCRIPT_FRACTION_HEIGHT_PT, "script_fraction");
        assertFamily(renderer, family, "S_{\\frac{1}{4}\\mathrm{圆}}=\\frac{1}{4}\\pi r^{2}",
            MathTypeStructureMetrics.Family.SCRIPT_FRACTION_MIXED,
            MathTypeStructureMetrics.SCRIPT_FRACTION_MIXED_HEIGHT_PT, "script_fraction_mixed");
        assertFamily(renderer, family, "\\frac{1}{2}", MathTypeStructureMetrics.Family.ORDINARY_FRACTION,
            MathTypeStructureMetrics.ORDINARY_FRACTION_HEIGHT_PT, "fraction");
        assertFamily(renderer, family, "\\frac{1+\\frac{a}{b}}{2}",
            MathTypeStructureMetrics.Family.NESTED_FRACTION,
            MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT, "nested_fraction");
        assertFamily(renderer, family, "\\frac{三角形ABD的面积}{三角形CBD的面积}",
            MathTypeStructureMetrics.Family.TEXT_FRACTION,
            MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT, "text_fraction");
        assertFamily(renderer, family, "\\frac{三角形ABD的面积}{三角形CBD的面积}=\\frac{AO}{CO}",
            MathTypeStructureMetrics.Family.TEXT_FRACTION,
            MathTypeStructureMetrics.TEXT_FRACTION_HEIGHT_PT, "text_fraction");
        assertFamily(renderer, family, "\\sqrt{2}", MathTypeStructureMetrics.Family.SQRT,
            MathTypeStructureMetrics.SQRT_HEIGHT_PT, "sqrt");
        assertFamily(renderer, family, "\\sqrt{1+\\sqrt{x}}", MathTypeStructureMetrics.Family.SQRT,
            MathTypeStructureMetrics.SQRT_HEIGHT_PT, "sqrt");
        assertFamily(renderer, family, "\\sqrt{x+\\sqrt{y+\\sqrt{z}}}",
            MathTypeStructureMetrics.Family.SQRT_NESTED,
            MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT, "sqrt_nested");
        assertFamily(renderer, family, "\\sqrt{1+\\frac{a}{b}}",
            MathTypeStructureMetrics.Family.SQRT_FRACTION,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT, "sqrt_fraction");
        assertFamily(renderer, family, "\\sqrt{1+\\sqrt{\\frac{a}{b}}}",
            MathTypeStructureMetrics.Family.SQRT_NESTED,
            MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT, "sqrt_nested");
        assertFamily(renderer, family, "\\sqrt{1+\\sqrt{\\frac{a}{b}+\\sqrt{z}}}",
            MathTypeStructureMetrics.Family.SQRT_NESTED,
            MathTypeStructureMetrics.SQRT_NESTED_HEIGHT_PT, "sqrt_nested");
        assertFamily(renderer, family, "\\frac{\\sqrt{a^{2}+b^{2}}}{2}",
            MathTypeStructureMetrics.Family.SQRT_FRACTION,
            MathTypeStructureMetrics.SQRT_FRACTION_HEIGHT_PT, "sqrt_fraction");
        assertFamily(renderer, family, "\\begin{array}{c}1\\\\2\\end{array}",
            MathTypeStructureMetrics.Family.ARRAY,
            MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT, "array");
        assertFamily(renderer, family, "\\begin{aligned} V&=abh\\\\ V&=Sh \\end{aligned}",
            MathTypeStructureMetrics.Family.ARRAY,
            MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT, "array");
        assertFamily(renderer, family,
            "\\begin{cases} \\frac{x}{y}=\\frac{7}{3}\\\\ \\frac{x+70}{y+70}=\\frac{7}{4} \\end{cases}",
            MathTypeStructureMetrics.Family.ARRAY,
            MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT, "array");
        assertFamily(renderer, family, "\\overline{AB}", MathTypeStructureMetrics.Family.ACCENT,
            MathTypeStructureMetrics.ACCENT_HEIGHT_PT, "accent");
    }

    @Test
    void sourceSampleMetricsExposeInkEvidenceWithoutChangingHeightContract() {
        MathTypeStructureMetrics.SourceSampleMetrics script =
            MathTypeStructureMetrics.sourceSampleMetrics(MathTypeStructureMetrics.Family.SCRIPT);
        MathTypeStructureMetrics.SourceSampleMetrics fraction =
            MathTypeStructureMetrics.sourceSampleMetrics(MathTypeStructureMetrics.Family.ORDINARY_FRACTION);
        MathTypeStructureMetrics.SourceSampleMetrics nestedFraction =
            MathTypeStructureMetrics.sourceSampleMetrics(MathTypeStructureMetrics.Family.NESTED_FRACTION);
        MathTypeStructureMetrics.SourceSampleMetrics sqrt =
            MathTypeStructureMetrics.sourceSampleMetrics(MathTypeStructureMetrics.Family.SQRT);
        MathTypeStructureMetrics.SourceSampleMetrics linear =
            MathTypeStructureMetrics.sourceSampleMetrics(MathTypeStructureMetrics.Family.LINEAR);

        assertEquals(MathTypeStructureMetrics.SCRIPT_CANDIDATE_HEIGHT_PT, script.candidateHeightPt(), 0.01d);
        assertEquals(MathTypeStructureMetrics.SCRIPT_HEIGHT_PT,
            MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.SCRIPT).heightPt(), 0.01d);
        assertEquals(0.922d, script.ink().widthRatio(), 0.001d);
        assertEquals(0.711d, script.ink().heightRatio(), 0.001d);
        assertEquals(0.487d, script.ink().centerYRatio(), 0.001d);
        assertEquals(3, script.ink().sampleCount());
        assertTrue(script.hasInkSamples());

        assertEquals(MathTypeStructureMetrics.ORDINARY_FRACTION_CANDIDATE_HEIGHT_PT,
            fraction.candidateHeightPt(), 0.01d);
        assertEquals(MathTypeStructureMetrics.ORDINARY_FRACTION_HEIGHT_PT,
            MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.ORDINARY_FRACTION).heightPt(), 0.01d);
        assertEquals(0.889d, fraction.ink().widthRatio(), 0.001d);
        assertEquals(4, fraction.ink().sampleCount());

        assertEquals(MathTypeStructureMetrics.NESTED_FRACTION_CANDIDATE_HEIGHT_PT,
            nestedFraction.candidateHeightPt(), 0.01d);
        assertEquals(MathTypeStructureMetrics.NESTED_FRACTION_HEIGHT_PT,
            MathTypeStructureMetrics.metrics(MathTypeStructureMetrics.Family.NESTED_FRACTION).heightPt(), 0.01d);
        assertEquals(0.911d, nestedFraction.ink().widthRatio(), 0.001d);
        assertEquals(2, nestedFraction.ink().sampleCount());

        assertEquals(MathTypeStructureMetrics.SQRT_HEIGHT_PT, sqrt.candidateHeightPt(), 0.01d);
        assertEquals(0.750d, sqrt.ink().heightRatio(), 0.001d);
        assertEquals(5, sqrt.ink().sampleCount());

        assertFalse(linear.hasInkSamples());
        assertEquals(MathTypeStructureMetrics.LINEAR_HEIGHT_PT, linear.candidateHeightPt(), 0.01d);
    }

    @Test
    void previewWidthCalibrationUsesStructureMetricsForLinearAndRoots() throws Exception {
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        Method estimate = LaTeXImageRenderer.class.getDeclaredMethod("estimateVectorHeightPt", String.class);
        Method calibrate = LaTeXImageRenderer.class.getDeclaredMethod("calibratePreviewMetrics", String.class,
            double.class, double.class, double.class);
        estimate.setAccessible(true);
        calibrate.setAccessible(true);

        assertCalibratedWidth(renderer, estimate, calibrate, "a+b=c", 40.0d,
            40.0d * 1.08d * MathTypeStructureMetrics.LINEAR_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\sqrt{x+1}", 40.0d,
            40.0d * MathTypeStructureMetrics.SQRT_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\sqrt{a^{2}+b^{2}}", 40.0d,
            40.0d * 0.96d * MathTypeStructureMetrics.SQRT_SCRIPT_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\sqrt{1+\\frac{a}{b}}", 40.0d,
            40.0d * 0.96d * MathTypeStructureMetrics.SQRT_FRACTION_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "T=2\\pi\\sqrt{\\frac{l}{g}}", 40.0d,
            40.0d * 0.96d * MathTypeStructureMetrics.SQRT_FRACTION_MIXED_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\frac{\\sqrt{a^{2}+b^{2}}}{2}", 40.0d,
            40.0d * 0.96d * MathTypeStructureMetrics.SQRT_FRACTION_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\sqrt{x}+\\sqrt{\\frac{a}{b}}", 40.0d,
            40.0d * 0.96d * MathTypeStructureMetrics.SQRT_FRACTION_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\sqrt{1+\\sqrt{\\frac{a}{b}}}", 40.0d,
            40.0d * 0.96d * MathTypeStructureMetrics.SQRT_NESTED_FRACTION_PREVIEW_WIDTH_SCALE);
        assertCalibratedWidth(renderer, estimate, calibrate, "\\sqrt{x+\\sqrt{y+\\sqrt{z}}}", 40.0d,
            40.0d * 0.86d);
    }

    private static void assertCalibratedHeight(LaTeXImageRenderer renderer, Method estimate, Method calibrate,
        String latex, double expectedHeightPt) throws Exception {
        double estimatedHeight = (double) estimate.invoke(renderer, latex);
        Object metrics = calibrate.invoke(renderer, latex, 40.0d, estimatedHeight, 4.0d);
        Method heightPt = metrics.getClass().getDeclaredMethod("heightPt");
        heightPt.setAccessible(true);

        assertEquals(expectedHeightPt, estimatedHeight, 0.01d);
        assertEquals(expectedHeightPt, (double) heightPt.invoke(metrics), 0.01d);
    }

    private static void assertCalibratedWidth(LaTeXImageRenderer renderer, Method estimate, Method calibrate,
        String latex, double inputWidthPt, double expectedWidthPt) throws Exception {
        double estimatedHeight = (double) estimate.invoke(renderer, latex);
        Object metrics = calibrate.invoke(renderer, latex, inputWidthPt, estimatedHeight, 4.0d);
        Method widthPt = metrics.getClass().getDeclaredMethod("widthPt");
        widthPt.setAccessible(true);

        assertEquals(expectedWidthPt, (double) widthPt.invoke(metrics), 0.01d);
    }

    private static void assertFamily(LaTeXImageRenderer renderer, Method method, String latex,
        MathTypeStructureMetrics.Family expectedFamily, double expectedHeightPt, String expectedPreviewClass)
        throws Exception {
        Object metrics = method.invoke(renderer, latex);
        Method family = metrics.getClass().getDeclaredMethod("family");
        Method heightPt = metrics.getClass().getDeclaredMethod("heightPt");
        Method sourceSeededHeight = metrics.getClass().getDeclaredMethod("sourceSeededHeight");
        Method previewClass = metrics.getClass().getDeclaredMethod("previewClass");
        family.setAccessible(true);
        heightPt.setAccessible(true);
        sourceSeededHeight.setAccessible(true);
        previewClass.setAccessible(true);

        assertEquals(expectedFamily, family.invoke(metrics));
        assertEquals(expectedHeightPt, (double) heightPt.invoke(metrics), 0.01d);
        assertTrue((boolean) sourceSeededHeight.invoke(metrics));
        assertEquals(expectedPreviewClass, previewClass.invoke(metrics));
    }

    @Test
    void runCommandTimesOutWithoutWaitingForStreamClose(@TempDir Path tempDir) throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("runCommand", List.class, Path.class, int.class);
        method.setAccessible(true);
        String java = Path.of(System.getProperty("java.home"), "bin", isWindows() ? "java.exe" : "java").toString();
        long start = System.nanoTime();
        Object result = method.invoke(
            new LaTeXImageRenderer(),
            List.of(java, "-cp", System.getProperty("java.class.path"), HangingProcess.class.getName()),
            tempDir,
            1
        );
        long elapsedMs = (System.nanoTime() - start) / 1_000_000L;
        Method exitCode = result.getClass().getDeclaredMethod("exitCode");
        exitCode.setAccessible(true);

        assertEquals(-1, exitCode.invoke(result));
        assertTrue(elapsedMs < 5000, "runCommand timeout must not wait for the child stream to close");
    }

    public static class HangingProcess {
        public static void main(String[] args) throws Exception {
            System.out.print("started");
            System.out.flush();
            Thread.sleep(60_000L);
        }
    }

    private void restoreProperty(String name, String value) {
        if (value == null) {
            System.clearProperty(name);
        } else {
            System.setProperty(name, value);
        }
    }

    private boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase().contains("win");
    }
}
