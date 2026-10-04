package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.latex.LaTeXParser.FormulaMetrics;
import com.lz.paperword.core.latex.LaTeXParser.FormulaStyleHints;
import com.lz.paperword.core.mtef.MtefRecordNormalizer;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import com.lz.paperword.core.render.WmfPreviewInspector;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;

import static com.lz.paperword.core.render.LaTeXImageRenderer.PreviewBackground.*;
import static org.junit.jupiter.api.Assertions.*;

class SizedEquationEmbeddingRegressionTest {
    @Test
    void twelvePointPreviewPreservesOriginalNativeEquationBytesForAllStylesAndBackgrounds() throws Exception {
        for (String latex : List.of("x_i^2+1", "\\frac{a+3}{b+5}", "\\sqrt{x+7}")) {
            var legacy = embed(latex, null, false, TRANSPARENT);
            for (boolean display : new boolean[]{false, true}) {
                var transparent = embed(latex, 12d, display, TRANSPARENT);
                var white = embed(latex, 12d, display, WHITE);
                assertArrayEquals(legacy.nativeData(), transparent.nativeData(), "12pt uses the original native FULL definition");
                assertArrayEquals(transparent.nativeData(), white.nativeData(), "background cannot change editable MTEF");
                assertFalse(Arrays.equals(transparent.preview(), white.preview()));
                assertFrameMatchesShape(transparent);
                assertFrameMatchesShape(white);
            }
        }
    }

    @Test
    void otherRequestedPointSizeIsWrittenToEditableMtefWithoutChangingMathCharacters() throws Exception {
        String latex = "x+7";
        var at12 = embed(latex, 12d, false, TRANSPARENT);
        var at9 = embed(latex, 9.5d, false, TRANSPARENT);
        var mtef12 = mtef(at12.nativeData());
        var mtef9 = mtef(at9.nativeData());
        assertFalse(Arrays.equals(mtef12, mtef9), "a smaller picture alone is not sufficient");
        var a = MtefRecordNormalizer.normalize(mtef12);
        var b = MtefRecordNormalizer.normalize(mtef9);
        assertEquals(a.records().stream().filter(r -> "CHAR".equals(r.name())).toList(),
            b.records().stream().filter(r -> "CHAR".equals(r.name())).toList());
        // Absolute SIZE(101): 9.5pt in the format's 1/32pt units = 304 (0x0130).
        assertTrue(contains(mtef9, new byte[]{9, 101, 0x30, 0x01}), "editable MTEF must contain requested 9.5pt SIZE");
        assertTrue(width(at9) < width(at12));
        assertFrameMatchesShape(at9);
    }

    @Test
    void sizedSourceMetricsConflictIsAnArgumentErrorBeforeDocumentMutation() throws Exception {
        try (var document = new XWPFDocument()) {
            var paragraph = document.createParagraph();
            var run = paragraph.createRun();
            var ast = new LaTeXParser().parseDetailed("x").ast();
            var hints = FormulaStyleHints.empty().withSourceMetrics(new FormulaMetrics(20, 12));
            var embedder = new MathTypeEmbedder();
            assertThrows(IllegalArgumentException.class, () -> embedder.embedEquationAtSize(
                paragraph, run, ast, "x", 12, false, 440, hints));
            assertThrows(IllegalArgumentException.class, () -> embedder.embedEquationAtSize(
                paragraph, run, ast, "\\pwmetrics{20,12}x", 12, false, 440, FormulaStyleHints.empty()));
            assertEquals(0, run.getCTR().sizeOfObjectArray());
            assertTrue(document.getPackage().getParts().stream().noneMatch(p -> p.getPartName().getName().contains("oleObject")));
        }
    }

    @Test
    void symbolsUseGeometricBaselineAndDoNotFallBackToLegacyHeightGuess() throws Exception {
        for (String latex : List.of("=", "+", "\\raisebox{3pt}{$x$}")) {
            var result = embed(latex, 12d, false, TRANSPARENT);
            var preview = new LaTeXImageRenderer().renderForOlePreviewAtSize(latex, 12, false);
            var position = Pattern.compile("(?:w:)?position[^>]*(?:w:)?val=\"(-?\\d+)\"").matcher(result.xml());
            assertTrue(position.find());
            assertEquals(-(int) Math.round(preview.depthPt() * 2), Integer.parseInt(position.group(1)), latex);
            assertFrameMatchesShape(result);
        }
    }

    private static Embedded embed(String latex, Double pointSize, boolean display,
                                  LaTeXImageRenderer.PreviewBackground background) throws Exception {
        try (var document = new XWPFDocument()) {
            var paragraph = document.createParagraph();
            var run = paragraph.createRun();
            var embedder = new MathTypeEmbedder();
            var ast = new LaTeXParser().parseDetailed(latex).ast();
            if (pointSize == null) embedder.embedEquation(paragraph, run, ast, latex);
            else embedder.embedEquationAtSize(paragraph, run, ast, latex, pointSize, display, 440,
                FormulaStyleHints.empty(), background);
            byte[] preview = null, nativeData = null;
            for (var part : document.getPackage().getParts()) {
                String name = part.getPartName().getName();
                if (name.endsWith("image_eq1.wmf")) preview = part.getInputStream().readAllBytes();
                if (name.endsWith("oleObject1.bin")) {
                    try (var filesystem = new POIFSFileSystem(part.getInputStream());
                         var stream = filesystem.createDocumentInputStream("Equation Native")) {
                        nativeData = stream.readAllBytes();
                    }
                }
            }
            assertNotNull(preview);
            assertNotNull(nativeData);
            return new Embedded(nativeData, preview, run.getCTR().xmlText());
        }
    }

    private static byte[] mtef(byte[] nativeData) {
        int offset = ByteBuffer.wrap(nativeData).order(ByteOrder.LITTLE_ENDIAN).getInt();
        return Arrays.copyOfRange(nativeData, offset, nativeData.length);
    }

    private static boolean contains(byte[] data, byte[] expected) {
        for (int i = 0; i <= data.length - expected.length; i++) {
            boolean equal = true;
            for (int j = 0; j < expected.length; j++) equal &= data[i+j] == expected[j];
            if (equal) return true;
        }
        return false;
    }

    private static void assertFrameMatchesShape(Embedded embedded) {
        var inspection = WmfPreviewInspector.inspect(embedded.preview());
        assertTrue(inspection.pureVector(), inspection.error());
        assertFalse(inspection.inkTouchesEdge());
        assertEquals(inspection.physicalWidth() * 72d / inspection.unitsPerInch(), width(embedded), 0.00051);
        var height = Pattern.compile("height:([0-9.]+)pt").matcher(embedded.xml());
        assertTrue(height.find());
        assertEquals(inspection.physicalHeight() * 72d / inspection.unitsPerInch(), Double.parseDouble(height.group(1)), 0.00051);
    }

    private static double width(Embedded embedded) {
        var width = Pattern.compile("width:([0-9.]+)pt").matcher(embedded.xml());
        assertTrue(width.find());
        return Double.parseDouble(width.group(1));
    }

    private record Embedded(byte[] nativeData, byte[] preview, String xml) { }
}
