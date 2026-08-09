package com.lz.paperword.core.docx;

import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

class MathTypeAlignmentRegressionTest {

    private static final Path REFERENCE_DOCX = Path.of("rebuild-assets/external/fraction-split-reference.docx");

    private final DocxBuilder builder = new DocxBuilder();

    @Test
    void shouldUseExpectedPreviewForFractionFormula() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithFormula("\\frac{1}{2}"));
        assertEquals("wmf", generated.previewExtension);
        assertWithin(generated.positionHalfPt, -24, 4.0d, "fraction baseline position");
        assertTrue(generated.styleHeightPt >= 12.0d, "fraction preview height should stay readable");
        assertTrue(generated.styleHeightPt <= 32.0d, "fraction preview height should not be oversized");
    }

    @Test
    void shouldNotUseScaledBitmapSizeAsWordDisplaySize() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithFormula("a^2+b^2=25"));

        assertEquals("wmf", generated.previewExtension);
        assertTrue(generated.styleHeightPt <= 18.0d,
            "inline formula display height should use logical points, not high-DPI bitmap pixels");
        assertTrue(generated.styleWidthPt <= 95.0d,
            "inline formula display width should use logical points, not high-DPI bitmap pixels");
    }

    @Test
    void shouldMatchScaledWmfPhysicalBoundsToWordDisplayBox() throws IOException {
        byte[] docx = buildDirectScaledOleDocx("a^2+b^2=c^2", 1.25d);
        ObjectMetrics metrics = extractFirstObjectMetrics(docx);
        Map<String, byte[]> entries = unzipBinaryEntries(docx);
        byte[] wmf = entries.get("word/media/image_eq1.wmf");

        assertNotNull(wmf, "scaled OLE preview must remain a POLYPOLYGON WMF media part");
        assertTrue(wmf.length >= 46, "WMF must include placeable and metafile headers");
        ByteBuffer header = ByteBuffer.wrap(wmf).order(ByteOrder.LITTLE_ENDIAN);
        int unitsPerInch = Short.toUnsignedInt(header.getShort(14));
        double frameWidthPt = (header.getShort(10) - header.getShort(6)) * 72d / unitsPerInch;
        double frameHeightPt = (header.getShort(12) - header.getShort(8)) * 72d / unitsPerInch;
        assertWithin(frameWidthPt, metrics.styleWidthPt, 0.08d,
            "scaled WMF physical width must equal the Word shape width");
        assertWithin(frameHeightPt, metrics.styleHeightPt, 0.08d,
            "scaled WMF physical height must equal the Word shape height");
        assertWithin(metrics.dxaOrig / 20d, metrics.styleWidthPt, 0.08d,
            "dxaOrig must equal the Word shape width");
        assertWithin((metrics.dyaOrig + 1) / 20d, metrics.styleHeightPt, 0.08d,
            "dyaOrig must equal the Word shape height");
    }

    @Test
    void shouldHonorExplicitShapeMetricsWhenPresent() throws IOException {
        byte[] docx = buildDocxWithContent("<p>$\\pwmetrics{12,8,20,14}a^2+b^2=25$</p>");
        ObjectMetrics generated = extractFirstObjectMetrics(docx);

        assertEquals("wmf", generated.previewExtension);
        assertWithin(generated.styleWidthPt, 20.0d, 0.6d, "shape width should come from explicit shape metrics");
        assertWithin(generated.styleHeightPt, 14.0d, 0.6d, "shape height should come from explicit shape metrics");
        assertTrue(generated.dxaOrig >= 395 && generated.dxaOrig <= 405, "dxaOrig should track explicit shape width");
        assertTrue(generated.dyaOrig >= 275 && generated.dyaOrig <= 285, "dyaOrig should track explicit shape height");
    }

    @Test
    void shouldStayCloseToReferenceFractionMetrics() throws IOException {
        assumeTrue(Files.exists(REFERENCE_DOCX), "reference docx should exist");

        ObjectMetrics reference = extractFirstObjectMetrics(Files.readAllBytes(REFERENCE_DOCX));
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithFormula("\\frac{1}{2}"));

        assertTrue(reference.previewExtension.equals("wmf") || reference.previewExtension.equals("emf"),
            "reference preview should be vector");
        assertEquals("wmf", generated.previewExtension);
        // 预览容器变化不应改变公式基线与版面尺寸。
        assertWithin(generated.positionHalfPt, reference.positionHalfPt, 8.0d, "baseline position");
        assertTrue(generated.styleHeightPt >= 12.0d, "style height should stay readable after preview shrink");
        assertTrue(generated.dyaOrig > 0, "dyaOrig should be positive");
        assertTrue(generated.styleWidthPt > 0, "style width should be positive");
    }

    @Test
    void shouldEmbedConcentrationCrossAsOleObject() throws IOException {
        byte[] docx = buildDocxWithContent(
            "浓度十字交叉：<br/>$$\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {20\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {10\\%} & {} & {} & {} & {20\\%}\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        assertTrue(documentXml.contains("<o:OLEObject"), "concentration cross should stay on OLE path");
        assertFalse(documentXml.contains("<w:drawing"), "concentration cross should not fall back to picture drawing");
        assertFalse(documentXml.contains("$$"), "display math delimiters should not leak into Word text");
        ObjectMetrics metrics = extractFirstObjectMetrics(docx);
        assertTrue(metrics.styleWidthPt <= 140.5d, "cross preview width should be capped for Word layout");
    }

    @Test
    void shouldKeepInlineMathPrefixWithoutLeakingDisplayDelimiters() throws IOException {
        byte[] docx = buildDocxWithContent("甲：$10\\times2=20(kg)$<br/>乙：$10\\times3=30(kg)$");
        String documentXml = unzipTextEntries(docx).get("word/document.xml");

        assertNotNull(documentXml);
        assertTrue(documentXml.contains("甲："), "inline formula prefix should stay as text");
        assertTrue(documentXml.contains("乙："), "second inline formula prefix should stay as text");
        assertFalse(documentXml.contains("$$"), "inline formulas should not introduce stray display delimiters");
    }

    @Test
    void shouldIgnoreRemovedPreviewSwitchAndKeepStrictVectorWmf() throws IOException {
        byte[] docx = buildDocxWithFormula("\\sqrt{a^{2}+b^{2}}");
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");

        assertNotNull(documentXml);
        assertNotNull(relsXml);
        assertTrue(documentXml.contains("<o:OLEObject"), "vector preview must keep editable MathType OLE");
        assertEquals("media/image_eq1.wmf",
            extractRelationshipTarget(relsXml, extractFirstImageRelId(documentXml)));
        ObjectMetrics metrics = extractFirstObjectMetrics(docx);
        assertEquals("wmf", metrics.previewExtension);
    }

    @Test
    void shouldPackageMathTypeOleWithNativeMtefStream() throws IOException {
        byte[] docx = buildDocxWithFormula("a^2+b^2=25");
        Map<String, byte[]> entries = unzipBinaryEntries(docx);
        String documentXml = new String(entries.get("word/document.xml"), StandardCharsets.UTF_8);
        String relsXml = new String(entries.get("word/_rels/document.xml.rels"), StandardCharsets.UTF_8);

        assertTrue(documentXml.contains("ProgID=\"Equation.DSMT4\""), "Word object should advertise MathType ProgID");

        String oleRelId = extractFirstOleRelId(documentXml);
        String oleTarget = extractRelationshipTarget(relsXml, oleRelId);
        assertEquals("embeddings/oleObject1.bin", oleTarget);

        byte[] oleBytes = entries.get("word/" + oleTarget);
        assertNotNull(oleBytes, "OLE binary part should exist");

        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(oleBytes))) {
            assertTrue(fs.getRoot().hasEntry("\u0001CompObj"), "OLE should contain CompObj stream");
            assertTrue(fs.getRoot().hasEntry("Equation Native"), "OLE should contain MathType Equation Native stream");

            byte[] equationNative = readOleStream(fs, "Equation Native");
            int headerSize = Short.toUnsignedInt(ByteBuffer.wrap(equationNative, 0, 2)
                .order(ByteOrder.LITTLE_ENDIAN).getShort());
            int mtefLength = ByteBuffer.wrap(equationNative, 8, 4)
                .order(ByteOrder.LITTLE_ENDIAN).getInt();

            assertTrue(headerSize >= 28, "Equation Native should have an OLE header");
            assertEquals(equationNative.length - headerSize, mtefLength, "MTEF length should match header metadata");
            assertEquals(5, equationNative[headerSize] & 0xFF, "MTEF should be v5");
            assertEquals(1, equationNative[headerSize + 1] & 0xFF, "MTEF platform should be Windows");
            assertEquals(0, equationNative[headerSize + 2] & 0xFF, "MTEF product should be MathType");
            String appKey = new String(equationNative, headerSize + 5, 5, StandardCharsets.US_ASCII);
            assertTrue(appKey.startsWith("DSMT"), "MTEF application key should identify Design Science MathType");
            byte[] mtef = java.util.Arrays.copyOfRange(equationNative, headerSize, equationNative.length);
            assertFalse(containsBytes(mtef, new byte[] {0x09, 0x65}),
                "OLE/MTEF body should not write an explicit point-size SIZE record");
        }
    }

    @Test
    void shouldWriteHashedFormulaTraceWithoutLeakingLatex() throws IOException {
        byte[] docx = buildDocxWithFormula("\\frac{1}{2}");
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        Matcher titleMatcher = Pattern.compile("<v:imagedata\\b[^>]*o:title=\"([^\"]*)\"").matcher(documentXml);
        assertTrue(titleMatcher.find(), "preview image should expose a diagnostic formula trace id");
        String title = titleMatcher.group(1);
        assertTrue(title.matches("pwf:1-[0-9a-f]{16}"), "trace id should include formula ordinal plus a short hash");
        assertFalse(title.contains("\\frac"), "trace id must not leak raw LaTeX");
    }

    @Test
    void shouldUseStableAsciiWhitespaceTraceHash() throws IOException {
        byte[] docx = buildDocxWithFormula("\\pwmetrics{1,2,3,4}\t\\pwstyle{trace}\n a\u00a0 +\t b ");
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        Matcher titleMatcher = Pattern.compile("<v:imagedata\\b[^>]*o:title=\"([^\"]*)\"").matcher(documentXml);
        assertTrue(titleMatcher.find(), "preview image should expose a diagnostic formula trace id");
        assertEquals("pwf:1-cb23f6635a581786", titleMatcher.group(1),
            "trace hash must match the Python request-to-DOCX whitespace normalization contract");
    }

    @Test
    void shouldUseSameTraceHashWhenEmbedderReceivesNbspDirectly()
        throws ReflectiveOperationException {
        Method method = MathTypeEmbedder.class.getDeclaredMethod("formulaTraceId", int.class, String.class);
        method.setAccessible(true);

        assertEquals("pwf:7-cb23f6635a581786", invokeFormulaTraceId(method, 7,
            "\\pwmetrics{1,2,3,4}\t\\pwstyle{trace}\n a\u00a0 +\t b "));
    }

    @Test
    void shouldDisambiguateDuplicateFormulaTraceIdsByDocumentOrdinal() throws IOException {
        List<String> titles = extractPreviewTitles(buildDocxWithContent("甲：$ABCD$<br/>乙：$ABCD$"));

        assertEquals(2, titles.size(), "both duplicate formulas should have preview titles");
        assertTrue(titles.get(0).matches("pwf:1-[0-9a-f]{16}"));
        assertTrue(titles.get(1).matches("pwf:2-[0-9a-f]{16}"));
        assertEquals(titles.get(0).substring(titles.get(0).indexOf('-') + 1),
            titles.get(1).substring(titles.get(1).indexOf('-') + 1),
            "duplicate formulas should share the content hash part");
        assertNotEquals(titles.get(0), titles.get(1),
            "duplicate formulas must remain unique through the document ordinal");
    }

    @Test
    void shouldResetTraceOrdinalsForEachBuildOnReusedBuilder() throws IOException {
        List<String> firstTitles = extractPreviewTitles(buildDocxWithContent("甲：$ABCD$"));
        List<String> secondTitles = extractPreviewTitles(buildDocxWithContent("乙：$ABCD$"));

        assertEquals(1, firstTitles.size());
        assertEquals(1, secondTitles.size());
        assertTrue(firstTitles.get(0).startsWith("pwf:1-"));
        assertEquals(firstTitles.get(0), secondTitles.get(0),
            "a reused DocxBuilder should reset formula trace ordinals for each document");
    }

    @Test
    void shouldKeepMultipleVerticalTemplatesAsOleObjects() throws IOException {
        byte[] docx = buildDocxWithContent(
            "整数加法：<br/>$$\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}$$"
                + "<br/>小数加法：<br/>$$\\begin{array}{rcr}12 & . & 50 \\\\ +3 & . & 75 \\\\ \\hline 16 & . & 25\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        assertNotNull(documentXml);

        Matcher matcher = Pattern.compile("<o:OLEObject\\b").matcher(documentXml);
        int count = 0;
        while (matcher.find()) {
            count++;
        }
        assertTrue(count >= 2, "multiple vertical templates should remain OLE formulas");
    }

    @Test
    void shouldClampWideArrayPreviewWidth() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithContent(
            "宽矩阵：<br/>$$\\begin{array}{rrrrrrrr}{} & {} & {3} & {4} & {0} & {} & {} & {} \\\\ {\\times} & {} & {5} & {3} & {0} & {0} & {} & {} \\\\ \\hline {} & {} & {} & {} & {} & {} & {} & {} \\\\ {+} & {} & {1} & {0} & {2} & {0} & {} & {} \\\\ {+} & {1} & {7} & {0} & {0} & {} & {} & {} \\\\ \\hline {} & {1} & {8} & {0} & {2} & {0} & {0} & {}\\end{array}$$"
        ));
        assertTrue(generated.styleWidthPt <= 300.5d, "wide arithmetic array preview width should stay capped");
    }

    @Test
    void shouldAllowDerivationArrayMoreWidthThanArithmeticLayout() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithContent(
            "推导：<br/>$$\\begin{array}{rl}&=\\frac{1}{2}\\times \\left(\\frac{1}{2\\times 3}-\\frac{1}{3\\times 4}+\\frac{1}{3\\times 4}-\\frac{1}{4\\times 5}+\\frac{1}{4\\times 5}-\\frac{1}{5\\times 6}+\\cdots +\\frac{1}{9\\times 10}-\\frac{1}{10\\times 11}\\right)+2\\times \\left(\\frac{1}{3}-\\frac{1}{4}+\\frac{1}{4}-\\frac{1}{5}+\\cdots +\\frac{1}{10}-\\frac{1}{11}\\right)\\\\&=\\frac{1}{12}-\\frac{1}{220}+\\frac{2}{3}-\\frac{2}{11}\\end{array}$$"
        ));
        assertTrue(generated.styleWidthPt > 300.5d, "derivation array should not use the arithmetic width cap");
        assertTrue(generated.styleWidthPt <= 420.5d, "derivation array preview width should still fit the page");
    }

    @Test
    void shouldClampLongInlineFormulaToPageWidth() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithContent(
            "长题干：$\\frac{1}{1\\times3}+\\frac{2}{3\\times5}+\\frac{2^2}{5\\times7}+\\cdots+\\frac{2^8}{17\\times19}-\\left(\\frac{2^3}{1\\times3\\times5}+\\frac{2^4}{3\\times5\\times7}+\\cdots+\\frac{2^{11}}{17\\times19\\times21}\\right)$"
        ));
        assertTrue(generated.styleWidthPt <= 500.5d, "long inline formula preview width should fit the page");
    }

    private String invokeFormulaTraceId(Method method, int formulaIndex, String latex)
        throws InvocationTargetException, IllegalAccessException {
        return (String) method.invoke(new MathTypeEmbedder(), formulaIndex, latex);
    }

    private byte[] buildDocxWithFormula(String latex) throws IOException {
        return buildDocxWithContent("<p>计算 $" + latex + "$</p>");
    }

    private byte[] buildDirectScaledOleDocx(String latex, double displayScale) throws IOException {
        LaTeXParser.DetailedParseResult parsed = new LaTeXParser().parseDetailed(latex);
        assertTrue(parsed.isSupported(), () -> parsed.diagnostics().toString());
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            var paragraph = document.createParagraph();
            var run = paragraph.createRun();
            new MathTypeEmbedder().embedEquation(paragraph, run, parsed.ast(), latex,
                displayScale, 420d);
            document.write(output);
            return output.toByteArray();
        }
    }

    private byte[] buildDocxWithContent(String content) throws IOException {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("MathType alignment regression");
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、分数题");

        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(6);
        question.setScore(5);
        question.setContent(content);

        section.setQuestions(List.of(question));
        request.setSections(List.of(section));
        return builder.build(request);
    }

    private ObjectMetrics extractFirstObjectMetrics(byte[] docxBytes) throws IOException {
        Map<String, String> entries = unzipTextEntries(docxBytes);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");

        assertNotNull(documentXml, "document.xml should exist");
        assertNotNull(relsXml, "document rels should exist");

        Matcher objectMatcher = Pattern.compile(
            "<w:rPr>.*?<w:position w:val=\"(-?\\d+)\".*?</w:rPr>.*?" +
                "<w:object w:dxaOrig=\"(\\d+)\" w:dyaOrig=\"(\\d+)\".*?" +
                "<v:shape[^>]*style=\"([^\"]+)\"[^>]*>.*?" +
                "<v:imagedata r:id=\"([^\"]+)\"",
            Pattern.DOTALL)
            .matcher(documentXml);

        assertTrue(objectMatcher.find(), "should contain at least one MathType object");

        int positionHalfPt = Integer.parseInt(objectMatcher.group(1));
        int dxaOrig = Integer.parseInt(objectMatcher.group(2));
        int dyaOrig = Integer.parseInt(objectMatcher.group(3));
        String style = objectMatcher.group(4);
        String imageRelId = objectMatcher.group(5);

        double styleWidthPt = parseStyleMetric(style, "width");
        double styleHeightPt = parseStyleMetric(style, "height");
        String previewTarget = extractRelationshipTarget(relsXml, imageRelId);
        assertNotNull(previewTarget, "image relationship should exist");

        String previewExtension = previewTarget.substring(previewTarget.lastIndexOf('.') + 1).toLowerCase();
        return new ObjectMetrics(positionHalfPt, dxaOrig, dyaOrig, styleWidthPt, styleHeightPt, previewExtension);
    }

    private String extractFirstImageRelId(String documentXml) {
        Matcher objectMatcher = Pattern.compile("<v:imagedata r:id=\"([^\"]+)\"").matcher(documentXml);
        assertTrue(objectMatcher.find(), "should contain preview image relationship");
        return objectMatcher.group(1);
    }

    private String extractFirstOleRelId(String documentXml) {
        Matcher objectMatcher = Pattern.compile("<o:OLEObject\\b[^>]*r:id=\"([^\"]+)\"").matcher(documentXml);
        assertTrue(objectMatcher.find(), "should contain OLE object relationship");
        return objectMatcher.group(1);
    }

    private List<String> extractPreviewTitles(byte[] docxBytes) throws IOException {
        String documentXml = unzipTextEntries(docxBytes).get("word/document.xml");
        assertNotNull(documentXml);
        Matcher matcher = Pattern.compile("<v:imagedata\\b[^>]*o:title=\"([^\"]*)\"").matcher(documentXml);
        java.util.ArrayList<String> titles = new java.util.ArrayList<>();
        while (matcher.find()) {
            titles.add(matcher.group(1));
        }
        return titles;
    }

    private Map<String, String> unzipTextEntries(byte[] docxBytes) throws IOException {
        Map<String, String> entries = new HashMap<>();
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(docxBytes), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.isDirectory()) {
                    continue;
                }
                entries.put(entry.getName(), new String(zis.readAllBytes(), StandardCharsets.UTF_8));
            }
        }
        return entries;
    }

    private Map<String, byte[]> unzipBinaryEntries(byte[] docxBytes) throws IOException {
        Map<String, byte[]> entries = new HashMap<>();
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(docxBytes), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.isDirectory()) {
                    continue;
                }
                entries.put(entry.getName(), zis.readAllBytes());
            }
        }
        return entries;
    }

    private byte[] readOleStream(POIFSFileSystem fs, String name) throws IOException {
        Entry entry = fs.getRoot().getEntry(name);
        assertTrue(entry instanceof DocumentEntry, name + " should be a document stream");
        try (DocumentInputStream in = new DocumentInputStream((DocumentEntry) entry)) {
            return in.readAllBytes();
        }
    }

    private boolean containsBytes(byte[] bytes, byte[] needle) {
        outer:
        for (int i = 0; i <= bytes.length - needle.length; i++) {
            for (int j = 0; j < needle.length; j++) {
                if (bytes[i + j] != needle[j]) {
                    continue outer;
                }
            }
            return true;
        }
        return false;
    }

    private String extractRelationshipTarget(String relsXml, String relId) {
        Matcher matcher = Pattern.compile(
            "<Relationship[^>]*Id=\"" + Pattern.quote(relId) + "\"[^>]*Target=\"([^\"]+)\"")
            .matcher(relsXml);
        if (!matcher.find()) {
            return null;
        }
        return matcher.group(1);
    }

    private double parseStyleMetric(String style, String property) {
        Matcher matcher = Pattern.compile(property + ":([0-9.]+)pt").matcher(style);
        assertTrue(matcher.find(), "style should contain " + property + " pt");
        return Double.parseDouble(matcher.group(1));
    }

    private void assertWithin(double actual, double expected, double tolerance, String label) {
        assertTrue(Math.abs(actual - expected) <= tolerance,
            label + " out of range, expected around " + expected + " but got " + actual);
    }

    private record ObjectMetrics(
        int positionHalfPt,
        int dxaOrig,
        int dyaOrig,
        double styleWidthPt,
        double styleHeightPt,
        String previewExtension
    ) {
    }
}
