package com.lz.paperword.core.docx;

import com.lz.paperword.core.mtef.MtefRecord;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
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

    private static final Path REFERENCE_DOCX = Path.of("d:/pdf2word/data/分数.docx");
    private static final Path DECIMAL_ALIGNMENT_REFERENCE_DOCX = Path.of("J:/lingzhi/data/小数点对齐.docx");
    private static final Path DIVISION_ALIGNMENT_REFERENCE_DOCX = Path.of("J:/lingzhi/data/除法参考.docx");

    private final DocxBuilder builder = new DocxBuilder();

    @Test
    void shouldUseExpectedLegacyPreviewForFractionFormula() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithFormula("\\frac{1}{2}"));
        assertEquals("png", generated.previewExtension);
        assertEquals(-24, generated.positionHalfPt);
        assertTrue(generated.styleHeightPt >= 20.0d, "fraction preview height should remain tall enough");
    }

    @Test
    void shouldStayCloseToReferenceFractionMetrics() throws IOException {
        assumeTrue(Files.exists(REFERENCE_DOCX), "reference docx should exist");

        ObjectMetrics reference = extractFirstObjectMetrics(Files.readAllBytes(REFERENCE_DOCX));
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithFormula("\\frac{1}{2}"));

        assertTrue(reference.previewExtension.equals("wmf") || reference.previewExtension.equals("emf"),
            "reference preview should be vector");
        assertEquals("png", generated.previewExtension);
        assertEquals(reference.positionHalfPt, generated.positionHalfPt);
        assertWithin(generated.styleHeightPt, reference.styleHeightPt, 12.0d, "style height");
        assertTrue(generated.dyaOrig > 0, "dyaOrig should be positive");
        assertTrue(generated.styleWidthPt > 0, "style width should be positive");
    }

    @Test
    void shouldRenderLegacyDivisionArrayAsPictureInsteadOfOleObject() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算下列竖式：<br/>$$\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");

        assertNotNull(documentXml);
        assertNotNull(relsXml);
        assertFalse(documentXml.contains("<o:OLEObject"), "legacy division array should no longer be embedded as OLE");
        assertTrue(documentXml.contains("<w:drawing"), "legacy division array should be inserted as drawing picture");
        assertTrue(relsXml.contains("image"), "drawing picture should keep image relationship");
    }

    @Test
    void shouldSplitDisplayFormulaIntoSeparateParagraphs() throws IOException {
        byte[] docx = buildDocxWithContent(
            "先按整数乘法计算：<br/>$$\\begin{array}{rrrrr} & 1 & 2 & 3 & \\\\ \\times & & 4 & 5 & \\\\ \\hline & & 6 & 1 & 5 \\\\ + & 4 & 9 & 2 & \\\\ \\hline & 5 & 5 & 3 & 5\\end{array}$$<br/>再确定小数点位置，得到 55.35。"
        );

        try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(docx))) {
            assertTrue(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("先按整数乘法计算")));
            assertTrue(document.getParagraphs().stream().anyMatch(p -> p.getText().contains("再确定小数点位置")));
            assertTrue(document.getParagraphs().stream().anyMatch(p -> p.getCTP().xmlText().contains("<o:OLEObject")),
                "display formula should occupy a dedicated paragraph");
        }
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
    void shouldKeepSubtractionVerticalTemplatesAsOleObjects() throws IOException {
        byte[] docx = buildDocxWithContent(
            "整数减法：<br/>$$\\begin{array}{rrrr} & 8 & 6 & 4 \\\\ - & 2 & 7 & 9 \\\\ \\hline & 5 & 8 & 5\\end{array}$$"
                + "<br/>小数减法：<br/>$$\\begin{array}{rcr}12 & . & 50 \\\\ -3 & . & 75 \\\\ \\hline 8 & . & 75\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        assertNotNull(documentXml);

        Matcher matcher = Pattern.compile("<o:OLEObject\\b").matcher(documentXml);
        int count = 0;
        while (matcher.find()) {
            count++;
        }
        assertTrue(count >= 2, "subtraction vertical templates should remain OLE formulas");
    }

    @Test
    void shouldKeepDecimalAlignmentAsMatrixOleObject() throws IOException {
        byte[] generatedDocx = buildDocxWithContent(
            "小数加法：<br/>$$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {+} & {3} & {.} & {75} \\\\ \\hline {} & {16} & {.} & {25}\\end{array}$$"
        );

        ObjectMetrics generated = extractFirstObjectMetrics(generatedDocx);
        byte[] generatedMtef = extractFirstEquationNative(generatedDocx);

        assertEquals(-42, generated.positionHalfPt, "decimal alignment baseline should follow reference");
        assertTrue(generated.styleWidthPt > 0, "decimal alignment style width should be positive");
        assertTrue(generated.styleHeightPt > 0, "decimal alignment style height should be positive");
        assertTrue(containsRecord(generatedMtef, MtefRecord.PILE), "generated decimal doc should contain PILE");
    }

    @Test
    void shouldRenderExplicitLongDivisionAsPicture() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算：<br/>$$\\longdiv[65]{13}{845}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        assertNotNull(documentXml);
        assertFalse(documentXml.contains("<o:OLEObject"), "explicit long division should be rendered as picture");
        assertTrue(documentXml.contains("<w:drawing"), "explicit long division should produce drawing picture");
    }

    @Test
    void shouldRenderStructuredThreeStepLongDivisionAsPicture() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算：<br/>$$\\longdiv[246]{5}{1234}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        assertFalse(documentXml.contains("<o:OLEObject"), "structured three-step long division should be picture only");
        assertTrue(documentXml.contains("<w:drawing"), "structured three-step long division should produce drawing");
    }

    @Test
    void shouldRenderExplicitLongDivisionWithoutQuotientAsPicture() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算：<br/>$$\\longdiv{13}{845}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        assertNotNull(documentXml);
        assertFalse(documentXml.contains("<o:OLEObject"), "long division without quotient should be rendered as picture");
        assertTrue(documentXml.contains("<w:drawing"), "long division without quotient should produce drawing");
    }

    @Test
    void shouldRenderCompositeLongDivisionWithStepsAsPicture() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算：<br/>$$\\begin{array}{l}{\\longdiv[65]{13}{845}} \\\\ {\\begin{array}{rrr}{7} & {8} & {} \\\\ \\hline {} & {6} & {5} \\\\ {} & {6} & {5} \\\\ \\hline {} & {} & {0}\\end{array}}\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        assertNotNull(documentXml);
        assertFalse(documentXml.contains("<o:OLEObject"), "composite long division should be picture only");
        assertTrue(documentXml.contains("<w:drawing"), "composite long division should produce drawing");
    }

    @Test
    void shouldRenderLongdivisionEnvironmentAsPicture() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算：<br/>$$\\begin{longdivision}{r|l}13 & 845 \\\\ \\hline & 65\\end{longdivision}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        assertFalse(documentXml.contains("<o:OLEObject"), "longdivision environment should be rendered as picture");
        assertTrue(documentXml.contains("<w:drawing"), "longdivision environment should produce drawing");
    }

    @Test
    void shouldKeepNonLongDivisionFormulasAsOleObjectsWhenLongDivisionExists() throws IOException {
        byte[] docx = buildDocxWithContent(
            "分数：<br/>$$\\frac{1}{2}$$<br/>长除法：<br/>$$\\longdiv[246]{5}{1234}$$<br/>小数：<br/>$$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {+} & {3} & {.} & {75} \\\\ \\hline {} & {16} & {.} & {25}\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        assertTrue(countMatches(documentXml, "<o:OLEObject\\b") >= 2, "fraction and decimal formula should remain OLE");
        assertTrue(countMatches(documentXml, "<w:drawing\\b") >= 1, "long division should be inserted as drawing");
    }

    @Test
    void shouldKeepLongDivisionAndOleCountsSeparated() throws IOException {
        byte[] docx = buildDocxWithContent(
            "加法：<br/>$$\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}$$"
                + "<br/>旧除法：<br/>$$\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}$$"
                + "<br/>显式长除法：<br/>$$\\longdiv[65]{13}{845}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        assertEquals(1, countMatches(documentXml, "<o:OLEObject\\b"), "only arithmetic formula should remain OLE");
        assertEquals(2, countMatches(documentXml, "<w:drawing\\b"), "both division formulas should be pictures");
    }

    @Test
    void shouldKeepAllCommonVerticalLayoutsAsOleObjects() throws IOException {
        byte[] docx = buildDocxWithContent(
            "加法：<br/>$$\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}$$"
                + "<br/>减法：<br/>$$\\begin{array}{rrrr} & 8 & 6 & 4 \\\\ - & 2 & 7 & 9 \\\\ \\hline & 5 & 8 & 5\\end{array}$$"
                + "<br/>乘法：<br/>$$\\begin{array}{rrrrr} & 1 & 2 & 3 & \\\\ \\times & & 4 & 5 & \\\\ \\hline & & 6 & 1 & 5 \\\\ + & 4 & 9 & 2 & \\\\ \\hline & 5 & 5 & 3 & 5\\end{array}$$"
                + "<br/>小数加法：<br/>$$\\begin{array}{rcr}12 & . & 50 \\\\ +3 & . & 75 \\\\ \\hline 16 & . & 25\\end{array}$$"
                + "<br/>小数减法：<br/>$$\\begin{array}{rcr}12 & . & 50 \\\\ -3 & . & 75 \\\\ \\hline 8 & . & 75\\end{array}$$"
                + "<br/>旧除法：<br/>$$\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}$$"
                + "<br/>显式长除法：<br/>$$\\longdiv[65]{13}{845}$$"
                + "<br/>显式长除法无商：<br/>$$\\longdiv{13}{845}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        assertNotNull(documentXml);
        assertTrue(countMatches(documentXml, "<o:OLEObject\\b") >= 5,
            "non-long-division layouts should remain OLE formulas");
        assertTrue(countMatches(documentXml, "<w:drawing\\b") >= 3,
            "all long-division variants should be rendered as pictures");
    }

    private byte[] buildDocxWithFormula(String latex) throws IOException {
        return buildDocxWithContent("<p>计算 $" + latex + "$</p>");
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
        Matcher objectMatcher = Pattern.compile("<o:OLEObject[^>]*r:id=\"([^\"]+)\"").matcher(documentXml);
        assertTrue(objectMatcher.find(), "should contain ole relationship");
        return objectMatcher.group(1);
    }

    private byte[] extractFirstEquationNative(byte[] docxBytes) throws IOException {
        return extractEquationNativeByIndex(docxBytes, 0);
    }

    private byte[] extractEquationNativeByIndex(byte[] docxBytes, int objectIndex) throws IOException {
        Map<String, String> entries = unzipTextEntries(docxBytes);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");

        assertNotNull(documentXml, "document.xml should exist");
        assertNotNull(relsXml, "document rels should exist");

        List<String> oleRelIds = extractOleRelIds(documentXml);
        assertTrue(objectIndex >= 0 && objectIndex < oleRelIds.size(), "requested ole object should exist");
        String oleTarget = extractRelationshipTarget(relsXml, oleRelIds.get(objectIndex));
        assertNotNull(oleTarget, "ole relationship should exist");

        byte[] oleBytes = readZipEntry(docxBytes, "word/" + oleTarget);
        assertNotNull(oleBytes, "ole embedding should exist");

        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(oleBytes))) {
            assertTrue(fs.getRoot().hasEntry("Equation Native"), "Equation Native stream should exist");
            if (!(fs.getRoot().getEntry("Equation Native") instanceof DocumentEntry)) {
                fail("Equation Native should be a document entry");
            }
            DocumentEntry entry = (DocumentEntry) fs.getRoot().getEntry("Equation Native");
            try (DocumentInputStream dis = new DocumentInputStream(entry)) {
                byte[] equationNative = dis.readAllBytes();
                int hdrSize = Short.toUnsignedInt((short) ((equationNative[0] & 0xFF) | ((equationNative[1] & 0xFF) << 8)));
                assertTrue(hdrSize > 0 && equationNative.length > hdrSize, "Equation Native header should be valid");
                byte[] mtef = new byte[equationNative.length - hdrSize];
                System.arraycopy(equationNative, hdrSize, mtef, 0, mtef.length);
                return mtef;
            }
        }
    }

    private List<String> extractOleRelIds(String documentXml) {
        List<String> relIds = new ArrayList<>();
        Matcher matcher = Pattern.compile("<o:OLEObject[^>]*r:id=\"([^\"]+)\"").matcher(documentXml);
        while (matcher.find()) {
            relIds.add(matcher.group(1));
        }
        assertFalse(relIds.isEmpty(), "should contain ole relationship");
        return relIds;
    }

    private LongDivisionOleInfo extractLongDivisionOleInfo(byte[] docxBytes) throws IOException {
        Map<String, String> entries = unzipTextEntries(docxBytes);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");
        assertNotNull(documentXml, "document.xml should exist");
        assertNotNull(relsXml, "document rels should exist");

        Matcher objectMatcher = Pattern.compile(
            "<w:rPr>.*?<w:position w:val=\"(-?\\d+)\".*?</w:rPr>.*?" +
                "<w:object w:dxaOrig=\"(\\d+)\" w:dyaOrig=\"(\\d+)\".*?" +
                "<v:shape[^>]*style=\"([^\"]+)\"[^>]*>.*?" +
                "<v:imagedata r:id=\"([^\"]+)\".*?" +
                "<o:OLEObject[^>]*r:id=\"([^\"]+)\"",
            Pattern.DOTALL)
            .matcher(documentXml);

        int index = 0;
        while (objectMatcher.find()) {
            byte[] mtef = extractEquationNativeByIndex(docxBytes, index);
            if (!containsByte(mtef, (byte) MtefRecord.TM_LDIV)) {
                index++;
                continue;
            }
            int positionHalfPt = Integer.parseInt(objectMatcher.group(1));
            int dxaOrig = Integer.parseInt(objectMatcher.group(2));
            int dyaOrig = Integer.parseInt(objectMatcher.group(3));
            double styleWidthPt = parseStyleMetric(objectMatcher.group(4), "width");
            double styleHeightPt = parseStyleMetric(objectMatcher.group(4), "height");
            return new LongDivisionOleInfo(
                positionHalfPt,
                dxaOrig,
                dyaOrig,
                styleWidthPt,
                styleHeightPt,
                extractDigitStream(mtef),
                extractTmLdivWindowHex(mtef)
            );
        }
        fail("should contain long division ole object");
        return null;
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

    private byte[] readZipEntry(byte[] docxBytes, String entryName) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(docxBytes), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (!entry.isDirectory() && entryName.equals(entry.getName())) {
                    return zis.readAllBytes();
                }
            }
        }
        return null;
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

    private int countMatches(String text, String pattern) {
        Matcher matcher = Pattern.compile(pattern).matcher(text == null ? "" : text);
        int count = 0;
        while (matcher.find()) {
            count++;
        }
        return count;
    }

    private boolean containsRecord(byte[] bytes, int recordType) {
        for (int i = 12; i < bytes.length; i++) {
            if ((bytes[i] & 0xFF) == recordType) {
                return true;
            }
        }
        return false;
    }

    private boolean containsByte(byte[] bytes, byte value) {
        for (byte current : bytes) {
            if (current == value) {
                return true;
            }
        }
        return false;
    }

    private String extractDigitStream(byte[] bytes) {
        StringBuilder digits = new StringBuilder();
        for (int i = 0; i <= bytes.length - 5; i++) {
            if ((bytes[i] & 0xFF) != MtefRecord.CHAR) {
                continue;
            }
            int options = bytes[i + 1] & 0xFF;
            int mtcode = (bytes[i + 3] & 0xFF) | ((bytes[i + 4] & 0xFF) << 8);
            char ch = (char) mtcode;
            if (Character.isDigit(ch)) {
                digits.append(ch);
            }
            i += ((options & MtefRecord.OPT_CHAR_ENC_CHAR_8) != 0) ? 5 : 4;
        }
        return digits.toString();
    }

    private String extractTmLdivWindowHex(byte[] mtef) {
        byte[] needle = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV};
        for (int i = 0; i <= mtef.length - needle.length; i++) {
            boolean match = true;
            for (int j = 0; j < needle.length; j++) {
                if (mtef[i + j] != needle[j]) {
                    match = false;
                    break;
                }
            }
            if (!match) {
                continue;
            }
            int start = Math.max(i - 24, 0);
            int end = Math.min(i + 80, mtef.length);
            StringBuilder builder = new StringBuilder();
            for (int k = start; k < end; k++) {
                builder.append(String.format("%02x", mtef[k] & 0xFF));
            }
            return builder.toString();
        }
        fail("should contain tmLDIV window");
        return "";
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

    private record LongDivisionOleInfo(
        int positionHalfPt,
        int dxaOrig,
        int dyaOrig,
        double styleWidthPt,
        double styleHeightPt,
        String digitStream,
        String tmLdivWindowHex
    ) {
    }
}
