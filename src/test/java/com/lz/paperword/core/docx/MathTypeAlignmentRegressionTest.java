package com.lz.paperword.core.docx;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayInputStream;
import java.io.IOException;
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

class MathTypeAlignmentRegressionTest {

    private static final Path REFERENCE_DOCX = Path.of("d:/pdf2word/data/分数.docx");

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
        assertTrue(Files.exists(REFERENCE_DOCX), "reference docx should exist");

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
    void shouldEmbedArrayFormulaAsOleObjectInsteadOfPictureFallback() throws IOException {
        byte[] docx = buildDocxWithContent(
            "计算下列竖式：<br/>$$\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");

        assertNotNull(documentXml);
        assertNotNull(relsXml);
        assertTrue(documentXml.contains("<o:OLEObject"), "complex array should still be embedded as OLE");
        assertFalse(documentXml.contains("<w:drawing"), "complex array should not fall back to picture drawing");
        assertNotNull(extractRelationshipTarget(relsXml, extractFirstImageRelId(documentXml)));
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
