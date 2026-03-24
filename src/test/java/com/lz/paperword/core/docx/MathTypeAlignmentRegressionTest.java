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
import static org.junit.jupiter.api.Assumptions.assumeTrue;

class MathTypeAlignmentRegressionTest {

    private static final Path REFERENCE_DOCX = Path.of("d:/pdf2word/data/分数.docx");

    private final DocxBuilder builder = new DocxBuilder();

    @Test
    void shouldUseExpectedPreviewForFractionFormula() throws IOException {
        ObjectMetrics generated = extractFirstObjectMetrics(buildDocxWithFormula("\\frac{1}{2}"));
        assertEquals("png", generated.previewExtension);
        assertEquals(-16, generated.positionHalfPt);
        assertTrue(generated.styleHeightPt >= 12.0d, "fraction preview height should remain positive after shrink");
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
    void shouldEmbedConcentrationCrossAsOleObject() throws IOException {
        byte[] docx = buildDocxWithContent(
            "浓度十字交叉：<br/>$$\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {20\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {10\\%} & {} & {} & {} & {20\\%}\\end{array}$$"
        );
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");

        assertNotNull(documentXml);
        assertTrue(documentXml.contains("<o:OLEObject"), "concentration cross should stay on OLE path");
        assertFalse(documentXml.contains("<w:drawing"), "concentration cross should not fall back to picture drawing");
    }

    @Test
    void shouldKeepCurrentLongDivisionAsOleObject() throws IOException {
        byte[] docx = buildDocxWithContent("长除法：<br/>$\\longdiv[246]{5}{1234}$");
        Map<String, String> entries = unzipTextEntries(docx);
        String documentXml = entries.get("word/document.xml");
        String relsXml = entries.get("word/_rels/document.xml.rels");

        assertNotNull(documentXml);
        assertNotNull(relsXml);
        assertTrue(documentXml.contains("<o:OLEObject"), "current long division should stay on OLE path");
        assertFalse(documentXml.contains("[\\longdiv[246]{5}{1234}]"), "current long division should not degrade to raw text");
        assertEquals("media/image_eq1.png", extractRelationshipTarget(relsXml, extractFirstImageRelId(documentXml)));
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
