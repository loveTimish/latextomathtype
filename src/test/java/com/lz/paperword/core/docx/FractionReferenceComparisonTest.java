package com.lz.paperword.core.docx;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

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

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class FractionReferenceComparisonTest {

    private static final Path OUTPUT_DIR = Path.of("target/generated-docs/reference-comparison");
    private static final Path OUTPUT_DOCX = OUTPUT_DIR.resolve("fraction-comparison.docx");
    private static final Path REFERENCE_DOCX = Path.of("d:/pdf2word/data/分数.docx");
    private static final String FRACTION_LATEX = "\\left(\\frac{1}{2}+\\frac{2}{3}\\right) \\div \\left(\\frac{5}{6}\\times\\frac{3}{5}\\right)";

    private final DocxBuilder builder = new DocxBuilder();

    @Test
    void shouldGenerateComparisonDocxNearReference() throws IOException {
        byte[] generated = buildFractionDocx();
        Files.createDirectories(OUTPUT_DIR);
        Files.write(OUTPUT_DOCX, generated);

        Assumptions.assumeTrue(Files.exists(REFERENCE_DOCX), "reference docx is only available in the Windows comparison environment");

        Metrics reference = extractMetrics(Files.readAllBytes(REFERENCE_DOCX));
        Metrics current = extractMetrics(generated);

        assertTrue(current.previewExtension.equals("png"));
        assertTrue(current.heightPt <= reference.heightPt * 1.25d,
            "generated formula should not be much taller than reference");
        assertTrue(current.heightPt >= reference.heightPt * 0.65d,
            "generated formula should not be too short");

        System.out.printf(
            "reference: pos=%d width=%.1fpt height=%.1fpt preview=%s%n",
            reference.positionHalfPt, reference.widthPt, reference.heightPt, reference.previewExtension);
        System.out.printf(
            "generated: pos=%d width=%.1fpt height=%.1fpt preview=%s file=%s%n",
            current.positionHalfPt, current.widthPt, current.heightPt, current.previewExtension, OUTPUT_DOCX);
    }

    private byte[] buildFractionDocx() throws IOException {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("分数公式对照");
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、计算题");

        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(6);
        question.setScore(5);
        question.setContent("<p>$" + FRACTION_LATEX + "$</p>");

        section.setQuestions(List.of(question));
        request.setSections(List.of(section));
        return builder.build(request);
    }

    private Metrics extractMetrics(byte[] docxBytes) throws IOException {
        Map<String, String> textEntries = unzipTextEntries(docxBytes);
        String documentXml = textEntries.get("word/document.xml");
        String relsXml = textEntries.get("word/_rels/document.xml.rels");
        assertNotNull(documentXml);
        assertNotNull(relsXml);

        Matcher matcher = Pattern.compile(
            "<w:rPr>.*?<w:position w:val=\"(-?\\d+)\".*?</w:rPr>.*?" +
                "<w:object w:dxaOrig=\"(\\d+)\" w:dyaOrig=\"(\\d+)\".*?" +
                "<v:shape[^>]*style=\"([^\"]+)\"[^>]*>.*?<v:imagedata r:id=\"([^\"]+)\"",
            Pattern.DOTALL)
            .matcher(documentXml);
        assertTrue(matcher.find(), "expected MathType object");

        String style = matcher.group(4);
        String relId = matcher.group(5);
        String target = extractRelationshipTarget(relsXml, relId);
        assertNotNull(target);

        return new Metrics(
            Integer.parseInt(matcher.group(1)),
            Integer.parseInt(matcher.group(2)),
            Integer.parseInt(matcher.group(3)),
            parseStyleMetric(style, "width"),
            parseStyleMetric(style, "height"),
            target.substring(target.lastIndexOf('.') + 1).toLowerCase()
        );
    }

    private Map<String, String> unzipTextEntries(byte[] docxBytes) throws IOException {
        Map<String, String> entries = new HashMap<>();
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(docxBytes), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (!entry.isDirectory()) {
                    entries.put(entry.getName(), new String(zis.readAllBytes(), StandardCharsets.UTF_8));
                }
            }
        }
        return entries;
    }

    private String extractRelationshipTarget(String relsXml, String relId) {
        Matcher matcher = Pattern.compile(
            "<Relationship[^>]*Id=\"" + Pattern.quote(relId) + "\"[^>]*Target=\"([^\"]+)\"")
            .matcher(relsXml);
        return matcher.find() ? matcher.group(1) : null;
    }

    private double parseStyleMetric(String style, String key) {
        Matcher matcher = Pattern.compile(key + ":([0-9.]+)pt").matcher(style);
        assertTrue(matcher.find(), "missing style metric " + key);
        return Double.parseDouble(matcher.group(1));
    }

    private record Metrics(
        int positionHalfPt,
        int dxaOrig,
        int dyaOrig,
        double widthPt,
        double heightPt,
        String previewExtension
    ) {
    }
}
