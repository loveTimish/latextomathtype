package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class FormulaGoldenCorpusDocxTest {

    @Test
    void shouldGenerateGoldenCorpusDocxWithEditableMathTypeOle() throws IOException {
        List<FormulaCase> requiredCases = loadRequiredCases();
        byte[] docx = new DocxBuilder().build(buildRequest(requiredCases));

        Path outputDir = Path.of("analysis", "formula-golden-corpus");
        Files.createDirectories(outputDir);

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path output = outputDir.resolve("formula-golden-corpus-" + timestamp + ".docx");
        Files.write(output, docx);

        try {
            Files.copy(output, outputDir.resolve("formula-golden-corpus-latest.docx"),
                StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
            // Timestamped document is the authoritative artifact if the alias is open in Word.
        }

        assertFormulaEmbeddingsAreValid(output, requiredCases);
        System.out.println("Formula golden corpus DOCX: " + output.toAbsolutePath());
    }

    private List<FormulaCase> loadRequiredCases() throws IOException {
        try (var stream = FormulaGoldenCorpusDocxTest.class.getResourceAsStream("/formula-golden-corpus.tsv")) {
            if (stream == null) {
                throw new IOException("formula-golden-corpus.tsv should be on the test classpath");
            }
            String text = new String(stream.readAllBytes(), StandardCharsets.UTF_8);
            List<FormulaCase> cases = new ArrayList<>();
            String[] lines = text.split("\\R");
            for (int i = 1; i < lines.length; i++) {
                if (lines[i].isBlank()) {
                    continue;
                }
                String[] parts = lines[i].split("\\t", 9);
                if (parts.length == 9 && "required".equals(parts[3])) {
                    cases.add(new FormulaCase(
                        parts[0],
                        parts[1],
                        parts[2],
                        Double.parseDouble(parts[4]),
                        Double.parseDouble(parts[5]),
                        Integer.parseInt(parts[6]),
                        Integer.parseInt(parts[7]),
                        parts[8]
                    ));
                }
            }
            return cases;
        }
    }

    private PaperExportRequest buildRequest(List<FormulaCase> requiredCases) {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("Formula Golden Corpus");
        paper.setScore(100);
        paper.setSuggestTime(45);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、结构族样本");
        List<QuestionDTO> questions = new ArrayList<>();
        for (int i = 0; i < requiredCases.size(); i++) {
            FormulaCase formula = requiredCases.get(i);
            questions.add(question(i + 1,
                "【" + visibleLabel(formula.subject()) + "/" + visibleLabel(formula.family()) + "】case " + (i + 1)
                    + " : $" + formula.latex() + "$"));
        }
        section.setQuestions(questions);
        request.setSections(List.of(section));
        return request;
    }

    private QuestionDTO question(int serialNumber, String content) {
        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(serialNumber);
        question.setQuestionType(6);
        question.setScore(10);
        question.setContent(content);
        return question;
    }

    private void assertFormulaEmbeddingsAreValid(Path docx, List<FormulaCase> expectedCases) throws IOException {
        int expectedFormulaCount = expectedCases.size();
        int wmfCount = 0;
        int oleCount = 0;
        int equationDsmt4Count = 0;
        int vectorWmfCount = 0;
        int bitmapWmfCount = 0;
        List<String> failures = new ArrayList<>();
        try (ZipFile zip = new ZipFile(docx.toFile())) {
            ZipEntry documentXml = zip.getEntry("word/document.xml");
            assertTrue(documentXml != null, "document.xml should exist");
            String xml = new String(zip.getInputStream(documentXml).readAllBytes(), StandardCharsets.UTF_8);
            assertFalse(xml.contains("\\frac"), "LaTeX fraction command must not leak into visible XML");
            assertFalse(xml.contains("\\sqrt"), "LaTeX sqrt command must not leak into visible XML");
            assertFalse(xml.contains("$"), "formula delimiters must not leak into visible XML");
            assertTrue(countOccurrences(xml, "<v:shape") >= expectedFormulaCount,
                "each formula should have at least one visible VML shape");
            assertEquals(expectedFormulaCount, countOccurrences(xml, "<v:imagedata"),
                "each formula shape should reference a WMF preview image");
            assertEquals(expectedFormulaCount, countOccurrences(xml, "<o:OLEObject"),
                "each formula shape should reference an OLE object");

            Enumeration<? extends ZipEntry> entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                if (name.startsWith("word/media/") && name.endsWith(".wmf")) {
                    wmfCount++;
                    FormulaCase expected = expectedCaseForMedia(name, expectedCases);
                    WmfSafety safety = inspectWmf(zip.getInputStream(entry).readAllBytes());
                    if (safety.vectorText() && !safety.hasBitmapFallback()) {
                        vectorWmfCount++;
                    }
                    if (safety.hasBitmapFallback()) {
                        bitmapWmfCount++;
                    }
                    if (expected != null) {
                        addFailureIf(failures, safety.extTextOutRecords() < expected.minTextRecords(),
                            expected.id() + " embedded WMF should draw enough text runs, got "
                                + safety.extTextOutRecords());
                        addFailureIf(failures, safety.polylineRecords() < expected.minPolylineRecords(),
                            expected.id() + " embedded WMF should draw required structure lines, got "
                                + safety.polylineRecords());
                        int expectedRadicals = countLatexCommand(expected.latex(), "\\sqrt");
                        if (expectedRadicals > 0) {
                            addFailureIf(failures, safety.fourPointPolylineRecords() < expectedRadicals,
                                expected.id() + " embedded WMF should preserve each radical as a four-point polyline");
                        }
                    }
                } else if (name.startsWith("word/embeddings/") && name.endsWith(".bin")) {
                    oleCount++;
                    byte[] ole = zip.getInputStream(entry).readAllBytes();
                    if (containsAscii(ole, "Equation.DSMT4")) {
                        equationDsmt4Count++;
                    }
                }
            }
        }
        assertEquals(expectedFormulaCount, wmfCount, "each formula should have a WMF preview");
        assertEquals(expectedFormulaCount, oleCount, "each formula should have an OLE embedding");
        assertEquals(expectedFormulaCount, equationDsmt4Count, "each OLE should be MathType Equation.DSMT4");
        assertEquals(expectedFormulaCount, vectorWmfCount, "each WMF should contain vector ExtTextOut text");
        assertEquals(0, bitmapWmfCount, "golden corpus WMFs must not use bitmap fallback records");
        assertTrue(failures.isEmpty(), String.join(System.lineSeparator(), failures));
    }

    private void addFailureIf(List<String> failures, boolean condition, String message) {
        if (condition) {
            failures.add(message);
        }
    }

    private FormulaCase expectedCaseForMedia(String name, List<FormulaCase> expectedCases) {
        String fileName = name.substring(name.lastIndexOf('/') + 1);
        if (!fileName.startsWith("image_eq") || !fileName.endsWith(".wmf")) {
            return null;
        }
        try {
            int index = Integer.parseInt(fileName.substring("image_eq".length(), fileName.length() - ".wmf".length()));
            return index >= 1 && index <= expectedCases.size() ? expectedCases.get(index - 1) : null;
        } catch (NumberFormatException ignored) {
            return null;
        }
    }

    private int countOccurrences(String text, String needle) {
        int count = 0;
        int index = 0;
        while ((index = text.indexOf(needle, index)) >= 0) {
            count++;
            index += needle.length();
        }
        return count;
    }

    private int countLatexCommand(String latex, String command) {
        int count = 0;
        int index = 0;
        while ((index = latex.indexOf(command, index)) >= 0) {
            int end = index + command.length();
            if (end >= latex.length() || !Character.isLetter(latex.charAt(end))) {
                count++;
            }
            index = end;
        }
        return count;
    }

    private String visibleLabel(String value) {
        return value.replace('_', '-');
    }

    private WmfSafety inspectWmf(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        boolean vectorText = false;
        boolean bitmapFallback = false;
        int extTextOut = 0;
        int polyline = 0;
        int fourPointPolyline = 0;
        int windowExtX = -1;
        int windowExtY = -1;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            int recordEnd = Math.min(data.length, offset + sizeWords * 2);
            if (function == 0x020C && offset + 10 <= recordEnd) {
                windowExtY = word(data, offset + 6);
                windowExtX = word(data, offset + 8);
            } else if (function == 0x0A32) {
                vectorText = true;
                extTextOut++;
            } else if (function == 0x0325 && offset + 8 <= recordEnd) {
                polyline++;
                if (word(data, offset + 6) >= 4) {
                    fourPointPolyline++;
                }
            } else if (isBitmapFunction(function)) {
                bitmapFallback = true;
            }
            offset += sizeWords * 2;
        }
        return new WmfSafety(vectorText, bitmapFallback, extTextOut, polyline, fourPointPolyline,
            windowExtX, windowExtY);
    }

    private boolean isBitmapFunction(int function) {
        return function == 0x0922 || function == 0x0940 || function == 0x0B23 || function == 0x0B41
            || function == 0x0D33 || function == 0x0F43;
    }

    private boolean containsAscii(byte[] data, String needle) {
        byte[] bytes = needle.getBytes(StandardCharsets.US_ASCII);
        outer:
        for (int i = 0; i <= data.length - bytes.length; i++) {
            for (int j = 0; j < bytes.length; j++) {
                if (data[i + j] != bytes[j]) {
                    continue outer;
                }
            }
            return true;
        }
        return false;
    }

    private boolean hasPlaceableHeader(byte[] data) {
        return data.length >= 22 && dword(data, 0) == 0x9AC6CDD7;
    }

    private int word(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort() & 0xffff;
    }

    private int dword(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
    }

    private record WmfSafety(
        boolean vectorText,
        boolean hasBitmapFallback,
        int extTextOutRecords,
        int polylineRecords,
        int fourPointPolylineRecords,
        int windowExtX,
        int windowExtY
    ) {
    }

    private record FormulaCase(
        String id,
        String subject,
        String family,
        double widthPt,
        double heightPt,
        int minTextRecords,
        int minPolylineRecords,
        String latex
    ) {
    }
}
