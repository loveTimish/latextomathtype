package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Enumeration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class UnifiedBoxFormulaCandidateDocxTest {

    @Test
    void shouldGenerateUnifiedBoxFormulaCandidateDocx() throws IOException {
        byte[] docx = new DocxBuilder().build(buildRequest());

        Path outputDir = Path.of("analysis", "trace-runs", "20260616-unified-box-candidate");
        Files.createDirectories(outputDir);

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path output = outputDir.resolve("unified-box-formula-candidate-" + timestamp + ".docx");
        Files.write(output, docx);

        Path latest = outputDir.resolve("unified-box-formula-candidate.docx");
        try {
            Files.copy(output, latest, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
            // The timestamped file is the authoritative artifact if the alias is open in Word.
        }

        System.out.println("Generated unified-box formula candidate: " + output.toAbsolutePath());
        assertTrue(Files.exists(output), "generated docx should exist");
        assertTrue(Files.size(output) > 1000, "generated docx should not be empty");
        assertFormulaEmbeddingsAreValid(output, 19);
    }

    private PaperExportRequest buildRequest() {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("统一盒模型公式候选样例");
        paper.setScore(100);
        paper.setSuggestTime(45);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、公式结构样例");
        section.setQuestions(List.of(
            question(1, """
                【线性与上下标】<br/>
                $S_{\\Delta AOB}:S_{\\Delta COD}=a^{2}:b^{2}=4:9$<br/>
                $a_{1}^{2}+b_{2}^{3}=c^{2}$<br/>
                $AO=1,\\ CO=\\frac{5}{3}$
                """),
            question(2, """
                【普通分式与面积链】<br/>
                $S_{\\Delta AOB}:S_{\\Delta COD}=a^{2}:b^{2}=4:9$，所以 $a:b=2:3$，<br/>
                $S_{\\Delta AOD}=S_{\\Delta COD}=1.2\\times\\frac{3}{2}=1.8$，<br/>
                $S_{\\text{梯形}ABCD}=1.2+1.8+1.8+2.7=7.5$
                """),
            question(3, """
                【中文长文本比例，避免竖向撑高】<br/>
                三角形$ABD$的面积/三角形$CBD$的面积$=\\frac{AO}{CO}$，所以 $\\frac{AO}{CO}=\\frac{3}{5}$，又 $AO=1$，所以 $CO=\\frac{5}{3}$。
                """),
            question(4, """
                【嵌套分式与上下标】<br/>
                $x^{\\frac{1}{2}}+a_{1}^{\\frac{2}{3}}$<br/>
                $\\frac{1+\\frac{a}{b}}{2+\\frac{c}{d}}=\\frac{b+a}{2b+\\frac{bc}{d}}$
                """),
            question(5, """
                【根号与根号内嵌套】<br/>
                $\\sqrt{8}=2\\sqrt{2}$<br/>
                $\\sqrt{1+\\frac{a_{1}^{2}}{\\frac{3}{5}}}$
                """),
            question(6, """
                【手绘线条结构】<br/>
                $\\overline{AB}=\\underline{CD}$<br/>
                $\\boxed{x+1}=\\frac{\\sqrt{a^{2}+b^{2}}}{2}$
                """)
        ));
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

    private void assertFormulaEmbeddingsAreValid(Path docx, int expectedFormulaCount) throws IOException {
        int vectorPreviewCount = 0;
        int equationDsmt4Count = 0;
        int oleCount = 0;
        try (ZipFile zip = new ZipFile(docx.toFile())) {
            ZipEntry documentXml = zip.getEntry("word/document.xml");
            assertTrue(documentXml != null, "document.xml should exist");
            String xml = new String(zip.getInputStream(documentXml).readAllBytes());
            assertFalse(xml.contains("\\frac"), "LaTeX fraction command must not leak into visible XML");
            assertFalse(xml.contains("\\sqrt"), "LaTeX sqrt command must not leak into visible XML");
            assertFalse(xml.contains("$"), "formula delimiters must not leak into visible XML");

            Enumeration<? extends ZipEntry> entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                if (name.startsWith("word/media/") && name.endsWith(".emf")) {
                    vectorPreviewCount++;
                } else if (name.startsWith("word/embeddings/") && name.endsWith(".bin")) {
                    oleCount++;
                    byte[] ole = zip.getInputStream(entry).readAllBytes();
                    if (containsAscii(ole, "Equation.DSMT4")) {
                        equationDsmt4Count++;
                    }
                }
            }
        }
        assertEquals(expectedFormulaCount, vectorPreviewCount,
            "each formula should have an EMF+ Dual vector preview");
        assertEquals(expectedFormulaCount, oleCount, "each formula should have an OLE embedding");
        assertEquals(expectedFormulaCount, equationDsmt4Count, "each OLE should be MathType Equation.DSMT4");
    }

    private boolean containsAscii(byte[] data, String needle) {
        byte[] bytes = needle.getBytes(java.nio.charset.StandardCharsets.US_ASCII);
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
}
