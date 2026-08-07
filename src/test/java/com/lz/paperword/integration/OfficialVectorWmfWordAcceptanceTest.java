package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class OfficialVectorWmfWordAcceptanceTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path OUTPUT = Path.of(
        "target/vector-acceptance/official-238-vector-wmf.docx");

    @Test
    void generateSingleWordDocumentWithEveryOfficialVectorPreview() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.word"),
            "Enable with -Dpaperword.acceptance.word=true");
        JsonNode corpus = new ObjectMapper().readTree(CORPUS.toFile());
        Set<String> examples = new LinkedHashSet<>();
        corpus.path("entries").forEach(entry -> entry.path("examples")
            .forEach(example -> examples.add(example.asText())));

        List<QuestionDTO> questions = new ArrayList<>();
        int index = 0;
        for (String example : examples) {
            index++;
            QuestionDTO question = new QuestionDTO();
            question.setSerialNumber(index);
            question.setQuestionType(6);
            question.setScore(1);
            question.setContent("$" + LaTeXParser.preNormalizeLatex(example) + "$");
            questions.add(question);
        }

        PaperExportRequest request = new PaperExportRequest();
        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("Official LaTeX Pure Vector WMF Acceptance");
        paper.setScore(questions.size());
        paper.setSuggestTime(30);
        request.setPaper(paper);
        SectionDTO section = new SectionDTO();
        section.setHeadline("MathJax SVG to Batik POLYPOLYGON WMF - 238 official examples");
        section.setQuestions(questions);
        request.setSections(List.of(section));

        Files.createDirectories(OUTPUT.getParent());
        Files.write(OUTPUT, new DocxBuilder(true).build(request));
        assertEquals(238, questions.size());
        assertTrue(Files.size(OUTPUT) > 1000);
        System.out.println("Generated official vector Word acceptance document: " + OUTPUT.toAbsolutePath());
    }
}
