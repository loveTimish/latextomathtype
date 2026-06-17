package com.lz.paperword.tools;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class GeneratedMathTypeTexReferenceDocxTest {

    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void shouldGenerateCurrentRendererDocxForWordMathTypeTexReference() throws IOException {
        Path sourceReport = Path.of(
            "analysis",
            "wmf-structure-metrics",
            "word-mathtype-tex-reference-expanded.source-report.json"
        );
        PaperExportRequest request = buildRequest(sourceReport);
        byte[] docx = new DocxBuilder().build(request);

        Path outputDir = Path.of("analysis", "wmf-structure-metrics", "generated-word-mathtype-tex-reference-expanded");
        Files.createDirectories(outputDir);
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path docxPath = outputDir.resolve("generated-word-mathtype-tex-reference-expanded-" + timestamp + ".docx");
        Path requestPath = outputDir.resolve("generated-word-mathtype-tex-reference-expanded-" + timestamp + ".request.json");
        Files.write(docxPath, docx);
        mapper.writerWithDefaultPrettyPrinter().writeValue(requestPath.toFile(), request);

        Files.write(outputDir.resolve("generated-word-mathtype-tex-reference-expanded.latest.txt"),
            List.of(docxPath.toAbsolutePath().toString(), requestPath.toAbsolutePath().toString()));

        assertTrue(Files.size(docxPath) > 1000, "generated DOCX should not be empty");
        assertEquals(20, request.getSections().get(0).getQuestions().size(), "current renderer comparable formula count should stay stable");
        System.out.println("Generated current renderer MathType reference DOCX: " + docxPath.toAbsolutePath());
        System.out.println("Generated request JSON: " + requestPath.toAbsolutePath());
    }

    private PaperExportRequest buildRequest(Path sourceReport) throws IOException {
        JsonNode root = mapper.readTree(sourceReport.toFile());
        List<QuestionDTO> questions = new ArrayList<>();
        for (JsonNode equation : root.path("equations")) {
            String formula = equation.path("output").asText("");
            if (formula.isBlank()) {
                continue;
            }
            if (formula.contains("\\sqrt[")) {
                continue;
            }
            int index = equation.path("formulaIndex").asInt(questions.size()) + 1;
            String rendererFormula = normalizeForCurrentRenderer(formula);
            QuestionDTO question = new QuestionDTO();
            question.setSerialNumber(index);
            question.setQuestionType(6);
            question.setScore(1);
            question.setContent("$" + rendererFormula + "$");
            questions.add(question);
        }

        PaperExportRequest request = new PaperExportRequest();
        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("Generated Word MathType TeX Reference Expanded");
        paper.setScore(questions.size());
        paper.setSuggestTime(10);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("Formula reference");
        section.setQuestions(questions);
        request.setSections(List.of(section));
        return request;
    }

    private String normalizeForCurrentRenderer(String formula) {
        return formula
            .replace("\\dfrac", "\\frac")
            .replace("\\cfrac", "\\frac");
    }
}
