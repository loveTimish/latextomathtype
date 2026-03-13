package com.lz.paperword.tools;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * One-shot offline generator:
 * - random sample 20 questions from converted_questions.json
 * - generate 5 paper docx files into output directory
 */
class RandomPaperBatchGeneratorTest {

    private static final Path SOURCE_JSON = Path.of("E:/lingzhi/extensions/paper-to-word/data/converted_questions.json");
    private static final Path OUTPUT_DIR = Path.of("E:/lingzhi/extensions/paper-to-word/output");
    private static final int QUESTIONS_PER_PAPER = 20;
    private static final int PAPER_COUNT = 5;

    private final ObjectMapper mapper = new ObjectMapper();
    private final DocxBuilder builder = new DocxBuilder(true);

    @Test
    void generateFiveRandomPapersFromConvertedQuestions() throws IOException {
        Assumptions.assumeTrue(
            Files.exists(SOURCE_JSON),
            "Source json file not found, skip offline generator test: " + SOURCE_JSON
        );

        JsonNode root = mapper.readTree(Files.readString(SOURCE_JSON));
        assertTrue(root.isArray(), "Source json must be an array.");
        assertTrue(root.size() >= QUESTIONS_PER_PAPER,
            "Question count must be >= " + QUESTIONS_PER_PAPER + ", actual: " + root.size());

        Files.createDirectories(OUTPUT_DIR);

        List<JsonNode> all = new ArrayList<>();
        root.forEach(all::add);
        Random random = new Random();
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        for (int i = 1; i <= PAPER_COUNT; i++) {
            Collections.shuffle(all, random);
            List<JsonNode> sample = all.subList(0, QUESTIONS_PER_PAPER);

            PaperExportRequest request = buildRequest(sample, i);
            byte[] docx = builder.build(request);

            Path output = OUTPUT_DIR.resolve("随机20题_第" + i + "套_" + timestamp + ".docx");
            Files.write(output, docx);
            System.out.println("Generated: " + output + " (" + docx.length + " bytes)");
        }
    }

    private PaperExportRequest buildRequest(List<JsonNode> sample, int paperIndex) {
        PaperExportRequest request = new PaperExportRequest();
        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("随机20题_第" + paperIndex + "套");
        paper.setSubjectType(2);
        paper.setStage(2);
        paper.setScore(100);
        paper.setSuggestTime(90);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、随机抽题（20题）");
        section.setQuestions(new ArrayList<>());

        int serial = 1;
        for (JsonNode item : sample) {
            QuestionDTO q = new QuestionDTO();
            q.setSerialNumber(serial++);
            q.setQuestionType(mapQuestionType(item));
            q.setContent(text(item, "stem"));
            q.setCorrect(mapCorrect(item, q.getQuestionType()));
            q.setAnalyze(text(item, "analysis"));
            q.setScore(5);

            if (item.has("options") && item.get("options").isArray()) {
                List<QuestionDTO.OptionDTO> opts = new ArrayList<>();
                for (JsonNode option : item.get("options")) {
                    QuestionDTO.OptionDTO dto = new QuestionDTO.OptionDTO();
                    dto.setPrefix(text(option, "key"));
                    dto.setContent(text(option, "text"));
                    opts.add(dto);
                }
                q.setOptions(opts.isEmpty() ? null : opts);
            }

            section.getQuestions().add(q);
        }

        request.setSections(List.of(section));
        return request;
    }

    private int mapQuestionType(JsonNode item) {
        String target = text(item, "target_question_type").toLowerCase();
        return switch (target) {
            case "single" -> 1;
            case "multiple", "multi" -> 2;
            case "judge" -> 3;
            case "blank" -> 4;
            case "essay" -> 5;
            case "calc", "calculation" -> 6;
            default -> {
                int raw = item.path("source_question_type_raw").asInt(6);
                if (raw >= 1 && raw <= 6) yield raw;
                yield 6;
            }
        };
    }

    private String mapCorrect(JsonNode item, int questionType) {
        if (questionType == 1 || questionType == 2) {
            if (item.has("correct_option_keys") && item.get("correct_option_keys").isArray()) {
                List<String> keys = new ArrayList<>();
                for (JsonNode key : item.get("correct_option_keys")) {
                    keys.add(key.asText(""));
                }
                return String.join(",", keys);
            }
            return "";
        }
        if (questionType == 3) {
            JsonNode judge = item.get("judge_answer");
            if (judge == null || judge.isNull()) return "";
            if (judge.isBoolean()) return judge.asBoolean() ? "正确" : "错误";
            return judge.asText("");
        }
        return text(item, "blank_answer");
    }

    private String text(JsonNode node, String field) {
        JsonNode value = node.get(field);
        if (value == null || value.isNull()) return "";
        return value.asText("");
    }
}

