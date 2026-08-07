package com.lz.paperword.core.latex;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.mtef.MtefWriter;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class OfficialLatexCoverageCompatibilityTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path REPORT = Path.of("target/official-latex-coverage/current-report.json");

    private final ObjectMapper objectMapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
    private final LaTeXParser parser = new LaTeXParser();
    private final MtefWriter writer = new MtefWriter();

    @Test
    void everyOfficialEntryProducesSupportedIrAndMtef() throws Exception {
        JsonNode corpus = objectMapper.readTree(CORPUS.toFile());
        List<Failure> failures = new ArrayList<>();
        Set<String> uniqueExamples = new LinkedHashSet<>();
        int successfulEntries = 0;

        for (JsonNode entry : corpus.path("entries")) {
            String command = entry.path("command").asText();
            boolean entryPassed = true;
            for (JsonNode exampleNode : entry.path("examples")) {
                String example = exampleNode.asText();
                uniqueExamples.add(example);
                LaTeXParser.DetailedParseResult parsed;
                try {
                    parsed = parser.parseDetailed(example);
                } catch (RuntimeException exception) {
                    failures.add(new Failure(command, example, "PARSE_EXCEPTION", exception.toString()));
                    entryPassed = false;
                    continue;
                }
                String expectedToken = command.startsWith("\\begin{") ? "\\begin" : command;
                if (!parsed.consumedCommands().contains(expectedToken)) {
                    failures.add(new Failure(command, example, "COMMAND_NOT_CONSUMED", expectedToken));
                    entryPassed = false;
                }
                for (LaTeXParser.ParseDiagnostic diagnostic : parsed.diagnostics()) {
                    if (diagnostic.severity() == LaTeXParser.DiagnosticSeverity.ERROR) {
                        failures.add(new Failure(command, example, diagnostic.code(), diagnostic.message()));
                        entryPassed = false;
                    }
                }
                if (!parsed.isSupported()) {
                    continue;
                }
                try {
                    MtefWriter.WriteReport written = writer.writeWithReport(parsed.mathIR());
                    assertTrue(written.bytes().length > 12);
                    assertTrue(written.normalization().recordCounts().getOrDefault("LINE", 0) > 0);
                } catch (RuntimeException | AssertionError exception) {
                    failures.add(new Failure(command, example, "MTEF_WRITE_FAILURE", exception.toString()));
                    entryPassed = false;
                }
            }
            if (entryPassed) {
                successfulEntries++;
            }
        }

        writeReport(corpus.path("entryCount").asInt(), successfulEntries, uniqueExamples.size(), failures);
        assertEquals(0, failures.size(), () ->
            failures.size() + " official coverage failures; inspect " + REPORT.toAbsolutePath());
    }

    private void writeReport(int totalEntries, int successfulEntries, int uniqueExamples, List<Failure> failures)
        throws Exception {
        Files.createDirectories(REPORT.getParent());
        Map<String, Integer> failuresByCode = new LinkedHashMap<>();
        for (Failure failure : failures) {
            failuresByCode.merge(failure.code(), 1, Integer::sum);
        }
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("totalEntries", totalEntries);
        report.put("successfulEntries", successfulEntries);
        report.put("failedEntries", totalEntries - successfulEntries);
        report.put("uniqueExamples", uniqueExamples);
        report.put("failureCount", failures.size());
        report.put("failuresByCode", failuresByCode);
        report.put("failures", failures);
        objectMapper.writeValue(REPORT.toFile(), report);
    }

    private record Failure(String command, String example, String code, String detail) {
    }
}
