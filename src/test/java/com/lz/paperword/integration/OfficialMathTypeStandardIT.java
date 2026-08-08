package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.docx.MathTypeEmbedder;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MathTypeDocxComparator;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class OfficialMathTypeStandardIT {

    private final ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
    private final LaTeXParser parser = new LaTeXParser();
    private final MathTypeDocxComparator comparator = new MathTypeDocxComparator();

    @Test
    void generatedOleMatchesPinnedMathTypeStandardByNormalizedStructure() throws Exception {
        Path manifestPath = Path.of(System.getProperty(
            "mathtype.standard.manifest", "target/mathtype-standard/manifest.json"));
        JsonNode manifest = mapper.readTree(manifestPath.toFile());
        boolean generateOnly = Boolean.getBoolean("mathtype.standard.generateOnly");
        boolean compareFormatted = Boolean.getBoolean("mathtype.standard.compareFormatted");
        List<Map<String, Object>> batchReports = new ArrayList<>();
        List<Map<String, Object>> failures = new ArrayList<>();
        int expectedObjects = 0;
        int generatedObjects = 0;
        int validOle = 0;
        int structureEqual = 0;
        int rawEqual = 0;
        int contaminatedStandards = 0;

        for (JsonNode batch : manifest.path("batches")) {
            Path formulasPath = Path.of(batch.path("formulas").asText());
            Path standardDocx = Path.of(batch.path("standardDocx").asText());
            Path generatedDocx = standardDocx.getParent().resolve("generated.docx");
            JsonNode formulas = mapper.readTree(formulasPath.toFile());
            generateDocx(formulas, generatedDocx);
            if (generateOnly) {
                batchReports.add(Map.of(
                    "name", batch.path("name").asText(),
                    "formulaCount", formulas.size(),
                    "generatedDocx", generatedDocx.toAbsolutePath().toString(),
                    "status", "generated-awaiting-mathtype-format"
                ));
                continue;
            }

            Path comparisonDocx = compareFormatted
                ? standardDocx.getParent().resolve("generated-formatted.docx")
                : generatedDocx;
            assertTrue(Files.isRegularFile(comparisonDocx), () -> "missing comparison DOCX: " + comparisonDocx);
            MathTypeDocxComparator.ComparisonReport comparison = comparator.compare(standardDocx, comparisonDocx);
            Path comparisonPath = standardDocx.getParent().resolve("structure-comparison.json");
            mapper.writeValue(comparisonPath.toFile(), comparison);
            expectedObjects += comparison.standardObjectCount();
            generatedObjects += comparison.generatedObjectCount();
            validOle += comparison.validGeneratedOleCount();
            structureEqual += comparison.normalizedStructureEqualCount();
            rawEqual += comparison.rawMtefEqualCount();
            batchReports.add(Map.of(
                "name", batch.path("name").asText(),
                "formulaCount", formulas.size(),
                "comparison", comparisonPath.toAbsolutePath().toString(),
                "passesExactStructureGate", comparison.passesExactStructureGate()
            ));
            for (MathTypeDocxComparator.FormulaComparison formula : comparison.formulas()) {
                if (!formula.standardLiteralLatexCommands().isEmpty()) {
                    contaminatedStandards++;
                }
                if (!formula.oleValid() || !formula.normalizedStructureEqual()) {
                    JsonNode source = formula.index() < formulas.size() ? formulas.get(formula.index()) : null;
                    Map<String, Object> failure = new LinkedHashMap<>();
                    failure.put("batch", batch.path("name").asText());
                    failure.put("batchFormulaIndex", formula.index());
                    failure.put("globalFormulaIndex", source == null ? -1 : source.path("globalIndex").asInt());
                    failure.put("formula", source == null ? "" : source.path("formula").asText());
                    failure.put("oleValid", formula.oleValid());
                    failure.put("normalizedStructureEqual", formula.normalizedStructureEqual());
                    failure.put("generatedError", formula.generatedError());
                    failure.put("standardLiteralLatexCommands", formula.standardLiteralLatexCommands());
                    failure.put("generatedLiteralLatexCommands", formula.generatedLiteralLatexCommands());
                    failures.add(failure);
                }
            }
        }

        Map<String, Object> finalReport = new LinkedHashMap<>();
        finalReport.put("schemaVersion", 1);
        finalReport.put("officialEntryCount", manifest.path("officialEntryCount").asInt());
        finalReport.put("selectedFormulaCount", manifest.path("selectedFormulaCount").asInt());
        finalReport.put("generationOnly", generateOnly);
        finalReport.put("expectedObjectCount", expectedObjects);
        finalReport.put("generatedObjectCount", generatedObjects);
        finalReport.put("validOleCount", validOle);
        finalReport.put("normalizedStructureEqualCount", structureEqual);
        finalReport.put("rawMtefEqualCount", rawEqual);
        finalReport.put("contaminatedStandardCount", contaminatedStandards);
        finalReport.put("batches", batchReports);
        finalReport.put("failureCount", failures.size());
        finalReport.put("failures", failures);
        Path reportPath = manifestPath.getParent().resolve(
            compareFormatted ? "final-formatted-structure-report.json" : "final-structure-report.json");
        mapper.writeValue(reportPath.toFile(), finalReport);

        assertEquals(expectedObjects, generatedObjects, "standard/generated OLE object count differs");
        assertEquals(generatedObjects, validOle, "every generated object must be valid Equation.DSMT4 OLE");
        assertEquals(generatedObjects, structureEqual,
            () -> failures.size() + " normalized MTEF mismatches; inspect " + reportPath.toAbsolutePath());
    }

    private void generateDocx(JsonNode formulas, Path output) throws Exception {
        Files.createDirectories(output.getParent());
        MathTypeEmbedder embedder = new MathTypeEmbedder();
        embedder.resetDocumentFormulaCounter();
        try (XWPFDocument document = new XWPFDocument()) {
            for (JsonNode item : formulas) {
                String latex = item.path("formula").asText();
                LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(latex);
                assertTrue(parsed.isSupported(), () -> latex + ": " + parsed.diagnostics());
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                embedder.embedEquation(paragraph, run, parsed.ast(), latex);
            }
            try (OutputStream stream = Files.newOutputStream(output)) {
                document.write(stream);
            }
        }
    }

}
