package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Compares every MathType-formatted generated batch in an official manifest. */
public final class OfficialMathTypeManifestComparatorCli {

    private OfficialMathTypeManifestComparatorCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 1) {
            System.err.println("usage: OfficialMathTypeManifestComparatorCli <manifest.json>");
            System.exit(2);
        }
        Path manifestPath = Path.of(args[0]).toAbsolutePath();
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        JsonNode manifest = mapper.readTree(manifestPath.toFile());
        MathTypeDocxComparator comparator = new MathTypeDocxComparator();
        List<Map<String, Object>> batches = new ArrayList<>();
        List<Map<String, Object>> failures = new ArrayList<>();
        int expected = 0;
        int generated = 0;
        int valid = 0;
        int equal = 0;

        for (JsonNode batch : manifest.path("batches")) {
            Path standard = Path.of(batch.path("standardDocx").asText());
            Path formatted = standard.getParent().resolve("generated-formatted.docx");
            JsonNode formulas = mapper.readTree(Path.of(batch.path("formulas").asText()).toFile());
            MathTypeDocxComparator.ComparisonReport comparison = comparator.compare(standard, formatted);
            Path comparisonPath = standard.getParent().resolve("formatted-structure-comparison.json");
            mapper.writeValue(comparisonPath.toFile(), comparison);
            expected += comparison.standardObjectCount();
            generated += comparison.generatedObjectCount();
            valid += comparison.validGeneratedOleCount();
            equal += comparison.normalizedStructureEqualCount();
            batches.add(Map.of(
                "name", batch.path("name").asText(),
                "formulaCount", formulas.size(),
                "comparison", comparisonPath.toString(),
                "passesExactStructureGate", comparison.passesExactStructureGate()
            ));
            for (MathTypeDocxComparator.FormulaComparison item : comparison.formulas()) {
                if (!item.oleValid() || !item.normalizedStructureEqual()) {
                    JsonNode source = item.index() < formulas.size() ? formulas.get(item.index()) : null;
                    Map<String, Object> failure = new LinkedHashMap<>();
                    failure.put("batch", batch.path("name").asText());
                    failure.put("index", item.index());
                    failure.put("formula", source == null ? "" : source.path("formula").asText());
                    failure.put("oleValid", item.oleValid());
                    failure.put("normalizedStructureEqual", item.normalizedStructureEqual());
                    failure.put("standardLiteralLatexCommands", item.standardLiteralLatexCommands());
                    failures.add(failure);
                }
            }
        }

        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("officialEntryCount", manifest.path("officialEntryCount").asInt());
        report.put("selectedFormulaCount", manifest.path("selectedFormulaCount").asInt());
        report.put("expectedObjectCount", expected);
        report.put("generatedObjectCount", generated);
        report.put("validOleCount", valid);
        report.put("normalizedStructureEqualCount", equal);
        report.put("batches", batches);
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        Path reportPath = manifestPath.getParent().resolve("final-formatted-structure-report.json");
        mapper.writeValue(reportPath.toFile(), report);
        System.out.printf("objects=%d valid=%d exactStructure=%d report=%s%n",
            generated, valid, equal, reportPath);
        if (expected != generated || generated != valid || generated != equal) {
            throw new IllegalStateException("formatted MathType structure gate failed; inspect " + reportPath);
        }
    }
}
