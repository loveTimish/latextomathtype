package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Mines native MathType objects that use a selected non-default template variation. */
public final class MtefTemplateVariationMinerCli {

    private MtefTemplateVariationMinerCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("usage: MtefTemplateVariationMinerCli <docx-dir> <selector> <out.json>");
            System.exit(2);
        }
        Path source = Path.of(args[0]).toAbsolutePath();
        int selector = Integer.parseInt(args[1]);
        Path output = Path.of(args[2]).toAbsolutePath();
        List<Path> documents;
        try (var paths = Files.list(source)) {
            documents = paths
                .filter(path -> path.getFileName().toString().toLowerCase().endsWith(".docx"))
                .sorted(Comparator.comparingInt(MtefTemplateVariationMinerCli::numericStem))
                .toList();
        }

        MathTypeDocxComparator inspector = new MathTypeDocxComparator();
        List<Map<String, Object>> matches = new ArrayList<>();
        int objectCount = 0;
        for (Path document : documents) {
            MathTypeDocxComparator.InspectionReport report = inspector.inspect(document);
            objectCount += report.formulas().size();
            for (int index = 0; index < report.formulas().size(); index++) {
                MathTypeDocxComparator.InspectedFormula formula = report.formulas().get(index);
                List<Integer> variations = formula.records().stream()
                    .filter(record -> "TMPL".equals(record.name()))
                    .filter(record -> record.selector() != null && record.selector() == selector)
                    .map(MtefRecordNormalizer.CanonicalRecord::variation)
                    .filter(variation -> variation != null && variation != 0)
                    .distinct()
                    .toList();
                if (variations.isEmpty()) {
                    continue;
                }
                Map<String, Object> match = new LinkedHashMap<>();
                match.put("document", document.toString());
                match.put("formulaIndex", index);
                match.put("entry", formula.entry());
                match.put("variations", variations);
                match.put("records", formula.records());
                match.put("mtefHex", formula.mtefHex());
                matches.add(match);
            }
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put("schemaVersion", 1);
        result.put("source", source.toString());
        result.put("selector", selector);
        result.put("documentCount", documents.size());
        result.put("objectCount", objectCount);
        result.put("matchCount", matches.size());
        result.put("matches", matches);
        if (output.getParent() != null) {
            Files.createDirectories(output.getParent());
        }
        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT).writeValue(output.toFile(), result);
        System.out.printf("documents=%d objects=%d matches=%d report=%s%n",
            documents.size(), objectCount, matches.size(), output);
    }

    private static int numericStem(Path path) {
        String name = path.getFileName().toString().replaceFirst("\\.[^.]+$", "");
        try {
            return Integer.parseInt(name);
        } catch (NumberFormatException ignored) {
            return Integer.MAX_VALUE;
        }
    }
}
