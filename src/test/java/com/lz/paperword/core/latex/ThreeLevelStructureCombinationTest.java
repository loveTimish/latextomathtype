package com.lz.paperword.core.latex;

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
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ThreeLevelStructureCombinationTest {

    private static final Path REPORT = Path.of(
        "target/official-latex-coverage/three-level-combinations.json");
    private final ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
    private final LaTeXParser parser = new LaTeXParser();
    private final MtefWriter writer = new MtefWriter();

    @Test
    void everyOrderedThreeLevelStructureCombinationProducesEditableMtef() throws Exception {
        List<StructureFamily> families = List.of(
            family("fraction", value -> "\\frac{" + value + "}{7}"),
            family("root", value -> "\\sqrt[3]{" + value + "}"),
            family("script", value -> "{" + value + "}_{i}^{2}"),
            family("fence", value -> "\\left(" + value + "\\right)"),
            family("accent", value -> "\\wideparen{" + value + "}"),
            family("over-under", value -> "\\overset{a}{\\underset{b}{" + value + "}}"),
            family("big-operator", value -> "\\sum_{i=0}^{" + value + "}x_i"),
            family("annotated-arrow", value -> "A\\xrightarrow{" + value + "}B"),
            family("matrix", value -> "\\begin{pmatrix}" + value + "&1\\\\2&3\\end{pmatrix}"),
            family("horizontal-brace", value -> "\\underbrace{" + value + "}_{n}")
        );
        Map<String, String> leaves = new LinkedHashMap<>();
        leaves.put("symbol", "\\alpha");
        leaves.put("relation", "x\\leq y");
        leaves.put("operator", "a+b\\times c");
        leaves.put("function", "\\sin x");
        leaves.put("number", "123");
        leaves.put("text", "\\text{rate}");

        List<Failure> failures = new ArrayList<>();
        List<Exclusion> exclusions = new ArrayList<>();
        Set<String> formulas = new LinkedHashSet<>();
        Set<String> normalizedStructures = new LinkedHashSet<>();
        int generated = 0;

        for (StructureFamily outer : families) {
            for (StructureFamily middle : families) {
                for (StructureFamily inner : families) {
                    for (Map.Entry<String, String> leaf : leaves.entrySet()) {
                        generated++;
                        String formula = outer.wrap().apply(middle.wrap().apply(inner.wrap().apply(leaf.getValue())));
                        if (!formulas.add(formula)) {
                            exclusions.add(new Exclusion(outer.name(), middle.name(), inner.name(), leaf.getKey(),
                                "duplicate-formula", formula));
                            continue;
                        }
                        try {
                            LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(formula);
                            if (!parsed.isSupported()) {
                                failures.add(new Failure(outer.name(), middle.name(), inner.name(), leaf.getKey(),
                                    formula, "PARSE_DIAGNOSTIC", parsed.diagnostics().toString()));
                                continue;
                            }
                            MtefWriter.WriteReport report = writer.writeWithReport(parsed.mathIR());
                            normalizedStructures.add(report.normalization().canonicalSignature());
                            if (report.bytes().length <= 12
                                    || report.normalization().recordCounts().getOrDefault("LINE", 0) == 0) {
                                failures.add(new Failure(outer.name(), middle.name(), inner.name(), leaf.getKey(),
                                    formula, "INVALID_MTEF", report.normalization().canonicalSignature()));
                            }
                        } catch (RuntimeException exception) {
                            failures.add(new Failure(outer.name(), middle.name(), inner.name(), leaf.getKey(),
                                formula, "EXCEPTION", exception.toString()));
                        }
                    }
                }
            }
        }

        Files.createDirectories(REPORT.getParent());
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("structureFamilies", families.stream().map(StructureFamily::name).toList());
        report.put("leafClasses", leaves.keySet());
        report.put("generatedCombinationCount", generated);
        report.put("executedFormulaCount", formulas.size());
        report.put("normalizedStructureCount", normalizedStructures.size());
        report.put("exclusionCount", exclusions.size());
        report.put("exclusions", exclusions);
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        mapper.writeValue(REPORT.toFile(), report);

        assertEquals(6_000, generated);
        assertEquals(0, failures.size(), () ->
            failures.size() + " three-level structure failures; inspect " + REPORT.toAbsolutePath());
    }

    private StructureFamily family(String name, Function<String, String> wrap) {
        return new StructureFamily(name, wrap);
    }

    private record StructureFamily(String name, Function<String, String> wrap) {}

    private record Exclusion(
        String outer, String middle, String inner, String leafClass, String reason, String formula
    ) {}

    private record Failure(
        String outer, String middle, String inner, String leafClass, String formula, String code, String detail
    ) {}
}
