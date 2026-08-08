package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import com.lz.paperword.core.render.SvgVectorEmfPlusRenderer;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ThreeLevelVectorEmfPlusAcceptanceTest {

    private static final Path REPORT = Path.of(
        "target/vector-acceptance/three-level-vector-emfplus.json");

    @Test
    void everyThreeLevelCombinationProducesStrictEmfPlusDual() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.threeLevelVector"),
            "Enable with -Dpaperword.acceptance.threeLevelVector=true");
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

        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        List<Failure> failures = new ArrayList<>();
        int generated = 0;
        int strictDual = 0;
        for (StructureFamily outer : families) {
            for (StructureFamily middle : families) {
                for (StructureFamily inner : families) {
                    for (Map.Entry<String, String> leaf : leaves.entrySet()) {
                        generated++;
                        String formula = outer.wrap().apply(middle.wrap().apply(inner.wrap().apply(leaf.getValue())));
                        try {
                            LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(formula);
                            if ("emf".equals(preview.extension())
                                    && "image/x-emf".equals(preview.contentType())
                                    && preview.widthPt() > 0d && preview.heightPt() > 0d
                                    && SvgVectorEmfPlusRenderer.isValidDualVector(preview.data())) {
                                strictDual++;
                            } else {
                                failures.add(new Failure(formula, "INVALID_EMFPLUS_DUAL",
                                    preview.extension() + "/" + preview.contentType()));
                            }
                        } catch (RuntimeException exception) {
                            failures.add(new Failure(formula, "EXCEPTION", exception.toString()));
                        }
                    }
                }
            }
        }

        Files.createDirectories(REPORT.getParent());
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("generatedCombinationCount", generated);
        report.put("strictEmfPlusDualCount", strictDual);
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT).writeValue(REPORT.toFile(), report);

        assertEquals(6_000, generated);
        assertEquals(6_000, strictDual);
        assertEquals(0, failures.size(), () ->
            failures.size() + " EMF+ combination failures; inspect " + REPORT.toAbsolutePath());
    }

    private static StructureFamily family(String name, Function<String, String> wrap) {
        return new StructureFamily(name, wrap);
    }

    private record StructureFamily(String name, Function<String, String> wrap) {}

    private record Failure(String formula, String code, String detail) {}
}
