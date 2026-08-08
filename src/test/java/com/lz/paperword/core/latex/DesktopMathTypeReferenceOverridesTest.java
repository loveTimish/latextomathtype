package com.lz.paperword.core.latex;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class DesktopMathTypeReferenceOverridesTest {

    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void referenceOverridesAreVersionedClassifiedAndParseable() throws Exception {
        JsonNode corpus = mapper.readTree(Path.of(
            "docs/reference/mathtype/latex-coverage-corpus.json").toFile());
        JsonNode overrides = mapper.readTree(Path.of(
            "docs/reference/mathtype/desktop-reference-overrides.json").toFile());
        assertEquals(1, overrides.path("schemaVersion").asInt());

        Set<String> officialExamples = new HashSet<>();
        for (JsonNode entry : corpus.path("entries")) {
            for (JsonNode example : entry.path("examples")) {
                officialExamples.add(example.asText());
            }
        }

        Set<String> formulas = new HashSet<>();
        for (JsonNode override : overrides.path("overrides")) {
            String formula = override.path("formula").asText();
            String mode = override.path("referenceMode").asText();
            assertTrue(formulas.add(formula), () -> "duplicate override: " + formula);
            assertTrue(officialExamples.contains(formula), () -> "not an official example: " + formula);
            assertFalse(override.path("reason").asText().isBlank(), () -> "missing reason: " + formula);
            assertTrue(mode.equals("tex-toggle-equivalent") || mode.equals("native-template"),
                () -> "invalid reference mode for " + formula + ": " + mode);

            if (mode.equals("tex-toggle-equivalent")) {
                String reference = override.path("referenceFormula").asText();
                assertFalse(reference.isBlank(), () -> "missing equivalent: " + formula);
                assertFalse(reference.contains("\\boxed") || reference.contains("\\newline")
                        || reference.matches(".*\\\\sf\\s.*")
                        || reference.contains("\\xLongleftarrow")
                        || reference.contains("\\largeT"),
                    () -> "equivalent still contains a known literal Texvc command: " + reference);
            } else {
                if (formula.contains("\\boxed")) {
                    assertEquals(37, override.path("selector").asInt(), "native box selector");
                    assertEquals(30, override.path("variation").asInt(), "native four-sided box variation");
                }
            }
        }
        assertEquals(131, formulas.size());
    }
}
