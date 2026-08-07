package com.lz.paperword.core.latex;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class WordMathTypeFormatLimitationsTest {

    @Test
    void limitationsAreVersionedOfficialAndEvidenceBacked() throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        JsonNode corpus = mapper.readTree(Path.of(
            "docs/reference/mathtype/latex-coverage-corpus.json").toFile());
        JsonNode limitations = mapper.readTree(Path.of(
            "docs/reference/mathtype/word-format-limitations.json").toFile());
        assertEquals(1, limitations.path("schemaVersion").asInt());
        assertEquals("7.11.1.462", limitations.path("environment").path("mathTypeVersion").asText());

        Map<String, Integer> official = new HashMap<>();
        Set<String> seen = new HashSet<>();
        int index = 0;
        for (JsonNode entry : corpus.path("entries")) {
            for (JsonNode example : entry.path("examples")) {
                String formula = example.asText();
                if (seen.add(formula)) {
                    official.put(formula, index++);
                }
            }
        }

        Set<String> limited = new HashSet<>();
        Map<String, Integer> countsByMode = new HashMap<>();
        for (JsonNode limitation : limitations.path("limitations")) {
            String formula = limitation.path("formula").asText();
            assertTrue(limited.add(formula), () -> "duplicate limitation: " + formula);
            assertTrue(official.containsKey(formula), () -> "not an official formula: " + formula);
            assertEquals(official.get(formula).intValue(), limitation.path("globalIndex").asInt());
            String mode = limitation.path("mode").asText();
            assertTrue(Set.of("preserve", "selection").contains(mode), () -> "invalid mode: " + mode);
            countsByMode.merge(mode, 1, Integer::sum);
            assertFalse(limitation.path("evidence").asText().isBlank());
            assertFalse(limitation.path("reason").asText().isBlank());
        }
        for (JsonNode group : limitations.path("selectionGroups")) {
            assertFalse(group.path("name").asText().isBlank());
            assertFalse(group.path("evidence").asText().isBlank());
            assertFalse(group.path("reason").asText().isBlank());
            for (JsonNode globalIndex : group.path("globalIndices")) {
                int value = globalIndex.asInt();
                String formula = official.entrySet().stream()
                    .filter(entry -> entry.getValue() == value)
                    .map(Map.Entry::getKey)
                    .findFirst()
                    .orElseThrow(() -> new AssertionError("not an official global index: " + value));
                assertTrue(limited.add(formula), () -> "duplicate limitation: " + formula);
                countsByMode.merge("selection", 1, Integer::sum);
            }
        }
        for (JsonNode group : limitations.path("shardGroups")) {
            assertFalse(group.path("name").asText().isBlank());
            assertFalse(group.path("evidence").asText().isBlank());
            assertFalse(group.path("reason").asText().isBlank());
            for (JsonNode globalIndex : group.path("globalIndices")) {
                int value = globalIndex.asInt();
                String formula = official.entrySet().stream()
                    .filter(entry -> entry.getValue() == value)
                    .map(Map.Entry::getKey)
                    .findFirst()
                    .orElseThrow(() -> new AssertionError("not an official global index: " + value));
                assertTrue(limited.add(formula), () -> "duplicate limitation: " + formula);
                countsByMode.merge("shard", 1, Integer::sum);
            }
        }
        assertEquals(66, limited.size());
        assertEquals(2, countsByMode.get("preserve"));
        assertEquals(22, countsByMode.get("selection"));
        assertEquals(42, countsByMode.get("shard"));
    }
}
