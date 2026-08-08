package com.lz.paperword.support;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class XscSourceReplacementRepairsTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void appliesOnlyThePinnedDocumentAndFormulaLocations() {
        ObjectNode request = requestWithQuestion(45,
            "$V_{\uFFFD }$:$V_{\uFFFD }$=1:12,$V_{\uFFFD }$:$V_{\uFFFD }$=1:16");

        JsonNode repaired = XscSourceReplacementRepairs.apply(39, request);
        String content = repaired.path("sections").get(0).path("questions").get(45)
            .path("content").asText();

        assertEquals(52, XscSourceReplacementRepairs.occurrenceCount());
        assertEquals(46, XscSourceReplacementRepairs.formulaRepairOccurrenceCount());
        assertEquals(37, XscSourceReplacementRepairs.replacementFormulaOccurrenceCount());
        assertEquals(21, XscSourceReplacementRepairs.documentCount());
        assertFalse(content.contains("\uFFFD"));
        assertTrue(content.contains("V_{\\text{\u7532} }"));
        assertEquals(2, occurrences(content, "V_{\\text{\u8F66} }"));
        assertTrue(content.contains("V_{\\text{\u4E59} }"));
        assertEquals(37, XscSourceReplacementRepairs
            .repairFor(39, "sections[0].questions[45].content#math1")
            .orElseThrow().formulaOrdinal());
    }

    @Test
    void rejectsUncataloguedReplacementCharacters() {
        ObjectNode request = requestWithQuestion(0, "$x+\uFFFD$");

        IllegalArgumentException failure = assertThrows(IllegalArgumentException.class,
            () -> XscSourceReplacementRepairs.apply(1, request));
        assertTrue(failure.getMessage().contains(
            "UNCATALOGUED_SOURCE_REPLACEMENT_CHARACTER"));
    }

    @Test
    void appliesPinnedSemanticPlaceholdersWithoutChangingValidSymbolsGlobally() {
        ObjectNode request = requestWithQuestion(62,
            "$\\frac{\u5317}{\u4EAC}=\\frac{\u5965\u8FD0\u4F1A}"
                + "{\\text{\u5FC3\u60F3\u4E8B\\Theta }}$");
        ObjectNode question = (ObjectNode) request.path("sections").get(0)
            .path("questions").get(62);
        question.put("analyze", "$\\frac{1}{9}=\\frac{\u5965\u8FD0\u4F1A}"
            + "{\\text{\u68A6\u60F3\\Theta \u771F}}$");

        JsonNode repaired = XscSourceReplacementRepairs.apply(78, request);
        JsonNode repairedQuestion = repaired.path("sections").get(0)
            .path("questions").get(62);

        assertTrue(repairedQuestion.path("content").asText().contains("\u5FC3\u60F3\u4E8B\u6210"));
        assertTrue(repairedQuestion.path("analyze").asText().contains("\u68A6\u60F3\u6210\u771F"));
        assertFalse(repairedQuestion.path("content").asText().contains("\\Theta"));
        assertEquals("semantic-placeholder", XscSourceReplacementRepairs
            .repairFor(78, "sections[0].questions[62].content#math1")
            .orElseThrow().kind());
        assertFalse(XscSourceReplacementRepairs
            .repairFor(78, "sections[0].questions[62].content#math1")
            .orElseThrow().reason().isBlank());
    }

    @Test
    void rejectsCatalogDriftInsteadOfGuessing() {
        ObjectNode request = requestWithQuestion(45, "$x$:$x$:$x$:$x$");

        IllegalArgumentException failure = assertThrows(IllegalArgumentException.class,
            () -> XscSourceReplacementRepairs.apply(39, request));
        assertTrue(failure.getMessage().contains("STALE_XSC_SOURCE_REPAIR"));
    }

    @Test
    void appliesToEveryConfiguredVersionedRequest() throws Exception {
        String configured = System.getProperty("xsc.sourceRepair.requestDir", "").trim();
        Assumptions.assumeTrue(!configured.isEmpty(),
            "Enable with -Dxsc.sourceRepair.requestDir=<versioned-request-dir>");
        Path requestDir = Path.of(configured);
        Assumptions.assumeTrue(Files.isDirectory(requestDir),
            "Missing versioned request directory: " + requestDir);

        for (int documentIndex = 1; documentIndex <= 551; documentIndex++) {
            Path request = requestDir.resolve("full-%02d.request.json".formatted(documentIndex));
            assertTrue(Files.isRegularFile(request), "missing request " + request);
            XscSourceReplacementRepairs.apply(documentIndex, JSON.readTree(request.toFile()));
        }
    }

    private static ObjectNode requestWithQuestion(int questionIndex, String content) {
        ObjectNode root = JSON.createObjectNode();
        ArrayNode sections = root.putArray("sections");
        ArrayNode questions = sections.addObject().putArray("questions");
        for (int index = 0; index <= questionIndex; index++) {
            questions.addObject().put("content", index == questionIndex ? content : "plain");
        }
        return root;
    }

    private static int occurrences(String value, String needle) {
        int count = 0;
        for (int from = 0; (from = value.indexOf(needle, from)) >= 0; from += needle.length()) {
            count++;
        }
        return count;
    }
}
