package com.lz.paperword.core.latex;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.HashSet;
import java.util.HexFormat;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class OfficialLatexCoverageCorpusTest {

    private static final Path SNAPSHOT = Path.of("docs/reference/mathtype/raw/latex-coverage.html");
    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private final ObjectMapper objectMapper = new ObjectMapper();

    @Test
    void corpusMatchesPinnedOfficialSnapshot() throws Exception {
        JsonNode corpus = objectMapper.readTree(CORPUS.toFile());

        assertEquals(1, corpus.path("schemaVersion").asInt());
        assertEquals("https://www.wiris.net/demo/editor/docs/latex-coverage/",
            corpus.path("sourceUrl").asText());
        assertEquals(551, corpus.path("entryCount").asInt());
        assertEquals(551, corpus.path("entries").size());
        assertEquals(sha256(SNAPSHOT), corpus.path("sourceSha256").asText());
    }

    @Test
    void everyOfficialEntryHasStableIdentityExamplesAndCategory() throws IOException {
        JsonNode entries = objectMapper.readTree(CORPUS.toFile()).path("entries");
        Set<String> commands = new HashSet<>();
        Set<String> categories = Set.of("symbol", "modifier", "style", "structure", "environment");

        for (JsonNode entry : entries) {
            assertTrue(entry.path("index").asInt() > 0);
            assertTrue(commands.add(entry.path("command").asText()),
                () -> "duplicate command: " + entry.path("command").asText());
            assertTrue(categories.contains(entry.path("category").asText()),
                () -> "unknown category: " + entry.path("category").asText());
            assertFalse(entry.path("invocation").asText().isBlank());
            assertTrue(entry.path("examples").isArray());
            assertFalse(entry.path("examples").isEmpty());
        }
    }

    private String sha256(Path path) throws Exception {
        MessageDigest digest = MessageDigest.getInstance("SHA-256");
        return HexFormat.of().formatHex(digest.digest(Files.readAllBytes(path)));
    }
}
