package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
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
import static org.junit.jupiter.api.Assertions.assertFalse;

/** Produces the exact OLE preview payloads used by the manual visual audit. */
class OfficialLatexVisualAuditArtifactTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path OUTPUT = Path.of("target/official-latex-coverage/visual-audit");
    private String previousPreviewFormat;

    @BeforeEach
    void selectClassicWmfBackend() {
        previousPreviewFormat = System.getProperty("paperword.ole.previewFormat");
        System.setProperty("paperword.ole.previewFormat", "wmf");
    }

    @AfterEach
    void restorePreviewBackend() {
        if (previousPreviewFormat == null) {
            System.clearProperty("paperword.ole.previewFormat");
        } else {
            System.setProperty("paperword.ole.previewFormat", previousPreviewFormat);
        }
    }

    @Test
    void exportOfficialWmfPreviewsForVisualAudit() throws Exception {
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        JsonNode corpus = mapper.readTree(CORPUS.toFile());
        Set<String> examples = new LinkedHashSet<>();
        corpus.path("entries").forEach(entry -> entry.path("examples")
            .forEach(example -> examples.add(example.asText())));

        Path wmfDirectory = OUTPUT.resolve("wmf");
        Files.createDirectories(wmfDirectory);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        List<Map<String, Object>> items = new ArrayList<>();
        int index = 0;
        for (String latex : examples) {
            index++;
            LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(latex);
            assertEquals("wmf", preview.extension(), latex);
            assertFalse(preview.placeholder(), latex);

            String id = String.format("official-%03d", index);
            String fileName = id + ".wmf";
            Files.write(wmfDirectory.resolve(fileName), preview.data());
            Map<String, Object> item = new LinkedHashMap<>();
            item.put("id", id);
            item.put("latex", latex);
            item.put("wmf", "wmf/" + fileName);
            item.put("widthPt", preview.widthPt());
            item.put("heightPt", preview.heightPt());
            item.put("wmfBytes", preview.data().length);
            item.put("payload", "WMF_POLYPOLYGON_VECTOR");
            items.add(item);
        }

        Map<String, Object> manifest = new LinkedHashMap<>();
        manifest.put("schemaVersion", 1);
        manifest.put("source", CORPUS.toString().replace('\\', '/'));
        manifest.put("officialEntryCount", corpus.path("entryCount").asInt());
        manifest.put("uniqueExampleCount", examples.size());
        manifest.put("previewPayload", "WMF_POLYPOLYGON_VECTOR");
        manifest.put("items", items);
        mapper.writeValue(OUTPUT.resolve("manifest.json").toFile(), manifest);
        assertEquals(238, items.size());
    }
}
