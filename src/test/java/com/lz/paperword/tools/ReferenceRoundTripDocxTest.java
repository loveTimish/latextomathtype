package com.lz.paperword.tools;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertTrue;

class ReferenceRoundTripDocxTest {

    private static final Path SOURCE_JSON = Path.of("target/reference-roundtrip/fraction-split-reference.request.json");
    private static final Path OUTPUT_DOCX = Path.of("target/reference-roundtrip/fraction-split-reference-regenerated.docx");

    private final ObjectMapper mapper = new ObjectMapper();
    private final DocxBuilder builder = new DocxBuilder(true);

    @Test
    void generateFractionReferenceRoundTripDocx() throws IOException {
        Assumptions.assumeTrue(Files.exists(SOURCE_JSON), "Missing reference roundtrip json: " + SOURCE_JSON);

        PaperExportRequest request = mapper.readValue(SOURCE_JSON.toFile(), PaperExportRequest.class);
        byte[] docx = builder.build(request);

        Files.createDirectories(OUTPUT_DOCX.getParent());
        Files.write(OUTPUT_DOCX, docx);

        assertTrue(Files.size(OUTPUT_DOCX) > 1000, "generated docx should not be empty");
    }
}
