package com.lz.paperword.tools;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Generates 10 compact Word samples from the xsc docx2tex batch requests.
 */
class XscBatch10DocxTest {

    private static final Path REQUEST_DIR = Path.of("D:/latextomathtype/analysis/batch10-requests");
    private static final Path OUTPUT_DIR = Path.of("D:/latextomathtype/analysis/batch10-docx");

    private final ObjectMapper mapper = new ObjectMapper();
    private final DocxBuilder builder = new DocxBuilder(true);

    @Test
    void generateTenXscBatchDocxFiles() throws IOException {
        Assumptions.assumeTrue(Files.exists(REQUEST_DIR), "Missing request dir: " + REQUEST_DIR);
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path runOutputDir = OUTPUT_DIR.resolve(timestamp);
        Files.createDirectories(runOutputDir);

        int generated = 0;
        for (int i = 1; i <= 10; i++) {
            Path requestPath = REQUEST_DIR.resolve("batch10-%02d.request.json".formatted(i));
            Assumptions.assumeTrue(Files.exists(requestPath), "Missing request json: " + requestPath);

            PaperExportRequest request = mapper.readValue(requestPath.toFile(), PaperExportRequest.class);
            byte[] docx = builder.build(request);
            Path output = runOutputDir.resolve("xsc测试集重建_%02d.docx".formatted(i));
            Files.write(output, docx);
            System.out.println("Generated xsc batch sample: " + output);
            assertTrue(Files.size(output) > 1000, "generated docx should not be empty: " + output);
            generated++;
        }

        assertEquals(10, generated, "should generate 10 documents");
    }
}
