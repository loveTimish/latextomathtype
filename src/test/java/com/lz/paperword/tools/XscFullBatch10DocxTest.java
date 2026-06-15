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
 * Generates full-document xsc Word samples from docx2tex .tex outputs.
 */
class XscFullBatch10DocxTest {

    private static final Path ANALYSIS_DIR = Path.of(
            System.getProperty("xsc.analysis.dir", "D:/latextomathtype/analysis"));
    private static final Path REQUEST_DIR = Path.of(
            System.getProperty("xsc.full.request.dir", ANALYSIS_DIR.resolve("batch10-full-requests").toString()));
    private static final Path OUTPUT_DIR = Path.of(
            System.getProperty("xsc.full.output.dir", ANALYSIS_DIR.resolve("batch10-full-docx").toString()));
    private static final Path RUN_OUTPUT_DIR = System.getProperty("xsc.full.run.output.dir") == null
            ? null
            : Path.of(System.getProperty("xsc.full.run.output.dir"));

    private final ObjectMapper mapper = new ObjectMapper();
    private final DocxBuilder builder = new DocxBuilder(true);

    @Test
    void generateTenFullXscDocxFiles() throws IOException {
        Assumptions.assumeTrue(Files.exists(REQUEST_DIR), "Missing request dir: " + REQUEST_DIR);
        String oldFallback = System.getProperty("paperword.wmf.allowTextFallback");
        System.setProperty("paperword.wmf.allowTextFallback", "false");
        try {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path runOutputDir = RUN_OUTPUT_DIR != null ? RUN_OUTPUT_DIR : OUTPUT_DIR.resolve(timestamp);
        Files.createDirectories(runOutputDir);

        int generated = 0;
        int start = Integer.getInteger("xsc.full.start", 1);
        int end = Integer.getInteger("xsc.full.end", 10);
        for (int i = start; i <= end; i++) {
            Path requestPath = REQUEST_DIR.resolve("full-%02d.request.json".formatted(i));
            Assumptions.assumeTrue(Files.exists(requestPath), "Missing request json: " + requestPath);

            PaperExportRequest request = mapper.readValue(requestPath.toFile(), PaperExportRequest.class);
            byte[] docx = builder.build(request);
            Path output = runOutputDir.resolve("xsc测试集完整重建_%02d.docx".formatted(i));
            Files.write(output, docx);
            System.out.println("Generated full xsc sample: " + output);
            assertTrue(Files.size(output) > 1000, "generated docx should not be empty: " + output);
            generated++;
        }

        assertEquals(end - start + 1, generated, "should generate requested documents");
        } finally {
            if (oldFallback == null) {
                System.clearProperty("paperword.wmf.allowTextFallback");
            } else {
                System.setProperty("paperword.wmf.allowTextFallback", oldFallback);
            }
        }
    }
}
