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

/**
 * Generates the xsc-size calibration sample produced by D:\latextomathtype\analysis.
 */
class CalibrationSampleDocxTest {

    private static final Path SOURCE_JSON = Path.of("D:/latextomathtype/analysis/sample.request.json");
    private static final Path OUTPUT_DOCX = Path.of("D:/latextomathtype/analysis/sample.generated.filtered.docx");

    private final ObjectMapper mapper = new ObjectMapper();
    private final DocxBuilder builder = new DocxBuilder(true);

    @Test
    void generateCalibrationSampleDocx() throws IOException {
        Assumptions.assumeTrue(Files.exists(SOURCE_JSON), "Missing calibration request json: " + SOURCE_JSON);

        PaperExportRequest request = mapper.readValue(SOURCE_JSON.toFile(), PaperExportRequest.class);
        byte[] docx = builder.build(request);

        Files.createDirectories(OUTPUT_DOCX.getParent());
        Files.write(OUTPUT_DOCX, docx);

        System.out.println("Generated calibration sample: " + OUTPUT_DOCX);
        assertTrue(Files.size(OUTPUT_DOCX) > 1000, "generated docx should not be empty");
    }
}
