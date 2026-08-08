package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefWriter;
import com.lz.paperword.core.ole.OlePackager;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import com.lz.paperword.core.render.SvgVectorEmfPlusRenderer;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;

class OfficialEmfPlusOlePreviewAcceptanceTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path REPORT = Path.of(
        "target/official-latex-coverage/emfplus-ole-preview-report.json");

    @Test
    void everyOfficialExampleProducesEditableOleAndStrictEmfPlusDualPreview() throws Exception {
        String previousFormat = System.getProperty("paperword.ole.previewFormat");
        System.setProperty("paperword.ole.previewFormat", "emf");
        try {
            runAcceptance();
        } finally {
            if (previousFormat == null) {
                System.clearProperty("paperword.ole.previewFormat");
            } else {
                System.setProperty("paperword.ole.previewFormat", previousFormat);
            }
        }
    }

    private void runAcceptance() throws Exception {
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        JsonNode corpus = mapper.readTree(CORPUS.toFile());
        Set<String> examples = new LinkedHashSet<>();
        corpus.path("entries").forEach(entry -> entry.path("examples")
            .forEach(node -> examples.add(node.asText())));

        LaTeXParser parser = new LaTeXParser();
        MtefWriter writer = new MtefWriter();
        OlePackager packager = new OlePackager();
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        List<Failure> failures = new ArrayList<>();
        List<Map<String, Object>> previews = new ArrayList<>();
        int validOleCount = 0;
        int validEmfPlusCount = 0;
        for (String latex : examples) {
            try {
                LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(latex);
                if (!parsed.isSupported()) {
                    failures.add(new Failure(latex, "UNSUPPORTED", parsed.diagnostics().toString()));
                    continue;
                }
                byte[] mtef = writer.writeWithReport(parsed.mathIR()).bytes();
                validateOle(packager.packageOle(mtef), mtef.length);
                validOleCount++;

                LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(latex);
                boolean valid = !preview.placeholder()
                    && "emf".equals(preview.extension())
                    && "image/x-emf".equals(preview.contentType())
                    && preview.widthPt() > 0d
                    && preview.heightPt() > 0d
                    && SvgVectorEmfPlusRenderer.isValidDualVector(preview.data());
                if (valid) {
                    validEmfPlusCount++;
                } else {
                    failures.add(new Failure(latex, "INVALID_EMFPLUS_DUAL",
                        "extension=" + preview.extension() + ", contentType=" + preview.contentType()
                            + ", placeholder=" + preview.placeholder() + ", bytes=" + preview.data().length));
                }
                Map<String, Object> item = new LinkedHashMap<>();
                item.put("latex", latex);
                item.put("widthPt", preview.widthPt());
                item.put("heightPt", preview.heightPt());
                item.put("emfBytes", preview.data().length);
                item.put("validEmfPlusDual", valid);
                previews.add(item);
            } catch (Exception exception) {
                failures.add(new Failure(latex, "EXCEPTION", exception.toString()));
            }
        }

        Files.createDirectories(REPORT.getParent());
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("officialEntryCount", corpus.path("entryCount").asInt());
        report.put("uniqueExampleCount", examples.size());
        report.put("validOleCount", validOleCount);
        report.put("validEmfPlusDualCount", validEmfPlusCount);
        report.put("previewPayload", "EMFPLUS_DUAL_OUTLINED_VECTOR");
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        report.put("previews", previews);
        mapper.writeValue(REPORT.toFile(), report);

        assertEquals(238, examples.size());
        assertEquals(238, validOleCount);
        assertEquals(238, validEmfPlusCount);
        assertEquals(0, failures.size(), () ->
            failures.size() + " EMF+ Dual failures; inspect " + REPORT.toAbsolutePath());
    }

    private static void validateOle(byte[] ole, int mtefLength) throws Exception {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(ole))) {
            byte[] compObj = read((DocumentEntry) fs.getRoot().getEntry("\u0001CompObj"));
            byte[] nativeStream = read((DocumentEntry) fs.getRoot().getEntry("Equation Native"));
            if (!new String(compObj, StandardCharsets.ISO_8859_1).contains("Equation.DSMT4")) {
                throw new IllegalStateException("CompObj ProgID is not Equation.DSMT4");
            }
            if (nativeStream.length < 28 || Short.toUnsignedInt(ByteBuffer.wrap(nativeStream, 0, 2)
                    .order(ByteOrder.LITTLE_ENDIAN).getShort()) != 28) {
                throw new IllegalStateException("Equation Native header is invalid");
            }
            int nativeLength = ByteBuffer.wrap(nativeStream, 8, 4)
                .order(ByteOrder.LITTLE_ENDIAN).getInt();
            if (nativeLength != mtefLength || nativeStream.length != 28 + mtefLength) {
                throw new IllegalStateException("Equation Native length mismatch");
            }
        }
    }

    private static byte[] read(DocumentEntry entry) throws Exception {
        try (DocumentInputStream input = new DocumentInputStream(entry)) {
            return input.readAllBytes();
        }
    }

    private record Failure(String latex, String code, String detail) {
    }
}
