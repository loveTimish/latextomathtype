package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefWriter;
import com.lz.paperword.core.ole.OlePackager;
import com.lz.paperword.core.render.LaTeXImageRenderer;
import com.lz.paperword.core.render.WmfPreviewInspector;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
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

class OfficialLatexOlePreviewAcceptanceTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path REPORT = Path.of("target/official-latex-coverage/ole-preview-report.json");

    private final ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
    private final LaTeXParser parser = new LaTeXParser();
    private final MtefWriter writer = new MtefWriter();
    private final OlePackager packager = new OlePackager();
    private final LaTeXImageRenderer renderer = new LaTeXImageRenderer();
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
    void everyOfficialExampleProducesEditableOleAndReadableUnclippedWmf() throws Exception {
        JsonNode corpus = mapper.readTree(CORPUS.toFile());
        Set<String> examples = new LinkedHashSet<>();
        corpus.path("entries").forEach(entry -> entry.path("examples").forEach(node -> examples.add(node.asText())));

        List<Failure> failures = new ArrayList<>();
        List<Map<String, Object>> previews = new ArrayList<>();
        int validOleCount = 0;
        int readableWmfCount = 0;
        int unclippedWmfCount = 0;
        int nonOverlappedWmfCount = 0;
        int bitmapWmfCount = 0;
        int vectorWmfCount = 0;
        for (String latex : examples) {
            try {
                LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(latex);
                if (!parsed.isSupported()) {
                    failures.add(new Failure(latex, "UNSUPPORTED", parsed.diagnostics().toString()));
                    continue;
                }
                MtefWriter.WriteReport written = writer.writeWithReport(parsed.mathIR());
                byte[] mtef = written.bytes();
                boolean hasVisibleCharacter = written.normalization().recordCounts().getOrDefault("CHAR", 0) > 0;
                validateOle(packager.packageOle(mtef), mtef.length);
                validOleCount++;

                LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(latex);
                WmfPreviewInspector.Inspection inspection = WmfPreviewInspector.inspect(preview.data());
                if (inspection.bitmapRecordCount() > 0) {
                    bitmapWmfCount++;
                } else {
                    vectorWmfCount++;
                }
                Map<String, Object> item = new LinkedHashMap<>();
                item.put("latex", latex);
                item.put("widthPt", preview.widthPt());
                item.put("heightPt", preview.heightPt());
                item.put("wmfBytes", preview.data().length);
                item.put("inspection", inspection);
                previews.add(item);

                if (preview.placeholder() || !"wmf".equals(preview.extension())) {
                    failures.add(new Failure(latex, "WMF_FALLBACK", "placeholder=" + preview.placeholder()
                        + ", extension=" + preview.extension()));
                } else if (!inspection.valid()) {
                    failures.add(new Failure(latex, "INVALID_WMF", inspection.error()));
                } else if (hasVisibleCharacter && inspection.foregroundPixels() == 0) {
                    failures.add(new Failure(latex, "BLANK_WMF", inspection.toString()));
                } else {
                    readableWmfCount++;
                    if (inspection.inkTouchesEdge()) {
                        failures.add(new Failure(latex, "CLIPPED_WMF", inspection.toString()));
                    } else {
                        unclippedWmfCount++;
                        if (inspection.foregroundDensity() > 0.80d) {
                            failures.add(new Failure(latex, "OVERLAPPED_OR_SOLID_WMF", inspection.toString()));
                        } else {
                            nonOverlappedWmfCount++;
                        }
                    }
                }
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
        report.put("readableWmfCount", readableWmfCount);
        report.put("unclippedWmfCount", unclippedWmfCount);
        report.put("nonOverlappedWmfCount", nonOverlappedWmfCount);
        report.put("bitmapWmfCount", bitmapWmfCount);
        report.put("vectorWmfCount", vectorWmfCount);
        report.put("previewPayload", bitmapWmfCount == 0 ? "WMF_POLYPOLYGON_VECTOR" : "MIXED_OR_BITMAP_WMF");
        report.put("passedCount", examples.size() - failures.size());
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        report.put("previews", previews);
        mapper.writeValue(REPORT.toFile(), report);

        assertEquals(238, examples.size());
        assertEquals(0, bitmapWmfCount, "Strict vector branch must not emit DIB-backed WMF previews");
        assertEquals(238, vectorWmfCount, "Every official example must emit a pure vector WMF");
        assertEquals(0, failures.size(), () -> failures.size() + " OLE/WMF failures; inspect " + REPORT.toAbsolutePath());
    }

    private void validateOle(byte[] ole, int mtefLength) throws Exception {
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
            int nativeLength = ByteBuffer.wrap(nativeStream, 8, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
            if (nativeLength != mtefLength || nativeStream.length != 28 + mtefLength) {
                throw new IllegalStateException("Equation Native length mismatch");
            }
        }
    }

    private byte[] read(DocumentEntry entry) throws Exception {
        try (DocumentInputStream input = new DocumentInputStream(entry)) {
            return input.readAllBytes();
        }
    }

    private record Failure(String latex, String code, String detail) {
    }
}
