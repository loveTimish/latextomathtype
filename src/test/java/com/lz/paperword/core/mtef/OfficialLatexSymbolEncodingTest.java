package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class OfficialLatexSymbolEncodingTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path REPORT = Path.of("target/official-latex-coverage/symbol-encoding-report.json");
    private final ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    @Test
    void reportsTypefaceMtcodeBits8AndVerificationSourceForEveryMappedOfficialSymbol() throws Exception {
        JsonNode entries = mapper.readTree(CORPUS.toFile()).path("entries");
        List<MtefCharMap.EncodingProfile> profiles = new ArrayList<>();
        Map<String, Integer> counts = new LinkedHashMap<>();

        for (JsonNode entry : entries) {
            if (!"symbol".equals(entry.path("category").asText())) {
                continue;
            }
            MtefCharMap.EncodingProfile profile = MtefCharMap.encodingProfile(entry.path("command").asText());
            if (profile == null) {
                continue;
            }
            profiles.add(profile);
            counts.merge(profile.source().name(), 1, Integer::sum);
            if (profile.mathTypeVerified()) {
                assertFalse(profile.source() == MtefCharMap.MappingSource.GENERATED_MATHJAX
                    || profile.source() == MtefCharMap.MappingSource.PROVISIONAL,
                    () -> "unverified mapping marked verified: " + profile.latex());
            }
        }

        Files.createDirectories(REPORT.getParent());
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("mappedOfficialSymbols", profiles.size());
        report.put("countsBySource", counts);
        report.put("profiles", profiles);
        mapper.writeValue(REPORT.toFile(), report);

        assertEquals(MtefCharMap.MappingSource.PROVISIONAL,
            MtefCharMap.encodingProfile("\\leftbarharpoon").source());
        assertFalse(MtefCharMap.encodingProfile("\\leftbarharpoon").mathTypeVerified());
    }

    @Test
    void usesMathTypeVerifiedRelationSymbolEncodingsAndDynamicEuclidDefinitions() {
        Map<String, int[]> expected = Map.ofEntries(
            Map.entry("\\bowtie", new int[] {MtefRecord.FN_MTEXTRA, 0xFFFD, 0x6E}),
            Map.entry("\\cong", new int[] {MtefRecord.FN_SYMBOL, 0x2245, 0x40}),
            Map.entry("\\sqsubset", new int[] {0x7F, 0x228F, 0xF0}),
            Map.entry("\\sqsupset", new int[] {0x7F, 0x2290, 0xF1}),
            Map.entry("\\smile", new int[] {MtefRecord.FN_MTEXTRA, 0x2323, 0x28}),
            Map.entry("\\frown", new int[] {MtefRecord.FN_MTEXTRA, 0x2322, 0x29}),
            Map.entry("\\sqsubseteq", new int[] {0x7F, 0x2291, 0xF4}),
            Map.entry("\\sqsupseteq", new int[] {0x7F, 0x2292, 0xF5}),
            Map.entry("\\ni", new int[] {MtefRecord.FN_TEXT_FE, 0x220B, -1}),
            Map.entry("\\vdash", new int[] {0x7E, 0x22A2, 0x90}),
            Map.entry("\\dashv", new int[] {0x7E, 0x22A3, 0x94})
        );
        expected.forEach((command, values) -> {
            MtefCharMap.EncodingProfile profile = MtefCharMap.encodingProfile(command);
            assertEquals(values[0], profile.typeface(), command);
            assertEquals(values[1], profile.mtcode(), command);
            assertEquals(values[2], profile.bits8(), command);
            assertTrue(profile.mathTypeVerified(), command);
        });

        String formula = "\\sqsubset,\\sqsupset,\\vdash,\\dashv";
        MtefWriter.WriteReport report = new MtefWriter().writeWithReport(new LaTeXParser().parseLaTeX(formula));
        Map<String, Integer> counts = report.normalization().recordCounts();
        assertEquals(2, counts.getOrDefault("ENCODING_DEF", 0));
        assertEquals(2, counts.getOrDefault("FONT_DEF", 0));
        assertEquals(2, counts.getOrDefault("FONT_STYLE_DEF", 0));
    }

    @Test
    void usesEncodingsVerifiedFromXscEquationNativeObjects() {
        Map<String, int[]> expected = Map.of(
            "\\cdot", new int[] {MtefRecord.FN_SYMBOL, 0x22C5, 0xD7},
            "\\cdots", new int[] {MtefRecord.FN_MTEXTRA, 0x22EF, 0x4C},
            "\\div", new int[] {MtefRecord.FN_SYMBOL, 0x00F7, 0xB8},
            "\\sim", new int[] {MtefRecord.FN_FUNCTION, 0x007E, -1},
            "\\vdots", new int[] {MtefRecord.FN_MTEXTRA, 0x22EE, 0x4D}
        );
        expected.forEach((command, values) -> {
            MtefCharMap.EncodingProfile profile = MtefCharMap.encodingProfile(command);
            assertEquals(values[0], profile.typeface(), command);
            assertEquals(values[1], profile.mtcode(), command);
            assertEquals(values[2], profile.bits8(), command);
            assertTrue(profile.mathTypeVerified(), command);
        });
    }
}
