package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class FormulaGoldenCorpusTest {

    @Test
    void requiredCorpusRendersAsSelfWrittenVectorWmf() throws IOException {
        List<Case> cases = loadCases();
        List<Result> results = new ArrayList<>();
        List<String> failures = new ArrayList<>();
        Set<String> requiredFamilies = new LinkedHashSet<>();
        Set<String> renderedFamilies = new LinkedHashSet<>();

        for (Case formula : cases) {
            if (formula.required()) {
                requiredFamilies.add(formula.family());
            }
            boolean canRender = VectorWmfFormulaRenderer.canRender(formula.latex());
            if (!canRender) {
                results.add(Result.gap(formula));
                if (formula.required()) {
                    failures.add(formula.id() + " should be supported by the vector renderer");
                }
                continue;
            }

            byte[] wmf = VectorWmfFormulaRenderer.render(formula.latex(), formula.widthPt(), formula.heightPt());
            assertNotNull(wmf, formula.id() + " should produce WMF bytes");
            WmfStats stats = inspectWmf(wmf);
            results.add(Result.rendered(formula, stats));
            renderedFamilies.add(formula.family());

            addFailureIf(failures, stats.extTextOutRecords() < formula.minTextRecords(),
                formula.id() + " should draw enough text runs, got " + stats.extTextOutRecords());
            addFailureIf(failures, stats.polylineRecords() < formula.minPolylineRecords(),
                formula.id() + " should draw required structure lines, got " + stats.polylineRecords());
            int expectedRadicals = countLatexCommand(formula.latex(), "\\sqrt");
            if (expectedRadicals > 0) {
                addFailureIf(failures, stats.fourPointPolylineRecords() < expectedRadicals,
                    formula.id() + " should preserve each radical as a polyline with at least four points, got "
                        + stats.fourPointPolylineRecords());
            }
            addFailureIf(failures, stats.stretchDibRecords() != 0, formula.id() + " must not use StretchDIB");
            addFailureIf(failures, stats.bitmapRecords() != 0, formula.id() + " must not use bitmap fallback records");
            addFailureIf(failures, stats.windowExtX() != (int) Math.round(formula.widthPt() * 20.0d),
                formula.id() + " should preserve requested physical width");
            addFailureIf(failures, stats.windowExtY() != (int) Math.round(formula.heightPt() * 20.0d),
                formula.id() + " should preserve requested physical height");
            addFailureIf(failures, stats.maxTextRightTwips() > stats.windowExtX() + 16,
                formula.id() + " text should stay inside the WMF box");
            addFailureIf(failures, stats.minTextXTwips() < -16, formula.id() + " text should not start off-canvas");
            addFailureIf(failures, stats.minPolylineXTwips() < -16,
                formula.id() + " structure lines should not start off-canvas");
            addFailureIf(failures, stats.maxPolylineXTwips() > stats.windowExtX() + 16,
                formula.id() + " structure lines should stay inside the WMF box");
            addFailureIf(failures, stats.maxTextYTwips() > stats.windowExtY() + 16,
                formula.id() + " text baseline should stay inside the WMF box");
            addFailureIf(failures, stats.maxPolylineYTwips() > stats.windowExtY() + 16,
                formula.id() + " structure lines should stay inside the WMF box");
        }

        writeReport(results);
        assertTrue(renderedFamilies.containsAll(requiredFamilies), "all required structure families must render");
        assertTrue(requiredFamilies.contains("sqrt_nested"), "nested radicals must stay explicit in the corpus");
        assertTrue(requiredFamilies.contains("sqrt_nested_fraction"),
            "nested radicals with fractions must stay explicit in the corpus");
        assertTrue(failures.isEmpty(), String.join(System.lineSeparator(), failures));
    }

    private static void addFailureIf(List<String> failures, boolean condition, String message) {
        if (condition) {
            failures.add(message);
        }
    }

    private static List<Case> loadCases() throws IOException {
        try (InputStream stream = FormulaGoldenCorpusTest.class.getResourceAsStream("/formula-golden-corpus.tsv")) {
            assertNotNull(stream, "formula-golden-corpus.tsv should be on the test classpath");
            String text = new String(stream.readAllBytes(), StandardCharsets.UTF_8);
            List<Case> cases = new ArrayList<>();
            String[] lines = text.split("\\R");
            for (int i = 1; i < lines.length; i++) {
                if (lines[i].isBlank()) {
                    continue;
                }
                String[] parts = lines[i].split("\\t", 9);
                assertEquals(9, parts.length, "invalid corpus row " + (i + 1));
                cases.add(new Case(
                    parts[0],
                    parts[1],
                    parts[2],
                    parts[3],
                    Double.parseDouble(parts[4]),
                    Double.parseDouble(parts[5]),
                    Integer.parseInt(parts[6]),
                    Integer.parseInt(parts[7]),
                    parts[8]
                ));
            }
            return cases;
        }
    }

    private static void writeReport(List<Result> results) throws IOException {
        Path outputDir = Path.of("analysis", "formula-golden-corpus");
        Files.createDirectories(outputDir);
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path output = outputDir.resolve("formula-golden-corpus-report-" + timestamp + ".json");
        Files.writeString(output, toJson(results), StandardCharsets.UTF_8);
        try {
            Files.copy(output, outputDir.resolve("formula-golden-corpus-report-latest.json"),
                StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
            // Timestamped report is the authoritative artifact if the alias is locked.
        }
        System.out.println("Formula golden corpus report: " + output.toAbsolutePath());
    }

    private static String toJson(List<Result> results) {
        StringBuilder out = new StringBuilder();
        out.append("{\n  \"cases\": [\n");
        for (int i = 0; i < results.size(); i++) {
            Result result = results.get(i);
            Case formula = result.formula();
            out.append("    {\n");
            out.append("      \"id\": \"").append(escape(formula.id())).append("\",\n");
            out.append("      \"subject\": \"").append(escape(formula.subject())).append("\",\n");
            out.append("      \"family\": \"").append(escape(formula.family())).append("\",\n");
            out.append("      \"expectation\": \"").append(escape(formula.expectation())).append("\",\n");
            out.append("      \"status\": \"").append(result.stats() == null ? "gap" : "rendered").append("\",\n");
            out.append("      \"latex\": \"").append(escape(formula.latex())).append("\"");
            if (result.stats() != null) {
                WmfStats stats = result.stats();
                out.append(",\n      \"wmf\": {\n");
                out.append("        \"extTextOutRecords\": ").append(stats.extTextOutRecords()).append(",\n");
                out.append("        \"polylineRecords\": ").append(stats.polylineRecords()).append(",\n");
                out.append("        \"stretchDibRecords\": ").append(stats.stretchDibRecords()).append(",\n");
                out.append("        \"bitmapRecords\": ").append(stats.bitmapRecords()).append(",\n");
                out.append("        \"windowExtX\": ").append(stats.windowExtX()).append(",\n");
                out.append("        \"windowExtY\": ").append(stats.windowExtY()).append(",\n");
                out.append("        \"minTextXTwips\": ").append(stats.minTextXTwips()).append(",\n");
                out.append("        \"maxTextRightTwips\": ").append(stats.maxTextRightTwips()).append(",\n");
                out.append("        \"maxTextYTwips\": ").append(stats.maxTextYTwips()).append(",\n");
                out.append("        \"minPolylineXTwips\": ").append(stats.minPolylineXTwips()).append(",\n");
                out.append("        \"maxPolylineXTwips\": ").append(stats.maxPolylineXTwips()).append(",\n");
                out.append("        \"maxPolylineYTwips\": ").append(stats.maxPolylineYTwips()).append(",\n");
                out.append("        \"maxPolylinePointCount\": ").append(stats.maxPolylinePointCount()).append(",\n");
                out.append("        \"fourPointPolylineRecords\": ").append(stats.fourPointPolylineRecords())
                    .append("\n");
                out.append("      }");
            }
            out.append("\n    }");
            if (i + 1 < results.size()) {
                out.append(",");
            }
            out.append("\n");
        }
        out.append("  ]\n}\n");
        return out.toString();
    }

    private static String escape(String value) {
        return value.replace("\\", "\\\\").replace("\"", "\\\"");
    }

    private static int countLatexCommand(String latex, String command) {
        int count = 0;
        int index = 0;
        while ((index = latex.indexOf(command, index)) >= 0) {
            int end = index + command.length();
            if (end >= latex.length() || !Character.isLetter(latex.charAt(end))) {
                count++;
            }
            index = end;
        }
        return count;
    }

    private static WmfStats inspectWmf(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        int extTextOut = 0;
        int polyline = 0;
        int stretchDib = 0;
        int bitmap = 0;
        int windowExtX = -1;
        int windowExtY = -1;
        int minTextX = Integer.MAX_VALUE;
        int maxTextRight = 0;
        int maxTextY = 0;
        int minPolylineX = Integer.MAX_VALUE;
        int maxPolylineX = 0;
        int maxPolylineY = 0;
        int maxPolylinePointCount = 0;
        int fourPointPolylineRecords = 0;
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            int recordEnd = Math.min(data.length, offset + sizeWords * 2);
            if (function == 0x020C && offset + 10 <= recordEnd) {
                windowExtY = word(data, offset + 6);
                windowExtX = word(data, offset + 8);
            } else if (function == 0x0A32 && offset + 14 <= recordEnd) {
                extTextOut++;
                int y = word(data, offset + 6);
                int x = word(data, offset + 8);
                int count = word(data, offset + 10);
                int options = word(data, offset + 12);
                int rectangleBytes = (options & 0x0006) == 0 ? 0 : 8;
                int textOffset = offset + 14 + rectangleBytes;
                int textBytes = count + (count & 1);
                int dxOffset = textOffset + textBytes;
                int dxTotal = 0;
                while (dxOffset + 2 <= recordEnd) {
                    dxTotal += word(data, dxOffset);
                    dxOffset += 2;
                }
                minTextX = Math.min(minTextX, x);
                maxTextRight = Math.max(maxTextRight, x + dxTotal);
                maxTextY = Math.max(maxTextY, y);
            } else if (function == 0x0325 && offset + 8 <= recordEnd) {
                polyline++;
                int pointCount = word(data, offset + 6);
                maxPolylinePointCount = Math.max(maxPolylinePointCount, pointCount);
                if (pointCount >= 4) {
                    fourPointPolylineRecords++;
                }
                for (int i = 0; i < pointCount && offset + 10 + i * 4 <= recordEnd; i++) {
                    int x = word(data, offset + 8 + i * 4);
                    minPolylineX = Math.min(minPolylineX, x);
                    maxPolylineX = Math.max(maxPolylineX, x);
                    maxPolylineY = Math.max(maxPolylineY, word(data, offset + 10 + i * 4));
                }
            } else if (function == 0x0F43) {
                stretchDib++;
            } else if (isBitmapFunction(function)) {
                bitmap++;
            }
            offset += sizeWords * 2;
        }
        return new WmfStats(extTextOut, polyline, stretchDib, bitmap, windowExtX, windowExtY,
            minTextX == Integer.MAX_VALUE ? 0 : minTextX, maxTextRight, maxTextY,
            minPolylineX == Integer.MAX_VALUE ? 0 : minPolylineX, maxPolylineX, maxPolylineY,
            maxPolylinePointCount, fourPointPolylineRecords);
    }

    private static boolean isBitmapFunction(int function) {
        return function == 0x0922 || function == 0x0940 || function == 0x0B23 || function == 0x0B41
            || function == 0x0D33 || function == 0x0F43;
    }

    private static boolean hasPlaceableHeader(byte[] data) {
        return data.length >= 22 && dword(data, 0) == 0x9AC6CDD7;
    }

    private static int word(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort() & 0xffff;
    }

    private static int dword(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
    }

    private record Case(
        String id,
        String subject,
        String family,
        String expectation,
        double widthPt,
        double heightPt,
        int minTextRecords,
        int minPolylineRecords,
        String latex
    ) {
        boolean required() {
            return "required".equals(expectation);
        }
    }

    private record Result(Case formula, WmfStats stats) {
        static Result rendered(Case formula, WmfStats stats) {
            return new Result(formula, stats);
        }

        static Result gap(Case formula) {
            return new Result(formula, null);
        }
    }

    private record WmfStats(
        int extTextOutRecords,
        int polylineRecords,
        int stretchDibRecords,
        int bitmapRecords,
        int windowExtX,
        int windowExtY,
        int minTextXTwips,
        int maxTextRightTwips,
        int maxTextYTwips,
        int minPolylineXTwips,
        int maxPolylineXTwips,
        int maxPolylineYTwips,
        int maxPolylinePointCount,
        int fourPointPolylineRecords
    ) {
    }
}
