package com.lz.paperword.core.render;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;

class OfficialVectorVisualComparisonTest {

    private static final Path CORPUS = Path.of("docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path OUTPUT = Path.of("target/vector-acceptance/official-visual");
    private static final int[] DPIS = {144, 300, 600};
    @Test
    void compareBatikSvgAndIndependentWmfRasterAtMultipleDpi() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.visual"),
            "Enable with -Dpaperword.acceptance.visual=true");
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        JsonNode corpus = mapper.readTree(CORPUS.toFile());
        Set<String> examples = new LinkedHashSet<>();
        corpus.path("entries").forEach(entry -> entry.path("examples")
            .forEach(example -> examples.add(example.asText())));
        Files.createDirectories(OUTPUT.resolve("pairs"));

        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        List<VisualMetric> metrics = new ArrayList<>();
        List<VisualFailure> failures = new ArrayList<>();
        int index = 0;
        for (String latex : examples) {
            index++;
            try {
                LaTeXImageRenderer.MathJaxSvgResult source = renderer.renderMathJaxSvgForAcceptance(latex);
                LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(latex);
                for (int dpi : DPIS) {
                    WmfPreviewInspector.RasterizedWmf wmf = WmfPreviewInspector.rasterize(preview.data(), dpi);
                    BufferedImage reference = SvgVectorWmfRenderer.rasterizeBatikReference(
                        source.svgBytes(), preview.widthPt(), preview.heightPt(),
                        wmf.image().getWidth(), wmf.image().getHeight());
                    Comparison comparison = compare(reference, wmf.image());
                    metrics.add(new VisualMetric(index, latex, dpi, comparison.ssim(), comparison.iou(),
                        comparison.missingMajorComponents(), comparison.colorMae()));
                    if (comparison.ssim() < 0.90d || comparison.iou() < 0.85d
                            || comparison.missingMajorComponents() != 0 || comparison.colorMae() > 18d) {
                        failures.add(new VisualFailure(index, latex, dpi, comparison));
                    }
                    if (dpi == 300) {
                        ImageIO.write(pair(reference, wmf.image()), "png", OUTPUT.resolve("pairs")
                            .resolve("official-%03d.png".formatted(index)).toFile());
                    }
                }
            } catch (Exception exception) {
                failures.add(new VisualFailure(index, latex, 0,
                    new Comparison(0d, 0d, 1, 255d, exception.toString())));
            }
        }

        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("officialEntryCount", corpus.path("entryCount").asInt());
        report.put("uniqueExampleCount", examples.size());
        report.put("dpi", DPIS);
        report.put("ssimThreshold", 0.90d);
        report.put("foregroundIouThreshold", 0.85d);
        report.put("comparisonCount", metrics.size());
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        report.put("metrics", metrics);
        mapper.writeValue(OUTPUT.resolve("report.json").toFile(), report);

        assertEquals(238, examples.size());
        assertEquals(0, failures.size(), () ->
            failures.size() + " vector visual failures; inspect " + OUTPUT.resolve("report.json").toAbsolutePath());
    }

    static BufferedImage pair(BufferedImage left, BufferedImage right) {
        int gap = 8;
        BufferedImage pair = new BufferedImage(left.getWidth() + gap + right.getWidth(),
            Math.max(left.getHeight(), right.getHeight()), BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = pair.createGraphics();
        graphics.setColor(Color.WHITE);
        graphics.fillRect(0, 0, pair.getWidth(), pair.getHeight());
        graphics.drawImage(left, 0, 0, null);
        graphics.setColor(new Color(220, 220, 220));
        graphics.fillRect(left.getWidth(), 0, gap, pair.getHeight());
        graphics.drawImage(right, left.getWidth() + gap, 0, null);
        graphics.dispose();
        return pair;
    }

    static Comparison compare(BufferedImage reference, BufferedImage actual) {
        int width = reference.getWidth();
        int height = reference.getHeight();
        if (width != actual.getWidth() || height != actual.getHeight()) {
            return new Comparison(0d, 0d, 1, 255d, "canvas dimensions differ");
        }
        int count = width * height;
        boolean[] referenceMask = new boolean[count];
        boolean[] actualMask = new boolean[count];
        double sumReference = 0d;
        double sumActual = 0d;
        int intersection = 0;
        int union = 0;
        double colorError = 0d;
        int colorSamples = 0;
        for (int i = 0; i < count; i++) {
            int x = i % width;
            int y = i / width;
            int referenceArgb = reference.getRGB(x, y);
            int actualArgb = actual.getRGB(x, y);
            int referenceAlpha = referenceArgb >>> 24;
            int actualAlpha = actualArgb >>> 24;
            referenceMask[i] = referenceAlpha >= 32;
            actualMask[i] = actualAlpha >= 32;
            if (referenceMask[i] || actualMask[i]) {
                union++;
                if (referenceMask[i] && actualMask[i]) {
                    intersection++;
                }
            }
            double referenceLuma = compositeLuma(referenceArgb);
            double actualLuma = compositeLuma(actualArgb);
            sumReference += referenceLuma;
            sumActual += actualLuma;
            if (referenceAlpha >= 200 && actualAlpha >= 200) {
                colorError += colorDistance(referenceArgb, actualArgb);
                colorSamples++;
            }
        }
        double meanReference = sumReference / count;
        double meanActual = sumActual / count;
        double varianceReference = 0d;
        double varianceActual = 0d;
        double covariance = 0d;
        for (int i = 0; i < count; i++) {
            int x = i % width;
            int y = i / width;
            double referenceValue = compositeLuma(reference.getRGB(x, y)) - meanReference;
            double actualValue = compositeLuma(actual.getRGB(x, y)) - meanActual;
            varianceReference += referenceValue * referenceValue;
            varianceActual += actualValue * actualValue;
            covariance += referenceValue * actualValue;
        }
        int denominator = Math.max(count - 1, 1);
        varianceReference /= denominator;
        varianceActual /= denominator;
        covariance /= denominator;
        double c1 = Math.pow(0.01d * 255d, 2d);
        double c2 = Math.pow(0.03d * 255d, 2d);
        double ssim = ((2d * meanReference * meanActual + c1) * (2d * covariance + c2))
            / ((meanReference * meanReference + meanActual * meanActual + c1)
            * (varianceReference + varianceActual + c2));
        double iou = union == 0 ? 1d : intersection / (double) union;
        int missing = missingMajorComponents(referenceMask, actualMask, width, height);
        double colorMae = colorSamples == 0 ? 0d : colorError / colorSamples;
        return new Comparison(ssim, iou, missing, colorMae, "");
    }

    private static int missingMajorComponents(boolean[] reference, boolean[] actual, int width, int height) {
        boolean[] visited = new boolean[reference.length];
        int referenceInk = 0;
        for (boolean value : reference) {
            if (value) {
                referenceInk++;
            }
        }
        int majorThreshold = Math.max(4, (int) Math.ceil(referenceInk * 0.001d));
        int missing = 0;
        int[] dx = {-1, 0, 1, -1, 1, -1, 0, 1};
        int[] dy = {-1, -1, -1, 0, 0, 1, 1, 1};
        for (int start = 0; start < reference.length; start++) {
            if (!reference[start] || visited[start]) {
                continue;
            }
            ArrayDeque<Integer> queue = new ArrayDeque<>();
            queue.add(start);
            visited[start] = true;
            int area = 0;
            int overlap = 0;
            while (!queue.isEmpty()) {
                int current = queue.removeFirst();
                area++;
                if (actual[current]) {
                    overlap++;
                }
                int x = current % width;
                int y = current / width;
                for (int direction = 0; direction < dx.length; direction++) {
                    int nx = x + dx[direction];
                    int ny = y + dy[direction];
                    if (nx < 0 || nx >= width || ny < 0 || ny >= height) {
                        continue;
                    }
                    int next = ny * width + nx;
                    if (reference[next] && !visited[next]) {
                        visited[next] = true;
                        queue.add(next);
                    }
                }
            }
            if (area >= majorThreshold && overlap < area * 0.50d) {
                missing++;
            }
        }
        return missing;
    }

    private static double compositeLuma(int argb) {
        double alpha = (argb >>> 24) / 255d;
        double red = ((argb >>> 16) & 0xFF) * alpha + 255d * (1d - alpha);
        double green = ((argb >>> 8) & 0xFF) * alpha + 255d * (1d - alpha);
        double blue = (argb & 0xFF) * alpha + 255d * (1d - alpha);
        return 0.2126d * red + 0.7152d * green + 0.0722d * blue;
    }

    private static double colorDistance(int left, int right) {
        return (Math.abs(((left >>> 16) & 0xFF) - ((right >>> 16) & 0xFF))
            + Math.abs(((left >>> 8) & 0xFF) - ((right >>> 8) & 0xFF))
            + Math.abs((left & 0xFF) - (right & 0xFF))) / 3d;
    }

    private record VisualMetric(int index, String latex, int dpi, double ssim,
                                double foregroundIou, int missingMajorComponents, double colorMae) {}

    private record VisualFailure(int index, String latex, int dpi, Comparison comparison) {}

    record Comparison(double ssim, double iou, int missingMajorComponents,
                      double colorMae, String error) {}
}
