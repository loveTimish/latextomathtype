package com.lz.paperword.core.render;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** Compares the serialized EMF+ paths, not renderer-internal geometry, with the SVG source. */
class OfficialEmfPlusVisualComparisonTest {

    private static final Path CORPUS = Path.of(
        "docs/reference/mathtype/latex-coverage-corpus.json");
    private static final Path OUTPUT = Path.of(
        "target/vector-acceptance/official-emfplus-visual");
    private static final int[] DPIS = {144, 300, 600};
    private static final double SSIM_THRESHOLD = 0.90d;
    private static final double IOU_THRESHOLD = 0.85d;
    private static final double MAX_SCALE_ERROR = 0.03d;
    private static final double MAX_CENTER_SHIFT_PX = 1.5d;
    private static final double MAX_ADDED_OUTSIDE_DILATION = 0.02d;
    private String previousPreviewFormat;

    @BeforeEach
    void selectEmfPlusBackend() {
        previousPreviewFormat = System.getProperty("paperword.ole.previewFormat");
        System.setProperty("paperword.ole.previewFormat", "emf");
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
    void compareSerializedEmfPlusAndMathJaxAtMultipleDpi() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.emfPlusVisual"),
            "Enable with -Dpaperword.acceptance.emfPlusVisual=true");
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        JsonNode corpus = mapper.readTree(CORPUS.toFile());
        Set<String> examples = new LinkedHashSet<>();
        corpus.path("entries").forEach(entry -> entry.path("examples")
            .forEach(example -> examples.add(example.asText())));
        Files.createDirectories(OUTPUT.resolve("pairs-300dpi"));

        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        List<VisualMetric> metrics = new ArrayList<>();
        List<VisualFailure> failures = new ArrayList<>();
        int index = 0;
        for (String latex : examples) {
            index++;
            try {
                LaTeXImageRenderer.MathJaxSvgResult source =
                    renderer.renderMathJaxSvgForAcceptance(latex);
                LaTeXImageRenderer.PreviewImage preview = renderer.renderForOlePreview(latex);
                BatikVectorSceneBuilder.VectorScene scene =
                    BatikVectorSceneBuilder.build(source.svgBytes());
                for (int dpi : DPIS) {
                    BufferedImage reference = rasterizeReference(scene,
                        preview.widthPt(), preview.heightPt(), dpi);
                    EmfPlusPreviewInspector.RasterizedEmfPlus decoded =
                        EmfPlusPreviewInspector.rasterize(preview.data(),
                            preview.widthPt(), preview.heightPt(), dpi);
                    BufferedImage actual = decoded.image();
                    OfficialVectorVisualComparisonTest.Comparison comparison =
                        OfficialVectorVisualComparisonTest.compare(reference, actual);
                    InkDelta ink = compareInkGeometry(reference, actual);
                    Set<Integer> expectedColors = new LinkedHashSet<>();
                    scene.shapes().forEach(shape -> expectedColors.add(shape.color().getRGB()));
                    boolean structureMismatch = decoded.pathFillCount() != scene.shapes().size()
                        || !decoded.colors().equals(expectedColors);
                    VisualMetric metric = new VisualMetric(index, latex, dpi,
                        comparison.ssim(), comparison.iou(),
                        comparison.missingMajorComponents(), comparison.colorMae(),
                        ink.widthScale(), ink.heightScale(), ink.centerShiftPx(),
                        ink.widthDifferencePx(), ink.heightDifferencePx(),
                        ink.addedOutsideDilatedReference(), scene.shapes().size(),
                        decoded.pathFillCount(), structureMismatch);
                    metrics.add(metric);
                    if (!passes(metric)) {
                        failures.add(new VisualFailure(index, latex, dpi, metric));
                    }
                    if (dpi == 300) {
                        ImageIO.write(OfficialVectorVisualComparisonTest.pair(reference, actual),
                            "png", OUTPUT.resolve("pairs-300dpi")
                                .resolve("official-%03d.png".formatted(index)).toFile());
                    }
                }
            } catch (Exception exception) {
                VisualMetric metric = VisualMetric.exception(index, latex,
                    exception.toString());
                failures.add(new VisualFailure(index, latex, 0, metric));
                metrics.add(metric);
            }
        }

        List<VisualMetric> byIou = metrics.stream()
            .filter(metric -> metric.dpi() > 0)
            .sorted(Comparator.comparingDouble(VisualMetric::foregroundIou))
            .limit(20)
            .toList();
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("backend", "MathJax SVG -> Batik scene -> serialized EMF+ Path/FillPath");
        report.put("officialEntryCount", corpus.path("entryCount").asInt());
        report.put("uniqueExampleCount", examples.size());
        report.put("dpi", DPIS);
        report.put("thresholds", Map.of(
            "ssim", SSIM_THRESHOLD,
            "foregroundIou", IOU_THRESHOLD,
            "maxScaleError", MAX_SCALE_ERROR,
            "maxRasterRoundingDifferencePx", 1,
            "maxCenterShiftPx", MAX_CENTER_SHIFT_PX,
            "maxAddedOutsideOnePixelDilation", MAX_ADDED_OUTSIDE_DILATION));
        report.put("comparisonCount", metrics.stream().filter(metric -> metric.dpi() > 0).count());
        report.put("failureCount", failures.size());
        report.put("worstTwentyByIou", byIou);
        report.put("failures", failures);
        report.put("metrics", metrics);
        mapper.writeValue(OUTPUT.resolve("report.json").toFile(), report);

        assertEquals(238, examples.size());
        assertEquals(238 * DPIS.length,
            metrics.stream().filter(metric -> metric.dpi() > 0).count());
        assertEquals(0, failures.size(), () -> failures.size()
            + " EMF+ visual failures; inspect " + OUTPUT.resolve("report.json").toAbsolutePath());
    }

    private static boolean passes(VisualMetric metric) {
        return metric.error().isEmpty()
            && metric.ssim() >= SSIM_THRESHOLD
            && metric.foregroundIou() >= IOU_THRESHOLD
            && metric.missingMajorComponents() == 0
            && metric.colorMae() <= 18d
            && (Math.abs(metric.widthScale() - 1d) <= MAX_SCALE_ERROR
                || metric.widthDifferencePx() <= 1)
            && (Math.abs(metric.heightScale() - 1d) <= MAX_SCALE_ERROR
                || metric.heightDifferencePx() <= 1)
            && metric.centerShiftPx() <= MAX_CENTER_SHIFT_PX
            && metric.addedOutsideDilatedReference() <= MAX_ADDED_OUTSIDE_DILATION
            && !metric.structureMismatch();
    }

    private static BufferedImage rasterizeReference(BatikVectorSceneBuilder.VectorScene scene,
                                                     double widthPt, double heightPt, int dpi) {
        int widthPx = Math.max(1, (int) Math.ceil(widthPt * dpi / 72d));
        int heightPx = Math.max(1, (int) Math.ceil(heightPt * dpi / 72d));
        double logicalWidth = widthPt * 96d / 72d;
        double logicalHeight = heightPt * 96d / 72d;
        double scale = Math.min(logicalWidth / scene.viewportWidth(),
            logicalHeight / scene.viewportHeight());
        double offsetX = (logicalWidth - scene.viewportWidth() * scale) / 2d;
        double offsetY = (logicalHeight - scene.viewportHeight() * scale) / 2d;
        AffineTransform sourceToLogical = new AffineTransform(
            scale, 0d, 0d, scale, offsetX, offsetY);
        AffineTransform sourceToPixels = AffineTransform.getScaleInstance(
            dpi / 96d, dpi / 96d);
        sourceToPixels.concatenate(sourceToLogical);

        BufferedImage image = new BufferedImage(widthPx, heightPx,
            BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = image.createGraphics();
        try {
            graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                RenderingHints.VALUE_ANTIALIAS_ON);
            graphics.setRenderingHint(RenderingHints.KEY_RENDERING,
                RenderingHints.VALUE_RENDER_QUALITY);
            graphics.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL,
                RenderingHints.VALUE_STROKE_PURE);
            for (BatikVectorSceneBuilder.PaintedShape painted : scene.shapes()) {
                graphics.setColor(painted.color());
                graphics.fill(sourceToPixels.createTransformedShape(painted.shape()));
            }
        } finally {
            graphics.dispose();
        }
        return image;
    }

    private static InkDelta compareInkGeometry(BufferedImage reference, BufferedImage actual) {
        InkGeometry expected = InkGeometry.of(reference);
        InkGeometry observed = InkGeometry.of(actual);
        if (expected.inkPixels() == 0 && observed.inkPixels() == 0) {
            return new InkDelta(1d, 1d, 0d, 0, 0, 0d);
        }
        if (expected.inkPixels() == 0 || observed.inkPixels() == 0) {
            return new InkDelta(0d, 0d, Double.POSITIVE_INFINITY,
                Integer.MAX_VALUE, Integer.MAX_VALUE, 1d);
        }
        double widthScale = observed.width() / (double) expected.width();
        double heightScale = observed.height() / (double) expected.height();
        double centerShift = Math.hypot(observed.centerX() - expected.centerX(),
            observed.centerY() - expected.centerY());
        int outside = 0;
        for (int y = 0; y < actual.getHeight(); y++) {
            for (int x = 0; x < actual.getWidth(); x++) {
                if (isInk(actual.getRGB(x, y)) && !hasInkNear(reference, x, y, 1)) {
                    outside++;
                }
            }
        }
        return new InkDelta(widthScale, heightScale, centerShift,
            Math.abs(observed.width() - expected.width()),
            Math.abs(observed.height() - expected.height()),
            outside / (double) observed.inkPixels());
    }

    private static boolean hasInkNear(BufferedImage image, int x, int y, int radius) {
        for (int dy = -radius; dy <= radius; dy++) {
            int sampleY = y + dy;
            if (sampleY < 0 || sampleY >= image.getHeight()) {
                continue;
            }
            for (int dx = -radius; dx <= radius; dx++) {
                int sampleX = x + dx;
                if (sampleX >= 0 && sampleX < image.getWidth()
                        && isInk(image.getRGB(sampleX, sampleY))) {
                    return true;
                }
            }
        }
        return false;
    }

    private static boolean isInk(int argb) {
        return (argb >>> 24) >= 32;
    }

    private record InkGeometry(int inkPixels, int minX, int minY, int maxX, int maxY,
                               double centerX, double centerY) {
        static InkGeometry of(BufferedImage image) {
            int minX = Integer.MAX_VALUE;
            int minY = Integer.MAX_VALUE;
            int maxX = Integer.MIN_VALUE;
            int maxY = Integer.MIN_VALUE;
            long sumX = 0L;
            long sumY = 0L;
            int ink = 0;
            for (int y = 0; y < image.getHeight(); y++) {
                for (int x = 0; x < image.getWidth(); x++) {
                    if (!isInk(image.getRGB(x, y))) {
                        continue;
                    }
                    ink++;
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                    sumX += x;
                    sumY += y;
                }
            }
            if (ink == 0) {
                return new InkGeometry(0, 0, 0, -1, -1, 0d, 0d);
            }
            return new InkGeometry(ink, minX, minY, maxX, maxY,
                sumX / (double) ink, sumY / (double) ink);
        }

        int width() {
            return inkPixels == 0 ? 0 : maxX - minX + 1;
        }

        int height() {
            return inkPixels == 0 ? 0 : maxY - minY + 1;
        }
    }

    private record InkDelta(double widthScale, double heightScale, double centerShiftPx,
                            int widthDifferencePx, int heightDifferencePx,
                            double addedOutsideDilatedReference) {
    }

    private record VisualMetric(int index, String latex, int dpi, double ssim,
                                double foregroundIou, int missingMajorComponents,
                                double colorMae, double widthScale, double heightScale,
                                double centerShiftPx, int widthDifferencePx,
                                int heightDifferencePx, double addedOutsideDilatedReference,
                                int sourceShapeCount, int serializedPathFillCount,
                                boolean structureMismatch, String error) {
        VisualMetric(int index, String latex, int dpi, double ssim,
                     double foregroundIou, int missingMajorComponents,
                     double colorMae, double widthScale, double heightScale,
                     double centerShiftPx, int widthDifferencePx,
                     int heightDifferencePx, double addedOutsideDilatedReference,
                     int sourceShapeCount, int serializedPathFillCount,
                     boolean structureMismatch) {
            this(index, latex, dpi, ssim, foregroundIou, missingMajorComponents,
                colorMae, widthScale, heightScale, centerShiftPx,
                widthDifferencePx, heightDifferencePx,
                addedOutsideDilatedReference, sourceShapeCount,
                serializedPathFillCount, structureMismatch, "");
        }

        static VisualMetric exception(int index, String latex, String error) {
            return new VisualMetric(index, latex, 0, 0d, 0d, 1, 255d,
                0d, 0d, Double.POSITIVE_INFINITY,
                Integer.MAX_VALUE, Integer.MAX_VALUE, 1d,
                0, 0, true, error);
        }
    }

    private record VisualFailure(int index, String latex, int dpi, VisualMetric metric) {
    }
}
