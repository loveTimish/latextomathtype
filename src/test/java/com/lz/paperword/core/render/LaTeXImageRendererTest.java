package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LaTeXImageRendererTest {

    @Test
    void shouldParseSvgPtDimensions() {
        byte[] svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="36pt" height="18pt" viewBox="0 0 36 18">
              <path d="M0 0h36v18H0z"/>
            </svg>
            """.getBytes(StandardCharsets.UTF_8);

        LaTeXImageRenderer.SvgDimensions dimensions = LaTeXImageRenderer.extractSvgDisplayDimensions(svg);
        assertEquals(36.0f, dimensions.widthPt(), 0.01f);
        assertEquals(18.0f, dimensions.heightPt(), 0.01f);
    }

    @Test
    void shouldParseSvgPxDimensions() {
        byte[] svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="40px" height="20px" viewBox="0 0 40 20">
              <path d="M0 0h40v20H0z"/>
            </svg>
            """.getBytes(StandardCharsets.UTF_8);

        LaTeXImageRenderer.SvgDimensions dimensions = LaTeXImageRenderer.extractSvgDisplayDimensions(svg);
        assertEquals(30.0f, dimensions.widthPt(), 0.05f);
        assertEquals(15.0f, dimensions.heightPt(), 0.05f);
    }

    @Test
    void shouldKeepLongDivisionPreviewHeaderOnly() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(), "\\longdiv[570]{6}{3420}");
        assertTrue(normalized.contains("\\overset{570}{\\overline{\\left)3420\\right.}}"));
        assertFalse(normalized.contains("\\underline{30}"), "预览图不应再自动补第一步乘积");
        assertFalse(normalized.contains("\\underline{42}"), "预览图不应再自动补后续步骤");
    }

    @Test
    void shouldReplaceCompositeLongDivisionHeaderInsideSingleBlock() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(
            new LaTeXImageRenderer(),
            "\\longdiv[570]{6}{3420}\\begin{array}{l}\\text{   }\\underline{30}\\\\\\text{    }42\\end{array}"
        );
        assertFalse(normalized.contains("\\longdiv[570]{6}{3420}"), "单块复合长除法预览时应替换掉原始 longdiv 命令");
        assertTrue(normalized.contains("\\overset{570}{\\overline{\\left)3420\\right.}}"), "预览图应保留长除法头部");
        assertTrue(normalized.contains("\\begin{array}{l}"), "单块复合长除法的步骤区应继续保留");
        assertTrue(normalized.contains("\\underline{30}"), "显式步骤区应继续进入预览渲染");
    }

    @Test
    void shouldWriteRenderedFormulaToDiskCache(@TempDir Path tempDir) throws Exception {
        String oldEnabled = System.getProperty("paperword.render.cache.enabled");
        String oldDir = System.getProperty("paperword.render.cache.dir");
        try {
            System.setProperty("paperword.render.cache.enabled", "true");
            System.setProperty("paperword.render.cache.dir", tempDir.toString());

            byte[] png = new LaTeXImageRenderer().renderToPng("x+" + UUID.randomUUID(), 13f);

            assertTrue(png.length > 0);
            try (var stream = Files.walk(tempDir)) {
                assertTrue(stream.anyMatch(path -> path.getFileName().toString().endsWith(".png")),
                    "rendered formulas should be written to the disk cache");
            }
        } finally {
            restoreProperty("paperword.render.cache.enabled", oldEnabled);
            restoreProperty("paperword.render.cache.dir", oldDir);
        }
    }

    @Test
    void cacheKeyIncludesWmfRenderProperties() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("cacheKey", String.class, String.class, float.class);
        method.setAccessible(true);
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        String oldScale = System.getProperty("paperword.wmf.textWidth.scale");
        try {
            System.setProperty("paperword.wmf.textWidth.scale", "1.00");
            String defaultKey = (String) method.invoke(renderer, "ole-target-10.00x10.00", "x+1", 12f);
            System.setProperty("paperword.wmf.textWidth.scale", "1.25");
            String tunedKey = (String) method.invoke(renderer, "ole-target-10.00x10.00", "x+1", 12f);

            assertFalse(defaultKey.equals(tunedKey), "WMF render tuning must invalidate preview cache keys");
        } finally {
            restoreProperty("paperword.wmf.textWidth.scale", oldScale);
        }
    }

    @Test
    void runCommandTimesOutWithoutWaitingForStreamClose(@TempDir Path tempDir) throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("runCommand", List.class, Path.class, int.class);
        method.setAccessible(true);
        String java = Path.of(System.getProperty("java.home"), "bin", isWindows() ? "java.exe" : "java").toString();
        long start = System.nanoTime();
        Object result = method.invoke(
            new LaTeXImageRenderer(),
            List.of(java, "-cp", System.getProperty("java.class.path"), HangingProcess.class.getName()),
            tempDir,
            1
        );
        long elapsedMs = (System.nanoTime() - start) / 1_000_000L;
        Method exitCode = result.getClass().getDeclaredMethod("exitCode");
        exitCode.setAccessible(true);

        assertEquals(-1, exitCode.invoke(result));
        assertTrue(elapsedMs < 5000, "runCommand timeout must not wait for the child stream to close");
    }

    public static class HangingProcess {
        public static void main(String[] args) throws Exception {
            System.out.print("started");
            System.out.flush();
            Thread.sleep(60_000L);
        }
    }

    private void restoreProperty(String name, String value) {
        if (value == null) {
            System.clearProperty(name);
        } else {
            System.setProperty(name, value);
        }
    }

    private boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase().contains("win");
    }
}
