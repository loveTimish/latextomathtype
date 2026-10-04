package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.lang.reflect.*;
import java.nio.file.*;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.*;
import static org.junit.jupiter.api.Assertions.*;

class RenderConcurrencyRegressionTest {
    private static long requestCount() throws Exception {
        Field field = LaTeXImageRenderer.class.getDeclaredField("mathJaxRequestId");
        field.setAccessible(true);
        return field.getLong(null);
    }
    private static void stopWorker() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("stopMathJaxWorker");
        method.setAccessible(true);
        method.invoke(null);
    }
    @Test
    void concurrentColdMissRendersOnlyOnce() throws Exception {
        String latex = "\\frac{x^2+" + Long.toUnsignedString(System.nanoTime()) + "}{x+1}";
        long before = requestCount();
        var gate = new CountDownLatch(1);
        try (var pool = Executors.newFixedThreadPool(8)) {
            var futures = new ArrayList<Future<LaTeXImageRenderer.PreviewImage>>();
            for (int i = 0; i < 8; i++) futures.add(pool.submit(() -> {
                gate.await();
                return new LaTeXImageRenderer().renderForOlePreview(latex);
            }));
            gate.countDown();
            byte[] first = futures.getFirst().get(20, TimeUnit.SECONDS).data();
            for (var future : futures) assertArrayEquals(first, future.get(20, TimeUnit.SECONDS).data());
            assertTrue(WmfPreviewInspector.inspect(first).pureVector());
        }
        assertEquals(1, requestCount() - before);
    }
    @Test
    void timeoutKillsWorkerAndNextRequestRecovers(@TempDir Path dir) throws Exception {
        String oldScript = System.getProperty("paperword.mathjax.script");
        String oldTimeout = System.getProperty("paperword.latex.timeout.seconds");
        String oldCache = System.getProperty("paperword.render.cache.enabled");
        Path script = dir.resolve("hang.cjs");
        Files.writeString(script, "process.stdin.resume(); setInterval(()=>{},1000);\n");
        stopWorker();
        try {
            System.setProperty("paperword.mathjax.script", script.toString());
            System.setProperty("paperword.latex.timeout.seconds", "1");
            System.setProperty("paperword.render.cache.enabled", "false");
            assertTimeoutPreemptively(Duration.ofSeconds(5), () ->
                assertThrows(IllegalStateException.class, () ->
                    new LaTeXImageRenderer().renderForOlePreview("x+" + System.nanoTime())));
            restore("paperword.mathjax.script", oldScript);
            restore("paperword.latex.timeout.seconds", oldTimeout);
            var image = new LaTeXImageRenderer().renderForOlePreview("x+" + System.nanoTime());
            assertTrue(WmfPreviewInspector.inspect(image.data()).pureVector());
        } finally {
            stopWorker();
            restore("paperword.mathjax.script", oldScript);
            restore("paperword.latex.timeout.seconds", oldTimeout);
            restore("paperword.render.cache.enabled", oldCache);
        }
    }
    @Test
    void failedMissIsNotRetained() throws Exception {
        String latex = "\\undefinedRegressionCommand" + Long.toUnsignedString(System.nanoTime());
        long before = requestCount();
        var renderer = new LaTeXImageRenderer();
        assertThrows(RuntimeException.class, () -> renderer.renderForOlePreview(latex));
        assertThrows(RuntimeException.class, () -> renderer.renderForOlePreview(latex));
        assertEquals(2, requestCount() - before);
    }
    @Test
    void interruptedWaiterDoesNotCancelSharedWork() throws Exception {
        String latex = "x+" + System.nanoTime();
        var memoryCache = new RenderMemoryCache(1024 * 1024, 10);
        var renderer = new LaTeXImageRenderer(memoryCache, null);
        Method keyMethod = LaTeXImageRenderer.class.getDeclaredMethod("cacheKey", String.class, String.class, float.class);
        keyMethod.setAccessible(true);
        String key = (String) keyMethod.invoke(renderer, "ole-wmf", latex, 9.02f);
        var pending = memoryCache.inFlight;
        var shared = new CompletableFuture<LaTeXImageRenderer.PreviewImage>();
        pending.put(key, shared);
        try (var pool = Executors.newSingleThreadExecutor()) {
            try {
                pool.submit(() -> {
                    Thread.currentThread().interrupt();
                    assertThrows(IllegalStateException.class, () -> renderer.renderForOlePreview(latex));
                    assertTrue(Thread.currentThread().isInterrupted());
                }).get(2, TimeUnit.SECONDS);
                assertFalse(shared.isDone(), "one caller cannot cancel work needed by other callers");
            } finally {
                // Release the simulated owner before ExecutorService.close(), even
                // if an uninterruptible waiter regression makes the assertion time out.
                shared.completeExceptionally(new IllegalStateException("test cleanup"));
            }
        } finally {
            pending.remove(key, shared);
        }
    }
    private static void restore(String name, String value) {
        if (value == null) System.clearProperty(name); else System.setProperty(name, value);
    }
}
