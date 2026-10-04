package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.FileSystems;
import java.nio.file.attribute.FileTime;
import java.net.URI;
import java.util.Map;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicLong;

import static org.junit.jupiter.api.Assertions.*;

class RenderCacheTest {
    @TempDir Path temporary;

    @Test
    void memoryUsesCombinedPayloadAndKeyWeightAndAccessOrder() {
        RenderMemoryCache cache = new RenderMemoryCache(540, 10);
        cache.put("a", new byte[10]); // 268 bytes, including key and object allowance.
        cache.put("b", new LaTeXImageRenderer.PreviewImage(new byte[1], 1, 1, "p", "i", false)); // 263 bytes.
        assertNotNull(cache.get("a"));
        cache.put("c", new byte[10]);
        assertNull(cache.get("b"));
        assertNotNull(cache.get("a"));
        assertNotNull(cache.get("c"));
        assertEquals(536L, cache.statistics().get("bytes"));
        assertEquals(1L, cache.statistics().get("evictions"));
        cache.put("a", new byte[1]);
        assertEquals(527L, cache.statistics().get("bytes"));
    }

    @Test
    void memoryRejectsOversizedValuesAndHonorsEntryAndDisabledLimits() {
        RenderMemoryCache cache = new RenderMemoryCache(1024, 1);
        cache.put("a", new byte[10]);
        cache.put("b", new byte[10]);
        assertNull(cache.get("a"));
        assertNotNull(cache.get("b"));
        cache.put("too large", new byte[1024]);
        assertNotNull(cache.get("b"));
        cache.put("k".repeat(600), new byte[1]);
        assertEquals(2L, cache.statistics().get("rejected"));
        for (RenderMemoryCache disabled : List.of(new RenderMemoryCache(0, 1), new RenderMemoryCache(1024, 0))) {
            disabled.put("a", new byte[1]);
            assertNull(disabled.get("a"));
            assertEquals(0L, disabled.statistics().get("bytes"));
        }
    }

    @Test
    void concurrentMemoryReadsAndWritesRemainBounded() throws Exception {
        RenderMemoryCache cache = new RenderMemoryCache(4000, 8);
        try (var executor = Executors.newFixedThreadPool(8)) {
            var tasks = new ArrayList<java.util.concurrent.Callable<Void>>();
            for (int thread = 0; thread < 8; thread++) {
                int worker = thread;
                tasks.add(() -> {
                    for (int i = 0; i < 500; i++) {
                        String key = worker + ":" + i;
                        cache.put(key, new byte[100]);
                        cache.get(key);
                        assertTrue(cache.statistics().get("bytes") <= 4000);
                        assertTrue(cache.statistics().get("entries") <= 8);
                    }
                    return null;
                });
            }
            for (var result : executor.invokeAll(tasks)) result.get();
        }
        assertTrue(cache.statistics().get("evictions") > 0);
    }

    @Test
    void diskUsesByteAndEntryQuotasWithLruAndDoesNotScanEveryHit() {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 400, 2, 400, 100_000, clock);
        cache.put("a", new byte[100]); // 152 bytes per record.
        cache.put("b", new byte[100]);
        assertNotNull(cache.get("a"));
        cache.put("c", new byte[100]);
        assertNull(cache.get("b"));
        assertNotNull(cache.get("a"));
        assertNotNull(cache.get("c"));
        long scans = cache.statistics().get("scans");
        for (int i = 0; i < 100; i++) assertNotNull(cache.get("a"));
        assertEquals(scans, cache.statistics().get("scans"));
        assertEquals(304L, cache.statistics().get("bytes"));
        assertEquals(2L, cache.statistics().get("entries"));
        cache.put("large", new byte[500]);
        assertEquals(1L, cache.statistics().get("rejected"));
    }

    @Test
    void diskEnforcesBytesEvenBelowEntryLimit() {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 250, 10, 250, 100_000, clock);
        cache.put("a", new byte[100]);
        cache.put("b", new byte[100]);
        assertNull(cache.get("a"));
        assertNotNull(cache.get("b"));
        assertEquals(152L, cache.statistics().get("bytes"));
    }

    @Test
    void diskTtlExpiresAfterWriteEvenWhenFrequentlyAccessed() {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 4096, 10, 4096, 100, clock);
        cache.put("a", new byte[] {1, 2, 3});
        clock.set(1099);
        assertArrayEquals(new byte[] {1, 2, 3}, cache.get("a"));
        clock.set(1100);
        assertNull(cache.get("a"));
        assertEquals(0L, cache.statistics().get("bytes"));
        assertEquals(1L, cache.statistics().get("expired"));
    }

    @Test
    void backwardClockNeverResurrectsFutureDatedEntries() {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 4096, 10, 4096, 1000, clock);
        cache.put("a", new byte[] {1});
        clock.set(999);
        assertNull(cache.get("a"));
    }

    @Test
    void freshProcessIndexSharesQuotaAndReadsExistingEntries() {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache first = disk(temporary, 4096, 2, 4096, 1000, clock);
        FormulaDiskCache second = disk(temporary, 4096, 2, 4096, 1000, clock);
        first.put("a", new byte[] {1});
        clock.incrementAndGet();
        second.put("b", new byte[] {2});
        assertArrayEquals(new byte[] {2}, first.get("b"));
        clock.incrementAndGet();
        second.put("c", new byte[] {3});
        assertNull(first.get("a"));
        assertArrayEquals(new byte[] {2}, first.get("b"));
        assertArrayEquals(new byte[] {3}, first.get("c"));
        long scans = first.statistics().get("scans");
        second.get("b");
        first.get("b");
        assertEquals(scans, first.statistics().get("scans"), "read-only refreshes must not ping-pong revisions");
    }

    @Test
    void diskDetectsCorruptionAndCanRewriteTheSameKey() throws Exception {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 4096, 10, 4096, 1000, clock);
        byte[] value = new byte[] {1, 2, 3};
        cache.put("a", value);
        Path file = entry("a");
        byte[] corrupt = Files.readAllBytes(file);
        corrupt[corrupt.length - 1] ^= 0x40;
        Files.write(file, corrupt);
        assertNull(cache.get("a"));
        assertFalse(Files.exists(file));
        assertEquals(1L, cache.statistics().get("corruptions"));
        cache.put("a", value);
        assertArrayEquals(value, cache.get("a"));
        Files.write(file, new byte[] {1});
        assertNull(cache.get("a"));
    }

    @Test
    void cleanupNeverTouchesOrdinaryFilesOrLegacyCacheOutsideNamespace() throws Exception {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 4096, 1, 4096, 100, clock);
        Path ordinary = temporary.resolve("ordinary.png");
        Files.writeString(ordinary, "user-owned");
        Path legacy = temporary.resolve("old-cache.properties");
        Files.writeString(legacy, "legacy cache, not owned by new namespace");
        cache.put("a", new byte[] {1});
        Path nestedOrdinary = temporary.resolve(FormulaDiskCache.NAMESPACE).resolve("notes.txt");
        Files.writeString(nestedOrdinary, "keep me");
        clock.set(2000);
        cache.put("b", new byte[] {2});
        assertEquals("user-owned", Files.readString(ordinary));
        assertTrue(Files.exists(legacy));
        assertEquals("keep me", Files.readString(nestedOrdinary));
        assertFalse(Files.exists(entry("a")));
    }

    @Test
    void unmarkedNamespaceIsNotAdoptedOrCleaned() throws Exception {
        Path namespace = Files.createDirectory(temporary.resolve(FormulaDiskCache.NAMESPACE));
        Path userFile = namespace.resolve(FormulaDiskCache.fileName("a"));
        Files.writeString(userFile, "user data");
        FormulaDiskCache cache = disk(temporary, 4096, 10, 4096, 1000, new AtomicLong(1000));
        cache.put("a", new byte[] {1});
        assertNull(cache.get("a"));
        assertEquals("user data", Files.readString(userFile));
        assertFalse(Files.exists(namespace.resolve(".owner")));
    }

    @Test
    void symlinkedNamespacesAndAncestorsAreNeverFollowed() throws Exception {
        Path outside = Files.createDirectory(temporary.resolve("outside"));
        Path root = Files.createDirectory(temporary.resolve("root"));
        Files.createSymbolicLink(root.resolve(FormulaDiskCache.NAMESPACE), outside);
        FormulaDiskCache cache = disk(root, 4096, 10, 4096, 1000, new AtomicLong(1000));
        cache.put("a", new byte[] {1});
        assertNull(cache.get("a"));
        try (var files = Files.list(outside)) { assertEquals(0, files.count()); }
        Path linked = Files.createSymbolicLink(temporary.resolve("linked"), outside);
        FormulaDiskCache throughParent = disk(linked.resolve("child"), 4096, 10, 4096, 1000, new AtomicLong(1000));
        throughParent.put("b", new byte[] {2});
        assertFalse(Files.exists(outside.resolve("child")));
    }

    @Test
    void symlinkedEntryIsNeitherReadNorOverwrittenAndTargetIsNeverDeleted() throws Exception {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 4096, 10, 4096, 1000, clock);
        cache.put("init", new byte[] {1});
        Path userFile = temporary.resolve("user.dat");
        Files.writeString(userFile, "keep me");
        Files.createSymbolicLink(entry("a"), userFile);
        assertNull(cache.get("a"));
        cache.put("a", new byte[] {2});
        assertTrue(Files.isSymbolicLink(entry("a")));
        assertEquals("keep me", Files.readString(userFile));
    }

    @Test
    void disabledDiskCacheDoesNotCreateDirectories() {
        FormulaDiskCache disabled = disk(temporary.resolve("disabled"), 0, 10, 4096, 1000, new AtomicLong(1000));
        disabled.put("a", new byte[] {1});
        assertNull(disabled.get("a"));
        assertFalse(Files.exists(temporary.resolve("disabled")));
    }

    @Test
    void concurrentDiskWritersPreserveAtomicRecordsAndQuota() throws Exception {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache cache = disk(temporary, 4096, 8, 4096, 1000, clock);
        try (var executor = Executors.newFixedThreadPool(6)) {
            var tasks = new ArrayList<java.util.concurrent.Callable<Void>>();
            for (int thread = 0; thread < 6; thread++) {
                int worker = thread;
                tasks.add(() -> {
                    for (int i = 0; i < 30; i++) {
                        String key = worker + ":" + i;
                        byte[] data = new byte[100];
                        Arrays.fill(data, (byte) worker);
                        cache.put(key, data);
                        byte[] result = cache.get(key);
                        if (result != null) assertArrayEquals(data, result);
                    }
                    return null;
                });
            }
            for (var result : executor.invokeAll(tasks)) result.get();
        }
        assertTrue(cache.statistics().get("entries") <= 8);
        assertTrue(cache.statistics().get("bytes") <= 4096);
        assertEquals(0L, cache.statistics().get("failures"));
    }

    @Test
    void rendererPreviewRoundTripPreservesAllGeometry() throws Exception {
        FormulaDiskCache disk = disk(temporary, 4096, 10, 4096, 1000, new AtomicLong(1000));
        LaTeXImageRenderer renderer = new LaTeXImageRenderer(new RenderMemoryCache(1024, 1), disk);
        var preview = new LaTeXImageRenderer.PreviewImage(new byte[] {1, 2, 3}, 23, 14,
            "wmf", "image/x-wmf", false, 2.125d, 17.25d, 10.5d);
        Method write = LaTeXImageRenderer.class.getDeclaredMethod("writePreviewToDisk", String.class, LaTeXImageRenderer.PreviewImage.class);
        Method read = LaTeXImageRenderer.class.getDeclaredMethod("readPreviewFromDisk", String.class);
        write.setAccessible(true);
        read.setAccessible(true);
        write.invoke(renderer, "fixture", preview);
        var restored = (LaTeXImageRenderer.PreviewImage) read.invoke(renderer, "fixture");
        assertNotNull(restored);
        assertArrayEquals(preview.data(), restored.data());
        assertEquals(preview.widthPx(), restored.widthPx());
        assertEquals(preview.heightPx(), restored.heightPx());
        assertEquals(preview.widthPt(), restored.widthPt());
        assertEquals(preview.heightPt(), restored.heightPt());
        assertEquals(preview.depthPt(), restored.depthPt());
        assertEquals(preview.contentType(), restored.contentType());
        assertEquals(preview.placeholder(), restored.placeholder());
        assertEquals(1L, renderer.cacheStatistics().get("disk.hits"));
        disk.put("broken", new byte[] {2, 0});
        assertNull(read.invoke(renderer, "broken"));
        assertFalse(Files.exists(entry("broken")));
    }

    @Test
    void recomputingPngAfterEvictionPreservesOutputAndDisabledDiskBehavior() {
        RenderMemoryCache memory = new RenderMemoryCache(1024 * 1024, 1);
        FormulaDiskCache disk = disk(temporary, 0, 0, 0, 0, new AtomicLong(1000));
        LaTeXImageRenderer renderer = new LaTeXImageRenderer(memory, disk);
        String oldLatex = System.getProperty("paperword.latex.command");
        try {
            System.setProperty("paperword.latex.command", temporary.resolve("missing-latex").toString());
            byte[] first = renderer.renderToPng("x+1");
            assertNotNull(first);
            renderer.renderToPng("x+2");
            byte[] recomputed = renderer.renderToPng("x+1");
            assertArrayEquals(first, recomputed);
            assertEquals(2L, renderer.cacheStatistics().get("memory.evictions"));
            assertFalse(Files.exists(temporary.resolve(FormulaDiskCache.NAMESPACE)));
        } finally {
            if (oldLatex == null) System.clearProperty("paperword.latex.command");
            else System.setProperty("paperword.latex.command", oldLatex);
        }
    }

    @Test
    void unsupportedDirectoryProviderStillRendersWithoutWritingCache() throws Exception {
        URI zipUri = URI.create("jar:" + temporary.resolve("cache.zip").toUri());
        try (var zip = FileSystems.newFileSystem(zipUri, Map.of("create", "true"))) {
            FormulaDiskCache disk = new FormulaDiskCache(new FormulaDiskCache.Config(
                zip.getPath("/cache"), 4096, 10, 4096, 1000, 60_000), () -> 1000);
            LaTeXImageRenderer renderer = new LaTeXImageRenderer(new RenderMemoryCache(1024 * 1024, 10), disk);
            String oldLatex = System.getProperty("paperword.latex.command");
            try {
                System.setProperty("paperword.latex.command", temporary.resolve("missing-latex").toString());
                assertTrue(renderer.renderToPng("x+1").length > 0);
                assertTrue(disk.statistics().get("failures") > 0);
                try (var files = Files.walk(zip.getPath("/"))) {
                    assertFalse(files.anyMatch(path -> path.toString().endsWith(".pwc")));
                }
            } finally {
                if (oldLatex == null) System.clearProperty("paperword.latex.command");
                else System.setProperty("paperword.latex.command", oldLatex);
            }
        }
    }

    @Test
    void oleGeometryAndBytesSurviveMemoryEviction() {
        RenderMemoryCache memory = new RenderMemoryCache(1024 * 1024, 1);
        FormulaDiskCache disabled = disk(temporary, 0, 0, 0, 0, new AtomicLong(1000));
        LaTeXImageRenderer renderer = new LaTeXImageRenderer(memory, disabled);
        var first = renderer.renderForOlePreview("\\frac{x+1}{2}", 41.25d, 23.75d);
        renderer.renderForOlePreview("x+2");
        var recomputed = renderer.renderForOlePreview("\\frac{x+1}{2}", 41.25d, 23.75d);
        assertArrayEquals(first.data(), recomputed.data());
        assertEquals(first.depthPt(), recomputed.depthPt());
        assertEquals(first.widthPt(), recomputed.widthPt());
        assertEquals(first.heightPt(), recomputed.heightPt());
        assertTrue(WmfPreviewInspector.inspect(recomputed.data()).pureVector());
        assertEquals(2L, renderer.cacheStatistics().get("memory.evictions"));
        assertEquals(0L, renderer.cacheStatistics().get("memory.inFlight"));
    }

    @Test
    void initialMaintenanceAppliesSmallerQuotaAndRemovesOnlyStaleNamedTemporaryFiles() throws Exception {
        AtomicLong clock = new AtomicLong(1000);
        FormulaDiskCache original = disk(temporary, 4096, 10, 4096, 10_000_000, clock);
        for (String key : List.of("a", "b", "c")) {
            original.put(key, new byte[] {1});
            clock.incrementAndGet();
        }
        Path namespace = temporary.resolve(FormulaDiskCache.NAMESPACE);
        Path stale = namespace.resolve("pending-" + "a".repeat(32) + ".tmp");
        Files.writeString(stale, "abandoned staging file");
        Files.setLastModifiedTime(stale, FileTime.fromMillis(0));
        Path fresh = namespace.resolve("pending-" + "b".repeat(32) + ".tmp");
        Files.writeString(fresh, "recent staging file");
        clock.set(3_700_000);
        Files.setLastModifiedTime(fresh, FileTime.fromMillis(clock.get()));
        Path ordinary = namespace.resolve("pending-user.tmp");
        Files.writeString(ordinary, "user file");
        FormulaDiskCache smaller = disk(temporary, 4096, 1, 4096, 10_000_000, clock);
        assertNotNull(smaller.get("c"));
        assertEquals(1L, smaller.statistics().get("entries"));
        assertEquals(2L, smaller.statistics().get("evictions"));
        assertFalse(Files.exists(stale));
        assertTrue(Files.exists(fresh));
        assertEquals("user file", Files.readString(ordinary));
    }

    private FormulaDiskCache disk(Path root, long bytes, int entries, long entryBytes, long ttl, AtomicLong clock) {
        if (bytes > 0 && entries > 0 && entryBytes > 52 && ttl > 0) {
            try (var directory = Files.newDirectoryStream(temporary)) {
                org.junit.jupiter.api.Assumptions.assumeTrue(directory instanceof SecureDirectoryStream<?>,
                    "This provider deliberately has no disk caching; fallback rendering is tested separately");
            } catch (IOException failure) {
                throw new UncheckedIOException(failure);
            }
        }
        return new FormulaDiskCache(new FormulaDiskCache.Config(root, bytes, entries, entryBytes, ttl, 60_000), clock::get);
    }

    private Path entry(String key) {
        return temporary.resolve(FormulaDiskCache.NAMESPACE).resolve(FormulaDiskCache.fileName(key));
    }
}
