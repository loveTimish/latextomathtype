package com.lz.paperword.core.render;

import com.lz.paperword.service.PaperExportService;
import java.lang.management.ManagementFactory;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.time.Duration;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import jdk.jfr.Recording;

/** Steady-state probe for synthetic formula-only DOCX exports; no production inputs. */
public final class WarmDocxProbe {
    static final com.sun.management.ThreadMXBean THREADS =
        (com.sun.management.ThreadMXBean) ManagementFactory.getThreadMXBean();
    static final com.sun.management.OperatingSystemMXBean OS =
        (com.sun.management.OperatingSystemMXBean) ManagementFactory.getOperatingSystemMXBean();
    static final java.lang.management.CompilationMXBean COMPILER = ManagementFactory.getCompilationMXBean();
    static final List<java.lang.management.GarbageCollectorMXBean> GC = ManagementFactory.getGarbageCollectorMXBeans();
    record Output(int kind, int round, byte[] bytes) { }

    static long gcMillis() { return GC.stream().mapToLong(x -> Math.max(0, x.getCollectionTime())).sum(); }
    static long allocations() { return THREADS.getThreadAllocatedBytes(Thread.currentThread().threadId()); }
    static Object cacheStats() throws Exception {
        try { return LaTeXImageRenderer.class.getMethod("cacheStatistics").invoke(new LaTeXImageRenderer()); }
        catch (NoSuchMethodException ignored) { return "not_available_in_baseline"; }
    }
    static void emit(String kind, Object value) throws Exception {
        System.out.println(kind + " " + LinuxBenchmark.JSON.writeValueAsString(value));
    }
    static String digest(byte[] bytes) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes));
    }
    static Map<String,String> immutableParts(byte[] bytes) throws Exception {
        var result = new TreeMap<String,String>();
        for (var e : LinuxBenchmark.unzip(bytes).entrySet()) {
            if (e.getKey().startsWith("word/media/") || e.getKey().startsWith("word/embeddings/"))
                result.put(e.getKey(), digest(e.getValue()));
        }
        return result;
    }
    public static void main(String[] args) throws Exception {
        ((ch.qos.logback.classic.Logger) org.slf4j.LoggerFactory.getLogger(org.slf4j.Logger.ROOT_LOGGER_NAME))
            .setLevel(ch.qos.logback.classic.Level.WARN);
        LinuxBenchmark.out = Path.of(args[0]); Files.createDirectories(LinuxBenchmark.out);
        int warmups = Integer.parseInt(args[1]), rounds = Integer.parseInt(args[2]);
        boolean profile = args.length > 3 && Boolean.parseBoolean(args[3]);
        System.setProperty("paperword.render.cache.enabled", "false");
        var svc = new PaperExportService();
        var requests = List.of(LinuxBenchmark.paper(40, 0), LinuxBenchmark.paper(40, 1));
        var outputs = new ArrayList<Output>();
        for (int round = 0; round < warmups; round++) {
            for (int order = 0; order < 2; order++) {
                int kind = (round + order) % 2;
                long start = System.nanoTime();
                byte[] bytes = svc.export(requests.get(kind));
                emit("WARMUP", Map.of("round", round, "kind", kind, "ms", (System.nanoTime() - start) / 1e6,
                    "bytes", bytes.length));
            }
        }
        long mathJaxBefore = LinuxBenchmark.requests();
        emit("STATE_BEFORE", Map.of("worker_requests", mathJaxBefore, "cache", cacheStats(),
            "compiled_ms", COMPILER.getTotalCompilationTime(), "gc_ms", gcMillis()));
        Recording recording = null;
        if (profile) {
            recording = new Recording();
            recording.enable("jdk.ExecutionSample").withPeriod(Duration.ofMillis(5));
            recording.enable("jdk.NativeMethodSample").withPeriod(Duration.ofMillis(5));
            recording.enable("jdk.JavaMonitorEnter").withThreshold(Duration.ZERO).withStackTrace();
            recording.enable("jdk.FileRead").withThreshold(Duration.ZERO).withStackTrace();
            recording.enable("jdk.FileWrite").withThreshold(Duration.ZERO).withStackTrace();
            recording.enable("jdk.GarbageCollection");
            recording.enable("jdk.Compilation").withThreshold(Duration.ZERO);
            recording.start();
        }
        for (int round = 0; round < rounds; round++) {
            for (int order = 0; order < 2; order++) {
                int kind = (round + order) % 2;
                long alloc = allocations(), compiled = COMPILER.getTotalCompilationTime(), gc = gcMillis();
                long processCpu = OS.getProcessCpuTime(), threadCpu = THREADS.getCurrentThreadCpuTime();
                long start = System.nanoTime();
                byte[] bytes = svc.export(requests.get(kind));
                long elapsed = System.nanoTime() - start;
                long cpuElapsed = THREADS.getCurrentThreadCpuTime() - threadCpu;
                long processElapsed = OS.getProcessCpuTime() - processCpu;
                long allocated = allocations() - alloc;
                outputs.add(new Output(kind, round, bytes));
                var metrics = new LinkedHashMap<String,Object>();
                metrics.put("case", kind == 0 ? "repeat40" : "unique40"); metrics.put("round", round);
                metrics.put("ms", elapsed / 1e6); metrics.put("thread_cpu_ms", cpuElapsed / 1e6);
                metrics.put("process_cpu_ms", processElapsed / 1e6); metrics.put("allocated_bytes", allocated);
                metrics.put("compilation_ms", COMPILER.getTotalCompilationTime() - compiled);
                metrics.put("gc_ms", gcMillis() - gc); metrics.put("bytes", bytes.length);
                emit("SAMPLE", metrics);
            }
        }
        if (recording != null) {
            recording.stop(); recording.dump(LinuxBenchmark.out.resolve("hot-export.jfr")); recording.close();
        }
        long workerDelta = LinuxBenchmark.requests() - mathJaxBefore;
        if (workerDelta != 0) throw new IllegalStateException("Not fully warm: additional MathJax requests = " + workerDelta);
        emit("STATE_AFTER", Map.of("worker_requests", LinuxBenchmark.requests(), "worker_requests_during_samples", workerDelta,
            "cache", cacheStats(), "compiled_ms", COMPILER.getTotalCompilationTime(), "gc_ms", gcMillis()));
        // All correctness/file work is deliberately outside the measured loop.
        var expectedMtef = new TreeMap<Integer,List<String>>();
        var expectedParts = new TreeMap<Integer,Map<String,String>>();
        for (Output output : outputs) {
            var mtef = LinuxBenchmark.check(output.bytes, 40, "paper_" + output.kind + "_" + output.round);
            var parts = immutableParts(output.bytes);
            if (expectedMtef.putIfAbsent(output.kind, mtef) != null && !expectedMtef.get(output.kind).equals(mtef))
                throw new IllegalStateException("MTEF changed across rounds");
            if (expectedParts.putIfAbsent(output.kind, parts) != null && !expectedParts.get(output.kind).equals(parts))
                throw new IllegalStateException("Media or OLE changed across rounds");
        }
        var validation = Map.of("documents", outputs.size(), "formulas", outputs.size() * 40,
            "mtef_sha256", expectedMtef, "part_sha256", expectedParts);
        Files.writeString(LinuxBenchmark.out.resolve("validation.json"), LinuxBenchmark.JSON.writeValueAsString(validation));
        emit("VALIDATED", Map.of("documents", outputs.size(), "formulas", outputs.size() * 40,
            "pure_vector_previews", outputs.size() * 40, "stable_rounds", true));
    }
}
