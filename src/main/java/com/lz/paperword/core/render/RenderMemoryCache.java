package com.lz.paperword.core.render;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ConcurrentHashMap;

/** Shared, byte-weighted access-order LRU. A zero limit disables retention, not rendering. */
final class RenderMemoryCache {
    private final long maxBytes;
    private final int maxEntries;
    private final LinkedHashMap<String, Entry> entries = new LinkedHashMap<>(16, .75f, true);
    final Map<String, CompletableFuture<LaTeXImageRenderer.PreviewImage>> inFlight = new ConcurrentHashMap<>();
    private long bytes, hits, misses, evictions, rejected;

    RenderMemoryCache(long maxBytes, int maxEntries) {
        if (maxBytes < 0 || maxEntries < 0) throw new IllegalArgumentException("Negative cache limit");
        this.maxBytes = maxBytes;
        this.maxEntries = maxEntries;
    }

    synchronized Object get(String key) {
        Entry entry = entries.get(key);
        if (entry == null) { misses++; return null; }
        hits++;
        return entry.value;
    }

    synchronized void put(String key, Object value) {
        long payloadBytes;
        if (value instanceof LaTeXImageRenderer.PreviewImage preview) {
            payloadBytes = preview.data().length + 2L * (preview.extension().length() + preview.contentType().length());
        } else if (value instanceof byte[] data) {
            payloadBytes = data.length;
        } else {
            throw new IllegalArgumentException("Unsupported render cache value");
        }
        // Includes the retained key and a conservative allowance for map/record overhead.
        long weight = 256L + 2L * key.length() + payloadBytes;
        Entry old = entries.remove(key);
        if (old != null) bytes -= old.weight;
        if (maxEntries == 0 || weight > maxBytes) { rejected++; return; }
        while (!entries.isEmpty() && (entries.size() >= maxEntries || bytes > maxBytes - weight)) {
            var iterator = entries.entrySet().iterator();
            bytes -= iterator.next().getValue().weight;
            iterator.remove();
            evictions++;
        }
        entries.put(key, new Entry(value, weight));
        bytes += weight;
    }

    synchronized Map<String, Long> statistics() {
        return Map.of("entries", (long) entries.size(), "bytes", bytes,
            "maxBytes", maxBytes, "maxEntries", (long) maxEntries, "hits", hits,
            "misses", misses, "evictions", evictions, "rejected", rejected,
            "inFlight", (long) inFlight.size());
    }

    private record Entry(Object value, long weight) { }
}
