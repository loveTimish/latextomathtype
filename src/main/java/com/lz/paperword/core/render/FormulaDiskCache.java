package com.lz.paperword.core.render;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.channels.FileLock;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.StandardOpenOption;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.attribute.FileTime;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.PriorityQueue;
import java.util.Set;
import java.util.UUID;
import java.util.function.LongSupplier;
import java.util.regex.Pattern;

/**
 * An application-owned, quota-limited cache. All file access is relative to securely opened
 * directories; providers without SecureDirectoryStream deliberately fall back to no disk cache.
 * The lock's revision keeps independent JVM indexes consistent without rescanning on hot hits.
 */
final class FormulaDiskCache {
    static final String NAMESPACE = "paperword-formula-cache-v1";
    private static final Path MARKER = Path.of(".owner");
    private static final Path LOCK = Path.of(".lock");
    private static final byte[] OWNER = "paperword-formula-render-cache-v1\n".getBytes(StandardCharsets.US_ASCII);
    private static final long MAGIC = 0x5057524341434831L;
    private static final int HEADER_BYTES = 8 + 8 + 32 + 4;
    private static final Pattern ENTRY_NAME = Pattern.compile("entry-[a-f0-9]{64}\\.pwc");
    private static final Pattern TEMP_NAME = Pattern.compile("pending-[a-f0-9]{32}\\.tmp");
    // Prevent overlapping JVM locks when a runtime property changes or tests use multiple stores.
    private static final Object JVM_LOCK = new Object();

    record Config(Path root, long maxBytes, int maxEntries, long maxEntryBytes,
                  long ttlMillis, long maintenanceMillis) {
        Config {
            root = root.toAbsolutePath().normalize();
            if (maxBytes < 0 || maxEntries < 0 || maxEntryBytes < 0 || ttlMillis < 0 || maintenanceMillis < 0) {
                throw new IllegalArgumentException("Negative disk cache limit");
            }
        }
    }

    private final Config config;
    private final LongSupplier clock;
    private final LinkedHashMap<String, Entry> index = new LinkedHashMap<>(16, .75f, true);
    private long bytes, revision = Long.MIN_VALUE, lastMaintenance = Long.MIN_VALUE;
    private long hits, misses, evictions, expired, corruptions, failures, rejected, writes, scans;

    FormulaDiskCache(Config config, LongSupplier clock) {
        this.config = config;
        this.clock = clock;
    }

    Config config() { return config; }

    byte[] get(String key) {
        synchronized (JVM_LOCK) {
            if (!enabled()) { misses++; return null; }
            try (SecureDirectoryStream<Path> directory = openOwnedDirectory();
                 FileChannel channel = openLock(directory);
                 FileLock lock = channel.tryLock()) {
                if (lock == null) { misses++; return null; }
                long now = clock.getAsLong();
                refresh(directory, channel, now);
                String name = fileName(key);
                Entry entry = index.get(name);
                if (entry == null) { misses++; return null; }
                if (isExpired(entry.created, now)) {
                    delete(directory, name);
                    expired++;
                    publishRevision(channel);
                    misses++;
                    return null;
                }
                byte[] data;
                try {
                    data = readRecord(directory, Path.of(name), now);
                } catch (IOException | RuntimeException invalid) {
                    delete(directory, name);
                    corruptions++;
                    publishRevision(channel);
                    misses++;
                    return null;
                }
                // Persist approximate access order for the next process, at most once per minute.
                if (now - entry.accessed >= 60_000L || now < entry.accessed) {
                    attributes(directory, Path.of(name)).setTimes(FileTime.fromMillis(now), null, null);
                    index.put(name, new Entry(entry.size, entry.created, now));
                }
                hits++;
                return data;
            } catch (IOException | RuntimeException unavailable) {
                failures++;
                misses++;
                return null;
            }
        }
    }

    void put(String key, byte[] data) {
        synchronized (JVM_LOCK) {
            long totalSize = HEADER_BYTES + (long) data.length;
            if (!accepts(data.length)) return;
            try (SecureDirectoryStream<Path> directory = openOwnedDirectory();
                 FileChannel channel = openLock(directory);
                 FileLock lock = channel.tryLock()) {
                if (lock == null) { rejected++; return; }
                long now = clock.getAsLong();
                refresh(directory, channel, now);
                String name = fileName(key);
                Path target = Path.of(name);
                // Never overwrite directories, symlinks, or unrelated entries.
                if (exists(directory, target) && !attributes(directory, target).readAttributes().isRegularFile()) {
                    rejected++;
                    return;
                }
                Path temporary = Path.of("pending-" + UUID.randomUUID().toString().replace("-", "") + ".tmp");
                try {
                    try (SeekableByteChannel out = directory.newByteChannel(temporary,
                        Set.of(StandardOpenOption.WRITE, StandardOpenOption.CREATE_NEW, LinkOption.NOFOLLOW_LINKS))) {
                        ByteBuffer header = ByteBuffer.allocate(HEADER_BYTES);
                        header.putLong(MAGIC).putLong(now).put(hash(data)).putInt(data.length).flip();
                        writeFully(out, header);
                        writeFully(out, ByteBuffer.wrap(data));
                    }
                    directory.move(temporary, directory, target);
                } finally {
                    if (exists(directory, temporary)) directory.deleteFile(temporary);
                }
                attributes(directory, target).setTimes(FileTime.fromMillis(now), null, null);
                Entry previous = index.put(name, new Entry(totalSize, now, now));
                if (previous != null) bytes -= previous.size;
                bytes += totalSize;
                evict(directory);
                publishRevision(channel);
                writes++;
            } catch (IOException | RuntimeException unavailable) {
                // A partial failed write never becomes a trusted in-memory disk index.
                revision = Long.MIN_VALUE;
                failures++;
            }
        }
    }

    void invalidate(String key) {
        synchronized (JVM_LOCK) {
            if (!enabled()) return;
            try (SecureDirectoryStream<Path> directory = openOwnedDirectory();
                 FileChannel channel = openLock(directory);
                 FileLock lock = channel.tryLock()) {
                if (lock == null) return;
                refresh(directory, channel, clock.getAsLong());
                delete(directory, fileName(key));
                corruptions++;
                publishRevision(channel);
            } catch (IOException | RuntimeException unavailable) {
                revision = Long.MIN_VALUE;
                failures++;
            }
        }
    }

    Map<String, Long> statistics() {
        synchronized (JVM_LOCK) {
            return Map.ofEntries(Map.entry("entries", (long) index.size()), Map.entry("bytes", bytes),
                Map.entry("maxBytes", config.maxBytes), Map.entry("maxEntries", (long) config.maxEntries),
                Map.entry("hits", hits), Map.entry("misses", misses), Map.entry("evictions", evictions),
                Map.entry("expired", expired), Map.entry("corruptions", corruptions),
                Map.entry("failures", failures), Map.entry("rejected", rejected),
                Map.entry("writes", writes), Map.entry("scans", scans));
        }
    }

    /** Reject before callers allocate a serialization buffer for an oversized record. */
    boolean accepts(long payloadBytes) {
        synchronized (JVM_LOCK) {
            if (!enabled() || payloadBytes <= 0 || payloadBytes > Integer.MAX_VALUE - HEADER_BYTES
                || payloadBytes > config.maxBytes - HEADER_BYTES || payloadBytes > config.maxEntryBytes - HEADER_BYTES) {
                rejected++;
                return false;
            }
            return true;
        }
    }

    boolean enabled() {
        return config.maxBytes > 0 && config.maxEntries > 0 && config.maxEntryBytes > HEADER_BYTES && config.ttlMillis > 0;
    }

    private void refresh(SecureDirectoryStream<Path> directory, FileChannel lock, long now) throws IOException {
        long diskRevision = readRevision(lock);
        if (revision != diskRevision || lastMaintenance == Long.MIN_VALUE || now < lastMaintenance
            || now - lastMaintenance >= config.maintenanceMillis) {
            boolean changed = scan(directory, now);
            revision = diskRevision;
            if (changed) publishRevision(lock);
            lastMaintenance = now;
        }
    }

    private boolean scan(SecureDirectoryStream<Path> directory, long now) throws IOException {
        scans++;
        boolean changed = false;
        index.clear();
        bytes = 0;
        PriorityQueue<Map.Entry<String, Entry>> retained = new PriorityQueue<>(Comparator
            .<Map.Entry<String, Entry>>comparingLong(e -> e.getValue().accessed).thenComparing(Map.Entry::getKey));
        // A directory stream can only be iterated once. Open '.' securely for each maintenance pass.
        try (SecureDirectoryStream<Path> listing = directory.newDirectoryStream(Path.of("."), LinkOption.NOFOLLOW_LINKS)) {
            for (Path path : listing) {
                Path name = path.getFileName();
                BasicFileAttributes attrs = attributes(directory, name).readAttributes();
                if (!attrs.isRegularFile()) continue;
                if (TEMP_NAME.matcher(name.toString()).matches()) {
                    if (now - attrs.lastModifiedTime().toMillis() >= 3_600_000L) directory.deleteFile(name);
                    continue;
                }
                if (!ENTRY_NAME.matcher(name.toString()).matches()) continue;
                long created;
                try (SeekableByteChannel input = directory.newByteChannel(name,
                    Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS))) {
                    if (attrs.size() <= HEADER_BYTES || attrs.size() > config.maxEntryBytes || attrs.size() > config.maxBytes) {
                        throw new IOException("Invalid cache record size");
                    }
                    ByteBuffer prefix = ByteBuffer.allocate(16);
                    readFully(input, prefix);
                    prefix.flip();
                    if (prefix.getLong() != MAGIC) throw new IOException("Invalid cache record header");
                    created = prefix.getLong();
                } catch (IOException invalid) {
                    directory.deleteFile(name);
                    corruptions++;
                    changed = true;
                    continue;
                }
                if (isExpired(created, now)) {
                    directory.deleteFile(name);
                    expired++;
                    changed = true;
                    continue;
                }
                Entry entry = new Entry(attrs.size(), created, attrs.lastModifiedTime().toMillis());
                retained.add(Map.entry(name.toString(), entry));
                bytes += entry.size;
                while (retained.size() > config.maxEntries || bytes > config.maxBytes) {
                    Map.Entry<String, Entry> victim = retained.remove();
                    directory.deleteFile(Path.of(victim.getKey()));
                    bytes -= victim.getValue().size;
                    evictions++;
                    changed = true;
                }
            }
        }
        while (!retained.isEmpty()) {
            Map.Entry<String, Entry> entry = retained.remove();
            index.put(entry.getKey(), entry.getValue());
        }
        return changed;
    }

    private void evict(SecureDirectoryStream<Path> directory) throws IOException {
        while (index.size() > config.maxEntries || bytes > config.maxBytes) {
            String name = index.keySet().iterator().next();
            delete(directory, name);
            evictions++;
        }
    }

    private void delete(SecureDirectoryStream<Path> directory, String name) throws IOException {
        if (exists(directory, Path.of(name))) {
            // deleteFile is relative and cannot follow a symlink swapped in by another actor.
            directory.deleteFile(Path.of(name));
        }
        Entry entry = index.remove(name);
        if (entry != null) bytes -= entry.size;
    }

    private byte[] readRecord(SecureDirectoryStream<Path> directory, Path name, long now) throws IOException {
        if (!attributes(directory, name).readAttributes().isRegularFile()) throw new IOException("Non-regular cache record");
        try (SeekableByteChannel input = directory.newByteChannel(name,
            Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS))) {
            long size = input.size();
            if (size <= HEADER_BYTES || size > config.maxEntryBytes || size > config.maxBytes || size > Integer.MAX_VALUE) {
                throw new IOException("Invalid cache record size");
            }
            ByteBuffer header = ByteBuffer.allocate(HEADER_BYTES);
            readFully(input, header);
            header.flip();
            if (header.getLong() != MAGIC || isExpired(header.getLong(), now)) throw new IOException("Invalid/expired cache record");
            byte[] expectedHash = new byte[32];
            header.get(expectedHash);
            int length = header.getInt();
            if (length != size - HEADER_BYTES) throw new IOException("Invalid cache record length");
            byte[] body = new byte[length];
            readFully(input, ByteBuffer.wrap(body));
            if (input.size() != size || !MessageDigest.isEqual(expectedHash, hash(body))) throw new IOException("Damaged cache record");
            return body;
        }
    }

    private boolean isExpired(long created, long now) {
        return created < 0 || created > now || now - created >= config.ttlMillis;
    }

    private SecureDirectoryStream<Path> openOwnedDirectory() throws IOException {
        Path namespace = config.root.resolve(NAMESPACE);
        boolean created = false;
        // Check/create one component at a time, never resolve a symlink as a cache directory.
        Path current = namespace.getRoot();
        for (Path component : namespace) {
            current = current.resolve(component);
            if (!Files.exists(current, LinkOption.NOFOLLOW_LINKS)) {
                try {
                    Files.createDirectory(current);
                    if (current.equals(namespace)) created = true;
                } catch (java.nio.file.FileAlreadyExistsException raced) { /* validate below */ }
            }
            if (!Files.isDirectory(current, LinkOption.NOFOLLOW_LINKS)) throw new IOException("Unsafe cache directory");
        }
        SecureDirectoryStream<Path> directory = openSecureDirectory(namespace);
        try {
            if (created) {
                try (SeekableByteChannel marker = directory.newByteChannel(MARKER,
                    Set.of(StandardOpenOption.WRITE, StandardOpenOption.CREATE_NEW, LinkOption.NOFOLLOW_LINKS))) {
                    writeFully(marker, ByteBuffer.wrap(OWNER));
                }
            }
            if (!attributes(directory, MARKER).readAttributes().isRegularFile()) throw new IOException("Unowned cache directory");
            try (SeekableByteChannel marker = directory.newByteChannel(MARKER,
                Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS))) {
                if (marker.size() != OWNER.length) throw new IOException("Unowned cache directory");
                ByteBuffer owner = ByteBuffer.allocate(OWNER.length);
                readFully(marker, owner);
                if (!Arrays.equals(OWNER, owner.array())) throw new IOException("Unowned cache directory");
            }
            return directory;
        } catch (IOException | RuntimeException failure) {
            directory.close();
            throw failure;
        }
    }

    private static SecureDirectoryStream<Path> openSecureDirectory(Path absolute) throws IOException {
        DirectoryStream<Path> root = Files.newDirectoryStream(absolute.getRoot());
        if (!(root instanceof SecureDirectoryStream<Path> secure)) {
            root.close();
            throw new IOException("Secure directory operations are unavailable");
        }
        try {
            for (Path component : absolute) {
                SecureDirectoryStream<Path> next = secure.newDirectoryStream(component, LinkOption.NOFOLLOW_LINKS);
                secure.close();
                secure = next;
            }
            return secure;
        } catch (IOException | RuntimeException failure) {
            secure.close();
            throw failure;
        }
    }

    private FileChannel openLock(SecureDirectoryStream<Path> directory) throws IOException {
        if (exists(directory, LOCK) && !attributes(directory, LOCK).readAttributes().isRegularFile()) {
            throw new IOException("Unsafe cache lock");
        }
        SeekableByteChannel channel = directory.newByteChannel(LOCK,
            Set.of(StandardOpenOption.CREATE, StandardOpenOption.READ, StandardOpenOption.WRITE, LinkOption.NOFOLLOW_LINKS));
        if (channel instanceof FileChannel fileChannel) return fileChannel;
        channel.close();
        throw new IOException("File locking is unavailable");
    }

    private long readRevision(FileChannel lock) throws IOException {
        if (lock.size() == 0) return 0;
        if (lock.size() != Long.BYTES) throw new IOException("Invalid cache lock revision");
        ByteBuffer revisionBytes = ByteBuffer.allocate(Long.BYTES);
        lock.position(0);
        readFully(lock, revisionBytes);
        return revisionBytes.flip().getLong();
    }

    private void publishRevision(FileChannel lock) throws IOException {
        revision++;
        lock.position(0);
        writeFully(lock, ByteBuffer.allocate(Long.BYTES).putLong(revision).flip());
    }

    private static BasicFileAttributeView attributes(SecureDirectoryStream<Path> directory, Path name) {
        return directory.getFileAttributeView(name, BasicFileAttributeView.class, LinkOption.NOFOLLOW_LINKS);
    }

    private static boolean exists(SecureDirectoryStream<Path> directory, Path name) throws IOException {
        try { attributes(directory, name).readAttributes(); return true; }
        catch (java.nio.file.NoSuchFileException missing) { return false; }
    }

    static String fileName(String key) {
        return "entry-" + HexFormat.of().formatHex(hash(key.getBytes(StandardCharsets.UTF_8))) + ".pwc";
    }

    private static byte[] hash(byte[] input) {
        try { return MessageDigest.getInstance("SHA-256").digest(input); }
        catch (NoSuchAlgorithmException impossible) { throw new IllegalStateException(impossible); }
    }

    private static void readFully(SeekableByteChannel channel, ByteBuffer buffer) throws IOException {
        while (buffer.hasRemaining()) if (channel.read(buffer) < 0) throw new IOException("Truncated cache record");
    }

    private static void writeFully(SeekableByteChannel channel, ByteBuffer buffer) throws IOException {
        while (buffer.hasRemaining()) channel.write(buffer);
    }

    private record Entry(long size, long created, long accessed) { }
}
