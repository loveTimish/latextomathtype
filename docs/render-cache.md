# Bounded formula-render caches

`LaTeXImageRenderer` shares a process-wide, thread-safe memory LRU across its
instances. PNG bytes and preview bytes (including physical dimensions and
baseline depth) use the **same** byte and entry budgets. Eviction only discards
reusable work; the next request computes the same output. Concurrent cold OLE
requests retain the existing per-key singleflight and interruption behavior.

## JVM properties

Pass these with `-D`, before starting Java. Memory limits are read once when the
renderer class initializes. Disk configuration is re-read on disk access; changing
it replaces the one current disk index, so old configuration indexes do not
accumulate.

| Property (`paperword.render.cache.` prefix) | Default | Meaning |
| --- | ---: | --- |
| `memory.maxBytes` | 67108864 (64 MiB) | Combined retained weight: payload bytes + UTF-16 key/string bytes + 256-byte per-entry object allowance |
| `memory.maxEntries` | 2048 | Combined PNG/preview entry limit |
| `enabled` | `true` | Existing **disk-only** switch, unchanged for compatibility |
| `dir` | `data/cache/formula-render` | Parent of the application-owned namespace |
| `disk.maxBytes` | 268435456 (256 MiB) | Total retained record-file bytes, including headers and metadata |
| `disk.maxEntries` | 8192 | Record-file entry limit |
| `disk.maxEntryBytes` | 16777216 (16 MiB) | Maximum individual record-file size; larger results are still returned normally |
| `disk.ttlSeconds` | 604800 (7 days) | Expire after write; a read does not extend the lifetime |
| `disk.maintenanceSeconds` | 60 | Full-directory maintenance interval on disk operations |

Setting either memory limit to `0` disables memory retention. Setting a disk
capacity/entry limit or TTL to `0` disables disk retention. A zero maintenance
interval scans on every disk operation and is useful only for diagnostics/tests.
Malformed or negative numeric properties fall back to their defaults. There is
no option for an unbounded cache. Byte weights are accounting estimates, not a
promise about exact JVM heap use: executing renders, in-flight responses, caller
references, parser objects, and transient serialization buffers are not retained
cache entries and are outside this limit.

Example with both cache layers disabled:

```sh
java -Dpaperword.render.cache.enabled=false \
  -Dpaperword.render.cache.memory.maxBytes=0 \
  -jar target/paper-to-word-1.0.0.jar
```

## Disk ownership, durability, and safety

New records live only in `<dir>/paperword-formula-cache-v1/`. A newly created
namespace receives an exact ownership marker. An existing directory without the
marker is **not** adopted or cleaned. Old sharded `.png`, `.wmf`, and `.properties`
cache files outside the namespace are not read, migrated, or removed; operators
may separately remove a known legacy cache after verifying its ownership.

A record contains metadata and image bytes together, with a length check and a
SHA-256 checksum. An atomic same-directory move publishes the entire record;
readers never combine metadata from one generation with image bytes from
another. Damaged records are cache misses and are discarded before retrying the
render. Oversized files are rejected before allocating a corresponding read
buffer. Cache errors are non-fatal to rendering.

Every directory component is opened with `NOFOLLOW_LINKS` using Java
`SecureDirectoryStream`; file reads, creation, replacement, and deletion are
relative to that securely opened directory. Symlinked cache directories and
entries are not followed. Only strictly named `entry-<sha256>.pwc` records and
stale application `pending-<uuid>.tmp` files inside the owned namespace are
eligible for cleanup. Ordinary files, directories, and legacy cache files are
left alone. Use a directory writable only by the application account.

The default Linux filesystem provider supports these secure directory APIs.
On a filesystem/provider without them (including common Windows providers),
**disk caching is deliberately unavailable** and rendering continues using the
bounded memory cache. This is a disk-cache capability difference, not a claim
that the Linux-only PDF converter works on Windows. No less-safe deletion
fallback is used.
Disk-cache failures can be observed through the statistics below.

A nonblocking OS file lock serializes mutations across JVMs. A revision in the
lock file invalidates another process's cached index after a write/deletion; a
refresh with no mutations does not produce another revision. Contended locks
cause an ordinary cache miss or skipped write, not an unbounded wait. Size and
entry limits are enforced before a completed write is reported. Disk LRU order
is exact within a process; access times are persisted at most once per minute,
so eviction order after restart or another process's write is approximate.
TTL is enforced on every disk hit and during maintenance, including backward
clock movement (future-dated entries are misses). The disk TTL is storage
retention, not formula freshness: a still-retained memory result remains valid
until capacity eviction; render-affecting configuration is part of its key.

The quota covers completed record files. The ownership marker and eight-byte
lock are fixed overhead; an atomic write temporarily needs room for one extra
record, up to the single-record limit. A process crash may leave a staging file;
maintenance removes strictly named staging files after one hour. Maintenance is
activity-driven, not a background timer: an idle application's expired cache
files remain until another disk operation. Filesystem permission/I/O failures
can prevent eviction; these are reported as failures and the operation degrades
to an uncached render.

## Cost and observability

Memory hit/insert operations are synchronized, O(1) except for the number of
entries evicted by an insertion. Memory hits perform **no disk access**. Disk
hits read one bounded record and check its checksum. A directory scan occurs at
initialization, after another cache instance/JVM changes the revision, or after
the maintenance interval; ordinary same-process hot hits do not scan the tree.
The scan's retained index is itself bounded by the configured quota and entry
limit. There is no recursive traversal. Hash/serialization/disk overhead applies
only to disk-cache operations, not to memory hits.

`new LaTeXImageRenderer().cacheStatistics()` returns an immutable map:

- `memory.entries`, `memory.bytes`, `memory.maxEntries`, `memory.maxBytes`
- `memory.hits`, `memory.misses`, `memory.evictions`, `memory.rejected`, `memory.inFlight`
- `disk.enabled`, plus (when configured) `disk.entries`, `disk.bytes`, `disk.maxEntries`, `disk.maxBytes`
- `disk.hits`, `disk.misses`, `disk.writes`, `disk.evictions`, `disk.expired`, `disk.corruptions`, `disk.rejected`, `disk.failures`, `disk.scans`

These counters are process-local and cumulative for the current cache instance;
changing the disk configuration starts new disk counters. Disk size/count
statistics reflect the most recent locked observation; another process can
change disk contents before the next observation. The existing OLE
singleflight performs a second lookup after acquiring ownership, so hit/miss
counts describe cache lookups, not HTTP requests. The public snapshot does not
scan disk and does not expose formulas or keys.

## Regression coverage

`RenderCacheTest` uses only synthetic data, temporary directories, injectable
small limits, and an injectable clock. It checks combined memory weights,
entry/byte eviction, access ordering, concurrent reads/writes, expiry boundaries,
clock reversal, sharing across separate indexes, checksums/truncation, owned
namespace handling, symlinks, ordinary-file preservation, disabled retention,
geometry serialization, deterministic PNG/WMF recomputation after eviction,
initial quota reduction, and selective stale-staging cleanup. Disk-only tests
require a secure directory provider; memory and disabled-cache tests still run
elsewhere. A separate ZIP-filesystem test verifies that an unsupported provider
still returns a rendered PNG without writing a disk cache entry.
`LaTeXImageRendererTest` and `RenderConcurrencyRegressionTest` continue to cover
vector rendering, render-configuration cache identities, singleflight failures,
worker timeouts, and interruption.

HTTP observability: `GET /api/export/cache/status` returns these process-local counters without formula content or filesystem paths.

`disk.enabled` reports that the disk layer is configured with nonzero limits; it is not an I/O health guarantee. Check `disk.failures`, writes, and misses for permission/provider failures.
