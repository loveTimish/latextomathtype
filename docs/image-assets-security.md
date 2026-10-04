# Local image assets: explicit, bounded and fail-closed

Both paper and layout Word exports use `ImageAssetLoader`. This applies to section
images, question images, layout IMAGE blocks, and page-background image fallbacks.
Formula previews use a separate renderer and are not affected.

## Deployment configuration

Local images are **disabled by default**. A document without image references can
still export normally. A request with an image reference fails with
`IMAGE_ASSET_DISABLED` until an administrator explicitly grants an asset directory.
There is no working-directory/project-directory fallback.

```bash
java \
  -Dpaperword.assets.root=/srv/paperword/images \
  -Dpaperword.assets.max-bytes=16777216 \
  -Dpaperword.assets.max-pixels=16000000 \
  -jar target/paper-to-word-1.0.0.jar
```

| JVM property | Environment variable | Default |
| --- | --- | --- |
| `paperword.assets.root` | `PAPERWORD_ASSETS_ROOT` | none; image reads disabled |
| `paperword.assets.windows-prefix` | `PAPERWORD_ASSETS_WINDOWS_PREFIX` | none; no legacy mapping |
| `paperword.assets.max-bytes` | `PAPERWORD_ASSETS_MAX_BYTES` | 16,777,216 bytes per image |
| `paperword.assets.max-pixels` | `PAPERWORD_ASSETS_MAX_PIXELS` | 16,000,000 pixels per image |

JVM properties take precedence over environment variables, including an explicitly
blank root (which disables reads) or blank Windows prefix (which disables mapping).
Limits must be decimal positive integers no larger than 2,147,483,639. Invalid
configuration fails explicitly; it does not quietly use an unlimited/default value.
Choose substantially smaller limits for a small heap or highly concurrent service:
full decoding can require multiple bytes per pixel plus decoder and DOCX overhead.
These are per-image limits, not a total-document or total-concurrent-export budget.

The setting source is deliberately JVM properties/environment variables, **not
Spring application.yml binding**. The Spring export services, CLI tools and direct
`new DocxBuilder(...)` / `new LayoutDocxBuilder(...)` calls all use the same settings.
To provide policy programmatically, use:

```java
ImageAssetLoader assets = new ImageAssetLoader(new ImageAssetConfig(
    Path.of("/srv/paperword/images"), null, 16 * 1024 * 1024L, 16_000_000L));
new DocxBuilder(true, assets).build(paperRequest);
new LayoutDocxBuilder(true, assets).build(layoutRequest);
```

Do not expose the loader configuration as request parameters. Choose the narrowest
asset directory that contains only images intended for export. Do not grant `/`,
a home directory, the whole application checkout, logs, keys, or arbitrary uploads.

## Exact path rules

- Relative paths are relative to the configured asset root, never `user.dir`
- Absolute native paths must be contained by that root, both lexically and after
  `toRealPath()` resolution; a sibling directory sharing the root's name prefix
  does not match
- Backslashes are treated as separators for portable request paths; no URL decoding
  or percent decoding takes place
- Any `..` path component is rejected, including paths that would normalize back
  into the root; NUL, URLs, drive-relative paths, UNC paths and alternate data
  streams are rejected
- Symlinks in an image path are rejected even if their target is inside the root;
  symlinks escaping the root are rejected by real-path containment first
- A configured root itself may be an operator-managed symlink: its canonical target
  is the allowed root for that read; image path components below it may not be links
- Directories, symbolic links, FIFO/named pipes, sockets and device files are not
  regular assets and are rejected before attempting to read their contents
- Missing, blank or invalid requested image paths abort the export. They never turn
  into a successful document with a silently missing image

### Legacy Windows paths on Linux

A single explicitly configured drive-qualified Windows directory may map to the
asset root. For example:

```bash
-Dpaperword.assets.root=/srv/paperword/images
'-Dpaperword.assets.windows-prefix=J:\latextomathtype\target\reference-roundtrip\generated-media'
```

`J:\latextomathtype\target\reference-roundtrip\generated-media\lesson\figure.png`
then maps exactly to `/srv/paperword/images/lesson/figure.png`. The prefix comparison
is case-insensitive and requires a directory-component boundary. The remainder
uses the Linux filesystem's case sensitivity. Another drive, a partial prefix,
UNC path, parent component or unmatched Windows prefix fails. There is no search
for `/target/`, `/rebuild-assets/`, `/latextomathtype/` or arbitrary matching suffixes.
Use a root-relative path to avoid legacy mapping entirely.

## Content and resource checks

- The file is read once into a bounded byte snapshot. That exact snapshot is both
  validated and embedded; no builder reopens the original filename
- Both pre-read file size and every read increment are checked against `max-bytes`
- Only actual PNG and common 8-bit grayscale/three-component JPEG are supported.
  GIF, BMP, TIFF, SVG, PDF and arbitrary bytes fail explicitly. Convert unsupported
  formats to PNG in a trusted ingestion step before export
- Filename extensions, declared MIME type, and request-provided width/height are
  not treated as proof of format or dimensions
- PNG IHDR / JPEG frame dimensions are inspected directly before calling ImageIO.
  Nonpositive dimensions and `width * height > max-pixels` are rejected using long
  arithmetic before allocating a decoded bitmap
- PNG chunk bounds/checksums and the complete IEND are checked. Obvious JPEG
  truncation/malformed headers are rejected. The built-in JDK PNG/JPEG reader then
  performs a full decode; decoder errors or warnings abort export
- ImageIO reads the in-memory snapshot with metadata ignored and without its disk
  cache. Only the supported JDK reader implementations are used
- Image display sizes fit both page axes before conversion to Word EMU units;
  extreme aspect ratios do not produce overflow-sized drawing dimensions

This is raster validation, not a general-purpose file sanitizer. Source image bytes
(including valid metadata) remain in the exported DOCX; only place exportable image
files inside the approved asset root.

## Filesystem trust and race boundary

This implementation targets Linux filesystems whose Java provider supports
`SecureDirectoryStream`. It opens directory components from the filesystem root
with `NOFOLLOW_LINKS`, uses descriptor-relative file opens with `NOFOLLOW_LINKS`,
rejects non-regular files, and compares file identity, size and modification time
before/after the bounded read. Providers without secure directory streams or stable
file identities fail closed. There is intentionally no weaker Windows/provider
fallback; a future native Windows safe-handle reader must provide an equivalent
policy before enabling filesystem reads there.

The allowed root, its contents and relevant directory ancestors must be managed by
trusted operators. **Do not let requesters or other untrusted processes create,
replace, hard-link or concurrently rewrite entries in that tree.** A read-only
asset mount, with service-account access limited to reads, is the recommended
production arrangement. Publish updates outside active exports or atomically swap
trusted immutable asset versions.

Java's standard file API cannot atomically test an opened file descriptor's regular
file type and open it with `O_NONBLOCK`. Therefore pre/post checks do **not** claim
protection against a malicious local writer swapping a checked regular file for a
FIFO just before open, or same-inode/ABA mutations engineered between observations.
A stable FIFO is rejected without opening it. This trust boundary is explicit;
applications needing hostile concurrent upload directories need a separate trusted
image-ingestion/snapshot service or a native safe-file reader. No application-level
allowlist can defend against an attacker with the service process's own privileges.

## Errors and migration

`ImageAssetException` is an unchecked exception with a stable `getCode()` string
and a readable message. Builders and services propagate it instead of swallowing
it. Codes include `IMAGE_ASSET_DISABLED`, `CONFIG`, `INVALID_PATH`, `OUTSIDE_ROOT`,
`NOT_FOUND`, `NOT_REGULAR`, `CHANGED`, `TOO_LARGE`, `TOO_MANY_PIXELS`,
`UNSUPPORTED_FORMAT`, `INVALID_CONTENT`, `IO`, and `EMBED_FAILED` (all with the
`IMAGE_ASSET_` prefix). API handlers should map capacity errors to 413 and invalid
request assets to 400, while distinguishing operator configuration/server failures.

Existing callers must explicitly configure the narrow asset root and migrate old
paths to root-relative names or one exact legacy prefix. Previously successful
exports which silently omitted missing images will now fail visibly.

## Focused regression suite

```bash
./.mvn/apache-maven-3.9.12/bin/mvn \
  -Dtest=ImageAssetLoaderTest,ImageAssetExportTest,DocxBuilderTest,LayoutDocxBuilderTest test
```

Tests create only synthetic raster images and canaries in temporary directories.
They cover fail-closed defaults, safe paths, sibling-prefix escapes, parent
components, Windows prefix boundaries, outside/inside symlinks, directory/FIFO
rejection, real-content format detection, byte/pixel boundaries, huge dimensions,
corrupt/truncated data, immutable snapshots, both image builders, background
fallbacks and matching policy through export services. The FIFO and symlink checks
are Linux-specific. No sensitive host files, external documents or user machines
are read.
