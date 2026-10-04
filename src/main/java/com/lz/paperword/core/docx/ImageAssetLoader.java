package com.lz.paperword.core.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.MemoryCacheImageInputStream;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.LinkOption;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.nio.file.StandardOpenOption;
import java.nio.file.attribute.BasicFileAttributeView;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Iterator;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.zip.CRC32;

/**
 * Reads only images in an explicitly allowed, trusted asset tree. Directory entries are
 * opened relative to pinned directory handles without following symbolic links. The same
 * bounded byte snapshot is validated and embedded, never reopening a request pathname.
 *
 * <p>The root and its contents must be managed by trusted operators, without untrusted
 * concurrent filesystem writers. Standard Java cannot atomically fstat/open a regular
 * file with O_NONBLOCK; see docs/image-assets-security.md for this deployment boundary.</p>
 */
public final class ImageAssetLoader {
    private static final byte[] PNG_SIGNATURE = {(byte) 137, 80, 78, 71, 13, 10, 26, 10};
    private final ImageAssetConfig config;

    public ImageAssetLoader(ImageAssetConfig config) {
        this.config = Objects.requireNonNull(config, "config");
    }

    public static ImageAssetLoader fromSystemProperties() {
        return new ImageAssetLoader(ImageAssetConfig.fromSystemProperties());
    }

    public ImageAsset load(String reference) {
        if (reference == null || reference.isBlank()) {
            throw failure("INVALID_PATH", "An image was requested without a file path");
        }
        if (config.root() == null) {
            throw failure("DISABLED", "Local images are disabled; configure paperword.assets.root to an approved image directory");
        }
        try {
            Path root;
            try {
                root = config.root().toRealPath();
                if (!Files.isDirectory(root, LinkOption.NOFOLLOW_LINKS)) {
                    throw failure("CONFIG", "paperword.assets.root must be an existing directory");
                }
            } catch (IOException e) {
                throw failure("CONFIG", "The configured image asset directory is unavailable", e);
            }
            Path candidate = resolve(root, reference);
            // Real-path containment is necessary even when a pathname is lexically inside root.
            Path real = candidate.toRealPath();
            if (!real.startsWith(root)) {
                throw failure("OUTSIDE_ROOT", "Image path resolves outside the approved image directory");
            }
            byte[] bytes = readBounded(candidate);
            return validate(bytes);
        } catch (NoSuchFileException e) {
            throw failure("NOT_FOUND", "A requested image file does not exist in the approved image directory", e);
        } catch (InvalidPathException e) {
            throw failure("INVALID_PATH", "The requested image path is invalid", e);
        } catch (IOException | SecurityException e) {
            throw failure("IO", "The requested image could not be read safely from the approved image directory", e);
        }
    }

    private Path resolve(Path root, String reference) {
        String value = reference.replace('\\', '/');
        if (value.indexOf('\0') >= 0 || ImageAssetConfig.hasTraversal(value)) {
            throw failure("INVALID_PATH", "Image paths must not contain NUL or parent-directory (..) components");
        }
        if (value.matches("^[A-Za-z]:/.*")) {
            String prefix = config.windowsPrefix();
            if (prefix == null || value.length() <= prefix.length()
                || !value.regionMatches(true, 0, prefix, 0, prefix.length())
                || value.charAt(prefix.length()) != '/') {
                throw failure("OUTSIDE_ROOT", "Windows image paths require a matching explicit paperword.assets.windows-prefix");
            }
            value = value.substring(prefix.length() + 1);
            // A suffix beginning '/' could otherwise make Path.resolve discard its root.
            if (value.isEmpty() || value.startsWith("/")) {
                throw failure("INVALID_PATH", "The mapped Windows image path must name a file below its configured prefix");
            }
        } else if (value.startsWith("//")) {
            throw failure("INVALID_PATH", "UNC image paths are not supported; use an approved root-relative path");
        }
        if (value.indexOf(':') >= 0) {
            throw failure("INVALID_PATH", "Image paths cannot be URLs, drive-relative paths, or alternate data streams");
        }
        Path path = Path.of(value);
        Path candidate = (path.isAbsolute() ? path : root.resolve(path)).normalize();
        if (!candidate.startsWith(root)) {
            throw failure("OUTSIDE_ROOT", "Image path is outside the approved image directory");
        }
        if (candidate.equals(root)) {
            throw failure("NOT_REGULAR", "An image must be a regular file, not the asset directory");
        }
        return candidate;
    }

    private byte[] readBounded(Path candidate) throws IOException {
        // Start at the filesystem root and pin every ancestor; never follow any symlink,
        // including an in-root symlink. Unsupported filesystem providers fail closed.
        try (SecureDirectoryStream<Path> directory = openDirectory(candidate.getParent())) {
            Path name = candidate.getFileName();
            BasicFileAttributes before = attributes(directory, name);
            if (!before.isRegularFile() || before.isSymbolicLink()) {
                throw failure("NOT_REGULAR", "A requested image is not a regular file; symlinks, devices and FIFOs are prohibited");
            }
            requireSize(before.size());
            if (before.fileKey() == null) {
                throw failure("CONFIG", "The image filesystem must expose stable file identities");
            }
            try (SeekableByteChannel channel = directory.newByteChannel(name,
                    Set.of(StandardOpenOption.READ, LinkOption.NOFOLLOW_LINKS))) {
                requireSize(channel.size());
                ByteArrayOutputStream output = new ByteArrayOutputStream((int) Math.min(before.size(), 8192));
                ByteBuffer chunk = ByteBuffer.allocate(8192);
                long total = 0;
                int read;
                while ((read = channel.read(chunk)) != -1) {
                    if (read == 0) {
                        continue;
                    }
                    total += read;
                    requireSize(total);
                    output.write(chunk.array(), 0, read);
                    chunk.clear();
                }
                BasicFileAttributes after = attributes(directory, name);
                if (!after.isRegularFile() || !Objects.equals(before.fileKey(), after.fileKey())
                    || before.size() != after.size() || before.size() != total
                    || channel.size() != total || !before.lastModifiedTime().equals(after.lastModifiedTime())) {
                    throw failure("CHANGED", "The requested image changed while being read; retry after asset updates finish");
                }
                return output.toByteArray();
            }
        }
    }

    private SecureDirectoryStream<Path> openDirectory(Path absolute) throws IOException {
        DirectoryStream<Path> start = Files.newDirectoryStream(absolute.getRoot());
        if (!(start instanceof SecureDirectoryStream<Path> secure)) {
            start.close();
            throw failure("CONFIG", "Secure local image reads require a filesystem with SecureDirectoryStream support");
        }
        try {
            for (Path component : absolute) {
                BasicFileAttributes attrs = attributes(secure, component);
                if (!attrs.isDirectory() || attrs.isSymbolicLink()) {
                    throw failure("NOT_REGULAR", "Image directory components must be real directories; symbolic links are prohibited");
                }
                SecureDirectoryStream<Path> next = secure.newDirectoryStream(component, LinkOption.NOFOLLOW_LINKS);
                secure.close();
                secure = next;
            }
            return secure;
        } catch (Throwable error) {
            secure.close();
            throw error;
        }
    }

    private BasicFileAttributes attributes(SecureDirectoryStream<Path> directory, Path path) throws IOException {
        return directory.getFileAttributeView(path, BasicFileAttributeView.class, LinkOption.NOFOLLOW_LINKS).readAttributes();
    }

    private void requireSize(long size) {
        if (size > config.maxBytes()) {
            throw failure("TOO_LARGE", "Image exceeds the configured byte limit of " + config.maxBytes());
        }
    }

    private ImageAsset validate(byte[] bytes) {
        ImageHeader header = header(bytes);
        checkPixels(header.width(), header.height());
        // Select only the JDK readers for the two supported raster formats. The explicit
        // header checks run before invoking ImageIO, so huge dimensions cannot reach read().
        ImageReader reader = null;
        Iterator<ImageReader> readers = ImageIO.getImageReadersByFormatName(header.format());
        String readerClass = "com.sun.imageio.plugins." + header.format() + "."
            + (header.format().equals("png") ? "PNG" : "JPEG") + "ImageReader";
        while (readers.hasNext()) {
            ImageReader candidate = readers.next();
            if (candidate.getClass().getName().equals(readerClass)) {
                reader = candidate;
                break;
            }
            candidate.dispose();
        }
        if (reader == null) {
            throw failure("CONFIG", "A required JDK image decoder is unavailable");
        }
        AtomicBoolean warning = new AtomicBoolean();
        reader.addIIOReadWarningListener((source, message) -> warning.set(true));
        try (MemoryCacheImageInputStream input = new MemoryCacheImageInputStream(new ByteArrayInputStream(bytes))) {
            reader.setInput(input, false, true);
            if (reader.getWidth(0) != header.width() || reader.getHeight(0) != header.height()) {
                throw failure("INVALID_CONTENT", "Image dimensions are inconsistent with its content");
            }
            BufferedImage decoded = reader.read(0);
            try {
                if (decoded == null || decoded.getWidth() != header.width() || decoded.getHeight() != header.height()
                    || warning.get()) {
                    throw failure("INVALID_CONTENT", "Image data is incomplete or corrupt");
                }
            } finally {
                if (decoded != null) {
                    decoded.flush();
                }
            }
            return new ImageAsset(bytes, header.width(), header.height(), header.format().equals("png")
                ? XWPFDocument.PICTURE_TYPE_PNG : XWPFDocument.PICTURE_TYPE_JPEG,
                "image." + header.format(), "image/" + header.format());
        } catch (IOException | RuntimeException e) {
            if (e instanceof ImageAssetException imageError) {
                throw imageError;
            }
            throw failure("INVALID_CONTENT", "Image data cannot be fully decoded as a supported PNG or JPEG", e);
        } finally {
            reader.dispose();
        }
    }

    private ImageHeader header(byte[] bytes) {
        if (bytes.length >= PNG_SIGNATURE.length && java.util.Arrays.equals(PNG_SIGNATURE,
                java.util.Arrays.copyOf(bytes, PNG_SIGNATURE.length))) {
            return pngHeader(bytes);
        }
        if (bytes.length >= 2 && u8(bytes, 0) == 0xff && u8(bytes, 1) == 0xd8) {
            return jpegHeader(bytes);
        }
        throw failure("UNSUPPORTED_FORMAT", "Image content must be PNG or JPEG; a filename extension or declared MIME type is not proof of format");
    }

    private ImageHeader pngHeader(byte[] bytes) {
        if (bytes.length < 33 || u32(bytes, 8) != 13 || u32(bytes, 12) != 0x49484452L) {
            throw failure("INVALID_CONTENT", "PNG must start with a complete IHDR chunk");
        }
        int width = (int) u32(bytes, 16);
        int height = (int) u32(bytes, 20);
        checkPixels(width, height);
        boolean idat = false;
        for (int pos = 8; pos <= bytes.length - 12;) {
            long length = u32(bytes, pos);
            if (length > bytes.length - pos - 12L) {
                throw failure("INVALID_CONTENT", "PNG contains a truncated or invalid chunk");
            }
            int size = (int) length;
            long type = u32(bytes, pos + 4);
            CRC32 crc = new CRC32();
            crc.update(bytes, pos + 4, size + 4);
            if (crc.getValue() != u32(bytes, pos + 8 + size)) {
                throw failure("INVALID_CONTENT", "PNG chunk checksum failed");
            }
            if (type == 0x49444154L) {
                idat = true;
            }
            pos += size + 12;
            if (type == 0x49454e44L) {
                if (size != 0 || !idat || pos != bytes.length) {
                    throw failure("INVALID_CONTENT", "PNG is incomplete or contains data after IEND");
                }
                return new ImageHeader("png", width, height);
            }
        }
        throw failure("INVALID_CONTENT", "PNG is missing its complete IEND chunk");
    }

    private ImageHeader jpegHeader(byte[] bytes) {
        if (bytes.length < 4 || u8(bytes, bytes.length - 2) != 0xff || u8(bytes, bytes.length - 1) != 0xd9) {
            throw failure("INVALID_CONTENT", "JPEG is truncated or contains data after its end marker");
        }
        int position = 2;
        while (position < bytes.length - 2) {
            if (u8(bytes, position++) != 0xff) {
                throw failure("INVALID_CONTENT", "JPEG contains an invalid header marker");
            }
            while (position < bytes.length && u8(bytes, position) == 0xff) {
                position++;
            }
            if (position >= bytes.length) {
                break;
            }
            int marker = u8(bytes, position++);
            if (marker == 0xda || marker == 0xd9 || marker == 0x00) {
                break;
            }
            if (marker == 0x01 || marker >= 0xd0 && marker <= 0xd7) {
                continue;
            }
            if (position + 2 > bytes.length) {
                break;
            }
            int length = u16(bytes, position);
            if (length < 2 || length > bytes.length - position) {
                break;
            }
            if (marker == 0xc0 || marker == 0xc1 || marker == 0xc2) {
                if (length < 8 || u8(bytes, position + 2) != 8) {
                    throw failure("INVALID_CONTENT", "Only complete 8-bit JPEG frames are supported");
                }
                int components = u8(bytes, position + 7);
                if ((components != 1 && components != 3) || length != 8 + 3 * components) {
                    throw failure("UNSUPPORTED_FORMAT", "Only grayscale or three-component JPEG images are supported");
                }
                return new ImageHeader("jpeg", u16(bytes, position + 5), u16(bytes, position + 3));
            }
            position += length;
        }
        throw failure("INVALID_CONTENT", "JPEG has no valid supported frame header");
    }

    private void checkPixels(int width, int height) {
        if (width <= 0 || height <= 0) {
            throw failure("INVALID_CONTENT", "Image width and height must be positive");
        }
        if ((long) width * height > config.maxPixels()) {
            throw failure("TOO_MANY_PIXELS", "Image exceeds the configured pixel limit of " + config.maxPixels());
        }
    }

    private static int u8(byte[] bytes, int at) {
        return bytes[at] & 0xff;
    }

    private static int u16(byte[] bytes, int at) {
        return u8(bytes, at) << 8 | u8(bytes, at + 1);
    }

    private static long u32(byte[] bytes, int at) {
        return (long) u8(bytes, at) << 24 | (long) u8(bytes, at + 1) << 16
            | (long) u8(bytes, at + 2) << 8 | u8(bytes, at + 3);
    }

    private static ImageAssetException failure(String suffix, String message) {
        return new ImageAssetException("IMAGE_ASSET_" + suffix, message);
    }

    private static ImageAssetException failure(String suffix, String message, Throwable cause) {
        return new ImageAssetException("IMAGE_ASSET_" + suffix, message, cause);
    }

    private record ImageHeader(String format, int width, int height) { }

    /** One bounded snapshot; builders must embed these bytes rather than reopen the path. */
    public record ImageAsset(byte[] data, int widthPx, int heightPx, int pictureType,
                             String fileName, String contentType) { }
}
