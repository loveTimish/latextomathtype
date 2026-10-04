package com.lz.paperword.core.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Arrays;
import java.util.List;
import java.util.zip.CRC32;

import static org.junit.jupiter.api.Assertions.*;

class ImageAssetLoaderTest {
    @TempDir Path temporary;

    @Test
    void missingRootFailsClosedEvenForExistingValidImage() throws Exception {
        Path image = image(temporary.resolve("canary.png"), "png", 2, 3);
        assertCode("DISABLED", new ImageAssetLoader(new ImageAssetConfig(null)), image.toString());
    }

    @Test
    void acceptsOnlyPathsUnderExplicitRootWithoutWorkingDirectoryFallback() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path nested = Files.createDirectory(root.resolve("nested"));
        Path source = image(nested.resolve("canary.png"), "png", 3, 2);
        ImageAssetLoader loader = loader(root);
        for (String reference : List.of("nested/canary.png", "./nested/canary.png", "nested\\canary.png", source.toString())) {
            ImageAssetLoader.ImageAsset result = loader.load(reference);
            assertEquals(3, result.widthPx());
            assertEquals(2, result.heightPx());
            assertEquals(XWPFDocument.PICTURE_TYPE_PNG, result.pictureType());
            assertArrayEquals(Files.readAllBytes(source), result.data());
        }
        assertCode("NOT_FOUND", loader, "canary.png");
    }

    @Test
    void refusesOutsideCanaryAbsolutePathsAndPrefixLookalikes() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path outside = Files.createDirectory(temporary.resolve("assets-other"));
        Path canary = image(outside.resolve("canary.png"), "png", 1, 1);
        assertCode("OUTSIDE_ROOT", loader(root), canary.toString());
    }

    @Test
    void refusesAllParentComponentsBeforeNormalization() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        image(root.resolve("canary.png"), "png", 1, 1);
        for (String path : List.of("../canary.png", "nested/../canary.png", "nested\\..\\canary.png",
                root + "/../assets/canary.png", "J:\\old\\assets\\..\\canary.png")) {
            assertCode("INVALID_PATH", loader(root), path);
        }
    }

    @Test
    void refusesBlankNulUrlsUncAndDriveRelativePaths() throws Exception {
        ImageAssetLoader loader = loader(Files.createDirectory(temporary.resolve("assets")));
        for (String path : Arrays.asList(null, "", " ", "a\0.png", "file:///tmp/canary.png", "https://example.invalid/image.png",
                "//server/share/canary.png", "\\\\server\\share\\canary.png", "J:canary.png", "a.png:stream")) {
            assertCode("INVALID_PATH", loader, path);
        }
    }

    @Test
    void mapsOneExplicitWindowsDirectoryWithNoSuffixGuessing() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Files.createDirectory(root.resolve("nested"));
        image(root.resolve("nested/canary.png"), "png", 4, 3);
        ImageAssetLoader mapped = new ImageAssetLoader(new ImageAssetConfig(root, "J:\\legacy\\images\\",
            ImageAssetConfig.DEFAULT_MAX_BYTES, ImageAssetConfig.DEFAULT_MAX_PIXELS));
        assertEquals(4, mapped.load("j:\\LEGACY\\IMAGES\\nested\\canary.png").widthPx());
        assertEquals(4, mapped.load("J:/legacy/images/nested/canary.png").widthPx());
        for (String path : List.of("K:/legacy/images/nested/canary.png", "J:/legacy/images-other/nested/canary.png",
                "J:/other/target/nested/canary.png", "J:/latextomathtype/nested/canary.png")) {
            assertCode("OUTSIDE_ROOT", mapped, path);
        }
        assertCode("OUTSIDE_ROOT", loader(root), "J:/legacy/images/nested/canary.png");
        assertCode("INVALID_PATH", mapped, "J:/legacy/images//nested/canary.png");
    }

    @Test
    @EnabledOnOs(OS.LINUX)
    void refusesSymlinkEscapesAndInRootSymlinks() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path outside = Files.createDirectory(temporary.resolve("outside"));
        image(outside.resolve("canary.png"), "png", 1, 1);
        image(root.resolve("inside.png"), "png", 1, 1);
        Files.createSymbolicLink(root.resolve("escape-file.png"), outside.resolve("canary.png"));
        Files.createSymbolicLink(root.resolve("escape-dir"), outside);
        Files.createSymbolicLink(root.resolve("inside-link.png"), root.resolve("inside.png"));
        Files.createDirectory(root.resolve("nested"));
        image(root.resolve("nested/nested.png"), "png", 1, 1);
        Files.createSymbolicLink(root.resolve("nested-link"), root.resolve("nested"));
        ImageAssetLoader loader = loader(root);
        assertCode("OUTSIDE_ROOT", loader, "escape-file.png");
        assertCode("OUTSIDE_ROOT", loader, "escape-dir/canary.png");
        assertCode("NOT_REGULAR", loader, "inside-link.png");
        assertCode("NOT_REGULAR", loader, "nested-link/nested.png");
    }

    @Test
    @EnabledOnOs(OS.LINUX)
    void rejectsAFifoWithoutOpeningOrBlocking() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path fifo = root.resolve("canary-fifo.png");
        Process createFifo = new ProcessBuilder("mkfifo", fifo.toString()).start();
        assertEquals(0, createFifo.waitFor());
        assertTimeoutPreemptively(Duration.ofSeconds(2), () -> assertCode("NOT_REGULAR", loader(root), "canary-fifo.png"));
    }

    @Test
    @EnabledOnOs(OS.LINUX)
    void rejectsSyntheticUnixSocketWithoutOpeningItsContents() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path socket = root.resolve("canary.sock");
        java.nio.channels.ServerSocketChannel channel;
        try {
            channel = java.nio.channels.ServerSocketChannel.open(java.net.StandardProtocolFamily.UNIX);
        } catch (java.net.SocketException denied) {
            org.junit.jupiter.api.Assumptions.assumeFalse(
                "Operation not permitted".equals(denied.getMessage()) || "Permission denied".equals(denied.getMessage()),
                "This runtime forbids creating a synthetic Unix-domain socket");
            throw denied;
        }
        try (java.nio.channels.ServerSocketChannel server = channel) {
            server.bind(java.net.UnixDomainSocketAddress.of(socket));
            assertCode("NOT_REGULAR", loader(root), "canary.sock");
        }
    }

    @Test
    void rejectsDirectoriesAndMissingImages() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Files.createDirectory(root.resolve("folder.png"));
        ImageAssetLoader loader = loader(root);
        assertCode("NOT_REGULAR", loader, "folder.png");
        assertCode("NOT_REGULAR", loader, ".");
        assertCode("NOT_FOUND", loader, "missing.png");
    }

    @Test
    void detectsContentInsteadOfTrustingExtensionOrDeclaredType() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        image(root.resolve("actually-png.jpg"), "png", 3, 2);
        image(root.resolve("actually-jpeg.png"), "jpeg", 3, 2);
        Files.writeString(root.resolve("fake.png"), "SYNTHETIC-CANARY-NOT-AN-IMAGE");
        ImageAssetLoader loader = loader(root);
        assertEquals(XWPFDocument.PICTURE_TYPE_PNG, loader.load("actually-png.jpg").pictureType());
        assertEquals("image/png", loader.load("actually-png.jpg").contentType());
        assertEquals(XWPFDocument.PICTURE_TYPE_JPEG, loader.load("actually-jpeg.png").pictureType());
        assertEquals("image/jpeg", loader.load("actually-jpeg.png").contentType());
        assertCode("UNSUPPORTED_FORMAT", loader, "fake.png");
        Files.write(root.resolve("empty.png"), new byte[0]);
        assertCode("UNSUPPORTED_FORMAT", loader, "empty.png");
    }

    @Test
    void intentionallyRejectsOtherFormatsEvenWhenImageIoCanDecodeThem() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        for (String format : List.of("gif", "bmp")) {
            image(root.resolve("canary." + format), format, 3, 2);
            assertCode("UNSUPPORTED_FORMAT", loader(root), "canary." + format);
        }
    }

    @Test
    void enforcesByteAndPixelLimitsInclusively() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path image = image(root.resolve("canary.png"), "png", 3, 2);
        long size = Files.size(image);
        assertCode("TOO_LARGE", new ImageAssetLoader(new ImageAssetConfig(root, null, size - 1, 6)), "canary.png");
        assertCode("TOO_MANY_PIXELS", new ImageAssetLoader(new ImageAssetConfig(root, null, size, 5)), "canary.png");
        assertEquals(3, new ImageAssetLoader(new ImageAssetConfig(root, null, size, 6)).load("canary.png").widthPx());
    }

    @Test
    void rejectsHugePngDimensionsBeforeAnyImageIoDecode() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path image = image(root.resolve("huge.png"), "png", 1, 1);
        byte[] bytes = Files.readAllBytes(image);
        ByteBuffer.wrap(bytes).putInt(16, 1_000_000).putInt(20, 1_000_000);
        updatePngCrc(bytes, 8);
        Files.write(image, bytes);
        assertTimeoutPreemptively(Duration.ofSeconds(2), () -> assertCode("TOO_MANY_PIXELS", loader(root), "huge.png"));
    }

    @Test
    void rejectsHugeJpegDimensionsBeforeAnyImageIoDecode() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        byte[] header = {(byte) 0xff, (byte) 0xd8, (byte) 0xff, (byte) 0xc0, 0, 11, 8,
            (byte) 0xff, (byte) 0xff, (byte) 0xff, (byte) 0xff, 1, 1, 0x11, 0,
            (byte) 0xff, (byte) 0xd9};
        Files.write(root.resolve("huge.jpeg"), header);
        assertTimeoutPreemptively(Duration.ofSeconds(2), () -> assertCode("TOO_MANY_PIXELS", loader(root), "huge.jpeg"));
    }

    @Test
    void rejectsTruncatedChecksummedAndUndecodablePngContent() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path image = image(root.resolve("canary.png"), "png", 3, 2);
        byte[] valid = Files.readAllBytes(image);
        Files.write(root.resolve("truncated.png"), Arrays.copyOf(valid, valid.length - 5));
        byte[] corrupt = valid.clone();
        corrupt[29] ^= 1;
        Files.write(root.resolve("bad-checksum.png"), corrupt);
        byte[] undecodable = valid.clone();
        // Invalidate the first IDAT's deflate stream but recompute its checksum.
        int idat = chunkPosition(undecodable, 0x49444154);
        undecodable[idat + 8] = 0;
        updatePngCrc(undecodable, idat);
        Files.write(root.resolve("bad-zlib.png"), undecodable);
        byte[] trailing = Arrays.copyOf(valid, valid.length + 1);
        Files.write(root.resolve("trailing.png"), trailing);
        for (String path : List.of("truncated.png", "bad-checksum.png", "bad-zlib.png", "trailing.png")) {
            assertCode("INVALID_CONTENT", loader(root), path);
        }
    }

    @Test
    void rejectsTruncatedOrCorruptJpegContent() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path image = image(root.resolve("canary.jpeg"), "jpeg", 3, 2);
        byte[] valid = Files.readAllBytes(image);
        Files.write(root.resolve("truncated.jpeg"), Arrays.copyOf(valid, valid.length - 2));
        byte[] headerOnly = Arrays.copyOf(valid, 32);
        headerOnly[30] = (byte) 0xff;
        headerOnly[31] = (byte) 0xd9;
        Files.write(root.resolve("no-image.jpeg"), headerOnly);
        assertCode("INVALID_CONTENT", loader(root), "truncated.jpeg");
        assertCode("INVALID_CONTENT", loader(root), "no-image.jpeg");
    }

    @Test
    void validatesConfigInsteadOfSilentlyDefaultingInvalidLimits() {
        assertConfigError(() -> new ImageAssetConfig(Path.of("relative")));
        assertConfigError(() -> new ImageAssetConfig(temporary, null, 0, 1));
        assertConfigError(() -> new ImageAssetConfig(temporary, null, 1, -1));
        assertConfigError(() -> new ImageAssetConfig(temporary, null, Long.MAX_VALUE, 1));
        assertConfigError(() -> new ImageAssetConfig(temporary, null, 1, Long.MAX_VALUE));
        assertConfigError(() -> new ImageAssetConfig(temporary, "J:/legacy/../images", 100, 100));
        assertConfigError(() -> new ImageAssetConfig(temporary, "images", 100, 100));
        assertConfigError(() -> new ImageAssetConfig(null, "J:/legacy/images", 100, 100));
    }

    @Test
    void unavailableConfiguredRootHasReadableConfigurationError() {
        ImageAssetException error = assertThrows(ImageAssetException.class,
            () -> loader(temporary.resolve("missing-root")).load("canary.png"));
        assertEquals("IMAGE_ASSET_CONFIG", error.getCode());
        assertTrue(error.getMessage().contains("configured image asset directory"));
    }

    @Test
    void snapshotRemainsUnchangedWhenSourceFileIsLaterReplaced() throws Exception {
        Path root = Files.createDirectory(temporary.resolve("assets"));
        Path image = image(root.resolve("canary.png"), "png", 3, 2);
        byte[] original = Files.readAllBytes(image);
        ImageAssetLoader.ImageAsset result = loader(root).load("canary.png");
        Files.writeString(image, "REPLACED-SYNTHETIC-CANARY");
        assertArrayEquals(original, result.data());
    }

    private ImageAssetLoader loader(Path root) {
        return new ImageAssetLoader(new ImageAssetConfig(root));
    }

    static Path image(Path path, String format, int width, int height) throws IOException {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        image.setRGB(0, 0, 0xff339966);
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        assertTrue(ImageIO.write(image, format, output), "Synthetic format writer must be available");
        Files.write(path, output.toByteArray());
        image.flush();
        return path;
    }

    private static int chunkPosition(byte[] bytes, int chunkType) {
        ByteBuffer buffer = ByteBuffer.wrap(bytes);
        for (int pos = 8; pos < bytes.length - 12; pos += buffer.getInt(pos) + 12) {
            if (buffer.getInt(pos + 4) == chunkType) {
                return pos;
            }
        }
        throw new AssertionError("Missing chunk in synthetic PNG");
    }

    private static void updatePngCrc(byte[] bytes, int chunkPosition) {
        ByteBuffer buffer = ByteBuffer.wrap(bytes);
        int size = buffer.getInt(chunkPosition);
        CRC32 crc = new CRC32();
        crc.update(bytes, chunkPosition + 4, size + 4);
        buffer.putInt(chunkPosition + 8 + size, (int) crc.getValue());
    }

    private static void assertConfigError(org.junit.jupiter.api.function.Executable executable) {
        assertEquals("IMAGE_ASSET_CONFIG", assertThrows(ImageAssetException.class, executable).getCode());
    }

    static void assertCode(String suffix, ImageAssetLoader loader, String reference) {
        ImageAssetException error = assertThrows(ImageAssetException.class, () -> loader.load(reference));
        assertEquals("IMAGE_ASSET_" + suffix, error.getCode());
        assertFalse(error.getMessage().isBlank());
    }
}
