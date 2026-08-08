package com.lz.paperword.core.render;

import java.awt.Font;
import java.awt.GraphicsEnvironment;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;

/** Loads and verifies the fonts used to outline SVG text in strict vector previews. */
final class BundledVectorFonts {

    static final String FONT_SET_ID = "vector-preview-fonts-v1";
    static final String OVERRIDE_DIRECTORY_PROPERTY = "paperword.vector.font.dir";
    private static final String RESOURCE_ROOT = "/fonts/vector-preview/";
    private static final List<FontAsset> FONT_ASSETS = List.of(
        new FontAsset("DejaVuSerif.ttf", "42d1edeb7952f31b1f96d767ed7030b08a39e0c372b0071641518864e2bffb51"),
        new FontAsset("DejaVuSerif-Bold.ttf", "c47b5527bcdc8dcf9ea8c77054454c5a884beaca2f44851a2a823ee639cbf07f"),
        new FontAsset("DejaVuSerif-Italic.ttf", "2e39b1d50f90b933b00c7bb54a96afd3f86419b3d717c7cf202e36f2d4973e47"),
        new FontAsset("DejaVuSerif-BoldItalic.ttf", "8d3dd3d31350309042ed226af82b34539bd773518e6107cb352712853ba80308"),
        new FontAsset("DejaVuMathTeXGyre.ttf", "40da67c0b6b03076504fbda4bf3e7b4f20b35999c98063019a3392fe9b1294fe"),
        new FontAsset("NotoSerifSC-VF.ttf", "a4aed9985a5916fbf6690456f8732a9fccd517938e353165d4142b4f11a39280")
    );

    private static volatile List<Font> fonts;

    private BundledVectorFonts() {
    }

    static List<Font> ensureRegistered() throws IOException {
        List<Font> current = fonts;
        if (current != null) {
            return current;
        }
        synchronized (BundledVectorFonts.class) {
            if (fonts != null) {
                return fonts;
            }
            List<Font> loaded = new ArrayList<>(FONT_ASSETS.size());
            for (FontAsset asset : FONT_ASSETS) {
                byte[] bytes = readFont(asset.name());
                String actual = sha256(bytes);
                if (!asset.sha256().equals(actual)) {
                    throw new IOException("bundled vector font hash mismatch for " + asset.name()
                        + ": expected " + asset.sha256() + ", actual " + actual);
                }
                try {
                    Font font = Font.createFont(Font.TRUETYPE_FONT, new ByteArrayInputStream(bytes));
                    GraphicsEnvironment.getLocalGraphicsEnvironment().registerFont(font);
                    loaded.add(font);
                } catch (Exception exception) {
                    throw new IOException("cannot register bundled vector font " + asset.name(), exception);
                }
            }
            fonts = List.copyOf(loaded);
            return fonts;
        }
    }

    private static byte[] readFont(String name) throws IOException {
        String override = System.getProperty(OVERRIDE_DIRECTORY_PROPERTY, "").trim();
        if (!override.isEmpty()) {
            Path path = Path.of(override).toAbsolutePath().normalize().resolve(name).normalize();
            if (!path.getParent().equals(Path.of(override).toAbsolutePath().normalize())) {
                throw new IOException("invalid vector font override path: " + path);
            }
            return Files.readAllBytes(path);
        }
        try (InputStream input = BundledVectorFonts.class.getResourceAsStream(RESOURCE_ROOT + name)) {
            if (input == null) {
                throw new IOException("bundled vector font is missing: " + name);
            }
            return input.readAllBytes();
        }
    }

    private static String sha256(byte[] bytes) throws IOException {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes));
        } catch (NoSuchAlgorithmException impossible) {
            throw new IOException("SHA-256 is unavailable", impossible);
        }
    }

    private record FontAsset(String name, String sha256) {
    }
}
