package com.lz.paperword.core.docx;

import java.nio.file.InvalidPathException;
import java.nio.file.Path;

/**
 * Explicit server-side image policy shared by Spring services and direct/CLI builders.
 * An absent root disables all requested local images; it never grants access to user.dir.
 */
public record ImageAssetConfig(Path root, String windowsPrefix, long maxBytes, long maxPixels) {
    public static final long DEFAULT_MAX_BYTES = 16L * 1024 * 1024;
    public static final long DEFAULT_MAX_PIXELS = 16_000_000L;

    public ImageAssetConfig {
        if (root != null && !root.isAbsolute()) {
            throw configError("paperword.assets.root must be an absolute directory path");
        }
        if (maxBytes <= 0 || maxBytes > Integer.MAX_VALUE - 8L) {
            throw configError("paperword.assets.max-bytes must be between 1 and 2147483639");
        }
        if (maxPixels <= 0 || maxPixels > Integer.MAX_VALUE - 8L) {
            throw configError("paperword.assets.max-pixels must be between 1 and 2147483639");
        }
        if (windowsPrefix != null && !windowsPrefix.isBlank()) {
            windowsPrefix = windowsPrefix.replace('\\', '/');
            while (windowsPrefix.endsWith("/")) {
                windowsPrefix = windowsPrefix.substring(0, windowsPrefix.length() - 1);
            }
            if (!windowsPrefix.matches("[A-Za-z]:/[^:]+") || hasTraversal(windowsPrefix)
                || windowsPrefix.contains("//") || windowsPrefix.indexOf('\0') >= 0) {
                throw configError("paperword.assets.windows-prefix must name one drive-qualified Windows directory");
            }
            if (root == null) {
                throw configError("paperword.assets.windows-prefix requires paperword.assets.root");
            }
        } else {
            windowsPrefix = null;
        }
    }

    public ImageAssetConfig(Path root) {
        this(root, null, DEFAULT_MAX_BYTES, DEFAULT_MAX_PIXELS);
    }

    /** JVM properties take precedence over matching uppercase underscore environment variables. */
    public static ImageAssetConfig fromSystemProperties() {
        String root = setting("root");
        try {
            return new ImageAssetConfig(root == null || root.isBlank() ? null : Path.of(root),
                setting("windows-prefix"), positiveLong("max-bytes", DEFAULT_MAX_BYTES),
                positiveLong("max-pixels", DEFAULT_MAX_PIXELS));
        } catch (InvalidPathException e) {
            throw new ImageAssetException("IMAGE_ASSET_CONFIG", "paperword.assets.root is not a valid directory path", e);
        }
    }

    private static long positiveLong(String name, long fallback) {
        String value = setting(name);
        if (value == null) {
            return fallback;
        }
        try {
            return Long.parseLong(value);
        } catch (NumberFormatException e) {
            throw configError("paperword.assets." + name + " must be a positive integer");
        }
    }

    private static String setting(String name) {
        String property = System.getProperty("paperword.assets." + name);
        return property != null ? property : System.getenv("PAPERWORD_ASSETS_" + name.toUpperCase(java.util.Locale.ROOT).replace('-', '_'));
    }

    static boolean hasTraversal(String path) {
        for (String component : path.split("/")) {
            if (component.equals("..")) {
                return true;
            }
        }
        return false;
    }

    private static ImageAssetException configError(String message) {
        return new ImageAssetException("IMAGE_ASSET_CONFIG", message);
    }
}
