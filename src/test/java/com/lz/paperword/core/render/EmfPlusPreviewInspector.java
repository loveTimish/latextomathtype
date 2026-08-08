package com.lz.paperword.core.render;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Path2D;
import java.awt.image.BufferedImage;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/** Independent test-side decoder for the generated EMF+ path stream. */
final class EmfPlusPreviewInspector {

    private static final int EMR_HEADER = 0x00000001;
    private static final int EMR_EOF = 0x0000000E;
    private static final int EMR_GDICOMMENT = 0x00000046;
    private static final int ENHMETA_SIGNATURE = 0x464D4520;
    private static final int EMF_PLUS_SIGNATURE = 0x2B464D45;
    private static final int EMF_PLUS_HEADER = 0x4001;
    private static final int EMF_PLUS_EOF = 0x4002;
    private static final int EMF_PLUS_OBJECT = 0x4008;
    private static final int EMF_PLUS_FILL_PATH = 0x4014;
    private static final int EMF_PLUS_SET_ANTI_ALIAS_MODE = 0x401E;
    private static final int EMF_PLUS_SET_PIXEL_OFFSET_MODE = 0x4022;
    private static final int GRAPHICS_VERSION_1_1 = 0xDBC01002;
    private static final int OBJECT_TYPE_PATH = 0x03;
    private static final int SOLID_COLOR_FLAG = 0x8000;
    private static final int LOGICAL_DPI = 96;

    private EmfPlusPreviewInspector() {
    }

    static RasterizedEmfPlus rasterize(byte[] emf, double widthPt, double heightPt, int dpi) {
        if (emf == null || emf.length < 108 || dpi <= 0
                || !Double.isFinite(widthPt) || !Double.isFinite(heightPt)
                || widthPt <= 0d || heightPt <= 0d) {
            throw new IllegalArgumentException("invalid EMF+ raster input");
        }
        if (u32(emf, 0) != EMR_HEADER || u32(emf, 40) != ENHMETA_SIGNATURE) {
            throw new IllegalArgumentException("not an enhanced metafile");
        }
        if (u32(emf, 48) != emf.length) {
            throw new IllegalArgumentException("EMF nBytes does not match payload length");
        }

        int widthPx = Math.max(1, (int) Math.ceil(widthPt * dpi / 72d));
        int heightPx = Math.max(1, (int) Math.ceil(heightPt * dpi / 72d));
        List<PaintedPath> fills = decodePlusStream(emf);
        BufferedImage image = new BufferedImage(widthPx, heightPx, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics = image.createGraphics();
        try {
            graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                RenderingHints.VALUE_ANTIALIAS_ON);
            graphics.setRenderingHint(RenderingHints.KEY_RENDERING,
                RenderingHints.VALUE_RENDER_QUALITY);
            graphics.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL,
                RenderingHints.VALUE_STROKE_PURE);
            AffineTransform toPixels = AffineTransform.getScaleInstance(
                dpi / (double) LOGICAL_DPI, dpi / (double) LOGICAL_DPI);
            for (PaintedPath fill : fills) {
                graphics.setColor(fill.color());
                graphics.fill(toPixels.createTransformedShape(fill.path()));
            }
        } finally {
            graphics.dispose();
        }
        Set<Integer> colors = new LinkedHashSet<>();
        fills.forEach(fill -> colors.add(fill.color().getRGB()));
        return new RasterizedEmfPlus(image, fills.size(), Set.copyOf(colors));
    }

    private static List<PaintedPath> decodePlusStream(byte[] emf) {
        int headerSize = checkedRecordSize(emf, 0, emf.length);
        if (headerSize < 88) {
            throw new IllegalArgumentException("truncated EMF header");
        }
        int offset = headerSize;
        int recordCount = 1;
        boolean sawClassicEof = false;
        boolean sawPlusHeader = false;
        boolean sawAntiAlias = false;
        boolean sawPixelOffset = false;
        boolean sawPlusEof = false;
        Map<Integer, Shape> objects = new HashMap<>();
        List<PaintedPath> fills = new ArrayList<>();
        while (offset < emf.length) {
            int type = u32(emf, offset);
            int size = checkedRecordSize(emf, offset, emf.length);
            recordCount++;
            if (type == EMR_GDICOMMENT) {
                int dataSize = u32(emf, offset + 8);
                if (dataSize < 4 || dataSize > size - 12
                        || u32(emf, offset + 12) != EMF_PLUS_SIGNATURE) {
                    throw new IllegalArgumentException("invalid EMF+ comment at " + offset);
                }
                int plusOffset = offset + 16;
                int plusEnd = offset + 12 + dataSize;
                while (plusOffset < plusEnd) {
                    int plusType = u16(emf, plusOffset);
                    int flags = u16(emf, plusOffset + 2);
                    int plusSize = u32(emf, plusOffset + 4);
                    int plusDataSize = u32(emf, plusOffset + 8);
                    if (plusSize < 12 || (plusSize & 3) != 0
                            || plusDataSize != plusSize - 12
                            || plusSize > plusEnd - plusOffset) {
                        throw new IllegalArgumentException(
                            "invalid EMF+ record at " + plusOffset);
                    }
                    switch (plusType) {
                        case EMF_PLUS_HEADER -> {
                            if (sawPlusHeader || plusDataSize != 16
                                    || (flags & 1) == 0
                                    || u32(emf, plusOffset + 12) != GRAPHICS_VERSION_1_1
                                    || u32(emf, plusOffset + 20) != LOGICAL_DPI
                                    || u32(emf, plusOffset + 24) != LOGICAL_DPI) {
                                throw new IllegalArgumentException("invalid EMF+ Dual header");
                            }
                            sawPlusHeader = true;
                        }
                        case EMF_PLUS_SET_ANTI_ALIAS_MODE -> {
                            requirePhase(sawPlusHeader && !sawPlusEof,
                                "anti-alias record outside EMF+ stream");
                            if (plusDataSize != 0 || flags != 0x000B) {
                                throw new IllegalArgumentException("unexpected EMF+ smoothing mode");
                            }
                            sawAntiAlias = true;
                        }
                        case EMF_PLUS_SET_PIXEL_OFFSET_MODE -> {
                            requirePhase(sawAntiAlias && !sawPlusEof,
                                "pixel-offset record outside EMF+ stream");
                            if (plusDataSize != 0 || flags != 0x0002) {
                                throw new IllegalArgumentException("unexpected EMF+ pixel-offset mode");
                            }
                            sawPixelOffset = true;
                        }
                        case EMF_PLUS_OBJECT -> {
                            requirePhase(sawPixelOffset && !sawPlusEof,
                                "path object outside EMF+ stream");
                            int objectType = (flags >>> 8) & 0x7F;
                            int objectId = flags & 0xFF;
                            if (objectType != OBJECT_TYPE_PATH || (flags & 0x8000) != 0) {
                                throw new IllegalArgumentException("unsupported EMF+ object type");
                            }
                            objects.put(objectId, decodePath(emf, plusOffset, plusDataSize));
                        }
                        case EMF_PLUS_FILL_PATH -> {
                            requirePhase(sawPixelOffset && !sawPlusEof,
                                "fill path outside EMF+ stream");
                            int objectId = flags & 0xFF;
                            if ((flags & SOLID_COLOR_FLAG) == 0
                                    || (flags & 0x7F00) != 0 || plusDataSize != 4) {
                                throw new IllegalArgumentException("unsupported EMF+ fill path");
                            }
                            Shape path = objects.get(objectId);
                            if (path == null) {
                                throw new IllegalArgumentException(
                                    "EMF+ fill references missing path object " + objectId);
                            }
                            fills.add(new PaintedPath(new Path2D.Double(path),
                                new Color(u32(emf, plusOffset + 12), true)));
                        }
                        case EMF_PLUS_EOF -> {
                            if (!sawPixelOffset || sawPlusEof || plusDataSize != 0) {
                                throw new IllegalArgumentException("invalid EMF+ EOF");
                            }
                            sawPlusEof = true;
                        }
                        default -> throw new IllegalArgumentException(String.format(
                            "unsupported EMF+ record 0x%04X", plusType));
                    }
                    plusOffset += plusSize;
                }
                if (plusOffset != plusEnd) {
                    throw new IllegalArgumentException("misaligned EMF+ comment payload");
                }
            } else if (type == EMR_EOF) {
                if (offset + size != emf.length) {
                    throw new IllegalArgumentException("classic EMF EOF is not last");
                }
                sawClassicEof = true;
            }
            offset += size;
        }
        if (offset != emf.length || recordCount != u32(emf, 52)
                || !sawPlusHeader || !sawAntiAlias || !sawPixelOffset
                || !sawPlusEof || !sawClassicEof) {
            throw new IllegalArgumentException("incomplete EMF+ Dual stream");
        }
        return List.copyOf(fills);
    }

    private static Shape decodePath(byte[] emf, int offset, int dataSize) {
        if (dataSize < 12 || u32(emf, offset + 12) != GRAPHICS_VERSION_1_1) {
            throw new IllegalArgumentException("invalid EMF+ path header");
        }
        int pointCount = u32(emf, offset + 16);
        int pointFlags = u32(emf, offset + 20);
        if (pointCount < 0 || pointFlags != 0) {
            throw new IllegalArgumentException("unsupported EMF+ path point encoding");
        }
        long required = 12L + pointCount * 9L;
        if (required > dataSize) {
            throw new IllegalArgumentException("truncated EMF+ path data");
        }
        int coordinates = offset + 24;
        int types = coordinates + pointCount * 8;
        float[] xs = new float[pointCount];
        float[] ys = new float[pointCount];
        for (int index = 0; index < pointCount; index++) {
            xs[index] = f32(emf, coordinates + index * 8);
            ys[index] = f32(emf, coordinates + index * 8 + 4);
            if (!Float.isFinite(xs[index]) || !Float.isFinite(ys[index])) {
                throw new IllegalArgumentException("non-finite EMF+ path coordinate");
            }
        }

        Path2D.Double path = new Path2D.Double(Path2D.WIND_EVEN_ODD);
        boolean figure = false;
        int index = 0;
        while (index < pointCount) {
            int pointType = emf[types + index] & 0xFF;
            int kind = pointType & 0x0F;
            boolean close = (pointType & 0x80) != 0;
            if ((pointType & 0x70) != 0) {
                throw new IllegalArgumentException("unsupported EMF+ path point flags");
            }
            if (kind == 0) {
                if (close) {
                    throw new IllegalArgumentException("empty EMF+ figure cannot close");
                }
                path.moveTo(xs[index], ys[index]);
                figure = true;
                index++;
            } else if (kind == 1) {
                requirePhase(figure, "line point before path start");
                path.lineTo(xs[index], ys[index]);
                if (close) {
                    path.closePath();
                    figure = false;
                }
                index++;
            } else if (kind == 3) {
                requirePhase(figure && index + 2 < pointCount,
                    "incomplete Bezier sequence");
                int secondType = emf[types + index + 1] & 0xFF;
                int thirdType = emf[types + index + 2] & 0xFF;
                if ((secondType & 0x0F) != 3 || (thirdType & 0x0F) != 3
                        || (pointType & 0x80) != 0 || (secondType & 0x80) != 0) {
                    throw new IllegalArgumentException("invalid EMF+ Bezier point sequence");
                }
                path.curveTo(xs[index], ys[index], xs[index + 1], ys[index + 1],
                    xs[index + 2], ys[index + 2]);
                if ((thirdType & 0x80) != 0) {
                    path.closePath();
                    figure = false;
                }
                index += 3;
            } else {
                throw new IllegalArgumentException("unsupported EMF+ path point type " + kind);
            }
        }
        return path;
    }

    private static int checkedRecordSize(byte[] bytes, int offset, int limit) {
        if (offset < 0 || offset > limit - 8) {
            throw new IllegalArgumentException("truncated EMF record header at " + offset);
        }
        int size = u32(bytes, offset + 4);
        if (size < 8 || (size & 3) != 0 || size > limit - offset) {
            throw new IllegalArgumentException("invalid EMF record size at " + offset);
        }
        return size;
    }

    private static void requirePhase(boolean condition, String message) {
        if (!condition) {
            throw new IllegalArgumentException(message);
        }
    }

    private static int u16(byte[] bytes, int offset) {
        requireRange(bytes, offset, 2);
        return (bytes[offset] & 0xFF) | ((bytes[offset + 1] & 0xFF) << 8);
    }

    private static int u32(byte[] bytes, int offset) {
        requireRange(bytes, offset, 4);
        return (bytes[offset] & 0xFF)
            | ((bytes[offset + 1] & 0xFF) << 8)
            | ((bytes[offset + 2] & 0xFF) << 16)
            | (bytes[offset + 3] << 24);
    }

    private static float f32(byte[] bytes, int offset) {
        return Float.intBitsToFloat(u32(bytes, offset));
    }

    private static void requireRange(byte[] bytes, int offset, int length) {
        if (offset < 0 || length < 0 || offset > bytes.length - length) {
            throw new IllegalArgumentException("truncated binary field at " + offset);
        }
    }

    record RasterizedEmfPlus(BufferedImage image, int pathFillCount, Set<Integer> colors) {
        RasterizedEmfPlus {
            colors = Set.copyOf(colors);
        }
    }

    private record PaintedPath(Shape path, Color color) {
    }
}
