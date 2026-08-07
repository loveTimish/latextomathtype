package com.lz.paperword.core.render;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Path2D;
import java.awt.image.BufferedImage;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/** Structural checks and independent rasterization for generated placeable WMF previews. */
public final class WmfPreviewInspector {

    private static final long PLACEABLE_KEY = 0x9AC6CDD7L;
    private static final int META_EOF = 0x0000;
    private static final int META_SET_MAP_MODE = 0x0103;
    private static final int META_SET_POLY_FILL_MODE = 0x0106;
    private static final int META_SELECT_OBJECT = 0x012D;
    private static final int META_DELETE_OBJECT = 0x01F0;
    private static final int META_SET_WINDOW_ORG = 0x020B;
    private static final int META_SET_WINDOW_EXT = 0x020C;
    private static final int META_CREATE_PEN = 0x02FA;
    private static final int META_CREATE_BRUSH = 0x02FC;
    private static final int META_POLYGON = 0x0324;
    private static final int META_POLYPOLYGON = 0x0538;
    private static final int META_TEXT_OUT = 0x0521;
    private static final int META_EXT_TEXT_OUT = 0x0A32;
    private static final Set<Integer> BITMAP_RECORDS = Set.of(0x0F43, 0x0F41, 0x0D5B, 0x0940);
    private static final Set<Integer> ALLOWED_RECORDS = Set.of(
        META_EOF, META_SET_MAP_MODE, META_SET_POLY_FILL_MODE, META_SELECT_OBJECT,
        META_DELETE_OBJECT, META_SET_WINDOW_ORG, META_SET_WINDOW_EXT,
        META_CREATE_PEN, META_CREATE_BRUSH, META_POLYGON, META_POLYPOLYGON,
        0x0F43);

    private static final int RASTER_DPI = 300;
    private static final int MAX_RASTER_SIDE = 4096;

    private WmfPreviewInspector() {
    }

    public static Inspection inspect(byte[] wmf) {
        if (wmf == null || wmf.length < 46) {
            return Inspection.invalid("WMF is shorter than the placeable and metafile headers");
        }
        ByteBuffer data = ByteBuffer.wrap(wmf).order(ByteOrder.LITTLE_ENDIAN);
        if (Integer.toUnsignedLong(data.getInt(0)) != PLACEABLE_KEY) {
            return Inspection.invalid("placeable WMF key is missing");
        }
        int left = data.getShort(6);
        int top = data.getShort(8);
        int right = data.getShort(10);
        int bottom = data.getShort(12);
        int unitsPerInch = Short.toUnsignedInt(data.getShort(14));
        if (right <= left || bottom <= top || unitsPerInch == 0) {
            return Inspection.invalid("placeable WMF bounds are empty");
        }
        int checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= Short.toUnsignedInt(data.getShort(offset));
        }
        if (checksum != Short.toUnsignedInt(data.getShort(20))) {
            return Inspection.invalid("placeable WMF checksum is invalid");
        }
        if (Short.toUnsignedInt(data.getShort(24)) != 9 || Short.toUnsignedInt(data.getShort(26)) < 0x0300) {
            return Inspection.invalid("WMF metafile header is invalid");
        }
        long declaredWords = Integer.toUnsignedLong(data.getInt(28));
        if (declaredWords * 2L != wmf.length - 22L) {
            return Inspection.invalid("WMF declared file size does not match payload length");
        }

        VectorRasterizer vector = new VectorRasterizer(right - left, bottom - top, unitsPerInch);
        int offset = 40;
        int recordCount = 0;
        int drawingRecordCount = 0;
        int bitmapRecordCount = 0;
        int textRecordCount = 0;
        int unsupportedRecordCount = 0;
        boolean eof = false;
        RasterStats dibRaster = null;
        Set<Integer> functions = new HashSet<>();
        while (offset + 6 <= wmf.length) {
            long sizeWords = Integer.toUnsignedLong(data.getInt(offset));
            if (sizeWords < 3 || sizeWords > Integer.MAX_VALUE / 2) {
                return Inspection.invalid("invalid WMF record size at byte " + offset);
            }
            int sizeBytes = (int) sizeWords * 2;
            if (offset + sizeBytes > wmf.length) {
                return Inspection.invalid("truncated WMF record at byte " + offset);
            }
            int function = Short.toUnsignedInt(data.getShort(offset + 4));
            functions.add(function);
            recordCount++;
            if (!ALLOWED_RECORDS.contains(function) && !BITMAP_RECORDS.contains(function)
                    && function != META_TEXT_OUT && function != META_EXT_TEXT_OUT) {
                unsupportedRecordCount++;
            }
            if (BITMAP_RECORDS.contains(function)) {
                bitmapRecordCount++;
                drawingRecordCount++;
                if (function == 0x0F43) {
                    dibRaster = inspectDib(data, offset + 6 + 22, offset + sizeBytes);
                }
            } else if (function == META_TEXT_OUT || function == META_EXT_TEXT_OUT) {
                textRecordCount++;
                drawingRecordCount++;
            } else if (function == META_POLYGON || function == META_POLYPOLYGON) {
                String error = vector.draw(function, data, offset + 6, offset + sizeBytes);
                if (error != null) {
                    return Inspection.invalid(error);
                }
                drawingRecordCount++;
            } else {
                String error = vector.applyState(function, data, offset + 6, offset + sizeBytes);
                if (error != null) {
                    return Inspection.invalid(error);
                }
            }
            if (function == META_EOF) {
                eof = true;
                offset += sizeBytes;
                break;
            }
            offset += sizeBytes;
        }
        if (!eof) {
            return Inspection.invalid("WMF EOF record is missing");
        }
        if (offset != wmf.length) {
            return Inspection.invalid("trailing bytes follow WMF EOF record");
        }
        if (dibRaster != null && dibRaster.error() != null) {
            return Inspection.invalid(dibRaster.error());
        }
        if (unsupportedRecordCount > 0) {
            return Inspection.invalid("WMF contains " + unsupportedRecordCount + " unsupported records: " + functions);
        }

        RasterStats raster = dibRaster != null ? dibRaster : vector.finish();
        return new Inspection(true, "", right - left, bottom - top, unitsPerInch,
            recordCount, drawingRecordCount, raster.width(), raster.height(), raster.foregroundPixels(),
            raster.foregroundDensity(), raster.touchesEdge(), bitmapRecordCount, textRecordCount,
            unsupportedRecordCount, dibRaster == null);
    }

    /** Independently rasterizes a validated pure-vector WMF at the inspector DPI. */
    public static RasterizedWmf rasterize(byte[] wmf) {
        return rasterize(wmf, RASTER_DPI);
    }

    /** Independently rasterizes a validated pure-vector WMF at the requested DPI. */
    public static RasterizedWmf rasterize(byte[] wmf, int dpi) {
        if (dpi <= 0 || dpi > 2400) {
            throw new IllegalArgumentException("invalid raster DPI: " + dpi);
        }
        Inspection inspection = inspect(wmf);
        if (!inspection.pureVector()) {
            throw new IllegalArgumentException("WMF is not a valid pure-vector preview: " + inspection.error());
        }
        ByteBuffer data = ByteBuffer.wrap(wmf).order(ByteOrder.LITTLE_ENDIAN);
        VectorRasterizer vector = new VectorRasterizer(
            inspection.physicalWidth(), inspection.physicalHeight(), inspection.unitsPerInch(), dpi);
        int offset = 40;
        while (offset + 6 <= wmf.length) {
            int sizeBytes = Math.toIntExact(Integer.toUnsignedLong(data.getInt(offset)) * 2L);
            int function = Short.toUnsignedInt(data.getShort(offset + 4));
            String error;
            if (function == META_POLYGON || function == META_POLYPOLYGON) {
                error = vector.draw(function, data, offset + 6, offset + sizeBytes);
            } else {
                error = vector.applyState(function, data, offset + 6, offset + sizeBytes);
            }
            if (error != null) {
                throw new IllegalArgumentException(error);
            }
            offset += sizeBytes;
            if (function == META_EOF) {
                break;
            }
        }
        return new RasterizedWmf(inspection, vector.finishImage());
    }

    private static final class VectorRasterizer {
        private final int physicalWidth;
        private final int physicalHeight;
        private final int widthPx;
        private final int heightPx;
        private final BufferedImage image;
        private final Graphics2D graphics;
        private final List<GdiObject> objects = new ArrayList<>();
        private int windowOrgX;
        private int windowOrgY;
        private int windowExtX;
        private int windowExtY;
        private Color brush;
        private int fillRule = Path2D.WIND_NON_ZERO;

        VectorRasterizer(int physicalWidth, int physicalHeight, int unitsPerInch) {
            this(physicalWidth, physicalHeight, unitsPerInch, RASTER_DPI);
        }

        VectorRasterizer(int physicalWidth, int physicalHeight, int unitsPerInch, int dpi) {
            this.physicalWidth = physicalWidth;
            this.physicalHeight = physicalHeight;
            double widthInches = physicalWidth / (double) unitsPerInch;
            double heightInches = physicalHeight / (double) unitsPerInch;
            double capScale = Math.min(1d, MAX_RASTER_SIDE /
                Math.max(Math.max(widthInches, heightInches) * dpi, 1d));
            this.widthPx = Math.max(1, (int) Math.ceil(widthInches * dpi * capScale));
            this.heightPx = Math.max(1, (int) Math.ceil(heightInches * dpi * capScale));
            this.image = new BufferedImage(widthPx, heightPx, BufferedImage.TYPE_INT_ARGB);
            this.graphics = image.createGraphics();
            graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        }

        String applyState(int function, ByteBuffer data, int start, int end) {
            if (function == META_SET_WINDOW_ORG && start + 4 <= end) {
                windowOrgY = data.getShort(start);
                windowOrgX = data.getShort(start + 2);
            } else if (function == META_SET_WINDOW_EXT && start + 4 <= end) {
                windowExtY = data.getShort(start);
                windowExtX = data.getShort(start + 2);
            } else if (function == META_SET_POLY_FILL_MODE && start + 2 <= end) {
                fillRule = Short.toUnsignedInt(data.getShort(start)) == 1
                    ? Path2D.WIND_EVEN_ODD : Path2D.WIND_NON_ZERO;
            } else if (function == META_CREATE_PEN) {
                allocate(new GdiObject(ObjectKind.PEN, null));
            } else if (function == META_CREATE_BRUSH) {
                if (start + 8 > end) {
                    return "truncated CREATEBRUSHINDIRECT record";
                }
                int style = Short.toUnsignedInt(data.getShort(start));
                long colorRef = Integer.toUnsignedLong(data.getInt(start + 2));
                Color color = style == 1 ? null : fromColorRef(colorRef);
                allocate(new GdiObject(ObjectKind.BRUSH, color));
            } else if (function == META_SELECT_OBJECT) {
                if (start + 2 > end) {
                    return "truncated SELECTOBJECT record";
                }
                int handle = Short.toUnsignedInt(data.getShort(start));
                if ((handle & 0x8000) != 0) {
                    int stock = handle & 0x7FFF;
                    if (stock == 5) {
                        brush = null;
                    } else if (stock <= 4) {
                        brush = switch (stock) {
                            case 0 -> Color.WHITE;
                            case 4 -> Color.BLACK;
                            default -> Color.GRAY;
                        };
                    }
                } else if (handle >= objects.size() || objects.get(handle) == null) {
                    return "SELECTOBJECT references missing handle " + handle;
                } else if (objects.get(handle).kind() == ObjectKind.BRUSH) {
                    brush = objects.get(handle).color();
                }
            } else if (function == META_DELETE_OBJECT) {
                if (start + 2 > end) {
                    return "truncated DELETEOBJECT record";
                }
                int handle = Short.toUnsignedInt(data.getShort(start));
                if (handle >= objects.size() || objects.get(handle) == null) {
                    return "DELETEOBJECT references missing handle " + handle;
                }
                objects.set(handle, null);
            }
            return null;
        }

        String draw(int function, ByteBuffer data, int start, int end) {
            if (windowExtX == 0 || windowExtY == 0) {
                return "polygon appears before non-zero window extents";
            }
            if (brush == null) {
                return "polygon is drawn without a selected solid brush";
            }
            int polygonCount;
            int cursor = start;
            int[] counts;
            if (function == META_POLYPOLYGON) {
                if (cursor + 2 > end) {
                    return "truncated POLYPOLYGON count";
                }
                polygonCount = Short.toUnsignedInt(data.getShort(cursor));
                cursor += 2;
                counts = new int[polygonCount];
                for (int i = 0; i < polygonCount; i++) {
                    if (cursor + 2 > end) {
                        return "truncated POLYPOLYGON contour counts";
                    }
                    counts[i] = Short.toUnsignedInt(data.getShort(cursor));
                    cursor += 2;
                }
            } else {
                if (cursor + 2 > end) {
                    return "truncated POLYGON count";
                }
                polygonCount = 1;
                counts = new int[]{Short.toUnsignedInt(data.getShort(cursor))};
                cursor += 2;
            }
            Path2D path = new Path2D.Double(fillRule);
            for (int count : counts) {
                if (count < 3 || cursor + count * 4L > end) {
                    return "invalid polygon point count " + count;
                }
                for (int i = 0; i < count; i++) {
                    int x = data.getShort(cursor);
                    int y = data.getShort(cursor + 2);
                    cursor += 4;
                    double px = (x - windowOrgX) * widthPx / (double) windowExtX;
                    double py = (y - windowOrgY) * heightPx / (double) windowExtY;
                    if (i == 0) {
                        path.moveTo(px, py);
                    } else {
                        path.lineTo(px, py);
                    }
                }
                path.closePath();
            }
            graphics.setColor(brush);
            graphics.fill(path);
            return null;
        }

        RasterStats finish() {
            graphics.dispose();
            int foreground = 0;
            boolean touchesEdge = false;
            for (int y = 0; y < heightPx; y++) {
                for (int x = 0; x < widthPx; x++) {
                    if (((image.getRGB(x, y) >>> 24) & 0xFF) != 0) {
                        foreground++;
                        if (x == 0 || y == 0 || x == widthPx - 1 || y == heightPx - 1) {
                            touchesEdge = true;
                        }
                    }
                }
            }
            double density = foreground / (double) ((long) widthPx * heightPx);
            return new RasterStats(widthPx, heightPx, foreground, density, touchesEdge, null);
        }

        BufferedImage finishImage() {
            graphics.dispose();
            return image;
        }

        private void allocate(GdiObject object) {
            for (int i = 0; i < objects.size(); i++) {
                if (objects.get(i) == null) {
                    objects.set(i, object);
                    return;
                }
            }
            objects.add(object);
        }

        private static Color fromColorRef(long value) {
            return new Color((int) (value & 0xFF), (int) ((value >>> 8) & 0xFF),
                (int) ((value >>> 16) & 0xFF));
        }
    }

    private static RasterStats inspectDib(ByteBuffer data, int offset, int recordEnd) {
        if (offset + 40 > recordEnd) {
            return RasterStats.invalid("WMF DIB header is truncated");
        }
        int headerSize = data.getInt(offset);
        int width = data.getInt(offset + 4);
        int storedHeight = data.getInt(offset + 8);
        int planes = Short.toUnsignedInt(data.getShort(offset + 12));
        int bitsPerPixel = Short.toUnsignedInt(data.getShort(offset + 14));
        int compression = data.getInt(offset + 16);
        if (headerSize < 40 || width <= 0 || storedHeight == 0 || planes != 1
                || bitsPerPixel != 24 || compression != 0) {
            return RasterStats.invalid("WMF DIB is not an uncompressed 24-bit bitmap");
        }
        int height = Math.abs(storedHeight);
        long strideLong = ((long) width * 3L + 3L) / 4L * 4L;
        long pixelsEnd = (long) offset + headerSize + strideLong * height;
        if (strideLong > Integer.MAX_VALUE || pixelsEnd > recordEnd) {
            return RasterStats.invalid("WMF DIB pixel data is truncated");
        }
        int stride = (int) strideLong;
        int pixelOffset = offset + headerSize;
        int foreground = 0;
        boolean touchesEdge = false;
        for (int storedY = 0; storedY < height; storedY++) {
            int y = storedHeight > 0 ? height - 1 - storedY : storedY;
            int row = pixelOffset + storedY * stride;
            for (int x = 0; x < width; x++) {
                int blue = Byte.toUnsignedInt(data.get(row + x * 3));
                int green = Byte.toUnsignedInt(data.get(row + x * 3 + 1));
                int red = Byte.toUnsignedInt(data.get(row + x * 3 + 2));
                if (Math.min(red, Math.min(green, blue)) < 245) {
                    foreground++;
                    if (x == 0 || x == width - 1 || y == 0 || y == height - 1) {
                        touchesEdge = true;
                    }
                }
            }
        }
        double density = foreground / (double) ((long) width * height);
        return new RasterStats(width, height, foreground, density, touchesEdge, null);
    }

    public record Inspection(
        boolean valid,
        String error,
        int physicalWidth,
        int physicalHeight,
        int unitsPerInch,
        int recordCount,
        int drawingRecordCount,
        int rasterWidth,
        int rasterHeight,
        int foregroundPixels,
        double foregroundDensity,
        boolean inkTouchesEdge,
        int bitmapRecordCount,
        int textRecordCount,
        int unsupportedRecordCount,
        boolean vector
    ) {
        private static Inspection invalid(String error) {
            return new Inspection(false, error, 0, 0, 0, 0, 0, -1, -1, -1, -1d,
                false, 0, 0, 0, false);
        }

        public boolean pureVector() {
            return valid && vector && bitmapRecordCount == 0 && textRecordCount == 0
                && unsupportedRecordCount == 0;
        }
    }

    public record RasterizedWmf(Inspection inspection, BufferedImage image) {
    }

    private enum ObjectKind { PEN, BRUSH }

    private record GdiObject(ObjectKind kind, Color color) {
    }

    private record RasterStats(int width, int height, int foregroundPixels,
                               double foregroundDensity, boolean touchesEdge, String error) {
        private static RasterStats invalid(String error) {
            return new RasterStats(0, 0, 0, -1d, false, error);
        }
    }
}
