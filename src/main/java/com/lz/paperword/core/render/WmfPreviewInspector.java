package com.lz.paperword.core.render;

import java.nio.ByteBuffer;
import java.nio.ByteOrder;

/** Structural and raster sanity checks for placeable WMF equation previews. */
public final class WmfPreviewInspector {

    private static final long PLACEABLE_KEY = 0x9AC6CDD7L;
    private static final int META_EOF = 0x0000;
    private static final int META_STRETCH_DIB = 0x0F43;
    private static final int META_POLYGON = 0x0324;
    private static final int META_POLYPOLYGON = 0x0538;

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

        int offset = 40;
        int recordCount = 0;
        int drawingRecordCount = 0;
        boolean eof = false;
        RasterStats raster = null;
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
            recordCount++;
            if (function == META_STRETCH_DIB) {
                raster = inspectDib(data, offset + 6 + 22, offset + sizeBytes);
                drawingRecordCount++;
            } else if (function == META_POLYGON || function == META_POLYPOLYGON) {
                drawingRecordCount++;
            } else if (function == META_EOF) {
                eof = true;
                break;
            }
            offset += sizeBytes;
        }
        if (!eof) {
            return Inspection.invalid("WMF EOF record is missing");
        }
        if (drawingRecordCount == 0) {
            return Inspection.invalid("WMF has no drawing records");
        }
        if (raster != null && raster.error() != null) {
            return Inspection.invalid(raster.error());
        }
        return new Inspection(true, "", right - left, bottom - top, unitsPerInch,
            recordCount, drawingRecordCount, raster == null ? -1 : raster.width(),
            raster == null ? -1 : raster.height(), raster == null ? -1 : raster.foregroundPixels(),
            raster == null ? -1d : raster.foregroundDensity(),
            raster != null && raster.touchesEdge());
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
        boolean inkTouchesEdge
    ) {
        private static Inspection invalid(String error) {
            return new Inspection(false, error, 0, 0, 0, 0, 0, -1, -1, -1, -1d, false);
        }
    }

    private record RasterStats(
        int width,
        int height,
        int foregroundPixels,
        double foregroundDensity,
        boolean touchesEdge,
        String error
    ) {
        private static RasterStats invalid(String error) {
            return new RasterStats(0, 0, 0, -1d, false, error);
        }
    }
}
