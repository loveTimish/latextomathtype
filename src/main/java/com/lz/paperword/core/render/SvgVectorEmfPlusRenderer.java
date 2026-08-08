package com.lz.paperword.core.render;

import java.awt.Color;
import java.awt.BasicStroke;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Area;
import java.awt.geom.PathIterator;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Adds an anti-aliased EMF+ path stream to the curve-preserving classic EMF fallback.
 *
 * <p>The result is an EMF+ Dual metafile. A classic EMF player sees native path records,
 * while an EMF+ player sees solid-color {@code EmfPlusFillPath} operations. Text has
 * already been outlined by {@link BatikVectorSceneBuilder}, so neither stream depends on
 * fonts installed on the playback machine.</p>
 */
public final class SvgVectorEmfPlusRenderer {

    /** Classic EMF uses a 0.01 mm virtual device grid, matching rclFrame exactly. */
    private static final double CLASSIC_DEVICE_UNITS_PER_PT = 2540d / 72d;
    private static final double EMF_PLUS_UNITS_PER_PT = 96d / 72d;
    private static final int EMF_PLUS_LOGICAL_DPI = 96;
    private static final int MAX_BYTE_ARRAY_SIZE = Integer.MAX_VALUE - 8;
    private static final int REFERENCE_DEVICE_WIDTH = 50_800;
    private static final int REFERENCE_DEVICE_HEIGHT = 28_575;
    private static final int REFERENCE_DEVICE_WIDTH_MM = 508;
    private static final int REFERENCE_DEVICE_HEIGHT_MM = 286;
    private static final int REFERENCE_DEVICE_WIDTH_UM = 508_000;
    private static final int REFERENCE_DEVICE_HEIGHT_UM = 285_750;
    static final String OPTICAL_COMPENSATION_PT_PROP = "paperword.ole.emfOpticalCompensationPt";
    /** Preserve the MathJax outline by default; compensation remains an opt-in diagnostic. */
    private static final double DEFAULT_OPTICAL_COMPENSATION_PT = 0.0d;
    private static final double MAX_OPTICAL_COMPENSATION_PT = 0.30d;

    private static final int EMR_HEADER = 0x00000001;
    private static final int EMR_POLYBEZIER_TO = 0x00000005;
    private static final int EMR_SET_WINDOW_EXT = 0x00000009;
    private static final int EMR_SET_VIEWPORT_EXT = 0x0000000B;
    private static final int EMR_EOF = 0x0000000E;
    private static final int EMR_SET_MAP_MODE = 0x00000011;
    private static final int EMR_SET_POLY_FILL_MODE = 0x00000013;
    private static final int EMR_MOVE_TO_EX = 0x0000001B;
    private static final int EMR_SELECT_OBJECT = 0x00000025;
    private static final int EMR_CREATE_BRUSH_INDIRECT = 0x00000027;
    private static final int EMR_DELETE_OBJECT = 0x00000028;
    private static final int EMR_LINE_TO = 0x00000036;
    private static final int EMR_BEGIN_PATH = 0x0000003B;
    private static final int EMR_END_PATH = 0x0000003C;
    private static final int EMR_CLOSE_FIGURE = 0x0000003D;
    private static final int EMR_FILL_PATH = 0x0000003E;
    private static final int EMR_GDICOMMENT = 0x00000046;
    private static final int ENHMETA_SIGNATURE = 0x464D4520;
    private static final int EMF_PLUS_SIGNATURE = 0x2B464D45;

    private static final int CLASSIC_BRUSH_HANDLE = 1;
    private static final int STOCK_NULL_BRUSH = 0x80000005;
    private static final int POLY_FILL_ALTERNATE = 1;
    private static final int MAP_MODE_ANISOTROPIC = 8;

    private static final int EMF_PLUS_HEADER = 0x4001;
    private static final int EMF_PLUS_EOF = 0x4002;
    private static final int EMF_PLUS_OBJECT = 0x4008;
    private static final int EMF_PLUS_FILL_PATH = 0x4014;
    private static final int EMF_PLUS_SET_ANTI_ALIAS_MODE = 0x401E;
    private static final int EMF_PLUS_SET_PIXEL_OFFSET_MODE = 0x4022;

    private static final int EMF_PLUS_DUAL = 0x0001;
    private static final int EMF_PLUS_VIDEO_DISPLAY = 0x00000001;
    private static final int GRAPHICS_VERSION_1_1 = 0xDBC01002;
    private static final int SMOOTHING_MODE_ANTI_ALIAS_8X8 = 0x000B;
    private static final int PIXEL_OFFSET_MODE_HIGH_QUALITY = 0x0002;
    private static final int OBJECT_TYPE_PATH = 0x03;
    private static final int PATH_OBJECT_ID = 0;
    private static final int SOLID_COLOR_FLAG = 0x8000;

    private static final int PATH_POINT_START = 0x00;
    private static final int PATH_POINT_LINE = 0x01;
    private static final int PATH_POINT_BEZIER = 0x03;
    private static final int PATH_POINT_CLOSE_SUBPATH = 0x80;

    private SvgVectorEmfPlusRenderer() {
    }

    public static boolean isValidDualVector(byte[] emf) {
        if (emf == null || emf.length < 88) {
            return false;
        }
        try {
            ClassicEmfLayout layout = inspectClassicEmf(emf);
            if (littleEndianInt(emf, 48) != emf.length
                    || littleEndianInt(emf, 52) != countEmfRecords(emf)) {
                return false;
            }
            if (layout.headerSize() != 108 || !isValidGeneratedHeader(emf)) {
                return false;
            }
            ClassicValidation classic = new ClassicValidation(
                Math.addExact(littleEndianInt(emf, 16), 1),
                Math.addExact(littleEndianInt(emf, 20), 1));
            EmfPlusValidation plus = new EmfPlusValidation();
            int offset = layout.headerSize();
            while (offset < emf.length) {
                int type = littleEndianInt(emf, offset);
                int size = littleEndianInt(emf, offset + 4);
                if (!isAllowedClassicRecord(type) || isBitmapRecord(type) || isTextRecord(type)) {
                    return false;
                }
                if (type == EMR_GDICOMMENT) {
                    PlusCommentKind kind = plus.acceptComment(emf, offset, size);
                    if (!classic.acceptComment(kind)) {
                        throw new IllegalArgumentException(
                            "unexpected EMF+ comment in classic phase at " + offset);
                    }
                } else if (!classic.acceptRecord(emf, offset, type, size, plus.isComplete())) {
                    throw new IllegalArgumentException(
                        "invalid classic EMF record 0x" + Integer.toHexString(type)
                            + " at " + offset);
                }
                offset += size;
            }
            int expectedHandles = classic.pathCount() == 0 ? 1 : 2;
            boolean valid = classic.isComplete()
                && plus.isComplete()
                && classic.pathCount() == plus.pathCount()
                && classic.rgbColors().equals(plus.rgbColors())
                && littleEndianUnsignedShort(emf, 56) == expectedHandles;
            if (!valid && Boolean.getBoolean("paperword.ole.emfValidationDebug")) {
                System.err.println("EMF+ final-state mismatch: classicComplete="
                    + classic.isComplete() + ", plusComplete=" + plus.isComplete()
                    + ", classicPaths=" + classic.pathCount() + ", plusPaths="
                    + plus.pathCount() + ", colorsMatch="
                    + classic.rgbColors().equals(plus.rgbColors()) + ", handles="
                    + littleEndianUnsignedShort(emf, 56) + ", expectedHandles="
                    + expectedHandles);
            }
            return valid;
        } catch (RuntimeException | SvgVectorWmfRenderer.SvgVectorWmfException e) {
            if (Boolean.getBoolean("paperword.ole.emfValidationDebug")) {
                System.err.println("EMF+ validation failed: " + e.getMessage());
            }
            return false;
        }
    }

    static byte[] render(byte[] svgBytes, double widthPt, double heightPt)
        throws SvgVectorWmfRenderer.SvgVectorWmfException {
        return render(svgBytes, widthPt, heightPt, opticalCompensationPt());
    }

    static byte[] render(byte[] svgBytes, double widthPt, double heightPt,
                         double opticalCompensationPt)
        throws SvgVectorWmfRenderer.SvgVectorWmfException {
        validateInput(svgBytes, widthPt, heightPt);
        if (!Double.isFinite(opticalCompensationPt) || opticalCompensationPt < 0d
                || opticalCompensationPt > MAX_OPTICAL_COMPENSATION_PT) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                "invalid EMF+ optical compensation: " + opticalCompensationPt + "pt");
        }

        BatikVectorSceneBuilder.VectorScene scene;
        try {
            scene = BatikVectorSceneBuilder.build(svgBytes);
        } catch (BatikVectorSceneBuilder.VectorSceneException e) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException(e.getMessage());
        }

        try {
            RenderGeometry geometry = renderGeometry(scene, widthPt, heightPt);
            List<EncodedPath> paths = encodePaths(scene, geometry, opticalCompensationPt);

            byte[] header = classicEmfHeader(
                geometry.classicDeviceWidth(), geometry.classicDeviceHeight(), !paths.isEmpty());
            byte[] headerComment = emfPlusComment(emfPlusHeader(
                EMF_PLUS_LOGICAL_DPI, EMF_PLUS_LOGICAL_DPI));
            byte[] antiAliasComment = emfPlusComment(emfPlusRecord(
                EMF_PLUS_SET_ANTI_ALIAS_MODE, SMOOTHING_MODE_ANTI_ALIAS_8X8, new byte[0]));
            byte[] pixelOffsetComment = emfPlusComment(emfPlusRecord(
                EMF_PLUS_SET_PIXEL_OFFSET_MODE, PIXEL_OFFSET_MODE_HIGH_QUALITY, new byte[0]));
            byte[] plusEofComment = emfPlusComment(emfPlusRecord(
                EMF_PLUS_EOF, 0, new byte[0]));

            ByteArrayOutputStream output = new ByteArrayOutputStream(64 * 1024);
            appendBytes(output, header);
            appendBytes(output, headerComment);
            appendBytes(output, emfRecordWithInt(EMR_SET_MAP_MODE, MAP_MODE_ANISOTROPIC));
            appendBytes(output, emfPointRecord(EMR_SET_WINDOW_EXT,
                new IntPoint(geometry.classicDeviceWidth(), geometry.classicDeviceHeight())));
            appendBytes(output, emfPointRecord(EMR_SET_VIEWPORT_EXT,
                new IntPoint(geometry.classicDeviceWidth(), geometry.classicDeviceHeight())));
            appendBytes(output, emfRecordWithInt(EMR_SET_POLY_FILL_MODE, POLY_FILL_ALTERNATE));
            for (EncodedPath path : paths) {
                writeClassicFilledPath(output, path);
            }
            appendBytes(output, antiAliasComment);
            appendBytes(output, pixelOffsetComment);
            for (EncodedPath path : paths) {
                appendBytes(output, path.comment());
            }
            appendBytes(output, plusEofComment);
            appendBytes(output, classicEofRecord());

            byte[] dualEmf = output.toByteArray();
            putLittleEndianInt(dualEmf, 48, dualEmf.length);
            putLittleEndianInt(dualEmf, 52, countEmfRecords(dualEmf));
            return dualEmf;
        } catch (ArithmeticException | IllegalArgumentException e) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                "EMF+ emit failed: " + e.getMessage());
        }
    }

    static double opticalCompensationPt() {
        String configured = System.getProperty(OPTICAL_COMPENSATION_PT_PROP);
        if (configured == null || configured.isBlank()) {
            return DEFAULT_OPTICAL_COMPENSATION_PT;
        }
        try {
            double value = Double.parseDouble(configured.trim());
            if (Double.isFinite(value) && value >= 0d && value <= MAX_OPTICAL_COMPENSATION_PT) {
                return value;
            }
        } catch (NumberFormatException ignored) {
            // Rejected below with the same stable diagnostic as an out-of-range value.
        }
        throw new IllegalArgumentException(
            "paperword.ole.emfOpticalCompensationPt must be between 0 and "
                + MAX_OPTICAL_COMPENSATION_PT + " points");
    }

    private static void validateInput(byte[] svgBytes, double widthPt, double heightPt)
        throws SvgVectorWmfRenderer.SvgVectorWmfException {
        if (svgBytes == null || svgBytes.length == 0) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException("empty SVG input");
        }
        if (!Double.isFinite(widthPt) || !Double.isFinite(heightPt)
                || widthPt <= 0d || heightPt <= 0d) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                "non-positive or non-finite target size: " + widthPt + "x" + heightPt);
        }
    }

    private static RenderGeometry renderGeometry(BatikVectorSceneBuilder.VectorScene scene,
                                                 double widthPt, double heightPt) {
        int classicDeviceWidth = checkedPositiveCeilToInt(
            widthPt * CLASSIC_DEVICE_UNITS_PER_PT, "classic EMF device width");
        int classicDeviceHeight = checkedPositiveCeilToInt(
            heightPt * CLASSIC_DEVICE_UNITS_PER_PT, "classic EMF device height");
        double scale = Math.min(
            classicDeviceWidth / scene.viewportWidth(),
            classicDeviceHeight / scene.viewportHeight());
        double offsetX = (classicDeviceWidth - scene.viewportWidth() * scale) / 2d;
        double offsetY = (classicDeviceHeight - scene.viewportHeight() * scale) / 2d;
        AffineTransform classicTransform = new AffineTransform(
            scale, 0d, 0d, scale, offsetX, offsetY);

        double plusWidth = widthPt * EMF_PLUS_UNITS_PER_PT;
        double plusHeight = heightPt * EMF_PLUS_UNITS_PER_PT;
        double plusScale = Math.min(
            plusWidth / scene.viewportWidth(), plusHeight / scene.viewportHeight());
        double plusOffsetX = (plusWidth - scene.viewportWidth() * plusScale) / 2d;
        double plusOffsetY = (plusHeight - scene.viewportHeight() * plusScale) / 2d;
        AffineTransform plusTransform = new AffineTransform(
            plusScale, 0d, 0d, plusScale, plusOffsetX, plusOffsetY);
        return new RenderGeometry(classicDeviceWidth, classicDeviceHeight,
            checkedPositiveCeilToInt(plusWidth, "EMF+ viewport width"),
            checkedPositiveCeilToInt(plusHeight, "EMF+ viewport height"),
            classicTransform, plusTransform);
    }

    private static List<EncodedPath> encodePaths(BatikVectorSceneBuilder.VectorScene scene,
                                                  RenderGeometry geometry,
                                                  double opticalCompensationPt) {
        List<EncodedPath> paths = new ArrayList<>();
        for (BatikVectorSceneBuilder.PaintedShape painted : scene.shapes()) {
            Shape classicShape = geometry.classicTransform().createTransformedShape(painted.shape());
            Shape plusShape = geometry.plusTransform().createTransformedShape(painted.shape());
            if (shouldOpticallyCompensate(plusShape, opticalCompensationPt)) {
                classicShape = expandFilledShape(classicShape,
                    opticalCompensationPt * CLASSIC_DEVICE_UNITS_PER_PT,
                    geometry.classicDeviceWidth(), geometry.classicDeviceHeight());
                plusShape = expandFilledShape(plusShape,
                    opticalCompensationPt * EMF_PLUS_UNITS_PER_PT,
                    geometry.plusViewportWidth(), geometry.plusViewportHeight());
            }
            PathData classicPath = toPathData(new Area(classicShape));
            PathData plusPath = toPathData(new Area(plusShape));
            if (classicPath.points().isEmpty() || plusPath.points().isEmpty()) {
                continue;
            }
            byte[] object = emfPlusPathObject(plusPath);
            byte[] fill = emfPlusFillPath(painted.color());
            paths.add(new EncodedPath(classicPath, painted.color(), emfPlusComment(object, fill)));
        }
        return paths;
    }

    private static boolean shouldOpticallyCompensate(Shape plusShape,
                                                       double opticalCompensationPt) {
        if (opticalCompensationPt <= 0d) {
            return false;
        }
        Rectangle2D bounds = plusShape.getBounds2D();
        if (bounds.isEmpty()) {
            return false;
        }
        double widthPt = bounds.getWidth() / EMF_PLUS_UNITS_PER_PT;
        double heightPt = bounds.getHeight() / EMF_PLUS_UNITS_PER_PT;
        double minorPt = Math.min(widthPt, heightPt);
        double majorPt = Math.max(widthPt, heightPt);
        boolean structuralRule = minorPt <= 0.75d && majorPt >= 2.5d
            && majorPt / Math.max(minorPt, 0.001d) >= 4d;
        return !structuralRule && majorPt <= 24d;
    }

    private static Shape expandFilledShape(Shape shape, double radiusUnits,
                                           double widthUnits, double heightUnits) {
        Area expanded = new Area(shape);
        BasicStroke stroke = new BasicStroke((float) (radiusUnits * 2d),
            BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND);
        expanded.add(new Area(stroke.createStrokedShape(shape)));
        expanded.intersect(new Area(new Rectangle2D.Double(0d, 0d, widthUnits, heightUnits)));
        return expanded;
    }

    private static byte[] classicEmfHeader(int classicDeviceWidth, int classicDeviceHeight,
                                           boolean hasPaths) {
        LittleEndianWriter header = new LittleEndianWriter(108);
        header.writeInt(EMR_HEADER);
        header.writeInt(108);
        header.writeRect(0, 0, classicDeviceWidth - 1, classicDeviceHeight - 1);
        header.writeRect(0, 0, classicDeviceWidth, classicDeviceHeight);
        header.writeInt(ENHMETA_SIGNATURE);
        header.writeInt(0x00010000);
        header.writeInt(0); // nBytes is finalized after all classic and EMF+ records exist.
        header.writeInt(0); // nRecords is finalized with nBytes.
        header.writeShort(hasPaths ? 2 : 1); // Handle zero plus the reusable brush, if present.
        header.writeShort(0);
        header.writeInt(0); // No description.
        header.writeInt(0);
        header.writeInt(0); // No palette.
        header.writeInt(REFERENCE_DEVICE_WIDTH);
        header.writeInt(REFERENCE_DEVICE_HEIGHT);
        header.writeInt(REFERENCE_DEVICE_WIDTH_MM);
        header.writeInt(REFERENCE_DEVICE_HEIGHT_MM);
        header.writeInt(0); // No pixel format descriptor.
        header.writeInt(0);
        header.writeInt(0); // No OpenGL records.
        header.writeInt(REFERENCE_DEVICE_WIDTH_UM);
        header.writeInt(REFERENCE_DEVICE_HEIGHT_UM);
        if (header.size() != 108) {
            throw new IllegalArgumentException("classic EMF header size mismatch");
        }
        return header.toByteArray();
    }

    private static void writeClassicFilledPath(ByteArrayOutputStream output, EncodedPath encoded) {
        appendBytes(output, classicCreateBrushRecord(encoded.color()));
        appendBytes(output, emfRecordWithInt(EMR_SELECT_OBJECT, CLASSIC_BRUSH_HANDLE));
        appendBytes(output, emfRecord(EMR_BEGIN_PATH));

        List<PathPoint> points = encoded.path().points();
        IntPoint current = null;
        int index = 0;
        while (index < points.size()) {
            PathPoint point = points.get(index);
            int kind = point.type() & 0x0F;
            boolean close = (point.type() & PATH_POINT_CLOSE_SUBPATH) != 0;
            if (kind == PATH_POINT_START) {
                current = integerPoint(point);
                appendBytes(output, emfPointRecord(EMR_MOVE_TO_EX, current));
                if (close) {
                    throw new IllegalArgumentException("empty classic EMF figure cannot be closed");
                }
                index++;
                continue;
            }
            if (current == null) {
                throw new IllegalArgumentException("classic EMF path starts without MOVETO");
            }
            if (kind == PATH_POINT_LINE) {
                current = integerPoint(point);
                appendBytes(output, emfPointRecord(EMR_LINE_TO, current));
                if (close) {
                    appendBytes(output, emfRecord(EMR_CLOSE_FIGURE));
                }
                index++;
                continue;
            }
            if (kind == PATH_POINT_BEZIER) {
                List<IntPoint> curve = new ArrayList<>();
                boolean closeCurve = false;
                while (index < points.size()
                        && (points.get(index).type() & 0x0F) == PATH_POINT_BEZIER) {
                    PathPoint curvePoint = points.get(index++);
                    closeCurve = (curvePoint.type() & PATH_POINT_CLOSE_SUBPATH) != 0;
                    if (closeCurve && index < points.size()
                            && (points.get(index).type() & 0x0F) == PATH_POINT_BEZIER) {
                        throw new IllegalArgumentException("close flag appears inside a Bezier sequence");
                    }
                    curve.add(integerPoint(curvePoint));
                }
                if (curve.isEmpty() || curve.size() % 3 != 0) {
                    throw new IllegalArgumentException("classic EMF Bezier point count is not divisible by three");
                }
                appendBytes(output, classicPolyBezierToRecord(current, curve));
                current = curve.getLast();
                if (closeCurve) {
                    appendBytes(output, emfRecord(EMR_CLOSE_FIGURE));
                }
                continue;
            }
            throw new IllegalArgumentException("unsupported classic EMF path point type: " + kind);
        }

        appendBytes(output, emfRecord(EMR_END_PATH));
        appendBytes(output, classicFillPathRecord(bounds(points)));
        appendBytes(output, emfRecordWithInt(EMR_SELECT_OBJECT, STOCK_NULL_BRUSH));
        appendBytes(output, emfRecordWithInt(EMR_DELETE_OBJECT, CLASSIC_BRUSH_HANDLE));
    }

    private static byte[] classicCreateBrushRecord(Color color) {
        LittleEndianWriter record = new LittleEndianWriter(24);
        record.writeInt(EMR_CREATE_BRUSH_INDIRECT);
        record.writeInt(24);
        record.writeInt(CLASSIC_BRUSH_HANDLE);
        record.writeInt(0); // BS_SOLID.
        record.writeInt(color.getRed() | (color.getGreen() << 8) | (color.getBlue() << 16));
        record.writeInt(0);
        return record.toByteArray();
    }

    private static byte[] classicPolyBezierToRecord(IntPoint current, List<IntPoint> points) {
        int size = checkedRecordCapacity(
            "classic EMR_POLYBEZIER_TO", 28L + (long) points.size() * 8L);
        LittleEndianWriter record = new LittleEndianWriter(size);
        record.writeInt(EMR_POLYBEZIER_TO);
        record.writeInt(size);
        List<IntPoint> boundsPoints = new ArrayList<>(
            checkedElementCount("classic Bezier bounds", (long) points.size() + 1L));
        boundsPoints.add(current);
        boundsPoints.addAll(points);
        record.writeRect(bounds(boundsPoints));
        record.writeInt(points.size());
        for (IntPoint point : points) {
            record.writeInt(point.x());
            record.writeInt(point.y());
        }
        return record.toByteArray();
    }

    private static byte[] classicFillPathRecord(Bounds bounds) {
        LittleEndianWriter record = new LittleEndianWriter(24);
        record.writeInt(EMR_FILL_PATH);
        record.writeInt(24);
        record.writeRect(bounds);
        return record.toByteArray();
    }

    private static byte[] classicEofRecord() {
        LittleEndianWriter record = new LittleEndianWriter(20);
        record.writeInt(EMR_EOF);
        record.writeInt(20);
        record.writeInt(0);
        record.writeInt(16);
        record.writeInt(20);
        return record.toByteArray();
    }

    private static byte[] emfRecord(int type) {
        LittleEndianWriter record = new LittleEndianWriter(8);
        record.writeInt(type);
        record.writeInt(8);
        return record.toByteArray();
    }

    private static byte[] emfRecordWithInt(int type, int value) {
        LittleEndianWriter record = new LittleEndianWriter(12);
        record.writeInt(type);
        record.writeInt(12);
        record.writeInt(value);
        return record.toByteArray();
    }

    private static byte[] emfPointRecord(int type, IntPoint point) {
        LittleEndianWriter record = new LittleEndianWriter(16);
        record.writeInt(type);
        record.writeInt(16);
        record.writeInt(point.x());
        record.writeInt(point.y());
        return record.toByteArray();
    }

    private static IntPoint integerPoint(PathPoint point) {
        return new IntPoint(
            checkedRoundToInt(point.x(), "classic EMF x coordinate"),
            checkedRoundToInt(point.y(), "classic EMF y coordinate"));
    }

    private static int checkedPositiveCeilToInt(double value, String description) {
        if (!Double.isFinite(value) || value <= 0d) {
            throw new IllegalArgumentException(
                description + " is non-positive, non-finite, or overflowed: " + value);
        }
        double rounded = Math.ceil(value);
        if (rounded > Integer.MAX_VALUE) {
            throw new IllegalArgumentException(
                description + " exceeds signed 32-bit range: " + value);
        }
        return (int) rounded;
    }

    static int checkedRoundToInt(float value, String description) {
        if (!Float.isFinite(value)) {
            throw new IllegalArgumentException(description + " is non-finite: " + value);
        }
        double rounded = Math.floor((double) value + 0.5d);
        if (rounded < Integer.MIN_VALUE || rounded > Integer.MAX_VALUE) {
            throw new IllegalArgumentException(
                description + " exceeds signed 32-bit range: " + value);
        }
        return (int) rounded;
    }

    static int checkedRecordCapacity(String description, long size) {
        if (size < 0L || size > MAX_BYTE_ARRAY_SIZE) {
            throw new IllegalArgumentException(
                description + " capacity exceeds supported byte-array range: " + size);
        }
        return (int) size;
    }

    private static int checkedElementCount(String description, long count) {
        if (count < 0L || count > Integer.MAX_VALUE) {
            throw new IllegalArgumentException(
                description + " element count exceeds signed 32-bit range: " + count);
        }
        return (int) count;
    }

    private static void appendBytes(ByteArrayOutputStream output, byte[] bytes) {
        checkedRecordCapacity("complete EMF", (long) output.size() + bytes.length);
        output.writeBytes(bytes);
    }

    private static Bounds bounds(List<PathPoint> points) {
        return bounds(points.stream().map(SvgVectorEmfPlusRenderer::integerPoint).toList());
    }

    private static Bounds bounds(Iterable<IntPoint> points) {
        int left = Integer.MAX_VALUE;
        int top = Integer.MAX_VALUE;
        int right = Integer.MIN_VALUE;
        int bottom = Integer.MIN_VALUE;
        for (IntPoint point : points) {
            left = Math.min(left, point.x());
            top = Math.min(top, point.y());
            right = Math.max(right, point.x());
            bottom = Math.max(bottom, point.y());
        }
        if (left == Integer.MAX_VALUE) {
            throw new IllegalArgumentException("classic EMF path has no bounds");
        }
        return new Bounds(left, top, right, bottom);
    }

    private static boolean isValidGeneratedHeader(byte[] emf) {
        return littleEndianInt(emf, 8) == 0
            && littleEndianInt(emf, 12) == 0
            && littleEndianInt(emf, 16) >= 0
            && littleEndianInt(emf, 20) >= 0
            && littleEndianInt(emf, 24) == 0
            && littleEndianInt(emf, 28) == 0
            && (long) littleEndianInt(emf, 32) == (long) littleEndianInt(emf, 16) + 1L
            && (long) littleEndianInt(emf, 36) == (long) littleEndianInt(emf, 20) + 1L
            && littleEndianInt(emf, 44) == 0x00010000
            && littleEndianUnsignedShort(emf, 58) == 0
            && littleEndianInt(emf, 60) == 0
            && littleEndianInt(emf, 64) == 0
            && littleEndianInt(emf, 68) == 0
            && littleEndianInt(emf, 72) == REFERENCE_DEVICE_WIDTH
            && littleEndianInt(emf, 76) == REFERENCE_DEVICE_HEIGHT
            && littleEndianInt(emf, 80) == REFERENCE_DEVICE_WIDTH_MM
            && littleEndianInt(emf, 84) == REFERENCE_DEVICE_HEIGHT_MM
            && littleEndianInt(emf, 88) == 0
            && littleEndianInt(emf, 92) == 0
            && littleEndianInt(emf, 96) == 0
            && littleEndianInt(emf, 100) == REFERENCE_DEVICE_WIDTH_UM
            && littleEndianInt(emf, 104) == REFERENCE_DEVICE_HEIGHT_UM;
    }

    private enum ClassicPhase {
        EXPECT_PLUS_HEADER,
        EXPECT_MAP_MODE,
        EXPECT_WINDOW_EXTENT,
        EXPECT_VIEWPORT_EXTENT,
        EXPECT_FILL_MODE,
        READY,
        EXPECT_BRUSH_SELECT,
        EXPECT_BEGIN_PATH,
        PATH,
        EXPECT_FILL_PATH,
        EXPECT_NULL_BRUSH,
        EXPECT_DELETE_BRUSH,
        PLUS_SECTION,
        COMPLETE
    }

    private enum PlusPhase {
        EXPECT_HEADER,
        EXPECT_ANTI_ALIAS,
        EXPECT_PIXEL_OFFSET,
        EXPECT_OBJECT_OR_EOF,
        EXPECT_FILL,
        COMPLETE
    }

    private enum PlusCommentKind {
        HEADER,
        ANTI_ALIAS,
        PIXEL_OFFSET,
        PATH_PAIR,
        EOF
    }

    private static final class ClassicValidation {
        private final int expectedWidth;
        private final int expectedHeight;
        private final List<Integer> rgbColors = new ArrayList<>();
        private ClassicPhase phase = ClassicPhase.EXPECT_PLUS_HEADER;
        private IntPoint current;
        private BoundsAccumulator pathBounds;
        private boolean figureActive;
        private boolean figureDrawable;
        private boolean pathDrawable;
        private int pathCount;

        ClassicValidation(int expectedWidth, int expectedHeight) {
            this.expectedWidth = expectedWidth;
            this.expectedHeight = expectedHeight;
        }

        boolean acceptComment(PlusCommentKind kind) {
            if (phase == ClassicPhase.EXPECT_PLUS_HEADER && kind == PlusCommentKind.HEADER) {
                phase = ClassicPhase.EXPECT_MAP_MODE;
                return true;
            }
            if (phase == ClassicPhase.READY && kind == PlusCommentKind.ANTI_ALIAS) {
                phase = ClassicPhase.PLUS_SECTION;
                return true;
            }
            return phase == ClassicPhase.PLUS_SECTION
                && (kind == PlusCommentKind.PIXEL_OFFSET
                    || kind == PlusCommentKind.PATH_PAIR
                    || kind == PlusCommentKind.EOF);
        }

        boolean acceptRecord(byte[] emf, int offset, int type, int size,
                             boolean plusComplete) {
            return switch (phase) {
                case EXPECT_MAP_MODE -> acceptIntRecord(
                    emf, offset, type, size, EMR_SET_MAP_MODE, MAP_MODE_ANISOTROPIC,
                    ClassicPhase.EXPECT_WINDOW_EXTENT);
                case EXPECT_WINDOW_EXTENT -> acceptExtentRecord(
                    emf, offset, type, size, EMR_SET_WINDOW_EXT,
                    expectedWidth, expectedHeight, ClassicPhase.EXPECT_VIEWPORT_EXTENT);
                case EXPECT_VIEWPORT_EXTENT -> acceptPositiveExtentRecord(
                    emf, offset, type, size, EMR_SET_VIEWPORT_EXT,
                    expectedWidth, expectedHeight,
                    ClassicPhase.EXPECT_FILL_MODE);
                case EXPECT_FILL_MODE -> acceptIntRecord(
                    emf, offset, type, size, EMR_SET_POLY_FILL_MODE, POLY_FILL_ALTERNATE,
                    ClassicPhase.READY);
                case READY -> acceptCreateBrush(emf, offset, type, size);
                case EXPECT_BRUSH_SELECT -> acceptIntRecord(
                    emf, offset, type, size, EMR_SELECT_OBJECT, CLASSIC_BRUSH_HANDLE,
                    ClassicPhase.EXPECT_BEGIN_PATH);
                case EXPECT_BEGIN_PATH -> acceptBeginPath(type, size);
                case PATH -> acceptPathRecord(emf, offset, type, size);
                case EXPECT_FILL_PATH -> acceptFillPath(emf, offset, type, size);
                case EXPECT_NULL_BRUSH -> acceptIntRecord(
                    emf, offset, type, size, EMR_SELECT_OBJECT, STOCK_NULL_BRUSH,
                    ClassicPhase.EXPECT_DELETE_BRUSH);
                case EXPECT_DELETE_BRUSH -> acceptDeleteBrush(
                    emf, offset, type, size);
                case PLUS_SECTION -> acceptEof(emf, offset, type, size, plusComplete);
                default -> false;
            };
        }

        private boolean acceptIntRecord(byte[] emf, int offset, int type, int size,
                                        int expectedType, int expectedValue,
                                        ClassicPhase next) {
            if (type != expectedType || size != 12
                    || littleEndianInt(emf, offset + 8) != expectedValue) {
                return false;
            }
            phase = next;
            return true;
        }

        private boolean acceptExtentRecord(byte[] emf, int offset, int type, int size,
                                            int expectedType, int width, int height,
                                            ClassicPhase next) {
            if (type != expectedType || size != 16
                    || littleEndianInt(emf, offset + 8) != width
                    || littleEndianInt(emf, offset + 12) != height) {
                return false;
            }
            phase = next;
            return true;
        }

        private boolean acceptPositiveExtentRecord(byte[] emf, int offset, int type, int size,
                                                    int expectedType, int width, int height,
                                                    ClassicPhase next) {
            if (type != expectedType || size != 16
                    || littleEndianInt(emf, offset + 8) != width
                    || littleEndianInt(emf, offset + 12) != height) {
                return false;
            }
            phase = next;
            return true;
        }

        private boolean acceptCreateBrush(byte[] emf, int offset, int type, int size) {
            if (type != EMR_CREATE_BRUSH_INDIRECT || size != 24
                    || littleEndianInt(emf, offset + 8) != CLASSIC_BRUSH_HANDLE
                    || littleEndianInt(emf, offset + 12) != 0
                    || littleEndianInt(emf, offset + 20) != 0) {
                return false;
            }
            int colorRef = littleEndianInt(emf, offset + 16);
            if ((colorRef & 0xFF000000) != 0) {
                return false;
            }
            int rgb = ((colorRef & 0xFF) << 16)
                | (colorRef & 0xFF00)
                | ((colorRef >>> 16) & 0xFF);
            rgbColors.add(rgb);
            phase = ClassicPhase.EXPECT_BRUSH_SELECT;
            return true;
        }

        private boolean acceptBeginPath(int type, int size) {
            if (type != EMR_BEGIN_PATH || size != 8) {
                return false;
            }
            current = null;
            pathBounds = new BoundsAccumulator();
            figureActive = false;
            figureDrawable = false;
            pathDrawable = false;
            phase = ClassicPhase.PATH;
            return true;
        }

        private boolean acceptPathRecord(byte[] emf, int offset, int type, int size) {
            if (type == EMR_MOVE_TO_EX) {
                if (size != 16 || (figureActive && !figureDrawable)) {
                    return false;
                }
                current = new IntPoint(
                    littleEndianInt(emf, offset + 8), littleEndianInt(emf, offset + 12));
                pathBounds.add(current);
                figureActive = true;
                figureDrawable = false;
                return true;
            }
            if (type == EMR_LINE_TO) {
                if (size != 16 || !figureActive || current == null) {
                    return false;
                }
                current = new IntPoint(
                    littleEndianInt(emf, offset + 8), littleEndianInt(emf, offset + 12));
                pathBounds.add(current);
                figureDrawable = true;
                pathDrawable = true;
                return true;
            }
            if (type == EMR_POLYBEZIER_TO) {
                return acceptPolyBezier(emf, offset, size);
            }
            if (type == EMR_CLOSE_FIGURE) {
                if (size != 8 || !figureActive || !figureDrawable) {
                    return false;
                }
                figureActive = false;
                return true;
            }
            if (type == EMR_END_PATH) {
                if (size != 8 || !pathDrawable
                        || (figureActive && !figureDrawable) || pathBounds.isEmpty()) {
                    return false;
                }
                phase = ClassicPhase.EXPECT_FILL_PATH;
                return true;
            }
            return false;
        }

        private boolean acceptPolyBezier(byte[] emf, int offset, int size) {
            if (current == null || !figureActive || size < 52) {
                return false;
            }
            int count = littleEndianInt(emf, offset + 24);
            long expectedSize = 28L + 8L * count;
            if (count <= 0 || count % 3 != 0 || expectedSize != size) {
                return false;
            }
            BoundsAccumulator recordBounds = new BoundsAccumulator();
            recordBounds.add(current);
            for (int index = 0; index < count; index++) {
                IntPoint point = new IntPoint(
                    littleEndianInt(emf, offset + 28 + index * 8),
                    littleEndianInt(emf, offset + 32 + index * 8));
                recordBounds.add(point);
                pathBounds.add(point);
                current = point;
            }
            if (!recordBounds.matches(emf, offset + 8)) {
                return false;
            }
            figureDrawable = true;
            pathDrawable = true;
            return true;
        }

        private boolean acceptFillPath(byte[] emf, int offset, int type, int size) {
            if (type != EMR_FILL_PATH || size != 24 || !pathBounds.matches(emf, offset + 8)) {
                return false;
            }
            phase = ClassicPhase.EXPECT_NULL_BRUSH;
            return true;
        }

        private boolean acceptDeleteBrush(byte[] emf, int offset, int type, int size) {
            if (type != EMR_DELETE_OBJECT || size != 12
                    || littleEndianInt(emf, offset + 8) != CLASSIC_BRUSH_HANDLE) {
                return false;
            }
            pathCount++;
            phase = ClassicPhase.READY;
            return true;
        }

        private boolean acceptEof(byte[] emf, int offset, int type, int size,
                                  boolean plusComplete) {
            if (!plusComplete || type != EMR_EOF || size != 20
                    || littleEndianInt(emf, offset + 8) != 0
                    || littleEndianInt(emf, offset + 12) != 16
                    || littleEndianInt(emf, offset + 16) != 20) {
                return false;
            }
            phase = ClassicPhase.COMPLETE;
            return true;
        }

        boolean isComplete() {
            return phase == ClassicPhase.COMPLETE;
        }

        int pathCount() {
            return pathCount;
        }

        List<Integer> rgbColors() {
            return rgbColors;
        }
    }

    private static final class EmfPlusValidation {
        private final boolean[] livePaths = new boolean[64];
        private final List<Integer> rgbColors = new ArrayList<>();
        private PlusPhase phase = PlusPhase.EXPECT_HEADER;
        private int pendingPathId = -1;
        private int pathCount;

        PlusCommentKind acceptComment(byte[] emf, int offset, int size) {
            if (size < 28 || littleEndianInt(emf, offset) != EMR_GDICOMMENT
                    || littleEndianInt(emf, offset + 8) != size - 12
                    || littleEndianInt(emf, offset + 12) != EMF_PLUS_SIGNATURE) {
                throw new IllegalArgumentException("invalid EMF+ comment bounds");
            }
            PlusPhase start = phase;
            List<Integer> types = new ArrayList<>(2);
            int cursor = offset + 16;
            int end = offset + size;
            while (cursor < end) {
                if (end - cursor < 12) {
                    throw new IllegalArgumentException("truncated EMF+ record header");
                }
                int type = littleEndianUnsignedShort(emf, cursor);
                int flags = littleEndianUnsignedShort(emf, cursor + 2);
                int plusSize = littleEndianInt(emf, cursor + 4);
                int plusDataSize = littleEndianInt(emf, cursor + 8);
                PlusPhase recordPhase = phase;
                if (plusSize < 12 || (plusSize & 3) != 0 || plusSize > end - cursor
                        || plusDataSize != plusSize - 12 || !isAllowedEmfPlusRecord(type)
                        || !acceptRecord(emf, cursor, type, flags, plusSize, plusDataSize)) {
                    throw new IllegalArgumentException("invalid EMF+ record 0x"
                        + Integer.toHexString(type) + " in phase " + recordPhase
                        + " at " + cursor);
                }
                types.add(type);
                cursor += plusSize;
            }
            if (cursor != end) {
                throw new IllegalArgumentException("EMF+ comment overrun");
            }
            return commentKind(start, types);
        }

        private boolean acceptRecord(byte[] emf, int offset, int type, int flags,
                                     int size, int dataSize) {
            return switch (phase) {
                case EXPECT_HEADER -> acceptHeader(emf, offset, type, flags, size, dataSize);
                case EXPECT_ANTI_ALIAS -> acceptProperty(
                    type, flags, size, dataSize, EMF_PLUS_SET_ANTI_ALIAS_MODE,
                    SMOOTHING_MODE_ANTI_ALIAS_8X8, PlusPhase.EXPECT_PIXEL_OFFSET);
                case EXPECT_PIXEL_OFFSET -> acceptProperty(
                    type, flags, size, dataSize, EMF_PLUS_SET_PIXEL_OFFSET_MODE,
                    PIXEL_OFFSET_MODE_HIGH_QUALITY, PlusPhase.EXPECT_OBJECT_OR_EOF);
                case EXPECT_OBJECT_OR_EOF -> type == EMF_PLUS_OBJECT
                    ? acceptPathObject(emf, offset, flags, size, dataSize)
                    : acceptPlusEof(type, flags, size, dataSize);
                case EXPECT_FILL -> acceptFill(emf, offset, type, flags, size, dataSize);
                case COMPLETE -> false;
            };
        }

        private boolean acceptHeader(byte[] emf, int offset, int type, int flags,
                                     int size, int dataSize) {
            if (type != EMF_PLUS_HEADER || flags != EMF_PLUS_DUAL || size != 28 || dataSize != 16
                    || littleEndianInt(emf, offset + 12) != GRAPHICS_VERSION_1_1
                    || littleEndianInt(emf, offset + 16) != EMF_PLUS_VIDEO_DISPLAY
                    || littleEndianInt(emf, offset + 20) != EMF_PLUS_LOGICAL_DPI
                    || littleEndianInt(emf, offset + 24) != EMF_PLUS_LOGICAL_DPI) {
                return false;
            }
            phase = PlusPhase.EXPECT_ANTI_ALIAS;
            return true;
        }

        private boolean acceptProperty(int type, int flags, int size, int dataSize,
                                       int expectedType, int expectedFlags, PlusPhase next) {
            if (type != expectedType || flags != expectedFlags || size != 12 || dataSize != 0) {
                return false;
            }
            phase = next;
            return true;
        }

        private boolean acceptPathObject(byte[] emf, int offset, int flags,
                                         int size, int dataSize) {
            int objectId = flags & 0xFF;
            int objectType = (flags >>> 8) & 0x7F;
            if ((flags & 0x8000) != 0 || objectType != OBJECT_TYPE_PATH
                    || objectId >= livePaths.length || livePaths[objectId]
                    || dataSize < 24
                    || littleEndianInt(emf, offset + 12) != GRAPHICS_VERSION_1_1
                    || littleEndianInt(emf, offset + 20) != 0) {
                return false;
            }
            int pointCount = littleEndianInt(emf, offset + 16);
            long expectedDataSize = (12L + 9L * pointCount + 3L) & ~3L;
            if (pointCount <= 0 || expectedDataSize != dataSize
                    || !hasValidPointPayload(emf, offset, size, pointCount)) {
                return false;
            }
            livePaths[objectId] = true;
            pendingPathId = objectId;
            phase = PlusPhase.EXPECT_FILL;
            return true;
        }

        private boolean hasValidPointPayload(byte[] emf, int offset, int size, int pointCount) {
            int pointsOffset = offset + 24;
            for (int index = 0; index < pointCount; index++) {
                float x = Float.intBitsToFloat(
                    littleEndianInt(emf, pointsOffset + index * 8));
                float y = Float.intBitsToFloat(
                    littleEndianInt(emf, pointsOffset + index * 8 + 4));
                if (!Float.isFinite(x) || !Float.isFinite(y)) {
                    return false;
                }
            }
            int typesOffset = pointsOffset + pointCount * 8;
            if (!hasValidPointTypes(emf, typesOffset, pointCount)) {
                return false;
            }
            int paddingOffset = typesOffset + pointCount;
            for (int index = paddingOffset; index < offset + size; index++) {
                if (emf[index] != 0) {
                    return false;
                }
            }
            return true;
        }

        private boolean hasValidPointTypes(byte[] emf, int offset, int count) {
            boolean figureActive = false;
            boolean figureDrawable = false;
            boolean anyDrawable = false;
            int index = 0;
            while (index < count) {
                int type = emf[offset + index] & 0xFF;
                int kind = type & 0x0F;
                if ((type & 0x70) != 0) {
                    return false;
                }
                if (kind == PATH_POINT_START) {
                    if ((type & PATH_POINT_CLOSE_SUBPATH) != 0
                            || (figureActive && !figureDrawable)) {
                        return false;
                    }
                    figureActive = true;
                    figureDrawable = false;
                    index++;
                    continue;
                }
                if (kind == PATH_POINT_LINE) {
                    if (!figureActive) {
                        return false;
                    }
                    figureDrawable = true;
                    anyDrawable = true;
                    if ((type & PATH_POINT_CLOSE_SUBPATH) != 0) {
                        figureActive = false;
                    }
                    index++;
                    continue;
                }
                if (kind != PATH_POINT_BEZIER || !figureActive) {
                    return false;
                }
                int runStart = index;
                while (index < count
                        && (emf[offset + index] & 0x0F) == PATH_POINT_BEZIER) {
                    int curveType = emf[offset + index] & 0xFF;
                    if ((curveType & 0x70) != 0) {
                        return false;
                    }
                    index++;
                }
                int runLength = index - runStart;
                if (runLength % 3 != 0) {
                    return false;
                }
                for (int curve = runStart; curve < index - 1; curve++) {
                    if ((emf[offset + curve] & PATH_POINT_CLOSE_SUBPATH) != 0) {
                        return false;
                    }
                }
                figureDrawable = true;
                anyDrawable = true;
                if ((emf[offset + index - 1] & PATH_POINT_CLOSE_SUBPATH) != 0) {
                    figureActive = false;
                }
            }
            return anyDrawable && (!figureActive || figureDrawable);
        }

        private boolean acceptFill(byte[] emf, int offset, int type, int flags,
                                   int size, int dataSize) {
            int objectId = flags & 0xFF;
            if (type != EMF_PLUS_FILL_PATH || size != 16 || dataSize != 4
                    || (flags & 0xFF00) != SOLID_COLOR_FLAG
                    || objectId >= livePaths.length || objectId != pendingPathId
                    || !livePaths[objectId]) {
                return false;
            }
            rgbColors.add(littleEndianInt(emf, offset + 12) & 0x00FFFFFF);
            livePaths[objectId] = false;
            pendingPathId = -1;
            pathCount++;
            phase = PlusPhase.EXPECT_OBJECT_OR_EOF;
            return true;
        }

        private boolean acceptPlusEof(int type, int flags, int size, int dataSize) {
            if (type != EMF_PLUS_EOF || flags != 0 || size != 12 || dataSize != 0
                    || pendingPathId != -1) {
                return false;
            }
            phase = PlusPhase.COMPLETE;
            return true;
        }

        private PlusCommentKind commentKind(PlusPhase start, List<Integer> types) {
            if (start == PlusPhase.EXPECT_HEADER && types.equals(List.of(EMF_PLUS_HEADER))) {
                return PlusCommentKind.HEADER;
            }
            if (start == PlusPhase.EXPECT_ANTI_ALIAS
                    && types.equals(List.of(EMF_PLUS_SET_ANTI_ALIAS_MODE))) {
                return PlusCommentKind.ANTI_ALIAS;
            }
            if (start == PlusPhase.EXPECT_PIXEL_OFFSET
                    && types.equals(List.of(EMF_PLUS_SET_PIXEL_OFFSET_MODE))) {
                return PlusCommentKind.PIXEL_OFFSET;
            }
            if (start == PlusPhase.EXPECT_OBJECT_OR_EOF
                    && types.equals(List.of(EMF_PLUS_OBJECT, EMF_PLUS_FILL_PATH))) {
                return PlusCommentKind.PATH_PAIR;
            }
            if (start == PlusPhase.EXPECT_OBJECT_OR_EOF
                    && types.equals(List.of(EMF_PLUS_EOF))) {
                return PlusCommentKind.EOF;
            }
            throw new IllegalArgumentException("unexpected EMF+ comment grouping");
        }

        boolean isComplete() {
            return phase == PlusPhase.COMPLETE;
        }

        int pathCount() {
            return pathCount;
        }

        List<Integer> rgbColors() {
            return rgbColors;
        }
    }

    private static final class BoundsAccumulator {
        private int left = Integer.MAX_VALUE;
        private int top = Integer.MAX_VALUE;
        private int right = Integer.MIN_VALUE;
        private int bottom = Integer.MIN_VALUE;

        void add(IntPoint point) {
            left = Math.min(left, point.x());
            top = Math.min(top, point.y());
            right = Math.max(right, point.x());
            bottom = Math.max(bottom, point.y());
        }

        boolean isEmpty() {
            return left == Integer.MAX_VALUE;
        }

        boolean matches(byte[] emf, int offset) {
            return !isEmpty()
                && littleEndianInt(emf, offset) == left
                && littleEndianInt(emf, offset + 4) == top
                && littleEndianInt(emf, offset + 8) == right
                && littleEndianInt(emf, offset + 12) == bottom;
        }
    }

    private static boolean isBitmapRecord(int type) {
        return (type >= 76 && type <= 81) || type == 114 || type == 116;
    }

    private static boolean isTextRecord(int type) {
        return type == 83 || type == 84;
    }

    private static PathData toPathData(Shape shape) {
        List<PathPoint> points = new ArrayList<>();
        PathIterator iterator = shape.getPathIterator(null);
        double[] coordinates = new double[6];
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        int figureStart = -1;

        while (!iterator.isDone()) {
            int segment = iterator.currentSegment(coordinates);
            switch (segment) {
                case PathIterator.SEG_MOVETO -> {
                    removeEmptyFigure(points, figureStart);
                    startX = coordinates[0];
                    startY = coordinates[1];
                    currentX = startX;
                    currentY = startY;
                    figureStart = points.size();
                    points.add(point(startX, startY, PATH_POINT_START));
                }
                case PathIterator.SEG_LINETO -> {
                    requireFigure(figureStart);
                    currentX = coordinates[0];
                    currentY = coordinates[1];
                    points.add(point(currentX, currentY, PATH_POINT_LINE));
                }
                case PathIterator.SEG_QUADTO -> {
                    requireFigure(figureStart);
                    double endX = coordinates[2];
                    double endY = coordinates[3];
                    points.add(point(
                        currentX + (coordinates[0] - currentX) * 2d / 3d,
                        currentY + (coordinates[1] - currentY) * 2d / 3d,
                        PATH_POINT_BEZIER));
                    points.add(point(
                        endX + (coordinates[0] - endX) * 2d / 3d,
                        endY + (coordinates[1] - endY) * 2d / 3d,
                        PATH_POINT_BEZIER));
                    points.add(point(endX, endY, PATH_POINT_BEZIER));
                    currentX = endX;
                    currentY = endY;
                }
                case PathIterator.SEG_CUBICTO -> {
                    requireFigure(figureStart);
                    points.add(point(coordinates[0], coordinates[1], PATH_POINT_BEZIER));
                    points.add(point(coordinates[2], coordinates[3], PATH_POINT_BEZIER));
                    points.add(point(coordinates[4], coordinates[5], PATH_POINT_BEZIER));
                    currentX = coordinates[4];
                    currentY = coordinates[5];
                }
                case PathIterator.SEG_CLOSE -> {
                    requireFigure(figureStart);
                    if (points.size() == figureStart + 1) {
                        points.remove(points.size() - 1);
                    } else {
                        int last = points.size() - 1;
                        PathPoint point = points.get(last);
                        points.set(last, point.withType(point.type() | PATH_POINT_CLOSE_SUBPATH));
                    }
                    currentX = startX;
                    currentY = startY;
                    figureStart = -1;
                }
                default -> throw new IllegalArgumentException("unknown Java2D path segment: " + segment);
            }
            iterator.next();
        }
        removeEmptyFigure(points, figureStart);
        return new PathData(List.copyOf(points));
    }

    private static void removeEmptyFigure(List<PathPoint> points, int figureStart) {
        if (figureStart >= 0 && points.size() == figureStart + 1) {
            points.remove(points.size() - 1);
        }
    }

    private static void requireFigure(int figureStart) {
        if (figureStart < 0) {
            throw new IllegalArgumentException("path segment appears before a start point");
        }
    }

    private static PathPoint point(double x, double y, int type) {
        if (!Double.isFinite(x) || !Double.isFinite(y)
                || x < -Float.MAX_VALUE || x > Float.MAX_VALUE
                || y < -Float.MAX_VALUE || y > Float.MAX_VALUE) {
            throw new IllegalArgumentException(
                "path coordinate is outside the finite float range: " + x + "," + y);
        }
        float floatX = (float) x;
        float floatY = (float) y;
        return new PathPoint(floatX, floatY, type);
    }

    private static byte[] emfPlusHeader(int dpiX, int dpiY) {
        LittleEndianWriter data = new LittleEndianWriter(16);
        data.writeInt(GRAPHICS_VERSION_1_1);
        data.writeInt(EMF_PLUS_VIDEO_DISPLAY);
        data.writeInt(dpiX);
        data.writeInt(dpiY);
        return emfPlusRecord(EMF_PLUS_HEADER, EMF_PLUS_DUAL, data.toByteArray());
    }

    private static byte[] emfPlusPathObject(PathData path) {
        int pointCount = path.points().size();
        int dataSize = checkedAlignedRecordCapacity(
            "EMF+ Path object data", 12L + (long) pointCount * 9L);
        LittleEndianWriter data = new LittleEndianWriter(dataSize);
        data.writeInt(GRAPHICS_VERSION_1_1);
        data.writeInt(pointCount);
        data.writeInt(0); // Absolute EmfPlusPointF coordinates; uncompressed point types.
        for (PathPoint point : path.points()) {
            data.writeFloat(point.x());
            data.writeFloat(point.y());
        }
        for (PathPoint point : path.points()) {
            data.writeByte(point.type());
        }
        data.padTo4();
        int flags = (OBJECT_TYPE_PATH << 8) | PATH_OBJECT_ID;
        return emfPlusRecord(EMF_PLUS_OBJECT, flags, data.toByteArray());
    }

    private static byte[] emfPlusFillPath(Color color) {
        LittleEndianWriter data = new LittleEndianWriter(4);
        data.writeInt(color.getRGB());
        return emfPlusRecord(EMF_PLUS_FILL_PATH,
            SOLID_COLOR_FLAG | PATH_OBJECT_ID, data.toByteArray());
    }

    private static byte[] emfPlusRecord(int type, int flags, byte[] data) {
        if ((data.length & 3) != 0) {
            throw new IllegalArgumentException("EMF+ record data is not 32-bit aligned");
        }
        int recordSize = checkedRecordCapacity("EMF+ record", 12L + data.length);
        LittleEndianWriter record = new LittleEndianWriter(recordSize);
        record.writeShort(type);
        record.writeShort(flags);
        record.writeInt(recordSize);
        record.writeInt(data.length);
        record.writeBytes(data);
        return record.toByteArray();
    }

    private static byte[] emfPlusComment(byte[]... records) {
        long recordsSize = 0L;
        for (byte[] record : records) {
            if ((record.length & 3) != 0) {
                throw new IllegalArgumentException("embedded EMF+ record is not 32-bit aligned");
            }
            recordsSize += record.length;
            checkedRecordCapacity("embedded EMF+ records", recordsSize);
        }
        int dataSize = checkedRecordCapacity("EMF+ comment data", 4L + recordsSize);
        int recordSize = checkedRecordCapacity("EMR_GDICOMMENT", 12L + dataSize);
        LittleEndianWriter comment = new LittleEndianWriter(recordSize);
        comment.writeInt(EMR_GDICOMMENT);
        comment.writeInt(recordSize);
        comment.writeInt(dataSize);
        comment.writeInt(EMF_PLUS_SIGNATURE);
        for (byte[] record : records) {
            comment.writeBytes(record);
        }
        return comment.toByteArray();
    }

    private static ClassicEmfLayout inspectClassicEmf(byte[] emf)
        throws SvgVectorWmfRenderer.SvgVectorWmfException {
        if (emf.length < 88 || littleEndianInt(emf, 0) != EMR_HEADER
                || littleEndianInt(emf, 40) != ENHMETA_SIGNATURE) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException("invalid classic EMF header");
        }
        int headerSize = littleEndianInt(emf, 4);
        if (headerSize < 88 || (headerSize & 3) != 0 || headerSize > emf.length) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                "invalid classic EMF header size: " + headerSize);
        }

        int offset = 0;
        int eofOffset = -1;
        while (offset < emf.length) {
            if (emf.length - offset < 8) {
                throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                    "truncated classic EMF record header at " + offset);
            }
            int type = littleEndianInt(emf, offset);
            int size = littleEndianInt(emf, offset + 4);
            if (size < 8 || (size & 3) != 0 || size > emf.length - offset) {
                throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                    "invalid classic EMF record size at " + offset + ": " + size);
            }
            if (type == EMR_EOF) {
                if (offset + size != emf.length) {
                    throw new SvgVectorWmfRenderer.SvgVectorWmfException(
                        "classic EMF has data after EMR_EOF");
                }
                eofOffset = offset;
            }
            offset += size;
        }
        if (eofOffset < 0) {
            throw new SvgVectorWmfRenderer.SvgVectorWmfException("classic EMF has no EMR_EOF");
        }
        return new ClassicEmfLayout(headerSize, eofOffset);
    }

    private static boolean isAllowedClassicRecord(int type) {
        return switch (type) {
            case EMR_SET_WINDOW_EXT, EMR_SET_VIEWPORT_EXT, EMR_EOF, EMR_SET_MAP_MODE,
                EMR_SET_POLY_FILL_MODE, EMR_MOVE_TO_EX, EMR_SELECT_OBJECT,
                EMR_CREATE_BRUSH_INDIRECT, EMR_DELETE_OBJECT, EMR_LINE_TO,
                EMR_BEGIN_PATH, EMR_END_PATH, EMR_CLOSE_FIGURE, EMR_FILL_PATH,
                EMR_POLYBEZIER_TO, EMR_GDICOMMENT -> true;
            default -> false;
        };
    }

    private static boolean isAllowedEmfPlusRecord(int type) {
        return switch (type) {
            case EMF_PLUS_HEADER, EMF_PLUS_EOF, EMF_PLUS_OBJECT, EMF_PLUS_FILL_PATH,
                EMF_PLUS_SET_ANTI_ALIAS_MODE, EMF_PLUS_SET_PIXEL_OFFSET_MODE -> true;
            default -> false;
        };
    }

    private static int countEmfRecords(byte[] emf) {
        int count = 0;
        int offset = 0;
        while (offset < emf.length) {
            if (emf.length - offset < 8) {
                throw new IllegalArgumentException("truncated EMF record header at " + offset);
            }
            int size = littleEndianInt(emf, offset + 4);
            if (size < 8 || (size & 3) != 0 || size > emf.length - offset) {
                throw new IllegalArgumentException("invalid EMF record size at " + offset + ": " + size);
            }
            if (count == Integer.MAX_VALUE) {
                throw new IllegalArgumentException(
                    "EMF record count exceeds signed 32-bit range");
            }
            count++;
            offset += size;
        }
        return count;
    }

    private static int checkedAlignedRecordCapacity(String description, long value) {
        if (value < 0L || value > Long.MAX_VALUE - 3L) {
            throw new IllegalArgumentException(description + " capacity overflow: " + value);
        }
        long aligned = (value + 3L) & ~3L;
        return checkedRecordCapacity(description, aligned);
    }

    private static int littleEndianInt(byte[] bytes, int offset) {
        return (bytes[offset] & 0xFF)
            | ((bytes[offset + 1] & 0xFF) << 8)
            | ((bytes[offset + 2] & 0xFF) << 16)
            | ((bytes[offset + 3] & 0xFF) << 24);
    }

    private static int littleEndianUnsignedShort(byte[] bytes, int offset) {
        return (bytes[offset] & 0xFF) | ((bytes[offset + 1] & 0xFF) << 8);
    }

    private static void putLittleEndianInt(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte) value;
        bytes[offset + 1] = (byte) (value >>> 8);
        bytes[offset + 2] = (byte) (value >>> 16);
        bytes[offset + 3] = (byte) (value >>> 24);
    }

    private record ClassicEmfLayout(int headerSize, int eofOffset) {
    }

    private record RenderGeometry(int classicDeviceWidth, int classicDeviceHeight,
                                  int plusViewportWidth, int plusViewportHeight,
                                  AffineTransform classicTransform,
                                  AffineTransform plusTransform) {
    }

    private record EncodedPath(PathData path, Color color, byte[] comment) {
    }

    private record PathData(List<PathPoint> points) {
    }

    private record PathPoint(float x, float y, int type) {
        PathPoint withType(int replacement) {
            return new PathPoint(x, y, replacement);
        }
    }

    private record IntPoint(int x, int y) {
    }

    private record Bounds(int left, int top, int right, int bottom) {
    }

    private static final class LittleEndianWriter {
        private final ByteArrayOutputStream output;

        LittleEndianWriter(int expectedSize) {
            output = new ByteArrayOutputStream(
                checkedRecordCapacity("little-endian record", expectedSize));
        }

        void writeByte(int value) {
            output.write(value);
        }

        void writeShort(int value) {
            output.write(value);
            output.write(value >>> 8);
        }

        void writeInt(int value) {
            output.write(value);
            output.write(value >>> 8);
            output.write(value >>> 16);
            output.write(value >>> 24);
        }

        void writeFloat(float value) {
            writeInt(Float.floatToRawIntBits(value));
        }

        void writeRect(int left, int top, int right, int bottom) {
            writeInt(left);
            writeInt(top);
            writeInt(right);
            writeInt(bottom);
        }

        void writeRect(Bounds bounds) {
            writeRect(bounds.left(), bounds.top(), bounds.right(), bounds.bottom());
        }

        void writeBytes(byte[] bytes) {
            output.writeBytes(bytes);
        }

        void padTo4() {
            while ((output.size() & 3) != 0) {
                output.write(0);
            }
        }

        int size() {
            return output.size();
        }

        byte[] toByteArray() {
            return output.toByteArray();
        }
    }
}
