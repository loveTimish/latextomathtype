package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.awt.image.BufferedImage;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class SvgVectorEmfPlusRendererTest {

    private static final int EMR_HEADER = 0x00000001;
    private static final int EMR_POLYBEZIER_TO = 0x00000005;
    private static final int EMR_SET_WINDOW_EXT = 0x00000009;
    private static final int EMR_SET_VIEWPORT_EXT = 0x0000000B;
    private static final int EMR_EOF = 0x0000000E;
    private static final int EMR_SET_MAP_MODE = 0x00000011;
    private static final int EMR_SELECT_OBJECT = 0x00000025;
    private static final int EMR_CREATE_BRUSH_INDIRECT = 0x00000027;
    private static final int EMR_FILLPATH = 0x0000003E;
    private static final int EMR_GDICOMMENT = 0x00000046;
    private static final int EMF_PLUS_SIGNATURE = 0x2B464D45;

    private static final String COLORED_CURVE_SVG =
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24pt\" height=\"12pt\" "
            + "viewBox=\"0 0 240 120\">"
            + "<path fill=\"#123456\" d=\"M10 100Q40 5 80 70C100 115 130 5 160 70L170 100Z\"/>"
            + "<rect fill=\"#d02030\" x=\"185\" y=\"20\" width=\"40\" height=\"80\"/>"
            + "</svg>";

    @Test
    void shouldEmitStructurallyCompleteEmfPlusDualWithClassicFallback() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(emf));
        List<EmfRecord> emfRecords = emfRecords(emf);
        List<PlusRecord> plusRecords = plusRecords(emf, emfRecords);

        assertEquals(EMR_HEADER, emfRecords.getFirst().type());
        assertEquals(EMR_EOF, emfRecords.getLast().type());
        assertEquals(108, emfRecords.getFirst().size());
        assertEquals(847, u32(emf, 32), "24pt frame width must be expressed in 0.01mm units");
        assertEquals(424, u32(emf, 36), "12pt frame height must be expressed in 0.01mm units");
        assertEquals(2, u16(emf, 56));
        assertEquals(emf.length, u32(emf, 48), "EMF header nBytes must include injected comments");
        assertEquals(emfRecords.size(), u32(emf, 52),
            "EMF header nRecords must include injected comment records");

        int headerSize = emfRecords.getFirst().size();
        EmfRecord firstAfterHeader = emfRecords.get(1);
        assertEquals(headerSize, firstAfterHeader.offset());
        assertEquals(EMR_GDICOMMENT, firstAfterHeader.type());
        assertEquals(0x4001, plusRecords.getFirst().type(),
            "EmfPlusHeader must be the first record after EMR_HEADER");
        assertEquals(0x0001, plusRecords.getFirst().flags(), "Header must mark EMF+ Dual");
        assertEquals(28, plusRecords.getFirst().size());
        assertEquals(16, plusRecords.getFirst().dataSize());
        assertEquals(0xDBC01002, u32(emf, plusRecords.getFirst().offset() + 12));
        assertEquals(1, u32(emf, plusRecords.getFirst().offset() + 16));
        assertEquals(96, u32(emf, plusRecords.getFirst().offset() + 20));
        assertEquals(96, u32(emf, plusRecords.getFirst().offset() + 24));

        assertEquals(0x401E, plusRecords.get(1).type());
        assertEquals(0x000B, plusRecords.get(1).flags(),
            "A=1 and SmoothingModeAntiAlias8x8=5 must encode as 0x000B");
        assertEquals(12, plusRecords.get(1).size());
        assertEquals(0, plusRecords.get(1).dataSize());

        assertEquals(0x4022, plusRecords.get(2).type());
        assertEquals(0x0002, plusRecords.get(2).flags(),
            "PixelOffsetModeHighQuality must use half-pixel sampling centers");
        assertEquals(12, plusRecords.get(2).size());
        assertEquals(0, plusRecords.get(2).dataSize());

        assertEquals(List.of(0x4001, 0x401E, 0x4022, 0x4008, 0x4014, 0x4008, 0x4014, 0x4002),
            plusRecords.stream().map(PlusRecord::type).toList());
        assertEquals(0x4002, plusRecords.getLast().type());
        assertEquals(12, plusRecords.getLast().size());
        assertEquals(0, plusRecords.getLast().dataSize());

        EmfRecord beforeEof = emfRecords.get(emfRecords.size() - 2);
        assertEquals(EMR_GDICOMMENT, beforeEof.type());
        assertEquals(EMF_PLUS_SIGNATURE, u32(emf, beforeEof.offset() + 12));
        assertEquals(0x4002, u16(emf, beforeEof.offset() + 16),
            "EmfPlusEOF comment must immediately precede EMR_EOF");

        assertTrue(emfRecords.stream().anyMatch(record -> record.type() == EMR_FILLPATH),
            "classic EMF fallback must retain vector fill records");
        assertTrue(emfRecords.stream().anyMatch(record -> record.type() == 0x00000005),
            "classic fallback must retain cubic EMR_POLYBEZIERTO records");
        assertTrue(emfRecords.stream().anyMatch(record -> record.type() == 0x00000011),
            "classic high-resolution coordinates require MM_ANISOTROPIC");
        assertTrue(emfRecords.stream().anyMatch(record -> record.type() == 0x00000009));
        assertTrue(emfRecords.stream().anyMatch(record -> record.type() == 0x0000000B));
        assertFalse(emfRecords.stream().anyMatch(record -> isBitmapRecord(record.type())),
            "classic fallback must not contain raster records");
        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(emf));
    }

    @Test
    void shouldUseOneExactClassicDeviceScaleForFractionalTargetSizes() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 17.3d, 7.7d);
        List<EmfRecord> records = emfRecords(emf);
        EmfRecord window = firstRecord(records, EMR_SET_WINDOW_EXT);
        EmfRecord viewport = firstRecord(records, EMR_SET_VIEWPORT_EXT);

        int boundsWidth = Math.addExact(u32(emf, 16), 1);
        int boundsHeight = Math.addExact(u32(emf, 20), 1);
        int frameWidth = u32(emf, 32);
        int frameHeight = u32(emf, 36);
        int windowWidth = u32(emf, window.offset() + 8);
        int windowHeight = u32(emf, window.offset() + 12);
        int viewportWidth = u32(emf, viewport.offset() + 8);
        int viewportHeight = u32(emf, viewport.offset() + 12);

        assertEquals((int) Math.ceil(17.3d * 2540d / 72d), frameWidth,
            "rclFrame width must use 0.01 mm units");
        assertEquals((int) Math.ceil(7.7d * 2540d / 72d), frameHeight,
            "rclFrame height must use 0.01 mm units");
        assertEquals(frameWidth, boundsWidth,
            "rclBounds width must describe the classic device coordinate space");
        assertEquals(frameHeight, boundsHeight,
            "rclBounds height must describe the classic device coordinate space");
        assertEquals(windowWidth, viewportWidth,
            "classic X mapping must be exactly 1:1");
        assertEquals(windowHeight, viewportHeight,
            "classic Y mapping must be exactly 1:1");
        assertEquals(boundsWidth, viewportWidth);
        assertEquals(boundsHeight, viewportHeight);
        assertEquals((long) windowWidth * viewportHeight,
            (long) windowHeight * viewportWidth,
            "classic X/Y scale factors must be identical");
        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(emf));
    }

    @Test
    void independentDecoderShouldRasterizeSerializedPathsAtPhysicalScale() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24pt\" height=\"12pt\" "
            + "viewBox=\"0 0 240 120\"><rect fill=\"#123456\" "
            + "x=\"24\" y=\"24\" width=\"96\" height=\"48\"/></svg>";
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            svg.getBytes(StandardCharsets.UTF_8), 24d, 12d);

        EmfPlusPreviewInspector.RasterizedEmfPlus raster =
            EmfPlusPreviewInspector.rasterize(emf, 24d, 12d, 300);
        BufferedImage image = raster.image();

        assertEquals(100, image.getWidth());
        assertEquals(50, image.getHeight());
        assertEquals(1, raster.pathFillCount());
        assertEquals(0xFF123456, image.getRGB(30, 20));
        assertEquals(0, image.getRGB(2, 2) >>> 24);
        assertEquals(java.util.Set.of(0xFF123456), raster.colors());
    }

    @Test
    void shouldSerializeAbsoluteCurvesCloseFlagsAndSolidArgbColors() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        List<PlusRecord> records = plusRecords(emf, emfRecords(emf));
        List<PlusRecord> objects = records.stream().filter(record -> record.type() == 0x4008).toList();
        List<PlusRecord> fills = records.stream().filter(record -> record.type() == 0x4014).toList();

        assertEquals(2, objects.size());
        assertEquals(2, fills.size());
        PlusRecord curve = objects.getFirst();
        assertEquals(0x0300, curve.flags(), "ObjectTypePath=3 and ObjectId=0");
        assertEquals(0xDBC01002, u32(emf, curve.offset() + 12));
        int pointCount = u32(emf, curve.offset() + 16);
        assertTrue(pointCount >= 7);
        assertEquals(0, u32(emf, curve.offset() + 20),
            "PathPointFlags=0 means absolute PointF and uncompressed point types");
        int pointTypes = curve.offset() + 24 + pointCount * 8;
        boolean hasBezier = false;
        boolean hasClose = false;
        for (int index = 0; index < pointCount; index++) {
            int type = emf[pointTypes + index] & 0xFF;
            hasBezier |= (type & 0x0F) == 0x03;
            hasClose |= (type & 0x80) != 0;
        }
        assertTrue(hasBezier, "quadratic/cubic SVG segments must remain cubic EMF+ Beziers");
        assertTrue(hasClose, "closed SVG contours must carry PathPointTypeCloseSubpath");
        assertEquals(0, curve.size() & 3);
        assertEquals(curve.size() - 12, curve.dataSize());

        assertEquals(0x8000, fills.getFirst().flags(),
            "S=1 selects inline solid ARGB; ObjectId remains zero");
        assertEquals(0xFF123456, u32(emf, fills.getFirst().offset() + 12));
        assertEquals(0xFFD02030, u32(emf, fills.get(1).offset() + 12));
        assertEquals(16, fills.getFirst().size());
        assertEquals(4, fills.getFirst().dataSize());
    }

    @Test
    void shouldCanonicalizeOverlapsAndPreserveRealHolesForAlternatePathFill() throws Exception {
        String svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40pt\" height=\"20pt\" "
                + "viewBox=\"0 0 400 200\">"
                + "<path fill=\"#000000\" fill-rule=\"nonzero\" "
                + "d=\"M10 10H150V150H10ZM100 10H240V150H100Z\"/>"
                + "<path fill=\"#202020\" fill-rule=\"evenodd\" "
                + "d=\"M260 10H390V150H260ZM290 40H360V120H290Z\"/>"
                + "</svg>";
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            svg.getBytes(StandardCharsets.UTF_8), 40d, 20d);
        List<PlusRecord> objects = plusRecords(emf, emfRecords(emf)).stream()
            .filter(record -> record.type() == 0x4008)
            .toList();

        assertEquals(2, objects.size());
        assertEquals(1, countPointType(emf, objects.getFirst(), 0x00, 0x0F),
            "overlapping nonzero contours must become one filled boundary before EMF+ alternate fill");
        assertEquals(2, countPointType(emf, objects.get(1), 0x00, 0x0F),
            "an actual even-odd hole must remain a separate subpath");
        assertEquals(2, countPointType(emf, objects.get(1), 0x80, 0x80),
            "both hole contours must remain explicitly closed");
    }

    @Test
    void shouldEmitLegalEmptyDualMetafileForAnEmptyScene() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"12pt\" height=\"6pt\" "
            + "viewBox=\"0 0 120 60\"/>";
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            svg.getBytes(StandardCharsets.UTF_8), 12d, 6d);
        List<EmfRecord> classic = emfRecords(emf);
        List<PlusRecord> plus = plusRecords(emf, classic);

        assertTrue(SvgVectorEmfPlusRenderer.isValidDualVector(emf));
        assertEquals(List.of(0x4001, 0x401E, 0x4022, 0x4002),
            plus.stream().map(PlusRecord::type).toList());
        assertEquals(1, u16(emf, 56),
            "an intentionally empty metafile must advertise no reusable brush handle");
        assertFalse(classic.stream().anyMatch(record -> record.type() == EMR_FILLPATH));
        assertFalse(classic.stream().anyMatch(record -> isBitmapRecord(record.type())));

        byte[] falselyNonempty = Arrays.copyOf(emf, emf.length);
        putU16(falselyNonempty, 56, 2);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(falselyNonempty),
            "the handle-table marker must distinguish a deliberate empty scene");
    }

    @Test
    void shouldRejectInvalidInputInsteadOfProducingMalformedMetafile() {
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorEmfPlusRenderer.render(new byte[0], 10d, 10d));
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorEmfPlusRenderer.render(
                COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), Double.NaN, 10d));
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorEmfPlusRenderer.render(
                "<not-svg/>".getBytes(StandardCharsets.UTF_8), 10d, 10d));
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorEmfPlusRenderer.render(
                COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 10d, 10d, 0.31d));
    }

    @Test
    void shouldRejectNumericAndRecordCapacityOverflowWithExplicitDiagnostics() {
        SvgVectorWmfRenderer.SvgVectorWmfException targetOverflow = assertThrows(
            SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorEmfPlusRenderer.render(
                COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 75_000_000d, 10d));
        assertTrue(targetOverflow.getMessage().contains(
            "classic EMF device width exceeds signed 32-bit range"));

        IllegalArgumentException coordinateOverflow = assertThrows(
            IllegalArgumentException.class,
            () -> SvgVectorEmfPlusRenderer.checkedRoundToInt(
                Float.MAX_VALUE, "classic EMF test coordinate"));
        assertTrue(coordinateOverflow.getMessage().contains(
            "classic EMF test coordinate exceeds signed 32-bit range"));

        IllegalArgumentException capacityOverflow = assertThrows(
            IllegalArgumentException.class,
            () -> SvgVectorEmfPlusRenderer.checkedRecordCapacity(
                "test EMF record", Integer.MAX_VALUE));
        assertTrue(capacityOverflow.getMessage().contains(
            "test EMF record capacity exceeds supported byte-array range"));
    }

    @Test
    void shouldCompensateSmallGlyphsWithoutThickeningStructuralRules() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"60pt\" height=\"12pt\" "
            + "viewBox=\"0 0 600 120\">"
            + "<rect x=\"40\" y=\"20\" width=\"70\" height=\"80\"/>"
            + "<rect x=\"160\" y=\"57\" width=\"360\" height=\"6\"/>"
            + "</svg>";
        byte[] plain = SvgVectorEmfPlusRenderer.render(
            svg.getBytes(StandardCharsets.UTF_8), 60d, 12d, 0d);
        byte[] compensated = SvgVectorEmfPlusRenderer.render(
            svg.getBytes(StandardCharsets.UTF_8), 60d, 12d, 0.20d);
        List<PlusRecord> plainObjects = pathObjects(plain);
        List<PlusRecord> compensatedObjects = pathObjects(compensated);

        FloatBounds plainGlyph = pathBounds(plain, plainObjects.getFirst());
        FloatBounds compensatedGlyph = pathBounds(compensated, compensatedObjects.getFirst());
        FloatBounds plainRule = pathBounds(plain, plainObjects.get(1));
        FloatBounds compensatedRule = pathBounds(compensated, compensatedObjects.get(1));

        assertTrue(compensatedGlyph.width() > plainGlyph.width());
        assertTrue(compensatedGlyph.height() > plainGlyph.height());
        assertEquals(plainRule.width(), compensatedRule.width(), 0.0001f);
        assertEquals(plainRule.height(), compensatedRule.height(), 0.0001f);
    }

    @Test
    void shouldPreserveExactOutlinesByDefault() {
        String property = SvgVectorEmfPlusRenderer.OPTICAL_COMPENSATION_PT_PROP;
        String previous = System.getProperty(property);
        try {
            System.clearProperty(property);
            assertEquals(0d, SvgVectorEmfPlusRenderer.opticalCompensationPt(), 0d);
        } finally {
            if (previous == null) {
                System.clearProperty(property);
            } else {
                System.setProperty(property, previous);
            }
        }
    }

    @Test
    void shouldRejectUnexpectedClassicAndEmfPlusRecords() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);

        byte[] unexpectedClassic = Arrays.copyOf(emf, emf.length);
        EmfRecord classicPathRecord = emfRecords(unexpectedClassic).stream()
            .filter(record -> record.type() == EMR_FILLPATH)
            .findFirst()
            .orElseThrow();
        putU32(unexpectedClassic, classicPathRecord.offset(), 0x00000018);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(unexpectedClassic));

        byte[] unexpectedPlus = Arrays.copyOf(emf, emf.length);
        PlusRecord pathObject = plusRecords(unexpectedPlus, emfRecords(unexpectedPlus)).stream()
            .filter(record -> record.type() == 0x4008)
            .findFirst()
            .orElseThrow();
        putU16(unexpectedPlus, pathObject.offset(), 0x4030);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(unexpectedPlus));
    }

    @Test
    void shouldRejectInexactCommentBoundsAndOutOfOrderProperties() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        List<EmfRecord> classic = emfRecords(emf);
        List<PlusRecord> plus = plusRecords(emf, classic);

        byte[] shortComment = Arrays.copyOf(emf, emf.length);
        PlusRecord object = plus.stream()
            .filter(record -> record.type() == 0x4008)
            .findFirst()
            .orElseThrow();
        EmfRecord objectComment = containingRecord(classic, object.offset());
        putU32(shortComment, objectComment.offset() + 8, 4);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(shortComment),
            "cbData must cover the complete comment, not merely a balanced record prefix");

        EmfRecord headerComment = classic.get(1);
        byte[] duplicateHeader = insertRecord(
            emf, headerComment.offset() + headerComment.size(), headerComment);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(duplicateHeader),
            "EmfPlusHeader must occur exactly once at the start of the plus stream");

        byte[] swappedProperties = Arrays.copyOf(emf, emf.length);
        PlusRecord antiAlias = plus.stream()
            .filter(record -> record.type() == 0x401E)
            .findFirst()
            .orElseThrow();
        PlusRecord pixelOffset = plus.stream()
            .filter(record -> record.type() == 0x4022)
            .findFirst()
            .orElseThrow();
        putU16(swappedProperties, antiAlias.offset(), pixelOffset.type());
        putU16(swappedProperties, antiAlias.offset() + 2, pixelOffset.flags());
        putU16(swappedProperties, pixelOffset.offset(), antiAlias.type());
        putU16(swappedProperties, pixelOffset.offset() + 2, antiAlias.flags());
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(swappedProperties),
            "anti-aliasing and pixel-offset records are unique and ordered");
    }

    @Test
    void shouldRejectInvalidPathObjectsPointPayloadsAndFillReferences() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        List<PlusRecord> plus = plusRecords(emf, emfRecords(emf));
        PlusRecord object = plus.stream()
            .filter(record -> record.type() == 0x4008)
            .findFirst()
            .orElseThrow();

        byte[] wrongObjectType = Arrays.copyOf(emf, emf.length);
        putU16(wrongObjectType, object.offset() + 2, 0x0200);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(wrongObjectType));

        byte[] danglingFillReference = Arrays.copyOf(emf, emf.length);
        putU16(danglingFillReference, object.offset() + 2, 0x0301);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(danglingFillReference),
            "FillPath must reference the live Path object declared immediately before it");

        byte[] impossiblePointCount = Arrays.copyOf(emf, emf.length);
        putU32(impossiblePointCount, object.offset() + 16, 0);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(impossiblePointCount));

        int pointCount = u32(emf, object.offset() + 16);
        int pointTypes = object.offset() + 24 + pointCount * 8;
        byte[] unknownPointType = Arrays.copyOf(emf, emf.length);
        unknownPointType[pointTypes + 1] = 0x02;
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(unknownPointType));

        byte[] nonFiniteCoordinate = Arrays.copyOf(emf, emf.length);
        putU32(nonFiniteCoordinate, object.offset() + 24, Float.floatToRawIntBits(Float.NaN));
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(nonFiniteCoordinate));
    }

    @Test
    void shouldRejectMalformedClassicBrushPathAndRecordLayouts() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        List<EmfRecord> records = emfRecords(emf);

        EmfRecord mapMode = firstRecord(records, EMR_SET_MAP_MODE);
        byte[] wrongMapMode = Arrays.copyOf(emf, emf.length);
        putU32(wrongMapMode, mapMode.offset() + 8, 7);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(wrongMapMode));

        EmfRecord viewport = firstRecord(records, EMR_SET_VIEWPORT_EXT);
        byte[] distortedViewport = Arrays.copyOf(emf, emf.length);
        putU32(distortedViewport, viewport.offset() + 8,
            u32(emf, viewport.offset() + 8) + 1);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(distortedViewport),
            "the classic fallback viewport must retain the renderer's physical scale");

        EmfRecord brush = firstRecord(records, EMR_CREATE_BRUSH_INDIRECT);
        byte[] patternedBrush = Arrays.copyOf(emf, emf.length);
        putU32(patternedBrush, brush.offset() + 12, 2);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(patternedBrush));

        EmfRecord select = firstRecord(records, EMR_SELECT_OBJECT);
        byte[] wrongBrushHandle = Arrays.copyOf(emf, emf.length);
        putU32(wrongBrushHandle, select.offset() + 8, 2);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(wrongBrushHandle));

        EmfRecord bezier = firstRecord(records, EMR_POLYBEZIER_TO);
        byte[] wrongBezierCount = Arrays.copyOf(emf, emf.length);
        putU32(wrongBezierCount, bezier.offset() + 24, u32(emf, bezier.offset() + 24) + 1);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(wrongBezierCount));

        EmfRecord fill = firstRecord(records, EMR_FILLPATH);
        byte[] wrongFillBounds = Arrays.copyOf(emf, emf.length);
        putU32(wrongFillBounds, fill.offset() + 8, u32(emf, fill.offset() + 8) + 1);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(wrongFillBounds));

        EmfRecord eof = records.getLast();
        byte[] wrongEofLayout = Arrays.copyOf(emf, emf.length);
        putU32(wrongEofLayout, eof.offset() + 12, 0);
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(wrongEofLayout));
    }

    @Test
    void shouldRejectAPathMissingFromOnlyOneSideOfTheDualMetafile() throws Exception {
        byte[] emf = SvgVectorEmfPlusRenderer.render(
            COLORED_CURVE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        List<EmfRecord> classic = emfRecords(emf);
        PlusRecord firstObject = plusRecords(emf, classic).stream()
            .filter(record -> record.type() == 0x4008)
            .findFirst()
            .orElseThrow();
        EmfRecord firstPathComment = containingRecord(classic, firstObject.offset());

        byte[] onePlusPathRemoved = removeRecord(emf, firstPathComment);
        assertEquals(1, plusRecords(onePlusPathRemoved, emfRecords(onePlusPathRemoved)).stream()
            .filter(record -> record.type() == 0x4008).count());
        assertEquals(1, plusRecords(onePlusPathRemoved, emfRecords(onePlusPathRemoved)).stream()
            .filter(record -> record.type() == 0x4014).count());
        assertFalse(SvgVectorEmfPlusRenderer.isValidDualVector(onePlusPathRemoved),
            "balanced EMF+ Object/Fill counts do not excuse a missing classic/plus path pair");
    }

    private static List<PlusRecord> pathObjects(byte[] emf) {
        return plusRecords(emf, emfRecords(emf)).stream()
            .filter(record -> record.type() == 0x4008)
            .toList();
    }

    private static EmfRecord firstRecord(List<EmfRecord> records, int type) {
        return records.stream()
            .filter(record -> record.type() == type)
            .findFirst()
            .orElseThrow();
    }

    private static EmfRecord containingRecord(List<EmfRecord> records, int nestedOffset) {
        return records.stream()
            .filter(record -> nestedOffset >= record.offset()
                && nestedOffset < record.offset() + record.size())
            .findFirst()
            .orElseThrow();
    }

    private static byte[] removeRecord(byte[] emf, EmfRecord removed) {
        byte[] result = new byte[emf.length - removed.size()];
        System.arraycopy(emf, 0, result, 0, removed.offset());
        System.arraycopy(emf, removed.offset() + removed.size(), result, removed.offset(),
            emf.length - removed.offset() - removed.size());
        putU32(result, 48, result.length);
        putU32(result, 52, u32(emf, 52) - 1);
        return result;
    }

    private static byte[] insertRecord(byte[] emf, int offset, EmfRecord source) {
        byte[] result = new byte[emf.length + source.size()];
        System.arraycopy(emf, 0, result, 0, offset);
        System.arraycopy(emf, source.offset(), result, offset, source.size());
        System.arraycopy(emf, offset, result, offset + source.size(), emf.length - offset);
        putU32(result, 48, result.length);
        putU32(result, 52, u32(emf, 52) + 1);
        return result;
    }

    private static FloatBounds pathBounds(byte[] emf, PlusRecord object) {
        int pointCount = u32(emf, object.offset() + 16);
        float minX = Float.POSITIVE_INFINITY;
        float minY = Float.POSITIVE_INFINITY;
        float maxX = Float.NEGATIVE_INFINITY;
        float maxY = Float.NEGATIVE_INFINITY;
        int pointsOffset = object.offset() + 24;
        for (int index = 0; index < pointCount; index++) {
            float x = Float.intBitsToFloat(u32(emf, pointsOffset + index * 8));
            float y = Float.intBitsToFloat(u32(emf, pointsOffset + index * 8 + 4));
            minX = Math.min(minX, x);
            minY = Math.min(minY, y);
            maxX = Math.max(maxX, x);
            maxY = Math.max(maxY, y);
        }
        return new FloatBounds(minX, minY, maxX, maxY);
    }

    private static List<EmfRecord> emfRecords(byte[] emf) {
        List<EmfRecord> records = new ArrayList<>();
        int offset = 0;
        while (offset < emf.length) {
            assertTrue(emf.length - offset >= 8, "truncated EMF record header at " + offset);
            int type = u32(emf, offset);
            int size = u32(emf, offset + 4);
            assertTrue(size >= 8 && (size & 3) == 0 && size <= emf.length - offset,
                "invalid EMF record size at " + offset + ": " + size);
            records.add(new EmfRecord(offset, type, size));
            offset += size;
        }
        assertEquals(emf.length, offset);
        return records;
    }

    private static List<PlusRecord> plusRecords(byte[] emf, List<EmfRecord> records) {
        List<PlusRecord> result = new ArrayList<>();
        for (EmfRecord record : records) {
            if (record.type() != EMR_GDICOMMENT || record.size() < 16
                    || u32(emf, record.offset() + 12) != EMF_PLUS_SIGNATURE) {
                continue;
            }
            int dataSize = u32(emf, record.offset() + 8);
            assertEquals(record.size() - 12, dataSize);
            int offset = record.offset() + 16;
            int end = record.offset() + 12 + dataSize;
            while (offset < end) {
                assertTrue(end - offset >= 12, "truncated EMF+ record header");
                int type = u16(emf, offset);
                int flags = u16(emf, offset + 2);
                int size = u32(emf, offset + 4);
                int plusDataSize = u32(emf, offset + 8);
                assertTrue(size >= 12 && (size & 3) == 0 && size <= end - offset);
                assertEquals(size - 12, plusDataSize);
                result.add(new PlusRecord(offset, type, flags, size, plusDataSize));
                offset += size;
            }
            assertEquals(end, offset);
        }
        return result;
    }

    private static boolean isBitmapRecord(int type) {
        return (type >= 76 && type <= 81) || type == 114 || type == 116;
    }

    private static int countPointType(byte[] emf, PlusRecord object, int expected, int mask) {
        int pointCount = u32(emf, object.offset() + 16);
        int typesOffset = object.offset() + 24 + pointCount * 8;
        int count = 0;
        for (int index = 0; index < pointCount; index++) {
            if ((emf[typesOffset + index] & mask) == expected) {
                count++;
            }
        }
        return count;
    }

    private static int u16(byte[] bytes, int offset) {
        return (bytes[offset] & 0xFF) | ((bytes[offset + 1] & 0xFF) << 8);
    }

    private static int u32(byte[] bytes, int offset) {
        return (bytes[offset] & 0xFF)
            | ((bytes[offset + 1] & 0xFF) << 8)
            | ((bytes[offset + 2] & 0xFF) << 16)
            | ((bytes[offset + 3] & 0xFF) << 24);
    }

    private static void putU16(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte) value;
        bytes[offset + 1] = (byte) (value >>> 8);
    }

    private static void putU32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte) value;
        bytes[offset + 1] = (byte) (value >>> 8);
        bytes[offset + 2] = (byte) (value >>> 16);
        bytes[offset + 3] = (byte) (value >>> 24);
    }

    private record EmfRecord(int offset, int type, int size) {
    }

    private record PlusRecord(int offset, int type, int flags, int size, int dataSize) {
    }

    private record FloatBounds(float minX, float minY, float maxX, float maxY) {
        float width() {
            return maxX - minX;
        }

        float height() {
            return maxY - minY;
        }
    }
}
