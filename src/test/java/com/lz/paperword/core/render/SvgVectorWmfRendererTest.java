package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** 真矢量 WMF 渲染器（SVG → POLYPOLYGON）的结构与严格模式测试。 */
class SvgVectorWmfRendererTest {

    private static final String SIMPLE_SVG =
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24pt\" height=\"12pt\" viewBox=\"0 0 2000 1000\">"
            + "<g stroke=\"#000000\" fill=\"#000000\" stroke-width=\"0\" transform=\"scale(1,1)\">"
            + "<path d=\"M10 10L100 10L100 100L10 100Z\"/>"
            + "<rect x=\"200\" y=\"50\" width=\"300\" height=\"20\"/>"
            + "</g></svg>";

    /** 带孔洞的字形（外轮廓 + 内轮廓）与二次曲线、T 反射。 */
    private static final String GLYPH_SVG =
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10pt\" height=\"10pt\" viewBox=\"0 0 500 500\">"
            + "<g transform=\"translate(0,0)\">"
            + "<path d=\"M50 50Q100 0 150 50T250 50L250 450L50 450ZM100 100L200 100L200 400L100 400Z\"/>"
            + "</g></svg>";

    @Test
    void shouldEmitPlaceableVectorWmfWithoutBitmapRecords() throws Exception {
        byte[] wmf = SvgVectorWmfRenderer.render(
            SIMPLE_SVG.getBytes(StandardCharsets.UTF_8), 24d, 12d);

        // placeable key
        assertEquals((byte) 0xD7, wmf[0]);
        assertEquals((byte) 0xCD, wmf[1]);
        assertEquals((byte) 0xC6, wmf[2]);
        assertEquals((byte) 0x9A, wmf[3]);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(wmf), "矢量 WMF 不应含位图记录");
        assertTrue(countRecords(wmf, 0x0538) >= 1, "Batik 应把 path + rect 输出为矢量轮廓");
        assertEquals(1, countRecords(wmf, 0x0000), "应恰好一条 EOF");
        assertEquals(1, countRecords(wmf, 0x0106), "应设置 WINDING 填充模式");
        assertEquals(1, countRecords(wmf, 0x02FC), "应创建画刷（stock 画刷不可靠）");
        // size 字段 = 记录区 WORD 数 + 9
        long sizeWords = u32(wmf, 22 + 6);
        assertEquals((wmf.length - 22) / 2, sizeWords, "标准头 size 字段应等于 (记录区+标准头)/2");
    }

    @Test
    void shouldKeepMultiContourGlyphInSinglePolyPolygon() throws Exception {
        byte[] wmf = SvgVectorWmfRenderer.render(
            GLYPH_SVG.getBytes(StandardCharsets.UTF_8), 10d, 10d);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(wmf));
        assertEquals(1, countRecords(wmf, 0x0538),
            "同一 path 的外轮廓+孔洞必须在同一条 POLYPOLYGON 中，否则孔洞被填死");
    }

    @Test
    void shouldRenderEllipseThroughBatik() throws Exception {
        String svg = SIMPLE_SVG.replace("</g>", "<ellipse cx=\"1\" cy=\"2\" rx=\"3\" ry=\"4\"/></g>");
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(wmf));
        assertTrue(countRecords(wmf, 0x0538) >= 1);
    }

    @Test
    void shouldRenderCubicPathCommand() throws Exception {
        String svg = SIMPLE_SVG.replace("M10 10L100 10L100 100L10 100Z",
            "M10 10C100 10 100 100 10 100Z");
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(wmf));
    }

    @Test
    void shouldRenderRotateTransform() throws Exception {
        String svg = SIMPLE_SVG.replace("scale(1,1)", "rotate(45)");
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(wmf));
    }

    @Test
    void shouldPreserveAspectRatioWhenTargetBoxDiffersFromSvgViewport() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"20pt\" height=\"10pt\" "
            + "viewBox=\"0 0 200 100\"><rect x=\"20\" y=\"20\" width=\"160\" height=\"60\"/></svg>";
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 20d, 20d);
        int[] bounds = drawingBounds(wmf);
        double ratio = (bounds[2] - bounds[0]) / (double) (bounds[3] - bounds[1]);

        assertTrue(ratio > 2.4d && ratio < 2.9d,
            "目标框比例不同也必须等比缩放，实际墨迹宽高比=" + ratio);
    }

    @Test
    void shouldExpandStrokeAndKeepColor() throws Exception {
        String svg = SIMPLE_SVG.replace("<path d=", "<path stroke=\"#ff0000\" stroke-width=\"10\" d=");
        SvgVectorWmfRenderer.VectorWmfResult result = SvgVectorWmfRenderer.renderDetailed(
            svg.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        assertTrue(result.colors().contains(0xFF0000));
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(result.bytes()));
    }

    @Test
    void shouldRenderSvgWithPhysicalViewportAndNoViewBox() throws Exception {
        String svg = SIMPLE_SVG.replace(" viewBox=\"0 0 2000 1000\"", "");
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(wmf));
    }

    @Test
    void shouldRenderNestedSvgLinesPolygonsAndUnicodeTextAsOutlines() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40pt\" height=\"20pt\" "
            + "viewBox=\"0 0 400 200\"><g fill=\"#007f00\" stroke=\"#0000ff\">"
            + "<svg x=\"10\" y=\"10\" width=\"180\" height=\"80\" viewBox=\"0 0 180 80\">"
            + "<polygon points=\"0,80 90,0 180,80\"/></svg>"
            + "<line x1=\"10\" y1=\"120\" x2=\"390\" y2=\"120\" stroke-width=\"6\" "
            + "stroke-dasharray=\"20 10\"/><text x=\"20\" y=\"180\" font-size=\"40\">§€中文К𝛼</text>"
            + "</g></svg>";
        SvgVectorWmfRenderer.VectorWmfResult result = SvgVectorWmfRenderer.renderDetailed(
            svg.getBytes(StandardCharsets.UTF_8), 40d, 20d);
        assertFalse(SvgVectorWmfRenderer.containsBitmapRecord(result.bytes()));
        assertTrue(result.colors().contains(0x007F00));
        assertTrue(result.colors().contains(0x0000FF));
        assertTrue(result.shapeCount() >= 3);
        assertTrue(result.outlinedCodePoints().contains((int) '§'));
        assertTrue(result.outlinedCodePoints().contains((int) '€'));
        assertTrue(result.outlinedCodePoints().contains((int) '中'));
        assertTrue(result.outlinedCodePoints().contains((int) 'К'));
        assertTrue(result.outlinedCodePoints().contains(0x1D6FC));
    }

    @Test
    void shouldRejectCodePointMissingFromFixedFontSet() {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10pt\" height=\"10pt\" "
            + "viewBox=\"0 0 10 10\"><text x=\"0\" y=\"8\">&#x10FFFF;</text></svg>";
        SvgVectorWmfRenderer.SvgVectorWmfException error = assertThrows(
            SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 10d, 10d));
        assertTrue(error.getMessage().contains("MISSING_GLYPH"));
    }

    @Test
    void shouldRejectEmbeddedRasterImage() {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" "
            + "width=\"10pt\" height=\"10pt\" viewBox=\"0 0 10 10\">"
            + "<image width=\"10\" height=\"10\" xlink:href=\"data:image/png;base64,"
            + "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZpSIAAAAASUVORK5CYII=\"/>"
            + "</svg>";
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 10d, 10d));
    }

    @Test
    void shouldEmitValidEmptyVectorWmfForLegalEmptyScene() throws Exception {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10pt\" height=\"5pt\" "
            + "viewBox=\"0 0 10 5\"><g fill=\"none\"/></svg>";
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 10d, 5d);
        WmfPreviewInspector.Inspection inspection = WmfPreviewInspector.inspect(wmf);

        assertTrue(inspection.valid(), inspection.error());
        assertTrue(inspection.pureVector());
        assertEquals(0, inspection.drawingRecordCount());
        assertEquals(0, inspection.foregroundPixels());
    }

    @Test
    void shouldApplyTransformsOnPathElementsThemselves() throws Exception {
        // MathJax 把字间距烘焙在 path 自身的 transform 上（translate(278,0) 等），
        // 漏掉会让同组兄弟字形全部叠在原点（lim 变 "lm" 块）。
        String svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"20pt\" height=\"10pt\" viewBox=\"0 0 1000 500\">"
                + "<g><path d=\"M0 0L100 0L100 100L0 100Z\"/>"
                + "<path d=\"M0 0L100 0L100 100L0 100Z\" transform=\"translate(500,0)\"/></g></svg>";
        byte[] wmf = SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 20d, 10d);
        int[] minX = {Integer.MAX_VALUE, Integer.MAX_VALUE};
        int[] maxX = {Integer.MIN_VALUE, Integer.MIN_VALUE};
        int poly = 0;
        int offset = 22 + 18;
        while (offset + 6 <= wmf.length) {
            long sizeWords = u32(wmf, offset);
            int func = (wmf[offset + 4] & 0xFF) | ((wmf[offset + 5] & 0xFF) << 8);
            if (func == 0x0538 && poly < 2) {
                int nPts = (wmf[offset + 8] & 0xFF) | ((wmf[offset + 9] & 0xFF) << 8);
                int p = offset + 10;
                for (int i = 0; i < nPts; i++) {
                    int x = (short) ((wmf[p] & 0xFF) | ((wmf[p + 1] & 0xFF) << 8));
                    minX[poly] = Math.min(minX[poly], x);
                    maxX[poly] = Math.max(maxX[poly], x);
                    p += 4;
                }
                poly++;
            }
            if (sizeWords < 3 || offset + sizeWords * 2 > wmf.length || func == 0) {
                break;
            }
            offset += (int) sizeWords * 2;
        }
        assertEquals(2, poly, "应有两条 POLYPOLYGON");
        assertTrue(maxX[0] < minX[1],
            "第二个 path 的 translate(500,0) 未生效：两条多边形 x 区间重叠 ["
                + minX[0] + "," + maxX[0] + "] vs [" + minX[1] + "," + maxX[1] + "]");
    }

    private static long u32(byte[] b, int off) {
        return (b[off] & 0xFFL)
            | ((b[off + 1] & 0xFFL) << 8)
            | ((b[off + 2] & 0xFFL) << 16)
            | ((b[off + 3] & 0xFFL) << 24);
    }

    private static int countRecords(byte[] wmf, int func) {
        int count = 0;
        int offset = 22 + 18;
        while (offset + 6 <= wmf.length) {
            long sizeWords = u32(wmf, offset);
            int f = (wmf[offset + 4] & 0xFF) | ((wmf[offset + 5] & 0xFF) << 8);
            if (f == func) {
                count++;
            }
            if (sizeWords < 3 || offset + sizeWords * 2 > wmf.length) {
                break;
            }
            if (f == 0) {
                break;
            }
            offset += (int) sizeWords * 2;
        }
        return count;
    }

    private static int[] drawingBounds(byte[] wmf) {
        int[] bounds = {Integer.MAX_VALUE, Integer.MAX_VALUE, Integer.MIN_VALUE, Integer.MIN_VALUE};
        int offset = 22 + 18;
        while (offset + 10 <= wmf.length) {
            long sizeWords = u32(wmf, offset);
            int func = (wmf[offset + 4] & 0xFF) | ((wmf[offset + 5] & 0xFF) << 8);
            if (func == 0x0538) {
                int polygonCount = (wmf[offset + 6] & 0xFF) | ((wmf[offset + 7] & 0xFF) << 8);
                int totalPoints = 0;
                int countsOffset = offset + 8;
                for (int i = 0; i < polygonCount; i++) {
                    totalPoints += (wmf[countsOffset + i * 2] & 0xFF)
                        | ((wmf[countsOffset + i * 2 + 1] & 0xFF) << 8);
                }
                int pointOffset = countsOffset + polygonCount * 2;
                for (int i = 0; i < totalPoints; i++) {
                    int x = (short) ((wmf[pointOffset] & 0xFF) | ((wmf[pointOffset + 1] & 0xFF) << 8));
                    int y = (short) ((wmf[pointOffset + 2] & 0xFF) | ((wmf[pointOffset + 3] & 0xFF) << 8));
                    bounds[0] = Math.min(bounds[0], x);
                    bounds[1] = Math.min(bounds[1], y);
                    bounds[2] = Math.max(bounds[2], x);
                    bounds[3] = Math.max(bounds[3], y);
                    pointOffset += 4;
                }
            }
            if (sizeWords < 3 || offset + sizeWords * 2 > wmf.length || func == 0) {
                break;
            }
            offset += (int) sizeWords * 2;
        }
        return bounds;
    }
}
