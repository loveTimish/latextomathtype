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
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24pt\" height=\"12pt\" viewBox=\"0 -1000 2000 1000\">"
            + "<g stroke=\"#000000\" fill=\"#000000\" stroke-width=\"0\" transform=\"scale(1,-1)\">"
            + "<path d=\"M10 -10L100 -10L100 -100L10 -100Z\"/>"
            + "<rect x=\"200\" y=\"-50\" width=\"300\" height=\"20\"/>"
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
        assertTrue(countRecords(wmf, 0x0538) >= 2, "path + rect 应各产生一条 POLYPOLYGON");
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
    void shouldRejectUnknownElement() {
        String svg = SIMPLE_SVG.replace("</g>", "<ellipse cx=\"1\" cy=\"2\" rx=\"3\" ry=\"4\"/></g>");
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d));
    }

    @Test
    void shouldRejectCubicPathCommand() {
        String svg = SIMPLE_SVG.replace("M10 -10L100 -10L100 -100L10 -100Z",
            "M10 -10C100 -10 100 -100 10 -100Z");
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d));
    }

    @Test
    void shouldRejectRotateTransform() {
        String svg = SIMPLE_SVG.replace("scale(1,-1)", "rotate(45)");
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d));
    }

    @Test
    void shouldRejectStrokeOnPath() {
        String svg = SIMPLE_SVG.replace("<path d=", "<path stroke=\"#ff0000\" d=");
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d));
    }

    @Test
    void shouldRejectMissingViewBox() {
        String svg = SIMPLE_SVG.replace(" viewBox=\"0 -1000 2000 1000\"", "");
        assertThrows(SvgVectorWmfRenderer.SvgVectorWmfException.class,
            () -> SvgVectorWmfRenderer.render(svg.getBytes(StandardCharsets.UTF_8), 24d, 12d));
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
}
