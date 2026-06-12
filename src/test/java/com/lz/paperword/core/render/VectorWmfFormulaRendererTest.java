package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class VectorWmfFormulaRendererTest {

    @Test
    void linearFormulaUsesVectorTextRecordsInsteadOfStretchDib() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("2.25\\div 0.9=2.5", 68.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x02FB), "vector WMF should create a font");
        assertTrue(records.contains(0x0A32), "vector WMF should draw formula text with ExtTextOut");
        assertFalse(records.contains(0x0F43), "linear vector WMF must not embed a DIB bitmap");
    }

    @Test
    void structuredFormulaIsNotClaimedAsVectorRenderableYet() {
        assertFalse(VectorWmfFormulaRenderer.canRender("\\frac{1}{2}"));
        assertFalse(VectorWmfFormulaRenderer.canRender("x^{2}"));
    }

    @Test
    void explicitFlatParenFenceCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\left ( { 第+十+一+届+华+杯+赛 } \\right )", 141.0d, 19.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void squarePlaceholderCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("+2=\\square", 40.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleArrayCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\begin{array}{ccccc} ABCD-EFGH2008 & \\end{array}", 89.0d, 49.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\begin{array}{c}1\\\\2\\end{array}"));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    private static List<Integer> records(byte[] data) {
        int offset = hasPlaceableHeader(data) ? 22 : 0;
        offset += 18;
        List<Integer> out = new ArrayList<>();
        while (offset + 6 <= data.length) {
            int sizeWords = dword(data, offset);
            int function = word(data, offset + 4);
            out.add(function);
            if (function == 0x0000 || sizeWords <= 0) {
                break;
            }
            offset += sizeWords * 2;
        }
        return out;
    }

    private static boolean hasPlaceableHeader(byte[] data) {
        return data.length >= 22 && dword(data, 0) == 0x9AC6CDD7;
    }

    private static int word(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 2).order(ByteOrder.LITTLE_ENDIAN).getShort() & 0xffff;
    }

    private static int dword(byte[] data, int offset) {
        return ByteBuffer.wrap(data, offset, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
    }
}
