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
        assertFalse(VectorWmfFormulaRenderer.canRender("x^{\\frac{1}{2}}"));
        assertFalse(VectorWmfFormulaRenderer.canRender("\\sqrt{\\frac{1}{2}}"));
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

    @Test
    void textCommandsCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\vartriangle =\\mathrm{9}+\\cdots", 70.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\mathrm{9}"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\vartriangle"));
        assertTrue(VectorWmfFormulaRenderer.canRender("1\\sim 9"));
        assertTrue(VectorWmfFormulaRenderer.canRender("\\bigcirc"));
        assertTrue(VectorWmfFormulaRenderer.canRender("(1+2+3+\\cdots +9)\\div 3=15"));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleScriptsCanRenderAsVectorText() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("C\\times D=kD ^ { 2 }", 92.0d, 13.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("a_{ 1 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("C\\times D=kD ^ { 2 }"));
        assertTrue(records.contains(0x02FB));
        assertTrue(records.contains(0x012D));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void simpleFractionsCanRenderAsVectorTextAndLines() throws IOException {
        byte[] wmf = VectorWmfFormulaRenderer.render("\\frac { 1 } { 15 }", 24.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender("\\frac { 1 } { 15 }"));
        assertTrue(VectorWmfFormulaRenderer.canRender(
            "\\left ( { 1+2+\\cdots +9+a+b+c } \\right )\\div 3=15+\\frac { a+b+c } { 3 }"
        ));
        assertTrue(records.contains(0x02FA), "fraction vector WMF should create a pen");
        assertTrue(records.contains(0x0325), "fraction vector WMF should draw a fraction bar");
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void fractionsWithSimpleScriptsCanRenderAsVectorTextAndLines() throws IOException {
        String latex = "\\frac { a_{ 1 } +a_{ 2 } +a_{ 3 } } { 3 }";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 54.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(VectorWmfFormulaRenderer.canRender("a=\\frac { k ^ { 2 } } { b-k }+k"));
        assertTrue(records.contains(0x02FA));
        assertTrue(records.contains(0x0325));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void leftBraceArraysCanRenderAsVectorText() throws IOException {
        String latex = "\\left \\{ \\begin{array}{l}B=2 \\\\,s=14\\end{array} \\right.";
        byte[] wmf = VectorWmfFormulaRenderer.render(latex, 42.0d, 28.0d);
        List<Integer> records = records(wmf);

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertTrue(records.contains(0x02FB));
        assertTrue(records.contains(0x0A32));
        assertFalse(records.contains(0x0F43));
    }

    @Test
    void emptyAndParenArraysCanRenderAsVectorText() throws IOException {
        String empty = "\\begin{array}{cccc} {} & \\end{array}";
        String paren = "\\left ( { \\begin{array}{cc} {} & 4, \\\\,8 & \\end{array} } \\right )";

        assertTrue(VectorWmfFormulaRenderer.canRender(empty));
        assertTrue(VectorWmfFormulaRenderer.canRender(paren));
        assertFalse(records(VectorWmfFormulaRenderer.render(empty, 18.0d, 13.0d)).contains(0x0F43));
        assertFalse(records(VectorWmfFormulaRenderer.render(paren, 34.0d, 33.0d)).contains(0x0F43));
    }

    @Test
    void adjacentArraysCanRenderAsVectorText() throws IOException {
        String adjacent = "\\begin{array}{l}cba \\\\,\\times abc\\end{array}\\begin{array}{l}c \\\\,b \\\\,b \\\\,a\\end{array}";
        String withText = "\\begin{array}{l}1b5 \\\\,\\times 5b1\\end{array}1b505\\begin{array}{l}1 \\\\,b \\\\,b \\\\,5\\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(adjacent));
        assertTrue(VectorWmfFormulaRenderer.canRender(withText));
        assertFalse(records(VectorWmfFormulaRenderer.render(adjacent, 60.0d, 50.0d)).contains(0x0F43));
        assertFalse(records(VectorWmfFormulaRenderer.render(withText, 92.0d, 50.0d)).contains(0x0F43));
    }

    @Test
    void widePuzzleArraysCanRenderAsVectorText() throws IOException {
        String latex = "\\begin{array}{cccccccc} ( & \\mathrm{7} & + & \\mathrm{9} & )\\div & \\mathrm{8} & = & \\mathrm{2} \\\\,{} & + & {} & - & {} & \\div & {} & {} \\\\,{} & \\mathrm{1}\\mathrm{1} & - & \\mathrm{1}\\mathrm{0} & - & \\mathrm{1} & = & \\mathrm{0} \\\\,{} & - & {} & - & {} & - & {} & {} \\\\,{} & \\mathrm{1}\\mathrm{2} & - & \\mathrm{3} & \\times & \\mathrm{4} & = & \\mathrm{0} \\\\,{} & - & {} & + & {} & \\div & {} & {} \\\\,{} & \\mathrm{5} & + & \\mathrm{6} & \\div & \\mathrm{2} & = & \\mathrm{8} \\\\,{} & \\|\\| & {} & \\|\\| & {} & \\|\\| & {} & {} \\\\,{} & \\mathrm{1} & {} & \\mathrm{2} & {} & \\mathrm{6} & {} & {} \\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertFalse(records(VectorWmfFormulaRenderer.render(latex, 150.0d, 110.0d)).contains(0x0F43));
    }

    @Test
    void nestedArrayCellsCanRenderAsVectorText() throws IOException {
        String latex = "\\begin{array}{ccccc} {} & {} & 11 & {} & {} \\\\,{} & 17 & {} & 13 & {} \\\\,23 & {} & 19 & {} & 15 \\\\,{} & 25 & {} & 21 & {} \\\\,{} & {} & 27 & {} & {} \\\\,\\to & \\begin{array}{ccccc} {} & \\end{array} & {} & {} & 27 \\\\,{} & {} & {} & 17 & {} \\\\,13 & {} & 23 & {} & 19 \\\\,{} & 15 & {} & 25 & {} \\\\,21 & {} & {} & {} & 11 \\\\,{} & {} & \\to & \\begin{array}{ccccc} {} & \\end{array} & {} \\\\,{} & 27 & {} & {} & {} \\\\,17 & {} & 13 & {} & 15 \\\\,{} & 19 & {} & 23 & {} \\\\,25 & {} & 21 & {} & {} \\\\,{} & 11 & {} & {} & \\to \\\\,\\begin{array}{ccc} 172713151923251121 & \\end{array} & \\end{array}";

        assertTrue(VectorWmfFormulaRenderer.canRender(latex));
        assertFalse(records(VectorWmfFormulaRenderer.render(latex, 168.0d, 205.0d)).contains(0x0F43));
    }

    @Test
    void standaloneScriptsAndEscapedUnderscoresCanRenderAsVectorText() throws IOException {
        byte[] script = VectorWmfFormulaRenderer.render("^ { \\mathrm{3} }", 12.0d, 13.0d);
        byte[] underline = VectorWmfFormulaRenderer.render("EF=\\_\\_\\_\\_\\_", 56.0d, 13.0d);

        assertTrue(VectorWmfFormulaRenderer.canRender("^ { \\mathrm{3} }"));
        assertTrue(VectorWmfFormulaRenderer.canRender("EF=\\_\\_\\_\\_\\_"));
        assertTrue(records(script).contains(0x0A32));
        assertTrue(records(underline).contains(0x0A32));
        assertFalse(records(script).contains(0x0F43));
        assertFalse(records(underline).contains(0x0F43));
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
