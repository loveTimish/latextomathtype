package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LaTeXImageRendererTest {

    @Test
    void shouldParseSvgPtDimensions() {
        byte[] svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="36pt" height="18pt" viewBox="0 0 36 18">
              <path d="M0 0h36v18H0z"/>
            </svg>
            """.getBytes(StandardCharsets.UTF_8);

        LaTeXImageRenderer.SvgDimensions dimensions = LaTeXImageRenderer.extractSvgDisplayDimensions(svg);
        assertEquals(36.0f, dimensions.widthPt(), 0.01f);
        assertEquals(18.0f, dimensions.heightPt(), 0.01f);
    }

    @Test
    void shouldParseSvgPxDimensions() {
        byte[] svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="40px" height="20px" viewBox="0 0 40 20">
              <path d="M0 0h40v20H0z"/>
            </svg>
            """.getBytes(StandardCharsets.UTF_8);

        LaTeXImageRenderer.SvgDimensions dimensions = LaTeXImageRenderer.extractSvgDisplayDimensions(svg);
        assertEquals(30.0f, dimensions.widthPt(), 0.05f);
        assertEquals(15.0f, dimensions.heightPt(), 0.05f);
    }

    @Test
    void shouldKeepLongDivisionPreviewHeaderOnly() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(new LaTeXImageRenderer(), "\\longdiv[570]{6}{3420}");
        assertTrue(normalized.contains("\\overset{570}{\\overline{\\left)3420\\right.}}"));
        assertFalse(normalized.contains("\\underline{30}"), "预览图不应再自动补第一步乘积");
        assertFalse(normalized.contains("\\underline{42}"), "预览图不应再自动补后续步骤");
    }

    @Test
    void shouldReplaceCompositeLongDivisionHeaderInsideSingleBlock() throws Exception {
        Method method = LaTeXImageRenderer.class.getDeclaredMethod("normalizeLatexForLocalRender", String.class);
        method.setAccessible(true);

        String normalized = (String) method.invoke(
            new LaTeXImageRenderer(),
            "\\longdiv[570]{6}{3420}\\begin{array}{l}\\text{   }\\underline{30}\\\\\\text{    }42\\end{array}"
        );
        assertFalse(normalized.contains("\\longdiv[570]{6}{3420}"), "单块复合长除法预览时应替换掉原始 longdiv 命令");
        assertTrue(normalized.contains("\\overset{570}{\\overline{\\left)3420\\right.}}"), "预览图应保留长除法头部");
        assertTrue(normalized.contains("\\begin{array}{l}"), "单块复合长除法的步骤区应继续保留");
        assertTrue(normalized.contains("\\underline{30}"), "显式步骤区应继续进入预览渲染");
    }
}
