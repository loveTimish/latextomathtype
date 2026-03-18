package com.lz.paperword.core.render;

import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;

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
}
