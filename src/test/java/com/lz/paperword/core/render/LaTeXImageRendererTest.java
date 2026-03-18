package com.lz.paperword.core.render;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
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
    void shouldBuildOfficialStyleXlopLongDivisionLatex() {
        LaTeXParser parser = new LaTeXParser();
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        String latex = renderer.buildLongDivisionPictureLatex(ast);

        assertTrue(!latex.contains("\\opset{columnwidth"),
            "官方样式不应继续手写 columnwidth");
        assertTrue(!latex.contains("voperator="),
            "官方样式不应继续手写 voperator");
        assertTrue(!latex.contains("voperation="),
            "官方样式不应继续手写 voperation");
        assertTrue(latex.contains("\\opidiv{1234}{5}"),
            "整数长除法应使用 xlop 的整数除法命令，避免生成小数商");
    }

    @Test
    void shouldExtractOnlyLongDivisionNodeFromCompositeTemplate() {
        LaTeXParser parser = new LaTeXParser();
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{l}{\\longdiv[246]{5}{1234}} \\\\ " +
                "{\\begin{array}{rrrr}{1} & {2} & {3} & {4} \\\\ \\hline {} & {1} & {0} & {}\\end{array}}\\end{array}"
        );
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        String latex = renderer.buildLongDivisionPictureLatex(ast);

        assertTrue(latex.contains("\\opidiv{1234}{5}"),
            "复合模板长除法应只抽取 longdiv 本体，并使用整数除法命令");
        assertTrue(!latex.contains("1} & {2}"),
            "复合模板中的手写参考步骤不应进入最终 xlop 公式");
    }

    @Test
    void shouldNormalizeLegacyDivisionArrayToOfficialXlopCommand() {
        LaTeXParser parser = new LaTeXParser();
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}"
        );
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        String latex = renderer.buildLongDivisionPictureLatex(ast);

        assertTrue(latex.contains("\\opidiv{845}{13}"),
            "旧除法模板应先正规化成官方 opidiv 命令");
    }

    @Test
    void shouldNormalizeLongdivisionEnvironmentToOfficialXlopCommand() {
        LaTeXParser parser = new LaTeXParser();
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{longdivision}{r|l}13 & 845 \\\\ \\hline & 65\\end{longdivision}"
        );
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        String latex = renderer.buildLongDivisionPictureLatex(ast);

        assertTrue(latex.contains("\\opidiv{845}{13}"),
            "longdivision 环境应先正规化成官方 opidiv 命令");
    }

    @Test
    void shouldFallbackToStructuredPreviewWhenXlopToolchainIsUnavailable() {
        String previousLatexCommand = System.getProperty("paperword.latex.command");
        try {
            System.setProperty("paperword.latex.command", "__missing_latex_for_longdivision__");
            LaTeXParser parser = new LaTeXParser();
            LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
            LaTeXImageRenderer renderer = new LaTeXImageRenderer();

            LaTeXImageRenderer.PreviewImage preview = renderer.renderLongDivisionPicture(ast, "\\longdiv[246]{5}{1234}");

            assertNotNull(preview, "当 xlop 工具链不可用时，应回退到本地图片兜底");
            assertNotNull(preview.data(), "工具链不可用时仍应产出可插图的图片数据");
            assertTrue(preview.data().length > 0, "回退图片数据不应为空");
        } finally {
            restoreProperty("paperword.latex.command", previousLatexCommand);
        }
    }

    @Test
    void shouldRenderOfficialXlopPreviewWhenToolchainAvailable() throws Exception {
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();
        Method renderSvgMethod = LaTeXImageRenderer.class.getDeclaredMethod(
            "renderSvgViaDvisvgm", String.class, float.class, List.class, boolean.class, java.nio.file.Path.class
        );
        renderSvgMethod.setAccessible(true);
        Method method = LaTeXImageRenderer.class.getDeclaredMethod(
            "renderViaDvisvgm", String.class, float.class, List.class, boolean.class
        );
        method.setAccessible(true);
        Method svgToPngMethod = LaTeXImageRenderer.class.getDeclaredMethod("svgToPng", byte[].class);
        svgToPngMethod.setAccessible(true);

        byte[] svg = (byte[]) renderSvgMethod.invoke(renderer, "\\opdiv{845}{13}", 14f, List.of("\\usepackage{xlop}"), false, null);
        byte[] transcodedPng = svg == null ? null : (byte[]) svgToPngMethod.invoke(renderer, svg);

        byte[] png = (byte[]) method.invoke(renderer, "\\opdiv{845}{13}", 14f, List.of("\\usepackage{xlop}"), false);

        assertNotNull(svg, "当前环境可用时，官方 xlop 渲染链应先返回 SVG");
        assertTrue(svg.length > 0, "官方 xlop 渲染链返回的 SVG 不应为空");
        assertNotNull(transcodedPng, "dvisvgm 产出的 SVG 应可继续转成 PNG");
        assertTrue(transcodedPng.length > 0, "SVG 转出来的 PNG 不应为空");
        assertNotNull(png, "当前环境可用时，官方 xlop 渲染链应返回 PNG");
        assertTrue(png.length > 0, "官方 xlop 渲染链返回的 PNG 不应为空");
    }

    @Test
    void shouldReturnPreviewForExplicitLongDivisionInCurrentEnvironment() {
        LaTeXParser parser = new LaTeXParser();
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        LaTeXImageRenderer renderer = new LaTeXImageRenderer();

        LaTeXImageRenderer.PreviewImage preview = renderer.renderLongDivisionPicture(ast, "\\longdiv[65]{13}{845}");

        assertNotNull(preview, "无论走 xlop 还是本地兜底，长除法都应返回可插图预览");
        assertNotNull(preview.data(), "长除法预览数据不应为空");
        assertTrue(preview.data().length > 0, "长除法预览图片字节不应为空");
    }

    private void restoreProperty(String key, String value) {
        if (value == null) {
            System.clearProperty(key);
            return;
        }
        System.setProperty(key, value);
    }
}
