package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.layout.VerticalLayoutSpec.Kind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionHeader;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionPlacedLine;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class VerticalLayoutNodeFactoryTest {

    private final VerticalLayoutNodeFactory factory = new VerticalLayoutNodeFactory();
    private final LaTeXParser parser = new LaTeXParser();

    @Test
    void shouldKeepCompositeLongDivisionAsStaircaseTextRows() {
        LaTeXNode longDivision = parser.parseLaTeX("\\longdiv[14300]{7}{100100}");
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            8,
            7,
            -1,
            List.of(),
            List.of(),
            List.of(),
            new LongDivisionHeader("7", "14300", "100100")
        );

        LaTeXNode normalized = factory.buildCompositeLongDivisionArray(
            longDivision,
            spec,
            List.of(
                new VerticalLayoutNodeFactory.RawLongDivisionLine("7", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("30", false),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("28", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("  21", false),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("  21", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("      00", false),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("      00", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("        0", false)
            )
        );

        LaTeXNode firstStepContent = firstCellContent(normalized, 1);
        LaTeXNode secondStepContent = firstCellContent(normalized, 2);
        LaTeXNode lastStepContent = firstCellContent(normalized, 8);

        assertEquals(LaTeXNode.Type.GROUP, firstStepContent.getType(), "下划线步骤行应保留为前缀文本加 underline 组合");
        assertTrue(containsNodeType(firstStepContent, LaTeXNode.Type.TEXT), "单块复合长除法应继续使用文本节点承载渲染结果");
        assertFalse(containsTabChar(firstStepContent), "单块复合长除法不应再写入 tab 锚点");
        assertEquals("      7", flattenChars(firstStepContent), "首步单数字乘积应在当前被除数块内右对齐");

        assertEquals(LaTeXNode.Type.TEXT, secondStepContent.getType(), "普通步骤行应继续使用文本节点承载对齐结果");
        assertFalse(containsTabChar(secondStepContent), "余数行不应再写入 tab 锚点");
        assertTrue(containsNodeType(secondStepContent, LaTeXNode.Type.TEXT), "余数行也应继续走文本渲染");
        assertEquals("      30", flattenChars(secondStepContent), "第二行应根据上一轮余数和下移数字落到正确列位");
        assertEquals("              0", flattenChars(lastStepContent), "最终余数 0 应落在最后一位被下移数字所在列");
    }

    @Test
    void shouldNormalizeRawLongDivisionLinesIntoPlacedColumns() {
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            8,
            7,
            -1,
            List.of(),
            List.of(),
            List.of(),
            new LongDivisionHeader("7", "14300", "100100")
        );

        List<LongDivisionPlacedLine> placedLines = factory.normalizeRawLongDivisionLines(
            spec,
            List.of(
                new VerticalLayoutNodeFactory.RawLongDivisionLine("7", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("30", false),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("  21", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("        0", false)
            )
        );

        assertEquals(4, placedLines.size(), "应把每条 raw line 转成结构化列坐标行");
        assertEquals(1, placedLines.get(0).startColumn(), "首步单数字乘积应在当前被除数块内右对齐");
        assertEquals(1, placedLines.get(0).underlineStartColumn(), "首步下划线应与单数字乘积的右对齐起点一致");
        assertEquals(1, placedLines.get(0).underlineEndColumn(), "单数字乘积的下划线应扩到当前被除数块宽度");
        assertEquals(2, placedLines.get(2).startColumn(), "后续步骤应按上一轮结果推进到新的被除数块");
        assertEquals(5, placedLines.get(3).startColumn(), "最终余数 0 应落在最后一位被下移数字所在列");
    }

    void shouldKeepTrailingZeroRowFromRawLongDivisionLines() {
        LaTeXNode longDivision = parser.parseLaTeX("\\longdiv[5]{2.5}{12.5}");
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            7,
            6,
            -1,
            List.of(),
            List.of(),
            List.of(),
            new LongDivisionHeader("2.5", "5", "12.5")
        );

        LaTeXNode normalized = factory.buildCompositeLongDivisionArray(
            longDivision,
            spec,
            List.of(
                new VerticalLayoutNodeFactory.RawLongDivisionLine("125", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("  0", false)
            )
        );
        assertEquals(3, normalized.getChildren().size(), "原始步骤区中的最终 0 应被完整保留");
        assertEquals("          0", flattenChars(firstCellContent(normalized, 2)), "小数整除收尾的 0 应按头部宽度和相对列位共同定位");
    }

    @Test
    void shouldRenderRawDecimalLongDivisionHeader() {
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            7,
            6,
            -1,
            List.of(),
            List.of(),
            List.of(),
            new LongDivisionHeader("2.5", "5", "12.5")
        );
        LaTeXNode longDivision = parser.parseLaTeX("\\longdiv[5]{2.5}{12.5}");

        LaTeXNode normalized = factory.buildCompositeLongDivisionArray(longDivision, spec, List.of());
        assertTrue(flattenChars(firstCellContent(normalized, 0)).contains("2.5"), "长除法头部应展示 LLM 返回的原始除数");
        assertTrue(flattenChars(firstCellContent(normalized, 0)).contains("12.5"), "长除法头部应展示原始被除数");
        assertTrue(flattenChars(firstCellContent(normalized, 0)).contains("5"), "长除法头部应继续保留显式商值");
    }

    @Test
    void shouldKeepDecimalQuotientRowsAsRightShiftedStaircase() {
        LaTeXNode longDivision = parser.parseLaTeX("\\longdiv[0.44]{5}{2.2}");
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            5,
            4,
            -1,
            List.of(),
            List.of(),
            List.of(),
            new LongDivisionHeader("5", "0.44", "2.2")
        );

        LaTeXNode normalized = factory.buildCompositeLongDivisionArray(
            longDivision,
            spec,
            List.of(
                new VerticalLayoutNodeFactory.RawLongDivisionLine("20", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("20", false),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("20", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("  0", false)
            )
        );

        assertEquals("    20", flattenChars(firstCellContent(normalized, 1)), "首个乘积应先补除数和除号宽度，再落到相对首列");
        assertEquals("    20", flattenChars(firstCellContent(normalized, 2)), "后续行应与首列保持统一原点");
        assertEquals("      0", flattenChars(firstCellContent(normalized, 4)), "最终余数 0 应继续保留相对右移后的列位");
    }

    @Test
    void shouldBuildCompositeLongDivisionInsideDividendPile() {
        LaTeXNode longDivision = parser.parseLaTeX("\\longdiv[570]{6}{3420}");
        VerticalLayoutSpec spec = new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            6,
            5,
            -1,
            List.of(),
            List.of(),
            List.of(),
            new LongDivisionHeader("6", "570", "3420")
        );

        LaTeXNode structured = factory.buildCompositeLongDivisionTemplateNode(
            longDivision,
            spec,
            List.of(
                new VerticalLayoutNodeFactory.RawLongDivisionLine("30", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("42", false),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("42", true),
                new VerticalLayoutNodeFactory.RawLongDivisionLine("0", false)
            )
        );

        assertEquals(LaTeXNode.Type.LONG_DIVISION, structured.getType(), "复合长除法应继续输出 tmLDIV 头部");
        LaTeXNode dividendPile = structured.getChildren().get(2);
        assertEquals("true", dividendPile.getMetadata(VerticalLayoutNodeFactory.RAW_PILE_CONTAINER),
            "步骤区应进入 dividend 槽位内的 raw pile，而不是挂在外层矩阵行上");
        assertEquals(5, dividendPile.getChildren().size(), "dividend 槽位内应包含被除数行和全部步骤行");
        assertTrue(containsTabChar(dividendPile), "dividend 槽位内的步骤应改由制表位驱动，而不是靠普通空格近似");
    }

    private LaTeXNode firstCellContent(LaTeXNode arrayNode, int rowIndex) {
        return arrayNode.getChildren().get(rowIndex).getChildren().get(0).getChildren().get(0);
    }

    private boolean containsTabChar(LaTeXNode node) {
        if (node == null) {
            return false;
        }
        if (node.getType() == LaTeXNode.Type.CHAR && "\t".equals(node.getValue())) {
            return true;
        }
        for (LaTeXNode child : node.getChildren()) {
            if (containsTabChar(child)) {
                return true;
            }
        }
        return false;
    }

    private boolean containsNodeType(LaTeXNode node, LaTeXNode.Type targetType) {
        if (node == null) {
            return false;
        }
        if (node.getType() == targetType) {
            return true;
        }
        for (LaTeXNode child : node.getChildren()) {
            if (containsNodeType(child, targetType)) {
                return true;
            }
        }
        return false;
    }

    private String flattenChars(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        appendChars(node, builder);
        return builder.toString();
    }

    private void appendChars(LaTeXNode node, StringBuilder builder) {
        if (node == null) {
            return;
        }
        if (node.getType() == LaTeXNode.Type.CHAR && node.getValue() != null) {
            builder.append(node.getValue());
        }
        for (LaTeXNode child : node.getChildren()) {
            appendChars(child, builder);
        }
    }
}
