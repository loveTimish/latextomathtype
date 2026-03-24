package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutCompiler.CrossMultiplicationLayout;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RuleSpan;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RowKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionStep;
import com.lz.paperword.core.layout.VerticalLayoutSpec.TabStopKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalSegment;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalTabStop;

import java.util.ArrayList;
import java.util.List;

/**
 * 将统一布局规格转换回适合 MTEF MATRIX 序列化的 AST 结构。
 */
public class VerticalLayoutNodeFactory {

    public static final String LONG_DIVISION_HEADER_ONLY = "longDivisionHeaderOnly";
    public static final String PRESERVE_RAW_ARRAY = "preserveRawArray";
    public static final String RAW_PILE_CONTAINER = "rawPileContainer";
    public static final String RAW_LINE_CONTAINER = "rawLineContainer";
    public static final String RAW_INLINE_CONTAINER = "rawInlineContainer";
    public static final String LONG_DIVISION_STEP_SPLIT = "longDivisionStepSplit";
    public static final String LONG_DIVISION_FIRST_STEP_SPLIT = "longDivisionFirstStepSplit";
    public static final String RAW_PILE_HALIGN = "rawPileHalign";
    public static final String RAW_PILE_VALIGN = "rawPileValign";
    public static final String RAW_PILE_RULER = "rawPileRuler";
    public static final String RAW_MATRIX_HALIGN = "rawMatrixHalign";

    public LaTeXNode buildArrayNode(VerticalLayoutSpec spec) {
        if (spec == null) {
            return null;
        }
        LaTeXNode array = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        array.setMetadata("columnSpec", resolveColumnSpec(spec));
        array.setMetadata("columnCount", String.valueOf(resolveColumnCount(spec)));
        array.setMetadata("columnLines", encodeZeros(resolveColumnCount(spec) + 1));
        array.setMetadata("rowLines", encodeRowLines(spec.rows().size(), spec.ruleSpans()));

        for (VerticalRow rowSpec : spec.rows()) {
            LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
            for (String text : resolveRowCells(spec, rowSpec)) {
                row.addChild(buildCell(text));
            }
            array.addChild(row);
        }
        return array;
    }

    public LaTeXNode buildCompositeLongDivisionArray(LaTeXNode originalOuterArray, VerticalLayoutSpec spec) {
        LaTeXNode normalized = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        normalized.getMetadata().putAll(originalOuterArray.getMetadata());
        normalized.setMetadata("columnCount", "1");
        normalized.setMetadata("columnSpec", "l");
        normalized.setMetadata("columnLines", "0,0");
        normalized.setMetadata("rowLines", "0,0,0");

        LaTeXNode firstRow = new LaTeXNode(LaTeXNode.Type.ROW);
        firstRow.addChild(copyCellContent(originalOuterArray.getChildren().get(0).getChildren().get(0), true));
        normalized.addChild(firstRow);

        LaTeXNode secondRow = new LaTeXNode(LaTeXNode.Type.ROW);
        secondRow.addChild(buildArrayNode(spec));
        normalized.addChild(secondRow);
        return normalized;
    }

    public LaTeXNode buildComputedLongDivisionArray(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        LaTeXNode normalized = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        normalized.setMetadata("columnCount", "1");
        normalized.setMetadata("columnSpec", "l");
        normalized.setMetadata("columnLines", "0,0");
        // 关键：标记为保留原始数组，避免被 writeArrayNode 的 composite long division 处理破坏
        normalized.setMetadata(PRESERVE_RAW_ARRAY, "true");

        // 第一行：简化的 tmLDIV 头部（只有除数、商、被除数，步骤在第二行）
        LaTeXNode firstRow = new LaTeXNode(LaTeXNode.Type.ROW);
        LaTeXNode headerCell = new LaTeXNode(LaTeXNode.Type.CELL);
        headerCell.addChild(buildSimpleLongDivisionHeader(longDivisionNode, spec));
        firstRow.addChild(headerCell);
        normalized.addChild(firstRow);

        // 后续行：使用真实空格字符逐行排版步骤区，避免再依赖复杂嵌套容器。
        if (spec != null && !spec.longDivisionSteps().isEmpty()) {
            for (LaTeXNode stepRow : buildComputedLongDivisionRows(spec)) {
                normalized.addChild(stepRow);
            }
        }
        normalized.setMetadata("rowLines", encodeZeros(normalized.getChildren().size() + 1));
        return normalized;
    }

    /**
     * 构建十字交叉的嵌套矩阵结构。
     *
     * <p>外层是 5x1 矩阵，每一行的唯一单元格内再放一个 1x3 矩阵，
     * 左右两个位置各自再嵌套一个 1x2 小矩阵，尽量贴近参考 MathML 结构。</p>
     */
    public LaTeXNode buildCrossMultiplicationArray(CrossMultiplicationLayout layout) {
        LaTeXNode outerArray = createPreservedArray(1, "l", Math.max(layout == null ? 0 : layout.rows().size(), 1));
        if (layout == null || layout.rows().isEmpty()) {
            outerArray.addChild(buildStackRow(new LaTeXNode(LaTeXNode.Type.GROUP)));
            outerArray.setMetadata("rowLines", encodeZeros(outerArray.getChildren().size() + 1));
            return outerArray;
        }
        for (List<String> rowCells : layout.rows()) {
            LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
            row.addChild(buildNodeCell(buildCrossMultiplicationLine(rowCells)));
            outerArray.addChild(row);
        }
        outerArray.setMetadata("rowLines", encodeZeros(outerArray.getChildren().size() + 1));
        return outerArray;
    }

    /**
     * 创建只包含头部的长除法节点（无步骤），用于 computed array 路径。
     * tmLDIV 模板只包含：除数（外部）、商（slot 0）、被除数（slot 1，仅数字）。
     * 步骤在 tmLDIV 之外，作为外层矩阵的后续行。
     */
    private LaTeXNode buildSimpleLongDivisionHeader(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        LaTeXNode header = new LaTeXNode(LaTeXNode.Type.LONG_DIVISION,
            longDivisionNode == null ? "\\longdiv" : longDivisionNode.getValue());
        if (longDivisionNode != null) {
            header.getMetadata().putAll(longDivisionNode.getMetadata());
        }
        header.setMetadata(LONG_DIVISION_HEADER_ONLY, "true");

        // 除数
        header.addChild(deepCopy(childAt(longDivisionNode, 0)));

        // 商（简化，不含步骤）
        String quotient = spec == null || spec.longDivisionHeader() == null ? "" : spec.longDivisionHeader().quotient();
        if (!quotient.isBlank()) {
            LaTeXNode quotientNode = new LaTeXNode(LaTeXNode.Type.GROUP);
            for (char ch : quotient.toCharArray()) {
                quotientNode.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
            }
            header.addChild(quotientNode);
        } else {
            header.addChild(new LaTeXNode(LaTeXNode.Type.GROUP));
        }

        // 被除数（简化，仅数字，不含 PILE/FRACTION 步骤）
        String dividend = flattenNodeText(childAt(longDivisionNode, 2));
        LaTeXNode dividendNode = new LaTeXNode(LaTeXNode.Type.GROUP);
        for (char ch : dividend.toCharArray()) {
            dividendNode.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        header.addChild(dividendNode);

        return header;
    }

    public LaTeXNode buildStructuredLongDivisionNode(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        group.addChild(buildSimpleLongDivisionHeader(longDivisionNode, spec));
        if (spec != null) {
            if (!spec.longDivisionSteps().isEmpty()) {
                for (int stepIndex = 0; stepIndex < spec.longDivisionSteps().size(); stepIndex++) {
                    LongDivisionStep step = spec.longDivisionSteps().get(stepIndex);
                    group.addChild(buildStructuredLongDivisionLine(step.remainderRow(), false));
                    group.addChild(buildStructuredLongDivisionLine(step.productRow(), false, stepIndex == 0));
                }
            } else {
                for (int index = 1; index < spec.rows().size(); index++) {
                    group.addChild(buildStructuredLongDivisionLine(spec.rows().get(index), false));
                }
            }
        }
        return group;
    }

    public LaTeXNode buildStructuredLongDivisionTemplateNode(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        LaTeXNode structured = new LaTeXNode(LaTeXNode.Type.LONG_DIVISION,
            longDivisionNode == null ? "\\longdiv" : longDivisionNode.getValue());
        if (longDivisionNode != null) {
            structured.getMetadata().putAll(longDivisionNode.getMetadata());
        }
        structured.setMetadata("forceSimpleTemplate", "true");

        structured.addChild(deepCopy(childAt(longDivisionNode, 0)));
        structured.addChild(buildLongDivisionQuotientNode(spec));
        // 方案 A：头部继续使用 tmLDIV，步骤区改为槽位内的 PILE + RULER + Tab。
        structured.addChild(buildStructuredLongDivisionDividendPile(longDivisionNode, spec));
        return structured;
    }

    private LaTeXNode buildCell(String text) {
        LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
        if (text == null || text.isEmpty()) {
            return cell;
        }
        if (VerticalLayoutCompiler.EXPLICIT_EMPTY_CELL_TOKEN.equals(text)) {
            // 显式 {} 占位需要保留单元格宽度，但不能显示可见字符。
            cell.addChild(buildTextNode(" "));
            return cell;
        }
        if (looksLikeSingleCommand(text)) {
            cell.addChild(new LaTeXNode(LaTeXNode.Type.COMMAND, text));
            return cell;
        }
        for (char ch : text.toCharArray()) {
            cell.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        return cell;
    }

    private LaTeXNode buildCrossMultiplicationLine(List<String> rowCells) {
        List<String> normalized = normalizeCrossRow(rowCells);
        LaTeXNode line = createPreservedArray(3, "ccc", 1);
        LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
        row.addChild(buildNodeCell(buildCrossMultiplicationPair(normalized.get(0), normalized.get(1))));
        row.addChild(buildCell(normalized.get(2)));
        row.addChild(buildNodeCell(buildCrossMultiplicationPair(normalized.get(3), normalized.get(4))));
        line.addChild(row);
        return line;
    }

    private LaTeXNode buildCrossMultiplicationPair(String left, String right) {
        LaTeXNode pair = createPreservedArray(2, "cc", 1);
        LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
        row.addChild(buildCell(left));
        row.addChild(buildCell(right));
        pair.addChild(row);
        return pair;
    }

    private List<String> normalizeCrossRow(List<String> rowCells) {
        List<String> normalized = new ArrayList<>(rowCells == null ? List.of() : rowCells);
        while (normalized.size() < 5) {
            normalized.add("");
        }
        if (normalized.size() > 5) {
            return new ArrayList<>(normalized.subList(0, 5));
        }
        return normalized;
    }

    private LaTeXNode createPreservedArray(int columnCount, String columnSpec, int rowCount) {
        LaTeXNode array = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        int safeColumnCount = Math.max(columnCount, 1);
        int safeRowCount = Math.max(rowCount, 1);
        array.setMetadata("columnCount", String.valueOf(safeColumnCount));
        array.setMetadata("columnSpec", columnSpec == null || columnSpec.isBlank() ? "c".repeat(safeColumnCount) : columnSpec);
        array.setMetadata("columnLines", encodeZeros(safeColumnCount + 1));
        array.setMetadata("rowLines", encodeZeros(safeRowCount + 1));
        array.setMetadata(PRESERVE_RAW_ARRAY, "true");
        return array;
    }

    private LaTeXNode copyCellContent(LaTeXNode cell, boolean markLongDivisionHeaderOnly) {
        LaTeXNode copy = new LaTeXNode(LaTeXNode.Type.CELL);
        if (cell == null) {
            return copy;
        }
        for (LaTeXNode child : cell.getChildren()) {
            LaTeXNode duplicated = deepCopy(child);
            if (markLongDivisionHeaderOnly && duplicated.getType() == LaTeXNode.Type.LONG_DIVISION) {
                duplicated = markLongDivisionHeaderOnly(duplicated);
            }
            copy.addChild(duplicated);
        }
        return copy;
    }

    private LaTeXNode deepCopy(LaTeXNode node) {
        LaTeXNode copy = new LaTeXNode(node.getType(), node.getValue());
        copy.getMetadata().putAll(node.getMetadata());
        for (LaTeXNode child : node.getChildren()) {
            copy.addChild(deepCopy(child));
        }
        return copy;
    }

    private String resolveColumnSpec(VerticalLayoutSpec spec) {
        if (spec.kind() == VerticalLayoutSpec.Kind.DECIMAL) {
            return "r";
        }
        return "r".repeat(Math.max(spec.columnCount(), 0));
    }

    private int resolveColumnCount(VerticalLayoutSpec spec) {
        return spec.columnCount();
    }

    private java.util.List<String> resolveRowCells(VerticalLayoutSpec spec, VerticalRow rowSpec) {
        return rowSpec.cells();
    }

    private LaTeXNode buildLongDivisionQuotientNode(VerticalLayoutSpec spec) {
        String quotient = spec == null || spec.longDivisionHeader() == null ? "" : spec.longDivisionHeader().quotient();
        LaTeXNode quotientNode = new LaTeXNode(LaTeXNode.Type.GROUP);
        if (quotient.isBlank()) {
            return quotientNode;
        }
        for (char ch : quotient.toCharArray()) {
            quotientNode.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        return quotientNode;
    }

    private LaTeXNode buildStructuredLongDivisionDividendBody(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        if (spec == null || spec.rows().isEmpty()) {
            LaTeXNode fallback = new LaTeXNode(LaTeXNode.Type.GROUP);
            fallback.addChild(deepCopy(childAt(longDivisionNode, 2)));
            return fallback;
        }
        return buildStructuredLongDivisionLines(spec);
    }

    private LaTeXNode buildStructuredLongDivisionDividendPile(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        if (spec == null || spec.rows().isEmpty()) {
            return buildStructuredLongDivisionDividendBody(longDivisionNode, spec);
        }

        List<LaTeXNode> lines = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < spec.rows().size(); rowIndex++) {
            VerticalRow rowSpec = spec.rows().get(rowIndex);
            List<VerticalSegment> segments = trimDivisionSegments(rowSpec.segments());
            lines.add(buildRawLine(buildTabbedSegmentGroup(segments, hasRuleBelowRow(spec, rowIndex))));
        }

        if (lines.isEmpty()) {
            return buildStructuredLongDivisionDividendBody(longDivisionNode, spec);
        }
        return buildRawPile(lines, 1, 2, trimDivisionTabStops(spec.tabStops()));
    }

    private LaTeXNode buildStructuredLongDivisionLines(VerticalLayoutSpec spec) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        for (int index = 0; index < spec.rows().size(); index++) {
            group.addChild(buildStructuredLongDivisionLine(spec.rows().get(index), index == 0));
        }
        return group;
    }

    private LaTeXNode buildStructuredLongDivisionLine(VerticalRow rowSpec, boolean inlineFirstRow) {
        return buildStructuredLongDivisionLine(rowSpec, inlineFirstRow, false);
    }

    private LaTeXNode buildStructuredLongDivisionLine(VerticalRow rowSpec, boolean inlineFirstRow, boolean firstStepProduct) {
        List<String> trimmedCells = trimDivisionColumns(rowSpec.cells());
        int firstContentIndex = findFirstNonEmptyCell(trimmedCells);
        String rowText = joinNonEmptyCells(trimmedCells);

        if (rowSpec.kind() == RowKind.STEP && firstContentIndex >= 0 && !rowText.isEmpty()) {
            LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
            group.setMetadata(LONG_DIVISION_STEP_SPLIT, "true");
            if (firstStepProduct) {
                group.setMetadata(LONG_DIVISION_FIRST_STEP_SPLIT, "true");
            }
            group.addChild(buildRawLine(buildTextGroup(List.of(rowText))));

            LaTeXNode underline = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\underline");
            underline.setMetadata(RAW_INLINE_CONTAINER, "true");
            underline.addChild(buildRawLine(buildTextGroup(List.of(rowText))));
            group.addChild(underline);
            return group;
        }
        LaTeXNode content = buildTextGroup(List.of(rowText));
        return inlineFirstRow ? content : buildRawLine(content);
    }

    private LaTeXNode buildStructuredLongDivisionMatrix(VerticalLayoutSpec spec) {
        LaTeXNode array = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        int columnCount = Math.max(resolveTrimmedColumnCount(spec), 1);
        array.setMetadata("columnCount", String.valueOf(columnCount));
        array.setMetadata("columnSpec", "r".repeat(columnCount));
        array.setMetadata("columnLines", encodeZeros(columnCount + 1));
        array.setMetadata("rowLines", encodeZeros(spec.rows().size() + 1));
        array.setMetadata(PRESERVE_RAW_ARRAY, "true");
        array.setMetadata(RAW_MATRIX_HALIGN, "0");

        for (VerticalRow rowSpec : spec.rows()) {
            array.addChild(buildStructuredLongDivisionRow(rowSpec));
        }
        return array;
    }

    private LaTeXNode buildStructuredLongDivisionRow(VerticalRow rowSpec) {
        LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
        List<String> trimmedCells = trimDivisionColumns(rowSpec.cells());
        int firstContentIndex = findFirstNonEmptyCell(trimmedCells);
        String rowText = joinNonEmptyCells(trimmedCells);

        if (rowSpec.kind() == RowKind.STEP && firstContentIndex >= 0 && !rowText.isEmpty()) {
            for (int index = 0; index < trimmedCells.size(); index++) {
                LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
                if (index == firstContentIndex) {
                    LaTeXNode underline = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\underline");
                    underline.addChild(buildTextGroup(List.of(rowText)));
                    cell.addChild(underline);
                }
                row.addChild(cell);
            }
            if (trimmedCells.isEmpty()) {
                row.addChild(new LaTeXNode(LaTeXNode.Type.CELL));
            }
            return row;
        }

        for (String text : trimmedCells) {
            row.addChild(buildCell(text));
        }
        if (trimmedCells.isEmpty()) {
            row.addChild(new LaTeXNode(LaTeXNode.Type.CELL));
        }
        return row;
    }

    private LaTeXNode buildStepsMatrixWithUnderlines(VerticalLayoutSpec spec) {
        LaTeXNode stepsArray = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        stepsArray.setMetadata("columnCount", "1");
        stepsArray.setMetadata("columnSpec", "l");
        stepsArray.setMetadata("columnLines", "0,0");
        stepsArray.setMetadata(PRESERVE_RAW_ARRAY, "true");

        List<LaTeXNode> stepRows = new ArrayList<>();
        for (LongDivisionStep step : spec.longDivisionSteps()) {
            List<String> productCells = trimDivisionColumns(step.productRow().cells());
            List<String> remainderCells = trimDivisionColumns(step.remainderRow().cells());

            LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
            LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);

            LaTeXNode productMatrix = buildSingleRowMatrix(productCells, 0);
            LaTeXNode remainderMatrix = buildSingleRowMatrix(remainderCells, 0);

            LaTeXNode underline = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\underline");
            underline.addChild(productMatrix);

            LaTeXNode productLine = buildRawLine(underline);
            LaTeXNode remainderLine = buildRawLine(remainderMatrix);

            LaTeXNode pile = buildRawPile(List.of(productLine, remainderLine), 0, 0);
            cell.addChild(pile);
            row.addChild(cell);
            stepRows.add(row);
        }

        StringBuilder rowLines = new StringBuilder("0");
        for (int i = 0; i < stepRows.size(); i++) {
            rowLines.append(",0");
        }
        stepsArray.setMetadata("rowLines", rowLines.toString());

        for (LaTeXNode row : stepRows) {
            stepsArray.addChild(row);
        }

        return stepsArray;
    }

    private List<LaTeXNode> buildComputedLongDivisionRows(VerticalLayoutSpec spec) {
        List<LaTeXNode> rows = new ArrayList<>();
        for (int stepIndex = 0; stepIndex < spec.longDivisionSteps().size(); stepIndex++) {
            LongDivisionStep step = spec.longDivisionSteps().get(stepIndex);
            String productText = buildFormulaAlignedDivisionText(step.productRow());
            if (!productText.isBlank()) {
                rows.add(buildSingleCellRow(buildIndentedUnderlineLine(productText)));
            }

            String remainderText = buildFormulaAlignedDivisionText(step.remainderRow());
            if (!remainderText.isBlank()) {
                rows.add(buildSingleCellRow(buildTextNode(remainderText)));
            }
        }
        return rows;
    }

    private LaTeXNode buildSingleCellRow(LaTeXNode content) {
        LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
        LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
        if (content != null) {
            cell.addChild(content);
        }
        row.addChild(cell);
        return row;
    }

    private LaTeXNode buildIndentedUnderlineLine(String text) {
        int prefixLength = countLeadingSpaces(text);
        String prefix = text.substring(0, prefixLength);
        String underlinedText = text.substring(prefixLength);

        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        if (!prefix.isEmpty()) {
            group.addChild(buildTextNode(prefix));
        }
        if (!underlinedText.isEmpty()) {
            LaTeXNode underline = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\underline");
            underline.addChild(buildTextNode(underlinedText));
            group.addChild(underline);
        }
        return group;
    }

    private LaTeXNode buildTextNode(String text) {
        LaTeXNode textNode = new LaTeXNode(LaTeXNode.Type.TEXT, "\\text");
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        if (text != null) {
            for (char ch : text.toCharArray()) {
                group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
            }
        }
        textNode.addChild(group);
        return textNode;
    }

    private String buildFormulaAlignedDivisionText(VerticalRow row) {
        if (row == null || row.cells().isEmpty()) {
            return "";
        }
        List<String> trimmedCells = trimDivisionColumns(row.cells());
        int leadingColumns = Math.max(findFirstNonEmptyCell(trimmedCells), 0);
        String digits = joinNonEmptyCells(trimmedCells);
        if (digits.isEmpty()) {
            return "";
        }

        // 统一空格公式：
        // 1. 每跨 1 个自然列，放大成 3 个空格；
        // 2. 再按位数补偿，保证个位列继续对齐。
        // 因而：
        // 1 位数 = 3*c - 1
        // 2 位数 = 3*c
        // 3 位数 = 3*c + 1
        // n 位数 = 3*c + n - 2
        int spaces = Math.max(leadingColumns * 3 + digits.length() - 2, 0);
        return " ".repeat(spaces) + digits;
    }

    private int countLeadingSpaces(String text) {
        int index = 0;
        while (index < text.length() && text.charAt(index) == ' ') {
            index++;
        }
        return index;
    }

    private VerticalLayoutSpec toComputedLongDivisionSteps(VerticalLayoutSpec spec) {
        if (spec == null || !spec.isLongDivision()) {
            return null;
        }
        int trimmedColumns = Math.max(spec.columnCount() - 2, 0);
        List<VerticalRow> stepRows = new ArrayList<>();
        for (VerticalRow row : spec.rows()) {
            if (row.kind() == RowKind.DIVIDEND) {
                continue;
            }
            List<String> trimmedCells = new ArrayList<>();
            for (int col = 2; col < row.cells().size(); col++) {
                trimmedCells.add(row.cells().get(col));
            }
            stepRows.add(new VerticalRow(row.kind(), trimmedCells));
        }
        if (stepRows.isEmpty()) {
            return new VerticalLayoutSpec(VerticalLayoutSpec.Kind.ARITHMETIC, trimmedColumns, -1, -1,
                List.of(), List.of(), List.of(), null);
        }

        List<RuleSpan> shiftedSpans = new ArrayList<>();
        for (RuleSpan span : spec.ruleSpans()) {
            int shiftedBoundary = span.boundaryIndex() - 1;
            if (shiftedBoundary <= 0 || shiftedBoundary > stepRows.size()) {
                continue;
            }
            shiftedSpans.add(new RuleSpan(shiftedBoundary,
                Math.max(span.startColumn() - 2, 0),
                Math.max(span.endColumn() - 2, 0)));
        }
        return new VerticalLayoutSpec(VerticalLayoutSpec.Kind.ARITHMETIC, trimmedColumns,
            Math.max(spec.anchorColumn() - 2, trimmedColumns - 1), -1,
            List.of(), stepRows, shiftedSpans, null);
    }

    private LaTeXNode markLongDivisionHeaderOnly(LaTeXNode node) {
        node.setMetadata(LONG_DIVISION_HEADER_ONLY, "true");
        return node;
    }

    private LaTeXNode buildSingleRowMatrix(List<String> cells) {
        return buildSingleRowMatrix(cells, null);
    }

    private LaTeXNode buildTextGroup(List<String> cells) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        if (cells == null || cells.isEmpty()) {
            return group;
        }
        for (String cell : cells) {
            if (cell == null || cell.isEmpty()) {
                continue;
            }
            if (looksLikeSingleCommand(cell)) {
                group.addChild(new LaTeXNode(LaTeXNode.Type.COMMAND, cell));
                continue;
            }
            for (char ch : cell.toCharArray()) {
                group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
            }
        }
        return group;
    }

    private LaTeXNode buildSingleRowMatrix(List<String> cells, Integer horizontalAlignment) {
        LaTeXNode array = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        int columnCount = Math.max(cells.size(), 1);
        array.setMetadata("columnCount", String.valueOf(columnCount));
        array.setMetadata("columnSpec", "r".repeat(columnCount));
        array.setMetadata("columnLines", encodeZeros(columnCount + 1));
        array.setMetadata("rowLines", "0,0");
        if (horizontalAlignment != null) {
            array.setMetadata(PRESERVE_RAW_ARRAY, "true");
            array.setMetadata(RAW_MATRIX_HALIGN, String.valueOf(horizontalAlignment));
        }

        LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
        for (String cell : cells) {
            row.addChild(buildCell(cell));
        }
        if (cells.isEmpty()) {
            row.addChild(buildCell(""));
        }
        array.addChild(row);
        return array;
    }

    private LaTeXNode buildFractionNode(LaTeXNode numerator, LaTeXNode denominator) {
        LaTeXNode fraction = new LaTeXNode(LaTeXNode.Type.FRACTION, "\\frac");
        fraction.addChild(numerator == null ? new LaTeXNode(LaTeXNode.Type.GROUP) : numerator);
        fraction.addChild(denominator == null ? new LaTeXNode(LaTeXNode.Type.GROUP) : denominator);
        return fraction;
    }

    private LaTeXNode buildStackRow(LaTeXNode content) {
        LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
        row.addChild(buildNodeCell(content));
        return row;
    }

    private LaTeXNode buildNodeCell(LaTeXNode content) {
        LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
        if (content != null) {
            cell.addChild(content);
        }
        return cell;
    }

    private LaTeXNode buildRawPile(List<LaTeXNode> children, int halign, int valign) {
        return buildRawPile(children, halign, valign, List.of());
    }

    private LaTeXNode buildRawPile(List<LaTeXNode> children, int halign, int valign, List<VerticalTabStop> tabStops) {
        LaTeXNode pile = new LaTeXNode(LaTeXNode.Type.GROUP);
        pile.setMetadata(RAW_PILE_CONTAINER, "true");
        pile.setMetadata(RAW_PILE_HALIGN, String.valueOf(halign));
        pile.setMetadata(RAW_PILE_VALIGN, String.valueOf(valign));
        if (tabStops != null && !tabStops.isEmpty()) {
            pile.setMetadata(RAW_PILE_RULER, encodeRawPileRuler(tabStops));
        }
        if (children != null) {
            for (LaTeXNode child : children) {
                if (child != null) {
                    pile.addChild(child);
                }
            }
        }
        return pile;
    }

    private LaTeXNode buildRawLine(LaTeXNode content) {
        LaTeXNode line = new LaTeXNode(LaTeXNode.Type.GROUP);
        line.setMetadata(RAW_LINE_CONTAINER, "true");
        if (content != null) {
            line.addChild(content);
        }
        return line;
    }

    private LaTeXNode buildTabbedSegmentGroup(List<VerticalSegment> segments, boolean underline) {
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        int currentStop = -1;
        for (VerticalSegment segment : segments) {
            int targetStop = Math.max(segment.endColumn(), currentStop);
            while (currentStop < targetStop) {
                group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, "\t"));
                currentStop++;
            }

            LaTeXNode textGroup = buildTextGroup(List.of(segment.text()));
            if (underline) {
                LaTeXNode underlineNode = new LaTeXNode(LaTeXNode.Type.COMMAND, "\\underline");
                underlineNode.addChild(textGroup);
                group.addChild(underlineNode);
            } else {
                group.addChild(textGroup);
            }
        }

        if (!segments.isEmpty()) {
            group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, "\t"));
        }
        return group;
    }

    private List<VerticalSegment> trimDivisionSegments(List<VerticalSegment> segments) {
        List<VerticalSegment> trimmed = new ArrayList<>();
        if (segments == null) {
            return trimmed;
        }
        for (VerticalSegment segment : segments) {
            if (segment.endColumn() < 2) {
                continue;
            }
            int startColumn = Math.max(segment.startColumn() - 2, 0);
            int endColumn = Math.max(segment.endColumn() - 2, startColumn);
            trimmed.add(new VerticalSegment(segment.text(), startColumn, endColumn, segment.underlined()));
        }
        return trimmed;
    }

    private List<VerticalTabStop> trimDivisionTabStops(List<VerticalTabStop> tabStops) {
        List<VerticalTabStop> trimmed = new ArrayList<>();
        if (tabStops == null) {
            return trimmed;
        }
        for (VerticalTabStop tabStop : tabStops) {
            if (tabStop.column() < 2) {
                continue;
            }
            trimmed.add(new VerticalTabStop(
                Math.max(tabStop.column() - 2, 0),
                tabStop.kind(),
                Math.max(tabStop.offset() - 480, 0)
            ));
        }
        return trimmed;
    }

    private boolean hasRuleBelowRow(VerticalLayoutSpec spec, int rowIndex) {
        int boundaryIndex = rowIndex + 1;
        return spec.ruleSpans().stream().anyMatch(span -> span.boundaryIndex() == boundaryIndex);
    }

    private String encodeRawPileRuler(List<VerticalTabStop> tabStops) {
        StringBuilder builder = new StringBuilder();
        for (VerticalTabStop tabStop : tabStops) {
            if (builder.length() > 0) {
                builder.append(';');
            }
            builder.append(encodeTabStopKind(tabStop.kind()))
                .append(':')
                .append(tabStop.offset());
        }
        return builder.toString();
    }

    private int encodeTabStopKind(TabStopKind kind) {
        return switch (kind == null ? TabStopKind.RIGHT : kind) {
            case LEFT -> 0;
            case CENTER -> 1;
            case RIGHT -> 2;
            case RELATION -> 3;
            case DECIMAL -> 4;
        };
    }

    /**
     * 构建步骤矩阵，每步是一个 FRACTION 模板（分子为乘积，分母为余数+下移数字）。
     * 矩阵只有一列，每行包含一个 FRACTION 节点。
     */
    private LaTeXNode buildStepsMatrixWithFractions(VerticalLayoutSpec spec) {
        LaTeXNode stepsArray = new LaTeXNode(LaTeXNode.Type.ARRAY, "\\array");
        stepsArray.setMetadata("columnCount", "1");
        stepsArray.setMetadata("columnSpec", "l");
        stepsArray.setMetadata("columnLines", "0,0");
        // 关键：标记为保留原始数组，避免被 compileArray 重新处理
        stepsArray.setMetadata(PRESERVE_RAW_ARRAY, "true");

        List<LaTeXNode> stepRows = new ArrayList<>();
        for (LongDivisionStep step : spec.longDivisionSteps()) {
            LaTeXNode row = new LaTeXNode(LaTeXNode.Type.ROW);
            LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);

            // 每个 FRACTION：分子=乘积行，分母=余数行
            List<String> productCells = trimDivisionColumns(step.productRow().cells());
            List<String> remainderCells = trimDivisionColumns(step.remainderRow().cells());

            // 使用带 horizontalAlignment 参数的版本，确保 PRESERVE_RAW_ARRAY 被设置
            LaTeXNode numerator = buildSingleRowMatrix(productCells, 0);
            LaTeXNode denominator = buildSingleRowMatrix(remainderCells, 0);
            cell.addChild(buildFractionNode(numerator, denominator));

            row.addChild(cell);
            stepRows.add(row);
        }

        // 添加行线和列线元数据
        int rowCount = stepRows.size();
        StringBuilder rowLines = new StringBuilder("0");
        for (int i = 0; i < rowCount; i++) {
            rowLines.append(",0");
        }
        stepsArray.setMetadata("rowLines", rowLines.toString());

        for (LaTeXNode row : stepRows) {
            stepsArray.addChild(row);
        }

        return stepsArray;
    }

    private List<String> trimDivisionColumns(List<String> cells) {
        List<String> trimmed = new ArrayList<>();
        if (cells == null) {
            return trimmed;
        }
        for (int col = 2; col < cells.size(); col++) {
            trimmed.add(cells.get(col));
        }
        return trimmed;
    }

    private List<String> placeRightAligned(String text, int columnCount) {
        int targetColumns = Math.max(columnCount, text == null ? 0 : text.length());
        List<String> cells = new ArrayList<>();
        for (int col = 0; col < targetColumns; col++) {
            cells.add("");
        }
        if (text == null || text.isEmpty()) {
            return cells;
        }
        int startColumn = Math.max(targetColumns - text.length(), 0);
        for (int index = 0; index < text.length() && startColumn + index < cells.size(); index++) {
            cells.set(startColumn + index, String.valueOf(text.charAt(index)));
        }
        return cells;
    }

    private int resolveLongDivisionDigitColumns(VerticalLayoutSpec spec) {
        return Math.max(spec == null ? 0 : spec.columnCount() - 2, 1);
    }

    private int resolveTrimmedColumnCount(VerticalLayoutSpec spec) {
        if (spec == null) {
            return 1;
        }
        int maxColumns = resolveLongDivisionDigitColumns(spec);
        for (VerticalRow row : spec.rows()) {
            maxColumns = Math.max(maxColumns, trimDivisionColumns(row.cells()).size());
        }
        return Math.max(maxColumns, 1);
    }

    private int findFirstNonEmptyCell(List<String> cells) {
        if (cells == null) {
            return -1;
        }
        for (int index = 0; index < cells.size(); index++) {
            String cell = cells.get(index);
            if (cell != null && !cell.isEmpty()) {
                return index;
            }
        }
        return -1;
    }

    private String joinNonEmptyCells(List<String> cells) {
        if (cells == null || cells.isEmpty()) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        for (String cell : cells) {
            if (cell != null && !cell.isEmpty()) {
                builder.append(cell);
            }
        }
        return builder.toString();
    }


    private String flattenNodeText(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR || node.getType() == LaTeXNode.Type.COMMAND) {
            return node.getValue() == null ? "" : node.getValue();
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flattenNodeText(child));
        }
        return builder.toString();
    }

    private boolean looksLikeSingleCommand(String text) {
        return text != null && text.startsWith("\\") && text.indexOf(' ') < 0;
    }

    private int findFirstRowIndex(List<VerticalRow> rows, RowKind kind) {
        for (int index = 0; index < rows.size(); index++) {
            if (rows.get(index).kind() == kind) {
                return index;
            }
        }
        return -1;
    }

    private LaTeXNode childAt(LaTeXNode node, int index) {
        if (node == null || index < 0 || index >= node.getChildren().size()) {
            return null;
        }
        return node.getChildren().get(index);
    }

    private String encodeZeros(int size) {
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < Math.max(size, 1); i++) {
            if (i > 0) {
                builder.append(',');
            }
            builder.append('0');
        }
        return builder.toString();
    }

    private String encodeRowLines(int rowCount, java.util.List<RuleSpan> ruleSpans) {
        int[] parts = new int[Math.max(rowCount + 1, 1)];
        for (RuleSpan span : ruleSpans) {
            if (span.boundaryIndex() >= 0 && span.boundaryIndex() < parts.length) {
                parts[span.boundaryIndex()] = 1;
            }
        }
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) {
                builder.append(',');
            }
            builder.append(parts[i]);
        }
        return builder.toString();
    }
}
