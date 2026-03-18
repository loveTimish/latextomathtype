package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RuleSpan;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RowKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionStep;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;

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
    public static final String RAW_PILE_HALIGN = "rawPileHalign";
    public static final String RAW_PILE_VALIGN = "rawPileValign";
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
        normalized.setMetadata("rowLines", "0,0,0");

        LaTeXNode firstRow = new LaTeXNode(LaTeXNode.Type.ROW);
        LaTeXNode headerCell = new LaTeXNode(LaTeXNode.Type.CELL);
        headerCell.addChild(markLongDivisionHeaderOnly(deepCopy(longDivisionNode)));
        firstRow.addChild(headerCell);
        normalized.addChild(firstRow);

        VerticalLayoutSpec stepSpec = toComputedLongDivisionSteps(spec);
        if (stepSpec != null && !stepSpec.rows().isEmpty()) {
            LaTeXNode secondRow = new LaTeXNode(LaTeXNode.Type.ROW);
            secondRow.addChild(buildArrayNode(stepSpec));
            normalized.addChild(secondRow);
        }
        return normalized;
    }

    public LaTeXNode buildStructuredLongDivisionNode(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        LaTeXNode structured = new LaTeXNode(LaTeXNode.Type.LONG_DIVISION,
            longDivisionNode == null ? "\\longdiv" : longDivisionNode.getValue());
        if (longDivisionNode != null) {
            structured.getMetadata().putAll(longDivisionNode.getMetadata());
        }
        structured.setMetadata(LONG_DIVISION_HEADER_ONLY, "true");
        structured.addChild(deepCopy(childAt(longDivisionNode, 0)));
        structured.addChild(buildLongDivisionQuotientNode(spec));
        structured.addChild(buildLongDivisionDividendStack(longDivisionNode, spec));
        return structured;
    }

    private LaTeXNode buildCell(String text) {
        LaTeXNode cell = new LaTeXNode(LaTeXNode.Type.CELL);
        if (text == null || text.isEmpty()) {
            return cell;
        }
        for (char ch : text.toCharArray()) {
            cell.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        return cell;
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
        if (quotient.isBlank()) {
            return new LaTeXNode(LaTeXNode.Type.GROUP);
        }
        int digitColumns = resolveLongDivisionDigitColumns(spec);
        return buildRawPile(List.of(buildSingleRowMatrix(placeRightAligned(quotient, digitColumns), 1)), 0, 0);
    }

    private LaTeXNode buildLongDivisionDividendStack(LaTeXNode longDivisionNode, VerticalLayoutSpec spec) {
        // 长除法被除数槽需要直接构造 PILE/LINE 子树，避免被 ARRAY 正规化成 MATRIX。
        LaTeXNode stack = buildRawPile(new ArrayList<>(), 1, 1);

        if (spec == null || spec.rows().isEmpty()) {
            stack.addChild(buildRawLine(deepCopy(childAt(longDivisionNode, 2))));
            return stack;
        }

        int dividendIndex = findFirstRowIndex(spec.rows(), RowKind.DIVIDEND);
        if (dividendIndex >= 0) {
            stack.addChild(buildSingleRowMatrix(trimDivisionColumns(spec.rows().get(dividendIndex).cells()), 1));
        }

        for (LongDivisionStep step : spec.longDivisionSteps()) {
            LaTeXNode numerator = buildSingleRowMatrix(trimDivisionColumns(step.productRow().cells()));
            LaTeXNode denominator = buildSingleRowMatrix(trimDivisionColumns(step.remainderRow().cells()));
            stack.addChild(buildRawLine(buildFractionNode(numerator, denominator)));
        }
        return stack;
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
        LaTeXNode pile = new LaTeXNode(LaTeXNode.Type.GROUP);
        pile.setMetadata(RAW_PILE_CONTAINER, "true");
        pile.setMetadata(RAW_PILE_HALIGN, String.valueOf(halign));
        pile.setMetadata(RAW_PILE_VALIGN, String.valueOf(valign));
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
