package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutSpec.Kind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionHeader;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionStep;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RowKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RuleSpan;
import com.lz.paperword.core.layout.VerticalLayoutSpec.TabStopKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalSegment;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalTabStop;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/**
 * 将 LaTeX AST 编译为统一的竖式布局语义。
 */
public class VerticalLayoutCompiler {

    public static final String EXPLICIT_EMPTY_CELL_TOKEN = "@@PW_EMPTY_CELL@@";
    private static final int CROSS_MULTIPLICATION_SIZE = 5;

    /**
     * 竖式正规化阶段需要吞掉的 LaTeX 间距命令。
     * 这些命令只负责排版留白，不能落成可见正文，否则会出现“×quad3”这类脏文本。
     */
    private static final Set<String> SPACING_COMMANDS = Set.of(
        "\\,", "\\;", "\\:", "\\!", "\\quad", "\\qquad", "\\hspace", "\\hskip"
    );

    public VerticalLayoutSpec compile(LaTeXNode node) {
        LaTeXNode root = unwrapRoot(node);
        if (root == null) {
            return null;
        }
        if (root.getType() == LaTeXNode.Type.LONG_DIVISION) {
            return compileExplicitLongDivision(root);
        }
        if (root.getType() == LaTeXNode.Type.ARRAY) {
            return compileArray(root);
        }
        for (LaTeXNode child : root.getChildren()) {
            VerticalLayoutSpec compiled = compile(child);
            if (compiled != null) {
                return compiled;
            }
        }
        return null;
    }

    public VerticalLayoutSpec compileArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY) {
            return null;
        }
        if (isDecimalArray(node)) {
            return compileDecimalArray(node);
        }
        return compileArithmeticArray(node);
    }

    /**
     * 识别十字交叉数组并提取 5x5 原始单元格文本。
     *
     * <p>这里不返回通用竖式规格，而是保留原始行列信息，
     * 供上层直接重建成嵌套 MATRIX，避免被算术数组正规化压平。</p>
     */
    public CrossMultiplicationLayout compileCrossMultiplicationArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY) {
            return null;
        }
        List<List<String>> rows = extractRowTexts(node);
        if (rows.size() != CROSS_MULTIPLICATION_SIZE) {
            return null;
        }
        List<List<String>> normalized = new ArrayList<>();
        for (List<String> row : rows) {
            normalized.add(normalizeCrossRow(row));
        }
        if (!matchesCrossMultiplicationPattern(normalized)) {
            return null;
        }
        return new CrossMultiplicationLayout(normalized);
    }

    public VerticalLayoutSpec compileExplicitLongDivision(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.LONG_DIVISION) {
            return null;
        }

        String divisor = flattenNodeText(childAt(node, 0));
        String quotient = flattenNodeText(childAt(node, 1));
        String dividend = flattenNodeText(childAt(node, 2));
        LongDivisionHeader header = new LongDivisionHeader(divisor, quotient, dividend);

        if (!divisor.matches("\\d+") || !dividend.matches("\\d+")) {
            return new VerticalLayoutSpec(Kind.LONG_DIVISION, 0, -1, -1, List.of(), List.of(), List.of(), header);
        }

        int divisorValue;
        try {
            divisorValue = Integer.parseInt(divisor);
        } catch (NumberFormatException ignored) {
            return new VerticalLayoutSpec(Kind.LONG_DIVISION, 0, -1, -1, List.of(), List.of(), List.of(), header);
        }
        if (divisorValue == 0) {
            return new VerticalLayoutSpec(Kind.LONG_DIVISION, 0, -1, -1, List.of(), List.of(), List.of(), header);
        }

        int columnCount = dividend.length() + 2;
        int anchorColumn = columnCount - 1;
        List<VerticalRow> rows = new ArrayList<>();
        List<RuleSpan> ruleSpans = new ArrayList<>();
        List<LongDivisionStep> longDivisionSteps = new ArrayList<>();

        rows.add(buildPlacedRow(RowKind.DIVIDEND, dividend, anchorColumn, columnCount));
        int remainder = 0;
        boolean started = false;
        for (int index = 0; index < dividend.length(); index++) {
            remainder = remainder * 10 + (dividend.charAt(index) - '0');
            if (!started && remainder < divisorValue) {
                continue;
            }
            started = true;

            int product = (remainder / divisorValue) * divisorValue;
            VerticalRow productRow = buildPlacedRow(RowKind.STEP, String.valueOf(product), index + 2, columnCount);
            rows.add(productRow);
            ruleSpans.add(ruleForBoundaryBelowRow(rows.size() - 1));

            remainder -= product;
            VerticalRow remainderRow;
            if (index < dividend.length() - 1) {
                int broughtDown = remainder * 10 + (dividend.charAt(index + 1) - '0');
                remainderRow = buildPlacedRow(RowKind.REMAINDER, String.valueOf(broughtDown), index + 3, columnCount);
                rows.add(remainderRow);
            } else {
                remainderRow = buildPlacedRow(RowKind.REMAINDER, String.valueOf(remainder), index + 2, columnCount);
                rows.add(remainderRow);
            }
            longDivisionSteps.add(new LongDivisionStep(productRow, remainderRow));
        }

        rows = withSegments(rows, ruleSpans);
        return new VerticalLayoutSpec(Kind.LONG_DIVISION, columnCount, anchorColumn, -1,
            buildTabStops(columnCount, -1), rows, ruleSpans, header, withSegments(longDivisionSteps));
    }

    private VerticalLayoutSpec compileArithmeticArray(LaTeXNode node) {
        List<List<String>> rawRows = extractRowTexts(node);
        if (rawRows.isEmpty()) {
            return null;
        }

        int declaredCols = parseDeclaredColumnCount(node.getMetadata("columnCount"));
        int maxCols = rawRows.stream().mapToInt(List::size).max().orElse(0);
        int targetCols = Math.max(maxCols, declaredCols);
        List<VerticalRow> rows = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < rawRows.size(); rowIndex++) {
            List<String> row = new ArrayList<>(rawRows.get(rowIndex));
            int padLeft = targetCols - row.size();
            while (padLeft-- > 0) {
                row.add(0, "");
            }
            while (row.size() < targetCols) {
                row.add("");
            }
            rows.add(new VerticalRow(resolveArrayRowKind(rowIndex, rawRows.size()), row));
        }

        List<RuleSpan> ruleSpans = compileRuleSpans(rows, parsePartitionArray(node.getMetadata("rowLines"), rows.size() + 1));
        rows = withSegments(rows, ruleSpans);
        return new VerticalLayoutSpec(Kind.ARITHMETIC, targetCols, targetCols - 1, -1,
            buildTabStops(targetCols, -1), rows, ruleSpans, null);
    }

    private VerticalLayoutSpec compileDecimalArray(LaTeXNode node) {
        List<List<String>> rawRows = extractRowTexts(node);
        if (rawRows.isEmpty()) {
            return null;
        }

        List<VerticalRow> rows = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < rawRows.size(); rowIndex++) {
            List<String> row = List.of(buildDecimalRowText(rawRows.get(rowIndex)));
            rows.add(new VerticalRow(resolveArrayRowKind(rowIndex, rawRows.size()), row));
        }

        List<RuleSpan> ruleSpans = compileRuleSpans(rows, parsePartitionArray(node.getMetadata("rowLines"), rows.size() + 1));
        rows = withSegments(rows, ruleSpans);
        return new VerticalLayoutSpec(Kind.DECIMAL, 1, 0, 0,
            buildTabStops(1, 0), rows, ruleSpans, null);
    }

    private String buildDecimalRowText(List<String> rawRow) {
        String sign = "";
        String integer = "";
        String fraction = "";
        boolean hasDecimalPoint = false;

        if (rawRow.size() >= 4) {
            sign = blankToEmpty(rawRow.get(0));
            integer = blankToEmpty(rawRow.get(1));
            hasDecimalPoint = ".".equals(blankToEmpty(rawRow.get(2)));
            fraction = blankToEmpty(rawRow.get(3));
        } else {
            String first = rawRow.isEmpty() ? "" : blankToEmpty(rawRow.get(0));
            if (!first.isEmpty() && (first.charAt(0) == '+' || first.charAt(0) == '-')) {
                sign = String.valueOf(first.charAt(0));
                integer = first.substring(1);
            } else {
                integer = first;
            }
            hasDecimalPoint = rawRow.size() > 1 && ".".equals(blankToEmpty(rawRow.get(1)));
            fraction = rawRow.size() > 2 ? blankToEmpty(rawRow.get(2)) : "";
        }

        StringBuilder builder = new StringBuilder();
        builder.append(sign);
        builder.append(integer);
        if (hasDecimalPoint) {
            builder.append('.');
        }
        builder.append(fraction);
        return builder.toString();
    }

    private List<RuleSpan> compileRuleSpans(List<VerticalRow> rows, int[] rowLines) {
        List<RuleSpan> spans = new ArrayList<>();
        for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
            if (rowLines.length <= rowIndex || rowLines[rowIndex] <= 0) {
                continue;
            }
            spans.add(ruleForBoundary(rows, rowIndex));
        }
        return spans;
    }

    private RuleSpan ruleForBoundaryBelowRow(int rowIndex) {
        return new RuleSpan(rowIndex + 1, 0, 0);
    }

    private RuleSpan ruleForBoundary(List<VerticalRow> rows, int currentRowIndex) {
        VerticalRow previous = rows.get(currentRowIndex - 1);
        VerticalRow current = rows.get(currentRowIndex);
        int start = Math.min(firstNonEmptyColumn(previous.cells()), firstNonEmptyColumn(current.cells()));
        int end = Math.max(lastNonEmptyColumn(previous.cells()), lastNonEmptyColumn(current.cells()));
        if (start < 0) {
            start = 0;
        }
        if (end < start) {
            end = start;
        }
        return new RuleSpan(currentRowIndex, start, end);
    }

    private VerticalRow buildPlacedRow(RowKind kind, String text, int endColumn, int columnCount) {
        List<String> cells = new ArrayList<>();
        for (int i = 0; i < columnCount; i++) {
            cells.add("");
        }
        if (text == null || text.isEmpty() || endColumn < 0) {
            return new VerticalRow(kind, cells);
        }
        int startColumn = Math.max(endColumn - text.length() + 1, 0);
        for (int index = 0; index < text.length() && startColumn + index < cells.size(); index++) {
            cells.set(startColumn + index, String.valueOf(text.charAt(index)));
        }
        return new VerticalRow(kind, cells);
    }

    private List<VerticalTabStop> buildTabStops(int columnCount, int decimalColumn) {
        List<VerticalTabStop> stops = new ArrayList<>();
        int unit = 240;
        for (int col = 0; col < columnCount; col++) {
            TabStopKind kind = (decimalColumn == col) ? TabStopKind.DECIMAL : TabStopKind.RIGHT;
            stops.add(new VerticalTabStop(col, kind, (col + 1) * unit));
        }
        return stops;
    }

    private List<VerticalRow> withSegments(List<VerticalRow> rows, List<RuleSpan> ruleSpans) {
        List<VerticalRow> result = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            VerticalRow row = rows.get(rowIndex);
            result.add(new VerticalRow(row.kind(), row.cells(), buildSegments(row.cells())));
        }
        return result;
    }

    private List<LongDivisionStep> withSegments(List<LongDivisionStep> steps) {
        if (steps == null || steps.isEmpty()) {
            return List.of();
        }
        List<LongDivisionStep> result = new ArrayList<>();
        for (LongDivisionStep step : steps) {
            result.add(new LongDivisionStep(
                new VerticalRow(step.productRow().kind(), step.productRow().cells(), buildSegments(step.productRow().cells())),
                new VerticalRow(step.remainderRow().kind(), step.remainderRow().cells(), buildSegments(step.remainderRow().cells()))
            ));
        }
        return result;
    }

    private List<LongDivisionStep> offsetLongDivisionSteps(List<LongDivisionStep> steps, int offset) {
        if (steps == null || steps.isEmpty()) {
            return List.of();
        }
        List<LongDivisionStep> shiftedSteps = new ArrayList<>();
        for (LongDivisionStep step : steps) {
            shiftedSteps.add(new LongDivisionStep(
                shiftRow(step.productRow(), offset),
                shiftRow(step.remainderRow(), offset)
            ));
        }
        return withSegments(shiftedSteps);
    }

    private VerticalRow shiftRow(VerticalRow row, int offset) {
        List<String> cells = new ArrayList<>();
        for (int i = 0; i < offset; i++) {
            cells.add("");
        }
        cells.addAll(row.cells());
        return new VerticalRow(row.kind(), cells);
    }

    private List<VerticalSegment> buildSegments(List<String> cells) {
        List<VerticalSegment> segments = new ArrayList<>();
        int col = 0;
        while (col < cells.size()) {
            while (col < cells.size() && blankToEmpty(cells.get(col)).isEmpty()) {
                col++;
            }
            if (col >= cells.size()) {
                break;
            }
            int start = col;
            StringBuilder builder = new StringBuilder();
            while (col < cells.size() && !blankToEmpty(cells.get(col)).isEmpty()) {
                builder.append(cells.get(col));
                col++;
            }
            segments.add(new VerticalSegment(builder.toString(), start, col - 1, false));
        }
        return segments;
    }

    private RowKind resolveArrayRowKind(int rowIndex, int totalRows) {
        return rowIndex == totalRows - 1 ? RowKind.RESULT : RowKind.OPERAND;
    }

    private List<List<String>> extractRowTexts(LaTeXNode arrayNode) {
        List<List<String>> rows = new ArrayList<>();
        for (LaTeXNode rowNode : arrayNode.getChildren()) {
            List<String> row = new ArrayList<>();
            for (LaTeXNode cell : rowNode.getChildren()) {
                row.add(extractCellText(cell));
            }
            trimRight(row);
            rows.add(row);
        }
        return rows;
    }

    private boolean isDecimalArray(LaTeXNode node) {
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                if (".".equals(extractCellText(cell))) {
                    return true;
                }
            }
        }
        return false;
    }

    private LaTeXNode unwrapRoot(LaTeXNode node) {
        if (node == null) {
            return null;
        }
        if (node.getType() == LaTeXNode.Type.ROOT && node.getChildren().size() == 1) {
            return node.getChildren().get(0);
        }
        return node;
    }

    private int[] parsePartitionArray(String encoded, int expectedSize) {
        int[] parts = new int[Math.max(expectedSize, 1)];
        if (encoded == null || encoded.isBlank()) {
            return parts;
        }
        String[] values = encoded.split(",");
        for (int i = 0; i < values.length && i < parts.length; i++) {
            try {
                parts[i] = Integer.parseInt(values[i].trim());
            } catch (NumberFormatException ignored) {
                parts[i] = 0;
            }
        }
        return parts;
    }

    private LaTeXNode childAt(LaTeXNode node, int index) {
        if (node == null || index < 0 || index >= node.getChildren().size()) {
            return null;
        }
        return node.getChildren().get(index);
    }

    private int firstNonEmptyColumn(List<String> row) {
        for (int i = 0; i < row.size(); i++) {
            if (!blankToEmpty(row.get(i)).isEmpty()) {
                return i;
            }
        }
        return -1;
    }

    private int lastNonEmptyColumn(List<String> row) {
        for (int i = row.size() - 1; i >= 0; i--) {
            if (!blankToEmpty(row.get(i)).isEmpty()) {
                return i;
            }
        }
        return -1;
    }

    private void trimRight(List<String> row) {
        while (!row.isEmpty()
            && blankToEmpty(row.get(row.size() - 1)).isEmpty()
            && !EXPLICIT_EMPTY_CELL_TOKEN.equals(row.get(row.size() - 1))) {
            row.remove(row.size() - 1);
        }
    }

    private String extractCellText(LaTeXNode cell) {
        if (cell != null && "true".equals(cell.getMetadata("explicitEmptyCell"))) {
            return EXPLICIT_EMPTY_CELL_TOKEN;
        }
        return flattenNodeText(cell);
    }

    private String flattenNodeText(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR) {
            return blankToEmpty(node.getValue());
        }
        if (node.getType() == LaTeXNode.Type.COMMAND) {
            if ("\\times".equals(node.getValue())) {
                return "×";
            }
            // LaTeX 转义字符在竖式编译阶段要还原成真实字符，不能把反斜杠也带进最终单元格文本。
            if ("\\%".equals(node.getValue())) {
                return "%";
            }
            if ("\\_".equals(node.getValue())) {
                return "_";
            }
            if ("\\{".equals(node.getValue())) {
                return "{";
            }
            if ("\\}".equals(node.getValue())) {
                return "}";
            }
            // 间距命令只保留版式意义，不应变成正文字符。
            if (SPACING_COMMANDS.contains(node.getValue())) {
                return "";
            }
            return blankToEmpty(node.getValue());
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flattenNodeText(child));
        }
        return builder.toString();
    }

    private String blankToEmpty(String text) {
        if (text == null || text.isBlank() || EXPLICIT_EMPTY_CELL_TOKEN.equals(text)) {
            return "";
        }
        return text;
    }

    private List<String> normalizeCrossRow(List<String> row) {
        List<String> normalized = new ArrayList<>(row == null ? List.of() : row);
        while (normalized.size() < CROSS_MULTIPLICATION_SIZE) {
            normalized.add("");
        }
        if (normalized.size() > CROSS_MULTIPLICATION_SIZE) {
            return new ArrayList<>(normalized.subList(0, CROSS_MULTIPLICATION_SIZE));
        }
        return normalized;
    }

    private boolean matchesCrossMultiplicationPattern(List<List<String>> rows) {
        if (rows.size() != CROSS_MULTIPLICATION_SIZE) {
            return false;
        }
        if (!isPercentLike(rows.get(0).get(0)) || !isPercentLike(rows.get(0).get(4))) {
            return false;
        }
        if (!isDiagonalArrow(rows.get(1).get(1)) || !isDiagonalArrow(rows.get(1).get(3))) {
            return false;
        }
        if (!isPercentLike(rows.get(2).get(2))) {
            return false;
        }
        if (!isDiagonalArrow(rows.get(3).get(1)) || !isDiagonalArrow(rows.get(3).get(3))) {
            return false;
        }
        return isPercentLike(rows.get(4).get(0)) && isPercentLike(rows.get(4).get(4));
    }

    private boolean isPercentLike(String text) {
        return !blankToEmpty(text).isEmpty() && blankToEmpty(text).contains("%");
    }

    private boolean isDiagonalArrow(String text) {
        String normalized = blankToEmpty(text);
        return "\\nearrow".equals(normalized)
            || "\\searrow".equals(normalized)
            || "\\nwarrow".equals(normalized)
            || "\\swarrow".equals(normalized);
    }

    private int parseDeclaredColumnCount(String encoded) {
        if (encoded == null || encoded.isBlank()) {
            return 0;
        }
        try {
            return Integer.parseInt(encoded.trim());
        } catch (NumberFormatException ignored) {
            return 0;
        }
    }

    /**
     * 十字交叉布局的原始 5x5 文本快照。
     */
    public record CrossMultiplicationLayout(List<List<String>> rows) {
        public CrossMultiplicationLayout {
            rows = rows == null ? List.of() : List.copyOf(rows.stream().map(List::copyOf).toList());
        }
    }

}
