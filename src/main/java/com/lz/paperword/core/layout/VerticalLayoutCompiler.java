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

/**
 * 将 LaTeX AST 编译为统一的竖式布局语义。
 */
public class VerticalLayoutCompiler {

    public VerticalLayoutSpec compile(LaTeXNode node) {
        LaTeXNode root = unwrapRoot(node);
        if (root == null) {
            return null;
        }
        if (isCompositeLongDivisionArray(root)) {
            return compileCompositeLongDivision(root);
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

    public boolean isCompositeLongDivisionArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY || node.getChildren().size() < 2) {
            return false;
        }
        return findFirstDescendant(node.getChildren().get(0), LaTeXNode.Type.LONG_DIVISION) != null
            && findNestedArray(node.getChildren().get(1), node) != null;
    }

    public VerticalLayoutSpec compileArray(LaTeXNode node) {
        if (node == null || node.getType() != LaTeXNode.Type.ARRAY) {
            return null;
        }
        if (isDivisionArray(node)) {
            return compileLegacyDivisionArray(node);
        }
        if (isDecimalArray(node)) {
            return compileDecimalArray(node);
        }
        return compileArithmeticArray(node);
    }

    public VerticalLayoutSpec compileCompositeLongDivision(LaTeXNode node) {
        LaTeXNode longDiv = findFirstDescendant(node.getChildren().get(0), LaTeXNode.Type.LONG_DIVISION);
        LaTeXNode stepsArray = findNestedArray(node.getChildren().get(1), node);
        if (longDiv == null || stepsArray == null) {
            return null;
        }

        VerticalLayoutSpec steps = compileArray(stepsArray);
        LongDivisionHeader header = new LongDivisionHeader(
            flattenNodeText(childAt(longDiv, 0)),
            flattenNodeText(childAt(longDiv, 1)),
            flattenNodeText(childAt(longDiv, 2))
        );
        if (steps == null) {
            steps = compileExplicitLongDivision(longDiv);
        }
        steps = offsetLayout(steps, 2, Kind.LONG_DIVISION, header);
        return new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            steps == null ? 0 : steps.columnCount(),
            steps == null ? -1 : steps.anchorColumn(),
            -1,
            steps == null ? List.of() : steps.tabStops(),
            steps == null ? List.of() : steps.rows(),
            steps == null ? List.of() : steps.ruleSpans(),
            header,
            steps == null ? List.of() : steps.longDivisionSteps()
        );
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

        boolean multiplication = isMultiplicationArray(node);
        int maxCols = rawRows.stream().mapToInt(List::size).max().orElse(0);
        int targetCols = multiplication ? maxCols + 1 : maxCols;
        List<VerticalRow> rows = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < rawRows.size(); rowIndex++) {
            List<String> row = new ArrayList<>(rawRows.get(rowIndex));
            int padLeft = targetCols - row.size() - (multiplication ? 1 : 0);
            while (padLeft-- > 0) {
                row.add(0, "");
            }
            if (multiplication) {
                row.add("");
            }
            while (row.size() < targetCols) {
                row.add("");
            }
            rows.add(new VerticalRow(resolveArrayRowKind(rowIndex, rawRows.size()), row));
        }

        List<RuleSpan> ruleSpans = compileRuleSpans(rows, parsePartitionArray(node.getMetadata("rowLines"), rows.size() + 1));
        rows = withSegments(rows, ruleSpans);
        return new VerticalLayoutSpec(Kind.ARITHMETIC, targetCols, targetCols - 1 - (multiplication ? 1 : 0), -1,
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

    private VerticalLayoutSpec compileLegacyDivisionArray(LaTeXNode node) {
        List<List<String>> rawRows = extractRowTexts(node);
        if (rawRows.isEmpty() || rawRows.get(0).size() < 2) {
            return null;
        }

        String divisor = rawRows.get(0).isEmpty() ? "" : rawRows.get(0).get(0);
        String dividend = rawRows.get(0).size() > 1 ? rawRows.get(0).get(1) : "";
        List<VerticalRow> rows = new ArrayList<>();
        List<RuleSpan> ruleSpans = new ArrayList<>();
        int maxCols = rawRows.stream().skip(1).mapToInt(List::size).max().orElse(1);
        int columnCount = Math.max(maxCols, 1) + 2;
        for (int rowIndex = 1; rowIndex < rawRows.size(); rowIndex++) {
            List<String> row = new ArrayList<>(rawRows.get(rowIndex));
            while (row.size() < columnCount - 2) {
                row.add(0, "");
            }
            row.add(0, "");
            row.add(0, "");
            rows.add(new VerticalRow(RowKind.STEP, row));
        }
        ruleSpans.addAll(compileRuleSpans(rows, parsePartitionArray(node.getMetadata("rowLines"), rawRows.size() + 1)));
        rows = withSegments(rows, ruleSpans);
        return new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            columnCount,
            columnCount - 1,
            -1,
            buildTabStops(columnCount, -1),
            rows,
            ruleSpans,
            new LongDivisionHeader(divisor, "", dividend)
        );
    }

    private VerticalLayoutSpec offsetLayout(VerticalLayoutSpec spec, int offset, Kind kind, LongDivisionHeader header) {
        if (spec == null) {
            return null;
        }
        List<VerticalRow> shiftedRows = new ArrayList<>();
        for (VerticalRow row : spec.rows()) {
            List<String> cells = new ArrayList<>();
            for (int i = 0; i < offset; i++) {
                cells.add("");
            }
            cells.addAll(row.cells());
            shiftedRows.add(new VerticalRow(row.kind(), cells));
        }
        List<RuleSpan> spans = new ArrayList<>(spec.ruleSpans());
        shiftedRows = withSegments(shiftedRows, spans);
        List<LongDivisionStep> shiftedSteps = offsetLongDivisionSteps(spec.longDivisionSteps(), offset);
        return new VerticalLayoutSpec(
            kind,
            spec.columnCount() + offset,
            spec.anchorColumn() + offset,
            spec.decimalColumn() < 0 ? -1 : spec.decimalColumn() + offset,
            buildTabStops(spec.columnCount() + offset, -1),
            shiftedRows,
            spans,
            header,
            shiftedSteps
        );
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
                row.add(flattenNodeText(cell));
            }
            trimRight(row);
            rows.add(row);
        }
        return rows;
    }

    private boolean isMultiplicationArray(LaTeXNode node) {
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                String text = flattenNodeText(cell);
                if (text.contains("\\times") || text.contains("×")) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean isDecimalArray(LaTeXNode node) {
        for (LaTeXNode row : node.getChildren()) {
            for (LaTeXNode cell : row.getChildren()) {
                if (".".equals(flattenNodeText(cell))) {
                    return true;
                }
            }
        }
        return false;
    }

    private boolean isDivisionArray(LaTeXNode node) {
        return node != null
            && node.getType() == LaTeXNode.Type.ARRAY
            && node.getMetadata("columnSpec") != null
            && node.getMetadata("columnSpec").contains("|");
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

    private LaTeXNode findFirstDescendant(LaTeXNode node, LaTeXNode.Type type) {
        if (node == null) {
            return null;
        }
        if (node.getType() == type) {
            return node;
        }
        for (LaTeXNode child : node.getChildren()) {
            LaTeXNode found = findFirstDescendant(child, type);
            if (found != null) {
                return found;
            }
        }
        return null;
    }

    private LaTeXNode findNestedArray(LaTeXNode node, LaTeXNode exclude) {
        if (node == null) {
            return null;
        }
        if (node != exclude && node.getType() == LaTeXNode.Type.ARRAY) {
            return node;
        }
        for (LaTeXNode child : node.getChildren()) {
            LaTeXNode found = findNestedArray(child, exclude);
            if (found != null) {
                return found;
            }
        }
        return null;
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
        while (!row.isEmpty() && blankToEmpty(row.get(row.size() - 1)).isEmpty()) {
            row.remove(row.size() - 1);
        }
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
            return blankToEmpty(node.getValue()).replace("\\", "");
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flattenNodeText(child));
        }
        return builder.toString();
    }

    private String blankToEmpty(String text) {
        return text == null || text.isBlank() ? "" : text;
    }

}
