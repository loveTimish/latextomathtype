package com.lz.paperword.core.layout;

import java.util.ArrayList;
import java.util.List;

/**
 * 统一的竖式布局语义模型。
 *
 * <p>预览层和 MTEF 本体都应从同一份布局数据生成，避免各自维护独立坐标逻辑。</p>
 */
public record VerticalLayoutSpec(
    Kind kind,
    int columnCount,
    int anchorColumn,
    int decimalColumn,
    List<VerticalTabStop> tabStops,
    List<VerticalRow> rows,
    List<RuleSpan> ruleSpans,
    LongDivisionHeader longDivisionHeader,
    List<LongDivisionStep> longDivisionSteps
) {
    public VerticalLayoutSpec {
        kind = kind == null ? Kind.ARITHMETIC : kind;
        columnCount = Math.max(columnCount, 0);
        anchorColumn = Math.max(anchorColumn, -1);
        decimalColumn = Math.max(decimalColumn, -1);
        tabStops = copyTabStops(tabStops);
        rows = copyRows(rows);
        ruleSpans = copyRuleSpans(ruleSpans);
        longDivisionSteps = copyLongDivisionSteps(longDivisionSteps);
    }

    public VerticalLayoutSpec(
        Kind kind,
        int columnCount,
        int anchorColumn,
        int decimalColumn,
        List<VerticalTabStop> tabStops,
        List<VerticalRow> rows,
        List<RuleSpan> ruleSpans,
        LongDivisionHeader longDivisionHeader
    ) {
        this(kind, columnCount, anchorColumn, decimalColumn, tabStops, rows, ruleSpans, longDivisionHeader, List.of());
    }

    public boolean isLongDivision() {
        return kind == Kind.LONG_DIVISION && longDivisionHeader != null;
    }

    public boolean hasStructuredLongDivisionSteps() {
        return isLongDivision() && !longDivisionSteps.isEmpty();
    }

    public enum Kind {
        ARITHMETIC,
        DECIMAL,
        LONG_DIVISION
    }

    public enum RowKind {
        OPERAND,
        RESULT,
        PARTIAL,
        STEP,
        REMAINDER,
        DIVIDEND,
        RULE
    }

    public enum TabStopKind {
        LEFT,
        CENTER,
        RIGHT,
        RELATION,
        DECIMAL
    }

    public record VerticalTabStop(int column, TabStopKind kind, int offset) {
        public VerticalTabStop {
            column = Math.max(column, 0);
            kind = kind == null ? TabStopKind.RIGHT : kind;
            offset = Math.max(offset, 0);
        }
    }

    public record VerticalSegment(String text, int startColumn, int endColumn, boolean underlined) {
        public VerticalSegment {
            text = text == null ? "" : text;
            startColumn = Math.max(startColumn, 0);
            endColumn = Math.max(endColumn, startColumn);
        }
    }

    public record VerticalRow(RowKind kind, List<String> cells, List<VerticalSegment> segments) {
        public VerticalRow {
            kind = kind == null ? RowKind.STEP : kind;
            cells = cells == null ? List.of() : List.copyOf(cells);
            segments = segments == null ? List.of() : List.copyOf(segments);
        }

        public VerticalRow(RowKind kind, List<String> cells) {
            this(kind, cells, List.of());
        }
    }

    public record RuleSpan(int boundaryIndex, int startColumn, int endColumn) {
        public RuleSpan {
            boundaryIndex = Math.max(boundaryIndex, 0);
            startColumn = Math.max(startColumn, 0);
            endColumn = Math.max(endColumn, startColumn);
        }
    }

    public record LongDivisionHeader(String divisor, String quotient, String dividend) {
        public LongDivisionHeader {
            divisor = divisor == null ? "" : divisor;
            quotient = quotient == null ? "" : quotient;
            dividend = dividend == null ? "" : dividend;
        }
    }

    public record LongDivisionStep(VerticalRow productRow, VerticalRow remainderRow) {
        public LongDivisionStep {
            productRow = productRow == null ? new VerticalRow(RowKind.STEP, List.of()) : productRow;
            remainderRow = remainderRow == null ? new VerticalRow(RowKind.REMAINDER, List.of()) : remainderRow;
        }
    }

    private static List<VerticalRow> copyRows(List<VerticalRow> source) {
        if (source == null || source.isEmpty()) {
            return List.of();
        }
        return List.copyOf(new ArrayList<>(source));
    }

    private static List<VerticalTabStop> copyTabStops(List<VerticalTabStop> source) {
        if (source == null || source.isEmpty()) {
            return List.of();
        }
        return List.copyOf(new ArrayList<>(source));
    }

    private static List<RuleSpan> copyRuleSpans(List<RuleSpan> source) {
        if (source == null || source.isEmpty()) {
            return List.of();
        }
        return List.copyOf(new ArrayList<>(source));
    }

    private static List<LongDivisionStep> copyLongDivisionSteps(List<LongDivisionStep> source) {
        if (source == null || source.isEmpty()) {
            return List.of();
        }
        return List.copyOf(new ArrayList<>(source));
    }
}
