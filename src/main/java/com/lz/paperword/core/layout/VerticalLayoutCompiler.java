package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.layout.VerticalLayoutSpec.Kind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionHeader;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionPlacedLine;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionPlacedRowType;
import com.lz.paperword.core.layout.VerticalLayoutSpec.LongDivisionStep;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RowKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.RuleSpan;
import com.lz.paperword.core.layout.VerticalLayoutSpec.TabStopKind;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalRow;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalSegment;
import com.lz.paperword.core.layout.VerticalLayoutSpec.VerticalTabStop;

import java.math.BigInteger;
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

        String rawDivisor = flattenNodeText(childAt(node, 0));
        String rawQuotient = flattenNodeText(childAt(node, 1));
        String rawDividend = flattenNodeText(childAt(node, 2));
        LongDivisionHeader header = new LongDivisionHeader(rawDivisor, rawQuotient, rawDividend);
        int fullColumnCount = Math.max(rawDividend.length() + 2, 3);
        // 长除法步骤区的位置改由 LLM 原始结果提供，编译层只保留头部语义，不再本地补算步骤。
        return new VerticalLayoutSpec(
            Kind.LONG_DIVISION,
            fullColumnCount,
            Math.max(fullColumnCount - 1, 0),
            -1,
            buildTabStops(fullColumnCount, -1),
            List.of(),
            List.of(),
            header
        );
    }

    List<LongDivisionPlacedLine> computeExpectedLongDivisionLines(LongDivisionHeader header) {
        if (header == null) {
            return List.of();
        }

        String divisorText = header.divisor();
        String dividendText = header.dividend();
        boolean allowDecimalExpansion = false;

        NormalizedLongDivisionOperands normalizedOperands = normalizeLongDivisionOperands(divisorText, dividendText);
        if (normalizedOperands != null) {
            divisorText = normalizedOperands.divisor();
            dividendText = normalizedOperands.dividend();
            allowDecimalExpansion = normalizedOperands.allowDecimalExpansion();
        }

        return computeLongDivisionPlacedLines(divisorText, dividendText, allowDecimalExpansion);
    }

    /**
     * 先把有限小数统一放大成整数除法，再复用同一套本地长除法步骤生成逻辑。
     */
    private NormalizedLongDivisionOperands normalizeLongDivisionOperands(String divisorText, String dividendText) {
        ParsedDecimal divisor = parseFiniteDecimal(divisorText);
        ParsedDecimal dividend = parseFiniteDecimal(dividendText);
        if (divisor == null || dividend == null) {
            return null;
        }
        int targetScale = Math.max(divisor.scale(), dividend.scale());
        String normalizedDivisor = scaleDecimalDigits(divisor, targetScale);
        String normalizedDividend = scaleDecimalDigits(dividend, targetScale);
        if (!isAsciiDigits(normalizedDivisor) || !isAsciiDigits(normalizedDividend)) {
            return null;
        }
        boolean allowDecimalExpansion = divisor.scale() > 0 || dividend.scale() > 0;
        return new NormalizedLongDivisionOperands(
            normalizedDivisor,
            normalizedDividend,
            allowDecimalExpansion,
            allowDecimalExpansion
        );
    }

    /**
     * 纯整数长除法统一在本地补出商与步骤；若商为有限小数，则继续补小数部分。
     */
    private ComputedLongDivision computeLongDivision(String divisorText, String dividendText, boolean allowDecimalExpansion) {
        if (!isAsciiDigits(divisorText) || !isAsciiDigits(dividendText)) {
            return null;
        }
        BigInteger divisor = new BigInteger(divisorText);
        BigInteger dividend = new BigInteger(dividendText);
        if (BigInteger.ZERO.equals(divisor)) {
            return null;
        }
        if (!hasFiniteDecimalQuotient(dividend, divisor)) {
            return null;
        }

        int columnCount = Math.max(dividendText.length(), 1);
        BigInteger current = BigInteger.ZERO;
        String currentText = "";
        int currentEndColumn = -1;
        int position = 0;
        int maxColumnCount = columnCount;
        boolean quotientStarted = false;
        boolean decimalStarted = false;
        StringBuilder quotient = new StringBuilder();
        List<LongDivisionStep> steps = new ArrayList<>();

        while (position < dividendText.length() && current.compareTo(divisor) < 0) {
            current = appendDigit(current, dividendText.charAt(position));
            currentText += dividendText.charAt(position);
            currentEndColumn = position;
            position++;
        }

        if (current.compareTo(divisor) < 0 && BigInteger.ZERO.equals(current)) {
            return new ComputedLongDivision("0", dividendText.length(), List.of());
        }

        while (current.compareTo(divisor) >= 0
            || (allowDecimalExpansion && !BigInteger.ZERO.equals(current) && position >= dividendText.length())) {
            if (current.compareTo(divisor) < 0) {
                if (!decimalStarted) {
                    if (quotient.isEmpty()) {
                        quotient.append('0');
                    }
                    quotient.append('.');
                    decimalStarted = true;
                    quotientStarted = true;
                }
                do {
                    current = appendDigit(current, '0');
                    currentText += '0';
                    currentEndColumn++;
                    maxColumnCount = Math.max(maxColumnCount, currentEndColumn + 1);
                    if (quotientStarted && current.compareTo(divisor) < 0) {
                        quotient.append('0');
                    }
                } while (current.compareTo(divisor) < 0);
            }

            quotientStarted = true;
            BigInteger[] divRem = current.divideAndRemainder(divisor);
            quotient.append(divRem[0].toString());

            BigInteger product = divRem[0].multiply(divisor);
            VerticalRow productRow = buildPlacedRow(RowKind.STEP, product.toString(), currentEndColumn, maxColumnCount);

            current = divRem[1];
            currentText = BigInteger.ZERO.equals(current) ? "" : current.toString();
            int remainderEndColumn = currentEndColumn;

            while (position < dividendText.length() && current.compareTo(divisor) < 0) {
                current = appendDigit(current, dividendText.charAt(position));
                currentText += dividendText.charAt(position);
                remainderEndColumn = position;
                if (quotientStarted && current.compareTo(divisor) < 0) {
                    quotient.append('0');
                }
                position++;
            }

            // 整除且后面已经没有更多位可下拉时，显式补一个最终余数 0，
            // 这样整数/小数长除法都能统一走后续的“末尾 0 收敛”排版逻辑。
            if (currentText.isEmpty() && BigInteger.ZERO.equals(current)) {
                currentText = "0";
            }
            VerticalRow remainderRow = currentText.isEmpty()
                ? new VerticalRow(RowKind.REMAINDER, List.of())
                : buildPlacedRow(RowKind.REMAINDER, currentText, remainderEndColumn, maxColumnCount);
            steps.add(new LongDivisionStep(productRow, remainderRow));
            currentEndColumn = remainderEndColumn;
            maxColumnCount = Math.max(maxColumnCount, remainderEndColumn + 1);
        }

        if (quotient.isEmpty()) {
            quotient.append('0');
        }
        return new ComputedLongDivision(quotient.toString(), maxColumnCount, steps);
    }

    private List<LongDivisionPlacedLine> computeLongDivisionPlacedLines(String divisorText,
                                                                        String dividendText,
                                                                        boolean allowDecimalExpansion) {
        if (!isAsciiDigits(divisorText) || !isAsciiDigits(dividendText)) {
            return List.of();
        }

        BigInteger divisor = new BigInteger(divisorText);
        BigInteger dividend = new BigInteger(dividendText);
        if (BigInteger.ZERO.equals(divisor)) {
            return List.of();
        }
        if (allowDecimalExpansion && !hasFiniteDecimalQuotient(dividend, divisor)) {
            return List.of();
        }

        BigInteger current = BigInteger.ZERO;
        String currentText = "";
        int currentEndColumn = -1;
        int position = 0;
        List<LongDivisionPlacedLine> placedLines = new ArrayList<>();

        while (position < dividendText.length() && current.compareTo(divisor) < 0) {
            current = appendDigit(current, dividendText.charAt(position));
            currentText += dividendText.charAt(position);
            currentEndColumn = position;
            position++;
        }

        if (current.compareTo(divisor) < 0 && BigInteger.ZERO.equals(current)) {
            return List.of();
        }

        while (current.compareTo(divisor) >= 0
            || (allowDecimalExpansion && !BigInteger.ZERO.equals(current) && position >= dividendText.length())) {
            if (current.compareTo(divisor) < 0) {
                do {
                    current = appendDigit(current, '0');
                    currentText += '0';
                    currentEndColumn++;
                } while (current.compareTo(divisor) < 0);
            }

            int activeEndColumn = currentEndColumn;
            int activeStartColumn = Math.max(activeEndColumn - currentText.length() + 1, 0);

            BigInteger[] divRem = current.divideAndRemainder(divisor);
            String productText = divRem[0].multiply(divisor).toString();
            int productEndColumn = activeEndColumn;
            int productStartColumn = Math.max(productEndColumn - productText.length() + 1, activeStartColumn);
            placedLines.add(new LongDivisionPlacedLine(
                LongDivisionPlacedRowType.PRODUCT,
                productText,
                productStartColumn,
                productEndColumn,
                activeStartColumn,
                activeEndColumn
            ));

            current = divRem[1];
            currentText = BigInteger.ZERO.equals(current) ? "" : current.toString();
            int remainderEndColumn = activeEndColumn;

            while (position < dividendText.length() && current.compareTo(divisor) < 0) {
                current = appendDigit(current, dividendText.charAt(position));
                currentText += dividendText.charAt(position);
                remainderEndColumn = position;
                position++;
            }

            if (currentText.isEmpty() && BigInteger.ZERO.equals(current)) {
                currentText = "0";
            }
            if (!currentText.isEmpty()) {
                int remainderStartColumn = Math.max(remainderEndColumn - currentText.length() + 1, 0);
                placedLines.add(new LongDivisionPlacedLine(
                    LongDivisionPlacedRowType.REMAINDER,
                    currentText,
                    remainderStartColumn,
                    remainderEndColumn,
                    -1,
                    -1
                ));
                if (position >= dividendText.length()
                    && BigInteger.ZERO.equals(current)
                    && currentText.length() > 1
                    && isAllZeroDigits(currentText)) {
                    placedLines.add(new LongDivisionPlacedLine(
                        LongDivisionPlacedRowType.PRODUCT,
                        currentText,
                        remainderStartColumn,
                        remainderEndColumn,
                        remainderStartColumn,
                        remainderEndColumn
                    ));
                    placedLines.add(new LongDivisionPlacedLine(
                        LongDivisionPlacedRowType.REMAINDER,
                        "0",
                        remainderEndColumn,
                        remainderEndColumn,
                        -1,
                        -1
                    ));
                    break;
                }
            }
            currentEndColumn = remainderEndColumn;
        }

        return placedLines;
    }

    private boolean isAllZeroDigits(String text) {
        if (text == null || text.isEmpty()) {
            return false;
        }
        for (int index = 0; index < text.length(); index++) {
            if (text.charAt(index) != '0') {
                return false;
            }
        }
        return true;
    }

    private BigInteger appendDigit(BigInteger current, char digit) {
        return current.multiply(BigInteger.TEN).add(BigInteger.valueOf(digit - '0'));
    }

    private boolean isAsciiDigits(String value) {
        if (value == null || value.isBlank()) {
            return false;
        }
        for (int index = 0; index < value.length(); index++) {
            if (!Character.isDigit(value.charAt(index))) {
                return false;
            }
        }
        return true;
    }

    private boolean hasFiniteDecimalQuotient(BigInteger dividend, BigInteger divisor) {
        BigInteger gcd = dividend.gcd(divisor);
        BigInteger reducedDivisor = divisor.divide(gcd);
        reducedDivisor = stripFactor(reducedDivisor, 2);
        reducedDivisor = stripFactor(reducedDivisor, 5);
        return BigInteger.ONE.equals(reducedDivisor);
    }

    private BigInteger stripFactor(BigInteger value, int factor) {
        BigInteger factorValue = BigInteger.valueOf(factor);
        while (!BigInteger.ZERO.equals(value) && BigInteger.ZERO.equals(value.mod(factorValue))) {
            value = value.divide(factorValue);
        }
        return value;
    }

    private ParsedDecimal parseFiniteDecimal(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        String normalized = value.trim();
        if (normalized.chars().filter(ch -> ch == '.').count() > 1) {
            return null;
        }
        if (!normalized.chars().allMatch(ch -> Character.isDigit(ch) || ch == '.')) {
            return null;
        }
        int dotIndex = normalized.indexOf('.');
        if (dotIndex < 0) {
            return new ParsedDecimal(stripLeadingZerosKeepZero(normalized), 0);
        }
        String integerPart = normalized.substring(0, dotIndex);
        String fractionPart = normalized.substring(dotIndex + 1);
        if (integerPart.isEmpty()) {
            integerPart = "0";
        }
        String digits = stripLeadingZerosKeepZero(integerPart + fractionPart);
        return new ParsedDecimal(digits, fractionPart.length());
    }

    private String scaleDecimalDigits(ParsedDecimal value, int targetScale) {
        String digits = value == null ? "" : value.digits();
        int zerosToAppend = Math.max(targetScale - (value == null ? 0 : value.scale()), 0);
        return stripLeadingZerosKeepZero(digits + "0".repeat(zerosToAppend));
    }

    private String stripLeadingZerosKeepZero(String value) {
        if (value == null || value.isEmpty()) {
            return "0";
        }
        int index = 0;
        while (index < value.length() - 1 && value.charAt(index) == '0') {
            index++;
        }
        return value.substring(index);
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
    private record ParsedDecimal(String digits, int scale) {
    }

    private record NormalizedLongDivisionOperands(String divisor,
                                                  String dividend,
                                                  boolean allowDecimalExpansion,
                                                  boolean normalizedFromDecimal) {
    }

    private record ComputedLongDivision(String quotient, int columnCount, List<LongDivisionStep> steps) {
    }

    public record CrossMultiplicationLayout(List<List<String>> rows) {
        public CrossMultiplicationLayout {
            rows = rows == null ? List.of() : List.copyOf(rows.stream().map(List::copyOf).toList());
        }
    }

}
