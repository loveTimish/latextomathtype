package com.lz.paperword.core.mathml;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Immutable semantic model for an explicit {@code longdivision} environment. */
public record LongDivisionSpec(
    String columnSpec,
    int columnCount,
    Expression divisor,
    Expression quotient,
    Expression dividend,
    List<Step> steps
) {
    public LongDivisionSpec {
        columnSpec = columnSpec == null ? "" : columnSpec;
        if (!columnSpec.matches("r+") || columnCount != columnSpec.length()) {
            throw new IllegalArgumentException("longdivision column specification must match r+");
        }
        divisor = Expression.require(divisor, "divisor");
        quotient = Expression.require(quotient, "quotient");
        dividend = Expression.require(dividend, "dividend");
        steps = List.copyOf(steps == null ? List.of() : steps);
        if (steps.isEmpty()) {
            throw new IllegalArgumentException("longdivision requires at least one explicit step");
        }
        for (Step step : steps) {
            if (step.endColumn() < 1 || step.endColumn() > columnCount) {
                throw new IllegalArgumentException("longdivision step end column is out of range");
            }
            if (step.ruleBelow() != null
                    && (step.ruleBelow().startColumn() < 1
                    || step.ruleBelow().endColumn() > columnCount
                    || step.ruleBelow().startColumn() > step.ruleBelow().endColumn())) {
                throw new IllegalArgumentException("longdivision cline range is out of range");
            }
        }
    }

    public static LongDivisionSpec from(MathIRNode node) {
        if (node == null || node.getType() != MathIRNode.Type.LONG_DIVISION) {
            throw new IllegalArgumentException("LONG_DIVISION MathIR node required");
        }
        String spec = value(node.getMetadata("columnSpec"));
        int columns = parseInt(node.getMetadata("columnCount"), spec.length());
        MathIRNode table = node.child(3);
        List<Step> rows = new ArrayList<>();
        if (table != null) {
            for (MathIRNode row : table.getChildren()) {
                MathIRNode content = null;
                for (MathIRNode cell : row.getChildren()) {
                    if (!cell.getChildren().isEmpty()) {
                        if (content != null) {
                            throw new IllegalArgumentException("longdivision row has multiple non-empty cells");
                        }
                        content = cell;
                    }
                }
                if (content == null) {
                    throw new IllegalArgumentException("longdivision row has no content");
                }
                int endColumn = parseInt(row.getMetadata("endColumn"), -1);
                RuleSpan rule = null;
                if (row.getMetadata("ruleStartColumn") != null) {
                    rule = new RuleSpan(
                        parseInt(row.getMetadata("ruleStartColumn"), -1),
                        parseInt(row.getMetadata("ruleEndColumn"), -1));
                }
                rows.add(new Step(endColumn, Expression.from(content), rule));
            }
        }
        return new LongDivisionSpec(spec, columns,
            Expression.from(node.child(0)), Expression.from(node.child(1)),
            Expression.from(node.child(2)), rows);
    }

    private static int parseInt(String value, int fallback) {
        try {
            return Integer.parseInt(value);
        } catch (RuntimeException ignored) {
            return fallback;
        }
    }

    private static String value(String value) {
        return value == null ? "" : value;
    }

    public record Step(int endColumn, Expression content, RuleSpan ruleBelow) {
        public Step {
            content = Expression.require(content, "step content");
        }
    }

    public record RuleSpan(int startColumn, int endColumn) {
    }

    /** Deep immutable snapshot so parser or writer mutation cannot alter the specification. */
    public record Expression(
        MathIRNode.Type type,
        String value,
        Map<String, String> metadata,
        List<Expression> children
    ) {
        public Expression {
            if (type == null) {
                throw new IllegalArgumentException("expression type is required");
            }
            metadata = Map.copyOf(metadata == null ? Map.of() : new LinkedHashMap<>(metadata));
            children = List.copyOf(children == null ? List.of() : children);
        }

        static Expression require(Expression expression, String name) {
            if (expression == null) {
                throw new IllegalArgumentException(name + " is required");
            }
            return expression;
        }

        public static Expression from(MathIRNode node) {
            MathIRNode safe = node == null ? new MathIRNode(MathIRNode.Type.SEQUENCE) : node;
            List<Expression> children = safe.getChildren().stream().map(Expression::from).toList();
            return new Expression(safe.getType(), safe.getValue(), safe.getMetadata(), children);
        }

        public MathIRNode toMathIR() {
            MathIRNode node = new MathIRNode(type, value);
            metadata.forEach(node::setMetadata);
            children.stream().map(Expression::toMathIR).forEach(node::addChild);
            return node;
        }
    }
}
