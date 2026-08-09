package com.lz.paperword.core.layout;

import com.lz.paperword.core.latex.LaTeXNode;

/** Deterministic MTEF layout-unit estimator used by structured vertical layouts. */
public final class MathStructureMeasurer {
    public static final int DEFAULT_COLUMN_WIDTH = 240;

    public int measure(LaTeXNode node) {
        if (node == null) {
            return 0;
        }
        return switch (node.getType()) {
            case CHAR -> measureText(node.getValue());
            case COMMAND -> Math.max(220, measureChildren(node));
            case FRACTION -> Math.max(measure(child(node, 0)), measure(child(node, 1))) + 96;
            case SQRT -> measureChildren(node) + 180;
            case SUPERSCRIPT, SUBSCRIPT -> measure(child(node, 0))
                + Math.max(120, (int) Math.round(measure(child(node, 1)) * 0.65d));
            case ARRAY -> measureArray(node);
            default -> measureChildren(node);
        };
    }

    private int measureArray(LaTeXNode node) {
        int maximum = 0;
        for (LaTeXNode row : node.getChildren()) {
            maximum = Math.max(maximum, measureChildren(row));
        }
        return maximum;
    }

    private int measureChildren(LaTeXNode node) {
        int width = 0;
        for (LaTeXNode child : node.getChildren()) {
            width = Math.addExact(width, measure(child));
        }
        return width;
    }

    private int measureText(String text) {
        if (text == null || text.isEmpty()) {
            return 0;
        }
        int width = 0;
        for (int codePoint : text.codePoints().toArray()) {
            width = Math.addExact(width, Character.isWhitespace(codePoint) ? 120 : DEFAULT_COLUMN_WIDTH);
        }
        return width;
    }

    private LaTeXNode child(LaTeXNode node, int index) {
        return index >= 0 && index < node.getChildren().size() ? node.getChildren().get(index) : null;
    }
}
