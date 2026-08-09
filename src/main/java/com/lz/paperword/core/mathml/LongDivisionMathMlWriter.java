package com.lz.paperword.core.mathml;

/** Serializes the structured long-division IR to standard MathML 3. */
public final class LongDivisionMathMlWriter {

    private static final double COLUMN_WIDTH_EM = 0.58d;
    private static final double MENCLOSE_RIGHT_INSET_EM = 0.20d;

    public String write(MathIRNode node) {
        LongDivisionSpec spec = LongDivisionSpec.from(node);
        StringBuilder xml = new StringBuilder(1024);
        xml.append("<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><semantics><mrow>");
        appendCompactVisual(xml, spec);
        xml.append("</mrow><annotation-xml encoding=\"application/mathml-presentation+xml\">");
        appendSemanticLongDivision(xml, spec);
        return xml.append("</annotation-xml></semantics></math>").toString();
    }

    private void appendCompactVisual(StringBuilder xml, LongDivisionSpec spec) {
        xml.append("<mtable columnalign=\"right right\" columnspacing=\"0.08em\" rowspacing=\"0.12em\">");
        if (!transparentChildren(spec.quotient()).isEmpty()) {
            xml.append("<mtr><mtd><mrow></mrow></mtd><mtd columnalign=\"right\">");
            appendStackOperand(xml, spec.quotient());
            xml.append("</mtd></mtr>");
        }
        xml.append("<mtr><mtd columnalign=\"right\">");
        appendStackOperand(xml, spec.divisor());
        xml.append("</mtd><mtd columnalign=\"right\"><menclose notation=\"longdiv\">")
            .append("<mpadded lspace=\"").append(formatEm(MENCLOSE_RIGHT_INSET_EM))
            .append("\" width=\"+0em\">");
        appendStackOperand(xml, spec.dividend());
        xml.append("</mpadded></menclose></mtd></mtr>")
            .append("<mtr><mtd><mrow></mrow></mtd><mtd columnalign=\"right\">")
            .append("<mtable columnalign=\"right\" rowspacing=\"0.08em\">");
        for (LongDivisionSpec.Step step : spec.steps()) {
            xml.append("<mtr><mtd columnalign=\"right\"><mrow>");
            if (step.ruleBelow() != null) {
                int span = step.ruleBelow().endColumn() - step.ruleBelow().startColumn() + 1;
                xml.append("<menclose notation=\"bottom\"><mtable width=\"")
                    .append(formatEm(span * COLUMN_WIDTH_EM))
                    .append("\" columnalign=\"right\" rowspacing=\"0em\">")
                    .append("<mtr><mtd columnalign=\"right\">");
                appendRowContent(xml, step.content());
                appendSpace(xml, Math.max(0, step.ruleBelow().endColumn() - step.endColumn())
                    * COLUMN_WIDTH_EM, 0d);
                xml.append("</mtd></mtr></mtable></menclose>");
                appendSpace(xml, (spec.columnCount() - step.ruleBelow().endColumn()) * COLUMN_WIDTH_EM, 0d);
            } else {
                appendRowContent(xml, step.content());
                appendSpace(xml, (spec.columnCount() - step.endColumn()) * COLUMN_WIDTH_EM, 0d);
            }
            xml.append("</mrow></mtd></mtr>");
        }
        xml.append("</mtable></mtd></mtr></mtable>");
    }

    private void appendSpace(StringBuilder xml, double widthEm, double heightEm) {
        if (widthEm <= 0d) {
            return;
        }
        xml.append("<mspace width=\"").append(formatEm(widthEm)).append("\"");
        if (heightEm > 0d) {
            xml.append(" height=\"").append(formatEm(heightEm)).append("\"");
        }
        xml.append("></mspace>");
    }

    private String formatEm(double value) {
        return String.format(java.util.Locale.ROOT, "%.3fem", value);
    }

    private void appendSemanticLongDivision(StringBuilder xml, LongDivisionSpec spec) {
        xml.append("<mlongdiv longdivstyle=\"lefttop\">");
        appendStackOperand(xml, spec.divisor());
        appendStackHeader(xml, spec.quotient());
        appendStackHeader(xml, spec.dividend());
        xml.append("<msgroup position=\"0\" shift=\"0\">");
        for (LongDivisionSpec.Step step : spec.steps()) {
            int rowPosition = spec.columnCount() - step.endColumn();
            int groupPosition = step.ruleBelow() == null
                ? rowPosition
                : spec.columnCount() - step.ruleBelow().endColumn();
            int relativeRowPosition = rowPosition - groupPosition;
            xml.append("<msgroup position=\"").append(groupPosition).append("\"><msrow");
            if (relativeRowPosition != 0) {
                xml.append(" position=\"").append(relativeRowPosition).append("\"");
            }
            xml.append('>');
            appendRowContent(xml, step.content());
            xml.append("</msrow>");
            if (step.ruleBelow() != null) {
                int length = step.ruleBelow().endColumn() - step.ruleBelow().startColumn() + 1;
                xml.append("<msline length=\"").append(length).append("\"/>");
            }
            xml.append("</msgroup>");
        }
        xml.append("</msgroup></mlongdiv>");
    }

    /**
     * mstack only splits an mn into digit columns when it is a direct msrow child.
     * Parser grouping nodes carry no visual semantics here, so unwrap them while
     * preserving real structures such as fractions and roots as one logical cell.
     */
    private void appendRowContent(StringBuilder xml, LongDivisionSpec.Expression expression) {
        xml.append("<mrow>");
        appendCoalesced(xml, transparentChildren(expression));
        xml.append("</mrow>");
    }

    private void appendStackOperand(StringBuilder xml, LongDivisionSpec.Expression expression) {
        java.util.List<LongDivisionSpec.Expression> children = transparentChildren(expression);
        if (children.isEmpty()) {
            xml.append("<mrow/>");
        } else if (children.size() == 1) {
            appendExpression(xml, children.get(0));
        } else if (numericLiteral(children) != null) {
            token(xml, "mn", numericLiteral(children));
        } else {
            xml.append("<mrow>");
            appendCoalesced(xml, children);
            xml.append("</mrow>");
        }
    }

    private void appendStackHeader(StringBuilder xml, LongDivisionSpec.Expression expression) {
        java.util.List<LongDivisionSpec.Expression> children = transparentChildren(expression);
        if (children.isEmpty()) {
            xml.append("<mrow/>");
        } else if (numericLiteral(children) != null) {
            token(xml, "mn", numericLiteral(children));
        } else {
            xml.append("<msrow>");
            appendCoalesced(xml, children);
            xml.append("</msrow>");
        }
    }

    private String numericLiteral(java.util.List<LongDivisionSpec.Expression> expressions) {
        if (expressions.stream().anyMatch(expression -> expression.value() == null)) {
            return null;
        }
        String value = expressions.stream().map(LongDivisionSpec.Expression::value)
            .collect(java.util.stream.Collectors.joining());
        return value.matches("[+-]?[0-9]+(?:[.,][0-9]+)?") ? value : null;
    }

    private java.util.List<LongDivisionSpec.Expression> transparentChildren(
            LongDivisionSpec.Expression expression) {
        java.util.List<LongDivisionSpec.Expression> result = new java.util.ArrayList<>();
        collectTransparent(expression, result);
        return result;
    }

    private void collectTransparent(LongDivisionSpec.Expression expression,
                                    java.util.List<LongDivisionSpec.Expression> output) {
        if (expression.type() == MathIRNode.Type.MATH
                || expression.type() == MathIRNode.Type.SEQUENCE
                || expression.type() == MathIRNode.Type.TABLE_CELL) {
            expression.children().forEach(child -> collectTransparent(child, output));
        } else {
            output.add(expression);
        }
    }

    private void appendCoalesced(StringBuilder xml,
                                 java.util.List<LongDivisionSpec.Expression> expressions) {
        for (int index = 0; index < expressions.size();) {
            LongDivisionSpec.Expression expression = expressions.get(index);
            if (expression.type() != MathIRNode.Type.NUMBER) {
                appendExpression(xml, expression);
                index++;
                continue;
            }
            StringBuilder number = new StringBuilder();
            while (index < expressions.size()
                    && expressions.get(index).type() == MathIRNode.Type.NUMBER) {
                number.append(expressions.get(index).value());
                index++;
            }
            token(xml, "mn", number.toString());
        }
    }

    private void appendExpression(StringBuilder xml, LongDivisionSpec.Expression expression) {
        switch (expression.type()) {
            case MATH, SEQUENCE, TABLE_ROW, TABLE_CELL -> container(xml, "mrow", expression);
            case IDENT -> token(xml, "mi", expression.value());
            case NUMBER -> token(xml, "mn", expression.value());
            case OPERATOR -> token(xml, "mo", expression.value());
            case TEXT -> token(xml, "mtext", expression.value());
            case STYLE -> container(xml, "mstyle", expression);
            case FRACTION -> binary(xml, "mfrac", expression);
            case SQRT -> container(xml, "msqrt", expression);
            case ROOT -> binary(xml, "mroot", expression);
            case SUB -> binary(xml, "msub", expression);
            case SUP -> binary(xml, "msup", expression);
            case SUBSUP -> ternary(xml, "msubsup", expression);
            case UNDER -> binary(xml, "munder", expression);
            case OVER, ARC -> binary(xml, "mover", expression);
            case UNDEROVER -> ternary(xml, "munderover", expression);
            case FENCE -> fenced(xml, expression);
            case TABLE -> table(xml, expression);
            case LONG_DIVISION -> throw new IllegalArgumentException("nested longdivision is not supported");
            case UNSUPPORTED -> throw new IllegalArgumentException("unsupported MathIR in longdivision: " + expression.value());
            default -> container(xml, "mrow", expression);
        }
    }

    private void token(StringBuilder xml, String tag, String value) {
        xml.append('<').append(tag).append('>');
        escape(xml, value);
        xml.append("</").append(tag).append('>');
    }

    private void container(StringBuilder xml, String tag, LongDivisionSpec.Expression expression) {
        xml.append('<').append(tag).append('>');
        appendCoalesced(xml, expression.children());
        xml.append("</").append(tag).append('>');
    }

    private void binary(StringBuilder xml, String tag, LongDivisionSpec.Expression expression) {
        xml.append('<').append(tag).append('>');
        appendChild(xml, expression, 0);
        appendChild(xml, expression, 1);
        xml.append("</").append(tag).append('>');
    }

    private void ternary(StringBuilder xml, String tag, LongDivisionSpec.Expression expression) {
        xml.append('<').append(tag).append('>');
        appendChild(xml, expression, 0);
        appendChild(xml, expression, 1);
        appendChild(xml, expression, 2);
        xml.append("</").append(tag).append('>');
    }

    private void appendChild(StringBuilder xml, LongDivisionSpec.Expression expression, int index) {
        if (index < expression.children().size()) {
            appendExpression(xml, expression.children().get(index));
        } else {
            xml.append("<mrow/>");
        }
    }

    private void fenced(StringBuilder xml, LongDivisionSpec.Expression expression) {
        xml.append("<mrow><mo fence=\"true\">");
        escape(xml, expression.metadata().getOrDefault("openDelimiter", "("));
        xml.append("</mo>");
        expression.children().forEach(child -> appendExpression(xml, child));
        xml.append("<mo fence=\"true\">");
        escape(xml, expression.metadata().getOrDefault("closeDelimiter", ")"));
        xml.append("</mo></mrow>");
    }

    private void table(StringBuilder xml, LongDivisionSpec.Expression expression) {
        xml.append("<mtable>");
        for (LongDivisionSpec.Expression row : expression.children()) {
            xml.append("<mtr>");
            for (LongDivisionSpec.Expression cell : row.children()) {
                xml.append("<mtd>");
                cell.children().forEach(child -> appendExpression(xml, child));
                xml.append("</mtd>");
            }
            xml.append("</mtr>");
        }
        xml.append("</mtable>");
    }

    private void escape(StringBuilder xml, String value) {
        if (value == null) {
            return;
        }
        value.codePoints().forEach(codePoint -> {
            switch (codePoint) {
                case '&' -> xml.append("&amp;");
                case '<' -> xml.append("&lt;");
                case '>' -> xml.append("&gt;");
                case '"' -> xml.append("&quot;");
                default -> xml.appendCodePoint(codePoint);
            }
        });
    }
}
