package com.lz.paperword.core.mathml;

/** Serializes structured long division, including its surrounding expression, to MathML 3. */
public final class LongDivisionMathMlWriter {

    private static final double COLUMN_WIDTH_EM = 0.58d;
    private static final double MENCLOSE_RIGHT_INSET_EM = 0.20d;

    public String write(MathIRNode node) {
        LongDivisionSpec spec = longDivisionSpec(node);
        StringBuilder xml = new StringBuilder(1024);
        xml.append("<math xmlns=\"http://www.w3.org/1998/Math/MathML\">");
        appendLongDivision(xml, spec);
        return xml.append("</math>").toString();
    }

    /**
     * Writes the complete parsed formula. Never select a single long-division child:
     * a division can have siblings, occur more than once, or be nested in a script.
     * The original IR is only read, so the independent MTEF path keeps its semantics.
     * Unsupported structures fail explicitly rather than losing their decorations.
     */
    public String writeExpression(MathIRNode node) {
        if (node == null) {
            throw new IllegalArgumentException("MathIR expression is required");
        }
        MathIRNode single = node;
        while ((single.getType() == MathIRNode.Type.MATH || single.getType() == MathIRNode.Type.SEQUENCE)
                && single.getChildren().size() == 1 && single.getMetadata("direction") == null) {
            single = single.child(0);
        }
        // Keep the established pure-division layout (and render cache bytes) unchanged.
        if (single.getType() == MathIRNode.Type.LONG_DIVISION) {
            return write(single);
        }
        StringBuilder xml = new StringBuilder(2048);
        xml.append("<math xmlns=\"http://www.w3.org/1998/Math/MathML\">");
        appendExpression(xml, LongDivisionSpec.Expression.from(node));
        return xml.append("</math>").toString();
    }

    private void appendLongDivision(StringBuilder xml, LongDivisionSpec spec) {
        xml.append("<semantics><mrow>");
        appendCompactVisual(xml, spec);
        xml.append("</mrow><annotation-xml encoding=\"application/mathml-presentation+xml\">");
        appendSemanticLongDivision(xml, spec);
        xml.append("</annotation-xml></semantics>");
    }

    private LongDivisionSpec longDivisionSpec(MathIRNode node) {
        if (node == null || node.getType() != MathIRNode.Type.LONG_DIVISION
                || node.getChildren().size() != 4 || node.child(3).getType() != MathIRNode.Type.TABLE) {
            throw new IllegalArgumentException("LONG_DIVISION requires three operands and a step table");
        }
        for (MathIRNode row : node.child(3).getChildren()) {
            if (row.getType() != MathIRNode.Type.TABLE_ROW
                    || row.getChildren().stream().anyMatch(cell -> cell.getType() != MathIRNode.Type.TABLE_CELL)) {
                throw new IllegalArgumentException("LONG_DIVISION requires structured step rows and cells");
            }
        }
        return LongDivisionSpec.from(node);
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
        if (expressions.stream().anyMatch(expression -> expression.value() == null
                || !expression.children().isEmpty()
                || (expression.type() != MathIRNode.Type.NUMBER && expression.type() != MathIRNode.Type.OPERATOR))) {
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
        if (expression.metadata().get("direction") == null && (expression.type() == MathIRNode.Type.MATH
                || expression.type() == MathIRNode.Type.SEQUENCE
                || expression.type() == MathIRNode.Type.TABLE_CELL)) {
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
                requireChildren(expressions.get(index), 0);
                number.append(expressions.get(index).value());
                index++;
            }
            token(xml, "mn", number.toString());
        }
    }

    private void appendExpression(StringBuilder xml, LongDivisionSpec.Expression expression) {
        switch (expression.type()) {
            case MATH, SEQUENCE, TABLE_ROW, TABLE_CELL -> container(xml, "mrow", expression);
            case IDENT, NUMBER, OPERATOR, TEXT -> expressionToken(xml, expression);
            case STYLE -> style(xml, expression);
            case FRACTION -> fraction(xml, expression);
            case SQRT -> container(xml, "msqrt", expression);
            case ROOT -> {
                requireChildren(expression, 2);
                xml.append("<mroot>");
                // Parser/IR order is degree, radicand; MathML order is radicand, degree.
                appendChild(xml, expression, 1);
                appendChild(xml, expression, 0);
                xml.append("</mroot>");
            }
            case SUB -> binary(xml, "msub", expression);
            case SUP -> binary(xml, "msup", expression);
            case SUBSUP -> ternary(xml, "msubsup", expression);
            case UNDER, OVER, ARC -> embellished(xml, expression);
            case UNDEROVER -> ternary(xml, "munderover", expression);
            case FENCE -> fenced(xml, expression);
            case TABLE -> table(xml, expression);
            case LONG_DIVISION -> appendLongDivision(xml, longDivisionSpec(expression.toMathIR()));
            case ENCLOSURE -> enclosure(xml, expression);
            case HBRACE, HBRACK -> horizontalFence(xml, expression);
            case ARROW -> arrow(xml, expression);
            case DIRAC -> dirac(xml, expression);
            case UNSUPPORTED -> throw unsupported(expression, "unsupported node");
        }
    }

    private void expressionToken(StringBuilder xml, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 0);
        String tag = switch (expression.type()) {
            case IDENT -> "mi";
            case NUMBER -> "mn";
            case TEXT -> "mtext";
            default -> "mo";
        };
        xml.append('<').append(tag);
        if ("function".equals(expression.metadata().get("role"))) {
            attribute(xml, "mathvariant", "normal");
        }
        if ("big-operator".equals(expression.metadata().get("role"))) {
            attribute(xml, "movablelimits", "false");
            if (expression.type() == MathIRNode.Type.OPERATOR) {
                attribute(xml, "largeop", "true");
            }
        }
        xml.append('>');
        escape(xml, expression.value());
        xml.append("</").append(tag).append('>');
    }

    private void fraction(StringBuilder xml, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 2);
        String style = expression.metadata().getOrDefault("fractionStyle", "auto");
        boolean explicit = "display".equals(style) || "continued".equals(style) || "text".equals(style);
        if (explicit) {
            xml.append("<mstyle displaystyle=\"").append(!"text".equals(style))
                .append("\" scriptlevel=\"0\">");
        } else if (!"auto".equals(style) && !"slash".equals(style)) {
            throw unsupported(expression, "fraction style " + style);
        }
        xml.append("<mfrac");
        if ("slash".equals(style)) {
            attribute(xml, "bevelled", "true");
        }
        xml.append('>');
        appendChild(xml, expression, 0);
        appendChild(xml, expression, 1);
        xml.append("</mfrac>");
        if (explicit) {
            xml.append("</mstyle>");
        }
    }

    private void style(StringBuilder xml, LongDivisionSpec.Expression expression) {
        String kind = expression.metadata().get("styleKind");
        if (kind == null) {
            throw unsupported(expression, "style without styleKind");
        }
        if ("vertical-shift".equals(kind)) {
            if (expression.metadata().containsKey("boxHeight") || expression.metadata().containsKey("boxDepth")) {
                throw unsupported(expression, "raisebox optional metrics");
            }
            xml.append("<mpadded");
            String shift = requiredMetadata(expression, "verticalShift");
            java.util.regex.Matcher dimension = java.util.regex.Pattern.compile(
                "^([+-]?(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+))\\s*(pt|px|em|ex)?$",
                java.util.regex.Pattern.CASE_INSENSITIVE).matcher(shift.trim());
            if (!dimension.matches()) {
                throw unsupported(expression, "raisebox dimension " + shift);
            }
            java.math.BigDecimal amount = new java.math.BigDecimal(dimension.group(1));
            String unit = dimension.group(2) == null ? "pt" : dimension.group(2).toLowerCase(java.util.Locale.ROOT);
            String magnitude = amount.abs().stripTrailingZeros().toPlainString() + unit;
            String offset = (amount.signum() < 0 ? "-" : "+") + magnitude;
            attribute(xml, "voffset", offset);
            // voffset moves ink only. Its height/depth must move with it or the
            // MathJax viewBox and the downstream vector frame can clip the ink.
            attribute(xml, "height", offset);
            attribute(xml, "depth", (amount.signum() < 0 ? "+" : "-") + magnitude);
            xml.append('>');
            appendCoalesced(xml, expression.children());
            xml.append("</mpadded>");
            return;
        }
        xml.append("<mstyle");
        switch (kind) {
            case "font" -> attribute(xml, "mathvariant", requiredMetadata(expression, "fontVariant"));
            case "size" -> attribute(xml, "mathsize", requiredMetadata(expression, "fontSizePt") + "pt");
            case "math-style" -> {
                String mathStyle = requiredMetadata(expression, "mathStyle");
                int level = switch (mathStyle) {
                    case "displaystyle", "textstyle" -> 0;
                    case "scriptstyle" -> 1;
                    case "scriptscriptstyle" -> 2;
                    default -> throw unsupported(expression, "math style " + mathStyle);
                };
                attribute(xml, "displaystyle", Boolean.toString("displaystyle".equals(mathStyle)));
                attribute(xml, "scriptlevel", Integer.toString(level));
            }
            case "color" -> attribute(xml, "mathcolor", rgbColor(expression));
            // Arbitrary CSS cannot be preserved by MathJax SVG; never silently strip it.
            default -> throw unsupported(expression, "style kind " + kind);
        }
        xml.append('>');
        appendCoalesced(xml, expression.children());
        xml.append("</mstyle>");
    }

    private String rgbColor(LongDivisionSpec.Expression expression) {
        if (!"rgb".equals(expression.metadata().get("colorModel"))) {
            throw unsupported(expression, "color model " + expression.metadata().get("colorModel"));
        }
        String[] components = requiredMetadata(expression, "colorValue").split(",");
        if (components.length != 3) {
            throw unsupported(expression, "invalid RGB color");
        }
        StringBuilder color = new StringBuilder("#");
        for (String component : components) {
            double value = Double.parseDouble(component.trim());
            if (!Double.isFinite(value) || value < 0 || value > 1) {
                throw unsupported(expression, "invalid RGB component");
            }
            color.append(String.format(java.util.Locale.ROOT, "%02x", Math.round(value * 255)));
        }
        return color.toString();
    }

    private void embellished(StringBuilder xml, LongDivisionSpec.Expression expression) {
        String accent = expression.metadata().get("accentCommand");
        boolean arc = expression.type() == MathIRNode.Type.ARC;
        String tag = expression.type() == MathIRNode.Type.UNDER ? "munder" : "mover";
        if (accent == null && !arc) {
            binary(xml, tag, expression);
            return;
        }
        requireChildren(expression, 1);
        String mark = arc ? "⌢" : switch (accent) {
            case "\\overline", "\\bar" -> "¯";
            case "\\underline" -> "_";
            case "\\hat" -> "^";
            case "\\tilde" -> "~";
            case "\\vec", "\\overrightarrow", "\\underrightarrow" -> "→";
            case "\\overleftarrow", "\\underleftarrow" -> "←";
            case "\\overleftrightarrow", "\\underleftrightarrow" -> "↔";
            case "\\dot" -> "˙";
            default -> throw unsupported(expression, "accent " + accent);
        };
        xml.append('<').append(tag);
        attribute(xml, "munder".equals(tag) ? "accentunder" : "accent", "true");
        xml.append('>');
        appendChild(xml, expression, 0);
        xml.append("<mo stretchy=\"true\">");
        escape(xml, mark);
        xml.append("</mo></").append(tag).append('>');
    }

    private void enclosure(StringBuilder xml, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 1);
        String notation = requiredMetadata(expression, "notation");
        if (!java.util.Set.of("box", "updiagonalstrike", "downdiagonalstrike",
                "updiagonalstrike downdiagonalstrike").contains(notation)) {
            throw unsupported(expression, "enclosure notation " + notation);
        }
        xml.append("<menclose");
        attribute(xml, "notation", notation);
        xml.append('>');
        appendChild(xml, expression, 0);
        xml.append("</menclose>");
    }

    private void horizontalFence(StringBuilder xml, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 2);
        boolean top = "top".equals(requiredMetadata(expression, "placement"));
        String tag = top ? "mover" : "munder";
        boolean annotation = !transparentChildren(expression.children().get(1)).isEmpty();
        if (annotation) {
            xml.append('<').append(tag).append('>');
        }
        xml.append('<').append(tag);
        attribute(xml, top ? "accent" : "accentunder", "true");
        xml.append('>');
        appendChild(xml, expression, 0);
        String mark = expression.type() == MathIRNode.Type.HBRACE ? (top ? "⏞" : "⏟") : (top ? "⎴" : "⎵");
        xml.append("<mo stretchy=\"true\">");
        escape(xml, mark);
        xml.append("</mo></").append(tag).append('>');
        if (annotation) {
            appendChild(xml, expression, 1);
            xml.append("</").append(tag).append('>');
        }
    }

    private void arrow(StringBuilder xml, LongDivisionSpec.Expression expression) {
        if (expression.children().size() < 1 || expression.children().size() > 2) {
            throw unsupported(expression, "invalid arrow operands");
        }
        String command = requiredMetadata(expression, "latexCommand");
        String symbol = switch (command) {
            case "\\xrightarrow", "\\xlongrightarrow" -> "→";
            case "\\xleftarrow", "\\xlongleftarrow" -> "←";
            case "\\xleftrightarrow", "\\xlongleftrightarrow" -> "↔";
            case "\\xLongrightarrow" -> "⇒";
            case "\\xLongleftarrow" -> "⇐";
            case "\\xLeftrightarrow", "\\xLongleftrightarrow" -> "⇔";
            case "\\xlongequal" -> "=";
            default -> throw unsupported(expression, "arrow " + command);
        };
        String tag = expression.children().size() == 2 ? "munderover" : "mover";
        xml.append('<').append(tag).append("><mo stretchy=\"true\" minsize=\"1.75em\">");
        escape(xml, symbol);
        xml.append("</mo>");
        if (expression.children().size() == 2) {
            appendChild(xml, expression, 1);
        }
        appendChild(xml, expression, 0);
        xml.append("</").append(tag).append('>');
    }

    private void dirac(StringBuilder xml, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 2);
        String command = requiredMetadata(expression, "latexCommand");
        xml.append("<mrow>");
        switch (command) {
            case "\\bra" -> {
                delimiter(xml, "⟨");
                appendChild(xml, expression, 0);
                delimiter(xml, "|");
            }
            case "\\ket" -> {
                delimiter(xml, "|");
                appendChild(xml, expression, 1);
                delimiter(xml, "⟩");
            }
            case "\\braket" -> {
                delimiter(xml, "⟨");
                appendChild(xml, expression, 0);
                delimiter(xml, expression.metadata().getOrDefault("middleDelimiter", "|"));
                appendChild(xml, expression, 1);
                delimiter(xml, "⟩");
            }
            default -> throw unsupported(expression, "Dirac command " + command);
        }
        xml.append("</mrow>");
    }

    private void requireChildren(LongDivisionSpec.Expression expression, int count) {
        if (expression.children().size() != count) {
            throw unsupported(expression, "expected " + count + " operands, got " + expression.children().size());
        }
    }

    private String requiredMetadata(LongDivisionSpec.Expression expression, String key) {
        String value = expression.metadata().get(key);
        if (value == null || value.isBlank()) {
            throw unsupported(expression, "missing " + key);
        }
        return value;
    }

    private IllegalArgumentException unsupported(LongDivisionSpec.Expression expression, String detail) {
        return new IllegalArgumentException("Unsupported MathIR in longdivision preview: "
            + expression.type() + " (" + detail + ")");
    }

    private void attribute(StringBuilder xml, String name, String value) {
        xml.append(' ').append(name).append("=\"");
        escape(xml, value);
        xml.append('"');
    }

    private void token(StringBuilder xml, String tag, String value) {
        xml.append('<').append(tag).append('>');
        escape(xml, value);
        xml.append("</").append(tag).append('>');
    }

    private void container(StringBuilder xml, String tag, LongDivisionSpec.Expression expression) {
        xml.append('<').append(tag);
        String direction = expression.metadata().get("direction");
        if (direction != null) {
            attribute(xml, "dir", direction);
        }
        xml.append('>');
        appendCoalesced(xml, expression.children());
        xml.append("</").append(tag).append('>');
    }

    private void binary(StringBuilder xml, String tag, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 2);
        xml.append('<').append(tag).append('>');
        appendChild(xml, expression, 0);
        appendChild(xml, expression, 1);
        xml.append("</").append(tag).append('>');
    }

    private void ternary(StringBuilder xml, String tag, LongDivisionSpec.Expression expression) {
        requireChildren(expression, 3);
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
        String binomialStyle = expression.metadata().get("binomialStyle");
        boolean explicitStyle = "dbinom".equals(binomialStyle) || "tbinom".equals(binomialStyle);
        if (binomialStyle != null && !"binom".equals(binomialStyle) && !explicitStyle) {
            throw unsupported(expression, "binomial style " + binomialStyle);
        }
        if (explicitStyle) {
            xml.append("<mstyle displaystyle=\"").append("dbinom".equals(binomialStyle))
                .append("\" scriptlevel=\"0\">");
        }
        xml.append("<mrow>");
        delimiter(xml, expression.metadata().getOrDefault("openDelimiter", "("));
        expression.children().forEach(child -> appendExpression(xml, child));
        delimiter(xml, expression.metadata().getOrDefault("closeDelimiter", ")"));
        xml.append("</mrow>");
        if (explicitStyle) {
            xml.append("</mstyle>");
        }
    }

    private void delimiter(StringBuilder xml, String value) {
        if (".".equals(value) || value.isEmpty()) {
            return;
        }
        String symbol = switch (value) {
            case "\\{" -> "{";
            case "\\}" -> "}";
            case "||", "\\|", "\\Vert", "\\lVert", "\\rVert" -> "‖";
            default -> value;
        };
        xml.append("<mo fence=\"true\" stretchy=\"true\">");
        escape(xml, symbol);
        xml.append("</mo>");
    }

    private void table(StringBuilder xml, LongDivisionSpec.Expression expression) {
        if ("true".equals(expression.metadata().get("binomialPile"))) {
            // A table forces textstyle and cannot distinguish dbinom/tbinom.
            // A rule-free fraction preserves inherited MathML style semantics.
            requireChildren(expression, 2);
            xml.append("<mfrac linethickness=\"0\">");
            for (LongDivisionSpec.Expression row : expression.children()) {
                requireChildren(row, 1);
                if (row.type() != MathIRNode.Type.TABLE_ROW
                        || row.children().get(0).type() != MathIRNode.Type.TABLE_CELL) {
                    throw unsupported(expression, "invalid binomial rows");
                }
                container(xml, "mrow", row.children().get(0));
            }
            xml.append("</mfrac>");
            return;
        }
        String open = expression.metadata().get("openDelimiter");
        String close = expression.metadata().get("closeDelimiter");
        if (open != null || close != null) {
            xml.append("<mrow>");
            delimiter(xml, open == null ? "" : open);
        }
        java.util.List<String> enclosure = new java.util.ArrayList<>();
        String[] columnLines = partitionLines(expression, "columnLines");
        String[] rowLines = partitionLines(expression, "rowLines");
        if (columnLines.length > 0) {
            if ("1".equals(columnLines[0])) enclosure.add("left");
            if ("1".equals(columnLines[columnLines.length - 1])) enclosure.add("right");
        }
        if (rowLines.length > 0) {
            if ("1".equals(rowLines[0])) enclosure.add("top");
            if ("1".equals(rowLines[rowLines.length - 1])) enclosure.add("bottom");
        }
        String rowLineStyle = expression.metadata().getOrDefault("rowLineStyle", "solid");
        if (!enclosure.isEmpty()) {
            if (!"solid".equals(rowLineStyle)) {
                throw unsupported(expression, "dashed outer table rules");
            }
            xml.append("<menclose");
            attribute(xml, "notation", String.join(" ", enclosure));
            xml.append('>');
        }
        xml.append("<mtable");
        String columnSpec = expression.metadata().get("columnSpec");
        if (columnSpec != null) {
            if (!columnSpec.matches("[lcr|]*")) {
                throw unsupported(expression, "table column specification " + columnSpec);
            }
            java.util.List<String> align = new java.util.ArrayList<>();
            for (char ch : columnSpec.toCharArray()) {
                switch (ch) {
                    case 'l' -> align.add("left");
                    case 'c' -> align.add("center");
                    case 'r' -> align.add("right");
                    default -> { }
                }
            }
            attribute(xml, "columnalign", String.join(" ", align));
        }
        appendPartitionLines(xml, "columnlines", columnLines, "solid");
        appendPartitionLines(xml, "rowlines", rowLines, rowLineStyle);
        xml.append('>');
        for (LongDivisionSpec.Expression row : expression.children()) {
            if (row.type() != MathIRNode.Type.TABLE_ROW || row.metadata().containsKey("clineBelow")) {
                throw unsupported(row, "table row or partial rule");
            }
            xml.append("<mtr>");
            for (LongDivisionSpec.Expression cell : row.children()) {
                if (cell.type() != MathIRNode.Type.TABLE_CELL) {
                    throw unsupported(cell, "table cell");
                }
                xml.append("<mtd>");
                appendCoalesced(xml, cell.children());
                xml.append("</mtd>");
            }
            xml.append("</mtr>");
        }
        xml.append("</mtable>");
        if (!enclosure.isEmpty()) {
            xml.append("</menclose>");
        }
        if (open != null || close != null) {
            delimiter(xml, close == null ? "" : close);
            xml.append("</mrow>");
        }
    }

    private String[] partitionLines(LongDivisionSpec.Expression expression, String name) {
        String value = expression.metadata().get(name);
        if (value == null) {
            return new String[0];
        }
        if (!value.matches("[01](,[01])*")) {
            throw unsupported(expression, "invalid table " + name);
        }
        return value.split(",");
    }

    private void appendPartitionLines(StringBuilder xml, String attribute, String[] parts, String style) {
        if (parts.length <= 2) {
            return;
        }
        java.util.List<String> lines = new java.util.ArrayList<>();
        for (int i = 1; i < parts.length - 1; i++) {
            lines.add("1".equals(parts[i]) ? style : "none");
        }
        attribute(xml, attribute, String.join(" ", lines));
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
