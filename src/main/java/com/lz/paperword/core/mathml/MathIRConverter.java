package com.lz.paperword.core.mathml;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.mtef.MtefCharMap;
import com.lz.paperword.core.mtef.MtefRecord;

import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Converts the current LaTeX AST into a MathML-aligned IR.
 *
 * <p>The design goal is minimum invasion: keep the parser and writer mostly intact, but place a
 * semantic normalization step between them so MTEF emission no longer depends directly on every
 * ad-hoc AST shape.</p>
 */
public class MathIRConverter {

    private static final Set<String> FUNCTION_COMMANDS = Set.of(
        "\\sin", "\\cos", "\\tan", "\\cot", "\\sec", "\\csc",
        "\\arcsin", "\\arccos", "\\arctan",
        "\\sinh", "\\cosh", "\\tanh",
        "\\log", "\\ln", "\\exp", "\\lim", "\\max", "\\min",
        "\\det", "\\dim", "\\gcd"
    );

    private static final Set<String> UNDER_OVER_BIGOPS = Set.of(
        "\\sum", "\\prod", "\\coprod", "\\bigcup", "\\bigcap", "\\bigvee", "\\bigwedge", "\\lim"
    );

    private static final Set<String> SCRIPT_BIGOPS = Set.of(
        "\\int", "\\iint", "\\iiint", "\\oint"
    );

    private static final Set<String> OVER_ACCENTS = Set.of(
        "\\overline", "\\bar", "\\hat", "\\tilde", "\\vec", "\\dot"
    );

    private static final Set<String> UNDER_ACCENTS = Set.of(
        "\\underline"
    );

    private static final Set<String> HORIZONTAL_BRACES = Set.of(
        "\\overbrace", "\\underbrace"
    );

    private static final Set<String> HORIZONTAL_BRACKETS = Set.of(
        "\\overbracket", "\\underbracket"
    );

    private static final Set<String> SPACING_COMMANDS = Set.of(
        "\\,", "\\;", "\\:", "\\!", "\\quad", "\\qquad", "\\hspace", "\\hskip"
    );

    public MathIRNode convert(LaTeXNode root) {
        MathIRNode math = new MathIRNode(MathIRNode.Type.MATH);
        if (root == null) {
            return math;
        }
        copyMetadata(root, math);
        appendConvertedChildren(root.getChildren(), math);
        return math;
    }

    public String dump(MathIRNode root) {
        StringBuilder builder = new StringBuilder();
        appendDump(root, builder, 0);
        return builder.toString();
    }

    private void appendDump(MathIRNode node, StringBuilder builder, int depth) {
        if (node == null) {
            return;
        }
        builder.append("  ".repeat(Math.max(depth, 0))).append(node.getType());
        if (node.getValue() != null && !node.getValue().isEmpty()) {
            builder.append('(').append(node.getValue()).append(')');
        }
        if (!node.getMetadata().isEmpty()) {
            builder.append(' ').append(node.getMetadata());
        }
        builder.append('\n');
        for (MathIRNode child : node.getChildren()) {
            appendDump(child, builder, depth + 1);
        }
    }

    private void appendConvertedChildren(List<LaTeXNode> sourceChildren, MathIRNode parent) {
        for (LaTeXNode child : sourceChildren) {
            MathIRNode converted = convertNode(child);
            if (converted != null) {
                parent.addChild(converted);
            }
        }
    }

    private MathIRNode convertNode(LaTeXNode node) {
        if (node == null) {
            return null;
        }
        return switch (node.getType()) {
            case ROOT -> convert(node);
            case GROUP -> convertLinearContainer(node, MathIRNode.Type.SEQUENCE);
            case CHAR -> convertCharNode(node);
            case COMMAND -> convertCommandNode(node);
            case FRACTION -> convertFractionNode(node);
            case SQRT -> convertSqrtNode(node);
            case SUPERSCRIPT -> convertSuperscriptNode(node);
            case SUBSCRIPT -> convertSubscriptNode(node);
            case TEXT -> convertTextNode(node);
            case ARRAY -> convertTableNode(node);
            case ROW -> convertLinearContainer(node, MathIRNode.Type.TABLE_ROW);
            case CELL -> convertLinearContainer(node, MathIRNode.Type.TABLE_CELL);
            case LONG_DIVISION -> convertLongDivisionNode(node);
        };
    }

    private MathIRNode convertLinearContainer(LaTeXNode node, MathIRNode.Type targetType) {
        MathIRNode container = new MathIRNode(targetType);
        copyMetadata(node, container);
        appendConvertedChildren(node.getChildren(), container);
        return container;
    }

    private MathIRNode convertCharNode(LaTeXNode node) {
        String value = node.getValue() == null ? "" : node.getValue();
        MathIRNode.Type type = classifyLiteralToken(value);
        MathIRNode ir = new MathIRNode(type, value);
        copyMetadata(node, ir);
        return ir;
    }

    private MathIRNode convertCommandNode(LaTeXNode node) {
        String command = node.getValue();
        if (command == null || command.isBlank()) {
            return unsupported("missing-command", node, List.of());
        }
        if (SPACING_COMMANDS.contains(command)) {
            return null;
        }
        if (command.startsWith("\\left")) {
            return convertFenceNode(node);
        }
        if (isHorizontalFenceCommand(command)) {
            return convertHorizontalFenceNode(node, null);
        }
        if (isEnclosureCommand(command)) {
            return convertEnclosureNode(node, command);
        }
        if (OVER_ACCENTS.contains(command)) {
            MathIRNode over = new MathIRNode(MathIRNode.Type.OVER);
            copyMetadata(node, over);
            over.setMetadata("accentCommand", command);
            over.addChild(convertArgument(childAt(node, 0)));
            return over;
        }
        if (UNDER_ACCENTS.contains(command)) {
            MathIRNode under = new MathIRNode(MathIRNode.Type.UNDER);
            copyMetadata(node, under);
            under.setMetadata("accentCommand", command);
            under.addChild(convertArgument(childAt(node, 0)));
            return under;
        }
        if (FUNCTION_COMMANDS.contains(command)) {
            MathIRNode function = new MathIRNode(MathIRNode.Type.IDENT, command.substring(1));
            copyMetadata(node, function);
            function.setMetadata("latexCommand", command);
            function.setMetadata("role", isBigOperatorCommand(command) ? "big-operator" : "function");
            if (isBigOperatorCommand(command)) {
                function.setMetadata("limitPlacement", limitPlacement(command));
            }
            return function;
        }

        MtefCharMap.CharEntry entry = MtefCharMap.lookup(command);
        if (entry != null) {
            MathIRNode mapped = new MathIRNode(classifyMappedCommand(command, entry), displayValue(command, entry));
            copyMetadata(node, mapped);
            mapped.setMetadata("latexCommand", command);
            if (isBigOperatorCommand(command)) {
                mapped.setMetadata("role", "big-operator");
                mapped.setMetadata("limitPlacement", limitPlacement(command));
            }
            return mapped;
        }

        return unsupported(command, node, node.getChildren());
    }

    private MathIRNode convertFenceNode(LaTeXNode node) {
        MathIRNode fence = new MathIRNode(MathIRNode.Type.FENCE);
        copyMetadata(node, fence);
        String left = firstNonBlank(node.getMetadata("leftDelimiter"), stripLeftCommand(node.getValue()));
        String right = firstNonBlank(node.getMetadata("rightDelimiter"), matchingRightDelimiter(left));
        fence.setMetadata("openDelimiter", left);
        fence.setMetadata("closeDelimiter", right);
        fence.setMetadata("latexCommand", node.getValue());
        fence.addChild(convertArgument(childAt(node, 0)));
        return fence;
    }

    private MathIRNode convertEnclosureNode(LaTeXNode node, String command) {
        MathIRNode enclosure = new MathIRNode(MathIRNode.Type.ENCLOSURE);
        copyMetadata(node, enclosure);
        enclosure.setMetadata("latexCommand", command);
        enclosure.setMetadata("notation", enclosureNotation(command));
        enclosure.addChild(convertArgument(childAt(node, 0)));
        return enclosure;
    }

    private MathIRNode convertHorizontalFenceNode(LaTeXNode node, LaTeXNode annotation) {
        String command = node == null ? null : node.getValue();
        MathIRNode.Type type = isHorizontalBracketCommand(command) ? MathIRNode.Type.HBRACK : MathIRNode.Type.HBRACE;
        MathIRNode fence = new MathIRNode(type);
        copyMetadata(node, fence);
        fence.setMetadata("latexCommand", command);
        fence.setMetadata("placement", isTopHorizontalFence(command) ? "top" : "bottom");
        fence.addChild(convertArgument(childAt(node, 0)));
        fence.addChild(convertArgument(annotation));
        return fence;
    }

    private MathIRNode convertFractionNode(LaTeXNode node) {
        MathIRNode fraction = new MathIRNode(MathIRNode.Type.FRACTION);
        copyMetadata(node, fraction);
        fraction.addChild(convertArgument(childAt(node, 0)));
        fraction.addChild(convertArgument(childAt(node, 1)));
        return fraction;
    }

    private MathIRNode convertSqrtNode(LaTeXNode node) {
        MathIRNode result = new MathIRNode(node.getChildren().size() == 2 ? MathIRNode.Type.ROOT : MathIRNode.Type.SQRT);
        copyMetadata(node, result);
        if (node.getChildren().size() == 2) {
            result.addChild(convertArgument(childAt(node, 0)));
            result.addChild(convertArgument(childAt(node, 1)));
        } else {
            result.addChild(convertArgument(childAt(node, 0)));
        }
        return result;
    }

    private MathIRNode convertSuperscriptNode(LaTeXNode node) {
        LaTeXNode baseAst = childAt(node, 0);
        LaTeXNode supAst = childAt(node, 1);

        if (isHorizontalFenceScript(baseAst, true)) {
            return convertHorizontalFenceNode(baseAst, supAst);
        }

        if (baseAst != null && baseAst.getType() == LaTeXNode.Type.SUBSCRIPT) {
            LaTeXNode innerBase = childAt(baseAst, 0);
            LaTeXNode subAst = childAt(baseAst, 1);
            if (isUnderOverOperator(innerBase)) {
                MathIRNode underover = new MathIRNode(MathIRNode.Type.UNDEROVER);
                copyMetadata(node, underover);
                underover.addChild(convertArgument(innerBase));
                underover.addChild(convertArgument(subAst));
                underover.addChild(convertArgument(supAst));
                return underover;
            }
            MathIRNode subsup = new MathIRNode(MathIRNode.Type.SUBSUP);
            copyMetadata(node, subsup);
            subsup.addChild(convertArgument(innerBase));
            subsup.addChild(convertArgument(subAst));
            subsup.addChild(convertArgument(supAst));
            return subsup;
        }

        if (isUnderOverOperator(baseAst)) {
            MathIRNode over = new MathIRNode(MathIRNode.Type.OVER);
            copyMetadata(node, over);
            over.addChild(convertArgument(baseAst));
            over.addChild(convertArgument(supAst));
            return over;
        }

        MathIRNode sup = new MathIRNode(MathIRNode.Type.SUP);
        copyMetadata(node, sup);
        sup.addChild(convertArgument(baseAst));
        sup.addChild(convertArgument(supAst));
        return sup;
    }

    private MathIRNode convertSubscriptNode(LaTeXNode node) {
        LaTeXNode baseAst = childAt(node, 0);
        LaTeXNode subAst = childAt(node, 1);

        if (isHorizontalFenceScript(baseAst, false)) {
            return convertHorizontalFenceNode(baseAst, subAst);
        }

        if (baseAst != null && baseAst.getType() == LaTeXNode.Type.SUPERSCRIPT) {
            LaTeXNode innerBase = childAt(baseAst, 0);
            LaTeXNode supAst = childAt(baseAst, 1);
            if (isUnderOverOperator(innerBase)) {
                MathIRNode underover = new MathIRNode(MathIRNode.Type.UNDEROVER);
                copyMetadata(node, underover);
                underover.addChild(convertArgument(innerBase));
                underover.addChild(convertArgument(subAst));
                underover.addChild(convertArgument(supAst));
                return underover;
            }
            MathIRNode subsup = new MathIRNode(MathIRNode.Type.SUBSUP);
            copyMetadata(node, subsup);
            subsup.addChild(convertArgument(innerBase));
            subsup.addChild(convertArgument(subAst));
            subsup.addChild(convertArgument(supAst));
            return subsup;
        }

        if (isUnderOverOperator(baseAst)) {
            MathIRNode under = new MathIRNode(MathIRNode.Type.UNDER);
            copyMetadata(node, under);
            under.addChild(convertArgument(baseAst));
            under.addChild(convertArgument(subAst));
            return under;
        }

        MathIRNode sub = new MathIRNode(MathIRNode.Type.SUB);
        copyMetadata(node, sub);
        sub.addChild(convertArgument(baseAst));
        sub.addChild(convertArgument(subAst));
        return sub;
    }

    private MathIRNode convertTextNode(LaTeXNode node) {
        MathIRNode text = new MathIRNode(MathIRNode.Type.TEXT, flattenText(node));
        copyMetadata(node, text);
        return text;
    }

    private MathIRNode convertTableNode(LaTeXNode node) {
        MathIRNode table = new MathIRNode(MathIRNode.Type.TABLE);
        copyMetadata(node, table);
        if ("cases".equals(node.getMetadata("environment"))) {
            table.setMetadata("openDelimiter", "{");
            table.setMetadata("closeDelimiter", ".");
        }
        appendConvertedChildren(node.getChildren(), table);
        return table;
    }

    private MathIRNode convertLongDivisionNode(LaTeXNode node) {
        MathIRNode longDivision = new MathIRNode(MathIRNode.Type.LONG_DIVISION);
        copyMetadata(node, longDivision);
        longDivision.addChild(convertArgument(childAt(node, 0)));
        longDivision.addChild(convertArgument(childAt(node, 1)));
        longDivision.addChild(convertArgument(childAt(node, 2)));
        return longDivision;
    }

    private MathIRNode convertArgument(LaTeXNode node) {
        if (node == null) {
            return new MathIRNode(MathIRNode.Type.SEQUENCE);
        }
        if (node.getType() == LaTeXNode.Type.GROUP) {
            return convertLinearContainer(node, MathIRNode.Type.SEQUENCE);
        }
        if (node.getType() == LaTeXNode.Type.ROOT) {
            MathIRNode sequence = new MathIRNode(MathIRNode.Type.SEQUENCE);
            copyMetadata(node, sequence);
            appendConvertedChildren(node.getChildren(), sequence);
            return sequence;
        }
        MathIRNode converted = convertNode(node);
        return converted != null ? converted : new MathIRNode(MathIRNode.Type.SEQUENCE);
    }

    private MathIRNode unsupported(String value, LaTeXNode source, List<LaTeXNode> originalChildren) {
        MathIRNode unsupported = new MathIRNode(MathIRNode.Type.UNSUPPORTED, value);
        copyMetadata(source, unsupported);
        unsupported.setMetadata("latex", value);
        for (LaTeXNode child : originalChildren) {
            MathIRNode converted = convertNode(child);
            if (converted != null) {
                unsupported.addChild(converted);
            }
        }
        return unsupported;
    }

    private MathIRNode.Type classifyLiteralToken(String value) {
        if (value == null || value.isEmpty()) {
            return MathIRNode.Type.TEXT;
        }
        if (value.chars().allMatch(Character::isDigit)) {
            return MathIRNode.Type.NUMBER;
        }
        if (value.chars().allMatch(Character::isLetter)) {
            return MathIRNode.Type.IDENT;
        }
        return MathIRNode.Type.OPERATOR;
    }

    private MathIRNode.Type classifyMappedCommand(String command, MtefCharMap.CharEntry entry) {
        if (isBigOperatorCommand(command)) {
            return MathIRNode.Type.OPERATOR;
        }
        if (entry.typeface() == MtefRecord.FN_LC_GREEK
            || entry.typeface() == MtefRecord.FN_UC_GREEK
            || entry.typeface() == MtefRecord.FN_VARIABLE
            || entry.typeface() == MtefRecord.FN_FUNCTION) {
            return MathIRNode.Type.IDENT;
        }
        if (entry.typeface() == MtefRecord.FN_NUMBER) {
            return MathIRNode.Type.NUMBER;
        }
        return MathIRNode.Type.OPERATOR;
    }

    private String displayValue(String command, MtefCharMap.CharEntry entry) {
        if (command != null && command.startsWith("\\") && FUNCTION_COMMANDS.contains(command)) {
            return command.substring(1);
        }
        return new String(Character.toChars(entry.mtcode()));
    }

    private boolean isEnclosureCommand(String command) {
        return "\\boxed".equals(command)
            || "\\cancel".equals(command)
            || "\\bcancel".equals(command)
            || "\\xcancel".equals(command);
    }

    private boolean isHorizontalBraceCommand(String command) {
        return HORIZONTAL_BRACES.contains(command);
    }

    private boolean isHorizontalBracketCommand(String command) {
        return HORIZONTAL_BRACKETS.contains(command);
    }

    private boolean isHorizontalFenceCommand(String command) {
        return isHorizontalBraceCommand(command) || isHorizontalBracketCommand(command);
    }

    private boolean isHorizontalFenceScript(LaTeXNode node, boolean overScript) {
        if (node == null || node.getType() != LaTeXNode.Type.COMMAND) {
            return false;
        }
        String command = node.getValue();
        return overScript ? isTopHorizontalFence(command) : isBottomHorizontalFence(command);
    }

    private boolean isTopHorizontalFence(String command) {
        return "\\overbrace".equals(command) || "\\overbracket".equals(command);
    }

    private boolean isBottomHorizontalFence(String command) {
        return "\\underbrace".equals(command) || "\\underbracket".equals(command);
    }

    private String enclosureNotation(String command) {
        return switch (command) {
            case "\\boxed" -> "box";
            case "\\cancel" -> "updiagonalstrike";
            case "\\bcancel" -> "downdiagonalstrike";
            case "\\xcancel" -> "updiagonalstrike downdiagonalstrike";
            default -> command;
        };
    }

    private boolean isBigOperatorCommand(String command) {
        return UNDER_OVER_BIGOPS.contains(command) || SCRIPT_BIGOPS.contains(command);
    }

    private boolean isUnderOverOperator(LaTeXNode node) {
        return node != null
            && node.getType() == LaTeXNode.Type.COMMAND
            && node.getValue() != null
            && UNDER_OVER_BIGOPS.contains(node.getValue());
    }

    private String limitPlacement(String command) {
        return UNDER_OVER_BIGOPS.contains(command) ? "underover" : "script";
    }

    private String flattenText(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR || node.getType() == LaTeXNode.Type.COMMAND) {
            return node.getValue() == null ? "" : decodeTextLiteral(node.getValue());
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flattenText(child));
        }
        return builder.toString();
    }

    private String decodeTextLiteral(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }
        if (value.length() == 2 && value.charAt(0) == '\\' && !Character.isLetter(value.charAt(1))) {
            return value.substring(1);
        }
        return value;
    }

    private String stripLeftCommand(String command) {
        if (command == null || !command.startsWith("\\left")) {
            return "(";
        }
        return command.substring("\\left".length());
    }

    private String matchingRightDelimiter(String left) {
        return switch (left) {
            case "(" -> ")";
            case "[" -> "]";
            case "{", "\\{" -> "}";
            case "|" -> "|";
            default -> left;
        };
    }

    private String firstNonBlank(String primary, String fallback) {
        return primary != null && !primary.isBlank() ? primary : fallback;
    }

    private void copyMetadata(LaTeXNode source, MathIRNode target) {
        if (source == null || target == null) {
            return;
        }
        for (Map.Entry<String, String> entry : source.getMetadata().entrySet()) {
            target.setMetadata(entry.getKey(), entry.getValue());
        }
    }

    private LaTeXNode childAt(LaTeXNode node, int index) {
        return node != null && index >= 0 && index < node.getChildren().size()
            ? node.getChildren().get(index)
            : null;
    }
}
