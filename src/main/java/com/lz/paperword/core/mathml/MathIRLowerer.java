package com.lz.paperword.core.mathml;

import com.lz.paperword.core.latex.LaTeXNode;

import java.util.Map;

/**
 * Lowers MathML-aligned IR back into the existing normalized LaTeX AST shapes expected by the
 * current MTEF writer implementation.
 */
public class MathIRLowerer {

    public LaTeXNode lower(MathIRNode node) {
        if (node == null) {
            return new LaTeXNode(LaTeXNode.Type.ROOT);
        }
        return switch (node.getType()) {
            case MATH -> lowerContainer(node, LaTeXNode.Type.ROOT);
            case SEQUENCE -> lowerContainer(node, LaTeXNode.Type.GROUP);
            case IDENT, NUMBER, OPERATOR -> lowerToken(node);
            case TEXT -> lowerText(node);
            case FRACTION -> lowerFraction(node);
            case SQRT -> lowerSqrt(node);
            case ROOT -> lowerRoot(node);
            case SUB -> lowerSub(node);
            case SUP -> lowerSup(node);
            case SUBSUP -> lowerSubSup(node);
            case UNDER -> lowerUnder(node);
            case OVER -> lowerOver(node);
            case UNDEROVER -> lowerUnderOver(node);
            case FENCE -> lowerFence(node);
            case ENCLOSURE -> lowerEnclosure(node);
            case TABLE -> lowerContainer(node, LaTeXNode.Type.ARRAY);
            case TABLE_ROW -> lowerContainer(node, LaTeXNode.Type.ROW);
            case TABLE_CELL -> lowerContainer(node, LaTeXNode.Type.CELL);
            case LONG_DIVISION -> lowerLongDivision(node);
            case UNSUPPORTED -> throw new UnsupportedOperationException(buildUnsupportedMessage(node));
        };
    }

    private LaTeXNode lowerContainer(MathIRNode node, LaTeXNode.Type type) {
        LaTeXNode lowered = new LaTeXNode(type, node.getValue());
        copyMetadata(node, lowered);
        for (MathIRNode child : node.getChildren()) {
            lowered.addChild(lower(child));
        }
        return lowered;
    }

    private LaTeXNode lowerToken(MathIRNode node) {
        String latexCommand = node.getMetadata("latexCommand");
        LaTeXNode lowered;
        if (latexCommand != null && !latexCommand.isBlank()) {
            lowered = new LaTeXNode(LaTeXNode.Type.COMMAND, latexCommand);
        } else {
            lowered = new LaTeXNode(LaTeXNode.Type.CHAR, node.getValue() == null ? "" : node.getValue());
        }
        copyMetadata(node, lowered);
        return lowered;
    }

    private LaTeXNode lowerText(MathIRNode node) {
        String command = node.getMetadata("latexCommand");
        LaTeXNode text = new LaTeXNode(LaTeXNode.Type.TEXT, command == null ? "\\text" : command);
        copyMetadata(node, text);
        LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
        String value = node.getValue() == null ? "" : node.getValue();
        for (char ch : value.toCharArray()) {
            group.addChild(new LaTeXNode(LaTeXNode.Type.CHAR, String.valueOf(ch)));
        }
        text.addChild(group);
        return text;
    }

    private LaTeXNode lowerFraction(MathIRNode node) {
        LaTeXNode fraction = new LaTeXNode(LaTeXNode.Type.FRACTION, "\\frac");
        copyMetadata(node, fraction);
        fraction.addChild(lowerArgument(node.child(0)));
        fraction.addChild(lowerArgument(node.child(1)));
        return fraction;
    }

    private LaTeXNode lowerSqrt(MathIRNode node) {
        LaTeXNode sqrt = new LaTeXNode(LaTeXNode.Type.SQRT, "\\sqrt");
        copyMetadata(node, sqrt);
        sqrt.addChild(lowerArgument(node.child(0)));
        return sqrt;
    }

    private LaTeXNode lowerRoot(MathIRNode node) {
        LaTeXNode root = new LaTeXNode(LaTeXNode.Type.SQRT, "\\sqrt");
        copyMetadata(node, root);
        root.addChild(lowerArgument(node.child(0)));
        root.addChild(lowerArgument(node.child(1)));
        return root;
    }

    private LaTeXNode lowerSub(MathIRNode node) {
        LaTeXNode sub = new LaTeXNode(LaTeXNode.Type.SUBSCRIPT, "_");
        copyMetadata(node, sub);
        sub.addChild(lowerArgument(node.child(0)));
        sub.addChild(lowerArgument(node.child(1)));
        return sub;
    }

    private LaTeXNode lowerSup(MathIRNode node) {
        LaTeXNode sup = new LaTeXNode(LaTeXNode.Type.SUPERSCRIPT, "^");
        copyMetadata(node, sup);
        sup.addChild(lowerArgument(node.child(0)));
        sup.addChild(lowerArgument(node.child(1)));
        return sup;
    }

    private LaTeXNode lowerSubSup(MathIRNode node) {
        LaTeXNode sub = new LaTeXNode(LaTeXNode.Type.SUBSCRIPT, "_");
        copyMetadata(node, sub);
        sub.addChild(lowerArgument(node.child(0)));
        sub.addChild(lowerArgument(node.child(1)));

        LaTeXNode sup = new LaTeXNode(LaTeXNode.Type.SUPERSCRIPT, "^");
        copyMetadata(node, sup);
        sup.addChild(sub);
        sup.addChild(lowerArgument(node.child(2)));
        return sup;
    }

    private LaTeXNode lowerUnder(MathIRNode node) {
        String accentCommand = node.getMetadata("accentCommand");
        if (accentCommand != null && !accentCommand.isBlank()) {
            LaTeXNode command = new LaTeXNode(LaTeXNode.Type.COMMAND, accentCommand);
            copyMetadata(node, command);
            command.addChild(lowerArgument(node.child(0)));
            return command;
        }
        return lowerBigOperatorScript(node, false);
    }

    private LaTeXNode lowerOver(MathIRNode node) {
        String accentCommand = node.getMetadata("accentCommand");
        if (accentCommand != null && !accentCommand.isBlank()) {
            LaTeXNode command = new LaTeXNode(LaTeXNode.Type.COMMAND, accentCommand);
            copyMetadata(node, command);
            command.addChild(lowerArgument(node.child(0)));
            return command;
        }
        return lowerBigOperatorScript(node, true);
    }

    private LaTeXNode lowerUnderOver(MathIRNode node) {
        LaTeXNode sub = new LaTeXNode(LaTeXNode.Type.SUBSCRIPT, "_");
        copyMetadata(node, sub);
        sub.addChild(lowerArgument(node.child(0)));
        sub.addChild(lowerArgument(node.child(1)));

        LaTeXNode sup = new LaTeXNode(LaTeXNode.Type.SUPERSCRIPT, "^");
        copyMetadata(node, sup);
        sup.addChild(sub);
        sup.addChild(lowerArgument(node.child(2)));
        return sup;
    }

    private LaTeXNode lowerBigOperatorScript(MathIRNode node, boolean over) {
        MathIRNode base = node.child(0);
        if (!isUnderOverOperator(base)) {
            throw new UnsupportedOperationException(buildUnsupportedMessage(node));
        }
        LaTeXNode script = new LaTeXNode(over ? LaTeXNode.Type.SUPERSCRIPT : LaTeXNode.Type.SUBSCRIPT, over ? "^" : "_");
        copyMetadata(node, script);
        script.addChild(lowerArgument(base));
        script.addChild(lowerArgument(node.child(1)));
        return script;
    }

    private LaTeXNode lowerFence(MathIRNode node) {
        String open = node.getMetadata("openDelimiter");
        String close = node.getMetadata("closeDelimiter");
        String command = switch (open) {
            case "(" -> "\\left(";
            case "[" -> "\\left[";
            case "{", "\\{" -> "\\left{";
            case "|" -> "\\left|";
            case "||" -> "\\left\\lVert";
            case "⌊" -> "\\left\\lfloor";
            case "⌈" -> "\\left\\lceil";
            case "." -> "\\left.";
            default -> throw new UnsupportedOperationException(buildUnsupportedMessage(node));
        };
        LaTeXNode fence = new LaTeXNode(LaTeXNode.Type.COMMAND, command);
        copyMetadata(node, fence);
        fence.setMetadata("leftDelimiter", open);
        fence.setMetadata("rightDelimiter", close);
        fence.addChild(lowerArgument(node.child(0)));
        return fence;
    }

    private LaTeXNode lowerEnclosure(MathIRNode node) {
        String command = node.getMetadata("latexCommand");
        if (command == null || command.isBlank()) {
            command = switch (node.getMetadata("notation")) {
                case "box" -> "\\boxed";
                case "updiagonalstrike" -> "\\cancel";
                case "downdiagonalstrike" -> "\\bcancel";
                case "updiagonalstrike downdiagonalstrike" -> "\\xcancel";
                default -> throw new UnsupportedOperationException(buildUnsupportedMessage(node));
            };
        }
        LaTeXNode enclosure = new LaTeXNode(LaTeXNode.Type.COMMAND, command);
        copyMetadata(node, enclosure);
        enclosure.addChild(lowerArgument(node.child(0)));
        return enclosure;
    }

    private LaTeXNode lowerLongDivision(MathIRNode node) {
        LaTeXNode longDivision = new LaTeXNode(LaTeXNode.Type.LONG_DIVISION, "\\longdiv");
        copyMetadata(node, longDivision);
        longDivision.addChild(lowerArgument(node.child(0)));
        longDivision.addChild(lowerArgument(node.child(1)));
        longDivision.addChild(lowerArgument(node.child(2)));
        return longDivision;
    }

    private LaTeXNode lowerArgument(MathIRNode node) {
        if (node == null) {
            return new LaTeXNode(LaTeXNode.Type.GROUP);
        }
        LaTeXNode lowered = lower(node);
        if (lowered.getType() == LaTeXNode.Type.ROOT) {
            LaTeXNode group = new LaTeXNode(LaTeXNode.Type.GROUP);
            group.getMetadata().putAll(lowered.getMetadata());
            for (LaTeXNode child : lowered.getChildren()) {
                group.addChild(child);
            }
            return group;
        }
        return lowered;
    }

    private boolean isUnderOverOperator(MathIRNode node) {
        return node != null
            && (node.getType() == MathIRNode.Type.OPERATOR || node.getType() == MathIRNode.Type.IDENT)
            && "big-operator".equals(node.getMetadata("role"))
            && "underover".equals(node.getMetadata("limitPlacement"));
    }

    private void copyMetadata(MathIRNode source, LaTeXNode target) {
        if (source == null || target == null) {
            return;
        }
        for (Map.Entry<String, String> entry : source.getMetadata().entrySet()) {
            target.setMetadata(entry.getKey(), entry.getValue());
        }
    }

    private String buildUnsupportedMessage(MathIRNode node) {
        return "Unsupported MathIR node for MTEF lowering: " + node;
    }
}
