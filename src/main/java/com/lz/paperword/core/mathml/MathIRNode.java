package com.lz.paperword.core.mathml;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * MathML-aligned internal representation node.
 *
 * <p>Phase 3 does not replace the existing LaTeX AST wholesale; instead it introduces a
 * canonical semantic layer between parser output and MTEF emission. The node structure here is
 * intentionally small and metadata-friendly so we can preserve existing layout-specific hints
 * while normalizing the stable MathML-like core subset.</p>
 */
public class MathIRNode {

    public enum Type {
        MATH,
        SEQUENCE,
        IDENT,
        NUMBER,
        OPERATOR,
        TEXT,
        FRACTION,
        SQRT,
        ROOT,
        SUB,
        SUP,
        SUBSUP,
        UNDER,
        OVER,
        UNDEROVER,
        FENCE,
        TABLE,
        TABLE_ROW,
        TABLE_CELL,
        LONG_DIVISION,
        UNSUPPORTED
    }

    private final Type type;
    private String value;
    private final List<MathIRNode> children = new ArrayList<>();
    private final Map<String, String> metadata = new LinkedHashMap<>();

    public MathIRNode(Type type) {
        this.type = type;
    }

    public MathIRNode(Type type, String value) {
        this.type = type;
        this.value = value;
    }

    public Type getType() {
        return type;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public List<MathIRNode> getChildren() {
        return children;
    }

    public void addChild(MathIRNode child) {
        if (child != null) {
            children.add(child);
        }
    }

    public MathIRNode child(int index) {
        return index >= 0 && index < children.size() ? children.get(index) : null;
    }

    public void setMetadata(String key, String value) {
        if (key == null) {
            return;
        }
        if (value == null) {
            metadata.remove(key);
            return;
        }
        metadata.put(key, value);
    }

    public String getMetadata(String key) {
        return metadata.get(key);
    }

    public Map<String, String> getMetadata() {
        return metadata;
    }

    @Override
    public String toString() {
        return "MathIRNode{" + type + ", value='" + value + "', children=" + children.size()
            + ", metadata=" + metadata + '}';
    }
}
