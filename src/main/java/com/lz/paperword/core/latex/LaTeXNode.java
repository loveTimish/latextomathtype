package com.lz.paperword.core.latex;

import java.util.ArrayList;
import java.util.List;

/**
 * LaTeX 抽象语法树（AST）节点类。
 *
 * <p>本类是 LaTeX 数学公式解析系统的核心数据结构，用于表示解析后的 LaTeX 表达式的树形结构。
 * 每个节点代表 LaTeX 表达式中的一个语法元素（字符、命令、分组、上下标等），
 * 通过父子关系构成一棵完整的抽象语法树。</p>
 *
 * <h3>AST 节点类型与子节点结构说明：</h3>
 * <ul>
 *   <li><b>ROOT</b>：根节点，包含整个表达式解析后的所有顶层子节点</li>
 *   <li><b>CHAR</b>：单个字符节点（字母、数字、运算符等），无子节点，value 存储字符值</li>
 *   <li><b>COMMAND</b>：LaTeX 命令节点（如 \alpha、\overline），value 存储命令名，
 *       一元命令（如 \overline）有一个子节点表示作用对象</li>
 *   <li><b>GROUP</b>：花括号分组 {...}，子节点为分组内的所有元素</li>
 *   <li><b>SUPERSCRIPT</b>：上标 base^{exp}，有两个子节点：children[0]=底数, children[1]=指数</li>
 *   <li><b>SUBSCRIPT</b>：下标 base_{sub}，有两个子节点：children[0]=底数, children[1]=下标</li>
 *   <li><b>FRACTION</b>：分数 \frac{num}{den}，有两个子节点：children[0]=分子, children[1]=分母</li>
 *   <li><b>SQRT</b>：根号 \sqrt{...} 或 \sqrt[n]{...}，
 *       无根次时有一个子节点（被开方数）；有根次时有两个子节点：children[0]=根次, children[1]=被开方数</li>
 *   <li><b>TEXT</b>：文本内容节点（如 \text{...}），value 存储命令名，子节点为文本内容</li>
 * </ul>
 *
 * <h3>AST 示例：</h3>
 * <p>对于 LaTeX 表达式 {@code \frac{x^{2}}{a+b}}，生成的 AST 结构为：</p>
 * <pre>
 * ROOT
 * └── FRACTION(\frac)
 *     ├── GROUP                    ← 分子
 *     │   └── SUPERSCRIPT(^)
 *     │       ├── CHAR(x)         ← 底数
 *     │       └── GROUP
 *     │           └── CHAR(2)     ← 指数
 *     └── GROUP                    ← 分母
 *         ├── CHAR(a)
 *         ├── CHAR(+)
 *         └── CHAR(b)
 * </pre>
 */
public class LaTeXNode {

    /**
     * LaTeX AST 节点类型枚举。
     *
     * <p>每种类型对应 LaTeX 表达式中的一种语法结构，
     * 决定了该节点在 AST 中的语义和子节点的组织方式。</p>
     */
    public enum Type {
        /** 根节点：包含整个 LaTeX 表达式的顶层节点，所有解析后的元素作为其子节点 */
        ROOT,
        /** 字符节点：表示单个字符（字母、数字、运算符如 +、-、=），value 中存储该字符 */
        CHAR,
        /** 命令节点：表示 LaTeX 命令（如 \frac、\alpha、\sqrt），value 中存储完整命令名（含反斜杠） */
        COMMAND,
        /** 分组节点：表示花括号包围的内容 {...}，子节点为分组内的各个元素 */
        GROUP,
        /** 上标节点：表示 base^{exp} 结构，children[0] 为底数，children[1] 为指数 */
        SUPERSCRIPT,
        /** 下标节点：表示 base_{sub} 结构，children[0] 为底数，children[1] 为下标内容 */
        SUBSCRIPT,
        /** 分数节点：表示 \frac{分子}{分母}，children[0] 为分子，children[1] 为分母 */
        FRACTION,
        /** 根号节点：表示 \sqrt{...} 或 \sqrt[n]{...}，可含 1~2 个子节点（根次 + 被开方数） */
        SQRT,
        /** 文本节点：表示非数学的文本内容（如 \text{...}、\mathrm{...}），子节点为文本内容 */
        TEXT
    }

    /** 节点类型（不可变），在构造时确定 */
    private final Type type;

    /** 节点的值：对于 CHAR 节点是字符本身，对于 COMMAND 节点是命令名（如 "\alpha"） */
    private String value;

    /** 子节点列表，用于构建树形结构。不同类型的节点对子节点数量和含义有不同约定 */
    private final List<LaTeXNode> children = new ArrayList<>();

    /**
     * 构造一个仅指定类型的节点（value 默认为 null）。
     * 通常用于创建 ROOT、GROUP 等不需要值的结构性节点。
     *
     * @param type 节点类型
     */
    public LaTeXNode(Type type) {
        this.type = type;
    }

    /**
     * 构造一个指定类型和值的节点。
     * 通常用于创建 CHAR（如 "x"）、COMMAND（如 "\frac"）等具有实际值的节点。
     *
     * @param type  节点类型
     * @param value 节点值（字符值或命令名）
     */
    public LaTeXNode(Type type, String value) {
        this.type = type;
        this.value = value;
    }

    /**
     * 获取节点类型。
     *
     * @return 节点的 {@link Type} 枚举值
     */
    public Type getType() {
        return type;
    }

    /**
     * 获取节点的值。
     *
     * @return 节点值字符串；CHAR 节点返回字符，COMMAND 节点返回命令名，结构性节点可能返回 null
     */
    public String getValue() {
        return value;
    }

    /**
     * 设置或更新节点的值。
     * 例如 \left 命令在解析时会将值更新为 "\left(" 等形式。
     *
     * @param value 新的节点值
     */
    public void setValue(String value) {
        this.value = value;
    }

    /**
     * 获取子节点列表。
     * 返回的是内部列表的直接引用，调用者可对其进行读取或遍历。
     *
     * @return 子节点列表
     */
    public List<LaTeXNode> getChildren() {
        return children;
    }

    /**
     * 向当前节点添加一个子节点。
     * 子节点按添加顺序排列，顺序对语义有重要影响
     * （如 FRACTION 的第一个子节点是分子，第二个是分母）。
     *
     * @param child 要添加的子节点
     */
    public void addChild(LaTeXNode child) {
        children.add(child);
    }

    /**
     * 返回节点的调试字符串表示，包含类型、值和子节点数量。
     * 用于调试和日志输出。
     *
     * @return 格式为 "LaTeXNode{TYPE, value='...', children=N}" 的字符串
     */
    @Override
    public String toString() {
        return "LaTeXNode{" + type + ", value='" + value + "', children=" + children.size() + '}';
    }
}
