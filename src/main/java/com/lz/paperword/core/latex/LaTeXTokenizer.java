package com.lz.paperword.core.latex;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/**
 * LaTeX 数学表达式词法分析器（Tokenizer）。
 *
 * <p>本类负责将 LaTeX 数学公式字符串拆分为一系列有类型的词法单元（Token），
 * 是整个 LaTeX 解析流水线的第一步。输出的 Token 列表将作为 {@link LaTeXParser} 的输入，
 * 用于构建抽象语法树（AST）。</p>
 *
 * <h3>词法分析规则：</h3>
 * <ol>
 *   <li><b>反斜杠命令</b>（\command）：遇到 '\' 时，判断下一个字符：
 *     <ul>
 *       <li>若为字母，则连续读取所有字母组成命令名（如 \frac、\alpha）→ COMMAND 类型</li>
 *       <li>若为非字母（如 \{、\}、\%），则作为转义字符处理 → COMMAND 类型</li>
 *     </ul>
 *   </li>
 *   <li><b>花括号</b>：'{' → LBRACE，'}' → RBRACE（用于分组）</li>
 *   <li><b>方括号</b>：'[' → LBRACKET，']' → RBRACKET（用于可选参数，如 \sqrt[3]{x}）</li>
 *   <li><b>上标符</b>：'^' → CARET（触发上标解析）</li>
 *   <li><b>下标符</b>：'_' → UNDERSCORE（触发下标解析）</li>
 *   <li><b>空白字符</b>：生成 WHITESPACE Token，默认由解析器在数学模式中忽略；文本命令可显式保留</li>
 *   <li><b>普通字符</b>：字母、数字、运算符（+、-、=、<、>等）→ CHAR 类型</li>
 * </ol>
 *
 * <h3>分词示例：</h3>
 * <p>输入：{@code \frac{x^{2}}{a+b}}</p>
 * <p>输出 Token 序列：</p>
 * <pre>
 * COMMAND(\frac) → LBRACE({) → CHAR(x) → CARET(^) → LBRACE({) → CHAR(2) → RBRACE(})
 * → RBRACE(}) → LBRACE({) → CHAR(a) → CHAR(+) → CHAR(b) → RBRACE(})
 * </pre>
 *
 * <h3>已知命令集合：</h3>
 * <p>{@code KNOWN_COMMANDS} 维护了一份常见 LaTeX 命令清单（不含反斜杠前缀），
 * 包括希腊字母、数学运算符、关系符号、三角函数、大型运算符、装饰命令等。
 * 目前该集合仅供参考，分词逻辑并不依赖它来判断命令有效性——任何以 '\' 开头后跟字母的序列
 * 都会被识别为 COMMAND 类型的 Token。</p>
 */
public class LaTeXTokenizer {

    /**
     * 词法单元（Token）类型枚举。
     *
     * <p>每种类型对应 LaTeX 数学公式中的一类词法元素，
     * 解析器（Parser）将根据 Token 类型决定后续的语法解析行为。</p>
     */
    public enum TokenType {
        /** LaTeX 命令：以反斜杠开头，如 \frac、\alpha、\sin；也包括转义字符如 \{、\} */
        COMMAND,
        /** 普通字符：单个字母、数字或运算符，如 a、b、1、+、=、<、> */
        CHAR,
        /** 左花括号 '{'：标记分组或命令参数的开始 */
        LBRACE,
        /** 右花括号 '}'：标记分组或命令参数的结束 */
        RBRACE,
        /** 左方括号 '['：标记可选参数的开始（如 \sqrt[n]{x} 中的根次） */
        LBRACKET,
        /** 右方括号 ']'：标记可选参数的结束 */
        RBRACKET,
        /** 上标符号 '^'：触发上标（SUPERSCRIPT）的解析 */
        CARET,
        /** 下标符号 '_'：触发下标（SUBSCRIPT）的解析 */
        UNDERSCORE,
        /** 空白字符：默认可被解析器忽略，但 \text{...} 这类文本命令会显式保留 */
        WHITESPACE
    }

    /**
     * 词法单元记录类。
     * 每个 Token 包含类型和原始值两个属性，使用 Java Record 实现（不可变）。
     *
     * @param type  Token 的类型（命令、字符、花括号等）
     * @param value Token 的原始字符串值（如 "\frac"、"x"、"{"）
     */
    public record Token(TokenType type, String value) {
        @Override
        public String toString() {
            return type + "(" + value + ")";
        }
    }

    /**
     * 已知 LaTeX 命令名集合（不含反斜杠前缀）。
     *
     * <p>包含以下类别的命令：</p>
     * <ul>
     *   <li>结构命令：frac、sqrt、left、right、begin、end</li>
     *   <li>希腊字母：alpha、beta、gamma、...、Omega（含大小写变体）</li>
     *   <li>二元运算符：times、div、pm、mp、cdot 等</li>
     *   <li>关系运算符：leq、geq、neq、approx、equiv 等</li>
     *   <li>集合与逻辑：subset、supset、in、cup、cap、forall、exists 等</li>
     *   <li>箭头符号：rightarrow、leftarrow、Rightarrow 等</li>
     *   <li>省略号：ldots、cdots、vdots、ddots</li>
     *   <li>三角函数与对数：sin、cos、tan、log、ln、exp、lim 等</li>
     *   <li>大型运算符：sum、int、prod、bigcup、bigcap 等</li>
     *   <li>装饰命令：overline、underline、hat、vec、bar 等</li>
     *   <li>文本/字体命令：text、mathrm、mathbf、mathit、mathcal、mathbb</li>
     *   <li>间距命令：quad、qquad</li>
     * </ul>
     */
    private static final Set<String> KNOWN_COMMANDS = Set.of(
        "frac", "sqrt", "left", "right", "begin", "end",
        "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon",
        "zeta", "eta", "theta", "vartheta", "iota", "kappa",
        "lambda", "mu", "nu", "xi", "pi", "varpi",
        "rho", "varrho", "sigma", "varsigma", "tau", "upsilon",
        "phi", "varphi", "chi", "psi", "omega",
        "Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi",
        "Sigma", "Upsilon", "Phi", "Psi", "Omega",
        "times", "div", "pm", "mp", "cdot", "bullet",
        "leq", "le", "geq", "ge", "neq", "ne",
        "approx", "equiv", "sim", "simeq", "cong", "propto",
        "perp", "parallel", "subset", "supset", "subseteq", "supseteq",
        "in", "notin", "ni", "cup", "cap",
        "emptyset", "varnothing", "exists", "forall",
        "neg", "lnot", "land", "wedge", "lor", "vee",
        "oplus", "otimes", "infty", "partial", "nabla",
        "angle", "triangle", "bigtriangleup", "therefore", "because",
        "rightarrow", "to", "leftarrow", "gets",
        "leftrightarrow", "Rightarrow", "Leftarrow", "Leftrightarrow",
        "uparrow", "downarrow",
        "ldots", "cdots", "vdots", "ddots",
        "prime", "degree",
        "sin", "cos", "tan", "cot", "sec", "csc",
        "arcsin", "arccos", "arctan",
        "sinh", "cosh", "tanh",
        "log", "ln", "exp", "lim", "max", "min",
        "det", "dim", "gcd",
        "sum", "int", "iint", "iiint", "oint", "prod", "coprod",
        "bigcup", "bigcap", "bigvee", "bigwedge",
        "overline", "underline", "hat", "tilde", "vec", "bar", "dot",
        "overbrace", "underbrace",
        "text", "mathrm", "mathbf", "mathit", "mathcal", "mathbb",
        "quad", "qquad",
        "not"
    );

    /**
     * 将 LaTeX 数学表达式字符串分词为 Token 列表。
     *
     * <p>核心分词算法采用单遍顺序扫描：从左到右逐字符读取输入字符串，
     * 根据当前字符类型决定生成何种 Token。对于反斜杠命令，采用贪婪匹配
     * （尽可能多地读取字母字符）。空白字符会作为 WHITESPACE token 保留，
     * 便于 \text{   } 这类文本命令显式承载长除法对齐空格。</p>
     *
     * <p>分词过程不做语法校验——即使输入的 LaTeX 语法不完整（如缺少右花括号），
     * 分词器仍会正常输出已识别的 Token 序列，语法错误留给后续的解析器处理。</p>
     *
     * @param input 原始 LaTeX 数学表达式字符串（不含 $ 分隔符）
     * @return 分词后的 Token 列表，按出现顺序排列
     */
    public List<Token> tokenize(String input) {
        List<Token> tokens = new ArrayList<>();
        int i = 0;
        int len = input.length();

        while (i < len) {
            char c = input.charAt(i);

            if (c == '\\') {
                // ---- 处理 LaTeX 命令或转义字符 ----
                // 遇到反斜杠，需要判断后续字符来确定 Token 类型
                i++;
                if (i >= len) break; // 反斜杠在末尾，忽略
                char next = input.charAt(i);
                // 特殊转义字符：\\、\{、\}、\%、\&、\#、\$、\_
                // 下一个字符不是字母时，将 \ + 下一字符 作为一个 COMMAND Token
                if (!Character.isLetter(next)) {
                    tokens.add(new Token(TokenType.COMMAND, "\\" + next));
                    i++;
                } else {
                    // 命令名由连续字母组成，如 \frac、\alpha、\overline
                    // 使用贪婪匹配：持续读取直到遇到非字母字符
                    StringBuilder cmd = new StringBuilder("\\");
                    while (i < len && Character.isLetter(input.charAt(i))) {
                        cmd.append(input.charAt(i));
                        i++;
                    }
                    tokens.add(new Token(TokenType.COMMAND, cmd.toString()));
                }
            } else if (c == '{') {
                // ---- 左花括号：分组/参数开始 ----
                tokens.add(new Token(TokenType.LBRACE, "{"));
                i++;
            } else if (c == '}') {
                // ---- 右花括号：分组/参数结束 ----
                tokens.add(new Token(TokenType.RBRACE, "}"));
                i++;
            } else if (c == '[') {
                // ---- 左方括号：可选参数开始（如 \sqrt[3]{x} 中的 [3]） ----
                tokens.add(new Token(TokenType.LBRACKET, "["));
                i++;
            } else if (c == ']') {
                // ---- 右方括号：可选参数结束 ----
                tokens.add(new Token(TokenType.RBRACKET, "]"));
                i++;
            } else if (c == '^') {
                // ---- 上标运算符：触发后续上标表达式的解析 ----
                tokens.add(new Token(TokenType.CARET, "^"));
                i++;
            } else if (c == '_') {
                // ---- 下标运算符：触发后续下标表达式的解析 ----
                tokens.add(new Token(TokenType.UNDERSCORE, "_"));
                i++;
            } else if (Character.isWhitespace(c)) {
                // ---- 空白字符：保留为 WHITESPACE，便于文本命令恢复显式缩进 ----
                int start = i;
                while (i < len && Character.isWhitespace(input.charAt(i))) {
                    i++;
                }
                tokens.add(new Token(TokenType.WHITESPACE, input.substring(start, i)));
            } else {
                // ---- 普通字符：字母、数字、运算符等，每个字符生成一个 CHAR Token ----
                tokens.add(new Token(TokenType.CHAR, String.valueOf(c)));
                i++;
            }
        }

        return tokens;
    }
}
