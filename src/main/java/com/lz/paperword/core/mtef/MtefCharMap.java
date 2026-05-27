package com.lz.paperword.core.mtef;

import java.util.HashMap;
import java.util.Map;

/**
 * LaTeX 命令到 MTEF 字符条目的反向映射表。
 *
 * <h2>概述</h2>
 * <p>本类维护了一张从 LaTeX 命令/字符到 MTEF 字符信息（typeface + mtcode）的静态映射表，
 * 是 LaTeX AST → MTEF 二进制转换管线中字符处理的核心组件。</p>
 *
 * <h2>设计背景</h2>
 * <p>在 Go 语言的原始项目（docxtolatex/eqn/chars.go + chars_extended.go）中，
 * 映射方向为 mtcode → LaTeX（用于读取 MTEF 并转换为 LaTeX）。
 * 本类将映射方向反转为 LaTeX → (typeface, mtcode)，以支持将 LaTeX 写入 MTEF 二进制格式。</p>
 *
 * <h2>MTEF 字符编码概念</h2>
 * <ul>
 *   <li><b>typeface（字型样式ID）</b>：对应 {@link MtefRecord} 中的 FN_* 常量，
 *       标识字符所属的字体样式类别。例如 FN_LC_GREEK(4) 表示小写希腊字母，
 *       FN_SYMBOL(6) 表示数学符号等。MathType 根据 typeface 选择对应的显示字体。</li>
 *   <li><b>mtcode（MathType 字符码）</b>：通常为 Unicode 码点，是字符在 MTEF 中的 16 位编码。
 *       CHAR 记录中通过 typeface + mtcode 组合唯一确定一个数学字符。</li>
 *   <li><b>bits8（Symbol 字体编码字节）</b>：MathType 原生格式中，对于使用 Symbol 字体的字符，
 *       会额外存储一个 8 位的字形位置字节（ENC_CHAR_8），表示字符在 Symbol 字体中的
 *       glyph 索引。这是历史遗留的编码方式，确保与旧版 MathType 兼容。</li>
 * </ul>
 *
 * <h2>映射分类</h2>
 * <ol>
 *   <li><b>希腊小写字母</b>：typeface = FN_LC_GREEK(4)，如 \alpha → U+03B1</li>
 *   <li><b>希腊大写字母</b>：typeface = FN_UC_GREEK(5)，如 \Gamma → U+0393</li>
 *   <li><b>数学符号</b>：typeface = FN_SYMBOL(6)，如 \times → U+00D7, \leq → U+2264</li>
 *   <li><b>ASCII 字符</b>：根据字符类型分配不同 typeface：
 *       <ul>
 *         <li>字母 a-z, A-Z → FN_VARIABLE(3)（变量斜体）</li>
 *         <li>数字 0-9 → FN_NUMBER(8)</li>
 *         <li>运算符 +, -, = → FN_SYMBOL(6)</li>
 *         <li>标点、括号 → FN_TEXT(1)</li>
 *       </ul>
 *   </li>
 *   <li><b>大型运算符</b>：typeface = FN_SYMBOL(6)，如 \sum → U+2211, \int → U+222B。
 *       这些符号在 MTEF 中通常作为 TMPL 记录内的 SYM 子记录使用。</li>
 *   <li><b>空格</b>：typeface = FN_SPACE(24)，不同宽度的数学空格</li>
 * </ol>
 *
 * <h2>在整体管线中的角色</h2>
 * <p>上层 MTEF 序列化器在遇到 LaTeX AST 中的字符/命令节点时，
 * 调用 {@link #lookup(String)} 或 {@link #lookupChar(char)} 获取对应的
 * typeface 和 mtcode，然后写入 CHAR 记录。对于 Symbol 字体的字符，
 * 还需调用 {@link #getSymbolBits8(int)} 获取额外的 8 位编码字节。</p>
 *
 * <p>移植并反转自 docxtolatex/eqn/chars.go + chars_extended.go</p>
 */
public final class MtefCharMap {

    /** 私有构造函数，防止实例化 */
    private MtefCharMap() {}

    /**
     * 字符映射条目：包含字型样式 ID（typeface）和 16 位 MathType 字符码（mtcode）。
     *
     * @param typeface 字型样式 ID，对应 {@link MtefRecord} 中的 FN_* 常量
     * @param mtcode   MathType 字符码，通常为 Unicode 码点
     */
    public record CharEntry(int typeface, int mtcode) {}

    /**
     * 核心映射表：LaTeX 命令/字符字符串 → MTEF 字符条目。
     * 使用 HashMap 实现 O(1) 查找性能。
     */
    private static final Map<String, CharEntry> LATEX_TO_MTEF = new HashMap<>();

    static {
        // ==================== 希腊小写字母（typeface = FN_LC_GREEK = 4）====================
        // 将 LaTeX 希腊字母命令映射到对应的 Unicode 码点
        putGreekLower("\\alpha",      0x03B1);  // α
        putGreekLower("\\beta",       0x03B2);  // β
        putGreekLower("\\gamma",      0x03B3);  // γ
        putGreekLower("\\delta",      0x03B4);  // δ
        putGreekLower("\\epsilon",    0x03B5);  // ε
        putGreekLower("\\varepsilon", 0x03B5);  // ε（varepsilon 与 epsilon 共享同一码点）
        putGreekLower("\\zeta",       0x03B6);  // ζ
        putGreekLower("\\eta",        0x03B7);  // η
        putGreekLower("\\theta",      0x03B8);  // θ
        putGreekLower("\\vartheta",   0x03D1);  // ϑ（theta 的变体形式）
        putGreekLower("\\iota",       0x03B9);  // ι
        putGreekLower("\\kappa",      0x03BA);  // κ
        putGreekLower("\\lambda",     0x03BB);  // λ
        putGreekLower("\\mu",         0x03BC);  // μ
        putGreekLower("\\nu",         0x03BD);  // ν
        putGreekLower("\\xi",         0x03BE);  // ξ
        putGreekLower("\\pi",         0x03C0);  // π
        putGreekLower("\\varpi",      0x03D6);  // ϖ（pi 的变体形式）
        putGreekLower("\\rho",        0x03C1);  // ρ
        putGreekLower("\\varrho",     0x03F1);  // ϱ（rho 的变体形式）
        putGreekLower("\\sigma",      0x03C3);  // σ
        putGreekLower("\\varsigma",   0x03C2);  // ς（sigma 的词尾形式）
        putGreekLower("\\tau",        0x03C4);  // τ
        putGreekLower("\\upsilon",    0x03C5);  // υ
        putGreekLower("\\phi",        0x03D5);  // ϕ（phi 的直立变体）
        putGreekLower("\\varphi",     0x03C6);  // φ（phi 的弯曲变体）
        putGreekLower("\\chi",        0x03C7);  // χ
        putGreekLower("\\psi",        0x03C8);  // ψ
        putGreekLower("\\omega",      0x03C9);  // ω

        // ==================== 希腊大写字母（typeface = FN_UC_GREEK = 5）====================
        // 注意：只有与拉丁字母外形不同的大写希腊字母才需要映射
        // （如 A=Alpha, B=Beta 等与拉丁字母相同的已由 ASCII 映射覆盖）
        putGreekUpper("\\Gamma",   0x0393);  // Γ
        putGreekUpper("\\Delta",   0x0394);  // Δ
        putGreekUpper("\\Theta",   0x0398);  // Θ
        putGreekUpper("\\Lambda",  0x039B);  // Λ
        putGreekUpper("\\Xi",      0x039E);  // Ξ
        putGreekUpper("\\Pi",      0x03A0);  // Π
        putGreekUpper("\\Sigma",   0x03A3);  // Σ
        putGreekUpper("\\Upsilon", 0x03A5);  // Υ
        putGreekUpper("\\Phi",     0x03A6);  // Φ
        putGreekUpper("\\Psi",     0x03A8);  // Ψ
        putGreekUpper("\\Omega",   0x03A9);  // Ω

        // ==================== 数学符号（typeface = FN_SYMBOL = 6）====================
        // 常用数学运算符、关系符、集合符号、逻辑符号、箭头等

        // 算术运算符
        putSymbol("\\times",     0x00D7);  // × 乘号
        putSymbol("\\div",       0x00F7);  // ÷ 除号
        putSymbol("\\pm",        0x00B1);  // ± 正负号
        putSymbol("\\mp",        0x2213);  // ∓ 负正号
        putSymbol("\\cdot",      0x22C5);  // ⋅ 点乘
        putSymbol("\\bullet",    0x2022);  // • 实心圆点
        putSymbol("\\circ",      0x2218);  // ∘ 空心圆（复合运算符）
        putSymbol("\\ast",       0x2217);  // ∗ 星号运算符
        putSymbol("\\star",      0x22C6);  // ⋆ 星号

        // 关系运算符
        putSymbol("\\leq",       0x2264);  // ≤ 小于等于
        putSymbol("\\le",        0x2264);  // ≤（\leq 的简写）
        putSymbol("\\geq",       0x2265);  // ≥ 大于等于
        putSymbol("\\ge",        0x2265);  // ≥（\geq 的简写）
        putSymbol("\\neq",       0x2260);  // ≠ 不等于
        putSymbol("\\ne",        0x2260);  // ≠（\neq 的简写）
        putSymbol("\\approx",    0x2248);  // ≈ 约等于
        putSymbol("\\equiv",     0x2261);  // ≡ 恒等于
        putSymbol("\\sim",       0x223C);  // ∼ 相似
        putSymbol("\\simeq",     0x2243);  // ≃ 渐近相等
        putSymbol("\\cong",      0x2245);  // ≅ 全等/同余
        putSymbol("\\propto",    0x221D);  // ∝ 正比于
        putSymbol("\\perp",      0x22A5);  // ⊥ 垂直
        putSymbol("\\parallel",  0x2225);  // ∥ 平行

        // 集合运算符
        putSymbol("\\subset",    0x2282);  // ⊂ 真子集
        putSymbol("\\supset",    0x2283);  // ⊃ 真超集
        putSymbol("\\subseteq",  0x2286);  // ⊆ 子集
        putSymbol("\\supseteq",  0x2287);  // ⊇ 超集
        putSymbol("\\in",        0x2208);  // ∈ 属于
        putSymbol("\\notin",     0x2209);  // ∉ 不属于
        putSymbol("\\ni",        0x220B);  // ∋ 包含（反向属于）
        putSymbol("\\cup",       0x222A);  // ∪ 并集
        putSymbol("\\cap",       0x2229);  // ∩ 交集
        putSymbol("\\emptyset",  0x2205);  // ∅ 空集
        putSymbol("\\varnothing",0x2205);  // ∅（空集的变体命令）

        // 逻辑运算符
        putSymbol("\\exists",    0x2203);  // ∃ 存在量词
        putSymbol("\\forall",    0x2200);  // ∀ 全称量词
        putSymbol("\\neg",       0x00AC);  // ¬ 逻辑非
        putSymbol("\\lnot",      0x00AC);  // ¬（\neg 的别名）
        putSymbol("\\land",      0x2227);  // ∧ 逻辑与
        putSymbol("\\wedge",     0x2227);  // ∧（\land 的别名）
        putSymbol("\\lor",       0x2228);  // ∨ 逻辑或
        putSymbol("\\vee",       0x2228);  // ∨（\lor 的别名）

        // 其他数学符号
        putSymbol("\\oplus",     0x2295);  // ⊕ 圆加（直和）
        putSymbol("\\otimes",    0x2297);  // ⊗ 圆乘（张量积）
        putSymbol("\\infty",     0x221E);  // ∞ 无穷大
        putSymbol("\\partial",   0x2202);  // ∂ 偏导数
        putSymbol("\\nabla",     0x2207);  // ∇ 梯度算子（nabla）
        putSymbol("\\angle",     0x2220);  // ∠ 角
        putText("\\triangle",     0x25B3);  // △ 几何三角形标记
        putText("\\bigtriangleup",0x25B3);  // △ Pandoc 常输出该写法
        putSymbol("\\therefore", 0x2234);  // ∴ 所以
        putSymbol("\\because",   0x2235);  // ∵ 因为

        // 箭头符号
        putSymbol("\\rightarrow",0x2192);  // → 右箭头
        putSymbol("\\to",        0x2192);  // →（\rightarrow 的别名）
        putSymbol("\\leftarrow", 0x2190);  // ← 左箭头
        putSymbol("\\gets",      0x2190);  // ←（\leftarrow 的别名）
        putSymbol("\\leftrightarrow", 0x2194);  // ↔ 双向箭头
        putSymbol("\\Rightarrow",0x21D2);  // ⇒ 双线右箭头（蕴含）
        putSymbol("\\Leftarrow", 0x21D0);  // ⇐ 双线左箭头
        putSymbol("\\Leftrightarrow", 0x21D4);  // ⇔ 双线双向箭头（等价）
        putSymbol("\\uparrow",   0x2191);  // ↑ 上箭头
        putSymbol("\\downarrow", 0x2193);  // ↓ 下箭头
        // 斜箭头改走 MathType 官方默认的 MT Extra 样式槽位。
        // 参考十字交叉.docx 中的 OLE/MTEF 就是用 FN_MTEXTRA + U+2197/U+2198，
        // 这样公式编辑器能识别为 MathType 自己的已知字体，而不是外部自定义字体。
        putMtExtra("\\nearrow", 0x2197);  // ↗ 右上箭头
        putMtExtra("\\searrow", 0x2198);  // ↘ 右下箭头
        putMtExtra("\\swarrow", 0x2199);  // ↙ 左下箭头
        putMtExtra("\\nwarrow", 0x2196);  // ↖ 左上箭头

        // 省略号
        putSymbol("\\ldots",     0x2026);  // … 底部省略号
        putSymbol("\\cdots",     0x22EF);  // ⋯ 居中省略号
        putSymbol("\\vdots",     0x22EE);  // ⋮ 垂直省略号
        putSymbol("\\ddots",     0x22F1);  // ⋱ 对角省略号

        // 其他
        putSymbol("\\prime",     0x2032);  // ′ 撇号
        putSymbol("\\degree",    0x00B0);  // ° 度数符号

        // 大型运算符（在 TMPL 模板记录内作为 SYM 子记录使用）
        putSymbol("\\int",       0x222B);  // ∫ 积分号
        putSymbol("\\iint",      0x222C);  // ∬ 二重积分
        putSymbol("\\iiint",     0x222D);  // ∭ 三重积分
        putSymbol("\\oint",      0x222E);  // ∮ 环路积分
        putSymbol("\\intop",     0x222B);  // ∫ generic integral-style big operator
        putSymbol("\\sum",       0x2211);  // ∑ 求和号
        putSymbol("\\sumop",     0x2211);  // ∑ generic summation-style big operator
        putSymbol("\\prod",      0x220F);  // ∏ 连乘号
        putSymbol("\\coprod",    0x2210);  // ∐ 余积号
        putSymbol("\\bigcup",    0x22C3);  // ⋃ 大并集
        putSymbol("\\bigcap",    0x22C2);  // ⋂ 大交集
        putSymbol("\\bigvee",    0x22C1);  // ⋁ 大逻辑或
        putSymbol("\\bigwedge",  0x22C0);  // ⋀ 大逻辑与
        putSymbol("\\biguplus",  0x2A04);  // ⨄ n-ary union with plus
        putSymbol("\\bigoplus",  0x2A01);  // ⨁ n-ary circled plus
        putSymbol("\\bigotimes", 0x2A02);  // ⨂ n-ary circled times

        // ==================== ASCII 可打印字符默认映射 ====================
        // 将 ASCII 可打印字符范围 '!' (0x21) 到 '~' (0x7E) 默认映射为变量样式 FN_VARIABLE
        // 使用 putIfAbsent 语义：如果 LaTeX 命令已在上面的映射中注册（如希腊字母、符号），
        // 则保留已有映射不覆盖。
        for (char c = '!'; c <= '~'; c++) {
            String key = String.valueOf(c);
            if (!LATEX_TO_MTEF.containsKey(key)) {
                LATEX_TO_MTEF.put(key, new CharEntry(MtefRecord.FN_VARIABLE, c));
            }
        }

        // 覆盖数字字符为 FN_NUMBER（数字样式），数字在公式中应使用直立体
        for (char c = '0'; c <= '9'; c++) {
            LATEX_TO_MTEF.put(String.valueOf(c), new CharEntry(MtefRecord.FN_NUMBER, c));
        }

        // 覆盖常用数学运算符为 FN_SYMBOL（符号字体），
        // 这样 MathType 会为这些字符应用运算符样式的间距和外观，
        // 使输出更接近 MathType 官方效果。
        LATEX_TO_MTEF.put("+", new CharEntry(MtefRecord.FN_SYMBOL, '+'));
        LATEX_TO_MTEF.put("-", new CharEntry(MtefRecord.FN_SYMBOL, 0x2212)); // 数学减号（U+2212），非 ASCII 连字符
        LATEX_TO_MTEF.put("=", new CharEntry(MtefRecord.FN_SYMBOL, '='));
        LATEX_TO_MTEF.put("<", new CharEntry(MtefRecord.FN_SYMBOL, '<'));
        LATEX_TO_MTEF.put(">", new CharEntry(MtefRecord.FN_SYMBOL, '>'));

        // 标点符号和括号使用 FN_TEXT（文本样式），保持直立体显示
        LATEX_TO_MTEF.put(",", new CharEntry(MtefRecord.FN_TEXT, ','));
        LATEX_TO_MTEF.put(".", new CharEntry(MtefRecord.FN_TEXT, '.'));
        LATEX_TO_MTEF.put(":", new CharEntry(MtefRecord.FN_TEXT, ':'));
        LATEX_TO_MTEF.put(";", new CharEntry(MtefRecord.FN_TEXT, ';'));
        LATEX_TO_MTEF.put("!", new CharEntry(MtefRecord.FN_TEXT, '!'));
        LATEX_TO_MTEF.put("?", new CharEntry(MtefRecord.FN_TEXT, '?'));
        LATEX_TO_MTEF.put("(", new CharEntry(MtefRecord.FN_TEXT, '('));
        LATEX_TO_MTEF.put(")", new CharEntry(MtefRecord.FN_TEXT, ')'));
        LATEX_TO_MTEF.put("[", new CharEntry(MtefRecord.FN_TEXT, '['));
        LATEX_TO_MTEF.put("]", new CharEntry(MtefRecord.FN_TEXT, ']'));
        LATEX_TO_MTEF.put("{", new CharEntry(MtefRecord.FN_TEXT, '{'));
        LATEX_TO_MTEF.put("}", new CharEntry(MtefRecord.FN_TEXT, '}'));
        LATEX_TO_MTEF.put("|", new CharEntry(MtefRecord.FN_TEXT, '|'));
        LATEX_TO_MTEF.put("/", new CharEntry(MtefRecord.FN_TEXT, '/'));

        // 空格字符映射：不同 LaTeX 空格命令对应不同宽度的 Unicode 空格
        LATEX_TO_MTEF.put(" ", new CharEntry(MtefRecord.FN_TEXT, ' '));         // 普通空格
        LATEX_TO_MTEF.put("\\,", new CharEntry(MtefRecord.FN_SPACE, 0x2006)); // 细空格（Six-Per-Em Space）
        LATEX_TO_MTEF.put("\\;", new CharEntry(MtefRecord.FN_SPACE, 0x2005)); // 中等空格（Four-Per-Em Space）
        LATEX_TO_MTEF.put("\\quad", new CharEntry(MtefRecord.FN_SPACE, 0x2001)); // 全方空格（Em Quad）

        // LaTeX 转义字符：需要反斜杠转义的特殊字符
        LATEX_TO_MTEF.put("\\%", new CharEntry(MtefRecord.FN_TEXT, '%'));
        LATEX_TO_MTEF.put("\\_", new CharEntry(MtefRecord.FN_TEXT, '_'));
        LATEX_TO_MTEF.put("\\{", new CharEntry(MtefRecord.FN_TEXT, '{'));
        LATEX_TO_MTEF.put("\\}", new CharEntry(MtefRecord.FN_TEXT, '}'));
    }

    /**
     * 辅助方法：将 LaTeX 命令注册为小写希腊字母映射。
     *
     * @param latex  LaTeX 命令，如 "\\alpha"
     * @param mtcode Unicode 码点，如 0x03B1
     */
    private static void putGreekLower(String latex, int mtcode) {
        LATEX_TO_MTEF.put(latex, new CharEntry(MtefRecord.FN_LC_GREEK, mtcode));
    }

    /**
     * 辅助方法：将 LaTeX 命令注册为大写希腊字母映射。
     *
     * @param latex  LaTeX 命令，如 "\\Gamma"
     * @param mtcode Unicode 码点，如 0x0393
     */
    private static void putGreekUpper(String latex, int mtcode) {
        LATEX_TO_MTEF.put(latex, new CharEntry(MtefRecord.FN_UC_GREEK, mtcode));
    }

    /**
     * 辅助方法：将 LaTeX 命令注册为数学符号映射。
     *
     * @param latex  LaTeX 命令，如 "\\times"
     * @param mtcode Unicode 码点，如 0x00D7
     */
    private static void putSymbol(String latex, int mtcode) {
        LATEX_TO_MTEF.put(latex, new CharEntry(MtefRecord.FN_SYMBOL, mtcode));
    }

    /**
     * 辅助方法：将 LaTeX 命令注册为普通文本字符映射。
     *
     * <p>几何题中的 △ 更适合使用普通文本字体渲染，而不是复用 Symbol/Greek 字形，
     * 否则容易显示成接近 Δ 的外观。</p>
     *
     * @param latex  LaTeX 命令，如 "\\bigtriangleup"
     * @param mtcode Unicode 码点，如 0x25B3
     */
    private static void putText(String latex, int mtcode) {
        LATEX_TO_MTEF.put(latex, new CharEntry(MtefRecord.FN_TEXT, mtcode));
    }

    /**
     * 辅助方法：将 LaTeX 命令注册为 MT Extra 样式字符。
     *
     * <p>MT Extra 是 MathType 官方默认字体之一，
     * 适合承载公式编辑器已知的特殊数学符号。</p>
     *
     * @param latex  LaTeX 命令，如 "\\nearrow"
     * @param mtcode Unicode 码点，如 0x2197
     */
    private static void putMtExtra(String latex, int mtcode) {
        LATEX_TO_MTEF.put(latex, new CharEntry(MtefRecord.FN_MTEXTRA, mtcode));
    }

    /**
     * 根据 LaTeX 命令或字符查找对应的 MTEF 字符条目。
     * <p>支持 LaTeX 命令（如 "\\alpha"）和单个字符（如 "x"）的查找。</p>
     *
     * @param latex LaTeX 命令或字符字符串
     * @return 对应的 {@link CharEntry}，包含 typeface 和 mtcode；未找到则返回 null
     */
    public static CharEntry lookup(String latex) {
        return LATEX_TO_MTEF.get(latex);
    }

    /**
     * 根据单个字符查找对应的 MTEF 字符条目。
     * <p>内部将字符转为字符串后查询映射表。</p>
     *
     * @param c 要查找的字符
     * @return 对应的 {@link CharEntry}；未找到则返回 null
     */
    public static CharEntry lookupChar(char c) {
        return LATEX_TO_MTEF.get(String.valueOf(c));
    }

    // ==================== Unicode → Symbol 字体编码映射（bits8）====================
    // MathType 的原生格式对使用 Symbol 字体的字符，除了存储 16 位的 mtcode（Unicode 码点）外，
    // 还会额外存储一个 8 位的字节（通过 OPT_CHAR_ENC_CHAR_8 选项标记），
    // 该字节表示字符在 Symbol 字体中的字形索引位置（glyph position）。
    // 这是因为 Symbol 字体使用的是自定义编码（非 Unicode），需要这个额外字节来定位字形。

    /**
     * Unicode 码点 → Symbol 字体 8 位字形位置映射表。
     * 用于生成 CHAR 记录中 ENC_CHAR_8 标志对应的额外字节。
     */
    private static final Map<Integer, Integer> UNICODE_TO_SYMBOL_BITS8 = new HashMap<>();
    static {
        // --- 希腊小写字母在 Symbol 字体中的位置 ---
        UNICODE_TO_SYMBOL_BITS8.put(0x03B1, 0x61); // α → 'a'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B2, 0x62); // β → 'b'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B3, 0x67); // γ → 'g'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B4, 0x64); // δ → 'd'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B5, 0x65); // ε → 'e'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B6, 0x7A); // ζ → 'z'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B7, 0x68); // η → 'h'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B8, 0x71); // θ → 'q'
        UNICODE_TO_SYMBOL_BITS8.put(0x03B9, 0x69); // ι → 'i'
        UNICODE_TO_SYMBOL_BITS8.put(0x03BA, 0x6B); // κ → 'k'
        UNICODE_TO_SYMBOL_BITS8.put(0x03BB, 0x6C); // λ → 'l'
        UNICODE_TO_SYMBOL_BITS8.put(0x03BC, 0x6D); // μ → 'm'
        UNICODE_TO_SYMBOL_BITS8.put(0x03BD, 0x6E); // ν → 'n'
        UNICODE_TO_SYMBOL_BITS8.put(0x03BE, 0x78); // ξ → 'x'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C0, 0x70); // π → 'p'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C1, 0x72); // ρ → 'r'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C2, 0x56); // ς（词尾 sigma）→ 'V'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C3, 0x73); // σ → 's'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C4, 0x74); // τ → 't'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C5, 0x75); // υ → 'u'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C6, 0x66); // φ → 'f'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C7, 0x63); // χ → 'c'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C8, 0x79); // ψ → 'y'
        UNICODE_TO_SYMBOL_BITS8.put(0x03C9, 0x77); // ω → 'w'
        UNICODE_TO_SYMBOL_BITS8.put(0x03D1, 0x4A); // ϑ (vartheta) → 'J'
        UNICODE_TO_SYMBOL_BITS8.put(0x03D5, 0x6A); // ϕ (phi variant) → 'j'
        UNICODE_TO_SYMBOL_BITS8.put(0x03D6, 0x76); // ϖ (varpi) → 'v'
        UNICODE_TO_SYMBOL_BITS8.put(0x03F1, 0x72); // ϱ (varrho) → 'r'

        // --- 希腊大写字母在 Symbol 字体中的位置 ---
        UNICODE_TO_SYMBOL_BITS8.put(0x0393, 0x47); // Γ → 'G'
        UNICODE_TO_SYMBOL_BITS8.put(0x0394, 0x44); // Δ → 'D'
        UNICODE_TO_SYMBOL_BITS8.put(0x0398, 0x51); // Θ → 'Q'
        UNICODE_TO_SYMBOL_BITS8.put(0x039B, 0x4C); // Λ → 'L'
        UNICODE_TO_SYMBOL_BITS8.put(0x039E, 0x58); // Ξ → 'X'
        UNICODE_TO_SYMBOL_BITS8.put(0x03A0, 0x50); // Π → 'P'
        UNICODE_TO_SYMBOL_BITS8.put(0x03A3, 0x53); // Σ → 'S'
        UNICODE_TO_SYMBOL_BITS8.put(0x03A5, 0x55); // Υ → 'U'
        UNICODE_TO_SYMBOL_BITS8.put(0x03A6, 0x46); // Φ → 'F'
        UNICODE_TO_SYMBOL_BITS8.put(0x03A8, 0x59); // Ψ → 'Y'
        UNICODE_TO_SYMBOL_BITS8.put(0x03A9, 0x57); // Ω → 'W'

        // --- 常用数学符号在 Symbol 字体中的位置 ---
        UNICODE_TO_SYMBOL_BITS8.put(0x00D7, 0xB4); // × 乘号
        UNICODE_TO_SYMBOL_BITS8.put(0x00F7, 0xB8); // ÷ 除号
        UNICODE_TO_SYMBOL_BITS8.put(0x00B1, 0xB1); // ± 正负号
        UNICODE_TO_SYMBOL_BITS8.put(0x2212, 0x2D); // − 数学减号（映射到 ASCII '-' 的位置）
        UNICODE_TO_SYMBOL_BITS8.put(0x00AC, 0xD8); // ¬ 逻辑非
        UNICODE_TO_SYMBOL_BITS8.put(0x00B0, 0xB0); // ° 度数
        UNICODE_TO_SYMBOL_BITS8.put(0x2264, 0xA3); // ≤
        UNICODE_TO_SYMBOL_BITS8.put(0x2265, 0xB3); // ≥
        UNICODE_TO_SYMBOL_BITS8.put(0x2260, 0xB9); // ≠
        UNICODE_TO_SYMBOL_BITS8.put(0x2248, 0xBB); // ≈
        UNICODE_TO_SYMBOL_BITS8.put(0x2261, 0xBA); // ≡
        UNICODE_TO_SYMBOL_BITS8.put(0x221D, 0xB5); // ∝
        UNICODE_TO_SYMBOL_BITS8.put(0x221E, 0xA5); // ∞
        UNICODE_TO_SYMBOL_BITS8.put(0x2202, 0xB6); // ∂
        UNICODE_TO_SYMBOL_BITS8.put(0x2207, 0xD1); // ∇
        UNICODE_TO_SYMBOL_BITS8.put(0x2208, 0xCE); // ∈
        UNICODE_TO_SYMBOL_BITS8.put(0x2209, 0xCF); // ∉
        UNICODE_TO_SYMBOL_BITS8.put(0x220B, 0x27); // ∋
        UNICODE_TO_SYMBOL_BITS8.put(0x2227, 0xD9); // ∧
        UNICODE_TO_SYMBOL_BITS8.put(0x2228, 0xDA); // ∨
        UNICODE_TO_SYMBOL_BITS8.put(0x2229, 0xC7); // ∩
        UNICODE_TO_SYMBOL_BITS8.put(0x222A, 0xC8); // ∪
        UNICODE_TO_SYMBOL_BITS8.put(0x2282, 0xCC); // ⊂
        UNICODE_TO_SYMBOL_BITS8.put(0x2283, 0xC9); // ⊃
        UNICODE_TO_SYMBOL_BITS8.put(0x2286, 0xCD); // ⊆
        UNICODE_TO_SYMBOL_BITS8.put(0x2287, 0xCA); // ⊇
        UNICODE_TO_SYMBOL_BITS8.put(0x2295, 0xC5); // ⊕
        UNICODE_TO_SYMBOL_BITS8.put(0x2297, 0xC4); // ⊗
        UNICODE_TO_SYMBOL_BITS8.put(0x22A5, 0x5E); // ⊥
        UNICODE_TO_SYMBOL_BITS8.put(0x2220, 0xD0); // ∠
        UNICODE_TO_SYMBOL_BITS8.put(0x2234, 0x5C); // ∴
        UNICODE_TO_SYMBOL_BITS8.put(0x2235, 0x5D); // ∵
        UNICODE_TO_SYMBOL_BITS8.put(0x2190, 0xAC); // ←
        UNICODE_TO_SYMBOL_BITS8.put(0x2191, 0xAD); // ↑
        UNICODE_TO_SYMBOL_BITS8.put(0x2192, 0xAE); // →
        UNICODE_TO_SYMBOL_BITS8.put(0x2193, 0xAF); // ↓
        UNICODE_TO_SYMBOL_BITS8.put(0x2194, 0xAB); // ↔
        UNICODE_TO_SYMBOL_BITS8.put(0x21D0, 0xDC); // ⇐
        UNICODE_TO_SYMBOL_BITS8.put(0x21D2, 0xDE); // ⇒
        UNICODE_TO_SYMBOL_BITS8.put(0x21D4, 0xDB); // ⇔
        UNICODE_TO_SYMBOL_BITS8.put(0x2205, 0xC6); // ∅
        UNICODE_TO_SYMBOL_BITS8.put(0x2200, 0x22); // ∀
        UNICODE_TO_SYMBOL_BITS8.put(0x2203, 0x24); // ∃
        UNICODE_TO_SYMBOL_BITS8.put(0x22C5, 0xD7); // ⋅
        UNICODE_TO_SYMBOL_BITS8.put(0x2026, 0xBC); // …
        UNICODE_TO_SYMBOL_BITS8.put(0x2032, 0xA2); // ′
        UNICODE_TO_SYMBOL_BITS8.put(0x223C, 0x7E); // ∼
        UNICODE_TO_SYMBOL_BITS8.put(0x2225, 0xBD); // ∥
        UNICODE_TO_SYMBOL_BITS8.put(0x25B3, 0x44); // △（复用 Delta 字形位置 'D'）

        // --- 大型运算符在 Symbol 字体中的位置（已通过 MathType 7 二进制输出确认）---
        UNICODE_TO_SYMBOL_BITS8.put(0x222B, 0xF2); // ∫ 积分号
        UNICODE_TO_SYMBOL_BITS8.put(0x2211, 0xE5); // ∑ 求和号
        UNICODE_TO_SYMBOL_BITS8.put(0x220F, 0xD5); // ∏ 连乘号
        UNICODE_TO_SYMBOL_BITS8.put(0x2210, 0xD5); // ∐ 余积号（复用连乘号位置）
        UNICODE_TO_SYMBOL_BITS8.put(0x222E, 0xF2); // ∮ 环路积分（复用积分号位置）

        // ASCII 范围内的字符在 Symbol 字体中保持相同位置（即 Unicode 码点 = Symbol 字体索引）
        for (int c = 0x20; c < 0x7F; c++) {
            UNICODE_TO_SYMBOL_BITS8.putIfAbsent(c, c);
        }
    }

    /**
     * 获取指定 mtcode（Unicode 码点）在 Symbol 字体中的 8 位编码字节（bits8）。
     * <p>此值用于写入 CHAR 记录的 ENC_CHAR_8 附加字节，
     * 只有当字符使用 Symbol 字体（或其他非 Unicode 编码字体）时才需要。</p>
     *
     * @param mtcode 字符的 MathType 码（Unicode 码点）
     * @return Symbol 字体中的字形位置（0x00-0xFF），若无映射则返回 -1
     */
    public static int getSymbolBits8(int mtcode) {
        Integer bits8 = UNICODE_TO_SYMBOL_BITS8.get(mtcode);
        return bits8 != null ? bits8 : -1;
    }
}
