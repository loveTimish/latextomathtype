package com.lz.paperword.core.mtef;

/**
 * MTEF (MathType Equation Format) 记录类型常量定义类。
 *
 * <h2>概述</h2>
 * <p>本类定义了 MTEF v5 二进制格式中所有记录类型、选项标志、字体样式、
 * 模板选择器、变体位以及修饰符等常量。这些常量是构建 MTEF 二进制数据流的基础。</p>
 *
 * <h2>MTEF 二进制格式简介</h2>
 * <p>MTEF（MathType Equation Format）是 MathType 公式编辑器使用的二进制格式，
 * 版本 5 是目前最广泛使用的版本。该格式用于将数学公式存储为 OLE 对象，
 * 嵌入到 Word 等 Office 文档中。</p>
 *
 * <h2>记录结构</h2>
 * <p>MTEF 数据由一系列记录（Record）组成，每条记录以一个标签字节（tag）开头，
 * 标识记录类型。不同类型的记录具有不同的后续字段结构：</p>
 * <ul>
 *   <li><b>END (0)</b>：标记一个记录列表的结束</li>
 *   <li><b>LINE (1)</b>：行记录，包含一行公式内容（字符、模板等）</li>
 *   <li><b>CHAR (2)</b>：字符记录，表示单个数学字符</li>
 *   <li><b>TMPL (3)</b>：模板记录，表示分数、根号、上下标等数学结构</li>
 *   <li><b>PILE (4)</b>：堆叠记录，表示多行垂直排列的内容</li>
 *   <li><b>MATRIX (5)</b>：矩阵记录</li>
 *   <li><b>EMBELL (6)</b>：修饰记录，如上点、上箭头等字符装饰</li>
 *   <li><b>SIZE (9)</b>：字号控制记录</li>
 *   <li><b>FONT_DEF (17)</b>：字体定义记录</li>
 * </ul>
 *
 * <h2>在整体管线中的角色</h2>
 * <p>本类作为 MTEF 生成管线中最底层的常量定义，被 {@link MtefCharMap}（字符映射）、
 * {@link MtefTemplateBuilder}（模板构建器）以及上层的 MTEF 二进制序列化器共同引用。
 * LaTeX AST → MTEF 二进制的转换过程中，每一步都需要引用本类中的常量来写入正确的字节。</p>
 *
 * <p>移植自 docxtolatex/eqn/record.go</p>
 */
public final class MtefRecord {

    /** 私有构造函数，防止实例化——本类仅提供静态常量 */
    private MtefRecord() {}

    // ======================== 记录类型（Record Types）========================
    // 每条 MTEF 记录的第一个字节即为记录类型标签（tag）

    /** 结束记录：标志一个记录列表（如 LINE 内的子记录序列）的结尾 */
    public static final int END            = 0;
    /** 行记录：公式中的一行内容，包含字符、模板等子记录 */
    public static final int LINE           = 1;
    /** 字符记录：表示一个数学字符，包含字体索引和字符编码 */
    public static final int CHAR           = 2;
    /** 模板记录：表示分数、根号、上下标等数学结构 */
    public static final int TMPL           = 3;
    /** 堆叠记录：垂直排列的多行内容（如 cases 环境） */
    public static final int PILE           = 4;
    /** 矩阵记录：表示矩阵结构 */
    public static final int MATRIX         = 5;
    /** 修饰记录：字符上的装饰，如点、横线、箭头等 */
    public static final int EMBELL         = 6;
    /** 标尺记录：定义制表位 */
    public static final int RULER          = 7;
    /** 字体样式定义记录：将字体样式ID映射到具体字体 */
    public static final int FONT_STYLE_DEF = 8;
    /** 字号记录：改变后续内容的字号大小 */
    public static final int SIZE           = 9;

    // 以下为 SIZE 记录的子类型，指定具体的字号层级
    /** 全尺寸：公式的标准字号 */
    public static final int FULL           = 10;
    /** 下标尺寸：第一级下标/上标的缩小字号 */
    public static final int SUB            = 11;
    /** 二级下标尺寸：下标的下标使用的更小字号 */
    public static final int SUB2           = 12;
    /** 符号尺寸：大型运算符（如求和号）的字号 */
    public static final int SYM            = 13;
    /** 子符号尺寸：大型运算符在下标位置时的字号 */
    public static final int SUBSYM         = 14;

    /** 颜色记录：设置后续内容的颜色 */
    public static final int COLOR          = 15;
    /** 颜色定义记录：定义颜色表中的颜色 */
    public static final int COLOR_DEF      = 16;
    /** 字体定义记录：在字体表中定义一个字体（名称+样式） */
    public static final int FONT_DEF       = 17;
    /** 公式首选项记录：包含公式的全局设置 */
    public static final int EQN_PREFS      = 18;
    /** 编码定义记录：定义字符编码方式 */
    public static final int ENCODING_DEF   = 19;
    /** 保留给未来使用的记录类型 */
    public static final int FUTURE         = 100;
    /** 根记录：MTEF 数据流的虚拟根节点（不实际写入二进制流） */
    public static final int ROOT           = 255;

    // ======================== 选项标志（Option Flags）========================
    // 记录标签后紧跟的 options 字节中的各个位标志

    /** 微移选项：表示记录包含水平/垂直微移数据 */
    public static final int OPT_NUDGE            = 0x08;
    /** 字符修饰选项：表示 CHAR 记录后跟随 EMBELL 修饰记录 */
    public static final int OPT_CHAR_EMBELL      = 0x01;
    /** 函数起始选项：标记该字符是函数名（如 sin、cos）的一部分 */
    public static final int OPT_CHAR_FUNC_START  = 0x02;
    /** 8位编码选项：CHAR 记录包含一个额外的 8 位字体编码字节（bits8） */
    public static final int OPT_CHAR_ENC_CHAR_8  = 0x04;
    /** 16位编码选项：CHAR 记录的 mtcode 使用 16 位（2字节）编码 */
    public static final int OPT_CHAR_ENC_CHAR_16 = 0x10;
    /** 无 mtcode 选项：CHAR 记录省略 mtcode 字段 */
    public static final int OPT_CHAR_ENC_NO_MTCODE = 0x20;
    /** 空行选项：标记 LINE 记录为空行（NULL LINE），不包含任何内容 */
    public static final int OPT_LINE_NULL        = 0x01;
    /** 行间距选项：LINE 记录包含行间距数据 */
    public static final int OPT_LINE_LSPACE      = 0x04;
    /** 行属性标尺选项：LINE 包含标尺引用 */
    public static final int OPT_LP_RULER         = 0x02;

    // ======================== 字体/字型样式（Font/Typeface Styles）========================
    // 这些 ID 用于 CHAR 记录中的 typeface 字段，标识字符所属的字体样式类别。
    // MathType 内部将不同类型的数学字符分配到不同的字体样式（Typeface Style）中。

    /** 文本样式：普通文本字符（标点符号、括号等） */
    public static final int FN_TEXT      = 1;
    /** 函数样式：数学函数名（sin, cos, log 等）使用的直立体 */
    public static final int FN_FUNCTION  = 2;
    /** 变量样式：数学变量（a, b, x, y 等）使用的斜体 */
    public static final int FN_VARIABLE  = 3;
    /** 小写希腊字母样式：α, β, γ 等 */
    public static final int FN_LC_GREEK  = 4;
    /** 大写希腊字母样式：Γ, Δ, Θ 等 */
    public static final int FN_UC_GREEK  = 5;
    /** 符号样式：数学运算符和特殊符号（±, ≤, ∈ 等），使用 Symbol 字体 */
    public static final int FN_SYMBOL    = 6;
    /** 向量/矩阵样式：粗体字母用于向量和矩阵名 */
    public static final int FN_VECTOR    = 7;
    /** 数字样式：数字 0-9 */
    public static final int FN_NUMBER    = 8;
    /** 用户自定义样式1 */
    public static final int FN_USER1     = 9;
    /** 用户自定义样式2 */
    public static final int FN_USER2     = 10;
    /** MT Extra 字体样式：MathType 扩展符号字体中的字符 */
    public static final int FN_MTEXTRA   = 11;
    /** 远东文本样式：用于中日韩等远东字符 */
    public static final int FN_TEXT_FE   = 12;
    /** 可扩展字体样式：可伸缩的大型分隔符 */
    public static final int FN_EXPAND    = 22;
    /** 标记样式 */
    public static final int FN_MARKER    = 23;
    /** 空格样式：不同宽度的数学空格字符 */
    public static final int FN_SPACE     = 24;

    // ======================== 模板选择器（Template Selectors）========================
    // 模板选择器定义了 TMPL 记录表示的数学结构类型。
    // 每种模板类型对应一种特定的数学排版结构，通过 variation 位进一步细化其外观。

    // --- 围栏/括号类模板（Fences）---
    /** 尖括号模板：⟨ ⟩ */
    public static final int TM_ANGLE    = 0;
    /** 圆括号模板：( ) */
    public static final int TM_PAREN    = 1;
    /** 花括号模板：{ } */
    public static final int TM_BRACE    = 2;
    /** 方括号模板：[ ] */
    public static final int TM_BRACK    = 3;
    /** 单竖线模板：| | （绝对值） */
    public static final int TM_BAR      = 4;
    /** 双竖线模板：‖ ‖（范数） */
    public static final int TM_DBAR     = 5;
    /** 下取整模板：⌊ ⌋ */
    public static final int TM_FLOOR    = 6;
    /** 上取整模板：⌈ ⌉ */
    public static final int TM_CEILING  = 7;
    /** 开方括号模板（单侧括号） */
    public static final int TM_OBRACK   = 8;

    // --- 区间模板 ---
    /** 区间模板：可组合开闭区间 */
    public static final int TM_INTERVAL = 9;

    // --- 根号模板（Radicals）---
    /** 根号模板：平方根 √ 或 n 次根 ⁿ√，通过 variation 区分 */
    public static final int TM_ROOT     = 10;

    // --- 分数模板（Fractions）---
    /** 分数模板：分子/分母结构，含水平分数线 */
    public static final int TM_FRACT    = 11;

    // --- 上划线/下划线模板（Over/Under bars）---
    /** 下划线模板：在内容下方添加横线 */
    public static final int TM_UBAR     = 12;
    /** 上划线模板：在内容上方添加横线 */
    public static final int TM_OBAR     = 13;

    // --- 箭头模板（Arrows）---
    /** 箭头模板：带上/下方内容的箭头 */
    public static final int TM_ARROW    = 14;

    // --- 积分模板（Integrals）---
    /** 积分模板：∫ 积分符号，可带上下限 */
    public static final int TM_INTEG    = 15;

    // --- 大型运算符模板（Sums, Products, etc.）---
    /** 求和模板：∑ 求和符号，可带上下限 */
    public static final int TM_SUM      = 16;
    /** 连乘模板：∏ 连乘符号，可带上下限 */
    public static final int TM_PROD     = 17;
    /** 余积模板：∐ 余积符号 */
    public static final int TM_COPROD   = 18;
    /** 并集模板：⋃ 大并集符号 */
    public static final int TM_UNION    = 19;
    /** 交集模板：⋂ 大交集符号 */
    public static final int TM_INTER    = 20;
    /** 积分运算符模板：类积分样式的自定义运算符 */
    public static final int TM_INTOP    = 21;
    /** 求和运算符模板：类求和样式的自定义运算符 */
    public static final int TM_SUMOP    = 22;

    // --- 极限模板（Limits）---
    /** 极限模板：lim 等极限表达式 */
    public static final int TM_LIM      = 23;

    // --- 水平大括号模板（Horizontal braces/brackets）---
    /** 水平花括号模板：内容上方或下方的水平大括号 ⏞ ⏟ */
    public static final int TM_HBRACE   = 24;
    /** 水平方括号模板 */
    public static final int TM_HBRACK   = 25;

    // --- 长除法模板 ---
    /** 长除法模板 */
    public static final int TM_LDIV     = 26;

    // --- 上标/下标模板（Sub/Superscripts）---
    /** 下标模板：x₂ 形式，基底字符写在模板之前的父 LINE 中 */
    public static final int TM_SUB      = 27;
    /** 上标模板：x² 形式，基底字符写在模板之前的父 LINE 中 */
    public static final int TM_SUP      = 28;
    /** 上下标同时模板：x₂³ 形式，基底字符写在模板之前 */
    public static final int TM_SUBSUP   = 29;

    // --- Dirac 记号模板 ---
    /** Dirac 括号模板：⟨ψ|φ⟩ 等量子力学记号 */
    public static final int TM_DIRAC    = 30;

    // --- 向量箭头模板 ---
    /** 向量箭头模板：字符上方的向量箭头 →  */
    public static final int TM_VEC      = 31;

    // --- 帽子、弧线、波浪模板（Hats, arcs, tilde）---
    /** 波浪号模板：字符上方的波浪号 ~ */
    public static final int TM_TILDE    = 32;
    /** 帽子模板：字符上方的尖帽 ^ */
    public static final int TM_HAT      = 33;
    /** 弧线模板：字符上方的弧线 ⌒ */
    public static final int TM_ARC      = 34;
    /** 联合状态模板（内部使用） */
    public static final int TM_JSTATUS  = 35;

    // --- 删除线模板（Overstrikes）---
    /** 删除线模板：在内容上划删除线 */
    public static final int TM_STRIKE   = 36;

    // --- 方框模板（Boxes）---
    /** 方框模板：将内容放在方框中 */
    public static final int TM_BOX      = 37;

    // --- 左侧脚注模板 ---
    /** 左侧脚本模板：左侧的上标/下标 */
    public static final int TM_LSCRIPT  = 44;

    // ======================== 模板变体位（Template Variation Bits）========================
    // 变体位用于进一步细化模板的外观和行为。
    // variation 字段通常为 1 字节，若最高位为 1 则扩展为 2 字节。

    // --- 围栏变体（Fence variations）---
    /** 显示左围栏（左括号） */
    public static final int TV_FENCE_L = 0x0001;
    /** 显示右围栏（右括号） */
    public static final int TV_FENCE_R = 0x0002;

    // --- 根号变体（Root variations）---
    /** 平方根：不显示根指数 */
    public static final int TV_ROOT_SQ  = 0;
    /** n 次根：显示根指数槽位 */
    public static final int TV_ROOT_NTH = 1;

    // --- 区间变体（Interval variations）---
    /** 左侧为左圆括号 ( */
    public static final int TV_INTV_LEFT_LP  = 0x0000;
    /** 左侧为右圆括号 ) */
    public static final int TV_INTV_LEFT_RP  = 0x0001;
    /** 左侧为左方括号 [ */
    public static final int TV_INTV_LEFT_LB  = 0x0002;
    /** 左侧为右方括号 ] */
    public static final int TV_INTV_LEFT_RB  = 0x0003;
    /** 右侧为左圆括号 ( */
    public static final int TV_INTV_RIGHT_LP = 0x0000;
    /** 右侧为右圆括号 ) */
    public static final int TV_INTV_RIGHT_RP = 0x0010;
    /** 右侧为左方括号 [ */
    public static final int TV_INTV_RIGHT_LB = 0x0020;
    /** 右侧为右方括号 ] */
    public static final int TV_INTV_RIGHT_RB = 0x0030;

    // --- 分数变体（Fraction variations）---
    /** 小型分数（行内缩小版） */
    public static final int TV_FR_SMALL = 0x0001;
    /** 斜线分数（如 a/b 而非 $\frac{a}{b}$） */
    public static final int TV_FR_SLASH = 0x0002;
    /** 基线分数 */
    public static final int TV_FR_BASE  = 0x0004;

    // --- 横线变体（Bar variations）---
    /** 双横线（如双下划线） */
    public static final int TV_BAR_DOUBLE = 0x0001;

    // --- 积分变体（Integral variations）---
    /** 单重积分 ∫ */
    public static final int TV_INT_1       = 0x0001;
    /** 二重积分 ∬ */
    public static final int TV_INT_2       = 0x0002;
    /** 三重积分 ∭ */
    public static final int TV_INT_3       = 0x0003;
    /** 环路积分 ∮ */
    public static final int TV_INT_LOOP    = 0x0004;
    /** 积分号可展开标志（高位字节） */
    public static final int TV_INT_EXPAND  = 0x0100;

    // --- 大型运算符上下限变体（Limit variations for big operators）---
    // 第4-5位（bit4, bit5）编码上下限的存在性（已通过 MathType 7 输出验证）
    /** 存在下限 */
    public static final int TV_BO_LOWER = 0x0010;
    /** 存在上限 */
    public static final int TV_BO_UPPER = 0x0020;
    /** 求和式上下限位置（位于算子正下/正上）；否则按积分式显示在右侧 */
    public static final int TV_BO_SUM   = 0x0040;

    // --- 上下标变体（Sub/sup variations）---
    /** 脚本前置：上下标出现在基底字符之前（如左侧上下标） */
    public static final int TV_SU_PRECEDES = 0x0001;

    // --- 向量箭头变体（Vec variations）---
    /** 左向箭头 */
    public static final int TV_VE_LEFT    = 0x0001;
    /** 右向箭头 */
    public static final int TV_VE_RIGHT   = 0x0002;
    /** 下方箭头 */
    public static final int TV_VE_UNDER   = 0x0004;
    /** 半箭头（鱼叉形）⇀ */
    public static final int TV_VE_HARPOON = 0x0008;

    // --- 删除线变体（Strike variations）---
    /** 水平删除线 */
    public static final int TV_ST_HORIZ = 0x0001;
    /** 从左下到右上的斜杠（/） */
    public static final int TV_ST_UP    = 0x0002;
    /** 从左上到右下的反斜杠（\\） */
    public static final int TV_ST_DOWN  = 0x0004;

    // --- 方框变体（Box variations）---
    /** 圆角边框 */
    public static final int TV_BX_ROUND  = 0x0001;
    /** 显示左边框 */
    public static final int TV_BX_LEFT   = 0x0002;
    /** 显示右边框 */
    public static final int TV_BX_RIGHT  = 0x0004;
    /** 显示上边框 */
    public static final int TV_BX_TOP    = 0x0008;
    /** 显示下边框 */
    public static final int TV_BX_BOTTOM = 0x0010;

    // --- 水平大括号变体（HBrace variations）---
    /** 大括号在上方（否则在下方） */
    public static final int TV_HB_TOP = 0x0001;

    // --- Dirac 记号变体（Dirac variations）---
    /** 存在左侧 bra 部分 */
    public static final int TV_DI_LEFT  = 0x0001;
    /** 存在右侧 ket 部分 */
    public static final int TV_DI_RIGHT = 0x0002;

    // ======================== 修饰类型（Embellishment Types）========================
    // 修饰（Embellishment）是附加在字符上方或下方的装饰性标记。
    // 在 MTEF 中，修饰记录跟在被修饰的 CHAR 记录之后。

    /** 单点修饰：ẋ（一阶导数点） */
    public static final int EMB_1DOT      = 2;
    /** 双点修饰：ẍ（二阶导数点） */
    public static final int EMB_2DOT      = 3;
    /** 三点修饰 */
    public static final int EMB_3DOT      = 4;
    /** 单撇号：x′ */
    public static final int EMB_1PRIME    = 5;
    /** 双撇号：x″ */
    public static final int EMB_2PRIME    = 6;
    /** 反撇号 */
    public static final int EMB_BPRIME    = 7;
    /** 波浪号修饰：x̃ */
    public static final int EMB_TILDE     = 8;
    /** 帽子修饰：x̂ */
    public static final int EMB_HAT       = 9;
    /** 否定线修饰：在字符上划斜线表示"非" */
    public static final int EMB_NOT       = 10;
    /** 右箭头修饰：字符上方右向箭头 */
    public static final int EMB_RARROW    = 11;
    /** 左箭头修饰：字符上方左向箭头 */
    public static final int EMB_LARROW    = 12;
    /** 双向箭头修饰 */
    public static final int EMB_BARROW    = 13;
    /** 右半箭头修饰 */
    public static final int EMB_R1ARROW   = 14;
    /** 左半箭头修饰 */
    public static final int EMB_L1ARROW   = 15;
    /** 上横线修饰：x̄ */
    public static final int EMB_MBAR      = 16;
    /** 上双横线修饰 */
    public static final int EMB_OBAR      = 17;
    /** 三撇号：x‴ */
    public static final int EMB_3PRIME    = 18;
    /** 下弧修饰（皱眉形） */
    public static final int EMB_FROWN     = 19;
    /** 上弧修饰（微笑形） */
    public static final int EMB_SMILE     = 20;
    /** 交叉横线修饰 */
    public static final int EMB_X_BARS    = 21;
    /** 上方竖线修饰 */
    public static final int EMB_UP_BAR    = 22;
    /** 下方竖线修饰 */
    public static final int EMB_DOWN_BAR  = 23;
    /** 四点修饰 */
    public static final int EMB_4DOT      = 24;
    /** 下方单点修饰 */
    public static final int EMB_U_1DOT    = 25;
    /** 下方双点修饰 */
    public static final int EMB_U_2DOT    = 26;
    /** 下方三点修饰 */
    public static final int EMB_U_3DOT    = 27;
    /** 下方四点修饰 */
    public static final int EMB_U_4DOT    = 28;
    /** 下方横线修饰 */
    public static final int EMB_U_BAR     = 29;
    /** 下方波浪号修饰 */
    public static final int EMB_U_TILDE   = 30;
    /** 下方下弧修饰 */
    public static final int EMB_U_FROWN   = 31;
    /** 下方上弧修饰 */
    public static final int EMB_U_SMILE   = 32;
    /** 下方右箭头修饰 */
    public static final int EMB_U_RARROW  = 33;
    /** 下方左箭头修饰 */
    public static final int EMB_U_LARROW  = 34;
    /** 下方双向箭头修饰 */
    public static final int EMB_U_BARROW  = 35;
    /** 下方右半箭头修饰 */
    public static final int EMB_U_R1ARROW = 36;
    /** 下方左半箭头修饰 */
    public static final int EMB_U_L1ARROW = 37;

    // ======================== MTEF 文件头常量（Header Constants）========================
    // MTEF 二进制流的头部包含版本号、平台、产品信息等，用于标识格式和兼容性。

    /** MTEF 格式版本号：当前使用 v5 */
    public static final int MTEF_VERSION       = 5;
    /** 平台标识：Windows = 1 */
    public static final int MTEF_PLATFORM_WIN  = 1;
    /** 平台标识：Mac = 0 */
    public static final int MTEF_PLATFORM_MAC  = 0;

    // 产品标识（依据 Wiris MathType SDK 规范）：
    // 0 = MathType, 1 = Equation Editor
    /** 产品标识：MathType */
    public static final int MTEF_PRODUCT_MATHTYPE   = 0;
    /** 产品标识：Equation Editor（微软公式编辑器） */
    public static final int MTEF_PRODUCT_EQN_EDITOR = 1;
    /** 产品主版本号：匹配已知可正常工作的模板 */
    public static final int MTEF_PRODUCT_VERSION    = 6;
    /** 产品子版本号 */
    public static final int MTEF_PRODUCT_SUBVERSION = 9;

    /**
     * 应用程序标识键：写入 MTEF 二进制头时使用的以 null 结尾的字符串。
     * "DSMT6" 表示 "Design Science MathType 6"，是 MathType 使用的标准标识。
     */
    public static final String MTEF_APPLICATION_KEY = "DSMT6";
}
