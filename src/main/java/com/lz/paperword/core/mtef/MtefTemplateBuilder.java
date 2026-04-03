package com.lz.paperword.core.mtef;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * MTEF 模板（TMPL）记录构建器。
 *
 * <h2>概述</h2>
 * <p>本类封装了 MTEF v5 二进制格式中各类数学结构模板记录（TMPL record）的生成逻辑。
 * 每个公开方法对应一种特定的数学排版结构，将模板记录的二进制头部写入输出流。
 * 调用方在写入模板头后，需要继续写入模板的各个槽位（slot）内容。</p>
 *
 * <h2>TMPL 记录的二进制结构</h2>
 * <p>根据 MTEF v5 规范和 MathType 7 的实际输出验证，TMPL 记录的字节布局如下：</p>
 * <pre>
 *   tag(1字节)          - 记录类型标签，固定为 TMPL(3)
 *   options(1字节)      - 选项标志（如是否包含微移数据），通常为 0x00
 *   [nudge]             - 可选的微移数据（仅当 options 包含 OPT_NUDGE 时存在）
 *   selector(1字节)     - 模板选择器，标识数学结构类型（如 TM_FRACT=11 表示分数）
 *   variation(1-2字节)  - 变体位，细化模板外观（如是否显示左/右括号）
 *   tmplOptions(1字节)  - 模板附加选项，大多数情况下为 0x00
 * </pre>
 *
 * <h3>variation 字段的编码规则</h3>
 * <p>variation 通常为 1 字节。如果最高位（bit7）被设置，则表示 variation 扩展为 2 字节：
 * 第一个字节的低 7 位为 variation 的低 7 位，第二个字节为 variation 的高 8 位。
 * 这种变长编码用于需要大于 127 的 variation 值（如积分模板的 TV_INT_EXPAND=0x0100）。</p>
 *
 * <h2>模板槽位（Slots）说明</h2>
 * <p>TMPL 头部之后是若干个槽位（slot），每个槽位是一个 LINE 记录（或 NULL LINE）。
 * 不同模板类型的槽位数量和含义不同：</p>
 * <ul>
 *   <li><b>分数 (TM_FRACT)</b>：2 个槽位 — 分子 + 分母</li>
 *   <li><b>根号 (TM_ROOT)</b>：平方根 1 个槽位（被开方数）；n次根 2 个槽位（根指数 + 被开方数）</li>
 *   <li><b>上标 (TM_SUP)</b>：2 个槽位 — NULL LINE（基底占位）+ 上标内容</li>
 *   <li><b>下标 (TM_SUB)</b>：2 个槽位 — NULL LINE（基底占位）+ 下标内容</li>
 *   <li><b>上下标 (TM_SUBSUP)</b>：3 个槽位 — NULL LINE（基底占位）+ 下标 + 上标</li>
 *   <li><b>围栏 (TM_PAREN 等)</b>：1 个槽位 — 括号内的内容</li>
 *   <li><b>大型运算符 (TM_SUM, TM_INTEG 等)</b>：最多 3 个槽位 — 操作数 + 下限 + 上限</li>
 * </ul>
 *
 * <h2>上下标模板的特殊处理</h2>
 * <p>对于 TM_SUB、TM_SUP、TM_SUBSUP 模板，基底字符不包含在模板记录内部，
 * 而是写在模板记录之前的父 LINE 中。模板内的第一个槽位（slot 0）是一个 NULL LINE，
 * 表示基底位置的占位符。在模板记录之前，还需要写入一个 SUB 字号记录（1字节），
 * 将后续上下标内容的字号缩小。</p>
 *
 * <h2>在整体管线中的角色</h2>
 * <p>本类位于 MTEF 生成管线的中间层。上层序列化器在遍历 LaTeX AST 时，
 * 遇到分数、根号、上下标等数学结构节点，会调用本类的相应方法写入 TMPL 记录头，
 * 然后递归地将子节点内容写入各个槽位的 LINE 记录中。本类引用 {@link MtefRecord}
 * 中定义的模板选择器和变体位常量。</p>
 */
public class MtefTemplateBuilder {

    /**
     * 写入 TMPL 记录的通用头部。
     * <p>依次写入：tag(TMPL=3) + options(0x00) + selector + variation + tmplOptions。
     * 这是所有模板类型共用的底层写入方法。</p>
     *
     * @param out         输出字节流
     * @param selector    模板选择器（TM_* 常量），标识数学结构类型
     * @param variation   变体位（TV_* 常量的组合），细化模板外观
     * @param tmplOptions 模板附加选项，大多数情况下为 0x00
     * @throws IOException 写入异常
     */
    private static void writeTemplateHeader(ByteArrayOutputStream out, int selector,
                                             int variation, int tmplOptions) throws IOException {
        out.write(MtefRecord.TMPL);       // tag：记录类型 = TMPL(3)
        out.write(0x00);                  // options：无特殊选项
        out.write(selector & 0xFF);       // selector：模板选择器（取低 8 位）
        writeVariation(out, variation);   // variation：变体位（1-2 字节变长编码）
        out.write(tmplOptions & 0xFF);    // tmplOptions：模板附加选项
    }

    /**
     * 写入 variation 字段（变长编码：1 或 2 字节）。
     * <p>编码规则：如果 variation 值可以用 7 位表示（0-127），则写入 1 字节。
     * 否则第一个字节的最高位设为 1 表示后续还有一个字节，
     * 低 7 位存储 variation 的低 7 位，第二个字节存储 variation 的高 8 位。</p>
     *
     * @param out       输出字节流
     * @param variation 变体值
     * @throws IOException 写入异常
     */
    private static void writeVariation(ByteArrayOutputStream out, int variation) throws IOException {
        if ((variation & ~0x7F) == 0) {
            // variation 在 0-127 范围内，单字节编码即可
            out.write(variation & 0x7F);
            return;
        }
        // variation > 127，需要双字节编码：
        // 第一字节：低 7 位 | 0x80（最高位置 1 表示有后续字节）
        out.write((variation & 0x7F) | 0x80);
        // 第二字节：高 8 位
        out.write((variation >> 8) & 0xFF);
    }

    // ===== 分数模板（Fraction）=====

    /**
     * 写入分数模板（TM_FRACT）的头部。
     * <p>生成标准分数结构：水平分数线上方为分子，下方为分母。
     * 调用方写入头部后需依次写入 2 个槽位：分子 LINE + 分母 LINE。</p>
     * <p>variation = 0x00 表示标准大小的水平分数。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeFractionHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_FRACT, 0x00, 0x00);
    }

    // ===== 根号模板（Roots）=====

    /**
     * 写入平方根模板（TM_ROOT, TV_ROOT_SQ）的头部。
     * <p>生成 √ 平方根结构，无根指数显示。
     * 调用方写入头部后需写入 1 个槽位：被开方数 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeSqrtHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_ROOT, MtefRecord.TV_ROOT_SQ, 0x00);
    }

    /**
     * 写入 n 次根模板（TM_ROOT, TV_ROOT_NTH）的头部。
     * <p>生成 ⁿ√ n 次根结构，左上角显示根指数。
     * 调用方写入头部后需写入 2 个槽位：根指数 LINE + 被开方数 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeNthRootHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_ROOT, MtefRecord.TV_ROOT_NTH, 0x00);
    }

    // ===== 上标/下标模板（Sub/Superscript）=====

    /**
     * 写入上标模板（TM_SUP）的头部。
     * <p>生成 x² 形式的上标结构。注意：基底字符（如 x）需要在调用本方法之前
     * 写入到父 LINE 中，模板内的第一个槽位（slot 0）应为 NULL LINE 占位符。
     * 调用方写入头部后需写入 2 个槽位：NULL LINE（基底占位）+ 上标内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeSuperscriptHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_SUP, 0x00, 0x00);
    }

    /**
     * 写入下标模板（TM_SUB）的头部。
     * <p>生成 x₂ 形式的下标结构。基底字符的处理方式与上标相同——
     * 写在模板之前的父 LINE 中，模板 slot 0 为 NULL LINE。
     * 调用方写入头部后需写入 2 个槽位：NULL LINE（基底占位）+ 下标内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeSubscriptHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_SUB, 0x00, 0x00);
    }

    /**
     * 写入同时上下标模板（TM_SUBSUP）的头部。
     * <p>生成 x₂³ 形式的同时包含上标和下标的结构。基底字符写在模板之前的父 LINE 中。
     * 调用方写入头部后需写入 3 个槽位：NULL LINE（基底占位）+ 下标 LINE + 上标 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeSubSuperscriptHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_SUBSUP, 0x00, 0x00);
    }

    // ===== 大型运算符模板（Big Operators）=====

    /**
     * 写入求和模板（TM_SUM）的头部。
     * <p>生成 ∑ 求和符号结构，可选地带有上限和/或下限。
     * variation 的 bit4 和 bit5 分别控制下限和上限的存在性。
     * 调用方写入头部后需写入对应数量的槽位：
     * 操作数 LINE + [下限 LINE] + [上限 LINE]。</p>
     *
     * @param out      输出字节流
     * @param hasLower 是否存在下限（如 i=1）
     * @param hasUpper 是否存在上限（如 n）
     * @throws IOException 写入异常
     */
    public static void writeSumHeader(ByteArrayOutputStream out, boolean hasLower, boolean hasUpper) throws IOException {
        int variation = 0;
        if (hasLower) variation |= MtefRecord.TV_BO_LOWER;  // bit4：下限存在
        if (hasUpper) variation |= MtefRecord.TV_BO_UPPER;  // bit5：上限存在
        writeTemplateHeader(out, MtefRecord.TM_SUM, variation, 0x00);
    }

    /**
     * 写入积分模板（TM_INTEG）的头部。
     * <p>生成 ∫ 积分符号结构，可选地带有上限和/或下限。
     * variation 的基础值为 TV_INT_1（单重积分），在此基础上
     * 通过 TV_BO_LOWER/TV_BO_UPPER 控制积分限的存在性。
     * 调用方写入头部后需写入对应数量的槽位：
     * 被积表达式 LINE + [下限 LINE] + [上限 LINE]。</p>
     *
     * @param out      输出字节流
     * @param hasLower 是否存在积分下限
     * @param hasUpper 是否存在积分上限
     * @throws IOException 写入异常
     */
    public static void writeIntegralHeader(ByteArrayOutputStream out, boolean hasLower, boolean hasUpper) throws IOException {
        writeIntegralHeader(out, "\\int", hasLower, hasUpper);
    }

    public static void writeIntegralHeader(ByteArrayOutputStream out, String integralCommand,
                                           boolean hasLower, boolean hasUpper) throws IOException {
        int variation = switch (integralCommand) {
            case "\\iint" -> MtefRecord.TV_INT_2;
            case "\\iiint" -> MtefRecord.TV_INT_3;
            case "\\oint" -> MtefRecord.TV_INT_LOOP;
            default -> MtefRecord.TV_INT_1;
        };
        if (hasLower) variation |= MtefRecord.TV_BO_LOWER;   // bit4：下限存在
        if (hasUpper) variation |= MtefRecord.TV_BO_UPPER;   // bit5：上限存在
        writeTemplateHeader(out, MtefRecord.TM_INTEG, variation, 0x00);
    }

    /**
     * 写入连乘模板（TM_PROD）的头部。
     * <p>生成 ∏ 连乘符号结构，可选地带有上限和/或下限。
     * 调用方写入头部后需写入对应数量的槽位：
     * 操作数 LINE + [下限 LINE] + [上限 LINE]。</p>
     *
     * @param out      输出字节流
     * @param hasLower 是否存在下限
     * @param hasUpper 是否存在上限
     * @throws IOException 写入异常
     */
    public static void writeProductHeader(ByteArrayOutputStream out, boolean hasLower, boolean hasUpper) throws IOException {
        int variation = 0;
        if (hasLower) variation |= MtefRecord.TV_BO_LOWER;  // bit4：下限存在
        if (hasUpper) variation |= MtefRecord.TV_BO_UPPER;  // bit5：上限存在
        writeTemplateHeader(out, MtefRecord.TM_PROD, variation, 0x00);
    }

    public static void writeLimitHeader(ByteArrayOutputStream out, boolean hasLower, boolean hasUpper) throws IOException {
        int variation = MtefRecord.TV_BO_SUM;
        if (hasLower) variation |= MtefRecord.TV_BO_LOWER;
        if (hasUpper) variation |= MtefRecord.TV_BO_UPPER;
        writeTemplateHeader(out, MtefRecord.TM_LIM, variation, 0x00);
    }

    // ===== 长除法模板（Long Division）=====

    /**
     * 写入长除法模板（TM_LDIV）的头部。
     *
     * <p>根据 Wiris 的 MTEF 文档，LDivBoxClass 的模板子对象顺序为：</p>
     * <ul>
     *   <li>dividend slot</li>
     *   <li>quotient slot（可选）</li>
     * </ul>
     *
     * <p>除数不在模板槽位内，而是写在模板之前的父级对象列表中。</p>
     *
     * @param out 输出字节流
     * @param hasUpper 是否包含商槽位
     * @throws IOException 写入异常
     */
    public static void writeLongDivisionHeader(ByteArrayOutputStream out, boolean hasUpper) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_LDIV, hasUpper ? 0x0001 : 0x0000, 0x00);
    }

    // ===== 包围/删除线模板（Enclosures）=====

    public static void writeStrikeHeader(ByteArrayOutputStream out, int variation) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_STRIKE, variation, 0x00);
    }

    public static void writeBoxHeader(ByteArrayOutputStream out) throws IOException {
        writeBoxHeader(out, MtefRecord.TV_BX_LEFT
            | MtefRecord.TV_BX_RIGHT
            | MtefRecord.TV_BX_TOP
            | MtefRecord.TV_BX_BOTTOM);
    }

    public static void writeBoxHeader(ByteArrayOutputStream out, int variation) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_BOX, variation, 0x00);
    }

    // ===== 围栏/括号模板（Fences）=====

    /**
     * 写入圆括号模板（TM_PAREN）的头部。
     * <p>生成 ( ... ) 圆括号围栏结构，左右括号均显示。
     * variation = TV_FENCE_L | TV_FENCE_R 表示同时显示左括号和右括号。
     * 调用方写入头部后需写入 1 个槽位：括号内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeParenHeader(ByteArrayOutputStream out) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_PAREN, true, true);
    }

    /**
     * 写入方括号模板（TM_BRACK）的头部。
     * <p>生成 [ ... ] 方括号围栏结构，左右括号均显示。
     * 调用方写入头部后需写入 1 个槽位：括号内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeBracketHeader(ByteArrayOutputStream out) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_BRACK, true, true);
    }

    /**
     * 写入花括号模板（TM_BRACE）的头部。
     * <p>生成 { ... } 花括号围栏结构，左右括号均显示。
     * 调用方写入头部后需写入 1 个槽位：括号内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeBraceHeader(ByteArrayOutputStream out) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_BRACE, true, true);
    }

    /**
     * 写入竖线围栏模板（TM_BAR）的头部。
     * <p>生成 | ... | 单竖线围栏结构（绝对值），左右竖线均显示。
     * 调用方写入头部后需写入 1 个槽位：括号内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeBarHeader(ByteArrayOutputStream out) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_BAR, true, true);
    }

    public static void writeFenceHeader(ByteArrayOutputStream out, int selector,
                                        boolean hasLeft, boolean hasRight) throws IOException {
        int variation = 0;
        if (hasLeft) variation |= MtefRecord.TV_FENCE_L;
        if (hasRight) variation |= MtefRecord.TV_FENCE_R;
        writeTemplateHeader(out, selector, variation, 0x00);
    }

    public static void writeDoubleBarHeader(ByteArrayOutputStream out, boolean hasLeft, boolean hasRight) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_DBAR, hasLeft, hasRight);
    }

    public static void writeFloorHeader(ByteArrayOutputStream out, boolean hasLeft, boolean hasRight) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_FLOOR, hasLeft, hasRight);
    }

    public static void writeCeilingHeader(ByteArrayOutputStream out, boolean hasLeft, boolean hasRight) throws IOException {
        writeFenceHeader(out, MtefRecord.TM_CEILING, hasLeft, hasRight);
    }

    // ===== 上划线/下划线及装饰模板（Embellishments）=====

    /**
     * 写入上划线模板（TM_OBAR）的头部。
     * <p>生成在内容上方添加横线的结构（如 x̄）。
     * 调用方写入头部后需写入 1 个槽位：被划线内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeOverlineHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_OBAR, 0x00, 0x00);
    }

    /**
     * 写入下划线模板（TM_UBAR）的头部。
     * <p>生成在内容下方添加横线的结构。
     * 调用方写入头部后需写入 1 个槽位：被划线内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeUnderlineHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_UBAR, 0x00, 0x00);
    }

    /**
     * 写入向量箭头模板（TM_VEC）的头部。
     * <p>生成字符上方的右向箭头（→），用于表示向量。
     * variation = TV_VE_RIGHT 表示箭头指向右方。
     * 调用方写入头部后需写入 1 个槽位：向量内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeVecHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_VEC, MtefRecord.TV_VE_RIGHT, 0x00);
    }

    /**
     * 写入帽子模板（TM_HAT）的头部。
     * <p>生成字符上方的尖帽符号（^），如 x̂。
     * 调用方写入头部后需写入 1 个槽位：被修饰内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeHatHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_HAT, 0x00, 0x00);
    }

    /**
     * 写入弧线模板（TM_ARC）的头部。
     * <p>生成内容上方的弧线（如 \overarc{AB}）。
     * HatBoxClass 模板只包含一个主 slot，由 MathType 自身负责绘制可伸缩弧线。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeArcHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_ARC, 0x00, 0x00);
    }

    /**
     * 写入波浪号模板（TM_TILDE）的头部。
     * <p>生成字符上方的波浪号（~），如 x̃。
     * 调用方写入头部后需写入 1 个槽位：被修饰内容 LINE。</p>
     *
     * @param out 输出字节流
     * @throws IOException 写入异常
     */
    public static void writeTildeHeader(ByteArrayOutputStream out) throws IOException {
        writeTemplateHeader(out, MtefRecord.TM_TILDE, 0x00, 0x00);
    }

    /**
     * 写入水平大括号模板（TM_HBRACE）的头部。
     * <p>生成内容上方（overbrace，⏞）或内容下方（underbrace，⏟）的水平可拉伸大括号。
     * 调用方写入头部后需写入两个槽位：
     * <ol>
     *   <li>main slot：被括住的内容 LINE</li>
     *   <li>small slot：标注文字（下标/上标内容）LINE，使用 SUB 缩小字号</li>
     * </ol>
     * 之后需要写入可拉伸的大括号字符（使用 FN_EXPAND 字体），然后用 END 关闭模板。
     * @param out 输出字节流
     * @param onTop 是否将大括号放在内容上方（true = overbrace，false = underbrace）
     * @throws IOException 写入异常
     */
    public static void writeHBraceHeader(ByteArrayOutputStream out, boolean onTop) throws IOException {
        int variation = onTop ? MtefRecord.TV_HB_TOP : 0x0000;
        writeTemplateHeader(out, MtefRecord.TM_HBRACE, variation, 0x00);
    }

    /**
     * 写入水平方括号模板（TM_HBRACK）的头部。
     * <p>生成内容上方（overbracket，⎴）或内容下方（underbracket，⎵）的水平可拉伸方括号。</p>
     * @param out 输出字节流
     * @param onTop 是否将方括号放在内容上方
     * @throws IOException 写入异常
     */
    public static void writeHBrackHeader(ByteArrayOutputStream out, boolean onTop) throws IOException {
        int variation = onTop ? MtefRecord.TV_HB_TOP : 0x0000;
        writeTemplateHeader(out, MtefRecord.TM_HBRACK, variation, 0x00);
    }
}
