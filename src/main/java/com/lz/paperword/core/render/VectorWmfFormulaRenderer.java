package com.lz.paperword.core.render;

import com.lz.paperword.core.latex.LaTeXParser;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Minimal vector WMF formula preview writer.
 *
 * <p>This intentionally starts with flat formulas only. Unsupported LaTeX keeps
 * using the existing TeX/DIB path until the vector layout engine grows real
 * fraction, script, radical, and matrix support.</p>
 */
final class VectorWmfFormulaRenderer {

    private static final int PLACEABLE_WMF_KEY = 0x9AC6CDD7;
    private static final int TWIPS_PER_POINT = 20;
    private static final int MM_ANISOTROPIC = 8;
    private static final int TRANSPARENT = 1;
    private static final int TA_BASELINE = 0x0018;
    private static final int ANSI_CHARSET = 0;
    private static final int SYMBOL_CHARSET = 2;
    private static final int GB2312_CHARSET = 134;
    private static final int PEN_OBJECT_INDEX = 9;
    private static final double LEFT_MARGIN_PT = 0.12d;
    private static final double TOP_MARGIN_PT = 0.20d;
    private static final String TEXT_WIDTH_SCALE_PROP = "paperword.wmf.textWidth.scale";
    private static final Charset GBK = Charset.forName("GBK");
    private static final Charset WINDOWS_1252 = Charset.forName("windows-1252");
    private static final FontRenderContext FONT_RENDER_CONTEXT = new FontRenderContext(new AffineTransform(), true, true);
    private static final String ANSI_PREVIEW_FACE = "Times New Roman";
    private static final double COMPACT_FRACTION_SCALE = 0.78d;
    private static final double COMPACT_FRACTION_LAYOUT_WIDTH_SCALE = 0.965d;
    private static final double COMPACT_FRACTION_NUMERATOR_Y_PT = -0.2d;
    private static final double COMPACT_FRACTION_DENOMINATOR_Y_PT = 14.0d;
    private static final double COMPACT_FRACTION_BAR_Y_PT = 13.6d;
    private static final double COMPACT_FRACTION_HEIGHT_PT = 25.5d;
    private static final double FRACTION_NUMERATOR_Y_PT = -2.4d;
    private static final double FRACTION_DENOMINATOR_Y_PT = 11.8d;
    private static final double FRACTION_BAR_Y_PT = 12.8d;
    private static final double FRACTION_HEIGHT_PT = 26.2d;
    private static final Font TIMES_FONT = new Font(ANSI_PREVIEW_FACE, Font.PLAIN, 12);
    private static final Font SYMBOL_FONT = new Font("Symbol", Font.PLAIN, 12);
    private static final Font CJK_FONT = new Font("SimSun", Font.PLAIN, 12);
    private static final double SHORT_GEOMETRY_LABEL_WIDTH_SCALE = 0.90d;
    private static final double SIMPLE_LINEAR_FONT_Y_SCALE = 0.92d;
    private static final double SHORT_LINEAR_EQUATION_FONT_Y_SCALE = 1.076d;
    private static final double SHORT_LINEAR_EQUATION_WIDTH_SCALE = 1.08d;
    private static final double FRACTION_SCRIPT_CHAIN_WIDTH_SCALE = 0.98d;
    private static final double SCRIPT_FONT_HEIGHT_SCALE = 0.90d;
    private static final double SCRIPT_GLYPH_WIDTH_SCALE = 0.84d;
    private static final double STANDALONE_TWO_DIGIT_WIDTH_SCALE = 0.79d;
    private static final double STANDALONE_SINGLE_S_WIDTH_SCALE = 1.065d;
    private static final double STANDALONE_BD_WIDTH_SCALE = 1.135d;
    private static final double STANDALONE_PAREN_POWER_WIDTH_SCALE = 0.924d;
    private static final double EQUATION_PAREN_POWER_WIDTH_SCALE = 0.908d;
    private static final double STANDALONE_UPPER_SUBSCRIPT_FONT_Y_SCALE = 0.956d;
    private static final double STANDALONE_UPPER_SUBSCRIPT_WIDTH_SCALE = 0.814d;
    private static final double STANDALONE_S1_SUBSCRIPT_WIDTH_SCALE = 0.790d;
    private static final double STANDALONE_S3_SUBSCRIPT_WIDTH_SCALE = 0.842d;
    private static final double STANDALONE_LOWER_SUPERSCRIPT_WIDTH_SCALE = 0.86d;
    private static final double STANDALONE_B_SUPERSCRIPT_WIDTH_SCALE = 0.779d;
    private static final double SHORT_SCRIPT_EQUATION_WIDTH_SCALE = 0.97d;
    private static final double SHORT_S2_EQUALS_TWO_WIDTH_SCALE = 0.969d;
    private static final double SHORT_A_EQUALS_ONE_WIDTH_SCALE = 1.025d;
    private static final double SHORT_B_EQUALS_TWO_WIDTH_SCALE = 0.965d;
    private static final double PAREN_POWER_FONT_Y_SCALE = 1.019d;
    private static final double SHORT_SCRIPT_EQUATION_FONT_Y_SCALE = 0.955d;
    private static final double SCRIPT_RELATION_WIDTH_SCALE = 0.975d;
    private static final double SCRIPT_RELATION_SIMPLE_RATIO_WIDTH_SCALE = 0.970d;
    private static final double SCRIPT_RELATION_LONG_CHAIN_WIDTH_SCALE = 0.9895d;
    private static final double SCRIPT_RELATION_TRIANGLE_WIDTH_SCALE = 0.98d;
    private static final double SCRIPT_RELATION_LONG_TRIANGLE_WIDTH_SCALE = 0.981d;
    private static final double SCRIPT_RELATION_FONT_Y_SCALE = 0.93d;
    private static final double SCRIPT_RELATION_COMPACT_FONT_Y_SCALE = 0.890d;
    private static final double SCRIPT_RELATION_LONG_CHAIN_FONT_Y_SCALE = 0.890d;
    private static final double SCRIPT_RELATION_TRIANGLE_FONT_Y_SCALE = 0.887d;
    private static final double SCRIPT_RELATION_LONG_TRIANGLE_FONT_Y_SCALE = 0.955d;
    private static final double SCRIPT_RELATION_EXACT_S1_EQUATION_FONT_Y_SCALE = 0.887d;
    private static final double SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION = 1.113d;
    private static final double SCRIPT_RELATION_TRIANGLE_WIDTH_COMPENSATION = 1.025d;
    private static final double SCRIPT_RELATION_LONG_TRIANGLE_WIDTH_COMPENSATION = 1.016d;
    private static final double SCRIPT_RELATION_LONG_CHAIN_WIDTH_COMPENSATION = 1.022d;
    private static final Pattern LEFT_RIGHT_PAREN = Pattern.compile(
        "\\\\left\\s*\\(\\s*(?:\\{\\s*)?(.*?)(?:\\s*})?\\s*\\\\right\\s*\\)"
    );
    private static final Pattern ARRAY_PATTERN = Pattern.compile(
        "\\\\begin\\{array}\\{[^}]*}\\s*(.*?)\\s*\\\\end\\{array}",
        Pattern.DOTALL
    );
    private static final Pattern LEFT_BRACE_ARRAY_PATTERN = Pattern.compile(
        "\\\\left\\s*\\\\\\{\\s*\\\\begin\\{array}\\{[^}]*}\\s*(.*?)\\s*\\\\end\\{array}\\s*\\\\right\\s*\\.",
        Pattern.DOTALL
    );
    private static final Pattern PAREN_ARRAY_PATTERN = Pattern.compile(
        "\\\\left\\s*\\(\\s*(?:\\{\\s*)?\\\\begin\\{array}\\{[^}]*}\\s*(.*?)\\s*\\\\end\\{array}\\s*(?:}\\s*)?\\\\right\\s*\\)",
        Pattern.DOTALL
    );
    private static final Pattern TEXT_COMMAND_PATTERN = Pattern.compile(
        "\\\\(?:mathrm|mathbf|mathit|textit|textbf|emph|text|boldsymbol)\\s*\\{\\s*([^{}]*)\\s*}"
    );

    private VectorWmfFormulaRenderer() {
    }

    static boolean canRender(String latex) {
        return layout(latex) != null;
    }

    static byte[] render(String latex, double widthPt, double heightPt) throws IOException {
        FormulaLayout layout = layout(latex);
        if (layout == null || widthPt <= 0d || heightPt <= 0d) {
            return null;
        }
        return renderLayout(latex, layout, widthPt, heightPt);
    }

    static byte[] renderFallbackText(String latex, double widthPt, double heightPt) throws IOException {
        String fallback = fallbackPlainText(latex);
        FormulaLayout layout = layoutFlatRuns(tokenizePlainFallback(fallback));
        return renderLayout(fallback, layout, widthPt, heightPt);
    }

    private static byte[] renderLayout(String latex, FormulaLayout layout, double widthPt, double heightPt)
        throws IOException {
        PreviewScale previewScale = previewScale(latex, layout, widthPt, heightPt);
        double offsetX = Math.max(0.0d, Math.min(2.0d, (widthPt - layout.widthPt() * previewScale.x()) / 2.0d));
        double offsetY = Math.max(0.0d, Math.min(1.5d, (heightPt - layout.heightPt() * previewScale.y()) / 2.0d));
        double shortScriptHeightScale = simpleShortScriptScale(latex);
        double shortScriptWidthScale = simpleShortScriptWidthScale(latex);
        double standaloneDigitWidthScale = standaloneDigitWidthScale(latex);
        double standaloneSingleLetterWidthScale = standaloneSingleLetterWidthScale(latex);
        double standaloneGeometryWidthScale = standaloneGeometryWidthScale(latex);
        double shortScriptEquationWidthScale = shortScriptEquationWidthScale(latex);
        double shortExactEquationWidthScale = shortExactEquationWidthScale(latex);
        double shortLinearEquationWidthScale = shortLinearEquationWidthScale(latex, heightPt);
        double fontYScale = fontYScale(latex);
        double relationFontYScale = scriptRelationFontYScale(latex);
        double relationHeightWidthCompensation = scriptRelationHeightWidthCompensation(latex);
        double regularFontWidthScale = standaloneSingleLetterWidthScale > 1.0d ? standaloneSingleLetterWidthScale
            : 1.0d;
        WmfBuilder builder = new WmfBuilder(widthPt, heightPt);
        builder.record(0x0103, out -> writeWord(out, MM_ANISOTROPIC)); // SetMapMode
        builder.record(0x0106, out -> writeWord(out, TRANSPARENT)); // SetBkMode
        builder.record(0x012E, out -> writeWord(out, TA_BASELINE)); // SetTextAlign
        builder.record(0x020B, out -> {
            writeShort(out, 0);
            writeShort(out, 0);
        }); // SetWindowOrg
        builder.record(0x020C, out -> {
            writeShort(out, toTwips(heightPt));
            writeShort(out, toTwips(widthPt));
        }); // SetWindowExt
        builder.record(0x0209, out -> {
            writeWord(out, 0);
            writeWord(out, 0);
            writeWord(out, 0);
        }); // SetTextColor black
        builder.record(0x02FB, out -> writeFont(out, ANSI_PREVIEW_FACE,
            12.0d * fontYScale * relationFontYScale * previewScale.y(), ANSI_CHARSET, regularFontWidthScale));
        builder.record(0x02FB, out -> writeFont(out, ANSI_PREVIEW_FACE,
            8.0d * shortScriptHeightScale * fontYScale * relationFontYScale * previewScale.y(), ANSI_CHARSET,
            shortScriptWidthScale));
        builder.record(0x02FB, out -> writeFont(out, ANSI_PREVIEW_FACE,
            22.0d * previewScale.y(), ANSI_CHARSET));
        builder.record(0x02FB, out -> writeFont(out, "Symbol",
            12.0d * fontYScale * relationFontYScale * previewScale.y(), SYMBOL_CHARSET, regularFontWidthScale));
        builder.record(0x02FB, out -> writeFont(out, "Symbol",
            8.0d * shortScriptHeightScale * fontYScale * relationFontYScale * previewScale.y(), SYMBOL_CHARSET,
            shortScriptWidthScale));
        builder.record(0x02FB, out -> writeFont(out, "Symbol", 22.0d * previewScale.y(), SYMBOL_CHARSET));
        builder.record(0x02FB, out -> writeFont(out, "SimSun",
            12.0d * fontYScale * relationFontYScale * previewScale.y(), GB2312_CHARSET, regularFontWidthScale));
        builder.record(0x02FB, out -> writeFont(out, "SimSun",
            8.0d * shortScriptHeightScale * fontYScale * relationFontYScale * previewScale.y(), GB2312_CHARSET,
            shortScriptWidthScale));
        builder.record(0x02FB, out -> writeFont(out, "SimSun", 22.0d * previewScale.y(), GB2312_CHARSET));
        builder.record(0x02FA, VectorWmfFormulaRenderer::writeBlackPen); // CreatePen
        builder.record(0x012D, out -> writeWord(out, 0)); // SelectObject

        if (!layout.lines().isEmpty()) {
            builder.record(0x012D, out -> writeWord(out, PEN_OBJECT_INDEX)); // SelectObject pen
            for (LineSegment line : layout.lines()) {
                int x1 = toTwips(LEFT_MARGIN_PT + offsetX + line.x1Pt() * previewScale.x());
                int y1 = toTwips(TOP_MARGIN_PT + offsetY + line.y1Pt() * previewScale.y());
                int x2 = toTwips(LEFT_MARGIN_PT + offsetX + line.x2Pt() * previewScale.x());
                int y2 = toTwips(TOP_MARGIN_PT + offsetY + line.y2Pt() * previewScale.y());
                builder.record(0x0325, out -> writePolyline(out, x1, y1, x2, y2));
            }
            builder.record(0x012D, out -> writeWord(out, 0));
        }
        int selectedFont = 0;
        for (PlacedText run : layout.runs()) {
            final int runBaseline = toTwips(TOP_MARGIN_PT + offsetY + run.baselinePt() * previewScale.y());
            double segmentX = run.xPt();
            for (EncodedText segment : encodeText(run.text(), run.cjk())) {
                int fontIndex = fontIndex(segment.kind(), run.script(), run.display());
                if (fontIndex != selectedFont) {
                    final int selectFontIndex = fontIndex;
                    builder.record(0x012D, out -> writeWord(out, selectFontIndex));
                    selectedFont = fontIndex;
                }
                final int runX = toTwips(LEFT_MARGIN_PT + offsetX + segmentX * previewScale.x());
                byte[] bytes = segment.bytes();
                double runWidthScale = run.widthScale() * standaloneDigitWidthScale * standaloneSingleLetterWidthScale
                    * standaloneGeometryWidthScale * shortScriptEquationWidthScale * shortExactEquationWidthScale
                    * shortLinearEquationWidthScale * relationHeightWidthCompensation;
                int[] dx = characterDxTwips(segment, run.script(), run.display(), previewScale.x(), shortScriptWidthScale,
                    runWidthScale);
                builder.record(0x0A32, out -> writeExtTextOut(out, runX, runBaseline, bytes, dx));
                segmentX += segment.widthPt() * sizeScale(run.script(), run.display())
                    * (run.script() ? shortScriptWidthScale : 1.0d) * runWidthScale;
            }
        }
        return builder.finish();
    }

    private static List<TextRun> tokenizePlainFallback(String text) {
        List<TextRun> out = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        boolean currentCjk = false;
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            boolean cjk = isCjk(String.valueOf(ch));
            if (!current.isEmpty() && cjk != currentCjk) {
                out.add(new TextRun(current.toString(), currentCjk));
                current.setLength(0);
            }
            current.append(ch);
            currentCjk = cjk;
        }
        if (!current.isEmpty()) {
            out.add(new TextRun(current.toString(), currentCjk));
        }
        return out.isEmpty() ? List.of(new TextRun(" ", false)) : out;
    }

    private static String fallbackPlainText(String latex) {
        String text = stripMetricsAndStyles(latex == null ? "" : latex);
        text = normalizeTextCommands(text);
        text = text.replaceAll("\\\\sqrt\\s*\\[[^]]*]\\s*\\{([^{}]*)}", "sqrt($1)");
        text = text.replaceAll("\\\\sqrt\\s*\\{([^{}]*)}", "sqrt($1)");
        String previous;
        do {
            previous = text;
            text = text.replaceAll("\\\\frac\\s*\\{([^{}]*)}\\s*\\{([^{}]*)}", "($1)/($2)");
        } while (!text.equals(previous));
        text = text.replaceAll("\\\\sum\\s*(_\\{[^{}]*})?\\s*(\\^\\{[^{}]*})?", "sum");
        text = text.replaceAll("\\\\begin\\{array}\\{[^}]*}", "");
        text = text.replaceAll("\\\\end\\{array}", "");
        text = text.replace("\\\\", " ");
        text = replaceCommandWords(text);
        text = text.replaceAll("[{}]", "");
        text = text.replaceAll("\\s+", " ").trim();
        return text.isBlank() ? " " : text;
    }

    private static String replaceCommandWords(String text) {
        return text
            .replace("\\alpha", "α")
            .replace("\\beta", "β")
            .replace("\\gamma", "γ")
            .replace("\\delta", "δ")
            .replace("\\pi", "π")
            .replace("\\times", "×")
            .replace("\\div", "÷")
            .replace("\\leq", "≤")
            .replace("\\le", "≤")
            .replace("\\geq", "≥")
            .replace("\\ge", "≥")
            .replace("\\neq", "≠")
            .replace("\\ne", "≠")
            .replace("\\cdots", "...")
            .replace("\\ldots", "...")
            .replace("\\pm", "±")
            .replaceAll("\\\\[A-Za-z]+", "");
    }

    private static PreviewScale previewScale(String latex, FormulaLayout layout, double widthPt, double heightPt) {
        double drawableWidthPt = Math.max(widthPt - LEFT_MARGIN_PT * 2.0d, 0.1d);
        double drawableHeightPt = Math.max(heightPt - TOP_MARGIN_PT * 2.0d, 0.1d);
        double widthScale = drawableWidthPt / Math.max(layout.widthPt() * textWidthScale(), 1.0d);
        double heightScale = drawableHeightPt / Math.max(layoutVerticalExtentPt(layout), 1.0d);
        boolean gridLike = isGridLikeLatex(latex);
        double xScale = Math.max(0.05d, Math.min(gridLike ? 8.0d : 2.25d, widthScale));
        double yScale = Math.max(0.05d, Math.min(1.0d, heightScale));
        if (!gridLike && heightScale < 0.98d) {
            xScale = Math.min(xScale, Math.max(0.05d, widthScale * heightScale));
        }
        return new PreviewScale(xScale, yScale);
    }

    private static double simpleShortScriptScale(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex))).trim();
        if (text.isEmpty() || hasFractionCommand(text) || text.contains("\\sqrt") || text.contains("\\begin")
            || text.contains("\\left") || text.contains("\\right") || text.contains("\\over")
            || text.contains("\\under") || text.contains("\\boxed")) {
            return 1.0d;
        }
        int firstOp = findFirstTopLevelScriptOperator(text);
        if (firstOp <= 0) {
            return 1.0d;
        }
        int atomEnd = skipWhitespaceBackward(text, firstOp);
        int atomStart = findAtomStart(text, atomEnd);
        if (atomStart != 0 || atomStart >= atomEnd) {
            return 1.0d;
        }
        ScriptGroup first = readScriptGroup(text, firstOp);
        if (first == null) {
            return 1.0d;
        }
        int cursor = skipWhitespaceForward(text, first.end());
        if (cursor >= text.length()) {
            return shortScriptAtom(text.substring(atomStart, atomEnd), first.body()) ? SCRIPT_FONT_HEIGHT_SCALE : 1.0d;
        }
        if (text.charAt(cursor) != '^' && text.charAt(cursor) != '_') {
            return 1.0d;
        }
        if (text.charAt(cursor) == text.charAt(firstOp)) {
            return 1.0d;
        }
        ScriptGroup second = readScriptGroup(text, cursor);
        if (second == null || skipWhitespaceForward(text, second.end()) != text.length()) {
            return 1.0d;
        }
        return shortScriptAtom(text.substring(atomStart, atomEnd), first.body())
            && shortScriptAtom(text.substring(atomStart, atomEnd), second.body()) ? SCRIPT_FONT_HEIGHT_SCALE : 1.0d;
    }

    private static double simpleShortScriptWidthScale(String latex) {
        return simpleShortScriptScale(latex) < 1.0d ? SCRIPT_GLYPH_WIDTH_SCALE : 1.0d;
    }

    private static double standaloneDigitWidthScale(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex))).trim();
        return text.matches("\\d{2}") ? STANDALONE_TWO_DIGIT_WIDTH_SCALE : 1.0d;
    }

    private static double standaloneSingleLetterWidthScale(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex))).trim();
        return "S".equals(text) ? STANDALONE_SINGLE_S_WIDTH_SCALE : 1.0d;
    }

    private static double standaloneGeometryWidthScale(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex))).trim();
        return "BD".equals(text) ? STANDALONE_BD_WIDTH_SCALE : 1.0d;
    }

    static boolean isStandaloneTallGeometryLabel(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex))).trim();
        return "AB".equals(text) || "BD".equals(text);
    }

    private static double standaloneParenPowerWidthScale(String latex) {
        String text = normalizeTextCommands(stripMetricsAndStyles(latex == null ? " " : latex))
            .replaceAll("\\s+", "");
        return text.matches("\\\\left\\([^()=]+\\\\right\\)\\^\\{?[A-Za-z0-9]{1,2}}?")
            ? STANDALONE_PAREN_POWER_WIDTH_SCALE : 1.0d;
    }

    private static double standaloneShortScriptWidthScale(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex)))
            .replaceAll("\\s+", "");
        if (text.isEmpty() || hasFractionCommand(text) || text.contains("\\sqrt") || text.contains("\\begin")
            || text.contains("\\left") || text.contains("\\right") || text.contains("\\over")
            || text.contains("\\under") || text.contains("\\boxed") || text.contains("=")
            || text.contains("\\colon") || text.contains("+") || text.contains("-")
            || text.contains("\\times") || text.contains("\\div")) {
            return 1.0d;
        }
        int op = findFirstTopLevelScriptOperator(text);
        if (op <= 0 || skipWhitespaceBackward(text, op) != op || findAtomStart(text, op) != 0) {
            return 1.0d;
        }
        ScriptGroup first = readScriptGroup(text, op);
        if (first == null || !shortScriptAtom(text.substring(0, op), first.body())) {
            return 1.0d;
        }
        int cursor = skipWhitespaceForward(text, first.end());
        if (cursor >= text.length()) {
            String base = text.substring(0, op);
            String body = first.body();
            if (first.operator() == '_' && base.matches("[A-Z]") && body.matches("\\d{1,2}")) {
                if ("S".equals(base) && "1".equals(body)) {
                    return STANDALONE_S1_SUBSCRIPT_WIDTH_SCALE;
                }
                if ("S".equals(base) && "3".equals(body)) {
                    return STANDALONE_S3_SUBSCRIPT_WIDTH_SCALE;
                }
                return STANDALONE_UPPER_SUBSCRIPT_WIDTH_SCALE;
            }
            if (first.operator() == '^' && base.matches("[a-z]") && body.matches("\\d{1,2}")) {
                return "b".equals(base) && "2".equals(body) ? STANDALONE_B_SUPERSCRIPT_WIDTH_SCALE
                    : STANDALONE_LOWER_SUPERSCRIPT_WIDTH_SCALE;
            }
            return 0.88d;
        }
        if ((text.charAt(cursor) != '^' && text.charAt(cursor) != '_') || text.charAt(cursor) == first.operator()) {
            return 1.0d;
        }
        ScriptGroup second = readScriptGroup(text, cursor);
        if (second == null || skipWhitespaceForward(text, second.end()) != text.length()) {
            return 1.0d;
        }
        return shortScriptAtom(text.substring(0, op), second.body()) ? 0.88d : 1.0d;
    }

    static double shortScriptEquationWidthScale(String latex) {
        return isShortScriptEquation(latex) ? SHORT_SCRIPT_EQUATION_WIDTH_SCALE : 1.0d;
    }

    static double shortExactEquationWidthScale(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex)))
            .replaceAll("\\s+", "");
        if ("a=1".equals(text)) {
            return SHORT_A_EQUALS_ONE_WIDTH_SCALE;
        }
        if ("S_{2}=2".equals(text)) {
            return SHORT_S2_EQUALS_TWO_WIDTH_SCALE;
        }
        return "b=2".equals(text) ? SHORT_B_EQUALS_TWO_WIDTH_SCALE : 1.0d;
    }

    static double scriptRelationWidthScale(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        if (text.isEmpty() || hasFractionCommand(raw) || raw.contains("\\sqrt") || raw.contains("\\begin")
            || raw.contains("\\left") || raw.contains("\\right") || raw.contains("\\over")
            || raw.contains("\\under") || raw.contains("\\boxed") || findFirstTopLevelScriptOperator(text) <= 0) {
            return 1.0d;
        }
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        if (relationOperators == 0 || standaloneShortScriptWidthScale(latex) < 0.999d) {
            return 1.0d;
        }
        if (triangleAreaScriptTermCount(text) >= 2 && relationOperators >= 4) {
            return isLongTriangleScriptRelation(raw) ? SCRIPT_RELATION_LONG_TRIANGLE_WIDTH_SCALE
                : SCRIPT_RELATION_TRIANGLE_WIDTH_SCALE;
        }
        if (isSimpleScriptRatioRelation(raw, text, relationOperators)) {
            return SCRIPT_RELATION_SIMPLE_RATIO_WIDTH_SCALE;
        }
        return relationOperators >= 6 ? SCRIPT_RELATION_LONG_CHAIN_WIDTH_SCALE : SCRIPT_RELATION_WIDTH_SCALE;
    }

    private static boolean isSimpleScriptRatioRelation(String raw, String text, int relationOperators) {
        return relationOperators == 3 && !raw.contains("\\bigtriangleup")
            && !isExactS1AreaEquation(text) && (text.contains("\\colon") || text.contains(":")) && text.contains("=");
    }

    static double scriptRelationFontYScale(String latex) {
        if (isExactS1AreaEquation(latex)) {
            return SCRIPT_RELATION_EXACT_S1_EQUATION_FONT_Y_SCALE;
        }
        if (isTriangleScriptRelation(latex)) {
            return isLongTriangleScriptRelation(latex) ? SCRIPT_RELATION_LONG_TRIANGLE_FONT_Y_SCALE
                : SCRIPT_RELATION_TRIANGLE_FONT_Y_SCALE;
        }
        if (isCompactScriptRelation(latex)) {
            return SCRIPT_RELATION_COMPACT_FONT_Y_SCALE;
        }
        if (isLongChainScriptRelation(latex)) {
            return SCRIPT_RELATION_LONG_CHAIN_FONT_Y_SCALE;
        }
        return scriptRelationFontYEligible(latex) ? SCRIPT_RELATION_FONT_Y_SCALE : 1.0d;
    }

    private static boolean isCompactScriptRelation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        return relationOperators >= 2 && relationOperators <= 3 && !raw.contains("\\bigtriangleup")
            && !text.contains("\\times") && !text.contains("\\div") && !text.contains("+") && !text.contains("-")
            && scriptRelationFontYEligible(latex);
    }

    private static boolean isTriangleScriptRelation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        return triangleAreaScriptTermCount(text) >= 2 && relationOperators >= 4 && scriptRelationFontYEligible(latex);
    }

    private static boolean isLongTriangleScriptRelation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        return triangleAreaScriptTermCount(text) >= 3 && relationOperators >= 8 && text.length() >= 120;
    }

    private static int triangleAreaScriptTermCount(String text) {
        int count = 0;
        int cursor = 0;
        while (cursor < text.length()) {
            int index = text.indexOf("S_{\\bigtriangleup", cursor);
            if (index < 0) {
                return count;
            }
            count++;
            cursor = index + 1;
        }
        return count;
    }

    private static boolean isLongChainScriptRelation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        return relationOperators >= 6 && triangleAreaScriptTermCount(text) < 2
            && !text.contains("\\times") && !text.contains("\\div") && !text.contains("+") && !text.contains("-")
            && scriptRelationFontYEligible(latex);
    }

    static double scriptRelationHeightWidthCompensation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        if (isExactS1AreaEquation(text)) {
            return SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION;
        }
        if (isLongChainScriptRelation(latex) && scriptRelationFontYScale(latex) < 0.999d) {
            return SCRIPT_RELATION_LONG_CHAIN_WIDTH_COMPENSATION;
        }
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        if (triangleAreaScriptTermCount(text) >= 2 && relationOperators >= 4 && scriptRelationFontYScale(latex) < 0.999d) {
            return isLongTriangleScriptRelation(latex) ? SCRIPT_RELATION_LONG_TRIANGLE_WIDTH_COMPENSATION
                : SCRIPT_RELATION_TRIANGLE_WIDTH_COMPENSATION;
        }
        return 1.0d;
    }

    private static boolean isExactS1AreaEquation(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex)))
            .replaceAll("\\s+", "");
        return "S_{1}=a^{2}=1".equals(text);
    }

    static boolean scriptRelationFontYEligible(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        if (scriptRelationWidthScale(latex) >= 0.999d) {
            return false;
        }
        int relationOperators = Math.max(topLevelRelationOperatorCount(normalizeFlatLatex(normalizeTextCommands(raw))),
            topLevelRelationOperatorCount(normalizeFlatLatex(raw)));
        if (relationOperators < 2) {
            return false;
        }
        return hasTopLevelEqualsOrColon(text) || hasTopLevelEqualsOrColon(normalizeFlatLatex(raw));
    }

    private static boolean hasTopLevelEqualsOrColon(String text) {
        int depth = 0;
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '\\') {
                int commandEnd = skipCommandName(text, i + 1);
                String command = text.substring(i + 1, commandEnd);
                if (depth == 0 && command.equals("colon")) {
                    return true;
                }
                i = commandEnd - 1;
                continue;
            }
            if (ch == '{') {
                depth++;
            } else if (ch == '}') {
                depth = Math.max(0, depth - 1);
            } else if (ch == '(' || ch == '[' || ch == '（' || ch == '【') {
                depth++;
            } else if (ch == ')' || ch == ']' || ch == '）' || ch == '】') {
                depth = Math.max(0, depth - 1);
            } else if (depth == 0 && (ch == '=' || ch == ':')) {
                return true;
            }
        }
        return false;
    }

    static boolean hasTopLevelRelationOperator(String text) {
        return topLevelRelationOperatorCount(text) > 0;
    }

    static int topLevelRelationOperatorCount(String text) {
        int depth = 0;
        int count = 0;
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '\\') {
                int commandEnd = skipCommandName(text, i + 1);
                String command = text.substring(i + 1, commandEnd);
                if (depth == 0 && (command.equals("colon") || command.equals("times") || command.equals("div"))) {
                    count++;
                }
                i = commandEnd - 1;
                continue;
            }
            if (ch == '{') {
                depth++;
            } else if (ch == '}') {
                depth = Math.max(0, depth - 1);
            } else if (ch == '(' || ch == '[' || ch == '（' || ch == '【') {
                depth++;
            } else if (ch == ')' || ch == ']' || ch == '）' || ch == '】') {
                depth = Math.max(0, depth - 1);
            } else if (depth == 0 && (ch == '=' || ch == ':' || ch == '+' || ch == '-')) {
                count++;
            }
        }
        return count;
    }

    private static double fontYScale(String latex) {
        String text = stripMetricsAndStyles(latex == null ? "" : latex);
        if (hasClosingFenceSuperscript(latex) && !hasFractionCommand(text) && !text.contains("\\sqrt")
            && !text.contains("\\begin") && !text.contains("\\over") && !text.contains("\\under")
            && !text.contains("\\boxed")) {
            return PAREN_POWER_FONT_Y_SCALE;
        }
        if (text.contains("\\left") || text.contains("\\right") || hasFractionCommand(text)
            || text.contains("\\sqrt") || text.contains("\\begin") || text.contains("\\over")
            || text.contains("\\under") || text.contains("\\boxed") || isStandaloneTallGeometryLabel(latex)) {
            return 1.0d;
        }
        if (isStandaloneUpperSubscript(latex)) {
            return SIMPLE_LINEAR_FONT_Y_SCALE * STANDALONE_UPPER_SUBSCRIPT_FONT_Y_SCALE;
        }
        return isShortScriptEquation(latex) ? SIMPLE_LINEAR_FONT_Y_SCALE * SHORT_SCRIPT_EQUATION_FONT_Y_SCALE
            : isShortLinearEquation(latex) ? SIMPLE_LINEAR_FONT_Y_SCALE * SHORT_LINEAR_EQUATION_FONT_Y_SCALE
            : SIMPLE_LINEAR_FONT_Y_SCALE;
    }

    static boolean isShortLinearEquation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        if (raw.contains("\\left") || raw.contains("\\right") || hasFractionCommand(raw) || raw.contains("\\sqrt")
            || raw.contains("\\begin") || raw.contains("\\over") || raw.contains("\\under")
            || raw.contains("\\boxed") || raw.contains("^") || raw.contains("_")) {
            return false;
        }
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        if (text.matches("[A-Za-z]=\\d+(?:[+\\-]\\d+)+")) {
            return false;
        }
        boolean hasMultiplicativeOperator = text.contains("×") || text.contains("÷") || text.contains("\\times")
            || text.contains("\\div") || text.contains("\\cdot") || text.contains("\\spot");
        return text.length() >= 5 && text.length() <= 16
            && (text.matches(".*[A-Za-z].*") || hasMultiplicativeOperator)
            && (text.contains("=") || hasMultiplicativeOperator)
            && text.matches("[A-Za-z0-9]+(?:[=×÷+\\-]|\\\\times|\\\\div|\\\\cdot|\\\\spot)[A-Za-z0-9=×÷+\\-]+")
            && topLevelRelationOperatorCount(text) >= 1;
    }

    static double shortLinearEquationWidthScale(String latex) {
        return shortLinearEquationWidthScale(latex, 12.0d);
    }

    static double shortLinearEquationWidthScale(String latex, double heightPt) {
        return heightPt <= 12.1d && isShortLinearEquation(latex) ? SHORT_LINEAR_EQUATION_WIDTH_SCALE : 1.0d;
    }

    static double fractionScriptChainWidthScale(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        if (!hasFractionCommand(raw) || raw.length() <= 45 || hasNestedFraction(raw)) {
            return 1.0d;
        }
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        boolean hasScript = text.contains("_") || text.contains("^");
        boolean hasAreaToken = raw.contains("\\bigtriangleup") || raw.contains("\\Delta") || raw.contains("\\Updelta")
            || raw.contains("梯形") || raw.contains("三角形");
        return hasScript && hasAreaToken ? FRACTION_SCRIPT_CHAIN_WIDTH_SCALE : 1.0d;
    }

    static double compactInlineFractionWidthScale(String latex) {
        return isCompactInlineFraction(stripMetricsAndStyles(latex == null ? "" : latex))
            ? COMPACT_FRACTION_LAYOUT_WIDTH_SCALE : 1.0d;
    }

    static boolean hasClosingFenceSuperscript(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex)))
            .replaceAll("\\s+", "");
        int cursor = 0;
        while (cursor < text.length()) {
            int op = nextScriptOperator(text, cursor);
            if (op < 0) {
                return false;
            }
            int atomEnd = skipWhitespaceBackward(text, op);
            int atomStart = findAtomStart(text, atomEnd);
            ScriptGroup group = readScriptGroup(text, op);
            if (atomStart >= 0 && atomStart < atomEnd && group != null && group.operator() == '^'
                && isClosingFenceScriptBase(text.substring(atomStart, atomEnd))) {
                return true;
            }
            cursor = group == null ? op + 1 : group.end();
        }
        return false;
    }

    static boolean isStandaloneUpperSubscript(String latex) {
        String text = normalizeFlatLatex(normalizeTextCommands(stripMetricsAndStyles(latex == null ? "" : latex)))
            .replaceAll("\\s+", "");
        if (text.isEmpty() || hasFractionCommand(text) || text.contains("\\sqrt") || text.contains("\\begin")
            || text.contains("\\left") || text.contains("\\right") || text.contains("\\over")
            || text.contains("\\under") || text.contains("\\boxed") || text.contains("=")
            || text.contains("\\colon") || text.contains("+") || text.contains("-")
            || text.contains("\\times") || text.contains("\\div")) {
            return false;
        }
        int op = findFirstTopLevelScriptOperator(text);
        if (op <= 0 || text.charAt(op) != '_' || findAtomStart(text, op) != 0) {
            return false;
        }
        ScriptGroup script = readScriptGroup(text, op);
        return script != null && script.end() == text.length()
            && text.substring(0, op).matches("[A-Z]") && script.body().matches("\\d{1,2}");
    }

    static boolean isShortScriptEquation(String latex) {
        String raw = stripMetricsAndStyles(latex == null ? "" : latex);
        if (raw.isBlank() || hasFractionCommand(raw) || raw.contains("\\sqrt") || raw.contains("\\begin")
            || raw.contains("\\left") || raw.contains("\\right") || raw.contains("\\over")
            || raw.contains("\\under") || raw.contains("\\boxed")) {
            return false;
        }
        String text = normalizeFlatLatex(normalizeTextCommands(raw)).replaceAll("\\s+", "");
        int relationOperators = topLevelRelationOperatorCount(text);
        int equals = text.indexOf('=');
        if (relationOperators != 1 || equals <= 0 || text.indexOf('=', equals + 1) >= 0 || equals >= text.length() - 1) {
            return false;
        }
        String left = text.substring(0, equals);
        String right = text.substring(equals + 1);
        int op = findFirstTopLevelScriptOperator(left);
        if (op <= 0 || findFirstTopLevelScriptOperator(left.substring(op + 1)) >= 0
            || skipWhitespaceBackward(left, op) != op || findAtomStart(left, op) != 0) {
            return false;
        }
        ScriptGroup script = readScriptGroup(left, op);
        return script != null && script.end() == left.length() && shortScriptAtom(left.substring(0, op), script.body())
            && right.matches("[A-Za-z0-9.]{1,4}");
    }

    private static boolean hasFractionCommand(String text) {
        return text != null && (text.contains("\\frac") || text.contains("\\dfrac") || text.contains("\\cfrac"));
    }

    private static int findFirstTopLevelScriptOperator(String text) {
        int depth = 0;
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '\\') {
                i = skipCommandName(text, i + 1) - 1;
                continue;
            }
            if (ch == '{') {
                depth++;
            } else if (ch == '}') {
                depth = Math.max(0, depth - 1);
            } else if (depth == 0 && (ch == '^' || ch == '_')) {
                return i;
            }
        }
        return -1;
    }

    private static int skipCommandName(String text, int start) {
        int cursor = start;
        while (cursor < text.length() && Character.isLetter(text.charAt(cursor))) {
            cursor++;
        }
        return Math.max(cursor, start + 1);
    }

    private static boolean shortScriptAtom(String base, String script) {
        return visibleAtomLength(base) <= 2 && visibleAtomLength(script) <= 2;
    }

    private static int visibleAtomLength(String text) {
        String normalized = normalizeTextCommands(text == null ? "" : text)
            .replaceAll("\\\\[A-Za-z]+", "x")
            .replaceAll("[{}\\s]", "");
        return normalized.codePointCount(0, normalized.length());
    }

    private static double layoutVerticalExtentPt(FormulaLayout layout) {
        double bottom = layout.heightPt();
        for (PlacedText run : layout.runs()) {
            bottom = Math.max(bottom, run.baselinePt() + (run.script() ? 2.8d : 3.5d));
        }
        for (LineSegment line : layout.lines()) {
            bottom = Math.max(bottom, Math.max(line.y1Pt(), line.y2Pt()) + 0.6d);
        }
        return bottom;
    }

    private static boolean isGridLikeLatex(String latex) {
        return latex != null && (latex.contains("\\begin{array}") || latex.contains("\\begin{cases}"));
    }

    private static double textWidthScale() {
        String value = System.getProperty(TEXT_WIDTH_SCALE_PROP);
        if (value == null || value.isBlank()) {
            return 1.0d;
        }
        try {
            return Math.max(0.75d, Math.min(1.35d, Double.parseDouble(value)));
        } catch (NumberFormatException ignored) {
            return 1.08d;
        }
    }

    private static FormulaLayout layout(String latex) {
        if (latex == null) {
            return null;
        }
        String text = LaTeXParser.preNormalizeLatex(stripMetricsAndStyles(latex)).trim();
        if (text.isBlank()) {
            return new FormulaLayout(List.of(new PlacedText(" ", false, false, false, 0.0d, 9.6d)), List.of(), 4.0d, 13.0d);
        }
        text = normalizeSpacingCommands(text);
        text = normalizeHorizontalBraceAnnotations(text);
        text = normalizeOverUnderSetCommands(text);
        FormulaLayout underline = layoutUnderline(text);
        if (underline != null) {
            return underline;
        }
        text = normalizeSimpleDecorations(text);
        text = normalizeTextCommands(text);
        text = normalizeWrappedScriptBases(text);
        FormulaLayout cancel = layoutCancel(text);
        if (cancel != null) {
            return cancel;
        }
        FormulaLayout boxed = layoutBoxed(text);
        if (boxed != null) {
            return boxed;
        }
        FormulaLayout boxedFragments = layoutBoxedFragments(text);
        if (boxedFragments != null) {
            return boxedFragments;
        }
        FormulaLayout compositeCases = layoutCompositeCases(text);
        if (compositeCases != null) {
            return compositeCases;
        }
        FormulaLayout cases = layoutCases(text);
        if (cases != null) {
            return cases;
        }
        FormulaLayout overline = layoutOverline(text);
        if (overline != null) {
            return overline;
        }
        FormulaLayout overarc = layoutOverarc(text);
        if (overarc != null) {
            return overarc;
        }
        FormulaLayout arrows = layoutXArrowFragments(text);
        if (arrows != null) {
            return arrows;
        }
        FormulaLayout sqrt = layoutSqrt(text);
        if (sqrt != null) {
            return sqrt;
        }
        Matcher leftBraceArray = LEFT_BRACE_ARRAY_PATTERN.matcher(text);
        if (leftBraceArray.matches()) {
            return layoutLeftBraceArray(leftBraceArray.group(1));
        }
        Matcher parenArray = PAREN_ARRAY_PATTERN.matcher(text);
        if (parenArray.matches()) {
            return layoutParenArray(parenArray.group(1));
        }
        FormulaLayout compositeArray = layoutCompositeArrays(text);
        if (compositeArray != null) {
            return compositeArray;
        }
        ArraySlice array = readArraySlice(text, 0);
        if (array != null && array.end() == text.length()) {
            ArraySlice singleCases = unwrapSingleCasesArray(array.body());
            if (singleCases != null) {
                return layoutLeftBraceArray(singleCases.body());
            }
            return layoutArray(array.body());
        }
        FormulaLayout fractions = layoutFractions(text);
        if (fractions != null) {
            double fractionScale = fractionScriptChainWidthScale(latex);
            if (fractionScale >= 0.999d) {
                fractionScale = compactInlineFractionWidthScale(latex);
            }
            return scaleLayoutX(fractions, fractionScale);
        }
        String flatText = normalizeFlatLatex(text);
        FormulaLayout scripts = layoutScripts(flatText);
        if (scripts != null) {
            return scaleLayoutX(scripts,
                standaloneParenPowerWidthScale(latex) * standaloneShortScriptWidthScale(latex)
                    * scriptRelationWidthScale(latex));
        }
        FormulaLayout standaloneScript = layoutStandaloneScript(flatText);
        if (standaloneScript != null) {
            return standaloneScript;
        }
        FormulaLayout flattenedScripts = layoutFlattenedScripts(flatText);
        if (flattenedScripts != null) {
            return flattenedScripts;
        }
        List<TextRun> runs = tokenizeFlat(flatText);
        if (runs == null) {
            return null;
        }
        return layoutFlatRuns(runs);
    }

    private static List<TextRun> tokenize(String latex) {
        if (latex == null || latex.isBlank()) {
            return null;
        }
        String text = stripMetricsAndStyles(latex).trim();
        return tokenizeFlat(text);
    }

    private static List<TextRun> tokenizeFlat(String text) {
        text = normalizeFlatLatex(text);
        text = normalizeTextCommands(text);
        if (text.contains("\\begin") || text.contains("\\frac") || text.contains("\\sqrt")
            || text.contains("\\over") || text.contains("^")) {
            return null;
        }
        List<TextRun> out = new ArrayList<>();
        for (int i = 0; i < text.length();) {
            char ch = text.charAt(i);
            if (Character.isWhitespace(ch) || ch == '{' || ch == '}') {
                i++;
                continue;
            }
            if (ch == '\\') {
                Command command = readCommand(text, i);
                if (command == null) {
                    return null;
                }
                out.add(new TextRun(command.text(), isCjk(command.text())));
                i = command.end();
                continue;
            }
            int end = i + 1;
            boolean cjk = isCjk(String.valueOf(ch));
            if (!cjk) {
                while (end < text.length()) {
                    char next = text.charAt(end);
                    if (next == '\\' || Character.isWhitespace(next) || next == '{' || next == '}'
                        || isCjk(String.valueOf(next))) {
                        break;
                    }
                    end++;
                }
            }
            out.add(new TextRun(text.substring(i, end), cjk));
            i = end;
        }
        return out.isEmpty() ? null : out;
    }

    private static FormulaLayout layoutFlatRuns(List<TextRun> runs) {
        List<PlacedText> placed = new ArrayList<>();
        double x = 0d;
        double baseline = 9.6d;
        for (TextRun run : runs) {
            x += leadingMathSpacingPt(run, false);
            placed.add(new PlacedText(run.text(), run.cjk(), false, false, x, baseline, runWidthScale(run, false)));
            x += estimatedRunWidthPt(run) + trailingMathSpacingPt(run, false);
        }
        return new FormulaLayout(placed, Math.max(x, 1.0d), 13.0d);
    }

    private static FormulaLayout layoutScripts(String text) {
        List<PlacedText> placed = new ArrayList<>();
        double x = 0d;
        double shrinkAccum = 0d;
        int cursor = 0;
        boolean sawScript = false;
        boolean equationParenPower = repeatedEquationParenPower(text);
        while (cursor < text.length()) {
            int op = nextScriptOperator(text, cursor);
            if (op < 0) {
                String tail = text.substring(cursor);
                if (!tail.isBlank()) {
                    List<TextRun> suffix = tokenizeFlat(tail);
                    if (suffix == null) {
                        return null;
                    }
                    x = placeRuns(placed, suffix, x, 9.6d, false);
                }
                break;
            }
            int atomEnd = skipWhitespaceBackward(text, op);
            int atomStart = findAtomStart(text, atomEnd);
            if (atomStart < cursor || atomStart >= atomEnd) {
                return null;
            }
            ScriptGroup scriptGroup = readScriptGroup(text, op);
            if (scriptGroup == null) {
                return null;
            }
            int nextCursor = scriptGroup.end();
            ScriptGroup pairedScript = null;
            int pairedOp = skipWhitespaceForward(text, nextCursor);
            if (pairedOp < text.length() && (text.charAt(pairedOp) == '^' || text.charAt(pairedOp) == '_')
                && text.charAt(pairedOp) != text.charAt(op)) {
                pairedScript = readScriptGroup(text, pairedOp);
                if (pairedScript == null) {
                    return null;
                }
                nextCursor = pairedScript.end();
            }
            String prefixText = text.substring(cursor, atomStart);
            if (!prefixText.isBlank()) {
                List<TextRun> prefix = tokenizeFlat(prefixText);
                if (prefix == null) {
                    return null;
                }
                x = placeRuns(placed, prefix, x, 9.6d, false);
            }
            List<TextRun> base = tokenizeFlat(text.substring(atomStart, atomEnd));
            List<TextRun> script = tokenizeFlat(scriptGroup.body());
            if (base == null || script == null) {
                return null;
            }
            double baseWidth = estimatedWidthPt(base);
            int segmentStart = placed.size();
            x = placeRuns(placed, base, x, 9.6d, false);
            String baseText = text.substring(atomStart, atomEnd);
            boolean closingFenceScriptBase = isClosingFenceScriptBase(baseText);
            double scriptBaseline = scriptBaselinePt(scriptGroup.operator(), closingFenceScriptBase);
            double scriptShift = scriptShiftPt(baseWidth);
            double scriptStart = x + scriptShift;
            double scriptRight = placeRuns(placed, script, scriptStart, scriptBaseline, true);
            if (pairedScript != null) {
                List<TextRun> pairedRuns = tokenizeFlat(pairedScript.body());
                if (pairedRuns == null) {
                    return null;
                }
                double pairedBaseline = scriptBaselinePt(pairedScript.operator(), closingFenceScriptBase);
                double pairedShift = scriptShiftPt(baseWidth);
                double pairedStart = x + pairedShift;
                scriptRight = Math.max(scriptRight, placeRuns(placed, pairedRuns, pairedStart, pairedBaseline, true));
            }
            double segmentRight = Math.max(x, scriptRight);
            if (equationParenPower && closingFenceScriptBase && scriptGroup.operator() == '^') {
                SegmentScale scaled = scalePlacedSegmentX(placed, segmentStart, EQUATION_PAREN_POWER_WIDTH_SCALE);
                shrinkAccum += Math.max(0.0d, segmentRight - scaled.rightPt());
                segmentRight = scaled.rightPt();
                x = scaled.rightPt();
            }
            x = segmentRight + scriptTailPt(baseWidth);
            sawScript = true;
            cursor = nextCursor;
        }
        return sawScript && !placed.isEmpty() ? new FormulaLayout(placed, Math.max(x + shrinkAccum, 1.0d), 13.0d)
            : null;
    }

    private static SegmentScale scalePlacedSegmentX(List<PlacedText> placed, int startIndex, double scale) {
        if (placed == null || startIndex < 0 || startIndex >= placed.size() || scale >= 0.999d || scale <= 0.0d) {
            return new SegmentScale(0.0d);
        }
        double anchor = placed.get(startIndex).xPt();
        double right = anchor;
        for (int i = startIndex; i < placed.size(); i++) {
            PlacedText run = placed.get(i);
            double x = anchor + (run.xPt() - anchor) * scale;
            PlacedText scaled = new PlacedText(run.text(), run.cjk(), run.script(), run.display(), x,
                run.baselinePt(), run.widthScale() * scale);
            placed.set(i, scaled);
            right = Math.max(right, x + scaledRunWidthPt(new TextRun(run.text(), run.cjk()), run.script()) * scale);
        }
        return new SegmentScale(right);
    }

    static boolean repeatedEquationParenPower(String text) {
        if (text == null || !text.contains("=")) {
            return false;
        }
        int count = 0;
        int cursor = 0;
        while (cursor < text.length()) {
            int op = nextScriptOperator(text, cursor);
            if (op < 0) {
                return count >= 2;
            }
            int atomEnd = skipWhitespaceBackward(text, op);
            int atomStart = findAtomStart(text, atomEnd);
            ScriptGroup group = readScriptGroup(text, op);
            if (atomStart >= 0 && atomStart < atomEnd && group != null && group.operator() == '^'
                && isClosingFenceScriptBase(text.substring(atomStart, atomEnd))) {
                count++;
                if (count >= 2) {
                    return true;
                }
            }
            cursor = group == null ? op + 1 : group.end();
        }
        return false;
    }

    private static FormulaLayout scaleLayoutX(FormulaLayout layout, double scale) {
        if (layout == null || scale >= 0.999d || scale <= 0.0d) {
            return layout;
        }
        List<PlacedText> runs = new ArrayList<>(layout.runs().size());
        for (PlacedText run : layout.runs()) {
            runs.add(new PlacedText(run.text(), run.cjk(), run.script(), run.display(),
                run.xPt() * scale, run.baselinePt(), run.widthScale() * scale));
        }
        List<LineSegment> lines = new ArrayList<>(layout.lines().size());
        for (LineSegment line : layout.lines()) {
            lines.add(new LineSegment(line.x1Pt() * scale, line.y1Pt(), line.x2Pt() * scale, line.y2Pt()));
        }
        return new FormulaLayout(runs, lines, layout.widthPt(), layout.heightPt());
    }

    private static double scriptBaselinePt(char operator, boolean closingFenceBase) {
        if (operator == '^') {
            return closingFenceBase ? 3.8d : 5.2d;
        }
        return 12.4d;
    }

    private static boolean isClosingFenceScriptBase(String text) {
        if (text == null) {
            return false;
        }
        String normalized = text.replaceAll("\\s+", "");
        return ")".equals(normalized) || "]".equals(normalized) || "}".equals(normalized);
    }

    private static double scriptShiftPt(double baseWidth) {
        if (baseWidth < 9.0d) {
            return -0.35d;
        }
        return baseWidth >= 18.0d ? 0.75d : 0.25d;
    }

    private static double scriptTailPt(double baseWidth) {
        if (baseWidth < 9.0d) {
            return 0.25d;
        }
        return baseWidth >= 18.0d ? 0.55d : 0.35d;
    }

    private static ScriptGroup readScriptGroup(String text, int op) {
        if (op >= text.length() || (text.charAt(op) != '^' && text.charAt(op) != '_')) {
            return null;
        }
        int groupStart = skipWhitespaceForward(text, op + 1);
        if (groupStart >= text.length()) {
            return null;
        }
        if (text.charAt(groupStart) == '{') {
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return null;
            }
            return new ScriptGroup(text.charAt(op), text.substring(groupStart + 1, groupEnd), groupEnd + 1);
        }
        int groupEnd = findSimpleScriptAtomEnd(text, groupStart);
        return groupEnd > groupStart ? new ScriptGroup(text.charAt(op), text.substring(groupStart, groupEnd), groupEnd) : null;
    }

    private static FormulaLayout layoutStandaloneScript(String text) {
        int op = skipWhitespaceForward(text, 0);
        if (op >= text.length() || text.charAt(op) != '^') {
            return null;
        }
        int groupStart = skipWhitespaceForward(text, op + 1);
        if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
            return null;
        }
        int groupEnd = findGroupEnd(text, groupStart);
        if (groupEnd < 0 || !text.substring(groupEnd + 1).isBlank()) {
            return null;
        }
        List<TextRun> script = tokenizeFlat(text.substring(groupStart + 1, groupEnd));
        if (script == null) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        double width = placeRuns(placed, script, 0.0d, 5.2d, true);
        return new FormulaLayout(placed, Math.max(width, 1.0d), 13.0d);
    }

    private static FormulaLayout layoutFlattenedScripts(String text) {
        if (!text.contains("^") && !text.contains("_")) {
            return null;
        }
        String flattened = text.replace("^", "").replace("_", "");
        List<TextRun> runs = tokenizeFlat(flattened);
        return runs == null ? null : layoutFlatRuns(runs);
    }

    private static FormulaLayout layoutFractions(String text) {
        boolean compactInlineFraction = isCompactInlineFraction(text);
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0d;
        int cursor = 0;
        boolean sawFraction = false;
        while (cursor < text.length()) {
            int frac = text.indexOf("\\frac", cursor);
            if (frac < 0) {
                String tail = text.substring(cursor);
                if (!tail.isBlank()) {
                    FormulaLayout suffix = layoutFractionPart(tail);
                    if (suffix == null) {
                        return null;
                    }
                    appendLayout(placed, lines, suffix, x, 4.8d);
                    x += suffix.widthPt();
                }
                break;
            }
            String prefixText = text.substring(cursor, frac);
            if (!prefixText.isBlank()) {
                FormulaLayout prefix = layoutFractionPart(prefixText);
                if (prefix == null) {
                    return null;
                }
                appendLayout(placed, lines, prefix, x, 4.8d);
                x += prefix.widthPt();
            }
            int numeratorStart = skipWhitespaceForward(text, frac + 5);
            if (numeratorStart >= text.length() || text.charAt(numeratorStart) != '{') {
                return null;
            }
            int numeratorEnd = findGroupEnd(text, numeratorStart);
            if (numeratorEnd < 0) {
                return null;
            }
            int denominatorStart = skipWhitespaceForward(text, numeratorEnd + 1);
            if (denominatorStart >= text.length() || text.charAt(denominatorStart) != '{') {
                return null;
            }
            int denominatorEnd = findGroupEnd(text, denominatorStart);
            if (denominatorEnd < 0) {
                return null;
            }
            FormulaLayout numerator = layoutFractionPart(text.substring(numeratorStart + 1, numeratorEnd));
            FormulaLayout denominator = layoutFractionPart(text.substring(denominatorStart + 1, denominatorEnd));
            if (numerator == null || denominator == null) {
                return null;
            }
            double numeratorWidth = numerator.widthPt();
            double denominatorWidth = denominator.widthPt();
            double fractionPad = compactInlineFraction ? 1.6d : 3.0d;
            double fractionContentWidth = compactInlineFraction
                ? Math.max(numeratorWidth, denominatorWidth) * COMPACT_FRACTION_SCALE
                : Math.max(numeratorWidth, denominatorWidth);
            double fractionWidth = fractionContentWidth + fractionPad;
            double fractionX = x + 1.0d;
            if (compactInlineFraction) {
                appendCompactInlineFractionLayout(placed, lines, numerator,
                    fractionX + (fractionWidth - numeratorWidth * COMPACT_FRACTION_SCALE) / 2.0d,
                    COMPACT_FRACTION_NUMERATOR_Y_PT);
                appendCompactInlineFractionLayout(placed, lines, denominator,
                    fractionX + (fractionWidth - denominatorWidth * COMPACT_FRACTION_SCALE) / 2.0d,
                    COMPACT_FRACTION_DENOMINATOR_Y_PT);
                lines.add(new LineSegment(fractionX, COMPACT_FRACTION_BAR_Y_PT,
                    fractionX + fractionWidth, COMPACT_FRACTION_BAR_Y_PT));
            } else {
                appendLayout(placed, lines, numerator, fractionX + (fractionWidth - numeratorWidth) / 2.0d,
                    FRACTION_NUMERATOR_Y_PT);
                appendLayout(placed, lines, denominator, fractionX + (fractionWidth - denominatorWidth) / 2.0d,
                    FRACTION_DENOMINATOR_Y_PT);
                lines.add(new LineSegment(fractionX, FRACTION_BAR_Y_PT, fractionX + fractionWidth, FRACTION_BAR_Y_PT));
            }
            x = fractionX + fractionWidth + (compactInlineFraction ? 0.8d : 1.4d);
            sawFraction = true;
            cursor = denominatorEnd + 1;
        }
        return sawFraction && !placed.isEmpty()
            ? new FormulaLayout(placed, lines, Math.max(x, 1.0d),
                compactInlineFraction ? COMPACT_FRACTION_HEIGHT_PT : FRACTION_HEIGHT_PT)
            : null;
    }

    private static boolean isCompactInlineFraction(String text) {
        if (text == null || !text.contains("\\frac")) {
            return false;
        }
        if (text.contains("\\begin") || text.contains("\\sqrt") || text.contains("\\sum")
            || text.contains("\\displaystyle") || text.contains("\\dfrac")) {
            return false;
        }
        int count = 0;
        int cursor = 0;
        while ((cursor = text.indexOf("\\frac", cursor)) >= 0) {
            count++;
            cursor += 5;
        }
        return count <= 2 && text.length() <= 80 && !hasNestedFraction(text) && hasInlineFractionContext(text);
    }

    private static boolean hasNestedFraction(String text) {
        int cursor = 0;
        while ((cursor = text.indexOf("\\frac", cursor)) >= 0) {
            int numeratorStart = skipWhitespaceForward(text, cursor + 5);
            if (numeratorStart >= text.length() || text.charAt(numeratorStart) != '{') {
                return false;
            }
            int numeratorEnd = findGroupEnd(text, numeratorStart);
            if (numeratorEnd < 0) {
                return false;
            }
            int denominatorStart = skipWhitespaceForward(text, numeratorEnd + 1);
            if (denominatorStart >= text.length() || text.charAt(denominatorStart) != '{') {
                return false;
            }
            int denominatorEnd = findGroupEnd(text, denominatorStart);
            if (denominatorEnd < 0) {
                return false;
            }
            if (text.substring(numeratorStart + 1, numeratorEnd).contains("\\frac")
                || text.substring(denominatorStart + 1, denominatorEnd).contains("\\frac")) {
                return true;
            }
            cursor = denominatorEnd + 1;
        }
        return false;
    }

    private static boolean hasInlineFractionContext(String text) {
        StringBuilder nonFractionText = new StringBuilder();
        int cursor = 0;
        while (cursor < text.length()) {
            int frac = text.indexOf("\\frac", cursor);
            if (frac < 0) {
                nonFractionText.append(text.substring(cursor));
                break;
            }
            nonFractionText.append(text, cursor, frac);
            int numeratorStart = skipWhitespaceForward(text, frac + 5);
            if (numeratorStart >= text.length() || text.charAt(numeratorStart) != '{') {
                return true;
            }
            int numeratorEnd = findGroupEnd(text, numeratorStart);
            if (numeratorEnd < 0) {
                return true;
            }
            int denominatorStart = skipWhitespaceForward(text, numeratorEnd + 1);
            if (denominatorStart >= text.length() || text.charAt(denominatorStart) != '{') {
                return true;
            }
            int denominatorEnd = findGroupEnd(text, denominatorStart);
            if (denominatorEnd < 0) {
                return true;
            }
            cursor = denominatorEnd + 1;
        }
        String context = normalizeFlatLatex(nonFractionText.toString())
            .replaceAll("\\\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\\b", "")
            .replaceAll("\\\\(?:left|right)\\s*\\.?", "")
            .replaceAll("[\\s{}\\[\\]()（）,，.。:：;；]", "");
        return !context.isBlank();
    }

    private static FormulaLayout layoutFractionPart(String text) {
        if (text == null || text.isBlank()) {
            return new FormulaLayout(List.of(new PlacedText(" ", false, false, false, 0.0d, 9.6d)), List.of(), 4.0d, 13.0d);
        }
        text = normalizeTextCommands(text);
        FormulaLayout arrows = layoutXArrowFragments(text);
        if (arrows != null) {
            return arrows;
        }
        FormulaLayout overline = layoutOverline(text);
        if (overline != null) {
            return overline;
        }
        text = normalizeFlatLatex(text);
        ArraySlice array = readArraySlice(text, 0);
        if (array != null && array.end() == text.length()) {
            return layoutArray(array.body());
        }
        ArraySlice cases = readCasesSlice(text, 0);
        if (cases != null && cases.end() == text.length()) {
            return layoutLeftBraceArray(cases.body());
        }
        if (text.contains("\\frac")) {
            FormulaLayout fractions = layoutFractions(text);
            if (fractions != null) {
                return fractions;
            }
        }
        FormulaLayout sqrt = layoutSqrt(text);
        if (sqrt != null) {
            return sqrt;
        }
        FormulaLayout scripts = layoutScripts(text);
        if (scripts != null) {
            return scripts;
        }
        List<TextRun> runs = tokenizeFlat(text);
        return runs == null ? null : layoutFlatRuns(runs);
    }

    private static FormulaLayout layoutXArrowFragments(String text) {
        List<String> markers = List.of("\\xrightarrow", "\\xleftarrow");
        MarkerHit firstHit = findNextMarker(text, markers, 0);
        if (firstHit == null) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0.0d;
        double height = 16.0d;
        int cursor = 0;
        boolean sawArrow = false;
        while (cursor < text.length()) {
            MarkerHit hit = findNextMarker(text, markers, cursor);
            if (hit == null) {
                FormulaLayout suffix = layoutXArrowSidePart(text.substring(cursor));
                if (suffix == null) {
                    return null;
                }
                appendLayout(placed, lines, suffix, x, 2.0d);
                x += suffix.widthPt();
                height = Math.max(height, suffix.heightPt() + 2.0d);
                break;
            }
            if (hit.start() > cursor) {
                FormulaLayout prefix = layoutXArrowSidePart(text.substring(cursor, hit.start()));
                if (prefix == null) {
                    return null;
                }
                appendLayout(placed, lines, prefix, x, 2.0d);
                x += prefix.widthPt();
                height = Math.max(height, prefix.heightPt() + 2.0d);
            }
            int groupStart = skipWhitespaceForward(text, hit.start() + hit.marker().length());
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                return null;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return null;
            }
            FormulaLayout arrow = layoutSingleXArrow(text.substring(groupStart + 1, groupEnd), "\\xleftarrow".equals(hit.marker()));
            appendLayout(placed, lines, arrow, x, 0.0d);
            x += arrow.widthPt();
            height = Math.max(height, arrow.heightPt());
            sawArrow = true;
            cursor = groupEnd + 1;
        }
        return sawArrow ? new FormulaLayout(placed, lines, Math.max(1.0d, x), height) : null;
    }

    private static FormulaLayout layoutXArrowSidePart(String text) {
        if (text == null || text.isBlank()) {
            return new FormulaLayout(List.of(), List.of(), 0.0d, 13.0d);
        }
        FormulaLayout fractions = text.contains("\\frac") ? layoutFractions(text) : null;
        if (fractions != null) {
            return fractions;
        }
        FormulaLayout scripts = layoutScripts(normalizeFlatLatex(text));
        if (scripts != null) {
            return scripts;
        }
        List<TextRun> runs = tokenizeFlat(text);
        return runs == null ? null : layoutFlatRuns(runs);
    }

    private static FormulaLayout layoutSingleXArrow(String label, boolean leftArrow) {
        List<TextRun> labelRuns = tokenizeFlat(normalizeTextCommands(label));
        if (labelRuns == null) {
            labelRuns = List.of(new TextRun(label == null ? "" : label.trim(), true));
        }
        double labelWidth = Math.max(0.0d, estimatedWidthPt(labelRuns));
        double width = Math.max(13.0d, labelWidth + 5.0d);
        List<PlacedText> placed = new ArrayList<>();
        double labelX = Math.max(0.0d, (width - labelWidth) / 2.0d);
        placeRuns(placed, labelRuns, labelX, 6.0d, true);
        List<LineSegment> lines = new ArrayList<>();
        double y = 11.0d;
        lines.add(new LineSegment(1.0d, y, width - 1.0d, y));
        if (leftArrow) {
            lines.add(new LineSegment(1.0d, y, 4.0d, y - 2.0d));
            lines.add(new LineSegment(1.0d, y, 4.0d, y + 2.0d));
        } else {
            lines.add(new LineSegment(width - 1.0d, y, width - 4.0d, y - 2.0d));
            lines.add(new LineSegment(width - 1.0d, y, width - 4.0d, y + 2.0d));
        }
        return new FormulaLayout(placed, lines, width, 15.0d);
    }

    private static void appendLayout(List<PlacedText> placed, List<LineSegment> lines, FormulaLayout layout,
        double dx, double dy) {
        for (PlacedText run : layout.runs()) {
            placed.add(new PlacedText(run.text(), run.cjk(), run.script(), run.display(), run.xPt() + dx,
                run.baselinePt() + dy));
        }
        for (LineSegment line : layout.lines()) {
            lines.add(new LineSegment(line.x1Pt() + dx, line.y1Pt() + dy, line.x2Pt() + dx, line.y2Pt() + dy));
        }
    }

    private static void appendLayoutAsScript(List<PlacedText> placed, List<LineSegment> lines, FormulaLayout layout,
        double dx, double dy) {
        appendScaledLayout(placed, lines, layout, 0.66d, dx, dy, true);
    }

    private static void appendScaledLayout(List<PlacedText> placed, List<LineSegment> lines, FormulaLayout layout,
        double scale, double dx, double dy, boolean script) {
        for (PlacedText run : layout.runs()) {
            placed.add(new PlacedText(run.text(), run.cjk(), script || run.script(), run.display(),
                dx + run.xPt() * scale, dy + run.baselinePt() * scale));
        }
        for (LineSegment line : layout.lines()) {
            lines.add(new LineSegment(dx + line.x1Pt() * scale, dy + line.y1Pt() * scale,
                dx + line.x2Pt() * scale, dy + line.y2Pt() * scale));
        }
    }

    private static void appendCompactInlineFractionLayout(List<PlacedText> placed, List<LineSegment> lines,
        FormulaLayout layout, double dx, double dy) {
        for (PlacedText run : layout.runs()) {
            placed.add(new PlacedText(run.text(), run.cjk(), run.script(), run.display(),
                dx + run.xPt() * COMPACT_FRACTION_SCALE, dy + run.baselinePt() * COMPACT_FRACTION_SCALE,
                run.widthScale() * COMPACT_FRACTION_SCALE));
        }
        for (LineSegment line : layout.lines()) {
            lines.add(new LineSegment(dx + line.x1Pt() * COMPACT_FRACTION_SCALE,
                dy + line.y1Pt() * COMPACT_FRACTION_SCALE,
                dx + line.x2Pt() * COMPACT_FRACTION_SCALE,
                dy + line.y2Pt() * COMPACT_FRACTION_SCALE));
        }
    }

    private static FormulaLayout layoutLeftBraceArray(String body) {
        FormulaLayout inner = layoutArray(body);
        if (inner == null) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        double braceBaseline = Math.max(16.0d, Math.min(inner.heightPt() - 1.0d, inner.heightPt() * 0.72d));
        placed.add(new PlacedText("{", false, false, true, 0.0d, braceBaseline));
        List<LineSegment> lines = new ArrayList<>();
        appendLayout(placed, lines, inner, 7.0d, 0.0d);
        return new FormulaLayout(placed, lines, inner.widthPt() + 8.5d, inner.heightPt());
    }

    private static FormulaLayout layoutCases(String text) {
        String begin = "\\begin{cases}";
        String end = "\\end{cases}";
        if (!text.startsWith(begin) || !text.endsWith(end)) {
            return null;
        }
        String body = text.substring(begin.length(), text.length() - end.length()).trim();
        if (body.isBlank()) {
            return null;
        }
        return layoutLeftBraceArray(body);
    }

    private static FormulaLayout layoutCompositeCases(String text) {
        String beginToken = "\\begin{cases}";
        int first = text.indexOf(beginToken);
        if (first < 0) {
            return null;
        }
        ArraySlice single = readCasesSlice(text, first);
        if (single != null && first == 0 && single.end() == text.length()) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0d;
        double height = 13.0d;
        int cursor = 0;
        boolean sawCases = false;
        while (cursor < text.length()) {
            int begin = text.indexOf(beginToken, cursor);
            if (begin < 0) {
                String tail = text.substring(cursor);
                if (!tail.isBlank()) {
                    FormulaLayout flat = layoutFlatText(tail);
                    if (flat == null) {
                        return null;
                    }
                    appendLayout(placed, lines, flat, x, 0.0d);
                    x += flat.widthPt() + 1.5d;
                    height = Math.max(height, flat.heightPt());
                }
                break;
            }
            String prefix = text.substring(cursor, begin);
            if (!prefix.isBlank()) {
                FormulaLayout flat = layoutFlatText(prefix);
                if (flat == null) {
                    return null;
                }
                appendLayout(placed, lines, flat, x, 0.0d);
                x += flat.widthPt() + 1.5d;
                height = Math.max(height, flat.heightPt());
            }
            ArraySlice slice = readCasesSlice(text, begin);
            if (slice == null) {
                return null;
            }
            FormulaLayout cases = layoutLeftBraceArray(slice.body());
            if (cases == null) {
                return null;
            }
            appendLayout(placed, lines, cases, x, 0.0d);
            x += cases.widthPt() + 2.0d;
            height = Math.max(height, cases.heightPt());
            sawCases = true;
            cursor = slice.end();
        }
        return sawCases && !placed.isEmpty() ? new FormulaLayout(placed, lines, Math.max(1.0d, x), height) : null;
    }

    private static FormulaLayout layoutParenArray(String body) {
        FormulaLayout inner = layoutArray(body);
        if (inner == null) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        double braceBaseline = Math.max(16.0d, Math.min(inner.heightPt() - 1.0d, inner.heightPt() * 0.72d));
        placed.add(new PlacedText("(", false, false, true, 0.0d, braceBaseline));
        List<LineSegment> lines = new ArrayList<>();
        appendLayout(placed, lines, inner, 7.0d, 0.0d);
        placed.add(new PlacedText(")", false, false, true, inner.widthPt() + 7.5d, braceBaseline));
        return new FormulaLayout(placed, lines, inner.widthPt() + 16.0d, inner.heightPt());
    }

    private static FormulaLayout layoutCompositeArrays(String text) {
        text = normalizeFlatFences(text);
        if (!text.contains("\\begin{array}") || text.indexOf("\\begin{array}") == text.lastIndexOf("\\begin{array}")) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0d;
        double height = 13.0d;
        int cursor = 0;
        boolean sawArray = false;
        while (cursor < text.length()) {
            int begin = text.indexOf("\\begin{array}", cursor);
            if (begin < 0) {
                String tail = text.substring(cursor);
                if (!tail.isBlank()) {
                    FormulaLayout flat = layoutFlatText(tail);
                    if (flat == null) {
                        return null;
                    }
                    appendLayout(placed, lines, flat, x, 0.0d);
                    x += flat.widthPt() + 1.5d;
                    height = Math.max(height, flat.heightPt());
                }
                break;
            }
            String prefix = text.substring(cursor, begin);
            if (!prefix.isBlank()) {
                FormulaLayout flat = layoutFlatText(prefix);
                if (flat == null) {
                    return null;
                }
                appendLayout(placed, lines, flat, x, 0.0d);
                x += flat.widthPt() + 1.5d;
                height = Math.max(height, flat.heightPt());
            }
            ArraySlice slice = readArraySlice(text, begin);
            if (slice == null || slice.body().contains("\\begin{array}")) {
                return null;
            }
            FormulaLayout array = layoutArray(slice.body());
            if (array == null) {
                return null;
            }
            appendLayout(placed, lines, array, x, 0.0d);
            x += array.widthPt() + 2.0d;
            height = Math.max(height, array.heightPt());
            sawArray = true;
            cursor = slice.end();
        }
        return sawArray && !placed.isEmpty() ? new FormulaLayout(placed, lines, Math.max(1.0d, x - 1.5d), height)
            : null;
    }

    private static String normalizeFlatFences(String text) {
        text = text.replaceAll("\\\\left\\s*\\(", "(");
        text = text.replaceAll("\\\\right\\s*\\)", ")");
        text = text.replaceAll("\\\\left\\s*\\[", "[");
        text = text.replaceAll("\\\\right\\s*]", "]");
        text = text.replaceAll("\\\\left\\s*\\|", "|");
        text = text.replaceAll("\\\\right\\s*\\|", "|");
        text = text.replaceAll("\\\\left\\s*\\\\\\{", "\\\\{");
        text = text.replaceAll("\\\\right\\s*\\\\}", "\\\\}");
        text = text.replaceAll("\\\\left\\s*\\.", "");
        text = text.replaceAll("\\\\right\\s*\\.", "");
        return text;
    }

    private static FormulaLayout layoutFlatText(String text) {
        List<TextRun> runs = tokenizeFlat(text);
        return runs == null ? null : layoutFlatRuns(runs);
    }

    private static FormulaLayout layoutArray(String body) {
        List<String> rowText = splitTopLevelRows(body);
        List<List<FormulaLayout>> rows = new ArrayList<>();
        List<Integer> hlineRows = new ArrayList<>();
        int columnCount = 0;
        int parsedRowIndex = 0;
        for (String row : rowText) {
            boolean hasHline = row.contains("\\hline");
            row = row.replace("\\hline", "").trim();
            if (hasHline) {
                hlineRows.add(parsedRowIndex);
            }
            if (row.isEmpty()) {
                continue;
            }
            List<String> cells = splitTopLevelCells(row);
            List<FormulaLayout> parsedRow = new ArrayList<>();
            for (String cell : cells) {
                String normalized = normalizeArrayCell(cell);
                FormulaLayout layout = normalized.isEmpty()
                    ? new FormulaLayout(List.of(), List.of(), 0.0d, 13.0d)
                    : layoutArrayCell(normalized);
                if (layout == null) {
                    return null;
                }
                parsedRow.add(layout);
            }
            rows.add(parsedRow);
            columnCount = Math.max(columnCount, parsedRow.size());
            parsedRowIndex++;
        }
        if (rows.isEmpty() || columnCount == 0) {
            return null;
        }
        double[] widths = new double[columnCount];
        int nonEmptyCells = 0;
        int totalCells = 0;
        for (List<FormulaLayout> row : rows) {
            for (int i = 0; i < row.size(); i++) {
                FormulaLayout cell = row.get(i);
                widths[i] = Math.max(widths[i], cell.widthPt());
                totalCells++;
                if (!cell.runs().isEmpty() || !cell.lines().isEmpty()) {
                    nonEmptyCells++;
                }
            }
        }
        boolean sparseGrid = columnCount >= 6 && totalCells > 0 && nonEmptyCells * 2 <= totalCells;
        for (int i = 0; i < widths.length; i++) {
            widths[i] = Math.max(widths[i], sparseGrid ? 8.0d : 4.0d);
        }
        double[] x = new double[columnCount];
        double cursor = 0d;
        double gap = sparseGrid ? 4.0d : 2.0d;
        for (int i = 0; i < columnCount; i++) {
            x[i] = cursor;
            cursor += widths[i] + gap;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double rowHeight = rows.size() > 1 ? 12.0d : 13.0d;
        double width = Math.max(cursor - gap, 1.0d);
        for (int rowIndex : hlineRows) {
            double y = Math.max(1.5d, rowIndex * rowHeight - 1.2d);
            lines.add(new LineSegment(0.0d, y, width, y));
        }
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            List<FormulaLayout> row = rows.get(rowIndex);
            double y = rowIndex * rowHeight;
            for (int col = 0; col < row.size(); col++) {
                FormulaLayout cell = row.get(col);
                double runX = x[col] + Math.max(0d, (widths[col] - cell.widthPt()) / 2.0d);
                appendLayout(placed, lines, cell, runX, y);
            }
        }
        double height = Math.max(13.0d, 2.5d + rows.size() * rowHeight);
        if (placed.isEmpty()) {
            placed.add(new PlacedText(" ", false, false, false, 0.0d, 9.6d));
        }
        return new FormulaLayout(placed, lines, width, height);
    }

    private static FormulaLayout layoutArrayCell(String text) {
        text = LaTeXParser.preNormalizeLatex(text);
        text = normalizeFlatFences(text);
        ArraySlice nested = readArraySlice(text, 0);
        if (nested != null && nested.end() == text.length()) {
            return layoutArray(nested.body());
        }
        ArraySlice nestedCases = readCasesSlice(text, 0);
        if (nestedCases != null && nestedCases.end() == text.length()) {
            return layoutLeftBraceArray(nestedCases.body());
        }
        FormulaLayout compositeCases = layoutCompositeCases(text);
        if (compositeCases != null) {
            return compositeCases;
        }
        FormulaLayout boxedFragments = layoutBoxedFragments(text);
        if (boxedFragments != null) {
            return boxedFragments;
        }
        FormulaLayout boxed = layoutBoxed(text);
        if (boxed != null) {
            return boxed;
        }
        FormulaLayout arrows = layoutXArrowFragments(text);
        if (arrows != null) {
            return arrows;
        }
        FormulaLayout fractions = layoutFractions(text);
        if (fractions != null) {
            return fractions;
        }
        FormulaLayout scripts = layoutScripts(text);
        if (scripts != null) {
            return scripts;
        }
        FormulaLayout standaloneScript = layoutStandaloneScript(text);
        if (standaloneScript != null) {
            return standaloneScript;
        }
        List<TextRun> runs = tokenizeFlat(text);
        return runs == null ? null : layoutFlatRuns(runs);
    }

    private static String normalizeArrayCell(String cell) {
        String normalized = cell.trim();
        if ("{}".equals(normalized)) {
            normalized = "";
        }
        while (normalized.startsWith(",\\begin{array}")) {
            normalized = normalized.substring(1).stripLeading();
        }
        return normalized;
    }

    private static List<String> splitTopLevelRows(String body) {
        List<String> rows = new ArrayList<>();
        int depth = 0;
        int start = 0;
        for (int i = 0; i < body.length() - 1; i++) {
            if (startsAnyMathArrayBegin(body, i)) {
                depth++;
                continue;
            }
            if (startsAnyMathArrayEnd(body, i)) {
                depth = Math.max(0, depth - 1);
                continue;
            }
            if (depth == 0 && body.charAt(i) == '\\' && body.charAt(i + 1) == '\\') {
                rows.add(body.substring(start, i));
                int next = i + 2;
                if (next < body.length() && body.charAt(next) == ',') {
                    next++;
                }
                start = next;
                i = next - 1;
            }
        }
        rows.add(body.substring(start));
        return rows;
    }

    private static List<String> splitTopLevelCells(String row) {
        List<String> cells = new ArrayList<>();
        int depth = 0;
        int start = 0;
        for (int i = 0; i < row.length(); i++) {
            if (startsAnyMathArrayBegin(row, i)) {
                depth++;
                continue;
            }
            if (startsAnyMathArrayEnd(row, i)) {
                depth = Math.max(0, depth - 1);
                continue;
            }
            if (depth == 0 && row.charAt(i) == '&' && !isEscaped(row, i)) {
                cells.add(row.substring(start, i));
                start = i + 1;
            }
        }
        cells.add(row.substring(start));
        return cells;
    }

    private static String normalizeFlatLatex(String text) {
        text = text.replace("\\_", "_");
        text = text.replace("\\enspace", " ");
        text = text.replace("\\:", " ");
        text = replaceThinSpaces(text);
        text = text.replaceAll("\\{\\s*([_^])\\s*\\{([^{}]*)}\\s*}", "$1{$2}");
        text = text.replaceAll("\\\\(?:mathrm|mathbf|mathit|textit|textbf|emph|text|boldsymbol)\\s*\\{\\s*(\\\\[A-Za-z]+)\\s*}", "$1");
        String previous;
        do {
            previous = text;
            text = text.replaceAll("\\\\begin\\{array}\\{[^}]*}\\s*\\\\end\\{array}", "");
        } while (!text.equals(previous));
        text = unwrapSingleLineArrayFragments(text);
        text = normalizeFlatFences(text);
        Matcher matcher = LEFT_RIGHT_PAREN.matcher(text);
        StringBuffer out = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(out, Matcher.quoteReplacement("(" + matcher.group(1).trim() + ")"));
        }
        matcher.appendTail(out);
        String normalized = out.toString();
        if (!normalized.contains("\\begin{array}")
            && !normalized.contains("\\begin{aligned}")
            && !normalized.contains("\\begin{cases}")) {
            normalized = removeFlatAlignmentAmpersands(normalized);
        }
        return normalized;
    }

    private static String removeFlatAlignmentAmpersands(String text) {
        if (text == null || text.indexOf('&') < 0) {
            return text;
        }
        StringBuilder out = new StringBuilder(text.length());
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '&') {
                if (i > 0 && text.charAt(i - 1) == '\\') {
                    if (out.length() > 0) {
                        out.setLength(out.length() - 1);
                    }
                    out.append('&');
                }
                continue;
            }
            out.append(ch);
        }
        return out.toString();
    }

    private static boolean isEscaped(String text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text.charAt(i) == '\\'; i--) {
            slashCount++;
        }
        return (slashCount & 1) == 1;
    }

    private static String normalizeHorizontalBraceAnnotations(String text) {
        String normalized = text;
        normalized = replaceAnnotatedBrace(normalized, "\\overset", "\\overbrace");
        normalized = replaceAnnotatedBrace(normalized, "\\underset", "\\underbrace");
        return normalized;
    }

    private static String unwrapSingleLineArrayFragments(String text) {
        String marker = "\\begin{array}";
        if (text == null || !text.contains(marker)) {
            return text;
        }
        StringBuilder out = new StringBuilder(text.length());
        int cursor = 0;
        while (cursor < text.length()) {
            int start = text.indexOf(marker, cursor);
            if (start < 0) {
                out.append(text.substring(cursor));
                break;
            }
            out.append(text, cursor, start);
            ArraySlice slice = readArraySlice(text, start);
            if (slice == null || slice.body().contains("\\\\")) {
                out.append(marker);
                cursor = start + marker.length();
                continue;
            }
            out.append(slice.body().replace('&', ' ').replaceAll("\\s+", " ").trim());
            cursor = slice.end();
        }
        return out.toString();
    }

    private static String replaceAnnotatedBrace(String text, String command, String braceCommand) {
        String marker = command + "{";
        int start = text.indexOf(marker);
        while (start >= 0) {
            int annotationStart = start + command.length();
            int annotationEnd = findGroupEnd(text, annotationStart);
            if (annotationEnd < 0) {
                return text;
            }
            int secondArgStart = skipWhitespaceForward(text, annotationEnd + 1);
            if (secondArgStart >= text.length() || text.charAt(secondArgStart) != '{') {
                start = text.indexOf(marker, start + marker.length());
                continue;
            }
            int secondArgEnd = findGroupEnd(text, secondArgStart);
            if (secondArgEnd < 0) {
                return text;
            }
            int braceStart = skipWhitespaceForward(text, secondArgStart + 1);
            if (!text.startsWith(braceCommand + "{", braceStart)) {
                start = text.indexOf(marker, start + marker.length());
                continue;
            }
            int bodyStart = braceStart + braceCommand.length();
            int bodyEnd = findGroupEnd(text, bodyStart);
            if (bodyEnd < 0 || bodyEnd > secondArgEnd) {
                return text;
            }
            String annotation = text.substring(annotationStart + 1, annotationEnd);
            String body = text.substring(bodyStart + 1, bodyEnd);
            String replacement = body + " " + annotation;
            text = text.substring(0, start) + replacement + text.substring(secondArgEnd + 1);
            start = text.indexOf(marker, start + replacement.length());
        }
        return text;
    }

    private static String normalizeTextCommands(String text) {
        boolean changed;
        do {
            String before = text;
            text = replaceTextColorCommands(text);
            text = unwrapTextLikeCommandBodies(text);
            Matcher matcher = TEXT_COMMAND_PATTERN.matcher(text);
            StringBuffer out = new StringBuffer();
            while (matcher.find()) {
                matcher.appendReplacement(out, Matcher.quoteReplacement(matcher.group(1).trim()));
            }
            matcher.appendTail(out);
            text = out.toString();
            changed = !text.equals(before);
        } while (changed);
        return text;
    }

    private static String unwrapTextLikeCommandBodies(String text) {
        List<String> commands = List.of("\\boldsymbol", "\\textnormal", "\\mathrm", "\\mathbf", "\\mathit",
            "\\textit", "\\textbf", "\\textrm", "\\emph", "\\text");
        StringBuilder out = new StringBuilder(text.length());
        for (int i = 0; i < text.length();) {
            String command = matchingCommandAt(text, i, commands);
            if (command == null) {
                out.append(text.charAt(i));
                i++;
                continue;
            }
            int groupStart = skipWhitespaceForward(text, i + command.length());
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                out.append(text.charAt(i));
                i++;
                continue;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                out.append(text.charAt(i));
                i++;
                continue;
            }
            out.append(unwrapTextLikeCommandBodies(text.substring(groupStart + 1, groupEnd).trim()));
            i = groupEnd + 1;
        }
        return out.toString();
    }

    private static String matchingCommandAt(String text, int index, List<String> commands) {
        for (String command : commands) {
            if (!text.startsWith(command, index)) {
                continue;
            }
            int end = index + command.length();
            if (end < text.length() && Character.isLetter(text.charAt(end))) {
                continue;
            }
            return command;
        }
        return null;
    }

    private static String normalizeSpacingCommands(String text) {
        return replaceThinSpaces(text.replace("\\enspace", " ")
            .replace("\\quad", " ")
            .replace("\\qquad", " ")
            .replace("\\:", " "));
    }

    private static String replaceThinSpaces(String text) {
        StringBuilder out = new StringBuilder(text.length());
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '\\' && i + 1 < text.length() && text.charAt(i + 1) == ','
                && (i == 0 || text.charAt(i - 1) != '\\')) {
                out.append(' ');
                i++;
            } else {
                out.append(ch);
            }
        }
        return out.toString();
    }

    private static String normalizeSimpleDecorations(String text) {
        for (String command : List.of("\\dot", "\\bar", "\\hat", "\\tilde", "\\vec", "\\underline")) {
            text = replaceUnaryCommandWithBody(text, command);
        }
        return text;
    }

    private static FormulaLayout layoutUnderline(String text) {
        String marker = "\\underline";
        if (!text.startsWith(marker)) {
            return null;
        }
        int groupStart = skipWhitespaceForward(text, marker.length());
        if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
            return null;
        }
        int groupEnd = findGroupEnd(text, groupStart);
        if (groupEnd < 0 || !text.substring(groupEnd + 1).isBlank()) {
            return null;
        }
        FormulaLayout inner = layoutFractionPart(text.substring(groupStart + 1, groupEnd));
        if (inner == null) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        appendLayout(placed, lines, inner, 0.0d, 0.0d);
        double width = Math.max(inner.widthPt(), 10.0d);
        double height = Math.max(inner.heightPt(), 13.0d);
        lines.add(new LineSegment(0.0d, height - 1.2d, width, height - 1.2d));
        return new FormulaLayout(placed, lines, width, height);
    }

    private static String normalizeOverUnderSetCommands(String text) {
        text = replaceBinaryCommandWithSecondArg(text, "\\overset");
        text = replaceBinaryCommandWithSecondArg(text, "\\underset");
        return text;
    }

    private static String replaceBinaryCommandWithSecondArg(String text, String command) {
        int start = text.indexOf(command);
        while (start >= 0) {
            int firstStart = skipWhitespaceForward(text, start + command.length());
            if (firstStart >= text.length() || text.charAt(firstStart) != '{') {
                start = text.indexOf(command, start + command.length());
                continue;
            }
            int firstEnd = findGroupEnd(text, firstStart);
            if (firstEnd < 0) {
                return text;
            }
            int secondStart = skipWhitespaceForward(text, firstEnd + 1);
            if (secondStart >= text.length() || text.charAt(secondStart) != '{') {
                start = text.indexOf(command, firstEnd + 1);
                continue;
            }
            int secondEnd = findGroupEnd(text, secondStart);
            if (secondEnd < 0) {
                return text;
            }
            String body = text.substring(secondStart + 1, secondEnd);
            text = text.substring(0, start) + body + text.substring(secondEnd + 1);
            start = text.indexOf(command, start + body.length());
        }
        return text;
    }

    private static String replaceUnaryCommandWithBody(String text, String command) {
        int start = text.indexOf(command);
        while (start >= 0) {
            int groupStart = skipWhitespaceForward(text, start + command.length());
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                start = text.indexOf(command, start + command.length());
                continue;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return text;
            }
            String body = text.substring(groupStart + 1, groupEnd);
            text = text.substring(0, start) + body + text.substring(groupEnd + 1);
            start = text.indexOf(command, start + body.length());
        }
        return text;
    }

    private static String normalizeWrappedScriptBases(String text) {
        StringBuilder out = new StringBuilder();
        for (int i = 0; i < text.length();) {
            if (text.charAt(i) == '{' && !isCommandArgumentBrace(text, i)) {
                int end = findGroupEnd(text, i);
                if (end > i && end + 1 < text.length() && (text.charAt(end + 1) == '^' || text.charAt(end + 1) == '_')) {
                    out.append(text, i + 1, end);
                    i = end + 1;
                    continue;
                }
            }
            out.append(text.charAt(i));
            i++;
        }
        return out.toString();
    }

    private static boolean isCommandArgumentBrace(String text, int braceIndex) {
        int i = braceIndex - 1;
        while (i >= 0 && Character.isWhitespace(text.charAt(i))) {
            i--;
        }
        if (i < 0 || !Character.isLetter(text.charAt(i))) {
            return false;
        }
        while (i >= 0 && Character.isLetter(text.charAt(i))) {
            i--;
        }
        return i >= 0 && text.charAt(i) == '\\';
    }

    private static String replaceTextColorCommands(String text) {
        String marker = "\\textcolor";
        int start = text.indexOf(marker);
        while (start >= 0) {
            int colorStart = skipWhitespaceForward(text, start + marker.length());
            if (colorStart >= text.length() || text.charAt(colorStart) != '{') {
                return text;
            }
            int colorEnd = findGroupEnd(text, colorStart);
            if (colorEnd < 0) {
                return text;
            }
            int bodyStart = skipWhitespaceForward(text, colorEnd + 1);
            if (bodyStart >= text.length() || text.charAt(bodyStart) != '{') {
                return text;
            }
            int bodyEnd = findGroupEnd(text, bodyStart);
            if (bodyEnd < 0) {
                return text;
            }
            String body = text.substring(bodyStart + 1, bodyEnd);
            text = text.substring(0, start) + body + text.substring(bodyEnd + 1);
            start = text.indexOf(marker, start + body.length());
        }
        return text;
    }

    private static FormulaLayout layoutBoxed(String text) {
        String marker = "\\boxed";
        if (!text.startsWith(marker)) {
            return null;
        }
        int groupStart = skipWhitespaceForward(text, marker.length());
        if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
            return null;
        }
        int groupEnd = findGroupEnd(text, groupStart);
        if (groupEnd < 0 || skipWhitespaceForward(text, groupEnd + 1) != text.length()) {
            return null;
        }
        String body = text.substring(groupStart + 1, groupEnd);
        FormulaLayout inner = body.isBlank()
            ? new FormulaLayout(List.of(), List.of(), 0.0d, 13.0d)
            : layoutArrayCell(body);
        if (inner == null) {
            return null;
        }
        double padX = 2.0d;
        double padTop = 2.0d;
        double padBottom = 2.0d;
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        appendLayout(placed, lines, inner, padX, padTop);
        double width = Math.max(6.0d, inner.widthPt() + padX * 2.0d);
        double height = Math.max(8.0d, inner.heightPt() + padTop + padBottom);
        lines.add(new LineSegment(0.0d, 0.5d, width, 0.5d));
        lines.add(new LineSegment(width, 0.5d, width, height - 0.5d));
        lines.add(new LineSegment(width, height - 0.5d, 0.0d, height - 0.5d));
        lines.add(new LineSegment(0.0d, height - 0.5d, 0.0d, 0.5d));
        return new FormulaLayout(placed, lines, width, height);
    }

    private static FormulaLayout layoutCancel(String text) {
        List<String> markers = List.of("\\xcancel", "\\bcancel", "\\cancel");
        MarkerHit hit = findNextMarker(text, markers, 0);
        if (hit == null || hit.start() != 0) {
            return null;
        }
        int groupStart = skipWhitespaceForward(text, hit.marker().length());
        if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
            return null;
        }
        int groupEnd = findGroupEnd(text, groupStart);
        if (groupEnd < 0 || !text.substring(groupEnd + 1).isBlank()) {
            return null;
        }
        FormulaLayout inner = layoutFractionPart(text.substring(groupStart + 1, groupEnd));
        if (inner == null) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        appendLayout(placed, lines, inner, 0.0d, 0.0d);
        double width = Math.max(inner.widthPt(), 6.0d);
        double height = Math.max(inner.heightPt(), 8.0d);
        if ("\\cancel".equals(hit.marker()) || "\\xcancel".equals(hit.marker())) {
            lines.add(new LineSegment(0.0d, height - 1.0d, width, 1.0d));
        }
        if ("\\bcancel".equals(hit.marker()) || "\\xcancel".equals(hit.marker())) {
            lines.add(new LineSegment(0.0d, 1.0d, width, height - 1.0d));
        }
        return new FormulaLayout(placed, lines, width, height);
    }

    private static FormulaLayout layoutBoxedFragments(String text) {
        String marker = "\\boxed";
        if (!text.contains(marker)) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0.0d;
        double height = 13.0d;
        boolean sawBoxed = false;
        int cursor = 0;
        while (cursor < text.length()) {
            int start = text.indexOf(marker, cursor);
            if (start < 0) {
                FormulaLayout suffix = layoutFractionPart(text.substring(cursor));
                if (suffix == null) {
                    return null;
                }
                appendLayout(placed, lines, suffix, x, 0.0d);
                x += suffix.widthPt();
                height = Math.max(height, suffix.heightPt());
                break;
            }
            if (start > cursor) {
                FormulaLayout prefix = layoutFractionPart(text.substring(cursor, start));
                if (prefix == null) {
                    return null;
                }
                appendLayout(placed, lines, prefix, x, 0.0d);
                x += prefix.widthPt();
                height = Math.max(height, prefix.heightPt());
            }
            int groupStart = skipWhitespaceForward(text, start + marker.length());
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                return null;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return null;
            }
            FormulaLayout boxed = layoutBoxed(text.substring(start, groupEnd + 1));
            if (boxed == null) {
                return null;
            }
            appendLayout(placed, lines, boxed, x, 0.0d);
            x += boxed.widthPt() + 1.8d;
            height = Math.max(height, boxed.heightPt());
            sawBoxed = true;
            cursor = groupEnd + 1;
        }
        return sawBoxed ? new FormulaLayout(placed, lines, Math.max(1.0d, x), height) : null;
    }

    private static FormulaLayout layoutOverline(String text) {
        String marker = "\\overline";
        if (!text.contains(marker)) {
            return null;
        }
        text = normalizeFlatFences(text);
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> overlines = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0.0d;
        double height = 14.0d;
        int cursor = 0;
        boolean sawOverline = false;
        while (cursor < text.length()) {
            int start = text.indexOf(marker, cursor);
            if (start < 0) {
                FormulaLayout suffix = layoutFractionPart(text.substring(cursor));
                if (suffix == null) {
                    return null;
                }
                appendLayout(placed, lines, suffix, x, 0.0d);
                x += suffix.widthPt();
                height = Math.max(height, suffix.heightPt());
                break;
            }
            if (start > cursor) {
                FormulaLayout prefix = layoutFractionPart(text.substring(cursor, start));
                if (prefix == null) {
                    return null;
                }
                appendLayout(placed, lines, prefix, x, 0.0d);
                x += prefix.widthPt();
                height = Math.max(height, prefix.heightPt());
            }
            int groupStart = skipWhitespaceForward(text, start + marker.length());
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                return null;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return null;
            }
            String body = text.substring(groupStart + 1, groupEnd);
            if (body.contains(marker)) {
                return null;
            }
            FormulaLayout bodyLayout = layoutBoxedAwarePart(body);
            if (bodyLayout == null) {
                return null;
            }
            appendLayout(placed, lines, bodyLayout, x, 0.0d);
            overlines.add(new LineSegment(x, 1.8d, x + bodyLayout.widthPt(), 1.8d));
            x += bodyLayout.widthPt();
            height = Math.max(height, bodyLayout.heightPt());
            sawOverline = true;
            cursor = groupEnd + 1;
            int scriptOp = skipWhitespaceForward(text, cursor);
            if (scriptOp < text.length() && (text.charAt(scriptOp) == '^' || text.charAt(scriptOp) == '_')) {
                ScriptGroup script = readScriptGroup(text, scriptOp);
                if (script == null) {
                    return null;
                }
                List<TextRun> scriptRuns = tokenizeFlat(script.body());
                if (scriptRuns == null) {
                    return null;
                }
                double scriptBaseline = script.operator() == '^' ? 5.2d : 12.4d;
                double scriptShift = bodyLayout.widthPt() >= 18.0d ? 0.75d : 0.55d;
                double scriptWidth = placeRuns(placed, scriptRuns, x + scriptShift, scriptBaseline, true) - x;
                x += Math.max(0.0d, scriptWidth);
                height = Math.max(height, script.operator() == '^' ? 14.0d : 16.0d);
                cursor = script.end();
            }
        }
        lines.addAll(overlines);
        return sawOverline && !placed.isEmpty() ? new FormulaLayout(placed, lines, Math.max(1.0d, x), height) : null;
    }

    private static FormulaLayout layoutBoxedAwarePart(String text) {
        if (text != null && text.contains("\\boxed")) {
            FormulaLayout boxed = layoutBoxedFragments(text);
            if (boxed != null) {
                return boxed;
            }
        }
        return layoutFractionPart(text);
    }

    private static FormulaLayout layoutSqrt(String text) {
        text = unwrapWholeGroup(text);
        String marker = "\\sqrt";
        if (!text.contains(marker)) {
            return null;
        }
        List<PlacedText> placed = new ArrayList<>();
        List<LineSegment> lines = new ArrayList<>();
        double x = 0.0d;
        double height = 16.0d;
        int cursor = 0;
        boolean sawSqrt = false;
        while (cursor < text.length()) {
            int start = text.indexOf(marker, cursor);
            if (start < 0) {
                FormulaLayout suffix = layoutFractionPart(text.substring(cursor));
                if (suffix == null) {
                    return null;
                }
                appendLayout(placed, lines, suffix, x, 0.0d);
                x += suffix.widthPt();
                height = Math.max(height, suffix.heightPt());
                break;
            }
            if (start > cursor) {
                FormulaLayout prefix = layoutFractionPart(text.substring(cursor, start));
                if (prefix == null) {
                    return null;
                }
                appendLayout(placed, lines, prefix, x, 0.0d);
                x += prefix.widthPt();
                height = Math.max(height, prefix.heightPt());
            }
            int groupStart = skipWhitespaceForward(text, start + marker.length());
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                return null;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return null;
            }
            FormulaLayout body = layoutFractionPart(text.substring(groupStart + 1, groupEnd));
            if (body == null) {
                return null;
            }
            double rootX = x;
            appendLayout(placed, lines, body, rootX + 6.0d, 1.2d);
            double width = body.widthPt() + 7.0d;
            double rootHeight = Math.max(body.heightPt() + 2.0d, 14.0d);
            lines.add(new LineSegment(rootX, rootHeight * 0.62d, rootX + 2.0d, rootHeight - 1.0d));
            lines.add(new LineSegment(rootX + 2.0d, rootHeight - 1.0d, rootX + 5.0d, 2.0d));
            lines.add(new LineSegment(rootX + 5.0d, 2.0d, rootX + width, 2.0d));
            x += width;
            height = Math.max(height, rootHeight);
            sawSqrt = true;
            cursor = groupEnd + 1;
        }
        return sawSqrt && (!placed.isEmpty() || !lines.isEmpty()) ? new FormulaLayout(placed, lines, Math.max(x, 1.0d), height) : null;
    }

    private static String unwrapWholeGroup(String text) {
        if (text == null || text.length() < 2 || text.charAt(0) != '{') {
            return text;
        }
        int end = findGroupEnd(text, 0);
        return end == text.length() - 1 ? text.substring(1, end).trim() : text;
    }

    private static FormulaLayout layoutOverarc(String text) {
        List<String> markers = List.of("\\overarc", "\\wideparen", "\\arc");
        if (markers.stream().noneMatch(text::contains)) {
            return null;
        }
        String flattened = text;
        List<LineSegment> arcs = new ArrayList<>();
        int searchFrom = 0;
        while (true) {
            MarkerHit hit = findNextMarker(flattened, markers, searchFrom);
            if (hit == null) {
                break;
            }
            int groupStart = skipWhitespaceForward(flattened, hit.start() + hit.marker().length());
            if (groupStart >= flattened.length() || flattened.charAt(groupStart) != '{') {
                return null;
            }
            int groupEnd = findGroupEnd(flattened, groupStart);
            if (groupEnd < 0) {
                return null;
            }
            String before = flattened.substring(0, hit.start());
            String body = flattened.substring(groupStart + 1, groupEnd);
            List<TextRun> bodyRuns = tokenizeFlat(body);
            if (bodyRuns == null || bodyRuns.isEmpty()) {
                return null;
            }
            double x = measureTextRuns(tokenizeFlat(before));
            double width = measureTextRuns(bodyRuns);
            double mid = x + width / 2.0d;
            double rise = Math.min(2.8d, Math.max(1.2d, width / 5.0d));
            arcs.add(new LineSegment(x, 3.0d, mid, 3.0d - rise));
            arcs.add(new LineSegment(mid, 3.0d - rise, x + width, 3.0d));
            flattened = before + body + flattened.substring(groupEnd + 1);
            searchFrom = hit.start() + body.length();
        }
        List<TextRun> allRuns = tokenizeFlat(flattened);
        if (allRuns == null) {
            return null;
        }
        FormulaLayout layout = layoutFlatRuns(allRuns);
        List<LineSegment> lines = new ArrayList<>(layout.lines());
        lines.addAll(arcs);
        return new FormulaLayout(layout.runs(), lines, layout.widthPt(), Math.max(layout.heightPt(), 15.0d));
    }

    private static MarkerHit findNextMarker(String text, List<String> markers, int start) {
        MarkerHit best = null;
        for (String marker : markers) {
            int index = text.indexOf(marker, start);
            if (index >= 0 && (best == null || index < best.start())) {
                best = new MarkerHit(marker, index);
            }
        }
        return best;
    }

    private static double measureTextRuns(List<TextRun> runs) {
        if (runs == null || runs.isEmpty()) {
            return 0.0d;
        }
        return layoutFlatRuns(runs).widthPt();
    }

    private static String stripMetricsAndStyles(String latex) {
        String text = latex.strip();
        boolean changed;
        do {
            changed = false;
            if (text.startsWith("\\pwmetrics{")) {
                int end = text.indexOf('}');
                if (end > 0) {
                    text = text.substring(end + 1).stripLeading();
                    changed = true;
                }
            }
            if (text.startsWith("\\pwstyle{")) {
                int end = text.indexOf('}');
                if (end > 0) {
                    text = text.substring(end + 1).stripLeading();
                    changed = true;
                }
            }
        } while (changed);
        return text;
    }

    private static Command readCommand(String text, int start) {
        int end = start + 1;
        while (end < text.length() && Character.isLetter(text.charAt(end))) {
            end++;
        }
        String name = text.substring(start + 1, end);
        if (name.isEmpty() && end < text.length()) {
            char symbol = text.charAt(end);
            String mappedSymbol = switch (symbol) {
                case '|' -> "|";
                case '_' -> "_";
                case '&' -> "&";
                case ':' -> " ";
                case '{' -> "{";
                case '}' -> "}";
                case '%' -> "%";
                default -> null;
            };
            return mappedSymbol == null ? null : new Command(mappedSymbol, end + 1);
        }
        String mapped = switch (name) {
            case "times" -> "×";
            case "div" -> "÷";
            case "le", "leq" -> "≤";
            case "ge", "geq" -> "≥";
            case "neq", "ne" -> "≠";
            case "approx" -> "≈";
            case "equiv" -> "≡";
            case "prec" -> "≺";
            case "succ" -> "≻";
            case "preceq" -> "⪯";
            case "succeq" -> "⪰";
            case "pm" -> "±";
            case "mp" -> "∓";
            case "lt" -> "<";
            case "gt" -> ">";
            case "cdot", "spot" -> "·";
            case "colon" -> ":";
            case "ast", "star" -> "∗";
            case "cdots", "ldots" -> "⋯";
            case "vdots" -> "⋮";
            case "parallel" -> "∥";
            case "nparallel" -> "∦";
            case "angle" -> "∠";
            case "measuredangle" -> "∡";
            case "sphericalangle" -> "∢";
            case "ulcorner" -> "⌜";
            case "urcorner" -> "⌝";
            case "llcorner" -> "⌞";
            case "lrcorner" -> "⌟";
            case "prime" -> "′";
            case "dprime" -> "″";
            case "trprime" -> "‴";
            case "euro" -> "€";
            case "sin" -> "sin";
            case "cos" -> "cos";
            case "tan" -> "tan";
            case "cot" -> "cot";
            case "sec" -> "sec";
            case "csc" -> "csc";
            case "arcsin" -> "arcsin";
            case "arccos" -> "arccos";
            case "arctan" -> "arctan";
            case "sinh" -> "sinh";
            case "cosh" -> "cosh";
            case "tanh" -> "tanh";
            case "log" -> "log";
            case "ln" -> "ln";
            case "exp" -> "exp";
            case "lim" -> "lim";
            case "max" -> "max";
            case "min" -> "min";
            case "km" -> "km";
            case "m" -> "m";
            case "cm" -> "cm";
            case "pi", "uppi" -> "π";
            case "nabla" -> "∇";
            case "Delta", "Updelta" -> "Δ";
            case "Theta", "Uptheta" -> "Θ";
            case "circ" -> "°";
            case "sim" -> "~";
            case "backsim" -> "∽";
            case "mid" -> "∣";
            case "nmid" -> "∤";
            case "in" -> "∈";
            case "notin" -> "∉";
            case "ni" -> "∋";
            case "subset" -> "⊂";
            case "supset" -> "⊃";
            case "subseteq" -> "⊆";
            case "supseteq" -> "⊇";
            case "cup" -> "∪";
            case "cap" -> "∩";
            case "emptyset", "varnothing" -> "∅";
            case "perp", "bot" -> "⊥";
            case "therefore" -> "∴";
            case "because" -> "∵";
            case "bigcirc", "Circle" -> "○";
            case "CIRCLE" -> "●";
            case "Sun" -> "☉";
            case "oplus" -> "⊕";
            case "ominus" -> "⊖";
            case "otimes" -> "⊗";
            case "odot" -> "⊙";
            case "oslash" -> "⊘";
            case "square" -> "□";
            case "blacksquare" -> "■";
            case "Diamond", "whitediamond" -> "◇";
            case "Diamondblack", "blackdiamond" -> "◆";
            case "bigstar" -> "★";
            case "whitestar" -> "☆";
            case "vartriangle", "bigtriangleup", "triangle" -> "△";
            case "bigtriangledown", "triangledown" -> "▽";
            case "to", "rightarrow" -> "→";
            case "leftarrow", "gets" -> "←";
            case "uparrow" -> "↑";
            case "downarrow" -> "↓";
            case "updownarrow" -> "↕";
            case "leftrightarrow" -> "↔";
            case "Rightarrow" -> "⇒";
            case "Leftarrow" -> "⇐";
            case "Uparrow" -> "⇑";
            case "Downarrow" -> "⇓";
            case "Leftrightarrow" -> "⇔";
            default -> null;
        };
        if (mapped == null) {
            return null;
        }
        return new Command(mapped, end);
    }

    private static boolean isCjk(String value) {
        for (int i = 0; i < value.length(); i++) {
            Character.UnicodeBlock block = Character.UnicodeBlock.of(value.charAt(i));
            if (block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION
                || block == Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS) {
                return true;
            }
        }
        return false;
    }

    private static double estimatedWidthPt(List<TextRun> runs) {
        double width = 0d;
        for (TextRun run : runs) {
            width += leadingMathSpacingPt(run, false) + estimatedRunWidthPt(run) + trailingMathSpacingPt(run, false);
        }
        return width;
    }

    private static double placeRuns(List<PlacedText> placed, List<TextRun> runs, double x, double baseline,
        boolean script) {
        double cursor = x;
        for (TextRun run : runs) {
            cursor += leadingMathSpacingPt(run, script);
            placed.add(new PlacedText(run.text(), run.cjk(), script, false, cursor, baseline,
                runWidthScale(run, script)));
            cursor += scaledRunWidthPt(run, script) + trailingMathSpacingPt(run, script);
        }
        return cursor;
    }

    private static int nextScriptOperator(String text, int start) {
        for (int i = start; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '^' || ch == '_') {
                return i;
            }
        }
        return -1;
    }

    private static int skipWhitespaceBackward(String text, int start) {
        int i = start;
        while (i > 0 && Character.isWhitespace(text.charAt(i - 1))) {
            i--;
        }
        return i;
    }

    private static int skipWhitespaceForward(String text, int start) {
        int i = start;
        while (i < text.length() && Character.isWhitespace(text.charAt(i))) {
            i++;
        }
        return i;
    }

    private static int findAtomStart(String text, int atomEnd) {
        int i = atomEnd - 1;
        if (i < 0) {
            return atomEnd;
        }
        char ch = text.charAt(i);
        if (ch == '}') {
            int depth = 1;
            i--;
            while (i >= 0) {
                char current = text.charAt(i);
                if (current == '}') {
                    depth++;
                } else if (current == '{') {
                    depth--;
                    if (depth == 0) {
                        return i;
                    }
                }
                i--;
            }
            return atomEnd;
        }
        if (Character.isLetterOrDigit(ch) || ch > 127) {
            return i;
        }
        return i;
    }

    private static int findSimpleScriptAtomEnd(String text, int start) {
        if (start >= text.length()) {
            return -1;
        }
        if (text.charAt(start) == '\\') {
            int end = start + 1;
            while (end < text.length() && Character.isLetter(text.charAt(end))) {
                end++;
            }
            return end > start + 1 ? end : Math.min(text.length(), start + 2);
        }
        return start + 1;
    }

    private static int findGroupEnd(String text, int groupStart) {
        int depth = 0;
        for (int i = groupStart; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '{') {
                depth++;
            } else if (ch == '}') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static ArraySlice readArraySlice(String text, int begin) {
        String beginToken = "\\begin{array}";
        String endToken = "\\end{array}";
        int bodyStart;
        if (text.startsWith(beginToken, begin)) {
            int specStart = skipWhitespaceForward(text, begin + beginToken.length());
            if (specStart >= text.length() || text.charAt(specStart) != '{') {
                return null;
            }
            int specEnd = findGroupEnd(text, specStart);
            if (specEnd < 0) {
                return null;
            }
            bodyStart = specEnd + 1;
        } else {
            String env = mathArrayEnvironmentAt(text, begin, true);
            if (env == null || "array".equals(env) || "cases".equals(env)) {
                return null;
            }
            beginToken = "\\begin{" + env + "}";
            endToken = "\\end{" + env + "}";
            bodyStart = begin + beginToken.length();
        }
        return readDelimitedArrayBody(text, bodyStart, beginToken, endToken);
    }

    private static ArraySlice readDelimitedArrayBody(String text, int bodyStart, String beginToken, String endToken) {
        if (bodyStart < 0 || bodyStart > text.length()) {
            return null;
        }
        int depth = 1;
        int cursor = bodyStart;
        while (cursor < text.length()) {
            int nextBegin = text.indexOf(beginToken, cursor);
            int nextEnd = text.indexOf(endToken, cursor);
            if (nextEnd < 0) {
                return null;
            }
            if (nextBegin >= 0 && nextBegin < nextEnd) {
                depth++;
                cursor = nextBegin + beginToken.length();
                continue;
            }
            depth--;
            if (depth == 0) {
                return new ArraySlice(text.substring(bodyStart, nextEnd).trim(), nextEnd + endToken.length());
            }
            cursor = nextEnd + endToken.length();
        }
        return null;
    }

    private static boolean startsAnyMathArrayBegin(String text, int index) {
        return mathArrayEnvironmentAt(text, index, true) != null;
    }

    private static boolean startsAnyMathArrayEnd(String text, int index) {
        return mathArrayEnvironmentAt(text, index, false) != null;
    }

    private static String mathArrayEnvironmentAt(String text, int index, boolean begin) {
        String prefix = begin ? "\\begin{" : "\\end{";
        if (!text.startsWith(prefix, index)) {
            return null;
        }
        int nameStart = index + prefix.length();
        int nameEnd = text.indexOf('}', nameStart);
        if (nameEnd < 0) {
            return null;
        }
        String env = text.substring(nameStart, nameEnd);
        return switch (env) {
            case "array", "aligned", "alignedat", "gathered", "matrix", "pmatrix", "bmatrix", "cases" -> env;
            default -> null;
        };
    }

    private static ArraySlice unwrapSingleCasesArray(String body) {
        String trimmed = body.trim();
        if (trimmed.endsWith("\\\\")) {
            trimmed = trimmed.substring(0, trimmed.length() - 2).trim();
        }
        ArraySlice cases = readCasesSlice(trimmed, 0);
        return cases != null && cases.end() == trimmed.length() ? cases : null;
    }

    private static ArraySlice readCasesSlice(String text, int begin) {
        String beginToken = "\\begin{cases}";
        if (!text.startsWith(beginToken, begin)) {
            return null;
        }
        String endToken = "\\end{cases}";
        int depth = 1;
        int cursor = begin + beginToken.length();
        while (cursor < text.length()) {
            int nextBegin = text.indexOf(beginToken, cursor);
            int nextEnd = text.indexOf(endToken, cursor);
            if (nextEnd < 0) {
                return null;
            }
            if (nextBegin >= 0 && nextBegin < nextEnd) {
                depth++;
                cursor = nextBegin + beginToken.length();
                continue;
            }
            depth--;
            if (depth == 0) {
                return new ArraySlice(text.substring(begin + beginToken.length(), nextEnd).trim(),
                    nextEnd + endToken.length());
            }
            cursor = nextEnd + endToken.length();
        }
        return null;
    }

    private static double estimatedRunWidthPt(TextRun run) {
        if (run.cjk()) {
            return measureTextPt(CJK_FONT, run.text(), 12.0d);
        }
        double width = 0d;
        for (EncodedText segment : encodeText(run.text(), false)) {
            width += segment.widthPt();
        }
        return width;
    }

    private static List<EncodedText> encodeText(String text, boolean cjk) {
        if (text == null || text.isEmpty()) {
            return List.of();
        }
        List<EncodedText> segments = new ArrayList<>();
        TextKind activeKind = null;
        StringBuilder active = new StringBuilder();
        for (int offset = 0; offset < text.length();) {
            int codePoint = text.codePointAt(offset);
            EncodableChar encodable = encodableChar(codePoint, cjk);
            TextKind kind = encodable.kind();
            String output = encodable.text();
            if (activeKind != null && activeKind != kind) {
                addEncodedSegment(segments, activeKind, active.toString());
                active.setLength(0);
            }
            activeKind = kind;
            active.append(output);
            offset += Character.charCount(codePoint);
        }
        if (activeKind != null && active.length() > 0) {
            addEncodedSegment(segments, activeKind, active.toString());
        }
        return segments;
    }

    private static void addEncodedSegment(List<EncodedText> segments, TextKind kind, String text) {
        Charset charset = kind == TextKind.CJK ? GBK : WINDOWS_1252;
        byte[] bytes = text.getBytes(charset);
        Font font = switch (kind) {
            case ANSI -> TIMES_FONT;
            case SYMBOL -> SYMBOL_FONT;
            case CJK -> CJK_FONT;
        };
        segments.add(new EncodedText(kind, text, bytes, measureTextPt(font, text, 12.0d)));
    }

    private static EncodableChar encodableChar(int codePoint, boolean forceCjk) {
        String value = new String(Character.toChars(codePoint));
        if (forceCjk || isCjk(value)) {
            return new EncodableChar(TextKind.CJK, gbkSafeText(codePoint));
        }
        int symbol = symbolByte(codePoint);
        if (symbol >= 0) {
            return new EncodableChar(TextKind.SYMBOL, String.valueOf((char) symbol));
        }
        String ansi = ansiSafeText(codePoint);
        if (canEncode(WINDOWS_1252, ansi)) {
            return new EncodableChar(TextKind.ANSI, ansi);
        }
        String gbk = gbkSafeText(codePoint);
        if (canEncode(GBK, gbk)) {
            return new EncodableChar(TextKind.CJK, gbk);
        }
        return new EncodableChar(TextKind.ANSI, asciiFallbackText(codePoint));
    }

    private static boolean canEncode(Charset charset, String text) {
        return charset.newEncoder().canEncode(text);
    }

    private static String ansiSafeText(int codePoint) {
        return switch (codePoint) {
            case 0x2212 -> "-";
            case 0x2018, 0x2019 -> "'";
            case 0x201C, 0x201D -> "\"";
            case 0x2026 -> "...";
            default -> new String(Character.toChars(codePoint));
        };
    }

    private static String gbkSafeText(int codePoint) {
        String value = new String(Character.toChars(codePoint));
        return canEncode(GBK, value) ? value : asciiFallbackText(codePoint);
    }

    private static String asciiFallbackText(int codePoint) {
        return switch (codePoint) {
            case 0x00B7, 0x2219, 0x2217 -> "*";
            case 0x2026, 0x22EF -> "...";
            case 0x22EE -> ":";
            case 0x2032 -> "'";
            case 0x2033 -> "''";
            case 0x2034 -> "'''";
            case 0x2205 -> "O";
            case 0x220A, 0x2AAF -> "<=";
            case 0x220B, 0x2AB0 -> ">=";
            case 0x2213 -> "-/+";
            case 0x2235 -> "...";
            case 0x2224 -> "|";
            case 0x2225 -> "||";
            case 0x2226 -> "||";
            case 0x25A1, 0x25A0 -> "[]";
            case 0x25CB, 0x25CF, 0x2609 -> "o";
            case 0x25B3, 0x25BD -> "△";
            case 0x25C7, 0x25C6 -> "<>";
            case 0x2605, 0x2606 -> "*";
            case 0x2285, 0x2296, 0x2297, 0x2299, 0x2298 -> "o";
            default -> "#";
        };
    }

    private static int symbolByte(int codePoint) {
        return switch (codePoint) {
            case 0x00B0 -> 0xB0; // degree
            case 0x00B1 -> 0xB1; // plus-minus
            case 0x00B7 -> 0xD7; // dot operator in Symbol
            case 0x00D7 -> 0xB4; // times
            case 0x00F7 -> 0xB8; // division
            case 0x03C0 -> 0x70; // pi
            case 0x0394 -> 0x44; // Delta
            case 0x0398 -> 0x51; // Theta
            case 0x2207 -> 0xD1; // nabla
            case 0x2208 -> 0xCE; // element of
            case 0x2209 -> 0xCF; // not element of
            case 0x220B -> 0x27; // contains as member
            case 0x2218 -> 0xB0; // ring operator fallback
            case 0x221A -> 0xD6; // radical
            case 0x221E -> 0xA5; // infinity
            case 0x2220 -> 0xD0; // angle
            case 0x2223 -> 0xBD; // divides
            case 0x2229 -> 0xC7; // cap
            case 0x222A -> 0xC8; // cup
            case 0x2234 -> 0x5C; // therefore
            case 0x2248 -> 0xBB; // approx
            case 0x2260 -> 0xB9; // not equal
            case 0x2261 -> 0xBA; // equivalent
            case 0x2264 -> 0xA3; // less equal
            case 0x2265 -> 0xB3; // greater equal
            case 0x2282 -> 0xCC; // subset
            case 0x2283 -> 0xC9; // superset
            case 0x2286 -> 0xCD; // subset equal
            case 0x2287 -> 0xCA; // superset equal
            case 0x22A5 -> 0x5E; // perpendicular
            case 0x2190 -> 0xAC; // left arrow
            case 0x2191 -> 0xAD; // up arrow
            case 0x2192 -> 0xAE; // right arrow
            case 0x2193 -> 0xAF; // down arrow
            case 0x2194 -> 0xAB; // left right arrow
            case 0x21D0 -> 0xDC; // double left arrow
            case 0x21D1 -> 0xDD; // double up arrow
            case 0x21D2 -> 0xDE; // double right arrow
            case 0x21D3 -> 0xDF; // double down arrow
            case 0x21D4 -> 0xDB; // double left right arrow
            default -> -1;
        };
    }

    private static double measureTextPt(Font font, String text, double sizePt) {
        if (text == null || text.isEmpty()) {
            return 0.0d;
        }
        Font sized = font.deriveFont((float) sizePt);
        return sized.getStringBounds(text, FONT_RENDER_CONTEXT).getWidth() * 72.0d / 96.0d;
    }

    private static int fontIndex(TextKind kind, boolean script, boolean display) {
        int sizeIndex = display ? 2 : (script ? 1 : 0);
        return switch (kind) {
            case ANSI -> sizeIndex;
            case SYMBOL -> 3 + sizeIndex;
            case CJK -> 6 + sizeIndex;
        };
    }

    private static double sizeScale(boolean script, boolean display) {
        if (display) {
            return 22.0d / 12.0d;
        }
        return script ? 8.0d / 12.0d : 1.0d;
    }

    private static double scaledRunWidthPt(TextRun run, boolean script) {
        double width = estimatedRunWidthPt(run);
        return script ? width * 0.66d : width;
    }

    private static double runWidthScale(TextRun run, boolean script) {
        if (script || run == null || run.cjk()) {
            return 1.0d;
        }
        return shortGeometryLabelWidthScale(run.text());
    }

    private static double shortGeometryLabelWidthScale(String text) {
        if (text == null) {
            return 1.0d;
        }
        String value = text.trim();
        if (value.length() < 2 || value.length() > 4 || !value.chars().allMatch(ch -> ch >= 'A' && ch <= 'Z')) {
            return 1.0d;
        }
        if (value.length() == 2) {
            return 0.82d;
        }
        return SHORT_GEOMETRY_LABEL_WIDTH_SCALE;
    }

    private static double leadingMathSpacingPt(TextRun run, boolean script) {
        if (!isMathBinaryOperator(run.text())) {
            return 0.0d;
        }
        return script ? 0.05d : 0.1d;
    }

    private static double trailingMathSpacingPt(TextRun run, boolean script) {
        if (!isMathBinaryOperator(run.text())) {
            return 0.0d;
        }
        return script ? 0.05d : 0.1d;
    }

    private static boolean isMathBinaryOperator(String text) {
        return "+-−=×÷*/<>≤≥≠".contains(text);
    }

    private static int[] characterDxTwips(EncodedText segment, boolean script, boolean display, double previewScale,
        double shortScriptScale, double runWidthScale) {
        if (segment.text().isEmpty() || segment.bytes().length == 0) {
            return null;
        }
        Font font = switch (segment.kind()) {
            case ANSI -> TIMES_FONT;
            case SYMBOL -> SYMBOL_FONT;
            case CJK -> CJK_FONT;
        };
        int[] dx = new int[segment.bytes().length];
        int byteIndex = 0;
        Charset charset = segment.kind() == TextKind.CJK ? GBK : WINDOWS_1252;
        for (int offset = 0; offset < segment.text().length();) {
            int codePoint = segment.text().codePointAt(offset);
            String ch = new String(Character.toChars(codePoint));
            double widthPt = measureTextPt(font, ch, 12.0d) * sizeScale(script, display)
                * (script ? shortScriptScale : 1.0d) * runWidthScale * previewScale;
            byte[] charBytes = ch.getBytes(charset);
            if (byteIndex >= dx.length) {
                break;
            }
            dx[byteIndex++] = Math.max(1, toTwips(widthPt));
            for (int i = 1; i < charBytes.length && byteIndex < dx.length; i++) {
                dx[byteIndex++] = 0;
            }
            offset += Character.charCount(codePoint);
        }
        return dx;
    }

    private static void writeExtTextOut(ByteArrayOutputStream out, int x, int y, byte[] text, int[] dx)
        throws IOException {
        writeShort(out, y);
        writeShort(out, x);
        writeShort(out, text.length);
        writeWord(out, 0);
        out.write(text);
        if ((text.length & 1) == 1) {
            out.write(0);
        }
        if (dx != null && dx.length == text.length) {
            for (int value : dx) {
                writeShort(out, value);
            }
        }
    }

    private static void writeBlackPen(ByteArrayOutputStream out) throws IOException {
        writeWord(out, 0); // PS_SOLID
        writeShort(out, Math.max(1, toTwips(0.45d)));
        writeShort(out, 0);
        writeDWord(out, 0);
    }

    private static void writePolyline(ByteArrayOutputStream out, int x1, int y1, int x2, int y2)
        throws IOException {
        writeWord(out, 2);
        writeShort(out, x1);
        writeShort(out, y1);
        writeShort(out, x2);
        writeShort(out, y2);
    }

    private static void writeFont(ByteArrayOutputStream out, String face, double sizePt, int charset) throws IOException {
        writeFont(out, face, sizePt, charset, 1.0d);
    }

    private static void writeFont(ByteArrayOutputStream out, String face, double sizePt, int charset,
        double widthScale) throws IOException {
        writeShort(out, -toTwips(sizePt));
        writeShort(out, Math.abs(widthScale - 1.0d) > 0.001d
            ? Math.max(1, toTwips(sizePt * widthScale * 0.45d)) : 0);
        writeShort(out, 0);
        writeShort(out, 0);
        writeWord(out, 400);
        out.write(0); // italic
        out.write(0); // underline
        out.write(0); // strikeout
        out.write(charset);
        out.write(0); // out precision
        out.write(0); // clip precision
        out.write(0); // quality
        out.write(0); // pitch/family
        byte[] faceBytes = face.getBytes(StandardCharsets.US_ASCII);
        out.write(faceBytes);
        out.write(0);
        for (int i = faceBytes.length + 1; i < 32; i++) {
            out.write(0);
        }
    }

    private static int toTwips(double pt) {
        return Math.max(1, (int) Math.round(pt * TWIPS_PER_POINT));
    }

    private static void writePlaceableHeader(ByteArrayOutputStream out, double widthPt, double heightPt)
        throws IOException {
        int inch = placeableUnitsPerInch(widthPt, heightPt);
        int right = Math.max((int) Math.round(widthPt / 72.0d * inch), 1);
        int bottom = Math.max((int) Math.round(heightPt / 72.0d * inch), 1);
        writeDWord(out, PLACEABLE_WMF_KEY);
        writeWord(out, 0);
        writeShort(out, 0);
        writeShort(out, 0);
        writeShort(out, right);
        writeShort(out, bottom);
        writeWord(out, inch);
        writeDWord(out, 0);
        byte[] bytes = out.toByteArray();
        int checksum = 0;
        for (int i = 0; i < 20; i += 2) {
            checksum ^= (bytes[i] & 0xff) | ((bytes[i + 1] & 0xff) << 8);
        }
        writeWord(out, checksum);
    }

    private static int placeableUnitsPerInch(double widthPt, double heightPt) {
        double maxInches = Math.max(widthPt, heightPt) / 72.0d;
        if (maxInches <= 0d) {
            return 1440;
        }
        int maxUnits = (int) Math.floor(32760.0d / maxInches);
        int[] candidates = {1440, 720, 360, 180, 120, 96, 72};
        for (int candidate : candidates) {
            if (candidate <= maxUnits) {
                return candidate;
            }
        }
        return Math.max(maxUnits, 1);
    }

    private static void writeWord(ByteArrayOutputStream out, int value) throws IOException {
        out.write(value & 0xff);
        out.write((value >>> 8) & 0xff);
    }

    private static void writeShort(ByteArrayOutputStream out, int value) throws IOException {
        writeWord(out, value & 0xffff);
    }

    private static void writeDWord(ByteArrayOutputStream out, long value) throws IOException {
        writeWord(out, (int) (value & 0xffff));
        writeWord(out, (int) ((value >>> 16) & 0xffff));
    }

    private record TextRun(String text, boolean cjk) {
    }

    private enum TextKind {
        ANSI,
        SYMBOL,
        CJK
    }

    private record EncodedText(TextKind kind, String text, byte[] bytes, double widthPt) {
    }

    private record PreviewScale(double x, double y) {
    }

    private record SegmentScale(double rightPt) {
    }

    private record EncodableChar(TextKind kind, String text) {
    }

    private record PlacedText(String text, boolean cjk, boolean script, boolean display, double xPt,
                              double baselinePt, double widthScale) {
        private PlacedText(String text, boolean cjk, boolean script, boolean display, double xPt,
                           double baselinePt) {
            this(text, cjk, script, display, xPt, baselinePt, 1.0d);
        }
    }

    private record LineSegment(double x1Pt, double y1Pt, double x2Pt, double y2Pt) {
    }

    private record MarkerHit(String marker, int start) {
    }

    private record ScriptGroup(char operator, String body, int end) {
    }

    private record FormulaLayout(List<PlacedText> runs, List<LineSegment> lines, double widthPt, double heightPt) {
        private FormulaLayout(List<PlacedText> runs, double widthPt, double heightPt) {
            this(runs, List.of(), widthPt, heightPt);
        }
    }

    private record Command(String text, int end) {
    }

    private record ArraySlice(String body, int end) {
    }

    private static final class WmfBuilder {
        private final ByteArrayOutputStream records = new ByteArrayOutputStream();
        private final double widthPt;
        private final double heightPt;
        private int maxRecordWords;
        private int objectCount;

        private WmfBuilder(double widthPt, double heightPt) {
            this.widthPt = widthPt;
            this.heightPt = heightPt;
        }

        private void record(int function, PayloadWriter writer) throws IOException {
            ByteArrayOutputStream payload = new ByteArrayOutputStream();
            writer.write(payload);
            byte[] bytes = payload.toByteArray();
            int words = 3 + ((bytes.length + 1) / 2);
            writeDWord(records, words);
            writeWord(records, function);
            records.write(bytes);
            if ((bytes.length & 1) == 1) {
                records.write(0);
            }
            if (function == 0x02FB || function == 0x02FA) {
                objectCount++;
            }
            maxRecordWords = Math.max(maxRecordWords, words);
        }

        private byte[] finish() throws IOException {
            record(0x0000, out -> { });
            byte[] recordBytes = records.toByteArray();
            int fileSizeWords = (18 + recordBytes.length) / 2;
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            writePlaceableHeader(out, widthPt, heightPt);
            writeWord(out, 1);
            writeWord(out, 9);
            writeWord(out, 0x0300);
            writeDWord(out, fileSizeWords);
            writeWord(out, Math.max(1, objectCount));
            writeDWord(out, maxRecordWords);
            writeWord(out, 0);
            out.write(recordBytes);
            return out.toByteArray();
        }
    }

    @FunctionalInterface
    private interface PayloadWriter {
        void write(ByteArrayOutputStream out) throws IOException;
    }
}
