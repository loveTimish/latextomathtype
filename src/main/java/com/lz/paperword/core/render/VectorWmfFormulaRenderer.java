package com.lz.paperword.core.render;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
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
    private static final int DEFAULT_DPI = 1440;
    private static final int MM_ANISOTROPIC = 8;
    private static final int TRANSPARENT = 1;
    private static final int TA_BASELINE = 0x0018;
    private static final Charset GBK = Charset.forName("GBK");
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
    private static final Pattern MATHRM_PATTERN = Pattern.compile("\\\\mathrm\\s*\\{\\s*([^{}]*)\\s*}");

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
        builder.record(0x02FB, out -> writeFont(out, "Times New Roman", 12.0d)); // CreateFontIndirect
        builder.record(0x02FB, out -> writeFont(out, "Times New Roman", 8.0d)); // CreateFontIndirect
        builder.record(0x02FB, out -> writeFont(out, "Times New Roman", 22.0d)); // CreateFontIndirect
        builder.record(0x02FA, VectorWmfFormulaRenderer::writeBlackPen); // CreatePen
        builder.record(0x012D, out -> writeWord(out, 0)); // SelectObject

        double widthScale = widthPt / Math.max(layout.widthPt(), 1.0d);
        widthScale = Math.max(0.70d, Math.min(1.25d, widthScale));
        double heightScale = heightPt / Math.max(layout.heightPt(), 1.0d);
        heightScale = Math.max(0.70d, Math.min(1.25d, heightScale));
        if (!layout.lines().isEmpty()) {
            builder.record(0x012D, out -> writeWord(out, 3)); // SelectObject pen
            for (LineSegment line : layout.lines()) {
                int x1 = toTwips(0.5d + line.x1Pt() * widthScale);
                int y1 = toTwips(1.0d + line.y1Pt() * heightScale);
                int x2 = toTwips(0.5d + line.x2Pt() * widthScale);
                int y2 = toTwips(1.0d + line.y2Pt() * heightScale);
                builder.record(0x0325, out -> writePolyline(out, x1, y1, x2, y2));
            }
            builder.record(0x012D, out -> writeWord(out, 0));
        }
        int selectedFont = 0;
        for (PlacedText run : layout.runs()) {
            int fontIndex = run.display() ? 2 : (run.script() ? 1 : 0);
            if (fontIndex != selectedFont) {
                builder.record(0x012D, out -> writeWord(out, fontIndex));
                selectedFont = fontIndex;
            }
            byte[] text = run.text().getBytes(run.cjk() ? GBK : StandardCharsets.ISO_8859_1);
            final int runX = toTwips(0.5d + run.xPt() * widthScale);
            final int runBaseline = toTwips(1.0d + run.baselinePt() * heightScale);
            builder.record(0x0A32, out -> writeExtTextOut(out, runX, runBaseline, text));
        }
        return builder.finish();
    }

    private static FormulaLayout layout(String latex) {
        if (latex == null || latex.isBlank()) {
            return null;
        }
        String text = stripMetricsAndStyles(latex).trim();
        text = normalizeTextCommands(text);
        Matcher leftBraceArray = LEFT_BRACE_ARRAY_PATTERN.matcher(text);
        if (leftBraceArray.matches()) {
            return layoutLeftBraceArray(leftBraceArray.group(1));
        }
        Matcher array = ARRAY_PATTERN.matcher(text);
        if (array.matches()) {
            return layoutArray(array.group(1));
        }
        FormulaLayout fractions = layoutFractions(text);
        if (fractions != null) {
            return fractions;
        }
        FormulaLayout scripts = layoutScripts(text);
        if (scripts != null) {
            return scripts;
        }
        List<TextRun> runs = tokenizeFlat(text);
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
            || text.contains("\\over") || text.contains("^") || text.contains("_")) {
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
            out.add(new TextRun(String.valueOf(ch), isCjk(String.valueOf(ch))));
            i++;
        }
        return out.isEmpty() ? null : out;
    }

    private static FormulaLayout layoutFlatRuns(List<TextRun> runs) {
        List<PlacedText> placed = new ArrayList<>();
        double x = 0d;
        double baseline = 9.6d;
        for (TextRun run : runs) {
            placed.add(new PlacedText(run.text(), run.cjk(), false, false, x, baseline));
            x += estimatedRunWidthPt(run);
        }
        return new FormulaLayout(placed, Math.max(x, 1.0d), 13.0d);
    }

    private static FormulaLayout layoutScripts(String text) {
        List<PlacedText> placed = new ArrayList<>();
        double x = 0d;
        int cursor = 0;
        boolean sawScript = false;
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
            int groupStart = skipWhitespaceForward(text, op + 1);
            if (groupStart >= text.length() || text.charAt(groupStart) != '{') {
                return null;
            }
            int groupEnd = findGroupEnd(text, groupStart);
            if (groupEnd < 0) {
                return null;
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
            List<TextRun> script = tokenizeFlat(text.substring(groupStart + 1, groupEnd));
            if (base == null || script == null) {
                return null;
            }
            double baseWidth = estimatedWidthPt(base);
            x = placeRuns(placed, base, x, 9.6d, false);
            double scriptBaseline = text.charAt(op) == '^' ? 5.4d : 12.1d;
            placeRuns(placed, script, x + 0.4d, scriptBaseline, true);
            x += Math.max(0d, estimatedScriptWidthPt(script) - baseWidth) + 1.5d;
            sawScript = true;
            cursor = groupEnd + 1;
        }
        return sawScript && !placed.isEmpty() ? new FormulaLayout(placed, Math.max(x, 1.0d), 13.0d) : null;
    }

    private static FormulaLayout layoutFractions(String text) {
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
                    List<TextRun> suffix = tokenizeFlat(tail);
                    if (suffix == null) {
                        return null;
                    }
                    x = placeRuns(placed, suffix, x, 14.4d, false);
                }
                break;
            }
            String prefixText = text.substring(cursor, frac);
            if (!prefixText.isBlank()) {
                List<TextRun> prefix = tokenizeFlat(prefixText);
                if (prefix == null) {
                    return null;
                }
                x = placeRuns(placed, prefix, x, 14.4d, false);
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
            double fractionWidth = Math.max(numeratorWidth, denominatorWidth) + 3.0d;
            double fractionX = x + 1.0d;
            appendLayout(placed, lines, numerator, fractionX + (fractionWidth - numeratorWidth) / 2.0d, -1.8d);
            appendLayout(placed, lines, denominator, fractionX + (fractionWidth - denominatorWidth) / 2.0d, 10.7d);
            lines.add(new LineSegment(fractionX, 12.2d, fractionX + fractionWidth, 12.2d));
            x = fractionX + fractionWidth + 1.0d;
            sawFraction = true;
            cursor = denominatorEnd + 1;
        }
        return sawFraction && !placed.isEmpty()
            ? new FormulaLayout(placed, lines, Math.max(x, 1.0d), 25.5d)
            : null;
    }

    private static FormulaLayout layoutFractionPart(String text) {
        FormulaLayout scripts = layoutScripts(text);
        if (scripts != null) {
            return scripts;
        }
        List<TextRun> runs = tokenizeFlat(text);
        return runs == null ? null : layoutFlatRuns(runs);
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

    private static FormulaLayout layoutArray(String body) {
        String[] rowText = body.split("\\\\\\\\");
        List<List<List<TextRun>>> rows = new ArrayList<>();
        int columnCount = 0;
        for (String row : rowText) {
            String[] cells = row.split("&", -1);
            List<List<TextRun>> parsedRow = new ArrayList<>();
            for (String cell : cells) {
                String normalized = cell.replace("{}", "").trim();
                List<TextRun> runs = normalized.isEmpty() ? List.of() : tokenizeFlat(normalized);
                if (runs == null) {
                    return null;
                }
                parsedRow.add(runs);
            }
            rows.add(parsedRow);
            columnCount = Math.max(columnCount, parsedRow.size());
        }
        if (rows.isEmpty() || columnCount == 0) {
            return null;
        }
        double[] widths = new double[columnCount];
        for (List<List<TextRun>> row : rows) {
            for (int i = 0; i < row.size(); i++) {
                widths[i] = Math.max(widths[i], estimatedWidthPt(row.get(i)));
            }
        }
        for (int i = 0; i < widths.length; i++) {
            widths[i] = Math.max(widths[i], 7.0d);
        }
        double[] x = new double[columnCount];
        double cursor = 0d;
        for (int i = 0; i < columnCount; i++) {
            x[i] = cursor;
            cursor += widths[i] + 2.0d;
        }
        List<PlacedText> placed = new ArrayList<>();
        double rowHeight = rows.size() > 1 ? 12.0d : 13.0d;
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            List<List<TextRun>> row = rows.get(rowIndex);
            double baseline = 9.6d + rowIndex * rowHeight;
            for (int col = 0; col < row.size(); col++) {
                List<TextRun> runs = row.get(col);
                double runX = x[col] + Math.max(0d, widths[col] - estimatedWidthPt(runs));
                for (TextRun run : runs) {
                    placed.add(new PlacedText(run.text(), run.cjk(), false, false, runX, baseline));
                    runX += estimatedRunWidthPt(run);
                }
            }
        }
        double width = Math.max(cursor - 2.0d, 1.0d);
        double height = Math.max(13.0d, 2.5d + rows.size() * rowHeight);
        return placed.isEmpty() ? null : new FormulaLayout(placed, width, height);
    }

    private static String normalizeFlatLatex(String text) {
        Matcher matcher = LEFT_RIGHT_PAREN.matcher(text);
        StringBuffer out = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(out, Matcher.quoteReplacement("(" + matcher.group(1).trim() + ")"));
        }
        matcher.appendTail(out);
        return out.toString();
    }

    private static String normalizeTextCommands(String text) {
        Matcher matcher = MATHRM_PATTERN.matcher(text);
        StringBuffer out = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(out, Matcher.quoteReplacement(matcher.group(1).trim()));
        }
        matcher.appendTail(out);
        return out.toString();
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
        String mapped = switch (name) {
            case "times" -> "×";
            case "div" -> "÷";
            case "le", "leq" -> "≤";
            case "ge", "geq" -> "≥";
            case "neq", "ne" -> "≠";
            case "lt" -> "<";
            case "gt" -> ">";
            case "cdot" -> "·";
            case "cdots" -> "⋯";
            case "pi" -> "π";
            case "circ" -> "°";
            case "sim" -> "~";
            case "bigcirc" -> "○";
            case "square" -> "□";
            case "vartriangle" -> "△";
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
            width += estimatedRunWidthPt(run);
        }
        return width;
    }

    private static double estimatedScriptWidthPt(List<TextRun> runs) {
        return estimatedWidthPt(runs) * 0.66d;
    }

    private static double placeRuns(List<PlacedText> placed, List<TextRun> runs, double x, double baseline,
        boolean script) {
        double cursor = x;
        for (TextRun run : runs) {
            placed.add(new PlacedText(run.text(), run.cjk(), script, false, cursor, baseline));
            cursor += script ? estimatedRunWidthPt(run) * 0.66d : estimatedRunWidthPt(run);
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

    private static double estimatedRunWidthPt(TextRun run) {
        if (run.cjk()) {
            return 9.0d * run.text().length();
        }
        double width = 0d;
        for (int i = 0; i < run.text().length(); i++) {
            char ch = run.text().charAt(i);
            if (Character.isDigit(ch)) {
                width += 6.8d;
            } else if (Character.isLetter(ch)) {
                width += 6.2d;
            } else if (".:".indexOf(ch) >= 0) {
                width += 3.4d;
            } else {
                width += 7.0d;
            }
        }
        return width;
    }

    private static void writeExtTextOut(ByteArrayOutputStream out, int x, int y, byte[] text) throws IOException {
        writeShort(out, y);
        writeShort(out, x);
        writeShort(out, text.length);
        writeWord(out, 0);
        out.write(text);
        if ((text.length & 1) == 1) {
            out.write(0);
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

    private static void writeFont(ByteArrayOutputStream out, String face, double sizePt) throws IOException {
        writeShort(out, -toTwips(sizePt));
        writeShort(out, 0);
        writeShort(out, 0);
        writeShort(out, 0);
        writeWord(out, 400);
        out.write(0); // italic
        out.write(0); // underline
        out.write(0); // strikeout
        out.write(0); // charset ANSI
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
        writeDWord(out, PLACEABLE_WMF_KEY);
        writeWord(out, 0);
        writeShort(out, 0);
        writeShort(out, 0);
        writeShort(out, toTwips(widthPt));
        writeShort(out, toTwips(heightPt));
        writeWord(out, DEFAULT_DPI);
        writeDWord(out, 0);
        byte[] bytes = out.toByteArray();
        int checksum = 0;
        for (int i = 0; i < 20; i += 2) {
            checksum ^= (bytes[i] & 0xff) | ((bytes[i + 1] & 0xff) << 8);
        }
        writeWord(out, checksum);
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

    private record PlacedText(String text, boolean cjk, boolean script, boolean display, double xPt,
        double baselinePt) {
    }

    private record LineSegment(double x1Pt, double y1Pt, double x2Pt, double y2Pt) {
    }

    private record FormulaLayout(List<PlacedText> runs, List<LineSegment> lines, double widthPt, double heightPt) {
        private FormulaLayout(List<PlacedText> runs, double widthPt, double heightPt) {
            this(runs, List.of(), widthPt, heightPt);
        }
    }

    private record Command(String text, int end) {
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
