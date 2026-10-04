package com.lz.paperword.core.docx;

import org.jsoup.Jsoup;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Conservative, opt-in soft wrapping for question stems. Formula source is opaque. */
final class ContentFlowLayout {
    private static final Pattern BOUNDARY = Pattern.compile(
        "(?is)<(?:p|div)\\b[^>]*>\\s*(?:<br\\s*/?>\\s*)?</(?:p|div)\\s*>|<table\\b[^>]*>.*?</table\\s*>|<img\\b[^>]*>|<br\\s*/?>|"
            + "</?(?:p|div|li|ul|ol|blockquote|h[1-6]|tr|td|th|hr)\\b[^>]*>|\\r\\n|[\\r\\n]");
    private static final Pattern SUBQUESTION = Pattern.compile(
        "^(?:[（(]\\s*(?:\\d+|[一二三四五六七八九十]+)\\s*[)）]|[①-⑳]|[一二三四五六七八九十]+[、．.]|第[一二三四五六七八九十\\d]+[步问]|[A-Ha-h][、．.)]|\\d+[)）]|\\d+[、．.](?!\\d)).*");
    private static final Pattern TRAILING_PUNCTUATION = Pattern.compile("[\\s\\p{P}]*");

    private ContentFlowLayout() { }

    static List<String> split(String html) {
        if (html == null || html.isBlank()) return List.of();
        ProtectedMath math = protectMath(html);
        List<String> result = new ArrayList<>();
        StringBuilder current = new StringBuilder();
        Matcher matcher = BOUNDARY.matcher(math.text());
        int offset = 0;
        int softBreaks = 0;
        boolean previousBreakWasBr = false;
        while (matcher.find()) {
            String chunk = math.text().substring(offset, matcher.start());
            if (!chunk.isBlank()) {
                appendChunk(result, current, chunk, softBreaks, math);
                softBreaks = 0;
                previousBreakWasBr = false;
            }
            String token = matcher.group();
            boolean br = token.toLowerCase(java.util.Locale.ROOT).startsWith("<br");
            boolean newline = token.charAt(0) != '<';
            if (br || newline) {
                // A source-formatting newline immediately after <br> is the same break.
                if (!(newline && previousBreakWasBr && chunk.isBlank())) softBreaks++;
                previousBreakWasBr = br;
            } else {
                flush(result, current);
                softBreaks = 0;
                previousBreakWasBr = false;
                String lower = token.toLowerCase(java.util.Locale.ROOT);
                if (lower.startsWith("<img") || lower.startsWith("<table")) result.add(token);
                else if ((lower.startsWith("<p") || lower.startsWith("<div")) && lower.contains("</")) result.add("");
                // Block tags are retained with their own boundaries. The original HTML
                // parser remains responsible for their content; no asset loading is added.
                else current.append(token);
            }
            offset = matcher.end();
        }
        appendChunk(result, current, math.text().substring(offset), softBreaks, math);
        flush(result, current);
        return result.stream().map(math::restore).toList();
    }

    private static void appendChunk(List<String> result, StringBuilder current, String chunk, int breaks, ProtectedMath math) {
        if (chunk.isBlank()) return;
        String next = chunk.trim();
        String previousText = visible(current.toString());
        String nextText = visible(next);
        if (breaks > 0 && !previousText.isEmpty()) {
            if (breaks > 1 || isStructuralLine(previousText) || startsSubquestion(nextText)
                    || math.isFormulaOnly(nextText) || math.isFormulaOnly(previousText)) {
                flush(result, current);
                // A blank source line is an intentional empty paragraph in flow mode.
                for (int i = 1; i < breaks; i++) result.add("");
            } else if (needsWordSpace(previousText, nextText)) {
                current.append(' ');
            }
        }
        current.append(next);
    }

    private static void flush(List<String> result, StringBuilder current) {
        String value = current.toString().trim();
        if (!visible(value).isEmpty()) result.add(value);
        current.setLength(0);
    }

    private static boolean startsSubquestion(String value) { return SUBQUESTION.matcher(value).matches(); }

    private static boolean isStructuralLine(String text) {
        return text.length() <= 60 && (text.endsWith(":") || text.endsWith("："));
    }

    private static String visible(String html) { return Jsoup.parse(html).body().text().trim(); }

    private static boolean needsWordSpace(String left, String right) {
        if (left.isEmpty() || right.isEmpty()) return false;
        int a = left.codePointBefore(left.length()), b = right.codePointAt(0);
        if (isCjk(a) || isCjk(b)) return false;
        if ("，。；：！？、,.!?;:)]}）】》".indexOf(b) >= 0) return false;
        if ("([{（【《".indexOf(a) >= 0) return false;
        return true;
    }

    private static boolean isCjk(int codePoint) {
        Character.UnicodeScript script = Character.UnicodeScript.of(codePoint);
        return script == Character.UnicodeScript.HAN || script == Character.UnicodeScript.HIRAGANA
            || script == Character.UnicodeScript.KATAKANA || script == Character.UnicodeScript.HANGUL
            || (codePoint >= 0x3000 && codePoint <= 0x303f) || (codePoint >= 0xff00 && codePoint <= 0xffef);
    }

    /** Explicit display delimiters may be followed by punctuation without becoming inline. */
    static boolean isDisplayFormula(String html) {
        if (html == null) return false;
        ProtectedMath protectedMath = protectMath(html);
        String text = visible(protectedMath.text());
        String token = protectedMath.marker() + "0\uE101";
        if (protectedMath.formulas().size() != 1 || !text.startsWith(token)) return false;
        String formula = protectedMath.formulas().getFirst().trim();
        return (formula.startsWith("$$") || formula.startsWith("\\["))
            && TRAILING_PUNCTUATION.matcher(text.substring(token.length())).matches();
    }

    /** Reuse the delimiter shielding in legacy splitting without changing its break policy. */
    static List<String> splitLegacy(String html) {
        if (html == null || html.isBlank()) return List.of();
        ProtectedMath math = protectMath(html);
        List<String> result = new ArrayList<>();
        for (String piece : math.text().split("(?i)<br\\s*/?>|(?:\\R\\s*){2,}")) {
            if (!piece.isBlank()) result.add(math.restore(piece.trim()));
        }
        return result;
    }

    private static ProtectedMath protectMath(String input) {
        List<String> formulas = new ArrayList<>();
        String marker = "\uE100";
        while (input.contains(marker)) marker += "\uE100";
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < input.length();) {
            String end = null;
            int openLength = 0;
            if (!escaped(input, i)) {
                if (input.startsWith("$$", i)) { end = "$$"; openLength = 2; }
                else if (input.charAt(i) == '$') { end = "$"; openLength = 1; }
                else if (input.startsWith("\\[", i)) { end = "\\]"; openLength = 2; }
                else if (input.startsWith("\\(", i)) { end = "\\)"; openLength = 2; }
                else if (input.startsWith("\\begin{", i)) {
                    int close = input.indexOf('}', i + 7);
                    if (close >= 0) {
                        String environment = input.substring(i + 7, close);
                        if (environment.matches("(?:array|aligned\\*?|align\\*?|[bBpPvV]?matrix|cases|gathered|split)")) {
                            end = "\\end{" + environment + "}";
                            openLength = close - i + 1;
                        }
                    }
                }
            }
            int close = -1;
            if (end != null) {
                int scan = i + openLength;
                while ((scan = input.indexOf(end, scan)) >= 0) {
                    if (!escaped(input, scan)) { close = scan + end.length(); break; }
                    scan += end.length();
                }
            }
            if (close >= 0) {
                String formula = input.substring(i, close).replaceAll("(?i)<br\\s*/?>", " ");
                result.append(marker).append(formulas.size()).append('\uE101');
                formulas.add(formula);
                i = close;
            } else {
                result.append(input.charAt(i++));
            }
        }
        return new ProtectedMath(result.toString(), formulas, marker);
    }

    private static boolean escaped(String text, int offset) {
        int slashes = 0;
        while (offset > 0 && text.charAt(--offset) == '\\') slashes++;
        return (slashes & 1) != 0;
    }

    private record ProtectedMath(String text, List<String> formulas, String marker) {
        boolean isFormulaOnly(String value) {
            return value.trim().matches("(?:" + Pattern.quote(marker) + "\\d+\uE101[\\s\\p{P}\\p{S}]*)+");
        }
        String restore(String value) {
            Matcher matcher = Pattern.compile(Pattern.quote(marker) + "(\\d+)\uE101").matcher(value);
            StringBuffer result = new StringBuffer();
            while (matcher.find()) matcher.appendReplacement(result, Matcher.quoteReplacement(formulas.get(Integer.parseInt(matcher.group(1)))));
            matcher.appendTail(result);
            return result.toString();
        }
    }
}
