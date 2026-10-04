package com.lz.paperword.core.docx;

import org.jsoup.Jsoup;
import org.junit.jupiter.api.Test;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class ContentFlowLayoutTest {
    private List<String> text(String html) {
        return ContentFlowLayout.split(html).stream().map(value -> Jsoup.parse(value).body().text()).toList();
    }

    @Test void joinsOnlySoftProseWrapsWithLanguageAwareSpacing() {
        assertEquals(List.of("甲乙丙丁"), text("甲\n乙<br/>丙<BR />丁"));
        assertEquals(List.of("Find the total number."), text("Find the\n total<br/>number."));
        assertEquals(List.of("Hello world"), text("<b>Hello</b><br/>\n<i>world</i>"));
    }

    @Test void preservesBlankLinesAndIntentionalBlockBoundaries() {
        assertEquals(List.of("甲", "", "乙"), text("甲\n\n乙"));
        assertEquals(List.of("甲", "", "乙"), text("甲<br/><br/>乙"));
        assertEquals(List.of("甲", "乙", "丙"), text("<p>甲</p><p>乙</p><div>丙</div>"));
        assertEquals(List.of("甲", "", "乙"), text("<p>甲</p><p><br/></p><p>乙</p>"));
    }

    @Test void keepsSubquestionsAndOptionsOnTheirOwnLines() {
        for (String marker : List.of("（1）", "(2)", "1)", "①", "一、", "（二）", "A.", "B、", "第一步")) {
            assertEquals(List.of("说明", marker + "内容"), text("说明\n" + marker + "内容"), marker);
        }
        assertEquals(List.of("计算：", "先相加"), text("计算：\n先相加"));
        assertEquals(List.of("The value is 2.5 units"), text("The value is\n2.5 units"));
    }

    @Test void formulaRowsAndStandaloneLinesAreOpaque() {
        for (String[] delimiters : new String[][]{{"$", "$"}, {"$$", "$$"}, {"\\(", "\\)"}, {"\\[", "\\]"}}) {
            String formula = delimiters[0] + "\\begin{aligned}a&=1\\\\\nb&=2\\end{aligned}" + delimiters[1];
            assertEquals(List.of("题干", formula, "下一段"), ContentFlowLayout.split("题干\n" + formula + "\n下一段"));
        }
        String array = "\\begin{array}{cc}a&b\\\\\nc&d\\end{array}";
        assertEquals(List.of(array), ContentFlowLayout.split(array));
        assertEquals(List.of("$a +b$"), ContentFlowLayout.split("$a<br/>+b$"));
        assertEquals(List.of("价格\\$5 下一步"), ContentFlowLayout.split("价格\\$5\nNext").stream().map(s -> s.replace("Next", "下一步")).toList());
    }

    @Test void literalPrivateUseTokensNeverCollideWithProtectedFormulas() {
        String literal = "\uE1000\uE101";
        assertEquals(List.of(literal + " $y$"), ContentFlowLayout.split(literal + " $y$"));
        assertEquals(List.of("$" + literal + "$ $z$"), ContentFlowLayout.split("$" + literal + "$ $z$"));
    }

    @Test void imageAndTableBlocksCannotBeJoinedAcross() {
        String table = "<table><tr><td>值</td></tr></table>";
        assertEquals(List.of("前", "<img src='x.png'>", "后"), ContentFlowLayout.split("前<img src='x.png'>后"));
        assertEquals(List.of("前", table, "后"), ContentFlowLayout.split("前" + table + "后"));
    }

    @Test void trailingPunctuationKeepsDisplayClassification() {
        for (String value : List.of("$$x=1$$", "$$x=1$$。", "\\[x=1\\]，", "<p>$$x=1$$；</p>")) {
            assertTrue(ContentFlowLayout.isDisplayFormula(value), value);
        }
        assertFalse(ContentFlowLayout.isDisplayFormula("$$x=1$$，求结果"));
        assertFalse(ContentFlowLayout.isDisplayFormula("文字$$x=1$$"));
        assertFalse(ContentFlowLayout.isDisplayFormula("$x=1$。"));
    }
}
