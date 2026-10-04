package com.lz.paperword.core.docx;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;

/** Physical exam-paper measurements, kept separate from legacy/compact formatting. */
final class ExamTypography {
    static final int BODY_INDENT_TWIPS = 480;
    static final double BODY_SIZE_PT = 12.0;
    static final double METADATA_SIZE_PT = 9.5;
    static final double FORMULA_MAX_WIDTH_PT = (11906 - 2 * 1247 - BODY_INDENT_TWIPS) / 20.0;

    private ExamTypography() { }

    static void styleRun(XWPFRun run, double fontSizePt) {
        run.setFontSize(fontSizePt);
        run.setColor("000000");
        run.setFontFamily("Times New Roman", XWPFRun.FontCharRange.ascii);
        run.setFontFamily("Times New Roman", XWPFRun.FontCharRange.hAnsi);
        run.setFontFamily("Times New Roman", XWPFRun.FontCharRange.cs);
        run.setFontFamily("宋体", XWPFRun.FontCharRange.eastAsia);
    }

    static void bodyParagraph(XWPFParagraph paragraph, boolean metadata, boolean formulaOrImage) {
        paragraph.setSpacingAfter(0);
        lineHeight(paragraph, metadata ? 280 : 380, formulaOrImage ? STLineSpacingRule.AT_LEAST : STLineSpacingRule.EXACT);
    }

    static void indent(XWPFParagraph paragraph, boolean numbered) {
        paragraph.setIndentationLeft(BODY_INDENT_TWIPS);
        CTPPr properties = paragraph.getCTP().getPPr();
        CTInd indentation = properties.getInd();
        // firstLine and hanging are mutually exclusive. Some Word-compatible
        // renderers prioritize even firstLine=0 over a nonzero hanging indent.
        if (numbered) {
            if (indentation.isSetFirstLine()) indentation.unsetFirstLine();
            paragraph.setIndentationHanging(BODY_INDENT_TWIPS);
        } else {
            if (indentation.isSetHanging()) indentation.unsetHanging();
            paragraph.setIndentationFirstLine(0);
        }
        if (numbered) {
            CTTabs tabs = properties.isSetTabs() ? properties.getTabs() : properties.addNewTabs();
            CTTabStop tab = tabs.addNewTab();
            tab.setVal(STTabJc.LEFT);
            tab.setPos(BigInteger.valueOf(BODY_INDENT_TWIPS));
        }
    }

    static void workingLine(XWPFParagraph paragraph) {
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);
        lineHeight(paragraph, 400, STLineSpacingRule.EXACT);
    }

    private static void lineHeight(XWPFParagraph paragraph, int twips, STLineSpacingRule.Enum rule) {
        CTPPr properties = paragraph.getCTP().isSetPPr() ? paragraph.getCTP().getPPr() : paragraph.getCTP().addNewPPr();
        CTSpacing spacing = properties.isSetSpacing() ? properties.getSpacing() : properties.addNewSpacing();
        spacing.setLine(BigInteger.valueOf(twips));
        spacing.setLineRule(rule);
    }
}
