package com.lz.paperword.integration;

import com.lz.paperword.core.docx.MathTypeEmbedder;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefWriter;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlBoolean;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrGeneral;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Generates a compact K12 formula reference whose equations remain editable
 * Equation.DSMT4 objects with strict outline-only EMF+ Dual previews.
 */
class K12CommonFormulaWordAcceptanceTest {

    private static final Path OUTPUT = Path.of(System.getProperty(
        "paperword.acceptance.k12CommonWord.output",
        "target/vector-acceptance/k12-common-formulas.docx"));
    private static final int EXPECTED_FORMULAS = 78;
    private static final int TABLE_WIDTH_DXA = 9_360;
    private static final int TABLE_INDENT_DXA = 120;
    private static final int LABEL_WIDTH_DXA = 2_700;
    private static final int FORMULA_WIDTH_DXA = 6_660;
    private static final double DISPLAY_SCALE = 1.20d;
    private static final double MAX_FORMULA_WIDTH_PT = 320.0d;

    @Test
    void everyCatalogFormulaParsesAndWritesEditableMtef() throws Exception {
        List<Formula> formulas = readCatalog();
        assertEquals(EXPECTED_FORMULAS, formulas.size());
        assertEquals(6, formulas.stream().map(Formula::group).distinct().count());
        assertTrue(formulas.stream().noneMatch(formula -> formula.latex().contains("\\longdiv")),
            "K12 accepted catalog must not claim unsupported full long division");

        Set<String> ids = new LinkedHashSet<>();
        LaTeXParser parser = new LaTeXParser();
        MtefWriter writer = new MtefWriter();
        for (Formula formula : formulas) {
            assertTrue(ids.add(formula.id()), () -> "duplicate formula id: " + formula.id());
            LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(formula.latex());
            assertTrue(parsed.isSupported(), () -> formula.id() + ": " + parsed.diagnostics());
            assertTrue(writer.write(parser.parseMathIR(formula.latex())).length > 0,
                () -> formula.id() + " generated empty MTEF");
        }
    }

    @Test
    void generateK12CommonFormulaReferenceWord() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.k12CommonWord"),
            "Enable with -Dpaperword.acceptance.k12CommonWord=true");

        List<Formula> formulas = readCatalog();
        assertEquals(EXPECTED_FORMULAS, formulas.size());
        Map<String, List<Formula>> byGroup = new LinkedHashMap<>();
        for (Formula formula : formulas) {
            byGroup.computeIfAbsent(formula.group(), ignored -> new ArrayList<>()).add(formula);
        }

        LaTeXParser parser = new LaTeXParser();
        MathTypeEmbedder embedder = new MathTypeEmbedder();
        try (XWPFDocument document = new XWPFDocument()) {
            configureLetterPage(document);
            configureStyles(document);
            addRunningHeaderAndFooter(document);
            addTitleBlock(document, formulas.size());

            int groupIndex = 0;
            for (Map.Entry<String, List<Formula>> group : byGroup.entrySet()) {
                groupIndex++;
                if (groupIndex == 2 || groupIndex == 3) {
                    document.createParagraph().createRun().addBreak(BreakType.PAGE);
                }
                addGroupHeading(document, groupIndex, group.getKey());
                XWPFTable table = createFormulaTable(document, groupIndex == 3);
                for (Formula formula : group.getValue()) {
                    addFormulaRow(table, formula, parser, embedder);
                }
                XWPFParagraph spacer = document.createParagraph();
                spacer.setSpacingBefore(0);
                spacer.setSpacingAfter(40);
            }

            Files.createDirectories(OUTPUT.getParent());
            try (OutputStream output = Files.newOutputStream(OUTPUT)) {
                document.write(output);
            }
        }

        assertTrue(Files.size(OUTPUT) > 10_000);
    }

    private static List<Formula> readCatalog() throws Exception {
        InputStream input = K12CommonFormulaWordAcceptanceTest.class
            .getResourceAsStream("/k12-common-formulas.tsv");
        assertNotNull(input, "missing k12-common-formulas.tsv");
        List<Formula> formulas = new ArrayList<>();
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(input, StandardCharsets.UTF_8))) {
            String line;
            int lineNumber = 0;
            while ((line = reader.readLine()) != null) {
                lineNumber++;
                if (lineNumber == 1 || line.isBlank()) {
                    continue;
                }
                String[] fields = line.split("\\t", 4);
                assertEquals(4, fields.length, "invalid catalog line " + lineNumber);
                formulas.add(new Formula(fields[0], fields[1], fields[2], fields[3]));
            }
        }
        return formulas;
    }

    private static void configureLetterPage(XWPFDocument document) {
        CTSectPr section = document.getDocument().getBody().isSetSectPr()
            ? document.getDocument().getBody().getSectPr()
            : document.getDocument().getBody().addNewSectPr();
        CTPageSz pageSize = section.isSetPgSz() ? section.getPgSz() : section.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(12_240));
        pageSize.setH(BigInteger.valueOf(15_840));
        CTPageMar margins = section.isSetPgMar() ? section.getPgMar() : section.addNewPgMar();
        margins.setTop(BigInteger.valueOf(1_440));
        margins.setBottom(BigInteger.valueOf(1_440));
        margins.setLeft(BigInteger.valueOf(1_440));
        margins.setRight(BigInteger.valueOf(1_440));
        margins.setHeader(BigInteger.valueOf(708));
        margins.setFooter(BigInteger.valueOf(708));
    }

    private static void configureStyles(XWPFDocument document) {
        XWPFStyles styles = document.createStyles();
        addParagraphStyle(styles, "Normal", "Normal", 11.0d, "000000", false,
            0, 120, 300, true);
        addParagraphStyle(styles, "K12Body", "K12 Body", 11.0d, "000000", false,
            0, 120, 300, false);
        addParagraphStyle(styles, "K12Title", "K12 Title", 22.0d, "0B2545", true,
            0, 80, 280, false);
        addParagraphStyle(styles, "K12Subtitle", "K12 Subtitle", 10.0d, "666666", false,
            0, 220, 280, false);
        addParagraphStyle(styles, "K12Heading1", "K12 Heading 1", 16.0d, "2E74B5", true,
            360, 200, 300, false);
    }

    private static void addParagraphStyle(XWPFStyles styles, String styleId, String name,
            double sizePt, String color, boolean bold, int before, int after, int line,
            boolean defaultStyle) {
        CTStyle style = CTStyle.Factory.newInstance();
        style.setStyleId(styleId);
        style.setType(STStyleType.PARAGRAPH);
        if (defaultStyle) {
            style.setDefault(XmlBoolean.Factory.newValue(true));
        }
        CTString styleName = style.addNewName();
        styleName.setVal(name);
        CTOnOff qFormat = style.addNewQFormat();
        qFormat.setVal(XmlBoolean.Factory.newValue(true));

        CTPPrGeneral pPr = style.addNewPPr();
        CTSpacing spacing = pPr.addNewSpacing();
        spacing.setBefore(BigInteger.valueOf(before));
        spacing.setAfter(BigInteger.valueOf(after));
        spacing.setLine(BigInteger.valueOf(line));
        spacing.setLineRule(STLineSpacingRule.AUTO);

        CTRPr rPr = style.addNewRPr();
        CTFonts fonts = rPr.addNewRFonts();
        fonts.setAscii("Calibri");
        fonts.setHAnsi("Calibri");
        fonts.setCs("Calibri");
        fonts.setEastAsia("Microsoft YaHei");
        CTHpsMeasure size = rPr.addNewSz();
        size.setVal(BigInteger.valueOf(Math.round(sizePt * 2.0d)));
        CTHpsMeasure csSize = rPr.addNewSzCs();
        csSize.setVal(BigInteger.valueOf(Math.round(sizePt * 2.0d)));
        rPr.addNewColor().setVal(color);
        if (bold) {
            rPr.addNewB().setVal(XmlBoolean.Factory.newValue(true));
            rPr.addNewBCs().setVal(XmlBoolean.Factory.newValue(true));
        }
        styles.addStyle(new XWPFStyle(style));
    }

    private static void addRunningHeaderAndFooter(XWPFDocument document) {
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph headerParagraph = header.createParagraph();
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        headerParagraph.setSpacingAfter(0);
        XWPFRun headerRun = headerParagraph.createRun();
        headerRun.setText("K12 常见公式参考手册");
        formatTextRun(headerRun, 9.0d, "777777", false);

        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph footerParagraph = footer.createParagraph();
        footerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        footerParagraph.setSpacingAfter(0);
        XWPFRun footerLabel = footerParagraph.createRun();
        footerLabel.setText("第 ");
        formatTextRun(footerLabel, 9.0d, "777777", false);
        CTSimpleField pageField = footerParagraph.getCTP().addNewFldSimple();
        pageField.setInstr("PAGE");
        CTR pageRun = pageField.addNewR();
        pageRun.addNewT().setStringValue("1");
        styleFieldRun(pageRun, 9.0d, "777777");
        XWPFRun footerSuffix = footerParagraph.createRun();
        footerSuffix.setText(" 页");
        formatTextRun(footerSuffix, 9.0d, "777777", false);
    }

    private static void addTitleBlock(XWPFDocument document, int formulaCount) {
        XWPFParagraph title = document.createParagraph();
        title.setStyle("K12Title");
        title.setAlignment(ParagraphAlignment.CENTER);
        title.createRun().setText("K12 常见公式参考手册");

        XWPFParagraph subtitle = document.createParagraph();
        subtitle.setStyle("K12Subtitle");
        subtitle.setAlignment(ParagraphAlignment.CENTER);
        subtitle.createRun().setText(formulaCount
            + " 个可编辑 MathType 公式 · 小学 / 初中 / 高中 / 物理 / 化学 / 竖式");
    }

    private static void addGroupHeading(XWPFDocument document, int index, String group) {
        XWPFParagraph heading = document.createParagraph();
        heading.setStyle("K12Heading1");
        heading.setKeepNext(true);
        if (index == 3) {
            heading.setSpacingBefore(120);
            heading.setSpacingAfter(80);
        }
        heading.createRun().setText(index + ". " + group);
    }

    private static XWPFTable createFormulaTable(XWPFDocument document, boolean compactRows) {
        XWPFTable table = document.createTable(1, 2);
        table.setWidth(TABLE_WIDTH_DXA);
        table.setWidthType(TableWidthType.DXA);
        table.setTableAlignment(TableRowAlign.LEFT);
        int verticalMargin = compactRows ? 40 : 80;
        table.setCellMargins(verticalMargin, 120, verticalMargin, 120);
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "CBD5E1");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "CBD5E1");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "CBD5E1");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "CBD5E1");
        table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 2, 0, "DCE3EA");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 2, 0, "DCE3EA");

        var tableProperties = table.getCTTbl().getTblPr();
        var indent = tableProperties.isSetTblInd()
            ? tableProperties.getTblInd() : tableProperties.addNewTblInd();
        indent.setW(BigInteger.valueOf(TABLE_INDENT_DXA));
        indent.setType(STTblWidth.DXA);
        var layout = tableProperties.isSetTblLayout()
            ? tableProperties.getTblLayout() : tableProperties.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);

        var grid = table.getCTTbl().getTblGrid();
        if (grid == null) {
            grid = table.getCTTbl().addNewTblGrid();
        }
        while (grid.sizeOfGridColArray() > 0) {
            grid.removeGridCol(0);
        }
        grid.addNewGridCol().setW(BigInteger.valueOf(LABEL_WIDTH_DXA));
        grid.addNewGridCol().setW(BigInteger.valueOf(FORMULA_WIDTH_DXA));

        XWPFTableRow header = table.getRow(0);
        header.setRepeatHeader(true);
        header.setCantSplitRow(true);
        configureCell(header.getCell(0), LABEL_WIDTH_DXA, "E8EEF5");
        configureCell(header.getCell(1), FORMULA_WIDTH_DXA, "E8EEF5");
        setHeaderCellText(header.getCell(0), "公式名称");
        setHeaderCellText(header.getCell(1), "公式");
        return table;
    }

    private static void addFormulaRow(XWPFTable table, Formula formula, LaTeXParser parser,
            MathTypeEmbedder embedder) {
        LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(formula.latex());
        assertTrue(parsed.isSupported(), () -> formula.id() + ": " + parsed.diagnostics());

        XWPFTableRow row = table.createRow();
        row.setCantSplitRow(true);
        XWPFTableCell labelCell = row.getCell(0);
        XWPFTableCell formulaCell = row.getCell(1);
        configureCell(labelCell, LABEL_WIDTH_DXA, null);
        configureCell(formulaCell, FORMULA_WIDTH_DXA, null);

        XWPFParagraph labelParagraph = firstParagraph(labelCell);
        labelParagraph.setStyle("K12Body");
        labelParagraph.setSpacingBefore(0);
        labelParagraph.setSpacingAfter(0);
        labelParagraph.setSpacingBetween(1.0d, LineSpacingRule.AUTO);
        XWPFRun idRun = labelParagraph.createRun();
        idRun.setText(formula.id() + "  ");
        formatTextRun(idRun, 9.0d, "6B7280", true);
        XWPFRun nameRun = labelParagraph.createRun();
        nameRun.setText(formula.name());
        formatTextRun(nameRun, 10.5d, "202020", false);

        XWPFParagraph equationParagraph = firstParagraph(formulaCell);
        equationParagraph.setStyle("K12Body");
        equationParagraph.setSpacingBefore(0);
        equationParagraph.setSpacingAfter(0);
        equationParagraph.setSpacingBetween(1.0d, LineSpacingRule.AUTO);
        XWPFRun equationRun = equationParagraph.createRun();
        embedder.embedEquation(equationParagraph, equationRun, parsed.ast(), formula.latex(),
            DISPLAY_SCALE, MAX_FORMULA_WIDTH_PT);
    }

    private static XWPFParagraph firstParagraph(XWPFTableCell cell) {
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        while (!paragraph.getRuns().isEmpty()) {
            paragraph.removeRun(0);
        }
        return paragraph;
    }

    private static void configureCell(XWPFTableCell cell, int widthDxa, String fill) {
        cell.setWidth(Integer.toString(widthDxa));
        cell.setWidthType(TableWidthType.DXA);
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        if (fill != null) {
            cell.setColor(fill);
        }
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        XWPFParagraph paragraph = firstParagraph(cell);
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        formatTextRun(run, 10.0d, "1F3A5F", true);
    }

    private static void formatTextRun(XWPFRun run, double size, String color, boolean bold) {
        run.setFontFamily("Calibri", XWPFRun.FontCharRange.ascii);
        run.setFontFamily("Calibri", XWPFRun.FontCharRange.hAnsi);
        run.setFontFamily("Microsoft YaHei", XWPFRun.FontCharRange.eastAsia);
        run.setFontSize(size);
        run.setColor(color);
        run.setBold(bold);
    }

    private static void styleFieldRun(CTR run, double size, String color) {
        CTRPr properties = run.isSetRPr() ? run.getRPr() : run.addNewRPr();
        CTFonts fonts = properties.addNewRFonts();
        fonts.setAscii("Calibri");
        fonts.setHAnsi("Calibri");
        fonts.setEastAsia("Microsoft YaHei");
        properties.addNewColor().setVal(color);
        properties.addNewSz().setVal(BigInteger.valueOf(Math.round(size * 2.0d)));
        properties.addNewSzCs().setVal(BigInteger.valueOf(Math.round(size * 2.0d)));
    }

    private record Formula(String id, String group, String name, String latex) {
    }
}
