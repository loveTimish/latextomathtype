package com.lz.paperword.core.docx;

import com.lz.paperword.model.layout.LayoutDocumentRequest;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;

import java.io.ByteArrayInputStream;
import java.math.BigInteger;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LayoutDocxBuilderTest {

    private final LayoutDocxBuilder builder = new LayoutDocxBuilder(false);

    @Test
    void writesMultipleLineSpacingAsAutoTwips() throws Exception {
        LayoutDocumentRequest.Style style = new LayoutDocumentRequest.Style();
        style.setLineSpacingMultiple(1.5d);

        try (XWPFDocument document = open(buildRequest(paragraphBlock("正文", style)))) {
            CTSpacing spacing = document.getParagraphArray(0).getCTP().getPPr().getSpacing();

            assertEquals(BigInteger.valueOf(360), spacing.getLine());
            assertEquals(STLineSpacingRule.AUTO, spacing.getLineRule());
        }
    }

    @Test
    void tableCellsRemoveBlockSpacingAndKeepConfiguredLineSpacing() throws Exception {
        LayoutDocumentRequest.Style style = new LayoutDocumentRequest.Style();
        style.setSpacingBeforeTwips(240);
        style.setSpacingAfterTwips(180);
        style.setLineSpacingMultiple(1.25d);

        LayoutDocumentRequest.Block tableBlock = new LayoutDocumentRequest.Block();
        tableBlock.setType(LayoutDocumentRequest.BlockType.TABLE);
        tableBlock.setStyle(style);
        LayoutDocumentRequest.TableCell tableCell = new LayoutDocumentRequest.TableCell();
        tableCell.setText("单元格");
        LayoutDocumentRequest.TableRow tableRow = new LayoutDocumentRequest.TableRow();
        tableRow.setCells(List.of(tableCell));
        tableBlock.setRows(List.of(tableRow));

        try (XWPFDocument document = open(buildRequest(tableBlock))) {
            XWPFTableCell cell = document.getTables().getFirst().getRow(0).getCell(0);
            CTSpacing spacing = cell.getParagraphs().getFirst().getCTP().getPPr().getSpacing();

            assertEquals(BigInteger.ZERO, spacing.getBefore());
            assertEquals(BigInteger.ZERO, spacing.getAfter());
            assertEquals(BigInteger.valueOf(300), spacing.getLine());
            assertEquals(STLineSpacingRule.AUTO, spacing.getLineRule());
        }
    }

    @Test
    void ignoresNonFiniteLineSpacing() throws Exception {
        LayoutDocumentRequest.Style style = new LayoutDocumentRequest.Style();
        style.setLineSpacingMultiple(Double.NaN);

        try (XWPFDocument document = open(buildRequest(paragraphBlock("正文", style)))) {
            CTSpacing spacing = document.getParagraphArray(0).getCTP().getPPr().getSpacing();

            assertFalse(spacing.isSetLine());
            assertFalse(spacing.isSetLineRule());
        }
    }

    @Test
    void abortsExportWithOriginalLatexWhenFormulaIsUnsupported() {
        LayoutDocumentRequest.Block block = paragraphBlock("$$\\unknownpackagecommand{x}$$",
            new LayoutDocumentRequest.Style());

        IllegalArgumentException error = assertThrows(IllegalArgumentException.class,
            () -> buildRequest(block));

        assertTrue(error.getMessage().contains("\\unknownpackagecommand{x}"));
        assertTrue(error.getMessage().contains("UNSUPPORTED_COMMAND"));
    }

    private byte[] buildRequest(LayoutDocumentRequest.Block block) throws Exception {
        LayoutDocumentRequest request = new LayoutDocumentRequest();
        LayoutDocumentRequest.Page page = new LayoutDocumentRequest.Page();
        page.setPageNumber(1);
        page.setBlocks(List.of(block));
        request.setPages(List.of(page));
        return builder.build(request);
    }

    private LayoutDocumentRequest.Block paragraphBlock(String text, LayoutDocumentRequest.Style style) {
        LayoutDocumentRequest.Block block = new LayoutDocumentRequest.Block();
        block.setType(LayoutDocumentRequest.BlockType.PARAGRAPH);
        block.setText(text);
        block.setStyle(style);
        return block;
    }

    private XWPFDocument open(byte[] docx) throws Exception {
        return new XWPFDocument(new ByteArrayInputStream(docx));
    }
}
