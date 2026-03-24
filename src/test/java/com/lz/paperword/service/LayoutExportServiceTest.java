package com.lz.paperword.service;

import com.lz.paperword.model.layout.LayoutDocumentRequest;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LayoutExportServiceTest {

    private final LayoutExportService service = new LayoutExportService();

    @Test
    void shouldExportVerticalAndCrossLayoutBlocksAsOleObjects() throws IOException {
        LayoutDocumentRequest request = new LayoutDocumentRequest();
        request.getDocument().setTitle("版面导出回归");

        LayoutDocumentRequest.Page page = new LayoutDocumentRequest.Page();
        page.setPageNumber(1);
        page.getBlocks().add(paragraphBlock("b1", "1. 用竖式计算 23×12"));
        page.getBlocks().add(paragraphBlock("b2",
            "【解答】\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {} \\\\ \\hline {} & {5} & {5} & {3} & {5} & {}\\end{array}"));
        page.getBlocks().add(paragraphBlock("b3", "2. 按十字交叉法配制 30% 溶液"));
        page.getBlocks().add(paragraphBlock("b4",
            "【解答】\\begin{array}{ccccc} 50\\% & & & & 10\\% \\\\ & {}\\searrow{} & & {}\\nearrow{} & \\\\ & & 30\\% & & \\\\ & {}\\nearrow{} & & {}\\searrow{} & \\\\ 20\\% & & & & 20\\% \\end{array}"));
        request.setPages(java.util.List.of(page));

        byte[] docx = service.exportLayoutDocument(request);
        assertNotNull(docx);
        assertTrue(docx.length > 100);

        String documentXml = unzipDocumentXml(docx);
        assertTrue(documentXml.contains("用竖式计算 23×12"));
        assertTrue(documentXml.contains("按十字交叉法配制 30% 溶液"));

        Matcher matcher = Pattern.compile("<o:OLEObject\\b").matcher(documentXml);
        int count = 0;
        while (matcher.find()) {
            count++;
        }
        assertTrue(count >= 2, "版面导出应把竖式和十字交叉都保留为 OLE 对象");
    }

    private LayoutDocumentRequest.Block paragraphBlock(String id, String text) {
        LayoutDocumentRequest.Block block = new LayoutDocumentRequest.Block();
        block.setId(id);
        block.setType(LayoutDocumentRequest.BlockType.PARAGRAPH);
        block.setText(text);
        block.setStyle(new LayoutDocumentRequest.Style());
        return block;
    }

    private String unzipDocumentXml(byte[] docx) throws IOException {
        try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(docx), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zipInputStream.getNextEntry()) != null) {
                if ("word/document.xml".equals(entry.getName())) {
                    return new String(zipInputStream.readAllBytes(), StandardCharsets.UTF_8);
                }
            }
        }
        throw new IOException("word/document.xml not found");
    }
}
