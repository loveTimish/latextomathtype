package com.lz.paperword.service;

import com.lz.paperword.core.docx.LayoutDocxBuilder;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * 负责将版面布局模型导出为保留分页与块结构的 Word 文档。
 */
@Service
public class LayoutExportService {

    private static final Logger log = LoggerFactory.getLogger(LayoutExportService.class);

    private final LayoutDocxBuilder oleDocxBuilder = new LayoutDocxBuilder(true);

    public byte[] exportLayoutDocument(LayoutDocumentRequest request) throws IOException {
        long start = System.currentTimeMillis();
        byte[] docx = sanitizeNestedOleRuns(oleDocxBuilder.build(request));
        log.info("Layout document exported in {}ms, size: {} bytes",
            System.currentTimeMillis() - start, docx.length);
        return docx;
    }

    /**
     * 兜底修正嵌套 OLE run 的坏结构，避免 Word 打开后对象关系错位。
     */
    private byte[] sanitizeNestedOleRuns(byte[] docx) throws IOException {
        try (ByteArrayInputStream input = new ByteArrayInputStream(docx);
             ZipInputStream zipIn = new ZipInputStream(input);
             ByteArrayOutputStream output = new ByteArrayOutputStream();
             ZipOutputStream zipOut = new ZipOutputStream(output)) {
            ZipEntry entry;
            while ((entry = zipIn.getNextEntry()) != null) {
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zipOut.putNextEntry(newEntry);
                byte[] content = zipIn.readAllBytes();
                if ("word/document.xml".equals(entry.getName())) {
                    String xml = new String(content, StandardCharsets.UTF_8);
                    xml = xml.replace("<w:r><w:r ", "<w:r ");
                    xml = xml.replace("</w:object></w:r></w:r>", "</w:object></w:r>");
                    content = xml.getBytes(StandardCharsets.UTF_8);
                }
                zipOut.write(content);
                zipOut.closeEntry();
                zipIn.closeEntry();
            }
            zipOut.finish();
            return output.toByteArray();
        }
    }
}
