package com.lz.paperword.controller;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import com.lz.paperword.service.LayoutExportService;
import com.lz.paperword.service.PaperExportService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

/**
 * REST controller for exporting exam papers to Word documents.
 */
@RestController
@RequestMapping("/api/export")
public class ExportController {

    private static final Logger log = LoggerFactory.getLogger(ExportController.class);

    private static final MediaType DOCX_MEDIA_TYPE = MediaType.parseMediaType(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

    private final PaperExportService exportService;
    private final LayoutExportService layoutExportService;

    public ExportController(PaperExportService exportService, LayoutExportService layoutExportService) {
        this.exportService = exportService;
        this.layoutExportService = layoutExportService;
    }

    /**
     * Export a paper to Word document with MathType formulas.
     *
     * POST /api/export/word
     * Content-Type: application/json
     * Body: PaperExportRequest JSON
     * Response: .docx binary file
     */
    @PostMapping("/word")
    public ResponseEntity<byte[]> exportWord(@RequestBody PaperExportRequest request) throws IOException {
        log.info("Received export request: {}",
            request.getPaper() != null ? request.getPaper().getName() : "unnamed");

        byte[] docxBytes = exportService.export(request);

        String fileName = "试卷.docx";
        if (request.getPaper() != null && request.getPaper().getName() != null) {
            fileName = request.getPaper().getName() + ".docx";
        }
        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8)
            .replace("+", "%20");

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(DOCX_MEDIA_TYPE);
        headers.set(HttpHeaders.CONTENT_DISPOSITION,
            "attachment; filename*=UTF-8''" + encodedFileName);
        headers.setContentLength(docxBytes.length);

        return ResponseEntity.ok()
            .headers(headers)
            .body(docxBytes);
    }

    /**
     * 直接按分页布局导出 Word，用于 PDF OCR 链路的块级版面保真输出。
     */
    @PostMapping("/layout-word")
    public ResponseEntity<byte[]> exportLayoutWord(@RequestBody LayoutDocumentRequest request) throws IOException {
        log.info("Received layout export request: {}",
            request.getDocument() != null ? request.getDocument().getTitle() : "unnamed");

        byte[] docxBytes = layoutExportService.exportLayoutDocument(request);
        String fileName = "layout-export.docx";
        if (request.getDocument() != null && request.getDocument().getTitle() != null
            && !request.getDocument().getTitle().isBlank()) {
            fileName = request.getDocument().getTitle() + ".docx";
        }
        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8)
            .replace("+", "%20");

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(DOCX_MEDIA_TYPE);
        headers.set(HttpHeaders.CONTENT_DISPOSITION,
            "attachment; filename*=UTF-8''" + encodedFileName);
        headers.setContentLength(docxBytes.length);

        return ResponseEntity.ok()
            .headers(headers)
            .body(docxBytes);
    }

    /**
     * Health check endpoint.
     */
    @GetMapping("/health")
    public ResponseEntity<String> health() {
        return ResponseEntity.ok("paper-to-word service is running");
    }
}
