package com.lz.paperword.controller;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import com.lz.paperword.service.*;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import java.io.IOException;
import java.util.Map;

@RestController
@RequestMapping("/api/export")
public final class PdfExportController {
    private final PaperExportService papers;
    private final LayoutExportService layouts;
    private final LinuxPdfExportService pdfs;
    public PdfExportController(PaperExportService papers, LayoutExportService layouts, LinuxPdfExportService pdfs) {
        this.papers=papers;this.layouts=layouts;this.pdfs=pdfs;
    }
    @PostMapping("/pdf")
    public ResponseEntity<byte[]> paper(@RequestBody PaperExportRequest request,
        @RequestParam(name="mode", defaultValue="compatible") LinuxPdfExportService.Mode mode) throws IOException {
        return response(pdfs.export(()->papers.export(request),mode));
    }
    @PostMapping("/layout-pdf")
    public ResponseEntity<byte[]> layout(@RequestBody LayoutDocumentRequest request,
        @RequestParam(name="mode", defaultValue="compatible") LinuxPdfExportService.Mode mode) throws IOException {
        return response(pdfs.export(()->layouts.exportLayoutDocument(request),mode));
    }
    @GetMapping("/pdf/status") public Map<String,Object> status() { return pdfs.status(); }
    private ResponseEntity<byte[]> response(byte[] bytes) {
        return ResponseEntity.ok().contentType(MediaType.APPLICATION_PDF).contentLength(bytes.length)
            .header(HttpHeaders.CONTENT_DISPOSITION,"attachment; filename=export.pdf").body(bytes);
    }
    @ExceptionHandler(LinuxPdfExportService.PdfExportException.class)
    public ResponseEntity<Map<String,String>> failure(LinuxPdfExportService.PdfExportException e) {
        HttpStatus status=switch(e.getCode()) {
            case "PDF_BUSY" -> HttpStatus.TOO_MANY_REQUESTS;
            case "PDF_DISABLED", "PDF_UNAVAILABLE", "PDF_CONFIGURATION" -> HttpStatus.SERVICE_UNAVAILABLE;
            case "PDF_TIMEOUT" -> HttpStatus.GATEWAY_TIMEOUT;
            case "PDF_INPUT_TOO_LARGE", "PDF_OUTPUT_TOO_LARGE" -> HttpStatus.PAYLOAD_TOO_LARGE;
            case "PDF_CONVERSION_FAILED" -> HttpStatus.UNPROCESSABLE_ENTITY;
            default -> HttpStatus.INTERNAL_SERVER_ERROR;
        };
        return ResponseEntity.status(status).body(Map.of("code",e.getCode(),"message",e.getMessage()));
    }
}
