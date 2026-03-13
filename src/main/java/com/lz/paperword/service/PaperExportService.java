package com.lz.paperword.service;

import com.lz.paperword.config.WindowsMathTypeProperties;
import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.IOException;

/**
 * Service for exporting exam papers to Word documents with MathType formulas.
 * This service can be used both via REST API and as a direct library call.
 */
@Service
public class PaperExportService {

    private static final Logger log = LoggerFactory.getLogger(PaperExportService.class);

    private final DocxBuilder oleDocxBuilder = new DocxBuilder(true);
    private final DocxBuilder draftDocxBuilder = new DocxBuilder(false);
    private final WindowsMathTypeProperties windowsMathTypeProperties;
    private final WindowsMathTypeClient windowsMathTypeClient;

    public PaperExportService(WindowsMathTypeProperties windowsMathTypeProperties,
                              WindowsMathTypeClient windowsMathTypeClient) {
        this.windowsMathTypeProperties = windowsMathTypeProperties;
        this.windowsMathTypeClient = windowsMathTypeClient;
    }

    /**
     * Export a paper to a .docx byte array.
     *
     * @param request the paper export data (title, sections, questions with LaTeX)
     * @return the .docx file content as byte array
     * @throws IOException if document generation fails
     */
    public byte[] export(PaperExportRequest request) throws IOException {
        log.info("Exporting paper to Word: {}", 
            request.getPaper() != null ? request.getPaper().getName() : "unnamed");

        long start = System.currentTimeMillis();
        byte[] docx;

        if (!windowsMathTypeProperties.isEnabled()) {
            docx = oleDocxBuilder.build(request);
        } else {
            byte[] draft = draftDocxBuilder.build(request);
            docx = windowsMathTypeClient.convert(draft);
        }

        long elapsed = System.currentTimeMillis() - start;

        log.info("Paper exported successfully in {}ms, size: {} bytes", elapsed, docx.length);
        return docx;
    }
}
