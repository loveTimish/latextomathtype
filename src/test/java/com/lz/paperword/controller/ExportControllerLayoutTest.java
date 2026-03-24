package com.lz.paperword.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import com.lz.paperword.service.LayoutExportService;
import com.lz.paperword.service.PaperExportService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.content;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.header;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@WebMvcTest(ExportController.class)
class ExportControllerLayoutTest {

    @Autowired
    private MockMvc mockMvc;

    @Autowired
    private ObjectMapper objectMapper;

    @MockBean
    private PaperExportService paperExportService;

    @MockBean
    private LayoutExportService layoutExportService;

    @Test
    void shouldExposeLayoutWordExportEndpoint() throws Exception {
        when(layoutExportService.exportLayoutDocument(any(LayoutDocumentRequest.class)))
            .thenReturn(new byte[]{0x50, 0x4B, 0x03, 0x04});

        LayoutDocumentRequest request = new LayoutDocumentRequest();
        request.getDocument().setTitle("layout-case");

        mockMvc.perform(post("/api/export/layout-word")
                .contentType(MediaType.APPLICATION_JSON)
                .content(objectMapper.writeValueAsBytes(request)))
            .andExpect(status().isOk())
            .andExpect(header().string("Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
            .andExpect(content().bytes(new byte[]{0x50, 0x4B, 0x03, 0x04}));
    }

    @Test
    void shouldKeepExistingPaperWordEndpointUntouched() throws Exception {
        when(paperExportService.export(any(PaperExportRequest.class)))
            .thenReturn(new byte[]{0x50, 0x4B, 0x03, 0x04});

        PaperExportRequest request = new PaperExportRequest();
        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("paper-case");
        request.setPaper(paper);

        mockMvc.perform(post("/api/export/word")
                .contentType(MediaType.APPLICATION_JSON)
                .content(objectMapper.writeValueAsBytes(request)))
            .andExpect(status().isOk())
            .andExpect(content().bytes(new byte[]{0x50, 0x4B, 0x03, 0x04}));
    }
}
