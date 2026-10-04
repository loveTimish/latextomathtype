package com.lz.paperword.controller;

import com.lz.paperword.core.docx.ImageAssetException;
import com.lz.paperword.service.*;
import org.junit.jupiter.api.Test;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;
import static org.mockito.Mockito.*;
import static org.mockito.ArgumentMatchers.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

class PdfExportControllerTest {
    @Test void exposesBothRoutesAndNativeComparison() throws Exception {
        var service=mock(LinuxPdfExportService.class);
        when(service.export(any(),any())).thenReturn("%PDF-test".getBytes());
        var mvc=MockMvcBuilders.standaloneSetup(new PdfExportController(mock(PaperExportService.class),mock(LayoutExportService.class),service)).build();
        mvc.perform(post("/api/export/pdf").contentType(MediaType.APPLICATION_JSON).content("{}"))
            .andExpect(status().isOk()).andExpect(content().contentType(MediaType.APPLICATION_PDF));
        verify(service).export(any(),eq(LinuxPdfExportService.Mode.compatible));
        mvc.perform(post("/api/export/layout-pdf?mode=native_layout").contentType(MediaType.APPLICATION_JSON).content("{}"))
            .andExpect(status().isOk());
        verify(service).export(any(),eq(LinuxPdfExportService.Mode.native_layout));
        mvc.perform(post("/api/export/pdf?mode=guess").contentType(MediaType.APPLICATION_JSON).content("{}"))
            .andExpect(status().isBadRequest());
    }
    @Test void errorsAreStructuredAndDoNotReturnPartialPdf() throws Exception {
        var service=mock(LinuxPdfExportService.class);
        var mvc=MockMvcBuilders.standaloneSetup(new PdfExportController(mock(PaperExportService.class),mock(LayoutExportService.class),service))
            .setControllerAdvice(new AssetErrorHandler()).build();
        String[] codes={"PDF_DISABLED","PDF_BUSY","PDF_TIMEOUT","PDF_INPUT_TOO_LARGE","PDF_CONVERSION_FAILED"};
        int[] statuses={503,429,504,413,422};
        for(int i=0;i<codes.length;i++) {
            doThrow(new LinuxPdfExportService.PdfExportException(codes[i],"test error")).when(service).export(any(),any());
            mvc.perform(post("/api/export/pdf").contentType(MediaType.APPLICATION_JSON).content("{}"))
                .andExpect(status().is(statuses[i])).andExpect(jsonPath("$.code").value(codes[i]));
        }
        doThrow(new ImageAssetException("IMAGE_ASSET_INVALID_CONTENT","Invalid PNG")).when(service).export(any(),any());
        mvc.perform(post("/api/export/pdf").contentType(MediaType.APPLICATION_JSON).content("{}"))
            .andExpect(status().isBadRequest()).andExpect(jsonPath("$.code").value("IMAGE_ASSET_INVALID_CONTENT"));
    }
}
