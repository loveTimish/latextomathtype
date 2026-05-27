package com.lz.paperword.model;

import lombok.Data;
import java.util.List;

@Data
public class PaperExportRequest {

    private PaperInfo paper;
    private List<SectionDTO> sections;

    @Data
    public static class PaperInfo {
        private String name;
        private Integer subjectType;
        private Integer stage;
        private Integer score;
        private Integer suggestTime;
    }
}
