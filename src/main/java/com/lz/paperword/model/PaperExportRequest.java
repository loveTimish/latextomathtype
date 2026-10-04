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
        /** Optional source document name used by compact reconstructed footers. */
        private String sourceName;
        /** Optional compact handout layout for reference-style regenerated documents. */
        private Boolean compactLayout;
        /** Optional: suppress generated question type labels in compact reconstructed handouts. */
        private Boolean hideQuestionTypeMetadata;
        /** Optional worksheet layout: keep question headers/workspace and solution paragraphs together. */
        private Boolean printLayout;
        /** Optional neutral PAGE/NUMPAGES footer in the non-compact layout. */
        private Boolean pageNumbers;
        /** Optional typography profile: "legacy" (default) or "exam"; exam is not compact. */
        private String typography;
        /** Optional exam preview background: "transparent" (default) or "white". */
        private String previewBackground;
    }
}
