package com.lz.paperword.model;

import lombok.Data;
import java.util.List;

@Data
public class SectionDTO {

    private String headline;
    /** Optional section-level images rendered below the section headline. */
    private List<String> images;
    /** Optional maximum display width in pixels for section-level images. */
    private Integer imageMaxWidthPx;
    /** Optional alignment for section-level images: left, center, or right. */
    private String imageAlignment;
    private List<QuestionDTO> questions;
}
