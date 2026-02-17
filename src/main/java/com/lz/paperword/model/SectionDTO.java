package com.lz.paperword.model;

import lombok.Data;
import java.util.List;

@Data
public class SectionDTO {

    private String headline;
    private List<QuestionDTO> questions;
}
