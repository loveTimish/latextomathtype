package com.lz.paperword.model;

import lombok.Data;
import java.util.List;

@Data
public class QuestionDTO {

    private Integer serialNumber;
    /**
     * 1.单选题 2.多选题 3.判断题 4.填空题 5.解答题 6.计算题
     */
    private Integer questionType;
    private String content;
    private List<OptionDTO> options;
    private String correct;
    private Integer score;
    private String analyze;

    @Data
    public static class OptionDTO {
        private String prefix;
        private String content;
    }
}
