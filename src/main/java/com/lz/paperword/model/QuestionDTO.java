package com.lz.paperword.model;

import lombok.Data;
import java.util.List;

@Data
public class QuestionDTO {

    private Integer serialNumber;
    /** 可选：教学阶段小节标题（如“课堂讲解/课堂练习/课后作业”），设置后会在该题前渲染一个小节标题 */
    private String phaseLabel;
    /**
     * 1.单选题 2.多选题 3.判断题 4.填空题 5.解答题 6.计算题
     */
    private Integer questionType;
    private String content;
    /** 题目配图的本地文件路径列表（按出现顺序渲染在题干之后、选项之前） */
    private List<String> images;
    private List<OptionDTO> options;
    private String correct;
    private Integer score;
    private String analyze;
    private String solution;
    private String difficulty;
    private String knowledgePoint;
    private List<String> tags;

    @Data
    public static class OptionDTO {
        private String prefix;
        private String content;
    }
}
