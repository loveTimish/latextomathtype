package com.lz.paperword.core.docx;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;

import java.util.List;

/**
 * 测试辅助类，用于构建简单的测试请求。
 */
class TestRequestBuilder {
    static PaperExportRequest build(String content) {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("Test Paper");
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("测试");

        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(6);
        question.setScore(5);
        question.setContent(content);

        section.setQuestions(List.of(question));
        request.setSections(List.of(section));
        return request;
    }
}
