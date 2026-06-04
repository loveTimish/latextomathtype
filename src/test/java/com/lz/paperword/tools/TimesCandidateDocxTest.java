package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

class TimesCandidateDocxTest {
    @Test
    void generateTimesCandidateDocx() throws Exception {
        String output = System.getProperty("latextomathtype.times.candidate.output",
            "target/reference-roundtrip/times-candidate.docx");
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("times candidate");
        paper.setSubjectType(2);
        paper.setStage(2);
        paper.setScore(1);
        paper.setSuggestTime(1);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("candidate");

        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(5);
        question.setScore(1);
        question.setContent("<p>$\\frac{1}{1\\times2}+\\frac{1}{2\\times3}+\\frac{99}{1\\times2\\times3\\times\\cdots\\times100}$</p>");
        section.setQuestions(List.of(question));
        request.setSections(List.of(section));

        byte[] docx = new DocxBuilder(true).build(request);
        Path out = Path.of(output);
        Files.createDirectories(out.getParent());
        Files.write(out, docx);
        assertTrue(Files.size(out) > 1000);
    }
}
