package com.lz.paperword.core.docx;

import com.lz.paperword.model.*;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class PrintLayoutRegressionTest {
    private PaperExportRequest request(boolean printable, boolean solved) {
        var paper=new PaperExportRequest.PaperInfo();paper.setPrintLayout(printable);paper.setPageNumbers(printable);
        var question=new QuestionDTO();question.setSerialNumber(1);question.setQuestionType(5);question.setContent("题干甲<br/>题干乙");question.setAnswerSpaceLines(5);
        if(solved){question.setSolution("步骤甲<br/>步骤乙<br/>步骤丙");question.setCorrect("答案");}
        var section=new SectionDTO();section.setQuestions(List.of(question));
        var request=new PaperExportRequest();request.setPaper(paper);request.setSections(List.of(section));return request;
    }
    private XWPFDocument build(PaperExportRequest request)throws Exception {
        return new XWPFDocument(new ByteArrayInputStream(new DocxBuilder(false).build(request)));
    }
    @Test void unansweredQuestionKeepsOnlyItsInitialWorkingSpace()throws Exception {
        try(var doc=build(request(true,false))) {
            var p=doc.getParagraphs();assertEquals(7,p.size());
            for(int i=0;i<p.size();i++) {
                assertTrue(p.get(i).getCTP().getPPr().isSetKeepLines());
                assertEquals(i<3,p.get(i).getCTP().getPPr().isSetKeepNext());
            }
            String footer=doc.getFooterList().getFirst()._getHdrFtr().xmlText();
            assertTrue(footer.contains("PAGE"));assertTrue(footer.contains("NUMPAGES"));
            assertFalse(footer.contains("教师版"));
        }
    }
    @Test void teacherStepsCanBreakOnlyAtParagraphBoundaries()throws Exception {
        try(var doc=build(request(true,true))) {
            var p=doc.getParagraphs();assertEquals(6,p.size());
            assertTrue(p.get(0).getCTP().getPPr().isSetKeepNext());
            assertTrue(p.get(1).getCTP().getPPr().isSetKeepNext());
            assertFalse(p.get(2).getCTP().getPPr().isSetKeepNext());
            assertTrue(p.get(3).getCTP().getPPr().isSetKeepNext());
            assertTrue(p.get(4).getCTP().getPPr().isSetKeepNext());
            assertTrue(p.stream().allMatch(v->v.getCTP().getPPr().isSetKeepLines()));
        }
    }
    @Test void defaultsDoNotForceNewPaginationOrFooter()throws Exception {
        try(var doc=build(request(false,false))) {
            assertTrue(doc.getFooterList().isEmpty());
            assertTrue(doc.getParagraphs().stream().allMatch(p->!p.getCTP().getPPr().isSetKeepNext()&&!p.getCTP().getPPr().isSetKeepLines()));
        }
    }
    @Test void singleSolutionStepKeepsItsAnswerOnTheSamePage()throws Exception {
        var request=request(true,true);request.getSections().getFirst().getQuestions().getFirst().setSolution("仅这一步");
        try(var doc=build(request)) {
            var p=doc.getParagraphs();assertEquals(4,p.size());
            assertTrue(p.get(2).getCTP().getPPr().isSetKeepNext());
            assertFalse(p.get(3).getCTP().getPPr().isSetKeepNext());
        }
    }
    @Test void workingSpaceIsBounded() {
        for(int lines:new int[]{-1,13,Integer.MAX_VALUE}) {
            var request=request(true,false);request.getSections().getFirst().getQuestions().getFirst().setAnswerSpaceLines(lines);
            assertThrows(IllegalArgumentException.class,()->new DocxBuilder(false).build(request));
        }
    }
    @Test void reusedBuilderResetsPrintFlagsAndTwoQuestionsStayIndependent()throws Exception {
        var builder=new DocxBuilder(false);var first=request(true,false);
        var second=new QuestionDTO();second.setSerialNumber(2);second.setQuestionType(4);second.setContent("第二题");
        first.getSections().getFirst().setQuestions(List.of(first.getSections().getFirst().getQuestions().getFirst(),second));
        try(var doc=new XWPFDocument(new ByteArrayInputStream(builder.build(first)))) {
            assertFalse(doc.getParagraphs().get(6).getCTP().getPPr().isSetKeepNext());
            assertEquals("2. 第二题",doc.getParagraphs().get(7).getText());
        }
        try(var doc=new XWPFDocument(new ByteArrayInputStream(builder.build(request(false,false))))) {
            assertTrue(doc.getFooterList().isEmpty());
            assertTrue(doc.getParagraphs().stream().noneMatch(p->p.getCTP().getPPr().isSetKeepLines()));
        }
    }
    @Test void pageNumbersAndPrintGroupingAreIndependentOptions()throws Exception {
        var onlyNumbers=request(false,false);onlyNumbers.getPaper().setPageNumbers(true);
        try(var doc=build(onlyNumbers)) {
            assertFalse(doc.getFooterList().isEmpty());
            assertTrue(doc.getParagraphs().stream().noneMatch(p->p.getCTP().getPPr().isSetKeepLines()));
        }
        var onlyGrouping=request(true,false);onlyGrouping.getPaper().setPageNumbers(false);
        try(var doc=build(onlyGrouping)) {
            assertTrue(doc.getFooterList().isEmpty());
            assertTrue(doc.getParagraphs().stream().allMatch(p->p.getCTP().getPPr().isSetKeepLines()));
        }
    }
}
