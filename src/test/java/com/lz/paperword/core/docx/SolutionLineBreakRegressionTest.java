package com.lz.paperword.core.docx;

import com.lz.paperword.model.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class SolutionLineBreakRegressionTest {
    @Test void explicitSolutionStepsRemainSeparateParagraphs() throws Exception {
        var question=new QuestionDTO();question.setSerialNumber(1);question.setQuestionType(5);
        question.setContent("求两个数的和。");
        question.setSolution("第一步：<b>先相加</b>。<br/>第二步：再检验。<BR />第三步：写答案。");
        var section=new SectionDTO();section.setQuestions(List.of(question));
        var request=new PaperExportRequest();request.setSections(List.of(section));
        try(var document=new XWPFDocument(new ByteArrayInputStream(new DocxBuilder(false).build(request)))) {
            List<String> paragraphs=document.getParagraphs().stream().map(p->p.getText()).toList();
            int first=paragraphs.indexOf("【解答】第一步：先相加。");
            assertTrue(first>=0,"The first step must not absorb the later steps: "+paragraphs);
            assertEquals("第二步：再检验。",paragraphs.get(first+1));
            assertEquals("第三步：写答案。",paragraphs.get(first+2));
        }
    }
    @Test void emptyHtmlLinesDoNotAddBlankSolutionParagraphs() throws Exception {
        var question=new QuestionDTO();question.setContent("题干");
        question.setSolution("步骤甲<br/><br/>  <br />步骤乙");
        var section=new SectionDTO();section.setQuestions(List.of(question));
        var request=new PaperExportRequest();request.setSections(List.of(section));
        try(var document=new XWPFDocument(new ByteArrayInputStream(new DocxBuilder(false).build(request)))) {
            List<String> paragraphs=document.getParagraphs().stream().map(p->p.getText()).toList();
            int first=paragraphs.indexOf("【解答】步骤甲");
            assertTrue(first>=0);assertEquals("步骤乙",paragraphs.get(first+1));
            assertEquals(3,paragraphs.size());
        }
    }
    @Test void htmlWrappingInsideMathDoesNotSplitOrLoseTheFormula()throws Exception {
        var method=DocxBuilder.class.getDeclaredMethod("normalizeHtmlLineBreaks",String.class);method.setAccessible(true);
        for(String[] pair:new String[][]{{"$","$"},{"$$","$$"},{"\\(","\\)"},{"\\[","\\]"}}) {
            String input="第一步 "+pair[0]+"a<br/>+b"+pair[1]+"<br/>第二步";
            assertEquals("第一步 "+pair[0]+"a +b"+pair[1]+"<br/>第二步",method.invoke(new DocxBuilder(false),input));
        }
    }
    @Test void escapedDollarDoesNotConsumeFollowingStepBreaks()throws Exception {
        var method=DocxBuilder.class.getDeclaredMethod("normalizeHtmlLineBreaks",String.class);method.setAccessible(true);
        assertEquals("价格\\$5<br/>下一步",method.invoke(new DocxBuilder(false),"价格\\$5<br/>下一步"));
    }
}
