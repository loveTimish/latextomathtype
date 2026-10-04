package com.lz.paperword.core.docx;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.model.*;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;

import java.io.ByteArrayInputStream;
import java.math.BigInteger;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ExamTypographyRegressionTest {
    private PaperExportRequest request(String typography, QuestionDTO... questions) {
        var paper = new PaperExportRequest.PaperInfo(); paper.setTypography(typography);
        var section = new SectionDTO(); section.setQuestions(List.of(questions));
        var request = new PaperExportRequest(); request.setPaper(paper); request.setSections(List.of(section));
        return request;
    }
    private QuestionDTO question(String content) {
        var question = new QuestionDTO(); question.setSerialNumber(1); question.setContent(content); return question;
    }
    private XWPFDocument build(PaperExportRequest request) throws Exception {
        return new XWPFDocument(new ByteArrayInputStream(new DocxBuilder(false).build(request)));
    }

    @Test void usesBlackTwelvePointBodyAndHangingTwentyFourPointNumbers() throws Exception {
        var question = question("正文 ABC<br/>续行内容<br/>$$\\frac{1}{2}$$。");
        question.setSerialNumber(12);
        question.setKnowledgePoint("合成知识点");
        try (var doc = build(request("exam", question))) {
            var paragraphs = doc.getParagraphs();
            assertEquals(480, paragraphs.getFirst().getIndentationLeft());
            assertEquals(480, paragraphs.getFirst().getIndentationHanging());
            assertFalse(paragraphs.getFirst().getCTP().getPPr().getInd().isSetFirstLine());
            assertTrue(paragraphs.getFirst().getText().startsWith("12.\t"));
            for (int i = 1; i < paragraphs.size(); i++) {
                assertEquals(480, paragraphs.get(i).getIndentationLeft());
                assertFalse(paragraphs.get(i).getCTP().getPPr().getInd().isSetHanging());
                assertEquals(0, paragraphs.get(i).getIndentationFirstLine());
            }
            for (XWPFRun run : paragraphs.getFirst().getRuns()) {
                assertEquals(12.0, run.getFontSizeAsDouble());
                assertEquals("000000", run.getColor());
                assertEquals("Times New Roman", run.getFontFamily(XWPFRun.FontCharRange.ascii));
                assertEquals("宋体", run.getFontFamily(XWPFRun.FontCharRange.eastAsia));
            }
            var bodySpacing = paragraphs.getFirst().getCTP().getPPr().getSpacing();
            assertEquals(BigInteger.valueOf(380), bodySpacing.getLine());
            assertEquals(STLineSpacingRule.EXACT, bodySpacing.getLineRule());
            assertEquals(STLineSpacingRule.AT_LEAST, paragraphs.get(2).getCTP().getPPr().getSpacing().getLineRule());
            assertNotEquals(ParagraphAlignment.CENTER, paragraphs.get(2).getAlignment());
            var metadata = paragraphs.getLast();
            assertEquals(9.5, metadata.getRuns().getFirst().getFontSizeAsDouble());
            assertEquals(BigInteger.valueOf(280), metadata.getCTP().getPPr().getSpacing().getLine());
            var margin = doc.getDocument().getBody().getSectPr().getPgMar();
            assertEquals(BigInteger.valueOf(1247), margin.getLeft());
            assertEquals(BigInteger.valueOf(1247), margin.getRight());
            assertEquals(BigInteger.valueOf(1020), margin.getTop());
            assertEquals(BigInteger.valueOf(1020), margin.getBottom());
            assertEquals(BigInteger.valueOf(425), margin.getFooter());
        }
    }

    @Test void formulaMetadataDoesNotGetAnExactHeight() throws Exception {
        var question = question("正文"); question.setKnowledgePoint("比值 $\\frac{1}{2}$");
        try (var doc = build(request("exam", question))) {
            var metadata = doc.getParagraphs().getLast();
            assertEquals(STLineSpacingRule.AT_LEAST, metadata.getCTP().getPPr().getSpacing().getLineRule());
            assertEquals(9.5, metadata.getRuns().getLast().getFontSizeAsDouble());
        }
    }

    @Test void workingLinesHaveExactTwentyPointHeightAndBoundedKeeps() throws Exception {
        var question = question("简短题干"); question.setQuestionType(5); question.setAnswerSpaceLines(12);
        try (var doc = build(request("exam", question))) {
            var paragraphs = doc.getParagraphs();
            assertEquals(13, paragraphs.size());
            for (int i = 1; i < paragraphs.size(); i++) {
                var spacing = paragraphs.get(i).getCTP().getPPr().getSpacing();
                assertEquals(BigInteger.valueOf(400), spacing.getLine());
                assertEquals(STLineSpacingRule.EXACT, spacing.getLineRule());
                assertEquals(BigInteger.ZERO, spacing.getAfter());
                if (i >= 2) assertFalse(paragraphs.get(i).getCTP().getPPr().isSetKeepNext());
            }
            assertTrue(paragraphs.getFirst().getCTP().getPPr().isSetKeepNext());
            assertTrue(paragraphs.get(1).getCTP().getPPr().isSetKeepNext());
        }
    }

    @Test void pageBreakIsAppliedOnceAtQuestionStartEvenWithNullStemOrPhase(@TempDir Path dir) throws Exception {
        Path image = dir.resolve("synthetic.png"); ImageAssetLoaderTest.image(image, "png", 10, 10);
        var question = question(null); question.setImages(List.of(image.toString())); question.setPageBreakBefore(true);
        var second = question("第二题"); second.setSerialNumber(2);
        for (String phase : new String[]{null, "练习"}) {
            question.setPhaseLabel(phase);
            var assets = new ImageAssetLoader(new ImageAssetConfig(dir, null,
                ImageAssetConfig.DEFAULT_MAX_BYTES, ImageAssetConfig.DEFAULT_MAX_PIXELS));
            try (var doc = new XWPFDocument(new ByteArrayInputStream(new DocxBuilder(false, assets).build(request("exam", question, second))))) {
                var paragraphs = doc.getParagraphs();
                assertTrue(paragraphs.getFirst().isPageBreak());
                assertEquals(1, paragraphs.stream().filter(XWPFParagraph::isPageBreak).count());
                assertFalse(paragraphs.getLast().isPageBreak());
                assertEquals(1, doc.getAllPictures().size());
                var imageParagraph = paragraphs.stream().filter(p -> p.getRuns().stream().anyMatch(r -> !r.getEmbeddedPictures().isEmpty())).findFirst().orElseThrow();
                assertEquals(STLineSpacingRule.AT_LEAST, imageParagraph.getCTP().getPPr().getSpacing().getLineRule());
            }
        }
    }

    @Test void flowAppliesOnlyToStemAndLeavesSolutionAnalysisSteps() throws Exception {
        var question = question("第一行\n继续<br/>最后"); question.setContentFlow("flow");
        question.setAnalyze("分析甲<br/>分析乙"); question.setSolution("解答甲<br/>解答乙"); question.setCorrect("答案文本");
        try (var doc = build(request(null, question))) {
            assertEquals(List.of("1. 第一行继续最后", "【解析】分析甲", "分析乙", "【解答】解答甲", "解答乙", "【答案】答案文本"),
                doc.getParagraphs().stream().map(XWPFParagraph::getText).toList());
        }
    }

    @Test void rejectsUnknownModesConflictsAndUnsupportedFlowAssets() {
        var question = question("题干");
        for (String mode : List.of("other", "", "EXAM")) {
            assertThrows(ExportRequestValidationException.class, () -> build(request(mode, question)));
            question.setContentFlow(mode);
            assertThrows(ExportRequestValidationException.class, () -> build(request(null, question)));
            question.setContentFlow(null);
        }
        var conflict = request("exam", question); conflict.getPaper().setCompactLayout(true);
        assertThrows(ExportRequestValidationException.class, () -> build(conflict));
        question.setContent("$\\pwmetrics{20,10}x$");
        assertThrows(ExportRequestValidationException.class, () -> build(request("exam", question)));
        question.setContentFlow("flow");
        for (String content : List.of("前<img src='x.png'>后", "前<table><tr><td>x</td></tr></table>后")) {
            question.setContent(content);
            assertThrows(ExportRequestValidationException.class, () -> build(request(null, question)));
        }
    }

    @Test void reusedBuilderRestoresLegacyAndFieldNamesRoundTrip() throws Exception {
        var builder = new DocxBuilder(false); var question = question("正文");
        builder.build(request("exam", question));
        try (var doc = new XWPFDocument(new ByteArrayInputStream(builder.build(request(null, question))))) {
            var paragraph = doc.getParagraphs().getFirst();
            assertEquals("1. 正文", paragraph.getText());
            assertEquals(10.5, paragraph.getRuns().getLast().getFontSizeAsDouble());
            assertFalse(paragraph.getCTP().getPPr().isSetKeepLines());
            assertFalse(paragraph.getCTP().getPPr().getSpacing().isSetLineRule());
        }
        question.setContentFlow("flow"); question.setPageBreakBefore(true);
        var mapper = new ObjectMapper();
        var restored = mapper.readValue(mapper.writeValueAsBytes(request("exam", question)), PaperExportRequest.class);
        assertEquals("exam", restored.getPaper().getTypography());
        assertEquals("flow", restored.getSections().getFirst().getQuestions().getFirst().getContentFlow());
        assertTrue(restored.getSections().getFirst().getQuestions().getFirst().getPageBreakBefore());
    }

    @Test void openingKeepsAreBoundedAndFinalConclusionStaysWithAnswer() throws Exception {
        var question = question("题干");
        question.setSolution("步骤甲<br/>步骤乙<br/>步骤丙<br/>结论"); question.setCorrect("答案");
        try (var doc = build(request("exam", question))) {
            var paragraphs = doc.getParagraphs();
            assertFalse(paragraphs.get(2).getCTP().getPPr().isSetKeepNext());
            assertTrue(paragraphs.get(3).getCTP().getPPr().isSetKeepNext());
            assertTrue(paragraphs.get(4).getCTP().getPPr().isSetKeepNext());
            assertFalse(paragraphs.getLast().getCTP().getPPr().isSetKeepNext());
        }
        question.setSolution(null); question.setCorrect(null);
        question.setContent(String.join("<br/>", java.util.Collections.nCopies(20, "长题干独立段落")));
        try (var doc = build(request("exam", question))) {
            assertTrue(doc.getParagraphs().stream().filter(p -> p.getCTP().getPPr().isSetKeepNext()).count() <= 5);
        }
    }
}
