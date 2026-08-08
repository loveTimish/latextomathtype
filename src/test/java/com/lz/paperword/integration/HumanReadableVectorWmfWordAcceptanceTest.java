package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.docx.MathTypeEmbedder;
import com.lz.paperword.core.latex.LaTeXParser;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class HumanReadableVectorWmfWordAcceptanceTest {

    private static final Path CORPUS = Path.of(
        "analysis/formula-golden-corpus/formula-golden-corpus-report-latest.json");
    private static final Path OUTPUT = Path.of(System.getProperty(
        "paperword.acceptance.humanWord.output",
        "target/vector-acceptance/vector-emfplus-dual-quality-review.docx"));
    /** 15pt review size: readable at 125% without the heavy 18.6pt appearance of the old 1.55x draft. */
    private static final double DISPLAY_SCALE = 1.25d;
    private static final Map<String, String> REVIEW_LATEX_OVERRIDES = Map.of(
        "k12-array-01",
        "\\left\\{\\begin{array}{c}x+y=1\\\\x-y=2\\end{array}\\right."
    );

    @Test
    void reviewEquationSystemUsesAnExplicitLeftBrace() {
        String latex = REVIEW_LATEX_OVERRIDES.get("k12-array-01");

        assertTrue(latex.startsWith("\\left\\{"));
        assertTrue(latex.endsWith("\\right."));
        assertTrue(new LaTeXParser().parseDetailed(latex).isSupported());
    }

    @Test
    void generateHumanReadableVectorQualityDocument() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.humanWord"),
            "Enable with -Dpaperword.acceptance.humanWord=true");

        JsonNode cases = new ObjectMapper().readTree(CORPUS.toFile()).path("cases");
        assertEquals(28, cases.size());

        LaTeXParser parser = new LaTeXParser();
        MathTypeEmbedder embedder = new MathTypeEmbedder();
        try (XWPFDocument document = new XWPFDocument()) {
            configureA4Page(document);
            addTitle(document);

            for (int index = 0; index < cases.size(); index++) {
                if (index == 15) {
                    XWPFParagraph pageBreak = document.createParagraph();
                    pageBreak.createRun().addBreak(BreakType.PAGE);
                }

                JsonNode item = cases.get(index);
                String id = item.path("id").asText();
                String latex = REVIEW_LATEX_OVERRIDES.getOrDefault(
                    id, item.path("latex").asText());
                LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(latex);
                assertTrue(parsed.isSupported(), () -> id + ": " + parsed.diagnostics());

                XWPFParagraph paragraph = document.createParagraph();
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                paragraph.setSpacingBefore(0);
                paragraph.setSpacingAfter(150);

                XWPFRun label = paragraph.createRun();
                label.setText(id + ":  ");
                label.setFontFamily("Arial");
                label.setFontSize(10);
                label.setColor("404040");

                XWPFRun equation = paragraph.createRun();
                embedder.embedEquation(paragraph, equation, parsed.ast(), latex,
                    DISPLAY_SCALE, 420.0d);
            }

            Files.createDirectories(OUTPUT.getParent());
            try (OutputStream output = Files.newOutputStream(OUTPUT)) {
                document.write(output);
            }
        }

        assertTrue(Files.size(OUTPUT) > 1_000);
    }

    private static void configureA4Page(XWPFDocument document) {
        CTSectPr section = document.getDocument().getBody().isSetSectPr()
            ? document.getDocument().getBody().getSectPr()
            : document.getDocument().getBody().addNewSectPr();
        CTPageSz pageSize = section.isSetPgSz() ? section.getPgSz() : section.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(11_906));
        pageSize.setH(BigInteger.valueOf(16_838));
        CTPageMar margins = section.isSetPgMar() ? section.getPgMar() : section.addNewPgMar();
        margins.setTop(BigInteger.valueOf(850));
        margins.setBottom(BigInteger.valueOf(850));
        margins.setLeft(BigInteger.valueOf(1_150));
        margins.setRight(BigInteger.valueOf(1_150));
    }

    private static void addTitle(XWPFDocument document) {
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        title.setSpacingAfter(90);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("Editable MathType Exact-Outline Vector Quality Review");
        titleRun.setBold(true);
        titleRun.setFontFamily("Arial");
        titleRun.setFontSize(15);

        XWPFParagraph note = document.createParagraph();
        note.setAlignment(ParagraphAlignment.CENTER);
        note.setSpacingAfter(180);
        XWPFRun noteRun = note.createRun();
        noteRun.setText("28 representative editable MathType OLE formulas - exact MathJax outlines in EMF+ Dual");
        noteRun.setFontFamily("Arial");
        noteRun.setFontSize(9);
        noteRun.setColor("666666");
    }
}
