package com.lz.paperword.core.mtef;

import com.lz.paperword.core.docx.MathTypeEmbedder;
import com.lz.paperword.core.latex.LaTeXParser;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/** Generates a DOCX containing editable MathType objects for supplied LaTeX formulas. */
public final class MathTypeFormulaDocxCli {

    private MathTypeFormulaDocxCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("usage: MathTypeFormulaDocxCli <output.docx> <latex> [<latex> ...]");
            System.exit(2);
        }
        Path output = Path.of(args[0]).toAbsolutePath();
        if (output.getParent() != null) {
            Files.createDirectories(output.getParent());
        }
        LaTeXParser parser = new LaTeXParser();
        MathTypeEmbedder embedder = new MathTypeEmbedder();
        embedder.resetDocumentFormulaCounter();
        try (XWPFDocument document = new XWPFDocument()) {
            for (int index = 1; index < args.length; index++) {
                String latex = args[index];
                LaTeXParser.DetailedParseResult parsed = parser.parseDetailed(latex);
                if (!parsed.isSupported()) {
                    throw new IllegalArgumentException(latex + ": " + parsed.diagnostics());
                }
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                embedder.embedEquation(paragraph, run, parsed.ast(), latex);
            }
            try (OutputStream stream = Files.newOutputStream(output)) {
                document.write(stream);
            }
        }
        System.out.printf("generated %d Equation.DSMT4 objects: %s%n", args.length - 1, output);
    }
}
