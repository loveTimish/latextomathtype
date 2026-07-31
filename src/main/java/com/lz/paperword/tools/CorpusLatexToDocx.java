package com.lz.paperword.tools;

import com.lz.paperword.core.docx.MathTypeEmbedder;
import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Build a docx from a corpus of LaTeX formulas (one per paragraph, each with
 * its image-name label), embedding every formula as a MathType OLE object
 * with the current renderer preview. Used to eyeball the effect of measured
 * structure parameters in Word.
 *
 * <p>Usage: CorpusLatexToDocx &lt;latex.tsv&gt; &lt;out.docx&gt; [limit]
 * TSV format: image-name&lt;TAB&gt;latex (see target/wmf-ruler/wmf-latex.tsv).
 */
public class CorpusLatexToDocx {

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("usage: CorpusLatexToDocx <latex.tsv> <out.docx> [limit]");
            System.exit(2);
        }
        Path tsv = Path.of(args[0]);
        Path out = Path.of(args[1]);
        int limit = args.length > 2 ? Integer.parseInt(args[2]) : Integer.MAX_VALUE;

        LaTeXParser parser = new LaTeXParser();
        MathTypeEmbedder embedder = new MathTypeEmbedder();
        embedder.resetDocumentFormulaCounter();

        int done = 0;
        int failed = 0;
        try (XWPFDocument doc = new XWPFDocument()) {
            List<String> lines = Files.readAllLines(tsv);
            for (String line : lines) {
                if (done >= limit) {
                    break;
                }
                String[] parts = line.split("\t", 2);
                if (parts.length != 2 || parts[1].isBlank()) {
                    continue;
                }
                String name = Path.of(parts[0]).getFileName().toString()
                    .replaceFirst("\\.bin$", "");
                // corpus TSV keeps the original $$ ... $$ display wrappers and
                // inter-token spaces; parseLaTeX expects bare math
                String latex = parts[1].trim()
                    .replaceFirst("^\\$\\$", "")
                    .replaceFirst("\\$\\$$", "")
                    .trim();
                XWPFParagraph p = doc.createParagraph();
                XWPFRun label = p.createRun();
                label.setText(name + ": ");
                label.setFontSize(8);
                try {
                    LaTeXNode ast = parser.parseLaTeX(latex);
                    embedder.embedEquation(p, p.createRun(), ast, latex);
                    done++;
                } catch (Exception e) {
                    XWPFRun err = p.createRun();
                    err.setText("[EMBED FAILED: " + e.getMessage() + "]");
                    failed++;
                }
            }
            out.toAbsolutePath().getParent().toFile().mkdirs();
            try (OutputStream os = Files.newOutputStream(out)) {
                doc.write(os);
            }
        }
        System.out.println("embedded=" + done + " failed=" + failed + " -> " + out);
    }
}
