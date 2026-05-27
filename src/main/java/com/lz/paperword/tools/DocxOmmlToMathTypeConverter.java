package com.lz.paperword.tools;

import com.lz.paperword.core.docx.MathTypeEmbedder;
import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * Convert an existing docx that stores equations as OMML into a docx that stores
 * the same equations as MathType OLE objects.
 *
 * <p>Pipeline:
 * <ol>
 *   <li>Use pandoc-generated .tex as the ordered LaTeX source of each equation</li>
 *   <li>Replace every {@code <m:oMath>} in document.xml with a temporary text placeholder</li>
 *   <li>Open the temporary docx with Apache POI and replace each placeholder run with MathType OLE</li>
 * </ol>
 */
public class DocxOmmlToMathTypeConverter {

    private static final Pattern INLINE_MATH_PATTERN =
        Pattern.compile("\\\\\\((.*?)\\\\\\)", Pattern.DOTALL);

    private static final Pattern OMML_PATTERN =
        Pattern.compile("<m:oMath\\b[^>]*>.*?</m:oMath>", Pattern.DOTALL);

    private static final Pattern PLACEHOLDER_PATTERN =
        Pattern.compile("__MT_PLACEHOLDER_(\\d+)__");

    private final LaTeXParser latexParser = new LaTeXParser();
    private final MathTypeEmbedder mathTypeEmbedder = new MathTypeEmbedder();

    public static void main(String[] args) throws Exception {
        Path inputDocx;
        Path inputTex;
        Path outputDocx;

        if (args.length >= 3) {
            inputDocx = Path.of(args[0]);
            inputTex = Path.of(args[1]);
            outputDocx = Path.of(args[2]);
        } else {
            String inputDocxProperty = System.getProperty("inputDocx");
            String inputTexProperty = System.getProperty("inputTex");
            String outputDocxProperty = System.getProperty("outputDocx");

            if (isBlank(inputDocxProperty) || isBlank(inputTexProperty) || isBlank(outputDocxProperty)) {
                System.err.println(
                    "Usage: DocxOmmlToMathTypeConverter <input.docx> <input.tex> <output.docx>"
                );
                System.err.println(
                    "Or pass -DinputDocx=... -DinputTex=... -DoutputDocx=..."
                );
                System.exit(1);
                return;
            }

            inputDocx = Path.of(inputDocxProperty);
            inputTex = Path.of(inputTexProperty);
            outputDocx = Path.of(outputDocxProperty);
        }

        if (!Files.exists(inputDocx) || !Files.exists(inputTex)) {
            System.err.println("Usage: DocxOmmlToMathTypeConverter <input.docx> <input.tex> <output.docx>");
            System.exit(1);
        }

        ConversionResult result = new DocxOmmlToMathTypeConverter().convert(inputDocx, inputTex, outputDocx);
        System.out.printf(
            "Converted %d equations from %s to %s%n",
            result.convertedEquations(),
            inputDocx,
            outputDocx
        );
    }

    public ConversionResult convert(Path inputDocx, Path inputTex, Path outputDocx) throws IOException {
        List<String> latexFormulas = extractLatexFormulas(inputTex);
        PreparedDocx preparedDocx = replaceOmmlWithPlaceholders(Files.readAllBytes(inputDocx));

        if (preparedDocx.placeholderCount() != latexFormulas.size()) {
            throw new IOException(
                "公式数量不一致：docx 中 OMML=" + preparedDocx.placeholderCount()
                    + "，tex 中 LaTeX=" + latexFormulas.size()
            );
        }

        if (outputDocx.getParent() != null) {
            Files.createDirectories(outputDocx.getParent());
        }

        try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(preparedDocx.docxBytes()))) {
            int converted = replacePlaceholders(document, latexFormulas);
            if (converted != latexFormulas.size()) {
                throw new IOException(
                    "占位符替换不完整：期望=" + latexFormulas.size() + "，实际=" + converted
                );
            }

            try (OutputStream outputStream = Files.newOutputStream(outputDocx)) {
                document.write(outputStream);
            }
            return new ConversionResult(latexFormulas.size(), outputDocx);
        }
    }

    private List<String> extractLatexFormulas(Path inputTex) throws IOException {
        String content = Files.readString(inputTex, StandardCharsets.UTF_8);
        Matcher matcher = INLINE_MATH_PATTERN.matcher(content);
        List<String> formulas = new ArrayList<>();
        while (matcher.find()) {
            String latex = matcher.group(1)
                .replace('\r', ' ')
                .replace('\n', ' ')
                .trim();
            formulas.add(normalizeLatexFormula(latex));
        }
        return formulas;
    }

    /**
     * Normalize pandoc-exported LaTeX into the subset that our MathType pipeline
     * and preview renderer handle more consistently.
     */
    private String normalizeLatexFormula(String latex) {
        String normalized = latex;

        // Geometry problems expect the standard triangle marker "△ABC".
        // Pandoc often exports OMML triangles as \bigtriangleup, while the
        // reference MathType document uses \triangle and renders better.
        normalized = normalized.replace("\\bigtriangleup", "\\triangle");

        return normalized;
    }

    private PreparedDocx replaceOmmlWithPlaceholders(byte[] inputDocxBytes) throws IOException {
        int placeholderCount = 0;
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(inputDocxBytes));
             ZipOutputStream zipOutputStream = new ZipOutputStream(outputStream)) {

            ZipEntry entry;
            while ((entry = zipInputStream.getNextEntry()) != null) {
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zipOutputStream.putNextEntry(newEntry);

                byte[] entryBytes = zipInputStream.readAllBytes();
                if ("word/document.xml".equals(entry.getName())) {
                    String documentXml = new String(entryBytes, StandardCharsets.UTF_8);
                    StringBuffer replacedXml = new StringBuffer();
                    Matcher matcher = OMML_PATTERN.matcher(documentXml);

                    while (matcher.find()) {
                        String placeholder = placeholderText(placeholderCount++);
                        String replacement =
                            "<w:r><w:t xml:space=\"preserve\">" + placeholder + "</w:t></w:r>";
                        matcher.appendReplacement(replacedXml, Matcher.quoteReplacement(replacement));
                    }
                    matcher.appendTail(replacedXml);
                    entryBytes = replacedXml.toString().getBytes(StandardCharsets.UTF_8);
                }

                zipOutputStream.write(entryBytes);
                zipOutputStream.closeEntry();
                zipInputStream.closeEntry();
            }
        }

        return new PreparedDocx(outputStream.toByteArray(), placeholderCount);
    }

    private int replacePlaceholders(XWPFDocument document, List<String> latexFormulas) {
        List<XWPFParagraph> paragraphs = collectParagraphs(document);
        int converted = 0;

        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = new ArrayList<>(paragraph.getRuns());
            for (XWPFRun run : runs) {
                String runText = run.text();
                if (runText == null) {
                    continue;
                }

                Matcher matcher = PLACEHOLDER_PATTERN.matcher(runText.trim());
                if (!matcher.matches()) {
                    continue;
                }

                int formulaIndex = Integer.parseInt(matcher.group(1));
                String latex = latexFormulas.get(formulaIndex);
                LaTeXNode ast = latexParser.parseLaTeX(latex);
                mathTypeEmbedder.embedEquation(paragraph, run, ast, latex);
                converted++;
            }
        }

        return converted;
    }

    private List<XWPFParagraph> collectParagraphs(XWPFDocument document) {
        List<XWPFParagraph> paragraphs = new ArrayList<>();
        collectBodyElements(document, paragraphs);

        for (XWPFHeader header : document.getHeaderList()) {
            collectBodyElements(header, paragraphs);
        }
        for (XWPFFooter footer : document.getFooterList()) {
            collectBodyElements(footer, paragraphs);
        }

        return paragraphs;
    }

    private void collectBodyElements(IBody body, List<XWPFParagraph> paragraphs) {
        for (IBodyElement bodyElement : body.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph paragraph) {
                paragraphs.add(paragraph);
            } else if (bodyElement instanceof XWPFTable table) {
                collectTableParagraphs(table, paragraphs);
            }
        }
    }

    private void collectTableParagraphs(XWPFTable table, List<XWPFParagraph> paragraphs) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                collectBodyElements(cell, paragraphs);
            }
        }
    }

    private String placeholderText(int index) {
        return "__MT_PLACEHOLDER_" + index + "__";
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }

    private record PreparedDocx(byte[] docxBytes, int placeholderCount) {}

    public record ConversionResult(int convertedEquations, Path outputPath) {}
}
