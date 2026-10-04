package com.lz.paperword.core.mathml;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefWriter;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilderFactory;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.TimeUnit;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

class MixedLongDivisionMathMlWriterTest {
    private static final String DIVISION = "\\begin{longdivision}{rr}{6}{2}{12}&12\\\\\\cline{1-2}&0\\end{longdivision}";
    private static final String SECOND = "\\begin{longdivision}{rr}{3}{5}{15}&15\\\\\\cline{1-2}&0\\end{longdivision}";
    private final LaTeXParser parser = new LaTeXParser();
    private final LongDivisionMathMlWriter writer = new LongDivisionMathMlWriter();

    @Test
    void pureDivisionKeepsEstablishedSerialization() {
        MathIRNode ir = parse(DIVISION);
        assertEquals(writer.write(ir.child(0)), writer.writeExpression(ir));
        assertEquals(writer.write(ir.child(0)), writer.writeExpression(parse("{" + DIVISION + "}")));
    }

    @Test
    void preservesPrefixSuffixAndBothDivisionsInOrder() throws Exception {
        String mathml = writer.writeExpression(parse("x+" + DIVISION + "=" + SECOND + "+y"));
        Document document = document(mathml);
        Element row = (Element) document.getDocumentElement().getFirstChild();
        assertEquals("mrow", row.getTagName());
        assertEquals(7, row.getChildNodes().getLength());
        assertEquals("x", row.getChildNodes().item(0).getTextContent());
        assertEquals("+", row.getChildNodes().item(1).getTextContent());
        assertEquals("semantics", row.getChildNodes().item(2).getNodeName());
        assertEquals("=", row.getChildNodes().item(3).getTextContent());
        assertEquals("semantics", row.getChildNodes().item(4).getNodeName());
        assertEquals("+", row.getChildNodes().item(5).getTextContent());
        assertEquals("y", row.getChildNodes().item(6).getTextContent());
        assertEquals(2, document.getElementsByTagName("mlongdiv").getLength());
    }

    @Test
    void preservesNestedFractionRootScriptsAndInvisibleFence() throws Exception {
        String latex = "\\left.\\frac{\\sqrt[3]{" + DIVISION + "}}{z_i^{2}}\\right|";
        Document document = document(writer.writeExpression(parse(latex)));
        Element root = (Element) document.getElementsByTagName("mroot").item(0);
        assertEquals("semantics", root.getFirstChild().getFirstChild().getNodeName());
        assertEquals("3", root.getLastChild().getTextContent());
        assertEquals(1, document.getElementsByTagName("mfrac").getLength());
        assertEquals(1, document.getElementsByTagName("msubsup").getLength());
        assertFalse(document.getDocumentElement().getTextContent().contains("."));
        assertEquals("|", document.getDocumentElement().getLastChild().getLastChild().getLastChild().getTextContent());
    }

    @Test
    void supportsDivisionAsScriptAndInsideAnotherDivision() throws Exception {
        String superscript = writer.writeExpression(parse("x^{" + DIVISION + "}"));
        assertEquals(1, document(superscript).getElementsByTagName("msup").getLength());
        assertTrue(superscript.contains("<msup><mi>x</mi><mrow><semantics>"));
        String nested = writer.writeExpression(parse(
            "\\begin{longdivision}{r}{x}{}{y}" + DIVISION + "\\end{longdivision}"));
        // One outer semantic annotation plus one copy of the inner division in each branch.
        assertEquals(3, document(nested).getElementsByTagName("mlongdiv").getLength());
    }

    @Test
    void preservesFontColorSizeAndMathStyleScopes() {
        String mathml = writer.writeExpression(parse(
            "\\mathbf{x}+{\\color{red}\\scriptstyle " + DIVISION + "}+{\\large y}"));
        assertTrue(mathml.contains("<mstyle mathvariant=\"bold\">"));
        assertTrue(mathml.contains("<mstyle mathcolor=\"#ff0000\">"));
        assertTrue(mathml.contains("<mstyle displaystyle=\"false\" scriptlevel=\"1\">"));
        assertTrue(mathml.contains("<mstyle mathsize=\"14.4pt\">"));
        assertTrue(mathml.endsWith("</mrow></math>"));
    }

    @Test
    void preservesAccentsArrowsBracesAndEnclosuresInsteadOfFlattening() {
        String mathml = writer.writeExpression(parse(
            "\\overline{x}+\\underbrace{a+b}_{n}+\\boxed{\\hat{z}}+\\xrightarrow[t]{u}+" + DIVISION));
        assertTrue(mathml.contains("<mover accent=\"true\">"));
        assertTrue(mathml.contains("<munder accentunder=\"true\">"));
        assertTrue(mathml.contains("<menclose notation=\"box\">"));
        assertTrue(mathml.contains("<munderover><mo stretchy=\"true\" minsize=\"1.75em\">→</mo>"));
        assertTrue(mathml.contains("⏟"));
        assertTrue(mathml.contains("¯"));
    }

    @Test
    void preservesAlignedTablesCasesAndPartitionRules() {
        String mathml = writer.writeExpression(parse(
            "\\begin{cases}" + DIVISION + "&x\\\\y&z\\end{cases}"));
        assertTrue(mathml.contains("<mo fence=\"true\" stretchy=\"true\">{</mo>"));
        assertTrue(mathml.contains("<mtable columnalign=\"left left\""));
        assertFalse(mathml.contains(">.</mo>"));
        String ruled = writer.writeExpression(parse(
            "\\begin{array}{|r|l|}\\hline " + DIVISION + "&x\\\\\\hline y&z\\\\\\hline\\end{array}"));
        assertTrue(ruled.contains("notation=\"left right top bottom\""));
        assertTrue(ruled.contains("columnlines=\"solid\" rowlines=\"solid\""));
    }

    @Test
    void failsClosedForUnsupportedOrMalformedNodesAnywhere() {
        MathIRNode expression = parse("x+" + DIVISION);
        expression.addChild(new MathIRNode(MathIRNode.Type.UNSUPPORTED, "\\unknown"));
        assertThrows(IllegalArgumentException.class, () -> writer.writeExpression(expression));
        MathIRNode malformed = new MathIRNode(MathIRNode.Type.FRACTION);
        malformed.addChild(parse(DIVISION));
        malformed.addChild(new MathIRNode(MathIRNode.Type.NUMBER, "1"));
        malformed.addChild(new MathIRNode(MathIRNode.Type.NUMBER, "2"));
        assertThrows(IllegalArgumentException.class, () -> writer.writeExpression(malformed));
        assertThrows(IllegalArgumentException.class, () -> writer.writeExpression(parse(
            "\\style{font-weight:bold}{x}+" + DIVISION)));
    }

    @Test
    void serializingWholeExpressionDoesNotChangeMtefSemantics() {
        MathIRNode expression = parse("x+\\frac{" + DIVISION + "}{z_i}+" + SECOND + "=y");
        String before = new MathIRConverter().dump(expression);
        byte[] mtef = new MtefWriter().write(expression);
        writer.writeExpression(expression);
        assertEquals(before, new MathIRConverter().dump(expression));
        assertArrayEquals(mtef, new MtefWriter().write(expression));
    }

    @Test
    void rendersMixedAndNestedExpressionsThroughPinnedMathJaxWithoutErrorGlyphs() throws Exception {
        List<String> formulas = List.of(
            DIVISION,
            "x+" + DIVISION + "+y",
            DIVISION + "+" + SECOND,
            "\\frac{" + DIVISION + "}{\\sqrt[3]{z_i^{2}}}",
            "\\left(" + DIVISION + "\\right)^{2}",
            "\\mathbf{x}+{\\color{red}\\scriptstyle " + DIVISION + "}",
            "\\overline{x}+\\underbrace{a+b}_{n}+\\boxed{z}+\\xrightarrow[t]{u}+" + DIVISION,
            "\\begin{cases}" + DIVISION + "&x\\\\y&z\\end{cases}");
        List<JsonNode> responses = renderMathJax(formulas);
        double pureWidth = responses.get(0).path("widthPt").asDouble();
        for (int i = 0; i < responses.size(); i++) {
            JsonNode response = responses.get(i);
            String svg = svg(response);
            if (i == 1 || i == 2) assertTrue(response.path("widthPt").asDouble() > pureWidth, formulas.get(i));
            if (i == 2) assertEquals(2, occurrences(svg, "data-mml-node=\"semantics\""));
            if (i == 5) assertTrue(svg.contains("#ff0000"), "red scope must reach SVG");
        }
    }

    @Test
    void raiseboxAdjustsBothDimensionsForEverySupportedUnit() {
        for (String unit : List.of("pt", "px", "em", "ex", "")) {
            String normalizedUnit = unit.isEmpty() ? "pt" : unit;
            String up = writer.writeExpression(parse("\\raisebox{20" + unit + "}{x}+" + DIVISION));
            String down = writer.writeExpression(parse("\\raisebox{-20" + unit + "}{x}+" + DIVISION));
            assertTrue(up.contains("voffset=\"+20" + normalizedUnit + "\" height=\"+20" + normalizedUnit
                + "\" depth=\"-20" + normalizedUnit + "\""));
            assertTrue(down.contains("voffset=\"-20" + normalizedUnit + "\" height=\"-20" + normalizedUnit
                + "\" depth=\"+20" + normalizedUnit + "\""));
        }
        MathIRNode expression = parse("\\raisebox{20pt}{x}+" + DIVISION);
        expression.child(0).setMetadata("verticalShift", "20cm");
        assertThrows(IllegalArgumentException.class, () -> writer.writeExpression(expression));
    }

    @Test
    void raisedAndLoweredInkFitsLiveMathJaxFrameBesideAndInsideDivision() throws Exception {
        List<String> formulas = List.of(
            "x+y+" + DIVISION,
            "\\raisebox{60pt}{x}+y+" + DIVISION,
            "\\raisebox{-60pt}{x}+y+" + DIVISION,
            DIVISION,
            DIVISION.replace("{6}", "{\\raisebox{60pt}{6}}"),
            DIVISION.replace("{6}", "{\\raisebox{-60pt}{6}}"));
        List<JsonNode> responses = renderMathJax(formulas);
        for (int base : List.of(0, 3)) {
            double originalHeight = responses.get(base).path("heightPt").asDouble();
            for (int shifted : List.of(base + 1, base + 2)) {
                assertTrue(responses.get(shifted).path("heightPt").asDouble() > originalHeight + 10,
                    "the frame must grow for " + formulas.get(shifted));
                assertSvgInkWithinViewBox(svg(responses.get(shifted)));
            }
        }
        assertTrue(responses.get(2).path("depthPt").asDouble() > responses.get(0).path("depthPt").asDouble() + 10,
            "lowered sibling must enlarge depth as well as the frame");
    }

    @Test
    void binomialDisplayAndTextQualifiersSurviveMathJaxSerialization() throws Exception {
        String display = "\\dbinom{\\frac{a}{b}}{c}+" + DIVISION;
        String text = "\\tbinom{\\frac{a}{b}}{c}+" + DIVISION;
        String displayMathml = writer.writeExpression(parse(display));
        String textMathml = writer.writeExpression(parse(text));
        assertTrue(displayMathml.contains("<mstyle displaystyle=\"true\" scriptlevel=\"0\">"));
        assertTrue(textMathml.contains("<mstyle displaystyle=\"false\" scriptlevel=\"0\">"));
        assertTrue(displayMathml.contains("<mfrac linethickness=\"0\">"));
        List<JsonNode> rendered = renderMathJax(List.of(display, text));
        assertTrue(rendered.get(0).path("widthPt").asDouble() > rendered.get(1).path("widthPt").asDouble(),
            "display binomial must retain larger contents and fences");
    }

    private void assertSvgInkWithinViewBox(String svg) throws Exception {
        Document document = document(svg);
        String[] parts = document.getDocumentElement().getAttribute("viewBox").split("\\s+");
        java.awt.geom.Rectangle2D frame = new java.awt.geom.Rectangle2D.Double(
            Double.parseDouble(parts[0]) - 1, Double.parseDouble(parts[1]) - 1,
            Double.parseDouble(parts[2]) + 2, Double.parseDouble(parts[3]) + 2);
        org.w3c.dom.NodeList paths = document.getElementsByTagName("path");
        assertTrue(paths.getLength() > 0);
        for (int i = 0; i < paths.getLength(); i++) {
            Element path = (Element) paths.item(i);
            java.awt.geom.AffineTransform transform = new java.awt.geom.AffineTransform();
            for (org.w3c.dom.Node ancestor = path; ancestor instanceof Element element; ancestor = ancestor.getParentNode()) {
                String attribute = element.getAttribute("transform");
                if (!attribute.isBlank()) {
                    transform.preConcatenate(org.apache.batik.parser.AWTTransformProducer.createAffineTransform(attribute));
                }
            }
            java.awt.Shape shape = org.apache.batik.parser.AWTPathProducer.createShape(
                new StringReader(path.getAttribute("d")), java.awt.geom.Path2D.WIND_NON_ZERO);
            java.awt.geom.Rectangle2D bounds = transform.createTransformedShape(shape).getBounds2D();
            assertTrue(frame.contains(bounds), "SVG path lies outside the viewBox: " + bounds + " vs " + frame);
        }
    }

    private String svg(JsonNode response) {
        return new String(Base64.getDecoder().decode(response.path("svgBase64").asText()), StandardCharsets.UTF_8);
    }

    private List<JsonNode> renderMathJax(List<String> formulas) throws Exception {
        Path node = Path.of(System.getProperty("paperword.mathjax.node.command", "node_modules/node/bin/node"));
        assumeTrue(Files.isExecutable(node), "local pinned Node is required for MathJax smoke");
        ObjectMapper mapper = new ObjectMapper();
        Process process = new ProcessBuilder(node.toAbsolutePath().toString(),
            "tools/mathjax/render_mathjax_svg.cjs", "--worker").redirectError(ProcessBuilder.Redirect.INHERIT).start();
        CompletableFuture<String> output = CompletableFuture.supplyAsync(() -> {
            try {
                return new String(process.getInputStream().readAllBytes(), StandardCharsets.UTF_8);
            } catch (java.io.IOException error) {
                throw new java.io.UncheckedIOException(error);
            }
        });
        try {
            for (int i = 0; i < formulas.size(); i++) {
                String mathml = writer.writeExpression(parse(formulas.get(i)));
                String request = mapper.writeValueAsString(Map.of("id", i, "inputFormat", "mathml",
                    "sourceBase64", Base64.getEncoder().encodeToString(mathml.getBytes(StandardCharsets.UTF_8))));
                process.getOutputStream().write((request + "\n").getBytes(StandardCharsets.UTF_8));
            }
            process.getOutputStream().close();
            assertTrue(process.waitFor(45, TimeUnit.SECONDS), "MathJax worker timed out");
            assertEquals(0, process.exitValue());
            List<String> responses = output.get(5, TimeUnit.SECONDS).lines().filter(line -> !line.isBlank()).toList();
            assertEquals(formulas.size(), responses.size());
            List<JsonNode> results = new java.util.ArrayList<>();
            for (int i = 0; i < responses.size(); i++) {
                JsonNode response = mapper.readTree(responses.get(i));
                assertTrue(response.path("ok").asBoolean(), formulas.get(i) + ": " + response);
                assertFalse(svg(response).contains("data-mml-node=\"merror\""), formulas.get(i));
                assertTrue(svg(response).contains("<path"), formulas.get(i));
                results.add(response);
            }
            return results;
        } finally {
            process.destroyForcibly();
        }
    }

    private MathIRNode parse(String latex) {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed(latex);
        assertTrue(result.isSupported(), latex + ": " + result.diagnostics());
        return result.mathIR();
    }

    private Document document(String mathml) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        return factory.newDocumentBuilder().parse(new InputSource(new StringReader(mathml)));
    }

    private int occurrences(String value, String needle) {
        return (value.length() - value.replace(needle, "").length()) / needle.length();
    }
}
