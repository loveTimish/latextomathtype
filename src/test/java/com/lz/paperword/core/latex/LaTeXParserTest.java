package com.lz.paperword.core.latex;

import com.lz.paperword.core.latex.LaTeXParser.ContentSegment;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class LaTeXParserTest {

    private final LaTeXParser parser = new LaTeXParser();

    @Test
    void rejectsReplacementCharactersBeforeFormulaSplitting() {
        String text = "$\\pwmetrics{17.400,16.000}V_{� }$："
            + "$\\pwmetrics{17.400,16.000}V_{� }$=1：12,"
            + "$\\pwmetrics{17.400,16.000}V_{� }$："
            + "$\\pwmetrics{17.400,16.000}V_{� }$=1：16。";

        IllegalArgumentException failure = assertThrows(IllegalArgumentException.class,
            () -> parser.parseText(text));
        assertTrue(failure.getMessage().contains("SOURCE_REPLACEMENT_CHARACTER"));
    }

    @Test
    void rejectsUnrecoverableSourceReplacementCharacterExplicitly() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed("V_{�}");
        LaTeXParser.DetailedParseResult numeric = parser.parseDetailed("1�+2");

        assertFalse(result.isSupported());
        assertTrue(result.diagnostics().stream().anyMatch(diagnostic ->
            "SOURCE_REPLACEMENT_CHARACTER".equals(diagnostic.code())));
        assertFalse(numeric.isSupported());
        assertTrue(numeric.normalizedLatex().contains("�"),
            "unknown source characters must not be deleted during normalization");
        IllegalArgumentException failure = assertThrows(IllegalArgumentException.class,
            () -> parser.parseLaTeX("V_{�}"));
        assertTrue(failure.getMessage().contains("SOURCE_REPLACEMENT_CHARACTER"));
    }

    @Test
    void parseTextRejectsFormerCorpusRepairPatterns() {
        for (String text : List.of(
                "$\\mathrm{� 路程}=\\mathrm{速度和}$",
                "$(75+60)� \\times 20=� 2700$",
                "$80-75=� 5$",
                "(54-27)�千米",
                "plain�text")) {
            IllegalArgumentException failure = assertThrows(IllegalArgumentException.class,
                () -> parser.parseText(text), text);
            assertTrue(failure.getMessage().contains("SOURCE_REPLACEMENT_CHARACTER"), text);
        }
    }

    @Test
    void preNormalizeDoesNotRewriteCorpusSpecificSymbols() {
        String latex = "\\text{相遇{\\blacksquare}{\\blacksquare}}"
            + "+\\text{追及{\\blacksquare}{\\blacksquare}}"
            + "+\\text{不合{\\blacksquare}意}"
            + "+\\text{心想事\\Theta }"
            + "+\\text{梦想\\Theta 真}";

        assertEquals(latex, LaTeXParser.preNormalizeLatex(latex));
    }

    @Test
    void preNormalizeRemovesOuterDollarMathDelimiters() {
        assertEquals("\\sqrt{1+x^2}", LaTeXParser.preNormalizeLatex("$\\sqrt{1+x^2}$"));
        assertEquals("\\sqrt{1+x^2}", LaTeXParser.preNormalizeLatex("$$\\sqrt{1+x^2}$$"));
    }

    @Test
    void detailedParseReportsConsumedCommandsAndSupportedIr() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed("\\frac{\\alpha}{x}");

        assertTrue(result.isSupported());
        assertEquals(List.of("\\frac", "\\alpha"), result.consumedCommands());
        assertTrue(result.diagnostics().isEmpty());
        assertNotNull(result.mathIR());
    }

    @Test
    void detailedParseReportsUnknownCommands() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed("\\definitelyUnsupported{x}");

        assertFalse(result.isSupported());
        assertTrue(result.diagnostics().stream().anyMatch(diagnostic ->
            "UNSUPPORTED_COMMAND".equals(diagnostic.code())
                && "\\definitelyUnsupported".equals(diagnostic.command())));
    }

    @Test
    void detailedParsePreservesLegacyStyleAsSemanticScope() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed("{\\bf x}");

        assertTrue(result.isSupported());
        assertTrue(result.diagnostics().isEmpty());
        assertEquals("bold", result.mathIR().child(0).child(0).getMetadata("fontVariant"));
    }

    @Test
    void stripsOnlyStandaloneAlignmentMarkers() {
        LaTeXParser.DetailedParseResult standalone =
            parser.parseDetailed("&=20.08\\times (200.9-200.7)");
        LaTeXParser.DetailedParseResult array =
            parser.parseDetailed("\\begin{array}{rl}x&=1\\\\y&=2\\end{array}");
        LaTeXParser.DetailedParseResult escaped = parser.parseDetailed("A\\&B");

        assertEquals("=20.08\\times (200.9-200.7)", standalone.normalizedLatex());
        assertTrue(standalone.isSupported());
        assertEquals("\\begin{array}{rl}x&=1\\\\y&=2\\end{array}", array.normalizedLatex());
        assertEquals("A\\&B", escaped.normalizedLatex());
    }

    @Test
    void parsesRaiseboxAsVerticalShiftStyle() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed("\\raisebox{-3pt}{2}");

        assertTrue(result.isSupported());
        assertEquals("vertical-shift", result.mathIR().child(0).getMetadata("styleKind"));
        assertEquals("-3.0", result.mathIR().child(0).getMetadata("verticalShiftPt"));
        assertEquals("2", result.mathIR().child(0).child(0).child(0).getValue());
    }

    @Test
    void detailedParseRejectsRaiseboxOptionalBoxMetrics() {
        for (String latex : List.of(
                "\\raisebox{-3pt}[8pt]{2}",
                "\\raisebox{-3pt}[8pt][2pt]{2}")) {
            LaTeXParser.DetailedParseResult result = parser.parseDetailed(latex);

            assertFalse(result.isSupported(), latex);
            assertTrue(result.diagnostics().stream().anyMatch(diagnostic ->
                "UNSUPPORTED_RAISEBOX_OPTIONAL_METRICS".equals(diagnostic.code())
                    && "\\raisebox".equals(diagnostic.command())), latex);
        }
    }

    @Test
    void repairsDanglingDelimiterAndLooseInternalMetricsPrefix() {
        List<ContentSegment> segments = parser.parseText(
            "before$$\\pwmetrics 40.600,13.000 0.5 \\times 1$after");

        ContentSegment formula = segments.stream().filter(ContentSegment::isMath).findFirst().orElseThrow();
        assertEquals("0.5 \\times 1", formula.rawText());
        assertNotNull(formula.metrics());
        assertEquals(40.6d, formula.metrics().wmfWidthPt(), 0.001d);
        assertEquals(13.0d, formula.metrics().wmfHeightPt(), 0.001d);
    }

    @Test
    void keepsEscapedDollarSignsInsideInlineFormula() {
        List<ContentSegment> segments = parser.parseText(
            "before$\\text{\\$\\$}\\;\\backslash\\text{sqrt2 \\$\\$}$after");

        List<ContentSegment> formulas = segments.stream().filter(ContentSegment::isMath).toList();
        assertEquals(1, formulas.size());
        assertEquals("\\text{\\$\\$}\\;\\backslash\\text{sqrt2 \\$\\$}", formulas.get(0).rawText());
    }

    @Test
    void doesNotTreatArrayLineSpacingAsDisplayMathDelimiter() {
        String latex = "\\sum\\nolimits_{\\begin{array}{c}a\\\\[0.1em]b\\\\[0.1em]c\\end{array}}";
        List<ContentSegment> segments = parser.parseText("$" + latex + "$");

        List<ContentSegment> formulas = segments.stream().filter(ContentSegment::isMath).toList();
        assertEquals(1, formulas.size());
        assertEquals(latex, formulas.get(0).rawText());
    }

    @Test
    void keepsNestedMathDelimitersInsideTextGroup() {
        String latex = "\\mbox{\\large$T$}_0^2";
        List<ContentSegment> segments = parser.parseText("$" + latex + "$");

        List<ContentSegment> formulas = segments.stream().filter(ContentSegment::isMath).toList();
        assertEquals(1, formulas.size());
        assertEquals(latex, formulas.get(0).rawText());
    }

    @Test
    void hoistsArrayAlignmentMarkerOutOfFontStyleGroup() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed(
            "\\begin{array}{l}21x\\mathbf{&=}140\\\\20x\\mathbf{&=}65\\end{array}");

        assertTrue(result.isSupported());
        assertEquals(
            "\\begin{array}{l}21x&\\mathbf{=}140\\\\20x&\\mathbf{=}65\\end{array}",
            result.normalizedLatex());
        LaTeXNode array = result.ast().getChildren().get(0);
        assertEquals(2, array.getChildren().size());
        assertEquals(2, array.getChildren().get(0).getChildren().size());
        assertEquals(2, array.getChildren().get(1).getChildren().size());
    }

    @Test
    void repairsNestedMathInsideTextColorAndKeepsScopedRgb() {
        List<ContentSegment> segments = parser.parseText(
            "发现规律$\\pwmetrics{40.600,13.000}\\textcolor{maroon}{$\\div$"
                + "\\pwmetrics{17.400,13.000}16=}$\\frac 5 256 $");

        List<ContentSegment> formulas = segments.stream().filter(ContentSegment::isMath).toList();
        assertEquals(2, formulas.size());
        assertEquals("\\textcolor{maroon}{\\div16=}", formulas.get(0).rawText());
        assertEquals(40.6d, formulas.get(0).metrics().wmfWidthPt(), 0.001d);
        LaTeXParser.DetailedParseResult colored = parser.parseDetailed(formulas.get(0).rawText());
        assertTrue(colored.isSupported());
        assertEquals("rgb", colored.mathIR().child(0).getMetadata("colorModel"));
        assertEquals("0.502,0,0", colored.mathIR().child(0).getMetadata("colorValue"));
        assertEquals("\\frac 5 256", formulas.get(1).rawText());
    }

    @Test
    void removesLooseNonContentIncludeGraphicsFromVisibleText() {
        List<ContentSegment> segments = parser.parseText(
            "题图 \\includegraphics[width=1\\textwidth] embeddings/oleObject475.bin 后文");

        assertEquals(1, segments.size());
        assertFalse(segments.get(0).isMath());
        assertEquals("题图  后文", segments.get(0).rawText());
    }

    @Test
    void removesPlainTextTableLayoutWhilePreservingCellsAndFollowingFormula() {
        List<ContentSegment> segments = parser.parseText(
            "\\begin table \\begin tabularx \\textwidth |p \\dimexpr 0.5\\linewidth | & 纯循环小数 & 混循环小数 "
                + "\\end tabularx \\end table % D2T: Empty equation removed!$x=1$");

        assertTrue(segments.stream().noneMatch(segment -> segment.rawText().contains("\\begin")));
        assertTrue(segments.stream().noneMatch(segment -> segment.rawText().contains("\\end")));
        assertTrue(segments.stream().anyMatch(segment -> !segment.isMath()
            && segment.rawText().contains("纯循环小数") && segment.rawText().contains("混循环小数")));
        assertTrue(segments.stream().anyMatch(segment -> segment.isMath() && "x=1".equals(segment.rawText())));
    }

    @Test
    void removesPlainTextTablePreambleWhenFirstCellIsNotEmpty() {
        List<ContentSegment> segments = parser.parseText(
            "\\begin table \\begin tabularx \\textwidth |p \\dimexpr 0.5\\linewidth-2\\arrayrulewidth | "
                + "景区 & 千岛湖 \\end tabularx \\end table 填空\\_\\_");

        String visible = segments.stream().map(ContentSegment::rawText).reduce("", String::concat);
        assertFalse(visible.contains("\\begin"));
        assertFalse(visible.contains("arrayrulewidth"));
        assertTrue(visible.contains("景区") && visible.contains("千岛湖"));
        assertTrue(visible.contains("填空__"));
    }

    @Test
    void testParseSimpleHtml() {
        List<ContentSegment> segments = parser.parseHtml("<p>已知 $x=3$，求 $y$ 的值</p>");
        // segments: "已知 ", "x=3", "，求 ", "y", " 的值"
        assertEquals(5, segments.size());
        assertFalse(segments.get(0).isMath());
        assertTrue(segments.get(0).rawText().contains("已知"));
        assertTrue(segments.get(1).isMath());
        assertEquals("x=3", segments.get(1).rawText());
        assertTrue(segments.get(3).isMath());
        assertEquals("y", segments.get(3).rawText());
        assertFalse(segments.get(4).isMath());
        assertTrue(segments.get(4).rawText().contains("的值"));
    }

    @Test
    void testParseFraction() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{x+1}{2}");
        assertNotNull(ast);
        assertEquals(LaTeXNode.Type.ROOT, ast.getType());
        assertEquals(1, ast.getChildren().size());

        LaTeXNode frac = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.FRACTION, frac.getType());
        assertEquals(2, frac.getChildren().size()); // numerator + denominator
    }

    @Test
    void testParseSuperscript() {
        LaTeXNode ast = parser.parseLaTeX("x^{2}");
        assertNotNull(ast);
        LaTeXNode sup = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, sup.getType());
        assertEquals(2, sup.getChildren().size()); // base + exponent
    }

    @Test
    void testParseSpacedSuperscriptFromDocxToLatex() {
        LaTeXNode ast = parser.parseLaTeX("45 ^ { \\circ }");
        assertNotNull(ast);
        assertEquals(2, ast.getChildren().size());
        assertEquals("4", flatten(ast.getChildren().get(0)));
        LaTeXNode sup = ast.getChildren().get(1);
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, sup.getType());
        assertEquals("5", flatten(sup.getChildren().get(0)));
        assertEquals("\\circ", flatten(sup.getChildren().get(1)));
    }

    @Test
    void testParseArrayLineBreaksWrittenAsThinSpacesByDocxToLatex() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{c} a=1 \\\\,b=2 \\\\,c=3 \\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals(3, array.getChildren().size());
        assertEquals("a=1", flatten(array.getChildren().get(0)));
        assertEquals("b=2", flatten(array.getChildren().get(1)));
        assertEquals("c=3", flatten(array.getChildren().get(2)));
    }

    @Test
    void testParseArrayLineBreakBeforeCommandKeepsCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{c} 6 \\\\,\\div \\\\,8 \\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals(3, array.getChildren().size());
        assertEquals("6", flatten(array.getChildren().get(0)));
        assertEquals("\\div", flatten(array.getChildren().get(1)));
        assertEquals("8", flatten(array.getChildren().get(2)));
    }

    @Test
    void testParseSubscript() {
        LaTeXNode ast = parser.parseLaTeX("a_{n}");
        assertNotNull(ast);
        LaTeXNode sub = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, sub.getType());
    }

    @Test
    void testParseSqrt() {
        LaTeXNode ast = parser.parseLaTeX("\\sqrt{x+1}");
        assertNotNull(ast);
        LaTeXNode sqrt = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SQRT, sqrt.getType());
        assertEquals(1, sqrt.getChildren().size()); // content
    }

    @Test
    void testParseNthRoot() {
        LaTeXNode ast = parser.parseLaTeX("\\sqrt[3]{x}");
        assertNotNull(ast);
        LaTeXNode sqrt = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SQRT, sqrt.getType());
        assertEquals(2, sqrt.getChildren().size()); // degree + content
    }

    @Test
    void testParseGreekLetters() {
        LaTeXNode ast = parser.parseLaTeX("\\alpha + \\beta");
        assertNotNull(ast);
        assertTrue(ast.getChildren().size() >= 2);
        assertEquals(LaTeXNode.Type.COMMAND, ast.getChildren().get(0).getType());
        assertEquals("\\alpha", ast.getChildren().get(0).getValue());
    }

    @Test
    void testParseComplexExpression() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{-b \\pm \\sqrt{b^{2}-4ac}}{2a}");
        assertNotNull(ast);
        // Should have a fraction at root level
        LaTeXNode frac = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.FRACTION, frac.getType());
    }

    @Test
    void testParsePlainTextOnly() {
        List<ContentSegment> segments = parser.parseHtml("<p>这是纯文本内容</p>");
        assertEquals(1, segments.size());
        assertFalse(segments.get(0).isMath());
        assertEquals("这是纯文本内容", segments.get(0).rawText());
    }

    @Test
    void testParseTextRestoresSafeLatexEscapesOnlyInPlainSegments() {
        List<ContentSegment> segments = parser.parseText(
            "\\raisebox -3pt 范围1\\textasciitilde 5，填空\\_\\_ $a_b$ 百分比\\%\\raisebox -");

        assertEquals(3, segments.size());
        assertEquals("范围1~ 5，填空__", segments.get(0).rawText());
        assertTrue(segments.get(1).isMath());
        assertEquals("a_b", segments.get(1).rawText());
        assertEquals("百分比%", segments.get(2).rawText());
    }

    @Test
    void testParseTextUnwrapsRaiseboxWithoutDroppingPlainTextOrEscapedDollar() {
        List<ContentSegment> segments = parser.parseText(
            "前\\raisebox{-3pt}[8pt][2pt]{完整\\{内容\\}，价格\\$5，"
                + "\\raisebox{1pt}{嵌套}}后");

        assertEquals(1, segments.size());
        assertFalse(segments.get(0).isMath());
        assertEquals("前完整{内容}，价格$5，嵌套后", segments.get(0).rawText());
    }

    @Test
    void testParseEmptyInput() {
        List<ContentSegment> segments = parser.parseHtml("");
        assertTrue(segments.isEmpty());
    }

    @Test
    void testParseBareLatexAsPlainTextWhenMissingDelimiters() {
        List<ContentSegment> segments = parser.parseText("【解答】\\frac{1}{2}+\\frac{1}{3}");
        assertEquals(1, segments.size());
        assertFalse(segments.get(0).isMath());
        assertEquals("【解答】\\frac{1}{2}+\\frac{1}{3}", segments.get(0).rawText());
    }

    @Test
    void testParseMultilineDisplayMathKeepsWholeBlock() {
        List<ContentSegment> segments = parser.parseText("【解答】\n$$\\begin{array}{ccccc}\n{30\\%} & {} & {} & {} & {20\\%} \\\\\n{} & {\\searrow} & {} & {\\nearrow} & {}\n\\end{array}$$");
        assertEquals(2, segments.size());
        assertFalse(segments.get(0).isMath());
        assertTrue(segments.get(1).isMath());
        assertTrue(segments.get(1).rawText().contains("\\begin{array}{ccccc}"));
    }

    @Test
    void testParseSumWithLimits() {
        LaTeXNode ast = parser.parseLaTeX("\\sum_{i=1}^{n}a_i");
        assertNotNull(ast);
        assertTrue(ast.getChildren().size() > 0);
    }

    @Test
    void testParseCoproductWithLimits() {
        LaTeXNode ast = parser.parseLaTeX("\\coprod_{i=1}^{n} A_i");

        assertNotNull(ast);
        assertEquals(2, ast.getChildren().size());
        assertEquals(LaTeXNode.Type.SUPERSCRIPT, ast.getChildren().get(0).getType());

        LaTeXNode superscript = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, superscript.getChildren().get(0).getType());
        assertEquals("\\coprod", superscript.getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testParseArrayForVerticalAddition() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rrrr", array.getMetadata("columnSpec"));
        assertEquals("4", array.getMetadata("columnCount"));
        assertEquals("0,0,0,0,0", array.getMetadata("columnLines"));
        assertEquals("0,0,1,0", array.getMetadata("rowLines"));
        assertEquals(3, array.getChildren().size());

        LaTeXNode firstRow = array.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ROW, firstRow.getType());
        assertEquals(4, firstRow.getChildren().size());
    }

    @Test
    void testParseArrayForDecimalAlignment() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcr}12 & . & 50 \\\\ +3 & . & 75 \\\\ \\hline 16 & . & 25\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rcr", array.getMetadata("columnSpec"));
        assertEquals("3", array.getMetadata("columnCount"));
        assertEquals(3, array.getChildren().size());
        assertEquals(".", array.getChildren().get(0).getChildren().get(1).getChildren().get(0).getValue());
    }

    @Test
    void testParseArrayForVerticalSubtraction() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 8 & 6 & 4 \\\\ - & 2 & 7 & 9 \\\\ \\hline & 5 & 8 & 5\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rrrr", array.getMetadata("columnSpec"));
        assertEquals(3, array.getChildren().size());
    }

    @Test
    void testParseArrayForDecimalSubtraction() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcr}12 & . & 50 \\\\ -3 & . & 75 \\\\ \\hline 8 & . & 75\\end{array}");
        assertNotNull(ast);
        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("rcr", array.getMetadata("columnSpec"));
        assertEquals(3, array.getChildren().size());
    }

    @Test
    void testParseExplicitLongDivisionCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode longDiv = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.LONG_DIVISION, longDiv.getType());
        assertEquals("\\longdiv", longDiv.getValue());
        assertEquals(3, longDiv.getChildren().size());
        assertEquals("13", flatten(longDiv.getChildren().get(0)));
        assertEquals("65", flatten(longDiv.getChildren().get(1)));
        assertEquals("845", flatten(longDiv.getChildren().get(2)));
    }

    @Test
    void testParseExplicitLongDivisionCommandWithoutQuotient() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{13}{845}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode longDiv = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.LONG_DIVISION, longDiv.getType());
        assertEquals(3, longDiv.getChildren().size());
        assertEquals("13", flatten(longDiv.getChildren().get(0)));
        assertEquals("", flatten(longDiv.getChildren().get(1)));
        assertEquals("845", flatten(longDiv.getChildren().get(2)));
    }

    @Test
    void parsesStructuredLongDivisionWithoutComputingAnyRows() {
        LaTeXParser.DetailedParseResult result = parser.parseDetailed(
            "\\begin{longdivision}{rrrr}{6}{570}{3420}"
                + "&30\\\\\\cline{1-2}&&42\\\\&&42\\\\\\cline{2-3}&&&0"
                + "\\end{longdivision}");

        assertTrue(result.isSupported(), () -> result.diagnostics().toString());
        LaTeXNode longDivision = result.ast().getChildren().get(0);
        assertEquals(LaTeXNode.Type.LONG_DIVISION, longDivision.getType());
        assertEquals("true", longDivision.getMetadata("structured"));
        assertEquals("rrrr", longDivision.getMetadata("columnSpec"));
        assertEquals(4, longDivision.getChildren().size());
        assertEquals("6", flatten(longDivision.getChildren().get(0)));
        assertEquals("570", flatten(longDivision.getChildren().get(1)));
        assertEquals("3420", flatten(longDivision.getChildren().get(2)));

        LaTeXNode steps = longDivision.getChildren().get(3);
        assertEquals(4, steps.getChildren().size());
        assertEquals("2", steps.getChildren().get(0).getMetadata("endColumn"));
        assertEquals("1", steps.getChildren().get(0).getMetadata("ruleStartColumn"));
        assertEquals("2", steps.getChildren().get(0).getMetadata("ruleEndColumn"));
        assertEquals("3", steps.getChildren().get(2).getMetadata("ruleEndColumn"));
        assertEquals("4", steps.getChildren().get(3).getMetadata("endColumn"));
        assertEquals("30", flatten(steps.getChildren().get(0)));
        assertEquals("42", flatten(steps.getChildren().get(1)));
        assertEquals("42", flatten(steps.getChildren().get(2)));
        assertEquals("0", flatten(steps.getChildren().get(3)));
    }

    @Test
    void rejectsMalformedStructuredLongDivisionInsteadOfFillingItIn() {
        LaTeXParser.DetailedParseResult invalidSpec = parser.parseDetailed(
            "\\begin{longdivision}{rc}{6}{}{12}&12\\end{longdivision}");
        LaTeXParser.DetailedParseResult multipleCells = parser.parseDetailed(
            "\\begin{longdivision}{rr}{6}{}{12}1&2\\end{longdivision}");
        LaTeXParser.DetailedParseResult invalidRule = parser.parseDetailed(
            "\\begin{longdivision}{rr}{6}{}{12}&12\\\\\\cline{1-3}&0\\end{longdivision}");
        LaTeXParser.DetailedParseResult orphanRule = parser.parseDetailed(
            "\\begin{longdivision}{rr}{6}{}{12}\\cline{1-2}&12\\end{longdivision}");

        assertTrue(invalidSpec.diagnostics().stream()
            .anyMatch(d -> "LONG_DIVISION_COLUMN_SPEC".equals(d.code())));
        assertTrue(multipleCells.diagnostics().stream()
            .anyMatch(d -> "LONG_DIVISION_MULTIPLE_CELLS".equals(d.code())));
        assertTrue(invalidRule.diagnostics().stream()
            .anyMatch(d -> "LONG_DIVISION_CLINE_RANGE".equals(d.code())));
        assertTrue(orphanRule.diagnostics().stream()
            .anyMatch(d -> "LONG_DIVISION_ORPHAN_CLINE".equals(d.code())));
    }

    @Test
    void testParseCompositeLongDivisionWithStepArrayInSingleBlock() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\longdiv[570]{6}{3420}\\begin{array}{l}\\text{   }\\underline{30}\\\\\\text{    }42\\\\\\text{    }\\underline{42}\\\\\\text{      }0\\end{array}"
        );
        assertNotNull(ast);
        assertEquals(2, ast.getChildren().size());
        assertEquals(LaTeXNode.Type.LONG_DIVISION, ast.getChildren().get(0).getType());
        assertEquals(LaTeXNode.Type.ARRAY, ast.getChildren().get(1).getType());
        assertEquals("l", ast.getChildren().get(1).getMetadata("columnSpec"));
        assertEquals(4, ast.getChildren().get(1).getChildren().size());
        assertTrue(flatten(ast.getChildren().get(1)).contains("    42"), "\\text{空格} 中的前导空格应被完整保留");
    }

    @Test
    void testParseConcentrationCrossArrayKeepsArrowCommandsInMathAst() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {10\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {20\\%} & {} & {} & {} & {20\\%}\\end{array}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("ccccc", array.getMetadata("columnSpec"));
        assertEquals(5, array.getChildren().size());
        assertTrue(flatten(array).contains("\\searrow"));
        assertTrue(flatten(array).contains("\\nearrow"));
    }

    @Test
    void testPreNormalizeKeepsArrayLineBreakBeforeSpace() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{c}1 \\\\ 2\\end{array}");

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(2, array.getChildren().size(),
            "array row break followed by a space must not be normalized as a control-space command");
    }

    @Test
    void trailingArrayRowBreakDoesNotCreateAnImplicitExtraRow() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{l}\\\\\\\\xy\\\\2\\\\\\\\\\end{array}");
        LaTeXNode array = ast.getChildren().get(0);

        assertEquals(5, array.getChildren().size());
    }

    @Test
    void unbracedTextCommandIgnoresItsDelimiterWhitespace() {
        LaTeXNode ast = parser.parseLaTeX("\\text x");
        LaTeXNode text = ast.getChildren().get(0);

        assertEquals("x", text.getChildren().get(0).getValue());
    }

    @Test
    void testParseMatrixEnvironmentPromotesToArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{matrix}1&2\\\\3&4\\end{matrix}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("matrix", array.getMetadata("environment"));
        assertEquals("cc", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParsePmatrixEnvironmentWrapsArrayWithFence() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("\\left(", fence.getValue());
        assertEquals("(", fence.getMetadata("leftDelimiter"));
        assertEquals(")", fence.getMetadata("rightDelimiter"));
        assertEquals(1, fence.getChildren().size());
        assertEquals(LaTeXNode.Type.ARRAY, fence.getChildren().get(0).getType());
        assertEquals("pmatrix", fence.getChildren().get(0).getMetadata("environment"));
    }

    @Test
    void testParseLeftRightPreservesBothDelimiters() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\{x+1\\right]");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("\\left{", fence.getValue());
        assertEquals("{", fence.getMetadata("leftDelimiter"));
        assertEquals("]", fence.getMetadata("rightDelimiter"));
    }

    @Test
    void testParseBoxedCommandAsUnaryEnclosure() {
        LaTeXNode ast = parser.parseLaTeX("\\boxed{x+1}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode boxed = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, boxed.getType());
        assertEquals("\\boxed", boxed.getValue());
        assertEquals(1, boxed.getChildren().size());
        assertEquals("x+1", flatten(boxed.getChildren().get(0)));
    }

    @Test
    void testParseCancelCommandAsUnaryEnclosure() {
        LaTeXNode ast = parser.parseLaTeX("\\xcancel{x+1}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode cancel = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, cancel.getType());
        assertEquals("\\xcancel", cancel.getValue());
        assertEquals(1, cancel.getChildren().size());
        assertEquals("x+1", flatten(cancel.getChildren().get(0)));
    }

    @Test
    void testParseLeftRightNormalizesExtendedFenceDelimiters() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\lfloor x \\right\\rceil");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("⌊", fence.getMetadata("leftDelimiter"));
        assertEquals("⌉", fence.getMetadata("rightDelimiter"));
    }

    @Test
    void testParseLeftRightNormalizesAngleFenceDelimiters() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\langle x+y \\right\\rangle");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("⟨", fence.getMetadata("leftDelimiter"));
        assertEquals("⟩", fence.getMetadata("rightDelimiter"));
        assertEquals("x+y", flatten(fence.getChildren().get(0)));
    }

    @Test
    void testParseLeftRightNormalizesOpenBracketFenceDelimiters() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\llbracket x+y \\right\\rrbracket");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("⟦", fence.getMetadata("leftDelimiter"));
        assertEquals("⟧", fence.getMetadata("rightDelimiter"));
        assertEquals("x+y", flatten(fence.getChildren().get(0)));
    }

    @Test
    void testParseSquareBracketsAsLiteralFormulaCharacters() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{1}{2}[\\frac{1}{a}-\\frac{1}{b}]");
        assertNotNull(ast);

        assertEquals(6, ast.getChildren().size());
        assertEquals(LaTeXNode.Type.FRACTION, ast.getChildren().get(0).getType());
        assertEquals("[", ast.getChildren().get(1).getValue());
        assertEquals(LaTeXNode.Type.FRACTION, ast.getChildren().get(2).getType());
        assertEquals("-", ast.getChildren().get(3).getValue());
        assertEquals(LaTeXNode.Type.FRACTION, ast.getChildren().get(4).getType());
        assertEquals("]", ast.getChildren().get(5).getValue());
    }

    @Test
    void testParseVmatrixEnvironmentWrapsArrayWithDoubleBarFence() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{Vmatrix}1&2\\\\3&4\\end{Vmatrix}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode fence = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, fence.getType());
        assertEquals("\\left\\lVert", fence.getValue());
        assertEquals("||", fence.getMetadata("leftDelimiter"));
        assertEquals("||", fence.getMetadata("rightDelimiter"));
        assertEquals("Vmatrix", fence.getChildren().get(0).getMetadata("environment"));
    }

    @Test
    void testParseAlignedEnvironmentPromotesToImplicitArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{aligned}a&=b\\\\c&=d\\end{aligned}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("aligned", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rl", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseAlignedEnvironmentAlternatesColumnSpecAcrossMultiplePairs() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{aligned}a&=b&c&=d\\end{aligned}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("aligned", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rlrl", array.getMetadata("columnSpec"));
        assertEquals("4", array.getMetadata("columnCount"));
    }

    @Test
    void testParseSplitEnvironmentPromotesToRelationPairArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{split}a&=b\\\\c&=d\\end{split}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("split", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rl", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseAlignStarEnvironmentPromotesToRelationPairArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{align*}a&=b\\\\c&=d\\end{align*}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("align*", array.getMetadata("environment"));
        assertEquals("relation-pairs", array.getMetadata("alignmentMode"));
        assertEquals("rl", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseCasesEnvironmentPromotesToImplicitArray() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{cases}x&1\\\\y&2\\end{cases}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("cases", array.getMetadata("environment"));
        assertEquals("ll", array.getMetadata("columnSpec"));
        assertEquals(2, array.getChildren().size());
    }

    @Test
    void testParseMultiplicationArrayMarksExplicitEmptyCells() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {}\\end{array}"
        );
        assertNotNull(ast);

        LaTeXNode array = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.ARRAY, array.getType());
        assertEquals("true", array.getChildren().get(0).getChildren().get(0).getMetadata("explicitEmptyCell"));
        assertEquals("true", array.getChildren().get(1).getChildren().get(1).getMetadata("explicitEmptyCell"));
    }

    @Test
    void parseTextDoesNotTreatMathArrayAsPlainTextTable() {
        List<ContentSegment> segments = parser.parseText(
            "$\\begin{array}{cc} {} & \\frac { 1 } { 4 } \\end{array}$");

        assertEquals(1, segments.size());
        LaTeXNode array = segments.get(0).ast().getChildren().get(0);
        assertEquals(2, array.getChildren().get(0).getChildren().size());
        assertEquals("true",
            array.getChildren().get(0).getChildren().get(0).getMetadata("explicitEmptyCell"));
    }

    private String flatten(LaTeXNode node) {
        if (node == null) {
            return "";
        }
        if (node.getType() == LaTeXNode.Type.CHAR || node.getType() == LaTeXNode.Type.COMMAND) {
            return node.getValue() == null ? "" : node.getValue();
        }
        StringBuilder builder = new StringBuilder();
        for (LaTeXNode child : node.getChildren()) {
            builder.append(flatten(child));
        }
        return builder.toString();
    }

    @Test
    void testParseOverarcAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\overarc{AB}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode overarc = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, overarc.getType());
        assertEquals("\\overarc", overarc.getValue());
        assertEquals(1, overarc.getChildren().size());
        assertEquals("AB", flatten(overarc.getChildren().get(0)));
    }

    @Test
    void testParseArcAliasAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\arc{AB}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode arc = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, arc.getType());
        assertEquals("\\arc", arc.getValue());
        assertEquals(1, arc.getChildren().size());
        assertEquals("AB", flatten(arc.getChildren().get(0)));
    }

    @Test
    void testParseOverparenAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\overparen{AB}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode overparen = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, overparen.getType());
        assertEquals("\\overparen", overparen.getValue());
        assertEquals(1, overparen.getChildren().size());
        assertEquals("AB", flatten(overparen.getChildren().get(0)));
    }

    @Test
    void testParseWideparenAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\wideparen{AB}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode wideparen = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, wideparen.getType());
        assertEquals("\\wideparen", wideparen.getValue());
        assertEquals(1, wideparen.getChildren().size());
        assertEquals("AB", flatten(wideparen.getChildren().get(0)));
    }

    @Test
    void testParseBraAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\bra{\\psi}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode bra = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, bra.getType());
        assertEquals("\\bra", bra.getValue());
        assertEquals(1, bra.getChildren().size());
        assertEquals("\\psi", bra.getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testParseKetAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\ket{\\psi}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode ket = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, ket.getType());
        assertEquals("\\ket", ket.getValue());
        assertEquals(1, ket.getChildren().size());
        assertEquals("\\psi", ket.getChildren().get(0).getChildren().get(0).getValue());
    }

    @Test
    void testParseBraketAsBinaryDiracCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\braket{\\psi|\\phi}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode braket = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, braket.getType());
        assertEquals("\\braket", braket.getValue());
        assertEquals("|", braket.getMetadata("middleDelimiter"));
        assertEquals(2, braket.getChildren().size());
        assertEquals("\\psi", braket.getChildren().get(0).getChildren().get(0).getValue());
        assertEquals("\\phi", braket.getChildren().get(1).getChildren().get(0).getValue());
    }

    @Test
    void testParseOverbraceAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\overbrace{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode overbrace = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, overbrace.getType());
        assertEquals("\\overbrace", overbrace.getValue());
        assertEquals(1, overbrace.getChildren().size());
        assertEquals("a+b+c", flatten(overbrace.getChildren().get(0)));
    }

    @Test
    void testParseUnderbraceAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\underbrace{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode underbrace = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, underbrace.getType());
        assertEquals("\\underbrace", underbrace.getValue());
        assertEquals(1, underbrace.getChildren().size());
        assertEquals("a+b+c", flatten(underbrace.getChildren().get(0)));
    }

    @Test
    void testPreNormalizeVisualUnderbraceCounterAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("a=1515\\cdots 151004个15︸\\times 333\\cdots 32008个3︸");

        LaTeXNode underbrace = ast.getChildren().get(2);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, underbrace.getType());
        assertEquals("\\underbrace", underbrace.getChildren().get(0).getValue());
        assertEquals("1515\\cdots15", flatten(underbrace.getChildren().get(0).getChildren().get(0)));
        assertEquals("1004个15", flatten(underbrace.getChildren().get(1)));
    }

    @Test
    void testPreNormalizeJoinedVisualUnderbraceCounterAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("=505050\\cdots 51004个5和1003个0︸\\times 999\\cdots 92008个9︸");

        LaTeXNode underbrace = ast.getChildren().get(1);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, underbrace.getType());
        assertEquals("\\underbrace", underbrace.getChildren().get(0).getValue());
        assertEquals("505050\\cdots5", flatten(underbrace.getChildren().get(0).getChildren().get(0)));
        assertEquals("1004个5和1003个0", flatten(underbrace.getChildren().get(1)));
    }

    @Test
    void testPreNormalizeCdotVisualUnderbraceCounterAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("88\\cdot \\cdot \\cdot 82007个8︸\\times 33\\cdot \\cdot \\cdot 32007个3︸");

        LaTeXNode underbrace = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, underbrace.getType());
        assertEquals("\\underbrace", underbrace.getChildren().get(0).getValue());
        assertEquals("88\\cdot\\cdot\\cdot8", flatten(underbrace.getChildren().get(0).getChildren().get(0)));
        assertEquals("2007个8", flatten(underbrace.getChildren().get(1)));
    }

    @Test
    void testPreNormalizeVariableVisualUnderbraceCounterAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("999\\cdots 9k个9︸=1000\\cdots 0k个0︸-1");

        LaTeXNode underbrace = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.SUBSCRIPT, underbrace.getType());
        assertEquals("\\underbrace", underbrace.getChildren().get(0).getValue());
        assertEquals("999\\cdots9", flatten(underbrace.getChildren().get(0).getChildren().get(0)));
        assertEquals("k个9", flatten(underbrace.getChildren().get(1)));
    }

    @Test
    void testParseOverbracketAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\overbracket{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode overbracket = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, overbracket.getType());
        assertEquals("\\overbracket", overbracket.getValue());
        assertEquals(1, overbracket.getChildren().size());
        assertEquals("a+b+c", flatten(overbracket.getChildren().get(0)));
    }

    @Test
    void testParseUnderbracketAsUnaryCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\underbracket{a+b+c}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode underbracket = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, underbracket.getType());
        assertEquals("\\underbracket", underbracket.getValue());
        assertEquals(1, underbracket.getChildren().size());
        assertEquals("a+b+c", flatten(underbracket.getChildren().get(0)));
    }

    @Test
    void testParseXrightarrowAsDedicatedArrowCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\xrightarrow{n\\to\\infty}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode arrow = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, arrow.getType());
        assertEquals("\\xrightarrow", arrow.getValue());
        assertEquals("TM_ARROW", arrow.getMetadata("templateFamily"));
        assertEquals("right", arrow.getMetadata("arrowDirection"));
        assertEquals("top", arrow.getMetadata("annotationPlacement"));
        assertEquals(1, arrow.getChildren().size());
        assertEquals("n\\to\\infty", flatten(arrow.getChildren().get(0)));
    }

    @Test
    void testParseXleftarrowAsDedicatedArrowCommand() {
        LaTeXNode ast = parser.parseLaTeX("\\xleftarrow{f}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode arrow = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, arrow.getType());
        assertEquals("\\xleftarrow", arrow.getValue());
        assertEquals("TM_ARROW", arrow.getMetadata("templateFamily"));
        assertEquals("left", arrow.getMetadata("arrowDirection"));
        assertEquals(1, arrow.getChildren().size());
        assertEquals("f", flatten(arrow.getChildren().get(0)));
    }

    @Test
    void testParseXrightarrowWithBottomAnnotation() {
        LaTeXNode ast = parser.parseLaTeX("\\xrightarrow[T]{n\\to\\infty}");
        assertNotNull(ast);
        assertEquals(1, ast.getChildren().size());

        LaTeXNode arrow = ast.getChildren().get(0);
        assertEquals(LaTeXNode.Type.COMMAND, arrow.getType());
        assertEquals("\\xrightarrow", arrow.getValue());
        assertEquals("TM_ARROW", arrow.getMetadata("templateFamily"));
        assertEquals("right", arrow.getMetadata("arrowDirection"));
        assertEquals("top-bottom", arrow.getMetadata("annotationPlacement"));
        assertEquals(2, arrow.getChildren().size());
        assertEquals("n\\to\\infty", flatten(arrow.getChildren().get(0)));
        assertEquals("T", flatten(arrow.getChildren().get(1)));
    }
}
