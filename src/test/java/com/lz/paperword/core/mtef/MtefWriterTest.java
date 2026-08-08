package com.lz.paperword.core.mtef;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.latex.LaTeXParser.FormulaMetrics;
import com.lz.paperword.core.latex.LaTeXParser.FormulaStyleHints;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class MtefWriterTest {

    private final LaTeXParser parser = new LaTeXParser();
    private final MtefWriter writer = new MtefWriter();

    @Test
    void writeWithReportReturnsStableCanonicalFormulaRecords() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{\\alpha}{2}");

        MtefWriter.WriteReport report = writer.writeWithReport(ast);

        assertTrue(report.bytes().length > 12);
        assertNotNull(report.mathIR());
        assertEquals(1, report.normalization().recordCounts().get("TMPL"));
        assertTrue(report.normalization().canonicalSignature().contains("TMPL:11:0"));
        assertTrue(report.normalization().canonicalSignature().contains("CHAR:"));
        assertTrue(report.normalization().contentOffset() > 0);
    }

    @Test
    void writeReportDefensivelyCopiesMtefBytes() {
        MtefWriter.WriteReport report = writer.writeWithReport(parser.parseLaTeX("x"));
        byte original = report.bytes()[0];

        byte[] mutable = report.bytes();
        mutable[0] = 0;

        assertEquals(original, report.bytes()[0]);
    }

    @Test
    void standaloneAlignmentMarkerDoesNotEnterMtefRecords() {
        String alignedRow = writer.writeWithReport(
            parser.parseLaTeX("&=20.08\\times (200.9-200.7)"))
            .normalization().canonicalSignature();
        String visibleFormula = writer.writeWithReport(
            parser.parseLaTeX("=20.08\\times (200.9-200.7)"))
            .normalization().canonicalSignature();

        assertEquals(visibleFormula, alignedRow);
    }

    @Test
    void raiseboxWritesNudgedMtefLine() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\raisebox{-3pt}{2}"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.LINE,
            (byte) MtefRecord.OPT_NUDGE,
            (byte) 0x80,
            (byte) 0xE0
        }), "-3pt raisebox should lower its MTEF line by 96 units");
    }

    @Test
    void textColorMaroonWritesRgbDefinitionAndRestoresBlack() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\textcolor{maroon}{\\div 16=}"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.COLOR_DEF, 0x00,
            (byte) 0xF6, 0x01, 0x00, 0x00, 0x00, 0x00,
            (byte) MtefRecord.COLOR, 0x02
        }), "maroon should be emitted as #800000-compatible RGB data");
        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.COLOR, 0x01
        }), "the scoped text color must restore the default black color");
    }

    @Test
    void testWriteSimpleChar() {
        LaTeXNode ast = parser.parseLaTeX("x");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 10); // header (12 bytes) + records
        // Check MTEF v5 header
        assertEquals(5, mtef[0] & 0xFF);  // version 5
        assertEquals(1, mtef[1] & 0xFF);  // platform Windows
        assertEquals(0, mtef[2] & 0xFF);  // product MathType (0)
        // product version / subversion may follow template (e.g. 7.0) or fallback constants.
        assertTrue((mtef[3] & 0xFF) > 0);
        assertTrue((mtef[4] & 0xFF) >= 0);
        // Application key "DSMTx\0"
        assertEquals('D', mtef[5] & 0xFF);
        assertEquals('S', mtef[6] & 0xFF);
        assertEquals('M', mtef[7] & 0xFF);
        assertEquals('T', mtef[8] & 0xFF);
        assertTrue((mtef[9] & 0xFF) >= '0' && (mtef[9] & 0xFF) <= '9');
        assertEquals(0,   mtef[10] & 0xFF); // null terminator
        // equation options may be 0/1 depending on template prefix source.
        int eqOptions = mtef[11] & 0xFF;
        assertTrue(eqOptions == 0 || eqOptions == 1);
    }

    @Test
    void testTimesDefaultsToReferenceSymbolEncoding() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{1}{1\\times2}");
        byte[] mtef = writer.write(ast);

        assertEquals(5, mtef[0] & 0xFF);
        assertEquals(1, mtef[1] & 0xFF);
        assertEquals(0, mtef[2] & 0xFF);
        assertEquals(7, mtef[3] & 0xFF);
        assertEquals(11, mtef[4] & 0xFF);
        assertTrue(containsAscii(mtef, "DSMT7\u0000"),
            "generated MTEF should use the pinned MathType 7.11 application key");
        assertFalse(containsAscii(mtef, "DSMT6\u0000"),
            "all generated equations must use the MathType 7 container identity");
        assertTrue(containsBytes(mtef, new byte[] {
                (byte) MtefRecord.CHAR,
                (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
                (byte) 0x86,
                (byte) 0xD7,
                0x00,
                (byte) 0xB4
            }),
            "\\times should default to the reference MathType Symbol encoding");
        assertFalse(containsBytes(mtef, new byte[] {
                (byte) MtefRecord.CHAR,
                0x00,
                (byte) 0x81,
                (byte) 0xD7,
                0x00
            }),
            "\\times should not fall back to the visually mismatched text multiplication sign");
    }

    @Test
    void testTimesSymbolEncodingCandidatesCanBeSelected() {
        String oldEncoding = System.getProperty("latextomathtype.mtef.times.encoding");
        try {
            System.setProperty("latextomathtype.mtef.times.encoding", "no-mtcode");
            LaTeXNode ast = parser.parseLaTeX("\\frac{1}{1\\times2}");
            byte[] mtef = new MtefWriter().write(ast);

            assertTrue(containsBytes(mtef, new byte[] {
                    (byte) MtefRecord.CHAR,
                    (byte) (MtefRecord.OPT_CHAR_ENC_NO_MTCODE | MtefRecord.OPT_CHAR_ENC_CHAR_8),
                    (byte) 0x86,
                    (byte) 0xB4
                }),
                "\\times should expose a Symbol bits8-only candidate for MathType UTF-8 locale testing");
        } finally {
            if (oldEncoding == null) {
                System.clearProperty("latextomathtype.mtef.times.encoding");
            } else {
                System.setProperty("latextomathtype.mtef.times.encoding", oldEncoding);
            }
        }
    }

    @Test
    void simpleExplicitParenthesesUseMathTypeFenceTemplate() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\left(x\\right)"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN,
            (byte) (MtefRecord.TV_FENCE_L | MtefRecord.TV_FENCE_R), 0x00
        }));
    }

    @Test
    void italicStyleDoesNotTurnNumbersIntoVariableGlyphs() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\mathit 0"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.CHAR, 0x00, (byte) (MtefRecord.FN_NUMBER | 0x80), '0', 0x00
        }));
    }

    @Test
    void mathAlphabetVariantsUsePinnedMathTypeGlyphEncodings() {
        byte[] blackboard = writer.write(parser.parseLaTeX("\\mathbb{N}\\mathbb{F}"));
        byte[] fraktur = writer.write(parser.parseLaTeX("\\mathfrak{x}"));
        byte[] sans = writer.write(parser.parseLaTeX("\\mathsf{x}"));

        assertEquals(7, blackboard[3] & 0xFF, "extended MathType typefaces require a DSMT7 header");
        assertEquals(11, blackboard[4] & 0xFF, "DSMT7 header must match pinned MathType 7.11");
        assertTrue(containsBytes(blackboard,
            ("TeX Input Language\0\\mathbb{N}\\mathbb{F}\0").getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
        assertEquals(7, fraktur[3] & 0xFF, "Euclid Math Two requires a DSMT7 header");
        assertEquals(7, sans[3] & 0xFF, "dynamic Euclid typefaces require a DSMT7 header");
        assertEquals(7, writer.write(parser.parseLaTeX("x"))[3] & 0xFF,
            "all formulas target the pinned MathType 7 container contract");
        assertTrue(containsBytes(blackboard, new byte[] {
            (byte) MtefRecord.CHAR, 0x04, (byte) 0x8B, 0x15, 0x21, (byte) 0xA5
        }));
        assertTrue(containsBytes(blackboard, new byte[] {
            (byte) MtefRecord.CHAR, 0x04, 0x7F, (byte) 0x85, (byte) 0xF0, 0x46
        }));
        assertTrue(containsBytes(fraktur, new byte[] {
            (byte) MtefRecord.CHAR, 0x04, 0x7F, 0x31, (byte) 0xF0, 0x78
        }));
        assertTrue(containsBytes(fraktur,
            ("EuclidFraktur\0" + (char) MtefRecord.FONT_DEF + "\u0007Euclid Fraktur\0")
                .getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
        assertTrue(containsBytes(sans, new byte[] {
            (byte) MtefRecord.CHAR, 0x00, 0x7F, 0x78, 0x00
        }));
        assertTrue(containsBytes(sans, new byte[] {
            (byte) MtefRecord.FONT_DEF, 0x05, 'A', 'r', 'i', 'a', 'l', 0,
            (byte) MtefRecord.FONT_STYLE_DEF, (byte) MtefRecord.FN_SYMBOL, 0
        }));
    }

    @Test
    void dynamicTypefaceSymbolPromotesTheWholeStreamToPinnedMathType7() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\emptyset, \\jmath, \\surd"));

        assertEquals(7, mtef[3] & 0xFF);
        assertEquals(11, mtef[4] & 0xFF);
        assertTrue(containsBytes(mtef,
            ("TeX Input Language\0\\emptyset, \\jmath, \\surd\0")
                .getBytes(java.nio.charset.StandardCharsets.US_ASCII)));
        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.CHAR, 0x04, 0x7F, 0x02, (byte) 0xED, (byte) 0xF8
        }));
    }

    @Test
    void sourceMeasuredRoundTripDoesNotAskMathTypeToReparseTex() {
        LaTeXNode ast = parser.parseLaTeX("a=1 \\\\ b=2");
        FormulaStyleHints hints = FormulaStyleHints.empty()
            .withSourceMetrics(new FormulaMetrics(40.0d, 28.0d));

        assertFalse(containsAscii(writer.write(ast, hints), "TeX Input Language"));
    }

    @Test
    void binomialUsesPileAndBoldEmbellishmentKeepsVectorTypeface() {
        byte[] binomial = writer.write(parser.parseLaTeX("\\binom 1 2"));
        byte[] boldHat = writer.write(parser.parseLaTeX("\\mathbf{\\hat u}"));

        Map<String, Integer> binomialRecords = MtefRecordNormalizer.normalize(binomial).recordCounts();
        assertEquals(1, binomialRecords.getOrDefault("PILE", 0));
        assertEquals(0, binomialRecords.getOrDefault("MATRIX", 0));
        assertTrue(containsBytes(boldHat, new byte[] {
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_EMBELL,
            (byte) (MtefRecord.FN_VECTOR | 0x80), 'u', 0x00
        }));
    }

    @Test
    void testWriteFraction() {
        LaTeXNode ast = parser.parseLaTeX("\\frac{1}{2}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 10);
        // Verify header
        assertEquals(5, mtef[0] & 0xFF);
    }

    @Test
    void testWriteSuperscript() {
        LaTeXNode ast = parser.parseLaTeX("x^{2}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 10);
    }

    @Test
    void testWriteSuperscriptThenTimesDoesNotRequireCommandValue() {
        LaTeXNode ast = parser.parseLaTeX("C ^{ 4 }\\times 6=36");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 10);
        assertEquals(5, mtef[0] & 0xFF);
    }

    @Test
    void testWriteSqrt() {
        LaTeXNode ast = parser.parseLaTeX("\\sqrt{x}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 10);
    }

    @Test
    void testWriteGreekLetter() {
        LaTeXNode ast = parser.parseLaTeX("\\alpha");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 5);
    }

    @Test
    void testWriteComplexFormula() {
        // Quadratic formula numerator
        LaTeXNode ast = parser.parseLaTeX("\\frac{-b \\pm \\sqrt{b^{2}-4ac}}{2a}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 20);
        assertEquals(5, mtef[0] & 0xFF);
    }

    @Test
    void testWriteSum() {
        LaTeXNode ast = parser.parseLaTeX("\\sum_{i=1}^{n}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 10);
    }

    @Test
    void testWriteLimitUsesTmLimTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\lim_{x\\to0}\\frac{\\sin x}{x}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        // 与 MathType 7 实测一致：\lim 用 tmSUMOP(0x16)，variation=0x50（TV_BO_SUM|TV_BO_LOWER）
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUMOP, 0x50, 0x00}),
            "limit should use tmSUMOP with lower-slot and summation-style placement");
        // 算子名 slot 由 SYM 字号记录包裹（真 MathType 结构）
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.SYM}),
            "tmSUMOP operator slot should be wrapped in a SYM size record");
        // 回归：operator slot 必须包含算子名字符（FN_FUNCTION 直立体），
        // 否则 MathType 打开时 "lim" 丢失（FN_FUNCTION=2 → typeface 0x82）
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x82, 0x6c}),
            "tmSUMOP operator slot should contain operator name char 'l'");
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x82, 0x69}),
            "tmSUMOP operator slot should contain operator name char 'i'");
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x82, 0x6d}),
            "tmSUMOP operator slot should contain operator name char 'm'");
    }

    @Test
    void testWriteDoubleIntegralUsesIntegralVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\iint_{D} x \\, dy \\, dx");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTEG, 0x12, 0x00}),
            "double integral should use tmINTEG with TV_INT_2 and lower limit bits");
    }

    @Test
    void testWriteIntegralWithOnlySpecialSymbolLowerLimitKeepsNullUpperSlot() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\int_{\\partial D}f(z)dz"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.SUB,
            (byte) MtefRecord.LINE, 0x00,
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
            (byte) (MtefRecord.FN_SYMBOL | 0x80), 0x02, 0x22, (byte) 0xB6,
            (byte) MtefRecord.CHAR, 0x00,
            (byte) (MtefRecord.FN_VARIABLE | 0x80), 0x44, 0x00,
            (byte) MtefRecord.END,
            (byte) MtefRecord.LINE, (byte) MtefRecord.OPT_LINE_NULL,
            (byte) MtefRecord.SYM
        }), "a lower-only integral must preserve the empty upper slot before SYM");
    }

    @Test
    void testWriteContourIntegralUsesLoopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\oint_C f(z) \\, dz");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTEG, 0x14, 0x00}),
            "contour integral should use tmINTEG with loop variation and lower limit bits");
    }

    @Test
    void testWriteCoproductUsesDedicatedTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\coprod_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_COPROD, 0x70, 0x00}),
            "coproduct should use tmCOPROD with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "coproduct should not fall back to tmSUM");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PROD, 0x70, 0x00}),
            "coproduct should not fall back to tmPROD");
    }

    @Test
    void testWriteBigcupUsesDedicatedUnionTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\bigcup_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UNION, 0x70, 0x00}),
            "bigcup should use tmUNION with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "bigcup should not fall back to tmSUM");
    }

    @Test
    void testWriteBigcapUsesDedicatedIntersectionTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\bigcap_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTER, 0x70, 0x00}),
            "bigcap should use tmINTER with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "bigcap should not fall back to tmSUM");
    }

    @Test
    void testWriteBigoplusUsesGenericSummationStyleTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\bigoplus_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUMOP, 0x70, 0x00}),
            "bigoplus should use tmSUMOP with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "bigoplus should not fall back to tmSUM");
    }

    @Test
    void testWriteIntopUsesGenericIntegralStyleTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\intop_{a}^{b} f(x)");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTOP, 0x30, 0x00}),
            "intop should use tmINTOP with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTEG, 0x31, 0x00}),
            "intop should not route through tmINTEG");
    }

    @Test
    void testWriteJointStatusUsesDedicatedTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\jstatus{AB}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_JSTATUS, 0x00, 0x00}),
            "jstatus should use tmJSTATUS");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HAT, 0x00, 0x00}),
            "jstatus should not fall back to tmHAT");
    }

    @Test
    void testWriteBoxedUsesTmBoxTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\boxed{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00}),
            "boxed should use tmBOX with all four square sides enabled");
    }

    @Test
    void testWriteCancelUsesTmStrikeUpVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\cancel{x}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_STRIKE, 0x02, 0x00}),
            "cancel should use tmSTRIKE with the up-diagonal slash variation");
    }

    @Test
    void testWriteXcancelUsesTmStrikeCrossVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\xcancel{x}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_STRIKE, 0x06, 0x00}),
            "xcancel should use tmSTRIKE with both diagonal variations enabled");
        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.COLOR, 0x01,
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_STRIKE, 0x06, 0x00,
            (byte) MtefRecord.COLOR, 0x00,
            (byte) MtefRecord.LINE, 0x00
        }), "MathType strike templates require black template state and automatic-color content");
    }

    @Test
    void testWriteOverarcUsesTmArcTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\overarc{AB}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "overarc should use tmARC instead of falling back to a generic accent path");
    }

    @Test
    void testWriteArcAliasUsesTmArcTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\arc{AB}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "arc alias should also route to tmARC");
    }

    @Test
    void testWriteDotTerminatesEmbellishmentList() {
        LaTeXNode ast = parser.parseLaTeX("0.1\\dot{6}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_EMBELL,
            (byte) (MtefRecord.FN_NUMBER - 128), 0x36, 0x00,
            (byte) MtefRecord.EMBELL, 0x00, (byte) MtefRecord.EMB_1DOT,
            (byte) MtefRecord.END
        }), "dot embellishment should be terminated so docx2tex can read it as \\dot{6}");
    }

    @Test
    void testWriteSingleCharacterBarAndHatAsMathTypeEmbellishments() {
        LaTeXNode ast = parser.parseLaTeX("\\bar x + \\hat y");
        byte[] mtef = writer.write(ast);

        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_EMBELL,
            (byte) (MtefRecord.FN_VARIABLE - 128), 0x78, 0x00,
            (byte) MtefRecord.EMBELL, 0x00, (byte) MtefRecord.EMB_OBAR,
            (byte) MtefRecord.END
        }));
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_EMBELL,
            (byte) (MtefRecord.FN_VARIABLE - 128), 0x79, 0x00,
            (byte) MtefRecord.EMBELL, 0x00, (byte) MtefRecord.EMB_HAT,
            (byte) MtefRecord.END
        }));
        assertFalse(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HAT, 0x00, 0x00
        }));
    }

    @Test
    void testNolimitsArrayUsesMathTypeSideLimitSumTemplate() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\sum\\nolimits_{\\begin{array}{c}a\\\\[0.1em]b\\\\[0.1em]c\\end{array}}");
        byte[] mtef = writer.write(ast);

        assertFalse(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x50, 0x00
        }));
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x10, 0x00
        }));
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX));
    }

    @Test
    void testLabeledBidirectionalArrowUsesMathTypeCompatibleLimitTemplate() {
        LaTeXNode ast = parser.parseLaTeX("A\\xleftrightarrow{\\cong}B");
        byte[] mtef = writer.write(ast);

        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00
        }));
        assertFalse(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x24, 0x00
        }));
    }

    @Test
    void testLongRightArrowWithScriptedLabelUsesMathTypeCompatibleLimitTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\mathbb{Q}\\xlongrightarrow{\\operatorname{ord}_p}\\mathbb{Z}");
        byte[] mtef = writer.write(ast);

        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00
        }));
        assertFalse(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x24, 0x00
        }));
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUB, 0x00, 0x00
        }), "the operator-name annotation must retain its p subscript");
    }

    @Test
    void testDoubleRightArrowAndLabeledEqualUseMathTypeCompatibleLimitTemplates() {
        byte[] doubleArrow = writer.write(parser.parseLaTeX("A\\xLongrightarrow{\\text{implies}}B"));
        byte[] labeledEqual = writer.write(parser.parseLaTeX(
            "A\\xlongequal[\\text{subscript}]{\\text{superscript}}B"));

        assertTrue(containsBytes(doubleArrow, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00
        }));
        assertTrue(containsBytes(labeledEqual, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x30, 0x00
        }));
        assertFalse(containsBytes(doubleArrow, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW
        }));
        assertFalse(containsBytes(labeledEqual, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW
        }));
    }

    @Test
    void testUnlabeledExpandableArrowDoesNotWritePhantomLimitSlots() {
        byte[] mtef = writer.write(parser.parseLaTeX("B\\xLeftrightarrow{}A"));

        assertFalse(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM
        }));
        MtefCharMap.CharEntry arrow = MtefCharMap.lookup("\\Leftrightarrow");
        assertNotNull(arrow);
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
            (byte) (arrow.typeface() | 0x80),
            (byte) (arrow.mtcode() & 0xFF), (byte) (arrow.mtcode() >> 8)
        }));
    }

    @Test
    void testUndersetPreservesCenteredLowerLimitTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\underset{!}{=}"));

        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x10, 0x00,
            (byte) MtefRecord.LINE, 0x00
        }));
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.CHAR, 0x00, (byte) (MtefRecord.FN_FUNCTION | 0x80), 0x21, 0x00
        }));
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.COLOR, 0x00,
            (byte) MtefRecord.LINE, (byte) MtefRecord.OPT_LINE_NULL,
            (byte) MtefRecord.END
        }), "underset must restore color before the empty upper slot");
    }

    @Test
    void testWriteOverparenUsesTmArcTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\overparen{AB}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "overparen should reuse tmARC");
    }

    @Test
    void testWriteWideparenUsesTmArcTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\wideparen{AB}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "wideparen should reuse tmARC");
    }

    @Test
    void testWriteBraUsesTmDiracWithLeftVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\bra{\\psi}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DIRAC, 0x01, 0x00}),
            "bra should use tmDIRAC with left slice variation");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE8, 0x27}),
            "bra should write U+27E8 left angle bracket with FN_EXPAND font");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x7C, 0x00}),
            "bra should write the Dirac vertical bar with FN_EXPAND font");
    }

    @Test
    void testWriteKetUsesTmDiracWithRightVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\ket{\\psi}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DIRAC, 0x02, 0x00}),
            "ket should use tmDIRAC with right slice variation");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x7C, 0x00}),
            "ket should write the Dirac vertical bar with FN_EXPAND font");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE9, 0x27}),
            "ket should write U+27E9 right angle bracket with FN_EXPAND font");
    }

    @Test
    void testWriteBraketUsesTmDiracWithDualVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\braket{a|b}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DIRAC, 0x03, 0x00}),
            "braket should use tmDIRAC with both left and right slice bits");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE8, 0x27}),
            "braket should write U+27E8 left angle bracket with FN_EXPAND font");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x7C, 0x00}),
            "braket should write the Dirac separator bar with FN_EXPAND font");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE9, 0x27}),
            "braket should write U+27E9 right angle bracket with FN_EXPAND font");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x61, 0x00}),
            "braket should serialize the left slot content explicitly");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x62, 0x00}),
            "braket should serialize the right slot content explicitly");
    }

    @Test
    void testWriteXrightarrowUsesMathTypeCompatibleLimitTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\xrightarrow{n\\to\\infty}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00}),
            "xrightarrow should use the MathType-format-compatible upper-slot form");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x24, 0x00}));
    }

    @Test
    void testWriteXleftarrowUsesMathTypeCompatibleLimitTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\xleftarrow{f}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00}),
            "xleftarrow should use the MathType-format-compatible upper-slot form");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x14, 0x00}));
    }

    @Test
    void testWriteXrightarrowWithBottomAnnotationUsesLimitTopAndBottomSlots() {
        LaTeXNode ast = parser.parseLaTeX("\\xrightarrow[b]{a}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x30, 0x00}),
            "xrightarrow[below]{above} should set both limit-slot variation bits");
        byte[] topChar = new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x61, 0x00};
        byte[] bottomChar = new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x62, 0x00};
        byte[] arrowChar = new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0x92, 0x21};
        assertTrue(containsBytes(mtef, topChar), "top annotation should still be serialized");
        assertTrue(containsBytes(mtef, bottomChar), "bottom annotation should be serialized");
        assertTrue(indexOfBytes(mtef, bottomChar) < indexOfBytes(mtef, topChar),
            "MathType limit slots must preserve bottom-before-top ordering");
    }

    @Test
    void testWriteXLongleftarrowUsesMathTypeFormatCompatibleLimitTemplate() {
        LaTeXNode ast = parser.parseLaTeX("B\\xLongleftarrow[\\text{seilpmi}]{}A");

        byte[] mtef = writer.write(ast);

        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x10, 0x00
        }), "xLongleftarrow with a lower label should use MathType's TM_LIM lower-slot form");
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
            (byte) 0x8B, (byte) 0xFD, (byte) 0xFF, 0x6E
        }), "Longleftarrow must use MathType's MT Extra legacy glyph position");
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.LINE, (byte) MtefRecord.OPT_LINE_NULL,
            (byte) MtefRecord.END, (byte) MtefRecord.FULL,
            (byte) MtefRecord.COLOR, 0x01
        }), "the template must close after the null upper label before restoring size and color");
    }

    @Test
    void testWriteOverbraceUsesTmHBRACEWithTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\overbrace{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACE, 0x01, 0x00}),
            "overbrace should use tmHBRACE with TV_HB_TOP variation bit set");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x37, (byte) 0xFE}),
            "overbrace should use MathType's native U+FE37 expandable glyph");
        assertTrue(containsBytes(mtef, new byte[]{
            (byte) MtefRecord.SUB, (byte) MtefRecord.LINE, 0x00,
            (byte) MtefRecord.END, (byte) MtefRecord.FULL
        }), "horizontal fence templates must retain their empty annotation slot");
    }

    @Test
    void testWriteUnderbraceUsesTmHBRACEWithNoTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\underbrace{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACE, 0x00, 0x00}),
            "underbrace should use tmHBRACE with TV_HB_TOP variation bit clear");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x38, (byte) 0xFE}),
            "underbrace should use MathType's native U+FE38 expandable glyph");
    }

    @Test
    void testWriteOverbracketUsesTmHBRACKWithTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\overbracket{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACK, 0x01, 0x00}),
            "overbracket should use tmHBRACK with TV_HB_TOP variation bit set");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x47, (byte) 0xFE}),
            "overbracket should use MathType's native U+FE47 expandable glyph");
    }

    @Test
    void testWriteUnderbracketUsesTmHBRACKWithNoTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\underbracket{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACK, 0x00, 0x00}),
            "underbracket should use tmHBRACK with TV_HB_TOP variation bit clear");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x48, (byte) 0xFE}),
            "underbracket should use MathType's native U+FE48 expandable glyph");
    }

    @Test
    void testCharMapLookup() {
        MtefCharMap.CharEntry alpha = MtefCharMap.lookup("\\alpha");
        assertNotNull(alpha);
        assertEquals(MtefRecord.FN_LC_GREEK, alpha.typeface());
        assertEquals(0x03B1, alpha.mtcode());

        MtefCharMap.CharEntry plus = MtefCharMap.lookupChar('+');
        assertNotNull(plus);
        // '+' is a math operator and should map to Symbol typeface.
        assertEquals(MtefRecord.FN_SYMBOL, plus.typeface());

        MtefCharMap.CharEntry nearrow = MtefCharMap.lookup("\\nearrow");
        assertNotNull(nearrow);
        assertEquals(MtefRecord.FN_MTEXTRA, nearrow.typeface());
        assertEquals(0x2197, nearrow.mtcode());

        MtefCharMap.CharEntry searrow = MtefCharMap.lookup("\\searrow");
        assertNotNull(searrow);
        assertEquals(MtefRecord.FN_MTEXTRA, searrow.typeface());
        assertEquals(0x2198, searrow.mtcode());
    }

    @Test
    void testWriteCrossArrayKeepsDiagonalArrowCommands() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{ccccc}{50\\%} & {} & {} & {} & {10\\%} \\\\ {} & {\\searrow} & {} & {\\nearrow} & {} \\\\ {} & {} & {30\\%} & {} & {} \\\\ {} & {\\nearrow} & {} & {\\searrow} & {} \\\\ {20\\%} & {} & {} & {} & {20\\%}\\end{array}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertFalse(containsAscii(mtef, "Segoe UI Symbol"),
            "cross array should no longer inject a custom Windows arrow font");
        assertFalse(containsAscii(mtef, "DejaVu Sans"),
            "cross array should no longer inject a custom Linux arrow font");
        assertTrue(containsAscii(mtef, "MT Extra"),
            "cross array should use MathType official MT Extra font slot");
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "cross array should still emit matrix records");
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x97, 0x21}), "should serialize \\nearrow as U+2197");
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x98, 0x21}), "should serialize \\searrow as U+2198");
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x8B, (byte) 0x97, 0x21}),
            "diagonal arrows should use FN_MTEXTRA typeface");
        assertFalse(containsBytes(mtef, new byte[]{(byte) 0x86, (byte) 0x97, 0x21}),
            "cross array should not keep FN_SYMBOL diagonal arrow typeface");
    }

    @Test
    void testWriteMatrixEnvironmentAsMatrixRecord() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{matrix}1&2\\\\3&4\\end{matrix}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "matrix environment should emit MATRIX record");
    }

    @Test
    void officialArrayMatrixHeadersMatchDesktopMathTypeAlignmentAndPartitions() {
        byte[] multiColumn = writer.write(parser.parseLaTeX(
            "\\begin{array}{cc|c}1&2&3\\\\x&y&z\\end{array}"));
        byte[] leftColumn = writer.write(parser.parseLaTeX(
            "\\begin{array}{l}a\\\\b\\end{array}"));

        MtefRecordNormalizer.CanonicalRecord multi = MtefRecordNormalizer.normalize(multiColumn).records()
            .stream().filter(record -> record.tag() == MtefRecord.MATRIX).findFirst().orElseThrow();
        MtefRecordNormalizer.CanonicalRecord left = MtefRecordNormalizer.normalize(leftColumn).records()
            .stream().filter(record -> record.tag() == MtefRecord.MATRIX).findFirst().orElseThrow();
        assertEquals(1, multi.matrixVerticalAlignment());
        assertEquals(1, multi.matrixHorizontalAlignment());
        assertEquals("00", multi.rowPartitions());
        assertEquals("00", multi.columnPartitions());
        assertEquals(0, left.matrixHorizontalAlignment());
    }

    @Test
    void testWriteAlignedEnvironmentAsPileWithRelationRuler() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{aligned}a&=b\\\\c&=d\\end{aligned}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "aligned should emit a ruler-backed PILE instead of a generic matrix");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x02, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01}),
            "aligned should expose right/relation tab stops for the two-column pair");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) ((MtefRecord.FN_TEXT & 0x7F) | 0x80), 0x09, 0x00}),
            "aligned pile lines should serialize tab characters between alignment segments");
    }

    @Test
    void testWriteAlignedEnvironmentWithMultipleRelationPairsAsPile() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{aligned}a&=b&c&=d\\end{aligned}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "multi-pair aligned should still emit a ruler-backed PILE");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x04, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01, 0x02, (byte) 0xD0, 0x02, 0x03, (byte) 0xC0, 0x03}),
            "multi-pair aligned should expose alternating right/relation tab stops for every pair");
        assertTrue(countOccurrences(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) ((MtefRecord.FN_TEXT & 0x7F) | 0x80), 0x09, 0x00}) >= 5,
            "multi-pair aligned should serialize one tab per column plus the closing tab");
    }

    @Test
    void testWriteSplitEnvironmentAsPileWithRelationRuler() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{split}a&=b\\\\c&=d\\end{split}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "split should share the aligned pile path instead of falling back to matrix");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x02, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01}),
            "split should expose the same right/relation tab stops as aligned");
    }

    @Test
    void testWriteAlignStarEnvironmentAsPileWithRelationRuler() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{align*}a&=b\\\\c&=d\\end{align*}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "align* should share the aligned pile path for the supported relation-pair subset");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x02, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01}),
            "align* should preserve the aligned relation-pair tab stops");
    }

    @Test
    void testWriteAlignOddColumnShapeFallsBackToMatrixBoundary() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{align}a&=b&c\\end{align}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX),
            "odd-column align stays on the generic matrix boundary until a fuller align model exists");
    }

    @Test
    void testWriteCasesEnvironmentUsesLeftBraceFenceAndMatrixContent() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{cases}x&1\\\\y&2\\end{cases}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BRACE, 0x01, 0x00}),
            "cases should use tmBRACE with only the left fence present");
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "cases content should still emit MATRIX record inside the fence");
    }

    @Test
    void testWriteVmatrixUsesDoubleBarTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{Vmatrix}1&2\\\\3&4\\end{Vmatrix}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DBAR, 0x03, 0x00}),
            "Vmatrix should use tmDBAR with both fences present");
    }

    @Test
    void testWriteFloorFenceUsesFloorTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\lfloor x+1 \\right\\rfloor");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_FLOOR, 0x03, 0x00}),
            "floor fence should use tmFLOOR with both fences present");
    }

    @Test
    void testWriteAngleFenceUsesAngleTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\langle x+1 \\right\\rangle");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ANGLE, 0x03, 0x00}),
            "angle fence should use tmANGLE with both fences present");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE8, 0x27}),
            "angle fence should keep the left angle bracket as FN_EXPAND char");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE9, 0x27}),
            "angle fence should keep the right angle bracket as FN_EXPAND char");
    }

    @Test
    void testWriteOpenBracketFenceUsesOpenBracketTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\left\\llbracket x+1 \\right\\rrbracket");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_OBRACK, 0x03, 0x00}),
            "open bracket fence should use tmOBRACK with both fences present");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE6, 0x27}),
            "open bracket fence should keep the left white bracket as FN_EXPAND char");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE7, 0x27}),
            "open bracket fence should keep the right white bracket as FN_EXPAND char");
    }

    @Test
    void testWriteSingleSidedBarFenceKeepsOnlyPresentSide() {
        LaTeXNode ast = parser.parseLaTeX("\\left. x \\right|");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BAR, 0x02, 0x00}),
            "single-sided right bar fence should encode only the right fence bit");
    }

    @Test
    void testWriteMixedIntervalFenceUsesTmInterval() {
        LaTeXNode ast = parser.parseLaTeX("\\left( x+1 \\right]");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTERVAL, 0x30, 0x00}),
            "mixed ( ] fences should use tmINTERVAL with left-LP and right-RB variation bits");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x28, 0x00}),
            "interval should keep the left parenthesis as an FN_EXPAND char");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x5D, 0x00}),
            "interval should keep the right bracket as an FN_EXPAND char");
    }

    @Test
    void testWriteEquationNumberFencesInsideExpressionUseParenTemplates() {
        LaTeXNode ast = parser.parseLaTeX("\\left ( { 1 } \\right )-\\left ( { 2 } \\right )");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(countOccurrences(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN}) >= 2,
            "single-digit equation number fences inside expressions should keep MathType's tmPAREN templates");
    }

    @Test
    void testExplicitParenFenceWithSourceMetricsUsesTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\left ( { a,b } \\right )");
        FormulaStyleHints hints = FormulaStyleHints.empty().withSourceMetrics(new FormulaMetrics(28.0d, 13.0d));
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN}),
            "xsc reference objects keep explicit \\left...\\right parens as MathType fence templates");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x28, 0x00}),
            "explicit paren template should write an expandable left paren");
    }

    @Test
    void testArithmeticDigitFencesStayFlat() {
        LaTeXNode ast = parser.parseLaTeX("5\\times (1)-(2)");
        byte[] mtef = writer.write(ast);
        int baselineParenTemplates = countOccurrences(writer.write(parser.parseLaTeX("5\\times 1-2")),
            new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN});

        assertNotNull(mtef);
        assertEquals(baselineParenTemplates,
            countOccurrences(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN}),
            "single-digit arithmetic operands should stay as flat parenthesis characters, not equation-number templates");
    }

    @Test
    void testFlatFenceRestoresFullSizeBeforeClosingAfterTrailingScript() {
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, false, false, false, false, null);
        byte[] mtef = writer.write(parser.parseLaTeX("(1^3+2^3)"), hints);
        var records = MtefRecordNormalizer.normalize(mtef).records();
        int close = -1;
        for (int index = 0; index < records.size(); index++) {
            if (records.get(index).tag() == MtefRecord.CHAR
                    && Integer.valueOf((int) ')').equals(records.get(index).mtcode())) {
                close = index;
                break;
            }
        }

        assertTrue(close > 0, "flat closing parenthesis must be present");
        assertEquals(MtefRecord.FULL, records.get(close - 1).tag(),
            "a trailing script must restore the containing line size before its closing delimiter");
    }

    @Test
    void testExplicitEmptyFirstArrayCellProducesAnEmptyMatrixSlot() {
        byte[] mtef = writer.write(parser.parseLaTeX(
            "\\begin{array}{cc} {} & \\frac14 \\end{array}"));
        var records = MtefRecordNormalizer.normalize(mtef).records();
        int matrix = -1;
        for (int index = 0; index < records.size(); index++) {
            if (records.get(index).tag() == MtefRecord.MATRIX) {
                matrix = index;
                break;
            }
        }

        assertTrue(matrix >= 0, "array must emit a MATRIX record");
        assertEquals(MtefRecord.LINE, records.get(matrix + 1).tag());
        assertEquals(MtefRecord.END, records.get(matrix + 2).tag(),
            "the explicit empty first cell must remain an empty LINE slot");
        assertEquals(MtefRecord.LINE, records.get(matrix + 3).tag(),
            "the fraction must remain in the second cell");
    }

    @Test
    void testExplicitEmptyFirstArrayCellSurvivesStructuredSecondCell() {
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, true, false, false, false, false, false, null);
        byte[] mtef = writer.write(parser.parseLaTeX(
            "\\begin{array}{cc} {} & \\frac { 1 } { 4 }\\times \\left( { 4.85\\div \\frac { 5 } { 18 }-3.6+6.15\\times 3\\frac { 3 } { 5 } } \\right)+\\left[ 5.5-1.75\\times \\left( { 1\\frac { 2 } { 3 }+\\frac { 19 } { 21 } } \\right) \\right] \\end{array}"), hints);
        var records = MtefRecordNormalizer.normalize(mtef).records();
        int matrix = -1;
        for (int index = 0; index < records.size(); index++) {
            if (records.get(index).tag() == MtefRecord.MATRIX) {
                matrix = index;
                break;
            }
        }

        assertTrue(matrix >= 0);
        assertEquals(MtefRecord.LINE, records.get(matrix + 1).tag());
        assertEquals(MtefRecord.END, records.get(matrix + 2).tag());
        assertEquals(MtefRecord.LINE, records.get(matrix + 3).tag());
    }

    @Test
    void testSourceFlatParenTemplateHintUsesTmParen() {
        LaTeXNode ast = parser.parseLaTeX("(105-5)");
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, true, false, false, false, null);
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN, 0x00, 0x04}),
            "source objects that use MathType's flat tmPAREN wrapper should preserve it via style hints");
    }

    @Test
    void testSourceFlatParenTemplateHintKeepsDecimalParensFlat() {
        LaTeXNode ast = parser.parseLaTeX("(0.099+0.111)\\div 2=0.105");
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, true, false, false, false, null);
        byte[] mtef = writer.write(ast, hints);
        int baselineParenTemplates = countOccurrences(writer.write(parser.parseLaTeX("0.099+0.111\\div 2=0.105"), hints),
            new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN});

        assertNotNull(mtef);
        assertEquals(baselineParenTemplates,
            countOccurrences(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN}),
            "decimal arithmetic parentheses should stay as flat source characters even when the source file has flat paren template hints");
    }

    @Test
    void testSourceLetterGroupObarHintWrapsAlphabeticRuns() {
        LaTeXNode ast = parser.parseLaTeX("(abc+def)");
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, false, true, false, false, null);
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        assertEquals(2,
            countOccurrences(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_OBAR, 0x00, 0x00}),
            "source objects that wrap alphabetic groups with tmOBAR should preserve those templates");
    }

    @Test
    void testCjkCharactersUseFarEastTextTypeface() {
        LaTeXNode ast = parser.parseLaTeX("S_{和}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) (MtefRecord.FN_TEXT_FE | 0x80), (byte) 0x8C, 0x54}),
            "CJK text inside formulas should use MathType's Far East text typeface instead of variable italics");
    }

    @Test
    void testWriteReversedIntervalFenceUsesTmInterval() {
        LaTeXNode ast = parser.parseLaTeX("\\left] x \\right(");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTERVAL, 0x03, 0x00}),
            "reversed ] ( fences should use tmINTERVAL with left-RB and right-LP variation bits");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x5D, 0x00}),
            "reversed interval should keep the left bracket glyph");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x28, 0x00}),
            "reversed interval should keep the right parenthesis glyph");
    }

    @Test
    void testCrossingFenceUsesTwoOneSidedTemplatesInSourceOrder() {
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, false, false, false, false, null);
        var records = MtefRecordNormalizer.normalize(writer.write(parser.parseLaTeX(
            "\\left[ 13\\times (\\frac{5}{7}+\\frac{5}{9}\\right]"), hints)).records();
        var bracketTemplates = records.stream()
            .filter(record -> record.tag() == MtefRecord.TMPL
                && Integer.valueOf(MtefRecord.TM_BRACK).equals(record.selector()))
            .toList();

        assertEquals(2, bracketTemplates.size());
        assertEquals(1, bracketTemplates.get(0).variation());
        assertEquals(2, bracketTemplates.get(1).variation());
        assertTrue(records.stream().anyMatch(record -> record.tag() == MtefRecord.CHAR
            && Integer.valueOf(MtefRecord.FN_FUNCTION | 0x80).equals(record.typeface())
            && Integer.valueOf((int) ')').equals(record.mtcode())));
        assertTrue(records.stream().noneMatch(record -> record.tag() == MtefRecord.TMPL
            && Integer.valueOf(MtefRecord.TM_BRACK).equals(record.selector())
            && Integer.valueOf(3).equals(record.variation())));
    }

    @Test
    void testWriteMathIrDirectlyForStableCoreSubset() {
        byte[] mtef = writer.write(parser.parseMathIR("\\begin{pmatrix}\\frac{1}{x}&\\sum_{i=1}^{n}a_i\\\\\\sqrt[3]{y}&z_0\\end{pmatrix}"));

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "IR path should still emit MATRIX record for pmatrix");
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "IR path should still emit TMPL records for fraction/root/scripts");
    }

    @Test
    void testWriteMathIrDirectlyForAlignedRelationPairs() {
        byte[] mtef = writer.write(parser.parseMathIR("\\begin{aligned}a&=b\\\\c&=d\\end{aligned}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "IR path should preserve aligned relation-pair semantics");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x02, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01}),
            "IR path should keep the aligned relation tab stops");
    }

    @Test
    void testWriteMathIrDirectlyForAlignedMultipleRelationPairs() {
        byte[] mtef = writer.write(parser.parseMathIR("\\begin{aligned}a&=b&c&=d\\end{aligned}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x04, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01, 0x02, (byte) 0xD0, 0x02, 0x03, (byte) 0xC0, 0x03}),
            "IR path should preserve all aligned relation-pair tab stops");
    }

    @Test
    void testWriteMathIrDirectlyForSplitRelationPairs() {
        byte[] mtef = writer.write(parser.parseMathIR("\\begin{split}a&=b\\\\c&=d\\end{split}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "IR path should route split through the aligned pile semantics");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x02, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01}),
            "IR path should preserve split relation-pair tab stops");
    }

    @Test
    void testWriteMathIrDirectlyForAlignStarRelationPairs() {
        byte[] mtef = writer.write(parser.parseMathIR("\\begin{align*}a&=b\\\\c&=d\\end{align*}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.PILE, MtefRecord.OPT_LP_RULER, 0x01, 0x02}),
            "IR path should route align* through the aligned pile semantics for the supported subset");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.RULER, 0x02, 0x02, (byte) 0xF0, 0x00, 0x03, (byte) 0xE0, 0x01}),
            "IR path should preserve align* relation-pair tab stops");
    }

    @Test
    void testWriteMathIrDirectlyForExtendedFenceSubset() {
        byte[] mtef = writer.write(parser.parseMathIR("\\left\\lVert x+1 \\right.") );

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DBAR, 0x01, 0x00}),
            "IR path should preserve single-sided double-bar fences");
    }

    @Test
    void testWriteMathIrDirectlyForAngleFenceSubset() {
        byte[] mtef = writer.write(parser.parseMathIR("\\left\\langle x+1 \\right\\rangle"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ANGLE, 0x03, 0x00}),
            "IR path should lower angle fences to tmANGLE");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE8, 0x27}),
            "IR path should keep the left angle bracket character");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE9, 0x27}),
            "IR path should keep the right angle bracket character");
    }

    @Test
    void testWriteMathIrDirectlyForOpenBracketFenceSubset() {
        byte[] mtef = writer.write(parser.parseMathIR("\\left\\llbracket x+1 \\right\\rrbracket"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_OBRACK, 0x03, 0x00}),
            "IR path should lower white square brackets to tmOBRACK");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE6, 0x27}),
            "IR path should keep the left white bracket character");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE7, 0x27}),
            "IR path should keep the right white bracket character");
    }

    @Test
    void testWriteMathIrDirectlyForIntervalFenceSubset() {
        byte[] mtef = writer.write(parser.parseMathIR("\\left( x+1 \\right]") );

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTERVAL, 0x30, 0x00}),
            "IR path should lower mixed paren/bracket fences to tmINTERVAL");
    }

    @Test
    void testWriteMathIrDirectlyForBoxAndStrikeEnclosures() {
        byte[] mtef = writer.write(parser.parseMathIR("\\boxed{\\xcancel{x}}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00}),
            "IR path should lower boxed to tmBOX");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_STRIKE, 0x06, 0x00}),
            "IR path should lower xcancel to tmSTRIKE with both diagonal bits");
    }

    @Test
    void testWriteMathIrDirectlyForArcTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\overarc{AB}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "IR path should lower overarc to tmARC");
    }

    @Test
    void testWriteMathIrDirectlyForOverparenArcTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\overparen{AB}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "IR path should lower overparen to tmARC");
    }

    @Test
    void testWriteMathIrDirectlyForWideparenArcTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\wideparen{AB}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARC, 0x00, 0x00}),
            "IR path should lower wideparen to tmARC");
    }

    @Test
    void testWriteMathIrDirectlyForBraDiracTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bra{\\psi}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DIRAC, 0x01, 0x00}),
            "IR path should lower bra to tmDIRAC left slice");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE8, 0x27}),
            "IR path should keep the left angle bracket character");
    }

    @Test
    void testWriteMathIrDirectlyForKetDiracTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\ket{\\psi}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DIRAC, 0x02, 0x00}),
            "IR path should lower ket to tmDIRAC right slice");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xE9, 0x27}),
            "IR path should keep the right angle bracket character");
    }

    @Test
    void testWriteMathIrDirectlyForBraketDiracTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\braket{a|b}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DIRAC, 0x03, 0x00}),
            "IR path should lower braket to tmDIRAC with both slice bits");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x61, 0x00}),
            "IR path should keep the left Dirac slot content");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x62, 0x00}),
            "IR path should keep the right Dirac slot content");
    }

    @Test
    void testWriteMathIrDirectlyForXrightarrowLimitTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\xrightarrow{f}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00}),
            "IR path should lower xrightarrow to the format-compatible limit template");
    }

    @Test
    void testWriteMathIrDirectlyForXleftarrowLimitTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\xleftarrow{g}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x20, 0x00}),
            "IR path should lower xleftarrow to the format-compatible limit template");
    }

    @Test
    void testWriteMathIrDirectlyForXleftarrowWithBottomAnnotation() {
        byte[] mtef = writer.write(parser.parseMathIR("\\xleftarrow[b]{a}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x30, 0x00}),
            "IR path should preserve both limit slots for xleftarrow");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x61, 0x00}),
            "IR path should preserve the top annotation payload");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x62, 0x00}),
            "IR path should preserve the bottom annotation payload");
    }

    @Test
    void testWriteMathIrDirectlyForHorizontalBracketTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\overbracket{x+1}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACK, 0x01, 0x00}),
            "IR path should lower overbracket to tmHBRACK");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, 0x47, (byte) 0xFE}),
            "IR path should keep the top square bracket expandable character");
    }

    @Test
    void testWriteMathIrDirectlyForLimitTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\lim_{n\\to\\infty} a_n"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUMOP, 0x50, 0x00}),
            "IR path should lower \\lim to tmSUMOP (matching real MathType 7)");
    }

    @Test
    void testWriteMathIrDirectlyForCoproductTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\coprod_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_COPROD, 0x70, 0x00}),
            "IR path should lower \\coprod to tmCOPROD");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "IR path should not route \\coprod through tmSUM fallback");
    }

    @Test
    void testWriteMathIrDirectlyForUnionTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bigcup_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UNION, 0x70, 0x00}),
            "IR path should lower \\bigcup to tmUNION");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "IR path should not route \\bigcup through tmSUM fallback");
    }

    @Test
    void testWriteMathIrDirectlyForIntersectionTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bigcap_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTER, 0x70, 0x00}),
            "IR path should lower \\bigcap to tmINTER");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "IR path should not route \\bigcap through tmSUM fallback");
    }

    @Test
    void testWriteMathIrDirectlyForSummationStyleGenericBigOperatorTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bigotimes_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUMOP, 0x70, 0x00}),
            "IR path should lower \\bigotimes to tmSUMOP");
    }

    @Test
    void testWriteMathIrDirectlyForIntegralStyleGenericBigOperatorTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\intop_{a}^{b} f(x)"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTOP, 0x30, 0x00}),
            "IR path should lower \\intop to tmINTOP");
    }

    @Test
    void testWriteMathIrDirectlyForJointStatusTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\jointstatus{AB}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_JSTATUS, 0x00, 0x00}),
            "IR path should lower \\jointstatus to tmJSTATUS");
    }

    @Test
    void testWriteUnsupportedMathIrFailsExplicitly() {
        UnsupportedOperationException ex = assertThrows(UnsupportedOperationException.class,
            () -> writer.write(parser.parseMathIR("\\foo{1}")));

        assertTrue(ex.getMessage().contains("Unsupported MathIR node"));
    }

    @Test
    void testWriteArrayAdditionAsMatrix() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 1 & 2 & 3 \\\\ + & 4 & 5 & 6 \\\\ \\hline & 5 & 7 & 9\\end{array}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(mtef.length > 20);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "array should emit MATRIX record");
    }

    @Test
    void testWriteArrayDivisionAsMatrix() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{r|l}13 & 845 \\\\ \\hline & 65 \\\\ & 78 \\\\ & 65 \\\\ & 0\\end{array}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "division template should emit MATRIX record");
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "long division should emit TMPL record");
    }

    @Test
    void testWriteDecimalArrayKeepsDotCharacter() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcr}12 & . & 50 \\\\ +3 & . & 75 \\\\ \\hline 16 & . & 25\\end{array}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.PILE), "decimal array should emit PILE record");
        assertTrue(containsBytes(mtef, new byte[]{'.', 0x00}), "decimal point should be serialized in MTEF");
    }

    @Test
    void testWriteArraySubtractionAsMatrix() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rrrr} & 8 & 6 & 4 \\\\ - & 2 & 7 & 9 \\\\ \\hline & 5 & 8 & 5\\end{array}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "subtraction array should emit MATRIX record");
    }

    @Test
    void testWriteDecimalSubtractionKeepsDotCharacter() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{array}{rcr}12 & . & 50 \\\\ -3 & . & 75 \\\\ \\hline 8 & . & 75\\end{array}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.PILE), "decimal subtraction should emit PILE record");
        assertTrue(containsBytes(mtef, new byte[]{'.', 0x00}), "decimal point should be preserved in subtraction");
    }

    @Test
    void testWriteMultiplicationArrayKeepsExplicitSpacerCells() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {}\\end{array}"
        );
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "multiplication array should emit MATRIX record");
        assertTrue(countOccurrences(mtef, new byte[]{0x20, 0x00}) >= 2,
            "explicit placeholder cells should be serialized as space-based spacer cells");
    }

    @Test
    void testWriteExplicitLongDivisionAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "explicit long division should emit MATRIX wrapper");
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "long division should emit TMPL record");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV, 0x01, 0x00}),
            "long division should use tmLDIV with upper slot variation");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UBAR, 0x00, 0x00}),
            "没有显式步骤区时不应再本地补出下划线步骤");
    }

    @Test
    void testWriteExplicitLongDivisionSerializesDividendBeforeQuotient() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[129]{12}{1548}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        String digitStream = extractDigitStream(mtef);
        // tmLDIV slot order: divisor (outside template), quotient (slot 0), dividend (slot 1)
        assertTrue(digitStream.contains("12"),
            "tmLDIV should write divisor before template");
        assertTrue(digitStream.contains("129"),
            "tmLDIV slot 0 should contain quotient");
        assertTrue(digitStream.contains("1548"),
            "tmLDIV slot 1 should contain dividend");
    }

    @Test
    void testWriteExplicitLongDivisionWithoutQuotientAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{13}{845}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "long division without quotient should emit MATRIX wrapper");
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "long division without quotient should emit TMPL record");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV, 0x00, 0x00}),
            "没有显式商时应保留头部模板而不是自动补商");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UBAR, 0x00, 0x00}),
            "没有显式步骤区时不应本地补出下划线步骤");
        String digitStream = extractDigitStream(mtef);
        assertFalse(digitStream.contains("65"), "导出端不应再本地补出商 65");
    }

    @Test
    void testWriteThreeStepLongDivisionHeaderDoesNotComputeUnderlinesLocally() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV, 0x01, 0x00}),
            "three-step long division should still use tmLDIV");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UBAR, 0x00, 0x00}),
            "只有头部时不应再本地补出三步下划线结构");
    }

    @Test
    void testWriteCompositeLongDivisionSingleBlockKeepsExplicitSteps() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\longdiv[570]{6}{3420}\\begin{array}{l}\\text{   }\\underline{30}\\\\\\text{    }42\\\\\\text{    }\\underline{42}\\\\\\text{      }0\\end{array}"
        );
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV, 0x01, 0x00}),
            "单块复合长除法仍应使用 tmLDIV 头部");
        assertTrue(countOccurrences(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UBAR, 0x00, 0x00}) >= 2,
            "单块复合长除法应保留原图中的下划线步骤");
        String digitStream = extractDigitStream(mtef);
        assertTrue(digitStream.contains("570"), "商应继续保留");
        assertTrue(digitStream.contains("3420"), "被除数应继续保留");
        assertTrue(digitStream.contains("30"), "显式步骤区中的第一步乘积应被写出");
        assertTrue(digitStream.contains("42"), "显式步骤区中的后续数字应被写出");
    }

    @Test
    void testWriteExplicitLongDivisionWithoutStructuredStepsFallsBackToSimpleHeader() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{x}{845}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV, 0x00, 0x00}),
            "non-numeric long division should still emit tmLDIV header");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UBAR, 0x00, 0x00}),
            "non-numeric long division should not synthesize structured underline steps");
    }

    @Test
    void testWriteDecimalLongDivisionHeaderKeepsRawValues() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[5]{2.5}{12.5}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LDIV, 0x01, 0x00}),
            "带显式商的小数长除法仍应使用带商槽的 tmLDIV");
        String digitStream = extractDigitStream(mtef);
        assertTrue(digitStream.contains("25"), "原始小数头部中的数字字符应被写出");
        assertTrue(digitStream.contains("125"), "原始小数被除数中的数字字符应被写出");
        assertTrue(digitStream.contains("5"), "显式商值应被写出");
    }

    @Test
    void testWriteLeadingSuperscriptUsesPrecedesVariation() {
        LaTeXNode ast = parser.parseLaTeX("{}^{a}x");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] tmpl = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUP, 0x01, 0x00};
        byte[] xChar = new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x78, 0x00};
        assertTrue(containsBytes(mtef, tmpl), "leading superscript should use tmSUP + tvSU_PRECEDES");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LSCRIPT, 0x00, 0x00}),
            "v5 stream should not rely on legacy tmLSCRIPT selector");
        assertTrue(indexOfBytes(mtef, tmpl) < indexOfBytes(mtef, xChar),
            "pre-script template must appear before the scripted item");
    }

    @Test
    void testWriteLeadingSubscriptUsesPrecedesVariation() {
        LaTeXNode ast = parser.parseLaTeX("{}_{i}x");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] tmpl = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUB, 0x01, 0x00};
        byte[] xChar = new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x78, 0x00};
        assertTrue(containsBytes(mtef, tmpl), "leading subscript should use tmSUB + tvSU_PRECEDES");
        assertTrue(indexOfBytes(mtef, tmpl) < indexOfBytes(mtef, xChar),
            "pre-script template must appear before the scripted item");
    }

    @Test
    void testWriteLeadingSubSupUsesPrecedesVariation() {
        LaTeXNode ast = parser.parseLaTeX("{}_{i}^{n}x");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] tmpl = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUBSUP, 0x01, 0x00};
        byte[] xChar = new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x83, 0x78, 0x00};
        assertTrue(containsBytes(mtef, tmpl), "leading sub/sup should use tmSUBSUP + tvSU_PRECEDES");
        assertTrue(indexOfBytes(mtef, tmpl) < indexOfBytes(mtef, xChar),
            "combined pre-script template must appear before the scripted item");
    }

    @Test
    void testWriteLeadingSupSubReversedSourceOrderStillUsesPrecedesVariation() {
        LaTeXNode ast = parser.parseLaTeX("{}^{n}_{i}x");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUBSUP, 0x01, 0x00}),
            "{}^{n}_{i}x should normalize to the same tmSUBSUP + tvSU_PRECEDES encoding");
    }

    @Test
    void testDivisionEquationChainUsesMathTypeBoxSegments() {
        LaTeXNode ast = parser.parseLaTeX("AB\\div C=DE\\div F=GH\\div I=3");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(3, countOccurrences(mtef, box),
            "MathType encodes divisor-chain operands as tmBOX 0x1e segments");
        assertTrue(containsBytes(mtef, new byte[]{0x02, 0x04, (byte) ((MtefRecord.FN_SYMBOL & 0x7F) | 0x80), (byte) 0xF7, 0x00, (byte) 0xB8}),
            "flat division chain should write \\div as the MathType Symbol division glyph");
        assertFalse(containsBytes(mtef, new byte[]{0x02, 0x00, (byte) MtefRecord.FN_VARIABLE, 0x5C, 0x00}),
            "flat division chain must not serialize the LaTeX command backslash as a variable");
    }

    @Test
    void testSingleDivisionUsesMathTypeBoxSegment() {
        LaTeXNode ast = parser.parseLaTeX("12\\div 3");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(1, countOccurrences(mtef, box),
            "MathType also boxes the divisor operand in a single division expression");
    }

    @Test
    void testSingleLetterDivisionStaysFlat() {
        LaTeXNode ast = parser.parseLaTeX("a\\div b");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(0, countOccurrences(mtef, box),
            "single-letter variable division in the xsc corpus matches the flat MathType body better");
    }

    @Test
    void testMultiplicationEquationUsesMathTypeBoxSegments() {
        LaTeXNode ast = parser.parseLaTeX("3\\times 4=12");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(3, countOccurrences(mtef, box),
            "MathType boxes the post-times operand and each result digit in simple multiplication equations");
        assertTrue(containsBytes(mtef, new byte[]{0x02, 0x04, (byte) ((MtefRecord.FN_SYMBOL & 0x7F) | 0x80), (byte) 0xD7, 0x00, (byte) 0xB4}),
            "flat multiplication equation should write \\times as the MathType Symbol multiplication glyph");
        assertFalse(containsBytes(mtef, new byte[]{0x02, 0x00, (byte) MtefRecord.FN_VARIABLE, 0x5C, 0x00}),
            "flat multiplication equation must not serialize the LaTeX command backslash as a variable");
    }

    @Test
    void testShortMultiplicationEquationWithLinearMetricsStaysFlat() {
        LaTeXNode ast = parser.parseLaTeX("7\\times 9=63");
        FormulaStyleHints hints = FormulaStyleHints.empty().withSourceMetrics(new FormulaMetrics(43.0d, 13.0d));
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(0, countOccurrences(mtef, box),
            "13pt xsc inline arithmetic uses a flat MathType character stream, not boxed operand templates");
        assertTrue(containsBytes(mtef, new byte[]{0x02, 0x04, (byte) ((MtefRecord.FN_SYMBOL & 0x7F) | 0x80), (byte) 0xD7, 0x00, (byte) 0xB4}),
            "flat multiplication equation should still write \\times as the MathType Symbol multiplication glyph");
    }

    @Test
    void testFullwidthParenthesesUseFarEastTextRecords() {
        LaTeXNode ast = parser.parseLaTeX("（n> m）");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{0x02, 0x00, (byte) 0x8C, 0x08, (byte) 0xFF}));
        assertTrue(containsBytes(mtef, new byte[]{0x02, 0x00, (byte) 0x8C, 0x09, (byte) 0xFF}));
        assertFalse(containsBytes(mtef, new byte[]{0x02, 0x00, (byte) MtefRecord.FN_VARIABLE, 0x08, (byte) 0xFF}));
    }

    @Test
    void testBoxedZeroWidthPlaceholderWritesEmptySlot() {
        LaTeXNode ast = parser.parseLaTeX("\\boxed{\u200d\u200d\u200d \u200d}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertEquals(1, countOccurrences(mtef, new byte[]{
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00
        }));
        assertFalse(containsBytes(mtef, new byte[]{0x02, 0x00, (byte) MtefRecord.FN_VARIABLE, 0x0D, 0x20}));
    }

    @Test
    void testAsciiFlatParensOverrideFlatParenTemplateHint() {
        LaTeXNode ast = parser.parseLaTeX("=(110+126)\\times 17\\div 2=2006");
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, true, false, false, false, null);
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        String chars = extractCharStream(mtef);
        assertTrue(chars.contains("=(110+126)"),
            "asciiFlatParens must preserve literal paren order even when flatParenTemplate is also present");
        assertFalse(chars.contains("=110+126()"),
            "asciiFlatParens must not move literal parens after the content");
    }

    @Test
    void testAsciiFlatParensFlattenStructuredImplicitFence() {
        LaTeXNode ast = parser.parseLaTeX("(1+\\frac{1}{2})");
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, false, false, false, false, null);
        byte[] mtef = writer.write(ast, hints);

        var records = MtefRecordNormalizer.normalize(mtef).records();
        assertFalse(records.stream().anyMatch(record ->
            record.tag() == MtefRecord.TMPL
                && Integer.valueOf(MtefRecord.TM_PAREN).equals(record.selector())));
        assertTrue(records.stream().anyMatch(record ->
            record.tag() == MtefRecord.CHAR && Integer.valueOf('(').equals(record.mtcode())));
        assertTrue(records.stream().anyMatch(record ->
            record.tag() == MtefRecord.CHAR && Integer.valueOf(')').equals(record.mtcode())));
    }

    @Test
    void testLegacyLiteralAsteriskAndEscapedUnderscoreUseFunctionTypeface() {
        var asterisk = MtefCharMap.lookup("*");
        var underscore = MtefCharMap.lookup("\\_");

        assertEquals(MtefRecord.FN_FUNCTION, asterisk.typeface());
        assertEquals((int) '*', asterisk.mtcode());
        assertEquals(MtefRecord.FN_FUNCTION, underscore.typeface());
        assertEquals((int) '_', underscore.mtcode());
    }

    @Test
    void testFullwidthParenHintFlattensStructuredImplicitFence() {
        LaTeXNode ast = parser.parseLaTeX("(1+\\frac{1}{2})");
        FormulaStyleHints hints = new FormulaStyleHints(
            false, false, false, false, false, false, false, false, false, true, null);
        var records = MtefRecordNormalizer.normalize(writer.write(ast, hints)).records();

        assertFalse(records.stream().anyMatch(record ->
            record.tag() == MtefRecord.TMPL
                && Integer.valueOf(MtefRecord.TM_PAREN).equals(record.selector())));
        assertTrue(records.stream().anyMatch(record -> Integer.valueOf(0xFF08).equals(record.mtcode())));
        assertTrue(records.stream().anyMatch(record -> Integer.valueOf(0xFF09).equals(record.mtcode())));
    }

    @Test
    void testMixedAsciiFullwidthParenHintPreservesPairOrder() {
        LaTeXNode ast = parser.parseLaTeX("(1+\\frac{1}{2})+(3+\\frac{1}{4})");
        FormulaStyleHints hints = new FormulaStyleHints(
            false, false, false, false, false, false, false, false, false, false,
            true, false, null);
        var parens = MtefRecordNormalizer.normalize(writer.write(ast, hints)).records().stream()
            .filter(record -> record.tag() == MtefRecord.CHAR)
            .filter(record -> Integer.valueOf('(').equals(record.mtcode())
                || Integer.valueOf(')').equals(record.mtcode())
                || Integer.valueOf(0xFF08).equals(record.mtcode())
                || Integer.valueOf(0xFF09).equals(record.mtcode()))
            .toList();

        assertEquals(List.of((int) '(', (int) ')', 0xFF08, 0xFF09),
            parens.stream().map(record -> record.mtcode()).toList());
    }

    @Test
    void testLegacyTextFeParenContentScopesAsciiTypeface() {
        LaTeXNode ast = parser.parseLaTeX("(1^3+2^3+3^3+\\cdots+10^3)");
        FormulaStyleHints hints = new FormulaStyleHints(
            false, false, false, false, false, false, false, false, false, true,
            false, true, null);
        var textFeCodes = MtefRecordNormalizer.normalize(writer.write(ast, hints)).records().stream()
            .filter(record -> record.tag() == MtefRecord.CHAR && Integer.valueOf(0x8C).equals(record.typeface()))
            .map(record -> record.mtcode()).toList();

        assertEquals(List.of(0xFF08, (int) '+', (int) '+', (int) '1', (int) '0', 0xFF09), textFeCodes);
    }


    @Test
    void testExplicitMathSpacePreservesMathTypeSpaceEncoding() {
        var records = MtefRecordNormalizer.normalize(writer.write(parser.parseLaTeX("1\\ \\ 2"))).records();

        assertEquals(2, records.stream().filter(record ->
            record.tag() == MtefRecord.CHAR
                && Integer.valueOf(0x98).equals(record.typeface())
                && Integer.valueOf(0xEF04).equals(record.mtcode())).count());
    }

    @Test
    void testAsciiFlatParensDoesNotFlattenExplicitFenceWithSourceMetrics() {
        LaTeXNode ast = parser.parseLaTeX("\\left(110+126\\right)");
        FormulaStyleHints hints = new FormulaStyleHints(
            true, false, false, false, false, false, true, false, false, false, new FormulaMetrics(80.0d, 20.0d));
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PAREN}),
            "explicit fences with source metrics still need a MathType fence template");
    }

    @Test
    void testShortMultiplicationEquationWithTallMetricsUsesBoxSegments() {
        LaTeXNode ast = parser.parseLaTeX("3\\times 4=12");
        FormulaStyleHints hints = FormulaStyleHints.empty().withSourceMetrics(new FormulaMetrics(60.0d, 18.0d));
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(3, countOccurrences(mtef, box),
            "18pt xsc multiplication candidates match MathType's boxed operand template pattern");
    }

    @Test
    void testTallVariableMultiplicationCandidateBoxesAllOperands() {
        LaTeXNode ast = parser.parseLaTeX("E\\times F+9=G5");
        FormulaStyleHints hints = FormulaStyleHints.empty().withSourceMetrics(new FormulaMetrics(75.0d, 18.0d));
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(4, countOccurrences(mtef, box),
            "18pt letter/digit multiplication candidates box both factors and each result character");
    }

    @Test
    void testMultiplicationEquationLeavesAddendFlat() {
        LaTeXNode ast = parser.parseLaTeX("3\\times 4+9=21");
        FormulaStyleHints hints = FormulaStyleHints.empty().withSourceMetrics(new FormulaMetrics(75.0d, 18.0d));
        byte[] mtef = writer.write(ast, hints);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(4, countOccurrences(mtef, box),
            "MathType keeps the +9 addend flat while boxing both factors and result digits");
    }

    @Test
    void testVariableMultiplicationEquationStaysFlat() {
        LaTeXNode ast = parser.parseLaTeX("H\\times M=36");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(0, countOccurrences(mtef, box),
            "variable multiplication equations did not match the boxed numeric pattern in the corpus");
    }

    @Test
    void testRepeatedMultiplicationEquationStaysFlat() {
        LaTeXNode ast = parser.parseLaTeX("1\\times 2\\times 3\\times 6=36");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(0, countOccurrences(mtef, box),
            "multi-factor multiplication equations need a separate MathType pattern");
    }

    @Test
    void testLargeNumberMultiplicationEquationStaysFlat() {
        LaTeXNode ast = parser.parseLaTeX("454\\times 229=103966");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        byte[] box = new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BOX, 0x1E, 0x00};
        assertEquals(0, countOccurrences(mtef, box),
            "large-number multiplication equations do not follow the one-digit boxed operand pattern");
    }

    // =====================================================================
    // typesize 差分状态机测试（与真 MathType 差分编码逐字节对齐）
    // 证据：real-src.docx 541 个真对象 + 课程文档 DSMT6/DSMT7 嵌套根式对象。
    // =====================================================================

    @Test
    void testNestedSqrtSkipsRedundantSubBeforeIndexSlot() {
        // 被开方数以 tmSUP 结尾时，字号上下文已是 SUB，
        // 真 MathType（DSMT7 oleObject117）在 index 槽前不再写 SUB。
        LaTeXNode ast = parser.parseLaTeX("\\sqrt{a^{2}+b^{2}}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        // 整个流只应有 2 条 SUB 记录（两个 tmSUP 各一条）；
        // 旧行为会在 ROOT index 槽前再写第 3 条。
        byte sub = (byte) MtefRecord.SUB;
        int subCount = 0;
        // 从表达式区域开始数（跳过前缀——前缀 EQN_PREFS 不含孤立 0x0b 记录语义，
        // 但保险起见用 END+SUB 序列统计 ROOT index 前的 SUB 是否出现）
        for (int i = 12; i < mtef.length; i++) {
            if (mtef[i] == sub) {
                subCount++;
            }
        }
        assertEquals(2, subCount,
            "radicand ends with tmSUP: size context already SUB, root index slot must not emit another SUB");
        // 序列校验：radicand LINE 的 END 之后应直接跟 index NULL LINE（00 00 01 01）。
        // 旧行为是 00 00 0b 01 01（END+END+SUB+NULL LINE），真 MathType 不写那个 SUB。
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.END, (byte) MtefRecord.END,
                        (byte) MtefRecord.COLOR, 0x00, (byte) MtefRecord.LINE, 0x01}),
            "radicand END must be followed directly by the index NULL LINE (no SUB in between)");
    }

    @Test
    void testSimpleSqrtStillEmitsSubBeforeIndexSlot() {
        // 被开方数以普通字符结尾（上下文 FULL）时，index 槽前必须写 SUB。
        LaTeXNode ast = parser.parseLaTeX("\\sqrt{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.END, (byte) MtefRecord.SUB,
                        (byte) MtefRecord.COLOR, 0x00, (byte) MtefRecord.LINE, 0x01}),
            "radicand ends at FULL size: root index slot needs the SUB typesize record");
    }

    @Test
    void testFractionWithPlainNumeratorHasNoFullBeforeFollowingSibling() {
        // 分子以普通字符结尾（上下文 FULL）且分式后还有兄弟时，
        // 真 MathType 不写 FULL（222/224 真分式验证）。
        LaTeXNode ast = parser.parseLaTeX("\\frac{a}{b}+c");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.END, (byte) MtefRecord.FULL,
                        (byte) MtefRecord.CHAR}),
            "fraction leaving FULL context must not be followed by a spurious FULL record");
    }

    @Test
    void testFractionWithScriptedNumeratorKeepsFullBetweenSlots() {
        // 分子以 tmSUB 结尾（上下文 SUB）时，槽间必须写 FULL 恢复分母字号
        // （真 MathType oleObject144 实测模式）。
        LaTeXNode ast = parser.parseLaTeX("\\frac{b_{8}}{81}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.END, (byte) MtefRecord.FULL,
                        (byte) MtefRecord.LINE, 0x00}),
            "numerator ending with a script template leaves SUB context: FULL required before denominator slot");
    }

    @Test
    void niceFractionUsesCompleteMathTypeSlashFractionVariation() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\nicefrac12"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_FRACT, 0x03, 0x00
        }), "nicefrac requires SMALL and SLASH variation bits");
        assertFalse(containsAscii(mtef, "TeX Input Language"),
            "unsupported MathType TeX source must not override native slash-fraction records");
    }

    @Test
    void squareRootUsesNativeTemplateWithoutTexReparseMetadata() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\sqrt x"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ROOT, 0x00, 0x00
        }), "the native square-root template must remain present");
        assertFalse(containsAscii(mtef, "TeX Input Language"),
            "Format Equations must consume the native root instead of reparsing the source tag");
        assertTrue(hasDefaultBlackColorState(mtef),
            "native root serialization must retain MathType's black color state");
    }

    @Test
    void layoutSensitiveStructuresUseNativeColorSerializationWithoutTexReparseMetadata() {
        for (String latex : new String[] {
            "\\cfrac{2}{1+\\cfrac21}",
            "\\frac{\\displaystyle\\frac12}3",
            "\\textstyle\\frac{\\textstyle\\frac12}2",
            "\\int\\limits_a^b"
        }) {
            byte[] mtef = writer.write(parser.parseLaTeX(latex));
            assertFalse(containsAscii(mtef, "TeX Input Language"), latex);
            assertTrue(hasDefaultBlackColorState(mtef), latex);
        }

        byte[] indexedRoot = writer.write(parser.parseLaTeX("\\sqrt[x+1]{2}"));
        assertTrue(containsBytes(indexedRoot, new byte[] {
            (byte) MtefRecord.COLOR, 0x01,
            (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ROOT, 0x01, 0x00,
            (byte) MtefRecord.COLOR, 0x00,
            (byte) MtefRecord.LINE, 0x00,
            (byte) MtefRecord.COLOR, 0x01,
            (byte) MtefRecord.CHAR, 0x00, (byte) 0x88, 0x32, 0x00
        }), "indexed root must scope each MathType template slot like TeXToggle");
    }

    @Test
    void directionCommandsDoNotWriteUnsupportedMathTypeTexMetadata() {
        assertFalse(containsAscii(writer.write(parser.parseLaTeX("\\ltr{x+3}")), "TeX Input Language"));
        assertFalse(containsAscii(writer.write(parser.parseLaTeX("\\rtl{x+3}")), "TeX Input Language"));
    }

    private boolean hasDefaultBlackColorState(byte[] mtef) {
        return containsBytes(mtef, new byte[] {
            (byte) MtefRecord.COLOR_DEF, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            (byte) MtefRecord.COLOR, 0x01
        });
    }

    @Test
    void extendedLabeledArrowsDoNotWriteUnsupportedMathTypeTexMetadata() {
        assertFalse(containsAscii(
            writer.write(parser.parseLaTeX("A\\xlongrightarrow{f}B")), "TeX Input Language"));
        assertFalse(containsAscii(
            writer.write(parser.parseLaTeX("A\\xLongleftarrow{f}B")), "TeX Input Language"));
        assertFalse(containsAscii(
            writer.write(parser.parseLaTeX("A\\xlongequal{f}B")), "TeX Input Language"));
    }

    @Test
    void xLongLeftArrowWritesMathTypeSupportedEquivalentTexMetadata() {
        byte[] mtef = writer.write(parser.parseLaTeX("B\\xLongleftarrow[\\text{seilpmi}]{}A"));

        assertTrue(containsAscii(mtef,
            "B\\mathop{\\Longleftarrow}\\limits_{\\rm seilpmi}A"));
        assertFalse(containsAscii(mtef, "\\xLongleftarrow"));
    }

    @Test
    void desktopEquivalentOfficialExamplesUseTheirPinnedTexMetadata() {
        byte[] mbox = writer.write(parser.parseLaTeX("\\mbox{\\large$T$}_0^2"));
        byte[] mixedText = writer.write(parser.parseLaTeX("\\text{If $x=0$ then $y=2$.}"));

        assertFalse(containsAscii(mbox, "TeX Input Language"));
        assertFalse(containsAscii(mbox, "\\mbox"));
        assertTrue(containsBytes(mbox, new byte[] {
            (byte) MtefRecord.CHAR, (byte) 0x80, (byte) (MtefRecord.FN_TEXT | 0x80), '\\', 0
        }));
        assertTrue(containsAscii(mixedText, "\\text{If }x=0\\text{ then }y=2\\text{.}"));
        assertFalse(containsAscii(mixedText, "$x=0$"));
        assertTrue(containsBytes(mixedText, new byte[] {
            (byte) MtefRecord.CHAR, 0, (byte) (MtefRecord.FN_TEXT | 0x80), 'I', 0,
            (byte) MtefRecord.CHAR, 0, (byte) (MtefRecord.FN_TEXT | 0x80), 'f', 0
        }));
        assertTrue(containsBytes(mixedText, new byte[] {
            (byte) MtefRecord.CHAR, 0, (byte) (MtefRecord.FN_VARIABLE | 0x80), 'x', 0
        }));
    }

    @Test
    void mathcalUppercaseUsesPinnedEuclidMathTwoEncodings() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\mathcal{LZF}"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
            0x7F, 0x12, 0x21, 0x4C
        }), "script L must use U+2112 at Euclid Math Two position 0x4C");
        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
            0x7F, 0x19, (byte) 0xF1, 0x5A
        }), "script Z must use MathType private code F119 at position 0x5A");
        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_CHAR_ENC_CHAR_8,
            0x7F, 0x31, 0x21, 0x46
        }), "script F must use U+2131 at Euclid Math Two position 0x46");
        assertTrue(containsAscii(mtef, "EuclidMath2"));
    }

    @Test
    void officialDynamicSymbolTypefacesUsePinnedEuclidFamilies() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\approxeq,\\barwedge"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.ENCODING_DEF, 'E', 'u', 'c', 'l', 'i', 'd', 'M', 'a', 't', 'h', '1', 0,
            (byte) MtefRecord.FONT_DEF, 0x07
        }), "typeface 0x7F symbols must bind Euclid Math One");
        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.ENCODING_DEF, 'E', 'u', 'c', 'l', 'i', 'd', 'M', 'a', 't', 'h', '2', 0,
            (byte) MtefRecord.FONT_DEF, 0x08
        }), "typeface 0x7E symbols must bind Euclid Math Two");

        byte[] reversed = writer.write(parser.parseLaTeX("\\sqcap,\\oslash,\\sqsubset,\\vdash"));
        assertEquals("EuclidMath2", MtefCharMap.lookup("\\sqcap").dynamicFontProfile());
        assertEquals("EuclidMath1", MtefCharMap.lookup("\\oslash").dynamicFontProfile());
        assertTrue(containsBytes(reversed, new byte[] {
            (byte) MtefRecord.ENCODING_DEF, 'E', 'u', 'c', 'l', 'i', 'd', 'M', 'a', 't', 'h', '2', 0,
            (byte) MtefRecord.FONT_DEF, 0x07
        }), "sqcap must bind typeface 0x7F to Euclid Math Two");
        assertTrue(containsBytes(reversed, new byte[] {
            (byte) MtefRecord.ENCODING_DEF, 'E', 'u', 'c', 'l', 'i', 'd', 'M', 'a', 't', 'h', '1', 0,
            (byte) MtefRecord.FONT_DEF, 0x08
        }), "oslash must bind typeface 0x7E to Euclid Math One");
    }

    @Test
    void extendedRelationsAndBlackboardFBindEuclidMathTwo() {
        assertEquals("EuclidMath2", MtefCharMap.lookup("\\Subset").dynamicFontProfile());

        byte[] relations = writer.write(parser.parseLaTeX("\\Subset,\\succapprox,\\supsetneqq"));
        assertTrue(containsAscii(relations, "EuclidMath2"));
        assertTrue(containsAscii(relations, "Euclid Math Two"));

        byte[] blackboard = writer.write(parser.parseLaTeX("\\mathbb{F}"));
        assertTrue(containsAscii(blackboard, "EuclidMath2"));
        assertTrue(containsAscii(blackboard, "Euclid Math Two"));
    }

    @Test
    void mathttUsesPinnedCourierFontStyleDefinition() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\mathtt x"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.FONT_STYLE_DEF, 0x03, 0x00,
            (byte) MtefRecord.CHAR, 0x00, 0x7F, 0x78, 0x00
        }), "mathtt must bind typeface 0x7F to the prefix Courier New font definition");
        assertFalse(containsAscii(mtef, "TeX Input Language"));
    }

    @Test
    void unicodeDiagonalArrowPairUsesTextFeWithoutUnsupportedTexMetadata() {
        byte[] mtef = writer.write(parser.parseLaTeX("\\nwsearrow\\neswarrow"));

        assertTrue(containsBytes(mtef, new byte[] {
            (byte) MtefRecord.CHAR, 0x00, (byte) 0x8C, 0x21, 0x29,
            (byte) MtefRecord.CHAR, 0x00, (byte) 0x8C, 0x22, 0x29
        }));
        assertFalse(containsAscii(mtef, "TeX Input Language"));
    }

    @Test
    void testLimFollowedByContentEmitsFullAfterTemplate() {
        // 真 MathType 在 tmSUMOP 模板结束后、后续兄弟之前写 FULL 复位
        // （operator 槽把上下文留在 SYM）。
        LaTeXNode ast = parser.parseLaTeX("\\lim_{x \\to 0} \\frac{\\sin x}{x}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.END, (byte) MtefRecord.FULL,
                        (byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_FRACT}),
            "tmSUMOP template must be followed by a FULL record before the next sibling");
        // 缩小槽（lower slot）内部的多字符内容保持 SUB 字号：
        // 'x'(fn=VARIABLE 0x83, mt=0x0078) 之后必须直接跟下一条 CHAR 记录，
        // 不能插入 FULL——真 MathType 的缩小槽内不写字号恢复记录。
        assertTrue(containsBytes(mtef, new byte[]{(byte) 0x83, 0x78, 0x00, (byte) MtefRecord.CHAR}),
            "reduced (SUB) slot content must stay at SUB size: no FULL between slot characters");
    }

    @Test
    void testSumVariationIncludesSummationStyleBit() {
        // 真 MathType tmSUM variation 恒含 TV_BO_SUM(0x40)：
        // oleObject37/39 实测 var=0x70 = SUM|LOWER|UPPER。
        LaTeXNode ast = parser.parseLaTeX("\\sum_{i=1}^{n} i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00,
                        (byte) MtefRecord.TM_SUM, 0x70, 0x00}),
            "sum with both limits should use variation 0x70 (TV_BO_SUM|TV_BO_LOWER|TV_BO_UPPER)");
    }

    private boolean containsRecord(byte[] bytes, int recordType) {
        for (int i = 12; i < bytes.length; i++) {
            if ((bytes[i] & 0xFF) == recordType) {
                return true;
            }
        }
        return false;
    }

    private boolean containsBytes(byte[] bytes, byte[] needle) {
        return indexOfBytes(bytes, needle) >= 0;
    }

    private int indexOfBytes(byte[] bytes, byte[] needle) {
        outer:
        for (int i = 0; i <= bytes.length - needle.length; i++) {
            for (int j = 0; j < needle.length; j++) {
                if (bytes[i + j] != needle[j]) {
                    continue outer;
                }
            }
            return i;
        }
        return -1;
    }

    private boolean containsAscii(byte[] bytes, String needle) {
        return new String(bytes, java.nio.charset.StandardCharsets.US_ASCII).contains(needle);
    }

    private String extractDigitStream(byte[] bytes) {
        StringBuilder digits = new StringBuilder();
        for (int i = 0; i <= bytes.length - 5; i++) {
            if ((bytes[i] & 0xFF) != MtefRecord.CHAR) {
                continue;
            }
            int options = bytes[i + 1] & 0xFF;
            int mtcode = (bytes[i + 3] & 0xFF) | ((bytes[i + 4] & 0xFF) << 8);
            char ch = (char) mtcode;
            if (Character.isDigit(ch)) {
                digits.append(ch);
            }
            i += ((options & MtefRecord.OPT_CHAR_ENC_CHAR_8) != 0) ? 5 : 4;
        }
        return digits.toString();
    }

    private String extractCharStream(byte[] bytes) {
        StringBuilder chars = new StringBuilder();
        for (int i = 0; i <= bytes.length - 5; i++) {
            if ((bytes[i] & 0xFF) != MtefRecord.CHAR) {
                continue;
            }
            int options = bytes[i + 1] & 0xFF;
            if ((options & MtefRecord.OPT_CHAR_ENC_NO_MTCODE) != 0) {
                continue;
            }
            int pos = i + 2;
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                pos += (pos < bytes.length && (bytes[pos] & 0xFF) == 0x80) ? 6 : 2;
            }
            if (pos + 2 >= bytes.length) {
                continue;
            }
            int mtcode = (bytes[pos + 1] & 0xFF) | ((bytes[pos + 2] & 0xFF) << 8);
            chars.append((char) mtcode);
            i += ((options & MtefRecord.OPT_CHAR_ENC_CHAR_8) != 0) ? 5 : 4;
        }
        return chars.toString();
    }

    private int countOccurrences(byte[] bytes, byte[] needle) {
        int count = 0;
        outer:
        for (int i = 0; i <= bytes.length - needle.length; i++) {
            for (int j = 0; j < needle.length; j++) {
                if (bytes[i + j] != needle[j]) {
                    continue outer;
                }
            }
            count++;
        }
        return count;
    }
}
