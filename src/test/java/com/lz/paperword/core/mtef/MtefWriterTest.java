package com.lz.paperword.core.mtef;

import com.lz.paperword.core.latex.LaTeXNode;
import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class MtefWriterTest {

    private final LaTeXParser parser = new LaTeXParser();
    private final MtefWriter writer = new MtefWriter();

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
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x50, 0x00}),
            "limit should use tmLim with lower-slot and summation-style placement");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.SYM}),
            "tmLim path should not emit a separate SYM operator record");
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
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_COPROD, 0x30, 0x00}),
            "coproduct should use tmCOPROD with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
            "coproduct should not fall back to tmSUM");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_PROD, 0x30, 0x00}),
            "coproduct should not fall back to tmPROD");
    }

    @Test
    void testWriteBigcupUsesDedicatedUnionTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\bigcup_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UNION, 0x30, 0x00}),
            "bigcup should use tmUNION with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
            "bigcup should not fall back to tmSUM");
    }

    @Test
    void testWriteBigcapUsesDedicatedIntersectionTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\bigcap_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTER, 0x30, 0x00}),
            "bigcap should use tmINTER with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
            "bigcap should not fall back to tmSUM");
    }

    @Test
    void testWriteBigoplusUsesGenericSummationStyleTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\bigoplus_{i=1}^{n} A_i");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUMOP, 0x30, 0x00}),
            "bigoplus should use tmSUMOP with both upper/lower limit bits");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
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
    void testWriteXrightarrowUsesTmArrowTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\xrightarrow{n\\to\\infty}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x24, 0x00}),
            "xrightarrow should use tmARROW with top-slot + right-arrow variation");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0x92, 0x21}),
            "xrightarrow should serialize the expandable right arrow glyph in FN_EXPAND");
    }

    @Test
    void testWriteXleftarrowUsesTmArrowTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\xleftarrow{f}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x14, 0x00}),
            "xleftarrow should use tmARROW with top-slot + left-arrow variation");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0x90, 0x21}),
            "xleftarrow should serialize the expandable left arrow glyph in FN_EXPAND");
    }

    @Test
    void testWriteOverbraceUsesTmHBRACEWithTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\overbrace{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACE, 0x01, 0x00}),
            "overbrace should use tmHBRACE with TV_HB_TOP variation bit set");
        // Should contain 0x23DE (⏞ overbrace character) in FN_EXPAND (encoded)
        // encodeTypeface(FN_EXPAND = 22) → (22 | 0x80) = 0x96
        // mtcode: 0x23DE → low byte 0xDE, high byte 0x23
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xDE, 0x23}),
            "overbrace should write U+23DE overbrace character with FN_EXPAND font");
    }

    @Test
    void testWriteUnderbraceUsesTmHBRACEWithNoTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\underbrace{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACE, 0x00, 0x00}),
            "underbrace should use tmHBRACE with TV_HB_TOP variation bit clear");
        // Should contain 0x23DF (⏟ underbrace character) in FN_EXPAND (encoded)
        // encodeTypeface(FN_EXPAND = 22) → (22 | 0x80) = 0x96
        // mtcode: 0x23DF → low byte 0xDF, high byte 0x23
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xDF, 0x23}),
            "underbrace should write U+23DF underbrace character with FN_EXPAND font");
    }

    @Test
    void testWriteOverbracketUsesTmHBRACKWithTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\overbracket{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACK, 0x01, 0x00}),
            "overbracket should use tmHBRACK with TV_HB_TOP variation bit set");
        // Should contain 0x23B4 (⎴ top square bracket) in FN_EXPAND (encoded)
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xB4, 0x23}),
            "overbracket should write U+23B4 top square bracket with FN_EXPAND font");
    }

    @Test
    void testWriteUnderbracketUsesTmHBRACKWithNoTopVariation() {
        LaTeXNode ast = parser.parseLaTeX("\\underbracket{x+1}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACK, 0x00, 0x00}),
            "underbracket should use tmHBRACK with TV_HB_TOP variation bit clear");
        // Should contain 0x23B5 (⎵ bottom square bracket) in FN_EXPAND (encoded)
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xB5, 0x23}),
            "underbracket should write U+23B5 bottom square bracket with FN_EXPAND font");
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
    void testWriteMathIrDirectlyForXrightarrowArrowTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\xrightarrow{f}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x24, 0x00}),
            "IR path should lower xrightarrow to tmARROW");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0x92, 0x21}),
            "IR path should keep the expandable right arrow glyph");
    }

    @Test
    void testWriteMathIrDirectlyForXleftarrowArrowTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\xleftarrow{g}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_ARROW, 0x14, 0x00}),
            "IR path should lower xleftarrow to tmARROW");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0x90, 0x21}),
            "IR path should keep the expandable left arrow glyph");
    }

    @Test
    void testWriteMathIrDirectlyForHorizontalBracketTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\overbracket{x+1}"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_HBRACK, 0x01, 0x00}),
            "IR path should lower overbracket to tmHBRACK");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.CHAR, 0x00, (byte) 0x96, (byte) 0xB4, 0x23}),
            "IR path should keep the top square bracket expandable character");
    }

    @Test
    void testWriteMathIrDirectlyForLimitTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\lim_{n\\to\\infty} a_n"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x50, 0x00}),
            "IR path should lower \\lim to tmLIM");
    }

    @Test
    void testWriteMathIrDirectlyForCoproductTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\coprod_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_COPROD, 0x30, 0x00}),
            "IR path should lower \\coprod to tmCOPROD");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
            "IR path should not route \\coprod through tmSUM fallback");
    }

    @Test
    void testWriteMathIrDirectlyForUnionTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bigcup_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_UNION, 0x30, 0x00}),
            "IR path should lower \\bigcup to tmUNION");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
            "IR path should not route \\bigcup through tmSUM fallback");
    }

    @Test
    void testWriteMathIrDirectlyForIntersectionTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bigcap_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_INTER, 0x30, 0x00}),
            "IR path should lower \\bigcap to tmINTER");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUM, 0x30, 0x00}),
            "IR path should not route \\bigcap through tmSUM fallback");
    }

    @Test
    void testWriteMathIrDirectlyForSummationStyleGenericBigOperatorTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\bigotimes_{i=1}^{n} A_i"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_SUMOP, 0x30, 0x00}),
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

    private boolean containsRecord(byte[] bytes, int recordType) {
        for (int i = 12; i < bytes.length; i++) {
            if ((bytes[i] & 0xFF) == recordType) {
                return true;
            }
        }
        return false;
    }

    private boolean containsBytes(byte[] bytes, byte[] needle) {
        outer:
        for (int i = 0; i <= bytes.length - needle.length; i++) {
            for (int j = 0; j < needle.length; j++) {
                if (bytes[i + j] != needle[j]) {
                    continue outer;
                }
            }
            return true;
        }
        return false;
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
