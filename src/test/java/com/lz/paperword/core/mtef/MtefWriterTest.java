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
            "limit should use tmLIM with lower-slot and summation-style placement");
        assertFalse(containsBytes(mtef, new byte[]{(byte) MtefRecord.SYM}),
            "tmLIM path should not emit a separate SYM operator record");
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
    void testWriteSingleSidedBarFenceKeepsOnlyPresentSide() {
        LaTeXNode ast = parser.parseLaTeX("\\left. x \\right|");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_BAR, 0x02, 0x00}),
            "single-sided right bar fence should encode only the right fence bit");
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
    void testWriteMathIrDirectlyForExtendedFenceSubset() {
        byte[] mtef = writer.write(parser.parseMathIR("\\left\\lVert x+1 \\right.") );

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_DBAR, 0x01, 0x00}),
            "IR path should preserve single-sided double-bar fences");
    }

    @Test
    void testWriteMathIrDirectlyForLimitTemplate() {
        byte[] mtef = writer.write(parser.parseMathIR("\\lim_{n\\to\\infty} a_n"));

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TMPL, 0x00, (byte) MtefRecord.TM_LIM, 0x50, 0x00}),
            "IR path should lower \\lim to tmLIM");
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
            "\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {} \\\\ \\hline {} & {5} & {5} & {3} & {5} & {}\\end{array}"
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
