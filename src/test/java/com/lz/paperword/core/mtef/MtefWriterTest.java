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
        assertEquals(MtefRecord.FN_SYMBOL, nearrow.typeface());
        assertEquals(0x2197, nearrow.mtcode());

        MtefCharMap.CharEntry searrow = MtefCharMap.lookup("\\searrow");
        assertNotNull(searrow);
        assertEquals(MtefRecord.FN_SYMBOL, searrow.typeface());
        assertEquals(0x2198, searrow.mtcode());
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
    void testWriteExplicitLongDivisionAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[65]{13}{845}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "explicit long division should emit MATRIX wrapper");
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "long division should emit TMPL record");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TM_LDIV, 0x01, 0x00}),
            "long division should use tmLDIV with upper slot variation");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TM_FRACT, 0x00, 0x00}),
            "long division steps should include fraction templates");
    }

    @Test
    void testWriteExplicitLongDivisionSerializesDividendBeforeQuotient() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[129]{12}{1548}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        String digitStream = extractDigitStream(mtef);
        assertTrue(digitStream.startsWith("121548"),
            "tmLDIV should still write divisor and dividend slot content first");
        assertTrue(digitStream.endsWith("129"),
            "tmLDIV should still write quotient slot after the structured dividend content");
    }

    @Test
    void testWriteExplicitLongDivisionWithoutQuotientAsTemplate() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv{13}{845}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "long division without quotient should emit MATRIX wrapper");
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "long division without quotient should emit TMPL record");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TM_LDIV, 0x00, 0x00}),
            "long division without quotient should use tmLDIV without upper slot");
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TM_FRACT, 0x00, 0x00}),
            "long division without quotient should still include fraction steps");
    }

    @Test
    void testWriteThreeStepLongDivisionKeepsStructuredFractions() {
        LaTeXNode ast = parser.parseLaTeX("\\longdiv[246]{5}{1234}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsBytes(mtef, new byte[]{(byte) MtefRecord.TM_LDIV, 0x01, 0x00}),
            "three-step long division should still use tmLDIV");
        assertTrue(countOccurrences(mtef, new byte[]{(byte) MtefRecord.TM_FRACT, 0x00, 0x00}) >= 3,
            "three-step long division should emit one fraction template per step");
    }

    @Test
    void testWriteLongDivisionEnvironmentStillUsesMatrixFallback() {
        LaTeXNode ast = parser.parseLaTeX("\\begin{longdivision}{r|l}13 & 845 \\\\ \\hline & 65\\end{longdivision}");
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "longdivision environment should remain MATRIX fallback");
    }

    @Test
    void testWriteCompositeLongDivisionWithStepsKeepsTemplateAndMatrix() {
        LaTeXNode ast = parser.parseLaTeX(
            "\\begin{array}{l}{\\longdiv[65]{13}{845}} \\\\ {\\begin{array}{rrr}{7} & {8} & {} \\\\ \\hline {} & {6} & {5} \\\\ {} & {6} & {5} \\\\ \\hline {} & {} & {0}\\end{array}}\\end{array}"
        );
        byte[] mtef = writer.write(ast);

        assertNotNull(mtef);
        assertTrue(containsRecord(mtef, MtefRecord.TMPL), "composite long division should still include tmLDIV");
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "composite long division should still include matrix steps");
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
