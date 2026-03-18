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
        assertTrue(containsRecord(mtef, MtefRecord.MATRIX), "decimal array should emit MATRIX record");
        assertTrue(containsBytes(mtef, new byte[]{'.', 0x00}), "decimal point should be serialized in MTEF");
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
}
