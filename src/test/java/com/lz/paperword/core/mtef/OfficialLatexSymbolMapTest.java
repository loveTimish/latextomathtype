package com.lz.paperword.core.mtef;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class OfficialLatexSymbolMapTest {

    @Test
    void pinnedMathTypeMappingsOverrideGenericAmsUnicodeMappings() {
        MtefCharMap.CharEntry approxeq = MtefCharMap.lookup("\\approxeq");
        MtefCharMap.CharEntry blackTriangleLeft = MtefCharMap.lookup("\\blacktriangleleft");

        assertNotNull(approxeq);
        assertEquals(0x224A, approxeq.mtcode());
        assertNotNull(blackTriangleLeft);
        assertEquals(0x25C0, blackTriangleLeft.mtcode());
    }

    @Test
    void curatedMathTypeMappingsTakePrecedenceOverGeneratedFallbacks() {
        MtefCharMap.CharEntry times = MtefCharMap.lookup("\\times");
        MtefCharMap.CharEntry nearrow = MtefCharMap.lookup("\\nearrow");

        assertEquals(MtefRecord.FN_SYMBOL, times.typeface());
        assertEquals(MtefRecord.FN_MTEXTRA, nearrow.typeface());
    }

    @Test
    void encodingProfileMarksPinnedMathTypeEncodingAsVerified() {
        MtefCharMap.EncodingProfile pinned = MtefCharMap.encodingProfile("\\approxeq");

        assertEquals(MtefCharMap.MappingSource.PINNED_MATHTYPE, pinned.source());
        assertEquals(0xAC, pinned.bits8());
        assertEquals(true, pinned.mathTypeVerified());
    }

    @Test
    void compositeBoxAndCircleSymbolsUseMathTypeLegacyEncodings() {
        assertPinned("\\boxminus", 0x7F, 0x229F, 0x27);
        assertPinned("\\boxplus", 0x7F, 0x229E, 0x28);
        assertPinned("\\boxtimes", 0x7F, 0x22A0, 0x29);
        assertPinned("\\Bumpeq", 0x7F, 0x224E, 0xA7);
        assertPinned("\\Cap", 0x7F, 0x22D2, 0xD3);
        assertPinned("\\checkmark", MtefRecord.FN_MTEXTRA, 0xFFFD, 0x6E);
        assertPinned("\\circeq", 0x7F, 0x2257, 0xA1);
        assertPinned("\\circledast", 0x7F, 0x229B, 0x23);
        assertPinned("\\circledcirc", 0x7F, 0x229A, 0x22);
        assertPinned("\\circleddash", 0x7F, 0x2296, 0x21);
        assertPinned("\\circledS", 0x7F, 0x24C8, 0x26);
        assertPinned("\\nwsearrow", MtefRecord.FN_TEXT_FE, 0x2921, -1);
        assertPinned("\\neswarrow", MtefRecord.FN_TEXT_FE, 0x2922, -1);
    }

    private void assertPinned(String command, int typeface, int mtcode, int bits8) {
        MtefCharMap.EncodingProfile profile = MtefCharMap.encodingProfile(command);
        assertEquals(MtefCharMap.MappingSource.PINNED_MATHTYPE, profile.source(), command);
        assertEquals(typeface, profile.typeface(), command);
        assertEquals(mtcode, profile.mtcode(), command);
        assertEquals(bits8, profile.bits8(), command);
        assertEquals(true, profile.mathTypeVerified(), command);
    }
}
