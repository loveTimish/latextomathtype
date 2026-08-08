package com.lz.paperword.core.mtef;

import com.lz.paperword.core.latex.LaTeXParser;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class MtefRecordNormalizerTest {

    @Test
    void ignoresDefinitionOrderingAndOptionalTopLevelLine() {
        MtefRecordNormalizer.CanonicalRecord line = record(MtefRecord.LINE, "LINE", false);
        MtefRecordNormalizer.CanonicalRecord colorDef = record(MtefRecord.COLOR_DEF, "COLOR_DEF", null);
        MtefRecordNormalizer.CanonicalRecord fontDef = record(MtefRecord.FONT_DEF, "FONT_DEF", null);
        MtefRecordNormalizer.CanonicalRecord character = new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.CHAR, "CHAR", 0, null, null, 0x83, (int) 'x', null, null, null, null);
        MtefRecordNormalizer.CanonicalRecord end = record(MtefRecord.END, "END", null);

        List<MtefRecordNormalizer.CanonicalRecord> standard = List.of(character, end);
        List<MtefRecordNormalizer.CanonicalRecord> generated = List.of(
            line, colorDef, fontDef, character, end, end);

        assertEquals(standard, MtefRecordNormalizer.canonicalFormulaRecords(generated));
    }

    @Test
    void nullLinesDoNotConsumeContainerEndRecords() {
        MtefRecordNormalizer.CanonicalRecord line = record(MtefRecord.LINE, "LINE", false);
        MtefRecordNormalizer.CanonicalRecord nullLine = record(MtefRecord.LINE, "LINE", true);
        MtefRecordNormalizer.CanonicalRecord end = record(MtefRecord.END, "END", null);

        assertEquals(List.of(nullLine, end),
            MtefRecordNormalizer.canonicalFormulaRecords(List.of(line, nullLine, end, end)));
    }

    @Test
    void removesRedundantBlackStatesButKeepsActualColorTransitions() {
        MtefRecordNormalizer.CanonicalRecord blackDefinition = new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.COLOR_DEF, "COLOR_DEF", 0, null, null, null, null, 0, null, null, null);
        MtefRecordNormalizer.CanonicalRecord redDefinition = new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.COLOR_DEF, "COLOR_DEF", 0, null, null, null, null, 1000 << 20, null, null, null);
        MtefRecordNormalizer.CanonicalRecord black = color(1);
        MtefRecordNormalizer.CanonicalRecord red = color(2);
        MtefRecordNormalizer.CanonicalRecord reset = color(0);

        List<MtefRecordNormalizer.CanonicalRecord> records = MtefRecordNormalizer.canonicalFormulaRecords(
            List.of(blackDefinition, black, redDefinition, red, red, reset));

        assertEquals(List.of(color(1000 << 20), color(0)), records);
    }

    @Test
    void removesInitialDefaultFullSizeState() {
        MtefRecordNormalizer.CanonicalRecord full = record(MtefRecord.FULL, "FULL", null);
        MtefRecordNormalizer.CanonicalRecord character = new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.CHAR, "CHAR", 0, null, null, 0x83, (int) 'x', null, null, null, null);

        assertEquals(List.of(character), MtefRecordNormalizer.canonicalFormulaRecords(List.of(full, character)));
    }

    @Test
    void symbolTypesizeDoesNotConsumeTheFollowingCharacterRecord() {
        byte[] mtef = new MtefWriter().write(new LaTeXParser().parseLaTeX("\\int_a^b"));
        List<MtefRecordNormalizer.CanonicalRecord> records = MtefRecordNormalizer.normalize(mtef).records();
        int symbolSize = -1;
        for (int index = 0; index < records.size(); index++) {
            if (records.get(index).tag() == MtefRecord.SYM) {
                symbolSize = index;
                break;
            }
        }

        assertEquals("LINE", records.get(symbolSize + 1).name());
        assertEquals("CHAR", records.get(symbolSize + 2).name());
        assertEquals(0x222B, records.get(symbolSize + 2).mtcode());
    }

    @Test
    void legacyModeCanonicalizesMathTypeSevenFlatParenthesisTemplates() {
        MtefRecordNormalizer.CanonicalRecord template = new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.TMPL, "TMPL", 0, MtefRecord.TM_PAREN, 3,
            null, null, 0, null, null, null);
        MtefRecordNormalizer.CanonicalRecord line = record(MtefRecord.LINE, "LINE", false);
        MtefRecordNormalizer.CanonicalRecord end = record(MtefRecord.END, "END", null);
        MtefRecordNormalizer.CanonicalRecord value = character(0x83, 'x');
        MtefRecordNormalizer.CanonicalRecord expandableLeft = character(0x96, '(');
        MtefRecordNormalizer.CanonicalRecord expandableRight = character(0x96, ')');

        assertEquals(
            List.of(character(0x82, '('), value, character(0x82, ')')),
            MtefRecordNormalizer.canonicalizeLegacyFenceRecords(List.of(
                template, line, value, end, expandableLeft, expandableRight, end)));
    }

    @Test
    void findsDsmtSixFormulaLineAfterColorStateWithoutFullRecord() {
        byte[] mtef = new byte[] {
            5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 1,
            (byte) MtefRecord.EQN_PREFS,
            0, 0, 0, 0,
            (byte) MtefRecord.COLOR_DEF, 4,
            0, 0, 0, 0, 0, 0, 'B', 'l', 'a', 'c', 'k', 0,
            (byte) MtefRecord.COLOR, 1,
            (byte) MtefRecord.LINE, 0,
            (byte) MtefRecord.CHAR, 0, (byte) 0x83, 'x', 0,
            (byte) MtefRecord.END
        };

        assertEquals(List.of(character(0x83, 'x')),
            MtefRecordNormalizer.normalize(mtef).records());
    }

    @Test
    void skipsDsmtSixEquationPreferenceArraysExactly() {
        byte[] mtef = new byte[] {
            5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
            (byte) MtefRecord.EQN_PREFS,
            0,                    // options
            1, 0x21, 0x2F,       // one size: 12pt
            2, 0x20, (byte) 0xF4, 0x50, (byte) 0xF0, // two spacing dimensions
            2, 1, 0, 0,          // two styles: one used, one unused
            (byte) MtefRecord.LINE, 0,
            (byte) MtefRecord.CHAR, 0, (byte) 0x82, '.', 0,
            (byte) MtefRecord.END
        };

        MtefRecordNormalizer.NormalizationReport report = MtefRecordNormalizer.normalize(mtef);

        assertEquals("CHAR:130:46:null", report.canonicalSignature());
    }

    @Test
    void normalizesLegacyAndCurrentEmptyEquationsEqually() {
        byte[] legacy = new byte[] {
            5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
            (byte) MtefRecord.EQN_PREFS, 0, 0, 0, 0,
            (byte) MtefRecord.FULL, (byte) MtefRecord.END
        };
        byte[] current = new byte[] {
            5, 1, 0, 7, 0, 'D', 'S', 'M', 'T', '7', 0, 0,
            (byte) MtefRecord.LINE, 0,
            (byte) MtefRecord.END, (byte) MtefRecord.END
        };

        assertEquals("END", MtefRecordNormalizer.normalize(legacy).canonicalSignature());
        assertEquals(
            MtefRecordNormalizer.normalize(legacy).canonicalSignature(),
            MtefRecordNormalizer.normalize(current).canonicalSignature());
    }

    @Test
    void recognizesLegacyRootLineWithInlineRulerData() {
        byte[] mtef = new byte[] {
            5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
            (byte) MtefRecord.EQN_PREFS, 0, 0, 0, 0,
            (byte) MtefRecord.FULL,
            (byte) MtefRecord.LINE, (byte) MtefRecord.OPT_LP_RULER,
            1, 0, (byte) 0xC8, 0x0D,
            (byte) MtefRecord.CHAR, 0, (byte) 0x86, '=', 0,
            (byte) MtefRecord.END, (byte) MtefRecord.END
        };

        assertEquals("CHAR:134:61:null|END",
            MtefRecordNormalizer.normalize(mtef).canonicalSignature());
    }

    @Test
    void canonicalizesLegacyExplicitTimesAsciiParentheses() {
        assertEquals(
            List.of(character(0x82, '('), character(0x82, ')')),
            MtefRecordNormalizer.normalizeLegacyCompatible(new byte[] {
                5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
                (byte) MtefRecord.LINE, 0,
                (byte) MtefRecord.CHAR, 0, 0x7F, '(', 0,
                (byte) MtefRecord.CHAR, 0, 0x7F, ')', 0,
                (byte) MtefRecord.END
            }).records());
    }

    @Test
    void canonicalizesLegacyFullwidthParenthesesToSameSemanticFenceChars() {
        assertEquals(
            List.of(character(0x82, '('), character(0x82, ')')),
            MtefRecordNormalizer.normalizeLegacyCompatible(new byte[] {
                5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
                (byte) MtefRecord.LINE, 0,
                (byte) MtefRecord.CHAR, 0, (byte) 0x8C, 0x08, (byte) 0xFF,
                (byte) MtefRecord.CHAR, 0, (byte) 0x8C, 0x09, (byte) 0xFF,
                (byte) MtefRecord.END
            }).records());
    }

    @Test
    void canonicalizesLegacyFarEastAsciiMathTokensBySemanticRole() {
        assertEquals(
            List.of(character(0x88, '6'), character(0x86, '+'), character(0x86, '='),
                character(0x86, '\u2212')),
            MtefRecordNormalizer.normalizeLegacyCompatible(new byte[] {
                5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
                (byte) MtefRecord.LINE, 0,
                (byte) MtefRecord.CHAR, 0, (byte) 0x8C, '6', 0,
                (byte) MtefRecord.CHAR, 0, (byte) 0x8C, '+', 0,
                (byte) MtefRecord.CHAR, 0, (byte) 0x8C, '=', 0,
                (byte) MtefRecord.CHAR, 0, (byte) 0x8C, '-', 0,
                (byte) MtefRecord.END
            }).records());
    }

    @Test
    void legacyStructureIgnoresCharacterNudgePayload() {
        assertEquals(
            "CHAR:134:61:null",
            MtefRecordNormalizer.normalizeLegacyCompatible(new byte[] {
                5, 1, 0, 6, 0, 'D', 'S', 'M', 'T', '6', 0, 0,
                (byte) MtefRecord.LINE, 0,
                (byte) MtefRecord.CHAR, (byte) MtefRecord.OPT_NUDGE,
                61, 0, (byte) 0x86, '=', 0,
                (byte) MtefRecord.END
            }).canonicalSignature());
    }

    @Test
    void unwrapsSingleColumnMatrixToLegacyRootLineLayout() {
        var matrix = new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.MATRIX, "MATRIX", 0, null, null, null, null, null,
            2, 1, 0, 0, 0, "", "", null);
        var line = record(MtefRecord.LINE, "LINE", false);
        var end = record(MtefRecord.END, "END", null);

        assertEquals(
            List.of(character(0x88, '1'), line, character(0x88, '2'), end, end, end),
            MtefRecordNormalizer.canonicalizeLegacySingleColumnMatrix(List.of(
                matrix, line, character(0x88, '1'), end,
                line, character(0x88, '2'), end, end, end)));
    }

    private static MtefRecordNormalizer.CanonicalRecord character(int typeface, char value) {
        return new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.CHAR, "CHAR", 0, null, null, typeface, (int) value,
            null, null, null, null);
    }

    private static MtefRecordNormalizer.CanonicalRecord color(int value) {
        return new MtefRecordNormalizer.CanonicalRecord(
            MtefRecord.COLOR, "COLOR", null, null, null, null, null, value, null, null, null);
    }

    private static MtefRecordNormalizer.CanonicalRecord record(int tag, String name, Boolean nullLine) {
        return new MtefRecordNormalizer.CanonicalRecord(
            tag, name, tag == MtefRecord.LINE ? 0 : null, null, null,
            null, null, null, null, null, nullLine);
    }
}
