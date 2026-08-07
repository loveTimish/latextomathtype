package com.lz.paperword.core.mtef;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

/** Converts an MTEF stream into stable formula-content records for golden comparisons. */
public final class MtefRecordNormalizer {

    private static final Map<Integer, String> TAG_NAMES = Map.ofEntries(
        Map.entry(0x00, "END"), Map.entry(0x01, "LINE"), Map.entry(0x02, "CHAR"),
        Map.entry(0x03, "TMPL"), Map.entry(0x04, "PILE"), Map.entry(0x05, "MATRIX"),
        Map.entry(0x06, "EMBELL"), Map.entry(0x07, "RULER"),
        Map.entry(0x08, "FONT_STYLE_DEF"), Map.entry(0x09, "SIZE"),
        Map.entry(0x0A, "FULL"), Map.entry(0x0B, "SUB"), Map.entry(0x0C, "SUB2"),
        Map.entry(0x0D, "SYM"), Map.entry(0x0E, "SUBSYM"), Map.entry(0x0F, "COLOR"),
        Map.entry(0x10, "COLOR_DEF"), Map.entry(0x11, "FONT_DEF"),
        Map.entry(0x12, "EQN_PREFS"), Map.entry(0x13, "ENCODING_DEF")
    );

    private MtefRecordNormalizer() {
    }

    public static NormalizationReport normalize(byte[] mtef) {
        if (mtef == null || mtef.length == 0) {
            return new NormalizationReport(0, 0, List.of(), Map.of(), "");
        }
        int contentOffset = formulaContentOffset(mtef);
        List<CanonicalRecord> records = new ArrayList<>();
        Map<String, Integer> counts = new LinkedHashMap<>();
        int offset = contentOffset;
        while (offset < mtef.length) {
            ParsedRecord parsed = parseRecord(mtef, offset);
            records.add(parsed.record());
            counts.merge(parsed.record().name(), 1, Integer::sum);
            if (parsed.nextOffset() <= offset) {
                break;
            }
            offset = parsed.nextOffset();
        }
        records = canonicalFormulaRecords(records);
        StringJoiner signature = new StringJoiner("|");
        records.stream().map(CanonicalRecord::signature).forEach(signature::add);
        return new NormalizationReport(
            contentOffset,
            mtef.length - contentOffset,
            records,
            counts,
            signature.toString()
        );
    }

    /** Canonicalizes MathType 6 flat ASCII parentheses against MathType 7 fence templates. */
    public static NormalizationReport normalizeLegacyCompatible(byte[] mtef) {
        NormalizationReport strict = normalize(mtef);
        List<CanonicalRecord> records = canonicalizeLegacyTypefaceRecords(
            canonicalizeLegacySingleColumnMatrix(
                canonicalizeLegacyFenceRecords(strict.records())));
        StringJoiner signature = new StringJoiner("|");
        records.stream().map(CanonicalRecord::signature).forEach(signature::add);
        return new NormalizationReport(
            strict.contentOffset(), strict.contentByteLength(), records,
            strict.recordCounts(), signature.toString());
    }

    static List<CanonicalRecord> canonicalizeLegacySingleColumnMatrix(List<CanonicalRecord> source) {
        if (source.size() < 4 || source.get(0).tag() != MtefRecord.MATRIX
                || !Integer.valueOf(1).equals(source.get(0).columns())
                || source.get(1).tag() != MtefRecord.LINE) {
            return source;
        }
        int matrixEnd = matchingEnd(source, 0);
        int firstLineEnd = matchingEnd(source, 1);
        if (matrixEnd <= firstLineEnd || firstLineEnd < 0) {
            return source;
        }
        List<CanonicalRecord> records = new ArrayList<>(source.size() - 2);
        records.addAll(source.subList(2, firstLineEnd));
        records.addAll(source.subList(firstLineEnd + 1, matrixEnd + 1));
        records.addAll(source.subList(matrixEnd + 1, source.size()));
        return List.copyOf(records);
    }

    private static List<CanonicalRecord> canonicalizeLegacyTypefaceRecords(List<CanonicalRecord> source) {
        return source.stream().map(record -> {
            if (record.tag() == MtefRecord.CHAR && isLegacyFlatParenthesis(record)) {
                int ascii = Integer.valueOf(0xFF08).equals(record.mtcode()) ? '(' :
                    Integer.valueOf(0xFF09).equals(record.mtcode()) ? ')' : record.mtcode();
                return new CanonicalRecord(record.tag(), record.name(), record.options(),
                    record.selector(), record.variation(), 0x82, ascii, record.value(),
                    record.rows(), record.columns(), record.matrixVerticalAlignment(),
                    record.matrixHorizontalAlignment(), record.matrixVerticalJustification(),
                    record.rowPartitions(), record.columnPartitions(), record.nullLine());
            }
            if (record.tag() == MtefRecord.CHAR && Integer.valueOf(0x8C).equals(record.typeface())
                    && record.mtcode() != null) {
                int mtcode = record.mtcode();
                if (mtcode >= '0' && mtcode <= '9') {
                    return withTypefaceAndMtcode(record, 0x88, mtcode);
                }
                if (mtcode == '+' || mtcode == '=') {
                    return withTypefaceAndMtcode(record, 0x86, mtcode);
                }
                if (mtcode == '-') {
                    return withTypefaceAndMtcode(record, 0x86, 0x2212);
                }
            }
            if (record.tag() == MtefRecord.CHAR && record.value() != null) {
                return new CanonicalRecord(record.tag(), record.name(), record.options(),
                    record.selector(), record.variation(), record.typeface(), record.mtcode(), null,
                    record.rows(), record.columns(), record.matrixVerticalAlignment(),
                    record.matrixHorizontalAlignment(), record.matrixVerticalJustification(),
                    record.rowPartitions(), record.columnPartitions(), record.nullLine());
            }
            return record;
        }).toList();
    }

    private static CanonicalRecord withTypefaceAndMtcode(CanonicalRecord record, int typeface, int mtcode) {
        return new CanonicalRecord(record.tag(), record.name(), record.options(),
            record.selector(), record.variation(), typeface, mtcode, record.value(),
            record.rows(), record.columns(), record.matrixVerticalAlignment(),
            record.matrixHorizontalAlignment(), record.matrixVerticalJustification(),
            record.rowPartitions(), record.columnPartitions(), record.nullLine());
    }

    private static boolean isLegacyFlatParenthesis(CanonicalRecord record) {
        if (Integer.valueOf(0xFF08).equals(record.mtcode())
                || Integer.valueOf(0xFF09).equals(record.mtcode())) {
            return true;
        }
        return (Integer.valueOf((int) '(').equals(record.mtcode())
                || Integer.valueOf((int) ')').equals(record.mtcode()))
            && record.typeface() != null
            && (record.typeface() == 0x7E || record.typeface() == 0x7F || record.typeface() == 0x82);
    }

    static List<CanonicalRecord> canonicalizeLegacyFenceRecords(List<CanonicalRecord> source) {
        List<CanonicalRecord> records = new ArrayList<>();
        int index = 0;
        while (index < source.size()) {
            CanonicalRecord record = source.get(index);
            if (record.tag() == MtefRecord.TMPL
                    && Integer.valueOf(MtefRecord.TM_PAREN).equals(record.selector())) {
                int templateEnd = matchingEnd(source, index);
                int lineIndex = index + 1;
                if (templateEnd > lineIndex && source.get(lineIndex).tag() == MtefRecord.LINE
                        && !Boolean.TRUE.equals(source.get(lineIndex).nullLine())) {
                    int lineEnd = matchingEnd(source, lineIndex);
                    List<CanonicalRecord> delimiters = lineEnd > lineIndex
                        ? source.subList(lineEnd + 1, templateEnd).stream()
                            .filter(candidate -> candidate.tag() == MtefRecord.CHAR)
                            .toList()
                        : List.of();
                    if (delimiters.size() >= 2 && isAsciiParenthesis(delimiters.get(0), '(')
                            && isAsciiParenthesis(delimiters.get(1), ')')) {
                        records.add(legacyParenthesis('('));
                        records.addAll(canonicalizeLegacyFenceRecords(source.subList(lineIndex + 1, lineEnd)));
                        records.add(legacyParenthesis(')'));
                        index = templateEnd + 1;
                        continue;
                    }
                }
            }
            records.add(record);
            index++;
        }
        return List.copyOf(records);
    }

    private static boolean isAsciiParenthesis(CanonicalRecord record, char value) {
        return record.mtcode() != null && record.mtcode() == value;
    }

    private static CanonicalRecord legacyParenthesis(char value) {
        return new CanonicalRecord(
            MtefRecord.CHAR, "CHAR", 0, null, null, 0x82, (int) value,
            null, null, null, null);
    }

    static List<CanonicalRecord> canonicalFormulaRecords(List<CanonicalRecord> parsedRecords) {
        Map<Integer, Integer> colors = new LinkedHashMap<>();
        colors.put(0, 0);
        colors.put(1, 0);
        int nextColorDefinition = 1;
        int currentColor = 0;
        List<CanonicalRecord> records = new ArrayList<>();
        for (CanonicalRecord record : parsedRecords) {
            if (record.tag() == MtefRecord.COLOR_DEF) {
                colors.put(nextColorDefinition++, record.value() == null ? 0 : record.value());
                continue;
            }
            if (record.tag() == MtefRecord.COLOR) {
                int index = record.value() == null ? 0 : record.value();
                nextColorDefinition = Math.max(nextColorDefinition, index + 1);
                int color = colors.getOrDefault(index, index == 0 || index == 1 ? 0 : index);
                if (color != currentColor) {
                    records.add(new CanonicalRecord(record.tag(), record.name(), record.options(),
                        record.selector(), record.variation(), record.typeface(), record.mtcode(), color,
                        record.rows(), record.columns(), record.nullLine()));
                    currentColor = color;
                }
                continue;
            }
            if (record.tag() == MtefRecord.ENCODING_DEF || record.tag() == MtefRecord.FONT_DEF
                    || record.tag() == MtefRecord.FONT_STYLE_DEF) {
                continue;
            }
            records.add(record);
        }
        if (!records.isEmpty() && records.get(0).tag() == MtefRecord.FULL) {
            records.remove(0);
        }
        if (!records.isEmpty() && records.get(0).tag() == MtefRecord.LINE
                && !Boolean.TRUE.equals(records.get(0).nullLine())) {
            int matchingEnd = matchingEnd(records, 0);
            if (matchingEnd > 0) {
                records.remove(matchingEnd);
                records.remove(0);
            }
        }
        return records;
    }

    private static int matchingEnd(List<CanonicalRecord> records, int start) {
        int depth = 0;
        for (int index = start; index < records.size(); index++) {
            int tag = records.get(index).tag();
            if ((tag == MtefRecord.LINE && !Boolean.TRUE.equals(records.get(index).nullLine()))
                    || tag == MtefRecord.TMPL || tag == MtefRecord.PILE
                    || tag == MtefRecord.MATRIX || tag == MtefRecord.EMBELL) {
                depth++;
            } else if (tag == MtefRecord.END && --depth == 0) {
                return index;
            }
        }
        return -1;
    }

    private static int formulaContentOffset(byte[] data) {
        int offset = recordStart(data);
        while (offset < data.length) {
            if (isLineRecord(data, offset)) {
                return offset;
            }
            if (unsigned(data[offset]) == MtefRecord.FULL
                    && offset + 1 < data.length
                    && unsigned(data[offset + 1]) == MtefRecord.END) {
                return offset;
            }
            int next = parseRecord(data, offset).nextOffset();
            if (next <= offset) {
                break;
            }
            offset = next;
        }
        return Math.min(recordStart(data), data.length);
    }

    private static int recordStart(byte[] data) {
        if (data.length <= 5) {
            return 0;
        }
        if (data.length >= 10 && new String(data, 5, 4, StandardCharsets.US_ASCII).equals("DSMT")) {
            int markerEnd = indexOfZero(data, 10);
            if (markerEnd >= 0) {
                int cursor = markerEnd + 1;
                if (data[9] == '7') {
                    byte[] texInput = "TeX Input Language\0".getBytes(StandardCharsets.US_ASCII);
                    int label = indexOf(data, texInput, cursor);
                    if (label >= 0 && label - cursor <= 8) {
                        int sourceEnd = indexOfZero(data, label + texInput.length);
                        if (sourceEnd >= 0) {
                            return Math.min(data.length, sourceEnd + 1);
                        }
                    }
                }
                return Math.min(data.length, cursor + 1);
            }
        }
        int markerEnd = indexOfZero(data, 5);
        return markerEnd >= 0 ? Math.min(data.length, markerEnd + 1) : 5;
    }

    private static boolean isFormulaLine(byte[] data, int offset) {
        return isLineRecord(data, offset)
            && lineHasMeaningfulTail(data, offset);
    }

    private static boolean isLineRecord(byte[] data, int offset) {
        return offset + 1 < data.length
            && unsigned(data[offset]) == MtefRecord.LINE
            && (unsigned(data[offset + 1]) & ~0x0F) == 0;
    }

    private static boolean lineHasMeaningfulTail(byte[] data, int start) {
        int offset = start + 2;
        while (offset < data.length) {
            int tag = unsigned(data[offset]);
            if (tag == MtefRecord.CHAR || tag == MtefRecord.TMPL
                || tag == MtefRecord.PILE || tag == MtefRecord.MATRIX) {
                return true;
            }
            if (tag == MtefRecord.END) {
                return false;
            }
            int next = parseRecord(data, offset).nextOffset();
            if (next <= offset) {
                return false;
            }
            offset = next;
        }
        return false;
    }

    private static ParsedRecord parseRecord(byte[] data, int offset) {
        int tag = unsigned(data[offset]);
        int cursor = offset + 1;
        Integer options = null;
        Integer selector = null;
        Integer variation = null;
        Integer typeface = null;
        Integer mtcode = null;
        Integer value = null;
        Integer rows = null;
        Integer columns = null;
        Integer matrixVerticalAlignment = null;
        Integer matrixHorizontalAlignment = null;
        Integer matrixVerticalJustification = null;
        String rowPartitions = null;
        String columnPartitions = null;
        Boolean nullLine = null;

        if (tag == MtefRecord.LINE && cursor < data.length) {
            options = unsigned(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor += nudgeLength(data, cursor);
            }
            if ((options & MtefRecord.OPT_LINE_LSPACE) != 0 && cursor < data.length) {
                value = unsigned(data[cursor++]);
            }
            if ((options & MtefRecord.OPT_LP_RULER) != 0 && cursor < data.length) {
                int stops = unsigned(data[cursor++]);
                cursor += stops * 3;
            }
            nullLine = (options & MtefRecord.OPT_LINE_NULL) != 0;
        } else if (tag == MtefRecord.CHAR && cursor < data.length) {
            options = unsigned(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor += nudgeLength(data, cursor);
            }
            if ((options & MtefRecord.OPT_CHAR_ENC_NO_MTCODE) == 0 && cursor + 2 < data.length) {
                typeface = unsigned(data[cursor++]);
                mtcode = littleEndian16(data, cursor);
                cursor += 2;
            }
            if ((options & MtefRecord.OPT_CHAR_ENC_CHAR_8) != 0 && cursor < data.length) {
                value = unsigned(data[cursor++]);
            }
            if ((options & MtefRecord.OPT_CHAR_ENC_CHAR_16) != 0 && cursor + 1 < data.length) {
                cursor += 2;
            }
        } else if (tag == MtefRecord.TMPL && cursor < data.length) {
            options = unsigned(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor += nudgeLength(data, cursor);
            }
            if (cursor < data.length) {
                selector = unsigned(data[cursor++]);
            }
            if (cursor < data.length) {
                variation = unsigned(data[cursor++]);
                if ((variation & 0x80) != 0 && cursor < data.length) {
                    variation |= unsigned(data[cursor++]) << 8;
                }
            }
            if (cursor < data.length) {
                value = unsigned(data[cursor++]);
            }
        } else if (tag == MtefRecord.PILE && cursor < data.length) {
            options = unsigned(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor += nudgeLength(data, cursor);
            }
            cursor += 2;
        } else if (tag == MtefRecord.MATRIX && cursor < data.length) {
            options = unsigned(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor += nudgeLength(data, cursor);
            }
            if (cursor < data.length) matrixVerticalAlignment = unsigned(data[cursor++]);
            if (cursor < data.length) matrixHorizontalAlignment = unsigned(data[cursor++]);
            if (cursor < data.length) matrixVerticalJustification = unsigned(data[cursor++]);
            if (cursor + 1 < data.length) {
                rows = unsigned(data[cursor++]);
                columns = unsigned(data[cursor++]);
                int rowBytes = packedPartitionBytes(rows + 1);
                int columnBytes = packedPartitionBytes(columns + 1);
                rowPartitions = hexSlice(data, cursor, rowBytes);
                cursor += rowBytes;
                columnPartitions = hexSlice(data, cursor, columnBytes);
                cursor += columnBytes;
            }
        } else if (tag == MtefRecord.RULER && cursor < data.length) {
            int stops = unsigned(data[cursor++]);
            cursor += stops * 3;
        } else if (tag == MtefRecord.FONT_STYLE_DEF || tag == MtefRecord.FONT_DEF
            || tag == MtefRecord.ENCODING_DEF) {
            cursor = skipNullTerminated(data, Math.min(cursor + 1, data.length));
        } else if (tag == MtefRecord.SIZE && cursor < data.length) {
            value = unsigned(data[cursor++]);
            if (value == 101 && cursor + 1 < data.length) {
                value = littleEndian16(data, cursor);
                cursor += 2;
            } else if (value == 100 && cursor + 2 < data.length) {
                cursor += 3;
            } else if (cursor < data.length && unsigned(data[cursor]) == 0x50) {
                cursor += Math.min(2, data.length - cursor);
            }
        } else if (tag == MtefRecord.EMBELL || tag == MtefRecord.COLOR) {
            if (cursor < data.length) {
                value = unsigned(data[cursor++]);
            }
        } else if (tag == MtefRecord.COLOR_DEF && cursor < data.length) {
            options = unsigned(data[cursor++]);
            int componentCount = (options & 0x01) != 0 ? 4 : 3;
            int packedColor = 0;
            for (int component = 0; component < componentCount && cursor + 1 < data.length; component++) {
                int componentValue = littleEndian16(data, cursor);
                cursor += 2;
                packedColor = componentCount == 3
                    ? (packedColor << 10) | (componentValue & 0x3FF)
                    : 31 * packedColor + componentValue;
            }
            value = packedColor;
            if ((options & 0x04) != 0) {
                cursor = skipNullTerminated(data, cursor);
            }
        } else if (tag == MtefRecord.EQN_PREFS) {
            cursor = skipEquationPreferences(data, cursor);
        } else if (tag >= MtefRecord.FUTURE && cursor < data.length) {
            int length = unsigned(data[cursor++]);
            cursor += length;
        }

        cursor = Math.max(offset + 1, Math.min(cursor, data.length));
        CanonicalRecord record = new CanonicalRecord(
            tag,
            TAG_NAMES.getOrDefault(tag, "REC_" + tag),
            options,
            selector,
            variation,
            typeface,
            mtcode,
            value,
            rows,
            columns,
            matrixVerticalAlignment,
            matrixHorizontalAlignment,
            matrixVerticalJustification,
            rowPartitions,
            columnPartitions,
            nullLine
        );
        return new ParsedRecord(record, cursor);
    }

    private static int skipEquationPreferences(byte[] data, int offset) {
        int cursor = Math.min(offset + 1, data.length); // options
        cursor = skipDimensionArray(data, cursor);     // sizes
        cursor = skipDimensionArray(data, cursor);     // spacing
        if (cursor >= data.length) {
            return data.length;
        }
        int styleCount = unsigned(data[cursor++]);
        for (int index = 0; index < styleCount && cursor < data.length; index++) {
            int fontDefinition = unsigned(data[cursor++]);
            if (fontDefinition != 0 && cursor < data.length) {
                cursor++; // character style
            }
        }
        return cursor;
    }

    private static int skipDimensionArray(byte[] data, int offset) {
        if (offset >= data.length) {
            return data.length;
        }
        int dimensionCount = unsigned(data[offset]);
        int byteOffset = offset + 1;
        int nibbleOffset = 0;
        int completedDimensions = 0;
        while (byteOffset + nibbleOffset / 2 < data.length
                && completedDimensions < dimensionCount) {
            int packed = unsigned(data[byteOffset + nibbleOffset / 2]);
            int nibble = (nibbleOffset & 1) == 0 ? packed >>> 4 : packed & 0x0F;
            nibbleOffset++;
            if (nibble == 0x0F) {
                completedDimensions++;
            }
        }
        return Math.min(data.length, byteOffset + (nibbleOffset + 1) / 2);
    }

    private static int nudgeLength(byte[] data, int offset) {
        if (offset + 1 >= data.length) {
            return 0;
        }
        return unsigned(data[offset]) == 0x80 || unsigned(data[offset + 1]) == 0x80 ? 6 : 2;
    }

    private static int packedPartitionBytes(int size) {
        return size <= 0 ? 0 : (size + 3) / 4;
    }

    private static String hexSlice(byte[] data, int offset, int length) {
        StringBuilder hex = new StringBuilder(length * 2);
        for (int index = 0; index < length && offset + index < data.length; index++) {
            hex.append(String.format("%02x", unsigned(data[offset + index])));
        }
        return hex.toString();
    }

    private static int skipNullTerminated(byte[] data, int offset) {
        int cursor = Math.max(0, offset);
        while (cursor < data.length && data[cursor] != 0) {
            cursor++;
        }
        return cursor < data.length ? cursor + 1 : cursor;
    }

    private static int indexOfZero(byte[] data, int offset) {
        for (int index = Math.max(0, offset); index < data.length; index++) {
            if (data[index] == 0) {
                return index;
            }
        }
        return -1;
    }

    private static int indexOf(byte[] data, byte[] needle, int offset) {
        outer:
        for (int index = Math.max(0, offset); index + needle.length <= data.length; index++) {
            for (int part = 0; part < needle.length; part++) {
                if (data[index + part] != needle[part]) {
                    continue outer;
                }
            }
            return index;
        }
        return -1;
    }

    private static int littleEndian16(byte[] data, int offset) {
        return unsigned(data[offset]) | (unsigned(data[offset + 1]) << 8);
    }

    private static int unsigned(byte value) {
        return value & 0xFF;
    }

    public record CanonicalRecord(
        int tag,
        String name,
        Integer options,
        Integer selector,
        Integer variation,
        Integer typeface,
        Integer mtcode,
        Integer value,
        Integer rows,
        Integer columns,
        Integer matrixVerticalAlignment,
        Integer matrixHorizontalAlignment,
        Integer matrixVerticalJustification,
        String rowPartitions,
        String columnPartitions,
        Boolean nullLine
    ) {
        public CanonicalRecord(int tag, String name, Integer options, Integer selector,
                               Integer variation, Integer typeface, Integer mtcode, Integer value,
                               Integer rows, Integer columns, Boolean nullLine) {
            this(tag, name, options, selector, variation, typeface, mtcode, value, rows, columns,
                null, null, null, null, null, nullLine);
        }

        public String signature() {
            if (tag == MtefRecord.CHAR) {
                return "CHAR:" + typeface + ':' + mtcode + ':' + value;
            }
            if (tag == MtefRecord.TMPL) {
                return "TMPL:" + selector + ':' + variation;
            }
            if (tag == MtefRecord.LINE) {
                return "LINE:null=" + nullLine;
            }
            if (tag == MtefRecord.MATRIX) {
                return "MATRIX:" + rows + 'x' + columns + ":align="
                    + matrixVerticalAlignment + ',' + matrixHorizontalAlignment + ','
                    + matrixVerticalJustification + ":parts=" + rowPartitions + ',' + columnPartitions;
            }
            if (tag == MtefRecord.SIZE || tag == MtefRecord.EMBELL || tag == MtefRecord.COLOR) {
                return name + ':' + value;
            }
            return name;
        }
    }

    public record NormalizationReport(
        int contentOffset,
        int contentByteLength,
        List<CanonicalRecord> records,
        Map<String, Integer> recordCounts,
        String canonicalSignature
    ) {
        public NormalizationReport {
            records = List.copyOf(records);
            recordCounts = Map.copyOf(recordCounts);
        }
    }

    private record ParsedRecord(CanonicalRecord record, int nextOffset) {
    }
}
