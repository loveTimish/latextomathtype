package com.lz.paperword.core.mtef;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HexFormat;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/** Exact normalized-MTEF comparison for MathType OLE objects embedded in DOCX files. */
public final class MathTypeDocxComparator {

    public InspectionReport inspect(Path docx) throws IOException {
        List<InspectedFormula> formulas = readFormulas(docx).stream()
            .map(formula -> new InspectedFormula(
                formula.entry(), formula.valid(), formula.error(), formula.sha256(),
                HexFormat.of().formatHex(formula.mtef()),
                formula.normalization().records()))
            .toList();
        return new InspectionReport(docx.toAbsolutePath().toString(), formulas);
    }

    public ComparisonReport compare(Path standardDocx, Path generatedDocx) throws IOException {
        List<OleFormula> standard = readFormulas(standardDocx);
        List<OleFormula> generated = readFormulas(generatedDocx);
        List<FormulaComparison> formulas = new ArrayList<>();
        int count = Math.max(standard.size(), generated.size());
        int validOle = 0;
        int structureEqual = 0;
        int rawMtefEqual = 0;
        for (int index = 0; index < count; index++) {
            OleFormula expected = index < standard.size() ? standard.get(index) : null;
            OleFormula actual = index < generated.size() ? generated.get(index) : null;
            boolean oleValid = actual != null && actual.valid();
            boolean strictStructure = expected != null && actual != null && expected.valid() && actual.valid()
                && expected.normalization().canonicalSignature()
                    .equals(actual.normalization().canonicalSignature());
            boolean legacyCandidate = !strictStructure && expected != null && actual != null
                && expected.valid() && actual.valid() && isLegacyDsmt(expected.mtef());
            MtefRecordNormalizer.NormalizationReport expectedReport = legacyCandidate
                ? MtefRecordNormalizer.normalizeLegacyCompatible(expected.mtef())
                : expected == null ? null : expected.normalization();
            MtefRecordNormalizer.NormalizationReport actualReport = legacyCandidate
                ? MtefRecordNormalizer.normalizeLegacyCompatible(actual.mtef())
                : actual == null ? null : actual.normalization();
            boolean legacyStructure = legacyCandidate
                && expectedReport.canonicalSignature().equals(actualReport.canonicalSignature());
            boolean sameStructure = strictStructure || legacyStructure;
            boolean sameBytes = expected != null && actual != null
                && MessageDigest.isEqual(expected.mtef(), actual.mtef());
            int firstDifference = sameStructure ? -1 : firstDifference(expectedReport, actualReport);
            if (oleValid) validOle++;
            if (sameStructure) structureEqual++;
            if (sameBytes) rawMtefEqual++;
            formulas.add(new FormulaComparison(
                index,
                expected == null ? null : expected.entry(),
                actual == null ? null : actual.entry(),
                oleValid,
                sameStructure,
                legacyStructure,
                sameBytes,
                expected == null ? "missing standard object" : expected.error(),
                actual == null ? "missing generated object" : actual.error(),
                expected == null ? null : expected.sha256(),
                actual == null ? null : actual.sha256(),
                firstDifference,
                recordAt(expectedReport, firstDifference),
                recordAt(actualReport, firstDifference),
                sameStructure || expectedReport == null ? List.of() : expectedReport.records(),
                sameStructure || actualReport == null ? List.of() : actualReport.records(),
                sameStructure || expected == null ? "" : HexFormat.of().formatHex(expected.mtef()),
                sameStructure || actual == null ? "" : HexFormat.of().formatHex(actual.mtef()),
                literalLatexCommands(expected),
                literalLatexCommands(actual)
            ));
        }
        return new ComparisonReport(
            standardDocx.toAbsolutePath().toString(),
            generatedDocx.toAbsolutePath().toString(),
            standard.size(),
            generated.size(),
            validOle,
            structureEqual,
            rawMtefEqual,
            formulas
        );
    }

    private boolean isLegacyDsmt(byte[] mtef) {
        return mtef != null && mtef.length > 9 && mtef[5] == 'D' && mtef[6] == 'S'
            && mtef[7] == 'M' && mtef[8] == 'T'
            && (mtef[9] == '5' || mtef[9] == '6');
    }

    private int firstDifference(MtefRecordNormalizer.NormalizationReport expected,
                                MtefRecordNormalizer.NormalizationReport actual) {
        if (expected == null || actual == null) {
            return -1;
        }
        List<MtefRecordNormalizer.CanonicalRecord> left = expected.records();
        List<MtefRecordNormalizer.CanonicalRecord> right = actual.records();
        int common = Math.min(left.size(), right.size());
        for (int index = 0; index < common; index++) {
            if (!left.get(index).signature().equals(right.get(index).signature())) {
                return index;
            }
        }
        return left.size() == right.size() ? -1 : common;
    }

    private MtefRecordNormalizer.CanonicalRecord recordAt(
            MtefRecordNormalizer.NormalizationReport report, int index) {
        if (report == null || index < 0 || index >= report.records().size()) {
            return null;
        }
        return report.records().get(index);
    }

    private List<String> literalLatexCommands(OleFormula formula) {
        if (formula == null || !formula.valid()) {
            return List.of();
        }
        List<String> commands = new ArrayList<>();
        List<MtefRecordNormalizer.CanonicalRecord> records = formula.normalization().records();
        for (int index = 0; index < records.size(); index++) {
            MtefRecordNormalizer.CanonicalRecord record = records.get(index);
            if (!isLiteralCharacter(record, '\\')) {
                continue;
            }
            StringBuilder command = new StringBuilder("\\");
            int cursor = index + 1;
            while (cursor < records.size()) {
                Integer mtcode = records.get(cursor).mtcode();
                if (mtcode == null || mtcode > 0x7F || !Character.isLetter((char) mtcode.intValue())) {
                    break;
                }
                command.append((char) mtcode.intValue());
                cursor++;
            }
            if (command.length() > 1) {
                commands.add(command.toString());
                index = cursor - 1;
            }
        }
        return List.copyOf(commands);
    }

    private boolean isLiteralCharacter(MtefRecordNormalizer.CanonicalRecord record, char expected) {
        return "CHAR".equals(record.name())
            && record.mtcode() != null
            && record.mtcode() == expected;
    }

    private List<OleFormula> readFormulas(Path docx) throws IOException {
        List<OleFormula> formulas = new ArrayList<>();
        try (ZipFile zip = new ZipFile(docx.toFile())) {
            List<? extends ZipEntry> embeddings = zip.stream()
                .filter(entry -> !entry.isDirectory())
                .filter(entry -> entry.getName().startsWith("word/embeddings/")
                    && entry.getName().toLowerCase().endsWith(".bin"))
                .sorted(Comparator.comparingInt(entry -> numericSuffix(entry.getName())))
                .toList();
            for (ZipEntry entry : embeddings) {
                formulas.add(readOle(entry.getName(), zip.getInputStream(entry).readAllBytes()));
            }
        }
        return formulas;
    }

    private OleFormula readOle(String entryName, byte[] oleBytes) {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(oleBytes))) {
            DirectoryEntry root = fs.getRoot();
            if (!root.hasEntry("Equation Native")) {
                return OleFormula.invalid(entryName, "Equation Native stream missing");
            }
            boolean progIdValid = containsEquationDsmt4(root);
            byte[] nativeBytes = readDocument(root, "Equation Native");
            if (nativeBytes.length < 28) {
                return OleFormula.invalid(entryName, "Equation Native shorter than 28-byte header");
            }
            ByteBuffer header = ByteBuffer.wrap(nativeBytes).order(ByteOrder.LITTLE_ENDIAN);
            int headerSize = Short.toUnsignedInt(header.getShort(0));
            long declaredLength = Integer.toUnsignedLong(header.getInt(8));
            if (headerSize < 12 || headerSize > nativeBytes.length) {
                return OleFormula.invalid(entryName, "invalid Equation Native header size " + headerSize);
            }
            if (declaredLength != nativeBytes.length - headerSize) {
                return OleFormula.invalid(entryName, "Equation Native cbObject mismatch");
            }
            byte[] mtef = java.util.Arrays.copyOfRange(nativeBytes, headerSize, nativeBytes.length);
            boolean dsmtHeader = mtef.length >= 10
                && new String(mtef, 5, 5, StandardCharsets.US_ASCII).startsWith("DSMT");
            if (!progIdValid || !dsmtHeader) {
                return OleFormula.invalid(entryName,
                    !progIdValid ? "CompObj does not identify Equation.DSMT4" : "MTEF DSMT header missing");
            }
            return new OleFormula(entryName, true, "", mtef, sha256(mtef),
                MtefRecordNormalizer.normalize(mtef));
        } catch (Exception exception) {
            return OleFormula.invalid(entryName, exception.toString());
        }
    }

    private boolean containsEquationDsmt4(DirectoryEntry root) throws IOException {
        for (Entry entry : root) {
            if (entry instanceof DocumentEntry document && entry.getName().contains("CompObj")) {
                byte[] bytes = readDocument(root, entry.getName());
                String ascii = new String(bytes, StandardCharsets.ISO_8859_1);
                String utf16 = new String(bytes, StandardCharsets.UTF_16LE);
                return ascii.contains("Equation.DSMT4") || utf16.contains("Equation.DSMT4");
            }
        }
        return false;
    }

    private byte[] readDocument(DirectoryEntry root, String name) throws IOException {
        try (DocumentInputStream input = new DocumentInputStream((DocumentEntry) root.getEntry(name));
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            input.transferTo(output);
            return output.toByteArray();
        }
    }

    private int numericSuffix(String name) {
        String digits = name.replaceAll("\\D+", "");
        return digits.isEmpty() ? Integer.MAX_VALUE : Integer.parseInt(digits);
    }

    private String sha256(byte[] bytes) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes));
    }

    public record FormulaComparison(
        int index,
        String standardEntry,
        String generatedEntry,
        boolean oleValid,
        boolean normalizedStructureEqual,
        boolean legacyNormalizationApplied,
        boolean rawMtefEqual,
        String standardError,
        String generatedError,
        String standardMtefSha256,
        String generatedMtefSha256,
        int firstNormalizedDifferenceIndex,
        MtefRecordNormalizer.CanonicalRecord standardDifference,
        MtefRecordNormalizer.CanonicalRecord generatedDifference,
        List<MtefRecordNormalizer.CanonicalRecord> standardNormalizedRecords,
        List<MtefRecordNormalizer.CanonicalRecord> generatedNormalizedRecords,
        String standardMtefHex,
        String generatedMtefHex,
        List<String> standardLiteralLatexCommands,
        List<String> generatedLiteralLatexCommands
    ) {}

    public record ComparisonReport(
        String standardDocx,
        String generatedDocx,
        int standardObjectCount,
        int generatedObjectCount,
        int validGeneratedOleCount,
        int normalizedStructureEqualCount,
        int rawMtefEqualCount,
        List<FormulaComparison> formulas
    ) {
        public boolean passesExactStructureGate() {
            return standardObjectCount == generatedObjectCount
                && generatedObjectCount == validGeneratedOleCount
                && generatedObjectCount == normalizedStructureEqualCount;
        }
    }

    public record InspectedFormula(
        String entry,
        boolean valid,
        String error,
        String mtefSha256,
        String mtefHex,
        List<MtefRecordNormalizer.CanonicalRecord> records
    ) {}

    public record InspectionReport(String docx, List<InspectedFormula> formulas) {}

    private record OleFormula(
        String entry,
        boolean valid,
        String error,
        byte[] mtef,
        String sha256,
        MtefRecordNormalizer.NormalizationReport normalization
    ) {
        private static OleFormula invalid(String entry, String error) {
            return new OleFormula(entry, false, error, new byte[0], "", MtefRecordNormalizer.normalize(new byte[0]));
        }
    }
}
