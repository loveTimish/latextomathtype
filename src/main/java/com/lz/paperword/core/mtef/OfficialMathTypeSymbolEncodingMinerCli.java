package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Mines positionally unambiguous CHAR encodings from the pinned MathType standard DOCX files. */
public final class OfficialMathTypeSymbolEncodingMinerCli {

    private static final Pattern TOKEN = Pattern.compile("\\\\[A-Za-z]+|\\\\[^A-Za-z\\s]|[^\\s{}]");
    private static final Pattern COMMAND = Pattern.compile("\\\\[A-Za-z]+");
    private static final Set<String> STRUCTURAL_RECORDS = Set.of("TMPL", "PILE", "MATRIX", "EMBELL");

    private OfficialMathTypeSymbolEncodingMinerCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("usage: OfficialMathTypeSymbolEncodingMinerCli <manifest.json> <out.tsv> <report.json>");
            System.exit(2);
        }
        Path manifestPath = Path.of(args[0]).toAbsolutePath();
        Path outputPath = Path.of(args[1]).toAbsolutePath();
        Path reportPath = Path.of(args[2]).toAbsolutePath();
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        JsonNode manifest = mapper.readTree(manifestPath.toFile());
        MathTypeDocxComparator inspector = new MathTypeDocxComparator();
        Map<String, List<Observation>> observations = new LinkedHashMap<>();
        int acceptedFormulas = 0;
        int rejectedFormulas = 0;

        for (JsonNode batch : manifest.path("batches")) {
            Path standard = Path.of(batch.path("standardDocx").asText());
            JsonNode formulas = mapper.readTree(Path.of(batch.path("formulas").asText()).toFile());
            MathTypeDocxComparator.InspectionReport inspection = inspector.inspect(standard);
            if (inspection.formulas().size() != formulas.size()) {
                throw new IllegalStateException(standard + " has " + inspection.formulas().size()
                    + " objects for " + formulas.size() + " formulas");
            }
            for (int index = 0; index < formulas.size(); index++) {
                String formula = formulas.get(index).path("formula").asText();
                List<String> tokens = tokens(formula);
                List<MtefRecordNormalizer.CanonicalRecord> records = inspection.formulas().get(index).records();
                List<MtefRecordNormalizer.CanonicalRecord> chars = records.stream()
                    .filter(record -> "CHAR".equals(record.name())).toList();
                boolean structural = records.stream().anyMatch(record -> STRUCTURAL_RECORDS.contains(record.name()));
                if (structural || tokens.size() != chars.size()) {
                    rejectedFormulas++;
                    continue;
                }
                acceptedFormulas++;
                for (int tokenIndex = 0; tokenIndex < tokens.size(); tokenIndex++) {
                    String token = tokens.get(tokenIndex);
                    if (!COMMAND.matcher(token).matches() || MtefCharMap.encodingProfile(token) == null) {
                        continue;
                    }
                    MtefRecordNormalizer.CanonicalRecord record = chars.get(tokenIndex);
                    if (record.typeface() == null || record.mtcode() == null) {
                        continue;
                    }
                    Encoding encoding = new Encoding(decodeTypeface(record.typeface()), record.mtcode(),
                        record.value() == null ? -1 : record.value(), record.options() == null ? 0 : record.options());
                    observations.computeIfAbsent(token, ignored -> new ArrayList<>())
                        .add(new Observation(batch.path("name").asText(), index, formula, encoding));
                }
            }
        }

        List<CommandResult> results = new ArrayList<>();
        List<String> lines = new ArrayList<>();
        lines.add("# latex-command<TAB>mtef-typeface<TAB>mtcode-hex<TAB>bits8-decimal");
        lines.add("# Mined from MathType 7.11.1.462 Equation Native objects; see the adjacent JSON report.");
        for (Map.Entry<String, List<Observation>> entry : observations.entrySet().stream()
                .sorted(Map.Entry.comparingByKey()).toList()) {
            Set<Encoding> distinct = new LinkedHashSet<>();
            entry.getValue().forEach(observation -> distinct.add(observation.encoding()));
            boolean unambiguous = distinct.size() == 1;
            Encoding encoding = unambiguous ? distinct.iterator().next() : null;
            results.add(new CommandResult(entry.getKey(), entry.getValue().size(), distinct,
                unambiguous, entry.getValue().stream().limit(5).toList()));
            if (unambiguous) {
                lines.add(entry.getKey() + "\t" + encoding.typeface() + "\t"
                    + Integer.toHexString(encoding.mtcode()).toUpperCase() + "\t" + encoding.bits8());
            }
        }
        results.sort(Comparator.comparing(CommandResult::command));
        Files.createDirectories(outputPath.getParent());
        Files.write(outputPath, lines, StandardCharsets.UTF_8);
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("sourceManifest", manifestPath.toString());
        report.put("acceptedFormulaCount", acceptedFormulas);
        report.put("rejectedFormulaCount", rejectedFormulas);
        report.put("observedCommandCount", results.size());
        report.put("unambiguousCommandCount", results.stream().filter(CommandResult::unambiguous).count());
        report.put("commands", results);
        Files.createDirectories(reportPath.getParent());
        mapper.writeValue(reportPath.toFile(), report);
        System.out.printf("acceptedFormulas=%d commands=%d unambiguous=%d tsv=%s report=%s%n",
            acceptedFormulas, results.size(), results.stream().filter(CommandResult::unambiguous).count(),
            outputPath, reportPath);
    }

    private static List<String> tokens(String formula) {
        List<String> result = new ArrayList<>();
        Matcher matcher = TOKEN.matcher(formula);
        while (matcher.find()) {
            result.add(matcher.group());
        }
        return result;
    }

    private static int decodeTypeface(int storedTypeface) {
        return storedTypeface >= 0x80 ? storedTypeface & 0x7F : storedTypeface;
    }

    private record Encoding(int typeface, int mtcode, int bits8, int options) {
    }

    private record Observation(String batch, int formulaIndex, String formula, Encoding encoding) {
    }

    private record CommandResult(String command, int observationCount, Set<Encoding> distinctEncodings,
                                 boolean unambiguous, List<Observation> samples) {
    }
}
