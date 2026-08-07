package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

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

/** Mines unambiguous MathType CHAR encodings from the pinned xsc source corpus. */
public final class XscSymbolEncodingMinerCli {

    private static final Pattern COMMAND = Pattern.compile("\\\\[A-Za-z]+");

    private XscSymbolEncodingMinerCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 4) {
            System.err.println("usage: XscSymbolEncodingMinerCli <source-docx-dir> <latex-report-dir>"
                + " <symbol-report.json> <out.json>");
            System.exit(2);
        }
        Path sourceDir = Path.of(args[0]).toAbsolutePath();
        Path reportDir = Path.of(args[1]).toAbsolutePath();
        Path symbolReport = Path.of(args[2]).toAbsolutePath();
        Path output = Path.of(args[3]).toAbsolutePath();
        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

        Set<String> candidates = new LinkedHashSet<>();
        for (JsonNode profile : mapper.readTree(symbolReport.toFile()).path("profiles")) {
            if (!profile.path("mathTypeVerified").asBoolean()) {
                candidates.add(profile.path("latex").asText());
            }
        }

        Map<String, List<Observation>> observations = new LinkedHashMap<>();
        Map<String, Integer> formulaCounts = new LinkedHashMap<>();
        MathTypeDocxComparator inspector = new MathTypeDocxComparator();
        List<Path> reports;
        try (var stream = Files.walk(reportDir, 2)) {
            reports = stream.filter(path -> path.getFileName().toString().endsWith(".report.json"))
                .sorted(Comparator.comparingInt(XscSymbolEncodingMinerCli::numericStem))
                .toList();
        }
        int inspectedDocuments = 0;
        int inspectedObjects = 0;
        for (Path reportPath : reports) {
            int documentIndex = numericStem(reportPath);
            JsonNode report = mapper.readTree(reportPath.toFile());
            List<FormulaTarget> targets = new ArrayList<>();
            for (JsonNode equation : report.path("equations")) {
                String formula = equation.path("output").asText();
                Set<String> commands = commands(formula);
                commands.retainAll(candidates);
                if (!commands.isEmpty()) {
                    targets.add(new FormulaTarget(equation.path("index").asInt(), formula, commands));
                    commands.forEach(command -> formulaCounts.merge(command, 1, Integer::sum));
                }
            }
            if (targets.isEmpty()) {
                continue;
            }
            Path source = sourceDir.resolve(documentIndex + ".docx");
            if (!Files.exists(source)) {
                throw new IllegalStateException("missing xsc source DOCX: " + source);
            }
            MathTypeDocxComparator.InspectionReport inspection = inspector.inspect(source);
            inspectedDocuments++;
            inspectedObjects += inspection.formulas().size();
            for (FormulaTarget target : targets) {
                if (target.index() < 1 || target.index() > inspection.formulas().size()) {
                    continue;
                }
                MathTypeDocxComparator.InspectedFormula formula = inspection.formulas().get(target.index() - 1);
                if (!formula.valid()) {
                    continue;
                }
                for (String command : target.commands()) {
                    MtefCharMap.EncodingProfile expected = MtefCharMap.encodingProfile(command);
                    if (expected == null) {
                        continue;
                    }
                    Set<Encoding> matches = new LinkedHashSet<>();
                    for (MtefRecordNormalizer.CanonicalRecord record : formula.records()) {
                        if ("CHAR".equals(record.name()) && record.mtcode() != null
                                && record.mtcode() == expected.mtcode()) {
                            matches.add(new Encoding(record.typeface(), record.mtcode(),
                                record.value() == null ? -1 : record.value(), record.options()));
                        }
                    }
                    if (matches.size() == 1) {
                        observations.computeIfAbsent(command, ignored -> new ArrayList<>())
                            .add(new Observation(documentIndex, target.index(), target.formula(),
                                matches.iterator().next()));
                    }
                }
            }
        }

        List<Map<String, Object>> commands = new ArrayList<>();
        for (String command : formulaCounts.keySet().stream().sorted().toList()) {
            List<Observation> items = observations.getOrDefault(command, List.of());
            Set<Encoding> encodings = items.stream().map(Observation::encoding)
                .collect(java.util.stream.Collectors.toCollection(LinkedHashSet::new));
            int formulas = formulaCounts.get(command);
            Map<String, Object> item = new LinkedHashMap<>();
            item.put("command", command);
            item.put("formulaCount", formulas);
            item.put("matchedFormulaCount", items.size());
            item.put("distinctEncodings", encodings);
            item.put("highConfidence", items.size() >= 2 && encodings.size() == 1);
            item.put("samples", items.stream().limit(5).toList());
            commands.add(item);
        }
        Map<String, Object> result = new LinkedHashMap<>();
        result.put("schemaVersion", 1);
        result.put("source", "xsc Equation Native corpus");
        result.put("inspectedDocuments", inspectedDocuments);
        result.put("inspectedObjects", inspectedObjects);
        result.put("commands", commands);
        Files.createDirectories(output.getParent());
        mapper.writeValue(output.toFile(), result);
        long verified = commands.stream().filter(item -> Boolean.TRUE.equals(item.get("highConfidence"))).count();
        System.out.printf("commands=%d highConfidence=%d report=%s%n", commands.size(), verified, output);
    }

    private static Set<String> commands(String formula) {
        Set<String> commands = new LinkedHashSet<>();
        Matcher matcher = COMMAND.matcher(formula);
        while (matcher.find()) {
            commands.add(matcher.group());
        }
        return commands;
    }

    private static int numericStem(Path path) {
        String name = path.getFileName().toString();
        String digits = name.replaceFirst("\\..*$", "").replaceAll("\\D+", "");
        if (digits.isEmpty() && path.getParent() != null) {
            digits = path.getParent().getFileName().toString().replaceAll("\\D+", "");
        }
        return digits.isEmpty() ? Integer.MAX_VALUE : Integer.parseInt(digits);
    }

    private record FormulaTarget(int index, String formula, Set<String> commands) {}
    private record Encoding(Integer typeface, Integer mtcode, int bits8, Integer options) {}
    private record Observation(int document, int formulaIndex, String formula, Encoding encoding) {}
}
