package com.lz.paperword.support;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.TextNode;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Applies pinned xsc extraction repairs only at their exact source locations. */
public final class XscSourceReplacementRepairs {

    private static final String RESOURCE = "/xsc-source-replacement-repairs.json";
    private static final Pattern PATH_TOKEN = Pattern.compile(
        "([A-Za-z][A-Za-z0-9]*)(?:\\[(\\d+)])?");
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Catalog CATALOG = loadCatalog();

    private XscSourceReplacementRepairs() {
    }

    public static JsonNode apply(int documentIndex, JsonNode requestRoot) {
        if (requestRoot == null || !requestRoot.isObject()) {
            throw new IllegalArgumentException("xsc request root must be a JSON object");
        }
        JsonNode repairedRoot = requestRoot.deepCopy();
        for (RepairInfo repair : CATALOG.byDocument().getOrDefault(documentIndex, List.of())) {
            applyOne(repairedRoot, repair);
        }
        List<String> remaining = new ArrayList<>();
        collectReplacementCharacters(repairedRoot, "", remaining);
        if (!remaining.isEmpty()) {
            throw new IllegalArgumentException(
                "UNCATALOGUED_SOURCE_REPLACEMENT_CHARACTER: xsc document " + documentIndex
                    + " still contains U+FFFD at " + remaining);
        }
        return repairedRoot;
    }

    public static Optional<RepairInfo> repairFor(int documentIndex, String sourceLocation) {
        return Optional.ofNullable(CATALOG.byLocation().get(new RepairKey(documentIndex, sourceLocation)));
    }

    public static int occurrenceCount() {
        return CATALOG.byLocation().size();
    }

    public static int formulaRepairOccurrenceCount() {
        return (int) CATALOG.byLocation().values().stream()
            .filter(repair -> repair.formulaOrdinal() > 0)
            .count();
    }

    public static int replacementFormulaOccurrenceCount() {
        return (int) CATALOG.byLocation().values().stream()
            .filter(repair -> repair.formulaOrdinal() > 0
                && repair.sourceLatex().indexOf('\uFFFD') >= 0)
            .count();
    }

    public static int occurrenceCount(int documentIndex) {
        return CATALOG.byDocument().getOrDefault(documentIndex, List.of()).size();
    }

    public static int documentCount() {
        return CATALOG.byDocument().size();
    }

    private static void applyOne(JsonNode root, RepairInfo repair) {
        Location location = parseLocation(repair.sourceLocation());
        ObjectNode owner = resolveOwner(root, location.pathTokens(), repair);
        JsonNode value = owner.get(location.fieldName());
        if (value == null || !value.isTextual()) {
            throw stale(repair, "target field is missing or is not text");
        }
        String content = value.asText();
        FormulaRange range;
        String repairScope;
        if (location.mathIndex() == 0) {
            range = new FormulaRange(0, content.length());
            repairScope = content;
        } else {
            List<FormulaRange> formulas = formulaRanges(content);
            if (location.mathIndex() > formulas.size()) {
                throw stale(repair, "formula index " + location.mathIndex()
                    + " exceeds field formula count " + formulas.size());
            }
            range = formulas.get(location.mathIndex() - 1);
            repairScope = content.substring(range.start(), range.end());
        }
        int sourceAt = repairScope.indexOf(repair.sourceLatex());
        if (sourceAt < 0) {
            throw stale(repair, "exact source text was not found in the indexed field/formula");
        }
        if (repairScope.indexOf(repair.sourceLatex(), sourceAt + repair.sourceLatex().length()) >= 0) {
            throw stale(repair, "exact source text occurs more than once in the indexed field/formula");
        }
        String repairedScope = repairScope.substring(0, sourceAt)
            + repair.repairedLatex()
            + repairScope.substring(sourceAt + repair.sourceLatex().length());
        owner.set(location.fieldName(), TextNode.valueOf(
            content.substring(0, range.start()) + repairedScope + content.substring(range.end())));
    }

    private static ObjectNode resolveOwner(JsonNode root, List<PathToken> tokens,
                                           RepairInfo repair) {
        JsonNode current = root;
        for (PathToken token : tokens) {
            current = current.get(token.name());
            if (current == null) {
                throw stale(repair, "missing JSON path token " + token.name());
            }
            if (token.index() != null) {
                if (!current.isArray() || token.index() < 0 || token.index() >= current.size()) {
                    throw stale(repair, "invalid JSON array index " + token.name()
                        + "[" + token.index() + "]");
                }
                current = current.get(token.index());
            }
        }
        if (!(current instanceof ObjectNode owner)) {
            throw stale(repair, "target field owner is not a JSON object");
        }
        return owner;
    }

    private static Location parseLocation(String sourceLocation) {
        int hash = sourceLocation.lastIndexOf('#');
        if (hash < 0 || hash + 1 >= sourceLocation.length()) {
            throw new IllegalStateException("invalid xsc repair sourceLocation: " + sourceLocation);
        }
        String selector = sourceLocation.substring(hash + 1);
        int mathIndex;
        if ("text".equals(selector)) {
            mathIndex = 0;
        } else if (selector.startsWith("math") && selector.length() > 4) {
            mathIndex = Integer.parseInt(selector.substring(4));
            if (mathIndex < 1) {
                throw new IllegalStateException(
                    "formula selector must be one-based: " + sourceLocation);
            }
        } else {
            throw new IllegalStateException("invalid xsc repair selector: " + sourceLocation);
        }
        String[] rawTokens = sourceLocation.substring(0, hash).split("\\.");
        if (rawTokens.length < 1 || mathIndex < 0) {
            throw new IllegalStateException("invalid xsc repair sourceLocation: " + sourceLocation);
        }
        List<PathToken> ownerTokens = new ArrayList<>();
        for (int index = 0; index < rawTokens.length - 1; index++) {
            ownerTokens.add(parsePathToken(rawTokens[index], sourceLocation));
        }
        PathToken field = parsePathToken(rawTokens[rawTokens.length - 1], sourceLocation);
        if (field.index() != null) {
            throw new IllegalStateException("indexed terminal field is not supported: " + sourceLocation);
        }
        return new Location(List.copyOf(ownerTokens), field.name(), mathIndex);
    }

    private static PathToken parsePathToken(String raw, String sourceLocation) {
        Matcher matcher = PATH_TOKEN.matcher(raw);
        if (!matcher.matches()) {
            throw new IllegalStateException("invalid path token in xsc repair: " + sourceLocation);
        }
        Integer index = matcher.group(2) == null ? null : Integer.valueOf(matcher.group(2));
        return new PathToken(matcher.group(1), index);
    }

    private static List<FormulaRange> formulaRanges(String content) {
        List<FormulaRange> formulas = new ArrayList<>();
        int searchFrom = 0;
        while (searchFrom < content.length()) {
            int start = nextUnescapedDollar(content, searchFrom);
            if (start < 0) {
                break;
            }
            int delimiterLength = start + 1 < content.length()
                && content.charAt(start + 1) == '$' ? 2 : 1;
            int end = closingDollar(content, start + delimiterLength, delimiterLength);
            if (end < 0) {
                throw new IllegalArgumentException("unclosed formula delimiter in xsc repair field");
            }
            formulas.add(new FormulaRange(start + delimiterLength, end));
            searchFrom = end + delimiterLength;
        }
        return List.copyOf(formulas);
    }

    private static int nextUnescapedDollar(String text, int from) {
        for (int index = Math.max(0, from); index < text.length(); index++) {
            if (text.charAt(index) == '$' && !isEscaped(text, index)) {
                return index;
            }
        }
        return -1;
    }

    private static int closingDollar(String text, int from, int delimiterLength) {
        for (int index = from; index < text.length(); index++) {
            if (text.charAt(index) != '$' || isEscaped(text, index)) {
                continue;
            }
            if (delimiterLength == 1) {
                return index;
            }
            if (index + 1 < text.length() && text.charAt(index + 1) == '$') {
                return index;
            }
        }
        return -1;
    }

    private static boolean isEscaped(String text, int index) {
        int backslashes = 0;
        for (int cursor = index - 1; cursor >= 0 && text.charAt(cursor) == '\\'; cursor--) {
            backslashes++;
        }
        return (backslashes & 1) != 0;
    }

    private static void collectReplacementCharacters(JsonNode node, String path,
                                                     List<String> locations) {
        if (node.isTextual()) {
            if (node.asText().indexOf('\uFFFD') >= 0) {
                locations.add(path.isEmpty() ? "/" : path);
            }
            return;
        }
        if (node.isArray()) {
            for (int index = 0; index < node.size(); index++) {
                collectReplacementCharacters(node.get(index), path + "/" + index, locations);
            }
            return;
        }
        if (node.isObject()) {
            node.fields().forEachRemaining(entry -> collectReplacementCharacters(
                entry.getValue(), path + "/" + escapePointer(entry.getKey()), locations));
        }
    }

    private static String escapePointer(String value) {
        return value.replace("~", "~0").replace("/", "~1");
    }

    private static IllegalArgumentException stale(RepairInfo repair, String detail) {
        return new IllegalArgumentException("STALE_XSC_SOURCE_REPAIR: document "
            + repair.documentIndex() + ", formula " + repair.formulaOrdinal() + ", "
            + repair.sourceLocation() + ": " + detail);
    }

    private static Catalog loadCatalog() {
        try (InputStream input = XscSourceReplacementRepairs.class.getResourceAsStream(RESOURCE)) {
            if (input == null) {
                throw new IllegalStateException("missing xsc source repair catalog: " + RESOURCE);
            }
            JsonNode root = JSON.readTree(input);
            if (root.path("schemaVersion").asInt() != 1 || !root.path("repairs").isArray()) {
                throw new IllegalStateException("invalid xsc source repair catalog schema");
            }
            Map<Integer, List<RepairInfo>> byDocument = new LinkedHashMap<>();
            Map<RepairKey, RepairInfo> byLocation = new HashMap<>();
            for (JsonNode repairNode : root.path("repairs")) {
                String reason = repairNode.path("reason").asText().trim();
                String source = decode(repairNode.path("sourceBase64").asText());
                String repaired = decode(repairNode.path("repairedBase64").asText());
                if (reason.isEmpty() || source.equals(repaired)
                        || repaired.indexOf('\uFFFD') >= 0) {
                    throw new IllegalStateException(
                        "xsc source repair needs a reason, must change the exact source, "
                            + "and cannot emit U+FFFD");
                }
                for (JsonNode occurrence : repairNode.path("occurrences")) {
                    RepairInfo info = new RepairInfo(
                        occurrence.path("documentIndex").asInt(),
                        occurrence.path("formulaOrdinal").asInt(),
                        occurrence.path("sourceLocation").asText(),
                        source,
                        repaired,
                        reason);
                    Location location = parseLocation(info.sourceLocation());
                    boolean invalidOrdinal = location.mathIndex() == 0
                        ? info.formulaOrdinal() != 0
                        : info.formulaOrdinal() < 1;
                    if (info.documentIndex() < 1 || invalidOrdinal
                            || info.sourceLocation().isBlank()) {
                        throw new IllegalStateException("invalid xsc source repair occurrence");
                    }
                    RepairKey key = new RepairKey(info.documentIndex(), info.sourceLocation());
                    if (byLocation.putIfAbsent(key, info) != null) {
                        throw new IllegalStateException("duplicate xsc source repair: " + key);
                    }
                    byDocument.computeIfAbsent(info.documentIndex(), ignored -> new ArrayList<>())
                        .add(info);
                }
            }
            Map<Integer, List<RepairInfo>> immutableByDocument = new LinkedHashMap<>();
            byDocument.forEach((document, repairs) ->
                immutableByDocument.put(document, List.copyOf(repairs)));
            return new Catalog(Map.copyOf(immutableByDocument), Map.copyOf(byLocation));
        } catch (IllegalArgumentException exception) {
            throw new IllegalStateException("invalid Base64 in xsc source repair catalog", exception);
        } catch (Exception exception) {
            if (exception instanceof IllegalStateException state) {
                throw state;
            }
            throw new IllegalStateException("cannot load xsc source repair catalog", exception);
        }
    }

    private static String decode(String base64) {
        return new String(Base64.getDecoder().decode(base64), StandardCharsets.UTF_8);
    }

    public record RepairInfo(int documentIndex, int formulaOrdinal, String sourceLocation,
                             String sourceLatex, String repairedLatex, String reason) {

        public String kind() {
            return sourceLatex.indexOf('\uFFFD') >= 0
                ? "replacement-character"
                : "semantic-placeholder";
        }
    }

    private record RepairKey(int documentIndex, String sourceLocation) {
    }

    private record Catalog(Map<Integer, List<RepairInfo>> byDocument,
                           Map<RepairKey, RepairInfo> byLocation) {
    }

    private record Location(List<PathToken> pathTokens, String fieldName, int mathIndex) {
    }

    private record PathToken(String name, Integer index) {
    }

    private record FormulaRange(int start, int end) {
    }
}
