package com.lz.paperword.integration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.mtef.MtefRecord;
import com.lz.paperword.core.mtef.MtefRecordNormalizer;
import com.lz.paperword.core.mtef.MtefWriter;
import com.lz.paperword.core.render.SvgVectorEmfPlusRenderer;
import com.lz.paperword.support.XscSourceReplacementRepairs;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.jsoup.Jsoup;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Deque;
import java.util.HashSet;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class XscEmfPlusCorpusAcceptanceTest {

    private static final String WORD_NS =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final String OFFICE_NS = "urn:schemas-microsoft-com:office:office";
    private static final String VML_NS = "urn:schemas-microsoft-com:vml";
    private static final String DOCUMENT_REL_NS =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String PACKAGE_REL_NS =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final String OLE_REL_TYPE = DOCUMENT_REL_NS + "/oleObject";
    private static final String IMAGE_REL_TYPE = DOCUMENT_REL_NS + "/image";
    private static final Pattern DOCUMENT_INDEX = Pattern.compile("_(\\d+)\\.docx$");
    private static final Pattern TRACE = Pattern.compile("pwf:(\\d+)-([0-9a-f]{16})");
    private static final Pattern TRACE_METRICS =
        Pattern.compile("^\\\\pwmetrics\\{[^}]+}[ \\t\\n\\x0B\\f\\r]*");
    private static final Pattern TRACE_STYLE =
        Pattern.compile("^\\\\pwstyle\\{[^}]*}[ \\t\\n\\x0B\\f\\r]*");
    private static final Pattern EDGE_SPACE =
        Pattern.compile("^[ \\t\\n\\x0B\\f\\r]+|[ \\t\\n\\x0B\\f\\r]+$");
    private static final Pattern INNER_SPACE = Pattern.compile("[ \\t\\n\\x0B\\f\\r]+");

    private static final Path CORPUS = Path.of(System.getProperty(
        "xsc.emfplus.corpus.dir", "target/xsc-emfplus-final"));
    private static final Path REQUESTS = Path.of(System.getProperty(
        "xsc.emfplus.requests.dir",
        "analysis/batch10-full-requests/20260807-pure-jvm-full"));
    private static final Path SOURCE_REPORTS = Path.of(System.getProperty(
        "xsc.emfplus.sourceReports.dir", "analysis/xsc-latex-fixed"));
    private static final Path REPORT = Path.of(System.getProperty(
        "xsc.emfplus.report",
        "target/vector-acceptance/xsc-emfplus-corpus.json"));
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final LaTeXParser LATEX_PARSER = new LaTeXParser();
    private static final ThreadLocal<MtefWriter> MTEF_WRITER =
        ThreadLocal.withInitial(MtefWriter::new);

    @Test
    void everyXscFormulaHasOneTraceableEditableOleAndStrictEmfPlusDualPreview() throws Exception {
        Assumptions.assumeTrue(Boolean.getBoolean("paperword.acceptance.xscEmfPlus"),
            "Enable with -Dpaperword.acceptance.xscEmfPlus=true");
        assertTrue(Files.isDirectory(CORPUS), "Missing xsc corpus: " + CORPUS);
        assertTrue(Files.isDirectory(REQUESTS), "Missing xsc requests: " + REQUESTS);

        int expectedDocuments = Integer.getInteger("xsc.emfplus.expectedDocuments", 551);
        int expectedFullFormulaCount = Integer.getInteger(
            "xsc.emfplus.expectedFormulaCount", 93_319);
        int expectedSourceRepairCount = Integer.getInteger(
            "xsc.emfplus.expectedSourceRepairCount", 52);
        int expectedFormulaRepairCount = Integer.getInteger(
            "xsc.emfplus.expectedFormulaRepairCount", 46);
        int expectedReplacementFormulaCount = Integer.getInteger(
            "xsc.emfplus.expectedReplacementFormulaCount", 37);
        int startDocument = Integer.getInteger("xsc.emfplus.startDocument", 1);
        int endDocument = Integer.getInteger("xsc.emfplus.endDocument", expectedDocuments);
        if (startDocument < 1 || endDocument < startDocument || endDocument > expectedDocuments) {
            throw new IllegalArgumentException("invalid xsc document scope: "
                + startDocument + ".." + endDocument + " of " + expectedDocuments);
        }

        List<IndexedPath> allDocuments = indexedFiles(CORPUS, DOCUMENT_INDEX, ".docx");
        List<Map<String, Object>> failures = new ArrayList<>();
        long rawDocxCount;
        try (var paths = Files.list(CORPUS)) {
            rawDocxCount = paths.filter(Files::isRegularFile)
                .filter(path -> path.getFileName().toString()
                    .toLowerCase(Locale.ROOT).endsWith(".docx"))
                .count();
        }
        if (rawDocxCount != allDocuments.size()) {
            addFailure(failures, 0, null, null, "DOCUMENT_NAME_NOT_INDEXED",
                "DOCX files=" + rawDocxCount + ", names matching " + DOCUMENT_INDEX
                    + "=" + allDocuments.size());
        }
        validateCompleteIndexSet(allDocuments, expectedDocuments, "corpus DOCX", failures);
        List<IndexedPath> documents = allDocuments.stream()
            .filter(item -> item.index() >= startDocument && item.index() <= endDocument)
            .toList();
        int parallelism = Math.max(1, Integer.getInteger("xsc.emfplus.parallelism",
            Math.min(4, Runtime.getRuntime().availableProcessors())));

        Counters totals = new Counters();
        List<Map<String, Object>> documentSummaries = new ArrayList<>();
        for (DocumentResult result : validateDocuments(documents, parallelism)) {
            totals.add(result.counters());
            failures.addAll(result.failures());
            documentSummaries.add(result.summary());
        }

        boolean fullScope = startDocument == 1 && endDocument == expectedDocuments;
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 4);
        report.put("corpus", CORPUS.toAbsolutePath().normalize().toString());
        report.put("requests", REQUESTS.toAbsolutePath().normalize().toString());
        report.put("documentStart", startDocument);
        report.put("documentEnd", endDocument);
        report.put("parallelism", parallelism);
        report.put("corpusDocumentCount", allDocuments.size());
        report.put("rawDocxCount", rawDocxCount);
        report.put("validatedDocumentCount", documents.size());
        report.put("expectedFormulaCount", totals.expectedFormulaCount);
        report.put("objectCount", totals.objectCount);
        report.put("traceMatchedCount", totals.traceMatchedCount);
        report.put("oleRelationshipCount", totals.oleRelationshipCount);
        report.put("oleCount", totals.oleCount);
        report.put("validOleCount", totals.validOleCount);
        report.put("mtefParsedCount", totals.mtefParsedCount);
        report.put("mtefBalancedCount", totals.mtefBalancedCount);
        report.put("mtefStructureComparableCount", totals.mtefStructureComparableCount);
        report.put("mtefStructureMatchedCount", totals.mtefStructureMatchedCount);
        report.put("emfRelationshipCount", totals.emfRelationshipCount);
        report.put("emfCount", totals.emfCount);
        report.put("validEmfPlusDualCount", totals.validEmfPlusDualCount);
        report.put("sourceRepairCount", totals.sourceRepairCount);
        report.put("formulaSourceRepairCount", totals.formulaSourceRepairCount);
        report.put("replacementFormulaRepairCount", totals.replacementFormulaRepairCount);
        report.put("sourceRepairApplications", totals.sourceRepairApplications);
        report.put("recoveredSourceReplacementCount", totals.recoveredSourceReplacementCount);
        report.put("sourceReplacementRecoveries", totals.sourceReplacementRecoveries);
        report.put("sourceReportsChecked", totals.sourceReportsChecked);
        report.put("expectedFullFormulaCount", expectedFullFormulaCount);
        report.put("expectedSourceRepairCount", expectedSourceRepairCount);
        report.put("expectedFormulaRepairCount", expectedFormulaRepairCount);
        report.put("expectedReplacementFormulaCount", expectedReplacementFormulaCount);
        report.put("failedDocumentCount", failures.stream()
            .map(failure -> failure.get("documentIndex")).distinct().count());
        report.put("failureCount", failures.size());
        report.put("failures", failures);
        report.put("documents", documentSummaries);
        boolean passed = rawDocxCount == allDocuments.size()
            && allDocuments.size() == expectedDocuments
            && documents.size() == endDocument - startDocument + 1
            && (!fullScope || totals.expectedFormulaCount == expectedFullFormulaCount)
            && (!fullScope || totals.oleCount == expectedFullFormulaCount)
            && (!fullScope || totals.sourceRepairCount == expectedSourceRepairCount)
            && (!fullScope || totals.formulaSourceRepairCount == expectedFormulaRepairCount)
            && (!fullScope
                || totals.replacementFormulaRepairCount == expectedReplacementFormulaCount)
            && totals.expectedFormulaCount == totals.objectCount
            && totals.objectCount == totals.traceMatchedCount
            && totals.objectCount == totals.oleRelationshipCount
            && totals.objectCount == totals.oleCount
            && totals.oleCount == totals.validOleCount
            && totals.oleCount == totals.mtefParsedCount
            && totals.oleCount == totals.mtefBalancedCount
            && totals.oleCount == totals.mtefStructureComparableCount
            && totals.oleCount == totals.mtefStructureMatchedCount
            && totals.objectCount == totals.emfRelationshipCount
            && totals.objectCount == totals.emfCount
            && totals.emfCount == totals.validEmfPlusDualCount
            && failures.isEmpty();
        report.put("passed", passed);
        Path reportParent = REPORT.toAbsolutePath().normalize().getParent();
        if (reportParent != null) {
            Files.createDirectories(reportParent);
        }
        JSON.enable(SerializationFeature.INDENT_OUTPUT).writeValue(REPORT.toFile(), report);

        assertEquals(expectedDocuments, allDocuments.size(), "xsc corpus document index coverage");
        assertEquals(rawDocxCount, allDocuments.size(), "every DOCX filename must expose its corpus index");
        assertEquals(endDocument - startDocument + 1, documents.size(), "selected xsc document scope");
        if (fullScope) {
            assertEquals(expectedFullFormulaCount, totals.expectedFormulaCount,
                "xsc request formula total");
            assertEquals(expectedFullFormulaCount, totals.oleCount,
                "xsc OLE formula total");
            assertEquals(expectedSourceRepairCount, totals.sourceRepairCount,
                "xsc source repair total");
            assertEquals(expectedFormulaRepairCount, totals.formulaSourceRepairCount,
                "xsc formula source repair total");
            assertEquals(expectedReplacementFormulaCount,
                totals.replacementFormulaRepairCount,
                "xsc U+FFFD formula repair total");
        }
        assertEquals(totals.expectedFormulaCount, totals.objectCount,
            "every request formula must have one w:object");
        assertEquals(totals.objectCount, totals.traceMatchedCount,
            "every w:object trace must match its request formula and ordinal");
        assertEquals(totals.objectCount, totals.oleRelationshipCount,
            "every formula must have one exclusive OLE relationship");
        assertEquals(totals.oleCount, totals.validOleCount,
            "every xsc OLE must be Equation.DSMT4");
        assertEquals(totals.oleCount, totals.mtefParsedCount,
            "every Equation Native payload must be parsed by the MTEF normalizer");
        assertEquals(totals.oleCount, totals.mtefBalancedCount,
            "every normalized MTEF tree must have balanced containers and END records");
        assertEquals(totals.oleCount, totals.mtefStructureComparableCount,
            "every source formula must produce an expected normalized MTEF tree");
        assertEquals(totals.oleCount, totals.mtefStructureMatchedCount,
            "every Equation Native payload must match its source formula's normalized MTEF tree");
        assertEquals(totals.objectCount, totals.emfRelationshipCount,
            "every formula must have one exclusive EMF relationship");
        assertEquals(totals.objectCount, totals.emfCount,
            "every formula object must resolve to one EMF preview");
        assertEquals(totals.emfCount, totals.validEmfPlusDualCount,
            "every xsc preview must be strict pure-vector EMF+ Dual");
        assertTrue(failures.isEmpty(), () -> failures.size()
            + " xsc validation failures; inspect " + REPORT.toAbsolutePath());
    }

    private static List<DocumentResult> validateDocuments(List<IndexedPath> documents,
                                                          int parallelism) throws Exception {
        if (parallelism == 1 || documents.size() <= 1) {
            return documents.stream().map(XscEmfPlusCorpusAcceptanceTest::validateDocument).toList();
        }
        ExecutorService executor = Executors.newFixedThreadPool(
            Math.min(parallelism, documents.size()));
        try {
            List<Future<DocumentResult>> futures = new ArrayList<>(documents.size());
            for (IndexedPath document : documents) {
                futures.add(executor.submit(() -> validateDocument(document)));
            }
            List<DocumentResult> results = new ArrayList<>(documents.size());
            for (Future<DocumentResult> future : futures) {
                results.add(future.get());
            }
            return List.copyOf(results);
        } finally {
            executor.shutdownNow();
        }
    }

    private static DocumentResult validateDocument(IndexedPath indexed) {
        int documentIndex = indexed.index();
        Path document = indexed.path();
        List<Map<String, Object>> failures = new ArrayList<>();
        Counters counters = new Counters();
        List<ExpectedFormula> expected = List.of();
        Path request = REQUESTS.resolve(String.format(Locale.ROOT,
            "full-%02d.request.json", documentIndex));
        try {
            expected = expectedFormulas(request, documentIndex, document, failures, counters);
            counters.expectedFormulaCount = expected.size();
            validateOptionalSourceReportCount(documentIndex, expected.size(), failures, counters);
        } catch (Exception exception) {
            addFailure(failures, documentIndex, document, null, "REQUEST_MANIFEST_INVALID",
                "Could not reconstruct formula manifest from " + request + ": "
                    + exceptionChain(exception));
        }

        try (ZipFile zip = new ZipFile(document.toFile())) {
            Document wordDocument = parseXml(zip, "word/document.xml");
            Map<String, RelationshipInfo> relationships = relationships(zip);
            NodeList objects = wordDocument.getElementsByTagNameNS(WORD_NS, "object");
            counters.objectCount = objects.getLength();
            Set<String> usedOleRelationshipIds = new LinkedHashSet<>();
            Set<String> usedImageRelationshipIds = new LinkedHashSet<>();
            Set<String> referencedOleEntries = new LinkedHashSet<>();
            Set<String> referencedEmfEntries = new LinkedHashSet<>();

            if (objects.getLength() != expected.size()) {
                addFailure(failures, documentIndex, document, null, "FORMULA_OBJECT_COUNT_MISMATCH",
                    "request formulas=" + expected.size() + ", w:object=" + objects.getLength());
            }
            for (int objectIndex = 0; objectIndex < objects.getLength(); objectIndex++) {
                int ordinal = objectIndex + 1;
                ExpectedFormula formula = objectIndex < expected.size() ? expected.get(objectIndex) : null;
                Element object = (Element) objects.item(objectIndex);
                validateFormulaObject(zip, relationships, object, documentIndex, document,
                    ordinal, formula, failures, counters, usedOleRelationshipIds,
                    usedImageRelationshipIds, referencedOleEntries, referencedEmfEntries);
            }

            Set<String> packageOleEntries = zipEntries(zip, "word/embeddings/", ".bin");
            Set<String> packageEmfEntries = zipEntries(zip, "word/media/", ".emf");
            validateSetEquality(failures, documentIndex, document, "OLE_PACKAGE_CLOSURE",
                "referenced OLE entries", referencedOleEntries,
                "packaged OLE entries", packageOleEntries);
            validateSetEquality(failures, documentIndex, document, "EMF_PACKAGE_CLOSURE",
                "referenced EMF entries", referencedEmfEntries,
                "packaged EMF entries", packageEmfEntries);

            Set<String> declaredOleRelationshipIds = relationshipIds(relationships, OLE_REL_TYPE, null);
            Set<String> declaredEmfRelationshipIds = relationshipIds(relationships, IMAGE_REL_TYPE, ".emf");
            validateSetEquality(failures, documentIndex, document, "OLE_RELATIONSHIP_CLOSURE",
                "used OLE relationship ids", usedOleRelationshipIds,
                "declared OLE relationship ids", declaredOleRelationshipIds);
            validateSetEquality(failures, documentIndex, document, "EMF_RELATIONSHIP_CLOSURE",
                "used EMF relationship ids", usedImageRelationshipIds,
                "declared EMF relationship ids", declaredEmfRelationshipIds);
        } catch (Exception exception) {
            addFailure(failures, documentIndex, document, null, "DOCX_INVALID",
                exceptionChain(exception));
        }

        Map<String, Object> summary = new LinkedHashMap<>();
        summary.put("documentIndex", documentIndex);
        summary.put("document", document.toString());
        summary.put("request", request.toString());
        summary.put("expectedFormulaCount", counters.expectedFormulaCount);
        summary.put("objectCount", counters.objectCount);
        summary.put("traceMatchedCount", counters.traceMatchedCount);
        summary.put("validOleCount", counters.validOleCount);
        summary.put("mtefParsedCount", counters.mtefParsedCount);
        summary.put("mtefBalancedCount", counters.mtefBalancedCount);
        summary.put("mtefStructureComparableCount", counters.mtefStructureComparableCount);
        summary.put("mtefStructureMatchedCount", counters.mtefStructureMatchedCount);
        summary.put("validEmfPlusDualCount", counters.validEmfPlusDualCount);
        summary.put("recoveredSourceReplacementCount", counters.recoveredSourceReplacementCount);
        summary.put("failureCount", failures.size());
        return new DocumentResult(counters, summary, List.copyOf(failures));
    }

    private static void validateFormulaObject(ZipFile zip,
                                              Map<String, RelationshipInfo> relationships,
                                              Element object,
                                              int documentIndex,
                                              Path document,
                                              int ordinal,
                                              ExpectedFormula expected,
                                              List<Map<String, Object>> failures,
                                              Counters counters,
                                              Set<String> usedOleRelationshipIds,
                                              Set<String> usedImageRelationshipIds,
                                              Set<String> referencedOleEntries,
                                              Set<String> referencedEmfEntries) {
        NodeList oleObjects = object.getElementsByTagNameNS(OFFICE_NS, "OLEObject");
        NodeList imageData = object.getElementsByTagNameNS(VML_NS, "imagedata");
        NodeList shapes = object.getElementsByTagNameNS(VML_NS, "shape");
        if (oleObjects.getLength() != 1 || imageData.getLength() != 1 || shapes.getLength() != 1) {
            addFailure(failures, documentIndex, document, expected,
                "FORMULA_OBJECT_CARDINALITY", "formula " + ordinal + " has OLEObject="
                    + oleObjects.getLength() + ", imagedata=" + imageData.getLength()
                    + ", shape=" + shapes.getLength());
            return;
        }

        Element oleObject = (Element) oleObjects.item(0);
        Element image = (Element) imageData.item(0);
        Element shape = (Element) shapes.item(0);
        String trace = image.getAttributeNS(OFFICE_NS, "title");
        Matcher traceMatcher = TRACE.matcher(trace);
        if (expected != null && traceMatcher.matches()
                && Integer.parseInt(traceMatcher.group(1)) == ordinal
                && trace.equals(expected.trace())) {
            counters.traceMatchedCount++;
        } else {
            addFailure(failures, documentIndex, document, expected, "FORMULA_TRACE_MISMATCH",
                "formula " + ordinal + " actualTrace=" + trace
                    + ", expectedTrace=" + (expected == null ? "<missing>" : expected.trace()));
        }

        String shapeId = shape.getAttribute("id");
        if (!"t".equals(shape.getAttributeNS(OFFICE_NS, "ole"))
                || !shapeId.equals(oleObject.getAttribute("ShapeID"))
                || !"Embed".equals(oleObject.getAttribute("Type"))
                || !"Equation.DSMT4".equals(oleObject.getAttribute("ProgID"))) {
            addFailure(failures, documentIndex, document, expected, "OLE_OBJECT_METADATA_INVALID",
                "formula " + ordinal + " shapeId=" + shapeId + ", oleShapeId="
                    + oleObject.getAttribute("ShapeID") + ", o:ole="
                    + shape.getAttributeNS(OFFICE_NS, "ole") + ", Type="
                    + oleObject.getAttribute("Type") + ", ProgID="
                    + oleObject.getAttribute("ProgID"));
        }

        String oleRelationshipId = oleObject.getAttributeNS(DOCUMENT_REL_NS, "id");
        String imageRelationshipId = image.getAttributeNS(DOCUMENT_REL_NS, "id");
        RelationshipInfo oleRelationship = validateRelationship(relationships, oleRelationshipId,
            OLE_REL_TYPE, ".bin", usedOleRelationshipIds, failures, documentIndex, document,
            expected, "OLE");
        RelationshipInfo imageRelationship = validateRelationship(relationships, imageRelationshipId,
            IMAGE_REL_TYPE, ".emf", usedImageRelationshipIds, failures, documentIndex, document,
            expected, "EMF");

        if (oleRelationship != null) {
            counters.oleRelationshipCount++;
            if (!referencedOleEntries.add(oleRelationship.entryName())) {
                addFailure(failures, documentIndex, document, expected, "OLE_TARGET_REUSED",
                    oleRelationship.entryName() + " is referenced by more than one formula object");
            }
            ZipEntry oleEntry = zip.getEntry(oleRelationship.entryName());
            if (oleEntry == null) {
                addFailure(failures, documentIndex, document, expected, "OLE_TARGET_MISSING",
                    oleRelationship.entryName());
            } else {
                counters.oleCount++;
                try (InputStream input = zip.getInputStream(oleEntry)) {
                    OlePayload ole = validateOleContainer(input.readAllBytes());
                    counters.validOleCount++;
                    try {
                        MtefRecordNormalizer.NormalizationReport normalized =
                            MtefRecordNormalizer.normalize(ole.mtef());
                        validateNormalizedMtef(normalized);
                        if (normalized.records().isEmpty()) {
                            throw new IllegalStateException("MTEF normalized record tree is empty");
                        }
                        counters.mtefParsedCount++;
                        if (expected != null && expected.expectedMtefRecords() != null) {
                            counters.mtefStructureComparableCount++;
                            if (expected.expectedMtefRecords().equals(normalized.records())) {
                                counters.mtefStructureMatchedCount++;
                            } else {
                                addFailure(failures, documentIndex, document, expected,
                                    "MTEF_STRUCTURE_MISMATCH", normalizedDifference(
                                        expected.expectedMtefRecords(), normalized.records()));
                            }
                        }
                    } catch (Exception exception) {
                        addFailure(failures, documentIndex, document, expected,
                            "MTEF_NORMALIZATION_INVALID", oleRelationship.entryName()
                                + ": sha256=" + ole.mtefSha256() + ": " + exceptionChain(exception));
                    }
                    try {
                        validateMtefStream(ole.mtef());
                        counters.mtefBalancedCount++;
                    } catch (Exception exception) {
                        addFailure(failures, documentIndex, document, expected,
                            "MTEF_BALANCE_INVALID", oleRelationship.entryName()
                                + ": sha256=" + ole.mtefSha256() + ": " + exceptionChain(exception));
                    }
                } catch (Exception exception) {
                    addFailure(failures, documentIndex, document, expected, "OLE_INVALID",
                        oleRelationship.entryName() + ": " + exceptionChain(exception));
                }
            }
        }

        if (imageRelationship != null) {
            counters.emfRelationshipCount++;
            if (!referencedEmfEntries.add(imageRelationship.entryName())) {
                addFailure(failures, documentIndex, document, expected, "EMF_TARGET_REUSED",
                    imageRelationship.entryName() + " is referenced by more than one formula object");
            }
            ZipEntry emfEntry = zip.getEntry(imageRelationship.entryName());
            if (emfEntry == null) {
                addFailure(failures, documentIndex, document, expected, "EMF_TARGET_MISSING",
                    imageRelationship.entryName());
            } else {
                counters.emfCount++;
                try (InputStream input = zip.getInputStream(emfEntry)) {
                    byte[] emf = input.readAllBytes();
                    if (SvgVectorEmfPlusRenderer.isValidDualVector(emf)) {
                        counters.validEmfPlusDualCount++;
                    } else {
                        addFailure(failures, documentIndex, document, expected,
                            "EMF_NOT_STRICT_VECTOR_DUAL", imageRelationship.entryName());
                    }
                } catch (Exception exception) {
                    addFailure(failures, documentIndex, document, expected, "EMF_INVALID",
                        imageRelationship.entryName() + ": " + exceptionChain(exception));
                }
            }
        }
    }

    private static RelationshipInfo validateRelationship(
            Map<String, RelationshipInfo> relationships,
            String relationshipId,
            String expectedType,
            String expectedSuffix,
            Set<String> usedRelationshipIds,
            List<Map<String, Object>> failures,
            int documentIndex,
            Path document,
            ExpectedFormula formula,
            String label) {
        if (relationshipId == null || relationshipId.isBlank()) {
            addFailure(failures, documentIndex, document, formula,
                label + "_RELATIONSHIP_ID_MISSING", "missing r:id");
            return null;
        }
        if (!usedRelationshipIds.add(relationshipId)) {
            addFailure(failures, documentIndex, document, formula,
                label + "_RELATIONSHIP_REUSED", relationshipId);
        }
        RelationshipInfo relationship = relationships.get(relationshipId);
        if (relationship == null) {
            addFailure(failures, documentIndex, document, formula,
                label + "_RELATIONSHIP_MISSING", relationshipId);
            return null;
        }
        if (!expectedType.equals(relationship.type())
                || relationship.external()
                || !relationship.entryName().toLowerCase(Locale.ROOT).endsWith(expectedSuffix)) {
            addFailure(failures, documentIndex, document, formula,
                label + "_RELATIONSHIP_INVALID", relationship.toString());
            return null;
        }
        return relationship;
    }

    private static OlePayload validateOleContainer(byte[] ole) throws Exception {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(ole))) {
            byte[] compObj = read((DocumentEntry) fs.getRoot().getEntry("\u0001CompObj"));
            byte[] nativeStream = read((DocumentEntry) fs.getRoot().getEntry("Equation Native"));
            String compObjAscii = new String(compObj, StandardCharsets.ISO_8859_1);
            String compObjUtf16 = new String(compObj, StandardCharsets.UTF_16LE);
            if (!compObjAscii.contains("Equation.DSMT4")
                    && !compObjUtf16.contains("Equation.DSMT4")) {
                throw new IllegalStateException("CompObj ProgID is not Equation.DSMT4");
            }
            if (nativeStream.length < 28) {
                throw new IllegalStateException("Equation Native is shorter than its 28-byte header");
            }
            ByteBuffer header = ByteBuffer.wrap(nativeStream).order(ByteOrder.LITTLE_ENDIAN);
            int headerSize = Short.toUnsignedInt(header.getShort(0));
            if (headerSize != 28) {
                throw new IllegalStateException("Equation Native header size is " + headerSize
                    + ", expected 28");
            }
            long nativeLength = Integer.toUnsignedLong(header.getInt(8));
            if (nativeLength <= 0 || nativeLength != nativeStream.length - headerSize) {
                throw new IllegalStateException("Equation Native cbObject mismatch: declared="
                    + nativeLength + ", actual=" + (nativeStream.length - headerSize));
            }
            byte[] mtef = Arrays.copyOfRange(nativeStream, headerSize, nativeStream.length);
            if (mtef.length < 10
                    || !"DSMT".equals(new String(mtef, 5, 4, StandardCharsets.US_ASCII))) {
                throw new IllegalStateException("MTEF DSMT header is missing");
            }
            String sha256 = HexFormat.of().formatHex(
                MessageDigest.getInstance("SHA-256").digest(mtef));
            return new OlePayload(mtef, sha256);
        }
    }

    private static void validateNormalizedMtef(MtefRecordNormalizer.NormalizationReport normalized) {
        if (normalized.contentOffset() < 0 || normalized.contentByteLength() <= 0
                || normalized.canonicalSignature().isBlank()) {
            throw new IllegalStateException("MTEF normalizer returned no formula content");
        }
        // MtefRecordNormalizer is also exercised here for canonical report generation.
        // Raw container balance is checked separately below because the current
        // normalizer treats EMBELL as a one-byte record while generated MTEF correctly
        // stores EMBELL(options, type), which would shift every later record.
    }

    private static String normalizedDifference(
            List<MtefRecordNormalizer.CanonicalRecord> expected,
            List<MtefRecordNormalizer.CanonicalRecord> actual) {
        int common = Math.min(expected.size(), actual.size());
        int firstDifference = common;
        for (int index = 0; index < common; index++) {
            if (!expected.get(index).equals(actual.get(index))) {
                firstDifference = index;
                break;
            }
        }
        return "firstDifference=" + firstDifference
            + ", expectedRecordCount=" + expected.size()
            + ", actualRecordCount=" + actual.size()
            + ", expectedRecord=" + (firstDifference < expected.size()
                ? expected.get(firstDifference) : "<end>")
            + ", actualRecord=" + (firstDifference < actual.size()
                ? actual.get(firstDifference) : "<end>");
    }

    private static int validateMtefStream(byte[] data) {
        int offset = strictRecordStart(data);
        Deque<Integer> containers = new ArrayDeque<>();
        int recordCount = 0;
        int rootEndCount = 0;
        while (offset < data.length) {
            StrictRecord record = parseStrictRecord(data, offset);
            recordCount++;
            if (record.tag() == MtefRecord.END) {
                if (containers.isEmpty()) {
                    rootEndCount++;
                    if (record.nextOffset() != data.length) {
                        throw new IllegalStateException("root END before byte " + data.length
                            + " at byte " + offset);
                    }
                } else {
                    containers.pop();
                }
            } else if (record.opensContainer()) {
                containers.push(record.tag());
            }
            offset = record.nextOffset();
        }
        if (offset != data.length) {
            throw new IllegalStateException("MTEF parser stopped at byte " + offset
                + " of " + data.length);
        }
        if (!containers.isEmpty()) {
            throw new IllegalStateException("unclosed MTEF container tags " + containers);
        }
        if (rootEndCount != 1) {
            throw new IllegalStateException("MTEF root END count is " + rootEndCount
                + ", expected 1");
        }
        return recordCount;
    }

    private static StrictRecord parseStrictRecord(byte[] data, int offset) {
        requireBytes(data, offset, 1, "record tag");
        int tag = Byte.toUnsignedInt(data[offset]);
        int cursor = offset + 1;
        boolean opensContainer = false;
        if (tag == MtefRecord.END || tag == MtefRecord.FULL || tag == MtefRecord.SUB
                || tag == MtefRecord.SUB2 || tag == MtefRecord.SYM || tag == MtefRecord.SUBSYM) {
            // No payload.
        } else if (tag == MtefRecord.LINE) {
            requireBytes(data, cursor, 1, "LINE options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor = skipStrictNudge(data, cursor);
            }
            if ((options & MtefRecord.OPT_LINE_LSPACE) != 0) {
                requireBytes(data, cursor, 1, "LINE spacing");
                cursor++;
            }
            if ((options & MtefRecord.OPT_LP_RULER) != 0) {
                requireBytes(data, cursor, 1, "LINE ruler count");
                int stops = Byte.toUnsignedInt(data[cursor++]);
                requireBytes(data, cursor, stops * 3, "LINE ruler stops");
                cursor += stops * 3;
            }
            opensContainer = (options & MtefRecord.OPT_LINE_NULL) == 0;
        } else if (tag == MtefRecord.CHAR) {
            requireBytes(data, cursor, 1, "CHAR options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor = skipStrictNudge(data, cursor);
            }
            if ((options & MtefRecord.OPT_CHAR_ENC_NO_MTCODE) == 0) {
                requireBytes(data, cursor, 3, "CHAR typeface and mtcode");
                cursor += 3;
            }
            if ((options & MtefRecord.OPT_CHAR_ENC_CHAR_8) != 0) {
                requireBytes(data, cursor, 1, "CHAR 8-bit value");
                cursor++;
            }
            if ((options & MtefRecord.OPT_CHAR_ENC_CHAR_16) != 0) {
                requireBytes(data, cursor, 2, "CHAR 16-bit value");
                cursor += 2;
            }
        } else if (tag == MtefRecord.TMPL) {
            requireBytes(data, cursor, 1, "TMPL options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor = skipStrictNudge(data, cursor);
            }
            requireBytes(data, cursor, 3, "TMPL selector, variation, options");
            cursor++; // selector
            int variation = Byte.toUnsignedInt(data[cursor++]);
            if ((variation & 0x80) != 0) {
                requireBytes(data, cursor, 1, "TMPL extended variation");
                cursor++;
            }
            cursor++; // template-specific options
            opensContainer = true;
        } else if (tag == MtefRecord.PILE) {
            requireBytes(data, cursor, 1, "PILE options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor = skipStrictNudge(data, cursor);
            }
            requireBytes(data, cursor, 2, "PILE alignment");
            cursor += 2;
            opensContainer = true;
        } else if (tag == MtefRecord.MATRIX) {
            requireBytes(data, cursor, 1, "MATRIX options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor = skipStrictNudge(data, cursor);
            }
            requireBytes(data, cursor, 5, "MATRIX header");
            cursor += 3;
            int rows = Byte.toUnsignedInt(data[cursor++]);
            int columns = Byte.toUnsignedInt(data[cursor++]);
            int partitions = packedPartitionBytes(rows + 1) + packedPartitionBytes(columns + 1);
            requireBytes(data, cursor, partitions, "MATRIX partitions");
            cursor += partitions;
            opensContainer = true;
        } else if (tag == MtefRecord.EMBELL) {
            requireBytes(data, cursor, 1, "EMBELL options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            if ((options & MtefRecord.OPT_NUDGE) != 0) {
                cursor = skipStrictNudge(data, cursor);
            }
            requireBytes(data, cursor, 1, "EMBELL type");
            cursor++;
            opensContainer = true;
        } else if (tag == MtefRecord.RULER) {
            requireBytes(data, cursor, 1, "RULER count");
            int stops = Byte.toUnsignedInt(data[cursor++]);
            requireBytes(data, cursor, stops * 3, "RULER stops");
            cursor += stops * 3;
        } else if (tag == MtefRecord.FONT_STYLE_DEF || tag == MtefRecord.FONT_DEF
                || tag == MtefRecord.ENCODING_DEF) {
            requireBytes(data, cursor, 1, "definition index");
            cursor++;
            cursor = skipStrictNullTerminated(data, cursor, "definition name");
        } else if (tag == MtefRecord.SIZE) {
            requireBytes(data, cursor, 1, "SIZE value");
            int value = Byte.toUnsignedInt(data[cursor++]);
            if (value == 101) {
                requireBytes(data, cursor, 2, "SIZE 16-bit value");
                cursor += 2;
            } else if (value == 100) {
                requireBytes(data, cursor, 3, "SIZE 32-bit value");
                cursor += 3;
            } else if (cursor < data.length && Byte.toUnsignedInt(data[cursor]) == 0x50) {
                requireBytes(data, cursor, 2, "SIZE encoded value");
                cursor += 2;
            }
        } else if (tag == MtefRecord.COLOR) {
            requireBytes(data, cursor, 1, "COLOR index");
            cursor++;
        } else if (tag == MtefRecord.COLOR_DEF) {
            requireBytes(data, cursor, 1, "COLOR_DEF options");
            int options = Byte.toUnsignedInt(data[cursor++]);
            int componentCount = (options & 0x01) != 0 ? 4 : 3;
            requireBytes(data, cursor, componentCount * 2, "COLOR_DEF components");
            cursor += componentCount * 2;
            if ((options & 0x04) != 0) {
                cursor = skipStrictNullTerminated(data, cursor, "COLOR_DEF name");
            }
        } else if (tag == MtefRecord.EQN_PREFS) {
            requireBytes(data, cursor, 1, "EQN_PREFS options");
            cursor++;
            cursor = skipStrictDimensionArray(data, cursor, "EQN_PREFS sizes");
            cursor = skipStrictDimensionArray(data, cursor, "EQN_PREFS spacing");
            requireBytes(data, cursor, 1, "EQN_PREFS style count");
            int styleCount = Byte.toUnsignedInt(data[cursor++]);
            for (int index = 0; index < styleCount; index++) {
                requireBytes(data, cursor, 1, "EQN_PREFS style font");
                int font = Byte.toUnsignedInt(data[cursor++]);
                if (font != 0) {
                    requireBytes(data, cursor, 1, "EQN_PREFS character style");
                    cursor++;
                }
            }
        } else if (tag >= MtefRecord.FUTURE) {
            requireBytes(data, cursor, 1, "FUTURE length");
            int length = Byte.toUnsignedInt(data[cursor++]);
            requireBytes(data, cursor, length, "FUTURE payload");
            cursor += length;
        } else {
            throw new IllegalStateException("unknown MTEF record tag " + tag + " at byte " + offset);
        }
        return new StrictRecord(tag, opensContainer, cursor);
    }

    private static int strictRecordStart(byte[] data) {
        if (data.length <= 5) {
            return 0;
        }
        if (data.length >= 10
                && "DSMT".equals(new String(data, 5, 4, StandardCharsets.US_ASCII))) {
            int markerEnd = strictIndexOfZero(data, 10);
            if (markerEnd < 0) {
                throw new IllegalStateException("unterminated DSMT application marker");
            }
            int cursor = markerEnd + 1;
            if (data[9] == '7') {
                byte[] label = "TeX Input Language\0".getBytes(StandardCharsets.US_ASCII);
                int labelOffset = strictIndexOf(data, label, cursor);
                if (labelOffset >= 0 && labelOffset - cursor <= 8) {
                    int sourceEnd = strictIndexOfZero(data, labelOffset + label.length);
                    if (sourceEnd < 0) {
                        throw new IllegalStateException("unterminated DSMT7 TeX source header");
                    }
                    return sourceEnd + 1;
                }
            }
            requireBytes(data, cursor, 1, "DSMT equation option");
            return cursor + 1;
        }
        int markerEnd = strictIndexOfZero(data, 5);
        if (markerEnd < 0) {
            throw new IllegalStateException("unterminated legacy MTEF application marker");
        }
        return markerEnd + 1;
    }

    private static int skipStrictNudge(byte[] data, int cursor) {
        requireBytes(data, cursor, 2, "nudge");
        // In the compact form a zero coordinate itself is encoded as 0x80.
        // The six-byte extended sentinel is therefore the pair 0x80,0x80.
        int length = Byte.toUnsignedInt(data[cursor]) == 0x80
            && Byte.toUnsignedInt(data[cursor + 1]) == 0x80 ? 6 : 2;
        requireBytes(data, cursor, length, "nudge");
        return cursor + length;
    }

    private static int skipStrictDimensionArray(byte[] data, int cursor, String label) {
        requireBytes(data, cursor, 1, label + " count");
        int count = Byte.toUnsignedInt(data[cursor++]);
        int nibble = 0;
        int completed = 0;
        while (completed < count) {
            requireBytes(data, cursor + nibble / 2, 1, label + " dimensions");
            int packed = Byte.toUnsignedInt(data[cursor + nibble / 2]);
            int value = (nibble & 1) == 0 ? packed >>> 4 : packed & 0x0F;
            nibble++;
            if (value == 0x0F) {
                completed++;
            }
        }
        return cursor + (nibble + 1) / 2;
    }

    private static int skipStrictNullTerminated(byte[] data, int cursor, String label) {
        int end = strictIndexOfZero(data, cursor);
        if (end < 0) {
            throw new IllegalStateException(label + " is not null terminated");
        }
        return end + 1;
    }

    private static int strictIndexOfZero(byte[] data, int from) {
        for (int index = Math.max(0, from); index < data.length; index++) {
            if (data[index] == 0) {
                return index;
            }
        }
        return -1;
    }

    private static int strictIndexOf(byte[] data, byte[] needle, int from) {
        outer:
        for (int index = Math.max(0, from); index + needle.length <= data.length; index++) {
            for (int part = 0; part < needle.length; part++) {
                if (data[index + part] != needle[part]) {
                    continue outer;
                }
            }
            return index;
        }
        return -1;
    }

    private static int packedPartitionBytes(int size) {
        return size <= 0 ? 0 : (size + 3) / 4;
    }

    private static void requireBytes(byte[] data, int offset, int length, String label) {
        if (offset < 0 || length < 0 || offset > data.length - length) {
            throw new IllegalStateException(label + " overruns MTEF at byte " + offset
                + " (need " + length + ", length " + data.length + ")");
        }
    }

    private static List<ExpectedFormula> expectedFormulas(Path request,
                                                          int documentIndex,
                                                          Path document,
                                                          List<Map<String, Object>> failures,
                                                          Counters counters) throws Exception {
        if (!Files.isRegularFile(request)) {
            throw new IllegalStateException("request file is missing");
        }
        counters.sourceRepairCount += XscSourceReplacementRepairs
            .occurrenceCount(documentIndex);
        JsonNode root = XscSourceReplacementRepairs.apply(
            documentIndex, JSON.readTree(request.toFile()));
        boolean compact = root.path("paper").path("compactLayout").asBoolean(false);
        List<ExpectedFormula> formulas = new ArrayList<>();
        JsonNode sections = root.path("sections");
        for (int sectionIndex = 0; sectionIndex < sections.size(); sectionIndex++) {
            JsonNode questions = sections.get(sectionIndex).path("questions");
            for (int questionIndex = 0; questionIndex < questions.size(); questionIndex++) {
                JsonNode question = questions.get(questionIndex);
                String base = "sections[" + sectionIndex + "].questions[" + questionIndex + "]";
                addExpectedFormulas(formulas, question.path("content").asText(null), base + ".content",
                    documentIndex, document, failures, counters);
                JsonNode options = question.path("options");
                for (int optionIndex = 0; optionIndex < options.size(); optionIndex++) {
                    addExpectedFormulas(formulas, options.get(optionIndex).path("content").asText(null),
                        base + ".options[" + optionIndex + "].content",
                        documentIndex, document, failures, counters);
                }
                addExpectedFormulas(formulas, question.path("knowledgePoint").asText(null),
                    base + ".knowledgePoint", documentIndex, document, failures, counters);
                if (compact) {
                    addExpectedFormulas(formulas, question.path("difficulty").asText(null),
                        base + ".difficulty", documentIndex, document, failures, counters);
                    addExpectedFormulas(formulas, joinedTags(question.path("tags")), base + ".tags",
                        documentIndex, document, failures, counters);
                } else {
                    addExpectedFormulas(formulas, joinedTags(question.path("tags")), base + ".tags",
                        documentIndex, document, failures, counters);
                    addExpectedFormulas(formulas, question.path("difficulty").asText(null),
                        base + ".difficulty", documentIndex, document, failures, counters);
                }
                addExpectedFormulas(formulas, question.path("analyze").asText(null), base + ".analyze",
                    documentIndex, document, failures, counters);
                addExpectedFormulas(formulas, question.path("solution").asText(null), base + ".solution",
                    documentIndex, document, failures, counters);
                addExpectedFormulas(formulas, question.path("correct").asText(null), base + ".correct",
                    documentIndex, document, failures, counters);
            }
        }
        return List.copyOf(formulas);
    }

    private static String joinedTags(JsonNode tags) {
        if (!tags.isArray() || tags.isEmpty()) {
            return null;
        }
        List<String> values = new ArrayList<>();
        tags.forEach(tag -> values.add(tag.asText()));
        return String.join("、", values);
    }

    private static void addExpectedFormulas(List<ExpectedFormula> formulas,
                                            String content,
                                            String sourceLocation,
                                            int documentIndex,
                                            Path document,
                                            List<Map<String, Object>> failures,
                                            Counters counters) {
        if (content == null || content.isBlank()) {
            return;
        }
        List<String> rawDelimitedFormulas = fallbackDelimitedFormulas(content);
        List<ManifestFormula> manifestFormulas;
        try {
            List<LaTeXParser.ContentSegment> parsedSegments = LATEX_PARSER.parseHtml(content).stream()
                .filter(LaTeXParser.ContentSegment::isMath)
                .toList();
            List<ManifestFormula> parsed = new ArrayList<>(parsedSegments.size());
            for (int formulaIndex = 0; formulaIndex < parsedSegments.size(); formulaIndex++) {
                LaTeXParser.ContentSegment segment = parsedSegments.get(formulaIndex);
                String formulaLocation = sourceLocation + "#math" + (formulaIndex + 1);
                XscSourceReplacementRepairs.RepairInfo repair = XscSourceReplacementRepairs
                    .repairFor(documentIndex, formulaLocation).orElse(null);
                String sourceLatex = repair != null
                    ? repair.sourceLatex()
                    : formulaIndex < rawDelimitedFormulas.size()
                        ? rawDelimitedFormulas.get(formulaIndex)
                        : segment.rawText();
                parsed.add(new ManifestFormula(segment, segment.rawText(), sourceLatex, null));
            }
            manifestFormulas = List.copyOf(parsed);
        } catch (Exception exception) {
            // Do not lose every earlier/later formula in a field just because one source
            // formula is invalid. Recover delimiters, then diagnose formulas individually.
            List<ManifestFormula> recovered = new ArrayList<>();
            for (int formulaIndex = 0; formulaIndex < rawDelimitedFormulas.size(); formulaIndex++) {
                String fallbackLatex = rawDelimitedFormulas.get(formulaIndex);
                String formulaLocation = sourceLocation + "#math" + (formulaIndex + 1);
                String sourceLatex = XscSourceReplacementRepairs
                    .repairFor(documentIndex, formulaLocation)
                    .map(XscSourceReplacementRepairs.RepairInfo::sourceLatex)
                    .orElse(fallbackLatex);
                try {
                    LaTeXParser.ContentSegment segment = LATEX_PARSER
                        .parseText("$" + fallbackLatex + "$").stream()
                        .filter(LaTeXParser.ContentSegment::isMath)
                        .findFirst()
                        .orElseThrow(() -> new IllegalStateException(
                            "formula delimiter was not recovered"));
                    recovered.add(new ManifestFormula(
                        segment, segment.rawText(), sourceLatex, null));
                } catch (Exception formulaFailure) {
                    recovered.add(new ManifestFormula(null,
                        stripRawFormulaAnnotations(fallbackLatex), sourceLatex, formulaFailure));
                }
            }
            manifestFormulas = List.copyOf(recovered);
        }
        int fieldFormula = 0;
        for (ManifestFormula manifest : manifestFormulas) {
            fieldFormula++;
            int ordinal = formulas.size() + 1;
            String latex = manifest.latex();
            String formulaLocation = sourceLocation + "#math" + fieldFormula;
            XscSourceReplacementRepairs.RepairInfo appliedRepair =
                XscSourceReplacementRepairs.repairFor(documentIndex, formulaLocation)
                    .orElse(null);
            if (appliedRepair != null) {
                if (appliedRepair.formulaOrdinal() != ordinal) {
                    throw new IllegalStateException("xsc source repair ordinal drift at "
                        + formulaLocation + ": catalog=" + appliedRepair.formulaOrdinal()
                        + ", actual=" + ordinal);
                }
                if (!appliedRepair.repairedLatex().equals(latex)) {
                    throw new IllegalStateException("xsc source repair output drift at "
                        + formulaLocation);
                }
                counters.formulaSourceRepairCount++;
                if (appliedRepair.sourceLatex().indexOf('\uFFFD') >= 0) {
                    counters.replacementFormulaRepairCount++;
                }
                addSourceRepairApplication(counters.sourceRepairApplications,
                    documentIndex, document, ordinal, formulaLocation, appliedRepair);
            }
            List<MtefRecordNormalizer.CanonicalRecord> expectedMtefRecords = null;
            Exception expectedMtefFailure = null;
            if (manifest.segment() != null) {
                try {
                    LaTeXParser.FormulaStyleHints styleHints = manifest.segment().styleHints()
                        .withSourceMetrics(manifest.segment().metrics());
                    MtefWriter.WriteReport expectedReport = MTEF_WRITER.get().writeWithReport(
                        manifest.segment().ast(), styleHints);
                    validateNormalizedMtef(expectedReport.normalization());
                    validateMtefStream(expectedReport.bytes());
                    expectedMtefRecords = expectedReport.normalization().records();
                } catch (Exception exception) {
                    expectedMtefFailure = exception;
                }
            }
            ExpectedFormula expected = new ExpectedFormula(ordinal,
                formulaLocation, latex, traceId(ordinal, latex),
                manifest.sourceLatex(), expectedMtefRecords);
            formulas.add(expected);
            if (manifest.sourceLatex().indexOf('\uFFFD') >= 0) {
                if (manifest.segment() != null && latex.indexOf('\uFFFD') < 0) {
                    counters.recoveredSourceReplacementCount++;
                    addSourceReplacementRecovery(counters.sourceReplacementRecoveries,
                        documentIndex, document, expected);
                } else {
                    addFailure(failures, documentIndex, document, expected,
                        "REQUEST_SOURCE_REPLACEMENT_CHARACTER",
                        replacementCharacterDetail(manifest.sourceLatex()));
                }
            }
            if (manifest.parseFailure() != null) {
                addFailure(failures, documentIndex, document, expected,
                    "REQUEST_FORMULA_UNPARSABLE", exceptionChain(manifest.parseFailure()));
            } else if (expectedMtefFailure != null) {
                addFailure(failures, documentIndex, document, expected,
                    "REQUEST_EXPECTED_MTEF_INVALID", exceptionChain(expectedMtefFailure));
            }
        }
    }

    private static void addSourceReplacementRecovery(
            List<Map<String, Object>> recoveries,
            int documentIndex,
            Path document,
            ExpectedFormula formula) {
        Map<String, Object> recovery = new LinkedHashMap<>();
        recovery.put("documentIndex", documentIndex);
        recovery.put("document", document.toString());
        recovery.put("formulaOrdinal", formula.ordinal());
        recovery.put("trace", formula.trace());
        recovery.put("sourceLocation", formula.sourceLocation());
        recovery.put("sourceLatex", formula.sourceLatex());
        recovery.put("repairedLatex", formula.latex());
        recovery.put("replacementCharacters", replacementCharacterDetail(formula.sourceLatex()));
        recoveries.add(recovery);
    }

    private static void addSourceRepairApplication(
            List<Map<String, Object>> applications,
            int documentIndex,
            Path document,
            int formulaOrdinal,
            String sourceLocation,
            XscSourceReplacementRepairs.RepairInfo repair) {
        Map<String, Object> application = new LinkedHashMap<>();
        application.put("documentIndex", documentIndex);
        application.put("document", document.toString());
        application.put("formulaOrdinal", formulaOrdinal);
        application.put("sourceLocation", sourceLocation);
        application.put("kind", repair.kind());
        application.put("reason", repair.reason());
        application.put("sourceLatex", repair.sourceLatex());
        application.put("repairedLatex", repair.repairedLatex());
        applications.add(application);
    }

    private static String replacementCharacterDetail(String sourceLatex) {
        List<Integer> utf16Indices = new ArrayList<>();
        List<Integer> codePointIndices = new ArrayList<>();
        int codePointIndex = 0;
        for (int offset = 0; offset < sourceLatex.length();) {
            int codePoint = sourceLatex.codePointAt(offset);
            if (codePoint == 0xFFFD) {
                utf16Indices.add(offset);
                codePointIndices.add(codePointIndex);
            }
            offset += Character.charCount(codePoint);
            codePointIndex++;
        }
        return "source LaTeX contains U+FFFD REPLACEMENT CHARACTER; utf16Indices="
            + utf16Indices + ", codePointIndices=" + codePointIndices
            + "; source encoding/content must be repaired before acceptance";
    }

    private static List<String> fallbackDelimitedFormulas(String html) {
        String text = Jsoup.parse(html).body().text()
            .replace("\\[", "$$").replace("\\]", "$$")
            .replace("\\(", "$").replace("\\)", "$");
        List<String> formulas = new ArrayList<>();
        int searchFrom = 0;
        while (searchFrom < text.length()) {
            int start = nextUnescapedDollar(text, searchFrom);
            if (start < 0) {
                break;
            }
            int delimiterLength = start + 1 < text.length() && text.charAt(start + 1) == '$' ? 2 : 1;
            int end = closingDollar(text, start + delimiterLength, delimiterLength);
            if (end < 0) {
                break;
            }
            String latex = text.substring(start + delimiterLength, end).trim();
            latex = latex.replaceFirst("^\\$\\s*(?=\\\\pwmetrics\\b)", "");
            formulas.add(stripRawFormulaAnnotations(latex));
            searchFrom = end + delimiterLength;
        }
        return List.copyOf(formulas);
    }

    private static int nextUnescapedDollar(String text, int from) {
        for (int index = Math.max(from, 0); index < text.length(); index++) {
            if (text.charAt(index) == '$' && !isEscaped(text, index)) {
                return index;
            }
        }
        return -1;
    }

    private static int closingDollar(String text, int cursor, int delimiterLength) {
        int braces = 0;
        while (cursor < text.length()) {
            char current = text.charAt(cursor);
            if (current == '\\') {
                cursor += Math.min(2, text.length() - cursor);
            } else if (current == '{') {
                braces++;
                cursor++;
            } else if (current == '}' && braces > 0) {
                braces--;
                cursor++;
            } else if (current == '$' && braces == 0
                    && (delimiterLength == 1
                        || cursor + 1 < text.length() && text.charAt(cursor + 1) == '$')) {
                return cursor;
            } else {
                cursor++;
            }
        }
        return -1;
    }

    private static boolean isEscaped(String text, int index) {
        int slashes = 0;
        for (int cursor = index - 1; cursor >= 0 && text.charAt(cursor) == '\\'; cursor--) {
            slashes++;
        }
        return (slashes & 1) != 0;
    }

    private static String stripRawFormulaAnnotations(String latex) {
        String normalized = latex == null ? "" : latex.trim();
        normalized = TRACE_METRICS.matcher(normalized).replaceFirst("");
        normalized = normalized
            .replaceAll("\\\\pwmetrics\\{[^}]+}\\s*", "")
            .replaceAll("\\\\pwstyle\\{[^}]*}\\s*", "")
            .trim();
        return normalized;
    }

    private static String traceId(int ordinal, String latex) {
        try {
            String normalized = latex == null ? "" : latex.replace('\u00A0', ' ');
            normalized = TRACE_METRICS.matcher(normalized).replaceFirst("");
            normalized = TRACE_STYLE.matcher(normalized).replaceFirst("");
            normalized = EDGE_SPACE.matcher(normalized).replaceAll("");
            normalized = INNER_SPACE.matcher(normalized).replaceAll(" ");
            byte[] digest = MessageDigest.getInstance("SHA-256")
                .digest(normalized.getBytes(StandardCharsets.UTF_8));
            return "pwf:" + ordinal + "-" + HexFormat.of().formatHex(digest, 0, 8);
        } catch (Exception exception) {
            throw new IllegalStateException("Could not compute formula trace", exception);
        }
    }

    private static void validateOptionalSourceReportCount(int documentIndex,
                                                          int expectedCount,
                                                          List<Map<String, Object>> failures,
                                                          Counters counters) {
        Path sourceReport = SOURCE_REPORTS.resolve(Integer.toString(documentIndex))
            .resolve(documentIndex + ".report.json");
        if (!Files.isRegularFile(sourceReport)) {
            return;
        }
        try {
            int sourceCount = JSON.readTree(sourceReport.toFile()).path("equations").size();
            counters.sourceReportsChecked++;
            if (sourceCount != expectedCount) {
                addFailure(failures, documentIndex, null, null, "SOURCE_REPORT_COUNT_MISMATCH",
                    sourceReport + " equations=" + sourceCount + ", request formulas=" + expectedCount);
            }
        } catch (Exception exception) {
            addFailure(failures, documentIndex, null, null, "SOURCE_REPORT_INVALID",
                sourceReport + ": " + exceptionChain(exception));
        }
    }

    private static Map<String, RelationshipInfo> relationships(ZipFile zip) throws Exception {
        Document document = parseXml(zip, "word/_rels/document.xml.rels");
        NodeList nodes = document.getElementsByTagNameNS(PACKAGE_REL_NS, "Relationship");
        Map<String, RelationshipInfo> relationships = new LinkedHashMap<>();
        for (int index = 0; index < nodes.getLength(); index++) {
            Element relationship = (Element) nodes.item(index);
            String id = relationship.getAttribute("Id");
            String target = relationship.getAttribute("Target");
            boolean external = "External".equalsIgnoreCase(relationship.getAttribute("TargetMode"));
            String entryName = external ? target : resolveWordTarget(target);
            RelationshipInfo previous = relationships.put(id, new RelationshipInfo(
                id, relationship.getAttribute("Type"), target, entryName, external));
            if (previous != null) {
                throw new IllegalStateException("duplicate relationship id " + id);
            }
        }
        return relationships;
    }

    private static String resolveWordTarget(String target) {
        String normalized = target.replace('\\', '/');
        if (normalized.startsWith("/")) {
            normalized = normalized.substring(1);
        } else {
            normalized = "word/" + normalized;
        }
        Deque<String> parts = new ArrayDeque<>();
        for (String part : normalized.split("/")) {
            if (part.isEmpty() || ".".equals(part)) {
                continue;
            }
            if ("..".equals(part)) {
                if (parts.isEmpty()) {
                    throw new IllegalStateException("relationship target escapes package root: " + target);
                }
                parts.removeLast();
            } else {
                parts.addLast(part);
            }
        }
        return String.join("/", parts);
    }

    private static Document parseXml(ZipFile zip, String entryName) throws Exception {
        ZipEntry entry = zip.getEntry(entryName);
        if (entry == null) {
            throw new IllegalStateException("missing DOCX part " + entryName);
        }
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
        factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
        factory.setXIncludeAware(false);
        factory.setExpandEntityReferences(false);
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        try (InputStream input = zip.getInputStream(entry)) {
            return factory.newDocumentBuilder().parse(input);
        }
    }

    private static Set<String> relationshipIds(Map<String, RelationshipInfo> relationships,
                                               String type,
                                               String suffix) {
        Set<String> ids = new LinkedHashSet<>();
        relationships.values().stream()
            .filter(relationship -> type.equals(relationship.type()))
            .filter(relationship -> suffix == null
                || relationship.entryName().toLowerCase(Locale.ROOT).endsWith(suffix))
            .map(RelationshipInfo::id)
            .forEach(ids::add);
        return ids;
    }

    private static Set<String> zipEntries(ZipFile zip, String prefix, String suffix) {
        Set<String> names = new LinkedHashSet<>();
        zip.stream()
            .filter(entry -> !entry.isDirectory())
            .map(ZipEntry::getName)
            .filter(name -> name.startsWith(prefix))
            .filter(name -> name.toLowerCase(Locale.ROOT).endsWith(suffix))
            .forEach(names::add);
        return names;
    }

    private static void validateSetEquality(List<Map<String, Object>> failures,
                                            int documentIndex,
                                            Path document,
                                            String code,
                                            String leftLabel,
                                            Set<String> left,
                                            String rightLabel,
                                            Set<String> right) {
        if (left.equals(right)) {
            return;
        }
        Set<String> onlyLeft = new LinkedHashSet<>(left);
        onlyLeft.removeAll(right);
        Set<String> onlyRight = new LinkedHashSet<>(right);
        onlyRight.removeAll(left);
        addFailure(failures, documentIndex, document, null, code,
            leftLabel + " only=" + onlyLeft + "; " + rightLabel + " only=" + onlyRight);
    }

    private static List<IndexedPath> indexedFiles(Path directory,
                                                  Pattern pattern,
                                                  String suffix) throws Exception {
        List<IndexedPath> files = new ArrayList<>();
        try (var paths = Files.list(directory)) {
            for (Path path : paths.filter(Files::isRegularFile)
                    .filter(candidate -> candidate.getFileName().toString()
                        .toLowerCase(Locale.ROOT).endsWith(suffix)).toList()) {
                Matcher matcher = pattern.matcher(path.getFileName().toString());
                if (matcher.find()) {
                    files.add(new IndexedPath(Integer.parseInt(matcher.group(1)), path));
                }
            }
        }
        return files.stream().sorted(Comparator.comparingInt(IndexedPath::index)).toList();
    }

    private static void validateCompleteIndexSet(List<IndexedPath> files,
                                                 int expectedDocuments,
                                                 String label,
                                                 List<Map<String, Object>> failures) {
        Map<Integer, List<Path>> byIndex = new LinkedHashMap<>();
        files.forEach(item -> byIndex.computeIfAbsent(item.index(), ignored -> new ArrayList<>())
            .add(item.path()));
        for (int index = 1; index <= expectedDocuments; index++) {
            List<Path> paths = byIndex.getOrDefault(index, List.of());
            if (paths.size() != 1) {
                addFailure(failures, index, null, null, "DOCUMENT_INDEX_COVERAGE",
                    label + " index " + index + " has " + paths.size() + " files: " + paths);
            }
        }
        byIndex.keySet().stream().filter(index -> index < 1 || index > expectedDocuments)
            .forEach(index -> addFailure(failures, index, null, null,
                "DOCUMENT_INDEX_OUT_OF_RANGE", label + " has unexpected index " + index));
    }

    private static byte[] read(DocumentEntry entry) throws Exception {
        try (DocumentInputStream input = new DocumentInputStream(entry)) {
            return input.readAllBytes();
        }
    }

    private static void addFailure(List<Map<String, Object>> failures,
                                   int documentIndex,
                                   Path document,
                                   ExpectedFormula formula,
                                   String code,
                                   String detail) {
        Map<String, Object> failure = new LinkedHashMap<>();
        failure.put("documentIndex", documentIndex);
        failure.put("document", document == null ? null : document.toString());
        failure.put("formulaOrdinal", formula == null ? null : formula.ordinal());
        failure.put("trace", formula == null ? null : formula.trace());
        failure.put("sourceLocation", formula == null ? null : formula.sourceLocation());
        failure.put("latex", formula == null ? null : formula.latex());
        failure.put("sourceLatex", formula == null ? null : formula.sourceLatex());
        failure.put("code", code);
        failure.put("detail", detail);
        failure.put("repro", "-Dpaperword.acceptance.xscEmfPlus=true"
            + " -Dxsc.emfplus.startDocument=" + documentIndex
            + " -Dxsc.emfplus.endDocument=" + documentIndex);
        failures.add(failure);
    }

    private static String exceptionChain(Throwable throwable) {
        List<String> chain = new ArrayList<>();
        Set<Throwable> seen = new HashSet<>();
        Throwable cursor = throwable;
        while (cursor != null && seen.add(cursor)) {
            chain.add(cursor.getClass().getName() + ": " + cursor.getMessage());
            cursor = cursor.getCause();
        }
        return String.join(" <- ", chain);
    }

    private record IndexedPath(int index, Path path) {
    }

    private record ManifestFormula(LaTeXParser.ContentSegment segment,
                                   String latex,
                                   String sourceLatex,
                                   Exception parseFailure) {
    }

    private record ExpectedFormula(
        int ordinal,
        String sourceLocation,
        String latex,
        String trace,
        String sourceLatex,
        List<MtefRecordNormalizer.CanonicalRecord> expectedMtefRecords
    ) {
        private ExpectedFormula {
            expectedMtefRecords = expectedMtefRecords == null
                ? null : List.copyOf(expectedMtefRecords);
        }
    }

    private record RelationshipInfo(String id, String type, String target,
                                    String entryName, boolean external) {
    }

    private record OlePayload(byte[] mtef, String mtefSha256) {
    }

    private record StrictRecord(int tag, boolean opensContainer, int nextOffset) {
    }

    private record DocumentResult(Counters counters,
                                  Map<String, Object> summary,
                                  List<Map<String, Object>> failures) {
    }

    private static final class Counters {
        private int expectedFormulaCount;
        private int objectCount;
        private int traceMatchedCount;
        private int oleRelationshipCount;
        private int oleCount;
        private int validOleCount;
        private int mtefParsedCount;
        private int mtefBalancedCount;
        private int mtefStructureComparableCount;
        private int mtefStructureMatchedCount;
        private int emfRelationshipCount;
        private int emfCount;
        private int validEmfPlusDualCount;
        private int sourceRepairCount;
        private int formulaSourceRepairCount;
        private int replacementFormulaRepairCount;
        private final List<Map<String, Object>> sourceRepairApplications = new ArrayList<>();
        private int recoveredSourceReplacementCount;
        private final List<Map<String, Object>> sourceReplacementRecoveries = new ArrayList<>();
        private int sourceReportsChecked;

        private void add(Counters other) {
            expectedFormulaCount += other.expectedFormulaCount;
            objectCount += other.objectCount;
            traceMatchedCount += other.traceMatchedCount;
            oleRelationshipCount += other.oleRelationshipCount;
            oleCount += other.oleCount;
            validOleCount += other.validOleCount;
            mtefParsedCount += other.mtefParsedCount;
            mtefBalancedCount += other.mtefBalancedCount;
            mtefStructureComparableCount += other.mtefStructureComparableCount;
            mtefStructureMatchedCount += other.mtefStructureMatchedCount;
            emfRelationshipCount += other.emfRelationshipCount;
            emfCount += other.emfCount;
            validEmfPlusDualCount += other.validEmfPlusDualCount;
            sourceRepairCount += other.sourceRepairCount;
            formulaSourceRepairCount += other.formulaSourceRepairCount;
            replacementFormulaRepairCount += other.replacementFormulaRepairCount;
            sourceRepairApplications.addAll(other.sourceRepairApplications);
            recoveredSourceReplacementCount += other.recoveredSourceReplacementCount;
            sourceReplacementRecoveries.addAll(other.sourceReplacementRecoveries);
            sourceReportsChecked += other.sourceReportsChecked;
        }
    }
}
