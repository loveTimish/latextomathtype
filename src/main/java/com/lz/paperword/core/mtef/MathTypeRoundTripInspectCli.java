package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.poifs.filesystem.*;
import org.w3c.dom.*;
import org.xml.sax.SAXException;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.security.MessageDigest;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/** Read-only, fail-closed observations of two DOCX MathType snapshots. No rendering or repair. */
public final class MathTypeRoundTripInspectCli {
    static final int MAX_ENTRIES = 4096, MAX_ENTRY = 16 * 1024 * 1024;
    static final long MAX_TOTAL = 128L * 1024 * 1024;
    static final int MAX_DEPTH = 64, MAX_RECORDS = 100_000, MAX_NODES = 500_000;
    private static final String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final String R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String O = "urn:schemas-microsoft-com:office:office";
    private static final String V = "urn:schemas-microsoft-com:vml";
    private static final String REL = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static final HexFormat HEX = HexFormat.of();

    public enum Status { OBSERVED_PRESERVED, OBSERVED_CHANGED, INCONCLUSIVE }
    public record Layer(Status status, String beforeSha256, String afterSha256, String note) {}
    public record StreamInfo(int length, String sha256) {}
    public record BinaryRecord(int offset, int tag, String hex) {}
    public record MtefInfo(boolean supported, String error, String headerSha256, String sequenceSha256,
                           String preferencesSha256, List<BinaryRecord> records) {}
    public record OleInfo(boolean valid, String error, String sha256, String rootClassId,
                          String rootModifiedFiletimeHex, Map<String, StreamInfo> streams,
                          String equationNativeSha256, MtefInfo mtef) {}
    public record WmfInfo(boolean supported, String error, String rawSha256, String effectiveSha256,
                          int recordCount, List<Integer> unknownFunctions, int ignoredPaddingBytes) {}
    public record FormulaInfo(int index, String ownerPart, String shapeId, String contextSha256,
                              String olePart, String previewPart, String frameSha256,
                              OleInfo ole, WmfInfo preview, List<String> issues) {}
    public record Inspection(String input, boolean valid, String packageSha256,
                             List<FormulaInfo> formulas, Map<String, String> contextHashes,
                             List<String> issues) {}
    public record FormulaComparison(Integer beforeIndex, Integer afterIndex, String matchedBy,
                                    Status status, Layer equationNative, Layer mtefSequence,
                                    Layer preferences, Layer otherStreams, Layer frame,
                                    Layer previewEffectiveRecords, Layer previewRaw,
                                    boolean containerBytesEqual, List<String> notes) {}
    public record ComparisonReport(int schemaVersion, Status status, Inspection before, Inspection after,
                                   Layer documentContext, List<FormulaComparison> objects,
                                   String pixelComparison, String semanticConclusion, List<String> issues) {}

    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("usage: MathTypeRoundTripInspectCli <before.docx> <after.docx> <NEW-report.json>");
            System.exit(2);
            return;
        }
        try {
            Path before = Path.of(args[0]), after = Path.of(args[1]), output = Path.of(args[2]);
            ComparisonReport report = compare(before, after);
            writeReport(report, output, before, after);
            System.out.printf("%s objects=%d/%d report=%s%n", report.status(),
                report.before().formulas().size(), report.after().formulas().size(), output.toAbsolutePath());
            System.exit(switch (report.status()) {
                case OBSERVED_PRESERVED -> 0;
                case OBSERVED_CHANGED -> 1;
                case INCONCLUSIVE -> 2;
            });
        } catch (Exception ex) {
            System.err.println("Inspection/report error: " + ex.getMessage());
            System.exit(2);
        }
    }

    /** Output is create-new only: existing files, input aliases, hardlinks and symlinks are never overwritten. */
    public static void writeReport(ComparisonReport report, Path output, Path... inputs) throws IOException {
        Path absolute = output.toAbsolutePath().normalize();
        for (Path input : inputs) {
            if (absolute.equals(input.toAbsolutePath().normalize())
                || (Files.exists(absolute) && Files.isSameFile(absolute, input))) {
                throw new IOException("report must not overwrite an input or input alias");
            }
        }
        try (OutputStream out = Files.newOutputStream(absolute,
                StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE, LinkOption.NOFOLLOW_LINKS)) {
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT).writeValue(out, report);
        }
    }

    public static ComparisonReport compare(Path before, Path after) {
        Inspection a = inspect(before), b = inspect(after);
        List<String> issues = new ArrayList<>();
        Layer context = layer(hashMap(a.contextHashes()), hashMap(b.contextHashes()), a.valid() && b.valid(),
            "Complete conservative Word XML context; only documented volatile metadata and relationship/shape identifiers are omitted");
        List<FormulaComparison> pairs = new ArrayList<>();
        Set<Integer> used = new HashSet<>();
        for (FormulaInfo x : a.formulas()) {
            String by = "stable unique shapeId";
            FormulaInfo y = uniqueMatch(x.shapeId(), a.formulas(), b.formulas(), used, true);
            if (y == null) {
                by = "unique paragraph text context";
                y = uniqueMatch(x.contextSha256(), a.formulas(), b.formulas(), used, false);
            }
            if (y == null && a.formulas().size() == 1 && b.formulas().size() == 1) {
                by = "single object in each document";
                y = b.formulas().get(0);
            }
            if (y == null) {
                pairs.add(unmatched(x.index(), null));
                issues.add("No unambiguous counterpart for before object " + x.index());
            } else {
                used.add(y.index());
                pairs.add(compareFormula(x, y, by));
            }
        }
        for (FormulaInfo y : b.formulas()) if (!used.contains(y.index())) pairs.add(unmatched(null, y.index()));
        boolean countChanged = a.formulas().size() != b.formulas().size();
        if (countChanged) issues.add("Observed object count changed");
        if (a.formulas().isEmpty() || b.formulas().isEmpty()) issues.add("No comparable MathType objects in one or both inputs");
        boolean incomplete = !a.valid() || !b.valid() || !a.issues().isEmpty() || !b.issues().isEmpty()
            || !issues.isEmpty() || pairs.stream().anyMatch(p -> p.status() == Status.INCONCLUSIVE);
        boolean changed = countChanged || context.status() == Status.OBSERVED_CHANGED
            || pairs.stream().anyMatch(p -> p.status() == Status.OBSERVED_CHANGED);
        Status status = changed ? Status.OBSERVED_CHANGED : incomplete ? Status.INCONCLUSIVE : Status.OBSERVED_PRESERVED;
        // An unreadable package cannot establish object deletion or document changes.
        if (!a.valid() || !b.valid()) status = Status.INCONCLUSIVE;
        return new ComparisonReport(1, status, a, b, context, List.copyOf(pairs),
            "NOT_MEASURED: no complete WMF/Word renderer or pixel comparison is used",
            "No mathematical-equivalence or native-editor round-trip guarantee is inferred from hashes or record normalization", issues);
    }

    private static FormulaInfo uniqueMatch(String key, List<FormulaInfo> a, List<FormulaInfo> b,
                                           Set<Integer> used, boolean shape) {
        if (key == null || key.isEmpty()) return null;
        if (a.stream().filter(f -> key.equals(shape ? f.shapeId() : f.contextSha256())).count() != 1) return null;
        List<FormulaInfo> candidates = b.stream().filter(f -> key.equals(shape ? f.shapeId() : f.contextSha256())).toList();
        return candidates.size() == 1 && !used.contains(candidates.get(0).index()) ? candidates.get(0) : null;
    }

    private static FormulaComparison unmatched(Integer a, Integer b) {
        Layer unknown = layer(null, null, false, "Object correspondence is missing or ambiguous; ZIP names and ordinals are not matching keys");
        return new FormulaComparison(a, b, "unmatched", Status.INCONCLUSIVE, unknown, unknown, unknown,
            unknown, unknown, unknown, unknown, false, List.of("No comparison inferred"));
    }

    private static FormulaComparison compareFormula(FormulaInfo a, FormulaInfo b, String by) {
        OleInfo x = a.ole(), y = b.ole();
        boolean validOle = x != null && y != null && x.valid() && y.valid();
        boolean mtef = validOle && x.mtef().supported() && y.mtef().supported();
        Layer nativeData = layer(x == null ? null : x.equationNativeSha256(), y == null ? null : y.equationNativeSha256(),
            validOle, "Exact complete Equation Native stream including header and preferences; a byte change alone is not a mathematical change");
        Layer sequence = layer(x == null || x.mtef() == null ? null : x.mtef().sequenceSha256(),
            y == null || y.mtef() == null ? null : y.mtef().sequenceSha256(), mtef,
            "Lossless ordered MTEF record bytes including coordinates, font, size, color and END records; excludes only separately reported EQN_PREFS");
        Layer prefs = layer(x == null || x.mtef() == null ? null : x.mtef().preferencesSha256(),
            y == null || y.mtef() == null ? null : y.mtef().preferencesSha256(), mtef, "Exact EQN_PREFS records in encounter order");
        Layer streams = layer(x == null ? null : otherStreams(x), y == null ? null : otherStreams(y), validOle,
            "Every other CFB stream, stream name/length and root CLSID; storage timestamps and allocation layout are reported separately");
        Layer frame = layer(a.frameSha256(), b.frameSha256(), a.issues().isEmpty() && b.issues().isEmpty(),
            "Full shape/run/paragraph/ancestor context including width, height, position, crop, hidden state, opacity and style references");
        WmfInfo u = a.preview(), v = b.preview();
        Layer preview = layer(u == null ? null : u.effectiveSha256(), v == null ? null : v.effectiveSha256(),
            u != null && v != null && u.supported() && v.supported(),
            "WMF effective record bytes; excludes only FaceName bytes after first NUL and odd EXTTEXTOUT string alignment byte. Record differences are not proof of pixel differences");
        Layer raw = layer(u == null ? null : u.rawSha256(), v == null ? null : v.rawSha256(),
            u != null && v != null, "Exact preview bytes (informational; not visual equivalence)");
        List<Layer> essential = List.of(nativeData, sequence, prefs, streams, frame, preview);
        Status status = essential.stream().anyMatch(l -> l.status() == Status.OBSERVED_CHANGED) ? Status.OBSERVED_CHANGED
            : essential.stream().anyMatch(l -> l.status() == Status.INCONCLUSIVE) ? Status.INCONCLUSIVE : Status.OBSERVED_PRESERVED;
        List<String> notes = new ArrayList<>();
        notes.addAll(a.issues()); notes.addAll(b.issues());
        if (!mtef) notes.add("MTEF parse incomplete or unsupported: semantic interpretation remains unknown even if native bytes match");
        if (validOle && !Objects.equals(x.sha256(), y.sha256()) && Objects.equals(x.streams(), y.streams()))
            notes.add("CFB bytes changed while every stream stayed equal; root timestamp/allocation metadata is separate from native data");
        return new FormulaComparison(a.index(), b.index(), by, status, nativeData, sequence, prefs, streams,
            frame, preview, raw, x != null && y != null && Objects.equals(x.sha256(), y.sha256()), notes);
    }

    private static String otherStreams(OleInfo ole) {
        Map<String, Object> values = new TreeMap<>(ole.streams());
        values.remove("Equation Native"); values.put("@rootClassId", ole.rootClassId());
        return hashMap(values);
    }
    private static Layer layer(String a, String b, boolean valid, String note) {
        return new Layer(!valid || a == null || b == null ? Status.INCONCLUSIVE
            : a.equals(b) ? Status.OBSERVED_PRESERVED : Status.OBSERVED_CHANGED, a, b, note);
    }

    public static Inspection inspect(Path input) {
        List<FormulaInfo> formulas = new ArrayList<>();
        List<String> issues = new ArrayList<>();
        Map<String, String> context = new TreeMap<>();
        String packageHash = null;
        try {
            if (Files.size(input) > MAX_TOTAL) throw new IOException("compressed package exceeds 128 MiB limit");
            Map<String, byte[]> parts = readPackage(input);
            packageHash = sha(Files.readAllBytes(input));
            Map<String, Document> xml = new TreeMap<>();
            for (Map.Entry<String, byte[]> entry : parts.entrySet()) {
                if (entry.getKey().endsWith(".xml") || entry.getKey().endsWith(".rels")) xml.put(entry.getKey(), parseXml(entry.getValue()));
            }
            if (!xml.containsKey("word/document.xml")) throw new IOException("word/document.xml is missing; this inspector supports transitional Word DOCX");
            Map<String, Map<String, Relationship>> rels = relationships(xml, parts, issues);
            Set<String> referencedOle = new HashSet<>();
            for (Map.Entry<String, Document> part : xml.entrySet()) {
                String name = part.getKey();
                if (!name.startsWith("word/") || !name.endsWith(".xml")) continue;
                NodeList objects = part.getValue().getElementsByTagNameNS(O, "OLEObject");
                Map<String, Relationship> edges = rels.getOrDefault(name, Map.of());
                for (int i = 0; i < objects.getLength(); i++) {
                    Element object = (Element) objects.item(i);
                    if (!object.getAttribute("ProgID").startsWith("Equation.")) {
                        issues.add("Unsupported non-MathType OLE object in " + name); continue;
                    }
                    List<String> objectIssues = new ArrayList<>();
                    Element wrapper = ancestor(object, W, "object");
                    if (wrapper == null) wrapper = (Element) object.getParentNode();
                    String id = object.getAttribute("ShapeID");
                    List<Element> shapes = descendants(wrapper, V, "shape").stream()
                        .filter(s -> id.equals(s.getAttribute("id"))).toList();
                    Element shape = shapes.size() == 1 ? shapes.get(0) : null;
                    if (shape == null) objectIssues.add("OLE ShapeID does not identify exactly one VML shape");
                    String olePart = target(edges, object.getAttributeNS(R, "id"), "/oleObject", objectIssues);
                    String previewPart = null;
                    if (shape != null) {
                        List<Element> images = descendants(shape, V, "imagedata");
                        if (images.size() != 1) objectIssues.add("VML shape does not have exactly one preview imagedata");
                        else previewPart = target(edges, images.get(0).getAttributeNS(R, "id"), "/image", objectIssues);
                    }
                    OleInfo ole = olePart == null ? null : inspectOle(parts.get(olePart));
                    if (olePart != null) referencedOle.add(olePart);
                    WmfInfo preview = previewPart == null ? null : inspectWmf(parts.get(previewPart));
                    Element paragraph = ancestor(object, W, "p");
                    String text = paragraph == null ? "" : paragraphText(paragraph);
                    String frame = frameContext(wrapper, edges);
                    formulas.add(new FormulaInfo(formulas.size(), name, id, text.isBlank() ? "" : sha(text.getBytes(StandardCharsets.UTF_8)),
                        olePart, previewPart, sha(frame.getBytes(StandardCharsets.UTF_8)), ole, preview, objectIssues));
                }
                context.put(name, sha(canonical(part.getValue().getDocumentElement(), edges, false).getBytes(StandardCharsets.UTF_8)));
            }
            for (String name : parts.keySet()) {
                if (name.startsWith("word/embeddings/") && !name.endsWith("/") && !referencedOle.contains(name)) issues.add("Unreferenced/unsupported embedding: " + name);
                // Unknown binary rendering dependencies are preserved conservatively, never silently dropped.
                if (name.startsWith("word/") && !name.endsWith("/") && !xml.containsKey(name) && !name.startsWith("word/media/")
                    && !name.startsWith("word/embeddings/")) context.put(name, sha(parts.get(name)));
            }
            if (formulas.size() > MAX_ENTRIES) throw new IOException("too many formula objects");
            return new Inspection(input.toAbsolutePath().toString(), true, packageHash, formulas, context, issues);
        } catch (Exception ex) {
            issues.add("Inspection stopped: " + ex.getClass().getSimpleName() + ": " + ex.getMessage());
            return new Inspection(input.toAbsolutePath().toString(), false, packageHash, formulas, context, issues);
        }
    }

    private record Relationship(String type, String target, boolean external, String digest) {}
    private static Map<String, Map<String, Relationship>> relationships(Map<String, Document> xml,
            Map<String, byte[]> parts, List<String> issues) throws IOException {
        Map<String, Map<String, Relationship>> all = new HashMap<>();
        for (Map.Entry<String, Document> e : xml.entrySet()) {
            String name = e.getKey();
            if (!name.endsWith(".rels")) continue;
            String owner;
            if (name.equals("_rels/.rels")) owner = "";
            else {
                int at = name.lastIndexOf("/_rels/");
                if (at < 0) throw new IOException("invalid relationship part path");
                owner = name.substring(0, at + 1) + name.substring(at + 7, name.length() - 5);
            }
            Map<String, Relationship> edges = new HashMap<>();
            for (Element r : descendants(e.getValue().getDocumentElement(), REL, "Relationship")) {
                String id = r.getAttribute("Id"), type = r.getAttribute("Type");
                boolean external = "External".equals(r.getAttribute("TargetMode"));
                String target = external ? r.getAttribute("Target") : resolveTarget(owner, r.getAttribute("Target"));
                if (external) issues.add("External relationship was not opened: " + owner + " [" + id + "]");
                else if (!parts.containsKey(target)) throw new IOException("missing relationship target: " + target);
                if (edges.put(id, new Relationship(type, target, external,
                    external ? sha(target.getBytes(StandardCharsets.UTF_8)) : sha(parts.get(target)))) != null)
                    throw new IOException("duplicate relationship Id");
            }
            all.put(owner, edges);
        }
        return all;
    }
    private static String target(Map<String, Relationship> edges, String id, String suffix, List<String> issues) {
        Relationship r = edges.get(id);
        if (r == null || r.external() || !r.type().equals(R + suffix)) {
            issues.add("Missing, external or wrong-type " + suffix + " relationship"); return null;
        }
        return r.target();
    }
    private static String resolveTarget(String owner, String target) throws IOException {
        try {
            URI uri = URI.create(target);
            if (uri.isAbsolute() || uri.getRawAuthority() != null || uri.getQuery() != null || uri.getFragment() != null)
                throw new IOException("non-package internal relationship target");
            String path = uri.getPath();
            if (path == null || path.isEmpty() || path.contains("\\") || path.contains(":")) throw new IOException("invalid target path");
            String base = owner.contains("/") ? owner.substring(0, owner.lastIndexOf('/') + 1) : "";
            Deque<String> stack = new ArrayDeque<>();
            for (String segment : (path.startsWith("/") ? path.substring(1) : base + path).split("/")) {
                if (segment.equals("..")) { if (stack.isEmpty()) throw new IOException("relationship escapes package"); stack.removeLast(); }
                else if (!segment.equals(".") && !segment.isEmpty()) stack.add(segment);
            }
            String result = String.join("/", stack); validatePartName(result); return result;
        } catch (IllegalArgumentException ex) { throw new IOException("invalid relationship URI", ex); }
    }

    static Map<String, byte[]> readPackage(Path path) throws IOException {
        Map<String, byte[]> parts = new LinkedHashMap<>(); long total = 0; int count = 0;
        try (ZipFile zip = new ZipFile(path.toFile())) {
            Enumeration<? extends ZipEntry> entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                if (++count > MAX_ENTRIES) throw new IOException("ZIP exceeds 4096 entry limit");
                String name = entry.getName(); validatePartName(entry.isDirectory() ? name.substring(0, name.length() - 1) : name);
                if (entry.getSize() > MAX_ENTRY) throw new IOException("ZIP entry exceeds 16 MiB limit");
                if (parts.containsKey(name)) throw new IOException("duplicate ZIP entry");
                byte[] data;
                try (InputStream in = zip.getInputStream(entry)) { data = bounded(in, MAX_ENTRY); }
                total += data.length;
                if (total > MAX_TOTAL) throw new IOException("ZIP exceeds 128 MiB expanded total limit");
                parts.put(name, data);
            }
        }
        return parts;
    }
    private static void validatePartName(String name) throws IOException {
        if (name.isEmpty() || name.startsWith("/") || name.contains("\\") || name.contains(":") || name.indexOf('\0') >= 0)
            throw new IOException("unsafe ZIP part path");
        for (String segment : name.split("/", -1)) if (segment.equals("..") || segment.equals(".") || segment.isEmpty())
            throw new IOException("unsafe ZIP part path segment");
    }
    private static byte[] bounded(InputStream in, int max) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream(); byte[] buffer = new byte[8192]; int n;
        while ((n = in.read(buffer)) != -1) {
            if ((long) out.size() + n > max) throw new IOException("bounded read exceeded limit");
            out.write(buffer, 0, n);
        }
        return out.toByteArray();
    }
    private static Document parseXml(byte[] data) throws Exception {
        DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
        f.setNamespaceAware(true); f.setXIncludeAware(false); f.setExpandEntityReferences(false);
        f.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
        f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        f.setFeature("http://xml.org/sax/features/external-general-entities", false);
        f.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        f.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, ""); f.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        f.setAttribute("http://www.oracle.com/xml/jaxp/properties/maxElementDepth", MAX_DEPTH);
        var builder = f.newDocumentBuilder();
        builder.setEntityResolver((publicId, systemId) -> { throw new SAXException("external XML resource forbidden"); });
        Document result = builder.parse(new ByteArrayInputStream(data));
        countNodes(result, 0, new int[]{0}); return result;
    }
    private static void countNodes(Node node, int depth, int[] count) throws IOException {
        if (depth > MAX_DEPTH || ++count[0] > MAX_NODES) throw new IOException("XML nesting/node count limit");
        for (Node child = node.getFirstChild(); child != null; child = child.getNextSibling()) countNodes(child, depth + 1, count);
    }
    private static List<Element> descendants(Element parent, String ns, String local) {
        NodeList nodes = parent.getElementsByTagNameNS(ns, local); List<Element> result = new ArrayList<>();
        for (int i = 0; i < nodes.getLength(); i++) result.add((Element) nodes.item(i));
        return result;
    }
    private static Element ancestor(Node child, String ns, String local) {
        for (Node n = child.getParentNode(); n instanceof Element e; n = n.getParentNode())
            if (ns.equals(e.getNamespaceURI()) && local.equals(e.getLocalName())) return e;
        return null;
    }
    private static String paragraphText(Element p) {
        StringBuilder out = new StringBuilder();
        for (Element t : descendants(p, W, "t")) { out.append(t.getTextContent()); out.append('\u001f'); }
        return out.toString();
    }
    private static String frameContext(Element wrapper, Map<String, Relationship> edges) {
        StringBuilder out = new StringBuilder(canonical(wrapper, edges, false));
        for (Node n = wrapper; n instanceof Element; n = n.getParentNode()) {
            int position = 0;
            for (Node prev = n.getPreviousSibling(); prev != null; prev = prev.getPreviousSibling())
                if (prev instanceof Element && Objects.equals(prev.getNamespaceURI(), n.getNamespaceURI())
                    && Objects.equals(prev.getLocalName(), n.getLocalName())) position++;
            out.append("@position:").append(n.getNamespaceURI()).append('|').append(n.getLocalName()).append(':').append(position);
        }
        for (Node n = wrapper.getParentNode(); n instanceof Element e; n = n.getParentNode()) {
            out.append(canonical(e, edges, true));
            // Actual run position among siblings is retained; changing it can move or hide an equation.
            if (W.equals(e.getNamespaceURI()) && "p".equals(e.getLocalName())) out.append(canonical(e, edges, false));
        }
        return out.toString();
    }
    private static String canonical(Node node, Map<String, Relationship> edges, boolean propertiesOnly) {
        StringBuilder out = new StringBuilder(); canonical(node, edges, propertiesOnly, out, 0); return out.toString();
    }
    private static void canonical(Node node, Map<String, Relationship> edges, boolean propertiesOnly,
                                  StringBuilder out, int depth) {
        if (depth > MAX_DEPTH) throw new IllegalArgumentException("XML canonical depth limit");
        if (node instanceof Element e) {
            String ns = Objects.toString(e.getNamespaceURI(), ""), local = e.getLocalName();
            if (W.equals(ns) && (local.equals("rsids") || local.equals("qFormat") || local.equals("proofErr"))) return;
            if (W.equals(ns) && (local.equals("bookmarkStart") || local.equals("bookmarkEnd")) && isGoBack(e)) return;
            out.append('<').append(ns).append('|').append(local);
            List<String> attrs = new ArrayList<>(); NamedNodeMap map = e.getAttributes();
            for (int i = 0; i < map.getLength(); i++) {
                Attr attr = (Attr) map.item(i); String ans = Objects.toString(attr.getNamespaceURI(), "");
                String key = attr.getLocalName() == null ? attr.getName() : attr.getLocalName();
                if (XMLConstants.XMLNS_ATTRIBUTE_NS_URI.equals(ans) || (W.equals(ans) && (key.startsWith("rsid") || key.equals("qFormat")))) continue;
                if ((V.equals(ns) && (key.equals("id") || key.equals("spid")))
                    || (O.equals(ns) && local.equals("OLEObject") && Set.of("ShapeID", "ObjectID").contains(key))) continue;
                String value = attr.getValue();
                if (R.equals(ans)) {
                    Relationship r = edges.get(value);
                    boolean formulaEdge = (V.equals(ns) && local.equals("imagedata") && ancestor(e, W, "object") != null) || (O.equals(ns) && local.equals("OLEObject"));
                    value = r == null ? "UNRESOLVED:" + value : r.type() + (formulaEdge ? "" : ":" + r.digest());
                }
                attrs.add(ans + '|' + key + '=' + value.length() + ':' + value);
            }
            Collections.sort(attrs); for (String attr : attrs) out.append(' ').append(attr);
            out.append('>');
            for (Node c = e.getFirstChild(); c != null; c = c.getNextSibling()) {
                if (!propertiesOnly || c instanceof Element ce && (ce.getLocalName().endsWith("Pr") || ce.getLocalName().equals("sectPr")))
                    canonical(c, edges, false, out, depth + 1);
            }
            out.append("</>");
        } else if ((node.getNodeType() == Node.TEXT_NODE || node.getNodeType() == Node.CDATA_SECTION_NODE)
            && (!node.getNodeValue().isBlank() || node.getParentNode() instanceof Element pe
                && W.equals(pe.getNamespaceURI()) && Set.of("t", "instrText", "delText", "delInstrText").contains(pe.getLocalName()))) {
            out.append('#').append(node.getNodeValue().length()).append(':').append(node.getNodeValue());
        }
    }
    private static boolean isGoBack(Element e) {
        if ("bookmarkStart".equals(e.getLocalName())) return "_GoBack".equals(e.getAttributeNS(W, "name"));
        String id = e.getAttributeNS(W, "id");
        for (Element start : descendants(e.getOwnerDocument().getDocumentElement(), W, "bookmarkStart"))
            if (id.equals(start.getAttributeNS(W, "id")) && "_GoBack".equals(start.getAttributeNS(W, "name"))) return true;
        return false;
    }

    static OleInfo inspectOle(byte[] data) {
        Map<String, StreamInfo> streams = new TreeMap<>();
        String raw = data == null ? null : sha(data), rootClass = null, modified = null, nativeSha = null;
        MtefInfo mtef = new MtefInfo(false, "Equation Native is absent", null, null, null, List.of());
        try {
            require(data != null && data.length >= 512 && data.length <= MAX_ENTRY, "invalid CFB length");
            require(HEX.formatHex(data, 0, 8).equals("d0cf11e0a1b11ae1"), "invalid CFB magic");
            int shift = u16(data, 30), major = u16(data, 26);
            require(u16(data, 28) == 0xfffe && ((major == 3 && shift == 9) || (major == 4 && shift == 12))
                && u16(data, 32) == 6, "unsupported CFB header");
            int sector = 1 << shift, sectors = data.length / sector - 1;
            require(data.length % sector == 0 && u32(data, 44) <= sectors && u32(data, 72) <= sectors
                && u32(data, 64) <= sectors, "CFB sector/count limit");
            long firstDirectory = u32(data, 48);
            require(firstDirectory < sectors, "invalid CFB directory sector");
            int rootAt = Math.toIntExact((firstDirectory + 1) * sector);
            require(data[rootAt + 66] == 5, "CFB root directory missing");
            modified = HEX.formatHex(data, rootAt + 108, rootAt + 116);
            preflightCfbDirectory(data, sector, sectors);
            Map<String, byte[]> payloads = new TreeMap<>();
            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(data))) {
                rootClass = fs.getRoot().getStorageClsid().toString();
                collectStreams(fs.getRoot(), "", streams, payloads, 0, new long[]{0, 0});
            }
            byte[] nativeBytes = payloads.get("Equation Native");
            require(nativeBytes != null, "Equation Native missing");
            nativeSha = sha(nativeBytes);
            require(nativeBytes.length >= 28, "truncated Equation Native header");
            int headerSize = u16(nativeBytes, 0);
            require(headerSize >= 28 && headerSize <= nativeBytes.length
                && u32(nativeBytes, 8) == nativeBytes.length - headerSize, "invalid Equation Native cbHdr/cbObject");
            mtef = inspectMtef(Arrays.copyOfRange(nativeBytes, headerSize, nativeBytes.length));
            return new OleInfo(true, "", raw, rootClass, modified, streams, nativeSha, mtef);
        } catch (Exception ex) {
            return new OleInfo(false, ex.getMessage(), raw, rootClass, modified, streams, nativeSha, mtef);
        }
    }
    /** Bound directory trees and declared stream sizes before POI traverses any attacker-controlled tree. */
    private static void preflightCfbDirectory(byte[] data, int sector, int sectors) throws IOException {
        int fatCount = Math.toIntExact(u32(data, 44));
        List<Integer> fatSectors = new ArrayList<>(); Set<Integer> seenFat = new HashSet<>();
        for (int p = 76; p < 512; p += 4) addFatSector(u32(data, p), sectors, fatSectors, seenFat);
        long dif = u32(data, 68); int difCount = Math.toIntExact(u32(data, 72)); Set<Long> seenDif = new HashSet<>();
        for (int i = 0; i < difCount; i++) {
            require(dif < sectors && seenDif.add(dif), "CFB DIFAT cycle/out-of-bounds");
            int start = Math.toIntExact((dif + 1) * sector);
            for (int p = start; p < start + sector - 4; p += 4) addFatSector(u32(data, p), sectors, fatSectors, seenFat);
            dif = u32(data, start + sector - 4);
        }
        require(fatSectors.size() == fatCount, "CFB FAT count mismatch");
        long[] fat = new long[Math.multiplyExact(fatCount, sector / 4)]; int index = 0;
        for (int sec : fatSectors) for (int p = (sec + 1) * sector; p < (sec + 2) * sector; p += 4) fat[index++] = u32(data, p);
        long next = u32(data, 48); Set<Long> seenDirectory = new HashSet<>(); ByteArrayOutputStream directory = new ByteArrayOutputStream();
        while (next != 0xfffffffeL) {
            require(next < sectors && next < fat.length && seenDirectory.add(next), "CFB directory chain cycle/out-of-bounds");
            require(directory.size() + sector <= MAX_ENTRIES * 128, "CFB directory entry limit");
            directory.write(data, Math.toIntExact((next + 1) * sector), sector); next = fat[(int) next];
        }
        byte[] properties = directory.toByteArray();
        for (int p = 0; p < properties.length; p += 128) {
            int type = properties[p + 66] & 255;
            require(type == 0 || type == 1 || type == 2 || type == 5, "unsupported CFB property type");
            if (type == 0) continue;
            int nameBytes = u16(properties, p + 64);
            require(nameBytes >= 2 && nameBytes <= 64 && (nameBytes & 1) == 0, "CFB property name length");
            if (type == 2 || type == 5) require(u32(properties, p + 124) == 0 && u32(properties, p + 120) <= MAX_ENTRY, "CFB declared stream size limit");
        }
        Set<Integer> visited = new HashSet<>(); validateCfbTree(properties, 0, 0, visited);
        for (int p = 0; p < properties.length; p += 128)
            require(properties[p + 66] == 0 || visited.contains(p / 128), "unreachable CFB property");
    }
    private static void addFatSector(long value, int sectors, List<Integer> fat, Set<Integer> seen) throws IOException {
        if (value == 0xffffffffL) return;
        require(value < sectors && seen.add((int) value), "CFB invalid/duplicate FAT sector"); fat.add((int) value);
    }
    private static void validateCfbTree(byte[] properties, long index, int depth, Set<Integer> visited) throws IOException {
        if (index == 0xffffffffL) return;
        require(depth <= MAX_DEPTH && index < properties.length / 128 && visited.add((int) index), "CFB directory tree cycle/depth/index limit");
        int p = (int) index * 128, type = properties[p + 66] & 255;
        require(type != 0, "CFB tree references empty property");
        validateCfbTree(properties, u32(properties, p + 68), depth + 1, visited);
        validateCfbTree(properties, u32(properties, p + 72), depth + 1, visited);
        if (type == 1 || type == 5) validateCfbTree(properties, u32(properties, p + 76), depth + 1, visited);
    }

    private static void collectStreams(DirectoryEntry directory, String prefix, Map<String, StreamInfo> streams,
            Map<String, byte[]> payloads, int depth, long[] budget) throws IOException {
        require(depth <= MAX_DEPTH, "CFB directory depth limit");
        for (Entry e : directory) {
            require(++budget[0] <= MAX_ENTRIES, "CFB entry count limit");
            String name = prefix + e.getName();
            if (e instanceof DirectoryEntry child) {
                require(!streams.containsKey(name + "/"), "duplicate CFB path");
                streams.put(name + "/", new StreamInfo(0, sha(child.getStorageClsid().toString().getBytes(StandardCharsets.UTF_8))));
                collectStreams(child, name + "/", streams, payloads, depth + 1, budget);
            } else if (e instanceof DocumentEntry document) {
                require(document.getSize() >= 0 && document.getSize() <= MAX_ENTRY, "CFB stream size limit");
                budget[1] += document.getSize(); require(budget[1] <= MAX_ENTRY, "CFB cumulative stream limit");
                byte[] bytes;
                try (DocumentInputStream in = new DocumentInputStream(document)) { bytes = bounded(in, MAX_ENTRY); }
                require(bytes.length == document.getSize() && !streams.containsKey(name), "CFB truncated/duplicate stream");
                streams.put(name, new StreamInfo(bytes.length, sha(bytes))); payloads.put(name, bytes);
            }
        }
    }

    /** Lossless bounded v5 tokenizer, not the existing lossy MtefRecordNormalizer. */
    static MtefInfo inspectMtef(byte[] data) {
        List<BinaryRecord> records = new ArrayList<>(); String header = null;
        ByteArrayOutputStream sequence = new ByteArrayOutputStream(), prefs = new ByteArrayOutputStream();
        boolean unsupported = false;
        try {
            Cursor c = new Cursor(data); require(c.get() == 5, "unsupported MTEF version");
            c.skip(4); c.string(); c.get(); header = sha(Arrays.copyOfRange(data, 0, c.at));
            sequence.writeBytes(Arrays.copyOfRange(data, 0, c.at));
            int depth = 0; boolean end = false;
            while (c.at < data.length) {
                require(records.size() < MAX_RECORDS, "MTEF record count limit");
                int start = c.at, tag = c.get(), option = 0;
                if (tag >= 1 && tag <= 6) { option = c.get(); if ((option & 8) != 0) c.nudge(); }
                switch (tag) {
                    case 0 -> { if (depth > 0) depth--; else end = true; }
                    case 1 -> {
                        require((option & ~15) == 0, "unsupported LINE options");
                        if ((option & 4) != 0) c.skip(2);
                        if ((option & 2) != 0) c.ruler();
                        if ((option & 1) == 0) depth++;
                    }
                    case 2 -> {
                        require((option & ~63) == 0 && (option & 20) != 20, "unsupported CHAR options");
                        c.integer(); // Typeface is always present, even when MTCode is omitted.
                        if ((option & 32) == 0) c.skip(2);
                        if ((option & 4) != 0) c.skip(1);
                        if ((option & 16) != 0) c.skip(2);
                        if ((option & 1) != 0) depth++;
                    }
                    case 3 -> {
                        require((option & ~8) == 0, "unsupported TMPL options");
                        c.get(); int variation = c.get(); if ((variation & 128) != 0) c.get(); c.get(); depth++;
                    }
                    case 4 -> { require((option & ~10) == 0, "unsupported PILE options"); c.skip(2); if ((option & 2) != 0) c.ruler(); depth++; }
                    case 5 -> { require((option & ~8) == 0, "unsupported MATRIX options"); c.skip(3); int rows = c.get(), cols = c.get(); c.skip((rows + 4) / 4 + (cols + 4) / 4); depth++; }
                    case 6 -> { require((option & ~8) == 0, "unsupported EMBELL options"); c.get(); }
                    case 7 -> c.skip(c.get() * 3);
                    case 8 -> { c.integer(); c.get(); }
                    case 9 -> { int size = c.get(); if (size == 101) c.skip(2); else if (size == 100) c.skip(3); else { require(size <= 7, "unsupported SIZE logical size"); c.get(); } }
                    case 10, 11, 12, 13, 14 -> { }
                    case 15 -> c.integer();
                    case 16 -> { int options = c.get(); require((options & ~7) == 0, "unsupported COLOR_DEF options"); c.skip((options & 1) == 0 ? 6 : 8); if ((options & 4) != 0) c.string(); }
                    case 17 -> { c.integer(); c.string(); }
                    case 18 -> {
                        require(c.get() == 0, "unsupported EQN_PREFS options"); c.dimensions(); c.dimensions();
                        int count = c.get(); for (int i = 0; i < count; i++) if (c.integer() != 0) c.get();
                    }
                    case 19 -> c.string();
                    default -> {
                        require(tag >= 100, "unknown MTEF tag " + tag);
                        c.skip(c.integer()); unsupported = true;
                    }
                }
                require(depth <= MAX_DEPTH, "MTEF nesting limit");
                require(!end || c.at == data.length, "MTEF bytes after terminal END");
                byte[] raw = Arrays.copyOfRange(data, start, c.at);
                records.add(new BinaryRecord(start, tag, HEX.formatHex(raw)));
                (tag == 18 ? prefs : sequence).writeBytes(raw);
            }
            require(depth == 0 && !records.isEmpty() && records.get(records.size() - 1).tag() == 0, "MTEF missing END/unclosed list");
            return new MtefInfo(!unsupported, unsupported ? "Opaque future MTEF record retained; interpretation unsupported" : "",
                header, sha(sequence.toByteArray()), sha(prefs.toByteArray()), records);
        } catch (Exception ex) {
            return new MtefInfo(false, ex.getMessage(), header, null, null, records);
        }
    }
    private static final class Cursor {
        final byte[] data; int at;
        Cursor(byte[] data) { this.data = data; }
        int get() throws IOException { require(at < data.length, "truncated MTEF at " + at); return data[at++] & 255; }
        void skip(int n) throws IOException { require(n >= 0 && (long) at + n <= data.length, "truncated MTEF at " + at); at += n; }
        int integer() throws IOException { int n = get(); return n == 255 ? get() | (get() << 8) : n; }
        void string() throws IOException { int start = at; while (get() != 0) require(at - start <= 65535, "MTEF string limit"); }
        void nudge() throws IOException { int x = get(), y = get(); if (x == 128 && y == 128) skip(4); }
        void ruler() throws IOException { require(get() == 7, "MTEF embedded RULER tag missing"); skip(get() * 3); }
        void dimensions() throws IOException {
            int count = get(), nibble = 0, start = at;
            for (int i = 0; i < count; i++) {
                int unit = nibble(start, nibble++); require(unit <= 4, "unsupported preference dimension unit");
                int digits = 0;
                while (true) {
                    int n = nibble(start, nibble++); if (n == 15) break;
                    require(n <= 11 && ++digits <= 1024, "invalid preference dimension");
                }
            }
            at = start + (nibble + 1) / 2;
        }
        int nibble(int start, int n) throws IOException {
            require(start + n / 2 < data.length, "truncated preference dimensions");
            int v = data[start + n / 2] & 255; return (n & 1) == 0 ? v >>> 4 : v & 15;
        }
    }

    /** WMF record equality with precisely two permitted padding exclusions, never a pixel comparator. */
    static WmfInfo inspectWmf(byte[] data) {
        String raw = data == null ? null : sha(data); int count = 0, ignored = 0;
        Set<Integer> unknown = new TreeSet<>();
        try {
            require(data != null && data.length >= 24 && data.length <= MAX_ENTRY, "missing/truncated WMF");
            byte[] effective = data.clone(); int header = 0;
            if (u32(data, 0) == 0x9ac6cdd7L) {
                require(data.length >= 46 && u16(data, 14) > 0, "invalid placeable WMF header");
                int checksum = 0; for (int p = 0; p < 20; p += 2) checksum ^= u16(data, p);
                require(checksum == u16(data, 20), "invalid placeable checksum"); header = 22;
            }
            require((u16(data, header) == 1 || u16(data, header) == 2) && u16(data, header + 2) == 9
                && (u16(data, header + 4) == 0x100 || u16(data, header + 4) == 0x300)
                && u32(data, header + 6) * 2 == data.length - header, "unsupported/invalid WMF header");
            int offset = header + 18; boolean eof = false; long maxRecord = 0;
            while (offset < data.length) {
                require(++count <= MAX_RECORDS && offset + 6 <= data.length, "WMF record limit/truncated header");
                long words = u32(data, offset); require(words >= 3 && words <= (data.length - offset) / 2, "invalid/truncated WMF record");
                int end = offset + Math.toIntExact(words * 2), fn = u16(data, offset + 4), p = offset + 6, n = end - p;
                maxRecord = Math.max(maxRecord, words);
                Integer fixed = WMF_FIXED.get(fn);
                if (fixed != null) require(n == fixed || ((fn == 0x0102 || fn == 0x012e) && n == 4), "incorrect fixed WMF record length: " + fn);
                else if (fn == 0x02fb) {
                    require(n == 50, "invalid LOGFONT length (requires 32-byte FaceName array)");
                    int firstZero = p + 18; while (firstZero < p + 50 && data[firstZero] != 0) firstZero++;
                    require(firstZero < p + 50, "unterminated LOGFONT FaceName");
                    for (int at = firstZero + 1; at < p + 50; at++) { effective[at] = 0; ignored++; }
                } else if (fn == 0x0a32) {
                    require(n >= 8, "truncated EXTTEXTOUT"); int length = u16(data, p + 4), options = u16(data, p + 6);
                    require(length <= 32767, "negative EXTTEXTOUT length");
                    // Only rectangle clipping/opacity flags are supported; other extensions remain unknown.
                    if ((options & ~6) != 0) unknown.add(fn);
                    int text = p + 8 + ((options & 6) != 0 ? 8 : 0), paddedEnd = text + length + (length & 1);
                    require(paddedEnd <= end && (end == paddedEnd || end - paddedEnd == length * 2), "invalid EXTTEXTOUT string/Dx length");
                    if ((length & 1) != 0) { effective[text + length] = 0; ignored++; }
                } else if (fn == 0x0521) {
                    require(n >= 6, "truncated TEXTOUT"); int length = u16(data, p);
                    require(n == 6 + length + (length & 1), "invalid TEXTOUT length");
                    // TEXTOUT padding is deliberately not normalized in this finite policy.
                } else if (fn == 0x0324 || fn == 0x0325) {
                    require(n >= 2 && n == 2L + u16(data, p) * 4L, "invalid polygon point count");
                } else if (fn == 0x0538) {
                    require(n >= 2, "truncated POLYPOLYGON"); int polygons = u16(data, p); long points = 0;
                    require(2L + polygons * 2L <= n, "truncated POLYPOLYGON count array");
                    for (int i = 0; i < polygons; i++) points += u16(data, p + 2 + 2 * i);
                    require(n == 2L + polygons * 2L + points * 4L, "invalid POLYPOLYGON points");
                } else if (fn == 0x0626) {
                    require(n >= 4, "truncated ESCAPE"); int escape = u16(data, p), bytes = u16(data, p + 2);
                    require(n == 4 + bytes + (bytes & 1), "invalid ESCAPE byte count");
                    // Historical MFCOMMENT (15) framing is supported and ALL payload bytes are retained.
                    // No embedded metafile is interpreted or rendered. Other escapes remain unsupported.
                    if (escape != 15) unknown.add(fn);
                } else unknown.add(fn);
                offset = end;
                if (fn == 0) { eof = true; require(offset == data.length, "bytes after WMF EOF"); break; }
            }
            require(eof && maxRecord <= u32(data, header + 12), "missing WMF EOF/invalid maximum record size");
            return new WmfInfo(unknown.isEmpty(), unknown.isEmpty() ? "" : "Unsupported WMF functions/options retained verbatim",
                raw, sha(effective), count, List.copyOf(unknown), ignored);
        } catch (Exception ex) {
            return new WmfInfo(false, ex.getMessage(), raw, null, count, List.copyOf(unknown), ignored);
        }
    }
    private static final Map<Integer, Integer> WMF_FIXED = Map.ofEntries(
        Map.entry(0x0000, 0), Map.entry(0x001e, 0), Map.entry(0x0127, 2),
        Map.entry(0x0102, 2), Map.entry(0x0103, 2), Map.entry(0x0104, 2), Map.entry(0x0105, 2),
        Map.entry(0x0106, 2), Map.entry(0x0107, 2), Map.entry(0x0108, 2), Map.entry(0x012d, 2),
        Map.entry(0x012e, 2), Map.entry(0x01f0, 2), Map.entry(0x0201, 4), Map.entry(0x0209, 4),
        Map.entry(0x020a, 4), Map.entry(0x020b, 4), Map.entry(0x020c, 4), Map.entry(0x020d, 4),
        Map.entry(0x020e, 4), Map.entry(0x020f, 4), Map.entry(0x0211, 4), Map.entry(0x0213, 4),
        Map.entry(0x0214, 4), Map.entry(0x02fa, 10), Map.entry(0x02fc, 8),
        Map.entry(0x0410, 8), Map.entry(0x0412, 8), Map.entry(0x0415, 8), Map.entry(0x0416, 8),
        Map.entry(0x0418, 8), Map.entry(0x041b, 8), Map.entry(0x041f, 8), Map.entry(0x061c, 12),
        Map.entry(0x0817, 16), Map.entry(0x081a, 16), Map.entry(0x0830, 16));
    private static int u16(byte[] b, int at) { return (b[at] & 255) | ((b[at + 1] & 255) << 8); }
    private static long u32(byte[] b, int at) { return Integer.toUnsignedLong(u16(b, at) | (u16(b, at + 2) << 16)); }
    private static void require(boolean condition, String message) throws IOException { if (!condition) throw new IOException(message); }
    private static String sha(byte[] bytes) {
        try { return HEX.formatHex(MessageDigest.getInstance("SHA-256").digest(bytes)); }
        catch (Exception ex) { throw new IllegalStateException(ex); }
    }
    private static String hashMap(Map<?, ?> values) {
        try { return sha(new ObjectMapper().writeValueAsBytes(values)); }
        catch (IOException ex) { throw new IllegalStateException(ex); }
    }
}
