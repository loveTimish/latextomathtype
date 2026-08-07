package com.lz.paperword.core.mtef;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/** Reports the compound-file streams of every MathType OLE object in a DOCX. */
public final class MathTypeOleContainerInspectCli {

    private MathTypeOleContainerInspectCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.err.println("usage: MathTypeOleContainerInspectCli <input.docx> <out.json>");
            System.exit(2);
        }
        Path input = Path.of(args[0]).toAbsolutePath();
        Path output = Path.of(args[1]).toAbsolutePath();
        List<Map<String, Object>> objects = new ArrayList<>();
        try (ZipFile zip = new ZipFile(input.toFile())) {
            List<? extends ZipEntry> embeddings = zip.stream()
                .filter(entry -> entry.getName().startsWith("word/embeddings/")
                    && entry.getName().toLowerCase().endsWith(".bin"))
                .sorted(Comparator.comparing(ZipEntry::getName))
                .toList();
            for (ZipEntry embedding : embeddings) {
                byte[] ole = zip.getInputStream(embedding).readAllBytes();
                Map<String, Object> object = new LinkedHashMap<>();
                object.put("entry", embedding.getName());
                List<Map<String, Object>> streams = new ArrayList<>();
                try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(ole))) {
                    collectStreams(fs.getRoot(), "", streams);
                }
                object.put("streams", streams);
                objects.add(object);
            }
        }
        Map<String, Object> report = new LinkedHashMap<>();
        report.put("schemaVersion", 1);
        report.put("docx", input.toString());
        report.put("objects", objects);
        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT).writeValue(output.toFile(), report);
        System.out.printf("objects=%d report=%s%n", objects.size(), output);
    }

    private static void collectStreams(DirectoryEntry directory, String prefix,
                                       List<Map<String, Object>> streams) throws Exception {
        for (Entry entry : directory) {
            String name = prefix + entry.getName();
            if (entry instanceof DirectoryEntry child) {
                collectStreams(child, name + "/", streams);
            } else if (entry instanceof DocumentEntry document) {
                byte[] bytes;
                try (DocumentInputStream input = new DocumentInputStream(document)) {
                    bytes = input.readAllBytes();
                }
                Map<String, Object> stream = new LinkedHashMap<>();
                stream.put("name", name);
                stream.put("length", bytes.length);
                stream.put("sha256", HexFormat.of().formatHex(
                    MessageDigest.getInstance("SHA-256").digest(bytes)));
                stream.put("prefixHex", HexFormat.of().formatHex(bytes, 0, Math.min(bytes.length, 96)));
                streams.add(stream);
            }
        }
        streams.sort(Comparator.comparing(stream -> (String) stream.get("name")));
    }
}
