package com.lz.paperword.core.mtef;

import com.lz.paperword.core.latex.LaTeXParser;
import com.lz.paperword.core.ole.OlePackager;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MathTypeDocxComparatorTest {

    @TempDir
    Path tempDir;

    @Test
    void validatesOleHeadersAndComparesNormalizedFormulaStructure() throws Exception {
        Path x = writeMinimalDocx("x", "x.docx");
        Path y = writeMinimalDocx("y", "y.docx");
        MathTypeDocxComparator comparator = new MathTypeDocxComparator();

        MathTypeDocxComparator.ComparisonReport identical = comparator.compare(x, x);
        assertTrue(identical.passesExactStructureGate());
        assertEquals(1, identical.validGeneratedOleCount());
        assertEquals(1, identical.rawMtefEqualCount());

        MathTypeDocxComparator.ComparisonReport different = comparator.compare(x, y);
        assertFalse(different.passesExactStructureGate());
        assertEquals(0, different.normalizedStructureEqualCount());
    }

    @Test
    void packagesMathType7ContainerMetadataAndPinnedNativeHeader() throws Exception {
        byte[] mtef = new MtefWriter().write(new LaTeXParser().parseLaTeX("\\Bbb x"));
        byte[] ole = new OlePackager().packageOle(mtef);

        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(ole))) {
            byte[] compObj = readStream((DocumentEntry) fs.getRoot().getEntry("\u0001CompObj"));
            byte[] nativeStream = readStream((DocumentEntry) fs.getRoot().getEntry("Equation Native"));
            assertTrue(new String(compObj, StandardCharsets.ISO_8859_1)
                .contains("MathType 7.0 Equation"));
            assertEquals("1c0000000200f1c3", java.util.HexFormat.of().formatHex(nativeStream, 0, 8));
            assertEquals(mtef.length,
                ByteBuffer.wrap(nativeStream, 8, 4).order(ByteOrder.LITTLE_ENDIAN).getInt());
            assertEquals("000000005c75e6080d3bdd000c001f08",
                java.util.HexFormat.of().formatHex(nativeStream, 12, 28));
        }

        byte[] ordinaryMtef = new MtefWriter().write(new LaTeXParser().parseLaTeX("x"));
        byte[] ordinaryOle = new OlePackager().packageOle(ordinaryMtef);
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(ordinaryOle))) {
            byte[] compObj = readStream((DocumentEntry) fs.getRoot().getEntry("\u0001CompObj"));
            assertTrue(new String(compObj, StandardCharsets.ISO_8859_1)
                .contains("MathType 7.0 Equation"));
        }
    }

    private byte[] readStream(DocumentEntry entry) throws Exception {
        try (DocumentInputStream input = new DocumentInputStream(entry)) {
            return input.readAllBytes();
        }
    }

    private Path writeMinimalDocx(String latex, String name) throws Exception {
        byte[] mtef = new MtefWriter().write(new LaTeXParser().parseLaTeX(latex));
        byte[] ole = new OlePackager().packageOle(mtef);
        Path path = tempDir.resolve(name);
        try (ZipOutputStream zip = new ZipOutputStream(Files.newOutputStream(path))) {
            zip.putNextEntry(new ZipEntry("word/embeddings/oleObject1.bin"));
            zip.write(ole);
            zip.closeEntry();
        }
        return path;
    }
}
