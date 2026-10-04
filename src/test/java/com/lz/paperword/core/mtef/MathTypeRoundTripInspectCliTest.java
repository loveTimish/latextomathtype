package com.lz.paperword.core.mtef;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.zip.*;

import static com.lz.paperword.core.mtef.MathTypeRoundTripInspectCli.Status.*;
import static org.junit.jupiter.api.Assertions.*;

/** Self-created fixtures only. No customer DOCX/OLE/WMF bytes are checked in. */
class MathTypeRoundTripInspectCliTest {
    @TempDir Path tmp;
    static final String NS = "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' "
        + "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
        + "xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:v='urn:schemas-microsoft-com:vml'";

    @Test void identicalKnownSnapshotsAreObservedPreservedWithoutPixelClaim() throws Exception {
        Path a = doc("a", paragraph(run("s1", "o1", "p1", "")), mtef('1'), wmf(0, 0, 0, 24, 20), "");
        var report = MathTypeRoundTripInspectCli.compare(a, a);
        assertEquals(OBSERVED_PRESERVED, report.status());
        assertTrue(report.pixelComparison().startsWith("NOT_MEASURED"));
        assertEquals(OBSERVED_PRESERVED, report.objects().get(0).equationNative().status());
    }

    @Test void relationshipIdsAndZipPartNumbersAreNotMatchingKeys() throws Exception {
        Path a = doc("a", paragraph(run("stable", "o1", "p1", "")), mtef('1'), wmf(0, 0, 0, 24, 20), "");
        Map<String, byte[]> entries = entries(a);
        entries.put("word/document.xml", document(paragraph(run("renamed", "o99", "p77", ""))));
        entries.put("word/_rels/document.xml.rels", rels("o99", "p77", "object900.bin", "image001.wmf", false));
        entries.put("word/embeddings/object900.bin", entries.remove("word/embeddings/object1.bin"));
        entries.put("word/media/image001.wmf", entries.remove("word/media/image1.wmf"));
        Path b = zip("renumbered", entries);
        var report = MathTypeRoundTripInspectCli.compare(a, b);
        assertEquals(OBSERVED_PRESERVED, report.status());
        assertEquals("unique paragraph text context", report.objects().get(0).matchedBy());
    }

    @Test void changedNumbersAndSizesRemainVisible() throws Exception {
        Path a = doc("a", paragraph(run("s", "o1", "p1", "")), mtef('1'), wmf(0, 0, 0, 24, 20), "");
        Path b = doc("b", paragraph(run("s", "o1", "p1", "")), mtef('2'), wmf(0, 0, 0, 24, 20), "");
        var r = MathTypeRoundTripInspectCli.compare(a, b);
        assertEquals(OBSERVED_CHANGED, r.objects().get(0).mtefSequence().status());
        byte[] sizeChanged = mtef('1'); sizeChanged[16] = 23;
        Path c = doc("c", paragraph(run("s", "o1", "p1", "")), sizeChanged, wmf(0, 0, 0, 24, 20), "");
        assertEquals(OBSERVED_CHANGED, MathTypeRoundTripInspectCli.compare(a, c).objects().get(0).mtefSequence().status());
    }

    @Test void losslessPreferencesAndNonNativeStreamsAreSeparate() throws Exception {
        byte[] a = concat(header(), new byte[]{18,0,0,0,0}, new byte[]{1,0,2,0,(byte)136,'1',0,0,0});
        byte[] b = concat(header(), new byte[]{18,0,1,0x21,0x2f,0,0}, new byte[]{1,0,2,0,(byte)136,'1',0,0,0});
        var x = MathTypeRoundTripInspectCli.inspectMtef(a); var y = MathTypeRoundTripInspectCli.inspectMtef(b);
        assertTrue(x.supported()); assertTrue(y.supported());
        assertEquals(x.sequenceSha256(), y.sequenceSha256()); assertNotEquals(x.preferencesSha256(), y.preferencesSha256());
        Path first = doc("stream1", paragraph(run("s", "o1", "p1", "")), a, wmf(0,0,0,24,20), "");
        Map<String,byte[]> parts = entries(first); parts.put("word/embeddings/object1.bin", ole(a, new byte[]{9}));
        Path second = zip("stream2", parts);
        assertEquals(OBSERVED_CHANGED, MathTypeRoundTripInspectCli.compare(first, second).objects().get(0).otherStreams().status());
    }

    @Test void onlyDocumentedWmfPaddingIsIgnored() {
        var a = MathTypeRoundTripInspectCli.inspectWmf(wmf(0,0,0,24,20));
        var b = MathTypeRoundTripInspectCli.inspectWmf(wmf(99,87,0,24,20));
        assertTrue(a.supported(), a.error()); assertTrue(b.supported(), b.error());
        assertNotEquals(a.rawSha256(), b.rawSha256()); assertEquals(a.effectiveSha256(), b.effectiveSha256());
        for (byte[] changed : List.of(wmf(0,0,0x0000ff,24,20), wmf(0,0,0,25,20), wmf(0,0,0,24,21)))
            assertNotEquals(a.effectiveSha256(), MathTypeRoundTripInspectCli.inspectWmf(changed).effectiveSha256());
        byte[] font = wmf(0,0,0,24,20); font[18+6+18] = 'Z';
        assertNotEquals(a.effectiveSha256(), MathTypeRoundTripInspectCli.inspectWmf(font).effectiveSha256());
    }

    @Test void hiddenRunOpacityGeometryAndStyleChangesAreObserved() throws Exception {
        String base = paragraph(run("s", "o1", "p1", ""));
        Path a = doc("base", base, mtef('1'), wmf(0,0,0,24,20), "");
        List<String> changes = List.of(base.replace("<w:r><w:object>", "<w:r><w:rPr><w:vanish/></w:rPr><w:object>"),
            base.replace("width:20pt", "width:21pt"), base.replace("height:10pt", "height:11pt"),
            base.replace("position:0", "position:2"), base.replace("opacity:1", "opacity:0"),
            base.replace("croptop='0'", "croptop='0.2'"),
            base.replace("<w:r><w:object>", "<w:r><w:rPr><w:rStyle w:val='Hidden'/></w:rPr><w:object>"));
        for (int i=0;i<changes.size();i++) {
            Path b = doc("changed"+i, changes.get(i), mtef('1'), wmf(0,0,0,24,20), "");
            assertEquals(OBSERVED_CHANGED, MathTypeRoundTripInspectCli.compare(a,b).status(), changes.get(i));
        }
        Path styled = doc("styles", base, mtef('1'), wmf(0,0,0,24,20), "<w:style w:styleId='Normal'><w:rPr><w:vanish/></w:rPr></w:style>");
        assertEquals(OBSERVED_CHANGED, MathTypeRoundTripInspectCli.compare(a,styled).documentContext().status());
    }

    @Test void whitespaceTextAndObjectSwapAreNotSilentlyLost() throws Exception {
        String objects = run("s1","o1","p1","")+run("s2","o1","p1","");
        Path a = doc("swapA", paragraph(objects),mtef('1'),wmf(0,0,0,24,20),"");
        Path b = doc("swapB",paragraph(run("s2","o1","p1","")+run("s1","o1","p1","")),mtef('1'),wmf(0,0,0,24,20),"");
        var report = MathTypeRoundTripInspectCli.compare(a,b);
        assertEquals(OBSERVED_CHANGED,report.status());
        assertTrue(report.objects().stream().allMatch(o->o.frame().status()==OBSERVED_CHANGED));
        Path c=doc("space1",paragraph("<w:r><w:t xml:space='preserve'> </w:t></w:r>"+objects),mtef('1'),wmf(0,0,0,24,20),"");
        Path d=doc("space2",paragraph("<w:r><w:t xml:space='preserve'>          </w:t></w:r>"+objects),mtef('1'),wmf(0,0,0,24,20),"");
        assertEquals(OBSERVED_CHANGED,MathTypeRoundTripInspectCli.compare(c,d).documentContext().status());
    }

    @Test void volatileBookmarksRsidAndGalleryMetadataDoNotChangeRenderingContext() throws Exception {
        String content=paragraph(run("s","o1","p1",""));
        Path a=doc("metadataA",content,mtef('1'),wmf(0,0,0,24,20),"<w:style w:styleId='Normal'/>");
        String changed=content.replace("<w:p>","<w:p w:rsidR='12345678'><w:bookmarkStart w:id='91' w:name='_GoBack'/><w:bookmarkEnd w:id='91'/>");
        Path b=doc("metadataB",changed,mtef('1'),wmf(0,0,0,24,20),"<w:style w:styleId='Normal' w:qFormat='1'><w:qFormat/></w:style>");
        assertEquals(OBSERVED_PRESERVED,MathTypeRoundTripInspectCli.compare(a,b).status());
    }

    @Test void countAndAmbiguousMatchingNeverUseZipOrdering() throws Exception {
        Path a=doc("countA",paragraph(run("duplicate","o1","p1","")+run("duplicate","o1","p1","")),mtef('1'),wmf(0,0,0,24,20),"");
        assertEquals(INCONCLUSIVE,MathTypeRoundTripInspectCli.compare(a,a).status());
        Path b=doc("countB",paragraph(run("s","o1","p1","")),mtef('1'),wmf(0,0,0,24,20),"");
        assertEquals(OBSERVED_CHANGED,MathTypeRoundTripInspectCli.compare(a,b).status());
    }

    @Test void unknownAndTruncatedDataRemainInconclusiveEvenWhenEqual() throws Exception {
        byte[] unknown=concat(header(),new byte[]{100,1,42,1,0,2,0,(byte)136,'1',0,0,0});
        var m=MathTypeRoundTripInspectCli.inspectMtef(unknown);
        assertFalse(m.supported()); assertNotNull(m.sequenceSha256());
        Path a=doc("future",paragraph(run("s","o1","p1","")),unknown,wmf(0,0,0,24,20),"");
        assertEquals(INCONCLUSIVE,MathTypeRoundTripInspectCli.compare(a,a).status());
        Path b=doc("futureChanged",paragraph(run("s","o1","p1","")).replace("width:20pt","width:30pt"),unknown,wmf(0,0,0,24,20),"");
        assertEquals(OBSERVED_CHANGED,MathTypeRoundTripInspectCli.compare(a,b).status());
        for(int i=0;i<mtef('1').length-1;i++) assertFalse(MathTypeRoundTripInspectCli.inspectMtef(Arrays.copyOf(mtef('1'),i)).supported());
        byte[] bad=wmf(0,0,0,24,20); ByteBuffer.wrap(bad).order(ByteOrder.LITTLE_ENDIAN).putInt(18,Integer.MAX_VALUE);
        assertFalse(MathTypeRoundTripInspectCli.inspectWmf(bad).supported());
        byte[] unk=wmf(0,0,0,24,20); unk[22]=(byte)0xff; unk[23]=(byte)0x7f;
        assertFalse(MathTypeRoundTripInspectCli.inspectWmf(unk).supported());
    }

    @Test void invalidNativeIsNotCalledMathematicalChange() throws Exception {
        Path a=doc("nativeA",paragraph(run("s","o1","p1","")),mtef('1'),wmf(0,0,0,24,20),"");
        Map<String,byte[]> parts=entries(a); parts.put("word/embeddings/object1.bin",new byte[]{1,2,3});
        Path b=zip("nativeB",parts);
        var r=MathTypeRoundTripInspectCli.compare(a,b);
        assertEquals(INCONCLUSIVE,r.objects().get(0).equationNative().status());
        assertEquals(INCONCLUSIVE,r.status());
    }

    @Test void xmlDtdExternalRelationshipsAndEscapingPathsFailClosed() throws Exception {
        Path a=doc("safe",paragraph(run("s","o1","p1","")),mtef('1'),wmf(0,0,0,24,20),"");
        Map<String,byte[]> parts=entries(a);
        parts.put("word/document.xml",("<!DOCTYPE a [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]>"+new String(document(paragraph("&xxe;")),StandardCharsets.UTF_8)).getBytes(StandardCharsets.UTF_8));
        assertFalse(MathTypeRoundTripInspectCli.inspect(zip("xxe",parts)).valid());
        parts=entries(a); parts.put("word/_rels/document.xml.rels",rels("o1","p1","https://example.invalid/a.bin","image1.wmf",true));
        assertEquals(INCONCLUSIVE,MathTypeRoundTripInspectCli.compare(zip("external",parts),zip("external2",parts)).status());
        parts=entries(a); parts.put("word/_rels/document.xml.rels",rels("o1","p1","../../../escape.bin","image1.wmf",false));
        assertFalse(MathTypeRoundTripInspectCli.inspect(zip("escape",parts)).valid());
        assertThrows(IOException.class,()->MathTypeRoundTripInspectCli.readPackage(zip("zipSlip",Map.of("../outside",new byte[0]))));
    }

    @Test void zipEntryCountAndExpandedEntryLimitsAreEnforced() throws Exception {
        Map<String,byte[]> many=new LinkedHashMap<>();
        for(int i=0;i<4097;i++) many.put("part"+i,new byte[0]);
        assertThrows(IOException.class,()->MathTypeRoundTripInspectCli.readPackage(zip("many",many)));
        assertThrows(IOException.class,()->MathTypeRoundTripInspectCli.readPackage(zip("large",Map.of("large",new byte[16*1024*1024+1]))));
    }

    @Test void cumulativeZipXmlDepthAndCfbCyclesAreBounded() throws Exception {
        byte[] chunk = new byte[16 * 1024 * 1024];
        Map<String, byte[]> parts = new LinkedHashMap<>();
        for (int i = 0; i < 9; i++) parts.put("chunk" + i, chunk);
        assertThrows(IOException.class, () -> MathTypeRoundTripInspectCli.readPackage(zip("cumulative", parts)));
        String deep = "<w:p>".repeat(70) + "</w:p>".repeat(70);
        assertFalse(MathTypeRoundTripInspectCli.inspect(zip("deep", Map.of("word/document.xml", document(deep)))).valid());
        byte[] cyclic = ole(mtef('1'), new byte[]{1});
        ByteBuffer cfb = ByteBuffer.wrap(cyclic).order(ByteOrder.LITTLE_ENDIAN);
        int sector = 1 << cfb.getShort(30), directory = (cfb.getInt(48) + 1) * sector;
        cfb.putInt(directory + 76, 0); // Root child points to the root itself.
        assertFalse(MathTypeRoundTripInspectCli.inspectOle(cyclic).valid());
        byte[] hugeStream = ole(mtef('1'), new byte[]{1});
        ByteBuffer.wrap(hugeStream).order(ByteOrder.LITTLE_ENDIAN).putInt(directory + 124, 1);
        assertFalse(MathTypeRoundTripInspectCli.inspectOle(hugeStream).valid());
    }

    @Test void reportCannotOverwriteInputHardlinkSymlinkOrExistingReport() throws Exception {
        Path a=doc("input",paragraph(run("s","o1","p1","")),mtef('1'),wmf(0,0,0,24,20),"");
        byte[] original=Files.readAllBytes(a); var report=MathTypeRoundTripInspectCli.compare(a,a);
        Path hard=tmp.resolve("hard.json"), sym=tmp.resolve("sym.json");
        Files.createLink(hard,a); Files.createSymbolicLink(sym,a);
        for(Path path:List.of(a,hard,sym)) assertThrows(IOException.class,()->MathTypeRoundTripInspectCli.writeReport(report,path,a));
        Path out=tmp.resolve("report.json"); MathTypeRoundTripInspectCli.writeReport(report,out,a);
        assertThrows(IOException.class,()->MathTypeRoundTripInspectCli.writeReport(report,out,a));
        assertArrayEquals(original,Files.readAllBytes(a));
    }

    static byte[] header(){return new byte[]{5,1,0,7,0,'D','S','M','T','7',0,1};}
    static byte[] mtef(char digit){return concat(header(),new byte[]{1,0,9,101,0,3,2,0,(byte)136,(byte)digit,0,0,0});}
    static byte[] concat(byte[]... arrays){ByteArrayOutputStream o=new ByteArrayOutputStream();for(byte[]a:arrays)o.writeBytes(a);return o.toByteArray();}
    static byte[] ole(byte[]m,byte[]other)throws Exception{
        byte[] nativeData=new byte[28+m.length]; ByteBuffer.wrap(nativeData).order(ByteOrder.LITTLE_ENDIAN).putShort((short)28).putInt(8,m.length);
        System.arraycopy(m,0,nativeData,28,m.length);
        try(POIFSFileSystem fs=new POIFSFileSystem();ByteArrayOutputStream o=new ByteArrayOutputStream()){
            fs.createDocument(new ByteArrayInputStream(nativeData),"Equation Native");
            fs.createDocument(new ByteArrayInputStream(other),"Auxiliary");fs.writeFilesystem(o);return o.toByteArray();
        }
    }
    static byte[] record(int fn,byte[]payload){ByteBuffer b=ByteBuffer.allocate(6+payload.length).order(ByteOrder.LITTLE_ENDIAN);b.putInt(b.capacity()/2).putShort((short)fn).put(payload);return b.array();}
    static byte[] wmf(int tail,int pad,int color,int fontHeight,int dx){
        byte[] font=new byte[50];ByteBuffer.wrap(font).order(ByteOrder.LITTLE_ENDIAN).putShort((short)fontHeight);font[18]='A';font[19]=0;Arrays.fill(font,20,50,(byte)tail);
        ByteBuffer text=ByteBuffer.allocate(12).order(ByteOrder.LITTLE_ENDIAN);text.putShort((short)2).putShort((short)3).putShort((short)1).putShort((short)0).put((byte)'1').put((byte)pad).putShort((short)dx);
        ByteBuffer pen=ByteBuffer.allocate(10).order(ByteOrder.LITTLE_ENDIAN);pen.putShort((short)0).putShort((short)1).putShort((short)0).putInt(color);
        byte[] body=concat(record(0x2fb,font),record(0x2fa,pen.array()),record(0xa32,text.array()),record(0,new byte[0]));
        ByteBuffer h=ByteBuffer.allocate(18).order(ByteOrder.LITTLE_ENDIAN);h.putShort((short)1).putShort((short)9).putShort((short)0x300).putInt((18+body.length)/2).putShort((short)2).putInt(28).putShort((short)0);return concat(h.array(),body);
    }
    static String run(String shape,String ole,String preview,String extra){return "<w:r>"+extra+"<w:object><v:shape id='"+shape+"' style='width:20pt;height:10pt;position:0;opacity:1'><v:imagedata r:id='"+preview+"' croptop='0'/></v:shape><o:OLEObject ProgID='Equation.DSMT4' ShapeID='"+shape+"' r:id='"+ole+"'/></w:object></w:r>";}
    static String paragraph(String runs){return "<w:p><w:r><w:t>Context</w:t></w:r>"+runs+"</w:p>";}
    static byte[] document(String content){return ("<w:document "+NS+"><w:body>"+content+"</w:body></w:document>").getBytes(StandardCharsets.UTF_8);}
    static byte[] rels(String oid,String pid,String obj,String img,boolean external){return ("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='"+oid+"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject' Target='"+(external?obj:"embeddings/"+obj)+"'"+(external?" TargetMode='External'":"")+"/><Relationship Id='"+pid+"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/"+img+"'/></Relationships>").getBytes(StandardCharsets.UTF_8);}
    Path doc(String name,String content,byte[]m,byte[]wmf,String styles)throws Exception{
        Map<String,byte[]> parts=new LinkedHashMap<>();parts.put("word/document.xml",document(content));parts.put("word/_rels/document.xml.rels",rels("o1","p1","object1.bin","image1.wmf",false));parts.put("word/embeddings/object1.bin",ole(m,new byte[]{1,2}));parts.put("word/media/image1.wmf",wmf);parts.put("word/styles.xml",("<w:styles "+NS+">"+styles+"</w:styles>").getBytes(StandardCharsets.UTF_8));return zip(name,parts);
    }
    Path zip(String name,Map<String,byte[]>parts)throws IOException{Path p=tmp.resolve(name+".docx");try(ZipOutputStream z=new ZipOutputStream(Files.newOutputStream(p))){for(var e:parts.entrySet()){z.putNextEntry(new ZipEntry(e.getKey()));z.write(e.getValue());z.closeEntry();}}return p;}
    static Map<String,byte[]> entries(Path p)throws IOException{Map<String,byte[]>m=new LinkedHashMap<>();try(ZipFile z=new ZipFile(p.toFile())){for(ZipEntry e:Collections.list(z.entries()))m.put(e.getName(),z.getInputStream(e).readAllBytes());}return m;}
}
