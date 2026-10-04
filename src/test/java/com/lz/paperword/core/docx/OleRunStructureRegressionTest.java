package com.lz.paperword.core.docx;

import com.lz.paperword.core.render.LaTeXImageRenderer;
import com.lz.paperword.core.latex.LaTeXParser;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPathFactory;
import java.io.*;
import java.util.List;
import java.util.zip.*;
import static org.junit.jupiter.api.Assertions.*;

class OleRunStructureRegressionTest {
    @Test
    void equationsAreDirectRunChildrenWithGeometricBaselineAndAttachedPoiRun() throws Exception {
        try(var doc = new XWPFDocument()) {
            var embedder=new MathTypeEmbedder();var paragraph=doc.createParagraph();
            paragraph.createRun().setText("中文 Ag ");
            var renderer=new LaTeXImageRenderer();
            for(String latex:List.of("x_i^2+1","\\frac{1}{2}","\\sqrt{x+1}")){
                var run=paragraph.createRun();
                embedder.embedEquation(paragraph,run,new LaTeXParser().parseDetailed(latex).ast(),latex);
                assertNotNull(run.getCTR().getRPr(),"POI's run must remain attached after XML replacement");
                int expected=-(int)Math.round(renderer.renderForOlePreview(latex).depthPt()*2);
                assertEquals(expected,Integer.parseInt(String.valueOf(run.getCTR().getRPr().getPositionArray(0).getVal())));
                paragraph.createRun().setText(" Ag 后文 ");
            }
            var out=new ByteArrayOutputStream();doc.write(out);
            try(var zip=new ZipInputStream(new ByteArrayInputStream(out.toByteArray()))){ZipEntry entry;boolean found=false;
                while((entry=zip.getNextEntry())!=null)if(entry.getName().equals("word/document.xml")){
                    var factory=DocumentBuilderFactory.newInstance();factory.setNamespaceAware(true);
                    var xml=factory.newDocumentBuilder().parse(new ByteArrayInputStream(zip.readAllBytes()));
                    var xp=XPathFactory.newInstance().newXPath();
                    assertEquals("0",xp.evaluate("count(//*[local-name()='r']/*[local-name()='r'])",xml));
                    assertEquals("3",xp.evaluate("count(//*[local-name()='r']/*[local-name()='object'])",xml));
                    assertEquals("3",xp.evaluate("count(//*[local-name()='r'][*[local-name()='object']]/*[local-name()='rPr']/*[local-name()='position'])",xml));found=true;
                }assertTrue(found);
            }
        }
    }
}
