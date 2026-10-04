package com.lz.paperword.service;

import com.lz.paperword.model.layout.LayoutDocumentRequest;
import org.junit.jupiter.api.Test;
import java.io.*;
import java.util.*;
import java.util.regex.*;
import java.util.zip.*;
import static org.junit.jupiter.api.Assertions.*;

class LayoutFormulaSizeRegressionTest {
    @Test void explicitBodySizeScalesFormulaPreviewAndDescentTogether() throws Exception {
        var request=new LayoutDocumentRequest();var page=new LayoutDocumentRequest.Page();
        for(int size:List.of(11,16)){
            var block=new LayoutDocumentRequest.Block();block.setType(LayoutDocumentRequest.BlockType.PARAGRAPH);
            block.setText("Ag $x_i^2+1$ 后文");block.getStyle().setFontSizePt(size);page.getBlocks().add(block);
        }
        request.setPages(List.of(page));byte[] docx=new LayoutExportService().exportLayoutDocument(request);
        String xml=null;try(var zip=new ZipInputStream(new ByteArrayInputStream(docx))){ZipEntry e;while((e=zip.getNextEntry())!=null)if(e.getName().equals("word/document.xml"))xml=new String(zip.readAllBytes(),java.nio.charset.StandardCharsets.UTF_8);}
        assertNotNull(xml);var boxes=Pattern.compile("style=\"width:([0-9.]+)pt;height:([0-9.]+)pt\"").matcher(xml);var widths=new ArrayList<Double>();var heights=new ArrayList<Double>();
        while(boxes.find()){widths.add(Double.parseDouble(boxes.group(1)));heights.add(Double.parseDouble(boxes.group(2)));}
        assertEquals(2,widths.size());assertEquals(16d/11,widths.get(1)/widths.get(0),0.08);assertEquals(16d/11,heights.get(1)/heights.get(0),0.08);
        var positions=Pattern.compile("<w:position w:val=\"(-?[0-9]+)\"").matcher(xml);var depths=new ArrayList<Integer>();while(positions.find())depths.add(Integer.parseInt(positions.group(1)));
        assertEquals(2,depths.size());assertEquals(depths.get(0)*16d/11,depths.get(1),1.5);
    }
}
