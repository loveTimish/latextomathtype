package com.lz.paperword.service;

import com.lz.paperword.model.*;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import com.lz.paperword.core.mtef.MathTypeDocxComparator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;
import java.util.zip.*;
import java.io.*;
import static org.junit.jupiter.api.Assertions.*;

class ExportConcurrencyRegressionTest {
    private PaperExportRequest paper(int kind) {
        var request = new PaperExportRequest();
        var info = new PaperExportRequest.PaperInfo();
        info.setName("Concurrent paper " + kind); info.setCompactLayout(kind % 2 == 0);
        request.setPaper(info); var section = new SectionDTO(); section.setHeadline("Test");
        var questions = new ArrayList<QuestionDTO>();
        for (int i=0; i<8; i++) { var q = new QuestionDTO();q.setSerialNumber(i+1);q.setQuestionType(6);
            q.setContent("$\\frac{x^{"+kind+"}+1}{x+"+kind+"}$"); questions.add(q); }
        section.setQuestions(questions);request.setSections(List.of(section));return request;
    }
    private LayoutDocumentRequest layout(int kind) {
        var request = new LayoutDocumentRequest(); var page = new LayoutDocumentRequest.Page();
        for(int i=0;i<8;i++){var block = new LayoutDocumentRequest.Block();block.setType(LayoutDocumentRequest.BlockType.FORMULA);block.setLatex("\\frac{x^{"+kind+"}+1}{x+"+kind+"}");page.getBlocks().add(block);}
        request.setPages(List.of(page));return request;
    }
    private List<String> nativeHashes(Path path, byte[] bytes) throws Exception {
        Files.write(path, bytes);var report = new MathTypeDocxComparator().inspect(path);
        assertEquals(8, report.formulas().size());assertTrue(report.formulas().stream().allMatch(f->f.valid()));
        return report.formulas().stream().map(f->f.mtefSha256()).toList();
    }
    @Test void sharedServicesDoNotMixDocumentState(@TempDir Path dir) throws Exception {
        var paperService = new PaperExportService(); var layoutService = new LayoutExportService();
        for(boolean layoutMode:List.of(false,true)){
            var expected = new HashMap<Integer,List<String>>();
            for(int kind=2;kind<6;kind++) expected.put(kind,nativeHashes(dir.resolve("expected-"+layoutMode+"-"+kind+".docx"),layoutMode?layoutService.exportLayoutDocument(layout(kind)):paperService.export(paper(kind))));
            try(var pool=Executors.newFixedThreadPool(4)){
                for(int round=0;round<3;round++){
                    var gate=new CountDownLatch(1);var futures=new ArrayList<Future<List<String>>>();
                    for(int kind=2;kind<6;kind++){int k=kind;int r=round;futures.add(pool.submit(()->{gate.await();return nativeHashes(dir.resolve("actual-"+layoutMode+"-"+r+"-"+k+".docx"),layoutMode?layoutService.exportLayoutDocument(layout(k)):paperService.export(paper(k)));}));}
                    gate.countDown();for(int i=0;i<4;i++)assertEquals(expected.get(i+2),futures.get(i).get(30,TimeUnit.SECONDS));
                }
            }
        }
    }
    @Test void vmlUsesLocaleIndependentDecimalPoints() throws Exception {
        Locale previous=Locale.getDefault();
        try { Locale.setDefault(Locale.GERMANY);byte[] bytes=new PaperExportService().export(paper(2));
            try(var zip=new ZipInputStream(new ByteArrayInputStream(bytes))){ZipEntry entry;boolean found=false;
                while((entry=zip.getNextEntry())!=null)if(entry.getName().equals("word/document.xml")){
                    String xml=new String(zip.readAllBytes(),java.nio.charset.StandardCharsets.UTF_8);
                    assertFalse(java.util.regex.Pattern.compile("(?:width|height):[0-9]+,[0-9]+pt").matcher(xml).find());
                    assertTrue(java.util.regex.Pattern.compile("width:[0-9]+\\.[0-9]+pt").matcher(xml).find());found=true;
                }assertTrue(found);
            }
        } finally {Locale.setDefault(previous);}
    }
}
