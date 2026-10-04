package com.lz.paperword.core.render;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.lz.paperword.model.*;
import com.lz.paperword.service.PaperExportService;
import com.lz.paperword.core.mtef.MathTypeDocxComparator;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;
import java.util.zip.*;
import java.io.*;
import java.lang.reflect.*;

/** Reproducible synthetic benchmark. No user data or external images. */
public final class LinuxBenchmark {
 static final ObjectMapper JSON = new ObjectMapper();
 static final String COMPLEX = "\\frac{-b+\\sqrt{b^2-4ac}}{2a}+\\sum_{k=1}^{n}\\frac{k^2}{k+1}";
 static Path out;
 static long requests() throws Exception { var f=LaTeXImageRenderer.class.getDeclaredField("mathJaxRequestId"); f.setAccessible(true); return f.getLong(null); }
 static void result(String name,long start,int count,int errors,Object detail)throws Exception {
   var m=new LinkedHashMap<String,Object>();m.put("case",name);m.put("ms",(System.nanoTime()-start)/1e6);m.put("count",count);m.put("errors",errors);m.put("heap_used_bytes",Runtime.getRuntime().totalMemory()-Runtime.getRuntime().freeMemory());m.put("detail",detail);System.out.println("BENCH "+JSON.writeValueAsString(m));
 }
 static PaperExportRequest paper(int n,int kind) {
  var p=new PaperExportRequest();var info=new PaperExportRequest.PaperInfo();info.setName("Synthetic benchmark "+kind);info.setCompactLayout(kind%2==0);p.setPaper(info);
  var s=new SectionDTO();s.setHeadline("Synthetic section");var qs=new ArrayList<QuestionDTO>();
  for(int i=0;i<n;i++){var q=new QuestionDTO();q.setSerialNumber(i+1);q.setQuestionType(6);q.setContent("$"+(kind==0?"x^2+1":COMPLEX+"+"+((kind==1)?i:kind))+"$");qs.add(q);}s.setQuestions(qs);p.setSections(List.of(s));return p;
 }
 static Map<String,byte[]> unzip(byte[] b)throws Exception { var out=new HashMap<String,byte[]>();try(var z=new ZipInputStream(new ByteArrayInputStream(b))){ZipEntry e;while((e=z.getNextEntry())!=null)out.put(e.getName(),z.readAllBytes());}return out; }
 static List<String> check(byte[] b,int count,String name)throws Exception {
  Path p=out.resolve(name+".docx");Files.write(p,b);var report=new MathTypeDocxComparator().inspect(p);
  if(report.formulas().size()!=count||report.formulas().stream().anyMatch(x->!x.valid()))throw new IllegalStateException("Invalid OLE count or MTEF");
  int wmf=0;for(var e:unzip(b).entrySet())if(e.getKey().endsWith(".wmf")){if(!WmfPreviewInspector.inspect(e.getValue()).pureVector())throw new IllegalStateException("not vector");wmf++;}
  if(wmf!=count)throw new IllegalStateException("Preview count "+wmf);
  return report.formulas().stream().map(x->x.mtefSha256()).toList();
 }
 static void performance()throws Exception {
  var r=new LaTeXImageRenderer();long t=System.nanoTime();r.renderForOlePreview("x^2+1");result("cold_simple",t,1,0,requests());
  t=System.nanoTime();for(int i=0;i<1000;i++)r.renderForOlePreview("x^2+1");result("warm_same_1000",t,1000,0,requests());
  t=System.nanoTime();for(int i=0;i<40;i++)r.renderForOlePreview(COMPLEX+"+"+i);result("complex_unique_40",t,40,0,requests());
  try(var pool=Executors.newFixedThreadPool(8)){
   for(int batch=0;batch<5;batch++){String tex=COMPLEX+"+"+(10000+batch);var gate=new CountDownLatch(1);var fs=new ArrayList<Future<LaTeXImageRenderer.PreviewImage>>();long before=requests();
    for(int i=0;i<8;i++)fs.add(pool.submit(()->{gate.await();return r.renderForOlePreview(tex);}));t=System.nanoTime();gate.countDown();byte[] first=null;int errors=0;
    for(var f:fs)try{byte[] bytes=f.get(30,TimeUnit.SECONDS).data();if(first==null)first=bytes;else if(!Arrays.equals(first,bytes))errors++;}catch(Exception e){errors++;}
    result("concurrent_same_cold_8_"+batch,t,8,errors,Map.of("worker_requests",requests()-before));
   }
  }
  var svc=new PaperExportService();
  for(int kind=0;kind<=1;kind++){var p=paper(40,kind);for(int round=0;round<3;round++){t=System.nanoTime();byte[] b=svc.export(p);long elapsed=System.nanoTime()-t;check(b,40,"paper_"+kind+"_"+round);result("paper40_"+(kind==0?"repeat":"unique")+"_"+round,System.nanoTime()-elapsed,1,0,Map.of("bytes",b.length));}}
 }
 static void distinct()throws Exception {
  var renderer=new LaTeXImageRenderer();
  for(int i=0;i<20;i++)renderer.renderForOlePreview(COMPLEX+"+"+(20000+i));
  long start=System.nanoTime();long before=requests();
  for(int i=0;i<60;i++)renderer.renderForOlePreview(COMPLEX+"+"+(30000+i));
  result("warm_jvm_distinct_60",start,60,0,Map.of("worker_requests",requests()-before));
 }
 static void coldPaper(int n)throws Exception {
  var svc=new PaperExportService();var request=paper(n,1);
  for(int round=0;round<3;round++){long t=System.nanoTime();byte[] b=svc.export(request);long elapsed=System.nanoTime()-t;check(b,n,"coldpaper_"+n+"_"+round);result("coldpaper"+n+"_"+round,System.nanoTime()-elapsed,1,0,Map.of("bytes",b.length,"worker_requests",requests()));}
 }
 static void compareOutputs(Path baseline)throws Exception {
  var comparator=new MathTypeDocxComparator();int docs=0;int formulas=0;
  try(var paths=Files.list(out)){for(Path optimized:paths.filter(x->x.toString().endsWith(".docx")).toList()){
   Path original=baseline.resolve(optimized.getFileName());var comparison=comparator.compare(original,optimized);
   if(!comparison.passesExactStructureGate()||comparison.rawMtefEqualCount()!=comparison.generatedObjectCount())throw new IllegalStateException("MTEF difference: "+optimized);
   var before=unzip(Files.readAllBytes(original));var after=unzip(Files.readAllBytes(optimized));
   for(var entry:before.entrySet())if(entry.getKey().startsWith("word/media/"))if(!Arrays.equals(entry.getValue(),after.get(entry.getKey())))throw new IllegalStateException("Preview changed "+entry.getKey());
   docs++;formulas+=comparison.generatedObjectCount();
  }}System.out.println("VALIDATED "+docs+" documents, "+formulas+" exact MTEF and preview pairs");
 }
 static void concurrency()throws Exception {
  var svc=new PaperExportService();var expected=new HashMap<Integer,List<String>>();for(int k=2;k<10;k++)expected.put(k,check(svc.export(paper(12,k)),12,"expected_"+k));
  long t=System.nanoTime();int errors=0,total=0;var messages=new ArrayList<String>();
  try(var pool=Executors.newFixedThreadPool(8)){for(int batch=0;batch<8;batch++){
   var gate=new CountDownLatch(1);var fs=new ArrayList<Future<String>>();int b=batch;
   for(int k=2;k<10;k++){int kind=k;fs.add(pool.submit(()->{gate.await();try{var got=check(svc.export(paper(12,kind)),12,"concurrent_"+b+"_"+kind);return got.equals(expected.get(kind))?"ok":"MTEF differs from serial reference";}catch(Exception e){return e.toString();}}));}
   gate.countDown();for(var f:fs){String s=f.get(45,TimeUnit.SECONDS);total++;if(!s.equals("ok")){errors++;if(messages.size()<8)messages.add(s);}}
  }}result("concurrent_service_64",t,total,errors,messages);
 }
 static void bugs()throws Exception {
  var svc=new PaperExportService();Locale.setDefault(Locale.GERMANY);var b=svc.export(paper(1,0));String xml=new String(unzip(b).get("word/document.xml"),java.nio.charset.StandardCharsets.UTF_8);System.out.println("BUG locale_comma="+java.util.regex.Pattern.compile("width:[0-9]+,[0-9]+pt").matcher(xml).find());Locale.setDefault(Locale.ROOT);
  Path canary=out.resolve("synthetic-canary.txt");Files.writeString(canary,"SYNTHETIC_TEST_CANARY_NO_PRIVATE_DATA");var p=paper(0,0);p.getSections().get(0).setImages(List.of(canary.toAbsolutePath().toString()));b=svc.export(p);boolean leaked=unzip(b).entrySet().stream().anyMatch(e->e.getKey().startsWith("word/media/")&&new String(e.getValue()).contains("SYNTHETIC_TEST_CANARY"));System.out.println("BUG text_canary_embedded="+leaked);
  String ld="\\begin{longdivision}{rrrr}{6}{570}{3420}&30\\\\\\cline{1-2}&&42\\\\&&42\\\\\\cline{2-3}&&&0\\end{longdivision}";var r=new LaTeXImageRenderer();var plain=r.renderMathJaxSvgForAcceptance(ld);var mixed=r.renderMathJaxSvgForAcceptance("x+"+ld+"+1");System.out.println("BUG mixed_longdivision_svg_identical="+Arrays.equals(plain.svgBytes(),mixed.svgBytes()));
 }
 static void timeout()throws Exception {
  Path script=out.resolve("hanging-worker.cjs");Files.writeString(script,"process.stdin.resume(); setInterval(()=>{},1000);\n");System.setProperty("paperword.mathjax.script",script.toAbsolutePath().toString());System.setProperty("paperword.latex.timeout.seconds","1");long t=System.nanoTime();try{new LaTeXImageRenderer().renderForOlePreview("x+987654");}catch(Exception e){result("hanging_worker_timeout",t,1,1,e.toString());}
 }
 public static void main(String[] args)throws Exception {
  ((ch.qos.logback.classic.Logger)org.slf4j.LoggerFactory.getLogger(org.slf4j.Logger.ROOT_LOGGER_NAME)).setLevel(ch.qos.logback.classic.Level.WARN);
  out=Path.of(args[1]);Files.createDirectories(out);if(System.getProperty("paperword.render.cache.enabled")==null)System.setProperty("paperword.render.cache.enabled","false");
  switch(args[0]){case "performance"->performance();case "distinct"->distinct();case "paper40"->coldPaper(40);case "paper100"->coldPaper(100);case "compare"->compareOutputs(Path.of(args[2]));case "concurrency"->concurrency();case "bugs"->bugs();case "timeout"->timeout();default->throw new IllegalArgumentException();}
 }
}
