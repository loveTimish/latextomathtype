package com.lz.paperword.core.render;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.util.*;
import java.nio.file.*;
public class CacheSoak {
 public static void main(String[] args)throws Exception {
  ((ch.qos.logback.classic.Logger)org.slf4j.LoggerFactory.getLogger(org.slf4j.Logger.ROOT_LOGGER_NAME)).setLevel(ch.qos.logback.classic.Level.WARN);
  var renderer=new LaTeXImageRenderer();var json=new ObjectMapper();byte[] first=null;long t=System.nanoTime();
  for(int i=0;i<100;i++) {
   var preview=renderer.renderForOlePreview("\\frac{x^{2}+"+i+"}{\\sqrt{y+1}}+\\sum_{k=1}^{n}k^2");
   if(i==0)first=preview.data().clone();
   if(!WmfPreviewInspector.inspect(preview.data()).pureVector())throw new AssertionError("not vector");
   if(i%10==9)System.out.println("SOAK "+json.writeValueAsString(Map.of("renders",i+1,"elapsed_ms",(System.nanoTime()-t)/1e6,"stats",renderer.cacheStatistics())));
  }
  byte[] repeated=renderer.renderForOlePreview("\\frac{x^{2}+0}{\\sqrt{y+1}}+\\sum_{k=1}^{n}k^2").data();
  if(!Arrays.equals(first,repeated))throw new AssertionError("eviction changed WMF bytes");
  var stats=renderer.cacheStatistics();
  if(stats.get("memory.bytes")>262144||stats.get("memory.entries")>16||stats.get("disk.bytes")>262144||stats.get("disk.entries")>16)throw new AssertionError("quota exceeded");
  System.out.println("SOAK_FINAL "+json.writeValueAsString(Map.of("sameAfterEviction",true,"stats",stats)));
 }
}
