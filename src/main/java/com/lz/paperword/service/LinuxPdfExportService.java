package com.lz.paperword.service;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.Semaphore;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicLong;

/** Opt-in, bounded Linux converter. Only converts DOCX created by our own exporters. */
@Service
public final class LinuxPdfExportService {
    private static final Logger log = LoggerFactory.getLogger(LinuxPdfExportService.class);
    public enum Mode { compatible, native_layout }
    @FunctionalInterface public interface DocxProducer { byte[] create() throws IOException; }
    public static final class PdfExportException extends RuntimeException {
        private final String code;
        public PdfExportException(String code, String message) { super(message); this.code=code; }
        public String getCode() { return code; }
    }
    private final Environment environment;
    private final Semaphore slots;
    private final AtomicLong completed = new AtomicLong(), failed = new AtomicLong();
    private final int concurrency;

    public LinuxPdfExportService(Environment environment) {
        this.environment=environment;
        this.concurrency=integer("max-concurrent", 1, 1, 8);
        this.slots=new Semaphore(concurrency, true);
    }
    private String setting(String name, String fallback) { return environment.getProperty("paperword.pdf."+name, fallback); }
    private int integer(String name, int fallback, int min, int max) {
        try { int value=Integer.parseInt(setting(name, Integer.toString(fallback)));
            if(value<min || value>max) throw new NumberFormatException(); return value;
        } catch (NumberFormatException e) { throw error("PDF_CONFIGURATION", "Invalid paperword.pdf."+name); }
    }
    private static PdfExportException error(String code, String text) { return new PdfExportException(code, text); }
    private Path requiredFile(String key) {
        String value=setting(key, "");
        if(value.isBlank()) throw error("PDF_UNAVAILABLE", "Configure paperword.pdf."+key+" before enabling PDF export");
        Path path=Path.of(value).toAbsolutePath().normalize();
        if(!Files.isRegularFile(path)) throw error("PDF_UNAVAILABLE", "Configured PDF dependency is missing: "+key);
        return path;
    }
    public Map<String,Object> status() {
        return Map.of("enabled", Boolean.parseBoolean(setting("enabled","false")), "active", concurrency-slots.availablePermits(),
            "maxConcurrent", concurrency, "completed", completed.get(), "failed", failed.get());
    }
    public byte[] export(DocxProducer producer, Mode mode) throws IOException {
        if(!Boolean.parseBoolean(setting("enabled", "false"))) throw error("PDF_DISABLED", "Linux PDF export is disabled; configure and enable paperword.pdf.enabled");
        if(!System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("linux")) throw error("PDF_UNAVAILABLE", "This PDF integration requires Linux");
        Path helper=requiredFile("helper"), jar=requiredFile("uno-jar");
        if(!Files.isExecutable(Path.of("/usr/bin/setsid")) || !Files.isExecutable(Path.of("/bin/kill")))
            throw error("PDF_UNAVAILABLE", "Linux PDF isolation requires /usr/bin/setsid and /bin/kill");
        if(!Files.isRegularFile(helper.resolveSibling("LibreOfficeBaselineExporter.java"))) throw error("PDF_UNAVAILABLE", "The PDF helper's Java source is missing");
        int timeout=integer("timeout-seconds", 120, 1, 600);
        long maxDocx=(long)integer("max-docx-mib", 32, 1, 256)*1024*1024;
        long maxPdf=(long)integer("max-pdf-mib", 64, 1, 256)*1024*1024;
        if(!slots.tryAcquire()) throw error("PDF_BUSY", "All PDF conversion slots are busy; retry later");
        Path work=null; Process process=null; Thread drain=null;
        BoundedLog output=new BoundedLog(); long start=System.nanoTime();
        try {
            byte[] docx=producer.create();
            if(docx.length>maxDocx) throw error("PDF_INPUT_TOO_LARGE", "Generated DOCX exceeds the PDF input byte limit");
            work=Files.createTempDirectory("paperword-pdf-");
            Path source=work.resolve("document.docx"), out=work.resolve("output");
            Files.write(source,docx,StandardOpenOption.CREATE_NEW);
            var command=List.of("/usr/bin/setsid",setting("python-command","python3"), helper.toString(), source.toString(), out.toString(),
                "--uno-jar",jar.toString(),"--soffice",setting("soffice-command","soffice"),"--java",setting("java-command","java"),"--timeout",Integer.toString(timeout),
                "--cjk-font",setting("cjk-font","Noto Serif CJK SC"),"--managed-process-group");
            try { process=new ProcessBuilder(command).redirectErrorStream(true).start(); }
            catch(IOException e) { throw error("PDF_UNAVAILABLE", "Cannot start the configured PDF helper; check Python and executable paths"); }
            Process running=process;
            drain=Thread.ofPlatform().daemon().name("paperword-pdf-output").start(()->output.drain(running.getInputStream()));
            if(!process.waitFor(timeout+5L,TimeUnit.SECONDS)) throw error("PDF_TIMEOUT", "PDF conversion exceeded its time limit");
            drain.join(1000);
            if(process.exitValue()!=0) {
                log.warn("PDF helper failed (exit {}): {}",process.exitValue(),output.text());
                // Logs retain details, while the HTTP error never exposes local source paths.
                String code=process.exitValue()==124 ? "PDF_TIMEOUT" : Set.of(69,126,127).contains(process.exitValue()) ? "PDF_UNAVAILABLE" : "PDF_CONVERSION_FAILED";
                throw error(code,"PDF conversion failed; inspect the server conversion log (exit "+process.exitValue()+")");
            }
            output.text().lines().filter(line->line.startsWith("CJK_FONT_FALLBACK ")).forEach(line->log.info("{}",line));
            Path pdf=out.resolve("document-"+(mode==Mode.native_layout?"native":"baseline-compatible")+".pdf");
            if(!Files.isRegularFile(pdf,LinkOption.NOFOLLOW_LINKS)) throw error("PDF_CONVERSION_FAILED", "PDF helper did not produce the requested output");
            long size=Files.size(pdf);
            if(size<5) throw error("PDF_CONVERSION_FAILED", "Converter returned an empty PDF output");
            if(size>maxPdf) throw error("PDF_OUTPUT_TOO_LARGE", "PDF output exceeds its byte limit");
            byte[] bytes=Files.readAllBytes(pdf);
            if(!new String(bytes,0,5,StandardCharsets.US_ASCII).equals("%PDF-")) throw error("PDF_CONVERSION_FAILED", "Converter returned an invalid PDF signature");
            completed.incrementAndGet();
            log.info("PDF conversion completed mode={} size={} elapsedMs={}",mode,size,Duration.ofNanos(System.nanoTime()-start).toMillis());
            return bytes;
        } catch(InterruptedException e) {
            Thread.currentThread().interrupt(); failed.incrementAndGet();
            throw error("PDF_INTERRUPTED", "PDF conversion was interrupted");
        } catch(IOException | RuntimeException e) { failed.incrementAndGet(); throw e; }
        finally {
            if(process!=null) stopTree(process);
            if(drain!=null && drain.isAlive()) drain.interrupt();
            if(work!=null) removeOwnedWork(work);
            slots.release();
        }
    }
    private static void stopTree(Process process) {
        // Snapshot descendants BEFORE terminating the parent, to avoid orphaning its children.
        boolean interrupted=Thread.interrupted();
        Set<ProcessHandle> children=new HashSet<>(process.descendants().toList());
        try {
            process.destroy(); // Allow our helper's signal handler to reap its process groups.
            try { process.waitFor(3,TimeUnit.SECONDS); } catch(InterruptedException ignored) { interrupted=true; }
            children.addAll(process.descendants().toList());
            for(ProcessHandle child:children) if(child.isAlive()) child.destroyForcibly();
            if(process.isAlive()) process.destroyForcibly();
            try { process.waitFor(2,TimeUnit.SECONDS); } catch(InterruptedException ignored) { interrupted=true; }
            // The helper and ALL compiler/LO/UNO descendants share this private
            // setsid group. Reap even orphaned soffice children if the helper hung.
            try {
                Process kill=new ProcessBuilder("/bin/kill","-KILL","--","-"+process.pid())
                    .redirectOutput(ProcessBuilder.Redirect.DISCARD).redirectError(ProcessBuilder.Redirect.DISCARD).start();
                if(!kill.waitFor(2,TimeUnit.SECONDS)) kill.destroyForcibly();
            } catch(IOException failure) { log.error("Could not reap private PDF process group",failure); }
              catch(InterruptedException ignored) { interrupted=true; }
        } finally { if(interrupted) Thread.currentThread().interrupt(); }
    }
    private static void removeOwnedWork(Path root) {
        // Files.walk does not follow symbolic links. Only this job's freshly created tree is removed.
        try(var paths=Files.walk(root)) {
            for(Path p:paths.sorted(Comparator.reverseOrder()).toList()) Files.deleteIfExists(p);
        } catch(IOException failure) { log.warn("Could not completely clean a private PDF job directory",failure); }
    }
    private static final class BoundedLog {
        private final StringBuilder text=new StringBuilder();
        void drain(InputStream input) { try(input) { byte[] b=new byte[4096]; int n;
            while((n=input.read(b))>=0) synchronized(this) { if(text.length()<65536) text.append(new String(b,0,Math.min(n,65536-text.length()),StandardCharsets.UTF_8)); }
        } catch(IOException ignored) {} }
        synchronized String text() { return text.toString(); }
    }
}
