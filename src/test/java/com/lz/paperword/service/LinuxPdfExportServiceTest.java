package com.lz.paperword.service;

import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.mock.env.MockEnvironment;
import java.nio.file.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicBoolean;
import static org.junit.jupiter.api.Assertions.*;

class LinuxPdfExportServiceTest {
    @TempDir Path temp;
    @BeforeEach void linuxOnly() { Assumptions.assumeTrue(System.getProperty("os.name").toLowerCase().contains("linux")); }
    private MockEnvironment configured(String body) throws Exception {
        Path script=temp.resolve("helper.py"), jar=temp.resolve("uno.jar");
        Files.writeString(script,"import pathlib,sys,time,os\nsource=pathlib.Path(sys.argv[1])\nout=pathlib.Path(sys.argv[2]);out.mkdir()\n"+body);
        Files.writeString(jar,"test fixture");
        Files.writeString(temp.resolve("LibreOfficeBaselineExporter.java"),"// test fixture");
        return new MockEnvironment().withProperty("paperword.pdf.enabled","true")
            .withProperty("paperword.pdf.helper",script.toString()).withProperty("paperword.pdf.uno-jar",jar.toString());
    }
    private static String success() {
        return "(out/'document-native.pdf').write_bytes(b'%PDF-native')\n(out/'document-baseline-compatible.pdf').write_bytes(b'%PDF-compatible')\n";
    }
    @Test void disabledDoesNotGenerateDocx() {
        var called=new AtomicBoolean();var service=new LinuxPdfExportService(new MockEnvironment());
        var e=assertThrows(LinuxPdfExportService.PdfExportException.class,()->service.export(()->{called.set(true);return new byte[1];},LinuxPdfExportService.Mode.compatible));
        assertEquals("PDF_DISABLED",e.getCode());assertFalse(called.get());
    }
    @Test void missingDependencyFailsBeforeGeneration() {
        var service=new LinuxPdfExportService(new MockEnvironment().withProperty("paperword.pdf.enabled","true"));
        assertEquals("PDF_UNAVAILABLE",assertThrows(LinuxPdfExportService.PdfExportException.class,()->service.export(()->new byte[1],LinuxPdfExportService.Mode.compatible)).getCode());
    }
    @Test void returnsRequestedPdfAndCleansSource() throws Exception {
        Path observed=temp.resolve("path.txt");
        var service=new LinuxPdfExportService(configured("pathlib.Path('"+observed+"').write_text(str(source.parent))\n"+success()));
        assertEquals("%PDF-compatible",new String(service.export(()->new byte[1],LinuxPdfExportService.Mode.compatible)));
        assertFalse(Files.exists(Path.of(Files.readString(observed))));
        assertEquals("%PDF-native",new String(service.export(()->new byte[1],LinuxPdfExportService.Mode.native_layout)));
        assertEquals(2L,service.status().get("completed"));
    }
    @Test void failedPartialOutputNeverReturnsSuccessAndCleans() throws Exception {
        Path observed=temp.resolve("path.txt");
        var service=new LinuxPdfExportService(configured("pathlib.Path('"+observed+"').write_text(str(source.parent))\n"+success()+"sys.exit(1)\n"));
        assertEquals("PDF_CONVERSION_FAILED",assertThrows(LinuxPdfExportService.PdfExportException.class,()->service.export(()->new byte[1],LinuxPdfExportService.Mode.compatible)).getCode());
        assertFalse(Files.exists(Path.of(Files.readString(observed))));
        assertEquals(0,service.status().get("active"));
    }
    @Test void busySlotRejectedBeforeDocxGenerationAndReleased() throws Exception {
        var service=new LinuxPdfExportService(configured(success()));
        var entered=new CountDownLatch(1);var release=new CountDownLatch(1);
        try(var executor=Executors.newSingleThreadExecutor()) {
            var first=executor.submit(()->service.export(()->{entered.countDown();try{release.await();}catch(InterruptedException e){Thread.currentThread().interrupt();}return new byte[1];},LinuxPdfExportService.Mode.compatible));
            assertTrue(entered.await(5,TimeUnit.SECONDS));
            try {
                assertEquals("PDF_BUSY",assertThrows(LinuxPdfExportService.PdfExportException.class,()->service.export(()->{fail("must not generate while busy");return null;},LinuxPdfExportService.Mode.compatible)).getCode());
            } finally { release.countDown(); }
            assertTrue(first.get(10,TimeUnit.SECONDS).length>5);
            assertEquals(0,service.status().get("active"));
        }
    }
    @Test void dependencyExitAndDeadlineHaveStableCodes() throws Exception {
        var missing=new LinuxPdfExportService(configured("sys.exit(69)\n"));
        assertEquals("PDF_UNAVAILABLE",assertThrows(LinuxPdfExportService.PdfExportException.class,()->missing.export(()->new byte[1],LinuxPdfExportService.Mode.compatible)).getCode());
        var deadline=new LinuxPdfExportService(configured("sys.exit(124)\n"));
        assertEquals("PDF_TIMEOUT",assertThrows(LinuxPdfExportService.PdfExportException.class,()->deadline.export(()->new byte[1],LinuxPdfExportService.Mode.compatible)).getCode());
    }
    @Test void oversizedDocxAndInvalidSignatureAreRejected() throws Exception {
        var env=configured(success()).withProperty("paperword.pdf.max-docx-mib","1");
        var service=new LinuxPdfExportService(env);
        assertEquals("PDF_INPUT_TOO_LARGE",assertThrows(LinuxPdfExportService.PdfExportException.class,()->service.export(()->new byte[1048577],LinuxPdfExportService.Mode.compatible)).getCode());
        var invalid=new LinuxPdfExportService(configured("(out/'document-baseline-compatible.pdf').write_bytes(b'not a PDF')\n"));
        assertEquals("PDF_CONVERSION_FAILED",assertThrows(LinuxPdfExportService.PdfExportException.class,()->invalid.export(()->new byte[1],LinuxPdfExportService.Mode.compatible)).getCode());
    }
    @Test void interruptKillsHelperAndReleasesSlot() throws Exception {
        Path pid=temp.resolve("pid.txt");
        var service=new LinuxPdfExportService(configured("pathlib.Path('"+pid+"').write_text(str(os.getpid()))\ntime.sleep(60)\n"));
        var code=new java.util.concurrent.atomic.AtomicReference<String>();var interrupted=new AtomicBoolean();
        Thread caller=new Thread(()->{try{service.export(()->new byte[1],LinuxPdfExportService.Mode.compatible);}catch(Exception e){code.set(((LinuxPdfExportService.PdfExportException)e).getCode());interrupted.set(Thread.currentThread().isInterrupted());}});
        caller.start();long deadline=System.nanoTime()+TimeUnit.SECONDS.toNanos(5);
        while(!Files.exists(pid)&&System.nanoTime()<deadline)Thread.sleep(10);
        assertTrue(Files.exists(pid));long process=Long.parseLong(Files.readString(pid));
        caller.interrupt();caller.join(5000);
        assertFalse(caller.isAlive());assertEquals("PDF_INTERRUPTED",code.get());assertTrue(interrupted.get());
        for(int i=0;i<50&&ProcessHandle.of(process).map(ProcessHandle::isAlive).orElse(false);i++)Thread.sleep(10);
        assertFalse(ProcessHandle.of(process).map(ProcessHandle::isAlive).orElse(false));
        assertEquals(0,service.status().get("active"));
    }
    @Test void stoppedHelperCannotLeakReparentedGroupChildOnInterrupt() throws Exception {
        Path orphan=temp.resolve("orphan.txt");
        String child="import subprocess,sys,pathlib; p=subprocess.Popen([sys.executable,'-c','import time;time.sleep(60)']); pathlib.Path('"+orphan+"').write_text(str(p.pid))";
        String body="import subprocess,signal\nsubprocess.run([sys.executable,'-c',"+quote(child)+"],check=True)\nos.kill(os.getpid(),signal.SIGSTOP)\ntime.sleep(60)\n";
        var service=new LinuxPdfExportService(configured(body));
        Thread caller=new Thread(()->{try{service.export(()->new byte[1],LinuxPdfExportService.Mode.compatible);}catch(Exception expected){}});
        caller.start();long deadline=System.nanoTime()+TimeUnit.SECONDS.toNanos(5);
        while(!Files.exists(orphan)&&System.nanoTime()<deadline)Thread.sleep(10);
        assertTrue(Files.exists(orphan));long pid=Long.parseLong(Files.readString(orphan));
        caller.interrupt();caller.join(8000);assertFalse(caller.isAlive());
        for(int i=0;i<100&&running(pid);i++)Thread.sleep(10);
        assertFalse(running(pid),"An orphaned LO-like process survived cancellation");
        assertEquals(0,service.status().get("active"));
    }
    private static String quote(String value) { return "'"+value.replace("\\","\\\\").replace("'","\\'")+"'"; }
    private static boolean running(long pid) throws Exception {
        Path stat=Path.of("/proc",Long.toString(pid),"stat");
        if(!Files.exists(stat))return false;
        String text=Files.readString(stat);return !text.substring(text.lastIndexOf(')')+2).startsWith("Z ");
    }
}
