package com.lz.paperword.core.docx;

import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import com.lz.paperword.model.layout.LayoutDocumentRequest;
import com.lz.paperword.service.LayoutExportService;
import com.lz.paperword.service.PaperExportService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class ImageAssetExportTest {
    @TempDir Path root;

    @Test
    void bothBuildersEmbedVerifiedBytesUsingActualFormat() throws Exception {
        Path image = ImageAssetLoaderTest.image(root.resolve("mislabeled.jpeg"), "png", 12, 8);
        ImageAssetLoader assets = new ImageAssetLoader(new ImageAssetConfig(root));
        PaperExportRequest paper = paper("mislabeled.jpeg");
        LayoutDocumentRequest layout = layout("mislabeled.jpeg");
        layout.getPages().getFirst().getBlocks().getFirst().setImageContentType("image/jpeg");
        layout.getPages().getFirst().getBlocks().getFirst().setImageWidthPx(999_999);
        layout.getPages().getFirst().getBlocks().getFirst().setImageHeightPx(999_999);
        for (byte[] docx : List.of(new DocxBuilder(false, assets).build(paper), new LayoutDocxBuilder(false, assets).build(layout))) {
            try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(docx))) {
                assertEquals(1, document.getAllPictures().size());
                assertEquals(XWPFDocument.PICTURE_TYPE_PNG, document.getAllPictures().getFirst().getPictureType());
                assertArrayEquals(Files.readAllBytes(image), document.getAllPictures().getFirst().getData());
            }
        }
    }

    @Test
    void extremeAspectRatiosAndRequestedDisplayWidthsStayWithinPageBounds() throws Exception {
        ImageAssetLoaderTest.image(root.resolve("tall.png"), "png", 1, 20_000);
        ImageAssetLoader assets = new ImageAssetLoader(new ImageAssetConfig(root));
        PaperExportRequest paper = paper("tall.png");
        paper.getSections().getFirst().setImages(List.of("tall.png"));
        paper.getSections().getFirst().setImageMaxWidthPx(Integer.MAX_VALUE);
        LayoutDocumentRequest layout = layout("tall.png");
        LayoutDocumentRequest.BoundingBox box = new LayoutDocumentRequest.BoundingBox();
        box.setWidth(Integer.MAX_VALUE);
        layout.getPages().getFirst().getBlocks().getFirst().setBbox(box);
        for (byte[] docx : List.of(new DocxBuilder(false, assets).build(paper), new LayoutDocxBuilder(false, assets).build(layout))) {
            try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(docx))) {
                for (var paragraph : document.getParagraphs()) {
                    for (var run : paragraph.getRuns()) {
                        for (var drawing : run.getCTR().getDrawingList()) {
                            for (var inline : drawing.getInlineList()) {
                                assertTrue(inline.getExtent().getCx() > 0 && inline.getExtent().getCx() <= 8_000_000);
                                assertTrue(inline.getExtent().getCy() > 0 && inline.getExtent().getCy() <= 11_000_000);
                            }
                        }
                    }
                }
            }
        }
    }

    @Test
    void defaultBuildersFailClosedAndDoNotSilentlyOmitRequestedImages() throws Exception {
        ImageAssetLoaderTest.image(root.resolve("canary.png"), "png", 1, 1);
        withProperties(Map.of("root", ""), () -> {
            assertFailure("DISABLED", () -> new DocxBuilder(false).build(paper(root.resolve("canary.png").toString())));
            assertFailure("DISABLED", () -> new LayoutDocxBuilder(false).build(layout(root.resolve("canary.png").toString())));
        });
    }

    @Test
    void bothBuildersPropagateMissingCorruptOutsideAndBlankImageErrors() throws Exception {
        Files.writeString(root.resolve("fake.png"), "SYNTHETIC-ONLY-NOT-IMAGE");
        ImageAssetLoader assets = new ImageAssetLoader(new ImageAssetConfig(root));
        for (String path : Arrays.asList("missing.png", "fake.png", "../outside.png", "", null)) {
            assertThrows(ImageAssetException.class, () -> new DocxBuilder(false, assets).build(paper(path)));
            assertThrows(ImageAssetException.class, () -> new LayoutDocxBuilder(false, assets).build(layout(path)));
        }
    }

    @Test
    void sectionAndQuestionImagesBothAbortOnFailure() {
        ImageAssetLoader assets = new ImageAssetLoader(new ImageAssetConfig(root));
        PaperExportRequest paper = paper("missing.png");
        SectionDTO section = paper.getSections().getFirst();
        section.setImages(section.getQuestions().getFirst().getImages());
        section.getQuestions().getFirst().setImages(List.of());
        assertFailure("NOT_FOUND", () -> new DocxBuilder(false, assets).build(paper));
    }

    @Test
    void fallbackPageBackgroundUsesSamePolicyAndDoesNotSkipErrors() throws Exception {
        ImageAssetLoader assets = new ImageAssetLoader(new ImageAssetConfig(root));
        LayoutDocumentRequest request = layout("unused.png");
        request.getPages().getFirst().setBlocks(List.of());
        request.getPages().getFirst().setBackgroundImagePath("missing.png");
        assertFailure("NOT_FOUND", () -> new LayoutDocxBuilder(false, assets).build(request));
        ImageAssetLoaderTest.image(root.resolve("background.png"), "png", 4, 3);
        request.getPages().getFirst().setBackgroundImagePath("background.png");
        try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(new LayoutDocxBuilder(false, assets).build(request)))) {
            assertEquals(1, document.getAllPictures().size());
        }
    }

    @Test
    void springServiceAndDirectBuildersUseTheSameSystemPropertyPolicy() throws Exception {
        ImageAssetLoaderTest.image(root.resolve("canary.png"), "png", 4, 3);
        withProperties(Map.of("root", root.toString(), "windows-prefix", "J:/legacy/assets", "max-bytes", "100000", "max-pixels", "12"), () -> {
            PaperExportRequest paper = paper("J:\\legacy\\assets\\canary.png");
            LayoutDocumentRequest layout = layout("J:\\legacy\\assets\\canary.png");
            for (byte[] docx : List.of(new PaperExportService().export(paper), new LayoutExportService().exportLayoutDocument(layout),
                    new DocxBuilder(false).build(paper), new LayoutDocxBuilder(false).build(layout))) {
                try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(docx))) {
                    assertEquals(1, document.getAllPictures().size());
                }
            }
            System.setProperty("paperword.assets.max-pixels", "11");
            assertFailure("TOO_MANY_PIXELS", () -> new PaperExportService().export(paper));
            assertFailure("TOO_MANY_PIXELS", () -> new LayoutExportService().exportLayoutDocument(layout));
        });
    }

    @Test
    void invalidSystemPropertyLimitsAreNotSilentlyIgnored() throws Exception {
        withProperties(Map.of("root", root.toString(), "max-bytes", "invalid"), () ->
            assertFailure("CONFIG", () -> new DocxBuilder(false)));
    }

    private static PaperExportRequest paper(String path) {
        PaperExportRequest request = new PaperExportRequest();
        PaperExportRequest.PaperInfo info = new PaperExportRequest.PaperInfo();
        info.setName("Synthetic image policy test");
        request.setPaper(info);
        QuestionDTO question = new QuestionDTO();
        question.setContent("Synthetic diagram");
        question.setImages(Arrays.asList(path));
        SectionDTO section = new SectionDTO();
        section.setQuestions(List.of(question));
        request.setSections(List.of(section));
        return request;
    }

    private static LayoutDocumentRequest layout(String path) {
        LayoutDocumentRequest request = new LayoutDocumentRequest();
        LayoutDocumentRequest.Block block = new LayoutDocumentRequest.Block();
        block.setType(LayoutDocumentRequest.BlockType.IMAGE);
        block.setImagePath(path);
        LayoutDocumentRequest.Page page = new LayoutDocumentRequest.Page();
        page.setBlocks(List.of(block));
        request.setPages(List.of(page));
        return request;
    }

    private static void assertFailure(String code, org.junit.jupiter.api.function.Executable action) {
        assertEquals("IMAGE_ASSET_" + code, assertThrows(ImageAssetException.class, action).getCode());
    }

    private static void withProperties(Map<String, String> settings, ThrowingRunnable action) throws Exception {
        Map<String, String> old = new HashMap<>();
        // Explicitly blank the optional mapping so an unrelated developer environment cannot broaden it.
        Map<String, String> effective = new HashMap<>(settings);
        effective.putIfAbsent("windows-prefix", "");
        effective.putIfAbsent("max-bytes", Long.toString(ImageAssetConfig.DEFAULT_MAX_BYTES));
        effective.putIfAbsent("max-pixels", Long.toString(ImageAssetConfig.DEFAULT_MAX_PIXELS));
        for (Map.Entry<String, String> entry : effective.entrySet()) {
            String key = "paperword.assets." + entry.getKey();
            old.put(key, System.getProperty(key));
            System.setProperty(key, entry.getValue());
        }
        try {
            action.run();
        } finally {
            old.forEach((key, value) -> {
                if (value == null) System.clearProperty(key); else System.setProperty(key, value);
            });
        }
    }

    @FunctionalInterface private interface ThrowingRunnable {
        void run() throws Exception;
    }
}
