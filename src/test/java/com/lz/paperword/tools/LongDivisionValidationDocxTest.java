package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import com.lz.paperword.core.mtef.MtefRecordNormalizer;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 生成本轮长除法回归样例文档，便于直接在 Word 中检查首行对齐与小数收尾。
 */
class LongDivisionValidationDocxTest {

    @Test
    void shouldGenerateLongDivisionValidationDocx() throws IOException {
        DocxBuilder builder = new DocxBuilder();
        byte[] docx = builder.build(buildRequest());

        Path outputDir = Path.of("target", "generated-docs");
        Files.createDirectories(outputDir);

        // 使用时间戳避免覆盖正在打开的旧文件，方便连续比对多次修正结果。
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path output = outputDir.resolve("longdiv-feedback-validation-" + timestamp + ".docx");
        Files.write(output, docx);

        // 固定别名仅用于快速定位；若 Word 占用旧文件导致覆盖失败，也不影响本次新文件生成。
        Path latestAlias = outputDir.resolve("longdiv-feedback-validation-latest.docx");
        try {
            Files.copy(output, latestAlias, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
            // 别名写入失败时保留时间戳文件即可。
        }

        System.out.println("Generated long division validation document: " + output);
        assertTrue(Files.exists(output), "生成的长除法验证文档应存在");
        assertTrue(Files.size(output) > 0, "生成的长除法验证文档不应为空");
        assertEditableVectorObjects(docx, 4);
    }

    private void assertEditableVectorObjects(byte[] docx, int expectedCount) throws IOException {
        int wmfCount = 0;
        List<byte[]> oleObjects = new ArrayList<>();
        try (ZipInputStream zip = new ZipInputStream(new ByteArrayInputStream(docx))) {
            ZipEntry entry;
            while ((entry = zip.getNextEntry()) != null) {
                String name = entry.getName();
                if (name.startsWith("word/media/") && name.endsWith(".wmf")) {
                    wmfCount++;
                } else if (name.startsWith("word/embeddings/") && name.endsWith(".bin")) {
                    oleObjects.add(zip.readAllBytes());
                }
            }
        }
        assertEquals(expectedCount, wmfCount, "每条长除法都应使用 WMF 预览");
        assertEquals(expectedCount, oleObjects.size(), "每条长除法都应嵌入可编辑 OLE");
        for (byte[] ole : oleObjects) {
            assertEquationDsmt4(ole);
        }
    }

    private void assertEquationDsmt4(byte[] ole) throws IOException {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(ole))) {
            byte[] compObj = read((DocumentEntry) fs.getRoot().getEntry("\u0001CompObj"));
            byte[] nativeStream = read((DocumentEntry) fs.getRoot().getEntry("Equation Native"));
            assertTrue(new String(compObj, StandardCharsets.ISO_8859_1).contains("Equation.DSMT4"),
                "OLE ProgID 应为 Equation.DSMT4");
            assertTrue(nativeStream.length >= 28, "Equation Native 必须包含标准 28 字节头");
            assertEquals(28, Short.toUnsignedInt(ByteBuffer.wrap(nativeStream, 0, 2)
                .order(ByteOrder.LITTLE_ENDIAN).getShort()), "Equation Native 头长度错误");
            int mtefLength = ByteBuffer.wrap(nativeStream, 8, 4).order(ByteOrder.LITTLE_ENDIAN).getInt();
            assertEquals(28 + mtefLength, nativeStream.length, "Equation Native 载荷长度错误");
            byte[] mtef = java.util.Arrays.copyOfRange(nativeStream, 28, nativeStream.length);
            MtefRecordNormalizer.NormalizationReport normalized = MtefRecordNormalizer.normalize(mtef);
            if (Boolean.getBoolean("paperword.longdivision.nativeHeaderCandidate")) {
                assertTrue(normalized.canonicalSignature().contains("TMPL:26:"),
                    "原生候选必须包含 MathType 专用 tmLDIV 模板");
                return;
            }
            assertTrue(normalized.canonicalSignature().contains("TMPL:26:1"),
                "结构化长除法必须使用 MathType 专用 tmLDIV 模板和可编辑商槽位");
            assertTrue(normalized.recordCounts().getOrDefault("PILE", 0) >= 1,
                "结构化长除法必须使用右对齐 PILE 组合头部与步骤");
            assertTrue(normalized.recordCounts().getOrDefault("MATRIX", 0) >= 1,
                "结构化长除法必须使用 MathType 可编辑的 MATRIX 容器");
            assertEquals(0, normalized.recordCounts().getOrDefault("RULER", 0),
                "结构化长除法 OLE 不得包含 MathType 无法激活的 RULER/TAB 布局");
        }
    }

    private byte[] read(DocumentEntry entry) throws IOException {
        try (DocumentInputStream input = new DocumentInputStream(entry)) {
            return input.readAllBytes();
        }
    }

    private PaperExportRequest buildRequest() {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("长除法回归验证");
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、长除法版式验证");
        section.setQuestions(List.of(createQuestion()));
        request.setSections(List.of(section));
        return request;
    }

    /** Structured examples are explicit layout data; none of these rows are computed locally. */
    private QuestionDTO createQuestion() {
        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(6);
        question.setScore(10);
        question.setContent("""
            请检查以下结构化长除法的商、逐行右对齐、横线跨度与编辑性：<br/>
            1. 整数：<br/>
            $$\\begin{longdivision}{rrrr}{6}{570}{3420}&30\\\\\\cline{1-2}&&42\\\\&&42\\\\\\cline{2-3}&&&0\\end{longdivision}$$<br/>
            2. 小数与空商：<br/>
            $$\\begin{longdivision}{rrrr}{2.5}{}{12.5}&&&125\\\\\\cline{1-4}&&&0\\end{longdivision}$$<br/>
            3. 嵌套分数、根式和特殊符号：<br/>
            $$\\begin{longdivision}{rrrr}{\\sqrt{2}}{\\frac{3}{2}}{3\\sqrt{2}}&&&\\frac{6}{\\sqrt{2}}\\\\\\cline{1-4}&&&\\alpha\\end{longdivision}$$<br/>
            4. 多项式：<br/>
            $$\\begin{longdivision}{rrrr}{x-1}{x+1}{x^{2}-1}&&&x^{2}-x\\\\\\cline{1-4}&&&x-1\\\\&&&x-1\\\\\\cline{2-4}&&&0\\end{longdivision}$$
            """);
        return question;
    }
}
