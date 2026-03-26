package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

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

    /**
     * 仅保留本轮需要人工核对的三个长除法样例，减少视觉干扰。
     */
    private QuestionDTO createQuestion() {
        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(6);
        question.setScore(10);
        question.setContent("""
            请重点检查以下长除法排版：<br/>
            1. 整数长除法首行是否按 LLM 步骤区位置渲染：<br/>
            $$\\longdiv[14300]{7}{100100}\\begin{array}{l}\\underline{7}\\\\30\\\\\\underline{28}\\\\\\text{  }21\\\\\\text{  }\\underline{21}\\\\\\text{      }00\\\\\\text{      }\\underline{00}\\\\\\text{        }0\\end{array}$$<br/>
            2. 小数长除法是否保留原图中的最终余数 0：<br/>
            $$\\longdiv[5]{2.5}{12.5}\\begin{array}{l}\\underline{125}\\\\\\text{  }0\\end{array}$$<br/>
            3. 带小数商的长除法是否保留显式商与最终余数 0：<br/>
            $$\\longdiv[0.44]{5}{2.2}\\begin{array}{l}\\underline{20}\\\\20\\\\\\underline{20}\\\\\\text{  }0\\end{array}$$
            """);
        return question;
    }
}
