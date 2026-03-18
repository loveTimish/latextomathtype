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

class GenerateVisibleVerticalWordTest {

    @Test
    void shouldGenerateVisibleWordForVerticalLayouts() throws IOException {
        DocxBuilder builder = new DocxBuilder();
        byte[] docx = builder.build(buildRequest());

        Path outputDir = Path.of("target", "generated-docs");
        Files.createDirectories(outputDir);

        // 使用时间戳生成唯一文件名，避免 Windows 下旧文档仍被占用时覆盖失败。
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss"));
        Path output = outputDir.resolve("vertical-layouts-visible-braced-template-real-steps-" + timestamp + ".docx");
        Files.write(output, docx);

        // 固定别名仅用于方便人工查找；若文件正被占用则忽略，不影响本次生成结果。
        Path latestAlias = outputDir.resolve("vertical-layouts-visible-braced-template-real-steps.docx");
        try {
            Files.copy(output, latestAlias, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
            // 别名写入失败通常是目标文件仍在 Word 中打开，本次唯一文件已成功生成即可。
        }

        assertTrue(Files.exists(output), "generated docx should exist");
        assertTrue(Files.size(output) > 0, "generated docx should not be empty");
    }

    private PaperExportRequest buildRequest() {
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("竖式展示样例");
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("一、竖式样例");

        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(6);
        question.setScore(20);
        question.setContent("""
            整数加法：<br/>
            $$\\begin{array}{rrrr}{} & {1} & {2} & {3} \\\\ {+} & {4} & {5} & {6} \\\\ \\hline {} & {5} & {7} & {9}\\end{array}$$<br/>
            整数减法：<br/>
            $$\\begin{array}{rrrr}{} & {8} & {6} & {4} \\\\ {-} & {2} & {7} & {9} \\\\ \\hline {} & {5} & {8} & {5}\\end{array}$$<br/>
            整数乘法：<br/>
            $$\\begin{array}{rrrrrr}{} & {} & {1} & {2} & {3} & {} \\\\ {\\times} & {} & {} & {4} & {5} & {} \\\\ \\hline {} & {} & {6} & {1} & {5} & {} \\\\ {+} & {4} & {9} & {2} & {} & {} \\\\ \\hline {} & {5} & {5} & {3} & {5} & {}\\end{array}$$<br/>
            小数加法：<br/>
            $$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {+} & {3} & {.} & {75} \\\\ \\hline {} & {16} & {.} & {25}\\end{array}$$<br/>
            小数减法：<br/>
            $$\\begin{array}{rcrl}{} & {12} & {.} & {50} \\\\ {-} & {3} & {.} & {75} \\\\ \\hline {} & {8} & {.} & {75}\\end{array}$$<br/>
            显式长除法（两步过程）：<br/>
            $$\\begin{array}{l}{\\longdiv[65]{13}{845}} \\\\ {\\begin{array}{rrr}{7} & {8} & {} \\\\ \\hline {} & {6} & {5} \\\\ {} & {6} & {5} \\\\ \\hline {} & {} & {0}\\end{array}}\\end{array}$$<br/>
            显式长除法（三步过程）：<br/>
            $$\\longdiv[246]{5}{1234}$$<br/>
            参考三步模板（对照用）：<br/>
            $$\\begin{array}{l}{\\longdiv[246]{5}{1234}} \\\\ {\\begin{array}{rrrr}{1} & {2} & {3} & {4} \\\\ \\hline {} & {1} & {0} & {} \\\\ {} & {2} & {3} & {} \\\\ \\hline {} & {2} & {0} & {} \\\\ {} & {3} & {4} & {} \\\\ \\hline {} & {3} & {0} & {} \\\\ {} & {} & {4} & {}\\end{array}}\\end{array}$$
            """);

        section.setQuestions(List.of(question));
        request.setSections(List.of(section));
        return request;
    }
}
