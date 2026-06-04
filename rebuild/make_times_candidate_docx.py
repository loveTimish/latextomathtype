# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import subprocess
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_JAVA = PROJECT_ROOT / "src/test/java/com/lz/paperword/tools/TimesCandidateDocxTest.java"


JAVA_SOURCE = r'''package com.lz.paperword.tools;

import com.lz.paperword.core.docx.DocxBuilder;
import com.lz.paperword.model.PaperExportRequest;
import com.lz.paperword.model.QuestionDTO;
import com.lz.paperword.model.SectionDTO;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

class TimesCandidateDocxTest {
    @Test
    void generateTimesCandidateDocx() throws Exception {
        String output = System.getProperty("latextomathtype.times.candidate.output",
            "target/reference-roundtrip/times-candidate.docx");
        PaperExportRequest request = new PaperExportRequest();

        PaperExportRequest.PaperInfo paper = new PaperExportRequest.PaperInfo();
        paper.setName("times candidate");
        paper.setSubjectType(2);
        paper.setStage(2);
        paper.setScore(1);
        paper.setSuggestTime(1);
        request.setPaper(paper);

        SectionDTO section = new SectionDTO();
        section.setHeadline("candidate");

        QuestionDTO question = new QuestionDTO();
        question.setSerialNumber(1);
        question.setQuestionType(5);
        question.setScore(1);
        question.setContent("<p>$\\frac{1}{1\\times2}+\\frac{1}{2\\times3}+\\frac{99}{1\\times2\\times3\\times\\cdots\\times100}$</p>");
        section.setQuestions(List.of(question));
        request.setSections(List.of(section));

        byte[] docx = new DocxBuilder(true).build(request);
        Path out = Path.of(output);
        Files.createDirectories(out.getParent());
        Files.write(out, docx);
        assertTrue(Files.size(out) > 1000);
    }
}
'''


def write_java_test() -> None:
    if TEMPLATE_JAVA.exists() and TEMPLATE_JAVA.read_text(encoding="utf-8") == JAVA_SOURCE:
        return
    TEMPLATE_JAVA.write_text(JAVA_SOURCE, encoding="utf-8")


def inspect_docx(docx: Path) -> dict[str, object]:
    with zipfile.ZipFile(docx) as zf:
        document_xml = zf.read("word/document.xml").decode("utf-8")
        embeddings = [name for name in zf.namelist() if name.startswith("word/embeddings/")]
    return {
        "docx": str(docx.resolve()),
        "bytes": docx.stat().st_size,
        "ole_objects": document_xml.count("<o:OLEObject"),
        "embeddings": len(embeddings),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--encoding", default="reference")
    parser.add_argument("--mode", default="")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--mvn", default=r"C:\Users\11703\.codex\tools\apache-maven-3.9.9\bin\mvn.cmd")
    args = parser.parse_args()

    write_java_test()
    cmd = [
        args.mvn,
        "-q",
        "-Dtest=com.lz.paperword.tools.TimesCandidateDocxTest",
        f"-Dlatextomathtype.mtef.times.encoding={args.encoding}",
        f"-Dlatextomathtype.times.candidate.output={args.output}",
        "test",
    ]
    if args.mode:
        cmd.insert(3, f"-Dlatextomathtype.mtef.times.mode={args.mode}")
    subprocess.run(cmd, cwd=PROJECT_ROOT, check=True)
    print(json.dumps(inspect_docx(PROJECT_ROOT / args.output), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
