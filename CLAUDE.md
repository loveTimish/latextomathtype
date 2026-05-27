# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and test commands

- Build runnable jar: `mvn clean package`
- Run all tests: `mvn test`
- Run one test class: `mvn -Dtest=LaTeXParserTest test`
- Run one test method: `mvn -Dtest=VerticalLayoutCompilerTest#shouldCompileExplicitLongDivisionHeaderAndSteps test`
- Run the app locally: `java -jar target/paper-to-word-1.0.0.jar`
- Dev mode via Spring Boot plugin: `mvn spring-boot:run`
- Evaluate the current artifact name if needed: `mvn -q -DskipTests help:evaluate -Dexpression=project.build.finalName -DforceStdout`

## What this project does

This is a Java 21 / Spring Boot 3 application that turns exam-paper or layout JSON into `.docx` files with editable MathType equations. The primary APIs are `POST /api/export/word` and `POST /api/export/layout-word`. The service can run in two modes:

- Pure Java mode: generate MathType-compatible OLE objects locally and embed them directly into the DOCX.
- Windows bridge mode: generate a draft DOCX with raw `$...$` markers, then send it to an external Windows service that uses local Word + MathType automation for final conversion.

The mode is controlled by `mathtype.windows.*` in `src/main/resources/application.yml`.

## High-level architecture

The top-level request flow is:

`ExportController -> PaperExportService -> DocxBuilder -> LaTeXParser + MathTypeEmbedder -> MtefWriter -> OlePackager`

Important pieces:

- `controller/ExportController.java`: HTTP entrypoint for paper export, layout export, and health endpoints.
- `service/PaperExportService.java`: chooses between pure-Java OLE generation and the Windows conversion service.
- `service/LayoutExportService.java`: exports `LayoutDocumentRequest` page/block layouts directly into DOCX while preserving OCR layout structure.
- `core/docx/DocxBuilder.java`: main orchestrator for building the Word document layout, question sections, options, answer areas, and inline/display math placement.
- `core/docx/LayoutDocxBuilder.java`: layout-preserving DOCX builder used by the OCR/PDF gateway service.
- `core/docx/MathTypeEmbedder.java`: creates `/word/embeddings/*.bin` OLE parts plus preview image parts, then injects the VML/OLE XML into the paragraph run.
- `core/ole/OlePackager.java`: packages MTEF bytes into a MathType OLE compound document.
- `core/render/LaTeXImageRenderer.java`: generates the preview image used before Word/MathType replaces it with native rendering.

## Math pipeline

The math pipeline is the core of the repository:

1. `LaTeXParser` extracts math from HTML/text, supporting `$...$`, `$$...$$`, `\(...\)`, `\[...\]`, and some bare LaTeX commands.
2. `LaTeXTokenizer` tokenizes LaTeX source.
3. `LaTeXParser` builds a `LaTeXNode` AST with recursive-descent parsing.
4. `MtefWriter` converts the AST to MTEF v5 bytes.
5. `OlePackager` wraps the MTEF payload as a MathType OLE object.
6. `MathTypeEmbedder` writes the OLE part and preview image into the DOCX package.

If you are changing formula behavior, read these together rather than in isolation:

- `core/latex/LaTeXParser.java`
- `core/latex/LaTeXNode.java`
- `core/mtef/MtefWriter.java`
- `core/mtef/MtefTemplateBuilder.java`
- `core/mtef/MtefCharMap.java`
- `core/docx/MathTypeEmbedder.java`

## Vertical arithmetic / long-division architecture

A substantial part of the codebase is dedicated to K12-style vertical layouts such as arithmetic columns, decimals, and long division.

The architecture is intentionally two-stage:

1. `core/layout/VerticalLayoutCompiler.java` normalizes ASTs for arrays / long-division expressions into a semantic `VerticalLayoutSpec`.
2. `core/layout/VerticalLayoutNodeFactory.java` converts that normalized layout back into AST structures that `MtefWriter` can serialize using matrix/pile/template records.

This means vertical math is not handled as a simple direct AST-to-MTEF mapping. If a vertical layout renders incorrectly, inspect the layout compiler/factory pair before changing low-level MTEF byte writing.

Key implication: many rendering issues that look like “MTEF bugs” are actually normalization or row/column/rule-span issues in the layout layer.

### Long-division editing rule

For long division, "looks close to the reference DOCX" and "remains editable in MathType" are not always the same goal.

- If MathType editability breaks, prefer editable structure over byte-level similarity to the reference shell.
- Do not put complex multi-step containers directly inside the `tmLDIV` dividend slot. That path can serialize, but MathType may refuse to edit it.
- The current stable approach is: keep the header as `tmLDIV`, and render the step area outside the template as ordinary rows that MathType can edit reliably.

### Current editable spacing strategy for long division

The current project uses a space-based editable layout for computed multi-step long division.

- Header: keep `tmLDIV` for divisor / quotient / dividend.
- Step area: emit ordinary text / underline rows outside the template instead of `RULER/tab-stop`.
- Leading spaces are calculated from the normalized vertical-layout columns rather than hardcoding every literal case.

Current spacing formula used by `VerticalLayoutNodeFactory`:

- Let `c` be the first non-empty column of the trimmed step row.
- Let `n` be the digit count of the visible row text.
- Leading spaces = `3 * c + n - 2`

Equivalent expansion:

- 1-digit row: `3 * c - 1`
- 2-digit row: `3 * c`
- 3-digit row: `3 * c + 1`

This rule was introduced specifically to make rows such as `10 / 23 / 20 / 34 / 30 / 4` line up in editable MathType output without relying on nested pile/ruler structures.

### Latest conclusion on single-block long division alignment

After switching OCR output to the new single-block form:

- `$$\longdiv[...]{}\n\begin{array}{l}...\end{array}$$`

the main remaining alignment issue is currently considered a **program-side layout problem**, not primarily an LLM recognition problem.

Observed conclusion:

- The OCR/LLM side is already preserving the original long-division steps and keeping them inside one display-math block.
- The current export path reconstructs the step area as single-column text rows with leading spaces (for example via `\text{   }42`), so final alignment depends on text-space rendering rather than preserved column semantics.
- Because of that, Word/MathType can show visible drift even when the OCR step order and space counts are correct.

Practical implication for future fixes:

- Do not ask the LLM to invent extra steps or to compensate for alignment issues by changing the source steps.
- Prefer fixing alignment in code by restoring explicit column/anchor semantics for the step rows, instead of relying only on raw text spaces inside the composite long-division block.
- Keep the rule that bare `\longdiv` must not auto-generate steps; only the original image steps should appear.

### Cross-multiplication editing rule

For concentration cross-multiplication / "十字交叉", MathType editability depends more on using MathType-known structures and fonts than on visually similar plain-text arrows.

- Do not use custom arrow fonts such as `Segoe UI Symbol` or `DejaVu Sans` for diagonal arrows in editable equations.
- Do not add extra font-definition records between the template prefix and the top-level `LINE` just for cross-multiplication arrows.
- The current stable approach is: keep the formula on the normal template-prefix path, detect the `5 x 5` cross-multiplication array in `VerticalLayoutCompiler`, and rebuild it as nested preserved matrices in `VerticalLayoutNodeFactory`.
- The nested structure should stay close to the reference MathML: outer vertical stack, each row as a `1 x 3` matrix, with the left/right positions themselves represented as `1 x 2` sub-matrices.
- Diagonal arrows should use MathType's official `MT Extra` slot (`FN_MTEXTRA`), matching the reference OLE in `J:\\lingzhi\\data\\十字交叉.docx`.

Key implication: if the cross opens in MathType but shows `?`, inspect the arrow typeface/slot first; if it opens but layout is wrong, inspect the nested-matrix reconstruction path first.

## MTEF writer details that matter

- `MtefWriter` prefers to reuse a known-good MTEF prefix extracted from `src/main/resources/mathtype-template/oleObject-template.bin` rather than hand-building every header byte.
- Long division has a special serialization path and does not follow the exact same outer LINE structure as ordinary expressions.
- The writer collaborates with `VerticalLayoutCompiler`, `VerticalLayoutNodeFactory`, and `MtefPileRulerWriter`; changes in one often require tests in the others.

When touching MTEF output, also review `TECHNICAL.md`, which explains the expected binary structure and DOCX/OLE embedding model.

## DOCX generation details that matter

- `DocxBuilder` parses HTML question content and splits `<br/>`-style content into separate paragraphs so stacked/vertical math does not get forced onto the same line as prose.
- `MathTypeEmbedder` computes run baseline offsets (`w:position`) heuristically based on formula type/height; inline alignment regressions are often here rather than in parser/MTEF code.
- The generated DOCX uses explicit package parts and relationships, not only high-level XWPF helpers, so package-level XML details matter.

## Configuration

`src/main/resources/application.yml` contains the important runtime settings:

- server port defaults to `8081`
- package logging for `com.lz.paperword` defaults to `DEBUG`
- `mathtype.windows.enabled=false` keeps the pure-Java embedding path
- `mathtype.windows.endpoint`, `timeout-seconds`, and `api-key` configure the optional Windows conversion bridge

## Tests to know about

The test suite is concentrated around parser, layout, rendering, and binary-format regressions.

Useful anchors:

- `core/latex/LaTeXParserTest`
- `core/layout/VerticalLayoutCompilerTest`
- `core/mtef/MtefWriterTest`
- `core/docx/MathTypeAlignmentRegressionTest`
- `tools/K12FormulaShowcaseDocxTest`
- `tools/GenerateVisibleVerticalWordTest`

There are also many focused debug/regression tests for long division and MTEF output. Prefer running a targeted test class or method when working in these areas.

Useful validation commands:

- Generate the current K12 showcase DOCX: `mvn "-Dtest=K12FormulaShowcaseDocxTest" test`
- Generate the focused vertical-layout sample DOCX: `mvn "-Dtest=GenerateVisibleVerticalWordTest" test`
- Run focused writer/layout/docx regressions: `mvn "-Dtest=VerticalLayoutCompilerTest,MtefWriterTest,MathTypeAlignmentRegressionTest" test`
