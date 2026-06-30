# latextomathtype

`latextomathtype` is a Java/Spring Boot service for exporting exam data and LaTeX formulas to Word `.docx` files with editable MathType OLE equations.

It is not a formula screenshot tool. The core goal is to generate Word documents whose formulas can still be opened by MathType on Windows, while preserving enough semantic structure for `docx2tex` round-trip validation.

```text
PaperExportRequest
  -> LaTeX tokenizer/parser
  -> Math IR
  -> MTEF v5 writer
  -> OLE2 MathType object
  -> WMF preview and Word display box
  -> DOCX
```

## Status

This repository currently focuses on production-style Word export for K12 math papers and DOCX reconstruction experiments.

| Area | Current state |
| --- | --- |
| Runtime | Spring Boot 3.3, Java 21 |
| Main output | `.docx` with MathType-compatible OLE formulas |
| Formula body | Pure Java MTEF/OLE writer, no desktop MathType dependency in the service path |
| OLE preview | Strict vector-preview path for MathType objects; preview failure is treated as an error |
| Validation | Java tests, OLE/POIFS inspection, Word/MathType spot checks, `docx2tex` round trips, physical-size comparison |
| Linux | Supported for service/runtime validation; final GUI editability checks still require Windows + Word + MathType |

## What It Does

- Exports `PaperExportRequest` JSON to Word through `POST /api/export/word`.
- Embeds formulas as MathType-compatible OLE2 objects with an `Equation Native` MTEF stream.
- Separates editable formula data from the visible Word display box, so layout can be calibrated without hardcoding formula font size inside MTEF.
- Generates OLE preview media as WMF and checks that generated DOCX files do not silently degrade to bitmap-only formula output.
- Supports a layout-oriented Word export endpoint for OCR/PDF reconstruction workflows.
- Provides reproducible validation scripts for reference DOCX round-trip, xsc corpus acceptance, OLE inspection, and `docx2tex` coverage.

## What It Is Not

- It does not try to produce byte-identical DOCX files.
- It does not require desktop MathType to run the server.
- It does not claim full LaTeX coverage. Unsupported or weakly covered constructs should be added with parser, MTEF, preview, and round-trip tests.
- It does not use PNG screenshots as the primary MathType equation representation.

## Requirements

| Tool | Use |
| --- | --- |
| JDK 21 | Build, test, and run the service |
| Maven 3.9.x | Build and test; the repository includes `.mvn/apache-maven-3.9.12` |
| Node.js | MathJax SVG worker for formula preview generation |
| `mathjax-full` | Installed by `npm install`; used by `tools/mathjax/render_mathjax_svg.cjs` |
| TeX Live or MiKTeX | Native TeX rendering path where enabled, including `latex` and `dvisvgm` |
| Python 3 | Reference rebuild and acceptance scripts |
| `docx2tex` | DOCX to LaTeX round-trip validation |
| Microsoft Word + MathType | Windows GUI spot check for double-click editability |
| Docker | Optional Linux/container smoke validation |

Windows validation scripts currently assume `docx2tex` is available at `J:\docx2tex\d2t.bat`. If your machine uses another path, update the script parameter or local script configuration before running full validation.

## Quick Start

Install the Node dependency used by the MathJax preview worker:

```powershell
npm install
npm run mathjax:smoke
```

Build the Java service with the bundled Maven:

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd clean package
```

Run the test suite:

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd test
```

Start the service:

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd spring-boot:run
```

Check the health endpoint:

```powershell
Invoke-WebRequest http://127.0.0.1:8081/api/export/health
```

Export the sample paper:

```powershell
Invoke-WebRequest `
  -Method Post `
  -Uri http://127.0.0.1:8081/api/export/word `
  -ContentType 'application/json; charset=utf-8' `
  -InFile exam-template.json `
  -OutFile target\paper.docx
```

Linux uses the same Maven wrapper path:

```bash
./.mvn/apache-maven-3.9.12/bin/mvn clean package
java -jar target/paper-to-word-1.0.0.jar
```

## API

Base path: `/api/export`

| Endpoint | Method | Response | Purpose |
| --- | --- | --- | --- |
| `/health` | `GET` | Text | Service health check |
| `/word` | `POST` | `.docx` | Export a structured paper request to Word with MathType OLE formulas |
| `/layout-word` | `POST` | `.docx` | Export block-level layout data for OCR/PDF reconstruction workflows |

Minimal request shape:

```json
{
  "paper": {
    "name": "2026 Math Paper",
    "subjectType": 1,
    "stage": 2,
    "score": 100,
    "suggestTime": 90
  },
  "sections": [
    {
      "headline": "I. Multiple Choice",
      "questions": [
        {
          "serialNumber": 1,
          "questionType": 1,
          "content": "Solve $x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}$.",
          "options": [
            { "prefix": "A", "content": "$1$" },
            { "prefix": "B", "content": "$2$" }
          ],
          "correct": "A",
          "score": 5,
          "analyze": "Example analysis with $\\frac{1}{2}$."
        }
      ]
    }
  ]
}
```

See [exam-template.json](exam-template.json) for a fuller Chinese exam sample.

## Architecture

| Module | Responsibility |
| --- | --- |
| `controller` | REST endpoints for standard and layout-oriented Word export |
| `service` | Request orchestration and export service boundaries |
| `core/latex` | Text/formula splitting, tokenization, and LaTeX parsing |
| `core/mathml` | Intermediate math representation and normalization |
| `core/mtef` | MathType MTEF v5 records, templates, character mapping, and writer tests |
| `core/ole` | OLE2 compound object packaging |
| `core/render` | MathJax/TeX preview rendering, WMF generation, and render caching |
| `core/docx` | DOCX construction, OLE embedding, VML shape sizing, and baseline placement |
| `core/layout` | Block-level layout export support |

The important design boundary is:

```text
MTEF/OLE body       controls MathType editability and semantic equation structure
Word display box   controls visible width, height, and baseline on the page
WMF preview        controls what Word paints before the OLE object is opened
```

Changing formula size inside MTEF is treated as a last resort. Most visual calibration belongs in the Word object display box and preview layer.

## Validation

The main acceptance chain is:

```text
reference DOCX
  -> docx2tex
  -> PaperExportRequest
  -> regenerated DOCX
  -> OLE, MathType, docx2tex, WMF, and layout checks
```

Run the reference round-trip on Windows:

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

The script rebuilds:

```text
target/reference-roundtrip/fraction-split-reference-regenerated.docx
```

It also writes comparison artifacts under:

```text
target/reference-roundtrip
```

Targeted checks:

```powershell
.\scripts\verify-mathtype-word.ps1 `
  -DocxPath target\reference-roundtrip\fraction-split-reference-regenerated.docx `
  -MinimumOleCount 400

.\scripts\verify-docx2tex-roundtrip.ps1 `
  -DocxPath target\reference-roundtrip\fraction-split-reference-regenerated.docx `
  -RequestJsonPath target\reference-roundtrip\fraction-split-reference.request.json `
  -OutDir target\reference-roundtrip\regenerated-docx2tex `
  -MinimumTimesCount 1000 `
  -MinimumCdotsCount 150 `
  -MinimumFractionCount 1500
```

Run only the reference-generation Java test:

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd `
  -q `
  -Dtest=com.lz.paperword.tools.ReferenceRoundTripDocxTest `
  test
```

## xsc Corpus Acceptance

The xsc pipeline validates full DOCX reconstruction, not only isolated formula export. The default source corpus is:

```text
F:\资料\xsc资料\word_files
```

Typical batch flow:

```powershell
.\scripts\run_xsc_docxtolatex.ps1 `
  -Start 1 `
  -End 155 `
  -OutRoot D:\latextomathtype\analysis\xsc-latex

python .\scripts\make_full_batch10_requests.py `
  --start 1 `
  --end 155 `
  --latex-root D:\latextomathtype\analysis\xsc-latex `
  --out-dir D:\latextomathtype\analysis\batch10-full-requests

.\scripts\run_xsc_acceptance.ps1 `
  -Start 141 `
  -End 155 `
  -LatexRoot D:\latextomathtype\analysis\xsc-latex

python .\scripts\summarize_xsc_full_acceptance.py
```

Recent full-corpus acceptance summary:

| Check | Result |
| --- | --- |
| Covered source range | `1..155` |
| Paired formula objects | `39551` |
| WMF width within 1% | `39551/39551` |
| WMF height within 1% | `39551/39551` |
| Generated non-WMF previews | `0` |
| MTEF clean pairs | `39248/39551` |
| Remaining hard suspects | `75` |
| Remaining low-tail failures | `51` |
| Remaining structural gaps | `126` |

The summary report is expected at:

```text
D:\latextomathtype\analysis\acceptance-summary\xsc-full-acceptance.json
```

## Linux And Docker

Package and run the executable jar:

```bash
./.mvn/apache-maven-3.9.12/bin/mvn -DskipTests package
java \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render \
  -jar target/paper-to-word-1.0.0.jar
```

Build and run the Docker image:

```bash
docker build -t latextomathtype:local .
docker run --rm -p 8081:8081 \
  -v latextomathtype-cache:/var/cache/latextomathtype/formula-render \
  latextomathtype:local
```

Run the Linux smoke script:

```bash
sh scripts/linux-smoke.sh
```

See [docs/linux-runtime.md](docs/linux-runtime.md) for package requirements and container smoke alternatives.

## Configuration

Most runtime knobs are Java system properties:

| Property | Default | Purpose |
| --- | --- | --- |
| `paperword.latex.command` | `latex` | Native LaTeX command path |
| `paperword.xelatex.command` | `xelatex` | XeLaTeX command path for CJK formulas where used |
| `paperword.dvisvgm.command` | `dvisvgm` | DVI/SVG conversion command path |
| `paperword.latex.timeout.seconds` | `20` | External render command timeout |
| `paperword.mathjax.node.command` | `node` | Node.js command used by the MathJax worker |
| `paperword.mathjax.script` | `tools/mathjax/render_mathjax_svg.cjs` | MathJax worker script |
| `paperword.render.cache.enabled` | `true` | Persistent formula render cache switch |
| `paperword.render.cache.dir` | `data/cache/formula-render` | Persistent formula render cache location |

The Spring Boot port is configured in [src/main/resources/application.yml](src/main/resources/application.yml) and currently defaults to `8081`.

## Repository Layout

```text
src/main/java/com/lz/paperword
  controller/        REST API
  service/           Export orchestration
  core/docx/         DOCX building and MathType embedding
  core/latex/        LaTeX splitting, tokenization, and parsing
  core/mathml/       Intermediate math representation
  core/mtef/         MTEF v5 writer and MathType records
  core/ole/          OLE2 packaging
  core/render/       Preview rendering, WMF generation, and cache
  core/layout/       Block layout export
  model/             Request DTOs

rebuild/             Reference reconstruction and layout comparison tools
rebuild-assets/      Reference DOCX files and extracted visual assets
scripts/             Validation, corpus, smoke, and inspection scripts
tools/mathjax/       MathJax SVG worker
docs/                Technical notes and validation plans
```

## Documentation

- [TECHNICAL.md](TECHNICAL.md): architecture, MTEF/OLE model, rendering boundary, and reference rebuild flow.
- [docs/linux-runtime.md](docs/linux-runtime.md): Linux and Docker runtime details.
- [docs/MathType-validation-plan.md](docs/MathType-validation-plan.md): validation layers and phase gates.
- [docs/MathType-support-matrix.md](docs/MathType-support-matrix.md): supported formula structures and known gaps.
- [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md): xsc corpus asset pipeline notes.

## Related Projects And References

- [transpect/docx2tex](https://github.com/transpect/docx2tex): DOCX to LaTeX conversion used as a round-trip validation tool.
- [plutext/docx4j](https://github.com/plutext/docx4j): a mature Java OpenXML library whose README structure is a useful contrast for quick project positioning.
- [MathJax](https://github.com/mathjax/MathJax-src): TeX/MathML/AsciiMath rendering engine used here through a local Node worker.
- [WIRIS MathType SDK: MTEF storage](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/how-mtef-is-stored-in-files-and-objects): reference for MathType native stream storage.
- [WIRIS MathType SDK: MTEF v5](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/mathtype-mtef-v5-mathtype-40-and-later): reference for MathType record structure.

## Boundaries

Generated documents aim for visual and semantic equivalence, not binary identity. Reference and regenerated files may differ in package internals, relationship IDs, media ordering, and paragraph internals.

The acceptance target is more practical:

- Formula objects remain MathType-compatible OLE objects.
- `Equation Native` streams are structurally valid.
- Word displays formulas at the expected width, height, and baseline.
- `docx2tex` can recover the expected LaTeX fragments.
- Windows + Word + MathType can open representative generated formulas for editing.
