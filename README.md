# latextomathtype

[中文说明](README.zh-CN.md)

Export exam data and LaTeX formulas to Word `.docx` with editable MathType OLE equations.

This is not a formula screenshot generator. The service writes MathType-compatible OLE objects, generates Word-visible previews, and validates the result with OLE inspection, Word/MathType spot checks, and `docx2tex` round trips.

```text
PaperExportRequest -> LaTeX parser -> Math IR -> MTEF v5 -> OLE2 -> MathJax/Batik EMF+ Dual preview -> DOCX
```

## Highlights

- Java 21 / Spring Boot service for Word paper export.
- `POST /api/export/word` returns a `.docx` with editable MathType formulas.
- Pure Java MTEF/OLE writer; the server path does not require desktop MathType.
- Formula body, strict vector EMF+ Dual preview, and Word display box are handled as separate layers.
- The default preview path outlines MathJax SVG through Batik and emits matching EMF+ and classic EMF paths; it contains no bitmap or text records and has no playback-time font dependency.
- Linux runtime is supported; final GUI editability checks use Windows + Word + MathType.

## Quick Start

For local development, install the locked MathJax dependency:

```powershell
npm install
npm run mathjax:smoke
```

Release deployments use the offline sidecar and do not run npm or access the network at runtime:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-vector-sidecar.ps1 -Platform all
```

Extract `vector-sidecar-windows-x64.zip` or `vector-sidecar-linux-x64.tar.gz` beside the application
JAR so that the directory is named `vector-sidecar`. The runtime validates Node `v24.9.0`, MathJax
`3.2.2`, and the worker bundle hash before accepting any render result.

Build, test, and run:

```powershell
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd clean package
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd test
.\.mvn\apache-maven-3.9.12\bin\mvn.cmd spring-boot:run
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

## API

| Endpoint | Method | Purpose |
| --- | --- | --- |
| `/api/export/health` | `GET` | Service health check |
| `/api/export/word` | `POST` | Export `PaperExportRequest` to Word |
| `/api/export/layout-word` | `POST` | Export block-level layout data to Word |

See [exam-template.json](exam-template.json) for the request shape.

## Core Design

The project deliberately keeps three concerns separate:

| Layer | Controls |
| --- | --- |
| MTEF/OLE body | MathType editability and formula semantics |
| EMF+ Dual vector preview | What Word paints before the OLE object is opened; MathJax/Batik outlines are stored as matching EMF+ and classic EMF paths, with bitmap and text records forbidden |
| Word display box | Visible width, height, and baseline on the page |

That boundary is the main reason layout tuning happens in the Word object shell instead of by forcing point-size records into the MTEF body.

## Known Limitation

- Complete long-division layout is not implemented. `\longdiv[quotient]{divisor}{dividend}` currently preserves only an editable quotient/divisor/dividend header. A following `array` is merely an adjacent caller-supplied structure and is not claimed as validated long-division semantics. Automatic subtraction rows, digit carry-down, underline alignment, and remainder placement remain unsupported; use ordinary division or an explicit quotient-remainder identity in production input.

## Validation

Main reference round trip:

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

This checks the regenerated reference DOCX through OLE inspection, Word/MathType probes where available, `docx2tex` coverage, and formula display-box comparison.

Full xsc corpus acceptance is documented in [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md). The latest recorded run covered all `551` reconstructed source documents and `93,319` trace-matched formula objects: all `93,319` passed OLE validation, MTEF parsing/balance/normalized-structure matching, and strict EMF+ Dual validation, with `0` failed documents and `0` failed formulas. The report also records `37` explicit recoveries of source `U+FFFD` replacement characters; none were rendered silently.

## Linux And Docker

Package and run on Linux:

```bash
./.mvn/apache-maven-3.9.12/bin/mvn -DskipTests package
java \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render \
  -jar target/paper-to-word-1.0.0.jar
```

Docker and package details are in [docs/linux-runtime.md](docs/linux-runtime.md).

## Project Layout

```text
src/main/java/com/lz/paperword
  controller/ service/ model/
  core/latex/ core/mathml/ core/mtef/ core/ole/ core/render/ core/docx/

scripts/   validation and corpus tools
rebuild/   reference reconstruction tools
docs/      technical notes and validation plans
```

## Documentation

- [TECHNICAL.md](TECHNICAL.md): architecture and implementation details.
- [docs/MathType-validation-plan.md](docs/MathType-validation-plan.md): validation layers and phase gates.
- [docs/MathType-support-matrix.md](docs/MathType-support-matrix.md): supported formula structures and gaps.
- [docs/linux-runtime.md](docs/linux-runtime.md): Linux and Docker runtime notes.
- [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md): xsc corpus pipeline notes.

## References

- [transpect/docx2tex](https://github.com/transpect/docx2tex)
- [MathJax](https://github.com/mathjax/MathJax-src)
- [WIRIS MathType SDK: MTEF storage](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/how-mtef-is-stored-in-files-and-objects)
- [WIRIS MathType SDK: MTEF v5](https://docs.wiris.com/en_US/mathtype-sdk-technical-documentation/mathtype-mtef-v5-mathtype-40-and-later)
