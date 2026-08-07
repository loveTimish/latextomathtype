# latextomathtype

[中文说明](README.zh-CN.md)

Export exam data and LaTeX formulas to Word `.docx` with editable MathType OLE equations.

This is not a formula screenshot generator. The service writes MathType-compatible OLE objects, generates Word-visible previews, and validates the result with OLE inspection, Word/MathType spot checks, and `docx2tex` round trips.

```text
PaperExportRequest -> LaTeX parser -> Math IR -> MTEF v5 -> OLE2 -> pure-vector WMF preview -> DOCX
```

## Highlights

- Java 21 / Spring Boot service for Word paper export.
- `POST /api/export/word` returns a `.docx` with editable MathType formulas.
- Pure Java MTEF/OLE writer; the server path does not require desktop MathType.
- Formula body, pure `POLYPOLYGON` WMF preview, and Word display box are handled as separate layers.
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
| Pure-vector WMF preview | What Word paints before the OLE object is opened; bitmap and text records are forbidden |
| Word display box | Visible width, height, and baseline on the page |

That boundary is the main reason layout tuning happens in the Word object shell instead of by forcing point-size records into the MTEF body.

## Validation

Main reference round trip:

```powershell
.\scripts\verify-reference-roundtrip.ps1
```

This checks the regenerated reference DOCX through OLE inspection, Word/MathType probes where available, `docx2tex` coverage, and formula display-box comparison.

Full xsc corpus acceptance is documented in [docs/xsc-latex-assets.md](docs/xsc-latex-assets.md). The latest recorded summary covered `1..155` source files, `39551` paired formula objects, and `0` generated non-WMF previews.

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
