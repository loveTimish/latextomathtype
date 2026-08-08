# Linux Runtime

`latextomathtype` uses the pure-Java MTEF/OLE export path. Start it with:

```bash
java \
  -Dpaperword.render.cache.enabled=true \
  -Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render \
  -jar target/paper-to-word-1.0.0.jar
```

Build the executable jar with:

```bash
./.mvn/apache-maven-3.9.12/bin/mvn -DskipTests package
```

Or build a Linux container image from the project root after packaging:

```bash
docker build -t latextomathtype:local .
docker run --rm -p 8081:8081 \
  -v latextomathtype-cache:/var/cache/latextomathtype/formula-render \
  latextomathtype:local
```

Check the service:

```bash
curl http://127.0.0.1:8081/api/export/health
```

For a direct Linux host smoke test after packaging:

```bash
sh scripts/linux-smoke.sh
```

The script starts the jar, verifies that the cache directory is writable, checks `/api/export/health`, and stops the process.

If Docker cannot pull `eclipse-temurin`, build a local smoke image from the already available Debian-based `postgres:16` image, then run the same smoke test:

```powershell
docker build -f Dockerfile.smoke -t latextomathtype:smoke .
powershell -ExecutionPolicy Bypass -File .\scripts\linux-smoke-docker.ps1
```

This path installs OpenJDK 21 and `curl` once in a local helper image, then runs `scripts/linux-smoke.sh` against the packaged jar.

## Formula Rendering

The OLE preview renderer uses one strict path:

```text
LaTeX -> locked MathJax SVG -> Batik vector scene -> WMF POLYPOLYGON
```

Build the offline Linux sidecar on a release machine:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-vector-sidecar.ps1 -Platform linux-x64
```

Extract `target/vector-sidecar-dist/vector-sidecar-linux-x64.tar.gz` beside the application JAR.
The extracted directory must be named `vector-sidecar`; it contains fixed Node `24.9.0`, MathJax
`3.2.2`, the worker, locked `node_modules`, and the required package metadata. No npm command or
network access occurs at runtime.

The worker reports its engine, Node version, MathJax version, and bundle hash on every response.
The Java side rejects any mismatch. Missing tools, unsupported SVG operations, missing glyphs,
coordinate overflow, bitmap records, or text records fail the export with the original LaTeX.
There is no PNG/DIB fallback.

Command overrides are development-only:

```bash
java \
  -Dpaperword.mathjax.node.command=/opt/dev-node/bin/node \
  -Dpaperword.mathjax.script=/workspace/tools/mathjax/render_mathjax_svg.cjs \
  -jar target/paper-to-word-1.0.0.jar
```

## MathType Boundary

The service does not route through a Windows MathType bridge. It always uses this project's pure-Java MTEF/OLE writer.

## Cache

Formula rendering uses two cache layers:

```text
memory cache: per JVM process
disk cache: persistent across restarts
```

Disk cache is enabled by default and keyed by normalized LaTeX, render mode, size, and renderer cache version. Use a persistent writable directory in production:

```bash
-Dpaperword.render.cache.enabled=true
-Dpaperword.render.cache.dir=/var/cache/latextomathtype/formula-render
```

To disable persistent cache:

```bash
-Dpaperword.render.cache.enabled=false
```
