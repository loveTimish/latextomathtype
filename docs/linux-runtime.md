# Linux Runtime

`latextomathtype` can run on Linux in pure-Java export mode. Set:

```bash
java \
  -Dmathtype.windows.enabled=false \
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

The script starts the jar with `mathtype.windows.enabled=false`, verifies that the cache directory is writable, checks `/api/export/health`, and stops the process.

If Docker cannot pull `eclipse-temurin`, build a local smoke image from the already available Debian-based `postgres:16` image, then run the same smoke test:

```powershell
docker build -f Dockerfile.smoke -t latextomathtype:smoke .
powershell -ExecutionPolicy Bypass -File .\scripts\linux-smoke-docker.ps1
```

This path installs OpenJDK 21 and `curl` once in a local helper image, then runs `scripts/linux-smoke.sh` against the packaged jar.

## Formula Rendering

The renderer first tries a native TeX pipeline:

```text
latex -> dvisvgm -> SVG -> PNG
```

Install these packages on Debian/Ubuntu-style systems:

```bash
apt-get update
apt-get install -y openjdk-21-jre-headless texlive-latex-base texlive-latex-extra texlive-fonts-recommended dvisvgm
```

If `latex` or `dvisvgm` is not installed, the service falls back to JLaTeXMath. That fallback still runs on Linux, but previews can differ slightly from native TeX and MathType.

Override command paths when needed:

```bash
java \
  -Dpaperword.latex.command=/usr/bin/latex \
  -Dpaperword.dvisvgm.command=/usr/bin/dvisvgm \
  -Dpaperword.latex.timeout.seconds=20 \
  -jar target/paper-to-word-1.0.0.jar
```

## MathType Boundary

`mathtype.windows.enabled=false` uses this project's pure-Java MTEF/OLE writer and is the Linux-compatible mode.

`mathtype.windows.enabled=true` is only for routing conversion to an external Windows MathType service. Linux cannot run desktop MathType locally.

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
