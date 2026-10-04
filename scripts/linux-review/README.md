# Linux review benchmark

All inputs are synthetic. No production exam data, credentials, external image URLs,
Windows automation, deployment or remote Git writes are used.

## Reproduction

Requirements: Java 21, Python 3, Node **24.9.0**, locked npm dependencies, Maven.
Run from a clean checkout of `bf56c59bc823a8a1bab518b4bc3a3d13d1a113d5`.
Use your environment's supported Maven proxy/settings if necessary.

The default Maven suite conditionally skips reference-document generators whose
external fixtures are absent. In particular,
`GeneratedMathTypeTexReferenceDocxTest` requires
`analysis/wmf-structure-metrics/word-mathtype-tex-reference-expanded.source-report.json`.
When that file is present, its parsing, formula-count and rendering assertions all
run normally; a malformed fixture or rendering failure is not skipped. A skip is
not an assertion that Word/MathType native reference compatibility has passed.

1. Before applying the patch, run `npm ci` and `sh .mvn/apache-maven-3.9.12/bin/mvn -DskipTests test-compile dependency:build-classpath -Dmdep.outputFile=target/dependencies.txt`.
2. Create `target/linux-review`, copy `target/classes` to `target/linux-review/baseline-classes`, and create `target/linux-review/classpath.txt` containing the **absolute** `target/test-classes`, absolute `target/classes`, and the contents of `target/dependencies.txt`, separated by `:`.
3. Apply the patch on the clean base. The single pinned HTML snapshot has an explicit CRLF checkout rule; after applying a patch to an existing working tree, run `git checkout-index --force -- docs/reference/mathtype/raw/latex-coverage.html` to materialize that rule. This rewrites only the unmodified reference snapshot from the index, so do not use it over local edits.
4. Ensure `node_modules/node/bin/node` is Node 24.9.0 (the local test environment used `npm install --no-save --package-lock=false node@24.9.0`). Verify package-lock.json is unchanged. Alternatively adjust the node path in these benchmark scripts to your verified Node 24.9.0 binary.
5. Run the selected Java suite: `sh .mvn/apache-maven-3.9.12/bin/mvn -Dpaperword.mathjax.node.command="$PWD/node_modules/node/bin/node" -Dtest="$(find src/test/java -name '*Test.java' ! -path '*/tools/*' -printf '%f\n' | sed 's/\.java$//' | paste -sd, -)" test`. Optional corpus/Office prerequisites remain skipped. On a JDK that disallows Mockito dynamic self-attach, pass its existing Byte Buddy agent at JVM startup with `-DargLine=-javaagent:/path/to/byte-buddy-agent.jar`; do not change OS security settings.
6. Compile the harness: `java com.sun.tools.javac.Main -proc:none -cp "$(cat target/linux-review/classpath.txt)" -d target/test-classes scripts/linux-review/LinuxBenchmark.java`.
7. Run sequentially: `python scripts/linux-review/run_benchmarks.py --mode performance --forks 3`, then modes `paper40 --forks 3`, `paper100 --forks 1`, and `distinct --forks 5`. Do not run other CPU-heavy tests at the same time.
8. `python scripts/linux-review/run_http_smoke.py` checks the current optimized classes: cold JVM startup, two successive HTTP exports, two OLE/WMF formulas, and one unchanged synthetic PNG per output. Use `--forks 1` for a short smoke. A historical comparison must explicitly pass `--variant baseline --variant optimized --baseline-script /path/to/original-checkout/tools/mathjax/render_mathjax_svg.cjs`; the original worker, package lock and dependencies must match those baseline classes. The script rejects a baseline without its original worker rather than silently mixing different renderer bundles.

The original performance report used an unchanged MathJax bundle on both sides.
For later changes to that bundle, use separate baseline and current checkouts with
their matching workers; copying Java classes alone is not a valid comparison.

The harness also accepts `concurrency`, `bugs`, `timeout`, and `compare` modes.
Always put an external time limit around `timeout` on the original code:
its deliberately nonresponsive worker exposes the deadlock being repaired.
Example: `timeout -k 2 8 java ... com.lz.paperword.core.render.LinuxBenchmark timeout target/linux-review/timeout-probe`.

`run_benchmarks.py` samples JVM and descendant RSS every 25 ms. RSS is observed,
not an exact instantaneous maximum; JVM heap-used snapshots are not leak tests.
The `paper40` and `paper100` modes start with an empty in-memory render cache.
The `performance` mode's later DOCX cases reuse formulas rendered earlier in that JVM.
Cache-disabled in these commands means **disk cache disabled**; the existing JVM cache remains enabled.

The reference commit and optimized code use the same Java flags, Node binary,
MathJax bundle, input formulas and cache settings. No renderer or correctness gate
is disabled for speed.
