# XSC MathType Linux WMF Goal

## Ultimate Goal

Build the `docxtolatex -> latextomathtype` rebuild path so the XSC test set under
`F:\资料\xsc资料\word_files` can be regenerated on Linux with results close to
official MathType output.

Target acceptance:

- Full XSC corpus can be rebuilt from source DOCX through `docxtolatex`, LaTeX,
  `latextomathtype`, and final DOCX generation.
- WMF preview physical size should match the source/test-set MathType object as
  tightly as possible, with the long-term gate set to relative error <= 1%.
- WMF preview should be vector WMF records (`ExtTextOut`, `Polyline`, etc.) and
  avoid bitmap `StretchDIB` fallback for supported high-frequency formulas.
- OLE/MTEF should remain editable and should match the source MathType MTEF
  structure as closely as practical, ignoring unavoidable header/container
  differences.
- Acceptance must be repeatable by scripts and reports, not by visual inspection
  alone.

## Current Branch

- Repo: `D:\latextomathtype\latextomathtype-wmf`
- Branch: `xsc-mathtype-linux-wmf-calibration`
- Remote: `origin -> https://github.com/loveTimish/latextomathtype.git`

## Work Completed In This Pass

This pass moved the vector WMF renderer further away from TeX/DIB fallback for
XSC-style arithmetic layouts.

Code changes:

- `VectorWmfFormulaRenderer` now parses top-level `array` blocks with balanced
  `\begin{array}` / `\end{array}` matching instead of regex-only extraction.
- Array rows and cells are split only at top level, so nested arrays no longer
  break the parser.
- Array cells can recursively lay out nested arrays, simple fractions, simple
  scripts, standalone scripts, and flat text.
- `\,` row-break residue before a nested array is normalized when it appears as
  `,\begin{array}` inside an array cell.
- Additional flat command mappings were added for escaped vertical bars and
  `\to`.
- Wide puzzle/vertical arithmetic arrays use tighter column sizing and centered
  cell placement to reduce over-wide generated WMF previews.
- Render cache version was bumped to force regeneration with the new vector
  behavior.

Regression tests added:

- Wide 8-column arithmetic puzzle array from the XSC set must render as vector
  WMF text/lines without `StretchDIB`.
- Nested array-cell layout from the XSC set must render as vector WMF text/lines
  without `StretchDIB`.

## Verification Performed

Completed:

```powershell
mvn -q "-Dtest=VectorWmfFormulaRendererTest,LaTeXImageRendererTest" test
```

Result: passed.

Partial/aborted verification:

- A 41-50 full acceptance run was attempted, but the user requested stopping
  before it completed.
- A generated single document exists at:
  `D:\latextomathtype\analysis\batch10-full-docx\20260612-171548\xsc测试集完整重建_41.docx`
- The full 41-50 acceptance summary was not regenerated after this patch.

Important earlier baseline before this pass:

- Previous 41-50 report:
  `D:\latextomathtype\analysis\template-rebuild\20260612-165141\template-acceptance-summary.json`
- That report still showed 3 `StretchDIB` fallbacks and a worst WMF width error
  of about 13.08%, concentrated in doc 41/doc 50 structures. The current code
  directly targets those classes, but the full acceptance rerun is still needed
  to prove corpus-level improvement.

## Pause Point

Do not treat the overall goal as complete yet.

The current safe stopping point is:

1. Code compiles for the renderer test slice.
2. Targeted vector-renderer tests pass.
3. Full 41-50 and full corpus acceptance are still pending.
4. The objective remains active conceptually, but work is paused at the user's
   request after pushing this checkpoint.

## Next Steps

When resuming, start with verification before more tuning.

1. Kill any stale Maven/Surefire Java process if a previous run was interrupted.

```powershell
Get-CimInstance Win32_Process |
  Where-Object { $_.Name -eq 'java.exe' -and $_.CommandLine -match 'surefire|XscFullBatch10DocxTest|maven|latextomathtype' } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
```

2. Run the renderer tests.

```powershell
cd D:\latextomathtype\latextomathtype-wmf
mvn -q "-Dtest=VectorWmfFormulaRendererTest,LaTeXImageRendererTest" test
```

3. Regenerate and summarize the 41-50 batch. Prefer running from the current
   PowerShell session rather than nesting `powershell -File`, because nested
   process invocation previously caused Chinese filename/path mojibake.

```powershell
& 'D:\latextomathtype\analysis\run_xsc_acceptance.ps1' `
  -Start 41 `
  -End 50 `
  -DatasetDir 'F:\资料\xsc资料\word_files' `
  -AnalysisDir 'D:\latextomathtype\analysis' `
  -LatexRoot 'D:\latextomathtype\analysis\xsc-latex' `
  -DocxToLatexDir 'D:\docxtolatex\docxtolatex'
```

4. Check the summary gates:

- `stretchDib` should drop to 0 for docs 41-50.
- `nonWmfGenerated` should remain 0.
- WMF width/height max error should be <= 1% for the target batch. If the 13%
  width case remains, inspect the corresponding metrics CSV row and tune only
  that layout class.
- MTEF reports should be checked for structure regressions, especially nested
  arrays and wide vertical arithmetic puzzles.

5. After 41-50 passes, expand to the full XSC corpus in batches and publish a
   single repeatable report containing:

- `size-report.json`
- `failures.json`
- worst-case CSV rows
- WMF record classification counts
- MTEF comparison summary

## Known Risks

- The current vector layout is still heuristic for MathType spacing. It is a
  better direction than bitmap fallback, but it is not yet a complete MathType
  layout engine.
- Array compression may need category-specific calibration. It improves the
  observed wide puzzle class but must be validated across the corpus to avoid
  over-compressing legitimate arrays.
- Nested arrays are now parsed structurally, but deeper structures such as
  radicals, complex nested fractions, and advanced accents still need explicit
  vector layout support.
