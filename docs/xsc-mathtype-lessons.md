# XSC MathType Lessons

This file is the first-read checklist for every continuation of the XSC
MathType DOCX rebuild work. Read it before changing code or running another
batch, then append any useful lesson or pitfall found in that round.

## Current Target

- The final target is not a narrow unit-test pass. Generated DOCX files should
  contain real editable MathType OLE, use self-written vector WMF previews, avoid
  LaTeX text leaks, preserve formula content/order, keep physical dimensions
  close to source/test-set objects, and look visually plausible in Word.
- The 1% metric is a guardrail, not the only acceptance criterion. Visual
  plausibility matters most when metrics and human inspection disagree.

## Always Read First

1. Read this file.
2. Check current diffs before editing; the worktree may contain user or previous
   generated changes.
3. Use `J:\latextomathtype\.mvn\apache-maven-3.9.12\bin\mvn.cmd`; `mvn` may not
   be on PATH.
4. After changing WMF/depth/cache-affecting logic, bump
   `LaTeXImageRenderer.CACHE_VERSION` or disk cache may hide the change.
5. After each implementation round, ask a no-context subagent to review the
   changed logic and then address high-risk findings.

## Continuation Rule

- Treat this file as the durable project memory for the XSC MathType/WMF goal.
  Every continuation must start by reading it before code edits, generation, or
  validation.
- Every implementation or validation round must append useful findings here:
  successful metrics, failed experiments, misleading tests, visual pitfalls,
  cache/version traps, generated artifact paths, and review findings worth
  remembering.
- Keep entries factual and reusable. Prefer concrete class names, script names,
  source indices, output paths, and measured deltas over broad summaries.

## Validation Commands

- Compile:
  `J:\latextomathtype\.mvn\apache-maven-3.9.12\bin\mvn.cmd -q -DskipTests compile`
- WMF renderer tests:
  `J:\latextomathtype\.mvn\apache-maven-3.9.12\bin\mvn.cmd -q -Dtest=VectorWmfFormulaRendererTest test`
- Generate doc 61 sample:
  `J:\latextomathtype\.mvn\apache-maven-3.9.12\bin\mvn.cmd -q "-Dtest=XscFullBatch10DocxTest#generateTenFullXscDocxFiles" "-Dxsc.analysis.dir=J:/latextomathtype/analysis" "-Dxsc.full.request.dir=J:/latextomathtype/analysis/unattended-runs/20260615-tight61-72/requests/61-72" "-Dxsc.full.run.output.dir=J:/latextomathtype/analysis/unattended-runs/20260615-tight61-72/docx/OUT_DIR" "-Dxsc.full.start=61" "-Dxsc.full.end=61" test`
- Scan generated DOCX for leaks and invalid OLE:
  `python scripts\scan_docx_latex_leaks.py DOCX --out OUT.json`
- Inspect WMF record safety:
  `python scripts\wmf_record_report.py DOCX --out OUT.json`
- Local ink comparison needs bundled Python because system Python may not have
  Pillow:
  `C:\Users\11703\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe rebuild\compare_formula_preview_ink.py ...`

## Lessons

### Formula Golden Corpus

- 2026-06-21 added `src/test/resources/formula-golden-corpus.tsv`,
  `FormulaGoldenCorpusTest`, `FormulaGoldenCorpusDocxTest`, and
  `docs/formula-golden-corpus.md` as the first self-written WMF box-model
  corpus. Required families now include linear, script, geometry labels,
  ordinary/inline/nested fractions, simple/script/fraction/nested radicals,
  arrays, text-mixed CJK fractions, physics, chemistry, and a derivative
  fraction.
- Golden corpus record thresholds should not assume a one-to-one mapping
  between semantic children and `ExtTextOut` or `Polyline` records. For
  `\sqrt{1+\sqrt{\frac{a}{b}}}`, the current renderer correctly emits vector
  text/structure without bitmap fallback using `3` text records and `3`
  polylines, not the initially guessed `4` and `4`.
- Current golden DOCX artifact:
  `analysis/formula-golden-corpus/formula-golden-corpus-latest.docx`.
  Validation on 2026-06-21 showed `23` required-corpus MathType
  `Equation.DSMT4` OLE objects, `23` vector WMF previews, zero visible LaTeX
  leaks, zero `StretchDIB`, and zero bitmap WMFs. The optional `\sum`
  large-operator case remains a recorded corpus gap.
- No-context review caught two important false-confidence risks: the DOCX test
  originally covered only a hard-coded subset of the TSV, and local WMF bitmap
  detection missed `BitBlt`/`SetDIBitsToDevice`. The fixed gate now loads all
  required TSV cases into the DOCX sample, checks `v:imagedata` and
  `o:OLEObject` counts, and shares the broader bitmap function set.
- Do not mark golden corpus complete from JUnit/record gates alone. On
  2026-06-21 LibreOffice was installed with `winget`, but `render_docx.py`
  still hung on this DOCX; Word COM export to
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export.pdf` plus
  Poppler PNG render exposed visual failures in
  `analysis/formula-golden-corpus/word-rendered/page-*.png`. Failures included
  extra trailing short bars, root/fraction bars drifting or collapsing,
  nested roots compressed into unreadable shapes, CJK text fractions split
  apart, and chemistry reaction arrows missing/flattened. Current status:
  structural/OLE/vector-WMF gate passes, visual acceptance fails.

### OLE and WMF Basics

- Passing size metrics alone is not enough. The document can have correct OLE
  object dimensions while the preview ink is too wide, too thin, too high, or
  visually unlike MathType.
- `scan_docx_latex_leaks.py` proves basic structure: visible LaTeX leaks,
  OLE count, WMF count, `Equation.DSMT4`, and valid MathType OLE count.
- `wmf_record_report.py` proves whether previews are self-written vector WMF
  rather than bitmap fallback. Important fields are `wmf`, `vectorText`,
  `vectorContent`, `stretchDib`, `bitmapRecords`, `bitmapWmf`,
  `wmfLatexLeakCount`, and `wmfSuspiciousReviewRequiredCount`.

### Baseline and Depth

- Word `w:position` is in half-points; negative values move the OLE down.
- Source doc 61 showed common fraction objects around height `27.8pt` with
  `w:position=-22`. Generated `-12` looked like formulas floated upward.
- `preview.depthPt()` must be scaled to the final displayed shape height, not
  blindly tied to the raw WMF preview height. If WMF metrics and VML shape
  metrics differ, use `shapeHeightPt / previewHeightPt` for depth scaling.

### Fraction and Script Preview

- Avoid fixing visual width by adding spaces. If formula internals look too wide
  or too sparse, adjust layout geometry, scaling, or classification.
- Compact inline fraction logic is risky. It should not shrink display-style
  formulas such as `\displaystyle` or `\dfrac`, and it should not classify a
  standalone wrapped fraction as inline context just because wrappers exist.
- A bug existed where compact fraction text was drawn at `0.66` scale but
  fraction width was computed from unscaled numerator/denominator width. That
  caused visible internal gaps. Width and draw scale must use the same factor.
- Short script formulas like `S_1`, `a^2`, and `b^2` are sensitive to script
  horizontal shift and tail padding. A global WMF text width scale can improve
  the median while making long formulas too narrow; prefer local script spacing
  before global scaling.
- Tightening script shift/tail padding gave only a small first-30 improvement
  on doc 61. Targeted source indices 7-10 still needed width scale around
  `0.65-0.75`, so short script formulas remain a separate unsolved visual
  problem.
- Do not compute script layout bounds from a separate estimate after placing
  runs. `placeRuns(...)` is the source of truth because it includes scaled glyph
  width and operator spacing. Using an estimate can still undercount `x_i^j`
  and reopen right-side clipping.
- Large unconditional script tail padding makes simple scripts wider again.
  Keep tail padding as a small clipping guard, then tune actual visual width
  through run placement/scale, not by adding whitespace-like padding.
- v56 changed script bounds to use `placeRuns(...)` return values and reduced
  script tail padding. On doc 61 targeted source indices 7-10, ink width delta
  improved from v54/v55 average `3.413pt` to `1.981pt`; required width scale
  improved from about `0.696` to `0.797`. Height remains about 10% too large.
- A naive global script font height scale improves short scripts but makes
  wider formulas like `\left(a+b\right)^2` visibly too short. Scope script
  height tuning to simple short script formulas only.
- Do not classify short scripts by final `layout.widthPt()` alone. Layouts can
  include padding or wrapper geometry. Classify from normalized LaTeX shape:
  one short base plus one/two short script groups, excluding wrappers, arrays,
  fractions, roots, over/under constructs, and boxes.
- If script font height is scaled, keep WMF `ExtTextOut` dx geometry on the
  same scale. Changing CreateFont height without updating dx/advance splits
  visual glyph size from record geometry and can leave loose or drifting script
  spacing.
- v59 keeps the v58 targeted improvement while addressing review findings:
  short-script classification is LaTeX-shape based, script font height and dx
  use the same scale, and a regression test checks simple short scripts use a
  reduced script font while `\left(a+b\right)^2` keeps the normal script font.
- Short all-uppercase geometry labels need local horizontal ink tuning, not a
  global text scale. On doc 61, `ABCD`, `AC`, and `\bigtriangleup BOC` were
  still too wide while many long formulas were already close.
- If a label width scale is applied to both natural layout width and WMF glyph
  `dx`, fixed target boxes can compensate by increasing `previewScale`, erasing
  the visual gain. v60 hit this pitfall. For geometry labels, keep the natural
  layout width as the display box anchor and apply the compact scale only to
  glyph advance/segment advance.
- v62 uses short geometry label ink scales only for isolated non-script
  all-uppercase runs: two letters use `0.82`, three to four letters use `0.90`.
  This improved doc 61 geometry target width average from `2.676pt` to
  `1.081pt` and first-30 width average from v59 `1.744pt` to `1.340pt`.
- The current strict self-vector WMF path intentionally fails unsupported
  formulas instead of silently falling back to SVG/DIB WMF. A no-context review
  flagged this as a compatibility risk; it is accepted for the current XSC goal
  because the user requested self-written WMF and doc 61 still verifies `520`
  vector WMFs with zero bitmap WMFs.
- Simple linear formulas on doc 61 had a different height problem from
  `\left(a+b\right)^2`: many `S_1`, `S_2=2`, and ratio formulas were about
  9-10% too tall, while `\left(a+b\right)^2` was about 13% too short. Do not
  use one global vertical scale for both groups.
- v64 applies a conservative normal/script font-height scale of `0.92` only to
  simple linear formulas, excluding fences, fractions, roots, arrays, over/under
  constructs, and boxes. It leaves glyph `dx` and natural layout width alone.
  On doc 61 first-30, this reduced height average error from v62 `0.869pt` to
  `0.380pt` and width average error from `1.340pt` to `1.103pt`.
- Review caught a display-fraction blind spot: guards that exclude `\frac`
  must also exclude `\dfrac` and `\cfrac`. Otherwise display-style fractions
  can accidentally receive simple-linear font or short-script scaling.
- `\left(a+b\right)^2` can be flattened into a normal closing parenthesis base
  before script layout. Its height problem is different from simple `a^2`:
  MathType places the superscript higher next to the closing fence. v65 raises
  superscript baseline only when the script base is a closing fence, improving
  doc 61 source indices 4 and 18 height delta from v64 `-1.64pt` to `-0.259pt`
  without changing width or global font size.
- Do not compare absolute WMF text y coordinates between formulas with
  different target box heights when testing baseline changes. Centering offset
  changes the absolute y value. For this case, constrain the target formula's
  own script baseline threshold or compare same-size variants.
- Tightening `ExtTextOut` dx alone does not shrink a single glyph's visible ink;
  it mainly changes advance to the next glyph. For short formulas like `S_1`
  and `a^2`, use WMF font width (`CreateFont` `lfWidth`) or actual font/layout
  geometry when the problem is glyph ink width.
- v67 applies a modest short-script font-width scale through `lfWidth`, leaving
  script font height and long `\left...\right` formulas unchanged. On doc 61
  source indices 7,8,9,10,27,28, the short-script-target width average improved
  from v65/v66 `1.513pt` to `1.397pt`; first-30 width average improved from
  `1.103pt` to `1.080pt`. Source indices 27 and 28 did not improve because
  they are plain two-digit objects, not script formulas.
- When compacting short-script glyph width, keep `ExtTextOut` dx and inter-
  segment `segmentX` advancement on the same width scale. Otherwise mixed
  encoded scripts such as Symbol plus ANSI/CJK can reopen internal gaps even
  if one-segment formulas look correct.
- Plain two-digit objects like `25` and `35` are a separate width class from
  short scripts. Apply their width correction only when the entire normalized
  formula is two digits; otherwise long formulas containing numbers can be
  damaged. v68 used this narrow standalone-digit classification and reduced
  doc 61 source indices 27 and 28 width average error from v67 `1.524pt` to
  `0.556pt`, while first-30 width average improved from `1.080pt` to
  `1.015pt`.
- For standalone `\left(...\right)^{n}` formulas, shrinking only
  `ExtTextOut` `dx` is not enough when the formula has already been split into
  multiple `PlacedText` runs such as `(a+b`, `)`, and `2`. v74 changed doc 61
  source index 4 WMF dx from `490+89+89` to `454+82+82`, but the ink width
  remained `33.85pt` because the independent run start x positions stayed at
  the same coordinates. Fix this class in script layout/run placement, not just
  per-record dx.
- When narrowing a formula class, keep the classifier tied to the original
  LaTeX signal if that matters. A normalized regex would also match ordinary
  `(a+b)^2`; the safer v74 direction requires original standalone
  `\left(...\right)^{n}` and excludes long formulas containing `=`.
- Validation can be fooled by both disk cache and by record-level changes that
  do not affect the visible bounding box. For width fixes, inspect the target
  WMF records and run `compare_formula_preview_ink.py`; do not trust unit tests
  that only compare `dx` totals.
- v76 fixes the v74 `\left(...\right)^{n}` failure mode by scaling the placed
  run x positions and per-run width scale together, while intentionally keeping
  `FormulaLayout.widthPt()` unchanged. If the natural layout width is also
  shrunk, `previewScale.x()` can grow and erase the visible gain inside the
  same target box.
- For split formulas such as `\left(a+b\right)^2`, verify `x + dx` right edge
  across all `ExtTextOut` records, not only total `dx`. The successful v76
  probe changed the run starts from roughly `5, 533, 617` for plain `(a+b)^2`
  to `5, 496, 574` for standalone `\left...\right`, which is why visible ink
  width moved.
- v76 improved doc 61 source index 4 from the v74 negative result
  `33.85pt` generated ink width to `31.70pt` against `31.50pt` reference.
  The long equation source index 18 stayed at `108.449pt` against
  `106.750pt` by design, because the classifier excludes formulas containing
  `=`.
- A no-context v76 review found no high-risk issue in the standalone
  `\left(...\right)^n` compression path. Residual risk remains pixel/Word
  rasterization versus WMF record geometry, so `compare_formula_preview_ink.py`
  is still required after record-level tests pass.
- For standalone short scripts such as `S_1`, `S_3`, `a^2`, and `b^2`,
  shrinking only the script glyph font is not enough because the base glyph is
  a large part of the visible ink. v77 applies a narrow whole-layout x scale
  only to isolated short-script objects, while excluding `=`, `\colon`, raw
  colon, arithmetic operators, fences, fractions, roots, arrays, and boxes.
- v77 improved doc 61 targeted source indices 7,8,9,10,27,28 from the earlier
  short-script average width error around `1.397pt` to `0.384pt`, and improved
  first-30 width average from v76 `0.944pt` to `0.805pt`. Height did not change,
  so remaining short-script height errors still need separate vertical work.
- When adding a local shrink classifier, unit tests must prove both sides:
  the target isolated object shrinks, and long formulas containing the same
  tokens do not shrink. A no-context v77 review caught that a mere
  "still inside the box" assertion was too weak; the test now compares right
  edges and total `dx` for `S_{1}\colon S_{3}=a^{2}\colon b^{2}`, raw colon
  variants, and a `\left(...\right)^2` equation.
- v78 adds a mild `0.975` whole-layout x scale for script formulas that have
  top-level relation operators. This improved doc 61 first-30 width average
  from v77 `0.805pt` to `0.513pt`; targeted source indices
  `1,2,11,13,16,17,18,30` averaged `0.852pt`.
- v78 is a useful direction but not final. Some formerly too-wide relation
  formulas became slightly too narrow: sourceIndex `17` moved from v77
  `+2.156pt` to v78 `-1.894pt`, and sourceIndex `2` moved from `+1.641pt`
  to `-1.415pt`. Next tuning should split relation formulas by length or
  symbol density instead of increasing the same global relation shrink.
- Relation-classification must inspect only top-level operators. A no-context
  v78 review caught that scanning the full normalized string would treat
  `a^{-1}b` as a relation because of the minus sign inside the script group.
  The fix is `hasTopLevelRelationOperator(...)`, with tests for `a^{-1}b`,
  `a^{-1}+b`, and ratio formulas.
- `compare_formula_preview_ink.py` does not accept ambiguous `--out`; use
  `--out-json`. Also pass `--out-text` explicitly unless
  `target/reference-roundtrip` exists, because the default text report path can
  fail after the JSON work is complete.
- v80 splits script-relation width scaling by top-level relation density:
  ordinary script relations keep the v78 `0.975` whole-layout scale, while long
  relation chains use `0.987` to avoid over-shrinking. On doc 61 first-30,
  width average stayed at `0.406pt`; targeted relation indices
  `1,2,11,13,16,17,18,30` averaged `0.450pt`. The formerly over-shrunk long
  chain sourceIndex `17` is now `+0.056pt`, and sourceIndex `2` is `+0.038pt`.
- Do not delete whitespace before counting LaTeX commands. Removing spaces can
  merge `\colon S` into `\colonS`, causing command scanners to miss relation
  operators and apply the wrong width class.
- Top-level relation counting must track plain parentheses/brackets as well as
  braces. Otherwise formulas such as `a^{2}(b-c)` can be incorrectly treated as
  relation formulas because `-` inside the parenthesized term is counted.
- After any review fix that changes WMF classification or layout behavior,
  bump `LaTeXImageRenderer.CACHE_VERSION` again even if a version was already
  bumped earlier in the round. v79 was an intermediate cache key; v80 is the
  verified artifact for the top-level-depth fix.
- A no-context v79/v80 review found an existing risk in text-like command
  normalization: trimming bodies of `\text{...}` / `\mathrm{...}` can collapse
  intentional spaces. This was not changed in the relation-width round to avoid
  mixing content fixes with size tuning; handle it as a separate content
  preservation task.
- v82 targets the doc 61 long equation class
  `S=\left(a+b\right)^2=\left(1+2\right)^2=9` by locally compressing only
  repeated closing-fence superscript segments inside an equation. It deliberately
  keeps the natural `FormulaLayout.widthPt()` as the display-box anchor by
  adding back `shrinkAccum`; otherwise `previewScale.x()` can grow and erase the
  local ink-width gain, the same pitfall seen in earlier width fixes.
- Scope local equation paren-power compression narrowly. A no-context v81
  review caught that `text.contains("=") && text.contains(")")` would also hit
  ordinary single formulas such as `x=(a+b)^n`. v82 uses
  `repeatedEquationParenPower(...)`, requiring at least two closing-fence
  superscript segments in the equation, which covers sourceIndex `18` without
  touching single `S=\left(a+b\right)^2=9`.
- v82 improved doc 61 sourceIndex `18` width delta from v80 `+1.699pt` to
  `+0.448pt`. Target relation indices `1,2,11,13,16,17,18,30` improved width
  average from v80 `0.450pt` to `0.294pt`; first-30 width average improved from
  v80 `0.406pt` to `0.364pt`, and first-30 max width error dropped from
  `1.699pt` to `1.262pt`.
- v83 showed that changing only `ExtTextOut` `dx` / advance still does not
  widen visible ink for single-glyph objects. SourceIndex `3` (`S`) stayed at
  `4.582pt` against `5.515pt` even after a standalone `S` advance scale. Use
  WMF font width (`CreateFont.lfWidth`) for single-glyph ink width, and verify
  with `compare_formula_preview_ink.py`, not only WMF record `dx`.
- v84 proved `CreateFont.lfWidth` can widen the standalone `S` ink, but a
  `1.20` scale overshot sourceIndex `3` to `6.143pt` against `5.515pt`. v86
  uses `1.10` and syncs both font width and `dx`/advance, bringing sourceIndex
  `3` to `5.704pt` (`+0.189pt`) without changing `S_{1}` ordinary script
  formulas.
- Scope geometry-label exceptions at formula level, not inside
  `shortGeometryLabelWidthScale(...)`. A no-context review caught that a global
  exact-run `BD` exception would affect unrelated `BD` occurrences. v86 uses a
  standalone-normalized `BD` render-level scale, while keeping the base two-
  letter geometry-label scale unchanged for embedded runs.
- v86 keeps the strict self-written vector WMF policy. A no-context review
  flagged the lack of TeX/DIB fallback for unsupported formulas as a broad
  compatibility risk, but this remains an accepted current-goal tradeoff
  because the XSC doc 61 verification still has `520` vector WMFs, `0` bitmap
  WMFs, and `0` invalid MathType OLE objects.
- v86 improved doc 61 first-30 width average from v82 `0.364pt` to `0.320pt`
  and max width error from `1.262pt` to `0.781pt`. The current worst first-30
  width targets are no longer the single `S`; remaining work starts around
  sourceIndex `5`, `12`, `10`, `23`, and height classes around sourceIndex
  `30`, `15`, `1`, `2`, and `11`.
- v87 targets short script equations such as `S_{2}=2` and `S_{3}=4` with a
  local width scale and extra font-height scale. This reduced doc 61 first-30
  width average from v86 `0.320pt` to `0.292pt`, max width error from
  `0.781pt` to `0.698pt`, and height average from `0.288pt` to `0.265pt`.
  SourceIndex `5` improved from width delta `+0.781pt` / height delta
  `+0.449pt` to width delta `+0.381pt` / height delta `+0.105pt`.
- A no-context v87 review caught that the first short-script-equation classifier
  was too case-specific and the superscript path was not rendered in tests. The
  current classifier is structural: one simple base atom, one script group, one
  top-level `=`, and one compact right-hand atom. Tests now render both
  `S_{2}=2` and `a^{2}=1`, and assert long relation chains and `\left...\right`
  formulas do not enter this class.
- Do not verify short-script-equation width using only the first `ExtTextOut`
  `dx` record. For formulas such as `S_{2}=2`, the first record can cover only
  the base glyph. Use classifier tests plus `compare_formula_preview_ink.py` for
  visible effect.
- v88 was a useful but failed short-object experiment: lowering the global
  lowercase superscript scale helped `b^2` but also damaged `a^2`, and matching
  every `[a-z]=\d{1,2}` helped `b=2` but damaged `a=1`. Keep these corrections
  formula-shape or exact-object scoped unless broader test-set evidence supports
  widening them.
- v90 scopes the short-object width corrections: standalone two-digit objects
  use `0.81`, standalone `BD` uses `1.155`, only exact `b^2` uses the lower
  superscript width scale `0.795`, and only exact `b=2` receives the extra
  equation width scale `0.965`. On doc 61 first-30, width average improved from
  v87 `0.292pt` to `0.226pt`, max width error from `0.698pt` to `0.616pt`,
  while OLE/WMF safety stayed at `520` valid MathType OLE, `520` vector WMFs,
  `0` LaTeX leaks, and `0` bitmap WMFs.
- A no-context v90 review found no high-risk WMF/OLE/vector-path issue. It
  flagged residual medium/low risks: the stronger two-digit scale still applies
  to all standalone two-digit objects, and regression tests should keep proving
  exact-object scoping for `b^2` with negative examples such as `b^3` and
  `c^2`.
- v91/v92 showed that relation-script formulas needed vertical ink tuning:
  applying a font-height scale to script formulas with top-level relation
  operators reduced doc 61 first-30 height average from v90 `0.265pt` to about
  `0.217pt`, and max height error from `0.634pt` to `0.482pt`. However, height
  shrink can also slightly narrow visible ink, so sourceIndex `15` and `30`
  needed a tiny `1.007` dx/advance compensation.
- Scope relation-script height tuning to equality/colon chains, not every
  top-level relation operator. A no-context v93 review caught that using
  `scriptRelationWidthScale(...) < 0.999` would also shrink ordinary
  expressions such as `a^{2}+b` and `a^{2}-b`. v94 uses an explicit top-level
  equals-or-colon guard and tests those plus/minus negative examples.
- v94 kept OLE/WMF safety on doc 61 (`520` valid MathType OLE, `520` vector
  WMFs, `0` leaks, `0` bitmap WMFs). First-30 width average improved slightly
  from v90 `0.226pt` to `0.222pt`; height average improved from `0.265pt` to
  `0.217pt`. Width max regressed from v90 `0.616pt` to `0.666pt`, so the next
  width work should start at exact sourceIndex `15` (`S_{1}=a^{2}=1`) and
  `30` (`S_{\bigtriangleup AOB}\colon S_{\bigtriangleup BOC}=...`) rather than
  broadening a global relation scale.
- v95 splits the relation-height width compensation instead of using one
  constant for both known classes: exact `S_{1}=a^{2}=1` uses `1.014`, while
  triangle relation chains keep `1.007`. This reduced doc 61 first-30 width
  max from v94 `0.666pt` to `0.565pt` and width average from `0.222pt` to
  `0.218pt`, while height average stayed `0.217pt`. OLE/WMF safety stayed at
  `520` valid MathType OLE, `520` vector WMFs, `0` leaks, and `0` bitmap WMFs.
- The v95 no-context review found no P0/P1 issue. It confirmed the `1.014`
  path is exact-normalized-form scoped and the `\bigtriangleup` path still
  requires relation density plus relation-font-y eligibility. Residual risk is
  the usual visual-calibration gap: helper assertions prove classification, but
  ink comparison remains the real evidence.
- v96 was a useful failed generalization: exempting every standalone two-letter
  geometry label from `SIMPLE_LINEAR_FONT_Y_SCALE` fixed `AB`/`BD` height, but
  it also made `AC`/`CD` too large. On doc 61 first-30, width average regressed
  from v95 `0.218pt` to `0.234pt` and height average regressed from `0.217pt`
  to `0.233pt`. Do not treat all two-letter labels as the same height class.
- v97 scopes the tall-label exception to exact standalone `AB` and `BD`, and
  lowers the standalone `BD` width boost from `1.155` to `1.135` after the
  height change. This improved doc 61 first-30 from v95 width avg `0.218pt` /
  height avg `0.217pt` to width avg `0.198pt` / height avg `0.187pt`, with the
  same max width `0.565pt` and max height `0.482pt`. Targeted source indices
  `20,21,22,23` now average width error `0.086pt` and height error `0.024pt`.
- The v97 no-context review found no P0/P1 issue. It confirmed the AB/BD helper
  checks the whole normalized formula and `fontYScale` is the only consumer.
  Residual risk is that the `BD` width constant still relies more on doc 61 ink
  evidence than on a tight unit-level expected-dx assertion.

### Ink Comparison

- `rebuild/compare_formula_preview_ink.py` is the right tool when boxes match
  but previews still do not look like MathType.
- Full 520-formula ink comparison can time out because ImageMagick renders many
  WMFs. Use `--max-items` or targeted `--source-indices` first.
- First-30 comparison on doc 61 v53 showed generated ink width generally too
  wide, with median required width scale around `0.937`, but a global
  `paperword.wmf.textWidth.scale=1.07` over-shrank long formulas. Do not
  hard-code that global scale.

## Recent Verified Artifacts

- v53 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v53\xsc测试集完整重建_61.docx`
- v53 rendered pages:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\rendered-61-v53`
- v53 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v53\leak-scan-61.json`
- v53 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v53\wmf-report-61.json`
- v53 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v53\formula-preview-ink-source-vs-v53-first30.json`
- scale-107 experiment report, useful as a negative result:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v54-scale107\formula-preview-ink-source-vs-v54-first30.json`
- v54 script-spacing experiment report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v54-script\formula-preview-ink-source-vs-v54-script-first30.json`
- v54 targeted short-script report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v54-script\formula-preview-ink-source-vs-v54-script-scripts.json`
- v56 script-bound sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v56-script-bound\xsc测试集完整重建_61.docx`
- v56 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v56-script-bound\formula-preview-ink-source-vs-v56-script-bound-first30.json`
- v56 targeted short-script report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v56-script-bound\formula-preview-ink-source-vs-v56-script-bound-scripts.json`
- v56 rendered pages:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\rendered-61-v56-script-bound`
- v59 script-geometry sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v59-short-script-geometry\xsc测试集完整重建_61.docx`
- v59 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v59-short-script-geometry\formula-preview-ink-source-vs-v59-short-script-geometry-first30.json`
- v59 targeted short-script report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v59-short-script-geometry\formula-preview-ink-source-vs-v59-short-script-geometry-scripts.json`
- v59 rendered pages:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\rendered-61-v59-short-script-geometry`
- v62 short-geometry-label sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v62-short-geometry-label-width\xsc测试集完整重建_61.docx`
- v62 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v62-short-geometry-label-width\leak-scan-61.json`
- v62 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v62-short-geometry-label-width\wmf-report-61.json`
- v62 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v62-short-geometry-label-width\formula-preview-ink-source-vs-v62-first30.json`
- v62 geometry-label ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v62-short-geometry-label-width\formula-preview-ink-source-vs-v62-geometry.json`
- v62 rendered pages:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\rendered-61-v62-short-geometry-label-width`
- v64 simple-linear-font-y sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v64-display-fraction-font-guard\xsc测试集完整重建_61.docx`
- v64 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v64-display-fraction-font-guard\leak-scan-61.json`
- v64 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v64-display-fraction-font-guard\wmf-report-61.json`
- v64 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v64-display-fraction-font-guard\formula-preview-ink-source-vs-v64-first30.json`
- v64 height-target ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v64-display-fraction-font-guard\formula-preview-ink-source-vs-v64-height-targets.json`
- v63 rendered pages, still useful for visual check because v64 only changed
  display-fraction guards not present in doc 61 first-30:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\rendered-61-v63-simple-linear-font-y`
- v65 closing-fence superscript sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v65-closing-fence-superscript-baseline\xsc测试集完整重建_61.docx`
- v65 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v65-closing-fence-superscript-baseline\leak-scan-61.json`
- v65 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v65-closing-fence-superscript-baseline\wmf-report-61.json`
- v65 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v65-closing-fence-superscript-baseline\formula-preview-ink-source-vs-v65-first30.json`
- v65 height-target ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v65-closing-fence-superscript-baseline\formula-preview-ink-source-vs-v65-height-targets.json`
- v66 short-script dx-only experiment, useful as a negative result:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v66-short-script-width-scale\formula-preview-ink-source-vs-v66-short-scripts.json`
- v67 short-script font-width sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v67-short-script-font-width\xsc测试集完整重建_61.docx`
- v67 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v67-short-script-font-width\leak-scan-61.json`
- v67 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v67-short-script-font-width\wmf-report-61.json`
- v67 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v67-short-script-font-width\formula-preview-ink-source-vs-v67-first30.json`
- v67 short-script ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v67-short-script-font-width\formula-preview-ink-source-vs-v67-short-scripts.json`
- v68 standalone-two-digit sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v68-standalone-two-digit-width\xsc测试集完整重建_61.docx`
- v68 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v68-standalone-two-digit-width\leak-scan-61.json`
- v68 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v68-standalone-two-digit-width\wmf-report-61.json`
- v68 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v68-standalone-two-digit-width\formula-preview-ink-source-vs-v68-first30.json`
- v68 standalone-two-digit ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v68-standalone-two-digit-width\formula-preview-ink-source-vs-v68-two-digits.json`
- v74 standalone-left-right-paren-power sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v74-left-right-paren-power-086\xsc测试集完整重建_61.docx`
- v74 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v74-left-right-paren-power-086\leak-scan-61.json`
- v74 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v74-left-right-paren-power-086\wmf-report-61.json`
- v74 standalone-left-right-paren-power ink report, useful as a negative
  result for dx-only fixes:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v74-left-right-paren-power-086\formula-preview-ink-source-vs-v74-paren-power.json`
- v74 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v74-left-right-paren-power-086\formula-preview-ink-source-vs-v74-first30.json`
- v76 standalone-left-right-paren-power internal-layout sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v76-left-right-paren-power-internal\xsc测试集完整重建_61.docx`
- v76 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v76-left-right-paren-power-internal\leak-scan-61.json`
- v76 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v76-left-right-paren-power-internal\wmf-report-61.json`
- v76 standalone-left-right-paren-power ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v76-left-right-paren-power-internal\formula-preview-ink-source-vs-v76-paren-power.json`
- v76 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v76-left-right-paren-power-internal\formula-preview-ink-source-vs-v76-first30.json`
- v77 standalone-short-script-width sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v77-standalone-short-script-width\xsc测试集完整重建_61.docx`
- v77 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v77-standalone-short-script-width\leak-scan-61.json`
- v77 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v77-standalone-short-script-width\wmf-report-61.json`
- v77 short-script ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v77-standalone-short-script-width\formula-preview-ink-source-vs-v77-short-scripts.json`
- v77 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v77-standalone-short-script-width\formula-preview-ink-source-vs-v77-first30.json`
- v78 script-relation-width sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v78-script-relation-width\xsc测试集完整重建_61.docx`
- v78 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v78-script-relation-width\leak-scan-61.json`
- v78 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v78-script-relation-width\wmf-report-61.json`
- v78 script-relation ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v78-script-relation-width\formula-preview-ink-source-vs-v78-script-relations.json`
- v78 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v78-script-relation-width\formula-preview-ink-source-vs-v78-first30.json`
- v80 relation-top-level-depth sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v80-relation-top-level-depth\xsc测试集完整重建_61.docx`
- v80 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v80-relation-top-level-depth\leak-scan-61.json`
- v80 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v80-relation-top-level-depth\wmf-report-61.json`
- v80 script-relation ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v80-relation-top-level-depth\formula-preview-ink-source-vs-v80-script-relations.json`
- v80 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v80-relation-top-level-depth\formula-preview-ink-source-vs-v80-first30.json`
- v82 repeated-equation-paren-power sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v82-repeated-equation-paren-power\xsc测试集完整重建_61.docx`
- v82 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v82-repeated-equation-paren-power\leak-scan-61.json`
- v82 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v82-repeated-equation-paren-power\wmf-report-61.json`
- v82 script-relation ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v82-repeated-equation-paren-power\formula-preview-ink-source-vs-v82-script-relations.json`
- v82 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v82-repeated-equation-paren-power\formula-preview-ink-source-vs-v82-first30.json`
- v86 short-label-width-synced sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v86-short-label-width-synced\xsc测试集完整重建_61.docx`
- v86 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v86-short-label-width-synced\leak-scan-61.json`
- v86 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v86-short-label-width-synced\wmf-report-61.json`
- v86 target short-label ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v86-short-label-width-synced\formula-preview-ink-source-vs-v86-targets.json`
- v86 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v86-short-label-width-synced\formula-preview-ink-source-vs-v86-first30.json`
- v87 short-script-equation sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v87-short-script-equation\xsc测试集完整重建_61.docx`
- v87 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v87-short-script-equation\leak-scan-61.json`
- v87 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v87-short-script-equation\wmf-report-61.json`
- v87 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v87-short-script-equation\formula-preview-ink-source-vs-v87-first30.json`
- v90 short-object scoped sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v90-short-object-width-scoped-review\xsc测试集完整重建_61.docx`
- v90 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v90-short-object-width-scoped-review\leak-scan-61.json`
- v90 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v90-short-object-width-scoped-review\wmf-report-61.json`
- v90 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v90-short-object-width-scoped-review\formula-preview-ink-source-vs-v90-first30.json`
- v90 short-object targeted ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v90-short-object-width-scoped-review\formula-preview-ink-source-vs-v90-short-objects.json`
- v94 script-relation height sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v94-script-relation-height-scoped\xsc测试集完整重建_61.docx`
- v94 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v94-script-relation-height-scoped\leak-scan-61.json`
- v94 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v94-script-relation-height-scoped\wmf-report-61.json`
- v94 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v94-script-relation-height-scoped\formula-preview-ink-source-vs-v94-first30.json`
- v95 relation-width compensation split sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v95-relation-width-comp-split\xsc测试集完整重建_61.docx`
- v95 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v95-relation-width-comp-split\leak-scan-61.json`
- v95 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v95-relation-width-comp-split\wmf-report-61.json`
- v95 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v95-relation-width-comp-split\formula-preview-ink-source-vs-v95-first30.json`
- v95 sourceIndex 15/30 targeted ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v95-relation-width-comp-split\formula-preview-ink-source-vs-v95-targets.json`
- v96 standalone-two-letter height experiment, useful as a negative result:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v96-standalone-geometry-label-y\formula-preview-ink-source-vs-v96-first30.json`
- v97 AB/BD scoped label-y sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v97-ab-bd-label-y-scoped\xsc测试集完整重建_61.docx`
- v97 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v97-ab-bd-label-y-scoped\leak-scan-61.json`
- v97 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v97-ab-bd-label-y-scoped\wmf-report-61.json`
- v97 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v97-ab-bd-label-y-scoped\formula-preview-ink-source-vs-v97-first30.json`
- v97 geometry-label targeted ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v97-ab-bd-label-y-scoped\formula-preview-ink-source-vs-v97-geometry-labels.json`
- v98 script-relation height experiment is a mixed result: lowering
  `SCRIPT_RELATION_FONT_Y_SCALE` from `0.965` to `0.945` improved first-30
  height average from v97 `0.187pt` to `0.171pt` and max from `0.482pt` to
  `0.449pt`, but short equation source index 15 widened in the wrong direction
  after insufficient compensation, with width max worsening from `0.565pt` to
  `0.616pt`. Do not submit a relation-height change without rebalancing the
  affected width compensation constants.
- v99 keeps the v98 relation height reduction and rebalance compensation:
  `SCRIPT_RELATION_SHORT_EQUATION_WIDTH_COMPENSATION=1.038` and
  `SCRIPT_RELATION_TRIANGLE_WIDTH_COMPENSATION=1.015`. On doc 61 first-30,
  width average improved from v97 `0.198pt` to `0.191pt`, width max from
  `0.565pt` to `0.551pt`, height average from `0.187pt` to `0.171pt`, and
  height max from `0.482pt` to `0.449pt`. Structural checks still showed `520`
  valid MathType OLE, `520` vector WMF, zero visible LaTeX leaks, and zero
  bitmap WMF.
- v99 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v99-relation-y-width-balanced\xsc测试集完整重建_61.docx`
- v99 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v99-relation-y-width-balanced\leak-scan-61.json`
- v99 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v99-relation-y-width-balanced\wmf-report-61.json`
- v99 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v99-relation-y-width-balanced\formula-preview-ink-source-vs-v99-first30.json`
- v99 no-context review found no P0/P1 high-risk issue. Residual risk is
  coverage: doc 61 first-30 does not prove every formula matched by
  `scriptRelationFontYEligible`, so the next broad run should include more
  relation chains and visual spot checks for short equations, triangle chains,
  and long ratio chains.
- The leak and WMF report scripts use `--out`, not `--json-out`. If they fail
  with argparse errors, rerun `scripts/scan_docx_latex_leaks.py ... --out ...`
  and `scripts/wmf_record_report.py ... --out ...` before judging the document.
- v100 lowers `SCRIPT_FONT_HEIGHT_SCALE` from `0.92` to `0.90` and bumps
  `LaTeXImageRenderer.CACHE_VERSION` to
  `v100-xsc-short-script-height-tightened`. This is a small positive scoped
  to simple short scripts: doc 61 first-30 width average improved from v99
  `0.191pt` to `0.188pt`, and height average from `0.171pt` to `0.164pt`.
  The worst height stayed `0.449pt`, so `S_1` / `S_3` are not solved by script
  font height alone; their remaining height error is likely base glyph or
  overall vertical layout, not subscript font size.
- v100 structural checks on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap
  WMF, zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
  Aligned physical target metrics remained within the 1% guardrail for all
  `520` objects: shape width/height max error `0`, WMF width max error about
  `0.178%`, WMF height max error about `0.177%`.
- v100 no-context review found no P0/P1 issue. It confirmed
  `SCRIPT_FONT_HEIGHT_SCALE` flows through `simpleShortScriptScale()` and does
  not touch OLE/MTEF or relation-chain writing, while the cache version reaches
  the renderer cache key.
- v100 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v100-short-script-height\xsc测试集完整重建_61.docx`
- v100 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v100-short-script-height\leak-scan-61.json`
- v100 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v100-short-script-height\wmf-report-61.json`
- v100 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v100-short-script-height\formula-preview-ink-source-vs-v100-first30.json`
- v100 short-script ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v100-short-script-height\formula-preview-ink-source-vs-v100-short-scripts.json`
- v100 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v100-short-script-height\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v101 targets the remaining `S_1` / `S_3` height error by scaling the base
  regular glyph Y for standalone uppercase subscript formulas, not by lowering
  `SCRIPT_FONT_HEIGHT_SCALE` again. v100 proved that standalone `S_{n}` height
  was dominated by the base `S` glyph and overall vertical layout, so continuing
  to shrink only the script font is the wrong knob for that class.
- v101 adds `isStandaloneUpperSubscript` for simple objects such as `S_{1}` and
  deliberately rejects relation chains such as
  `S_{1}\colon S_{3}=a^{2}\colon b^{2}`. Keep relation-chain compensation
  separate; the remaining worst first-30 heights after v101 are still relation
  chains/sourceIndex `30`/`15` and simple linear items `1`/`2`/`11`, not the
  standalone `S_n` cases.
- v101 structural checks on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap
  WMF, zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
  Aligned physical target metrics remained within the 1% guardrail for all
  `520` objects: shape width/height max error `0`, WMF width max error about
  `0.178%`, WMF height max error about `0.177%`.
- v101 preview ink results: first-30 height average improved from v100
  `0.164pt` to `0.142pt`, max from `0.449pt` to `0.431pt`. Targeted short
  scripts source indices `7,8,9,10` improved height average from v100 `0.278pt`
  to `0.105pt`, and max from `0.449pt` to `0.112pt`.
- v101 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v101-upper-subscript-font-y\xsc测试集完整重建_61.docx`
- v101 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v101-upper-subscript-font-y\leak-scan-61.json`
- v101 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v101-upper-subscript-font-y\wmf-report-61.json`
- v101 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v101-upper-subscript-font-y\formula-preview-ink-source-vs-v101-first30.json`
- v101 short-script ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v101-upper-subscript-font-y\formula-preview-ink-source-vs-v101-short-scripts.json`
- v101 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v101-upper-subscript-font-y\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v102 relation-height experiment with `SCRIPT_RELATION_FONT_Y_SCALE=0.912`,
  `SCRIPT_RELATION_SHORT_EQUATION_WIDTH_COMPENSATION=1.058`, and
  `SCRIPT_RELATION_TRIANGLE_WIDTH_COMPENSATION=1.035` was over-aggressive.
  It improved targeted relation-chain height deltas for source indices `15`
  and `30` from about `0.425/0.431pt` to about `0.222/0.228pt`, but widened
  the worst first-30 width max from v101 `0.551pt` to `0.601pt` and pushed
  single short equations source indices `5`/`6` too short vertically. Do not
  reapply a global relation Y reduction without excluding single-relation
  formulas or rechecking short equations.
- v102b is a cache-version caution: changing relation constants while keeping
  `LaTeXImageRenderer.CACHE_VERSION` at the previous value reused stale WMF
  previews, so the generated ink metrics were identical to the older attempt.
  Any WMF/font/metric tuning must bump cache version, even if the output
  directory changes.
- v104 keeps the relation adjustment narrower: `SCRIPT_RELATION_FONT_Y_SCALE`
  is `0.93`, width compensations are `1.048` for `S_{1}=a^{2}=1` and `1.025`
  for bigtriangle relation chains, and `scriptRelationFontYEligible` now
  requires at least two top-level relation operators. This preserves chain
  formulas such as `S_{3}=4=b^{2}` while excluding single-relation short
  formulas such as `S_{2}=2`.
- v104 on doc 61 is a positive height step but not final: first-30 height avg
  improved from v101 `0.142pt` to `0.119pt`, and height max from `0.431pt` to
  `0.330pt`. Structural and physical gates still passed with `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero bitmap WMF,
  target WMF width/height `520/520` within 1%, and shape width/height `520/520`
  within 1%. Width remains the next issue: first-30 width avg moved from v101
  `0.188pt` to `0.191pt`, and width max worsened from `0.551pt` to `0.601pt`.
- v104 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v104-relation-chain-height\xsc测试集完整重建_61.docx`
- v104 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v104-relation-chain-height\leak-scan-61.json`
- v104 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v104-relation-chain-height\wmf-report-61.json`
- v104 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v104-relation-chain-height\formula-preview-ink-source-vs-v104-first30.json`
- v104 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v104-relation-chain-height\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v105 fixes the v104 triangle-chain width regression by adding a separate
  `SCRIPT_RELATION_TRIANGLE_WIDTH_SCALE=0.98`. The key lesson is that
  `scriptRelationHeightWidthCompensation` changes WMF `ExtTextOut` `dx`
  advances but does not change the layout width used by `previewScale`; for
  long formulas near the display box edge, dx-only widening can be clipped.
  Use `scriptRelationWidthScale` when the correction must happen in layout
  space.
- v105 on doc 61 improved first-30 ink width average from v104 `0.191pt` to
  `0.172pt` and width max from `0.601pt` to `0.465pt`, while preserving v104
  height results: height average `0.119pt`, height max `0.330pt`. It is also
  better than v101 on width avg/max. The previous triangle worst sourceIndex
  `30` dropped out of the worst-width list; remaining width targets are
  sourceIndex `15` (`S_{1}=a^{2}=1`, `-0.465pt`), sourceIndex `14` (`a=1`,
  `-0.455pt`), and sourceIndex `18` repeated paren equation (`+0.448pt`).
- v105 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, target WMF width/height `520/520` within 1%, and shape width/height
  `520/520` within 1%.
- v105 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v105-triangle-chain-width\xsc测试集完整重建_61.docx`
- v105 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v105-triangle-chain-width\leak-scan-61.json`
- v105 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v105-triangle-chain-width\wmf-report-61.json`
- v105 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v105-triangle-chain-width\formula-preview-ink-source-vs-v105-first30.json`
- v105 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v105-triangle-chain-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v106 adds an exact `S_{1}=a^{2}=1` relation-equation balance instead of
  changing the whole script-relation class. The formula now uses
  `SCRIPT_RELATION_EXACT_S1_EQUATION_FONT_Y_SCALE=0.905` and
  `SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION=1.075`; neighboring
  formulas such as `S_{3}=4=b^{2}` and `S_{2}=2=a\times b` keep their previous
  behavior. This is deliberately narrow because v102 already showed that broad
  relation Y reductions damage short equations.
- v106 is a small positive step, not a final solve for doc 61. First-30 width
  average changed from v105 `0.172pt` to `0.171pt`, width max from `0.465pt`
  to `0.455pt`, height average from `0.119pt` to `0.116pt`, and height max
  stayed `0.330pt`. SourceIndex `15` improved from width/height deltas
  `-0.465pt/+0.323pt` to `-0.414pt/+0.222pt`. Do not keep increasing this
  exact compensation blindly; the next top width target is sourceIndex `14`
  (`a=1`, `-0.455pt`), while the next top height target is still sourceIndex
  `30` (`+0.330pt`).
- v106 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, target WMF width/height `520/520` within 1%, and target shape
  width/height `520/520` within 1%.
- v106 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v106-s1-area-equation-balance\xsc测试集完整重建_61.docx`
- v106 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v106-s1-area-equation-balance\leak-scan-61.json`
- v106 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v106-s1-area-equation-balance\wmf-report-61.json`
- v106 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v106-s1-area-equation-balance\formula-preview-ink-source-vs-v106-first30.json`
- v106 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v106-s1-area-equation-balance\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v107 targets the short exact equation `a=1` through
  `shortExactEquationWidthScale`, not through relation-chain logic. It adds
  `SHORT_A_EQUALS_ONE_WIDTH_SCALE=1.025` while preserving `b=2` at `0.965`
  and leaving ratio/proportion formulas such as `a\colon b=5\colon 7` at
  `1.0`. This is the right entry point for short ordinary equations because
  `a=1` has a width problem but almost no height problem.
- v107 on doc 61 moved sourceIndex `14` (`a=1`) out of the worst-width list.
  First-30 width average improved from v106 `0.171pt` to `0.157pt`, width max
  from `0.455pt` to `0.448pt`, and height average/max stayed `0.116pt` /
  `0.330pt`. The new top width residuals are sourceIndex `18` (`+0.448pt`)
  and sourceIndex `15` (`-0.414pt`); the top height residual remains
  sourceIndex `30` (`+0.330pt`).
- v107 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, target WMF width/height `520/520` within 1%, and target shape
  width/height `520/520` within 1%.
- v107 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v107-a-equals-one-width\xsc测试集完整重建_61.docx`
- v107 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v107-a-equals-one-width\leak-scan-61.json`
- v107 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v107-a-equals-one-width\wmf-report-61.json`
- v107 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v107-a-equals-one-width\formula-preview-ink-source-vs-v107-first30.json`
- v107 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v107-a-equals-one-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v108 narrows only repeated parenthesized-power equations by changing
  `EQUATION_PAREN_POWER_WIDTH_SCALE` from `0.93` to `0.925` and bumping the
  cache to `v108-xsc-repeated-paren-width`. This is deliberately scoped to
  `repeatedEquationParenPower`, so single formulas like
  `S=\left(a+b\right)^2=9` and standalone `\left(a+b\right)^2` keep their
  existing behavior.
- A failed test-tightening attempt showed that the existing WMF unit-level
  assertions (`<2200`) are too coarse for proving this small change; reducing
  them to `<2190` failed even though the real DOCX ink metric improved. For
  micro width tuning, keep unit tests as behavioral guards and trust the
  rendered DOCX ink comparison for acceptance evidence.
- v108 on doc 61 improved sourceIndex `18`
  (`S=\left(a+b\right)^{2}=\left(1+2\right)^{2}=9`) from width delta
  `+0.448pt` to `+0.348pt`. First-30 width average improved from v107
  `0.157pt` to `0.154pt`, and width max improved from `0.448pt` to `0.414pt`.
  Height average/max stayed `0.116pt` / `0.330pt`. The next visible targets
  are sourceIndex `15` (`S_{1}=a^{2}=1`, `-0.414pt`) and the relation-chain
  height group led by sourceIndex `30` (`+0.330pt`).
- v108 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v108 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v108-repeated-paren-width\xsc测试集完整重建_61.docx`
- v108 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v108-repeated-paren-width\leak-scan-61.json`
- v108 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v108-repeated-paren-width\wmf-report-61.json`
- v108 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v108-repeated-paren-width\formula-preview-ink-source-vs-v108-first30.json`
- v108 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v108-repeated-paren-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v109 continues the exact `S_{1}=a^{2}=1` branch with a small balance change:
  `SCRIPT_RELATION_EXACT_S1_EQUATION_FONT_Y_SCALE` moves from `0.905` to
  `0.900`, and `SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION` moves
  from `1.075` to `1.085`. This is not whitespace padding; it uses the
  existing exact formula class to make the rendered MathType-style object a bit
  shorter and wider.
- The reason this combination works is that `previewScale` limits horizontal
  scale when a non-grid formula is height-compressed. For `S_{1}=a^{2}=1`, the
  residual was both too tall and too narrow, so lowering the exact Y scale and
  increasing the exact dx compensation is a coherent paired move. Do not apply
  this reasoning to broad relation formulas without rechecking width, because
  earlier v102-style global height cuts widened regressions elsewhere.
- v109 on doc 61 improved sourceIndex `15` from width/height deltas
  `-0.414pt/+0.222pt` to `-0.314pt/+0.171pt`. First-30 width average improved
  from v108 `0.154pt` to `0.151pt`, width max from `0.414pt` to `0.381pt`,
  height average from `0.116pt` to `0.114pt`, and height max stayed `0.330pt`.
  The next visible width residuals are sourceIndex `5` (`S_{2}=2`,
  `+0.381pt`) and sourceIndex `18` (`+0.348pt`); the next top height residual
  remains the relation-chain group led by sourceIndex `30` (`+0.330pt`).
- v109 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v109 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v109-s1-area-balance\xsc测试集完整重建_61.docx`
- v109 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v109-s1-area-balance\leak-scan-61.json`
- v109 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v109-s1-area-balance\wmf-report-61.json`
- v109 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v109-s1-area-balance\formula-preview-ink-source-vs-v109-first30.json`
- v109 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v109-s1-area-balance\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v110 adds an exact `S_{2}=2` width scale instead of tightening the whole
  short-script-equation class. `SHORT_S2_EQUALS_TWO_WIDTH_SCALE=0.985` is
  applied through `shortExactEquationWidthScale`, alongside the existing
  `a=1` and `b=2` exact corrections. This avoids changing broader formulas
  such as `a^{2}=1`, `S^{2}=4`, or `S_{2}=2.5`.
- The v110 choice is based on v109 residuals: sourceIndex `5` (`S_{2}=2`) was
  the top first-30 width error at `+0.381pt` while its height error was only
  `+0.105pt`. A pure exact-width adjustment is therefore safer than another
  font-Y change.
- v110 on doc 61 improved sourceIndex `5` width delta from `+0.381pt` to
  `+0.231pt`. First-30 width average improved from v109 `0.151pt` to
  `0.146pt`, width max from `0.381pt` to `0.348pt`, and height average/max
  stayed `0.114pt` / `0.330pt`. The next visible width target is again
  sourceIndex `18` (`+0.348pt`), and the top height residual remains the
  relation-chain group led by sourceIndex `30` (`+0.330pt`).
- v110 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v110 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v110-s2-short-equation-width\xsc测试集完整重建_61.docx`
- v110 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v110-s2-short-equation-width\leak-scan-61.json`
- v110 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v110-s2-short-equation-width\wmf-report-61.json`
- v110 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v110-s2-short-equation-width\formula-preview-ink-source-vs-v110-first30.json`
- v110 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v110-s2-short-equation-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v111 continues the repeated parenthesized-power equation branch by moving
  `EQUATION_PAREN_POWER_WIDTH_SCALE` from `0.925` to `0.920` and bumping the
  cache to `v111-xsc-repeated-paren-width`. This is still scoped by
  `repeatedEquationParenPower(...)`, so the single formula
  `S=\left(a+b\right)^2=9` remains outside this correction; the intent is not
  to use spacing, but to shrink the closing-fence superscript segments inside
  repeated equation forms.
- v111 on doc 61 improved sourceIndex `18`
  (`S=\left(a+b\right)^{2}=\left(1+2\right)^{2}=9`) width delta from
  `+0.348pt` to `+0.248pt`. First-30 width average improved from v110
  `0.146pt` to `0.142pt`, and width max improved from `0.348pt` to
  `0.327pt`. Height average/max stayed `0.114pt` / `0.330pt`. The next top
  visible width residuals are sourceIndex `1` (`+0.327pt`), sourceIndex `15`
  (`-0.314pt`), and sourceIndex `7` (`+0.275pt`); the top height residual is
  still the relation-chain group led by sourceIndex `30` (`+0.330pt`).
- v111 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v111 no-context review found no P0/P1/P2 issues. The review specifically
  checked that the scale is only applied when `equationParenPower` is true,
  that `S=\left(a+b\right)^2=9` is covered as a false boundary case, and that
  `CACHE_VERSION` participates in the cache key.
- v111 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v111-repeated-paren-width\xsc测试集完整重建_61.docx`
- v111 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v111-repeated-paren-width\leak-scan-61.json`
- v111 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v111-repeated-paren-width\wmf-report-61.json`
- v111 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v111-repeated-paren-width\formula-preview-ink-source-vs-v111-first30.json`
- v111 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v111-repeated-paren-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v112 keeps the final submitted change narrow: only exact
  `S_{1}=a^{2}=1` gets a width-compensation bump from `1.085` to `1.093`,
  with cache version `v112-xsc-s1-area-width`. This is scoped through
  `isExactS1AreaEquation(...)` and therefore does not alter ordinary relation
  chains, long relation chains, or `\bigtriangleup` relation chains.
- Two v112 experiments were intentionally discarded before commit. Lowering
  `SCRIPT_RELATION_FONT_Y_SCALE` from `0.93` to `0.925` improved width max but
  worsened first-30 height average from `0.114pt` to `0.119pt` and did not
  move the `0.330pt` max height. Lowering ordinary
  `SCRIPT_RELATION_WIDTH_SCALE` from `0.975` to `0.970` over-shrank exact
  `S_{1}=a^{2}=1`, worsening its width delta to `-0.515pt`. Do not repeat
  those broad relation moves without a new mechanism that separates exact S1,
  ordinary ratio chains, long chains, and triangle chains.
- v112 final on doc 61 improved sourceIndex `15` (`S_{1}=a^{2}=1`) width
  delta from `-0.314pt` to `-0.264pt`. First-30 width average improved from
  v111 `0.142pt` to `0.141pt`, p90 from `0.264pt` to `0.263pt`, and
  height average/max stayed `0.114pt` / `0.330pt`. Width max stayed `0.327pt`,
  now led by sourceIndex `1`; top height remains the relation-chain group led
  by sourceIndex `30`.
- v112 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v112 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v112-s1-area-width\xsc测试集完整重建_61.docx`
- v112 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v112-s1-area-width\leak-scan-61.json`
- v112 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v112-s1-area-width\wmf-report-61.json`
- v112 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v112-s1-area-width\formula-preview-ink-source-vs-v112-first30.json`
- v112 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v112-s1-area-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v113b narrows only standalone uppercase subscript formulas by changing
  `STANDALONE_UPPER_SUBSCRIPT_WIDTH_SCALE` from `0.82` to `0.814`, with cache
  version `v113b-xsc-upper-subscript-balanced`. The scope remains limited by
  `standaloneShortScriptWidthScale(...)` and `isStandaloneUpperSubscript(...)`:
  formulas containing `=`, `\colon`, `+`, `-`, `\times`, `\div`,
  `\left`/`\right`, fractions, roots, or arrays do not take this path.
- The first v113 trial used `0.805` and improved sourceIndex `7`
  (`S_{1}`) from `+0.275pt` to `+0.175pt`, but it over-shrank the same
  standalone-uppercase-subscript class at sourceIndex `9` from about
  `-0.194pt` to `-0.294pt`. The final `0.814` is the balanced value: sourceIndex
  `7` improves to `+0.225pt`, while sourceIndex `9` lands at `-0.244pt`, so the
  two visible class representatives are both closer than the over-tightened
  trial.
- A same-round retune must still bump `LaTeXImageRenderer.CACHE_VERSION`.
  Reusing `v113-xsc-upper-subscript-width` after changing `0.805` to `0.814`
  produced identical ink numbers because cached previews were reused. Treat
  unchanged ink output after a renderer constant edit as a cache-version smell
  before trusting the measurement.
- v113b first-30 ink on doc 61 keeps first-30 width average at `0.141pt`,
  improves p90 from v112 `0.263pt` to `0.260pt`, and keeps height average/max at
  `0.114pt` / `0.330pt`. The remaining top residuals are still relation-chain
  and exact/short-number cases rather than this standalone `S_{1}` class.
- v113b structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v113b no-context review found no P0/P1/P2 issues. It confirmed that the
  standalone upper-subscript scale is excluded from `S_{1}\colon S_{3}=...`,
  `S_{1}=a^{2}=1`, `a^{2}`, and `S=\left(a+b\right)^2=9`, and suggested a
  future harder test that directly asserts the scale path or a narrow
  `S_{1}` golden width.
- v113b sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v113b-upper-subscript-balanced\xsc测试集完整重建_61.docx`
- v113b leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v113b-upper-subscript-balanced\leak-scan-61.json`
- v113b WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v113b-upper-subscript-balanced\wmf-report-61.json`
- v113b first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v113b-upper-subscript-balanced\formula-preview-ink-source-vs-v113b-first30.json`
- v113b aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v113b-upper-subscript-balanced\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v114 adds a separate simple-ratio relation width branch instead of touching
  the broad `SCRIPT_RELATION_WIDTH_SCALE`. `SCRIPT_RELATION_SIMPLE_RATIO_WIDTH_SCALE`
  is `0.970` and only applies when there are exactly three top-level relation
  operators, the formula contains a raw `:` or `\colon` plus `=`, and the
  formula is not a `\bigtriangleup` relation or exact `S_{1}=a^{2}=1`. This
  preserves long chains at `0.987`, triangle chains at `0.98`, and ordinary
  three-part equations such as `S_{3}=4=b^{2}` at `0.975`.
- v114 first hit a useful test boundary: the old test expected raw `:` and
  `\colon` versions of `S_{1}:S_{3}=a^{2}:b^{2}` to have identical WMF text
  advance. The simple-ratio branch must therefore recognize both raw colon and
  `\colon`, not just `\colon`, or the renderer will split visually equivalent
  authoring styles.
- v114 on doc 61 improved the first-30 ink width average from v113b `0.141pt`
  to `0.131pt`, median from `0.131pt` to `0.110pt`, p90 from `0.260pt` to
  `0.249pt`, and max from `0.327pt` to `0.264pt`. The former top width residual
  sourceIndex `1` (`S_{1}\colon S_{3}=a^{2}\colon b^{2}`) dropped out of the
  worst-width list. Height average/max stayed `0.114pt` / `0.330pt`; the
  relation-chain height issue remains for a later mechanism.
- v114 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v114 no-context review found no P0/P1/P2 issues. It confirmed the new branch
  does not affect exact `S_{1}=a^{2}=1`, long chains, triangle chains,
  `S_{3}=4=b^{2}`, or `a^{2}+b`, and recommended adding direct scale assertions
  for raw colon, exact S1, and `S_{3}=4=b^{2}`; those assertions were added
  before commit.
- v114 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v114-simple-ratio-width\xsc测试集完整重建_61.docx`
- v114 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v114-simple-ratio-width\leak-scan-61.json`
- v114 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v114-simple-ratio-width\wmf-report-61.json`
- v114 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v114-simple-ratio-width\formula-preview-ink-source-vs-v114-first30.json`
- v114 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v114-simple-ratio-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v115 adds a triangle-relation-only font Y branch instead of lowering the
  broad `SCRIPT_RELATION_FONT_Y_SCALE` again. Triangle relation chains that
  contain `\bigtriangleup`, have at least four top-level relation operators,
  and already satisfy the script-relation font-Y eligibility now use
  `SCRIPT_RELATION_TRIANGLE_FONT_Y_SCALE = 0.905`; exact `S_{1}=a^{2}=1`
  remains on its existing `0.90` branch, while ordinary relation chains keep
  `0.93`.
- The value is intentionally class-specific. The v112 broad trial at `0.925`
  worsened height average and did not move the worst `0.330pt` case, so v115
  targets the triangle residual directly: previous triangle behavior was
  effectively `0.93`, and `0.93 * 0.9728` is about `0.905`.
- v115 keeps the existing triangle height-width compensation active because
  `scriptRelationHeightWidthCompensation(...)` still sees
  `scriptRelationFontYScale(latex) < 0.999` for triangle chains. The helper
  `isTriangleScriptRelation(...)` calls `scriptRelationFontYEligible(...)`,
  which depends on `scriptRelationWidthScale(...)`; it does not recurse back
  into `scriptRelationFontYScale(...)`.
- v115 on doc 61 improved first-30 ink height average from v114 `0.114pt` to
  `0.111pt`, p90 from `0.311pt` to `0.264pt`, and max from `0.330pt` to
  `0.311pt`. The sourceIndex `30` triangle relation height residual improved
  from `+0.330pt` to `+0.228pt`.
- v115 has a small width tradeoff: first-30 ink width average moved from
  v114 `0.131pt` to `0.132pt`, median from `0.110pt` to `0.116pt`, and max
  stayed `0.264pt`. Do not present this as final convergence; it is a net
  height improvement with an acceptable but measurable width cost.
- Remaining visible height work after v115 is mainly ordinary relation groups,
  especially sourceIndex `1`, `2`, and `11` at about `+0.311pt`, plus paren-power
  cases such as sourceIndex `4` and `18` around `-0.259pt`. Avoid another broad
  relation-Y retune without splitting these classes and rechecking width.
- v115 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v115 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v115-triangle-relation-height\xsc测试集完整重建_61.docx`
- v115 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v115-triangle-relation-height\leak-scan-61.json`
- v115 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v115-triangle-relation-height\wmf-report-61.json`
- v115 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v115-triangle-relation-height\formula-preview-ink-source-vs-v115-first30.json`
- v115 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v115-triangle-relation-height\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v116 was a rejected broad experiment: changing the generic
  `SCRIPT_RELATION_FONT_Y_SCALE` from `0.93` to `0.906` improved the ordinary
  short relation height representatives, but it also hit long relation/additive
  equations. On doc 61 first-30 ink, sourceIndex `1`, `2`, and `11` height
  residuals moved from `+0.311pt` to `+0.210pt`, but width max worsened from
  v115 `0.264pt` to `0.413pt`, width average from `0.132pt` to `0.146pt`, and
  height average from `0.111pt` to `0.119pt`. Do not reapply a broad relation
  font-Y reduction without a separate width compensation plan.
- v117 keeps the generic relation font-Y at `0.93` and adds a compact relation
  branch at `0.906`. The compact branch only applies to eligible script
  relations with two or three top-level relation operators, excludes
  `\bigtriangleup`, and excludes formulas containing `\times`, `\div`, `+`, or
  `-`. This covers short ratios such as `S_{1}\colon S_{3}=a^{2}\colon b^{2}`
  and `S_{3}=4=b^{2}` without touching the long chain
  `S_{1}\colon S_{3}\colon S_{2}\colon S_{4}=...` or the additive total
  equation `S=S_{1}+S_{2}+S_{3}+S_{4}=...`.
- v117 on doc 61 improved first-30 ink versus v115: width average
  `0.132pt -> 0.130pt`, width median `0.116pt -> 0.094pt`, width max stayed
  `0.264pt`; height average `0.111pt -> 0.104pt`, height p90
  `0.264pt -> 0.231pt`, and height max stayed `0.311pt`. SourceIndex `1` and
  `11` height residuals improved from `+0.311pt` to `+0.210pt`; sourceIndex
  `2` remains `+0.311pt` because it is the protected long-chain case.
- v117 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v117 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v117-compact-relation-height\xsc测试集完整重建_61.docx`
- v117 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v117-compact-relation-height\leak-scan-61.json`
- v117 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v117-compact-relation-height\wmf-report-61.json`
- v117 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v117-compact-relation-height\formula-preview-ink-source-vs-v117-first30.json`
- v117 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v117-compact-relation-height\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v118d fixes a cache/process pitfall before keeping the result: a previous
  v118 paren-power experiment reused a cache version and produced stale,
  byte-identical WMF for the target formulas. After any renderer metric or WMF
  logic change, use a new `LaTeXImageRenderer.CACHE_VERSION`; do not judge a
  failed trial unless the cache key proves the new renderer actually ran.
- v118d adds a narrow closing-fence-superscript height branch for paren-power
  formulas. `PAREN_POWER_FONT_Y_SCALE = 1.019` applies only when a superscript
  is attached to a closing fence such as `\right)` or `)`, and still excludes
  fractions, sqrt, `\begin`, `\over`, `\under`, and boxed formulas. Width is
  compensated with `STANDALONE_PAREN_POWER_WIDTH_SCALE = 0.924` and
  `EQUATION_PAREN_POWER_WIDTH_SCALE = 0.916`.
- v118d on doc 61 improved the first-30 ink height average from v117 `0.104pt`
  to `0.094pt` and p90 from `0.231pt` to `0.210pt`. The target paren-power
  cases sourceIndex `4` and `18` improved from about `-0.259pt` height residual
  to about `-0.111pt`. This is a real visual improvement but not final
  convergence: sourceIndex `2` remains the max height residual at about
  `+0.311pt`.
- v118d has a recorded width tradeoff. Width average and median stayed at
  v117 levels (`0.130pt` and `0.094pt`), but p90 moved from `0.249pt` to
  `0.260pt` and max from `0.264pt` to `0.298pt`, mainly from sourceIndex `18`.
  Keep this branch only as a scoped height improvement; do not hide the width
  cost when planning the next round.
- v118d structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v118d sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v118d-paren-power-height-width\xsc测试集完整重建_61.docx`
- v118d leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v118d-paren-power-height-width\leak-scan-61.json`
- v118d WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v118d-paren-power-height-width\wmf-report-61.json`
- v118d first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v118d-paren-power-height-width\formula-preview-ink-source-vs-v118d-first30.json`
- v118d aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v118d-paren-power-height-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v119c targets the remaining long pure ratio-chain height residual instead of
  broad relation-Y retuning. Long non-triangle script relations with at least
  six top-level relation operators, and without `+`, `-`, `\times`, or `\div`,
  use `SCRIPT_RELATION_LONG_CHAIN_FONT_Y_SCALE = 0.906`. This intentionally
  covers `S_{1}\colon S_{3}\colon S_{2}\colon S_{4}=...` and protects additive
  totals such as `S=S_{1}+S_{2}+S_{3}+S_{4}=...`.
- v119 first lowered long-chain Y without enough width compensation. It reduced
  first-30 height max from v118d `0.311pt` to `0.228pt`, but widened the worst
  width residual to `0.413pt`; do not keep that form. v119c adds
  `SCRIPT_RELATION_LONG_CHAIN_WIDTH_COMPENSATION = 1.022`, bringing first-30
  width max back to v118d's `0.298pt` while preserving the height gain.
- v119c on doc 61 first-30 ink versus v118d: width average stayed `0.130pt`,
  median stayed `0.094pt`, p90 stayed `0.260pt`, and max stayed `0.298pt`.
  Height average improved from `0.094pt` to `0.091pt`, median stayed `0.080pt`,
  p90 stayed `0.210pt`, and max improved from `0.311pt` to `0.228pt`.
  SourceIndex `2` height moved from `+0.311pt` to `+0.210pt`; its width stayed
  about `-0.263pt`.
- v119c structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v119c repeated the cache-version lesson: after changing only width
  compensation, regenerating into a new output directory was not enough; the
  old `v119b` cache key reused stale WMF and made v119b look unchanged. Bump
  `LaTeXImageRenderer.CACHE_VERSION` for every renderer metric change before
  comparing visual output.
- `rebuild\compare_formula_preview_ink.py` needs Pillow. If plain `python`
  reports `ModuleNotFoundError: No module named 'PIL'`, use the bundled runtime
  `C:\Users\11703\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`.
- v119c sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v119c-long-chain-relation-height-width\xsc测试集完整重建_61.docx`
- v119c leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v119c-long-chain-relation-height-width\leak-scan-61.json`
- v119c WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v119c-long-chain-relation-height-width\wmf-report-61.json`
- v119c first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v119c-long-chain-relation-height-width\formula-preview-ink-source-vs-v119c-first30.json`
- v119c aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v119c-long-chain-relation-height-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v120 targets the triangle-ratio relation height only. Lowering
  `SCRIPT_RELATION_TRIANGLE_FONT_Y_SCALE` from `0.905` to `0.887` improved
  sourceIndex `30` height residual from about `+0.228pt` to `+0.127pt`.
  Do not compensate this case by widening the triangle relation branch:
  `SCRIPT_RELATION_TRIANGLE_WIDTH_COMPENSATION = 1.045` and `1.030` both
  overextended the test formula boundary, so keep the prior `1.025`.
- v120 is a scoped height tradeoff, not a full convergence. On doc 61 first-30
  ink versus v119c, height average improved from `0.091pt` to `0.087pt`, p90
  improved from `0.210pt` to `0.175pt`, and max improved from `0.228pt` to
  `0.210pt`. Width max stayed `0.298pt`, but width average moved from
  `0.130pt` to `0.132pt` and median from `0.094pt` to `0.110pt`; record this
  honestly when comparing future rounds.
- v120 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v120 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v120-triangle-relation-height\xsc测试集完整重建_61.docx`
- v120 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v120-triangle-relation-height\leak-scan-61.json`
- v120 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v120-triangle-relation-height\wmf-report-61.json`
- v120 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v120-triangle-relation-height\formula-preview-ink-source-vs-v120-first30.json`
- v120 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v120-triangle-relation-height\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v121 targets compact script relations only. Lowering
  `SCRIPT_RELATION_COMPACT_FONT_Y_SCALE` from `0.906` to `0.890` improves the
  repeated compact relation cases such as `S_{1}\colon S_{3}=a^{2}\colon b^{2}`
  and `S_{3}=4=b^{2}` without touching long-chain, triangle, exact `S_{1}`,
  additive-total, or standalone script formulas.
- v121 on doc 61 first-30 ink versus v120: height average improved from
  `0.087pt` to `0.080pt`, p90 improved from `0.175pt` to `0.114pt`, and max
  stayed `0.210pt`. SourceIndex `1` and `11` improved from about `+0.210pt`
  height residual to about `+0.108pt`. Width max stayed `0.298pt`, while width
  average moved slightly from `0.132pt` to `0.133pt` and median from `0.110pt`
  to `0.126pt`; keep recording this small width cost.
- Do not keep the v121b long-chain height experiment as-is. Lowering
  `SCRIPT_RELATION_LONG_CHAIN_FONT_Y_SCALE` from `0.906` to `0.890` improved
  first-30 height max from `0.210pt` to `0.171pt`, but worsened sourceIndex `2`
  width residual from about `-0.263pt` to about `-0.313pt`, making width max
  worse than the v120/v121 baseline. Any future long-chain height work needs
  separate width compensation, not just lowering the Y scale.
- v121 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v121 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v121-compact-relation-height\xsc测试集完整重建_61.docx`
- v121 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v121-compact-relation-height\leak-scan-61.json`
- v121 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v121-compact-relation-height\wmf-report-61.json`
- v121 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v121-compact-relation-height\formula-preview-ink-source-vs-v121-first30.json`
- v121 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v121-compact-relation-height\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v122c targets only exact `S_{1}=a^{2}=1`. Lowering
  `SCRIPT_RELATION_EXACT_S1_EQUATION_FONT_Y_SCALE` from `0.90` to `0.887`
  improves sourceIndex `15` height residual from about `+0.171pt` to
  `+0.120pt`; keep `SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION` at
  `1.100` with this height scale, because the same height scale with `1.093`
  worsened sourceIndex `15` width residual to about `-0.314pt`.
- Do not keep the v122b exact-S1 experiment as-is. Lowering exact S1 Y further
  to `0.878` removed sourceIndex `15` from the worst-height list, but it made
  the first-30 width max worse at `0.314pt`. For this family, height tightening
  and width compensation must be tuned together.
- v122c on doc 61 first-30 ink versus v121: height average improved from
  `0.080pt` to `0.079pt`, p90 improved from `0.114pt` to `0.113pt`, and max
  stayed `0.210pt`. Width average stayed `0.133pt`, median stayed `0.126pt`,
  p90 stayed `0.260pt`, and max stayed `0.298pt`.
- v122c structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v122c sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v122c-exact-s1-relation-height-width\xsc测试集完整重建_61.docx`
- v122c leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v122c-exact-s1-relation-height-width\leak-scan-61.json`
- v122c WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v122c-exact-s1-relation-height-width\wmf-report-61.json`
- v122c first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v122c-exact-s1-relation-height-width\formula-preview-ink-source-vs-v122c-first30.json`
- v122c aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v122c-exact-s1-relation-height-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v123b targets long-chain script relations such as
  `S_{1}\colon S_{3}\colon S_{2}\colon S_{4}=a^{2}\colon b^{2}\colon ab\colon ab`.
  Lowering `SCRIPT_RELATION_LONG_CHAIN_FONT_Y_SCALE` from `0.906` to `0.890`
  fixes the old sourceIndex `2` height outlier, but only if the actual X
  layout scale is widened from `0.987` to `0.9895`. Do not rely on
  `SCRIPT_RELATION_LONG_CHAIN_WIDTH_COMPENSATION` alone for this case; that
  changes the preview box more than the ink width.
- Do not keep the v123 first attempt as-is. `SCRIPT_RELATION_LONG_CHAIN_FONT_Y_SCALE
  = 0.890` with `SCRIPT_RELATION_LONG_CHAIN_WIDTH_COMPENSATION = 1.026`
  improved first-30 height max from `0.210pt` to `0.127pt`, but worsened width
  max from `0.298pt` to `0.313pt`. The kept v123b branch returns the
  compensation to `1.022` and widens the long-chain ink scale instead.
- v123b on doc 61 first-30 ink versus v122c: height average improved from
  `0.079pt` to `0.075pt`, p90 improved from `0.113pt` to `0.112pt`, and max
  improved from `0.210pt` to `0.127pt`. Width average improved from `0.133pt`
  to `0.117pt`, median improved from `0.126pt` to `0.082pt`, p90 improved from
  `0.260pt` to `0.246pt`, and max stayed `0.298pt`.
- v123b structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal target
  shape width/height `520/520` within 1%.
- v123b sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v123b-long-chain-relation-height-xscale\xsc测试集完整重建_61.docx`
- v123b leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v123b-long-chain-relation-height-xscale\leak-scan-61.json`
- v123b WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v123b-long-chain-relation-height-xscale\wmf-report-61.json`
- v123b first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v123b-long-chain-relation-height-xscale\formula-preview-ink-source-vs-v123b-first30.json`
- v123b aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v123b-long-chain-relation-height-xscale\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v124 targets repeated parenthesized-power equations such as
  `S=\left(a+b\right)^2=\left(1+2\right)^2=9`, which is doc 61 formula
  object/sourceIndex `18`. Do not treat this object as an `AB`/geometry label;
  mapping through `61.tex` and `oleObject18` shows it is the repeated
  parenthesized-power equation. Lowering `EQUATION_PAREN_POWER_WIDTH_SCALE`
  from `0.916` to `0.913` reduced its visible ink width residual from
  `+0.298pt` to `+0.248pt`.
- v124 on doc 61 first-30 ink versus v123b: width average improved from
  `0.117pt` to `0.115pt`, p90 improved from `0.246pt` to `0.244pt`, and max
  improved from `0.298pt` to `0.264pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`.
- v124 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v124 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v124-repeated-paren-equation-width\xsc测试集完整重建_61.docx`
- v124 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v124-repeated-paren-equation-width\leak-scan-61.json`
- v124 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v124-repeated-paren-equation-width\wmf-report-61.json`
- v124 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v124-repeated-paren-equation-width\formula-preview-ink-source-vs-v124-first30.json`
- v124 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v124-repeated-paren-equation-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v125 retunes only standalone two-digit objects by changing
  `STANDALONE_TWO_DIGIT_WIDTH_SCALE` from `0.81` to `0.79`, with cache version
  `v125-standalone-two-digit-width`. Keep this scoped to formulas whose entire
  normalized text is exactly two digits; long formulas containing numbers such
  as `S=25+35` must not inherit the standalone digit scale.
- v125 on doc 61 first-30 ink versus v124: width average improved from
  `0.115pt` to `0.105pt`, p90 improved from `0.244pt` to `0.232pt`, and max
  stayed `0.264pt`. Height average, p90, and max stayed `0.075pt`, `0.112pt`,
  and `0.127pt`. The two standalone digit targets improved from v124
  sourceIndex `27/28` width residuals `+0.242pt/+0.260pt` to
  `+0.089pt/+0.108pt`.
- v125 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v125 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v125-standalone-two-digit-width\xsc测试集完整重建_61.docx`
- v125 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v125-standalone-two-digit-width\leak-scan-61.json`
- v125 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v125-standalone-two-digit-width\wmf-report-61.json`
- v125 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v125-standalone-two-digit-width\formula-preview-ink-source-vs-v125-first30.json`
- v125 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v125-standalone-two-digit-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v126 targets only exact `S_{1}=a^{2}=1` width after the v122c height tuning.
  Raising `SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION` from `1.100`
  to `1.107` moves sourceIndex `15` visible ink width residual from v125
  `-0.264pt` to `-0.213pt` without changing height. This is the opposite
  direction from the failed v122b/v122c-side experiment that lowered the
  compensation to `1.093` and made width worse.
- v126 on doc 61 first-30 ink versus v125: width average improved from
  `0.105pt` to `0.103pt`, p90 improved from `0.232pt` to `0.226pt`, and max
  improved from `0.264pt` to `0.248pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`.
- Keep the exact-S1 width compensation scoped through
  `isExactS1AreaEquation(...)`. Broad relation compensation still risks
  damaging compact relations such as `S_{3}=4=b^{2}` and additive totals; do
  not merge this path with generic script-relation width scaling.
- v126 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v126 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v126-exact-s1-relation-width\xsc测试集完整重建_61.docx`
- v126 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v126-exact-s1-relation-width\leak-scan-61.json`
- v126 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v126-exact-s1-relation-width\wmf-report-61.json`
- v126 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v126-exact-s1-relation-width\formula-preview-ink-source-vs-v126-first30.json`
- v126 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v126-exact-s1-relation-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- v127 continues the repeated parenthesized-power equation tuning after v124.
  Lowering `EQUATION_PAREN_POWER_WIDTH_SCALE` from `0.913` to `0.910` moves
  sourceIndex `18` visible ink width residual from v126 `+0.248pt` to
  `+0.198pt` without changing height. The scope remains
  `repeatedEquationParenPower(...)`: the formula must contain `=` and at least
  two closing-fence superscript segments, so single equations such as
  `S=\left(a+b\right)^2=9` stay out of this local compression path.
- v127 on doc 61 first-30 ink versus v126: width average improved from
  `0.103pt` to `0.102pt`, p90 improved from `0.226pt` to `0.214pt`, and max
  improved from `0.248pt` to `0.244pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`.
- v127 tightens the regression test threshold for repeated parenthesized-power
  equations from `leftRightEquation * 1.46` to `* 1.455`, still relative rather
  than exact so future small calibration is possible while avoiding a return to
  the wider v124/v126 state.
- v127 structural and physical gates still passed on doc 61: `520` valid
  MathType OLE, `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE,
  zero bitmap WMF, zero WMF LaTeX leaks, zero review-required suspicious WMF
  text, ordinal target WMF width/height `520/520` within 1%, and ordinal
  target shape width/height `520/520` within 1%.
- v127 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\xsc测试集完整重建_61.docx`
- v127 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\leak-scan-61.json`
- v127 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\wmf-report-61.json`
- v127 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\formula-preview-ink-source-vs-v127-first30.json`
- v127 aligned physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\pair-metrics-aligned\4-3-4 蝴蝶模型_summary.json`
- Added `scripts/measure_wmf_formula_glyphs.py` as the next diagnostic wheel.
  It reads a DOCX, maps each formula object to its WMF preview, parses
  placeable header, `SetWindowExt`, `CreateFontIndirect`, `SelectObject`,
  `ExtTextOut`, `TextOut`, and `Polyline`, and can optionally render the
  preview media with ImageMagick to measure real ink bbox. Keep the distinction
  explicit: `ExtTextOut` `dx[]` is record-level advance/layout geometry, not
  proof that Word-visible ink changed.
- The glyph tool reports three layers per formula: outer box metrics
  (`v:shape`, `w:dxaOrig/w:dyaOrig`, WMF placeable/window extents), run metrics
  (`fontFace`, `fontHeightPt`, `fontWidthPt`, `charset`, `xPt`, `baselinePt`,
  `advanceWidthPt`, `rightEdgePt`, and per-character left/right/dx), and optional
  ImageMagick preview-media ink metrics
  (`magickInkLeft/Top/Right/Bottom/Width/HeightPt`). Use it with the existing
  source-vs-generated ink comparison when deciding whether a width change is
  visual or only record-level.
- First v127 glyph run on doc 61 indexes `5,7,9,15,18` shows why this tool is
  needed. For sourceIndex `18`, `recordWidth=109.1pt` and ImageMagick preview
  ink is `107.5pt` wide under fixed `-density 144`; the record runs are
  continuous and show no artificial spacing. For sourceIndex `15`,
  `recordWidth=47.1pt` but ImageMagick preview ink is `43.229pt` wide, so
  further tuning cannot rely on total `dx` alone; font width/height and real
  renderer ink must be checked together.
- v127 glyph outputs:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\wmf-glyph-metrics-5-7-9-15-18.json`
- v127 glyph text summary:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\wmf-glyph-metrics-5-7-9-15-18.txt`
- v127 glyph CSV exports:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\wmf-glyph-runs-5-7-9-15-18.csv`
  and
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v127-repeated-paren-equation-width\wmf-glyph-chars-5-7-9-15-18.csv`
- No-context review caught an important measurement pitfall: generated WMFs use
  about `20` logical units per point, but official/source MathType WMFs can use
  different `SetWindowExt` units. The glyph tool now derives separate X/Y
  `unitsPerPt` from `SetWindowExt` and the WMF placeable physical size when
  available, falling back to `20` only when the physical scale is missing. Do
  not compare source-run advances before checking the reported unit scale.
- Source WMF record metrics can still exceed the visible box even after unit
  scaling because official MathType previews may contain overlapping or
  alternate-font text runs. Treat source WMF run metrics as diagnostic evidence,
  not as an acceptance target by themselves; use ink/page render metrics for the
  final visual judgment.
- The glyph tool now keeps a real WMF object table: font, pen, brush, palette,
  and region creation all occupy object indexes; `DeleteObject` frees indexes;
  `SelectObject` updates the active font only when the selected object is a
  font. This avoids shifting font handles when official/source WMFs create pens
  or brushes before fonts.
- `ExtTextOut` parsing now handles optional `ETO_OPAQUE` / `ETO_CLIPPED`
  rectangles before reading text and `dx[]`. Without this, clipped official WMF
  text records can shift the parsed raw text and every per-character advance.
- ImageMagick WMF ink measurement is explicitly named `magickInk*` and rendered
  with fixed `-density 144 -background white -alpha remove -alpha off`. This is
  an automated preview-media diagnostic, not Word/GDI truth; use Word/page ink
  checks before treating it as final visual evidence.
- `scripts/measure_wmf_formula_glyphs.py --with-ink` depends on both ImageMagick
  and Pillow. The script reports `pillowAvailable` and per-formula
  `magickInkError`; use `--require-magick-ink` when a run must fail instead of
  silently producing record-only diagnostics.
- v128 targets only standalone uppercase `S_{1}` and `S_{3}` subscript labels.
  The v127 residuals had opposite signs for these two labels, so the renderer
  now keeps the generic uppercase-subscript scale intact and applies exact local
  width scales only for standalone `S_{1}` (`0.790`) and standalone `S_{3}`
  (`0.842`). Do not broaden this into relation formulas such as
  `S_{1}=a^{2}=1` or `S_{3}=4=b^{2}`.
- v128 on doc 61 first-30 visible ink versus v127: width average improved from
  `0.102pt` to `0.090pt`, p90 improved from `0.214pt` to `0.190pt`, and max
  improved from `0.244pt` to `0.231pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`. The previous standalone `S_{1}` and
  `S_{3}` worst residuals exited the worst-width list.
- v128 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap WMF,
  zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v128 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
  Do not use raw ordinal source comparison here: the source has extra objects,
  so ordinal metrics can show false huge ratios even when target metrics are
  correct.
- `rebuild/compare_formula_preview_ink.py --indices` is zero-based DOCX object
  indexing. To compare sourceIndex `1..30`, pass `--indices 0,1,...,29`; passing
  `1..30` compares sourceIndex `2..31` and can create a false regression around
  sourceIndex `31`.
- v128 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\xsc测试集完整重建_61.docx`
- v128 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\leak-scan-61.json`
- v128 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\wmf-report-61.json`
- v128 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\formula-preview-ink-source-vs-v128-first30.json`
- v128 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- Added `scripts/compare_wmf_formula_glyphs.py` to compare two JSON reports
  from `scripts/measure_wmf_formula_glyphs.py`. It aligns by 1-based
  `objectIndex`, writes formula/run diff CSVs, and ranks worst
  record-width, ImageMagick ink-width, and ink-height deltas. Keep this as a
  diagnostic attribution tool: it explains whether a renderer change moved
  WMF records, preview-media ink, baselines, or run placement.
- The glyph inspector `--indices` argument is 1-based object indexing because
  it filters `box.index + 1`. This is intentionally different from
  `rebuild/compare_formula_preview_ink.py --indices`, which is zero-based.
  Always check the script help/source before reusing an index list across these
  tools.
- A v127-to-v128 glyph diff on doc 61 indexes `5,7,9,15,18` confirms the
  intended narrow effect: standalone `S_{1}` record width changed by `-0.250pt`,
  standalone `S_{3}` record width changed by `+0.300pt`, and `S_{2}=2`,
  `S_{1}=a^{2}=1`, plus the repeated parenthesized-power equation stayed at
  `0.000pt` record-width delta. The same diff reports `0.000pt`
  `magickInkWidthDeltaPt` for `S_{1}`/`S_{3}`, which is another reminder that
  ImageMagick preview-media ink can miss small real-layout changes; use
  Word/page ink checks for final visual acceptance.
- v128 glyph metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\wmf-glyph-metrics-5-7-9-15-18.json`
- v128 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v128-standalone-s1-s3-width\wmf-glyph-diff-v127-v128-5-7-9-15-18.json`
- v129 targets only exact `S_{2}=2` after v128. Lowering
  `SHORT_S2_EQUALS_TWO_WIDTH_SCALE` from `0.985` to `0.976` reduced doc 61
  sourceIndex `5` Word/page ink width residual from v128 `+0.231pt` to v129
  `+0.181pt`, while v128-to-v129 glyph diff across worst-first30 indexes shows
  only objectIndex `5` moved at record level (`-0.100pt` record width); sampled
  `S_{1}=a^{2}=1`, repeated parenthesized-power, standalone `S`, `b^{2}`,
  `b=2`, `CD`, `25`, `35`, and triangle-chain relations stayed at `0.000pt`
  record-width delta.
- v129 on doc 61 first-30 visible ink versus v128: width average improved from
  `0.090pt` to `0.088pt`, p90 improved from `0.190pt` to `0.182pt`, and max
  improved from `0.231pt` to `0.213pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`.
- v129 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap WMF,
  zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v129 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
- v129 also reconfirmed that official/source MathType WMF record structure is
  not a calibration target by itself: source WMFs often use about `32` logical
  units per point and overlapping `x=0` runs, so source-vs-generated glyph diff
  is useful for attribution but Word/page ink remains the visual acceptance
  evidence.
- v129 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v129-exact-s2-equals-two-width\xsc测试集完整重建_61.docx`
- v129 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v129-exact-s2-equals-two-width\leak-scan-61.json`
- v129 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v129-exact-s2-equals-two-width\wmf-report-61.json`
- v129 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v129-exact-s2-equals-two-width\formula-preview-ink-source-vs-v129-first30.json`
- v129 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v129-exact-s2-equals-two-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- v129 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v129-exact-s2-equals-two-width\wmf-glyph-diff-v128-v129-worst-first30.json`
- v130 targets only exact `S_{1}=a^{2}=1` after v129. Increasing
  `SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION` from `1.107` to
  `1.113` reduced doc 61 sourceIndex `15` Word/page ink width residual from
  v129 `-0.213pt` to v130 `-0.163pt`. Do not read this as a generic script
  relation scale: it is intentionally narrower than standalone `S_{1}` and
  unrelated to long parenthesized-power equations.
- v130 on doc 61 first-30 visible ink versus v129: width average improved from
  `0.088pt` to `0.087pt`, p90 improved from `0.182pt` to `0.170pt`, and max
  improved from `0.213pt` to `0.198pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`.
- v130 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap WMF,
  zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v130 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
- v129-to-v130 glyph diff across worst-first30 indexes confirms the intended
  narrow effect: only objectIndex `15` moved materially, with record width
  `+0.100pt` and `magickInkWidthDeltaPt=+0.502pt`; sampled `S_{2}=2`,
  repeated parenthesized-power, standalone `S`, `b^{2}`, `b=2`, `CD`, `25`,
  `35`, and triangle-chain relations stayed at `0.000pt` record-width delta.
  Word/page ink remains the visual acceptance evidence because ImageMagick ink
  can move more than the record advance.
- v130 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v130-exact-s1-equation-width\xsc测试集完整重建_61.docx`
- v130 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v130-exact-s1-equation-width\leak-scan-61.json`
- v130 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v130-exact-s1-equation-width\wmf-report-61.json`
- v130 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v130-exact-s1-equation-width\formula-preview-ink-source-vs-v130-first30.json`
- v130 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v130-exact-s1-equation-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- v130 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v130-exact-s1-equation-width\wmf-glyph-diff-v129-v130-worst-first30.json`
- v131 continues the repeated parenthesized-power equation tuning after v130.
  Lowering `EQUATION_PAREN_POWER_WIDTH_SCALE` from `0.910` to `0.908` reduced
  doc 61 sourceIndex `18`
  `S=\left(a+b\right)^{2}=\left(1+2\right)^{2}=9` Word/page ink width residual
  from v130 `+0.198pt` to v131 `+0.148pt`; height stayed `-0.111pt`.
  The scope is still `repeatedEquationParenPower(...)`, so single equations
  such as `S=\left(a+b\right)^2=9` and standalone `\left(a+b\right)^2` stay
  out of this local compression path.
- v131 on doc 61 first-30 visible ink versus v130: width average improved from
  `0.087pt` to `0.085pt`, p90 improved from `0.170pt` to `0.164pt`, and max
  improved from `0.198pt` to `0.189pt`. Height average, p90, and max stayed
  `0.075pt`, `0.112pt`, and `0.127pt`. After this, the worst width residual
  moved to sourceIndex `3` standalone `S` at `+0.189pt`.
- v131 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap WMF,
  zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v131 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
- v130-to-v131 glyph diff across worst-first30 indexes confirms the intended
  narrow effect: only objectIndex `18` moved materially, with record width
  `-0.050pt`; sampled `S_{1}=a^{2}=1`, `S_{2}=2`, standalone `S`,
  standalone `\left(a+b\right)^{2}`, `b^{2}`, `b=2`, `CD`, `25`, `35`, and
  triangle-chain relations stayed at `0.000pt` record-width delta. ImageMagick
  ink did not move, while Word/page ink improved, so keep using Word/page ink as
  visual acceptance evidence.
- v131 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v131-repeated-paren-equation-width\xsc测试集完整重建_61.docx`
- v131 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v131-repeated-paren-equation-width\leak-scan-61.json`
- v131 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v131-repeated-paren-equation-width\wmf-report-61.json`
- v131 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v131-repeated-paren-equation-width\formula-preview-ink-source-vs-v131-first30.json`
- v131 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v131-repeated-paren-equation-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- v131 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v131-repeated-paren-equation-width\wmf-glyph-diff-v130-v131-worst-first30.json`
- v131 review lesson: when a calibration changes a record coordinate by only a
  few logical units, regression tests must use a threshold tight enough to fail
  on the previous constant. A broad relative assertion such as comparing a
  repeated parenthesized-power equation to a shorter single equation can prove
  compactness but still allow the exact `0.908` backslide to `0.910`; the test
  now caps the repeated equation `maxTextRightCoordinate` below `2189` logical
  units, which is above the v131 value and below the v130/backslide value.
- v132 revisits standalone `S` after it became the first-30 worst width
  residual again. Lowering `STANDALONE_SINGLE_S_WIDTH_SCALE` from `1.10` to
  `1.065` reduced doc 61 sourceIndex `3` visible Word/page ink from v131
  `+0.189pt` to below the worst-width list; first-30 width average improved
  from `0.085pt` to `0.079pt`, p90 from `0.164pt` to `0.153pt`, and max from
  `0.189pt` to `0.181pt`. Height metrics stayed unchanged.
- v132 confirms the old v83/v84 lesson still holds: ImageMagick ink did not
  move for standalone `S`, while Word/page ink did, so do not use
  `magickInkWidthPt` as the final signal for tiny single-glyph calibration.
  The v131-to-v132 glyph diff showed only objectIndex `3` moving materially:
  record width `-0.350pt`, sampled `S_{1}` relation formulas, `S_{2}=2`,
  parenthesized powers, `b^{2}`, `b=2`, `CD`, `25`, `35`, and triangle-chain
  relations stayed at `0.000pt` record-width delta.
- v132 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap WMF,
  zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v132 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
- v132 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v132-standalone-s-width\xsc测试集完整重建_61.docx`
- v132 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v132-standalone-s-width\leak-scan-61.json`
- v132 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v132-standalone-s-width\wmf-report-61.json`
- v132 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v132-standalone-s-width\formula-preview-ink-source-vs-v132-first30.json`
- v132 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v132-standalone-s-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- v132 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v132-standalone-s-width\wmf-glyph-diff-v131-v132-worst-first30.json`
- v133 retunes exact `S_{2}=2` after it remained the doc 61 first-30 worst
  width residual. Lowering `SHORT_S2_EQUALS_TWO_WIDTH_SCALE` from `0.976` to
  `0.969` reduced sourceIndex `5` visible Word/page ink width residual from
  v132 `+0.181pt` to v133 `+0.131pt`. First-30 width average improved from
  `0.079pt` to `0.077pt`, p90 from `0.153pt` to `0.150pt`, and max from
  `0.181pt` to `0.169pt`. Height average, p90, and max stayed `0.075pt`,
  `0.112pt`, and `0.127pt`.
- v133 confirms that exact short-script equation tuning should stay local:
  the v132-to-v133 glyph diff showed only objectIndex `5` moving materially
  with record width `-0.100pt`; sampled `S_{1}` relation formulas,
  standalone parenthesized powers, `b^{2}`, `b=2`, repeated parenthesized
  equations, `CD`, `25`, `35`, and triangle-chain relations stayed at
  `0.000pt` record-width delta. ImageMagick ink again stayed unchanged while
  Word/page ink improved, so keep treating `magickInk*` as a diagnostic layer,
  not the visual acceptance signal.
- v133 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap
  WMF, zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v133 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
- v133 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v133-exact-s2-equation-width\xsc测试集完整重建_61.docx`
- v133 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v133-exact-s2-equation-width\leak-scan-61.json`
- v133 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v133-exact-s2-equation-width\wmf-report-61.json`
- v133 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v133-exact-s2-equation-width\formula-preview-ink-source-vs-v133-first30.json`
- v133 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v133-exact-s2-equation-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- v133 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v133-exact-s2-equation-width\wmf-glyph-diff-v132-v133-worst-first30.json`
- v134 retunes standalone `b^{2}` after it became the doc 61 first-30 worst
  width residual. Lowering `STANDALONE_B_SUPERSCRIPT_WIDTH_SCALE` from `0.795`
  to `0.779` moved sourceIndex `10` out of the worst-width list; first-30
  width average improved from v133 `0.077pt` to v134 `0.074pt`, median from
  `0.072pt` to `0.069pt`, and max from `0.169pt` to `0.163pt`. Height
  average, p90, and max stayed `0.075pt`, `0.112pt`, and `0.127pt`.
- v134 keeps standalone `b^{2}` tuning isolated from relation formulas:
  the v133-to-v134 glyph diff showed only objectIndex `10` moving materially
  with record width `-0.150pt`; sampled `S_{1}` relation formulas,
  standalone parenthesized powers, exact `S_{2}=2`, `b=2`, repeated
  parenthesized equations, `CD`, `BD`, `25`, `35`, and triangle-chain
  relations stayed at `0.000pt` record-width delta. ImageMagick ink again
  stayed unchanged while Word/page ink improved, reinforcing that
  `magickInk*` is diagnostic rather than the acceptance signal.
- v134 structural gates on doc 61 still passed: `520` valid MathType OLE,
  `520` vector WMF, zero visible LaTeX leaks, zero invalid OLE, zero bitmap
  WMF, zero WMF LaTeX leaks, and zero review-required suspicious WMF text.
- v134 target physical metrics on doc 61 still passed when aligned with
  `full-61.request.json`: target WMF width `520/520` within 1% with max error
  `0.178%`, target WMF height `520/520` within 1% with max error `0.177%`,
  and target shape width/height both `520/520` within 1% with max error `0`.
- v134 sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\xsc测试集完整重建_61.docx`
- v134 leak scan:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\leak-scan-61.json`
- v134 WMF report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\wmf-report-61.json`
- v134 first-30 ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\formula-preview-ink-source-vs-v134-first30.json`
- v134 target physical metrics:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\pair-metrics-target\4-3-4 蝴蝶模型_summary.json`
- v134 glyph diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\wmf-glyph-diff-v133-v134-worst-first30.json`
- A v135 micro-step that changed only
  `SCRIPT_RELATION_EXACT_S1_EQUATION_WIDTH_COMPENSATION` from `1.113` to
  `1.116` was intentionally not kept: doc 61 first-30 Word/page ink stayed
  unchanged, with sourceIndex `15` still at width delta `-0.163pt`. Do not
  commit record-only or constant-only changes unless the Word/page ink report
  or another acceptance-level visual metric moves in the right direction.
- `scripts/measure_wmf_formula_glyphs.py` now emits byte-level `ExtTextOut`
  dx rows via `--out-bytes-csv` in addition to decoded character rows. This is
  necessary for Symbol/CJK/DBCS runs: objectIndex `30` in the v134 doc 61
  sample decodes the triangle marker through `SimSun` as `¡÷`, so decoded
  characters are useful for human hints but raw byte dx is the safer
  calibration evidence.
- The glyph inspector has a `--self-test` path that constructs a tiny WMF with
  `CreatePenIndirect` before `CreateFontIndirect`, selects the font at object
  index `1`, and uses `ExtTextOut` with `ETO_CLIPPED`. Keep this self-test
  passing before relying on the tool for glyph-size diagnosis.
- v134 byte-level glyph sample:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\wmf-glyph-bytes-byte-sample.csv`
- Source MathType WMF record geometry is not structurally comparable to the
  generated vector WMF for many formulas. In the doc 61 first-30 worst-width
  sample, source objectIndex `15` (`S_{1}=a^{2}=1`) has `5` runs while the
  generated WMF has `6`; objectIndex `30` has `6` source runs versus `14`
  generated runs. Large source/generated `recordWidthDeltaPt` values in this
  situation are diagnostic only, not a direct tuning target.
- `scripts/compare_wmf_formula_glyphs.py` now reports
  `runStructureComparable`, byte counts, advance-sum deltas, and
  `recordWidthTrust`. Treat `low_run_structure_mismatch` as a stop sign for
  renderer constants: use Word/page ink or another acceptance-level visual
  metric before changing widths.
- `runStructureComparable` must require identical run keys
  `(runIndex, fontFace, text)`, not just equal run count and equal concatenated
  text. Same joined text with different per-run fonts is still structurally
  different and should stay low-trust.
- v134 source/generated glyph structural diff:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\wmf-glyph-diff-source-generated-worst-first30-v3.txt`
- A doc 61 all-formula Word/page ink run over `520` formulas can exceed a
  five-minute command timeout because it renders both reference and generated
  preview media. `rebuild/compare_formula_preview_ink.py` now supports
  `--save-reference-rows`, `--save-generated-rows`, `--load-reference-rows`,
  `--load-generated-rows`, and `--progress-every` so full visual checks can be
  split into reusable measurement and fast comparison phases.
- The first cache-mode verification used doc 61 first 30 formulas and matched
  the previous v134 metrics exactly: width avg `0.074pt`, p90 `0.150pt`, max
  `0.163pt`; height avg `0.075pt`, p90 `0.112pt`, max `0.127pt`.
- Cache load mode must not require ImageMagick when both reference and
  generated rows are loaded, and loaded rows must still honor `--indices`,
  `--source-indices`, and `--max-items`. A regression check used a missing
  `--magick` path plus `--source-indices 1,2,3` and produced `paired_count=3`.
- v134 first-30 cached Word/page ink rows:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\formula-preview-ink-reference-v134-first30-rows.json`
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\formula-preview-ink-generated-v134-first30-rows.json`
- `rebuild/compare_formula_preview_ink.py --measure-only` is needed for
  single-side full-document measurement. Using `--source-indices 1` as a
  bootstrap filter only measures one formula on both sides, which is not a
  valid way to pre-cache the reference document.
- Long single-side measurement must checkpoint as it runs. The first reference
  full-doc attempt timed out after five minutes; after adding
  `--checkpoint-every` and `--resume-reference-rows`, the partial cache
  survived at `401/520` rows and the second run completed the remaining
  `119` rows.
- Resume caches must still be filtered by the current `--indices`,
  `--source-indices`, and `--max-items`. A regression check resumed from a
  full `520`-row reference cache with `--source-indices 1,2,3` and saved only
  rows `0,1,2`. `--measure-only` intentionally rejects `--load-*-rows`; use
  `--resume-*-rows` when the command is expected to write or extend a cache.
- v134 doc 61 full Word/page ink cache is now available:
  reference rows `520/520`,
  generated rows `520/520`, and loaded comparison `paired_count=520`.
  The raw all-index statistics are polluted by DOCX-local object-index
  misalignment in later pages, so do not use the raw `avg=35.304pt` width
  error as a renderer metric.
- The full-doc ink comparison now reports `alignment_suspicious` using context
  similarity and extreme width/height scale. In v134 doc 61,
  `alignment_suspicious_count=212`, leaving `aligned_paired_count=308`. This
  means the next large improvement is not another first-30 width constant; it
  is better source/generated formula alignment and then long-formula family
  calibration.
- v134 full-doc guarded ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\formula-preview-ink-source-vs-v134-all-alignment-guard-v2.txt`
- Full-doc preview ink comparison must not use docx2tex report ordinal as the
  source DOCX object index. In doc 61, `source-report-61.json` has `520`
  converted equations but `docObjectIndex` spans `1..527` with `7` gaps. The
  correct source-row mapping is `docObjectIndex - 1`; using the report ordinal
  created fake huge deltas and a misleading full-doc average. The fixed
  LaTeX-aligned cached run produced `488` usable pairs, width avg `1.057pt`,
  height avg `0.756pt`, and only `12` context/scale suspicious rows, with an
  explicit warning that coverage is not full-document acceptance.
- Latest LaTeX-aligned preview ink report:
  `J:\latextomathtype\analysis\unattended-runs\20260615-tight61-72\docx\61-v134-standalone-b-squared-width\formula-preview-ink-source-vs-v134-all-latex-aligned.txt`
- Exact LaTeX keys miss formulas when the source docx2tex report keeps style
  wrappers such as `\mathrm{三角形...}` or `\mathrm{cm}` while the generated
  request contains bare Chinese text or bare `cm`. A safe fallback is to use
  `ordinal_latex_key` only for keys that are unique on both remaining sides,
  mark the row `pair_method=ordinal_key`, and warn that those pairs are
  diagnostic only. Do not pair duplicate collapsed keys by encounter order.
- With exact plus unique ordinal-key fallback, doc 61 cached preview ink
  comparison reaches `513/520` usable pairs: `495` exact plus `25` fallback,
  with `7` source docObjectIndex gaps intentionally left unpaired. Width avg is
  `1.084pt`, height avg is `0.863pt`, and `alignment_suspicious_count=14`.
  The report still warns that generated request order is assumed to match DOCX
  preview order because there is no explicit renderer object map yet.
- Generated DOCX previews now carry a non-leaking formula trace id on
  `<v:imagedata o:title="pwf:<sha256-prefix>">`. This gives future
  source/request/generated alignment an explicit renderer object map without
  writing raw LaTeX into the Word XML or WMF preview.
- The Java and Python trace hash inputs must stay byte-for-byte compatible:
  mirror the request-to-DOCX path, strip only leading `\pwmetrics{...}` and
  `\pwstyle{...}`, normalize NBSP to ordinary space, then trim/collapse only
  ASCII whitespace `[ \t\n\v\f\r]`. Do not hash `normalize_latex_key(...)`,
  because canonical LaTeX keys intentionally rewrite style wrappers and would
  diverge from the DOCX metadata written by `MathTypeEmbedder`.
- A contract fixture for trace ids is now locked in Java:
  `\pwmetrics{1,2,3,4}\t\pwstyle{trace}\n a\u00a0 +\t b ` must normalize to
  `a + b` through the actual DOCX builder path and hash to
  `pwf:cb23f6635a581786`. Keep the Python helper in
  `rebuild/compare_formula_preview_ink.py` matching that exact request-to-DOCX
  whitespace protocol.
- `scripts/compare_docx_pair_metrics.py` now preserves request `rawLatex`, and
  `rebuild/extract_formula_boxes.py` exposes `image_title`, so
  `rebuild/compare_formula_preview_ink.py` can match generated preview rows by
  trace id when the generated DOCX was produced by this newer build.
- Trace matching is intentionally strict: only a trace id that is unique in the
  request and unique in the generated DOCX row set is trusted. Duplicate or
  otherwise unmatched trace rows are marked `generated_trace_status=unmatched`
  and `alignment_suspicious` instead of silently falling back to an authoritative
  visual metric.
- Old v134 doc 61 cached rows have `generated_trace_count=0` because they were
  generated before preview trace ids existed. Regenerate and remeasure the DOCX
  before using trace ids as alignment evidence on the full test set.
- A small generated trace candidate proved the Java/Python metadata path:
  `target/reference-roundtrip/trace-candidate.docx` extracted one object with
  `image_title=pwf:ec640c8e884a30bf`.
- `scripts/measure_wmf_formula_glyphs.py` must not independently scrape every
  JSON string with a loose `$...$` regex. For generated DOCX files that carry
  `pwf:` preview titles, it now trusts only trace ids that are unique in both
  the generated DOCX rows and the request formulas. Duplicate or unmatched trace
  rows are labeled `trace_duplicate` / `trace_unmatched` instead of silently
  falling back to ordinal formula labels. Without trace titles it falls back to
  request order only as diagnostic context and labels rows with
  `formulaMatchMethod=ordinal_fallback`.
  Reports keep `formula` as the stripped display body for compatibility, and
  add `rawFormula`, `formulaKey`, `formulaTraceId`, and `formulaMatchMethod`.
  Standard generated request JSON may not preserve `_mathOrder`, because
  `make_full_batch10_requests.py` strips internal request fields before writing.
- Trace-backed glyph labeling was verified on
  `target/reference-roundtrip/trace-candidate.docx` with a matching temporary
  request: `scripts/measure_wmf_formula_glyphs.py` reported
  `match=trace_unique trace=pwf:ec640c8e884a30bf` for the first object.
- Content-only trace ids are not enough for real test-set alignment. A doc 61
  trace run had `520/520` preview titles but only `271/520` unique generated
  trace matches because repeated formulas shared the same `pwf:<hash>`. New
  generated DOCX previews use `pwf:<documentOrdinal>-<sha256-prefix>`, keeping
  the normalized content hash non-leaking while making duplicate formulas
  uniquely traceable.
- The trace ordinal must be document-scoped, not merely a long-lived
  `MathTypeEmbedder` instance counter. `PaperExportService` and several tests
  reuse `DocxBuilder` instances, so `DocxBuilder.build(...)` and
  `LayoutDocxBuilder.build(...)` must reset the embedder's document formula
  counter before writing a new DOCX. Regression tests should cover duplicate
  formulas in one document and repeated builds through the same builder.
- Python trace helpers must index both the new `pwf:<ordinal>-<hash>` format
  and legacy `pwf:<hash>` titles. Legacy duplicates remain diagnostic only:
  `scripts/measure_wmf_formula_glyphs.py` labels them `trace_duplicate` /
  `trace_unmatched` but falls back to request ordinal for formula text context
  instead of treating the labels as authoritative matches.
- After ordinal trace ids, regenerated doc 61 at
  `J:\latextomathtype\analysis\trace-runs\20260616-doc61-ordinal-trace\docx\xsc测试集完整重建_61.docx`
  has `520` MathType OLE objects, `520` WMF previews, `0` LaTeX leaks, and
  `520/520` unique ordinal trace titles. Cached full-doc ink comparison now
  reports `generated_trace_count=520`, `generated_trace_match_count=520`,
  `usable_latex_pair_count=513`, `aligned_paired_count=499`, and
  `alignment_suspicious_count=14`; width/height deltas are not solved yet
  (`aligned_ink_width_abs_delta_pt avg=1.072pt`,
  `aligned_ink_height_abs_delta_pt avg=0.829pt`), so the next renderer work
  should target the remaining height-heavy and long-formula families rather
  than more alignment plumbing.
- Compact inline fractions were height-heavy because their numerator and
  denominator slots were rendered as 8pt script text inside a compact layout.
  The better mechanism is to keep compact x placement and compressed `dx`
  advances, but render the fraction-slot text at readable 12pt height and
  widen the numerator/denominator baseline spacing. On doc 61 worst-height
  samples, generated `magickInkHeightPt` moved from about `12.884pt` to
  `19.821pt` for simple inline fractions and from `15.5pt` to `23.5pt` for
  wide CJK ratio fractions; old residuals around `-9pt` dropped to roughly
  `-1.1..-3.1pt` on the sampled worst rows.
- WMF preview geometry edits must bump `LaTeXImageRenderer.CACHE_VERSION`
  before regenerating DOCX evidence. A same-version rebuild reused old cached
  WMF media and falsely showed no change even after renderer code changed.
  Treat unchanged measured font heights after a renderer edit as a cache-key
  suspect before tuning more constants.
- Do not compact nested fractions with the simple inline-fraction slot layout.
  `count <= 2` alone still allows constructs such as
  `CO=\frac{\frac{1}{2}}{3}`; those need non-compact vertical geometry to
  avoid overlapping slots.
- When promoting compact inline fraction slots from 8pt script-sized text to
  readable text, do not blindly force every child run to non-script. Real script
  children inside a numerator/denominator, such as `\frac{S_{1}}{ABC}`, must
  remain script-sized while the main slot text stays readable.
- Short linear formulas in 12pt-high boxes, such as `AE=EF=FB`,
  `BE=DF=1`, and `12\times 2=24`, were visually too small after the generic
  simple-linear font Y scale: generated WMF used about `9.75pt` text while the
  reference MathType WMF samples used about `10.5pt`. A scoped short-linear
  equation classifier can raise font height without adding spaces or changing
  fraction/script layout. On doc 61 v137, sample glyph metrics moved these rows
  to `10.5pt` text and `7.0pt` Magick ink height; full-doc width average only
  improved slightly (`aligned_ink_width_abs_delta_pt avg 1.072pt -> 1.066pt`),
  so this is a readability fix, not the final width solution.
- The short-linear readable-height classifier must require an equality,
  multiplication, division, or dot-product operator. Do not let plain short
  addition/subtraction expressions such as `A+B+C`, `12+34`, or `AB-CD` match
  this calibration; those are near-neighbor negative cases for the regression
  test.
- `scripts/run_xsc_unattended_acceptance.ps1` may call system `python` for
  compare steps, which can fail on Chinese source paths if that Python/runtime
  decodes arguments incorrectly. A failed compare after `RUN mvn-*` can still
  leave a valid generated DOCX; verify with the bundled Codex Python and the
  known source path before discarding the run.
- Short-linear width compensation must be scoped by rendered box height, not
  only by LaTeX text. A broad `1.08` width scale for every short linear
  equation improved target 12pt samples such as `AE=EF=FB` and `BE=DF=1`, but
  widened many already-close 12.75pt formulas and made doc 61 full comparison
  worse (`aligned_ink_width_abs_delta_pt avg 1.066pt -> 1.076pt`). Gating the
  same width compensation to `heightPt <= 12.1` preserved the target sample
  gains and improved the doc 61 aligned width average to `1.043pt`; keep future
  local calibrations tied to the physical box family that actually needs them.
- Non-compact fraction rows in the 27.75pt-high family were height-light even
  when their glyph font was already readable. The fix was vertical geometry,
  not font size: move the numerator slot up, denominator slot down, fraction
  bar with it, and keep compact inline fractions on their separate geometry.
  On doc 61 v140, representative rows such as
  `S_{\Delta DEO}=\frac{2}{3}...` moved from about `20.812pt` Magick ink height
  to `21.804pt`, while compact inline rows like `48\times\frac{1}{4}=12`
  stayed unchanged. Full comparison improved height without widening formulas:
  `aligned_ink_height_abs_delta_pt avg 0.366pt -> 0.318pt`, max
  `4.773pt -> 3.484pt`, with width avg unchanged at `1.043pt`.
- Fraction formulas return from `layoutFractions(...)` before the generic
  `layoutScripts(...)` branch, so `scriptRelationWidthScale(...)` cannot fix
  fraction-heavy area chains. For doc 61, the remaining width-heavy cluster was
  `\frac` + script/area formulas such as
  `S_{\bigtriangleup ENF}=\frac{9}{...}...`; applying a scoped `0.98` X scale
  to only non-nested fraction layouts with script/area tokens reduced these
  rows while leaving compact inline fractions (`48\times\frac{1}{4}=12`) and
  plain algebraic fractions (`EF=\frac{1}{2}(a+2a)=...`) unchanged. v142 moved
  `aligned_ink_width_abs_delta_pt avg 1.043pt -> 0.981pt` with height avg
  unchanged at `0.318pt`. Do not treat every long `S_` fraction as an area
  formula; a no-context review caught that overmatch. Require explicit area
  semantics such as `\bigtriangleup`, `\Delta`, `\Updelta`, `梯形`, or `三角形`
  so ordinary sequences like `S_{n}=\frac{...}{...}+b^{2}` remain unscaled.
- For long triangle/script ratio chains such as
  `S_{\bigtriangleup GEF}\colon ...=4\colon9\colon6\colon6`, check glyph
  metrics before changing `dx`: source and generated boxes were both
  `368.25x17.25pt`, and generated record width was not smaller, but generated
  main text used about `9.8pt` while MathType reference used about `10.5pt`.
  A scoped long-triangle relation font-height branch is the right lever; do not
  add spaces or only widen `ExtTextOut dx`. v143 improved doc 61 width max
  `5.517pt -> 4.766pt` and aligned width avg `0.981pt -> 0.976pt`, with a small
  height avg tradeoff `0.318pt -> 0.320pt`.
- Compact inline fractions were still a shared width/height family after the
  readable-slot change: samples such as `48\times\frac{1}{4}=12`,
  `EF=\frac{1}{2}(a+2a)=\frac{3}{2}a`, and
  `S_{梯形EFCD}=\frac{7}{12}S` were slightly too wide and about `2pt` too
  short in Magick ink. A small overall compact-inline X scale plus a 1pt wider
  numerator/denominator slot spread moved sample ink heights to the reference
  `21.804pt` family and improved doc 61 v144 aligned averages
  (`width 0.976pt -> 0.919pt`, `height 0.320pt -> 0.233pt`). When reading the
  full report, keep suspicious ordinal-key outliers separate; v144 introduced a
  large suspicious width row, but the aligned max remained the old long-triangle
  row and the aligned averages improved.
- Fraction-family width scales must be mutually exclusive. A no-context review
  caught that compact inline fractions and area-chain fractions can both match a
  short `S_{\bigtriangleup ...}=\frac{...}{...}S_{梯形...}` formula, which would
  multiply `0.98 * 0.965` and over-compress. Apply the area-chain scale first
  and use the compact-inline scale only when the area-chain scale did not match;
  keep a regression test for that intersection.
- MathType-like WMF formula variables need font-style parity before more spacing
  tweaks: Latin variable runs such as `S`, `AOB`, `AB`, `a`, and `b` should use
  Times New Roman Italic, while digits, operators, Symbol glyphs, and CJK text
  stay non-italic. This is a scoped font-selection change, not a reason to add
  spaces or inflate `ExtTextOut` dx. In v145, doc 61 kept 520 valid MathType
  OLE/WMF objects and 0 LaTeX leaks; aligned width average improved
  `0.919pt -> 0.879pt`, aligned width max improved `4.766pt -> 4.266pt`, and
  height average stayed effectively flat (`0.233pt -> 0.234pt`).
- When adding WMF preview font objects, keep object indices explicit and covered
  by tests. Adding three italic Times New Roman fonts after the original
  ANSI/Symbol/CJK 9-font table required moving `PEN_OBJECT_INDEX` from `9` to
  `12`; otherwise later polyline/fraction-bar drawing can select a font object
  instead of the pen. Tests should assert both italic LOGFONT selection for
  Latin letter records and non-italic selection for numeric-only records.
- A no-context review caught a real style-semantics gap in the v145 italic
  approach: the main layout path still calls `normalizeTextCommands(...)` before
  most tokenization, so wrappers such as `\mathrm{AB}` and `\text{AB}` can be
  flattened before WMF font selection sees them. Do not claim full LaTeX style
  fidelity yet. The next robust fix is to preserve style metadata through
  `TextRun` / `PlacedText` across all layout branches, instead of trying to infer
  roman/text intent from the already-flattened characters.
- Unified formula layout must be introduced as a guarded path, not a broad
  replacement. A no-context review caught that letting every `\sqrt` enter the
  unified parser bypassed the calibrated `layoutSqrt(...)` path and would also
  silently drop optional root indexes such as `\sqrt[3]{8}`. Keep simple roots
  on the old path until the unified parser explicitly emits root indexes.
- Text-heavy CJK ratios should usually be content-split, not forced into one
  MathType fraction object. A full `\frac{三角形ABD的面积}{三角形CBD的面积}` preview
  either becomes too tall or too small after Word object scaling; the readable
  K12-style form is ordinary Chinese text plus small formula objects such as
  `$ABD$`, `$CBD$`, and `$\frac{AO}{CO}$`.
- Preview cache version must be bumped whenever WMF layout or preview box
  metrics change. v146 initially appeared unchanged in Word/PDF because
  `CACHE_VERSION` still pointed at v145 and reused old WMF previews; visual QA
  is not trustworthy until the cache key is invalidated.
- Formula layout needs a single standard glyph anchor before any scale tuning.
  For the Q1-style families, use one base glyph size and derive scripts,
  display symbols, and compact inline fraction slots from that anchor. Do not
  let the font object size and the `ExtTextOut` advance both apply unrelated
  shrink factors to the same script run, or scripts will end up too small even
  when the preview box looks reasonable. Cache version bumps must follow any
  such glyph-scale change so Word does not reuse stale WMF previews.
- Keep the unified box model scoped until each structure is calibrated. A
  no-context review caught that routing text-heavy fractions and complex
  roots through the same unified parser can create a semantic downgrade:
  `\frac{三角形ABD的面积}{...}` becomes a slash fraction, and
  `\sqrt{1+\frac{a_{1}^{2}}{\frac{3}{5}}}` can be hard-compressed by height
  scaling. The current safer rule is: standard glyph anchoring covers common
  script/inline-fraction families, unified box covers controllable nested
  ordinary fractions, while text-heavy fractions and roots stay on their
  established vector paths until source/ink calibration exists for them.
- For the standard glyph ladder, use source WMF `10.5pt` as the base font
  anchor and a visual script ratio (`0.78` in the current build) as a separate
  renderer choice. Do not confuse that visual ratio with the source record
  script ratio (`~0.577`), because the latter does not directly map to Word ink.
  Also keep `ExtTextOut` advance ratios independent from vertical
  `previewScale.y()`, or non-square target boxes can create width drift.
- `\dfrac` and `\cfrac` must be included anywhere the renderer estimates
  fraction physical height, depth, nesting, or preview class. Counting only
  `\frac` makes display fractions look like linear formulas and invites Word
  to squeeze them into a 13pt box.
- The source-WMF parameter table must report coverage, not just numbers. In the
  current 67-470 corpus, `linear`, `script`, `fraction`, and `accent` are
  usable; `nested_fraction` and `text_fraction` are thin; `sqrt`,
  `sqrt_fraction`, `script_fraction`, and `array` are missing. Missing samples
  should appear as `sampleStatus=missing` with a conservative renderer action,
  otherwise future work can mistake "no row" for "no special handling needed".
- `\dfrac` and `\cfrac` support must be end-to-end. Updating only
  `LaTeXImageRenderer` height/depth estimation is not enough: the vector WMF
  layout path (`layoutFractions`, compact/nested detection, `tokenizeFlat`
  guards, and inline-context scanning) must also use the shared fraction-command
  scanner. A regression should assert these commands produce vector fraction
  bars and no DIB fallback.
- Fraction-command scanners must require a LaTeX command boundary after
  `\frac`, `\dfrac`, and `\cfrac`. Plain substring scans can misclassify
  command prefixes such as `\fraction` as real fractions, which then pollutes
  both the source-WMF parameter table and the renderer height/layout route.
- Nested-fraction classification must compare all supported fraction commands
  together, not only same-command nesting. A `\frac{1+\dfrac{a}{b}}{2}` shape
  is still a nested fraction for source statistics and renderer height-family
  selection.
- `\cfrac` must follow the same display-style routing as `\dfrac` in compact
  inline checks. It is easy to include it in the shared fraction scanner but
  forget old explicit exclusions, which squeezes contextual formulas such as
  `x=\cfrac{a}{b}` into compact inline fraction spacing.
- Source structure classification must keep every renderer-supported
  array-like environment out of scalar buckets. Include `alignedat`, `gathered`,
  `pmatrix`, `bmatrix`, and `cases` with `array`, or formulas that contain both
  arrays and fractions/scripts will poison scalar parameter tables.
- Separate numeric WMF parameter samples from global source-report coverage.
  A structure may have no usable WMF rows in the active numbered corpus but
  still have source-report candidates elsewhere. Report `coverageCount`
  separately and do not use coverage-only samples as renderer constants until
  their source DOCX WMF has been joined and inspected.
- After array-like classification was aligned with the renderer, the 67-470
  corpus produced 9 actual `array` WMF rows at 33pt height; these were
  previously hidden inside scalar buckets. `sqrt` and `sqrt_fraction` still had
  zero global source-report coverage in the available reports, so radicals
  remain generated-candidate/old-path territory until real MathType source
  samples are found.
- Do not blindly change `scaleLayoutX` to return a scaled layout width. In this
  renderer the stored layout width is also the later `previewScale` denominator;
  keeping the old width preserves the intended local horizontal compression.
  Changing it to the scaled width cancels or reverses several existing width
  corrections and made standalone script/paren-power tests fail.

### 2026-06-16 Word/MathType reference extraction

- Added `scripts/extract_word_mathtype_reference.ps1` as a first-pass reference extractor. It opens a DOCX through Word COM, records every inline OLE object's ProgID, width, height, range positions, exports a Word PDF, and can open one `Equation.DSMT4` object in MathType for a screenshot.
- The sample run on `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate.docx` produced `analysis/trace-runs/20260616-unified-box-candidate/word-reference/word-mathtype-reference.json`, `word-reference.pdf`, and `mathtype-formula-2.png`; it found `19` MathType OLE objects.
- ImageMagick PDF rendering failed on this machine because Ghostscript (`gswin64c.exe`) was missing. Do not treat that as a Word export failure; the PDF was created successfully by Word COM.
- MathType editor screenshot for formula index `2` showed the formula body is large and natural inside MathType, while Word's inline object box for that formula is `60.75 x 26.25 pt`. Use this split to diagnose whether a defect is in the OLE/MTEF body, WMF preview ink, or Word display box.

### 2026-06-16 Reference extraction UI minimization

- `scripts/extract_word_mathtype_reference.ps1` now defaults to `Word.Application.Visible = $false`, `ScreenUpdating = $false`, and `DisplayAlerts = 0`.
- Added `-NoMathTypeUi` so reference extraction can run with no Word UI and no MathType popup. In that mode the script still records OLE sizes, exports the Word PDF, and skips the spotcheck screenshot.
- MathType popup suppression is only partial when an actual `OLEFormat.Edit()` spotcheck is requested; that path still needs a visible editor window by design.
- Script-slot fractions such as `x^{\frac{1}{2}}` should not reuse the normal
  stacked fraction box. The stacked form placed the numerator baseline near
  `1.2pt` inside a `26.3pt` WMF and Word clipped it into broken dots/short
  strokes. In script context, parse grouped `\frac` with script metadata and use
  a slash fraction (`1/2`) on the same script-size ladder. Bump
  `LaTeXImageRenderer.CACHE_VERSION` after this kind of parser/layout change;
  otherwise DOCX generation can silently reuse stale WMF previews.
- Nested display fractions need their own physical height family. A formula like
  `\frac{1+\frac{a}{b}}{2+\frac{c}{d}}=...` can put the lowest denominator
  baseline around `25.15pt`; if the Word shape remains the single-fraction
  `26.3pt` family, the result looks cramped or clipped. Treat nested fractions
  separately from both script fractions and text-heavy slash ratios.
- Do not tune WMF layout from visual inspection first. Build a source-WMF
  structure corpus: read MathType source WMF records, join each object to the
  source-report LaTeX, classify by structure, then summarize font height,
  script ratio, record advance, baseline span, object height, and margins. The
  first helper for this is `scripts/aggregate_wmf_structure_metrics.py`; use it
  to derive per-structure constants before changing renderer code.
- Source report alignment must be trusted by `docObjectIndex`, not by silent
  ordinal fallback. Some reports have object gaps, and ordinal joins can put
  the wrong LaTeX structure on a WMF object, poisoning per-structure constants.
- The current source-WMF corpus gives stable seed values for common structures:
  main MathType preview font is about `10.5pt`, script record font ratio is
  about `0.577`, dominant linear boxes are `12.75-13pt`, dominant script boxes
  are `16pt`, dominant ordinary fraction boxes are `28pt`, and observed nested
  fraction boxes cluster around `53-60pt`. Treat these as renderer seed
  families, not final visual acceptance.
- Source MathType WMFs in the current corpus use record `0x0626` heavily and
  expose zero parsed `META_POLYLINE` records, so `horizontalLineCenterRatio`
  and line width fields are diagnostic-null for now. Do not infer that fractions
  or radicals have no bars; add `0x0626` parsing or ink-bbox evidence before
  using line geometry as a renderer parameter.
- Source WMF text baselines are often anchored at `0`, so `baselineCenterRatio`
  and crude `topMarginPt` are not safe Word-box placement parameters. Use
  physical height buckets and rendered ink/Word PDF checks for vertical
  placement until a better baseline model exists.
- After deriving source-WMF candidates, the validation loop is: apply a scoped
  renderer change, bump `LaTeXImageRenderer.CACHE_VERSION`, regenerate DOCX,
  inspect WMF records, compare rendered ink/Word PDF, and spot-check Word object
  boxes. The source statistics reduce blind tuning; they do not replace visual
  and physical-size verification.

### 2026-06-17 Word TeXToggle source WMF references

- Word 16 COM can run MathType's Word macro `MTCommand_TeXToggle` in hidden
  mode. Selecting TeX text such as `$\\sqrt{1+\\frac{a_1}{b}}$` and running that
  macro produced a real `Equation.DSMT4` inline OLE with an official MathType
  WMF preview, without opening the MathType editor UI.
- Added `scripts/generate_word_mathtype_tex_reference.ps1` to generate small
  official-reference DOCX files from explicit TeX formulas and a matching
  source-report JSON. This is the preferred path for missing structures such as
  radicals before changing renderer constants.
- `C:\Program Files (x86)\MathType\Office Support\BlankEqn.doc` converts to a
  DOCX with one empty MathType OLE, but the converted preview has zero text runs.
  It is useful only as an installation/object sanity check, not as a structure
  calibration sample.
- The first hidden TeXToggle sample
  `analysis/wmf-structure-metrics/word-mathtype-tex-reference.docx` matched all
  six generated formulas to source-report objects with zero skipped unmapped
  WMFs. Seed physical heights were: symbol/linear `14.25pt`, script `18.75pt`,
  ordinary fraction `30.75pt`, nested fraction `60pt`, sqrt `18pt`, and
  sqrt+fraction `35.25pt`.
- Official MathType source WMFs generated by TeXToggle use a 12pt main font and
  about `0.583` script font ratio in this sample. This differs from some older
  corpus-derived `10.5pt` preview seeds, so prefer structure-specific sample
  provenance over a single global font constant.
- Parameter aggregation must skip WMF objects that have no source-report or no
  `docObjectIndex` mapping. Treating those objects as empty LaTeX silently
  classifies them as `linear` and pollutes standard glyph parameters. Report
  skipped unmapped WMFs separately from missing source DOCX files.
- The standard glyph renderer path should not absorb display or nested fractions
  just because they contain scripts. Only compact inline fractions may share the
  standard glyph ladder; display/nested/script-fraction formulas should route to
  the fraction or unified-box layout families.

### 2026-06-17 TeXToggle expansion and trust checks

- `scripts/generate_word_mathtype_tex_reference.ps1` must verify each macro
  conversion against Word's actual `InlineShapes` collection. A fixed sleep plus
  assumed `docObjectIndex = formula ordinal` is not trustworthy: if
  `MTCommand_TeXToggle` fails or inserts an unexpected object, the source-report
  would look valid and poison the parameter table. The script now requires each
  formula to add exactly one `Equation.DSMT4` object and writes the actual
  `inlineShapeIndex`, `formulaIndex`, `ProgID`, width, and height.
- The expanded hidden-Word TeXToggle sample has 22 formulas and verified
  `22/22` `Equation.DSMT4` objects. It covers symbol/linear, script,
  script-fraction, ordinary fraction, display fraction, nested fraction, sqrt,
  sqrt+fraction, and optional root examples.
- `scripts/aggregate_wmf_structure_metrics.py` now supports
  `--extra-docx-source-report DOCX=SOURCE_REPORT`, so the main XSC source corpus
  can be merged with generated official MathType reference samples. The combined
  report `analysis/wmf-structure-metrics/combined-xsc-tex-toggle-summary.txt`
  had `4404` mapped formulas and `0` skipped unmapped WMFs in this run.
- Combined source seeds should be read by provenance. XSC corpus values dominate
  common structures (`linear`, `script`, `fraction`, `accent`, `array`), while
  TeXToggle fills structures missing from XSC (`script_fraction`, `sqrt`,
  `sqrt_fraction`). Current combined seed buckets include `sqrt` at `18pt`
  (with one optional-root sample at `20.25pt`) and `sqrt_fraction` at `35.25pt`.
- Root-containing formulas must be classified before generic fraction families
  for preview physical height. A formula like `\sqrt{1+\frac{a}{b}}` should use
  the `sqrt_fraction` family around `35.25pt`, not the nested-fraction `53-60pt`
  family, because the vector renderer deliberately keeps roots outside the
  generic unified-box fraction path.
- Do not route true nested display fractions through the compact unified-box
  path until that parser has its own tall fraction metrics. A no-context review
  caught that `layoutUnifiedBox(...)` accepted `hasNestedFraction(text)` while
  `FractionFormulaBox` still used compact inline heights around `13.0+10.8pt`;
  this contradicts the source-WMF nested-fraction family around `53.25/60pt`.
  Keep script-slot fractions such as `x^{\frac{1}{2}}` on the slash/unified
  path, but route nested display fractions through the tall display-fraction
  path.
- If a structure seed height was already set from MathType source statistics,
  avoid applying a later generic visual multiplier to that same height. v156
  removed the blanket `\sqrt` `heightScale *= 1.16` because `sqrt` and
  `sqrt_fraction` were already seeded at `18.0pt` and `35.25pt`; keeping the
  multiplier would drift them to about `20.9pt` and `40.9pt`.
- v156 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-014359.docx`.
  Word COM extraction found `19` InlineShapes and `19` `Equation.DSMT4` OLE
  objects. `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks and
  `19` valid MathType OLE objects; `wmf_record_report.py` reported `19` vector
  WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX leaks.
- v157 introduced `MathTypeStructureMetrics` as the Java-side seed table for
  source-WMF structure constants. `LaTeXImageRenderer.estimateVectorHeightPt`
  now reads linear/script/fraction/nested-fraction/text-fraction/sqrt/
  sqrt-fraction/array/accent heights from that table, and
  `VectorWmfFormulaRenderer` uses it for standard glyph font size/script ratio
  and ordinary/compact fraction height families. This is not the full final
  parameter-table integration yet; many local placement constants still need
  later source/ink-backed migration.
- A no-context v157 review found the important split: outer Word shape height
  can be source-seeded while internal WMF ink geometry still uses old local
  constants. The first correction was to move ordinary stacked fraction height
  from the old `26.2pt` constant to the source-WMF `28.0pt` seed and make the
  unified fraction box split that same family into above/below metrics. Do not
  broaden `layoutUnifiedBox`; keep it limited to script-slot fractions until
  nested display fraction geometry has its own tall parser model.
- v157 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-014933.docx`.
  Word COM extraction found `19` InlineShapes and `19` `Equation.DSMT4` OLE
  objects. `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks and
  `19` valid MathType OLE objects; `wmf_record_report.py` reported `19` vector
  WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX leaks.
  `git ls-files src/main/java/com/lz/paperword/core/render/MathTypeStructureMetrics.java`
  was empty at this point, so remember to include this new source file when
  staging/committing.
- A no-context v158/v159 review caught a real source-table integration gap:
  `\frac{\sqrt{a^{2}+b^{2}}}{2}` was estimated as the `sqrt_fraction`
  `35.25pt` family by `LaTeXImageRenderer`, but `layoutFractions(...)` still
  placed the internal WMF as an ordinary `28pt` fraction. This is the
  problematic "source-seeded outer box, old internal geometry" class. v159
  routes non-compact fractions containing `\sqrt` to the `sqrt_fraction`
  height family inside `VectorWmfFormulaRenderer` as well.
- WMF polyline tests must read `META_POLYLINE` point y coordinates at
  `offset+10` and `offset+14`, not `offset+8` and `offset+12`. The latter are x
  coordinates and can make radical-line height assertions meaningless.
- Do not require text baselines in `\sqrt{1+\frac{a}{b}}` to exceed the
  `18pt` simple-root family just to prove `sqrt_fraction` routing. The radical
  line/overall window height is the better structural assertion; text baselines
  may remain within the old range while the root enclosure correctly grows.
- v159 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-015812.docx`.
  Word COM extraction found `19` InlineShapes and `19` `Equation.DSMT4` OLE
  objects. `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks and
  `19` valid MathType OLE objects; `wmf_record_report.py` reported `19` vector
  WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX leaks.
- The `script_fraction` source bucket is still thin (`24.0pt/33.75pt`), and
  current renderer policy remains "slash fraction only inside script context."
  Do not blindly map `x^{\frac{1}{2}}` to ordinary fraction height; give it a
  separate calibration pass once script-slot source/ink evidence is stronger.
- v160 adds a narrow Java seed for `script_fraction` height families:
  pure script-slot fractions such as `x^{\frac{1}{2}}+a^{\frac{2}{3}}` use the
  TeXToggle `24.0pt` family, while mixed formulas with both a script-slot
  fraction and a top-level ordinary fraction such as
  `S_{\frac{1}{4}\mathrm{圆}}=\frac{1}{4}\pi r^{2}` use the `33.75pt` family.
  This is intentionally a height-family classifier, not a broad internal
  spacing calibration.
- Do not treat every formula containing both scripts and fractions as
  `script_fraction`. The safe v160 classifier requires a fraction inside a
  script group; it only uses the pure `24.0pt` family when there is no
  top-level fraction command. Ordinary `\frac{1}{2}` remains on the ordinary
  fraction family, and nested display fractions still keep their separate tall
  family.
- v160 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-020512.docx`.
  Word COM extraction found `19` InlineShapes and `19` `Equation.DSMT4` OLE
  objects. `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks and
  `19` valid MathType OLE objects; `wmf_record_report.py` reported `19` vector
  WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX leaks.
- A no-context v161 review caught a broad parameter-table contract problem:
  `estimateVectorHeightPt(...)` returned source-seeded heights from
  `MathTypeStructureMetrics`, but `calibratePreviewMetrics(...)` then applied
  legacy `heightScale` multipliers. Examples included ordinary fractions
  shrinking from `28.0pt` to `26.32pt`, text-heavy fractions getting `0.62x`,
  and arrays receiving old line-count multipliers. This contradicts the rule
  that source-seeded physical heights should not be scaled again.
- v161 keeps legacy width calibration but guards source-seeded physical heights
  from second scaling. Covered seeded families are array, fraction,
  nested_fraction, text_fraction, sqrt, sqrt_fraction, script_fraction, script,
  and accent. Old height multipliers remain only for non-table fallback shapes.
  Regression tests now call `calibratePreviewMetrics(...)` by reflection and
  assert final height still equals the structure seed for ordinary fraction,
  nested fraction, text-heavy fraction, sqrt, sqrt_fraction, root-in-fraction,
  pure script_fraction, mixed script_fraction, array, and accent.
- v161 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-021324.docx`.
  Word COM extraction found `19` InlineShapes and `19` `Equation.DSMT4` OLE
  objects. `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks and
  `19` valid MathType OLE objects; `wmf_record_report.py` reported `19` vector
  WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX leaks.
  The report listed `4` suspicious script-text samples with
  `requiresReview=false`, so they remain diagnostic encoding samples rather
  than a leak failure.
- A no-context v162 review caught that `estimateVectorHeightPt`,
  source-seeded-height guarding, and preview width class selection still had
  parallel structure classifiers. This made `MathTypeStructureMetrics`
  less reusable and allowed future structures to be wired into one path but not
  another. v162 introduces `MathTypeStructureMetrics.Family` and
  `FamilyMetrics`, then makes `LaTeXImageRenderer.classifyStructureFamily(...)`
  the single query used for estimated height, source-seeded-height guarding,
  and preview class.
- Keep the structure-family contract covered by tests. The v162 contract test
  asserts `family`, `heightPt`, `sourceSeededHeight`, and `previewClass` for
  linear, script, pure script_fraction, mixed script_fraction, ordinary
  fraction, nested_fraction, text_fraction, sqrt, sqrt_fraction, root-in-
  fraction, array row count, and accent. This is the guard against collapsing
  `sqrt_fraction` into generic `sqrt` or `script_fraction` into generic
  `fraction` when adding future parameters.
- v162 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-021818.docx`.
  Word COM extraction found `19` InlineShapes and `19` `Equation.DSMT4` OLE
  objects. `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks and
  `19` valid MathType OLE objects; `wmf_record_report.py` reported `19` vector
  WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX leaks.
- v164 wired the source-WMF `array` height family into both explicit arrays and
  renderer-supported array-like environments (`aligned`, `alignedat`,
  `gathered`, `matrix`, `pmatrix`, `bmatrix`, and `cases`). A no-context review
  caught that classifying only `\begin{array}` lets `cases/aligned` fall back to
  scalar height families even though `VectorWmfFormulaRenderer` renders them as
  array layouts.
- For array-like source heights, count top-level rows from the first array-like
  body instead of splitting the whole LaTeX string on every `\\`. A raw split
  can count nested array/cases line breaks and create a mismatch between the
  Word shape height and `layoutArray(...)` internals.
- `MathTypeStructureMetrics.ARRAY_SOURCE_HEIGHT_PT` is now the 33pt two-row
  source family, with `ARRAY_ROW_HEIGHT_PT` at 16.5pt. Keep future array tuning
  anchored to this source family before changing local row spacing.
- If Word COM reference extraction times out while exporting PDF or closing
  Word, run a narrower OLE-only check before judging the candidate invalid. For
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-023121.docx`,
  the minimal COM check found `InlineShapes=19` and `Equation.DSMT4=19` after
  clearing hidden no-title `WINWORD` processes left by the timed-out extraction.
- v165 moved plain accent internals one step closer to the source table:
  `layoutOverline(...)` and `layoutUnderline(...)` now use
  `MathTypeStructureMetrics.ACCENT_HEIGHT_PT` for their internal layout height,
  with overline/underline line placement constants kept in the same metrics
  class. This avoids the old split where the Word shape used the 15.75pt
  source family but the WMF internals still used 13-14pt local boxes.
- Add internal WMF tests for every source-seeded family as it is migrated. The
  v165 accent test renders `\overline{AB}` and `\underline{AB}` at 15.75pt and
  asserts `windowExtY`, polyline y, text y, vector content, and no DIB fallback.
  Outer `LaTeXImageRenderer` height tests alone are not enough.
- v165 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-024209.docx`.
  `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks, `19` WMF media,
  and `19` valid MathType OLE objects. `wmf_record_report.py` reported `19`
  vector WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX
  leaks. A minimal Word COM check found `InlineShapes=19` and
  `Equation.DSMT4=19`.
- v166 moved text-heavy fractions onto their own source-seeded internal WMF
  family. `LaTeXImageRenderer` already classified `text_fraction` as `30pt`,
  but `VectorWmfFormulaRenderer.layoutFractions(...)` still rendered it through
  ordinary/compact fraction geometry. Keep the ordinary `28pt` fraction shape
  centered inside the `30pt` text-fraction family until source `0x0626` or ink
  data gives better numerator/bar/denominator offsets.
- `isCompactInlineFraction(...)` must exclude `hasTextHeavyFraction(...)`.
  Mixed formulas such as
  `\frac{三角形ABD的面积}{三角形CBD的面积}=\frac{AO}{CO}` contain a small inline
  fraction, but the visible dominant structure is still the text-heavy
  fraction. Letting compact routing win recreates the old cramped 18pt profile.
- v166 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-024818.docx`.
  `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks, `19` WMF media,
  and `19` valid MathType OLE objects. `wmf_record_report.py` reported `19`
  vector WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX
  leaks. A minimal Word COM check found `InlineShapes=19` and
  `Equation.DSMT4=19`.
- v167 moved true nested display fractions onto the `nested_fraction` internal
  WMF height family. Before the change, rendering
  `\frac{1+\frac{a}{b}}{2+\frac{c}{d}}` in a 53.25pt box still put the deepest
  text around 35pt and the lowest bar around 27pt, leaving the source-seeded
  lower half mostly empty. `layoutFractions(...)` now uses a nested branch
  separate from compact, sqrt_fraction, and text_fraction.
- The current nested split is intentionally conservative: outer nested
  fractions use `NESTED_FRACTION_ABOVE_PT=30pt` and the source total
  `53.25pt`, while inner fractions keep the ordinary recursive model. This
  fixes the "outer shape tall, internal ink ordinary" mismatch without
  pretending source `0x0626` line geometry has been decoded.
- v167 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-025308.docx`.
  `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks, `19` WMF media,
  and `19` valid MathType OLE objects. `wmf_record_report.py` reported `19`
  vector WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX
  leaks. A minimal Word COM check found `InlineShapes=19` and
  `Equation.DSMT4=19`.
- v168 keeps script-slot fractions as slash fractions. Do not "fix" pure
  `script_fraction` by turning `x^{\frac{1}{2}}` into a stacked display
  fraction; earlier Word clipping showed the slash policy is intentional for
  script context. The source-table integration point is the unified-box
  `FormulaLayout.heightPt()`, not the script-slot fraction drawing semantics.
- `VectorWmfFormulaRenderer.layoutUnifiedBox(...)` now pins pure script-slot
  fractions to `SCRIPT_FRACTION_HEIGHT_PT=24pt` and mixed script-slot plus
  top-level fraction formulas to `SCRIPT_FRACTION_MIXED_HEIGHT_PT=33.75pt`.
  This mirrors `LaTeXImageRenderer` classification while preserving the
  existing forced-slash child fraction behavior.
- v168 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-030322.docx`.
  `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks, `19` WMF media,
  and `19` valid MathType OLE objects. `wmf_record_report.py` reported `19`
  vector WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX
  leaks. A minimal Word COM check found `InlineShapes=19` and
  `Equation.DSMT4=19`.
- A no-context v168 review caught a real classifier mismatch: the vector
  renderer's original `hasFractionInsideScript(...)` used a depth-blind script
  operator scan, while `LaTeXImageRenderer` only treats top-level script slots
  as `script_fraction`. Without the fix,
  `\frac{x^{\frac{1}{2}}}{2}` could be misread as `script_fraction_mixed` or
  fail to render after correcting the route. Use a depth-aware top-level script
  scanner for source-family classification, and let `layoutFractionPart(...)`
  call `layoutUnifiedBox(...)` so display-fraction children such as
  `x^{\frac{1}{2}}` still render as slash script-slot fractions inside the
  correct outer display/nested fraction path.
- v169 moved the remaining simple-root WMF geometry constants into
  `MathTypeStructureMetrics`: root body x/y offset, width pad, checkmark
  x-points, top y, bottom pad, and left descent ratios. This does not claim a
  new visual calibration; it makes `layoutSqrt(...)` and `SqrtFormulaBox` share
  the same parameter-table entry points before later source/ink tuning.
- Root-line tests should not assert absolute x coordinates after final preview
  fitting. `fitLayoutToTarget(...)` can add left padding or scale x distances.
  The reliable invariant is the root checkmark's relative shape ratio plus y
  placement within a small twip tolerance. v169 added
  `sqrtRootShapeCoordinatesComeFromStructureMetrics()` for that.
- v169 generated
  `analysis/trace-runs/20260616-unified-box-candidate/unified-box-formula-candidate-20260617-031105.docx`.
  `scan_docx_latex_leaks.py` reported `0` visible LaTeX leaks, `19` WMF media,
  and `19` valid MathType OLE objects. `wmf_record_report.py` reported `19`
  vector WMFs, `0` bitmap WMFs, `0` StretchDIB records, and `0` WMF LaTeX
  leaks. A minimal Word COM check found `InlineShapes=19` and
  `Equation.DSMT4=19`.
- v170 fixed the Python source-WMF parameter chain to match the Java structure
  family contract for script-slot fractions. `script_fraction` (`24pt`) and
  `script_fraction_mixed` (`33.75pt`) are now separate candidate buckets; do not
  merge them when copying source seeds into `MathTypeStructureMetrics`.
- `aggregate_wmf_structure_metrics.py` must scan only top-level script slots
  when deciding `script_fraction` / `script_fraction_mixed`. A nested/display
  fraction such as `\frac{x^{\frac{1}{2}}}{2}` remains `nested_fraction`, not
  mixed script-fraction. This mirrors the depth-aware Java scanner and prevents
  display fractions from being polluted by child script-slot fractions.
- Candidate reports now distinguish directly usable renderer fields from
  missing parameters. `dominantHeightPt`, `mainFontPt`, and source
  `scriptRatio` can seed families, but bar/radical geometry, text-heavy
  numerator/denominator baselines, root index placement, array spacing, and
  script-slot/top-level baseline relationships still need `0x0626` decoding or
  rendered ink evidence.
- `sourceScriptFontRatio` is a source WMF record ratio, not the same thing as
  the Java visual script ratio. Keep `VISUAL_SCRIPT_FONT_RATIO` ink-validated;
  do not replace it blindly with the source `0.577/0.583` ratio.
- `coverageCount` in parameter candidates is discovery-only. Calibration
  strength comes from mapped WMF `count`, not source-report-only coverage. The
  regenerated combined report used the broad `--source-report-root analysis`
  root and restored `4404` mapped formulas with `0` skipped unmapped WMFs; a
  narrower source-report root gave `811` skipped rows and should not be treated
  as the final corpus.
- v171 wired ImageMagick ink bbox into
  `scripts/aggregate_wmf_structure_metrics.py`. Use `--with-ink` explicitly and
  use `--require-magick-ink` for calibration claims; the aggregate script must
  fail if any formula has `magickInkError`, because `inspect_docx(...)` bypasses
  the CLI-only failure check in `measure_wmf_formula_glyphs.py`.
- System Python may not have Pillow. For ink aggregation, use the bundled
  Python path from Codex desktop:
  `C:\Users\11703\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`.
  With system Python, `--require-magick-ink` now fails fast on
  `pillow_not_available` instead of writing an empty ink report.
- Empty rendered ink is not a successful measurement. Aggregation excludes
  `inkCount=0` from `magickInkWidthPt`, `magickInkHeightPt`, and ink ratio
  stats, and reports `magickBlankInkCount` separately so blank renders cannot
  silently pull medians toward zero.
- Ink ratios in the aggregate report are visible ink divided by the Word shape
  canvas, not WMF placeable/window units. Keep comparing `shapeWidthPt`,
  `wmfPlaceableWidthPt`, and Word spot checks before copying ink ratios into
  renderer constants.
- The TeXToggle expanded sample now produces a usable ink candidate report at
  `analysis/wmf-structure-metrics/word-mathtype-tex-reference-expanded-ink-summary.json`.
  The validation run had `22` formulas, `22` ink samples, `0` ink errors, and
  `0` blank ink samples; structures such as `script`, `fraction`, `sqrt`, and
  `sqrt_fraction` expose `magickInkWidthRatio`, `magickInkHeightRatio`, and
  `magickInkCenterYRatio` plus per-structure ink sample counts.
- v172 added `MathTypeStructureMetrics.SourceSampleMetrics` and `InkMetrics` as
  the Java-side read point for source/ink evidence. These values are calibration
  evidence, not automatic layout transforms: thin TeXToggle samples such as
  `script_fraction` (`n=1`) and `nested_fraction` (`n=2`) must not directly
  drive glyph scaling without Word spot checks and broader corpus support.
- Keep `sourceSampleMetrics(...)` separate from `metrics(...)`. `metrics(...)`
  is the active height contract used by `LaTeXImageRenderer` and
  `VectorWmfFormulaRenderer`; `sourceSampleMetrics(...)` records provenance
  fields such as source main font, source script ratio, ink width/height ratio,
  ink centerY ratio, and ink sample count for future per-structure tuning.
- The current ink ratios are most useful for comparing generated previews
  against official source WMFs by structure. They should be used to decide the
  next targeted geometry change, not copied blindly into `previewScale`, font
  width, or `ExtTextOut dx`.
- v173 added `--out-candidates-java` to
  `scripts/aggregate_wmf_structure_metrics.py`. It writes a reviewable Java
  switch fragment such as
  `analysis/wmf-structure-metrics/word-mathtype-tex-reference-expanded-ink-source-sample-metrics.javafrag`
  so `MathTypeStructureMetrics.sourceSampleMetrics(...)` can be refreshed from
  candidate JSON instead of hand-copying numbers.
- Treat the generated Java fragment as evidence, not as an automatic patch. The
  TeXToggle-only sample proposes candidate heights such as ordinary fraction
  `30.75pt` and nested fraction `60pt`, while the active renderer contract may
  intentionally stay on combined/XSC-backed heights such as `28pt` and
  `53.25pt`. Review provenance and Word spot checks before syncing active
  heights.
- `SourceSampleMetrics.candidateHeightPt` is named deliberately. It records the
  source candidate height from a report; the active renderer height remains
  `MathTypeStructureMetrics.metrics(...)`. Do not infer that every candidate
  height should immediately replace the active height family.
- v174 fixed the first concrete candidate/active-height mismatch in the Java
  evidence table. `SourceSampleMetrics.candidateHeightPt` now records the
  TeXToggle source candidate height for structures such as script `18.75pt`,
  ordinary fraction `30.75pt`, and nested fraction `60pt`, while
  `metrics(...)` keeps the active renderer heights (`16pt`, `28pt`, `53.25pt`)
  until broader source/Word evidence justifies changing rendered output.
- Tests should assert both values when they intentionally differ. This keeps the
  parameter table honest: source evidence can disagree with active rendering
  without silently becoming either stale documentation or an unreviewed visual
  change.
- v175 extended `scripts/compare_wmf_formula_glyphs.py` from object-index-only
  comparison into a structure-aware diff tool. Use `--pair-by formula-key` when
  comparing source MathType WMFs against generated WMFs from the same formula
  list; it now reports per-structure ink width/height deltas, record-width
  deltas, low-trust record-only cases, and worst examples.
- Generated-vs-source comparison needs a generated DOCX built from the same
  formula corpus, not the separate 19-formula unified-box demo. v175 added
  `GeneratedMathTypeTexReferenceDocxTest` to build a current-renderer DOCX and
  request JSON from
  `analysis/wmf-structure-metrics/word-mathtype-tex-reference-expanded.source-report.json`.
- Current renderer comparability gaps are explicit: the MTEF/OLE path rejected
  `\dfrac` / `\cfrac`, so the analysis generator normalizes them to `\frac`;
  strict self-vector WMF still rejects indexed roots such as `\sqrt[3]{8}=2`,
  so those are skipped for now rather than mixed into size calibration.
- First same-formula source-vs-generated diff artifact:
  `analysis/wmf-structure-metrics/generated-word-mathtype-tex-reference-expanded/generated-vs-source-20260617-034550-diff.txt`.
  It paired 20 formulas by formula key with zero missing pairs. Structure
  averages showed linear formulas much too narrow (`inkW_avg=-14.162pt`),
  simple sqrt too narrow (`inkW_avg=-10.646pt`), script_fraction too wide and
  too short (`inkW_avg=+14.266pt`, `inkH_avg=-8.5pt`), ordinary fraction width
  roughly near source but height short (`inkW_avg=-0.57pt`,
  `inkH_avg=-3.303pt`), and nested_fraction too wide but shorter
  (`inkW_avg=+11.118pt`, `inkH_avg=-5.28pt`).
- Treat record-width deltas from source MathType WMFs cautiously in this diff:
  official source WMFs often group text into different runs than the generated
  self-written WMF. v175's diff labels many rows `low_run_structure_mismatch`;
  for next renderer changes, prioritize same-formula ink bbox deltas and use
  record/run rows for diagnosis, not as a direct pass/fail metric.
- v176 added source-driven preview width scale entry points for linear, sqrt,
  and sqrt_fraction families. On the TeXToggle same-formula corpus, linear
  average ink-width delta improved from `-14.162pt` to `-6.103pt`, and sqrt
  improved from `-10.646pt` to `+1.469pt`. The generated artifacts were
  `analysis/wmf-structure-metrics/generated-word-mathtype-tex-reference-expanded/generated-word-mathtype-tex-reference-expanded-20260617-085037.docx`
  and `generated-vs-source-20260617-085037-diff.txt`.
- Do not treat all plain `sqrt` formulas as one width class. v176 made
  ordinary roots like `\sqrt{8}=2\sqrt{2}` and `\sqrt{x+1}` much closer, but
  over-expanded `\sqrt{a^{2}+b^{2}}` from the source `41.274pt` ink width to
  `52.874pt`. Root body formulas containing scripts need their own width
  parameter.
- v177 split script-body roots via
  `MathTypeStructureMetrics.SQRT_SCRIPT_PREVIEW_WIDTH_SCALE`. The same
  `\sqrt{a^{2}+b^{2}}` ink width moved from v176 `52.874pt` to `42.732pt`
  against source `41.274pt`, while ordinary-root improvements were preserved.
  The generated artifacts were
  `analysis/wmf-structure-metrics/generated-word-mathtype-tex-reference-expanded/generated-word-mathtype-tex-reference-expanded-20260617-085433.docx`
  and `generated-vs-source-20260617-085433-diff.txt`.
- After v177, the next evidence-backed targets are no longer plain sqrt.
  Worst same-formula ink deltas are script relation formula
  `S_{\Delta AOB}:S_{\Delta COD}=a^{2}:b^{2}=4:9` (`-16.351pt`),
  nested fraction `\frac{1+\frac{a}{b}}{2+\frac{c}{d}}` (`+16.227pt`),
  mixed linear/fraction `AO=1,\ CO=\frac{5}{3}` (`-16.162pt`),
  and script-slot fraction `x^{\frac{1}{2}}+a^{\frac{2}{3}}` (`+14.266pt`,
  height `-8.500pt`). Handle these as separate structure classes rather than
  by broad global width scaling.
- v178 targets visible line quality rather than object size. Fraction bars and
  radicals share the WMF pen, so line weight should live in
  `MathTypeStructureMetrics.STRUCTURE_LINE_WIDTH_PT`; the first visual-quality
  pass changed it from the old hard-coded `0.45pt` to `0.32pt`. Bump
  `LaTeXImageRenderer.CACHE_VERSION` after this change or Word will keep old
  cached line previews.
- Root signs should be emitted as one continuous `META_POLYLINE` with four
  points, not three independent two-point records. Independent root segments
  show visible cap/joint artifacts in Word. v178 updated both `layoutSqrt(...)`
  and `SqrtFormulaBox.emit(...)` to use a multi-point `LineSegment.polyline`.
  The generated v178 sample
  `analysis/wmf-structure-metrics/generated-word-mathtype-tex-reference-expanded/generated-word-mathtype-tex-reference-expanded-20260617-090452.docx`
  stayed self-vector (`20` WMFs, `20` vectorText, `20` vectorContent,
  `0` StretchDIB, `0` bitmap WMF). In its WMF report, simple root records such
  as image_eq16/image_eq17 now contain one radical polyline instead of three
  separate root-segment polylines.
- When parsing WMF polylines in tests, read the declared point count and all
  points. Old helpers assumed exactly two points, which would hide regressions
  once roots, arcs, or future brackets use real multi-point paths.
- v181 fixed two visual-quality issues exposed by the golden corpus Word/PNG
  gate. First, `appendScaledLayout(...)` and compact-fraction scaling must
  preserve all `LineSegment` polyline points; reducing a four-point radical to
  its first and last point makes nested roots collapse into stray lines. Second,
  ordinary `\rightarrow` / `\to` should use self-written polyline arrows rather
  than Symbol-font glyph fallback, because Word/PDF rendered the Symbol right
  arrow byte as a short dash even after WMF records contained the expected
  symbol byte. The regenerated golden corpus kept `23` valid MathType OLE,
  `23` vector WMFs, zero LaTeX leaks, zero StretchDIB/bitmap WMFs, and page 4
  showed a visible arrow in the chemistry reaction. Visual acceptance still
  fails overall: nested radicals, radicals containing fractions, text-heavy CJK
  fractions, nested fraction proportions, and chemistry spacing still need
  structure-specific layout work.
- A no-context v181 review caught that preserving polyline points only in
  scaled append paths was incomplete. Plain `appendLayout(...)` is also used by
  nested radicals, so it must translate all `LineSegment` points rather than
  rebuilding a two-point segment from x1/y1/x2/y2. The follow-up test now
  asserts `\sqrt{1+\sqrt{x}}` emits two four-point radical polylines, and the
  golden corpus records `maxPolylinePointCount`. The same review also caught
  that simple arrows inside array cells need the new self-drawn arrow path, and
  that `chemistry-equation-01` should require `3` polyline records so a
  `\rightarrow` regression cannot pass as text-only WMF.
- v182 fixed another polyline-preservation hole in the unified formula box
  path. `emitScaled(...)`, `emitScaledScript(...)`, and `scaleLayoutX(...)`
  previously rebuilt structure lines from only x1/y1/x2/y2, which would discard
  intermediate points for four-point radical signs whenever the structure was
  placed in a scaled fraction slot, script slot, or horizontally compressed
  layout. Use `LineSegment.scaled(...)` and `LineSegment.scaledX(...)` for
  these transforms so future multi-point roots/arrows/brackets survive.
- v182 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-latest.docx`.
  `scan_docx_latex_leaks.py` still reported `23` MathType
  `Equation.DSMT4` OLE objects, `23` WMF previews, and zero visible LaTeX
  leaks. `wmf_record_report.py` still reported `23` vector WMFs, zero
  `StretchDIB`, zero bitmap records, and zero WMF LaTeX leaks. Word COM export
  plus Poppler PNG render refreshed
  `analysis/formula-golden-corpus/word-rendered/page-1.png` through
  `page-4.png`.
- v182 visual status is still not acceptable. Page 2 shows that roots no longer
  collapse into stray two-point lines, but ordinary and nested radicals remain
  too thin/tall and nested-root proportions are cramped. Page 3 shows
  `text_mixed` CJK fractions still look like two stacked text lines rather than
  a proper long fraction with a visible centered bar. Page 4 keeps the
  chemistry arrow visible, but chemistry spacing and style remain unnatural.
- A no-context v182 review caught three gate/geometry gaps. First,
  `layoutVerticalExtentPt(...)` must compute line bounds from every polyline
  point, because a radical's deepest point is usually an intermediate point,
  not y1/y2. Second, corpus validation should count four-point radical
  polylines against the actual number of `\sqrt` commands, not merely check
  `maxPolylinePointCount >= 4`. Third, `FormulaGoldenCorpusDocxTest` must apply
  the TSV structure thresholds to the embedded WMFs as well; OLE/vector-text
  counts alone can pass text-only or collapsed-radical previews. The DOCX gate
  should not compare embedded WMF dimensions to TSV display-box dimensions,
  because `DocxBuilder` recalibrates the actual object size by structure family.
- v183 is a stage improvement, not visual acceptance. It raised the text-heavy
  fraction family to `33pt`, made its numerator/bar/denominator y positions
  explicit, and moved radical bottom/left-descent geometry through
  `MathTypeStructureMetrics`. The golden DOCX still passes the
  structural/OLE/vector gates, and Word COM + Poppler render shows case 15
  text-mixed fractions now read as real long fractions instead of two stacked
  CJK text rows. Residual visual failures remain obvious: nested radicals and
  radicals containing fractions are still cramped, with too-close inner/outer
  root shapes and overlong left descenders.
- Do not keep derived text-fraction offset constants after switching the family
  to explicit source-like y positions. A stale
  `TEXT_FRACTION_VERTICAL_OFFSET_PT` implies ordinary fraction geometry still
  drives text-heavy fraction layout, which is false and can mislead the next
  tuning round.
- A no-context v183 review caught two real risks before commit. First,
  `STRUCTURE_LINE_WIDTH_PT` is global for all polylines, so changing it to fix
  one CJK fraction line can also change roots, arrows, boxes, and future grid
  lines. Keep it at the v178 `0.32pt` until there is a per-structure pen model
  or visual proof for every affected structure. Second, detecting a
  text-heavy fraction at formula level must not force ordinary sibling
  fractions such as trailing `AO/CO` to use text-heavy y coordinates. Compute
  text-heavy fraction y placement per `\frac{...}{...}` node, while the overall
  formula can still use the taller text-fraction height family.
- v184 lifts only tall radical checkmark turns through
  `MathTypeStructureMetrics.sqrtBottomPadPt(rootHeightPt)`. Keep the ordinary
  `SQRT_BOTTOM_PAD_PT` for simple `18pt` roots, but use a larger tall-root pad
  for `sqrt_fraction` / nested-fraction height roots. Record probes showed
  `\sqrt{1+\frac{a}{b}}` moved the radical turn from `643twips` to `529twips`,
  and `\sqrt{1+\sqrt{\frac{a}{b}}}` moved outer/inner fraction-root turns from
  about `683/667twips` to `569/553twips`, while `\sqrt{x+1}` stayed unchanged.
- v184 Word COM + Poppler visual status: page 2/page 3 show shorter descenders
  for roots containing fractions, including `T=2\pi\sqrt{\frac{l}{g}}`, so the
  change is a real visual improvement. It is not final acceptance. Nested roots
  with fractions still look like crowded parallel vertical strokes; next tuning
  should address nested-root spacing/body placement rather than further raising
  every tall-root turn.
- A no-context v184 review flagged that root-height thresholding is broader
  than a named structure family and that record-coordinate tests can overclaim
  visual repair. Keep this round's claim narrow: tall `sqrt_fraction` and
  `sqrt_nested_fraction` turns are lifted and Word PNG shows shorter descenders,
  but root-in-fraction and nested-root proportions still need separate visual
  calibration. Do not describe v184 as completing radical rendering.
- v185 is another stage improvement, not acceptance. It adds a nested-root-only
  body y offset (`SQRT_NESTED_BODY_Y_EXTRA_PT = 2.2pt`) when the body of a
  `\sqrt{...}` contains another `\sqrt`. Record probes moved nested radical top
  gaps from `24twips` to `68twips` for simple nested, nested-fraction, and deep
  nested roots. This reduces the earlier Word visual failure where inner and
  outer radical bars/stems looked glued together.
- v185 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-latest.docx`; the leak
  scanner still reported `23` MathType `Equation.DSMT4` OLE objects, `23` WMF
  previews, and zero visible LaTeX leaks. `wmf_record_report.py` still reported
  `23` vector WMFs, zero `StretchDIB`, zero bitmap records, and zero WMF LaTeX
  leaks. Word COM export plus Poppler PNG render refreshed
  `analysis/formula-golden-corpus/word-rendered/page-1.png` through
  `page-4.png`.
- v185 Word PNG status: page 2 case 11/12 nested radicals are visibly clearer
  than v184 because the radical bars no longer sit on top of each other. It is
  still not final. Page 3 case 13 deep nested radicals look too flat/compressed,
  and page 3 case 18 physics `sqrt_fraction` remains crowded around the root
  body and fraction. Continue with a per-structure layout model for nested
  depth, root-in-fraction placement, and radical stroke proportions rather than
  treating one global nested offset as the final solution.
- A no-context v185 review caught that `bodyText.contains("\\sqrt")` was too
  broad for nested-root detection, because text-like command bodies can contain
  a literal `\sqrt` without being a structural radical. Use
  `hasSqrtCommandOutsideText(...)` for the nested-root offset so `\text{...}`,
  `\mathrm{...}`, and related style/text command bodies do not change root body
  placement. The same review also pointed out that production fixed-height
  rendering can still compress deep nested radicals even when record-level top
  gaps improve, so keep a production-envelope test and continue treating Word
  PNG review as required evidence.
- v187 is a scoped deep-nested-radical improvement, not acceptance. It adds
  `sqrtCommandDepthOutsideText(...)`, applies `SQRT_NESTED_BODY_Y_EXTRA_PT` per
  nested depth in `layoutSqrt(...)`, and introduces a `SQRT_NESTED` structure
  family at `64pt` only for non-fraction roots with structural depth greater
  than `2`. This keeps two-level cases such as `\sqrt{1+\sqrt{x}}` on the
  ordinary `18pt` root family, while giving `\sqrt{x+\sqrt{y+\sqrt{z}}}` enough
  production height to avoid the worst vertical flattening.
- A bad intermediate v186-style threshold (`sqrtCommandDepthOutsideText > 1`)
  inflated two-level nested roots to `64pt`, moved case 12 to a later page, and
  made the golden corpus visually worse even though the structural tests still
  passed. After narrowing the threshold to `> 2`, bump cache again; the verified
  cache key for this round is `v187-deep-nested-root-family`.
- v187 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-latest.docx` and
  `formula-golden-corpus-20260621-214150.docx`. The leak scanner still reported
  `23` MathType `Equation.DSMT4` OLE objects, `23` WMF previews, and zero
  visible LaTeX leaks. `wmf_record_report.py` still reported `23` vector WMFs,
  zero `StretchDIB`, zero bitmap records, zero WMF LaTeX leaks, and zero
  suspicious reviews required.
- v187 Word COM + Poppler PNG status: page 2 restores cases 7-12 after the bad
  threshold experiment; case 11/12 keep the v185 visual improvement without
  becoming oversized. Page 3 case 13 deep nested radicals have clearer layer
  separation than v185, but still look too high/flat and are not final. Page 3
  case 18 physics `sqrt_fraction` remains crowded. Page 4 still shows the
  chemistry arrow, but chemistry spacing and higher-math fraction styling remain
  future work.
- A no-context v187 review caught a real production-classification blocker:
  checking `hasFractionCommand(text)` before root depth meant
  `\sqrt{1+\sqrt{\frac{a}{b}+\sqrt{z}}}` stayed in the `SQRT_FRACTION`
  `35.25pt` family instead of the new `SQRT_NESTED` `64pt` family. The fix is
  depth-first classification for non-indexed structural roots: if root depth is
  greater than `2`, choose `SQRT_NESTED` before checking ordinary
  `sqrt_fraction`. Add tests for estimate height, calibrated height, and
  `classifyStructureFamily(...)` so this exact deep-nested-with-fraction case
  cannot regress silently.
- v188 targets a different `sqrt_fraction` failure: mixed linear/root formulas
  such as `T=2\pi\sqrt{\frac{l}{g}}` were using the same narrow
  `SQRT_FRACTION_PREVIEW_WIDTH_SCALE` as pure roots. The golden metrics showed
  case 18 at only `38.25pt` wide against the TSV's `86pt` reference width,
  while pure case 9/10 should not be widened. Use a separate
  `SQRT_FRACTION_MIXED_PREVIEW_WIDTH_SCALE` only when the formula is not a
  top-level fraction and has visible top-level text outside the `\sqrt{...}`
  group.
- v188 validation after this width split: case 18 grew from `38.25pt` to
  `49.77pt` shape width and record width from `33.05pt` to `43.15pt`, while
  case 9 stayed `28.68pt` and case 10 stayed `47.81pt`. Structural gates still
  reported `23` valid MathType OLE objects, `23` vector WMFs, zero LaTeX leaks,
  zero bitmap/StretchDIB WMFs, and zero suspicious-review-required WMF text.
  Word PNG page 3 shows case 18 less horizontally cramped, but the internal
  `l/g` fraction under the root is still not MathType-quality. Continue with a
  root-body fraction layout pass; do not call v188 final visual acceptance.
- v189 scopes the first root-body fraction pass to ordinary `\frac` whose
  numerator and denominator are short simple atoms, then applies
  `SQRT_BODY_FRACTION_SCALE = 0.78` only to that root body. Do not broaden this
  to every whole fraction under a radical: no-context review caught that
  `\sqrt{\frac{1+\frac{a}{b}}{2+\frac{c}{d}}}`, root-containing fractions,
  text-heavy fractions, `\dfrac`, and `\cfrac` would otherwise inherit the same
  shrink and risk undoing nested-fraction/text-fraction calibration.
- v189 also adds a small root-body fraction geometry adjustment:
  `SQRT_BODY_FRACTION_LEFT_ADJUST_PT = -0.8` and
  `SQRT_BODY_FRACTION_TOP_PAD_PT = 2.4`. Keep tests for both right coverage and
  left protrusion; an earlier `-1.2pt` left shift was review-risky because the
  body could sit left of the radical top turn.
- v189 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-222128.docx`
  and `analysis/formula-golden-corpus/word-rendered-v189-final/page-*.png`.
  The leak scanner reported `23` `Equation.DSMT4` OLE objects, `23` WMF media,
  zero visible LaTeX leaks, and zero invalid MathType OLE. `wmf_record_report`
  reported `23` vector-text/vector-content WMFs, zero `StretchDIB`, zero bitmap
  records, zero WMF LaTeX leaks, and zero review-required suspicious text.
  Glyph metrics kept case 18 at `49.77x35.25pt`, with the root-body `l/g`
  baselines at about `7.15pt` and `18.10pt`, so the internal fraction is more
  compact than v188. Word page 3 confirms it is more readable, but visual
  acceptance is still not final: `T=2\pi` to root spacing remains too loose and
  the radical left leg/proportion still does not match MathType quality.
- v190 narrows the next root-body fraction pass to a small left tuck:
  `SQRT_BODY_FRACTION_LEFT_ADJUST_PT = -1.6` and cache key
  `v190-sqrt-body-fraction-tuck`. An attempted `-2.4pt` shift failed the
  left-protrusion guard, so keep the root-body fraction inside the radical top
  turn instead of chasing visual tightness by over-shifting.
- v190 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-222902.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v190.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v190/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero visible
  LaTeX leaks, zero bitmap/StretchDIB WMFs, and zero review-required suspicious
  WMF text.
- v190 glyph metrics moved case 18's root-body `l/g` fraction left by about
  `2.2pt` and reduced record width from about `43.6pt` to `41.35pt` while
  keeping the shape at `49.77x35.25pt`. Word page 3 shows a small improvement,
  not final acceptance: the fraction is less floaty, but the radical left leg,
  root top bar, and MathType-like proportion still need a more principled root
  drawing pass.
- No-context v190 review found no blocking issue, but flagged that the new
  guard is only WMF record/bar-level. Future acceptance should compare glyph
  ink or at least text-run min/max x against the radical top turn and top bar;
  a fraction bar can stay inside while numerator/denominator ink protrudes or
  visually crosses.
- v191 narrows only tall radical checkmark geometry: ordinary roots keep
  `SQRT_CHECK_TOP_X_PT = 5.0`, while roots taller than the ordinary family use
  `SQRT_TALL_CHECK_TOP_X_PT = 4.0`. This lets compact root-body fractions use
  `SQRT_BODY_FRACTION_LEFT_ADJUST_PT = -2.4` without repeating the old v190
  left-protrusion failure, because the radical top turn also moved left.
- v191 review correctly warned that reusing `-2.4pt` looked like the old failed
  v190 experiment. Do not rely on the parameter value alone; require direct
  guards that the compact body fraction bar and glyph start at or to the right
  of the radical top turn, and that nested fraction root top bars still cover
  their body text after the tall-root branch.
- v191 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-225053.docx`.
  Structural scans stayed clean: `23` MathType OLE objects, `23` WMF previews,
  zero visible LaTeX leaks, zero bitmap/StretchDIB WMFs, and zero
  review-required suspicious WMF text. Word PNG evidence is under
  `analysis/formula-golden-corpus/word-rendered-v191/page-*.png`; page 3 case
  18 has less empty space between the root turn and `l/g`, but the root is still
  visibly hand-drawn rather than MathType-quality.
- v191 does not move glyph runs for case 18; `glyph-v191-final.txt` keeps `l`
  at `x=37.85pt` and `g` at `x=36.75pt`. The improvement comes from the radical
  polyline geometry and the compact body placement guard, not from character
  sizing. A future pass still needs a more principled radical stroke model or
  ink-level comparison, especially for nested radicals.
- v192 increases only `SQRT_NESTED_BODY_Y_EXTRA_PT` from `2.2` to `3.0` and
  bumps the cache key to `v192-deep-sqrt-top-gap`. The intent is to make deep
  nested radicals read as separate stacked structures instead of compressing the
  body glyphs into the top band.
- v192 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-225504.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v192.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v192/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero visible
  LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs, and zero
  review-required suspicious WMF text.
- v192 glyph evidence: case 13 nested radical baselines moved from about
  `16.9/20.3/21.5pt` in v191 to about `18.5/22.7/23.9pt`, so nested content is
  more separated vertically while the shape remains `41.28x64.0pt`. The change
  is still not final MathType-quality; it is a local improvement for deep nested
  roots.
- v192 side effect: because `SQRT_NESTED_BODY_Y_EXTRA_PT` is global, case 11/12
  and case 18 also shift slightly. Case 18 record width changed from about
  `41.35pt` to `40.25pt`, and the root-body `l/g` run starts moved left by
  about `1.1pt`. Word page 3 still looks acceptable, but future work should
  consider splitting shallow nested-root spacing from deep nested-root spacing
  if shallow roots start drifting.
- v192 process note: the requested no-context review was launched, but did not
  return before this round's cutoff, and a replacement review could not be
  spawned because the agent thread limit was reached. Treat this as a workflow
  gap for the round, not as review acceptance.
- v193 splits the root-body horizontal slot instead of changing ordinary root
  padding: simple whole fractions directly under a radical now use
  `SQRT_BODY_FRACTION_LEFT_PAD_PT = 5.4`, while ordinary radical bodies still
  use `SQRT_BODY_LEFT_PAD_PT = 6.0`. Keep the existing
  `SQRT_BODY_FRACTION_LEFT_ADJUST_PT = -2.4`; this pass is a conservative
  `0.6pt` extra tuck for compact root-body fractions, not a broad root layout
  rewrite.
- v193 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-232223.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v193.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v193/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v193 glyph evidence is intentionally narrow: case 12's root-body `a/b` runs
  moved from `x=19.85pt` to `x=19.35pt`, and case 18's `l/g` runs moved from
  about `36.75/35.65pt` to `35.90/34.80pt`. Case 13 deep nested radical
  metrics did not change. Word page 2/3 show no left protrusion, but the root
  checkmark/left-leg shape is still visibly hand-drawn and should be the next
  larger design pass.
- v193 no-context review found no blocking risk and agreed that splitting the
  compact root-body fraction pad is better scoped than changing ordinary root
  padding. It flagged the first pad guard as too weak, so the test now also
  pins the fraction-bar gap after the radical top turn to a narrow positive
  `8..18` twip range.
- v194 introduces a tall-root-only checkmark midpoint:
  `SQRT_TALL_CHECK_MID_X_PT = 1.4`, while ordinary roots keep
  `SQRT_CHECK_MID_X_PT = 2.0` and tall roots keep
  `SQRT_TALL_CHECK_TOP_X_PT = 4.0`. This targets the visual complaint that tall
  radical rising strokes looked too vertical, without reopening ordinary-root
  spacing or the v191/v193 left-protrusion guards.
- v194 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-233201.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v194.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v194/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v194 does not move text glyphs; glyph metrics for case 12/13/18 stay at the
  v193 positions. Word page 2/3 show no structural regression and the tall-root
  rising stroke is slightly less vertical, but the radical is still a straight
  polyline with a mechanical look. A future root-shape pass should consider a
  richer multi-segment or curve-approximated radical stroke rather than only
  moving the existing four points.
- v194 no-context review found no blocking risk and confirmed that the
  tall-root-only midpoint is better scoped than changing the global
  `SQRT_CHECK_MID_X_PT`. It also noted a future cleanup: the tall-root threshold
  `SQRT_HEIGHT_PT + 4.0` is duplicated between layout and
  `sqrtBottomPadPt(...)`, so extract it before larger radical-shape changes.
- v195 moves tall radicals from a strict four-point polyline to a five-point
  multi-segment stroke by adding `SQRT_TALL_CHECK_LOW_X_PT = 1.1` and
  `SQRT_TALL_CHECK_LOW_Y_RATIO = 0.76`. Ordinary radicals remain four-point.
  This is the first pass toward a less mechanical radical stroke while staying
  inside the self-written WMF vector path.
- v195 required test helpers to stop assuming every radical is exactly four
  points. Use `isRadicalPolyline(pointCount >= 4)`, `radicalTopX(...)`, and
  `radicalTopY(...)` when checking radical geometry. A five-point tall radical
  makes the top turn the penultimate point, not `x(2)/y(2)`. This same mistake
  already caused failing tests in this round.
- v195 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-235244.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v195.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v195/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v195 visual status: page 2/3 show no broken root, extra short line, or
  boundary regression, and text glyph metrics remain at the v193/v194
  positions. The deep nested case 13 outer radical now has a visible shoulder,
  but the shoulder is still a bit abrupt; future work should tune the low-point
  ratio or add a second shoulder segment only after keeping the `>=4` radical
  guards stable.
- v195 no-context review found no blocking risk. The `pointCount >= 4`
  radical heuristic is currently safe because renderer polylines are only box
  sqrt, production normal sqrt, and production tall sqrt; if future bracket or
  decoration lines also use four-plus-point polylines, add a shape heuristic or
  explicit structure tag. The JSON field `fourPointPolylineRecords` is now a
  compatibility name; failure text should say "radical polyline with at least
  four points".
- v196 adds a second tall-root shoulder point:
  `SQRT_TALL_CHECK_SHOULDER_X_PT = 1.9` and
  `SQRT_TALL_CHECK_SHOULDER_Y_RATIO = 0.55`, so tall radicals now emit six
  points: start, low shoulder, bottom turn, middle shoulder, top turn, and top
  bar end. Ordinary radicals remain four-point, and the cache key is
  `v196-tall-sqrt-shoulder`.
- v196 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260621-235643.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v196.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v196/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v196 glyph evidence is intentionally unchanged for text: case 13 still has
  `x/y/z` runs at about `5.65/19.2/32.75pt`, and case 18 keeps `l/g` at about
  `35.9/34.8pt`. WMF bytes changed for tall-root cases such as image_eq9,
  image_eq11, image_eq12, image_eq13, and image_eq18, proving the extra
  radical point was written without moving glyph runs.
- v196 visual status: page 2/3 show no broken root, short-line regression, or
  bitmap fallback. The case13 zoom shows the outer tall radical has a more
  controllable shoulder, but both case13 and case18 still look too straight and
  thin on the left leg. The next useful pass should tune the left-leg/low-turn
  geometry or nested-root spacing, not keep adding generic points.
- v196 no-context review could not be spawned because the agent thread limit
  was reached. Treat this as a process gap for the round, not as review
  acceptance.
- v197 extracts the tall-radical threshold into
  `SQRT_TALL_EXTRA_HEIGHT_PT` and `isTallSqrt(...)`, then routes both
  `layoutSqrt(...)` and `sqrtBottomPadPt(...)` through the shared predicate.
  This removes the duplicated `SQRT_HEIGHT_PT + 4.0` guard that v194 review
  had flagged before further radical-shape tuning.
- v197 lowers the tall-root low shoulder by changing
  `SQRT_TALL_CHECK_LOW_Y_RATIO` from `0.76` to `0.80`. In WMF coordinates this
  moves the low shoulder downward, making the left-leg bottom turn more
  concentrated without moving glyph runs or changing ordinary four-point
  radicals. Cache key: `v197-tall-sqrt-threshold`.
- v197 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-001710.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v197.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v197/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v197 visual status: page 2/3 and zooms show no broken radical, short-line
  regression, or fold-back from the lower shoulder. The change is still a
  local improvement only; case12 nested-fraction roots remain too tight, and
  case13/case18 still look too straight compared with MathType.
- v197 no-context review found no blocking issue and confirmed that
  `isTallSqrt(...)`, `sqrtBottomPadPt(...)`, and `layoutSqrt(...)` now share
  the same tall-root predicate. The review correctly noted that a renderer
  threshold test should not assume the requested WMF outer height forces the
  internal root height; after one failing attempt, the regression now verifies
  the real routing instead: simple `\sqrt{x}` stays four-point, while
  `\sqrt{1+\frac{a}{b}}` switches to the multi-point tall radical.
- v198 adds `SQRT_NESTED_BODY_LEFT_EXTRA_PT = 1.2` and applies it only when a
  sqrt body contains another structural `\sqrt` and is not the whole-body
  compact fraction case. The same extra is added to the sqrt layout width, so
  the body is not moved right without expanding the WMF box. Cache key:
  `v198-nested-sqrt-body-pad`.
- v198 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-003415.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v198-final.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v198-final/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v198 glyph evidence shows the nested-root body pad is real, not just page
  zoom noise: case12 moves `1+` from about `7.15pt` to `7.95pt` and the
  fraction `a/b` from about `19.35pt` to `19.7pt`; case13 moves the nested
  `x/y/z` runs right, while case18 pure sqrt-fraction remains unchanged.
- v198 visual status: page 2/3 and zooms show no broken radicals, short-line
  regression, or bitmap fallback. This is still only a local improvement:
  case12 remains too narrow and mechanical, and case13/case18 radicals still
  need a better root-stroke model rather than more generic spacing tweaks.
- v198 no-context review found no blocker in the nested-sqrt condition, width
  accounting, or cache bump, but correctly flagged that the first horizontal
  gap assertion was too weak because ordinary body pad alone could satisfy it.
  The regression was tightened with a `\sqrt{\sqrt{x}}` control that expects
  the inner top-turn gap to include `SQRT_BODY_LEFT_PAD_PT` plus
  `SQRT_NESTED_BODY_LEFT_EXTRA_PT`.
- v199 opens the non-compact tall sqrt stroke instead of changing glyph runs:
  `SQRT_TALL_CHECK_LOW_X_PT`, `SQRT_TALL_CHECK_MID_X_PT`,
  `SQRT_TALL_CHECK_SHOULDER_X_PT`, and `SQRT_TALL_CHECK_TOP_X_PT` move right
  to make nested/tall roots less vertical and mechanical. Cache key:
  `v199-open-tall-sqrt-stroke`.
- v199 learned that pure sqrt-body fractions cannot share the wider tall-root
  turn. The first full test run failed because `\sqrt{\frac{l}{g}}` let the
  fraction bar protrude left of the radical top turn. Compact tall roots now
  use separate `SQRT_TALL_COMPACT_CHECK_*` point constants that preserve the
  old narrower coverage, while mixed/nested roots use the opened stroke.
- v199 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-004100.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v199.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v199/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v199 visual status: case13's outer nested root is visibly less vertical,
  case18 compact sqrt-fraction remains covered, and case12 has no broken
  radical or short-line regression. Remaining gap: root strokes still look
  too straight/hard at the top bar and case12's inner fraction body is still
  narrow; next passes should address stroke weight/corner shape or nested
  fraction slot sizing, not widen every root indiscriminately.
- v199 no-context review found no blocker. It confirmed that the opened tall
  stroke is scoped away from `compactBodyFraction`, all compact low/shoulder/
  mid/top X constants are consumed, the fraction-bar protrusion and top-bar
  coverage guards remain in tests, the new anti-vertical-leg assertion covers
  the widened tall root, and the v199 DOCX/PDF/PNG paths exist.
- v200 adds a scoped nested sqrt-body fraction scale instead of character
  special-casing. `SQRT_NESTED_BODY_FRACTION_SCALE = 0.90` applies only when a
  whole simple fraction is laid out inside an existing sqrt body
  (`sqrtDepth > 0`); standalone pure sqrt-body fractions keep
  `SQRT_BODY_FRACTION_SCALE = 0.78`. Cache key:
  `v200-nested-sqrt-fraction-scale`.
- v200 passes `sqrtDepth` through the `layoutFractionPart(...)`,
  `layoutFractions(...)`, `layoutSqrt(...)`, and common wrapper recursion
  (`underline`, `cancel`, `boxed`, `overline`) so nested scale is a structural
  decision even through decoration wrappers. This avoids data-specific fixes
  such as widening `a/b` but not `l/g`, and keeps physics case18
  `\sqrt{\frac{l}{g}}` unchanged while widening case12's inner
  `\sqrt{\frac{a}{b}}`.
- v200 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-005815.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v200-final.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v200-final/page-*.png`.
  Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v200 glyph evidence: case12's inner fraction glyph advance widens from about
  `2.85pt` to `3.15pt`, while case18 `l/g` stays at about `2.7/4.85pt`.
  Visual zoom confirms case12 is more readable and has no broken radical or
  short-line regression, but the inner radical/fraction spacing is still hard
  and not MathType-like enough.
- v200 no-context review found no direct blocker for case12, but caught three
  useful hardening items: wrapper layouts could drop `sqrtDepth`, the first
  glyph-width assertion compared different characters, and widened nested
  fraction-bar coverage was not guarded. The final patch fixes all three before
  commit.
- v201 moves fraction-bar length from local magic numbers to structure profiles
  in `MathTypeStructureMetrics`: compact, ordinary, text-heavy, nested, and
  sqrt-body fractions now choose separate inset/overhang values. Cache key:
  `v201-fraction-bar-profiles`.
- v201 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-010906.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v201.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v201/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v201 visual result: case15 text-heavy fractions now have bars that visually
  cover the wide CJK numerator/denominator text instead of stopping after a
  large inset, and the trailing `AO/CO` fraction is also closer to the glyph
  width. This is a real local improvement, but not a goal-level finish:
  case12 and case18 still look hard because the radical stroke/body placement
  model is too straight and cramped.
- v202 increases `SQRT_NESTED_BODY_LEFT_EXTRA_PT` from `1.2pt` to `2.2pt` and
  bumps the cache key to `v202-nested-sqrt-body-gap`. This is deliberately
  scoped to bodies that structurally contain another sqrt, so ordinary roots
  and compact `\sqrt{\frac{l}{g}}` keep their previous placement.
- v202 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-011741.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v202.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v202/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v202 glyph evidence: nested root bodies moved right without touching case18.
  Case11 `1+` moved from about `8.1pt` to `8.7pt`, case12 `1+` from about
  `7.75pt` to `8.35pt`, and case13 `x/y/z` moved right by about `0.35-0.85pt`;
  case18 `l/g` stayed at about `35.9/34.8pt`. Visual zoom shows a small spacing
  improvement but not a finished MathType look: the radical stroke itself is
  still too straight and the inner nested root still reads too vertical.
- v203 opens the non-compact tall-root shoulder only:
  `SQRT_TALL_CHECK_SHOULDER_X_PT` moves from `2.7pt` to `3.15pt` and
  `SQRT_TALL_CHECK_SHOULDER_Y_RATIO` moves from `0.55` to `0.50`. The tall
  midpoint and all compact tall-root constants stay unchanged, so compact
  `\sqrt{\frac{l}{g}}` coverage remains protected. Cache key:
  `v203-open-tall-sqrt-shoulder`.
- v203 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-012532.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v203.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v203/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v203 visual result is only a small stroke-shape improvement: case13's tall
  nested root is a little less pinched near the top turn, while case18 remains
  unchanged. Case12 is still not acceptable enough because the inner radical
  left leg and the fraction body still crowd together; the next useful step is
  a real nested-radical body/scale model or a richer radical stroke than a
  single WMF polyline.
- v204 adds a metric-level clearance for compact fraction bodies inside nested
  radicals only: `SQRT_NESTED_BODY_FRACTION_LEFT_EXTRA_PT = 0.85pt`. The guard
  is `compactBodyFraction && sqrtDepth > 0`, so standalone compact roots such
  as case18 `\sqrt{\frac{l}{g}}` keep their previous slot while case12's inner
  `\sqrt{\frac{a}{b}}` gets a wider gap between the radical top turn and the
  fraction bar/body. Cache key: `v204-pad-nested-sqrt-fraction`.
- v204 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-013417.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v204.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v204/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v204 glyph evidence shows the change is scoped: case12 `a/b` moves from
  about `19.05pt` to `19.9pt`, while case18 `l/g` stays at about
  `35.9/34.8pt`. Visual zoom says this is only a local readability
  improvement, not a final pass: the inner radical still looks too straight and
  mechanical. The next round should replace the radical stroke model with a
  richer multi-segment/curved-looking WMF polyline profile rather than adding
  more padding.
- v205 opens the compact tall-radical stroke profile without moving the
  fraction body: `SQRT_TALL_COMPACT_CHECK_SHOULDER_X_PT` moves from `1.9pt` to
  `2.65pt`, and compact tall roots now use their own
  `SQRT_TALL_COMPACT_CHECK_SHOULDER_Y_RATIO = 0.46`. An attempted
  `SQRT_TALL_COMPACT_CHECK_TOP_X_PT = 4.25pt` failed the structural guard
  because the radical top turn moved right of the compact fraction bar; keeping
  the top turn at `4.0pt` preserves coverage while still opening the shoulder.
  Cache key: `v205-open-compact-sqrt-stroke`.
- v205 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-014001.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v205.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v205/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v205 visual result is safe but still modest. Case18 keeps the same glyph
  placement (`l/g` around `35.9/34.8pt`) and the compact radical shoulder is a
  little more open, but Word's rendered PNG still looks mechanical. Case12 does
  not regress, but the inner radical still has a long straight leg. The next
  meaningful improvement should add more explicit stroke shape to compact tall
  roots, for example a separate short left hook/low shoulder segment or a
  slightly thicker/curved-looking multi-polyline profile, not another body
  position tweak.
- v206 adds the first explicit multi-polyline compact tall-root stroke: compact
  tall radicals now draw a short hook line before the main radical via
  `SQRT_TALL_COMPACT_HOOK_X_PT = 0.55pt`,
  `SQRT_TALL_COMPACT_HOOK_START_Y_RATIO = 0.62`, and
  `SQRT_TALL_COMPACT_HOOK_END_Y_RATIO = 0.52`. This changes only WMF polyline
  geometry; glyph positions and advances remain unchanged. Cache key:
  `v206-compact-sqrt-hook`.
- v206 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-014617.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v206.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v206/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v206 evidence: case12's WMF polyline count rises from `3` to `4`, and case18
  rises from `2` to `3`, while glyph metrics stay stable (`a/b` around
  `19.9pt`, `l/g` around `35.9/34.8pt`). Word PNG visual check shows the short
  hook is visible and reduces the single-straight-leg look, but the radical is
  still too thin and mechanical. The next useful step is likely a thicker
  structure-line profile or a second close parallel stroke for compact roots,
  provided structural scans continue to reject bitmap fallback and line
  protrusion.
- v207 thickens compact tall radicals without changing the global WMF pen:
  because `STRUCTURE_LINE_WIDTH_PT` is shared by fraction bars, roots, accents,
  arrows, and boxes, globally increasing it would also thicken text-mixed and
  ordinary fraction bars. Instead v207 adds a root-only near-parallel shadow
  stroke controlled by `SQRT_TALL_COMPACT_SHADOW_X_OFFSET_PT = 0.18pt` and
  `SQRT_TALL_COMPACT_SHADOW_Y_OFFSET_PT = 0.0pt`. Cache key:
  `v207-compact-sqrt-shadow-stroke`.
- v207 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-015319.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v207.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v207/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v207 evidence: case12's WMF polyline count rises from `4` to `5`, and case18
  rises from `3` to `4`, while glyph metrics stay stable (`a/b` around
  `19.9pt`, `l/g` around `35.9/34.8pt`). Word PNG visual check shows the
  compact radical is visibly less hairline and does not smear into a dirty
  blob. Remaining gap: case12 still looks cramped and the radical/top-bar joint
  is still too angular. Next useful work should focus on top-turn geometry and
  nested radical/body scale, not global line width.
- v208 candidate was rejected, not shipped as a renderer change. Extending the
  inner radical top bar for `\sqrt{1+\sqrt{\frac{a}{b}}}` by `0.85pt`, then
  even by `0.35pt`, looked like a harmless coverage tweak but still changed the
  WMF/vector right bbox enough for Word preview scaling to shrink case12 glyphs:
  case12 `a/b` moved from the v207 `19.9pt` area to about `19.5pt`, and `1+`
  moved from about `8.2pt` to `8.1pt`. Removing the extra width from layout was
  not enough because the stroke endpoint itself still participates in the
  rendered bbox.
- The committed v208 lesson is a guard, not a geometry bump: keep the cache key
  at `v207-compact-sqrt-shadow-stroke`, remove the top-overhang constant, and
  add a regression assertion that nested sqrt-body fraction text right edge
  stays at least `22.8pt`. This prevents future "make the top bar longer"
  tweaks from silently paying for coverage by shrinking the inner glyph slot.
  Future root work should change stroke profile inside the existing bbox, or
  explicitly grow the physical shape and accept the size change with Word PNG
  evidence.
- v209 improves compact tall radicals without changing glyph scale or the right
  bbox: it inserts an internal top-lead point between the compact shoulder and
  top turn (`SQRT_TALL_COMPACT_CHECK_TOP_LEAD_X_PT = 3.55pt`,
  `SQRT_TALL_COMPACT_CHECK_TOP_LEAD_Y_RATIO = 0.19`) and applies it to both the
  root-only shadow stroke and the main compact tall radical. `topBarEnd`,
  `bodyX`, `scaledBodyWidth`, and text advances are unchanged. Cache key:
  `v209-compact-sqrt-top-lead`.
- v209 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-023037.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v209.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v209/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v209 evidence: case12 and case18 keep the v207 glyph metrics (`case12 a/b`
  still around `19.9pt/right=22.9pt`, `case18 l/g` still around
  `35.9/34.8pt`) while their WMF byte size rises only for the compact
  sqrt-fraction cases due to the extra internal point. Word crops
  `case12-v209-zoom.png`, `case18-v209-zoom.png`, and
  `case15-v209-text-mixed.png` show a slightly less abrupt top-turn transition
  and no text-mixed regression. Remaining gap: the visual improvement is
  modest; case12 still has a mechanical inner radical and a long lower leg.
  Next work should either tune the lower/mid compact leg inside the existing
  bbox or add rendered-ink bbox measurement before larger shape changes.
- v210 lifts only the compact tall sqrt-fraction lower check point by adding
  `SQRT_TALL_COMPACT_CHECK_LOW_Y_RATIO = 0.76` while ordinary tall radicals keep
  `SQRT_TALL_CHECK_LOW_Y_RATIO = 0.80`. The renderer applies the same compact
  lower-point ratio to the main radical and the root-only shadow stroke; it does
  not change `topBarEnd`, `bodyX`, `scaledBodyWidth`, text advances, or the
  compact top-turn X. Cache key: `v210-compact-sqrt-low-lift`.
- v210 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-024011.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v210.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v210/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v210 evidence: case12 glyph metrics stayed at the v209 level (`1+ x=8.2pt`,
  `a/b x=19.9pt/right=22.9pt`), and case18 stayed stable (`l x=35.9pt`,
  `g x=34.8pt/right=39.65pt`). Word crops `case12-v210-zoom.png`,
  `case18-v210-wide.png`, and `case15-v210-text-mixed.png` show a small
  positive change: the compact lower leg is less long, with no mixed-text
  regression. Remaining gap: roots still look like mechanical polylines, not
  MathType-style strokes. The next useful round should introduce a richer
  compact radical stroke profile, such as an extra lower-curve point or
  short diagonal transition inside the same bbox, rather than repeatedly nudging
  a single Y ratio.
- v211 adds that richer compact radical profile as an internal lower-transition
  point: `SQRT_TALL_COMPACT_CHECK_LOWER_TRANSITION_X_PT = 1.28pt` and
  `SQRT_TALL_COMPACT_CHECK_LOWER_TRANSITION_Y_RATIO = 0.68`. The point is
  inserted between the compact low point and mid point for both the main radical
  and the root-only shadow stroke. It stays inside the existing bbox and does
  not move `topBarEnd`, `bodyX`, `scaledBodyWidth`, `rootX`, `rootHeight`, or
  text advances. Cache key: `v211-compact-sqrt-lower-transition`.
- v211 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-024757.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v211.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v211/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v211 evidence: case12 and case18 glyph metrics are unchanged from v210
  (`case12 a/b x=19.9pt/right=22.9pt`, `case18 l x=35.9pt`,
  `g x=34.8pt/right=39.65pt`), proving the extra point did not trigger the
  v208-style Word preview shrink. Word crops `case12-v211-zoom.png`,
  `case18-v211-wide.png`, and `case15-v211-text-mixed.png` show a slightly
  smoother lower leg and no mixed-text regression. Remaining gap: the radical
  still reads as a mechanical polyline because the parallel shadow stroke traces
  every segment uniformly. The next useful step is to tune the compact root
  stroke model itself, for example by limiting the shadow to selected segments
  or using a dedicated root pen/profile, instead of adding more internal points.
- v212 changes the compact tall sqrt-fraction shadow from a full copied radical
  outline into a local lower-leg stroke. The shadow still starts at
  `rootX + SQRT_TALL_COMPACT_SHADOW_X_OFFSET_PT` and follows the compact low
  and lower-transition points, but now stops at the compact mid point instead
  of duplicating the shoulder, top-lead, and top-turn segments. This keeps the
  root-only thickening where the lower leg needs weight while reducing the
  double-line corner effect near the upper turn. Cache key:
  `v212-compact-sqrt-local-shadow`.
- v212 validation regenerated
  `analysis/formula-golden-corpus/formula-golden-corpus-20260622-025507.docx`,
  exported Word PDF
  `analysis/formula-golden-corpus/formula-golden-corpus-word-export-v212.pdf`,
  and PNG pages under
  `analysis/formula-golden-corpus/word-rendered-v212/page-*.png`. Structural
  scans stayed clean: `23` MathType OLE objects, `23` WMF previews, zero
  visible LaTeX leaks, zero invalid MathType OLE, zero bitmap/StretchDIB WMFs,
  and zero review-required suspicious WMF text.
- v212 evidence: case12 and case18 glyph metrics remain unchanged
  (`case12 a/b x=19.9pt/right=22.9pt`, `case18 l x=35.9pt`,
  `g x=34.8pt/right=39.65pt`), so localizing the shadow does not affect Word
  preview scale. Word crops `case12-v212-zoom.png`, `case18-v212-wide.png`, and
  `case15-v212-text-mixed.png` show fewer double-line upper-corner artifacts
  and no text-mixed regression. Remaining gap: the upper radical is now cleaner
  but still a little thin and straight, so the next useful work should tune a
  real compact-root stroke profile or root-only pen width rather than returning
  to full-length shadow copying.
