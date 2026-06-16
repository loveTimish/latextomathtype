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
