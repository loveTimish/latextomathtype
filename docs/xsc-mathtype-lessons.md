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
